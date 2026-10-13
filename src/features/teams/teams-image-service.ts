import "server-only";
import { generateStoredChatImage } from "@/features/chat-page/chat-services/chat-api/image/image-generation-service";
import { createHash } from "crypto";
import { GenerateSasUrl, GetBlob, UploadBlob } from "@/features/common/services/azure-storage";
import { GetImageFromStore, UploadImageToStore } from "@/features/chat-page/chat-services/chat-image-service";
import { buildFaithfulImagePrompt } from "@/features/chat-page/chat-services/chat-api/image/image-intent";
import { isTeamsImageEditRequest } from "./teams-image-intent";
import { OpenAIDALLEInstance } from "@/features/common/services/openai";
import { uniqueId } from "@/features/common/util";
import { toFile } from "openai";
import sharp from "sharp";
import type { TeamsStoredFile } from "./teams-file-policy";

type ImageReply = string | {
  type: "message";
  text: string;
  attachments: { contentType: "image/jpeg"; contentUrl: string; name: string }[];
};

const requests = new Map<string, Promise<ImageReply>>();

export function createTeamsImageReply(props: {
  message: string;
  conversationId: string;
  activityId: string;
  uploadedFiles?: TeamsStoredFile[];
}): Promise<ImageReply> {
  const key = `${props.conversationId}\n${props.activityId}`;
  const existing = props.activityId ? requests.get(key) : undefined;
  if (existing) return existing;
  const request = generate(props);
  if (props.activityId) {
    requests.set(key, request);
    if (requests.size > 500) requests.delete(requests.keys().next().value!);
  }
  return request;
}

async function readImage(threadId: string, name: string): Promise<Buffer | null> {
  const stored = await GetImageFromStore(threadId, name);
  return stored.status === "OK" ? streamBuffer(stored.response) : null;
}

async function streamBuffer(stream: unknown): Promise<Buffer> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of stream as AsyncIterable<Uint8Array>) {
    size += chunk.length;
    if (size > 50 * 1024 * 1024) throw new Error("Image input too large");
    chunks.push(Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

async function readReference(file: TeamsStoredFile, threadId: string): Promise<Buffer> {
  const url = new URL(file.url);
  const account = (process.env.AZURE_STORAGE_ACCOUNT_NAME ?? "").trim();
  const path = decodeURIComponent(url.pathname);
  if (url.protocol !== "https:" || url.hostname !== `${account}.blob.core.windows.net` ||
      !path.startsWith(`/dl-link/teams-upload/${threadId}/`)) throw new Error("Invalid stored image reference");
  const stored = await GetBlob("dl-link", path.slice("/dl-link/".length));
  if (stored.status !== "OK") throw new Error("Image reference not found");
  return sharp(await streamBuffer(stored.response)).rotate().png().toBuffer();
}

async function generate(props: { message: string; conversationId: string; uploadedFiles?: TeamsStoredFile[] }): Promise<ImageReply> {
  const threadId = `teams-${createHash("sha256").update(props.conversationId).digest("hex").slice(0, 32)}`;
  try {
    const editing = isTeamsImageEditRequest(props.message);
    let fileName: string;
    if (editing) {
      const base = await readImage(threadId, "__latest__.png");
      if (!base) return "編集対象の画像が見つかりません。この会話で元の画像を生成してから、差し替えを依頼してください。";
      const references = (props.uploadedFiles ?? []).filter(file => /^(png|jpg|jpeg|webp)$/.test(file.extension));
      if (/(添付|アップロード)/i.test(props.message) && references.length === 0) return "差し替えに使う添付画像を確認できませんでした。ロゴを画像ファイルとして再添付してください。";
      const buffers = [base, ...await Promise.all(references.map(file => readReference(file, threadId)))];
      const files = await Promise.all(buffers.map((buffer, index) => toFile(buffer, `image-${index + 1}.png`, { type: "image/png" })));
      const prompt = buildFaithfulImagePrompt(props.message, props.message, "edit") +
        "\n画像1が編集対象です。画像2以降はユーザーが添付した参照画像です。添付ロゴの差し替えでは、参照画像のロゴの形・配色・文字・綴りを忠実に使用してください。別の文字や新しいロゴを作らず、対象箇所以外の車体・構図・背景を維持してください。";
      const response = await OpenAIDALLEInstance().images.edit({
        model: process.env.AZURE_OPENAI_DALLE_API_DEPLOYMENT_NAME!,
        image: files.length === 1 ? files[0] : files,
        prompt, size: "auto", quality: "high",
      }, { signal: AbortSignal.timeout(240_000), timeout: 240_000, maxRetries: 0 });
      if (!response.data?.[0]?.b64_json) throw new Error("Invalid edit response");
      const buffer = Buffer.from(response.data[0].b64_json, "base64");
      fileName = `${uniqueId()}.png`;
      for (const name of [fileName, "__base__.png", "__latest__.png"]) {
        if ((await UploadImageToStore(threadId, name, buffer)).status !== "OK") throw new Error("Image storage failed");
      }
    } else {
      const result = await generateStoredChatImage({
        prompt: props.message, threadId, userMessage: props.message,
        signal: AbortSignal.timeout(240_000),
      });
      if ("error" in result) throw new Error("Image generation failed");
      fileName = result.fileName;
    }
    const buffer = await readImage(threadId, fileName);
    if (!buffer) throw new Error("Image output not found");
    const preview = await sharp(buffer).resize({ width: 1024, height: 1024, fit: "inside", withoutEnlargement: true }).jpeg({ quality: 80 }).toBuffer();
    if (preview.length > 1024 * 1024) throw new Error("Image preview too large");
    const previewPath = `${threadId}/${fileName}.preview.jpg`;
    if ((await UploadBlob("images", previewPath, preview, false, "image/jpeg")).status !== "OK") throw new Error("Preview storage failed");
    const [signed, original] = await Promise.all([
      GenerateSasUrl("images", previewPath, "image/jpeg"),
      GenerateSasUrl("images", `${threadId}/${fileName}`, "image/png"),
    ]);
    if (signed.status !== "OK" || original.status !== "OK") throw new Error("Image URL signing failed");
    return {
      type: "message",
      text: `画像を${editing ? "編集" : "生成"}しました。\n\n[画像を開く](${original.response})`,
      attachments: [{ contentType: "image/jpeg", contentUrl: signed.response, name: "generated-image.jpg" }],
    };
  } catch (error) {
    const failure = error as { name?: string; status?: number; code?: string };
    console.error("[teams-image] generation, editing or delivery failed", {
      name: failure?.name, status: failure?.status, code: failure?.code,
    });
    return "画像の生成・編集または表示準備に失敗しました。時間をおいてもう一度お試しください。";
  }
}
