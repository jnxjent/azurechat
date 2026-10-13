import "server-only";
import { generateStoredChatImage } from "@/features/chat-page/chat-services/chat-api/image/image-generation-service";
import { createHash } from "crypto";
import { GenerateSasUrl } from "@/features/common/services/azure-storage";

const requests = new Map<string, Promise<string>>();

export function createTeamsImageReply(props: {
  message: string;
  conversationId: string;
  activityId: string;
}): Promise<string> {
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

async function generate(props: { message: string; conversationId: string }): Promise<string> {
  const threadId = `teams-${createHash("sha256").update(props.conversationId).digest("hex").slice(0, 32)}`;
  try {
    const result = await generateStoredChatImage({
      prompt: props.message, threadId, userMessage: props.message,
      signal: AbortSignal.timeout(240_000),
    });
    if ("error" in result) {
      return "画像を生成できませんでした。時間をおいてもう一度お試しください。";
    }
    // Teams cannot use the AC image route requiring a browser login.
    const signed = await GenerateSasUrl("images", `${threadId}/${result.fileName}`);
    if (signed.status !== "OK") return "画像は生成されましたが、表示用リンクを取得できませんでした。";
    return `画像を生成しました。\n\n![生成画像](${signed.response})\n\n[画像を開く](${signed.response})`;
  } catch {
    console.error("[teams-image] generation or delivery failed");
    return "画像を生成できませんでした。時間をおいてもう一度お試しください。";
  }
}
