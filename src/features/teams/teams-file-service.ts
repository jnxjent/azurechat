import "server-only";

import { randomUUID } from "crypto";
import {
  DownloadBlobAsText,
  GenerateSasUrl,
  UploadBlob,
} from "@/features/common/services/azure-storage";
import {
  assertSafeTeamsDownloadUrl,
  MAX_TEAMS_FILE_BYTES,
  MAX_TEAMS_TOTAL_BYTES,
  parseTeamsFileCandidates,
  type TeamsAttachmentLike,
  type TeamsStoredFile,
  validateTeamsFileBytes,
} from "./teams-file-policy";

const DOWNLOAD_TIMEOUT_MS = 30_000;
const latestUploadPointerName = (threadId: string) =>
  `thread-${threadId}-teams-upload-latest.json`;

export async function receiveTeamsFiles(props: {
  attachments: readonly TeamsAttachmentLike[] | undefined;
  threadId: string;
}): Promise<TeamsStoredFile[]> {
  const candidates = parseTeamsFileCandidates(props.attachments);
  if (candidates.length === 0) return [];

  const downloaded: Array<{
    buffer: Buffer;
    extension: string;
    fileName: string;
  }> = [];
  let totalBytes = 0;
  for (const candidate of candidates) {
    const buffer = await downloadTeamsFile(candidate.downloadUrl);
    validateTeamsFileBytes(candidate.fileName, candidate.extension, buffer);
    totalBytes += buffer.length;
    if (totalBytes > MAX_TEAMS_TOTAL_BYTES) {
      throw new Error("添付ファイルの合計サイズは40MBまでです。");
    }
    downloaded.push({
      buffer,
      extension: candidate.extension,
      fileName: candidate.fileName,
    });
  }

  const stored: TeamsStoredFile[] = [];
  for (const file of downloaded) {
    const blobPath = `teams-upload/${props.threadId}/${Date.now()}-${randomUUID()}-${file.fileName}`;
    const upload = await UploadBlob("dl-link", blobPath, file.buffer);
    if (upload.status !== "OK") {
      throw new Error(`「${file.fileName}」をAzure Storageへ保存できませんでした。`);
    }
    const sas = await GenerateSasUrl("dl-link", blobPath);
    if (sas.status !== "OK") {
      throw new Error(`「${file.fileName}」の参照URLを作成できませんでした。`);
    }
    stored.push({
      extension: file.extension,
      fileName: file.fileName,
      savedAt: Date.now(),
      size: file.buffer.length,
      url: sas.response,
    });
  }

  const pointer = await UploadBlob(
    "dl-link",
    latestUploadPointerName(props.threadId),
    Buffer.from(JSON.stringify(stored))
  );
  if (pointer.status !== "OK") {
    throw new Error("Teams添付の会話ポインターを保存できませんでした。");
  }

  console.log("[teams-file] stored", {
    count: stored.length,
    files: stored.map((file) => ({
      extension: file.extension,
      name: file.fileName,
      size: file.size,
    })),
    threadId: props.threadId,
  });
  return stored;
}

export async function readLatestTeamsFiles(
  threadId: string
): Promise<TeamsStoredFile[]> {
  const response = await DownloadBlobAsText(
    "dl-link",
    latestUploadPointerName(threadId)
  );
  if (response.status !== "OK") return [];
  try {
    const parsed = JSON.parse(response.response) as unknown;
    return Array.isArray(parsed)
      ? parsed.filter(isValidStoredFile).slice(0, 3)
      : [];
  } catch {
    return [];
  }
}

async function downloadTeamsFile(url: string): Promise<Buffer> {
  assertSafeTeamsDownloadUrl(url);
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), DOWNLOAD_TIMEOUT_MS);
  try {
    const response = await fetch(url, {
      headers: { Accept: "application/octet-stream" },
      redirect: "follow",
      signal: controller.signal,
    });
    if (!response.ok) {
      throw new Error(`Teams添付の取得に失敗しました: HTTP ${response.status}`);
    }
    assertSafeTeamsDownloadUrl(response.url);
    const contentLength = Number(response.headers.get("content-length") ?? 0);
    if (contentLength > MAX_TEAMS_FILE_BYTES) {
      throw new Error("Teams添付は25MBを超えているため取得できません。");
    }
    if (!response.body) throw new Error("Teams添付の内容が空です。");

    const chunks: Buffer[] = [];
    const reader = response.body.getReader();
    let total = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > MAX_TEAMS_FILE_BYTES) {
        await reader.cancel();
        throw new Error("Teams添付は25MBを超えているため取得できません。");
      }
      chunks.push(Buffer.from(value));
    }
    return Buffer.concat(chunks);
  } finally {
    clearTimeout(timeout);
  }
}

function isValidStoredFile(value: unknown): value is TeamsStoredFile {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const file = value as Partial<TeamsStoredFile>;
  return Boolean(
    typeof file.extension === "string" &&
      typeof file.fileName === "string" &&
      typeof file.savedAt === "number" &&
      typeof file.size === "number" &&
      typeof file.url === "string"
  );
}
