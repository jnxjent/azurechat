import { isIP } from "net";

export const TEAMS_FILE_DOWNLOAD_CONTENT_TYPE =
  "application/vnd.microsoft.teams.file.download.info";
export const MAX_TEAMS_FILE_COUNT = 3;
export const MAX_TEAMS_FILE_BYTES = 25 * 1024 * 1024;
export const MAX_TEAMS_TOTAL_BYTES = 40 * 1024 * 1024;

const ALLOWED_EXTENSIONS = new Set([
  "pdf",
  "docx",
  "xlsx",
  "pptx",
  "png",
  "jpg",
  "jpeg",
  "webp",
]);

export type TeamsAttachmentLike = {
  contentType?: string;
  contentUrl?: string;
  content?: unknown;
  name?: string;
};

export type TeamsFileCandidate = {
  downloadUrl: string;
  extension: string;
  fileName: string;
  uniqueId?: string;
};

export type TeamsStoredFile = {
  extension: string;
  fileName: string;
  savedAt: number;
  size: number;
  url: string;
};

export function parseTeamsFileCandidates(
  attachments: readonly TeamsAttachmentLike[] | undefined
): TeamsFileCandidate[] {
  const fileAttachments = (attachments ?? []).filter(
    (attachment) =>
      attachment.contentType?.toLowerCase() ===
      TEAMS_FILE_DOWNLOAD_CONTENT_TYPE
  );
  if (fileAttachments.length > MAX_TEAMS_FILE_COUNT) {
    throw new Error(
      `一度に添付できるファイルは${MAX_TEAMS_FILE_COUNT}件までです。`
    );
  }

  return fileAttachments.map((attachment) => {
    const content = isRecord(attachment.content) ? attachment.content : {};
    const downloadUrl = stringValue(content.downloadUrl);
    const originalName = attachment.name?.trim() || "";
    const contentFileType = stringValue(content.fileType).replace(/^\./, "");
    const extension = (
      extensionFromName(originalName) || contentFileType
    ).toLowerCase();

    if (!downloadUrl) {
      throw new Error("Teams添付のダウンロードURLを取得できませんでした。");
    }
    assertSafeTeamsDownloadUrl(downloadUrl);
    if (!extension || !ALLOWED_EXTENSIONS.has(extension)) {
      throw new Error(
        `未対応のファイル形式です: ${originalName || contentFileType || "不明"}`
      );
    }

    return {
      downloadUrl,
      extension,
      fileName: sanitizeTeamsFileName(
        originalName || `teams-upload.${extension}`,
        extension
      ),
      uniqueId: stringValue(content.uniqueId) || undefined,
    };
  });
}

export function assertSafeTeamsDownloadUrl(value: string): void {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error("Teams添付のダウンロードURLが不正です。");
  }
  const host = url.hostname.toLowerCase();
  if (
    url.protocol !== "https:" ||
    !host ||
    host === "localhost" ||
    host.endsWith(".localhost") ||
    host.endsWith(".local") ||
    isIP(host) !== 0
  ) {
    throw new Error("Teams添付のダウンロード先が許可されていません。");
  }
}

export function validateTeamsFileBytes(
  fileName: string,
  extension: string,
  buffer: Buffer
): void {
  if (buffer.length === 0) {
    throw new Error(`「${fileName}」の内容が空です。`);
  }
  if (buffer.length > MAX_TEAMS_FILE_BYTES) {
    throw new Error(
      `「${fileName}」は25MBを超えているためアップロードできません。`
    );
  }

  const zipBased = new Set(["docx", "xlsx", "pptx"]);
  const valid =
    (extension === "pdf" && buffer.subarray(0, 5).toString() === "%PDF-") ||
    (extension === "png" &&
      buffer.subarray(0, 8).equals(
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
      )) ||
    ((extension === "jpg" || extension === "jpeg") &&
      buffer[0] === 0xff &&
      buffer[1] === 0xd8 &&
      buffer[2] === 0xff) ||
    (extension === "webp" &&
      buffer.subarray(0, 4).toString("ascii") === "RIFF" &&
      buffer.subarray(8, 12).toString("ascii") === "WEBP") ||
    (zipBased.has(extension) &&
      buffer[0] === 0x50 &&
      buffer[1] === 0x4b);

  if (!valid) {
    throw new Error(
      `「${fileName}」の拡張子と実際のファイル形式が一致しません。`
    );
  }
}

export function sanitizeTeamsFileName(
  fileName: string,
  extension: string
): string {
  const withoutPath = fileName.split(/[\\/]/).pop() ?? "";
  const base = withoutPath
    .replace(/\.[^.]+$/, "")
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, "_")
    .replace(/[. ]+$/g, "")
    .trim()
    .slice(0, 80);
  return `${base || "teams-upload"}.${extension}`;
}

export function referencesTeamsUpload(message: string): boolean {
  return /(添付|アップロード|このファイル|今のファイル|先ほどのファイル|さっきのファイル)/i.test(
    message
  );
}

export function stripTeamsAttachmentMarkup(message: string): string {
  return message
    .replace(/<attachment\b[^>]*>[\s\S]*?<\/attachment>/gi, " ")
    .replace(/<attachment\b[^>]*\/?>/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extensionFromName(fileName: string): string {
  return fileName.match(/\.([a-z0-9]+)$/i)?.[1] ?? "";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function stringValue(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}
