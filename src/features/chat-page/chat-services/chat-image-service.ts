// src/features/chat-page/chat-services/chat-image-service.ts
"use server";
import "server-only";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { GetBlob, UploadBlob } from "../../common/services/azure-storage";
import { ChatThreadModel } from "./models";

const IMAGE_CONTAINER_NAME = "images";
const LATEST_ATTACHMENT_BLOB_NAME = "__latest_attachment__.img";
const LATEST_ATTACHMENT_META_BLOB_NAME = "__latest_attachment__.json";
const DEFAULT_ATTACHMENT_TTL_MS = 30 * 60 * 1000;
const MAX_IMAGE_ATTACHMENT_BYTES = 50 * 1024 * 1024;
// ★ まず NEXT_PUBLIC_IMAGE_URL を優先し、なければ NEXTAUTH_URL + /api/images
const IMAGE_API_PATH =
  process.env.NEXT_PUBLIC_IMAGE_URL ||
  (process.env.NEXTAUTH_URL + "/api/images");

export const GetBlobPath = (threadId: string, blobName: string): string => {
  return `${threadId}/${blobName}`;
};

export const UploadImageToStore = async (
  threadId: string,
  fileName: string,
  imageData: Buffer
): Promise<ServerActionResponse<string>> => {
  return await UploadBlob(
    IMAGE_CONTAINER_NAME,
    `${threadId}/${fileName}`,
    imageData
  );
};

export const GetImageFromStore = async (
  threadId: string,
  fileName: string
): Promise<ServerActionResponse<ReadableStream>> => {
  const blobPath = GetBlobPath(threadId, fileName);
  return await GetBlob(IMAGE_CONTAINER_NAME, blobPath);
};

type LatestImageAttachmentMetadata = {
  fileName: string;
  contentType: string;
  savedAt: string;
  size: number;
  consumedAt?: string;
};

export type LatestImageAttachment = LatestImageAttachmentMetadata & {
  buffer: Buffer;
};

function resolveAttachmentTtlMs(): number {
  const configured = Number(process.env.IMAGE_ATTACHMENT_TTL_MS);
  return Number.isFinite(configured) && configured > 0
    ? configured
    : DEFAULT_ATTACHMENT_TTL_MS;
}

/**
 * Keep a short-lived server-side copy of the paperclip image. The normal file
 * upload remains unchanged; this copy is only a fallback reference for GPT
 * Image when the chat message loses its in-memory data URL.
 */
export const SaveLatestImageAttachment = async (
  formData: FormData
): Promise<ServerActionResponse<LatestImageAttachmentMetadata>> => {
  const threadId = String(formData.get("id") ?? "").trim();
  const file = formData.get("file");
  if (!threadId || !(file instanceof File)) {
    return {
      status: "ERROR",
      errors: [{ message: "Image attachment or chat thread is missing." }],
    };
  }

  if (!file.size || file.size >= MAX_IMAGE_ATTACHMENT_BYTES) {
    return {
      status: "ERROR",
      errors: [{ message: "Image attachment must be smaller than 50 MB." }],
    };
  }

  const buffer = Buffer.from(await file.arrayBuffer());
  const metadata: LatestImageAttachmentMetadata = {
    fileName: file.name || "attachment.img",
    contentType: file.type || "application/octet-stream",
    savedAt: new Date().toISOString(),
    size: buffer.length,
  };

  const imageResult = await UploadImageToStore(
    threadId,
    LATEST_ATTACHMENT_BLOB_NAME,
    buffer
  );
  if (imageResult.status !== "OK") return imageResult;

  const metaResult = await UploadImageToStore(
    threadId,
    LATEST_ATTACHMENT_META_BLOB_NAME,
    Buffer.from(JSON.stringify(metadata), "utf-8")
  );
  if (metaResult.status !== "OK") return metaResult;

  return { status: "OK", response: metadata };
};

export const LoadLatestImageAttachment = async (
  threadId: string
): Promise<LatestImageAttachment | null> => {
  const metaResult = await GetImageFromStore(
    threadId,
    LATEST_ATTACHMENT_META_BLOB_NAME
  );
  if (metaResult.status !== "OK" || !metaResult.response) return null;

  try {
    const metadata = JSON.parse(
      await new Response(metaResult.response as any).text()
    ) as LatestImageAttachmentMetadata;
    const savedAtMs = Date.parse(metadata.savedAt);
    if (
      metadata.consumedAt ||
      !Number.isFinite(savedAtMs) ||
      Date.now() - savedAtMs > resolveAttachmentTtlMs()
    ) {
      return null;
    }

    const imageResult = await GetImageFromStore(
      threadId,
      LATEST_ATTACHMENT_BLOB_NAME
    );
    if (imageResult.status !== "OK" || !imageResult.response) return null;
    const buffer = Buffer.from(
      await new Response(imageResult.response as any).arrayBuffer()
    );
    if (!buffer.length || buffer.length >= MAX_IMAGE_ATTACHMENT_BYTES) {
      return null;
    }
    return { ...metadata, buffer };
  } catch (error) {
    console.warn("[image-attachment] Failed to load latest attachment:", error);
    return null;
  }
};

export const ConsumeLatestImageAttachment = async (
  threadId: string
): Promise<void> => {
  const metaResult = await GetImageFromStore(
    threadId,
    LATEST_ATTACHMENT_META_BLOB_NAME
  );
  if (metaResult.status !== "OK" || !metaResult.response) return;

  try {
    const metadata = JSON.parse(
      await new Response(metaResult.response as any).text()
    ) as LatestImageAttachmentMetadata;
    if (!metadata.savedAt || metadata.consumedAt) return;
    metadata.consumedAt = new Date().toISOString();
    await UploadImageToStore(
      threadId,
      LATEST_ATTACHMENT_META_BLOB_NAME,
      Buffer.from(JSON.stringify(metadata), "utf-8")
    );
  } catch (error) {
    // The image edit already succeeded. A cleanup failure must not hide its
    // result; expiry remains the secondary protection against stale reuse.
    console.warn("[image-attachment] Failed to consume attachment:", error);
  }
};

export const GetImageUrl = (threadId: string, fileName: string): string => {
  // ?t=...&img=... を付けるだけ（余分なスラッシュを入れない）
  const params = `?t=${threadId}&img=${fileName}`;
  return `${IMAGE_API_PATH}${params}`; // ← ここがポイント（末尾に / を付けない）
};

export const GetThreadAndImageFromUrl = (
  urlString: string
): ServerActionResponse<{ threadId: string; imgName: string }> => {
  const url = new URL(urlString);
  const threadId = url.searchParams.get("t");
  const imgName = url.searchParams.get("img");

  if (!threadId || !imgName) {
    return {
      status: "ERROR",
      errors: [
        {
          message:
            "Invalid URL, threadId and/or imgName not formatted correctly.",
        },
      ],
    };
  }

  return {
    status: "OK",
    response: {
      threadId,
      imgName,
    },
  };
};

/* -------------------------------------------------------------------------- */
/* ★ 追加：スレッドに「元絵」と「最新画像」を記録／取得するためのヘルパー */
/* -------------------------------------------------------------------------- */

export const RegisterImageOnThread = (
  thread: ChatThreadModel,
  fileName: string
): void => {
  if (!thread.originalImageFileName) {
    thread.originalImageFileName = fileName;
  }
  thread.lastImageFileName = fileName;
};

export const GetBaseImageFileNameForOverlay = (
  thread: ChatThreadModel
): string | undefined => {
  return thread.originalImageFileName;
};

export const GetImageUrlFromThread = (
  thread: ChatThreadModel
): string | undefined => {
  const base = thread.originalImageFileName; // ★ 元絵のみ
  if (!base) return undefined;
  return GetImageUrl(thread.id, base);
};

/* -------------------------------------------------------------------------- */
/* ★ NEW: overlay state JSON を Blob に保存/取得                              */
/* -------------------------------------------------------------------------- */

export type OverlayState = {
  align: "left" | "center" | "right";
  vAlign: "top" | "middle" | "bottom";
  offsetX: number;
  offsetY: number;
  size: "small" | "medium" | "large" | "xlarge";
  text: string;
  color?: string;
  fontFamily?: "gothic" | "mincho" | "meiryo";
  bold?: boolean;
  italic?: boolean;
};

const OVERLAY_STATE_BLOB_NAME = "__overlay_state__.json";

export const SaveOverlayStateToStore = async (
  threadId: string,
  state: OverlayState
): Promise<ServerActionResponse<string>> => {
  const json = JSON.stringify(state ?? {}, null, 2);
  const buf = Buffer.from(json, "utf-8");
  return await UploadBlob(
    IMAGE_CONTAINER_NAME,
    `${threadId}/${OVERLAY_STATE_BLOB_NAME}`,
    buf
  );
};

export const LoadOverlayStateFromStore = async (
  threadId: string
): Promise<ServerActionResponse<OverlayState | null>> => {
  const blobPath = `${threadId}/${OVERLAY_STATE_BLOB_NAME}`;
  const res = await GetBlob(IMAGE_CONTAINER_NAME, blobPath);

  if (res.status !== "OK") {
    // 未作成は普通に起きるので「null」で返す（ERROR扱いにしない）
    return { status: "OK", response: null };
  }

  try {
    // GetBlob が ReadableStream を返す前提（Nodeの fetch Response で読める）
    const stream = res.response!;
    const text = await new Response(stream as any).text();
    const obj = JSON.parse(text || "null");
    if (!obj) return { status: "OK", response: null };
    return { status: "OK", response: obj as OverlayState };
  } catch (e: any) {
    return {
      status: "ERROR",
      errors: [{ message: "Failed to parse overlay state JSON: " + String(e) }],
    };
  }
};
