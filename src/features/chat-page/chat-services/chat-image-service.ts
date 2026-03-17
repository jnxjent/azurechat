// src/features/chat-page/chat-services/chat-image-service.ts
"use server";
import "server-only";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { GetBlob, UploadBlob } from "../../common/services/azure-storage";
import { ChatThreadModel } from "./models";

const IMAGE_CONTAINER_NAME = "images";
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
