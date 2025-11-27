// src/features/chat-page/chat-services/chat-image-service.ts
"use server";
import "server-only";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { GetBlob, UploadBlob } from "../../common/services/azure-storage";

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
  return `${IMAGE_API_PATH}${params}`;  // ← ここがポイント（末尾に / を付けない）
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
