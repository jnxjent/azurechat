
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

/**
 * 画像ファイル名をスレッドに登録するヘルパー。
 *
 * - originalImageFileName:
 *     まだ未設定なら、最初の1回だけここに入る（＝元絵）
 * - lastImageFileName:
 *     毎回、最新の画像ファイル名で上書きされる
 *
 * ※ ChatThreadModel を mutate するだけの純粋関数。
 *   実際の永続化（DB更新など）は呼び出し側で行う想定。
 */
export const RegisterImageOnThread = (
  thread: ChatThreadModel,
  fileName: string
): void => {
  if (!thread.originalImageFileName) {
    thread.originalImageFileName = fileName;
  }
  thread.lastImageFileName = fileName;
};

/**
 * 文字入れ（オーバーレイ）の「ベースとなる画像」のファイル名を返す。
 *
 * ✅ フォールバックなし：
 *   - つねに originalImageFileName（＝元絵）のみを使う
 *   - 設定されていない場合は undefined を返す
 *
 * これにより、
 * 「2回目以降の編集でも常に元絵に対して文字を描画」
 * となり、過去の文字レイヤーが重なり続けることを防ぐ。
 */
export const GetBaseImageFileNameForOverlay = (
  thread: ChatThreadModel
): string | undefined => {
  return thread.originalImageFileName;
};

/**
 * 便利ヘルパー：スレッドから「元絵」の画像URLを返す。
 *
 * - originalImageFileName が存在しない場合は undefined を返す。
 * - 余計なフォールバックは行わない（＝元絵がなければ処理自体を見直すべき状態）。
 */
export const GetImageUrlFromThread = (
  thread: ChatThreadModel
): string | undefined => {
  const base = thread.originalImageFileName; // ★ 元絵のみ
  if (!base) return undefined;
  return GetImageUrl(thread.id, base);
};

