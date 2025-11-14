// src/features/chat-page/chat-services/chat-api/chat-image-service.ts
"use server";
import "server-only";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { GetBlob, UploadBlob } from "../../common/services/azure-storage";

const IMAGE_CONTAINER_NAME = "images";
// 相対パスのほうがローカル/本番で安全（NEXTAUTH_URL未設定事故を防ぐ）
const IMAGE_API_PATH = "/api/images";

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
  // /api/images?t=<threadId>&img=<fileName> 形式に統一（末尾の / を排除）
  const params = `?t=${threadId}&img=${fileName}`;
  return `${IMAGE_API_PATH}${params}`;
};

// ===================== ここから追加：overlay 呼び出し（textB64 対応） =====================

type OverlayTargetRect = { x: number; y: number; w: number; h: number };

export type OverlayPayload = {
  /** UTF-8のプレーン文字列（内部で base64 化して textB64 を付与） */
  text: string;

  /** どちらかで指定してください（sourceBlobName が優先） */
  sourceBlobName?: string;        // 例: "thread123/base.png"
  threadId?: string;              // 例: "thread123"
  sourceFileName?: string;        // 例: "base.png"

  /** 省略時は route 側で自動命名（<元名>__text_<ts>.png） */
  resultBlobName?: string;

  /** 以降は route.ts と同じパラメータ */
  fontSize?: number;              // 例: 72
  bottomMargin?: number;          // 例: 80
  targetRect?: OverlayTargetRect; // 例: { x, y, w, h }
  debugRect?: boolean;            // ガイド矩形の描画有無
  overwrite?: boolean;            // trueで同名上書き
};

/**
 * /api/images/overlay を叩く最小関数。
 * - text は UTF-8 → Base64 に変換し、text と textB64 の両方を送る（後方互換）。
 * - 呼び出し側は threadId+sourceFileName または sourceBlobName のどちらかでOK。
 */
export async function OverlayTextOnImage(
  payload: OverlayPayload
): Promise<ServerActionResponse<{ blobUrl: string; resultBlobName: string }>> {
  try {
    const {
      text,
      sourceBlobName,
      threadId,
      sourceFileName,
      resultBlobName,
      fontSize,
      bottomMargin,
      targetRect,
      debugRect,
      overwrite,
    } = payload;

    // sourceBlobName が無ければ threadId + sourceFileName から合成
    const resolvedSource =
      sourceBlobName ??
      (threadId && sourceFileName ? GetBlobPath(threadId, sourceFileName) : undefined);

    if (!resolvedSource) {
      return {
        status: "ERROR",
        errors: [{ message: "sourceBlobName または (threadId + sourceFileName) を指定してください。" }],
      };
    }

    // ★ text を UTF-8 → Base64 に変換（バックエンドは textB64 優先）
    const textB64 = text ? Buffer.from(text, "utf8").toString("base64") : undefined;

    const res = await fetch(`${IMAGE_API_PATH}/overlay`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      // 後方互換のため text も残し、textB64 を追加
      body: JSON.stringify({
        sourceBlobName: resolvedSource,
        resultBlobName,
        text,
        textB64,
        fontSize,
        bottomMargin,
        targetRect,
        debugRect,
        overwrite,
      }),
      cache: "no-store",
    });

    const json = await res.json().catch(() => ({} as any));

    if (!res.ok) {
      return {
        status: "ERROR",
        errors: [{ message: json?.error || `overlay HTTP ${res.status}` }],
      };
    }

    return {
      status: "OK",
      response: {
        blobUrl: json.blobUrl as string,
        resultBlobName: (json.resultBlobName as string) ?? "",
      },
    };
  } catch (e: any) {
    return {
      status: "ERROR",
      errors: [{ message: e?.message ?? String(e) }],
    };
  }
}

// ===================== 追加ここまで =====================

export const GetThreadAndImageFromUrl = (
  urlString: string
): ServerActionResponse<{ threadId: string; imgName: string }> => {
  // Get threadId and img from query parameters t and img
  const url = new URL(urlString);
  const threadId = url.searchParams.get("t");
  const imgName = url.searchParams.get("img");

  // Check if threadId and img are valid
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
