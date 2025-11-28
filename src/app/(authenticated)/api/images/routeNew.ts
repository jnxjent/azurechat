// src/app/(authenticated)/api/images/route.ts
import { NextRequest } from "next/server";
import { GetImageFromStore } from "@/features/chat-page/chat-services/chat-image-service";

export const runtime = "nodejs";

/**
 * 画像取得エンドポイント
 * /api/images?t=<threadId>&img=<blobName>
 *
 * - 通常は img をそのまま探す
 * - 見つからなければ img + ".png" も再トライ（拡張子抜け対策）
 */
export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url);
  const threadId = searchParams.get("t");
  const imgName = searchParams.get("img");

  if (!threadId || !imgName) {
    return new Response(
      JSON.stringify({
        error: "invalid_request",
        detail: "Missing query parameters: t (threadId) and img (image name) are required.",
      }),
      {
        status: 400,
        headers: { "Content-Type": "application/json" },
      }
    );
  }

  // --- 1st try: imgName（拡張子無しのまま）
  let blobResult = await GetImageFromStore(threadId, imgName);

  // --- 2nd try: 拡張子 ".p
