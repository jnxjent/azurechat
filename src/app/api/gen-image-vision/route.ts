// src/app/api/gen-image-vision/route.ts
export const runtime = "nodejs";
export const dynamic = "force-dynamic";

import { NextRequest } from "next/server";
import {
  BlobServiceClient,
  StorageSharedKeyCredential,
} from "@azure/storage-blob";

// ★ 環境変数
const VISION_ENDPOINT = `https://${process.env.AZURE_OPENAI_VISION_API_INSTANCE_NAME}.openai.azure.com`;
const VISION_KEY = process.env.AZURE_OPENAI_VISION_API_KEY!;
const VISION_DEPLOYMENT =
  process.env.AZURE_OPENAI_VISION_API_DEPLOYMENT_NAME!;
const VISION_API_VERSION =
  process.env.AZURE_OPENAI_VISION_API_VERSION || "2024-02-01";

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    // ★ offsetX / offsetY も受け取る（Edit ツールから渡ってくる）
    const {
      imageUrl,
      text,
      instruction,
      offsetX = 0,
      offsetY = 0,
    } = body as {
      imageUrl: string;
      text: string;
      instruction?: string;
      offsetX?: number;
      offsetY?: number;
    };

    if (!imageUrl || !text) {
      return new Response(
        JSON.stringify({ error: "imageUrl and text are required" }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    // ★ localhost URLを Base64 Data URL に変換
    let visionAccessibleUrl = imageUrl;

    if (imageUrl.includes("/api/images?")) {
      try {
        const url = new URL(imageUrl, "http://localhost");
        const threadId = url.searchParams.get("t");
        const imgName = url.searchParams.get("img");

        if (threadId && imgName) {
          console.log("[Vision] Fetching image as base64...");

          // Azure Storageから画像を取得してBase64化
          const imageBuffer = await getImageBufferFromStorage(
            threadId,
            imgName
          );
          const base64Image = imageBuffer.toString("base64");
          visionAccessibleUrl = `data:image/png;base64,${base64Image}`;

          console.log(
            "[Vision] Converted to base64 data URL (length:",
            base64Image.length,
            ")"
          );
        }
      } catch (e) {
        console.error("[Vision] Failed to get image as base64:", e);
      }
    }

    // 1. Visionでプラカード位置を検出
    console.log("[Vision] Detecting placard position...");
    const placard = await detectPlacardWithVision(
      visionAccessibleUrl,
      instruction || ""
    );

    if (!placard) {
      return new Response(
        JSON.stringify({ error: "Failed to detect placard position" }),
        { status: 500, headers: { "Content-Type": "application/json" } }
      );
    }

    console.log("[Vision] Detected:", placard);

    // 2. 既存のgen-imageを呼び出し
    const textB64 = Buffer.from(text, "utf8").toString("base64");

    // ★ デバッグ：エンコード・デコードテスト
    const decoded = Buffer.from(textB64, "base64").toString("utf8");
    console.log("[Vision Debug] Original text:", text);
    console.log("[Vision Debug] Base64:", textB64);
    console.log("[Vision Debug] Decoded back:", decoded);
    console.log(
      "[Vision Debug] Match:",
      text === decoded ? "YES" : "NO"
    );

    // -------- ここから座標計算を修正（下寄せ＋offsetY対応） --------
    const imgHeight = 1024; // 現状は 1024x1024 前提で処理しているのでそのまま

    // Vision が返したセンターY
    const centerY = placard.y;

    // デフォルトは「センターより少し下（+150px）、ただし下端から80pxは余白」
    let baseY = centerY + 150;
    if (baseY > imgHeight - 80) {
      baseY = imgHeight - 80;
    }

    // ユーザーの「もう少し上/下」を offsetY で反映
    //   offsetY > 0 : 下に移動
    //   offsetY < 0 : 上に移動
    let finalY = baseY + offsetY;

    // 画面外に出ないように上下にマージン
    const minY = 80;
    const maxY = imgHeight - 40;
    if (finalY < minY) finalY = minY;
    if (finalY > maxY) finalY = maxY;

    // gen-image は bottom 基準で marginBottom を見ている想定
    const calculatedMarginBottom = imgHeight - finalY;

    console.log("[Vision Debug] Placard center Y:", centerY);
    console.log("[Vision Debug] imgHeight:", imgHeight);
    console.log(
      "[Vision Debug] Final textY:",
      finalY,
      "(from bottom:",
      calculatedMarginBottom,
      ")"
    );
    // ------------------------------------------------------------

    const genResponse = await fetch(`${req.nextUrl.origin}/api/gen-image`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        imageUrl,
        textB64,
        fontSize: placard.fontSize,
        align: "center", // 水平方向は中央のまま
        vAlign: "bottom",
        marginBottom: calculatedMarginBottom,
        width: 1024,
        height: 1024,
        // ★ 今後左右にも動かしたくなったら、/api/gen-image 側を拡張して offsetX を渡す
      }),
    });

    return genResponse;
  } catch (err: any) {
    console.error("[Vision] Error:", err);
    return new Response(
      JSON.stringify({ error: err?.message || "Unknown error" }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}

// ★ Azure Storageから画像を取得
async function getImageBufferFromStorage(
  threadId: string,
  imgName: string
): Promise<Buffer> {
  const storageAccount = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const accountKey = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const containerName =
    process.env.AZURE_STORAGE_CONTAINER_NAME || "images";

  const credential = new StorageSharedKeyCredential(
    storageAccount,
    accountKey
  );
  const blobServiceClient = new BlobServiceClient(
    `https://${storageAccount}.blob.core.windows.net`,
    credential
  );

  const containerClient =
    blobServiceClient.getContainerClient(containerName);
  const blobClient = containerClient.getBlobClient(
    `${threadId}/${imgName}`
  );

  console.log("[Vision] Downloading blob from storage...");
  const buffer = await blobClient.downloadToBuffer();
  console.log(
    "[Vision] Downloaded successfully, size:",
    buffer.length,
    "bytes"
  );

  return buffer;
}

async function detectPlacardWithVision(
  imageUrl: string,
  instruction: string
): Promise<{ x: number; y: number; fontSize: number } | null> {
  try {
    // ★ プロンプトを改善
    const prompt = instruction
      ? `この画像には、キャラクターが白いプラカード（看板）を持っています。
ユーザーの指示: "${instruction}"

**白いプラカード（看板）の中心位置**と適切なフォントサイズをJSON形式で返してください。
プラカードは画像の下半分、キャラクターの下部にある白い矩形領域です。
"Nice to meet you!"という英語テキストが表示されている白い看板の中心位置を特定してください。
画像サイズは1024x1024です。

**必ずJSON形式で返してください：**
{
  "x": プラカードの中心X座標（通常512付近）,
  "y": プラカードの中心Y座標（450-600の範囲、白い看板の中心）,
  "fontSize": 適切なフォントサイズ（40-60）
}

例: {"x": 512, "y": 520, "fontSize": 48}`
      : `この画像には、キャラクターが白いプラカード（看板）を持っています。

**白いプラカード（看板）の中心位置**を特定してください。
プラカードは画像の下半分、キャラクターの下部にある白い矩形領域です。
"Nice to meet you!"という英語テキストが表示されている白い看板の中心位置を特定してください。
画像サイズは1024x1024です。

**必ずJSON形式で返してください：**
{
  "x": プラカードの中心X座標（通常512付近）,
  "y": プラカードの中心Y座標（450-600の範囲、白い看板の中心）,
  "fontSize": 適切なフォントサイズ（40-60）
}

例: {"x": 512, "y": 520, "fontSize": 48}`;

    const apiUrl = `${VISION_ENDPOINT}/openai/deployments/${VISION_DEPLOYMENT}/chat/completions?api-version=${VISION_API_VERSION}`;

    console.log("[Vision] Calling API with image...");

    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": VISION_KEY,
      },
      body: JSON.stringify({
        messages: [
          {
            role: "user",
            content: [
              { type: "text", text: prompt },
              {
                type: "image_url",
                image_url: {
                  url: imageUrl,
                  detail: "low",
                },
              },
            ],
          },
        ],
        response_format: { type: "json_object" },
        max_tokens: 300,
        temperature: 0.3,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("[Vision] API Error:", response.status, errorText);
      return null;
    }

    const result = await response.json();
    const content = result.choices?.[0]?.message?.content;

    if (!content) {
      console.error("[Vision] No content in response");
      return null;
    }

    const coords = JSON.parse(content);
    console.log("[Vision] Parsed coordinates:", coords);

    // ★ Y座標の範囲チェックと補正
    let y = coords.y || 520;
    if (y > 650 || y < 400) {
      console.warn(
        `[Vision] Y coordinate ${y} seems wrong, adjusting to 520`
      );
      y = 520;
    }

    return {
      x: coords.x || 512,
      y,
      fontSize: coords.fontSize || 48,
    };
  } catch (e: any) {
    console.error("[Vision] Detection failed:", e);
    return null;
  }
}
