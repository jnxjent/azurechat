"use server";
import "server-only";

import { uniqueId } from "@/features/common/util";
import { GetImageUrl, UploadImageToStore } from "@/features/chat-page/chat-services/chat-image-service";

// gpt-image-2 supports: 1024x1024, 1024x1536, 1536x1024
function compatSizeStr(size: string): string {
  if (size === "1792x1024") return "1536x1024";
  if (size === "1024x1792") return "1024x1536";
  return size;
}

export async function executeCreateImage(
  args: { prompt: string; text?: string; size?: string },
  threadId: string,
  userMessage: string,
  signal: AbortSignal
) {
  const prompt = (args?.prompt || "").trim();
  const text = (args?.text || "").trim(); // プラカードに入れるテキスト
  const size = compatSizeStr((args?.size || "1024x1024").trim());

  console.log("createImage called with prompt:", prompt);
  console.log("createImage text for placard:", text || "(none)");

  if (!prompt) return "No prompt provided";
  if (prompt.length >= 4000)
    return "Prompt is too long, it must be less than 4000 characters";

  const endpointRaw = process.env.AZURE_OPENAI_ENDPOINT || "";
  const endpoint = endpointRaw.replace(/\/+$/, "");
  const apiKey = process.env.AZURE_OPENAI_API_KEY || "";
  const deployment = process.env.AZURE_OPENAI_IMAGE_DEPLOYMENT || "";
  const apiVersion =
    process.env.AZURE_OPENAI_IMAGE_API_VERSION || "2025-04-01-preview";

  if (!endpoint || !/^https:\/\/.+\.openai\.azure\.com$/i.test(endpoint)) {
    return {
      error:
        "Image generation is not configured: invalid AZURE_OPENAI_ENDPOINT.",
    };
  }
  if (!apiKey)
    return {
      error:
        "Image generation is not configured: missing AZURE_OPENAI_API_KEY.",
    };
  if (!deployment)
    return {
      error:
        "Image generation is not configured: missing AZURE_OPENAI_IMAGE_DEPLOYMENT.",
    };

  // Step 1: Azure OpenAI DALL-E で画像生成（テキストなし）
  const imageGenUrl = `${endpoint}/openai/deployments/${encodeURIComponent(
    deployment
  )}/images/generations?api-version=${encodeURIComponent(apiVersion)}`;

  let json: any;
  try {
    const res = await fetch(imageGenUrl, {
      method: "POST",
      headers: { "api-key": apiKey, "Content-Type": "application/json" },
      body: JSON.stringify({
        prompt:
          prompt +
          (text ? ". プラカードには何も書かれていない。文字は入れない。" : ""),
        n: 1,
        size,
      }),
      signal,
      cache: "no-store",
    });

    const responseText = await res.text();
    if (!res.ok) {
      let errCode = "";
      let errMsg = "";
      try { const j = JSON.parse(responseText); errCode = j?.error?.code ?? ""; errMsg = j?.error?.message ?? ""; } catch {}
      console.error("[image/create] API error", {
        status: res.status, code: errCode, message: errMsg,
        deployment, apiVersion, size,
      });
      return {
        error: `There was an error creating the image: HTTP ${res.status}.`,
      };
    }
    try {
      json = JSON.parse(responseText);
    } catch {
      return { error: "Invalid JSON response from Azure." };
    }
  } catch (error) {
    console.error("🔴 error while calling Azure image gen:\n", error);
    return { error: "There was an error creating the image: " + error };
  }

  const data0 = json?.data?.[0];
  const b64 = data0?.b64_json as string | undefined;
  const urlDirect = data0?.url as string | undefined;

  if (!b64 && !urlDirect)
    return { error: "Invalid API response: no data[0].b64_json/url." };

  try {
    // まず画像を Blob に保存
    let baseImageUrl: string;

    if (b64) {
      const imageName = `${uniqueId()}.png`;
      await UploadImageToStore(threadId, imageName, Buffer.from(b64, "base64"));
      baseImageUrl = GetImageUrl(threadId, imageName);
    } else {
      baseImageUrl = urlDirect!;
    }

    // テキスト指定がある場合は Vision API で合成
    if (text) {
      console.log("🎨 Using Vision API to add text to placard...");

      const visionBaseUrl =
        process.env.VISION_API_BASE_URL ||
        (process.env.WEBSITE_HOSTNAME
          ? `https://${process.env.WEBSITE_HOSTNAME}`
          : process.env.NEXTAUTH_URL || "http://localhost:3000");

      console.log("[Vision] base URL for overlay:", visionBaseUrl);

      try {
        const visionResponse = await fetch(
          `${visionBaseUrl}/api/gen-image-vision`,
          {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              imageUrl: baseImageUrl,
              text: text,
              // 詳細なスタイル指定は add_text_to_existing_image 側で扱う想定
            }),
            signal,
          }
        );

        if (visionResponse.ok) {
          const visionResult = await visionResponse.json();
          const generatedPath = visionResult?.imageUrl as
            | string
            | undefined;

          if (generatedPath) {
            console.log("✅ Vision API successfully added text");

            const fs = require("fs");
            const path = require("path");
            const finalImageName = `${uniqueId()}.png`;
            const finalImagePath = path.join(
              process.cwd(),
              "public",
              generatedPath
            );
            const finalImageBuffer = fs.readFileSync(finalImagePath);

            await UploadImageToStore(
              threadId,
              finalImageName,
              finalImageBuffer
            );

            const finalImageUrl = GetImageUrl(threadId, finalImageName);

            return `画像に「${text}」を入れて作成しました。\n\n![画像](${finalImageUrl})\n\n[画像を開く](${finalImageUrl})`;
          }
        }

        console.warn(
          "⚠️ Vision API failed, returning base image without text"
        );
      } catch (visionError) {
        console.error("🔴 Vision API error:", visionError);
        console.warn("⚠️ Falling back to base image without text");
      }
    }

    // テキストなし・または Vision 失敗時
    return `画像を1枚作成しました。\n\n![画像](${baseImageUrl})\n\n[画像を開く](${baseImageUrl})${
      text
        ? '\n\nこのままでも使えますが、プラカードに入れたい日本語の文字があれば、自然な言い方で指定してください。例:「がんばろう」と入れて。\n後から文字を入れることもできます。'
        : ""
    }`;
  } catch (error) {
    console.error("🔴 error while storing image:\n", error);
    return { error: "There was an error storing the image: " + error };
  }
}
