"use server";
import "server-only";

import { uniqueId } from "@/features/common/util";
import { GetImageUrl, UploadImageToStore } from "@/features/chat-page/chat-services/chat-image-service";

export async function executeAddTextToExistingImage(
  args: {
    imageUrl: string;
    text: string;
    styleHint?: string;
    font?: string;
    color?: string;
    size?: string;
    offsetX?: number;
    offsetY?: number;
  },
  threadId: string,
  signal: AbortSignal
) {
  const imageUrl = (args?.imageUrl || "").trim();
  const text = (args?.text || "").trim();

  if (!imageUrl) {
    return { error: "imageUrl is required for add_text_to_existing_image." };
  }
  if (!text) {
    return { error: "text is required for add_text_to_existing_image." };
  }

  console.log("🖋 add_text_to_existing_image called:", {
    imageUrl,
    text,
    styleHint: args.styleHint || "(none)",
    font: args.font || "(default)",
    color: args.color || "(default)",
    size: args.size || "(default)",
    offsetX: args.offsetX ?? 0,
    offsetY: args.offsetY ?? 0,
  });

  const visionBaseUrl =
    process.env.VISION_API_BASE_URL ||
    (process.env.WEBSITE_HOSTNAME
      ? `https://${process.env.WEBSITE_HOSTNAME}`
      : process.env.NEXTAUTH_URL || "http://localhost:3000");

  console.log("[Vision] base URL for overlay (existing image):", visionBaseUrl);

  try {
    const res = await fetch(`${visionBaseUrl}/api/gen-image-vision`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        imageUrl,
        text,
        styleHint: args.styleHint ?? "",
        font: args.font ?? "Meiryo",    // デフォルト Meiryo
        color: args.color ?? "white",   // デフォルト白
        size: args.size ?? "large",     // デフォルトやや大きめ
        offsetX: args.offsetX ?? 0,
        offsetY: args.offsetY ?? 0,
      }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error(
        "🔴 Vision API returned non-200 in add_text_to_existing_image:",
        res.status,
        t
      );
      return {
        error: `Vision API error in add_text_to_existing_image: HTTP ${res.status}`,
      };
    }

    const json = await res.json();
    const generatedPath = json?.imageUrl as string | undefined;

    if (!generatedPath) {
      return {
        error:
          "Vision API response is missing 'imageUrl' in add_text_to_existing_image.",
      };
    }

    const fs = require("fs");
    const path = require("path");
    const finalImagePath = path.join(process.cwd(), "public", generatedPath);
    const finalImageBuffer = fs.readFileSync(finalImagePath);

    const finalImageName = `${uniqueId()}.png`;
    await UploadImageToStore(threadId, finalImageName, finalImageBuffer);
    const finalImageUrl = GetImageUrl(threadId, finalImageName);

    return `既存の画像に「${text}」というテキストを追加しました。\n\n![画像](${finalImageUrl})\n\n[画像を開く](${finalImageUrl})`;
  } catch (error) {
    console.error(
      "🔴 error while calling Vision API in add_text_to_existing_image:\n",
      error
    );
    return {
      error:
        "There was an error while adding text to the existing image: " + error,
    };
  }
}
