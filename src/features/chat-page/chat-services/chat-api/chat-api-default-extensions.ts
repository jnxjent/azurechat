// src/features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts
"use server";
import "server-only";

import { OpenAIDALLEInstance } from "@/features/common/services/openai"; // ★ 追加

import { ServerActionResponse } from "@/features/common/server-action-response";
import { uniqueId } from "@/features/common/util";
import { GetImageUrl, UploadImageToStore } from "../chat-image-service";
import { ChatThreadModel } from "../models";

import {
  buildSendOptionsFromMode,
  canonicalizeMode,
  type ThinkingModeInput,
} from "@/features/chat-page/chat-services/chat-api/reasoning-utils";

type ThinkingModeAPI = "normal" | "thinking" | "fast";

function normalizeThinkingMode(
  input?: ThinkingModeAPI | ThinkingModeInput
): ThinkingModeAPI {
  const c = canonicalizeMode(input as any);
  return c as ThinkingModeAPI;
}

function buildExternalImageUrl(threadId: string, fileName: string): string {
  const publicBase = process.env.NEXT_PUBLIC_IMAGE_URL;
  if (publicBase) {
    const base = publicBase.replace(/\/+$/, "");
    return `${base}/?t=${threadId}&img=${fileName}`;
  }

  const nextAuth = process.env.NEXTAUTH_URL;
  if (nextAuth) {
    const base = nextAuth.replace(/\/+$/, "");
    return `${base}/api/images/?t=${threadId}&img=${fileName}`;
  }

  return GetImageUrl(threadId, fileName);
}

type StyleParams = {
  font?: string;
  size?: "small" | "medium" | "large" | "xlarge";
  sizeAdjust?: "larger" | "smaller";
  align?: "left" | "center" | "right";
  vAlign?: "top" | "middle" | "bottom";
  bottomMargin?: number;
  offsetX?: number;
  offsetY?: number;
  color?: string;
};

type TextLayout = {
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

const lastTextLayoutByThread = new Map<string, TextLayout>();

function parseStyleHint(styleHint?: string): StyleParams {
  if (!styleHint) return {};
  const s = styleHint.replace(/\s+/g, "").toLowerCase();

  const p: StyleParams = {};

  if (s.includes("特大") || s.includes("ドーン") || s.includes("めちゃ大")) {
    p.size = "xlarge";
  } else if (
    s.includes("大きめ") ||
    s.includes("大きく") ||
    s.includes("大きい")
  ) {
    p.size = "large";
  } else if (s.includes("小さめ") || s.includes("小さい") || s.includes("控えめ")) {
    p.size = "small";
  } else if (s.includes("普通") || s.includes("標準")) {
    p.size = "medium";
  }

  if (
    s.includes("もう少し大きく") ||
    s.includes("もうちょっと大きく") ||
    s.includes("もっと大きく") ||
    s.includes("さらに大きく") ||
    s.includes("ちょい大きく")
  ) {
    p.sizeAdjust = "larger";
  } else if (
    s.includes("もう少し小さく") ||
    s.includes("もうちょっと小さく") ||
    s.includes("もっと小さく") ||
    s.includes("さらに小さく") ||
    s.includes("ちょい小さく")
  ) {
    p.sizeAdjust = "smaller";
  }

  if (
    s.includes("一番下") ||
    s.includes("最下部") ||
    s.includes("フッター") ||
    s.includes("下部") ||
    s.includes("下の方") ||
    s.includes("下側")
  ) {
    p.vAlign = "bottom";
    p.bottomMargin = 80;
  }

  if (
    s.includes("一番上") ||
    s.includes("最上部") ||
    s.includes("上端") ||
    s.includes("画面の上") ||
    s.includes("上部") ||
    s.includes("上の方") ||
    s.includes("上側")
  ) {
    p.vAlign = "top";
  }

  if (
    !p.vAlign &&
    (s.includes("真ん中") ||
      s.includes("センター") ||
      s.includes("中心") ||
      s.includes("中央"))
  ) {
    p.vAlign = "middle";
  }

  if (s.includes("左上")) { p.align = "left"; p.vAlign = "top"; }
  if (s.includes("右上")) { p.align = "right"; p.vAlign = "top"; }
  if (s.includes("左下")) { p.align = "left"; p.vAlign = "bottom"; p.bottomMargin = 80; }
  if (s.includes("右下")) { p.align = "right"; p.vAlign = "bottom"; p.bottomMargin = 80; }

  if (!p.align) {
    if (
      s.includes("左寄せ") ||
      s.includes("左側") ||
      s.includes("左端") ||
      (s.includes("左") && !s.includes("中央") && !s.includes("真ん中"))
    ) {
      p.align = "left";
    } else if (
      s.includes("右寄せ") ||
      s.includes("右側") ||
      s.includes("右端") ||
      (s.includes("右") && !s.includes("中央") && !s.includes("真ん中"))
    ) {
      p.align = "right";
    } else if (
      s.includes("中央") ||
      s.includes("真ん中") ||
      s.includes("センター") ||
      s.includes("中寄せ")
    ) {
      p.align = "center";
    }
  }

  if (s.includes("少し右") || s.includes("ちょい右") || s.includes("やや右")) p.offsetX = (p.offsetX ?? 0) + 80;
  if (s.includes("少し左") || s.includes("ちょい左") || s.includes("やや左")) p.offsetX = (p.offsetX ?? 0) - 80;
  if (s.includes("少し上") || s.includes("ちょい上") || s.includes("やや上")) p.offsetY = (p.offsetY ?? 0) - 60;
  if (s.includes("少し下") || s.includes("ちょい下") || s.includes("やや下")) p.offsetY = (p.offsetY ?? 0) + 60;

  if (s.includes("→") || s.includes("➡") || s.includes("➜") || s.includes("右矢印")) p.offsetX = (p.offsetX ?? 0) + 80;
  if (s.includes("←") || s.includes("⬅") || s.includes("左矢印")) p.offsetX = (p.offsetX ?? 0) - 80;
  if (s.includes("↑") || s.includes("⬆") || s.includes("上矢印")) p.offsetY = (p.offsetY ?? 0) - 60;
  if (s.includes("↓") || s.includes("⬇") || s.includes("下矢印")) p.offsetY = (p.offsetY ?? 0) + 60;

  if (s.includes("メイリオ")) p.font = "Meiryo";
  if (s.includes("游ゴシック") || s.includes("游ｺﾞｼｯｸ")) p.font = "Yu Gothic";
  if (s.includes("ゴシック")) p.font = "Yu Gothic";
  if (s.includes("明朝")) p.font = "Yu Mincho";
  if (s.includes("手書き") || s.includes("手書き風")) p.font = "Comic Sans MS";

  if (s.includes("白文字") || s.includes("白")) p.color = "#ffffff";
  if (s.includes("黒文字") || s.includes("黒")) p.color = "#000000";
  if (s.includes("赤文字") || s.includes("赤")) p.color = "red";
  if (s.includes("青文字") || s.includes("青")) p.color = "blue";
  if (s.includes("黄色") || s.includes("黄")) p.color = "yellow";

  return p;
}

export const GetDefaultExtensions = async (props: {
  chatThread: ChatThreadModel;
  userMessage: string;
  signal: AbortSignal;
  mode?: ThinkingModeAPI;
}): Promise<ServerActionResponse<Array<any>>> => {
  const defaultExtensions: Array<any> = [];

  const currentMode = normalizeThinkingMode(props.mode ?? "normal");
  const modeOpts = buildSendOptionsFromMode(currentMode);

  console.log("🧠 Reasoning Mode Applied:", {
    mode: currentMode,
    reasoning_effort: modeOpts.reasoning_effort,
    temperature: modeOpts.temperature,
  });

  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeCreateImage(
          args,
          props.chatThread,
          props.userMessage,
          props.signal,
          modeOpts
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          prompt: { type: "string" },
          text: { type: "string" },
          size: {
            type: "string",
            enum: ["1024x1024", "1024x1792", "1792x1024"],
          },
        },
        required: ["prompt"],
      },
      description:
        "Use this tool ONLY when user clearly asks for a NEW image to be created. " +
        "If user wants to MODIFY or add text to an ALREADY GENERATED image, you MUST NOT call this tool. " +
        "Instead, call add_text_to_existing_image with the previous image URL." +
        "After this tool returns a url, you MUST display the image using Markdown image syntax: ![image](url). Never output the URL as plain text.",
      name: "create_img",
    },
  });

  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeAddTextToExistingImage(
          args,
          props.chatThread,
          props.userMessage,
          props.signal,
          modeOpts
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          imageUrl: {
            type: "string",
            description:
              "URL of the existing image. If the user says 'this image', use the URL that was returned previously (for example from create_img).",
          },
          text: {
            type: "string",
            description:
              "Japanese text to overlay on the image. " +
              "CRITICAL: If the user is ONLY adjusting position, size, or color (words like '右に', 'もう少し大きく', '赤色に'), " +
              "you MUST use the EXACT same text from the previous image. Do NOT shorten, modify, or change the text content in any way.",
          },
          styleHint: {
            type: "string",
            description:
              "Natural language hint for font size, color, position such as '大きめの白文字で、下部中央に', '少し上に', '➡ で少し右へ', 'もう少し大きく', etc.",
          },
          font: {
            type: "string",
            description: "Font family name if explicitly requested (e.g., 'Meiryo').",
          },
          color: {
            type: "string",
            description: "Text color (e.g., 'white', '#ffffff').",
          },
          size: {
            type: "string",
            description: "Rough size hint like 'small', 'medium', 'large'.",
          },
          offsetX: {
            type: "number",
            description:
              "Horizontal offset in pixels. Positive moves text to the right, negative to the left.",
          },
          offsetY: {
            type: "number",
            description:
              "Vertical offset in pixels. Positive moves text downward, negative upward.",
          },
        },
        required: ["imageUrl", "text"],
      },
      description:
        "Use this tool when the user wants to add or adjust text on an EXISTING image, for example 'この絵に 2026 謹賀新年 と入れて' or 'もう少し下に', 'そこから➡で右に', 'もう少し大きく'. " +
        "CRITICAL RULE: When the user is ONLY requesting position/size/color adjustments, " +
        "you MUST preserve the EXACT text from the previous image without any modifications.",
      name: "add_text_to_existing_image",
    },
  });

  return { status: "OK", response: defaultExtensions };
};

// ---------------- 画像生成（NEW image 用）★ OpenAIDALLEInstance方式に変更 ----------------
async function executeCreateImage(
  args: { prompt: string; text?: string; size?: string },
  chatThread: ChatThreadModel,
  userMessage: string,
  signal: AbortSignal,
  modeOpts?: {
    reasoning_effort?: "low" | "medium" | "high";
    temperature?: number;
  }
) {
  const prompt = (args?.prompt || "").trim();

  console.log("createImage called with prompt:", prompt);

  if (!prompt) return "No prompt provided";
  if (prompt.length >= 4000)
    return "Prompt is too long, it must be less than 4000 characters";

  // ★ OpenAIDALLEInstance() 方式に変更（gpt-image-1.5対応）
  const openAI = OpenAIDALLEInstance();

  let response;
  try {
    response = await openAI.images.generate(
      { model: "gpt-image-1.5", prompt },
      { signal }
    );
  } catch (error) {
    console.error("🔴 error while calling Azure image gen:\n", error);
    return { error: "There was an error creating the image: " + error };
  }

  if (!response.data?.[0]?.b64_json) {
    return { error: "Invalid API response: no b64_json." };
  }

  try {
    const imageName = `${uniqueId()}.png`;
    const buffer = Buffer.from(response.data[0].b64_json, "base64");

    await UploadImageToStore(chatThread.id, imageName, buffer);
    await UploadImageToStore(chatThread.id, "__base__.png", buffer);

    lastTextLayoutByThread.delete(chatThread.id);
    console.log("🗑️ Cleared text layout for thread:", chatThread.id);

    const baseImageUrl = buildExternalImageUrl(chatThread.id, imageName);
    return { revised_prompt: prompt, url: baseImageUrl };
  } catch (error) {
    console.error("🔴 error while storing image:\n", error);
    return { error: "There was an error storing the image: " + error };
  }
}

// ---------------- 既存画像への文字追加（EDIT 用・Vision 不使用） ----------------
async function executeAddTextToExistingImage(
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
  chatThread: ChatThreadModel,
  userMessage: string,
  signal: AbortSignal,
  modeOpts?: {
    reasoning_effort?: "low" | "medium" | "high";
    temperature?: number;
  }
) {
  const explicitUrl = (args?.imageUrl || "").trim();
  let text = (args?.text || "").trim();
  const styleHint = (args?.styleHint || "").trim();

  const baseImageUrl = buildExternalImageUrl(chatThread.id, "__base__.png");

  console.log("🗺️ lastTextLayoutByThread MAP状態:", {
    threadId: chatThread.id,
    hasEntry: lastTextLayoutByThread.has(chatThread.id),
    mapSize: lastTextLayoutByThread.size,
    allKeys: Array.from(lastTextLayoutByThread.keys()),
    currentValue: lastTextLayoutByThread.get(chatThread.id),
  });

  console.log("🖋 add_text_to_existing_image called:", {
    passedImageUrl: explicitUrl,
    usedBaseImageUrl: baseImageUrl,
    text,
    styleHint,
    argsOffsetX: args?.offsetX,
    argsOffsetY: args?.offsetY,
  });

  if (!text) {
    return { error: "text is required for add_text_to_existing_image." };
  }

  const hintSource = styleHint || userMessage || "";
  const parsed = parseStyleHint(hintSource);

  console.log("🔍 parsed style hint:", parsed);

  const last = lastTextLayoutByThread.get(chatThread.id);
  console.log("📍 last layout from Map:", last);

  if (last?.text && text !== last.text) {
    console.warn("⚠️ Text content changed:", {
      previous: last.text,
      current: text,
      userMessage,
    });

    const lowerMsg = (userMessage || "").toLowerCase();
    const isExplicitChange =
      lowerMsg.includes("変更") || lowerMsg.includes("変える") || lowerMsg.includes("書き換え");

    if (!isExplicitChange) {
      console.warn("⚠️⚠️ Text changed without explicit request. Using previous text.");
      text = last.text;
    }
  }

  const align: "left" | "center" | "right" =
    parsed.align !== undefined ? parsed.align : last?.align ?? "center";

  const vAlign: "top" | "middle" | "bottom" =
    parsed.vAlign !== undefined ? parsed.vAlign : last?.vAlign ?? "middle";

  console.log("✅ resolved align/vAlign:", { align, vAlign });

  let size: "small" | "medium" | "large" | "xlarge" =
    (args.size as any) ?? parsed.size ?? last?.size ?? "large";

  if (parsed.sizeAdjust === "larger") {
    const sizeOrder: Array<"small" | "medium" | "large" | "xlarge"> = ["small", "medium", "large", "xlarge"];
    const currentIndex = sizeOrder.indexOf(size);
    if (currentIndex >= 0 && currentIndex < sizeOrder.length - 1) {
      const oldSize = size;
      size = sizeOrder[currentIndex + 1];
      console.log(`📏 Size adjusted larger: ${oldSize} → ${size}`);
    }
  } else if (parsed.sizeAdjust === "smaller") {
    const sizeOrder: Array<"small" | "medium" | "large" | "xlarge"> = ["small", "medium", "large", "xlarge"];
    const currentIndex = sizeOrder.indexOf(size);
    if (currentIndex > 0) {
      const oldSize = size;
      size = sizeOrder[currentIndex - 1];
      console.log(`📏 Size adjusted smaller: ${oldSize} → ${size}`);
    }
  }

  const color = args.color ?? parsed.color ?? last?.color ?? "white";

  console.log("🎨 color resolution:", {
    argsColor: args.color,
    parsedColor: parsed.color,
    lastColor: last?.color,
    finalColor: color,
  });

  const fontHint = ((styleHint || "") + " " + (args.font || "") + " " + (parsed.font || "")).toLowerCase();

  let fontFamily: "gothic" | "mincho" | "meiryo" = last?.fontFamily ?? "gothic";

  if (fontHint.includes("明朝") || fontHint.includes("mincho") || fontHint.includes("serif")) {
    fontFamily = "mincho";
  } else if (fontHint.includes("メイリオ") || fontHint.includes("meiryo")) {
    fontFamily = "meiryo";
  } else if (fontHint.includes("ゴシック") || fontHint.includes("gothic")) {
    fontFamily = "gothic";
  }

  console.log("🔤 fontFamily resolution:", {
    fontHint,
    lastFontFamily: last?.fontFamily,
    finalFontFamily: fontFamily,
  });

  const lowerHintAll = (hintSource || "").toLowerCase();

  const boldOff = hintSource.includes("太字やめ") || hintSource.includes("太字解除") || hintSource.includes("太字をやめ") || hintSource.includes("太字を解除") || hintSource.includes("通常") || lowerHintAll.includes("not bold") || lowerHintAll.includes("no bold");
  const italicOff = hintSource.includes("斜体やめ") || hintSource.includes("斜体解除") || hintSource.includes("イタリックやめ") || hintSource.includes("イタリック解除") || hintSource.includes("斜体をやめ") || hintSource.includes("斜体を解除") || lowerHintAll.includes("not italic") || lowerHintAll.includes("no italic");
  const boldOn = hintSource.includes("太字") || hintSource.includes("ボールド") || lowerHintAll.includes("bold");
  const italicOn = hintSource.includes("イタリック") || hintSource.includes("斜体") || lowerHintAll.includes("italic");

  const bold = boldOff ? false : boldOn ? true : (last?.bold ?? false);
  const italic = italicOff ? false : italicOn ? true : (last?.italic ?? false);

  console.log("📝 bold/italic resolution:", {
    lastBold: last?.bold,
    lastItalic: last?.italic,
    finalBold: bold,
    finalItalic: italic,
  });

  const positionSpecified =
    parsed.align !== undefined ||
    parsed.vAlign !== undefined ||
    /左上|右上|左下|右下|一番上|一番下|中央|真ん中|センター|上部|下部/.test(hintSource);

  const deltaOffsetX = (parsed.offsetX ?? 0) + (typeof args.offsetX === "number" ? args.offsetX : 0);
  const deltaOffsetY = (parsed.offsetY ?? 0) + (typeof args.offsetY === "number" ? args.offsetY : 0);

  const baseOffsetX = positionSpecified ? 0 : (last?.offsetX ?? 0);
  const baseOffsetY = positionSpecified ? 0 : (last?.offsetY ?? 0);

  const offsetX = baseOffsetX + deltaOffsetX;
  const offsetY = baseOffsetY + deltaOffsetY;

  console.log("📐 offset calculation:", {
    positionSpecified,
    baseOffsetX,
    baseOffsetY,
    parsedOffsetX: parsed.offsetX,
    parsedOffsetY: parsed.offsetY,
    argsOffsetX: args.offsetX,
    argsOffsetY: args.offsetY,
    deltaOffsetX,
    deltaOffsetY,
    finalOffsetX: offsetX,
    finalOffsetY: offsetY,
  });

  const bottomMargin = parsed.bottomMargin;

  lastTextLayoutByThread.set(chatThread.id, {
    align,
    vAlign,
    offsetX,
    offsetY,
    size,
    text,
    color,
    fontFamily,
    bold,
    italic,
  });

  console.log("💾 saved to Map:", {
    threadId: chatThread.id,
    saved: lastTextLayoutByThread.get(chatThread.id),
    mapSizeAfter: lastTextLayoutByThread.size,
  });

  const baseUrl =
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000");

  const genImageBase = baseUrl.replace(/\/+$/, "");
  console.log("[gen-image] base URL for overlay:", genImageBase);
  console.log("[gen-image] resolved style params:", { align, vAlign, size, color, fontFamily, bold, italic, offsetX, offsetY, bottomMargin });

  try {
    const resp = await fetch(`${genImageBase}/api/gen-image`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        imageUrl: baseImageUrl,
        text,
        align,
        vAlign,
        size,
        color,
        offsetX,
        offsetY,
        bottomMargin,
        autoDetectPlacard: false,
        fontFamily,
        bold,
        italic,
      }),
    });

    if (!resp.ok) {
      const t = await resp.text().catch(() => "");
      console.error("🔴 /api/gen-image failed in edit:", resp.status, t);
      return { error: `Text overlay failed: HTTP ${resp.status}` };
    }

    const result = await resp.json();
    const generatedPath = result?.imageUrl as string | undefined;

    if (!generatedPath) {
      console.error("🔴 gen-image edit returned no imageUrl");
      return { error: "gen-image edit returned no imageUrl" };
    }

    const fs = require("fs");
    const path = require("path");
    const finalImageName = `${uniqueId()}.png`;
    const finalImagePath = path.join(
      process.cwd(),
      "public",
      generatedPath.startsWith("/") ? generatedPath.slice(1) : generatedPath
    );
    const finalImageBuffer = fs.readFileSync(finalImagePath);

    await UploadImageToStore(chatThread.id, finalImageName, finalImageBuffer);

    const finalImageUrl = buildExternalImageUrl(chatThread.id, finalImageName);

    return { revised_prompt: text, url: finalImageUrl };
  } catch (err) {
    console.error("🔴 error in executeAddTextToExistingImage (simple):", err);
    return { error: "There was an error adding text to the existing image: " + err };
  }
}