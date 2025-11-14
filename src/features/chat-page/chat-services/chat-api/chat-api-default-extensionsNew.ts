// src/features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts
"use server";
import "server-only";

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

/** standard を normal へ、その他はそのまま（保険） */
function normalizeThinkingMode(input?: ThinkingModeAPI | ThinkingModeInput): ThinkingModeAPI {
  const c = canonicalizeMode(input as any);
  return c as ThinkingModeAPI;
}

export const GetDefaultExtensions = async (props: {
  chatThread: ChatThreadModel;
  userMessage: string;
  signal: AbortSignal;
  mode?: ThinkingModeAPI; // "normal" | "thinking" | "fast"
}): Promise<ServerActionResponse<Array<any>>> => {
  const defaultExtensions: Array<any> = [];

  const currentMode = normalizeThinkingMode(props.mode ?? "normal");
  const modeOpts = buildSendOptionsFromMode(currentMode);

  console.log("🧠 Reasoning Mode Applied:", {
    mode: currentMode, // normal | thinking | fast
    reasoning_effort: modeOpts.reasoning_effort,
    temperature: modeOpts.temperature,
  });

  // ★ 画像生成ツール（Vision API対応版）
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeCreateImage(
          args,
          props.chatThread.id,
          props.userMessage,
          props.signal,
          modeOpts
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          prompt: { type: "string" },
          text: { type: "string" }, // ★ 追加：プラカードに入れるテキスト
          size: { type: "string", enum: ["1024x1024", "1024x1792", "1792x1024"] },
        },
        required: ["prompt"],
      },
      description:
        "Use this tool ONLY when the user explicitly asks to create an image. If the prompt mentions placard, sign, board, or banner, you should also extract the text to display on it and pass it as 'text' parameter. Call it at most once per user request.",
      name: "create_img",
    },
  });

  return { status: "OK", response: defaultExtensions };
};

// ---------------- 画像生成（Vision API対応版） ----------------
async function executeCreateImage(
  args: { prompt: string; text?: string; size?: string },
  threadId: string,
  userMessage: string,
  signal: AbortSignal,
  modeOpts?: { reasoning_effort?: "low" | "medium" | "high"; temperature?: number }
) {
  const prompt = (args?.prompt || "").trim();
  const text = (args?.text || "").trim(); // ★ プラカードに入れるテキスト
  const size = (args?.size || "1024x1024").trim();

  console.log("createImage called with prompt:", prompt);
  console.log("createImage text for placard:", text || "(none)");
  console.log("🧩 reasoning_effort in request:", modeOpts?.reasoning_effort || "none");

  if (!prompt) return "No prompt provided";
  if (prompt.length >= 4000) return "Prompt is too long, it must be less than 4000 characters";

  const endpointRaw = process.env.AZURE_OPENAI_ENDPOINT || "";
  const endpoint = endpointRaw.replace(/\/+$/, "");
  const apiKey = process.env.AZURE_OPENAI_API_KEY || "";
  const deployment = process.env.AZURE_OPENAI_IMAGE_DEPLOYMENT || "";
  const apiVersion = process.env.AZURE_OPENAI_API_VERSION || "2025-04-01-preview";

  if (!endpoint || !/^https:\/\/.+\.openai\.azure\.com$/i.test(endpoint)) {
    return { error: "Image generation is not configured: invalid AZURE_OPENAI_ENDPOINT." };
  }
  if (!apiKey) return { error: "Image generation is not configured: missing AZURE_OPENAI_API_KEY." };
  if (!deployment) return { error: "Image generation is not configured: missing AZURE_OPENAI_IMAGE_DEPLOYMENT." };

  // ★ Step 1: Azure OpenAI DALL-Eで画像生成（テキストなし）
  const imageGenUrl = `${endpoint}/openai/deployments/${encodeURIComponent(
    deployment
  )}/images/generations?api-version=${encodeURIComponent(apiVersion)}`;

  let json: any;
  try {
    const res = await fetch(imageGenUrl, {
      method: "POST",
      headers: { "api-key": apiKey, "Content-Type": "application/json" },
      body: JSON.stringify({
        prompt: prompt + (text ? ". プラカードには何も書かれていない。文字は入れない。" : ""),
        n: 1,
        size,
        response_format: "b64_json",
        reasoning_effort: modeOpts?.reasoning_effort,
        temperature: modeOpts?.temperature,
      }),
      signal,
      cache: "no-store",
    });

    const responseText = await res.text();
    if (!res.ok) {
      return { error: `There was an error creating the image: HTTP ${res.status}.` };
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

  if (!b64 && !urlDirect) return { error: "Invalid API response: no data[0].b64_json/url." };

  try {
    // ★ まず画像を保存
    let baseImageUrl: string;
    
    if (b64) {
      const imageName = `${uniqueId()}.png`;
      await UploadImageToStore(threadId, imageName, Buffer.from(b64, "base64"));
      baseImageUrl = GetImageUrl(threadId, imageName);
    } else {
      baseImageUrl = urlDirect!;
    }

    // ★ Step 2: テキストがある場合、Vision APIでプラカードにテキスト配置
    if (text) {
      console.log("🎨 Using Vision API to add text to placard...");
      
      try {
        const visionResponse = await fetch(`${process.env.NEXTAUTH_URL || 'http://localhost:3000'}/api/gen-image-vision`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            imageUrl: baseImageUrl,
            text: text
          }),
          signal,
        });

        if (visionResponse.ok) {
          const visionResult = await visionResponse.json();
          
          if (visionResult.imageUrl) {
            console.log("✅ Vision API successfully added text");
            // /generated/xxx.png を Azure Storageにアップロード
            const finalImageName = `${uniqueId()}.png`;
            
            // /generated/ から読み込んでAzure Storageに保存
            const fs = require('fs');
            const path = require('path');
            const finalImagePath = path.join(process.cwd(), 'public', visionResult.imageUrl);
            const finalImageBuffer = fs.readFileSync(finalImagePath);
            
            await UploadImageToStore(threadId, finalImageName, finalImageBuffer);
            
            const finalImageUrl = GetImageUrl(threadId, finalImageName);
            
            // ★ Markdown形式でクリック可能なリンクを返す
            return `画像に「${text}」を入れて作成しました。\n\n![画像](${finalImageUrl})\n\n[画像を開く](${finalImageUrl})`;
          }
        }
        
        console.warn("⚠️ Vision API failed, returning base image without text");
      } catch (visionError) {
        console.error("🔴 Vision API error:", visionError);
        console.warn("⚠️ Falling back to base image without text");
      }
    }

    // ★ テキストなし、またはVision API失敗時 - Markdown形式で返す
    return `画像を1枚作成しました。\n\n![画像](${baseImageUrl})\n\n[画像を開く](${baseImageUrl})${text ? '\n\nこのままでも使えますが、プラカードに入れたい日本語の文字があれば、自然な言い方で指定してください。例:「がんばろう」と入れて。\n後から文字を入れることもできます。' : ''}`;
    
  } catch (error) {
    console.error("🔴 error while storing image:\n", error);
    return { error: "There was an error storing the image: " + error };
  }
}