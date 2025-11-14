// features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts
"use server";
import "server-only";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { uniqueId } from "@/features/common/util";
import { GetImageUrl, UploadImageToStore } from "../chat-image-service";
import { ChatThreadModel } from "../models";

// ✅ reasoning-effort 関連（fast/normal/standard/auto/thinking をそのまま受ける）
import {
  buildSendOptionsFromMode,
  type ThinkingModeInput,
} from "@/features/chat-page/chat-services/chat-api/reasoning-utils";

/**
 * デフォルト拡張（画像生成など）。reasoning_effort はここで一元適用。
 */
export const GetDefaultExtensions = async (props: {
  chatThread: ChatThreadModel;
  userMessage: string;
  signal: AbortSignal;
  /** UIから渡される現在モード: "standard" | "auto" | "fast" | "thinking" | （互換で "normal" も許容） */
  mode?: ThinkingModeInput;
}): Promise<ServerActionResponse<Array<any>>> => {
  const defaultExtensions: Array<any> = [];

  // =========================================================
  // 💬 reasoning_effort（熟考度）をモードに応じて組み込み
  //    - 余計な正規化はしない（utils 側に委譲）
  //    - auto を動的化するため userMessage を渡す
  // =========================================================
  const currentMode: ThinkingModeInput = props.mode ?? "auto";
  const modeOpts = buildSendOptionsFromMode(currentMode, props.userMessage);
  // modeOpts: { reasoning_effort: "low"|"medium"|"high", temperature: number }

  console.log("🧠 Reasoning Mode Applied:", {
    mode: currentMode, // "auto" | "fast" | "normal" | "standard" | "thinking"
    reasoning_effort: modeOpts.reasoning_effort,
    temperature: modeOpts.temperature,
  });

  // =========================================================
  // 🖼️ Image creation tool
  // =========================================================
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeCreateImage(
          args,
          props.chatThread.id,
          props.userMessage,
          props.signal,
          modeOpts // reasoning_effort / temperature をそのまま渡す（将来拡張にも備える）
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          prompt: { type: "string" },
          size: { type: "string", enum: ["1024x1024", "1024x1792", "1792x1024"] },
        },
        required: ["prompt"],
      },
      description:
        "Use this tool ONLY when the user explicitly asks to create an image. Call it at most once per user request.",
      name: "create_img",
    },
  });

  return {
    status: "OK",
    response: defaultExtensions,
  };
};

// =========================================================
// 🖼️ Azure OpenAI 画像生成エンドポイント呼び出し
// =========================================================
async function executeCreateImage(
  args: { prompt: string; size?: string },
  threadId: string,
  userMessage: string,
  signal: AbortSignal,
  modeOpts?: { reasoning_effort?: "low" | "medium" | "high"; temperature?: number }
) {
  const prompt = (args?.prompt || "").trim();
  const size = (args?.size || "1024x1024").trim();

  console.log("createImage called with prompt:", prompt);
  console.log("🧩 reasoning_effort in request:", modeOpts?.reasoning_effort || "none");

  if (!prompt) {
    return "No prompt provided";
  }
  if (prompt.length >= 4000) {
    return "Prompt is too long, it must be less than 4000 characters";
  }

  // ---- 環境変数の読み取りと検証 ----
  const endpointRaw = process.env.AZURE_OPENAI_ENDPOINT || "";
  const endpoint = endpointRaw.replace(/\/+$/, "");
  const apiKey = process.env.AZURE_OPENAI_API_KEY || "";
  const deployment = process.env.AZURE_OPENAI_IMAGE_DEPLOYMENT || "";
  const apiVersion = process.env.AZURE_OPENAI_API_VERSION || "2025-04-01-preview";

  console.debug("IMG_DEBUG", {
    endpoint,
    deployment,
    apiVersion,
    hasAzureKey: !!apiKey,
    hasOpenAIKey: !!process.env.OPENAI_API_KEY,
    userMessagePreview: (userMessage || "").slice(0, 24),
  });

  if (!endpoint || !/^https:\/\/.+\.openai\.azure\.com$/i.test(endpoint)) {
    return {
      error:
        "Image generation is not configured: invalid AZURE_OPENAI_ENDPOINT (expected https://<resource>.openai.azure.com).",
    };
  }
  if (!apiKey) {
    return { error: "Image generation is not configured: missing AZURE_OPENAI_API_KEY." };
  }
  if (!deployment) {
    return { error: "Image generation is not configured: missing AZURE_OPENAI_IMAGE_DEPLOYMENT." };
  }

  const url = `${endpoint}/openai/deployments/${encodeURIComponent(
    deployment
  )}/images/generations?api-version=${encodeURIComponent(apiVersion)}`;

  console.debug("IMG_CALL", { url, deployment, apiVersion, size });

  // ---- Azure 画像生成 呼び出し ----
  let json: any;
  try {
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "api-key": apiKey,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        prompt,
        n: 1,
        size,
        response_format: "b64_json",
        // ※ 画像モデルでは無視される可能性が高いが、デバッグ・将来拡張用に付与
        reasoning_effort: modeOpts?.reasoning_effort,
        temperature: modeOpts?.temperature,
      }),
      signal,
      cache: "no-store",
    });

    const text = await res.text();
    const preview = text.slice(0, 512);
    console.debug("IMG_RES", { status: res.status, ok: res.ok, preview });

    if (!res.ok) {
      return {
        error:
          `There was an error creating the image: HTTP ${res.status}. ` +
          `Please verify endpoint/deployment/api-version/key.`,
      };
    }

    try {
      json = JSON.parse(text);
    } catch {
      return { error: "Invalid JSON response from Azure." };
    }
  } catch (error) {
    console.error("🔴 error while calling Azure image gen:\n", error);
    return { error: "There was an error creating the image: " + error };
  }

  // ---- 正常応答の検証 ----
  const data0 = json?.data?.[0];
  const b64 = data0?.b64_json as string | undefined;
  const urlDirect = data0?.url as string | undefined;

  if (!b64 && !urlDirect) {
    return { error: "Invalid API response: no data[0].b64_json/url." };
  }

  // ---- ストレージへ保存 ----
  try {
    if (b64) {
      const imageName = `${uniqueId()}.png`;
      await UploadImageToStore(threadId, imageName, Buffer.from(b64, "base64"));

      return {
        revised_prompt: data0?.revised_prompt,
        url: GetImageUrl(threadId, imageName),
      };
    } else {
      return {
        revised_prompt: data0?.revised_prompt,
        url: urlDirect,
      };
    }
  } catch (error) {
    console.error("🔴 error while storing image:\n", error);
    return { error: "There was an error storing the image: " + error };
  }
}
