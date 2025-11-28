// src/features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts
"use server";
import "server-only";

import { ServerActionResponse } from "@/features/common/server-action-response";
import { uniqueId } from "@/features/common/util";
import {
  GetImageUrl,
  UploadImageToStore,
  // RegisterImageOnThread,  // ← もう使わないのでコメントアウトでOK
  // GetImageUrlFromThread, // ← 同上
} from "../chat-image-service";
import { ChatThreadModel } from "../models";

import {
  buildSendOptionsFromMode,
  canonicalizeMode,
  type ThinkingModeInput,
} from "@/features/chat-page/chat-services/chat-api/reasoning-utils";

type ThinkingModeAPI = "normal" | "thinking" | "fast";

/** standard を normal へ、その他はそのまま（保険） */
function normalizeThinkingMode(
  input?: ThinkingModeAPI | ThinkingModeInput
): ThinkingModeAPI {
  const c = canonicalizeMode(input as any);
  return c as ThinkingModeAPI;
}

/**
 * 画像URLを組み立てる共通ヘルパー
 * - NEXT_PUBLIC_IMAGE_URL があればそれを最優先（https://xxx.azurewebsites.net/api/images）
 * - なければ NEXTAUTH_URL + /api/images を使う
 * - どちらも無ければ、最後の保険として GetImageUrl() にフォールバック
 *
 * 生成される形式の例:
 *   https://.../api/images/?t=<threadId>&img=<fileName>
 */
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

  // 環境変数が未設定だった場合の最後の保険
  return GetImageUrl(threadId, fileName);
}

/* ------------------------------------------------------------------ */
/* NL スタイルヒント → パラメータ変換                                  */
/* ------------------------------------------------------------------ */

type StyleParams = {
  font?: string;
  size?: "small" | "medium" | "large" | "xlarge";
  align?: "left" | "center" | "right";
  vAlign?: "top" | "middle" | "bottom";
  bottomMargin?: number;
  offsetX?: number;
  offsetY?: number;
  color?: string;
};

function parseStyleHint(styleHint?: string): StyleParams {
  if (!styleHint) return {};
  // かなり雑でも良いので、まずは正規化
  const s = styleHint.replace(/\s+/g, "").toLowerCase();

  const p: StyleParams = {};

  // ---- サイズ系 ----
  if (s.includes("特大") || s.includes("ドーン") || s.includes("めちゃ大")) {
    p.size = "xlarge";
  } else if (s.includes("大きめ") || s.includes("大きく") || s.includes("大きい")) {
    p.size = "large";
  } else if (s.includes("小さめ") || s.includes("小さい") || s.includes("控えめ")) {
    p.size = "small";
  } else if (s.includes("普通") || s.includes("標準")) {
    p.size = "medium";
  }

  // ---- 垂直位置（下 / 上 / 真ん中）----
  // 1) まず bottom を決める
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

  // 2) 次に top を決める（「一番上の中央」のような場合も top を優先）
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

  // 3) 上下が指定されていない場合だけ中央扱い
  if (
    !p.vAlign &&
    (s.includes("真ん中") ||
      s.includes("中央") ||
      s.includes("センター") ||
      s.includes("中心"))
  ) {
    p.vAlign = "middle";
  }

  // ---- 水平位置（左 / 中央 / 右）----
  if (s.includes("左上") || s.includes("左下")) {
    // ４隅ショートカットでまとめて処理するのでここでは何もしない
  } else if (
    s.includes("左寄せ") ||
    s.includes("左側") ||
    s.includes("左端") ||
    s.includes("左")
  ) {
    p.align = "left";
  }

  if (s.includes("右上") || s.includes("右下")) {
    // ４隅ショートカットでまとめて処理するのでここでは何もしない
  } else if (
    s.includes("右寄せ") ||
    s.includes("右側") ||
    s.includes("右端") ||
    s.includes("右")
  ) {
    p.align = "right";
  }

  if (
    s.includes("中央") ||
    s.includes("真ん中") ||
    s.includes("センター") ||
    s.includes("中寄せ")
  ) {
    p.align = "center";
  }

  // ---- ４隅ショートカット ----
  if (s.includes("左上")) {
    p.align = "left";
    p.vAlign = "top";
  }
  if (s.includes("右上")) {
    p.align = "right";
    p.vAlign = "top";
  }
  if (s.includes("左下")) {
    p.align = "left";
    p.vAlign = "bottom";
    p.bottomMargin = 80;
  }
  if (s.includes("右下")) {
    p.align = "right";
    p.vAlign = "bottom";
    p.bottomMargin = 80;
  }

  // ---- 微調整（少し右 / 少し上 など）----
  if (s.includes("少し右") || s.includes("ちょい右") || s.includes("やや右")) {
    p.offsetX = (p.offsetX ?? 0) + 80;
  }
  if (s.includes("少し左") || s.includes("ちょい左") || s.includes("やや左")) {
    p.offsetX = (p.offsetX ?? 0) - 80;
  }
  if (s.includes("少し上") || s.includes("ちょい上") || s.includes("やや上")) {
    p.offsetY = (p.offsetY ?? 0) - 60;
  }
  if (s.includes("少し下") || s.includes("ちょい下") || s.includes("やや下")) {
    p.offsetY = (p.offsetY ?? 0) + 60;
  }

  // ---- 矢印による移動指定（→ ← ↑ ↓）----
  // ここは「増分」として扱う（あとでベース offset に加算）
  if (s.includes("→") || s.includes("➡") || s.includes("➜") || s.includes("右矢印")) {
    p.offsetX = (p.offsetX ?? 0) + 80;
  }
  if (s.includes("←") || s.includes("⬅") || s.includes("左矢印")) {
    p.offsetX = (p.offsetX ?? 0) - 80;
  }
  if (s.includes("↑") || s.includes("⬆") || s.includes("上矢印")) {
    p.offsetY = (p.offsetY ?? 0) - 60;
  }
  if (s.includes("↓") || s.includes("⬇") || s.includes("下矢印")) {
    p.offsetY = (p.offsetY ?? 0) + 60;
  }

  // ---- フォント ----
  if (s.includes("メイリオ")) p.font = "Meiryo";
  if (s.includes("游ゴシック") || s.includes("游ｺﾞｼｯｸ")) p.font = "Yu Gothic";
  if (s.includes("ゴシック")) p.font = "Yu Gothic";
  if (s.includes("明朝")) p.font = "Yu Mincho";
  if (s.includes("手書き") || s.includes("手書き風")) {
    // 実際には別の手書きフォントに差し替えてOK
    p.font = "Comic Sans MS";
  }

  // ---- 色 ----
  if (s.includes("白文字") || s.includes("白")) p.color = "#ffffff";
  if (s.includes("黒文字") || s.includes("黒")) p.color = "#000000";
  if (s.includes("赤文字") || s.includes("赤")) p.color = "red";
  if (s.includes("青文字") || s.includes("青")) p.color = "blue";
  if (s.includes("黄色") || s.includes("黄")) p.color = "yellow";

  return p;
}

/* ------------------------------------------------------------------ */

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

  // ★ 画像生成ツール（新しく描く用）
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
          text: { type: "string" }, // ※ 今は無視して「元絵だけ」生成。文字入れは add_text_to_existing_image を使う想定。
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
        "Instead, call add_text_to_existing_image with the previous image URL.",
      name: "create_img",
    },
  });

  // ★ 既存画像に文字だけ足すツール（Vision を使わないシンプル版）
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
            description: "Japanese text to overlay on the image.",
          },
          styleHint: {
            type: "string",
            description:
              "Natural language hint for font size, color, position such as '大きめの白文字で、下部中央に', '少し上に', '➡ で少し右へ', etc.",
          },
          font: {
            type: "string",
            description:
              "Font family name if explicitly requested (e.g., 'Meiryo').",
          },
          color: {
            type: "string",
            description: "Text color (e.g., 'white', '#ffffff').",
          },
          size: {
            type: "string",
            description:
              "Rough size hint like 'small', 'medium', 'large'. You can infer from the user's request.",
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
        "Use this tool when the user wants to add or adjust text on an EXISTING image, for example 'この絵に 2026 謹賀新年 と入れて' or 'もう少し下に', 'そこから➡で右に'. " +
        "This version does NOT use Azure Vision; it simply overlays text on top of the existing image using the /api/gen-image route.",
      name: "add_text_to_existing_image",
    },
  });

  return { status: "OK", response: defaultExtensions };
};

// ---------------- 画像生成（NEW image 用） ----------------
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
  const size = (args?.size || "1024x1024").trim();

  console.log("createImage called with prompt:", prompt);
  console.log("createImage (initial) will NOT add text overlay in this version.");
  console.log(
    "🧩 reasoning_effort in request:",
    modeOpts?.reasoning_effort || "none"
  );

  if (!prompt) return "No prompt provided";
  if (prompt.length >= 4000)
    return "Prompt is too long, it must be less than 4000 characters";

  const endpointRaw = process.env.AZURE_OPENAI_ENDPOINT || "";
  const endpoint = endpointRaw.replace(/\/+$/, "");
  const apiKey = process.env.AZURE_OPENAI_API_KEY || "";
  const deployment = process.env.AZURE_OPENAI_IMAGE_DEPLOYMENT || "";
  const apiVersion =
    process.env.AZURE_OPENAI_API_VERSION || "2025-04-01-preview";

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

  const imageGenUrl = `${endpoint}/openai/deployments/${encodeURIComponent(
    deployment
  )}/images/generations?api-version=${encodeURIComponent(apiVersion)}`;

  let json: any;
  try {
    const res = await fetch(imageGenUrl, {
      method: "POST",
      headers: { "api-key": apiKey, "Content-Type": "application/json" },
      body: JSON.stringify({
        // ★ ここではテキスト無しのクリーンな画像だけを生成
        prompt,
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
    let baseImageUrl: string;

    if (b64) {
      const imageName = `${uniqueId()}.png`;
      const buffer = Buffer.from(b64, "base64");

      // 通常のランダム名で保存（ユーザーに見せる用）
      await UploadImageToStore(chatThread.id, imageName, buffer);

      // ★ スレッド共通の「元絵」として __base__.png にも同じ内容を保存
      //   - 同じスレッドで create_img を再実行したら、ここで __base__.png を上書き
      await UploadImageToStore(chatThread.id, "__base__.png", buffer);

      // 表示用 URL は従来どおりランダム名を使う
      baseImageUrl = buildExternalImageUrl(chatThread.id, imageName);
    } else {
      // まれに URL で返るケース（今はほぼ無いはずだが念のため）
      baseImageUrl = urlDirect!;
    }

    // ★ このシンプル版では、ここで文字入れは行わない（元絵だけ返す）
    return {
      revised_prompt: prompt,
      url: baseImageUrl,
    };
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
  // LLM から渡された URL はログ用に保持（実際のベースには使わない）
  const explicitUrl = (args?.imageUrl || "").trim();
  const text = (args?.text || "").trim();
  const styleHint = (args?.styleHint || "").trim();

  // ★ ここが今回の本丸：
  //   - ベース画像は常に「threadId/__base__.png」を使う
  //   - explicitUrl（直前に表示していた画像）は参照しない
  const baseImageUrl = buildExternalImageUrl(chatThread.id, "__base__.png");

  console.log("🖋 add_text_to_existing_image (simple) called:", {
    passedImageUrl: explicitUrl,
    usedBaseImageUrl: baseImageUrl,
    text,
    styleHint,
    offsetX: args?.offsetX,
    offsetY: args?.offsetY,
  });

  if (!text) {
    return {
      error: "text is required for add_text_to_existing_image.",
    };
  }

  // ★ styleHint + userMessage からスタイルを推定
  const hintSource = styleHint || userMessage || "";
  const parsed = parseStyleHint(hintSource);

  // 明示指定があればそれを優先しつつ、パース結果をデフォルトとして利用
  const align: "left" | "center" | "right" =
    (parsed.align as any) ?? "center";
  const vAlign: "top" | "middle" | "bottom" =
    (parsed.vAlign as any) ?? "bottom";
  const size: "small" | "medium" | "large" | "xlarge" =
    (args.size as any) ?? parsed.size ?? "large";
  const color = args.color ?? parsed.color ?? "white";
  const font = args.font ?? parsed.font;

  // ★ 累積移動：args の offset をベースに、styleHint 由来の増分を足す
  const baseOffsetX =
    typeof args.offsetX === "number" ? args.offsetX : 0;
  const baseOffsetY =
    typeof args.offsetY === "number" ? args.offsetY : 0;

  const offsetX = baseOffsetX + (parsed.offsetX ?? 0);
  const offsetY = baseOffsetY + (parsed.offsetY ?? 0);

  const bottomMargin = parsed.bottomMargin; // route.ts 側で undefined ならデフォルト 80 が効く

  const baseUrl =
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME
      ? `https://${process.env.WEBSITE_HOSTNAME}`
      : "http://localhost:3000");

  const genImageBase = baseUrl.replace(/\/+$/, "");
  console.log("[gen-image] base URL for overlay:", genImageBase);
  console.log("[gen-image] resolved style params:", {
    align,
    vAlign,
    size,
    color,
    font,
    offsetX,
    offsetY,
    bottomMargin,
  });

  try {
    const resp = await fetch(`${genImageBase}/api/gen-image`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        imageUrl: baseImageUrl, // ← ★ 毎回 __base__.png を元絵として使う
        text,
        align,
        vAlign,
        size, // small/medium/large/xlarge を route.ts 側で fontSize にマップ
        color,
        font,
        offsetX,
        offsetY,
        bottomMargin,
        autoDetectPlacard: false, // ★ ここで完全に OFF にする
      }),
    });

    if (!resp.ok) {
      const t = await resp.text().catch(() => "");
      console.error(
        "🔴 /api/gen-image failed in edit:",
        resp.status,
        t
      );
      return {
        error: `Text overlay failed: HTTP ${resp.status}`,
      };
    }

    const result = await resp.json();
    const generatedPath = result?.imageUrl as string | undefined;

    if (!generatedPath) {
      console.error("🔴 gen-image edit returned no imageUrl");
      return { error: "gen-image edit returned no imageUrl" };
    }

    // /generated/xxx.png を Azure Storage の images コンテナに保存し直す
    const fs = require("fs");
    const path = require("path");
    const finalImageName = `${uniqueId()}.png`;
    const finalImagePath = path.join(
      process.cwd(),
      "public",
      generatedPath.startsWith("/")
        ? generatedPath.slice(1)
        : generatedPath
    );
    const finalImageBuffer = fs.readFileSync(finalImagePath);

    await UploadImageToStore(
      chatThread.id,
      finalImageName,
      finalImageBuffer
    );

    const finalImageUrl = buildExternalImageUrl(
      chatThread.id,
      finalImageName
    );

    return {
      revised_prompt: text,
      url: finalImageUrl,
    };
  } catch (err) {
    console.error("🔴 error in executeAddTextToExistingImage (simple):", err);
    return {
      error: "There was an error adding text to the existing image: " + err,
    };
  }
}
