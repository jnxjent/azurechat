// src/app/(authenticated)/api/chat/route.ts
import { ChatAPIEntry } from "@/features/chat-page/chat-services/chat-api/chat-api";
import { UserPrompt } from "@/features/chat-page/chat-services/models";

export const runtime = "nodejs";

// 🚀 キャッシュ禁止（既存）
export const dynamic = "force-dynamic";
export const revalidate = 0;

/** UIの3値 */
type ThinkingModeUI = "standard" | "thinking" | "fast";
/** APIで使う3値（standardはnormalへ） */
type ThinkingModeAPI = "normal" | "thinking" | "fast";

function uiToApi(m?: ThinkingModeUI | null): ThinkingModeAPI {
  if (!m) return "normal";
  if (m === "thinking") return "thinking";
  if (m === "fast") return "fast";
  return "normal"; // standard → normal
}

type UserPromptWithMode = UserPrompt & {
  thinkingMode?: ThinkingModeUI;
  apiThinkingMode?: ThinkingModeAPI;
};

export async function POST(req: Request) {
  const formData = await req.formData();

  // 本文（JSON文字列）
  const content = formData.get("content");
  const multimodalImage = formData.get("image-base64");
  const uiThinkingMode = formData.get("thinkingMode") as ThinkingModeUI | null;

  if (typeof content !== "string") {
    return new Response(
      JSON.stringify({
        error: "missing_content",
        message: "`content` must be a JSON string.",
      }),
      { status: 400, headers: { "content-type": "application/json" } }
    );
  }

  let parsed: any;
  try {
    parsed = JSON.parse(content);
  } catch {
    return new Response(
      JSON.stringify({
        error: "invalid_json",
        message: "`content` is not valid JSON.",
      }),
      { status: 400, headers: { "content-type": "application/json" } }
    );
  }

  // UI→API 正規化
  const apiThinkingMode = uiToApi(uiThinkingMode);

  if (process.env.NODE_ENV !== "production") {
    console.log("🚦 route.ts resolved (Body only):", {
      resolvedUI: uiThinkingMode ?? "standard",
      resolvedAPI: apiThinkingMode,
    });
  }

  // 🧠 API に渡すプロンプト生成
  const userPrompt: UserPromptWithMode = {
    ...(parsed as UserPromptWithMode),
    thinkingMode: uiThinkingMode ?? "standard",
    apiThinkingMode,
    multimodalImage:
      typeof multimodalImage === "string" && multimodalImage.length > 0
        ? multimodalImage
        : "",
  };

  // 🆕 GPTに渡す履歴を最新30件に制限（速度・Token・応答安定性◎）
  // UserPromptWithMode の型定義上は `messages` が無い（`message` はある）ため、
  // unknown 経由で安全に存在確認してから slice する。
  const upAny = userPrompt as unknown as {
    messages?: unknown;
    history?: unknown;
  };

  if (Array.isArray(upAny.messages)) {
    upAny.messages = upAny.messages.slice(-30);
  }
  if (Array.isArray(upAny.history)) {
    upAny.history = upAny.history.slice(-30);
  }

  // LLMへ送信
  return await ChatAPIEntry(userPrompt, req.signal);
}
