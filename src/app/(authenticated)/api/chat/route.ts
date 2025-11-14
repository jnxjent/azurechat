// src/app/(authenticated)/api/chat/route.ts
import { ChatAPIEntry } from "@/features/chat-page/chat-services/chat-api/chat-api";
import { UserPrompt } from "@/features/chat-page/chat-services/models";

export const runtime = "nodejs";

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

/** 最小ガード：直前 assistant の tool_calls に紐付かない tool を除外 */
function fixOrphanTools(messages: any[]) {
  if (!Array.isArray(messages)) return messages;
  const out: any[] = [];
  let lastAssistantToolIds: Set<string> | null = null;

  for (const m of messages) {
    if (m?.role === "assistant") {
      lastAssistantToolIds = null;
      if (Array.isArray(m.tool_calls) && m.tool_calls.length > 0) {
        lastAssistantToolIds = new Set(
          m.tool_calls.map((tc: any) => tc?.id).filter(Boolean)
        );
      } else if (Array.isArray(m.tool_calls)) {
        delete (m as any).tool_calls; // 空配列は削除
      }
      out.push(m);
      continue;
    }

    if (m?.role === "tool") {
      if (lastAssistantToolIds && m.tool_call_id && lastAssistantToolIds.has(m.tool_call_id)) {
        out.push(m);
      }
      continue; // 孤立toolは落とす
    }

    // user / system
    lastAssistantToolIds = null;
    out.push(m);
  }
  return out;
}

/** リカバリ用：tool を全除去＆assistant.tool_calls も除去（最後の手段） */
function hardSanitize(messages: any[]) {
  if (!Array.isArray(messages)) return messages;
  const out: any[] = [];
  for (const m of messages) {
    if (m?.role === "tool") continue; // すべて落とす
    if (m?.role === "assistant") {
      if (m.tool_calls) delete (m as any).tool_calls;
    }
    out.push(m);
  }
  return out.length ? out : [{ role: "user", content: "（ツール出力を無視して続行）" }];
}

export async function POST(req: Request) {
  const formData = await req.formData();

  // 本文（JSON文字列）
  const content = formData.get("content");
  const multimodalImage = formData.get("image-base64");
  const uiThinkingMode = formData.get("thinkingMode") as ThinkingModeUI | null;

  if (typeof content !== "string") {
    return new Response(
      JSON.stringify({ error: "missing_content", message: "`content` must be a JSON string." }),
      { status: 400, headers: { "content-type": "application/json" } }
    );
  }

  let parsed: any;
  try {
    parsed = JSON.parse(content);
  } catch {
    return new Response(
      JSON.stringify({ error: "invalid_json", message: "`content` is not valid JSON." }),
      { status: 400, headers: { "content-type": "application/json" } }
    );
  }

  // 送信元が messages を持っている場合：まずは通常サニタイズ
  if (Array.isArray(parsed?.messages)) {
    parsed.messages = fixOrphanTools(parsed.messages);
    if (parsed.messages.length === 0) {
      parsed.messages = [{ role: "user", content: "（ツール出力を無視して続行）" }];
    }
  }

  // UI→API 正規化
  const apiThinkingMode = uiToApi(uiThinkingMode);

  if (process.env.NODE_ENV !== "production") {
    console.log("🚦 route.ts resolved (Body only):", {
      resolvedUI: uiThinkingMode ?? "standard",
      resolvedAPI: apiThinkingMode,
    });
  }

  // API に渡すプロンプト
  const userPrompt: UserPromptWithMode = {
    ...(parsed as UserPromptWithMode),
    thinkingMode: uiThinkingMode ?? "standard",
    apiThinkingMode,
    multimodalImage:
      typeof multimodalImage === "string" && multimodalImage.length > 0 ? multimodalImage : undefined,
  };

  // 念のためこちら側も整合
  if (Array.isArray((userPrompt as any).messages)) {
    (userPrompt as any).messages = fixOrphanTools((userPrompt as any).messages);
    if ((userPrompt as any).messages.length === 0) {
      (userPrompt as any).messages = [{ role: "user", content: "（ツール出力を無視して続行）" }];
    }
  }

  // 実行＆400特定メッセージ時は自動リトライ（ハードサニタイズ）
  try {
    return await ChatAPIEntry(userPrompt, req.signal);
  } catch (e: any) {
    const msg: string = e?.error?.message || e?.message || "";
    if (/role 'tool' must be a response to a preceeding message with 'tool_calls'/.test(msg)) {
      // リカバリ：tool をすべて落として再送
      if (Array.isArray((userPrompt as any).messages)) {
        (userPrompt as any).messages = hardSanitize((userPrompt as any).messages);
      }
      // parsed 側も同期（ChatAPIEntry がこちらを参照する場合に備える）
      if (Array.isArray(parsed?.messages)) {
        parsed.messages = hardSanitize(parsed.messages);
      }
      return await ChatAPIEntry(userPrompt, req.signal);
    }
    throw e;
  }
}
