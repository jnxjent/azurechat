// src/app/(authenticated)/api/chat/route.ts
import { ChatAPIEntry } from "@/features/chat-page/chat-services/chat-api/chat-api";
import { UserPrompt } from "@/features/chat-page/chat-services/models";

export const runtime = "nodejs";

/** UI の 3値 */
type ThinkingModeUI = "standard" | "thinking" | "fast";
/** API の 3値（standard → normal に正規化） */
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

  // 本文（JSON 文字列）
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
