// src/features/chat-page/chat-services/chat-api/reasoning-utils.ts

/** UI/互換で飛んでくる可能性のある表記 */
export type ThinkingModeInput = "standard" | "thinking" | "fast" | "normal";

/** 内部で使う正規化済み3値 */
export type ThinkingModeCanonical = "normal" | "thinking" | "fast";

export type ReasoningOptions = {
  reasoning_effort: "low" | "medium" | "high";
  temperature: number;
};

/** 表記ゆれ → 正規化（standard→normal） */
export function canonicalizeMode(m?: ThinkingModeInput | null): ThinkingModeCanonical {
  const v = (m ?? "normal").toLowerCase();
  if (v === "thinking" || v === "deep") return "thinking";
  if (v === "fast" || v === "quick") return "fast";
  if (v === "standard" || v === "normal") return "normal";
  return "normal";
}

/** プリセット：fast=low / normal=medium / thinking=high */
const PRESETS: Record<ThinkingModeCanonical, ReasoningOptions> = {
  fast:     { reasoning_effort: "low",    temperature: 0.10 },
  normal:   { reasoning_effort: "medium", temperature: 0.20 },
  thinking: { reasoning_effort: "high",   temperature: 0.40 },
};

/** モードに応じて送信オプションを返す（auto は存在しない想定） */
export function buildSendOptionsFromMode(
  mode?: ThinkingModeInput | ThinkingModeCanonical | null
): ReasoningOptions {
  const canon = canonicalizeMode(mode as any);
  return PRESETS[canon] ?? PRESETS.normal;
}

/** ログ向け（確認用） */
export function buildSendOptionsWithMeta(
  mode?: ThinkingModeInput | ThinkingModeCanonical | null
): { displayMode: string; canonicalMode: ThinkingModeCanonical; options: ReasoningOptions } {
  const displayMode = (mode ?? "normal") as string;
  const canonicalMode = canonicalizeMode(mode as any);
  const options = buildSendOptionsFromMode(mode);
  return { displayMode, canonicalMode, options };
}
