const EXPLICIT_WORD_REPLACEMENT_PATTERNS = [
  /[「『"']([^」』"']+)[」』"']\s*(?:→|->|=>)\s*[「『"']([^」』"']*)[」』"']/,
  /(?:【誤】|[（(]誤[）)])\s*[^【】\r\n]+?\s*(?:\r?\n\s*)?(?:【正】|[（(]正[）)])\s*[^【】\r\n]*/,
  /[「『"']([^」』"']+)[」』"']\s*を\s*[「『"']([^」』"']*)[」』"']\s*に/,
];

export function hasExplicitWordReplacementInstruction(value: string): boolean {
  const instruction = String(value ?? "").trim();
  return (
    instruction.length > 0 &&
    EXPLICIT_WORD_REPLACEMENT_PATTERNS.some((pattern) => pattern.test(instruction))
  );
}

export type WordEditInstructionSource =
  | "tool"
  | "tool-explicit"
  | "user-explicit"
  | "user-original";

/**
 * Keep the user's literal A→B direction authoritative when it is explicit.
 * For a broad proofreading request, preserve the concrete closed replacement
 * list produced after document search instead of replacing it with the broad
 * original request and losing every applicable edit.
 */
export function resolveWordEditInstruction(
  toolInstruction: string,
  originalUserInstruction?: string
): { instruction: string; source: WordEditInstructionSource } {
  const tool = String(toolInstruction ?? "").trim();
  const original = String(originalUserInstruction ?? "").trim();

  if (!original) return { instruction: tool, source: "tool" };
  if (hasExplicitWordReplacementInstruction(original)) {
    return { instruction: original, source: "user-explicit" };
  }
  if (hasExplicitWordReplacementInstruction(tool)) {
    return { instruction: tool, source: "tool-explicit" };
  }
  return { instruction: original, source: "user-original" };
}
