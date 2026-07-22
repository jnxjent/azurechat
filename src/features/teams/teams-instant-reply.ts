const GREETING_REPLIES = new Map<string, string>([
  ["こんにちは", "こんにちは！ご用件をどうぞ。"],
  ["こんにちわ", "こんにちは！ご用件をどうぞ。"],
  ["こんばんは", "こんばんは！ご用件をどうぞ。"],
  ["こんばんわ", "こんばんは！ご用件をどうぞ。"],
  ["おはよう", "おはようございます！ご用件をどうぞ。"],
  ["おはようございます", "おはようございます！ご用件をどうぞ。"],
  ["hello", "Hello! How can I help?"],
  ["hi", "Hello! How can I help?"],
  ["hey", "Hello! How can I help?"],
]);

/**
 * Returns a deterministic response only for a greeting-only Teams message.
 * Substantive messages such as "こんにちは、規程を探して" must continue to
 * Azure AI Search and Azure OpenAI.
 */
export function resolveTeamsInstantReply(message: string): string | null {
  const normalized = message
    .normalize("NFKC")
    .trim()
    .toLowerCase()
    .replace(/[\s!！?？。、,.…]+$/g, "")
    .trim();

  return GREETING_REPLIES.get(normalized) ?? null;
}
