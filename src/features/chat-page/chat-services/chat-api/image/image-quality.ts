export const GPT_IMAGE_QUALITIES = [
  "low",
  "medium",
  "high",
  "auto",
] as const;

export type GptImageQuality = (typeof GPT_IMAGE_QUALITIES)[number];

/** Preserve the API's existing auto behavior for omitted or invalid values. */
export function normalizeGptImageQuality(value: unknown): GptImageQuality {
  const normalized = String(value ?? "").trim().toLowerCase();
  return GPT_IMAGE_QUALITIES.includes(normalized as GptImageQuality)
    ? (normalized as GptImageQuality)
    : "auto";
}
