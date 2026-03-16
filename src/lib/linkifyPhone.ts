// src/lib/linkifyPhone.ts

export type TextPart =
  | { type: "text"; value: string }
  | { type: "phone"; value: string; telHref: string };

const ZENKAKU_MAP: Record<string, string> = {
  "０": "0",
  "１": "1",
  "２": "2",
  "３": "3",
  "４": "4",
  "５": "5",
  "６": "6",
  "７": "7",
  "８": "8",
  "９": "9",
  "＋": "+",
  "－": "-",
  "（": "(",
  "）": ")",
  "　": " ",
};

function toHalfWidth(input: string): string {
  return input.replace(/[０-９＋－（）　]/g, (char) => ZENKAKU_MAP[char] ?? char);
}

export function normalizePhoneForTel(phone: string): string {
  return toHalfWidth(phone).replace(/[\s\-()]/g, "");
}

const PHONE_REGEX =
  /(?:\+81[-\s]?\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}|0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}|[０-９＋－（）\s]{10,20})/g;

function looksLikePhone(raw: string): boolean {
  const normalized = normalizePhoneForTel(raw);
  const digitsOnly = normalized.replace(/[^\d+]/g, "");
  const digitCount = digitsOnly.replace(/[^\d]/g, "").length;
  return digitCount >= 10 && digitCount <= 15;
}

export function splitTextWithPhones(text: string): TextPart[] {
  if (!text) {
    return [{ type: "text", value: "" }];
  }

  const parts: TextPart[] = [];
  let lastIndex = 0;

  // ★ matchAll → exec ループに変更（downlevelIteration不要）
  let match: RegExpExecArray | null;
  PHONE_REGEX.lastIndex = 0;
  while ((match = PHONE_REGEX.exec(text)) !== null) {
    const raw = match[0];
    const index = match.index ?? 0;

    if (!looksLikePhone(raw)) {
      continue;
    }

    if (index > lastIndex) {
      parts.push({
        type: "text",
        value: text.slice(lastIndex, index),
      });
    }

    parts.push({
      type: "phone",
      value: raw,
      telHref: `tel:${normalizePhoneForTel(raw)}`,
    });

    lastIndex = index + raw.length;
  }

  if (lastIndex < text.length) {
    parts.push({
      type: "text",
      value: text.slice(lastIndex),
    });
  }

  return parts;
}