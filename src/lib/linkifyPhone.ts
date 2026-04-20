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

/**
 * 全角→半角の最低限変換
 */
function toHalfWidth(input: string): string {
  return input.replace(/[０-９＋－（）　]/g, (char) => ZENKAKU_MAP[char] ?? char);
}

/**
 * tel: リンク用に電話番号を正規化
 * - 全角数字を半角へ
 * - 空白、ハイフン、括弧を除去
 */
export function normalizePhoneForTel(phone: string): string {
  return toHalfWidth(phone).replace(/[\s\-()]/g, "");
}

/**
 * 日本の一般的な電話番号をざっくり検出
 * 例:
 * - 03-1234-5678
 * - 090-1234-5678
 * - +81-90-1234-5678
 * - ０５３-１２３-４５６７
 */
const PHONE_REGEX =
  /(?:\+81[-\s]?\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}|0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}|[０-９＋－（）\s]{10,20})/g;

/**
 * 電話番号として妥当かを軽く判定
 * - 数字数が少なすぎるものを除外
 */
function looksLikePhone(raw: string): boolean {
  const normalized = normalizePhoneForTel(raw);
  const digitsOnly = normalized.replace(/[^\d+]/g, "");
  const digitCount = digitsOnly.replace(/[^\d]/g, "").length;
  return digitCount >= 10 && digitCount <= 15;
}

/**
 * プレーンテキストを「通常文字列」と「電話番号」に分解
 */
export function splitTextWithPhones(text: string): TextPart[] {
  if (!text) {
    return [{ type: "text", value: "" }];
  }

  const parts: TextPart[] = [];
  let lastIndex = 0;

  const matches = Array.from(text.matchAll(PHONE_REGEX));  // ← 修正

  for (const match of matches) {
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