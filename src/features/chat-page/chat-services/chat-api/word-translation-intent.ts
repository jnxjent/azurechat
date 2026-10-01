export const WORD_TRANSLATION_LANGUAGE_PATTERNS = [
  ["en", /英語|英訳|English/i],
  ["pt", /ポルトガル語|Portuguese/i],
  ["vi", /ベトナム語|Vietnamese/i],
  ["id", /インドネシア語|Indonesian/i],
  ["zh-CN", /中国語|簡体字|Chinese/i],
  ["ko", /韓国語|ハングル|Korean/i],
  ["es", /スペイン語|Spanish/i],
  ["fil", /タガログ語|フィリピノ語?|Tagalog|Filipino/i],
] as const;

export type WordTranslationLanguage =
  (typeof WORD_TRANSLATION_LANGUAGE_PATTERNS)[number][0];

export function detectWholeWordTranslationLanguage(
  text: string,
  hasWordSource: boolean
): WordTranslationLanguage | null {
  if (!hasWordSource) return null;
  const hasTranslationIntent =
    /翻訳|英訳|訳して?|(?:英語|ポルトガル語|ベトナム語|インドネシア語|中国語|簡体字|韓国語|スペイン語|タガログ語|フィリピノ語?)訳/i.test(
      text
    );
  if (!hasTranslationIntent) return null;
  const targetLanguage =
    WORD_TRANSLATION_LANGUAGE_PATTERNS.find(([, pattern]) =>
      pattern.test(text)
    )?.[0] ?? "en";
  const referencesWholeWord =
    /添付|Word|ワード|\.docx|全文|全体|全部|丸ごと/i.test(text);
  const requestsPartialTranslation =
    /(?:第?\d+|一|二|三)(?:段落|章|節)|一部|特定(?:の)?箇所|選択(?:した)?(?:箇所|範囲)|タイトル|見出し/i.test(
      text
    );
  const explicitlyWhole = /全文|全体|全部|丸ごと/i.test(text);
  return referencesWholeWord &&
    (!requestsPartialTranslation || explicitlyWhole)
    ? targetLanguage
    : null;
}
