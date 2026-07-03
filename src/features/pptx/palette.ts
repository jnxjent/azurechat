export type PptxPalette = {
  canvas: string; surface: string; titleBg: string; headerBg: string;
  accentA: string; accentB: string; headerText: string; bodyText: string;
  mutedText: string; sectionBg: string; tableHeaderBg: string;
  tableHeaderText: string; tableAltBg: string; border: string;
};

type NamedPaletteBase = {
  main: string; accent: string; accent_light: string; main_light: string; text_muted: string;
  labelJa: string; shortDesc: string; mood: string; recommendedFor: string;
};

export const PPTX_NAMED_PALETTES: Record<string, NamedPaletteBase> = {
  navy_orange:    { main:"13294B", accent:"F5821F", accent_light:"EEF2F9", main_light:"E4E8F0", text_muted:"6B7488",
    labelJa:"ネイビー×オレンジ", shortDesc:"濃紺×オレンジ", mood:"知的・信頼・力強い", recommendedFor:"IT・DX・経営・提案書" },
  forest_amber:   { main:"1B4D3E", accent:"F4A300", accent_light:"FBEFD5", main_light:"E3EDE8", text_muted:"5E6E66",
    labelJa:"深緑×アンバー",       shortDesc:"深緑×琥珀色", mood:"誠実・活力・自然",   recommendedFor:"採用・農業・エコ" },
  burgundy_gold:  { main:"8C1D18", accent:"E0A33B", accent_light:"F7ECD6", main_light:"F3E5E4", text_muted:"6E5A58",
    labelJa:"バーガンディ×ゴールド", shortDesc:"深赤×金",   mood:"格式・老舗・高級感", recommendedFor:"製造・高級品・老舗企業" },
  teal_coral:     { main:"0E4D5C", accent:"EE6C4D", accent_light:"FBE6DE", main_light:"DCE9EC", text_muted:"5A6B70",
    labelJa:"ティール×コーラル",    shortDesc:"青緑×珊瑚色", mood:"清潔・温かい・誠実", recommendedFor:"医療・環境・産廃" },
  charcoal_terra: { main:"333333", accent:"C15F3C", accent_light:"F3E3DA", main_light:"ECECEA", text_muted:"6E6E6E",
    labelJa:"チャコール×テラコッタ", shortDesc:"炭色×煉瓦色", mood:"重厚・実直・安定", recommendedFor:"建設・土木・インフラ" },
};

// 色パレット一覧テキスト（ユーザー向けメッセージ用）
export function pptxPaletteListText(): string {
  return PPTX_PALETTE_KEYS.map((k, i) => {
    const m = PPTX_NAMED_PALETTES[k];
    return `${i + 1}. **${m.labelJa}**（${m.mood}）：${m.recommendedFor}`;
  }).join("\n");
}

export const PPTX_PALETTE_KEYS = Object.keys(PPTX_NAMED_PALETTES);

export function buildPaletteFromKey(key: string): PptxPalette | undefined {
  const s = PPTX_NAMED_PALETTES[key];
  if (!s) return undefined;
  return {
    canvas:"FFFFFF", surface:"FFFFFF", titleBg:s.main, headerBg:s.main, accentA:s.main,
    accentB:s.accent, headerText:"FFFFFF", bodyText:s.main, mutedText:s.text_muted,
    sectionBg:s.accent_light, tableHeaderBg:s.main, tableHeaderText:"FFFFFF",
    tableAltBg:s.main_light, border:s.main_light,
  };
}

export type PaletteResolution = {
  accentColor: string;
  paletteKey?: string;
  palette?: PptxPalette;
};

// 判定順: 日本語複合名 → 英語ID → 個別日本語KW → 数字参照(1〜5) → 基本色
export function resolvePptxPaletteInstruction(s: string): PaletteResolution | null {
  const t = s.toLowerCase();
  let key: string | null = null;

  // 日本語複合名（最優先: navy_orange が navy 単体に誤マッチしないよう）
  if      (/(ネイビー.{0,2}オレンジ|ネイビーオレンジ)/.test(s))           key = "navy_orange";
  else if (/(深緑.{0,2}アンバー|深緑アンバー|フォレスト.{0,2}アンバー)/.test(s)) key = "forest_amber";
  else if (/(バーガンディ.{0,2}ゴールド|バーガンディゴールド|ワインレッド.{0,2}ゴールド)/.test(s)) key = "burgundy_gold";
  else if (/(ティール.{0,2}コーラル|ティールコーラル|青緑.{0,2}コーラル)/.test(s)) key = "teal_coral";
  else if (/(チャコール.{0,2}テラコッタ|チャコールテラコッタ)/.test(s))   key = "charcoal_terra";
  // 英語ID
  else if (/navy.orange/.test(t))    key = "navy_orange";
  else if (/forest.amber/.test(t))   key = "forest_amber";
  else if (/burgundy.gold/.test(t))  key = "burgundy_gold";
  else if (/teal.coral/.test(t))     key = "teal_coral";
  else if (/charcoal.terra/.test(t)) key = "charcoal_terra";
  // 個別日本語KW (amber=forest_amber, gold=burgundy_gold に注意)
  else if (/(バーガンディ|ワインレッド|深赤|burgundy|maroon|crimson)/.test(s)) key = "burgundy_gold";
  else if (/(ゴールド|金色|gold)/.test(t))                                key = "burgundy_gold";
  else if (/(ティール|コーラル|サンゴ色?|青緑|teal|coral|cyan|turquoise)/.test(s)) key = "teal_coral";
  else if (/(チャコール|テラコッタ|煉瓦色?|炭色?|charcoal|terra)/.test(s)) key = "charcoal_terra";
  else if (/(深緑|フォレスト|アンバー|琥珀色?|forest|amber)/.test(s))     key = "forest_amber";
  else if (/(ネイビー|navy|紺)/.test(s))                                  key = "navy_orange";

  if (key) {
    const palette = buildPaletteFromKey(key)!;
    return { paletteKey: key, palette, accentColor: PPTX_NAMED_PALETTES[key].main };
  }

  // 数字参照: "1で"/"2で"... → palette by index (色名未検出時のみ)
  // 候補を個別評価し、ページ番号文脈（P3で/スライド3で/ページ3で等）のみ除外
  let paletteNumMatch: RegExpMatchArray | null = null;
  const paletteNumberRe = /([1-5１-５])[でに]/g;
  let m: RegExpExecArray | null;
  while ((m = paletteNumberRe.exec(s)) !== null) {
    const before = s.slice(Math.max(0, m.index - 10), m.index);
    if (/[Pp](?:age)?\s*$/.test(before) || /スライド\s*$/.test(before) || /ページ\s*$/.test(before)) continue;
    paletteNumMatch = m;
    break;
  }
  if (paletteNumMatch) {
    const d = paletteNumMatch[1];
    const n = d >= "1" && d <= "5" ? parseInt(d) : "１２３４５".indexOf(d) + 1;
    const paletteKey = PPTX_PALETTE_KEYS[n - 1];
    if (paletteKey) {
      const palette = buildPaletteFromKey(paletteKey)!;
      return { paletteKey, palette, accentColor: PPTX_NAMED_PALETTES[paletteKey].main };
    }
  }

  // 基本色（パレット非該当）
  if (/(赤|red)/.test(t))              return { accentColor: "C00000" };
  if (/(青|blue)/.test(t))             return { accentColor: "2F5597" };
  if (/(緑|green)/.test(t))            return { accentColor: "548235" };
  if (/(紫|purple)/.test(t))           return { accentColor: "7030A0" };
  if (/(オレンジ|orange|橙)/.test(t)) return { accentColor: "C55A11" };
  if (/(黄|yellow)/.test(t))           return { accentColor: "BF9000" };
  if (/(ピンク|pink)/.test(t))         return { accentColor: "C0508A" };
  if (/(グレー|gray|grey)/.test(t))    return { accentColor: "666666" };

  return null;
}
