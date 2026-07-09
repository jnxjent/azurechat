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
  coral_orange:   { main:"1B5E3B", accent:"E07038", accent_light:"FCEADE", main_light:"E1EDE6", text_muted:"5B6E62",
    labelJa:"深緑×コーラルオレンジ", shortDesc:"深緑×珊瑚橙", mood:"活力・誠実・フレッシュ", recommendedFor:"産廃・環境サービス・営業資料（暖色系）" },
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

// 判定順: ①複合パレット名明示 → ②基本色明示 → ③日本語複合名 → ④英語ID → ⑤個別KW → ⑥数字参照(1〜6) → ⑦基本色
export function resolvePptxPaletteInstruction(s: string): PaletteResolution | null {
  const t = s.toLowerCase();

  // 変更先色の明示判定ヘルパー（変更動詞との近接 ≤12文字で判定）
  // LLMが「既存配色（ティール×コーラル）の雰囲気を維持しつつ赤系に」のような instruction を渡した際、
  // 既存配色名が先にマッチして誤ったパレットが返される問題を防ぐ。
  const isExplicitTarget = (colorPat: string): boolean => {
    const fwd = new RegExp(`(?:${colorPat}).{0,12}(?:に|へ|で|基調|トーン|統一|変更|変えて|寄せ)`, "i");
    const bwd = new RegExp(`(?:に|へ|で|基調|トーン|統一|変更|変えて|寄せ).{0,12}(?:${colorPat})`, "i");
    return fwd.test(s) || bwd.test(s);
  };
  const paletteOf = (key: string): PaletteResolution => {
    const palette = buildPaletteFromKey(key)!;
    return { paletteKey: key, palette, accentColor: PPTX_NAMED_PALETTES[key].main };
  };

  // ① 複合パレット名の変更先明示（コーラルオレンジ等が基本色「オレンジ」より先に捕捉されるよう最優先）
  if (isExplicitTarget("コーラルオレンジ|coral.orange|深緑.{0,4}コーラルオレンジ|深緑.{0,4}珊瑚橙|サンゴ.{0,4}オレンジ")) return paletteOf("coral_orange");
  if (isExplicitTarget("ティール.{0,4}コーラル|teal.coral|青緑.{0,4}コーラル"))                                              return paletteOf("teal_coral");
  if (isExplicitTarget("バーガンディ.{0,4}ゴールド|burgundy.gold|ワインレッド.{0,4}ゴールド"))                                return paletteOf("burgundy_gold");
  if (isExplicitTarget("深緑.{0,4}アンバー|forest.amber|フォレスト.{0,4}アンバー"))                                          return paletteOf("forest_amber");
  if (isExplicitTarget("ネイビー.{0,4}オレンジ|navy.orange"))                                                                 return paletteOf("navy_orange");
  if (isExplicitTarget("チャコール.{0,4}テラコッタ|charcoal.terra"))                                                          return paletteOf("charcoal_terra");

  // ② 基本色の変更先明示（複合名より後に評価 → コーラルオレンジが①で捕捉済みなのでオレンジ単体に安全にマッチ）
  // 赤・緑・紺はパレット全体を返す（単色だと色体系が崩れるため）
  // 青はパレット候補がなく単色指定（navy_orange は紺ベースで「青系」とは別）
  if (isExplicitTarget("赤|赤系|赤色|赤のトーン|レッド|red")) return paletteOf("burgundy_gold");
  if (isExplicitTarget("緑|緑系|グリーン|green"))             return paletteOf("forest_amber");
  if (isExplicitTarget("紺|ネイビー|navy"))                    return paletteOf("navy_orange");
  if (isExplicitTarget("青|青系|ブルー|blue"))  return { accentColor: "2F5597" };
  if (isExplicitTarget("紫|パープル|purple"))   return { accentColor: "7030A0" };
  if (isExplicitTarget("オレンジ|orange|橙"))   return { accentColor: "C55A11" };
  if (isExplicitTarget("黄|イエロー|yellow"))   return { accentColor: "BF9000" };
  if (isExplicitTarget("ピンク|pink"))          return { accentColor: "C0508A" };
  if (isExplicitTarget("グレー|gray|grey"))     return { accentColor: "666666" };

  let key: string | null = null;

  // 日本語複合名（最優先: navy_orange が navy 単体に誤マッチしないよう）
  if      (/(ネイビー.{0,2}オレンジ|ネイビーオレンジ)/.test(s))           key = "navy_orange";
  else if (/(深緑.{0,2}アンバー|深緑アンバー|フォレスト.{0,2}アンバー)/.test(s)) key = "forest_amber";
  else if (/(バーガンディ.{0,2}ゴールド|バーガンディゴールド|ワインレッド.{0,2}ゴールド)/.test(s)) key = "burgundy_gold";
  else if (/(ティール.{0,2}コーラル|ティールコーラル|青緑.{0,2}コーラル)/.test(s)) key = "teal_coral";
  else if (/(チャコール.{0,2}テラコッタ|チャコールテラコッタ)/.test(s))   key = "charcoal_terra";
  else if (/(深緑.{0,4}コーラルオレンジ|深緑.{0,4}珊瑚橙|コーラルオレンジ|コーラル.{0,4}オレンジ|サンゴ.{0,4}オレンジ|coral.orange|green.coral.orange)/.test(s)) key = "coral_orange";
  // 英語ID
  else if (/navy.orange/.test(t))    key = "navy_orange";
  else if (/forest.amber/.test(t))   key = "forest_amber";
  else if (/burgundy.gold/.test(t))  key = "burgundy_gold";
  else if (/teal.coral/.test(t))     key = "teal_coral";
  else if (/charcoal.terra/.test(t)) key = "charcoal_terra";
  else if (/coral.orange/.test(t))   key = "coral_orange";
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
  const paletteNumberRe = /([1-6１-６])番?[でに]/g;
  let m: RegExpExecArray | null;
  while ((m = paletteNumberRe.exec(s)) !== null) {
    const before = s.slice(Math.max(0, m.index - 10), m.index);
    if (/[Pp](?:age)?\s*$/.test(before) || /スライド\s*$/.test(before) || /ページ\s*$/.test(before)) continue;
    paletteNumMatch = m;
    break;
  }
  if (paletteNumMatch) {
    const d = paletteNumMatch[1];
    const n = d >= "1" && d <= "6" ? parseInt(d) : "１２３４５６".indexOf(d) + 1;
    const paletteKey = PPTX_PALETTE_KEYS[n - 1];
    if (paletteKey) {
      const palette = buildPaletteFromKey(paletteKey)!;
      return { paletteKey, palette, accentColor: PPTX_NAMED_PALETTES[paletteKey].main };
    }
  }

  // 基本色（パレット非該当）
  // 「赤」はバーガンディ×ゴールドパレット全体を適用（単色指定では色体系が崩れるため）
  if (/(赤|red)/.test(t)) {
    const palette = buildPaletteFromKey("burgundy_gold")!;
    return { paletteKey: "burgundy_gold", palette, accentColor: PPTX_NAMED_PALETTES["burgundy_gold"].main };
  }
  if (/(青|blue)/.test(t))             return { accentColor: "2F5597" };
  if (/(緑|green)/.test(t))            return { accentColor: "548235" };
  if (/(紫|purple)/.test(t))           return { accentColor: "7030A0" };
  if (/(オレンジ|orange|橙)/.test(t)) return { accentColor: "C55A11" };
  if (/(黄|yellow)/.test(t))           return { accentColor: "BF9000" };
  if (/(ピンク|pink)/.test(t))         return { accentColor: "C0508A" };
  if (/(グレー|gray|grey)/.test(t))    return { accentColor: "666666" };

  return null;
}
