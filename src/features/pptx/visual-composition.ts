/** Deterministic visual planning shared by generation and structural editing.
 * The input is a table of source facts; planning never adds or estimates facts.
 */
export type VisualComposition = { kind: "process" | "comparison" | "classification" | "bar"; columns: string[]; rows: string[][]; values?: number[] };
export const PPT_VISUAL_COMPOSITION_RULES = [
  "Organize each slide around one conclusion, then supporting facts, then qualifications/sources. Prefer short noun phrases over paragraphs; never remove a material qualification to make a stronger claim.",
  "Choose a diagram for the relationship, not decoration: a process uses ordered steps, classification uses named categories, a contrast uses two sides on the same comparison axes, numeric comparisons use a bar chart only with a common metric, period and unit.",
  "For process diagrams supply a table with columns 工程 and 内容, one row per step (2–5). For classification supply 分類 and 内容 (2–4 rows). For a two-sided comparison supply 項目 plus the two actual group names (2–4 rows). Copy facts from the source, keeping sequence and exceptions.",
  "For a quantitative comparison supply a two-column table: entity name, then the exact metric header containing period AND unit, e.g. 2025年度売上高（億円）. Values must be nonnegative and directly sourced on that same basis. If fiscal periods, segment/consolidated scope or units differ, use a descriptive table; do not manufacture a common basis, rank or chart value.",
  "Keep sources and caveats visible. Do not invent bar heights, barrier scores, icons with factual meaning, or a causal arrow unless the source states the relationship. Do not make every slide a table or a card grid.",
].join("\n");

export function buildVisualComposition(title: string, columns: string[], rows: string[][]): VisualComposition | undefined {
  if (!Array.isArray(columns) || !Array.isArray(rows) || rows.length < 2 || rows.length > 6 ||
      columns.some(c => typeof c !== "string" || !c.trim()) ||
      rows.some(r => !Array.isArray(r) || r.length !== columns.length || r.some(c => typeof c !== "string" || !c.trim()))) return;
  const base = { columns: [...columns], rows: rows.map(r => [...r]) };
  if (columns.length === 2 && rows.length <= 5 && /工程|ステップ|手順|段階/.test(columns[0])) return { ...base, kind: "process" };
  if (columns.length === 2 && rows.length <= 4 && /分類|区分|種類/.test(columns[0])) return { ...base, kind: "classification" };
  if (columns.length === 3 && rows.length <= 4 && /項目|比較軸|観点/.test(columns[0]) && /比較|対比|二極|違い/.test(title)) return { ...base, kind: "comparison" };
  if (columns.length !== 2 || !/売上|数量|件数|金額|人数/.test(columns[1]) ||
      !/(?:19|20)\d{2}(?:年|年度)/.test(columns[1]) ||
      !/[（(](?:円|千円|万円|百万円|億円|件|人|台|トン|t)[）)]/.test(columns[1])) return;
  // Require a single period in the metric header and no row-specific footnotes,
  // units, signs, percentages or estimated/ranged values.
  if ((columns[1].match(/(?:19|20)\d{2}/g) ?? []).length !== 1) return;
  if (!rows.every(r => /^(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(r[1]) && !/年度|月期|連結|単体|部門/.test(r[0]))) return;
  const values = rows.map(r => Number(r[1].replace(/,/g, "")));
  if (values.some(v => !Number.isFinite(v) || v > Number.MAX_SAFE_INTEGER) || Math.max(...values) <= 0) return;
  return { ...base, kind: "bar", values };
}

/** Company profiles need named rows, not a paragraph containing a list of names. */
export function requiresCompanyCards(instruction: string): boolean {
  return /(?:上位|主要|トップ|大手)[^。\n]{0,12}(?:企業|会社|プレイヤー)|業界(?:トップ|上位)\s*[0-9０-９一二三四五六七八九十]+/i.test(instruction);
}

export function hasCompanyRows(slides: unknown[]): boolean {
  return slides.length > 0 && slides.every((slide: any) =>
    slide?.kind === "table" && Array.isArray(slide.table?.columns) &&
    slide.table.columns.some((name: unknown) => typeof name === "string" && /企業|会社|プレイヤー/.test(name)) &&
    Array.isArray(slide.table?.rows) && slide.table.rows.length > 0);
}
