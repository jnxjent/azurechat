import { buildVisualComposition, type VisualComposition } from "./visual-composition";

/**
 * Structural PPTX edits handled by scripts/pptx_structure_edit.py:
 * slide insertion, filling an empty slide, and font-size changes.
 *
 * Detection is deliberately narrow. Requests that also ask for color, logo,
 * image or text replacement return null and keep using the existing
 * edit_pptx routes unchanged.
 */

export type PptxStructureContent = {
  title: string;
  kind: "bullets" | "table" | "stats" | "two_column";
  keyMessage?: string;
  stats?: { label: string; value: string; period?: string; unit?: string }[];
  bullets?: string[];
  table?: { columns: string[]; rows: string[][] };
  sources?: { label: string; url: string }[];
  composition?: VisualComposition;
  visualLayout?: "table";
  caveats?: string[];
};

export type PptxStructureOp =
  | ({ type: "add_slide"; position?: number; blank?: boolean } & Partial<PptxStructureContent>)
  | ({ type: "fill_slide"; slideIndex: number } & Partial<PptxStructureContent>)
  | { type: "font_size"; scope: "title" | "body" | "table" | "all"; mode: "absolute" | "relative"; value: number; slides?: number[] };

export type PptxStructureIntent =
  | { kind: "add_slide"; count: number; requestedCount: number; position?: number; blank: boolean; unsupported: string[] }
  | { kind: "fill_slide"; slideIndex: number; unsupported: string[] }
  | { kind: "font_size"; scope: "title" | "body" | "table" | "all"; mode: "absolute" | "relative"; value: number; slides?: number[]; unsupported: string[] };

export const MAX_ADDED_SLIDES = 3;

const KANJI_NUMBERS: Record<string, number> = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
const toNumber = (value: string) => KANJI_NUMBERS[value] ?? Number(value);

const SLIDE_ADD_RE =
  /(?:スライド|ページ)(?![にへ上右左下])[^。、\n]{0,8}(?:追加|挿入|足して|加えて)|(?:追加|挿入)[^。、\n]{0,6}(?:スライド|ページ)|(?:\d+|[一二三四五六七八九十])\s*枚[^。、\n]{0,4}(?:追加|挿入|足して|加えて)/;
const FONT_SIZE_RE =
  /フォントサイズ|文字(?:の)?サイズ|字(?:の)?(?:大きさ|サイズ)|\d+(?:\.\d+)?\s*(?:pt|ポイント)|(?:文字|字|フォント|タイトル|見出し|本文|表)[^。\n]{0,6}(?:大きく|小さく)/i;
// Other edit kinds keep the existing edit_pptx routing (no behavior change).
const OTHER_EDIT_RE =
  /(?<![特景])色|カラー|color|ロゴ|logo|画像|写真|image|背景|テーマ|アクセント|字体|書体|游|メイリオ|ゴシック|明朝|「[^」]+」[^。\n]{0,12}(?:変更|変え|修正|差し替え|置き換え|にして)/i;
// Existing-slide content edits (bullet_add / 増量). A new slide may still use bullets.
const EXISTING_CONTENT_EDIT_RE = /箇条書き|bullet|ブレット|増量|詳しく/i;
const UNSUPPORTED_RE: { re: RegExp; label: string }[] = [
  { re: /レイアウト.{0,6}最適化|重なり.{0,4}解消|配置.{0,4}(?:修正|変更|調整)|再レイアウト|位置.{0,4}調整/, label: "レイアウト最適化・shape移動" },
  { re: /スピーカーノート|ノート.{0,4}(?:追加|冒頭|末尾|記録)|speaker\s*note/i, label: "スピーカーノート追加" },
  { re: /再構成|作り直し|全体.{0,6}(?:見直し|再生成)/, label: "内容の再構成・作り直し" },
];

function pageNumbers(text: string): number[] {
  const pages: number[] = [];
  const add = (n: number) => { if (n >= 1 && !pages.includes(n)) pages.push(n); };
  Array.from(text.matchAll(/(?:P|ページ|スライド)\s*(\d{1,3})(?![\d枚])/gi)).forEach((m) => add(Number(m[1])));
  Array.from(text.matchAll(/(\d{1,3})\s*(?:枚目|ページ目)/g)).forEach((m) => add(Number(m[1])));
  return pages;
}

function insertPosition(text: string): number | undefined {
  let m = text.match(/最後から\s*(\d{1,3})\s*(?:枚目|番目)/);
  if (m) return -Number(m[1]);
  m = text.match(/(?:P|ページ|スライド)\s*(\d{1,3})\s*(?:枚目)?\s*の\s*(?:後|あと|次)/i) ?? text.match(/(\d{1,3})\s*(?:枚目|ページ目)\s*の\s*(?:後|あと|次)/);
  if (m) return Number(m[1]);
  m = text.match(/(?:P|ページ|スライド)\s*(\d{1,3})\s*(?:枚目)?\s*の\s*前/i) ?? text.match(/(\d{1,3})\s*(?:枚目|ページ目)\s*の\s*前/);
  if (m) return Math.max(0, Number(m[1]) - 1);
  m = text.match(/(\d{1,3})\s*(?:枚目|ページ目)\s*(?:に|として)/);
  if (m) return Math.max(0, Number(m[1]) - 1);
  if (/(?:表紙|1枚目|P1)\s*の\s*(?:後|あと|次)|2枚目(?:に|として)/i.test(text)) return 1;
  if (/先頭|一番前|最初に/.test(text)) return 0;
  return undefined; // end of deck
}

export function detectPptxStructureEdit(message: string): PptxStructureIntent | null {
  const text = message.normalize("NFKC");
  if (OTHER_EDIT_RE.test(text)) return null;
  const unsupported = UNSUPPORTED_RE.filter(({ re }) => re.test(text)).map(({ label }) => label);

  const emptyTarget = /空白|空の|空き|白紙|何もない/.test(text);
  const pages = pageNumbers(text);
  const fill =
    emptyTarget && pages.length === 1 &&
    /(?:に|へ)[^。\n]{0,40}(?:入れ|書|記載|追加|載せ|まとめ)/.test(text) &&
    !/(?:後|あと|次|前)\s*に[^。\n]{0,20}(?:スライド|ページ)[^。\n]{0,8}(?:追加|挿入)/.test(text) &&
    !/(?:空白|空の|白紙)\s*(?:の)?\s*(?:スライド|ページ)\s*を\s*(?:\d+\s*枚)?\s*(?:追加|挿入)/.test(text);
  if (fill) return { kind: "fill_slide", slideIndex: pages[0] - 1, unsupported };

  if (SLIDE_ADD_RE.test(text) && /pptx?|パワーポイント|パワポ|スライド|ページ|資料|枚/i.test(text) &&
      !/(?:スライド|ページ|P\s*\d+)\s*(?:に|へ)[^。\n]{0,12}箇条書き/i.test(text)) {
    const countMatch = text.match(/(\d+|[一二三四五六七八九十])\s*枚[^。、\n]{0,4}(?:追加|挿入|足して|加えて)|(?:スライド|ページ)を?\s*(\d+|[一二三四五六七八九十])\s*枚/);
    const requestedCount = countMatch ? toNumber(countMatch[1] ?? countMatch[2]) : 1;
    return {
      kind: "add_slide",
      requestedCount,
      count: Math.max(1, Math.min(MAX_ADDED_SLIDES, requestedCount)),
      position: insertPosition(text),
      blank: emptyTarget && !/(?:情報|内容|結果|回答|まとめ|表|一覧|について)/.test(text),
      unsupported,
    };
  }

  if (EXISTING_CONTENT_EDIT_RE.test(text)) return null;
  if (FONT_SIZE_RE.test(text) && !/フォント\s*(?:を|は)\s*[^。\n]{1,12}(?:に変更|に変え|にして)(?![^。\n]*(?:pt|ポイント|大き|小さ))/.test(text)) {
    const scope = /表/.test(text) ? "table" : /タイトル|見出し/.test(text) ? "title" : /本文/.test(text) ? "body" : "all";
    const absolute = text.match(/(\d+(?:\.\d+)?)\s*(?:pt|ポイント)/i);
    const step = /かなり|もっと|大幅/.test(text) ? 4 : 2;
    const smaller = /小さく|縮小|下げ/.test(text);
    if (!absolute && !/大きく|小さく|拡大|縮小|上げ|下げ/.test(text)) return null;
    return {
      kind: "font_size",
      scope,
      mode: absolute ? "absolute" : "relative",
      value: absolute ? Number(absolute[1]) : smaller ? -step : step,
      ...(pages.length ? { slides: pages.map((n) => n - 1) } : {}),
      unsupported,
    };
  }
  return null;
}

/** True for chat requests that target the thread's PPTX (used for routing). */
export function isPptxStructureEditRequest(message: string): boolean {
  // 「ページ」だけでは Wiki 等の通常依頼と区別できないため、PPTを示す語を必須にする。
  return detectPptxStructureEdit(message) !== null && /pptx?|パワーポイント|パワポ|スライド|資料/i.test(message.normalize("NFKC"));
}

const str = (value: unknown, max: number) => String(value ?? "").trim().slice(0, max);

function sanitizeContent(op: Record<string, unknown>): Partial<PptxStructureContent> {
  const out: Partial<PptxStructureContent> = { title: str(op.title, 60).replace(/[。．.]+$/, "") };
  if (typeof op.keyMessage === "string") out.keyMessage = str(op.keyMessage, 30);
  if (Array.isArray(op.stats)) out.stats = op.stats.slice(0, 4).map((raw) => {
    const item = (raw && typeof raw === "object" ? raw : {}) as Record<string, unknown>;
    return { label: str(item.label, 40), value: str(item.value, 30), period: str(item.period, 30), unit: str(item.unit, 20) };
  }).filter(item => item.label && item.value);
  if (op.kind === "stats" && out.stats && out.stats.length >= 2) out.kind = "stats";
  const table = op.table as { columns?: unknown; rows?: unknown } | undefined;
  if ((op.kind === "table" || op.kind === "two_column") && table && Array.isArray(table.columns) && Array.isArray(table.rows)) {
    const columns = table.columns.map((c) => str(c, 30)).filter(Boolean).slice(0, 5);
    const rows = table.rows.filter(Array.isArray).map((row) => (row as unknown[]).map((c) => str(c, 80)).slice(0, columns.length)).filter((row) => row.some(Boolean)).slice(0, 6);
    if (columns.length && rows.length) Object.assign(out, { kind: "table", table: { columns, rows } });
  } else if (Array.isArray(op.bullets)) {
    const bullets = op.bullets.map((b) => str(b, 120)).filter(Boolean).slice(0, 8);
    if (bullets.length) Object.assign(out, { kind: "bullets", bullets });
  }
  if (op.kind === "two_column" && out.table && Array.isArray(op.bullets)) {
    out.kind = "two_column";
    out.bullets = op.bullets.map(b => str(b, 80)).filter(Boolean).slice(0, 4);
    if (!out.bullets.length) out.kind = "table";
  }
  if (Array.isArray(op.caveats)) {
    out.caveats = op.caveats.map(c => str(c, 120)).filter(Boolean).slice(0, 2);
  }
  if (Array.isArray(op.sources)) {
    const sources = op.sources
      .map((s) => ({ label: str((s as Record<string, unknown>)?.label, 80), url: str((s as Record<string, unknown>)?.url, 500) }))
      .filter((s) => /^https?:\/\/\S+$/i.test(s.url))
      .slice(0, 3);
    if (sources.length) out.sources = sources;
  }
  if (op.visualLayout === "table") out.visualLayout = "table";
  if (out.table && !out.visualLayout) out.composition = buildVisualComposition(out.title ?? "", out.table.columns, out.table.rows);
  return out;
}

const normalizeForMatch = (value: string) =>
  value.normalize("NFKC").replace(/株式会社|\(株\)|㈱|有限会社/g, "").replace(/[\s,，、・]/g, "").toLowerCase();

/**
 * Keep only content grounded in the source text: every number must appear in
 * the source, the key (name) column of a table must appear verbatim, and
 * source URLs must be ones that were actually returned.
 */
export function verifyPptxStructureContent(slides: unknown, sourceText: string): { slides: PptxStructureContent[]; warnings: string[] } {
  const source = normalizeForMatch(sourceText);
  const numbersOk = (text: string) =>
    (text.normalize("NFKC").replace(/(\d),(\d)/g, "$1$2").match(/\d+(?:\.\d+)?/g) ?? []).every((n) => new RegExp(`(?<![\\d.])${n.replace(/\./g, "\\.")}(?![\\d.])`).test(source));
  const warnings: string[] = [];
  const kept: PptxStructureContent[] = [];
  for (const raw of Array.isArray(slides) ? slides : []) {
    const content = sanitizeContent({ ...(raw as Record<string, unknown>) });
    if (!content.title) continue;
    if (content.keyMessage && !source.includes(normalizeForMatch(content.keyMessage))) {
      delete content.keyMessage;
      warnings.push(`「${content.title}」: 根拠を確認できない結論文を掲載していません。`);
    }
    if (content.stats) {
      // Label, value, unit and period must co-occur in one source line.
      content.stats = content.stats.filter(item => sourceText.split(/\r?\n/).some(line =>
        [item.label, item.value, item.period, item.unit].filter(Boolean).every(value => normalizeForMatch(line).includes(normalizeForMatch(value!)))));
      if (content.kind === "stats" && content.stats.length < 2) content.kind = undefined;
    }
    if (content.kind === "two_column" && content.bullets) content.bullets = content.bullets.filter(b => source.includes(normalizeForMatch(b)));
    if (content.caveats) content.caveats = content.caveats.filter(numbersOk);
    let dropped = 0;
    if ((content.kind === "table" || content.kind === "two_column") && content.table) {
      const keyColumn = content.table.columns.findIndex((_, c) => content.table!.rows.some((row) => row[c] && !/^[\d.\s%]+$/.test(row[c])));
      const rows = content.table.rows.filter((row) => {
        const ok = row.every(numbersOk) && (keyColumn < 0 || !row[keyColumn] || source.includes(normalizeForMatch(row[keyColumn])));
        if (!ok) dropped++;
        return ok;
      });
      content.table = { ...content.table, rows };
      content.composition = content.visualLayout ? undefined : buildVisualComposition(content.title ?? "", content.table.columns, rows);
      if (content.composition?.kind === "bar") {
        const exact = (value: string) => value.normalize("NFKC").replace(/[\s,|]/g, "").toLowerCase();
        const lines = sourceText.split(/\r?\n/).map(exact);
        if (!exact(sourceText).includes(exact(content.table.columns[1])) ||
            !rows.every(row => lines.some(line => line.includes(row.map(exact).join(""))))) {
          // Missing entity/value association or comparison basis: keep the
          // readable table instead of visually asserting a quantitative rank.
          content.composition = undefined;
          content.visualLayout = "table";
          warnings.push(`「${content.title}」: 同じ比較条件と企業別の数値対応を元文で確認できないため、グラフにせず表で表示します。`);
        }
      }
      if (!rows.length) content.kind = undefined;
      else if (content.kind === "two_column" && !content.bullets?.length) content.kind = "table";
    } else if (content.kind === "bullets" && content.bullets) {
      content.bullets = content.bullets.filter((b) => {
        const ok = numbersOk(b);
        if (!ok) dropped++;
        return ok;
      });
      if (!content.bullets.length) content.kind = undefined;
    }
    if (content.sources) content.sources = content.sources.filter((s) => sourceText.includes(s.url));
    if (dropped) warnings.push(`「${content.title}」: 元の回答で確認できない値を含む${dropped}件を掲載していません。`);
    if (content.kind) kept.push(content as PptxStructureContent);
  }
  return { slides: kept, warnings };
}

/** Whitelist the ops passed to the Python script. */
export function sanitizePptxStructureOps(value: unknown): PptxStructureOp[] {
  if (!Array.isArray(value)) return [];
  const ops: PptxStructureOp[] = [];
  for (const raw of value.slice(0, 6)) {
    if (!raw || typeof raw !== "object") continue;
    const op = raw as Record<string, unknown>;
    if (op.type === "add_slide") {
      const position = Number.isInteger(op.position) ? Math.max(-200, Math.min(200, op.position as number)) : undefined;
      ops.push({ type: "add_slide", ...(position !== undefined ? { position } : {}), ...(op.blank === true ? { blank: true, title: str(op.title, 60) } : sanitizeContent(op)) });
    } else if (op.type === "fill_slide" && Number.isInteger(op.slideIndex) && (op.slideIndex as number) >= 0) {
      ops.push({ type: "fill_slide", slideIndex: op.slideIndex as number, ...sanitizeContent(op) });
    } else if (op.type === "font_size" && ["title", "body", "table", "all"].includes(String(op.scope))) {
      const value = Number(op.value);
      if (!Number.isFinite(value) || value === 0) continue;
      const slides = Array.isArray(op.slides) ? op.slides.filter((n) => Number.isInteger(n) && (n as number) >= 0).slice(0, 100) as number[] : undefined;
      ops.push({ type: "font_size", scope: op.scope as "title" | "body" | "table" | "all", mode: op.mode === "absolute" ? "absolute" : "relative", value: Math.max(-40, Math.min(200, value)), ...(slides?.length ? { slides } : {}) });
    }
  }
  return ops;
}
