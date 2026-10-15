import PptxGenJS from "pptxgenjs";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { buildPaletteFromKey } from "@/features/pptx/palette";

/**
 * Excelの比較表（行=項目、列=現状/案1/案2…）を、見栄えのよいPowerPoint1枚に清書する。
 * LLMを使わず決定的に処理し、セルの文言は変更しない。
 */

export interface ExcelTableSlideSpec {
  title: string;
  headers: string[];
  rows: string[][];
  /** 「現状」など比較の基準列。案の列で基準と異なる行を強調する。無い場合は null。 */
  baselineColumn: number | null;
}

type ItemKind = "lead" | "bullet" | "note";
interface CellItem { kind: ItemKind; text: string; changed: boolean }

const MAX_COLUMNS = 6;
const MAX_ROWS = 15;

const BASELINE_HEADER_PATTERN = /現状|現行|変更前|従来|As[\s-]?Is/i;
const BULLET_PATTERN = /^[・•●◆■□○◇\-－‐*＊]\s*/;
const SENTENCE_END_PATTERN = /[。．.）)」』!！?？：:]$/;

const cellText = (value: unknown): string =>
  String(value ?? "").replace(/\r\n?/g, "\n").replace(/[ \t　]+$/gm, "").trim();

/** 最初の表らしいシートを読み、見出し行・項目行に整理する。表が無ければ null。 */
export function parseExcelTable(buffer: Buffer, fallbackTitle: string): ExcelTableSlideSpec | null {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const raw = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false, blankrows: false })
      .map(row => (row as unknown[]).map(cellText));
    const spec = buildSpec(raw, fallbackTitle);
    if (spec) return spec;
  }
  return null;
}

export function buildSpec(raw: string[][], fallbackTitle: string): ExcelTableSlideSpec | null {
  const width = Math.max(0, ...raw.map(row => row.length));
  const usedColumns = Array.from({ length: width }, (_, c) => c).filter(c => raw.some(row => row[c]));
  const grid = raw.map(row => usedColumns.map(c => row[c] ?? "")).filter(row => row.some(Boolean));

  const headerIndex = grid.findIndex(row => row.filter(Boolean).length >= 2);
  if (headerIndex < 0) return null;
  const titleRow = grid.slice(0, headerIndex).find(row => row.filter(Boolean).length === 1);
  const title = titleRow?.find(Boolean) || fallbackTitle;

  const headers = grid[headerIndex];
  if (headers.length < 2 || headers.length > MAX_COLUMNS) return null;

  // 項目列が空の行は、直前の項目の続き（Excelの縦結合・複数行入力）として連結する。
  const rows: string[][] = [];
  for (const row of grid.slice(headerIndex + 1)) {
    const previous = rows[rows.length - 1];
    if (!row[0] && previous) {
      row.forEach((value, c) => {
        if (value) previous[c] = previous[c] ? `${previous[c]}\n${value}` : value;
      });
    } else {
      rows.push([...row]);
    }
  }
  if (!rows.length || rows.length > MAX_ROWS) return null;

  const baseline = headers.findIndex((header, c) => c >= 1 && BASELINE_HEADER_PATTERN.test(header));
  return {
    title,
    headers,
    rows,
    baselineColumn: baseline >= 1 && baseline < headers.length - 1 ? baseline : null,
  };
}

const comparable = (text: string): string =>
  text.normalize("NFKC").replace(BULLET_PATTERN, "").replace(/[\s、。,.]/g, "");

const parenDepth = (text: string): number =>
  (text.match(/[（(]/g)?.length ?? 0) - (text.match(/[）)]/g)?.length ?? 0);

/** セル文字列を、リード文・箇条書き・補足（括弧書き／箇条書きの説明）に分ける。 */
export function toCellItems(text: string, baselineText: string | null): CellItem[] {
  const items: CellItem[] = [];
  let openParen = 0;
  let underBullet = false;
  for (const line of text.split("\n").map(value => value.trim()).filter(Boolean)) {
    const last = items[items.length - 1];
    if (openParen > 0 && last) {
      last.text += line;
      openParen += parenDepth(line);
      continue;
    }
    if (BULLET_PATTERN.test(line)) {
      items.push({ kind: "bullet", text: line.replace(BULLET_PATTERN, ""), changed: false });
      underBullet = true;
      continue;
    }
    if (/^[（(]/.test(line) || underBullet) {
      // 折り返しで分かれた補足は1つにつなぐ。
      if (last?.kind === "note" && !SENTENCE_END_PATTERN.test(last.text) && !/^[（(]/.test(line)) last.text += line;
      else items.push({ kind: "note", text: line, changed: false });
      openParen = Math.max(0, parenDepth(line));
      continue;
    }
    items.push({ kind: "lead", text: line, changed: false });
  }
  if (baselineText !== null) {
    const base = comparable(baselineText);
    for (const item of items) {
      if (item.kind === "lead") item.changed = !base.includes(comparable(item.text));
    }
  }
  return items;
}

// ── 描画 ───────────────────────────────────────────────
const W = 13.33;
const H = 7.5;
const MARGIN_X = 0.45;
const TABLE_TOP = 1.3;
const TABLE_BOTTOM_MARGIN = 0.35;
const LABEL_COLUMN_W = 1.55;

const COLOR = {
  title: "1F4E3D",
  dark: "1F4E3D",
  teal: "2A9D8F",
  baselineHead: "7F8C86",
  labelBg: "EEF4F1",
  border: "D9DEDB",
  text: "333333",
  note: "6B6B6B",
  highlight: "B8741A",
};

/** 全角を1、半角を0.55文字として幅を見積もる。 */
const visualLength = (text: string): number =>
  Array.from(text).reduce((sum, ch) => sum + (/[\u0000-ÿ｡-ﾟ]/.test(ch) ? 0.55 : 1), 0);

function estimateCellHeight(items: CellItem[], widthInches: number, fontSize: number): number {
  const usable = Math.max(0.5, widthInches - 0.2);
  let height = 0.14;
  for (const item of items) {
    const size = item.kind === "note" ? fontSize - 1.5 : fontSize;
    const prefix = item.kind === "lead" ? 0 : 1.5;
    const perLine = Math.max(4, Math.floor((usable * 72) / size));
    const lines = Math.max(1, Math.ceil((visualLength(item.text) + prefix) / perLine));
    height += (lines * size * 1.3) / 72;
  }
  return height;
}

function columnWidths(spec: ExcelTableSlideSpec): number[] {
  const rest = W - MARGIN_X * 2 - LABEL_COLUMN_W;
  const weights = spec.headers.slice(1).map((header, i) => {
    const volume = spec.rows.reduce((sum, row) => sum + visualLength(row[i + 1] ?? ""), visualLength(header));
    return Math.pow(Math.max(20, volume), 0.6);
  });
  const total = weights.reduce((a, b) => a + b, 0);
  const minimum = Math.min(1.6, rest / weights.length);
  const raw = weights.map(weight => Math.max(minimum, (rest * weight) / total));
  const scale = rest / raw.reduce((a, b) => a + b, 0);
  return [LABEL_COLUMN_W, ...raw.map(value => value * scale)];
}

function headerFill(spec: ExcelTableSlideSpec, column: number, colors = COLOR): string {
  if (column === 0) return colors.dark;
  if (column === spec.baselineColumn) return colors.baselineHead;
  const proposalIndex = column - 1 - (spec.baselineColumn !== null && column > spec.baselineColumn ? 1 : 0);
  return proposalIndex % 2 === 0 ? colors.teal : colors.dark;
}

function runsFor(items: CellItem[], fontSize: number, colors = COLOR): PptxGenJS.TextProps[] {
  return items.map((item, index) => {
    const breakLine = index < items.length - 1;
    if (item.kind === "bullet") {
      return { text: `• ${item.text}`, options: { bold: true, color: colors.dark, fontSize, breakLine, paraSpaceBefore: index ? 3 : 0 } };
    }
    if (item.kind === "note") {
      return { text: `　${item.text}`, options: { color: colors.note, fontSize: fontSize - 1.5, breakLine } };
    }
    return {
      text: item.text,
      options: item.changed
        ? { bold: true, color: colors.highlight, fontSize, breakLine }
        : { color: colors.text, fontSize, breakLine },
    };
  });
}

/** 「体制（出席者、人数）」のような長い項目名は括弧の前で改行する。 */
const labelText = (label: string): string =>
  visualLength(label) > 6 ? label.replace(/^([^（(\n]+)([（(])/, "$1\n$2") : label;

const TABLE_AVAILABLE_H = H - TABLE_TOP - TABLE_BOTTOM_MARGIN;
const HEADER_ROW_H = 0.5;
/** 行高の見積もり誤差に備えた余裕。 */
const HEIGHT_SAFETY = 1.1;

export interface ExcelTableSlideLayout {
  fontSize: number;
  widths: number[];
  rowHeights: number[];
  cells: CellItem[][][];
}

/** 1枚に収まる最大の文字サイズと行高を求める。最小サイズでも収まらなければ null。 */
export function layoutExcelTableSlide(spec: ExcelTableSlideSpec): ExcelTableSlideLayout | null {
  const widths = columnWidths(spec);
  const cells = spec.rows.map(row => spec.headers.map((_, c) => {
    if (c === 0) return [] as CellItem[];
    const isProposal = spec.baselineColumn !== null && c > spec.baselineColumn;
    return toCellItems(row[c] ?? "", isProposal ? row[spec.baselineColumn!] ?? "" : null);
  }));
  for (let size = 12; size >= 7; size -= 0.5) {
    const rowHeights = spec.rows.map((row, r) => HEIGHT_SAFETY * Math.max(
      0.36,
      estimateCellHeight(labelText(row[0] ?? "").split("\n").map(text => ({ kind: "lead" as const, text, changed: false })), widths[0], size),
      ...cells[r].map((items, c) => (c === 0 ? 0 : estimateCellHeight(items, widths[c], size))),
    ));
    const total = HEADER_ROW_H + rowHeights.reduce((a, b) => a + b, 0);
    if (total <= TABLE_AVAILABLE_H) {
      // 余白は最終行に回し、表をスライド下端まで揃える。
      rowHeights[rowHeights.length - 1] += TABLE_AVAILABLE_H - total;
      return { fontSize: size, widths, rowHeights, cells };
    }
  }
  return null;
}

export async function renderExcelTableSlide(spec: ExcelTableSlideSpec, fontFace = "Meiryo", paletteKey?: string): Promise<Buffer> {
  const palette = paletteKey ? buildPaletteFromKey(paletteKey) : undefined;
  const colors = palette ? { ...COLOR, title: palette.titleBg, dark: palette.headerBg, teal: palette.accentB, labelBg: palette.tableAltBg, border: palette.border, text: palette.bodyText, note: palette.mutedText, highlight: palette.accentB } : COLOR;
  const layout = layoutExcelTableSlide(spec);
  if (!layout) throw new Error("表がスライド1枚に収まりません。");
  const { fontSize, widths, rowHeights, cells } = layout;

  const border = { type: "solid" as const, pt: 0.75, color: colors.border };
  const borders = [border, border, border, border] as [typeof border, typeof border, typeof border, typeof border];
  const headerRow: PptxGenJS.TableCell[] = spec.headers.map((header, c) => ({
    text: header,
    options: {
      bold: true, color: "FFFFFF", fontSize: fontSize + 1.5, align: "center", valign: "middle",
      fill: { color: headerFill(spec, c, colors) }, border: borders, margin: [4, 6, 4, 6],
    },
  }));
  const bodyRows: PptxGenJS.TableCell[][] = spec.rows.map((row, r) => spec.headers.map((_, c) => (
    c === 0
      ? {
          text: labelText(row[0] ?? ""),
          options: {
            bold: true, color: colors.dark, fontSize, align: "center", valign: "middle",
            fill: { color: colors.labelBg }, border: borders, margin: [5, 6, 5, 6],
          },
        }
      : {
          text: cells[r][c].length ? runsFor(cells[r][c], fontSize, colors) : "",
          options: {
            align: "left", valign: "top", fill: { color: "FFFFFF" }, border: borders, margin: [5, 8, 5, 8],
          },
        }
  )));

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.title = spec.title;
  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addText(spec.title, {
    x: MARGIN_X, y: 0.35, w: W - MARGIN_X * 2, h: 0.6,
    fontFace, fontSize: 24, bold: true, color: colors.title, valign: "middle", fit: "shrink",
  });

  const hasHighlight = cells.some(row => row.some(items => items.some(item => item.changed)));
  if (hasHighlight) {
    slide.addShape("rect", {
      x: W - MARGIN_X - 2.35, y: 0.98, w: 0.17, h: 0.17,
      fill: { color: colors.highlight }, line: { color: colors.highlight, width: 0 },
    });
    slide.addText(`= ${spec.headers[spec.baselineColumn!]}からの変更点`, {
      x: W - MARGIN_X - 2.12, y: 0.9, w: 2.12, h: 0.32,
      fontFace, fontSize: 10, color: colors.note, valign: "middle",
    });
  }

  slide.addTable([headerRow, ...bodyRows], {
    x: MARGIN_X, y: TABLE_TOP, w: W - MARGIN_X * 2,
    colW: widths, rowH: [HEADER_ROW_H, ...rowHeights], fontFace,
  });

  const buffer = (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
  // This PptxGenJS version assigns table IDs independently of text/shape IDs.
  // This renderer has no animations or connectors referencing shape IDs;
  // preserve existing IDs and renumber only duplicates in its single slide.
  const zip = await JSZip.loadAsync(buffer);
  const slideFile = zip.file("ppt/slides/slide1.xml");
  if (!slideFile) throw new Error("生成したPowerPointのスライドが見つかりません。");
  const xml = await slideFile.async("string");
  const ids = Array.from(xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g), match => Number(match[1]));
  let nextId = Math.max(0, ...ids) + 1;
  const seen = new Set<number>();
  const patched = xml.replace(/(<p:cNvPr\b[^>]*\bid=")(\d+)(")/g, (tag, prefix, value, suffix) => {
    const id = Number(value);
    if (seen.has(id)) return `${prefix}${nextId++}${suffix}`;
    seen.add(id);
    return tag;
  });
  zip.file("ppt/slides/slide1.xml", patched);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
