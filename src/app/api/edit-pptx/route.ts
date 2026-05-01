export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import JSZip from "jszip";
import { promises as fs } from "node:fs";
import { constants as fsConstants } from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import * as XLSX from "xlsx";
import {
  BlobSASPermissions,
  BlobServiceClient,
  StorageSharedKeyCredential,
  generateBlobSASQueryParameters,
} from "@azure/storage-blob";
import { OpenAIInstance, OpenAIDALLEInstance } from "@/features/common/services/openai";
import { uniqueId } from "@/features/common/util";

const execFileAsync = promisify(execFile);

type SlideSummary = {
  slideIndex: number;
  texts: string[];
};

type ImageInsert = {
  slideIndex: number;
  imagePrompt: string;
  position?: "top-right" | "top-left" | "bottom-right" | "bottom-left" | "center";
  nearText?: string;   // このテキストを含む Shape の隣に配置（指定時は position より優先）
  anchorSide?: "left" | "right" | "above" | "below"; // nearText Shape のどちら側か（default: "right"）
  widthPct?: number;
  imagePath?: string; // set after DALL-E generation, before Python call
};

type EditPlan = {
  deckEdits?: {
    accentColor?: string | null;
    fontFace?: string | null;
    preserveTextColors?: boolean;
  };
  slideEdits?: Array<{
    slideIndex: number;
    replaceText?: Array<{
      find: string;
      replace: string;
    }>;
  }>;
  imageInserts?: ImageInsert[];
};

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"');
}

function normalizeHexColor(value: string | undefined | null): string | undefined {
  const normalized = String(value ?? "").replace("#", "").trim().toUpperCase();
  return /^[0-9A-F]{6}$/.test(normalized) ? normalized : undefined;
}

function extractShapeTextByParagraph(shapeXml: string): string[] {
  const paragraphs: string[] = [];
  const paraRe = /<a:p(?:\s[^>]*)?>[\s\S]*?<\/a:p>/g;
  let pm: RegExpExecArray | null;
  while ((pm = paraRe.exec(shapeXml)) !== null) {
    const paraXml = pm[0];
    const runRe = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
    let rm: RegExpExecArray | null;
    let paraText = "";
    while ((rm = runRe.exec(paraXml)) !== null) {
      paraText += decodeXmlEntities(rm[1]);
    }
    if (paraText.trim()) paragraphs.push(paraText);
  }
  return paragraphs;
}

function splitIntoShapes(slideXml: string): string[] {
  const shapes: string[] = [];
  const re = /<p:sp[\s>][\s\S]*?<\/p:sp>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(slideXml)) !== null) shapes.push(m[0]);
  return shapes;
}

async function extractSlideSummaries(pptxBuffer: Buffer): Promise<SlideSummary[]> {
  const zip = await JSZip.loadAsync(pptxBuffer);
  const slideEntries = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)/)?.[1] ?? "0", 10);
      const numB = parseInt(b.match(/slide(\d+)/)?.[1] ?? "0", 10);
      return numA - numB;
    });

  const summaries: SlideSummary[] = [];
  for (let si = 0; si < slideEntries.length; si++) {
    const xml = await zip.files[slideEntries[si]].async("string");
    const shapes = splitIntoShapes(xml);
    const texts = shapes
      .flatMap((shapeXml) => extractShapeTextByParagraph(shapeXml))
      .map((text) => text.trim())
      .filter(Boolean);
    summaries.push({ slideIndex: si, texts: texts.slice(0, 20) });
  }
  return summaries;
}

function parseDirectAccentColor(instruction: string): string | undefined {
  const t = instruction.toLowerCase();
  if (/(赤|red)/.test(t)) return "C00000";
  if (/(青|blue)/.test(t)) return "2F5597";
  if (/(緑|green)/.test(t)) return "548235";
  if (/(紫|purple)/.test(t)) return "7030A0";
  if (/(オレンジ|orange|橙)/.test(t)) return "C55A11";
  if (/(黄|yellow)/.test(t)) return "BF9000";
  if (/(ピンク|pink)/.test(t)) return "C0508A";
  return undefined;
}

async function buildEditPlan(slides: SlideSummary[], instruction: string): Promise<EditPlan> {
  const openai = OpenAIInstance();

  const systemPrompt = `You convert a natural-language PowerPoint editing request into a safe JSON edit plan.
Return JSON only in this shape:
{
  "deckEdits": {
    "accentColor": "RRGGBB" | null,
    "fontFace": string | null,
    "preserveTextColors": boolean
  },
  "slideEdits": [
    {
      "slideIndex": number,
      "replaceText": [{ "find": string, "replace": string }]
    }
  ],
  "imageInserts": [
    {
      "slideIndex": number,
      "imagePrompt": string,
      "nearText": string,
      "anchorSide": "left" | "right" | "above" | "below",
      "position": "top-right" | "top-left" | "bottom-right" | "bottom-left" | "center",
      "widthPct": number
    }
  ]
}

Rules:
- If the user asks to change the deck color or tone, set deckEdits.accentColor to a practical 6-digit hex.
- Deck color change means accent shapes, fills, and lines. It does not mean changing body text color.
- preserveTextColors should usually be true unless the user explicitly asks to recolor text.
- Only emit text replacements when the user explicitly wants wording changed.
- slideIndex is zero-based.
- If the user asks to add any visual element to a slide — including icons, illustrations, images, photos, vehicles, animals, objects, characters, or marks — populate imageInserts[]. Also treat descriptive phrases like "○○の絵" / "○○のイラスト" / "○○（説明）" as image requests even if the word "画像" or "image" is not used.
  - imagePrompt: concise English DALL-E prompt (e.g. "white industrial garbage truck packer, side view, clean flat illustration, white background").
  - nearText: if the user says "next to X" / "横に" / "〜の横" / "〜の隣" etc., set nearText to the shortest unique text string from the slide. The image will be placed adjacent to that shape.
  - anchorSide: "right" for "横に/右に", "left" for "左に", "above" for "上に", "below" for "下に". Default "right".
  - position: fallback when nearText shape is not found. Use "center" for main illustrations, "top-right" for decorative corner icons. Never use "top-left".
  - widthPct: 6-8 for small inline icons, 12-15 for corner decorations, 30-50 for main illustrations. Default 35 for large illustrations, 13 for decorative icons.
- Do NOT add descriptive text to the slide via replaceText when the user is clearly describing a visual/image to insert.
- Keep the JSON minimal. Use null or [] when not needed.`;

  const userPrompt = `Instruction:
${instruction}

Slides:
${JSON.stringify(slides, null, 2)}

Return JSON only.`;

  const res = await openai.chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 1500,
  });

  const content = res.choices[0]?.message?.content ?? "{}";
  const parsed = JSON.parse(content) as EditPlan;
  parsed.deckEdits ??= {};
  parsed.slideEdits ??= [];

  const directAccent = parseDirectAccentColor(instruction);
  // imageInserts がある場合、画像プロンプト内（「青系アイコン」等）の色を
  // テキスト全体の色変更と誤認しないよう、parseDirectAccentColor を適用しない。
  // LLM が明示的に accentColor を返した場合のみテキスト色を変更する。
  const hasImageInserts = (parsed.imageInserts?.length ?? 0) > 0;
  if (!normalizeHexColor(parsed.deckEdits.accentColor) && directAccent && !hasImageInserts) {
    parsed.deckEdits.accentColor = directAccent;
  } else {
    parsed.deckEdits.accentColor = normalizeHexColor(parsed.deckEdits.accentColor) ?? null;
  }
  parsed.deckEdits.preserveTextColors = parsed.deckEdits.preserveTextColors !== false;

  return parsed;
}

function parseAzureBlobUrl(fileUrl: string): {
  containerName: string;
  blobPath: string;
} | null {
  try {
    const urlObj = new URL(fileUrl.split("?")[0]);
    if (!urlObj.hostname.includes(".blob.core.windows.net")) return null;
    const parts = urlObj.pathname.split("/").filter(Boolean);
    if (parts.length < 2) return null;
    return {
      containerName: parts[0],
      blobPath: parts.slice(1).join("/"),
    };
  } catch {
    return null;
  }
}

async function downloadBlobDirectFromStorage(
  containerName: string,
  blobPath: string
): Promise<Buffer | null> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY;
  if (!acc || !key) return null;

  try {
    const connStr = `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`;
    const svc = BlobServiceClient.fromConnectionString(connStr);
    const cc = svc.getContainerClient(containerName);
    return await cc.getBlockBlobClient(blobPath).downloadToBuffer();
  } catch (e) {
    console.warn(`[edit-pptx] downloadBlobDirectFromStorage failed (${containerName}/${blobPath}):`, String((e as any)?.message ?? e));
    return null;
  }
}

async function downloadBlob(fileUrl: string, threadId?: string): Promise<Buffer> {
  const res = await fetch(fileUrl);
  if (res.ok) {
    return Buffer.from(await res.arrayBuffer());
  }

  const blobRef = parseAzureBlobUrl(fileUrl);
  if (blobRef && (res.status === 403 || res.status === 404)) {
    const directBuffer = await downloadBlobDirectFromStorage(
      blobRef.containerName,
      blobRef.blobPath
    );
    if (directBuffer) {
      console.warn(
        `[edit-pptx] recovered blob download via account key: ${blobRef.containerName}/${blobRef.blobPath}`
      );
      return directBuffer;
    }
  }

  if (
    (res.status === 403 || res.status === 404) &&
    blobRef &&
    (blobRef.containerName === "dl-link" || blobRef.containerName === "pptx")
  ) {
    const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME;
    const key = process.env.AZURE_STORAGE_ACCOUNT_KEY;
    if (acc && key) {
      const connStr = `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`;
      const svc = BlobServiceClient.fromConnectionString(connStr);
      const cc = svc.getContainerClient(blobRef.containerName);

      if (blobRef.containerName === "dl-link") {
        // dl-link は {threadId}/{filename} 構造
        const blobPathParts = blobRef.blobPath.split("/").filter(Boolean);
        const effectiveThreadId = threadId?.trim() || blobPathParts[0];
        if (effectiveThreadId) {
          for await (const blob of cc.listBlobsFlat({ prefix: `${effectiveThreadId}/` })) {
            if (blob.name.toLowerCase().endsWith(".pptx")) {
              return await cc.getBlockBlobClient(blob.name).downloadToBuffer();
            }
          }
        }
      } else {
        // pptx コンテンツは {threadId}_edited_{uniqueId}.pptx などフラット構造
        // threadId プレフィックスで前方一致検索、スラッシュなし！
        const effectiveThreadId = threadId?.trim();
        if (effectiveThreadId) {
          for await (const blob of cc.listBlobsFlat({ prefix: effectiveThreadId })) {
            if (blob.name.toLowerCase().endsWith(".pptx")) {
              return await cc.getBlockBlobClient(blob.name).downloadToBuffer();
            }
          }
        }
      }
    }
  }

  throw new Error(`Failed to download file: HTTP ${res.status}`);
}

async function uploadToBlob(buffer: Buffer, fileName: string): Promise<string> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const containerName = "pptx";

  const cred = new StorageSharedKeyCredential(acc, key);
  const svc = BlobServiceClient.fromConnectionString(
    `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
  const cc = svc.getContainerClient(containerName);
  await cc.createIfNotExists({ access: "blob" });

  const bbc = cc.getBlockBlobClient(fileName);
  await bbc.uploadData(buffer, {
    blobHTTPHeaders: {
      blobContentType:
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      blobContentDisposition: `attachment; filename="${fileName}"`,
    },
  });

  const sas = generateBlobSASQueryParameters(
    {
      containerName,
      blobName: fileName,
      expiresOn: new Date(Date.now() + 24 * 60 * 60 * 1000),
      permissions: BlobSASPermissions.parse("r"),
    },
    cred
  );
  return `${bbc.url}?${sas}`;
}

// ─────────────────────────────────────────────
// Excel サポート
// ─────────────────────────────────────────────

const EXCEL_EXTENSIONS = new Set([".xlsx", ".xls", ".xlsm"]);

function getFileExtension(url: string): string {
  try {
    const pathname = new URL(url.split("?")[0]).pathname;
    return path.extname(pathname).toLowerCase();
  } catch {
    return path.extname(url.split("?")[0]).toLowerCase();
  }
}

function isExcelFile(ext: string): boolean {
  return EXCEL_EXTENSIONS.has(ext);
}

type SheetSummary = {
  sheetName: string;
  rowCount: number;
  colCount: number;
  columns: Array<{
    letter: string;
    header: string;
    headerFillColor?: string;  // RRGGBB (existing header background color)
    headerFontColor?: string;  // RRGGBB (existing header font color)
    headerBold?: boolean;
  }>;
  sampleRows: Array<Record<string, string>>;
};

type ExcelEditPlan = {
  sheetEdits?: Array<{
    sheetName: string;
    setCells?: Array<{ address: string; value: string | number }>;
    replaceText?: Array<{ find: string; replace: string }>;
  }>;
  formatEdits?: Array<{
    sheetName: string;
    range: string;
    bold?: boolean;
    fontColor?: string;
    fillColor?: string;
  }>;
  copyRowColorEdits?: Array<{
    sheetName: string;
    targetColumn: string;   // 列記号 ("G") またはヘッダー名 ("対応")
    referenceColumn: string; // 色を参照する列記号またはヘッダー名
    startRow?: number;       // デフォルト 2（ヘッダー行をスキップ）
  }>;
  borderEdits?: Array<{
    sheetName: string;
    range: string;           // "A1:D10" など
    style?: string;          // "thin" | "medium" | "thick" | "hair" | "dashed"
    edges?: string;          // "all" | "outer" | "inner" | "top" | "bottom" | "left" | "right"
  }>;
};

/** xlsx (zip) の styles.xml と theme.xml を解析し、各シートのヘッダー行セルスタイルを返す。
 *  戻り値: sheetIndex(0始まり) → colLetter → {fillColor, fontColor, bold}
 */
async function extractXlsxHeaderStyles(
  buffer: Buffer,
  sheetCount: number
): Promise<Map<number, Record<string, { fillColor?: string; fontColor?: string; bold?: boolean }>>> {
  const result = new Map<number, Record<string, { fillColor?: string; fontColor?: string; bold?: boolean }>>();
  try {
    const zip = await JSZip.loadAsync(buffer);

    // --- テーマカラー解決 (theme index → RRGGBB) ---
    const themeColors: string[] = [];
    const themeFile = zip.files["xl/theme/theme1.xml"];
    if (themeFile) {
      const themeXml = await themeFile.async("string");
      const clrM = /<a:clrScheme[\s\S]*?>([\s\S]*?)<\/a:clrScheme>/.exec(themeXml);
      if (clrM) {
        for (const tag of ["dk1","lt1","dk2","lt2","accent1","accent2","accent3","accent4","accent5","accent6","hlink","folHlink"]) {
          const re = new RegExp(`<a:${tag}>[\\s\\S]*?(?:val="([0-9A-Fa-f]{6})"|lastClr="([0-9A-Fa-f]{6})")`);
          const m = re.exec(clrM[1]);
          themeColors.push(m ? (m[1] || m[2] || "000000").toUpperCase() : "000000");
        }
      }
    }

    // --- styles.xml 解极E---
    const stylesFile = zip.files["xl/styles.xml"];
    if (!stylesFile) return result;
    const stylesXml = await stylesFile.async("string");

    const fills: Array<string | undefined> = [];
    const fillsXml = /<fills[^>]*>([\s\S]*?)<\/fills>/.exec(stylesXml)?.[1] ?? "";
    for (const fm of Array.from(fillsXml.matchAll(/<fill>([\s\S]*?)<\/fill>/g))) {
      const body = fm[1];
      const rgbM = /<fgColor\s[^>]*rgb="[A-Fa-f0-9]{2}([A-Fa-f0-9]{6})"/.exec(body);
      const thM  = /<fgColor\s[^>]*theme="(\d+)"/.exec(body);
      if (rgbM) fills.push(rgbM[1].toUpperCase());
      else if (thM) fills.push(themeColors[parseInt(thM[1], 10)]);
      else fills.push(undefined);
    }

    const fonts: Array<{ color?: string; bold?: boolean }> = [];
    const fontsXml = /<fonts[^>]*>([\s\S]*?)<\/fonts>/.exec(stylesXml)?.[1] ?? "";
    for (const fm of Array.from(fontsXml.matchAll(/<font>([\s\S]*?)<\/font>/g))) {
      const body = fm[1];
      const bold = /<b[\s\/>]/.test(body);
      const rgbM = /<color\s[^>]*rgb="[A-Fa-f0-9]{2}([A-Fa-f0-9]{6})"/.exec(body);
      const thM  = /<color\s[^>]*theme="(\d+)"/.exec(body);
      let color: string | undefined;
      if (rgbM) color = rgbM[1].toUpperCase();
      else if (thM) color = themeColors[parseInt(thM[1], 10)];
      fonts.push({ color, bold: bold || undefined });
    }

    const xfEntries: Array<{ fillId: number; fontId: number }> = [];
    const xfsXml = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? "";
    for (const xm of Array.from(xfsXml.matchAll(/<xf\s([^>]*?)(?:\/>|>)/g))) {
      const attrs = xm[1];
      xfEntries.push({
        fillId: parseInt(/fillId="(\d+)"/.exec(attrs)?.[1] ?? "0", 10),
        fontId: parseInt(/fontId="(\d+)"/.exec(attrs)?.[1] ?? "0", 10),
      });
    }

    // --- 各シートのヘッダー行スタイルを抽出 ---
    for (let i = 0; i < sheetCount; i++) {
      const sheetFile = zip.files[`xl/worksheets/sheet${i + 1}.xml`];
      if (!sheetFile) continue;
      const sheetXml = await sheetFile.async("string");
      const firstRowM = /<row\s[^>]*r="1"[^>]*>([\s\S]*?)<\/row>/.exec(sheetXml);
      if (!firstRowM) continue;

      const sheetStyles: Record<string, { fillColor?: string; fontColor?: string; bold?: boolean }> = {};
      for (const cm of Array.from(firstRowM[1].matchAll(/<c\s([^>]*)>/g))) {
        const attrs = cm[1];
        const colLetter = /r="([A-Z]+)\d+"/.exec(attrs)?.[1];
        const sIdx = parseInt(/\bs="(\d+)"/.exec(attrs)?.[1] ?? "0", 10);
        if (!colLetter) continue;
        const xf = xfEntries[sIdx];
        if (!xf) continue;
        const fillColor = fills[xf.fillId];
        const font = fonts[xf.fontId];
        if (fillColor || font?.color || font?.bold) {
          sheetStyles[colLetter] = { fillColor, fontColor: font?.color, bold: font?.bold };
        }
      }
      result.set(i, sheetStyles);
    }
  } catch {
    // スタイル取得失敗は無視して空マップを返す
  }
  return result;
}

async function extractSheetSummaries(buffer: Buffer): Promise<SheetSummary[]> {
  const wb = XLSX.read(buffer, { type: "buffer", sheetStubs: false });
  const headerStylesMap = await extractXlsxHeaderStyles(buffer, wb.SheetNames.length);

  return wb.SheetNames.map((sheetName, sheetIdx) => {
    const ws = wb.Sheets[sheetName];
    const rows: string[][] = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: "",
      blankrows: false,
    }) as string[][];

    const headerRow = (rows[0] ?? []).map(String);
    const sheetStyles = headerStylesMap.get(sheetIdx) ?? {};
    const columns = headerRow.map((header, i) => {
      const letter = XLSX.utils.encode_col(i);
      const style = sheetStyles[letter];
      return {
        letter,
        header,
        ...(style?.fillColor ? { headerFillColor: style.fillColor } : {}),
        ...(style?.fontColor ? { headerFontColor: style.fontColor } : {}),
        ...(style?.bold ? { headerBold: true } : {}),
      };
    });
    const sampleRows = rows.slice(1, 6).map((row) => {
      const obj: Record<string, string> = {};
      columns.forEach(({ letter, header }, i) => {
        const key = header ? `${letter}(${header})` : letter;
        obj[key] = String(row[i] ?? "");
      });
      return obj;
    });
    const ref = ws["!ref"];
    const range = ref ? XLSX.utils.decode_range(ref) : null;

    return {
      sheetName,
      rowCount: range ? range.e.r + 1 : rows.length,
      colCount: range ? range.e.c + 1 : columns.length,
      columns,
      sampleRows,
    };
  });
}

async function buildExcelEditPlan(
  sheets: SheetSummary[],
  instruction: string
): Promise<ExcelEditPlan> {
  const openai = OpenAIInstance();

  const systemPrompt = `You convert a natural-language Excel editing request into a safe JSON edit plan.
Return JSON only in this shape:
{
  "sheetEdits": [
    {
      "sheetName": string,
      "setCells": [{ "address": "A1", "value": "new value" }],
      "replaceText": [{ "find": "old", "replace": "new" }]
    }
  ],
  "formatEdits": [
    {
      "sheetName": string,
      "range": "A1:D1",
      "bold": true | false,
      "fontColor": "RRGGBB",
      "fillColor": "RRGGBB"
    }
  ],
  "copyRowColorEdits": [
    {
      "sheetName": string,
      "targetColumn": "G",   // MUST be a column letter (A, B, C...) from the columns list
      "referenceColumn": "A", // MUST be a column letter (A, B, C...) from the columns list
      "startRow": 2           // first row to copy (default 2, skip header)
    }
  ],
  "borderEdits": [
    {
      "sheetName": string,
      "range": "A1:D10",     // cell range to apply borders
      "style": "thin",        // "thin" | "medium" | "thick" | "hair" | "dashed" (default "thin")
      "edges": "all"          // "all" | "outer" | "inner" | "top" | "bottom" | "left" | "right" (default "all")
    }
  ]
}

Rules:
- sheetName must match one of the provided sheet names exactly.
- The "columns" array in each sheet lists column letters AND header names. Each column may also have headerFillColor, headerFontColor, headerBold indicating its current header cell style.
- Always use the column LETTER (e.g. "D") in setCells addresses, copyRowColorEdits, and formatEdits ranges  Enever use the header name as a substitute for a column letter.
- Use setCells when the user wants to add or update cell values. The value field can be a literal (number or string) OR an Excel formula string starting with "=".
  - When the user asks to calculate a value, use an Excel formula. Examples:
    - 合計 =SUM(B2:B10)
    - 平均 =AVERAGE(B2:B10)
    - 偏差値 (standard score, population std dev, ALL Excel versions compatible — do NOT use STDEV.P or STDEVP):
      =IFERROR(ROUND(50+10*(B2-AVERAGE($B$2:$B$4))/SQRT(SUMPRODUCT(($B$2:$B$4-AVERAGE($B$2:$B$4))^2)/COUNT($B$2:$B$4)),2),"")
      Replace B with the score column letter and 4 with the last data row number (rowCount - 1 if row 1 is header).
      Output one setCells entry per data row (from row 2 to last data row), adjusting only the lone "B2" row number each time (keep $B$2:$B$4 fixed).
      IMPORTANT: Never use STDEV.P, STDEVP, or any function containing a dot. Never prefix "@" before a function name.
  - NEVER add a header name via setCells  Eset the header through the setCells value directly in row 1.
- Use replaceText ONLY when the user explicitly asks to find and replace text content. NEVER use replaceText for formatting operations.
- Use formatEdits for bold/color changes. fontColor and fillColor are 6-digit hex (no #).
- AUTO HEADER STYLING: Whenever you add a new column (i.e. setCells includes a row-1 cell for a column that does not yet exist), you MUST ALWAYS emit a formatEdits entry for that new header cell that copies the style of the existing styled headers:
  1. Find any existing column that has headerFillColor set  Euse its headerFillColor, headerFontColor, and headerBold.
  2. Apply all three (fillColor, fontColor, bold) to the new header cell range (e.g. "C1").
  3. This rule is unconditional  Eapply it even if the user did not mention design or style.
- DESIGN MATCHING (explicit): When the user additionally says "デザインを合わせて" / "同じデザイン" etc., also emit copyRowColorEdits to copy row-by-row background colors from an adjacent styled column to the new column.
- Use copyRowColorEdits when the user wants a column's background colors to match those of another column row-by-row. targetColumn and referenceColumn MUST be column letters from the "columns" list.
- Use borderEdits when the user asks to add borders (枠・罫線/border), frame cells, or make the sheet look cleaner. Infer the data range from the sheet summary. Use edges="all" for full grid, "outer" for outer frame only.
- NEVER set a cell value to an empty string unless the user explicitly asks to clear that cell.
- NEVER modify header row values (row 1) unless the user explicitly asks to change column names.
- Only emit the operations the user actually requested. Keep the JSON minimal.`;

  const userPrompt = `Instruction: ${instruction}

Sheets:
${JSON.stringify(sheets, null, 2)}

Return JSON only.`;

  const res = await openai.chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 1500,
  });

  const content = res.choices[0]?.message?.content ?? "{}";
  const parsed = JSON.parse(content) as ExcelEditPlan;
  parsed.sheetEdits ??= [];
  parsed.formatEdits ??= [];
  parsed.copyRowColorEdits ??= [];
  parsed.borderEdits ??= [];
  return parsed;
}

/** 新列に対して既存ヘッダーの完全なスタイル（fill+font+bold）と枠線を強制注入する。
 *  LLMがfillだけ生成してfontを欠いても、この関数で上書き補完する。
 */
function injectMissingHeaderStyles(
  plan: ExcelEditPlan,
  sheets: SheetSummary[]
): ExcelEditPlan {
  const sheetMap = new Map(sheets.map((s) => [s.sheetName, s]));

  for (const sheetEdit of plan.sheetEdits ?? []) {
    const sheet = sheetMap.get(sheetEdit.sheetName);
    if (!sheet) continue;

    // スタイルを持つ既存行をテンプレートとして使う
    const templateCol = sheet.columns.find((c) => c.headerFillColor);
    if (!templateCol) continue;

    const existingLetters = new Set(sheet.columns.map((c) => c.letter));

    // setCells でrow1に追加される新列を収集
    const newHeaderLetters = new Set<string>();
    for (const cell of sheetEdit.setCells ?? []) {
      const m = /^([A-Z]+)1$/.exec(cell.address);
      if (m && !existingLetters.has(m[1])) newHeaderLetters.add(m[1]);
    }

    for (const colLetter of Array.from(newHeaderLetters)) {
      const headerRange = `${colLetter}1`;

      // LLMが生成したこの列のformatEditsを削除して完全版で上書き！fillだけの不完全なものを排除する
      plan.formatEdits = (plan.formatEdits ?? []).filter(
        (fe) => !(fe.sheetName === sheetEdit.sheetName && fe.range === headerRange)
      );
      plan.formatEdits.push({
        sheetName: sheetEdit.sheetName,
        range: headerRange,
        bold: templateCol.headerBold ?? true,
        fontColor: templateCol.headerFontColor ?? "FFFFFF",
        fillColor: templateCol.headerFillColor!,
      });

      // 新列の全セル（ヘッダー+データ行）に枠線を注入
      // setCells の最大行数から列データ末尾を算出
      let maxRow = sheet.rowCount;
      for (const cell of sheetEdit.setCells ?? []) {
        const cm = /^([A-Z]+)(\d+)$/.exec(cell.address);
        if (cm && cm[1] === colLetter) maxRow = Math.max(maxRow, parseInt(cm[2], 10));
      }
      plan.borderEdits = plan.borderEdits ?? [];
      plan.borderEdits.push({
        sheetName: sheetEdit.sheetName,
        range: `${colLetter}1:${colLetter}${maxRow}`,
        style: "thin",
        edges: "all",
      });
    }
  }

  return plan;
}

async function resolveEditExcelScriptPath(): Promise<string> {
  const candidates = [
    path.join(process.cwd(), "src", "scripts", "edit_excel.py"),
    path.join(process.cwd(), "scripts", "edit_excel.py"),
    "/home/site/wwwroot/src/scripts/edit_excel.py",
    "/home/site/wwwroot/scripts/edit_excel.py",
  ];
  for (const candidate of candidates) {
    try {
      await fs.access(candidate, fsConstants.R_OK);
      return candidate;
    } catch {
      // try next
    }
  }
  throw new Error(`edit_excel.py not found. Checked: ${candidates.join(", ")}`);
}

async function uploadExcelToBlob(buffer: Buffer, fileName: string): Promise<string> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const containerName = "xlsx";

  const cred = new StorageSharedKeyCredential(acc, key);
  const svc = BlobServiceClient.fromConnectionString(
    `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
  const cc = svc.getContainerClient(containerName);
  await cc.createIfNotExists({ access: "blob" });

  const bbc = cc.getBlockBlobClient(fileName);
  await bbc.uploadData(buffer, {
    blobHTTPHeaders: {
      blobContentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      blobContentDisposition: `attachment; filename="${fileName}"`,
    },
  });

  const sas = generateBlobSASQueryParameters(
    {
      containerName,
      blobName: fileName,
      expiresOn: new Date(Date.now() + 24 * 60 * 60 * 1000),
      permissions: BlobSASPermissions.parse("r"),
    },
    cred
  );
  return `${bbc.url}?${sas}`;
}

async function runPythonEditExcel(
  inputBuffer: Buffer,
  inputExt: string,
  plan: ExcelEditPlan,
  threadId: string
) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-xlsx-"));
  const inputPath = path.join(tempDir, `input${inputExt}`);
  const outputPath = path.join(tempDir, "output.xlsx");
  const planPath = path.join(tempDir, "plan.json");
  const scriptPath = await resolveEditExcelScriptPath();

  try {
    await fs.writeFile(inputPath, inputBuffer);
    await fs.writeFile(planPath, JSON.stringify(plan), "utf8");

    const pythonBin = process.platform === "win32" ? "python" : "python3";

    if (process.platform !== "win32") {
      try {
        await execFileAsync(pythonBin, ["-c", "import openpyxl"]);
      } catch {
        throw new Error(
          "openpyxl がサーバーにインストールされていません。" +
          "startup.sh の設定を確認してください。"
        );
      }
    }

    const { stdout, stderr } = await execFileAsync(pythonBin, [
      scriptPath,
      "--input", inputPath,
      "--output", outputPath,
      "--plan", planPath,
    ]);

    if (stderr?.trim()) {
      console.warn("[edit-excel] python stderr:", stderr.trim());
    }

    const outputBuffer = await fs.readFile(outputPath);
    const pythonResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};
    const fileName = `${threadId || uniqueId()}_edited_${uniqueId()}.xlsx`;
    const downloadUrl = await uploadExcelToBlob(outputBuffer, fileName);

    return {
      downloadUrl,
      fileName,
      changedSheets: Number(pythonResult.changedSheets ?? 0),
      totalSheets: Number(pythonResult.totalSheets ?? 0),
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

// ─────────────────────────────────────────────
// PDF → Excel 変換
// ─────────────────────────────────────────────

async function resolveConvertPdfToWordScriptPath(): Promise<string> {
  const candidates = [
    path.join(process.cwd(), "src", "scripts", "pdf_to_word.py"),
    path.join(process.cwd(), "scripts", "pdf_to_word.py"),
    "/home/site/wwwroot/src/scripts/pdf_to_word.py",
    "/home/site/wwwroot/scripts/pdf_to_word.py",
  ];
  for (const candidate of candidates) {
    try {
      await fs.access(candidate, fsConstants.R_OK);
      return candidate;
    } catch {
      // try next
    }
  }
  throw new Error(`pdf_to_word.py not found. Checked: ${candidates.join(", ")}`);
}

async function runPythonPdfToWord(inputBuffer: Buffer, threadId: string, mode: "layout" | "editable" = "layout") {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-pdf2docx-"));
  const inputPath = path.join(tempDir, "input.pdf");
  const outputPath = path.join(tempDir, "output.docx");
  const scriptPath = await resolveConvertPdfToWordScriptPath();

  const pyEnv = process.platform !== "win32"
    ? {
        ...process.env,
        PYTHONPATH: `/home/site/python-packages${process.env.PYTHONPATH ? `:${process.env.PYTHONPATH}` : ""}`,
        LD_LIBRARY_PATH: `/home/site/python-packages${process.env.LD_LIBRARY_PATH ? `:${process.env.LD_LIBRARY_PATH}` : ""}`,
      }
    : process.env;

  try {
    await fs.writeFile(inputPath, inputBuffer);

    const pythonBin = process.platform === "win32" ? "python" : "python3";
    const { stdout, stderr } = await execFileAsync(pythonBin, [
      scriptPath,
      "--input", inputPath,
      "--output", outputPath,
      "--mode", mode,
    ], { env: pyEnv });

    if (stderr?.trim()) {
      console.warn("[pdf-to-word] python stderr:", stderr.trim());
    }

    const pythonResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};

    if (pythonResult.engine === "none") {
      return { engine: "none" as const };
    }

    const outputBuffer = await fs.readFile(outputPath);
    const fileName = `${threadId || uniqueId()}_converted_${uniqueId()}.docx`;
    const downloadUrl = await uploadWordToBlob(outputBuffer, fileName);

    return {
      downloadUrl,
      fileName,
      paragraphs: Number(pythonResult.paragraphs ?? 0),
      tables: Number(pythonResult.tables ?? 0),
      engine: String(pythonResult.engine ?? ""),
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

async function resolveConvertPdfScriptPath(): Promise<string> {
  const candidates = [
    path.join(process.cwd(), "src", "scripts", "pdf_to_excel.py"),
    path.join(process.cwd(), "scripts", "pdf_to_excel.py"),
    "/home/site/wwwroot/src/scripts/pdf_to_excel.py",
    "/home/site/wwwroot/scripts/pdf_to_excel.py",
  ];
  for (const candidate of candidates) {
    try {
      await fs.access(candidate, fsConstants.R_OK);
      return candidate;
    } catch {
      // try next
    }
  }
  throw new Error(`pdf_to_excel.py not found. Checked: ${candidates.join(", ")}`);
}

async function correctExcelAccountNames(excelBuffer: Buffer): Promise<Buffer> {
  const openai = OpenAIInstance();
  const wb = XLSX.read(excelBuffer, { type: "buffer" });

  const isAccountName = (val: string) =>
    val.trim().length > 0 && !/^[\d,，△▲\-()（）\s]+$/.test(val.trim());

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const data = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" }) as string[][];
    if (data.length === 0) continue;

    // 科目名が入る列を自動検出（各列の isAccountName 候補が 3件以上ある列をすべて処理）
    const maxCols = Math.max(...data.map(row => row.length));
    const colIndices: number[] = [];
    for (let c = 0; c < maxCols; c++) {
      const count = data.filter(row => isAccountName(String(row[c] ?? ""))).length;
      if (count >= 3) colIndices.push(c);
    }

    for (const colIdx of colIndices) {
      // 科目名候補のみに絞る（数値・金額・空セルは対象外）
      const candidates: { rowIdx: number; name: string }[] = [];
      for (let i = 0; i < data.length; i++) {
        const val = String(data[i]?.[colIdx] ?? "").trim();
        if (isAccountName(val)) candidates.push({ rowIdx: i, name: val });
      }
      if (candidates.length === 0) continue;

      console.log(`[pdf-to-excel] correctExcelAccountNames: sheet="${sheetName}" col=${colIdx} candidates=${candidates.length}`);
      console.log(`[pdf-to-excel] sample input:`, candidates.slice(0, 10).map(r => r.name));

      // バッチサイズ 80 件ずつ処理
      const BATCH = 80;
      for (let batchStart = 0; batchStart < candidates.length; batchStart += BATCH) {
        const batch = candidates.slice(batchStart, batchStart + BATCH);

        try {
          const res = await openai.chat.completions.create({
            model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
            messages: [{
              role: "user",
              content:
                "以下は日本の財務諸表（貸借対照表・損益計算書・販売費及び一般管理費・製造原価報告書）の列データです。\n" +
                "TKC会計ソフトのPDFは文字を1文字ずつ配置するため、OCR後に文字化けや断片化が発生します。\n" +
                "例: 「Ⅱ 上 原壳価」→「Ⅱ 売上原価」、「壳」→「売」、「資 本金」→「資本金」\n\n" +
                "ルール:\n" +
                "1. 入力と同じ順序・同じ件数で返してください\n" +
                "2. 科目名の文字化けを正しい勘定科目名に補正してください\n" +
                "3. 補正不要な行はそのまま返してください\n" +
                "4. ローマ数字（Ⅰ Ⅱ Ⅲ）や括弧も含めてそのまま返してください\n\n" +
                `以下の入力リスト(${batch.length}件)を補正し、JSON形式 {"corrected": [...]} で返してください:\n` +
                JSON.stringify(batch.map(r => r.name)),
            }],
            response_format: { type: "json_object" },
            max_completion_tokens: 4000,
          });

          const content = res.choices[0]?.message?.content ?? "{}";
          console.log(`[pdf-to-excel] LLM response (batch ${batchStart}):`, content.slice(0, 300));

          const parsed = JSON.parse(content);
          const corrected: unknown[] = parsed.corrected ?? Object.values(parsed)[0];
          if (!Array.isArray(corrected) || corrected.length !== batch.length) {
            console.warn(`[pdf-to-excel] length mismatch: expected=${batch.length} got=${Array.isArray(corrected) ? corrected.length : "non-array"}`);
            continue;
          }

          for (let i = 0; i < batch.length; i++) {
            const newVal = String(corrected[i] ?? "").trim();
            const oldVal = batch[i].name;
            // 空文字返却・変更なし・明らかな誤補正はスキップ（原文維持）
            if (!newVal || newVal === oldVal) continue;
            const addr = XLSX.utils.encode_cell({ r: batch[i].rowIdx, c: colIdx });
            ws[addr] = { t: "s", v: newVal };
            console.log(`[pdf-to-excel] corrected: "${oldVal}" → "${newVal}"`);
          }
        } catch (e) {
          console.warn(`[pdf-to-excel] LLM correction failed sheet=${sheetName} col=${colIdx} batch=${batchStart}:`, String((e as any)?.message ?? e));
        }
      }
    }
  }

  return Buffer.from(XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
}

async function runPythonPdfToExcel(inputBuffer: Buffer, threadId: string, fileUrl?: string) {
  const inputExt = fileUrl ? (getFileExtension(fileUrl) || ".pdf") : ".pdf";
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-pdf2xl-"));
  const inputPath = path.join(tempDir, `input${inputExt}`);
  const outputPath = path.join(tempDir, "output.xlsx");
  const scriptPath = await resolveConvertPdfScriptPath();

  // PYTHONPATH・LD_LIBRARY_PATH を明示的に設定！startup.sh が動いていない環境でも動作させるため
  const pyEnv = process.platform !== "win32"
    ? {
        ...process.env,
        PYTHONPATH: `/home/site/python-packages${process.env.PYTHONPATH ? `:${process.env.PYTHONPATH}` : ""}`,
        LD_LIBRARY_PATH: `/home/site/python-packages${process.env.LD_LIBRARY_PATH ? `:${process.env.LD_LIBRARY_PATH}` : ""}`,
      }
    : process.env;

  try {
    await fs.writeFile(inputPath, inputBuffer);

    const pythonBin = process.platform === "win32" ? "python" : "python3";

    const { stdout, stderr } = await execFileAsync(pythonBin, [
      scriptPath,
      "--input", inputPath,
      "--output", outputPath,
    ], { env: pyEnv });

    if (stderr?.trim()) {
      console.warn("[pdf-to-excel] python stderr:", stderr.trim());
    }

    const rawBuffer = await fs.readFile(outputPath);
    const pythonResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};
    let outputBuffer: Buffer = Buffer.from(rawBuffer);
    try {
      outputBuffer = await correctExcelAccountNames(rawBuffer);
    } catch (e) {
      console.warn("[pdf-to-excel] LLM account name correction skipped (fallback to raw):", String((e as any)?.message ?? e));
    }
    const fileName = `${threadId || uniqueId()}_converted_${uniqueId()}.xlsx`;
    const downloadUrl = await uploadExcelToBlob(outputBuffer, fileName);

    return {
      downloadUrl,
      fileName,
      sheets: Number(pythonResult.sheets ?? 0),
      tables: Number(pythonResult.tables ?? 0),
      pages: Number(pythonResult.pages ?? 0),
      engine: String(pythonResult.engine ?? ""),
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

// ─────────────────────────────────────────────
// Word サポート
// ─────────────────────────────────────────────

const WORD_EXTENSIONS = new Set([".docx"]);

function isWordFile(ext: string): boolean {
  return WORD_EXTENSIONS.has(ext);
}

type DocParagraph = { style: string; text: string };

type WordDocSummary = {
  paragraphs: DocParagraph[];
  totalParagraphs: number;
};

type WordEditPlan = {
  replaceText?: Array<{ find: string; replace: string }>;
  formatRuns?: Array<{
    matchText?: string;   // omit to apply to ALL paragraphs
    bold?: boolean;
    italic?: boolean;
    fontSize?: number;
    fontColor?: string;
    fontFace?: string;
  }>;
  addParagraphs?: Array<{
    text: string;
    style?: string;       // "Normal" | "Heading1" | "Heading2" | "List Bullet"
    bold?: boolean;
    italic?: boolean;
    fontSize?: number;
    fontColor?: string;
    fontFace?: string;
  }>;
};

async function extractDocSummary(buffer: Buffer): Promise<WordDocSummary> {
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.files["word/document.xml"].async("string");

  const paraRe = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
  const paragraphs: DocParagraph[] = [];
  let pm: RegExpExecArray | null;
  while ((pm = paraRe.exec(docXml)) !== null) {
    const paraXml = pm[0];
    const styleMatch = /<w:pStyle\s+w:val="([^"]+)"/.exec(paraXml);
    const style = styleMatch ? styleMatch[1] : "Normal";
    const textRe = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
    let text = "";
    let tm: RegExpExecArray | null;
    while ((tm = textRe.exec(paraXml)) !== null) {
      text += decodeXmlEntities(tm[1]);
    }
    if (text.trim()) paragraphs.push({ style, text: text.trim() });
  }

  return { paragraphs: paragraphs.slice(0, 50), totalParagraphs: paragraphs.length };
}

async function buildWordEditPlan(
  summary: WordDocSummary,
  instruction: string
): Promise<WordEditPlan> {
  const openai = OpenAIInstance();

  const systemPrompt = `You convert a natural-language Word document editing request into a safe JSON edit plan.
Return JSON only in this shape:
{
  "replaceText": [{ "find": "old text", "replace": "new text" }],
  "formatRuns": [
    {
      "matchText": "paragraph text to find (omit to apply to ALL paragraphs)",
      "bold": true,
      "italic": false,
      "fontSize": 14,
      "fontColor": "RRGGBB",
      "fontFace": "Yu Gothic"
    }
  ],
  "addParagraphs": [
    {
      "text": "paragraph text to append",
      "style": "Normal",
      "bold": false,
      "fontSize": 12,
      "fontFace": "Meiryo"
    }
  ]
}

Rules:
- replaceText: use when the user wants to change specific wording.
- formatRuns: use when the user wants to apply bold/italic/font size/color/font face.
  - matchText: substring found in the target paragraph. OMIT matchText entirely (do not include the key) when the user wants to format ALL paragraphs or the whole document.
  - fontSize: points. If the user says "4倍" or "4x", multiply the likely current size (11pt default) by 4 → 44. "2倍" → 22. "大きく" → 18.
  - fontColor: 6-digit hex without #.
  - fontFace: font name string. "ゴシック" → "Yu Gothic", "明朝" → "Yu Mincho", "メイリオ" → "Meiryo". Use the exact font name as a string.
- addParagraphs: use when the user wants to INSERT or APPEND new text to the document.
  - style: "Normal" | "Heading1" | "Heading2" | "List Bullet" (default "Normal").
  - fontFace: font name string. Apply the same mapping as formatRuns. "ゴシック" → "Yu Gothic", "明朝" → "Yu Mincho", "メイリオ" → "Meiryo".
  - Paragraphs are appended at the end of the document.
- Only emit operations the user actually requested. Keep JSON minimal.`;

  const userPrompt = `Instruction: ${instruction}

Document summary:
${JSON.stringify(summary, null, 2)}

Return JSON only.`;

  const res = await openai.chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 1000,
  });

  const content = res.choices[0]?.message?.content ?? "{}";
  const parsed = JSON.parse(content) as WordEditPlan;
  parsed.replaceText ??= [];
  parsed.formatRuns ??= [];
  parsed.addParagraphs ??= [];
  return parsed;
}

async function resolveEditWordScriptPath(): Promise<string> {
  const candidates = [
    path.join(process.cwd(), "src", "scripts", "edit_word.py"),
    path.join(process.cwd(), "scripts", "edit_word.py"),
    "/home/site/wwwroot/src/scripts/edit_word.py",
    "/home/site/wwwroot/scripts/edit_word.py",
  ];
  for (const candidate of candidates) {
    try {
      await fs.access(candidate, fsConstants.R_OK);
      return candidate;
    } catch {
      // try next
    }
  }
  throw new Error(`edit_word.py not found. Checked: ${candidates.join(", ")}`);
}

async function uploadWordToBlob(buffer: Buffer, fileName: string): Promise<string> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const containerName = "docx";

  const cred = new StorageSharedKeyCredential(acc, key);
  const svc = BlobServiceClient.fromConnectionString(
    `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
  const cc = svc.getContainerClient(containerName);
  await cc.createIfNotExists({ access: "blob" });

  const bbc = cc.getBlockBlobClient(fileName);
  await bbc.uploadData(buffer, {
    blobHTTPHeaders: {
      blobContentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      blobContentDisposition: `attachment; filename="${fileName}"`,
    },
  });

  const sas = generateBlobSASQueryParameters(
    {
      containerName,
      blobName: fileName,
      expiresOn: new Date(Date.now() + 24 * 60 * 60 * 1000),
      permissions: BlobSASPermissions.parse("r"),
    },
    cred
  );
  return `${bbc.url}?${sas}`;
}

async function runPythonEditWord(
  inputBuffer: Buffer,
  plan: WordEditPlan,
  threadId: string
) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-docx-"));
  const inputPath = path.join(tempDir, "input.docx");
  const outputPath = path.join(tempDir, "output.docx");
  const planPath = path.join(tempDir, "plan.json");
  const scriptPath = await resolveEditWordScriptPath();

  // Explicitly set Python paths for App Service (startup.sh installs to /home/site/python-packages)
  const pyEnv = process.platform !== "win32"
    ? {
        ...process.env,
        PYTHONPATH: `/home/site/python-packages${process.env.PYTHONPATH ? `:${process.env.PYTHONPATH}` : ""}`,
        LD_LIBRARY_PATH: `/home/site/python-packages${process.env.LD_LIBRARY_PATH ? `:${process.env.LD_LIBRARY_PATH}` : ""}`,
      }
    : process.env;

  try {
    await fs.writeFile(inputPath, inputBuffer);
    await fs.writeFile(planPath, JSON.stringify(plan), "utf8");

    const pythonBin = process.platform === "win32" ? "python" : "python3";

    if (process.platform !== "win32") {
      try {
        await execFileAsync(pythonBin, ["-c", "import docx"], { env: pyEnv });
      } catch {
        throw new Error(
          "python-docx がサーバーにインストールされていません。" +
          "startup.sh の設定を確認してください。"
        );
      }
    }

    const { stdout, stderr } = await execFileAsync(pythonBin, [
      scriptPath,
      "--input", inputPath,
      "--output", outputPath,
      "--plan", planPath,
    ], { env: pyEnv });

    if (stderr?.trim()) {
      console.warn("[edit-word] python stderr:", stderr.trim());
    }

    const outputBuffer = await fs.readFile(outputPath);
    const pythonResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};
    const fileName = `${threadId || uniqueId()}_edited_${uniqueId()}.docx`;
    const downloadUrl = await uploadWordToBlob(outputBuffer, fileName);

    return {
      downloadUrl,
      fileName,
      changedParagraphs: Number(pythonResult.changedParagraphs ?? 0),
      totalParagraphs: Number(pythonResult.totalParagraphs ?? 0),
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

// ─────────────────────────────────────────────
// DALL-E (PPTX 用)
// ─────────────────────────────────────────────

async function generateDalleImage(prompt: string): Promise<Buffer | null> {
  try {
    const openai = OpenAIDALLEInstance();
    // gpt-image-1 は response_format 非対応のため省略。
    // レスポンスは b64_json または url のどちらかで返る。
    const response = await openai.images.generate({
      prompt,
      n: 1,
      size: "1024x1024",
    } as Parameters<typeof openai.images.generate>[0]);

    const item = response.data?.[0] as any;

    // b64_json で返ってきた場合
    if (item?.b64_json) {
      return Buffer.from(item.b64_json, "base64");
    }

    // url で返ってきた場合
    if (item?.url) {
      const imageRes = await fetch(item.url);
      if (!imageRes.ok) {
        console.warn("[edit-pptx] Failed to fetch generated image:", imageRes.status);
        return null;
      }
      return Buffer.from(await imageRes.arrayBuffer());
    }

    console.warn("[edit-pptx] DALL-E returned no image data. raw:", JSON.stringify(response));
    return null;
  } catch (e: any) {
    console.warn("[edit-pptx] DALL-E image generation failed:", String(e?.message ?? e));
    return null;
  }
}

async function resolveEditPptxScriptPath(): Promise<string> {
  const candidates = [
    path.join(process.cwd(), "src", "scripts", "edit_pptx.py"),
    path.join(process.cwd(), "scripts", "edit_pptx.py"),
    "/home/site/wwwroot/src/scripts/edit_pptx.py",
    "/home/site/wwwroot/scripts/edit_pptx.py",
  ];

  for (const candidate of candidates) {
    try {
      await fs.access(candidate, fsConstants.R_OK);
      return candidate;
    } catch {
      // try next candidate
    }
  }

  throw new Error(`edit_pptx.py not found. Checked: ${candidates.join(", ")}`);
}

async function runPythonEdit(inputBuffer: Buffer, plan: EditPlan, threadId: string) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-pptx-"));
  const inputPath = path.join(tempDir, "input.pptx");
  const outputPath = path.join(tempDir, "output.pptx");
  const planPath = path.join(tempDir, "plan.json");
  const scriptPath = await resolveEditPptxScriptPath();

  try {
    await fs.writeFile(inputPath, inputBuffer);

    // imageInserts がある場合は DALL-E で画像生成し、一時ファイルパスをプランに追記
    const requestedImages = plan.imageInserts?.length ?? 0;
    let generatedImages = 0;
    if (requestedImages > 0) {
      for (let i = 0; i < plan.imageInserts!.length; i++) {
        const insert = plan.imageInserts![i];
        console.log(`[edit-pptx] generating image ${i}: "${insert.imagePrompt}"`);
        const imageBuffer = await generateDalleImage(insert.imagePrompt);
        if (imageBuffer) {
          const imagePath = path.join(tempDir, `image_${i}.png`);
          await fs.writeFile(imagePath, imageBuffer);
          insert.imagePath = imagePath;
          generatedImages++;
          console.log(`[edit-pptx] image ${i} generated`);
        } else {
          console.warn(`[edit-pptx] image ${i} skipped: DALL-E failed`);
        }
      }
    }

    await fs.writeFile(planPath, JSON.stringify(plan), "utf8");

    // Azure App Service (Linux) は python3、Windows ローカルは python
    const pythonBin = process.platform === "win32" ? "python" : "python3";

    // python-pptx / lxml は startup.sh でインストール済みであることを前提とする
    // runtime install は行わず、未インストールの場合は明示エラーを返す
    if (process.platform !== "win32") {
      try {
        await execFileAsync(pythonBin, ["-c", "import pptx, lxml"]);
      } catch {
        throw new Error(
          "python-pptx または lxml がサーバーにインストールされていません。" +
          "startup.sh の設定を確認してください。"
        );
      }
    }

    const { stdout, stderr } = await execFileAsync(pythonBin, [
      scriptPath,
      "--input",
      inputPath,
      "--output",
      outputPath,
      "--plan",
      planPath,
    ]);

    if (stderr?.trim()) {
      console.warn("[edit-pptx] python stderr:", stderr.trim());
    }

    const outputBuffer = await fs.readFile(outputPath);
    const pythonResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};
    const fileName = `${threadId || uniqueId()}_edited_${uniqueId()}.pptx`;
    const downloadUrl = await uploadToBlob(outputBuffer, fileName);

    const insertedImages = Number(pythonResult.insertedImages ?? 0);
    const imageWarning =
      requestedImages > 0 && insertedImages < requestedImages
        ? `画像挿入: ${requestedImages}件要汁E/ ${insertedImages}件成功`
        : undefined;

    return {
      downloadUrl,
      fileName,
      changedSlides: Number(pythonResult.changedSlides ?? 0),
      totalSlides: Number(pythonResult.totalSlides ?? 0),
      requestedImages,
      insertedImages,
      ...(imageWarning ? { imageWarning } : {}),
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { fileUrl, instruction, threadId, action, mode } = body as {
      fileUrl: string;
      instruction: string;
      threadId: string;
      action?: string;
      mode?: string;
    };

    if (!fileUrl?.trim() || (!instruction?.trim() && action !== "pdf_to_excel" && action !== "pdf_to_word")) {
      return NextResponse.json(
        { ok: false, error: "fileUrl and instruction are required" },
        { status: 400 }
      );
    }

    const ext = getFileExtension(fileUrl);
    console.log(`[edit-pptx] fileUrl =`, fileUrl.substring(0, 80));
    console.log(`[edit-pptx] ext =`, ext, "action =", action ?? "(none)", "instruction =", instruction.substring(0, 120));

    // PDF → Excel 変換
    if (action === "pdf_to_excel") {
      const pdfBuffer = await downloadBlob(fileUrl, threadId);
      const result = await runPythonPdfToExcel(pdfBuffer, threadId, fileUrl);
      console.log("[pdf-to-excel] result:", JSON.stringify(result));
      return NextResponse.json({ ok: true, ...result });
    }

    // PDF → Word 変換
    if (action === "pdf_to_word") {
      const pdfBuffer = await downloadBlob(fileUrl, threadId);
      const wordMode = mode === "editable" ? "editable" : "layout";
      const result = await runPythonPdfToWord(pdfBuffer, threadId, wordMode);
      console.log("[pdf-to-word] result:", JSON.stringify(result));
      return NextResponse.json({ ok: true, ...result });
    }

    // Excel ファイル (.xlsx / .xls / .xlsm) の場合は Excel 専用フローへ
    if (isExcelFile(ext)) {
      const excelBuffer = await downloadBlob(fileUrl, threadId);
      const sheets = await extractSheetSummaries(excelBuffer);
      console.log("[edit-excel] sheets:", JSON.stringify(sheets.map(s => ({ name: s.sheetName, columns: s.columns }))));
      const rawPlan = await buildExcelEditPlan(sheets, instruction);
      const plan = injectMissingHeaderStyles(rawPlan, sheets);

      console.log("[edit-excel] plan:", JSON.stringify(plan));

      const result = await runPythonEditExcel(excelBuffer, ext, plan, threadId);
      return NextResponse.json({ ok: true, ...result });
    }

    // Word ファイル (.docx) の場合は Word 専用フローへ
    if (isWordFile(ext)) {
      const wordBuffer = await downloadBlob(fileUrl, threadId);
      const summary = await extractDocSummary(wordBuffer);
      const plan = await buildWordEditPlan(summary, instruction);

      console.log("[edit-word] plan:", JSON.stringify(plan));

      const result = await runPythonEditWord(wordBuffer, plan, threadId);
      return NextResponse.json({ ok: true, ...result });
    }

    // PPTX フロー（既存！）
    const pptxBuffer = await downloadBlob(fileUrl, threadId);
    const slides = await extractSlideSummaries(pptxBuffer);
    const plan = await buildEditPlan(slides, instruction);

    console.log("[edit-pptx] plan:", JSON.stringify(plan));

    const result = await runPythonEdit(pptxBuffer, plan, threadId);

    return NextResponse.json({
      ok: true,
      ...result,
    });
  } catch (e: any) {
    console.error("[edit-pptx] error:", e);
    return NextResponse.json(
      { ok: false, error: String(e?.message ?? e) },
      { status: 500 }
    );
  }
}
