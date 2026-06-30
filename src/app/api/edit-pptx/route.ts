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


type ImageInsert = {
  slideIndex: number;  // -1 = 全スライドに適用
  imagePrompt?: string; // DALL-E 生成用プロンプト（imageUrl がない場合）
  imageUrl?: string;   // 既存画像URL（ロゴ等）— 指定時は DALL-E をスキップ
  position?: "top-right" | "top-left" | "bottom-right" | "bottom-left" | "center";
  nearText?: string;   // このテキストを含む Shape の隣に配置（指定時は position より優先）
  anchorSide?: "left" | "right" | "above" | "below"; // nearText Shape のどちら側か（default: "right"）
  widthPct?: number;
  imagePath?: string; // set after download/generation, before Python call
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
      replace?: string;
      appendToRun?: string;
    }>;
    addBullets?: Array<{
      afterText: string;
      texts: string[];
    }>;
    copyShapeBlock?: {
      headingShapeName: string;
      descShapeName: string;
      headingText: string;
      descText: string;
      groupShapeNames?: string[];
    };
    convertToCards?: {
      cards: Array<{
        heading: string;
        body: string;
        iconKey?: string;
      }>;
    };
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

// 各 <a:t> テキストを個別に返す（Python run.text.replace の対象単位と一致）
function extractShapeRunTexts(shapeXml: string): string[] {
  const runs: string[] = [];
  const runRe = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
  let rm: RegExpExecArray | null;
  while ((rm = runRe.exec(shapeXml)) !== null) {
    const t = decodeXmlEntities(rm[1]).trim();
    if (t.length >= 2) runs.push(t); // 1文字以下は記号・装飾の可能性が高いのでスキップ
  }
  return runs;
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


// タイトルshapeを正確に検出して構造化スライドデータを返す（runs = Python置換単位のテキスト）
function extractShapeName(shapeXml: string): string {
  return shapeXml.match(/<p:cNvPr[^>]*\bname="([^"]+)"/)?.[1] ?? "";
}

async function extractSlidesStructured(
  pptxBuffer: Buffer
): Promise<Array<{ slideIndex: number; title: string; bullets: string[]; runs: string[]; shapes: Array<{ name: string; texts: string[] }> }>> {
  const zip = await JSZip.loadAsync(pptxBuffer);
  const slideEntries = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)/)?.[1] ?? "0", 10);
      const numB = parseInt(b.match(/slide(\d+)/)?.[1] ?? "0", 10);
      return numA - numB;
    });

  const result: Array<{ slideIndex: number; title: string; bullets: string[]; runs: string[]; shapes: Array<{ name: string; texts: string[] }> }> = [];
  for (let si = 0; si < slideEntries.length; si++) {
    const xml = await zip.files[slideEntries[si]].async("string");
    const shapeXmls = splitIntoShapes(xml);

    let title = "";
    const bullets: string[] = [];
    const runs: string[] = []; // 非タイトルshapeのrun単位テキスト（Python置換検証用）
    const shapes: Array<{ name: string; texts: string[] }> = []; // shape単位の構造（heading+description検出用）

    for (const shapeXml of shapeXmls) {
      // <p:ph type="title"> or <p:ph type="ctrTitle"> marks the title placeholder
      const isTitle = /<p:ph[^/]*type="(?:title|ctrTitle)"/.test(shapeXml);
      const texts = extractShapeTextByParagraph(shapeXml)
        .map((t) => t.trim())
        .filter(Boolean);
      if (isTitle && !title) {
        title = texts.join(" ").slice(0, 80);
      } else {
        bullets.push(...texts.slice(0, 8));
        runs.push(...extractShapeRunTexts(shapeXml));
        const name = extractShapeName(shapeXml);
        if (texts.length > 0) {
          shapes.push({ name, texts: texts.slice(0, 6) });
        }
      }
    }

    // Fallback: no title placeholder found — promote first bullet
    if (!title && bullets.length > 0) {
      title = bullets.shift() ?? "";
    }

    result.push({
      slideIndex: si,
      title,
      bullets: bullets.slice(0, 10),
      runs: runs.slice(0, 80).map((r) => r.slice(0, 120)),
      shapes: shapes.slice(0, 30),
    });
  }
  return result;
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

// URL付きの画像挿入指示をLLM不要で直接EditPlanに変換する
function tryBuildDirectPlan(instruction: string, _slides: unknown[]): EditPlan | null {
  const urlMatch = instruction.match(/https?:\/\/[^\s)>\]　「」）]+/);
  if (!urlMatch) return null;
  const matchedUrl = urlMatch[0];

  // PPTX/PDF/DOCX/XLSX など非画像URLは画像挿入対象外
  const isImageUrl = /\.(?:png|jpg|jpeg|webp|gif|bmp)(\?|$)/i.test(matchedUrl);
  if (!isImageUrl) return null;
  const imageUrl = matchedUrl;

  // スライドインデックス決定（デフォルト: 表紙=0）
  let slideIndex = 0;
  const allSlidesRe = /全スライド|全ページ|all slides|すべてのスライド|全ての/;
  const slideNumRe = /スライド[　 ]*(\d+)|(\d+)[　 ]*枚目|(\d+)[　 ]*ページ/;
  if (allSlidesRe.test(instruction)) {
    slideIndex = -1;
  } else {
    const numMatch = instruction.match(slideNumRe);
    if (numMatch) {
      slideIndex = parseInt(numMatch[1] ?? numMatch[2] ?? numMatch[3], 10) - 1;
    }
    // 「表紙」指定があっても index=0 のまま（デフォルトと同じ）
  }

  // 配置・サイズ
  const isLogo = /ロゴ|logo/i.test(instruction);
  const widthPct = isLogo ? 12 : 15;
  let position: "top-right" | "top-left" | "bottom-right" | "bottom-left" | "center" = "top-right";
  if (/左上|top[\s-]?left/i.test(instruction)) position = "top-left";
  else if (/左下|bottom[\s-]?left/i.test(instruction)) position = "bottom-left";
  else if (/右下|bottom[\s-]?right/i.test(instruction)) position = "bottom-right";
  else if (/右上|top[\s-]?right/i.test(instruction)) position = "top-right";
  else if (/中央|center/i.test(instruction)) position = "center";

  return {
    deckEdits: { accentColor: null, fontFace: null, preserveTextColors: true },
    slideEdits: [],
    imageInserts: [{ slideIndex, imageUrl, position, widthPct }],
  };
}

async function buildEditPlan(
  slides: Array<{ slideIndex: number; title: string; bullets: string[]; runs: string[] }>,
  instruction: string
): Promise<EditPlan> {
  // 決定的に解釈できる画像URL挿入はLLMをスキップして直接プランを生成
  const direct = tryBuildDirectPlan(instruction, slides);
  if (direct && (direct.imageInserts?.length ?? 0) > 0 &&
      (direct.slideEdits?.length ?? 0) === 0 &&
      !direct.deckEdits?.accentColor) {
    // URL挿入のみで色・テキスト編集がない場合は直接返す
    const hasOtherIntent = /色|カラー|color|フォント|font|テキスト|text|文字|変更|replace|recolor/i.test(instruction);
    if (!hasOtherIntent) {
      console.log("[buildEditPlan] using direct plan (URL image insert, no LLM needed)");
      return direct;
    }
  }

  const openai = OpenAIInstance();
  const pptModel = process.env.AZURE_OPENAI_PPT_DEPLOYMENT_NAME ?? process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!;

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
      "imageUrl": string,
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
- CRITICAL: "find" MUST be a short verbatim substring (≤ 80 chars) that appears in runs[] of the target slide. runs[] contains individual PowerPoint text-run units (shorter than full bullets). NEVER use a full bullet/paragraph sentence as "find" — it spans multiple runs and will not match. Pick a distinctive short phrase from runs[].
- slideIndex is zero-based.
- If the user asks to add an icon, illustration, image, or mark to a slide, populate imageInserts[].
  - imagePrompt: concise English DALL-E prompt describing the image (e.g. "robot icon, flat design, simple, white background"). Use this when NO existing URL is provided.
  - imageUrl: if the user provides an existing image URL (logo, photo, etc.), set imageUrl to that URL and omit imagePrompt. Do NOT generate a DALL-E prompt when a URL is given.
  - slideIndex: use -1 to apply the image to ALL slides (useful for logos/watermarks). Use a specific index for a single slide.
  - nearText: if the user says "next to X", "beside X", "横に", "〇〇の横" etc., set nearText to the shortest unique text string found in the slide (e.g. "ボット" not the full sentence). The image will be placed adjacent to the shape containing that text.
  - anchorSide: which side of the nearText shape to place the image. Use "right" for "横に/右に", "left" for "左に", "above" for "上に", "below" for "下に". Default "right".
  - position: fallback position if nearText shape is not found. Use "top-right" for decorative icons and logos. Never use "top-left" as it may overlap slide content.
  - widthPct: image width as percentage of slide width. Use 6-8 for small inline icons, 10-15 for logos, 30-50 for large illustrations. Default 12 for logos (imageUrl), 8 when nearText is set, 13 otherwise.
- Only emit imageInserts when the user explicitly requests an image or visual element.
- Keep the JSON minimal. Use null or [] when not needed.`;

  const userPrompt = `Instruction:
${instruction}

Slides:
${JSON.stringify(slides, null, 2)}

Return JSON only.`;

  const res = await openai.chat.completions.create({
    model: pptModel,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 4000,
  });

  const finishReason = res.choices[0]?.finish_reason;
  if (finishReason && finishReason !== "stop") {
    console.warn(`[buildEditPlan] finish_reason=${finishReason}`);
  }
  const content = res.choices[0]?.message?.content || "{}";

  let parsed: EditPlan;
  try {
    parsed = JSON.parse(content) as EditPlan;
  } catch (e) {
    console.warn(
      `[buildEditPlan] JSON.parse failed (finish_reason=${finishReason}), content[:300]=${content.slice(0, 300)}`
    );
    // 直接プランで救済できるか試みる
    const fallback = tryBuildDirectPlan(instruction, slides);
    if (fallback) {
      console.log("[buildEditPlan] using direct plan as JSON-parse fallback");
      return fallback;
    }
    // 最終フォールバック: 空プラン（500にしない）
    parsed = { deckEdits: {}, slideEdits: [], imageInserts: [] };
  }

  parsed.deckEdits ??= {};
  parsed.slideEdits ??= [];

  const directAccent = parseDirectAccentColor(instruction);
  // imageInserts がある場合は画像プロンプト内の色指定（「青系アイコン」等）を
  // デッキ全体の色変更と誤認しないよう parseDirectAccentColor を適用しない。
  // LLM が明示的に accentColor を返した場合のみデッキ色を変更する。
  const hasImageInserts = (parsed.imageInserts?.length ?? 0) > 0;
  if (!normalizeHexColor(parsed.deckEdits.accentColor) && directAccent && !hasImageInserts) {
    parsed.deckEdits.accentColor = directAccent;
  } else {
    parsed.deckEdits.accentColor = normalizeHexColor(parsed.deckEdits.accentColor) ?? null;
  }
  parsed.deckEdits.preserveTextColors = parsed.deckEdits.preserveTextColors !== false;

  // instruction に画像URLが含まれている場合、LLMが誤って imagePrompt を返しても
  // imageUrl に差し替えて DALL-E フォールバックを防ぐ
  if (parsed.imageInserts?.length) {
    const urlInInstruction = instruction.match(/https?:\/\/[^\s)>\]　「」）]+/)?.[0];
    if (urlInInstruction) {
      for (const ins of parsed.imageInserts) {
        if (!ins.imageUrl && ins.imagePrompt) {
          ins.imageUrl = urlInInstruction;
          delete ins.imagePrompt;
        }
      }
    }
  }

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
        // pptx コンテナは {threadId}_edited_{uniqueId}.pptx などフラット構造
        // threadId プレフィックスで前方一致検索（スラッシュなし）
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

async function uploadToBlob(buffer: Buffer, blobKey: string, displayFileName?: string): Promise<string> {
  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
  const containerName = "pptx";

  const cred = new StorageSharedKeyCredential(acc, key);
  const svc = BlobServiceClient.fromConnectionString(
    `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
  const cc = svc.getContainerClient(containerName);
  await cc.createIfNotExists({ access: "blob" });

  // Blob key は ASCII のみ（URL短縮）。DL時のファイル名は Content-Disposition で指定
  const bbc = cc.getBlockBlobClient(blobKey);
  await bbc.uploadData(buffer, {
    blobHTTPHeaders: {
      blobContentType:
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      blobContentDisposition: `attachment; filename*=UTF-8''${encodeURIComponent(displayFileName ?? blobKey)}`,
    },
  });

  const sas = generateBlobSASQueryParameters(
    {
      containerName,
      blobName: blobKey,
      expiresOn: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
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

/** Blob URL またはファイル名からベース名を取り出し、suffix を付けた出力ファイル名を生成する。 */
function buildOutputFileName(sourceUrl: string | undefined, suffix: string): string {
  let base = "";
  if (sourceUrl) {
    try {
      const pathname = new URL(sourceUrl.split("?")[0]).pathname;
      const decoded = decodeURIComponent(path.basename(pathname));
      base = path.basename(decoded, path.extname(decoded));
    } catch {
      base = path.basename(sourceUrl.split("?")[0], path.extname(sourceUrl.split("?")[0]));
    }
  }
  // 空・UUID のみ・先頭が threadId パターンの場合はフォールバック
  if (!base || /^[0-9a-f-]{32,}$/i.test(base)) base = "output";
  // Blob Storage で使えない文字を除去（最大100文字）
  const safe = base.replace(/[<>:"/\\|?*\x00-\x1f]/g, "_").slice(0, 100);
  return `${safe}${suffix}`;
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
  allRowsText: string; // 全行を "R行番号: 値1 | 値2 | ..." 形式で結合したテキスト
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
  chartEdits?: Array<{
    sheetName: string;
    chartType: string;       // "line" | "bar" | "scatter" | "pie"
    title: string;           // LLM が内容から生成したグラフタイトル
    xColumn: string;         // X軸の列記号
    yColumns: string[];      // Y軸の列記号リスト
    xLabel?: string;         // X軸ラベル
    yLabel?: string;         // Y軸ラベル（単位があれば含める）
    insertCell?: string;     // グラフ画像の左上セル (例: "F1")
    yDivisor?: number;       // Y値をこの数で割って表示 (例: 1000 → 千円, 10000 → 万円)
    yTickFormat?: string;    // Y軸目盛のフォーマット: "comma"=カンマ区切り整数, "plain"=そのまま整数, "auto"=自動(デフォルト)
    yMin?: number;           // Y軸の表示最小値（表示単位変換後）
    yMax?: number;           // Y軸の表示最大値（表示単位変換後）
    yTickStep?: number;      // Y軸の目盛間隔（表示単位変換後）
    seriesColors?: string[]; // 系列ごとの色 (matplotlib色名 or "#RRGGBB")
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
    // !ref の開始列を取得して実際のExcel列記号にオフセット補正する
    // （データがA列以外から始まる場合にXLSX.jsの配列インデックスがズレるため）
    const ref = ws["!ref"];
    const startColIdx = ref ? XLSX.utils.decode_range(ref).s.c : 0;
    const columns = headerRow.map((header, i) => {
      const letter = XLSX.utils.encode_col(i + startColIdx);
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

    // 全行を "R行番号: 値1 | 値2 | ..." 形式でテキスト化（最大200行）
    const range = ref ? XLSX.utils.decode_range(ref) : null;
    const startRowIdx = range ? range.s.r : 0;

    const allRowsText = rows.slice(0, 200).map((row, rowIdx) => {
      const cells = (row as unknown[]).map((cell, colOffset) => {
        if (cell === null || cell === undefined) return "";
        const colLetter = XLSX.utils.encode_col(startColIdx + colOffset);
        const prefix = `${colLetter}=`;
        if (cell instanceof Date) return `${prefix}${cell.toLocaleDateString("ja-JP")}`;
        return `${prefix}${String(cell).trim()}`;
      }).filter(Boolean);
      return `R${startRowIdx + rowIdx + 1}: ${cells.join(" | ")}`;
    }).join("\n");

    return {
      sheetName,
      rowCount: range ? range.e.r + 1 : rows.length,
      colCount: range ? range.e.c + 1 : columns.length,
      columns,
      sampleRows,
      allRowsText,
    };
  });
}

async function buildExcelEditPlan(
  sheets: SheetSummary[],
  instruction: string,
  previousChartEdits?: ExcelEditPlan["chartEdits"]
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
  ],
  "chartEdits": [
    {
      "sheetName": string,
      "chartType": "line",   // "line" | "bar" | "scatter" | "pie"
      "title": string,        // Read the actual data and write a concise descriptive Japanese title
      "xColumn": "A",         // column letter for X axis (categories / time axis)
      "yColumns": ["B"],      // column letter(s) for Y axis values
      "xLabel": string,       // X axis label (use the column header, e.g. "月")
      "yLabel": string,       // Y axis label with unit AFTER yDivisor (e.g. "売上（千円）")
      "insertCell": "F1",     // top-left anchor cell for the chart image
      "yDivisor": 1,          // divide all Y values by this for display (1=no change, 1000=千円, 10000=万円)
      "yTickFormat": "auto",  // Y axis tick format: "comma"=comma-separated integers (e.g. 1,200), "plain"=plain integer, "auto"=default
      "yMin": number,         // optional Y axis minimum AFTER yDivisor
      "yMax": number,         // optional Y axis maximum AFTER yDivisor
      "yTickStep": number,    // optional Y axis tick interval AFTER yDivisor
      "seriesColors": ["#2196F3"] // optional per-series colors (matplotlib color names or hex "#RRGGBB"). One entry per yColumn.
    }
  ]
}

Rules:
- sheetName must match one of the provided sheet names exactly.
- The "columns" array in each sheet lists column letters AND header names. Each column may also have headerFillColor, headerFontColor, headerBold indicating its current header cell style.
- Always use the column LETTER (e.g. "D") in setCells addresses, copyRowColorEdits, and formatEdits ranges — never use the header name as a substitute for a column letter.
- Use setCells when the user wants to add or update cell values. The value field can be a literal (number or string) OR an Excel formula string starting with "=".
  - When the user asks to calculate a value, use an Excel formula. Examples:
    - 合計 =SUM(B2:B10)
    - 平均 =AVERAGE(B2:B10)
    - 偏差値 (standard score, population std dev, ALL Excel versions compatible — do NOT use STDEV.P or STDEVP):
      =IFERROR(ROUND(50+10*(B2-AVERAGE($B$2:$B$4))/SQRT(SUMPRODUCT(($B$2:$B$4-AVERAGE($B$2:$B$4))^2)/COUNT($B$2:$B$4)),2),"")
      Replace B with the score column letter and 4 with the last data row number (rowCount - 1 if row 1 is header).
      Output one setCells entry per data row (from row 2 to last data row), adjusting only the lone "B2" row number each time (keep $B$2:$B$4 fixed).
      IMPORTANT: Never use STDEV.P, STDEVP, or any function containing a dot. Never prefix "@" before a function name.
  - NEVER add a header name via setCells — set the header through the setCells value directly in row 1.
- Use replaceText ONLY when the user explicitly asks to find and replace text content. NEVER use replaceText for formatting operations.
- Use formatEdits for bold/color changes. fontColor and fillColor are 6-digit hex (no #).
- AUTO HEADER STYLING: Whenever you add a new column (i.e. setCells includes a row-1 cell for a column that does not yet exist), you MUST ALWAYS emit a formatEdits entry for that new header cell that copies the style of the existing styled headers:
  1. Find any existing column that has headerFillColor set — use its headerFillColor, headerFontColor, and headerBold.
  2. Apply all three (fillColor, fontColor, bold) to the new header cell range (e.g. "C1").
  3. This rule is unconditional — apply it even if the user did not mention design or style.
- DESIGN MATCHING (explicit): When the user additionally says "デザインを合わせて" / "同じデザイン" etc., also emit copyRowColorEdits to copy row-by-row background colors from an adjacent styled column to the new column.
- Use copyRowColorEdits when the user wants a column's background colors to match those of another column row-by-row. targetColumn and referenceColumn MUST be column letters from the "columns" list.
- Use borderEdits when the user asks to add borders (枠・罫線・border), frame cells, or make the sheet look cleaner. Infer the data range from the sheet summary. Use edges="all" for full grid, "outer" for outer frame only.
- NEVER set a cell value to an empty string unless the user explicitly asks to clear that cell.
- NEVER modify header row values (row 1) unless the user explicitly asks to change column names.
- Use chartEdits when the user asks to CREATE or MODIFY a chart (グラフ・折れ線グラフ・棒グラフ・散布図・円グラフ・チャートを作成/修正/変更・タイトルを変えて・縦軸を変えて・単位を変えて):
  - When MODIFYING an existing chart (タイトル変更・軸変更・単位変更など), re-emit the full chartEdits entry with ALL parameters including unchanged ones. The old chart will be automatically replaced.
  - Infer chartType: 折れ線→"line", 棒→"bar", 散布図→"scatter", 円→"pie". Default "line".
  - Read all data rows to understand the content, then write a short descriptive Japanese title (e.g. "月別売上推移").
  - xColumn is usually the leftmost column (categories, dates, labels). yColumns are the numeric value columns.
  - xLabel: use the header of xColumn. yLabel: use the yColumn header as the base, then ALWAYS append the unit in parentheses when yDivisor != 1. Examples: header="売上", yDivisor=1000 → yLabel="売上（千円）"; header="売上", yDivisor=10000 → yLabel="売上（万円）"; header="売上", yDivisor=1 → yLabel="売上".
  - yDivisor: set when the user wants a different unit scale. 千円→1000, 万円→10000, 百万円→1000000, 億円→100000000. Default 1 (no scaling). Check the raw data values to determine appropriate divisor (e.g., values like 1200000 are in 円; ÷1000 gives 千円).
  - yTickFormat: "comma" when user asks for comma-separated numbers (数字3桁・カンマ区切り・3桁区切り・1,200形式). Default "auto" (no comma). Use "comma" also when yDivisor is applied and round numbers are expected.
  - yMin/yMax/yTickStep: set only when the user asks to adjust vertical axis numbers/range/tick interval (e.g. "縦軸を0から1500に", "目盛を250刻みに"). Values are AFTER yDivisor.
  - seriesColors: set when the user asks to change bar/line/scatter colors. Use matplotlib color names ("red","blue","green","orange","purple","yellow","black","gray","pink","cyan") or hex "#RRGGBB". Convert Japanese color names: 赤→"red", 青→"blue", 緑→"green", 橙/オレンジ→"orange", 紫→"purple", 黄→"yellow", 黒→"black", 灰/グレー→"gray", ピンク→"pink". One entry per yColumn. For a single series: ["red"].
  - insertCell: find the rightmost column that has data, then place the chart 2 columns further right at row 1. Example: data in C:D → use "F1". NEVER overlap with any data column.
- Only emit the operations the user actually requested. Keep the JSON minimal.`;

  const sheetsSummary = sheets.map((s) => {
    const colDesc = s.columns.map((c) => {
      const style = [
        c.headerFillColor ? `fill:${c.headerFillColor}` : "",
        c.headerFontColor ? `font:${c.headerFontColor}` : "",
        c.headerBold ? "bold" : "",
      ].filter(Boolean).join(",");
      return style ? `${c.letter}(${c.header})[${style}]` : `${c.letter}(${c.header})`;
    }).join(", ");
    return `=== Sheet: ${s.sheetName} | rows:${s.rowCount} cols:${s.colCount} ===
Columns: ${colDesc}
All rows:
${s.allRowsText}`;
  }).join("\n\n");

  const prevChartCtx = previousChartEdits?.length
    ? `\n\n[EXISTING CHART STATE — preserve all unlisted parameters]\n${JSON.stringify(previousChartEdits, null, 2)}\nWhen emitting chartEdits, start from these parameters and apply ONLY the changes the user requested. Copy all other fields as-is.`
    : "";

  const userPrompt = `Instruction: ${instruction}${prevChartCtx}

${sheetsSummary}

Return JSON only.`;

  const res = await openai.chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 4096,
  });

  const content = res.choices[0]?.message?.content ?? "{}";
  const parsed = JSON.parse(content) as ExcelEditPlan;
  parsed.sheetEdits ??= [];
  parsed.formatEdits ??= [];
  parsed.copyRowColorEdits ??= [];
  parsed.borderEdits ??= [];
  parsed.chartEdits ??= [];
  return parsed;
}

function injectMissingHeaderStyles(
  plan: ExcelEditPlan,
  sheets: SheetSummary[]
): ExcelEditPlan {
  const sheetMap = new Map(sheets.map((s) => [s.sheetName, s]));

  for (const sheetEdit of plan.sheetEdits ?? []) {
    const sheet = sheetMap.get(sheetEdit.sheetName);
    if (!sheet) continue;

    const templateCol = sheet.columns.find((c) => c.headerFillColor);
    if (!templateCol) continue;

    const existingLetters = new Set(sheet.columns.map((c) => c.letter));

    const newHeaderLetters = new Set<string>();
    for (const cell of sheetEdit.setCells ?? []) {
      const m = /^([A-Z]+)1$/.exec(cell.address);
      if (m && !existingLetters.has(m[1])) newHeaderLetters.add(m[1]);
    }

    for (const colLetter of Array.from(newHeaderLetters)) {
      const headerRange = `${colLetter}1`;

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

function normalizeExcelChartColumns(
  plan: ExcelEditPlan,
  sheets: SheetSummary[]
): ExcelEditPlan {
  const sheetMap = new Map(sheets.map((s) => [s.sheetName, s]));

  const normalizeColumn = (sheet: SheetSummary, colSpec: string | undefined): string | undefined => {
    const spec = String(colSpec ?? "").trim().toUpperCase();
    if (!spec) return spec;
    if (sheet.columns.some((c) => c.letter === spec)) return spec;
    if (!/^[A-Z]+$/.test(spec)) return spec;

    const ordinal = XLSX.utils.decode_col(spec);
    const actual = sheet.columns[ordinal]?.letter;
    return actual ?? spec;
  };

  for (const chartEdit of plan.chartEdits ?? []) {
    const sheet = sheetMap.get(chartEdit.sheetName);
    if (!sheet) continue;

    const xColumn = normalizeColumn(sheet, chartEdit.xColumn);
    if (xColumn) chartEdit.xColumn = xColumn;

    chartEdit.yColumns = (chartEdit.yColumns ?? []).map((col) => {
      return normalizeColumn(sheet, col) ?? col;
    });
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
      blobContentDisposition: `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
    },
  });

  const sas = generateBlobSASQueryParameters(
    {
      containerName,
      blobName: fileName,
      expiresOn: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
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
    const pyEnv = process.platform !== "win32"
      ? {
          ...process.env,
          PYTHONPATH: `/home/site/python-packages${process.env.PYTHONPATH ? `:${process.env.PYTHONPATH}` : ""}`,
        }
      : process.env;

    if (process.platform !== "win32") {
      try {
        await execFileAsync(pythonBin, ["-c", "import openpyxl"], { env: pyEnv });
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
    ], { env: pyEnv });

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

async function buildAccountNameCorrectionPlan(excelBuffer: Buffer): Promise<ExcelEditPlan> {
  const openai = OpenAIInstance();
  const wb = XLSX.read(excelBuffer, { type: "buffer" });
  const isAccountName = (val: string) =>
    val.trim().length > 0 && !/^[\d,，△▲\-()（）\s]+$/.test(val.trim());
  const sheetEdits: NonNullable<ExcelEditPlan["sheetEdits"]> = [];
  const CORRECTION_TIMEOUT_MS = 90_000;
  const startTime = Date.now();

  outer:
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" }) as string[][];
    if (!data.length) continue;
    const maxCols = Math.max(...data.map((row) => row.length));
    const colIndices: number[] = [];
    for (let c = 0; c < maxCols; c++) {
      const count = data.filter((row) => isAccountName(String(row[c] ?? ""))).length;
      if (count >= 3) colIndices.push(c);
    }

    const setCells: Array<{ address: string; value: string }> = [];

    for (const colIdx of colIndices) {
      if (Date.now() - startTime > CORRECTION_TIMEOUT_MS) {
        console.warn("[pdf-to-excel] correction time limit reached, skipping remaining columns");
        break outer;
      }

      const candidates: { rowIdx: number; name: string }[] = [];
      for (let i = 0; i < data.length; i++) {
        const val = String(data[i]?.[colIdx] ?? "").trim();
        if (isAccountName(val)) candidates.push({ rowIdx: i, name: val });
      }
      if (candidates.length === 0) continue;

      const BATCH = 40;
      for (let batchStart = 0; batchStart < candidates.length; batchStart += BATCH) {
        if (Date.now() - startTime > CORRECTION_TIMEOUT_MS) {
          console.warn("[pdf-to-excel] correction time limit reached, skipping remaining batches");
          break outer;
        }

        const batch = candidates.slice(batchStart, batchStart + BATCH);
        try {
          const res = await openai.chat.completions.create({
            model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
            messages: [
              {
                role: "user",
                content:
                  "以下は日本の財務諸表（貸借対照表・損益計算書・販売費及び一般管理費・製造原価報告書）の列データです。\n" +
                  "TKC会計ソフトのPDFは文字を1文字ずつ配置するため、OCR後に文字化けや断片化が発生します。\n" +
                  "例: 「Ⅱ 上 原壳価」→「Ⅱ 売上原価」、「壳」→「売」、「資 本金」→「資本金」\n" +
                  "例: 「短\\n期貸付\\n金」→「短期貸付金」（改行で分割された文字を結合する）\n\n" +
                  "ルール:\n1. 入力と同じ順序・同じ件数で返してください\n" +
                  "2. 科目名の文字化けや改行による断片化を正しい勘定科目名に補正してください\n" +
                  "3. 補正不要な行はそのまま返してください\n" +
                  "4. ローマ数字（Ⅰ Ⅱ Ⅲ）や括弧も含めてそのまま返してください\n\n" +
                  `以下の入力リスト(${batch.length}件)を補正し、JSON形式 {"corrected": [...]} で返してください:\n` +
                  JSON.stringify(batch.map((r) => r.name)),
              },
            ],
            response_format: { type: "json_object" },
            max_completion_tokens: 8000,
          });
          const finishReason = res.choices[0]?.finish_reason;
          const content = res.choices[0]?.message?.content ?? "{}";
          console.log(`[pdf-to-excel] correction batch finish_reason=${finishReason} items=${batch.length}`);
          try {
            const parsed = JSON.parse(content);
            const corrected: unknown[] = parsed.corrected ?? Object.values(parsed)[0];
            if (!Array.isArray(corrected) || corrected.length !== batch.length) continue;
            for (let i = 0; i < batch.length; i++) {
              const newVal = String(corrected[i] ?? "").trim();
              const oldVal = batch[i].name;
              if (!newVal || newVal === oldVal) continue;
              const addr = XLSX.utils.encode_cell({ r: batch[i].rowIdx, c: colIdx });
              setCells.push({ address: addr, value: newVal });
            }
          } catch {
            console.warn(`[pdf-to-excel] LLM correction JSON parse failed: finish_reason=${finishReason} content_head=${content.slice(0, 200)}`);
          }
        } catch (e) {
          console.warn(`[pdf-to-excel] LLM correction failed:`, String((e as any)?.message ?? e));
        }
      }
    }

    if (setCells.length > 0) {
      sheetEdits.push({ sheetName, setCells });
    }
  }

  return { sheetEdits };
}

async function runPythonPdfToExcel(inputBuffer: Buffer, threadId: string, fileUrl?: string) {
  const inputExt = fileUrl ? (getFileExtension(fileUrl) || ".pdf") : ".pdf";
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-pdf2xl-"));
  const inputPath = path.join(tempDir, `input${inputExt}`);
  const outputPath = path.join(tempDir, "output.xlsx");
  const scriptPath = await resolveConvertPdfScriptPath();

  // PYTHONPATH・LD_LIBRARY_PATH を明示的に設定（startup.sh が動いていない環境でも動作させるため）
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
    let outputBuffer = rawBuffer;
    try {
      const correctionPlan = await buildAccountNameCorrectionPlan(rawBuffer);
      if (correctionPlan.sheetEdits?.length) {
        const correctionDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-correct-"));
        try {
          const corrInPath = path.join(correctionDir, "input.xlsx");
          const corrOutPath = path.join(correctionDir, "output.xlsx");
          const corrPlanPath = path.join(correctionDir, "plan.json");
          const editScriptPath = await resolveEditExcelScriptPath();
          await fs.writeFile(corrInPath, rawBuffer);
          await fs.writeFile(corrPlanPath, JSON.stringify(correctionPlan), "utf8");
          await execFileAsync(pythonBin, [editScriptPath, "--input", corrInPath, "--output", corrOutPath, "--plan", corrPlanPath], { env: pyEnv });
          outputBuffer = await fs.readFile(corrOutPath);
          console.log(`[pdf-to-excel] applied ${correctionPlan.sheetEdits.reduce((n, s) => n + (s.setCells?.length ?? 0), 0)} account name corrections`);
        } finally {
          await fs.rm(correctionDir, { recursive: true, force: true });
        }
      }
    } catch (e) {
      console.warn("[pdf-to-excel] account name correction failed, using uncorrected file:", String((e as any)?.message ?? e));
      outputBuffer = rawBuffer;
    }

    const pythonResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};
    const fileName = buildOutputFileName(fileUrl, "_変換後.xlsx");
    const downloadUrl = await uploadExcelToBlob(outputBuffer, fileName);

    // シート名一覧を取得（refine_excel_pages ツールがLLMに提示するため）
    let sheetNames: string[] = [];
    try {
      const wb = XLSX.read(outputBuffer, { type: "buffer" });
      sheetNames = wb.SheetNames;
    } catch {
      // 取得失敗は非致命的
    }

    return {
      downloadUrl,
      fileName,
      sheetNames,
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
// Excel 精度向上（GPT-4V リファイン）
// ─────────────────────────────────────────────

async function runPythonRefineExcelPages(
  excelBuffer: Buffer,
  targetSheets: string[],
  threadId: string,
  outputFileName?: string
): Promise<{ downloadUrl: string; fileName: string; refined: number; skipped: number }> {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-refine-"));
  const excelInPath = path.join(tempDir, "input.xlsx");
  const excelOutPath = path.join(tempDir, "output.xlsx");
  const scriptPath = await resolveConvertPdfScriptPath();

  const pyEnv = process.platform !== "win32"
    ? {
        ...process.env,
        PYTHONPATH: `/home/site/python-packages${process.env.PYTHONPATH ? `:${process.env.PYTHONPATH}` : ""}`,
        LD_LIBRARY_PATH: `/home/site/python-packages${process.env.LD_LIBRARY_PATH ? `:${process.env.LD_LIBRARY_PATH}` : ""}`,
      }
    : process.env;

  try {
    await fs.writeFile(excelInPath, excelBuffer);

    const pythonBin = process.platform === "win32" ? "python" : "python3";
    const { stdout, stderr } = await execFileAsync(pythonBin, [
      scriptPath,
      "--input", excelInPath,
      "--output", excelOutPath,
      "--refine-sheets", targetSheets.join(","),
    ], { env: pyEnv });

    if (stderr?.trim()) {
      console.warn("[refine-excel] python stderr:", stderr.trim());
    }

    const outputBuffer = await fs.readFile(excelOutPath);
    const pythonResult = stdout?.trim() ? JSON.parse(stdout.trim()) : {};
    const fileName = outputFileName ?? buildOutputFileName(undefined, "_精度向上後.xlsx");
    const downloadUrl = await uploadExcelToBlob(outputBuffer, fileName);

    return {
      downloadUrl,
      fileName,
      refined: Number(pythonResult.refined ?? 0),
      skipped: Number(pythonResult.skipped ?? 0),
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
  }>;
  trackChanges?: boolean;
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
      "fontSize": 12
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
  - Paragraphs are appended at the end of the document.
  - NEVER use addParagraphs to append a summary of changes made (e.g. "修正箇所一覧", "変更点一覧"). Word comments are inserted automatically at each change location.
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
    max_completion_tokens: 4000,
  });

  const raw = res.choices[0]?.message?.content ?? "{}";
  const finishReason = res.choices[0]?.finish_reason;
  if (finishReason === "length") {
    console.warn("[buildWordEditPlan] LLM response truncated (finish_reason=length). Plan may be incomplete.");
  }
  let parsed: WordEditPlan;
  try {
    parsed = JSON.parse(raw) as WordEditPlan;
  } catch {
    console.error("[buildWordEditPlan] JSON.parse failed. raw=", raw.slice(0, 200));
    throw new Error("編集プランの生成に失敗しました（JSONパースエラー）。指示を短くして再試行してください。");
  }
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

async function uploadWordToBlob(buffer: Buffer, fileName: string, displayName?: string): Promise<string> {
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
      blobContentDisposition: `attachment; filename*=UTF-8''${encodeURIComponent(displayName ?? fileName)}`,
    },
  });

  const sas = generateBlobSASQueryParameters(
    {
      containerName,
      blobName: fileName,
      expiresOn: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
      permissions: BlobSASPermissions.parse("r"),
    },
    cred
  );
  return `${bbc.url}?${sas}`;
}

async function runPythonEditWord(
  inputBuffer: Buffer,
  plan: WordEditPlan,
  threadId: string,
  originalFileName?: string
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
    const blobKey = `${threadId || uniqueId()}_edited_${uniqueId()}.docx`;
    const displayName = originalFileName
      ? `${originalFileName.replace(/\.docx$/i, "")}_rev1.docx`
      : blobKey;
    const downloadUrl = await uploadWordToBlob(outputBuffer, blobKey, displayName);

    return {
      downloadUrl,
      fileName: displayName,
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

async function runPythonEdit(inputBuffer: Buffer, plan: EditPlan, threadId: string, fileBaseName?: string) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "azurechat-pptx-"));
  const inputPath = path.join(tempDir, "input.pptx");
  const outputPath = path.join(tempDir, "output.pptx");
  const planPath = path.join(tempDir, "plan.json");
  const scriptPath = await resolveEditPptxScriptPath();

  try {
    await fs.writeFile(inputPath, inputBuffer);

    // imageInserts がある場合: slideIndex=-1 を全スライドに展開し、画像を準備
    const requestedImages = plan.imageInserts?.length ?? 0;
    let generatedImages = 0;
    if (requestedImages > 0) {
      // slideIndex: -1 → 全スライドに複製展開（Python側に渡す前に解決）
      const totalSlideCount = (() => {
        try {
          // 実際のスライド数は Python 実行後にしか分からないため、-1 はそのまま渡す
          // Python 側で -1 を全スライドと解釈する
          return -1;
        } catch { return -1; }
      })();
      void totalSlideCount; // 現在は Python 側で処理

      const urlCache = new Map<string, Buffer>();
      for (let i = 0; i < plan.imageInserts!.length; i++) {
        const insert = plan.imageInserts![i];
        if (insert.imageUrl) {
          // 既存画像URL（ロゴ等）をダウンロード — 同じURLは1回のみ
          const cached = urlCache.get(insert.imageUrl);
          if (cached) {
            const imagePath = path.join(tempDir, `image_${i}.png`);
            await fs.writeFile(imagePath, cached);
            insert.imagePath = imagePath;
            generatedImages++;
            console.log(`[edit-pptx] image ${i} reused from cache`);
          } else {
            try {
              console.log(`[edit-pptx] downloading image ${i}: ${insert.imageUrl.slice(0, 80)}`);
              const resp = await fetch(insert.imageUrl);
              if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
              const buf = Buffer.from(await resp.arrayBuffer());
              urlCache.set(insert.imageUrl, buf);
              const imagePath = path.join(tempDir, `image_${i}.png`);
              await fs.writeFile(imagePath, buf);
              insert.imagePath = imagePath;
              generatedImages++;
              console.log(`[edit-pptx] image ${i} downloaded (${buf.length} bytes)`);
            } catch (e) {
              console.warn(`[edit-pptx] image ${i} download failed:`, e);
            }
          }
        } else if (insert.imagePrompt) {
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
          "サーバー管理者に startup.sh の設定を確認してください。"
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
    const safeDisplay = (fileBaseName || "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim()
      .slice(0, 60);
    const displayFileName = safeDisplay ? `${safeDisplay}.pptx` : undefined;
    const blobKey = `pptx_${uniqueId().slice(0, 8)}.pptx`;  // ASCII のみ（URL短縮）
    const fileName = displayFileName ?? blobKey;
    const downloadUrl = await uploadToBlob(outputBuffer, blobKey, displayFileName);

    const insertedImages = Number(pythonResult.insertedImages ?? 0);
    const imageWarning =
      requestedImages > 0 && insertedImages < requestedImages
        ? `画像挿入: ${requestedImages}件要求 / ${insertedImages}件成功`
        : undefined;

    return {
      downloadUrl,
      fileName,
      outputBuffer,
      changedSlides: Number(pythonResult.changedSlides ?? 0),
      changedSlideIndices: Array.isArray(pythonResult.changedSlideIndices)
        ? (pythonResult.changedSlideIndices as number[])
        : [],
      totalSlides: Number(pythonResult.totalSlides ?? 0),
      requestedImages,
      insertedImages,
      charsBefore: Number(pythonResult.charsBefore ?? 0),
      charsAfter: Number(pythonResult.charsAfter ?? 0),
      ...(imageWarning ? { imageWarning } : {}),
      ...(Array.isArray(pythonResult.outOfRangeSlides) && pythonResult.outOfRangeSlides.length > 0
        ? { outOfRangeSlides: pythonResult.outOfRangeSlides as number[] }
        : {}),
      ...(Array.isArray(pythonResult.overflowCandidates) && pythonResult.overflowCandidates.length > 0
        ? { overflowCandidates: pythonResult.overflowCandidates as string[] }
        : {}),
      ...(Array.isArray(pythonResult.layoutWarnings) && pythonResult.layoutWarnings.length > 0
        ? { layoutWarnings: pythonResult.layoutWarnings as string[] }
        : {}),
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

/** edit-pptx 後に変更スライドのみ Vision レビューして目視警告を返す（対応3）*/
async function runVisionReviewAfterEdit(
  outputBuffer: Buffer,
  changedSlideIndices: number[],
  baseUrl: string
): Promise<string[]> {
  try {
    const reviewForm = new FormData();
    reviewForm.append(
      "pptx",
      new Blob([new Uint8Array(outputBuffer)], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" })
    );
    reviewForm.append("title", "edit-review");

    const res = await fetch(`${baseUrl}/api/vision-review-pptx`, {
      method: "POST",
      body: reviewForm,
      signal: AbortSignal.timeout(90_000),
    });
    if (!res.ok) return [];

    const review = await res.json() as { fixes?: Array<{ slideIndex: number; field: string; value: string; reason?: string }> };
    const warnings: string[] = [];
    const overflowFields = new Set(["fitTextToShape", "fontScaleDown", "trimText", "syncItemDecorations", "copyItemDecoration", "alignItemGroup", "fallbackLayout"]);

    for (const fix of (review.fixes ?? [])) {
      if (!changedSlideIndices.includes(fix.slideIndex)) continue;
      if (overflowFields.has(fix.field)) {
        warnings.push(`Page${fix.slideIndex + 1}: ${fix.reason ?? fix.field}`);
      }
    }

    if (warnings.length > 0) {
      console.log(`[edit-pptx] vision post-review warnings: ${warnings.join(" / ")}`);
    }
    return warnings;
  } catch (e) {
    console.warn("[edit-pptx] vision post-review failed (non-fatal):", String((e as Error)?.message ?? e).slice(0, 120));
    return [];
  }
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { fileUrl, instruction, threadId, action, mode, previousChartEdits, outputBaseName, plan: incomingPlan, trackChanges, excelFileUrl, targetSheets, outputFileName } = body as {
      fileUrl: string;
      instruction: string;
      threadId: string;
      action?: string;
      mode?: string;
      previousChartEdits?: ExcelEditPlan["chartEdits"];
      outputBaseName?: string;
      plan?: EditPlan;
      trackChanges?: boolean;
      excelFileUrl?: string;
      targetSheets?: string[];
      outputFileName?: string;
    };

    // Excel 指定シート精度向上（GPT-4V リファイン）— fileUrl 不要なので先に処理
    if (action === "refine_excel_pages") {
      if (!excelFileUrl?.trim() || !Array.isArray(targetSheets) || targetSheets.length === 0) {
        return NextResponse.json({ ok: false, error: "excelFileUrl and targetSheets are required" }, { status: 400 });
      }
      const excelBuffer = await downloadBlob(excelFileUrl, threadId);
      const result = await runPythonRefineExcelPages(excelBuffer, targetSheets, threadId, outputFileName);
      console.log("[refine-excel] result:", JSON.stringify(result));
      return NextResponse.json({ ok: true, ...result });
    }

    if (!fileUrl?.trim() || (!instruction?.trim() && action !== "pdf_to_excel" && action !== "pdf_to_word" && action !== "extract_pptx_summary" && action !== "apply_pptx_plan")) {
      return NextResponse.json(
        { ok: false, error: "fileUrl and instruction are required" },
        { status: 400 }
      );
    }

    const ext = getFileExtension(fileUrl);
    console.log(`[edit-pptx] fileUrl =`, fileUrl.substring(0, 80));
    const instructionText = instruction ?? "";
    console.log(`[edit-pptx] ext =`, ext, "action =", action ?? "(none)", "instruction =", instructionText.substring(0, 120));

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

    // PPTX スライド構造抽出（編集・アップロードなし）
    if (action === "extract_pptx_summary") {
      const pptxBuffer = await downloadBlob(fileUrl, threadId);
      const slides = await extractSlidesStructured(pptxBuffer);
      console.log(`[extract_pptx_summary] extracted ${slides.length} slides`);
      return NextResponse.json({ ok: true, slides, totalSlides: slides.length });
    }

    // EditPlan を直接受け取って既存PPTXに適用（replaceText のみ・レイアウト再生成なし）
    if (action === "apply_pptx_plan") {
      if (!incomingPlan || !Array.isArray(incomingPlan.slideEdits)) {
        return NextResponse.json({ ok: false, error: "plan.slideEdits is required" }, { status: 400 });
      }
      // deckEdits / imageInserts は内容増量では不要のため除去。item 単位バリデーション付き。
      const safeplan: EditPlan = {
        slideEdits: incomingPlan.slideEdits
          .map((se) => {
            const validReplace = (se.replaceText ?? []).filter((item) => {
              const find = String((item as any).find ?? "").trim();
              const replace = String((item as any).replace ?? "").trim();
              const appendToRun = String((item as any).appendToRun ?? "").trim();
              return find && (replace || appendToRun);
            });
            const validBullets = (se.addBullets ?? []).filter((item) => {
              const texts = Array.isArray((item as any).texts)
                ? (item as any).texts.filter((t: unknown) => String(t ?? "").trim())
                : [];
              return texts.length > 0;
            });
            const csb = (se as any).copyShapeBlock;
            const rawCsbGroup: unknown = csb?.groupShapeNames;
            const validGroupNames: string[] | undefined =
              Array.isArray(rawCsbGroup) && rawCsbGroup.length >= 2
                ? (rawCsbGroup as unknown[]).map((n) => String(n).trim()).filter(Boolean)
                : undefined;
            const validCopyBlock =
              csb &&
              typeof csb.headingShapeName === "string" && csb.headingShapeName.trim() &&
              typeof csb.descShapeName === "string" && csb.descShapeName.trim()
                ? { headingShapeName: csb.headingShapeName.trim(), descShapeName: csb.descShapeName.trim(),
                    headingText: String(csb.headingText ?? "").trim(), descText: String(csb.descText ?? "").trim(),
                    ...(validGroupNames ? { groupShapeNames: validGroupNames } : {}) }
                : undefined;
            const rawConvertToCards = (se as any).convertToCards;
            const validCards = Array.isArray(rawConvertToCards?.cards)
              ? rawConvertToCards.cards
                  .map((card: any) => ({
                    heading: String(card?.heading ?? "").trim().slice(0, 40),
                    body: String(card?.body ?? "").trim().slice(0, 180),
                    iconKey: String(card?.iconKey ?? "").trim().slice(0, 24),
                  }))
                  .filter((card: { heading: string; body: string }) => card.heading || card.body)
                  .slice(0, 6)
              : [];
            const validConvertToCards = validCards.length > 0
              ? { cards: validCards }
              : undefined;
            const hasValid = validReplace.length > 0 || validBullets.length > 0 || !!validCopyBlock || !!validConvertToCards;
            if (!hasValid) return null;
            return {
              slideIndex: se.slideIndex,
              ...(validReplace.length > 0 ? { replaceText: validReplace } : {}),
              ...(validBullets.length > 0 ? { addBullets: validBullets } : {}),
              ...(validCopyBlock ? { copyShapeBlock: validCopyBlock } : {}),
              ...(validConvertToCards ? { convertToCards: validConvertToCards } : {}),
            };
          })
          .filter((se): se is NonNullable<typeof se> => se !== null),
      };
      const pptxBuffer = await downloadBlob(fileUrl, threadId);
      const safeBaseName = outputBaseName ?? "内容増量";
      const result = await runPythonEdit(pptxBuffer, safeplan, threadId, safeBaseName);
      console.log(`[apply_pptx_plan] changedSlides=${result.changedSlides} charsBefore=${result.charsBefore} charsAfter=${result.charsAfter}`);

      // 対応3: addBullets/copyShapeBlock 後に Vision レビューして目視警告を返す
      let applyVisualWarnings: string[] | undefined;
      const hasApplyBulletOrShapeEdit = (safeplan.slideEdits ?? []).some(
        (se) =>
          Array.isArray((se as any).addBullets) ||
          (se as any).copyShapeBlock ||
          (se as any).convertToCards ||
          Array.isArray((se as any).replaceText)
      );
      if (
        process.env.PPTX_EDIT_VISION_REVIEW_ENABLED === "true" &&
        hasApplyBulletOrShapeEdit &&
        result.changedSlides > 0 &&
        result.changedSlideIndices.length > 0
      ) {
        const baseUrl = (
          process.env.NEXTAUTH_URL ||
          (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
        ).replace(/\/+$/, "");
        const warnings = await runVisionReviewAfterEdit(result.outputBuffer, result.changedSlideIndices, baseUrl);
        if (warnings.length > 0) applyVisualWarnings = warnings;
      }

      const { outputBuffer: _applyBuf, ...applyResult } = result;
      void _applyBuf;
      return NextResponse.json({ ok: true, ...applyResult, ...(applyVisualWarnings ? { visualWarnings: applyVisualWarnings } : {}) });
    }

    // Excel ファイル (.xlsx / .xls / .xlsm) の場合は Excel 専用フローへ
    if (isExcelFile(ext)) {
      const excelBuffer = await downloadBlob(fileUrl, threadId);
      const sheets = await extractSheetSummaries(excelBuffer);
      console.log("[edit-excel] sheets:", JSON.stringify(sheets.map(s => ({ name: s.sheetName, columns: s.columns }))));
      const rawPlan = await buildExcelEditPlan(sheets, instruction, previousChartEdits);
      const plan = injectMissingHeaderStyles(normalizeExcelChartColumns(rawPlan, sheets), sheets);

      console.log("[edit-excel] plan:", JSON.stringify(plan));

      const result = await runPythonEditExcel(excelBuffer, ext, plan, threadId);
      return NextResponse.json({
        ok: true,
        ...result,
        appliedChartEdits: plan.chartEdits?.length ? plan.chartEdits : undefined,
      });
    }

    // Word ファイル (.docx) の場合は Word 専用フローへ
    if (isWordFile(ext)) {
      const wordBuffer = await downloadBlob(fileUrl, threadId);
      const summary = await extractDocSummary(wordBuffer);
      const plan = await buildWordEditPlan(summary, instruction);
      plan.trackChanges = !!trackChanges;

      console.log("[edit-word] plan:", JSON.stringify(plan));

      const rawName = fileUrl.split("?")[0].split("/").pop() ?? "";
      const originalFileName = (() => { try { return decodeURIComponent(rawName); } catch { return rawName; } })();

      const result = await runPythonEditWord(wordBuffer, plan, threadId, originalFileName || undefined);
      return NextResponse.json({ ok: true, ...result });
    }

    // PPTX フロー（既存）
    const pptxBuffer = await downloadBlob(fileUrl, threadId);
    const slides = await extractSlidesStructured(pptxBuffer);
    let plan = await buildEditPlan(slides, instruction);

    // deckEdits はユーザーが明示的にスライド追加・削除を要求した場合のみ許可
    // (それ以外の編集指示で accentColor/fontFace が混入すると別物化する)
    const deckEditAllowed =
      /スライド.{0,8}(追加|挿入|削除)|ページ.{0,8}(追加|削除)|デザイン.{0,8}変更|カラー.{0,8}変更|色.{0,8}(変え|変更|かえ|にして|替え)|フォント.{0,8}(変え|変更|かえ)|アクセント|(緑|青|赤|黄|紫|オレンジ|ピンク|ネイビー|グレー|グリーン|ブルー|レッド|green|blue|red|yellow|purple|orange|pink|navy|gray).{0,10}(にして|にかえ|に変え|に変更)/.test(instruction);
    if (!deckEditAllowed && plan.deckEdits) {
      console.log("[edit-pptx] stripping deckEdits (not explicitly requested)");
      plan = { ...plan, deckEdits: undefined };
    }

    console.log("[edit-pptx] plan:", JSON.stringify(plan));

    const slideCountBefore = slides.length;
    const result = await runPythonEdit(pptxBuffer, plan, threadId, outputBaseName);

    // 変更なし検証: 0件置換は無言成功ではなくエラーとして返す
    // deckEdits（テーマ色・フォント変更）はchangedSlides=0でもtheme XMLが更新されるため除外
    const hasDeckChange = !!(plan.deckEdits?.accentColor || plan.deckEdits?.fontFace);
    if ((result.changedSlides ?? 0) <= 0 && (plan.imageInserts?.length ?? 0) === 0 && !hasDeckChange) {
      return NextResponse.json({
        ok: false,
        error: "置換対象が見つからず、内容は変更されませんでした。文字列が複数のrunに分割されているか、指定した文字列が見つかりません。",
      });
    }
    // スライド数変化検証: 編集依頼でページ数が増減するのは異常
    if (result.totalSlides !== slideCountBefore) {
      console.error(`[edit-pptx] slide count changed ${slideCountBefore} → ${result.totalSlides}`);
      return NextResponse.json({
        ok: false,
        error: `スライド数が変化しました（${slideCountBefore}枚 → ${result.totalSlides}枚）。編集ではなく再生成が発生した可能性があります。`,
      });
    }

    // 対応3: addBullets/copyShapeBlock 後に Vision レビューして目視警告を返す
    // PPTX_EDIT_VISION_REVIEW_ENABLED=true の場合のみ実行（レイテンシ ~30秒）
    let visualWarnings: string[] | undefined;
    const hasBulletOrShapeEdit = (plan.slideEdits ?? []).some(
      (se) => Array.isArray((se as any).addBullets) || (se as any).copyShapeBlock
    );
    if (
      process.env.PPTX_EDIT_VISION_REVIEW_ENABLED === "true" &&
      hasBulletOrShapeEdit &&
      result.changedSlides > 0 &&
      result.changedSlideIndices.length > 0
    ) {
      const baseUrl = (
        process.env.NEXTAUTH_URL ||
        (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
      ).replace(/\/+$/, "");
      const warnings = await runVisionReviewAfterEdit(result.outputBuffer, result.changedSlideIndices, baseUrl);
      if (warnings.length > 0) visualWarnings = warnings;
    }

    // outputBuffer は外部に公開しない（サイズが大きいため）
    const { outputBuffer: _buf, ...safeResult } = result;
    void _buf;
    return NextResponse.json({
      ok: true,
      ...safeResult,
      ...(visualWarnings ? { visualWarnings } : {}),
    });
  } catch (e: any) {
    console.error("[edit-pptx] error:", e);
    return NextResponse.json(
      { ok: false, error: String(e?.message ?? e) },
      { status: 500 }
    );
  }
}
