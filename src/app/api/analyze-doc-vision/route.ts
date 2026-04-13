export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import path from "node:path";
import { BlobServiceClient } from "@azure/storage-blob";
import { OpenAIVisionInstance } from "@/features/common/services/openai";

const MAX_PAGES = 30;
const PDF_RENDER_SCALE = 1.75;
const VISION_DETAIL: "low" | "high" | "auto" = "auto";
const VISION_MAX_COMPLETION_TOKENS = 900;
const VISION_MAX_RETRIES = 3;
const RETRY_BASE_DELAY_MS = 1500;

type VisualBlock = {
  kind: "callout" | "node" | "badge" | "figure";
  role?: "primary" | "supporting" | "annotation";
  groupId?: string;
  text: string;
  x: number;
  y: number;
  w: number;
  h: number;
  emphasis?: boolean;
};

type Connector = {
  from: number;
  to: number;
  label?: string;
  style?: "arrow" | "line";
  relationshipType?: "flow" | "compare" | "annotation" | "support";
};

export type ConversationTurn = {
  speakerRole: string;
  speakerType?: "agent" | "customer" | "staff" | "other";
  text: string;
  turnIndex: number;
};

type SlideAnalysis = {
  slideTitle: string;
  bullets: string[];
  layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram" | "conversation";
  tableRows?: string[][];
  columns?: Array<{ header: string; bullets: string[] }>;
  visualBlocks?: VisualBlock[];
  connectors?: Connector[];
  conversationStyle?: "chat-ui" | "interview" | "dialog-list";
  conversationTurns?: ConversationTurn[];
};

function extractJsonObject(text: string): string {
  const trimmed = String(text ?? "").trim();
  if (!trimmed) {
    throw new Error("Vision returned empty content");
  }

  const fencedMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  if (fencedMatch?.[1]) {
    return fencedMatch[1].trim();
  }

  const firstBrace = trimmed.indexOf("{");
  const lastBrace = trimmed.lastIndexOf("}");
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    return trimmed.slice(firstBrace, lastBrace + 1);
  }

  return trimmed;
}

function parseVisionJson(text: string): Record<string, unknown> {
  const candidate = extractJsonObject(text);
  return JSON.parse(candidate) as Record<string, unknown>;
}

function extractPartialSlide(raw: string, pageIndex: number): SlideAnalysis {
  const titleMatch = raw.match(/"slideTitle"\s*:\s*"([^"]+)"/);
  const title = titleMatch?.[1] ?? `Slide ${pageIndex + 1}`;
  const bulletsArrayMatch = raw.match(/"bullets"\s*:\s*\[([^\]]*)/);
  const bullets: string[] = [];
  if (bulletsArrayMatch?.[1]) {
    const itemMatches = Array.from(bulletsArrayMatch[1].matchAll(/"([^"]+)"/g));
    bullets.push(...itemMatches.map((m) => m[1]).filter(Boolean));
  }
  return { slideTitle: title, bullets, layoutType: "bullets" };
}

function buildVisionPrompt(pageIndex: number, totalPages: number, mode?: "faithful" | "redesign"): string {
  const faithfulRules = mode === "faithful" ? `
FAITHFUL MODE — strict preservation rules:
- This is an EXISTING slide deck. Preserve the original structure faithfully.
- Do NOT summarize, rewrite, or compress content. Extract text as-is.
- If the page is already a bullet-point slide, keep all bullets in the same order.
- Do NOT merge multiple points into one. Keep each bullet as a separate item.
- Section headers visible on the slide (e.g. colored boxes with text) should be output as "[H] header text" bullets.
- Preserve the slide title exactly as it appears. Do NOT add "[H]" to the title.
- For conversation pages: extract EVERY utterance verbatim. Do NOT shorten or summarize dialogue.` : `
- Use "[H] ..." prefix for section headers within bullets.`;

  return `Analyze page ${pageIndex + 1} of ${totalPages}. Return JSON only.

Schema:
{
  "slideTitle": "Main title",
  "layoutType": "bullets",
  "bullets": ["[H] Section header", "Bullet A", "Bullet B"],
  "tableRows": [["Header 1", "Header 2"], ["Value 1", "Value 2"]],
  "columns": [
    { "header": "Column 1", "bullets": ["Point A", "Point B"] },
    { "header": "Column 2", "bullets": ["Point C", "Point D"] }
  ],
  "conversationStyle": "chat-ui",
  "conversationTurns": [
    { "speakerRole": "ボット", "speakerType": "agent", "text": "ご相談ありがとうございます。まず住所を教えてください。", "turnIndex": 0 },
    { "speakerRole": "ユーザー", "speakerType": "customer", "text": "浜松市中区〇〇です。", "turnIndex": 1 }
  ],
  "visualBlocks": [
    { "kind": "callout", "role": "annotation", "groupId": "g1", "text": "Label", "x": 8, "y": 20, "w": 24, "h": 16, "emphasis": true }
  ],
  "connectors": [
    { "from": 0, "to": 1, "label": "relates", "style": "arrow", "relationshipType": "annotation" }
  ]
}

Rules:
- Use "title" for title-only pages.
- Use "bullets" for standard content pages.
- Use "table" when a table is the primary visual.
- Use "multi-column" for clear parallel columns.
- Use "diagram" when the page is mainly boxes, arrows, callouts, or a visual relationship map (NOT a chat conversation).
- Use "conversation" when the page shows a chat UI or role-play dialogue with alternating speakers.
  Conversation detection triggers when ANY of these are true:
  * Speaker labels appear (ボット/Bot, ユーザー/User, or similar roles) alternating in sequence
  * Speech bubbles or rounded message boxes appear in a vertical sequence
  * The page title contains ロールプレイ, チャット, 会話, dialogue, or role-play
  * Alternating left/right message boxes are visible (chat-style UI)
- For conversation pages: populate conversationTurns[] with every speaker turn in order.
  * Set speakerRole to the EXACT label shown on the slide (e.g. "ボット", "ユーザー", "面接官", "応募者", "Bot", "User").
  * Set speakerType: "agent" for bots/AI/staff roles, "customer" for user/client roles, "staff" for human staff, "other" otherwise.
  * Extract full text verbatim (do NOT summarize). Leave visualBlocks empty.
  * Set conversationStyle: "chat-ui" for L/R bubble layouts, "interview" for Q&A role-contrast layouts, "dialog-list" for sequential vertical lists without bubbles.
${mode === "faithful" ? "- Extract ALL text as-is. Do NOT summarize or compress. Preserve every bullet and label." : "- Keep bullets concise. Keep item text short, usually under 45 characters."}
- For diagram pages, include 2-8 visualBlocks and 0-8 connectors.
- Set visualBlocks.role to "primary", "supporting", or "annotation" when possible.
- Add groupId when multiple blocks belong to the same area, lane, or subtopic.
- Set connectors.relationshipType to "flow", "compare", "annotation", or "support" when possible.
- visualBlocks coordinates use a 0-100 canvas.
- Preserve important labels and hierarchy from the page.
- Return valid JSON only.
${faithfulRules}`;
}

async function analyzePageWithVision(
  base64Image: string,
  pageIndex: number,
  totalPages: number,
  mode?: "faithful" | "redesign"
): Promise<SlideAnalysis> {
  const openai = OpenAIVisionInstance();

  for (let attempt = 1; attempt <= VISION_MAX_RETRIES; attempt++) {
    try {
      const res = await openai.chat.completions.create({
        model: process.env.AZURE_OPENAI_VISION_API_DEPLOYMENT_NAME!,
        messages: [
          {
            role: "user",
            content: [
              {
                type: "text",
                text: buildVisionPrompt(pageIndex, totalPages, mode),
              },
              {
                type: "image_url",
                image_url: {
                  url: `data:image/png;base64,${base64Image}`,
                  detail: VISION_DETAIL,
                },
              },
            ],
          },
        ],
        response_format: { type: "json_object" },
        // faithfulモードは全テキスト抽出のため出力が長くなる → 上限を2倍に拡大
        max_completion_tokens: mode === "faithful" ? VISION_MAX_COMPLETION_TOKENS * 2 : VISION_MAX_COMPLETION_TOKENS,
      });

      const text = res.choices[0]?.message?.content ?? "";
      const finishReason = res.choices[0]?.finish_reason;

      // トークン上限で切れた場合はリトライ対象とする
      if (finishReason === "length" || !text) {
        console.warn(`[analyze-doc-vision] page ${pageIndex + 1} truncated (finish_reason=${finishReason}), attempt ${attempt}`);
        if (attempt < VISION_MAX_RETRIES) {
          await new Promise((resolve) => setTimeout(resolve, RETRY_BASE_DELAY_MS * attempt));
          continue;
        }
        // 最終試行でも切れた → 部分抽出してフォールバック
        console.warn(`[analyze-doc-vision] page ${pageIndex + 1} still truncated after ${VISION_MAX_RETRIES} attempts, using partial extraction`);
        return extractPartialSlide(text, pageIndex);
      }

      let parsed: Record<string, unknown>;
      try {
        parsed = parseVisionJson(text);
      } catch (parseErr) {
        console.warn(`[analyze-doc-vision] page ${pageIndex + 1} JSON parse failed (attempt ${attempt}):`, parseErr);
        if (attempt < VISION_MAX_RETRIES) {
          await new Promise((resolve) => setTimeout(resolve, RETRY_BASE_DELAY_MS * attempt));
          continue;
        }
        console.warn(`[analyze-doc-vision] page ${pageIndex + 1} parse failed after all retries, using partial extraction`);
        return extractPartialSlide(text, pageIndex);
      }

      const VALID_LAYOUTS = ["title", "bullets", "table", "multi-column", "diagram", "conversation"] as const;
      type ValidLayout = (typeof VALID_LAYOUTS)[number];
      const layoutType: ValidLayout = VALID_LAYOUTS.includes(parsed.layoutType as ValidLayout)
        ? (parsed.layoutType as ValidLayout)
        : "bullets";

      const tableRows = Array.isArray(parsed.tableRows)
        ? parsed.tableRows
            .map((row: unknown) =>
              Array.isArray(row) ? row.map((cell) => String(cell ?? "").trim()) : []
            )
            .filter((row: string[]) => row.some(Boolean))
        : [];

      const columns = Array.isArray(parsed.columns)
        ? parsed.columns
            .map((col: unknown) => {
              if (!col || typeof col !== "object") return null;
              const c = col as Record<string, unknown>;
              return {
                header: String(c.header ?? "").trim(),
                bullets: Array.isArray(c.bullets)
                  ? c.bullets.map((b) => String(b ?? "").trim()).filter(Boolean)
                  : [],
              };
            })
            .filter(
              (
                c: { header: string; bullets: string[] } | null
              ): c is { header: string; bullets: string[] } => c !== null && c.header.length > 0
            )
        : [];

      const visualBlocks: VisualBlock[] = Array.isArray(parsed.visualBlocks)
        ? parsed.visualBlocks
            .map((block: unknown): VisualBlock | null => {
              if (!block || typeof block !== "object") return null;
              const b = block as Record<string, unknown>;
              const kind = String(b.kind ?? "").trim();
              if (!["callout", "node", "badge", "figure"].includes(kind)) return null;
              const textValue = String(b.text ?? "").trim();
              const x = Number(b.x);
              const y = Number(b.y);
              const w = Number(b.w);
              const h = Number(b.h);
              if (!textValue || !Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(w) || !Number.isFinite(h)) {
                return null;
              }
              const roleRaw = String(b.role ?? "").trim();
              return {
                kind: kind as VisualBlock["kind"],
                role: (roleRaw === "primary" || roleRaw === "supporting" || roleRaw === "annotation")
                  ? (roleRaw as VisualBlock["role"])
                  : undefined,
                groupId: String(b.groupId ?? "").trim() || undefined,
                text: textValue,
                x: Math.max(0, Math.min(100, x)),
                y: Math.max(0, Math.min(100, y)),
                w: Math.max(6, Math.min(100, w)),
                h: Math.max(6, Math.min(100, h)),
                emphasis: Boolean(b.emphasis),
              };
            })
            .filter((block): block is VisualBlock => block !== null)
        : [];

      const connectors: Connector[] = Array.isArray(parsed.connectors)
        ? parsed.connectors
            .map((conn: unknown): Connector | null => {
              if (!conn || typeof conn !== "object") return null;
              const c = conn as Record<string, unknown>;
              const from = Number(c.from);
              const to = Number(c.to);
              if (!Number.isInteger(from) || !Number.isInteger(to)) return null;
              const relRaw = String(c.relationshipType ?? "").trim();
              return {
                from,
                to,
                label: String(c.label ?? "").trim(),
                style: String(c.style ?? "arrow").trim() === "line" ? "line" : "arrow",
                relationshipType: (relRaw === "flow" || relRaw === "compare" || relRaw === "annotation" || relRaw === "support")
                  ? (relRaw as Connector["relationshipType"])
                  : undefined,
              };
            })
            .filter((conn): conn is Connector => conn !== null)
        : [];

      const conversationTurns: ConversationTurn[] = Array.isArray(parsed.conversationTurns)
        ? parsed.conversationTurns
            .map((turn: unknown, idx: number): ConversationTurn | null => {
              if (!turn || typeof turn !== "object") return null;
              const t = turn as Record<string, unknown>;
              // Accept speakerRole (new) or fall back to legacy speaker field
              const speakerRole = String(t.speakerRole ?? t.speaker ?? "").trim();
              const text = String(t.text ?? "").trim();
              if (!text || !speakerRole) return null;
              const rawType = String(t.speakerType ?? "").trim().toLowerCase();
              const speakerType: ConversationTurn["speakerType"] =
                rawType === "agent" || rawType === "customer" || rawType === "staff" || rawType === "other"
                  ? (rawType as ConversationTurn["speakerType"])
                  : undefined;
              return {
                speakerRole,
                speakerType,
                text,
                turnIndex: typeof t.turnIndex === "number" ? t.turnIndex : idx,
              };
            })
            .filter((t): t is ConversationTurn => t !== null)
        : [];

      const VALID_CONV_STYLES = ["chat-ui", "interview", "dialog-list"] as const;
      type ConvStyle = (typeof VALID_CONV_STYLES)[number];
      const rawConvStyle = String(parsed.conversationStyle ?? "").trim();
      const conversationStyle: ConvStyle | undefined = VALID_CONV_STYLES.includes(rawConvStyle as ConvStyle)
        ? (rawConvStyle as ConvStyle)
        : undefined;

      // Fix 5: post-parse structural validation
      const uniqueRoles = new Set(conversationTurns.map((t) => t.speakerRole));
      const resolvedLayout: ValidLayout =
        conversationTurns.length >= 2
          ? "conversation"
          : conversationTurns.length === 1 && uniqueRoles.size >= 1 && visualBlocks.length < 2
            ? "conversation"
            : layoutType === "conversation" && conversationTurns.length >= 1
              ? "conversation"
              : layoutType === "diagram" && visualBlocks.length >= 2
                ? "diagram"
                : layoutType === "multi-column" && columns.length >= 2
                  ? "multi-column"
                  : tableRows.length > 0 && layoutType === "table"
                    ? "table"
                    : layoutType === "title"
                      ? "title"
                      : "bullets";

      // Infer conversationStyle when not explicitly returned
      const resolvedConvStyle: ConvStyle | undefined =
        resolvedLayout === "conversation"
          ? (conversationStyle ?? "chat-ui")
          : undefined;

      return {
        slideTitle: String(parsed.slideTitle ?? `Slide ${pageIndex + 1}`),
        bullets: Array.isArray(parsed.bullets)
          ? parsed.bullets.map((bullet: unknown) => String(bullet)).filter(Boolean)
          : [],
        layoutType: resolvedLayout,
        tableRows,
        columns,
        visualBlocks,
        connectors,
        conversationStyle: resolvedConvStyle,
        conversationTurns: conversationTurns.length > 0 ? conversationTurns : undefined,
      };
    } catch (error: any) {
      const status = error?.status ?? error?.response?.status;
      const message = String(error?.message ?? error);
      const isJsonParseFailure =
        error instanceof SyntaxError ||
        message.includes("Unexpected end of JSON input") ||
        message.includes("Vision returned empty content");
      const isRetriable =
        status === 429 ||
        status === 500 ||
        status === 503 ||
        isJsonParseFailure;

      if (!isRetriable || attempt === VISION_MAX_RETRIES) {
        throw error;
      }

      const delayMs = RETRY_BASE_DELAY_MS * attempt;
      console.warn(
        `[analyze-doc-vision] retry ${attempt}/${VISION_MAX_RETRIES} for page ${pageIndex + 1} after ${isJsonParseFailure ? "invalid JSON response" : `HTTP ${status ?? "unknown"}`}`
      );
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
  }

  throw new Error("Vision analysis failed");
}

async function renderPdfPages(
  pdfBuffer: Buffer,
  maxPages: number,
  onPage: (base64Image: string, pageIndex: number, totalPages: number) => Promise<void>
): Promise<number> {
  /* eslint-disable */
  const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.js");
  const { createCanvas, DOMMatrix: NapiDOMMatrix } = require("@napi-rs/canvas");
  // Node.js には DOMMatrix がないためポリフィル（RadialAxialShading 描画時に必要）
  if (typeof globalThis.DOMMatrix === "undefined") {
    (globalThis as any).DOMMatrix = NapiDOMMatrix;
  }
  // require.resolve はこのコードパスでは webpack にバンドルされてモジュールID(数値)を返すため使用不可
  // エラーログより standalone の node_modules は /node_modules/（ルート直下）に配置される
  if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
    const { existsSync } = require("node:fs");
    const workerCandidates = [
      // standalone ビルド: ルート直下 /node_modules/（最優先）
      "/node_modules/pdfjs-dist/legacy/build/pdf.worker.js",
      // 通常の Node.js 実行（ローカル dev など）
      path.join(process.cwd(), "node_modules", "pdfjs-dist", "legacy", "build", "pdf.worker.js"),
    ];
    const found = workerCandidates.find((p) => existsSync(p));
    console.log("[analyze-doc-vision] pdf.worker resolved:", found ?? "(none found, using fallback)");
    pdfjsLib.GlobalWorkerOptions.workerSrc = found ?? workerCandidates[1];
  }
  /* eslint-enable */

  const NodeCanvasFactory = {
    create(width: number, height: number) {
      const canvas = createCanvas(width, height);
      const context = canvas.getContext("2d");
      return { canvas, context };
    },
    reset(
      canvasAndContext: { canvas: any; context: any },
      width: number,
      height: number
    ) {
      canvasAndContext.canvas.width = width;
      canvasAndContext.canvas.height = height;
    },
    destroy(canvasAndContext: { canvas: any; context: any }) {
      canvasAndContext.canvas.width = 0;
      canvasAndContext.canvas.height = 0;
    },
  };

  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(pdfBuffer),
    canvasFactory: NodeCanvasFactory,
    disableWorker: true,
    disableFontFace: true,
    nativeImageDecoderSupport: "none",
  });

  const pdf = await loadingTask.promise;
  const totalPages = Math.min(pdf.numPages, maxPages);
  console.log(`[analyze-doc-vision] PDF totalPages=${totalPages} (raw=${pdf.numPages})`);

  try {
    for (let i = 1; i <= totalPages; i++) {
      console.log(`[analyze-doc-vision] rendering page ${i}/${totalPages}`);
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale: PDF_RENDER_SCALE });
      const { canvas, context } = NodeCanvasFactory.create(
        Math.floor(viewport.width),
        Math.floor(viewport.height)
      );

      try {
        await page.render({
          canvasContext: context,
          viewport,
          canvasFactory: NodeCanvasFactory,
        }).promise;

        const pngBuffer = canvas.toBuffer("image/png");
        console.log(`[analyze-doc-vision] page ${i} rendered (${pngBuffer.length} bytes), calling Vision API`);
        await onPage(pngBuffer.toString("base64"), i - 1, totalPages);
        console.log(`[analyze-doc-vision] page ${i} Vision API done`);
      } finally {
        page.cleanup();
        NodeCanvasFactory.destroy({ canvas, context });
      }
    }
  } finally {
    pdf.destroy();
  }

  return totalPages;
}

function imageBufferToBase64(buffer: Buffer): string {
  return buffer.toString("base64");
}

async function tryBlobListingFallback(
  fileUrl: string
): Promise<{ buffer: Buffer; effectiveUrl: string } | null> {
  try {
    const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME;
    const key = process.env.AZURE_STORAGE_ACCOUNT_KEY;
    if (!acc || !key) return null;

    const urlObj = new URL(fileUrl.split("?")[0]);
    const parts = urlObj.pathname.split("/").filter(Boolean);
    if (parts.length < 2) return null;

    const containerName = parts[0];
    const threadId = parts[1];
    const prefix = `${threadId}/`;

    const connStr = `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`;
    const blobServiceClient = BlobServiceClient.fromConnectionString(connStr);
    const containerClient = blobServiceClient.getContainerClient(containerName);

    const supported = [".pdf", ".png", ".jpg", ".jpeg", ".webp", ".gif"];

    for await (const blob of containerClient.listBlobsFlat({ prefix })) {
      const nameLower = blob.name.toLowerCase();
      if (supported.some((ext) => nameLower.endsWith(ext))) {
        console.log(`[analyze-doc-vision] fallback: found blob ${blob.name}`);
        const blockBlobClient = containerClient.getBlockBlobClient(blob.name);
        const buffer = await blockBlobClient.downloadToBuffer();
        return {
          buffer,
          effectiveUrl: `https://${acc}.blob.core.windows.net/${containerName}/${blob.name}`,
        };
      }
    }

    console.warn(
      `[analyze-doc-vision] fallback: no supported file under prefix ${prefix}`
    );
    return null;
  } catch (e) {
    console.error("[analyze-doc-vision] blob listing fallback error:", e);
    return null;
  }
}

async function downloadFile(
  fileUrl: string
): Promise<{ buffer: Buffer; effectiveUrl: string }> {
  const res = await fetch(fileUrl);
  if (!res.ok) {
    if (
      res.status === 404 &&
      fileUrl.includes(".blob.core.windows.net/dl-link/")
    ) {
      console.warn("[analyze-doc-vision] 404 on blob URL, trying listing fallback");
      const fallback = await tryBlobListingFallback(fileUrl);
      if (fallback) return fallback;
    }
    throw new Error(`Failed to download file: HTTP ${res.status} - ${fileUrl}`);
  }
  const ab = await res.arrayBuffer();
  return { buffer: Buffer.from(ab), effectiveUrl: fileUrl };
}

function detectMimeType(
  url: string,
  buffer: Buffer
): "pdf" | "image" | "unknown" {
  const lower = url.toLowerCase().split("?")[0];
  if (lower.endsWith(".pdf")) return "pdf";
  if (
    lower.endsWith(".png") ||
    lower.endsWith(".jpg") ||
    lower.endsWith(".jpeg") ||
    lower.endsWith(".webp") ||
    lower.endsWith(".gif")
  ) {
    return "image";
  }
  if (buffer[0] === 0x25 && buffer[1] === 0x50) return "pdf";
  if (buffer[0] === 0x89 && buffer[1] === 0x50) return "image";
  if (buffer[0] === 0xff && buffer[1] === 0xd8) return "image";
  return "unknown";
}

export type AnalyzeDocVisionRequest = {
  fileUrl: string;
  maxPages?: number;
  mode?: "faithful" | "redesign";
};

export type AnalyzeDocVisionResponse = {
  ok: boolean;
  slides?: Array<{
    title: string;
    bullets: string[];
    layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram" | "conversation";
    tableRows?: string[][];
    columns?: Array<{ header: string; bullets: string[] }>;
    visualBlocks?: VisualBlock[];
    connectors?: Connector[];
    conversationStyle?: "chat-ui" | "interview" | "dialog-list";
    conversationTurns?: ConversationTurn[];
  }>;
  totalPages?: number;
  error?: string;
};

export async function analyzeDocVision(
  fileUrl: string,
  maxPages: number = MAX_PAGES,
  mode?: "faithful" | "redesign"
): Promise<AnalyzeDocVisionResponse> {
  if (!fileUrl?.trim()) {
    return { ok: false, error: "fileUrl is required" };
  }

  console.log("[analyze-doc-vision] fileUrl =", fileUrl.substring(0, 80));

  const { buffer, effectiveUrl } = await downloadFile(fileUrl);
  const mimeType = detectMimeType(effectiveUrl, buffer);

  console.log("[analyze-doc-vision] mimeType =", mimeType);

  const slides: Array<{
    title: string;
    bullets: string[];
    layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram" | "conversation";
    tableRows?: string[][];
    columns?: Array<{ header: string; bullets: string[] }>;
    visualBlocks?: VisualBlock[];
    connectors?: Connector[];
    conversationStyle?: "chat-ui" | "interview" | "dialog-list";
    conversationTurns?: ConversationTurn[];
  }> = [];
  let totalPages = 0;

  if (mimeType === "pdf") {
    totalPages = await renderPdfPages(
      buffer,
      maxPages,
      async (base64Image, pageIndex, pageCount) => {
        const result = await analyzePageWithVision(base64Image, pageIndex, pageCount, mode);
        slides.push({
          title: result.slideTitle,
          bullets: result.bullets,
          layoutType: result.layoutType,
          tableRows: result.tableRows,
          columns: result.columns,
          visualBlocks: result.visualBlocks,
          connectors: result.connectors,
          conversationStyle: result.conversationStyle,
          conversationTurns: result.conversationTurns,
        });
      }
    );
    console.log("[analyze-doc-vision] PDF pages analyzed:", totalPages);
  } else if (mimeType === "image") {
    totalPages = 1;
    const result = await analyzePageWithVision(imageBufferToBase64(buffer), 0, 1, mode);
    slides.push({
      title: result.slideTitle,
      bullets: result.bullets,
      layoutType: result.layoutType,
      tableRows: result.tableRows,
      columns: result.columns,
      visualBlocks: result.visualBlocks,
      connectors: result.connectors,
      conversationStyle: result.conversationStyle,
      conversationTurns: result.conversationTurns,
    });
    console.log("[analyze-doc-vision] Single image analyzed");
  } else {
    return { ok: false, error: "Unsupported file type. Supported: PDF, PNG, JPG, WEBP, GIF" };
  }

  return { ok: true, slides, totalPages };
}

export async function POST(
  req: NextRequest
): Promise<NextResponse<AnalyzeDocVisionResponse>> {
  try {
    const body: AnalyzeDocVisionRequest = await req.json();
    const { fileUrl, maxPages = MAX_PAGES, mode } = body;
    const result = await analyzeDocVision(fileUrl, maxPages, mode);
    return NextResponse.json(result, { status: result.ok ? 200 : (result.error === "fileUrl is required" ? 400 : 500) });
  } catch (e: any) {
    console.error("[analyze-doc-vision] error:", e);
    return NextResponse.json(
      { ok: false, error: String(e?.message ?? e) },
      { status: 500 }
    );
  }
}
