export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
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

type SlideAnalysis = {
  slideTitle: string;
  bullets: string[];
  layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram";
  tableRows?: string[][];
  columns?: Array<{ header: string; bullets: string[] }>;
  visualBlocks?: VisualBlock[];
  connectors?: Connector[];
};

function buildVisionPrompt(pageIndex: number, totalPages: number): string {
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
  "visualBlocks": [
    { "kind": "callout", "role": "annotation", "groupId": "g1", "text": "Label", "x": 8, "y": 20, "w": 24, "h": 16, "emphasis": true },
    { "kind": "node", "role": "primary", "groupId": "g1", "text": "Central idea", "x": 38, "y": 35, "w": 24, "h": 18 }
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
- Use "diagram" when the page is mainly boxes, speech bubbles, arrows, callouts, or a visual relationship map.
- Keep bullets concise and use "[H] ..." for section headers if helpful.
- For diagram pages, include 2-8 visualBlocks and 0-8 connectors.
- Set visualBlocks.role to "primary", "supporting", or "annotation" when possible.
- Add groupId when multiple blocks belong to the same area, lane, or subtopic.
- Set connectors.relationshipType to "flow", "compare", "annotation", or "support" when possible.
- visualBlocks coordinates use a 0-100 canvas.
- Preserve important labels and hierarchy from the page.
- Keep item text short, usually under 45 characters.
- Return valid JSON only.`;
}

async function analyzePageWithVision(
  base64Image: string,
  pageIndex: number,
  totalPages: number
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
                text: buildVisionPrompt(pageIndex, totalPages),
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
        max_completion_tokens: VISION_MAX_COMPLETION_TOKENS,
      });

      const text = res.choices[0]?.message?.content ?? "{}";
      const parsed = JSON.parse(text);

      const VALID_LAYOUTS = ["title", "bullets", "table", "multi-column", "diagram"] as const;
      type ValidLayout = (typeof VALID_LAYOUTS)[number];
      const layoutType: ValidLayout = VALID_LAYOUTS.includes(parsed.layoutType)
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
              ): c is {
                header: string;
                bullets: string[];
              } => c !== null && c.header.length > 0
            )
        : [];

      const visualBlocks = Array.isArray(parsed.visualBlocks)
        ? parsed.visualBlocks
            .map((block: unknown) => {
              if (!block || typeof block !== "object") return null;
              const b = block as Record<string, unknown>;
              const kind = String(b.kind ?? "").trim();
              if (!["callout", "node", "badge", "figure"].includes(kind)) return null;
              const textValue = String(b.text ?? "").trim();
              const x = Number(b.x);
              const y = Number(b.y);
              const w = Number(b.w);
              const h = Number(b.h);
              if (
                !textValue ||
                !Number.isFinite(x) ||
                !Number.isFinite(y) ||
                !Number.isFinite(w) ||
                !Number.isFinite(h)
              ) {
                return null;
              }
              return {
                kind: kind as VisualBlock["kind"],
                role:
                  String(b.role ?? "").trim() === "primary" ||
                  String(b.role ?? "").trim() === "supporting" ||
                  String(b.role ?? "").trim() === "annotation"
                    ? (String(b.role).trim() as VisualBlock["role"])
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
            .filter((block: VisualBlock | null): block is VisualBlock => block !== null)
        : [];

      const connectors = Array.isArray(parsed.connectors)
        ? parsed.connectors
            .map((conn: unknown) => {
              if (!conn || typeof conn !== "object") return null;
              const c = conn as Record<string, unknown>;
              const from = Number(c.from);
              const to = Number(c.to);
              if (!Number.isInteger(from) || !Number.isInteger(to)) return null;
              return {
                from,
                to,
                label: String(c.label ?? "").trim(),
                style:
                  String(c.style ?? "arrow").trim() === "line"
                    ? ("line" as const)
                    : ("arrow" as const),
                relationshipType:
                  String(c.relationshipType ?? "").trim() === "flow" ||
                  String(c.relationshipType ?? "").trim() === "compare" ||
                  String(c.relationshipType ?? "").trim() === "annotation" ||
                  String(c.relationshipType ?? "").trim() === "support"
                    ? (String(c.relationshipType).trim() as Connector["relationshipType"])
                    : undefined,
              };
            })
            .filter((conn: Connector | null): conn is Connector => conn !== null)
        : [];

      const resolvedLayout: ValidLayout =
        layoutType === "diagram" && visualBlocks.length >= 2
          ? "diagram"
          : layoutType === "multi-column" && columns.length >= 2
            ? "multi-column"
            : tableRows.length > 0 && layoutType === "table"
              ? "table"
              : layoutType === "title"
                ? "title"
                : "bullets";

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
      };
    } catch (error: any) {
      const status = error?.status ?? error?.response?.status;
      const isRetriable = status === 429 || status === 500 || status === 503;

      if (!isRetriable || attempt === VISION_MAX_RETRIES) {
        throw error;
      }

      const delayMs = RETRY_BASE_DELAY_MS * attempt;
      console.warn(
        `[analyze-doc-vision] retry ${attempt}/${VISION_MAX_RETRIES} for page ${pageIndex + 1} after HTTP ${status ?? "unknown"}`
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
  const { createCanvas } = require("@napi-rs/canvas");
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
    disableFontFace: true,
    nativeImageDecoderSupport: "none",
  });

  const pdf = await loadingTask.promise;
  const totalPages = Math.min(pdf.numPages, maxPages);

  try {
    for (let i = 1; i <= totalPages; i++) {
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
        await onPage(pngBuffer.toString("base64"), i - 1, totalPages);
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
};

export type AnalyzeDocVisionResponse = {
  ok: boolean;
  slides?: Array<{
    title: string;
    bullets: string[];
    layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram";
    tableRows?: string[][];
    columns?: Array<{ header: string; bullets: string[] }>;
    visualBlocks?: VisualBlock[];
    connectors?: Connector[];
  }>;
  totalPages?: number;
  error?: string;
};

export async function POST(
  req: NextRequest
): Promise<NextResponse<AnalyzeDocVisionResponse>> {
  try {
    const body: AnalyzeDocVisionRequest = await req.json();
    const { fileUrl, maxPages = MAX_PAGES } = body;

    if (!fileUrl?.trim()) {
      return NextResponse.json(
        { ok: false, error: "fileUrl is required" },
        { status: 400 }
      );
    }

    console.log("[analyze-doc-vision] fileUrl =", fileUrl.substring(0, 80));

    const { buffer, effectiveUrl } = await downloadFile(fileUrl);
    const mimeType = detectMimeType(effectiveUrl, buffer);

    console.log("[analyze-doc-vision] mimeType =", mimeType);

    const slides: Array<{
      title: string;
      bullets: string[];
      layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram";
      tableRows?: string[][];
      columns?: Array<{ header: string; bullets: string[] }>;
      visualBlocks?: VisualBlock[];
      connectors?: Connector[];
    }> = [];
    let totalPages = 0;

    if (mimeType === "pdf") {
      totalPages = await renderPdfPages(
        buffer,
        maxPages,
        async (base64Image, pageIndex, pageCount) => {
          const result = await analyzePageWithVision(
            base64Image,
            pageIndex,
            pageCount
          );
          slides.push({
            title: result.slideTitle,
            bullets: result.bullets,
            layoutType: result.layoutType,
            tableRows: result.tableRows,
            columns: result.columns,
            visualBlocks: result.visualBlocks,
            connectors: result.connectors,
          });
        }
      );
      console.log("[analyze-doc-vision] PDF pages analyzed:", totalPages);
    } else if (mimeType === "image") {
      totalPages = 1;
      const result = await analyzePageWithVision(
        imageBufferToBase64(buffer),
        0,
        1
      );
      slides.push({
        title: result.slideTitle,
        bullets: result.bullets,
        layoutType: result.layoutType,
        tableRows: result.tableRows,
        columns: result.columns,
        visualBlocks: result.visualBlocks,
        connectors: result.connectors,
      });
      console.log("[analyze-doc-vision] Single image analyzed");
    } else {
      return NextResponse.json(
        {
          ok: false,
          error: "Unsupported file type. Supported: PDF, PNG, JPG, WEBP, GIF",
        },
        { status: 400 }
      );
    }

    return NextResponse.json({
      ok: true,
      slides,
      totalPages,
    });
  } catch (e: any) {
    console.error("[analyze-doc-vision] error:", e);
    return NextResponse.json(
      { ok: false, error: String(e?.message ?? e) },
      { status: 500 }
    );
  }
}
