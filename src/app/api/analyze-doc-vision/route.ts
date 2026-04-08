export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import { OpenAIVisionInstance } from "@/features/common/services/openai";

const MAX_PAGES = 30;
const PDF_RENDER_SCALE = 1.0;
const VISION_DETAIL: "low" | "high" | "auto" = "low";
const VISION_MAX_COMPLETION_TOKENS = 400;
const VISION_MAX_RETRIES = 3;
const RETRY_BASE_DELAY_MS = 1500;

type SlideAnalysis = {
  slideTitle: string;
  bullets: string[];
};

function buildVisionPrompt(pageIndex: number, totalPages: number): string {
  return `この資料の ${pageIndex + 1}/${totalPages} ページを解析してください。以下のJSON形式だけを返してください。
{
  "slideTitle": "ページのメインタイトル",
  "bullets": [
    "重要なポイントや本文の要約",
    "グラフや表の内容・数値・傾向",
    "画像や図の説明"
  ]
}

ルール:
- slideTitle はページの主題を簡潔に表す
- bullets にはページの重要情報を箇条書きで入れる
- グラフがあれば数値や傾向を含める
- 表があれば主要な行列内容を含める
- 図や画像があれば意味を説明する
- JSON以外の文字は出力しない`;
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
        temperature: 0.2,
      });

      const text = res.choices[0]?.message?.content ?? "{}";
      const parsed = JSON.parse(text);
      return {
        slideTitle: String(parsed.slideTitle ?? `スライド${pageIndex + 1}`),
        bullets: Array.isArray(parsed.bullets)
          ? parsed.bullets.map(String).filter(Boolean)
          : [],
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

async function downloadFile(fileUrl: string): Promise<Buffer> {
  const res = await fetch(fileUrl);
  if (!res.ok) {
    throw new Error(`Failed to download file: HTTP ${res.status} - ${fileUrl}`);
  }
  const ab = await res.arrayBuffer();
  return Buffer.from(ab);
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
  slides?: Array<{ title: string; bullets: string[] }>;
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

    const buffer = await downloadFile(fileUrl);
    const mimeType = detectMimeType(fileUrl, buffer);

    console.log("[analyze-doc-vision] mimeType =", mimeType);

    const slides: Array<{ title: string; bullets: string[] }> = [];
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
          slides.push({ title: result.slideTitle, bullets: result.bullets });
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
      slides.push({ title: result.slideTitle, bullets: result.bullets });
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
