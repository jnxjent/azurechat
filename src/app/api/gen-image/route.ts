// File: src/app/api/gen-image/route.ts
// AzureChat GPT5 画像生成・文字入れ統合ルート（完全修正版）
// - ベース画像生成（Azure OpenAI Images）
// - 既存画像に日本語テキスト追加（SVG + sharp）
// - Blob Storage 読込対応
// - プラカード自動検出対応
// - 豆腐文字完全排除（file:/// font 読込）
// ------------------------------------------------------

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

import { NextRequest } from "next/server";
import sharp from "sharp";
import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import {
  BlobServiceClient,
  StorageSharedKeyCredential,
} from "@azure/storage-blob";

// -------------------- ENV --------------------------
const AZ_ENDPOINT = process.env.AZURE_OPENAI_ENDPOINT!;
const AZ_KEY = process.env.AZURE_OPENAI_API_KEY!;
const DEPLOYMENT = process.env.AZURE_OPENAI_IMAGE_DEPLOYMENT!;
const API_VERSION =
  process.env.AZURE_OPENAI_API_VERSION || "2024-12-01-preview";

const STORAGE_ACCOUNT_NAME = process.env.AZURE_STORAGE_ACCOUNT_NAME;
const STORAGE_ACCOUNT_KEY = process.env.AZURE_STORAGE_ACCOUNT_KEY;
const STORAGE_CONTAINER_NAME =
  process.env.AZURE_STORAGE_CONTAINER_NAME || "images";

// -------------------- Utils ------------------------
function normalizeSpaces(input: string) {
  return String(input || "")
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function sanitizePrompt(raw: string) {
  let s = normalizeSpaces(raw);

  s = s
    .replace(/プラカード/g, "無地のボード")
    .replace(/サインボード|サイン・?ボード/g, "無地のボード")
    .replace(/メッセージボード/g, "無地のボード")
    .replace(/横断幕|バナー/g, "無地の布")
    .replace(/ポスター/g, "無地のフレーム")
    .replace(/持って|掲げ(る|て)/g, "そばにある")
    .replace(/スローガン|抗議|デモ|プロテスト|政治|選挙/g, "ファミリー向け");

  if (!/文字は入れない/.test(s)) s += "。文字は入れない。";
  if (!/非政治的/.test(s)) s += " 非政治的。";
  if (!/家族向け|ファミリー向け/.test(s)) s += " 家族向け。";
  if (!/ロゴや商標は含まない/.test(s)) s += " ロゴや商標は含まない。";

  return s;
}

function fallbackPrompt() {
  return "可愛い三毛猫のイラスト。柔らかな水彩で、背景はシンプル。文字は入れない。非政治的。家族向け。ロゴや商標は含まない。";
}

function escapeXml(s: string) {
  return String(s).replace(/[&<>"']/g, (ch) =>
    ch === "&"
      ? "&amp;"
      : ch === "<"
      ? "&lt;"
      : ch === ">"
      ? "&gt;"
      : ch === '"'
      ? "&quot;"
      : "&#39;"
  );
}

function pickNumber(v: any, def: number) {
  const n = Number(v);
  return Number.isFinite(n) ? n : def;
}

function decodeB64OrEmpty(b64: any): string {
  if (b64 == null) return "";
  try {
    return Buffer.from(String(b64), "base64").toString("utf8");
  } catch (e) {
    console.error("textB64 decode failed:", e);
    return "";
  }
}

function pickAlign(v: any): "left" | "center" | "right" {
  return v === "left" || v === "center" || v === "right" ? v : "center";
}

function pickVAlign(v: any): "top" | "middle" | "bottom" {
  return v === "top" || v === "middle" ? v : "bottom";
}

// -------------------- Placard Detection -------------
async function detectWhiteRectangle(
  imageBuffer: Buffer
): Promise<{ x: number; y: number; w: number; h: number } | null> {
  try {
    const image = sharp(imageBuffer);
    const metadata = await image.metadata();
    const { width, height } = metadata;

    if (!width || !height) return null;

    const maxDim = 512;
    const scale = Math.min(1, maxDim / Math.max(width, height));
    const resizedWidth = Math.round(width * scale);
    const resizedHeight = Math.round(height * scale);

    const { data, info } = await image
      .resize(resizedWidth, resizedHeight)
      .ensureAlpha()
      .raw()
      .toBuffer({ resolveWithObject: true });

    const w = info.width;
    const h = info.height;
    const threshold = 200;

    let minX = w,
      maxX = 0,
      minY = h,
      maxY = 0,
      white = 0;

    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const idx = (y * w + x) * 4;
        const r = data[idx],
          g = data[idx + 1],
          b = data[idx + 2];

        if (r > threshold && g > threshold && b > threshold) {
          white++;
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
      }
    }

    if (white < 100) return null;

    const rectW = maxX - minX;
    const rectH = maxY - minY;

    if (rectW < 50 || rectH < 30) return null;

    return {
      x: Math.round(minX / scale),
      y: Math.round(minY / scale),
      w: Math.round(rectW / scale),
      h: Math.round(rectH / scale),
    };
  } catch {
    return null;
  }
}

// -------------------- Text Compose ------------------
// ★★★ 豆腐文字 ZERO（file:/// フォント読込）
async function composeTextOnImageBase(
  baseImage: Buffer,
  opts: {
    text?: string;
    width: number;
    height: number;
    fontSize: number;
    strokeWidth: number;
    align: "left" | "center" | "right";
    vAlign: "top" | "middle" | "bottom";
    marginBottom: number;
    fill?: string;
    stroke?: string;
    autoDetectPlacard?: boolean;
  }
) {
  const {
    text = "",
    width,
    height,
    fontSize,
    strokeWidth,
    align,
    vAlign,
    marginBottom,
    fill = "#ffffff",
    stroke = "rgba(0,0,0,0.4)",
    autoDetectPlacard = false,
  } = opts;

  if (!text) return baseImage;

  // ------- ★ Azure Linux 本番で絶対に存在するパス ----------
  const absoluteFontPath =
    "file:///home/site/wwwroot/public/fonts/NotoSansJP-Regular.ttf";

  let x, y;
  let effectiveFontSize = fontSize;
  let anchor: "start" | "middle" | "end";

  if (autoDetectPlacard) {
    const rect = await detectWhiteRectangle(baseImage);
    if (rect) {
      x = rect.x + rect.w / 2;
      y = rect.y + rect.h / 2;
      anchor = "middle";

      const chars = Math.max(Array.from(text).length, 1);
      const maxFsW = Math.floor((rect.w * 0.8) / chars);
      const maxFsH = Math.floor(rect.h * 0.6);
      effectiveFontSize = Math.max(
        Math.min(fontSize, maxFsW, maxFsH),
        20
      );
    } else {
      x = align === "left" ? 40 : align === "right" ? width - 40 : width / 2;
      y =
        vAlign === "top"
          ? 40 + fontSize
          : vAlign === "middle"
          ? height / 2
          : height - marginBottom;
      anchor = align === "left" ? "start" : align === "right" ? "end" : "middle";
    }
  } else {
    x = align === "left" ? 40 : align === "right" ? width - 40 : width / 2;
    y =
      vAlign === "top"
        ? 40 + fontSize
        : vAlign === "middle"
        ? height / 2
        : height - marginBottom;
    anchor = align === "left" ? "start" : align === "right" ? "end" : "middle";
  }

  const svg = `
<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <style>
      @font-face {
        font-family: 'MyJP';
        src: url("${absoluteFontPath}") format("truetype");
      }
      .label {
        font-family: 'MyJP', sans-serif;
        font-size: ${effectiveFontSize}px;
        fill: ${fill};
        paint-order: stroke;
        stroke: ${stroke};
        stroke-width: ${strokeWidth}px;
        dominant-baseline: central;
        text-anchor: ${anchor};
      }
    </style>
  </defs>
  <text x="${x}" y="${y}" class="label">${escapeXml(text)}</text>
</svg>`.trim();

  return await sharp(baseImage)
    .composite([{ input: Buffer.from(svg), left: 0, top: 0 }])
    .png()
    .toBuffer();
}

// -------------------- Save to /public/generated -------
async function saveToPublicGenerated(buf: Buffer) {
  const id = crypto.randomUUID();
  const dir = path.join(process.cwd(), "public", "generated");
  await fs.mkdir(dir, { recursive: true });
  await fs.writeFile(path.join(dir, `${id}.png`), buf);
  return `/generated/${id}.png`;
}

// -------------------- Base Image Loader ----------------
async function getBaseImageBufferFromSource(
  req: NextRequest,
  imageUrl?: string,
  imageB64?: string
) {
  // --- base64
  if (imageB64) {
    const b64 = imageB64.replace(/^data:image\/\w+;base64,/, "");
    return Buffer.from(b64, "base64");
  }

  if (!imageUrl) throw new Error("image source not provided");

  // --- Blob SDK
  if (imageUrl.includes("/api/images")) {
    try {
      const u = new URL(imageUrl, "http://localhost");
      const t = u.searchParams.get("t");
      const img = u.searchParams.get("img");

      if (t && img && STORAGE_ACCOUNT_NAME && STORAGE_ACCOUNT_KEY) {
        const credential = new StorageSharedKeyCredential(
          STORAGE_ACCOUNT_NAME,
          STORAGE_ACCOUNT_KEY
        );
        const blobService = new BlobServiceClient(
          `https://${STORAGE_ACCOUNT_NAME}.blob.core.windows.net`,
          credential
        );
        const container = blobService.getContainerClient(
          STORAGE_CONTAINER_NAME
        );

        const candidates = [img, img.includes(".") ? img : `${img}.png`];
        let lastErr = null;

        for (const cand of candidates) {
          try {
            const blobPath = `${t}/${cand}`;
            const blob = container.getBlobClient(blobPath);

            const dl = await blob.download();
            const chunks: Buffer[] = [];

            for await (const ch of dl.readableStreamBody!) {
              chunks.push(Buffer.from(ch as any));
            }

            return Buffer.concat(chunks);

          } catch (e) {
            lastErr = e;
          }
        }
        throw lastErr;
      }
    } catch {}
  }

  // --- data URL
  if (imageUrl.startsWith("data:image/")) {
    const b64 = imageUrl.replace(/^data:image\/\w+;base64,/, "");
    return Buffer.from(b64, "base64");
  }

  // --- absolute URL
  if (/^https?:\/\//i.test(imageUrl)) {
    const cookie = req.headers.get("cookie") || "";
    const headers = cookie ? { cookie } : undefined;

    const res = await fetch(imageUrl, {
      cache: "no-store",
      headers,
      redirect: "follow",
    });

    if (!res.ok) throw new Error(`fetch failed: ${res.status}`);
    const buf = Buffer.from(await res.arrayBuffer());
    return buf;
  }

  // --- relative (/generated/xxx.png)
  const rel = imageUrl.startsWith("/") ? imageUrl.slice(1) : imageUrl;
  return await fs.readFile(path.join(process.cwd(), "public", rel));
}

// -------------------- Azure Images Gen -----------------
async function generateImageWithGuards({
  prompt,
  width,
  height,
  timeoutMs,
}: {
  prompt: string;
  width: number;
  height: number;
  timeoutMs: number;
}): Promise<Buffer> {
  const url = `${AZ_ENDPOINT.replace(
    /\/+$/,
    ""
  )}/openai/deployments/${DEPLOYMENT}/images/generations?api-version=${API_VERSION}`;

  async function callOnce(p: string) {
    const controller = new AbortController();
    const tm = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json", "api-key": AZ_KEY },
        body: JSON.stringify({
          prompt: p,
          size: `${width}x${height}`,
          response_format: "b64_json",
        }),
        signal: controller.signal,
      }).finally(() => clearTimeout(tm));

      if (!res.ok) {
        const text = await res.text().catch(() => "");
        return { ok: false as const, status: res.status, text };
      }

      const json = await res.json();
      const b64 = json?.data?.[0]?.b64_json;
      if (!b64)
        return { ok: false as const, status: 502, text: "No image returned" };

      return { ok: true as const, buf: Buffer.from(b64, "base64") };
    } catch (e: any) {
      if (e?.name === "AbortError")
        throw new Error(`Image generation timeout (${timeoutMs} ms)`);
      throw e;
    }
  }

  const safe = sanitizePrompt(prompt);
  const first = await callOnce(safe);

  if (first.ok) return first.buf;

  const policy =
    first.status === 400 &&
    /policy/i.test(first.text || "");

  if (policy) {
    const fb = fallbackPrompt();
    const second = await callOnce(fb);
    if (second.ok) return second.buf;

    throw new Error(`Images API policy violation: ${second.text || ""}`);
  }

  throw new Error(
    `Images API error ${first.status}: ${first.text || "unknown"}`
  );
}

// -------------------- Main Handler ---------------------
export async function POST(req: NextRequest) {
  const started = Date.now();

  try {
    const body = await req.json().catch(() => ({} as any));

    const width = pickNumber(body.width, 1024);
    const height = pickNumber(body.height, 1024);
    const baseFontSize = pickNumber(body.fontSize, 64);
    const strokeWidth = pickNumber(body.strokeWidth, 6);
    const align = pickAlign(body.align);
    const vAlign = pickVAlign(body.vAlign);
    const marginBottom = pickNumber(
      body.bottomMargin ?? body.marginBottom,
      80
    );

    const rawColor =
      body.color != null ? normalizeSpaces(String(body.color)) : undefined;
    const fill = rawColor ?? String(body.fill ?? "#ffffff");
    const stroke = String(body.stroke ?? "rgba(0,0,0,0.4)");

    const rawSize = body.size ? String(body.size).toLowerCase() : undefined;
    const sizeMap: any = {
      small: 32,
      medium: 40,
      large: 48,
      xlarge: 64,
    };
    const sizeFont = rawSize ? sizeMap[rawSize] : undefined;
    const effectiveFontSize = sizeFont ?? baseFontSize;

    const text: string = (() => {
      const t = decodeB64OrEmpty(body.textB64);
      if (t) return t;
      return String(body.text ?? "");
    })();

    const autoDetectPlacard = body.autoDetectPlacard === true;
    const timeoutMs = pickNumber(body.timeoutMs, 90000);

    // ---- 既存画像に追記
    if (body.imageUrl || body.imageB64) {
      const baseImage = await getBaseImageBufferFromSource(
        req,
        body.imageUrl,
        body.imageB64
      );

      const out = await composeTextOnImageBase(baseImage, {
        text,
        width,
        height,
        fontSize: effectiveFontSize,
        strokeWidth,
        align,
        vAlign,
        marginBottom,
        fill,
        stroke,
        autoDetectPlacard,
      });

      const imageUrl = await saveToPublicGenerated(out);

      return new Response(JSON.stringify({ imageUrl }), {
        headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
      });
    }

    // ---- 新規生成
    const prompt = String(body.prompt ?? "");
    if (!prompt) {
      return new Response(
        JSON.stringify({
          error: "invalid_request",
          detail: "prompt または imageUrl が必要です。",
        }),
        {
          status: 400,
          headers: { "Content-Type": "application/json" },
        }
      );
    }

    const baseImage = await generateImageWithGuards({
      prompt,
      width,
      height,
      timeoutMs,
    });

    if (!text) {
      const imageUrl = await saveToPublicGenerated(baseImage);
      return new Response(JSON.stringify({ imageUrl }), {
        headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
      });
    }

    const out = await composeTextOnImageBase(baseImage, {
      text,
      width,
      height,
      fontSize: effectiveFontSize,
      strokeWidth,
      align,
      vAlign,
      marginBottom,
      fill,
      stroke,
      autoDetectPlacard,
    });

    const imageUrl = await saveToPublicGenerated(out);

    return new Response(JSON.stringify({ imageUrl }), {
      headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
    });
  } catch (err: any) {
    const elapsed = Date.now() - started;
    console.error("gen-image failed:", err?.message, `(elapsed ${elapsed}ms)`);

    return new Response(JSON.stringify({ error: err?.message || "Unknown" }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
}
