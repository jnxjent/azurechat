// File: src/app/api/gen-image/route.ts
// AzureChat GPT5 画像生成・文字入れ統合ルート（フォント指定対応版）

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
  if (v === undefined || v === null) return def;
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
  if (v === "top" || v === "middle" || v === "bottom") return v;
  return "middle";
}

function clamp(n: number, min: number, max: number) {
  return Math.max(min, Math.min(max, n));
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
async function composeTextOnImageBase(
  baseImage: Buffer,
  opts: {
    text?: string;

    // ※ body.width/height は “ヒント” 扱いにし、実画像メタを優先する
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
    fontFamily?: "gothic" | "mincho" | "meiryo";
    bold?: boolean;
    italic?: boolean;
    offsetX?: number;
    offsetY?: number;
  }
) {
  const {
    text = "",
    width: widthHint,
    height: heightHint,
    fontSize,
    strokeWidth,
    align,
    vAlign,
    marginBottom,
    fill = "#ffffff",
    stroke = "rgba(0,0,0,0.4)",
    autoDetectPlacard = false,
    fontFamily = "gothic",
    bold = false,
    italic = false,
    offsetX = 0,
    offsetY = 0,
  } = opts;

  if (!text) return baseImage;

  // ★★★★★ 重要：実画像サイズを優先（キャンバス認識ズレ対策）
  const meta = await sharp(baseImage).metadata();
  const width = meta.width ?? widthHint ?? 1024;
  const height = meta.height ?? heightHint ?? 1024;

  // --- フォントファイルの選択 ---
  const fontFileNameProd =
    fontFamily === "mincho"
      ? "/home/site/wwwroot/public/fonts/NotoSerifJP-Regular.otf"
      : "/home/site/wwwroot/public/fonts/NotoSansJP-Regular.ttf";

  const fontFileNameLocal =
    fontFamily === "mincho"
      ? path.join(process.cwd(), "public", "fonts", "NotoSerifJP-Regular.otf")
      : path.join(process.cwd(), "public", "fonts", "NotoSansJP-Regular.ttf");

  const fontFilePathProd = fontFileNameProd;
  const fontFilePathLocal = fontFileNameLocal;

  const absoluteFontPath =
    process.env.WEBSITE_SITE_NAME || process.env.NODE_ENV === "production"
      ? `file://${fontFilePathProd}`
      : `file://${fontFilePathLocal.replace(/\\/g, "/")}`;

  const cssFontFamily =
    fontFamily === "mincho"
      ? "'MyJP', 'Noto Serif JP', 'Yu Mincho', serif"
      : fontFamily === "meiryo"
      ? "'MyJP', 'Meiryo', 'Yu Gothic', 'Noto Sans JP', sans-serif"
      : "'MyJP', 'Noto Sans JP', 'Yu Gothic', sans-serif";

  let x: number;
  let y: number;
  let effectiveFontSize = fontSize;
  let anchor: "start" | "middle" | "end";

  const paddingX = 40; // 左右の余白
  const paddingTop = 40; // 上余白（中心座標計算用）
  const paddingBottomExtra = 20; // 下側の最低余白（marginBottomとは別）

  // ★ y は「文字の中心」なので top/bottom は fontSize/2 を考慮
  const topYCenter = paddingTop + effectiveFontSize / 2;
  const bottomYCenter = height - marginBottom - effectiveFontSize / 2;

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
      x = align === "left" ? paddingX : align === "right" ? width - paddingX : width / 2;
      y =
        vAlign === "top"
          ? topYCenter
          : vAlign === "middle"
          ? height / 2
          : bottomYCenter;

      anchor = align === "left" ? "start" : align === "right" ? "end" : "middle";
    }
  } else {
    x = align === "left" ? paddingX : align === "right" ? width - paddingX : width / 2;
    y =
      vAlign === "top"
        ? topYCenter
        : vAlign === "middle"
        ? height / 2
        : bottomYCenter;

    anchor = align === "left" ? "start" : align === "right" ? "end" : "middle";
  }

  // オフセット反映（「少し上」「➡で右に」など）
  x += offsetX;
  y += offsetY;

  // ★★★★★ 重要：キャンバス外に出ないよう clamp
  // x：左右 paddingX 内に保持
  x = clamp(x, paddingX, width - paddingX);

  // y：上は topYCenter 以上、下は (height - marginBottom - fontSize/2) 以下
  // ただし bottom が極端に小さい/負になる場合に備えて保険
  const yMin = topYCenter;
  const yMax = Math.max(
    yMin,
    height - paddingBottomExtra - effectiveFontSize / 2
  );
  y = clamp(y, yMin, yMax);

  const fontWeight = bold ? "700" : "400";
  const fontStyle = italic ? "italic" : "normal";

  const svg = `
<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <style>
      @font-face {
        font-family: 'MyJP';
        src: url("${absoluteFontPath}") format("truetype");
      }
      .label {
        font-family: ${cssFontFamily};
        font-size: ${effectiveFontSize}px;
        font-weight: ${fontWeight};
        font-style: ${fontStyle};
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
  if (imageB64) {
    const b64 = imageB64.replace(/^data:image\/\w+;base64,/, "");
    return Buffer.from(b64, "base64");
  }

  if (!imageUrl) throw new Error("image source not provided");

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
        let lastErr: any = null;

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
    } catch {
      // fallthrough
    }
  }

  if (imageUrl.startsWith("data:image/")) {
    const b64 = imageUrl.replace(/^data:image\/\w+;base64,/, "");
    return Buffer.from(b64, "base64");
  }

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
    first.status === 400 && /policy/i.test(first.text || "");

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

    console.log(
      "[gen-image] align/vAlign from body >>>",
      body.align,
      body.vAlign
    );

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
    const sizeMap: Record<string, number> = {
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

    const fontFamily =
      body.fontFamily as "gothic" | "mincho" | "meiryo" | undefined;
    const bold = body.bold === true;
    const italic = body.italic === true;
    const offsetX = pickNumber(body.offsetX, 0);
    const offsetY = pickNumber(body.offsetY, 0);

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
        fontFamily,
        bold,
        italic,
        offsetX,
        offsetY,
      });

      const imageUrl = await saveToPublicGenerated(out);

      return new Response(JSON.stringify({ imageUrl }), {
        headers: {
          "Content-Type": "application/json",
          "Cache-Control": "no-store",
        },
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
        headers: {
          "Content-Type": "application/json",
          "Cache-Control": "no-store",
        },
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
      fontFamily,
      bold,
      italic,
      offsetX,
      offsetY,
    });

    const imageUrl = await saveToPublicGenerated(out);

    return new Response(JSON.stringify({ imageUrl }), {
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store",
      },
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
