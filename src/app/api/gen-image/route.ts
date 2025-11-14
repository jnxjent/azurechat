// src/app/api/gen-image/route.ts
// 1) Azure OpenAI Images でベース画像生成（安全ガード＋自動リトライ）
// 2) 既存画像(imageUrl/imageB64)に日本語テキスト後付け（textB64対応）
// 3) プラカード自動認識機能（autoDetectPlacard）
// 4) 生成結果は public/generated に保存し、{ imageUrl } を JSON で返す
export const runtime = "nodejs";
export const dynamic = "force-dynamic";

import { NextRequest } from "next/server";
import sharp from "sharp";
import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";

const AZ_ENDPOINT = process.env.AZURE_OPENAI_ENDPOINT!;
const AZ_KEY = process.env.AZURE_OPENAI_API_KEY!;
const DEPLOYMENT = process.env.AZURE_OPENAI_IMAGE_DEPLOYMENT!;
const API_VERSION = process.env.AZURE_OPENAI_API_VERSION || "2024-12-01-preview";

// ---------- utils ----------
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
async function loadFontAsDataURL() {
  const fontPath = path.join(process.cwd(), "public", "fonts", "NotoSansJP-Regular.ttf");
  const fontBuf = await fs.readFile(fontPath);
  return `data:font/ttf;base64,${fontBuf.toString("base64")}`;
}
function escapeXml(s: string) {
  return String(s).replace(/[&<>"']/g, (ch) =>
    ch === "&" ? "&amp;" : ch === "<" ? "&lt;" : ch === ">" ? "&gt;" : ch === '"' ? "&quot;" : "&#39;"
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
  return (v === "left" || v === "center" || v === "right") ? v : "center";
}
function pickVAlign(v: any): "top" | "middle" | "bottom" {
  return (v === "top" || v === "middle" || v === "bottom") ? v : "bottom";
}

// ---------- プラカード自動認識 ----------
/**
 * 画像から白い矩形領域（プラカード）を検出
 * RGB値が全て200以上のピクセルを「白」と判定し、それらを囲む矩形を返す
 * ★ メモリ効率化：配列を使わず直接min/maxを計算
 */
async function detectWhiteRectangle(imageBuffer: Buffer): Promise<{ x: number; y: number; w: number; h: number } | null> {
  try {
    const image = sharp(imageBuffer);
    const metadata = await image.metadata();
    const { width, height } = metadata;
    
    if (!width || !height) {
      console.warn('[detectWhiteRectangle] Invalid image dimensions');
      return null;
    }

    // 処理速度のため画像を縮小（最大512px）
    const maxDim = 512;
    const scale = Math.min(1, maxDim / Math.max(width, height));
    const resizedWidth = Math.round(width * scale);
    const resizedHeight = Math.round(height * scale);

    console.log(`[detectWhiteRectangle] Analyzing image: ${width}x${height} -> ${resizedWidth}x${resizedHeight}`);

    // 画像をRGBA配列に変換
    const { data, info } = await image
      .resize(resizedWidth, resizedHeight, { fit: 'inside' })
      .ensureAlpha()
      .raw()
      .toBuffer({ resolveWithObject: true });

    const w = info.width;
    const h = info.height;

    // 白色ピクセルの境界を直接計算（配列を使わない）
    const threshold = 200;
    let minX = w;
    let maxX = 0;
    let minY = h;
    let maxY = 0;
    let whitePixelCount = 0;

    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const idx = (y * w + x) * 4;
        const r = data[idx];
        const g = data[idx + 1];
        const b = data[idx + 2];
        
        // 白っぽいピクセル
        if (r > threshold && g > threshold && b > threshold) {
          whitePixelCount++;
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
      }
    }

    if (whitePixelCount < 100) {
      console.warn(`[detectWhiteRectangle] Not enough white pixels found: ${whitePixelCount}`);
      return null;
    }

    const rectW = maxX - minX;
    const rectH = maxY - minY;

    // 矩形が小さすぎる場合は無視
    if (rectW < 50 || rectH < 30) {
      console.warn(`[detectWhiteRectangle] Rectangle too small: ${rectW}x${rectH}`);
      return null;
    }

    // 元の画像サイズにスケールバック
    const result = {
      x: Math.round(minX / scale),
      y: Math.round(minY / scale),
      w: Math.round(rectW / scale),
      h: Math.round(rectH / scale)
    };

    console.log(`[detectWhiteRectangle] Detected placard: x=${result.x}, y=${result.y}, w=${result.w}, h=${result.h} (${whitePixelCount} white pixels)`);
    return result;

  } catch (e) {
    console.error('[detectWhiteRectangle] Detection failed:', e);
    return null;
  }
}

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
    autoDetectPlacard?: boolean; // ★ 新規パラメータ
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

  const fontDataURL = await loadFontAsDataURL();
  
  let x: number;
  let y: number;
  let effectiveFontSize = fontSize;
  let anchor: "start" | "middle" | "end";

  // ★ プラカード自動検出
  if (autoDetectPlacard) {
    console.log('[composeTextOnImageBase] Attempting placard detection...');
    const rect = await detectWhiteRectangle(baseImage);
    
    if (rect) {
      console.log(`[composeTextOnImageBase] Using detected placard position`);
      
      // プラカードの中央に配置
      x = rect.x + rect.w / 2;
      y = rect.y + rect.h / 2;
      anchor = "middle";
      
      // フォントサイズを自動調整（矩形に収まるように）
      const chars = Math.max(Array.from(text).length, 1);
      const maxFsByWidth = Math.floor((rect.w * 0.8) / chars);
      const maxFsByHeight = Math.floor(rect.h * 0.6);
      effectiveFontSize = Math.min(fontSize, maxFsByWidth, maxFsByHeight);
      
      // 最小フォントサイズの保証
      effectiveFontSize = Math.max(effectiveFontSize, 20);
      
      console.log(`[composeTextOnImageBase] Font size adjusted: ${fontSize} -> ${effectiveFontSize}`);
    } else {
      console.warn('[composeTextOnImageBase] Placard detection failed, using default position');
      x = align === "left" ? 40 : align === "right" ? width - 40 : width / 2;
      y = vAlign === "top" ? 40 + fontSize : vAlign === "middle" ? height / 2 : height - marginBottom;
      anchor = align === "left" ? "start" : align === "right" ? "end" : "middle";
    }
  } else {
    // 従来の配置ロジック
    x = align === "left" ? 40 : align === "right" ? width - 40 : width / 2;
    y = vAlign === "top" ? 40 + fontSize : vAlign === "middle" ? height / 2 : height - marginBottom;
    anchor = align === "left" ? "start" : align === "right" ? "end" : "middle";
  }

  const svg = `
<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <style>
      @font-face {
        font-family: 'MyJP';
        src: url('${fontDataURL}') format('truetype');
        font-weight: normal;
        font-style: normal;
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

  return await sharp(baseImage).composite([{ input: Buffer.from(svg), left: 0, top: 0 }]).png().toBuffer();
}

async function saveToPublicGenerated(buf: Buffer) {
  const id = crypto.randomUUID();
  const dir = path.join(process.cwd(), "public", "generated");
  await fs.mkdir(dir, { recursive: true });
  await fs.writeFile(path.join(dir, `${id}.png`), buf);
  return `/generated/${id}.png`;
}

// 相対URL・dataURL・絶対URLの全てに対応して Buffer を返す
async function getBaseImageBufferFromSource(req: NextRequest, imageUrl?: string, imageB64?: string) {
  if (imageB64) {
    const b64 = imageB64.replace(/^data:image\/\w+;base64,/, "");
    return Buffer.from(b64, "base64");
  }
  if (!imageUrl) throw new Error("image source not provided");

  // ★ /api/images?t=xxx&img=yyy の場合、Azure Storageから直接取得
  if (imageUrl.includes('/api/images?')) {
    try {
      const url = new URL(imageUrl, 'http://localhost');
      const threadId = url.searchParams.get('t');
      const imgName = url.searchParams.get('img');
      
      if (threadId && imgName) {
        const storageAccount = process.env.AZURE_STORAGE_ACCOUNT_NAME;
        const containerName = process.env.AZURE_STORAGE_CONTAINER_NAME || "images";
        const directUrl = `https://${storageAccount}.blob.core.windows.net/${containerName}/${threadId}/${imgName}`;
        
        console.log(`[getBaseImageBufferFromSource] Fetching from Azure Storage: ${directUrl}`);
        const res = await fetch(directUrl, { cache: "no-store" });
        if (res.ok) {
          console.log(`[getBaseImageBufferFromSource] Successfully fetched from Azure Storage`);
          return Buffer.from(await res.arrayBuffer());
        }
        console.warn(`[getBaseImageBufferFromSource] Azure Storage fetch failed (${res.status}), falling back`);
      }
    } catch (e) {
      console.error('[getBaseImageBufferFromSource] Failed to fetch from Azure Storage:', e);
    }
  }

  // data URL
  if (imageUrl.startsWith("data:image/")) {
    const b64 = imageUrl.replace(/^data:image\/\w+;base64,/, "");
    return Buffer.from(b64, "base64");
  }

  // 絶対URL
  if (/^https?:\/\//i.test(imageUrl)) {
    const res = await fetch(imageUrl, { cache: "no-store" });
    if (!res.ok) throw new Error(`fetch imageUrl failed ${res.status}: ${await res.text().catch(() => "")}`);
    return Buffer.from(await res.arrayBuffer());
  }

  // 相対パス（例: /generated/xxx.png）→ public 直読み
  const rel = imageUrl.startsWith("/") ? imageUrl.slice(1) : imageUrl;
  const abs = path.join(process.cwd(), "public", rel);
  return await fs.readFile(abs);
}

// ---------- Azure Images 呼び出し（ガード＋フォールバック） ----------
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
  const url = `${AZ_ENDPOINT.replace(/\/+$/, "")}/openai/deployments/${DEPLOYMENT}/images/generations?api-version=${API_VERSION}`;

  async function callOnce(p: string) {
    const controller = new AbortController();
    const tm = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const payload = { prompt: p, size: `${width}x${height}`, response_format: "b64_json" };
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json", "api-key": AZ_KEY },
        body: JSON.stringify(payload),
        signal: controller.signal,
      }).finally(() => clearTimeout(tm));

      if (!res.ok) {
        const t = await res.text().catch(() => "");
        return { ok: false as const, status: res.status, text: t };
      }
      const json = await res.json();
      const b64 = json?.data?.[0]?.b64_json;
      if (!b64) return { ok: false as const, status: 502, text: "No image in response" };
      return { ok: true as const, buf: Buffer.from(b64, "base64") };
    } catch (e: any) {
      if (e?.name === "AbortError") throw new Error(`Timeout while generating image (aborted after ${timeoutMs} ms)`);
      throw e;
    }
  }

  const safe = sanitizePrompt(prompt);
  const first = await callOnce(safe);
  if (first.ok) return first.buf;

  const policy = first.status === 400 && /content_policy_violation|ResponsibleAIPolicyViolation/i.test(first.text || "");
  if (policy) {
    const fb = fallbackPrompt();
    const second = await callOnce(fb);
    if (second.ok) return second.buf;
    throw new Error(`Images API rejected by policy. detail=${second.text || first.text || "policy_violation"}`);
  }
  throw new Error(`Images API error ${first.status}: ${first.text || "unknown"}`);
}

// ---------- route ----------
export async function POST(req: NextRequest) {
  const started = Date.now();
  try {
    const body = await req.json().catch(() => ({}));

    // 共通パラメータ
    const width = pickNumber(body.width, 1024);
    const height = pickNumber(body.height, 1024);
    const fontSize = pickNumber(body.fontSize, 64);
    const strokeWidth = pickNumber(body.strokeWidth, 6);
    const align = pickAlign(body.align);
    const vAlign = pickVAlign(body.vAlign);

    // ✅ 名称の正規化：bottomMargin（新）優先、互換でmarginBottom（旧）
    const marginBottom = pickNumber(body.bottomMargin ?? body.marginBottom, 80);

    const fill: string = String(body.fill ?? "#ffffff");
    const stroke: string = String(body.stroke ?? "rgba(0,0,0,0.4)");

    // ✅ 日本語対応：textB64（新）優先、互換でtext（旧）
    const text: string = (() => {
      const t = decodeB64OrEmpty(body.textB64);
      if (t) return t;
      return String(body.text ?? "");
    })();

    // ★ プラカード自動認識フラグ
    const autoDetectPlacard: boolean = body.autoDetectPlacard === true;

    const timeoutMs = pickNumber(body.timeoutMs, 90_000);

    // --- 分岐A：既存画像に追記 ---
    if (body.imageUrl || body.imageB64) {
      const baseImage = await getBaseImageBufferFromSource(req, body.imageUrl, body.imageB64);
      const out = await composeTextOnImageBase(baseImage, {
        text,
        width,
        height,
        fontSize,
        strokeWidth,
        align,
        vAlign,
        marginBottom,
        fill,
        stroke,
        autoDetectPlacard, // ★ 追加
      });
      const imageUrl = await saveToPublicGenerated(out);
      return new Response(JSON.stringify({ imageUrl }), {
        headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
      });
    }

    // --- 分岐B：新規生成 ---
    const prompt: string = String(body.prompt ?? "");
    if (!prompt) {
      return new Response(
        JSON.stringify({
          error: "invalid_request",
          detail: "Either {prompt} or {imageUrl & text} or {imageB64 & text} is required.",
        }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }

    const baseImage = await generateImageWithGuards({ prompt, width, height, timeoutMs });

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
      fontSize,
      strokeWidth,
      align,
      vAlign,
      marginBottom,
      fill,
      stroke,
      autoDetectPlacard, // ★ 追加
    });
    const imageUrl = await saveToPublicGenerated(out);
    return new Response(JSON.stringify({ imageUrl }), {
      headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
    });
  } catch (err: any) {
    const elapsed = Date.now() - started;
    const msg = err?.message || "Unknown error";
    console.error("gen-image route failed:", msg, `(elapsed ${elapsed}ms)`);
    return new Response(JSON.stringify({ error: msg }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
}