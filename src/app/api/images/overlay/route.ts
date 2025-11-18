// src/app/api/images/overlay/route.ts
// 日本語フォント(TTF)をSVGに埋め込み → sharpで合成 → Blob保存（SAS不要・キャッシュ抑止）
export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import sharp from "sharp";
import { BlobServiceClient, StorageSharedKeyCredential } from "@azure/storage-blob";
import fs from "node:fs";
import path from "node:path";

//=================== Storage ===================//
const accountName = process.env.AZURE_STORAGE_ACCOUNT_NAME!;
const accountKey  = process.env.AZURE_STORAGE_ACCOUNT_KEY!;
const containerName = process.env.AZURE_STORAGE_CONTAINER_NAME || "images";

//=================== Font ===================//
const FONT_FILE   = "NotoSansJP-Regular.ttf";
const FONT_WEIGHT = 400;

function loadFontDataUrl(): { dataUrl: string | null; reason?: string } {
  try {
    const fontPath = path.join(process.cwd(), "public", "fonts", FONT_FILE);
    if (!fs.existsSync(fontPath)) return { dataUrl: null, reason: `Font file not found at ${fontPath}` };
    const buf = fs.readFileSync(fontPath);
    return { dataUrl: `data:font/truetype;base64,${buf.toString("base64")}` };
  } catch (e: any) {
    return { dataUrl: null, reason: e?.message ?? String(e) };
  }
}

function escXml(s: string) {
  return (s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&apos;");
}

//=================== Smart JSON decode ===================//
// リクエストの生バイトを読み、UTF-8優先でJSONパース。
// 必要に応じて Shift_JIS / EUC-JP / ISO-2022-JP にフォールバック。
async function readJsonSmart(req: NextRequest): Promise<any> {
  const raw = Buffer.from(await req.arrayBuffer());
  const ct = req.headers.get("content-type") || "";
  const m  = /charset=([^;]+)/i.exec(ct);
  const hinted = (m?.[1] || "").toLowerCase();

  const candidates = [
    hinted,
    "utf-8",
    "shift_jis",      // Windows-31J/CP932互換
    "euc-jp",
    "iso-2022-jp",
  ].filter(Boolean) as string[];

  for (const cs of candidates) {
    try {
      // TextDecoder は Node.js 18+ で各エンコーディングに対応
      const txt = new TextDecoder(cs as any, { fatal: false }).decode(raw);
      return JSON.parse(txt);
    } catch {
      // 次の候補へ
    }
  }
  // 最後にUTF-8でのbest-effortを試みる（JSON壊れなら例外）
  const txt = new TextDecoder("utf-8").decode(raw);
  return JSON.parse(txt);
}

//=================== SVG builder ===================//
/**
 * targetRect が渡された場合：
 *  - 矩形中央に配置
 *  - 矩形サイズに収まるようフォントサイズ自動調整
 *  - debugRect=true で水色ガイド表示
 */
function buildSvg(
  width: number,
  height: number,
  text: string,
  fontSize = 72,
  bottomMargin = 80,
  fontDataUrl: string,
  targetRect?: { x:number, y:number, w:number, h:number },
  debugRect: boolean = false
) {
  const family = "NotoSansJPEmbed";
  const fontFace = `@font-face{font-family:'${family}';src:url('${fontDataUrl}') format('truetype');font-weight:${FONT_WEIGHT};font-style:normal;}`;
  const safe = escXml(text);

  // 既定：下部中央
  let x = width / 2;
  let y = height - bottomMargin;
  let fs = fontSize;
  let rectSvg = "";

  // targetRect があれば中央配置＆自動リサイズ
  if (targetRect) {
    const r = targetRect;
    const chars = Math.max(Array.from(safe).length, 1);
    const maxFsByWidth  = Math.floor((r.w * 0.9) / chars); // ざっくり1文字=フォントサイズ想定
    const maxFsByHeight = Math.floor(r.h * 0.8);
    fs = Math.min(fontSize, maxFsByWidth, maxFsByHeight);

    x = r.x + r.w / 2;
    y = r.y + r.h / 2; // 縦中央
    if (debugRect) {
      rectSvg = `<rect x="${r.x}" y="${r.y}" width="${r.w}" height="${r.h}" fill="none" stroke="#00BFFF" stroke-width="3" />`;
    }
  }

  return Buffer.from(
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<svg width="${width}" height="${height}" xmlns="http://www.w3.org/2000/svg">
      <style>
        ${fontFace}
        .t{
          font:${fs}px '${family}', sans-serif;
          font-weight:${FONT_WEIGHT};
          fill:#fff;
          paint-order:stroke;
          stroke:#000;
          stroke-width:6px;
          text-rendering:geometricPrecision;
        }
      </style>
      ${rectSvg}
      <text x="${x}" y="${y}" text-anchor="middle" dominant-baseline="middle" class="t">${safe}</text>
    </svg>`
  );
}

//=================== Handlers ===================//
export async function GET() { return new Response("ok"); }

export async function POST(req: NextRequest) {
  const startedAt = Date.now();
  try {
    // ★ 生バイト→スマートJSONデコード（UTF-8優先＋各種日本語エンコーディングfallback）
    let {
      sourceBlobName, text, textB64,
      fontSize = 72, bottomMargin = 80,
      resultBlobName, overwrite = false,
      targetRect, debugRect = false
    } = await readJsonSmart(req);

    if (!accountName || !accountKey) {
      return NextResponse.json({ error: "Storage env missing" }, { status: 500 });
    }
    if (!sourceBlobName) {
      return NextResponse.json({ error: "sourceBlobName is required" }, { status: 400 });
    }

    // 文字エンコード安全化：textB64 が来ていたら最優先
    if (!text && textB64) {
      try { text = Buffer.from(String(textB64), "base64").toString("utf8"); } catch {}
    }
    if (!text) {
      return NextResponse.json({ error: "text or textB64 is required" }, { status: 400 });
    }
    text = String(text).normalize("NFC");

    // フォント読み込み
    const { dataUrl, reason } = loadFontDataUrl();
    if (!dataUrl) {
      return NextResponse.json(
        { error: "Font not embedded. Place public/fonts/NotoSansJP-Regular.ttf", detail: reason, fontEmbedded: false },
        { status: 500 }
      );
    }

    // Blob取得
    const cred = new StorageSharedKeyCredential(accountName, accountKey);
    const svc  = new BlobServiceClient(`https://${accountName}.blob.core.windows.net`, cred);
    const cont = svc.getContainerClient(containerName);
    const src  = cont.getBlockBlobClient(sourceBlobName);
    if (!(await src.exists())) {
      return NextResponse.json({ error: `Blob not found: ${sourceBlobName}` }, { status: 404 });
    }
    const baseBuf = await src.downloadToBuffer();

    // 合成
    const meta   = await sharp(baseBuf).metadata();
    const width  = meta.width  ?? 1024;
    const height = meta.height ?? 1024;
    const svg    = buildSvg(width, height, text, fontSize, bottomMargin, dataUrl, targetRect, debugRect);

    const outBuf = await sharp(baseBuf)
      .composite([{ input: svg, top: 0, left: 0 }])
      .png()
      .toBuffer();

    // 保存（キャッシュ抑止）
    const name = resultBlobName ?? `${sourceBlobName.replace(/\.[^.]+$/, "")}__text_${Date.now()}.png`;
    const dst  = cont.getBlockBlobClient(encodeURI(name));
    if (!overwrite && (await dst.exists())) {
      return NextResponse.json({ error: `Result exists: ${name}` }, { status: 409 });
    }

    await dst.uploadData(outBuf, {
      blobHTTPHeaders: {
        blobContentType: "image/png",
        blobCacheControl: "no-cache, no-store, must-revalidate",
      },
    });

    return NextResponse.json({
      message: "ok",
      resultBlobName: name,
      blobUrl: `https://${accountName}.blob.core.windows.net/${containerName}/${encodeURIComponent(name)}`,
      fontEmbedded: true,
      tookMs: Date.now() - startedAt,
      // デバッグしたいときは受信文字を見られるようにコメントアウトを外す
      // receivedText: text,
    });
  } catch (e: any) {
    console.error("overlay error >>>", e);
    return NextResponse.json({ error: e?.message ?? String(e) }, { status: 500 });
  }
}
