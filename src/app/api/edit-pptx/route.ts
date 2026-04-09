export const runtime = "nodejs";

/**
 * POST /api/edit-pptx
 *
 * 既存 PPTX ファイルをダウンロードし、ユーザーの指示に従ってテキスト内容だけを
 * GPT で書き換えて返す。背景・画像・レイアウト・フォント等のデザインは保持する。
 *
 * Body: { fileUrl: string; instruction: string; threadId: string }
 * Response: { ok: true; downloadUrl: string; fileName: string }
 *           | { ok: false; error: string }
 */

import { NextRequest, NextResponse } from "next/server";
import JSZip from "jszip";
import {
  BlobSASPermissions,
  BlobServiceClient,
  StorageSharedKeyCredential,
  generateBlobSASQueryParameters,
} from "@azure/storage-blob";
import { OpenAIInstance } from "@/features/common/services/openai";
import { uniqueId } from "@/features/common/util";

// ── XML helpers ────────────────────────────────────────────────────────────

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"');
}

function encodeXmlEntities(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/'/g, "&apos;")
    .replace(/"/g, "&quot;");
}

/**
 * Extract shape text from a single <p:sp> block.
 * Each <a:p> paragraph becomes one line.
 * Within a paragraph, all <a:t> runs are concatenated (no separator).
 * Returns { paragraphs: string[], totalRuns: number }
 */
function extractShapeTextByParagraph(shapeXml: string): string[] {
  const paragraphs: string[] = [];
  const paraRe = /<a:p(?:\s[^>]*)?>[\s\S]*?<\/a:p>/g;
  let pm: RegExpExecArray | null;
  while ((pm = paraRe.exec(shapeXml)) !== null) {
    const paraXml = pm[0];
    // Collect all <a:t> texts within this paragraph (concatenate runs)
    const runRe = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
    let rm: RegExpExecArray | null;
    let paraText = "";
    while ((rm = runRe.exec(paraXml)) !== null) {
      paraText += decodeXmlEntities(rm[1]);
    }
    paragraphs.push(paraText);
  }
  return paragraphs;
}

/**
 * Replace text in a single <p:sp> shape XML.
 * newParagraphs[i] maps to the i-th <a:p> that HAS runs.
 * Within each paragraph:
 *   - first <a:t> gets the new text
 *   - subsequent <a:t> in same paragraph are emptied
 * Paragraphs without runs (e.g. spacing-only) are preserved unchanged.
 */
function replaceShapeText(shapeXml: string, newParagraphs: string[]): string {
  let paraWithRunsIndex = 0;

  return shapeXml.replace(
    /(<a:p(?:\s[^>]*)?>)([\s\S]*?)(<\/a:p>)/g,
    (full, open, body, close) => {
      const hasRun = /<a:r[\s>]/.test(body);
      if (!hasRun) return full; // spacing/format-only paragraph — keep as-is

      const targetText = paraWithRunsIndex < newParagraphs.length
        ? newParagraphs[paraWithRunsIndex]
        : "";
      paraWithRunsIndex++;

      // First <a:t> in this paragraph gets the full text; rest emptied
      let isFirst = true;
      const newBody = body.replace(
        /(<a:t(?:\s[^>]*)?>)([^<]*)(<\/a:t>)/g,
        (_full: string, atOpen: string, _old: string, atClose: string) => {
          if (isFirst) {
            isFirst = false;
            return `${atOpen}${encodeXmlEntities(targetText)}${atClose}`;
          }
          return `${atOpen}${atClose}`;
        }
      );
      return `${open}${newBody}${close}`;
    }
  );
}

/** Split slide XML into <p:sp>…</p:sp> shape blocks, returning them in order. */
function splitIntoShapes(slideXml: string): string[] {
  const shapes: string[] = [];
  const re = /<p:sp[\s>][\s\S]*?<\/p:sp>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(slideXml)) !== null) shapes.push(m[0]);
  return shapes;
}

/**
 * Apply revised texts back into the slide XML.
 * revisions: shapeIndex (in ALL shapes) → new paragraphs array
 */
function applyRevisionsToSlide(
  slideXml: string,
  revisions: Map<number, string[]>
): string {
  if (revisions.size === 0) return slideXml;
  let shapeIndex = 0;
  return slideXml.replace(/<p:sp[\s>][\s\S]*?<\/p:sp>/g, (shapeXml) => {
    const idx = shapeIndex++;
    if (revisions.has(idx)) {
      return replaceShapeText(shapeXml, revisions.get(idx)!);
    }
    return shapeXml;
  });
}

// ── GPT revision ───────────────────────────────────────────────────────────

type SlideTextMap = {
  slideIndex: number;
  shapes: { shapeIndex: number; text: string }[];
};

async function reviseWithGPT(
  slideMaps: SlideTextMap[],
  instruction: string
): Promise<SlideTextMap[]> {
  const openai = OpenAIInstance();

  const systemPrompt = `あなたはPowerPointスライドのテキストを編集するエキスパートです。
ユーザーの指示に従って、スライドのテキスト内容を改良してください。

必須ルール:
- 必ず {"slides": [...]} 形式のJSONを返すこと（slides キーは必須）
- 指示に無関係なスライド・テキストボックスも含め、全スライドをそのまま返すこと
- テキストボックス内の改行（\\n）は維持すること
- テキストを変更した場合のみ text フィールドを更新し、変更しない場合は元の text をそのまま返すこと`;

  const userPrompt = `以下のスライドテキスト構造に対して、次の指示を適用してください。

指示: ${instruction}

スライドテキスト:
${JSON.stringify(slideMaps, null, 2)}

{"slides": [...]} 形式で全スライドを返してください。`;

  const res = await openai.chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
    response_format: { type: "json_object" },
    max_completion_tokens: 4000,
  });

  const content = res.choices[0]?.message?.content ?? "{}";
  console.log("[edit-pptx] GPT raw response (first 300):", content.substring(0, 300));

  const parsed = JSON.parse(content);

  // Must return { slides: [...] } — find the array regardless of wrapper key
  let arr: unknown = null;
  if (Array.isArray(parsed)) {
    arr = parsed;
  } else {
    // Try common wrapper keys
    for (const key of ["slides", "result", "data", "output", "pptx"]) {
      if (Array.isArray(parsed[key])) { arr = parsed[key]; break; }
    }
    // Last resort: first array-valued key
    if (!arr) {
      for (const val of Object.values(parsed)) {
        if (Array.isArray(val)) { arr = val; break; }
      }
    }
  }

  if (!Array.isArray(arr) || arr.length === 0) {
    console.warn("[edit-pptx] GPT returned no usable array, applying no changes");
    return slideMaps;
  }

  return arr as SlideTextMap[];
}

// ── Blob helpers ───────────────────────────────────────────────────────────

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
  } catch {
    return null;
  }
}

async function downloadBlob(fileUrl: string, threadId?: string): Promise<Buffer> {
  // Try direct fetch first (SAS URL)
  const res = await fetch(fileUrl);
  if (res.ok) {
    return Buffer.from(await res.arrayBuffer());
  }

  // Recovery flow: if SAS is broken but the blob path itself is valid,
  // bypass the SAS URL and download directly with the server-side account key.
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

  // Fallback: list blobs to find actual PPTX
  if (
    (res.status === 403 || res.status === 404) &&
    blobRef &&
    (blobRef.containerName === "dl-link" || blobRef.containerName === "pptx")
  ) {
    const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME;
    const key = process.env.AZURE_STORAGE_ACCOUNT_KEY;
    if (acc && key) {
      const blobPathParts = blobRef.blobPath.split("/").filter(Boolean);
      const effectiveThreadId = threadId?.trim() || blobPathParts[0];
      if (effectiveThreadId) {
        const connStr = `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`;
        const svc = BlobServiceClient.fromConnectionString(connStr);
        const cc = svc.getContainerClient(blobRef.containerName);
        for await (const blob of cc.listBlobsFlat({ prefix: `${effectiveThreadId}/` })) {
          if (blob.name.toLowerCase().endsWith(".pptx")) {
            return await cc.getBlockBlobClient(blob.name).downloadToBuffer();
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
    { containerName, blobName: fileName, expiresOn: new Date(Date.now() + 24 * 60 * 60 * 1000), permissions: BlobSASPermissions.parse("r") },
    cred
  );
  return `${bbc.url}?${sas}`;
}

// ── Main handler ───────────────────────────────────────────────────────────

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { fileUrl, instruction, threadId } = body as {
      fileUrl: string;
      instruction: string;
      threadId: string;
    };

    if (!fileUrl?.trim() || !instruction?.trim()) {
      return NextResponse.json(
        { ok: false, error: "fileUrl and instruction are required" },
        { status: 400 }
      );
    }

    console.log("[edit-pptx] fileUrl =", fileUrl.substring(0, 80));
    console.log("[edit-pptx] instruction =", instruction.substring(0, 120));

    // 1. Download PPTX
    const pptxBuffer = await downloadBlob(fileUrl, threadId);

    // 2. Open ZIP
    const zip = await JSZip.loadAsync(pptxBuffer);

    // 3. Find all slide XMLs in order
    const slideEntries = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)/)?.[1] ?? "0", 10);
        const numB = parseInt(b.match(/slide(\d+)/)?.[1] ?? "0", 10);
        return numA - numB;
      });

    if (slideEntries.length === 0) {
      return NextResponse.json({ ok: false, error: "スライドが見つかりませんでした。PATHが正しくない可能性があります。" }, { status: 400 });
    }

    console.log(`[edit-pptx] Found ${slideEntries.length} slides`);

    // 4. Extract text per shape from each slide
    //    text = paragraphs joined by "\n" (each paragraph = all runs concatenated)
    const slideMaps: SlideTextMap[] = [];

    for (let si = 0; si < slideEntries.length; si++) {
      const entryName = slideEntries[si];
      const xml = await zip.files[entryName].async("string");
      const shapes = splitIntoShapes(xml);

      const shapeTexts: { shapeIndex: number; text: string }[] = [];
      for (let shi = 0; shi < shapes.length; shi++) {
        const paragraphs = extractShapeTextByParagraph(shapes[shi]);
        // Only include shapes that have actual text
        const combined = paragraphs.join("\n").trim();
        if (combined.length > 0) {
          // Store as "\n"-joined string for GPT readability
          shapeTexts.push({ shapeIndex: shi, text: combined });
        }
      }

      if (shapeTexts.length > 0) {
        slideMaps.push({ slideIndex: si, shapes: shapeTexts });
      }
    }

    console.log(`[edit-pptx] Extracted text from ${slideMaps.length} slides with text`);

    // 5. Call GPT for revision
    const revised = await reviseWithGPT(slideMaps, instruction);

    // 6. Build revision maps per slide
    //    value = string[] (one entry per paragraph with runs)
    const revisionsBySlide = new Map<number, Map<number, string[]>>();
    for (const slideMap of revised) {
      const shapeMap = new Map<number, string[]>();
      for (const shape of slideMap.shapes) {
        const original = slideMaps
          .find((s) => s.slideIndex === slideMap.slideIndex)
          ?.shapes.find((sh) => sh.shapeIndex === shape.shapeIndex)?.text;
        if (original !== undefined && shape.text !== original) {
          // Split back into paragraphs for replaceShapeText
          shapeMap.set(shape.shapeIndex, shape.text.split("\n"));
        }
      }
      if (shapeMap.size > 0) {
        revisionsBySlide.set(slideMap.slideIndex, shapeMap);
      }
    }

    const changedSlides = revisionsBySlide.size;
    console.log(`[edit-pptx] ${changedSlides} slides will be modified`);

    // 7. Apply revisions to slide XMLs in ZIP
    for (let si = 0; si < slideEntries.length; si++) {
      const revisions = revisionsBySlide.get(si);
      if (!revisions || revisions.size === 0) continue;

      const entryName = slideEntries[si];
      const originalXml = await zip.files[entryName].async("string");
      const revisedXml = applyRevisionsToSlide(originalXml, revisions);

      // Use STORE (uncompressed) for modified XML entries.
      // Office accepts mixed compression in a single PPTX ZIP.
      zip.file(entryName, Buffer.from(revisedXml, "utf8"), {
        compression: "STORE",
        binary: false,
      });
    }

    // 8. Re-package — use STORE globally to prevent JSZip from re-compressing
    // entries it doesn't understand (e.g. already-compressed image streams).
    // File size increases slightly but correctness is guaranteed.
    const outputBuffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "STORE",
    });

    const fileName = `${threadId ?? uniqueId()}_edited_${uniqueId()}.pptx`;
    const downloadUrl = await uploadToBlob(outputBuffer, fileName);

    return NextResponse.json({
      ok: true,
      downloadUrl,
      fileName,
      changedSlides,
      totalSlides: slideEntries.length,
    });
  } catch (e: any) {
    console.error("[edit-pptx] error:", e);
    return NextResponse.json(
      { ok: false, error: String(e?.message ?? e) },
      { status: 500 }
    );
  }
}
