export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import JSZip from "jszip";
import { promises as fs } from "node:fs";
import { constants as fsConstants } from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import {
  BlobSASPermissions,
  BlobServiceClient,
  StorageSharedKeyCredential,
  generateBlobSASQueryParameters,
} from "@azure/storage-blob";
import { OpenAIInstance } from "@/features/common/services/openai";
import { uniqueId } from "@/features/common/util";

const execFileAsync = promisify(execFile);

type SlideSummary = {
  slideIndex: number;
  texts: string[];
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
  ]
}

Rules:
- If the user asks to change the deck color or tone, set deckEdits.accentColor to a practical 6-digit hex.
- Deck color change means accent shapes, fills, and lines. It does not mean changing body text color.
- preserveTextColors should usually be true unless the user explicitly asks to recolor text.
- Only emit text replacements when the user explicitly wants wording changed.
- slideIndex is zero-based.
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
  if (!normalizeHexColor(parsed.deckEdits.accentColor) && directAccent) {
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
    const fileName = `${threadId || uniqueId()}_edited_${uniqueId()}.pptx`;
    const downloadUrl = await uploadToBlob(outputBuffer, fileName);

    return {
      downloadUrl,
      fileName,
      changedSlides: Number(pythonResult.changedSlides ?? 0),
      totalSlides: Number(pythonResult.totalSlides ?? 0),
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

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
