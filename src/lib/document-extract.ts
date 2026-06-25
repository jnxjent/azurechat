// src/lib/document-extract.ts
// Shared text-extraction utilities used by sl-sync.

import { DocumentIntelligenceInstance } from "@/features/common/services/document-intelligence";

const CHUNK_SIZE = 2300;
const CHUNK_OVERLAP = Math.floor(CHUNK_SIZE * 0.25);

export async function extractExcelText(buffer: ArrayBuffer): Promise<string[]> {
  const XLSX = require("xlsx");
  const workbook = XLSX.read(Buffer.from(buffer), {
    type: "buffer",
    cellFormula: false,
    cellHTML: false,
    cellNF: false,
    sheetStubs: false,
  });

  const docs: string[] = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows: unknown[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });
    if (!rows.length) continue;

    const lines: string[] = [`=== シート: ${sheetName} ===`];
    for (const row of rows) {
      const cells = row.map((cell) => {
        if (cell === null || cell === undefined) return "";
        if (cell instanceof Date) return cell.toLocaleDateString("ja-JP");
        return String(cell).trim();
      });
      if (cells.every((c) => c === "")) continue;
      lines.push(cells.join(" | "));
    }
    if (lines.length > 1) docs.push(lines.join("\n"));
  }
  return docs;
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"');
}

export async function extractWordText(buffer: ArrayBuffer): Promise<string[]> {
  try {
    const JSZipModule = await import("jszip");
    const JSZip = JSZipModule.default ?? JSZipModule;
    const zip = await (JSZip as any).loadAsync(Buffer.from(new Uint8Array(buffer)));
    const docXml = await zip.files["word/document.xml"]?.async("string");
    if (!docXml) return [];

    const paragraphs: string[] = [];
    const paraRe = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
    let pm: RegExpExecArray | null;
    while ((pm = paraRe.exec(docXml)) !== null) {
      const paraXml = pm[0];
      const textRe = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
      let text = "";
      let tm: RegExpExecArray | null;
      while ((tm = textRe.exec(paraXml)) !== null) {
        text += decodeXmlEntities(tm[1]);
      }
      if (text.trim()) paragraphs.push(text.trim());
    }
    return paragraphs;
  } catch {
    return [];
  }
}

const _parsed = parseInt(process.env.DOC_INTELLIGENCE_PAGE_CHUNK ?? "60", 10);
const DOC_INTELLIGENCE_PAGE_CHUNK =
  Number.isFinite(_parsed) && _parsed > 0 ? _parsed : 60;

async function getPdfPageCount(buffer: ArrayBuffer): Promise<number> {
  try {
    const pdfjsLib = await import("pdfjs-dist/legacy/build/pdf.js");
    const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(buffer) });
    const pdf = await loadingTask.promise;
    return pdf.numPages;
  } catch {
    return 0;
  }
}

export async function extractWithDocumentIntelligence(
  buffer: ArrayBuffer
): Promise<string[]> {
  const totalPages = await getPdfPageCount(buffer);
  if (totalPages === 0) {
    console.warn("[doc-extract] Could not determine page count, falling back to single call");
    const client = DocumentIntelligenceInstance();
    const poller = await client.beginAnalyzeDocument("prebuilt-read", buffer);
    const { paragraphs } = await poller.pollUntilDone();
    return (paragraphs ?? []).map((p) => p.content).filter(Boolean);
  }

  console.log(`[doc-extract] totalPages=${totalPages} chunkSize=${DOC_INTELLIGENCE_PAGE_CHUNK}`);
  const client = DocumentIntelligenceInstance();
  const allParagraphs: string[] = [];

  for (let pageStart = 1; pageStart <= totalPages; pageStart += DOC_INTELLIGENCE_PAGE_CHUNK) {
    const pageEnd = Math.min(pageStart + DOC_INTELLIGENCE_PAGE_CHUNK - 1, totalPages);
    const pages = `${pageStart}-${pageEnd}`;
    console.log(`[doc-extract] beginAnalyzeDocument pages=${pages}`);

    const poller = await client.beginAnalyzeDocument("prebuilt-read", buffer, { pages });
    const result = await poller.pollUntilDone();
    allParagraphs.push(...(result.paragraphs ?? []).map((p) => p.content).filter(Boolean));
  }

  console.log(`[doc-extract] total paragraphs extracted: ${allParagraphs.length}`);
  return allParagraphs;
}

export async function extractMsgText(buffer: ArrayBuffer): Promise<string[]> {
  try {
    const MsgReaderModule = require("@kenjiuno/msgreader");
    const MsgReader = MsgReaderModule.default ?? MsgReaderModule;
    const reader = new MsgReader(Buffer.from(buffer));
    const data = reader.getFileData();

    const lines: string[] = [];
    if (data.subject) lines.push(`件名: ${data.subject}`);
    if (data.senderName || data.senderEmail) {
      lines.push(`送信者: ${[data.senderName, data.senderEmail].filter(Boolean).join(" ")}`);
    }
    if (Array.isArray(data.recipients) && data.recipients.length > 0) {
      const toList = data.recipients
        .map((r: any) => [r.name, r.email].filter(Boolean).join(" "))
        .join(", ");
      lines.push(`宛先: ${toList}`);
    }
    if (data.messageDeliveryTime) lines.push(`日時: ${data.messageDeliveryTime}`);

    let body: string = data.body ?? "";
    if (!body && data.bodyHtml) {
      body = data.bodyHtml.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
    }
    if (body) lines.push(body);

    const fullText = lines.join("\n").trim();
    console.log(`[doc-extract] .msg extracted: subject="${data.subject ?? ""}" bodyLen=${body.length}`);
    return fullText ? [fullText] : [];
  } catch (e) {
    console.warn("[doc-extract] .msg parse failed:", e);
    return [];
  }
}

export async function extractTextFromBuffer(
  buffer: ArrayBuffer,
  fileName: string
): Promise<string[]> {
  const lower = fileName.toLowerCase();
  if (lower.endsWith(".xlsx") || lower.endsWith(".xlsm")) {
    return extractExcelText(buffer);
  }
  if (lower.endsWith(".msg")) {
    return extractMsgText(buffer);
  }
  return extractWithDocumentIntelligence(buffer);
}

export function chunkWithOverlap(text: string): string[] {
  if (text.length <= CHUNK_SIZE) return [text];
  const chunks: string[] = [];
  let start = 0;
  while (start < text.length) {
    chunks.push(text.substring(start, Math.min(start + CHUNK_SIZE, text.length)));
    start += CHUNK_SIZE - CHUNK_OVERLAP;
  }
  return chunks;
}
