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

export async function extractWithDocumentIntelligence(
  buffer: ArrayBuffer
): Promise<string[]> {
  const client = DocumentIntelligenceInstance();
  const poller = await client.beginAnalyzeDocument("prebuilt-read", buffer);
  const { paragraphs } = await poller.pollUntilDone();
  if (!paragraphs) return [];
  return paragraphs.map((p) => p.content).filter(Boolean);
}

export async function extractTextFromBuffer(
  buffer: ArrayBuffer,
  fileName: string
): Promise<string[]> {
  const lower = fileName.toLowerCase();
  if (lower.endsWith(".xlsx") || lower.endsWith(".xlsm")) {
    return extractExcelText(buffer);
  }
  if (lower.endsWith(".docx")) {
    return extractWordText(buffer);
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
