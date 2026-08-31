// src/lib/document-extract.ts
// Shared text-extraction utilities used by sl-sync.

import { DocumentIntelligenceInstance } from "@/features/common/services/document-intelligence";

const CHUNK_SIZE = 2300;
const CHUNK_OVERLAP = Math.floor(CHUNK_SIZE * 0.25);
const EMPTY_PDF_PAGE_TEXT = "[PAGE_EMPTY]";

export type ExtractedIndexChunk = {
  content: string;
  chunkIndex: number;
  pageStart: number | null;
  pageEnd: number | null;
};

export type ExtractedIndexDocument = {
  chunks: ExtractedIndexChunk[];
  pageCount: number | null;
  hasPageMetadata: boolean;
};

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

function cleanExtractedText(value: string): string {
  return value
    .replace(/\u0000/g, "")
    .replace(/\r\n?/g, "\n")
    .replace(/[\t ]+\n/g, "\n")
    .trim();
}

/**
 * Decode a plain-text upload locally. Japanese business documents are
 * commonly UTF-8, UTF-16, or CP932/Shift-JIS.
 */
export function extractPlainText(buffer: ArrayBuffer): string[] {
  const bytes = new Uint8Array(buffer);
  if (bytes.length === 0) return [];

  let text = "";
  if (
    bytes.length >= 3 &&
    bytes[0] === 0xef &&
    bytes[1] === 0xbb &&
    bytes[2] === 0xbf
  ) {
    text = new TextDecoder("utf-8").decode(bytes.subarray(3));
  } else if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    text = new TextDecoder("utf-16le").decode(bytes.subarray(2));
  } else if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    // TextDecoder does not consistently expose utf-16be in every Node build.
    const swapped = new Uint8Array(bytes.length - 2);
    for (let i = 2; i + 1 < bytes.length; i += 2) {
      swapped[i - 2] = bytes[i + 1];
      swapped[i - 1] = bytes[i];
    }
    text = new TextDecoder("utf-16le").decode(swapped);
  } else {
    try {
      text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    } catch {
      // WHATWG Encoding maps shift_jis to Windows-31J/CP932 in Node.js.
      text = new TextDecoder("shift_jis").decode(bytes);
    }
  }

  const cleaned = cleanExtractedText(text);
  return cleaned ? [cleaned] : [];
}

/** Extract text from the legacy OLE-based Word .doc format. */
export async function extractLegacyWordText(
  buffer: ArrayBuffer
): Promise<string[]> {
  try {
    const WordExtractor = require("word-extractor");
    const extractor = new WordExtractor();
    const document = await extractor.extract(Buffer.from(buffer));
    const sections = [
      document.getHeaders?.(),
      document.getBody?.(),
      document.getTextboxes?.(),
      document.getFootnotes?.(),
      document.getEndnotes?.(),
    ]
      .map((part: unknown) => cleanExtractedText(String(part ?? "")))
      .filter(Boolean);
    const text = sections.join("\n");
    return text ? [text] : [];
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    throw new Error(
      `旧形式Word（.doc）を読み取れませんでした。パスワード保護、破損、または未対応の旧形式である可能性があります。${detail ? ` (${detail})` : ""}`
    );
  }
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

function normalizeReportedPageNumber(
  reportedPage: number,
  requestedStart: number,
  requestedEnd: number
): number {
  if (reportedPage >= requestedStart && reportedPage <= requestedEnd) {
    return reportedPage;
  }
  const requestedCount = requestedEnd - requestedStart + 1;
  if (reportedPage >= 1 && reportedPage <= requestedCount) {
    return requestedStart + reportedPage - 1;
  }
  return reportedPage;
}

async function extractPdfPagesWithDocumentIntelligence(
  buffer: ArrayBuffer
): Promise<{ pageCount: number; pages: Array<{ pageNumber: number; content: string }> }> {
  const totalPages = await getPdfPageCount(buffer);
  if (totalPages === 0) {
    throw new Error("PDF page count could not be determined");
  }

  console.log(`[doc-extract] totalPages=${totalPages} chunkSize=${DOC_INTELLIGENCE_PAGE_CHUNK}`);
  const client = DocumentIntelligenceInstance();
  const pageText = new Map<number, string[]>();

  for (let pageStart = 1; pageStart <= totalPages; pageStart += DOC_INTELLIGENCE_PAGE_CHUNK) {
    const pageEnd = Math.min(pageStart + DOC_INTELLIGENCE_PAGE_CHUNK - 1, totalPages);
    const pages = `${pageStart}-${pageEnd}`;
    console.log(`[doc-extract] beginAnalyzeDocument pages=${pages}`);

    const poller = await client.beginAnalyzeDocument("prebuilt-read", buffer, { pages });
    const result = await poller.pollUntilDone();
    const pagesWithLines = (result.pages ?? []).filter((page) => (page.lines?.length ?? 0) > 0);

    if (pagesWithLines.length > 0) {
      for (const page of pagesWithLines) {
        const pageNumber = normalizeReportedPageNumber(page.pageNumber, pageStart, pageEnd);
        if (pageNumber < 1 || pageNumber > totalPages) continue;
        const lines = (page.lines ?? []).map((line) => line.content).filter(Boolean);
        const current = pageText.get(pageNumber) ?? [];
        current.push(...lines);
        pageText.set(pageNumber, current);
      }
      continue;
    }

    // Defensive fallback for service responses that omit page lines.
    for (const paragraph of result.paragraphs ?? []) {
      const reportedPage = paragraph.boundingRegions?.[0]?.pageNumber;
      if (!reportedPage || !paragraph.content) continue;
      const pageNumber = normalizeReportedPageNumber(reportedPage, pageStart, pageEnd);
      if (pageNumber < 1 || pageNumber > totalPages) continue;
      const current = pageText.get(pageNumber) ?? [];
      current.push(paragraph.content);
      pageText.set(pageNumber, current);
    }
  }

  const extractedPages = Array.from({ length: totalPages }, (_, index) => {
    const pageNumber = index + 1;
    const content = (pageText.get(pageNumber) ?? []).join("\n").trim();
    return {
      pageNumber,
      content: content || EMPTY_PDF_PAGE_TEXT,
    };
  });

  console.log(
    `[doc-extract] PDF pages extracted: total=${totalPages} empty=${extractedPages.filter((page) => page.content === EMPTY_PDF_PAGE_TEXT).length}`
  );
  return { pageCount: totalPages, pages: extractedPages };
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
  if (
    lower.endsWith(".xlsx") ||
    lower.endsWith(".xlsm") ||
    lower.endsWith(".xls")
  ) {
    return extractExcelText(buffer);
  }
  if (lower.endsWith(".txt")) {
    return extractPlainText(buffer);
  }
  if (lower.endsWith(".doc")) {
    return extractLegacyWordText(buffer);
  }
  if (lower.endsWith(".msg")) {
    return extractMsgText(buffer);
  }
  return extractWithDocumentIntelligence(buffer);
}

export async function extractIndexDocumentFromBuffer(
  buffer: ArrayBuffer,
  fileName: string
): Promise<ExtractedIndexDocument> {
  if (fileName.toLowerCase().endsWith(".pdf")) {
    const extracted = await extractPdfPagesWithDocumentIntelligence(buffer);
    const chunks: ExtractedIndexChunk[] = [];

    for (const page of extracted.pages) {
      for (const content of chunkWithOverlap(page.content)) {
        chunks.push({
          content,
          chunkIndex: chunks.length,
          pageStart: page.pageNumber,
          pageEnd: page.pageNumber,
        });
      }
    }

    return {
      chunks,
      pageCount: extracted.pageCount,
      hasPageMetadata: true,
    };
  }

  const textParts = await extractTextFromBuffer(buffer, fileName);
  const chunks = chunkWithOverlap(textParts.join("\n")).map((content, chunkIndex) => ({
    content,
    chunkIndex,
    pageStart: null,
    pageEnd: null,
  }));
  return { chunks, pageCount: null, hasPageMetadata: false };
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
