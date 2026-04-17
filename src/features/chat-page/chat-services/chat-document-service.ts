"use server";
import "server-only";

import { userHashedId } from "@/features/auth-page/helpers";
import { HistoryContainer } from "@/features/common/services/cosmos";

import { RevalidateCache } from "@/features/common/navigation-helpers";
import { ServerActionResponse } from "@/features/common/server-action-response";
import { DocumentIntelligenceInstance } from "@/features/common/services/document-intelligence";
import { uniqueId } from "@/features/common/util";
import { SqlQuerySpec } from "@azure/cosmos";
import { EnsureIndexIsCreated } from "./azure-ai-search/azure-ai-search";
import { CHAT_DOCUMENT_ATTRIBUTE, ChatDocumentModel } from "./models";
import {
  GenerateSasUrl,
  UploadBlob,
} from "@/features/common/services/azure-storage";

// ─────────────────────────────────────────────
// アップロード上限（バイト）
// デフォルトは 100MB。環境変数 MAX_UPLOAD_DOCUMENT_SIZE で上書き可能。
// ─────────────────────────────────────────────
const DEFAULT_MAX_UPLOAD_DOCUMENT_SIZE = 100 * 1024 * 1024; // 100MB

function resolveMaxUploadDocumentSize(): number {
  const raw = process.env.MAX_UPLOAD_DOCUMENT_SIZE;
  if (!raw) {
    return DEFAULT_MAX_UPLOAD_DOCUMENT_SIZE;
  }

  const parsed = Number(raw);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    // 不正値の場合はログを出してデフォルトにフォールバック
    console.warn(
      `[Upload] MAX_UPLOAD_DOCUMENT_SIZE is invalid (${raw}), using default ${DEFAULT_MAX_UPLOAD_DOCUMENT_SIZE} bytes`
    );
    return DEFAULT_MAX_UPLOAD_DOCUMENT_SIZE;
  }

  return parsed;
}

const MAX_UPLOAD_DOCUMENT_SIZE = resolveMaxUploadDocumentSize();

const CHUNK_SIZE = 2300;
// 25% overlap
const CHUNK_OVERLAP = CHUNK_SIZE * 0.25;

const DOCUMENT_CONTAINER_NAME = "dl-link";

export const UploadDocumentToStore = async (
  _threadId: string,
  fileName: string,
  fileData: Buffer
): Promise<ServerActionResponse<string>> => {
  const uploadResponse = await UploadBlob(DOCUMENT_CONTAINER_NAME, fileName, fileData);
  if (uploadResponse.status !== "OK") {
    return uploadResponse;
  }

  return await GenerateSasUrl(DOCUMENT_CONTAINER_NAME, fileName);
};

export const UploadDocument = async (formData: FormData) => {
  const threadId = String(formData.get("id"));
  const file: File | null = formData.get("file") as unknown as File;
  const fileName = formData.get("fileName") as string;
  const blob = new Blob([file], { type: file.type });
  const buff = await blob.arrayBuffer();
  const uploadResponse = await UploadDocumentToStore(
    threadId,
    `${threadId}/${fileName}`,
    Buffer.from(buff)
  );
  return uploadResponse;
};

export const CrackDocument = async (
  formData: FormData
): Promise<ServerActionResponse<string[]>> => {
  try {
    const response = await EnsureIndexIsCreated();
    if (response.status === "OK") {
      const fileResponse = await LoadFile(formData);
      if (fileResponse.status === "OK") {
        const splitDocuments = await ChunkDocumentWithOverlap(
          fileResponse.response.join("\n")
        );

        return {
          status: "OK",
          response: splitDocuments,
        };
      }

      return fileResponse;
    }

    return response;
  } catch (e) {
    return {
      status: "ERROR",
      errors: [
        {
          message: `${e}`,
        },
      ],
    };
  }
};

// ---------- Excel (.xlsx / .xlsm) テキスト抽出 ----------
async function extractExcelText(buffer: ArrayBuffer): Promise<string[]> {
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const XLSX = require("xlsx");

  const workbook = XLSX.read(Buffer.from(buffer), {
    type: "buffer",
    cellFormula: false, // 数式は読まない（値のみ）
    cellHTML: false,
    cellNF: false,
    sheetStubs: false,
  });

  const docs: string[] = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];

    // シート全体をJSON（行の配列）に変換
    const rows: unknown[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,          // 1-based配列形式（ヘッダー行も含む）
      defval: "",         // 空セルは空文字
      blankrows: false,   // 完全空行はスキップ
    });

    if (!rows.length) continue;

    const lines: string[] = [`=== シート: ${sheetName} ===`];

    for (const row of rows) {
      // undefined/null を空文字に変換し、区切り文字「|」で結合
      const cells = row.map((cell) => {
        if (cell === null || cell === undefined) return "";
        // Dateオブジェクトは日付文字列に変換
        if (cell instanceof Date) return cell.toLocaleDateString("ja-JP");
        return String(cell).trim();
      });
      // 全セルが空の行はスキップ
      if (cells.every((c) => c === "")) continue;
      lines.push(cells.join(" | "));
    }

    if (lines.length > 1) {
      docs.push(lines.join("\n"));
    }
  }

  return docs;
}

// Excel拡張子判定
function isExcelFile(fileName: string): boolean {
  const lower = fileName.toLowerCase();
  return lower.endsWith(".xlsx") || lower.endsWith(".xlsm");
}

// Word拡張子判定（.docx のみ対応。旧形式 .doc は非対応）
function isWordFile(fileName: string): boolean {
  return fileName.toLowerCase().endsWith(".docx");
}

function _decodeXmlEntities(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"');
}

async function extractWordText(buffer: ArrayBuffer): Promise<string[]> {
  try {
    const JSZipModule = await import("jszip");
    const JSZip = JSZipModule.default ?? JSZipModule;
    const zip = await (JSZip as any).loadAsync(Buffer.from(new Uint8Array(buffer)));

    const fileKeys = Object.keys(zip.files);
    console.log(`[extractWordText] zip entries (first 10):`, fileKeys.slice(0, 10));

    const docXml = await zip.files["word/document.xml"]?.async("string");
    if (!docXml) {
      console.warn(`[extractWordText] word/document.xml not found in zip`);
      return [];
    }

    const paragraphs: string[] = [];
    const paraRe = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
    let pm: RegExpExecArray | null;
    while ((pm = paraRe.exec(docXml)) !== null) {
      const paraXml = pm[0];
      const textRe = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
      let text = "";
      let tm: RegExpExecArray | null;
      while ((tm = textRe.exec(paraXml)) !== null) {
        text += _decodeXmlEntities(tm[1]);
      }
      if (text.trim()) paragraphs.push(text.trim());
    }
    console.log(`[extractWordText] extracted ${paragraphs.length} paragraphs`);
    return paragraphs;
  } catch (e: any) {
    console.error(`[extractWordText] failed:`, String(e?.message ?? e));
    return [];
  }
}

const LoadFile = async (
  formData: FormData
): Promise<ServerActionResponse<string[]>> => {
  try {
    const file: File | null = formData.get("file") as unknown as File;

    if (file && file.size < MAX_UPLOAD_DOCUMENT_SIZE) {
      const buffer = await file.arrayBuffer();

      // Excel ファイル (.xlsx / .xlsm) は SheetJS で全シート抽出
      if (isExcelFile(file.name)) {
        console.log(`[LoadFile] Excel extraction: ${file.name}`);
        const docs = await extractExcelText(buffer);
        if (!docs.length) {
          return {
            status: "ERROR",
            errors: [{ message: "Excelファイルの内容が空か読み取れませんでした。" }],
          };
        }
        return { status: "OK", response: docs };
      }

      // Word ファイル (.docx) は JSZip でテキスト抽出を試みる
      if (isWordFile(file.name)) {
        console.log(`[LoadFile] Word extraction: ${file.name}`);
        const docs = await extractWordText(buffer);
        if (docs.length > 0) {
          return { status: "OK", response: docs };
        }
        // テキスト0件 = 画像埋め込み型Word（EMF等）
        // 旧SDK (ai-form-recognizer) はDOCX非対応のためDIには渡さず、案内メッセージを返す
        console.log(`[LoadFile] Word has no text (image-based docx). Returning guidance.`);
        return {
          status: "OK",
          response: [
            "このWordファイルは画像埋め込み型のため、テキストとして読み取ることができませんでした。",
            "PDFとして保存してアップロードするか、「このWordをExcelに変換して」と指示することで表データを抽出できます。",
          ],
        };
      }

      // その他のファイルは Azure Document Intelligence で抽出
      const client = DocumentIntelligenceInstance();

      const poller = await client.beginAnalyzeDocument(
        "prebuilt-read",
        buffer
      );
      const { paragraphs } = await poller.pollUntilDone();

      const docs: Array<string> = [];

      if (paragraphs) {
        for (const paragraph of paragraphs) {
          docs.push(paragraph.content);
        }
      }

      return {
        status: "OK",
        response: docs,
      };
    } else {
      return {
        status: "ERROR",
        errors: [
          {
            message: `File is too large and must be less than ${MAX_UPLOAD_DOCUMENT_SIZE} bytes.`,
          },
        ],
      };
    }
  } catch (e) {
    return {
      status: "ERROR",
      errors: [
        {
          message: `${e}`,
        },
      ],
    };
  }
};

export const FindAllChatDocuments = async (
  chatThreadID: string
): Promise<ServerActionResponse<ChatDocumentModel[]>> => {
  try {
    const querySpec: SqlQuerySpec = {
      query:
        "SELECT * FROM root r WHERE r.type=@type AND r.chatThreadId = @threadId AND r.isDeleted=@isDeleted",
      parameters: [
        {
          name: "@type",
          value: CHAT_DOCUMENT_ATTRIBUTE,
        },
        {
          name: "@threadId",
          value: chatThreadID,
        },
        {
          name: "@isDeleted",
          value: false,
        },
      ],
    };

    const { resources } = await HistoryContainer()
      .items.query<ChatDocumentModel>(querySpec)
      .fetchAll();

    if (resources) {
      return {
        status: "OK",
        response: resources,
      };
    } else {
      return {
        status: "ERROR",
        errors: [
          {
            message: "No documents found",
          },
        ],
      };
    }
  } catch (e) {
    return {
      status: "ERROR",
      errors: [
        {
          message: `${e}`,
        },
      ],
    };
  }
};

export const CreateChatDocument = async (
  fileName: string,
  chatThreadID: string
): Promise<ServerActionResponse<ChatDocumentModel>> => {
  try {
    const modelToSave: ChatDocumentModel = {
      chatThreadId: chatThreadID,
      id: uniqueId(),
      userId: await userHashedId(),
      createdAt: new Date(),
      type: CHAT_DOCUMENT_ATTRIBUTE,
      isDeleted: false,
      name: fileName,
    };

    const { resource } =
      await HistoryContainer().items.upsert<ChatDocumentModel>(modelToSave);
    RevalidateCache({
      page: "chat",
      params: chatThreadID,
    });

    if (resource) {
      return {
        status: "OK",
        response: resource,
      };
    }

    return {
      status: "ERROR",
      errors: [
        {
          message: "Unable to save chat document",
        },
      ],
    };
  } catch (e) {
    return {
      status: "ERROR",
      errors: [
        {
          message: `${e}`,
        },
      ],
    };
  }
};

export async function ChunkDocumentWithOverlap(
  document: string
): Promise<string[]> {
  const chunks: string[] = [];

  if (document.length <= CHUNK_SIZE) {
    // If the document is smaller than the desired chunk size, return it as a single chunk.
    chunks.push(document);
    return chunks;
  }

  let startIndex = 0;

  // Split the document into chunks of the desired size, with overlap.
  while (startIndex < document.length) {
    const endIndex = startIndex + CHUNK_SIZE;
    const chunk = document.substring(startIndex, endIndex);
    chunks.push(chunk);
    startIndex = endIndex - CHUNK_OVERLAP;
  }

  return chunks;
}
