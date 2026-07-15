import "server-only";

import { createHash } from "crypto";
import {
  DownloadBlobAsText,
  GenerateSasUrl,
  UploadBlob,
} from "@/features/common/services/azure-storage";
import {
  findTeamsOfficeFileCandidates,
  type TeamsOfficeFileCandidate,
} from "./teams-search-service";

export type TeamsOfficeRequest =
  | {
      action: "pdf_to_excel";
      fileQuery: string;
    }
  | {
      action: "refine_excel_sheets";
      targetSheets: string[];
    };

export function parseTeamsOfficeRequest(
  message: string
): TeamsOfficeRequest | null {
  const normalized = message.trim().normalize("NFKC");
  const asksForExcelRefinement =
    /(excel|エクセル|xlsx)/i.test(normalized) &&
    /(再変換|再抽出|精度|読み直|再読込|もう一度変換)/i.test(normalized);

  if (asksForExcelRefinement) {
    return {
      action: "refine_excel_sheets",
      targetSheets: extractTargetSheetNames(normalized),
    };
  }

  const asksForExcel = /(excel|エクセル|xlsx)/i.test(normalized);
  const asksForConversion = /(変換|出力|作成|にして|して)/i.test(normalized);
  const hasDocumentSource =
    /(sharepoint|\bsp\b|\bsl\b|pdf|word|docx|財務諸表)/i.test(normalized);

  if (!asksForExcel || !asksForConversion || !hasDocumentSource) return null;

  const quoted =
    normalized.match(/「([^」]+)」/)?.[1] ??
    normalized.match(/『([^』]+)』/)?.[1] ??
    normalized.match(/["“]([^"”]+)["”]/)?.[1];
  const fileQuery = quoted?.trim() || extractUnquotedFileQuery(normalized);

  if (!fileQuery) return null;
  return { action: "pdf_to_excel", fileQuery };
}

export async function executeTeamsOfficeRequest(props: {
  request: TeamsOfficeRequest;
  conversationId: string;
  userEmail?: string | null;
}): Promise<string> {
  if (props.request.action === "refine_excel_sheets") {
    if (props.request.targetSheets.length === 0) {
      return "再変換するシート名を指定してください。例: 「P2のシートだけ再変換して」";
    }

    const result = await refineExcelSheets({
      threadId: buildTeamsThreadId(props.conversationId),
      targetSheets: props.request.targetSheets,
    });
    return formatExcelRefinementResult(result);
  }

  const search = await findTeamsOfficeFileCandidates({
    query: props.request.fileQuery,
    userEmail: props.userEmail,
    extensions: ["pdf", "docx"],
  });

  if (search.exactMatches.length === 0) {
    return buildNotFoundMessage(props.request.fileQuery, search.suggestions);
  }
  if (search.exactMatches.length > 1) {
    return buildMultipleFilesMessage(search.exactMatches);
  }

  const file = search.exactMatches[0];
  const result = await convertDocumentToExcel({
    fileUrl: file.url,
    fileName: file.name,
    threadId: buildTeamsThreadId(props.conversationId),
  });

  if (result && typeof result === "object" && "error" in result) {
    return `Excelへの変換に失敗しました。\n\n${String(result.error)}`;
  }
  if (
    !result ||
    typeof result !== "object" ||
    !("downloadUrl" in result) ||
    typeof result.downloadUrl !== "string"
  ) {
    return "Excelへの変換は完了しましたが、ダウンロードリンクを取得できませんでした。";
  }

  const outputName =
    "fileName" in result && typeof result.fileName === "string"
      ? result.fileName
      : file.name.replace(/\.(pdf|docx)$/i, ".xlsx");
  const detail =
    "message" in result && typeof result.message === "string"
      ? `\n\n${result.message}`
      : "";

  return `Excelへの変換が完了しました。${detail}\n\n📊 [${escapeMarkdownLinkText(
    outputName
  )}](${result.downloadUrl})`;
}

type TeamsExcelPointer = {
  url: string;
  fileName: string;
  savedAt: number;
  sheetNames?: string[];
};

const excelPointerBlobName = (threadId: string) =>
  `thread-${threadId}-excel-latest.json`;

async function convertDocumentToExcel(props: {
  fileUrl: string;
  fileName: string;
  threadId: string;
}): Promise<Record<string, unknown>> {
  const sourceUrl = await resolveTeamsOfficeSourceUrl(props);
  const response = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      fileUrl: sourceUrl,
      instruction: "",
      threadId: props.threadId,
      action: "pdf_to_excel",
      outputBaseName: props.fileName,
    }),
  });

  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    return {
      error: `PDF→Excel変換に失敗しました: HTTP ${response.status}${
        detail ? ` ${detail.slice(0, 200)}` : ""
      }`,
    };
  }

  const result = (await response.json()) as Record<string, unknown>;
  if (result.engine === "none") {
    return { error: "このファイルから表データを抽出できませんでした。" };
  }
  if (typeof result.downloadUrl !== "string") {
    return { error: "ダウンロードURLが取得できませんでした。" };
  }

  const fileName =
    typeof result.fileName === "string"
      ? result.fileName
      : props.fileName.replace(/\.(pdf|docx)$/i, ".xlsx");
  const sheetNames = Array.isArray(result.sheetNames)
    ? result.sheetNames.filter(
        (item): item is string => typeof item === "string"
      )
    : undefined;
  await saveExcelPointer(props.threadId, {
    url: result.downloadUrl,
    fileName,
    savedAt: Date.now(),
    sheetNames,
  });

  const tables = Number(result.tables ?? 0);
  const sheets = Number(result.sheets ?? 0);
  const pages = Number(result.pages ?? 0);
  return {
    ...result,
    fileName,
    message:
      tables > 0
        ? `${pages}ページを変換しました（テーブル${tables}個、${sheets}シート）。`
        : `${pages}ページを変換しました。`,
  };
}

async function resolveTeamsOfficeSourceUrl(props: {
  fileUrl: string;
  fileName: string;
  threadId: string;
}): Promise<string> {
  const url = new URL(props.fileUrl);
  if (!url.hostname.endsWith(".sharepoint.com")) return props.fileUrl;

  const tenantId = requiredOfficeEnv("AZURE_AD_TENANT_ID");
  const clientId = requiredOfficeEnv("AZURE_AD_CLIENT_ID");
  const clientSecret = requiredOfficeEnv("AZURE_AD_CLIENT_SECRET");
  const tokenResponse = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: "client_credentials",
        scope: "https://graph.microsoft.com/.default",
      }),
    }
  );
  const tokenData = (await tokenResponse.json().catch(() => ({}))) as {
    access_token?: string;
    error_description?: string;
  };
  if (!tokenResponse.ok || !tokenData.access_token) {
    throw new Error(
      `SharePoint認証に失敗しました: ${
        tokenData.error_description ?? `HTTP ${tokenResponse.status}`
      }`
    );
  }

  const pathParts = decodeURIComponent(url.pathname).split("/").filter(Boolean);
  const sitesIndex = pathParts.indexOf("sites");
  if (sitesIndex < 0 || !pathParts[sitesIndex + 1]) {
    throw new Error("SharePointサイトをURLから特定できませんでした。");
  }
  const siteName = pathParts[sitesIndex + 1];
  const graphHeaders = { Authorization: `Bearer ${tokenData.access_token}` };
  const siteResponse = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${url.hostname}:/sites/${encodeURIComponent(
      siteName
    )}:`,
    { headers: graphHeaders }
  );
  const siteData = (await siteResponse.json().catch(() => ({}))) as {
    id?: string;
  };
  if (!siteResponse.ok || !siteData.id) {
    throw new Error(`SharePointサイトの取得に失敗しました: HTTP ${siteResponse.status}`);
  }

  const contentResponse = url.pathname.includes("/_layouts/")
    ? await downloadSharePointLayoutFile({
        siteId: siteData.id,
        fileName: url.searchParams.get("file") ?? props.fileName,
        headers: graphHeaders,
      })
    : await downloadSharePointPathFile({
        siteId: siteData.id,
        librarySegment: pathParts[sitesIndex + 2] ?? "",
        filePath: pathParts.slice(sitesIndex + 3).join("/"),
        headers: graphHeaders,
      });

  if (!contentResponse?.ok) {
    throw new Error(
      `SharePointファイルの取得に失敗しました: HTTP ${
        contentResponse?.status ?? "unknown"
      }`
    );
  }

  const safeFileName = props.fileName.replace(/[\\/:*?"<>|]/g, "_");
  const blobPath = `${props.threadId}/${safeFileName}`;
  const upload = await UploadBlob(
    "dl-link",
    blobPath,
    Buffer.from(await contentResponse.arrayBuffer())
  );
  if (upload.status !== "OK") {
    throw new Error("SharePointファイルの一時保存に失敗しました。");
  }
  const sas = await GenerateSasUrl("dl-link", blobPath);
  if (sas.status !== "OK") {
    throw new Error("SharePointファイルの一時URLを生成できませんでした。");
  }
  return sas.response;
}

async function downloadSharePointPathFile(props: {
  siteId: string;
  librarySegment: string;
  filePath: string;
  headers: { Authorization: string };
}): Promise<Response | null> {
  const drivesResponse = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${props.siteId}/drives`,
    { headers: props.headers }
  );
  const drivesData = (await drivesResponse.json().catch(() => ({}))) as {
    value?: Array<{ id?: string; name?: string; webUrl?: string }>;
  };
  const drive = drivesData.value?.find((item) => {
    const webSegment = decodeURIComponent(
      String(item.webUrl ?? "").split("/").pop() ?? ""
    );
    return item.name === props.librarySegment || webSegment === props.librarySegment;
  });
  if (!drive?.id || !props.filePath) return null;

  return fetch(
    `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${props.filePath
      .split("/")
      .map(encodeURIComponent)
      .join("/")}:/content`,
    { headers: props.headers }
  );
}

async function downloadSharePointLayoutFile(props: {
  siteId: string;
  fileName: string;
  headers: { Authorization: string };
}): Promise<Response | null> {
  const drivesResponse = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${props.siteId}/drives`,
    { headers: props.headers }
  );
  const drivesData = (await drivesResponse.json().catch(() => ({}))) as {
    value?: Array<{ id?: string }>;
  };

  for (const drive of drivesData.value ?? []) {
    if (!drive.id) continue;
    const searchResponse = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${drive.id}/root/search(q='${encodeURIComponent(
        props.fileName
      )}')`,
      { headers: props.headers }
    );
    const searchData = (await searchResponse.json().catch(() => ({}))) as {
      value?: Array<{ id?: string; name?: string; file?: unknown }>;
    };
    const item = searchData.value?.find(
      (candidate) =>
        candidate.file &&
        candidate.name?.toLowerCase() === props.fileName.toLowerCase()
    );
    if (item?.id) {
      return fetch(
        `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${item.id}/content`,
        { headers: props.headers }
      );
    }
  }
  return null;
}

async function refineExcelSheets(props: {
  threadId: string;
  targetSheets: string[];
}): Promise<Record<string, unknown>> {
  const pointer = await readExcelPointer(props.threadId);
  if (!pointer?.url) {
    return {
      error:
        "このTeams会話で変換したExcelが見つかりません。先にPDF／WordをExcelへ変換してください。",
    };
  }

  const sheetToProcess = props.targetSheets[0];
  const outputFileName = buildRefinedFileName(pointer.fileName);
  const response = await fetch(`${getOfficeApiBaseUrl()}/api/edit-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      excelFileUrl: pointer.url,
      targetSheets: [sheetToProcess],
      outputFileName,
      instruction: "",
      threadId: props.threadId,
      action: "refine_excel_pages",
    }),
  });

  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    return {
      error: `Excel再変換に失敗しました: HTTP ${response.status}${
        detail ? ` ${detail.slice(0, 200)}` : ""
      }`,
    };
  }

  const result = (await response.json()) as Record<string, unknown>;
  const didRefine = Number(result.refined ?? 0) > 0;
  if (!didRefine || typeof result.downloadUrl !== "string") {
    return {
      error: `「${sheetToProcess}」は再変換できませんでした。最新Excelは更新していません。`,
    };
  }

  const fileName =
    typeof result.fileName === "string" ? result.fileName : outputFileName;
  await saveExcelPointer(props.threadId, {
    ...pointer,
    url: result.downloadUrl,
    fileName,
    savedAt: Date.now(),
  });

  return {
    ...result,
    fileName,
    processedSheet: sheetToProcess,
    message: `「${sheetToProcess}」をVision AIで再変換しました。`,
  };
}

async function saveExcelPointer(
  threadId: string,
  pointer: TeamsExcelPointer
): Promise<void> {
  const response = await UploadBlob(
    "dl-link",
    excelPointerBlobName(threadId),
    Buffer.from(JSON.stringify(pointer))
  );
  if (response.status !== "OK") {
    throw new Error("Failed to save Teams Excel pointer.");
  }
}

async function readExcelPointer(
  threadId: string
): Promise<TeamsExcelPointer | null> {
  const response = await DownloadBlobAsText(
    "dl-link",
    excelPointerBlobName(threadId)
  );
  if (response.status !== "OK") return null;

  try {
    return JSON.parse(response.response) as TeamsExcelPointer;
  } catch {
    return null;
  }
}

function buildRefinedFileName(fileName: string): string {
  const base = fileName
    .replace(/\.xlsx$/i, "")
    .replace(/(_精度向上後)+$/, "")
    .replace(/[　 ﻿<>:"/\\|?*\x00-\x1f]/g, "_");
  const revision = base.match(/^(.*?)_rev(\d+)$/i);
  return revision
    ? `${revision[1]}_rev${Number(revision[2]) + 1}.xlsx`
    : `${base}_rev1.xlsx`;
}

function getOfficeApiBaseUrl(): string {
  return (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME
      ? `https://${process.env.WEBSITE_HOSTNAME}`
      : "http://localhost:3000")
  ).replace(/\/+$/, "");
}

function requiredOfficeEnv(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`Missing environment variable: ${name}`);
  return value;
}

function extractTargetSheetNames(message: string): string[] {
  const names = new Set<string>();
  const quotedPattern = /[「『"“]([^」』"”]+)[」』"”]/g;
  let match: RegExpExecArray | null;
  while ((match = quotedPattern.exec(message)) !== null) {
    const value = match[1]?.trim();
    if (value && !/再変換|再抽出/.test(value)) names.add(value);
  }

  // ハイフンを含むシート名（例: P3-T2）は分割せず、一つの名前として扱う。
  // 複合名を先に取り除くことで、後段の単一名検索が P3 / T2 を重複追加しないようにする。
  const compoundSheetPattern = /\b([a-z]+\d+(?:-[a-z]+\d+)+)\b/gi;
  const messageWithoutCompoundNames = message.replace(
    compoundSheetPattern,
    (value) => {
      names.add(value.toUpperCase());
      return " ".repeat(value.length);
    }
  );

  const cellStylePattern = /\b([a-z]+\d+)\b/gi;
  while ((match = cellStylePattern.exec(messageWithoutCompoundNames)) !== null) {
    if (match[1]) names.add(match[1].toUpperCase());
  }

  return Array.from(names);
}

function formatExcelRefinementResult(result: unknown): string {
  if (result && typeof result === "object" && "error" in result) {
    return `Excelシートの再変換に失敗しました。\n\n${String(result.error)}`;
  }
  if (
    !result ||
    typeof result !== "object" ||
    !("downloadUrl" in result) ||
    typeof result.downloadUrl !== "string"
  ) {
    return "Excelシートの再変換は完了しましたが、ダウンロードリンクを取得できませんでした。";
  }

  const fileName =
    "fileName" in result && typeof result.fileName === "string"
      ? result.fileName
      : "refined.xlsx";
  const message =
    "message" in result && typeof result.message === "string"
      ? result.message
      : "指定シートを再変換しました。";

  return `${message}\n\n📊 [${escapeMarkdownLinkText(fileName)}](${result.downloadUrl})`;
}

function extractUnquotedFileQuery(message: string): string | null {
  const patterns = [
    /(?:sharepoint|\bsp\b)(?:上|内)?(?:にある|の)?\s*(?:pdf|word|docx)?\s*(?:ファイル)?\s*(.+?)\s*を\s*(?:excel|エクセル|xlsx)/i,
    /\bsl\b(?:上|内)?(?:にある|の)?\s*(?:pdf|word|docx)?\s*(?:ファイル)?\s*(.+?)\s*を\s*(?:excel|エクセル|xlsx)/i,
    /(?:pdf|word|docx)(?:ファイル)?\s*(.+?)\s*を\s*(?:excel|エクセル|xlsx)/i,
  ];

  for (const pattern of patterns) {
    const match = message.match(pattern)?.[1]?.trim();
    if (match) return match;
  }
  return null;
}

function buildNotFoundMessage(
  query: string,
  suggestions: TeamsOfficeFileCandidate[]
): string {
  if (suggestions.length === 0) {
    return `「${query}」に一致する、アクセス可能なPDF／Wordファイルが見つかりませんでした。`;
  }

  const list = suggestions.map((file) => `- ${file.name}`).join("\n");
  return `「${query}」に完全一致するファイルが見つかりませんでした。候補は次のとおりです。\n\n${list}\n\nファイル名を「」で囲んで再度指定してください。`;
}

function buildMultipleFilesMessage(files: TeamsOfficeFileCandidate[]): string {
  const list = files.slice(0, 10).map((file) => `- ${file.name}`).join("\n");
  return `複数のファイルが見つかりました。変換するファイル名を「」で囲んで指定してください。\n\n${list}`;
}

function buildTeamsThreadId(conversationId: string): string {
  return `teams-${createHash("sha256")
    .update(conversationId)
    .digest("hex")
    .slice(0, 32)}`;
}

function escapeMarkdownLinkText(value: string): string {
  return value.replace(/([\\[\]])/g, "\\$1");
}
