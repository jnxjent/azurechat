// src/features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts
"use server";
import "server-only";

import { DownloadBlobAsText, GenerateSasUrl, UploadBlob } from "@/features/common/services/azure-storage";
import { OpenAIDALLEInstance, OpenAIInstance } from "@/features/common/services/openai";
import { ServerActionResponse } from "@/features/common/server-action-response";
import { uniqueId } from "@/features/common/util";
import { GetImageUrl, UploadImageToStore } from "../chat-image-service";
import { FindTopChatMessagesForCurrentUser } from "../chat-message-service";
import { FindAllChatDocuments } from "../chat-document-service";
import { ChatThreadModel } from "../models";
import { BlobServiceClient } from "@azure/storage-blob";
import { analyzeDocVision } from "@/app/api/analyze-doc-vision/handler";
import { SimpleSearch, SimilaritySearch, ExtensionSimilaritySearch } from "@/features/chat-page/chat-services/azure-ai-search/azure-ai-search";
import { userSession } from "@/features/auth-page/helpers";

import {
  buildSendOptionsFromMode,
  canonicalizeMode,
  type ThinkingModeInput,
} from "@/features/chat-page/chat-services/chat-api/reasoning-utils";

type ThinkingModeAPI = "normal" | "thinking" | "fast";

/** standard を normal へ、その他はそのまま（保険） */
function normalizeThinkingMode(
  input?: ThinkingModeAPI | ThinkingModeInput
): ThinkingModeAPI {
  const c = canonicalizeMode(input as any);
  return c as ThinkingModeAPI;
}

/**
 * 画像URLを組み立てる共通ヘルパー
 */
function buildExternalImageUrl(threadId: string, fileName: string): string {
  const publicBase = process.env.NEXT_PUBLIC_IMAGE_URL;
  if (publicBase) {
    const base = publicBase.replace(/\/+$/, "");
    return `${base}/?t=${threadId}&img=${fileName}`;
  }

  const nextAuth = process.env.NEXTAUTH_URL;
  if (nextAuth) {
    const base = nextAuth.replace(/\/+$/, "");
    return `${base}/api/images/?t=${threadId}&img=${fileName}`;
  }

  return GetImageUrl(threadId, fileName);
}

/**
 * SAS なし Azure Blob URL を {container, path} に分解する。
 * SAS 付き・非 Blob URL は null を返す。
 */
function parseBlobRawUrl(rawUrl: string | null | undefined): { container: string; path: string } | null {
  if (!rawUrl?.trim()) return null;
  try {
    const obj = new URL(rawUrl);
    const isAzureBlob =
      obj.hostname.endsWith(".blob.core.windows.net") ||
      obj.host === "127.0.0.1:10000" ||
      obj.host === "localhost:10000";
    if (!isAzureBlob || obj.searchParams.has("sig")) return null;
    const parts = obj.pathname.split("/").filter(Boolean);
    if (parts.length < 2) return null;
    return { container: parts[0], path: parts.slice(1).join("/") };
  } catch {
    return null;
  }
}

async function resolveDocumentUrlForVision(
  fileUrl: string,
  threadId: string
): Promise<string> {
  try {
    /*
    if (sourceFileUrls.length > 1) {
      const mergedSlides: Array<{
        title: string;
        bullets: string[];
        layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram" | "conversation" | "stat_callouts" | "card_grid" | "icon_rows" | "metric-cards" | "process-cards" | "timeline" | "company-overview" | "closing";
        tableRows?: string[][];
        columns?: Array<{ header: string; bullets: string[] }>;
        conversationStyle?: "chat-ui" | "interview" | "dialog-list";
        conversationTurns?: Array<{
          speakerRole: string;
          speakerType?: "agent" | "customer" | "staff" | "other";
          text: string;
          turnIndex: number;
        }>;
      }> = [];
      let mergedTotalPages = 0;

      for (const currentFileUrl of sourceFileUrls) {
        const resolvedFileUrl = await resolveDocumentUrlForVision(
          currentFileUrl,
          chatThread.id
        );
        console.log("[convert_doc_to_pptx] Analyzing document with Vision API:", {
          sourceFile: extractFileNameFromDocumentUrl(currentFileUrl),
          resolvedUrl: resolvedFileUrl.substring(0, 80),
        });
        const analyzeResult = await analyzeDocVision(resolvedFileUrl, maxPages ?? 30, mode);

        if (!analyzeResult?.ok || !analyzeResult.slides?.length) {
          console.error("[convert_doc_to_pptx] analyze-doc-vision failed:", analyzeResult?.error);
          return { error: analyzeResult?.error ?? "ドキュメント解析結果を取得できませんでした。" };
        }

        mergedSlides.push(...analyzeResult.slides);
        mergedTotalPages += analyzeResult.totalPages ?? analyzeResult.slides.length;
      }

      const mergedTitle =
        mergedSlides[0]?.title ||
        derivedTitle ||
        presentationTitle?.trim() ||
        "プレゼンテーション";

      console.log("[convert_doc_to_pptx] Title sources:", {
        derivedTitle,
        presentationTitle,
        deckPreferences,
        firstSlideTitle: mergedSlides[0]?.title,
        finalTitle: mergedTitle,
      });
      console.log("[convert_doc_to_pptx] Aggregated deck:", {
        fileCount: sourceFileUrls.length,
        totalPages: mergedTotalPages,
        slideCount: mergedSlides.length,
      });

      const pptxRes = await fetch(`${baseUrl}/api/gen-pptx`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          title: mergedTitle,
          slides: mergedSlides,
          threadId: chatThread.id,
          fontFace,
          designInstruction: deckPreferences.designInstruction,
          deckPreferences,
          mode,
        }),
      });

      if (!pptxRes.ok) {
        const t = await pptxRes.text().catch(() => "");
        console.error("[convert_doc_to_pptx] gen-pptx failed:", pptxRes.status, t);
        return { error: `PowerPoint生成に失敗しました: HTTP ${pptxRes.status}` };
      }

      const pptxResult = await pptxRes.json();
      if (!pptxResult?.downloadUrl) {
        return { error: "ダウンロードURLを取得できませんでした。" };
      }

      return {
        downloadUrl: pptxResult.downloadUrl,
        fileName: pptxResult.fileName,
        totalPages: mergedTotalPages,
        message: `${sourceFileUrls.length}件の資料をまとめて${mergedTotalPages}ページ分を解析し、PowerPointを生成しました。`,
      };
    }
    /*
    if (sourceFileUrls.length > 1) {
      const mergedSlides: Array<{
        title: string;
        bullets: string[];
        layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram" | "conversation" | "stat_callouts" | "card_grid" | "icon_rows" | "metric-cards" | "process-cards" | "timeline" | "company-overview" | "closing";
        tableRows?: string[][];
        columns?: Array<{ header: string; bullets: string[] }>;
        conversationStyle?: "chat-ui" | "interview" | "dialog-list";
        conversationTurns?: Array<{
          speakerRole: string;
          speakerType?: "agent" | "customer" | "staff" | "other";
          text: string;
          turnIndex: number;
        }>;
      }> = [];
      let mergedTotalPages = 0;

      for (const currentFileUrl of sourceFileUrls) {
        const resolvedFileUrl = await resolveDocumentUrlForVision(
          currentFileUrl,
          chatThread.id
        );
        console.log("[convert_doc_to_pptx] Analyzing document with Vision API:", {
          sourceFile: extractFileNameFromDocumentUrl(currentFileUrl),
          resolvedUrl: resolvedFileUrl.substring(0, 80),
        });
        const analyzeResult = await analyzeDocVision(resolvedFileUrl, maxPages ?? 30, mode);

        if (!analyzeResult?.ok || !analyzeResult.slides?.length) {
          console.error("[convert_doc_to_pptx] analyze-doc-vision failed:", analyzeResult?.error);
          return { error: analyzeResult?.error ?? "ドキュメント解析結果を取得できませんでした。" };
        }

        mergedSlides.push(...analyzeResult.slides);
        mergedTotalPages += analyzeResult.totalPages ?? analyzeResult.slides.length;
      }

      const mergedTitle =
        mergedSlides[0]?.title ||
        derivedTitle ||
        presentationTitle?.trim() ||
        "プレゼンテーション";

      console.log("[convert_doc_to_pptx] Title sources:", {
        derivedTitle,
        presentationTitle,
        deckPreferences,
        firstSlideTitle: mergedSlides[0]?.title,
        finalTitle: mergedTitle,
      });
      console.log("[convert_doc_to_pptx] Aggregated deck:", {
        fileCount: sourceFileUrls.length,
        totalPages: mergedTotalPages,
        slideCount: mergedSlides.length,
      });

      const pptxRes = await fetch(`${baseUrl}/api/gen-pptx`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          title: mergedTitle,
          slides: mergedSlides,
          threadId: chatThread.id,
          fontFace,
          designInstruction: deckPreferences.designInstruction,
          deckPreferences,
          mode,
        }),
      });

      if (!pptxRes.ok) {
        const t = await pptxRes.text().catch(() => "");
        console.error("[convert_doc_to_pptx] gen-pptx failed:", pptxRes.status, t);
        return { error: `PowerPoint生成に失敗しました: HTTP ${pptxRes.status}` };
      }

      const pptxResult = await pptxRes.json();
      if (!pptxResult?.downloadUrl) {
        return { error: "ダウンロードURLを取得できませんでした。" };
      }

      return {
        downloadUrl: pptxResult.downloadUrl,
        fileName: pptxResult.fileName,
        totalPages: mergedTotalPages,
        message: `${sourceFileUrls.length}件の資料をまとめて${mergedTotalPages}ページ分を解析し、PowerPointを生成しました。`,
      };
    }
    */
    const url = new URL(fileUrl);
    const isSharePointUrl = url.hostname.includes("sharepoint.com");
    const isAzureBlobWithoutSas =
      url.hostname.includes(".blob.core.windows.net") && !url.searchParams.has("sig");

    if (!isSharePointUrl && !isAzureBlobWithoutSas) {
      return fileUrl;
    }

    const fileName = extractFileNameFromDocumentUrl(fileUrl);
    if (!fileName) {
      return fileUrl;
    }

    const resolvedBlobPath = await findThreadDocumentBlobPath(threadId, fileName);
    if (resolvedBlobPath) {
      const sasResponse = await GenerateSasUrl("dl-link", resolvedBlobPath);
      if (sasResponse.status === "OK" && sasResponse.response) {
        console.log(
          `[convert_doc_to_pptx] Resolved document URL to SAS for thread ${threadId}: ${resolvedBlobPath}`
        );
        return sasResponse.response;
      }
    }

    // blob未キャッシュのSharePointファイル → Graph APIでダウンロードしてblobに保存
    if (isSharePointUrl) {
      const spSas = await downloadSharePointFileToBlob(fileUrl, threadId, fileName);
      if (spSas) return spSas;
      console.warn(
        `[convert_doc_to_pptx] Graph API download failed for ${fileName}, falling back to direct URL`
      );
    }
  } catch (error) {
    console.warn("[convert_doc_to_pptx] Failed to resolve document URL for Vision:", error);
  }

  return fileUrl;
}

/**
 * SharePoint ファイルを Graph API (app-only token) でダウンロードし、
 * Azure Blob Storage の dl-link/${threadId}/${fileName} にキャッシュして SAS URL を返す。
 */
async function downloadSharePointFileToBlob(
  sharePointUrl: string,
  threadId: string,
  fileName: string
): Promise<string | null> {
  try {
    const tenantId = process.env.AZURE_AD_TENANT_ID?.trim();
    const clientId = process.env.AZURE_AD_CLIENT_ID?.trim();
    const clientSecret = process.env.AZURE_AD_CLIENT_SECRET?.trim();
    if (!tenantId || !clientId || !clientSecret) {
      console.warn("[convert_doc_to_pptx] Azure AD env vars not set, skipping Graph download");
      return null;
    }

    // 1. app-only トークン取得
    const tokenRes = await fetch(
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
    const tokenData: any = await tokenRes.json().catch(() => ({}));
    if (!tokenRes.ok || !tokenData.access_token) {
      console.warn("[convert_doc_to_pptx] Graph token fetch failed:", tokenData.error_description ?? tokenData.error);
      return null;
    }
    const accessToken: string = tokenData.access_token;

    // 2. SharePoint URL を分解
    const urlObj = new URL(sharePointUrl);

    // 2a. _layouts/15/Doc.aspx?sourcedoc={GUID} 形式の場合: Graph API でファイル名検索してダウンロード
    // SP REST API は Sites.ReadAll (Graph) 権限のみでは使えないため、Graph drive search を使う
    if (urlObj.pathname.includes("/_layouts/")) {
      // URL の file= パラメータからファイル名を取得（なければ引数の fileName を使う）
      const fileNameParam = urlObj.searchParams.get("file") ?? fileName;

      // /_layouts より前のパスがサイトパス
      const layoutsIdx = urlObj.pathname.indexOf("/_layouts");
      const sitePath = urlObj.pathname.substring(0, layoutsIdx); // e.g. "/sites/SiteName"
      const sitePathParts = sitePath.split("/").filter(Boolean);
      const siteIdx = sitePathParts.indexOf("sites");
      if (siteIdx < 0) {
        console.warn("[downloadSharePointFileToBlob] Cannot extract site name from _layouts URL");
        return null;
      }
      const siteName2 = sitePathParts[siteIdx + 1];

      // Graph API でサイト ID 解決
      const siteRes2 = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${urlObj.hostname}:/sites/${siteName2}:`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      const siteData2: any = await siteRes2.json().catch(() => ({}));
      if (!siteRes2.ok || !siteData2.id) {
        console.warn("[downloadSharePointFileToBlob] site resolve failed for _layouts URL:", siteData2.error?.message ?? siteRes2.status);
        return null;
      }
      const siteId2: string = siteData2.id;

      // ファイル名で Graph API drive 検索
      const driveSearchRes = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId2}/drive/search(q='${encodeURIComponent(fileNameParam)}')`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      const driveSearchData: any = await driveSearchRes.json().catch(() => ({}));
      const foundItem = (driveSearchData.value ?? []).find(
        (item: any) =>
          item.name?.toLowerCase() === fileNameParam.toLowerCase() && item.file
      );
      if (!foundItem) {
        console.warn("[downloadSharePointFileToBlob] Graph drive search: no match for", fileNameParam);
        return null;
      }
      const driveId2: string = foundItem.parentReference?.driveId;
      const itemId2: string = foundItem.id;
      if (!driveId2 || !itemId2) return null;

      // driveItem content をダウンロード
      const contentRes = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId2}/items/${itemId2}/content`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      if (!contentRes.ok) {
        console.warn("[downloadSharePointFileToBlob] Graph driveItem download failed:", contentRes.status, fileNameParam);
        return null;
      }

      const fileBuffer2 = Buffer.from(await contentRes.arrayBuffer());
      const blobPath2 = `${threadId}/${fileName}`;
      const upResult2 = await UploadBlob("dl-link", blobPath2, fileBuffer2);
      if (upResult2.status !== "OK") {
        console.warn("[downloadSharePointFileToBlob] Blob upload failed after Graph search download");
        return null;
      }
      const sasRes2 = await GenerateSasUrl("dl-link", blobPath2);
      if (sasRes2.status === "OK" && sasRes2.response) {
        console.log(`[edit_sp_pptx] SP file cached via Graph drive search: ${blobPath2}`);
        return sasRes2.response;
      }
      return null;
    }

    // 2b. 通常 SP パス URL の場合: site + library + file path を取得
    const hostname = urlObj.hostname;
    const decodedPath = decodeURIComponent(urlObj.pathname);
    const pathParts = decodedPath.split("/").filter(Boolean);
    // 例: ["sites", "AzureChatxSharepointTestSite", "SL", "j.nomoto", "file.pdf"]
    const siteIndex = pathParts.indexOf("sites");
    if (siteIndex < 0 || siteIndex + 2 >= pathParts.length) return null;
    const siteName = pathParts[siteIndex + 1];
    const librarySegment = pathParts[siteIndex + 2]; // ライブラリのURLセグメント (例: "SL")
    const filePathWithinLibrary = pathParts.slice(siteIndex + 3).join("/"); // ライブラリ内のパス

    // 3. Graph API でサイト ID を解決
    const siteRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${hostname}:/sites/${siteName}:`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const siteData: any = await siteRes.json().catch(() => ({}));
    if (!siteRes.ok || !siteData.id) {
      console.warn("[convert_doc_to_pptx] Graph site resolve failed:", siteData.error?.message ?? siteRes.status);
      return null;
    }
    const siteId: string = siteData.id;

    // 4. ドライブ一覧からライブラリに対応するドライブを特定
    const drivesRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const drivesData: any = await drivesRes.json().catch(() => ({}));
    if (!drivesRes.ok || !drivesData.value?.length) {
      console.warn("[convert_doc_to_pptx] Graph drives fetch failed:", drivesData.error?.message ?? drivesRes.status);
      return null;
    }
    const matchedDrive = drivesData.value.find((d: any) => {
      const webUrlSlug = decodeURIComponent(String(d.webUrl ?? "").split("/").pop() ?? "");
      return d.name === librarySegment || webUrlSlug === librarySegment;
    });
    if (!matchedDrive) {
      console.warn(
        `[convert_doc_to_pptx] Drive not found for library "${librarySegment}". Available: ${drivesData.value.map((d: any) => d.name).join(", ")}`
      );
      return null;
    }
    const driveId: string = matchedDrive.id;

    // 5. ライブラリ内のパスでファイルをダウンロード
    const fileRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePathWithinLibrary}:/content`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    if (!fileRes.ok) {
      console.warn("[convert_doc_to_pptx] Graph file download failed:", fileRes.status, `(drive=${matchedDrive.name}, path=${filePathWithinLibrary})`);
      return null;
    }

    // 6. Azure Blob Storage にキャッシュ
    const fileBuffer = Buffer.from(await fileRes.arrayBuffer());
    const blobPath = `${threadId}/${fileName}`;
    const uploadResult = await UploadBlob("dl-link", blobPath, fileBuffer);
    if (uploadResult.status !== "OK") {
      console.warn("[convert_doc_to_pptx] Blob upload failed after Graph download");
      return null;
    }

    // 7. SAS URL 生成
    const sasResponse = await GenerateSasUrl("dl-link", blobPath);
    if (sasResponse.status === "OK" && sasResponse.response) {
      console.log(`[convert_doc_to_pptx] SP file cached to blob via Graph: ${blobPath}`);
      return sasResponse.response;
    }

    return null;
  } catch (e) {
    console.warn("[convert_doc_to_pptx] downloadSharePointFileToBlob error:", e);
    return null;
  }
}

async function findThreadDocumentBlobPath(
  threadId: string,
  fileName: string
): Promise<string | null> {
  const directPath = `${threadId}/${fileName}`;
  const direct = await GenerateSasUrl("dl-link", directPath);
  if (direct.status === "OK" && direct.response) {
    const headRes = await fetch(direct.response, { method: "HEAD" }).catch(() => null);
    if (headRes?.ok) {
      return directPath;
    }
  }

  const acc = process.env.AZURE_STORAGE_ACCOUNT_NAME;
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY;
  if (!acc || !key) return null;

  const connStr = `DefaultEndpointsProtocol=https;AccountName=${acc};AccountKey=${key};EndpointSuffix=core.windows.net`;
  const svc = BlobServiceClient.fromConnectionString(connStr);
  const cc = svc.getContainerClient("dl-link");
  const target = fileName.trim().toLowerCase();
  for await (const blob of cc.listBlobsFlat({ prefix: `${threadId}/` })) {
    const blobName = blob.name.split("/").pop()?.trim().toLowerCase();
    if (blobName === target) {
      return blob.name;
    }
  }

  return null;
}

function extractFileNameFromDocumentUrl(fileUrl: string): string | null {
  try {
    const url = new URL(fileUrl);
    const sharePointFileName = url.searchParams.get("file");
    if (sharePointFileName?.trim()) {
      return decodeURIComponent(sharePointFileName).trim();
    }

    const pathFileName = decodeURIComponent(url.pathname.split("/").pop() ?? "").trim();
    if (!pathFileName || pathFileName.toLowerCase() === "doc.aspx") {
      return null;
    }

    return pathFileName;
  } catch {
    return null;
  }
}

function extractPresentationTitleFromFileUrl(fileUrl: string): string | null {
  const fileName = extractFileNameFromDocumentUrl(fileUrl);
  if (!fileName) {
    return null;
  }

  const title = fileName.replace(/\.[^.]+$/, "").trim();
  return title || null;
}

function normalizeDocumentUrlInput(value: unknown): string {
  const raw = String(value ?? "").trim();
  if (!raw) {
    return "";
  }

  const labelMatch = raw.match(/^(?:file_url|fileUrl)\s*:\s*(.+)$/i);
  const candidate = labelMatch?.[1]?.trim() ?? raw;
  const firstHttpIndex = candidate.search(/https?:\/\//i);
  const normalized = firstHttpIndex >= 0 ? candidate.slice(firstHttpIndex).trim() : candidate;

  try {
    return new URL(normalized).toString();
  } catch {
    return normalized;
  }
}

function isHttpUrl(value: string): boolean {
  try {
    const parsed = new URL(value);
    return parsed.protocol === "http:" || parsed.protocol === "https:";
  } catch {
    return false;
  }
}

function extractLatestPptxUrlFromMessages(messages: string[]): string | null {
  const urlPattern = /https?:\/\/[^\s)\]]+\.pptx(?:\?[^\s)\]]*)?/gi;
  for (const message of messages) {
    const matches = message.match(urlPattern);
    if (matches?.length) {
      return matches[matches.length - 1];
    }
  }
  return null;
}

/** Markdownリンク [DisplayName.pptx](URL) からURL+表示名を両方取得する */
function extractLatestPptxInfoFromMessages(messages: string[]): { url: string; displayName: string | null } | null {
  const mdPattern = /\[([^\]]+?\.pptx)\]\((https?:\/\/[^\s)]+\.pptx(?:\?[^\s)]*)?)\)/gi;
  const urlPattern = /https?:\/\/[^\s)\]]+\.pptx(?:\?[^\s)\]]*)?/gi;
  for (const message of messages) {
    mdPattern.lastIndex = 0;
    let mdMatch: RegExpExecArray | null;
    let lastMdMatch: RegExpExecArray | null = null;
    while ((mdMatch = mdPattern.exec(message)) !== null) {
      lastMdMatch = mdMatch;
    }
    if (lastMdMatch) {
      return { url: lastMdMatch[2], displayName: lastMdMatch[1].replace(/\.pptx$/i, "").trim() };
    }
    urlPattern.lastIndex = 0;
    const urlMatches = message.match(urlPattern);
    if (urlMatches?.length) {
      return { url: urlMatches[urlMatches.length - 1], displayName: null };
    }
  }
  return null;
}

function extractLatestXlsxUrlFromMessages(messages: string[]): string | null {
  // messages は createdAt DESC（新しい順）で渡される前提
  // 最初にヒットした URL を即 return することで「最新」を確保する
  const urlPattern = /https?:\/\/[^\s)\]"']+\.(?:xlsx|xls|xlsm)(?:\?[^\s)\]"']*)?/gi;
  for (const message of messages) {
    const matches = message.match(urlPattern);
    if (matches?.length) {
      // Blob URL（blob.core.windows.net）を優先、なければ最後の一致
      const blobUrl = matches.find((u) => u.includes("blob.core.windows.net"));
      return blobUrl ?? matches[matches.length - 1];
    }
  }
  return null;
}

// ---------- スレッド単位の最新 Excel URL ポインタ (Blob Storage) ----------

type ExcelPointer = { url: string; fileName: string; savedAt: number; sourceFileQuery?: string; chartEdits?: object[] };
const EXCEL_PTR_BLOB = (threadId: string) => `thread-${threadId}-excel-latest.json`;

async function saveLatestExcelUrl(
  threadId: string,
  url: string,
  fileName: string,
  sourceFileQuery?: string,
  chartEdits?: object[]
): Promise<void> {
  try {
    const data: ExcelPointer = { url, fileName, savedAt: Date.now(), sourceFileQuery, chartEdits };
    await UploadBlob("dl-link", EXCEL_PTR_BLOB(threadId), Buffer.from(JSON.stringify(data)));
    console.log(`[excel-ptr] saved pointer for thread ${threadId}: ${fileName} (query: ${sourceFileQuery ?? "-"}, charts: ${chartEdits?.length ?? 0})`);
  } catch (e) {
    console.warn("[excel-ptr] save failed:", e);
  }
}

async function readLatestExcelPtr(threadId: string): Promise<ExcelPointer | null> {
  try {
    const res = await DownloadBlobAsText("dl-link", EXCEL_PTR_BLOB(threadId));
    if (res.status !== "OK") return null;
    return JSON.parse(res.response) as ExcelPointer;
  } catch {
    return null;
  }
}

// -----------------------------------------------------------------------

async function resolveLatestXlsxUrlFromThread(chatThreadId: string): Promise<string | null> {
  try {
    // 0th: Blob ポインタ vs 新規アップロードを比較して新しい方を使う
    const [ptr, docsResponse] = await Promise.all([
      readLatestExcelPtr(chatThreadId),
      FindAllChatDocuments(chatThreadId),
    ]);

    const latestUploadDoc = docsResponse.status === "OK"
      ? docsResponse.response
          .filter((doc) => /\.(xlsx|xls|xlsm)$/i.test(doc.name))
          .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())[0]
      : null;
    const latestUploadTime = latestUploadDoc ? new Date(latestUploadDoc.createdAt).getTime() : 0;

    if (ptr?.url) {
      if (latestUploadTime > ptr.savedAt) {
        // 新規アップロードがポインタより新しい → アップロードを優先
        console.log(`[resolveLatestXlsx] newer upload (${latestUploadDoc!.name}) > pointer, using upload`);
        const sasRes = await GenerateSasUrl("dl-link", `${chatThreadId}/${latestUploadDoc!.name}`);
        if (sasRes.status === "OK") return sasRes.response;
      }
      console.log(`[resolveLatestXlsx] using blob pointer: ${ptr.fileName}`);
      return ptr.url;
    }

    // 1st: scan message history for xlsx URLs
    const historyResponse = await FindTopChatMessagesForCurrentUser(chatThreadId, 20);
    if (historyResponse.status === "OK") {
      const messages = historyResponse.response
        .map((message) => String(message.content ?? "").trim())
        .filter(Boolean);
      const fromHistory = extractLatestXlsxUrlFromMessages(messages);
      if (fromHistory) return fromHistory;
    }

    // 2nd: fall back to ChatDocuments (already fetched above)
    if (latestUploadDoc) {
      const sasRes = await GenerateSasUrl("dl-link", `${chatThreadId}/${latestUploadDoc.name}`);
      if (sasRes.status === "OK") return sasRes.response;
    }
  } catch {
    // ignore
  }
  return null;
}

function extractLatestDocxUrlFromMessages(messages: string[]): string | null {
  const urlPattern = /https?:\/\/[^\s)\]]+\.docx(?:\?[^\s)\]]*)?/gi;
  for (const message of messages) {
    const matches = message.match(urlPattern);
    if (matches?.length) {
      return matches[matches.length - 1];
    }
  }
  return null;
}

function extractLatestPdfOrDocxUrlFromMessages(messages: string[]): string | null {
  const urlPattern = /https?:\/\/[^\s)\]]+\.(?:pdf|docx)(?:\?[^\s)\]]*)?/gi;
  for (const message of messages) {
    const matches = message.match(urlPattern);
    if (matches?.length) {
      return matches[matches.length - 1];
    }
  }
  return null;
}

async function resolveLatestDocxUrlFromThread(chatThreadId: string): Promise<string | null> {
  try {
    const historyResponse = await FindTopChatMessagesForCurrentUser(chatThreadId, 20);
    if (historyResponse.status !== "OK") return null;
    const messages = historyResponse.response
      .map((message) => String(message.content ?? "").trim())
      .filter(Boolean);
    return extractLatestDocxUrlFromMessages(messages);
  } catch {
    return null;
  }
}

async function resolveLatestPdfOrDocxUrlFromThread(chatThreadId: string): Promise<string | null> {
  try {
    const historyResponse = await FindTopChatMessagesForCurrentUser(chatThreadId, 20);
    if (historyResponse.status !== "OK") return null;
    const messages = historyResponse.response
      .map((message) => String(message.content ?? "").trim())
      .filter(Boolean);
    return extractLatestPdfOrDocxUrlFromMessages(messages);
  } catch {
    return null;
  }
}

async function resolveLatestPptxUrlFromThread(chatThreadId: string): Promise<string | null> {
  try {
    const historyResponse = await FindTopChatMessagesForCurrentUser(chatThreadId, 20);
    if (historyResponse.status !== "OK") return null;
    const messages = historyResponse.response
      .map((message) => String(message.content ?? "").trim())
      .filter(Boolean);
    return extractLatestPptxUrlFromMessages(messages);
  } catch {
    return null;
  }
}

/** スレッドの最新アップロード画像URL（png/jpg/jpeg/webp等）を抽出する */
function extractLatestImageUrlFromMessages(messages: string[]): string | null {
  // file_url: ライン優先（アップロードされたファイルを示す）
  const fileUrlLineRe = /(?:^|[\n\r])(?:file_url|fileUrl)\s*:\s*(https?:\/\/[^\s\n\r]+\.(?:png|jpg|jpeg|webp|gif|bmp)(?:\?[^\s\n\r]*)?)/gi;
  const imageUrlRe = /https?:\/\/[^\s)\]]+\.(?:png|jpg|jpeg|webp|gif|bmp)(?:\?[^\s)\]]*)?/gi;
  for (const message of messages) {
    fileUrlLineRe.lastIndex = 0;
    let lastFileUrl: string | null = null;
    let m: RegExpExecArray | null;
    while ((m = fileUrlLineRe.exec(message)) !== null) { lastFileUrl = m[1]; }
    if (lastFileUrl) return lastFileUrl;
    const fallback = message.match(imageUrlRe);
    if (fallback?.length) return fallback[fallback.length - 1];
  }
  return null;
}

async function resolveLatestImageUrlFromThread(chatThreadId: string): Promise<string | null> {
  try {
    const historyResponse = await FindTopChatMessagesForCurrentUser(chatThreadId, 20);
    if (historyResponse.status !== "OK") return null;
    const messages = historyResponse.response
      .map((m) => String(m.content ?? "").trim())
      .filter(Boolean);
    return extractLatestImageUrlFromMessages(messages);
  } catch {
    return null;
  }
}

/** Markdownリンクの表示名（displayName）も含めて返す版 */
async function resolveLatestPptxInfoFromThread(chatThreadId: string): Promise<{ url: string; displayName: string | null } | null> {
  try {
    const historyResponse = await FindTopChatMessagesForCurrentUser(chatThreadId, 20);
    if (historyResponse.status !== "OK") return null;
    const messages = historyResponse.response
      .map((m) => String(m.content ?? "").trim())
      .filter(Boolean);
    return extractLatestPptxInfoFromMessages(messages);
  } catch {
    return null;
  }
}

type DeckPreferences = {
  designInstruction?: string;
  language?: "ja" | "en";
  fontScale?: "small" | "medium" | "large" | "xlarge";
  accentColor?: string;
  avoidEnglishLabels?: boolean;
  recentDesignNotes?: string[];
};

/* ------------------------------------------------------------------ */
/* NL スタイルヒント → パラメータ変換                                  */
/* ------------------------------------------------------------------ */

type StyleParams = {
  font?: string;
  size?: "small" | "medium" | "large" | "xlarge";
  sizeAdjust?: "larger" | "smaller";
  align?: "left" | "center" | "right";
  vAlign?: "top" | "middle" | "bottom";
  bottomMargin?: number;
  offsetX?: number;
  offsetY?: number;
  color?: string;
};

/** ★ スレッドごとの「直近のテキスト位置」を保持する状態 */
type TextLayout = {
  align: "left" | "center" | "right";
  vAlign: "top" | "middle" | "bottom";
  offsetX: number;
  offsetY: number;
  size: "small" | "medium" | "large" | "xlarge";
  text: string;
  color?: string;
  fontFamily?: "gothic" | "mincho" | "meiryo";
  bold?: boolean;
  italic?: boolean;
};

const lastTextLayoutByThread = new Map<string, TextLayout>();

function parseStyleHint(styleHint?: string): StyleParams {
  if (!styleHint) return {};
  const s = styleHint.replace(/\s+/g, "").toLowerCase();

  const p: StyleParams = {};

  // ---- サイズ系（絶対指定）----
  if (s.includes("特大") || s.includes("ドーン") || s.includes("めちゃ大")) {
    p.size = "xlarge";
  } else if (
    s.includes("大きめ") ||
    s.includes("大きく") ||
    s.includes("大きい")
  ) {
    p.size = "large";
  } else if (
    s.includes("小さめ") ||
    s.includes("小さい") ||
    s.includes("控えめ")
  ) {
    p.size = "small";
  } else if (s.includes("普通") || s.includes("標準")) {
    p.size = "medium";
  }

  // ★ サイズ系（相対指定）★
  if (
    s.includes("もう少し大きく") ||
    s.includes("もうちょっと大きく") ||
    s.includes("もっと大きく") ||
    s.includes("さらに大きく") ||
    s.includes("ちょい大きく")
  ) {
    p.sizeAdjust = "larger";
  } else if (
    s.includes("もう少し小さく") ||
    s.includes("もうちょっと小さく") ||
    s.includes("もっと小さく") ||
    s.includes("さらに小さく") ||
    s.includes("ちょい小さく")
  ) {
    p.sizeAdjust = "smaller";
  }

  // ---- 垂直位置（下 / 上 / 真ん中）----
  if (
    s.includes("一番下") ||
    s.includes("最下部") ||
    s.includes("フッター") ||
    s.includes("下部") ||
    s.includes("下の方") ||
    s.includes("下側")
  ) {
    p.vAlign = "bottom";
    p.bottomMargin = 80;
  }

  if (
    s.includes("一番上") ||
    s.includes("最上部") ||
    s.includes("上端") ||
    s.includes("画面の上") ||
    s.includes("上部") ||
    s.includes("上の方") ||
    s.includes("上側")
  ) {
    p.vAlign = "top";
  }

  // ★ 中央判定は最後に（他の位置指定がない場合のみ）
  if (
    !p.vAlign &&
    (s.includes("真ん中") ||
      s.includes("センター") ||
      s.includes("中心") ||
      s.includes("中央"))
  ) {
    p.vAlign = "middle";
  }

  // ---- ４隅ショートカット（水平位置より先に処理）----
  if (s.includes("左上")) {
    p.align = "left";
    p.vAlign = "top";
  }
  if (s.includes("右上")) {
    p.align = "right";
    p.vAlign = "top";
  }
  if (s.includes("左下")) {
    p.align = "left";
    p.vAlign = "bottom";
    p.bottomMargin = 80;
  }
  if (s.includes("右下")) {
    p.align = "right";
    p.vAlign = "bottom";
    p.bottomMargin = 80;
  }

  // ---- 水平位置（左 / 右 を先に、中央は最後）----
  // ★ 4隅で既に設定済みの場合はスキップ
  if (!p.align) {
    if (
      s.includes("左寄せ") ||
      s.includes("左側") ||
      s.includes("左端") ||
      (s.includes("左") && !s.includes("中央") && !s.includes("真ん中"))
    ) {
      p.align = "left";
    } else if (
      s.includes("右寄せ") ||
      s.includes("右側") ||
      s.includes("右端") ||
      (s.includes("右") && !s.includes("中央") && !s.includes("真ん中"))
    ) {
      p.align = "right";
    } else if (
      s.includes("中央") ||
      s.includes("真ん中") ||
      s.includes("センター") ||
      s.includes("中寄せ")
    ) {
      p.align = "center";
    }
  }

  // ---- 微調整（少し右 / 少し上 など）----
  if (
    s.includes("少し右") ||
    s.includes("ちょい右") ||
    s.includes("やや右")
  ) {
    p.offsetX = (p.offsetX ?? 0) + 80;
  }
  if (
    s.includes("少し左") ||
    s.includes("ちょい左") ||
    s.includes("やや左")
  ) {
    p.offsetX = (p.offsetX ?? 0) - 80;
  }
  if (
    s.includes("少し上") ||
    s.includes("ちょい上") ||
    s.includes("やや上")
  ) {
    p.offsetY = (p.offsetY ?? 0) - 60;
  }
  if (
    s.includes("少し下") ||
    s.includes("ちょい下") ||
    s.includes("やや下")
  ) {
    p.offsetY = (p.offsetY ?? 0) + 60;
  }

  // ---- 矢印による移動指定（→ ← ↑ ↓）----
  if (
    s.includes("→") ||
    s.includes("➡") ||
    s.includes("➜") ||
    s.includes("右矢印")
  ) {
    p.offsetX = (p.offsetX ?? 0) + 80;
  }
  if (s.includes("←") || s.includes("⬅") || s.includes("左矢印")) {
    p.offsetX = (p.offsetX ?? 0) - 80;
  }
  if (s.includes("↑") || s.includes("⬆") || s.includes("上矢印")) {
    p.offsetY = (p.offsetY ?? 0) - 60;
  }
  if (s.includes("↓") || s.includes("⬇") || s.includes("下矢印")) {
    p.offsetY = (p.offsetY ?? 0) + 60;
  }

  // ---- フォント ----
  if (s.includes("メイリオ")) p.font = "Meiryo";
  if (s.includes("游ゴシック") || s.includes("游ｺﾞｼｯｸ"))
    p.font = "Yu Gothic";
  if (s.includes("ゴシック")) p.font = "Yu Gothic";
  if (s.includes("明朝")) p.font = "Yu Mincho";
  if (s.includes("手書き") || s.includes("手書き風")) {
    p.font = "Comic Sans MS";
  }

  // ---- 色 ----
  if (s.includes("白文字") || s.includes("白")) p.color = "#ffffff";
  if (s.includes("黒文字") || s.includes("黒")) p.color = "#000000";
  if (s.includes("赤文字") || s.includes("赤")) p.color = "red";
  if (s.includes("青文字") || s.includes("青")) p.color = "blue";
  if (s.includes("黄色") || s.includes("黄")) p.color = "yellow";

  return p;
}

/* ------------------------------------------------------------------ */

export const GetDefaultExtensions = async (props: {
  chatThread: ChatThreadModel;
  userMessage: string;
  signal: AbortSignal;
  mode?: ThinkingModeAPI;
}): Promise<ServerActionResponse<Array<any>>> => {
  const defaultExtensions: Array<any> = [];

  const currentMode = normalizeThinkingMode(props.mode ?? "normal");
  const modeOpts = buildSendOptionsFromMode(currentMode);

  console.log("🧠 Reasoning Mode Applied:", {
    mode: currentMode,
    reasoning_effort: modeOpts.reasoning_effort,
    temperature: modeOpts.temperature,
  });

  // ★ 画像生成ツール（新しく描く用）
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeCreateImage(
          args,
          props.chatThread,
          props.signal,
          modeOpts
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          prompt: { type: "string" },
          text: { type: "string" },
          size: {
            type: "string",
            enum: ["1024x1024", "1024x1792", "1792x1024"],
          },
        },
        required: ["prompt"],
      },
      description:
        "Use this tool ONLY when user clearly asks for a NEW image to be created. " +
        "If user wants to MODIFY or add text to an ALREADY GENERATED image, you MUST NOT call this tool. " +
        "Instead, call add_text_to_existing_image with the previous image URL." +
        "After this tool returns a url, you MUST display the image using Markdown image syntax: ![image](url). Never output the URL as plain text.",
      name: "create_img",
    },
  });

  // ★ 既存画像に文字だけ足すツール（Vision を使わないシンプル版）
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeAddTextToExistingImage(
          args,
          props.chatThread,
          props.userMessage,
          props.signal,
          modeOpts
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          imageUrl: {
            type: "string",
            description:
              "URL of the existing image. If the user says 'this image', use the URL that was returned previously (for example from create_img).",
          },
          text: {
            type: "string",
            description:
              "Japanese text to overlay on the image. " +
              "CRITICAL: If the user is ONLY adjusting position, size, or color (words like '右に', 'もう少し大きく', '赤色に'), " +
              "you MUST use the EXACT same text from the previous image. Do NOT shorten, modify, or change the text content in any way.",
          },
          styleHint: {
            type: "string",
            description:
              "Natural language hint for font size, color, position such as '大きめの白文字で、下部中央に', '少し上に', '➡ で少し右へ', 'もう少し大きく', etc.",
          },
          font: {
            type: "string",
            description:
              "Font family name if explicitly requested (e.g., 'Meiryo').",
          },
          color: {
            type: "string",
            description: "Text color (e.g., 'white', '#ffffff').",
          },
          size: {
            type: "string",
            description: "Rough size hint like 'small', 'medium', 'large'.",
          },
          offsetX: {
            type: "number",
            description:
              "Horizontal offset in pixels. Positive moves text to the right, negative to the left.",
          },
          offsetY: {
            type: "number",
            description:
              "Vertical offset in pixels. Positive moves text downward, negative upward.",
          },
        },
        required: ["imageUrl", "text"],
      },
      description:
        "Use this tool when the user wants to add or adjust text on an EXISTING image, for example 'この絵に 2026 謹賀新年 と入れて' or 'もう少し下に', 'そこから➡で右に', 'もう少し大きく'. " +
        "CRITICAL RULE: When the user is ONLY requesting position/size/color adjustments, " +
        "you MUST preserve the EXACT text from the previous image without any modifications.",
      name: "add_text_to_existing_image",
    },
  });

  // ★ PowerPoint 生成ツール（テキストベース）
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeCreatePptx(args, props.chatThread, props.userMessage),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          title: {
            type: "string",
            description: "プレゼンテーション全体のタイトル",
          },
          slides: {
            type: "array",
            description:
              "スライドのリスト。\n" +
              "【重要】会社紹介・会社概要の場合は layoutType を積極的に使い分けること:\n" +
              "  - 最初の「表紙」スライドは不要（自動生成される）\n" +
              "  - 会社概要スライド → layoutType='company-overview' + metrics + leadText + callout\n" +
              "  - 強み・工程・フロー（3ステップ程度） → layoutType='process-cards' + steps + benefits\n" +
              "  - 比較・競合 → layoutType='multi-column'\n" +
              "  - お問い合わせ・次のステップ → layoutType='closing'\n" +
              "  - その他 → layoutType='bullets'（3〜4項目）\n" +
              "【提案書モード】枚数を12〜16枚に増やし課題→提案→根拠→比較→効果→ロードマップの流れで構成。",
            items: {
              type: "object",
              properties: {
                title: { type: "string", description: "スライドのタイトル" },
                bullets: {
                  type: "array",
                  items: { type: "string" },
                  description: "bullets/closing レイアウト時の内容リスト。1〜2文の具体的な記述。3〜4項目。",
                },
                layoutType: {
                  type: "string",
                  enum: ["bullets", "multi-column", "table", "diagram", "company-overview", "process-cards", "closing", "metric-cards", "timeline"],
                  description:
                    "レイアウト種別。\n" +
                    "bullets=箇条書きカード（デフォルト）\n" +
                    "company-overview=会社概要（leadText+metrics+callout 必須）\n" +
                    "metric-cards=数値KPIカード4枚（metrics 必須）\n" +
                    "process-cards=工程・プロセスフロー（steps+benefits 必須）\n" +
                    "timeline=タイムライン（steps 必須）\n" +
                    "multi-column=2〜3列比較（columns 必須）\n" +
                    "table=表形式（tableRows 必須）\n" +
                    "closing=締め・お問い合わせ（bullets使用）",
                },
                // company-overview 専用フィールド
                leadText: {
                  type: "string",
                  description: "company-overview: 左パネルに表示するリード文（会社の説明文2〜4文）",
                },
                metrics: {
                  type: "array",
                  description: "company-overview / metric-cards: 数値カード（最大4件）",
                  items: {
                    type: "object",
                    properties: {
                      label: { type: "string", description: "カードのラベル（例: '創業', '本社', '上場'）" },
                      value: { type: "string", description: "カードのメイン数値・テキスト（例: '1952年', '東証プライム'）" },
                      note: { type: "string", description: "カードの補足（例: '70年以上の実績'）" },
                      iconKey: { type: "string", description: "アイコン: calendar/location/stock/network/people/chart/building/gear/verified/star" },
                      colorRole: {
                        type: "string",
                        enum: ["primary", "accent", "neutral"],
                        description:
                          "カードの色役割。意味に基づいて設定すること（インデックス順サイクルは禁止）。\n" +
                          "primary=深緑（基本情報・所在地・設立など）\n" +
                          "accent=銅色（数値実績・上場・差別化ポイントなど強調したい項目）\n" +
                          "neutral=ダークグリーン（補足・背景情報）\n" +
                          "例: 創業→primary, 東証プライム→accent, 本社→neutral, 取引先数→accent",
                      },
                    },
                    required: ["label", "value"],
                  },
                },
                callout: {
                  type: "object",
                  description: "company-overview: 左パネル下部のコールアウトボックス（社名の由来・補足情報など）",
                  properties: {
                    title: { type: "string", description: "コールアウトのタイトル（例: '社名の由来'）" },
                    body: { type: "string", description: "コールアウトの本文" },
                  },
                  required: ["title", "body"],
                },
                // process-cards 専用フィールド
                subtitle: {
                  type: "string",
                  description: "process-cards: カード群の上に表示する説明文（1文）",
                },
                steps: {
                  type: "array",
                  description: "process-cards / timeline: 各ステップの内容（2〜4件）",
                  items: {
                    type: "object",
                    properties: {
                      title: { type: "string", description: "ステップのタイトル（例: '収集運搬'）" },
                      body: { type: "string", description: "ステップの説明文（1〜2文）" },
                      iconKey: {
                        type: "string",
                        description:
                          "アイコン識別子。必ず指定すること。\n" +
                          "廃棄物系: truck / gear / archive / shield / coins / leaf / eye\n" +
                          "汎用: building / people / chart / star / verified / lightbulb / rocket / network",
                      },
                    },
                    required: ["title", "body"],
                  },
                },
                benefits: {
                  type: "array",
                  items: { type: "string" },
                  description: "process-cards: スライド下部に表示するメリット行（2〜4項目、例: '不適正処理リスクの排除'）",
                },
                // multi-column 専用フィールド
                columns: {
                  type: "array",
                  description: "multi-column: 各列のデータ",
                  items: {
                    type: "object",
                    properties: {
                      header: { type: "string" },
                      bullets: { type: "array", items: { type: "string" } },
                    },
                    required: ["header", "bullets"],
                  },
                },
                tableRows: {
                  type: "array",
                  description: "table: 1行目ヘッダー、以降データ行",
                  items: { type: "array", items: { type: "string" } },
                },
              },
              required: ["title", "bullets"],
            },
          },
          proposalMode: {
            type: "boolean",
            description:
              "提案書モード。true にすると「1スライド1テーマ×12〜16枚構成」で、課題→提案→根拠→比較→効果→ロードマップの流れで自動展開する。" +
              "ユーザーが「提案書で」「しっかりした資料で」「営業資料として」「お客様向けに」と言った場合、または文字が少ない・内容が薄いと指摘された場合は true にすること。" +
              "【禁止】ユーザーが「7枚」「8枚」「10枚以下」など具体的な少ない枚数を指定した場合は false にすること（指定枚数を優先）。",
          },
          fontFace: {
            type: "string",
            description: "PowerPointで使うフォント名。例: 'Meiryo', 'Yu Gothic', 'Yu Mincho'",
          },
          designInstruction: {
            type: "string",
            description:
              "デザイン・色調の指示。業種感を必ず含めること。\n" +
              "【廃棄物処理・環境・インフラ・サステナ系】→ '廃棄物処理・環境配慮・信頼感をテーマに、深緑ベースの落ち着いたデザイン。会社紹介資料' のようにキーワード(廃棄物/環境/産廃)を含めること。\n" +
              "例: '医療・製薬向けの清潔感ある白と青', 'IT・DX提案書らしいモダンなグラデーション', '廃棄物処理業の信頼感・環境意識を表現した深緑テーマ'",
          },
          palette: {
            type: "string",
            enum: ["navy_orange", "forest_amber", "burgundy_gold", "teal_coral", "charcoal_terra"],
            description:
              "【カラーパレット選択】コンテンツの業種・用途・ターゲット感から必ず判断して設定すること。\n" +
              "  navy_orange   = 紺×オレンジ → IT・AI・DX・経営・役員・システム・テクノロジー企業（落ち着いたプロ感）\n" +
              "  forest_amber  = 深緑×琥珀  → 採用・人材募集・インターン・新卒リクルート・人の成長・農業・食品・エコ\n" +
              "    ↑「人が育つ・生命感・成長」イメージ → 採用/研修/インターン系はこれ\n" +
              "  burgundy_gold = 深赤×金    → 伝統・高級・老舗・製造業・工業・ものづくり・品質重視\n" +
              "  teal_coral    = 青緑×珊瑚  → 産廃・廃棄物処理・リサイクル・医療・ヘルス・動的な産業系企業\n" +
              "    ↑廃棄物処理業・環境サービス会社の会社紹介はこれ（会社の動的でモダンな印象）\n" +
              "  charcoal_terra= 炭×煉瓦   → 建設・土木・インフラ・重工業・プラント・施設管理\n" +
              "【判断例】\n" +
              "  産廃会社の会社紹介 → teal_coral\n" +
              "  DX人材採用・インターン募集 → forest_amber\n" +
              "  AzureChat/AI/DX経営報告 → navy_orange\n" +
              "  廃棄物処理施設・プラント建設 → charcoal_terra",
          },
        },
        required: ["title", "slides"],
      },
      description:
        "ユーザーがテーマや内容を指定してPowerPoint（PPTX）を新規作成するツール。\n" +
        "テキストベースでスライド構成を作る場合に使用する。\n" +
        "【最重要・ツール選択ルール】\n" +
        "・PDFをそのままPPTに変換する場合 → convert_doc_to_pptx を使うこと。\n" +
        "・会話で既にスライド構成を議論済みで、PDFは参考資料として内容を拡充・追記する場合 → このツール（create_pptx）を使うこと。\n" +
        "  この場合、まず sl_doc_search や会話コンテキストでPDF内容を把握し、前の会話のスライド構成をベースに各スライドの bullets を肉付けした上で slides パラメータに設定して呼ぶこと。\n" +
        "【提案書モード】ユーザーが「提案書」「営業資料」「お客様向け」「しっかりした資料」と言った場合は proposalMode=true にして、12〜16枚構成で作ること。\n" +
        "【経営向け再構築モード】複数の定期レポートや四半期報告書（例: Q1〜Q4 議事録・活動報告PDF）から経営層・役員向けPPTを作る場合：\n" +
        "  ① slides パラメータを時系列（Q1→Q4）で組まないこと。以下の9カテゴリで構成すること:\n" +
        "    1. 目的・位置づけ（なぜこのツール/施策が必要か）\n" +
        "    2. 現在使える主な機能（ビジネス機能として整理。技術仕様でなく「何ができるか」「何の業務に使えるか」）\n" +
        "    3. 利用状況・KPI・運用実績（アクティブ率・件数・満足度などの数値。四半期をまたぐ場合はトレンドを統合）\n" +
        "    4. 拡張・連携状況（SharePoint検索、RAG、Salesforce、議事郎連携など。議事郎は独立スライド不可、ここに統合）\n" +
        "    5. セキュリティ・ガバナンス・運用基盤\n" +
        "    6. コスト・投資対効果（費用・ROI・削減効果）\n" +
        "    7. 課題・リスク・改善要望\n" +
        "    8. 今後のロードマップ\n" +
        "    9. 経営判断が必要な論点（意思決定を促す締めスライド）\n" +
        "  ② 各カテゴリのbulletsは、全ての参照ドキュメントから関連情報を集約・統合して記述すること。\n" +
        "  ③ スライドタイトルに「Q1」「Q2」「Q3」「Q4」「第1四半期」などの時系列ラベルを含めないこと。\n" +
        "【重要】会話中にすでにPPTXが生成・編集された実績がある場合、色・デザイン・テキスト変更・ロゴ追加・画像追加・添付画像挿入はすべて edit_pptx を使うこと。このツールは完全新規作成専用。\n" +
        "【禁止】会話中にPPTXリンクが存在する状態で「ロゴを追加して」「画像を入れて」「添付を表紙に」などと言われた場合、絶対にこのツールを使わないこと。\n" +
        "【palette 選択】ユーザーの業種・用途・ターゲット層を読み取り、必ず palette を設定すること。\n" +
        "  IT/AI/DX/経営/役員向け → navy_orange\n" +
        "  採用・人材募集・インターン・新卒向け → forest_amber（人の成長・緑のイメージ）\n" +
        "  産廃・廃棄物処理・リサイクル・環境サービス → teal_coral（動的な産業系）\n" +
        "  伝統・製造・老舗 → burgundy_gold、建設・土木・インフラ → charcoal_terra\n" +
        "ユーザーが業種・用途を言及した場合は designInstruction に業種感を含めること。\n" +
        "【重要】会社紹介・提案書の場合、slides の bullets には [会社名] [設立年] 等のプレースホルダーを使わず、知っている限りの具体的な情報を入れること（ツール実行時に自動でWeb検索して補完される）。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式でユーザーに提示すること。リンクテキストは displayName フィールドを使うこと（例: [ミダック会社紹介.pptx](downloadUrl)）。",
      name: "create_pptx",
    },
  });

  // ★ ドキュメント（PDF・画像）→ PPTX 変換ツール（Vision API使用・高精度）
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeConvertDocToPptx(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileUrl: {
            type: "string",
            description:
              "変換するファイルのURL（Azure BlobのURL）。会話コンテキストの file_url または fileUrl から取得すること。",
          },
          fileUrls: {
            type: "array",
            items: { type: "string" },
            description:
              "追加で取り込む複数ファイルURLの配列。fileUrl と一緒に渡すと、1つのPPTにまとめて生成する。",
          },
          presentationTitle: {
            type: "string",
            description:
              "プレゼンテーション全体のタイトル（省略可能、省略時はファイル名から自動設定）",
          },
          fontFace: {
            type: "string",
            description: "PowerPointで使うフォント名。例: 'Meiryo', 'Yu Gothic', 'Yu Mincho'",
          },
          designInstruction: {
            type: "string",
            description:
              "ユーザーの自然言語指示を反映してPPTのLook&Feelingを整えるための自然言語指示。例: 'ecoで洗練された役員向け' 'ポップで親しみやすく図解多め' '高級感のある提案書トーン'",
          },
          maxPages: {
            type: "number",
            description: "変換する最大ページ数（省略可能、デフォルト30）",
          },
          mode: {
            type: "string",
            enum: ["faithful", "redesign"],
            description:
              "変換モード。'faithful'=忠実変換（元ページ数維持・自動タイトルスライドなし・デザインAI最小化）。" +
              "「そのまま」「忠実に」「原本に近く」「ページ数を変えずに」などの場合は 'faithful' を指定。" +
              "デフォルトは 'redesign'（デザイン自動改善）。",
          },
        },
        required: [],
      },
      description:
        "ユーザーがアップロードしたPDF・画像ファイルをPowerPoint（PPTX）に変換するツール。\n" +
        "Vision APIを使って各ページを視覚的に解析するため、グラフ・表・図も含めて高精度に変換できる。\n" +
        "使用タイミング：ユーザーが「PPTに変換して」「スライドにして」「PPT化して」と言い、かつ会話コンテキストにfile_urlがある場合。\n" +
        "【禁止】会話で既にスライド構成を議論済みで、PDFは参考資料として内容を拡充・追記するだけの場合は、このツールを使わないこと。その場合は create_pptx を使うこと。\n" +
        "【重要】fileUrlは必ず会話コンテキストの 'file_url:' または 'fileUrl:' で始まる行から取得すること（blob.core.windows.net のURLを優先）。\n" +
        "検索結果の引用（citation本文中）に含まれるSharePointのリンクは使わないこと。'file_url:' 行から得たBlobURLであれば使ってよい。\n" +
        "「そのまま変換」「忠実に変換」「原本に近く」など正確な再現が求められる場合は mode='faithful' を指定すること。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "convert_doc_to_pptx",
    },
  });

  // ★ SharePoint SL文書をPPTに変換するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeConvertSpToPptx(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileQuery: {
            type: "string",
            description: "変換したいSharePointファイルの名前またはキーワード。例: '営業資料2024.pdf'",
          },
          mode: {
            type: "string",
            enum: ["faithful", "redesign"],
            description:
              "変換モード。'faithful'=忠実変換（ページ数維持）。'redesign'=デザイン自動改善（デフォルト）。",
          },
        },
        required: ["fileQuery"],
      },
      description:
        "SharePointのSLライブラリにある文書（PDF）をPowerPoint（PPTX）に変換するツール。\n" +
        "使用タイミング：会話コンテキストに file_url が存在しない状態で、ユーザーがSP/SLの資料名を挙げてPPT変換を求めた場合。\n" +
        "例: 「SPの営業資料2024.pdfをPPTにして」「SLにある〇〇をスライドにして」\n" +
        "【重要】会話コンテキストに file_url が既にある場合は convert_doc_to_pptx を使うこと（このツールは不要）。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。\n" +
        "複数候補がある場合はリストを提示してユーザーに選ばせること。",
      name: "convert_sp_to_pptx",
    },
  });

  // ★ 既存 PPTX を指示に従って改良するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeEditPptx(
          {
            ...args,
            fileUrl:
              String(args?.fileUrl ?? "").trim() ||
              (await resolveLatestPptxUrlFromThread(props.chatThread.id)) ||
              "",
          },
          props.chatThread
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileUrl: {
            type: "string",
            description:
              "編集対象のPPTXファイルのURL（省略可能）。省略した場合はこのスレッドで直近に生成・編集したPPTXを自動的に使用する。",
          },
          instruction: {
            type: "string",
            description:
              "ユーザーの編集指示。例: '色を青に変えて', 'フォントを游ゴシックに', '全体のトーンを力強く', '3枚目のタイトルをXXXに変えて', 'ロゴを追加して', '表紙に画像を追加'",
          },
          imageUrl: {
            type: "string",
            description:
              "挿入する画像のURL。会話コンテキストに 'file_url:' で始まる画像（png/jpg/jpeg/webp等）がある場合、そのURLをここに設定すること。ロゴ・添付画像挿入の場合は必須。DALL-Eで生成しないこと。",
          },
        },
        required: ["instruction"],
      },
      description:
        "このスレッドで生成・編集した既存PPTXを自然言語の指示に従って改良するツール。\n" +
        "【絶対ルール】会話中にPPTXが生成・編集された実績がある場合は、必ずこのツールを使うこと。create_pptx / convert_doc_to_pptx は使わないこと。\n" +
        "【最優先ケース】以下は必ずこのツールを使う：\n" +
        "- 「ロゴを追加して」「画像を追加して」「添付画像を入れて」「表紙にロゴを入れて」など画像・ロゴ挿入\n" +
        "- 「色を変えて」「緑にして」「赤くして」「青にして」などの色変更\n" +
        "- 「フォントを変えて」「もっとポップに」などデザイン変更\n" +
        "- 「〜に変えて」「〜を修正して」などテキスト編集\n" +
        "【imageUrl】ユーザーが画像をアップロードしている場合（会話コンテキストの file_url: 行に png/jpg/webp のURL）、imageUrl にそのURLを必ず設定すること。\n" +
        "fileUrlは省略可（スレッド内の直近PPTXを自動取得）。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式でユーザーに提示すること。リンクテキストは displayName フィールドを使うこと（例: [AzureChat機能紹介_ロゴ追加.pptx](downloadUrl)）。",
      name: "edit_pptx",
    },
  });

  // ★ SharePoint SL の PPTX を指示に従って編集するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) => await executeEditSpPptx(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileQuery: {
            type: "string",
            description: "編集したいSharePointのPPTXファイルの名前またはキーワード。例: '営業資料2024.pptx'",
          },
          instruction: {
            type: "string",
            description: "編集指示。例: 'Matrix映画風の色味に変えて'、'フォントを游ゴシックに'、'表紙のタイトルをXXXに変更して'",
          },
        },
        required: ["fileQuery", "instruction"],
      },
      description:
        "SharePointのSLライブラリにあるPPTXファイルを自然言語の指示に従って編集するツール。\n" +
        "使用タイミング：ユーザーがSP/SL上のPPTXの色・フォント・テキストを変更したい場合。\n" +
        "例: 「SPにある営業資料をMatrix風の色にして」「SLの〇〇.pptxのフォントを変えて」\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "edit_sp_pptx",
    },
  });

  // ★ SharePoint SL の Excel ファイルを編集するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) => await executeEditSpExcel(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileQuery: {
            type: "string",
            description: "編集したいSharePointのExcelファイルの名前またはキーワード。例: '売上データ.xlsx'",
          },
          instruction: {
            type: "string",
            description: "編集指示。例: '折れ線グラフを作成して'、'棒グラフにして'、'縦軸を千円単位にして'",
          },
          previousChartEdits: {
            type: "array",
            description:
              "【グラフ修正時は必須】直前の edit_sp_excel / edit_excel のtool結果に含まれる appliedChartEdits の値をそのまま渡すこと。前回のグラフ設定が引き継がれ、指定した項目だけ変更される。",
            items: { type: "object" },
          },
        },
        required: ["fileQuery", "instruction"],
      },
      description:
        "SharePointのSLライブラリにあるExcelファイル（.xlsx/.xls/.xlsm）を自然言語の指示に従って編集するツール。\n" +
        "使用タイミング：ユーザーがSP/SL上のExcelのグラフ作成・セル編集・書式変更などを求める場合。\n" +
        "例: 「SPにある売上データ.xlsxをグラフ化して」「SLの〇〇.xlsxに折れ線グラフを追加して」\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "edit_sp_excel",
    },
  });

  // ★ テキスト・表データから Excel ファイルを新規作成するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) => await executeCreateExcel(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          content: {
            type: "string",
            description:
              "Excelに出力するデータ全文。テキスト・表・数値をそのまま渡す。タブ区切り・CSV・箇条書き等いずれでも可。",
          },
          title: {
            type: "string",
            description:
              "ブック/シートのタイトル。省略時はcontentから自動推定する。",
          },
          instruction: {
            type: "string",
            description:
              "書式・構成の指示。例: '1行目をヘッダーにして' '複数シートに分けて' '合計行を追加して'",
          },
        },
        required: ["content"],
      },
      description:
        "ユーザーが指定したテキストや表データからExcelファイル（.xlsx）を新規作成するツール。\n" +
        "使用タイミング：ユーザーが「Excelにして」「Excelで出力して」「表をExcelにして」「xlsx にして」と言い、かつアップロードファイルがない場合。\n" +
        "既存Excelファイルの編集は edit_excel ツールを使うこと（このツールは新規作成専用）。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "create_excel",
    },
  });

  // ★ アップロードされた Excel ファイルを指示に従って編集するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeEditExcel(
          {
            ...args,
            fileUrl:
              String(args?.fileUrl ?? "").trim() ||
              (await resolveLatestXlsxUrlFromThread(props.chatThread.id)) ||
              "",
          },
          props.chatThread
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileUrl: {
            type: "string",
            description:
              "編集対象のExcelファイルのURL。アップロードまたはこのスレッドで作成されたxlsx/xls/xlsmのURL。省略時はスレッド内の最新Excelを自動解決する。",
          },
          instruction: {
            type: "string",
            description:
              "ユーザーの編集指示。例: 'A1セルを「売上合計」に変えて', '1行目を太字・背景色を青に', '「旧社名」を「新社名」に置換して', '折れ線グラフを作成してシート内に追加して', '棒グラフにして', '棒を赤に'。",
          },
          previousChartEdits: {
            type: "array",
            description:
              "【グラフ修正時は必須】直前の edit_excel / edit_sp_excel のtool結果に含まれる appliedChartEdits の値をそのまま渡すこと。" +
              "これにより前回のグラフ設定（chartType・title・yDivisor・seriesColors等）が自動的に引き継がれ、指定した項目だけ変更される。" +
              "グラフを新規作成する場合は省略してよい。",
            items: { type: "object" },
          },
        },
        required: ["instruction"],
      },
      description:
        "このスレッドのExcelファイル（アップロードまたはcreate_excelで作成）を自然言語の指示に従って編集するツール。\n" +
        "使用タイミング：ExcelファイルへのセルA値変更・テキスト置換・書式変更（太字・色・罫線・枠・border）・整形・見やすくする・グラフ作成/修正（折れ線グラフ・棒グラフ・散布図・円グラフ・チャート・タイトル変更・縦軸/横軸ラベル変更・単位変更・目盛調整）等を求める場合。\n" +
        "重要：グラフ・縦軸・横軸・単位に関する指示は必ずこのツールで処理すること。「画像なので数値が読めない」は誤り — このツールがExcelの元データを直接読み取る。\n" +
        "fileUrl が省略された場合はスレッド内の最新Excelを自動的に使用する。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "edit_excel",
    },
  });

  // ★ テキストから Word ファイルを新規作成するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) => await executeCreateWord(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          content: {
            type: "string",
            description:
              "Wordドキュメントに記載するテキスト全文。ユーザーが指定した内容をそのまま渡す。",
          },
          title: {
            type: "string",
            description:
              "ドキュメントのタイトル。省略時はcontentから自動推定する。",
          },
          instruction: {
            type: "string",
            description:
              "書式・スタイルの指示。例: '見出しを使って整理して' '箇条書きにして' '表形式でまとめて'",
          },
          fontFace: {
            type: "string",
            description: "使用フォント名。例: 'Meiryo', 'Yu Gothic', 'Yu Mincho'（省略時: Meiryo）",
          },
        },
        required: ["content"],
      },
      description:
        "ユーザーが指定したテキストや内容からWordファイル（.docx）を新規作成するツール。\n" +
        "使用タイミング：ユーザーが「Wordにして」「Wordで作って」「Word文書を作成して」「docxにして」と言った場合。\n" +
        "既存Wordファイルの編集は edit_word ツールを使うこと（このツールは新規作成専用）。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "create_word",
    },
  });

  // ★ アップロードされた Word ファイルを指示に従って編集するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeEditWord(
          {
            ...args,
            fileUrl:
              String(args?.fileUrl ?? "").trim() ||
              (await resolveLatestDocxUrlFromThread(props.chatThread.id)) ||
              "",
          },
          props.chatThread
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileUrl: {
            type: "string",
            description:
              "編集対象のWordファイルのURL。アップロードまたはこのスレッドで作成された.docxのURL。省略時はスレッド内の最新Wordを自動解決する。",
          },
          instruction: {
            type: "string",
            description:
              "ユーザーの編集指示。例: '「旧社名」を「新社名」に置換して', 'タイトルを太字・赤色にして', '第1章の見出しを16ptにして'",
          },
        },
        required: ["instruction"],
      },
      description:
        "このスレッドのWordファイル（アップロードまたはcreate_wordで作成）を自然言語の指示に従って編集するツール。\n" +
        "使用タイミング：Wordファイルへのテキスト置換・書式変更（太字・色・フォントサイズ）を求める場合。\n" +
        "fileUrl が省略された場合はスレッド内の最新Wordを自動的に使用する。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "edit_word",
    },
  });

  // ★ アップロードされた PDF ファイルを Word に変換するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) => await executeConvertPdfToWord(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileUrl: {
            type: "string",
            description:
              "変換対象のPDFファイルのURL。このスレッドでアップロードされた.pdfのURLを指定する。",
          },
          mode: {
            type: "string",
            enum: ["layout", "editable"],
            description:
              "layout: 見た目・レイアウト再現優先（pdf2docx使用）。editable: テキスト・表を編集可能な形で抽出優先（Doc Intelligence使用）。",
          },
        },
        required: ["fileUrl"],
      },
      description:
        "このスレッドでアップロードされたPDFファイルをWord（.docx）に変換するツール。\n" +
        "使用タイミング：ユーザーがPDFをWordに変換したいと言った場合。\n" +
        "mode=layout: 「WordにしてWordに変換して」など見た目重視の場合。\n" +
        "mode=editable: 「編集可能なWordに」「表を編集できるWordに」「テキストとして抽出」など編集重視の場合。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "convert_pdf_to_word",
    },
  });

  // ★ アップロードされた PDF ファイルを Excel に変換するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) =>
        await executeConvertPdfToExcel(
          {
            ...args,
            fileUrl:
              String(args?.fileUrl ?? "").trim() ||
              (await resolveLatestPdfOrDocxUrlFromThread(props.chatThread.id)) ||
              "",
          },
          props.chatThread
        ),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileUrl: {
            type: "string",
            description:
              "変換対象のPDF/WordファイルのURL。このスレッドでアップロードされた.pdf/.docxのURL。省略時はスレッド内の最新PDF/Wordを自動解決する。",
          },
        },
        required: [],
      },
      description:
        "このスレッドでアップロードされたPDFまたはWord（.docx）ファイルをExcel（.xlsx）に変換するツール。\n" +
        "使用タイミング：ユーザーがPDF/WordをExcelに変換したいと言った場合。\n" +
        "fileUrl は省略可能。省略するとスレッド内の最新PDF/Wordを自動的に使用する。\n" +
        "テーブルはシートに、テーブルがない場合はテキストを「Text」シートに出力する。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "convert_pdf_to_excel",
    },
  });

  return { status: "OK", response: defaultExtensions };
};

// ---------------- SP文書検索（提案書コンテキスト） ----------------

/**
 * 提案書生成前に AI Search（SharePoint文書）を複数クエリで検索し、
 * 参照可能な社内文書のテキストをまとめて返す。
 * LLMの事前学習知識ではなく、実際のSP文書を提案内容に反映させるための関数。
 */
async function fetchSpContextForProposal(
  topic: string,
  inputSlides: Array<{ title: string; bullets: string[] }>,
  deptLower: string
): Promise<string> {
  try {
    // タイトル + 各スライドタイトルから検索クエリを生成（最大4クエリ）
    const queries = [topic, ...inputSlides.map((s) => s.title)]
      .filter(Boolean)
      .slice(0, 4);

    const seen = new Set<string>();
    const excerpts: string[] = [];

    for (const query of queries) {
      const result = await SimilaritySearch(query, 6, "isSlDoc eq true", deptLower);
      if (result.status !== "OK") continue;

      for (const item of result.response) {
        const content = item.document.pageContent?.trim();
        const source = item.document.metadata || "";
        if (!content || seen.has(content)) continue;
        seen.add(content);
        // 1件あたり最大600文字に切り詰めて過大なトークン消費を防ぐ
        excerpts.push(`【出典: ${source}】\n${content.slice(0, 600)}`);
      }
    }

    console.log(`[proposalMode] SP文書取得: ${excerpts.length}件 (queries=${queries.length})`);
    return excerpts.slice(0, 15).join("\n\n---\n\n");
  } catch (e) {
    console.warn("[proposalMode] fetchSpContextForProposal failed:", e);
    return "";
  }
}

// ─── PromptIntent: ユーザー意図の構造化 ─────────────────────────────────────

type PromptIntentLocal = {
  documentPurpose: "proposal"|"company-intro"|"recruitment"|"training"|"analysis"|"internal"|"ir"|"campaign"|"other";
  audience: "executive"|"customer"|"employee"|"candidate"|"general";
  designFreedom: "conservative"|"balanced"|"expressive";
  toneKeywords: string[];
  colorDirectives?: { primary?: string; accent?: string; background?: string };
  layoutDirectives: { preferTwoColumn?: boolean; includeTables?: boolean; avoidBulletOnly?: boolean; preferMetrics?: boolean; preferProcess?: boolean };
  styleGuardrails: { allowModernDark?: boolean; allowPlayful?: boolean; allowGlass?: boolean; maxAccentIntensity?: "low"|"medium"|"high" };
};

function parsePromptIntent(text: string): PromptIntentLocal {
  const h = text.toLowerCase();
  const has = (...words: string[]) => words.some((w) => h.includes(w));

  // documentPurpose
  let documentPurpose: PromptIntentLocal["documentPurpose"] = "other";
  if      (has("採用","recruit","人材","求人","hiring"))                         documentPurpose = "recruitment";
  else if (has("キャンペーン","イベント","告知","campaign","event"))              documentPurpose = "campaign";
  else if (has("提案","proposal","営業提案"))                                    documentPurpose = "proposal";
  else if (has("会社紹介","会社概要","初回訪問","company profile","紹介資料"))   documentPurpose = "company-intro";
  else if (has("研修","training","教育","onboard","オンボード"))                 documentPurpose = "training";
  else if (has("分析","調査","市場","analysis","リサーチ","research"))           documentPurpose = "analysis";
  else if (has("ir ","ir、","ir。","投資家","株主","決算","investor"))            documentPurpose = "ir";
  else if (has("社内","internal","報告","レポート"))                             documentPurpose = "internal";
  else if (has("営業","商談","提案"))                                            documentPurpose = "proposal";

  // audience
  let audience: PromptIntentLocal["audience"] = "general";
  if      (has("役員","経営層","executive","board","ceo","社長"))    audience = "executive";
  else if (has("顧客","お客様","customer","クライアント","取引先"))  audience = "customer";
  else if (has("候補者","求職者","candidate","job seeker"))          audience = "candidate";
  else if (has("社員","employee","スタッフ","従業員","メンバー"))    audience = "employee";

  // designFreedom
  const isExpressive = has("fancy","華やか","かっこよく","インパクト","bold","個性的","派手","モダン","creative");
  const isConservative = has("上品","信頼感","堅め","堅実","営業向け","シンプル","落ち着い","フォーマル","品よく");
  let designFreedom: PromptIntentLocal["designFreedom"] = "balanced";
  if (isExpressive && !isConservative) designFreedom = "expressive";
  else if (isConservative)             designFreedom = "conservative";
  // guardrail: proposal/ir/executive + expressive → balanced
  if (designFreedom === "expressive" && (documentPurpose === "proposal" || documentPurpose === "ir" || audience === "executive")) {
    designFreedom = "balanced";
  }

  // toneKeywords
  const toneKeywords = ["fancy","モダン","エレガント","bold","上品","信頼感","親しみ","明るい","シンプル","クール","professional","minimal","impactful"]
    .filter((kw) => h.includes(kw));

  // colorDirectives: HEX (#RRGGBB or RRGGBB) → 最初の2つ
  const hexMatches = Array.from(text.matchAll(/#?([0-9A-Fa-f]{6})\b/g));
  const colorMapping: Record<string, string> = {
    "ネイビー":"0B2540","navy":"0B2540","紺":"0B3060",
    "オレンジ":"F97316","orange":"F97316",
    "青":"2563EB","ブルー":"2563EB","blue":"2563EB",
    "赤":"DC2626","red":"DC2626",
    "緑":"16A34A","グリーン":"16A34A","green":"16A34A",
    "黄":"EAB308","yellow":"EAB308",
    "黒":"0F172A","ブラック":"0F172A","black":"0F172A",
    "白":"F8FAFC","white":"F8FAFC",
    "グレー":"6B7280","gray":"6B7280","grey":"6B7280",
    "紫":"7C3AED","パープル":"7C3AED","purple":"7C3AED","violet":"7C3AED",
    "ピンク":"EC4899","pink":"EC4899",
    "ティール":"0D9488","teal":"0D9488","水色":"38BDF8",
    "インディゴ":"4F46E5","indigo":"4F46E5",
  };

  const colorDirectives: PromptIntentLocal["colorDirectives"] = {};
  // HEX 優先
  if (hexMatches.length >= 1) colorDirectives.primary = hexMatches[0][1].toUpperCase();
  if (hexMatches.length >= 2) colorDirectives.accent  = hexMatches[1][1].toUpperCase();
  // カラーワードで補完
  let foundPrimary = Boolean(colorDirectives.primary);
  for (const [word, hex] of Object.entries(colorMapping)) {
    if (!h.includes(word.toLowerCase())) continue;
    if (!foundPrimary) { colorDirectives.primary = hex; foundPrimary = true; }
    else if (!colorDirectives.accent) { colorDirectives.accent = hex; break; }
  }

  // layoutDirectives
  const layoutDirectives: PromptIntentLocal["layoutDirectives"] = {
    preferTwoColumn: has("2列","二列","左右","比較","two column","two-column","サイドバイサイド"),
    includeTables:   has("表","テーブル","一覧表","比較表","table","matrix"),
    avoidBulletOnly: has("箇条書きだけにしない","単調にしない","バリエーション","メリハリ","変化","飽きない"),
    preferMetrics:   has("数値","kpi","実績","指標","metric","定量","数字","数"),
    preferProcess:   has("手順","流れ","プロセス","ステップ","process","step","工程","フロー"),
  };

  // styleGuardrails
  const styleGuardrails: PromptIntentLocal["styleGuardrails"] = {
    allowModernDark: designFreedom === "expressive" || has("dark","モダンダーク","黒","black","ダーク"),
    allowPlayful:    designFreedom === "expressive" && !["proposal","ir","company-intro"].includes(documentPurpose),
    allowGlass:      designFreedom !== "conservative",
    maxAccentIntensity: designFreedom === "conservative" ? "low" : designFreedom === "expressive" ? "high" : "medium",
  };

  return {
    documentPurpose, audience, designFreedom, toneKeywords,
    colorDirectives: Object.keys(colorDirectives).length > 0 ? colorDirectives : undefined,
    layoutDirectives,
    styleGuardrails,
  };
}

// ---------------- BraveSearch + スライド補完 ----------------

async function searchBrave(query: string): Promise<string> {
  const apiKey = process.env.BRAVE_SUBSCRIPTION_TOKEN;
  if (!apiKey) return "";
  const controller = new AbortController();
  const tid = setTimeout(() => controller.abort(), 6000);
  try {
    const params = new URLSearchParams({ q: query, count: "5" });
    const res = await fetch(
      `https://api.search.brave.com/res/v1/web/search?${params.toString()}`,
      {
        headers: {
          "Accept": "application/json",
          "X-Subscription-Token": apiKey,
        },
        signal: controller.signal,
      }
    );
    clearTimeout(tid);
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      console.warn("[brave-search] HTTP", res.status, "query:", query, "body:", body.slice(0, 200));
      return "";
    }
    const data = await res.json();
    const results = (data.web?.results ?? []) as Array<{
      title?: string;
      description?: string;
      extra_snippets?: string[];
    }>;
    const text = results
      .slice(0, 5)
      .map((r) => {
        const snippets = (r.extra_snippets ?? []).slice(0, 2).join(" ");
        return `【${r.title ?? ""}】${r.description ?? ""} ${snippets}`.trim();
      })
      .filter(Boolean)
      .join("\n");
    console.log(`[brave-search] OK: ${results.length}件 query="${query}"`);
    return text.slice(0, 3500);
  } catch (e: any) {
    clearTimeout(tid);
    console.warn("[brave-search] failed (query:", query, "):", e?.message ?? e);
    return "";
  }
}

// ---- HTMLページ本文取得 ----
async function fetchPageText(url: string, maxChars = 3000): Promise<string> {
  const controller = new AbortController();
  const tid = setTimeout(() => controller.abort(), 5000);
  try {
    const res = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml",
        "Accept-Language": "ja,en;q=0.9",
      },
      signal: controller.signal,
    });
    clearTimeout(tid);
    const ct = res.headers.get("content-type") ?? "";
    if (!res.ok || !ct.includes("text/html")) return "";
    const html = await res.text();
    const text = html
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<[^>]+>/g, " ")
      .replace(/&[a-z#0-9]+;/gi, " ")
      .replace(/\s{2,}/g, " ")
      .trim();
    return text.slice(0, maxChars);
  } catch {
    clearTimeout(tid);
    return "";
  }
}

// ---- Brave検索 + ページ本文収集 ----
type BraveWebEvidence = { snippets: string; pages: string };

async function collectWebEvidence(query: string): Promise<BraveWebEvidence> {
  const apiKey = process.env.BRAVE_SUBSCRIPTION_TOKEN;
  if (!apiKey) return { snippets: "", pages: "" };

  const controller = new AbortController();
  const tid = setTimeout(() => controller.abort(), 8000);
  let braveData: any = null;
  try {
    const params = new URLSearchParams({ q: query, count: "8" });
    const res = await fetch(`https://api.search.brave.com/res/v1/web/search?${params}`, {
      headers: { Accept: "application/json", "X-Subscription-Token": apiKey },
      signal: controller.signal,
    });
    clearTimeout(tid);
    if (res.ok) braveData = await res.json();
  } catch {
    clearTimeout(tid);
  }

  const results: Array<{ title?: string; description?: string; extra_snippets?: string[]; url?: string }> =
    braveData?.web?.results ?? [];

  const snippets = results
    .slice(0, 8)
    .map((r) => {
      const extras = (r.extra_snippets ?? []).slice(0, 3).join(" ");
      return `【${r.title ?? ""}】${r.description ?? ""} ${extras}`.trim();
    })
    .filter(Boolean)
    .join("\n")
    .slice(0, 4000);

  const candidateUrls = results
    .slice(0, 5)
    .map((r) => r.url ?? "")
    .filter((u) => u.startsWith("http"));

  const pageTexts = await Promise.allSettled(
    candidateUrls.slice(0, 4).map((url) => fetchPageText(url, 2500))
  );

  const pages = pageTexts
    .filter((r): r is PromiseFulfilledResult<string> => r.status === "fulfilled" && Boolean(r.value))
    .map((r) => r.value)
    .join("\n---\n")
    .slice(0, 8000);

  console.log(`[collectWebEvidence] query="${query}" snippets=${snippets.length}c pages=${pages.length}c`);
  return { snippets, pages };
}

// ---- LLM事実抽出 ----
type CompanyFacts = {
  companyName: string;
  industry: string;
  business: string[];
  strengths: string[];
  metrics: Array<{ label: string; value: string; note?: string }>;
  cautions: string[];
};

// ---- 会社紹介用中間ブリーフ（Web本文 → 用途別構造化JSON） ----
type CompanyBrief = {
  companyName: string;
  audience: string;
  purpose: string;
  companyOverview: string;
  businessAreas: string[];
  serviceFlow: Array<{ title: string; body: string }>;
  strengths: string[];
  metrics: Array<{ label: string; value: string; note?: string }>;
  proofPoints: string[];
  recommendedSlideOutline: Array<{ slideTitle: string; layoutType: string; keyConcept: string }>;
};

function cleanWebText(raw: string): string {
  return raw
    .replace(/https?:\/\/\S+/g, "")
    .replace(/[\t\r]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ 　]{3,}/g, " ")
    .replace(/[^\S\n]{4,}/g, " ")
    .replace(/(\d[\d,. ]{20,})/g, "")
    .trim();
}

async function extractCompanyFacts(
  companyName: string,
  evidence: BraveWebEvidence
): Promise<CompanyFacts> {
  const empty: CompanyFacts = {
    companyName,
    industry: "",
    business: [],
    strengths: [],
    metrics: [],
    cautions: [],
  };

  const rawCombined = [evidence.snippets, evidence.pages].filter(Boolean).join("\n\n");
  if (!rawCombined) return empty;
  const combined = cleanWebText(rawCombined).slice(0, 5000);

  try {
    const openai = OpenAIInstance();
    const completion = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME ?? "",
      max_completion_tokens: 3000,
      response_format: { type: "json_object" } as const,
      messages: [
        {
          role: "system",
          content:
            `You are a fact extractor. Extract facts about "${companyName}" from the web text below. ` +
            `Fill in this exact JSON structure (omit array items you cannot verify): ` +
            `{"companyName":"${companyName}","industry":"","business":[""],"strengths":[""],"metrics":[{"label":"創業","value":"","note":""},{"label":"本社","value":"","note":""},{"label":"上場","value":"","note":""},{"label":"従業員","value":"","note":""}],"cautions":[]} ` +
            `Important: if a fact is clearly mentioned, include it. Do not leave everything empty. Output JSON only.`,
        },
        {
          role: "user",
          content: `Company: ${companyName}\n\nWeb text:\n${combined}`,
        },
      ],
    });

    const raw = completion.choices[0]?.message?.content ?? "";
    console.log("[extractCompanyFacts] raw:", raw.slice(0, 1000));
    const stripped = raw.replace(/```(?:json)?/gi, "").replace(/```/g, "");
    const match = stripped.match(/\{[\s\S]*\}/);
    if (!match) {
      console.warn("[extractCompanyFacts] no JSON object found in response");
      return empty;
    }
    const parsed = JSON.parse(match[0]);
    return { ...empty, ...parsed } as CompanyFacts;
  } catch (e) {
    console.warn("[extractCompanyFacts] failed:", e);
    return empty;
  }
}

// ---- Web本文 → 用途別中間ブリーフ構築 ----
function detectAudienceAndPurpose(userPrompt: string, title: string): { audience: string; purpose: string } {
  const text = `${userPrompt} ${title}`;
  const audience =
    /初回訪問/.test(text) ? "初回訪問先の担当者" :
    /採用/.test(text) ? "求職者・採用候補者" :
    /社内|内部/.test(text) ? "社内関係者" :
    /投資家|IR/.test(text) ? "投資家・アナリスト" :
    /営業/.test(text) ? "見込み顧客・営業先" :
    "ビジネス関係者";
  const purpose =
    /初回訪問/.test(text) ? "初回訪問用会社紹介" :
    /採用/.test(text) ? "採用向け会社紹介" :
    /IR|投資家/.test(text) ? "IR・投資家向け説明" :
    /営業資料/.test(text) ? "営業資料" :
    /提案/.test(text) ? "提案書" :
    "会社紹介";
  return { audience, purpose };
}

async function buildCompanyBrief(
  companyName: string,
  userPrompt: string,
  title: string,
  evidence: BraveWebEvidence
): Promise<CompanyBrief> {
  const { audience, purpose } = detectAudienceAndPurpose(userPrompt, title);

  const emptyBrief: CompanyBrief = {
    companyName,
    audience,
    purpose,
    companyOverview: "",
    businessAreas: [],
    serviceFlow: [],
    strengths: [],
    metrics: [],
    proofPoints: [],
    recommendedSlideOutline: [],
  };

  const rawCombined = [evidence.snippets, evidence.pages].filter(Boolean).join("\n\n");
  if (!rawCombined) return emptyBrief;
  const webText = cleanWebText(rawCombined).slice(0, 7000);

  try {
    const openai = OpenAIInstance();
    const completion = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME ?? "",
      max_completion_tokens: 4000,
      response_format: { type: "json_object" } as const,
      messages: [
        {
          role: "system",
          content: `You are a business intelligence analyst. Read web content about a company and produce a structured CompanyBrief JSON for a presentation.

CRITICAL: "当社" (our company) = the PRESENTER's company, NOT "${companyName}". Always refer to "${companyName}" by its actual name, never "当社".

Extract ONLY facts explicitly stated in the web text. Do NOT invent.

Output this exact JSON (all text in Japanese):
{
  "companyName": "official name",
  "audience": "${audience}",
  "purpose": "${purpose}",
  "companyOverview": "2-4 sentence overview in Japanese",
  "businessAreas": ["事業領域1", "事業領域2", "事業領域3"],
  "serviceFlow": [{"title": "ステップ名", "body": "説明"}],
  "strengths": ["強み1", "強み2", "強み3"],
  "metrics": [{"label": "創業", "value": "1952年", "note": "詳細"}, {"label": "本社", "value": "東京都", "note": "住所"}, {"label": "従業員", "value": "500名", "note": "時点"}],
  "proofPoints": ["実績・証拠1", "実績・証拠2"],
  "recommendedSlideOutline": [
    {"slideTitle": "スライドタイトル", "layoutType": "company-overview|stat_callouts|card_grid|icon_rows|process-cards|multi-column|closing", "keyConcept": "このスライドで伝えること"}
  ]
}

Rules:
- businessAreas: 3-5 items
- serviceFlow: 2-4 steps if a process is described, empty array otherwise
- strengths: 3-5 items
- metrics: include founding year, location, headcount, stock listing if found. value MAX 15 chars.
- proofPoints: concrete evidence (client count, certifications, awards, rankings)
- recommendedSlideOutline: 6-8 slides with VARIED layoutTypes (no consecutive repeats). Always end with "closing".
- Output JSON only.`,
        },
        {
          role: "user",
          content: `会社名: ${companyName}\n閲覧対象者: ${audience}\n資料の目的: ${purpose}\n\nWebから取得した情報:\n${webText}`,
        },
      ],
    });

    const raw = completion.choices[0]?.message?.content ?? "";
    console.log("[buildCompanyBrief] raw:", raw.slice(0, 500));
    const stripped = raw.replace(/```(?:json)?/gi, "").replace(/```/g, "");
    const match = stripped.match(/\{[\s\S]*\}/);
    if (!match) {
      console.warn("[buildCompanyBrief] no JSON object found");
      return emptyBrief;
    }
    const parsed = JSON.parse(match[0]);
    const brief: CompanyBrief = {
      companyName: parsed.companyName || companyName,
      audience: parsed.audience || audience,
      purpose: parsed.purpose || purpose,
      companyOverview: parsed.companyOverview || "",
      businessAreas: Array.isArray(parsed.businessAreas) ? parsed.businessAreas : [],
      serviceFlow: Array.isArray(parsed.serviceFlow) ? parsed.serviceFlow : [],
      strengths: Array.isArray(parsed.strengths) ? parsed.strengths : [],
      metrics: Array.isArray(parsed.metrics) ? parsed.metrics : [],
      proofPoints: Array.isArray(parsed.proofPoints) ? parsed.proofPoints : [],
      recommendedSlideOutline: Array.isArray(parsed.recommendedSlideOutline) ? parsed.recommendedSlideOutline : [],
    };
    console.log(`[buildCompanyBrief] done: areas=${brief.businessAreas.length} strengths=${brief.strengths.length} metrics=${brief.metrics.length} outline=${brief.recommendedSlideOutline.length}`);
    return brief;
  } catch (e) {
    console.warn("[buildCompanyBrief] failed:", e);
    return emptyBrief;
  }
}

// ---- LLMスライド設計 ----
async function planCompanyProfileSlides(
  title: string,
  brief: CompanyBrief,
  userPrompt: string,
  designInstruction?: string
): Promise<RawPptSlide[]> {
  try {
    const openai = OpenAIInstance();

    const outlineHint = brief.recommendedSlideOutline.length > 0
      ? `\n\n## Recommended Slide Outline (from brief — follow this structure)\n` +
        brief.recommendedSlideOutline.map((o, i) =>
          `${i + 1}. "${o.slideTitle}" → layoutType="${o.layoutType}" — ${o.keyConcept}`
        ).join("\n")
      : "";

    const systemPrompt = `You are an expert PowerPoint presentation designer. Design 7-8 company profile slides in Japanese for "${brief.companyName}". You are the DECISION MAKER for visual design — layout choice, information hierarchy, and text treatment are YOUR responsibility.

## CRITICAL: "当社" の定義
"当社" はこのプレゼンを作成している依頼者側の会社を指します。紹介対象は「${brief.companyName}」です。スライド内で「当社」という言葉は使わず、必ず「${brief.companyName}」または「同社」と表記してください。

## Data Source Rule
Use ONLY information from the CompanyBrief provided. Do NOT invent facts. If a field is empty, omit that content.

## Available layoutTypes (vary across slides — no consecutive repeats)

- "bullets": Bullet list. Use ONLY when no better layout fits. Fields: title, bullets (max 4 items)
- "stat_callouts": 3 large KPI numbers. Use when you have 3+ numeric facts. Fields: title, statCallouts ([{value,unit,label}×3]), bullets (2-3 insights)
- "card_grid": Icon+heading+body card grid (3-6 cards). Use for businessAreas, strengths. Fields: title, cards ([{iconKey,heading,body}×3-6])
- "icon_rows": Icon rows (3-4 rows). Use for proofPoints, capabilities. Fields: title, cards ([{iconKey,heading,body,statusLabel?}×3-4])
- "company-overview": Overview with lead text + metrics. Use companyOverview as leadText. Fields: title, leadText (2-4 sentences), metrics (max 4), callout?, bullets[]
- "metric-cards": KPI emphasis. Fields: title, metrics (max 4), bullets[]
- "process-cards": Step flow. Use serviceFlow as steps. Fields: title, subtitle, steps ([{title,body,iconKey}×2-4]), benefits (2-4), bullets[]
- "timeline": Horizontal steps. Fields: title, subtitle?, steps (3-5), benefits?, bullets[]
- "multi-column": Side-by-side. Fields: title, columns (2-3: {header, bullets[]}), bullets[]
- "closing": Call to action. Fields: title, bullets (3-4 next-step items)

## Metric Card Rules (CRITICAL)
- value: MAX 15 chars (city only, year only, short number)
- note: full detail
- iconKey: calendar/location/stock/network/people/chart/building/gear/verified/star
- colorRole: alternate "primary"/"accent"/"neutral" across cards

## Mandatory Content Rules — EMPTY SLIDES ARE FORBIDDEN
Every slide MUST have at least one non-empty field from: bullets / cards / metrics / steps / statCallouts / leadText. A slide with only a title and empty arrays is INVALID.

- card_grid / icon_rows → cards[] MUST have 3+ items. Each card MUST have iconKey + heading + body.
- process-cards → steps[] MUST have 2+ items. Each step MUST have title + body.
- stat_callouts → statCallouts[] MUST have 3 items. Each MUST have value + unit + label.
- company-overview → leadText MUST be 2-4 sentences.
- closing → bullets[] MUST have 3-4 concrete next steps.

## Design Rules
1. Cover slide is auto-generated — do NOT include a "表紙" slide
2. VARY layoutType — target: company-overview + stat_callouts + card_grid + icon_rows + closing
3. Numbers/KPIs → stat_callouts (not plain bullets)
4. 3+ parallel items → card_grid (not bullets)
5. Process/flow → process-cards or icon_rows (not bullets)
6. Total: 7-8 slides${outlineHint}

Return ONLY this JSON:
{"slides":[{"title":"...","bullets":[],"layoutType":"company-overview","leadText":"...","metrics":[{"label":"創業","value":"1952年","note":"1952年4月","iconKey":"calendar","colorRole":"primary"}]},{"title":"...","bullets":[],"layoutType":"stat_callouts","statCallouts":[{"value":"457","unit":"名","label":"従業員数"},{"value":"1952","unit":"年","label":"創業"},{"value":"94","unit":"%","label":"顧客満足度"}]},{"title":"...","bullets":[],"layoutType":"card_grid","cards":[{"iconKey":"gear","heading":"廃棄物処理","body":"産業廃棄物の収集・運搬・処理を一括対応"},...]},{"title":"まとめ・次のステップ","bullets":["ご不明点はお気軽にご相談ください","導入事例・実績資料をご用意しています","個別提案・現地訪問も対応可能です"],"layoutType":"closing"}]}`;

    const completion = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME ?? "",
      max_completion_tokens: 8000,
      response_format: { type: "json_object" } as const,
      messages: [
        { role: "system", content: systemPrompt },
        {
          role: "user",
          content: `タイトル: ${title}
ユーザー要求: ${userPrompt.slice(0, 400)}
デザイン指示: ${designInstruction ?? "プロフェッショナル・信頼感"}
閲覧対象者: ${brief.audience}
資料の目的: ${brief.purpose}

会社ブリーフ（一次資料 — これだけを根拠にしてください）:
${JSON.stringify(brief, null, 2)}`,
        },
      ],
    });

    const choice = completion.choices[0];
    console.log("[planCompanyProfileSlides] finish_reason:", choice?.finish_reason, "usage:", JSON.stringify(completion.usage));
    const raw = choice?.message?.content ?? "";
    console.log("[planCompanyProfileSlides] raw:", raw.slice(0, 1000));

    if (!raw) {
      console.warn("[planCompanyProfileSlides] empty response");
      return [];
    }

    const stripped = raw.replace(/```(?:json)?/gi, "").replace(/```/g, "").trim();

    // JSON.parse 全体 → .slides を読む（最も安全）
    let parsed: RawPptSlide[] | null = null;
    try {
      const fullObj = JSON.parse(stripped);
      if (Array.isArray(fullObj)) {
        parsed = fullObj;
      } else if (Array.isArray(fullObj?.slides)) {
        parsed = fullObj.slides;
      }
    } catch {
      // フォールバック: 配列部分だけ抽出
      const arrMatch = stripped.match(/\[[\s\S]*\]/);
      if (!arrMatch) {
        console.warn("[planCompanyProfileSlides] no JSON array in response");
        return [];
      }
      parsed = JSON.parse(arrMatch[0]);
    }
    if (!Array.isArray(parsed)) {
      console.warn("[planCompanyProfileSlides] parsed is not an array");
      return [];
    }
    return parsed
      .filter((s) => s.title)
      .map((s) => ({
        ...s,
        bullets: Array.isArray(s.bullets) ? s.bullets : [],
        columns: Array.isArray(s.columns) ? s.columns : undefined,
        tableRows: Array.isArray(s.tableRows) ? s.tableRows : undefined,
        metrics: Array.isArray(s.metrics) ? s.metrics : undefined,
        steps: Array.isArray(s.steps) ? s.steps : undefined,
        benefits: Array.isArray(s.benefits) ? s.benefits : undefined,
      }));
  } catch (e) {
    console.error("[planCompanyProfileSlides] error:", e);
    return [];
  }
}

function buildPptxSearchQuery(title: string, slides: RawPptSlide[] = []): string | null {
  const sourceText = [
    title,
    ...slides.flatMap((s) => [
      s.title,
      ...(s.bullets ?? []),
      ...(s.columns ?? []).flatMap((col) => [col.header, ...(col.bullets ?? [])]),
      ...(s.tableRows ?? []).flat(),
    ]),
  ].join(" ");

  if (!/紹介|会社|提案|営業資料|PR|プロフィール|Profile/.test(sourceText)) return null;

  const quoted = sourceText.match(/[「『"']([^」』"']{2,30})[」』"']/)?.[1];
  const companyLike =
    quoted ||
    sourceText.match(/(?:株式会社|有限会社|合同会社|（株）|\(株\))\s*([^\s、。・:：]{2,30})/)?.[1] ||
    sourceText.match(/([ァ-ヶー一-龠A-Za-z0-9]{2,30})(?:の)?(?:会社紹介|紹介資料|営業資料|提案書|プロフィール|Profile)/)?.[1];

  const target = (companyLike ?? title)
    .replace(/^(?:株式会社|有限会社|合同会社|（株）|\(株\))/, "")
    .replace(/（[^）]*）|\([^)]*\)/g, "")
    .replace(/^(?:会社紹介|紹介資料|営業資料|提案書|プロフィール|Profile)$/, "")
    .trim()
    .split(/[\s　]/)[0];

  if (!target || target.length < 2) return null;
  return `${target} 会社概要 事業内容 実績`;
}

type RawPptSlide = {
  title: string;
  bullets: string[];
  layoutType?: string;
  columns?: Array<{ header: string; bullets: string[] }>;
  tableRows?: string[][];
  leadText?: string;
  metrics?: Array<{
    label: string;
    value: string;        // 表示用短縮値（LLMが設定: 最大15文字）
    note?: string;        // 補足詳細（LLMが設定）
    iconKey?: string;
    displayValue?: string;
    colorRole?: "primary" | "accent" | "neutral";
  }>;
  callout?: { title: string; body: string };
  subtitle?: string;
  steps?: Array<{ title: string; body: string; iconKey?: string }>;
  benefits?: string[];
  cards?: Array<{ iconKey?: string; heading: string; body: string; statusLabel?: string }>;
  statCallouts?: Array<{ value: string; unit: string; label: string }>;
  // LLMデザイン判断フィールド
  visualIntent?: string;
  density?: "low" | "medium" | "high";
  textTreatment?: "short" | "normal" | "explanatory";
};

// Brave結果からキー事実を正規表現で抽出（LLM呼び出しなし・切れる心配なし）
function extractFactsFromWeb(webContext: string): Record<string, string> {
  const facts: Record<string, string> = {};
  const text = webContext.replace(/【[^】]*】/g, " "); // タイトル部分を除去して本文優先

  const foundingM = text.match(/(?:19|20)(\d{2})年(?:の)?(?:創業|設立)/);
  if (foundingM) facts["創業"] = foundingM[0].replace(/(?:創業|設立)/, "").trim();

  const locM = text.match(/(静岡県浜松市|浜松市(?:[^、。\s]{0,6})?|静岡県(?:[^、。\s]{0,10})?)/);
  if (locM) facts["本社"] = locM[1].trim();

  if (/東証プライム/.test(text)) facts["上場"] = "東証プライム";
  else if (/東証スタンダード/.test(text)) facts["上場"] = "東証スタンダード";
  else if (/東証グロース/.test(text)) facts["上場"] = "東証グロース";

  const clientM = text.match(/約?([\d,，万]+)\s*社(?:以上)?(?:の取引|との取引|との契約)?/);
  if (clientM) facts["取引先"] = `約${clientM[1].replace(/[，]/g, ",")}社`;

  const stockM = text.match(/\((\d{4})\)/);
  if (stockM) facts["証券コード"] = stockM[1];

  // 従業員数
  const empM = text.match(/従業員(?:数)?[：:は]?\s*約?([\d,，]+)\s*名/);
  if (empM) facts["従業員"] = `約${empM[1].replace(/[，]/g, ",")}名`;

  // 売上高・営業収益
  const revM = text.match(/(?:売上高|営業収益)[：:は]?\s*約?([\d,，.]+)\s*(?:億円|百億円)/);
  if (revM) facts["売上"] = `${revM[1]}億円`;

  // 施設数・拠点数
  const facilityM = text.match(/(?:施設数?|処理施設)[：:は]?\s*約?([\d]+)\s*(?:ヵ所|箇所|か所|施設)/);
  if (facilityM) facts["施設"] = `${facilityM[1]}施設`;
  const baseM = text.match(/(?:拠点数?)[：:は]?\s*約?([\d]+)\s*(?:ヵ所|箇所|か所|拠点)/);
  if (baseM) facts["拠点"] = `${baseM[1]}拠点`;

  // 処理能力（廃棄物特有）
  const capM = text.match(/(?:処理能力|年間処理量)[：:は]?\s*約?([\d,，万]+)\s*(?:トン|t)/);
  if (capM) facts["処理能力"] = `約${capM[1]}t/年`;

  console.log("[enrich-slides] extracted facts:", facts);
  return facts;
}

function applyFact(text: string, facts: Record<string, string>): string {
  let t = text;
  // プレースホルダー置換（[〇〇] 形式）
  if (facts["創業"])    t = t.replace(/\[(?:創業年?|設立年?|創業年度|設立年度)\]/g, facts["創業"]);
  if (facts["本社"])    t = t.replace(/\[(?:本社|所在地|住所|拠点|市区町村)\]/g, facts["本社"]);
  if (facts["上場"])    t = t.replace(/\[(?:上場|市場区分|証券取引所|上場市場)\]/g, facts["上場"]);
  if (facts["取引先"])  t = t.replace(/\[(?:取引先数?|顧客数?|取引社数?|取引先)\]/g, facts["取引先"]);
  if (facts["証券コード"]) t = t.replace(/\[(?:証券コード|コード|銘柄コード)\]/g, facts["証券コード"]);
  // 「YYYY年」形式の補完（[YYYY]）
  if (facts["創業"])    t = t.replace(/\[YYYY\]/g, facts["創業"]);
  return t;
}

function enrichSlidesWithWebData(slides: RawPptSlide[], webContext: string): Promise<RawPptSlide[]> {
  if (!webContext) return Promise.resolve(slides);

  const facts = extractFactsFromWeb(webContext);
  if (Object.keys(facts).length === 0) return Promise.resolve(slides);

  let applied = 0;
  const result = slides.map((s) => {
    const updated = {
      ...s,
      bullets: [...(s.bullets ?? [])],
      metrics: s.metrics?.map((m) => ({ ...m })),
      callout: s.callout ? { ...s.callout } : undefined,
      steps: s.steps?.map((st) => ({ ...st })),
    };

    // leadText
    if (updated.leadText) {
      const n = applyFact(updated.leadText, facts);
      if (n !== updated.leadText) { updated.leadText = n; applied++; }
    }
    // metrics
    updated.metrics?.forEach((m) => {
      const nv = applyFact(m.value, facts);
      if (nv !== m.value) { m.value = nv; applied++; }
      if (m.note) { const nn = applyFact(m.note, facts); if (nn !== m.note) { m.note = nn; applied++; } }
    });
    // callout.body
    if (updated.callout?.body) {
      const n = applyFact(updated.callout.body, facts);
      if (n !== updated.callout.body) { updated.callout.body = n; applied++; }
    }
    // bullets（先頭3件のみ）
    updated.bullets.slice(0, 3).forEach((b, i) => {
      const n = applyFact(b, facts);
      if (n !== b) { updated.bullets[i] = n; applied++; }
    });

    return updated;
  });

  console.log(`[enrich-slides] regex applied ${applied} enrichments from ${Object.keys(facts).length} facts`);
  return Promise.resolve(result);
}

// ---------------- SharePoint コンテンツを使ったPPTスライド補充 ----------------

/**
 * ユーザーメッセージから "SharePointにある〇〇" パターンを検出し、
 * 検索クエリ文字列を返す。見つからなければ null。
 */
function extractSharePointDocQuery(userMessage: string): string | null {
  // "SharePointにある[文書名]" / "SharePointの[文書名]" パターン
  const m = userMessage.match(/Share\s*Point[にのの上]ある([^\s　、。!！?？\n]{3,60})/i)
         ?? userMessage.match(/Share\s*Point[にのの上]([^\s　、。!！?？\n]{3,60}(?:報告|資料|ドキュメント|書類|一覧|まとめ)[^\s　、。!！?？\n]*)/i);
  if (!m?.[1]) return null;

  // 末尾の助詞・動詞句を除去 ("を参考に" / "を参照して" 等)
  const doc = m[1]
    .replace(/[をはがにの]*(?:参考|参照|もと|確認|把握|読ん|見て)[^\s]*/g, "")
    .replace(/[をはがにの]+$/, "")
    .trim();
  return doc.length >= 2 ? doc : null;
}

/** SharePoint インデックスを検索してスライド補充用テキストを返す */
async function searchSpForPptxContent(docQuery: string): Promise<string> {
  const apiKey    = process.env.AZURE_SEARCH_API_KEY?.trim()    || "";
  const searchName = process.env.AZURE_SEARCH_NAME?.trim()      || "";
  const indexName  = process.env.AZURE_SEARCH_INDEX_NAME?.trim() || "";
  if (!apiKey || !searchName || !indexName) return "";

  const session  = await userSession();
  const deptLower = session?.slDept?.toLowerCase().trim() || null;

  console.log(`[create_pptx] SP search: "${docQuery}" dept=${deptLower}`);

  const result = await ExtensionSimilaritySearch({
    searchText: docQuery,
    vectors: ["embedding"],
    apiKey,
    searchName,
    indexName,
    filter: undefined,   // ACL フィルタに委ねる
    deptLower,
    userHash: undefined, // buildSearchAclFilter が userHashedId() でフォールバック
    top: 10,
  });

  if (result.status !== "OK" || result.response.length === 0) {
    console.log("[create_pptx] SP search: 結果なし");
    return "";
  }

  const content = result.response
    .map((r, i) => `[${i}] ${r.document.metadata ?? ""}\n${r.document.pageContent}`)
    .join("\n---\n");
  console.log(`[create_pptx] SP search: ${result.response.length}件取得`);
  return content;
}

/**
 * LLM を使って SP ドキュメント内容でスライドの bullet を書き直す。
 * 構造（title・layoutType）は維持し、内容のみ SP 情報で充填する。
 */
async function enrichSlidesWithDocContent(
  slides: RawPptSlide[],
  docContent: string,
  title: string,
  userPrompt: string
): Promise<RawPptSlide[]> {
  if (!docContent || !slides.length) return slides;

  const openai = OpenAIInstance();
  const slideSkeleton = JSON.stringify(
    slides.map((s) => ({
      title: s.title,
      bullets: s.bullets,
      layoutType: s.layoutType,
      metrics: s.metrics,
      steps: s.steps,
    }))
  );

  try {
    const completion = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME ?? "",
      max_completion_tokens: 6000,
      response_format: { type: "json_object" } as const,
      messages: [
        {
          role: "system",
          content:
            "You are a presentation strategist specializing in executive communications. " +
            "Given a target slide structure and reference document content, rewrite each slide with relevant facts, numbers, and details. " +
            "Use ONLY information from the document — never invent facts not present in the document. " +
            "Preserve the same slide count, titles, and layoutTypes. " +
            "\n\n" +
            "RESTRUCTURING MANDATE — CRITICAL:\n" +
            "Do NOT map document content to slides in chronological or document-page order.\n" +
            "Treat ALL document content as a flat pool of facts, then ASSIGN each fact to the slide whose TOPIC best matches — regardless of which quarter, section, or page it appeared in.\n" +
            "If multiple periods (Q1/Q2/Q3/Q4) reported the same metric, synthesize them: show the latest value or the trend (e.g., 'Q1時点40名→Q4現在55名').\n" +
            "If a slide topic is 'KPIと利用実績', pull ALL KPI data from ALL parts of the document.\n" +
            "If a slide topic is 'コスト・投資対効果', pull ALL cost/budget information, not just one quarter's mention.\n" +
            "Related tools like 議事郎/議事録アプリ should be presented as USE CASES of the main product, not as separate products.\n" +
            "\n" +
            "For executive audiences: each slide must answer a business question ('なぜ重要か' / '何ができるか' / '投資上の意味は何か'), not just describe a time period.\n" +
            "\n" +
            "For bullets: concrete and specific (avoid vague placeholders). " +
            "For metrics: use numeric values from the document if available. " +
            "IMPORTANT: All text in bullets, leadText, callout, steps body MUST be in polite Japanese (です/ます調). " +
            "Do NOT use noun-ending style (体言止め) or abrupt verb endings (〜する、〜実施). " +
            "CRITICAL — complete sentences only: metric.note / card.body / bullets / steps.body must each end at a natural boundary " +
            "(句点「。」, closing parenthesis「）」, closing quote「」」, or a period). " +
            "NEVER produce mid-sentence cuts — always include the closing quote and full thought. " +
            "When shortening, shorten to the nearest preceding sentence boundary, not by character count. " +
            "Output JSON: {\"slides\": [/* same structure as input */]}",
        },
        {
          role: "user",
          content:
            `プレゼンタイトル: ${title}\nユーザー要求: ${userPrompt.slice(0, 300)}\n\n` +
            `## スライド骨格 (JSON):\n${slideSkeleton}\n\n` +
            `## 参照ドキュメント (SharePoint):\n${docContent.slice(0, 7000)}`,
        },
      ],
    });

    const raw = completion.choices[0]?.message?.content ?? "";
    const stripped = raw.replace(/```(?:json)?/gi, "").replace(/```/g, "");
    const match = stripped.match(/\{[\s\S]*\}/);
    if (!match) return slides;

    const parsed = JSON.parse(match[0]);
    const newSlides = parsed.slides as RawPptSlide[];
    if (!Array.isArray(newSlides) || newSlides.length === 0) return slides;
    console.log(`[create_pptx] enrichSlidesWithDocContent: ${newSlides.length}枚をSP内容で補充`);
    return newSlides;
  } catch (e) {
    console.warn("[create_pptx] enrichSlidesWithDocContent failed:", e);
    return slides;
  }
}

// ---------------- 提案書スライド展開 ----------------
type ProposalSlide = {
  title: string;
  bullets: string[];
  layoutType?: string;
  columns?: Array<{ header: string; bullets: string[] }>;
  tableRows?: string[][];
  leadText?: string;
  metrics?: Array<{ label: string; value: string; note?: string; iconKey?: string }>;
  callout?: { title: string; body: string };
  subtitle?: string;
  steps?: Array<{ title: string; body: string; iconKey?: string }>;
  benefits?: string[];
  // 新レイアウト型用フィールド
  statCallouts?: Array<{ value: string; unit: string; label: string }>;
  cards?: Array<{ iconKey?: string; heading: string; body: string; statusLabel?: string }>;
};

async function expandToProposalSlides(
  title: string,
  inputSlides: ProposalSlide[],
  designHint?: string,
  deptLower?: string,
  webContext?: string
): Promise<ProposalSlide[]> {
  try {
    const openai = OpenAIInstance();
    const inputSummary = inputSlides.length
      ? inputSlides.map((s) => `- ${s.title}: ${(s.bullets ?? []).slice(0, 2).join(" / ")}`).join("\n")
      : "（初期スライドなし）";

    // SharePoint文書を検索してコンテキストとして取得
    const spContext = deptLower
      ? await fetchSpContextForProposal(title, inputSlides, deptLower)
      : "";

    const spSection = spContext
      ? `\n\n【社内SharePoint文書（必ず内容を反映させること。LLMの事前学習知識より優先すること）】\n${spContext}`
      : "";

    const webSection = webContext
      ? `\n\n【Web検索結果（会社・業界の公開情報 - プレースホルダー不可、実データを使うこと）】\n${webContext}`
      : "";

    const systemPrompt = `あなたは営業提案書のスライド構成の専門家です。与えられたタイトル・初期スライド・社内文書・Web情報を元に、12〜16枚の提案書スライドを生成してください。

【最重要1】社内SharePoint文書が提供されている場合は、その内容（数値・事例・実績・規程・方針）を必ずスライドの bullets に盛り込むこと。
【最重要2】Web検索結果が提供されている場合は、会社の実際のデータ（創業年・所在地・従業員数・事業内容・実績など）を bullets に直接使うこと。[〇〇]等のプレースホルダーは絶対に使わないこと。

【構成の流れ（必須）】
1. 表紙（タイトルスライド）
2. 課題・背景（顧客が抱える問題）
3. 現状の問題点（具体的な課題の深掘り）
4. 提案概要（一言で伝える解決策）
5〜7. 提案詳細（サービス内容・特徴・強みを2〜3スライドで）
8. 根拠・実績（数値・事例・実績。SP文書の数値を使うこと）
9. 他社比較（layoutType="multi-column"、3列比較を推奨）
10. 導入効果（layoutType="table"、効果を数値で）
11. コスト感・導入ロードマップ
12. まとめ・次のステップ

【使用できる layoutType と必須フィールド】
- "bullets": 箇条書き3〜4項目。フィールド: title, bullets (max 4)
- "stat_callouts": 数値KPI3つを大きく表示。フィールド: title, statCallouts ([{value,unit,label}×3]), bullets (インサイト2〜3件)
- "card_grid": アイコン付きカード3〜6枚グリッド。フィールド: title, cards ([{iconKey,heading,body}×3〜6])
- "icon_rows": アイコン行3〜4本（ステータスピル付き可）。フィールド: title, cards ([{iconKey,heading,body,statusLabel?}×3〜4])
- "process-cards": ステップフロー。フィールド: title, subtitle, steps ([{title,body,iconKey}×2〜4]), benefits (2〜4)
- "multi-column": 比較2〜3列。フィールド: title, columns ([{header,bullets[]}×2〜3])
- "table": 構造化表。フィールド: title, tableRows (1行目=ヘッダー)
- "closing": CTAまとめ。フィールド: title, bullets (3〜4件)

【各スライドのルール】
- bullets は3〜4項目のみ（詰め込まない）
- 各 bullet は具体的な1〜2文。キーワードのみ禁止
- 数値・実績・KPIが出てきたら stat_callouts に振り分けること（表に詰めない）
- 機能・強み・特徴を3〜6つ並べるなら card_grid を使うこと（箇条書きにしない）
- 手順・プロセス・対応状況なら icon_rows または process-cards を使うこと
- 「表紙」タイトルのスライドは生成しないこと（自動生成される）

必ず以下のJSON形式で返すこと（配列のみ、説明文なし）:
[{"title":"...","bullets":["..."],"layoutType":"bullets"}]`;

    const userPrompt = `タイトル: ${title}
デザインヒント: ${designHint ?? "ビジネス向け"}
初期スライド:
${inputSummary}${spSection}${webSection}`;

    const completion = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME ?? "",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt },
      ],
      max_completion_tokens: 8000,
    });

    const propChoice = completion.choices[0];
    console.log("[proposalMode] finish_reason:", propChoice?.finish_reason, "usage:", JSON.stringify(completion.usage));
    const raw = propChoice?.message?.content ?? "";
    const jsonMatch = raw.match(/\[[\s\S]*\]/);
    if (!jsonMatch) {
      console.warn("[proposalMode] Failed to extract JSON from response, raw(200):", raw.slice(0, 200));
      return inputSlides;
    }

    const parsed: ProposalSlide[] = JSON.parse(jsonMatch[0]);
    return parsed
      .filter((s) => s.title)
      .map((s) => ({
        ...s,
        bullets: Array.isArray(s.bullets) ? s.bullets : [],
        columns: Array.isArray(s.columns) ? s.columns : undefined,
        tableRows: Array.isArray(s.tableRows) ? s.tableRows : undefined,
      }));
  } catch (e) {
    console.error("[proposalMode] expandToProposalSlides error:", e);
    return inputSlides;
  }
}

// ---------------- LLMレビュー&修正 ----------------

async function reviewAndRefineSlides(
  title: string,
  slides: RawPptSlide[],
  designInstruction?: string
): Promise<RawPptSlide[]> {
  try {
    const openai = OpenAIInstance();
    const prompt = `あなたはB2B営業資料に強いプレゼンテーションデザイナーです。
以下のスライドJSONを見て、不自然・ダサい箇所を修正してください。

チェック項目:
1. タイトル・本文がプロンプトの転記になっていないか（閲覧者視点の表現に書き直す）
2. colorRole が意味ベースか（数値・実績・差別化 → accent、基本情報 → primary、補足 → neutral）
3. bullets が自然な箇条書きか（1〜2文。ただし意味が完結する文にすること）
4. layoutType が内容に合っているか
5. metrics/steps/bullets の情報量が多すぎないか（各最大4項目）
6. 【文体統一】bullets・leadText・callout・steps の本文はすべて「です/ます調」に統一すること。体言止め・言い切り（〜する、〜推進、〜実施）は「〜しています」「〜できます」等に書き直す。
7. 【未完文禁止】metric.note / card.body / bullets / steps.body はすべて句点「。」・閉じ括弧「）」・閉じ引用符「」」で終わること。「ユーザーアンケートで『同僚に薦め」のような途中切れは絶対禁止。短縮する場合も直前の文末まで含めること。
8. 【経営向けストーリー確認】タイトルやbulletsに「Q1」「Q2」「Q3」「Q4」「第1四半期」「第2四半期」など時系列ラベルが複数のスライドに散在していた場合、それは「定期レポートを時系列に並べた構成」になっています。経営層向け資料では、以下のアーク構造が正しい姿です：目的・位置づけ → 主な機能 → 利用状況・KPI → 拡張・連携状況 → セキュリティ・ガバナンス → コスト・投資対効果 → 課題・リスク → ロードマップ → 経営判断が必要な論点。時系列構造を検知した場合、各スライドのtitleをカテゴリ軸に書き直し、bulletsを該当カテゴリに適合した内容に整理してください。「議事郎」などの連携ツールは独立スライドを作らず、「連携・拡張状況」スライドのbulletsに統合すること。

重要: metrics・steps・colorRole・iconKey・layoutType・leadText・callout フィールドは削除しないこと。
変更不要なスライドはそのまま返すこと。

元タイトル: ${title}
デザイン指示: ${designInstruction ?? "なし"}
スライドJSON:
${JSON.stringify(slides)}

{"slides":[...]} の形式でJSONのみ返してください。`;

    const res = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_PPT_DEPLOYMENT_NAME ?? process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
      messages: [{ role: "user", content: prompt }],
      max_completion_tokens: 8000,
      response_format: { type: "json_object" },
    });

    const raw = res.choices[0]?.message?.content ?? "";
    const parsed = JSON.parse(raw);
    const refined: RawPptSlide[] = parsed?.slides;

    // 構造検証: スライド数・title・layoutType・bullets が壊れていたら元に戻す
    if (!Array.isArray(refined) || refined.length === 0) {
      console.warn("[reviewSlides] empty result, using original");
      return slides;
    }
    if (refined.length < slides.length * 0.7) {
      console.warn(`[reviewSlides] too few slides (${refined.length} < ${slides.length}), using original`);
      return slides;
    }
    const hasStructure = refined.every(
      (s) => typeof s.title === "string" && Array.isArray(s.bullets)
    );
    if (!hasStructure) {
      console.warn("[reviewSlides] structure broken, using original");
      return slides;
    }

    console.log(`[reviewSlides] refined ${slides.length} → ${refined.length} slides`);
    return refined.map((s, i) => ({
      // 元スライドのフィールドをベースに、レビュー結果で上書き（重要フィールドの消失を防ぐ）
      ...slides[i],
      ...s,
      // 空配列はレビュー結果を採用せず元スライドを維持
      bullets:      (Array.isArray(s.bullets)      && s.bullets.length      > 0) ? s.bullets      : (slides[i]?.bullets      ?? []),
      metrics:      (Array.isArray(s.metrics)      && s.metrics.length      > 0) ? s.metrics      : slides[i]?.metrics,
      steps:        (Array.isArray(s.steps)        && s.steps.length        > 0) ? s.steps        : slides[i]?.steps,
      cards:        (Array.isArray(s.cards)        && s.cards.length        > 0) ? s.cards        : slides[i]?.cards,
      statCallouts: (Array.isArray(s.statCallouts) && s.statCallouts.length > 0) ? s.statCallouts : slides[i]?.statCallouts,
      benefits:     (Array.isArray(s.benefits)     && s.benefits.length     > 0) ? s.benefits     : slides[i]?.benefits,
    }));
  } catch (e) {
    console.warn("[reviewSlides] failed, using original slides:", e);
    return slides;
  }
}

/** 各PDFのスライドをタイトル＋bullets のテキストブロックに変換する（経営向け再構築用の事実プール） */
function buildDocSummaryFromSlides(
  fileName: string,
  slides: Array<{ title: string; bullets?: string[] }>
): string {
  const lines = [`【${fileName}】`];
  for (const slide of slides) {
    lines.push(`■ ${slide.title}`);
    for (const bullet of (slide.bullets ?? [])) {
      lines.push(`  ・${bullet}`);
    }
  }
  return lines.join("\n");
}

/** 複数PDFのスライドを経営向け9カテゴリに再構築する（per-doc中間要約で情報源を確保） */
async function restructureSlidesForExecutive(
  title: string,
  mergedSlides: RawPptSlide[],
  perDocSummaries: string[],
  designInstruction?: string
): Promise<RawPptSlide[]> {
  try {
    const openai = OpenAIInstance();
    const summaryBlock = perDocSummaries.length > 0
      ? `\n\n=== 各ドキュメントの中間要約（事実プール）===\n${perDocSummaries.join("\n\n")}\n========================`
      : "";

    const prompt = `あなたはB2B経営層向けプレゼンテーションの構成エキスパートです。
複数の四半期レポートや会議録をマージしたスライドJSONと、各PDFの中間要約を受け取り、経営層向けの9カテゴリ構成に再整理してください。${summaryBlock}

再整理ルール:
1. 以下の9カテゴリ軸でスライドを構成すること:
   目的・位置づけ → 主な機能 → 利用状況・KPI → 拡張・連携状況 → セキュリティ・ガバナンス → コスト・投資対効果 → 課題・リスク → ロードマップ → 経営判断が必要な論点
2. 各PDFの中間要約を「事実プール」として扱い、四半期ごとの時系列構造は崩す
3. 固有名詞・数値・四半期由来の根拠（例: Q1実績◯件、Q3計画）は削除せずカテゴリのbulletsに組み込む
4. bullets: 各bullet 45〜90文字、1カテゴリあたり3〜5項目（数値・固有名詞は短縮しない）
5. 情報量を増やす方向で整理すること。圧縮・省略禁止
6. metrics・steps・colorRole・iconKey・layoutType・leadText・callout フィールドは削除しないこと
7. 「議事郎」などの連携ツールは独立スライドを作らず「拡張・連携状況」スライドのbulletsに統合すること
8. すべての文末は「です/ます調」にすること

元タイトル: ${title}
デザイン指示: ${designInstruction ?? "なし"}
マージ済みスライドJSON:
${JSON.stringify(mergedSlides)}

{"slides":[...]} の形式でJSONのみ返してください。`;

    const res = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_PPT_DEPLOYMENT_NAME ?? process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME!,
      messages: [{ role: "user", content: prompt }],
      max_completion_tokens: 10000,
      response_format: { type: "json_object" },
    });

    const raw = res.choices[0]?.message?.content ?? "";
    const parsed = JSON.parse(raw);
    const restructured: RawPptSlide[] = parsed?.slides;

    if (!Array.isArray(restructured) || restructured.length === 0) {
      console.warn("[restructureExec] empty result, using original");
      return mergedSlides;
    }
    if (restructured.length < 5) {
      console.warn(`[restructureExec] too few slides (${restructured.length}), using original`);
      return mergedSlides;
    }
    const hasStructure = restructured.every(
      (s) => typeof s.title === "string" && Array.isArray(s.bullets)
    );
    if (!hasStructure) {
      console.warn("[restructureExec] structure broken, using original");
      return mergedSlides;
    }

    console.log(`[restructureExec] restructured ${mergedSlides.length} → ${restructured.length} slides`);
    return restructured;
  } catch (e) {
    console.warn("[restructureExec] failed, using original slides:", e);
    return mergedSlides;
  }
}

// ---------------- 会社紹介モード ----------------

function detectCompanyProfileMode(
  title: string,
  slides: RawPptSlide[],
  designInstruction?: string
): boolean {
  const text = `${title} ${(designInstruction ?? "")}`.toLowerCase();
  // "機能紹介資料" は製品機能紹介であり会社紹介ではないため除外
  const hasProfile = /会社紹介|(?<!機能)紹介資料|company profile|初回訪問|初回営業/.test(text);
  const hasSmallDeck = slides.length <= 10;
  return hasProfile && hasSmallDeck;
}

const TITLE_SUFFIXES =
  /[\s　]*(会社紹介|紹介資料|営業資料|提案書|会社概要|初回訪問|COMPANY\s*PROFILE|Company\s*Profile|プロフィール|Profile)/gi;

function extractCompanyNameFromTitle(title: string): string {
  const cleaned = title
    .replace(/（[^）]*）|\([^)]*\)/g, "")
    .replace(TITLE_SUFFIXES, "")
    .trim();

  const quoted = cleaned.match(/[「『"']([^」』"']{2,20})[」』"']/)?.[1];
  if (quoted) return quoted;

  // 株式会社などのプレフィックスを除去してから先頭語を返す
  const noPrefix = cleaned.replace(/^(株式会社|有限会社|合同会社|（株）|\(株\))\s*/, "");
  return (noPrefix.split(/[\s　]/)[0] ?? cleaned).slice(0, 20);
}

// ---------------- Python レンダラー経由 PowerPoint 生成 ----------------

async function executeCreatePptxPython(
  args: {
    title: string;
    slides: RawPptSlide[];
    palette: string;
    designInstruction?: string;
  },
  chatThread: ChatThreadModel
) {
  const { title, slides, palette, designInstruction } = args;

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    const res = await fetch(`${baseUrl}/api/gen-pptx-profile`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        title,
        slides: (slides ?? []).map((s) => ({
          title:      s.title,
          bullets:    Array.isArray(s.bullets) ? s.bullets : [],
          layoutType: s.layoutType,
          leadText:   s.leadText,
          callout:    s.callout,
          metrics:    s.metrics,
          steps:      s.steps,
          benefits:   s.benefits,
          subtitle:   s.subtitle,
        })),
        palette,
        designInstruction,
        threadId: chatThread.id,
        fileBaseName: generatePptxDisplayName(title).replace(/\.pptx$/i, ""),
      }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error("[create_pptx_py] gen-pptx-profile failed:", res.status, t);
      return { error: `PowerPoint生成に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();
    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    console.log(`[create_pptx_py] palette=${result.palette} → ${result.fileName}`);
    return {
      downloadUrl: result.downloadUrl,
      fileName:    result.fileName,
      displayName: generatePptxDisplayName(title),
      palette:     result.palette,
      message:     "PowerPoint file created successfully.",
    };
  } catch (e: any) {
    console.error("[create_pptx_py] error:", e);
    return { error: "PowerPoint生成中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- PowerPoint 生成 ----------------

function generatePptxDisplayName(title: string): string {
  const clean = title
    .replace(/（[^）]*）|\([^)]*\)/g, "")
    .replace(/[\\/:*?"<>|【】「」『』〔〕]/g, "")
    .replace(/\s+/g, "")
    .trim()
    .slice(0, 30);
  return `${clean || "プレゼンテーション"}.pptx`;
}

async function executeCreatePptx(
  args: {
    title: string;
    slides: RawPptSlide[];
    proposalMode?: boolean;
    fontFace?: string;
    designInstruction?: string;
    palette?: string;
  },
  chatThread: ChatThreadModel,
  userMessage?: string
) {
  const { title, slides, proposalMode, fontFace, designInstruction, palette } = args ?? {};

  if (!title || !slides?.length) {
    return { error: "title and slides are required." };
  }

  // PromptIntent を finalSlides 生成前に解析し、以降のプロンプトへ伝搬する
  const intentSource = [designInstruction ?? "", title, userMessage ?? ""].filter(Boolean).join(" ");
  const promptIntent = parsePromptIntent(intentSource);
  const ld = promptIntent.layoutDirectives;
  console.log(
    `[PromptIntent] purpose=${promptIntent.documentPurpose} audience=${promptIntent.audience} ` +
    `freedom=${promptIntent.designFreedom} twoCol=${!!ld.preferTwoColumn} tables=${!!ld.includeTables} ` +
    `metrics=${!!ld.preferMetrics} process=${!!ld.preferProcess}` +
    (promptIntent.colorDirectives?.primary ? ` colors=${promptIntent.colorDirectives.primary}/${promptIntent.colorDirectives.accent ?? "?"}` : "")
  );

  // layoutDirectives をデザイン指示文に追加してスライド設計 LLM に伝搬
  const layoutHints: string[] = [];
  if (ld.preferTwoColumn) layoutHints.push("2列レイアウト(multi-column)を少なくとも1枚含めること");
  if (ld.includeTables)   layoutHints.push("表形式(table)のスライドを少なくとも1枚含めること");
  if (ld.preferMetrics)   layoutHints.push("数値・KPIを強調するmetric-cardsを使うこと");
  if (ld.preferProcess)   layoutHints.push("手順・フローにはprocess-cardsまたはtimelineを使うこと");
  if (ld.avoidBulletOnly) layoutHints.push("箇条書きのみのスライドが連続しないようレイアウトを変化させること");
  const layoutHintText = layoutHints.length > 0 ? `【レイアウト要件】${layoutHints.join("。")}` : "";

  const searchQuery = buildPptxSearchQuery(title, slides);
  let finalSlides: RawPptSlide[] = slides;

  // ★ SharePoint 参照検出: "SharePointにある〇〇を参考に" パターンがあればSP優先
  const spDocQuery = userMessage ? extractSharePointDocQuery(userMessage) : null;
  if (spDocQuery) {
    const spContent = await searchSpForPptxContent(spDocQuery);
    if (spContent) {
      finalSlides = await enrichSlidesWithDocContent(slides, spContent, title, userMessage ?? "");
    }
  } else if (proposalMode) {
    // 提案書モード: 12〜16枚展開（Brave snippetのみ継続使用）
    let webContext = "";
    if (searchQuery) {
      webContext = await searchBrave(searchQuery);
    }
    const session = await userSession();
    const deptLower = (session?.slDept ?? "others").toLowerCase().trim();
    finalSlides = await expandToProposalSlides(title, slides, designInstruction, deptLower, webContext);
  } else if (!proposalMode && detectCompanyProfileMode(title, slides, designInstruction)) {
    // 会社紹介モード: Web事実収集 → CompanyBrief構築 → LLMスライド設計
    const companyName = extractCompanyNameFromTitle(title);
    const query = companyName
      ? `${companyName} 会社概要 事業内容 実績`
      : (searchQuery || `${title} 会社概要 事業内容`);
    console.log("[create_pptx] company profile mode — collectWebEvidence:", query);
    const evidence = await collectWebEvidence(query);
    const brief = await buildCompanyBrief(companyName, userMessage ?? "", title, evidence);
    console.log(`[create_pptx] brief built: areas=${brief.businessAreas.length} strengths=${brief.strengths.length} metrics=${brief.metrics.length} outline=${brief.recommendedSlideOutline.length}`);
    const planned = await planCompanyProfileSlides(
      title, brief, userMessage ?? "", designInstruction
    );
    if (planned.length > 0) {
      finalSlides = planned;
    } else {
      // フォールバック: スニペットでregex補完
      const snippetContext = evidence.snippets;
      if (snippetContext) finalSlides = await enrichSlidesWithWebData(slides, snippetContext);
    }
  } else if (searchQuery) {
    // 通常モード: Brave snippetでregex補完
    const webContext = await searchBrave(searchQuery);
    if (webContext) finalSlides = await enrichSlidesWithWebData(slides, webContext);
  }

  // LLMレビュー: スライド内容を見直して不自然な箇所を修正（layoutHintText でレイアウト要件を伝搬）
  const reviewInstruction = [designInstruction, layoutHintText].filter(Boolean).join(" / ");
  finalSlides = await reviewAndRefineSlides(title, finalSlides, reviewInstruction);

  const explicitInstruction = designInstruction?.trim() ||
    (proposalMode
      ? "提案書スタイル：課題→解決策→根拠→効果の流れを視覚的に表現。濃紺ベース、見出しは白抜き太字、重要数値は大きく強調。スライドごとにレイアウトを変化させ、比較スライドは表形式、プロセスはフロー図で表現すること。"
      : "プロフェッショナルで信頼感のあるビジネス向けデザイン。見出しは太字で視認性高く、数値・実績は強調表示。スライド間でレイアウトに変化をつけること。");
  const deckPreferences: DeckPreferences = { designInstruction: explicitInstruction };

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    const res = await fetch(`${baseUrl}/api/gen-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        title,
        slides: (finalSlides ?? []).map((s) => ({
          title: s.title,
          bullets: Array.isArray(s.bullets) ? s.bullets : [],
          ...(s.layoutType    ? { layoutType: s.layoutType }       : {}),
          ...(s.columns       ? { columns: s.columns }             : {}),
          ...(s.tableRows     ? { tableRows: s.tableRows }         : {}),
          ...(s.leadText      ? { leadText: s.leadText }           : {}),
          ...(s.metrics       ? { metrics: s.metrics }             : {}),
          ...(s.callout       ? { callout: s.callout }             : {}),
          ...(s.subtitle      ? { subtitle: s.subtitle }           : {}),
          ...(s.steps         ? { steps: s.steps }                 : {}),
          ...(s.benefits      ? { benefits: s.benefits }           : {}),
          ...(s.cards         ? { cards: s.cards }                 : {}),
          ...(s.statCallouts  ? { statCallouts: s.statCallouts }   : {}),
          ...(s.visualIntent  ? { visualIntent: s.visualIntent }   : {}),
          ...(s.density       ? { density: s.density }             : {}),
          ...(s.textTreatment ? { textTreatment: s.textTreatment } : {}),
        })),
        threadId: chatThread.id,
        fontFace,
        designInstruction: explicitInstruction,
        deckPreferences,
        fileBaseName: generatePptxDisplayName(title).replace(/\.pptx$/i, ""),
        promptIntent,
        ...(palette ? { palette } : {}),
      }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error("[create_pptx] gen-pptx failed:", res.status, t);
      return { error: `PowerPoint生成に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();
    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
      displayName: generatePptxDisplayName(title),
      message: "PowerPoint file created successfully.",
    };
  } catch (e: any) {
    console.error("[create_pptx] error:", e);
    return { error: "PowerPoint生成中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- ドキュメント → PPTX 変換（Vision API使用） ----------------
async function executeConvertDocToPptx(
  args: {
    fileUrl: string;
    fileUrls?: string[];
    presentationTitle?: string;
    fontFace?: string;
    designInstruction?: string;
    maxPages?: number;
    mode?: "faithful" | "redesign";
  },
  chatThread: ChatThreadModel
) {
  const { fileUrl, fileUrls, presentationTitle, fontFace, designInstruction, maxPages, mode } = args ?? {};
  const sourceFileUrls = Array.from(
    new Set(
      [fileUrl, ...(Array.isArray(fileUrls) ? fileUrls : [])]
        .map((value) => normalizeDocumentUrlInput(value))
        .filter(Boolean)
    )
  );
  const derivedTitle = sourceFileUrls[0] ? extractPresentationTitleFromFileUrl(sourceFileUrls[0]) : null;
  // PDF→PPT変換はスレッド履歴からスタイルを引き継がない（各変換が独立）
  const explicitInstruction = designInstruction?.trim() || undefined;
  const deckPreferences: DeckPreferences = explicitInstruction
    ? { designInstruction: explicitInstruction }
    : {};

  if (sourceFileUrls.length === 0) {
    return { error: "fileUrl is required." };
  }

  const invalidFileUrl = sourceFileUrls.find((value) => !isHttpUrl(value));
  if (invalidFileUrl) {
    return {
      error: `fileUrl ??????'file_url:' ? 'fileUrl:' ?????????URL ????????????: ${invalidFileUrl}`,
    };
  }

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    if (sourceFileUrls.length > 1) {
      const mergedSlides: Array<{
        title: string;
        bullets: string[];
        layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram" | "conversation" | "stat_callouts" | "card_grid" | "icon_rows" | "metric-cards" | "process-cards" | "timeline" | "company-overview" | "closing";
        tableRows?: string[][];
        columns?: Array<{ header: string; bullets: string[] }>;
        conversationStyle?: "chat-ui" | "interview" | "dialog-list";
        conversationTurns?: Array<{
          speakerRole: string;
          speakerType?: "agent" | "customer" | "staff" | "other";
          text: string;
          turnIndex: number;
        }>;
      }> = [];
      let mergedTotalPages = 0;
      const perDocSummaries: string[] = [];

      for (const currentFileUrl of sourceFileUrls) {
        const resolvedFileUrl = await resolveDocumentUrlForVision(
          currentFileUrl,
          chatThread.id
        );
        console.log("[convert_doc_to_pptx] Analyzing document with Vision API:", {
          sourceFile: extractFileNameFromDocumentUrl(currentFileUrl),
          resolvedUrl: resolvedFileUrl.substring(0, 80),
        });
        const analyzeResult = await analyzeDocVision(resolvedFileUrl, maxPages ?? 30, mode);

        if (!analyzeResult?.ok || !analyzeResult.slides?.length) {
          console.error("[convert_doc_to_pptx] analyze-doc-vision failed:", analyzeResult?.error);
          return { error: analyzeResult?.error ?? "ドキュメント解析結果を取得できませんでした。" };
        }

        mergedSlides.push(...analyzeResult.slides);
        mergedTotalPages += analyzeResult.totalPages ?? analyzeResult.slides.length;
        perDocSummaries.push(buildDocSummaryFromSlides(
          extractFileNameFromDocumentUrl(currentFileUrl) ?? currentFileUrl,
          analyzeResult.slides
        ));
      }

      const mergedTitle =
        mergedSlides[0]?.title ||
        derivedTitle ||
        presentationTitle?.trim() ||
        "プレゼンテーション";

      console.log("[convert_doc_to_pptx] Title sources:", {
        derivedTitle,
        presentationTitle,
        deckPreferences,
        firstSlideTitle: mergedSlides[0]?.title,
        finalTitle: mergedTitle,
      });
      console.log("[convert_doc_to_pptx] Aggregated deck:", {
        fileCount: sourceFileUrls.length,
        totalPages: mergedTotalPages,
        slideCount: mergedSlides.length,
      });

      // 複数ドキュメントのマージ後: 経営向け再構築（四半期時系列ではなくカテゴリ軸に整理）
      const isExecutiveContext =
        /経営|役員|幹部|経営層|executive|management/i.test(
          [mergedTitle, designInstruction ?? ""].join(" ")
        ) ||
        (sourceFileUrls.length >= 2 &&
          /Q[1-4]|[1-4]Q|第[1-4]四半期|四半期|report|議事録|会議録/i.test(
            sourceFileUrls.join(" ")
          ));

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let finalMergedSlides: any[] = mergedSlides;
      if (mode !== "faithful" && isExecutiveContext && mergedSlides.length > 4) {
        console.log("[convert_doc_to_pptx] Executive context detected — running restructure pass");
        finalMergedSlides = await restructureSlidesForExecutive(
          mergedTitle,
          mergedSlides as unknown as RawPptSlide[],
          perDocSummaries,
          designInstruction
        );
      }

      const pptxRes = await fetch(`${baseUrl}/api/gen-pptx`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          title: mergedTitle,
          slides: finalMergedSlides,
          threadId: chatThread.id,
          fontFace,
          designInstruction: deckPreferences.designInstruction,
          deckPreferences,
          mode,
          fileBaseName: generatePptxDisplayName(mergedTitle).replace(/\.pptx$/i, ""),
        }),
      });

      if (!pptxRes.ok) {
        const t = await pptxRes.text().catch(() => "");
        console.error("[convert_doc_to_pptx] gen-pptx failed:", pptxRes.status, t);
        return { error: `PowerPoint生成に失敗しました: HTTP ${pptxRes.status}` };
      }

      const pptxResult = await pptxRes.json();
      if (!pptxResult?.downloadUrl) {
        return { error: "ダウンロードURLを取得できませんでした。" };
      }

      return {
        downloadUrl: pptxResult.downloadUrl,
        fileName: pptxResult.fileName,
        displayName: generatePptxDisplayName(mergedTitle),
        totalPages: mergedTotalPages,
        message: `${sourceFileUrls.length}件の資料をまとめて${mergedTotalPages}ページ分を解析し、PowerPointを生成しました。`,
      };
    }
    // Step 1: Vision API でドキュメントを解析してスライド構造を取得
    const resolvedFileUrl = await resolveDocumentUrlForVision(
      fileUrl,
      chatThread.id
    );
    console.log("[convert_doc_to_pptx] Analyzing document with Vision API:", resolvedFileUrl.substring(0, 80));
    const analyzeResult = await analyzeDocVision(resolvedFileUrl, maxPages ?? 30, mode);

    if (!analyzeResult?.ok || !analyzeResult.slides?.length) {
      console.error("[convert_doc_to_pptx] analyze-doc-vision failed:", analyzeResult?.error);
      return { error: analyzeResult?.error ?? "ドキュメントの解析結果が空でした。" };
    }

    const slides: Array<{
      title: string;
      bullets: string[];
      layoutType?: "title" | "bullets" | "table" | "multi-column" | "diagram" | "conversation";
      tableRows?: string[][];
      columns?: Array<{ header: string; bullets: string[] }>;
      conversationStyle?: "chat-ui" | "interview" | "dialog-list";
      conversationTurns?: Array<{
        speakerRole: string;
        speakerType?: "agent" | "customer" | "staff" | "other";
        text: string;
        turnIndex: number;
      }>;
    }> = analyzeResult.slides;
    const totalPages: number = analyzeResult.totalPages ?? slides.length;

    // タイトルを決定（指定がなければ最初のスライドのタイトルを使う）
    const title =
      slides[0]?.title ||
      derivedTitle ||
      presentationTitle?.trim() ||
      "プレゼンテーション";

    console.log("[convert_doc_to_pptx] Title sources:", {
      derivedTitle,
      presentationTitle,
      deckPreferences,
      firstSlideTitle: slides[0]?.title,
      finalTitle: title,
    });
    console.log(`[convert_doc_to_pptx] Analyzed ${totalPages} pages → ${slides.length} slides`);

    // Step 2: 解析結果から PPTX を生成
    const pptxRes = await fetch(`${baseUrl}/api/gen-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        title,
        slides,
        threadId: chatThread.id,
        fontFace,
        designInstruction: deckPreferences.designInstruction,
        deckPreferences,
        mode,
        fileBaseName: generatePptxDisplayName(title).replace(/\.pptx$/i, ""),
      }),
    });

    if (!pptxRes.ok) {
      const t = await pptxRes.text().catch(() => "");
      console.error("[convert_doc_to_pptx] gen-pptx failed:", pptxRes.status, t);
      return { error: `PowerPoint生成に失敗しました: HTTP ${pptxRes.status}` };
    }

    const pptxResult = await pptxRes.json();
    if (!pptxResult?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    return {
      downloadUrl: pptxResult.downloadUrl,
      fileName: pptxResult.fileName,
      displayName: generatePptxDisplayName(title),
      totalPages,
      message: `${totalPages}ページをVision APIで解析し、PowerPointファイルを生成しました。`,
    };
  } catch (e: any) {
    console.error("[convert_doc_to_pptx] error:", e);
    return { error: "変換中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- editLabel 抽出ヘルパー ----------------
function buildEditLabel(instruction: string): string {
  const cleaned = instruction.replace(/https?:\/\/\S+/g, "").replace(/（[^）]*）|\([^)]*\)/g, "");

  // ロゴ（画像URLがある場合も含む）
  if (/ロゴ|logo/i.test(instruction)) return "ロゴ追加";
  // 画像
  if (/画像|写真|イラスト|image|photo/i.test(cleaned)) return "画像追加";
  // 色・カラー + 具体的な色名（「文字色」「タイトル文字を赤に」も色変更として扱うため先に判定）
  if (/色|カラー|color|青|赤|緑|黄|白|黒|紫|オレンジ|ピンク|グレー|グリーン|ブルー|レッド/i.test(cleaned)) return "色変更";
  // フォント・フォントサイズ
  if (/フォント|font|文字サイズ|字体/i.test(cleaned)) return "フォント変更";
  // 文言・テキスト・文字変更
  if (/文言|テキスト|文字|コピー|見出し|タイトル|本文/i.test(cleaned)) return "文言修正";
  // レイアウト・構成
  if (/レイアウト|配置|構成|並び|整列|スライド追加|ページ追加/i.test(cleaned)) return "レイアウト変更";

  // フォールバック: 応答文語句を除去して短縮
  const stripped = cleaned
    .slice(0, 40)
    .replace(/以下|変更|行った|行いました|対応しました|確認ください|してください|して下さい|お願いします|てください|ください|します|しました|している|する|した/g, "")
    .replace(/[をにがはでのへとからまで（）()、。！!？?\s　]/g, "")
    .trim();
  return stripped.slice(0, 8) || "編集済み";
}

// ---------------- 既存 PPTX 改良 ----------------
async function executeEditPptx(
  args: { fileUrl?: string; instruction: string; imageUrl?: string },
  chatThread: ChatThreadModel
) {
  let { fileUrl, instruction, imageUrl: argImageUrl } = args ?? {};

  if (!instruction?.trim()) {
    return { error: "instructionは必須です。編集内容を指定してください。" };
  }

  // 画像URL解決: LLMがimageUrlを省略した場合のフォールバック
  // ロゴ/画像/添付の指示 かつ instruction にURLがない場合、スレッド最新アップロード画像URLを自動注入
  const needsImageUrl = /ロゴ|logo|画像|写真|添付|イラスト|image|photo/i.test(instruction);
  const resolvedImageUrl = argImageUrl?.trim() ||
    (needsImageUrl && !/https?:\/\//.test(instruction)
      ? (await resolveLatestImageUrlFromThread(chatThread.id)) ?? ""
      : "");
  if (resolvedImageUrl && !/https?:\/\//.test(instruction)) {
    instruction = `${resolvedImageUrl} ${instruction.trim()}`;
  }

  // ── 未対応操作の早期検出 ──────────────────────────────────────────────────
  // edit_pptx で実行できない操作が含まれる場合は即座に返却し、
  // LLMが「対応済み」と虚偽表示するのを防ぐ。
  const UNSUPPORTED_EDIT_PATTERNS: { re: RegExp; label: string }[] = [
    { re: /(?:新規|新しい|空白)?スライド(?![にへ上右左下])[^。、\n]{0,6}(追加|挿入)|(?:新規|新しい|空白)?ページ(?![にへ上右左下])[^。、\n]{0,6}(追加|挿入)|(追加|挿入)[^。、\n]{0,6}(?:新規|新しい|空白)?スライド/, label: "スライド追加・挿入" },
    { re: /空白.{0,8}スライド|空.{0,4}スライド|スライド.{0,4}空白|P\d+.{0,6}空|本文.{0,8}追加|箇条書き.{0,8}追加/, label: "空白スライドへの本文・箇条書き追加" },
    { re: /フォントサイズ|\d+\s*pt|\d+\s*ポイント|タイトル.{0,6}サイズ|文字.{0,4}(大き|小さ|サイズ)/, label: "フォントサイズ変更" },
    { re: /レイアウト.{0,6}最適化|重なり.{0,4}解消|配置.{0,4}(修正|変更|調整)|再レイアウト|位置.{0,4}調整/, label: "レイアウト最適化・shape移動" },
    { re: /スピーカーノート|ノート.{0,4}(追加|冒頭|末尾|記録)|speaker\s*note/i, label: "スピーカーノート追加" },
    { re: /再構成|作り直し|内容.{0,6}(整理|再生成|分離)|全体.{0,6}(見直し|修正|再生成)|を分ける|を分離/, label: "内容の再構成・作り直し" },
  ];
  const unsupportedFound = UNSUPPORTED_EDIT_PATTERNS.filter(({ re }) => re.test(instruction));
  if (unsupportedFound.length > 0) {
    const labels = unsupportedFound.map((u) => u.label).join("、");
    return {
      error: `この編集は既存PPTX編集では対応できません。PPTXを再生成する必要があります。\n\n未対応の要求: ${labels}\n\n対応可能な編集: ロゴ・画像挿入、アクセントカラー変更、既存文字列の置換`,
    };
  }

  // Markdownリンクの表示名を優先取得（fileUrl + displayName を一括解決）
  const originalFileUrl = fileUrl?.trim() ?? "";
  const threadPptxInfo = await resolveLatestPptxInfoFromThread(chatThread.id);
  if (!fileUrl?.trim()) {
    fileUrl = threadPptxInfo?.url ?? "";
  }

  if (!fileUrl?.trim()) {
    return {
      error:
        "編集対象のPPTXが見つかりませんでした。このスレッドでPPTXを生成するか、PPTのURLを指定してください。",
    };
  }

  const editLabel = buildEditLabel(instruction);

  // 元ファイル名: Markdownリンクの表示名は fileUrl が一致する場合のみ使用
  // SASクエリは異なっても同一Blobなら一致とみなすため origin + pathname(decode) で比較する
  const blobKey = (u: string) => { try { const p = new URL(u); return (p.origin + decodeURIComponent(p.pathname)).toLowerCase(); } catch { return u; } };
  const isSameAsThreadPptx = !originalFileUrl || blobKey(originalFileUrl) === blobKey(threadPptxInfo?.url ?? "");
  const inputBaseName = (isSameAsThreadPptx ? threadPptxInfo?.displayName : null) ??
    (() => {
      try {
        const urlPath = new URL(fileUrl).pathname;
        const decoded = decodeURIComponent(urlPath.split("/").pop() ?? "");
        const base = decoded
          .replace(/\.[^.]+$/, "")
          .replace(/_edited_[A-Za-z0-9]{6,12}$/i, "")
          .replace(/_[A-Za-z0-9]{6,12}$/, "")
          .trim();
        return /^pptx$/i.test(base) ? "" : base;
      } catch { return ""; }
    })();

  const outputBaseName = inputBaseName ? `${inputBaseName}_${editLabel}` : editLabel;

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    const res = await fetch(`${baseUrl}/api/edit-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fileUrl, instruction, threadId: chatThread.id, outputBaseName }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error("[edit_pptx] edit-pptx failed:", res.status, t);
      return { error: `PPTX編集に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();
    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    const baseMessage = `${result.changedSlides}枚のスライドを編集しました（全${result.totalSlides}枚）。`;
    const imageMessage =
      result.requestedImages > 0
        ? result.insertedImages === result.requestedImages
          ? `画像${result.insertedImages}件を挿入しました。`
          : `⚠️ ${result.imageWarning}`
        : "";

    const editDisplayName = `${outputBaseName}.pptx`;

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
      displayName: editDisplayName,
      changedSlides: result.changedSlides,
      totalSlides: result.totalSlides,
      message: imageMessage ? `${baseMessage} ${imageMessage}` : baseMessage,
    };
  } catch (e: any) {
    console.error("[edit_pptx] error:", e);
    return { error: "PPTX編集中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- Excel 編集 ----------------
async function executeEditExcel(
  args: { fileUrl?: string; instruction: string; previousChartEdits?: object[]; sourceFileQuery?: string },
  chatThread: ChatThreadModel
) {
  const { fileUrl, instruction, previousChartEdits: llmPreviousChartEdits, sourceFileQuery } = args ?? {};

  if (!instruction?.trim()) {
    return { error: "instructionは必須です。編集内容を指定してください。" };
  }

  if (!fileUrl?.trim()) {
    return {
      error:
        "編集対象のExcelファイルが見つかりませんでした。このスレッドでExcelファイルをアップロードしてください。",
    };
  }

  // LLMが previousChartEdits を渡さなかった場合はポインタから自動補完（LLM依存を排除）
  const existingPtr = await readLatestExcelPtr(chatThread.id);
  const previousChartEdits = llmPreviousChartEdits?.length
    ? llmPreviousChartEdits
    : existingPtr?.chartEdits;
  if (previousChartEdits?.length) {
    console.log(`[edit_excel] previousChartEdits: ${previousChartEdits.length} entries (source: ${llmPreviousChartEdits?.length ? "LLM" : "pointer"})`);
  }

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    const res = await fetch(`${baseUrl}/api/edit-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fileUrl, instruction, threadId: chatThread.id, previousChartEdits }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error("[edit_excel] edit-pptx route failed:", res.status, t);
      return { error: `Excel編集に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();
    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    // ポインタ更新（sourceFileQuery は引数優先 → 既存ポインタ引き継ぎ、chartEdits も保持）
    await saveLatestExcelUrl(
      chatThread.id,
      result.downloadUrl,
      result.fileName ?? "edited.xlsx",
      sourceFileQuery ?? existingPtr?.sourceFileQuery,
      result.appliedChartEdits ?? existingPtr?.chartEdits
    );

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
      changedSheets: result.changedSheets,
      totalSheets: result.totalSheets,
      message: `${result.changedSheets}シートを編集しました（全${result.totalSheets}シート）。`,
      ...(result.appliedChartEdits ? { appliedChartEdits: result.appliedChartEdits } : {}),
    };
  } catch (e: any) {
    console.error("[edit_excel] error:", e);
    return { error: "Excel編集中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- Excel 新規作成 ----------------
async function executeCreateExcel(
  args: { content: string; title?: string; instruction?: string },
  chatThread: ChatThreadModel
) {
  const { content, title, instruction } = args ?? {};

  if (!content?.trim() && !title?.trim()) {
    return { error: "content を指定してください。作成するデータを入力してください。" };
  }

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    const res = await fetch(`${baseUrl}/api/gen-excel`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        content: content ?? "",
        title: title ?? "",
        instruction: instruction ?? "",
        threadId: chatThread.id,
      }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error("[create_excel] gen-excel failed:", res.status, t);
      return { error: `Excel作成に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();
    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
      sheets: result.sheets,
      totalRows: result.totalRows,
      message: `Excelファイルを作成しました（${result.sheets}シート、${result.totalRows}行）。`,
    };
  } catch (e: any) {
    console.error("[create_excel] error:", e);
    return { error: "Excel作成中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- Word 新規作成 ----------------
async function executeCreateWord(
  args: { content: string; title?: string; instruction?: string; fontFace?: string },
  chatThread: ChatThreadModel
) {
  const { content, title, instruction, fontFace } = args ?? {};

  if (!content?.trim() && !title?.trim()) {
    return { error: "content を指定してください。作成する内容を入力してください。" };
  }

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    const res = await fetch(`${baseUrl}/api/gen-word`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        content: content ?? "",
        title: title ?? "",
        instruction: instruction ?? "",
        fontFace: fontFace ?? "Meiryo",
        threadId: chatThread.id,
      }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error("[create_word] gen-word failed:", res.status, t);
      return { error: `Word作成に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();
    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
      paragraphs: result.paragraphs,
      tables: result.tables,
      message: `Wordファイルを作成しました（${result.paragraphs}段落、テーブル${result.tables}個）。`,
    };
  } catch (e: any) {
    console.error("[create_word] error:", e);
    return { error: "Word作成中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- Word 編集 ----------------
async function executeEditWord(
  args: { fileUrl?: string; instruction: string },
  chatThread: ChatThreadModel
) {
  const { fileUrl, instruction } = args ?? {};

  if (!instruction?.trim()) {
    return { error: "instructionは必須です。編集内容を指定してください。" };
  }

  if (!fileUrl?.trim()) {
    return {
      error:
        "編集対象のWordファイルが見つかりませんでした。このスレッドでWordファイルをアップロードしてください。",
    };
  }

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    const res = await fetch(`${baseUrl}/api/edit-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fileUrl, instruction, threadId: chatThread.id }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error("[edit_word] edit-pptx route failed:", res.status, t);
      return { error: `Word編集に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();
    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
      changedParagraphs: result.changedParagraphs,
      totalParagraphs: result.totalParagraphs,
      message: `${result.changedParagraphs}箇所を編集しました（全${result.totalParagraphs}段落）。`,
    };
  } catch (e: any) {
    console.error("[edit_word] error:", e);
    return { error: "Word編集中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- PDF → Excel 変換 ----------------
async function executeConvertPdfToExcel(
  args: { fileUrl?: string },
  chatThread: ChatThreadModel
) {
  let { fileUrl } = args ?? {};

  if (!fileUrl?.trim()) {
    fileUrl = (await resolveLatestPdfOrDocxUrlFromThread(chatThread.id)) ?? "";
  }

  if (!fileUrl?.trim()) {
    return {
      error:
        "変換対象のPDF/Wordファイルが見つかりませんでした。このスレッドでPDFまたはWordファイルをアップロードしてください。",
    };
  }

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    const res = await fetch(`${baseUrl}/api/edit-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fileUrl, instruction: "", threadId: chatThread.id, action: "pdf_to_excel" }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error("[convert_pdf_to_excel] route failed:", res.status, t);
      return { error: `PDF→Excel変換に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();

    // 画像埋め込み型 Word（EMF 等）で抽出不可だった場合
    if (result?.engine === "none") {
      return {
        error:
          "このWordファイルは画像埋め込み型のため、表データを抽出できませんでした。\n" +
          "WordをPDF形式で保存してからアップロードし、再度「Excelに変換して」とお試しください。",
      };
    }

    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    const tableInfo = result.tables > 0
      ? `テーブル${result.tables}個を${result.sheets}シートに変換`
      : `テキストを「Text」シートに出力`;

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
      sheets: result.sheets,
      tables: result.tables,
      pages: result.pages,
      message: `${result.pages}ページを変換しました（${tableInfo}）。`,
    };
  } catch (e: any) {
    console.error("[convert_pdf_to_excel] error:", e);
    return { error: "PDF→Excel変換中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- PDF → Word 変換 ----------------
async function executeConvertPdfToWord(
  args: { fileUrl?: string; mode?: "layout" | "editable" },
  chatThread: ChatThreadModel
) {
  const { fileUrl, mode = "layout" } = args ?? {};

  if (!fileUrl?.trim()) {
    return {
      error: "変換対象のPDFファイルが見つかりませんでした。このスレッドでPDFファイルをアップロードしてください。",
    };
  }

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    const res = await fetch(`${baseUrl}/api/edit-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fileUrl, instruction: "", threadId: chatThread.id, action: "pdf_to_word", mode }),
    });

    if (!res.ok) {
      const t = await res.text().catch(() => "");
      console.error("[convert_pdf_to_word] route failed:", res.status, t);
      return { error: `PDF→Word変換に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();

    if (result?.engine === "none") {
      return {
        error: "PDFの変換に失敗しました。スキャン画像のみのPDFの場合はテキスト抽出ができません。",
      };
    }

    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
      paragraphs: result.paragraphs,
      tables: result.tables,
      message: `PDFをWordに変換しました（段落${result.paragraphs}件、表${result.tables}件）。`,
    };
  } catch (e: any) {
    console.error("[convert_pdf_to_word] error:", e);
    return { error: "PDF→Word変換中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- SharePoint SL文書 → PPT変換 ----------------
async function executeConvertSpToPptx(
  args: { fileQuery: string; mode?: "faithful" | "redesign" },
  chatThread: ChatThreadModel
) {
  const { fileQuery, mode } = args ?? {};
  console.log(`[convert_sp_to_pptx] called with fileQuery="${fileQuery}" mode=${mode}`);

  if (!fileQuery?.trim()) {
    return { error: "fileQuery（ファイル名またはキーワード）を指定してください。" };
  }

  // 現在ユーザーの部署情報を取得してACLフィルタに渡す
  const currentUser = await userSession();
  const deptLower = currentUser?.slDept?.toLowerCase() ?? undefined;

  // AI Search でアクセス可能な全SL文書を取得（"*"検索）し、クライアント側でファイル名フィルタ
  // ※ fileQuery をページ本文テキスト検索に使うとファイル名がヒットしない場合があるため
  // top: 200 でページネーション無限ループ・接続 aborted を防止
  const searchResult = await SimpleSearch("*", "isSlDoc eq true", deptLower, 200);
  const searchCount =
    searchResult.status === "OK" ? searchResult.response.length : 0;
  console.log(
    `[convert_sp_to_pptx] SimpleSearch returned status=${searchResult.status} count=${searchCount}`
  );

  if (searchResult.status !== "OK" || !searchResult.response.length) {
    return { error: "アクセス可能なSharePointファイルが見つかりませんでした。" };
  }

  const allDocs = searchResult.response;

  // ファイル名でクライアント側フィルタリング（部分一致・大文字小文字無視）
  const queryLower = fileQuery.trim().toLowerCase();
  const matched = allDocs.filter(({ document: doc }) => {
    const name = (doc.metadata ?? "").toLowerCase();
    return name.includes(queryLower) || queryLower.includes(name.replace(/\.pdf$/i, ""));
  });

  console.log(`[convert_sp_to_pptx] name-matched count=${matched.length} (query="${fileQuery}")`);

  if (!matched.length) {
    // フォールバック: 全候補を提示
    const allFiles = Array.from(
      new Map(
        allDocs.map(({ document: doc }) => [
          doc.effectiveFileUrl || doc.fileUrl,
          doc.metadata || "不明",
        ])
      ).entries()
    );
    const list = allFiles.map(([, name], i) => `${i + 1}. ${name}`).join("\n");
    return {
      multipleFiles: true,
      message: `「${fileQuery}」に一致するファイルが見つかりませんでした。\nアクセス可能なSLファイル一覧です：\n\n${list}\n\nファイル名を指定してください。`,
    };
  }

  // URLをキーにしてユニークファイルを抽出（同名ファイルが別フォルダにある場合を考慮）
  const seen = new Map<string, { fileName: string; url: string }>();
  for (const { document: doc } of matched) {
    const url = doc.effectiveFileUrl || doc.fileUrl;
    const name = doc.metadata || url.split("/").pop() || "file";
    if (!seen.has(url)) seen.set(url, { fileName: name, url });
  }

  const candidates = Array.from(seen.values());

  // 複数ファイルがヒットした場合はリスト返却
  if (candidates.length > 1) {
    const list = candidates
      .map((c, i) => `${i + 1}. ${c.fileName}`)
      .join("\n");
    return {
      multipleFiles: true,
      message: `「${fileQuery}」で複数のファイルが見つかりました。どれを変換しますか？\n\n${list}\n\nファイル名を指定して再度お試しください。`,
    };
  }

  const { fileName, url } = candidates[0];

  // PDF以外は変換不可
  if (!fileName.toLowerCase().endsWith(".pdf")) {
    const ext = fileName.split(".").pop()?.toUpperCase() ?? "不明";
    const hint =
      ext === "PPTX" || ext === "PPT"
        ? "（すでにPowerPointファイルです）"
        : ext === "DOCX" || ext === "DOC"
        ? "（WordファイルはPPT変換に対応していません）"
        : "";
    return { error: `「${fileName}」はPDFファイルではないため、PPTに変換できません。${hint}` };
  }

  console.log(`[convert_sp_to_pptx] Converting SP file: ${fileName}`);
  console.log("[convert_sp_to_pptx] original url =", url.substring(0, 100));

  const baseUrl = (
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : "http://localhost:3000")
  ).replace(/\/+$/, "");

  try {
    // Step 1: SP URL → Blob SAS URL に解決（Graph API経由キャッシュ含む）
    const resolvedUrl = await resolveDocumentUrlForVision(url, chatThread.id);
    console.log("[convert_sp_to_pptx] resolved url =", resolvedUrl.substring(0, 100));

    // Step 2: Vision API でPDF解析
    const analyzeResult = await analyzeDocVision(resolvedUrl, 30, mode);
    if (!analyzeResult?.ok || !analyzeResult.slides?.length) {
      console.error("[convert_sp_to_pptx] analyze-doc-vision failed:", analyzeResult?.error);
      return { error: analyzeResult?.error ?? "PDFの解析に失敗しました。" };
    }

    const { slides, totalPages } = analyzeResult;
    const title = slides[0]?.title || fileName.replace(/\.pdf$/i, "");

    console.log(`[convert_sp_to_pptx] Analyzed ${totalPages} pages → ${slides.length} slides`);

    // Step 2: PPTX 生成
    const pptxRes = await fetch(`${baseUrl}/api/gen-pptx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        slides,
        title,
        threadId: chatThread.id,
        deckPreferences: {},
        mode,
        fileBaseName: generatePptxDisplayName(title).replace(/\.pptx$/i, ""),
      }),
    });

    if (!pptxRes.ok) {
      const t = await pptxRes.text().catch(() => "");
      console.error("[convert_sp_to_pptx] gen-pptx failed:", pptxRes.status, t);
      return { error: `PowerPoint生成に失敗しました: HTTP ${pptxRes.status}` };
    }

    const pptxResult = await pptxRes.json();
    if (!pptxResult?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    return {
      downloadUrl: pptxResult.downloadUrl,
      fileName: pptxResult.fileName,
      displayName: generatePptxDisplayName(title),
      totalPages,
      message: `SharePointの「${fileName}」（${totalPages}ページ）をPowerPointに変換しました。`,
    };
  } catch (e: any) {
    console.error("[convert_sp_to_pptx] error:", e);
    return { error: "変換中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- SharePoint SL の PPTX を編集 ----------------
async function executeEditSpPptx(
  args: { fileQuery: string; instruction: string },
  chatThread: ChatThreadModel
) {
  const { fileQuery, instruction } = args ?? {};

  if (!fileQuery?.trim()) return { error: "fileQuery（ファイル名またはキーワード）を指定してください。" };
  if (!instruction?.trim()) return { error: "instruction（編集内容）を指定してください。" };

  // 1. AI Search でアクセス可能な全 SL 文書を取得し、クライアント側でフィルタ
  const currentUser = await userSession();
  const deptLower = currentUser?.slDept?.toLowerCase() ?? undefined;

  const searchResult = await SimpleSearch("*", "isSlDoc eq true", deptLower, 200);
  if (searchResult.status !== "OK" || !searchResult.response.length) {
    return { error: "アクセス可能なSharePointファイルが見つかりませんでした。" };
  }

  // 2. PPTX ファイルをファイル名でフィルタ
  const queryLower = fileQuery.trim().toLowerCase();
  const matched = searchResult.response.filter(({ document: doc }) => {
    const name = (doc.metadata ?? "").toLowerCase();
    return (
      name.endsWith(".pptx") &&
      (name.includes(queryLower) || queryLower.includes(name.replace(/\.pptx$/i, "")))
    );
  });

  console.log(`[edit_sp_pptx] pptx-matched count=${matched.length} (query="${fileQuery}")`);

  if (!matched.length) {
    return { error: `「${fileQuery}」に一致するPPTXファイルが見つかりませんでした。` };
  }

  // 3. URL でユニーク化（同一ファイルが複数チャンクとして登録されている場合を考慮）
  const seen = new Map<string, { fileName: string; sourceUrl: string; effectiveFileUrl: string | null }>();
  for (const { document: doc } of matched) {
    const key = doc.effectiveFileUrl || doc.fileUrl;
    if (key && !seen.has(key)) {
      seen.set(key, {
        fileName: doc.metadata ?? "",
        sourceUrl: doc.fileUrl,
        effectiveFileUrl: doc.effectiveFileUrl ?? null,
      });
    }
  }

  const candidates = Array.from(seen.values());

  const uniqueFileNamesPptx = new Set(
    candidates.map((c) => c.fileName.toLowerCase().replace(/\.pptx$/i, ""))
  );

  let chosenPptx = candidates[0];
  if (candidates.length > 1) {
    if (uniqueFileNamesPptx.size === 1) {
      console.log(`[edit_sp_pptx] ${candidates.length} duplicates of "${candidates[0].fileName}" found — auto-selecting first (highest relevance)`);
    } else {
      const list = Array.from(uniqueFileNamesPptx).map((n, i) => `${i + 1}. ${n}`).join("\n");
      return {
        multipleFiles: true,
        message: `「${fileQuery}」で複数の異なるファイルが見つかりました。どれを編集しますか？\n\n${list}\n\nファイル名を指定して再度お試しください。`,
      };
    }
  }

  const { fileName, sourceUrl, effectiveFileUrl } = chosenPptx;
  console.log(`[edit_sp_pptx] target: ${fileName} sourceUrl=${sourceUrl.substring(0, 100)}`);

  // 4. SAS URL を解決する
  //    優先順位: ① effectiveFileUrl が Blob raw URL → GenerateSasUrl
  //             ② SP 直パス URL → downloadSharePointFileToBlob (Graph API)
  let resolvedUrl: string | null = null;

  // ① effectiveFileUrl が SAS なし Blob URL の場合
  const blobParsed = parseBlobRawUrl(effectiveFileUrl);
  if (blobParsed) {
    const sasRes = await GenerateSasUrl(blobParsed.container, blobParsed.path);
    if (sasRes.status === "OK" && sasRes.response) {
      resolvedUrl = sasRes.response;
      console.log(`[edit_sp_pptx] Resolved via GenerateSasUrl: ${blobParsed.path}`);
    }
  }

  // ② SP URL → Graph API でダウンロードしてキャッシュ
  if (!resolvedUrl) {
    const urlForDownload = effectiveFileUrl || sourceUrl;
    const spSas = await downloadSharePointFileToBlob(urlForDownload, chatThread.id, fileName);
    if (spSas) {
      resolvedUrl = spSas;
      console.log(`[edit_sp_pptx] Resolved via Graph API download`);
    }
  }

  if (!resolvedUrl) {
    console.warn(`[edit_sp_pptx] Could not resolve to blob URL:`, sourceUrl);
    return { error: `「${fileName}」のダウンロードURLを取得できませんでした。` };
  }

  // 5. edit-pptx API に委託
  return executeEditPptx({ fileUrl: resolvedUrl, instruction }, chatThread);
}

// ---------------- SharePoint SL の Excel を編集 ----------------
async function executeEditSpExcel(
  args: { fileQuery: string; instruction: string; previousChartEdits?: object[] },
  chatThread: ChatThreadModel
) {
  const { fileQuery, instruction, previousChartEdits } = args ?? {};

  if (!fileQuery?.trim()) return { error: "fileQuery（ファイル名またはキーワード）を指定してください。" };
  if (!instruction?.trim()) return { error: "instruction（編集内容）を指定してください。" };

  // 0. このスレッドで前回編集した Blob ポインタがあり、fileQuery と同名なら SP 再取得をスキップ
  //    照合は sourceFileQuery（元のSPファイル名）を優先する（編集済みファイル名は "_edited_" が入るため）
  const ptr = await readLatestExcelPtr(chatThread.id);
  if (ptr?.url) {
    const ptrMatch = (ptr.sourceFileQuery ?? ptr.fileName).toLowerCase().replace(/\.(xlsx|xls|xlsm)$/i, "");
    const queryBase = fileQuery.trim().toLowerCase().replace(/\.(xlsx|xls|xlsm)$/i, "");
    if (ptrMatch.includes(queryBase) || queryBase.includes(ptrMatch)) {
      console.log(`[edit_sp_excel] Using saved blob URL for "${ptr.fileName}" (source: "${ptr.sourceFileQuery ?? "-"}", skipping SP fetch)`);
      return executeEditExcel({ fileUrl: ptr.url, instruction, previousChartEdits, sourceFileQuery: fileQuery }, chatThread);
    }
  }

  // 1. fileQuery でテキスト検索 + SL文書フィルタ（200件制限を回避するためクエリで絞る）
  const currentUser = await userSession();
  const deptLower = currentUser?.slDept?.toLowerCase() ?? undefined;

  const searchResult = await SimpleSearch(fileQuery, "isSlDoc eq true", deptLower, 50);
  if (searchResult.status !== "OK" || !searchResult.response.length) {
    return { error: "アクセス可能なSharePointファイルが見つかりませんでした。" };
  }

  // 2. Excel ファイルをファイル名でフィルタ
  //    metadata が空/別形式の場合に備えて fileUrl / effectiveFileUrl からもファイル名を取得する
  const queryLower = fileQuery.trim().toLowerCase();
  const matched = searchResult.response.filter(({ document: doc }) => {
    const metaName = (doc.metadata ?? "").trim().toLowerCase();
    const urlName = (extractFileNameFromDocumentUrl(doc.effectiveFileUrl || doc.fileUrl) ?? "").toLowerCase();
    // resolvedName と同じロジック: metaName がExcel拡張子付きなら採用、そうでなければ urlName
    const name = /\.(xlsx|xls|xlsm)$/i.test(metaName) ? metaName : (urlName || metaName);
    return (
      /\.(xlsx|xls|xlsm)$/i.test(name) &&
      (name.includes(queryLower) || queryLower.includes(name.replace(/\.(xlsx|xls|xlsm)$/i, "")))
    );
  });

  console.log(`[edit_sp_excel] xlsx-matched count=${matched.length} (query="${fileQuery}")`);

  if (!matched.length) {
    return { error: `「${fileQuery}」に一致するExcelファイルが見つかりませんでした。` };
  }

  // 3. URL でユニーク化
  const seen = new Map<string, { fileName: string; sourceUrl: string; effectiveFileUrl: string | null }>();
  for (const { document: doc } of matched) {
    const key = doc.effectiveFileUrl || doc.fileUrl;
    if (key && !seen.has(key)) {
      // metadata がExcel拡張子付きファイル名の場合に採用、それ以外は URL から取得する
      const metaName = (doc.metadata ?? "").trim();
      const urlName = extractFileNameFromDocumentUrl(doc.effectiveFileUrl || doc.fileUrl) ?? "";
      const resolvedName = /\.(xlsx|xls|xlsm)$/i.test(metaName) ? metaName : (urlName || metaName);
      seen.set(key, {
        fileName: resolvedName,
        sourceUrl: doc.fileUrl,
        effectiveFileUrl: doc.effectiveFileUrl ?? null,
      });
    }
  }

  const candidates = Array.from(seen.values());

  // ファイル名（拡張子除く）でグループ化し、同一名が複数あれば最初の1件を自動選択
  // 異なるファイル名が複数ある場合のみユーザーに選択を促す
  const uniqueFileNames = new Set(
    candidates.map((c) => c.fileName.toLowerCase().replace(/\.(xlsx|xls|xlsm)$/i, ""))
  );

  let chosen = candidates[0];
  if (candidates.length > 1) {
    if (uniqueFileNames.size === 1) {
      // 同じファイルの重複アップロード → 検索スコア最高（先頭）を使用
      console.log(`[edit_sp_excel] ${candidates.length} duplicates of "${candidates[0].fileName}" found — auto-selecting first (highest relevance)`);
    } else {
      // 本当に異なるファイルが複数ある → ユーザーに確認
      const list = Array.from(uniqueFileNames).map((n, i) => `${i + 1}. ${n}`).join("\n");
      return {
        multipleFiles: true,
        message: `「${fileQuery}」で複数の異なるファイルが見つかりました。どれを編集しますか？\n\n${list}\n\nファイル名を指定して再度お試しください。`,
      };
    }
  }

  const { fileName, sourceUrl, effectiveFileUrl } = chosen;
  console.log(`[edit_sp_excel] target: ${fileName} sourceUrl=${sourceUrl.substring(0, 100)}`);

  // 4. SAS URL を解決する
  //    優先順位: ① effectiveFileUrl が Blob raw URL → GenerateSasUrl
  //             ② SP 直パス URL → downloadSharePointFileToBlob (Graph API)
  let resolvedUrl: string | null = null;

  const blobParsed = parseBlobRawUrl(effectiveFileUrl);
  if (blobParsed) {
    const sasRes = await GenerateSasUrl(blobParsed.container, blobParsed.path);
    if (sasRes.status === "OK" && sasRes.response) {
      resolvedUrl = sasRes.response;
      console.log(`[edit_sp_excel] Resolved via GenerateSasUrl: ${blobParsed.path}`);
    }
  }

  if (!resolvedUrl) {
    const urlForDownload = effectiveFileUrl || sourceUrl;
    const spSas = await downloadSharePointFileToBlob(urlForDownload, chatThread.id, fileName);
    if (spSas) {
      resolvedUrl = spSas;
      console.log(`[edit_sp_excel] Resolved via Graph API download`);
    }
  }

  if (!resolvedUrl) {
    console.warn(`[edit_sp_excel] Could not resolve to blob URL:`, sourceUrl);
    return { error: `「${fileName}」のダウンロードURLを取得できませんでした。` };
  }

  // 5. edit_excel に委託（sourceFileQuery を渡してポインタ保存を集約）
  return executeEditExcel({ fileUrl: resolvedUrl, instruction, previousChartEdits, sourceFileQuery: fileQuery }, chatThread);
}

// ---------------- 画像生成（NEW image 用） ----------------
async function executeCreateImage(
  args: { prompt: string; text?: string; size?: string },
  chatThread: ChatThreadModel,
  signal?: AbortSignal,
  modeOpts?: {
    reasoning_effort?: "minimal" | "medium" | "high";
    temperature?: number;
  }
) {
  const prompt = (args?.prompt || "").trim();

  console.log("createImage called with prompt:", prompt);

  if (!prompt) return "No prompt provided";
  if (prompt.length >= 4000)
    return "Prompt is too long, it must be less than 4000 characters";

  const openAI = OpenAIDALLEInstance();

  let response;
  try {
    response = await openAI.images.generate(
      {
        model: "gpt-image-1.5",
        prompt,
      },
      { signal }
    );
  } catch (error) {
    console.error("🔴 error while calling Azure image gen:\n", error);
    return { error: "There was an error creating the image: " + error };
  }

  if (!response.data?.[0]?.b64_json) {
    return { error: "Invalid API response: no b64_json." };
  }

  try {
    const imageName = `${uniqueId()}.png`;
    const buffer = Buffer.from(response.data[0].b64_json, "base64");

    await UploadImageToStore(chatThread.id, imageName, buffer);
    await UploadImageToStore(chatThread.id, "__base__.png", buffer);

    lastTextLayoutByThread.delete(chatThread.id);
    console.log("🗑️ Cleared text layout for thread:", chatThread.id);

    const baseImageUrl = buildExternalImageUrl(chatThread.id, imageName);
    return { revised_prompt: prompt, url: baseImageUrl };
  } catch (error) {
    console.error("🔴 error while storing image:\n", error);
    return { error: "There was an error storing the image: " + error };
  }
}

// ---------------- 既存画像への文字追加（EDIT 用・Vision 不使用） ----------------
async function executeAddTextToExistingImage(
  args: {
    imageUrl: string;
    text: string;
    styleHint?: string;
    font?: string;
    color?: string;
    size?: string;
    offsetX?: number;
    offsetY?: number;
  },
  chatThread: ChatThreadModel,
  userMessage: string,
  signal: AbortSignal,
  modeOpts?: {
    reasoning_effort?: "minimal" | "medium" | "high";
    temperature?: number;
  }
) {
  const explicitUrl = (args?.imageUrl || "").trim();
  let text = (args?.text || "").trim();
  const styleHint = (args?.styleHint || "").trim();

  const baseImageUrl = buildExternalImageUrl(chatThread.id, "__base__.png");

  console.log("🗺️ lastTextLayoutByThread MAP状態:", {
    threadId: chatThread.id,
    hasEntry: lastTextLayoutByThread.has(chatThread.id),
    mapSize: lastTextLayoutByThread.size,
    allKeys: Array.from(lastTextLayoutByThread.keys()),
    currentValue: lastTextLayoutByThread.get(chatThread.id),
  });

  console.log("🖋 add_text_to_existing_image called:", {
    passedImageUrl: explicitUrl,
    usedBaseImageUrl: baseImageUrl,
    text,
    styleHint,
    argsOffsetX: args?.offsetX,
    argsOffsetY: args?.offsetY,
  });

  if (!text) {
    return { error: "text is required for add_text_to_existing_image." };
  }

  const hintSource = styleHint || userMessage || "";
  const parsed = parseStyleHint(hintSource);

  console.log("🔍 parsed style hint:", parsed);

  const last = lastTextLayoutByThread.get(chatThread.id);
  console.log("📍 last layout from Map:", last);

  if (last?.text && text !== last.text) {
    console.warn("⚠️ Text content changed:", {
      previous: last.text,
      current: text,
      userMessage,
    });

    const lowerMsg = (userMessage || "").toLowerCase();
    const isExplicitChange =
      lowerMsg.includes("変更") ||
      lowerMsg.includes("変える") ||
      lowerMsg.includes("書き換え");

    if (!isExplicitChange) {
      console.warn("⚠️⚠️ Text changed without explicit request. Using previous text.");
      text = last.text;
    }
  }

  const align: "left" | "center" | "right" =
    parsed.align !== undefined ? parsed.align : last?.align ?? "center";

  const vAlign: "top" | "middle" | "bottom" =
    parsed.vAlign !== undefined ? parsed.vAlign : last?.vAlign ?? "middle";

  console.log("✅ resolved align/vAlign:", { align, vAlign });

  let size: "small" | "medium" | "large" | "xlarge" =
    (args.size as any) ?? parsed.size ?? last?.size ?? "large";

  if (parsed.sizeAdjust === "larger") {
    const sizeOrder: Array<"small" | "medium" | "large" | "xlarge"> = [
      "small",
      "medium",
      "large",
      "xlarge",
    ];
    const currentIndex = sizeOrder.indexOf(size);
    if (currentIndex >= 0 && currentIndex < sizeOrder.length - 1) {
      const oldSize = size;
      size = sizeOrder[currentIndex + 1];
      console.log(`📏 Size adjusted larger: ${oldSize} → ${size}`);
    }
  } else if (parsed.sizeAdjust === "smaller") {
    const sizeOrder: Array<"small" | "medium" | "large" | "xlarge"> = [
      "small",
      "medium",
      "large",
      "xlarge",
    ];
    const currentIndex = sizeOrder.indexOf(size);
    if (currentIndex > 0) {
      const oldSize = size;
      size = sizeOrder[currentIndex - 1];
      console.log(`📏 Size adjusted smaller: ${oldSize} → ${size}`);
    }
  }

  const color = args.color ?? parsed.color ?? last?.color ?? "white";

  console.log("🎨 color resolution:", {
    argsColor: args.color,
    parsedColor: parsed.color,
    lastColor: last?.color,
    finalColor: color,
  });

  const fontHint = (
    (styleHint || "") +
    " " +
    (args.font || "") +
    " " +
    (parsed.font || "")
  ).toLowerCase();

  let fontFamily: "gothic" | "mincho" | "meiryo" =
    last?.fontFamily ?? "gothic";

  if (
    fontHint.includes("明朝") ||
    fontHint.includes("mincho") ||
    fontHint.includes("serif")
  ) {
    fontFamily = "mincho";
  } else if (fontHint.includes("メイリオ") || fontHint.includes("meiryo")) {
    fontFamily = "meiryo";
  } else if (fontHint.includes("ゴシック") || fontHint.includes("gothic")) {
    fontFamily = "gothic";
  }

  console.log("🔤 fontFamily resolution:", {
    fontHint,
    lastFontFamily: last?.fontFamily,
    finalFontFamily: fontFamily,
  });

  const lowerHintAll = (hintSource || "").toLowerCase();

  const boldOff =
    hintSource.includes("太字やめ") ||
    hintSource.includes("太字解除") ||
    hintSource.includes("太字をやめ") ||
    hintSource.includes("太字を解除") ||
    hintSource.includes("通常") ||
    lowerHintAll.includes("not bold") ||
    lowerHintAll.includes("no bold");

  const italicOff =
    hintSource.includes("斜体やめ") ||
    hintSource.includes("斜体解除") ||
    hintSource.includes("イタリックやめ") ||
    hintSource.includes("イタリック解除") ||
    hintSource.includes("斜体をやめ") ||
    hintSource.includes("斜体を解除") ||
    lowerHintAll.includes("not italic") ||
    lowerHintAll.includes("no italic");

  const boldOn =
    hintSource.includes("太字") ||
    hintSource.includes("ボールド") ||
    lowerHintAll.includes("bold");

  const italicOn =
    hintSource.includes("イタリック") ||
    hintSource.includes("斜体") ||
    lowerHintAll.includes("italic");

  const bold = boldOff ? false : boldOn ? true : (last?.bold ?? false);
  const italic = italicOff ? false : italicOn ? true : (last?.italic ?? false);

  console.log("📝 bold/italic resolution:", {
    lastBold: last?.bold,
    lastItalic: last?.italic,
    finalBold: bold,
    finalItalic: italic,
  });

  const positionSpecified =
    parsed.align !== undefined ||
    parsed.vAlign !== undefined ||
    /左上|右上|左下|右下|一番上|一番下|中央|真ん中|センター|上部|下部/.test(
      hintSource
    );

  const deltaOffsetX =
    (parsed.offsetX ?? 0) +
    (typeof args.offsetX === "number" ? args.offsetX : 0);
  const deltaOffsetY =
    (parsed.offsetY ?? 0) +
    (typeof args.offsetY === "number" ? args.offsetY : 0);

  const baseOffsetX = positionSpecified ? 0 : (last?.offsetX ?? 0);
  const baseOffsetY = positionSpecified ? 0 : (last?.offsetY ?? 0);

  const offsetX = baseOffsetX + deltaOffsetX;
  const offsetY = baseOffsetY + deltaOffsetY;

  console.log("📐 offset calculation:", {
    positionSpecified,
    baseOffsetX,
    baseOffsetY,
    parsedOffsetX: parsed.offsetX,
    parsedOffsetY: parsed.offsetY,
    argsOffsetX: args.offsetX,
    argsOffsetY: args.offsetY,
    deltaOffsetX,
    deltaOffsetY,
    finalOffsetX: offsetX,
    finalOffsetY: offsetY,
  });

  const bottomMargin = parsed.bottomMargin;

  lastTextLayoutByThread.set(chatThread.id, {
    align,
    vAlign,
    offsetX,
    offsetY,
    size,
    text,
    color,
    fontFamily,
    bold,
    italic,
  });

  console.log("💾 saved to Map:", {
    threadId: chatThread.id,
    saved: lastTextLayoutByThread.get(chatThread.id),
    mapSizeAfter: lastTextLayoutByThread.size,
  });

  const baseUrl =
    process.env.NEXTAUTH_URL ||
    (process.env.WEBSITE_HOSTNAME
      ? `https://${process.env.WEBSITE_HOSTNAME}`
      : "http://localhost:3000");

  const genImageBase = baseUrl.replace(/\/+$/, "");
  console.log("[gen-image] base URL for overlay:", genImageBase);
  console.log("[gen-image] resolved style params:", {
    align,
    vAlign,
    size,
    color,
    fontFamily,
    bold,
    italic,
    offsetX,
    offsetY,
    bottomMargin,
  });

  try {
    const resp = await fetch(`${genImageBase}/api/gen-image`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        imageUrl: baseImageUrl,
        text,
        align,
        vAlign,
        size,
        color,
        offsetX,
        offsetY,
        bottomMargin,
        autoDetectPlacard: false,
        fontFamily,
        bold,
        italic,
      }),
    });

    if (!resp.ok) {
      const t = await resp.text().catch(() => "");
      console.error("🔴 /api/gen-image failed in edit:", resp.status, t);
      return { error: `Text overlay failed: HTTP ${resp.status}` };
    }

    const result = await resp.json();
    const generatedPath = result?.imageUrl as string | undefined;

    if (!generatedPath) {
      console.error("🔴 gen-image edit returned no imageUrl");
      return { error: "gen-image edit returned no imageUrl" };
    }

    const fs = require("fs");
    const path = require("path");
    const finalImageName = `${uniqueId()}.png`;
    const finalImagePath = path.join(
      process.cwd(),
      "public",
      generatedPath.startsWith("/") ? generatedPath.slice(1) : generatedPath
    );
    const finalImageBuffer = fs.readFileSync(finalImagePath);

    await UploadImageToStore(chatThread.id, finalImageName, finalImageBuffer);

    const finalImageUrl = buildExternalImageUrl(chatThread.id, finalImageName);

    return { revised_prompt: text, url: finalImageUrl };
  } catch (err) {
    console.error("🔴 error in executeAddTextToExistingImage (simple):", err);
    return { error: "There was an error adding text to the existing image: " + err };
  }
}
