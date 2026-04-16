// src/features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts
"use server";
import "server-only";

import { GenerateSasUrl, UploadBlob } from "@/features/common/services/azure-storage";
import { OpenAIDALLEInstance } from "@/features/common/services/openai";
import { ServerActionResponse } from "@/features/common/server-action-response";
import { uniqueId } from "@/features/common/util";
import { GetImageUrl, UploadImageToStore } from "../chat-image-service";
import { FindTopChatMessagesForCurrentUser } from "../chat-message-service";
import { ChatThreadModel } from "../models";
import { BlobServiceClient } from "@azure/storage-blob";
import { analyzeDocVision } from "@/app/api/analyze-doc-vision/route";
import { SimpleSearch } from "@/features/chat-page/chat-services/azure-ai-search/azure-ai-search";
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
        await executeCreatePptx(args, props.chatThread),
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
            description: "スライドのリスト",
            items: {
              type: "object",
              properties: {
                title: {
                  type: "string",
                  description: "スライドのタイトル",
                },
                bullets: {
                  type: "array",
                  items: { type: "string" },
                  description: "箇条書きの内容リスト",
                },
              },
              required: ["title", "bullets"],
            },
          },
          fontFace: {
            type: "string",
            description: "PowerPointで使うフォント名。例: 'Meiryo', 'Yu Gothic', 'Yu Mincho'",
          },
          designInstruction: {
            type: "string",
            description: "デザイン・色調の指示。例: '赤いトーンで力強く', '青を基調にしたシンプルなデザイン', 'ポップで明るい印象'",
          },
        },
        required: ["title", "slides"],
      },
      description:
        "ユーザーがテーマや内容を指定してPowerPoint（PPTX）を新規作成するツール。\n" +
        "テキストベースでスライド構成を作る場合に使用する。\n" +
        "ファイル（PDF・画像）をPPTに変換する場合は、代わりに convert_doc_to_pptx ツールを使うこと。\n" +
        "【重要】会話中にすでにPPTXが生成・編集された実績がある場合、色・デザイン・テキスト変更は edit_pptx を使うこと。このツールは完全新規作成専用。\n" +
        "ユーザーが色やデザインを指定した場合は designInstruction に渡すこと（例：「赤いトーン」→ designInstruction='赤いトーンで力強く'）。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
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
              "ユーザーの編集指示。例: '色を青に変えて', 'フォントを游ゴシックに', '全体のトーンを力強く', '3枚目のタイトルをXXXに変えて'",
          },
        },
        required: ["instruction"],
      },
      description:
        "このスレッドで生成・編集した既存PPTXを自然言語の指示に従って改良するツール。\n" +
        "【重要】fileUrlは省略可能。省略するとスレッド内の最新PPTXを自動で参照する。\n" +
        "使用タイミング：\n" +
        "【最優先】会話中にPPTXが生成・編集された実績がある場合は、このツールを使うこと。\n" +
        "- ユーザーが「色を変えて」「緑にして」「赤くして」「青にして」などの色変更を求める場合\n" +
        "- ユーザーが「フォントを変えて」「もっとポップに」などデザイン変更を求める場合\n" +
        "- ユーザーが「〜に変えて」「〜を修正して」「〜を追加して」などテキスト編集を求める場合\n" +
        "fileUrlは省略可（スレッド内の直近PPTXを自動取得）。ユーザーがファイルをアップしていなくても呼び出せる。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
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

  // ★ アップロードされた Excel ファイルを指示に従って編集するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) => await executeEditExcel(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileUrl: {
            type: "string",
            description:
              "編集対象のExcelファイルのURL。このスレッドでアップロードされたxlsx/xls/xlsmのURLを指定する。",
          },
          instruction: {
            type: "string",
            description:
              "ユーザーの編集指示。例: 'A1セルを「売上合計」に変えて', '1行目を太字・背景色を青に', '「旧社名」を「新社名」に置換して'",
          },
        },
        required: ["fileUrl", "instruction"],
      },
      description:
        "このスレッドでアップロードされたExcelファイル（xlsx/xls/xlsm）を自然言語の指示に従って編集するツール。\n" +
        "使用タイミング：ユーザーがExcelファイルをアップロードし、セル値の変更・テキスト置換・書式変更（太字・色）を求める場合。\n" +
        "fileUrl は会話コンテキストのアップロードURLを使用する。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "edit_excel",
    },
  });

  // ★ アップロードされた Word ファイルを指示に従って編集するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) => await executeEditWord(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileUrl: {
            type: "string",
            description:
              "編集対象のWordファイルのURL。このスレッドでアップロードされた.docxのURLを指定する。",
          },
          instruction: {
            type: "string",
            description:
              "ユーザーの編集指示。例: '「旧社名」を「新社名」に置換して', 'タイトルを太字・赤色にして', '第1章の見出しを16ptにして'",
          },
        },
        required: ["fileUrl", "instruction"],
      },
      description:
        "このスレッドでアップロードされたWordファイル（.docx）を自然言語の指示に従って編集するツール。\n" +
        "使用タイミング：ユーザーがWordファイルをアップロードし、テキスト置換・書式変更（太字・色・フォントサイズ）を求める場合。\n" +
        "fileUrl は会話コンテキストのアップロードURLを使用する。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "edit_word",
    },
  });

  // ★ アップロードされた PDF ファイルを Excel に変換するツール
  defaultExtensions.push({
    type: "function",
    function: {
      function: async (args: any) => await executeConvertPdfToExcel(args, props.chatThread),
      parse: (input: string) => JSON.parse(input),
      parameters: {
        type: "object",
        properties: {
          fileUrl: {
            type: "string",
            description:
              "変換対象のPDFファイルのURL。このスレッドでアップロードされた.pdfのURLを指定する。",
          },
        },
        required: ["fileUrl"],
      },
      description:
        "このスレッドでアップロードされたPDFファイルをExcel（.xlsx）に変換するツール。\n" +
        "使用タイミング：ユーザーがPDFをExcelに変換したいと言った場合。\n" +
        "PDFのテーブルはシートに、テーブルがない場合はテキストを「Text」シートに出力する。\n" +
        "ツールが返した downloadUrl を必ずMarkdownリンク形式 [ファイル名](downloadUrl) でユーザーに提示すること。",
      name: "convert_pdf_to_excel",
    },
  });

  return { status: "OK", response: defaultExtensions };
};

// ---------------- PowerPoint 生成 ----------------
async function executeCreatePptx(
  args: {
    title: string;
    slides: Array<{ title: string; bullets: string[] }>;
    fontFace?: string;
    designInstruction?: string;
  },
  chatThread: ChatThreadModel
) {
  const { title, slides, fontFace, designInstruction } = args ?? {};

  if (!title || !slides?.length) {
    return { error: "title and slides are required." };
  }

  // Each PPT creation is independent — do not accumulate style from thread history.
  const explicitInstruction = designInstruction?.trim() || undefined;
  const deckPreferences: DeckPreferences = explicitInstruction
    ? { designInstruction: explicitInstruction }
    : {};

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
        slides,
        threadId: chatThread.id,
        fontFace,
        designInstruction: explicitInstruction,
        deckPreferences,
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
      totalPages,
      message: `${totalPages}ページをVision APIで解析し、PowerPointファイルを生成しました。`,
    };
  } catch (e: any) {
    console.error("[convert_doc_to_pptx] error:", e);
    return { error: "変換中にエラーが発生しました: " + String(e?.message ?? e) };
  }
}

// ---------------- 既存 PPTX 改良 ----------------
async function executeEditPptx(
  args: { fileUrl?: string; instruction: string },
  chatThread: ChatThreadModel
) {
  let { fileUrl, instruction } = args ?? {};

  if (!instruction?.trim()) {
    return { error: "instructionは必須です。編集内容を指定してください。" };
  }

  if (!fileUrl?.trim()) {
    fileUrl = (await resolveLatestPptxUrlFromThread(chatThread.id)) ?? "";
  }

  if (!fileUrl?.trim()) {
    return {
      error:
        "編集対象のPPTXが見つかりませんでした。このスレッドでPPTXを生成するか、PPTのURLを指定してください。",
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

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
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
        "編集対象のExcelファイルが見つかりませんでした。このスレッドでExcelファイルをアップロードしてください。",
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
      console.error("[edit_excel] edit-pptx route failed:", res.status, t);
      return { error: `Excel編集に失敗しました: HTTP ${res.status}` };
    }

    const result = await res.json();
    if (!result?.downloadUrl) {
      return { error: "ダウンロードURLが取得できませんでした。" };
    }

    return {
      downloadUrl: result.downloadUrl,
      fileName: result.fileName,
      changedSheets: result.changedSheets,
      totalSheets: result.totalSheets,
      message: `${result.changedSheets}シートを編集しました（全${result.totalSheets}シート）。`,
    };
  } catch (e: any) {
    console.error("[edit_excel] error:", e);
    return { error: "Excel編集中にエラーが発生しました: " + String(e?.message ?? e) };
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
  const { fileUrl } = args ?? {};

  if (!fileUrl?.trim()) {
    return {
      error:
        "変換対象のPDFファイルが見つかりませんでした。このスレッドでPDFファイルをアップロードしてください。",
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
      message: `${result.pages}ページのPDFを変換しました（${tableInfo}）。`,
    };
  } catch (e: any) {
    console.error("[convert_pdf_to_excel] error:", e);
    return { error: "PDF→Excel変換中にエラーが発生しました: " + String(e?.message ?? e) };
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

  if (candidates.length > 1) {
    const list = candidates.map((c, i) => `${i + 1}. ${c.fileName}`).join("\n");
    return {
      multipleFiles: true,
      message: `「${fileQuery}」で複数のファイルが見つかりました。どれを編集しますか？\n\n${list}\n\nファイル名を指定して再度お試しください。`,
    };
  }

  const { fileName, sourceUrl, effectiveFileUrl } = candidates[0];
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
