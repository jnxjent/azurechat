// src/features/chat-page/chat-services/chat-api/chat-api.ts
"use server";
import "server-only";

// ★ SF拡張の Extension ID（環境変数化）
const SF_EXTENSION_ID = process.env.SF_EXTENSION_ID || "";

import { getCurrentUser } from "@/features/auth-page/helpers";
import { GenerateSasUrl } from "@/features/common/services/azure-storage";
import { isSalesforceAllowedEmail } from "@/features/common/services/salesforce-access";
import {
  resolveSalesforceRoute,
} from "@/features/common/services/salesforce-routing";
import { FindAllExtensionForCurrentUser } from "@/features/extensions-page/extension-services/extension-service";
import { CHAT_DEFAULT_SYSTEM_PROMPT } from "@/features/theme/theme-config";
import { ChatCompletionStreamingRunner } from "openai/resources/beta/chat/completions";
import { ChatApiRAG } from "../chat-api/chat-api-rag";
import { FindAllChatDocuments } from "../chat-document-service";
import {
  CreateChatMessage,
  FindTopChatMessagesForCurrentUser,
} from "../chat-message-service";
import {
  AddExtensionToChatThread,
  EnsureChatThreadOperation,
} from "../chat-thread-service";
import { LoadLatestImageAttachment, UploadImageToStore } from "../chat-image-service";
import { LoadPendingPptxEdit } from "../pptx-pending-edit-service";
import { resolvePptxPaletteInstruction } from "@/features/pptx/palette";
import { isExplicitPptOutputRequest, isNewPptCreationRequest, isAttachedPptConversionRequest, isExcelTableToSlideRequest, resolveOfficeChatRoute } from "@/features/pptx/output-intent";
import { ChatThreadModel, UserPrompt } from "../models";
import { mapOpenAIChatMessages } from "../utils";
import { GetDefaultExtensions } from "./chat-api-default-extensions";
import { GetDynamicExtensions } from "./chat-api-dynamic-extensions";
import { ChatApiExtensions } from "./chat-api-extension";
import { ChatApiMultimodal } from "./chat-api-multimodal";
import {
  isPptxAssetPlacementRequest,
  isSharePointImageRequest,
  resolveRequiredImageToolName,
} from "./image/image-intent";
import { OpenAIStream } from "./open-ai-stream";
import {
  isDeskNetsAgentEnabled,
} from "@/features/desknets-agent/desknets-agent-client";
import { createDeskNetsAgentTool } from "@/features/desknets-agent/desknets-agent-tool";
import {
  shouldRouteToDeskNetsAgent,
} from "@/features/desknets-agent/desknets-agent-intent";

type ChatTypes = "extensions" | "chat-with-file" | "multimodal";

type ThinkingModeUI = "standard" | "thinking" | "fast";
type ThinkingModeAPI = "normal" | "thinking" | "fast";

type UserPromptWithMode = UserPrompt & {
  thinkingMode?: ThinkingModeUI;
  apiThinkingMode?: ThinkingModeAPI;
};

function uiToApi(mode?: ThinkingModeUI): ThinkingModeAPI {
  if (mode === "thinking") return "thinking";
  if (mode === "fast") return "fast";
  return "normal";
}

function normalizeImageAttachments(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value.map((item) => String(item ?? "").trim()).filter(Boolean);
  }
  const single = String(value ?? "").trim();
  return single ? [single] : [];
}

/** ★最小ガード：直前 assistant.tool_calls に紐付かない tool を history から除外 */
function fixOrphanToolsInline(messages: any[]) {
  if (!Array.isArray(messages)) return messages;
  const out: any[] = [];
  let lastAssistantToolIds: Set<string> | null = null;

  for (const m of messages) {
    if (m?.role === "assistant") {
      lastAssistantToolIds = null;
      if (Array.isArray(m.tool_calls) && m.tool_calls.length > 0) {
        lastAssistantToolIds = new Set(
          m.tool_calls.map((tc: any) => tc?.id).filter(Boolean)
        );
      } else if (Array.isArray(m.tool_calls)) {
        // 空配列は削除（ノイズ防止）
        delete (m as any).tool_calls;
      }
      out.push(m);
      continue;
    }
    if (m?.role === "tool") {
      // 直前 assistant の tool_calls に一致しない tool は落とす
      if (
        lastAssistantToolIds &&
        m.tool_call_id &&
        lastAssistantToolIds.has(m.tool_call_id)
      ) {
        out.push(m);
      }
      continue;
    }
    // user / system が来たら直前の tool 関連はリセット
    lastAssistantToolIds = null;
    out.push(m);
  }
  return out;
}

export const ChatAPIEntry = async (props: UserPrompt, signal: AbortSignal) => {
  const currentChatThreadResponse = await EnsureChatThreadOperation(props.id);
  if (currentChatThreadResponse.status !== "OK") {
    return new Response("", { status: 401 });
  }
  const currentChatThread = currentChatThreadResponse.response;
  const user = await getCurrentUser();
  const hasSalesforceExtension =
    Boolean(SF_EXTENSION_ID) &&
    currentChatThread.extension.includes(SF_EXTENSION_ID);
  const salesforceAllowed = isSalesforceAllowedEmail(user.email);
  const salesforceRouting = resolveSalesforceRoute({
    message: props.message,
    isSalesforceAllowed: salesforceAllowed,
    hasSalesforceExtension,
    defaultToSalesforce: hasSalesforceExtension,
  });
  const executionChatThread: ChatThreadModel =
    salesforceRouting.route === "salesforce" || !hasSalesforceExtension
      ? currentChatThread
      : {
          ...currentChatThread,
          extension: currentChatThread.extension.filter(
            (extensionId) => extensionId !== SF_EXTENSION_ID
          ),
        };

  console.log("[SF Route]", {
    channel: "web",
    allowed: salesforceAllowed,
    intent: salesforceRouting.intent,
    route: salesforceRouting.route,
  });

  const p = props as UserPromptWithMode;
  const imageAttachmentUrls = normalizeImageAttachments(
    (p as any).multimodalImage
  );
  const referencesAnAttachedImage =
    /(?:添付|アップロード|ロゴ|画像ファイル|reference\s+image|attached\s+(?:logo|image))/i.test(
      props.message
    );
  const referencesSharePointImage = isSharePointImageRequest(props.message);
  const storedImageAttachment =
    imageAttachmentUrls.length === 0 && (referencesAnAttachedImage || isAttachedPptConversionRequest(props.message))
      ? await LoadLatestImageAttachment(currentChatThread.id)
      : null;
  const imageAttachmentCountForRouting =
    imageAttachmentUrls.length +
    (storedImageAttachment ? 1 : 0) +
    (referencesSharePointImage ? 1 : 0);
  const newPptRequest = isNewPptCreationRequest(props.message);
  const explicitPptRequest = isExplicitPptOutputRequest(props.message);
  const pptxAssetPlacementRequest = !newPptRequest && isPptxAssetPlacementRequest(props.message);
  const accentReplyText = props.message.trim();
  const mayBePendingPptxAccentReply =
    accentReplyText.length <= 100 &&
    !!resolvePptxPaletteInstruction(accentReplyText) &&
    !/(?:作成|変更|編集|追加|削除|挿入|配置|生成|直して|変えて|入れて|replace|change|edit|add|remove|insert|place)/i.test(
      accentReplyText
    );
  const pendingPptxEdit = mayBePendingPptxAccentReply
    ? await LoadPendingPptxEdit(currentChatThread.id).catch(() => null)
    : null;
  const requiredImageToolName = pendingPptxEdit
    ? "edit_pptx"
    : pptxAssetPlacementRequest
    ? "edit_pptx"
    : explicitPptRequest ? undefined : resolveRequiredImageToolName(
        props.message,
        imageAttachmentCountForRouting
      );
  const shouldUseImageEditTools = Boolean(requiredImageToolName);
  const resolvedMode: ThinkingModeAPI =
    p.apiThinkingMode ?? uiToApi(p.thinkingMode) ?? "normal";

  if (process.env.NODE_ENV !== "production") {
    console.log("📨 ChatAPIEntry received modes:", {
      apiThinkingMode: p.apiThinkingMode,
      uiThinkingMode: p.thinkingMode,
      resolvedMode,
      imageAttachmentCount: imageAttachmentUrls.length,
      imageAttachmentCountForRouting,
      storedImageAttachment: storedImageAttachment?.fileName ?? null,
      pendingPptxAccentReply: !!pendingPptxEdit,
      requiredImageToolName: requiredImageToolName ?? null,
    });
  }

  // 並列取得（extensions に mode を渡す）
  const [history, docs, extension] = await Promise.all([
    _getHistory(currentChatThread),
    _getDocuments(currentChatThread),
    _getExtensions({
      chatThread: executionChatThread,
      userMessage: props.message,
      imageAttachmentUrls,
      signal,
      mode: resolvedMode,
    }),
  ]);
  const latestDoc = [...docs].sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())[0];
  const excelTableSlideRequest = !!latestDoc && /\.(?:xlsx|xlsm|xls)$/i.test(latestDoc.name) &&
    isExcelTableToSlideRequest(props.message);
  const attachedPdfConversion = !excelTableSlideRequest && explicitPptRequest && !newPptRequest &&
    /添付|アップロード|PDF|この画像|このファイル/i.test(props.message) &&
    !/参考|追記|拡充|肉付け/i.test(props.message) &&
    docs.some(doc => /\.pdf$/i.test(doc.name));
  if (attachedPdfConversion) {
    const pdfDocs = docs.filter(doc => /\.pdf$/i.test(doc.name));
    const sources = await Promise.all(pdfDocs.map(async doc => {
      const sas = await GenerateSasUrl("dl-link", `${currentChatThread.id}/${doc.name}`);
      if (sas.status !== "OK") throw new Error("添付PDFのURLを取得できませんでした。");
      return `file_name: ${doc.name}\nfile_url: ${sas.response}`;
    }));
    history.push({ role: "system", content: `現在の会話に添付されたPDFです。利用者の添付画像という表現はこのPDFを指します。convert_doc_to_pptxへ渡してください。\n${sources.join("\n\n")}` });
    console.log("[PPT attachment] PDF conversion sources resolved", { count: sources.length });
  }
  let imageConversionCandidates = imageAttachmentUrls.length ? imageAttachmentUrls : storedImageAttachment
    ? [`data:${storedImageAttachment.contentType};base64,${storedImageAttachment.buffer.toString("base64")}`] : [];
  // UploadDocument also retains source image files. Use them even when the short-lived
  // image-reference metadata has expired or was consumed by an earlier image tool.
  if (!imageConversionCandidates.length && !attachedPdfConversion && isAttachedPptConversionRequest(props.message)) {
    imageConversionCandidates = await Promise.all(docs.filter(doc => /\.(?:png|jpe?g|webp)$/i.test(doc.name)).map(async doc => {
      const sas = await GenerateSasUrl("dl-link", `${currentChatThread.id}/${doc.name}`);
      if (sas.status !== "OK") throw new Error("添付画像のURLを取得できませんでした。");
      return sas.response;
    }));
  }
  const attachedImageConversion = !excelTableSlideRequest && !attachedPdfConversion && !pptxAssetPlacementRequest &&
    isAttachedPptConversionRequest(props.message) && imageConversionCandidates.length > 0;
  if (attachedImageConversion) {
    const imageConversionSources = await Promise.all(imageConversionCandidates.map(async (source, index) => {
      if (!source.startsWith("data:")) return source;
      const match = source.match(/^data:image\/(?:png|jpeg|jpg|webp);base64,([A-Za-z0-9+/=\r\n]+)$/);
      if (!match) throw new Error("添付画像の形式が不正です。PNG・JPEG・WEBPを使用してください。");
      const name = `ppt-source-${Date.now()}-${index}.png`;
      const uploaded = await UploadImageToStore(currentChatThread.id, name, Buffer.from(match[1], "base64"));
      if (uploaded.status !== "OK") throw new Error("PowerPoint変換用の画像を保存できませんでした。");
      const sas = await GenerateSasUrl("images", `${currentChatThread.id}/${name}`);
      if (sas.status !== "OK") throw new Error("添付画像のURLを取得できませんでした。");
      return sas.response;
    }));
    history.push({ role: "system", content: `添付画像そのものをPowerPointへ変換する依頼です。convert_doc_to_pptxを使い、mode=faithfulで画像を残し文字を編集可能にしてください。\n${imageConversionSources.map(url => `file_url: ${url}`).join("\n\n")}` });
  }


  // 2ターン目以降の「候補2に変更して」などでも、同じAgentセッションを継続する。
  if (
    isDeskNetsAgentEnabled() &&
    shouldRouteToDeskNetsAgent(props.message, history)
  ) {
    extension.push(
      createDeskNetsAgentTool(currentChatThread.id, props.message)
    );
    console.log("[DeskNetsAgent] Native tool enabled", {
      chatThreadId: currentChatThread.id,
      chatTypeHint: docs.length > 0 ? "rag" : "extensions",
    });
  }

  currentChatThread.personaMessage = `${CHAT_DEFAULT_SYSTEM_PROMPT} \n\n ${currentChatThread.personaMessage}`;
  executionChatThread.personaMessage = currentChatThread.personaMessage;

  const chatType: ChatTypes = resolveOfficeChatRoute({
    explicitPptRequest: explicitPptRequest || excelTableSlideRequest,
    requiredTools: salesforceRouting.route !== "normal" || shouldUseImageEditTools,
    imageCount: imageAttachmentUrls.length,
    documentCount: docs.length,
  });
  console.log("[Office route]", { chatType, explicitPptRequest, newPptRequest, imageCount: imageAttachmentUrls.length, documentCount: docs.length });

  await CreateChatMessage({
    name: user.name,
    content: props.message,
    role: "user",
    chatThreadId: currentChatThread.id,
    multiModalImage: (p as any).multimodalImage,
  });

  let runner: ChatCompletionStreamingRunner;
  switch (chatType) {
    case "chat-with-file":
      runner = await ChatApiRAG({
        chatThread: currentChatThread,
        userMessage: props.message,
        history,
        signal,
      });
      break;
    case "multimodal":
      runner = ChatApiMultimodal({
        chatThread: currentChatThread,
        userMessage: props.message,
        file: (p as any).multimodalImage,
        signal,
      });
      break;
    case "extensions":
    default:
      runner = await ChatApiExtensions({
        chatThread: executionChatThread,
        userMessage: props.message,
        history,
        extensions: extension,
        requiredToolName: (excelTableSlideRequest ? "excel_table_to_slide" : attachedPdfConversion || attachedImageConversion ? "convert_doc_to_pptx" : newPptRequest ? "create_pptx" : requiredImageToolName),
        loginEmail: user.email,
        salesforceRouting,
        signal,
        mode: resolvedMode,
      });
      break;
  }

  const readableStream = OpenAIStream({ runner, chatThread: currentChatThread });
  return new Response(readableStream, {
    headers: { "Cache-Control": "no-cache", Connection: "keep-alive" },
  });
};

const _getHistory = async (chatThread: ChatThreadModel) => {
  const historyResponse =
    await FindTopChatMessagesForCurrentUser(chatThread.id);
  if (historyResponse.status === "OK") {
    const historyResults = historyResponse.response;
    // DB → OpenAI 形式へ
    const mapped = mapOpenAIChatMessages(historyResults).reverse();
    // ★ここで一発サニタイズ（孤立 tool を除去）
    return fixOrphanToolsInline(mapped);
  }
  console.error("🔴 Error on getting history:", historyResponse.errors);
  return [];
};

const _getDocuments = async (chatThread: ChatThreadModel) => {
  const docsResponse = await FindAllChatDocuments(chatThread.id);
  if (docsResponse.status === "OK") {
    return docsResponse.response;
  }
  console.error("🔴 Error on AI search:", docsResponse.errors);
  return [];
};

const normalizeExtensionName = (name: string): string =>
  name.toLowerCase().replace(/[\s_-]+/g, "");

const isBraveSearchExtensionName = (name: string): boolean => {
  const normalized = normalizeExtensionName(name);
  return (
    normalized === "bravesearch" ||
    (normalized.includes("brave") && normalized.includes("search"))
  );
};

const knownBraveSearchExtensionIds = new Set<string>();

/**
 * Enable the existing BraveSearch extension in ordinary Chat without making
 * each user turn it on for every thread. Set
 * BRAVE_SEARCH_DEFAULT_ENABLED=false to opt out.
 */
const ensureDefaultBraveSearchExtension = async (
  chatThread: ChatThreadModel
): Promise<void> => {
  if (
    process.env.BRAVE_SEARCH_DEFAULT_ENABLED?.trim().toLowerCase() === "false"
  ) {
    return;
  }
  if (chatThread.extension.includes(SF_EXTENSION_ID)) {
    return;
  }
  if (
    chatThread.extension.some((id) => knownBraveSearchExtensionIds.has(id))
  ) {
    return;
  }

  const available = await FindAllExtensionForCurrentUser();
  if (available.status !== "OK") {
    console.warn(
      "[BraveSearch] Could not load extensions for default enablement."
    );
    return;
  }

  const braveExtension = available.response.find((extension) =>
    isBraveSearchExtensionName(extension.name)
  );
  if (!braveExtension) {
    return;
  }
  knownBraveSearchExtensionIds.add(braveExtension.id);
  if (chatThread.extension.includes(braveExtension.id)) return;

  // Use it immediately even if the database update temporarily fails.
  chatThread.extension.push(braveExtension.id);
  const persisted = await AddExtensionToChatThread({
    chatThreadId: chatThread.id,
    extensionId: braveExtension.id,
  });
  if (persisted.status === "OK") {
    console.log("[BraveSearch] Enabled by default for Chat thread", {
      chatThreadId: chatThread.id,
      extensionId: braveExtension.id,
    });
  } else {
    console.warn("[BraveSearch] Default extension persistence failed", {
      chatThreadId: chatThread.id,
      extensionId: braveExtension.id,
    });
  }
};

const _getExtensions = async (props: {
  chatThread: ChatThreadModel;
  userMessage: string;
  imageAttachmentUrls?: string[];
  signal: AbortSignal;
  mode: ThinkingModeAPI;
}) => {
  const extension: Array<any> = [];

  await ensureDefaultBraveSearchExtension(props.chatThread);

  // ★ このスレッドが SF 拡張を持っているか？
  const hasSfExtension =
    Array.isArray(props.chatThread.extension) &&
    props.chatThread.extension.includes(SF_EXTENSION_ID);

  // ★ SF スレッドのときは、汎用のデフォルト拡張（画像ツールなど）をスキップして高速化
  if (!hasSfExtension) {
    const response = await GetDefaultExtensions({
      chatThread: props.chatThread,
      userMessage: props.userMessage,
      imageAttachmentUrls: props.imageAttachmentUrls,
      signal: props.signal,
      mode: props.mode, // ← ここが“断絶”をつなぐ肝
    });
    if (response.status === "OK" && response.response.length > 0) {
      extension.push(...response.response);
    }
  } else if (process.env.NODE_ENV !== "production") {
    console.log(
      "[SF] SF_EXTENSION_ID detected. Skipping default (image) extensions for speed."
    );
  }

  const dynamicExtensionsResponse = await GetDynamicExtensions({
    extensionIds: props.chatThread.extension,
  });
  if (
    dynamicExtensionsResponse.status === "OK" &&
    dynamicExtensionsResponse.response.length > 0
  ) {
    extension.push(...dynamicExtensionsResponse.response);
  }

  return extension;
};
