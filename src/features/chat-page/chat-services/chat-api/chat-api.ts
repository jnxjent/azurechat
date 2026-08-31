// src/features/chat-page/chat-services/chat-api/chat-api.ts
"use server";
import "server-only";

// ★ SF拡張の Extension ID（環境変数化）
const SF_EXTENSION_ID = process.env.SF_EXTENSION_ID || "";

import { getCurrentUser } from "@/features/auth-page/helpers";
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
import { LoadLatestImageAttachment } from "../chat-image-service";
import { LoadPendingPptxEdit } from "../pptx-pending-edit-service";
import { resolvePptxPaletteInstruction } from "@/features/pptx/palette";
import { ChatThreadModel, UserPrompt } from "../models";
import { mapOpenAIChatMessages } from "../utils";
import { GetDefaultExtensions } from "./chat-api-default-extensions";
import { GetDynamicExtensions } from "./chat-api-dynamic-extensions";
import { ChatApiExtensions } from "./chat-api-extension";
import { ChatApiMultimodal } from "./chat-api-multimodal";
import {
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

/**
 * A logo/image attached for an existing deck is a PowerPoint asset, not a
 * request to compose a new raster image with GPT Image.
 */
function isPptxAssetPlacementRequest(message: string): boolean {
  const value = String(message ?? "");
  const referencesExistingDeck =
    /(?:添付|上記|この|今の|先ほど|さきほど|出力した|生成した|作成した|既存の).{0,16}(?:PPTX?|PowerPoint|スライド|資料)/i.test(
      value
    ) ||
    /(?:PPTX?|PowerPoint|スライド|資料).{0,16}(?:編集|修正|変更|追加)/i.test(
      value
    );
  const asksForAssetPlacement =
    /(?:ロゴ|logo|添付画像|画像).{0,20}(?:入れ|挿入|配置|載せ|追加|右肩|表紙|フロント)/i.test(
      value
    ) ||
    /(?:入れ|挿入|配置|載せ|追加).{0,20}(?:ロゴ|logo|添付画像)/i.test(
      value
    );
  return referencesExistingDeck && asksForAssetPlacement;
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
    imageAttachmentUrls.length === 0 && referencesAnAttachedImage
      ? await LoadLatestImageAttachment(currentChatThread.id)
      : null;
  const imageAttachmentCountForRouting =
    imageAttachmentUrls.length +
    (storedImageAttachment ? 1 : 0) +
    (referencesSharePointImage ? 1 : 0);
  const pptxAssetPlacementRequest = isPptxAssetPlacementRequest(props.message);
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
    : resolveRequiredImageToolName(
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
  const [user, history, docs, extension] = await Promise.all([
    getCurrentUser(),
    _getHistory(currentChatThread),
    _getDocuments(currentChatThread),
    _getExtensions({
      chatThread: currentChatThread,
      userMessage: props.message,
      imageAttachmentUrls,
      signal,
      mode: resolvedMode,
    }),
  ]);

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

  let chatType: ChatTypes = "extensions";
  if (shouldUseImageEditTools) {
    chatType = "extensions";
  } else if (imageAttachmentUrls.length > 0) {
    chatType = "multimodal";
  } else if (docs.length > 0) {
    chatType = "chat-with-file";
  } else if (extension.length > 0) {
    chatType = "extensions";
  }

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
        chatThread: currentChatThread,
        userMessage: props.message,
        history,
        extensions: extension,
        requiredToolName: requiredImageToolName,
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
