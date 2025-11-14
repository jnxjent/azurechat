// src/features/chat-page/chat-services/chat-api/chat-api.ts
"use server";
import "server-only";

import { getCurrentUser } from "@/features/auth-page/helpers";
import { CHAT_DEFAULT_SYSTEM_PROMPT } from "@/features/theme/theme-config";
import { ChatCompletionStreamingRunner } from "openai/resources/beta/chat/completions";
import { ChatApiRAG } from "../chat-api/chat-api-rag";
import { FindAllChatDocuments } from "../chat-document-service";
import {
  CreateChatMessage,
  FindTopChatMessagesForCurrentUser,
} from "../chat-message-service";
import { EnsureChatThreadOperation } from "../chat-thread-service";
import { ChatThreadModel, UserPrompt } from "../models";
import { mapOpenAIChatMessages } from "../utils";
import { GetDefaultExtensions } from "./chat-api-default-extensions";
import { GetDynamicExtensions } from "./chat-api-dynamic-extensions";
import { ChatApiExtensions } from "./chat-api-extension";
import { ChatApiMultimodal } from "./chat-api-multimodal";
import { OpenAIStream } from "./open-ai-stream";

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
      if (lastAssistantToolIds && m.tool_call_id && lastAssistantToolIds.has(m.tool_call_id)) {
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
  const resolvedMode: ThinkingModeAPI = p.apiThinkingMode ?? uiToApi(p.thinkingMode) ?? "normal";

  if (process.env.NODE_ENV !== "production") {
    console.log("📨 ChatAPIEntry received modes:", {
      apiThinkingMode: p.apiThinkingMode,
      uiThinkingMode: p.thinkingMode,
      resolvedMode,
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
      signal,
      mode: resolvedMode,
    }),
  ]);

  currentChatThread.personaMessage = `${CHAT_DEFAULT_SYSTEM_PROMPT} \n\n ${currentChatThread.personaMessage}`;

  let chatType: ChatTypes = "extensions";
  if ((p as any).multimodalImage && (p as any).multimodalImage.length > 0) {
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
        signal,
      });
      break;
  }

  const readableStream = OpenAIStream({ runner, chatThread: currentChatThread });
  return new Response(readableStream, {
    headers: { "Cache-Control": "no-cache", Connection: "keep-alive" },
  });
};

const _getHistory = async (chatThread: ChatThreadModel) => {
  const historyResponse = await FindTopChatMessagesForCurrentUser(chatThread.id);
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

const _getExtensions = async (props: {
  chatThread: ChatThreadModel;
  userMessage: string;
  signal: AbortSignal;
  mode: ThinkingModeAPI;
}) => {
  const extension: Array<any> = [];

  const response = await GetDefaultExtensions({
    chatThread: props.chatThread,
    userMessage: props.userMessage,
    signal: props.signal,
    mode: props.mode, // ← ここが“断絶”をつなぐ肝
  });
  if (response.status === "OK" && response.response.length > 0) {
    extension.push(...response.response);
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
