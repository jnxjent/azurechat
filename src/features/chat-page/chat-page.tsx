"use client";
import { ChatInput } from "@/features/chat-page/chat-input/chat-input";
import { chatStore, useChat } from "@/features/chat-page/chat-store";
import { ChatLoading } from "@/features/ui/chat/chat-message-area/chat-loading";
import { ChatMessageArea } from "@/features/ui/chat/chat-message-area/chat-message-area";
import ChatMessageContainer from "@/features/ui/chat/chat-message-area/chat-message-container";
import ChatMessageContentArea from "@/features/ui/chat/chat-message-area/chat-message-content";
import { useChatScrollAnchor } from "@/features/ui/chat/chat-message-area/use-chat-scroll-anchor";
import { Download } from "lucide-react";
import { useSession } from "next-auth/react";
import { FC, useEffect, useRef } from "react";
import { ExtensionModel } from "../extensions-page/extension-services/models";
import { ChatHeader } from "./chat-header/chat-header";
import {
  ChatDocumentModel,
  ChatMessageModel,
  ChatThreadModel,
} from "./chat-services/models";
import MessageContent from "./message-content";

interface ChatPageProps {
  messages: Array<ChatMessageModel>;
  chatThread: ChatThreadModel;
  chatDocuments: Array<ChatDocumentModel>;
  extensions: Array<ExtensionModel>;
  isAdmin: boolean; // ← 追加
}

// ChatMessageArea が受け取れるロール（UI 用）
type ChatUiRole = "user" | "system" | "assistant" | "tool";

function toUiRole(role: ChatMessageModel["role"]): ChatUiRole {
  if (role === "function") {
    return "assistant";
  }
  return role as ChatUiRole;
}

export const ChatPage: FC<ChatPageProps> = (props) => {
  const { data: session } = useSession();

  // Server Component 側で確定済みの isAdmin を使う
  const isAdmin = props.isAdmin;

  useEffect(() => {
    chatStore.initChatSession({
      chatThread: props.chatThread,
      messages: props.messages,
      userName: session?.user?.name!,
    });
  }, [props.messages, session?.user?.name, props.chatThread]);

  const { messages, loading } = useChat();

  const current = useRef<HTMLDivElement>(null);

  useChatScrollAnchor({ ref: current });

  return (
    <main className="flex flex-1 relative flex-col">
      <ChatHeader
        chatThread={props.chatThread}
        chatDocuments={props.chatDocuments}
        extensions={props.extensions}
        isAdmin={isAdmin}
      />
      <ChatMessageContainer ref={current}>
        <ChatMessageContentArea>
          {messages.map((message, index) => {
            // アシスタントメッセージの直前にある tool メッセージの downloadUrl を取得
            // → 「Page 8をカード型デザインにしました。」の下にもダウンロードボタンを表示
            let assistantDownloadUrl: string | null = null;
            let assistantDownloadName: string | null = null;
            if (message.role === "assistant" && index > 0) {
              const prev = messages[index - 1];
              if (prev.role === "tool" || prev.role === "function") {
                try {
                  const obj = JSON.parse(prev.content);
                  if (
                    typeof obj?.downloadUrl === "string" &&
                    /^https?:\/\/[^/]+\.blob\.core\.windows\.net\//i.test(obj.downloadUrl)
                  ) {
                    assistantDownloadUrl = obj.downloadUrl;
                    assistantDownloadName = obj.displayName ?? obj.fileName ?? null;
                  }
                } catch {}
              }
            }

            return (
              <ChatMessageArea
                key={message.id}
                profileName={message.name}
                role={toUiRole(message.role)}
                onCopy={() => {
                  navigator.clipboard.writeText(message.content);
                }}
                profilePicture={
                  message.role === "assistant"
                    ? "/ai-icon.png"
                    : session?.user?.image
                }
              >
                <MessageContent message={message} />
                {assistantDownloadUrl && (
                  <a
                    href={assistantDownloadUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="mt-2 inline-flex items-center gap-1.5 rounded-md border px-3 py-1.5 text-sm font-medium hover:bg-accent transition-colors"
                  >
                    <Download size={14} strokeWidth={2} />
                    {assistantDownloadName ?? "ダウンロード"}
                  </a>
                )}
              </ChatMessageArea>
            );
          })}
          {loading === "loading" && <ChatLoading />}
        </ChatMessageContentArea>
      </ChatMessageContainer>
      <ChatInput isAdmin={isAdmin} />
    </main>
  );
};