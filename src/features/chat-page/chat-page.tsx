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
  isAdmin: boolean;
}

type ChatUiRole = "user" | "system" | "assistant" | "tool";

function toUiRole(role: ChatMessageModel["role"]): ChatUiRole {
  if (role === "function") return "assistant";
  return role as ChatUiRole;
}

const BLOB_URL_RE = /^https?:\/\/[^/]+\.blob\.core\.windows\.net\//i;

export const ChatPage: FC<ChatPageProps> = (props) => {
  const { data: session } = useSession();
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
            let assistantDownloadUrl: string | null = null;
            let assistantDownloadName: string | null = null;
            if (message.role === "assistant" && index > 0) {
              for (let i = index - 1; i >= 0; i--) {
                const prev = messages[i];
                if (prev.role !== "tool" && prev.role !== "function") break;
                try {
                  const obj = JSON.parse(prev.content);
                  if (
                    typeof obj?.downloadUrl === "string" &&
                    BLOB_URL_RE.test(obj.downloadUrl)
                  ) {
                    assistantDownloadUrl = obj.downloadUrl;
                    assistantDownloadName = obj.displayName ?? obj.fileName ?? null;
                    break;
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
                  <div className="not-prose mt-2">
                    <a
                      href={assistantDownloadUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1.5 rounded-md border px-3 py-1.5 text-sm font-medium hover:bg-accent transition-colors"
                    >
                      <Download size={14} strokeWidth={2} />
                      {assistantDownloadName ?? "ダウンロード"}
                    </a>
                  </div>
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
