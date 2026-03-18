"use client";

import {
  ResetInputRows,
  onKeyDown,
  onKeyUp,
  useChatInputDynamicHeight,
} from "@/features/chat-page/chat-input/use-chat-input-dynamic-height";

import { AttachFile } from "@/features/ui/chat/chat-input-area/attach-file";
import {
  ChatInputActionArea,
  ChatInputForm,
  ChatInputPrimaryActionArea,
  ChatInputSecondaryActionArea,
} from "@/features/ui/chat/chat-input-area/chat-input-area";
import { ChatTextInput } from "@/features/ui/chat/chat-input-area/chat-text-input";
import { ImageInput } from "@/features/ui/chat/chat-input-area/image-input";
import { Microphone } from "@/features/ui/chat/chat-input-area/microphone";
import { StopChat } from "@/features/ui/chat/chat-input-area/stop-chat";
import { SubmitChat } from "@/features/ui/chat/chat-input-area/submit-chat";
import React, { useRef, useState } from "react";
import { useSession } from "next-auth/react";
import { chatStore, useChat } from "../chat-store";
import { fileStore, useFileStore } from "./file/file-store";
import { PromptSlider } from "./prompt/prompt-slider";
import {
  speechToTextStore,
  useSpeechToText,
} from "./speech/use-speech-to-text";
import {
  textToSpeechStore,
  useTextToSpeech,
} from "./speech/use-text-to-speech";

// ★ 思考モードトグル（3段階：標準→熟考→即答）
// import {
//   ModeCycleButton,
//   type ThinkingMode,
// } from "@/components/chatModeToggle";

type UploadScope = "common" | "personal";

export const ChatInput = () => {
  const { loading, input, chatThreadId } = useChat();
  const { uploadButtonLabel } = useFileStore();
  const { isPlaying } = useTextToSpeech();
  const { isMicrophoneReady } = useSpeechToText();
  const { rows } = useChatInputDynamicHeight();

  const submitButton = React.useRef<HTMLButtonElement>(null);
  const formRef = useRef<HTMLFormElement>(null);

  // ★ モードは ChatInput 側で保持（hidden input で route に渡す）
  const mode = "standard";

  // ★ 管理者判定
  const { data: session } = useSession();
  const adminEmails = (process.env.NEXT_PUBLIC_SL_ADMIN_EMAILS ?? "")
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);

  const me = (session?.user as any)?.email?.toLowerCase?.() ?? "";
  const isAdmin = adminEmails.includes(me);

  // 新仕様:
  // - 管理者は 共通 / 個人 を選択可能
  // - 一般メンバーは uploadScope を送らず、サーバ側で personal 固定 / 非SP部署はBlob判定
  const [uploadScope, setUploadScope] = useState<UploadScope>("personal");

  const submit = () => {
    if (formRef.current) {
      formRef.current.requestSubmit();
    }
  };

  return (
    <ChatInputForm
      ref={formRef}
      onSubmit={(e) => {
        e.preventDefault();
        chatStore.submitChat(e);
      }}
      status={uploadButtonLabel}
    >
      {/* ★ サーバに確実に届くよう hidden input を同送 */}
      <input type="hidden" name="thinkingMode" value={mode} />

      <ChatTextInput
        onBlur={(e) => {
          if (e.currentTarget.value.replace(/\s/g, "").length === 0) {
            ResetInputRows();
          }
        }}
        onKeyDown={(e) => {
          onKeyDown(e, submit);
        }}
        onKeyUp={(e) => {
          onKeyUp(e);
        }}
        value={input}
        rows={rows}
        onChange={(e) => {
          chatStore.updateInput(e.currentTarget.value);
        }}
      />

      <ChatInputActionArea>
        <ChatInputSecondaryActionArea>
          <AttachFile
            onClick={(formData) =>
              fileStore.onFileChange({
                formData,
                chatThreadId,
                uploadScope: isAdmin ? uploadScope : undefined,
              })
            }
          />

          <PromptSlider />

          {/* ★ 管理者のみ：アップロード先トグル（共通 / 個人） */}
          {isAdmin && (
            <div className="flex items-center gap-1 text-xs text-muted-foreground">
              <span>UP先:</span>

              {(
                [
                  { value: "common" as const, label: "共通" },
                  { value: "personal" as const, label: "個人" },
                ] as const
              ).map((item) => (
                <button
                  key={item.value}
                  type="button"
                  onClick={() => setUploadScope(item.value)}
                  className={`px-2 py-0.5 rounded border text-xs transition-colors ${
                    uploadScope === item.value
                      ? "bg-primary text-primary-foreground border-primary"
                      : "border-muted-foreground hover:bg-muted"
                  }`}
                  aria-pressed={uploadScope === item.value}
                >
                  {item.label}
                </button>
              ))}
            </div>
          )}

          {/* ★ 思考モードトグル（標準→熟考→即答→…） */}
          {/* <ModeCycleButton value={mode} onChange={setMode} /> */}
        </ChatInputSecondaryActionArea>

        <ChatInputPrimaryActionArea>
          <ImageInput />
          <Microphone
            startRecognition={() => speechToTextStore.startRecognition()}
            stopRecognition={() => speechToTextStore.stopRecognition()}
            isPlaying={isPlaying}
            stopPlaying={() => textToSpeechStore.stopPlaying()}
            isMicrophoneReady={isMicrophoneReady}
          />
          {loading === "loading" ? (
            <StopChat stop={() => chatStore.stopGeneratingMessages()} />
          ) : (
            <SubmitChat ref={submitButton} />
          )}
        </ChatInputPrimaryActionArea>
      </ChatInputActionArea>
    </ChatInputForm>
  );
};