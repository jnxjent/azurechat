import { AI_NAME } from "@/features/theme/theme-config";
import { ChatCompletionStreamingRunner } from "openai/resources/beta/chat/completions";
import { CreateChatMessage } from "../chat-message-service";
import {
  AzureChatCompletion,
  AzureChatCompletionAbort,
  ChatThreadModel,
} from "../models";

export const OpenAIStream = (props: {
  runner: ChatCompletionStreamingRunner;
  chatThread: ChatThreadModel;
}) => {
  const encoder = new TextEncoder();

  const { runner, chatThread } = props;

  const readableStream = new ReadableStream({
    async start(controller) {
      const streamResponse = (event: string, value: string) => {
        controller.enqueue(encoder.encode(`event: ${event} \n`));
        controller.enqueue(encoder.encode(`data: ${value} \n\n`));
      };

      let lastMessage = "";

      runner
        .on("content", (content) => {
          const completion = runner.currentChatCompletionSnapshot;
          if (completion) {
            const response: AzureChatCompletion = {
              type: "content",
              response: completion,
            };
            lastMessage = completion.choices[0].message.content ?? "";
            streamResponse(response.type, JSON.stringify(response));
          }
        })

        // ✅ GPT-5 ツール呼び出し（新）
        .on("toolCall", async (toolCall) => {
          // UI/履歴用に保存：role は "tool"
          await CreateChatMessage({
            // name はUI表示用（DBスキーマに合わせて任意）
            name: toolCall.function?.name ?? "tool",
            content: toolCall.function?.arguments ?? "",
            role: "tool",
            chatThreadId: chatThread.id,
            // 可能なら追加（CreateChatMessage/DBに項目を用意）
            tool_call_id: toolCall.id,
          });

          const response: AzureChatCompletion = {
            type: "toolCall",
            response: toolCall,
          };
          streamResponse(response.type, JSON.stringify(response));
        })

        .on("toolCallResult", async (toolCallResult) => {
          // 実行結果を保存：role は "tool"、content は文字列
          await CreateChatMessage({
            name: "tool",
            content:
              typeof toolCallResult === "string"
                ? toolCallResult
                : JSON.stringify(toolCallResult),
            role: "tool",
            chatThreadId: chatThread.id,
            tool_call_id: (toolCallResult as any)?.id, // あれば保存
          });

          const response: AzureChatCompletion = {
            type: "toolCallResult",
            response: toolCallResult,
          };
          streamResponse(response.type, JSON.stringify(response));
        })

        // （オプション）後方互換：古い functionCall イベントが来た場合でも "tool" として保存
        .on("functionCall", async (functionCall) => {
          await CreateChatMessage({
            name: functionCall.name ?? "tool",
            content: functionCall.arguments ?? "",
            role: "tool",                    // ← 重要: function は使わない
            chatThreadId: chatThread.id,
            tool_call_id: (functionCall as any)?.id,
          });

          const response: AzureChatCompletion = {
            type: "functionCall",            // イベント名はそのまま流す（UI依存なら toolCall に寄せてもOK）
            response: functionCall,
          };
          streamResponse(response.type, JSON.stringify(response));
        })

        .on("functionCallResult", async (functionCallResult) => {
          await CreateChatMessage({
            name: "tool",
            content:
              typeof functionCallResult === "string"
                ? functionCallResult
                : JSON.stringify(functionCallResult),
            role: "tool",                    // ← 重要
            chatThreadId: chatThread.id,
            tool_call_id: (functionCallResult as any)?.id,
          });

          const response: AzureChatCompletion = {
            type: "functionCallResult",
            response: functionCallResult,
          };
          streamResponse(response.type, JSON.stringify(response));
        })

        .on("abort", (error) => {
          const response: AzureChatCompletionAbort = {
            type: "abort",
            response: "Chat aborted",
          };
          streamResponse(response.type, JSON.stringify(response));
          controller.close();
        })
        .on("error", async (error) => {
          console.log("🔴 error", error);
          const response: AzureChatCompletion = {
            type: "error",
            response: error.message,
          };

          // if there is an error still save the last message even though it is not complete
          await CreateChatMessage({
            name: AI_NAME,
            content: lastMessage,
            role: "assistant",
            chatThreadId: props.chatThread.id,
          });

          streamResponse(response.type, JSON.stringify(response));
          controller.close();
        })
        .on("finalContent", async (content: string) => {
          await CreateChatMessage({
            name: AI_NAME,
            content: content,
            role: "assistant",
            chatThreadId: props.chatThread.id,
          });

          const response: AzureChatCompletion = {
            type: "finalContent",
            response: content,
          };
          streamResponse(response.type, JSON.stringify(response));
          controller.close();
        });
    },
  });

  return readableStream;
};
