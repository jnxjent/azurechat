// src/features/chat-page/chat-services/chat-api/open-ai-stream.ts
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
      let controllerClosed = false;

      const closeController = () => {
        if (controllerClosed) return;
        controllerClosed = true;
        try {
          controller.close();
        } catch {
          // already closed by client disconnect
        }
      };

      const streamResponse = (event: string, value: string) => {
        if (controllerClosed) return;
        try {
          controller.enqueue(encoder.encode(`event: ${event} \n`));
          controller.enqueue(encoder.encode(`data: ${value} \n\n`));
        } catch {
          controllerClosed = true;
        }
      };

      let lastMessage = "";

      // 🔹 ツール呼び出し（GPT-5 runTools → functionCall にマッピング）
      runner
        .on("functionCall", (fnCall: any) => {
          try {
            const fn = (fnCall as any).function ?? {};
            const name = fn.name ?? "tool";
            const args =
              typeof fn.arguments === "string"
                ? fn.arguments
                : JSON.stringify(fn.arguments ?? {});

            const response: AzureChatCompletion = {
              type: "functionCall",
              response: {
                name,
                arguments: args,
              },
            };

            streamResponse(response.type, JSON.stringify(response));
          } catch (e) {
            console.log("⚠️ functionCall mapping error:", e);
          }
        })
        // 🔹 ツール実行結果 → functionCallResult にマッピング
        .on("functionCallResult", (fnResult: any) => {
          try {
            const payload =
              typeof fnResult === "string"
                ? fnResult
                : JSON.stringify(fnResult);

            const response: AzureChatCompletion = {
              type: "functionCallResult",
              response: payload,
            };

            streamResponse(response.type, JSON.stringify(response));
          } catch (e) {
            console.log("⚠️ functionCallResult mapping error:", e);
          }
        })
        // 🔹 通常のコンテンツ delta
        .on("content", () => {
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
        .on("abort", () => {
          const response: AzureChatCompletionAbort = {
            type: "abort",
            response: "Chat aborted",
          };
          streamResponse(response.type, JSON.stringify(response));
          closeController();
        })
        .on("error", async (error: any) => {
          console.log("🔴 error", error);
          const response: AzureChatCompletion = {
            type: "error",
            response: error?.message ?? String(error),
          };

          if (lastMessage) {
            await CreateChatMessage({
              name: AI_NAME,
              content: lastMessage,
              role: "assistant",
              chatThreadId: chatThread.id,
            });
          }

          streamResponse(response.type, JSON.stringify(response));
          closeController();
        })
        .on("finalContent", async (content: string) => {
          await CreateChatMessage({
            name: AI_NAME,
            content: content,
            role: "assistant",
            chatThreadId: chatThread.id,
          });

          const response: AzureChatCompletion = {
            type: "finalContent",
            response: content,
          };
          streamResponse(response.type, JSON.stringify(response));
          closeController();
        });
    },
  });

  return readableStream;
};
