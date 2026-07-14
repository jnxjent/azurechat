// src/features/chat-page/chat-services/chat-api/open-ai-stream.ts
import { AI_NAME } from "@/features/theme/theme-config";
import { ChatCompletionStreamingRunner } from "openai/resources/beta/chat/completions";
import { CreateChatMessage } from "../chat-message-service";
import {
  AzureChatCompletion,
  AzureChatCompletionAbort,
  ChatThreadModel,
} from "../models";

/**
 * LLM が壊れた Markdown リンクを生成した場合に修復する。
 * 壊れたパターン例: https://[host](https://host/path?sig=...)/path?sig=...
 * 正しい形式:      [fileName](https://host/path?sig=...)
 */
function repairBrokenMarkdownUrls(text: string): string {
  // Pattern: https://[linkText](https://innerUrl)/trailing
  // The innerUrl inside () is the complete correct URL.
  return text.replace(
    /https?:\/\/\[([^\]]*)\]\((https?:\/\/[^)]+)\)([^\s"'<>\]]*)/g,
    (_match, _linkText, innerUrl, trailing) => {
      // trailing may be empty or a duplicate path fragment — discard it
      const fullUrl = trailing
        ? innerUrl.includes("?") ? innerUrl : innerUrl + trailing
        : innerUrl;
      const pathPart = fullUrl.split("?")[0];
      const fileName =
        decodeURIComponent(pathPart.split("/").filter(Boolean).pop() ?? "") ||
        "ファイルをダウンロード";
      return `[${fileName}](${fullUrl})`;
    }
  );
}

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
      // functionCallResult をまとめて収集し finalContent で DB 保存する
      // （revalidate 後の props.messages でも tool メッセージが残るようにするため）
      const pendingToolResults: string[] = [];

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

            // SSE 送信と同時に DB 保存用に収集（finalContent で保存）
            pendingToolResults.push(payload);

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
          const repairedContent = repairBrokenMarkdownUrls(content);
          if (repairedContent !== content) {
            console.warn("[open-ai-stream] repaired broken markdown URL in finalContent");
          }

          // tool results を assistant より先に保存することで
          // revalidate 後の props.messages でも messages[index-1] が tool になる
          for (const result of pendingToolResults) {
            await CreateChatMessage({
              name: "tool",
              content: result,
              role: "tool",
              chatThreadId: chatThread.id,
            });
          }

          await CreateChatMessage({
            name: AI_NAME,
            content: repairedContent,
            role: "assistant",
            chatThreadId: chatThread.id,
          });

          const response: AzureChatCompletion = {
            type: "finalContent",
            response: repairedContent,
          };
          streamResponse(response.type, JSON.stringify(response));
          closeController();
        });
    },
  });

  return readableStream;
};
