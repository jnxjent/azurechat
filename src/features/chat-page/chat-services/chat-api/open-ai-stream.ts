// src/features/chat-page/chat-services/chat-api/open-ai-stream.ts
import { AI_NAME } from "@/features/theme/theme-config";
import { ChatCompletionStreamingRunner } from "openai/resources/beta/chat/completions";
import { CreateChatMessage } from "../chat-message-service";
import {
  AzureChatCompletion,
  AzureChatCompletionAbort,
  ChatThreadModel,
} from "../models";
import {
  CitationMarkupItem,
  extractCitationItems,
  formatCitationMarkup,
  removeCitationMarkup,
} from "@/features/ui/markdown/citation-markup";

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

function buildToolResultFallbackContent(toolResults: string[]): string {
  const latest = toolResults.at(-1) ?? "";
  if (!latest) return "処理が完了しました。";

  try {
    const parsed = JSON.parse(latest) as Record<string, unknown>;
    const imageUrl = typeof parsed.url === "string" ? parsed.url : "";
    if (imageUrl) {
      return `画像の処理が完了しました。\n\n![生成画像](${imageUrl})`;
    }
    const downloadUrl =
      typeof parsed.downloadUrl === "string" ? parsed.downloadUrl : "";
    if (downloadUrl) {
      return `ファイルの処理が完了しました。\n\n[ファイルを開く](${downloadUrl})`;
    }
    if (typeof parsed.error === "string" && parsed.error) {
      return parsed.error;
    }
    if (typeof parsed.message === "string" && parsed.message) {
      return parsed.message;
    }
    const result =
      typeof parsed.result === "object" && parsed.result !== null
        ? (parsed.result as Record<string, unknown>)
        : undefined;
    if (
      result &&
      typeof result.assistantMessage === "string" &&
      result.assistantMessage
    ) {
      return result.assistantMessage;
    }
  } catch {
    // String tool results are already suitable as a fallback response.
  }

  return latest;
}

function extractToolCitationItems(toolResults: string[]): CitationMarkupItem[] {
  const items: CitationMarkupItem[] = [];
  const resultRe =
    /^\s*\[\d+\]\.\s*file name:\s*(.+?)\s*\r?\n\s*file id:\s*([^\s\r\n]+)/gim;

  for (const result of toolResults) {
    items.push(...extractCitationItems(result));
    let match: RegExpExecArray | null;
    while ((match = resultRe.exec(result)) !== null) {
      const name = match[1].trim();
      const id = match[2].trim();
      if (name && id) items.push({ name, id });
    }
    resultRe.lastIndex = 0;
  }

  return Array.from(new Map(items.map((item) => [item.id, item])).values());
}

function hasGeneratedPptxResult(toolResults: string[]): boolean {
  return toolResults.some((toolResult) => {
    try {
      const parsed = JSON.parse(toolResult) as Record<string, unknown>;
      const fileName =
        typeof parsed.fileName === "string" ? parsed.fileName : "";
      const displayName =
        typeof parsed.displayName === "string" ? parsed.displayName : "";
      const downloadUrl =
        typeof parsed.downloadUrl === "string" ? parsed.downloadUrl : "";
      const message = typeof parsed.message === "string" ? parsed.message : "";

      return (
        /\.pptx$/i.test(fileName) ||
        /\.pptx$/i.test(displayName) ||
        /\.pptx(?:$|[?#])/i.test(downloadUrl) ||
        (Boolean(downloadUrl) && /PowerPoint/i.test(message))
      );
    } catch {
      return false;
    }
  });
}

/**
 * Rebuild citations from IDs that were actually created by server-side tools.
 * This prevents model-specific Markdoc formatting differences from leaking into UI.
 */
function normalizeFinalCitations(
  content: string,
  toolResults: string[]
): string {
  // Search results may be used to ground a generated deck, but they are not
  // citations for the short PPTX download response. Do not expose every
  // retrieval candidate as a citation list after PPTX creation or editing.
  if (hasGeneratedPptxResult(toolResults)) {
    const body = removeCitationMarkup(content).trimEnd();
    console.log("[open-ai-stream] suppressed citations for PPTX result");
    return body;
  }

  const available = extractToolCitationItems(toolResults);
  if (!available.length) return content;

  const availableById = new Map(available.map((item) => [item.id, item]));
  const requested = extractCitationItems(content)
    .map((item) => availableById.get(item.id))
    .filter((item): item is CitationMarkupItem => Boolean(item));

  // Prefer the model-selected subset. If its tag was malformed beyond parsing,
  // use one server-created citation per document name instead of losing citations.
  const selected = requested.length
    ? requested
    : Array.from(
        new Map(
          available.map((item) => [item.name.toLocaleLowerCase(), item])
        ).values()
      );
  const citation = formatCitationMarkup(selected);
  if (!citation) return content;

  const body = removeCitationMarkup(content).trimEnd();
  console.log(
    `[open-ai-stream] normalized citations selected=${selected.length} available=${available.length}`
  );
  return body ? `${body}\n\n${citation}` : citation;
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
      let finalized = false;
      // functionCallResult をまとめて収集し finalContent で DB 保存する
      // （revalidate 後の props.messages でも tool メッセージが残るようにするため）
      const pendingToolResults: string[] = [];

      const finalizeStream = async (content: string) => {
        if (finalized) return;
        finalized = true;

        const citationNormalizedContent = normalizeFinalCitations(
          content || "",
          pendingToolResults
        );
        const repairedContent = repairBrokenMarkdownUrls(
          citationNormalizedContent
        );
        if (repairedContent !== citationNormalizedContent) {
          console.warn(
            "[open-ai-stream] repaired broken markdown URL in finalContent"
          );
        }

        try {
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
        } catch (error) {
          // A persistence failure must not leave the browser stream open
          // forever after an otherwise successful tool execution.
          console.error("[open-ai-stream] final message persistence failed", error);
        } finally {
          const response: AzureChatCompletion = {
            type: "finalContent",
            response: repairedContent,
          };
          streamResponse(response.type, JSON.stringify(response));
          closeController();
        }
      };

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
          if (finalized) return;
          finalized = true;
          const response: AzureChatCompletionAbort = {
            type: "abort",
            response: "Chat aborted",
          };
          streamResponse(response.type, JSON.stringify(response));
          closeController();
        })
        .on("error", async (error: any) => {
          if (finalized) return;
          finalized = true;
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
        .on("finalContent", (content: string) => {
          void finalizeStream(content);
        })
        .on("end", () => {
          if (finalized) return;
          const fallbackContent = pendingToolResults.length
            ? buildToolResultFallbackContent(pendingToolResults)
            : lastMessage || "処理が完了しました。";
          console.warn(
            "[open-ai-stream] runner ended without finalContent; using fallback"
          );
          void finalizeStream(fallbackContent);
        });
    },
  });

  return readableStream;
};
