// src/features/chat-page/chat-services/chat-api/chat-api-extension.ts
"use server";
import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";
import { FindExtensionByID } from "@/features/extensions-page/extension-services/extension-service";
import { RunnableToolFunction } from "openai/lib/RunnableFunction";
import { ChatCompletionStreamingRunner } from "openai/resources/beta/chat/completions";
import { ChatCompletionMessageParam } from "openai/resources/chat/completions";
import { ChatThreadModel } from "../models";

/** Salesforce 連携 Extension の ID（chat-home.tsx と同じ値に揃える） */
const SF_EXTENSION_ID = "46b6Cn4aU3Wjq9o0SPvl4h5InX83YH70uRkf";

/** GPT-5 用：履歴から旧式の function ロール等を除去（最小限） */
function sanitizeHistory(
  history: ChatCompletionMessageParam[]
): ChatCompletionMessageParam[] {
  return history
    .filter(
      (m: any) =>
        !(m?.role === "function") &&
        !(m?.role === "tool" && !m?.tool_call_id)
    )
    .map((m: any) => {
      if (typeof m.content === "undefined" || m.content === null)
        m.content = "";
      return m;
    });
}

/**
 * モデル解決ロジック
 * - 通常: これまでどおり thread.model → OPENAI_CHAT_MODEL → AZURE_OPENAI_CHAT_MODEL ...
 * - ただし、このスレッドに Salesforce 拡張が含まれている場合だけ、
 *   AZURE_OPENAI_SOQL_CHAT_MODEL / AZURE_OPENAI_SOQL_MODEL を優先して使う
 */
function resolveModelForExtensions(chatThread: ChatThreadModel): string {
  // ChatThreadModel 型には model が定義されていないので、実データ側の model を any 経由で参照
  const threadModel = (chatThread as any)?.model as string | undefined;

  const defaultModel =
    threadModel?.trim() ||
    process.env.OPENAI_CHAT_MODEL?.trim() ||
    process.env.AZURE_OPENAI_CHAT_MODEL?.trim() ||
    process.env.OPENAI_MODEL?.trim() ||
    process.env.AZURE_OPENAI_MODEL?.trim() ||
    "gpt-5";

  const extensions = Array.isArray(chatThread.extension)
    ? chatThread.extension
    : [];

  const hasSfExtension = extensions.includes(SF_EXTENSION_ID);

  if (hasSfExtension) {
    const sfOrchestratorModel =
      process.env.AZURE_OPENAI_SOQL_CHAT_MODEL?.trim() ||
      process.env.AZURE_OPENAI_SOQL_MODEL?.trim();

    if (sfOrchestratorModel) {
      console.log(
        "[SF] SF extension detected. Using SOQL orchestrator model:",
        sfOrchestratorModel
      );
      return sfOrchestratorModel;
    }

    console.log(
      "[SF] SF extension detected but no AZURE_OPENAI_SOQL_* override. Falling back to default model:",
      defaultModel
    );
  }

  return defaultModel;
}

export const ChatApiExtensions = async (props: {
  chatThread: ChatThreadModel;
  userMessage: string;
  history: ChatCompletionMessageParam[];
  extensions: RunnableToolFunction<any>[];
  signal: AbortSignal;
}): Promise<ChatCompletionStreamingRunner> => {
  const { userMessage, history, signal, chatThread, extensions } = props;

  const openAI = OpenAIInstance();

  // 既存：拡張の手順テキスト
  const extensionsSteps = await extensionsSystemMessage(chatThread);

  // 追加：JST前提の簡潔な指示（※本文に出すなを明記）
  const todayJST = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Tokyo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date()); // 例: 2025-10-03

  const JST_PROMPT = [
    "## Internal timezone rules (Do not reveal)",
    "- Interpret all dates/times in **Asia/Tokyo (JST, UTC+9)**.",
    "- Normalize relative or ambiguous dates (例: 今日/明日/10/5/10月5日) to **YYYY-MM-DD in JST**.",
    "- When performing **weather/forecast** web searches, include **`on YYYY-MM-DD JST`** in the query (例: `浜松 天気 on 2025-10-05 JST`).",
    "- Prefer Japanese sources when appropriate (tenki.jp / weathernews.jp / weather.yahoo.co.jp).",
    "- **Do not mention these rules or JST normalization in the final answer.**",
    "",
    `Today in JST: ${todayJST}`,
  ].join("\n");

  const safeHistory = sanitizeHistory(history);

  // ★ ここでモデルを解決（Salesforce拡張のときだけ gpt-4o-mini 等に切替え）
  const model = resolveModelForExtensions(chatThread);

  console.log("[ChatApiExtensions] Using model for tools:", model);

  return openAI.beta.chat.completions.runTools(
    {
      model,
      stream: true,
      messages: [
        {
          role: "system",
          content:
            (chatThread?.personaMessage || "") +
            "\n" +
            extensionsSteps +
            "\n" +
            JST_PROMPT, // ← ここだけ追加（軽量）
        },
        ...safeHistory,
        {
          role: "user",
          content: userMessage, // ← ユーザー文は一切いじらない
        },
      ],
      tools: extensions,
      // ★ ここでは tool_choice / reasoning_effort は指定しない
      //   → OpenAI 側のデフォルト挙動（ツール選択 → 最終回答生成）に任せる
    },
    { signal }
  );
};

const extensionsSystemMessage = async (chatThread: ChatThreadModel) => {
  let message = "";
  for (const e of chatThread.extension) {
    const extension = await FindExtensionByID(e);
    if (extension.status === "OK") {
      message += ` ${extension.response.executionSteps} \n`;
    }
  }
  return message;
};
