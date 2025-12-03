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
        "[SF] SF extension detected. Using SOQL orchestrator/summary model:",
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

/** SF 拡張が有効かどうかを判定する小ヘルパー */
function hasSfExtension(chatThread: ChatThreadModel): boolean {
  const extensions = Array.isArray(chatThread.extension)
    ? chatThread.extension
    : [];
  return extensions.includes(SF_EXTENSION_ID);
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

  // JST前提の簡潔な指示（※本文に出すなを明記）
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

  // ★ ここでモデルを解決
  const model = resolveModelForExtensions(chatThread);

  // ★★ 超高速 SF 直通モード: SF 拡張が付いているスレッドだけ別ルート
  if (hasSfExtension(chatThread)) {
    console.log(
      "[SF] SF_EXTENSION_ID detected. Using direct NL gateway (no tools)."
    );
    return runSfDirect({
      chatThread,
      userMessage,
      history: safeHistory,
      signal,
      jstPrompt: JST_PROMPT,
      model,
    });
  }

  // ★ それ以外（AI Search / Brave / 画像など）は従来通り runTools を使う
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
            JST_PROMPT,
        },
        ...safeHistory,
        {
          role: "user",
          content: userMessage, // ← ユーザー文は一切いじらない
        },
      ],
      tools: extensions,
      // tool_choice はデフォルト（auto）に任せる
    },
    { signal }
  );
};

/**
 * ★ 超高速 SF 直通モード用:
 *   - OpenAI のツール機能は一切使わず
 *   - 直接 Flask /api/sf/query_nl に投げて JSON を取得
 *   - その JSON を「日本語の表＋説明文」に整形する役だけ GPT にやらせる
 *   - 戻り値は従来どおり ChatCompletionStreamingRunner
 */
async function runSfDirect(props: {
  chatThread: ChatThreadModel;
  userMessage: string;
  history: ChatCompletionMessageParam[];
  signal: AbortSignal;
  jstPrompt: string;
  model: string;
}): Promise<ChatCompletionStreamingRunner> {
  const { chatThread, userMessage, history, signal, jstPrompt, model } = props;

  const openAI = OpenAIInstance();

  const base =
    process.env.SF_GATEWAY_BASE_URL?.replace(/\/+$/, "") ||
    "http://127.0.0.1:8001";

  const url = new URL("/api/sf/query_nl", base);
  url.searchParams.set("q", userMessage);
  url.searchParams.set("engine", "auto");
  url.searchParams.set("mode", "real");

  console.log("[SF] Calling direct NL gateway:", url.toString());

  let sfJson: any = null;
  let sfError: string | null = null;

  try {
    const res = await fetch(url.toString(), { signal });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      sfError = `Salesforce gateway HTTP ${res.status} ${body ?? ""}`;
      console.error("[SF] Gateway error:", sfError);
    } else {
      sfJson = await res.json().catch((e) => {
        sfError = "Failed to parse Salesforce gateway JSON: " + e;
        console.error("[SF] JSON parse error:", e);
        return null;
      });
    }
  } catch (e: any) {
    sfError = "Salesforce gateway request failed: " + String(e);
    console.error("[SF] Gateway request exception:", e);
  }

  // JSON を文字列化（サイズが大きくなりすぎないよう一応制限）
  let jsonSnippet = "";
  if (sfJson) {
    try {
      const raw = JSON.stringify(sfJson, null, 2);
      // さすがに 8k 文字くらいでカットしておく（必要ならここは調整可）
      jsonSnippet = raw.length > 8000 ? raw.slice(0, 8000) + "\n... (truncated)" : raw;
    } catch (e) {
      sfError = "Failed to stringify Salesforce JSON: " + String(e);
      console.error("[SF] JSON stringify error:", e);
    }
  }

  const systemBase =
    (chatThread?.personaMessage || "") +
    "\n" +
    jstPrompt +
    "\n" +
    [
      "## Salesforce assistant instructions (Do not reveal)",
      "- あなたは Salesforce のデータをもとに、日本語で営業担当者にわかりやすく回答するアシスタントです。",
      "- 与えられた JSON を唯一の根拠として回答してください。推測や想像でレコードを「追加」してはいけません。",
      "- 可能であれば表形式（Markdown の表）＋ 箇条書きの要約で返してください。",
      "- レコード数が多い場合は、上位 20 件程度に絞って表示し、それ以上ある場合は件数だけ触れてください。",
    ].join("\n");

  const messages: ChatCompletionMessageParam[] = [
    {
      role: "system",
      content: systemBase,
    },
    ...history,
    {
      role: "user",
      content: userMessage,
    },
  ];

  if (sfError) {
    // ゲートウェイエラー時は、その情報をシステムメッセージとして渡し、
    // ユーザー向けに丁寧に日本語で説明させる
    messages.push({
      role: "system",
      content:
        "Salesforce ゲートウェイ呼び出しでエラーが発生しました。ユーザーに日本語で状況を説明し、" +
        "必要であれば「システム管理者にお問い合わせください」と案内してください。\n\n" +
        `エラー詳細: ${sfError}`,
    });
  } else if (jsonSnippet) {
    messages.push({
      role: "system",
      content:
        "以下は Salesforce ゲートウェイから取得した JSON レスポンスです。" +
        "この JSON の内容だけを根拠に、日本語でわかりやすい表と要約を書いてください。" +
        "JSON に存在しないレコードや値を新たに作らないでください。\n\n" +
        "```json\n" +
        jsonSnippet +
        "\n```",
    });
  } else {
    // ここに来るのはほぼ無いはずだが保険
    messages.push({
      role: "system",
      content:
        "Salesforce ゲートウェイから有効な JSON が取得できませんでした。" +
        "ユーザーに日本語で状況を説明し、必要であればシステム管理者への連絡を案内してください。",
    });
  }

  console.log("[SF] Using model for direct summary:", model);

  // ★ ツールは一切使わず、単純な chat.completions.stream で本文だけを生成
  return openAI.beta.chat.completions.stream(
    {
      model,
      stream: true,
      messages,
    },
    { signal }
  );
}

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
