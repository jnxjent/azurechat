import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";
import type {
  ChatCompletionMessageParam,
  ChatCompletionMessageToolCall,
} from "openai/resources/chat/completions";
import {
  searchTeamsKnowledgeWithTarget,
  type TeamsSearchSource,
} from "./teams-search-service";
import {
  resolveBraveSearchRequest,
  searchBraveWeb,
} from "./teams-brave-search-service";
import { resolveTeamsInstantReply } from "./teams-instant-reply";
import { buildTeamsWebQuery } from "./teams-web-query";
import {
  parseTeamsInternalSearchToolArguments,
  TEAMS_INTERNAL_SEARCH_INSTRUCTIONS,
  TEAMS_INTERNAL_SEARCH_TOOL,
  TEAMS_INTERNAL_SEARCH_TOOL_NAME,
} from "./teams-search-tool";

const MAX_HISTORY_MESSAGES = 20;
const MAX_INTERNAL_SEARCH_CALLS = 3;
const conversations = new Map<string, ChatCompletionMessageParam[]>();

export type TeamsChatResult =
  | { type: "reply"; text: string }
  | { type: "reset"; text: string };

export async function createTeamsChatReply(props: {
  conversationId: string;
  message: string;
  userEmail?: string | null;
}): Promise<TeamsChatResult> {
  const message = props.message.trim();
  if (!message) {
    return { type: "reply", text: "メッセージを入力してください。" };
  }

  if (message.toLowerCase() === "/reset") {
    conversations.delete(props.conversationId);
    return { type: "reset", text: "この会話の履歴をリセットしました。" };
  }

  const instantReply = resolveTeamsInstantReply(message);
  if (instantReply) {
    return { type: "reply", text: instantReply };
  }

  assertAzureOpenAIConfiguration();

  const history = conversations.get(props.conversationId) ?? [];
  const braveRequest = resolveBraveSearchRequest(message);
  const webSearch = braveRequest.enabled
    ? await searchBraveWeb({
        query: braveRequest.query,
        startIndex: 1,
      })
    : { context: "", sources: [] };
  const allSources = [...webSearch.sources];
  const userContent = buildUserContent(
    braveRequest.query,
    "",
    webSearch.context,
    false,
    braveRequest.enabled
  );
  const baseSystemPrompt =
    process.env.TEAMS_SYSTEM_PROMPT?.trim() ||
    "あなたはAzureChatのAIアシスタントです。ユーザーと同じ言語で、正確かつ簡潔に回答してください。Teamsで読みやすいMarkdownを使用してください。社内情報については提供された検索資料だけを根拠に回答し、根拠がない場合は分からないと回答してください。検索資料を使用した文には必ず [1] の形式で出典番号を付けてください。";
  const messages: ChatCompletionMessageParam[] = [
    {
      role: "system",
      content: `${baseSystemPrompt}\n\n${TEAMS_INTERNAL_SEARCH_INSTRUCTIONS}`,
    },
    ...history,
    { role: "user", content: userContent },
  ];

  const openai = OpenAIInstance();
  let answer = "";
  let internalSearchCalls = 0;
  let forceFinalAnswer = false;

  while (!answer) {
    const completion = await openai.chat.completions.create({
      model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME ?? "",
      messages,
      ...(!braveRequest.skipInternalSearch && !forceFinalAnswer
        ? { tools: [TEAMS_INTERNAL_SEARCH_TOOL], tool_choice: "auto" as const }
        : {}),
    });
    const assistantMessage = completion.choices[0]?.message;
    if (!assistantMessage) {
      throw new Error("Azure OpenAI returned no assistant message.");
    }

    const toolCalls = forceFinalAnswer
      ? []
      : assistantMessage.tool_calls?.filter(
          (toolCall) => toolCall.type === "function"
        ) ?? [];
    if (toolCalls.length === 0) {
      answer = assistantMessage.content?.trim() ?? "";
      break;
    }

    messages.push({
      role: "assistant",
      content: assistantMessage.content,
      tool_calls: toolCalls,
    });

    for (const toolCall of toolCalls) {
      const toolContent = await executeInternalSearchTool({
        toolCall,
        userEmail: props.userEmail,
        allSources,
        searchCallAllowed: internalSearchCalls < MAX_INTERNAL_SEARCH_CALLS,
      });
      if (toolCall.function.name === TEAMS_INTERNAL_SEARCH_TOOL_NAME) {
        internalSearchCalls++;
      }
      messages.push({
        role: "tool",
        tool_call_id: toolCall.id,
        content: toolContent,
      });
    }

    forceFinalAnswer = internalSearchCalls >= MAX_INTERNAL_SEARCH_CALLS;
  }

  if (!answer) {
    throw new Error("Azure OpenAI returned an empty response.");
  }
  const answerWithSources = appendSources(answer, allSources);

  const updatedHistory: ChatCompletionMessageParam[] = [
    ...history,
    { role: "user", content: message },
    { role: "assistant", content: answerWithSources },
  ];
  conversations.set(
    props.conversationId,
    updatedHistory.slice(-MAX_HISTORY_MESSAGES)
  );

  return { type: "reply", text: answerWithSources };
}

async function executeInternalSearchTool(props: {
  toolCall: ChatCompletionMessageToolCall;
  userEmail?: string | null;
  allSources: TeamsSearchSource[];
  searchCallAllowed: boolean;
}): Promise<string> {
  if (props.toolCall.function.name !== TEAMS_INTERNAL_SEARCH_TOOL_NAME) {
    return "未対応のツールです。別のツールを呼ばず、回答を作成してください。";
  }
  if (!props.searchCallAllowed) {
    return "社内検索回数の上限に達しました。取得済みの資料だけで回答してください。";
  }

  try {
    const args = parseTeamsInternalSearchToolArguments(
      props.toolCall.function.arguments
    );
    const result = await searchTeamsKnowledgeWithTarget({
      ...args,
      userEmail: props.userEmail,
    });
    if (!result.userEmail) {
      return "Teamsユーザーのメールアドレスを解決できないため、アクセス権付き社内検索を実行できませんでした。";
    }

    const context = mergeSearchSources(
      result.context,
      result.sources,
      props.allSources
    );
    return `社内検索資料:\n${
      context || "該当する社内文書は見つかりませんでした。"
    }`;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.warn("[teams-search-tool] execution failed", message);
    return `社内検索を実行できませんでした: ${message.slice(0, 300)}`;
  }
}

function mergeSearchSources(
  context: string,
  incoming: TeamsSearchSource[],
  allSources: TeamsSearchSource[]
): string {
  const indexMap = new Map<number, number>();
  for (const source of incoming) {
    const existing = allSources.find(
      (item) =>
        item.kind === source.kind &&
        item.name === source.name &&
        (item.url ?? "") === (source.url ?? "")
    );
    if (existing) {
      indexMap.set(source.index, existing.index);
      continue;
    }

    const nextIndex =
      allSources.reduce((max, item) => Math.max(max, item.index), 0) + 1;
    indexMap.set(source.index, nextIndex);
    allSources.push({ ...source, index: nextIndex });
  }

  return context.replace(/\[(\d+)\]/g, (match, rawIndex: string) => {
    const mapped = indexMap.get(Number(rawIndex));
    return mapped ? `[${mapped}]` : match;
  });
}

function buildUserContent(
  message: string,
  internalContext: string,
  webContext: string,
  internalSearchPerformed: boolean,
  webSearchPerformed: boolean
): string {
  const sections = [`質問:\n${message}`];
  if (internalSearchPerformed) {
    sections.push(
      `社内検索資料:\n${
        internalContext || "該当する社内文書は見つかりませんでした。"
      }`
    );
  }
  if (webSearchPerformed) {
    const weatherQuery = buildTeamsWebQuery(message);
    if (weatherQuery.enriched) {
      sections.push(
        `天気回答要件:\n対象日時をJSTで確認してください。検索資料全体を突き合わせ、取得できた範囲で「天気」「最高気温」「最低気温」「降水確率」を簡潔に回答してください。一つの資料で値が「-」でも、他の資料に具体値があればそちらを採用してください。資料にない項目は推測せず「確認できません」としてください。\n検索条件: ${weatherQuery.query}`
      );
    }
    sections.push(
      `Web検索資料:\n${webContext || "Web検索結果は見つかりませんでした。"}`
    );
  }
  return sections.join("\n\n");
}

function appendSources(answer: string, sources: TeamsSearchSource[]): string {
  if (sources.length === 0) return answer;

  const citedIndexes = new Set(
    Array.from(answer.matchAll(/\[(\d+)\]/g), (match) => Number(match[1]))
  );
  const citedSources = sources.filter((source) =>
    citedIndexes.has(source.index)
  );
  if (citedSources.length === 0) return answer;

  const formatSources = (items: TeamsSearchSource[]) =>
    items.slice(0, 5).map((source) => {
      const index = `[${source.index}]`;
      const name = escapeMarkdownLinkText(compactSourceName(source.name));
      return source.url
        ? `${index} [${name}](${source.url})`
        : `${index} ${name}`;
    });

  const internal = formatSources(
    citedSources.filter((source) => source.kind === "internal")
  );
  const web = formatSources(
    citedSources.filter((source) => source.kind === "web")
  );
  const sections = [
    ...(internal.length > 0 ? [`社内: ${internal.join(" ・ ")}`] : []),
    ...(web.length > 0 ? [`Web: ${web.join(" ・ ")}`] : []),
  ].join("\n");

  return `${answer}\n\n**参照:** ${sections}`;
}

function escapeMarkdownLinkText(value: string): string {
  return value.replace(/([\\[\]])/g, "\\$1");
}

function compactSourceName(value: string): string {
  const normalized = value.replace(/\s+/g, " ").trim();
  return normalized.length > 60
    ? `${normalized.slice(0, 57)}...`
    : normalized;
}

function assertAzureOpenAIConfiguration(): void {
  const required = [
    "AZURE_OPENAI_API_KEY",
    "AZURE_OPENAI_API_INSTANCE_NAME",
    "AZURE_OPENAI_API_DEPLOYMENT_NAME",
    "AZURE_OPENAI_API_VERSION",
  ];
  const missing = required.filter((name) => !process.env[name]?.trim());

  if (missing.length > 0) {
    throw new Error(
      `Azure OpenAI configuration is incomplete: ${missing.join(", ")}`
    );
  }
}
