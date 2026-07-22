import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";
import type { ChatCompletionMessageParam } from "openai/resources/chat/completions";
import {
  searchTeamsKnowledge,
  type TeamsSearchSource,
} from "./teams-search-service";
import {
  resolveBraveSearchRequest,
  searchBraveWeb,
} from "./teams-brave-search-service";
import { resolveTeamsInstantReply } from "./teams-instant-reply";

const MAX_HISTORY_MESSAGES = 20;
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
  const internalSearch = braveRequest.skipInternalSearch
    ? { context: "", sources: [] as TeamsSearchSource[] }
    : await searchTeamsKnowledge({
        query: braveRequest.query,
        userEmail: props.userEmail,
      });
  const webSearch = braveRequest.enabled
    ? await searchBraveWeb({
        query: braveRequest.query,
        startIndex: internalSearch.sources.length + 1,
      })
    : { context: "", sources: [] };
  const allSources = [...internalSearch.sources, ...webSearch.sources];
  const userContent = buildUserContent(
    braveRequest.query,
    internalSearch.context,
    webSearch.context,
    !braveRequest.skipInternalSearch,
    braveRequest.enabled
  );
  const messages: ChatCompletionMessageParam[] = [
    {
      role: "system",
      content:
        process.env.TEAMS_SYSTEM_PROMPT?.trim() ||
        "あなたはAzureChatのAIアシスタントです。ユーザーと同じ言語で、正確かつ簡潔に回答してください。Teamsで読みやすいMarkdownを使用してください。社内情報については提供された検索資料だけを根拠に回答し、根拠がない場合は分からないと回答してください。検索資料を使用した文には必ず [1] の形式で出典番号を付けてください。",
    },
    ...history,
    { role: "user", content: userContent },
  ];

  const completion = await OpenAIInstance().chat.completions.create({
    model: process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME ?? "",
    messages,
  });

  const answer = completion.choices[0]?.message?.content?.trim();
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
