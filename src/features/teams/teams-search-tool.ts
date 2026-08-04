import type { ChatCompletionTool } from "openai/resources/chat/completions";

import type { SlSearchScope } from "@/lib/sl-search-target";

export type TeamsInternalSearchToolArguments = {
  query: string;
  scope?: SlSearchScope;
  folder?: string;
};

const VALID_SCOPES = new Set<SlSearchScope>([
  "all",
  "personal",
  "dept_common",
  "global_common",
]);

export const TEAMS_INTERNAL_SEARCH_TOOL_NAME = "teams_internal_search";

export const TEAMS_INTERNAL_SEARCH_INSTRUCTIONS =
  "社内・SharePoint・部署共通・全社共通・個人資料について、文書を読んで回答する必要がある場合はteams_internal_searchを使用してください。" +
  "フォルダー指定は助詞や表記にかかわらず意味から解釈してください。たとえば「可能性調査というフォルダー」「『可能性調査』のフォルダーで」「可能性調査フォルダー内」は、いずれもfolder=可能性調査です。" +
  "部署共通はscope=dept_common、全社共通はscope=global_common、個人・自分・私の資料はscope=personalです。" +
  "queryには検索場所を示す語を除き、探す内容・比較項目・文書名を具体的に設定してください。" +
  "一般知識、雑談、天気、Web情報だけで答える質問ではこのツールを使用しないでください。";

export const TEAMS_INTERNAL_SEARCH_TOOL: ChatCompletionTool = {
  type: "function",
  function: {
    name: TEAMS_INTERNAL_SEARCH_TOOL_NAME,
    description:
      "アクセス権を適用して社内SharePoint文書を検索します。検索範囲やフォルダー名はユーザーの自然な表現から意味的に抽出してください。",
    parameters: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description:
            "検索内容。検索場所の表現を除き、文書名・対象・比較項目などを含めます。",
        },
        scope: {
          type: "string",
          enum: ["all", "personal", "dept_common", "global_common"],
          description:
            "検索範囲。部署共通はdept_common、全社共通はglobal_common、個人資料はpersonal、明示がなければall。",
        },
        folder: {
          type: "string",
          description:
            "明示されたSharePointフォルダー名。引用符、「という」、助詞「の」「で」「から」「内」などは含めません。",
        },
      },
      required: ["query"],
      additionalProperties: false,
    },
  },
};

export function parseTeamsInternalSearchToolArguments(
  value: string
): TeamsInternalSearchToolArguments {
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    throw new Error("Teams internal search tool returned invalid JSON.");
  }

  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error("Teams internal search tool arguments must be an object.");
  }

  const input = parsed as Record<string, unknown>;
  const query = normalizeText(input.query);
  if (!query) {
    throw new Error("Teams internal search tool query is required.");
  }

  const rawScope = normalizeText(input.scope);
  if (rawScope && !VALID_SCOPES.has(rawScope as SlSearchScope)) {
    throw new Error(`Unsupported Teams internal search scope: ${rawScope}`);
  }

  const folder = normalizeText(input.folder);
  return {
    query,
    ...(rawScope ? { scope: rawScope as SlSearchScope } : {}),
    ...(folder ? { folder } : {}),
  };
}

function normalizeText(value: unknown): string {
  if (typeof value !== "string") return "";
  return value.normalize("NFKC").trim().slice(0, 500);
}
