import "server-only";

import type { TeamsSearchSource } from "./teams-search-service";

const BRAVE_SEARCH_ENDPOINT =
  "https://api.search.brave.com/res/v1/web/search";
const DEFAULT_RESULT_COUNT = 5;
const MAX_WEB_CONTEXT_CHARACTERS = 10_000;

type BraveWebResult = {
  title?: string;
  url?: string;
  description?: string;
  extra_snippets?: string[];
};

export type TeamsBraveSearchResult = {
  context: string;
  sources: TeamsSearchSource[];
};

export function resolveBraveSearchRequest(message: string): {
  enabled: boolean;
  query: string;
  skipInternalSearch: boolean;
} {
  const trimmed = message.trim();
  const explicit = trimmed.match(/^\/(?:web|brave)\s+([\s\S]+)$/i);
  if (explicit?.[1]?.trim()) {
    return {
      enabled: true,
      query: explicit[1].trim(),
      skipInternalSearch: true,
    };
  }

  const mode = process.env.TEAMS_BRAVE_SEARCH_MODE?.trim().toLowerCase();
  if (mode === "off") {
    return { enabled: false, query: trimmed, skipInternalSearch: false };
  }

  const internalKnowledgePattern =
    /(社内|規程|規定|就業|人事|総務|申請|手順書|マニュアル|sharepoint|sl文書)/i;
  const clearlyPublicPattern =
    /(天気|天候|気温|降水|台風|ニュース|株価|為替|選挙結果|スポーツ結果)/i;
  const skipInternalSearch =
    clearlyPublicPattern.test(trimmed) && !internalKnowledgePattern.test(trimmed);

  if (mode === "always") {
    return { enabled: true, query: trimmed, skipInternalSearch };
  }

  const autoSearchPattern =
    /(最新|現在|今日|ニュース|天気|株価|為替|価格|検索して|調べて|ウェブ|web|brave|インターネット|外部情報|公式サイト)/i;
  return {
    enabled: autoSearchPattern.test(trimmed),
    query: trimmed,
    skipInternalSearch,
  };
}

export async function searchBraveWeb(props: {
  query: string;
  startIndex: number;
}): Promise<TeamsBraveSearchResult> {
  const apiKey = process.env.BRAVE_SUBSCRIPTION_TOKEN?.trim();
  if (!apiKey) {
    throw new Error("BRAVE_SUBSCRIPTION_TOKEN is not configured.");
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 10_000);

  try {
    const params = new URLSearchParams({
      q: props.query,
      count: String(resolveResultCount()),
    });
    const response = await fetch(`${BRAVE_SEARCH_ENDPOINT}?${params}`, {
      headers: {
        Accept: "application/json",
        "X-Subscription-Token": apiKey,
      },
      signal: controller.signal,
    });

    if (!response.ok) {
      const detail = await response.text().catch(() => "");
      throw new Error(
        `Brave Search returned HTTP ${response.status}${
          detail ? `: ${detail.slice(0, 200)}` : ""
        }`
      );
    }

    const data = (await response.json()) as {
      web?: { results?: BraveWebResult[] };
    };
    return formatResults(data.web?.results ?? [], props.startIndex);
  } finally {
    clearTimeout(timeout);
  }
}

export function isTeamsBraveSearchConfigured(): boolean {
  return Boolean(process.env.BRAVE_SUBSCRIPTION_TOKEN?.trim());
}

function formatResults(
  results: BraveWebResult[],
  startIndex: number
): TeamsBraveSearchResult {
  const sources: TeamsSearchSource[] = [];
  const sections: string[] = [];
  let totalCharacters = 0;

  for (const result of results) {
    const url = normalizeHttpsUrl(result.url);
    if (!url) continue;

    const index = startIndex + sources.length;
    const name = result.title?.trim() || new URL(url).hostname;
    const snippets = [
      result.description?.trim(),
      ...(result.extra_snippets ?? []).slice(0, 2).map((item) => item.trim()),
    ]
      .filter(Boolean)
      .join(" ");

    sources.push({ index, name, url, kind: "web" });
    if (!snippets) continue;

    const section = `[${index}] ${name}\nURL: ${url}\n${snippets}`;
    const remaining = MAX_WEB_CONTEXT_CHARACTERS - totalCharacters;
    if (remaining <= 0) break;
    const clipped = section.slice(0, remaining);
    sections.push(clipped);
    totalCharacters += clipped.length;
  }

  console.log("[teams-brave-search] completed", {
    results: results.length,
    sources: sources.length,
  });

  return { context: sections.join("\n\n---\n\n"), sources };
}

function resolveResultCount(): number {
  const parsed = Number.parseInt(
    process.env.TEAMS_BRAVE_SEARCH_TOP ?? "",
    10
  );
  if (!Number.isFinite(parsed)) return DEFAULT_RESULT_COUNT;
  return Math.min(Math.max(parsed, 1), 10);
}

function normalizeHttpsUrl(value?: string): string | null {
  const normalized = value?.trim();
  if (!normalized) return null;

  try {
    const url = new URL(normalized);
    return url.protocol === "https:" ? url.toString() : null;
  } catch {
    return null;
  }
}
