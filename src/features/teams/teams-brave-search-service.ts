import "server-only";

import type { TeamsSearchSource } from "./teams-search-service";
import { requiresTeamsInternalSearch } from "./teams-search-intent";
import { buildTeamsWebQuery } from "./teams-web-query";
import {
  extractWeatherPageText,
  isTrustedWeatherUrl,
} from "./teams-weather-page";

const BRAVE_SEARCH_ENDPOINT =
  "https://api.search.brave.com/res/v1/web/search";
const DEFAULT_RESULT_COUNT = 5;
const MAX_WEB_CONTEXT_CHARACTERS = 10_000;
const MAX_WEATHER_PAGES = 3;
const WEATHER_PAGE_TIMEOUT_MS = 6_000;

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
  const skipInternalSearch = !requiresTeamsInternalSearch(trimmed);
  if (mode === "off") {
    return { enabled: false, query: trimmed, skipInternalSearch };
  }

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
    const effectiveQuery = buildTeamsWebQuery(props.query);
    const params = new URLSearchParams({
      q: effectiveQuery.query,
      count: String(resolveResultCount()),
      extra_snippets: "true",
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
    const results = data.web?.results ?? [];
    const weatherPageText = effectiveQuery.enriched
      ? await loadWeatherPageText(results)
      : new Map<string, string>();
    return formatResults(
      results,
      props.startIndex,
      effectiveQuery.enriched,
      weatherPageText
    );
  } finally {
    clearTimeout(timeout);
  }
}

export function isTeamsBraveSearchConfigured(): boolean {
  return Boolean(process.env.BRAVE_SUBSCRIPTION_TOKEN?.trim());
}

function formatResults(
  results: BraveWebResult[],
  startIndex: number,
  queryEnriched: boolean,
  weatherPageText: Map<string, string>
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
      ...(result.extra_snippets ?? []).slice(0, 5).map((item) => item.trim()),
    ]
      .filter(Boolean)
      .join(" ");

    sources.push({ index, name, url, kind: "web" });
    const pageText = weatherPageText.get(url);
    if (!snippets && !pageText) continue;
    const section = `[${index}] ${name}\nURL: ${url}\n${snippets}${
      pageText ? `\nページ本文の気象情報:\n${pageText}` : ""
    }`;
    const remaining = MAX_WEB_CONTEXT_CHARACTERS - totalCharacters;
    if (remaining <= 0) break;
    const clipped = section.slice(0, remaining);
    sections.push(clipped);
    totalCharacters += clipped.length;
  }

  console.log("[teams-brave-search] completed", {
    results: results.length,
    sources: sources.length,
    queryEnriched,
  });

  return { context: sections.join("\n\n---\n\n"), sources };
}

async function loadWeatherPageText(
  results: BraveWebResult[]
): Promise<Map<string, string>> {
  const urls = Array.from(
    new Set(
      results
        .map((result) => normalizeHttpsUrl(result.url))
        .filter((url): url is string => Boolean(url && isTrustedWeatherUrl(url)))
    )
  ).slice(0, MAX_WEATHER_PAGES);

  const loaded = await Promise.all(
    urls.map(async (url) => [url, await fetchWeatherPageText(url)] as const)
  );
  const successful = loaded.filter((entry) => Boolean(entry[1]));
  console.log("[teams-brave-search] weather pages loaded", {
    requested: urls.length,
    loaded: successful.length,
  });
  return new Map(successful);
}

async function fetchWeatherPageText(url: string): Promise<string> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), WEATHER_PAGE_TIMEOUT_MS);
  try {
    const response = await fetch(url, {
      headers: {
        Accept: "text/html,application/xhtml+xml",
        "User-Agent": "AzureChat-TeamsBot/1.0",
      },
      redirect: "follow",
      signal: controller.signal,
    });
    if (!response.ok || !isTrustedWeatherUrl(response.url)) return "";
    const contentType = response.headers.get("content-type") ?? "";
    if (!contentType.toLowerCase().includes("text/html")) return "";
    return extractWeatherPageText((await response.text()).slice(0, 750_000));
  } catch (error) {
    console.warn("[teams-brave-search] weather page fetch skipped", {
      host: new URL(url).hostname,
      error: error instanceof Error ? error.message : String(error),
    });
    return "";
  } finally {
    clearTimeout(timeout);
  }
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
