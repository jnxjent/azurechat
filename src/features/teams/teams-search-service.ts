import "server-only";

import {
  ExtensionSimilaritySearch,
  type DocumentSearchResponse,
} from "@/features/chat-page/chat-services/azure-ai-search/azure-ai-search";
import { hashValue } from "@/features/auth-page/helpers";
import { resolveSlAccess } from "@/lib/sl-dept";

const DEFAULT_SEARCH_TOP = 8;
const MAX_CONTEXT_CHARACTERS = 24_000;

export type TeamsSearchSource = {
  index: number;
  name: string;
  url?: string;
  kind: "internal" | "web";
};

export type TeamsSearchResult = {
  context: string;
  sources: TeamsSearchSource[];
  userEmail: string;
  dept: string;
};

export type TeamsOfficeFileCandidate = {
  name: string;
  url: string;
};

export async function searchTeamsKnowledge(props: {
  query: string;
  userEmail?: string | null;
}): Promise<TeamsSearchResult> {
  assertSearchConfiguration();

  const userEmail = resolveTeamsSearchUserEmail(props.userEmail);
  if (!userEmail) {
    throw new Error(
      "Teams AI Search user is unresolved. Set TEAMS_TEST_USER_EMAIL for local Playground testing."
    );
  }

  const access = resolveSlAccess(userEmail);
  const response = await ExtensionSimilaritySearch({
    searchText: props.query,
    vectors: ["embedding"],
    apiKey: requiredEnv("AZURE_SEARCH_API_KEY"),
    searchName: requiredEnv("AZURE_SEARCH_NAME"),
    indexName: requiredEnv("AZURE_SEARCH_INDEX_NAME"),
    filter: "isSlDoc eq true",
    deptLower: access.dept,
    userHash: hashValue(userEmail),
    top: resolveSearchTop(),
  });

  if (response.status !== "OK") {
    const detail = response.errors?.map((error) => error.message).join("; ");
    throw new Error(`Teams AI Search failed${detail ? `: ${detail}` : "."}`);
  }

  const formatted = formatSearchResults(response.response);
  console.log("[teams-search] completed", {
    dept: access.dept,
    role: access.role,
    results: response.response.length,
    sources: formatted.sources.length,
  });

  return {
    ...formatted,
    userEmail,
    dept: access.dept,
  };
}

export async function findTeamsOfficeFileCandidates(props: {
  query: string;
  userEmail?: string | null;
  extensions: string[];
}): Promise<{
  exactMatches: TeamsOfficeFileCandidate[];
  suggestions: TeamsOfficeFileCandidate[];
}> {
  assertSearchConfiguration();

  const userEmail = resolveTeamsSearchUserEmail(props.userEmail);
  if (!userEmail) {
    throw new Error(
      "Teams Office user is unresolved. Set TEAMS_TEST_USER_EMAIL for local Playground testing."
    );
  }

  const access = resolveSlAccess(userEmail);
  const response = await ExtensionSimilaritySearch({
    searchText: props.query,
    vectors: ["embedding"],
    apiKey: requiredEnv("AZURE_SEARCH_API_KEY"),
    searchName: requiredEnv("AZURE_SEARCH_NAME"),
    indexName: requiredEnv("AZURE_SEARCH_INDEX_NAME"),
    filter: "isSlDoc eq true",
    deptLower: access.dept,
    userHash: hashValue(userEmail),
    top: 20,
  });

  if (response.status !== "OK") {
    const detail = response.errors?.map((error) => error.message).join("; ");
    throw new Error(`Teams Office file search failed${detail ? `: ${detail}` : "."}`);
  }

  const allowedExtensions = new Set(
    props.extensions.map((extension) => extension.toLowerCase())
  );
  const unique = new Map<string, TeamsOfficeFileCandidate>();

  for (const item of response.response) {
    const document = item.document;
    const name = document.metadata?.trim();
    const url = (document.effectiveFileUrl ?? document.fileUrl)?.trim();
    if (!name || !url || !/^https:\/\//i.test(url)) continue;

    const extension = name.match(/\.([a-z0-9]+)$/i)?.[1]?.toLowerCase();
    if (!extension || !allowedExtensions.has(extension)) continue;
    if (!unique.has(url)) unique.set(url, { name, url });
  }

  const suggestions = Array.from(unique.values());
  const normalizedQuery = normalizeFileSearchText(props.query);
  const exactMatches = suggestions.filter((candidate) => {
    const normalizedName = normalizeFileSearchText(
      candidate.name.replace(/\.[^.]+$/i, "")
    );
    return (
      normalizedName.includes(normalizedQuery) ||
      normalizedQuery.includes(normalizedName)
    );
  });

  return {
    exactMatches,
    suggestions: suggestions.slice(0, 5),
  };
}

export function isTeamsSearchConfigured(): boolean {
  return [
    "AZURE_SEARCH_API_KEY",
    "AZURE_SEARCH_NAME",
    "AZURE_SEARCH_INDEX_NAME",
    "AZURE_OPENAI_API_KEY",
    "AZURE_OPENAI_API_INSTANCE_NAME",
    "AZURE_OPENAI_API_EMBEDDINGS_DEPLOYMENT_NAME",
  ].every((name) => Boolean(process.env[name]?.trim()));
}

function formatSearchResults(results: DocumentSearchResponse[]): {
  context: string;
  sources: TeamsSearchSource[];
} {
  const sourceByKey = new Map<string, TeamsSearchSource>();
  const sections: string[] = [];
  let totalCharacters = 0;

  for (const result of results) {
    const document = result.document;
    const name = document.metadata?.trim() || "名称不明の文書";
    const url = cleanUrl(document.fileUrl);
    const key = `${name}\n${url ?? ""}`;
    let source = sourceByKey.get(key);

    if (!source) {
      source = {
        index: sourceByKey.size + 1,
        name,
        kind: "internal",
        ...(url ? { url } : {}),
      };
      sourceByKey.set(key, source);
    }

    const pageContent = document.pageContent?.trim();
    if (!pageContent) continue;

    const section = `[${source.index}] ${name}\n${pageContent}`;
    const remaining = MAX_CONTEXT_CHARACTERS - totalCharacters;
    if (remaining <= 0) break;

    const clipped = section.slice(0, remaining);
    sections.push(clipped);
    totalCharacters += clipped.length;
  }

  return {
    context: sections.join("\n\n---\n\n"),
    sources: Array.from(sourceByKey.values()),
  };
}

function resolveTeamsSearchUserEmail(
  explicitEmail?: string | null
): string | null {
  const explicit = normalizeEmail(explicitEmail);
  if (explicit) return explicit;

  if (process.env.NODE_ENV !== "production") {
    return (
      normalizeEmail(process.env.TEAMS_TEST_USER_EMAIL) ??
      normalizeEmail(process.env.SL_LOCAL_DEFAULT_EMAIL)
    );
  }

  return null;
}

function resolveSearchTop(): number {
  const parsed = Number.parseInt(process.env.TEAMS_SEARCH_TOP ?? "", 10);
  if (!Number.isFinite(parsed)) return DEFAULT_SEARCH_TOP;
  return Math.min(Math.max(parsed, 1), 20);
}

function cleanUrl(value?: string | null): string | undefined {
  const normalized = value?.trim();
  if (!normalized || !/^https:\/\//i.test(normalized)) return undefined;
  return normalized;
}

function normalizeEmail(value?: string | null): string | null {
  const normalized = value?.trim().toLowerCase();
  return normalized || null;
}

function normalizeFileSearchText(value: string): string {
  return value
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\.(pdf|docx|xlsx)$/i, "")
    .replace(/[\s\u3000「」『』【】()（）・_\-]/g, "");
}

function requiredEnv(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`Missing environment variable: ${name}`);
  return value;
}

function assertSearchConfiguration(): void {
  if (!isTeamsSearchConfigured()) {
    throw new Error("Teams AI Search configuration is incomplete.");
  }
}
