import "server-only";

import {
  ExtensionSimilaritySearch,
  SimpleSearch,
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
    console.warn(
      "[teams-search] skipped because the channel did not provide a resolvable user"
    );
    return { context: "", sources: [], userEmail: "", dept: "" };
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

export async function searchTeamsKnowledgeByFileFamily(props: {
  query: string;
  searchQueries: string[];
  fileFamily: string;
  userEmail?: string | null;
}): Promise<TeamsSearchResult> {
  assertSearchConfiguration();

  const userEmail = resolveTeamsSearchUserEmail(props.userEmail);
  if (!userEmail) {
    return { context: "", sources: [], userEmail: "", dept: "" };
  }

  const access = resolveSlAccess(userEmail);
  const queries = Array.from(
    new Set(props.searchQueries.map((item) => item.trim()).filter(Boolean))
  );
  const responses = await Promise.all(
    (queries.length ? queries : [props.query]).map((searchText) =>
      ExtensionSimilaritySearch({
        searchText,
        vectors: ["embedding"],
        apiKey: requiredEnv("AZURE_SEARCH_API_KEY"),
        searchName: requiredEnv("AZURE_SEARCH_NAME"),
        indexName: requiredEnv("AZURE_SEARCH_INDEX_NAME"),
        filter: "isSlDoc eq true",
        deptLower: access.dept,
        userHash: hashValue(userEmail),
        top: 20,
      })
    )
  );

  const failed = responses.find((response) => response.status !== "OK");
  if (failed) {
    const detail = failed.errors?.map((error) => error.message).join("; ");
    throw new Error(
      `Teams targeted AI Search failed${detail ? `: ${detail}` : "."}`
    );
  }

  const normalizedFamily = normalizeFileSearchText(props.fileFamily);
  const unique = new Map<string, DocumentSearchResponse>();
  for (const response of responses) {
    if (response.status !== "OK") continue;
    for (const item of response.response) {
      const name = item.document.metadata?.trim() ?? "";
      if (
        !normalizedFamily ||
        !normalizeFileSearchText(name).includes(normalizedFamily)
      ) {
        continue;
      }
      const key = [
        item.document.effectiveFileUrl ?? item.document.fileUrl ?? name,
        item.document.pageContent ?? "",
      ].join("\n");
      if (!unique.has(key)) unique.set(key, item);
    }
  }

  const formatted = formatSearchResults(Array.from(unique.values()));
  console.log("[teams-search] targeted document search completed", {
    dept: access.dept,
    fileFamily: props.fileFamily,
    queries: queries.length,
    matchedChunks: unique.size,
    sources: formatted.sources.length,
  });
  return { ...formatted, userEmail, dept: access.dept };
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
  const userHash = hashValue(userEmail);
  const filenameList = await SimpleSearch(
    "*",
    "isSlDoc eq true",
    access.dept,
    1000,
    userHash
  );
  const response =
    filenameList.status === "OK" && filenameList.response.length > 0
      ? filenameList
      : await ExtensionSimilaritySearch({
    searchText: props.query,
    vectors: ["embedding"],
    apiKey: requiredEnv("AZURE_SEARCH_API_KEY"),
    searchName: requiredEnv("AZURE_SEARCH_NAME"),
    indexName: requiredEnv("AZURE_SEARCH_INDEX_NAME"),
    filter: "isSlDoc eq true",
    deptLower: access.dept,
    userHash,
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

  console.log("[teams-office-search] filename-first completed", {
    dept: access.dept,
    query: props.query,
    scannedDocuments:
      response.status === "OK" ? response.response.length : 0,
    candidateFiles: suggestions.length,
    exactMatches: exactMatches.length,
  });

  return {
    exactMatches,
    suggestions: suggestions.slice(0, 5),
  };
}

export async function searchTeamsKnowledgeByOfficeFileName(props: {
  fileName: string;
  userEmail?: string | null;
}): Promise<TeamsSearchResult> {
  assertSearchConfiguration();

  const userEmail = resolveTeamsSearchUserEmail(props.userEmail);
  if (!userEmail) {
    throw new Error(
      "Teams Office user is unresolved. Set TEAMS_TEST_USER_EMAIL for local Playground testing."
    );
  }

  const access = resolveSlAccess(userEmail);
  const response = await SimpleSearch(
    "*",
    "isSlDoc eq true",
    access.dept,
    1000,
    hashValue(userEmail)
  );
  if (response.status !== "OK") {
    const detail = response.errors?.map((error) => error.message).join("; ");
    throw new Error(
      `Teams Office document read failed${detail ? `: ${detail}` : "."}`
    );
  }

  const expected = normalizeFileSearchText(props.fileName);
  const matches = response.response.filter(({ document }) => {
    const metadataName = document.metadata?.trim() ?? "";
    return normalizeFileSearchText(metadataName) === expected;
  });
  const formatted = formatSearchResults(matches);
  console.log("[teams-office-search] exact file content completed", {
    dept: access.dept,
    fileName: props.fileName,
    chunks: matches.length,
    contextCharacters: formatted.context.length,
  });

  return { ...formatted, userEmail, dept: access.dept };
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
  if (process.env.NODE_ENV !== "production") {
    const localTestEmail =
      normalizeEmail(process.env.TEAMS_TEST_USER_EMAIL) ??
      normalizeEmail(process.env.SL_LOCAL_DEFAULT_EMAIL);
    if (localTestEmail) return localTestEmail;
  }

  return normalizeEmail(explicitEmail);
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
