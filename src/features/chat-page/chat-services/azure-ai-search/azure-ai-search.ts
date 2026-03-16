"use server";
import "server-only";

import { userHashedId } from "@/features/auth-page/helpers";
import { ServerActionResponse } from "@/features/common/server-action-response";
import {
  AzureAISearchIndexClientInstance,
  AzureAISearchInstance,
} from "@/features/common/services/ai-search";
import { OpenAIEmbeddingInstance } from "@/features/common/services/openai";
import { uniqueId } from "@/features/common/util";
import {
  AzureKeyCredential,
  SearchClient,
  SearchIndex,
} from "@azure/search-documents";

export interface AzureSearchDocumentIndex {
  id: string;
  pageContent: string;
  embedding?: number[];
  user: string;
  chatThreadId: string;
  metadata: string;
  fileUrl: string;
  dept: string;
  isSlDoc: boolean | null;
}

export type DocumentSearchResponse = {
  score: number;
  document: AzureSearchDocumentIndex;
};

function escapeODataValue(value: string): string {
  return String(value ?? "").replace(/'/g, "''");
}

function combineFilters(a?: string, b?: string): string | undefined {
  const aa = (a ?? "").trim();
  const bb = (b ?? "").trim();

  if (!aa) return bb || undefined;
  if (!bb) return aa || undefined;

  return `(${aa}) and (${bb})`;
}

/**
 * ACL統一関数
 * - 個人文書: isSlDoc != true かつ user一致
 * - SL文書  : isSlDoc == true かつ dept一致 or common
 *
 * 使い分け:
 * - deptLower === null      → ACLを付けない（明示的無効化）
 * - deptLower === undefined → fallbackで "others"
 * - userHash が渡された場合はそれを使う（Route Handler経由）
 * - userHash が未指定の場合は userHashedId() を呼ぶ（Server Action経由）
 */
async function buildSearchAclFilter(
  deptLower?: string | null,
  userHash?: string
): Promise<string | undefined> {
  if (deptLower === null) return undefined;

  const normalizedDept = (deptLower ?? "others").toLowerCase().trim();
  const d = escapeODataValue(normalizedDept);

  const resolvedUserHash = userHash ?? (await userHashedId());

  const userFilter = `(isSlDoc ne true and user eq '${escapeODataValue(resolvedUserHash)}')`;
  const deptFilter = `(isSlDoc eq true and (dept eq '${d}' or dept eq 'common'))`;

  return `(${userFilter} or ${deptFilter})`;
}

// -------------------------------------------------------
// Search
// -------------------------------------------------------

export const SimpleSearch = async (
  searchText?: string,
  filter?: string,
  deptLower?: string | null
): Promise<ServerActionResponse<Array<DocumentSearchResponse>>> => {
  try {
    const instance = AzureAISearchInstance<AzureSearchDocumentIndex>();

    const scopeFilter = await buildSearchAclFilter(deptLower);
    const finalFilter = combineFilters(filter, scopeFilter);

    const searchResults = await instance.search(searchText ?? "*", {
      filter: finalFilter,
    });

    const results: Array<DocumentSearchResponse> = [];

    for await (const result of searchResults.results) {
      results.push({
        score: result.score,
        document: result.document,
      });
    }

    return { status: "OK", response: results };
  } catch (e) {
    return { status: "ERROR", errors: [{ message: `${e}` }] };
  }
};

export const SimilaritySearch = async (
  searchText: string,
  k: number,
  filter?: string,
  deptLower?: string | null
): Promise<ServerActionResponse<Array<DocumentSearchResponse>>> => {
  try {
    const openai = OpenAIEmbeddingInstance();

    const embeddings = await openai.embeddings.create({
      input: searchText,
      model: "",
    });

    const searchClient = AzureAISearchInstance<AzureSearchDocumentIndex>();

    const scopeFilter = await buildSearchAclFilter(deptLower);
    const finalFilter = combineFilters(filter, scopeFilter);

    const searchResults = await searchClient.search(searchText, {
      top: k,
      filter: finalFilter,
      vectorSearchOptions: {
        queries: [
          {
            vector: embeddings.data[0].embedding,
            fields: ["embedding"],
            kind: "vector",
            kNearestNeighborsCount: 10,
          },
        ],
      },
    });

    const results: Array<DocumentSearchResponse> = [];

    for await (const result of searchResults.results) {
      results.push({
        score: result.score,
        document: result.document,
      });
    }

    return { status: "OK", response: results };
  } catch (e) {
    return { status: "ERROR", errors: [{ message: `${e}` }] };
  }
};

export const ExtensionSimilaritySearch = async (props: {
  searchText: string;
  vectors: string[];
  apiKey: string;
  searchName: string;
  indexName: string;
  filter?: string;
  deptLower?: string | null;
  userHash?: string; // Route Handler経由で渡すuserHash
}): Promise<ServerActionResponse<Array<DocumentSearchResponse>>> => {
  try {
    const openai = OpenAIEmbeddingInstance();

    const {
      searchText,
      vectors,
      apiKey,
      searchName,
      indexName,
      filter,
      deptLower,
      userHash,
    } = props;

    const embeddings = await openai.embeddings.create({
      input: searchText,
      model: "",
    });

    const endpoint = `https://${searchName}.search.windows.net`;

    const searchClient = new SearchClient(
      endpoint,
      indexName,
      new AzureKeyCredential(apiKey),
      { allowInsecureConnection: process.env.NODE_ENV === "development" }
    );

    // userHash を外から受け取り buildSearchAclFilter に渡す
    const scopeFilter = await buildSearchAclFilter(deptLower, userHash);
    const finalFilter = combineFilters(filter, scopeFilter);

    console.log("[SEARCH:Extension] deptLower =", deptLower);
    console.log("[SEARCH:Extension] userHash =", userHash ? "***" : "(none)");
    console.log("[SEARCH:Extension] finalFilter =", finalFilter);

    const searchResults = await searchClient.search(searchText, {
      top: 3,
      filter: finalFilter,
      vectorSearchOptions: {
        queries: [
          {
            vector: embeddings.data[0].embedding,
            fields: vectors,
            kind: "vector",
            kNearestNeighborsCount: 10,
          },
        ],
      },
    });

    const results: Array<DocumentSearchResponse> = [];

    for await (const result of searchResults.results) {
      const document = result.document as Record<string, unknown>;
      const newDocument: Record<string, unknown> = {};

      for (const key in document) {
        if (!vectors.includes(key)) {
          newDocument[key] = document[key];
        }
      }

      results.push({
        score: result.score,
        document: newDocument as unknown as AzureSearchDocumentIndex,
      });
    }

    return { status: "OK", response: results };
  } catch (e) {
    return { status: "ERROR", errors: [{ message: `${e}` }] };
  }
};

// -------------------------------------------------------
// Indexing
// -------------------------------------------------------

export const IndexDocuments = async (
  fileName: string,
  fileUrl: string,
  docs: string[],
  chatThreadId: string,
  dept: string,
  isSlDoc: boolean
): Promise<Array<ServerActionResponse<boolean>>> => {
  try {
    const documentsToIndex: AzureSearchDocumentIndex[] = [];
    const currentUserHash = await userHashedId();

    for (const doc of docs) {
      documentsToIndex.push({
        id: uniqueId(),
        chatThreadId,
        user: isSlDoc ? "" : currentUserHash,
        pageContent: doc,
        metadata: fileName,
        fileUrl,
        embedding: [],
        dept: (dept ?? "others").toLowerCase().trim(),
        isSlDoc,
      });
    }

    const instance = AzureAISearchInstance<AzureSearchDocumentIndex>();

    const embeddingsResponse = await EmbedDocuments(documentsToIndex);

    if (embeddingsResponse.status !== "OK") {
      return [embeddingsResponse];
    }

    const uploadResponse = await instance.uploadDocuments(
      embeddingsResponse.response
    );

    const response: Array<ServerActionResponse<boolean>> = [];

    uploadResponse.results.forEach((r) => {
      if (r.succeeded) {
        response.push({ status: "OK", response: true });
      } else {
        response.push({
          status: "ERROR",
          errors: [{ message: `${r.errorMessage}` }],
        });
      }
    });

    return response;
  } catch (e) {
    return [{ status: "ERROR", errors: [{ message: `${e}` }] }];
  }
};

// -------------------------------------------------------
// Embed
// -------------------------------------------------------

export const EmbedDocuments = async (
  documents: AzureSearchDocumentIndex[]
): Promise<ServerActionResponse<Array<AzureSearchDocumentIndex>>> => {
  try {
    const openai = OpenAIEmbeddingInstance();

    const embeddings = await openai.embeddings.create({
      input: documents.map((d) => d.pageContent),
      model: "",
    });

    const embeddedDocuments = documents.map((doc, index) => ({
      ...doc,
      embedding: embeddings.data[index]?.embedding ?? [],
    }));

    return {
      status: "OK",
      response: embeddedDocuments,
    };
  } catch (e) {
    return {
      status: "ERROR",
      errors: [{ message: `${e}` }],
    };
  }
};

// -------------------------------------------------------
// Index helpers
// -------------------------------------------------------

export const GetSearchIndex = async (
  indexName: string
): Promise<ServerActionResponse<SearchIndex>> => {
  try {
    const client = AzureAISearchIndexClientInstance();
    const index = await client.getIndex(indexName);
    return { status: "OK", response: index };
  } catch (e) {
    return { status: "ERROR", errors: [{ message: `${e}` }] };
  }
};

export const DeleteDocuments = async (
  chatThreadId: string
): Promise<Array<ServerActionResponse<boolean>>> => {
  try {
    const safeChatThreadId = escapeODataValue(chatThreadId);

    const documentsInChatResponse = await SimpleSearch(
      undefined,
      `chatThreadId eq '${safeChatThreadId}'`,
      null
    );

    if (documentsInChatResponse.status !== "OK") {
      return [
        {
          status: "ERROR",
          errors: documentsInChatResponse.errors ?? [
            { message: "Failed to search documents before delete." },
          ],
        },
      ];
    }

    const instance = AzureAISearchInstance<AzureSearchDocumentIndex>();
    const deletedResponse = await instance.deleteDocuments(
      documentsInChatResponse.response.map((r) => r.document)
    );

    const response: Array<ServerActionResponse<boolean>> = [];
    deletedResponse.results.forEach((r) => {
      if (r.succeeded) {
        response.push({ status: "OK", response: true });
      } else {
        response.push({
          status: "ERROR",
          errors: [{ message: `${r.errorMessage}` }],
        });
      }
    });

    return response;
  } catch (e) {
    return [{ status: "ERROR", errors: [{ message: `${e}` }] }];
  }
};

export const EnsureIndexIsCreated = async (): Promise<
  ServerActionResponse<boolean>
> => {
  try {
    await AzureAISearchIndexClientInstance().getIndex(
      process.env.AZURE_SEARCH_INDEX_NAME!
    );
    return { status: "OK", response: true };
  } catch {
    return { status: "OK", response: false };
  }
};