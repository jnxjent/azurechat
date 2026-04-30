import { ExtensionSimilaritySearch } from "../azure-ai-search/azure-ai-search";
import { CreateCitations, FormatCitations } from "../citation-service";

function getRequiredHeader(req: Request, name: string): string {
  const value = req.headers.get(name)?.trim();
  if (!value) {
    throw new Error(`[EXT-RAG] Missing required header: ${name}`);
  }
  return value;
}

export const SearchAzureAISimilarDocuments = async (
  req: Request,
  deptLower: string,
  userHash: string | null  // route.ts から受け取る
) => {
  try {
    const body = await req.json();
    const search = String(body.search ?? "").trim();

    if (!search) {
      return new Response(
        JSON.stringify({
          status: "ERROR",
          errors: [{ message: "Missing search text." }],
        }),
        {
          headers: { "Content-Type": "application/json; charset=utf-8" },
          status: 400,
        }
      );
    }

    if (!userHash) {
      return new Response(
        JSON.stringify({
          status: "ERROR",
          errors: [{ message: "User not authenticated." }],
        }),
        {
          headers: { "Content-Type": "application/json; charset=utf-8" },
          status: 401,
        }
      );
    }

    const vectorsHeader = getRequiredHeader(req, "vectors");
    const apiKey = getRequiredHeader(req, "apiKey");
    const searchName = getRequiredHeader(req, "searchName");
    const indexName = getRequiredHeader(req, "indexName");

    const vectors = vectorsHeader
      .split(",")
      .map((v) => v.trim())
      .filter(Boolean);

    console.log("[EXT-RAG] deptLower =", deptLower);
    console.log("[EXT-RAG] userHash =", userHash ? "***" : "(none)");
    console.log("[EXT-RAG] searchName =", searchName);
    console.log("[EXT-RAG] indexName =", indexName);
    console.log("[EXT-RAG] vectors =", vectors);

    const result = await ExtensionSimilaritySearch({
      apiKey,
      searchName,
      indexName,
      vectors,
      searchText: search,
      deptLower,
      userHash,
    });

    if (result.status !== "OK") {
      console.error("🔴 Retrieving documents", result.errors);

      return new Response(JSON.stringify(result), {
        headers: { "Content-Type": "application/json; charset=utf-8" },
        status: 500,
      });
    }

    const withoutEmbedding = FormatCitations(result.response);

    const citationResponse = await CreateCitations(withoutEmbedding, userHash);

    const allCitations = [];

    for (const citation of citationResponse) {
      if (citation.status === "OK") {
        allCitations.push({
          id: citation.response.id,
          content: citation.response.content,
        });
      }
    }

    return new Response(JSON.stringify(allCitations), {
      headers: { "Content-Type": "application/json; charset=utf-8" },
    });
  } catch (e) {
    console.error("🔴 Retrieving documents", e);

    return new Response(
      JSON.stringify({
        error: true,
        message: e instanceof Error ? e.message : String(e),
      }),
      {
        headers: { "Content-Type": "application/json; charset=utf-8" },
        status: 500,
      }
    );
  }
};