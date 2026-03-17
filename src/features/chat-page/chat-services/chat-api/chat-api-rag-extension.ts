import { getToken } from "next-auth/jwt";
import { decideDept, getUserEmailFromJwtToken } from "@/lib/sl-dept";
import { ExtensionSimilaritySearch } from "../azure-ai-search/azure-ai-search";
import { CreateCitations, FormatCitations } from "../citation-service";

/**
 * Request から deptLower を解決
 * - next-auth token から email を取得
 * - email -> decideDept()
 * - fallback は SL_DEPT_DEFAULT
 */
async function resolveDeptLowerFromRequest(req: Request): Promise<string> {
  try {
    const cookieHeader = req.headers.get("cookie") ?? "";

    const token = await getToken({
      req: {
        headers: {
          cookie: cookieHeader,
        },
        cookies: Object.fromEntries(
          cookieHeader
            .split(";")
            .map((part) => part.trim())
            .filter(Boolean)
            .map((part) => {
              const eq = part.indexOf("=");
              if (eq < 0) return [part, ""];
              const key = part.slice(0, eq);
              const value = part.slice(eq + 1);
              return [key, value];
            })
        ),
      } as any,
      secret: process.env.NEXTAUTH_SECRET!,
    }).catch(() => null);

    const email = token ? getUserEmailFromJwtToken(token) : null;

    return decideDept({
      requestedDept: undefined,
      userEmail: email,
    });

  } catch {

    return (
      (process.env.SL_DEPT_DEFAULT ?? "others")
        .toLowerCase()
        .trim() || "others"
    );
  }
}

export const SearchAzureAISimilarDocuments = async (req: Request) => {

  try {

    const body = await req.json();
    const search = body.search as string;

    const vectors = req.headers.get("vectors") as string;
    const apiKey = req.headers.get("apiKey") as string;
    const searchName = req.headers.get("searchName") as string;
    const indexName = req.headers.get("indexName") as string;

    // authorization header には userHash が入る想定
    const userId = (req.headers.get("authorization") as string) ?? "";

    const deptLower = await resolveDeptLowerFromRequest(req);

    console.log("[EXT-RAG] deptLower =", deptLower);

    const result = await ExtensionSimilaritySearch({
      apiKey,
      searchName,
      indexName,
      vectors: vectors.split(","),
      searchText: search,
      deptLower,
      userHash: userId || undefined,  // ★追加
    });

    if (result.status !== "OK") {

      console.error("🔴 Retrieving documents", result.errors);

      return new Response(JSON.stringify(result));
    }

    const withoutEmbedding = FormatCitations(result.response);

    const citationResponse = await CreateCitations(
      withoutEmbedding,
      userId
    );

    const allCitations = [];

    for (const citation of citationResponse) {

      if (citation.status === "OK") {

        allCitations.push({
          id: citation.response.id,
          content: citation.response.content,
        });
      }
    }

    return new Response(JSON.stringify(allCitations));

  } catch (e) {

    console.error("🔴 Retrieving documents", e);

    return new Response(JSON.stringify(e));
  }
};