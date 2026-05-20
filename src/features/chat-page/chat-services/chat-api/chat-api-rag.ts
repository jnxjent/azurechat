"use server";
import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";
import { ChatCompletionStreamingRunner } from "openai/resources/beta/chat/completions";
import { ChatCompletionMessageParam } from "openai/resources/chat/completions";
import { ExtensionSimilaritySearch } from "../azure-ai-search/azure-ai-search";
import { CreateCitations, FormatCitations } from "../citation-service";
import { ChatCitationModel, ChatThreadModel } from "../models";
import { GetDefaultExtensions } from "./chat-api-default-extensions";
import { FindAllChatDocuments } from "../chat-document-service";
import { GenerateSasUrl } from "@/features/common/services/azure-storage";

// dept判定ユーティリティ
import { decideDept, getEffectiveSlUserEmail, getUserEmailFromJwtToken, resolveSlAccess } from "@/lib/sl-dept";
import { getToken } from "next-auth/jwt";
import { cookies } from "next/headers";
import { hashValue, userSession } from "@/features/auth-page/helpers";

// OData filter用にシングルクォートをエスケープ
function odataEscape(v: string) {
  return String(v ?? "").replace(/'/g, "''");
}

type UserContext = {
  email: string | null;
  deptLower: string;
  userHash: string | null;
};

/**
 * サーバ側で「ユーザーの email / deptLower / userHash」を決める
 * - token から email を抜く
 * - email → dept 判定（sl-dept.ts）
 * - email → hashValue(email)
 * - fallback は SL_DEPT_DEFAULT
 */
export async function resolveUserContext(): Promise<UserContext> {
  try {
    const session = await userSession().catch(() => null);
    const cookieStore = await cookies();

    const token = await getToken({
      req: {
        headers: {
          cookie: cookieStore.toString(),
        },
        cookies: Object.fromEntries(
          cookieStore.getAll().map((c) => [c.name, c.value])
        ),
      } as any,
      secret: process.env.NEXTAUTH_SECRET!,
    }).catch(() => null);

    const email = getEffectiveSlUserEmail(
      token ? getUserEmailFromJwtToken(token) : null
    );

    const deptLower = email
      ? resolveSlAccess(email).dept
      : session?.slDept?.trim().toLowerCase() ||
        decideDept({
          requestedDept: undefined,
          userEmail: email,
        });

    const userHash = email ? hashValue(email) : null;

    console.log("[RAG-EXT:resolveUserContext] email =", email);
    console.log("[RAG-EXT:resolveUserContext] deptLower =", deptLower);
    console.log(
      "[RAG-EXT:resolveUserContext] userHash =",
      userHash ? "***" : "(none)"
    );

    return {
      email,
      deptLower,
      userHash,
    };
  } catch {
    const deptLower =
      (process.env.SL_DEPT_DEFAULT ?? "cp").toLowerCase().trim() || "cp";

    console.log("[RAG-EXT:resolveUserContext] fallback email = (none)");
    console.log("[RAG-EXT:resolveUserContext] fallback deptLower =", deptLower);
    console.log("[RAG-EXT:resolveUserContext] fallback userHash = (none)");

    return {
      email: null,
      deptLower,
      userHash: null,
    };
  }
}

function getRequiredEnv(name: string): string {
  const value = process.env[name];
  if (!value || !value.trim()) {
    throw new Error(`[RAG-EXT] Missing environment variable: ${name}`);
  }
  return value.trim();
}

async function resolveThreadDocumentBlobUrls(chatThreadId: string): Promise<string[]> {
  const docsResponse = await FindAllChatDocuments(chatThreadId);
  if (docsResponse.status !== "OK") {
    return [];
  }

  const urls = await Promise.all(
    docsResponse.response.map(async (doc) => {
      const sas = await GenerateSasUrl("dl-link", `${chatThreadId}/${doc.name}`);
      return sas.status === "OK" ? sas.response : null;
    })
  );

  return Array.from(new Set(urls.filter((url): url is string => Boolean(url))));
}

export const ChatApiRAG = async (props: {
  chatThread: ChatThreadModel;
  userMessage: string;
  history: ChatCompletionMessageParam[];
  signal: AbortSignal;
}): Promise<ChatCompletionStreamingRunner> => {
  const { chatThread, userMessage, history, signal } = props;

  const openAI = OpenAIInstance();
  const { email, deptLower, userHash } = await resolveUserContext();

  console.log("[RAG-EXT] email =", email);
  console.log("[RAG-EXT] deptLower =", deptLower);
  console.log("[RAG-EXT] userHash =", userHash ? "***" : "(none)");

  const filter = `(chatThreadId eq '${odataEscape(chatThread.id)}' or isSlDoc eq true)`;
  console.log("[RAG-EXT] base filter =", filter);

  const apiKey = getRequiredEnv("AZURE_SEARCH_API_KEY");
  const searchName = getRequiredEnv("AZURE_SEARCH_NAME");
  const indexName = getRequiredEnv("AZURE_SEARCH_INDEX_NAME");

  const documentResponse = await ExtensionSimilaritySearch({
    searchText: userMessage,
    vectors: ["embedding"],
    apiKey,
    searchName,
    indexName,
    filter,
    deptLower,
    userHash: userHash ?? undefined,
  });

  const documents: ChatCitationModel[] = [];
  const uploadedBlobUrls = await resolveThreadDocumentBlobUrls(chatThread.id);

  if (documentResponse.status === "OK") {
    const withoutEmbedding = FormatCitations(documentResponse.response);
    const citationResponse = await CreateCitations(withoutEmbedding);
    citationResponse.forEach((c) => {
      if (c.status === "OK") {
        documents.push(c.response);
      }
    });
  } else {
    console.error("[RAG-EXT] ExtensionSimilaritySearch error:", documentResponse.errors);
  }

  const content = documents
    .map((result, index) => {
      const page = result.content.document.pageContent;
      const displayUrl =
        result.content.document.effectiveFileUrl ??
        result.content.document.fileUrl ??
        "";
      // このスレッドにアップロードされたファイルのみ file_url を出す
      // 他スレッド由来のSLドキュメントは認証が必要なため除外（convert_doc_to_pptxで誤使用防止）
      // SLファイルは fileUrl=SP webUrl / effectiveFileUrl=Blob URL なので、Blob URLを優先する
      const isThisThread = result.content.document.chatThreadId === chatThread.id;
      const blobUrl = isThisThread
        ? (result.content.document.effectiveFileUrl ?? result.content.document.fileUrl ?? displayUrl)
        : null;
      return `[${index}]. file name: ${result.content.document.metadata}
file id: ${result.id}${blobUrl ? `\nfile_url: ${blobUrl}` : ""}
${page}`;
    })
    .join("\n------\n");

  // ファイルURLリスト（convert_doc_to_pptx ツールに渡すため）
  // このスレッドにアップロードされたファイルのみを対象にする
  const fileUrls = uploadedBlobUrls;
  const hasUploadedFile = fileUrls.length > 0;

  const xlsxUrls = fileUrls.filter((u) => /\.(xlsx|xls|xlsm)(?:\?|$)/i.test(u));
  const fileUrlHint = hasUploadedFile
    ? `\n- The uploaded document file URLs are:\n${fileUrls.map((u, i) => `  [${i}] ${u}`).join("\n")}\n- If the user asks to convert the document to PowerPoint, use the convert_doc_to_pptx tool with the file_url from above.${xlsxUrls.length > 0 ? `\n- CRITICAL: このスレッドにExcelファイル（.xlsx）がアップロードされています。ユーザーが「グラフにして」「折れ線グラフ」「棒グラフ」「グラフ化して」「チャートを作成して」と言った場合、必ず edit_excel ツールを fileUrl=${xlsxUrls[0]} で呼び出すこと。検索結果にPNGファイルが含まれていても、それはExcelとは無関係の知識ベースの画像であり、ユーザーのExcelファイルではない。` : ""}`
    : "\n- 【重要】このスレッドにアップロードされたファイルは存在しません。ユーザーがSharePoint/SLの資料名を挙げてPPT変換を要求した場合は、必ず convert_sp_to_pptx ツールを使うこと。convert_doc_to_pptx は使わないこと。";

  const _userMessage = `
- Review the following content from documents uploaded by the user and create a final answer.
- If you don't know the answer, just say that you don't know. Don't try to make up an answer.
- You must always include a citation at the end of your answer and don't include full stop after the citations.
- IMPORTANT: If the user asks to compare multiple documents or find contradictions across files: (1) First call sl_doc_search with a broad query (e.g. "IR議事録") to discover available document names. (2) Then call sl_doc_search once per discovered document using "company name + document type + keyword" queries. (3) Only answer after collecting content from all documents. Never answer based solely on the initial context when multi-document comparison is requested.
- If the user asks to create a PowerPoint or slides from the document content, use the convert_doc_to_pptx tool with the file_url from the document context below. This tool uses Vision API for high-quality conversion.${fileUrlHint}

----------------
content:
${content}

----------------
question:
${userMessage}
`;

  // ★ デフォルトツール（create_pptx 等）を RAG モードでも有効にする
  const extensionsResponse = await GetDefaultExtensions({
    chatThread,
    userMessage,
    signal,
  });
  const tools = extensionsResponse.status === "OK" ? extensionsResponse.response : [];

  // ★ sl_doc_search ツール：LLMが複数文書を横断検索するために複数回呼び出し可能
  tools.push({
    type: "function",
    function: {
      name: "sl_doc_search",
      description:
        "SharePointの個人・部署・全社共通ドキュメントを検索します。\n" +
        "【2段階で使うこと】\n" +
        "① 比較対象の文書名が不明な場合: mode=\"discover\" で広いクエリ（例:「IR議事録」）を1回呼び出し、返ってくる file name から文書名・会社名を把握する。\n" +
        "② 文書名が判明したら: mode=\"content\" で「会社名 + 文書種別 + キーワード」の形式で文書ごとに個別呼び出しする（複数回）。\n" +
        "例：最初に mode=discover で「IR議事録」→ 次に mode=content で「野村アセット IR議事録 社長コメント」「セイタキャピタル IR議事録 社長コメント」と個別検索。",
      parameters: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description:
              "検索クエリ。会社名・ファイル名・キーワードを組み合わせると精度が上がります。例：「セイタキャピタル IR議事録 社長コメント」",
          },
          mode: {
            type: "string",
            enum: ["discover", "content"],
            description:
              "discover: 文書名の一覧取得（広いクエリ向け、上位32件）。content: 個社別の本文取得（絞ったクエリ向け、上位8件）。省略時はcontent扱い。",
          },
        },
        required: ["query"],
      },
      function: async (args: { query: string; mode?: string }) => {
        const effectiveTop = args.mode === "discover" ? 32 : 8;
        console.log("[sl_doc_search] query =", args.query, "mode =", args.mode ?? "content", "top =", effectiveTop);
        const searchResult = await ExtensionSimilaritySearch({
          searchText: args.query,
          vectors: ["embedding"],
          apiKey,
          searchName,
          indexName,
          filter,
          deptLower,
          userHash: userHash ?? undefined,
          top: effectiveTop,
        });

        if (searchResult.status !== "OK") {
          console.error("[sl_doc_search] error:", searchResult.errors);
          return "検索エラーが発生しました";
        }

        if (searchResult.response.length === 0) {
          return "該当する文書が見つかりませんでした";
        }

        const withoutEmbedding = FormatCitations(searchResult.response);
        const citationResponse = await CreateCitations(withoutEmbedding);

        return searchResult.response
          .map((r, i) => {
            const cit = citationResponse[i];
            const id = cit?.status === "OK" ? cit.response.id : r.document.id;
            return `[${i}]. file name: ${r.document.metadata}\nfile id: ${id}\n${r.document.pageContent}`;
          })
          .join("\n---\n");
      },
      parse: (input: string) => JSON.parse(input),
    },
  });

  if (tools.length > 0) {
    return openAI.beta.chat.completions.runTools(
      {
        model: "",
        stream: true,
        messages: [
          { role: "system", content: chatThread.personaMessage },
          ...history,
          { role: "user", content: _userMessage },
        ],
        tools,
      },
      { signal }
    );
  }

  // ツールなしフォールバック
  return openAI.beta.chat.completions.stream(
    {
      model: "",
      stream: true,
      messages: [
        { role: "system", content: chatThread.personaMessage },
        ...history,
        { role: "user", content: _userMessage },
      ],
    },
    { signal }
  );
};
