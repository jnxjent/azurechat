"use server";
import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";
import { ChatCompletionStreamingRunner } from "openai/resources/beta/chat/completions";
import { ChatCompletionMessageParam } from "openai/resources/chat/completions";
import { ExtensionSimilaritySearch } from "../azure-ai-search/azure-ai-search";
import {
  buildSlSearchTargetFilter,
  inferSlSearchTarget,
  stripSlSearchTargetTerms,
} from "@/lib/sl-search-target";
import type { SlSearchScope } from "@/lib/sl-search-target";
import { CreateCitations, FormatCitations } from "../citation-service";
import { ChatCitationModel, ChatThreadModel } from "../models";
import { GetDefaultExtensions } from "./chat-api-default-extensions";
import { FindAllChatDocuments } from "../chat-document-service";
import { GenerateSasUrl } from "@/features/common/services/azure-storage";
import { createSharePointPdfSummaryTool } from "../sharepoint-summary-tool";
import { createDeskNetsAgentTool } from "@/features/desknets-agent/desknets-agent-tool";
import { isDeskNetsAgentEnabled } from "@/features/desknets-agent/desknets-agent-client";
import {
  shouldRouteToDeskNetsAgent,
} from "@/features/desknets-agent/desknets-agent-intent";

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
  const forceDeskNetsAgent =
    isDeskNetsAgentEnabled() &&
    shouldRouteToDeskNetsAgent(userMessage, history);

  const openAI = OpenAIInstance();
  const { email, deptLower, userHash } = await resolveUserContext();

  console.log("[RAG-EXT] email =", email);
  console.log("[RAG-EXT] deptLower =", deptLower);
  console.log("[RAG-EXT] userHash =", userHash ? "***" : "(none)");

  const baseFilter = `(chatThreadId eq '${odataEscape(chatThread.id)}' or isSlDoc eq true)`;
  const inferredTarget = inferSlSearchTarget(userMessage);
  const initialTargetFilter = buildSlSearchTargetFilter(inferredTarget);
  const initialSearchText = stripSlSearchTargetTerms(userMessage, inferredTarget);
  const filter = initialTargetFilter
    ? `(${baseFilter}) and (${initialTargetFilter})`
    : baseFilter;
  console.log("[RAG-EXT] base filter =", baseFilter);
  console.log("[RAG-EXT] inferred search target =", inferredTarget);
  console.log("[RAG-EXT] effective search text =", initialSearchText);
  console.log("[RAG-EXT] effective filter =", filter);

  const apiKey = getRequiredEnv("AZURE_SEARCH_API_KEY");
  const searchName = getRequiredEnv("AZURE_SEARCH_NAME");
  const indexName = getRequiredEnv("AZURE_SEARCH_INDEX_NAME");

  const documentResponse = inferredTarget.folderUncertain
    ? null
    : await ExtensionSimilaritySearch({
        searchText: initialSearchText,
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

  if (documentResponse?.status === "OK") {
    const withoutEmbedding = FormatCitations(documentResponse.response);
    const citationResponse = await CreateCitations(withoutEmbedding);
    citationResponse.forEach((c) => {
      if (c.status === "OK") {
        documents.push(c.response);
      }
    });
  } else if (documentResponse) {
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
    ? `\n- The uploaded document file URLs are:\n${fileUrls.map((u, i) => `  [${i}] ${u}`).join("\n")}\n- PDFをそのままPPTに変換する場合は convert_doc_to_pptx を上記 file_url で呼ぶこと。ただし、会話でスライド構成が既に議論済みでPDFは参考資料として使う場合は create_pptx を使うこと（convert_doc_to_pptx は使わないこと）。${xlsxUrls.length > 0 ? `\n- CRITICAL: このスレッドにExcelファイル（.xlsx）がアップロードされています。ユーザーが「グラフにして」「折れ線グラフ」「棒グラフ」「グラフ化して」「チャートを作成して」と言った場合、必ず edit_excel ツールを fileUrl=${xlsxUrls[0]} で呼び出すこと。検索結果にPNGファイルが含まれていても、それはExcelとは無関係の知識ベースの画像であり、ユーザーのExcelファイルではない。` : ""}`
    : "\n- 【重要】このスレッドにアップロードされたファイルは存在しません。ユーザーがSharePoint/SLの資料名を挙げてPPT変換を要求した場合は、必ず convert_sp_to_pptx ツールを使うこと。convert_doc_to_pptx は使わないこと。";

  const _userMessage = `
- Review the following content from documents uploaded by the user and create a final answer.
- If you don't know the answer, just say that you don't know. Don't try to make up an answer.
- You must always include a citation at the end of your answer and don't include full stop after the citations.
- OVERRIDE FOR PDF SUMMARY: If the user asks to summarize one named SharePoint PDF (including 要約, 全文要約, 全体要約, or 全ページ要約), call summarize_sp_pdf instead of sl_doc_search unless they explicitly request only a specific topic or page range. Preserve every [p.N] page reference returned by the tool and do not claim full-document coverage from ordinary search results.
- PDF SUMMARY LINK RULE: Use only the exact citation tag returned by summarize_sp_pdf. Never output a raw SharePoint URL or add a separate Markdown link such as '原文を開く'.
- PDF SUMMARY TO WORD: If the request also asks for Word/docx output, complete both calls in order: summarize_sp_pdf, then create_word. Pass an explicitly requested page count to summarize_sp_pdf.targetPages. Copy the short Word出力用summaryRef returned by summarize_sp_pdf into create_word.summaryRef, set create_word.content='[summaryRef]', create_word.formatMode='markdown', create_word.title='元ファイル名 要約', and create_word.fileName to the exact recommended '元ファイル名_要約.docx'. Never copy or rewrite the long summary into create_word.content and do not use convert_pdf_to_word for this case.
- IMPORTANT: If the user asks to compare multiple documents or find contradictions across files: (1) First call sl_doc_search with a broad query (e.g. "IR議事録") to discover available document names. (2) Then call sl_doc_search once per discovered document using "company name + document type + keyword" queries. (3) Only answer after collecting content from all documents. Never answer based solely on the initial context when multi-document comparison is requested.
${inferredTarget.folderUncertain ? "- IMPORTANT: The user referred to a SharePoint folder, but its name could not be determined with confidence. Ask the user which folder name to search. Do not search broadly and do not answer from unrelated documents until the folder is clarified." : ""}
- PowerPoint生成のツール選択ルール（厳守）：
  ① ユーザーがPDF/文書を「そのままPPTに変換して」「スライド化して」と言った場合 → convert_doc_to_pptx を使うこと。
  ② 会話の中でスライド構成（タイトル・箇条書き）が既に議論・提示されており、「PDFを参考に内容を拡充して」「追記して」「厚くして」とPDFはあくまで参考資料として扱う場合 → create_pptx を使うこと。この場合、前の会話のスライド構成を slides パラメータのベースとし、文書コンテキストや sl_doc_search の結果で各スライドの bullets を肉付けした上で呼び出すこと。convert_doc_to_pptx は使わないこと。${fileUrlHint}

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
  tools.push(
    createSharePointPdfSummaryTool({
      deptLower,
      userHash: userHash ?? undefined,
      signal,
      userMessage,
      chatThreadId: chatThread.id,
    })
  );

  // RAG経路でもDeskNet's Native Toolを利用可能にする。
  // chat-api.ts側で追加したToolは、RAGが独自にTool一覧を構成するため再追加が必要。
  if (
    isDeskNetsAgentEnabled() &&
    shouldRouteToDeskNetsAgent(userMessage, history)
  ) {
    tools.push(createDeskNetsAgentTool(chatThread.id, userMessage));
    console.log("[DeskNetsAgent] Native tool enabled in RAG path", {
      chatThreadId: chatThread.id,
    });
  }

  // ★ sl_doc_search ツール：LLMが複数文書を横断検索するために複数回呼び出し可能
  tools.push({
    type: "function",
    function: {
      name: "sl_doc_search",
      description:
        "SharePointの個人・部署・全社共通ドキュメントを検索します。\n" +
        "ユーザーが検索先を指定した場合はscopeとfolderへ必ず反映してください。\n" +
        "例：「個人ファイルから」「個人フォルダーから」→scope=personal（folderは設定しない）、「部署共通から」→scope=dept_common、「全社共通から」→scope=global_common、「〇〇フォルダーから」→folder=〇〇。「部署共通フォルダー内の可能性調査というフォルダー」→scope=dept_common, folder=可能性調査。引用符は不要です。フォルダー名を一意に判断できない場合は検索せずユーザーに確認してください。\n" +
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
          scope: {
            type: "string",
            enum: ["all", "personal", "dept_common", "global_common"],
            description:
              "検索範囲。指定なし/allは閲覧可能な全資料、personalは自分の個人資料、dept_commonは部署共通、global_commonは全社共通。",
          },
          folder: {
            type: "string",
            description:
              "検索対象のSharePointフォルダー名。ユーザーが『〇〇フォルダーから』などと指定した場合だけ設定します。",
          },
        },
        required: ["query"],
      },
      function: async (args: {
        query: string;
        mode?: string;
        scope?: SlSearchScope;
        folder?: string;
      }) => {
        if (inferredTarget.folderUncertain && !args.folder?.trim()) {
          return "【要確認】検索対象のフォルダー名を一意に判断できません。ユーザーにフォルダー名を確認してから、folderを指定して再検索してください。";
        }
        const effectiveTop = args.mode === "discover" ? 32 : 8;
        const targetFilter = buildSlSearchTargetFilter(args);
        const searchQuery = stripSlSearchTargetTerms(args.query, args);
        const scopedFilter = targetFilter
          ? `(${filter}) and (${targetFilter})`
          : filter;
        console.log("[sl_doc_search] query =", args.query, "effectiveQuery =", searchQuery, "mode =", args.mode ?? "content", "scope =", args.scope ?? "all", "folder =", args.folder ?? "(none)", "top =", effectiveTop);
        const searchResult = await ExtensionSimilaritySearch({
          searchText: searchQuery,
          vectors: ["embedding"],
          apiKey,
          searchName,
          indexName,
          filter: scopedFilter,
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
        ...(forceDeskNetsAgent
          ? {
              tool_choice: {
                type: "function",
                function: { name: "desknets_schedule_agent" },
              },
            }
          : {}),
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
