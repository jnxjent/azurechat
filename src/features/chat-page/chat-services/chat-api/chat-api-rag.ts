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
async function resolveUserContext(): Promise<UserContext> {
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
      const isThisThread = result.content.document.chatThreadId === chatThread.id;
      const blobUrl = isThisThread ? (result.content.document.fileUrl ?? displayUrl) : null;
      return `[${index}]. file name: ${result.content.document.metadata}
file id: ${result.id}${blobUrl ? `\nfile_url: ${blobUrl}` : ""}
${page}`;
    })
    .join("\n------\n");

  // ファイルURLリスト（convert_doc_to_pptx ツールに渡すため）
  // このスレッドにアップロードされたファイルのみを対象にする
  const fileUrls = uploadedBlobUrls;
  const fileUrlHint =
    fileUrls.length > 0
      ? `\n- The uploaded document file URLs are:\n${fileUrls.map((u, i) => `  [${i}] ${u}`).join("\n")}\n- If the user asks to convert the document to PowerPoint, use the convert_doc_to_pptx tool with the file_url from above.`
      : "";

  const _userMessage = `
- Review the following content from documents uploaded by the user and create a final answer.
- If you don't know the answer, just say that you don't know. Don't try to make up an answer.
- You must always include a citation at the end of your answer and don't include full stop after the citations.
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
