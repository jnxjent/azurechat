"use server";
import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";
import {
  ChatCompletionStreamingRunner,
  ChatCompletionStreamParams,
} from "openai/resources/beta/chat/completions";
import { ChatCompletionMessageParam } from "openai/resources/chat/completions";
import { ExtensionSimilaritySearch } from "../azure-ai-search/azure-ai-search";
import { CreateCitations, FormatCitations } from "../citation-service";
import { ChatCitationModel, ChatThreadModel } from "../models";

// dept判定ユーティリティ
import { decideDept, getUserEmailFromJwtToken, resolveSlAccess } from "@/lib/sl-dept";
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

    const email = token ? getUserEmailFromJwtToken(token) : null;

    const deptLower = email
      ? resolveSlAccess(email).dept
      : session?.slDept?.trim().toLowerCase() ||
        decideDept({
          requestedDept: undefined,
          userEmail: email,
        });

    const userHash = email ? hashValue(email) : null;

    return {
      email,
      deptLower,
      userHash,
    };
  } catch {
    const deptLower =
      (process.env.SL_DEPT_DEFAULT ?? "cp").toLowerCase().trim() || "cp";

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

  // 業務条件は chatThreadId のみ
  // ACL は azure-ai-search.ts 側の buildSearchAclFilter() に一本化
  const filter = `(chatThreadId eq '${odataEscape(chatThread.id)}' or isSlDoc eq true)`;

  console.log("[RAG-EXT] base filter =", filter);

  // ※ 環境変数名は、あなたの既存プロジェクトに合わせて必要なら読み替えてください
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
    userHash: userHash ?? undefined, // ★ null → undefined に変換
  });

  const documents: ChatCitationModel[] = [];

  if (documentResponse.status === "OK") {
    const withoutEmbedding = FormatCitations(documentResponse.response);

    // 既存シグネチャを維持
    const citationResponse = await CreateCitations(withoutEmbedding);

    citationResponse.forEach((c) => {
      if (c.status === "OK") {
        documents.push(c.response);
      }
    });
  } else {
    console.error(
      "[RAG-EXT] ExtensionSimilaritySearch error:",
      documentResponse.errors
    );
  }

  const content = documents
    .map((result, index) => {
      const page = result.content.document.pageContent;
      return `[${index}]. file name: ${result.content.document.metadata}
file id: ${result.id}
${page}`;
    })
    .join("\n------\n");

  const _userMessage = `
- Review the following content from documents uploaded by the user and create a final answer.
- If you don't know the answer, just say that you don't know. Don't try to make up an answer.
- You must always include a citation at the end of your answer and don't include full stop after the citations.

----------------
content:
${content}

----------------
question:
${userMessage}
`;

  const stream: ChatCompletionStreamParams = {
    model: "",
    stream: true,
    messages: [
      { role: "system", content: chatThread.personaMessage },
      ...history,
      { role: "user", content: _userMessage },
    ],
  };

  return openAI.beta.chat.completions.stream(stream, { signal });
};
