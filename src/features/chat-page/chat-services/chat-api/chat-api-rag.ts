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
import { decideDept, getUserEmailFromJwtToken } from "@/lib/sl-dept";
import { getToken } from "next-auth/jwt";
import { cookies } from "next/headers";

// OData filter用にシングルクォートをエスケープ
function odataEscape(v: string) {
  return String(v ?? "").replace(/'/g, "''");
}

/**
 * サーバ側で「ユーザーの deptLower」を決める
 * - token から email を抜く
 * - token 無い場合は SL_LOCAL_EMAIL を使用（Local開発用）
 * - email → dept 判定（sl-dept.ts）
 * - fallback は SL_DEPT_DEFAULT
 */
async function resolveDeptLower(): Promise<string> {
  try {
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

    let email = token ? getUserEmailFromJwtToken(token) : null;

    // ★ Local fallback
    if (!email && process.env.SL_LOCAL_EMAIL) {
      email = process.env.SL_LOCAL_EMAIL;
    }

    console.log("[DOC] email =", email);

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
  const deptLower = await resolveDeptLower();

  console.log("[RAG-EXT] deptLower =", deptLower);

  // 業務条件は chatThreadId のみ
  const filter = `chatThreadId eq '${odataEscape(chatThread.id)}'`;

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
  });

  const documents: ChatCitationModel[] = [];

  if (documentResponse.status === "OK") {
    const withoutEmbedding = FormatCitations(documentResponse.response);
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