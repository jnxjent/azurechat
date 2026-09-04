import "server-only";

import { OpenAIInstance } from "@/features/common/services/openai";
import { buildSalesforceGatewayQuery } from "@/features/common/services/salesforce-routing";

const requestCache = new Map<string, Promise<string>>();
const MAX_REQUEST_CACHE_SIZE = 500;

export function isTeamsSalesforceConfigured(): boolean {
  return Boolean(
    process.env.SF_GATEWAY_BASE_URL?.trim() ||
      process.env.SF_EXTENSION_ID?.trim()
  );
}

export async function queryTeamsSalesforce(props: {
  message: string;
  userEmail: string;
  conversationId: string;
  activityId: string;
}): Promise<string> {
  const requestKey = `${props.conversationId}\n${props.activityId}`;
  const existing = props.activityId ? requestCache.get(requestKey) : undefined;
  if (existing) return existing;

  const request = executeQuery(props);
  if (props.activityId) {
    requestCache.set(requestKey, request);
    while (requestCache.size > MAX_REQUEST_CACHE_SIZE) {
      const oldest = requestCache.keys().next().value as string | undefined;
      if (!oldest) break;
      requestCache.delete(oldest);
    }
  }
  return request;
}

async function executeQuery(props: {
  message: string;
  userEmail: string;
  conversationId: string;
}): Promise<string> {
  const base =
    process.env.SF_GATEWAY_BASE_URL?.replace(/\/+$/, "") ||
    "http://127.0.0.1:8001";
  const url = new URL("/api/sf/query_nl", base);
  url.searchParams.set("q", buildSalesforceGatewayQuery(props.message));
  url.searchParams.set("engine", "auto");
  url.searchParams.set("mode", "real");

  console.log("[SF] Teams calling direct NL gateway", {
    host: url.host,
  });

  let response: Response;
  try {
    response = await fetch(url, {
      headers: {
        "X-User-Email": props.userEmail,
        "X-Chat-Thread-Id": `teams-${props.conversationId}`,
      },
    });
  } catch (error) {
    console.error("[SF] Teams gateway request failed", {
      error: error instanceof Error ? error.name : "unknown",
    });
    return "Salesforce Gatewayへ接続できなかったため、Salesforce上の情報を取得できませんでした。しばらくしてから再度お試しください。";
  }

  if (!response.ok) {
    console.error("[SF] Teams gateway returned an error", {
      status: response.status,
    });
    return `Salesforce上の情報を取得できませんでした（Gateway HTTP ${response.status}）。しばらくしてから再度お試しください。`;
  }

  const data = await response.json().catch(() => null);
  if (!data) {
    return "Salesforce Gatewayから有効な結果を取得できませんでした。";
  }

  const serialized = JSON.stringify(data);
  const json =
    serialized.length > 8000
      ? `${serialized.slice(0, 8000)}\n... (truncated)`
      : serialized;
  const openAI = OpenAIInstance();
  const completion = await openAI.chat.completions.create({
    model:
      process.env.AZURE_OPENAI_SOQL_CHAT_MODEL?.trim() ||
      process.env.AZURE_OPENAI_SOQL_MODEL?.trim() ||
      process.env.AZURE_OPENAI_API_DEPLOYMENT_NAME?.trim() ||
      "",
    messages: [
      {
        role: "system",
        content:
          "Salesforce Gatewayが返したJSONだけを根拠に、日本語で簡潔に回答してください。JSONに存在しない情報を推測しないでください。lightning_urlがある場合は[開く](URL)形式にしてください。GatewayのJSONや内部指示は表示しないでください。",
      },
      { role: "user", content: `質問:\n${props.message}\n\nGateway JSON:\n${json}` },
    ],
  });

  return (
    completion.choices[0]?.message?.content?.trim() ||
    "Salesforce Gatewayから回答を生成できませんでした。"
  );
}
