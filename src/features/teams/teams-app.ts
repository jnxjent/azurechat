import "server-only";

import { App } from "@microsoft/teams.apps";
import { NextHttpAdapter } from "./next-http-adapter";
import { createTeamsChatReply } from "./teams-chat-service";
import { isTeamsSearchConfigured } from "./teams-search-service";
import { isTeamsBraveSearchConfigured } from "./teams-brave-search-service";
import {
  executeTeamsOfficeRequest,
  parseTeamsOfficeRequest,
} from "./teams-office-service";

export const TEAMS_MESSAGING_ENDPOINT = "/api/teams/messages" as const;

type TeamsRuntime = {
  app: App;
  adapter: NextHttpAdapter;
};

let runtimePromise: Promise<TeamsRuntime> | undefined;

export function isTeamsEnabled(): boolean {
  return process.env.TEAMS_ENABLED?.trim().toLowerCase() === "true";
}

export function isTeamsConfigured(): boolean {
  return Boolean(
    resolveClientId() && resolveClientSecret() && resolveTenantId()
  );
}

export function isTeamsAiSearchConfigured(): boolean {
  return isTeamsSearchConfigured();
}

export function isTeamsWebSearchConfigured(): boolean {
  return isTeamsBraveSearchConfigured();
}

export async function getTeamsRuntime(): Promise<TeamsRuntime> {
  if (!isTeamsEnabled()) {
    throw new Error("Teams integration is disabled. Set TEAMS_ENABLED=true.");
  }

  if (!runtimePromise) {
    runtimePromise = createTeamsRuntime().catch((error) => {
      runtimePromise = undefined;
      throw error;
    });
  }

  return runtimePromise;
}

async function createTeamsRuntime(): Promise<TeamsRuntime> {
  const skipAuth = isLocalAuthSkipped();
  const clientId =
    resolveClientId() ??
    (skipAuth ? "00000000-0000-0000-0000-000000000011" : undefined);
  const clientSecret =
    resolveClientSecret() ?? (skipAuth ? "local-playground-only" : undefined);
  const tenantId =
    resolveTenantId() ??
    (skipAuth ? "00000000-0000-0000-0000-000000000001" : undefined);

  if (!clientId || !clientSecret || !tenantId) {
    throw new Error(
      "Teams configuration is incomplete. Set TEAMS_BOT_ID, TEAMS_BOT_SECRET, and TEAMS_TENANT_ID (or reuse the AZURE_AD_* values)."
    );
  }

  const adapter = new NextHttpAdapter();

  const app = new App({
    clientId,
    clientSecret,
    tenantId,
    httpServerAdapter: adapter,
    messagingEndpoint: TEAMS_MESSAGING_ENDPOINT,
    skipAuth,
    activity: {
      mentions: { stripText: true },
    },
  });

  app.on("message", async ({ activity, send, api }) => {
    const text = activity.text?.trim() || "(empty message)";
    const conversationId = activity.conversation?.id;

    if (!conversationId) {
      await send("会話を識別できなかったため、回答を作成できませんでした。");
      return;
    }

    try {
      const userEmail = await resolveActivityUserEmail({ activity, api });
      const officeRequest = parseTeamsOfficeRequest(text);
      if (officeRequest) {
        const startMessage =
          officeRequest.action === "pdf_to_excel"
            ? `「${officeRequest.fileQuery}」を検索し、Excelへの変換を開始します。完了までしばらくお待ちください。`
            : `Excelの「${officeRequest.targetSheets.join(", ") || "指定シート"}」を再変換します。完了までしばらくお待ちください。`;
        await send(startMessage);
        const officeReply = await executeTeamsOfficeRequest({
          request: officeRequest,
          conversationId,
          userEmail,
        });
        await send(officeReply);
        return;
      }

      const result = await createTeamsChatReply({
        conversationId,
        message: text,
        userEmail,
      });
      await send(result.text);
    } catch (error) {
      console.error("[teams] AzureChat reply failed", error);
      await send(
        "AzureChatへの接続中にエラーが発生しました。しばらくしてから再度お試しください。"
      );
    }
  });

  await app.initialize();
  return { app, adapter };
}

async function resolveActivityUserEmail(props: {
  activity: {
    channelId?: unknown;
    conversation?: { id?: string };
    from?: {
      id?: string;
      properties?: Record<string, unknown>;
    };
  };
  api: {
    conversations: {
      members: (conversationId: string) => {
        getById: (id: string) => Promise<{
          email?: string;
          userPrincipalName?: string;
        }>;
      };
    };
  };
}): Promise<string | null> {
  const directEmail = firstEmail(
    props.activity.from?.properties?.email,
    props.activity.from?.properties?.userPrincipalName
  );
  if (directEmail) return directEmail;

  const conversationId = props.activity.conversation?.id;
  const memberId = props.activity.from?.id;
  if (
    String(props.activity.channelId ?? "") !== "msteams" ||
    !conversationId ||
    !memberId
  ) {
    return null;
  }

  try {
    const member = await props.api.conversations
      .members(conversationId)
      .getById(memberId);
    return firstEmail(member.email, member.userPrincipalName);
  } catch (error) {
    console.warn("[teams] Failed to resolve conversation member", error);
    return null;
  }
}

function firstEmail(...values: unknown[]): string | null {
  for (const value of values) {
    if (typeof value !== "string") continue;
    const normalized = value.trim().toLowerCase();
    if (normalized.includes("@")) return normalized;
  }
  return null;
}

function resolveClientId(): string | undefined {
  return clean(process.env.TEAMS_BOT_ID) ?? clean(process.env.AZURE_AD_CLIENT_ID);
}

function resolveClientSecret(): string | undefined {
  return (
    clean(process.env.TEAMS_BOT_SECRET) ??
    clean(process.env.AZURE_AD_CLIENT_SECRET)
  );
}

function resolveTenantId(): string | undefined {
  return (
    clean(process.env.TEAMS_TENANT_ID) ??
    clean(process.env.AZURE_AD_TENANT_ID)
  );
}

function clean(value: string | undefined): string | undefined {
  const normalized = value?.trim();
  return normalized || undefined;
}

function isLocalAuthSkipped(): boolean {
  return (
    process.env.NODE_ENV !== "production" &&
    process.env.TEAMS_SKIP_AUTH?.trim().toLowerCase() === "true"
  );
}
