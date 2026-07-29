import "server-only";

import { App } from "@microsoft/teams.apps";
import { NextHttpAdapter } from "./next-http-adapter";
import { createTeamsChatReply } from "./teams-chat-service";
import { isTeamsSearchConfigured } from "./teams-search-service";
import { isTeamsBraveSearchConfigured } from "./teams-brave-search-service";
import {
  buildTeamsThreadId,
  executeTeamsOfficeRequest,
  parseTeamsOfficeRequest,
  registerTeamsUploadedOfficeFiles,
  type TeamsOfficeRequest,
} from "./teams-office-service";
import {
  recordTeamsChatTurn,
  recordTeamsThreadUsage,
} from "./teams-usage-service";
import { resolveTeamsInstantReply } from "./teams-instant-reply";
import { requiresTeamsInternalSearch } from "./teams-search-intent";
import { readLatestTeamsFiles, receiveTeamsFiles } from "./teams-file-service";
import {
  referencesTeamsUpload,
  stripTeamsAttachmentMarkup,
  type TeamsStoredFile,
} from "./teams-file-policy";

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
    const text = stripTeamsAttachmentMarkup(activity.text?.trim() ?? "");
    const conversationId = activity.conversation?.id;
    const activityId = String(activity.id ?? "").trim();

    if (!conversationId) {
      await send("会話を識別できなかったため、回答を作成できませんでした。");
      return;
    }

    try {
      const teamsThreadId = buildTeamsThreadId(conversationId);
      let uploadedFiles: TeamsStoredFile[] = [];
      try {
        uploadedFiles = await receiveTeamsFiles({
          attachments: activity.attachments,
          threadId: teamsThreadId,
        });
        if (uploadedFiles.length > 0) {
          await registerTeamsUploadedOfficeFiles(teamsThreadId, uploadedFiles);
          await send(buildTeamsFileReceivedMessage(uploadedFiles));
        } else if (text && referencesTeamsUpload(text)) {
          uploadedFiles = await readLatestTeamsFiles(teamsThreadId);
        }
      } catch (error) {
        console.error("[teams-file] receive failed", error);
        await send(
          `ファイルを受信できませんでした。${safeTeamsFileErrorMessage(error)}`
        );
        return;
      }

      const teamsUserId = activity.from?.id;
      if (!text && uploadedFiles.length > 0) {
        if (teamsUserId) {
          await recordTeamsThreadUsage({ conversationId, teamsUserId });
          await recordCompletedTeamsTurn({
            conversationId,
            activityId,
            teamsUserId,
          });
        }
        return;
      }

      const messageText = text || "(empty message)";
      const instantReply = resolveTeamsInstantReply(messageText);
      if (instantReply) {
        // Send first so a greeting is not delayed by Graph, Cosmos DB,
        // Azure AI Search, or Azure OpenAI.
        await send(instantReply);
        if (teamsUserId) {
          await recordTeamsThreadUsage({ conversationId, teamsUserId });
          await recordCompletedTeamsTurn({
            conversationId,
            activityId,
            teamsUserId,
          });
        }
        return;
      }

      if (teamsUserId) {
        await recordTeamsThreadUsage({ conversationId, teamsUserId });
      } else {
        console.warn("[teams] Usage statistics skipped: user ID is missing");
      }

      const routingText = appendTeamsFileNamesForRouting(
        messageText,
        uploadedFiles
      );
      const officeRequest = parseTeamsOfficeRequest(routingText);

      // Office operations and internal ACL-aware search require the Teams
      // member's email. General chat and Web search do not.
      const userEmail = officeRequest || requiresTeamsInternalSearch(text)
        ? await resolveActivityUserEmail({ activity, api })
        : null;

      if (officeRequest) {
        const startMessage = buildOfficeStartMessage(officeRequest);
        await send(startMessage);
        const officeReply = await executeTeamsOfficeRequest({
          request: officeRequest,
          conversationId,
          uploadedFiles,
          userEmail,
        });
        await send(officeReply);
        await recordCompletedTeamsTurn({
          conversationId,
          activityId,
          teamsUserId,
        });
        return;
      }

      const result = await createTeamsChatReply({
        conversationId,
        message: messageText,
        userEmail,
      });
      await send(result.text);
      if (result.type === "reply") {
        await recordCompletedTeamsTurn({
          conversationId,
          activityId,
          teamsUserId,
        });
      }
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

function buildTeamsFileReceivedMessage(files: TeamsStoredFile[]): string {
  const list = files
    .map(
      (file) =>
        `- ${file.fileName}（${Math.max(1, Math.ceil(file.size / 1024))} KB）`
    )
    .join("\n");
  return [
    "ファイルを受信しました。",
    "",
    list,
    "",
    "PDFはExcel・Word・PowerPoint変換、Excel・Word・PowerPointは既存の編集機能で利用できます。",
  ].join("\n");
}

function appendTeamsFileNamesForRouting(
  message: string,
  files: TeamsStoredFile[]
): string {
  if (files.length === 0) return message;
  return `${message}\n添付ファイル: ${files
    .map((file) => file.fileName)
    .join("、")}`;
}

function safeTeamsFileErrorMessage(error: unknown): string {
  const message = error instanceof Error ? error.message.trim() : "";
  return message && message.length <= 180
    ? `\n\n${message}`
    : "対応形式とファイルサイズを確認してください。";
}

async function recordCompletedTeamsTurn(props: {
  conversationId: string;
  activityId: string;
  teamsUserId: string | undefined;
}): Promise<void> {
  if (!props.teamsUserId) {
    console.warn("[teams] Chat turn statistics skipped: user ID is missing");
    return;
  }
  if (!props.activityId) {
    console.warn("[teams] Chat turn statistics skipped: activity ID is missing");
    return;
  }
  await recordTeamsChatTurn({
    conversationId: props.conversationId,
    activityId: props.activityId,
    teamsUserId: props.teamsUserId,
  });
}

function buildOfficeStartMessage(request: TeamsOfficeRequest): string {
  if (request.action === "ppt_color_help") {
    return "利用可能なPowerPointの色パターンを確認します。";
  }
  if (request.action === "edit_latest_ppt_color") {
    return "直前に作成したPowerPointの色味を変更します。完了までしばらくお待ちください。";
  }
  if (request.action === "edit_latest_excel") {
    return "直前に作成したExcelファイルを編集します。完了までしばらくお待ちください。";
  }
  if (request.action === "edit_latest_word") {
    return "直前に作成したWordファイルを編集します。完了までしばらくお待ちください。";
  }
  if (request.action === "proofread_sp_word") {
    return "SharePointのWordファイルを検索し、誤字・誤記を変更履歴付きで修正します。完了までしばらくお待ちください。";
  }
  if (request.action === "edit_latest_ppt") {
    return `直前に作成したPowerPointの${request.targetPages
      .map((page) => `P${page}`)
      .join("、")}を編集します。完了までしばらくお待ちください。`;
  }
  if (request.action === "create_ppt_from_sharepoint") {
    return `SharePoint資料を検索し、「${request.title}」のPowerPoint作成を開始します。完了までしばらくお待ちください。`;
  }
  if (request.action === "create_excel") {
    return `「${request.title}」のExcel作成を開始します。完了までしばらくお待ちください。`;
  }
  if (request.action === "create_word") {
    return `「${request.title}」のWord作成を開始します。完了までしばらくお待ちください。`;
  }
  if (request.action === "create_ppt") {
    return `「${request.title}」のPowerPoint作成を開始します。完了までしばらくお待ちください。`;
  }
  if (request.action === "pdf_to_excel") {
    return `「${request.fileQuery}」を検索し、Excelへの変換を開始します。完了までしばらくお待ちください。`;
  }
  if (request.action === "pdf_to_word") {
    return `「${request.fileQuery}」を検索し、Wordへの変換を開始します。完了までしばらくお待ちください。`;
  }
  if (request.action === "pdf_to_ppt") {
    return `「${request.fileQuery}」を検索し、PowerPointへの変換を開始します。完了までしばらくお待ちください。`;
  }
  if (request.action === "translate_pdf_to_pptx") {
    const languageNames = {
      en: "英語",
      pt: "ポルトガル語",
      vi: "ベトナム語",
      id: "インドネシア語",
      "zh-CN": "中国語（簡体字）",
      ko: "韓国語",
      es: "スペイン語",
      fil: "タガログ語",
    } as const;
    return `PDFの日本語を${languageNames[request.targetLanguage]}へ翻訳し、編集可能なPowerPointを作成します。完了までしばらくお待ちください。`;
  }
  return `Excelの「${
    request.targetSheets.join(", ") || "指定シート"
  }」を再変換します。完了までしばらくお待ちください。`;
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
