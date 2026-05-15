// app/(authenticated)/api/sl/sync-check/route.ts
//
// SharePoint SLフォルダと Azure AI Search Index の同期チェック API
//
// 変更点:
// - session access token 方式を廃止
// - app-only token (client credentials) 方式へ変更
// - x-sl-sync-key で保護された手動実行APIとして利用
//
// 必要な env:
// - AZURE_AD_TENANT_ID
// - AZURE_AD_CLIENT_ID
// - AZURE_AD_CLIENT_SECRET
// - SL_SYNC_CHECK_KEY

export const runtime = "nodejs";
export const maxDuration = 300;

import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { options as authOptions } from "@/features/auth-page/auth-api";
import { runSlSync } from "@/lib/sl-sync";

// -------------------------------------------------------
// Optional endpoint guard
// -------------------------------------------------------
async function requireSyncKey(req: NextRequest) {
  const session = await getServerSession(authOptions);
  if ((session?.user as any)?.email) return;

  const required = (process.env.SL_SYNC_CHECK_KEY ?? "").trim();
  if (!required) return;

  const got = (req.headers.get("x-sl-sync-key") ?? "").trim();
  if (!got || got !== required) {
    throw new Error("Forbidden: invalid x-sl-sync-key");
  }
}

// -------------------------------------------------------
// App-only Graph token (client credentials)
// -------------------------------------------------------
async function getAppOnlyAccessToken(): Promise<string> {
  const tenantId = (process.env.AZURE_AD_TENANT_ID ?? "").trim();
  const clientId = (process.env.AZURE_AD_CLIENT_ID ?? "").trim();
  const clientSecret = (process.env.AZURE_AD_CLIENT_SECRET ?? "").trim();

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      "Missing Azure AD env vars (AZURE_AD_TENANT_ID / AZURE_AD_CLIENT_ID / AZURE_AD_CLIENT_SECRET)"
    );
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    cache: "no-store",
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "client_credentials",
      scope: "https://graph.microsoft.com/.default",
    }),
  });

  const data: any = await res.json().catch(() => ({}));

  if (!res.ok) {
    console.error("[SL sync-check] App token fetch failed:", data);
    throw new Error(`Failed to get app-only access token: ${JSON.stringify(data)}`);
  }

  const accessToken = data?.access_token as string | undefined;
  if (!accessToken) {
    throw new Error("No access_token returned from Azure AD");
  }

  return accessToken;
}

// -------------------------------------------------------
// Concurrency guard — prevent parallel sync runs
// -------------------------------------------------------
let syncRunning = false;

// -------------------------------------------------------
// Route handlers
// -------------------------------------------------------
export async function OPTIONS() {
  return NextResponse.json({}, { status: 200 });
}

export async function POST(req: NextRequest) {
  if (syncRunning) {
    console.warn("[SL sync-check] Sync already running, rejecting concurrent request");
    return NextResponse.json({ ok: false, error: "同期が既に実行中です。完了をお待ちください。" }, { status: 429 });
  }

  syncRunning = true;
  try {
    await requireSyncKey(req);

    const accessToken = await getAppOnlyAccessToken();

    const url = new URL(req.url);
    const apply = url.searchParams.get("apply") === "true";
    const indexNew = url.searchParams.get("indexNew") === "true";
    const batchSize = Math.max(
      1,
      Math.min(20, parseInt(url.searchParams.get("batchSize") ?? "5", 10) || 5)
    );

    const result = await runSlSync({
      accessToken,
      apply,
      indexNew,
      batchSize,
    });

    return NextResponse.json(result);
  } catch (e: any) {
    console.error("[SL sync-check] Error:", e);
    const msg = String(e?.message ?? e);
    const status = msg.startsWith("Forbidden:") ? 403 : 500;
    return NextResponse.json({ ok: false, error: msg }, { status });
  } finally {
    syncRunning = false;
  }
}