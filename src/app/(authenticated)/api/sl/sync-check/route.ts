// app/(authenticated)/api/sl/sync-check/route.ts
//
// 委任トークン（セッション）で SharePoint SLフォルダを参照し、
// Indexに残っている孤立ドキュメントを削除する。
// - SharePoint: Graph (delegate / session accessToken)
// - Azure AI Search: REST API (api-key)

export const runtime = "nodejs";

import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { getToken } from "next-auth/jwt";
import { options as authOptions } from "@/features/auth-page/auth-api";
import { getDeptConfig, getAllowedDepts } from "@/lib/sl-dept";

// -------------------------------------------------------
// Optional endpoint guard
// -------------------------------------------------------
async function requireSyncKey(req: NextRequest) {
  // ブラウザからの管理者セッションがあればOK
  const session = await getServerSession(authOptions);
  if ((session?.user as any)?.email) return;

  // Logic App等からはキーで認証
  const required = (process.env.SL_SYNC_CHECK_KEY ?? "").trim();
  if (!required) return;

  const got = (req.headers.get("x-sl-sync-key") ?? "").trim();
  if (!got || got !== required) {
    throw new Error("Forbidden: invalid x-sl-sync-key");
  }
}

// -------------------------------------------------------
// 委任Token取得（publish/route.ts と同じ考え方）
// -------------------------------------------------------
async function getValidAccessToken(req: NextRequest): Promise<string | null> {
  const token = await getToken({ req });
  if (!token) return null;

  const accessToken = (token as any).accessToken as string | undefined;
  const expiresAt = (token as any).accessTokenExpiresAt as number | undefined;
  const refreshToken = (token as any).refreshToken as string | undefined;

  if (!accessToken) return null;

  const nowSec = Math.floor(Date.now() / 1000);
  if (expiresAt && nowSec < expiresAt - 60) {
    return accessToken;
  }

  if (!refreshToken) return accessToken;

  const tenantId = process.env.AZURE_AD_TENANT_ID;
  const clientId = process.env.AZURE_AD_CLIENT_ID;
  const clientSecret = process.env.AZURE_AD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) return accessToken;

  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        cache: "no-store",
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          grant_type: "refresh_token",
          refresh_token: refreshToken,
          scope: "openid profile email offline_access User.Read Files.ReadWrite",
        }),
      }
    );

    const data: any = await res.json().catch(() => ({}));
    if (!res.ok) {
      console.error("[SL sync-check] Token refresh failed:", data);
      return accessToken;
    }

    return data.access_token as string;
  } catch (e) {
    console.error("[SL sync-check] Token refresh error:", e);
    return accessToken;
  }
}

// -------------------------------------------------------
// Helpers
// -------------------------------------------------------
function encodeGraphPath(path: string): string {
  return (path ?? "")
    .split("/")
    .filter((s) => s.length > 0)
    .map((seg) => encodeURIComponent(seg))
    .join("/");
}

async function resolveSiteId(accessToken: string, siteUrl: string): Promise<string> {
  const url = new URL(siteUrl);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${url.hostname}:${url.pathname}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
      cache: "no-store",
    }
  );
  if (!res.ok) throw new Error(`Failed to get site: ${await res.text()}`);
  const json = await res.json();
  return json.id as string;
}

async function resolveDriveId(
  accessToken: string,
  siteId: string,
  driveName: string
): Promise<string> {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
      cache: "no-store",
    }
  );
  if (!res.ok) throw new Error(`Failed to get drives: ${await res.text()}`);
  const json = await res.json();
  const drive = (json.value ?? []).find((d: any) => d.name === driveName);
  if (!drive) {
    const names = (json.value ?? []).map((d: any) => d.name).join(", ");
    throw new Error(`Drive "${driveName}" not found. Available: ${names}`);
  }
  return drive.id as string;
}

// -------------------------------------------------------
// Graph API: SharePoint folder file names
// -------------------------------------------------------
async function getSpFileNames(
  accessToken: string,
  siteUrl: string,
  driveName: string,
  folder: string
): Promise<Set<string>> {
  const siteId = await resolveSiteId(accessToken, siteUrl);
  const driveId = await resolveDriveId(accessToken, siteId, driveName);

  const folderPath = encodeGraphPath(folder);

  const fileNames = new Set<string>();
  let nextUrl: string | null =
    `https://graph.microsoft.com/v1.0/drives/${driveId}` +
    `/root:/${folderPath}:/children?$select=name,file&$top=200`;

  while (nextUrl) {
    const res: Response = await fetch(nextUrl, {  // ← Response 型を明示
      headers: { Authorization: `Bearer ${accessToken}` },
      cache: "no-store",
    });

    if (res.status === 404) break; // folder not found -> empty
    if (!res.ok) throw new Error(`Failed to list folder: ${await res.text()}`);

    const json = await res.json();
    for (const item of json.value ?? []) {
      if (item?.file && item?.name) fileNames.add(String(item.name));
    }
    nextUrl = json["@odata.nextLink"] ?? null;
  }

  return fileNames;
}

// -------------------------------------------------------
// Azure Search: dept docs
// -------------------------------------------------------
async function getIndexDocs(
  dept: string
): Promise<Array<{ id: string; fileName: string }>> {
  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  
  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars (AZURE_SEARCH_ENDPOINT/API_KEY/INDEX_NAME)");
  }

  const docs: Array<{ id: string; fileName: string }> = [];
  let skip = 0;
  const top = 200;

  while (true) {
    const res = await fetch(
      `${endpoint}/indexes/${indexName}/docs?api-version=2024-07-01` +
        `&$select=id,fileUrl,dept` +
        `&$filter=dept eq '${dept.replace(/'/g, "''")}'` +
        `&$top=${top}&$skip=${skip}`,
      {
        headers: {
          "api-key": apiKey,
          "Content-Type": "application/json",
        },
        cache: "no-store",
      }
    );

    if (!res.ok) throw new Error(`Search query failed: ${await res.text()}`);

    const json = await res.json();
    const items: any[] = json.value ?? [];
    if (items.length === 0) break;

    for (const item of items) {
      const rawUrl: string = item.fileUrl ?? "";
      const decoded = safeDecodeURIComponent(rawUrl);
      const fileName = decoded.split("/").pop() ?? "";
      if (item.id && fileName) docs.push({ id: String(item.id), fileName });
    }

    skip += items.length;
    if (items.length < top) break;
  }

  return docs;
}

function safeDecodeURIComponent(v: string): string {
  try {
    return decodeURIComponent(v);
  } catch {
    return v;
  }
}

// -------------------------------------------------------
// Azure Search: delete docs by ids
// -------------------------------------------------------
async function deleteIndexDocs(ids: string[]): Promise<void> {
  if (ids.length === 0) return;

  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;

  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars (AZURE_SEARCH_ENDPOINT/API_KEY/INDEX_NAME)");
  }

  const body = {
    value: ids.map((id) => ({ "@search.action": "delete", id })),
  };

  const res = await fetch(
    `${endpoint}/indexes/${indexName}/docs/index?api-version=2024-07-01`,
    {
      method: "POST",
      headers: { "api-key": apiKey, "Content-Type": "application/json" },
      body: JSON.stringify(body),
      cache: "no-store",
    }
  );

  if (!res.ok) throw new Error(`Delete failed: ${await res.text()}`);
  console.log(`[SL sync-check] Deleted ${ids.length} index docs`);
}

// -------------------------------------------------------
// Route handlers
// -------------------------------------------------------
export async function OPTIONS() {
  return NextResponse.json({}, { status: 200 });
}

export async function POST(req: NextRequest) {
  try {
    await requireSyncKey(req);

    const accessToken = await getValidAccessToken(req);
    if (!accessToken) {
      return NextResponse.json(
        { ok: false, error: "No access token in session" },
        { status: 401 }
      );
    }

    const results: Record<string, any> = {};

    for (const dept of getAllowedDepts()) {
      try {
        const { siteUrl, driveName, folder } = getDeptConfig(dept);

        const spFiles = await getSpFileNames(accessToken, siteUrl, driveName, folder);
        console.log(`[SL sync-check] dept=${dept} SP files=${spFiles.size}`);

        const indexDocs = await getIndexDocs(dept);
        console.log(`[SL sync-check] dept=${dept} indexed files=${indexDocs.length}`);

        const orphanIds = indexDocs
          .filter((doc) => !spFiles.has(doc.fileName))
          .map((doc) => doc.id);

        if (orphanIds.length === 0) {
          results[dept] = {
            spFiles: spFiles.size,
            indexDocs: indexDocs.length,
            deleted: 0,
          };
          continue;
        }

        console.log(`[SL sync-check] dept=${dept} orphans=${orphanIds.length}`);
        await deleteIndexDocs(orphanIds);

        results[dept] = {
          spFiles: spFiles.size,
          indexDocs: indexDocs.length,
          deleted: orphanIds.length,
        };
      } catch (deptErr: any) {
        console.error(`[SL sync-check] dept=${dept} error:`, deptErr);
        results[dept] = { error: String(deptErr?.message ?? deptErr) };
      }
    }

    return NextResponse.json({ ok: true, results });
  } catch (e: any) {
    console.error("[SL sync-check] Error:", e);
    const msg = String(e?.message ?? e);
    const status = msg.startsWith("Forbidden:") ? 403 : 500;
    return NextResponse.json({ ok: false, error: msg }, { status });
  }
}