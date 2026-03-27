// app/(authenticated)/api/sl/sync-check/route.ts
//
// 委任トークン（セッション）で SharePoint SLフォルダを参照し、
// Indexに残っている孤立ドキュメントを削除する。
// - SharePoint: Graph (delegate / session accessToken)
// - Azure AI Search: REST API (api-key)
//
// v3 変更点（最終版）:
//   - SP走査: SL直下のみ → baseFolder配下を再帰走査
//   - 照合キー: ファイル名（BlobURLと一致するため）
//   - 安全ガード: fetchFailed=true（404等）の場合は削除スキップ
//   - 安全ガード: 子フォルダのfetchFailedを親に伝播

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
// 委任Token取得
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
    { headers: { Authorization: `Bearer ${accessToken}` }, cache: "no-store" }
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
    { headers: { Authorization: `Bearer ${accessToken}` }, cache: "no-store" }
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
// Graph API: baseFolder配下を再帰走査してファイル名を収集
//
// 照合キーはファイル名（item.name）
// IndexのfileUrlがBlobURLのため、relativePath照合は不可
//
// 返値:
//   fileNames   : Set<string> — ファイル名のSet
//   fetchFailed : true = 404など参照失敗（削除禁止）
//                 false = 正常取得（0件でも削除許可）
// -------------------------------------------------------
async function collectFileNamesRecursive(
  accessToken: string,
  driveId: string,
  currentFolderPath: string,
  fileNames: Set<string>
): Promise<{ fetchFailed: boolean }> {
  const encoded = encodeGraphPath(currentFolderPath);
  let nextUrl: string | null =
    `https://graph.microsoft.com/v1.0/drives/${driveId}` +
    `/root:/${encoded}:/children?$select=name,file,folder&$top=200`;

  while (nextUrl) {
    const res: Response = await fetch(nextUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
      cache: "no-store",
    });

    if (res.status === 404) {
      console.warn(`[SL sync-check] Folder not found (404): ${currentFolderPath}`);
      return { fetchFailed: true };
    }
    if (!res.ok) {
      throw new Error(`Failed to list folder "${currentFolderPath}": ${await res.text()}`);
    }

    const json = await res.json();
    for (const item of json.value ?? []) {
      if (item?.file && item?.name) {
        // ファイル → ファイル名を収集
        fileNames.add(String(item.name));
      } else if (item?.folder && item?.name) {
        // サブフォルダ → 再帰。子で失敗したら親にも伝播して即返す
        const child = await collectFileNamesRecursive(
          accessToken,
          driveId,
          `${currentFolderPath}/${item.name}`,
          fileNames
        );
        if (child.fetchFailed) {
          console.warn(
            `[SL sync-check] Child folder fetch failed: ${currentFolderPath}/${item.name} — propagating to parent`
          );
          return { fetchFailed: true };
        }
      }
    }
    nextUrl = json["@odata.nextLink"] ?? null;
  }

  return { fetchFailed: false };
}

async function getSpFileNames(
  accessToken: string,
  siteUrl: string,
  driveName: string,
  baseFolder: string
): Promise<{ fileNames: Set<string>; fetchFailed: boolean }> {
  const siteId = await resolveSiteId(accessToken, siteUrl);
  const driveId = await resolveDriveId(accessToken, siteId, driveName);

  const fileNames = new Set<string>();
  const { fetchFailed } = await collectFileNamesRecursive(
    accessToken,
    driveId,
    baseFolder,
    fileNames
  );

  console.log(
    `[SL sync-check] SP recursive scan: baseFolder="${baseFolder}" total=${fileNames.size} fetchFailed=${fetchFailed}`
  );
  return { fileNames, fetchFailed };
}

// -------------------------------------------------------
// Azure Search: dept docs（ファイル名で照合）
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
        headers: { "api-key": apiKey, "Content-Type": "application/json" },
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
        const { siteUrl, driveName, folder: baseFolder } = getDeptConfig(dept);

        // SP側: ファイル名のSetを再帰取得
        const { fileNames: spFileNames, fetchFailed } = await getSpFileNames(
          accessToken,
          siteUrl,
          driveName,
          baseFolder
        );
        console.log(
          `[SL sync-check] dept=${dept} SP fileNames=${spFileNames.size} fetchFailed=${fetchFailed}`
        );

        // ★ 安全ガード: 参照失敗（404等）の場合は削除をスキップ
        if (fetchFailed) {
          console.warn(
            `[SL sync-check] dept=${dept} SP fetch failed — skipping deletion to avoid data loss`
          );
          results[dept] = {
            spFileNames: 0,
            indexDocs: "unknown",
            deleted: 0,
            skipped: "sp_fetch_failed",
          };
          continue;
        }

        // Index側: ファイル名で取得
        const indexDocs = await getIndexDocs(dept);
        console.log(`[SL sync-check] dept=${dept} indexed docs=${indexDocs.length}`);

        // 照合: ファイル名で比較
        const orphanIds = indexDocs
          .filter((doc) => !spFileNames.has(doc.fileName))
          .map((doc) => doc.id);

        if (orphanIds.length === 0) {
          results[dept] = {
            spFileNames: spFileNames.size,
            indexDocs: indexDocs.length,
            deleted: 0,
          };
          continue;
        }

        console.log(`[SL sync-check] dept=${dept} orphans=${orphanIds.length}`);
        await deleteIndexDocs(orphanIds);

        results[dept] = {
          spFileNames: spFileNames.size,
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