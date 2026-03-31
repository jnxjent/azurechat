// src/lib/sl-sync.ts
//
// SharePoint SLフォルダと Azure AI Search Index を照合し、
// 孤立した Index ドキュメントを検出 / 削除する共通ロジック
//
// - SharePoint: Microsoft Graph (delegated token)
// - Azure AI Search: REST API (api-key)
//
// ポイント:
// - route.ts / Timer から共通利用できる
// - apply=false なら dry-run
// - apply=true なら削除実行
// - SP参照失敗(fetchFailed=true)時は削除スキップ

import { getDeptConfig, getAllowedDepts } from "@/lib/sl-dept";

// -------------------------------------------------------
// Types
// -------------------------------------------------------
export type SlSyncDeptResult = {
  spFileNames: number;
  indexDocs: number | "unknown";
  orphanIds?: string[];
  deleted: number;
  skipped?: string;
  error?: string;
};

export type SlSyncResult = {
  ok: true;
  mode: "dry-run" | "apply";
  results: Record<string, SlSyncDeptResult>;
};

export type RunSlSyncParams = {
  accessToken: string;
  apply?: boolean;
};

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

function safeDecodeURIComponent(v: string): string {
  try {
    return decodeURIComponent(v);
  } catch {
    return v;
  }
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

  if (!res.ok) {
    throw new Error(`Failed to get site: ${await res.text()}`);
  }

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

  if (!res.ok) {
    throw new Error(`Failed to get drives: ${await res.text()}`);
  }

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
    const res = await fetch(nextUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
      cache: "no-store",
    });

    if (res.status === 404) {
      console.warn(`[SL sync] Folder not found (404): ${currentFolderPath}`);
      return { fetchFailed: true };
    }

    if (!res.ok) {
      throw new Error(`Failed to list folder "${currentFolderPath}": ${await res.text()}`);
    }

    const json = await res.json();

    for (const item of json.value ?? []) {
      if (item?.file && item?.name) {
        fileNames.add(String(item.name));
      } else if (item?.folder && item?.name) {
        const child = await collectFileNamesRecursive(
          accessToken,
          driveId,
          `${currentFolderPath}/${item.name}`,
          fileNames
        );

        if (child.fetchFailed) {
          console.warn(
            `[SL sync] Child folder fetch failed: ${currentFolderPath}/${item.name} — propagating to parent`
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
    `[SL sync] SP recursive scan: baseFolder="${baseFolder}" total=${fileNames.size} fetchFailed=${fetchFailed}`
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
        `&$filter=dept eq '${dept.replace(/'/g, "''")}' and isSlDoc eq true` +
        `&$top=${top}&$skip=${skip}`,
      {
        headers: { "api-key": apiKey, "Content-Type": "application/json" },
        cache: "no-store",
      }
    );

    if (!res.ok) {
      throw new Error(`Search query failed: ${await res.text()}`);
    }

    const json = await res.json();
    const items: any[] = json.value ?? [];
    if (items.length === 0) break;

    for (const item of items) {
      const rawUrl: string = item.fileUrl ?? "";
      const decoded = safeDecodeURIComponent(rawUrl);
      const fileName = decoded.split("/").pop() ?? "";
      if (item.id && fileName) {
        docs.push({ id: String(item.id), fileName });
      }
    }

    skip += items.length;
    if (items.length < top) break;
  }

  return docs;
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

  if (!res.ok) {
    throw new Error(`Delete failed: ${await res.text()}`);
  }

  console.log(`[SL sync] Deleted ${ids.length} index docs`);
}

// -------------------------------------------------------
// Main
// -------------------------------------------------------
export async function runSlSync({
  accessToken,
  apply = false,
}: RunSlSyncParams): Promise<SlSyncResult> {
  const results: Record<string, SlSyncDeptResult> = {};

  for (const dept of getAllowedDepts()) {
    try {
      const { siteUrl, driveName, folder: baseFolder } = getDeptConfig(dept);

      const { fileNames: spFileNames, fetchFailed } = await getSpFileNames(
        accessToken,
        siteUrl,
        driveName,
        baseFolder
      );

      console.log(
        `[SL sync] dept=${dept} SP fileNames=${spFileNames.size} fetchFailed=${fetchFailed}`
      );

      if (fetchFailed) {
        console.warn(
          `[SL sync] dept=${dept} SP fetch failed — skipping deletion to avoid data loss`
        );

        results[dept] = {
          spFileNames: 0,
          indexDocs: "unknown",
          deleted: 0,
          skipped: "sp_fetch_failed",
        };
        continue;
      }

      const indexDocs = await getIndexDocs(dept);
      console.log(`[SL sync] dept=${dept} indexed docs=${indexDocs.length}`);

      const orphanIds = indexDocs
        .filter((doc) => !spFileNames.has(doc.fileName))
        .map((doc) => doc.id);

      if (orphanIds.length === 0) {
        results[dept] = {
          spFileNames: spFileNames.size,
          indexDocs: indexDocs.length,
          deleted: 0,
          orphanIds: [],
        };
        continue;
      }

      console.log(`[SL sync] dept=${dept} orphans=${orphanIds.length}`);

      if (apply) {
        await deleteIndexDocs(orphanIds);
      }

      results[dept] = {
        spFileNames: spFileNames.size,
        indexDocs: indexDocs.length,
        deleted: apply ? orphanIds.length : 0,
        orphanIds,
      };
    } catch (deptErr: any) {
      console.error(`[SL sync] dept=${dept} error:`, deptErr);
      results[dept] = {
        spFileNames: 0,
        indexDocs: 0,
        deleted: 0,
        error: String(deptErr?.message ?? deptErr),
      };
    }
  }

  return {
    ok: true,
    mode: apply ? "apply" : "dry-run",
    results,
  };
}