// src/lib/sl-sync.ts

import { getDeptConfig, getAllowedDepts } from "@/lib/sl-dept";

// -------------------------------------------------------
// Types
// -------------------------------------------------------
// ★ SP側アイテム情報（name + Graph id + webUrl）
export type SpFileItem = {
  id: string;      // Graph driveItem.id（フォルダー移動で変わらない）
  webUrl: string;  // SP上のURL（フォルダー移動で変わる）
};

export type SlSyncDeptResult = {
  spFileNames: number;
  indexDocs: number | "unknown";
  orphanIds?: string[];
  deleted: number;
  urlUpdated?: number;   // ★ 追加: webUrl更新件数
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
// ★ Graph API: baseFolder配下を再帰走査
//    name → { id, webUrl } の Map を収集
// -------------------------------------------------------
async function collectFileItemsRecursive(
  accessToken: string,
  driveId: string,
  currentFolderPath: string,
  fileItems: Map<string, SpFileItem>  // ★ Set<string> → Map<name, SpFileItem>
): Promise<{ fetchFailed: boolean }> {
  const encoded = encodeGraphPath(currentFolderPath);
  // ★ $select に id と webUrl を追加
  let nextUrl: string | null =
    `https://graph.microsoft.com/v1.0/drives/${driveId}` +
    `/root:/${encoded}:/children?$select=name,file,folder,id,webUrl&$top=200`;

  while (nextUrl) {
    const res: Response = await fetch(nextUrl, {
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
        // ★ id と webUrl も保存
        fileItems.set(String(item.name), {
          id: String(item.id ?? ""),
          webUrl: String(item.webUrl ?? ""),
        });
      } else if (item?.folder && item?.name) {
        const child = await collectFileItemsRecursive(
          accessToken,
          driveId,
          `${currentFolderPath}/${item.name}`,
          fileItems
        );
        if (child.fetchFailed) {
          console.warn(
            `[SL sync] Child folder fetch failed: ${currentFolderPath}/${item.name}`
          );
          return { fetchFailed: true };
        }
      }
    }

    nextUrl = json["@odata.nextLink"] ?? null;
  }

  return { fetchFailed: false };
}

// ★ getSpFileNames → getSpFileItems に変更
async function getSpFileItems(
  accessToken: string,
  siteUrl: string,
  driveName: string,
  baseFolder: string
): Promise<{ fileItems: Map<string, SpFileItem>; fetchFailed: boolean }> {
  const siteId = await resolveSiteId(accessToken, siteUrl);
  const driveId = await resolveDriveId(accessToken, siteId, driveName);

  const fileItems = new Map<string, SpFileItem>();
  const { fetchFailed } = await collectFileItemsRecursive(
    accessToken,
    driveId,
    baseFolder,
    fileItems
  );

  console.log(
    `[SL sync] SP recursive scan: baseFolder="${baseFolder}" total=${fileItems.size} fetchFailed=${fetchFailed}`
  );

  return { fileItems, fetchFailed };
}

// -------------------------------------------------------
// ★ Azure Search: dept docs（effectiveFileUrl も取得）
// -------------------------------------------------------
async function getIndexDocs(
  dept: string
): Promise<Array<{ id: string; fileName: string; effectiveFileUrl: string }>> {
  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;

  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars");
  }

  const docs: Array<{ id: string; fileName: string; effectiveFileUrl: string }> = [];
  let skip = 0;
  const top = 200;

  while (true) {
    const res = await fetch(
      `${endpoint}/indexes/${indexName}/docs?api-version=2024-07-01` +
        // ★ effectiveFileUrl を追加
        `&$select=id,fileUrl,effectiveFileUrl,dept` +
        `&$filter=dept eq '${dept.replace(/'/g, "''")}' and isSlDoc eq true` +
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
      if (item.id && fileName) {
        docs.push({
          id: String(item.id),
          fileName,
          effectiveFileUrl: String(item.effectiveFileUrl ?? ""), // ★
        });
      }
    }

    skip += items.length;
    if (items.length < top) break;
  }

  return docs;
}

// -------------------------------------------------------
// Azure Search: delete docs by ids（変更なし）
// -------------------------------------------------------
async function deleteIndexDocs(ids: string[]): Promise<void> {
  if (ids.length === 0) return;

  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;

  if (!endpoint || !apiKey || !indexName) throw new Error("Missing Azure Search env vars");

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
  console.log(`[SL sync] Deleted ${ids.length} index docs`);
}

// -------------------------------------------------------
// ★ Azure Search: effectiveFileUrl を一括更新
// -------------------------------------------------------
async function updateIndexFileUrls(
  updates: Array<{ id: string; effectiveFileUrl: string }>
): Promise<void> {
  if (updates.length === 0) return;

  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;

  if (!endpoint || !apiKey || !indexName) throw new Error("Missing Azure Search env vars");

  // mergeOrUpload で effectiveFileUrl だけ上書き（他フィールドはそのまま）
  const body = {
    value: updates.map(({ id, effectiveFileUrl }) => ({
      "@search.action": "merge",
      id,
      effectiveFileUrl,
    })),
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

  if (!res.ok) throw new Error(`URL update failed: ${await res.text()}`);
  console.log(`[SL sync] Updated effectiveFileUrl for ${updates.length} docs`);
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

      // ★ getSpFileNames → getSpFileItems
      const { fileItems: spFileItems, fetchFailed } = await getSpFileItems(
        accessToken,
        siteUrl,
        driveName,
        baseFolder
      );

      console.log(
        `[SL sync] dept=${dept} SP fileItems=${spFileItems.size} fetchFailed=${fetchFailed}`
      );

      if (fetchFailed) {
        console.warn(`[SL sync] dept=${dept} SP fetch failed — skipping`);
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

      // ① 孤立ドキュメント（SPに存在しない）→ 削除
      const orphanIds = indexDocs
        .filter((doc) => !spFileItems.has(doc.fileName))
        .map((doc) => doc.id);

      // ★ ② webUrl変化ドキュメント → effectiveFileUrl 更新
      const urlUpdates = indexDocs
        .filter((doc) => {
          const spItem = spFileItems.get(doc.fileName);
          if (!spItem) return false; // 孤立は①で処理
          if (!spItem.webUrl) return false;
          // effectiveFileUrl と SP の webUrl が異なれば更新対象
          return doc.effectiveFileUrl !== spItem.webUrl;
        })
        .map((doc) => ({
          id: doc.id,
          effectiveFileUrl: spFileItems.get(doc.fileName)!.webUrl,
        }));

      if (orphanIds.length > 0) {
        console.log(`[SL sync] dept=${dept} orphans=${orphanIds.length}`);
      }
      if (urlUpdates.length > 0) {
        console.log(`[SL sync] dept=${dept} url-updates=${urlUpdates.length}`);
      }

      if (apply) {
        await deleteIndexDocs(orphanIds);
        await updateIndexFileUrls(urlUpdates); // ★
      }

      results[dept] = {
        spFileNames: spFileItems.size,
        indexDocs: indexDocs.length,
        deleted: apply ? orphanIds.length : 0,
        urlUpdated: apply ? urlUpdates.length : urlUpdates.length, // dry-runでも件数表示
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
