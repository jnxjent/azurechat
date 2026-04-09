// src/lib/sl-sync.ts

import { getAllowedDepts, getDeptConfig } from "@/lib/sl-dept";

export type SpFileItem = {
  id: string;
  name: string;
  webUrl: string;
  sourceSiteUrl: string;
  relativePath: string;
};

export type SlSyncDeptResult = {
  spFileNames: number;
  indexDocs: number | "unknown";
  orphanIds?: string[];
  deleted: number;
  urlUpdated?: number;
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

type IndexDoc = {
  id: string;
  fileName: string;
  fileUrl: string;
  effectiveFileUrl: string;
  slScope: string | null;
  relativePath: string | null;
  spItemId: string | null;
};

type ScopeKind = "global_common" | "dept_common" | "personal";

type GlobalCommonConfig = {
  siteUrl: string;
  driveName: string;
  folder: string;
};

type SpInventory = {
  allItems: SpFileItem[];
  byName: Map<string, SpFileItem[]>;
  byId: Map<string, SpFileItem>;
};

function encodeGraphPath(path: string): string {
  return (path ?? "")
    .split("/")
    .filter(Boolean)
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

function normalizeSiteUrl(siteUrl: string): string {
  return (siteUrl ?? "").replace(/\/+$/, "").toLowerCase();
}

function normalizeFolderPath(path: string): string {
  return (path ?? "")
    .replace(/\\/g, "/")
    .replace(/^\/+/, "")
    .replace(/\/+$/, "");
}

function pathStartsWith(path: string, prefix: string): boolean {
  const p = normalizeFolderPath(path).toLowerCase();
  const f = normalizeFolderPath(prefix).toLowerCase();
  return p === f || p.startsWith(`${f}/`);
}

function getDecodedPathnameFromWebUrl(webUrl: string): string {
  try {
    const u = new URL(webUrl);
    return safeDecodeURIComponent(u.pathname);
  } catch {
    return safeDecodeURIComponent(webUrl);
  }
}

function getGlobalCommonConfig(): GlobalCommonConfig | null {
  const siteUrl = process.env.SL_COMMON_SITE_URL;
  const driveName = process.env.SL_COMMON_DRIVE_NAME;
  const folder = process.env.SL_COMMON_FOLDER || "Common";

  if (!siteUrl || !driveName) return null;
  return { siteUrl, driveName, folder };
}

function deriveDeptCommonFolder(baseFolder: string): string {
  return `${normalizeFolderPath(baseFolder)}/Common`;
}

function isWithinFolderByWebUrl(webUrl: string, folderPath: string): boolean {
  const decodedPathname = getDecodedPathnameFromWebUrl(webUrl).toLowerCase();
  const normalizedFolder = normalizeFolderPath(folderPath).toLowerCase();

  return (
    decodedPathname.includes(`/${normalizedFolder}/`) ||
    decodedPathname.endsWith(`/${normalizedFolder}`)
  );
}

function resolveScopeFromLocation(params: {
  webUrl: string;
  sourceSiteUrl: string;
  deptSiteUrl: string;
  deptBaseFolder: string;
  globalCommonSiteUrl?: string | null;
  globalCommonFolder?: string | null;
}): ScopeKind {
  const {
    webUrl,
    sourceSiteUrl,
    deptSiteUrl,
    deptBaseFolder,
    globalCommonSiteUrl,
    globalCommonFolder,
  } = params;

  const sourceSite = normalizeSiteUrl(sourceSiteUrl);
  const deptSite = normalizeSiteUrl(deptSiteUrl);
  const globalSite = normalizeSiteUrl(globalCommonSiteUrl || "");
  const deptCommonFolder = deriveDeptCommonFolder(deptBaseFolder);

  if (
    globalSite &&
    sourceSite === globalSite &&
    globalCommonFolder &&
    isWithinFolderByWebUrl(webUrl, globalCommonFolder)
  ) {
    return "global_common";
  }

  if (sourceSite === deptSite && isWithinFolderByWebUrl(webUrl, deptCommonFolder)) {
    return "dept_common";
  }

  return "personal";
}

function buildInventory(items: SpFileItem[]): SpInventory {
  const byName = new Map<string, SpFileItem[]>();
  const byId = new Map<string, SpFileItem>();

  for (const item of items) {
    const key = item.name.toLowerCase();
    const bucket = byName.get(key) ?? [];
    bucket.push(item);
    byName.set(key, bucket);

    if (item.id) {
      byId.set(item.id, item);
    }
  }

  return { allItems: items, byName, byId };
}

function extractRelativePathFromWebUrl(
  webUrl: string,
  roots: Array<{ siteUrl: string; folder: string }>
): string | null {
  const decodedPathname = getDecodedPathnameFromWebUrl(webUrl).toLowerCase();

  for (const root of roots) {
    const sitePath = (() => {
      try {
        return safeDecodeURIComponent(new URL(root.siteUrl).pathname).toLowerCase();
      } catch {
        return "";
      }
    })();

    if (!sitePath || !decodedPathname.startsWith(sitePath)) {
      continue;
    }

    const rootFolder = normalizeFolderPath(root.folder).toLowerCase();
    const marker = `/${rootFolder}/`;
    const markerIndex = decodedPathname.indexOf(marker);

    if (markerIndex >= 0) {
      return decodedPathname.slice(markerIndex + 1);
    }

    const endMarker = `/${rootFolder}`;
    if (decodedPathname.endsWith(endMarker)) {
      return rootFolder;
    }
  }

  return null;
}

function resolveIndexRelativePath(
  doc: Pick<IndexDoc, "effectiveFileUrl" | "fileUrl">,
  deptSiteUrl: string,
  deptBaseFolder: string,
  globalCommon?: GlobalCommonConfig | null
): string | null {
  const roots = [{ siteUrl: deptSiteUrl, folder: deptBaseFolder }];
  if (globalCommon) {
    roots.push({
      siteUrl: globalCommon.siteUrl,
      folder: globalCommon.folder,
    });
  }

  return (
    extractRelativePathFromWebUrl(doc.effectiveFileUrl, roots) ??
    extractRelativePathFromWebUrl(doc.fileUrl, roots)
  );
}

/**
 * spItemId を使って Graph API でドライブ内のアイテムを直接ルックアップ。
 * scan 範囲外（基点フォルダより上位）に移動されたファイルを追跡するために使用。
 */
async function lookupSpItemByIdInDrive(
  accessToken: string,
  siteUrl: string,
  driveName: string,
  itemId: string
): Promise<SpFileItem | null> {
  try {
    const siteId = await resolveSiteId(accessToken, siteUrl);
    const driveId = await resolveDriveId(accessToken, siteId, driveName);

    const res = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}` +
        `?$select=name,file,id,webUrl,parentReference,deleted`,
      { headers: { Authorization: `Bearer ${accessToken}` }, cache: "no-store" }
    );

    if (res.status === 404) return null; // ファイルが実際に削除済み（ゴミ箱も空）
    if (!res.ok) {
      console.warn(`[SL sync] lookupSpItemById failed (${res.status}): ${await res.text()}`);
      return null;
    }

    const item = await res.json();

    // ★ ゴミ箱に入ったアイテムは deleted ファセットが付く → 削除済みとして扱いorphan化
    if (item?.deleted) {
      console.log(`[SL sync] lookupSpItemByIdInDrive: item in Recycle Bin, treating as deleted: itemId=${itemId}`);
      return null;
    }

    if (!item?.file || !item?.name) return null; // フォルダ等はスキップ

    // parentReference.path = "/drives/{id}/root:/folder/path" 形式
    const parentPath = (() => {
      const raw: string = item.parentReference?.path ?? "";
      const rootIdx = raw.indexOf("root:");
      if (rootIdx < 0) return "";
      return normalizeFolderPath(safeDecodeURIComponent(raw.slice(rootIdx + 5)));
    })();

    const relativePath = parentPath
      ? normalizeFolderPath(`${parentPath}/${item.name}`)
      : normalizeFolderPath(String(item.name));

    return {
      id: String(item.id),
      name: String(item.name),
      webUrl: String(item.webUrl),
      sourceSiteUrl: siteUrl,
      relativePath,
    };
  } catch (e) {
    console.warn(`[SL sync] lookupSpItemByIdInDrive error:`, e);
    return null;
  }
}

function findMatchingSpItem(doc: IndexDoc, inventory: SpInventory): SpFileItem | null {
  // 第1優先: SP item ID（ファイル移動後も不変）
  if (doc.spItemId) {
    const byId = inventory.byId.get(doc.spItemId);
    if (byId) return byId;
  }

  // 第2優先: relativePath の完全一致（spItemId 未保存の旧ドキュメント向け）
  if (doc.relativePath) {
    const exact = inventory.allItems.find(
      (item) => item.relativePath.toLowerCase() === doc.relativePath?.toLowerCase()
    );
    if (exact) return exact;
  }

  // 第3優先: ファイル名一致（同名が1件のみの場合）
  const sameName = inventory.byName.get(doc.fileName.toLowerCase()) ?? [];
  if (sameName.length === 1) {
    return sameName[0];
  }

  return null;
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
  const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
    headers: { Authorization: `Bearer ${accessToken}` },
    cache: "no-store",
  });
  if (!res.ok) throw new Error(`Failed to get drives: ${await res.text()}`);
  const json = await res.json();
  const drive = (json.value ?? []).find((d: any) => d.name === driveName);
  if (!drive) {
    const names = (json.value ?? []).map((d: any) => d.name).join(", ");
    throw new Error(`Drive "${driveName}" not found. Available: ${names}`);
  }
  return drive.id as string;
}

async function collectFileItemsRecursive(
  accessToken: string,
  driveId: string,
  currentFolderPath: string,
  sourceSiteUrl: string,
  fileItems: SpFileItem[],
  isRoot = false
): Promise<{ fetchFailed: boolean; rootMissing: boolean }> {
  const encoded = encodeGraphPath(currentFolderPath);
  let nextUrl: string | null =
    `https://graph.microsoft.com/v1.0/drives/${driveId}` +
    `/root:/${encoded}:/children?$select=name,file,folder,id,webUrl,parentReference&$top=200`;

  while (nextUrl) {
    const res: Response = await fetch(nextUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
      cache: "no-store",
    });

    if (res.status === 404) {
      console.warn(`[SL sync] Folder not found (404): ${currentFolderPath}`);
      // isRoot=true ならベースフォルダ自体が存在しない → rootMissing
      return { fetchFailed: true, rootMissing: isRoot };
    }

    if (!res.ok) {
      throw new Error(`Failed to list folder "${currentFolderPath}": ${await res.text()}`);
    }

    const json: any = await res.json();

    for (const item of json.value ?? []) {
      if (item?.file && item?.name) {
        fileItems.push({
          id: String(item.id ?? ""),
          name: String(item.name ?? ""),
          webUrl: String(item.webUrl ?? ""),
          sourceSiteUrl,
          relativePath: normalizeFolderPath(`${currentFolderPath}/${item.name}`),
        });
      } else if (item?.folder && item?.name) {
        const child = await collectFileItemsRecursive(
          accessToken,
          driveId,
          `${currentFolderPath}/${item.name}`,
          sourceSiteUrl,
          fileItems,
          false  // 子フォルダは isRoot=false
        );
        if (child.fetchFailed) {
          console.warn(`[SL sync] Child folder fetch failed: ${currentFolderPath}/${item.name}`);
          return { fetchFailed: true, rootMissing: false };
        }
      }
    }

    nextUrl = json["@odata.nextLink"] ?? null;
  }

  return { fetchFailed: false, rootMissing: false };
}

async function getSpFileItemsForFolder(
  accessToken: string,
  siteUrl: string,
  driveName: string,
  folderPath: string
): Promise<{ inventory: SpInventory; fetchFailed: boolean; rootMissing: boolean }> {
  const siteId = await resolveSiteId(accessToken, siteUrl);
  const driveId = await resolveDriveId(accessToken, siteId, driveName);

  const fileItems: SpFileItem[] = [];
  const { fetchFailed, rootMissing } = await collectFileItemsRecursive(
    accessToken,
    driveId,
    folderPath,
    siteUrl,
    fileItems,
    true  // ベースフォルダ呼び出しは isRoot=true
  );

  console.log(
    `[SL sync] SP recursive scan: site="${siteUrl}" folder="${folderPath}" total=${fileItems.length} fetchFailed=${fetchFailed} rootMissing=${rootMissing}`
  );

  return { inventory: buildInventory(fileItems), fetchFailed, rootMissing };
}

async function getSpFileItems(
  accessToken: string,
  deptSiteUrl: string,
  deptDriveName: string,
  deptBaseFolder: string
): Promise<{ inventory: SpInventory; fetchFailed: boolean; rootMissing: boolean }> {
  const deptScan = await getSpFileItemsForFolder(
    accessToken,
    deptSiteUrl,
    deptDriveName,
    deptBaseFolder
  );

  if (deptScan.fetchFailed) {
    // rootMissing を上位に伝播させる（ベースフォルダ不在の情報を保持）
    return { inventory: buildInventory([]), fetchFailed: true, rootMissing: deptScan.rootMissing };
  }

  const merged = [...deptScan.inventory.allItems];
  const globalCommon = getGlobalCommonConfig();

  if (globalCommon) {
    const globalScan = await getSpFileItemsForFolder(
      accessToken,
      globalCommon.siteUrl,
      globalCommon.driveName,
      globalCommon.folder
    );

    if (!globalScan.fetchFailed) {
      merged.push(...globalScan.inventory.allItems);
    } else {
      console.warn(
        `[SL sync] Global common scan skipped: ${globalCommon.siteUrl} / ${globalCommon.folder}`
      );
    }
  }

  console.log(
    `[SL sync] SP merged scan: deptBaseFolder="${deptBaseFolder}" total=${merged.length}`
  );

  return { inventory: buildInventory(merged), fetchFailed: false, rootMissing: false };
}

async function getIndexDocs(
  dept: string,
  deptSiteUrl: string,
  deptBaseFolder: string,
  globalCommon?: GlobalCommonConfig | null
): Promise<IndexDoc[]> {
  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;

  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars");
  }

  const docs: IndexDoc[] = [];
  let skip = 0;
  const top = 200;

  while (true) {
    const res = await fetch(
      `${endpoint}/indexes/${indexName}/docs?api-version=2024-07-01` +
        `&$select=id,metadata,fileUrl,effectiveFileUrl,dept,slScope,spItemId` +
        `&$filter=(dept eq '${dept.replace(/'/g, "''")}' or slScope eq 'global_common') and isSlDoc eq true` +
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
      const fileUrl = String(item.fileUrl ?? "");
      const effectiveFileUrl = String(item.effectiveFileUrl ?? "");
      // metadata が最も信頼できるファイル名。URLパースをフォールバックとする
      const fileName =
        String(item.metadata ?? "").trim() ||
        safeDecodeURIComponent(effectiveFileUrl).split("/").pop() ||
        safeDecodeURIComponent(fileUrl).split("/").pop() ||
        "";

      if (item.id && fileName) {
        const doc: IndexDoc = {
          id: String(item.id),
          fileName,
          fileUrl,
          effectiveFileUrl,
          slScope: item.slScope == null ? null : String(item.slScope),
          relativePath: null,
          spItemId: item.spItemId ? String(item.spItemId) : null,
        };
        doc.relativePath = resolveIndexRelativePath(
          doc,
          deptSiteUrl,
          deptBaseFolder,
          globalCommon
        );
        docs.push(doc);
      }
    }

    skip += items.length;
    if (items.length < top) break;
  }

  return docs;
}

async function getIndexDocsGlobalCommon(
  globalCommon: GlobalCommonConfig
): Promise<IndexDoc[]> {
  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;

  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars");
  }

  const docs: IndexDoc[] = [];
  let skip = 0;
  const top = 200;

  while (true) {
    const res = await fetch(
      `${endpoint}/indexes/${indexName}/docs?api-version=2024-07-01` +
        `&$select=id,metadata,fileUrl,effectiveFileUrl,dept,slScope,spItemId` +
        `&$filter=slScope eq 'global_common' and isSlDoc eq true` +
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
      const fileUrl = String(item.fileUrl ?? "");
      const effectiveFileUrl = String(item.effectiveFileUrl ?? "");
      const fileName =
        String(item.metadata ?? "").trim() ||
        safeDecodeURIComponent(effectiveFileUrl).split("/").pop() ||
        safeDecodeURIComponent(fileUrl).split("/").pop() ||
        "";

      if (item.id && fileName) {
        const doc: IndexDoc = {
          id: String(item.id),
          fileName,
          fileUrl,
          effectiveFileUrl,
          slScope: item.slScope == null ? null : String(item.slScope),
          relativePath: null,
          spItemId: item.spItemId ? String(item.spItemId) : null,
        };
        doc.relativePath = resolveIndexRelativePath(
          doc,
          globalCommon.siteUrl,
          globalCommon.folder,
          null
        );
        docs.push(doc);
      }
    }

    skip += items.length;
    if (items.length < top) break;
  }

  return docs;
}

async function deleteIndexDocs(ids: string[]): Promise<void> {
  if (ids.length === 0) return;

  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;

  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars");
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
  console.log(`[SL sync] Deleted ${ids.length} index docs`);
}

async function updateIndexDocs(
  updates: Array<{
    id: string;
    effectiveFileUrl: string;
    slScope?: ScopeKind;
    slOwner?: string | null;
  }>
): Promise<void> {
  if (updates.length === 0) return;

  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;

  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars");
  }

  const body = {
    value: updates.map(({ id, effectiveFileUrl, slScope, slOwner }) => {
      const doc: any = {
        "@search.action": "merge",
        id,
        effectiveFileUrl,
      };
      if (slScope !== undefined) doc.slScope = slScope;
      if (slOwner !== undefined) doc.slOwner = slOwner;
      return doc;
    }),
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

  if (!res.ok) throw new Error(`Index update failed: ${await res.text()}`);
  console.log(`[SL sync] Updated ${updates.length} docs`);
}

export async function runSlSync({
  accessToken,
  apply = false,
}: RunSlSyncParams): Promise<SlSyncResult> {
  const results: Record<string, SlSyncDeptResult> = {};

  for (const dept of getAllowedDepts()) {
    try {
      const { siteUrl, driveName, folder: baseFolder } = getDeptConfig(dept);
      const globalCommon = getGlobalCommonConfig();

      const { inventory, fetchFailed, rootMissing } = await getSpFileItems(
        accessToken,
        siteUrl,
        driveName,
        baseFolder
      );

      console.log(
        `[SL sync] dept=${dept} SP fileItems=${inventory.allItems.length} fetchFailed=${fetchFailed} rootMissing=${rootMissing}`
      );

      if (fetchFailed) {
        if (rootMissing) {
          // ベースフォルダ自体が存在しない → SP上にファイルは0件確定
          // インデックス上の孤立ドキュメントをすべて削除する
          const indexDocs = await getIndexDocs(dept, siteUrl, baseFolder, globalCommon);
          const orphanIds = indexDocs
            .filter((doc) => doc.slScope !== "global_common")
            .map((doc) => doc.id);
          console.log(
            `[SL sync] dept=${dept} base folder missing, orphans=${orphanIds.length} apply=${apply}`
          );
          if (apply && orphanIds.length > 0) {
            await deleteIndexDocs(orphanIds);
          }
          results[dept] = {
            spFileNames: 0,
            indexDocs: indexDocs.length,
            deleted: apply ? orphanIds.length : 0,
            orphanIds,
          };
        } else {
          // 子フォルダの一時的な 404 → 安全のためスキップ
          results[dept] = {
            spFileNames: 0,
            indexDocs: "unknown",
            deleted: 0,
            skipped: "sp_fetch_failed",
          };
        }
        continue;
      }

      const indexDocs = await getIndexDocs(dept, siteUrl, baseFolder, globalCommon);
      console.log(`[SL sync] dept=${dept} indexed docs=${indexDocs.length}`);

      const matchedDocs = indexDocs.map((doc) => ({
        doc,
        spItem: findMatchingSpItem(doc, inventory),
      }));

      // scan 範囲外（基点フォルダより上位など）へ移動したファイルを Graph API で直接追跡
      const scanMissed = matchedDocs.filter(
        (entry) => !entry.spItem && entry.doc.spItemId && entry.doc.slScope !== "global_common"
      );
      if (scanMissed.length > 0) {
        console.log(`[SL sync] dept=${dept} looking up ${scanMissed.length} scan-missed items by spItemId`);
        for (const entry of scanMissed) {
          const found = await lookupSpItemByIdInDrive(
            accessToken,
            siteUrl,
            driveName,
            entry.doc.spItemId!
          );
          if (found) {
            entry.spItem = found;
            console.log(`[SL sync] dept=${dept} found outside scan scope: ${entry.doc.fileName} → ${found.webUrl}`);
          }
        }
      }

      const orphanIds = matchedDocs
        .filter(({ doc }) => doc.slScope !== "global_common")
        .filter(({ spItem }) => !spItem)
        .map(({ doc }) => doc.id);

      console.log(`[SL sync] dept=${dept} orphans=${orphanIds.length} apply=${apply}`);

      const docUpdates = matchedDocs
        .filter((entry): entry is { doc: IndexDoc; spItem: SpFileItem } => Boolean(entry.spItem))
        .filter(({ doc, spItem }) => {
          const desiredScope = resolveScopeFromLocation({
            webUrl: spItem.webUrl,
            sourceSiteUrl: spItem.sourceSiteUrl,
            deptSiteUrl: siteUrl,
            deptBaseFolder: baseFolder,
            globalCommonSiteUrl: globalCommon?.siteUrl ?? null,
            globalCommonFolder: globalCommon?.folder ?? null,
          });

          return doc.effectiveFileUrl !== spItem.webUrl || doc.slScope !== desiredScope;
        })
        .map(({ doc, spItem }) => {
          const desiredScope = resolveScopeFromLocation({
            webUrl: spItem.webUrl,
            sourceSiteUrl: spItem.sourceSiteUrl,
            deptSiteUrl: siteUrl,
            deptBaseFolder: baseFolder,
            globalCommonSiteUrl: globalCommon?.siteUrl ?? null,
            globalCommonFolder: globalCommon?.folder ?? null,
          });

          return {
            id: doc.id,
            effectiveFileUrl: spItem.webUrl,
            slScope: desiredScope,
            ...(desiredScope === "personal" ? {} : { slOwner: null }),
          };
        });

      if (apply) {
        await deleteIndexDocs(orphanIds);
        await updateIndexDocs(docUpdates);
      }

      results[dept] = {
        spFileNames: inventory.allItems.length,
        indexDocs: indexDocs.length,
        deleted: apply ? orphanIds.length : 0,
        urlUpdated: docUpdates.length,
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

  try {
    const globalCommon = getGlobalCommonConfig();
    if (globalCommon) {
      const globalScan = await getSpFileItemsForFolder(
        accessToken,
        globalCommon.siteUrl,
        globalCommon.driveName,
        globalCommon.folder
      );

      if (!globalScan.fetchFailed) {
        const globalIndexDocs = await getIndexDocsGlobalCommon(globalCommon);
        const matchedDocs = globalIndexDocs.map((doc) => ({
          doc,
          spItem: findMatchingSpItem(doc, globalScan.inventory),
        }));

        const globalOrphanIds = matchedDocs
          .filter(({ spItem }) => !spItem)
          .map(({ doc }) => doc.id);

        if (apply && globalOrphanIds.length > 0) {
          await deleteIndexDocs(globalOrphanIds);
        }

        results["global_common"] = {
          spFileNames: globalScan.inventory.allItems.length,
          indexDocs: globalIndexDocs.length,
          deleted: apply ? globalOrphanIds.length : 0,
          orphanIds: globalOrphanIds,
        };
      } else {
        results["global_common"] = {
          spFileNames: 0,
          indexDocs: 0,
          deleted: 0,
          skipped: "sp_fetch_failed",
        };
      }
    }
  } catch (globalErr: any) {
    console.error(`[SL sync] global_common error:`, globalErr);
    results["global_common"] = {
      spFileNames: 0,
      indexDocs: 0,
      deleted: 0,
      error: String(globalErr?.message ?? globalErr),
    };
  }

  return {
    ok: true,
    mode: apply ? "apply" : "dry-run",
    results,
  };
}
