// src/lib/sl-sync.ts

import { getDeptConfig, getAllowedDepts } from "@/lib/sl-dept";

export type SpFileItem = {
  id: string;
  webUrl: string;
  sourceSiteUrl: string;
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
  effectiveFileUrl: string;
  slScope: string | null;
};

type ScopeKind = "global_common" | "dept_common" | "personal";

type GlobalCommonConfig = {
  siteUrl: string;
  driveName: string;
  folder: string;
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

  return {
    siteUrl,
    driveName,
    folder,
  };
}

function deriveDeptCommonFolder(baseFolder: string): string {
  return `${normalizeFolderPath(baseFolder)}/Common`;
}

function isWithinFolderByWebUrl(webUrl: string, folderPath: string): boolean {
  const decodedPathname = getDecodedPathnameFromWebUrl(webUrl).toLowerCase();
  const normalizedFolder = normalizeFolderPath(folderPath).toLowerCase();

  return (
    decodedPathname.includes(`/${normalizedFolder.toLowerCase()}/`) ||
    decodedPathname.endsWith(`/${normalizedFolder.toLowerCase()}`)
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

async function collectFileItemsRecursive(
  accessToken: string,
  driveId: string,
  currentFolderPath: string,
  sourceSiteUrl: string,
  fileItems: Map<string, SpFileItem>
): Promise<{ fetchFailed: boolean }> {
  const encoded = encodeGraphPath(currentFolderPath);
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
        fileItems.set(String(item.name), {
          id: String(item.id ?? ""),
          webUrl: String(item.webUrl ?? ""),
          sourceSiteUrl,
        });
      } else if (item?.folder && item?.name) {
        const child = await collectFileItemsRecursive(
          accessToken,
          driveId,
          `${currentFolderPath}/${item.name}`,
          sourceSiteUrl,
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

async function getSpFileItemsForFolder(
  accessToken: string,
  siteUrl: string,
  driveName: string,
  folderPath: string
): Promise<{ fileItems: Map<string, SpFileItem>; fetchFailed: boolean }> {
  const siteId = await resolveSiteId(accessToken, siteUrl);
  const driveId = await resolveDriveId(accessToken, siteId, driveName);

  const fileItems = new Map<string, SpFileItem>();
  const { fetchFailed } = await collectFileItemsRecursive(
    accessToken,
    driveId,
    folderPath,
    siteUrl,
    fileItems
  );

  console.log(
    `[SL sync] SP recursive scan: site="${siteUrl}" folder="${folderPath}" total=${fileItems.size} fetchFailed=${fetchFailed}`
  );

  return { fileItems, fetchFailed };
}

async function getSpFileItems(
  accessToken: string,
  deptSiteUrl: string,
  deptDriveName: string,
  deptBaseFolder: string
): Promise<{ fileItems: Map<string, SpFileItem>; fetchFailed: boolean }> {
  const merged = new Map<string, SpFileItem>();

  const deptScan = await getSpFileItemsForFolder(
    accessToken,
    deptSiteUrl,
    deptDriveName,
    deptBaseFolder
  );

  if (deptScan.fetchFailed) {
    return { fileItems: merged, fetchFailed: true };
  }

  deptScan.fileItems.forEach((item, name) => {
    merged.set(name, item);
  });

  const globalCommon = getGlobalCommonConfig();
  if (globalCommon) {
    const globalScan = await getSpFileItemsForFolder(
      accessToken,
      globalCommon.siteUrl,
      globalCommon.driveName,
      globalCommon.folder
    );

    if (!globalScan.fetchFailed) {
      globalScan.fileItems.forEach((item, name) => {
        merged.set(name, item);
      });
    } else {
      console.warn(
        `[SL sync] Global common scan skipped: ${globalCommon.siteUrl} / ${globalCommon.folder}`
      );
    }
  }

  console.log(
    `[SL sync] SP merged scan: deptBaseFolder="${deptBaseFolder}" total=${merged.size}`
  );

  return { fileItems: merged, fetchFailed: false };
}

async function getIndexDocs(dept: string): Promise<IndexDoc[]> {
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
        `&$select=id,fileUrl,effectiveFileUrl,dept,slScope` +
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
      const rawUrl: string = item.fileUrl ?? "";
      const decoded = safeDecodeURIComponent(rawUrl);
      const fileName = decoded.split("/").pop() ?? "";

      if (item.id && fileName) {
        docs.push({
          id: String(item.id),
          fileName,
          effectiveFileUrl: String(item.effectiveFileUrl ?? ""),
          slScope: item.slScope == null ? null : String(item.slScope),
        });
      }
    }

    skip += items.length;
    if (items.length < top) break;
  }

  return docs;
}

// ★ 追加: global_common専用のIndex取得
async function getIndexDocsGlobalCommon(): Promise<IndexDoc[]> {
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
        `&$select=id,fileUrl,effectiveFileUrl,dept,slScope` +
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
      const rawUrl: string = item.fileUrl ?? "";
      const decoded = safeDecodeURIComponent(rawUrl);
      const fileName = decoded.split("/").pop() ?? "";

      if (item.id && fileName) {
        docs.push({
          id: String(item.id),
          fileName,
          effectiveFileUrl: String(item.effectiveFileUrl ?? ""),
          slScope: item.slScope == null ? null : String(item.slScope),
        });
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
      const globalCommon = getGlobalCommonConfig();

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

      // global_common はこの dept sync で削除しない
      const orphanIds = indexDocs
        .filter((doc) => doc.slScope !== "global_common")
        .filter((doc) => !spFileItems.has(doc.fileName))
        .map((doc) => doc.id);

      const docUpdates = indexDocs
        .filter((doc) => {
          const spItem = spFileItems.get(doc.fileName);
          if (!spItem) return false;
          if (!spItem.webUrl) return false;

          const desiredScope = resolveScopeFromLocation({
            webUrl: spItem.webUrl,
            sourceSiteUrl: spItem.sourceSiteUrl,
            deptSiteUrl: siteUrl,
            deptBaseFolder: baseFolder,
            globalCommonSiteUrl: globalCommon?.siteUrl ?? null,
            globalCommonFolder: globalCommon?.folder ?? null,
          });

          const urlChanged = doc.effectiveFileUrl !== spItem.webUrl;
          const scopeChanged = doc.slScope !== desiredScope;

          return urlChanged || scopeChanged;
        })
        .map((doc) => {
          const spItem = spFileItems.get(doc.fileName)!;

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

      if (orphanIds.length > 0) {
        console.log(`[SL sync] dept=${dept} orphans=${orphanIds.length}`);
      }
      if (docUpdates.length > 0) {
        console.log(`[SL sync] dept=${dept} updates=${docUpdates.length}`);
      }

      if (apply) {
        await deleteIndexDocs(orphanIds);
        await updateIndexDocs(docUpdates);
      }

      results[dept] = {
        spFileNames: spFileItems.size,
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

  // ★ 追加: global_common 専用 orphan 削除
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
        const globalIndexDocs = await getIndexDocsGlobalCommon();
        console.log(`[SL sync] global_common indexed docs=${globalIndexDocs.length}`);

        const globalOrphanIds = globalIndexDocs
          .filter((doc) => !globalScan.fileItems.has(doc.fileName))
          .map((doc) => doc.id);

        if (globalOrphanIds.length > 0) {
          console.log(`[SL sync] global_common orphans=${globalOrphanIds.length}`);
        }

        if (apply && globalOrphanIds.length > 0) {
          await deleteIndexDocs(globalOrphanIds);
        }

        results["global_common"] = {
          spFileNames: globalScan.fileItems.size,
          indexDocs: globalIndexDocs.length,
          deleted: apply ? globalOrphanIds.length : 0,
          orphanIds: globalOrphanIds,
        };
      } else {
        console.warn(`[SL sync] global_common SP scan failed — skipping orphan delete`);
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