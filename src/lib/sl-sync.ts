// src/lib/sl-sync.ts

import { createHash, randomUUID } from "crypto";
import { getAllowedDepts, getDeptConfig } from "@/lib/sl-dept";
import { OpenAIEmbeddingInstance } from "@/features/common/services/openai";
import { extractIndexDocumentFromBuffer } from "./document-extract";

export type SpFileItem = {
  id: string;
  name: string;
  webUrl: string;
  sourceSiteUrl: string;
  relativePath: string;
  contentTag: string;
  lastModifiedAt: string | null;
};

export type SlSyncDeptResult = {
  spFileNames: number;
  indexDocs: number | "unknown";
  orphanIds?: string[];
  deleted: number;
  urlUpdated?: number;
  newIndexed?: number;
  newSkipped?: number;
  reindexCandidates?: number;
  reindexCandidateNames?: string[];
  reindexed?: number;
  reindexFailed?: number;
  unindexedCount?: number;
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
  indexNew?: boolean;
  batchSize?: number;
  reindexOnly?: boolean;
};

type IndexDoc = {
  id: string;
  fileName: string;
  fileUrl: string;
  effectiveFileUrl: string;
  slScope: string | null;
  relativePath: string | null;
  storedRelativePath: string | null;
  spItemId: string | null;
  spContentTag: string | null;
  spLastModifiedAt: string | null;
  indexingVersion: number | null;
};

type NewIndexDoc = {
  id: string;
  pageContent: string;
  embedding: number[];
  metadata: string;
  fileUrl: string;
  effectiveFileUrl: string;
  chatThreadId: string;
  user: string;
  dept: string;
  isSlDoc: true;
  slScope: "global_common" | "dept_common" | "personal";
  slOwner: string | null;
  spItemId: string | null;
  relativePath?: string | null;
  spContentTag?: string | null;
  spLastModifiedAt?: string | null;
  indexingVersion?: number;
  chunkIndex?: number;
  documentChunkCount?: number;
  pageStart?: number;
  pageEnd?: number;
  documentPageCount?: number;
};

type MatchedIndexDoc = {
  doc: IndexDoc;
  spItem: SpFileItem | null;
  lookupFailed?: boolean;
};

type ReindexCandidate = {
  item: SpFileItem;
  oldDocIds: string[];
  reason: "content_changed" | "legacy_missing_tag" | "page_metadata_missing";
};

type ReindexJob = {
  resultKey: string;
  candidate: ReindexCandidate;
  dept: string;
  siteUrl: string;
  driveName: string;
  baseFolder: string;
  globalCommon?: GlobalCommonConfig | null;
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

function hashValue(value: string): string {
  return createHash("sha256").update(value.trim().toLowerCase()).digest("hex");
}

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
  const normalized = (path ?? "")
    .replace(/\\/g, "/")
    .replace(/^\/+/, "")
    .replace(/\/+$/, "");
  return normalized === "." ? "" : normalized;
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
  const base = normalizeFolderPath(baseFolder);
  const commonSub = process.env.SL_COMMON_SUBFOLDER ?? "Common";
  return base ? `${base}/${commonSub}` : commonSub;
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
  itemRelativePath?: string;
  globalCommonSiteUrl?: string | null;
  globalCommonFolder?: string | null;
}): ScopeKind {
  const {
    webUrl,
    sourceSiteUrl,
    deptSiteUrl,
    deptBaseFolder,
    itemRelativePath,
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

  // Files directly in baseFolder (not in a user subfolder) are dept_common
  if (itemRelativePath !== undefined) {
    const normalizedBase = normalizeFolderPath(deptBaseFolder).toLowerCase();
    const normalizedRel = normalizeFolderPath(itemRelativePath).toLowerCase();
    const rest = normalizedBase
      ? normalizedRel.startsWith(normalizedBase + "/")
        ? normalizedRel.slice(normalizedBase.length + 1)
        : ""
      : normalizedRel;
    if (!rest.includes("/")) return "dept_common";
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
): Promise<SpFileItem | null | "error"> {
  try {
    const siteId = await resolveSiteId(accessToken, siteUrl);
    const driveId = await resolveDriveId(accessToken, siteId, driveName);

    const res = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}` +
        `?$select=name,file,id,webUrl,parentReference,deleted,cTag,eTag,lastModifiedDateTime`,
      { headers: { Authorization: `Bearer ${accessToken}` }, cache: "no-store" }
    );

    if (res.status === 404) return null; // ファイルが実際に削除済み（ゴミ箱も空）
    if (!res.ok) {
      console.warn(`[SL sync] lookupSpItemById failed (${res.status}): ${await res.text()}`);
      return "error"; // スロットリング等の一時エラー → 孤立扱いしない
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
      contentTag: String(item.cTag ?? item.eTag ?? ""),
      lastModifiedAt: item.lastModifiedDateTime
        ? String(item.lastModifiedDateTime)
        : null,
    };
  } catch (e) {
    console.warn(`[SL sync] lookupSpItemByIdInDrive error:`, e);
    return "error"; // ネットワークエラー等 → 孤立扱いしない
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
  const select =
    "name,file,folder,id,webUrl,parentReference,cTag,eTag,lastModifiedDateTime";
  let nextUrl: string | null = encoded
    ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encoded}:/children?$select=${select}&$top=200`
    : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=${select}&$top=200`;

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
          contentTag: String(item.cTag ?? item.eTag ?? ""),
          lastModifiedAt: item.lastModifiedDateTime
            ? String(item.lastModifiedDateTime)
            : null,
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
  globalCommon?: GlobalCommonConfig | null,
  hasRelativePath = false,
  hasChangeTrackingFields = false,
  hasPageMetadataFields = false
): Promise<IndexDoc[]> {
  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;

  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars");
  }

  const selectFields = [
    "id",
    "metadata",
    "fileUrl",
    "effectiveFileUrl",
    "dept",
    "slScope",
    "spItemId",
    ...(hasRelativePath ? ["relativePath"] : []),
    ...(hasChangeTrackingFields
      ? ["spContentTag", "spLastModifiedAt"]
      : []),
    ...(hasPageMetadataFields ? ["indexingVersion"] : []),
  ].join(",");

  const docs: IndexDoc[] = [];
  let skip = 0;
  const top = 200;

  while (true) {
    const res = await fetch(
      `${endpoint}/indexes/${indexName}/docs?api-version=2024-07-01` +
        `&$select=${selectFields}` +
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
        const storedRelativePath = hasRelativePath && item.relativePath ? String(item.relativePath) : null;
        const doc: IndexDoc = {
          id: String(item.id),
          fileName,
          fileUrl,
          effectiveFileUrl,
          slScope: item.slScope == null ? null : String(item.slScope),
          relativePath: null,
          storedRelativePath,
          spItemId: item.spItemId ? String(item.spItemId) : null,
          spContentTag: hasChangeTrackingFields && item.spContentTag
            ? String(item.spContentTag)
            : null,
          spLastModifiedAt:
            hasChangeTrackingFields && item.spLastModifiedAt
              ? String(item.spLastModifiedAt)
              : null,
          indexingVersion:
            hasPageMetadataFields && Number.isInteger(item.indexingVersion)
              ? Number(item.indexingVersion)
              : null,
        };
        doc.relativePath = storedRelativePath ?? resolveIndexRelativePath(
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
  globalCommon: GlobalCommonConfig,
  hasRelativePath = false,
  hasChangeTrackingFields = false,
  hasPageMetadataFields = false
): Promise<IndexDoc[]> {
  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;

  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars");
  }

  const selectFields = [
    "id",
    "metadata",
    "fileUrl",
    "effectiveFileUrl",
    "dept",
    "slScope",
    "spItemId",
    ...(hasRelativePath ? ["relativePath"] : []),
    ...(hasChangeTrackingFields
      ? ["spContentTag", "spLastModifiedAt"]
      : []),
    ...(hasPageMetadataFields ? ["indexingVersion"] : []),
  ].join(",");

  const docs: IndexDoc[] = [];
  let skip = 0;
  const top = 200;

  while (true) {
    const res = await fetch(
      `${endpoint}/indexes/${indexName}/docs?api-version=2024-07-01` +
        `&$select=${selectFields}` +
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
        const storedRelativePath = hasRelativePath && item.relativePath ? String(item.relativePath) : null;
        const doc: IndexDoc = {
          id: String(item.id),
          fileName,
          fileUrl,
          effectiveFileUrl,
          slScope: item.slScope == null ? null : String(item.slScope),
          relativePath: null,
          storedRelativePath,
          spItemId: item.spItemId ? String(item.spItemId) : null,
          spContentTag: hasChangeTrackingFields && item.spContentTag
            ? String(item.spContentTag)
            : null,
          spLastModifiedAt:
            hasChangeTrackingFields && item.spLastModifiedAt
              ? String(item.spLastModifiedAt)
              : null,
          indexingVersion:
            hasPageMetadataFields && Number.isInteger(item.indexingVersion)
              ? Number(item.indexingVersion)
              : null,
        };
        doc.relativePath = storedRelativePath ?? resolveIndexRelativePath(
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

  for (let offset = 0; offset < ids.length; offset += 1000) {
    const batchIds = ids.slice(offset, offset + 1000);
    const res = await fetch(
      `${endpoint}/indexes/${indexName}/docs/index?api-version=2024-07-01`,
      {
        method: "POST",
        headers: { "api-key": apiKey, "Content-Type": "application/json" },
        body: JSON.stringify({
          value: batchIds.map((id) => ({ "@search.action": "delete", id })),
        }),
        cache: "no-store",
      }
    );

    await assertIndexActionSucceeded(res, "Delete");
  }
  console.log(`[SL sync] Deleted ${ids.length} index docs`);
}

async function ensureRelativePathField(): Promise<boolean> {
  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;
  if (!endpoint || !apiKey || !indexName) return false;

  const getRes = await fetch(
    `${endpoint}/indexes/${indexName}?api-version=2024-07-01`,
    { headers: { "api-key": apiKey }, cache: "no-store" }
  );
  if (!getRes.ok) {
    console.warn("[SL sync] ensureRelativePathField: failed to get index schema");
    return false;
  }

  const index = await getRes.json();
  const fields: any[] = index.fields ?? [];

  if (fields.some((f: any) => f.name === "relativePath")) {
    console.log("[SL sync] relativePath field already exists in index");
    return true;
  }

  // Only attempt to add the field when explicitly opted in via env var
  if (process.env.SL_AUTO_ENSURE_RELATIVE_PATH_FIELD !== "true") {
    console.warn("[SL sync] relativePath field missing; set SL_AUTO_ENSURE_RELATIVE_PATH_FIELD=true to add it automatically");
    return false;
  }

  fields.push({
    name: "relativePath",
    type: "Edm.String",
    searchable: true,
    filterable: true,
    retrievable: true,
    sortable: false,
    facetable: false,
  });
  index.fields = fields;

  const putRes = await fetch(
    `${endpoint}/indexes/${indexName}?api-version=2024-07-01`,
    {
      method: "PUT",
      headers: { "api-key": apiKey, "Content-Type": "application/json" },
      body: JSON.stringify(index),
      cache: "no-store",
    }
  );

  if (!putRes.ok) {
    console.error("[SL sync] ensureRelativePathField: failed to update schema:", await putRes.text());
    return false;
  }

  console.log("[SL sync] ensureRelativePathField: relativePath field added to index");
  return true;
}

async function hasSharePointChangeTrackingFields(): Promise<boolean> {
  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;
  if (!endpoint || !apiKey || !indexName) return false;

  const res = await fetch(
    `${endpoint}/indexes/${indexName}?api-version=2024-07-01`,
    { headers: { "api-key": apiKey }, cache: "no-store" }
  );
  if (!res.ok) {
    console.warn(
      "[SL sync] Failed to inspect index schema for SharePoint change tracking"
    );
    return false;
  }

  const index = await res.json();
  const fieldNames = new Set<string>(
    (index.fields ?? []).map((field: any) => String(field.name ?? ""))
  );
  const requiredFields = ["spContentTag", "spLastModifiedAt"];
  const missingFields = requiredFields.filter(
    (fieldName) => !fieldNames.has(fieldName)
  );
  const hasFields = missingFields.length === 0;

  if (!hasFields) {
    console.warn(
      `[SL sync] Change reindexing disabled for ${indexName}: missing fields ${missingFields.join(", ")}`
    );
  }
  return hasFields;
}

async function hasSharePointPageMetadataFields(): Promise<boolean> {
  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;
  if (!endpoint || !apiKey || !indexName) return false;

  const res = await fetch(
    `${endpoint}/indexes/${indexName}?api-version=2024-07-01`,
    { headers: { "api-key": apiKey }, cache: "no-store" }
  );
  if (!res.ok) {
    console.warn("[SL sync] Failed to inspect page metadata index fields");
    return false;
  }

  const index = await res.json();
  const fieldNames = new Set<string>(
    (index.fields ?? []).map((field: any) => String(field.name ?? ""))
  );
  const requiredFields = [
    "indexingVersion",
    "chunkIndex",
    "documentChunkCount",
    "pageStart",
    "pageEnd",
    "documentPageCount",
  ];
  const missingFields = requiredFields.filter((fieldName) => !fieldNames.has(fieldName));
  if (missingFields.length > 0) {
    console.warn(
      `[SL sync] Page-aware indexing disabled for ${indexName}: missing fields ${missingFields.join(", ")}`
    );
    return false;
  }
  return true;
}

async function addNewIndexDocs(docs: NewIndexDoc[]): Promise<void> {
  if (docs.length === 0) return;

  const endpoint = process.env.AZURE_SEARCH_ENDPOINT;
  const apiKey = process.env.AZURE_SEARCH_API_KEY;
  const indexName = process.env.AZURE_SEARCH_INDEX_NAME;

  if (!endpoint || !apiKey || !indexName) {
    throw new Error("Missing Azure Search env vars");
  }

  for (let offset = 0; offset < docs.length; offset += 500) {
    const batchDocs = docs.slice(offset, offset + 500);
    const res = await fetch(
      `${endpoint}/indexes/${indexName}/docs/index?api-version=2024-07-01`,
      {
        method: "POST",
        headers: { "api-key": apiKey, "Content-Type": "application/json" },
        body: JSON.stringify({
          value: batchDocs.map((doc) => ({
            "@search.action": "upload",
            ...doc,
          })),
        }),
        cache: "no-store",
      }
    );

    await assertIndexActionSucceeded(res, "Index upload");
  }
  console.log(`[SL sync] Indexed ${docs.length} new docs`);
}

async function assertIndexActionSucceeded(
  response: Response,
  operation: string
): Promise<void> {
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`${operation} failed: ${text}`);
  }

  if (!text) return;
  const payload = JSON.parse(text);
  const failures = (payload.value ?? []).filter(
    (result: any) => result.status === false
  );
  if (failures.length > 0) {
    const details = failures
      .slice(0, 5)
      .map((result: any) => `${result.key}: ${result.errorMessage ?? "failed"}`)
      .join("; ");
    throw new Error(
      `${operation} partially failed (${failures.length} documents): ${details}`
    );
  }
}

async function downloadSpFile(
  accessToken: string,
  driveId: string,
  itemId: string
): Promise<ArrayBuffer> {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
      cache: "no-store",
      redirect: "follow",
    }
  );
  if (!res.ok) throw new Error(`Download failed (${res.status}): ${await res.text()}`);
  return res.arrayBuffer();
}

function findUnindexedSpItems(
  inventory: SpInventory,
  indexDocs: IndexDoc[]
): SpFileItem[] {
  const indexedBySpItemId = new Set<string>();
  const indexedByRelPath = new Set<string>();

  for (const doc of indexDocs) {
    if (doc.spItemId) indexedBySpItemId.add(doc.spItemId);
    if (doc.relativePath) indexedByRelPath.add(doc.relativePath.toLowerCase());
  }

  console.log(
    `[SL sync] findUnindexedSpItems: indexDocs=${indexDocs.length} bySpItemId=${indexedBySpItemId.size} byRelPath=${indexedByRelPath.size}`
  );

  return inventory.allItems.filter((item) => {
    if (item.id && indexedBySpItemId.has(item.id)) return false;
    if (item.relativePath && indexedByRelPath.has(item.relativePath.toLowerCase())) return false;
    console.log(`[SL sync] UNINDEXED: name=${item.name} spItemId=${item.id} relPath=${item.relativePath}`);
    return true;
  });
}

function resolvePersonalOwnerHash(
  spItem: SpFileItem,
  baseFolder: string
): string | null {
  const domain = (process.env.SL_PERSONAL_EMAIL_DOMAIN ?? "").trim();
  if (!domain) {
    console.warn(
      "[SL sync] SL_PERSONAL_EMAIL_DOMAIN is not set — personal slOwner cannot be determined"
    );
    return null;
  }

  const normalizedBase = normalizeFolderPath(baseFolder).toLowerCase();
  const normalizedPath = normalizeFolderPath(spItem.relativePath).toLowerCase();

  let rest: string;
  if (normalizedBase) {
    if (!normalizedPath.startsWith(normalizedBase + "/")) return null;
    rest = normalizedPath.slice(normalizedBase.length + 1);
  } else {
    rest = normalizedPath;
  }

  const firstSegment = rest.split("/")[0];
  if (!firstSegment) return null;

  const commonSubfolder = (process.env.SL_COMMON_SUBFOLDER ?? "common").toLowerCase();
  if (firstSegment.toLowerCase() === commonSubfolder) return null;

  // フォルダ名が既知ユーザーのメールローカルパートでない場合は null を返す。
  // 例: SL/新フォルダ/ のような任意フォルダ → dept_common として扱わせる。
  const targetEmail = `${firstSegment}@${domain}`;
  const isKnownUser = getAllowedDepts().some((dept) => {
    const key = `SL_DEPT_BY_EMAIL_${dept.toUpperCase()}`;
    return (process.env[key] ?? "")
      .split(",")
      .some((s) => s.trim().toLowerCase() === targetEmail);
  });
  if (!isKnownUser) return null;

  return hashValue(`${firstSegment}@${domain}`);
}

async function indexNewSpFiles(params: {
  accessToken: string;
  dept: string;
  siteUrl: string;
  driveName: string;
  baseFolder: string;
  unindexedItems: SpFileItem[];
  batchSize: number;
  globalCommon?: GlobalCommonConfig | null;
  hasRelativePath?: boolean;
  hasChangeTrackingFields?: boolean;
  hasPageMetadataFields?: boolean;
}): Promise<{ indexed: number; skipped: number }> {
  const {
    accessToken,
    dept,
    siteUrl,
    driveName,
    baseFolder,
    unindexedItems,
    batchSize,
    globalCommon,
    hasRelativePath = false,
    hasChangeTrackingFields = false,
    hasPageMetadataFields = false,
  } = params;

  const batch = unindexedItems.slice(0, batchSize);
  if (batch.length === 0) return { indexed: 0, skipped: 0 };

  const siteId = await resolveSiteId(accessToken, siteUrl);
  const driveId = await resolveDriveId(accessToken, siteId, driveName);
  const openai = OpenAIEmbeddingInstance();

  let indexed = 0;
  let skipped = 0;

  for (const item of batch) {
    try {
      console.log(`[SL sync] Indexing new SP file: ${item.name} (id=${item.id})`);

      const buffer = await downloadSpFile(accessToken, driveId, item.id);

      const extractedDocument = await extractIndexDocumentFromBuffer(buffer, item.name);
      const allChunks = extractedDocument.chunks;
      if (allChunks.length === 0) {
        console.warn(`[SL sync] No text extracted from ${item.name}, skipping`);
        skipped++;
        continue;
      }

      const scope = resolveScopeFromLocation({
        webUrl: item.webUrl,
        sourceSiteUrl: item.sourceSiteUrl,
        deptSiteUrl: siteUrl,
        deptBaseFolder: baseFolder,
        itemRelativePath: item.relativePath,
        globalCommonSiteUrl: globalCommon?.siteUrl ?? null,
        globalCommonFolder: globalCommon?.folder ?? null,
      });

      const slOwner =
        scope === "personal" ? resolvePersonalOwnerHash(item, baseFolder) : null;

      // フォルダ名が既知ユーザーに対応しない場合は dept_common に格下げしてインデックス
      const effectiveScope: ScopeKind = scope === "personal" && slOwner === null
        ? "dept_common"
        : scope;

      if (effectiveScope !== scope) {
        console.log(`[SL sync] ${item.name}: unknown personal folder → treating as dept_common`);
      }

      const EMBED_BATCH = 50;
      const docsToIndex: NewIndexDoc[] = [];
      let firstEmbeddingDim = 0;

      for (let b = 0; b < allChunks.length; b += EMBED_BATCH) {
        const batchChunks = allChunks.slice(b, b + EMBED_BATCH);
        const embeddingRes = await openai.embeddings.create({
          input: batchChunks.map((chunk) => chunk.content),
          model: "",
        });

        const batchEmbeddingDim = embeddingRes.data[0]?.embedding?.length ?? 0;
        if (firstEmbeddingDim === 0) firstEmbeddingDim = batchEmbeddingDim;

        console.log(
          `[SL sync] Embedding batch: offset=${b} chunks=${batchChunks.length} embData=${embeddingRes.data.length} firstEmbLen=${batchEmbeddingDim}`
        );

        for (let j = 0; j < batchChunks.length; j++) {
          docsToIndex.push({
            id: randomUUID(),
            pageContent: batchChunks[j].content,
            embedding: embeddingRes.data[j]?.embedding ?? [],
            metadata: item.name,
            fileUrl: item.webUrl,
            effectiveFileUrl: item.webUrl,
            chatThreadId: "sl-auto",
            user: "",
            dept: dept.toLowerCase(),
            isSlDoc: true,
            slScope: effectiveScope,
            slOwner: slOwner ?? null,
            spItemId: item.id,
            ...(hasPageMetadataFields && extractedDocument.hasPageMetadata
              ? {
                  indexingVersion: 2,
                  chunkIndex: batchChunks[j].chunkIndex,
                  documentChunkCount: allChunks.length,
                  pageStart: batchChunks[j].pageStart ?? undefined,
                  pageEnd: batchChunks[j].pageEnd ?? undefined,
                  documentPageCount: extractedDocument.pageCount ?? undefined,
                }
              : {}),
            ...(hasRelativePath ? { relativePath: item.relativePath ?? null } : {}),
            ...(hasChangeTrackingFields
              ? {
                  spContentTag: item.contentTag || null,
                  spLastModifiedAt: item.lastModifiedAt,
                }
              : {}),
          });
        }
      }

      const emptyEmbCount = docsToIndex.filter((d) => d.embedding.length === 0).length;
      if (emptyEmbCount > 0) {
        console.warn(`[SL sync] WARNING: ${emptyEmbCount}/${docsToIndex.length} docs have empty embedding for ${item.name}`);
      }

      await addNewIndexDocs(docsToIndex);
      indexed++;
      console.log(
        `[SL sync] Indexed ${item.name}: scope=${effectiveScope} chunks=${allChunks.length} embeddingDim=${firstEmbeddingDim}`
      );
    } catch (e) {
      console.error(`[SL sync] Failed to index ${item.name}:`, e);
      // sentinel エントリを登録して次回Syncで再キューされないようにする
      try {
        const sentinelText = `[INDEXING_FAILED] ${item.name}`;
        const sentinelEmbRes = await openai.embeddings.create({
          input: [sentinelText],
          model: "",
        });
        const sentinelEmbedding = sentinelEmbRes.data[0]?.embedding ?? [];
        const sentinelId = `sl_error_${hashValue(item.id || item.name)}`;
        await addNewIndexDocs([{
          id: sentinelId,
          pageContent: sentinelText,
          embedding: sentinelEmbedding,
          metadata: JSON.stringify({ indexingError: true, fileName: item.name }),
          fileUrl: item.webUrl,
          effectiveFileUrl: item.webUrl,
          chatThreadId: "system",
          user: "system",
          dept: dept,
          isSlDoc: true,
          slScope: "dept_common",
          slOwner: null,
          spItemId: item.id,
          relativePath: item.relativePath ?? null,
          ...(hasChangeTrackingFields
            ? {
                spContentTag: item.contentTag || null,
                spLastModifiedAt: item.lastModifiedAt,
              }
            : {}),
        }]);
        console.warn(`[SL sync] Sentinel registered for ${item.name} to prevent re-queue`);
      } catch (sentinelErr) {
        console.error(`[SL sync] Sentinel registration also failed for ${item.name}:`, sentinelErr);
      }
      skipped++;
    }
  }

  return { indexed, skipped };
}

function findReindexCandidates(
  matchedDocs: MatchedIndexDoc[],
  hasPageMetadataFields: boolean
): ReindexCandidate[] {
  const groups = new Map<
    string,
    { item: SpFileItem; docs: IndexDoc[] }
  >();

  for (const entry of matchedDocs) {
    if (!entry.spItem?.id || !entry.spItem.contentTag) continue;
    const current = groups.get(entry.spItem.id) ?? {
      item: entry.spItem,
      docs: [],
    };
    current.docs.push(entry.doc);
    groups.set(entry.spItem.id, current);
  }

  const changed: ReindexCandidate[] = [];
  const legacy: ReindexCandidate[] = [];

  for (const { item, docs } of Array.from(groups.values())) {
    const storedTags = docs
      .map((doc) => doc.spContentTag)
      .filter((tag): tag is string => Boolean(tag));
    const hasMissingTag = storedTags.length !== docs.length;
    const contentChanged = storedTags.some((tag) => tag !== item.contentTag);
    const pageMetadataMissing =
      hasPageMetadataFields &&
      item.name.toLowerCase().endsWith(".pdf") &&
      docs.some((doc) => doc.indexingVersion !== 2);

    if (!hasMissingTag && !contentChanged && !pageMetadataMissing) continue;

    const candidate: ReindexCandidate = {
      item,
      oldDocIds: docs.map((doc) => doc.id),
      reason: contentChanged
        ? "content_changed"
        : pageMetadataMissing
          ? "page_metadata_missing"
          : "legacy_missing_tag",
    };
    if (contentChanged) changed.push(candidate);
    else legacy.push(candidate);
  }

  // Real SharePoint changes take priority over the gradual legacy migration.
  const modifiedAt = (candidate: ReindexCandidate) => {
    const value = Date.parse(candidate.item.lastModifiedAt ?? "");
    return Number.isNaN(value) ? 0 : value;
  };
  const newestFirst = (a: ReindexCandidate, b: ReindexCandidate) =>
    modifiedAt(b) - modifiedAt(a);
  return [
    ...changed.sort(newestFirst),
    ...legacy.sort(newestFirst),
  ];
}

async function reindexSpFiles(params: {
  accessToken: string;
  dept: string;
  siteUrl: string;
  driveName: string;
  baseFolder: string;
  candidates: ReindexCandidate[];
  batchSize: number;
  globalCommon?: GlobalCommonConfig | null;
  hasRelativePath: boolean;
  hasChangeTrackingFields: boolean;
  hasPageMetadataFields: boolean;
}): Promise<{ reindexed: number; failed: number }> {
  const batch = params.candidates.slice(0, params.batchSize);
  let reindexed = 0;
  let failed = 0;

  for (const candidate of batch) {
    console.log(
      `[SL sync] Reindexing ${candidate.item.name}: reason=${candidate.reason} oldChunks=${candidate.oldDocIds.length}`
    );
    const result = await indexNewSpFiles({
      accessToken: params.accessToken,
      dept: params.dept,
      siteUrl: params.siteUrl,
      driveName: params.driveName,
      baseFolder: params.baseFolder,
      unindexedItems: [candidate.item],
      batchSize: 1,
      globalCommon: params.globalCommon,
      hasRelativePath: params.hasRelativePath,
      hasChangeTrackingFields: params.hasChangeTrackingFields,
      hasPageMetadataFields: params.hasPageMetadataFields,
    });

    if (result.indexed !== 1) {
      failed++;
      continue;
    }

    // Keep the previous searchable chunks until the replacement is complete.
    await deleteIndexDocs(candidate.oldDocIds);
    reindexed++;
  }

  return { reindexed, failed };
}

async function updateIndexDocs(
  updates: Array<{
    id: string;
    effectiveFileUrl: string;
    slScope?: ScopeKind;
    slOwner?: string | null;
    relativePath?: string | null;
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
    value: updates.map(({ id, effectiveFileUrl, slScope, slOwner, relativePath }) => {
      const doc: any = {
        "@search.action": "merge",
        id,
        effectiveFileUrl,
      };
      if (slScope !== undefined) doc.slScope = slScope;
      if (slOwner !== undefined) doc.slOwner = slOwner;
      if (relativePath !== undefined) doc.relativePath = relativePath;
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
  indexNew = false,
  batchSize = 5,
  reindexOnly = false,
}: RunSlSyncParams): Promise<SlSyncResult> {
  const hasRelativePath = await ensureRelativePathField();
  const hasChangeTrackingFields =
    await hasSharePointChangeTrackingFields();
  const hasPageMetadataFields =
    await hasSharePointPageMetadataFields();

  const results: Record<string, SlSyncDeptResult> = {};
  const reindexJobs: ReindexJob[] = [];

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
          const indexDocs = await getIndexDocs(
            dept,
            siteUrl,
            baseFolder,
            globalCommon,
            hasRelativePath,
            hasChangeTrackingFields,
            hasPageMetadataFields
          );
          const orphanIds = indexDocs
            .filter((doc) => doc.slScope !== "global_common")
            .map((doc) => doc.id);
          console.log(
            `[SL sync] dept=${dept} base folder missing, orphans=${orphanIds.length} apply=${apply}`
          );
          if (apply && !reindexOnly && orphanIds.length > 0) {
            await deleteIndexDocs(orphanIds);
          }
          results[dept] = {
            spFileNames: 0,
            indexDocs: indexDocs.length,
            deleted: apply && !reindexOnly ? orphanIds.length : 0,
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

      const indexDocs = await getIndexDocs(
        dept,
        siteUrl,
        baseFolder,
        globalCommon,
        hasRelativePath,
        hasChangeTrackingFields,
        hasPageMetadataFields
      );
      console.log(`[SL sync] dept=${dept} indexed docs=${indexDocs.length}`);

      const matchedDocs: MatchedIndexDoc[] = indexDocs.map((doc) => ({
        doc,
        spItem: findMatchingSpItem(doc, inventory),
      }));

      // scan 範囲外（基点フォルダより上位など）へ移動したファイルを Graph API で直接追跡
      const scanMissed = matchedDocs.filter(
        (entry) => !entry.spItem && entry.doc.spItemId && entry.doc.slScope !== "global_common"
      );
      if (scanMissed.length > 0) {
        const uniqueSpItemIds = Array.from(new Set(scanMissed.map((e) => e.doc.spItemId!)));
        console.log(`[SL sync] dept=${dept} looking up ${uniqueSpItemIds.length} unique spItemIds (${scanMissed.length} docs)`);
        const lookupCache = new Map<string, SpFileItem | null | "error">();
        for (const spItemId of uniqueSpItemIds) {
          lookupCache.set(spItemId, await lookupSpItemByIdInDrive(accessToken, siteUrl, driveName, spItemId));
        }
        for (const entry of scanMissed) {
          const found = lookupCache.get(entry.doc.spItemId!);
          if (found === "error") {
            entry.lookupFailed = true;
            console.warn(`[SL sync] dept=${dept} lookup error for ${entry.doc.fileName}, skipping orphan`);
          } else if (found) {
            entry.spItem = found;
            console.log(`[SL sync] dept=${dept} found outside scan scope: ${entry.doc.fileName} → ${found.webUrl}`);
          }
        }
      }

      const orphanIds = matchedDocs
        .filter(({ doc }) => doc.slScope !== "global_common")
        .filter(({ spItem, lookupFailed }) => !spItem && !lookupFailed)
        .map(({ doc }) => doc.id);

      const gcExcluded = matchedDocs.filter(({ doc, spItem }) => doc.slScope === "global_common" && !spItem).length;
      if (gcExcluded > 0) {
        console.log(`[SL sync] dept=${dept} global_common docs outside scan (excluded from orphan): ${gcExcluded}`);
      }
      console.log(`[SL sync] dept=${dept} orphans=${orphanIds.length} apply=${apply}`);

      // global_common ドキュメントは global_common ブロックが管理する。
      // dept ループで slScope を上書きするとドキュメントが消失するため除外する。
      const docUpdates = matchedDocs
        .filter((entry): entry is { doc: IndexDoc; spItem: SpFileItem } => Boolean(entry.spItem))
        .filter(({ doc }) => doc.slScope !== "global_common")
        .filter(({ doc, spItem }) => {
          const desiredScope = resolveScopeFromLocation({
            webUrl: spItem.webUrl,
            sourceSiteUrl: spItem.sourceSiteUrl,
            deptSiteUrl: siteUrl,
            deptBaseFolder: baseFolder,
            itemRelativePath: spItem.relativePath,
            globalCommonSiteUrl: globalCommon?.siteUrl ?? null,
            globalCommonFolder: globalCommon?.folder ?? null,
          });

          return doc.effectiveFileUrl !== spItem.webUrl
            || doc.slScope !== desiredScope
            || (hasRelativePath && doc.storedRelativePath !== spItem.relativePath);
        })
        .map(({ doc, spItem }) => {
          const desiredScope = resolveScopeFromLocation({
            webUrl: spItem.webUrl,
            sourceSiteUrl: spItem.sourceSiteUrl,
            deptSiteUrl: siteUrl,
            deptBaseFolder: baseFolder,
            itemRelativePath: spItem.relativePath,
            globalCommonSiteUrl: globalCommon?.siteUrl ?? null,
            globalCommonFolder: globalCommon?.folder ?? null,
          });

          const slOwner = desiredScope === "personal"
            ? resolvePersonalOwnerHash(spItem, baseFolder)
            : null;

          // フォルダ名が既知ユーザーに対応しない場合は dept_common に格下げ
          const effectiveScope: ScopeKind = desiredScope === "personal" && slOwner === null
            ? "dept_common"
            : desiredScope;

          return {
            id: doc.id,
            effectiveFileUrl: spItem.webUrl,
            slScope: effectiveScope,
            slOwner,
            ...(hasRelativePath ? { relativePath: spItem.relativePath } : {}),
          };
        });

      const reindexCandidates = hasChangeTrackingFields || hasPageMetadataFields
        ? findReindexCandidates(
            matchedDocs.filter(({ doc }) => doc.slScope !== "global_common"),
            hasPageMetadataFields
          )
        : [];

      if (apply && !reindexOnly) {
        await deleteIndexDocs(orphanIds);
        await updateIndexDocs(docUpdates);
      }
      if (apply && indexNew) {
        reindexJobs.push(
          ...reindexCandidates.map((candidate) => ({
            resultKey: dept,
            candidate,
            dept,
            siteUrl,
            driveName,
            baseFolder,
            globalCommon,
          }))
        );
      }

      const deptResult: SlSyncDeptResult = {
        spFileNames: inventory.allItems.length,
        indexDocs: indexDocs.length,
        deleted: apply && !reindexOnly ? orphanIds.length : 0,
        urlUpdated: docUpdates.length,
        reindexCandidates: reindexCandidates.length,
        reindexCandidateNames: reindexCandidates
          .slice(0, 5)
          .map((candidate) => candidate.item.name),
        reindexed: 0,
        reindexFailed: 0,
        orphanIds,
      };

      if (indexNew && !reindexOnly) {
        const unindexed = findUnindexedSpItems(inventory, indexDocs);
        console.log(
          `[SL sync] dept=${dept} unindexedSPFiles=${unindexed.length} apply=${apply}`
        );
        if (apply && unindexed.length > 0) {
          const { indexed, skipped } = await indexNewSpFiles({
            accessToken,
            dept,
            siteUrl,
            driveName,
            baseFolder,
            unindexedItems: unindexed,
            batchSize,
            globalCommon,
            hasRelativePath,
            hasChangeTrackingFields,
            hasPageMetadataFields,
          });
          deptResult.newIndexed = indexed;
          deptResult.newSkipped = skipped;
        } else {
          deptResult.unindexedCount = unindexed.length;
        }
      }

      results[dept] = deptResult;
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
        const globalIndexDocs = await getIndexDocsGlobalCommon(
          globalCommon,
          hasRelativePath,
          hasChangeTrackingFields,
          hasPageMetadataFields
        );
        console.log(`[SL sync] global_common SP files=${globalScan.inventory.allItems.length} indexed docs=${globalIndexDocs.length}`);
        const matchedDocs: MatchedIndexDoc[] = globalIndexDocs.map((doc) => ({
          doc,
          spItem: findMatchingSpItem(doc, globalScan.inventory),
        }));

        // scan 範囲外（Common フォルダ外）へ移動したファイルを Graph API で直接追跡
        const globalScanMissed = matchedDocs.filter(
          (entry) => !entry.spItem && entry.doc.spItemId
        );
        if (globalScanMissed.length > 0) {
          const uniqueGlobalSpItemIds = Array.from(new Set(globalScanMissed.map((e) => e.doc.spItemId!)));
          console.log(`[SL sync] global_common looking up ${uniqueGlobalSpItemIds.length} unique spItemIds (${globalScanMissed.length} docs)`);
          const globalLookupCache = new Map<string, SpFileItem | null | "error">();
          for (const spItemId of uniqueGlobalSpItemIds) {
            globalLookupCache.set(spItemId, await lookupSpItemByIdInDrive(accessToken, globalCommon.siteUrl, globalCommon.driveName, spItemId));
          }
          for (const entry of globalScanMissed) {
            const found = globalLookupCache.get(entry.doc.spItemId!);
            if (found === "error") {
              entry.lookupFailed = true;
              console.warn(`[SL sync] global_common lookup error for ${entry.doc.fileName}, skipping orphan`);
            } else if (found) {
              entry.spItem = found;
              console.log(`[SL sync] global_common found outside scan scope: ${entry.doc.fileName} → ${found.webUrl}`);
            }
          }
        }

        const globalOrphanIds = matchedDocs
          .filter(({ spItem, lookupFailed }) => !spItem && !lookupFailed)
          .map(({ doc }) => doc.id);

        const docUpdates = matchedDocs
          .filter(
            (entry): entry is { doc: IndexDoc; spItem: SpFileItem } =>
              Boolean(entry.spItem)
          )
          .filter(({ doc, spItem }) =>
            doc.effectiveFileUrl !== spItem.webUrl
            || (hasRelativePath && doc.storedRelativePath !== spItem.relativePath)
          )
          .map(({ doc, spItem }) => ({
            id: doc.id,
            effectiveFileUrl: spItem.webUrl,
            ...(hasRelativePath ? { relativePath: spItem.relativePath } : {}),
          }));

        const reindexCandidates = hasChangeTrackingFields || hasPageMetadataFields
          ? findReindexCandidates(matchedDocs, hasPageMetadataFields)
          : [];
        if (globalOrphanIds.length > 0) {
          console.log(
            `[SL sync] global_common orphans=${globalOrphanIds.length} apply=${apply}`
          );
        }

        if (apply && !reindexOnly) {
          if (globalOrphanIds.length > 0) await deleteIndexDocs(globalOrphanIds);
          if (docUpdates.length > 0) await updateIndexDocs(docUpdates);
        }
        if (apply && indexNew) {
          reindexJobs.push(
            ...reindexCandidates.map((candidate) => ({
              resultKey: "global_common",
              candidate,
              dept: "common",
              siteUrl: globalCommon.siteUrl,
              driveName: globalCommon.driveName,
              baseFolder: globalCommon.folder,
              globalCommon,
            }))
          );
        }

        const gcResult: SlSyncDeptResult = {
          spFileNames: globalScan.inventory.allItems.length,
          indexDocs: globalIndexDocs.length,
          deleted: apply && !reindexOnly ? globalOrphanIds.length : 0,
          urlUpdated: docUpdates.length,
          reindexCandidates: reindexCandidates.length,
          reindexCandidateNames: reindexCandidates
            .slice(0, 5)
            .map((candidate) => candidate.item.name),
          reindexed: 0,
          reindexFailed: 0,
          orphanIds: globalOrphanIds,
        };

        if (indexNew && !reindexOnly) {
          const gcUnindexed = findUnindexedSpItems(globalScan.inventory, globalIndexDocs);
          console.log(
            `[SL sync] global_common unindexedSPFiles=${gcUnindexed.length} apply=${apply}`
          );
          if (apply && gcUnindexed.length > 0) {
            const { indexed, skipped } = await indexNewSpFiles({
              accessToken,
              dept: "common",
              siteUrl: globalCommon.siteUrl,
              driveName: globalCommon.driveName,
              baseFolder: globalCommon.folder,
              unindexedItems: gcUnindexed,
              batchSize,
              globalCommon,
              hasRelativePath,
              hasChangeTrackingFields,
              hasPageMetadataFields,
            });
            gcResult.newIndexed = indexed;
            gcResult.newSkipped = skipped;
          } else {
            gcResult.unindexedCount = gcUnindexed.length;
          }
        }

        results["global_common"] = gcResult;
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

  // Prioritize real content changes across every department, then spend any
  // remaining budget on the gradual migration of legacy untagged documents.
  const selectedReindexJobs = reindexJobs
    .sort((a, b) => {
      if (a.candidate.reason !== b.candidate.reason) {
        return a.candidate.reason === "content_changed" ? -1 : 1;
      }
      const aModified = Date.parse(a.candidate.item.lastModifiedAt ?? "");
      const bModified = Date.parse(b.candidate.item.lastModifiedAt ?? "");
      return (
        (Number.isNaN(bModified) ? 0 : bModified) -
        (Number.isNaN(aModified) ? 0 : aModified)
      );
    })
    .slice(0, batchSize);

  for (const job of selectedReindexJobs) {
    try {
      const reindexResult = await reindexSpFiles({
        accessToken,
        dept: job.dept,
        siteUrl: job.siteUrl,
        driveName: job.driveName,
        baseFolder: job.baseFolder,
        candidates: [job.candidate],
        batchSize: 1,
        globalCommon: job.globalCommon,
        hasRelativePath,
        hasChangeTrackingFields,
        hasPageMetadataFields,
      });
      const resultRow = results[job.resultKey];
      if (resultRow) {
        resultRow.reindexed =
          (resultRow.reindexed ?? 0) + reindexResult.reindexed;
        resultRow.reindexFailed =
          (resultRow.reindexFailed ?? 0) + reindexResult.failed;
      }
    } catch (error) {
      console.error(
        `[SL sync] Reindex job failed for ${job.candidate.item.name}`,
        error
      );
      const resultRow = results[job.resultKey];
      if (resultRow) {
        resultRow.reindexFailed = (resultRow.reindexFailed ?? 0) + 1;
      }
    }
  }

  return {
    ok: true,
    mode: apply ? "apply" : "dry-run",
    results,
  };
}
