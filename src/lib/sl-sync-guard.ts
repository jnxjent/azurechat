import { createHash } from "node:crypto";
import { BlobServiceClient, type ContainerClient } from "@azure/storage-blob";
import { extractWordText } from "./document-extract";
import { recordSlSyncPreOcrFailure as recordPreOcrFailure, reserveSlSyncBudget, rolloverSlSyncLedger, SL_SYNC_CONTINUING_RETRY_MAX_PAGES, type SlSyncLedger } from "./sl-sync-budget";

export class SlSyncGuardBlockedError extends Error {
  constructor(public readonly reason: string) {
    super(`SL sync OCR blocked: ${reason}`);
    this.name = "SlSyncGuardBlockedError";
  }
}

export type SlSyncFileIdentity = {
  sourceSite: string;
  driveId: string;
  itemId: string;
  contentTag: string | null;
};

const CONTAINER = "sl-sync-guard";
const hash = (value: string) => createHash("sha256").update(value).digest("hex");
const positiveInt = (name: string, fallback: number) => {
  const raw = process.env[name];
  if (raw === undefined) return fallback;
  const value = Number(raw.trim());
  if (!/^[1-9]\d*$/.test(raw.trim()) || !Number.isSafeInteger(value)) {
    throw new SlSyncGuardBlockedError(`invalid_${name.toLowerCase()}`);
  }
  return value;
};

// August 2026: 112,629 processed pages cost JPY 30,190. Reserve headroom
// below that pre-incident monthly usage; this is a page cap, not a yen cap.
const limits = () => ({
  daily: positiveInt("SL_SYNC_DAILY_OCR_PAGE_LIMIT", 5000),
  monthly: monthlyLimit(),
  perFile: positiveInt("SL_SYNC_MAX_FILE_OCR_PAGES", 500),
  attempts: positiveInt("SL_SYNC_MAX_FILE_ATTEMPTS", 2),
  totalAttempts: positiveInt("SL_SYNC_MAX_FILE_TOTAL_ATTEMPTS", 2),
});

function monthlyLimit(): number {
  const month = process.env.SL_SYNC_TEMPORARY_LIMIT_MONTH?.trim();
  const raw = process.env.SL_SYNC_TEMPORARY_MONTHLY_OCR_PAGE_LIMIT?.trim();
  if (month || raw) {
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month ?? "") ||
        !raw || !/^[1-9]\d*$/.test(raw) || !Number.isSafeInteger(Number(raw))) {
      throw new SlSyncGuardBlockedError("invalid_temporary_monthly_limit");
    }
    if (period().month === month) return Number(raw);
  }
  return positiveInt("SL_SYNC_MONTHLY_OCR_PAGE_LIMIT", 100000);
}

export function slSyncPauseReason(now = new Date()): string | null {
  const value = process.env.SL_SYNC_PAUSE_AT?.trim();
  if (!value) return null;
  // Require an explicit time zone so App Service UTC and Japan time agree.
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:Z|[+-]\d{2}:\d{2})$/.test(value)) {
    return "invalid_pause_at";
  }
  const end = Date.parse(value);
  if (!Number.isFinite(end)) return "invalid_pause_at";
  return now.getTime() >= end ? "scheduled_pause" : null;
}

function period(): { day: string; month: string } {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Tokyo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const get = (type: string) => parts.find((part) => part.type === type)?.value ?? "";
  const month = `${get("year")}-${get("month")}`;
  return { month, day: `${month}-${get("day")}` };
}

function storage(): { container: ContainerClient; ledgerName: string } {
  const account = process.env.AZURE_STORAGE_ACCOUNT_NAME?.trim();
  const key = process.env.AZURE_STORAGE_ACCOUNT_KEY?.trim();
  const guardConnection = process.env.SL_SYNC_GUARD_STORAGE_CONNECTION_STRING?.trim();
  const endpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT?.trim();
  if ((!guardConnection && (!account || !key)) || !endpoint) {
    throw new Error("SL sync guard requires Azure Storage and Document Intelligence settings");
  }
  const containerName = process.env.SL_SYNC_GUARD_CONTAINER_NAME?.trim() || CONTAINER;
  if (!/^[a-z0-9](?:[a-z0-9-]{1,61}[a-z0-9])$/.test(containerName) ||
      containerName.includes("--")) {
    throw new Error("Invalid SL sync guard container name");
  }
  const service = BlobServiceClient.fromConnectionString(
    guardConnection ||
      `DefaultEndpointsProtocol=https;AccountName=${account};AccountKey=${key};EndpointSuffix=core.windows.net`
  );
  return {
    container: service.getContainerClient(containerName),
    ledgerName: `budget-${hash(endpoint.toLowerCase()).slice(0, 24)}.json`,
  };
}

function fileAttemptId(file: SlSyncFileIdentity): string {
  return hash([
    process.env.AZURE_SEARCH_INDEX_NAME?.trim() ?? "",
    file.sourceSite,
    file.driveId,
    file.itemId,
    file.contentTag ?? "",
  ].join("|"));
}

// Read the shared attempt ledger before selecting a batch. Exhausted files
// must not keep occupying its first slots on every timer run.
export async function getSlSyncExcludedCandidatePositions(
  files: SlSyncFileIdentity[]
): Promise<Set<number>> {
  if (files.length === 0) return new Set();
  try {
    const max = limits();
    return await withLedger((ledger) => {
      const excluded = new Set<number>();
      files.forEach((file, position) => {
        const attempt = ledger.attempts[fileAttemptId(file)];
        if (!attempt) return;
        const oversized = attempt.blockedPageLimit !== undefined &&
          attempt.blockedPageLimit >= max.perFile;
        const budget = attempt.deferredBudget;
        const deferred = budget !== undefined &&
          (budget.reason === "daily_page_limit"
            ? budget.period === ledger.day && max.daily <= budget.limit
            : budget.period === ledger.month && max.monthly <= budget.limit);
        const exhausted = attempt.pages !== undefined &&
          attempt.count >= max.attempts &&
          (attempt.pages > SL_SYNC_CONTINUING_RETRY_MAX_PAGES || attempt.day === ledger.day);
        const cumulativeExhausted =
          (attempt.totalCount ?? attempt.count) >= max.totalAttempts;
        if (attempt.succeeded || oversized || deferred || exhausted || cumulativeExhausted) {
          excluded.add(position);
        }
      });
      return excluded;
    });
  } catch (error) {
    if (error instanceof SlSyncGuardBlockedError) throw error;
    console.error("[SL sync guard] Could not read attempt ledger:", error);
    throw new SlSyncGuardBlockedError("budget_store_unavailable");
  }
}

async function withLedger<T>(change: (ledger: SlSyncLedger) => Promise<T> | T): Promise<T> {
  const { container, ledgerName } = storage();
  await container.createIfNotExists();
  const lock = container.getBlockBlobClient(`${ledgerName}.lock`);
  try {
    await lock.uploadData(Buffer.alloc(0), { conditions: { ifNoneMatch: "*" } });
  } catch (error: any) {
    if (error?.statusCode !== 409 && error?.statusCode !== 412) throw error;
  }
  const lease = lock.getBlobLeaseClient();
  let acquired = false;
  for (let retry = 0; retry < 30; retry++) {
    try {
      await lease.acquireLease(60);
      acquired = true;
      break;
    } catch (error: any) {
      if (error?.statusCode !== 409 && error?.statusCode !== 412) throw error;
      await new Promise((resolve) => setTimeout(resolve, 500));
    }
  }
  if (!acquired) throw new Error("SL sync guard lock unavailable");
  try {
    const blob = container.getBlockBlobClient(ledgerName);
    let ledger: SlSyncLedger;
    let etag: string | undefined;
    try {
      etag = (await blob.getProperties()).etag;
      ledger = JSON.parse((await blob.downloadToBuffer()).toString()) as SlSyncLedger;
      if (ledger.version !== 1 || !ledger.attempts) throw new Error("Invalid SL sync guard ledger");
    } catch (error: any) {
      if (error?.statusCode !== 404) throw error;
      // Existing consumption is unknown on first deployment. Start with the
      // current month exhausted, then reset automatically next month. An
      // operator can seed a known current-month value before first use.
      const rawSeed = process.env.SL_SYNC_INITIAL_MONTHLY_OCR_PAGES?.trim();
      const seeded = rawSeed ? Number(rawSeed) : NaN;
      const initialPages = Number.isSafeInteger(seeded) && seeded >= 0
        ? seeded : limits().monthly;
      ledger = {
        version: 1,
        day: period().day,
        month: period().month,
        dailyPages: 0,
        monthlyPages: initialPages,
        attempts: {},
      };
      // Persist the initial exhausted-month state even when the first claim
      // is blocked. Otherwise each call would initialize it again forever.
      const created = await blob.uploadData(Buffer.from(JSON.stringify(ledger)), {
        blobHTTPHeaders: { blobContentType: "application/json" },
        conditions: { ifNoneMatch: "*" },
      });
      etag = created.etag;
    }
    const now = period();
    rolloverSlSyncLedger(ledger, now.day, now.month, new Date());
    const result = await change(ledger);
    if (!etag) throw new Error("SL sync guard ledger ETag missing");
    await blob.uploadData(Buffer.from(JSON.stringify(ledger)), {
      blobHTTPHeaders: { blobContentType: "application/json" },
      conditions: { ifMatch: etag },
    });
    return result;
  } finally {
    await lease.releaseLease();
  }
}

async function estimatePages(buffer: ArrayBuffer, fileName: string): Promise<number> {
  const name = fileName.toLowerCase();
  if (/\.(xlsx|xlsm|xls|txt|doc|msg)$/.test(name)) return 0; // Local extraction.
  if (/\.(png|jpe?g|bmp|gif|webp)$/.test(name)) return 1;
  if (/\.pdf$/.test(name)) {
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.js");
    // pdfjs may transfer its input to a worker; keep the original bytes for DI.
    const pdf = await pdfjs.getDocument({ data: new Uint8Array(buffer.slice(0)) }).promise;
    if (!Number.isSafeInteger(pdf.numPages) || pdf.numPages < 1) {
      throw new SlSyncGuardBlockedError("pdf_page_count_unknown");
    }
    return pdf.numPages;
  }
  if (/\.pptx$/.test(name)) {
    const jszip = await import("jszip");
    const zip = await (jszip.default ?? jszip).loadAsync(Buffer.from(buffer));
    const count = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path)).length;
    if (!count) throw new SlSyncGuardBlockedError("pptx_page_count_unknown");
    return count;
  }
  if (/\.docx$/.test(name)) {
    const text = (await extractWordText(buffer)).join("\n");
    if (!text) throw new SlSyncGuardBlockedError("docx_page_count_unknown");
    // Document Intelligence bills Word in 3,000-character page blocks.
    return Math.max(1, Math.ceil((text.length * 1.25) / 3000) + 1);
  }
  if (/\.tiff?$/.test(name)) {
    const sharp = (await import("sharp")).default;
    const count = (await sharp(Buffer.from(buffer)).metadata()).pages;
    if (!count) throw new SlSyncGuardBlockedError("tiff_page_count_unknown");
    return count;
  }
  throw new SlSyncGuardBlockedError("unsupported_page_count");
}

export async function claimSlSyncFile(params: {
  sourceSite: string;
  driveId: string;
  itemId: string;
  contentTag: string | null;
  fileName: string;
  buffer: ArrayBuffer;
}): Promise<string> {
  const pause = slSyncPauseReason();
  if (pause) throw new SlSyncGuardBlockedError(pause);
  let pages: number;
  try {
    pages = await estimatePages(params.buffer, params.fileName);
  } catch (error) {
    if (error instanceof SlSyncGuardBlockedError) throw error;
    throw new SlSyncGuardBlockedError("page_count_unknown");
  }
  const max = limits();
  // A successful production index must not suppress the TestSite index: they
  // share the page budget, but each needs its own file-version attempt record.
  const id = fileAttemptId(params);
  try {
    const reason = await withLedger((ledger) => {
      const pause = slSyncPauseReason();
      if (pause) throw new SlSyncGuardBlockedError(pause);
      const reason = reserveSlSyncBudget(ledger, id, pages, max, new Date());
      if (!reason) {
        console.log(`[SL sync guard] reserved ${pages} pages; daily=${ledger.dailyPages}/${max.daily} monthly=${ledger.monthlyPages}/${max.monthly}`);
      }
      return reason;
    });
    if (reason) throw new SlSyncGuardBlockedError(reason);
    return id;
  } catch (error) {
    if (error instanceof SlSyncGuardBlockedError) throw error;
    console.error("[SL sync guard] Persistent budget unavailable:", error);
    throw new SlSyncGuardBlockedError("budget_store_unavailable");
  }
}

export async function recordSlSyncFilePreOcrFailure(file: SlSyncFileIdentity): Promise<{ count: number; limit: number }> {
  const id = fileAttemptId(file);
  try {
    const max = limits();
    const count = await withLedger((ledger) =>
      recordPreOcrFailure(ledger, id, max, new Date())
    );
    return { count, limit: max.totalAttempts };
  } catch (error) {
    if (error instanceof SlSyncGuardBlockedError) throw error;
    console.error("[SL sync guard] Could not record pre-OCR failure:", error);
    throw new SlSyncGuardBlockedError("budget_store_unavailable");
  }
}

export async function markSlSyncFileIndexed(id: string): Promise<void> {
  await withLedger((ledger) => {
    const attempt = ledger.attempts[id];
    if (!attempt) throw new Error("SL sync guard claim missing");
    attempt.succeeded = true;
    attempt.updatedAt = new Date().toISOString();
  });
}
