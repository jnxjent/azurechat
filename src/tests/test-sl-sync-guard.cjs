// Exercises the production guard and its Blob ledger without Azure, Graph, or OCR calls.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");
const JSZip = require("jszip");

function loadTs(file, dependencies = {}) {
  const sourcePath = path.join(__dirname, "..", "lib", file);
  const compiled = ts.transpileModule(fs.readFileSync(sourcePath, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 },
    fileName: sourcePath,
  });
  const loaded = new Module(sourcePath, module);
  loaded.filename = sourcePath;
  loaded.paths = Module._nodeModulePaths(path.dirname(sourcePath));
  const originalRequire = loaded.require.bind(loaded);
  loaded.require = (name) => Object.hasOwn(dependencies, name)
    ? dependencies[name] : originalRequire(name);
  loaded._compile(compiled.outputText, sourcePath);
  return loaded.exports;
}

const blobs = new Map();
let revision = 0;
const missing = () => Object.assign(new Error("Blob not found"), { statusCode: 404 });
const conflict = () => Object.assign(new Error("Blob conflict"), { statusCode: 409 });

function blobClient(name) {
  return {
    async getProperties() {
      const blob = blobs.get(name);
      if (!blob) throw missing();
      return { etag: blob.etag };
    },
    async downloadToBuffer() {
      const blob = blobs.get(name);
      if (!blob) throw missing();
      return blob.data;
    },
    async uploadData(data, options = {}) {
      const previous = blobs.get(name);
      if (options.conditions?.ifNoneMatch === "*" && previous) throw conflict();
      if (options.conditions?.ifMatch && previous?.etag !== options.conditions.ifMatch) {
        throw conflict();
      }
      const etag = String(++revision);
      blobs.set(name, { data: Buffer.from(data), etag, leased: previous?.leased ?? false });
      return { etag };
    },
    getBlobLeaseClient() {
      return {
        async acquireLease() {
          const blob = blobs.get(name);
          if (blob.leased) throw conflict();
          blob.leased = true;
        },
        async releaseLease() { blobs.get(name).leased = false; },
      };
    },
  };
}

const fakeStorage = {
  BlobServiceClient: {
    fromConnectionString(connection) {
      assert.equal(connection, "mock-storage");
      return {
        getContainerClient(container) {
          assert.equal(container, "sl-sync-guard");
          return {
            async createIfNotExists() {},
            getBlockBlobClient: blobClient,
          };
        },
      };
    },
  },
};

const budget = loadTs("sl-sync-budget.ts");
const guard = loadTs("sl-sync-guard.ts", {
  "@azure/storage-blob": fakeStorage,
  "./document-extract": { extractWordText: async () => [] },
  "./sl-sync-budget": budget,
  "pdfjs-dist/legacy/build/pdf.js": {
    getDocument({ data }) {
      return { promise: Promise.resolve({ numPages: new DataView(data.buffer).getUint16(0, true) }) };
    },
  },
});

const file = (itemId) => ({
  sourceSite: "https://example.test/sharepoint",
  driveId: "drive",
  itemId,
  contentTag: "v1",
  fileName: "test.png",
  buffer: new ArrayBuffer(0),
});
const blocked = (reason) => (error) => {
  assert.equal(error?.reason, reason);
  return true;
};

async function main() {
  const monthParts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Tokyo", year: "numeric", month: "2-digit",
  }).formatToParts(new Date());
  const currentMonth = `${monthParts.find((part) => part.type === "year").value}-${monthParts.find((part) => part.type === "month").value}`;
  process.env.SL_SYNC_GUARD_STORAGE_CONNECTION_STRING = "mock-storage";
  process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = "https://mock-di-one.test/";
  process.env.AZURE_SEARCH_INDEX_NAME = "dl_index_phase15";
  process.env.SL_SYNC_DAILY_OCR_PAGE_LIMIT = "10";
  process.env.SL_SYNC_MONTHLY_OCR_PAGE_LIMIT = "100";
  process.env.SL_SYNC_TEMPORARY_LIMIT_MONTH = currentMonth;
  process.env.SL_SYNC_TEMPORARY_MONTHLY_OCR_PAGE_LIMIT = "4";
  process.env.SL_SYNC_INITIAL_MONTHLY_OCR_PAGES = "2";
  process.env.SL_SYNC_MAX_FILE_OCR_PAGES = "200";
  process.env.SL_SYNC_MAX_FILE_ATTEMPTS = "2";
  process.env.SL_SYNC_MAX_FILE_TOTAL_ATTEMPTS = "2";
  delete process.env.SL_SYNC_PAUSE_AT;

  const prodId = await guard.claimSlSyncFile(file("same"));
  process.env.AZURE_SEARCH_INDEX_NAME = "dl_index_phase15_test";
  const testId = await guard.claimSlSyncFile(file("same"));
  assert.notEqual(prodId, testId, "each Search index needs its own attempt record");
  await assert.rejects(guard.claimSlSyncFile(file("extra")), blocked("monthly_page_limit"));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([file("extra")])], [0],
    "a file that does not fit the remaining monthly budget must not block smaller files");
  process.env.SL_SYNC_TEMPORARY_MONTHLY_OCR_PAGE_LIMIT = "5";
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([file("extra")])], [],
    "raising the monthly limit must release a deferred file");
  process.env.SL_SYNC_TEMPORARY_MONTHLY_OCR_PAGE_LIMIT = "4";
  const ledger = [...blobs.entries()].find(([name]) => name.endsWith(".json"));
  assert.equal(JSON.parse(ledger[1].data).monthlyPages, 4, "both sites share one page budget");

  process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = "https://mock-di-two.test/";
  process.env.SL_SYNC_INITIAL_MONTHLY_OCR_PAGES = "0";
  process.env.SL_SYNC_TEMPORARY_MONTHLY_OCR_PAGE_LIMIT = "100";
  const retryId = await guard.claimSlSyncFile(file("retry"));
  assert.equal(retryId, await guard.claimSlSyncFile(file("retry")));
  await assert.rejects(guard.claimSlSyncFile(file("retry")), blocked("total_attempt_limit"));
  const exhausted = ["retry", "retry", "retry", "retry", "retry", "next"].map((itemId) => file(itemId));
  const excluded = await guard.getSlSyncExcludedCandidatePositions(exhausted);
  assert.deepEqual([...excluded], [0, 1, 2, 3, 4]);
  assert.equal(exhausted.filter((_, position) => !excluded.has(position)).slice(0, 5)[0].itemId,
    "next", "a retry-exhausted first batch must not block later files");
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([
    { ...file("retry"), contentTag: "v2" },
  ])], [], "a changed file version may be retried");
  const retryBlob = [...blobs.keys()].find((name) => name.endsWith(".json") &&
    JSON.parse(blobs.get(name).data).attempts[retryId]);
  const retryRecord = JSON.parse(blobs.get(retryBlob).data);
  retryRecord.attempts[retryId].day = "2000-01-01";
  await blobClient(retryBlob).uploadData(Buffer.from(JSON.stringify(retryRecord)));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([file("retry")])], [0],
    "two cumulative attempts remain exhausted on the next day");
  await assert.rejects(guard.claimSlSyncFile(file("retry")), blocked("total_attempt_limit"));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([file("retry")])], [0],
    "two cumulative attempts stop a small file from blocking later files");
  await guard.markSlSyncFileIndexed(retryId);
  await assert.rejects(guard.claimSlSyncFile(file("retry")), blocked("already_indexed_this_version"));

  process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = "https://mock-di-pre-ocr.test/";
  const badDownload = file("bad-download");
  assert.deepEqual(await guard.recordSlSyncFilePreOcrFailure(badDownload), { count: 1, limit: 2 });
  assert.deepEqual(await guard.recordSlSyncFilePreOcrFailure(badDownload), { count: 2, limit: 2 });
  assert.deepEqual(await guard.recordSlSyncFilePreOcrFailure(badDownload), { count: 2, limit: 2 });
  const preOcrBatch = ["bad-download", "bad-download", "next-file"].map((itemId) => file(itemId));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions(preOcrBatch)], [0, 1],
    "a failed download or page count must release the first batch slot");
  const preOcrExcluded = await guard.getSlSyncExcludedCandidatePositions(preOcrBatch);
  assert.equal(preOcrBatch.filter((_, position) => !preOcrExcluded.has(position))[0].itemId,
    "next-file", "the next candidate must enter the batch");
  const preOcrBlob = [...blobs.entries()].find(([name, blob]) =>
    name.endsWith(".json") && Object.values(JSON.parse(blob.data).attempts).some((entry) =>
      entry.totalCount === 2 && !Object.hasOwn(entry, "pages")));
  assert.ok(preOcrBlob, "pre-OCR failures must be written to the shared ledger");
  assert.equal(JSON.parse(preOcrBlob[1].data).monthlyPages, 0,
    "a failed download must not charge OCR pages");
  await assert.rejects(guard.claimSlSyncFile(badDownload), blocked("total_attempt_limit"));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([
    { ...badDownload, contentTag: "v2" },
  ])], [], "a corrected file version can be retried");

  process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = "https://mock-di-large.test/";
  process.env.SL_SYNC_DAILY_OCR_PAGE_LIMIT = "1000";
  process.env.SL_SYNC_TEMPORARY_MONTHLY_OCR_PAGE_LIMIT = "10000";
  process.env.SL_SYNC_MAX_FILE_OCR_PAGES = "500";
  const pdf = (itemId, pages) => ({
    ...file(itemId), fileName: "large.pdf", buffer: new Uint16Array([pages]).buffer,
  });
  await guard.claimSlSyncFile(pdf("large", 400));
  await guard.claimSlSyncFile(pdf("large", 400));
  await assert.rejects(guard.claimSlSyncFile(pdf("large", 400)), blocked("total_attempt_limit"));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([pdf("large", 400)])], [0],
    "a failed large PDF must stop occupying the next batch");
  await assert.rejects(guard.claimSlSyncFile(pdf("oversized", 600)), blocked("file_page_limit"));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([pdf("oversized", 600)])], [0]);
  process.env.SL_SYNC_MAX_FILE_OCR_PAGES = "600";
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([pdf("oversized", 600)])], [],
    "raising the page limit must release an oversized PDF");

  process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = "https://mock-di-excel.test/";
  process.env.SL_SYNC_MAX_EXCEL_SOURCE_BYTES = "10000";
  process.env.SL_SYNC_MAX_EXCEL_XML_BYTES = "100";
  process.env.SL_SYNC_MAX_EXCEL_CHUNKS = "2";
  const bigXml = new JSZip();
  bigXml.file("xl/worksheets/sheet1.xml", "x".repeat(101));
  const largeXlsx = { ...file("large-xlsx"), fileName: "large.xlsx",
    buffer: await bigXml.generateAsync({ type: "nodebuffer" }) };
  await assert.rejects(guard.claimSlSyncFile(largeXlsx), blocked("excel_xml_bytes_limit"));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([largeXlsx, file("next")])], [0],
    "an oversized XLSX must release its batch slot without embedding");
  await assert.rejects(guard.claimSlSyncFile({ ...largeXlsx, itemId: "large-xlsm", fileName: "large.xlsm" }),
    blocked("excel_xml_bytes_limit"));
  process.env.SL_SYNC_MAX_EXCEL_XML_BYTES = "101";
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([largeXlsx])], [],
    "raising the XML limit must allow the same file version");

  process.env.SL_SYNC_MAX_EXCEL_SOURCE_BYTES = "100";
  const largeXls = { ...file("large-xls"), fileName: "large.xls", buffer: Buffer.alloc(101) };
  await assert.rejects(guard.claimSlSyncFile(largeXls), blocked("excel_source_bytes_limit"));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([largeXls])], [0],
    "an oversized XLS must be excluded before parsing");
  process.env.SL_SYNC_MAX_EXCEL_SOURCE_BYTES = "102";
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([largeXls])], []);

  const manyChunks = { ...file("many-chunks"), fileName: "many.xls", buffer: Buffer.alloc(10) };
  await guard.claimSlSyncFile(manyChunks);
  await assert.rejects(guard.enforceSlSyncExcelChunkLimit(manyChunks, 3), blocked("excel_chunks_limit"));
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([manyChunks])], [0],
    "an Excel file with too many chunks must be excluded before embedding");
  process.env.SL_SYNC_MAX_EXCEL_CHUNKS = "3";
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([manyChunks])], [],
    "raising the chunk limit must release the file");
  assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([
    { ...manyChunks, contentTag: "v2" },
  ])], [], "a changed Excel file version may be inspected again");

  const actualExcel = process.env.SL_SYNC_EXCEL_TEST_FILE;
  if (actualExcel) {
    process.env.SL_SYNC_MAX_EXCEL_SOURCE_BYTES = String(5 * 1024 * 1024);
    process.env.SL_SYNC_MAX_EXCEL_XML_BYTES = String(20 * 1024 * 1024);
    const realFile = { ...file("real-excel"), fileName: "test2.xlsx", buffer: fs.readFileSync(actualExcel) };
    await assert.rejects(guard.claimSlSyncFile(realFile), blocked("excel_xml_bytes_limit"));
    assert.deepEqual([...await guard.getSlSyncExcludedCandidatePositions([realFile])], [0]);
  }

  process.env.SL_SYNC_DAILY_OCR_PAGE_LIMIT = "bad";
  await assert.rejects(guard.claimSlSyncFile(file("invalid")), blocked("invalid_sl_sync_daily_ocr_page_limit"));
  process.env.SL_SYNC_DAILY_OCR_PAGE_LIMIT = "10";
  process.env.SL_SYNC_MAX_FILE_OCR_PAGES = "200";
  process.env.SL_SYNC_PAUSE_AT = "2026-09-21T00:00:00+09:00";
  assert.equal(guard.slSyncPauseReason(new Date("2026-09-20T14:59:59Z")), null);
  assert.equal(guard.slSyncPauseReason(new Date("2026-09-20T15:00:00Z")), "scheduled_pause");
  process.env.SL_SYNC_PAUSE_AT = "2026-09-13T00:00:00+09:00";
  await assert.rejects(guard.claimSlSyncFile(file("paused")), blocked("scheduled_pause"));
  process.env.SL_SYNC_PAUSE_AT = "2026-09-21";
  assert.equal(guard.slSyncPauseReason(), "invalid_pause_at");
  delete process.env.SL_SYNC_PAUSE_AT;

  process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = "https://mock-di-three.test/";
  process.env.SL_SYNC_TEMPORARY_MONTHLY_OCR_PAGE_LIMIT = "1";
  const results = await Promise.allSettled([
    guard.claimSlSyncFile(file("concurrent-a")),
    guard.claimSlSyncFile(file("concurrent-b")),
  ]);
  assert.equal(results.filter((result) => result.status === "fulfilled").length, 1);
  assert.equal(results.filter((result) => result.status === "rejected").length, 1);
  assert.equal(results.find((result) => result.status === "rejected").reason.reason, "monthly_page_limit");

  // A temporary limit for another month must not replace the normal limit.
  const [year, month] = currentMonth.split("-").map(Number);
  process.env.SL_SYNC_TEMPORARY_LIMIT_MONTH = new Date(Date.UTC(year, month, 1)).toISOString().slice(0, 7);
  process.env.SL_SYNC_MONTHLY_OCR_PAGE_LIMIT = "2";
  process.env.SL_SYNC_TEMPORARY_MONTHLY_OCR_PAGE_LIMIT = "100";
  process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = "https://mock-di-four.test/";
  await guard.claimSlSyncFile(file("normal-a"));
  await guard.claimSlSyncFile(file("normal-b"));
  await assert.rejects(guard.claimSlSyncFile(file("normal-c")), blocked("monthly_page_limit"));

  delete process.env.SL_SYNC_TEMPORARY_MONTHLY_OCR_PAGE_LIMIT;
  await assert.rejects(guard.claimSlSyncFile(file("invalid-override")), blocked("invalid_temporary_monthly_limit"));

  console.log("SL sync guard local tests passed (no Azure or OCR calls)");
}

main().catch((error) => { console.error(error); process.exitCode = 1; });
