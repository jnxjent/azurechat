const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const sourcePath = path.join(__dirname, "..", "lib", "sl-sync-budget.ts");
const compiled = ts.transpileModule(fs.readFileSync(sourcePath, "utf8"), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 },
  fileName: sourcePath,
});
const loaded = new Module(sourcePath, module);
loaded.filename = sourcePath;
loaded.paths = Module._nodeModulePaths(path.dirname(sourcePath));
loaded._compile(compiled.outputText, sourcePath);
const { recordSlSyncPreOcrFailure, reserveSlSyncBudget, rolloverSlSyncLedger } = loaded.exports;

const now = new Date("2026-09-14T00:00:00Z");
const max = { daily: 5, monthly: 8, perFile: 500, attempts: 2, totalAttempts: 2 };
const ledger = {
  version: 1, day: "2026-09-14", month: "2026-09",
  dailyPages: 0, monthlyPages: 0, attempts: {},
};

assert.equal(reserveSlSyncBudget(ledger, "a", 4, max, now), null);
assert.equal(reserveSlSyncBudget(ledger, "b", 2, max, now), "daily_page_limit");
assert.equal(ledger.attempts.b.count, 0, "deferred work must not consume an attempt");
assert.equal(ledger.attempts.b.deferredBudget.period, ledger.day);
assert.equal(reserveSlSyncBudget(ledger, "a", 1, max, now), null);
assert.equal(reserveSlSyncBudget(ledger, "a", 1, max, now), "total_attempt_limit");
assert.equal(ledger.monthlyPages, 5, "failed attempts still consume their reserved pages");

rolloverSlSyncLedger(ledger, "2026-09-15", "2026-09", new Date("2026-09-15T00:00:00Z"));
assert.equal(ledger.dailyPages, 0);
assert.equal(reserveSlSyncBudget(ledger, "a", 1, max, now), "total_attempt_limit",
  "a small file must stay stopped after two cumulative attempts across days");
assert.equal(reserveSlSyncBudget(ledger, "b", 4, max, now), "monthly_page_limit");
assert.equal(reserveSlSyncBudget(ledger, "b", 2, max, now), null);
assert.equal(reserveSlSyncBudget(ledger, "c", 501, max, now), "file_page_limit");
assert.equal(ledger.attempts.c.blockedPageLimit, 500, "oversized files must be skipped on later runs");
assert.equal(reserveSlSyncBudget(ledger, "c", -1, max, now), "page_count_unknown");

ledger.attempts.b.succeeded = true;
rolloverSlSyncLedger(ledger, "2026-10-01", "2026-10", new Date("2026-10-01T00:00:00Z"));
assert.equal(ledger.monthlyPages, 0);
assert.equal(reserveSlSyncBudget(ledger, "b", 1, max, now), "already_indexed_this_version");

const largeLedger = {
  version: 1, day: "2026-09-14", month: "2026-09",
  dailyPages: 0, monthlyPages: 0, attempts: {},
};
const largeMax = { daily: 1000, monthly: 10000, perFile: 500, attempts: 2, totalAttempts: 2 };
assert.equal(reserveSlSyncBudget(largeLedger, "large", 400, largeMax, now), null);
assert.equal(reserveSlSyncBudget(largeLedger, "large", 400, largeMax, now), null);
assert.equal(reserveSlSyncBudget(largeLedger, "large", 400, largeMax, now), "total_attempt_limit");
rolloverSlSyncLedger(largeLedger, "2026-09-15", "2026-09", new Date("2026-09-15T00:00:00Z"));
assert.equal(reserveSlSyncBudget(largeLedger, "large", 400, largeMax, now), "total_attempt_limit",
  "a failed 400-page file does not restart automatically the next day");
assert.equal(reserveSlSyncBudget(largeLedger, "too-large", 600, largeMax, now), "file_page_limit");
assert.equal(reserveSlSyncBudget(largeLedger, "too-large", 600,
  { ...largeMax, perFile: 600 }, now), null,
  "raising the file-page limit makes a previously oversized file eligible");

const preOcrLedger = {
  version: 1, day: "2026-09-14", month: "2026-09",
  dailyPages: 0, monthlyPages: 0, attempts: {},
};
assert.equal(recordSlSyncPreOcrFailure(preOcrLedger, "bad-pdf", max, now), 1);
assert.equal(recordSlSyncPreOcrFailure(preOcrLedger, "bad-pdf", max, now), 2);
assert.equal(recordSlSyncPreOcrFailure(preOcrLedger, "bad-pdf", max, now), 2,
  "a third pre-OCR failure must not increase the count");
assert.equal(preOcrLedger.attempts["bad-pdf"].count, 0,
  "pre-OCR failures do not consume an OCR attempt");
assert.equal(preOcrLedger.dailyPages, 0);
assert.equal(preOcrLedger.monthlyPages, 0);
assert.equal(reserveSlSyncBudget(preOcrLedger, "bad-pdf", 1, max, now), "total_attempt_limit",
  "a valid later download must still respect pre-OCR failures");
assert.equal(reserveSlSyncBudget(preOcrLedger, "mixed", 1, max, now), null);
assert.equal(recordSlSyncPreOcrFailure(preOcrLedger, "mixed", max, now), 2,
  "one OCR attempt plus one pre-OCR failure reaches the same cumulative limit");

console.log("SL sync budget regression tests passed");
