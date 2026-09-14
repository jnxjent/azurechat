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
const { reserveSlSyncBudget, rolloverSlSyncLedger } = loaded.exports;

const now = new Date("2026-09-14T00:00:00Z");
const max = { daily: 5, monthly: 8, perFile: 4, attempts: 2 };
const ledger = {
  version: 1, day: "2026-09-14", month: "2026-09",
  dailyPages: 0, monthlyPages: 0, attempts: {},
};

assert.equal(reserveSlSyncBudget(ledger, "a", 4, max, now), null);
assert.equal(reserveSlSyncBudget(ledger, "b", 2, max, now), "daily_page_limit");
assert.equal(ledger.attempts.b, undefined, "deferred work must not consume an attempt");
assert.equal(reserveSlSyncBudget(ledger, "a", 1, max, now), null);
assert.equal(reserveSlSyncBudget(ledger, "a", 1, max, now), "attempt_limit");
assert.equal(ledger.monthlyPages, 5, "failed attempts still consume their reserved pages");

rolloverSlSyncLedger(ledger, "2026-09-15", "2026-09", new Date("2026-09-15T00:00:00Z"));
assert.equal(ledger.dailyPages, 0);
assert.equal(reserveSlSyncBudget(ledger, "a", 1, max, now), "attempt_limit", "day rollover must not reset attempts");
assert.equal(reserveSlSyncBudget(ledger, "b", 4, max, now), "monthly_page_limit");
assert.equal(reserveSlSyncBudget(ledger, "b", 3, max, now), null);
assert.equal(reserveSlSyncBudget(ledger, "c", 5, max, now), "file_page_limit");
assert.equal(reserveSlSyncBudget(ledger, "c", -1, max, now), "page_count_unknown");

ledger.attempts.b.succeeded = true;
rolloverSlSyncLedger(ledger, "2026-10-01", "2026-10", new Date("2026-10-01T00:00:00Z"));
assert.equal(ledger.monthlyPages, 0);
assert.equal(reserveSlSyncBudget(ledger, "b", 1, max, now), "already_indexed_this_version");

console.log("SL sync budget regression tests passed");
