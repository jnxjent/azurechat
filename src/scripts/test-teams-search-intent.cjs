const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");
const ts = require("typescript");

function loadTypeScriptModule(relativePath) {
  const fileName = path.resolve(relativePath);
  const source = fs.readFileSync(fileName, "utf8");
  const javascript = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
    },
  }).outputText;
  const loaded = new Module(fileName);
  loaded.filename = fileName;
  loaded.paths = Module._nodeModulePaths(path.dirname(fileName));
  loaded._compile(javascript, fileName);
  return loaded.exports;
}

const { requiresTeamsInternalSearch } = loadTypeScriptModule(
  "features/teams/teams-search-intent.ts"
);

assert.equal(requiresTeamsInternalSearch("社内の書類によればどうなっていますか"), true);
assert.equal(requiresTeamsInternalSearch("社内規程によれば申請期限はいつですか"), true);
assert.equal(requiresTeamsInternalSearch("SharePointの資料を調べて"), true);
assert.equal(requiresTeamsInternalSearch("SPにある規程を確認して"), true);
assert.equal(requiresTeamsInternalSearch("AI Searchで検索して"), true);

assert.equal(requiresTeamsInternalSearch("浜松の明日の天気は？"), false);
assert.equal(requiresTeamsInternalSearch("一般的な有給休暇の制度を教えて"), false);
assert.equal(requiresTeamsInternalSearch("Pythonのコードを書いて"), false);
assert.equal(requiresTeamsInternalSearch("こんにちは"), false);

console.log("Teams internal search intent tests passed.");

const { buildTeamsWebQuery } = loadTypeScriptModule(
  "features/teams/teams-web-query.ts"
);
const fixedNow = new Date("2026-07-21T06:00:00.000Z");
const weather = buildTeamsWebQuery("浜松の明日の天気は？", fixedNow);
assert.equal(weather.enriched, true);
assert.match(weather.query, /2026-07-22/);
assert.match(weather.query, /最高気温 最低気温 降水確率/);

const general = buildTeamsWebQuery("最新のAIニュース", fixedNow);
assert.deepEqual(general, { query: "最新のAIニュース", enriched: false });

console.log("Teams Web query tests passed.");
