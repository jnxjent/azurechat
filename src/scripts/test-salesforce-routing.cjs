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

const routing = loadTypeScriptModule(
  "features/common/services/salesforce-routing.ts"
);
const access = loadTypeScriptModule(
  "features/common/services/salesforce-access.ts"
);

const cases = [
  ["こんにちは", true, "normal"],
  ["大映産業の住所は？", true, "normal"],
  ["Salesforceによれば、大映産業の住所は？", true, "salesforce"],
  ["セールスフォースで大映産業を調べて", true, "salesforce"],
  ["SFによれば、大映産業の住所を教えて", true, "salesforce"],
  ["ＳＦ連携で今月の商談を調べて", true, "salesforce"],
  ["Salesforceから取引先情報を取得して", true, "salesforce"],
  ["セールスフォース上の情報を確認して", true, "salesforce"],
  ["Salesforceで日栄産業の住所は？", true, "salesforce"],
  ["おすすめのSF映画は？", true, "normal"],
  ["SF小説を書いて", true, "normal"],
  ["おすすめのＳＦ映画は？", true, "normal"],
  ["Salesforceとは？", true, "knowledge"],
  ["Salesforceの使い方を教えて", true, "knowledge"],
  ["Salesforceでエラーが出たときのQAは？", true, "knowledge"],
  ["Salesforce連携のQAをSharePointから探して", true, "knowledge"],
  ["Salesforceの申請手順を社内資料から調べて", true, "knowledge"],
  ["Salesforceによれば、大映産業の住所は？", false, "denied"],
  ["Salesforce連携のQAをSharePointから探して", false, "knowledge"],
];

for (const [message, allowed, expectedRoute] of cases) {
  const result = routing.resolveSalesforceRoute({
    message,
    isSalesforceAllowed: allowed,
    hasSalesforceExtension: true,
  });
  assert.equal(result.route, expectedRoute, `${message}: ${result.route}`);
}

assert.equal(
  routing.resolveSalesforceRoute({
    message: "Salesforceで日栄産業を調べて",
    isSalesforceAllowed: true,
    hasSalesforceExtension: false,
  }).route,
  "normal"
);
assert.equal(
  access.isSalesforceAllowedEmail(
    " J.NOMOTO@MIDAC.JP ",
    "other@example.com, j.nomoto@midac.jp\n"
  ),
  true
);
assert.equal(access.isSalesforceAllowedEmail("other@example.com", ""), false);
assert.equal(
  routing.buildSalesforceGatewayQuery("Salesforceで日栄産業を調べて"),
  "取引先の日栄産業について教えて"
);
assert.equal(
  routing.buildSalesforceGatewayQuery("Salesforceで日栄産業の商談を調べて"),
  "日栄産業の商談を調べて"
);
assert.equal(
  routing.buildSalesforceGatewayQuery("Salesforceによれば、大映産業の住所は？"),
  "大映産業の住所は?"
);
assert.equal(
  routing.buildSalesforceGatewayQuery(
    "Salesforceで日栄産業株式会社の住所を教えて"
  ),
  "日栄産業株式会社の住所を教えて"
);

const webEntrySource = fs.readFileSync(
  "features/chat-page/chat-services/chat-api/chat-api.ts",
  "utf8"
);
const webExtensionSource = fs.readFileSync(
  "features/chat-page/chat-services/chat-api/chat-api-extension.ts",
  "utf8"
);
const teamsSource = fs.readFileSync("features/teams/teams-app.ts", "utf8");
assert.match(webEntrySource, /extensionId !== SF_EXTENSION_ID/);
assert.match(
  webExtensionSource,
  /salesforceRouting\.route === "salesforce"/
);
assert.match(webExtensionSource, /salesforceRouting\.route === "denied"/);
assert.match(teamsSource, /resolveSalesforceRoute\(/);
assert.match(teamsSource, /salesforceRouting\.route === "salesforce"/);
assert.match(
  teamsSource,
  /const localDevEmail = isLocalAuthSkipped\(\)[\s\S]*?process\.env\.NEXT_PUBLIC_DEV_USER_EMAIL[\s\S]*?if \(localDevEmail\)[\s\S]*?return localDevEmail;/
);
assert.doesNotMatch(
  teamsSource,
  /firstEmail\(\s*props\.activity\.from\?\.properties/
);

console.log(`Salesforce routing tests passed (${cases.length + 14} assertions).`);
