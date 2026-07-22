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
      esModuleInterop: true,
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

const policy = loadTypeScriptModule("features/teams/teams-file-policy.ts");

const candidates = policy.parseTeamsFileCandidates([
  {
    contentType: "application/vnd.microsoft.teams.file.download.info",
    name: "決算資料.pdf",
    content: {
      downloadUrl: "https://tenant.sharepoint.com/download?id=1",
      fileType: "pdf",
      uniqueId: "item-1",
    },
  },
  { contentType: "application/vnd.microsoft.card.adaptive", content: {} },
]);
assert.equal(candidates.length, 1);
assert.equal(candidates[0].fileName, "決算資料.pdf");
assert.equal(candidates[0].extension, "pdf");

assert.throws(
  () =>
    policy.parseTeamsFileCandidates([
      {
        contentType: "application/vnd.microsoft.teams.file.download.info",
        name: "malware.exe",
        content: { downloadUrl: "https://tenant.sharepoint.com/download" },
      },
    ]),
  /未対応/
);
assert.throws(
  () => policy.assertSafeTeamsDownloadUrl("http://tenant.sharepoint.com/a"),
  /許可/
);
assert.throws(
  () => policy.assertSafeTeamsDownloadUrl("https://127.0.0.1/a"),
  /許可/
);

policy.validateTeamsFileBytes(
  "sample.pdf",
  "pdf",
  Buffer.from("%PDF-1.7\ncontent")
);
policy.validateTeamsFileBytes(
  "sample.docx",
  "docx",
  Buffer.from([0x50, 0x4b, 0x03, 0x04])
);
assert.throws(
  () =>
    policy.validateTeamsFileBytes(
      "fake.pdf",
      "pdf",
      Buffer.from("not a pdf")
    ),
  /一致しません/
);

assert.equal(
  policy.stripTeamsAttachmentMarkup(
    '<attachment id="1">file</attachment> Excelに変換して'
  ),
  "Excelに変換して"
);
assert.equal(policy.referencesTeamsUpload("先ほどのファイルを編集して"), true);
assert.equal(policy.referencesTeamsUpload("東京の天気は？"), false);

console.log("Teams file policy tests passed.");
