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

const { resolveTeamsInstantReply } = loadTypeScriptModule(
  "features/teams/teams-instant-reply.ts"
);

assert.equal(resolveTeamsInstantReply("こんにちは"), "こんにちは！ご用件をどうぞ。");
assert.equal(resolveTeamsInstantReply(" こんにちわ！ "), "こんにちは！ご用件をどうぞ。");
assert.equal(resolveTeamsInstantReply("HELLO!!"), "Hello! How can I help?");
assert.equal(resolveTeamsInstantReply("こんばんは。"), "こんばんは！ご用件をどうぞ。");

// A greeting plus a real question must use the normal RAG/OpenAI route.
assert.equal(resolveTeamsInstantReply("こんにちは。搬入規程を教えて"), null);
assert.equal(resolveTeamsInstantReply("おはよう、昨日の資料を要約して"), null);
assert.equal(resolveTeamsInstantReply("搬入受付時の連絡先を教えて"), null);

console.log("Teams instant reply tests passed.");
