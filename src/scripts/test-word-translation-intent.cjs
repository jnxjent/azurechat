const assert = require("node:assert/strict");
const fs = require("node:fs");
const Module = require("node:module");
const path = require("node:path");
const ts = require("typescript");

const fileName = path.resolve(
  "features/chat-page/chat-services/chat-api/word-translation-intent.ts"
);
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

const { detectWholeWordTranslationLanguage } = loaded.exports;

assert.equal(
  detectWholeWordTranslationLanguage(
    "添付を中国語訳してWORDで出力して",
    true
  ),
  "zh-CN"
);
assert.equal(
  detectWholeWordTranslationLanguage("添付を英訳して", true),
  "en"
);
assert.equal(
  detectWholeWordTranslationLanguage("添付をベトナム語に翻訳して", true),
  "vi"
);
assert.equal(
  detectWholeWordTranslationLanguage("添付を中国語訳して", false),
  null
);
assert.equal(
  detectWholeWordTranslationLanguage("Wordの第1段落だけ中国語訳して", true),
  null
);
assert.equal(
  detectWholeWordTranslationLanguage("Wordの社名を置換して", true),
  null
);
assert.equal(
  detectWholeWordTranslationLanguage("Word全文の第1章を中国語訳して", true),
  "zh-CN"
);

console.log("Word translation intent tests passed.");
