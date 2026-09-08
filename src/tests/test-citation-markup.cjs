const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const sourcePath = path.join(
  __dirname,
  "..",
  "features",
  "ui",
  "markdown",
  "citation-markup.ts"
);
const source = fs.readFileSync(sourcePath, "utf8");
const compiled = ts.transpileModule(source, {
  compilerOptions: {
    module: ts.ModuleKind.CommonJS,
    target: ts.ScriptTarget.ES2022,
  },
  fileName: sourcePath,
});
const loaded = new Module(sourcePath, module);
loaded.filename = sourcePath;
loaded.paths = Module._nodeModulePaths(path.dirname(sourcePath));
loaded._compile(compiled.outputText, sourcePath);

const {
  removeCitationMarkup,
  removeOpenAIInternalCitationMarkup,
} = loaded.exports;

const fileCitation = "\uE200filecite\uE202turn0file0\uE201";
const searchCitation =
  "\uE200cite\uE202turn0search0\uE202turn0search1\uE201";

assert.equal(
  removeOpenAIInternalCitationMarkup(`変換しました。${fileCitation}`),
  "変換しました。",
  "a leaked file citation must be removed"
);
assert.equal(
  removeOpenAIInternalCitationMarkup(`回答です。 ${searchCitation}`),
  "回答です。",
  "a leaked multi-source search citation must be removed"
);
assert.equal(
  removeOpenAIInternalCitationMarkup(
    `変換しました。\n${fileCitation}\n\n次の説明です。`
  ),
  "変換しました。\n\n次の説明です。",
  "removing a citation must not leave excessive blank lines"
);
assert.equal(
  removeOpenAIInternalCitationMarkup(
    `通常引用 {% citation items=[{name:\"資料.pdf\",id:\"doc-1\"}] /%}`
  ),
  `通常引用 {% citation items=[{name:\"資料.pdf\",id:\"doc-1\"}] /%}`,
  "AzureChat citation markup must remain available to the UI"
);
assert.equal(
  removeCitationMarkup(
    `完了しました。${fileCitation}\n{% citation items=[{name:\"資料.pdf\",id:\"doc-1\"}] /%}`
  ),
  "完了しました。",
  "server normalization must remove both citation formats"
);

console.log("Citation markup regression tests: 5 assertions passed");
