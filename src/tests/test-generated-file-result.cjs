const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const ts = require("typescript");

const sourcePath = path.join(
  __dirname,
  "..",
  "features",
  "chat-page",
  "chat-services",
  "chat-api",
  "generated-file-result.ts"
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
const { hasGeneratedFileResult } = loaded.exports;

const json = (value) => JSON.stringify(value);

assert.equal(
  hasGeneratedFileResult([
    json({ fileName: "source.docx", fileUrl: "https://example/source.docx" }),
  ]),
  false,
  "a search/source result without downloadUrl must retain citations"
);
assert.equal(
  hasGeneratedFileResult([
    json({ downloadUrl: "https://example/output", fileName: "edited.docx" }),
  ]),
  true,
  "a generated Word result must suppress search citations"
);
assert.equal(
  hasGeneratedFileResult([
    json({ downloadUrl: "https://example/%E4%BF%AE%E6%AD%A3%E7%89%88.docx?sig=x" }),
  ]),
  true,
  "an encoded generated Word URL must be recognized"
);
assert.equal(
  hasGeneratedFileResult([
    json({ downloadUrl: "https://example/output", message: "Word編集が完了しました" }),
  ]),
  true,
  "a generated file message may identify an extensionless download URL"
);
assert.equal(
  hasGeneratedFileResult([
    json({
      downloadUrl: "https://example/failed.docx",
      fileName: "failed.docx",
      error: "failed",
    }),
  ]),
  false,
  "an error result must not suppress citations"
);

console.log("Generated file result regression tests: 5 assertions passed");
