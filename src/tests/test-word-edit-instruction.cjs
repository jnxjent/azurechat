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
  "word-edit-instruction.ts"
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
  hasExplicitWordReplacementInstruction,
  resolveWordEditInstruction,
} = loaded.exports;

assert.equal(
  hasExplicitWordReplacementInstruction("「太平興産」を「大平興産」に置換"),
  true
);
assert.equal(
  hasExplicitWordReplacementInstruction("（誤）とういう考え方\n（正）という考え方"),
  true
);
assert.equal(
  hasExplicitWordReplacementInstruction("誤字・誤記を修正履歴とともに修正してください"),
  false
);

const broadUserRequest =
  "SharePointにある議事録の誤字・誤記を修正履歴とともに修正したWordで出力してください";
const reviewedPairs =
  "「太平興産」を「大平興産」に置換、「稼働率とういう考え方」を「稼働率という考え方」に置換";
assert.deepEqual(resolveWordEditInstruction(reviewedPairs, broadUserRequest), {
  instruction: reviewedPairs,
  source: "tool-explicit",
});

const explicitUserRequest = "「大平興産」を「太平興産」に置換してください";
assert.deepEqual(resolveWordEditInstruction(reviewedPairs, explicitUserRequest), {
  instruction: explicitUserRequest,
  source: "user-explicit",
});

assert.deepEqual(resolveWordEditInstruction("表題を太字にする", broadUserRequest), {
  instruction: broadUserRequest,
  source: "user-original",
});

console.log("Word edit instruction regression tests: 6 assertions passed");
