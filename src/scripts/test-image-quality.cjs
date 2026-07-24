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

const { normalizeGptImageQuality } = loadTypeScriptModule(
  "features/chat-page/chat-services/chat-api/image/image-quality.ts"
);

assert.equal(normalizeGptImageQuality("high"), "high");
assert.equal(normalizeGptImageQuality("HIGH"), "high");
assert.equal(normalizeGptImageQuality(" medium "), "medium");
assert.equal(normalizeGptImageQuality("low"), "low");
assert.equal(normalizeGptImageQuality("auto"), "auto");
assert.equal(normalizeGptImageQuality(undefined), "auto");
assert.equal(normalizeGptImageQuality("hd"), "auto");
assert.equal(normalizeGptImageQuality("高画質"), "auto");

const extensionSource = fs.readFileSync(
  "features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts",
  "utf8"
);
const qualityEnums = extensionSource.match(
  /enum: \["low", "medium", "high", "auto"\]/g
);
assert.equal(qualityEnums?.length, 2, "create and edit schemas need quality");
const imageApiOptions = extensionSource.match(
  /size: normalizeGptImageSize\(args\?\.size\),\s*quality,/g
);
assert.ok(
  (imageApiOptions?.length ?? 0) >= 2,
  "create and edit API calls need quality"
);
assert.match(
  extensionSource,
  /const quality = normalizeGptImageQuality\(args\?\.quality\)/
);

console.log("GPT Image quality tests passed.");
