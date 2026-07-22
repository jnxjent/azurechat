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

const { extractWeatherPageText, isTrustedWeatherUrl } = loadTypeScriptModule(
  "features/teams/teams-weather-page.ts"
);

assert.equal(isTrustedWeatherUrl("https://www.jma.go.jp/bosai/forecast/"), true);
assert.equal(isTrustedWeatherUrl("https://weather.yahoo.co.jp/weather/"), true);
assert.equal(isTrustedWeatherUrl("https://example.com/weather"), false);
assert.equal(isTrustedWeatherUrl("http://www.jma.go.jp/weather"), false);

const text = extractWeatherPageText(`
  <html><head><style>.hidden { display:none }</style><script>bad()</script></head>
  <body><h1>東京の明日の天気</h1><div>晴れ</div>
  <p>最高気温 36&#8451; / 最低気温 27℃</p>
  <p>降水確率 10&amp;#37;</p></body></html>
`);
assert.match(text, /東京の明日の天気/);
assert.match(text, /最高気温 36℃/);
assert.match(text, /最低気温 27℃/);
assert.match(text, /降水確率/);
assert.doesNotMatch(text, /bad\(\)/);

console.log("Teams weather page extraction tests passed.");
