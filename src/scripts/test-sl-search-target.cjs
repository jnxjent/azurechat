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

const {
  buildSlSearchTargetFilter,
  inferSlSearchTarget,
  stripSlSearchTargetTerms,
} = loadTypeScriptModule("lib/sl-search-target.ts");

assert.deepEqual(inferSlSearchTarget("個人ファイルの中から販売計画を検索して"), {
  scope: "personal",
});
assert.deepEqual(inferSlSearchTarget("全社共通から旅費規程を調べて"), {
  scope: "global_common",
});
assert.deepEqual(inferSlSearchTarget("部署共通資料から申請書を探して"), {
  scope: "dept_common",
});
assert.deepEqual(
  inferSlSearchTarget("営業改革プロジェクトフォルダーから議事録を探して"),
  { folder: "営業改革プロジェクト" }
);
assert.deepEqual(
  inferSlSearchTarget("「営業改革」フォルダーから議事録を探して"),
  { folder: "営業改革" }
);
assert.deepEqual(
  inferSlSearchTarget(
    "個人フォルダーにある「ABCD㈱財務諸表」から、この会社の財務分析をして"
  ),
  { scope: "personal" }
);
assert.deepEqual(
  inferSlSearchTarget(
    "部署共通フォルダーの中にある、ユーザーが作成した「可能性調査」というフォルダーから候補地を調べて"
  ),
  { scope: "dept_common", folder: "可能性調査" }
);
assert.deepEqual(
  inferSlSearchTarget(
    "部署共通フォルダーの中にある、ユーザーが作成した可能性調査というフォルダーから候補地を調べて"
  ),
  { scope: "dept_common", folder: "可能性調査" }
);
assert.deepEqual(
  inferSlSearchTarget("部署共通フォルダーの中にあるプロジェクト用フォルダーを調べて"),
  { scope: "dept_common", folderUncertain: true }
);
assert.deepEqual(inferSlSearchTarget("販売計画を検索して"), {});

assert.equal(
  buildSlSearchTargetFilter({ scope: "personal" }),
  "(isSlDoc eq true and slScope eq 'personal')"
);
assert.equal(
  buildSlSearchTargetFilter({ scope: "global_common", folder: "全社 規程" }),
  "(isSlDoc eq true and slScope eq 'global_common') and (isSlDoc eq true and search.ismatch('全社 規程', 'relativePath', 'simple', 'all'))"
);
assert.equal(buildSlSearchTargetFilter({ scope: "all" }), undefined);
assert.equal(
  stripSlSearchTargetTerms("個人フォルダー ABCD㈱財務諸表", {
    scope: "personal",
  }),
  "ABCD㈱財務諸表"
);
assert.equal(
  stripSlSearchTargetTerms("営業改革プロジェクトフォルダーから 議事録", {
    folder: "営業改革プロジェクト",
  }),
  "議事録"
);
assert.equal(
  stripSlSearchTargetTerms(
    "部署共通フォルダーの中にある「可能性調査」というフォルダーから 候補地",
    { scope: "dept_common", folder: "可能性調査" }
  ),
  "候補地"
);
assert.equal(
  stripSlSearchTargetTerms(
    "部署共通フォルダーの中にある ユーザーが作成した可能性調査というフォルダーから 候補地",
    { scope: "dept_common", folder: "可能性調査" }
  ),
  "候補地"
);

console.log("SharePoint search target tests passed.");
