const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");
const ts = require("typescript");

function loadOfficeService() {
  const fileName = path.resolve("features/teams/teams-office-service.ts");
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
  const originalRequire = loaded.require.bind(loaded);
  loaded.require = (request) => {
    if (request === "server-only") return {};
    if (request === "@/features/common/services/azure-storage") return {};
    if (request === "@/features/pptx/palette") {
      return {
        pptxPaletteListText: () => "",
        resolvePptxPaletteInstruction: () => null,
      };
    }
    if (
      request === "./teams-search-service" ||
      request === "./teams-ppt-plan-service" ||
      request === "./teams-word-proofread-service"
    ) {
      return {};
    }
    return originalRequire(request);
  };
  loaded._compile(javascript, fileName);
  return loaded.exports;
}

const { parseTeamsOfficeRequest } = loadOfficeService();

assert.equal(
  parseTeamsOfficeRequest(
    "添付PDFをExcelに変換して\n添付ファイル: report.pdf"
  )?.action,
  "pdf_to_excel"
);
assert.equal(
  parseTeamsOfficeRequest(
    "このファイルをWordに変換して\n添付ファイル: report.pdf"
  )?.action,
  "pdf_to_word"
);
assert.equal(
  parseTeamsOfficeRequest(
    "PowerPointに変換して\n添付ファイル: report.pdf"
  )?.action,
  "pdf_to_ppt"
);

const refine = parseTeamsOfficeRequest(
  "P2のシートを再変換して\n添付ファイル: report.xlsx"
);
assert.equal(refine?.action, "refine_excel_sheets");
assert.deepEqual(refine?.targetSheets, ["P2"]);

assert.equal(
  parseTeamsOfficeRequest(
    "このExcelをグラフにして\n添付ファイル: report.xlsx"
  )?.action,
  "edit_latest_excel"
);
assert.equal(
  parseTeamsOfficeRequest(
    "このWordの誤字を修正して\n添付ファイル: report.docx"
  )?.action,
  "edit_latest_word"
);

console.log("Teams Office attachment routing tests passed.");
