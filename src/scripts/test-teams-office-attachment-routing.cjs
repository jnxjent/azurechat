const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");
const ts = require("typescript");

function loadOfficeService() {
  const storedBlobs = new Map();
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
    if (request === "@/features/common/services/azure-storage") {
      return {
        DownloadBlobAsText: async (container, blobName) => {
          const value = storedBlobs.get(`${container}/${blobName}`);
          return value === undefined
            ? { status: "ERROR", response: "" }
            : { status: "OK", response: value };
        },
        GenerateSasUrl: async () => ({
          status: "OK",
          response: "https://example.test/generated",
        }),
        UploadBlob: async (container, blobName, buffer) => {
          storedBlobs.set(`${container}/${blobName}`, buffer.toString("utf8"));
          return { status: "OK" };
        },
      };
    }
    if (request === "@/features/pptx/palette") {
      return {
        pptxPaletteListText: () => "",
        resolvePptxPaletteInstruction: () => null,
      };
    }
    if (request === "@/features/auth-page/helpers") {
      return { hashValue: (value) => value };
    }
    if (request === "@/lib/sl-dept") {
      return { resolveSlAccess: () => ({ dept: "bm" }) };
    }
    if (
      request ===
      "@/features/chat-page/chat-services/sharepoint-summary-service"
    ) {
      return { summarizeSharePointPdf: async () => ({}) };
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

const {
  buildTeamsThreadId,
  executeTeamsOfficeRequest,
  parseTeamsOfficeRequest,
  registerTeamsUploadedOfficeFiles,
} = loadOfficeService();

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
const summaryRequest = parseTeamsOfficeRequest(
  "SharePointにある「令和7年度環境白書.pdf」を先頭から最終ページまで全文要約し、約10ページ（7,000～8,000文字）のWordファイルにしてください"
);
assert.equal(summaryRequest?.action, "summarize_sp_pdf");
assert.equal(summaryRequest?.fileQuery, "令和7年度環境白書.pdf");
assert.equal(summaryRequest?.targetPages, 10);
assert.equal(summaryRequest?.targetCharsLow, 7000);
assert.equal(summaryRequest?.targetCharsHigh, 8000);
assert.equal(
  parseTeamsOfficeRequest(
    "PowerPointに変換して\n添付ファイル: report.pdf"
  )?.action,
  "pdf_to_ppt"
);

const chineseTranslation = parseTeamsOfficeRequest(
  "添付PDFの日本語を中国語に翻訳して\n添付ファイル: ごみ分別.pdf"
);
assert.equal(chineseTranslation?.action, "translate_pdf_to_pptx");
assert.equal(chineseTranslation?.targetLanguage, "zh-CN");

const portugueseTranslation = parseTeamsOfficeRequest(
  "次は同じものをポルトガル語に変換して"
);
assert.equal(portugueseTranslation?.action, "translate_pdf_to_pptx");
assert.equal(portugueseTranslation?.targetLanguage, "pt");

for (const [languageName, languageCode] of [
  ["英語", "en"],
  ["ポルトガル語", "pt"],
  ["ベトナム語", "vi"],
  ["インドネシア語", "id"],
  ["中国語", "zh-CN"],
  ["韓国語", "ko"],
  ["スペイン語", "es"],
  ["タガログ語", "fil"],
]) {
  assert.equal(
    parseTeamsOfficeRequest(
      `添付PDFを${languageName}に翻訳して\n添付ファイル: guide.pdf`
    )?.targetLanguage,
    languageCode
  );
}

assert.equal(
  parseTeamsOfficeRequest(
    "添付を英語にして編集可能なPPTXで出力して\n添付ファイル: guide.pdf"
  )?.action,
  "translate_pdf_to_pptx"
);
assert.equal(
  parseTeamsOfficeRequest("中国語の営業資料をPowerPointで新規作成して")
    ?.action,
  "create_ppt"
);
assert.equal(
  parseTeamsOfficeRequest("P2のタイトルだけを英語に変更して")?.action,
  "edit_latest_ppt"
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

async function testPdfTranslationExecutionAndFollowup() {
  const conversationId = "translation-test-conversation";
  const teamsThreadId = buildTeamsThreadId(conversationId);
  const sourceFile = {
    extension: "pdf",
    fileName: "ごみ分別.pdf",
    savedAt: Date.now(),
    size: 100,
    url: "https://example.test/source.pdf?sig=test",
  };
  await registerTeamsUploadedOfficeFiles(teamsThreadId, [sourceFile]);

  const requests = [];
  const originalFetch = global.fetch;
  global.fetch = async (_url, init) => {
    requests.push(JSON.parse(init.body));
    const targetLanguage = requests.at(-1).targetLanguage;
    return {
      ok: true,
      status: 200,
      json: async () => ({
        downloadUrl: `https://example.test/${targetLanguage}.pptx`,
        fileName: `ごみ分別_${targetLanguage}.pptx`,
        pages: 2,
      }),
    };
  };

  try {
    const firstReply = await executeTeamsOfficeRequest({
      request: chineseTranslation,
      conversationId,
      uploadedFiles: [sourceFile],
    });
    assert.match(firstReply, /中国語（簡体字）/);
    assert.equal(requests[0].action, "translate_pdf_to_pptx");
    assert.equal(requests[0].targetLanguage, "zh-CN");
    assert.equal(requests[0].fileUrl, sourceFile.url);

    const followupReply = await executeTeamsOfficeRequest({
      request: portugueseTranslation,
      conversationId,
      uploadedFiles: [],
    });
    assert.match(followupReply, /ポルトガル語/);
    assert.equal(requests[1].targetLanguage, "pt");
    assert.equal(requests[1].fileUrl, sourceFile.url);
  } finally {
    global.fetch = originalFetch;
  }
}

testPdfTranslationExecutionAndFollowup()
  .then(() => {
    console.log("Teams Office attachment routing tests passed.");
  })
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
