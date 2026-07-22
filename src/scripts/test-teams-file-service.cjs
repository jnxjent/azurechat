const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");
const ts = require("typescript");

function transpileTypeScript(fileName) {
  return ts.transpileModule(fs.readFileSync(fileName, "utf8"), {
    compilerOptions: {
      esModuleInterop: true,
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
    },
  }).outputText;
}

function loadTypeScriptModule(fileName, mocks = {}) {
  const absolutePath = path.resolve(fileName);
  const loaded = new Module(absolutePath);
  loaded.filename = absolutePath;
  loaded.paths = Module._nodeModulePaths(path.dirname(absolutePath));
  const originalRequire = loaded.require.bind(loaded);
  loaded.require = (request) =>
    Object.prototype.hasOwnProperty.call(mocks, request)
      ? mocks[request]
      : originalRequire(request);
  loaded._compile(transpileTypeScript(absolutePath), absolutePath);
  return loaded.exports;
}

async function main() {
  const policy = loadTypeScriptModule(
    "features/teams/teams-file-policy.ts"
  );
  const uploads = [];
  let pointerText = "";
  const storage = {
    async UploadBlob(container, blobPath, content) {
      const buffer = Buffer.from(content);
      uploads.push({ blobPath, buffer, container });
      if (blobPath.endsWith("-teams-upload-latest.json")) {
        pointerText = buffer.toString("utf8");
      }
      return { status: "OK", response: blobPath };
    },
    async GenerateSasUrl(container, blobPath) {
      assert.equal(container, "dl-link");
      return {
        status: "OK",
        response: `https://storage.example.test/${encodeURIComponent(blobPath)}`,
      };
    },
    async DownloadBlobAsText(container, blobPath) {
      assert.equal(container, "dl-link");
      assert.ok(blobPath.endsWith("-teams-upload-latest.json"));
      return { status: "OK", response: pointerText };
    },
  };
  const service = loadTypeScriptModule(
    "features/teams/teams-file-service.ts",
    {
      "server-only": {},
      "@/features/common/services/azure-storage": storage,
      "./teams-file-policy": policy,
    }
  );

  const originalFetch = global.fetch;
  const pdf = Buffer.from("%PDF-1.7\nlocal integration test");
  let fetchedUrl = "";
  global.fetch = async (url, options) => {
    fetchedUrl = String(url);
    assert.equal(options.headers.Accept, "application/octet-stream");
    assert.equal(options.redirect, "follow");
    assert.ok(options.signal);
    const response = new Response(pdf, {
      status: 200,
      headers: { "content-length": String(pdf.length) },
    });
    Object.defineProperty(response, "url", {
      value: "https://tenant.sharepoint.com/download/report.pdf",
    });
    return response;
  };

  try {
    const stored = await service.receiveTeamsFiles({
      attachments: [
        {
          contentType: policy.TEAMS_FILE_DOWNLOAD_CONTENT_TYPE,
          name: "report.pdf",
          content: {
            downloadUrl:
              "https://tenant.sharepoint.com/download/report.pdf?token=test",
            fileType: "pdf",
            uniqueId: "drive-item-1",
          },
        },
      ],
      threadId: "local-integration-test",
    });

    assert.equal(
      fetchedUrl,
      "https://tenant.sharepoint.com/download/report.pdf?token=test"
    );
    assert.equal(stored.length, 1);
    assert.equal(stored[0].fileName, "report.pdf");
    assert.equal(stored[0].extension, "pdf");
    assert.equal(stored[0].size, pdf.length);
    assert.ok(stored[0].url.startsWith("https://storage.example.test/"));

    assert.equal(uploads.length, 2);
    assert.equal(uploads[0].container, "dl-link");
    assert.match(
      uploads[0].blobPath,
      /^teams-upload\/local-integration-test\/.*-report\.pdf$/
    );
    assert.deepEqual(uploads[0].buffer, pdf);
    assert.equal(
      uploads[1].blobPath,
      "thread-local-integration-test-teams-upload-latest.json"
    );

    const latest = await service.readLatestTeamsFiles(
      "local-integration-test"
    );
    assert.deepEqual(latest, stored);

    const uploadsBeforeInvalidFile = uploads.length;
    global.fetch = async () => {
      const response = new Response("not a pdf", { status: 200 });
      Object.defineProperty(response, "url", {
        value: "https://tenant.sharepoint.com/download/fake.pdf",
      });
      return response;
    };
    await assert.rejects(
      service.receiveTeamsFiles({
        attachments: [
          {
            contentType: policy.TEAMS_FILE_DOWNLOAD_CONTENT_TYPE,
            name: "fake.pdf",
            content: {
              downloadUrl: "https://tenant.sharepoint.com/download/fake.pdf",
              fileType: "pdf",
            },
          },
        ],
        threadId: "local-integration-test",
      })
    );
    assert.equal(uploads.length, uploadsBeforeInvalidFile);
  } finally {
    global.fetch = originalFetch;
  }

  console.log("Teams file service integration tests passed.");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
