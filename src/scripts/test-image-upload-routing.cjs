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
      esModuleInterop: true,
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

const intent = loadTypeScriptModule(
  "features/chat-page/chat-services/chat-api/image/image-intent.ts"
);
assert.equal(
  intent.resolveRequiredImageToolName(
    "添付ロゴを使って白いパッカー車をデザインして",
    1
  ),
  "edit_existing_image"
);
assert.equal(
  intent.resolveRequiredImageToolName(
    "荷台のMidacの社名を添付のロゴに差し替えて",
    1
  ),
  "edit_existing_image"
);
assert.equal(
  intent.resolveRequiredImageToolName(
    "荷台のMidacを添付ロゴに置き換えて。他は変更しないで",
    1
  ),
  "edit_existing_image"
);
assert.equal(
  intent.resolveRequiredImageToolName(
    "今の絵に「安全第一」という文字を加えて",
    1
  ),
  "add_text_to_existing_image"
);
assert.equal(
  intent.resolveRequiredImageToolName("添付画像の内容を説明して", 1),
  undefined
);
assert.equal(
  intent.resolveRequiredImageToolName("添付ロゴを使ってデザインして", 0),
  undefined
);
assert.equal(
  intent.isNewImageReferenceCompositionRequest(
    "添付ロゴを使って白いパッカー車をデザインして"
  ),
  true
);
assert.equal(
  intent.extractSharePointImageQuery("SPにあるLogoを使って"),
  "Logo"
);
assert.equal(
  intent.extractSharePointImageQuery(
    "SharePointにあるmidac_logo.pngを使ってパッカー車をデザインして"
  ),
  "midac_logo.png"
);
assert.equal(intent.extractSharePointImageQuery("添付ロゴを使って"), null);
assert.equal(intent.isSupportedImageReferenceUrl("midac_logo.png"), false);
assert.equal(
  intent.isSupportedImageReferenceUrl("https://example.com/midac_logo.png"),
  true
);
assert.equal(
  intent.isSupportedImageReferenceUrl("data:image/png;base64,iVBORw0KGgo="),
  true
);
assert.equal(
  intent.sanitizeImageLocationForLog(
    "https://account.blob.core.windows.net/dl-link/thread/logo.png?sv=1&sig=secret"
  ),
  "https://account.blob.core.windows.net/dl-link/thread/logo.png"
);
assert.equal(
  intent.sanitizeImageLocationForLog("data:image/png;base64,secret"),
  "chat-attachment:data-url"
);
assert.equal(
  intent.sanitizeImageLocationForLog("thread:__latest__.png"),
  "thread:__latest__.png"
);
assert.equal(
  intent.resolveRequiredImageToolName(
    "SPにあるLogoを使ってパッカー車をデザインして",
    1
  ),
  "edit_existing_image"
);

const imageFiles = loadTypeScriptModule(
  "features/ui/chat/chat-input-area/input-image-store.ts"
);
assert.equal(
  imageFiles.isSupportedChatImageFile({ name: "logo.png", type: "image/png" }),
  true
);
assert.equal(
  imageFiles.isSupportedChatImageFile({
    name: "rules.pdf",
    type: "application/pdf",
  }),
  false
);

const messageContent = fs.readFileSync(
  "features/chat-page/message-content.tsx",
  "utf8"
);
assert.ok(messageContent.includes('overlayText=""'));
assert.ok(!messageContent.includes("extractOverlayText(message.content)"));

const fileStore = fs.readFileSync(
  "features/chat-page/chat-input/file/file-store.ts",
  "utf8"
);
const imageFlow = fileStore.slice(
  fileStore.indexOf("public async onImageFileChange"),
  fileStore.indexOf("public async onFileChange")
);
assert.ok(imageFlow.includes("UploadDocument(formData)"));
assert.ok(imageFlow.includes("SaveLatestImageAttachment(formData)"));
assert.ok(imageFlow.includes("Indexing image metadata"));
assert.ok(imageFlow.includes("IndexDocuments("));
assert.ok(imageFlow.includes("publishToSharePoint(file"));
assert.ok(!imageFlow.includes("CrackDocument("));
assert.ok(!imageFlow.includes("CreateChatDocument("));

const chatInput = fs.readFileSync(
  "features/chat-page/chat-input/chat-input.tsx",
  "utf8"
);
assert.ok(chatInput.includes("fileStore.onImageFileChange"));
assert.ok(chatInput.includes("fileStore.onFileChange"));
assert.ok(
  chatInput.indexOf("InputImageStore.SetFile(selectedFile)") <
    chatInput.indexOf("fileStore.onImageFileChange")
);
assert.ok(chatInput.includes('disabled={loading !== "idle"}'));
assert.ok(chatInput.includes('if (loading === "file upload")'));

const extensions = fs.readFileSync(
  "features/chat-page/chat-services/chat-api/chat-api-extension.ts",
  "utf8"
);
assert.ok(extensions.includes("function: { name: requiredToolName }"));

const chatApi = fs.readFileSync(
  "features/chat-page/chat-services/chat-api/chat-api.ts",
  "utf8"
);
assert.ok(chatApi.includes("imageAttachmentCountForRouting"));
assert.ok(chatApi.includes("LoadLatestImageAttachment(currentChatThread.id)"));
assert.ok(chatApi.includes("referencesSharePointImage"));

const imageExtensions = fs.readFileSync(
  "features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts",
  "utf8"
);
assert.ok(imageExtensions.includes("timeout: requestTimeoutMs"));
assert.ok(imageExtensions.includes("maxRetries: 0"));
assert.ok(imageExtensions.includes("inputBytes: imageBuffers.map"));
const referenceSelection = imageExtensions.slice(
  imageExtensions.indexOf("const inferredReferenceUrls"),
  imageExtensions.indexOf("const referenceUrls")
);
assert.ok(!referenceSelection.includes("legacyImageUrl"));
assert.ok(imageExtensions.includes("currentAttachmentUrls.length > 0"));
assert.ok(imageExtensions.includes("candidateReferenceUrls"));
assert.ok(imageExtensions.includes("normalizeAzureEditImage"));
assert.ok(imageExtensions.includes("imageContentHash"));
assert.ok(imageExtensions.includes("URL did not return a supported image"));
assert.ok(imageExtensions.includes("Failed to convert WebP to PNG"));
assert.ok(imageExtensions.includes("LoadLatestImageAttachment(chatThread.id)"));
assert.ok(imageExtensions.includes("storedAttachment?.fileName"));
assert.ok(imageExtensions.includes("ConsumeLatestImageAttachment(chatThread.id)"));
assert.ok(imageExtensions.includes("extractSharePointImageQuery(userMessage)"));
assert.ok(imageExtensions.includes("sharePointReference:"));
assert.ok(imageExtensions.includes("readImageBufferFromConfiguredBlob"));
assert.ok(imageExtensions.includes("SP image loaded via Blob SDK"));
assert.ok(imageExtensions.includes("Image URL fetch failed"));
assert.ok(imageExtensions.includes("sanitizeImageLocationForLog(resolvedBaseUrl)"));
assert.ok(imageExtensions.includes("isSupportedImageReferenceUrl"));
assert.ok(
  imageExtensions.includes("Ignored invalid model referenceImageUrls")
);
assert.ok(
  imageExtensions.includes(
    "Ignored model referenceImageUrls because a SharePoint image was resolved"
  )
);
assert.ok(
  imageExtensions.includes("sharePointImageReference\n    ? []")
);
const candidateReferenceSelection = imageExtensions.slice(
  imageExtensions.indexOf("const candidateReferenceUrls"),
  imageExtensions.indexOf("const referenceUrls")
);
assert.ok(
  !candidateReferenceSelection.includes("sharePointImageReference.resolvedUrl")
);
assert.ok(
  imageExtensions.includes(
    "if (sharePointImageBuffer && !sharePointImageUsedAsBase)"
  )
);
assert.ok(imageExtensions.includes("[edit_existing_image] output saved:"));

const openAiStream = fs.readFileSync(
  "features/chat-page/chat-services/chat-api/open-ai-stream.ts",
  "utf8"
);
assert.ok(openAiStream.includes('.on("end"'));
assert.ok(openAiStream.includes("buildToolResultFallbackContent"));
assert.ok(openAiStream.includes("runner ended without finalContent"));

const chatStoreSource = fs.readFileSync(
  "features/chat-page/chat-store.tsx",
  "utf8"
);
const finalContentHandler = chatStoreSource.slice(
  chatStoreSource.indexOf('case "finalContent"'),
  chatStoreSource.indexOf("default:", chatStoreSource.indexOf('case "finalContent"'))
);
assert.ok(finalContentHandler.includes("assistantResponseId ?? uniqueId()"));
assert.ok(finalContentHandler.includes("this.addToMessages(mappedFinalContent)"));
const baseSelection = imageExtensions.slice(
  imageExtensions.indexOf("let resolvedBaseUrl"),
  imageExtensions.indexOf("if (!inputBuffer) {")
);
assert.ok(baseSelection.includes("let inputBuffer = latestBuffer"));
assert.ok(
  baseSelection.indexOf("let inputBuffer = latestBuffer") <
    baseSelection.indexOf("explicitBaseUrl")
);

console.log("Image upload/edit regression tests: 71 assertions passed");
