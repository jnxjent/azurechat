const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');

function load(file, mocks = {}, globals = {}) {
  const exports = {};
  const code = ts.transpileModule(fs.readFileSync(path.resolve(file), 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
  }).outputText;
  vm.runInNewContext(code, { exports, require(name) {
    if (name === 'server-only') return {};
    if (name in mocks) return mocks[name];
    return require(name);
  }, process, console, Buffer, AbortSignal, URL, ...globals });
  return exports;
}

async function main() {
  const image = load('features/chat-page/chat-services/chat-api/image/image-intent.ts');
  const logoRequest = '今出した廃棄物処理車の荷台の「Midac」という文字を添付Logoに差し替えて絵を出して';
  assert.equal(image.isPptxAssetPlacementRequest(logoRequest), false);
  assert.equal(image.resolveRequiredImageToolName(logoRequest, 1), 'edit_existing_image');
  for (const message of ['今のPPTのロゴを添付Logoに差し替えて', 'このスライドに添付ロゴを入れて']) {
    assert.equal(image.isPptxAssetPlacementRequest(message), true);
  }
  const intent = load('features/teams/teams-image-intent.ts');
  const truck = '産業廃棄物収取運搬車（パッカー車）を以下でデザインして。\n白基調で、荷台に社名として、Midacと入れて';
  assert.equal(intent.isTeamsImageGenerationRequest(truck), true);
  assert.equal(intent.isTeamsImageGenerationRequest('猫の絵を描いて'), true);
  assert.equal(intent.isTeamsImageGenerationRequest(logoRequest), true);
  assert.equal(intent.isTeamsImageEditRequest(logoRequest), true);
  assert.equal(intent.isTeamsImageEditRequest('白い車の絵を描いて、Midacと入れて'), false);
  for (const text of ['画像生成用プロンプトを考えて', 'PPTのデザインを変更して', '画像の生成方法を教えて', '社内の日報を検索して']) {
    assert.equal(intent.isTeamsImageGenerationRequest(text), false);
  }
  const routing = load('features/common/services/salesforce-routing.ts');
  const access = { isSalesforceAllowed: true, hasSalesforceExtension: true };
  assert.equal(routing.resolveSalesforceRoute({ ...access, message: '私の部下の最近の日報を挙げて' }).route, 'normal');
  assert.equal(routing.resolveSalesforceRoute({ ...access, message: 'SFで私の部下の最近の日報を挙げて' }).route, 'salesforce');
  assert.equal(routing.resolveSalesforceRoute({ ...access, isSalesforceAllowed: false, message: 'Salesforceで日報を一覧表示して' }).route, 'denied');
  let imageCalls = 0, editCalls = 0;
  const signedPaths = [];
  const thread = `teams-${require('crypto').createHash('sha256').update('A').digest('hex').slice(0, 32)}`;
  const baseBytes = Buffer.from('original truck');
  const logoBytes = Buffer.from('attached MIDAC logo');
  const sharpMock = buffer => ({ rotate() { return this; }, png() { return this; },
    resize() { return this; }, jpeg() { return this; }, toBuffer: async () => buffer });
  let missingBase = false;
  process.env.AZURE_STORAGE_ACCOUNT_NAME = 'teststore';
  const service = load('features/teams/teams-image-service.ts', {
    '@/features/chat-page/chat-services/chat-api/image/image-generation-service': {
      generateStoredChatImage: async (args) => {
        imageCalls++;
        assert.equal(args.prompt, truck);
        assert.equal(args.userMessage, truck);
        assert.equal(args.threadId, `teams-${require('crypto').createHash('sha256').update('A').digest('hex').slice(0, 32)}`);
        return { fileName: 'generated.png', url: 'https://ac.example/api/images' };
      },
    },
    '@/features/common/services/azure-storage': {
      GenerateSasUrl: async (container, blob, type) => {
        assert.equal(container, 'images'); signedPaths.push([blob, type]);
        return { status: 'OK', response: `https://blob.example/${blob}?sig=test` };
      },
      UploadBlob: async (container, name, buffer, returnName, type) => {
        assert.equal(type, 'image/jpeg'); return { status: 'OK' };
      },
      GetBlob: async (container, name) => {
        assert.equal(container, 'dl-link');
        assert.ok(name.startsWith(`teams-upload/${thread}/`));
        return { status: 'OK', response: (async function* () { yield logoBytes; })() };
      },
    },
    '@/features/chat-page/chat-services/chat-image-service': {
      GetImageFromStore: async () => missingBase ? { status: 'NOT_FOUND' } :
        { status: 'OK', response: (async function* () { yield baseBytes; })() },
      UploadImageToStore: async () => ({ status: 'OK' }),
    },
    '@/features/chat-page/chat-services/chat-api/image/image-intent': image,
    './teams-image-intent': intent,
    '@/features/common/services/openai': { OpenAIDALLEInstance: () => ({ images: { edit: async input => {
      editCalls++;
      assert.equal(input.image.length, 2);
      assert.ok(input.image[0].equals(baseBytes));
      assert.ok(input.image[1].equals(logoBytes));
      assert.match(input.prompt, /形・配色・文字・綴り/);
      return { data: [{ b64_json: Buffer.from('edited truck').toString('base64') }] };
    } } }) },
    '@/features/common/util': { uniqueId: () => 'edited' },
    openai: { toFile: async buffer => buffer },
    sharp: { default: sharpMock },
  });
  const props = { message: truck, conversationId: 'A', activityId: 'event1' };
  const [first, duplicate] = await Promise.all([service.createTeamsImageReply(props), service.createTeamsImageReply(props)]);
  assert.equal(imageCalls, 1);
  assert.equal(first, duplicate);
  assert.equal(first.attachments[0].contentType, 'image/jpeg');
  assert.match(first.attachments[0].contentUrl, /preview\.jpg/);
  assert.ok(signedPaths.some(([blob, type]) => blob === `${thread}/generated.png` && type === 'image/png'));
  assert.match(first.text, /blob\.example/);
  assert.doesNotMatch(first.text, /ac\.example/);
  const uploadedFiles = [{ extension: 'png', fileName: 'logo.png', size: logoBytes.length,
    savedAt: Date.now(), url: `https://teststore.blob.core.windows.net/dl-link/teams-upload/${thread}/logo.png?sig=test` }];
  const edited = await service.createTeamsImageReply({ ...props, activityId: 'edit1', message: logoRequest, uploadedFiles });
  assert.equal(editCalls, 1);
  assert.equal(imageCalls, 1, 'edit must not call fresh generation');
  assert.match(edited.text, /編集/);
  assert.match(await service.createTeamsImageReply({ ...props, activityId: 'edit2', message: logoRequest }), /再添付/);
  missingBase = true;
  assert.match(await service.createTeamsImageReply({ ...props, activityId: 'edit3', message: logoRequest, uploadedFiles }), /編集対象/);
  assert.equal(editCalls, 1);

  let receivedUrl, receivedHeaders, summary;
  const sf = load('features/teams/teams-salesforce-service.ts', {
    '@/features/common/services/salesforce-routing': routing,
    '@/features/common/services/openai': { OpenAIInstance: () => ({ chat: { completions: {
      create: async (request) => { summary = request; return { choices: [{ message: { content: '日報一覧です' } }] }; },
    } } }) },
  }, { fetch: async (url, options) => {
    receivedUrl = new URL(url); receivedHeaders = options.headers;
    return { ok: true, json: async () => ({ ok: true, count: 1, total_count: 2000,
      items: [{ ReportContent: 'x'.repeat(9000) }], table_instruction: '対象はGatewayで絞り込み済みです' }) };
  } });
  await sf.queryTeamsSalesforce({ message: 'SFで私の部下の日報を一覧表示して', userEmail: 'user@example.com', conversationId: 'A', activityId: 'sf1' });
  assert.equal(receivedUrl.searchParams.get('q'), '私の部下の日報を一覧表示して');
  assert.equal(receivedHeaders['X-User-Email'], 'user@example.com');
  assert.match(summary.messages[0].content, /Gatewayで絞り込み済み/);
  const json = summary.messages[1].content.split('Gateway JSON:\n')[1];
  assert.equal(JSON.parse(json).items[0].ReportContent.length, 9000);
  assert.equal(JSON.parse(json).total_count, 2000);

  let generationInput, failStorage = false;
  const writes = [];
  const generation = load('features/chat-page/chat-services/chat-api/image/image-generation-service.ts', {
    '@/features/common/services/openai': { OpenAIDALLEInstance: () => ({ images: { generate: async input => {
      generationInput = input; return { data: [{ b64_json: Buffer.from('png').toString('base64') }] };
    } } }) },
    '@/features/common/util': { uniqueId: () => 'image-id' },
    '../../chat-image-service': { UploadImageToStore: async (thread, name) => {
      writes.push([thread, name]); return { status: failStorage ? 'ERROR' : 'OK' };
    } },
    './image-intent': image,
    './image-quality': load('features/chat-page/chat-services/chat-api/image/image-quality.ts'),
  });
  const generated = await generation.generateStoredChatImage({ threadId: 'A', userMessage: truck, prompt: truck });
  assert.equal(generated.fileName, 'image-id.png');
  assert.equal(generationInput.prompt, truck);
  assert.equal(writes.length, 3);
  assert.equal(writes[1][1], '__base__.png');
  assert.equal(writes[2][1], '__latest__.png');
  failStorage = true;
  assert.ok('error' in await generation.generateStoredChatImage({ threadId: 'A', userMessage: truck, prompt: truck }));
  console.log('Teams SF/image fixes: routing, logo edit, image delivery/deduplication, SF scope/integrity passed');
}
main().catch(error => { console.error(error); process.exitCode = 1; });
