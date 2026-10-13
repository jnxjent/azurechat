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
  for (const text of ['画像生成用プロンプトを考えて', 'PPTのデザインを変更して', '画像の生成方法を教えて', '社内の日報を検索して']) {
    assert.equal(intent.isTeamsImageGenerationRequest(text), false);
  }
  const routing = load('features/common/services/salesforce-routing.ts');
  const access = { isSalesforceAllowed: true, hasSalesforceExtension: true };
  assert.equal(routing.resolveSalesforceRoute({ ...access, message: '私の部下の最近の日報を挙げて' }).route, 'normal');
  assert.equal(routing.resolveSalesforceRoute({ ...access, message: 'SFで私の部下の最近の日報を挙げて' }).route, 'salesforce');
  assert.equal(routing.resolveSalesforceRoute({ ...access, isSalesforceAllowed: false, message: 'Salesforceで日報を一覧表示して' }).route, 'denied');
  let imageCalls = 0, signedPath;
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
      GenerateSasUrl: async (container, blob) => {
        assert.equal(container, 'images'); signedPath = blob;
        return { status: 'OK', response: 'https://blob.example/generated.png?sig=test' };
      },
    },
  });
  const props = { message: truck, conversationId: 'A', activityId: 'event1' };
  const [first, duplicate] = await Promise.all([service.createTeamsImageReply(props), service.createTeamsImageReply(props)]);
  assert.equal(imageCalls, 1);
  assert.equal(first, duplicate);
  assert.equal(signedPath, `teams-${require('crypto').createHash('sha256').update('A').digest('hex').slice(0, 32)}/generated.png`);
  assert.match(first, /blob\.example/);
  assert.doesNotMatch(first, /ac\.example/);

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
