const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const root = path.resolve(__dirname, '..');
const transpile = s => ts.transpileModule(s, { compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.CommonJS } }).outputText;
const identity = { exports: {} };
vm.runInNewContext(transpile(fs.readFileSync(path.join(root, 'features/pptx/company-identity.ts'), 'utf8')), identity);
const helpers = identity.exports;
const prompt = '当社㈱ミダックホールディングスは産業廃棄物一貫処理会社です。初回客先訪問用の営業資料を8枚のスライドで作成してください。内容はHPから入手してください。';
assert.equal(helpers.extractPresentationCompanyName('初回客先訪問用 営業資料', prompt), 'ミダックホールディングス');
assert.equal(helpers.extractPresentationCompanyName('営業資料', prompt.normalize('NFKC')), 'ミダックホールディングス');
assert.equal(helpers.extractPresentationCompanyName('初回客先訪問用 営業資料', '会社紹介してください'), '');
helpers.assertCompanyIdentity('㈱ミダックホールディングス', '株式会社ミダックホールディングス');
assert.throws(() => helpers.assertCompanyIdentity('ミダックホールディングス', '楽天グループ株式会社'));

const source = fs.readFileSync(path.join(root, 'features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts'), 'utf8');
// Bound the real functions without loading authentication or Azure clients.
const collectBody = source.slice(source.indexOf('async function collectWebEvidence('), source.indexOf('\n}', source.indexOf('async function collectWebEvidence(')) + 2);
const sharedBody = source.slice(source.indexOf('export async function createSharedCompanyProfilePptPlan('), source.indexOf('\n}', source.indexOf('): Promise<SharedCompanyProfilePptPlan> {')) + 2);
let results = [], pages = {}, actualCompany = 'ミダックホールディングス', plans = 0;
const context = {
  exports: {}, ...helpers, console, URL, URLSearchParams, AbortController, setTimeout, clearTimeout,
  process: { env: { BRAVE_SUBSCRIPTION_TOKEN: 'test-double' } },
  fetch: async () => ({ ok: true, json: async () => ({ web: { results } }) }),
  fetchPageText: async url => pages[url] || '',
  extractCompanyNameFromTitle: helpers.extractPresentationCompanyName,
  buildCompanyBrief: async () => ({ companyName: actualCompany }),
  extractRequestedTotalSlideCount: () => 8,
  planCompanyProfileSlides: async () => { plans++; return Array.from({ length: 7 }, () => ({ title: '事業内容' })); },
  PPT_PRODUCTION_INSTRUCTION: '', removePptProductionNotes: v => v, parsePromptIntent: () => ({}),
};
vm.runInNewContext(transpile(collectBody + '\n' + sharedBody), context);
async function main() {
  const run = () => context.exports.createSharedCompanyProfilePptPlan({ title: '初回客先訪問用 営業資料', userPrompt: prompt });
  await assert.rejects(() => context.exports.createSharedCompanyProfilePptPlan({ title: '営業資料', userPrompt: '会社紹介してください' }), /会社名/);
  results = [{ title: '楽天グループ会社概要', url: 'https://corp.rakuten.co.jp/company/' }];
  await assert.rejects(run, /公式サイト/);
  assert.equal(plans, 0, 'unrelated search results must not reach slide planning');
  results = [{ title: '株式会社ミダックホールディングス 会社概要', url: 'https://www.midac.jp/company/' }];
  pages = { 'https://www.midac.jp/company/': '楽天グループ株式会社の事業概要' };
  await assert.rejects(run, /公式サイト/);
  assert.equal(plans, 0, 'wrong-company page must not reach slide planning');
  pages = { 'https://www.midac.jp/company/': '株式会社ミダックホールディングス 産業廃棄物一貫処理' };
  actualCompany = '楽天グループ株式会社';
  await assert.rejects(run, /一致しない/);
  assert.equal(plans, 0, 'wrong-company brief must not reach slide planning');
  actualCompany = '株式会社ミダックホールディングス';
  const plan = await run();
  assert.equal(plan.companyName, 'ミダックホールディングス');
  assert.equal(plan.slides.length, 7);
  assert.equal(plans, 1);
  const explicit = await context.exports.createSharedCompanyProfilePptPlan({ title: '営業資料', userPrompt: prompt + ' 公式URL: https://www.midac.jp/company/' });
  assert.equal(explicit.officialDomain, 'midac.jp');
  console.log('Company identity and wrong-company stop tests passed (search/LLM doubles).');
}
main().catch(e => { console.error(e); process.exitCode = 1; });
