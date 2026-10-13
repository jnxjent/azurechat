const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const ts = require('typescript');

// Scripted service doubles. Each Vision call takes the next queued item (an Error is thrown).
const visionQueue = [];
const visionCalls = [];
let diLines = [];
const reply = (content, finish_reason = 'stop') => ({ choices: [{ finish_reason, message: { content: typeof content === 'string' ? content : JSON.stringify(content) } }] });
const openaiDouble = { OpenAIVisionInstance: () => ({ chat: { completions: { create: async params => {
  visionCalls.push(params);
  const next = visionQueue.shift();
  if (!next) throw new Error('unexpected Vision call');
  if (next instanceof Error) throw next;
  return next;
} } } }) };
// Document Intelligence page in pixels of the (enlarged) image it receives.
const diDouble = { DocumentIntelligenceInstance: () => ({ beginAnalyzeDocument: async () => ({ pollUntilDone: async () => {
  let content = '';
  const lines = diLines.map(line => {
    const offset = content.length; content += line.text + '\n';
    const words = line.words.map(([text, x0, x1]) => ({ content: text, span: { offset: offset + line.text.indexOf(text), length: text.length },
      polygon: [{ x: x0, y: line.y0 }, { x: x1, y: line.y0 }, { x: x1, y: line.y1 }, { x: x0, y: line.y1 }] }));
    return { content: line.text, spans: [{ offset, length: line.text.length }], words: () => words[Symbol.iterator](),
      polygon: [{ x: words[0].polygon[0].x, y: line.y0 }, { x: words[words.length - 1].polygon[1].x, y: line.y0 }, { x: words[words.length - 1].polygon[1].x, y: line.y1 }, { x: words[0].polygon[0].x, y: line.y1 }] };
  });
  return { pages: [{ width: 1600, height: 1200, lines }] };
} }) }) };

const cache = new Map();
function load(relative) {
  const filename = path.resolve(relative);
  if (cache.has(filename)) return cache.get(filename);
  const mod = new Module(filename); mod.filename = filename;
  mod.paths = Module._nodeModulePaths(path.dirname(filename));
  const requireOriginal = mod.require.bind(mod);
  mod.require = name => {
    if (name === 'server-only') return {};
    if (name === '@/features/common/services/openai') return openaiDouble;
    if (name === '@/features/common/services/document-intelligence') return diDouble;
    if (name.startsWith('@/')) return load(name.slice(2) + '.ts');
    return requireOriginal(name);
  };
  mod._compile(ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true } }).outputText, filename);
  cache.set(filename, mod.exports); return mod.exports;
}
const intent = load('features/pptx/output-intent.ts');
assert(intent.isImagePptConversionRequest('添付PNGを画像はそのままで文字を編集可能なPPTに変換して'));
assert(!intent.isImagePptConversionRequest('添付ロゴ画像を参考に会社紹介のPPTを作成して'));
assert(!intent.isNewPptCreationRequest('この画像を編集可能なPPTに変換して'));
const prompt = '当社㈱ミダックホールディングスは産業廃棄物一貫処理会社です。初回客先訪問用の営業資料を8枚のスライドで作成してください（PPTで出力してください）。色は、さわやかなGreen系で。内容はHPから入手してください。添付会社Logoを各スライドの右上には①してください。';
assert(intent.isNewPptCreationRequest(prompt));
assert(!intent.isNewPptCreationRequest('既存のPPTを修正してください'));
assert(!intent.isExplicitPptOutputRequest('ロゴを使って画像を作成して'));
assert.equal(intent.pdfPptConversionMode('PDFをPPTにして'), 'faithful');
assert.equal(intent.pdfPptConversionMode('PDFを編集可能なPPTに再構成して'), 'redesign');
// "Editable text" alone must keep source images (faithful), not switch to the redesign path.
assert.equal(intent.pdfPptConversionMode('画像はそのままで、文字だけ編集可能なPPTに変換して'), 'faithful');

async function main() {
  const { createCanvas } = require('@napi-rs/canvas');
  const sharp = require('sharp');
  const JSZip = require('jszip');
  const canvas = createCanvas(400, 300), ctx = canvas.getContext('2d');
  ctx.fillStyle = 'white'; ctx.fillRect(0, 0, 400, 300);
  ctx.strokeStyle = '#ddd';
  for (let y = 10; y < 250; y += 140) for (let x = 10; x < 350; x += 100) { ctx.fillStyle = '#eee'; ctx.fillRect(x,y,90,120); ctx.strokeRect(x, y, 90, 120); }
  const panels = load('features/pptx/pdf-page-panels.ts').findPdfPagePanels(ctx.getImageData(0, 0, 400, 300).data, 400, 300);
  assert.equal(panels.length, 8);

  const layoutModule = load('features/pptx/pdf-editable-layout.ts');
  const { normalizeEditablePdfLayout: normalize, recognizeEditablePdfLayout: recognize, isMisreadCorrection } = layoutModule;
  const { createPdfImageDeck, eraseTextRegions } = load('features/pptx/pdf-image-deck.ts');
  const { addLogoToPptx } = load('features/pptx/logo-overlay.ts');

  // Normalisation drops only the broken boxes and clamps small overflows.
  const normalized = normalize({ texts: [
    { text: '本文', x: '0.1', y: '0.9', w: '0.8', h: '0.12', color: '#123456' },
    { text: '範囲外', x: 10, y: .1, w: .2, h: .1 },
    { text: '座標なし', x: .1 },
    { text: '   ', x: .1, y: .1, w: .1, h: .1 },
  ] });
  assert.equal(normalized.layout.texts.length, 1);
  assert.equal(normalized.layout.texts[0].color, '123456');
  assert(normalized.layout.texts[0].fontSize > 0);
  assert(normalized.layout.texts[0].y + normalized.layout.texts[0].h <= 1);
  assert.equal(normalized.issues.length, 2);
  assert.deepEqual(normalize({ texts: [] }).layout.texts, [], 'a page without text is valid');
  assert.throws(() => normalize({ texts: 'x' }), /配列/);

  assert(isMisreadCorrection('従業員敗', '従業員数'));
  assert(isMisreadCorrection('収集這艘', '収集運搬'));
  assert(isMisreadCorrection('東京 はか', '東京 ほか'));
  assert(isMisreadCorrection('愛知·岐阜·三重·静岡·滋賀·東京 はか', '愛知・岐阜・三重・静岡・滋賀・東京 ほか'), 'width variants are not edits');
  assert(isMisreadCorrection('中間処理施設(破碎·選別)', '中間処理施設（破砕・選別）'));
  assert(!isMisreadCorrection('その他産業廃棄物', 'その他施設産業廃棄物'), 'additions are refused');
  assert(!isMisreadCorrection('億頼にお応えします。', 'ご要望にお応えします。'), 'rewrites are refused');
  assert(!isMisreadCorrection('愛知・岐阜・三重・静岡・滋賀・東京 はか', '愛知・岐阜・三重・静岡・滋賀・東京'), 'drops are refused');

  const page = createCanvas(400, 300), pctx = page.getContext('2d');
  pctx.fillStyle = '#ffffff'; pctx.fillRect(0, 0, 400, 300);
  pctx.fillStyle = '#2e7d32'; pctx.fillRect(40, 150, 120, 100); // a "photo" that must survive
  pctx.fillStyle = '#000000'; pctx.font = '28px sans-serif'; pctx.fillText('TITLE', 40, 60);
  const data = 'data:image/png;base64,' + page.toBuffer('image/png').toString('base64');

  // Vision fallback (no Document Intelligence): a layout whose boxes are mostly broken is retried with feedback.
  delete process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT; delete process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY;
  visionCalls.length = 0;
  visionQueue.push(
    Object.assign(new Error('rate limited'), { status: 429 }),
    reply({ texts: [{ text: 'TITLE', x: 5, y: 5, w: 1, h: 1 }, { text: 'x', x: 2, y: 2, w: 1, h: 1 }] }),
    reply({ texts: [{ text: 'TITLE', x: .1, y: .1, w: .4, h: .12, fontSize: .09 }] }),
  );
  const fallback = await recognize(data);
  assert.equal(visionCalls.length, 3, '429 retried, then invalid layout retried');
  assert.match(visionCalls[2].messages[0].content[0].text, /前回の解析の問題/);
  assert.deepEqual(fallback.texts.map(t => t.text), ['TITLE']);

  // Persistent bad answers still return the best attempt instead of failing the whole document.
  // The final empty answer must not replace the earlier attempt that found text.
  visionQueue.push(reply('not json'), reply({ texts: [{ text: 'A', x: .1, y: .1, w: .1, h: .1 }, { text: 'B', x: 3, y: 3, w: 1, h: 1 }, { text: 'C', x: .1 }] }), reply('{}'));
  assert.deepEqual((await recognize(data)).texts.map(t => t.text), ['A']);

  // Document Intelligence path: positions from OCR, Vision corrects glyphs and flags logo text.
  process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = 'https://di.test/'; process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY = 'k';
  diLines = [
    // Heading and logo on one baseline: split at the wide gap so the logo can be excluded.
    { text: '会社概要 MIDAC', y0: 100, y1: 180, words: [['会', 150, 200], ['社', 205, 255], ['概', 260, 310], ['要', 315, 365], ['MIDAC', 1100, 1500]] },
    { text: '従業員敗', y0: 300, y1: 340, words: [['従', 150, 190], ['業', 192, 232], ['員', 234, 274], ['敗', 276, 316]] },
    { text: '億頼にお応えします。', y0: 400, y1: 440, words: [['億頼にお応えします。', 150, 550]] },
  ];
  visionCalls.length = 0;
  visionQueue.push(reply({ lines: [
    { i: 0, text: '会社概要', keep: true, bold: true },
    { i: 1, text: 'MIDAC', keep: false },
    { i: 2, text: '従業員数', keep: true },
    { i: 3, text: 'ご要望にお応えします。', keep: true },
  ] }));
  const di = await recognize(data);
  assert.equal(visionCalls.length, 1);
  assert.match(visionCalls[0].messages[0].content[0].text, /"会社概要".*"MIDAC"/s, 'logo segment is sent separately');
  assert.deepEqual(di.texts.map(t => t.text), ['会社概要', '従業員数', '億頼にお応えします。']);
  assert.equal(di.texts[0].bold, true);
  assert(Math.abs(di.texts[0].x - 150 / 1600) < 1e-6 && Math.abs(di.texts[0].w - 215 / 1600) < 1e-6);

  // A widely spaced contact footer stays in one editable line even if Vision
  // labels both its company name and contact information as logo text.
  diLines = [
    { text: '（株）テストホールディングス', y0: 1080, y1: 1110, words: [['（株）テストホールディングス', 150, 900]] },
    { text: '運行管理者：テスト（内線：1234）', y0: 1120, y1: 1150, words: [['運', 150, 175], ['行', 250, 275], ['管理者：テスト（内線：1234）', 350, 1200]] },
  ];
  visionQueue.push(reply({ lines: [{ i: 0, text: diLines[0].text, keep: false }, { i: 1, text: diLines[1].text, keep: false }] }));
  assert.deepEqual((await recognize(data)).texts.map(t => t.text), diLines.map(line => line.text));
  assert(!layoutModule.isContactFooterText({ text: '写真の看板', y: .2 }, [{ text: '内線：1234', y: .95 }]));

  // Even when Vision calls a multi-colour Latin mark ordinary text, it stays in the image; plain Latin text stays editable.
  const branded = createCanvas(400, 300), bctx = branded.getContext('2d');
  bctx.fillStyle = '#ffffff'; bctx.fillRect(0, 0, 400, 300);
  bctx.font = 'bold 36px sans-serif';
  bctx.fillStyle = '#1f4fa8'; bctx.fillText('LO', 250, 60);
  bctx.fillStyle = '#f08c1e'; bctx.fillText('GO', 305, 60);
  bctx.fillStyle = '#222222'; bctx.font = '24px sans-serif'; bctx.fillText('Plain text', 40, 200);
  const brandedData = 'data:image/png;base64,' + branded.toBuffer('image/png').toString('base64');
  diLines = [
    { text: 'LOGO', y0: 4 * 28, y1: 4 * 64, words: [['LOGO', 4 * 248, 4 * 362]] },
    { text: 'Plain text', y0: 4 * 180, y1: 4 * 206, words: [['Plain', 4 * 38, 4 * 95], ['text', 4 * 100, 4 * 145]] },
  ];
  visionQueue.push(reply({ lines: [{ i: 0, text: 'LOGO', keep: true }, { i: 1, text: 'Plain text', keep: true }] }));
  assert.deepEqual((await recognize(brandedData)).texts.map(t => t.text), ['Plain text']);

  // A page without readable text keeps the source page and makes no Vision call.
  diLines = []; visionCalls.length = 0;
  const textless = await recognize(data);
  assert.deepEqual(textless.texts, []);
  assert.equal(visionCalls.length, 0);

  // A failed correction keeps the raw OCR text rather than failing.
  diLines = [{ text: 'TITLE', y0: 110, y1: 250, words: [['TITLE', 160, 700]] }];
  visionQueue.push(reply('', 'length'), reply('broken'));
  assert.deepEqual((await recognize(data)).texts.map(t => t.text), ['TITLE']);

  // Erasing removes glyphs inside the box and leaves the photo pixels untouched.
  const titleBox = { x: 36 / 400, y: 36 / 300, w: 92 / 400, h: 30 / 300, fontSize: 24 / 300 };
  const erased = await sharp(await eraseTextRegions(Buffer.from(data.split(',')[1], 'base64'), [titleBox])).raw().toBuffer({ resolveWithObject: true });
  const px = (x, y) => Array.from(erased.data.subarray((y * 400 + x) * erased.info.channels, (y * 400 + x) * erased.info.channels + 3));
  for (let y = 40; y < 62; y += 3) for (let x = 40; x < 125; x += 3) assert(Math.min(...px(x, y)) > 235, `glyph pixel left at ${x},${y}`);
  assert.deepEqual(px(100, 200), [0x2e, 0x7d, 0x32], 'photo pixels are preserved');

  const editableLayout = { backgroundColor: 'FFFFFF', texts: [{ ...titleBox, text: '編集できる会社案内', color: '000000', bold: true, align: 'left' }] };
  const editableDeck = await createPdfImageDeck([{ sourceImageDataUrl: data, sourceWidth: 400, sourceHeight: 300, editableLayout }, { sourceImageDataUrl: data, sourceWidth: 400, sourceHeight: 300, editableLayout: textless }]);
  const editableZip = await JSZip.loadAsync(editableDeck);
  const editableXml = await editableZip.file('ppt/slides/slide1.xml').async('string');
  assert(editableXml.includes('編集できる会社案内'), 'OCR text must be editable DrawingML text');
  assert.equal((editableXml.match(/<p:pic>/g) || []).length, 1, 'one source image with text erased');
  assert(editableXml.includes('<p:sp>'), 'editable text shape must exist');
  assert(!/normAutofit/.test(editableXml), 'no shrink-on-edit autofit');
  const textlessXml = await editableZip.file('ppt/slides/slide2.xml').async('string');
  assert.equal((textlessXml.match(/<p:pic>/g) || []).length, 1);
  assert(!textlessXml.includes('<a:t>'), 'textless page has no text shapes');

  const logo = createCanvas(200, 80); logo.getContext('2d').fillRect(0,0,200,80);
  const deck = await createPdfImageDeck([{sourceImageDataUrl:data,sourceWidth:400,sourceHeight:300}, {sourceImageDataUrl:data,sourceWidth:300,sourceHeight:400}]);
  const overlaid = await addLogoToPptx(deck, 'data:image/png;base64,' + logo.toBuffer('image/png').toString('base64'));
  const zip = await JSZip.loadAsync(overlaid);
  for (let i=1; i<=2; i++) {
    const xml = await zip.file(`ppt/slides/slide${i}.xml`).async('string');
    assert(xml.includes('Company logo')); assert.equal((xml.match(/<p:pic>/g)||[]).length,2);
    assert((await zip.file(`ppt/slides/_rels/slide${i}.xml.rels`).async('string')).includes('azurechat-company-logo.png'));
  }
  assert.equal(visionQueue.length, 0, 'every scripted Vision response was consumed');
  console.log('PPT output regression checks passed (service doubles; run scripts/verify-pdf-editable-ppt-real.cjs for the real PDF + Vision check)');
}
const routing = load('features/pptx/output-intent.ts');
const salesPrompt = '当社㈱ミダックホールディングスは産業廃棄物一貫処理会社です。初回客先訪問用の営業資料を8枚のスライドで作成してください（PPTで出力してください）。色は、さわやかなGreen系で。内容はHPから入手してください。添付会社Logoを各スライドの右上に配置してください。';
assert(routing.isNewPptCreationRequest(salesPrompt));
assert(routing.isCompanyProfileContent(salesPrompt, 7));
assert(routing.requestsOfficialWebsiteContent('内容はHPから入手してください'));
assert(!routing.isCompanyProfileContent('AzureChatの機能紹介資料を作成してください', 7));
const rendering = load('features/teams/teams-ppt-plan-service.ts').buildTeamsPptRenderingOptions(salesPrompt, '営業資料');
assert.equal(rendering.palette, 'forest_amber');
assert(!rendering.designInstruction.includes('ネイビー'));
assert.deepEqual(routing.removePptProductionNotes([{title:'事業内容',bullets:['一貫処理会社です。','右上に添付ロゴ「midac_logo_16.png」を配置']}]), [{title:'事業内容',bullets:['一貫処理会社です。']}]);
for (const counts of [{imageCount:1,documentCount:0},{imageCount:0,documentCount:1},{imageCount:1,documentCount:1}]) {
  assert.equal(routing.resolveOfficeChatRoute({explicitPptRequest:routing.isExplicitPptOutputRequest(salesPrompt),requiredTools:false,...counts}), 'extensions');
}
assert.equal(routing.resolveOfficeChatRoute({explicitPptRequest:false,requiredTools:false,imageCount:1,documentCount:0}), 'multimodal');
main().catch(e => {console.error(e);process.exitCode=1;});
