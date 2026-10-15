const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const ts = require('typescript');
const XLSX = require('xlsx');

const cache = new Map();
function load(relative) {
  const filename = path.resolve(relative);
  if (cache.has(filename)) return cache.get(filename);
  const mod = new Module(filename); mod.filename = filename;
  mod.paths = Module._nodeModulePaths(path.dirname(filename));
  const requireOriginal = mod.require.bind(mod);
  mod.require = name => (name.startsWith('@/') ? load(name.slice(2) + '.ts') : requireOriginal(name));
  mod._compile(ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true } }).outputText, filename);
  cache.set(filename, mod.exports); return mod.exports;
}
const intent = load('features/pptx/output-intent.ts');
const slide = load('features/pptx/excel-table-slide.ts');
const palette = load('features/pptx/palette.ts');
for (const [message, key] of [
  ['全体に赤っぽく', 'burgundy_gold'], ['バーガンディゴールドで', 'burgundy_gold'],
  ['全体的にブルーっぽく', 'blue_sky'], ['青っぽい', 'blue_sky'],
  ['ブルー×スカイブルー', 'blue_sky'], ['7でお願いします', 'blue_sky'],
  ['青緑っぽく', 'teal_coral'], ['ティール×コーラルで', 'teal_coral'],
  ['ネイビーで', 'navy_orange'], ['緑っぽく', 'forest_amber'],
]) assert.equal(palette.resolvePptxPaletteInstruction(message)?.paletteKey, key, message);
assert.equal(palette.PPTX_PALETTE_KEYS[0], 'navy_orange');
assert.equal(palette.PPTX_PALETTE_KEYS[5], 'coral_orange');
assert.match(palette.pptxPaletteListText(), /青×明るい水色/);
assert.equal(intent.isExcelTableToSlideRequest('このExcelを青色で見栄えのいいPPT1枚にして'), true);

// ── 依頼文の判定 ──
for (const message of [
  '見栄えのいいPPTにして',
  'このExcelを見栄えのいいパワポにしてください',
  '添付の表をスライド1枚にまとめて',
  '一枚のPowerPointに綺麗にして',
]) assert.equal(intent.isExcelTableToSlideRequest(message), true, message);
for (const message of [
  'このExcelを元に10枚のPPTを作成して',
  'PPTの色を変えて見栄えよくして',
  '見栄えのいいExcelにして',
  'グラフを綺麗にして',
  '添付PDFを見栄えのいいPPTにして',
  '先ほどのPPTを見栄えよくして',
  'Excelを参考に見栄えのいいPPTを作って',
]) assert.equal(intent.isExcelTableToSlideRequest(message), false, message);

// 既存の判定は変更していない（新しい判定は呼び出し側で優先させるだけ）。
assert.equal(intent.isExplicitPptOutputRequest('見栄えのいいPPTにして'), true);
assert.equal(intent.isNewPptCreationRequest('見栄えのいいPPTにして'), false);

// ── 表の読み取り（経営会議サンプル相当。項目列が空の行は続きとして連結） ──
const sample = [
  ['ミダックグループ経営会議について（体制、運営の見直し）'],
  [],
  ['項目', '現状', '案1', '案2（他社事例を参考）'],
  ['会議目的', '報告（質疑あり）', '報告（質疑あり）', '議論（意思決定の場）'],
  ['体制（出席者、人数）', '取締役、部・室長、次長', '取締役、部・室長、次長', '取締役、執行役員総勢10名程度'],
  ['', '子会社社長、総勢24名', '子会社社長、総勢24名', '（部室長は必要に応じ陪席）'],
  ['会議資料（送付）', '開催前日夜', '開催２日前', '開催２日前'],
  ['報告（説明）', '資料の大半を読み上げ', '冒頭、グループ損益及び資金繰りを事務局より報告', '冒頭、グループ損益及び資金繰りを事務局より報告'],
  ['', '（重点項目を黄色マークしているが、', '部・室、会社については、以下の内容を報告（説明）', '担当役員より以下について報告（説明）'],
  ['', '部・会社によって、バラツキあり）', '・業績（計画）の遅れやブレに対する対策（リカバリープラン）', '・セグメント業績（計画）の遅れやブレに対する対策（リカバリープラン）'],
  ['', '', '・主要なトピック（黄色マークでポイントのみ説明）', '・リソースの配分・投資の判断'],
  ['', '', 'プロジェクトの進捗状況、重要顧客の動向、重大なトラブルやク', '新規事業への参入・撤退、大規模な設備投資、M＆A、重要な採用や組織変更'],
  ['', '', 'レームの発生状況', '・中長期的な経営戦略、方針の策定'],
  ['', '', '・外部環境の変化（黄色マークでポイントのみ説明）', '来期以降の事業計画、経営理念、戦略のブラッシュアップ、全社的な重要課題の設定'],
  ['', '', '競合の動き、市場・法改正・業界トレンドなどの変化', '・全社をまたぐ重要課題の解決'],
  ['', '', '', '1部門だけでは解決できない部署間の利害調整、リスク管理（コンプライアンスやセ'],
  ['', '', '', 'キュリティ等）の協議'],
];
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([[], ...sample.map(row => ['', ...row])]), 'Sheet1');
const spec = slide.parseExcelTable(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }), 'fallback');
assert.ok(spec);
assert.equal(spec.title, 'ミダックグループ経営会議について（体制、運営の見直し）');
assert.deepEqual(spec.headers, ['項目', '現状', '案1', '案2（他社事例を参考）']);
assert.equal(spec.rows.length, 4);
assert.equal(spec.baselineColumn, 1);

const kinds = items => items.map(item => `${item.kind}${item.changed ? '*' : ''}:${item.text}`);
assert.deepEqual(kinds(slide.toCellItems(spec.rows[0][2], spec.rows[0][1])), ['lead:報告（質疑あり）']);
assert.deepEqual(kinds(slide.toCellItems(spec.rows[2][2], spec.rows[2][1])), ['lead*:開催２日前']);
assert.deepEqual(kinds(slide.toCellItems(spec.rows[1][3], spec.rows[1][1])), ['lead*:取締役、執行役員総勢10名程度', 'note:（部室長は必要に応じ陪席）']);
assert.deepEqual(kinds(slide.toCellItems(spec.rows[3][1], null)), ['lead:資料の大半を読み上げ', 'note:（重点項目を黄色マークしているが、部・会社によって、バラツキあり）']);
const report1 = slide.toCellItems(spec.rows[3][2], spec.rows[3][1]);
assert.deepEqual(report1.filter(item => item.kind === 'note').map(item => item.text), [
  'プロジェクトの進捗状況、重要顧客の動向、重大なトラブルやクレームの発生状況',
  '競合の動き、市場・法改正・業界トレンドなどの変化',
]);
assert.equal(report1.filter(item => item.kind === 'bullet').length, 3);

// ── 1枚への収まり ──
const AVAILABLE_H = 7.5 - 1.3 - 0.35;
const layout = slide.layoutExcelTableSlide(spec);
assert.ok(layout, 'sample must fit on one slide');
assert.ok(Math.abs(0.5 + layout.rowHeights.reduce((a, b) => a + b, 0) - AVAILABLE_H) < 1e-6, 'table ends at the bottom margin');
const manyRows = slide.buildSpec([['項目', '現状', '案1'], ...Array.from({ length: 15 }, (_, i) => [`項目${i}`, '短い', '短い'])], 'x');
assert.ok(manyRows);
assert.equal(slide.layoutExcelTableSlide(manyRows), null, '15 rows do not fit and must not be generated');
const longText = slide.buildSpec([['項目', '現状', '案1'], ...Array.from({ length: 4 }, (_, i) => [`項目${i}`, 'あ'.repeat(600), 'い'.repeat(600)])], 'x');
assert.equal(slide.layoutExcelTableSlide(longText), null, 'long cells do not fit and must not be generated');

// 表の形でないシート・列が多すぎる表は扱わない。
assert.equal(slide.buildSpec([['メモだけ']], 'x'), null);
assert.equal(slide.buildSpec([Array.from({ length: 8 }, (_, i) => `列${i}`), Array(8).fill('a')], 'x'), null);

module.exports = { spec, render: slide.renderExcelTableSlide };
async function testRenderedPackage() {
  const JSZip = require('jszip');
  for (const source of [spec, { ...spec, baselineColumn: null }]) {
    const zip = await JSZip.loadAsync(await slide.renderExcelTableSlide(source));
    const slideNames = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name));
    assert.equal(slideNames.length, 1);
    const xml = await zip.file(slideNames[0]).async('string');
    const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map(match => match[1]);
    assert.ok(ids.length >= 3, 'generated slide must contain a title and table');
    assert.equal(new Set(ids).size, ids.length, 'shape IDs must be unique within the generated slide');
    assert.ok(xml.includes('<a:tbl>'), 'table must remain editable');
    assert.ok(xml.includes(source.title), 'title must be preserved');
    const theme = await zip.file('ppt/theme/theme1.xml').async('string');
    for (const font of theme.matchAll(/<a:(?:majorFont|minorFont)>([\s\S]*?)<\/a:(?:majorFont|minorFont)>/g)) {
      assert.match(font[1], /<a:latin\b[^>]*\/>\s*<a:ea\b[^>]*\/>\s*<a:cs\b/, 'theme must retain required East Asian font declaration');
    }
  }
  await assert.rejects(() => slide.renderExcelTableSlide(longText));
  const JSZipPalette = require('jszip');
  const blueZip = await JSZipPalette.loadAsync(await slide.renderExcelTableSlide(spec, 'Meiryo', 'blue_sky'));
  const blueXml = await blueZip.file('ppt/slides/slide1.xml').async('string');
  assert.ok(blueXml.includes('val="2463A6"'), 'requested blue main color must reach the output');
  assert.ok(blueXml.includes('val="66B5E8"'), 'requested sky blue accent must reach the output');
}
if (require.main === module) testRenderedPackage()
  .then(() => console.log('Excel table slide tests passed (including generated PPTX shape IDs).'))
  .catch(error => { console.error(error); process.exitCode = 1; });
