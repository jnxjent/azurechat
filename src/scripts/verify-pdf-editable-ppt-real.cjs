// Real end-to-end check of PDF -> editable PPT (faithful mode) with the real PDF and the
// real Document Intelligence / Azure OpenAI Vision services configured in .env.local.
// Nothing is uploaded: the PDF is read from disk and the PPTX is written to <outDir>.
//   node scripts/verify-pdf-editable-ppt-real.cjs <file.pdf> <outDir>
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const ts = require('typescript');
const sharp = require('sharp');
const JSZip = require('jszip');

for (const line of fs.readFileSync('.env.local', 'utf8').split(/\r?\n/)) {
  const m = line.match(/^([A-Z0-9_]+)=(.*)$/);
  if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
}
const cache = new Map();
function load(file) {
  const filename = path.resolve(file);
  if (cache.has(filename)) return cache.get(filename);
  const mod = new Module(filename); mod.filename = filename; mod.paths = Module._nodeModulePaths(path.dirname(filename));
  const requireOriginal = mod.require.bind(mod);
  const resolveTs = base => ['.ts', '.tsx', '/index.ts'].map(ext => base + ext).find(f => fs.existsSync(f));
  mod.require = name => {
    if (name === 'server-only') return {};
    if (name.startsWith('@/')) return load(resolveTs(name.slice(2)));
    if (name.startsWith('.') && !/\.(c?js|json)$/.test(name)) return load(resolveTs(path.resolve(path.dirname(filename), name)));
    return requireOriginal(name);
  };
  cache.set(filename, mod.exports);
  mod._compile(ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true, jsx: ts.JsxEmit.ReactJSX } }).outputText, filename);
  cache.set(filename, mod.exports); return mod.exports;
}

async function main() {
  const [pdfPath, outDir] = process.argv.slice(2);
  if (!pdfPath || !outDir) throw new Error('usage: node scripts/verify-pdf-editable-ppt-real.cjs <file.pdf> <outDir>');
  fs.mkdirSync(outDir, { recursive: true });
  const pdf = fs.readFileSync(pdfPath);
  const realFetch = global.fetch;
  global.fetch = async (url, init) => String(url).startsWith('https://local.test/') ? new Response(pdf) : realFetch(url, init);
  console.log('Vision deployment:', process.env.AZURE_OPENAI_VISION_API_DEPLOYMENT_NAME, '/ Document Intelligence:', Boolean(process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT));
  const started = Date.now();
  const result = await load('app/api/analyze-doc-vision/handler.ts').analyzeDocVision('https://local.test/source.pdf', 30, 'faithful');
  assert(result.ok, result.error);
  console.log(`analyzed ${result.slides.length} slide(s) in ${Math.round((Date.now() - started) / 1000)}s`);
  fs.writeFileSync(path.join(outDir, 'layouts.json'), JSON.stringify(result.slides.map(s => s.editableLayout), null, 1));
  const deck = await load('features/pptx/pdf-image-deck.ts').createPdfImageDeck(result.slides);
  const output = path.join(outDir, 'output.pptx');
  fs.writeFileSync(output, deck);

  const zip = await JSZip.loadAsync(deck);
  for (const [index, slide] of result.slides.entries()) {
    const n = index + 1;
    const xml = await zip.file(`ppt/slides/slide${n}.xml`).async('string');
    const texts = slide.editableLayout?.texts ?? [];
    const runs = [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'));
    assert.deepEqual(runs.join('\n'), texts.map(t => t.text).join('\n'), `slide ${n}: every recognised line is editable text`);
    assert.equal((xml.match(/<p:pic>/g) || []).length, 1, `slide ${n}: exactly one source image`);
    // Every pixel outside the (padded) text boxes must be the original pixel.
    const rel = await zip.file(`ppt/slides/_rels/slide${n}.xml.rels`).async('string');
    const media = rel.match(/Target="\.\.\/media\/([^"]+)"/)[1];
    const embedded = await sharp(await zip.file(`ppt/media/${media}`).async('nodebuffer')).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const source = await sharp(Buffer.from(slide.sourceImageDataUrl.split(',')[1], 'base64')).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const { width: W, height: H, channels: C } = source.info;
    assert.equal(embedded.info.width, W); assert.equal(embedded.info.height, H);
    const inText = (x, y) => texts.some(t => {
      const pad = Math.max(2, Math.round(Math.max(t.fontSize * H, t.h * H / 3) * .3));
      return x >= Math.floor(t.x * W) - pad && x < Math.ceil((t.x + t.w) * W) + pad && y >= Math.floor(t.y * H) - pad && y < Math.ceil((t.y + t.h) * H) + pad;
    });
    let changedOutside = 0, changedInside = 0, outside = 0;
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const p = (y * W + x) * C;
      const differs = embedded.data[p] !== source.data[p] || embedded.data[p + 1] !== source.data[p + 1] || embedded.data[p + 2] !== source.data[p + 2];
      if (inText(x, y)) { if (differs) changedInside++; } else { outside++; if (differs) changedOutside++; }
    }
    assert.equal(changedOutside, 0, `slide ${n}: pixels outside text regions must be the original image`);
    console.log(`slide ${n}: ${texts.length} editable line(s), ${(outside / (W * H) * 100).toFixed(0)}% of pixels identical to source, ${changedInside} text pixel(s) erased`);
  }
  console.log('verified', output);
}
main().catch(e => { console.error(e); process.exitCode = 1; });
