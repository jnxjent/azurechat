const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const http = require('node:http');
const ts = require('typescript');
const sharp = require('sharp');
let received;
function load(relative) {
  const filename = path.resolve(relative), mod = new Module(filename);
  mod.filename = filename; mod.paths = Module._nodeModulePaths(path.dirname(filename));
  const original = mod.require.bind(mod);
  mod.require = name => {
    if (name.endsWith('/image-form-data')) return load('features/ui/chat/chat-input-area/image-form-data.ts');
    if (name.endsWith('/chat-api')) return { ChatAPIEntry: async props => { received = props.multimodalImage; return new Response('ok'); } };
    return original(name);
  };
  mod._compile(ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText, filename);
  return mod.exports;
}
async function main() {
  const localFixture = 'C:/Users/021213/Downloads/飲酒運転ポスター.png';
  const source = process.argv[2] || fs.existsSync(localFixture)
    ? fs.readFileSync(process.argv[2] || localFixture)
    : await sharp(require('node:crypto').randomBytes(800 * 800 * 3), { raw: { width: 800, height: 800, channels: 3 } }).png().toBuffer();
  const original = `data:image/png;base64,${source.toString('base64')}`;
  assert(original.length > 1048576, 'fixture must reproduce the 1 MiB limit');
  const { POST } = load('app/(authenticated)/api/chat/route.ts');
  const { prepareImageFormData } = load('features/ui/chat/chat-input-area/image-form-data.ts');
  const server = http.createServer(async (req, res) => {
    try {
      const chunks = []; for await (const chunk of req) chunks.push(chunk);
      const request = new Request('http://local.test/api/chat', { method: 'POST', headers: req.headers, body: Buffer.concat(chunks) });
      const response = await POST(request); res.writeHead(response.status); res.end(await response.text());
    } catch (error) { res.writeHead(500); res.end(String(error)); }
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  try {
    const send = async fixed => {
      const form = new FormData(); form.set('content', JSON.stringify({ id: 'local-test', message: '添付ファイルをPPTにして' }));
      form.set('image-base64', original);
      if (fixed) prepareImageFormData(form);
      const response = await fetch(`http://127.0.0.1:${server.address().port}`, { method: 'POST', body: form });
      assert.equal(response.status, 200, await response.text());
    };
    await send(false);
    assert.equal(received.length, 1048576);
    await assert.rejects(sharp(Buffer.from(received.split(',')[1], 'base64')).png().toBuffer(), /libspng read error/);
    await send(true);
    assert.equal(received, original, 'real HTTP multipart must retain every character');
    const decoded = Buffer.from(received.split(',')[1], 'base64');
    assert.deepEqual(decoded, source, 'image bytes must match original exactly');
    await sharp(decoded).png().toBuffer();
    console.log(`PASS: old route reproduced libspng error; fixed HTTP route retains ${original.length} chars / ${source.length} original bytes.`);
  } finally { await new Promise(resolve => server.close(resolve)); }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
