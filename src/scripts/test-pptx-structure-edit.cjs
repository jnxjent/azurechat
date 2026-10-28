const fs=require('node:fs'),ts=require('typescript'),assert=require('node:assert/strict'),path=require('node:path');
const m={exports:{}};
function pure(name){const mod={exports:{}};new Function('require','module','exports',ts.transpileModule(fs.readFileSync(path.join(__dirname,'../features/pptx/'+name+'.ts'),'utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2020}}).outputText)(id=>pure(id.replace('./','')),mod,mod.exports);return mod.exports;}
Object.assign(m.exports,pure('structure-edit-intent'));
const {detectPptxStructureEdit:detect,isPptxStructureEditRequest:routed,verifyPptxStructureContent:verify,sanitizePptxStructureOps:sanitize}=m.exports;

// Slide insertion
let r=detect('①のPPTXに、②の情報のスライドを一枚追加して');
assert.equal(r.kind,'add_slide');assert.equal(r.count,1);assert.equal(r.position,undefined);assert.equal(r.blank,false);
r=detect('P3の後に業界大手5社のスライドを2枚追加して');assert.equal(r.kind,'add_slide');assert.equal(r.position,3);assert.equal(r.count,2);
r=detect('スライドを5枚追加して');assert.equal(r.count,3);assert.equal(r.requestedCount,5);
assert.equal(detect('最後から2枚目にまとめのスライドを追加して').position,-2);
assert.equal(detect('3枚目として会社概要のスライドを挿入して').position,2);
assert.equal(detect('空白スライドを追加して').blank,true);
assert.equal(detect('各社の特色をまとめたスライドを追加して').kind,'add_slide','特色 is not a color request');
assert.equal(detect('上の内容を箇条書きでスライドを1枚追加して').kind,'add_slide','a new slide may use bullets');
r=detect('スライドを追加して、レイアウトも最適化して');assert.equal(r.kind,'add_slide');assert.deepEqual(r.unsupported,['レイアウト最適化・shape移動']);

// Empty slide body
r=detect('P5（空白）にまとめを入れて');assert.equal(r.kind,'fill_slide');assert.equal(r.slideIndex,4);

// Font size
r=detect('タイトルの文字を大きくして');assert.deepEqual([r.kind,r.scope,r.mode,r.value],['font_size','title','relative',2]);
r=detect('P2の本文を14ptにして');assert.deepEqual([r.scope,r.mode,r.value,r.slides],['body','absolute',14,[1]]);
assert.equal(detect('表の文字をかなり小さくして').value,-4);

// Existing routes keep handling these (no behavior change)
for(const text of ['P2の色を緑にして','ロゴを右上に追加して','P2に箇条書きを追加して','P2のスライドに箇条書きを追加','フォントをメイリオに変更して','「旧社名」を「新社名」に変更して','P3の内容を詳しくして','画像をスライドに追加して'])
  assert.equal(detect(text),null,text);

// Chat routing: requires a PPT word
assert.equal(routed('①のPPTXに②の情報のスライドを一枚追加して'),true);
assert.equal(routed('Wikiのページを追加して'),false);
assert.equal(routed('メモに追加して'),false);
assert.equal(routed('タイトルの文字を大きくして'),false,'no PPT word');
assert.equal(routed('資料のタイトルの文字を大きくして'),true);

// Grounding against the chat answer
const source='上位5社: 大栄環境株式会社 売上730億円 / TREホールディングス 900億円 https://example.com/ir';
const checked=verify([{title:'業界大手',kind:'table',table:{columns:['企業','売上'],rows:[['大栄環境','730億円'],['TREホールディングス','900億円'],['架空商事','500億円'],['大栄環境','999億円']]},sources:[{label:'IR',url:'https://example.com/ir'},{label:'x',url:'https://fake.example/x'}]}],source);
assert.equal(checked.slides[0].table.rows.length,2,'invented company and number are dropped');
assert.deepEqual(checked.slides[0].sources.map(s=>s.url),['https://example.com/ir']);
assert.match(checked.warnings[0],/2件/);
assert.equal(verify([{title:'x',kind:'bullets',bullets:['売上は123億円']}],source).slides.length,0,'nothing grounded → no slide');

// Whitelist for the Python plan
const ops=sanitize([{type:'add_slide',position:3,title:'T',kind:'table',table:{columns:['a','b','c','d','e','f'],rows:Array.from({length:9},()=>['1','2'])}},{type:'drop_all'},{type:'font_size',scope:'all',mode:'relative',value:0},{type:'fill_slide',slideIndex:-1}]);
assert.equal(ops.length,1);assert.equal(ops[0].table.columns.length,5);assert.equal(ops[0].table.rows.length,6);
console.log('PASS pptx structure edit intent, routing, grounding and op whitelist');

let content=verify([{title:'結論。',kind:'stats',keyMessage:'売上999億円',stats:[{label:'A社',value:'100',unit:'億円',period:'2025年度'},{label:'B社',value:'80',unit:'億円',period:'2025年度'}]}],'A社 2025年度 100億円\nB社 2025年度 80億円');
assert.equal(content.slides[0].keyMessage,undefined);assert.equal(content.slides[0].title,'結論');assert.equal(content.slides[0].kind,'stats');
assert.equal(verify([{title:'結論',kind:'stats',stats:[{label:'A社',value:'100'},{label:'架空社',value:'999'}]}],'A社 100').slides.length,0);
content=sanitize([{type:'add_slide',title:'比較。',kind:'two_column',bullets:['A社は回収'],table:{columns:['企業','特徴'],rows:[['A社','回収']]},unknown:'discard'}]);
assert.equal(content[0].kind,'two_column');assert.equal(content[0].unknown,undefined);assert.equal(content[0].title,'比較');
const safeStats=sanitize([{type:'add_slide',title:'数値',kind:'stats',stats:[{label:'A',value:'10',extra:99},{label:'B',value:'20'}]}]);assert.equal(safeStats[0].stats[0].extra,undefined);
console.log('PASS key message evidence, stats/two_column whitelist and title punctuation');
