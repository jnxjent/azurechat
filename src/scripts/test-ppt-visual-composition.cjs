const fs=require('node:fs'),path=require('node:path'),ts=require('typescript'),assert=require('node:assert/strict');
function pure(name){const m={exports:{}};new Function('require','module','exports',ts.transpileModule(fs.readFileSync(path.join(__dirname,'../features/pptx/'+name+'.ts'),'utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS}}).outputText)(id=>pure(id.replace('./','')),m,m.exports);return m.exports;}
const {buildVisualComposition:plan}=pure('visual-composition');
const source=[['回収','回収する'],['処理','選別する'],['処分','埋立する']];
const before=JSON.stringify(source),p=plan('流れ',['工程','内容'],source);
assert.equal(p.kind,'process');assert.equal(JSON.stringify(source),before);assert.deepEqual(p.rows,source);
assert.equal(plan('分類',['分類','内容'],source).kind,'classification');
assert.equal(plan('二極化の比較',['項目','地域企業','大手'],[['顧客','地域','広域'],['設備','回収車','処理設備']]).kind,'comparison');
const bars=plan('企業比較',['企業','2025年度売上高（億円）'],[['A社','1,187'],['B社','802']]);
assert.equal(bars.kind,'bar');assert.deepEqual(bars.values,[1187,802]);
for(const header of ['売上高（億円）','2025年度売上高','2024年と2025年売上高（億円）'])assert.equal(plan('比較',['企業',header],[['A','10'],['B','20']]),undefined);
for(const val of ['約20','20〜30','20億円','20%','-20','1,18'])assert.equal(plan('比較',['企業','2025年度売上高（億円）'],[['A','10'],['B',val]]),undefined);
assert.equal(plan('比較',['企業','2025年度売上高（億円）'],[['A 単体','10'],['B 連結','20']]),undefined);
const {sanitizePptxStructureOps}=pure('structure-edit-intent');
const ops=sanitizePptxStructureOps([{type:'add_slide',title:'工程',kind:'table',table:{columns:['工程','内容'],rows:source},composition:{kind:'bar',values:[99]}}]);
assert.equal(ops[0].composition.kind,'process','Recompute plan from canonical table; never trust caller-supplied diagram data');
const {verifyPptxStructureContent:verify}=pure('structure-edit-intent');
const chartSlide={title:'企業比較',kind:'table',table:{columns:['企業','2025年度売上高（億円）'],rows:[['A社','1,187'],['B社','802']]}};
const verified=verify([chartSlide],'企業 | 2025年度売上高（億円）\nA社 | 1,187\nB社 | 802');
assert.equal(verified.slides[0].composition.kind,'bar');
const incomplete=verify([chartSlide],'A社の売上は1,187。B社の売上は802。');
assert.equal(incomplete.slides[0].visualLayout,'table');
assert.equal(sanitizePptxStructureOps([{type:'add_slide',...incomplete.slides[0]}])[0].composition,undefined,'API sanitation preserves table fallback');
const {drawVisualComposition}=pure('visual-composition-renderer');
for(const composition of [p,bars,plan('二極化',['項目','地域企業','大手'],[['顧客','地域','広域'],['設備','回収車','処理設備']])]){
 const texts=[],shapes=[];drawVisualComposition({addText:(t,o)=>texts.push([t,o]),addShape:(s,o)=>shapes.push([s,o])},composition,{x:.5,y:1.5,w:12,h:4.5,primary:'2C7A58',accent:'D98B24',text:'223333',muted:'556666',border:'CCDADA',canvas:'FFFFFF',fontFace:'Meiryo'});
 for(const row of composition.rows)for(const cell of row)assert.ok(texts.some(([text])=>text===cell));
 assert.ok(shapes.every(([,o])=>o.w>=0&&o.h>=0));
}
console.log('PASS visual composition: editable diagrams preserve cells; mixed/unknown chart bases rejected; generation/edit planner shared');

const {requiresCompanyCards,hasCompanyRows}=pure('visual-composition');
for(const text of ['「業界上位企業」スライドを1枚追加してください','上記業界トップ５の情報を先のPPTにスライド追加して']) assert.equal(requiresCompanyCards(text),true);
assert.equal(requiresCompanyCards('まとめを追加して'),false);
assert.equal(hasCompanyRows([{kind:'bullets',bullets:['A社、B社']}]),false);
assert.equal(hasCompanyRows([{kind:'table',table:{columns:['企業','特徴'],rows:[['A社','回収']]}}]),true);
