const fs=require('node:fs'),path=require('node:path'),ts=require('typescript'),assert=require('node:assert/strict');
const root=path.join(__dirname,'..');
const source=ts.createSourceFile('extensions.ts',fs.readFileSync(path.join(root,'features/chat-page/chat-services/chat-api/chat-api-default-extensions.ts'),'utf8'),ts.ScriptTarget.Latest,true);
let handler,imageCheck;
function visit(node){
  if(ts.isFunctionDeclaration(node)&&node.name?.text==='hasExplicitPptxImageInsertRequest')imageCheck=node.getText(source);
  if(ts.isObjectLiteralExpression(node)&&node.properties.some(p=>p.name?.getText(source)==='name'&&p.initializer?.text==='edit_pptx'))
    handler=node.properties.find(p=>p.name?.getText(source)==='function').initializer.getText(source);
  ts.forEachChild(node,visit);
}
visit(source);assert(handler&&imageCheck,'Use the actual edit_pptx tool entry');
const mod={exports:{}};
function pure(name){const m={exports:{}};new Function('require','module','exports',ts.transpileModule(fs.readFileSync(path.join(root,'features/pptx/'+name+'.ts'),'utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS}}).outputText)(id=>pure(id.replace('./','')),m,m.exports);return m.exports;}
Object.assign(mod.exports,pure('structure-edit-intent'));
let reads=0,calls=0;
function entry(message){
  return new Function('props','detectPptxStructureEdit','resolvePptxEditImageSource','resolveLatestStoredImageDataUrl','resolveLatestPptxInfoFromThread','executeEditPptx',
    ts.transpileModule(imageCheck+'\nconst handler='+handler+'; return handler;',{compilerOptions:{target:ts.ScriptTarget.ES2022}}).outputText)(
      {userMessage:message,chatThread:{id:'fixture'}},mod.exports.detectPptxStructureEdit,()=>'',async()=>{reads++;return null;},
      async()=>({url:'https://fixture.test/deck.pptx'}),async args=>{calls++;return args;});
}
(async()=>{
  for(const message of ['アドバイス通り「業界上位企業」スライドを1枚追加してください','上記業界トップ５の情報を先のPPTにスライド追加して']){
    const result=await entry(message)({instruction:'画像を追加して、イラストを差し替えて'});
    assert.equal(result.instruction,message);assert.equal(result.fileUrl,'https://fixture.test/deck.pptx');assert.equal(result.error,undefined);
  }
  assert.equal(reads,0);assert.equal(calls,2);
  for(const message of ['PPTの表紙の画像を差し替えて','スライドを追加して、画像を入れて']){
    const result=await entry(message)({instruction:message});assert.match(result.error,/添付画像/);
  }
  assert.equal(reads,2);assert.equal(calls,2);
  console.log('PASS actual PPT tool entry: structural requests ignore invented image requirements; explicit image requests retain attachment checks');
})().catch(error=>{console.error(error);process.exitCode=1;});
