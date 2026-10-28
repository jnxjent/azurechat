const fs=require('node:fs'),path=require('node:path'),ts=require('typescript'),assert=require('node:assert/strict'),{EventEmitter}=require('node:events');
const saved=[];let release;const gate=new Promise(resolve=>release=resolve);
function load(file){const m={exports:{}};new Function('require','module','exports',ts.transpileModule(fs.readFileSync(path.join(__dirname,'..',file),'utf8'),{compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2022}}).outputText)(id=>{
 if(id==='@/features/theme/theme-config')return {AI_NAME:'Assistant'};
 if(id==='../chat-message-service')return {CreateChatMessage:async item=>{await gate;saved.push(item);}};
 if(id==='@/features/ui/markdown/citation-markup')return load('features/ui/markdown/citation-markup.ts');
 if(id==='./generated-file-result')return load('features/chat-page/chat-services/chat-api/generated-file-result.ts');
 throw Error(id);
},m,m.exports);return m.exports;}
(async()=>{
 const {OpenAIStream}=load('features/chat-page/chat-services/chat-api/open-ai-stream.ts');
 const runner=new EventEmitter();runner.abort=()=>runner.emit('abort');
 const reader=OpenAIStream({runner,chatThread:{id:'test'}}).getReader();
 const tool={integration:'pptx_structure_edit',downloadUrl:'https://example.invalid/edited.pptx',message:'スライドを追加しました（P4）。\n⚠ 長い文章を省略しました。'};
 runner.emit('functionCallResult',JSON.stringify(tool));runner.emit('end');
 let visible='';for(let i=0;i<4;i++){const chunk=await reader.read();visible+=new TextDecoder().decode(chunk.value);}
 assert.ok(visible.includes('event: finalContent'),'final message reaches browser before persistence finishes');
 assert.ok(visible.includes('スライドを追加しました'));assert.ok(visible.includes('長い文章を省略しました'));assert.ok(visible.includes('https://example.invalid/edited.pptx'));
 assert.equal(saved.length,0);release();let rest='';while(true){const r=await reader.read();if(r.done)break;rest+=new TextDecoder().decode(r.value);}
 assert.equal(saved.filter(r=>r.role==='assistant').length,1);assert.ok(saved.at(-1).content.includes('長い文章を省略しました'));
 assert.ok(!rest.includes('event: finalContent'),'exactly one completion event');
 const other=new EventEmitter();const response=new Response(OpenAIStream({runner:other,chatThread:{id:'other'}})).text();
 other.emit('functionCallResult',JSON.stringify({downloadUrl:'https://example.invalid/other.docx',message:'other'}));other.emit('end');
 assert.ok((await response).includes('ファイルの処理が完了しました。'),'other file fallback unchanged');
 console.log('PASS structural edit completion before persistence, warning/link preserved, one saved reply, other file route unchanged');
})().catch(e=>{console.error(e);process.exitCode=1;});
