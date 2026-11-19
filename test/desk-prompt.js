const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
const source=fs.readFileSync(require.resolve('../contentScript.js'),'utf8');const start=source.indexOf('globalThis.HkDeskBridge = {');assert.ok(start>=0);
const copied=[],messages=[];
const context={pageKey:'https://a.test/',collectGuidePageData:async()=>({text:'current article'}),refreshPaletteFromStorage:async()=>['#fff5b8'],collectExistingTagsByUsage:async()=>['AI'],buildReadingGuidePrompt:(data,palette,tags)=>JSON.stringify({data,palette,tags}),navigator:{clipboard:{writeText:async text=>copied.push(text)}},chrome:{runtime:{sendMessage:async message=>messages.push(message)}},t:key=>key};
vm.createContext(context);
const end = source.indexOf('\n};', start);
assert.ok(end > start);
const bridge=vm.runInContext(source.slice(start, end + 3) + ';HkDeskBridge',context);
(async()=>{
 await bridge.copyPrompt();assert.equal(copied.length,1);assert.match(copied[0],/current article/);assert.match(copied[0],/AI/);assert.equal(messages[0].type,'HK_DESK_PROMPT_SOURCE');
 context.collectGuidePageData=async()=>{context.pageKey='https://b.test/';return{text:'other page'}};
 await assert.rejects(bridge.copyPrompt(),/sourceMismatch/);assert.equal(copied.length,1,'navigation cannot copy a prompt for the wrong page');

 context.aiSettings={provider:'gemini-web'};
 let applied;
 context.applyReadingGuideResponse=async(...args)=>{applied=args;return{success:true,context:'summary'}};
 await assert.rejects(bridge.applyReply('https://a.test/','reply'),/sourceMismatch/);
 assert.equal(applied,undefined,'wrong-page reply never enters the application pipeline');
 const result=await bridge.applyReply('https://b.test/','reply');assert.equal(result.success,true);assert.equal(applied[1],'gemini-web');assert.equal(applied[3],'https://b.test/');
 console.log('Direct desk prompt uses current article and binds its source');
})().catch(e=>{console.error(e);process.exitCode=1});
