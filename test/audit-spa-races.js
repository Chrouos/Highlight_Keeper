const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { test } = require('node:test');
const source = fs.readFileSync(require.resolve('../contentScript.js'), 'utf8');
const sidebar = fs.readFileSync(require.resolve('../sidepanel.js'), 'utf8');
const A = 'https://example.com/a', B = 'https://example.com/b';
const deferred = () => { let resolve; const promise = new Promise(r => { resolve = r; }); return {promise, resolve}; };
const extract = (text, start, end) => { const first = text.indexOf(start), last = text.indexOf(end, first); assert.ok(first >= 0 && last > first); return text.slice(first, last); };
const run = (code, context, exported) => { vm.createContext(context); vm.runInContext(code + '\nglobalThis.subject = ' + exported, context); return context.subject; };

test('deleting notes keeps the authorized source across SPA navigation', async () => {
  const gate = deferred(), removed = [], operations = [];
  const records = Object.fromEntries(['meta','hkGeneratedNotes','map','hkPersonalNotes'].map(k => [k,{[A]:{text:'A'},[B]:{text:'B'}}]));
  const context = {pageKey:A, document:{querySelectorAll:()=>[]}, HIGHLIGHT_CLASS:'mark', unwrapHighlightElement(){}, previewData:[], previewUndo:null,
    hidePreviewConfirmBar(){}, destroyGuidedBar(){}, PAGE_META_KEY:'meta', MINDMAP_STORAGE_KEY:'map',
    highlightPanelState:{notesByPage:{[A]:{},[B]:{}}}, refreshHighlightPanelIfVisible:async()=>{}, t:x=>x,
    chrome:{storage:{local:{get:()=>gate.promise,set:async updates=>Object.assign(records,updates),remove:async key=>removed.push(key)}},
      runtime:{sendMessage:async message=>{operations.push(message);return {success:true};}}}};
  const subject = run(extract(source,'const deletePageNotes =','const applyPreviewHighlight ='),context,'deletePageNotes');
  const pending = subject(A); context.pageKey=B; gate.resolve(structuredClone(records)); await pending;
  assert.ok(records.meta[B], 'navigation must preserve B metadata');
  assert.equal(records.meta[A],undefined);
  assert.deepEqual(removed,[A]); assert.equal(operations[0].operations[0].pageKey,A);
});

test('a stale restore cannot add old article marks to the new DOM', async () => {
  const gate=deferred(); let wrapped=0;
  const context={pageKey:A,domainPolicyLoaded:Promise.resolve(),domainAllowsHighlights:()=>true, storage:{get:()=>gate.promise}, document:{querySelector:()=>null}, HIGHLIGHT_ATTR:'mark',
    resolveRangeSnapshot:()=>({range:{collapsed:false}}), wrapRangeWithHighlight:()=>{wrapped++;}, setAllMarksMetadata(){},updateHighlightEntries:async()=>{},console};
  const subject=run(extract(source,'const restoreHighlights =','// 同一篇文章的識別碼'),context,'restoreHighlights');
  const pending=subject(); context.pageKey=B; gate.resolve({[A]:[{id:'old',range:{},color:'#fff5b8'}]});await pending;
  assert.equal(wrapped,0,'old restoration must stop before wrapping ranges');
});

test('blocking a domain while highlights are loading prevents restoration', async () => {
  const gate=deferred();let wrapped=0,allowed=true;
  const context={pageKey:A,domainPolicyLoaded:Promise.resolve(),domainAllowsHighlights:()=>allowed,storage:{get:()=>gate.promise},document:{querySelector:()=>null},HIGHLIGHT_ATTR:'mark',resolveRangeSnapshot:()=>({range:{collapsed:false}}),wrapRangeWithHighlight:()=>{wrapped++;},setAllMarksMetadata(){},updateHighlightEntries:async()=>{},console};
  const subject=run(extract(source,'const restoreHighlights =','// 同一篇文章的識別碼'),context,'restoreHighlights');
  const pending=subject();await Promise.resolve();allowed=false;gate.resolve({[A]:[{id:'old',range:{},color:'#fff5b8'}]});await pending;
  assert.equal(wrapped,0,'the policy is checked again after storage returns');
});

test('legacy migration preserves its initial destination after navigation',async()=>{
  const gate=deferred(), legacy=A+'?utm_source=old';
  const records={[legacy]:[{id:'old',url:legacy}], [B]:[{id:'new',url:B}], meta:{[legacy]:{title:'A'},[B]:{title:'B'}}};
  let reads=0;
  const context={pageKey:A,URL,console, PAGE_META_KEY:'meta',MINDMAP_STORAGE_KEY:'map', normalizePageKey:url=>url.split('?')[0],
    storage:{get:async key=>{if(++reads===1)return gate.promise;return structuredClone(records);},set:async updates=>Object.assign(records,updates), remove:async keys=>keys.forEach(key=>delete records[key])}};
  const subject=run(extract(source,'const articleIdentity =','// 重整後找不到對應文字'),context,'migrateLegacyPageData');
  const pending=subject();context.pageKey=B;gate.resolve({});await pending;
  assert.equal(records[B].length,1,'B highlights must remain untouched');
  assert.equal(records[A]?.[0].id,'old');assert.equal(records[A][0].url,A);
  assert.equal(records.meta[B].title,'B');assert.equal(records.meta[A].title,'A');
});

test('sidebar newest invocation wins when tab queries resolve out of order',async()=>{
  const old=deferred();let queries=0;
  const state={pageKey:'',page:null,loadToken:0,tabId:null};
  const context={state,isAIWebsite:()=>false,flushNote:async()=>{},setNoteStatus(){},setStatus(){},render(){},hasGuide:()=>false,t:x=>x,
    els:{response:{value:''},personalNote:{value:''},noPage:{},pageTags:{},aiControls:{}},
    GENERATED_NOTES_KEY:'generated',PERSONAL_NOTES_KEY:'personal',MINDMAP_KEY:'map',PAGE_META_KEY:'meta',console,
    chrome:{tabs:{query:()=>++queries===1?old.promise:Promise.resolve([{id:2,url:B}])},storage:{local:{get:async()=>({})}}},
    sendToTab:async(id,message)=>message.type==='GET_PAGE_HIGHLIGHTS'?{success:true,data:{url:id===1?A:B}}:{success:true}};
  const subject=run(extract(sidebar,'const loadActiveTab =','// 標註變動時只更新清單'),context,'loadActiveTab');
  const first=subject();await subject();assert.equal(state.pageKey,B);
  old.resolve([{id:1,url:A}]);await first;
  assert.equal(state.pageKey,B,'late earlier query cannot replace the latest active article');
});

test('stale restore attempts cannot report orphan state for a newer page',async()=>{
  const gate=deferred();let notices=0,metaWrites=0;
  const context={pageKey:A,document:{title:'Article A'},migrateLegacyPageData:async()=>{}, restoreHighlights:()=>gate.promise,
    checkPendingFocusHighlight(){},orphanHighlightCount:0,updateOrphanNotice(){},
    ensurePageMetaTitle:async()=>{metaWrites++;},HIGHLIGHT_RETRY_DELAYS:[],
    notifyOrphanHighlights(){notices++;},window:{setTimeout(){throw new Error('unexpected retry');}}};
  const subject=run(extract(source,'const runRestoreAttempt =','const applyHighlight ='),context,'runRestoreAttempt');
  const pending=subject(1);context.pageKey=B;context.document.title='Article B';gate.resolve({total:1,visible:0});await pending;
  assert.equal(metaWrites,0,'old restore cannot write the newer page title');
  assert.equal(notices,0,'old restore cannot show an orphan warning on the newer page');
});

test('deleting a guide item cannot write old highlights into the new page',async()=>{
 const gate=deferred(),writes=[],records={hkGeneratedNotes:{[A]:{guideTerms:[{text:'term'}]},[B]:{guideTerms:[{text:'term'}]}}};
 const context={pageKey:A,HkParsers:{looseKey:x=>x},HIGHLIGHT_ATTR:'mark',document:{querySelectorAll:()=>[]},unwrapHighlightElement(){},
  getStoredHighlights:()=>gate.promise,setStoredHighlights:async(entries,key)=>writes.push({entries,key:key||context.pageKey}),
  previewData:[],previewUndo:null,showPreviewConfirmBar(){},hidePreviewConfirmBar(){},highlightPanelState:{notesByPage:{}},
  refreshHighlightPanelIfVisible:async()=>{},chrome:{storage:{local:{get:async()=>structuredClone(records),set:async v=>Object.assign(records,v)}}}};
 const subject=run(extract(source,'const deleteGuideItem =','// 刪除這篇文章的全部筆記'),context,'deleteGuideItem');
 const pending=subject({text:'term',kind:'term'});context.pageKey=B;gate.resolve([{id:'one',text:'term',kind:'term'}]);await pending;
 assert.equal(writes[0].key,A);assert.equal(records.hkGeneratedNotes[B].guideTerms.length,1);
 assert.equal(records.hkGeneratedNotes[A].guideTerms.length,0);
});

test('delete confirmation does not carry over to a different article',async()=>{
 let sent=0;
 const context={state:{pageKey:A,tabId:1},els:{deletePageNotes:{classList:{add(){},remove(){}},textContent:''}},t:x=>x,
 window:{setTimeout:()=>1,clearTimeout(){}},sendToTab:async()=>{sent++;return{success:true}},loadActiveTab:async()=>{},setStatus(){},console};
 const subject=run(extract(sidebar,'let deleteArmTimer =','/* ── 分享與匯出'),context,'deletePageNotes');
 await subject();context.state.pageKey=B;context.state.tabId=2;await subject();
 assert.equal(sent,0,'a different article needs its own confirmation');
});
