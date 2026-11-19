const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
const source=fs.readFileSync(require.resolve('../contentScript.js'),'utf8');const article='https://a.test/',other='https://b.test/';
const extract=(name,next)=>source.slice(source.indexOf('const '+name+' ='),source.indexOf(next,source.indexOf('const '+name+' =')));
(async()=>{
 let saved;const context={pageKey:article,storage:{set:async update=>saved=update},MINDMAP_STORAGE_KEY:'hkMindmaps',document:{title:'Article A'},getStoredMindmaps:async()=>{context.pageKey=other;return{[other]:{outline:'keep B'}}}};
 vm.createContext(context);const save=vm.runInContext(extract('saveMindmap','const normalizeAutoHighlightItems')+';saveMindmap',context);
 await save('A outline','A title',article);assert.equal(saved.hkMindmaps[article].outline,'A outline');assert.equal(saved.hkMindmaps[other].outline,'keep B');
 context.pageKey=article;context.HkParsers={parsePageTags:()=>['AI']};context.getPageTags=async()=>{context.pageKey=other;return['existing']};let tagTarget;
 context.setPageTags=async(key,tags)=>{tagTarget=key};context.highlightPanelState={pageMeta:{}};
 const applyTags=vm.runInContext(extract('applyAiPageTags','const applyNotePayload')+';applyAiPageTags',context);
 await applyTags('AI',article);assert.equal(tagTarget,article);assert.equal(context.highlightPanelState.pageMeta[other],undefined);
 console.log('AI mindmap and tags stay bound to the source during navigation');
})().catch(error=>{console.error(error);process.exitCode=1});
