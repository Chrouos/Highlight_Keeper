const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
const source=fs.readFileSync(require.resolve('../contentScript.js'),'utf8');const start=source.indexOf('const handleHighlightClick =');const end=source.indexOf('chrome.runtime.onMessage.addListener',start);
class Element{};const mark={dataset:{hkNote:'Saved thought'},getAttribute:()=> 'one'};const target=new Element();target.closest=()=>mark;let editors=0,previews=0,located=0;
const context={Element,HIGHLIGHT_CLASS:'hk-highlight',HIGHLIGHT_ATTR:'data-highlight-id',findTranslatedHighlightAtPoint:()=>null,hideFloatingButton(){},hideNoteTooltip(){},openHighlightMenu:async()=>{editors++},showNoteTooltip:()=>{previews++},HkReadingDeskUI:{locateExisting:()=>{located++;return true}},t:k=>k,console};
vm.createContext(context);const click=vm.runInContext(source.slice(start,end)+';handleHighlightClick',context);const event={target,preventDefault(){},stopPropagation(){}};
click(event);assert.equal(editors,1,'click opens the editor even when a card exists');assert.equal(located,0);assert.equal(previews,0);
context.HkReadingDeskUI.locateExisting=()=>false;click(event);assert.equal(previews,0);assert.equal(editors,2,'click without a card also opens editor');
console.log('Highlight clicks locate existing thoughts or show a read-only preview');

const hoverStart=source.indexOf('document.addEventListener("mouseover",');
const hoverEnd=source.indexOf('document.addEventListener("mouseout",',hoverStart);
let hover;const hoverContext={document:{addEventListener:(_type,fn)=>{hover=fn}},HIGHLIGHT_CLASS:'hk-highlight',HIGHLIGHT_ATTR:'data-highlight-id',HkPageAnnotations:{isShown:()=>true},clearTimeout(){},tooltipHideTimer:0,showNoteTooltip:()=>{previews++}};
vm.runInNewContext(source.slice(hoverStart,hoverEnd),hoverContext);
hover({target:{closest:()=>mark}});
assert.equal(previews,1,'hover still shows the full thought when a margin annotation is visible');
