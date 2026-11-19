const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
class Element {constructor(tag){this.tagName=tag;this.children=[];this.style={};this.events={};this.attrs={};this.textContent='';}appendChild(child){this.children.push(child);return child;}setAttribute(k,v){this.attrs[k]=v;}addEventListener(k,fn){this.events[k]=fn;}}
const source=fs.readFileSync(require.resolve('../contentScript.js'),'utf8');const start=source.indexOf('const createDeskExtractIcon = () => {');const end=source.indexOf('const hideFloatingButton =',start);
let extracted=0;const context={floatingButton:null,FLOATING_BUTTON_ID:'toolbar',document:{createElement:tag=>new Element(tag),createElementNS:(_,tag)=>new Element(tag),body:new Element('body')},chrome:{runtime:{}},t:k=>k,console,hideFloatingButton(){},showHkToast(){},HkReadingDeskUI:{extract:async()=>{extracted++;}}};vm.createContext(context);
const toolbar=vm.runInContext(source.slice(start,end)+';ensureFloatingButton()',context);const desk=toolbar.children[2];
assert.equal(desk.attrs['aria-label'],'desk.extract');assert.equal(desk.title,'desk.extract');assert.equal(desk.textContent,'');assert.equal(desk.children[0].tagName,'svg');assert.equal(desk.children[0].attrs['aria-hidden'],'true');
let prevented=false;desk.events.mousedown({preventDefault(){prevented=true}});assert.ok(prevented,'click preserves selected article text');
(async()=>{await desk.events.click();assert.equal(extracted,1);console.log('Selection toolbar icon and extract action passed');})().catch(e=>{console.error(e);process.exitCode=1});
