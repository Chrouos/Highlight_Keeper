const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const source=fs.readFileSync(require.resolve('../contentScript.js'),'utf8');
const start=source.indexOf('let inlineThoughtSlot=null,inlineThoughtBlock=null;');
const end=source.indexOf('globalThis.HkDeskBridge =',start);
let slot;
const block={tagName:'P',after(node){slot=node;node.isConnected=true;}};
const mark={getBoundingClientRect:()=>({top:100,bottom:120}),closest:()=>block};
const document={querySelectorAll:()=>[mark],createElement:()=>({
  dataset:{},style:{height:''},setAttribute(){},remove(){this.isConnected=false;},
  getBoundingClientRect:()=>({left:0,top:120,width:500})
})};
const context={document,innerWidth:800,innerHeight:600,HIGHLIGHT_ATTR:'data-highlight',CSS:{escape:id=>id}};
vm.createContext(context);
vm.runInContext(source.slice(start,end)+';globalThis.reserve=reserveInlineThought;',context);
context.reserve('a');
assert.equal(slot.style.height,'64px');
context.reserve('a',250);
assert.equal(slot.style.height,'266px');
context.reserve('a');
assert.equal(slot.style.height,'266px','looking up a slot preserves the measured height');
context.reserve('a',250);
assert.equal(slot.style.height,'266px','repeated layouts keep the same height');
console.log('Inline thought slot keeps its measured height across layouts');
