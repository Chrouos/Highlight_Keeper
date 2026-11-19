const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const source=fs.readFileSync(require.resolve('../reading-annotations.js'),'utf8');
const start=source.indexOf('  const pageAreaIsClear = item => {');
const end=source.indexOf('  const render=()=>{',start);
const rect={left:620,right:710,top:178,bottom:194};
const element={closest:()=>null,getClientRects:()=>[rect]};
const label={textContent:'Contributors',parentElement:element};
let nodes=[label];
const document={
  body:{},querySelectorAll:()=>[],
  createTreeWalker:()=>({index:-1,nextNode(){this.index++;this.currentNode=nodes[this.index];return !!this.currentNode;}}),
  createRange:()=>({selectNodeContents(){},getClientRects:()=>[rect]})
};
const context={document,NodeFilter:{SHOW_TEXT:4},getComputedStyle:()=>({display:'block',visibility:'visible',opacity:'1'})};
vm.createContext(context);
vm.runInContext(source.slice(start,end)+';globalThis.isClear=pageAreaIsClear;',context);
const area={x:600,y:130,width:240,height:180};
assert.equal(context.isClear(area),false,'a short sidebar label between sample points blocks the margin');
nodes=[];
assert.equal(context.isClear(area),true,'an empty margin remains available');
console.log('Reading margin detects text throughout the proposed area');
