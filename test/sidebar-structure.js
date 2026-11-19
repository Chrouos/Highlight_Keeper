const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const source=fs.readFileSync(require.resolve('../manager.js'),'utf8');
const setup=source.slice(0,source.indexOf('const PAGE_META_KEY'));
let flattened=false,removed=false;
const summary={remove(){removed=true;backup.children.shift();}},child={};
const backup={children:[summary,child],open:false,querySelector:()=>summary,replaceWith(...children){flattened=children.length===1&&children[0]===child;}};
vm.runInNewContext(setup,{window:{location:{search:'?embed=sync'}},URLSearchParams,document:{documentElement:{classList:{add(){}}},querySelector:()=>backup}});
assert(flattened&&removed,'embedded backup has no second disclosure');
console.log('Sidebar embedded backup opens without nested disclosure');
