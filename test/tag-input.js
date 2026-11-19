const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
for(const [file,name,next] of [['manager.js','normalizeTagsInput','// ── 標籤整理'],['contentScript.js','parseTags','const isValidPageKey']]){
 const source=fs.readFileSync(require.resolve('../'+file),'utf8');const start=source.indexOf('const '+name+' =');const end=source.indexOf(next,start);
 const parse=vm.runInNewContext(source.slice(start,end)+';'+name);
 assert.deepEqual(Array.from(parse('systems design，人工智慧、open source\nAI, AI')),['systems design','人工智慧','open source','AI'],file+' preserves tag names and accepts natural separators');
}
console.log('Tag input consistency passed');
