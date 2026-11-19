const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
require('../parsers.js');
const source=fs.readFileSync(require.resolve('../auto-sync.js'),'utf8');
const a=source.indexOf('  const parseRemotePages ='),b=source.indexOf('  const canonicalJson',a);
const parse=vm.runInNewContext(source.slice(a,b)+';parseRemotePages',{root:globalThis,URL});
const page={url:'https://example.com/a',entries:[{id:'a',text:'saved'}]};
assert.equal(parse(JSON.stringify({pages:[page]})).length,1);
assert.equal(parse(JSON.stringify([{url:page.url,entries:[{range:{text:'legacy text'}}]}])).length,1);
for(const entries of ['broken',[null],[{}],[{text:{bad:true}}],[{text:'valid',note:{bad:true}}]]){
 assert.throws(()=>parse(JSON.stringify([{...page,entries}])),'invalid highlight aborts complete backup');
}
for(const cards of [[null],[{id:'a',type:'unknown'}],[{id:'a',type:'quote'}]]){
 assert.throws(()=>parse(JSON.stringify([{url:page.url,readingDesk:{cards}}])),'invalid desk card aborts complete backup');
}
assert.throws(()=>parse(JSON.stringify([page,page])),'duplicate pages are not silently discarded');
console.log('Remote schema validation rejects damaged records and preserves legacy ranges');
