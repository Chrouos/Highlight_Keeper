const assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs');
let paths=[];
const sandbox={TypeError,AbortController,TextEncoder,btoa,setTimeout:(fn,ms)=>setTimeout(fn,ms===15000?100:0),clearTimeout,
 fetch:async(url,init)=>{
 const path=new URL(url).pathname.replace('/repos/o/r','');paths.push(init.method+' '+path);
 if(path.startsWith('/git/commits/'))throw new TypeError('Failed to fetch');
 const payload=path.startsWith('/commits/')?{commit:{tree:{sha:'base-tree'}}}:path==='/git/blobs'?{sha:'blob'}:path==='/git/trees'?{sha:'tree'}:path==='/git/commits'?{sha:'new-commit'}:{};
 return{ok:true,json:async()=>payload};
 }};
vm.runInNewContext(fs.readFileSync(require.resolve('../share-link.js'),'utf8'),sandbox);
(async()=>{
 const sha=await sandbox.HkShareLink.commitFilesToGithub({repo:'o/r',token:'secret',branch:'main'},[{path:'backup.json',content:'[]'}],'backup',{parentSha:'old'});
 assert.equal(sha,'new-commit');assert(paths.includes('GET /commits/old'));
 assert.equal(paths.filter(p=>p==='POST /git/commits').length,1);
 console.log('Commit read fallback retains original parent and does not repeat writes');
})().catch(e=>{console.error(e);process.exitCode=1});
