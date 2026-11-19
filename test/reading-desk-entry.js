const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
const source=fs.readFileSync(require.resolve('../background.js'),'utf8');
const start=source.indexOf('const openReadingDesk = async');const end=source.indexOf('chrome.action.onClicked',start);
const created=[];let opened=0;
const context={chrome:{runtime:{getURL:path=>'chrome-extension://test/'+path},tabs:{sendMessage:async()=>{throw new Error('no content script')},create:async item=>created.push(item)}},openSidePanel:()=>{opened++;throw new Error('gesture expired')},console};
vm.createContext(context);vm.runInContext(source.slice(start,end)+';globalThis.openDesk=openReadingDesk;',context);
(async()=>{await context.openDesk({id:1,url:'https://example.com'});assert.equal(opened,0,'async fallback cannot depend on sidePanel user gesture');assert.equal(created[0].url,'chrome-extension://test/manager.html');console.log('Reading desk toolbar fallback passed');})().catch(e=>{console.error(e);process.exitCode=1});
