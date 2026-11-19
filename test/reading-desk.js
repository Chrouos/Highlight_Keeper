const assert = require('node:assert/strict');
require('../reading-desk-model.js');
require('../parsers.js');
const M = globalThis.HkReadingDesk;
const initial = M.normalize({ cards: [{ id: 'quote:a', type: 'quote', highlightId: 'a', x: -20, y: Infinity }, { id: 'note:b', type: 'note', text: '心得', x: 35, y: 60 }] });
assert.equal(initial.cards[0].x, 0);
assert.equal(initial.cards[0].y, 70);
assert.equal(M.normalize({cards:[{id:'x',type:'quote'}]}).cards.length, 0);
const patched = M.patch(initial, { type: 'upsert', card: { id: 'note:b', type: 'note', text: '新心得', x: 35, y: 60 } });
assert.equal(patched.cards[1].text, '新心得');
assert.equal(initial.cards[1].text, '心得');
assert.equal(M.patch(patched, { type: 'remove', id: 'quote:a' }).cards.length, 1);
assert.deepEqual(M.clamp(1000, -9, 280, 200, 400, 300), {x:120,y:0});
const url='https://example.com/article';
const pages=HkParsers.collectBackupPages({hkReadingDesks:{[url]:patched}});
assert.equal(pages.length,1);
assert.equal(pages[0].readingDesk.cards[1].text,'新心得');
const result=HkParsers.pagesToStorageUpdates(pages,{});
assert.equal(result.updates.hkReadingDesks[url].cards[1].text,'新心得');
console.log('Reading desk model and backup tests passed');
require('../reading-desk-store.js');
(async()=>{
 let data={};
 const store=HkReadingDeskStore.createStore({get:async()=>structuredClone(data),set:async updates=>{await new Promise(r=>setImmediate(r));Object.assign(data,structuredClone(updates));}});
 await Promise.all([store.patch('https://a.test/',{type:'upsert',card:{id:'one',type:'note',text:'A'}}),store.patch('https://b.test/',{type:'upsert',card:{id:'two',type:'note',text:'B'}})]);
 assert.equal(data.hkReadingDesks['https://a.test/'].cards[0].text,'A');
 assert.equal(data.hkReadingDesks['https://b.test/'].cards[0].text,'B');
 await assert.rejects(store.patch('javascript:alert(1)',{type:'remove',id:'one'}));
 console.log('Reading desk concurrent persistence tests passed');
})().catch(e=>{console.error(e);process.exitCode=1;});
(async()=>{
 const data={};const store=HkReadingDeskStore.createStore({get:async()=>structuredClone(data),set:async u=>Object.assign(data,structuredClone(u))});
 assert.equal(typeof store.exclusive,'function','sync and desk writes share one queue');
 await Promise.all([
   store.exclusive(async()=>{const old=structuredClone(data);await new Promise(r=>setImmediate(r));await Object.assign(data,{hkReadingDesks:{...(old.hkReadingDesks||{}),'https://sync.test/':M.normalize({cards:[{id:'sync',type:'note',text:'remote'}]})}});}),
   store.patch('https://desk.test/',{type:'upsert',card:{id:'draft',type:'note',text:'new thought'}}),
 ]);
 assert.ok(data.hkReadingDesks['https://sync.test/']);assert.ok(data.hkReadingDesks['https://desk.test/']);
 console.log('Reading desk sync serialization passed');
})().catch(e=>{console.error(e);process.exitCode=1;});
(async()=>{
 let data={hkReadingDesks:{'https://a.test/':M.normalize({cards:[{id:'n',type:'note',text:'preserve'}]})}};
 const store=HkReadingDeskStore.createStore({get:async()=>structuredClone(data),set:async u=>Object.assign(data,structuredClone(u))});
 assert.equal(typeof store.batch,'function','import/delete use the same shared writer');
 await store.batch([{pageKey:'https://b.test/',value:M.normalize({cards:[{id:'b',type:'note',text:'imported'}]})}]);
 assert.equal(data.hkReadingDesks['https://a.test/'].cards[0].text,'preserve');
 await store.batch([{pageKey:'https://b.test/',value:null}]);assert.equal(data.hkReadingDesks['https://b.test/'],undefined);
 console.log('Reading desk batch import/delete passed');
})().catch(e=>{console.error(e);process.exitCode=1;});
const packed=M.arrange([{id:'a',x:700,y:80,width:320,height:280},{id:'b',x:700,y:120,width:320,height:280},{id:'c',x:700,y:180,width:320,height:280}],1200,800);
assert.equal(packed.length,3);
for(let i=0;i<packed.length;i++)for(let j=i+1;j<packed.length;j++){
 const a=packed[i],b=packed[j];assert.ok(a.x+a.width<=b.x || b.x+b.width<=a.x || a.y+a.height<=b.y || b.y+b.height<=a.y,'pinned notes do not overlap');
}
assert.ok(packed.every(p=>p.x>=8&&p.y>=8&&p.x+p.width<=1192&&p.y+p.height<=704));
console.log('Pinned card packing avoids overlap and stays within the viewport');
(async()=>{
 const url='https://clear.test/';let data={ [url]:[{id:'a',note:'thought'},{id:'b',note:'other'}], hkPersonalNotes:{[url]:{text:'keep article notes'}}, hkReadingDesks:{[url]:M.normalize({cards:[{id:'quote:a',type:'quote',highlightId:'a',text:'quote'},{id:'note:n',type:'note',text:'keep free note'}]})}, 'https://other.test/':[{id:'keep'}] };
 const store=HkReadingDeskStore.createStore({get:async()=>structuredClone(data),set:async u=>Object.assign(data,structuredClone(u))});
 assert.equal(typeof store.deleteHighlights,'function');
 const one=await store.deleteHighlights(url,['a']);assert.equal(one.count,1);assert.deepEqual(data[url].map(h=>h.id),['b']);assert.equal(data.hkReadingDesks[url].cards.length,1);assert.equal(data.hkReadingDesks[url].cards[0].text,'keep free note');
 const all=await store.deleteHighlights(url,null);assert.equal(all.count,1);assert.deepEqual(data[url],[]);assert.equal(data.hkPersonalNotes[url].text,'keep article notes');assert.equal(data['https://other.test/'][0].id,'keep');
 console.log('Highlight deletion preserves free notes, article notes and other pages');
})().catch(e=>{console.error(e);process.exitCode=1;});
(async()=>{
 const url='https://undo.test/';let data={[url]:[{id:'a',note:'original'}],hkReadingDesks:{[url]:M.normalize({cards:[{id:'q',type:'quote',highlightId:'a'}]})}};
 const store=HkReadingDeskStore.createStore({get:async()=>structuredClone(data),set:async u=>Object.assign(data,structuredClone(u))});
 const removed=await store.deleteHighlights(url,['a']);
 data[url].push({id:'new',note:'keep new work'});
 await store.restoreDeleted(url,removed.undo);
 assert.deepEqual(data[url].map(h=>h.id),['new','a']);
 assert.equal(data.hkReadingDesks[url].cards[0].id,'q');
 data[url].find(h=>h.id==='a').note='newer';
 await store.restoreDeleted(url,removed.undo);
 assert.equal(data[url].find(h=>h.id==='a').note,'newer');
 console.log('Undo merges deleted highlights and cards without overwriting newer work');
})().catch(e=>{console.error(e);process.exitCode=1});
