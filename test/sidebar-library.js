const assert = require('node:assert/strict');
require('../sidebar-library-data.js');
const data = { __hk_page_meta__: { 'https://a.test/': { title: 'AI article', tags: ['研究'] }, 'https://ghost.test/': { title: 'Visited only' } }, 'https://a.test/': [{ text: 'Gemini', note: '重點', createdAt: 10 }], hkPersonalNotes: { 'https://b.test/': { text: '自己的筆記', updatedAt: 20 } } };
const pages = globalThis.HkSidebarLibrary.pages(data);
assert.equal(pages.length, 2);
assert.equal(pages[0].url, 'https://b.test/');
assert.equal(globalThis.HkSidebarLibrary.search(pages, 'Gemini 研究').length, 1);
assert.equal(globalThis.HkSidebarLibrary.search(pages, '自己的').length, 1);
assert.equal(globalThis.HkSidebarLibrary.search(pages, '不存在').length, 0);
console.log('Sidebar library tests passed');
const articleNotes=HkSidebarLibrary.articleNotes([
 {id:'a',text:'Original quote',note:'Canonical thought'}, {id:'b',text:'No thoughts yet',note:''},
],{cards:[{id:'quote:a',type:'quote',highlightId:'a',text:'Original quote'},{id:'note:c',type:'note',text:'Independent thought'},{id:'note:empty',type:'note',text:''}]});
assert.equal(articleNotes.length,2);
assert.equal(articleNotes.find(n=>n.type==='quote').text,'Canonical thought');
assert.equal(articleNotes.find(n=>n.type==='note').text,'Independent thought');
assert.equal(articleNotes.filter(n=>n.id==='quote:a').length,1);
console.log('Article notes overview uses canonical thoughts and free notes');
