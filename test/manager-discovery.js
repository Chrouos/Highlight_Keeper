const assert = require('node:assert/strict');
require('../tag-organizer.js');
const api = globalThis.HkTagOrganizer;
assert.equal(api.matches('Gemini API\n速度優化', 'API Gemini'), true);
assert.equal(api.matches('Gemini API', 'API OpenAI'), false);
assert.equal(api.matches('ＡＩ 筆記', 'ai'), true);
assert.match(api.snippet('x'.repeat(200) + '速度優化測試', '速度'), /速度優化/);
const meta = { a: { title: 'A', tags: ['AI', '人工智慧', '其他'] }, b: { tags: ['其他'] } };
assert.deepEqual(api.mergeTags(meta, ['AI', '人工智慧'], '人工智慧').a.tags, ['人工智慧', '其他']);
assert.deepEqual(meta.a.tags, ['AI', '人工智慧', '其他']);
assert.deepEqual(api.mergeTags(meta, ['AI'], '').a.tags, ['人工智慧', '其他']);
assert.deepEqual(api.addTags(meta, ['a', 'c'], ['其他', '新分類']).a.tags, ['AI', '人工智慧', '其他', '新分類']);
assert.deepEqual(api.addTags(meta, ['a', 'c'], ['新分類']).c.tags, ['新分類']);
assert.deepEqual(api.addTags(meta, ['a'], ['新分類']).b, meta.b);

assert.equal(api.matchesTags(['AI'], ['AI', '系統設計']), true);
assert.equal(api.matchesTags(['其他'], ['AI', '系統設計']), false);
assert.equal(api.matchesTags([], ['AI', '__hk_untagged__']), true);
assert.equal(api.matchesTags(['AI'], []), true);
assert.deepEqual(api.paginate([1,2,3,4,5], 2, 2), { items: [3,4], page: 2, pages: 3, total: 5, start: 3, end: 4 });
assert.equal(api.paginate([1], 9, 20).page, 1);
assert.equal(api.paginate([], 1, 20).start, 0);

// Exercise the manager's actual combined filter, including OR semantics and domain/search intersection.
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require.resolve('../manager.js'), 'utf8');
const body = source.slice(source.indexOf('const getFilteredPages = () => {'), source.indexOf('// 把 bulk 備份的 pages'));
const state = { pages: [
 { url: 'https://a.test/1', title: 'Gemini', tags: ['AI'] },
 { url: 'https://a.test/2', title: 'Gemini', tags: ['系統設計'] },
 { url: 'https://b.test/3', title: 'Gemini', tags: ['AI', '系統設計'] },
 { url: 'https://a.test/4', title: 'Other', tags: [] },
], activeTags: new Set(['AI', '系統設計']), searchTerm: '', activeDomain: '' };
const context = { state, window: { HkTagOrganizer: api }, matchesSearch: (p, q) => api.matches(p.title, q), pageDomain: (url) => new URL(url).hostname };
vm.createContext(context);
vm.runInContext(body + ';globalThis.filtered = getFilteredPages;', context);
assert.equal(context.filtered().length, 3); // A page matching both tags appears only once.
state.activeDomain = 'a.test'; state.searchTerm = 'Gemini';
assert.equal(context.filtered().length, 2);
state.activeTags = new Set(['__hk_untagged__']);
assert.equal(context.filtered().length, 0);
state.searchTerm = '';
assert.equal(context.filtered().length, 1);
state.activeTags.clear();
assert.equal(context.filtered().length, 3);
console.log('Manager discovery tests passed');
