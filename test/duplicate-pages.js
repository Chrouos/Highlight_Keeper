const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
require("../shared.js");
const source = fs.readFileSync(require.resolve("../manager.js"), "utf8");
const start = source.indexOf("const canonicalPageKey =");
const end = source.indexOf("// 實際把群組合併寫回 storage", start);
assert.ok(start >= 0 && end > start);
const daily = "https://ai.kitmori.com/zh/daily";
const article = "https://ai.kitmori.com/zh/article/gemini-4-carbon-is-now-available";
const context = { URL, window: { HkUrlKey: globalThis.HkUrlKey }, pageDomain: () => "ai.kitmori.com", state: {
  pages: [{ url: daily }, { url: article }, { url: `${article}?utm_source=mail` }],
  meta: { [daily]: { title: "Same title" }, [article]: { title: "Same title" } },
} };
vm.createContext(context);
const groups = vm.runInContext(source.slice(start, end) + "\nfindDuplicateGroups()", context);
assert.equal(groups.length, 1);
assert.equal(groups[0].canonical, article);
assert.deepEqual(Array.from(groups[0].pages, (p) => p.url).sort(), [article, `${article}?utm_source=mail`].sort());
console.log("Duplicate page grouping passed: distinct paths remain separate and tracking variants merge.");

(async () => {
  const mergeStart = source.indexOf("const mergeDuplicateGroups =");
  const mergeEnd = source.indexOf("// 點「整理重複頁面」", mergeStart);
  const variant = `${article}?utm_source=mail`;
  const store = { __hk_page_meta__: {}, hkGeneratedNotes: {}, hkMindmaps: {}, hkPersonalNotes: { [article]: { text: "first" }, [variant]: { text: "different" } } };
  let writes = 0;
  const mergeContext = { PAGE_META_KEY: "__hk_page_meta__", GENERATED_NOTES_KEY: "hkGeneratedNotes", MINDMAP_KEY: "hkMindmaps", PERSONAL_NOTES_KEY: "hkPersonalNotes", mergeEntryArrays: (x) => x.flat(), chrome: { storage: { local: { get: async () => structuredClone(store), set: async () => { writes++; }, remove: async () => { writes++; } } } } };
  vm.createContext(mergeContext);
  const merge = vm.runInContext(source.slice(mergeStart, mergeEnd) + "\nmergeDuplicateGroups", mergeContext);
  const result = await merge([{ canonical: article, pages: [{ url: article, entries: [] }, { url: variant, entries: [] }] }]);
  assert.equal(writes, 0, "conflicting notes are never overwritten or deleted");
  assert.equal(result.skipped, 1);
  delete store.hkPersonalNotes[variant];
  const ok = await merge([{ canonical: article, pages: [{ url: article, entries: [{ id: "a" }] }, { url: variant, entries: [{ id: "b" }] }] }]);
  assert.equal(ok.mergedGroups, 1);
  assert.equal(ok.mergedExtra, 1);
  assert.equal(writes, 2, "nonconflicting variants are saved and old keys removed");
  console.log("Duplicate page protection passed: conflicting notes are kept intact.");
})().catch((error) => { console.error(error); process.exitCode = 1; });
