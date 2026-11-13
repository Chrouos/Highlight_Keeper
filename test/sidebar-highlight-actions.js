const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const source = fs.readFileSync(require.resolve("../contentScript.js"), "utf8");
const start = source.indexOf('  if (message?.type === "UPDATE_PAGE_HIGHLIGHT" || message?.type === "DELETE_PAGE_HIGHLIGHT")');
const end = source.indexOf('  if (message?.type === "GET_PAGE_HIGHLIGHTS")', start);
assert.ok(start >= 0 && end > start);
const entry = { id: "one", note: "old", color: "#123456" };
let writes = 0;
const context = {
  pageKey: "https://example.com/article", CSS: { escape: (s) => s }, HIGHLIGHT_ATTR: "data-highlight-id",
  document: { querySelectorAll: () => [] }, unwrapHighlightElement() {}, t: (key) => key,
  findHighlightEntry: async (id) => id === entry.id ? entry : null,
  updateHighlightEntry: async (_id, changes) => { writes++; Object.assign(entry, changes); },
  deleteHighlightEntry: async () => { writes++; }, setAllMarksMetadata() {},
};
vm.createContext(context);
const handler = vm.runInContext(`(message, sendResponse) => { ${source.slice(start, end)} }`, context);
const send = (message) => new Promise((resolve) => handler(message, resolve));
(async () => {
  let result = await send({ type: "UPDATE_PAGE_HIGHLIGHT", id: "one", pageKey: "https://example.com/other", note: "new" });
  assert.equal(result.success, false);
  assert.equal(writes, 0, "a stale page cannot edit highlights");
  result = await send({ type: "UPDATE_PAGE_HIGHLIGHT", id: "one", pageKey: context.pageKey, note: "new" });
  assert.equal(result.success, true);
  assert.equal(entry.note, "new");
  assert.equal(entry.color, "#123456", "editing a note keeps its color");
  result = await send({ type: "DELETE_PAGE_HIGHLIGHT", id: "missing", pageKey: context.pageKey });
  assert.equal(result.success, false);
  assert.equal(writes, 1, "unknown IDs cannot delete another highlight");
  console.log("Sidebar highlight actions passed: stale pages and unknown IDs are rejected.");
})().catch((error) => { console.error(error); process.exitCode = 1; });
