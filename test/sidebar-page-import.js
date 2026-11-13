const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
require("../shared.js");
require("../parsers.js");
const source = fs.readFileSync(require.resolve("../contentScript.js"), "utf8");
const start = source.indexOf('  if (message?.type === "IMPORT_PAGE_HIGHLIGHTS")');
const end = source.indexOf('  if (message?.type === "SET_PAGE_TAGS")', start);
assert.ok(start >= 0 && end > start);
const page = "https://example.com/article";
const backup = { type: "highlight-keeper-bulk", pages: [{ url: page, title: "A", entries: [{ id: "one", text: "hello", color: "#fff5b8" }], personalNote: "my note" }] };
let data = {};
const context = {
  pageKey: page, HkParsers: globalThis.HkParsers, normalizePageKey: globalThis.HkUrlKey.normalizePageKey,
  storage: { get: async () => structuredClone(data), set: async (updates) => Object.assign(data, structuredClone(updates)) },
  parseImportedHighlightsPayload: () => { throw new Error("legacy parser must not handle bulk"); },
  importHighlightsFromEntries: async () => { throw new Error("legacy import must not handle bulk"); },
  attemptRestoreHighlights() {}, t: (key) => key,
};
vm.createContext(context);
const handle = vm.runInContext(`(message, sendResponse) => { ${source.slice(start, end)} }`, context);
const send = (message) => new Promise((resolve) => handle(message, resolve));
(async () => {
  const message = { type: "IMPORT_PAGE_HIGHLIGHTS", pageKey: page, text: JSON.stringify(backup) };
  let result = await send(message);
  assert.equal(result.success, true);
  assert.equal(result.count, 1);
  assert.equal(data[page][0].text, "hello");
  assert.equal(data.hkPersonalNotes[page].text, "my note", "a full backup restores personal notes too");
  data[page][0].text = "local change";
  result = await send(message);
  assert.equal(result.skipped, 1);
  assert.equal(data[page][0].text, "local change", "existing page is not overwritten");
  const before = JSON.stringify(data);
  result = await send({ ...message, text: JSON.stringify({ pages: [{ ...backup.pages[0], url: "https://example.com/other" }] }) });
  assert.equal(result.success, false, "a different page cannot be imported as the current page");
  assert.equal(JSON.stringify(data), before);
  console.log("Sidebar page import passed: full backup round-trips and existing data is preserved.");
})().catch((error) => { console.error(error); process.exitCode = 1; });
