const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
require("../shared.js");
const source = fs.readFileSync(require.resolve("../contentScript.js"), "utf8");
const resolverStart = source.indexOf("const normalizePageKey =");
const resolverEnd = source.indexOf("let pageKey = resolvePageKey();", resolverStart);
assert.ok(resolverStart >= 0 && resolverEnd > resolverStart);
const resolverContext = {
  window: { HkUrlKey: globalThis.HkUrlKey, location: { href: "https://ai.kitmori.com/zh/article/gemini-4-carbon-is-now-available" } },
  document: { querySelector: () => ({ href: "https://ai.kitmori.com/zh/daily" }) },
};
vm.createContext(resolverContext);
const initialKey = vm.runInContext(source.slice(resolverStart, resolverEnd) + "\nresolvePageKey()", resolverContext);
assert.equal(initialKey, resolverContext.window.location.href, "a stale canonical URL must not override the current article on first injection");
resolverContext.window.location.href = "https://ai.kitmori.com/zh/article/gemini-4-carbon-is-now-available?id=7&utm_source=mail";
assert.equal(vm.runInContext("resolvePageKey()", resolverContext), "https://ai.kitmori.com/zh/article/gemini-4-carbon-is-now-available?id=7", "known content parameters stay while tracking is removed");
const start = source.indexOf('const LOCATION_CHANGE_EVENT = "hk:locationchange"');
const end = source.indexOf("// ── 分享連結收件端", start);
assert.ok(start >= 0 && end > start);
const interval = [];
const timeouts = [];
const notified = [];
const windowMock = {
  location: { href: "https://ai.kitmori.com/zh/daily" },
  setInterval: (fn) => interval.push(fn),
  setTimeout: (fn) => { timeouts.push(fn); return timeouts.length; },
  clearTimeout() {}, addEventListener() {}, dispatchEvent() {},
};
const context = {
  window: windowMock, history: { pushState() {}, replaceState() {} }, Event: class {},
  pageKey: "https://ai.kitmori.com/zh/daily", highlightPanelState: {}, orphanHighlightCount: 0, orphanNotified: false,
  document: { querySelectorAll: () => [] }, HIGHLIGHT_CLASS: "hk-highlight", unwrapHighlightElement() {},
  updateOrphanNotice() {}, attemptRestoreHighlights() {}, refreshHighlightPanelIfVisible() {},
  notifyPageChanged: () => notified.push(context.pageKey), normalizePageKey: globalThis.HkUrlKey.normalizePageKey,
};
vm.createContext(context);
vm.runInContext(source.slice(start, end), context);
assert.ok(interval.length, "detect route changes made in the page's separate JavaScript world");
windowMock.location.href = "https://ai.kitmori.com/zh/article/gemini-4-carbon-is-now-available";
interval[0]();
while (timeouts.length) timeouts.shift()();
assert.equal(context.pageKey, windowMock.location.href, "article and daily list use different keys");
assert.deepEqual(notified, [windowMock.location.href], "the sidebar is informed of the new page");
windowMock.location.href += "?utm_source=test";
interval[0]();
while (timeouts.length) timeouts.shift()();
assert.equal(notified.length, 1, "tracking parameters do not switch the article");
windowMock.location.href = "https://ai.kitmori.com/zh/daily";
interval[0]();
while (timeouts.length) timeouts.shift()();
assert.equal(context.pageKey, windowMock.location.href, "returning to the list restores its own key");
windowMock.location.href = "https://ai.kitmori.com/zh/article/another";
vm.runInContext("synchronizeCurrentLocation()", context);
assert.equal(context.pageKey, windowMock.location.href, "a request immediately after navigation uses the new page");
console.log("SPA navigation passed: route changes update the article key and notify the sidebar.");
