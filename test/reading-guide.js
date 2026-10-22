// Regression cases from reading-guide replies and sidebar tab switching.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
require("../shared.js");
require("../parsers.js");
const P = globalThis.HkParsers;
assert.deepEqual(P.getReadingGuideLimits("short article"), { termsMax: 3, highlightsMin: 3, highlightsMax: 4 });
assert.deepEqual(P.getReadingGuideLimits("word ".repeat(900)), { termsMax: 4, highlightsMin: 5, highlightsMax: 7 });
assert.deepEqual(P.getReadingGuideLimits("word ".repeat(2000)), { termsMax: 5, highlightsMin: 8, highlightsMax: 12 });
assert.equal(P.getReadingGuideLimits("文".repeat(5000)).highlightsMax, 12, "long CJK articles receive the same passage budget");
const articleUrl = "https://elevenlabs.io/zh/blog/realtime";
const otherUrl = "https://huggingface.co/blog/nvidia/nemotron-retriever-world-class-agentic-retrieval";
const oldReply = `===脈絡===
因為整合語音元件很複雜，所以採用端到端模型。
===專有名詞===
1. Eleven Realtime

名詞：Eleven Realtime

解釋：整合語音互動的語音轉語音模型。

2. Full-Duplex-Bench

名詞：Full-Duplex-Bench

解釋：評估停頓、打斷與輪次交接。
===重點標註===
原文：Eleven Realtime is a speech-to-speech model.

#fff5b8

備註：整合原本分開的元件，降低系統開發的複雜度。

原文：Max supports complex tasks.

重點：選型要看任務複雜度，而不只看成本。`;
const markdownReply = ["```markdown", `來源文章：${articleUrl}`, oldReply
  .replace("===脈絡===", "## 論述脈絡")
  .replace("===專有名詞===", "## 專有名詞")
  .replace("===重點標註===", "## 重點標註")
  .replaceAll("名詞：", "- **名詞：**")
  .replaceAll("解釋：", "  - **解釋：**")
  .replaceAll("原文：", "- **原文：**")
  .replaceAll("備註：", "  - **備註：**"), "```"].join("\n");
const parsed = P.parseReadingGuideResponse(oldReply);
assert.equal(parsed.terms.length, 2, "blank lines must not separate terms from definitions");
assert.equal(parsed.highlights.length, 2);
assert.ok(parsed.highlights[0].reason.includes("複雜度"), "blank lines must not lose notes");
const markdown = P.parseReadingGuideResponse(markdownReply);
assert.equal(markdown.sourceUrl, articleUrl);
assert.deepEqual(markdown.terms, parsed.terms, "Markdown and legacy replies retain the same explanations");
assert.deepEqual(markdown.highlights, parsed.highlights);
assert.ok(!markdown.highlights.at(-1).reason.includes("```"));
assert.equal(P.matchesReadingGuideSource(markdown, otherUrl, "Eleven Realtime"), false);
assert.equal(P.matchesReadingGuideSource(markdown, articleUrl, "translated article"), true);
assert.equal(P.matchesReadingGuideSource(parsed, otherUrl, "NVIDIA embedding models"), false);
assert.equal(P.matchesReadingGuideSource(parsed, articleUrl, "Eleven Realtime is a speech-to-speech model."), true);
assert.equal(P.parseHighlightBlocks('原文：「Exact quote」\n\n註解：Useful explanation')[0].text, "Exact quote");
assert.equal(P.parseHighlightBlocks('原文：Quote\n備註：First sentence.\nSecond sentence.')[0].reason, "First sentence.\nSecond sentence.");
assert.equal(P.parseReadingGuideResponse('## **專有名詞**\n名詞：Example\n解釋：First sentence.\nSecond sentence.').terms[0].explanation,
  "First sentence.\nSecond sentence.");

async function checkPersistence() {
  const source = fs.readFileSync(require.resolve("../contentScript.js"), "utf8");
  const start = source.indexOf("const saveReadingGuideContext =");
  const end = source.indexOf("// 統計所有頁面", start);
  const entries = [];
  const saved = {};
  let noteUpdates = 0;
  const context = {
    pageKey: articleUrl, HkParsers: P, DEFAULT_COLOR: "#fff5b8",
    aiSettings: { categories: [], usePreview: true }, previewData: [],
    highlightPanelState: { notesByPage: {} },
    chrome: { storage: { local: { get: async () => ({ hkGeneratedNotes: saved }) } } },
    getPagePlainText: () => "Eleven Realtime is a speech-to-speech model. Max supports complex tasks. Full-Duplex-Bench",
    toHexColor: (color) => color,
    normalizeWhitespace: (text) => text.replace(/\s+/g, " ").trim(),
    t: (key) => key, refreshPaletteFromStorage: async () => ["#fff5b8", "#dfeefc"],
    collectColorUsageCounts: async () => ({}), sortPaletteByUsage: (palette) => palette,
    collectPageHighlights: async () => entries, getStoredHighlights: async () => entries,
    setStoredHighlights: async (next) => entries.splice(0, entries.length, ...next),
    formatAiNote: (reason) => `AI: ${reason}`, setAllMarksMetadata: () => { noteUpdates++; },
    saveGeneratedNote: async (url, record) => { saved[url] = record; },
    showPreviewConfirmBar: () => {}, ensurePageMetaTitle: async () => {}, document: { title: "Article" },
  };
  context.applyAutoHighlightTasksBatch = async (tasks) => {
    context.previewData.push(...tasks.map((data) => ({ data: { ...data, url: articleUrl } })));
    return { unmatched: [], appliedCount: tasks.length };
  };
  vm.createContext(context);
  vm.runInContext(source.slice(start, end) + "\nglobalThis.applyGuide = applyReadingGuideResponse;", context);
  const result = await context.applyGuide(markdownReply);
  assert.equal(result.previewing, true);
  assert.equal(saved[articleUrl].guideTerms.length, 2, "preview terms must be persisted before confirmation");
  assert.ok(saved[articleUrl].guideHighlights[0].reason.includes("複雜度"));
  assert.equal(entries.length, 0, "preview must not commit marks");
  entries.push({ id: "existing", text: "Eleven Realtime", kind: "passage", note: "old note" });
  await context.applyGuide(markdownReply);
  assert.equal(context.previewData.length, 4, "reapplying must not duplicate unconfirmed marks");
  assert.equal(entries[0].kind, "term");
  assert.ok(entries[0].note.includes("語音轉語音模型"), "reapplying must update existing explanations");
  assert.ok(noteUpdates > 0);
  const before = JSON.stringify(saved);
  await assert.rejects(() => context.applyGuide(markdownReply.replace(articleUrl, otherUrl)), /sourceMismatch/);
  assert.equal(JSON.stringify(saved), before, "a wrong-page reply must not mutate saved notes");
  context.getPagePlainText = () => "word ".repeat(2000);
  context.previewData = [];
  entries.splice(0);
  const longReply = `來源文章：${articleUrl}\n## 專有名詞\n` +
    Array.from({ length: 5 }, (_, i) => `名詞：Term ${i}\n解釋：This is a useful definition.`).join("\n\n") +
    "\n## 重點標註\n" + Array.from({ length: 10 }, (_, i) => `原文：Passage ${i}\n重點：This section explains a different important result.`).join("\n\n");
  const longResult = await context.applyGuide(longReply);
  assert.equal(longResult.terms.length, 5);
  assert.equal(longResult.highlights.length, 10, "terms must not consume the long article's passage budget");
  assert.equal(saved[articleUrl].guideHighlights.length, 10, "all long-article passages must survive saving");
}

async function checkSidebar() {
  const nodes = new Map();
  const node = () => ({ value: "", children: [], classList: { toggle() {}, add() {} },
    addEventListener() {}, setAttribute() {}, append(...items) { this.children.push(...items); },
    appendChild(item) { this.children.push(item); }, replaceChildren() { this.children = []; } });
  const events = [];
  const local = { hkGeneratedNotes: {} };
  const session = {};
  let active = { id: 1, url: articleUrl };
  const tabs = { 1: active, 2: { id: 2, url: otherUrl } };
  const context = {
    URL, console, setTimeout, clearTimeout,
    document: { getElementById: (id) => { if (!nodes.has(id)) nodes.set(id, node()); return nodes.get(id); },
      createElement: node, querySelector: node, querySelectorAll: () => [] },
    window: { HkI18n: { t: (key) => key, initI18n: async () => {}, applyDOMTranslations() {}, onLangChange() {} }, setTimeout, clearTimeout },
    chrome: {
      tabs: { query: async () => [active], get: async (id) => tabs[id], onActivated: { addListener() {} }, onUpdated: { addListener() {} },
        sendMessage: async (id, message) => {
          if (message.type === "GET_PAGE_HIGHLIGHTS") return { success: true, data: { url: tabs[id].url, title: "Article", pageText: "Article text", highlights: [] } };
          if (message.type === "BUILD_READING_GUIDE_PROMPT") return { success: true, provider: "chatgpt", configured: true, prompt: "Prompt" };
          if (message.type === "APPLY_READING_GUIDE_RESPONSE") {
            assert.equal(message.pageKey, tabs[id].url);
            events.forEach((callback) => callback({ [tabs[id].url]: { newValue: [] } }, "local"));
            return { success: true, terms: parsed.terms, highlights: parsed.highlights, context: "Saved context", appliedCount: 4 };
          }
        } },
      storage: { local: { get: async () => local }, session: { get: async () => session, set: async (value) => Object.assign(session, value) },
        onChanged: { addListener: (callback) => events.push(callback) } },
    },
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(require.resolve("../sidepanel.js"), "utf8") + "\nglobalThis.sidebar = { state, loadActiveTab, applyGuideResponse };", context);
  await new Promise((resolve) => setImmediate(resolve));
  const sidebar = context.sidebar;
  assert.equal(sidebar.state.pageKey, articleUrl);
  nodes.get("aiResponse").value = "unsent response";
  active = { id: 9, url: "https://chatgpt.com/" };
  await sidebar.loadActiveTab();
  assert.equal(sidebar.state.pageKey, articleUrl, "opening ChatGPT must retain the article");
  assert.equal(nodes.get("aiResponse").value, "unsent response");
  active = tabs[2];
  await sidebar.loadActiveTab();
  assert.equal(nodes.get("aiResponse").value, "", "switching articles must clear an old response draft");
  await sidebar.applyGuideResponse("reply");
  assert.equal(sidebar.state.guideTerms.length, 2, "storage events during apply must not discard results");
  assert.equal(sidebar.state.applying, false);
  local.hkGeneratedNotes[otherUrl] = { guideTerms: parsed.terms, guideHighlights: parsed.highlights };
  await sidebar.loadActiveTab();
  assert.equal(sidebar.state.guideTerms.length, 2, "refresh must retain preview explanations");
  nodes.get("aiResponse").value = "old article reply";
  tabs[2].url = "https://example.com/new-article";
  await sidebar.loadActiveTab();
  assert.equal(nodes.get("aiResponse").value, "", "navigation in the same tab must clear an old response too");
  active = { id: 9, url: "https://chatgpt.com/" };
  sidebar.state.page = null;
  await sidebar.loadActiveTab();
  assert.equal(sidebar.state.page, null, "opening on ChatGPT without a saved source must not throw");
  session.hkReadingGuideSourceTabId = 1;
  await sidebar.loadActiveTab();
  assert.equal(sidebar.state.pageKey, articleUrl, "the saved source can be restored from the AI tab");
}

(async () => {
  await checkPersistence();
  await checkSidebar();
  console.log("Reading-guide regressions passed: legacy/Markdown formatting, notes, source validation, preview persistence, repeat apply, and tab switching.");
})().catch((error) => { console.error(error); process.exitCode = 1; });
