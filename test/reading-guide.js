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

// One reply format, however it reaches the clipboard: code-block copy, rendered copy with
// escaped Markdown, or the legacy ===marker=== format.
const unifiedReply = `來源文章：https://example.com/a_b
## 脈絡
主線說明。
## 名詞解釋
- **名詞：** RAG
  - **解釋：** 先檢索再生成。
## 重點整理
- **原文：** Exact quote
  - **分類：** #fff5b8
  - **重點：** 為何重要。
## 閱讀地圖
# 文章主題
- 背景
  - 問題
- 結論
## 標籤
AI, 檢索`;
const unified = P.parseReadingGuideResponse(unifiedReply);
assert.equal(unified.context, "主線說明。");
assert.equal(unified.terms[0].text, "RAG");
assert.equal(unified.highlights[0].reason, "為何重要。");
assert.equal(P.parseMindmapOutline(unified.mindmap).title, "文章主題");
assert.equal(P.parseMindmapOutline(unified.mindmap).children[0].children[0].label, "問題");
assert.deepEqual(P.parsePageTags(unified.tags), ["AI", "檢索"]);
const escapeMarkdown = (text) => text.replace(/[=#*_-]/g, (char) => `\\${char}`);
assert.deepEqual(P.parseReadingGuideResponse(escapeMarkdown(unifiedReply)), unified, "escaped Markdown parses identically");
const legacyMarkers = P.parseReadingGuideResponse(escapeMarkdown("===脈絡===\n主線說明。\n===重點===\n原文：Exact quote\n重點：為何重要。"));
assert.equal(legacyMarkers.context, "主線說明。", "\\=\\=\\= markers are recognized");
assert.equal(legacyMarkers.highlights[0].text, "Exact quote");
const renderedCopy = P.parseReadingGuideResponse("脈絡\n主線說明。\n重點整理\n原文：Exact quote\n重點：為何重要。");
assert.equal(renderedCopy.context, "主線說明。", "headings survive losing their ## prefix");
assert.equal(renderedCopy.highlights.length, 1);
assert.equal(P.parseSharedMarkdown(unifiedReply), null, "a guide reply is not mistaken for shared notes");

// Notes read line by line: new replies keep their lines, old walls of text get broken at display.
const lined = P.parseReadingGuideResponse(`## 名詞解釋
- **名詞：** A/B 測試
  - **解釋：** 把使用者分到不同版本。
  - **例子：** 部分請求用新模型。
  - **在本文：** 只是可能原因。
- **名詞：** Empty first line
  - **解釋：**
    說明在下一行。
## 重點整理
- **原文：** Exact quote
  - **分類：** #fff5b8
  - **重點：** 這段在講什麼。
  - **為何重要：** 它和主線的關係。`);
assert.equal(lined.terms[0].explanation, "把使用者分到不同版本。\n例子：部分請求用新模型。\n在本文：只是可能原因。");
assert.equal(lined.terms[1].explanation, "說明在下一行。");
assert.equal(lined.highlights[0].reason, "這段在講什麼。\n為何重要：它和主線的關係。");
assert.equal(P.parseReadingGuideResponse("## 名詞解釋\n名詞：X\n解釋：第一句。\n說明：第二句。").terms[0].explanation, "第一句。\n第二句。", "a second label adds a line");
assert.equal(P.parseHighlightBlocks("原文：Q\n重點：第一句。\n說明：第二句。")[0].reason, "第一句。\n第二句。");
const wall = "AI：A/B 測試是將不同使用者或請求分配到不同版本，藉此比較系統表現的方法。例如部分 Codex 請求可能由新版模型處理，其餘維持舊版。本文將它列為新模型流量出現的可能原因，但沒有證據確認 OpenAI 確實正在執行 A/B 測試。";
assert.equal(P.breakNoteLines(wall).split("\n").length, 3, "an old note breaks at each sentence");
assert.equal(P.breakNoteLines("短句。還是短句。"), "短句。還是短句。", "short notes stay on one line");
assert.equal(P.breakNoteLines("First line.\n\nSecond line."), "First line.\n\nSecond line.", "existing breaks are kept");
const english = "The recall went up to 3.5 percent in the test, e.g. Foo beat Bar. The cost also rose! Is it worth it?";
assert.equal(P.breakNoteLines(english).split("\n").length, 3, "decimals and abbreviations do not split");
assert.equal(P.breakNoteLines("「這是一段夠長的引用句子，長到需要被斷開來才好閱讀。」下一句也很長，一樣要留在自己的那一行才行。").split("\n")[0].endsWith("」"), true, "closing quotes stay with their sentence");
assert.equal(P.looseKey("“Don’t” — stop"), P.looseKey('"Don\'t" - stop'), "curly and straight quotes match");
assert.equal(P.pickTermText("Model ID（模型識別碼）", "The model id is new."), "Model ID", "translation in brackets is ignored");
assert.equal(P.pickTermText("模型識別碼 (Model ID)", "The Model ID is new."), "Model ID", "the side found in the article wins");
assert.equal(P.pickTermText("Staged Rollout（分階段推出）", "Nothing about it."), "Staged Rollout", "falls back to the first part");
assert.equal(P.pickTermText("GPT (beta)", "Try GPT (beta) today."), "GPT (beta)", "brackets in the article itself are kept");
const pick = P.pickTermText;
assert.equal(pick("A/B Testing（A/B 測試）", "either as an A/B test or a rollout"), "A/B test", "-ing form finds the base word and keeps the article's wording");
assert.equal(pick("A/B Test", "They were A/B testing it."), "A/B testing", "base word finds the -ing form");
assert.equal(pick("Staged Rollouts", "the start of a staged\n rollout."), "staged\n rollout", "plural finds singular across a line break");
assert.equal(pick("Rollout", "Two rollouts shipped."), "rollouts", "singular finds plural");
assert.equal(pick("Policy", "Three policies apply."), "policies", "-y finds -ies");
assert.equal(pick("Policies", "One policy applies."), "policy", "-ies finds -y");
assert.equal(pick("Caching", "Use a cache here."), "cache", "dropped e is restored");
assert.equal(pick("Cache", "We are caching it."), "caching", "e is dropped before -ing");
assert.equal(pick("Embedding", "We embed it."), "embed", "doubled letter is undone");
assert.equal(pick("Run", "The running job."), "running", "doubled letter is added");
assert.equal(pick("Testing", "We test it. Then testing again."), "Testing", "exact wording wins over a variant");
assert.equal(pick("Rate", "A rat ran past."), "Rate", "rate does not match rat");
assert.equal(pick("String", "Say str here."), "String", "string does not reduce to str");
assert.equal(pick("Status", "Statu here."), "Status", "-us is not a plural");
assert.equal(pick("Test", "the latest contest"), "Test", "a word inside another word does not count");
assert.equal(pick("Model ID", "the model identifier"), "Model ID", "short words stay exact");
assert.equal(pick("Testing Tools", "testing tool"), "testing tool", "only the head word flexes");
assert.equal(pick("Tests Tool", "testing tool"), "Tests Tool", "words before the head stay exact");
assert.equal(pick("測試", "這是 test"), "測試", "non-English terms are left alone");
assert.equal(pick("C++ (語言)", "C++ is old"), "C++", "regex characters are safe");
assert.equal(pick("Full-Duplex-Bench", "The full-duplex-benches ran."), "full-duplex-benches", "hyphens in a term are safe");
assert.equal(pick("gpt-6.2-astra（新模型）", "only gpt-6.2-astra shows up"), "gpt-6.2-astra", "model names with dots and hyphens are safe");
assert.equal(pick("[RAG] {x} a|b $1 ^c", "nothing"), "[RAG] {x} a|b $1 ^c", "regex syntax in a term does not throw");
const multiLineMd = P.pageToMarkdown({ title: "T", url: "https://example.com/a", entries: [{ text: "Q", note: "第一行\n第二行" }] });
assert.equal(P.parseSharedMarkdown(multiLineMd).items[0].note, "第一行\n第二行", "shared Markdown keeps note lines");

async function checkPersistence() {
  const source = fs.readFileSync(require.resolve("../contentScript.js"), "utf8");
  const start = source.indexOf("const saveReadingGuideContext =");
  const end = source.indexOf("// 統計所有頁面", start);
  const entries = [];
  const saved = {};
  let noteUpdates = 0;
  const context = {
    pageKey: articleUrl, HkParsers: P, DEFAULT_COLOR: "#fff5b8",
    MINDMAP_STORAGE_KEY: "hkMindmaps", PAGE_META_KEY: "__hk_page_meta__",
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
  const requests = [];
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
          requests.push(message);
          if (message.type === "GET_PAGE_HIGHLIGHTS") return { success: true, data: { url: tabs[id].url, title: "Article", pageText: "Article text", highlights: [] } };
          if (message.type === "BUILD_READING_GUIDE_PROMPT") return { success: true, provider: "chatgpt", configured: true, prompt: "Prompt" };
          if (message.type === "UPDATE_PAGE_HIGHLIGHT") return { success: true };
          if (message.type === "SET_PAGE_TAGS") return { success: true };
          if (message.type === "IMPORT_PAGE_HIGHLIGHTS") return { success: true, count: 1 };
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
  vm.runInContext(fs.readFileSync(require.resolve("../sidepanel.js"), "utf8") + "\nglobalThis.sidebar = { state, loadActiveTab, applyGuideResponse, activateMainView, changePageHighlight, savePageTags, importPageHighlights };", context);
  await new Promise((resolve) => setImmediate(resolve));
  const sidebar = context.sidebar;
  sidebar.activateMainView("settings");
  assert.equal(nodes.get("view-settings").hidden, false, "settings is available in the side panel");
  assert.equal(nodes.get("view-reading").hidden, true, "the reading view is hidden when settings is open");
  sidebar.activateMainView("reading");
  assert.equal(nodes.get("view-reading").hidden, false, "returning to reading restores the view");
  assert.equal(sidebar.state.pageKey, articleUrl);
  await sidebar.changePageHighlight("mark-1", { note: "note" });
  assert.ok(requests.slice(-2).some((request) => request.type === "GET_PAGE_HIGHLIGHTS"), "the list refreshes after saving a highlight");
  assert.ok(requests.some((request) => request.type === "UPDATE_PAGE_HIGHLIGHT" && request.id === "mark-1" && request.pageKey === articleUrl && request.note === "note"), "updates target the original page and exact highlight");
  await sidebar.savePageTags(["alpha", "beta"]);
  assert.ok(requests.some((request) => request.type === "SET_PAGE_TAGS" && request.pageKey === articleUrl && request.tags.join(",") === "alpha,beta"), "tags are saved for the active page");
  assert.equal((await sidebar.importPageHighlights("{}")).count, 1);
  assert.ok(requests.some((request) => request.type === "IMPORT_PAGE_HIGHLIGHTS" && request.pageKey === articleUrl), "imports are bound to the active page");
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
  assert.equal(nodes.get("moreTools").hidden, false, "tools remain accessible without an article");
  session.hkReadingGuideSourceTabId = 1;
  await sidebar.loadActiveTab();
  assert.equal(sidebar.state.pageKey, articleUrl, "the saved source can be restored from the AI tab");
}

(async () => {
  await checkPersistence();
  await checkSidebar();
  console.log("Reading-guide regressions passed: legacy/Markdown formatting, notes, source validation, preview persistence, repeat apply, and tab switching.");
})().catch((error) => { console.error(error); process.exitCode = 1; });
