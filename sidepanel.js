/* 閱讀導覽側欄：整個擴充功能唯一的閱讀與 AI 入口。 */
const PERSONAL_NOTES_KEY = "hkPersonalNotes";
const GENERATED_NOTES_KEY = "hkGeneratedNotes";
const MINDMAP_KEY = "hkMindmaps";
const PAGE_META_KEY = "__hk_page_meta__";
const GITHUB_SETTINGS_KEY = "hkGithubSyncSettings";
const SYNC_STATE_KEY = "hkAutoSyncState";
const AI_SETTINGS_KEY = "hkAISettings";
const WEB_PROVIDERS = ["chatgpt", "gemini-web"];

const els = {
  title: document.getElementById("pageTitle"),
  url: document.getElementById("pageUrl"),
  status: document.getElementById("status"),
  aiControls: document.getElementById("aiControls"),
  noPage: document.getElementById("noPage"),
  context: document.getElementById("contextText"),
  emptyContext: document.getElementById("emptyContext"),
  mapSection: document.getElementById("mapSection"),
  map: document.getElementById("readingMap"),
  startGuide: document.getElementById("startGuideBtn"),
  terms: document.getElementById("termsList"),
  emptyTerms: document.getElementById("emptyTerms"),
  highlights: document.getElementById("highlightsList"),
  emptyHighlights: document.getElementById("emptyHighlights"),
  termsCount: document.getElementById("termsCount"),
  highlightsCount: document.getElementById("highlightsCount"),
  personalNote: document.getElementById("personalNote"),
  noteStatus: document.getElementById("noteSaveStatus"),
  generate: document.getElementById("generateBtn"),
  copyPrompt: document.getElementById("copyPromptBtn"),
  openWebAI: document.getElementById("openWebAiBtn"),
  response: document.getElementById("aiResponse"),
  applyResponse: document.getElementById("applyResponseBtn"),
  settings: document.getElementById("settingsBtn"),
  refresh: document.getElementById("refreshBtn"),
  tabs: document.querySelector(".tabs"),
  moreTools: document.getElementById("moreTools"),
  shareLink: document.getElementById("shareLinkBtn"),
  copyNotes: document.getElementById("copyNotesBtn"),
  downloadNotes: document.getElementById("downloadNotesBtn"),
  openTools: document.getElementById("openToolsBtn"),
  openManager: document.getElementById("openManagerBtn"),
  langSelect: document.getElementById("langSelect"),
  syncNotice: document.getElementById("syncNotice"),
  deletePageNotes: document.getElementById("deletePageNotesBtn"),
};

const t = (key, params) => window.HkI18n.t(key, params);
const state = {
  tabId: null,
  pageKey: "",
  page: null,
  context: "",
  guideTerms: [],
  guideHighlights: [],
  mindmapOutline: "",
  renderedOutline: null,
  provider: "chatgpt",
  configured: false,
  prompt: "",
  noteText: "",
  noteDirty: false,
  noteTimer: 0,
  loadToken: 0,
  activeTab: "context",
  applying: false,
};

const isAIWebsite = (url) => {
  try { return ["chatgpt.com", "gemini.google.com"].includes(new URL(url).hostname); }
  catch (_error) { return false; }
};

const guideStatus = (result) => {
  const summary = t(result.previewing ? "sidebar.guidePreview" : "sidebar.guideSummary", {
    terms: result.terms?.length || 0,
    highlights: result.highlights?.length || 0,
    count: result.previewing ? (result.previewCount ?? result.appliedCount ?? 0) : (result.appliedCount || 0),
  });
  return summary + (result.unmatchedCount ? t("sidebar.guideUnmatched", { count: result.unmatchedCount }) : "");
};

const setStatus = (message = "", isError = false) => {
  els.status.textContent = message;
  els.status.hidden = !message;
  els.status.classList.toggle("is-error", isError);
};

const setNoteStatus = (key) => {
  els.noteStatus.textContent = t(key);
};

const wait = (ms) => new Promise((resolve) => window.setTimeout(resolve, ms));

const sendToTab = async (tabId, message) => {
  try {
    return await chrome.tabs.sendMessage(tabId, message);
  } catch (firstError) {
    const tab = await chrome.tabs.get(tabId).catch(() => null);
    if (!tab?.url || !/^https?:/i.test(tab.url)) throw firstError;
    // 頁面還在載入時內容腳本會自己注入；載入完成後 onUpdated 會再讀一次。
    if (tab.status === "loading") throw firstError;
    // 重新整理的瞬間分頁可能還標示 complete：先等一下再試，真的沒有內容腳本才注入。
    await wait(300);
    const retried = await chrome.tabs.sendMessage(tabId, message).then(
      (response) => ({ response }),
      () => null
    );
    if (retried) return retried.response;
    const latest = await chrome.tabs.get(tabId).catch(() => null);
    if (latest?.status !== "complete") throw firstError;
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["shared.js", "parsers.js", "i18n.js", "highlightDom.js", "contentScript.js"],
    });
    await wait(180);
    return chrome.tabs.sendMessage(tabId, message);
  }
};

const saveNote = async (pageKey, text) => {
  if (!pageKey) return;
  const stored = await chrome.storage.local.get([PERSONAL_NOTES_KEY, PAGE_META_KEY]);
  const notes = { ...(stored[PERSONAL_NOTES_KEY] || {}) };
  notes[pageKey] = { text, updatedAt: Date.now() };
  const updates = { [PERSONAL_NOTES_KEY]: notes };
  const title = state.pageKey === pageKey ? state.page?.title : "";
  if (title && !stored[PAGE_META_KEY]?.[pageKey]?.title) {
    updates[PAGE_META_KEY] = {
      ...(stored[PAGE_META_KEY] || {}),
      [pageKey]: { ...stored[PAGE_META_KEY]?.[pageKey], title, updatedAt: Date.now() },
    };
  }
  await chrome.storage.local.set(updates);
};

const flushNote = async () => {
  if (state.noteTimer) window.clearTimeout(state.noteTimer);
  state.noteTimer = 0;
  if (!state.noteDirty || !state.pageKey) return;
  const pageKey = state.pageKey;
  const text = els.personalNote.value;
  state.noteDirty = false;
  setNoteStatus("sidebar.notesSaving");
  try {
    await saveNote(pageKey, text);
    if (state.pageKey === pageKey) setNoteStatus("sidebar.notesSaved");
  } catch (error) {
    console.debug("儲存個人筆記失敗", error);
    state.noteDirty = state.pageKey === pageKey;
    if (state.noteDirty) setNoteStatus("sidebar.notesUnsaved");
  }
};

const breakLines = (text) => window.HkParsers?.breakNoteLines(text) ?? String(text ?? "");

// 一行一個區塊；「例子：」這類開頭的小標加粗，掃一眼就知道每行在講什麼。
const fillNoteLines = (container, text) => {
  breakLines(text).split("\n").forEach((line) => {
    const content = line.trim();
    if (!content) return;
    const row = document.createElement("span");
    row.className = "note-line";
    const labeled = content.match(/^([^：:\s]{1,8}[：:])\s*(.*)$/);
    if (labeled && !/^https?:/i.test(content)) {
      const label = document.createElement("strong");
      label.textContent = labeled[1];
      row.append(label, labeled[2]);
    } else {
      row.textContent = content;
    }
    container.appendChild(row);
  });
};

const makeItemButton = (item, kind) => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "item-button";
  // 色條對應頁面上的標註顏色，讓清單與原文對得起來。
  if (item.color) button.style?.setProperty("--item-color", item.color);
  if (item.unavailable && !item.covered) {
    button.disabled = true;
    button.classList.add("is-unavailable");
  }
  if ((item.unavailable && !item.covered) || item.notOnPage) {
    const unavailable = document.createElement("span");
    // 名詞本來就可能只是概念，沒有標記仍有解釋，不用警示。
    const termOnly = kind === "term" && item.unavailable && !item.notApplied;
    unavailable.className = termOnly ? "item-kind is-quiet" : "item-kind";
    unavailable.textContent = t(item.notApplied ? "sidebar.notApplied" : termOnly ? "sidebar.termNotInArticle" : "sidebar.unavailable");
    button.appendChild(unavailable);
  }
  button.addEventListener("click", async () => {
    if (!state.tabId) return;
    try {
      const result = await sendToTab(state.tabId, {
        type: "FOCUS_SAVED_HIGHLIGHT",
        text: item.text,
        kind,
      });
      if (!result?.success) setStatus(result?.error || t("sidebar.noMatch"), true);
      else setStatus("");
    } catch (_error) {
      setStatus(t("sidebar.noMatch"), true);
    }
  });
  const text = document.createElement("span");
  text.className = "item-text";
  text.textContent = item.label || item.text || "";
  button.appendChild(text);
  const note = item.explanation || item.reason || item.note || "";
  if (note) {
    const description = document.createElement("span");
    description.className = "item-note";
    fillNoteLines(description, note);
    button.appendChild(description);
  }
  return button;
};

const deleteItem = async (item, kind) => {
  try {
    const result = await sendToTab(state.tabId, {
      type: "DELETE_GUIDE_ITEM",
      text: item.text,
      kind,
      pageKey: state.pageKey,
    });
    if (!result?.success) throw new Error(result?.error);
    await loadActiveTab();
  } catch (error) {
    setStatus(error?.message || t("sidebar.deleteError"), true);
  }
};

const makeItemRow = (item, kind) => {
  const row = document.createElement("div");
  row.className = "item-row";
  const remove = document.createElement("button");
  remove.type = "button";
  remove.className = "item-delete";
  remove.textContent = "×";
  remove.title = t("sidebar.deleteItem");
  remove.setAttribute("aria-label", t("sidebar.deleteItem"));
  remove.addEventListener("click", () => deleteItem(item, kind));
  row.append(makeItemButton(item, kind), remove);
  return row;
};

const renderItems = (container, empty, items, kind) => {
  container.replaceChildren();
  items.forEach((item) => container.appendChild(makeItemRow(item, kind)));
  empty.hidden = items.length > 0;
};

const buildMapList = (nodes, depth) => {
  const list = document.createElement("ul");
  list.className = "map-list";
  nodes.forEach((node) => {
    const item = document.createElement("li");
    item.className = "map-node";
    if (node.children.length) {
      const details = document.createElement("details");
      details.open = depth < 2;
      const summary = document.createElement("summary");
      summary.textContent = node.label;
      details.append(summary, buildMapList(node.children, depth + 1));
      item.appendChild(details);
    } else {
      item.textContent = node.label;
    }
    list.appendChild(item);
  });
  return list;
};

// 只有大綱變動才重建，重繪時不會把使用者收合的節點又展開。
const renderReadingMap = () => {
  if (state.renderedOutline === state.mindmapOutline) return;
  state.renderedOutline = state.mindmapOutline;
  els.map.replaceChildren();
  const tree = state.mindmapOutline
    ? window.HkParsers?.parseMindmapOutline(state.mindmapOutline, state.page?.title || "")
    : null;
  if (!tree) return;
  const rootTitle = document.createElement("p");
  rootTitle.className = "map-root";
  rootTitle.textContent = tree.title;
  els.map.append(rootTitle, buildMapList(tree.children, 0));
};

// 與內容腳本相同的寬鬆比對：AI 引用和實際標到的文字可能差在空白與標點。
const itemKey = (text) => window.HkParsers?.looseKey(text || "") ?? (text || "").trim().toLowerCase();

const mergeItems = (stored, extra) => {
  const items = stored.map((item) => ({ ...item, reason: item.note }));
  const indexes = new Map(items.map((item, index) => [itemKey(item.text), index]));
  extra.forEach((item) => {
    const key = itemKey(item.text);
    if (key && indexes.has(key)) {
      const index = indexes.get(key);
      items[index] = { ...items[index], ...item, unavailable: false };
    } else if (key) {
      indexes.set(key, items.length);
      // 當初有對到原文、現在卻沒標記：是預覽沒確認就離開，不是原文不存在。
      items.push({ ...item, notApplied: !item.unavailable && !item.covered, unavailable: true });
    }
  });
  // AI 依文章順序列出項目，照這個順序才能帶著讀；手動標註排在後面。
  const guideOrder = new Map(extra.map((item, index) => [itemKey(item.text), index]));
  const rank = (item) => guideOrder.get(itemKey(item.text)) ?? extra.length;
  return items.sort((a, b) => rank(a) - rank(b));
};

const render = () => {
  const page = state.page;
  els.title.textContent = page?.title || "—";
  els.url.textContent = page?.url || "";
  els.context.textContent = breakLines(state.context || "");
  els.emptyContext.hidden = Boolean(state.context);

  const storedHighlights = Array.isArray(page?.highlights) ? page.highlights : [];
  const visibleHighlights = [...storedHighlights, ...(page?.previewHighlights || [])];
  const terms = mergeItems(visibleHighlights.filter((item) => item.kind === "term"), state.guideTerms);
  const highlights = mergeItems(visibleHighlights.filter((item) => item.kind !== "term"), state.guideHighlights);
  renderItems(els.terms, els.emptyTerms, terms, "term");
  renderItems(els.highlights, els.emptyHighlights, highlights, "passage");
  els.termsCount.textContent = String(terms.length);
  els.termsCount.hidden = terms.length === 0;
  els.highlightsCount.textContent = String(highlights.length);
  els.highlightsCount.hidden = highlights.length === 0;

  renderReadingMap();
  els.mapSection.hidden = !state.mindmapOutline && visibleHighlights.length === 0;
  els.startGuide.disabled = visibleHighlights.length === 0;

  const isWebProvider = WEB_PROVIDERS.includes(state.provider);
  els.generate.hidden = isWebProvider;
  els.openWebAI.hidden = !isWebProvider;
  els.openWebAI.textContent = t(state.provider === "gemini-web" ? "sidebar.openGemini" : "sidebar.openChatgpt");
  els.applyResponse.disabled = state.applying;
  els.generate.disabled = state.applying;
  els.noPage.hidden = Boolean(page);
  els.aiControls.hidden = !page;
  els.moreTools.hidden = !page;
  els.deletePageNotes.hidden = !page;
  els.tabs.hidden = !page;
  els.personalNote.disabled = !page;
  activateTab(state.activeTab);
};

const activateTab = (name) => {
  state.activeTab = name;
  document.querySelectorAll("[role=tab]").forEach((tab) => {
    const selected = tab.dataset.tab === name;
    tab.setAttribute("aria-selected", String(selected));
    tab.tabIndex = selected ? 0 : -1;
  });
  document.querySelectorAll("[role=tabpanel]").forEach((pane) => {
    pane.hidden = !state.page || pane.id !== `pane-${name === "personal-notes" ? "notes" : name}`;
  });
};

// 背景同步有衝突或失敗時才出現；平常不佔位子。
const renderSyncNotice = (syncState) => {
  const conflicts = syncState?.conflicts?.length || 0;
  const failed = syncState?.status === "error" && syncState.errorCode !== "remote-moved";
  els.syncNotice.hidden = !conflicts && !failed;
  els.syncNotice.textContent = conflicts
    ? t("sidebar.syncConflicts", { count: conflicts })
    : failed ? t("sidebar.syncError") : "";
};

const hasGuide = () => Boolean(state.context || state.guideTerms.length || state.guideHighlights.length);

const loadActiveTab = async () => {
  const previousPageKey = state.pageKey;
  let [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  if (isAIWebsite(tab?.url)) {
    // Keep the source article while the user chats with an AI in another tab.
    if (state.page) return;
    const saved = await chrome.storage.session.get("hkReadingGuideSourceTabId");
    tab = Number.isInteger(saved.hkReadingGuideSourceTabId)
      ? await chrome.tabs.get(saved.hkReadingGuideSourceTabId).catch(() => null) : null;
  }
  const token = ++state.loadToken;
  await flushNote();
  if (token !== state.loadToken) return;
  if (state.tabId !== tab?.id) els.response.value = "";
  state.tabId = tab?.id ?? null;
  state.pageKey = "";
  state.page = null;
  state.context = "";
  state.guideTerms = [];
  state.guideHighlights = [];
  state.mindmapOutline = "";
  state.prompt = "";
  state.noteText = "";
  state.noteDirty = false;
  els.personalNote.value = "";
  setNoteStatus("sidebar.notesSaved");
  setStatus("");
  render();
  if (!tab?.id || !/^https?:/i.test(tab.url || "")) {
    setStatus(t("sidebar.unsupported"), true);
    return;
  }

  try {
    const response = await sendToTab(tab.id, { type: "GET_PAGE_HIGHLIGHTS" });
    if (token !== state.loadToken) return;
    if (!response?.success || !response.data?.url) throw new Error(response?.error || t("sidebar.extractError"));
    if (previousPageKey && previousPageKey !== response.data.url) els.response.value = "";
    state.page = response.data;
    state.pageKey = response.data.url;
    const [stored, promptInfo] = await Promise.all([
      chrome.storage.local.get([GENERATED_NOTES_KEY, PERSONAL_NOTES_KEY, MINDMAP_KEY]),
      sendToTab(tab.id, { type: "BUILD_READING_GUIDE_PROMPT" }),
    ]);
    if (token !== state.loadToken) return;
    const generated = stored[GENERATED_NOTES_KEY]?.[state.pageKey] || {};
    const personal = stored[PERSONAL_NOTES_KEY]?.[state.pageKey] || {};
    state.context = generated.context || generated.note || "";
    state.guideTerms = generated.guideTerms || generated.unlocatedTerms || [];
    state.guideHighlights = generated.guideHighlights || generated.unlocatedHighlights || [];
    state.mindmapOutline = stored[MINDMAP_KEY]?.[state.pageKey]?.outline || "";
    state.noteText = personal.text || "";
    els.personalNote.value = state.noteText;
    state.provider = promptInfo?.provider || "chatgpt";
    state.configured = Boolean(promptInfo?.configured);
    state.prompt = promptInfo?.prompt || "";
    if (!promptInfo?.success) setStatus(t("sidebar.extractError"), true);
    // 換文章時才決定收合：已整理過的把位子留給閱讀，沒整理過就直接看到 AI 步驟。
    // 同一篇文章因標註變動重新載入時，不動使用者正在操作的區塊。
    if (previousPageKey !== state.pageKey) els.aiControls.open = !hasGuide();
    render();
    if (!response.data.pageText?.trim()) setStatus(t("sidebar.extractError"), true);
  } catch (error) {
    if (token !== state.loadToken) return;
    console.debug("讀取側欄文章資料失敗", error);
    state.page = null;
    render();
    setStatus(t("sidebar.extractError"), true);
  }
};

const applyGuideResponse = async (text) => {
  if (!state.tabId || !text || state.applying) return;
  const request = { tabId: state.tabId, pageKey: state.pageKey, loadToken: state.loadToken, provider: state.provider };
  const isCurrent = () => request.tabId === state.tabId && request.pageKey === state.pageKey && request.loadToken === state.loadToken;
  state.applying = true;
  render();
  setStatus(t("sidebar.generating"));
  try {
    const result = await sendToTab(request.tabId, {
      type: "APPLY_READING_GUIDE_RESPONSE",
      text,
      provider: request.provider,
      pageKey: request.pageKey,
    });
    if (!isCurrent()) return;
    if (!result?.success) throw new Error(result?.error || t("sidebar.parseError"));
    if (els.response.value.trim() === text.trim()) els.response.value = "";
    if (result.shared) {
      // 別人分享的筆記 Markdown：內容已寫入 storage，整頁重讀最單純。
      state.applying = false;
      await loadActiveTab();
      setStatus(result.summary);
      return;
    }
    const refreshed = await sendToTab(request.tabId, { type: "GET_PAGE_HIGHLIGHTS" }).catch(() => null);
    if (!isCurrent()) return;
    if (refreshed?.success) state.page = refreshed.data;
    state.context = result.context || state.context;
    state.guideTerms = result.terms || [];
    state.guideHighlights = result.highlights || [];
    els.aiControls.open = false;
    render();
    setStatus(guideStatus(result));
  } catch (error) {
    if (!isCurrent()) return;
    setStatus(error?.message || t("sidebar.parseError"), true);
  } finally {
    state.applying = false;
    render();
  }
};

const generateGuide = async () => {
  if (!state.tabId || state.applying) return;
  const request = { tabId: state.tabId, pageKey: state.pageKey, loadToken: state.loadToken, provider: state.provider };
  const isCurrent = () => request.tabId === state.tabId && request.pageKey === state.pageKey && request.loadToken === state.loadToken;
  state.applying = true;
  render();
  setStatus(t("sidebar.generating"));
  try {
    const result = await sendToTab(request.tabId, { type: "GENERATE_READING_GUIDE", pageKey: request.pageKey });
    if (!isCurrent()) return;
    if (result?.manual) {
      state.provider = result.provider || "chatgpt";
      render();
      setStatus(t("sidebar.manualPasteHint"));
      return;
    }
    if (!result?.success) throw new Error(result?.error || t("sidebar.generateError"));
    const refreshed = await sendToTab(request.tabId, { type: "GET_PAGE_HIGHLIGHTS" }).catch(() => null);
    if (!isCurrent()) return;
    if (refreshed?.success) state.page = refreshed.data;
    state.context = result.context || state.context;
    state.guideTerms = result.terms || [];
    state.guideHighlights = result.highlights || [];
    els.aiControls.open = false;
    render();
    setStatus(guideStatus(result));
  } catch (error) {
    if (!isCurrent()) return;
    setStatus(error?.message || t("sidebar.generateError"), true);
  } finally {
    state.applying = false;
    render();
  }
};

// 按下當下才重組 Prompt，新加的標註與目前選取的文字才會算進去。
const copyPromptToClipboard = async () => {
  const info = await sendToTab(state.tabId, { type: "BUILD_READING_GUIDE_PROMPT" }).catch(() => null);
  if (info?.success && info.prompt) state.prompt = info.prompt;
  if (!state.pageKey || !state.prompt) throw new Error(t("sidebar.promptCopyError"));
  await navigator.clipboard.writeText(state.prompt);
  await chrome.storage.session.set({ hkReadingGuideSourceTabId: state.tabId });
};

/* ── 刪除這篇文章的筆記 ─────────────────────────────── */
// 不可復原，所以要按第二次才執行；不用 confirm()，對話框會蓋住正在讀的文章。
let deleteArmTimer = 0;
const disarmDelete = () => {
  window.clearTimeout(deleteArmTimer);
  deleteArmTimer = 0;
  els.deletePageNotes.classList.remove("is-armed");
  els.deletePageNotes.textContent = t("sidebar.deletePageNotes");
};
const deletePageNotes = async () => {
  if (!deleteArmTimer) {
    els.deletePageNotes.classList.add("is-armed");
    els.deletePageNotes.textContent = t("sidebar.deletePageNotesConfirm");
    deleteArmTimer = window.setTimeout(disarmDelete, 5000);
    return;
  }
  disarmDelete();
  // 待存的個人筆記要先取消，否則刪完又被自動儲存寫回來。
  window.clearTimeout(state.noteTimer);
  state.noteTimer = 0;
  state.noteDirty = false;
  try {
    const result = await sendToTab(state.tabId, { type: "DELETE_PAGE_NOTES", pageKey: state.pageKey });
    if (!result?.success) throw new Error(result?.error);
    els.response.value = "";
    await loadActiveTab();
    els.aiControls.open = true;
    els.moreTools.open = false;
    setStatus(t("sidebar.pageNotesDeleted"));
  } catch (error) {
    setStatus(error?.message || t("sidebar.deleteError"), true);
  }
};

/* ── 分享與匯出 ─────────────────────────────────────── */
// 讀出本頁的匯出格式資料 {url,title,tags,entries,note,mindmap}。
const collectPageEntry = async () => {
  const key = state.pageKey;
  if (!key) throw new Error(t("popup.errNoTabSimple"));
  const stored = await chrome.storage.local.get([
    key, PAGE_META_KEY, GENERATED_NOTES_KEY, MINDMAP_KEY, PERSONAL_NOTES_KEY,
  ]);
  const entries = Array.isArray(stored[key]) ? stored[key] : [];
  const meta = stored[PAGE_META_KEY]?.[key] || {};
  const note = stored[GENERATED_NOTES_KEY]?.[key] || null;
  const mindmap = stored[MINDMAP_KEY]?.[key] || null;
  const personal = stored[PERSONAL_NOTES_KEY]?.[key]?.text;
  if (!entries.length && !note && !mindmap && !personal) throw new Error(t("popup.errNoNotes"));
  return {
    url: key,
    title: meta.title || state.page?.title || "",
    tags: Array.isArray(meta.tags) ? meta.tags : [],
    entries,
    note,
    mindmap,
    ...(typeof personal === "string" ? { personalNote: personal } : {}),
  };
};

// 預設把筆記壓進「原文網址#hk=…」；放不下時改 commit 到 GitHub 備份 repo 分享 raw 連結。
const shareLink = async () => {
  els.shareLink.disabled = true;
  try {
    const pageEntry = await collectPageEntry();
    const fragmentLink = await window.HkShareLink.buildFragmentShareUrl(pageEntry);
    if (fragmentLink) {
      await navigator.clipboard.writeText(fragmentLink);
      setStatus(t("popup.statusLinkCopied"));
      return;
    }
    const stored = await chrome.storage.local.get(GITHUB_SETTINGS_KEY);
    const settings = stored[GITHUB_SETTINGS_KEY] || {};
    if (window.HkShareLink.validateGithubSettings(settings)) throw new Error(t("popup.errLinkTooLong"));
    setStatus(t("popup.statusLinkUploading"));
    const rawLink = await window.HkShareLink.commitPageToGithub(settings, pageEntry);
    await navigator.clipboard.writeText(rawLink);
    setStatus(t("popup.statusLinkCopiedGithub"));
  } catch (error) {
    setStatus(error?.message || t("popup.errLinkShare"), true);
  } finally {
    els.shareLink.disabled = false;
  }
};

const copyNotesMarkdown = async () => {
  try {
    const markdown = window.HkParsers.pageToMarkdown(await collectPageEntry(), {
      tags: t("notesMd.tags"),
      highlights: t("notesMd.highlights"),
      summary: t("notesMd.summary"),
      context: t("notesMd.context"),
      personalNote: t("notesMd.personalNote"),
      mindmap: t("notesMd.mindmap"),
    });
    await navigator.clipboard.writeText(markdown);
    setStatus(t("popup.statusNotesCopied"));
  } catch (error) {
    setStatus(error?.message || t("popup.errNotesCopy"), true);
  }
};

const downloadNotesJson = async () => {
  try {
    const pageEntry = await collectPageEntry();
    const payload = { type: "highlight-keeper-bulk", version: 2, exportedAt: Date.now(), pages: [pageEntry] };
    const name = (pageEntry.title || pageEntry.url)
      .replace(/[\/\\:*?"<>|]+/g, "-")
      .replace(/\s+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 60) || "notes";
    const href = URL.createObjectURL(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" }));
    const link = document.createElement("a");
    link.href = href;
    link.download = `${name}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(href);
    setStatus(t("popup.statusNotesDownloaded"));
  } catch (error) {
    setStatus(error?.message || t("popup.errNotesDownload"), true);
  }
};

document.querySelectorAll("[role=tab]").forEach((tab) => {
  tab.addEventListener("click", () => activateTab(tab.dataset.tab));
  tab.addEventListener("keydown", (event) => {
    if (!["ArrowLeft", "ArrowRight"].includes(event.key)) return;
    event.preventDefault();
    const tabs = [...document.querySelectorAll("[role=tab]")];
    const index = tabs.indexOf(tab);
    const next = tabs[(index + (event.key === "ArrowRight" ? 1 : tabs.length - 1)) % tabs.length];
    next.focus();
    activateTab(next.dataset.tab);
  });
});

els.refresh.addEventListener("click", loadActiveTab);
els.generate.addEventListener("click", generateGuide);
els.copyPrompt.addEventListener("click", async () => {
  try {
    await copyPromptToClipboard();
    setStatus(t("sidebar.promptCopied"));
  } catch (_error) {
    setStatus(t("sidebar.promptCopyError"), true);
  }
});
// 只開網站；複製 Prompt 只有一個入口。
els.openWebAI.addEventListener("click", () => {
  const url = state.provider === "gemini-web" ? "https://gemini.google.com/app" : "https://chatgpt.com/";
  chrome.tabs.create({ url }).catch(() => setStatus(t("sidebar.noMatch"), true));
});
els.applyResponse.addEventListener("click", async () => {
  let text = els.response.value.trim();
  // 框內沒東西就直接讀剪貼簿，按一下即可帶入。
  if (!text) {
    text = (await navigator.clipboard.readText().catch(() => ""))?.trim() || "";
    els.response.value = text;
  }
  if (!text) {
    els.response.focus();
    setStatus(t("sidebar.pasteEmpty"), true);
    return;
  }
  applyGuideResponse(text);
});
// 貼上即自動套用（setTimeout 讓 textarea 先吃到剪貼簿內容）。
els.response.addEventListener("paste", () => {
  window.setTimeout(() => {
    const text = els.response.value.trim();
    if (text) applyGuideResponse(text);
  }, 30);
});
els.settings.addEventListener("click", async () => {
  if (!state.tabId) return;
  try {
    await sendToTab(state.tabId, { type: "OPEN_AI_SETTINGS" });
    setStatus("");
  } catch (_error) {
    setStatus(t("sidebar.settingsError"), true);
  }
});
els.startGuide.addEventListener("click", async () => {
  try {
    const result = await sendToTab(state.tabId, { type: "START_GUIDED_READING" });
    setStatus(result?.success ? "" : t("sidebar.noGuideSteps"), !result?.success);
  } catch (_error) {
    setStatus(t("sidebar.noGuideSteps"), true);
  }
});
els.deletePageNotes.addEventListener("click", deletePageNotes);
els.shareLink.addEventListener("click", shareLink);
els.copyNotes.addEventListener("click", copyNotesMarkdown);
els.downloadNotes.addEventListener("click", downloadNotesJson);
els.openTools.addEventListener("click", async () => {
  try {
    await sendToTab(state.tabId, { type: "OPEN_PAGE_PANEL" });
  } catch (_error) {
    setStatus(t("sidebar.settingsError"), true);
  }
});
const openManager = () => {
  chrome.tabs.create({ url: chrome.runtime.getURL("manager.html") });
};
els.openManager.addEventListener("click", openManager);
els.syncNotice.addEventListener("click", openManager);
els.langSelect.addEventListener("change", (event) => window.HkI18n.setLang(event.target.value));
els.personalNote.addEventListener("input", () => {
  state.noteDirty = true;
  state.noteText = els.personalNote.value;
  setNoteStatus("sidebar.notesUnsaved");
  if (state.noteTimer) window.clearTimeout(state.noteTimer);
  state.noteTimer = window.setTimeout(flushNote, 450);
});

chrome.runtime?.onMessage?.addListener((message) => {
  if (message?.type === "HK_PAGE_CHANGED" && message.pageKey === state.pageKey) loadActiveTab();
});
chrome.tabs.onActivated.addListener(() => loadActiveTab());
chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (tabId !== state.tabId) return;
  if (changeInfo.status === "complete" || changeInfo.url) loadActiveTab();
});
chrome.storage.onChanged.addListener((changes, area) => {
  if (area === "local" && changes[SYNC_STATE_KEY]) renderSyncNotice(changes[SYNC_STATE_KEY].newValue);
  if (area !== "local" || !state.pageKey) return;
  if (changes[AI_SETTINGS_KEY]) {
    state.provider = changes[AI_SETTINGS_KEY].newValue?.provider || "chatgpt";
    render();
  }
  if (changes[MINDMAP_KEY]) {
    state.mindmapOutline = changes[MINDMAP_KEY].newValue?.[state.pageKey]?.outline || "";
    render();
  }
  if (state.applying && (changes[GENERATED_NOTES_KEY] || changes[state.pageKey])) return;
  if (changes[GENERATED_NOTES_KEY]) {
    const next = changes[GENERATED_NOTES_KEY].newValue || {};
    state.context = next[state.pageKey]?.context || next[state.pageKey]?.note || "";
    state.guideTerms = next[state.pageKey]?.guideTerms || next[state.pageKey]?.unlocatedTerms || [];
    state.guideHighlights = next[state.pageKey]?.guideHighlights || next[state.pageKey]?.unlocatedHighlights || [];
    render();
  }
  if (changes[PERSONAL_NOTES_KEY] && !state.noteDirty) {
    const next = changes[PERSONAL_NOTES_KEY].newValue || {};
    state.noteText = next[state.pageKey]?.text || "";
    els.personalNote.value = state.noteText;
    setNoteStatus("sidebar.notesSaved");
  }
  if (changes[state.pageKey]) loadActiveTab();
});

window.HkI18n.initI18n().then(() => {
  window.HkI18n.applyDOMTranslations();
  els.langSelect.value = window.HkI18n.getLang?.() || "zh";
  render();
  chrome.storage.local.get(SYNC_STATE_KEY).then((stored) => renderSyncNotice(stored[SYNC_STATE_KEY]));
});
window.HkI18n.onLangChange(() => {
  window.HkI18n.applyDOMTranslations();
  render();
});
activateTab("context");
loadActiveTab();
