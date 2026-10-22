const PERSONAL_NOTES_KEY = "hkPersonalNotes";
const GENERATED_NOTES_KEY = "hkGeneratedNotes";

const els = {
  title: document.getElementById("pageTitle"),
  url: document.getElementById("pageUrl"),
  status: document.getElementById("status"),
  aiControls: document.getElementById("aiControls"),
  noPage: document.getElementById("noPage"),
  context: document.getElementById("contextText"),
  emptyContext: document.getElementById("emptyContext"),
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
  manualPaste: document.getElementById("manualPaste"),
  response: document.getElementById("aiResponse"),
  applyResponse: document.getElementById("applyResponseBtn"),
  settings: document.getElementById("settingsBtn"),
  refresh: document.getElementById("refreshBtn"),
  tabs: document.querySelector(".tabs"),
};

const t = (key, params) => window.HkI18n.t(key, params);
const state = {
  tabId: null,
  pageKey: "",
  page: null,
  context: "",
  guideTerms: [],
  guideHighlights: [],
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
  const stored = await chrome.storage.local.get(PERSONAL_NOTES_KEY);
  const notes = { ...(stored[PERSONAL_NOTES_KEY] || {}) };
  notes[pageKey] = { text, updatedAt: Date.now() };
  await chrome.storage.local.set({ [PERSONAL_NOTES_KEY]: notes });
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

const makeItemButton = (item, kind) => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "item-button";
  if (item.unavailable) {
    button.disabled = true;
    button.classList.add("is-unavailable");
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
  const kindLabel = document.createElement("span");
  kindLabel.className = "item-kind";
  kindLabel.textContent = t(kind === "term" ? "sidebar.termKind" : "sidebar.passageKind");
  if (item.unavailable) kindLabel.textContent += ` · ${t("sidebar.unavailable")}`;
  const text = document.createElement("span");
  text.className = "item-text";
  text.textContent = item.text || "";
  button.append(kindLabel, text);
  const note = item.explanation || item.reason || item.note || "";
  if (note) {
    const description = document.createElement("span");
    description.className = "item-note";
    description.textContent = note;
    button.appendChild(description);
  }
  return button;
};

const renderItems = (container, empty, items, kind) => {
  container.replaceChildren();
  items.forEach((item) => container.appendChild(makeItemButton(item, kind)));
  empty.hidden = items.length > 0;
};

const render = () => {
  const page = state.page;
  els.title.textContent = page?.title || "—";
  els.url.textContent = page?.url || "";
  els.context.textContent = state.context || "";
  els.emptyContext.hidden = Boolean(state.context);

  const storedHighlights = Array.isArray(page?.highlights) ? page.highlights : [];
  const visibleHighlights = [...storedHighlights, ...(page?.previewHighlights || [])];
  const mergeItems = (stored, extra) => {
    const items = stored.map((item) => ({ ...item, reason: item.note }));
    const indexes = new Map(items.map((item, index) => [(item.text || "").trim().toLowerCase(), index]));
    extra.forEach((item) => {
      const key = (item.text || "").trim().toLowerCase();
      if (key && indexes.has(key)) {
        const index = indexes.get(key);
        items[index] = { ...items[index], ...item, unavailable: false };
      } else if (key) {
        indexes.set(key, items.length);
        items.push({ ...item, unavailable: true });
      }
    });
    return items;
  };
  const terms = mergeItems(visibleHighlights.filter((item) => item.kind === "term"), state.guideTerms);
  const highlights = mergeItems(visibleHighlights.filter((item) => item.kind !== "term"), state.guideHighlights);
  renderItems(els.terms, els.emptyTerms, terms, "term");
  renderItems(els.highlights, els.emptyHighlights, highlights, "passage");
  els.termsCount.textContent = String(terms.length);
  els.termsCount.hidden = terms.length === 0;
  els.highlightsCount.textContent = String(highlights.length);
  els.highlightsCount.hidden = highlights.length === 0;

  const isWebProvider = ["chatgpt", "gemini-web"].includes(state.provider);
  els.generate.hidden = isWebProvider;
  els.copyPrompt.hidden = !isWebProvider;
  els.openWebAI.hidden = !isWebProvider;
  els.openWebAI.textContent = t(state.provider === "gemini-web" ? "sidebar.openGemini" : "sidebar.openChatgpt");
  els.applyResponse.disabled = state.applying;
  els.generate.disabled = state.applying;
  els.manualPaste.hidden = !isWebProvider;
  els.settings.hidden = !state.pageKey;
  els.noPage.hidden = Boolean(page);
  els.aiControls.hidden = !page;
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
  if (token !== state.loadToken) return;
  if (state.tabId !== tab?.id) els.response.value = "";
  state.tabId = tab?.id ?? null;
  state.pageKey = "";
  state.page = null;
  state.context = "";
  state.guideTerms = [];
  state.guideHighlights = [];
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
      chrome.storage.local.get([GENERATED_NOTES_KEY, PERSONAL_NOTES_KEY]),
      sendToTab(tab.id, { type: "BUILD_READING_GUIDE_PROMPT" }),
    ]);
    if (token !== state.loadToken) return;
    const generated = stored[GENERATED_NOTES_KEY]?.[state.pageKey] || {};
    const personal = stored[PERSONAL_NOTES_KEY]?.[state.pageKey] || {};
    state.context = generated.context || generated.note || "";
    state.guideTerms = generated.guideTerms || generated.unlocatedTerms || [];
    state.guideHighlights = generated.guideHighlights || generated.unlocatedHighlights || [];
    state.noteText = personal.text || "";
    els.personalNote.value = state.noteText;
    state.provider = promptInfo?.provider || "chatgpt";
    state.configured = Boolean(promptInfo?.configured);
    state.prompt = promptInfo?.prompt || "";
    if (!promptInfo?.success) setStatus(t("sidebar.extractError"), true);
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
      prompt: state.prompt,
      provider: request.provider,
      pageKey: request.pageKey,
    });
    if (!isCurrent()) return;
    if (!result?.success) throw new Error(result?.error || t("sidebar.parseError"));
    const refreshed = await sendToTab(request.tabId, { type: "GET_PAGE_HIGHLIGHTS" }).catch(() => null);
    if (!isCurrent()) return;
    if (refreshed?.success) state.page = refreshed.data;
    state.context = result.context || state.context;
    state.guideTerms = result.terms || [];
    state.guideHighlights = result.highlights || [];
    if (els.response.value.trim() === text.trim()) els.response.value = "";
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
    await navigator.clipboard.writeText(state.prompt);
    await chrome.storage.session.set({ hkReadingGuideSourceTabId: state.tabId });
    setStatus(t("sidebar.promptCopied"));
  } catch (_error) {
    setStatus(t("sidebar.promptCopyError"), true);
  }
});
els.openWebAI.addEventListener("click", async () => {
  try {
    if (!state.pageKey || !state.prompt) return;
    await navigator.clipboard.writeText(state.prompt);
    await chrome.storage.session.set({ hkReadingGuideSourceTabId: state.tabId });
    const url = state.provider === "gemini-web" ? "https://gemini.google.com/app" : "https://chatgpt.com/";
    await chrome.tabs.create({ url });
    setStatus(t("sidebar.promptCopied"));
  } catch (_error) {
    setStatus(t("sidebar.promptCopyError"), true);
  }
});
els.applyResponse.addEventListener("click", () => applyGuideResponse(els.response.value.trim()));
els.settings.addEventListener("click", async () => {
  if (!state.tabId) return;
  try {
    await sendToTab(state.tabId, { type: "OPEN_AI_SETTINGS" });
    setStatus("");
  } catch (_error) {
    setStatus(t("sidebar.settingsError"), true);
  }
});
els.personalNote.addEventListener("input", () => {
  state.noteDirty = true;
  state.noteText = els.personalNote.value;
  setNoteStatus("sidebar.notesUnsaved");
  if (state.noteTimer) window.clearTimeout(state.noteTimer);
  state.noteTimer = window.setTimeout(flushNote, 450);
});

chrome.tabs.onActivated.addListener(() => loadActiveTab());
chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (tabId !== state.tabId) return;
  if (changeInfo.status === "complete" || changeInfo.url) loadActiveTab();
});
chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== "local" || !state.pageKey) return;
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
  render();
});
window.HkI18n.onLangChange(() => {
  window.HkI18n.applyDOMTranslations();
  render();
});
activateTab("context");
loadActiveTab();
