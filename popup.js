const colorInput = document.getElementById("color");
const addColorBtn = document.getElementById("addColorBtn");
const paletteListEl = document.getElementById("paletteList");
const openPanelBtn = document.getElementById("openPanelBtn");
const openManagerBtn = document.getElementById("openManagerBtn");
const shareLinkBtn = document.getElementById("shareLinkBtn");
const copyNotesBtn = document.getElementById("copyNotesBtn");
const downloadNotesBtn = document.getElementById("downloadNotesBtn");
const langSelect = document.getElementById("langSelect");
const statusEl = document.querySelector(".status");

const t = (key, params) => HkI18n.t(key, params);

// 網址正規化共用自 shared.js（manifest/popup.html 已先載入）。
const normalizePageKey = (href) =>
  window.HkUrlKey ? window.HkUrlKey.normalizePageKey(href) : href;

// 與 manager.js / contentScript 一致的 storage 主鍵，用來組出「當頁」的匯出資料。
const PAGE_META_KEY = "__hk_page_meta__";
const GENERATED_NOTES_KEY = "hkGeneratedNotes";
const MINDMAP_KEY = "hkMindmaps";
const GITHUB_SETTINGS_KEY = "hkGithubSyncSettings";

const DEFAULT_COLOR = "#ffeb3b";
const DEFAULT_PALETTE = [
  "#ffeb3b",
  "#ffa726",
  "#81c784",
  "#64b5f6",
  "#f48fb1",
  "#c792ea",
];

let palette = [...DEFAULT_PALETTE];
let isInitializing = true;
let panelSide = "right";
const RECEIVER_MISSING_ERROR = "Receiving end does not exist.";

const normalizeColor = (value) => {
  if (typeof value !== "string") return DEFAULT_COLOR;
  const trimmed = value.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(trimmed)) return trimmed;
  if (/^#[0-9a-f]{3}$/.test(trimmed)) {
    return `#${trimmed[1]}${trimmed[1]}${trimmed[2]}${trimmed[2]}${trimmed[3]}${trimmed[3]}`;
  }
  const match = trimmed.match(
    /^rgba?\((\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})(?:,\s*(\d+(?:\.\d+)?))?\)$/
  );
  if (match) {
    const toHex = (component) =>
      Math.max(0, Math.min(255, Number(component)))
        .toString(16)
        .padStart(2, "0");
    return `#${toHex(match[1])}${toHex(match[2])}${toHex(match[3])}`;
  }
  return DEFAULT_COLOR;
};

const normalizePalette = (input) => {
  if (!Array.isArray(input)) return [...DEFAULT_PALETTE];
  const seen = new Set();
  const result = [];
  input.forEach((value) => {
    if (typeof value !== "string") return;
    const color = normalizeColor(value);
    if (!seen.has(color)) {
      seen.add(color);
      result.push(color);
    }
  });
  return result.length ? result : [...DEFAULT_PALETTE];
};

const setStatus = (message, isError = false) => {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "#d93025" : "#1a73e8";
};

// popup 很長、狀態列在最底部——回饋常被推到看不見。分享等操作用這個把它捲進視野。
const setStatusVisible = (message, isError = false) => {
  setStatus(message, isError);
  try {
    statusEl.scrollIntoView({ block: "center", behavior: "smooth" });
  } catch (_e) {}
};

const injectContentAssets = async (tabId) => {
  try {
    await chrome.scripting?.insertCSS({
      target: { tabId },
      files: ["contentStyles.css"],
    });
  } catch (error) {
    console.debug("注入面板樣式失敗", error);
  }
  try {
    await chrome.scripting?.executeScript({
      target: { tabId },
      files: ["shared.js", "parsers.js", "i18n.js", "contentScript.js"],
    });
  } catch (error) {
    console.debug("注入內容腳本失敗", error);
    throw error;
  }
};

const sendMessageToTab = async (tabId, payload) => {
  try {
    return await chrome.tabs.sendMessage(tabId, payload);
  } catch (error) {
    const message =
      chrome.runtime.lastError?.message || error?.message || "";
    if (message.includes(RECEIVER_MISSING_ERROR)) {
      await injectContentAssets(tabId);
      await new Promise((resolve) => setTimeout(resolve, 150));
      return chrome.tabs.sendMessage(tabId, payload);
    }
    throw new Error(message || t("popup.errCannotSendMessage"));
  }
};

const persistColor = async (color) => {
  try {
    await chrome.storage?.local.set({ hkLastColor: color });
  } catch (error) {
    console.debug("儲存顏色失敗", error);
  }
};

const persistPalette = async (nextPalette) => {
  palette = normalizePalette(nextPalette);
  try {
    await chrome.storage?.local.set({ hkColorPalette: palette });
  } catch (error) {
    console.debug("儲存顏色列表失敗", error);
  }
  renderPalette();
  return palette;
};

const renderPalette = () => {
  paletteListEl.innerHTML = "";
  if (!palette.length) {
    const placeholder = document.createElement("p");
    placeholder.className = "palette-empty";
    placeholder.textContent = t("popup.emptyPalette");
    paletteListEl.appendChild(placeholder);
    return;
  }

  palette.forEach((color, index) => {
    const item = document.createElement("div");
    item.className = "palette-item";
    item.dataset.index = String(index);
    if (color === colorInput.value.toLowerCase()) {
      item.classList.add("is-active");
    }

    const applyBtn = document.createElement("button");
    applyBtn.type = "button";
    applyBtn.className = "palette-swatch";
    applyBtn.style.backgroundColor = color;
    applyBtn.title = t("popup.useColor", { color });
    applyBtn.addEventListener("click", () => selectColor(color));

    const editInput = document.createElement("input");
    editInput.type = "color";
    editInput.className = "palette-edit";
    editInput.value = color;
    editInput.title = t("popup.editColor", { color });
    editInput.addEventListener("input", (event) => {
      updatePaletteColor(index, event.target.value);
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "palette-remove";
    deleteBtn.textContent = t("popup.deleteColor");
    deleteBtn.title = t("popup.deleteColorTitle", { color });
    deleteBtn.addEventListener("click", () => removePaletteColor(index));

    item.appendChild(applyBtn);
    item.appendChild(editInput);
    item.appendChild(deleteBtn);
    paletteListEl.appendChild(item);
  });
};

const selectColor = async (color, silent = false) => {
  const normalized = normalizeColor(color);
  if (colorInput.value.toLowerCase() !== normalized) {
    colorInput.value = normalized;
  }
  await persistColor(normalized);
  if (!silent) {
    setStatus(t("popup.statusSelected", { color: normalized }));
  }
  renderPalette();
};

const updatePaletteColor = async (index, color) => {
  const previousColor = palette[index];
  const normalized = normalizeColor(color);
  const nextPalette = [...palette];
  nextPalette[index] = normalized;
  await persistPalette(nextPalette);
  if (colorInput.value.toLowerCase() === previousColor) {
    await selectColor(normalized, true);
  }
  setStatus(t("popup.statusUpdated", { color: normalized }));
  window.requestAnimationFrame(() => {
    const nextInput = paletteListEl.querySelector(
      `.palette-item[data-index="${index}"] .palette-edit`
    );
    if (nextInput instanceof HTMLInputElement) {
      nextInput.focus();
    }
  });
};

const removePaletteColor = async (index) => {
  const removed = palette[index];
  const nextPalette = palette.filter((_, i) => i !== index);
  await persistPalette(nextPalette);
  const current = colorInput.value.toLowerCase();
  if (!palette.includes(current)) {
    await selectColor(palette[0] ?? DEFAULT_COLOR, true);
  }
  setStatus(t("popup.statusDeleted", { color: removed }));
};

const handleAddColor = async () => {
  const color = normalizeColor(colorInput.value);
  if (palette.includes(color)) {
    setStatus(t("popup.statusDuplicate"));
    return;
  }
  await persistPalette([...palette, color]);
  setStatus(t("popup.statusAdded", { color }));
};

const aiModeLabel = (provider) => {
  switch (provider) {
    case "gemini":
      return t("popup.aiModeGemini");
    case "chatgpt":
      return t("popup.aiModeCopy");
    case "openai":
    default:
      return t("popup.aiModeOpenai");
  }
};

// 找出「當頁」的 storage 主鍵：contentScript 以 canonical 連結為主鍵，popup 不在
// 頁面內，先輕量讀取該頁 canonical（不觸發全文擷取），讓鍵與標註鍵一致。
const resolveCurrentPageKey = async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.url) return { tab: tab || null, key: "" };
  let canonical = "";
  try {
    const [{ result } = {}] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => document.querySelector('link[rel="canonical"]')?.href || "",
    });
    if (typeof result === "string") canonical = result;
  } catch (_e) {
    // 受限頁面（chrome://、商店等）無法注入；退回網址。
  }
  return { tab, key: normalizePageKey(canonical || tab.url) };
};

const loadPageInfo = async () => {
  const markEl = document.getElementById("markCount");
  const modeEl = document.getElementById("aiMode");
  try {
    const stored = await chrome.storage?.local.get("hkAISettings");
    const provider = stored?.hkAISettings?.provider || "openai";
    if (modeEl) modeEl.textContent = aiModeLabel(provider);
  } catch (_e) {
    if (modeEl) modeEl.textContent = "—";
  }
  try {
    // Read the count straight from storage — waking the content script for
    // this forced a full-page text extraction and made the popup feel slow.
    const { key } = await resolveCurrentPageKey();
    if (!key) {
      if (markEl) markEl.textContent = "—";
      return;
    }
    const stored = await chrome.storage?.local.get(key);
    const entries = stored?.[key];
    if (markEl) {
      markEl.textContent = Array.isArray(entries) ? String(entries.length) : "0";
    }
  } catch (_e) {
    if (markEl) markEl.textContent = "—";
  }
};

// 讀出「當頁」的匯出格式資料 {url,title,tags,entries,note,mindmap}；沒筆記回 null。
const collectCurrentPageEntry = async () => {
  const { key } = await resolveCurrentPageKey();
  if (!key) return { error: t("popup.errNoTabSimple") };
  const stored = await chrome.storage?.local.get([
    key,
    PAGE_META_KEY,
    GENERATED_NOTES_KEY,
    MINDMAP_KEY,
  ]);
  const entries = Array.isArray(stored?.[key]) ? stored[key] : [];
  const meta = stored?.[PAGE_META_KEY]?.[key] || {};
  const note = stored?.[GENERATED_NOTES_KEY]?.[key] || null;
  const mindmap = stored?.[MINDMAP_KEY]?.[key] || null;
  if (!entries.length && !note && !mindmap) {
    return { error: t("popup.errNoNotes") };
  }
  return {
    pageEntry: {
      url: key,
      title: meta.title || "",
      tags: Array.isArray(meta.tags) ? meta.tags : [],
      entries,
      note,
      mindmap,
    },
  };
};

// 複製分享連結：預設把筆記壓進「原文網址#hk=…」（零設定）；
// 筆記太多放不下時，改 commit 到 GitHub 備份 repo 分享 raw 連結（免 Pages）。
const shareCurrentPageLink = async () => {
  if (shareLinkBtn) shareLinkBtn.disabled = true;
  try {
    const { pageEntry, error } = await collectCurrentPageEntry();
    if (error) {
      setStatusVisible(error, true);
      return;
    }
    const fragmentLink = await window.HkShareLink.buildFragmentShareUrl(pageEntry);
    if (fragmentLink) {
      await navigator.clipboard.writeText(fragmentLink);
      setStatusVisible(t("popup.statusLinkCopied"));
      return;
    }
    // fragment 放不下 → 走 GitHub raw（沿用備份設定，不需開 Pages）
    const stored = await chrome.storage?.local.get(GITHUB_SETTINGS_KEY);
    const settings = stored?.[GITHUB_SETTINGS_KEY] || {};
    if (window.HkShareLink.validateGithubSettings(settings)) {
      setStatusVisible(t("popup.errLinkTooLong"), true);
      return;
    }
    setStatusVisible(t("popup.statusLinkUploading"));
    const rawLink = await window.HkShareLink.commitPageToGithub(settings, pageEntry);
    await navigator.clipboard.writeText(rawLink);
    setStatusVisible(t("popup.statusLinkCopiedGithub"));
  } catch (error) {
    console.debug("複製分享連結失敗", error);
    setStatusVisible(error?.message || t("popup.errLinkShare"), true);
  } finally {
    if (shareLinkBtn) shareLinkBtn.disabled = false;
  }
};

// 複製本頁筆記成 Markdown（貼到聊天／筆記軟體都可讀）。
const copyCurrentPageNotes = async () => {
  try {
    const { pageEntry, error } = await collectCurrentPageEntry();
    if (error) {
      setStatusVisible(error, true);
      return;
    }
    const markdown = window.HkParsers.pageToMarkdown(pageEntry, {
      tags: t("notesMd.tags"),
      highlights: t("notesMd.highlights"),
      summary: t("notesMd.summary"),
      mindmap: t("notesMd.mindmap"),
    });
    await navigator.clipboard.writeText(markdown);
    setStatusVisible(t("popup.statusNotesCopied"));
  } catch (error) {
    console.debug("複製本頁筆記失敗", error);
    setStatusVisible(error?.message || t("popup.errNotesCopy"), true);
  }
};

// 下載本頁筆記 JSON（給也用 Highlight Keeper 的人「匯入多個 JSON」）。
const downloadCurrentPageNotes = async () => {
  try {
    const { pageEntry, error } = await collectCurrentPageEntry();
    if (error) {
      setStatusVisible(error, true);
      return;
    }
    const payload = {
      type: "highlight-keeper-bulk",
      version: 2,
      exportedAt: Date.now(),
      pages: [pageEntry],
    };
    const name =
      (pageEntry.title || pageEntry.url)
        .replace(/[\/\\:*?"<>|]+/g, "-")
        .replace(/\s+/g, "-")
        .replace(/^-+|-+$/g, "")
        .slice(0, 60) || "notes";
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: "application/json",
    });
    const href = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = href;
    link.download = `${name}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(href);
    setStatusVisible(t("popup.statusNotesDownloaded"));
  } catch (error) {
    console.debug("下載本頁筆記失敗", error);
    setStatusVisible(error?.message || t("popup.errNotesDownload"), true);
  }
};

const loadInitialState = async () => {
  await HkI18n.initI18n();
  HkI18n.applyDOMTranslations();
  if (langSelect) {
    langSelect.value = HkI18n.getLang();
    langSelect.addEventListener("change", (event) => {
      HkI18n.setLang(event.target.value);
    });
  }
  HkI18n.onLangChange(() => {
    renderPalette();
    loadPageInfo();
    if (langSelect) langSelect.value = HkI18n.getLang();
  });

  try {
    const stored = await chrome.storage?.local.get([
      "hkLastColor",
      "hkColorPalette",
      "hkPanelSide",
    ]);
    palette = normalizePalette(stored?.hkColorPalette);
    renderPalette();
    const initialColor = stored?.hkLastColor
      ? normalizeColor(stored.hkLastColor)
      : palette[0] ?? DEFAULT_COLOR;
    colorInput.value = initialColor;
    await persistColor(initialColor);
    panelSide =
      stored?.hkPanelSide === "left" || stored?.hkPanelSide === "right"
        ? stored.hkPanelSide
        : "right";
  } catch (error) {
    console.debug("載入初始設定失敗", error);
    palette = [...DEFAULT_PALETTE];
    renderPalette();
    colorInput.value = DEFAULT_COLOR;
    panelSide = "right";
  } finally {
    isInitializing = false;
  }
};

colorInput.addEventListener("input", (event) => {
  if (isInitializing) return;
  selectColor(event.target.value);
});

addColorBtn.addEventListener("click", handleAddColor);

openPanelBtn.addEventListener("click", async () => {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) {
      setStatus(t("popup.errNoTab"), true);
      return;
    }
    const response = await sendMessageToTab(tab.id, {
      type: "OPEN_PAGE_PANEL",
      side: panelSide,
    });
    if (!response?.success) {
      throw new Error(response?.error ?? t("popup.errOpenPanel"));
    }
    setStatus(t("popup.statusPanelOpened"));
  } catch (error) {
    setStatus(error?.message || t("popup.errOpenPanel2"), true);
  }
});

openManagerBtn?.addEventListener("click", () => {
  const url = chrome.runtime.getURL("manager.html");
  chrome.tabs.create({ url });
});

shareLinkBtn?.addEventListener("click", shareCurrentPageLink);
copyNotesBtn?.addEventListener("click", copyCurrentPageNotes);
downloadNotesBtn?.addEventListener("click", downloadCurrentPageNotes);

document.getElementById("aiHighlightBtn")?.addEventListener("click", async () => {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) { setStatus(t("popup.errNoTabSimple"), true); return; }
    const response = await sendMessageToTab(tab.id, { type: "BUILD_AI_HIGHLIGHT_PROMPT" });
    if (!response?.success || !response.prompt) {
      throw new Error(response?.error ?? t("popup.errCannotTrigger"));
    }
    await navigator.clipboard.writeText(response.prompt);
    setStatus(t("popup.statusAiCopied"));
  } catch (error) {
    setStatus(error?.message || t("popup.errAiTrigger"), true);
  }
});

const aiPasteInput = document.getElementById("aiPasteInput");

// 套用框內的 AI 回覆（重點／摘要／標籤／心智圖）。按鈕與「貼上即套用」共用。
const applyPastedAiResponse = async () => {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) { setStatus(t("popup.errNoTabSimple"), true); return; }
    let text = aiPasteInput?.value?.trim() || "";
    // 框內沒東西就直接讀剪貼簿（已要求 clipboardRead），按一下即可帶入。
    if (!text) {
      try {
        text = (await navigator.clipboard.readText())?.trim() || "";
        if (aiPasteInput && text) aiPasteInput.value = text;
      } catch (_e) {
        /* 剪貼簿空或被拒 → 維持空字串 */
      }
    }
    if (!text) {
      aiPasteInput?.focus();
      setStatus(t("popup.errPasteEmpty"), true);
      return;
    }
    setStatus(t("popup.statusApplyingPaste"));
    const response = await sendMessageToTab(tab.id, {
      type: "APPLY_AI_RESPONSE",
      text,
    });
    if (!response?.success) {
      throw new Error(response?.error ?? t("popup.errCannotTrigger"));
    }
    setStatus(response.summary || t("popup.statusPasteApplied"));
    if (aiPasteInput) aiPasteInput.value = "";
    loadPageInfo();
  } catch (error) {
    setStatus(error?.message || t("popup.errAiTrigger"), true);
  }
};

document.getElementById("aiPasteApplyBtn")?.addEventListener("click", applyPastedAiResponse);

// 貼上即自動套用（setTimeout 讓 textarea 先吃到剪貼簿內容），與頁面面板一致。
// 註：右鍵選單貼上會走這條；鍵盤 Cmd/Ctrl+V 走下方 keydown（popup 收不到原生快捷鍵）。
aiPasteInput?.addEventListener("paste", () => {
  window.setTimeout(() => {
    if (aiPasteInput.value.trim()) applyPastedAiResponse();
  }, 30);
});

// macOS 的擴充 popup 沒有「編輯」選單，收不到 Cmd+A/C/X/V/Z，這裡用 JS 補回來。
aiPasteInput?.addEventListener("keydown", (event) => {
  if (!(event.metaKey || event.ctrlKey) || event.altKey) return;
  const el = aiPasteInput;
  switch (event.key.toLowerCase()) {
    case "a":
      event.preventDefault();
      el.select();
      break;
    case "c":
      event.preventDefault();
      document.execCommand("copy");
      break;
    case "x":
      event.preventDefault();
      document.execCommand("cut");
      break;
    case "z":
      event.preventDefault();
      document.execCommand(event.shiftKey ? "redo" : "undo");
      break;
    case "v":
      event.preventDefault();
      navigator.clipboard
        .readText()
        .then((clip) => {
          const text = clip || "";
          if (!text) return;
          const start = el.selectionStart ?? el.value.length;
          const end = el.selectionEnd ?? el.value.length;
          el.value = el.value.slice(0, start) + text + el.value.slice(end);
          const caret = start + text.length;
          el.selectionStart = el.selectionEnd = caret;
          if (el.value.trim()) applyPastedAiResponse();
        })
        .catch(() => {
          /* 剪貼簿空或被拒 */
        });
      break;
    default:
      break;
  }
});

document.getElementById("clearPageBtn")?.addEventListener("click", async () => {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) { setStatus(t("popup.errNoTabSimple"), true); return; }
    const response = await sendMessageToTab(tab.id, { type: "CLEAR_PAGE_HIGHLIGHTS" });
    if (!response?.success) throw new Error(response?.error ?? t("popup.errCannotClear"));
    setStatus(t("popup.statusCleared"));
    loadPageInfo();
  } catch (error) {
    setStatus(error?.message || t("popup.errClear"), true);
  }
});

chrome.storage?.onChanged.addListener((changes, areaName) => {
  if (areaName !== "local") return;
  if (changes.hkColorPalette && !isInitializing) {
    palette = normalizePalette(changes.hkColorPalette.newValue);
    renderPalette();
  }
  if (changes.hkLastColor && !isInitializing) {
    const nextColor = normalizeColor(changes.hkLastColor.newValue);
    colorInput.value = nextColor;
    renderPalette();
  }
  if (changes.hkPanelSide) {
    panelSide =
      changes.hkPanelSide.newValue === "left" ? "left" : "right";
  }
  if (changes.hkAISettings) {
    loadPageInfo();
  }
});

loadInitialState();
loadPageInfo();
