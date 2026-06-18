const colorInput = document.getElementById("color");
const addColorBtn = document.getElementById("addColorBtn");
const paletteListEl = document.getElementById("paletteList");
const openPanelBtn = document.getElementById("openPanelBtn");
const openManagerBtn = document.getElementById("openManagerBtn");
const langSelect = document.getElementById("langSelect");
const statusEl = document.querySelector(".status");

const t = (key, params) => HkI18n.t(key, params);

// 網址正規化共用自 shared.js（manifest/popup.html 已先載入）。
const normalizePageKey = (href) =>
  window.HkUrlKey ? window.HkUrlKey.normalizePageKey(href) : href;

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
      files: ["shared.js", "i18n.js", "contentScript.js"],
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
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.url) {
      if (markEl) markEl.textContent = "—";
      return;
    }
    // contentScript 以 canonical 連結為主鍵；popup 不在頁面內，先嘗試讀取該頁
    // canonical（輕量、不觸發全文擷取），讓計數鍵與標註鍵一致，否則會顯示成 0。
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
    const key = normalizePageKey(canonical || tab.url);
    const stored = await chrome.storage?.local.get(key);
    const entries = stored?.[key];
    if (markEl) {
      markEl.textContent = Array.isArray(entries) ? String(entries.length) : "0";
    }
  } catch (_e) {
    if (markEl) markEl.textContent = "—";
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
