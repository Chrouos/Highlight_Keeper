const colorInput = document.getElementById("color");
const addColorBtn = document.getElementById("addColorBtn");
const paletteListEl = document.getElementById("paletteList");
const openPanelBtn = document.getElementById("openPanelBtn");
const openManagerBtn = document.getElementById("openManagerBtn");
const statusEl = document.querySelector(".status");

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
      files: ["contentScript.js"],
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
      return chrome.tabs.sendMessage(tabId, payload);
    }
    throw new Error(message || "無法傳送訊息");
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
    placeholder.textContent = "目前沒有顏色，請新增一個色票。";
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
    applyBtn.title = `使用 ${color}`;
    applyBtn.addEventListener("click", () => selectColor(color));

    const editInput = document.createElement("input");
    editInput.type = "color";
    editInput.className = "palette-edit";
    editInput.value = color;
    editInput.title = `編輯顏色 ${color}`;
    editInput.addEventListener("input", (event) => {
      updatePaletteColor(index, event.target.value);
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "palette-remove";
    deleteBtn.textContent = "刪除";
    deleteBtn.title = `刪除顏色 ${color}`;
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
    setStatus(`已選擇顏色 ${normalized}`);
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
  setStatus(`色票顏色已更新為 ${normalized}`);
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
  setStatus(`已刪除顏色 ${removed}`);
};

const handleAddColor = async () => {
  const color = normalizeColor(colorInput.value);
  if (palette.includes(color)) {
    setStatus("此顏色已存在於色票中");
    return;
  }
  await persistPalette([...palette, color]);
  setStatus(`已新增顏色 ${color}`);
};

const loadInitialState = async () => {
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
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) {
    setStatus("找不到目前的分頁，無法開啟面板。", true);
    return;
  }
  try {
    const response = await sendMessageToTab(tab.id, {
      type: "OPEN_PAGE_PANEL",
      side: panelSide,
    });
    if (!response?.success) {
      throw new Error(response?.error ?? "無法開啟頁面面板");
    }
    setStatus("已在頁面顯示面板");
  } catch (error) {
    setStatus(error?.message || "無法開啟頁面面板。", true);
  }
});

openManagerBtn?.addEventListener("click", () => {
  const url = chrome.runtime.getURL("manager.html");
  chrome.tabs.create({ url });
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
});

loadInitialState();
