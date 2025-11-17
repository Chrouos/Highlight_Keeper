const HIGHLIGHT_CLASS = "hk-highlight";
const HIGHLIGHT_ATTR = "data-highlight-id";
const MAX_SELECTOR_DEPTH = 6;
const TEXT_CONTEXT_CHARS = 60;
const TEXT_PARENT_SKIP_TAGS = new Set(["SCRIPT", "STYLE", "NOSCRIPT"]);

const pageKey = new URL(window.location.href).href;

const storage = chrome.storage?.local;
const DEFAULT_COLOR = "#ffeb3b";
const DEFAULT_PALETTE = [
  "#ffeb3b",
  "#ffa726",
  "#81c784",
  "#64b5f6",
  "#f48fb1",
  "#c792ea",
];
const PAGE_META_KEY = "__hk_page_meta__";
const DEFAULT_AI_PROMPT = `你是一位筆記整理助手，根據提供的網頁全文與標註內容，整理出淺顯易懂的筆記。優先考慮使用者標注段落，將重點控制在五百字以內，輸出內容需要像說故事一樣有脈絡的說明。`;
const DEFAULT_GEMINI_MODEL = "gemini-2.0-flash";
const MODEL_OPTIONS = {
  openai: [
    { value: "gpt-4o-mini", label: "GPT-4o mini" },
    { value: "gpt-4o", label: "GPT-4o" },
  ],
  gemini: [
    { value: "gemini-2.0-flash", label: "Gemini 2.0 Flash" },
    { value: "gemini-2.5-flash", label: "Gemini 2.5 Flash" },
  ],
};
let colorPalette = [...DEFAULT_PALETTE];
let currentColor = DEFAULT_COLOR;
let floatingButton = null;
const FLOATING_BUTTON_ID = "hk-floating-btn";
const FLOATING_BUTTON_MARGIN = 8;
let selectionDebounceTimer = null;
const HIGHLIGHT_MENU_ID = "hk-highlight-menu";
let highlightMenu = null;
let highlightMenuEls = null;
let activeHighlight = null;
let activeHighlightId = null;
let highlightMenuStatusTimer = null;
const HIGHLIGHT_PANEL_ID = "hk-page-panel";
const HIGHLIGHT_PANEL_POSITION_KEY = "hkPanelPosition";
const HIGHLIGHT_PANEL_FONT_SCALE_KEY = "hkPanelFontScale";
const PANEL_FONT_SCALE_MIN = 0.85;
const PANEL_FONT_SCALE_MAX = 1.35;
const PANEL_FONT_SCALE_STEP = 0.1;
const PANEL_DRAG_MARGIN = 12;
let highlightPanel = null;
let highlightPanelEls = null;
let highlightPanelVisible = false;
let highlightPanelPreferredSide = "right";
const highlightPanelState = {
  side: "right",
  activeKey: pageKey,
  searchTerm: "",
  activeTag: null,
  activeTab: "page",
  searchPageFilter: null,
  allPages: {},
  pageMeta: {},
  allTags: [],
  currentEntries: [],
  notesByPage: {},
  position: null,
  fontScale: 1,
};
let panelStatusTimer = null;
let archiveStatusTimer = null;
let panelPreferencesPromise;
let panelDragState = null;
let aiSettings = {
  provider: "openai",
  openaiKey: "",
  openaiModel: "gpt-4o-mini",
  geminiKey: "",
  geminiModel: DEFAULT_GEMINI_MODEL,
  prompt: DEFAULT_AI_PROMPT,
};
let isGeneratingNote = false;
const HIGHLIGHT_RETRY_DELAYS = [450, 1500, 3500];

const ensureFloatingButton = () => {
  if (floatingButton) return floatingButton;
  floatingButton = document.createElement("button");
  floatingButton.type = "button";
  floatingButton.id = FLOATING_BUTTON_ID;
  floatingButton.className = "hk-floating-btn";
  floatingButton.setAttribute("aria-label", "標註選取文字");
  floatingButton.title = "標註選取文字";
  const iconSrc =
    chrome?.runtime?.id && typeof chrome.runtime.getURL === "function"
      ? chrome.runtime.getURL("Icon/32.png")
      : null;
  if (iconSrc) {
    floatingButton.innerHTML = `<img src="${iconSrc}" class="hk-floating-btn-image" alt="" aria-hidden="true" />`;
  } else {
    floatingButton.textContent = "HL";
  }
  floatingButton.style.display = "none";
  floatingButton.addEventListener("mousedown", (event) => {
    // Prevent losing selection before highlight is applied.
    event.preventDefault();
  });
  floatingButton.addEventListener("click", async (event) => {
    event.preventDefault();
    try {
      await applyHighlight(currentColor);
    } catch (error) {
      console.debug("無法套用 highlight 按鈕動作", error);
    } finally {
      hideFloatingButton();
    }
  });
  document.body.appendChild(floatingButton);
  return floatingButton;
};

const hideFloatingButton = () => {
  if (floatingButton) {
    floatingButton.style.display = "none";
  }
};

const positionFloatingButton = (rect) => {
  const button = ensureFloatingButton();
  const { innerWidth, innerHeight } = window;
  const buttonRect = button.getBoundingClientRect();
  const width = buttonRect.width || 32;
  const height = buttonRect.height || 32;

  let top = rect.top - height - FLOATING_BUTTON_MARGIN;
  let left = rect.right - width;

  if (top < FLOATING_BUTTON_MARGIN) {
    top = rect.bottom + FLOATING_BUTTON_MARGIN;
  }

  if (left < FLOATING_BUTTON_MARGIN) {
    left = rect.left;
  }

  if (left + width > innerWidth - FLOATING_BUTTON_MARGIN) {
    left = innerWidth - width - FLOATING_BUTTON_MARGIN;
  }

  if (top + height > innerHeight - FLOATING_BUTTON_MARGIN) {
    top = innerHeight - height - FLOATING_BUTTON_MARGIN;
  }

  button.style.top = `${Math.max(FLOATING_BUTTON_MARGIN, top)}px`;
  button.style.left = `${Math.max(FLOATING_BUTTON_MARGIN, left)}px`;
  button.style.display = "flex";
};

const showFloatingButton = (range) => {
  const rect = range.getBoundingClientRect();
  if (!rect || (rect.width === 0 && rect.height === 0)) {
    hideFloatingButton();
    return;
  }
  positionFloatingButton(rect);
};

const closeHighlightMenu = () => {
  if (highlightMenuStatusTimer) {
    window.clearTimeout(highlightMenuStatusTimer);
    highlightMenuStatusTimer = null;
  }
  if (highlightMenuEls?.status) {
    highlightMenuEls.status.textContent = "";
  }
  if (highlightMenu) {
    highlightMenu.style.display = "none";
  }
  activeHighlight = null;
  activeHighlightId = null;
};

const setHighlightMenuStatus = (message, isError = false) => {
  if (!highlightMenuEls?.status) return;
  if (highlightMenuStatusTimer) {
    window.clearTimeout(highlightMenuStatusTimer);
    highlightMenuStatusTimer = null;
  }
  highlightMenuEls.status.textContent = message;
  highlightMenuEls.status.style.color = isError ? "#d93025" : "#1a73e8";
  if (message) {
    highlightMenuStatusTimer = window.setTimeout(() => {
      if (highlightMenuEls?.status) {
        highlightMenuEls.status.textContent = "";
      }
    }, 2000);
  }
};

const toHexColor = (value) => {
  if (!value) return DEFAULT_COLOR;
  if (value.startsWith("#")) {
    return value.toLowerCase();
  }
  const match = value.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (!match) return DEFAULT_COLOR;
  const [_, r, g, b] = match;
  const toHex = (component) =>
    Number(component).toString(16).padStart(2, "0").toLowerCase();
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
};

const sanitizePalette = (input) => {
  if (!Array.isArray(input)) {
    return [...DEFAULT_PALETTE];
  }
  const seen = new Set();
  const sanitized = [];
  input.forEach((value) => {
    if (typeof value !== "string") return;
    const color = toHexColor(value.trim());
    if (!seen.has(color)) {
      seen.add(color);
      sanitized.push(color);
    }
  });
  return sanitized.length ? sanitized : [...DEFAULT_PALETTE];
};

const setColorPaletteState = (palette) => {
  colorPalette = sanitizePalette(palette);
  renderHighlightMenuSwatches();
};

const persistColorPalette = async (nextPalette) => {
  const sanitized = sanitizePalette(nextPalette);
  colorPalette = sanitized;
  renderHighlightMenuSwatches();
  if (!storage) return;
  try {
    await storage.set({ hkColorPalette: sanitized });
  } catch (error) {
    console.debug("儲存顏色色票失敗", error);
  }
};

const loadPalette = async () => {
  if (!storage) return colorPalette;
  try {
    const stored = await storage.get("hkColorPalette");
    const palette = sanitizePalette(stored?.hkColorPalette);
    setColorPaletteState(palette);
    return colorPalette;
  } catch (error) {
    console.debug("讀取顏色清單失敗", error);
    setColorPaletteState(DEFAULT_PALETTE);
    return colorPalette;
  }
};

const refreshPaletteFromStorage = async () => {
  const latest = await loadPalette();
  return latest;
};

const renderHighlightMenuSwatches = () => {
  if (!highlightMenuEls?.swatchGroup) return;
  highlightMenuEls.swatchGroup.innerHTML = "";
  colorPalette.forEach((color) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "hk-menu-swatch";
    button.style.backgroundColor = color;
    button.dataset.color = color;
    button.title = color;
    button.addEventListener("click", () => applyColorChange(color));
    highlightMenuEls.swatchGroup.appendChild(button);
  });
};

const getPagePlainText = () => {
  const raw = document.body?.innerText || "";
  const normalized = raw.replace(/\n{3,}/g, "\n\n").trim();
  return normalized.slice(0, 15000);
};

const collectPageHighlights = async () => {
  const highlights = await getStoredHighlights();
  return highlights.map((item) => ({
    id: item.id,
    text: item.text ?? "",
    note: item.note ?? "",
    color: item.color,
    createdAt: item.createdAt,
    url: item.url ?? pageKey,
    range: item.range,
  }));
};

const persistAISettings = async () => {
  try {
    await chrome.storage?.local.set({ hkAISettings: aiSettings });
  } catch (error) {
    console.debug("儲存 AI 設定失敗", error);
  }
};

const updateAiKeyVisibility = () => {
  const groups = highlightPanelEls?.aiKeyGroups ?? [];
  groups.forEach((group) => {
    const provider = group.dataset.provider;
    const isVisible = provider === aiSettings.provider;
    group.classList.toggle("is-visible", isVisible);
    group.hidden = !isVisible;
    if (!isVisible) {
      const field = group.querySelector("input");
      if (field) {
        field.value = aiSettings[provider === "openai" ? "openaiKey" : "geminiKey"] ?? "";
      }
    }
  });
};

const populateAiModelSelect = () => {
  const select = highlightPanelEls?.aiModelSelect;
  if (!select) return;
  const provider = aiSettings.provider;
  const options = MODEL_OPTIONS[provider] ?? [];
  select.innerHTML = "";
  options.forEach((option) => {
    const opt = document.createElement("option");
    opt.value = option.value;
    opt.textContent = option.label;
    select.appendChild(opt);
  });
  const currentValue =
    provider === "openai" ? aiSettings.openaiModel : aiSettings.geminiModel;
  if (options.some((option) => option.value === currentValue)) {
    select.value = currentValue;
  } else if (options.length) {
    select.value = options[0].value;
    if (provider === "openai") {
      aiSettings.openaiModel = select.value;
    } else {
      aiSettings.geminiModel = select.value;
    }
    persistAISettings();
  }
};

const setAiPanelStatus = (message, isError = false) => {
  const statusEl = highlightPanelEls?.aiStatus;
  if (!statusEl) return;
  statusEl.textContent = message || "";
  statusEl.classList.toggle("is-error", Boolean(isError));
};

const updateGenerateAvailability = () => {
  const generateBtn = highlightPanelEls?.aiGenerateBtn;
  if (!generateBtn) return;
  const provider = aiSettings.provider;
  const key = provider === "openai" ? aiSettings.openaiKey : aiSettings.geminiKey;
  const hasKey = Boolean(key?.trim());
  generateBtn.disabled = isGeneratingNote || !hasKey;
  generateBtn.textContent = isGeneratingNote ? "產生中..." : "產生筆記";
};

const applyAiSettingsToUI = () => {
  const providerSelect = highlightPanelEls?.aiProviderSelect;
  const modelSelect = highlightPanelEls?.aiModelSelect;
  const promptField = highlightPanelEls?.aiPromptField;
  const openaiInput = highlightPanelEls?.aiOpenaiKeyInput;
  const geminiInput = highlightPanelEls?.aiGeminiKeyInput;

  if (providerSelect) {
    providerSelect.value = aiSettings.provider;
  }
  populateAiModelSelect();
  if (modelSelect) {
    modelSelect.value =
      aiSettings.provider === "openai"
        ? aiSettings.openaiModel
        : aiSettings.geminiModel;
  }
  if (promptField) {
    promptField.value = aiSettings.prompt ?? DEFAULT_AI_PROMPT;
  }
  if (openaiInput) {
    openaiInput.value = aiSettings.openaiKey ?? "";
  }
  if (geminiInput) {
    geminiInput.value = aiSettings.geminiKey ?? "";
  }
  updateAiKeyVisibility();
  updateGenerateAvailability();
};

const loadAISettings = async () => {
  try {
    const stored = await chrome.storage?.local.get("hkAISettings");
    if (stored?.hkAISettings) {
      aiSettings = {
        ...aiSettings,
        ...stored.hkAISettings,
      };
    }
  } catch (error) {
    console.debug("讀取 AI 設定失敗", error);
  } finally {
    applyAiSettingsToUI();
  }
};

const buildNotePrompt = (pageData) => {
  const basePrompt = aiSettings.prompt?.trim() || DEFAULT_AI_PROMPT;
  const highlightLines = (pageData.highlights || [])
    .map((item, idx) => {
      const noteText = item.note ? `（註解：${item.note.trim()}）` : "";
      return `${idx + 1}. ${item.text.trim()}${noteText}`;
    })
    .join("\n");

  return `${basePrompt}

### 網頁資訊
- 標題：${pageData.title}
- URL：${pageData.url}

### 原文內容（可能已截斷）
${pageData.pageText}

### 使用者標註
${highlightLines || "（尚未加入標註）"}
`;
};

const callOpenAI = async (key, prompt) => {
  const payload = {
    model: aiSettings.openaiModel || "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content:
          "You are a note-taking assistant who produces concise, easy-to-read study notes in Traditional Chinese.",
      },
      { role: "user", content: prompt },
    ],
    temperature: 0.4,
  };

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${key}`,
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`OpenAI API 錯誤：${errorText}`);
  }

  const json = await response.json();
  const text = json?.choices?.[0]?.message?.content;
  if (!text) {
    throw new Error("OpenAI 回傳內容為空");
  }
  return text.trim();
};

const callGemini = async (key, prompt) => {
  const model = aiSettings.geminiModel || DEFAULT_GEMINI_MODEL;
  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1/models/${model}:generateContent?key=${encodeURIComponent(
      key
    )}`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        contents: [
          {
            role: "user",
            parts: [{ text: prompt }],
          },
        ],
      }),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    if (response.status === 404 && model !== DEFAULT_GEMINI_MODEL) {
      aiSettings.geminiModel = DEFAULT_GEMINI_MODEL;
      await persistAISettings();
      return callGemini(key, prompt);
    }
    throw new Error(`Gemini API 錯誤：${errorText}`);
  }

  const json = await response.json();
  const parts = json?.candidates?.[0]?.content?.parts;
  const text = parts?.map((part) => part.text).join("\n");
  if (!text) {
    throw new Error("Gemini 回傳內容為空");
  }
  return text.trim();
};

const saveGeneratedNote = async (pageUrl, notePayload) => {
  try {
    const stored = await chrome.storage?.local.get("hkGeneratedNotes");
    const nextNotes = {
      ...(stored?.hkGeneratedNotes ?? {}),
      [pageUrl]: notePayload,
    };
    await chrome.storage?.local.set({ hkGeneratedNotes: nextNotes });
  } catch (error) {
    console.debug("儲存筆記失敗", error);
  }
};

const handleGenerateAiNote = async () => {
  if (isGeneratingNote || !highlightPanelEls?.aiGenerateBtn) return;
  const provider = aiSettings.provider;
  const apiKey =
    provider === "openai"
      ? aiSettings.openaiKey?.trim()
      : aiSettings.geminiKey?.trim();

  if (!apiKey) {
    setAiPanelStatus("請先輸入 API Key", true);
    updateGenerateAvailability();
    return;
  }

  try {
    isGeneratingNote = true;
    updateGenerateAvailability();
    setAiPanelStatus("產生筆記中…");
    const pageData = {
      title: document.title,
      url: pageKey,
      pageText: getPagePlainText(),
      highlights: await collectPageHighlights(),
    };
    const prompt = buildNotePrompt(pageData);
    const noteText =
      provider === "openai"
        ? await callOpenAI(apiKey, prompt)
        : await callGemini(apiKey, prompt);

    const notePayload = {
      note: noteText,
      provider,
      model: provider === "openai" ? aiSettings.openaiModel : aiSettings.geminiModel,
      prompt,
      generatedAt: Date.now(),
      url: pageData.url,
    };

    highlightPanelState.notesByPage = {
      ...highlightPanelState.notesByPage,
      [pageData.url]: notePayload,
    };
    if (highlightPanelState.activeKey !== pageKey) {
      highlightPanelState.activeKey = pageKey;
    }
    highlightPanelState.activeTab = "ai-note";
    applyHighlightPanelTabState();
    await saveGeneratedNote(pageData.url, notePayload);
    updateAiNoteSection(notePayload);
    setAiPanelStatus("已完成筆記產生");
  } catch (error) {
    console.debug("產生 AI 筆記失敗", error);
    setAiPanelStatus(error?.message || "無法產生筆記", true);
  } finally {
    isGeneratingNote = false;
    updateGenerateAvailability();
  }
};

const formatTimestamp = (value) => {
  if (!value) return "";
  try {
    return new Intl.DateTimeFormat(undefined, {
      dateStyle: "medium",
      timeStyle: "short",
    }).format(new Date(value));
  } catch (_error) {
    return "";
  }
};

const truncateText = (text, limit = 140) => {
  if (!text) return "";
  const normalized = text.trim().replace(/\s+/g, " ");
  if (normalized.length <= limit) return normalized;
  return `${normalized.slice(0, limit)}…`;
};

const parseTags = (input) => {
  if (typeof input !== "string") return [];
  return Array.from(
    new Set(
      input
        .split(/[,\s]+/)
        .map((tag) => tag.trim())
        .filter(Boolean)
    )
  );
};

const isValidPageKey = (key) => {
  if (typeof key !== "string") return false;
  try {
    const url = new URL(key);
    return ["http:", "https:"].includes(url.protocol);
  } catch (_error) {
    return false;
  }
};

const getPageDisplayName = (url) => {
  if (!url) return "未知頁面";
  try {
    const parsed = new URL(url);
    const path =
      parsed.pathname && parsed.pathname !== "/"
        ? parsed.pathname.replace(/\/$/, "")
        : "/";
    return `${parsed.hostname}${path}`;
  } catch (_error) {
    return url;
  }
};

const applyHighlightPanelSideClasses = () => {
  if (!highlightPanel) return;
  const panel = highlightPanel;
  const hasCustomPosition = isValidPanelPosition(highlightPanelState.position);
  if (hasCustomPosition) {
    const clamped = clampHighlightPanelPosition(highlightPanelState.position);
    if (clamped) {
      highlightPanelState.position = clamped;
      panel.classList.add("hk-panel-floating");
      panel.classList.remove("hk-panel-left", "hk-panel-right");
      panel.style.left = `${clamped.x}px`;
      panel.style.top = `${clamped.y}px`;
      panel.style.right = "auto";
      panel.style.bottom = "auto";
    }
  } else {
    panel.classList.remove("hk-panel-floating");
    panel.classList.toggle("hk-panel-left", highlightPanelState.side === "left");
    panel.classList.toggle(
      "hk-panel-right",
      highlightPanelState.side !== "left"
    );
    panel.style.removeProperty("left");
    panel.style.removeProperty("top");
    panel.style.removeProperty("right");
    panel.style.removeProperty("bottom");
  }
  updateFontControlAvailability();
};

const setHighlightPanelSide = async (side, persist = true) => {
  const resolved = side === "left" ? "left" : "right";
  highlightPanelState.side = resolved;
  highlightPanelPreferredSide = resolved;
  if (persist) {
    await clearHighlightPanelPosition(true);
  }
  applyHighlightPanelSideClasses();
  if (persist && storage) {
    try {
      await storage.set({ hkPanelSide: resolved });
    } catch (error) {
      console.debug("儲存面板位置失敗", error);
    }
  }
};

const applyHighlightPanelTabState = () => {
  const allowedTabs = ["page", "archive", "search", "ai-note"];
  const resolved = allowedTabs.includes(highlightPanelState.activeTab)
    ? highlightPanelState.activeTab
    : "page";
  highlightPanelState.activeTab = resolved;
  const { tabButtons, tabPanels } = highlightPanelEls ?? {};
  const buttonMap = {
    page: tabButtons?.page,
    archive: tabButtons?.archive,
    search: tabButtons?.search,
    "ai-note": tabButtons?.ai,
  };
  const panelMap = {
    page: tabPanels?.page,
    archive: tabPanels?.archive,
    search: tabPanels?.search,
    "ai-note": tabPanels?.ai,
  };

  allowedTabs.forEach((tab) => {
    const isActive = resolved === tab;
    const btn = buttonMap[tab];
    if (btn) {
      btn.classList.toggle("is-active", isActive);
      btn.setAttribute("aria-selected", isActive ? "true" : "false");
      btn.setAttribute("tabindex", isActive ? "0" : "-1");
    }
    const panel = panelMap[tab];
    if (panel) {
      panel.classList.toggle("is-active", isActive);
      panel.hidden = !isActive;
      panel.setAttribute("aria-hidden", isActive ? "false" : "true");
    }
  });
};

const isValidPanelPosition = (value) =>
  !!value &&
  typeof value.x === "number" &&
  typeof value.y === "number" &&
  Number.isFinite(value.x) &&
  Number.isFinite(value.y);

const clampHighlightPanelPosition = (position) => {
  if (!position) return null;
  const margin = PANEL_DRAG_MARGIN;
  const panelRect = highlightPanel?.getBoundingClientRect();
  const width = panelRect?.width || 320;
  const height = panelRect?.height || 460;
  const viewportWidth =
    window.innerWidth || document.documentElement?.clientWidth || width;
  const viewportHeight =
    window.innerHeight || document.documentElement?.clientHeight || height;
  const maxLeft = Math.max(margin, viewportWidth - width - margin);
  const maxTop = Math.max(margin, viewportHeight - height - margin);
  const sanitizedX = Number.isFinite(position.x) ? position.x : margin;
  const sanitizedY = Number.isFinite(position.y) ? position.y : margin;
  return {
    x: Math.min(Math.max(margin, sanitizedX), maxLeft),
    y: Math.min(Math.max(margin, sanitizedY), maxTop),
  };
};

const updateFontControlAvailability = () => {
  const controls = highlightPanelEls?.fontControls;
  if (!controls) return;
  const scale = highlightPanelState.fontScale ?? 1;
  const minLocked = scale <= PANEL_FONT_SCALE_MIN + 0.005;
  const maxLocked = scale >= PANEL_FONT_SCALE_MAX - 0.005;
  if (controls.decrease) {
    controls.decrease.disabled = minLocked;
  }
  if (controls.increase) {
    controls.increase.disabled = maxLocked;
  }
};

const applyHighlightPanelFontScale = () => {
  if (!highlightPanel) return;
  const scale = highlightPanelState.fontScale ?? 1;
  highlightPanel.style.setProperty(
    "--hk-panel-font-scale",
    scale.toString()
  );
  updateFontControlAvailability();
};

const clampPanelFontScale = (value) =>
  Math.min(
    Math.max(value, PANEL_FONT_SCALE_MIN),
    PANEL_FONT_SCALE_MAX
  );

const setHighlightPanelFontScale = async (scale, persist = true) => {
  const numeric = Number(scale);
  const clamped = clampPanelFontScale(Number.isFinite(numeric) ? numeric : 1);
  highlightPanelState.fontScale = Number(clamped.toFixed(2));
  applyHighlightPanelFontScale();
  if (persist && storage) {
    try {
      await storage.set({
        [HIGHLIGHT_PANEL_FONT_SCALE_KEY]: highlightPanelState.fontScale,
      });
    } catch (error) {
      console.debug("儲存面板字級失敗", error);
    }
  }
};

const adjustHighlightPanelFontScale = (delta) => {
  const current = highlightPanelState.fontScale ?? 1;
  setHighlightPanelFontScale(current + delta);
};

const setHighlightPanelPosition = async (x, y, persist = true) => {
  const clamped = clampHighlightPanelPosition({
    x: Number(x),
    y: Number(y),
  });
  if (!clamped) return;
  highlightPanelState.position = clamped;
  applyHighlightPanelSideClasses();
  if (persist && storage) {
    try {
      await storage.set({
        [HIGHLIGHT_PANEL_POSITION_KEY]: clamped,
      });
    } catch (error) {
      console.debug("儲存面板座標失敗", error);
    }
  }
};

const clearHighlightPanelPosition = async (persist = true) => {
  const hadPosition = Boolean(highlightPanelState.position);
  highlightPanelState.position = null;
  applyHighlightPanelSideClasses();
  if (persist && storage) {
    try {
      if (typeof storage.remove === "function") {
        await storage.remove(HIGHLIGHT_PANEL_POSITION_KEY);
      } else {
        await storage.set({ [HIGHLIGHT_PANEL_POSITION_KEY]: null });
      }
    } catch (error) {
      console.debug("清除面板座標失敗", error);
    }
  }
  return hadPosition;
};

const isPanelDragDisallowedTarget = (target) => {
  if (!(target instanceof Element)) return false;
  return Boolean(
    target.closest(
      "button, input, textarea, select, a, label, [role=\"button\"], [contenteditable=\"true\"]"
    )
  );
};

const handlePanelPointerDown = (event) => {
  if (!highlightPanel || !highlightPanelEls?.dragHandle) return;
  if (event.pointerType === "mouse" && event.button !== 0) return;
  if (isPanelDragDisallowedTarget(event.target)) return;
  event.preventDefault();
  const rect = highlightPanel.getBoundingClientRect();
  const initialPosition =
    clampHighlightPanelPosition({ x: rect.left, y: rect.top }) || {
      x: rect.left,
      y: rect.top,
    };
  const existingPosition = isValidPanelPosition(highlightPanelState.position)
    ? highlightPanelState.position
    : null;
  panelDragState = {
    pointerId: event.pointerId,
    startX: event.clientX,
    startY: event.clientY,
    originX: existingPosition?.x ?? initialPosition.x,
    originY: existingPosition?.y ?? initialPosition.y,
    initialized: Boolean(existingPosition),
  };
  highlightPanel.classList.add("is-dragging");
  if (panelDragState.initialized) {
    applyHighlightPanelSideClasses();
  }
  event.currentTarget?.setPointerCapture?.(event.pointerId);
};

const handlePanelPointerMove = (event) => {
  if (!panelDragState || event.pointerId !== panelDragState.pointerId) return;
  event.preventDefault();
  if (!panelDragState.initialized) {
    const starting = clampHighlightPanelPosition({
      x: panelDragState.originX,
      y: panelDragState.originY,
    });
    if (starting) {
      highlightPanelState.position = starting;
      panelDragState.originX = starting.x;
      panelDragState.originY = starting.y;
      panelDragState.initialized = true;
      applyHighlightPanelSideClasses();
    }
  }
  if (!panelDragState.initialized) return;
  const deltaX = event.clientX - panelDragState.startX;
  const deltaY = event.clientY - panelDragState.startY;
  const nextX = panelDragState.originX + deltaX;
  const nextY = panelDragState.originY + deltaY;
  setHighlightPanelPosition(nextX, nextY, false);
};

const handlePanelPointerEnd = (event) => {
  if (!panelDragState || event.pointerId !== panelDragState.pointerId) return;
  event.preventDefault();
  event.currentTarget?.releasePointerCapture?.(event.pointerId);
  highlightPanel?.classList.remove("is-dragging");
  const hasPosition = isValidPanelPosition(highlightPanelState.position);
  const finalPos = hasPosition
    ? highlightPanelState.position
    : panelDragState.initialized
      ? {
          x: panelDragState.originX,
          y: panelDragState.originY,
        }
      : null;
  panelDragState = null;
  if (finalPos) {
    setHighlightPanelPosition(finalPos.x, finalPos.y, true);
  }
};

const loadPanelPreferences = async () => {
  if (!storage) return;
  try {
    const stored = await storage.get([
      "hkPanelSide",
      HIGHLIGHT_PANEL_POSITION_KEY,
      HIGHLIGHT_PANEL_FONT_SCALE_KEY,
    ]);
    const side =
      stored?.hkPanelSide === "left" || stored?.hkPanelSide === "right"
        ? stored.hkPanelSide
        : "right";
    highlightPanelState.side = side;
    highlightPanelPreferredSide = side;
    const savedPosition = stored?.[HIGHLIGHT_PANEL_POSITION_KEY];
    if (isValidPanelPosition(savedPosition)) {
      highlightPanelState.position = savedPosition;
    }
    const savedScale = stored?.[HIGHLIGHT_PANEL_FONT_SCALE_KEY];
    if (typeof savedScale === "number" && Number.isFinite(savedScale)) {
      highlightPanelState.fontScale = clampPanelFontScale(savedScale);
    }
    applyHighlightPanelSideClasses();
    applyHighlightPanelFontScale();
  } catch (error) {
    console.debug("讀取面板設定失敗", error);
  }
};

const updateHighlightPanelTagFilters = () => {
  if (!highlightPanelEls?.tagsContainer) return;
  const container = highlightPanelEls.tagsContainer;
  container.innerHTML = "";

  if (!highlightPanelState.allTags.length) {
    const hint = document.createElement("span");
    hint.className = "hk-panel-tags-empty";
    hint.textContent = "尚未建立標籤";
    container.appendChild(hint);
    return;
  }

  const clearBtn = document.createElement("button");
  clearBtn.type = "button";
  clearBtn.className = "hk-panel-tag-chip hk-panel-tag-clear";
  clearBtn.textContent = "全部";
  clearBtn.classList.toggle("is-active", !highlightPanelState.activeTag);
  clearBtn.addEventListener("click", () => {
    highlightPanelState.activeTag = null;
    updateHighlightPanelTagFilters();
    renderHighlightPanel();
  });
  container.appendChild(clearBtn);

  highlightPanelState.allTags.forEach((tag) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "hk-panel-tag-chip";
    chip.textContent = tag;
    chip.classList.toggle("is-active", highlightPanelState.activeTag === tag);
    chip.addEventListener("click", () => {
      highlightPanelState.activeTag =
        highlightPanelState.activeTag === tag ? null : tag;
      updateHighlightPanelTagFilters();
      renderHighlightPanel();
    });
    container.appendChild(chip);
  });
};

const updateSearchPageFilterOptions = () => {
  const select = highlightPanelEls?.searchPageSelect;
  if (!select) return;
  const options = [];
  const addOption = (value, label) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    options.push(option);
  };
  addOption("__all__", "全部頁面");
  const keys = Object.keys(highlightPanelState.allPages || {}).sort((a, b) =>
    getPageDisplayName(a).localeCompare(getPageDisplayName(b))
  );
  keys.forEach((key) => {
    const title = highlightPanelState.pageMeta?.[key]?.title?.trim();
    addOption(key, title || getPageDisplayName(key));
  });
  select.innerHTML = "";
  options.forEach((option) => select.appendChild(option));
  const activeValue =
    highlightPanelState.searchPageFilter && keys.includes(highlightPanelState.searchPageFilter)
      ? highlightPanelState.searchPageFilter
      : "__all__";
  select.value = activeValue;
  highlightPanelState.searchPageFilter =
    activeValue === "__all__" ? null : activeValue;
};

const updateTagSuggestionDropdown = () => {
  if (!highlightPanelEls?.tagInput || !highlightPanelEls?.suggestionDropdown)
    return;
  const dropdown = highlightPanelEls.suggestionDropdown;
  const value = highlightPanelEls.tagInput.value.trim().toLowerCase();
  dropdown.innerHTML = "";
  if (!value) {
    dropdown.classList.remove("is-visible");
    return;
  }
  const matches = highlightPanelState.allTags.filter((tag) =>
    tag.toLowerCase().includes(value)
  );
  if (!matches.length) {
    dropdown.classList.remove("is-visible");
    return;
  }
  dropdown.classList.add("is-visible");
  matches.slice(0, 6).forEach((tag) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "hk-panel-tag-suggestion";
    item.textContent = tag;
    item.addEventListener("click", () => {
      highlightPanelEls.tagInput.value = tag;
      dropdown.classList.remove("is-visible");
      highlightPanelEls.addTagBtn?.click();
    });
    dropdown.appendChild(item);
  });
};

const setPanelStatus = (message, isError = false) => {
  if (!highlightPanelEls?.exportStatus) return;
  const el = highlightPanelEls.exportStatus;
  if (panelStatusTimer) {
    window.clearTimeout(panelStatusTimer);
    panelStatusTimer = null;
  }
  el.textContent = message || "";
  el.classList.toggle("is-error", isError);
  if (message) {
    panelStatusTimer = window.setTimeout(() => {
      if (highlightPanelEls?.exportStatus) {
        highlightPanelEls.exportStatus.textContent = "";
        highlightPanelEls.exportStatus.classList.remove("is-error");
      }
    }, 2400);
  }
  if (highlightPanelState.activeTab === "archive") {
    setArchiveStatus(message, isError);
  }
};

const setArchiveStatus = (message, isError = false) => {
  if (!highlightPanelEls?.archiveStatus) return;
  const el = highlightPanelEls.archiveStatus;
  if (archiveStatusTimer) {
    window.clearTimeout(archiveStatusTimer);
    archiveStatusTimer = null;
  }
  el.textContent = message || "";
  el.classList.toggle("is-error", isError);
  if (message) {
    archiveStatusTimer = window.setTimeout(() => {
      if (highlightPanelEls?.archiveStatus) {
        highlightPanelEls.archiveStatus.textContent = "";
        highlightPanelEls.archiveStatus.classList.remove("is-error");
      }
    }, 2600);
  }
};

const updateExportButtonsState = () => {
  const copyBtn = highlightPanelEls?.copyBtn;
  const downloadBtn = highlightPanelEls?.downloadBtn;
  if (!copyBtn || !downloadBtn) return;
  const hasEntries = (highlightPanelState.currentEntries?.length ?? 0) > 0;
  [copyBtn, downloadBtn].forEach((btn) => {
    btn.disabled = !hasEntries;
    btn.classList.toggle("is-disabled", !hasEntries);
    btn.title = hasEntries
      ? btn.textContent
      : "無標註可匯出";
  });
};

const deleteHighlightFromPanel = async (entry) => {
  try {
    const targetUrl = entry.pageUrl ?? pageKey;
    await deleteHighlightEntry(entry.id, targetUrl);
    if (targetUrl === pageKey) {
      const targetEl = document.querySelector(
        `[${HIGHLIGHT_ATTR}="${entry.id}"]`
      );
      if (targetEl) {
        unwrapHighlightElement(targetEl);
      }
    }
    setPanelStatus("已刪除標註");
    await refreshHighlightPanelData();
    await renderHighlightPanel();
  } catch (error) {
    console.debug("刪除標註失敗", error);
    setPanelStatus("刪除失敗", true);
  }
};

const updateAiNoteSection = (noteData) => {
  const {
    aiNoteSection,
    aiNoteContent,
    aiNoteMeta,
    aiNoteCopyBtn,
    aiNoteEmpty,
  } = highlightPanelEls ?? {};
  if (!aiNoteSection || !aiNoteContent || !aiNoteMeta || !aiNoteCopyBtn) {
    return;
  }

  if (aiNoteCopyBtn._hkResetTimer) {
    window.clearTimeout(aiNoteCopyBtn._hkResetTimer);
    aiNoteCopyBtn._hkResetTimer = null;
  }
  aiNoteCopyBtn.classList.remove("is-error");

  if (noteData?.note) {
    aiNoteContent.textContent = noteData.note;
    aiNoteContent.style.display = "block";
    aiNoteEmpty.style.display = "none";
    const providerLabel = noteData.provider === "openai" ? "OpenAI" : "Gemini";
    const modelLabel = noteData.model ? `${providerLabel} · ${noteData.model}` : providerLabel;
    const generatedLabel = noteData.generatedAt
      ? `${modelLabel} · ${formatTimestamp(noteData.generatedAt)}`
      : modelLabel;
    aiNoteMeta.textContent = generatedLabel;
    aiNoteCopyBtn.disabled = false;
    aiNoteCopyBtn.textContent = "複製";
    aiNoteCopyBtn.onclick = async () => {
      if (aiNoteCopyBtn._hkResetTimer) {
        window.clearTimeout(aiNoteCopyBtn._hkResetTimer);
        aiNoteCopyBtn._hkResetTimer = null;
      }
      try {
        await navigator.clipboard.writeText(noteData.note);
        aiNoteCopyBtn.textContent = "已複製";
      } catch (error) {
        console.debug("複製 AI 筆記失敗", error);
        aiNoteCopyBtn.textContent = "複製失敗";
        aiNoteCopyBtn.classList.add("is-error");
      } finally {
        aiNoteCopyBtn._hkResetTimer = window.setTimeout(() => {
          aiNoteCopyBtn.textContent = "複製";
          aiNoteCopyBtn.classList.remove("is-error");
          aiNoteCopyBtn._hkResetTimer = null;
        }, 1800);
      }
    };
  } else {
    aiNoteContent.textContent = "";
    aiNoteContent.style.display = "none";
    aiNoteEmpty.style.display = "block";
    aiNoteMeta.textContent = "";
    aiNoteCopyBtn.disabled = true;
    aiNoteCopyBtn.onclick = null;
    aiNoteCopyBtn.textContent = "複製";
  }
};

const getCurrentPanelEntries = () => highlightPanelState.currentEntries ?? [];

const buildHighlightExportEntry = (entry) => {
  if (!entry) return null;
  const url = entry.pageUrl || entry.url || pageKey;
  const range = entry.range;
  if (!url || !range) return null;
  const payload = {
    id: entry.id,
    url,
    color: toHexColor(entry.color || DEFAULT_COLOR),
    text: entry.text ?? "",
    note: entry.note ?? "",
    range,
    createdAt: entry.createdAt ?? Date.now(),
  };
  const tags = Array.isArray(entry.tags)
    ? entry.tags
    : Array.isArray(entry.pageTags)
    ? entry.pageTags
    : null;
  if (tags?.length) {
    payload.tags = Array.from(
      new Set(tags.map((tag) => String(tag).trim()).filter(Boolean))
    );
  }
  const title = highlightPanelState.pageMeta[url]?.title;
  if (title) {
    payload.title = title;
  }
  return payload;
};

const exportHighlights = async (mode) => {
  const entries = getCurrentPanelEntries();
  if (!entries.length) {
    setPanelStatus("沒有可匯出的標註", true);
    return;
  }
  const lines = entries.map(({ text, note }) => {
    const trimmed = (text ?? "").trim();
    const noteLine = note ? `（註解：${note.trim()}）` : "";
    return [trimmed, noteLine].filter(Boolean).join(" ");
  });
  const payloadText = lines.join("\n\n");

  if (mode === "copy" && navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(payloadText);
      setPanelStatus("已複製純文字");
    } catch (error) {
      console.debug("複製標註失敗", error);
      setPanelStatus("複製失敗", true);
    }
    return;
  }

  const exportEntries = entries
    .map((entry) => buildHighlightExportEntry(entry))
    .filter(Boolean);
  if (!exportEntries.length) {
    setPanelStatus("沒有可匯出的標註", true);
    return;
  }
  const payload = {
    type: "highlight-keeper",
    version: 1,
    exportedAt: Date.now(),
    total: exportEntries.length,
    entries: exportEntries,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  const filenameBase = highlightPanelState.activeKey
    ? getPageDisplayName(highlightPanelState.activeKey).replaceAll(/[\\/]+/g, "_")
    : "highlights";
  link.download = `${filenameBase}-highlights.json`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  setPanelStatus("已下載 JSON");
};

const parseImportedHighlightsPayload = (rawText) => {
  let parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (_error) {
    throw new Error("JSON 格式不正確");
  }
  if (Array.isArray(parsed)) return parsed;
  if (parsed && Array.isArray(parsed.entries)) return parsed.entries;
  if (parsed && typeof parsed === "object") return [parsed];
  return [];
};

const sanitizeImportedAnchors = (anchors) => {
  if (!anchors || typeof anchors !== "object") return undefined;
  const sanitizeBoundary = (boundary) => {
    if (!boundary || typeof boundary !== "object") return undefined;
    const css =
      typeof boundary.css === "string" && boundary.css.trim()
        ? boundary.css.trim()
        : undefined;
    const offset =
      typeof boundary.textOffset === "number" && Number.isFinite(boundary.textOffset)
        ? boundary.textOffset
        : undefined;
    if (!css && typeof offset === "undefined") return undefined;
    return { css, textOffset: offset };
  };
  const sanitized = {
    version:
      typeof anchors.version === "number" && Number.isFinite(anchors.version)
        ? anchors.version
        : undefined,
    start: sanitizeBoundary(anchors.start),
    end: sanitizeBoundary(anchors.end),
    quote:
      anchors.quote && typeof anchors.quote === "object"
        ? {
            exact: typeof anchors.quote.exact === "string" ? anchors.quote.exact : "",
            prefix: typeof anchors.quote.prefix === "string" ? anchors.quote.prefix : "",
            suffix: typeof anchors.quote.suffix === "string" ? anchors.quote.suffix : "",
          }
        : undefined,
  };
  if (!sanitized.start && !sanitized.end && !sanitized.quote) {
    return undefined;
  }
  return sanitized;
};

const normalizeImportedHighlightEntry = (entry, index) => {
  if (!entry || typeof entry !== "object") return null;
  const url =
    typeof entry.url === "string"
      ? entry.url
      : typeof entry.pageUrl === "string"
      ? entry.pageUrl
      : null;
  if (!url) return null;
  const range = entry.range;
  if (
    !range ||
    typeof range.startXPath !== "string" ||
    typeof range.endXPath !== "string"
  ) {
    return null;
  }
  const color = toHexColor(entry.color || DEFAULT_COLOR);
  const normalized = {
    id: `hk-import-${Date.now()}-${Math.floor(
      Math.random() * 100000
    )}-${index}`,
    color,
    text: typeof entry.text === "string" ? entry.text : range.text ?? "",
    note: typeof entry.note === "string" ? entry.note : "",
    range: {
      startXPath: range.startXPath,
      startOffset: Number(range.startOffset) || 0,
      endXPath: range.endXPath,
      endOffset: Number(range.endOffset) || 0,
      text: range.text ?? entry.text ?? "",
      anchors: sanitizeImportedAnchors(range.anchors),
    },
    url,
    createdAt:
      typeof entry.createdAt === "number" && Number.isFinite(entry.createdAt)
        ? entry.createdAt
        : Date.now(),
    tags: Array.isArray(entry.tags)
      ? entry.tags.map((tag) => String(tag).trim()).filter(Boolean)
      : parseTags(entry.tags ?? ""),
    title:
      typeof entry.title === "string" && entry.title.trim()
        ? entry.title.trim()
        : undefined,
  };
  return normalized;
};

const importHighlightsFromEntries = async (rawEntries) => {
  if (!storage) throw new Error("無法存取瀏覽器儲存空間");
  const normalized = rawEntries
    .map((entry, index) => normalizeImportedHighlightEntry(entry, index))
    .filter(Boolean);
  if (!normalized.length) {
    throw new Error("沒有可匯入的標註");
  }
  const grouped = new Map();
  normalized.forEach((entry) => {
    if (!grouped.has(entry.url)) {
      grouped.set(entry.url, []);
    }
    grouped.get(entry.url).push(entry);
  });
  const urls = Array.from(grouped.keys());
  const existing = await storage.get(urls);
  const updates = {};
  grouped.forEach((list, url) => {
    const current = Array.isArray(existing[url]) ? existing[url] : [];
    updates[url] = [
      ...current,
      ...list.map(({ title, ...rest }) => rest),
    ];
  });
  if (Object.keys(updates).length) {
    await storage.set(updates);
  }
  await Promise.all(
    normalized
      .filter((entry) => entry.title)
      .map((entry) => ensurePageMetaTitle(entry.url, entry.title))
  );
  const newColors = Array.from(
    new Set(
      normalized
        .map((entry) => entry.color)
        .filter((color) => !colorPalette.includes(color))
    )
  );
  if (newColors.length) {
    await persistColorPalette([...colorPalette, ...newColors]);
  }
  return normalized.length;
};

const handleImportFileChange = async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    const text = await file.text();
    const rawEntries = parseImportedHighlightsPayload(text);
    if (!rawEntries.length) {
      throw new Error("檔案中沒有標註資料");
    }
    const importedCount = await importHighlightsFromEntries(rawEntries);
    setPanelStatus(`已匯入 ${importedCount} 則標註`);
    await refreshHighlightPanelData();
    await renderHighlightPanel();
    attemptRestoreHighlights();
  } catch (error) {
    console.debug("匯入標註失敗", error);
    setPanelStatus(error.message || "匯入失敗", true);
  } finally {
    event.target.value = "";
  }
};

const renderPageTagEditor = async () => {
  if (!highlightPanelEls?.pageTagSection) return;
  const {
    pageTagSection,
    pageTagHint,
    pageTagList,
    tagInput,
    addTagBtn,
    suggestionDropdown,
  } = highlightPanelEls;
  updateTagSuggestionDropdown();

  highlightPanelState.activeKey = pageKey;
  const activeKey = pageKey;
  const meta = highlightPanelState.pageMeta || {};
  const currentPageTags = Array.isArray(meta[activeKey]?.tags)
    ? meta[activeKey].tags
    : [];

  tagInput.disabled = false;
  addTagBtn.disabled = false;
  pageTagHint.textContent = "輸入後按 Enter 或按新增；點擊標籤可移除";

  pageTagList.innerHTML = "";
  pageTagList.classList.remove("is-disabled");

  if (!currentPageTags.length) {
    const empty = document.createElement("span");
    empty.className = "hk-panel-tags-empty";
    empty.textContent = "尚未為此頁設定標籤";
    pageTagList.appendChild(empty);
  } else {
    currentPageTags.forEach((tag) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "hk-panel-tag-chip";
      chip.textContent = tag;
      chip.title = "移除此標籤";
      chip.addEventListener("click", async () => {
        const nextTags = currentPageTags.filter((item) => item !== tag);
        await setPageTags(activeKey, nextTags);
        highlightPanelState.pageMeta[activeKey] = {
          ...(highlightPanelState.pageMeta[activeKey] ?? {}),
          tags: nextTags,
        };
        if (highlightPanelState.activeTag === tag) {
          highlightPanelState.activeTag = null;
        }
        await refreshHighlightPanelData();
        await renderHighlightPanel();
      });
      pageTagList.appendChild(chip);
    });
  }

  const handleAddTag = async () => {
    const value = tagInput.value.trim();
    if (!value) return;
    const tags = Array.from(new Set([...currentPageTags, value]));
    await setPageTags(activeKey, tags);
    highlightPanelState.pageMeta[activeKey] = {
      ...(highlightPanelState.pageMeta[activeKey] ?? {}),
      tags,
    };
    tagInput.value = "";
    await refreshHighlightPanelData();
    await renderHighlightPanel();
    updateTagSuggestionDropdown();
  };

  addTagBtn.onclick = handleAddTag;
  tagInput.onkeydown = (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      handleAddTag();
    }
  };

  const handleInput = () => {
    updateTagSuggestionDropdown();
  };

  tagInput.oninput = handleInput;
  tagInput.onfocus = handleInput;
  tagInput.onblur = () => {
    window.setTimeout(() => {
      highlightPanelEls?.suggestionDropdown?.classList.remove("is-visible");
    }, 120);
  };

  updateTagSuggestionDropdown();
};

const refreshHighlightPanelData = async () => {
  if (!storage) {
    highlightPanelState.allPages = { [pageKey]: [] };
    highlightPanelState.pageMeta = {};
    highlightPanelState.allTags = [];
    return;
  }
  try {
    const everything = await storage.get(null);
    const storedMeta = await getPageMeta();
    const pages = {};
    Object.entries(everything).forEach(([key, value]) => {
      if (isValidPageKey(key) && Array.isArray(value)) {
        pages[key] = value;
      }
    });
    if (!pages[pageKey]) {
      pages[pageKey] = await getStoredHighlights(pageKey);
    }

    const allKeys = new Set([
      ...Object.keys(pages),
      ...Object.keys(storedMeta),
    ]);
    const combinedMeta = {};
    const tagSet = new Set();
    let metaChanged = false;

    allKeys.forEach((url) => {
      const existing = storedMeta[url];
      const entryTags = (pages[url] ?? [])
        .flatMap((entry) =>
          Array.isArray(entry?.tags)
            ? entry.tags
            : parseTags(entry?.tags ?? "")
        )
        .map((tag) => tag.trim())
        .filter(Boolean);
      const metaTags = Array.isArray(existing?.tags)
        ? existing.tags.map((tag) => tag.trim()).filter(Boolean)
        : [];
      const merged = Array.from(new Set([...metaTags, ...entryTags])).sort((a, b) =>
        a.localeCompare(b)
      );
      merged.forEach((tag) => tagSet.add(tag));
      combinedMeta[url] = {
        ...(existing ?? {}),
        tags: merged,
      };
      if (!combinedMeta[url].title && url === pageKey) {
        combinedMeta[url].title = document.title || "";
      }
      if (
        merged.length !== metaTags.length ||
        merged.some((tag, idx) => tag !== metaTags[idx])
      ) {
        metaChanged = true;
      }
    });

    if (metaChanged) {
      await setPageMeta(combinedMeta);
    }

    highlightPanelState.allPages = pages;
    highlightPanelState.pageMeta = combinedMeta;
    highlightPanelState.allTags = Array.from(tagSet).sort((a, b) =>
      a.localeCompare(b)
    );
    try {
      const storedNotes = await chrome.storage?.local.get("hkGeneratedNotes");
      highlightPanelState.notesByPage = storedNotes?.hkGeneratedNotes ?? {};
    } catch (noteError) {
      console.debug("讀取筆記快取失敗", noteError);
      highlightPanelState.notesByPage = {};
    }
    if (
      highlightPanelState.activeTag &&
      !tagSet.has(highlightPanelState.activeTag)
    ) {
      highlightPanelState.activeTag = null;
    }
    highlightPanelState.activeKey = pageKey;
  } catch (error) {
    console.debug("讀取所有標註資料失敗", error);
    const fallbackEntries = await getStoredHighlights();
    const fallbackMeta = await getPageMeta();
    Object.entries(fallbackMeta).forEach(([url, meta]) => {
      if (Array.isArray(meta?.tags)) {
        fallbackMeta[url] = {
          ...meta,
          tags: Array.from(new Set(meta.tags.map((tag) => tag.trim()).filter(Boolean))).sort(
            (a, b) => a.localeCompare(b)
          ),
        };
      }
    });
    highlightPanelState.allPages = { [pageKey]: fallbackEntries };
    highlightPanelState.pageMeta = fallbackMeta;
    if (!highlightPanelState.pageMeta[pageKey]) {
      highlightPanelState.pageMeta[pageKey] = {};
    }
    if (!highlightPanelState.pageMeta[pageKey].title) {
      highlightPanelState.pageMeta[pageKey].title = document.title || "";
    }
    const entryTags = fallbackEntries
      .flatMap((entry) =>
        Array.isArray(entry?.tags) ? entry.tags : parseTags(entry?.tags ?? "")
      )
      .filter((tag) => tag.trim().length);
    const metaTags = Object.values(fallbackMeta)
      .flatMap((meta) => (Array.isArray(meta?.tags) ? meta.tags : []))
      .filter((tag) => tag.trim().length);
    highlightPanelState.allTags = Array.from(
      new Set([...entryTags, ...metaTags])
    ).sort((a, b) => a.localeCompare(b));
    try {
      const storedNotes = await chrome.storage?.local.get("hkGeneratedNotes");
      highlightPanelState.notesByPage = storedNotes?.hkGeneratedNotes ?? {};
    } catch (noteError) {
      highlightPanelState.notesByPage = {};
    }
    if (
      highlightPanelState.activeTag &&
      !highlightPanelState.allTags.includes(highlightPanelState.activeTag)
    ) {
      highlightPanelState.activeTag = null;
    }
    highlightPanelState.activeKey = pageKey;
  }
  updateHighlightPanelTagFilters();
};

const collectPanelEntries = (scope = "current", opts = {}) => {
  const allPages = highlightPanelState.allPages;
  if (!allPages) return [];
  const pageMeta = highlightPanelState.pageMeta || {};
  const fetchPageTags = (url) => {
    const meta = pageMeta[url];
    if (Array.isArray(meta?.tags)) {
      return meta.tags;
    }
    return [];
  };
  const includeAll =
    scope === "all" || (scope === "search-filter" && !opts.pageFilter);
  if (includeAll) {
    return Object.entries(allPages)
      .filter(([url]) => (opts.pageFilter ? url === opts.pageFilter : true))
      .flatMap(([url, items]) =>
        items.map((entry) => ({
          ...entry,
          pageUrl: url,
          pageTags: fetchPageTags(url),
        }))
      );
  }
  const targetKey =
    scope === "current"
      ? pageKey
      : (typeof scope === "string" && isValidPageKey(scope) && scope) || pageKey;
  const entries = allPages[targetKey] ?? [];
  const pageTags = fetchPageTags(targetKey);
  return entries.map((entry) => ({
    ...entry,
    pageUrl: targetKey,
    pageTags,
  }));
};

const createPanelEntryElement = (entry, options = {}) => {
  const item = document.createElement("article");
  item.className = "hk-panel-item";
  item.setAttribute("data-highlight-id", entry.id);

  const meta = document.createElement("div");
  meta.className = "hk-panel-meta";

  const colorDot = document.createElement("span");
  colorDot.className = "hk-panel-color";
  colorDot.style.backgroundColor = entry.color ?? DEFAULT_COLOR;

  meta.appendChild(colorDot);

  if (options.showTitle) {
    const title = document.createElement("button");
    title.type = "button";
    title.className = "hk-panel-entry-title";
    const pageTitle =
      highlightPanelState.pageMeta?.[entry.pageUrl]?.title?.trim() || "";
    title.textContent = pageTitle || getPageDisplayName(entry.pageUrl);
    title.addEventListener("click", () => {
      window.open(entry.pageUrl, "_blank", "noopener");
    });
    meta.appendChild(title);
  } else {
    const pageLabel = document.createElement("span");
    pageLabel.className = "hk-panel-page";
    pageLabel.textContent =
      entry.pageUrl === pageKey ? "本頁" : getPageDisplayName(entry.pageUrl);
    meta.appendChild(pageLabel);
  }

  if (options.showUrl) {
    const permalink = document.createElement("a");
    permalink.className = "hk-panel-entry-url";
    permalink.href = entry.pageUrl;
    permalink.target = "_blank";
    permalink.rel = "noopener";
    permalink.textContent = entry.pageUrl.replace(/^https?:\/\//, "");
    meta.appendChild(permalink);
  }

  const timestamp = document.createElement("span");
  timestamp.className = "hk-panel-time";
  timestamp.textContent = formatTimestamp(entry.createdAt);
  meta.appendChild(timestamp);

  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "hk-panel-delete-btn";
  deleteBtn.textContent = "刪除";
  deleteBtn.addEventListener("click", async (event) => {
    event.stopPropagation();
    await deleteHighlightFromPanel(entry);
  });
  meta.appendChild(deleteBtn);

  const text = document.createElement("p");
  text.className = "hk-panel-text";
  text.textContent = truncateText(entry.text) || "(無文字)";

  item.appendChild(meta);
  item.appendChild(text);

  if (entry.pageTags?.length && !options.hideTags) {
    const tagsRow = document.createElement("div");
    tagsRow.className = "hk-panel-tags-row";
    entry.pageTags.forEach((tag) => {
      const chip = document.createElement("span");
      chip.className = "hk-panel-tag-chip";
      chip.textContent = tag;
      tagsRow.appendChild(chip);
    });
    item.appendChild(tagsRow);
  }

  if (entry.pageUrl === pageKey) {
    item.addEventListener("click", () => focusHighlightElement(entry.id));
  } else {
    item.classList.add("is-external");
    item.addEventListener("click", () => {
      window.open(entry.pageUrl, "_blank", "noopener");
    });
  }

  return item;
};

const renderPageEntries = (entries) => {
  const list = highlightPanelEls?.pageList;
  const placeholder = highlightPanelEls?.pagePlaceholder;
  if (!list || !placeholder) return;
  highlightPanelState.currentEntries = entries;
  list.innerHTML = "";
  if (!entries.length) {
    placeholder.style.display = "block";
    updateExportButtonsState();
    setPanelStatus("");
    return;
  }
  placeholder.style.display = "none";
  entries.forEach((entry) => {
    list.appendChild(createPanelEntryElement(entry));
  });
  updateExportButtonsState();
  setPanelStatus("");
};

const renderSearchEntries = (entries) => {
  const list = highlightPanelEls?.searchList;
  const placeholder = highlightPanelEls?.searchPlaceholder;
  if (!list || !placeholder) return;
  list.innerHTML = "";
  const hasFilter =
    Boolean(highlightPanelState.searchTerm.trim()) ||
    Boolean(highlightPanelState.activeTag);
  if (!entries.length) {
    placeholder.textContent = hasFilter
      ? "沒有符合搜尋條件的筆記。"
      : "目前尚未有跨頁標註。";
    placeholder.style.display = "block";
    return;
  }
  placeholder.style.display = "none";
  entries.forEach((entry) => {
    list.appendChild(
      createPanelEntryElement(entry, {
        showTitle: true,
        showUrl: true,
        hideTags: true,
      })
    );
  });
};

const focusHighlightElement = (id) => {
  if (!id) return;
  const element = document.querySelector(`[${HIGHLIGHT_ATTR}="${id}"]`);
  if (!element) return;
  closeHighlightMenu();
  element.scrollIntoView({ behavior: "smooth", block: "center" });
  element.classList.add("hk-highlight-focus");
  window.setTimeout(() => {
    element.classList.remove("hk-highlight-focus");
  }, 1600);
};

const ensureHighlightPanel = () => {
  if (highlightPanel) return highlightPanel;
  const panel = document.createElement("aside");
  panel.id = HIGHLIGHT_PANEL_ID;
  panel.className = "hk-page-panel";
  panel.style.display = "none";

  const header = document.createElement("div");
  header.className = "hk-panel-header";

  const title = document.createElement("h2");
  title.className = "hk-panel-title";
  title.textContent = "此頁標註";

  const fontControls = document.createElement("div");
  fontControls.className = "hk-panel-font-controls";

  const fontDecreaseBtn = document.createElement("button");
  fontDecreaseBtn.type = "button";
  fontDecreaseBtn.className = "hk-panel-font-btn";
  fontDecreaseBtn.setAttribute("aria-label", "縮小文字");
  fontDecreaseBtn.textContent = "A-";
  fontDecreaseBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    adjustHighlightPanelFontScale(-PANEL_FONT_SCALE_STEP);
  });

  const fontIncreaseBtn = document.createElement("button");
  fontIncreaseBtn.type = "button";
  fontIncreaseBtn.className = "hk-panel-font-btn";
  fontIncreaseBtn.setAttribute("aria-label", "放大文字");
  fontIncreaseBtn.textContent = "A+";
  fontIncreaseBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    adjustHighlightPanelFontScale(PANEL_FONT_SCALE_STEP);
  });

  fontControls.appendChild(fontDecreaseBtn);
  fontControls.appendChild(fontIncreaseBtn);

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "hk-panel-close";
  closeBtn.setAttribute("aria-label", "關閉標註面板");
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => closeHighlightPanel());

  header.appendChild(title);
  header.appendChild(fontControls);
  header.appendChild(closeBtn);
  header.addEventListener("pointerdown", handlePanelPointerDown);
  header.addEventListener("pointermove", handlePanelPointerMove);
  header.addEventListener("pointerup", handlePanelPointerEnd);
  header.addEventListener("pointercancel", handlePanelPointerEnd);

  const tabs = document.createElement("div");
  tabs.className = "hk-panel-tabs";
  tabs.setAttribute("role", "tablist");
  tabs.setAttribute("aria-label", "面板內容切換");

  const pageTabBtn = document.createElement("button");
  pageTabBtn.type = "button";
  pageTabBtn.className = "hk-panel-tab-btn";
  pageTabBtn.id = "hk-panel-tab-page";
  pageTabBtn.setAttribute("role", "tab");
  pageTabBtn.setAttribute("aria-controls", "hk-panel-tabpanel-page");
  pageTabBtn.textContent = "標註";
  pageTabBtn.addEventListener("click", () => {
    if (highlightPanelState.activeTab === "page") return;
    highlightPanelState.activeTab = "page";
    applyHighlightPanelTabState();
    renderHighlightPanel().catch((error) =>
      console.debug("重新整理標註面板失敗", error)
    );
  });

  const archiveTabBtn = document.createElement("button");
  archiveTabBtn.type = "button";
  archiveTabBtn.className = "hk-panel-tab-btn";
  archiveTabBtn.id = "hk-panel-tab-archive";
  archiveTabBtn.setAttribute("role", "tab");
  archiveTabBtn.setAttribute("aria-controls", "hk-panel-tabpanel-archive");
  archiveTabBtn.textContent = "存檔";
  archiveTabBtn.addEventListener("click", () => {
    if (highlightPanelState.activeTab === "archive") return;
    highlightPanelState.activeTab = "archive";
    applyHighlightPanelTabState();
    renderHighlightPanel().catch((error) =>
      console.debug("重新整理標註面板失敗", error)
    );
  });

  const searchTabBtn = document.createElement("button");
  searchTabBtn.type = "button";
  searchTabBtn.className = "hk-panel-tab-btn";
  searchTabBtn.id = "hk-panel-tab-search";
  searchTabBtn.setAttribute("role", "tab");
  searchTabBtn.setAttribute("aria-controls", "hk-panel-tabpanel-search");
  searchTabBtn.textContent = "搜尋";
  searchTabBtn.addEventListener("click", () => {
    if (highlightPanelState.activeTab === "search") return;
    highlightPanelState.activeTab = "search";
    applyHighlightPanelTabState();
    renderHighlightPanel().catch((error) =>
      console.debug("重新整理標註面板失敗", error)
    );
  });

  const aiTabBtn = document.createElement("button");
  aiTabBtn.type = "button";
  aiTabBtn.className = "hk-panel-tab-btn";
  aiTabBtn.id = "hk-panel-tab-ai-note";
  aiTabBtn.setAttribute("role", "tab");
  aiTabBtn.setAttribute("aria-controls", "hk-panel-tabpanel-ai-note");
  aiTabBtn.textContent = "AI 筆記";
  aiTabBtn.addEventListener("click", () => {
    if (highlightPanelState.activeTab === "ai-note") return;
    highlightPanelState.activeTab = "ai-note";
    applyHighlightPanelTabState();
    renderHighlightPanel().catch((error) =>
      console.debug("重新整理標註面板失敗", error)
    );
  });

  tabs.appendChild(pageTabBtn);
  tabs.appendChild(archiveTabBtn);
  tabs.appendChild(searchTabBtn);
  tabs.appendChild(aiTabBtn);

  const pageTabPanel = document.createElement("div");
  pageTabPanel.className = "hk-panel-tab-content hk-panel-tab-content-page";
  pageTabPanel.id = "hk-panel-tabpanel-page";
  pageTabPanel.setAttribute("role", "tabpanel");
  pageTabPanel.setAttribute("aria-labelledby", pageTabBtn.id);

  const archiveTabPanel = document.createElement("div");
  archiveTabPanel.className = "hk-panel-tab-content hk-panel-tab-content-archive";
  archiveTabPanel.id = "hk-panel-tabpanel-archive";
  archiveTabPanel.setAttribute("role", "tabpanel");
  archiveTabPanel.setAttribute("aria-labelledby", archiveTabBtn.id);

  const searchTabPanel = document.createElement("div");
  searchTabPanel.className = "hk-panel-tab-content hk-panel-tab-content-search";
  searchTabPanel.id = "hk-panel-tabpanel-search";
  searchTabPanel.setAttribute("role", "tabpanel");
  searchTabPanel.setAttribute("aria-labelledby", searchTabBtn.id);

  const aiTabPanel = document.createElement("div");
  aiTabPanel.className = "hk-panel-tab-content hk-panel-tab-content-ai";
  aiTabPanel.id = "hk-panel-tabpanel-ai-note";
  aiTabPanel.setAttribute("role", "tabpanel");
  aiTabPanel.setAttribute("aria-labelledby", aiTabBtn.id);

  const searchControls = document.createElement("div");
  searchControls.className = "hk-panel-controls";

  const searchWrapper = document.createElement("label");
  searchWrapper.className = "hk-panel-search";
  searchWrapper.textContent = "搜尋筆記";
  const searchInput = document.createElement("input");
  searchInput.type = "search";
  searchInput.className = "hk-panel-search-input";
  searchInput.placeholder = "輸入關鍵字或標籤";
  searchInput.addEventListener("input", (event) => {
    highlightPanelState.searchTerm = event.target.value ?? "";
    renderHighlightPanel();
  });
  searchWrapper.appendChild(searchInput);
  searchControls.appendChild(searchWrapper);

  const tagsContainer = document.createElement("div");
  tagsContainer.className = "hk-panel-tags";
  searchControls.appendChild(tagsContainer);

  const pageFilterSelect = document.createElement("select");
  pageFilterSelect.className = "hk-panel-search-select";
  pageFilterSelect.addEventListener("change", (event) => {
    const value = event.target.value;
    highlightPanelState.searchPageFilter =
      value === "__all__" ? null : value || null;
    renderHighlightPanel();
  });
  searchControls.appendChild(pageFilterSelect);

  const pageTagSection = document.createElement("div");
  pageTagSection.className = "hk-panel-page-tags";
  const pageTagHeader = document.createElement("div");
  pageTagHeader.className = "hk-panel-page-tags-header";
  const pageTagTitle = document.createElement("span");
  pageTagTitle.textContent = "頁面工具";
  pageTagHeader.appendChild(pageTagTitle);
  const pageTagHint = document.createElement("span");
  pageTagHint.className = "hk-panel-page-tags-hint";
  pageTagHeader.appendChild(pageTagHint);
  pageTagSection.appendChild(pageTagHeader);

  const exportActions = document.createElement("div");
  exportActions.className = "hk-panel-export-actions";
  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "hk-panel-export-btn";
  copyBtn.textContent = "複製內容";
  copyBtn.addEventListener("click", () => exportHighlights("copy"));
  const downloadBtn = document.createElement("button");
  downloadBtn.type = "button";
  downloadBtn.className = "hk-panel-export-btn";
  downloadBtn.textContent = "下載 JSON";
  downloadBtn.addEventListener("click", () => exportHighlights("download"));
  const importLabel = document.createElement("label");
  importLabel.className = "hk-panel-export-btn hk-panel-import-label";
  importLabel.textContent = "匯入 JSON";
  exportActions.appendChild(copyBtn);
  exportActions.appendChild(downloadBtn);
  exportActions.appendChild(importLabel);
  pageTagSection.appendChild(exportActions);

  const exportStatus = document.createElement("div");
  exportStatus.className = "hk-panel-export-status";
  exportStatus.setAttribute("role", "status");
  exportStatus.setAttribute("aria-live", "polite");
  pageTagSection.appendChild(exportStatus);

  const importInput = document.createElement("input");
  importInput.type = "file";
  importInput.accept = "application/json";
  importInput.className = "hk-panel-import-input";
  importInput.addEventListener("click", (event) => {
    event.target.value = "";
  });
  importInput.addEventListener("change", handleImportFileChange);
  importLabel.appendChild(importInput);

  const tagInputRow = document.createElement("div");
  tagInputRow.className = "hk-panel-tag-input-row";
  const tagInputWrapper = document.createElement("div");
  tagInputWrapper.className = "hk-panel-tag-input-wrapper";

  const tagInput = document.createElement("input");
  tagInput.type = "text";
  tagInput.className = "hk-panel-tag-input";
  tagInput.placeholder = "輸入標籤後按 Enter";

  const suggestionDropdown = document.createElement("div");
  suggestionDropdown.className = "hk-panel-tag-suggestions";
  suggestionDropdown.setAttribute("role", "listbox");
  suggestionDropdown.setAttribute("aria-label", "標籤建議");

  tagInputWrapper.appendChild(tagInput);
  tagInputWrapper.appendChild(suggestionDropdown);
  tagInputRow.appendChild(tagInputWrapper);

  const addTagBtn = document.createElement("button");
  addTagBtn.type = "button";
  addTagBtn.className = "hk-panel-tag-add";
  addTagBtn.textContent = "新增";
  tagInputRow.appendChild(addTagBtn);

  pageTagSection.appendChild(tagInputRow);

  const pageTagList = document.createElement("div");
  pageTagList.className = "hk-panel-page-tag-list";
  pageTagSection.appendChild(pageTagList);

  const pageArchiveSection = document.createElement("div");
  pageArchiveSection.className = "hk-panel-archive";
  const pageArchiveHint = document.createElement("p");
  pageArchiveHint.className = "hk-panel-archive-hint";
  pageArchiveHint.textContent = "本頁筆記";
  const pageArchiveActions = document.createElement("div");
  pageArchiveActions.className = "hk-panel-archive-actions";
  const archiveCopyBtn = document.createElement("button");
  archiveCopyBtn.type = "button";
  archiveCopyBtn.className = "hk-panel-export-btn";
  archiveCopyBtn.textContent = "複製文字";
  archiveCopyBtn.addEventListener("click", () => exportHighlights("copy"));
  const archiveDownloadBtn = document.createElement("button");
  archiveDownloadBtn.type = "button";
  archiveDownloadBtn.className = "hk-panel-export-btn";
  archiveDownloadBtn.textContent = "下載 JSON";
  archiveDownloadBtn.addEventListener("click", () => exportHighlights("download"));
  const archiveImportLabel = document.createElement("label");
  archiveImportLabel.className = "hk-panel-export-btn hk-panel-import-label";
  archiveImportLabel.textContent = "匯入 JSON";
  const archiveImportInput = document.createElement("input");
  archiveImportInput.type = "file";
  archiveImportInput.accept = "application/json";
  archiveImportInput.className = "hk-panel-import-input";
  archiveImportInput.addEventListener("click", (event) => {
    event.target.value = "";
  });
  archiveImportInput.addEventListener("change", handleImportFileChange);
  archiveImportLabel.appendChild(archiveImportInput);
  pageArchiveActions.appendChild(archiveCopyBtn);
  pageArchiveActions.appendChild(archiveDownloadBtn);
  pageArchiveActions.appendChild(archiveImportLabel);
  const archiveSection = document.createElement("div");
  archiveSection.className = "hk-panel-archive";
  const archiveHint = document.createElement("p");
  archiveHint.className = "hk-panel-archive-hint";
  archiveHint.textContent = "全部筆記";
  const archiveButtons = document.createElement("div");
  archiveButtons.className = "hk-panel-archive-actions";
  const archiveAllDownloadBtn = document.createElement("button");
  archiveAllDownloadBtn.type = "button";
  archiveAllDownloadBtn.className = "hk-panel-export-btn";
  archiveAllDownloadBtn.textContent = "下載全部筆記";
  archiveAllDownloadBtn.addEventListener("click", handleDownloadAllHighlights);
  const archiveAllImportLabel = document.createElement("label");
  archiveAllImportLabel.className = "hk-panel-export-btn hk-panel-import-label";
  archiveAllImportLabel.textContent = "匯入全部筆記";
  const archiveAllImportInput = document.createElement("input");
  archiveAllImportInput.type = "file";
  archiveAllImportInput.accept = "application/json";
  archiveAllImportInput.multiple = true;
  archiveAllImportInput.className = "hk-panel-import-input";
  archiveAllImportInput.addEventListener("click", (event) => {
    event.target.value = "";
  });
  archiveAllImportInput.addEventListener("change", handleBulkImportChange);
  archiveAllImportLabel.appendChild(archiveAllImportInput);
  archiveButtons.appendChild(archiveAllDownloadBtn);
  archiveButtons.appendChild(archiveAllImportLabel);
  const archiveStatus = document.createElement("div");
  archiveStatus.className = "hk-panel-export-status";
  archiveStatus.setAttribute("role", "status");
  archiveStatus.setAttribute("aria-live", "polite");
  pageArchiveSection.appendChild(pageArchiveHint);
  pageArchiveSection.appendChild(pageArchiveActions);
  archiveSection.appendChild(archiveHint);
  archiveSection.appendChild(archiveButtons);
  archiveSection.appendChild(archiveStatus);

  const aiSettingsSection = document.createElement("section");
  aiSettingsSection.className = "hk-panel-ai-settings";

  const providerField = document.createElement("div");
  providerField.className = "hk-panel-ai-field";
  const providerLabel = document.createElement("label");
  providerLabel.className = "hk-panel-ai-label";
  const providerSelectId = "hk-panel-ai-provider";
  providerLabel.setAttribute("for", providerSelectId);
  providerLabel.textContent = "服務提供者";
  const aiProviderSelect = document.createElement("select");
  aiProviderSelect.id = providerSelectId;
  aiProviderSelect.className = "hk-panel-ai-select";
  [
    ["openai", "OpenAI"],
    ["gemini", "Google Gemini"],
  ].forEach(([value, label]) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    aiProviderSelect.appendChild(option);
  });
  providerField.appendChild(providerLabel);
  providerField.appendChild(aiProviderSelect);
  aiSettingsSection.appendChild(providerField);

  const modelField = document.createElement("div");
  modelField.className = "hk-panel-ai-field";
  const modelLabel = document.createElement("label");
  const modelSelectId = "hk-panel-ai-model";
  modelLabel.setAttribute("for", modelSelectId);
  modelLabel.className = "hk-panel-ai-label";
  modelLabel.textContent = "模型";
  const aiModelSelect = document.createElement("select");
  aiModelSelect.id = modelSelectId;
  aiModelSelect.className = "hk-panel-ai-select";
  modelField.appendChild(modelLabel);
  modelField.appendChild(aiModelSelect);
  aiSettingsSection.appendChild(modelField);

  const aiKeyGroupsContainer = document.createElement("div");
  aiKeyGroupsContainer.className = "hk-panel-ai-key-groups";

  const openaiGroup = document.createElement("div");
  openaiGroup.className = "hk-panel-ai-key-group";
  openaiGroup.dataset.provider = "openai";
  const openaiLabel = document.createElement("label");
  const openaiInputId = "hk-panel-ai-openai-key";
  openaiLabel.setAttribute("for", openaiInputId);
  openaiLabel.textContent = "OpenAI API Key";
  const openaiInput = document.createElement("input");
  openaiInput.type = "password";
  openaiInput.id = openaiInputId;
  openaiInput.className = "hk-panel-ai-input";
  openaiInput.placeholder = "sk-...";
  openaiInput.autocomplete = "off";
  openaiGroup.appendChild(openaiLabel);
  openaiGroup.appendChild(openaiInput);
  openaiGroup.hidden = true;

  const geminiGroup = document.createElement("div");
  geminiGroup.className = "hk-panel-ai-key-group";
  geminiGroup.dataset.provider = "gemini";
  const geminiLabel = document.createElement("label");
  const geminiInputId = "hk-panel-ai-gemini-key";
  geminiLabel.setAttribute("for", geminiInputId);
  geminiLabel.textContent = "Gemini API Key";
  const geminiInput = document.createElement("input");
  geminiInput.type = "password";
  geminiInput.id = geminiInputId;
  geminiInput.className = "hk-panel-ai-input";
  geminiInput.placeholder = "AIza...";
  geminiInput.autocomplete = "off";
  geminiGroup.appendChild(geminiLabel);
  geminiGroup.appendChild(geminiInput);
  geminiGroup.hidden = true;

  aiKeyGroupsContainer.appendChild(openaiGroup);
  aiKeyGroupsContainer.appendChild(geminiGroup);
  aiSettingsSection.appendChild(aiKeyGroupsContainer);

  const promptField = document.createElement("div");
  promptField.className = "hk-panel-ai-field";
  const promptLabel = document.createElement("label");
  const promptTextareaId = "hk-panel-ai-prompt";
  promptLabel.setAttribute("for", promptTextareaId);
  promptLabel.className = "hk-panel-ai-label";
  promptLabel.textContent = "自訂 Prompt";
  const promptTextarea = document.createElement("textarea");
  promptTextarea.id = promptTextareaId;
  promptTextarea.className = "hk-panel-ai-textarea";
  promptTextarea.rows = 4;
  promptField.appendChild(promptLabel);
  promptField.appendChild(promptTextarea);
  aiSettingsSection.appendChild(promptField);

  const aiActions = document.createElement("div");
  aiActions.className = "hk-panel-ai-actions";
  const aiStatus = document.createElement("span");
  aiStatus.className = "hk-panel-ai-status";
  aiStatus.setAttribute("role", "status");
  aiStatus.setAttribute("aria-live", "polite");
  const aiGenerateBtn = document.createElement("button");
  aiGenerateBtn.type = "button";
  aiGenerateBtn.className = "hk-panel-ai-generate";
  aiGenerateBtn.textContent = "產生筆記";
  aiActions.appendChild(aiStatus);
  aiActions.appendChild(aiGenerateBtn);
  aiSettingsSection.appendChild(aiActions);

  const aiNoteSection = document.createElement("section");
  aiNoteSection.className = "hk-panel-ai-note";
  const aiNoteHeader = document.createElement("div");
  aiNoteHeader.className = "hk-panel-ai-note-header";
  const aiNoteTitle = document.createElement("span");
  aiNoteTitle.className = "hk-panel-ai-note-title";
  aiNoteTitle.textContent = "AI 筆記";
  const aiNoteMeta = document.createElement("span");
  aiNoteMeta.className = "hk-panel-ai-note-meta";
  const aiNoteCopyBtn = document.createElement("button");
  aiNoteCopyBtn.type = "button";
  aiNoteCopyBtn.className = "hk-panel-ai-note-copy";
  aiNoteCopyBtn.textContent = "複製";
  aiNoteCopyBtn.disabled = true;
  aiNoteHeader.appendChild(aiNoteTitle);
  aiNoteHeader.appendChild(aiNoteMeta);
  aiNoteHeader.appendChild(aiNoteCopyBtn);
  const aiNoteContent = document.createElement("pre");
  aiNoteContent.className = "hk-panel-ai-note-content";
  aiNoteContent.textContent = "";
  const aiNoteEmpty = document.createElement("div");
  aiNoteEmpty.className = "hk-panel-ai-note-empty";
  aiNoteEmpty.textContent = "尚未產生筆記。";
  aiNoteSection.appendChild(aiNoteHeader);
  aiNoteSection.appendChild(aiNoteContent);
  aiNoteSection.appendChild(aiNoteEmpty);

  const pageList = document.createElement("div");
  pageList.className = "hk-panel-list hk-panel-page-list";

  const pagePlaceholder = document.createElement("p");
  pagePlaceholder.className = "hk-panel-placeholder";
  pagePlaceholder.textContent = "本頁尚未建立標註。";

  const searchList = document.createElement("div");
  searchList.className = "hk-panel-list hk-panel-search-list";

  const searchPlaceholder = document.createElement("p");
  searchPlaceholder.className = "hk-panel-placeholder";
  searchPlaceholder.textContent = "目前沒有符合的筆記。";

  panel.appendChild(header);
  panel.appendChild(tabs);
  pageTabPanel.appendChild(pageTagSection);
  pageTabPanel.appendChild(pageList);
  pageTabPanel.appendChild(pagePlaceholder);
  searchTabPanel.appendChild(searchControls);
  searchTabPanel.appendChild(searchList);
  searchTabPanel.appendChild(searchPlaceholder);
  archiveTabPanel.appendChild(pageArchiveSection);
  archiveTabPanel.appendChild(archiveSection);
  aiTabPanel.appendChild(aiSettingsSection);
  aiTabPanel.appendChild(aiNoteSection);
  panel.appendChild(pageTabPanel);
  panel.appendChild(archiveTabPanel);
  panel.appendChild(searchTabPanel);
  panel.appendChild(aiTabPanel);

  highlightPanel = panel;
  highlightPanelEls = {
    container: panel,
    dragHandle: header,
    tabs,
    pageList,
    pagePlaceholder,
    searchList,
    searchPlaceholder,
    searchInput,
    tagsContainer,
    searchPageSelect: pageFilterSelect,
    pageTagSection,
    pageTagHint,
    copyBtn,
    downloadBtn,
    tagInput,
    addTagBtn,
    suggestionDropdown,
    pageTagList,
    exportStatus,
    archiveStatus,
    aiSettingsSection,
    aiNoteSection,
    aiNoteContent,
    aiNoteMeta,
    aiNoteCopyBtn,
    aiNoteEmpty,
    aiProviderSelect,
    aiModelSelect,
    aiPromptField: promptTextarea,
    aiOpenaiKeyInput: openaiInput,
    aiGeminiKeyInput: geminiInput,
    aiKeyGroups: [openaiGroup, geminiGroup],
    aiGenerateBtn,
    aiStatus,
    fontControls: {
      decrease: fontDecreaseBtn,
      increase: fontIncreaseBtn,
    },
    tabButtons: {
      page: pageTabBtn,
      archive: archiveTabBtn,
      search: searchTabBtn,
      ai: aiTabBtn,
    },
    tabPanels: {
      page: pageTabPanel,
      archive: archiveTabPanel,
      search: searchTabPanel,
      ai: aiTabPanel,
    },
  };

  aiProviderSelect.addEventListener("change", (event) => {
    const value = event.target.value === "gemini" ? "gemini" : "openai";
    aiSettings.provider = value;
    populateAiModelSelect();
    updateAiKeyVisibility();
    updateGenerateAvailability();
    persistAISettings();
  });

  aiModelSelect.addEventListener("change", (event) => {
    if (aiSettings.provider === "openai") {
      aiSettings.openaiModel = event.target.value;
    } else {
      aiSettings.geminiModel = event.target.value;
    }
    persistAISettings();
  });

  openaiInput.addEventListener("input", (event) => {
    aiSettings.openaiKey = event.target.value;
    updateGenerateAvailability();
    persistAISettings();
  });

  geminiInput.addEventListener("input", (event) => {
    aiSettings.geminiKey = event.target.value;
    updateGenerateAvailability();
    persistAISettings();
  });

  promptTextarea.addEventListener("input", (event) => {
    aiSettings.prompt = event.target.value;
    persistAISettings();
  });

  aiGenerateBtn.addEventListener("click", handleGenerateAiNote);

  document.body.appendChild(panel);
  applyHighlightPanelSideClasses();
  applyHighlightPanelTabState();
  updateExportButtonsState();
  applyAiSettingsToUI();
  applyHighlightPanelFontScale();
  setAiPanelStatus("");
  return panel;
};

const renderHighlightPanel = async () => {
  ensureHighlightPanel();
  if (!highlightPanelEls) return;

  highlightPanelState.activeKey = pageKey;

  await refreshHighlightPanelData();
  applyHighlightPanelSideClasses();
  await renderPageTagEditor();
  updateHighlightPanelTagFilters();
  updateSearchPageFilterOptions();

  const pageEntries = collectPanelEntries("current").sort(
    (a, b) => (b.createdAt ?? 0) - (a.createdAt ?? 0)
  );
  renderPageEntries(pageEntries);

  const searchTerm = highlightPanelState.searchTerm.trim().toLowerCase();
  const activeTag = highlightPanelState.activeTag;
  const searchEntries = collectPanelEntries("all", {
    pageFilter: highlightPanelState.searchPageFilter,
  })
    .filter((entry) => {
      const tags = Array.isArray(entry.pageTags) ? entry.pageTags : [];
      const matchesTag = !activeTag || tags.includes(activeTag);
      if (!matchesTag) return false;
      if (!searchTerm) return true;
      const haystacks = [entry.text, entry.note, entry.pageUrl, tags.join(" ")]
        .filter(Boolean)
        .map((value) => String(value).toLowerCase());
      return haystacks.some((text) => text.includes(searchTerm));
    })
    .sort((a, b) => (b.createdAt ?? 0) - (a.createdAt ?? 0));
  renderSearchEntries(searchEntries);

  if (highlightPanelEls?.searchInput) {
    highlightPanelEls.searchInput.value = highlightPanelState.searchTerm;
  }

  applyHighlightPanelTabState();

  const noteKey = pageKey;
  updateAiNoteSection(highlightPanelState.notesByPage[noteKey]);
};

const openHighlightPanel = async (options = {}) => {
  ensureHighlightPanel();
  if (panelPreferencesPromise) {
    try {
      await panelPreferencesPromise;
    } catch (error) {
      console.debug("載入面板設定失敗:", error);
    }
  }
  const { side, view } = options ?? {};
  if (typeof side === "string") {
    await setHighlightPanelSide(side);
  } else if (highlightPanelState.side !== highlightPanelPreferredSide) {
    await setHighlightPanelSide(highlightPanelPreferredSide, false);
  } else {
    applyHighlightPanelSideClasses();
  }

  if (view === "ai-note") {
    highlightPanelState.activeTab = "ai-note";
  } else if (view === "search") {
    highlightPanelState.activeTab = "search";
  } else if (view === "archive") {
    highlightPanelState.activeTab = "archive";
  } else if (view === "highlights" || view === "page") {
    highlightPanelState.activeTab = "page";
  }

  highlightPanelVisible = true;
  highlightPanel.style.display = "flex";
  highlightPanel.setAttribute("aria-hidden", "false");
  await renderHighlightPanel();
  return true;
};

const closeHighlightPanel = () => {
  if (!highlightPanel) return false;
  highlightPanelVisible = false;
  highlightPanel.style.display = "none";
  highlightPanel.setAttribute("aria-hidden", "true");
  return true;
};

const toggleHighlightPanel = async () => {
  if (highlightPanelVisible) {
    closeHighlightPanel();
    return false;
  }
  await openHighlightPanel();
  return true;
};

const refreshHighlightPanelIfVisible = async () => {
  if (!highlightPanelVisible) return;
  await renderHighlightPanel();
};

const applyColorChange = async (color) => {
  if (!activeHighlight || !activeHighlightId) return;
  const nextColor = toHexColor(color || DEFAULT_COLOR);
  setHighlightMetadata(activeHighlight, {
    color: nextColor,
    note: activeHighlight.dataset.hkNote ?? "",
  });
  currentColor = nextColor;
  if (highlightMenuEls?.colorInput) {
    highlightMenuEls.colorInput.value = nextColor;
  }
  if (!colorPalette.includes(nextColor)) {
    await persistColorPalette([...colorPalette, nextColor]);
  }
  try {
    await updateHighlightEntry(activeHighlightId, { color: nextColor });
    try {
      await chrome.storage?.local.set({ hkLastColor: nextColor });
    } catch (storageError) {
      console.debug("更新預設 highlight 顏色失敗", storageError);
    }
    setHighlightMenuStatus("顏色已更新");
    await refreshHighlightPanelIfVisible();
  } catch (error) {
    console.debug("更新 highlight 顏色失敗", error);
    setHighlightMenuStatus("無法更新顏色", true);
  }
};

const handleSaveNote = async () => {
  if (!activeHighlight || !activeHighlightId || !highlightMenuEls?.noteField) {
    return;
  }
  const note = highlightMenuEls.noteField.value ?? "";
  const trimmedNote = note.trim();
  const normalizedColor =
    activeHighlight.dataset.hkColor ||
    highlightMenuEls.colorInput?.value ||
    DEFAULT_COLOR;
  setHighlightMetadata(activeHighlight, { color: normalizedColor, note: trimmedNote });
  try {
    await updateHighlightEntry(activeHighlightId, { note: trimmedNote });
    if (highlightMenuEls.noteField) {
      highlightMenuEls.noteField.value = trimmedNote;
    }
    setHighlightMenuStatus("註解已儲存");
    await refreshHighlightPanelIfVisible();
  } catch (error) {
    console.debug("儲存 highlight 註解失敗", error);
    setHighlightMenuStatus("無法儲存註解", true);
  }
};

const handleDeleteHighlight = async () => {
  if (!activeHighlight || !activeHighlightId) return;
  const elementToRemove = activeHighlight;
  const idToRemove = activeHighlightId;
  closeHighlightMenu();
  unwrapHighlightElement(elementToRemove);
  try {
    await deleteHighlightEntry(idToRemove);
    await refreshHighlightPanelIfVisible();
    updateTagSuggestionDropdown();
  } catch (error) {
    console.debug("刪除 highlight 失敗", error);
  }
};

const ensureHighlightMenu = () => {
  if (highlightMenu) return highlightMenu;

  const container = document.createElement("div");
  container.id = HIGHLIGHT_MENU_ID;
  container.className = "hk-highlight-menu";
  container.style.display = "none";

  const header = document.createElement("div");
  header.className = "hk-menu-header";
  header.textContent = "標註設定";
  container.appendChild(header);

  const colorWrapper = document.createElement("div");
  colorWrapper.className = "hk-menu-section hk-menu-colors";

  const colorLabel = document.createElement("label");
  colorLabel.className = "hk-menu-label";
  colorLabel.textContent = "顏色";

  const colorInput = document.createElement("input");
  colorInput.type = "color";
  colorInput.className = "hk-menu-color-input";
  colorInput.value = DEFAULT_COLOR;
  colorLabel.appendChild(colorInput);

  const swatchGroup = document.createElement("div");
  swatchGroup.className = "hk-menu-swatches";
  colorWrapper.appendChild(colorLabel);
  colorWrapper.appendChild(swatchGroup);
  container.appendChild(colorWrapper);

  const noteWrapper = document.createElement("div");
  noteWrapper.className = "hk-menu-section";
  const noteLabel = document.createElement("label");
  noteLabel.className = "hk-menu-label";
  noteLabel.textContent = "註解";
  const noteField = document.createElement("textarea");
  noteField.className = "hk-menu-note";
  noteField.rows = 6;
  noteField.placeholder = "輸入註解...";
  noteWrapper.appendChild(noteLabel);
  noteWrapper.appendChild(noteField);
  container.appendChild(noteWrapper);

  const actions = document.createElement("div");
  actions.className = "hk-menu-actions";
  const saveNoteBtn = document.createElement("button");
  saveNoteBtn.type = "button";
  saveNoteBtn.className = "hk-menu-btn hk-menu-btn-primary";
  saveNoteBtn.textContent = "儲存註解";
  saveNoteBtn.addEventListener("click", handleSaveNote);

  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "hk-menu-btn hk-menu-btn-danger";
  deleteBtn.textContent = "刪除標註";
  deleteBtn.addEventListener("click", handleDeleteHighlight);

  actions.appendChild(saveNoteBtn);
  actions.appendChild(deleteBtn);
  container.appendChild(actions);

  const status = document.createElement("div");
  status.className = "hk-menu-status";
  container.appendChild(status);

  colorInput.addEventListener("input", (event) => {
    const next = event.target.value;
    applyColorChange(next);
  });

  container.addEventListener("mousedown", (event) => {
    event.stopPropagation();
  });

  highlightMenu = container;
  highlightMenuEls = {
    container,
    colorInput,
    noteField,
    status,
    swatchGroup,
  };

  renderHighlightMenuSwatches();

  document.body.appendChild(container);
  return container;
};

const positionHighlightMenu = (highlightEl) => {
  const menu = ensureHighlightMenu();
  if (!highlightEl) return;
  const margin = 12;
  menu.style.visibility = "hidden";
  menu.style.opacity = "0";
  menu.style.display = "flex";

  window.requestAnimationFrame(() => {
    const rect = highlightEl.getBoundingClientRect();
    const menuRect = menu.getBoundingClientRect();
    let top = rect.bottom + margin;
    let left = rect.left;

    if (left + menuRect.width > window.innerWidth - margin) {
      left = window.innerWidth - menuRect.width - margin;
    }
    if (left < margin) {
      left = margin;
    }

    if (top + menuRect.height > window.innerHeight - margin) {
      top = rect.top - menuRect.height - margin;
      if (top < margin) {
        top = window.innerHeight - menuRect.height - margin;
      }
    }

    menu.style.top = `${Math.max(margin, top)}px`;
    menu.style.left = `${Math.max(margin, left)}px`;
    menu.style.visibility = "visible";
    menu.style.opacity = "1";
  });
};

const openHighlightMenu = async (highlightEl) => {
  if (!highlightEl) return;
  const id = highlightEl.getAttribute(HIGHLIGHT_ATTR);
  if (!id) return;

  await refreshPaletteFromStorage();

  ensureHighlightMenu();
  renderHighlightMenuSwatches();

  activeHighlight = highlightEl;
  activeHighlightId = id;
  setHighlightMenuStatus("");

  try {
    const record = await findHighlightEntry(id);
    const color =
      record?.color ||
      highlightEl.dataset.hkColor ||
      highlightEl.getAttribute("data-highlight-color") ||
      DEFAULT_COLOR;
    const note =
      record?.note ??
      highlightEl.dataset.hkNote ??
      highlightEl.getAttribute("data-highlight-note") ??
      "";
    if (highlightMenuEls?.colorInput) {
      highlightMenuEls.colorInput.value = toHexColor(color);
    }
    if (highlightMenuEls?.noteField) {
      highlightMenuEls.noteField.value = note;
    }
  } catch (error) {
    console.debug("讀取 highlight 設定失敗", error);
  }

  if (highlightMenu) {
    highlightMenu.style.display = "flex";
    positionHighlightMenu(highlightEl);
  }
};

const updateCurrentColor = async () => {
  if (!storage) return;
  try {
    const stored = await storage.get("hkLastColor");
    const color = stored?.hkLastColor;
    if (typeof color === "string" && color.trim()) {
      currentColor = toHexColor(color);
    }
  } catch (error) {
    console.debug("讀取預設顏色失敗", error);
  }
};

loadPalette();
updateCurrentColor();
panelPreferencesPromise = loadPanelPreferences();
loadAISettings();

chrome.storage?.onChanged.addListener((changes, areaName) => {
  if (areaName !== "local") return;
  if (changes.hkLastColor) {
    const nextColor = changes.hkLastColor.newValue;
    if (typeof nextColor === "string" && nextColor.trim()) {
      currentColor = toHexColor(nextColor);
    } else {
      currentColor = DEFAULT_COLOR;
    }
  }
  if (changes.hkColorPalette) {
    setColorPaletteState(changes.hkColorPalette.newValue);
  }
  if (changes.hkPanelSide) {
    const nextSide =
      changes.hkPanelSide.newValue === "left" ? "left" : "right";
    highlightPanelState.side = nextSide;
    highlightPanelPreferredSide = nextSide;
    applyHighlightPanelSideClasses();
  }
  if (changes[HIGHLIGHT_PANEL_POSITION_KEY]) {
    const nextPosition = changes[HIGHLIGHT_PANEL_POSITION_KEY].newValue;
    highlightPanelState.position = isValidPanelPosition(nextPosition)
      ? nextPosition
      : null;
    applyHighlightPanelSideClasses();
  }
  if (changes[HIGHLIGHT_PANEL_FONT_SCALE_KEY]) {
    const nextScale = Number(changes[HIGHLIGHT_PANEL_FONT_SCALE_KEY].newValue);
    if (Number.isFinite(nextScale)) {
      highlightPanelState.fontScale = clampPanelFontScale(nextScale);
      applyHighlightPanelFontScale();
    }
  }
  if (changes.hkAISettings) {
    aiSettings = {
      ...aiSettings,
      ...(changes.hkAISettings.newValue ?? {}),
    };
    applyAiSettingsToUI();
  }
  if (changes.hkGeneratedNotes) {
    highlightPanelState.notesByPage = changes.hkGeneratedNotes.newValue ?? {};
    if (highlightPanelVisible) {
      updateAiNoteSection(highlightPanelState.notesByPage[pageKey]);
    }
  }
  if (highlightPanelVisible) {
    const affectedKeys = Object.keys(changes);
    const shouldRefresh =
      affectedKeys.includes(pageKey) ||
      affectedKeys.some((key) => isValidPageKey(key));
    if (shouldRefresh) {
      renderHighlightPanel().catch((error) =>
        console.debug("重新整理標註面板失敗", error)
      );
    }
  }
});

const isEditableElement = (node) => {
  if (!(node instanceof HTMLElement)) return false;
  const tag = node.tagName.toLowerCase();
  return (
    node.isContentEditable ||
    tag === "input" ||
    tag === "textarea" ||
    tag === "select"
  );
};

const getNodeXPath = (node) => {
  if (!node || node === document) {
    return "";
  }

  const parts = [];
  let currentNode = node;

  while (currentNode && currentNode !== document) {
    let index = 0;
    let sibling = currentNode;

    if (currentNode.nodeType === Node.TEXT_NODE) {
      while (sibling.previousSibling) {
        sibling = sibling.previousSibling;
        if (sibling.nodeType === Node.TEXT_NODE) {
          index += 1;
        }
      }
      parts.unshift(`text()[${index + 1}]`);
      currentNode = currentNode.parentNode;
    } else if (currentNode.nodeType === Node.ELEMENT_NODE) {
      while (sibling.previousSibling) {
        sibling = sibling.previousSibling;
        if (
          sibling.nodeType === Node.ELEMENT_NODE &&
          sibling.nodeName === currentNode.nodeName
        ) {
          index += 1;
        }
      }
      const tagName = currentNode.nodeName.toLowerCase();
      parts.unshift(`${tagName}[${index + 1}]`);
      currentNode = currentNode.parentNode;
    } else {
      currentNode = currentNode.parentNode;
    }
  }

  return `/${parts.join("/")}`;
};

const getNodeFromXPath = (xpath) => {
  if (!xpath) return null;
  const result = document.evaluate(
    xpath,
    document,
    null,
    XPathResult.FIRST_ORDERED_NODE_TYPE,
    null
  );
  return result.singleNodeValue;
};

const getNodeLength = (node) => {
  if (!node) return 0;
  if (node.nodeType === Node.TEXT_NODE) {
    return node.nodeValue?.length ?? 0;
  }
  return node.childNodes.length;
};

const clampOffset = (node, offset) => {
  const max = getNodeLength(node);
  return Math.min(Math.max(offset, 0), max);
};

const cssEscape = (value) => {
  if (typeof value !== "string") return "";
  if (typeof CSS !== "undefined" && typeof CSS.escape === "function") {
    return CSS.escape(value);
  }
  return value.replace(/([^\w-])/g, "\\$1");
};

const getSiblingIndexOfType = (element) => {
  if (!element || !element.parentElement) return 1;
  let index = 1;
  let sibling = element.previousElementSibling;
  while (sibling) {
    if (sibling.tagName === element.tagName) {
      index += 1;
    }
    sibling = sibling.previousElementSibling;
  }
  return index;
};

const getCssSelector = (element) => {
  if (!element || element.nodeType !== Node.ELEMENT_NODE) return null;
  const parts = [];
  let current = element;
  let depth = 0;
  while (current && current.nodeType === Node.ELEMENT_NODE && depth < MAX_SELECTOR_DEPTH) {
    let selector = current.tagName.toLowerCase();
    if (current.id) {
      selector += `#${cssEscape(current.id)}`;
      parts.unshift(selector);
      break;
    }
    if (current.classList && current.classList.length) {
      const classes = Array.from(current.classList)
        .slice(0, 2)
        .map((cls) => `.${cssEscape(cls)}`)
        .join("");
      selector += classes;
    }
    const index = getSiblingIndexOfType(current);
    if (index > 1) {
      selector += `:nth-of-type(${index})`;
    }
    parts.unshift(selector);
    current = current.parentElement;
    depth += 1;
  }
  if (!parts.length && element.id) {
    return `${element.tagName.toLowerCase()}#${cssEscape(element.id)}`;
  }
  return parts.join(" > ");
};

const getAnchorElement = (node) => {
  if (!node) return document.body;
  if (node.nodeType === Node.ELEMENT_NODE) return node;
  let current = node.parentElement || node.parentNode;
  let depth = 0;
  while (current && depth < MAX_SELECTOR_DEPTH) {
    if (
      current.id ||
      (current.classList && current.classList.length) ||
      !current.parentElement
    ) {
      return current;
    }
    current = current.parentElement;
    depth += 1;
  }
  return current || document.body;
};

const shouldSkipTextNode = (node) => {
  if (!node || node.nodeType !== Node.TEXT_NODE) return true;
  const parentTag = node.parentNode?.nodeName;
  return TEXT_PARENT_SKIP_TAGS.has(parentTag);
};

const getNextNodeInDocument = (node) => {
  if (!node) return null;
  if (node.firstChild) {
    return node.firstChild;
  }
  let current = node;
  while (current) {
    if (current.nextSibling) {
      return current.nextSibling;
    }
    current = current.parentNode;
  }
  return null;
};

const getPreviousNodeInDocument = (node) => {
  if (!node) return null;
  if (node.previousSibling) {
    let current = node.previousSibling;
    while (current && current.lastChild) {
      current = current.lastChild;
    }
    return current;
  }
  return node.parentNode;
};

const getNextTextNode = (node) => {
  let current = getNextNodeInDocument(node);
  while (current) {
    if (current.nodeType === Node.TEXT_NODE && !shouldSkipTextNode(current)) {
      return current;
    }
    current = getNextNodeInDocument(current);
  }
  return null;
};

const getPreviousTextNode = (node) => {
  let current = getPreviousNodeInDocument(node);
  while (current) {
    if (current.nodeType === Node.TEXT_NODE && !shouldSkipTextNode(current)) {
      return current;
    }
    current = getPreviousNodeInDocument(current);
  }
  return null;
};

const findTextNodeInSubtree = (root, forward = true) => {
  if (!root) return null;
  if (root.nodeType === Node.TEXT_NODE && !shouldSkipTextNode(root)) {
    return root;
  }
  const walker = document.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT,
    {
      acceptNode(node) {
        return shouldSkipTextNode(node)
          ? NodeFilter.FILTER_REJECT
          : NodeFilter.FILTER_ACCEPT;
      },
    },
    false
  );
  if (forward) {
    return walker.nextNode();
  }
  let result = null;
  let current = walker.nextNode();
  while (current) {
    result = current;
    current = walker.nextNode();
  }
  return result;
};

const resolveTextNodeFromBoundary = (container, offset, searchForward) => {
  if (!container) return null;
  if (container.nodeType === Node.TEXT_NODE) {
    return {
      node: container,
      offset: clampOffset(container, offset),
    };
  }
  const childNodes = container.childNodes;
  if (childNodes && childNodes.length) {
    if (searchForward) {
      for (let i = offset; i < childNodes.length; i += 1) {
        const candidate = findTextNodeInSubtree(childNodes[i], true);
        if (candidate) {
          return { node: candidate, offset: 0 };
        }
      }
    } else {
      for (let i = Math.min(offset - 1, childNodes.length - 1); i >= 0; i -= 1) {
        const candidate = findTextNodeInSubtree(childNodes[i], false);
        if (candidate) {
          return {
            node: candidate,
            offset: candidate.nodeValue?.length ?? 0,
          };
        }
      }
    }
  }
  const fallback = searchForward
    ? getNextTextNode(container)
    : getPreviousTextNode(container);
  if (fallback) {
    return {
      node: fallback,
      offset:
        searchForward && fallback.nodeType === Node.TEXT_NODE
          ? 0
          : fallback.nodeValue?.length ?? 0,
    };
  }
  return null;
};

const getRangeBoundaryInfo = (range, isStart) => {
  if (!range) return { node: null, offset: 0 };
  const container = isStart ? range.startContainer : range.endContainer;
  const offset = isStart ? range.startOffset : range.endOffset;
  const resolved = resolveTextNodeFromBoundary(container, offset, isStart);
  if (resolved) return resolved;
  if (container && container.nodeType === Node.TEXT_NODE) {
    return { node: container, offset: clampOffset(container, offset) };
  }
  return { node: null, offset: 0 };
};

const collectContextText = (node, offset, direction, maxLength = TEXT_CONTEXT_CHARS) => {
  if (!node || maxLength <= 0) return "";
  if (node.nodeType !== Node.TEXT_NODE) return "";
  let remaining = maxLength;
  let result = "";
  let currentNode = node;
  let currentOffset = clampOffset(node, offset);
  while (currentNode && remaining > 0) {
    if (!shouldSkipTextNode(currentNode)) {
      const value = currentNode.nodeValue || "";
      if (direction < 0) {
        const sliceEnd = currentNode === node ? currentOffset : value.length;
        const sliceStart = Math.max(0, sliceEnd - remaining);
        result = value.slice(sliceStart, sliceEnd) + result;
        remaining -= sliceEnd - sliceStart;
      } else {
        const sliceStart = currentNode === node ? currentOffset : 0;
        const sliceEnd = Math.min(value.length, sliceStart + remaining);
        result += value.slice(sliceStart, sliceEnd);
        remaining -= sliceEnd - sliceStart;
      }
    }
    if (remaining <= 0) break;
    currentNode =
      direction < 0
        ? getPreviousTextNode(currentNode)
        : getNextTextNode(currentNode);
    if (currentNode && currentNode.nodeType === Node.TEXT_NODE) {
      currentOffset = direction < 0 ? currentNode.nodeValue?.length ?? 0 : 0;
    } else {
      currentOffset = 0;
    }
  }
  return direction < 0 ? result.slice(-maxLength) : result.slice(0, maxLength);
};

const getElementTextOffset = (element, boundaryNode, boundaryOffset) => {
  if (!element || !boundaryNode) return null;
  const walker = document.createTreeWalker(
    element,
    NodeFilter.SHOW_TEXT,
    {
      acceptNode(node) {
        return shouldSkipTextNode(node)
          ? NodeFilter.FILTER_REJECT
          : NodeFilter.FILTER_ACCEPT;
      },
    },
    false
  );
  let accumulated = 0;
  let current = walker.nextNode();
  let lastNode = null;
  while (current) {
    const length = current.nodeValue?.length ?? 0;
    if (current === boundaryNode) {
      return accumulated + Math.min(boundaryOffset, length);
    }
    accumulated += length;
    lastNode = current;
    current = walker.nextNode();
  }
  if (lastNode === boundaryNode) {
    const length = lastNode.nodeValue?.length ?? 0;
    return accumulated + Math.min(boundaryOffset, length);
  }
  return null;
};

const buildBoundaryAnchor = (boundary) => {
  if (!boundary?.node || boundary.node.nodeType !== Node.TEXT_NODE) return null;
  const anchorElement = getAnchorElement(boundary.node);
  if (!anchorElement) return null;
  const selector = getCssSelector(anchorElement);
  if (!selector) return null;
  const textOffset = getElementTextOffset(anchorElement, boundary.node, boundary.offset);
  return {
    css: selector,
    textOffset,
  };
};

const buildRangeAnchors = (range) => {
  try {
    const startBoundary = getRangeBoundaryInfo(range, true);
    const endBoundary = getRangeBoundaryInfo(range, false);
    return {
      version: 1,
      createdAt: Date.now(),
      start: buildBoundaryAnchor(startBoundary),
      end: buildBoundaryAnchor(endBoundary),
      quote: {
        exact: range.toString(),
        prefix: startBoundary.node
          ? collectContextText(startBoundary.node, startBoundary.offset, -1)
          : "",
        suffix: endBoundary.node
          ? collectContextText(endBoundary.node, endBoundary.offset, 1)
          : "",
      },
    };
  } catch (error) {
    console.debug("建立標註 anchors 失敗", error);
    return null;
  }
};

const serializeRange = (range) => {
  const anchors = buildRangeAnchors(range);
  return {
    startXPath: getNodeXPath(range.startContainer),
    startOffset: range.startOffset,
    endXPath: getNodeXPath(range.endContainer),
    endOffset: range.endOffset,
    text: range.toString(),
    anchors: anchors ?? undefined,
  };
};

const deserializeRange = (data) => {
  if (!data) return null;
  const startNode = getNodeFromXPath(data.startXPath);
  const endNode = getNodeFromXPath(data.endXPath);

  if (!startNode || !endNode) return null;

  const range = document.createRange();
  range.setStart(startNode, clampOffset(startNode, data.startOffset));
  range.setEnd(endNode, clampOffset(endNode, data.endOffset));
  return range;
};

const moveTextBoundary = (node, offset, distance) => {
  if (!node || node.nodeType !== Node.TEXT_NODE) return null;
  let remaining = Math.abs(distance);
  let currentNode = node;
  let currentOffset = clampOffset(node, offset);
  if (distance >= 0) {
    while (currentNode) {
      const value = currentNode.nodeValue ?? "";
      const available = value.length - currentOffset;
      if (remaining <= available) {
        return {
          node: currentNode,
          offset: currentOffset + remaining,
        };
      }
      remaining -= available;
      currentNode = getNextTextNode(currentNode);
      currentOffset = 0;
    }
    return null;
  }
  while (currentNode) {
    if (remaining <= currentOffset) {
      return {
        node: currentNode,
        offset: currentOffset - remaining,
      };
    }
    remaining -= currentOffset;
    currentNode = getPreviousTextNode(currentNode);
    if (!currentNode) break;
    currentOffset = currentNode.nodeValue?.length ?? 0;
  }
  return null;
};

const locateTextPosition = (nodes, index) => {
  if (!Array.isArray(nodes)) return null;
  for (let i = 0; i < nodes.length; i += 1) {
    const entry = nodes[i];
    if (index < entry.start) break;
    if (index <= entry.end) {
      return {
        node: entry.node,
        offset: clampOffset(entry.node, index - entry.start),
      };
    }
  }
  const lastEntry = nodes[nodes.length - 1];
  if (lastEntry && index >= lastEntry.end) {
    return {
      node: lastEntry.node,
      offset: clampOffset(lastEntry.node, lastEntry.node.nodeValue?.length ?? 0),
    };
  }
  return null;
};

const buildDocumentTextIndex = () => {
  if (!document?.body) return null;
  const walker = document.createTreeWalker(
    document.body,
    NodeFilter.SHOW_TEXT,
    {
      acceptNode(node) {
        return shouldSkipTextNode(node)
          ? NodeFilter.FILTER_REJECT
          : NodeFilter.FILTER_ACCEPT;
      },
    },
    false
  );
  const nodes = [];
  let text = "";
  let current = walker.nextNode();
  while (current) {
    const value = current.nodeValue ?? "";
    if (value) {
      const start = text.length;
      text += value;
      nodes.push({ node: current, start, end: text.length });
    }
    current = walker.nextNode();
  }
  return { text, nodes };
};

const scanForQuoteMatch = (haystack, needle, prefix, suffix) => {
  if (!haystack || !needle) return null;
  let fromIndex = 0;
  while (fromIndex <= haystack.length) {
    const index = haystack.indexOf(needle, fromIndex);
    if (index === -1) break;
    const before = prefix
      ? haystack.slice(Math.max(0, index - prefix.length), index)
      : "";
    const after = suffix
      ? haystack.slice(index + needle.length, index + needle.length + suffix.length)
      : "";
    const prefixOk = !prefix || before.endsWith(prefix);
    const suffixOk = !suffix || after.startsWith(suffix);
    if (prefixOk && suffixOk) {
      return index;
    }
    fromIndex = index + 1;
  }
  return null;
};

const findQuoteMatchIndex = (haystack, needle, prefix, suffix) => {
  const attempts = [
    { prefix, suffix },
    { prefix, suffix: "" },
    { prefix: "", suffix },
    { prefix: "", suffix: "" },
  ];
  for (const attempt of attempts) {
    const index = scanForQuoteMatch(
      haystack,
      needle,
      attempt.prefix || "",
      attempt.suffix || ""
    );
    if (index !== null && index !== undefined) {
      return index;
    }
  }
  return null;
};

const resolveBoundaryFromCssAnchor = (anchor) => {
  if (!anchor?.css) return null;
  const targetOffset =
    typeof anchor.textOffset === "number" ? Math.max(anchor.textOffset, 0) : null;
  if (targetOffset === null) return null;
  let element = null;
  try {
    element = document.querySelector(anchor.css);
  } catch (_error) {
    element = null;
  }
  if (!element) return null;
  const walker = document.createTreeWalker(
    element,
    NodeFilter.SHOW_TEXT,
    {
      acceptNode(node) {
        return shouldSkipTextNode(node)
          ? NodeFilter.FILTER_REJECT
          : NodeFilter.FILTER_ACCEPT;
      },
    },
    false
  );
  let accumulated = 0;
  let current = walker.nextNode();
  let lastNode = null;
  while (current) {
    const value = current.nodeValue ?? "";
    const length = value.length;
    if (accumulated + length >= targetOffset) {
      const offset = Math.max(0, Math.min(length, targetOffset - accumulated));
      return {
        node: current,
        offset,
      };
    }
    accumulated += length;
    lastNode = current;
    current = walker.nextNode();
  }
  if (lastNode && targetOffset === accumulated) {
    return {
      node: lastNode,
      offset: lastNode.nodeValue?.length ?? 0,
    };
  }
  return null;
};

const resolveRangeFromCssAnchors = (anchors, textLength) => {
  if (!anchors) return null;
  const startBoundary = anchors.start ? resolveBoundaryFromCssAnchor(anchors.start) : null;
  const endBoundary = anchors.end ? resolveBoundaryFromCssAnchor(anchors.end) : null;
  const normalizedLength = Math.max(
    textLength || 0,
    typeof anchors.quote?.exact === "string" ? anchors.quote.exact.length : 0
  );
  if (startBoundary && endBoundary) {
    try {
      const range = document.createRange();
      range.setStart(startBoundary.node, startBoundary.offset);
      range.setEnd(endBoundary.node, endBoundary.offset);
      if (!range.collapsed || normalizedLength === 0) {
        return range;
      }
    } catch (_error) {
      return null;
    }
  }
  if (startBoundary && !endBoundary && normalizedLength > 0) {
    const derivedEnd = moveTextBoundary(
      startBoundary.node,
      startBoundary.offset,
      normalizedLength
    );
    if (derivedEnd) {
      const range = document.createRange();
      range.setStart(startBoundary.node, startBoundary.offset);
      range.setEnd(derivedEnd.node, derivedEnd.offset);
      if (!range.collapsed) return range;
    }
  }
  if (!startBoundary && endBoundary && normalizedLength > 0) {
    const derivedStart = moveTextBoundary(
      endBoundary.node,
      endBoundary.offset,
      -normalizedLength
    );
    if (derivedStart) {
      const range = document.createRange();
      range.setStart(derivedStart.node, derivedStart.offset);
      range.setEnd(endBoundary.node, endBoundary.offset);
      if (!range.collapsed) return range;
    }
  }
  return null;
};

const resolveRangeFromTextAnchors = (anchors) => {
  const quote = anchors?.quote;
  const exact = typeof quote?.exact === "string" ? quote.exact : "";
  if (!exact) return null;
  const index = buildDocumentTextIndex();
  if (!index || !index.text) return null;
  const matchIndex = findQuoteMatchIndex(
    index.text,
    exact,
    typeof quote?.prefix === "string" ? quote.prefix : "",
    typeof quote?.suffix === "string" ? quote.suffix : ""
  );
  if (matchIndex === null || matchIndex === undefined) return null;
  const startPosition = locateTextPosition(index.nodes, matchIndex);
  const endPosition = locateTextPosition(index.nodes, matchIndex + exact.length);
  if (!startPosition || !endPosition) return null;
  const range = document.createRange();
  range.setStart(startPosition.node, startPosition.offset);
  range.setEnd(endPosition.node, endPosition.offset);
  return range;
};

const resolveRangeSnapshot = (snapshot) => {
  if (!snapshot) {
    return { range: null, snapshot: null, updated: false };
  }
  const directRange = deserializeRange(snapshot);
  if (directRange) {
    return { range: directRange, snapshot, updated: false };
  }
  const cssRange = resolveRangeFromCssAnchors(snapshot.anchors, snapshot.text?.length);
  if (cssRange) {
    const normalized = serializeRange(cssRange.cloneRange());
    return { range: cssRange, snapshot: normalized, updated: true };
  }
  const textRange = resolveRangeFromTextAnchors(snapshot.anchors);
  if (textRange) {
    const normalized = serializeRange(textRange.cloneRange());
    return { range: textRange, snapshot: normalized, updated: true };
  }
  return { range: null, snapshot, updated: false };
};

const createHighlightElement = (color, id) => {
  const mark = document.createElement("mark");
  mark.className = HIGHLIGHT_CLASS;
  mark.setAttribute(HIGHLIGHT_ATTR, id);
  mark.style.backgroundColor = color;
  mark.style.setProperty("--hk-highlight-color", color);
  mark.style.padding = "0";
  mark.style.margin = "0";
  mark.style.color = "inherit";
  return mark;
};

const setHighlightMetadata = (element, { color, note }) => {
  if (!element) return;
  if (color) {
    element.style.backgroundColor = color;
    element.style.setProperty("--hk-highlight-color", color);
    if (element.dataset) {
      element.dataset.hkColor = color;
    }
    element.setAttribute("data-highlight-color", color);
  }
  if (typeof note === "string") {
    const trimmed = note.trim();
    if (trimmed) {
      element.dataset.hkNote = trimmed;
      element.setAttribute("data-highlight-note", trimmed);
      element.setAttribute("title", trimmed);
    } else {
      delete element.dataset.hkNote;
      element.removeAttribute("data-highlight-note");
      element.removeAttribute("title");
    }
  }
};

const unwrapHighlightElement = (highlightEl) => {
  if (!highlightEl) return;
  const parent = highlightEl.parentNode;
  if (!parent) {
    highlightEl.remove();
    return;
  }
  while (highlightEl.firstChild) {
    parent.insertBefore(highlightEl.firstChild, highlightEl);
  }
  parent.removeChild(highlightEl);
  parent.normalize();
};

const wrapRangeWithHighlight = (range, color, id) => {
  const highlightEl = createHighlightElement(color, id);
  const extracted = range.extractContents();
  highlightEl.appendChild(extracted);
  range.insertNode(highlightEl);
  return highlightEl;
};

const getStoredHighlights = async (key = pageKey) => {
  if (!storage) return [];
  const existing = await storage.get(key);
  return existing[key] ?? [];
};

const setStoredHighlights = async (highlights, key = pageKey) => {
  if (!storage) return;
  await storage.set({ [key]: highlights });
};

const getPageMeta = async () => {
  if (!storage) return {};
  const stored = await storage.get(PAGE_META_KEY);
  return stored[PAGE_META_KEY] ?? {};
};

const setPageMeta = async (meta) => {
  if (!storage) return;
  await storage.set({ [PAGE_META_KEY]: meta });
};

const ensurePageMetaTitle = async (url, title) => {
  if (!storage) return;
  if (typeof title !== "string" || !title.trim()) return;
  try {
    const meta = await getPageMeta();
    const existing = meta[url] ?? {};
    if (existing.title === title) return;
    meta[url] = {
      ...existing,
      title: title.trim(),
      updatedAt: Date.now(),
    };
    await setPageMeta(meta);
    if (highlightPanelState.pageMeta[url]) {
      highlightPanelState.pageMeta[url] = {
        ...highlightPanelState.pageMeta[url],
        title: title.trim(),
      };
    }
  } catch (error) {
    console.debug("更新頁面標題失敗", error);
  }
};

const getPageTags = async (key = pageKey) => {
  const meta = await getPageMeta();
  const tags = meta[key]?.tags;
  if (!Array.isArray(tags)) return [];
  return Array.from(new Set(tags.map((tag) => tag.trim()).filter(Boolean)));
};

const setPageTags = async (key, tags) => {
  const meta = await getPageMeta();
  meta[key] = {
    ...(meta[key] ?? {}),
    tags: Array.from(new Set(tags.map((tag) => tag.trim()).filter(Boolean))),
    updatedAt: Date.now(),
  };
  await setPageMeta(meta);
};

const saveHighlight = async (entry) => {
  if (!storage) return;
  const highlights = await getStoredHighlights(entry.url ?? pageKey);
  await setStoredHighlights([...highlights, entry], entry.url ?? pageKey);
};

const updateHighlightEntry = async (id, partial, key = pageKey) => {
  if (!storage) return false;
  const highlights = await getStoredHighlights(key);
  let updated = false;
  const next = highlights.map((item) => {
    if (item.id !== id) return item;
    updated = true;
    return { ...item, ...partial };
  });
  if (updated) {
    await setStoredHighlights(next, key);
  }
  return updated;
};

const deleteHighlightEntry = async (id, key = pageKey) => {
  if (!storage) return false;
  const highlights = await getStoredHighlights(key);
  const next = highlights.filter((item) => item.id !== id);
  if (next.length !== highlights.length) {
    await setStoredHighlights(next, key);
    return true;
  }
  return false;
};

const findHighlightEntry = async (id, key = pageKey) => {
  if (!storage) return null;
  const highlights = await getStoredHighlights(key);
  return highlights.find((item) => item.id === id) ?? null;
};

const restoreHighlights = async () => {
  if (!storage) return { total: 0, visible: 0 };
  try {
    const saved = await storage.get(pageKey);
    const allHighlights = Array.isArray(saved[pageKey]) ? saved[pageKey] : [];
    const highlights = allHighlights.filter((item) => Boolean(item?.range));
    let visibleCount = 0;

    for (const highlight of highlights) {
      const alreadyExists = document.querySelector(
        `[${HIGHLIGHT_ATTR}="${highlight.id}"]`
      );
      if (alreadyExists) {
        visibleCount += 1;
        continue;
      }

      const resolved = resolveRangeSnapshot(highlight.range);
      const range = resolved.range;
      if (!range || range.collapsed) continue;

      try {
        const highlightEl = wrapRangeWithHighlight(range, highlight.color, highlight.id);
        setHighlightMetadata(highlightEl, {
          color: highlight.color,
          note: highlight.note ?? "",
          tags: Array.isArray(highlight.tags)
            ? highlight.tags
            : parseTags(highlight.tags ?? ""),
        });
        visibleCount += 1;
        if (resolved.updated && resolved.snapshot) {
          await updateHighlightEntry(
            highlight.id,
            {
              range: resolved.snapshot,
              text: resolved.snapshot.text ?? highlight.text ?? highlight.range?.text ?? "",
            },
            pageKey
          );
        }
      } catch (error) {
        console.debug("無法還原 highlight:", highlight, error);
      }
    }

    return { total: highlights.length, visible: visibleCount };
  } catch (error) {
    console.debug("載入 highlight 失敗:", error);
    return { total: 0, visible: 0 };
  }
};

const attemptRestoreHighlights = async (attempt = 0) => {
  await ensurePageMetaTitle(pageKey, document.title);
  const { total, visible } = await restoreHighlights();
  if (!total) return;
  if (visible >= total || attempt >= HIGHLIGHT_RETRY_DELAYS.length) return;
  const nextDelay = HIGHLIGHT_RETRY_DELAYS[attempt] ?? 1200;
  window.setTimeout(() => attemptRestoreHighlights(attempt + 1), nextDelay);
};

const applyHighlight = async (color) => {
  const selection = window.getSelection();

  if (!selection || selection.rangeCount === 0) {
    throw new Error("請先選取要標註的文字。");
  }

  const range = selection.getRangeAt(0);

  if (selection.isCollapsed || range.collapsed) {
    throw new Error("選取內容為空，請選擇文字後再試。");
  }

  if (isEditableElement(range.commonAncestorContainer)) {
    throw new Error("無法在可編輯欄位內標註。");
  }

  const snapshot = serializeRange(range.cloneRange());
  const highlightId = `hk-${Date.now()}-${Math.floor(Math.random() * 100000)}`;

  const normalizedColor = toHexColor(color || DEFAULT_COLOR);
  const highlightEl = wrapRangeWithHighlight(range, normalizedColor, highlightId);
  setHighlightMetadata(highlightEl, { color: normalizedColor, note: "" });
  selection.removeAllRanges();

  await saveHighlight({
    id: highlightId,
    color: normalizedColor,
    text: snapshot.text,
    range: snapshot,
    url: pageKey,
    createdAt: Date.now(),
    note: "",
  });
  if (!colorPalette.includes(normalizedColor)) {
    await persistColorPalette([...colorPalette, normalizedColor]);
  }
  await ensurePageMetaTitle(pageKey, document.title);
  await refreshHighlightPanelIfVisible();
};

const handleSelectionIntent = () => {
  const activeEl = document.activeElement;
  const isInsideHighlightUi =
    activeEl instanceof Node &&
    ((highlightMenu && highlightMenu.contains(activeEl)) ||
      (highlightPanel && highlightPanel.contains(activeEl)));
  if (
    highlightMenu &&
    highlightMenu.style.display !== "none" &&
    isInsideHighlightUi
  ) {
    return;
  }
  if (selectionDebounceTimer) {
    window.clearTimeout(selectionDebounceTimer);
  }
  selectionDebounceTimer = window.setTimeout(() => {
    const selection = window.getSelection();
    const anchorNode = selection?.anchorNode;
    const anchorElement =
      anchorNode && anchorNode.nodeType === Node.TEXT_NODE
        ? anchorNode.parentNode
        : anchorNode;

    if (
      highlightMenu &&
      highlightMenu.style.display !== "none" &&
      anchorElement instanceof Node &&
      highlightMenu.contains(anchorElement)
    ) {
      return;
    }

    if (!selection || selection.rangeCount === 0) {
      hideFloatingButton();
      return;
    }

    const range = selection.getRangeAt(0);
    const ancestor =
      range.commonAncestorContainer instanceof HTMLElement
        ? range.commonAncestorContainer
        : range.commonAncestorContainer?.parentElement;

    if (
      selection.isCollapsed ||
      range.collapsed ||
      !ancestor ||
      isEditableElement(ancestor)
    ) {
      hideFloatingButton();
      return;
    }

    closeHighlightMenu();
    showFloatingButton(range);
  }, 60);
};

const handleHighlightClick = (event) => {
  const target = event.target;
  if (!(target instanceof Element)) return;
  const highlightEl = target.closest(`.${HIGHLIGHT_CLASS}`);
  if (!highlightEl) return;
  event.preventDefault();
  event.stopPropagation();
  hideFloatingButton();
  openHighlightMenu(highlightEl).catch((error) =>
    console.debug("開啟 highlight 面板失敗", error)
  );
};

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type === "GET_PAGE_HIGHLIGHTS") {
    (async () => {
      try {
        const highlights = await collectPageHighlights();
        const payload = {
          title: document.title,
          url: pageKey,
          pageText: getPagePlainText(),
          highlights,
        };
        sendResponse({ success: true, data: payload });
      } catch (error) {
        console.debug("取得標註資料失敗", error);
        sendResponse({ success: false, error: error?.message || "無法取得標註資料" });
      }
    })();
    return true;
  }
  if (message?.type === "SHOW_AI_NOTE") {
    const payload = message.payload || {};
    const targetKey = payload.url || pageKey;
    highlightPanelState.notesByPage = {
      ...highlightPanelState.notesByPage,
      [targetKey]: payload,
    };
    if (message.focusPanel) {
      highlightPanelState.activeKey = targetKey;
      highlightPanelState.activeTab = "ai-note";
    }
    openHighlightPanel({
      pageKey: targetKey,
      view: message.focusPanel ? "ai-note" : highlightPanelState.activeTab,
    }).then(() => {
      updateAiNoteSection(payload);
    });
    sendResponse?.({ success: true });
    return true;
  }
  if (message?.type === "OPEN_PAGE_PANEL") {
    openHighlightPanel({
      side: message.side,
      pageKey: message.pageKey,
      view: message.view,
    })
      .then(() => sendResponse({ success: true, visible: true }))
      .catch((error) =>
        sendResponse({ success: false, error: error?.message })
      );
    return true;
  }
  if (message?.type === "SET_PANEL_SIDE") {
    setHighlightPanelSide(message.side)
      .then(() => sendResponse({ success: true, side: highlightPanelState.side }))
      .catch((error) =>
        sendResponse({ success: false, error: error?.message })
      );
    return true;
  }
  if (message?.type === "TOGGLE_PAGE_PANEL") {
    toggleHighlightPanel()
      .then((visible) => sendResponse({ success: true, visible }))
      .catch((error) =>
        sendResponse({ success: false, error: error?.message })
      );
    return true;
  }
  if (message?.type === "APPLY_HIGHLIGHT") {
    const color = message.color || "#ffeb3b";
    applyHighlight(color)
      .then(() => {
        currentColor = color;
        hideFloatingButton();
        closeHighlightMenu();
        sendResponse({ success: true });
      })
      .catch((error) =>
        sendResponse({ success: false, error: error?.message })
      );
    return true;
  }
  return undefined;
});

if (document.readyState === "loading") {
  document.addEventListener(
    "DOMContentLoaded",
    () => attemptRestoreHighlights(),
    {
      once: true,
    }
  );
} else {
  attemptRestoreHighlights();
}

window.addEventListener("load", () => attemptRestoreHighlights(), {
  once: true,
});

document.addEventListener("mouseup", handleSelectionIntent);
document.addEventListener("keyup", handleSelectionIntent);
document.addEventListener("selectionchange", handleSelectionIntent);
document.addEventListener("click", handleHighlightClick, true);
document.addEventListener(
  "mousedown",
  (event) => {
    const target = event.target;
    if (floatingButton && !floatingButton.contains(target)) {
      hideFloatingButton();
    }
    if (
      highlightMenu &&
      highlightMenu.style.display !== "none" &&
      !highlightMenu.contains(target) &&
      !(activeHighlight && activeHighlight.contains(target))
    ) {
      closeHighlightMenu();
    }
  },
  true
);
window.addEventListener(
  "scroll",
  () => {
    hideFloatingButton();
    closeHighlightMenu();
  },
  true
);
window.addEventListener(
  "resize",
  () => {
    closeHighlightMenu();
    if (highlightPanelVisible) {
      renderHighlightPanel().catch((error) =>
        console.debug("重新渲染標註面板失敗", error)
      );
    }
    if (highlightPanelState.position && highlightPanel) {
      setHighlightPanelPosition(
        highlightPanelState.position.x,
        highlightPanelState.position.y,
        false
      );
    }
  },
  true
);
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeHighlightMenu();
    hideFloatingButton();
    closeHighlightPanel();
  }
});
const collectAllPageHighlights = async () => {
  if (!storage) throw new Error("無法使用儲存空間");
  const everything = await storage.get(null);
  const meta = everything[PAGE_META_KEY] || {};
  const pages = Object.entries(everything)
    .filter(([key, value]) => isValidPageKey(key) && Array.isArray(value))
    .map(([url, entries]) => ({
      url,
      title: meta[url]?.title || "",
      entries,
    }));
  return { pages, meta };
};

const handleDownloadAllHighlights = async () => {
  try {
    const { pages } = await collectAllPageHighlights();
    if (!pages.length) {
      setArchiveStatus("沒有筆記可下載", true);
      return;
    }
    const payload = {
      type: "highlight-keeper-bulk",
      version: 1,
      exportedAt: Date.now(),
      pages,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "highlight-keeper-all.json";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    setArchiveStatus("已下載全部筆記");
  } catch (error) {
    console.debug("下載全部筆記失敗", error);
    setArchiveStatus(error?.message || "下載全部筆記失敗", true);
  }
};

const mergeBulkImportPayload = async (files) => {
  if (!files?.length) return;
  try {
    setArchiveStatus("匯入中…");
    const fileTexts = await Promise.all(Array.from(files).map((file) => file.text()));
    const allEntries = fileTexts.flatMap((text) =>
      parseImportedHighlightsPayload(text)
    );
    const normalized = allEntries
      .map((entry, index) => normalizeImportedHighlightEntry(entry, index))
      .filter(Boolean);
    if (!normalized.length) {
      throw new Error("沒有可匯入的筆記");
    }
    const grouped = new Map();
    normalized.forEach((entry) => {
      if (!grouped.has(entry.url)) {
        grouped.set(entry.url, []);
      }
      grouped.get(entry.url).push(entry);
    });
    const urls = Array.from(grouped.keys());
    const existing = await storage.get(urls);
    const updates = {};
    const skipped = [];
    grouped.forEach((list, url) => {
      if (Array.isArray(existing[url]) && existing[url].length) {
        skipped.push(url);
        return;
      }
      updates[url] = list.map(({ title, ...rest }) => rest);
    });
    if (!Object.keys(updates).length) {
      throw new Error("所有頁面皆已有筆記，已全部跳過");
    }
    await storage.set(updates);
    await Promise.all(
      normalized
        .filter((entry) => entry.title)
        .map((entry) => ensurePageMetaTitle(entry.url, entry.title))
    );
    setArchiveStatus(
      `成功匯入 ${Object.keys(updates).length} 個頁面，跳過 ${skipped.length} 個頁面`
    );
    await refreshHighlightPanelData();
    await renderHighlightPanel();
    attemptRestoreHighlights();
  } catch (error) {
    console.debug("匯入全部筆記失敗", error);
    setArchiveStatus(error?.message || "匯入全部筆記失敗", true);
  }
};

const handleBulkImportChange = (event) => {
  const files = event.target.files;
  mergeBulkImportPayload(files);
};
