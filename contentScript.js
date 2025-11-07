const HIGHLIGHT_CLASS = "hk-highlight";
const HIGHLIGHT_ATTR = "data-highlight-id";

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
let highlightPanel = null;
let highlightPanelEls = null;
let highlightPanelVisible = false;
let highlightPanelPreferredSide = "right";
const PANEL_ALL_KEY = "__all__";
const highlightPanelState = {
  side: "right",
  activeKey: pageKey,
  searchTerm: "",
  activeTag: null,
  activeTab: "highlights",
  allPages: {},
  pageMeta: {},
  allTags: [],
  currentEntries: [],
  notesByPage: {},
};
let panelStatusTimer = null;
let panelPreferencesPromise;
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
      ? chrome.runtime.getURL("Icon.png")
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
  highlightPanel.classList.toggle(
    "hk-panel-left",
    highlightPanelState.side === "left"
  );
  highlightPanel.classList.toggle(
    "hk-panel-right",
    highlightPanelState.side !== "left"
  );
  if (highlightPanelEls?.sideButtons) {
    highlightPanelEls.sideButtons.forEach((btn) => {
      btn.classList.toggle(
        "is-active",
        btn.dataset.side === highlightPanelState.side
      );
    });
  }
};

const setHighlightPanelSide = async (side, persist = true) => {
  const resolved = side === "left" ? "left" : "right";
  highlightPanelState.side = resolved;
  highlightPanelPreferredSide = resolved;
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
  const activeTab =
    highlightPanelState.activeTab === "ai-note" ? "ai-note" : "highlights";
  highlightPanelState.activeTab = activeTab;
  const { tabButtons, tabPanels } = highlightPanelEls ?? {};
  const highlightBtn = tabButtons?.highlights;
  const aiBtn = tabButtons?.ai;
  const highlightPanelContent = tabPanels?.highlights;
  const aiPanelContent = tabPanels?.ai;

  if (highlightBtn) {
    const isActive = activeTab === "highlights";
    highlightBtn.classList.toggle("is-active", isActive);
    highlightBtn.setAttribute("aria-selected", isActive ? "true" : "false");
    highlightBtn.setAttribute("tabindex", isActive ? "0" : "-1");
  }

  if (aiBtn) {
    const isActive = activeTab === "ai-note";
    aiBtn.classList.toggle("is-active", isActive);
    aiBtn.setAttribute("aria-selected", isActive ? "true" : "false");
    aiBtn.setAttribute("tabindex", isActive ? "0" : "-1");
  }

  if (highlightPanelContent) {
    const showHighlights = activeTab === "highlights";
    highlightPanelContent.classList.toggle("is-active", showHighlights);
    highlightPanelContent.hidden = !showHighlights;
    highlightPanelContent.setAttribute(
      "aria-hidden",
      showHighlights ? "false" : "true"
    );
  }

  if (aiPanelContent) {
    const showAi = activeTab === "ai-note";
    aiPanelContent.classList.toggle("is-active", showAi);
    aiPanelContent.hidden = !showAi;
    aiPanelContent.setAttribute("aria-hidden", showAi ? "false" : "true");
  }
};

const loadPanelPreferences = async () => {
  if (!storage) return;
  try {
    const stored = await storage.get("hkPanelSide");
    const side =
      stored?.hkPanelSide === "left" || stored?.hkPanelSide === "right"
        ? stored.hkPanelSide
        : "right";
    highlightPanelState.side = side;
    highlightPanelPreferredSide = side;
    applyHighlightPanelSideClasses();
  } catch (error) {
    console.debug("讀取面板設定失敗", error);
  }
};

const updateHighlightPanelSelectOptions = () => {
  if (!highlightPanelEls?.pageSelect) return;
  const select = highlightPanelEls.pageSelect;
  select.innerHTML = "";

  const allOption = document.createElement("option");
  allOption.value = PANEL_ALL_KEY;
  allOption.textContent = "全部頁面";
  select.appendChild(allOption);

  const keys = Object.keys(highlightPanelState.allPages).sort((a, b) =>
    getPageDisplayName(a).localeCompare(getPageDisplayName(b))
  );

  if (!keys.includes(pageKey)) {
    keys.unshift(pageKey);
  }

  keys.forEach((key) => {
    const option = document.createElement("option");
    option.value = key;
    const metaTitle = highlightPanelState.pageMeta[key]?.title;
    const displayName = getPageDisplayName(key);
    const labelBase = metaTitle ? `${metaTitle} — ${displayName}` : displayName;
    option.textContent =
      key === pageKey ? `${labelBase}（本頁）` : labelBase;
    select.appendChild(option);
  });

  const activeValue =
    highlightPanelState.activeKey === PANEL_ALL_KEY
      ? PANEL_ALL_KEY
      : highlightPanelState.activeKey;
  select.value =
    Array.from(select.options).some((opt) => opt.value === activeValue)
      ? activeValue
      : pageKey;
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

  const activeKey = highlightPanelState.activeKey;
  const meta = highlightPanelState.pageMeta || {};
  const currentPageTags = Array.isArray(meta[activeKey]?.tags)
    ? meta[activeKey].tags
    : [];

  const isAllView = activeKey === PANEL_ALL_KEY;
  tagInput.disabled = isAllView;
  addTagBtn.disabled = isAllView;
  pageTagHint.textContent = isAllView
    ? "切換到特定頁面後可編輯標籤"
    : "輸入後按 Enter 或按新增；點擊標籤可移除";

  pageTagList.innerHTML = "";
  if (isAllView) {
    pageTagList.classList.add("is-disabled");
    tagInput.value = "";
    suggestionDropdown.classList.remove("is-visible");
    addTagBtn.onclick = null;
    tagInput.onkeydown = null;
    return;
  }
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
    if (
      highlightPanelState.activeKey !== PANEL_ALL_KEY &&
      !pages[highlightPanelState.activeKey]
    ) {
      highlightPanelState.activeKey = pageKey;
    }
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
  }
  updateHighlightPanelSelectOptions();
  updateHighlightPanelTagFilters();
};

const collectPanelEntries = () => {
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
  if (highlightPanelState.activeKey === PANEL_ALL_KEY) {
    return Object.entries(allPages).flatMap(([url, items]) =>
      items.map((entry) => ({
        ...entry,
        pageUrl: url,
        pageTags: fetchPageTags(url),
      }))
    );
  }
  const targetKey = highlightPanelState.activeKey || pageKey;
  const entries = allPages[targetKey] ?? [];
  const pageTags = fetchPageTags(targetKey);
  return entries.map((entry) => ({
    ...entry,
    pageUrl: targetKey,
    pageTags,
  }));
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

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "hk-panel-close";
  closeBtn.setAttribute("aria-label", "關閉標註面板");
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => closeHighlightPanel());

  header.appendChild(title);
  header.appendChild(closeBtn);

  const tabs = document.createElement("div");
  tabs.className = "hk-panel-tabs";
  tabs.setAttribute("role", "tablist");
  tabs.setAttribute("aria-label", "面板內容切換");

  const highlightsTabBtn = document.createElement("button");
  highlightsTabBtn.type = "button";
  highlightsTabBtn.className = "hk-panel-tab-btn";
  highlightsTabBtn.id = "hk-panel-tab-highlights";
  highlightsTabBtn.setAttribute("role", "tab");
  highlightsTabBtn.setAttribute("aria-controls", "hk-panel-tabpanel-highlights");
  highlightsTabBtn.textContent = "標註";
  highlightsTabBtn.addEventListener("click", () => {
    if (highlightPanelState.activeTab === "highlights") return;
    highlightPanelState.activeTab = "highlights";
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

  tabs.appendChild(highlightsTabBtn);
  tabs.appendChild(aiTabBtn);

  const highlightsTabPanel = document.createElement("div");
  highlightsTabPanel.className =
    "hk-panel-tab-content hk-panel-tab-content-highlights";
  highlightsTabPanel.id = "hk-panel-tabpanel-highlights";
  highlightsTabPanel.setAttribute("role", "tabpanel");
  highlightsTabPanel.setAttribute("aria-labelledby", highlightsTabBtn.id);

  const aiTabPanel = document.createElement("div");
  aiTabPanel.className = "hk-panel-tab-content hk-panel-tab-content-ai";
  aiTabPanel.id = "hk-panel-tabpanel-ai-note";
  aiTabPanel.setAttribute("role", "tabpanel");
  aiTabPanel.setAttribute("aria-labelledby", aiTabBtn.id);

  const controls = document.createElement("div");
  controls.className = "hk-panel-controls";

  const selectLabel = document.createElement("label");
  selectLabel.className = "hk-panel-select-label";
  selectLabel.textContent = "顯示標註來源";

  const pageSelect = document.createElement("select");
  pageSelect.className = "hk-panel-select";
  pageSelect.addEventListener("change", (event) => {
    const value = event.target.value;
    highlightPanelState.activeKey =
      value === PANEL_ALL_KEY ? PANEL_ALL_KEY : value;
    renderHighlightPanel();
  });
  selectLabel.appendChild(pageSelect);

  const searchWrapper = document.createElement("label");
  searchWrapper.className = "hk-panel-search";
  searchWrapper.textContent = "關鍵字搜尋";
  const searchInput = document.createElement("input");
  searchInput.type = "search";
  searchInput.className = "hk-panel-search-input";
  searchInput.placeholder = "輸入文字或標籤";
  searchInput.addEventListener("input", (event) => {
    highlightPanelState.searchTerm = event.target.value ?? "";
    renderHighlightPanel();
  });
  searchWrapper.appendChild(searchInput);

  const sideGroup = document.createElement("div");
  sideGroup.className = "hk-panel-side-group";

  const rightBtn = document.createElement("button");
  rightBtn.type = "button";
  rightBtn.dataset.side = "right";
  rightBtn.className = "hk-panel-side-btn";
  rightBtn.textContent = "右側";
  rightBtn.addEventListener("click", () => setHighlightPanelSide("right"));
  const leftBtn = document.createElement("button");
  leftBtn.type = "button";
  leftBtn.dataset.side = "left";
  leftBtn.className = "hk-panel-side-btn";
  leftBtn.textContent = "左側";
  leftBtn.addEventListener("click", () => setHighlightPanelSide("left"));

  sideGroup.appendChild(leftBtn);
  sideGroup.appendChild(rightBtn);

  controls.appendChild(selectLabel);
  controls.appendChild(searchWrapper);
  controls.appendChild(sideGroup);

  const tagsContainer = document.createElement("div");
  tagsContainer.className = "hk-panel-tags";
  controls.appendChild(tagsContainer);

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
  const importBtn = document.createElement("button");
  importBtn.type = "button";
  importBtn.className = "hk-panel-export-btn";
  importBtn.textContent = "匯入 JSON";
  exportActions.appendChild(copyBtn);
  exportActions.appendChild(downloadBtn);
  exportActions.appendChild(importBtn);
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
  importInput.addEventListener("change", handleImportFileChange);
  importBtn.addEventListener("click", () => {
    importInput.value = "";
    importInput.click();
  });
  pageTagSection.appendChild(importInput);

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

  const list = document.createElement("div");
  list.className = "hk-panel-list";

  const placeholder = document.createElement("p");
  placeholder.className = "hk-panel-placeholder";
  placeholder.textContent = "正在載入標註...";

  panel.appendChild(header);
  panel.appendChild(tabs);
  highlightsTabPanel.appendChild(controls);
  highlightsTabPanel.appendChild(pageTagSection);
  highlightsTabPanel.appendChild(list);
  highlightsTabPanel.appendChild(placeholder);
  aiTabPanel.appendChild(aiSettingsSection);
  aiTabPanel.appendChild(aiNoteSection);
  panel.appendChild(highlightsTabPanel);
  panel.appendChild(aiTabPanel);

  highlightPanel = panel;
  highlightPanelEls = {
    container: panel,
    tabs,
    list,
    placeholder,
    pageSelect,
    sideButtons: [leftBtn, rightBtn],
    searchInput,
    tagsContainer,
    pageTagSection,
    pageTagHint,
    copyBtn,
    downloadBtn,
    tagInput,
    addTagBtn,
    suggestionDropdown,
    pageTagList,
    exportStatus,
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
    tabButtons: {
      highlights: highlightsTabBtn,
      ai: aiTabBtn,
    },
    tabPanels: {
      highlights: highlightsTabPanel,
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

  applyHighlightPanelSideClasses();
  applyHighlightPanelTabState();
  updateExportButtonsState();
  applyAiSettingsToUI();
  setAiPanelStatus("");
  document.body.appendChild(panel);
  return panel;
};

const setHighlightPanelPlaceholder = (message) => {
  ensureHighlightPanel();
  if (highlightPanelEls?.placeholder) {
    highlightPanelEls.placeholder.textContent = message;
    highlightPanelEls.placeholder.style.display = "block";
  }
  if (highlightPanelEls?.list) {
    highlightPanelEls.list.innerHTML = "";
  }
};

const renderHighlightPanel = async () => {
  ensureHighlightPanel();
  if (!highlightPanelEls?.list) return;

  applyHighlightPanelTabState();
  await refreshHighlightPanelData();
  applyHighlightPanelSideClasses();
  updateHighlightPanelTagFilters();
  await renderPageTagEditor();

  const searchTerm = highlightPanelState.searchTerm.trim().toLowerCase();
  const activeTag = highlightPanelState.activeTag;

  const filtered = collectPanelEntries().filter((entry) => {
    const tags = Array.isArray(entry.pageTags) ? entry.pageTags : [];
    const matchesTag = !activeTag || tags.includes(activeTag);
    if (!matchesTag) return false;
    if (!searchTerm) return true;
    const haystacks = [entry.text, entry.note, entry.pageUrl, tags.join(" ")]
      .filter(Boolean)
      .map((value) => String(value).toLowerCase());
    return haystacks.some((text) => text.includes(searchTerm));
  });

  const entries = filtered.sort(
    (a, b) => (b.createdAt ?? 0) - (a.createdAt ?? 0)
  );

  highlightPanelState.currentEntries = entries;
  highlightPanelEls.list.innerHTML = "";
  highlightPanelEls.list.dataset.entriesCount = String(entries.length);

  if (!entries.length) {
    const noteKey =
      highlightPanelState.activeKey === PANEL_ALL_KEY
        ? pageKey
        : highlightPanelState.activeKey;
    updateAiNoteSection(highlightPanelState.notesByPage[noteKey]);
    const message =
      highlightPanelState.activeKey === PANEL_ALL_KEY
        ? "目前沒有任何標註。"
        : "該頁面尚未建立標註。";
    setHighlightPanelPlaceholder(message);
    updateExportButtonsState();
    setPanelStatus("");
    return;
  }

  if (highlightPanelEls.placeholder) {
    highlightPanelEls.placeholder.style.display = "none";
  }

  updateExportButtonsState();
  setPanelStatus("");

  entries.forEach((entry) => {
    const item = document.createElement("article");
    item.className = "hk-panel-item";
    item.setAttribute("data-highlight-id", entry.id);

    const meta = document.createElement("div");
    meta.className = "hk-panel-meta";

    const colorDot = document.createElement("span");
    colorDot.className = "hk-panel-color";
    colorDot.style.backgroundColor = entry.color ?? DEFAULT_COLOR;

    const pageLabel = document.createElement("span");
    pageLabel.className = "hk-panel-page";
    pageLabel.textContent =
      entry.pageUrl === pageKey ? "本頁" : getPageDisplayName(entry.pageUrl);

    const timestamp = document.createElement("span");
    timestamp.className = "hk-panel-time";
    timestamp.textContent = formatTimestamp(entry.createdAt);

    meta.appendChild(colorDot);
    if (entry.note) {
      const note = document.createElement("span");
      note.className = "hk-panel-note";
      note.textContent = entry.note;
      meta.appendChild(note);
    }
    meta.appendChild(pageLabel);
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

    if (entry.pageTags?.length) {
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

    highlightPanelEls.list.appendChild(item);
  });

  const noteKey =
    highlightPanelState.activeKey === PANEL_ALL_KEY
      ? pageKey
      : highlightPanelState.activeKey;
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
  const { side, pageKey: targetPage, view } = options ?? {};
  if (typeof targetPage === "string") {
    highlightPanelState.activeKey =
      targetPage === PANEL_ALL_KEY ? PANEL_ALL_KEY : targetPage;
  } else if (view === PANEL_ALL_KEY) {
    highlightPanelState.activeKey = PANEL_ALL_KEY;
  }
  if (view === "ai-note" || view === "highlights") {
    highlightPanelState.activeTab = view;
  }
  if (side) {
    await setHighlightPanelSide(side);
  } else if (highlightPanelState.side !== highlightPanelPreferredSide) {
    await setHighlightPanelSide(highlightPanelPreferredSide, false);
  } else {
    applyHighlightPanelSideClasses();
  }
  applyHighlightPanelTabState();
  highlightPanelVisible = true;
  highlightPanel.style.display = "flex";
  highlightPanel.setAttribute("aria-hidden", "false");
  await renderHighlightPanel();
  if (highlightPanelEls?.searchInput) {
    highlightPanelEls.searchInput.value = highlightPanelState.searchTerm;
  }
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
  noteField.rows = 3;
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
      const noteKey =
        highlightPanelState.activeKey === PANEL_ALL_KEY
          ? pageKey
          : highlightPanelState.activeKey;
      updateAiNoteSection(highlightPanelState.notesByPage[noteKey]);
    }
  }
  if (highlightPanelVisible) {
    const affectedKeys = Object.keys(changes);
    const shouldRefresh =
      highlightPanelState.activeKey === PANEL_ALL_KEY ||
      affectedKeys.includes(highlightPanelState.activeKey) ||
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

const serializeRange = (range) => {
  return {
    startXPath: getNodeXPath(range.startContainer),
    startOffset: range.startOffset,
    endXPath: getNodeXPath(range.endContainer),
    endOffset: range.endOffset,
    text: range.toString(),
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

      const range = deserializeRange(highlight.range);
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
      highlightPanelVisible &&
      highlightPanel &&
      !highlightPanel.contains(target)
    ) {
      closeHighlightPanel();
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
