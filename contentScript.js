const HIGHLIGHT_CLASS = "hk-highlight";
const HIGHLIGHT_ATTR = "data-highlight-id";
const MAX_SELECTOR_DEPTH = 6;
const TEXT_CONTEXT_CHARS = 60;
const TEXT_PARENT_SKIP_TAGS = new Set(["SCRIPT", "STYLE", "NOSCRIPT"]);

// 網址正規化共用自 shared.js（manifest 內已於本檔前載入）：去掉 #錨點與已知
// 追蹤參數，否則同一篇文章帶不同參數會被當成不同頁，標註就會「不見」。
const normalizePageKey = (href) =>
  window.HkUrlKey ? window.HkUrlKey.normalizePageKey(href) : href;

// 優先用頁面宣告的 canonical 連結當鍵：它是去掉所有追蹤／信件參數後最穩定的
// 文章網址（例如 Substack 的 /p/slug）。沒有 canonical 才退回剝參數的網址。
const resolvePageKey = () => {
  try {
    const link = document.querySelector('link[rel="canonical"]');
    const href = link?.href;
    if (href && /^https?:/i.test(href)) return normalizePageKey(href);
  } catch (_e) {}
  return normalizePageKey(window.location.href);
};

let pageKey = resolvePageKey();

// ── i18n（與 popup／manager 共用 i18n.js 的字典；manifest 已先注入 i18n.js） ──
const HkI18n = typeof window !== "undefined" ? window.HkI18n : null;
const t = (key, params) => (HkI18n ? HkI18n.t(key, params) : key);
// 純解析函式（manifest 內已於本檔前載入 parsers.js）
const HkParsers = window.HkParsers;
const HkHighlightDom = window.HkHighlightDom;
HkI18n?.initI18n?.();
// 語言切換時：面板是用 JS textContent 直接建立的，整個重建最單純可靠
// （面板狀態 highlightPanelState 另外保存，分頁/側邊不受影響）。
HkI18n?.onLangChange?.(() => {
  try {
    rebuildHighlightPanelForLang?.();
  } catch (_e) {}
});

const storage = chrome.storage?.local;
const DEFAULT_COLOR = "#fff5b8";
// 淺色系預設：底色夠淡，黑字疊上去仍清楚，整頁不會被螢光色蓋掉。
const DEFAULT_PALETTE = [
  "#fff5b8",
  "#ffe6cc",
  "#dff2e0",
  "#dfeefc",
  "#fce4ee",
  "#ece3fa",
];
// 舊版飽和色 → 新版淺色。升級後使用者的色票與既有標註都還存著舊色，
// 只換程式裡的預設值是看不到差別的，所以另外做搬移。
const LEGACY_COLOR_MAP = {
  "#ffeb3b": "#fff5b8",
  "#ffa726": "#ffe6cc",
  "#81c784": "#dff2e0",
  "#64b5f6": "#dfeefc",
  "#f48fb1": "#fce4ee",
  "#c792ea": "#ece3fa",
  // 1.3.0 的中間深度
  "#ffeb85": "#fff5b8",
  "#ffd9b0": "#ffe6cc",
  "#cdeccf": "#dff2e0",
  "#cfe4fa": "#dfeefc",
  "#fbd3e3": "#fce4ee",
  "#e3d7f7": "#ece3fa",
};
const PAGE_META_KEY = "__hk_page_meta__";
const PREVIOUS_DEFAULT_AI_PROMPT = `你是一位筆記整理助手，根據提供的網頁全文與標註內容，整理出淺顯易懂的筆記。優先考慮使用者標注段落，將重點控制在五百字以內，輸出內容需要像說故事一樣有脈絡的說明。`;
const PREVIOUS_CONCISE_AI_PROMPT = `整理文章主線，說明前因、結果、轉折與結論，像講清楚一段故事。只寫 3～5 句，最多 350 個中文字；保留必要細節，不重複文章內容，不補沒有根據的推論。`;
const PREVIOUS_DEFAULT_AUTO_HIGHLIGHT_PROMPT = `你是一位文章結構分析助手。請全面標註文章中各個關鍵面向的重要片段，確保覆蓋文章的完整論述架構，不要只挑少數幾句。`;
const PREVIOUS_CONCISE_AUTO_HIGHLIGHT_PROMPT = `只標出文章最值得回看的專有名詞與關鍵段落，附一行簡短解釋。少而準，不要把每個論點都標出。`;
const PREVIOUS_BRIEF_AUTO_HIGHLIGHT_PROMPT = `只挑最重要的專有名詞與關鍵段落，合計最多 5 個。每項附 20 字內的簡短說明；不重複標註，也不補原文沒有的背景。`;
const PREVIOUS_COMBINED_AUTO_HIGHLIGHT_PROMPT = `只挑最重要的專有名詞與關鍵段落，合計最多 5 個。名詞要說明「它是什麼、在本文扮演什麼角色」，不要只改寫原句；每項最多 35 字。`;
const DEFAULT_AI_PROMPT = `用 3～5 句講清楚文章主線，串起前因、結果、轉折與結論。不要逐段改寫；說明各段之間為何有關、這件事為何重要。最多 300 個中文字，只寫原文支持的內容。`;
const DEFAULT_AUTO_HIGHLIGHT_PROMPT = `依文章篇幅挑選重要段落，涵蓋主要章節、方法、證據與限制，不重複同一論點。專有名詞另外計算，不占段落名額；名詞要說明「它是什麼、在本文扮演什麼角色」，不要只改寫原句。`;
const DEFAULT_READING_GUIDE_PROMPT = `# 任務：製作閱讀導覽

你是一位會講清楚觀念的閱讀老師。請依下方文章，帶領初學者順著脈絡讀懂全文。標註要少，解釋要足以讓初學者理解。

## 回覆規則

- 這是「指令＋資料」，請直接產生結果：不要反問、不要確認任務、不要加開場白或結語。
- 整份回覆放在**單一**標記為 markdown 的程式碼區塊，區塊外不加文字，方便一鍵複製。
- 第一行、區塊標題與欄位名稱都照「回覆範本」原樣保留，不翻譯、不改寫。
- 不使用 === 分隔線，Markdown 符號前不加反斜線。
- 本文主張與背景補充必須分清楚，不要寫無關段落或泛泛而談。
- 寫給快速掃讀的人：一行只講一件事，句子要短。不要把多個意思擠在同一段，寧可多換行。

## 各區塊要求

### 脈絡
分成 3～5 個短段落，段落之間空一行，每段最多 2 句：問題從哪裡來、作者怎麼處理、證據支持什麼、限制或取捨在哪裡。自然串起前因、結果與轉折，不逐段翻述；全部約 200～400 個中文字，短文可以更短。

### 名詞解釋
只挑理解本文必要、且文章裡逐字出現的專有名詞，與重點段落分開計算。「名詞」欄照文章的寫法，不加翻譯或括號，中文意思寫進解釋。每個名詞分行寫：「解釋」一句白話說它是什麼；「例子」一句具體例子，沒有好例子就省略這行；「在本文」一句說它在這篇的作用，以及本文有沒有證實。每行不超過 40 個中文字。不可只改寫相鄰原句；一般背景知識須標明「補充」，不要捏造本文沒有的結果。

### 重點整理
挑選關鍵段落，涵蓋全文主要章節，不要只挑開頭或結論。同一個論點不重複標；方法、重要結果、數據的意義、取捨與限制都應納入考量。原文保留文章的語言，逐字引用連續文字，不翻譯、不加省略號。分行寫：「重點」一句說這段在講什麼；「為何重要」一句說它和主線的關係。每行不超過 40 個中文字，不能只是重述數字或原句。

### 閱讀地圖
用大綱畫出建議的閱讀路線，讓人由上而下就能掌握全文。主幹依原文的論述順序；子節點是父節點的展開、原因、做法或例子。重要分支展開到第 3～4 層，節點 15 字以內，內容只來自原文。

### 標籤
3～6 個代表全文主題的標籤，半形逗號分隔，不加 # 號。`;
// 分類預設名依 UI 語言產生（使用者自訂後會覆蓋這些預設）。
const getDefaultCategories = () => [
  { name: t("cat.motivation"), color: "#ece3fa" },
  { name: t("cat.method"), color: "#dfeefc" },
  { name: t("cat.pros"), color: "#dff2e0" },
  { name: t("cat.cons"), color: "#ffe6cc" },
  { name: t("cat.conclusion"), color: "#fff5b8" },
];
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
  chatgpt: [
    { value: "chatgpt-web", label: "網頁版（免 API）" },
  ],
};
let colorPalette = [...DEFAULT_PALETTE];
let currentColor = DEFAULT_COLOR;
let floatingButton = null;
const FLOATING_BUTTON_ID = "hk-floating-btn";
const FLOATING_BUTTON_MARGIN = 8;
let floatingTranslateCard = null;
let translateTargetLang = "zh-TW";
const TRANSLATE_LANGS = [
  { code: "zh-TW", label: "繁體中文" },
  { code: "zh-CN", label: "簡體中文" },
  { code: "en", label: "English" },
  { code: "ja", label: "日本語" },
  { code: "ko", label: "한국어" },
  { code: "fr", label: "Français" },
  { code: "de", label: "Deutsch" },
  { code: "es", label: "Español" },
  { code: "pt", label: "Português" },
  { code: "ru", label: "Русский" },
  { code: "ar", label: "العربية" },
];
let selectionDebounceTimer = null;
const HIGHLIGHT_MENU_ID = "hk-highlight-menu";
let highlightMenu = null;
let highlightMenuEls = null;
let activeHighlight = null;
let activeHighlightId = null;
let highlightMenuStatusTimer = null;
const HIGHLIGHT_NOTE_TOOLTIP_ID = "hk-note-tooltip";
let highlightNoteTooltip = null;
let tooltipHideTimer = null;
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
  provider: "chatgpt",
  openaiKey: "",
  openaiModel: "gpt-4o-mini",
  geminiKey: "",
  geminiModel: DEFAULT_GEMINI_MODEL,
  prompt: DEFAULT_AI_PROMPT,
  autoHighlightPrompt: DEFAULT_AUTO_HIGHLIGHT_PROMPT,
  categories: getDefaultCategories(),
  usePreview: true,
  selectionOnly: false,
  // 延伸補充：讓模型先想過一輪，再補文章「沒寫到」的背景／反面說法／延伸讀物。
  deepResearch: false,
};
let previewData = [];
const HIGHLIGHT_RETRY_DELAYS = [450, 1500, 3500];
const MINDMAP_STORAGE_KEY = "hkMindmaps";
// While a batch apply / import is running we write storage many times; this
// flag stops chrome.storage.onChanged from re-rendering the panel mid-batch.
let panelRefreshSuppressed = false;
let panelRefreshDebounceTimer = null;
let persistAISettingsTimer = null;

// ── Translation API ────────────────────────────────────
const callGoogleTranslate = async (text, targetLang) => {
  const url =
    `https://translate.googleapis.com/translate_a/single` +
    `?client=gtx&sl=auto&tl=${encodeURIComponent(targetLang)}&dt=t&dt=ld` +
    `&q=${encodeURIComponent(text)}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(t("translate.errRequest", { status: res.status }));
  const data = await res.json();
  const translated = (data[0] ?? []).map((c) => c[0] ?? "").join("");
  const detected = data[8]?.[0]?.[0] || data[2] || "auto";
  return { translated, detected };
};

const LANG_NAMES = new Intl.DisplayNames(["zh-Hant"], { type: "language" });
const langLabel = (code) => {
  try {
    return LANG_NAMES.of(code) || code;
  } catch {
    return code;
  }
};

// ── Translate card ─────────────────────────────────────
const hideTranslateCard = () => {
  if (floatingTranslateCard) floatingTranslateCard.style.display = "none";
};

const ensureTranslateCard = () => {
  if (floatingTranslateCard) return floatingTranslateCard;
  const card = document.createElement("div");
  card.className = "hk-translate-card";
  card.style.display = "none";
  card.addEventListener("mousedown", (e) => e.stopPropagation());
  document.body.appendChild(card);
  floatingTranslateCard = card;
  return card;
};

const showTranslateCard = async (text, anchorRect, highlightEl = null) => {
  const card = ensureTranslateCard();
  card.innerHTML = "";

  // header row
  const header = document.createElement("div");
  header.className = "hk-tc-header";

  const detectedSpan = document.createElement("span");
  detectedSpan.className = "hk-tc-detected";
  detectedSpan.textContent = t("translate.detecting");

  const controls = document.createElement("div");
  controls.className = "hk-tc-controls";

  const arrowSpan = document.createElement("span");
  arrowSpan.textContent = "→";
  arrowSpan.className = "hk-tc-arrow";

  const select = document.createElement("select");
  select.className = "hk-tc-select";
  TRANSLATE_LANGS.forEach(({ code, label }) => {
    const opt = document.createElement("option");
    opt.value = code;
    opt.textContent = label;
    if (code === translateTargetLang) opt.selected = true;
    select.appendChild(opt);
  });

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "hk-tc-close";
  closeBtn.setAttribute("aria-label", t("translate.closeAria"));
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", hideTranslateCard);

  controls.appendChild(arrowSpan);
  controls.appendChild(select);
  header.appendChild(detectedSpan);
  header.appendChild(controls);
  header.appendChild(closeBtn);
  card.appendChild(header);

  const resultEl = document.createElement("div");
  resultEl.className = "hk-tc-result";
  resultEl.textContent = t("translate.translating");
  card.appendChild(resultEl);

  // position (viewport coords — card uses position:fixed)
  card.style.display = "block";
  card.style.visibility = "hidden";
  const margin = 8;
  const cardW = card.offsetWidth || 320;
  const cardH = card.offsetHeight || 80;
  let top = anchorRect.bottom + margin;
  let left = anchorRect.left;
  if (left + cardW > window.innerWidth - margin) {
    left = window.innerWidth - cardW - margin;
  }
  if (anchorRect.bottom + margin + cardH > window.innerHeight) {
    top = anchorRect.top - margin - cardH;
  }
  card.style.left = `${Math.max(margin, left)}px`;
  card.style.top = `${Math.max(margin, top)}px`;
  card.style.visibility = "visible";

  const appendTranslationToNote = async (translated) => {
    if (!highlightEl) return;
    const highlightId = highlightEl.getAttribute(HIGHLIGHT_ATTR);
    if (!highlightId) return;
    const existing = highlightEl.dataset.hkNote?.trim() ?? "";
    const withoutOldTranslation = existing.replace(/\n?\n?\[翻譯\][^\n]*/g, "").trim();
    const newNote = withoutOldTranslation ? `${withoutOldTranslation}\n\n[翻譯] ${translated}` : `[翻譯] ${translated}`;
    const color = highlightEl.dataset.hkColor || highlightEl.style.backgroundColor || DEFAULT_COLOR;
    setHighlightMetadata(highlightEl, { color, note: newNote });
    if (highlightMenuEls?.noteField && activeHighlightId === highlightId) {
      highlightMenuEls.noteField.value = newNote;
    }
    try {
      await updateHighlightEntry(highlightId, { note: newNote });
      await refreshHighlightPanelIfVisible();
    } catch (err) {
      console.debug("附加翻譯到備註失敗", err);
    }
  };

  let lastTranslated = null;
  const doTranslate = async (targetLang) => {
    resultEl.textContent = t("translate.translating");
    resultEl.className = "hk-tc-result";
    try {
      const { translated, detected } = await callGoogleTranslate(text, targetLang);
      detectedSpan.textContent = langLabel(detected);
      resultEl.textContent = translated;
      if (translated !== lastTranslated) {
        lastTranslated = translated;
        appendTranslationToNote(translated);
      }
    } catch (err) {
      resultEl.textContent = err?.message || t("translate.errFallback");
      resultEl.className = "hk-tc-result hk-tc-error";
    }
  };

  select.addEventListener("change", () => {
    translateTargetLang = select.value;
    doTranslate(translateTargetLang);
  });

  doTranslate(translateTargetLang);
};

// ── Floating toolbar ───────────────────────────────────
let _floatingSelectionRange = null;

const ensureFloatingButton = () => {
  if (floatingButton) return floatingButton;

  // toolbar wrapper
  const toolbar = document.createElement("div");
  toolbar.id = FLOATING_BUTTON_ID;
  toolbar.className = "hk-floating-toolbar";
  toolbar.style.display = "none";

  // highlight button
  const hlBtn = document.createElement("button");
  hlBtn.type = "button";
  hlBtn.className = "hk-floating-btn";
  hlBtn.setAttribute("aria-label", t("toolbar.highlightAria"));
  hlBtn.title = t("toolbar.highlightTitle");
  hlBtn.textContent = "HL";

  if (chrome?.runtime?.id && typeof chrome.runtime.getURL === "function") {
    try {
      const iconSrc = chrome.runtime.getURL("Icon/32.png");
      if (iconSrc) {
        const img = new Image();
        img.decoding = "async";
        img.addEventListener("load", () => {
          hlBtn.textContent = "";
          hlBtn.style.setProperty("background-image", `url("${iconSrc}")`, "important");
          hlBtn.style.setProperty("background-repeat", "no-repeat", "important");
          hlBtn.style.setProperty("background-position", "center", "important");
          hlBtn.style.setProperty("background-size", "22px 22px", "important");
        });
        img.addEventListener("error", () => {
          hlBtn.textContent = "HL";
          hlBtn.style.removeProperty("background-image");
        });
        img.src = iconSrc;
      }
    } catch (error) {
      console.debug("無法載入浮動按鈕圖示", error);
    }
  }

  hlBtn.addEventListener("mousedown", (e) => e.preventDefault());
  hlBtn.addEventListener("click", async (e) => {
    e.preventDefault();
    try {
      await applyHighlight(currentColor);
    } catch (error) {
      console.debug("無法套用 highlight 按鈕動作", error);
    } finally {
      hideFloatingButton();
    }
  });

  // translate button
  const trBtn = document.createElement("button");
  trBtn.type = "button";
  trBtn.className = "hk-floating-tr-btn";
  trBtn.setAttribute("aria-label", t("toolbar.translateAria"));
  trBtn.title = t("toolbar.translateTitle");
  trBtn.textContent = t("toolbar.translateText");

  trBtn.addEventListener("mousedown", (e) => e.preventDefault());
  trBtn.addEventListener("click", (e) => {
    e.preventDefault();
    const sel = window.getSelection();
    const text = sel?.toString().trim();
    if (!text) return;
    const rect = _floatingSelectionRange
      ? _floatingSelectionRange.getBoundingClientRect()
      : toolbar.getBoundingClientRect();
    const anchorEl = sel?.anchorNode?.nodeType === Node.TEXT_NODE
      ? sel.anchorNode.parentElement
      : sel?.anchorNode;
    const highlightEl = anchorEl?.closest?.(`.${HIGHLIGHT_CLASS}`) ?? null;
    showTranslateCard(text, rect, highlightEl);
  });

  toolbar.appendChild(hlBtn);
  toolbar.appendChild(trBtn);
  document.body.appendChild(toolbar);
  floatingButton = toolbar;
  return toolbar;
};

const hideFloatingButton = () => {
  if (floatingButton) floatingButton.style.display = "none";
  hideTranslateCard();
};

const positionFloatingButton = (rect) => {
  const toolbar = ensureFloatingButton();
  toolbar.style.visibility = "hidden";
  toolbar.style.display = "flex";

  const { innerWidth, innerHeight } = window;
  const tbRect = toolbar.getBoundingClientRect();
  const width = tbRect.width || 80;
  const height = tbRect.height || 38;

  let top = rect.top - height - FLOATING_BUTTON_MARGIN;
  let left = rect.right - width;

  if (top < FLOATING_BUTTON_MARGIN) top = rect.bottom + FLOATING_BUTTON_MARGIN;
  if (left < FLOATING_BUTTON_MARGIN) left = rect.left;
  if (left + width > innerWidth - FLOATING_BUTTON_MARGIN) {
    left = innerWidth - width - FLOATING_BUTTON_MARGIN;
  }
  if (top + height > innerHeight - FLOATING_BUTTON_MARGIN) {
    top = innerHeight - height - FLOATING_BUTTON_MARGIN;
  }

  toolbar.style.top = `${Math.max(FLOATING_BUTTON_MARGIN, top)}px`;
  toolbar.style.left = `${Math.max(FLOATING_BUTTON_MARGIN, left)}px`;
  toolbar.style.visibility = "visible";
};

const showFloatingButton = (range) => {
  const rect = range.getBoundingClientRect();
  if (!rect || (rect.width === 0 && rect.height === 0)) {
    hideFloatingButton();
    return;
  }
  _floatingSelectionRange = range;
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

// 色票搬移：只有在「整組色票都還是舊預設」時才自動換掉，
// 使用者自己調過顏色就不動他的設定。
const migrateLegacyPalette = (palette) => {
  if (!Array.isArray(palette) || !palette.length) return palette;
  const allLegacy = palette.every((color) => LEGACY_COLOR_MAP[color]);
  if (!allLegacy) return palette;
  return palette.map((color) => LEGACY_COLOR_MAP[color] || color);
};

const loadPalette = async () => {
  if (!storage) return colorPalette;
  try {
    const stored = await storage.get("hkColorPalette");
    const raw = sanitizePalette(stored?.hkColorPalette);
    const palette = migrateLegacyPalette(raw);
    setColorPaletteState(palette);
    if (palette.join() !== raw.join()) {
      storage
        .set({ hkColorPalette: palette })
        .catch((e) => console.debug("搬移舊色票失敗", e));
    }
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

// 樣板雜訊：導覽、頁首頁尾、側欄、隱藏元素，以及本擴充自己注入的 UI。
// 這些混進「原文內容」會浪費 token，也會誘使 AI 把選單文字當成重點。
const BOILERPLATE_SELECTOR = [
  "nav",
  "header",
  "footer",
  "aside",
  "script",
  "style",
  "noscript",
  '[role="navigation"]',
  '[role="banner"]',
  '[role="contentinfo"]',
  '[aria-hidden="true"]',
  ".hk-page-panel",
  ".hk-mindmap-overlay",
  "#hk-guided-bar",
].join(", ");

const getPagePlainText = () => {
  let sourceText = "";
  // 1) 優先抓明確的主內容容器（一頁通常只有一個 <main>），最能去掉導覽雜訊。
  const main =
    document.querySelector("main") ||
    document.querySelector('[role="main"]') ||
    document.querySelector("article");
  if (main && (main.innerText || "").trim().length >= 200) {
    sourceText = main.innerText;
  } else {
    // 2) 後備：複製 body，移除導覽/頁首/頁尾等樣板後再取文字（不動到實際頁面）。
    const clone = document.body?.cloneNode(true);
    if (clone) {
      clone
        .querySelectorAll(BOILERPLATE_SELECTOR)
        .forEach((el) => el.remove());
      sourceText = clone.innerText || clone.textContent || "";
    } else {
      sourceText = document.body?.innerText || "";
    }
  }
  const normalized = sourceText.replace(/\n{3,}/g, "\n\n").trim();
  return normalized.slice(0, 60000);
};

const collectPageHighlights = async () => {
  const highlights = await getStoredHighlights();
  return highlights.map((item) => ({
    id: item.id,
    text: item.text ?? "",
    note: item.note ?? "",
    color: item.color,
    ...(item.kind ? { kind: item.kind } : {}),
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

// For high-frequency inputs (prompt textarea, API key) — one write after the
// user stops typing instead of one write per keystroke.
const schedulePersistAISettings = () => {
  if (persistAISettingsTimer) window.clearTimeout(persistAISettingsTimer);
  persistAISettingsTimer = window.setTimeout(() => {
    persistAISettingsTimer = null;
    persistAISettings();
  }, 400);
};

const isLikelyProviderApiKey = (provider, value) => {
  const key = typeof value === "string" ? value.trim() : "";
  if (provider === "openai") return /^sk-/i.test(key);
  if (provider === "gemini") return /^AIza/i.test(key);
  return false;
};

const updateAiKeyVisibility = () => {
  const groups = highlightPanelEls?.aiKeyGroups ?? [];
  const isWebProvider = ["chatgpt", "gemini-web"].includes(aiSettings.provider);
  groups.forEach((group) => {
    const provider = group.dataset.provider;
    const isVisible = provider === "web" ? isWebProvider : provider === aiSettings.provider;
    group.classList.toggle("is-visible", isVisible);
    group.hidden = !isVisible;
    if (!isVisible) {
      const field = group.querySelector("input");
      if (field) {
        field.value = aiSettings[provider === "openai" ? "openaiKey" : "geminiKey"] ?? "";
      }
    }
  });
  const modelField = highlightPanelEls?.aiModelSelect?.closest(".hk-panel-ai-field");
  if (modelField) {
    modelField.hidden = ["chatgpt", "gemini-web"].includes(aiSettings.provider);
  }
  const webInfo = highlightPanelEls?.aiWebInfo;
  if (webInfo) webInfo.textContent = t(aiSettings.provider === "gemini-web" ? "panel.geminiWebInfo" : "panel.chatgptInfo");
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
    opt.textContent =
      option.value === "chatgpt-web" ? t("provider.chatgptWeb") : option.label;
    select.appendChild(opt);
  });
  const currentValue = provider === "openai" ? aiSettings.openaiModel : aiSettings.geminiModel;
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

const applyAiSettingsToUI = () => {
  const providerSelect = highlightPanelEls?.aiProviderSelect;
  const modelSelect = highlightPanelEls?.aiModelSelect;
  const promptField = highlightPanelEls?.aiPromptField;
  const autoHighlightPromptField = highlightPanelEls?.aiAutoHighlightPromptField;
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
  if (autoHighlightPromptField) {
    autoHighlightPromptField.value =
      aiSettings.autoHighlightPrompt ?? DEFAULT_AUTO_HIGHLIGHT_PROMPT;
  }
  if (openaiInput) {
    openaiInput.value = aiSettings.openaiKey ?? "";
  }
  if (geminiInput) {
    geminiInput.value = aiSettings.geminiKey ?? "";
  }
  const previewCb = highlightPanelEls?.aiPreviewCheckbox;
  if (previewCb) previewCb.checked = aiSettings.usePreview ?? true;
  const selOnlyCb = highlightPanelEls?.aiSelOnlyCheckbox;
  if (selOnlyCb) selOnlyCb.checked = aiSettings.selectionOnly ?? false;
  const deepCb = highlightPanelEls?.aiDeepResearchCheckbox;
  if (deepCb) deepCb.checked = aiSettings.deepResearch !== false;
  updateAiKeyVisibility();
  renderCategoryList();
};

const renderCategoryList = () => {
  const catList = highlightPanelEls?.aiCatList;
  if (!catList) return;
  catList.innerHTML = "";
  const cats = Array.isArray(aiSettings.categories) ? aiSettings.categories : [];
  cats.forEach((cat, idx) => {
    const row = document.createElement("div");
    row.className = "hk-panel-ai-cat-row";
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "hk-panel-ai-cat-name";
    nameInput.value = cat.name;
    nameInput.placeholder = t("panel.categoryPlaceholder");
    nameInput.addEventListener("input", (e) => {
      aiSettings.categories[idx].name = e.target.value;
      schedulePersistAISettings();
    });
    const colorInput = document.createElement("input");
    colorInput.type = "color";
    colorInput.className = "hk-panel-ai-cat-color";
    colorInput.value = cat.color;
    colorInput.addEventListener("input", (e) => {
      aiSettings.categories[idx].color = e.target.value;
      schedulePersistAISettings();
    });
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "hk-panel-ai-cat-del";
    delBtn.textContent = "✕";
    delBtn.addEventListener("click", () => {
      aiSettings.categories.splice(idx, 1);
      persistAISettings();
      renderCategoryList();
    });
    row.appendChild(colorInput);
    row.appendChild(nameInput);
    row.appendChild(delBtn);
    catList.appendChild(row);
  });
};

// 「自動畫重點」用的是分類色，跟色票一樣會留在 storage 裡。
// 同樣採保守規則：整組都還是舊預設色才自動換淺，使用者調過就不動。
const migrateLegacyCategories = (categories) => {
  if (!Array.isArray(categories) || !categories.length) return categories;
  const allLegacy = categories.every((c) => LEGACY_COLOR_MAP[toHexColor(c?.color || "")]);
  if (!allLegacy) return categories;
  return categories.map((c) => ({
    ...c,
    color: LEGACY_COLOR_MAP[toHexColor(c.color)] || c.color,
  }));
};

const loadAISettings = async () => {
  try {
    const stored = await chrome.storage?.local.get("hkAISettings");
    const savedSettings = stored?.hkAISettings;
    if (stored?.hkAISettings) {
      aiSettings = {
        ...aiSettings,
        ...stored.hkAISettings,
      };
    }
    let promptMigrated = false;
    if (
      (aiSettings.provider === "openai" && !isLikelyProviderApiKey("openai", aiSettings.openaiKey)) ||
      (aiSettings.provider === "gemini" && !isLikelyProviderApiKey("gemini", aiSettings.geminiKey))
    ) {
      aiSettings.provider = "chatgpt";
      promptMigrated = true;
    }
    if ([PREVIOUS_DEFAULT_AI_PROMPT, PREVIOUS_CONCISE_AI_PROMPT].includes(savedSettings?.prompt)) {
      aiSettings.prompt = DEFAULT_AI_PROMPT;
      promptMigrated = true;
    }
    if (
      [PREVIOUS_DEFAULT_AUTO_HIGHLIGHT_PROMPT, PREVIOUS_CONCISE_AUTO_HIGHLIGHT_PROMPT, PREVIOUS_BRIEF_AUTO_HIGHLIGHT_PROMPT, PREVIOUS_COMBINED_AUTO_HIGHLIGHT_PROMPT]
        .includes(savedSettings?.autoHighlightPrompt)
    ) {
      aiSettings.autoHighlightPrompt = DEFAULT_AUTO_HIGHLIGHT_PROMPT;
      promptMigrated = true;
    }
    if (
      [PREVIOUS_DEFAULT_AI_PROMPT, PREVIOUS_CONCISE_AI_PROMPT].includes(savedSettings?.prompt) &&
      [PREVIOUS_DEFAULT_AUTO_HIGHLIGHT_PROMPT, PREVIOUS_CONCISE_AUTO_HIGHLIGHT_PROMPT, PREVIOUS_BRIEF_AUTO_HIGHLIGHT_PROMPT, PREVIOUS_COMBINED_AUTO_HIGHLIGHT_PROMPT]
        .includes(savedSettings?.autoHighlightPrompt)
    ) {
      aiSettings.deepResearch = false;
      promptMigrated = true;
    }
    const migrated = migrateLegacyCategories(aiSettings.categories);
    if (migrated !== aiSettings.categories || promptMigrated) {
      aiSettings.categories = migrated;
      persistAISettings();
    }
  } catch (error) {
    console.debug("讀取 AI 設定失敗", error);
  } finally {
    applyAiSettingsToUI();
  }
};

const normalizeWhitespace = (input) =>
  typeof input === "string" ? input.replace(/\s+/g, " ").trim() : "";

// 色彩使用次數只用於排序色票提示，變動不頻繁；快取避免每次複製 Prompt 都全量讀。
// 任何標註寫入（setStoredHighlights）會把它設為 null 以失效。
let colorUsageCache = null;
const collectColorUsageCounts = async () => {
  if (!storage) return {};
  if (colorUsageCache) return colorUsageCache;
  try {
    const all = await storage.get(null);
    const counts = {};
    Object.entries(all).forEach(([key, entries]) => {
      if (!isValidPageKey(key) || !Array.isArray(entries)) return;
      entries.forEach((entry) => {
        if (!entry || typeof entry !== "object") return;
        const color = toHexColor(entry.color || DEFAULT_COLOR);
        counts[color] = (counts[color] || 0) + 1;
      });
    });
    colorUsageCache = counts;
    return counts;
  } catch (error) {
    console.debug("讀取顏色使用次數失敗", error);
    return {};
  }
};

const sortPaletteByUsage = (palette, counts = {}) => {
  const order = new Map();
  palette.forEach((color, index) => {
    order.set(color, index);
  });
  return [...palette].sort((a, b) => {
    const diff = (counts[b] || 0) - (counts[a] || 0);
    if (diff !== 0) return diff;
    return (order.get(a) ?? 0) - (order.get(b) ?? 0);
  });
};

// 所有入口共用：閱讀導覽要讀的頁面資料（設定「只標選取的文字」時改用選取範圍）。
const collectGuidePageData = async () => {
  const selection = aiSettings.selectionOnly ? window.getSelection()?.toString()?.trim() : "";
  return {
    title: document.title,
    url: pageKey,
    pageText: selection || getPagePlainText(),
    highlights: await collectPageHighlights(),
  };
};

// 整個擴充功能唯一的 Prompt：一次取得脈絡、名詞解釋、重點整理、閱讀地圖與標籤。
const buildReadingGuidePrompt = (pageData, palette, existingTags = []) => {
  const limits = HkParsers.getReadingGuideLimits(pageData.pageText);
  const customInstructions = [
    aiSettings.prompt?.trim() && aiSettings.prompt.trim() !== DEFAULT_AI_PROMPT
      ? `- 脈絡補充要求：${aiSettings.prompt.trim()}`
      : "",
    aiSettings.autoHighlightPrompt?.trim() && aiSettings.autoHighlightPrompt.trim() !== DEFAULT_AUTO_HIGHLIGHT_PROMPT
      ? `- 標註補充要求：${aiSettings.autoHighlightPrompt.trim()}`
      : "",
  ].filter(Boolean).join("\n");
  const categoryLines = Array.isArray(aiSettings.categories) && aiSettings.categories.length
    ? aiSettings.categories.map((category) => `- #${category.name} (${category.color})`).join("\n")
    : sanitizePalette(palette).map((color) => `- ${color}`).join("\n");
  const existing = (pageData.highlights || [])
    .map((item) => `- ${normalizeWhitespace(item.text)}`)
    .filter((line) => line.length > 2)
    .join("\n");
  // 延伸掛在各重點底下，讀的時候補充就在那句話旁邊。
  const deepFields = aiSettings.deepResearch === true
    ? "\n  - **延伸：** 原文沒說、但值得知道的一句話（沒有就省略）\n  - **來源：** 可查證的名字或搜尋關鍵詞（不確定就省略，不要寫網址）"
    : "";
  return `${DEFAULT_READING_GUIDE_PROMPT}

## 回覆範本

~~~markdown
來源文章：${pageData.url}

## 脈絡
第一段：問題從哪裡來。

第二段：作者怎麼處理。

第三段：證據與限制。

## 名詞解釋
- **名詞：** A/B test
  - **解釋：** 把使用者分到不同版本，比較哪個表現好。
  - **例子：** 部分請求用新模型，其餘維持舊版。
  - **在本文：** 被列為流量變化的可能原因，但沒有證據證實。

## 重點整理
- **原文：** 文章中的連續原句
  - **分類：** #分類或色碼
  - **重點：** 這段在講什麼
  - **為何重要：** 它和主線的關係${deepFields}

## 閱讀地圖
# 文章主題
- 背景與問題
  - 面向
    - 具體例子
- 核心做法
- 結論與限制

## 標籤
標籤一, 標籤二, 標籤三
~~~

## 這篇文章的設定

- 重點整理：${limits.highlightsMin}～${limits.highlightsMax} 段；名詞解釋：最多 ${limits.termsMax} 個。短文沒有足夠的獨立重點可少選，不要湊數；長文需覆蓋分散在全文的重要章節。
- ${t("prompt.outputLang")}
${customInstructions ? `${customInstructions}\n` : ""}
### 可用分類與顏色
${categoryLines || "- #fff5b8"}
${existingTags.length ? `\n### 既有標籤（語意相符時優先沿用）\n${existingTags.join("、")}\n` : ""}
### 已有標註（請避免重複）
${existing || "（無）"}

## 文章資料

- 標題：${pageData.title}
- 網址：${pageData.url}

### 文章全文
${pageData.pageText}`;
};

const saveReadingGuideContext = async (context, provider, model, prompt, targetPageKey, unmatchedTerms = [], unmatchedHighlights = [], guideItems = []) => {
  if (!targetPageKey) return;
  const stored = await chrome.storage.local.get("hkGeneratedNotes");
  const previous = stored?.hkGeneratedNotes?.[targetPageKey] || {};
  const payload = {
    ...previous,
    ...(context?.trim() ? { context: context.trim() } : {}),
    unlocatedTerms: unmatchedTerms.map((item) => ({ text: item.text, explanation: item.reason, unavailable: true })),
    unlocatedHighlights: unmatchedHighlights.map((item) => ({ text: item.text, reason: item.reason, unavailable: true })),
    guideTerms: guideItems.filter((item) => item.kind === "term").map((item) => ({ ...item, explanation: item.reason })),
    guideHighlights: guideItems.filter((item) => item.kind !== "term"),
    formatVersion: 2,
    provider,
    model,
    prompt,
    generatedAt: Date.now(),
    url: targetPageKey,
  };
  await saveGeneratedNote(targetPageKey, payload);
  highlightPanelState.notesByPage = {
    ...highlightPanelState.notesByPage,
    [targetPageKey]: payload,
  };
};

// 預覽中按取消要回到套用前的樣子，不只撤掉標記。
let previewUndo = null;
const capturePreviewUndo = async (targetPageKey) => {
  const stored = await chrome.storage.local.get(["hkGeneratedNotes", MINDMAP_STORAGE_KEY, PAGE_META_KEY]);
  return {
    pageKey: targetPageKey,
    note: stored.hkGeneratedNotes?.[targetPageKey],
    mindmap: stored[MINDMAP_STORAGE_KEY]?.[targetPageKey],
    meta: stored[PAGE_META_KEY]?.[targetPageKey],
    entries: await getStoredHighlights(targetPageKey),
  };
};

const applyReadingGuideResponse = async (rawText, prompt = "", provider = "chatgpt", model = "chatgpt-web", targetPageKey = pageKey) => {
  if (targetPageKey !== pageKey) throw new Error(t("sidebar.parseError"));
  const parsed = HkParsers.parseReadingGuideResponse(rawText, {
    categories: aiSettings.categories,
    toHexColor,
    defaultColor: DEFAULT_COLOR,
  });
  if (!parsed.context && !parsed.terms.length && !parsed.highlights.length && !parsed.mindmap) {
    throw new Error(t("sidebar.parseError"));
  }
  if (!HkParsers.matchesReadingGuideSource(parsed, targetPageKey, getPagePlainText())) {
    throw new Error(t("sidebar.sourceMismatch"));
  }
  const undoSnapshot = previewUndo || await capturePreviewUndo(targetPageKey);
  const latestPalette = await refreshPaletteFromStorage();
  if (targetPageKey !== pageKey) throw new Error(t("sidebar.parseError"));
  const usageCounts = await collectColorUsageCounts();
  const palette = sortPaletteByUsage(latestPalette, usageCounts);
  const existing = [
    ...await collectPageHighlights(),
    ...previewData.map((item) => item.data).filter((item) => item.url === targetPageKey),
  ];
  if (targetPageKey !== pageKey) throw new Error(t("sidebar.parseError"));
  const seen = new Set(existing.map((item) => HkParsers.looseKey(item.text)));
  const pageText = getPagePlainText();
  const limits = HkParsers.getReadingGuideLimits(pageText);
  const guideItems = [
    ...parsed.terms.slice(0, limits.termsMax).map((term) => {
      const text = normalizeWhitespace(HkParsers.pickTermText(term.text, pageText));
      const label = normalizeWhitespace(term.text);
      return { text, ...(label !== text ? { label } : {}), kind: "term", reason: term.explanation, color: palette[1] || palette[0] || DEFAULT_COLOR };
    }),
    ...parsed.highlights.slice(0, limits.highlightsMax).map((item) => ({ ...item, text: normalizeWhitespace(item.text), kind: "passage" })),
  ].filter((item) => item.text);
  const tasks = guideItems.filter((item) => {
    const key = HkParsers.looseKey(item.text);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  // Reapplying a reply should refresh explanations on existing marks too.
  const entries = await getStoredHighlights(targetPageKey);
  if (targetPageKey !== pageKey) throw new Error(t("sidebar.sourceMismatch"));
  let updatedExisting = false;
  const revisedEntries = entries.map((entry) => {
    const item = guideItems.find((candidate) => HkParsers.looseKey(candidate.text) === HkParsers.looseKey(entry.text));
    if (!item) return entry;
    const note = formatAiNote(item.reason, item.extra, item.source);
    if (entry.note === note && entry.kind === item.kind) return entry;
    updatedExisting = true;
    setAllMarksMetadata(entry.id, { note, kind: item.kind });
    return { ...entry, note, kind: item.kind };
  });
  if (updatedExisting) await setStoredHighlights(revisedEntries, targetPageKey);
  previewData.forEach((preview) => {
    if (preview.data.url !== targetPageKey) return;
    const item = guideItems.find((candidate) => HkParsers.looseKey(candidate.text) === HkParsers.looseKey(preview.data.text));
    if (!item) return;
    preview.data.note = formatAiNote(item.reason, item.extra, item.source);
    preview.data.kind = item.kind;
    setAllMarksMetadata(preview.data.id, { note: preview.data.note, kind: item.kind });
  });

  let appliedCount = 0;
  let previewing = previewData.some((item) => item.data.url === targetPageKey);
  let unmatchedTasks = [];
  if (tasks.length) {
    const result = await applyAutoHighlightTasksBatch(tasks, aiSettings.usePreview);
    if (targetPageKey !== pageKey) throw new Error(t("sidebar.parseError"));
    unmatchedTasks = result.unmatched;
    appliedCount = result.appliedCount;
    previewing = aiSettings.usePreview && previewData.length > 0;
    if (previewing) showPreviewConfirmBar(previewData.length);
    else await ensurePageMetaTitle(pageKey, document.title);
  }
  const unmatchedByKey = new Map(unmatchedTasks.map((item) => [HkParsers.looseKey(item.text), item]));
  const unmatchedKeys = new Set(unmatchedByKey.keys());
  const isCovered = (item) => Boolean(unmatchedByKey.get(HkParsers.looseKey(item.text))?.covered);
  const savedGuideItems = guideItems.map((item) => ({
    ...item,
    unavailable: unmatchedKeys.has(HkParsers.looseKey(item.text)),
    ...(isCovered(item) ? { covered: true } : {}),
  }));
  if (previewData.some((item) => item.data.url === targetPageKey)) previewUndo = undoSnapshot;
  await saveReadingGuideContext(parsed.context, provider, model, prompt, targetPageKey,
    unmatchedTasks.filter((item) => item.kind === "term"),
    unmatchedTasks.filter((item) => item.kind !== "term"), savedGuideItems);
  if (parsed.mindmap && parseMindmapOutline(parsed.mindmap)) {
    await saveMindmap(parsed.mindmap, parseMindmapOutline(parsed.mindmap).title);
  }
  if (parsed.tags) await applyAiPageTags(parsed.tags);
  return {
    success: true,
    context: parsed.context,
    terms: savedGuideItems.filter((item) => item.kind === "term").map((item) => ({
      text: item.text,
      explanation: item.reason,
      ...(unmatchedKeys.has(HkParsers.looseKey(item.text)) ? { unavailable: true } : {}),
      ...(item.covered ? { covered: true } : {}),
      ...(item.label ? { label: item.label } : {}),
    })),
    highlights: savedGuideItems.filter((item) => item.kind !== "term").map((item) => ({
      ...item,
      ...(unmatchedKeys.has(HkParsers.looseKey(item.text)) ? { unavailable: true } : {}),
    })),
    appliedCount,
    unmatchedCount: unmatchedTasks.filter((item) => !item.covered && item.kind !== "term").length,
    previewCount: previewData.filter((item) => item.data.url === targetPageKey).length,
    previewing,
  };
};

// 統計所有頁面的既有標籤（依使用頁數排序），給 AI 標籤任務優先重用，
// 避免每次都發明只用一次的新標籤、標籤越長越多。
const collectExistingTagsByUsage = async (limit = 30) => {
  try {
    const stored = await chrome.storage?.local.get(PAGE_META_KEY);
    const meta = stored?.[PAGE_META_KEY] || {};
    const counts = new Map();
    Object.values(meta).forEach((entry) => {
      (Array.isArray(entry?.tags) ? entry.tags : []).forEach((tag) => {
        const name = String(tag).trim();
        if (name) counts.set(name, (counts.get(name) || 0) + 1);
      });
    });
    return [...counts.entries()]
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .slice(0, limit)
      .map(([name]) => name);
  } catch (_e) {
    return [];
  }
};

// Parse a `# root` + indented `- ` outline into a tree.
const parseMindmapOutline = (rawText) =>
  HkParsers.parseMindmapOutline(rawText, document.title || t("mindmap.title"));

const getStoredMindmaps = async () => {
  if (!storage) return {};
  try {
    const stored = await storage.get(MINDMAP_STORAGE_KEY);
    return stored?.[MINDMAP_STORAGE_KEY] ?? {};
  } catch (_e) {
    return {};
  }
};

const saveMindmap = async (outlineText, title) => {
  if (!storage) return;
  const maps = await getStoredMindmaps();
  maps[pageKey] = {
    outline: outlineText,
    title: title || document.title || "",
    generatedAt: Date.now(),
  };
  await storage.set({ [MINDMAP_STORAGE_KEY]: maps });
};

const normalizeAutoHighlightItems = (payload, palette) => {
  const source = Array.isArray(payload)
    ? payload
    : Array.isArray(payload?.highlights)
    ? payload.highlights
    : Array.isArray(payload?.items)
    ? payload.items
    : [];
  const normalizedPalette = sanitizePalette(palette);
  const allowedColors = new Set(normalizedPalette);
  // also allow any category colors defined in current aiSettings
  if (Array.isArray(aiSettings.categories)) {
    aiSettings.categories.forEach((c) => {
      if (c?.color) allowedColors.add(toHexColor(c.color));
    });
  }
  const fallbackColor = normalizedPalette[0] || DEFAULT_COLOR;
  const seenText = new Set();

  return source
    .map((item) => {
      if (!item || typeof item !== "object") return null;
      const text = normalizeWhitespace(item.text);
      if (!text) return null;
      const dedupeKey = text.toLowerCase();
      if (seenText.has(dedupeKey)) return null;
      seenText.add(dedupeKey);

      const candidateColor =
        typeof item.color === "string" ? toHexColor(item.color.trim()) : fallbackColor;
      const color = allowedColors.has(candidateColor) ? candidateColor : fallbackColor;
      const reason =
        typeof item.reason === "string"
          ? normalizeWhitespace(item.reason).slice(0, 180)
          : "";
      // 延伸寫在同一個重點區塊裡，跟著重點一起變成這則標註的註解。
      const extra =
        typeof item.extra === "string"
          ? normalizeWhitespace(item.extra).slice(0, 240)
          : "";
      // 來源刻意不收網址：模型編造連結的機率遠高於編造書名／論文名，
      // 我們只存「查得到的名字」，要驗證時再由前端組搜尋連結。
      const source =
        typeof item.source === "string"
          ? normalizeWhitespace(item.source).replace(/https?:\/\/\S+/g, "").slice(0, 120).trim()
          : "";
      const kind = item.kind === "term" ? "term" : item.kind === "passage" ? "passage" : undefined;
      return { text, color, reason, extra, source, ...(kind ? { kind } : {}) };
    })
    .filter(Boolean);
};

// 標註註解最多三行：「AI：重點」／「延伸：…」／「來源：…」。
const formatAiNote = (reason, extra, source) => {
  const lines = [];
  if (reason) lines.push(`AI：${reason}`);
  if (extra) lines.push(`${t("ai.extraNotePrefix")}${extra}`);
  if (extra && source) lines.push(`${t("ai.sourceNotePrefix")}${source}`);
  return lines.join("\n");
};

const rangeTouchesExistingHighlight = (range) => {
  const boundaries = [range.startContainer, range.endContainer];
  for (const boundary of boundaries) {
    const element =
      boundary?.nodeType === Node.ELEMENT_NODE
        ? boundary
        : boundary?.parentElement || boundary?.parentNode;
    if (element?.closest?.(`.${HIGHLIGHT_CLASS}`)) {
      return true;
    }
  }
  try {
    const fragment = range.cloneContents();
    return Boolean(fragment.querySelector?.(`.${HIGHLIGHT_CLASS}`));
  } catch (_error) {
    return true;
  }
};

const buildNormalizedSearchIndex = () => {
  const index = buildDocumentTextIndex();
  if (!index || !Array.isArray(index.nodes) || !index.nodes.length) return null;
  const rawText = index.text || "";
  const normalizedToRaw = [];
  let normalizedText = "";
  let lastWasWhitespace = false;
  for (let i = 0; i < rawText.length; i += 1) {
    const char = rawText[i];
    const isWhitespace = /\s/.test(char);
    if (isWhitespace) {
      if (!lastWasWhitespace) {
        normalizedText += " ";
        normalizedToRaw.push(i);
      }
      lastWasWhitespace = true;
      continue;
    }
    normalizedText += char;
    normalizedToRaw.push(i);
    lastWasWhitespace = false;
  }
  return {
    ...index,
    rawTextLength: rawText.length,
    normalizedText,
    normalizedToRaw,
  };
};

const findIndexesInText = (haystack, needle) => {
  const indexes = [];
  let fromIndex = 0;
  while (fromIndex <= haystack.length) {
    const index = haystack.indexOf(needle, fromIndex);
    if (index === -1) break;
    indexes.push(index);
    fromIndex = index + 1;
  }
  return indexes;
};

const normalizedSpanToRawSpan = (mapping, start, length, rawTextLength) => {
  if (!Array.isArray(mapping) || !length) return null;
  const startRaw = mapping[start];
  const endRawChar = mapping[start + length - 1];
  if (!Number.isFinite(startRaw) || !Number.isFinite(endRawChar)) return null;
  const endRaw = Math.min(rawTextLength, endRawChar + 1);
  if (endRaw <= startRaw) return null;
  return { startRaw, endRaw };
};

const createRangeFromRawOffsets = (nodes, startRaw, endRaw) => {
  const startPosition = locateTextPosition(nodes, startRaw);
  const endPosition = locateTextPosition(nodes, endRaw);
  if (!startPosition || !endPosition) return null;
  const range = document.createRange();
  range.setStart(startPosition.node, startPosition.offset);
  range.setEnd(endPosition.node, endPosition.offset);
  if (range.collapsed) return null;
  return range;
};

const findRangeForAutoHighlightText = (targetText, consumedSpans) => {
  const normalizedNeedle = normalizeWhitespace(targetText);
  if (!normalizedNeedle) return null;
  const searchIndex = buildNormalizedSearchIndex();
  if (!searchIndex?.normalizedText) return null;

  const attempts = [
    {
      haystack: searchIndex.normalizedText,
      needle: normalizedNeedle,
    },
    {
      haystack: searchIndex.normalizedText.toLowerCase(),
      needle: normalizedNeedle.toLowerCase(),
    },
  ];

  for (const attempt of attempts) {
    const indexes = findIndexesInText(attempt.haystack, attempt.needle);
    for (const start of indexes) {
      const rawSpan = normalizedSpanToRawSpan(
        searchIndex.normalizedToRaw,
        start,
        attempt.needle.length,
        searchIndex.rawTextLength
      );
      if (!rawSpan) continue;
      const spanKey = `${rawSpan.startRaw}:${rawSpan.endRaw}`;
      if (consumedSpans?.has(spanKey)) continue;
      const range = createRangeFromRawOffsets(
        searchIndex.nodes,
        rawSpan.startRaw,
        rawSpan.endRaw
      );
      if (!range) continue;
      const ancestor =
        range.commonAncestorContainer instanceof HTMLElement
          ? range.commonAncestorContainer
          : range.commonAncestorContainer?.parentElement;
      if (ancestor && isEditableElement(ancestor)) continue;
      if (rangeTouchesExistingHighlight(range)) continue;
      consumedSpans?.add(spanKey);
      return range;
    }
  }
  return null;
};

const applyAutoHighlightRange = async (range, color, reason, extra = "", source = "") => {
  const snapshot = serializeRange(range.cloneRange());
  const text = normalizeWhitespace(snapshot.text);
  if (!text) return null;
  const highlightId = `hk-ai-${Date.now()}-${Math.floor(Math.random() * 100000)}`;
  const normalizedColor = toHexColor(color || DEFAULT_COLOR);
  const note = formatAiNote(reason, extra, source);
  const highlightEl = wrapRangeWithHighlight(range, normalizedColor, highlightId);
  setAllMarksMetadata(highlightId, { color: normalizedColor, note });
  await saveHighlight({
    id: highlightId,
    color: normalizedColor,
    text: snapshot.text,
    range: snapshot,
    url: pageKey,
    createdAt: Date.now(),
    note,
  });
  return {
    id: highlightId,
    color: normalizedColor,
    text,
  };
};

// ── Batch auto-highlight ─────────────────────────────────
// The old path rebuilt the full-document text index for EVERY item and wrote
// storage once per item (each write re-rendering the panel). Here we build the
// index once, locate every span, apply from the end of the document backwards
// (so earlier raw offsets stay valid while the DOM mutates), then persist all
// entries with a single storage write.
// AI 引用時常改動空白與標點（直引號／彎引號、破折號、中英文之間的空格、跨段落）。
// 比對時把這些差異摺疊掉，字元一對一對應回原文位置。
const foldForMatching = (char) => {
  if (/[\s\u200b-\u200d\ufeff]/.test(char)) return "";
  if (/['‘’＇]/.test(char)) return "'";
  if (/["“”＂]/.test(char)) return '"';
  if (/[-‐-―−]/.test(char)) return "-";
  const lower = char.toLowerCase();
  return lower.length === 1 ? lower : char;
};

const buildFoldedIndex = (searchIndex) => {
  if (searchIndex.folded) return searchIndex.folded;
  const raw = searchIndex.text || "";
  const toRaw = [];
  let text = "";
  for (let i = 0; i < raw.length; i += 1) {
    const folded = foldForMatching(raw[i]);
    if (!folded) continue;
    text += folded;
    toRaw.push(i);
  }
  searchIndex.folded = { text, toRaw };
  return searchIndex.folded;
};

// 英文名詞不標在單字中間（test 不標 latest 的字尾）；中文沒有這個問題。
const isWholeWordSpan = (searchIndex, span) => {
  const raw = searchIndex.text || "";
  const isWordChar = (char) => Boolean(char) && /[A-Za-z0-9]/.test(char);
  const cutsStart = isWordChar(raw[span.startRaw]) && isWordChar(raw[span.startRaw - 1]);
  const cutsEnd = isWordChar(raw[span.endRaw - 1]) && isWordChar(raw[span.endRaw]);
  return !cutsStart && !cutsEnd;
};

const findFoldedSpan = (searchIndex, targetText, takenIntervals, accept) => {
  const needle = [...String(targetText || "")].map(foldForMatching).join("");
  if (!needle) return null;
  const { text, toRaw } = buildFoldedIndex(searchIndex);
  for (const start of findIndexesInText(text, needle)) {
    const span = { startRaw: toRaw[start], endRaw: toRaw[start + needle.length - 1] + 1 };
    if (takenIntervals.some((iv) => span.startRaw < iv.endRaw && span.endRaw > iv.startRaw)) continue;
    if (accept(searchIndex, span)) return span;
  }
  return null;
};

const findSpanForTextInIndex = (searchIndex, targetText, takenIntervals, accept = () => true) => {
  const normalizedNeedle = normalizeWhitespace(targetText);
  if (!normalizedNeedle || !searchIndex?.normalizedText) return null;
  const attempts = [
    { haystack: searchIndex.normalizedText, needle: normalizedNeedle },
    {
      haystack: searchIndex.normalizedTextLower,
      needle: normalizedNeedle.toLowerCase(),
    },
  ];
  for (const attempt of attempts) {
    const indexes = findIndexesInText(attempt.haystack, attempt.needle);
    for (const start of indexes) {
      const rawSpan = normalizedSpanToRawSpan(
        searchIndex.normalizedToRaw,
        start,
        attempt.needle.length,
        searchIndex.rawTextLength
      );
      if (!rawSpan) continue;
      const overlaps = takenIntervals.some(
        (iv) => rawSpan.startRaw < iv.endRaw && rawSpan.endRaw > iv.startRaw
      );
      if (overlaps || !accept(searchIndex, rawSpan)) continue;
      return rawSpan;
    }
  }
  return findFoldedSpan(searchIndex, targetText, takenIntervals, accept);
};

const applyAutoHighlightTasksBatch = async (tasks, usePreview) => {
  const targetPageKey = pageKey;
  const searchIndex = buildNormalizedSearchIndex();
  if (searchIndex) {
    searchIndex.normalizedTextLower = searchIndex.normalizedText.toLowerCase();
  }
  let skippedCount = 0;
  const located = [];
  const unmatched = [];
  const takenIntervals = [];
  if (searchIndex) {
    // 段落優先、長的優先：名詞常出現在重點句裡，不能讓名詞把整句擠掉。
    const ordered = [...tasks].sort((a, b) =>
      (a.kind === "term") - (b.kind === "term") || String(b.text).length - String(a.text).length
    );
    for (const task of ordered) {
      const accept = task.kind === "term" ? isWholeWordSpan : undefined;
      const span = findSpanForTextInIndex(searchIndex, task.text, takenIntervals, accept);
      if (!span) {
        skippedCount += 1;
        // 文字其實在頁面上，只是落在別的標記裡：不是「找不到」。
        task.covered = Boolean(findSpanForTextInIndex(searchIndex, task.text, [], accept));
        unmatched.push(task);
        continue;
      }
      takenIntervals.push(span);
      located.push({ task, span });
    }
  } else {
    skippedCount = tasks.length;
    unmatched.push(...tasks);
  }

  // Apply back-to-front: wrapping a later range only splits text nodes inside
  // it, so cached node references / offsets before it remain valid.
  located.sort((a, b) => b.span.startRaw - a.span.startRaw);

  const newEntries = [];
  let appliedCount = 0;
  for (const { task, span } of located) {
    try {
      const range = createRangeFromRawOffsets(
        searchIndex.nodes,
        span.startRaw,
        span.endRaw
      );
      if (!range) {
        skippedCount += 1;
        unmatched.push(task);
        continue;
      }
      const ancestor =
        range.commonAncestorContainer instanceof HTMLElement
          ? range.commonAncestorContainer
          : range.commonAncestorContainer?.parentElement;
      if (
        (ancestor && isEditableElement(ancestor)) ||
        rangeTouchesExistingHighlight(range)
      ) {
        skippedCount += 1;
        task.covered = rangeTouchesExistingHighlight(range);
        unmatched.push(task);
        continue;
      }
      if (usePreview) {
        const result = applyPreviewHighlight(range, task.color, task.reason, task.extra, task.source, task.kind);
        if (result) {
          previewData.push(result);
          appliedCount += 1;
        } else {
          skippedCount += 1;
          unmatched.push(task);
        }
        continue;
      }
      const snapshot = serializeRange(range.cloneRange());
      const text = normalizeWhitespace(snapshot.text);
      if (!text) {
        skippedCount += 1;
        unmatched.push(task);
        continue;
      }
      const highlightId = `hk-ai-${Date.now()}-${Math.floor(Math.random() * 100000)}`;
      const normalizedColor = toHexColor(task.color || DEFAULT_COLOR);
      const note = formatAiNote(task.reason, task.extra, task.source);
      wrapRangeWithHighlight(range, normalizedColor, highlightId);
      setAllMarksMetadata(highlightId, { color: normalizedColor, note, kind: task.kind });
      newEntries.push({
        id: highlightId,
        color: normalizedColor,
        text: snapshot.text,
        range: snapshot,
        url: targetPageKey,
        createdAt: Date.now(),
        note,
        ...(task.kind ? { kind: task.kind } : {}),
      });
      appliedCount += 1;
    } catch (error) {
      console.debug("套用 AI 重點失敗", error);
      skippedCount += 1;
    }
  }

  if (newEntries.length) {
    panelRefreshSuppressed = true;
    try {
      const existing = await getStoredHighlights(targetPageKey);
      await setStoredHighlights([...existing, ...newEntries], targetPageKey);
    } finally {
      panelRefreshSuppressed = false;
    }
  }
  return { appliedCount, skippedCount, unmatched };
};

const callOpenAI = async (key, prompt, options = {}) => {
  const systemPrompt =
    options.systemPrompt ||
    "You are a note-taking assistant who produces concise, easy-to-read study notes in Traditional Chinese.";
  const payload = {
    model: aiSettings.openaiModel || "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content: systemPrompt,
      },
      { role: "user", content: prompt },
    ],
    temperature: 0.4,
  };
  if (options.jsonMode) {
    payload.response_format = { type: "json_object" };
  }

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
    if (options.jsonMode && response.status === 400) {
      return callOpenAI(key, prompt, { ...options, jsonMode: false });
    }
    throw new Error(t("ai.errOpenaiApi", { error: errorText }));
  }

  const json = await response.json();
  const text = json?.choices?.[0]?.message?.content;
  if (!text) {
    throw new Error(t("ai.errOpenaiEmpty"));
  }
  return text.trim();
};

const callGemini = async (key, prompt, options = {}) => {
  const model = aiSettings.geminiModel || DEFAULT_GEMINI_MODEL;
  const payload = {
    contents: [
      {
        role: "user",
        parts: [{ text: prompt }],
      },
    ],
  };
  if (options.jsonMode) {
    payload.generationConfig = {
      responseMimeType: "application/json",
    };
  }
  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1/models/${model}:generateContent?key=${encodeURIComponent(
      key
    )}`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    if (response.status === 404 && model !== DEFAULT_GEMINI_MODEL) {
      aiSettings.geminiModel = DEFAULT_GEMINI_MODEL;
      await persistAISettings();
      return callGemini(key, prompt, options);
    }
    if (options.jsonMode && response.status === 400) {
      return callGemini(key, prompt, { ...options, jsonMode: false });
    }
    throw new Error(t("ai.errGeminiApi", { error: errorText }));
  }

  const json = await response.json();
  const parts = json?.candidates?.[0]?.content?.parts;
  const text = parts?.map((part) => part.text).join("\n");
  if (!text) {
    throw new Error(t("ai.errGeminiEmpty"));
  }
  return text.trim();
};

const saveGeneratedNote = async (pageUrl, notePayload) => {
  try {
    const stored = await chrome.storage?.local.get("hkGeneratedNotes");
    const previous = stored?.hkGeneratedNotes?.[pageUrl] || {};
    const nextNotes = {
      ...(stored?.hkGeneratedNotes ?? {}),
      [pageUrl]: { ...previous, ...notePayload, formatVersion: 2 },
    };
    await chrome.storage?.local.set({ hkGeneratedNotes: nextNotes });
  } catch (error) {
    console.debug("儲存筆記失敗", error);
  }
};

const clearPageHighlights = async () => {
  document.querySelectorAll(`.${HIGHLIGHT_CLASS}`).forEach(unwrapHighlightElement);
  previewData = [];
  previewUndo = null;
  hidePreviewConfirmBar();
  await setStoredHighlights([], pageKey);
  await refreshHighlightPanelIfVisible();
  setAiPanelStatus(t("ai.statusClearedAndReady"));
};

// 刪除側欄的單一項目：頁面標記、預覽與 AI 導覽清單都要一起移除，否則會以「不在頁面上」再出現。
const deleteGuideItem = async ({ text, kind }) => {
  const sameItem = (item) => {
    if (HkParsers.looseKey(item?.text || "") !== HkParsers.looseKey(text || "")) return false;
    return kind === "term" ? item.kind === "term" : item.kind !== "term";
  };
  const removeMarks = (id) => {
    document.querySelectorAll(`[${HIGHLIGHT_ATTR}="${id}"]`).forEach(unwrapHighlightElement);
  };
  const entries = await getStoredHighlights();
  entries.filter(sameItem).forEach((item) => removeMarks(item.id));
  if (entries.some(sameItem)) await setStoredHighlights(entries.filter((item) => !sameItem(item)));
  previewData = previewData.filter((item) => {
    if (!sameItem(item.data)) return true;
    removeMarks(item.data.id);
    return false;
  });
  if (previewData.length) showPreviewConfirmBar(previewData.length);
  else {
    previewUndo = null;
    hidePreviewConfirmBar();
  }
  const stored = await chrome.storage.local.get("hkGeneratedNotes");
  const note = stored?.hkGeneratedNotes?.[pageKey];
  if (note) {
    const keepTerm = (item) => !sameItem({ ...item, kind: "term" });
    const keepPassage = (item) => !sameItem({ ...item, kind: "passage" });
    const next = {
      ...note,
      guideTerms: (note.guideTerms || []).filter(keepTerm),
      unlocatedTerms: (note.unlocatedTerms || []).filter(keepTerm),
      guideHighlights: (note.guideHighlights || []).filter(keepPassage),
      unlocatedHighlights: (note.unlocatedHighlights || []).filter(keepPassage),
    };
    highlightPanelState.notesByPage = { ...highlightPanelState.notesByPage, [pageKey]: next };
    await chrome.storage.local.set({ hkGeneratedNotes: { ...stored.hkGeneratedNotes, [pageKey]: next } });
  }
  await refreshHighlightPanelIfVisible();
};

// 刪除這篇文章的全部筆記：標記、脈絡、名詞、地圖、標籤與個人筆記。
const deletePageNotes = async () => {
  document.querySelectorAll(`.${HIGHLIGHT_CLASS}`).forEach(unwrapHighlightElement);
  previewData = [];
  previewUndo = null;
  hidePreviewConfirmBar();
  destroyGuidedBar();
  const keys = [PAGE_META_KEY, "hkGeneratedNotes", MINDMAP_STORAGE_KEY, "hkPersonalNotes"];
  const stored = await chrome.storage.local.get(keys);
  const updates = {};
  keys.forEach((key) => {
    if (!stored[key]?.[pageKey]) return;
    const { [pageKey]: _removed, ...rest } = stored[key];
    updates[key] = rest;
  });
  if (Object.keys(updates).length) await chrome.storage.local.set(updates);
  await chrome.storage.local.remove(pageKey);
  const { [pageKey]: _note, ...notes } = highlightPanelState.notesByPage || {};
  highlightPanelState.notesByPage = notes;
  await refreshHighlightPanelIfVisible();
};

const applyPreviewHighlight = (range, color, reason, extra = "", source = "", kind) => {
  const snapshot = serializeRange(range.cloneRange());
  const text = normalizeWhitespace(snapshot.text);
  if (!text) return null;
  const highlightId = `hk-ai-preview-${Date.now()}-${Math.floor(Math.random() * 100000)}`;
  const normalizedColor = toHexColor(color || DEFAULT_COLOR);
  const note = formatAiNote(reason, extra, source);
  const highlightEl = wrapRangeWithHighlight(range, normalizedColor, highlightId);
  if (highlightEl) highlightEl.classList.add("hk-highlight-preview");
  setAllMarksMetadata(highlightId, { color: normalizedColor, note, kind });
  document.querySelectorAll(`[${HIGHLIGHT_ATTR}="${highlightId}"]`).forEach(el => el.classList.add("hk-highlight-preview"));
  return {
    el: highlightEl,
    data: {
      id: highlightId,
      color: normalizedColor,
      text: snapshot.text,
      range: snapshot,
      url: pageKey,
      createdAt: Date.now(),
      note,
      ...(kind ? { kind } : {}),
    },
  };
};

const showPreviewConfirmBar = (count) => {
  let bar = document.getElementById("hk-preview-bar");
  if (!bar) {
    bar = document.createElement("div");
    bar.id = "hk-preview-bar";
    bar.className = "hk-preview-bar";
    const msg = document.createElement("span");
    msg.id = "hk-preview-bar-msg";
    msg.className = "hk-preview-bar-msg";
    const confirmBtn = document.createElement("button");
    confirmBtn.type = "button";
    confirmBtn.className = "hk-preview-bar-confirm";
    confirmBtn.textContent = t("panel.previewConfirm");
    confirmBtn.addEventListener("click", () => confirmPreviewHighlights());
    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "hk-preview-bar-cancel";
    cancelBtn.textContent = t("panel.previewCancel");
    cancelBtn.addEventListener("click", () => cancelPreviewHighlights());
    bar.appendChild(msg);
    bar.appendChild(confirmBtn);
    bar.appendChild(cancelBtn);
    document.body.appendChild(bar);
  }
  const msg = document.getElementById("hk-preview-bar-msg");
  if (msg) msg.textContent = t("panel.previewMessage", { count });
  bar.hidden = false;
};

const hidePreviewConfirmBar = () => {
  const bar = document.getElementById("hk-preview-bar");
  if (bar) bar.hidden = true;
};

const confirmPreviewHighlights = async () => {
  const toSave = [...previewData];
  previewData = [];
  previewUndo = null;
  hidePreviewConfirmBar();
  const entries = [];
  toSave.forEach((item) => {
    document
      .querySelectorAll(`[${HIGHLIGHT_ATTR}="${item.data.id}"]`)
      .forEach((el) => el.classList.remove("hk-highlight-preview"));
    entries.push(item.data);
  });
  if (entries.length) {
    panelRefreshSuppressed = true;
    try {
      const existing = await getStoredHighlights();
      await setStoredHighlights([...existing, ...entries]);
    } finally {
      panelRefreshSuppressed = false;
    }
  }
  await ensurePageMetaTitle(pageKey, document.title);
  await refreshHighlightPanelIfVisible();
  setAiPanelStatus(t("ai.applied", { count: toSave.length }));
  notifyPageChanged();
};

const restorePreviewUndo = async (undo) => {
  const key = undo.pageKey;
  const stored = await chrome.storage.local.get(["hkGeneratedNotes", MINDMAP_STORAGE_KEY, PAGE_META_KEY]);
  const restore = (map, value) => {
    const next = { ...(map || {}) };
    if (value === undefined) delete next[key];
    else next[key] = value;
    return next;
  };
  const notes = restore(stored.hkGeneratedNotes, undo.note);
  highlightPanelState.notesByPage = notes;
  await chrome.storage.local.set({
    hkGeneratedNotes: notes,
    [MINDMAP_STORAGE_KEY]: restore(stored[MINDMAP_STORAGE_KEY], undo.mindmap),
    [PAGE_META_KEY]: restore(stored[PAGE_META_KEY], undo.meta),
  });
  // 重新套用時可能改過既有標記的註解，也一起還原。
  await setStoredHighlights(undo.entries, key);
  undo.entries.forEach((entry) => setAllMarksMetadata(entry.id, { note: entry.note, kind: entry.kind }));
};

// 側欄不一定收得到 storage 事件（內容沒變時），明確通知它重讀。
const notifyPageChanged = () => {
  chrome.runtime.sendMessage({ type: "HK_PAGE_CHANGED", pageKey }).catch(() => {});
};

const cancelPreviewHighlights = async () => {
  const toRemove = [...previewData];
  const undo = previewUndo;
  previewData = [];
  previewUndo = null;
  hidePreviewConfirmBar();
  toRemove.forEach((item) => {
    document.querySelectorAll(`[${HIGHLIGHT_ATTR}="${item.data.id}"]`).forEach(unwrapHighlightElement);
  });
  try {
    if (undo?.pageKey === pageKey) await restorePreviewUndo(undo);
  } catch (error) {
    console.debug("還原預覽前的筆記失敗", error);
  }
  await refreshHighlightPanelIfVisible();
  setAiPanelStatus(t("ai.statusPreviewCancelled"));
  notifyPageChanged();
};

// AI 回傳的「標籤」區塊：解析後併入目前頁面既有標籤（聯集，不覆蓋手動加的），
// 並同步面板狀態讓標籤即時顯示。回傳實際新增的標籤數。
const applyAiPageTags = async (rawTagsSection) => {
  const aiTags = HkParsers.parsePageTags(rawTagsSection);
  if (!aiTags.length) return 0;
  const existing = await getPageTags(pageKey);
  const merged = Array.from(new Set([...existing, ...aiTags]));
  await setPageTags(pageKey, merged);
  const prevMeta = highlightPanelState.pageMeta?.[pageKey] || {};
  highlightPanelState.pageMeta = {
    ...highlightPanelState.pageMeta,
    [pageKey]: { ...prevMeta, tags: merged },
  };
  return merged.length - existing.length;
};

const applyNotePayload = async (noteText, { context } = {}) => {
  const previous = highlightPanelState.notesByPage[pageKey] || {};
  const notePayload = {
    ...previous,
    ...(typeof noteText === "string" && noteText.trim() ? { note: noteText } : {}),
    ...(typeof context === "string" && context.trim() ? { context: context.trim() } : {}),
    provider: aiSettings.provider === "gemini-web" ? "gemini-web" : "chatgpt",
    model: aiSettings.provider === "gemini-web" ? "gemini-web" : "chatgpt-web",
    generatedAt: Date.now(),
    url: pageKey,
  };
  highlightPanelState.notesByPage = {
    ...highlightPanelState.notesByPage,
    [pageKey]: notePayload,
  };
  await saveGeneratedNote(pageKey, notePayload);
  return notePayload;
};

// 套用「分享的筆記」到當前頁面：標註走 AI 套用管線（文字重新定位、
// 保留顏色與註解），摘要／心智圖／標籤／標題各自寫入。
// payload: { title, tags, hl:[{text,note,color}], note, mindmap }
const applySharedPayload = async (payload, { overwrite = false } = {}) => {
  if (overwrite) {
    await clearPageHighlights();
  }
  let highlight = null;
  const items = Array.isArray(payload.hl) ? payload.hl : [];
  if (items.length) {
    const latestPalette = await refreshPaletteFromStorage();
    const usageCounts = await collectColorUsageCounts();
    const preferredPalette = sortPaletteByUsage(latestPalette, usageCounts);
    const tasks = normalizeAutoHighlightItems(
      items.map((item) => ({
        text: item.text,
        color: item.color,
        reason: item.note,
        kind: item.kind,
      })),
      preferredPalette
    );
    if (tasks.length) {
      setAiPanelStatus(t("ai.statusApplyingHighlights"));
      const { appliedCount, skippedCount } = await applyAutoHighlightTasksBatch(
        tasks,
        false // 分享匯入不走預覽，直接套用
      );
      highlight = { appliedCount, skippedCount, previewing: false };
    }
  }
  const noteText = (payload.note || "").trim();
  const context = typeof payload.context === "string" ? payload.context.trim() : "";
  if (noteText || context) {
    await applyNotePayload(noteText, { context });
  }
  if (typeof payload.personalNote === "string") {
    const stored = await chrome.storage.local.get("hkPersonalNotes");
    const personalNotes = { ...(stored?.hkPersonalNotes || {}) };
    personalNotes[pageKey] = { text: payload.personalNote, updatedAt: Date.now() };
    await chrome.storage.local.set({ hkPersonalNotes: personalNotes });
  }
  const outline = (payload.mindmap || "").trim();
  if (outline) {
    try {
      const tree = parseMindmapOutline(outline);
      if (tree) {
        await saveMindmap(outline, tree.title);
      }
    } catch (mapError) {
      console.debug("寫入分享心智圖失敗", mapError);
    }
  }
  let tagsAdded = 0;
  const tags = Array.isArray(payload.tags) ? payload.tags.filter(Boolean) : [];
  if (tags.length) {
    try {
      tagsAdded = await applyAiPageTags(tags.join(", "));
    } catch (tagError) {
      console.debug("匯入分享標籤失敗", tagError);
    }
  }
  await ensurePageMetaTitle(pageKey, payload.title || document.title);
  if (!highlight && !noteText && !context && typeof payload.personalNote !== "string") {
    throw new Error(t("ai.errUnrecognized"));
  }
  return { highlight, noteImported: Boolean(noteText), tagsAdded };
};

// 貼上的是別人分享的筆記 Markdown：轉成分享 payload 走同一條匯入管線。
const importSharedMarkdown = async (shared) => {
  const applied = await applySharedPayload({
    title: shared.title,
    tags: shared.tags,
    hl: shared.items.map((item) => ({ text: item.text, note: item.note, kind: item.kind })),
    note: shared.summary,
    context: shared.context,
    ...(typeof shared.personalNote === "string" ? { personalNote: shared.personalNote } : {}),
    mindmap: shared.outline,
  });
  const otherPage = shared.url && normalizePageKey(shared.url) !== pageKey;
  const summary = t("ai.appliedCount", { applied: applied.highlight?.appliedCount || 0 });
  return {
    success: true,
    shared: true,
    summary: (otherPage ? t("share.mdOtherPage") : "") + summary,
  };
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
  const visibleTabs = ["page", "search"];
  const resolved = visibleTabs.includes(highlightPanelState.activeTab)
    ? highlightPanelState.activeTab
    : "page";
  highlightPanelState.activeTab = resolved;

  const { tabButtons, tabPanels } = highlightPanelEls ?? {};
  visibleTabs.forEach((tab) => {
    const isActive = resolved === tab;
    const btn = tabButtons?.[tab];
    if (btn) {
      btn.classList.toggle("is-active", isActive);
      btn.setAttribute("aria-selected", isActive ? "true" : "false");
      btn.setAttribute("tabindex", isActive ? "0" : "-1");
    }
    const panel = tabPanels?.[tab];
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
    hint.textContent = t("panel.tagsEmpty");
    container.appendChild(hint);
    return;
  }

  const clearBtn = document.createElement("button");
  clearBtn.type = "button";
  clearBtn.className = "hk-panel-tag-chip hk-panel-tag-clear";
  clearBtn.textContent = t("panel.tagsAll");
  clearBtn.classList.toggle("is-active", !highlightPanelState.activeTag);
  clearBtn.addEventListener("click", () => {
    highlightPanelState.activeTag = null;
    updateHighlightPanelTagFilters();
    renderPanelViews();
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
      renderPanelViews();
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
    setPanelStatus(t("ai.statusHighlightDeleted"));
    await refreshHighlightPanelData();
    await renderHighlightPanel();
  } catch (error) {
    console.debug("刪除標註失敗", error);
    setPanelStatus(t("ai.errDeleteHighlight"), true);
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
    ...(entry.kind ? { kind: entry.kind } : {}),
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
    setPanelStatus(t("ai.errNoHighlightExport"), true);
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
      setPanelStatus(t("ai.statusTextCopied"));
    } catch (error) {
      console.debug("複製標註失敗", error);
      setPanelStatus(t("ai.errTextCopy"), true);
    }
    return;
  }

  const exportEntries = entries
    .map((entry) => buildHighlightExportEntry(entry))
    .filter(Boolean);
  if (!exportEntries.length) {
    setPanelStatus(t("ai.errNoHighlightExport"), true);
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
  setPanelStatus(t("ai.statusJsonDownloaded"));
};

const parseImportedHighlightsPayload = (rawText) => {
  let parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (_error) {
    throw new Error(t("ai.errJsonInvalid"));
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
    ...(entry.kind === "term" || entry.kind === "passage" ? { kind: entry.kind } : {}),
  };
  return normalized;
};

const importHighlightsFromEntries = async (rawEntries) => {
  if (!storage) throw new Error(t("ai.errStorageAccess"));
  const normalized = rawEntries
    .map((entry, index) => normalizeImportedHighlightEntry(entry, index))
    .filter(Boolean);
  if (!normalized.length) {
    throw new Error(t("ai.errNoHighlightImport"));
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
      throw new Error(t("ai.errNoHighlightInFile"));
    }
    const importedCount = await importHighlightsFromEntries(rawEntries);
    setPanelStatus(t("ai.importedCount", { count: importedCount }));
    await refreshHighlightPanelData();
    await renderHighlightPanel();
    attemptRestoreHighlights();
  } catch (error) {
    console.debug("匯入標註失敗", error);
    setPanelStatus(error.message || t("ai.errImportFailed"), true);
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
  pageTagHint.textContent = t("panel.tagInputHint");

  pageTagList.innerHTML = "";
  pageTagList.classList.remove("is-disabled");

  if (!currentPageTags.length) {
    const empty = document.createElement("span");
    empty.className = "hk-panel-tags-empty";
    empty.textContent = t("panel.tagEmptyState");
    pageTagList.appendChild(empty);
  } else {
    currentPageTags.forEach((tag) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "hk-panel-tag-chip";
      chip.textContent = tag;
      chip.title = t("panel.tagRemoveTitle");
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
  // Color shown as a left vertical bar — see .hk-panel-item::before
  item.style.setProperty("--hk-color", entry.color ?? DEFAULT_COLOR);

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
  deleteBtn.textContent = t("popup.deleteColor");
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

  const noteText = typeof entry.note === "string" ? entry.note.trim() : "";
  if (noteText) {
    const note = document.createElement("p");
    note.className = "hk-panel-item-note";
    // 註解可能有兩行（AI：重點 / 延伸：…）。分行渲染，延伸那行另外標色，
    // 掃列表時一眼分得出「這是原文在講什麼」與「這是額外補的」。
    const extraPrefix = t("ai.extraNotePrefix");
    const sourcePrefix = t("ai.sourceNotePrefix");
    HkParsers.breakNoteLines(noteText).split("\n").forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed) return;

      // 來源：模型憑記憶寫的，一律當成「線索」而非事實。
      // 做成搜尋連結——我們不信任模型給的網址，但可以幫他一鍵去查。
      if (trimmed.startsWith(sourcePrefix)) {
        const query = trimmed.slice(sourcePrefix.length).trim();
        if (!query) return;
        const link = document.createElement("a");
        link.className = "hk-note-line hk-note-line-source";
        link.href = `https://www.google.com/search?q=${encodeURIComponent(query)}`;
        link.target = "_blank";
        link.rel = "noopener";
        link.textContent = `${sourcePrefix}${query}`;
        link.title = t("ai.sourceSearchHint");
        link.addEventListener("click", (event) => event.stopPropagation());
        note.appendChild(link);
        return;
      }

      const row = document.createElement("span");
      row.className = trimmed.startsWith(extraPrefix)
        ? "hk-note-line hk-note-line-extra"
        : "hk-note-line";
      row.textContent = trimmed;
      note.appendChild(row);
    });
    item.appendChild(note);
  }

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

// 依目前的失聯數量，更新本頁清單上方的常駐提示。
const updateOrphanNotice = () => {
  const el = highlightPanelEls?.pageOrphanNotice;
  if (!el) return;
  if (orphanHighlightCount > 0) {
    el.textContent = t("panel.orphanNotice", { count: orphanHighlightCount });
    el.style.display = "block";
  } else {
    el.style.display = "none";
  }
};

const renderPageEntries = (entries) => {
  const list = highlightPanelEls?.pageList;
  const placeholder = highlightPanelEls?.pagePlaceholder;
  if (!list || !placeholder) return;
  updateOrphanNotice();
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
  const scrollTarget =
    translatedHighlightOverlays.get(id)?.range?.startContainer?.parentElement ||
    getHighlightVisualElement(element);
  scrollTarget?.scrollIntoView?.({ behavior: "smooth", block: "center" });
  const focusElement = getHighlightVisualElement(element);
  focusElement?.classList.add("hk-highlight-focus");
  window.setTimeout(() => {
    focusElement?.classList.remove("hk-highlight-focus");
  }, 1600);
};

// 管理頁「跳到原文」：寫入 hkFocusHighlight 後開分頁，本頁載入時撈出並聚焦該標註。
const FOCUS_HIGHLIGHT_KEY = "hkFocusHighlight";
const checkPendingFocusHighlight = async () => {
  if (!storage) return;
  try {
    const stored = await storage.get(FOCUS_HIGHLIGHT_KEY);
    const req = stored?.[FOCUS_HIGHLIGHT_KEY];
    if (!req || req.url !== pageKey || !req.id) return;
    if (req.at && Date.now() - req.at > 60000) {
      await storage.remove(FOCUS_HIGHLIGHT_KEY);
      return;
    }
    const element = document.querySelector(`[${HIGHLIGHT_ATTR}="${req.id}"]`);
    if (!element) return; // 還沒還原到，下次 restore 再試
    await storage.remove(FOCUS_HIGHLIGHT_KEY);
    focusHighlightElement(req.id);
  } catch (error) {
    console.debug("聚焦指定標註失敗", error);
  }
};

// 掃過所有頁面的標註，把舊的飽和色換成對應的新淺色。回傳換掉的筆數。
const recolorLegacyHighlights = async () => {
  if (!storage) throw new Error(t("ai.errStorageUnavailable"));
  const all = await storage.get(null);
  const patch = {};
  let changed = 0;

  Object.entries(all || {}).forEach(([key, value]) => {
    if (!/^https?:/i.test(key) || !Array.isArray(value)) return;
    let touched = false;
    const next = value.map((entry) => {
      const mapped = LEGACY_COLOR_MAP[toHexColor(entry?.color || "")];
      if (!mapped) return entry;
      touched = true;
      changed += 1;
      return { ...entry, color: mapped };
    });
    if (touched) patch[key] = next;
  });

  if (!changed) return 0;
  await storage.set(patch);

  // 本頁的標記直接改色，不用重新整理。
  document.querySelectorAll(`.${HIGHLIGHT_CLASS}`).forEach((el) => {
    const mapped = LEGACY_COLOR_MAP[toHexColor(el.dataset?.hkColor || "")];
    if (mapped) setHighlightMetadata(el, { color: mapped });
  });
  colorUsageCache = null;
  return changed;
};

// 設定抽屜裡的色票管理（原本住在 popup，重做後移進面板）。
const renderPanelPalette = () => {
  const listEl = highlightPanelEls?.paletteList;
  if (!listEl) return;
  listEl.innerHTML = "";

  if (!colorPalette.length) {
    const empty = document.createElement("p");
    empty.className = "hk-panel-palette-empty";
    empty.textContent = t("popup.emptyPalette");
    listEl.appendChild(empty);
    return;
  }

  colorPalette.forEach((color, index) => {
    const item = document.createElement("div");
    item.className = "hk-panel-swatch-item";
    if (color === currentColor) item.classList.add("is-active");

    // 點色塊 = 設為之後畫重點的預設色
    const useBtn = document.createElement("button");
    useBtn.type = "button";
    useBtn.className = "hk-panel-swatch";
    useBtn.style.backgroundColor = color;
    useBtn.title = t("popup.useColor", { color });
    useBtn.addEventListener("click", async () => {
      currentColor = color;
      try {
        await chrome.storage?.local.set({ hkLastColor: color });
      } catch (error) {
        console.debug("儲存預設顏色失敗", error);
      }
      renderPanelPalette();
    });

    const editInput = document.createElement("input");
    editInput.type = "color";
    editInput.className = "hk-panel-swatch-edit";
    editInput.value = color;
    editInput.title = t("popup.editColor", { color });
    editInput.addEventListener("change", async (event) => {
      const next = toHexColor(event.target.value);
      const nextPalette = [...colorPalette];
      nextPalette[index] = next;
      await persistColorPalette(nextPalette);
      renderPanelPalette();
      renderHighlightMenuSwatches();
    });

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "hk-panel-swatch-remove";
    removeBtn.textContent = "×";
    removeBtn.title = t("popup.deleteColorTitle", { color });
    removeBtn.addEventListener("click", async () => {
      await persistColorPalette(colorPalette.filter((_, i) => i !== index));
      renderPanelPalette();
      renderHighlightMenuSwatches();
    });

    item.appendChild(useBtn);
    item.appendChild(editInput);
    item.appendChild(removeBtn);
    listEl.appendChild(item);
  });
};

const ensureHighlightPanel = () => {
  if (highlightPanel) {
    if (!document.body.contains(highlightPanel)) {
      document.body.appendChild(highlightPanel);
    }
    return highlightPanel;
  }
  const panel = document.createElement("aside");
  panel.id = HIGHLIGHT_PANEL_ID;
  panel.className = "hk-page-panel";
  panel.style.display = "none";

  const header = document.createElement("div");
  header.className = "hk-panel-header";

  const title = document.createElement("h2");
  title.className = "hk-panel-title";
  title.textContent = t("panel.title");

  // Font controls live in the settings drawer now
  const fontControls = document.createElement("div");
  fontControls.className = "hk-panel-drawer-font-row";

  const fontDecreaseBtn = document.createElement("button");
  fontDecreaseBtn.type = "button";
  fontDecreaseBtn.className = "hk-panel-font-btn";
  fontDecreaseBtn.setAttribute("aria-label", t("panel.fontDecreaseTitle"));
  fontDecreaseBtn.textContent = "A−";
  fontDecreaseBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    adjustHighlightPanelFontScale(-PANEL_FONT_SCALE_STEP);
  });

  const fontIncreaseBtn = document.createElement("button");
  fontIncreaseBtn.type = "button";
  fontIncreaseBtn.className = "hk-panel-font-btn";
  fontIncreaseBtn.setAttribute("aria-label", t("panel.fontIncreaseTitle"));
  fontIncreaseBtn.textContent = "A+";
  fontIncreaseBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    adjustHighlightPanelFontScale(PANEL_FONT_SCALE_STEP);
  });

  fontControls.appendChild(fontDecreaseBtn);
  fontControls.appendChild(fontIncreaseBtn);

  const settingsBtn = document.createElement("button");
  settingsBtn.type = "button";
  settingsBtn.className = "hk-panel-iconbtn hk-panel-settings-btn";
  settingsBtn.setAttribute("aria-label", t("panel.settings"));
  settingsBtn.setAttribute("aria-expanded", "false");
  settingsBtn.innerHTML = "&#9881;"; // ⚙

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "hk-panel-close";
  closeBtn.setAttribute("aria-label", t("panel.closePanel"));
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => closeHighlightPanel());

  header.appendChild(title);
  header.appendChild(settingsBtn);
  header.appendChild(closeBtn);
  header.addEventListener("pointerdown", handlePanelPointerDown);
  header.addEventListener("pointermove", handlePanelPointerMove);
  header.addEventListener("pointerup", handlePanelPointerEnd);
  header.addEventListener("pointercancel", handlePanelPointerEnd);

  const tabs = document.createElement("div");
  tabs.className = "hk-panel-tabs";
  tabs.setAttribute("role", "tablist");
  tabs.setAttribute("aria-label", t("panel.tabsAria"));

  const pageTabBtn = document.createElement("button");
  pageTabBtn.type = "button";
  pageTabBtn.className = "hk-panel-tab-btn";
  pageTabBtn.id = "hk-panel-tab-page";
  pageTabBtn.setAttribute("role", "tab");
  pageTabBtn.setAttribute("aria-controls", "hk-panel-tabpanel-page");
  pageTabBtn.textContent = t("panel.tabPage");
  pageTabBtn.addEventListener("click", () => {
    if (highlightPanelState.activeTab === "page") return;
    highlightPanelState.activeTab = "page";
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
  searchTabBtn.textContent = t("panel.tagsAll");
  searchTabBtn.addEventListener("click", () => {
    if (highlightPanelState.activeTab === "search") return;
    highlightPanelState.activeTab = "search";
    applyHighlightPanelTabState();
    renderHighlightPanel().catch((error) =>
      console.debug("重新整理標註面板失敗", error)
    );
  });

  tabs.appendChild(pageTabBtn);
  tabs.appendChild(searchTabBtn);

  const pageTabPanel = document.createElement("div");
  pageTabPanel.className = "hk-panel-tab-content hk-panel-tab-content-page";
  pageTabPanel.id = "hk-panel-tabpanel-page";
  pageTabPanel.setAttribute("role", "tabpanel");
  pageTabPanel.setAttribute("aria-labelledby", pageTabBtn.id);

  const searchTabPanel = document.createElement("div");
  searchTabPanel.className = "hk-panel-tab-content hk-panel-tab-content-search";
  searchTabPanel.id = "hk-panel-tabpanel-search";
  searchTabPanel.setAttribute("role", "tabpanel");
  searchTabPanel.setAttribute("aria-labelledby", searchTabBtn.id);

  const searchControls = document.createElement("div");
  searchControls.className = "hk-panel-controls";

  const searchWrapper = document.createElement("label");
  searchWrapper.className = "hk-panel-search";
  searchWrapper.textContent = t("panel.searchHeader");
  const searchInput = document.createElement("input");
  searchInput.type = "search";
  searchInput.className = "hk-panel-search-input";
  searchInput.placeholder = t("panel.searchPlaceholder");
  let searchDebounceTimer = null;
  searchInput.addEventListener("input", (event) => {
    highlightPanelState.searchTerm = event.target.value ?? "";
    if (searchDebounceTimer) window.clearTimeout(searchDebounceTimer);
    searchDebounceTimer = window.setTimeout(() => {
      searchDebounceTimer = null;
      renderPanelViews();
    }, 150);
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
    renderPanelViews();
  });
  searchControls.appendChild(pageFilterSelect);

  const pageTagSection = document.createElement("div");
  pageTagSection.className = "hk-panel-page-tags";
  // Header redundant with drawer section title — keep only the hint span
  const pageTagHeader = document.createElement("div");
  pageTagHeader.className = "hk-panel-page-tags-header";
  const pageTagHint = document.createElement("span");
  pageTagHint.className = "hk-panel-page-tags-hint";
  pageTagHeader.appendChild(pageTagHint);
  pageTagSection.appendChild(pageTagHeader);

  const clearPageBtn = document.createElement("button");
  clearPageBtn.type = "button";
  clearPageBtn.className = "hk-panel-clear-page-btn";
  clearPageBtn.textContent = t("panel.clearPage");
  clearPageBtn.title = t("panel.clearPageTitle");
  clearPageBtn.addEventListener("click", async () => {
    if (!window.confirm(t("panel.clearConfirm"))) return;
    await clearPageHighlights();
  });
  pageTagSection.appendChild(clearPageBtn);

  const exportActions = document.createElement("div");
  exportActions.className = "hk-panel-export-actions";
  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "hk-panel-export-btn";
  copyBtn.textContent = t("panel.copyContent");
  copyBtn.addEventListener("click", () => exportHighlights("copy"));
  const downloadBtn = document.createElement("button");
  downloadBtn.type = "button";
  downloadBtn.className = "hk-panel-export-btn";
  downloadBtn.textContent = t("panel.downloadJson");
  downloadBtn.addEventListener("click", () => exportHighlights("download"));
  const importLabel = document.createElement("label");
  importLabel.className = "hk-panel-export-btn hk-panel-import-label";
  importLabel.textContent = t("panel.importJson");
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
  tagInput.placeholder = t("panel.tagInputPlaceholder");

  const suggestionDropdown = document.createElement("div");
  suggestionDropdown.className = "hk-panel-tag-suggestions";
  suggestionDropdown.setAttribute("role", "listbox");
  suggestionDropdown.setAttribute("aria-label", t("panel.tagSuggestions"));

  tagInputWrapper.appendChild(tagInput);
  tagInputWrapper.appendChild(suggestionDropdown);
  tagInputRow.appendChild(tagInputWrapper);

  const addTagBtn = document.createElement("button");
  addTagBtn.type = "button";
  addTagBtn.className = "hk-panel-tag-add";
  addTagBtn.textContent = t("panel.tagAdd");
  tagInputRow.appendChild(addTagBtn);

  pageTagSection.appendChild(tagInputRow);

  const pageTagList = document.createElement("div");
  pageTagList.className = "hk-panel-page-tag-list";
  pageTagSection.appendChild(pageTagList);

  const pageArchiveSection = document.createElement("div");
  pageArchiveSection.className = "hk-panel-archive";
  const pageArchiveHint = document.createElement("p");
  pageArchiveHint.className = "hk-panel-archive-hint";
  pageArchiveHint.textContent = t("panel.archivePageHint");
  const pageArchiveActions = document.createElement("div");
  pageArchiveActions.className = "hk-panel-archive-actions";
  const archiveCopyBtn = document.createElement("button");
  archiveCopyBtn.type = "button";
  archiveCopyBtn.className = "hk-panel-export-btn";
  archiveCopyBtn.textContent = t("panel.archiveCopy");
  archiveCopyBtn.addEventListener("click", () => exportHighlights("copy"));
  const archiveDownloadBtn = document.createElement("button");
  archiveDownloadBtn.type = "button";
  archiveDownloadBtn.className = "hk-panel-export-btn";
  archiveDownloadBtn.textContent = t("panel.archiveDownload");
  archiveDownloadBtn.addEventListener("click", () => exportHighlights("download"));
  const archiveImportLabel = document.createElement("label");
  archiveImportLabel.className = "hk-panel-export-btn hk-panel-import-label";
  archiveImportLabel.textContent = t("panel.archiveImport");
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
  archiveHint.textContent = t("panel.archiveAllHint");
  const archiveButtons = document.createElement("div");
  archiveButtons.className = "hk-panel-archive-actions";
  const archiveAllDownloadBtn = document.createElement("button");
  archiveAllDownloadBtn.type = "button";
  archiveAllDownloadBtn.className = "hk-panel-export-btn";
  archiveAllDownloadBtn.textContent = t("panel.archiveAllDownload");
  archiveAllDownloadBtn.addEventListener("click", handleDownloadAllHighlights);
  const archiveAllImportLabel = document.createElement("label");
  archiveAllImportLabel.className = "hk-panel-export-btn hk-panel-import-label";
  archiveAllImportLabel.textContent = t("panel.archiveAllImport");
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

  // ── Shared settings ──────────────────────────────────
  const aiConfigBlock = document.createElement("div");
  aiConfigBlock.className = "hk-panel-ai-config-block";

  const providerField = document.createElement("div");
  providerField.className = "hk-panel-ai-field";
  const providerLabel = document.createElement("label");
  providerLabel.className = "hk-panel-ai-label";
  const providerSelectId = "hk-panel-ai-provider";
  providerLabel.setAttribute("for", providerSelectId);
  providerLabel.textContent = t("panel.providerLabel");
  const aiProviderSelect = document.createElement("select");
  aiProviderSelect.id = providerSelectId;
  aiProviderSelect.className = "hk-panel-ai-select";
  [
    ["chatgpt", t("provider.chatgptWeb")],
    ["gemini-web", t("provider.geminiWeb")],
    ["openai", "OpenAI API"],
    ["gemini", "Google Gemini API"],
  ].forEach(([value, label]) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    aiProviderSelect.appendChild(option);
  });
  providerField.appendChild(providerLabel);
  providerField.appendChild(aiProviderSelect);
  aiConfigBlock.appendChild(providerField);

  const modelField = document.createElement("div");
  modelField.className = "hk-panel-ai-field";
  const modelLabel = document.createElement("label");
  const modelSelectId = "hk-panel-ai-model";
  modelLabel.setAttribute("for", modelSelectId);
  modelLabel.className = "hk-panel-ai-label";
  modelLabel.textContent = t("panel.modelLabel");
  const aiModelSelect = document.createElement("select");
  aiModelSelect.id = modelSelectId;
  aiModelSelect.className = "hk-panel-ai-select";
  modelField.appendChild(modelLabel);
  modelField.appendChild(aiModelSelect);
  aiConfigBlock.appendChild(modelField);

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

  const chatgptGroup = document.createElement("div");
  chatgptGroup.className = "hk-panel-ai-key-group";
  chatgptGroup.dataset.provider = "web";
  const chatgptInfo = document.createElement("p");
  chatgptInfo.className = "hk-panel-ai-chatgpt-info";
  chatgptInfo.textContent = t("panel.chatgptInfo");
  chatgptGroup.appendChild(chatgptInfo);
  chatgptGroup.hidden = true;

  aiKeyGroupsContainer.appendChild(openaiGroup);
  aiKeyGroupsContainer.appendChild(geminiGroup);
  aiKeyGroupsContainer.appendChild(chatgptGroup);
  aiConfigBlock.appendChild(aiKeyGroupsContainer);
  aiSettingsSection.appendChild(aiConfigBlock);

  // status (shared)
  // status banner — promoted to top of panel so users always see it
  const aiStatus = document.createElement("div");
  aiStatus.className = "hk-panel-ai-status hk-panel-status-banner";
  aiStatus.setAttribute("role", "status");
  aiStatus.setAttribute("aria-live", "polite");

  // ── 自動畫重點 block ─────────────────────────────────
  const aiHighlightBlock = document.createElement("div");
  aiHighlightBlock.className = "hk-panel-ai-block hk-panel-ai-hl-block";

  const aiHlHead = document.createElement("div");
  aiHlHead.className = "hk-panel-ai-block-head";
  const aiHlTitle = document.createElement("span");
  aiHlTitle.className = "hk-panel-ai-block-title";
  aiHlTitle.textContent = t("panel.autoHighlightTitle");
  const aiHlDesc = document.createElement("p");
  aiHlDesc.className = "hk-panel-ai-block-desc";
  aiHlDesc.textContent = t("panel.autoHighlightDesc");
  aiHlHead.appendChild(aiHlTitle);
  aiHlHead.appendChild(aiHlDesc);

  const autoHighlightPromptField = document.createElement("div");
  autoHighlightPromptField.className = "hk-panel-ai-field";
  const autoHighlightPromptLabel = document.createElement("label");
  const autoHighlightPromptTextareaId = "hk-panel-ai-auto-highlight-prompt";
  autoHighlightPromptLabel.setAttribute("for", autoHighlightPromptTextareaId);
  autoHighlightPromptLabel.className = "hk-panel-ai-label";
  autoHighlightPromptLabel.textContent = "Prompt";
  const autoHighlightPromptTextarea = document.createElement("textarea");
  autoHighlightPromptTextarea.id = autoHighlightPromptTextareaId;
  autoHighlightPromptTextarea.className = "hk-panel-ai-textarea";
  autoHighlightPromptTextarea.rows = 3;
  autoHighlightPromptField.appendChild(autoHighlightPromptLabel);
  autoHighlightPromptField.appendChild(autoHighlightPromptTextarea);

  // Category editor
  const catField = document.createElement("div");
  catField.className = "hk-panel-ai-field";
  const catLabel = document.createElement("label");
  catLabel.className = "hk-panel-ai-label";
  catLabel.textContent = t("panel.categoryLabel");
  const catList = document.createElement("div");
  catList.className = "hk-panel-ai-cat-list";
  const catAddBtn = document.createElement("button");
  catAddBtn.type = "button";
  catAddBtn.className = "hk-panel-ai-cat-add";
  catAddBtn.textContent = t("panel.addCategory");
  catAddBtn.addEventListener("click", () => {
    if (!Array.isArray(aiSettings.categories)) aiSettings.categories = [];
    aiSettings.categories.push({ name: "新分類", color: "#a8a8a8" });
    persistAISettings();
    renderCategoryList();
  });
  const catResetBtn = document.createElement("button");
  catResetBtn.type = "button";
  catResetBtn.className = "hk-panel-ai-cat-reset";
  catResetBtn.textContent = t("panel.resetCategories");
  catResetBtn.addEventListener("click", () => {
    aiSettings.categories = getDefaultCategories();
    persistAISettings();
    renderCategoryList();
  });
  catField.appendChild(catLabel);
  catField.appendChild(catList);
  catField.appendChild(catAddBtn);
  catField.appendChild(catResetBtn);

  // Options row
  const hlOptionsRow = document.createElement("div");
  hlOptionsRow.className = "hk-panel-ai-options-row";
  const previewCheckLabel = document.createElement("label");
  previewCheckLabel.className = "hk-panel-ai-check-row";
  const previewCheckbox = document.createElement("input");
  previewCheckbox.type = "checkbox";
  previewCheckbox.addEventListener("change", (e) => {
    aiSettings.usePreview = e.target.checked;
    persistAISettings();
  });
  previewCheckLabel.appendChild(previewCheckbox);
  previewCheckLabel.append("先預覽再確認");
  const selOnlyCheckLabel = document.createElement("label");
  selOnlyCheckLabel.className = "hk-panel-ai-check-row";
  const selOnlyCheckbox = document.createElement("input");
  selOnlyCheckbox.type = "checkbox";
  selOnlyCheckbox.addEventListener("change", (e) => {
    aiSettings.selectionOnly = e.target.checked;
    persistAISettings();
  });
  selOnlyCheckLabel.appendChild(selOnlyCheckbox);
  selOnlyCheckLabel.append("只標選取的文字");

  // 延伸補充：Prompt 會多要一個「===延伸===」區塊，接在摘要後面。
  const deepCheckLabel = document.createElement("label");
  deepCheckLabel.className = "hk-panel-ai-check-row";
  deepCheckLabel.title = t("panel.deepResearchHint");
  const deepCheckbox = document.createElement("input");
  deepCheckbox.type = "checkbox";
  deepCheckbox.addEventListener("change", (e) => {
    aiSettings.deepResearch = e.target.checked;
    persistAISettings();
  });
  deepCheckLabel.appendChild(deepCheckbox);
  deepCheckLabel.append(t("panel.deepResearch"));

  hlOptionsRow.appendChild(previewCheckLabel);
  hlOptionsRow.appendChild(selOnlyCheckLabel);
  hlOptionsRow.appendChild(deepCheckLabel);

  aiHighlightBlock.appendChild(aiHlHead);
  aiHighlightBlock.appendChild(catField);
  aiHighlightBlock.appendChild(autoHighlightPromptField);
  aiHighlightBlock.appendChild(hlOptionsRow);
  aiSettingsSection.appendChild(aiHighlightBlock);

  // ── AI 摘要筆記 block ────────────────────────────────
  const aiNoteGenBlock = document.createElement("div");
  aiNoteGenBlock.className = "hk-panel-ai-block hk-panel-ai-note-gen-block";

  const aiNoteGenHead = document.createElement("div");
  aiNoteGenHead.className = "hk-panel-ai-block-head";
  const aiNoteGenTitle = document.createElement("span");
  aiNoteGenTitle.className = "hk-panel-ai-block-title";
  aiNoteGenTitle.textContent = t("panel.noteSummaryTitle");
  const aiNoteGenDesc = document.createElement("p");
  aiNoteGenDesc.className = "hk-panel-ai-block-desc";
  aiNoteGenDesc.textContent = t("panel.noteSummaryDesc");
  aiNoteGenHead.appendChild(aiNoteGenTitle);
  aiNoteGenHead.appendChild(aiNoteGenDesc);

  const promptField = document.createElement("div");
  promptField.className = "hk-panel-ai-field";
  const promptLabel = document.createElement("label");
  const promptTextareaId = "hk-panel-ai-prompt";
  promptLabel.setAttribute("for", promptTextareaId);
  promptLabel.className = "hk-panel-ai-label";
  promptLabel.textContent = "Prompt";
  const promptTextarea = document.createElement("textarea");
  promptTextarea.id = promptTextareaId;
  promptTextarea.className = "hk-panel-ai-textarea";
  promptTextarea.rows = 4;
  promptField.appendChild(promptLabel);
  promptField.appendChild(promptTextarea);

  aiNoteGenBlock.appendChild(aiNoteGenHead);
  aiNoteGenBlock.appendChild(promptField);
  aiSettingsSection.appendChild(aiNoteGenBlock);

  // 失聯（orphan）標註提示：重整後找不到對應文字時常駐顯示在本頁清單上方。
  const pageOrphanNotice = document.createElement("div");
  pageOrphanNotice.className = "hk-panel-orphan-notice";
  pageOrphanNotice.style.display = "none";

  const pageList = document.createElement("div");
  pageList.className = "hk-panel-list hk-panel-page-list";

  const pagePlaceholder = document.createElement("p");
  pagePlaceholder.className = "hk-panel-placeholder";
  pagePlaceholder.textContent = t("panel.emptyPage");

  const searchList = document.createElement("div");
  searchList.className = "hk-panel-list hk-panel-search-list";

  const searchPlaceholder = document.createElement("p");
  searchPlaceholder.className = "hk-panel-placeholder";
  searchPlaceholder.textContent = t("panel.emptySearch");

  // 閱讀內容與 AI 流程都在側欄；這裡只放入口，面板專注在工具。
  const toolCard = document.createElement("div");
  toolCard.className = "hk-tool-card";

  const openGuideBtn = document.createElement("button");
  openGuideBtn.type = "button";
  openGuideBtn.className = "hk-tool-open-guide";
  openGuideBtn.textContent = t("panel.openGuide");
  openGuideBtn.addEventListener("click", () => {
    chrome.runtime.sendMessage({ type: "OPEN_SIDE_PANEL" }).catch((error) =>
      setAiPanelStatus(error?.message || t("sidebar.settingsError"), true)
    );
  });

  toolCard.appendChild(openGuideBtn);
  toolCard.appendChild(aiStatus);

  // ── 色票管理（原本在 popup，重做後放進設定抽屜）────────
  const paletteSection = document.createElement("div");
  paletteSection.className = "hk-panel-palette";

  const paletteList = document.createElement("div");
  paletteList.className = "hk-panel-palette-list";

  const paletteAddRow = document.createElement("div");
  paletteAddRow.className = "hk-panel-palette-add";

  const paletteAddInput = document.createElement("input");
  paletteAddInput.type = "color";
  paletteAddInput.className = "hk-panel-palette-input";
  paletteAddInput.value = toHexColor(currentColor || DEFAULT_COLOR);

  const paletteAddBtn = document.createElement("button");
  paletteAddBtn.type = "button";
  paletteAddBtn.className = "hk-panel-palette-add-btn";
  paletteAddBtn.textContent = t("popup.addToPalette");
  paletteAddBtn.addEventListener("click", async () => {
    const color = toHexColor(paletteAddInput.value);
    if (colorPalette.includes(color)) {
      setAiPanelStatus(t("popup.statusDuplicate"), true);
      return;
    }
    await persistColorPalette([...colorPalette, color]);
    currentColor = color;
    try {
      await chrome.storage?.local.set({ hkLastColor: color });
    } catch (error) {
      console.debug("儲存預設顏色失敗", error);
    }
    renderPanelPalette();
    renderHighlightMenuSwatches();
  });

  // 既有標註仍留著舊的飽和色 → 一鍵重上色（動到資料，先確認）。
  const recolorBtn = document.createElement("button");
  recolorBtn.type = "button";
  recolorBtn.className = "hk-panel-palette-recolor";
  recolorBtn.textContent = t("panel.recolorLegacy");
  recolorBtn.title = t("panel.recolorLegacyHint");
  recolorBtn.addEventListener("click", async () => {
    if (!window.confirm(t("panel.recolorConfirm"))) return;
    recolorBtn.disabled = true;
    try {
      const count = await recolorLegacyHighlights();
      setAiPanelStatus(
        count
          ? t("panel.recolorDone", { count })
          : t("panel.recolorNone"),
        !count
      );
      if (count) await refreshHighlightPanelIfVisible();
    } catch (error) {
      setAiPanelStatus(error?.message || t("ai.errStorageAccess"), true);
    } finally {
      recolorBtn.disabled = false;
    }
  });

  paletteAddRow.appendChild(paletteAddInput);
  paletteAddRow.appendChild(paletteAddBtn);
  paletteSection.appendChild(paletteList);
  paletteSection.appendChild(paletteAddRow);
  paletteSection.appendChild(recolorBtn);

  // ── Settings drawer (overlay, slides over the panel body) ─
  const drawer = document.createElement("div");
  drawer.className = "hk-panel-drawer";
  drawer.setAttribute("role", "region");
  drawer.setAttribute("aria-label", t("panel.settingsTitle"));

  const drawerHeader = document.createElement("div");
  drawerHeader.className = "hk-panel-drawer-header";

  const drawerTitle = document.createElement("h3");
  drawerTitle.className = "hk-panel-drawer-title";
  drawerTitle.textContent = t("panel.settingsTitle");

  const drawerCloseBtn = document.createElement("button");
  drawerCloseBtn.type = "button";
  drawerCloseBtn.className = "hk-panel-close";
  drawerCloseBtn.setAttribute("aria-label", t("panel.closeSettings"));
  drawerCloseBtn.textContent = "×";

  drawerHeader.appendChild(drawerTitle);
  drawerHeader.appendChild(drawerCloseBtn);

  const drawerBody = document.createElement("div");
  drawerBody.className = "hk-panel-drawer-body";

  const makeDrawerSection = (titleText, ...children) => {
    const section = document.createElement("section");
    section.className = "hk-panel-drawer-section";
    if (titleText) {
      const h = document.createElement("h4");
      h.className = "hk-panel-drawer-section-title";
      h.textContent = titleText;
      section.appendChild(h);
    }
    children.forEach((c) => c && section.appendChild(c));
    return section;
  };

  const displaySection = makeDrawerSection("顯示", fontControls);
  const paletteSectionWrap = makeDrawerSection(
    t("popup.paletteTitle"),
    paletteSection
  );
  const pageToolsSection = makeDrawerSection("本頁工具", pageTagSection);
  const pageArchiveSectionWrap = makeDrawerSection("本頁備份", pageArchiveSection);
  const allArchiveSectionWrap = makeDrawerSection("全部備份", archiveSection);
  const aiSettingsWrap = makeDrawerSection("AI 服務", aiSettingsSection);

  drawerBody.appendChild(displaySection);
  drawerBody.appendChild(paletteSectionWrap);
  drawerBody.appendChild(pageToolsSection);
  drawerBody.appendChild(pageArchiveSectionWrap);
  drawerBody.appendChild(allArchiveSectionWrap);
  drawerBody.appendChild(aiSettingsWrap);

  drawer.appendChild(drawerHeader);
  drawer.appendChild(drawerBody);

  const toggleDrawer = (open) => {
    const isOpen = typeof open === "boolean" ? open : !drawer.classList.contains("is-open");
    drawer.classList.toggle("is-open", isOpen);
    settingsBtn.setAttribute("aria-expanded", isOpen ? "true" : "false");
  };
  settingsBtn.addEventListener("click", () => toggleDrawer());
  drawerCloseBtn.addEventListener("click", () => toggleDrawer(false));

  // ── Assemble panel ─────────────────────────────────────
  panel.appendChild(header);
  panel.appendChild(toolCard);
  panel.appendChild(tabs);
  pageTabPanel.appendChild(pageOrphanNotice);
  pageTabPanel.appendChild(pageList);
  pageTabPanel.appendChild(pagePlaceholder);
  searchTabPanel.appendChild(searchControls);
  searchTabPanel.appendChild(searchList);
  searchTabPanel.appendChild(searchPlaceholder);
  panel.appendChild(pageTabPanel);
  panel.appendChild(searchTabPanel);
  panel.appendChild(drawer);

  highlightPanel = panel;
  highlightPanelEls = {
    container: panel,
    dragHandle: header,
    settingsBtn,
    drawer,
    drawerCloseBtn,
    toggleDrawer,
    paletteList,
    paletteAddInput,
    tabs,
    pageList,
    pagePlaceholder,
    pageOrphanNotice,
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
    aiProviderSelect,
    aiModelSelect,
    aiPromptField: promptTextarea,
    aiAutoHighlightPromptField: autoHighlightPromptTextarea,
    aiOpenaiKeyInput: openaiInput,
    aiGeminiKeyInput: geminiInput,
    aiKeyGroups: [openaiGroup, geminiGroup, chatgptGroup],
    aiWebInfo: chatgptInfo,
    aiStatus,
    aiCatList: catList,
    aiPreviewCheckbox: previewCheckbox,
    aiSelOnlyCheckbox: selOnlyCheckbox,
    aiDeepResearchCheckbox: deepCheckbox,
    fontControls: {
      decrease: fontDecreaseBtn,
      increase: fontIncreaseBtn,
    },
    tabButtons: {
      page: pageTabBtn,
      search: searchTabBtn,
    },
    tabPanels: {
      page: pageTabPanel,
      search: searchTabPanel,
    },
  };

  aiProviderSelect.addEventListener("change", (event) => {
    const v = event.target.value;
    const value = ["gemini", "chatgpt", "gemini-web"].includes(v) ? v : "openai";
    aiSettings.provider = value;
    populateAiModelSelect();
    updateAiKeyVisibility();
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
    schedulePersistAISettings();
  });

  geminiInput.addEventListener("input", (event) => {
    aiSettings.geminiKey = event.target.value;
    schedulePersistAISettings();
  });

  promptTextarea.addEventListener("input", (event) => {
    aiSettings.prompt = event.target.value;
    schedulePersistAISettings();
  });

  autoHighlightPromptTextarea.addEventListener("input", (event) => {
    aiSettings.autoHighlightPrompt = event.target.value;
    schedulePersistAISettings();
  });

  document.body.appendChild(panel);
  applyHighlightPanelSideClasses();
  applyHighlightPanelTabState();
  updateExportButtonsState();
  applyAiSettingsToUI();
  applyHighlightPanelFontScale();
  renderPanelPalette();
  refreshPaletteFromStorage()
    .then(() => renderPanelPalette())
    .catch(() => {});
  setAiPanelStatus("");
  return panel;
};

// View-only pass: renders lists from cached state without touching storage.
// Used for search keystrokes / filter changes, where re-reading all of
// chrome.storage on every input made the panel feel sluggish.
const renderPanelViews = () => {
  if (!highlightPanelEls) return;

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

  const searchInputEl = highlightPanelEls?.searchInput;
  if (searchInputEl && searchInputEl.value !== highlightPanelState.searchTerm) {
    searchInputEl.value = highlightPanelState.searchTerm;
  }

  applyHighlightPanelTabState();
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
  renderPanelViews();
};

// Coalesces storage-change re-renders into one render per burst.
const schedulePanelRefresh = () => {
  if (!highlightPanelVisible || panelRefreshSuppressed) return;
  if (panelRefreshDebounceTimer) window.clearTimeout(panelRefreshDebounceTimer);
  panelRefreshDebounceTimer = window.setTimeout(() => {
    panelRefreshDebounceTimer = null;
    if (!highlightPanelVisible || panelRefreshSuppressed) return;
    renderHighlightPanel().catch((error) =>
      console.debug("重新整理標註面板失敗", error)
    );
  }, 180);
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

  if (view === "settings") {
    highlightPanelEls?.toggleDrawer?.(true);
  } else if (view === "search") {
    highlightPanelState.activeTab = "search";
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

// 語言切換時整個重建面板，讓所有 JS 建立的標籤跟著新語言；保留開啟狀態。
const rebuildHighlightPanelForLang = () => {
  if (!highlightPanel) return;
  const wasVisible = highlightPanelVisible;
  try {
    highlightPanel.remove();
  } catch (_e) {}
  highlightPanel = null;
  highlightPanelEls = null;
  if (wasVisible) {
    openHighlightPanel().catch((error) =>
      console.debug("語言切換重建面板失敗", error)
    );
  }
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
  setAllMarksMetadata(activeHighlightId, {
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
    setHighlightMenuStatus(t("menu.colorUpdated"));
    await refreshHighlightPanelIfVisible();
  } catch (error) {
    console.debug("更新 highlight 顏色失敗", error);
    setHighlightMenuStatus(t("menu.colorUpdateFailed"), true);
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
  setAllMarksMetadata(activeHighlightId, { color: normalizedColor, note: trimmedNote });
  try {
    await updateHighlightEntry(activeHighlightId, { note: trimmedNote });
    if (highlightMenuEls.noteField) {
      highlightMenuEls.noteField.value = trimmedNote;
    }
    setHighlightMenuStatus(t("menu.noteSaved"));
    await refreshHighlightPanelIfVisible();
  } catch (error) {
    console.debug("儲存 highlight 註解失敗", error);
    setHighlightMenuStatus(t("menu.noteSaveFailed"), true);
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
  header.textContent = t("detail.heading");
  container.appendChild(header);

  const colorWrapper = document.createElement("div");
  colorWrapper.className = "hk-menu-section hk-menu-colors";

  const colorLabel = document.createElement("label");
  colorLabel.className = "hk-menu-label";
  colorLabel.textContent = t("detail.colorLabel");

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
  noteLabel.textContent = t("detail.noteLabel");
  const menuTranslateBtn = document.createElement("button");
  menuTranslateBtn.type = "button";
  menuTranslateBtn.className = "hk-menu-translate-btn";
  menuTranslateBtn.textContent = t("detail.translate");
  menuTranslateBtn.title = t("detail.translateTitle");
  menuTranslateBtn.addEventListener("click", () => {
    if (!activeHighlight) return;
    const text = getHighlightText(activeHighlight);
    if (!text) return;
    const rect = getHighlightVisualRect(activeHighlight);
    showTranslateCard(text, rect, activeHighlight);
  });
  noteLabel.appendChild(menuTranslateBtn);
  const noteField = document.createElement("textarea");
  noteField.className = "hk-menu-note";
  noteField.rows = 6;
  noteField.placeholder = t("detail.notePlaceholder");
  noteWrapper.appendChild(noteLabel);
  noteWrapper.appendChild(noteField);
  container.appendChild(noteWrapper);

  const actions = document.createElement("div");
  actions.className = "hk-menu-actions";
  const saveNoteBtn = document.createElement("button");
  saveNoteBtn.type = "button";
  saveNoteBtn.className = "hk-menu-btn hk-menu-btn-primary";
  saveNoteBtn.textContent = t("detail.saveNote");
  saveNoteBtn.addEventListener("click", handleSaveNote);

  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "hk-menu-btn hk-menu-btn-danger";
  deleteBtn.textContent = t("detail.deleteHighlight");
  deleteBtn.addEventListener("click", handleDeleteHighlight);

  actions.appendChild(saveNoteBtn);
  actions.appendChild(deleteBtn);
  container.appendChild(actions);

  const status = document.createElement("div");
  status.className = "hk-menu-status";
  container.appendChild(status);

  colorInput.addEventListener("input", (event) => {
    const nextColor = toHexColor(event.target.value || DEFAULT_COLOR);
    if (activeHighlight && activeHighlightId) {
      setAllMarksMetadata(activeHighlightId, {
        color: nextColor,
        note: activeHighlight.dataset.hkNote ?? "",
      });
    }
    currentColor = nextColor;
  });
  colorInput.addEventListener("change", (event) => {
    applyColorChange(event.target.value);
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
    const rect = getHighlightVisualRect(highlightEl);
    if (!rect) return;
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
  // 任一頁面標註陣列在別處變動 → 失效色彩統計快取
  if (Object.keys(changes).some((k) => /^https?:/i.test(k))) {
    colorUsageCache = null;
  }
  // 管理頁要求聚焦本頁某標註（分頁已開著的情況）
  if (changes[FOCUS_HIGHLIGHT_KEY]?.newValue) {
    checkPendingFocusHighlight();
  }
  if (changes.hkLastColor) {
    const nextColor = changes.hkLastColor.newValue;
    if (typeof nextColor === "string" && nextColor.trim()) {
      currentColor = toHexColor(nextColor);
    } else {
      currentColor = DEFAULT_COLOR;
    }
    renderPanelPalette();
  }
  if (changes.hkColorPalette) {
    setColorPaletteState(changes.hkColorPalette.newValue);
    renderPanelPalette();
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
  }
  {
    const affectedKeys = Object.keys(changes);
    const shouldRefresh =
      affectedKeys.includes(pageKey) ||
      affectedKeys.some((key) => isValidPageKey(key));
    if (shouldRefresh) {
      schedulePanelRefresh();
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
  const expected = normalizeWhitespace(snapshot.text ?? "");

  const rangeTextMatches = (range) => {
    if (!range || range.collapsed) return false;
    if (!expected) return true;
    return normalizeWhitespace(range.toString()) === expected;
  };

  const directRange = deserializeRange(snapshot);
  if (rangeTextMatches(directRange)) {
    return { range: directRange, snapshot, updated: false };
  }

  const cssRange = resolveRangeFromCssAnchors(snapshot.anchors, snapshot.text?.length);
  if (rangeTextMatches(cssRange)) {
    const normalized = serializeRange(cssRange.cloneRange());
    return { range: cssRange, snapshot: normalized, updated: true };
  }

  const textRange = resolveRangeFromTextAnchors(snapshot.anchors);
  if (textRange && !textRange.collapsed) {
    const normalized = serializeRange(textRange.cloneRange());
    return { range: textRange, snapshot: normalized, updated: true };
  }

  // Last resort: text search using snapshot.text directly (handles missing/stale anchors)
  if (expected) {
    const syntheticAnchors = {
      quote: {
        exact: snapshot.text,
        prefix: "",
        suffix: "",
      },
    };
    const fallbackRange = resolveRangeFromTextAnchors(syntheticAnchors);
    if (fallbackRange && !fallbackRange.collapsed) {
      const normalized = serializeRange(fallbackRange.cloneRange());
      return { range: fallbackRange, snapshot: normalized, updated: true };
    }
  }

  return { range: null, snapshot, updated: false };
};

const translatedHighlightOverlays = new Map();
const TRANSLATED_OVERLAY_GROUP_CLASS = "hk-highlight-overlay-group";
const TRANSLATED_OVERLAY_CLASS = "hk-highlight-overlay";
const translatedCssHighlightColors = new Map();
let translatedCssHighlightStyle = null;

const getTranslatedVisualRanges = (range) => {
  const getLayer = HkHighlightDom?.getTranslationLayer;
  const getScript = HkHighlightDom?.getDominantTextScript;
  if (typeof getLayer !== "function") return [range.cloneRange()];
  const textNodes = getTextNodesInRange(range);
  const layeredNodes = textNodes.filter((node) => getLayer(node));
  const translatedDocument = Boolean(
    HkHighlightDom?.isGoogleTranslatedDocument?.(document)
  );
  if (!layeredNodes.length && !translatedDocument) return [range.cloneRange()];

  const targetLayer = getLayer(range.startContainer);
  const startText =
    range.startContainer?.nodeType === Node.TEXT_NODE
      ? range.startContainer.nodeValue?.slice(range.startOffset)
      : textNodes[0]?.nodeValue || "";
  const targetScript =
    typeof getScript === "function" ? getScript(startText) : "other";
  const ranges = [];
  textNodes.forEach((node) => {
    if ((getLayer(node) || null) !== (targetLayer || null)) return;
    if (translatedDocument && targetScript !== "other" && typeof getScript === "function") {
      const nodeScript = getScript(node.nodeValue || "");
      if (nodeScript !== "other" && nodeScript !== targetScript) return;
    }
    let start = 0;
    let end = node.length;
    if (node === range.startContainer && range.startContainer.nodeType === Node.TEXT_NODE) {
      start = range.startOffset;
    }
    if (node === range.endContainer && range.endContainer.nodeType === Node.TEXT_NODE) {
      end = range.endOffset;
    }
    if (start >= end) return;
    const visualRange = document.createRange();
    visualRange.setStart(node, start);
    visualRange.setEnd(node, end);
    ranges.push(visualRange);
  });
  return ranges.length ? ranges : [range.cloneRange()];
};

const getTranslatedHighlightRects = (rangeOrRanges) => {
  const ranges = Array.isArray(rangeOrRanges)
    ? rangeOrRanges
    : [rangeOrRanges];
  return ranges.flatMap((range) =>
    Array.from(range?.getClientRects?.() ?? []).filter(
      (rect) => rect.width > 0 && rect.height > 0
    )
  );
};

const getTranslatedHighlightBounds = (entry) => {
  const rects = getTranslatedHighlightRects(
    entry?.visualRanges || entry?.range
  );
  if (!rects.length) return null;
  const left = Math.min(...rects.map((rect) => rect.left));
  const top = Math.min(...rects.map((rect) => rect.top));
  const right = Math.max(...rects.map((rect) => rect.right));
  const bottom = Math.max(...rects.map((rect) => rect.bottom));
  return {
    bottom,
    height: bottom - top,
    left,
    right,
    top,
    width: right - left,
  };
};

const supportsCssHighlights = () =>
  Boolean(
    globalThis.CSS?.highlights && typeof globalThis.Highlight === "function"
  );

const getTranslatedCssHighlightName = (id) =>
  `hkText_${String(id).replace(/[^a-zA-Z0-9_-]/g, "_")}`;

const rebuildTranslatedCssHighlightStyle = () => {
  if (!translatedCssHighlightColors.size) {
    translatedCssHighlightStyle?.remove();
    translatedCssHighlightStyle = null;
    return;
  }
  if (!translatedCssHighlightStyle) {
    translatedCssHighlightStyle = document.createElement("style");
    translatedCssHighlightStyle.setAttribute("translate", "no");
    translatedCssHighlightStyle.setAttribute("data-translate", "no");
    (document.head || document.documentElement).appendChild(
      translatedCssHighlightStyle
    );
  }
  translatedCssHighlightStyle.textContent = Array.from(
    translatedCssHighlightColors,
    ([name, color]) =>
      `::highlight(${name}) { background-color: ${toHexColor(
        color
      )}; color: inherit; }`
  ).join("\n");
};

const renderTranslatedHighlightOverlay = (entry) => {
  if (!entry?.group || !entry.range) return false;
  if (entry.mode === "css") {
    translatedCssHighlightColors.set(entry.cssName, entry.color);
    rebuildTranslatedCssHighlightStyle();
    return true;
  }
  const rects = getTranslatedHighlightRects(
    entry.visualRanges || entry.range
  );
  entry.shadowRoot?.replaceChildren();
  rects.forEach((rect) => {
    const overlay = document.createElement("span");
    overlay.className = `${HIGHLIGHT_CLASS} ${TRANSLATED_OVERLAY_CLASS}`;
    overlay.setAttribute(HIGHLIGHT_ATTR, entry.id);
    overlay.setAttribute("aria-hidden", "true");
    overlay.style.setProperty("position", "fixed", "important");
    overlay.style.setProperty("left", `${rect.left}px`, "important");
    overlay.style.setProperty("top", `${rect.top}px`, "important");
    overlay.style.setProperty("width", `${rect.width}px`, "important");
    overlay.style.setProperty("height", `${rect.height}px`, "important");
    overlay.style.setProperty("padding", "0", "important");
    overlay.style.setProperty("margin", "0", "important");
    overlay.style.setProperty("display", "block", "important");
    overlay.style.setProperty("pointer-events", "none", "important");
    overlay.style.setProperty("z-index", "2147483646", "important");
    overlay.style.setProperty("box-sizing", "border-box", "important");
    overlay.style.setProperty("border-radius", "2px", "important");
    overlay.style.setProperty("background-color", entry.color, "important");
    setHighlightMetadata(overlay, {
      color: entry.color,
      note: entry.note ?? "",
    });
    entry.shadowRoot?.appendChild(overlay);
  });
  return rects.length > 0;
};

const createTranslatedHighlightOverlay = (range, color, id) => {
  const group = document.createElement("span");
  group.className = `${HIGHLIGHT_CLASS} ${TRANSLATED_OVERLAY_GROUP_CLASS}`;
  group.setAttribute(HIGHLIGHT_ATTR, id);
  group.setAttribute("data-hk-overlay", "true");
  group.setAttribute("translate", "no");
  group.setAttribute("data-translate", "no");
  group.setAttribute("aria-hidden", "true");
  group.style.setProperty("position", "fixed", "important");
  group.style.setProperty("left", "0", "important");
  group.style.setProperty("top", "0", "important");
  group.style.setProperty("width", "0", "important");
  group.style.setProperty("height", "0", "important");
  group.style.setProperty("pointer-events", "none", "important");
  group.style.setProperty("z-index", "2147483646", "important");
  group.style.setProperty("contain", "strict", "important");

  const entry = {
    color,
    group,
    id,
    note: "",
    range: range.cloneRange(),
    visualRanges: getTranslatedVisualRanges(range),
    mode: supportsCssHighlights() ? "css" : "shadow",
    cssName: getTranslatedCssHighlightName(id),
    shadowRoot: null,
  };
  group.dataset.hkText = entry.visualRanges
    .map((visualRange) => visualRange.toString())
    .join("")
    .trim();
  if (entry.mode === "css") {
    try {
      entry.cssHighlight = new globalThis.Highlight(...entry.visualRanges);
      globalThis.CSS.highlights.set(entry.cssName, entry.cssHighlight);
    } catch (_error) {
      entry.mode = "shadow";
    }
  }
  if (entry.mode === "shadow") {
    entry.shadowRoot = group.attachShadow({ mode: "open" });
  }
  translatedHighlightOverlays.set(id, entry);
  document.documentElement.appendChild(group);
  if (!renderTranslatedHighlightOverlay(entry)) {
    if (entry.mode === "css") {
      globalThis.CSS.highlights.delete(entry.cssName);
      translatedCssHighlightColors.delete(entry.cssName);
      rebuildTranslatedCssHighlightStyle();
    }
    translatedHighlightOverlays.delete(id);
    group.remove();
    return null;
  }
  return group;
};

const updateTranslatedHighlightOverlays = () => {
  translatedHighlightOverlays.forEach((entry, id) => {
    if (!entry.group.isConnected) {
      translatedHighlightOverlays.delete(id);
      return;
    }
    if (entry.mode !== "css") renderTranslatedHighlightOverlay(entry);
  });
};

const getHighlightVisualElement = (highlightEl) => {
  if (highlightEl?.classList?.contains(TRANSLATED_OVERLAY_GROUP_CLASS)) {
    const entry = translatedHighlightOverlays.get(
      highlightEl.getAttribute(HIGHLIGHT_ATTR)
    );
    if (entry?.mode === "shadow") {
      return entry.shadowRoot?.querySelector(`.${TRANSLATED_OVERLAY_CLASS}`);
    }
    return highlightEl;
  }
  return highlightEl;
};

const getHighlightVisualRect = (highlightEl) => {
  const entry = translatedHighlightOverlays.get(
    highlightEl?.getAttribute?.(HIGHLIGHT_ATTR)
  );
  if (entry?.mode === "css") return getTranslatedHighlightBounds(entry);
  return getHighlightVisualElement(highlightEl)?.getBoundingClientRect?.();
};

const getHighlightText = (highlightEl) =>
  highlightEl?.dataset?.hkText || highlightEl?.textContent?.trim() || "";

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

const setHighlightMetadata = (element, { color, note, kind }) => {
  if (!element) return;
  const overlayId = element.getAttribute?.(HIGHLIGHT_ATTR);
  const overlayEntry = overlayId
    ? translatedHighlightOverlays.get(overlayId)
    : null;
  if (color) {
    element.style.backgroundColor = color;
    element.style.setProperty("--hk-highlight-color", color);
    if (overlayEntry) overlayEntry.color = color;
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
    } else {
      delete element.dataset.hkNote;
      element.removeAttribute("data-highlight-note");
    }
    if (overlayEntry) overlayEntry.note = trimmed;
  }
  if (kind === "term" || kind === "passage") {
    element.dataset.hkKind = kind;
    element.setAttribute("data-hk-kind", kind);
  }
  if (overlayEntry && element === overlayEntry.group) {
    renderTranslatedHighlightOverlay(overlayEntry);
  }
};

const setAllMarksMetadata = (id, meta) => {
  document.querySelectorAll(`[${HIGHLIGHT_ATTR}="${id}"]`).forEach((el) => {
    setHighlightMetadata(el, meta);
  });
};

const unwrapHighlightElement = (highlightEl) => {
  if (!highlightEl) return;
  const overlayGroup = highlightEl.classList?.contains(
    TRANSLATED_OVERLAY_GROUP_CLASS
  )
    ? highlightEl
    : highlightEl.closest?.(`.${TRANSLATED_OVERLAY_GROUP_CLASS}`);
  if (overlayGroup) {
    const id = overlayGroup.getAttribute(HIGHLIGHT_ATTR);
    const entry = id ? translatedHighlightOverlays.get(id) : null;
    if (entry?.mode === "css") {
      globalThis.CSS?.highlights?.delete(entry.cssName);
      translatedCssHighlightColors.delete(entry.cssName);
      rebuildTranslatedCssHighlightStyle();
    }
    if (id) translatedHighlightOverlays.delete(id);
    overlayGroup.remove();
    return;
  }
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

const getTextNodesInRange = (range) => {
  const ancestor = range.commonAncestorContainer;
  if (ancestor.nodeType === Node.TEXT_NODE) return [ancestor];
  const nodes = [];
  const walker = document.createTreeWalker(ancestor, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {
    if (range.intersectsNode(node)) nodes.push(node);
  }
  return nodes;
};

const wrapRangeTextNodeWalking = (range, color, id) => {
  const textNodes = getTextNodesInRange(range);
  if (!textNodes.length) return null;
  let firstMark = null;
  for (let i = 0; i < textNodes.length; i++) {
    const node = textNodes[i];
    let start = 0;
    let end = node.length;
    if (node === range.startContainer && range.startContainer.nodeType === Node.TEXT_NODE) {
      start = range.startOffset;
    }
    if (node === range.endContainer && range.endContainer.nodeType === Node.TEXT_NODE) {
      end = range.endOffset;
    }
    if (start >= end) continue;
    if (end < node.length) node.splitText(end);
    const selectedNode = start > 0 ? node.splitText(start) : node;
    const mark = createHighlightElement(color, id);
    selectedNode.parentNode.insertBefore(mark, selectedNode);
    mark.appendChild(selectedNode);
    if (!firstMark) firstMark = mark;
  }
  return firstMark;
};

const BLOCK_TAGS = new Set([
  "P", "DIV", "LI", "ARTICLE", "SECTION", "BLOCKQUOTE",
  "H1", "H2", "H3", "H4", "H5", "H6",
  "TD", "TH", "HEADER", "FOOTER", "MAIN", "ASIDE", "FIGURE",
]);

const nearestBlockAncestor = (node) => {
  let el = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
  while (el && el !== document.body) {
    if (BLOCK_TAGS.has(el.tagName)) return el;
    el = el.parentElement;
  }
  return document.body;
};

// 把 range 兩端的空白（含換行、不斷行空白、全形空白）往內縮，避免標註背景與
// 註解點線下線「多跑出來一點點」：通常是選取或 AI 比對時把尾端空白也圈了進來。
const HL_WS = new Set([" ", "\t", "\n", "\r", " ", "　", "​"]);
const trimRangeWhitespace = (range) => {
  let guard = 0;
  while (guard++ < 10000) {
    const { startContainer, startOffset } = range;
    if (
      startContainer.nodeType === Node.TEXT_NODE &&
      startOffset < (startContainer.nodeValue?.length ?? 0) &&
      HL_WS.has(startContainer.nodeValue[startOffset])
    ) {
      range.setStart(startContainer, startOffset + 1);
    } else break;
  }
  guard = 0;
  while (guard++ < 10000) {
    const { endContainer, endOffset } = range;
    if (
      endContainer.nodeType === Node.TEXT_NODE &&
      endOffset > 0 &&
      HL_WS.has(endContainer.nodeValue?.[endOffset - 1])
    ) {
      range.setEnd(endContainer, endOffset - 1);
    } else break;
  }
  return range;
};

const wrapRangeWithHighlight = (range, color, id) => {
  // 先把兩端空白縮掉，否則尾端空白會讓背景／下線多出來一截。
  trimRangeWhitespace(range);
  if (range.collapsed) return null;
  const highlightMode =
    HkHighlightDom?.getHighlightMode?.(document, range) ??
    (HkHighlightDom?.shouldUseTextNodeWrapping?.(document, range)
      ? "css-highlight"
      : "inline");
  if (highlightMode === "css-highlight") {
    return createTranslatedHighlightOverlay(range, color, id);
  }
  // If range start and end are in the SAME block, use extractContents — one mark,
  // handles <em>/<strong>/injected spans correctly without breaking anything.
  const startBlock = nearestBlockAncestor(range.startContainer);
  const endBlock = nearestBlockAncestor(range.endContainer);

  const useTextNodeWrapping =
    HkHighlightDom?.shouldUseTextNodeWrapping?.(document, range) ?? false;

  if (startBlock === endBlock && !useTextNodeWrapping) {
    try {
      const mark = createHighlightElement(color, id);
      const extracted = range.extractContents();
      mark.appendChild(extracted);
      range.insertNode(mark);
      return mark;
    } catch (_e) {
      // fall through to text-node walking
    }
  }

  // Range crosses block boundaries (e.g. <li> → <li>) — text-node walking
  // preserves block structure, creates multiple marks sharing the same ID.
  return wrapRangeTextNodeWalking(range, color, id);
};

const getStoredHighlights = async (key = pageKey) => {
  if (!storage) return [];
  const existing = await storage.get(key);
  return existing[key] ?? [];
};

// 輕量 toast：固定在畫面底部、自動消失。用於把「存檔失敗」這類過去被
// console.debug 默默吞掉的錯誤，明確告訴使用者。
let hkToastTimer = null;
const showHkToast = (message, { isError = false, duration = 4000 } = {}) => {
  try {
    let toast = document.getElementById("hk-toast");
    if (!toast) {
      toast = document.createElement("div");
      toast.id = "hk-toast";
      toast.className = "hk-toast";
      toast.setAttribute("role", "status");
      toast.setAttribute("aria-live", "polite");
      (document.body || document.documentElement).appendChild(toast);
    }
    toast.textContent = message;
    toast.classList.toggle("is-error", isError);
    toast.classList.add("is-visible");
    if (hkToastTimer) window.clearTimeout(hkToastTimer);
    hkToastTimer = window.setTimeout(() => {
      toast.classList.remove("is-visible");
    }, duration);
  } catch (_e) {
    /* 連 toast 都失敗就算了，不要再丟錯打斷流程 */
  }
};

const isQuotaError = (error) => {
  const msg = (error?.message || chrome.runtime?.lastError?.message || "").toLowerCase();
  return msg.includes("quota") || msg.includes("exceeded");
};

const setStoredHighlights = async (highlights, key = pageKey) => {
  if (!storage) return;
  colorUsageCache = null; // 標註有變動 → 讓色彩統計快取失效
  try {
    await storage.set({ [key]: highlights });
  } catch (error) {
    showHkToast(
      isQuotaError(error)
        ? t("storage.quotaToast")
        : t("storage.saveFailed"),
      { isError: true }
    );
    throw error;
  }
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
        setAllMarksMetadata(highlight.id, {
          color: highlight.color,
          note: highlight.note ?? "",
          kind: highlight.kind,
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

// 同一篇文章的識別碼：當 pageKey 沒有查詢字串（canonical 乾淨網址）時，
// 用「origin + 路徑」比對，才能把帶不同信件／追蹤參數的舊鍵也認出來；
// 若 pageKey 仍帶查詢字串（查詢決定內容的站，如 youtube ?v=），則嚴格比對正規化網址。
// SPA 導航時 pageKey 會被重算，所以這裡用函式即時讀「當前」pageKey，
// 避免初始注入時固定下來的舊判斷套用到新頁。
const pageKeyHasQuery = () => {
  try {
    return Boolean(new URL(pageKey).search);
  } catch (_e) {
    return false;
  }
};
const articleIdentity = (url) => {
  try {
    const u = new URL(url);
    return pageKeyHasQuery()
      ? normalizePageKey(url)
      : `${u.origin}${u.pathname}`;
  } catch (_e) {
    return null;
  }
};

// 把舊版用「完整網址（含追蹤／信件參數）」存的資料，搬到正規化後的 pageKey。
// 掃描整個 storage，找出與本文同一識別碼的舊鍵並合併；搬完刪掉舊鍵。一次性、冪等。
const LEGACY_MIGRATED_FLAG = "hkLegacyMigrated";
let legacyMigrationPromise = null;
const migrateLegacyPageData = () => {
  if (legacyMigrationPromise) return legacyMigrationPromise;
  legacyMigrationPromise = (async () => {
    if (!storage) return;
    try {
      const targetId = articleIdentity(pageKey);
      if (!targetId) return;
      // 便宜判斷，避免每頁載入都 get(null) 全量掃描：
      // ① 全域已標記「無殘留舊式網址鍵」→ 直接 skip
      // ② 本頁已有標註 → 之前造訪時就已合併過任何舊鍵，無需再掃
      const pre = await storage.get([pageKey, LEGACY_MIGRATED_FLAG]);
      if (pre[LEGACY_MIGRATED_FLAG]) return;
      if (Array.isArray(pre[pageKey]) && pre[pageKey].length) return;

      const all = await storage.get(null);
      const writes = {};
      const removes = [];

      // 1) 標註：頂層以 http(s) 網址為鍵、值為陣列。把同一識別碼的舊鍵併入 pageKey。
      const merged = Array.isArray(all[pageKey]) ? [...all[pageKey]] : [];
      const seenIds = new Set(merged.map((it) => it?.id).filter(Boolean));
      let touchedHighlights = false;
      for (const key of Object.keys(all)) {
        if (key === pageKey || !/^https?:/i.test(key)) continue;
        if (!Array.isArray(all[key])) continue;
        if (articleIdentity(key) !== targetId) continue;
        for (const item of all[key]) {
          const normalized =
            item && typeof item === "object" ? { ...item, url: pageKey } : item;
          if (normalized?.id && seenIds.has(normalized.id)) continue;
          if (normalized?.id) seenIds.add(normalized.id);
          merged.push(normalized);
        }
        removes.push(key);
        touchedHighlights = true;
      }
      if (touchedHighlights && merged.length) writes[pageKey] = merged;

      // 2) 心智圖、3) 生成筆記、4) 頁面 meta（皆為 { [url]: value } 物件）
      const moveSubKey = (storeKey) => {
        const map = all[storeKey];
        if (!map || typeof map !== "object") return;
        let changed = false;
        const next = { ...map };
        for (const key of Object.keys(map)) {
          if (key === pageKey || !/^https?:/i.test(key)) continue;
          if (articleIdentity(key) !== targetId) continue;
          if (next[pageKey] === undefined) next[pageKey] = map[key];
          delete next[key];
          changed = true;
        }
        if (changed) writes[storeKey] = next;
      };
      moveSubKey(MINDMAP_STORAGE_KEY);
      moveSubKey("hkGeneratedNotes");
      moveSubKey(PAGE_META_KEY);

      if (Object.keys(writes).length) await storage.set(writes);
      if (removes.length) await storage.remove(removes);

      // 若已無「帶追蹤參數／錨點」的舊式網址鍵，標記完成；之後不再全量掃描。
      const remaining = { ...all };
      removes.forEach((k) => delete remaining[k]);
      const hasDirty = Object.keys(remaining).some(
        (k) => /^https?:/i.test(k) && k !== normalizePageKey(k)
      );
      if (!hasDirty) await storage.set({ [LEGACY_MIGRATED_FLAG]: true });
    } catch (error) {
      console.debug("舊版標註資料搬移失敗", error);
    }
  })();
  return legacyMigrationPromise;
};

// 重整後找不到對應文字（頁面改版／內容被移除）的標註數，給面板與提示用。
let orphanHighlightCount = 0;
let orphanNotified = false;
const notifyOrphanHighlights = (orphanCount, total) => {
  orphanHighlightCount = orphanCount;
  updateOrphanNotice();
  if (orphanNotified || orphanCount <= 0) return;
  orphanNotified = true;
  showHkToast(
    t("panel.orphanToast", { orphan: orphanCount, total }),
    { isError: true, duration: 6500 }
  );
};

const attemptRestoreHighlights = async (attempt = 0) => {
  if (attempt === 0) await migrateLegacyPageData();
  await ensurePageMetaTitle(pageKey, document.title);
  const { total, visible } = await restoreHighlights();
  checkPendingFocusHighlight();
  if (!total) {
    orphanHighlightCount = 0;
    updateOrphanNotice();
    return;
  }
  if (visible >= total) {
    // 全部還原成功（可能是這次或之前重試補上的）→ 清掉殘留的失聯狀態。
    orphanHighlightCount = 0;
    updateOrphanNotice();
    return;
  }
  if (attempt < HIGHLIGHT_RETRY_DELAYS.length) {
    // 還有重試額度：內容可能還在延遲載入，先別判定為失聯。
    const nextDelay = HIGHLIGHT_RETRY_DELAYS[attempt] ?? 1200;
    window.setTimeout(() => attemptRestoreHighlights(attempt + 1), nextDelay);
    return;
  }
  // 重試用盡仍有缺 → 視為失聯，提示使用者。
  notifyOrphanHighlights(total - visible, total);
};

const applyHighlight = async (color) => {
  const selection = window.getSelection();

  if (!selection || selection.rangeCount === 0) {
    throw new Error(t("highlight.errNoSelection"));
  }

  const range = selection.getRangeAt(0);

  if (selection.isCollapsed || range.collapsed) {
    throw new Error(t("highlight.errEmptySelection"));
  }

  if (isEditableElement(range.commonAncestorContainer)) {
    throw new Error(t("highlight.errEditable"));
  }

  const snapshot = serializeRange(range.cloneRange());
  const highlightId = `hk-${Date.now()}-${Math.floor(Math.random() * 100000)}`;

  const normalizedColor = toHexColor(color || DEFAULT_COLOR);
  const highlightEl = wrapRangeWithHighlight(range, normalizedColor, highlightId);
  setAllMarksMetadata(highlightId, { color: normalizedColor, note: "" });
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

const findTranslatedHighlightAtPoint = (event) => {
  const x = event.clientX;
  const y = event.clientY;
  if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
  for (const entry of translatedHighlightOverlays.values()) {
    const rects =
      entry.mode === "css"
        ? getTranslatedHighlightRects(entry.visualRanges || entry.range)
        : [getHighlightVisualRect(entry.group)].filter(Boolean);
    for (const rect of rects) {
      if (
        x >= rect.left &&
        x <= rect.right &&
        y >= rect.top &&
        y <= rect.bottom
      ) {
        return entry.group;
      }
    }
  }
  return null;
};

const handleHighlightClick = (event) => {
  const target = event.target;
  if (!(target instanceof Element)) return;
  const highlightEl =
    target.closest(`.${HIGHLIGHT_CLASS}`) ||
    findTranslatedHighlightAtPoint(event);
  if (!highlightEl) return;
  event.preventDefault();
  event.stopPropagation();
  hideFloatingButton();
  openHighlightMenu(highlightEl).catch((error) =>
    console.debug("開啟 highlight 面板失敗", error)
  );
};

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type === "BUILD_READING_GUIDE_PROMPT") {
    (async () => {
      try {
        const requestPageKey = pageKey;
        const pageData = await collectGuidePageData();
        const palette = await refreshPaletteFromStorage();
        const existingTags = await collectExistingTagsByUsage();
        if (requestPageKey !== pageKey) throw new Error(t("sidebar.extractError"));
        sendResponse({
          success: true,
          prompt: buildReadingGuidePrompt(pageData, palette, existingTags),
          pageKey: requestPageKey,
          provider: aiSettings.provider,
          configured: ["chatgpt", "gemini-web"].includes(aiSettings.provider) || Boolean(
            isLikelyProviderApiKey(
              aiSettings.provider,
              aiSettings.provider === "openai" ? aiSettings.openaiKey : aiSettings.geminiKey
            )
          ),
        });
      } catch (error) {
        sendResponse({ success: false, error: error?.message || t("sidebar.extractError") });
      }
    })();
    return true;
  }
  if (message?.type === "GENERATE_READING_GUIDE") {
    (async () => {
      try {
        if (message.pageKey && message.pageKey !== pageKey) throw new Error(t("sidebar.sourceMismatch"));
        const provider = aiSettings.provider;
        if (["chatgpt", "gemini-web"].includes(provider)) {
          sendResponse({ success: false, manual: true, provider, error: t("sidebar.manualPaste") });
          return;
        }
        const apiKey = provider === "openai" ? aiSettings.openaiKey?.trim() : aiSettings.geminiKey?.trim();
        if (!isLikelyProviderApiKey(provider, apiKey)) throw new Error(t("ai.errNoApiKey"));
        const requestPageKey = pageKey;
        const pageData = await collectGuidePageData();
        const palette = await refreshPaletteFromStorage();
        const existingTags = await collectExistingTagsByUsage();
        if (requestPageKey !== pageKey) throw new Error(t("sidebar.parseError"));
        const prompt = buildReadingGuidePrompt(pageData, palette, existingTags);
        const model = provider === "openai" ? aiSettings.openaiModel : aiSettings.geminiModel;
        const response = provider === "openai"
          ? await callOpenAI(apiKey, prompt)
          : await callGemini(apiKey, prompt);
        if (requestPageKey !== pageKey) throw new Error(t("sidebar.parseError"));
        sendResponse(await applyReadingGuideResponse(response, prompt, provider, model, requestPageKey));
      } catch (error) {
        console.debug("產生閱讀導覽失敗", error);
        sendResponse({ success: false, error: error?.message || t("sidebar.generateError") });
      }
    })();
    return true;
  }
  if (message?.type === "APPLY_READING_GUIDE_RESPONSE") {
    (async () => {
      try {
        if (message.pageKey && message.pageKey !== pageKey) throw new Error(t("sidebar.sourceMismatch"));
        const text = typeof message.text === "string" ? message.text.trim() : "";
        if (!text) throw new Error(t("sidebar.parseError"));
        const shared = HkParsers.parseSharedMarkdown(HkParsers.unescapeAiMarkdown(text));
        if (shared) {
          sendResponse(await importSharedMarkdown(shared));
          return;
        }
        const requestPageKey = pageKey;
        const provider = message.provider === "gemini-web" ? "gemini-web" : "chatgpt";
        const model = provider === "gemini-web" ? "gemini-web" : "chatgpt-web";
        const result = await applyReadingGuideResponse(text, message.prompt || "", provider, model, requestPageKey);
        sendResponse(result);
      } catch (error) {
        sendResponse({ success: false, error: error?.message || t("sidebar.parseError") });
      }
    })();
    return true;
  }
  if (message?.type === "FOCUS_SAVED_HIGHLIGHT") {
    (async () => {
      try {
        const entries = await getStoredHighlights();
        const targetText = HkParsers.looseKey(message.text || "");
        const entry = entries.find((item) => {
          if (HkParsers.looseKey(item?.text || "") !== targetText) return false;
          if (message.kind === "term") return item.kind === "term";
          return !message.kind || item.kind === "passage" || !item.kind;
        });
        const preview = previewData.find((item) =>
          HkParsers.looseKey(item.data?.text || "") === targetText &&
          (!message.kind || item.data?.kind === message.kind)
        );
        const targetEntry = entry || preview?.data;
        if (!targetEntry) {
          // 名詞落在重點句裡時沒有自己的標記，改跳到包含它的那個標記。
          const container = [...document.querySelectorAll(`.${HIGHLIGHT_CLASS}`)].find((el) =>
            HkParsers.looseKey(el.textContent || "").includes(targetText)
          );
          if (!container || !targetText) throw new Error(t("sidebar.noMatch"));
          focusHighlightElement(container.getAttribute(HIGHLIGHT_ATTR));
          sendResponse({ success: true });
          return;
        }
        const target = document.querySelector(`[${HIGHLIGHT_ATTR}="${targetEntry.id}"]`);
        if (!target) throw new Error(t("sidebar.noMatch"));
        focusHighlightElement(targetEntry.id);
        sendResponse({ success: true });
      } catch (error) {
        sendResponse({ success: false, error: error?.message || t("sidebar.noMatch") });
      }
    })();
    return true;
  }
  if (message?.type === "GET_PAGE_HIGHLIGHTS") {
    (async () => {
      try {
        const highlights = await collectPageHighlights();
        const payload = {
          title: document.title,
          url: pageKey,
          pageText: getPagePlainText(),
          // 存了但頁面上找不到標記的，讓側欄標示出來，不要看起來一切正常。
          highlights: highlights.map((item) => ({
            ...item,
            notOnPage: !document.querySelector(`[${HIGHLIGHT_ATTR}="${item.id}"]`),
          })),
          previewHighlights: previewData.map((item) => item.data),
        };
        sendResponse({ success: true, data: payload });
      } catch (error) {
        console.debug("取得標註資料失敗", error);
        sendResponse({ success: false, error: error?.message || "無法取得標註資料" });
      }
    })();
    return true;
  }
  if (message?.type === "OPEN_AI_SETTINGS") {
    openHighlightPanel({ view: "settings" });
    sendResponse({ success: true });
    return false;
  }
  if (message?.type === "DELETE_GUIDE_ITEM" || message?.type === "DELETE_PAGE_NOTES") {
    (async () => {
      try {
        if (message.pageKey && message.pageKey !== pageKey) throw new Error(t("sidebar.sourceMismatch"));
        if (message.type === "DELETE_GUIDE_ITEM") await deleteGuideItem(message);
        else await deletePageNotes();
        sendResponse({ success: true });
      } catch (error) {
        console.debug("刪除筆記失敗", error);
        sendResponse({ success: false, error: error?.message });
      }
    })();
    return true;
  }
  if (message?.type === "START_GUIDED_READING") {
    const hasSteps = collectGuidedSteps().length > 0;
    if (hasSteps) startGuidedReading();
    sendResponse({ success: hasSteps });
    return false;
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
    // 快捷鍵不帶顏色 → 用目前選的顏色
    const color = message.color || currentColor || DEFAULT_COLOR;
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

// ── SPA 導航偵測 ─────────────────────────────────────────────
// 內容腳本只在初次注入時跑一次，SPA（Substack、Medium、新聞站）用
// history API 換頁不會重新注入，於是 pageKey 卡在舊頁、新標註存錯鍵、
// 舊頁標註也殘留在新畫面上。這裡監看網址變化，換頁時清舊標、重算
// pageKey 並重新還原。
const LOCATION_CHANGE_EVENT = "hk:locationchange";
let urlChangeTimer = null;

const onUrlChanged = (next) => {
  // 清掉舊頁殘留的 in-page 標註，避免疊到新畫面（沿用 clearPageHighlights 的做法）。
  try {
    document
      .querySelectorAll(`.${HIGHLIGHT_CLASS}`)
      .forEach(unwrapHighlightElement);
  } catch (_e) {}

  pageKey = next;
  highlightPanelState.activeKey = pageKey;

  // 重設失聯狀態：新頁要重新統計，不可沿用舊頁的數字／已通知旗標。
  orphanHighlightCount = 0;
  orphanNotified = false;
  try {
    updateOrphanNotice();
  } catch (_e) {}

  // 重新還原；restoreHighlights 對已存在的標記是冪等的，重試階梯會處理
  // SPA 延遲渲染。attemptRestoreHighlights 內已會 ensurePageMetaTitle。
  attemptRestoreHighlights(0);

  // 面板若開著，讓它跟著切到新頁。
  try {
    refreshHighlightPanelIfVisible();
  } catch (_e) {}
};

const handleLocationChange = () => {
  let next;
  try {
    next = normalizePageKey(window.location.href);
  } catch (_e) {
    return;
  }
  if (next === pageKey) return;
  // debounce：SPA 換頁常連續觸發多個事件／mutation，合併成一次。
  if (urlChangeTimer) window.clearTimeout(urlChangeTimer);
  urlChangeTimer = window.setTimeout(() => {
    urlChangeTimer = null;
    let settled;
    try {
      settled = normalizePageKey(window.location.href);
    } catch (_e) {
      return;
    }
    if (settled === pageKey) return;
    onUrlChanged(settled);
  }, 300);
};

// 攔截 pushState／replaceState：原生不會發事件，補一個合成事件好讓我們監聽。
// 用旗標防止重複注入（如 popup 的 executeScript 後備路徑）時二度包裝／重複註冊。
if (!window.__hkUrlWatcherInstalled) {
  window.__hkUrlWatcherInstalled = true;
  try {
    const wrapHistoryMethod = (name) => {
      const original = history[name];
      if (typeof original !== "function" || original.__hkWrapped) return;
      const wrapped = function (...args) {
        const result = original.apply(this, args);
        try {
          window.dispatchEvent(new Event(LOCATION_CHANGE_EVENT));
        } catch (_e) {}
        return result;
      };
      wrapped.__hkWrapped = true;
      history[name] = wrapped;
    };
    wrapHistoryMethod("pushState");
    wrapHistoryMethod("replaceState");
  } catch (_e) {}

  window.addEventListener("popstate", handleLocationChange);
  window.addEventListener("hashchange", handleLocationChange);
  window.addEventListener(LOCATION_CHANGE_EVENT, handleLocationChange);
}

// ── 分享連結收件端 ─────────────────────────────────────────────
// 兩種來源，偵測到就跳頂部橫幅、一鍵匯入到自己的擴充：
// 1. 原文網址帶 #hk=<壓縮筆記>（fragment 分享，零設定）
// 2. raw.githubusercontent.com 上的 highlight-keeper-bulk JSON（GitHub 分享）
const SHARE_FRAGMENT_PREFIX = "#hk=";

const decompressShareToken = async (token) => {
  const b64 = token.replace(/-/g, "+").replace(/_/g, "/");
  const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
  const stream = new Blob([bytes])
    .stream()
    .pipeThrough(new DecompressionStream("deflate-raw"));
  return JSON.parse(await new Response(stream).text());
};

// 橫幅 UI（inline style，避開頁面 CSS 干擾）。回傳 {bar,msg,setActions,mkBtn}。
const createShareBanner = () => {
  const bar = document.createElement("div");
  bar.setAttribute(
    "style",
    [
      "position:fixed",
      "top:0",
      "left:0",
      "right:0",
      "z-index:2147483647",
      "display:flex",
      "align-items:center",
      "gap:12px",
      "flex-wrap:wrap",
      "padding:10px 16px",
      "background:#1c1917",
      "color:#f5f5f4",
      "font:14px/1.4 -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",
      "box-shadow:0 2px 10px rgba(0,0,0,.25)",
    ].join(";")
  );
  const msg = document.createElement("span");
  msg.style.flex = "1";
  msg.style.minWidth = "180px";
  const actions = document.createElement("span");
  actions.style.display = "flex";
  actions.style.gap = "8px";
  actions.style.flexWrap = "wrap";

  const mkBtn = (label, primary) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = label;
    btn.setAttribute(
      "style",
      [
        "font:inherit",
        "font-weight:600",
        "cursor:pointer",
        "border-radius:8px",
        "padding:7px 14px",
        "border:1px solid " + (primary ? "#f59e0b" : "#57534e"),
        "background:" + (primary ? "#f59e0b" : "transparent"),
        "color:" + (primary ? "#1c1917" : "#f5f5f4"),
      ].join(";")
    );
    return btn;
  };

  const closeBtn = mkBtn("✕", false);
  closeBtn.title = t("share.close");
  closeBtn.style.padding = "7px 10px";
  closeBtn.addEventListener("click", () => bar.remove());

  bar.appendChild(msg);
  bar.appendChild(actions);
  bar.appendChild(closeBtn);

  const setActions = (buttons) => {
    actions.innerHTML = "";
    buttons.forEach((btn) => actions.appendChild(btn));
  };
  return { bar, msg, setActions, mkBtn };
};

// fragment 分享：匯入到「當前頁面」（走 AI 套用管線，標註即套即見）。
const startFragmentShareImport = (payload) => {
  const { bar, msg, setActions, mkBtn } = createShareBanner();

  const stripHashFromUrl = () => {
    try {
      history.replaceState(null, "", location.pathname + location.search);
    } catch (_e) {}
  };

  const finish = async (overwrite) => {
    msg.textContent = t("share.importing");
    setActions([]);
    try {
      const res = await applySharedPayload(payload, { overwrite });
      await refreshHighlightPanelIfVisible();
      msg.textContent = t("share.done", {
        count: res.highlight?.appliedCount || 0,
      });
      stripHashFromUrl();
      window.setTimeout(() => bar.remove(), 6000);
    } catch (error) {
      console.debug("匯入分享筆記失敗", error);
      msg.textContent = error?.message || t("share.fail");
    }
  };

  const promptImport = async () => {
    let existing = [];
    try {
      const stored = await chrome.storage.local.get(pageKey);
      existing = Array.isArray(stored?.[pageKey]) ? stored[pageKey] : [];
    } catch (_e) {}
    if (!existing.length) return finish(false);
    // 本頁已有標註 → 併入／覆蓋／取消
    msg.textContent = t("share.conflict", { count: existing.length });
    const mergeBtn = mkBtn(t("share.merge"), true);
    const overwriteBtn = mkBtn(t("share.overwrite"), false);
    const cancelBtn = mkBtn(t("share.cancel"), false);
    mergeBtn.addEventListener("click", () => finish(false));
    overwriteBtn.addEventListener("click", () => finish(true));
    cancelBtn.addEventListener("click", () => {
      msg.textContent = t("share.cancelled");
      setActions([]);
      window.setTimeout(() => bar.remove(), 2500);
    });
    setActions([mergeBtn, overwriteBtn, cancelBtn]);
  };

  const count = Array.isArray(payload.hl) ? payload.hl.length : 0;
  msg.textContent = `🔖 ${t("share.detected", {
    title: payload.title || "",
    count,
  })}`;
  const importBtn = mkBtn(t("share.importBtn"), true);
  importBtn.addEventListener("click", () => promptImport());
  setActions([importBtn]);
  document.body.appendChild(bar);
};

// GitHub raw 分享：整頁 JSON（含原始 entries）直接寫進 storage，
// 開啟原文時就會還原。與 manager 匯入相同的資料形狀。
const startRawShareImport = (page) => {
  const { bar, msg, setActions, mkBtn } = createShareBanner();

  const entryKey = (entry) =>
    entry?.id
      ? `id:${entry.id}`
      : `t:${entry?.text || ""}|${entry?.note || ""}`;

  const importPage = async (mode) => {
    const url = page.url;
    const cur = await chrome.storage.local.get([
      url,
      PAGE_META_KEY,
      "hkGeneratedNotes",
      MINDMAP_STORAGE_KEY,
      "hkPersonalNotes",
    ]);
    const existingEntries = Array.isArray(cur[url]) ? cur[url] : [];
    let entries;
    if (mode === "merge") {
      const seen = new Set(existingEntries.map(entryKey));
      entries = existingEntries.slice();
      page.entries.forEach((entry) => {
        const key = entryKey(entry);
        if (!seen.has(key)) {
          seen.add(key);
          entries.push(entry);
        }
      });
    } else {
      entries = page.entries.slice();
    }

    const metaAll = { ...(cur[PAGE_META_KEY] || {}) };
    const metaEntry = { ...(metaAll[url] || {}) };
    if (page.title && (mode !== "merge" || !metaEntry.title)) {
      metaEntry.title = page.title;
    }
    const tagSet = new Set([
      ...(mode === "merge" && Array.isArray(metaEntry.tags) ? metaEntry.tags : []),
      ...(Array.isArray(page.tags) ? page.tags : []),
    ]);
    if (tagSet.size) metaEntry.tags = [...tagSet];
    metaAll[url] = metaEntry;

    const notesAll = { ...(cur.hkGeneratedNotes || {}) };
    if (page.note) {
      if (mode === "merge" && notesAll[url]) {
        notesAll[url] = {
          ...page.note,
          ...notesAll[url],
          ...(notesAll[url].context || !page.note.context ? {} : { context: page.note.context }),
        };
      } else {
        notesAll[url] = page.note;
      }
    }
    const mapsAll = { ...(cur[MINDMAP_STORAGE_KEY] || {}) };
    if (page.mindmap && (mode !== "merge" || !mapsAll[url])) {
      mapsAll[url] = page.mindmap;
    }
    const personalNotesAll = { ...(cur.hkPersonalNotes || {}) };
    if (
      typeof page.personalNote === "string" &&
      (mode !== "merge" || !personalNotesAll[url]?.text)
    ) {
      personalNotesAll[url] = { text: page.personalNote, updatedAt: Date.now() };
    }

    await chrome.storage.local.set({
      [url]: entries,
      [PAGE_META_KEY]: metaAll,
      hkGeneratedNotes: notesAll,
      [MINDMAP_STORAGE_KEY]: mapsAll,
      hkPersonalNotes: personalNotesAll,
    });
    return entries.length;
  };

  const finish = async (mode) => {
    msg.textContent = t("share.importing");
    setActions([]);
    try {
      const total = await importPage(mode);
      msg.textContent = "";
      msg.appendChild(
        document.createTextNode(t("share.rawDone", { count: total }) + " ")
      );
      const openLink = document.createElement("a");
      openLink.href = page.url;
      openLink.textContent = t("share.openSource");
      openLink.setAttribute("style", "color:#fbbf24;font-weight:600");
      msg.appendChild(openLink);
    } catch (error) {
      console.debug("匯入分享筆記失敗", error);
      msg.textContent = error?.message || t("share.fail");
    }
  };

  const promptImport = async () => {
    let existing = [];
    try {
      const stored = await chrome.storage.local.get(page.url);
      existing = Array.isArray(stored?.[page.url]) ? stored[page.url] : [];
    } catch (_e) {}
    if (!existing.length) return finish("overwrite");
    msg.textContent = t("share.conflict", { count: existing.length });
    const mergeBtn = mkBtn(t("share.merge"), true);
    const overwriteBtn = mkBtn(t("share.overwrite"), false);
    const cancelBtn = mkBtn(t("share.cancel"), false);
    mergeBtn.addEventListener("click", () => finish("merge"));
    overwriteBtn.addEventListener("click", () => finish("overwrite"));
    cancelBtn.addEventListener("click", () => {
      msg.textContent = t("share.cancelled");
      setActions([]);
    });
    setActions([mergeBtn, overwriteBtn, cancelBtn]);
  };

  const count = Array.isArray(page.entries) ? page.entries.length : 0;
  msg.textContent = `🔖 ${t("share.detected", {
    title: page.title || "",
    count,
  })}`;
  const importBtn = mkBtn(t("share.importBtn"), true);
  importBtn.addEventListener("click", () => promptImport());
  setActions([importBtn]);
  document.body.appendChild(bar);
};

// 進入點：判斷當前頁是不是分享連結（fragment 或 raw JSON）。
const detectSharedNotes = async () => {
  try {
    // 1) fragment：原文網址#hk=…
    const hash = location.hash || "";
    if (hash.startsWith(SHARE_FRAGMENT_PREFIX)) {
      const payload = await decompressShareToken(
        hash.slice(SHARE_FRAGMENT_PREFIX.length)
      );
      if (
        payload?.v === 1 &&
        (payload.hl?.length || payload.note || payload.context || payload.personalNote || payload.mindmap)
      ) {
        startFragmentShareImport(payload);
      }
      return;
    }
    // 2) raw.githubusercontent 上的備份／分享 JSON
    if (location.hostname === "raw.githubusercontent.com") {
      const text = document.body?.innerText || "";
      if (text.length > 5 * 1024 * 1024) return;
      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch (_e) {
        return;
      }
      if (parsed?.type !== "highlight-keeper-bulk") return;
      const pages = HkParsers.normalizeBulkPages(parsed.pages);
      if (pages.length === 1) startRawShareImport(pages[0]);
    }
  } catch (error) {
    console.debug("分享筆記偵測失敗", error);
  }
};

detectSharedNotes();

const ensureNoteTooltip = () => {
  if (highlightNoteTooltip) return highlightNoteTooltip;
  const el = document.createElement("div");
  el.id = HIGHLIGHT_NOTE_TOOLTIP_ID;
  el.className = "hk-note-tooltip";
  el.setAttribute("role", "tooltip");
  document.body.appendChild(el);
  highlightNoteTooltip = el;
  return el;
};

const showNoteTooltip = (text, targetEl) => {
  if (!text || !targetEl) return;
  const tooltip = ensureNoteTooltip();
  tooltip.textContent = HkParsers.breakNoteLines(text);

  // measure while invisible to avoid position flash
  tooltip.style.visibility = "hidden";
  tooltip.style.display = "block";

  const rect = targetEl.getBoundingClientRect();
  const scrollX = window.scrollX;
  const scrollY = window.scrollY;
  const margin = 12;

  const tooltipWidth = tooltip.offsetWidth;
  const tooltipHeight = tooltip.offsetHeight;
  const vpWidth = window.innerWidth;
  const vpHeight = window.innerHeight;

  let top = rect.bottom + scrollY + margin;
  let left = rect.left + scrollX + rect.width / 2 - tooltipWidth / 2;

  if (left < scrollX + margin) left = scrollX + margin;
  if (left + tooltipWidth > scrollX + vpWidth - margin) {
    left = scrollX + vpWidth - tooltipWidth - margin;
  }
  if (rect.bottom + margin + tooltipHeight > vpHeight) {
    top = rect.top + scrollY - margin - tooltipHeight;
  }

  tooltip.style.left = `${left}px`;
  tooltip.style.top = `${top}px`;
  tooltip.style.visibility = "visible";
};

const hideNoteTooltip = () => {
  if (!highlightNoteTooltip) return;
  highlightNoteTooltip.style.display = "none";
  highlightNoteTooltip.style.visibility = "hidden";
};

// ── Guided Reading ─────────────────────────────────────
let guidedReadingState = null;

const collectGuidedSteps = () => {
  const seen = new Set();
  const steps = [];
  document.querySelectorAll(`.${HIGHLIGHT_CLASS}`).forEach((el) => {
    const id = el.getAttribute(HIGHLIGHT_ATTR);
    if (!id || seen.has(id)) return;
    seen.add(id);
    steps.push({ id, el });
  });
  return steps;
};

const destroyGuidedBar = () => {
  const existing = document.getElementById("hk-guided-bar");
  if (existing) existing.remove();
  if (guidedReadingState?.focusedEl) {
    guidedReadingState.focusedEl.classList.remove("hk-guided-focus");
  }
  guidedReadingState = null;
};

const renderGuidedBar = () => {
  const state = guidedReadingState;
  if (!state) return;
  const { steps, index } = state;
  const step = steps[index];
  const el = step.el;

  if (state.focusedEl && state.focusedEl !== el) {
    state.focusedEl.classList.remove("hk-guided-focus");
  }
  el.classList.add("hk-guided-focus");
  state.focusedEl = el;

  el.scrollIntoView({ behavior: "smooth", block: "center" });

  const bar = document.getElementById("hk-guided-bar");
  if (!bar) return;

  const note = el.dataset.hkNote?.trim() ?? "";
  const color = el.style.backgroundColor || "#fff5b8";
  const text = el.textContent?.trim() ?? "";

  bar.querySelector(".hk-guided-dot").style.backgroundColor = color;
  bar.querySelector(".hk-guided-text").textContent = text.length > 80 ? text.slice(0, 80) + "…" : text;
  bar.querySelector(".hk-guided-note").textContent = HkParsers.breakNoteLines(note);
  bar.querySelector(".hk-guided-note").style.display = note ? "block" : "none";
  bar.querySelector(".hk-guided-counter").textContent = `${index + 1} / ${steps.length}`;
  bar.querySelector(".hk-guided-prev").disabled = index === 0;
  bar.querySelector(".hk-guided-next").disabled = index === steps.length - 1;
};

const startGuidedReading = () => {
  destroyGuidedBar();
  const steps = collectGuidedSteps();
  if (!steps.length) {
    setAiPanelStatus(t("ai.errNoHighlightGuided"), true);
    return;
  }

  guidedReadingState = { steps, index: 0, focusedEl: null };

  const bar = document.createElement("div");
  bar.id = "hk-guided-bar";
  bar.className = "hk-guided-bar";

  const topRow = document.createElement("div");
  topRow.className = "hk-guided-top";

  const dot = document.createElement("span");
  dot.className = "hk-guided-dot";

  const text = document.createElement("span");
  text.className = "hk-guided-text";

  const counter = document.createElement("span");
  counter.className = "hk-guided-counter";

  const exitBtn = document.createElement("button");
  exitBtn.type = "button";
  exitBtn.className = "hk-guided-exit";
  exitBtn.textContent = t("guided.exit");
  exitBtn.addEventListener("click", destroyGuidedBar);

  topRow.appendChild(dot);
  topRow.appendChild(text);
  topRow.appendChild(counter);
  topRow.appendChild(exitBtn);

  const note = document.createElement("div");
  note.className = "hk-guided-note";

  const navRow = document.createElement("div");
  navRow.className = "hk-guided-nav";

  const prevBtn = document.createElement("button");
  prevBtn.type = "button";
  prevBtn.className = "hk-guided-prev";
  prevBtn.textContent = t("guided.prev");
  prevBtn.addEventListener("click", () => {
    if (guidedReadingState.index > 0) {
      guidedReadingState.index--;
      renderGuidedBar();
    }
  });

  const nextBtn = document.createElement("button");
  nextBtn.type = "button";
  nextBtn.className = "hk-guided-next";
  nextBtn.textContent = t("guided.next");
  nextBtn.addEventListener("click", () => {
    if (guidedReadingState.index < guidedReadingState.steps.length - 1) {
      guidedReadingState.index++;
      renderGuidedBar();
    }
  });

  navRow.appendChild(prevBtn);
  navRow.appendChild(nextBtn);

  bar.appendChild(topRow);
  bar.appendChild(note);
  bar.appendChild(navRow);
  document.body.appendChild(bar);

  renderGuidedBar();
};

document.addEventListener("mouseover", (event) => {
  const target = event.target?.closest?.(`.${HIGHLIGHT_CLASS}`);
  if (!target) return;
  const note = target.dataset.hkNote?.trim();
  if (!note) return;
  clearTimeout(tooltipHideTimer);
  showNoteTooltip(note, target);
});

document.addEventListener("mouseout", (event) => {
  const from = event.target?.closest?.(`.${HIGHLIGHT_CLASS}`);
  if (!from) return;
  const to = event.relatedTarget?.closest?.(`.${HIGHLIGHT_CLASS}`);
  if (to === from) return;
  tooltipHideTimer = setTimeout(hideNoteTooltip, 60);
});

document.addEventListener("mouseup", handleSelectionIntent);
document.addEventListener("keyup", handleSelectionIntent);
document.addEventListener("selectionchange", handleSelectionIntent);
document.addEventListener("click", handleHighlightClick, true);
document.addEventListener(
  "mousedown",
  (event) => {
    const target = event.target;
    const inToolbar = floatingButton && floatingButton.contains(target);
    const inCard = floatingTranslateCard && floatingTranslateCard.contains(target);
    if (floatingButton && !inToolbar && !inCard) {
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
  (event) => {
    updateTranslatedHighlightOverlays();
    // Ignore internal scrolling inside the highlight menu (e.g. scrolling the
    // note textarea after pasting a long note), otherwise the menu would close
    // the moment the textarea overflows.
    const target = event.target;
    if (
      highlightMenu &&
      target instanceof Node &&
      highlightMenu.contains(target)
    ) {
      return;
    }
    hideFloatingButton();
    closeHighlightMenu();
    hideNoteTooltip();
  },
  true
);
window.addEventListener(
  "resize",
  () => {
    updateTranslatedHighlightOverlays();
    closeHighlightMenu();
    // No data changed on resize — just keep the floating panel inside the
    // viewport. (A full re-render here re-read all of storage and caused
    // visible jank while resizing.)
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
  if (event.key !== "Escape") return;
  // 由內而外關閉：翻譯卡 → 設定抽屜 → 面板，別一次全收掉。
  if (floatingTranslateCard && floatingTranslateCard.style.display !== "none") {
    hideTranslateCard();
    return;
  }
  if (highlightPanelEls?.drawer?.classList.contains("is-open")) {
    highlightPanelEls.toggleDrawer?.(false);
    return;
  }
  closeHighlightMenu();
  hideFloatingButton();
  closeHighlightPanel();
});
const collectAllPageHighlights = async () => {
  if (!storage) throw new Error(t("ai.errStorageUnavailable"));
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
      setArchiveStatus(t("archive.nothingToDownload"), true);
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
    setArchiveStatus(t("archive.downloaded"));
  } catch (error) {
    console.debug("下載全部筆記失敗", error);
    setArchiveStatus(error?.message || t("archive.downloadFailed"), true);
  }
};

const mergeBulkImportPayload = async (files) => {
  if (!files?.length) return;
  try {
    setArchiveStatus(t("archive.importing"));
    const fileTexts = await Promise.all(Array.from(files).map((file) => file.text()));
    const allEntries = fileTexts.flatMap((text) =>
      parseImportedHighlightsPayload(text)
    );
    const normalized = allEntries
      .map((entry, index) => normalizeImportedHighlightEntry(entry, index))
      .filter(Boolean);
    if (!normalized.length) {
      throw new Error(t("manager.errNoImport"));
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
      throw new Error(t("ai.allSkipped"));
    }
    await storage.set(updates);
    await Promise.all(
      normalized
        .filter((entry) => entry.title)
        .map((entry) => ensurePageMetaTitle(entry.url, entry.title))
    );
    setArchiveStatus(
      t("ai.bulkImported", { imported: Object.keys(updates).length, skipped: skipped.length })
    );
    await refreshHighlightPanelData();
    await renderHighlightPanel();
    attemptRestoreHighlights();
  } catch (error) {
    console.debug("匯入全部筆記失敗", error);
    setArchiveStatus(error?.message || t("archive.importFailed"), true);
  }
};

const handleBulkImportChange = (event) => {
  const files = event.target.files;
  mergeBulkImportPayload(files);
};
