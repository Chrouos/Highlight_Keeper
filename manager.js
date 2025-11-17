const PAGE_META_KEY = "__hk_page_meta__";
const DEFAULT_COLOR = "#ffeb3b";
const GITHUB_SETTINGS_KEY = "hkGithubSyncSettings";
const GITHUB_DEFAULT_SETTINGS = {
  token: "",
  repo: "",
  branch: "main",
  path: "backups/highlight-keeper.json",
};

const state = {
  pages: [],
  meta: {},
  notes: {},
  searchTerm: "",
};
let detailCurrentPageUrl = "";
let githubSettings = { ...GITHUB_DEFAULT_SETTINGS };

const statusEl = document.getElementById("managerStatus");
const listEl = document.getElementById("pageList");
const pageCountEl = document.getElementById("pageCount");
const downloadBtn = document.getElementById("downloadAllBtn");
const importInput = document.getElementById("bulkImportInput");
const closeBtn = document.getElementById("closeManagerBtn");
const searchInput = document.getElementById("managerSearch");
const githubTokenInput = document.getElementById("githubToken");
const githubRepoInput = document.getElementById("githubRepo");
const githubBranchInput = document.getElementById("githubBranch");
const githubPathInput = document.getElementById("githubPath");
const githubDownloadBtn = document.getElementById("githubDownloadBtn");
const githubUploadBtn = document.getElementById("githubUploadBtn");
const githubStatusEl = document.getElementById("githubSyncStatus");
const detailOverlayId = "hk-manager-detail";

const setStatus = (message, isError = false) => {
  if (!statusEl) return;
  statusEl.textContent = message || "";
  statusEl.classList.toggle("is-error", Boolean(isError));
};

const setGithubStatus = (message, isError = false) => {
  if (!githubStatusEl) return;
  githubStatusEl.textContent = message || "";
  githubStatusEl.classList.toggle("is-error", Boolean(isError));
};

const setGithubActionsDisabled = (disabled) => {
  if (githubUploadBtn) {
    githubUploadBtn.disabled = disabled;
  }
  if (githubDownloadBtn) {
    githubDownloadBtn.disabled = disabled;
  }
};

const getPageDisplayName = (url) => {
  try {
    const parsed = new URL(url);
    const path = parsed.pathname && parsed.pathname !== "/"
      ? parsed.pathname.replace(/\/$/, "")
      : "/";
    return `${parsed.hostname}${path}`;
  } catch (_error) {
    return url;
  }
};

const isValidPageKey = (key) => {
  if (typeof key !== "string") return false;
  try {
    const url = new URL(key);
    return url.protocol === "http:" || url.protocol === "https:";
  } catch (_error) {
    return false;
  }
};

const toHexColor = (value) => {
  if (!value) return DEFAULT_COLOR;
  if (value.startsWith("#")) return value.toLowerCase();
  const match = value.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (!match) return DEFAULT_COLOR;
  const [, r, g, b] = match;
  const toHex = (component) =>
    Number(component).toString(16).padStart(2, "0").toLowerCase();
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
};

const parseTags = (input) => {
  if (typeof input !== "string") return [];
  return Array.from(
    new Set(
      input
        .split(/[\s,]+/)
        .map((tag) => tag.trim())
        .filter(Boolean)
    )
  );
};

const applyGithubSettingsToFields = () => {
  if (githubTokenInput) {
    githubTokenInput.value = githubSettings.token || "";
  }
  if (githubRepoInput) {
    githubRepoInput.value = githubSettings.repo || "";
  }
  if (githubBranchInput) {
    githubBranchInput.value = githubSettings.branch || "";
  }
  if (githubPathInput) {
    githubPathInput.value = githubSettings.path || "";
  }
};

const persistGithubSettings = (updates = {}) => {
  const normalized = { ...updates };
  if (typeof normalized.repo === "string") {
    normalized.repo = normalized.repo.trim();
  }
  if (typeof normalized.branch === "string") {
    normalized.branch = normalized.branch.trim() || GITHUB_DEFAULT_SETTINGS.branch;
  }
  if (typeof normalized.path === "string") {
    normalized.path = normalized.path.replace(/^\/+/, "");
  }
  githubSettings = {
    ...githubSettings,
    ...normalized,
  };
  return chrome.storage.local
    .set({ [GITHUB_SETTINGS_KEY]: githubSettings })
    .catch((error) => {
      console.debug("儲存 GitHub 設定失敗", error);
    });
};

const loadGithubSettings = async () => {
  try {
    const stored = await chrome.storage.local.get(GITHUB_SETTINGS_KEY);
    const saved = stored?.[GITHUB_SETTINGS_KEY] || {};
    githubSettings = {
      ...GITHUB_DEFAULT_SETTINGS,
      ...saved,
    };
    githubSettings = {
      ...githubSettings,
      token: githubSettings.token || "",
      repo: githubSettings.repo?.trim() || "",
      branch: githubSettings.branch?.trim() || GITHUB_DEFAULT_SETTINGS.branch,
      path: (githubSettings.path || GITHUB_DEFAULT_SETTINGS.path).replace(
        /^\/+/,
        ""
      ),
    };
  } catch (error) {
    console.debug("讀取 GitHub 設定失敗", error);
    githubSettings = { ...GITHUB_DEFAULT_SETTINGS };
  } finally {
    applyGithubSettingsToFields();
  }
};

const bindGithubInput = (element, key) => {
  if (!element) return;
  element.addEventListener("input", (event) => {
    let value = event.target.value;
    if (key !== "token") {
      value = value.trim();
    }
    if (key === "path") {
      value = value.replace(/^\/+/, "");
      if (value !== event.target.value) {
        event.target.value = value;
      }
    } else if (key !== "token" && value !== event.target.value) {
      event.target.value = value;
    }
    persistGithubSettings({ [key]: value });
  });
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
  const url = typeof entry.url === "string"
    ? entry.url
    : typeof entry.pageUrl === "string"
    ? entry.pageUrl
    : null;
  if (!url) return null;
  const range = entry.range;
  if (!range || typeof range.startXPath !== "string" || typeof range.endXPath !== "string") {
    return null;
  }
  const color = toHexColor(entry.color || DEFAULT_COLOR);
  return {
    id: `hk-import-${Date.now()}-${Math.floor(Math.random() * 100000)}-${index}`,
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
};

const fetchAllPages = async () => {
  const all = await chrome.storage.local.get(null);
  const meta = all[PAGE_META_KEY] || {};
  const notes = all.hkGeneratedNotes || {};
  const pages = Object.entries(all)
    .filter(([key, value]) => isValidPageKey(key) && Array.isArray(value))
    .map(([url, entries]) => {
      const latest = entries.reduce(
        (max, entry) => Math.max(max, Number(entry?.createdAt) || 0),
        0
      );
      return {
        url,
        title: meta[url]?.title?.trim() || getPageDisplayName(url),
        total: entries.length,
        updatedAt: latest,
        tags: Array.isArray(meta[url]?.tags) ? meta[url].tags : [],
        entries,
      };
    })
    .sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  state.pages = pages;
  state.meta = meta;
  state.notes = notes;
};

const matchesSearch = (page, term) => {
  if (!term) return true;
  const normalized = term.trim().toLowerCase();
  if (!normalized) return true;
  const haystacks = [
    page.title,
    page.url,
    ...(Array.isArray(page.tags) ? page.tags : []),
    ...(page.entries || []).flatMap((entry) => [
      entry.text,
      entry.note,
      ...(Array.isArray(entry.tags) ? entry.tags : []),
    ]),
  ]
    .filter(Boolean)
    .map((value) => String(value).toLowerCase());
  return haystacks.some((text) => text.includes(normalized));
};

const renderPageList = () => {
  if (!listEl || !pageCountEl) return;
  listEl.innerHTML = "";
  const filtered = state.pages.filter((page) =>
    matchesSearch(page, state.searchTerm)
  );
  pageCountEl.textContent = `共 ${filtered.length} 個頁面`;
  if (!filtered.length) {
    const empty = document.createElement("p");
    empty.textContent = state.searchTerm
      ? "沒有符合搜尋條件的頁面。"
      : "目前尚未建立任何筆記。";
    empty.className = "hk-manager-meta";
    listEl.appendChild(empty);
    return;
  }
  filtered.forEach((page) => {
    const card = document.createElement("article");
    card.className = "hk-manager-card";
    card.addEventListener("click", () => openPageDetail(page));
    const title = document.createElement("h3");
    title.textContent = page.title;
    card.appendChild(title);
    const urlLink = document.createElement("a");
    urlLink.href = page.url;
    urlLink.target = "_blank";
    urlLink.rel = "noopener";
    urlLink.className = "hk-manager-url";
    urlLink.textContent = page.url;
    card.appendChild(urlLink);
    const meta = document.createElement("div");
    meta.className = "hk-manager-meta";
    const updatedText = page.updatedAt
      ? new Intl.DateTimeFormat(undefined, {
          dateStyle: "medium",
          timeStyle: "short",
        }).format(new Date(page.updatedAt))
      : "未知時間";
    meta.textContent = `筆記數：${page.total} · 最後更新：${updatedText}`;
    card.appendChild(meta);
    listEl.appendChild(card);
  });
};

const refreshManager = async () => {
  await fetchAllPages();
  renderPageList();
};

const buildFullExportPayload = () => ({
  type: "highlight-keeper-bulk",
  version: 1,
  exportedAt: Date.now(),
  pages: state.pages.map((page) => ({
    url: page.url,
    title: state.meta[page.url]?.title || "",
    entries: page.entries,
  })),
});

const parseGithubBackupPayload = (rawText) => {
  let parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (_error) {
    throw new Error("GitHub 備份檔案不是有效的 JSON 格式");
  }
  const pages = Array.isArray(parsed?.pages)
    ? parsed.pages
    : Array.isArray(parsed)
    ? parsed
    : [];
  if (!pages.length) {
    throw new Error("GitHub 備份中沒有頁面資料");
  }
  const normalized = pages
    .map((page) => {
      if (!page || typeof page !== "object") return null;
      const url =
        typeof page.url === "string"
          ? page.url
          : typeof page.pageUrl === "string"
          ? page.pageUrl
          : null;
      if (!url || !isValidPageKey(url)) return null;
      const entries = Array.isArray(page.entries) ? page.entries : [];
      if (!entries.length) return null;
      return {
        url,
        title: typeof page.title === "string" ? page.title : "",
        entries,
      };
    })
    .filter(Boolean);
  if (!normalized.length) {
    throw new Error("GitHub 備份裡沒有可匯入的筆記");
  }
  return normalized;
};

const downloadAllPages = async () => {
  await fetchAllPages();
  if (!state.pages.length) {
    setStatus("沒有筆記可匯出", true);
    return;
  }
  const payload = buildFullExportPayload();
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
  setStatus("已下載全部筆記");
};

const ensureDetailOverlay = () => {
  let overlay = document.getElementById(detailOverlayId);
  if (overlay) return overlay;

  overlay = document.createElement("div");
  overlay.id = detailOverlayId;
  overlay.className = "hk-manager-detail is-hidden";

  const backdrop = document.createElement("div");
  backdrop.className = "hk-manager-detail-backdrop";
  backdrop.addEventListener("click", () => closePageDetail());

  const dialog = document.createElement("div");
  dialog.className = "hk-manager-detail-dialog";

  const header = document.createElement("header");
  header.className = "hk-manager-detail-header";

  const headingWrap = document.createElement("div");
  const titleEl = document.createElement("h3");
  titleEl.id = "hk-manager-detail-title";
  const urlEl = document.createElement("a");
  urlEl.id = "hk-manager-detail-url";
  urlEl.className = "hk-manager-detail-url";
  urlEl.target = "_blank";
  urlEl.rel = "noopener";
  const metaEl = document.createElement("div");
  metaEl.id = "hk-manager-detail-meta";
  metaEl.className = "hk-manager-detail-meta";
  headingWrap.appendChild(titleEl);
  headingWrap.appendChild(urlEl);
  headingWrap.appendChild(metaEl);

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "hk-manager-detail-close";
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => closePageDetail());

  header.appendChild(headingWrap);
  header.appendChild(closeBtn);

  const body = document.createElement("div");
  body.className = "hk-manager-detail-body";

  const entriesSection = document.createElement("section");
  entriesSection.className = "hk-manager-detail-section hk-manager-detail-section-notes";
  const entriesTitle = document.createElement("h4");
  entriesTitle.textContent = "筆記";
  const tagsRow = document.createElement("div");
  tagsRow.className = "hk-manager-detail-tags";
  const tagsLabel = document.createElement("span");
  tagsLabel.textContent = "頁面標籤";
  const tagsChips = document.createElement("div");
  tagsChips.id = "hk-manager-detail-tags-chips";
  tagsChips.className = "hk-manager-tags-chips";
  const tagsInput = document.createElement("input");
  tagsInput.id = "hk-manager-detail-tags-input";
  tagsInput.className = "hk-manager-input";
  tagsInput.placeholder = "以逗號或空白分隔多個 Tags";
  const tagsButton = document.createElement("button");
  tagsButton.type = "button";
  tagsButton.className = "hk-manager-btn hk-manager-btn-muted";
  tagsButton.textContent = "套用標籤";
  tagsButton.addEventListener("click", () => savePageTagsFromDetail());
  tagsInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      savePageTagsFromDetail();
    }
  });
  tagsRow.appendChild(tagsLabel);
  tagsRow.appendChild(tagsChips);
  tagsRow.appendChild(tagsInput);
  tagsRow.appendChild(tagsButton);
  const entriesList = document.createElement("div");
  entriesList.id = "hk-manager-detail-entries";
  entriesList.className = "hk-manager-detail-list";
  entriesSection.appendChild(entriesTitle);
  entriesSection.appendChild(tagsRow);
  entriesSection.appendChild(entriesList);

  const aiSection = document.createElement("section");
  aiSection.className = "hk-manager-detail-section hk-manager-detail-section-ai";
  aiSection.id = "hk-manager-detail-ai";
  const aiTitle = document.createElement("h4");
  aiTitle.textContent = "AI 紀錄";
  const aiContent = document.createElement("p");
  aiContent.id = "hk-manager-detail-ai-content";
  aiContent.className = "hk-manager-detail-ai";
  aiSection.appendChild(aiTitle);
  aiSection.appendChild(aiContent);

  body.appendChild(entriesSection);
  body.appendChild(aiSection);

  dialog.appendChild(header);
  dialog.appendChild(body);

  overlay.appendChild(backdrop);
  overlay.appendChild(dialog);
  document.body.appendChild(overlay);
  return overlay;
};

const closePageDetail = () => {
  const overlay = document.getElementById(detailOverlayId);
  if (!overlay) return;
  overlay.classList.add("is-hidden");
  detailCurrentPageUrl = "";
};

const renderPageDetail = (page) => {
  const overlay = ensureDetailOverlay();
  const titleEl = overlay.querySelector("#hk-manager-detail-title");
  const urlEl = overlay.querySelector("#hk-manager-detail-url");
  const metaEl = overlay.querySelector("#hk-manager-detail-meta");
  const entriesList = overlay.querySelector("#hk-manager-detail-entries");
  const aiContent = overlay.querySelector("#hk-manager-detail-ai-content");
  const aiSection = overlay.querySelector("#hk-manager-detail-ai");
  const tagsChips = overlay.querySelector("#hk-manager-detail-tags-chips");
  const tagsInput = overlay.querySelector("#hk-manager-detail-tags-input");
  if (
    !titleEl ||
    !urlEl ||
    !metaEl ||
    !entriesList ||
    !aiContent ||
    !aiSection ||
    !tagsChips ||
    !tagsInput
  )
    return;

  detailCurrentPageUrl = page.url;
  titleEl.textContent = page.title;
  urlEl.textContent = page.url;
  urlEl.href = page.url;
  metaEl.textContent = `筆記 ${page.total} 則 · 最後更新：${
    page.updatedAt
      ? new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short" }).format(
          new Date(page.updatedAt)
        )
      : "未知時間"
  }`;

  entriesList.innerHTML = "";
  if (!page.entries?.length) {
    const empty = document.createElement("p");
    empty.className = "hk-manager-detail-empty";
    empty.textContent = "尚無筆記。";
    entriesList.appendChild(empty);
  } else {
    page.entries.forEach((entry) => {
      const item = document.createElement("article");
      item.className = "hk-manager-detail-item";
      const text = document.createElement("p");
      text.className = "hk-manager-detail-text";
      text.textContent = entry.text || "(無內容)";
      item.appendChild(text);
      const trimmedNote = entry.note?.trim();
      if (trimmedNote) {
        const note = document.createElement("p");
        note.className = "hk-manager-detail-note is-clickable";
        note.textContent = trimmedNote;
        note.setAttribute("title", "點擊複製註解");
        note.addEventListener("click", async () => {
          try {
            await navigator.clipboard.writeText(trimmedNote);
            setStatus("已複製註解");
          } catch (error) {
            console.debug("複製註解失敗", error);
            setStatus("無法複製註解", true);
          }
        });
        item.appendChild(note);
      }
      entriesList.appendChild(item);
    });
  }

  const sanitizedTags = Array.isArray(page.tags)
    ? Array.from(new Set(page.tags.map((tag) => tag.trim()).filter(Boolean)))
    : [];
  tagsChips.innerHTML = "";
  if (!sanitizedTags.length) {
    const emptyTag = document.createElement("span");
    emptyTag.className = "hk-manager-detail-empty";
    emptyTag.textContent = "尚無標籤";
    tagsChips.appendChild(emptyTag);
  } else {
    sanitizedTags.forEach((tag) => {
      const chip = document.createElement("span");
      chip.className = "hk-manager-tag-chip";
      chip.textContent = tag;
      tagsChips.appendChild(chip);
    });
  }
  tagsInput.value = sanitizedTags.join(", ");

  const aiNote = state.notes?.[page.url];
  if (aiNote?.note) {
    aiContent.textContent = aiNote.note;
    aiSection.style.display = "";
  } else {
    aiContent.textContent = "尚未產生 AI 紀錄。";
    aiSection.style.display = "";
  }
};

const openPageDetail = (page) => {
  renderPageDetail(page);
  const overlay = ensureDetailOverlay();
  overlay.classList.remove("is-hidden");
};

const normalizeTagsInput = (input) => {
  if (!input) return [];
  return Array.from(
    new Set(
      input
        .split(/[,\\s]+/)
        .map((tag) => tag.trim())
        .filter(Boolean)
    )
  );
};

const savePageTagsInStorage = async (url, tags) => {
  const stored = await chrome.storage.local.get(PAGE_META_KEY);
  const current = stored[PAGE_META_KEY] || {};
  const existing = current[url] || {};
  const normalized = Array.isArray(tags) ? tags : [];
  const next = {
    ...existing,
    tags: normalized,
    updatedAt: Date.now(),
  };
  await chrome.storage.local.set({
    [PAGE_META_KEY]: {
      ...current,
      [url]: next,
    },
  });
};

const savePageTagsFromDetail = async () => {
  if (!detailCurrentPageUrl) return;
  const overlay = ensureDetailOverlay();
  const input = overlay.querySelector("#hk-manager-detail-tags-input");
  const chips = overlay.querySelector("#hk-manager-detail-tags-chips");
  if (!input || !chips) return;
  const tags = normalizeTagsInput(input.value);
  try {
    await savePageTagsInStorage(detailCurrentPageUrl, tags);
    state.meta[detailCurrentPageUrl] = {
      ...(state.meta[detailCurrentPageUrl] || {}),
      tags,
    };
    state.pages = state.pages.map((page) =>
      page.url === detailCurrentPageUrl ? { ...page, tags } : page
    );
    renderPageDetail(
      state.pages.find((page) => page.url === detailCurrentPageUrl) || {
        url: detailCurrentPageUrl,
        title: state.meta[detailCurrentPageUrl]?.title || detailCurrentPageUrl,
        total: 0,
        updatedAt: null,
        entries: [],
        tags,
      }
    );
    renderPageList();
    setStatus("已更新頁面標籤");
  } catch (error) {
    console.debug("更新頁面標籤失敗", error);
    setStatus("無法更新頁面標籤", true);
  }
};

const encodeContentToBase64 = (text) => {
  const encoder = new TextEncoder();
  const bytes = encoder.encode(text);
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
};

const decodeBase64ToText = (encoded) => {
  if (!encoded) return "";
  const sanitized = encoded.replace(/\s/g, "");
  const binary = atob(sanitized);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  const decoder = new TextDecoder();
  return decoder.decode(bytes);
};

const buildRepoApiBase = (repo) => {
  if (typeof repo !== "string") return null;
  const [owner, name] = repo.split("/").map((part) => part?.trim());
  if (!owner || !name) return null;
  return `https://api.github.com/repos/${encodeURIComponent(
    owner
  )}/${encodeURIComponent(name)}`;
};

const buildContentPath = (path) => {
  if (typeof path !== "string") return "";
  return path
    .split("/")
    .map((segment) => segment.trim())
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");
};

const getGithubSettingsSnapshot = () => {
  const token =
    githubTokenInput?.value?.trim() || githubSettings.token?.trim() || "";
  const repo =
    githubRepoInput?.value?.trim() || githubSettings.repo?.trim() || "";
  const branch =
    githubBranchInput?.value?.trim() ||
    githubSettings.branch?.trim() ||
    GITHUB_DEFAULT_SETTINGS.branch;
  const pathInput =
    githubPathInput?.value?.trim() ||
    githubSettings.path?.trim() ||
    GITHUB_DEFAULT_SETTINGS.path;
  const normalizedPath = pathInput.replace(/^\/+/, "");
  return {
    token,
    repo,
    branch: branch || GITHUB_DEFAULT_SETTINGS.branch,
    path: normalizedPath,
  };
};

const validateGithubSettings = (settings) => {
  if (!settings.token) return "請輸入 GitHub Token";
  if (!settings.repo || !settings.repo.includes("/")) {
    return "請輸入 owner/repo 格式的儲存庫";
  }
  if (!settings.path) return "請輸入檔案路徑";
  return null;
};

const fetchGithubFileSha = async (settings) => {
  const repoBase = buildRepoApiBase(settings.repo);
  if (!repoBase) throw new Error("儲存庫格式不正確");
  const encodedPath = buildContentPath(settings.path);
  const url = `${repoBase}/contents/${encodedPath}?ref=${encodeURIComponent(
    settings.branch
  )}`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${settings.token}`,
      Accept: "application/vnd.github+json",
    },
  });
  if (response.status === 404) {
    return null;
  }
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`GitHub 讀取檔案失敗：${errorText}`);
  }
  const json = await response.json();
  return json?.sha ?? null;
};

const fetchGithubBackupContent = async (settings) => {
  const repoBase = buildRepoApiBase(settings.repo);
  if (!repoBase) throw new Error("儲存庫格式不正確");
  const encodedPath = buildContentPath(settings.path);
  const url = `${repoBase}/contents/${encodedPath}?ref=${encodeURIComponent(
    settings.branch
  )}`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${settings.token}`,
      Accept: "application/vnd.github+json",
    },
  });
  if (response.status === 404) {
    throw new Error("GitHub 上找不到指定的備份檔案");
  }
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`GitHub 讀取檔案失敗：${errorText}`);
  }
  const json = await response.json();
  if (!json?.content) {
    throw new Error("GitHub 回傳的檔案內容為空");
  }
  return decodeBase64ToText(json.content);
};

const uploadHighlightsToGithub = async () => {
  const settings = getGithubSettingsSnapshot();
  const validationError = validateGithubSettings(settings);
  if (validationError) {
    setGithubStatus(validationError, true);
    return;
  }
  setGithubActionsDisabled(true);
  setGithubStatus("上傳中…");
  try {
    await fetchAllPages();
    if (!state.pages.length) {
      throw new Error("目前沒有筆記可上傳");
    }
    const payload = buildFullExportPayload();
    const content = encodeContentToBase64(JSON.stringify(payload, null, 2));
    let existingSha = null;
    try {
      existingSha = await fetchGithubFileSha(settings);
    } catch (error) {
      console.debug("查詢 GitHub 既有檔案失敗", error);
    }
    const repoBase = buildRepoApiBase(settings.repo);
    if (!repoBase) {
      throw new Error("儲存庫格式不正確");
    }
    const encodedPath = buildContentPath(settings.path);
    const url = `${repoBase}/contents/${encodedPath}`;
    const body = {
      message: `backup: highlight-keeper (${new Date().toISOString()})`,
      content,
      branch: settings.branch,
    };
    if (existingSha) {
      body.sha = existingSha;
    }
    const response = await fetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${settings.token}`,
        "Content-Type": "application/json",
        Accept: "application/vnd.github+json",
      },
      body: JSON.stringify(body),
    });
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`GitHub API 錯誤：${errorText}`);
    }
    setGithubStatus("已成功上傳到 GitHub");
    githubSettings = { ...githubSettings, ...settings };
    persistGithubSettings(githubSettings);
  } catch (error) {
    setGithubStatus(error?.message || "上傳失敗", true);
  } finally {
    setGithubActionsDisabled(false);
  }
};

const savePageMetaTitles = async (updates) => {
  if (!Object.keys(updates).length) return;
  const stored = await chrome.storage.local.get(PAGE_META_KEY);
  const current = stored[PAGE_META_KEY] || {};
  await chrome.storage.local.set({
    [PAGE_META_KEY]: {
      ...current,
      ...updates,
    },
  });
};

const applyGithubBackupPages = async (pages) => {
  if (!Array.isArray(pages) || !pages.length) {
    setGithubStatus("GitHub 備份中沒有可匯入的頁面", true);
    return;
  }
  await fetchAllPages();
  const seen = new Set();
  const uniquePages = pages.filter((page) => {
    if (seen.has(page.url)) return false;
    seen.add(page.url);
    return true;
  });
  const urls = uniquePages.map((page) => page.url);
  const existing = await chrome.storage.local.get(urls);
  const overlapping = uniquePages
    .filter((page) => Array.isArray(existing[page.url]) && existing[page.url].length)
    .map((page) => page.url);
  let pagesToImport = uniquePages;
  let overwrittenCount = overlapping.length;
  if (overlapping.length) {
    const samples = overlapping.slice(0, 3).map((url) => `- ${getPageDisplayName(url)}`);
    const confirmMessage = [
      `有 ${overlapping.length} 個頁面在本機已有筆記。`,
      ...samples,
      overlapping.length > samples.length ? "..." : null,
      "要覆蓋這些頁面並改用 GitHub 版本嗎？",
      "按「取消」則只匯入新的頁面並保留本機資料。",
    ]
      .filter(Boolean)
      .join("\n");
    const shouldOverride = window.confirm(confirmMessage);
    if (!shouldOverride) {
      pagesToImport = uniquePages.filter((page) => !overlapping.includes(page.url));
      overwrittenCount = 0;
    }
  }
  if (!pagesToImport.length) {
    setGithubStatus("已取消匯入，本機筆記保持不變。");
    return;
  }
  const updates = Object.fromEntries(
    pagesToImport.map((page) => [page.url, page.entries])
  );
  const metaUpdates = {};
  pagesToImport.forEach((page) => {
    if (page.title) {
      metaUpdates[page.url] = {
        ...(state.meta[page.url] || {}),
        title: page.title,
      };
    }
  });
  await chrome.storage.local.set(updates);
  await savePageMetaTitles(metaUpdates);
  const importedCount = pagesToImport.length;
  const skippedOverlap = overlapping.length - overwrittenCount;
  const statusParts = [
    `已從 GitHub 匯入 ${importedCount} 個頁面`,
    overwrittenCount ? `覆蓋 ${overwrittenCount} 個已存在頁面` : null,
    skippedOverlap > 0 ? `保留 ${skippedOverlap} 個本機版本` : null,
  ].filter(Boolean);
  setGithubStatus(statusParts.join("，"));
  await refreshManager();
};

const downloadHighlightsFromGithub = async () => {
  const settings = getGithubSettingsSnapshot();
  const validationError = validateGithubSettings(settings);
  if (validationError) {
    setGithubStatus(validationError, true);
    return;
  }
  setGithubActionsDisabled(true);
  setGithubStatus("從 GitHub 下載中…");
  try {
    const content = await fetchGithubBackupContent(settings);
    const pages = parseGithubBackupPayload(content);
    await applyGithubBackupPages(pages);
    githubSettings = { ...githubSettings, ...settings };
    persistGithubSettings(githubSettings);
  } catch (error) {
    setGithubStatus(error?.message || "下載失敗", true);
  } finally {
    setGithubActionsDisabled(false);
  }
};

const importMultipleFiles = async (files) => {
  if (!files?.length) return;
  setStatus("解析匯入檔案中…");
  const filePromises = Array.from(files).map(async (file) => {
    const text = await file.text();
    return parseImportedHighlightsPayload(text);
  });
  const fileEntries = await Promise.all(filePromises);
  const flattened = fileEntries.flat();
  const normalized = flattened
    .map((entry, index) => normalizeImportedHighlightEntry(entry, index))
    .filter(Boolean);
  if (!normalized.length) {
    setStatus("沒有可匯入的筆記", true);
    return;
  }
  const grouped = new Map();
  normalized.forEach((entry) => {
    if (!grouped.has(entry.url)) {
      grouped.set(entry.url, []);
    }
    grouped.get(entry.url).push(entry);
  });
  const urls = Array.from(grouped.keys());
  const existing = await chrome.storage.local.get(urls);
  const updates = {};
  const metaUpdates = {};
  const skipped = [];
  grouped.forEach((list, url) => {
    const current = existing[url];
    if (Array.isArray(current) && current.length) {
      skipped.push(url);
      return;
    }
    updates[url] = list.map(({ title, ...rest }) => rest);
    const titleEntry = list.find((item) => item.title);
    if (titleEntry?.title) {
      metaUpdates[url] = {
        ...(state.meta[url] || {}),
        title: titleEntry.title,
      };
    }
  });
  const importedPages = Object.keys(updates).length;
  if (!importedPages) {
    setStatus("所有頁面皆已有筆記，已忽略匯入。", true);
    return;
  }
  await chrome.storage.local.set(updates);
  await savePageMetaTitles(metaUpdates);
  setStatus(
    `成功匯入 ${importedPages} 個頁面，跳過 ${skipped.length} 個已存在的頁面。`
  );
  await refreshManager();
};

const init = () => {
  refreshManager().catch((error) => {
    console.debug("載入筆記失敗", error);
    setStatus("無法載入筆記", true);
  });
  downloadBtn?.addEventListener("click", () => {
    downloadAllPages().catch((error) => {
      console.debug("下載全部筆記失敗", error);
      setStatus(error?.message || "下載失敗", true);
    });
  });
  importInput?.addEventListener("change", (event) => {
    importMultipleFiles(event.target.files).catch((error) => {
      console.debug("匯入失敗", error);
      setStatus(error?.message || "匯入失敗", true);
    });
    event.target.value = "";
  });
  closeBtn?.addEventListener("click", () => {
    window.close();
  });
  searchInput?.addEventListener("input", (event) => {
    state.searchTerm = event.target.value ?? "";
    renderPageList();
  });
  bindGithubInput(githubTokenInput, "token");
  bindGithubInput(githubRepoInput, "repo");
  bindGithubInput(githubBranchInput, "branch");
  bindGithubInput(githubPathInput, "path");
  githubDownloadBtn?.addEventListener("click", () => {
    downloadHighlightsFromGithub();
  });
  githubUploadBtn?.addEventListener("click", () => {
    uploadHighlightsToGithub();
  });
  loadGithubSettings().catch((error) => {
    console.debug("初始化 GitHub 設定失敗", error);
  });
};

init();
