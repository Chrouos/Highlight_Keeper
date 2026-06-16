const PAGE_META_KEY = "__hk_page_meta__";
const DEFAULT_COLOR = "#ffeb3b";
const GITHUB_SETTINGS_KEY = "hkGithubSyncSettings";
const GITHUB_DEFAULT_SETTINGS = {
  token: "",
  repo: "",
  branch: "main",
  path: "backups/highlight-keeper.json",
};

const MINDMAP_KEY = "hkMindmaps";
const GENERATED_NOTES_KEY = "hkGeneratedNotes";

const state = {
  pages: [],
  meta: {},
  notes: {},
  mindmaps: {},
  searchTerm: "",
  sortBy: "updated",
  activeTag: "",
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
const sortSelect = document.getElementById("managerSort");
const tagFilterEl = document.getElementById("managerTagFilter");
const githubTokenInput = document.getElementById("githubToken");
const githubRepoInput = document.getElementById("githubRepo");
const githubBranchInput = document.getElementById("githubBranch");
const githubPathInput = document.getElementById("githubPath");
const githubDownloadBtn = document.getElementById("githubDownloadBtn");
const githubUploadBtn = document.getElementById("githubUploadBtn");
const githubStatusEl = document.getElementById("githubSyncStatus");
const detailOverlayId = "hk-manager-detail";
const confirmOverlayId = "hk-manager-confirm";

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
  const notes = all[GENERATED_NOTES_KEY] || {};
  const mindmaps = all[MINDMAP_KEY] || {};
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
  state.mindmaps = mindmaps;
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

const sortPages = (pages) => {
  const sorted = [...pages];
  if (state.sortBy === "count") {
    sorted.sort((a, b) => (b.total || 0) - (a.total || 0));
  } else if (state.sortBy === "title") {
    sorted.sort((a, b) =>
      String(a.title).localeCompare(String(b.title), undefined, {
        sensitivity: "base",
      })
    );
  } else {
    sorted.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  }
  return sorted;
};

// 收集所有頁面的標籤聯集，畫成可點選的篩選晶片。
const renderTagFilter = () => {
  if (!tagFilterEl) return;
  const tagSet = new Set();
  state.pages.forEach((page) => {
    (Array.isArray(page.tags) ? page.tags : []).forEach((tag) => {
      const trimmed = String(tag).trim();
      if (trimmed) tagSet.add(trimmed);
    });
  });
  tagFilterEl.innerHTML = "";
  if (!tagSet.size) {
    tagFilterEl.style.display = "none";
    return;
  }
  tagFilterEl.style.display = "flex";
  const makeChip = (label, value) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "hk-manager-filter-chip";
    chip.textContent = label;
    chip.classList.toggle("is-active", state.activeTag === value);
    chip.addEventListener("click", () => {
      state.activeTag = state.activeTag === value ? "" : value;
      renderPageList();
    });
    return chip;
  };
  tagFilterEl.appendChild(makeChip("全部", ""));
  Array.from(tagSet)
    .sort((a, b) => a.localeCompare(b))
    .forEach((tag) => tagFilterEl.appendChild(makeChip(tag, tag)));
};

const renderPageList = () => {
  if (!listEl || !pageCountEl) return;
  listEl.innerHTML = "";
  renderTagFilter();
  let filtered = state.pages.filter((page) =>
    matchesSearch(page, state.searchTerm)
  );
  if (state.activeTag) {
    filtered = filtered.filter(
      (page) =>
        Array.isArray(page.tags) && page.tags.includes(state.activeTag)
    );
  }
  filtered = sortPages(filtered);
  pageCountEl.textContent = `共 ${filtered.length} 個頁面`;
  if (!filtered.length) {
    const empty = document.createElement("p");
    empty.textContent =
      state.searchTerm || state.activeTag
        ? "沒有符合條件的頁面。"
        : "目前尚未建立任何筆記。";
    empty.className = "hk-manager-meta";
    listEl.appendChild(empty);
    return;
  }
  filtered.forEach((page) => {
    const card = document.createElement("article");
    card.className = "hk-manager-card";
    card.addEventListener("click", () => openPageDetail(page));

    const deletePageBtn = document.createElement("button");
    deletePageBtn.type = "button";
    deletePageBtn.className = "hk-manager-card-delete";
    deletePageBtn.textContent = "刪除";
    deletePageBtn.title = "刪除整頁筆記";
    deletePageBtn.addEventListener("click", async (event) => {
      event.stopPropagation();
      const ok = await openConfirmDialog({
        title: "刪除整頁",
        message: `確定刪除「${page.title}」的所有標註、摘要與心智圖嗎？此動作無法復原。`,
        confirmLabel: "刪除",
        cancelLabel: "取消",
      });
      if (!ok) return;
      try {
        await deletePage(page.url);
        if (detailCurrentPageUrl === page.url) closePageDetail();
        await refreshManager();
        setStatus("已刪除整頁筆記");
      } catch (error) {
        console.debug("刪除整頁失敗", error);
        setStatus("刪除失敗", true);
      }
    });
    card.appendChild(deletePageBtn);

    const title = document.createElement("h3");
    title.textContent = page.title;
    card.appendChild(title);
    const urlLink = document.createElement("a");
    urlLink.href = page.url;
    urlLink.target = "_blank";
    urlLink.rel = "noopener";
    urlLink.className = "hk-manager-url";
    urlLink.textContent = page.url;
    urlLink.addEventListener("click", (event) => event.stopPropagation());
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

// 刪除整頁：清掉標註、頁面 meta、AI 摘要與心智圖。無法復原。
const deletePage = async (url) => {
  const stored = await chrome.storage.local.get([
    PAGE_META_KEY,
    GENERATED_NOTES_KEY,
    MINDMAP_KEY,
  ]);
  const meta = { ...(stored[PAGE_META_KEY] || {}) };
  const notes = { ...(stored[GENERATED_NOTES_KEY] || {}) };
  const maps = { ...(stored[MINDMAP_KEY] || {}) };
  delete meta[url];
  delete notes[url];
  delete maps[url];
  await chrome.storage.local.remove(url);
  await chrome.storage.local.set({
    [PAGE_META_KEY]: meta,
    [GENERATED_NOTES_KEY]: notes,
    [MINDMAP_KEY]: maps,
  });
};

// 刪除單筆標註：優先用 id 比對，沒有 id 才退回索引。
const deleteHighlightEntry = async (url, entry, index) => {
  const stored = await chrome.storage.local.get(url);
  const entries = Array.isArray(stored[url]) ? stored[url] : [];
  const next = entry?.id
    ? entries.filter((item) => item?.id !== entry.id)
    : entries.filter((_, i) => i !== index);
  if (next.length === entries.length) return false;
  await chrome.storage.local.set({ [url]: next });
  return true;
};

const buildFullExportPayload = () => ({
  type: "highlight-keeper-bulk",
  version: 2,
  exportedAt: Date.now(),
  pages: state.pages.map((page) => ({
    url: page.url,
    title: state.meta[page.url]?.title || "",
    tags: Array.isArray(page.tags) ? page.tags : [],
    entries: page.entries,
    note: state.notes[page.url] || null,
    mindmap: state.mindmaps[page.url] || null,
  })),
});

// 把 bulk 備份的 pages 陣列正規化成 {url,title,tags,entries,note,mindmap}。
const normalizeBulkPages = (pages) => {
  if (!Array.isArray(pages)) return [];
  return pages
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
        tags: Array.isArray(page.tags)
          ? page.tags.map((t) => String(t).trim()).filter(Boolean)
          : [],
        entries,
        note: page.note && typeof page.note === "object" ? page.note : null,
        mindmap:
          page.mindmap && typeof page.mindmap === "object" ? page.mindmap : null,
      };
    })
    .filter(Boolean);
};

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
  const normalized = normalizeBulkPages(pages);
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

const ensureConfirmOverlay = () => {
  let overlay = document.getElementById(confirmOverlayId);
  if (overlay) return overlay;

  overlay = document.createElement("div");
  overlay.id = confirmOverlayId;
  overlay.className = "hk-manager-confirm is-hidden";
  overlay.setAttribute("role", "dialog");
  overlay.setAttribute("aria-modal", "true");
  overlay.setAttribute("aria-labelledby", "hk-manager-confirm-title");
  overlay.setAttribute("aria-describedby", "hk-manager-confirm-message");

  const backdrop = document.createElement("div");
  backdrop.className = "hk-manager-confirm-backdrop";

  const dialog = document.createElement("div");
  dialog.className = "hk-manager-confirm-dialog";
  dialog.tabIndex = -1;

  const title = document.createElement("h3");
  title.id = "hk-manager-confirm-title";

  const message = document.createElement("p");
  message.id = "hk-manager-confirm-message";
  message.className = "hk-manager-confirm-message";

  const actions = document.createElement("div");
  actions.className = "hk-manager-confirm-actions";

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.id = "hk-manager-confirm-cancel";
  cancelBtn.className = "hk-manager-btn hk-manager-btn-ghost";
  cancelBtn.textContent = "取消";

  const confirmBtn = document.createElement("button");
  confirmBtn.type = "button";
  confirmBtn.id = "hk-manager-confirm-ok";
  confirmBtn.className = "hk-manager-btn";
  confirmBtn.textContent = "確定";

  actions.appendChild(cancelBtn);
  actions.appendChild(confirmBtn);

  dialog.appendChild(title);
  dialog.appendChild(message);
  dialog.appendChild(actions);

  overlay.appendChild(backdrop);
  overlay.appendChild(dialog);
  document.body.appendChild(overlay);
  return overlay;
};

const openConfirmDialog = ({ title, message, confirmLabel, cancelLabel }) =>
  new Promise((resolve) => {
    const overlay = ensureConfirmOverlay();
    const titleEl = overlay.querySelector("#hk-manager-confirm-title");
    const messageEl = overlay.querySelector("#hk-manager-confirm-message");
    const confirmBtn = overlay.querySelector("#hk-manager-confirm-ok");
    const cancelBtn = overlay.querySelector("#hk-manager-confirm-cancel");
    const backdrop = overlay.querySelector(".hk-manager-confirm-backdrop");

    if (!titleEl || !messageEl || !confirmBtn || !cancelBtn || !backdrop) {
      resolve(false);
      return;
    }

    titleEl.textContent = title || "確認動作";
    messageEl.textContent = message || "";
    confirmBtn.textContent = confirmLabel || "確定";
    cancelBtn.textContent = cancelLabel || "取消";

    let resolved = false;
    const cleanup = (result) => {
      if (resolved) return;
      resolved = true;
      overlay.classList.add("is-hidden");
      overlay.removeEventListener("keydown", onKeydown);
      confirmBtn.onclick = null;
      cancelBtn.onclick = null;
      backdrop.onclick = null;
      resolve(result);
    };

    const onKeydown = (event) => {
      if (event.key === "Escape") {
        cleanup(false);
      }
    };

    confirmBtn.onclick = () => cleanup(true);
    cancelBtn.onclick = () => cleanup(false);
    backdrop.onclick = () => cleanup(false);
    overlay.addEventListener("keydown", onKeydown);

    overlay.classList.remove("is-hidden");
    confirmBtn.focus();
  });

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

  const mindmapSection = document.createElement("section");
  mindmapSection.className =
    "hk-manager-detail-section hk-manager-detail-section-mindmap";
  mindmapSection.id = "hk-manager-detail-mindmap";
  const mindmapHead = document.createElement("div");
  mindmapHead.className = "hk-manager-detail-mindmap-head";
  const mindmapTitle = document.createElement("h4");
  mindmapTitle.textContent = "心智圖";
  const mindmapCopyBtn = document.createElement("button");
  mindmapCopyBtn.type = "button";
  mindmapCopyBtn.id = "hk-manager-detail-mindmap-copy";
  mindmapCopyBtn.className = "hk-manager-btn hk-manager-btn-muted";
  mindmapCopyBtn.textContent = "複製大綱";
  mindmapHead.appendChild(mindmapTitle);
  mindmapHead.appendChild(mindmapCopyBtn);
  const mindmapContent = document.createElement("pre");
  mindmapContent.id = "hk-manager-detail-mindmap-content";
  mindmapContent.className = "hk-manager-detail-mindmap";
  mindmapSection.appendChild(mindmapHead);
  mindmapSection.appendChild(mindmapContent);

  body.appendChild(entriesSection);
  body.appendChild(aiSection);
  body.appendChild(mindmapSection);

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
    page.entries.forEach((entry, index) => {
      const item = document.createElement("article");
      item.className = "hk-manager-detail-item";

      const itemHead = document.createElement("div");
      itemHead.className = "hk-manager-detail-item-head";
      const text = document.createElement("p");
      text.className = "hk-manager-detail-text";
      text.textContent = entry.text || "(無內容)";
      itemHead.appendChild(text);
      const delEntryBtn = document.createElement("button");
      delEntryBtn.type = "button";
      delEntryBtn.className = "hk-manager-detail-item-delete";
      delEntryBtn.textContent = "刪除";
      delEntryBtn.title = "刪除這筆標註";
      delEntryBtn.addEventListener("click", async () => {
        const ok = await openConfirmDialog({
          title: "刪除標註",
          message: "確定刪除這筆標註嗎？此動作無法復原。",
          confirmLabel: "刪除",
          cancelLabel: "取消",
        });
        if (!ok) return;
        try {
          await deleteHighlightEntry(page.url, entry, index);
          await fetchAllPages();
          const updated = state.pages.find((p) => p.url === page.url);
          if (updated) renderPageDetail(updated);
          else closePageDetail();
          renderPageList();
          setStatus("已刪除標註");
        } catch (error) {
          console.debug("刪除標註失敗", error);
          setStatus("刪除失敗", true);
        }
      });
      itemHead.appendChild(delEntryBtn);
      item.appendChild(itemHead);

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

  // 心智圖：有大綱才顯示這個區塊
  const mindmapSection = overlay.querySelector("#hk-manager-detail-mindmap");
  const mindmapContent = overlay.querySelector(
    "#hk-manager-detail-mindmap-content"
  );
  const mindmapCopyBtn = overlay.querySelector(
    "#hk-manager-detail-mindmap-copy"
  );
  if (mindmapSection && mindmapContent && mindmapCopyBtn) {
    const outline = state.mindmaps?.[page.url]?.outline?.trim();
    if (outline) {
      mindmapSection.style.display = "";
      mindmapContent.textContent = outline;
      mindmapCopyBtn.disabled = false;
      mindmapCopyBtn.onclick = async () => {
        try {
          await navigator.clipboard.writeText(outline);
          setStatus("已複製心智圖大綱");
        } catch (error) {
          console.debug("複製心智圖大綱失敗", error);
          setStatus("複製失敗", true);
        }
      };
    } else {
      mindmapSection.style.display = "none";
      mindmapContent.textContent = "";
      mindmapCopyBtn.onclick = null;
    }
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
  if (Array.isArray(json)) {
    throw new Error("GitHub 路徑指向資料夾，請指定備份檔案");
  }
  if (json?.content) {
    return decodeBase64ToText(json.content);
  }
  if (json?.download_url) {
    const rawResponse = await fetch(json.download_url, {
      headers: {
        Authorization: `Bearer ${settings.token}`,
        Accept: "application/vnd.github+json",
      },
    });
    if (!rawResponse.ok) {
      const errorText = await rawResponse.text();
      throw new Error(`GitHub 讀取檔案失敗：${errorText}`);
    }
    return await rawResponse.text();
  }
  if (json?.sha) {
    const blobResponse = await fetch(`${repoBase}/git/blobs/${json.sha}`, {
      headers: {
        Authorization: `Bearer ${settings.token}`,
        Accept: "application/vnd.github+json",
      },
    });
    if (!blobResponse.ok) {
      const errorText = await blobResponse.text();
      throw new Error(`GitHub 讀取檔案失敗：${errorText}`);
    }
    const blob = await blobResponse.json();
    if (!blob?.content) {
      throw new Error("GitHub 回傳的檔案內容為空");
    }
    return decodeBase64ToText(blob.content);
  }
  throw new Error("GitHub 回傳的檔案內容為空");
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

// 還原 AI 摘要與心智圖（v2 備份才有；只覆蓋有帶資料的頁面）。
const saveNotesAndMindmaps = async (noteUpdates, mindmapUpdates) => {
  if (!Object.keys(noteUpdates).length && !Object.keys(mindmapUpdates).length) {
    return;
  }
  const stored = await chrome.storage.local.get([
    GENERATED_NOTES_KEY,
    MINDMAP_KEY,
  ]);
  const nextNotes = { ...(stored[GENERATED_NOTES_KEY] || {}), ...noteUpdates };
  const nextMaps = { ...(stored[MINDMAP_KEY] || {}), ...mindmapUpdates };
  await chrome.storage.local.set({
    [GENERATED_NOTES_KEY]: nextNotes,
    [MINDMAP_KEY]: nextMaps,
  });
};

// 從一批已決定要匯入的頁面，組出 meta（title+tags）／notes／mindmaps 並寫入。
const applyImportedPageExtras = async (pages) => {
  const metaUpdates = {};
  const noteUpdates = {};
  const mindmapUpdates = {};
  pages.forEach((page) => {
    const metaEntry = { ...(state.meta[page.url] || {}) };
    let touched = false;
    if (page.title) {
      metaEntry.title = page.title;
      touched = true;
    }
    if (Array.isArray(page.tags) && page.tags.length) {
      metaEntry.tags = page.tags;
      touched = true;
    }
    if (touched) metaUpdates[page.url] = metaEntry;
    if (page.note) noteUpdates[page.url] = page.note;
    if (page.mindmap) mindmapUpdates[page.url] = page.mindmap;
  });
  await savePageMetaTitles(metaUpdates);
  await saveNotesAndMindmaps(noteUpdates, mindmapUpdates);
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
  const overlapping = uniquePages.filter(
    (page) => Array.isArray(existing[page.url]) && existing[page.url].length
  );
  const toSortedJson = (value) => JSON.stringify(value ?? null);
  const isSameEntries = (currentEntries, nextEntries) =>
    toSortedJson(currentEntries) === toSortedJson(nextEntries);
  const identical = overlapping.filter((page) =>
    isSameEntries(existing[page.url], page.entries)
  );
  const changed = overlapping.filter(
    (page) => !identical.includes(page)
  );
  const changedUrls = changed.map((page) => page.url);
  const identicalUrls = identical.map((page) => page.url);
  let pagesToImport = uniquePages.filter((page) => !identicalUrls.includes(page.url));
  let overwrittenCount = changed.length;
  if (changed.length) {
    const samples = changedUrls.slice(0, 3).map((url) => `- ${getPageDisplayName(url)}`);
    const confirmMessage = [
      `有 ${changed.length} 個頁面在本機已有筆記，且內容不同。`,
      ...samples,
      changed.length > samples.length ? "..." : null,
      "要覆蓋這些頁面並改用 GitHub 版本嗎？",
      "按「取消」則只匯入新的頁面並保留本機資料。",
    ]
      .filter(Boolean)
      .join("\n");
    const shouldOverride = await openConfirmDialog({
      title: "確認匯入方式",
      message: confirmMessage,
      confirmLabel: "覆蓋本機版本",
      cancelLabel: "只匯入新頁面",
    });
    if (!shouldOverride) {
      pagesToImport = pagesToImport.filter((page) => !changedUrls.includes(page.url));
      overwrittenCount = 0;
    }
  }
  if (!pagesToImport.length) {
    if (identical.length) {
      setGithubStatus("已是最新版本，沒有需要更新的筆記。");
    } else {
      setGithubStatus("已取消匯入，本機筆記保持不變。");
    }
    return;
  }
  const updates = Object.fromEntries(
    pagesToImport.map((page) => [page.url, page.entries])
  );
  await chrome.storage.local.set(updates);
  await applyImportedPageExtras(pagesToImport);
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
  const texts = await Promise.all(Array.from(files).map((file) => file.text()));

  // 兩種格式並存：bulk（{pages:[...]}，含摘要/心智圖/標籤）與舊的單筆標註陣列。
  const bulkByUrl = new Map();
  const looseEntries = [];
  texts.forEach((text) => {
    let parsed = null;
    try {
      parsed = JSON.parse(text);
    } catch (_error) {
      return;
    }
    if (Array.isArray(parsed?.pages)) {
      normalizeBulkPages(parsed.pages).forEach((page) => {
        if (!bulkByUrl.has(page.url)) bulkByUrl.set(page.url, page);
      });
    } else {
      const arr = Array.isArray(parsed)
        ? parsed
        : Array.isArray(parsed?.entries)
        ? parsed.entries
        : parsed && typeof parsed === "object"
        ? [parsed]
        : [];
      looseEntries.push(...arr);
    }
  });

  // 單筆標註：正規化後分組成頁面（無摘要/心智圖）；bulk 已涵蓋的頁面就略過。
  const looseByUrl = new Map();
  looseEntries
    .map((entry, index) => normalizeImportedHighlightEntry(entry, index))
    .filter(Boolean)
    .forEach((entry) => {
      if (bulkByUrl.has(entry.url)) return;
      if (!looseByUrl.has(entry.url)) {
        looseByUrl.set(entry.url, {
          url: entry.url,
          title: entry.title || "",
          tags: [],
          entries: [],
          note: null,
          mindmap: null,
        });
      }
      looseByUrl.get(entry.url).entries.push(entry);
    });

  const candidatePages = [...bulkByUrl.values(), ...looseByUrl.values()];
  if (!candidatePages.length) {
    setStatus("沒有可匯入的筆記", true);
    return;
  }

  const existing = await chrome.storage.local.get(candidatePages.map((p) => p.url));
  const toImport = [];
  const skipped = [];
  candidatePages.forEach((page) => {
    const current = existing[page.url];
    if (Array.isArray(current) && current.length) {
      skipped.push(page.url);
      return;
    }
    toImport.push(page);
  });
  if (!toImport.length) {
    setStatus("所有頁面皆已有筆記，已忽略匯入。", true);
    return;
  }

  const updates = Object.fromEntries(
    toImport.map((page) => [
      page.url,
      page.entries.map(({ title, ...rest }) => rest),
    ])
  );
  await chrome.storage.local.set(updates);
  await applyImportedPageExtras(toImport);
  setStatus(
    `成功匯入 ${toImport.length} 個頁面，跳過 ${skipped.length} 個已存在的頁面。`
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
  sortSelect?.addEventListener("change", (event) => {
    state.sortBy = event.target.value || "updated";
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
