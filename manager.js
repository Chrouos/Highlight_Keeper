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
  selected: new Set(),
};
let detailCurrentPageUrl = "";
let githubSettings = { ...GITHUB_DEFAULT_SETTINGS };

const statusEl = document.getElementById("managerStatus");
const listEl = document.getElementById("pageList");
const pageCountEl = document.getElementById("pageCount");
const downloadBtn = document.getElementById("downloadAllBtn");
const mergeBtn = document.getElementById("mergeDuplicatesBtn");
const importInput = document.getElementById("bulkImportInput");
const closeBtn = document.getElementById("closeManagerBtn");
const searchInput = document.getElementById("managerSearch");
const sortSelect = document.getElementById("managerSort");
const tagFilterEl = document.getElementById("managerTagFilter");
const langSelect = document.getElementById("managerLang");

// i18n（共用 i18n.js；以一般 script 在本模組前載入，window.HkI18n 可用）
const HkI18n = typeof window !== "undefined" ? window.HkI18n : null;
const t = (key, params) => (HkI18n ? HkI18n.t(key, params) : key);
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
  // 預先算好每頁的小寫搜尋字串，避免每次按鍵都重掃所有標註。
  pages.forEach((page) => {
    const parts = [page.title, page.url, ...(page.tags || [])];
    (page.entries || []).forEach((entry) => {
      if (entry?.text) parts.push(entry.text);
      if (entry?.note) parts.push(entry.note);
      if (Array.isArray(entry?.tags)) parts.push(...entry.tags);
    });
    page._search = parts.filter(Boolean).join("\n").toLowerCase();
  });
  state.pages = pages;
  state.meta = meta;
  state.notes = notes;
  state.mindmaps = mindmaps;
};

const matchesSearch = (page, term) => {
  if (!term) return true;
  const normalized = term.trim().toLowerCase();
  if (!normalized) return true;
  // 用 fetchAllPages 預算好的 _search 字串；舊資料保險用 fallback。
  const haystack =
    typeof page._search === "string"
      ? page._search
      : [page.title, page.url].filter(Boolean).join("\n").toLowerCase();
  return haystack.includes(normalized);
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
  tagFilterEl.appendChild(makeChip(t("manager.filterAll"), ""));
  Array.from(tagSet)
    .sort((a, b) => a.localeCompare(b))
    .forEach((tag) => tagFilterEl.appendChild(makeChip(tag, tag)));
};

const ensureSelectionBar = () => {
  let bar = document.getElementById("managerSelectionBar");
  if (bar) return bar;
  bar = document.createElement("div");
  bar.id = "managerSelectionBar";
  bar.className = "hk-manager-selection-bar";
  if (listEl?.parentNode) listEl.parentNode.insertBefore(bar, listEl);
  return bar;
};

// 把目前選取的頁面打包下載（沿用全部下載的流程，只是過濾頁面）。
const exportSelectedPages = () => {
  const urls = new Set(state.selected);
  const payload = buildFullExportPayload();
  payload.pages = payload.pages.filter((p) => urls.has(p.url));
  if (!payload.pages.length) return;
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "highlight-keeper-selected.json";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  setStatus(t("manager.statusExportedSelected", { count: payload.pages.length }));
};

const bulkDeleteSelected = async () => {
  const urls = [...state.selected];
  if (!urls.length) return;
  const ok = await openConfirmDialog({
    title: t("manager.bulkDeleteTitle"),
    message: t("manager.bulkDeleteConfirm", { count: urls.length }),
    confirmLabel: t("manager.deleteSelected"),
    cancelLabel: t("manager.btnCancel"),
  });
  if (!ok) return;
  for (const url of urls) {
    try {
      await deletePage(url);
    } catch (error) {
      console.debug("批次刪除失敗", error);
    }
  }
  state.selected.clear();
  if (detailCurrentPageUrl && urls.includes(detailCurrentPageUrl)) {
    closePageDetail();
  }
  await refreshManager();
  setStatus(t("manager.statusBulkDeleted", { count: urls.length }));
};

const renderSelectionBar = (filtered) => {
  const bar = ensureSelectionBar();
  bar.innerHTML = "";
  if (!state.pages.length) {
    bar.style.display = "none";
    return;
  }
  bar.style.display = "";
  const filteredUrls = filtered.map((p) => p.url);
  const allChecked =
    filteredUrls.length > 0 && filteredUrls.every((u) => state.selected.has(u));
  const someChecked = filteredUrls.some((u) => state.selected.has(u));

  const allLabel = document.createElement("label");
  allLabel.className = "hk-manager-select-all";
  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.checked = allChecked;
  cb.indeterminate = !allChecked && someChecked;
  cb.addEventListener("change", () => {
    if (cb.checked) filteredUrls.forEach((u) => state.selected.add(u));
    else filteredUrls.forEach((u) => state.selected.delete(u));
    renderPageList();
  });
  const labelText = document.createElement("span");
  labelText.textContent = state.selected.size
    ? t("manager.selectedCount", { count: state.selected.size })
    : t("manager.selectAll");
  allLabel.appendChild(cb);
  allLabel.appendChild(labelText);
  bar.appendChild(allLabel);

  if (state.selected.size > 0) {
    const actions = document.createElement("div");
    actions.className = "hk-manager-selection-actions";
    const exportBtn = document.createElement("button");
    exportBtn.type = "button";
    exportBtn.className = "hk-manager-btn hk-manager-btn-ghost";
    exportBtn.textContent = t("manager.exportSelected");
    exportBtn.addEventListener("click", () => exportSelectedPages());
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "hk-manager-btn hk-manager-btn-danger";
    delBtn.textContent = t("manager.deleteSelected");
    delBtn.addEventListener("click", () => bulkDeleteSelected());
    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.className = "hk-manager-btn hk-manager-btn-ghost";
    clearBtn.textContent = t("manager.clearSelection");
    clearBtn.addEventListener("click", () => {
      state.selected.clear();
      renderPageList();
    });
    actions.appendChild(exportBtn);
    actions.appendChild(delBtn);
    actions.appendChild(clearBtn);
    bar.appendChild(actions);
  }
};

const renderPageList = () => {
  if (!listEl || !pageCountEl) return;
  listEl.innerHTML = "";
  renderTagFilter();
  // 清掉已不存在頁面的選取
  const existingUrls = new Set(state.pages.map((p) => p.url));
  [...state.selected].forEach((u) => {
    if (!existingUrls.has(u)) state.selected.delete(u);
  });
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
  pageCountEl.textContent = t("manager.statusPageCount", { count: filtered.length });
  renderSelectionBar(filtered);
  if (!filtered.length) {
    const empty = document.createElement("p");
    empty.textContent =
      state.searchTerm || state.activeTag
        ? t("manager.emptySearch")
        : t("manager.emptyAll");
    empty.className = "hk-manager-meta";
    listEl.appendChild(empty);
    return;
  }
  filtered.forEach((page) => {
    const card = document.createElement("article");
    card.className = "hk-manager-card";
    if (state.selected.has(page.url)) card.classList.add("is-selected");
    card.addEventListener("click", () => openPageDetail(page));

    const selectCb = document.createElement("input");
    selectCb.type = "checkbox";
    selectCb.className = "hk-manager-card-select";
    selectCb.checked = state.selected.has(page.url);
    selectCb.title = t("manager.selectPage");
    selectCb.addEventListener("click", (event) => event.stopPropagation());
    selectCb.addEventListener("change", () => {
      if (selectCb.checked) state.selected.add(page.url);
      else state.selected.delete(page.url);
      card.classList.toggle("is-selected", selectCb.checked);
      renderSelectionBar(filtered);
    });
    card.appendChild(selectCb);

    const deletePageBtn = document.createElement("button");
    deletePageBtn.type = "button";
    deletePageBtn.className = "hk-manager-card-delete";
    deletePageBtn.textContent = t("manager.deletePage");
    deletePageBtn.title = t("manager.deletePageTitle");
    deletePageBtn.addEventListener("click", async (event) => {
      event.stopPropagation();
      const ok = await openConfirmDialog({
        title: t("manager.confirmDeletePageTitle"),
        message: t("manager.confirmDeletePage", { title: page.title }),
        confirmLabel: t("manager.deletePage"),
        cancelLabel: t("manager.btnCancel"),
      });
      if (!ok) return;
      try {
        await deletePage(page.url);
        if (detailCurrentPageUrl === page.url) closePageDetail();
        await refreshManager();
        setStatus(t("manager.statusPageDeleted"));
      } catch (error) {
        console.debug("刪除整頁失敗", error);
        setStatus(t("manager.statusDeleteFail"), true);
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
      : t("manager.unknownTime");
    meta.textContent = t("manager.pageMetaList", { total: page.total, updated: updatedText });
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

// ── 整理重複頁面 ────────────────────────────────────────────────
// 同一條路徑只差網址參數（?ref=、追蹤碼、參數順序…）會被存成好幾筆，
// 這裡用 shared.js 的 normalizePageKey 收斂成「主頁面」再合併。
const canonicalPageKey = (url) =>
  window.HkUrlKey ? window.HkUrlKey.normalizePageKey(url) : url;

// 把多個頁面的標註陣列合成一份，依 id（沒有 id 才退回整筆 JSON）去重，
// 再依建立時間排序，避免合併後出現重複或亂序。
const mergeEntryArrays = (arrays) => {
  const seen = new Set();
  const merged = [];
  arrays.forEach((entries) => {
    (Array.isArray(entries) ? entries : []).forEach((entry) => {
      const sig = entry?.id ? `id:${entry.id}` : `raw:${JSON.stringify(entry)}`;
      if (seen.has(sig)) return;
      seen.add(sig);
      merged.push(entry);
    });
  });
  merged.sort(
    (a, b) => (Number(a?.createdAt) || 0) - (Number(b?.createdAt) || 0)
  );
  return merged;
};

// 找出正規化後 key 相同、卻分裂成多筆的頁面群組。
const findDuplicateGroups = () => {
  const groups = new Map();
  state.pages.forEach((page) => {
    const canonical = canonicalPageKey(page.url);
    if (!groups.has(canonical)) groups.set(canonical, []);
    groups.get(canonical).push(page);
  });
  return [...groups.entries()]
    .filter(([, pages]) => pages.length > 1)
    .map(([canonical, pages]) => ({
      canonical,
      // 最後更新的排前面，標題／摘要／心智圖優先採用較新的那份。
      pages: [...pages].sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0)),
    }));
};

// 實際把群組合併寫回 storage：標註合一、meta/摘要/心智圖收斂到主頁面，
// 再移除其餘舊 key。一次寫入，避免每筆都觸發面板重繪。
const mergeDuplicateGroups = async (groups) => {
  const stored = await chrome.storage.local.get([
    PAGE_META_KEY,
    GENERATED_NOTES_KEY,
    MINDMAP_KEY,
  ]);
  const meta = { ...(stored[PAGE_META_KEY] || {}) };
  const notes = { ...(stored[GENERATED_NOTES_KEY] || {}) };
  const maps = { ...(stored[MINDMAP_KEY] || {}) };
  const entryUpdates = {};
  const keysToRemove = [];
  let mergedExtra = 0;

  groups.forEach(({ canonical, pages }) => {
    const urls = pages.map((p) => p.url);
    entryUpdates[canonical] = mergeEntryArrays(pages.map((p) => p.entries));
    mergedExtra += pages.length - 1;

    // meta：title 取第一個非空（已依更新時間排序），tags 取聯集。
    const tagSet = new Set();
    let title;
    let note;
    let mindmap;
    pages.forEach((p) => {
      const m = meta[p.url];
      if (m?.title && !title) title = m.title;
      (Array.isArray(m?.tags) ? m.tags : []).forEach((tag) => tagSet.add(tag));
      if (note === undefined && notes[p.url] != null) note = notes[p.url];
      if (mindmap === undefined && maps[p.url] != null) mindmap = maps[p.url];
    });

    // 先清掉群組內所有舊 key 的附屬資料，再把合併結果寫到主頁面。
    urls.forEach((u) => {
      delete meta[u];
      delete notes[u];
      delete maps[u];
      if (u !== canonical) keysToRemove.push(u);
    });
    const newMeta = {};
    if (title) newMeta.title = title;
    if (tagSet.size) newMeta.tags = [...tagSet];
    if (Object.keys(newMeta).length) meta[canonical] = newMeta;
    if (note !== undefined) notes[canonical] = note;
    if (mindmap !== undefined) maps[canonical] = mindmap;
  });

  await chrome.storage.local.set({
    ...entryUpdates,
    [PAGE_META_KEY]: meta,
    [GENERATED_NOTES_KEY]: notes,
    [MINDMAP_KEY]: maps,
  });
  if (keysToRemove.length) await chrome.storage.local.remove(keysToRemove);
  return mergedExtra;
};

// 點「整理重複頁面」：先預覽要合併哪些群組，確認後才動手。
const runMergeDuplicates = async () => {
  const groups = findDuplicateGroups();
  if (!groups.length) {
    setStatus(t("manager.mergeNoneFound"));
    return;
  }
  const extra = groups.reduce((sum, g) => sum + g.pages.length - 1, 0);
  const samples = groups.slice(0, 8).map(({ canonical, pages }) => {
    const name = getPageDisplayName(canonical);
    return `• ${name}  (${pages.length} → 1)`;
  });
  const message = [
    t("manager.mergeConfirmIntro", { groups: groups.length, extra }),
    ...samples,
    groups.length > samples.length
      ? t("manager.mergeConfirmMore", { count: groups.length - samples.length })
      : null,
    "",
    t("manager.mergeConfirmHint"),
  ]
    .filter((line) => line !== null)
    .join("\n");
  const ok = await openConfirmDialog({
    title: t("manager.mergeConfirmTitle"),
    message,
    confirmLabel: t("manager.mergeConfirmBtn"),
    cancelLabel: t("manager.btnCancel"),
  });
  if (!ok) return;
  const mergedExtra = await mergeDuplicateGroups(groups);
  if (detailCurrentPageUrl) closePageDetail();
  state.selected.clear();
  await refreshManager();
  setStatus(
    t("manager.statusMerged", { groups: groups.length, extra: mergedExtra })
  );
};

// 跳到原文：寫入聚焦請求後開分頁，內容腳本載入時會捲動並閃爍該標註。
const openHighlightInPage = async (url, entryId) => {
  try {
    if (entryId) {
      await chrome.storage.local.set({
        hkFocusHighlight: { url, id: entryId, at: Date.now() },
      });
    }
  } catch (error) {
    console.debug("寫入聚焦請求失敗", error);
  }
  chrome.tabs.create({ url });
};

// 刪除單筆標註：優先用 id 比對，沒有 id 才退回索引。
const deleteHighlightEntry = async (url, entry, index) => {
  const stored = await chrome.storage.local.get(url);
  const entries = Array.isArray(stored[url]) ? stored[url] : [];
  const next = entry?.id
    ? entries.filter((item) => item?.id !== entry.id)
    : entries.filter((_, i) => i !== index);
  if (next.length === entries.length) return false;
  if (next.length === 0) {
    // 刪到最後一筆 → 整頁清掉（含 meta/摘要/心智圖），不留 0 筆空頁
    await deletePage(url);
  } else {
    await chrome.storage.local.set({ [url]: next });
  }
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
// 純解析邏輯共用自 parsers.js（manager.html 已先載入）。
const normalizeBulkPages = (pages) =>
  window.HkParsers ? window.HkParsers.normalizeBulkPages(pages) : [];

const parseGithubBackupPayload = (rawText) => {
  let parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (_error) {
    throw new Error(t("manager.errGithubBackupInvalid"));
  }
  const pages = Array.isArray(parsed?.pages)
    ? parsed.pages
    : Array.isArray(parsed)
    ? parsed
    : [];
  if (!pages.length) {
    throw new Error(t("manager.errGithubBackupEmpty"));
  }
  const normalized = normalizeBulkPages(pages);
  if (!normalized.length) {
    throw new Error(t("manager.errGithubBackupNoImport"));
  }
  return normalized;
};

const downloadAllPages = async () => {
  await fetchAllPages();
  if (!state.pages.length) {
    setStatus(t("manager.errNoNotesExport"), true);
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
  setStatus(t("manager.statusAllDownloaded"));
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
  cancelBtn.textContent = t("manager.btnCancel");

  const thirdBtn = document.createElement("button");
  thirdBtn.type = "button";
  thirdBtn.id = "hk-manager-confirm-third";
  thirdBtn.className = "hk-manager-btn hk-manager-btn-muted";
  thirdBtn.hidden = true;

  const confirmBtn = document.createElement("button");
  confirmBtn.type = "button";
  confirmBtn.id = "hk-manager-confirm-ok";
  confirmBtn.className = "hk-manager-btn";
  confirmBtn.textContent = t("manager.btnConfirm");

  actions.appendChild(cancelBtn);
  actions.appendChild(thirdBtn);
  actions.appendChild(confirmBtn);

  dialog.appendChild(title);
  dialog.appendChild(message);
  dialog.appendChild(actions);

  overlay.appendChild(backdrop);
  overlay.appendChild(dialog);
  document.body.appendChild(overlay);
  return overlay;
};

// 預設兩鈕回傳 boolean（向後相容）。傳入 thirdLabel 時變三鈕，
// 回傳 "confirm" / "third" / "cancel"。
const openConfirmDialog = ({
  title,
  message,
  confirmLabel,
  cancelLabel,
  thirdLabel,
}) =>
  new Promise((resolve) => {
    const overlay = ensureConfirmOverlay();
    const titleEl = overlay.querySelector("#hk-manager-confirm-title");
    const messageEl = overlay.querySelector("#hk-manager-confirm-message");
    const confirmBtn = overlay.querySelector("#hk-manager-confirm-ok");
    const cancelBtn = overlay.querySelector("#hk-manager-confirm-cancel");
    const thirdBtn = overlay.querySelector("#hk-manager-confirm-third");
    const backdrop = overlay.querySelector(".hk-manager-confirm-backdrop");

    if (!titleEl || !messageEl || !confirmBtn || !cancelBtn || !backdrop) {
      resolve(thirdLabel ? "cancel" : false);
      return;
    }

    const triState = Boolean(thirdLabel);
    titleEl.textContent = title || t("manager.dialogTitle");
    messageEl.textContent = message || "";
    confirmBtn.textContent = confirmLabel || t("manager.btnConfirm");
    cancelBtn.textContent = cancelLabel || t("manager.btnCancel");
    if (thirdBtn) {
      thirdBtn.hidden = !triState;
      if (triState) thirdBtn.textContent = thirdLabel;
    }

    let resolved = false;
    const cleanup = (result) => {
      if (resolved) return;
      resolved = true;
      overlay.classList.add("is-hidden");
      overlay.removeEventListener("keydown", onKeydown);
      confirmBtn.onclick = null;
      cancelBtn.onclick = null;
      if (thirdBtn) thirdBtn.onclick = null;
      backdrop.onclick = null;
      resolve(result);
    };

    const cancelValue = triState ? "cancel" : false;
    const onKeydown = (event) => {
      if (event.key === "Escape") cleanup(cancelValue);
    };

    confirmBtn.onclick = () => cleanup(triState ? "confirm" : true);
    cancelBtn.onclick = () => cleanup(cancelValue);
    if (thirdBtn) thirdBtn.onclick = () => cleanup("third");
    backdrop.onclick = () => cleanup(cancelValue);
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
  entriesTitle.textContent = t("manager.notesSection");
  const tagsRow = document.createElement("div");
  tagsRow.className = "hk-manager-detail-tags";
  const tagsLabel = document.createElement("span");
  tagsLabel.textContent = t("manager.tagsLabel");
  const tagsChips = document.createElement("div");
  tagsChips.id = "hk-manager-detail-tags-chips";
  tagsChips.className = "hk-manager-tags-chips";
  const tagsInput = document.createElement("input");
  tagsInput.id = "hk-manager-detail-tags-input";
  tagsInput.className = "hk-manager-input";
  tagsInput.placeholder = t("manager.tagsPlaceholder");
  const tagsButton = document.createElement("button");
  tagsButton.type = "button";
  tagsButton.className = "hk-manager-btn hk-manager-btn-muted";
  tagsButton.textContent = t("manager.applyTags");
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
  aiTitle.textContent = t("manager.aiRecord");
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
  mindmapTitle.textContent = t("manager.mindmapSection");
  const mindmapCopyBtn = document.createElement("button");
  mindmapCopyBtn.type = "button";
  mindmapCopyBtn.id = "hk-manager-detail-mindmap-copy";
  mindmapCopyBtn.className = "hk-manager-btn hk-manager-btn-muted";
  mindmapCopyBtn.textContent = t("manager.mindmapCopy");
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
  metaEl.textContent = t("manager.pageMetaDetail", {
    total: page.total,
    updated: page.updatedAt
      ? new Intl.DateTimeFormat(undefined, {
          dateStyle: "medium",
          timeStyle: "short",
        }).format(new Date(page.updatedAt))
      : t("manager.unknownTime"),
  });

  entriesList.innerHTML = "";
  if (!page.entries?.length) {
    const empty = document.createElement("p");
    empty.className = "hk-manager-detail-empty";
    empty.textContent = t("manager.emptyNotes");
    entriesList.appendChild(empty);
  } else {
    page.entries.forEach((entry, index) => {
      const item = document.createElement("article");
      item.className = "hk-manager-detail-item";

      const itemHead = document.createElement("div");
      itemHead.className = "hk-manager-detail-item-head";
      const text = document.createElement("p");
      text.className = "hk-manager-detail-text";
      text.textContent = entry.text || t("manager.emptyContent");
      itemHead.appendChild(text);
      const jumpBtn = document.createElement("button");
      jumpBtn.type = "button";
      jumpBtn.className = "hk-manager-detail-item-jump";
      jumpBtn.textContent = t("manager.jumpToSource");
      jumpBtn.title = t("manager.jumpToSourceTitle");
      jumpBtn.addEventListener("click", () =>
        openHighlightInPage(page.url, entry.id)
      );
      itemHead.appendChild(jumpBtn);
      const delEntryBtn = document.createElement("button");
      delEntryBtn.type = "button";
      delEntryBtn.className = "hk-manager-detail-item-delete";
      delEntryBtn.textContent = t("manager.deletePage");
      delEntryBtn.title = t("manager.deleteEntryTitle");
      delEntryBtn.addEventListener("click", async () => {
        const ok = await openConfirmDialog({
          title: t("manager.confirmDeleteEntryTitle"),
          message: t("manager.confirmDeleteEntry"),
          confirmLabel: t("manager.deletePage"),
          cancelLabel: t("manager.btnCancel"),
        });
        if (!ok) return;
        try {
          await deleteHighlightEntry(page.url, entry, index);
          await fetchAllPages();
          const updated = state.pages.find((p) => p.url === page.url);
          if (updated) renderPageDetail(updated);
          else closePageDetail();
          renderPageList();
          setStatus(t("ai.statusHighlightDeleted"));
        } catch (error) {
          console.debug("刪除標註失敗", error);
          setStatus(t("manager.statusDeleteFail"), true);
        }
      });
      itemHead.appendChild(delEntryBtn);
      item.appendChild(itemHead);

      const trimmedNote = entry.note?.trim();
      if (trimmedNote) {
        const note = document.createElement("p");
        note.className = "hk-manager-detail-note is-clickable";
        note.textContent = trimmedNote;
        note.setAttribute("title", t("manager.clickToCopyNote"));
        note.addEventListener("click", async () => {
          try {
            await navigator.clipboard.writeText(trimmedNote);
            setStatus(t("manager.statusNoteCopied"));
          } catch (error) {
            console.debug("複製註解失敗", error);
            setStatus(t("manager.errNoteCopy"), true);
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
    emptyTag.textContent = t("manager.emptyTags");
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
    aiContent.textContent = t("manager.emptyAiNote");
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
          setStatus(t("manager.statusMindmapCopied"));
        } catch (error) {
          console.debug("複製心智圖大綱失敗", error);
          setStatus(t("mindmap.copyFail"), true);
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
    setStatus(t("manager.statusTagsUpdated"));
  } catch (error) {
    console.debug("更新頁面標籤失敗", error);
    setStatus(t("manager.errTagsUpdate"), true);
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
  if (!settings.token) return t("manager.errNoToken");
  if (!settings.repo || !settings.repo.includes("/")) {
    return t("manager.errNoRepo");
  }
  if (!settings.path) return t("manager.errNoPath");
  return null;
};

const GITHUB_API_VERSION = "2022-11-28";
const githubHeaders = (token) => ({
  Authorization: `Bearer ${token}`,
  Accept: "application/vnd.github+json",
  "X-GitHub-Api-Version": GITHUB_API_VERSION,
});

// 把 GitHub 錯誤回應整理成精簡可行動的訊息。
// 只有「真的回 HTML 整頁」（邊緣在驗證前就擋下）才換成檢查清單；
// JSON 錯誤一律原樣顯示 message 與 errors 細節，方便看出是哪個欄位／權限出問題。
const describeGithubError = (status, errorText) => {
  const raw = (errorText || "").trim();
  if (/^<!?(?:doctype|html)\b/i.test(raw) || /^</.test(raw)) {
    return t("manager.errGithubBadRequest", { status: status || 400 });
  }
  try {
    const json = JSON.parse(raw);
    const parts = [];
    if (json?.message) parts.push(json.message);
    if (Array.isArray(json?.errors) && json.errors.length) {
      const detail = json.errors
        .map((e) =>
          e?.message || [e?.resource, e?.field, e?.code].filter(Boolean).join(" ")
        )
        .filter(Boolean)
        .join("; ");
      if (detail) parts.push(detail);
    }
    if (parts.length) {
      return `${parts.join(" — ")}${status ? `（${status}）` : ""}`;
    }
  } catch (_e) {
    /* 非 JSON，往下走 */
  }
  return raw.slice(0, 300) || `HTTP ${status}`;
};

// 統一的 GitHub REST 呼叫：自動帶版本標頭、JSON body，失敗時丟出整理過的訊息。
const githubApiRequest = async (
  repoBase,
  path,
  token,
  { method = "GET", body } = {}
) => {
  const response = await fetch(`${repoBase}${path}`, {
    method,
    headers: {
      ...githubHeaders(token),
      ...(body ? { "Content-Type": "application/json" } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!response.ok) {
    const text = await response.text();
    console.debug("GitHub 請求失敗", method, path, response.status, text.slice(0, 500));
    const error = new Error(describeGithubError(response.status, text));
    error.status = response.status;
    throw error;
  }
  return response.json();
};

// 經由 Git Data API 上傳：blob → tree → commit → 更新 ref。
// 關鍵差異：檔案路徑放在 JSON body（tree[].path），不進 URL，
// 避開 contents 端點對 URL 路徑編碼的邊緣限制（會回「invalid request」整頁 HTML），
// 也能處理較大的備份檔。
const uploadViaGitDataApi = async (settings, payloadText) => {
  const repoBase = buildRepoApiBase(settings.repo);
  if (!repoBase) throw new Error(t("manager.errRepoInvalid"));
  const { token, branch, path } = settings;
  const refPath = `/git/ref/heads/${encodeURIComponent(branch)}`;

  // 1. 分支最新 commit（不存在 → 之後建立新分支）
  let latestSha = null;
  try {
    const refData = await githubApiRequest(repoBase, refPath, token);
    latestSha = refData?.object?.sha || null;
  } catch (error) {
    if (error?.status !== 404) throw error; // 404＝分支還沒建立，其餘照丟
  }
  let baseTree = null;
  if (latestSha) {
    const commitData = await githubApiRequest(
      repoBase,
      `/git/commits/${latestSha}`,
      token
    );
    baseTree = commitData?.tree?.sha || null;
  }

  // 2. blob（base64 內容）
  const blob = await githubApiRequest(repoBase, "/git/blobs", token, {
    method: "POST",
    body: { content: encodeContentToBase64(payloadText), encoding: "base64" },
  });

  // 3. tree（檔案路徑只在這裡出現）
  const tree = await githubApiRequest(repoBase, "/git/trees", token, {
    method: "POST",
    body: {
      ...(baseTree ? { base_tree: baseTree } : {}),
      tree: [{ path, mode: "100644", type: "blob", sha: blob.sha }],
    },
  });

  // 4. commit
  const commit = await githubApiRequest(repoBase, "/git/commits", token, {
    method: "POST",
    body: {
      message: `backup: highlight-keeper (${new Date().toISOString()})`,
      tree: tree.sha,
      parents: latestSha ? [latestSha] : [],
    },
  });

  // 5. 更新既有分支，或建立新分支
  if (latestSha) {
    await githubApiRequest(repoBase, `/git/refs/heads/${encodeURIComponent(branch)}`, token, {
      method: "PATCH",
      body: { sha: commit.sha },
    });
  } else {
    await githubApiRequest(repoBase, "/git/refs", token, {
      method: "POST",
      body: { ref: `refs/heads/${branch}`, sha: commit.sha },
    });
  }
};

const fetchGithubBackupContent = async (settings) => {
  const repoBase = buildRepoApiBase(settings.repo);
  if (!repoBase) throw new Error(t("manager.errRepoInvalid"));
  const encodedPath = buildContentPath(settings.path);
  const url = `${repoBase}/contents/${encodedPath}?ref=${encodeURIComponent(
    settings.branch
  )}`;
  const response = await fetch(url, { headers: githubHeaders(settings.token) });
  if (response.status === 404) {
    throw new Error(t("manager.errGithubNotFound"));
  }
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(describeGithubError(response.status, errorText));
  }
  const json = await response.json();
  if (Array.isArray(json)) {
    throw new Error(t("manager.errGithubPathIsDir"));
  }
  if (json?.content) {
    return decodeBase64ToText(json.content);
  }
  if (json?.download_url) {
    const rawResponse = await fetch(json.download_url, {
      headers: githubHeaders(settings.token),
    });
    if (!rawResponse.ok) {
      const errorText = await rawResponse.text();
      throw new Error(describeGithubError(rawResponse.status, errorText));
    }
    return await rawResponse.text();
  }
  if (json?.sha) {
    const blobResponse = await fetch(`${repoBase}/git/blobs/${json.sha}`, {
      headers: githubHeaders(settings.token),
    });
    if (!blobResponse.ok) {
      const errorText = await blobResponse.text();
      throw new Error(t("manager.githubReadFail", { error: errorText }));
    }
    const blob = await blobResponse.json();
    if (!blob?.content) {
      throw new Error(t("manager.errGithubEmpty"));
    }
    return decodeBase64ToText(blob.content);
  }
  throw new Error(t("manager.errGithubEmpty"));
};

const uploadHighlightsToGithub = async () => {
  const settings = getGithubSettingsSnapshot();
  const validationError = validateGithubSettings(settings);
  if (validationError) {
    setGithubStatus(validationError, true);
    return;
  }
  setGithubActionsDisabled(true);
  setGithubStatus(t("manager.statusUploading"));
  try {
    await fetchAllPages();
    if (!state.pages.length) {
      throw new Error(t("manager.errNoNotesUpload"));
    }
    let payload = buildFullExportPayload();

    // 先看遠端現況：若遠端有本機沒有的頁面（多半是另一台裝置存的），
    // 盲目上傳會把它們洗掉 → 讓使用者選「合併／覆蓋／取消」。
    let remotePages = null;
    try {
      const remoteText = await fetchGithubBackupContent(settings);
      remotePages = parseGithubBackupPayload(remoteText);
    } catch (_e) {
      remotePages = null; // 遠端不存在或空 → 首次上傳，無衝突
    }
    if (remotePages && remotePages.length) {
      const localUrls = new Set(payload.pages.map((p) => p.url));
      const remoteOnly = remotePages.filter((p) => !localUrls.has(p.url));
      if (remoteOnly.length) {
        const samples = remoteOnly
          .slice(0, 3)
          .map((p) => `- ${getPageDisplayName(p.url)}`);
        const choice = await openConfirmDialog({
          title: t("manager.uploadConflictTitle"),
          message: [
            t("manager.uploadConflictIntro", { count: remoteOnly.length }),
            ...samples,
            remoteOnly.length > samples.length ? "…" : null,
            t("manager.uploadConflictAsk"),
          ]
            .filter(Boolean)
            .join("\n"),
          confirmLabel: t("manager.uploadConflictMerge"),
          thirdLabel: t("manager.uploadConflictOverwrite"),
          cancelLabel: t("manager.btnCancel"),
        });
        if (choice === "cancel") {
          setGithubStatus(t("manager.uploadCancelled"));
          return;
        }
        if (choice === "confirm") {
          // 合併：把遠端獨有頁面併進上傳內容，避免刪掉另一台裝置的資料
          payload = { ...payload, pages: [...payload.pages, ...remoteOnly] };
        }
        // choice === "third" → 直接覆蓋，payload 維持本機版本
      }
    }

    await uploadViaGitDataApi(settings, JSON.stringify(payload, null, 2));
    setGithubStatus(t("manager.statusUploaded"));
    githubSettings = { ...githubSettings, ...settings };
    persistGithubSettings(githubSettings);
  } catch (error) {
    setGithubStatus(error?.message || t("manager.errUpload"), true);
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
    setGithubStatus(t("manager.errGithubNoPages"), true);
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
      t("manager.importConflictIntro", { count: changed.length }),
      ...samples,
      changed.length > samples.length ? "..." : null,
      t("manager.importConflictAsk"),
      t("manager.importConflictHint"),
    ]
      .filter(Boolean)
      .join("\n");
    const shouldOverride = await openConfirmDialog({
      title: t("manager.confirmImportTitle"),
      message: confirmMessage,
      confirmLabel: t("manager.confirmImportOverwrite"),
      cancelLabel: t("manager.confirmImportNewOnly"),
    });
    if (!shouldOverride) {
      pagesToImport = pagesToImport.filter((page) => !changedUrls.includes(page.url));
      overwrittenCount = 0;
    }
  }
  if (!pagesToImport.length) {
    if (identical.length) {
      setGithubStatus(t("manager.statusAlreadyLatest"));
    } else {
      setGithubStatus(t("manager.statusImportCancelled"));
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
    t("manager.statusGithubImported", { count: importedCount }),
    overwrittenCount ? t("manager.statusOverwritten", { count: overwrittenCount }) : null,
    skippedOverlap > 0 ? t("manager.statusKeptLocal", { count: skippedOverlap }) : null,
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
  setGithubStatus(t("manager.statusDownloading"));
  try {
    const content = await fetchGithubBackupContent(settings);
    const pages = parseGithubBackupPayload(content);
    await applyGithubBackupPages(pages);
    githubSettings = { ...githubSettings, ...settings };
    persistGithubSettings(githubSettings);
  } catch (error) {
    setGithubStatus(error?.message || t("manager.errDownload"), true);
  } finally {
    setGithubActionsDisabled(false);
  }
};

const importMultipleFiles = async (files) => {
  if (!files?.length) return;
  setStatus(t("manager.statusParsing"));
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
    setStatus(t("manager.errNoImport"), true);
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
    setStatus(t("manager.errAllExist"), true);
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
    t("manager.statusImportPartial", { imported: toImport.length, skipped: skipped.length })
  );
  await refreshManager();
};

const initI18n = async () => {
  if (!HkI18n) return;
  await HkI18n.initI18n();
  HkI18n.applyDOMTranslations();
  if (langSelect) {
    langSelect.value = HkI18n.getLang();
    langSelect.addEventListener("change", (event) => {
      HkI18n.setLang(event.target.value);
    });
  }
  HkI18n.onLangChange(() => {
    HkI18n.applyDOMTranslations();
    if (langSelect) langSelect.value = HkI18n.getLang();
    renderPageList();
    if (detailCurrentPageUrl) {
      const page = state.pages.find((p) => p.url === detailCurrentPageUrl);
      if (page) renderPageDetail(page);
    }
  });
};

const init = async () => {
  await initI18n();
  refreshManager().catch((error) => {
    console.debug("載入筆記失敗", error);
    setStatus(t("manager.errLoadNotes"), true);
  });
  downloadBtn?.addEventListener("click", () => {
    downloadAllPages().catch((error) => {
      console.debug("下載全部筆記失敗", error);
      setStatus(error?.message || t("manager.errDownload"), true);
    });
  });
  importInput?.addEventListener("change", (event) => {
    importMultipleFiles(event.target.files).catch((error) => {
      console.debug(t("manager.errImport"), error);
      setStatus(error?.message || t("manager.errImport"), true);
    });
    event.target.value = "";
  });
  mergeBtn?.addEventListener("click", () => {
    runMergeDuplicates().catch((error) => {
      console.debug("整理重複頁面失敗", error);
      setStatus(error?.message || t("manager.statusDeleteFail"), true);
    });
  });
  closeBtn?.addEventListener("click", () => {
    window.close();
  });
  let searchDebounce = null;
  searchInput?.addEventListener("input", (event) => {
    state.searchTerm = event.target.value ?? "";
    if (searchDebounce) clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => renderPageList(), 150);
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
