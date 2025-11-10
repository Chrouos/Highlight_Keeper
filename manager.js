const PAGE_META_KEY = "__hk_page_meta__";
const DEFAULT_COLOR = "#ffeb3b";

const state = {
  pages: [],
  meta: {},
  searchTerm: "",
};

const statusEl = document.getElementById("managerStatus");
const listEl = document.getElementById("pageList");
const pageCountEl = document.getElementById("pageCount");
const downloadBtn = document.getElementById("downloadAllBtn");
const importInput = document.getElementById("bulkImportInput");
const closeBtn = document.getElementById("closeManagerBtn");
const searchInput = document.getElementById("managerSearch");

const setStatus = (message, isError = false) => {
  if (!statusEl) return;
  statusEl.textContent = message || "";
  statusEl.classList.toggle("is-error", Boolean(isError));
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
        entries,
      };
    })
    .sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  state.pages = pages;
  state.meta = meta;
};

const matchesSearch = (page, term) => {
  if (!term) return true;
  const normalized = term.trim().toLowerCase();
  if (!normalized) return true;
  const haystacks = [
    page.title,
    page.url,
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

const downloadAllPages = async () => {
  await fetchAllPages();
  if (!state.pages.length) {
    setStatus("沒有筆記可匯出", true);
    return;
  }
  const payload = {
    type: "highlight-keeper-bulk",
    version: 1,
    exportedAt: Date.now(),
    pages: state.pages.map((page) => ({
      url: page.url,
      title: state.meta[page.url]?.title || "",
      entries: page.entries,
    })),
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
  setStatus("已下載全部筆記");
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
};

init();
