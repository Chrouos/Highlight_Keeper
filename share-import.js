/* 分享頁一鍵匯入：只在「本擴充自己的 viewer 分享頁」動作。
   偵測到分享筆記 → 頁面上方跳橫幅「要匯入你的 Highlight Keeper 嗎？」→
   一鍵寫進本機儲存；該頁已有筆記則問「覆蓋／合併／取消」。
   跑在 content script（isolated world），透過 chrome.storage 寫入與擴充其他頁共用的儲存。
   注入在 <all_urls>，非分享頁一開始就 return，成本極低。 */
(function () {
  // 只認我們自己的 viewer（靜態 HTML 內嵌 <meta name="hk-share">）＋ hash 參數。
  if (!document.querySelector('meta[name="hk-share"]')) return;
  const params = new URLSearchParams((location.hash || "").replace(/^#/, ""));
  const owner = params.get("o");
  const repo = params.get("r");
  const branch = params.get("b");
  const file = params.get("f");
  if (!owner || !repo || !branch || !file) return;
  if (!chrome?.storage?.local) return;

  // 告訴頁面「擴充在場」，讓 viewer 略過自動下載（改由這裡的一鍵匯入處理）。
  try {
    document.documentElement.setAttribute("data-hk-ext", "1");
  } catch (_e) {}

  const PAGE_META_KEY = "__hk_page_meta__";
  const GENERATED_NOTES_KEY = "hkGeneratedNotes";
  const MINDMAP_KEY = "hkMindmaps";

  const lang =
    (window.HkI18n && HkI18n.getLang && HkI18n.getLang()) ||
    navigator.language ||
    "en";
  const isZh = String(lang).toLowerCase().indexOf("zh") === 0;
  const T = {
    detected: isZh ? "偵測到分享筆記" : "Shared notes detected",
    importBtn: isZh ? "匯入到我的 Highlight Keeper" : "Import to my Highlight Keeper",
    marks: (n) => (isZh ? `${n} 筆標註` : `${n} highlight${n === 1 ? "" : "s"}`),
    importing: isZh ? "匯入中…" : "Importing…",
    done: (n) =>
      isZh ? `✓ 已匯入（本頁共 ${n} 筆）` : `✓ Imported (${n} total on this page)`,
    fail: isZh ? "匯入失敗，請稍後再試" : "Import failed, please try again",
    empty: isZh ? "沒有可匯入的筆記" : "Nothing to import",
    conflict: (n) =>
      isZh
        ? `這頁你已經有 ${n} 筆筆記，要怎麼處理？`
        : `You already have ${n} notes on this page. What to do?`,
    overwrite: isZh ? "覆蓋" : "Overwrite",
    merge: isZh ? "合併" : "Merge",
    cancel: isZh ? "取消" : "Cancel",
    cancelled: isZh ? "已取消匯入" : "Import cancelled",
    close: isZh ? "關閉" : "Close",
  };

  const rawUrl =
    "https://raw.githubusercontent.com/" +
    encodeURIComponent(owner) +
    "/" +
    encodeURIComponent(repo) +
    "/" +
    encodeURIComponent(branch) +
    "/" +
    file.split("/").map(encodeURIComponent).join("/");

  const entryKey = (e) =>
    e && e.id
      ? "id:" + e.id
      : "t:" + ((e && e.text) || "") + "|" + ((e && e.note) || "");

  const mergeEntries = (existing, incoming) => {
    const seen = new Set(existing.map(entryKey));
    const out = existing.slice();
    incoming.forEach((e) => {
      const k = entryKey(e);
      if (!seen.has(k)) {
        seen.add(k);
        out.push(e);
      }
    });
    return out;
  };

  // 把單頁寫進儲存。mode: 'overwrite' | 'merge'（新頁一律當 overwrite）。
  const importPage = async (page, mode) => {
    const url = page.url;
    const cur = await chrome.storage.local.get([
      url,
      PAGE_META_KEY,
      GENERATED_NOTES_KEY,
      MINDMAP_KEY,
    ]);
    const existingEntries = Array.isArray(cur[url]) ? cur[url] : [];
    const entries =
      mode === "merge"
        ? mergeEntries(existingEntries, page.entries)
        : page.entries.slice();

    const metaAll = { ...(cur[PAGE_META_KEY] || {}) };
    const metaEntry = { ...(metaAll[url] || {}) };
    if (mode === "merge") {
      if (!metaEntry.title && page.title) metaEntry.title = page.title;
      const tagSet = new Set([...(metaEntry.tags || []), ...(page.tags || [])]);
      if (tagSet.size) metaEntry.tags = [...tagSet];
    } else {
      if (page.title) metaEntry.title = page.title;
      if (page.tags && page.tags.length) metaEntry.tags = page.tags;
    }
    metaAll[url] = metaEntry;

    const notesAll = { ...(cur[GENERATED_NOTES_KEY] || {}) };
    if (mode === "merge") {
      if (!notesAll[url] && page.note) notesAll[url] = page.note;
    } else if (page.note) {
      notesAll[url] = page.note;
    }

    const mapsAll = { ...(cur[MINDMAP_KEY] || {}) };
    if (mode === "merge") {
      if (!mapsAll[url] && page.mindmap) mapsAll[url] = page.mindmap;
    } else if (page.mindmap) {
      mapsAll[url] = page.mindmap;
    }

    await chrome.storage.local.set({
      [url]: entries,
      [PAGE_META_KEY]: metaAll,
      [GENERATED_NOTES_KEY]: notesAll,
      [MINDMAP_KEY]: mapsAll,
    });
    return entries.length;
  };

  // ── 橫幅 UI（inline style，避開頁面 CSS 干擾） ──────────────────
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
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.setAttribute(
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
    return b;
  };

  const closeBtn = mkBtn("✕", false);
  closeBtn.title = T.close;
  closeBtn.style.padding = "7px 10px";
  closeBtn.addEventListener("click", () => bar.remove());

  bar.appendChild(msg);
  bar.appendChild(actions);
  bar.appendChild(closeBtn);

  const setActions = (buttons) => {
    actions.innerHTML = "";
    buttons.forEach((b) => actions.appendChild(b));
  };

  const finish = async (page, mode) => {
    msg.textContent = T.importing;
    setActions([]);
    try {
      const total = await importPage(page, mode);
      msg.textContent = T.done(total);
    } catch (error) {
      console.debug("匯入分享筆記失敗", error);
      msg.textContent = T.fail;
    }
  };

  const promptImport = async (page) => {
    const existing = await chrome.storage.local.get(page.url);
    const existingEntries = Array.isArray(existing[page.url])
      ? existing[page.url]
      : [];
    if (!existingEntries.length) {
      return finish(page, "overwrite"); // 新頁：直接寫入
    }
    // 已有筆記：問覆蓋／合併／取消
    msg.textContent = T.conflict(existingEntries.length);
    const ov = mkBtn(T.overwrite, true);
    const mg = mkBtn(T.merge, false);
    const cn = mkBtn(T.cancel, false);
    ov.addEventListener("click", () => finish(page, "overwrite"));
    mg.addEventListener("click", () => finish(page, "merge"));
    cn.addEventListener("click", () => {
      msg.textContent = T.cancelled;
      setActions([]);
    });
    setActions([ov, mg, cn]);
  };

  const showBanner = (page) => {
    const count = Array.isArray(page.entries) ? page.entries.length : 0;
    msg.textContent = `🔖 ${T.detected}：${page.title || ""}（${T.marks(count)}）`;
    const importBtn = mkBtn(T.importBtn, true);
    importBtn.addEventListener("click", () => promptImport(page));
    setActions([importBtn]);
    document.body.appendChild(bar);
  };

  fetch(rawUrl, { cache: "no-cache" })
    .then((res) => (res.ok ? res.text() : null))
    .then((text) => {
      if (text == null) return;
      let data;
      try {
        data = JSON.parse(text);
      } catch (_e) {
        return;
      }
      const rawPages = Array.isArray(data && data.pages)
        ? data.pages
        : Array.isArray(data)
        ? data
        : [];
      const pages =
        window.HkParsers && window.HkParsers.normalizeBulkPages
          ? window.HkParsers.normalizeBulkPages(rawPages)
          : [];
      if (!pages.length) return; // 沒有可匯入的頁（例如只有 AI 筆記無標註）
      showBanner(pages[0]);
    })
    .catch((error) => {
      console.debug("讀取分享筆記失敗", error);
    });
})();
