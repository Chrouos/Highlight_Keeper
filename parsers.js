/* 純解析函式（無 DOM／chrome 依賴），抽出供 contentScript／manager 共用並可單元測試。
   依賴（分類、顏色轉換、預設色、fallback 標題）一律由呼叫端以參數傳入。
   以一般 script 在各情境最先載入，掛在 window.HkParsers / globalThis.HkParsers。 */
(function (root) {
  const isHttpUrl = (key) => {
    if (typeof key !== "string") return false;
    try {
      const url = new URL(key);
      return url.protocol === "http:" || url.protocol === "https:";
    } catch (_e) {
      return false;
    }
  };

  // 區塊標記別名（中英雙語）：===重點===／===Highlights=== 等都認。
  const AI_SECTION_ALIASES = {
    highlights: ["重點", "畫重點", "標註", "highlights", "marks"],
    note: ["摘要", "筆記", "摘要筆記", "summary", "notes"],
    mindmap: ["心智圖", "mindmap"],
  };

  const splitAiSections = (rawText) => {
    if (typeof rawText !== "string" || !rawText.trim()) return null;
    const aliasToKey = new Map();
    Object.entries(AI_SECTION_ALIASES).forEach(([key, names]) => {
      names.forEach((name) => aliasToKey.set(name.toLowerCase(), key));
    });
    const headRe = /^\s*(?:[=＝—-]{2,}|[【\[]|#{2,4})\s*([^=＝【】\[\]#\s]{1,14})\s*(?:[=＝—-]{2,}|[】\]])?\s*$/;
    const sections = {};
    let currentKey = null;
    let buffer = [];
    let foundAny = false;
    const flush = () => {
      if (!currentKey) return;
      const body = buffer.join("\n").trim();
      if (body) sections[currentKey] = body;
    };
    for (const line of rawText.split(/\r?\n/)) {
      const match = line.match(headRe);
      const key = match ? aliasToKey.get(match[1].trim().toLowerCase()) : undefined;
      if (key) {
        flush();
        currentKey = key;
        buffer = [];
        foundAny = true;
        continue;
      }
      buffer.push(line);
    }
    flush();
    return foundAny ? sections : null;
  };

  const parseMindmapOutline = (rawText, fallbackTitle = "心智圖") => {
    if (typeof rawText !== "string" || !rawText.trim()) return null;
    const lines = rawText
      .split(/\r?\n/)
      .map((line) => line.replace(/\s+$/, ""))
      .filter((line) => line.trim());
    let title = "";
    const root = { label: "", children: [] };
    const stack = [{ node: root, depth: -1 }];
    const itemRe = /^(\s*)(?:[-*•·]|\d+[.)])\s+(.*)$/;

    for (const line of lines) {
      const headingMatch = line.match(/^#{1,3}\s+(.*)$/);
      if (headingMatch && !title) {
        title = headingMatch[1].trim();
        continue;
      }
      const itemMatch = line.match(itemRe);
      if (!itemMatch) {
        if (!title && stack.length === 1) title = line.trim();
        continue;
      }
      const indent = itemMatch[1].replace(/\t/g, "  ").length;
      const label = itemMatch[2].replace(/\*\*/g, "").trim();
      if (!label) continue;
      const node = { label, children: [] };
      while (stack.length > 1 && indent <= stack[stack.length - 1].depth) {
        stack.pop();
      }
      stack[stack.length - 1].node.children.push(node);
      stack.push({ node, depth: indent });
    }
    if (!root.children.length) return null;
    return { title: title || fallbackTitle, children: root.children };
  };

  // 大綱啟發式：多為清單行、且沒有「原文：／Source：」區塊。
  const looksLikeMindmapOutline = (text) => {
    if (typeof text !== "string") return false;
    if (/^(原文|source)\s*[:：]/im.test(text)) return false;
    const lines = text.split(/\r?\n/).filter((l) => l.trim());
    if (lines.length < 3) return false;
    const listLines = lines.filter((l) => /^\s*(?:[-*•·]|\d+[.)])\s+/.test(l));
    return listLines.length >= 3 && listLines.length >= lines.length * 0.6;
  };

  // 解析「原文：／#分類／重點：」三行式區塊（中英雙語前綴皆認）。
  // deps: { categories[], toHexColor(fn), defaultColor }
  const parseHighlightBlocks = (rawText, deps = {}) => {
    if (typeof rawText !== "string" || !rawText.trim()) return [];
    const toHexColor =
      typeof deps.toHexColor === "function" ? deps.toHexColor : (c) => c;
    const defaultColor = deps.defaultColor || "#ffeb3b";
    const cats = Array.isArray(deps.categories) ? deps.categories : [];
    const catByName = new Map();
    cats.forEach((c) => {
      if (c?.name) catByName.set(String(c.name).trim().toLowerCase(), toHexColor(c.color));
    });
    const fallbackColor = toHexColor(cats[0]?.color || defaultColor);
    const hexRe = /^#?[0-9a-fA-F]{6}$/;

    const resolveColor = (tag) => {
      const cleaned = (tag || "").trim().replace(/^#/, "");
      if (!cleaned) return fallbackColor;
      const byName = catByName.get(cleaned.toLowerCase());
      if (byName) return byName;
      if (hexRe.test(cleaned)) return toHexColor(`#${cleaned}`);
      return fallbackColor;
    };

    const stripPrefix = (line, labels) => {
      for (const label of labels) {
        const re = new RegExp(`^${label}\\s*[:：]\\s*`, "i");
        if (re.test(line)) return line.replace(re, "").trim();
      }
      return null;
    };

    const items = [];
    const blocks = rawText.split(/\n\s*\n/);
    for (const block of blocks) {
      const lines = block.split("\n").map((l) => l.trim()).filter(Boolean);
      let text = "";
      let tag = "";
      let reason = "";
      for (const line of lines) {
        const asText = stripPrefix(line, [
          "原文", "原句", "片段", "source", "quote", "excerpt", "text",
        ]);
        if (asText !== null) { text = asText; continue; }
        const asReason = stripPrefix(line, [
          "重點", "摘要", "說明", "理由", "point", "summary", "note", "reason", "insight",
        ]);
        if (asReason !== null) { reason = asReason; continue; }
        const tagMatch = line.match(/#\s*([^\s#:：]+)/);
        if (
          tagMatch &&
          (line.startsWith("#") ||
            /^(分類|顏色|標籤|category|color|colour|tag)\s*[:：]/i.test(line))
        ) {
          tag = tagMatch[1].trim();
        }
      }
      if (!text) continue;
      items.push({ text, color: resolveColor(tag), reason });
    }
    return items;
  };

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
        if (!url || !isHttpUrl(url)) return null;
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

  root.HkParsers = {
    isHttpUrl,
    AI_SECTION_ALIASES,
    splitAiSections,
    parseMindmapOutline,
    looksLikeMindmapOutline,
    parseHighlightBlocks,
    normalizeBulkPages,
  };
})(typeof window !== "undefined" ? window : globalThis);
