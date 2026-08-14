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
    tags: ["標籤", "tags", "tag", "關鍵字", "keywords"],
    extra: [
      "延伸",
      "延伸補充",
      "補充",
      "延伸閱讀",
      "extra",
      "background",
      "further",
    ],
  };

  // 把 AI「標籤」區塊解析成乾淨的字串陣列：
  // 逗號／頓號／分號／換行／井字號都當分隔符，去掉項目符號與首尾標點，
  // 去重、限制每個標籤長度，最多取 max 個（避免一次塞爆）。
  const parsePageTags = (rawText, max = 8) => {
    if (typeof rawText !== "string" || !rawText.trim()) return [];
    const seen = new Set();
    const result = [];
    rawText
      .replace(/[,，、;；#＃]+/g, "\n")
      .split(/\n+/)
      .forEach((chunk) => {
        const tag = chunk
          .replace(/^[\s•\-*–—]+/, "")
          .replace(/^["'「『（(\[]+/, "")
          .replace(/["'」』）)\]\s.。]+$/, "")
          .trim();
        if (!tag || tag.length > 24) return;
        const key = tag.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        result.push(tag);
      });
    return result.slice(0, max);
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
    const defaultColor = deps.defaultColor || "#fff5b8";
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
      let extra = "";
      let source = "";
      for (const line of lines) {
        const asText = stripPrefix(line, [
          "原文", "原句", "片段", "source", "quote", "excerpt", "text",
        ]);
        if (asText !== null) { text = asText; continue; }
        // 「延伸／來源」都要先於「重點」比對：同一個區塊裡三者可能並存。
        // 注意不要把 "source" 放進來源別名 —— 它已經是「原文」的英文別名。
        const asSource = stripPrefix(line, [
          "來源", "出處", "參考來源", "參考", "ref", "reference", "citation",
        ]);
        if (asSource !== null) { source = asSource; continue; }
        const asExtra = stripPrefix(line, [
          "延伸", "延伸補充", "補充", "背景", "extra", "background", "context",
        ]);
        if (asExtra !== null) { extra = asExtra; continue; }
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
      // 「來源」只有在有延伸時才有意義（是延伸那句話的出處）。
      // 明確寫「無／none」的一律當成沒有，不要把佔位字當來源存下去。
      const cleanSource = /^(無|沒有|none|n\/a|-|—)$/i.test(source.trim())
        ? ""
        : source.trim();
      items.push({
        text,
        color: resolveColor(tag),
        reason,
        extra,
        source: extra ? cleanSource : "",
      });
    }
    return items;
  };

  // 解析「延伸」區塊：每組兩行「原文：…／延伸：…」，回傳 [{text, extra}]。
  // 之所以要求 AI 附上原文片段，是為了把補充掛回「那一句」，而不是丟一大段
  // 看不出在講哪裡的補充。
  const parseExtraBlocks = (rawText) => {
    if (typeof rawText !== "string" || !rawText.trim()) return [];
    const stripPrefix = (line, labels) => {
      for (const label of labels) {
        const re = new RegExp(`^[-*•·\\s]*${label}\\s*[:：]\\s*`, "i");
        if (re.test(line)) return line.replace(re, "").trim();
      }
      return null;
    };
    const items = [];
    rawText.split(/\n\s*\n/).forEach((block) => {
      const lines = block.split("\n").map((l) => l.trim()).filter(Boolean);
      let text = "";
      const extras = [];
      lines.forEach((line) => {
        const asText = stripPrefix(line, [
          "原文", "原句", "片段", "source", "quote", "excerpt", "text",
        ]);
        if (asText !== null) {
          text = asText;
          return;
        }
        const asExtra = stripPrefix(line, [
          "延伸", "補充", "延伸補充", "背景", "另一面", "extra", "background", "note",
        ]);
        if (asExtra !== null && asExtra) extras.push(asExtra);
      });
      const extra = extras.join(" ").trim();
      if (text && extra) items.push({ text, extra });
    });
    return items;
  };

  // 比對用的寬鬆正規化：去空白、標點與大小寫差異，讓 AI 回抄的片段
  // 即使少了句號或換行也對得上原本的標註。
  const looseKey = (input) =>
    typeof input === "string"
      ? input
          .toLowerCase()
          .replace(/\s+/g, "")
          .replace(/[.,;:!?"'`()\[\]{}<>«»—–\-。，、；：！？「」『』（）【】…·]/g, "")
      : "";

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

  // 把單頁匯出格式 {url,title,tags,entries,note,mindmap} 轉成可分享的 Markdown。
  // 區塊標題由呼叫端以 labels 傳入（i18n 由呼叫端處理，維持本檔無相依）。
  // labels: { tags, highlights, summary, mindmap }
  const pageToMarkdown = (page, labels = {}) => {
    if (!page || typeof page !== "object") return "";
    const L = {
      tags: labels.tags || "標籤",
      highlights: labels.highlights || "重點",
      summary: labels.summary || "摘要筆記",
      mindmap: labels.mindmap || "心智圖大綱",
    };
    const lines = [];
    const title = typeof page.title === "string" ? page.title.trim() : "";
    if (title) lines.push(`# ${title}`);
    if (page.url) lines.push(String(page.url));

    const tags = Array.isArray(page.tags)
      ? page.tags.map((t) => String(t).trim()).filter(Boolean)
      : [];
    if (tags.length) {
      lines.push(`${L.tags}：${tags.map((t) => `#${t}`).join(" ")}`);
    }

    const entries = Array.isArray(page.entries) ? page.entries : [];
    if (entries.length) {
      lines.push("", `## ${L.highlights}`);
      entries.forEach((entry) => {
        const text = (entry?.text || "").trim();
        if (!text) return;
        lines.push(`- ${text.replace(/\n+/g, " ")}`);
        const note = (entry?.note || "").trim();
        if (note) lines.push(`  > ${note.replace(/\n+/g, " ")}`);
      });
    }

    const summary =
      page.note && typeof page.note.note === "string"
        ? page.note.note.trim()
        : "";
    if (summary) lines.push("", `## ${L.summary}`, summary);

    const outline =
      page.mindmap && typeof page.mindmap.outline === "string"
        ? page.mindmap.outline.trim()
        : "";
    if (outline) lines.push("", `## ${L.mindmap}`, outline);

    return lines.join("\n").trim();
  };

  // 解析 pageToMarkdown 產出的分享 Markdown（「複製本頁筆記」的逆向）。
  // 中英區塊標題都認；不是這個格式回 null，讓呼叫端走其他解析路徑。
  const MD_SECTION_ALIASES = {
    highlights: ["重點", "highlights"],
    summary: ["摘要筆記", "summary"],
    mindmap: ["心智圖大綱", "mind map outline", "mindmap"],
  };
  const MD_TAGS_PREFIX = /^(標籤|tags)\s*[:：]\s*/i;

  const parseSharedMarkdown = (rawText) => {
    if (typeof rawText !== "string" || !rawText.trim()) return null;
    const lines = rawText.split(/\r?\n/);
    const aliasToKey = new Map();
    Object.entries(MD_SECTION_ALIASES).forEach(([key, names]) => {
      names.forEach((name) => aliasToKey.set(name.toLowerCase(), key));
    });

    const result = {
      url: "",
      title: "",
      tags: [],
      items: [],
      summary: "",
      outline: "",
    };
    let section = null; // null | highlights | summary | mindmap
    const summaryLines = [];
    const outlineLines = [];

    for (const rawLine of lines) {
      const line = rawLine.replace(/\s+$/, "");
      const heading = line.match(/^##\s+(.+)$/);
      if (heading) {
        const key = aliasToKey.get(heading[1].trim().toLowerCase());
        if (!key) return null; // 出現不認得的 ## 區塊 → 不是我們的格式
        section = key;
        continue;
      }
      if (!section) {
        // 檔頭區：# 標題、網址行、標籤行
        const titleMatch = line.match(/^#\s+(.+)$/);
        if (titleMatch && !result.title) {
          result.title = titleMatch[1].trim();
          continue;
        }
        if (!result.url && isHttpUrl(line.trim())) {
          result.url = line.trim();
          continue;
        }
        if (MD_TAGS_PREFIX.test(line)) {
          result.tags = line
            .replace(MD_TAGS_PREFIX, "")
            .split(/\s+/)
            .map((t) => t.replace(/^#/, "").trim())
            .filter(Boolean);
          continue;
        }
        continue;
      }
      if (section === "highlights") {
        const item = line.match(/^-\s+(.*)$/);
        if (item) {
          result.items.push({ text: item[1].trim(), note: "" });
          continue;
        }
        const note = line.match(/^\s+>\s*(.*)$/);
        if (note && result.items.length) {
          const last = result.items[result.items.length - 1];
          last.note = last.note ? `${last.note}\n${note[1]}` : note[1];
        }
        continue;
      }
      if (section === "summary") {
        summaryLines.push(line);
        continue;
      }
      if (section === "mindmap") {
        outlineLines.push(line);
      }
    }

    result.summary = summaryLines.join("\n").trim();
    result.outline = outlineLines.join("\n").trim();
    result.items = result.items.filter((item) => item.text);
    // 至少要有「重點」條列或摘要其中之一，且有標題或網址，才視為分享格式。
    if (!result.items.length && !result.summary) return null;
    if (!result.title && !result.url) return null;
    return result;
  };

  root.HkParsers = {
    isHttpUrl,
    AI_SECTION_ALIASES,
    splitAiSections,
    parsePageTags,
    parseMindmapOutline,
    looksLikeMindmapOutline,
    parseHighlightBlocks,
    parseExtraBlocks,
    looseKey,
    normalizeBulkPages,
    pageToMarkdown,
    parseSharedMarkdown,
  };
})(typeof window !== "undefined" ? window : globalThis);
