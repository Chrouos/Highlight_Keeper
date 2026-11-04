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

  // 從 AI 網頁的畫面複製回覆時，Markdown 符號常被加上反斜線（\=\=\=、\#\#、\*\*、\_），
  // 所有解析前先還原，貼上來源不同也能得到相同結果。
  const unescapeAiMarkdown = (text) =>
    String(text ?? "").replace(/\\([\\`*_{}[\]()#+\-.!=~>|:])/g, "$1");

  // 顯示用：長句一句一行，已存的舊筆記不用重跑 AI 也好讀。不改動儲存的內容。
  const BREAK_NOTE_MIN_LENGTH = 40;
  const breakNoteLines = (text) =>
    String(text ?? "")
      .split("\n")
      .map((line) =>
        Array.from(line).length < BREAK_NOTE_MIN_LENGTH
          ? line
          : line
              .replace(/([。！？]+[」』）)"”]*)\s*(?=\S)/g, "$1\n")
              .replace(/(?<!\b(?:e\.g|i\.e|vs|etc|Dr|Mr|Ms|Mrs|No))([.!?]["”')]?)\s+(?=[A-Z\u3400-\u9fff])/g, "$1\n")
      )
      .join("\n");

  // 區塊標記別名（中英雙語）：## 重點整理／===重點===／===Highlights=== 等都認。
  const AI_SECTION_ALIASES = {
    highlights: ["重點", "重點標註", "重點整理", "畫重點", "標註", "highlights", "marks"],
    context: ["脈絡", "主線", "論述脈絡", "context", "mainthread", "articlecontext"],
    terms: ["專有名詞", "名詞解釋", "名詞", "terms", "glossary"],
    note: ["摘要", "筆記", "摘要筆記", "summary", "notes"],
    mindmap: ["閱讀地圖", "導讀地圖", "心智圖", "心智圖大綱", "mindmap", "readingmap"],
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
    const headRe = /^\s*(?:[=＝—-]{2,}|[【\[]|#{1,6})\s*([^=＝【】\[\]#\s]{1,14})\s*(?:[=＝—-]{2,}|[】\]])?\s*$/;
    // 複製畫面文字時 ## 可能整個消失；只接受不會出現在內文的完整標題。
    const bareHeadings = new Set(["論述脈絡", "脈絡", "名詞解釋", "專有名詞", "重點整理", "重點標註", "閱讀地圖"]);
    const sections = {};
    let currentKey = null;
    let buffer = [];
    let foundAny = false;
    const flush = () => {
      if (!currentKey) return;
      const body = buffer.join("\n").trim();
      if (body) sections[currentKey] = body;
    };
    for (const line of unescapeAiMarkdown(rawText).split(/\r?\n/)) {
      if (/^\s*(?:```|~~~)/.test(line)) continue;
      const plainLine = line.replace(/\*\*|__/g, "");
      const match = plainLine.match(headRe);
      const bare = bareHeadings.has(plainLine.trim()) ? plainLine.trim() : "";
      const key = aliasToKey.get((match ? match[1].trim() : bare).toLowerCase());
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
  const normalizeAiFieldLine = (line) => String(line || "")
    .trim()
    .replace(/^[-*+•]\s+/, "")
    .replace(/^\d+[.)、]\s*/, "")
    .replace(/\*\*|__/g, "")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/^([^：\s]{1,8}：)\s+/, "$1");

  const cleanAiQuote = (value) => String(value || "").trim()
    .replace(/^(?:「([\s\S]*)」|“([\s\S]*)”|"([\s\S]*)")$/, (_all, a, b, c) => a ?? b ?? c)
    .trim();

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
    // A new quote starts a record. Blank lines inside an AI reply are formatting,
    // not record boundaries: ChatGPT often inserts one between every field.
    let text = "", tag = "", reason = "", extra = "", source = "", lastField = "";
    const flush = () => {
      if (!text) return;
      const cleanSource = /^(無|沒有|none|n\/a|-|—)$/i.test(source.trim()) ? "" : source.trim();
      items.push({ text, color: resolveColor(tag), reason, extra, source: extra ? cleanSource : "" });
    };
    const lines = unescapeAiMarkdown(rawText).split(/\r?\n/).map(normalizeAiFieldLine).filter(Boolean);
      for (const line of lines) {
        const asText = stripPrefix(line, [
          "原文", "原句", "片段", "source", "quote", "excerpt", "text",
        ]);
        if (asText !== null) {
          flush();
          text = cleanAiQuote(asText);
          tag = reason = extra = source = "";
          lastField = "";
          continue;
        }
        // 「延伸／來源」都要先於「重點」比對：同一個區塊裡三者可能並存。
        // 注意不要把 "source" 放進來源別名 —— 它已經是「原文」的英文別名。
        const asSource = stripPrefix(line, [
          "來源", "出處", "參考來源", "參考", "ref", "reference", "citation",
        ]);
        if (asSource !== null) { source = asSource; lastField = ""; continue; }
        const asExtra = stripPrefix(line, [
          "延伸", "延伸補充", "補充", "背景", "extra", "background", "context",
        ]);
        if (asExtra !== null) { extra = asExtra; lastField = ""; continue; }
        const asReason = stripPrefix(line, [
          "重點", "摘要", "說明", "理由", "備註", "註解", "解釋", "point", "summary", "note", "reason", "insight",
        ]);
        if (asReason !== null) { reason = reason && asReason ? `${reason}\n${asReason}` : reason || asReason; lastField = "reason"; continue; }
        const tagMatch = line.match(/#\s*([^\s#:：]+)/);
        if (
          tagMatch &&
          (line.startsWith("#") ||
            /^(分類|顏色|標籤|category|color|colour|tag)\s*[:：]/i.test(line))
        ) {
          tag = tagMatch[1].trim();
          lastField = "";
        } else if (lastField === "reason" && !/^(?:#|```|~~~)/.test(line)) {
          reason = reason ? `${reason}\n${line}` : line;
        }
      }
    flush();
    return items;
  };

  // 將新版閱讀指南與舊版「重點／摘要」格式統一成側欄可用資料。
  // 專有名詞區塊接受「名詞／解釋」格式，也接受既有「原文／重點」格式。
  const parseReadingGuideResponse = (rawText, deps = {}) => {
    const sections = splitAiSections(rawText) || {};
    const context = Array.from(String(sections.context || sections.note || "").trim())
      .slice(0, 1200)
      .join("");
    const termsText = sections.terms || "";
    const parsedTerms = [];
    let currentTerm = null;
    const saveCurrentTerm = () => {
      if (currentTerm?.text && currentTerm?.explanation) {
        parsedTerms.push({
          text: currentTerm.text,
          explanation: Array.from(currentTerm.explanation).slice(0, 600).join(""),
        });
      }
      currentTerm = null;
    };
    termsText.split(/\r?\n/).forEach((rawLine) => {
      const line = normalizeAiFieldLine(rawLine);
      const term = line.match(/^(?:名詞|專有名詞|詞條|term)\s*[:：]\s*(.+)$/i);
      const definition = line.match(/^(?:解釋|說明|定義|備註|註解|explanation|definition|note)\s*[:：]\s*(.*)$/i);
      if (term) {
        saveCurrentTerm();
        currentTerm = { text: cleanAiQuote(term[1]), explanation: "" };
      } else if (definition && currentTerm) {
        const value = definition[1].trim();
        if (value) currentTerm.explanation = currentTerm.explanation ? `${currentTerm.explanation}\n${value}` : value;
        currentTerm.hasDefinition = true;
      } else if (line && currentTerm?.hasDefinition && !/^\s*(?:\d+[.)]|#|```|~~~)/.test(rawLine)) {
        currentTerm.explanation = currentTerm.explanation ? `${currentTerm.explanation}\n${line}` : line;
      }
    });
    saveCurrentTerm();
    if (!parsedTerms.length && termsText) {
      parseHighlightBlocks(termsText, deps).forEach((item) => {
        if (item.text && item.reason) {
          parsedTerms.push({ text: item.text, explanation: Array.from(item.reason).slice(0, 600).join("") });
        }
      });
    }

    const highlights = parseHighlightBlocks(sections.highlights || "", deps)
      .map((item) => ({
        ...item,
        reason: String(item.reason || "").slice(0, 600),
        kind: "passage",
      }));
    const seen = new Set();
    const terms = parsedTerms
      .filter((item) => {
        const key = item.text.toLocaleLowerCase();
        if (!item.text || !item.explanation || seen.has(key)) return false;
        seen.add(key);
        return true;
      })
      .slice(0, 5)
      .map((item) => ({ text: item.text, explanation: item.explanation }));
    const sourceLine = unescapeAiMarkdown(rawText).split(/\r?\n/)
      .map(normalizeAiFieldLine)
      .find((line) => /^(?:來源文章|文章網址|source url)\s*[:：]/i.test(line));
    const sourceUrl = sourceLine?.match(/https?:\/\/[^\s<>)\]]+/i)?.[0] || "";
    return {
      context,
      terms,
      highlights: highlights.slice(0, 12),
      sourceUrl,
      mindmap: sections.mindmap || "",
      tags: sections.tags || "",
    };
  };

  // Count CJK characters and Latin words separately so English articles are not
  // treated as longer merely because their words use more characters.
  const getReadingGuideLimits = (pageText) => {
    const text = String(pageText || "");
    const cjk = (text.match(/[\u3400-\u9fff]/g) || []).length;
    const words = (text.replace(/[\u3400-\u9fff]/g, " ").match(/[a-z0-9]+(?:[-'][a-z0-9]+)*/gi) || []).length;
    const length = cjk + words * 2.5;
    if (length >= 4500) return { termsMax: 5, highlightsMin: 8, highlightsMax: 12 };
    if (length >= 1800) return { termsMax: 4, highlightsMin: 5, highlightsMax: 7 };
    return { termsMax: 3, highlightsMin: 3, highlightsMax: 4 };
  };

  const matchesReadingGuideSource = (guide, pageUrl, pageText) => {
    if (guide.sourceUrl) {
      const normalize = root.HkUrlKey?.normalizePageKey || ((url) => url);
      return normalize(guide.sourceUrl) === normalize(pageUrl);
    }
    // Legacy replies have no URL. Require at least one quote/term from this page
    // before any notes or highlights are saved to it.
    const normalizeText = (text) => String(text || "").replace(/\s+/g, " ").trim().toLowerCase();
    const article = normalizeText(pageText);
    return [...guide.terms, ...guide.highlights].some((item) => {
      const quote = normalizeText(item.text);
      return quote && article.includes(quote);
    });
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
          .replace(/[.,;:!?"'`()\[\]{}<>«»—–\-。，、；：！？「」『』（）【】…·“”‘’]/g, "")
      : "";

  // 英文字的詞形變化：testing / tested / tests 都算 test。只產生真的會出現的寫法，
  // 所以 rate 不會對到 rat，string 也不會被拆成 str。
  const englishWordForms = (word) => {
    const lower = word.toLowerCase();
    if (lower.length < 3) return [lower];
    const roots = new Set([lower]);
    const addStem = (stem) => {
      if (stem.length < 3 || !/[aeiouy]/.test(stem)) return;
      roots.add(stem);
      roots.add(`${stem}e`);
      if (/([b-df-hj-np-tv-z])\1$/.test(stem)) roots.add(stem.slice(0, -1));
    };
    if (lower.endsWith("ies") && lower.length >= 5) roots.add(`${lower.slice(0, -3)}y`);
    else if (lower.endsWith("es") && /(s|x|z|ch|sh)es$/.test(lower)) roots.add(lower.slice(0, -2));
    else if (lower.endsWith("s") && !/(ss|us|is)$/.test(lower)) roots.add(lower.slice(0, -1));
    if (lower.endsWith("ing")) addStem(lower.slice(0, -3));
    if (lower.endsWith("ed")) addStem(lower.slice(0, -2));
    const forms = new Set();
    roots.forEach((item) => {
      forms.add(item);
      if (item.endsWith("e")) {
        forms.add(`${item}s`);
        forms.add(`${item}d`);
        forms.add(`${item.slice(0, -1)}ing`);
      } else if (/[^aeiou]y$/.test(item)) {
        forms.add(`${item.slice(0, -1)}ies`);
      } else {
        forms.add(/(s|x|z|ch|sh)$/.test(item) ? `${item}es` : `${item}s`);
        forms.add(`${item}ed`);
        forms.add(`${item}ing`);
        const last = item.slice(-1);
        forms.add(`${item}${last}ed`);
        forms.add(`${item}${last}ing`);
      }
    });
    return [...forms].sort((a, b) => b.length - a.length);
  };

  // 比對完整單字的正規式，最後一個英文字（中心詞）可換成指定詞形；名詞不是英文結尾時回傳 null。
  const buildTermPattern = (term, toForms) => {
    const tokens = String(term || "").trim().split(/\s+/).filter(Boolean);
    const lastWord = tokens.length ? tokens[tokens.length - 1].match(/^(.*?)([A-Za-z]+)$/) : null;
    if (!lastWord) return null;
    const escape = (value) => value.replace(/[.*+?^${}()|[\]\\/]/g, "\\$&");
    const lead = [...tokens.slice(0, -1).map(escape), escape(lastWord[1])];
    const prefix = `${lead.slice(0, -1).join("[\\s-]+")}${lead.length > 1 ? "[\\s-]+" : ""}${lead[lead.length - 1]}`;
    const forms = toForms(lastWord[2]).map(escape).join("|");
    return new RegExp(`(?<![\\p{L}\\p{N}])${prefix}(?:${forms})(?![\\p{L}\\p{N}])`, "iu");
  };

  // 容許詞形不同，回傳文章裡實際的寫法；對不到回傳 null。
  const findTermVariant = (term, pageText) =>
    String(pageText || "").match(buildTermPattern(term, englishWordForms) || /(?!)/)?.[0] || null;

  // AI 常把名詞寫成「Model ID（模型識別碼）」。標記要用文章裡真的出現的那一邊。
  const pickTermText = (text, pageText) => {
    const raw = String(text || "").trim();
    const parts = raw.match(/^(.+?)\s*[（(]\s*(.+?)\s*[）)]$/);
    const candidates = parts ? [parts[1], parts[2]] : [raw];
    const page = looseKey(String(pageText || ""));
    // 英文名詞要是完整單字：test 不算出現在 latest 裡。
    const isOnPage = (item) => {
      if (!looseKey(item) || !page.includes(looseKey(item))) return false;
      const exact = buildTermPattern(item, (word) => [word]);
      return exact ? exact.test(String(pageText || "")) : true;
    };
    const found = [raw, ...candidates].find(isOnPage);
    if (found) return found;
    // 完全相同的寫法優先；都沒有才看詞形變化（A/B Testing → A/B test）。
    for (const item of candidates) {
      const variant = findTermVariant(item, pageText);
      if (variant) return variant;
    }
    return candidates[0];
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
        const note = page.note && typeof page.note === "object" ? page.note : null;
        const mindmap =
          page.mindmap && typeof page.mindmap === "object" ? page.mindmap : null;
        const hasPersonalNote = typeof page.personalNote === "string";
        const personalNote = hasPersonalNote ? page.personalNote : "";
        if (!entries.length && !note && !mindmap && !personalNote) return null;
        return {
          url,
          title: typeof page.title === "string" ? page.title : "",
          tags: Array.isArray(page.tags)
            ? page.tags.map((t) => String(t).trim()).filter(Boolean)
            : [],
          entries,
          note,
          mindmap,
          ...(hasPersonalNote ? { personalNote } : {}),
        };
      })
      .filter(Boolean);
  };

  const STORAGE_KEYS = {
    meta: "__hk_page_meta__",
    notes: "hkGeneratedNotes",
    mindmaps: "hkMindmaps",
    personal: "hkPersonalNotes",
  };

  // AI 筆記裡的 prompt 是整篇文章全文，不需要備份，也不算內容變動。
  const exportNote = (note) => {
    if (!note || typeof note !== "object") return null;
    const { prompt: _prompt, ...rest } = note;
    return rest;
  };

  // 這一頁在本機有沒有真正的內容（只存了標題不算）。
  const pageHasContent = (all, url) =>
    (Array.isArray(all[url]) && all[url].length > 0) ||
    Boolean(all[STORAGE_KEYS.notes]?.[url]) ||
    Boolean(all[STORAGE_KEYS.mindmaps]?.[url]) ||
    typeof all[STORAGE_KEYS.personal]?.[url]?.text === "string";

  // 從 storage 全部內容組出備份頁面；下載全部與自動同步都用這一份。
  const collectBackupPages = (all) => {
    const meta = all[STORAGE_KEYS.meta] || {};
    const notes = all[STORAGE_KEYS.notes] || {};
    const mindmaps = all[STORAGE_KEYS.mindmaps] || {};
    const personal = all[STORAGE_KEYS.personal] || {};
    const urls = new Set(
      [
        ...Object.keys(all).filter((key) => Array.isArray(all[key])),
        ...Object.keys(meta),
        ...Object.keys(notes),
        ...Object.keys(mindmaps),
        ...Object.keys(personal),
      ].filter(isHttpUrl)
    );
    return normalizeBulkPages(
      [...urls].map((url) => ({
        url,
        title: meta[url]?.title || "",
        tags: Array.isArray(meta[url]?.tags) ? meta[url].tags : [],
        entries: Array.isArray(all[url]) ? all[url] : [],
        note: exportNote(notes[url]),
        mindmap: mindmaps[url] || null,
        ...(typeof personal[url]?.text === "string" ? { personalNote: personal[url].text } : {}),
      }))
    );
  };

  // 把備份頁面轉成要寫入 storage 的內容。本機已有內容的頁面跳過，不覆蓋。
  const pagesToStorageUpdates = (pages, all, now = Date.now()) => {
    const meta = { ...(all[STORAGE_KEYS.meta] || {}) };
    const notes = { ...(all[STORAGE_KEYS.notes] || {}) };
    const mindmaps = { ...(all[STORAGE_KEYS.mindmaps] || {}) };
    const personal = { ...(all[STORAGE_KEYS.personal] || {}) };
    const updates = {};
    const skipped = [];
    const imported = [];
    pages.forEach((page) => {
      if (pageHasContent(all, page.url)) {
        skipped.push(page.url);
        return;
      }
      imported.push(page.url);
      if (page.entries.length) updates[page.url] = page.entries;
      if (page.title || page.tags.length) {
        meta[page.url] = { ...meta[page.url], title: page.title, tags: page.tags, updatedAt: now };
      }
      if (page.note) notes[page.url] = exportNote(page.note);
      if (page.mindmap) mindmaps[page.url] = page.mindmap;
      if (typeof page.personalNote === "string") personal[page.url] = { text: page.personalNote, updatedAt: now };
    });
    if (imported.length) {
      updates[STORAGE_KEYS.meta] = meta;
      updates[STORAGE_KEYS.notes] = notes;
      updates[STORAGE_KEYS.mindmaps] = mindmaps;
      updates[STORAGE_KEYS.personal] = personal;
    }
    return { updates, imported, skipped };
  };

  // 清掉舊版留下的多餘資料：只逛過的頁面標題、AI 筆記裡的文章全文。回傳要寫入的部分。
  const cleanStoredData = (all) => {
    const updates = {};
    const meta = all[STORAGE_KEYS.meta];
    if (meta) {
      const kept = Object.fromEntries(
        Object.entries(meta).filter(([url, value]) => value?.tags?.length || pageHasContent(all, url))
      );
      if (Object.keys(kept).length !== Object.keys(meta).length) updates[STORAGE_KEYS.meta] = kept;
    }
    const notes = all[STORAGE_KEYS.notes];
    if (notes && Object.values(notes).some((note) => note && "prompt" in note)) {
      updates[STORAGE_KEYS.notes] = Object.fromEntries(
        Object.entries(notes).map(([url, note]) => [url, exportNote(note) || note])
      );
    }
    return updates;
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
      context: labels.context || "脈絡",
      personalNote: labels.personalNote || "我的筆記",
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
        const kindPrefix = entry?.kind === "term" ? "〔專有名詞〕 " : "";
        lines.push(`- ${kindPrefix}${text.replace(/\n+/g, " ")}`);
        const note = (entry?.note || "").trim();
        if (note) note.split(/\n+/).forEach((noteLine) => lines.push(`  > ${noteLine}`));
      });
    }

    const summary =
      page.note && typeof page.note.note === "string"
        ? page.note.note.trim()
        : "";
    if (summary) lines.push("", `## ${L.summary}`, summary);

    const context =
      page.note && typeof page.note.context === "string"
        ? page.note.context.trim()
        : "";
    if (context) lines.push("", `## ${L.context}`, context);

    const personalNote =
      typeof page.personalNote === "string" ? page.personalNote.trim() : "";
    if (personalNote) {
      lines.push(
        "",
        `## ${L.personalNote}`,
        ...personalNote.split("\n").map((line) => `> ${line}`)
      );
    }

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
    context: ["脈絡", "文章脈絡", "context"],
    personalNote: ["我的筆記", "個人筆記", "personal notes", "my notes"],
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
      context: "",
      personalNote: "",
      outline: "",
    };
    let section = null; // null | highlights | summary | mindmap
    const summaryLines = [];
    const contextLines = [];
    const personalNoteLines = [];
    let hasPersonalNoteSection = false;
    const outlineLines = [];

    for (const rawLine of lines) {
      const line = rawLine.replace(/\s+$/, "");
      const heading = line.match(/^##\s+(.+)$/);
      if (heading) {
        const key = aliasToKey.get(heading[1].trim().toLowerCase());
        if (!key && section === "personalNote") {
          personalNoteLines.push(line);
          continue;
        }
        if (!key) return null; // 出現不認得的 ## 區塊 → 不是我們的格式
        section = key;
        if (key === "personalNote") hasPersonalNoteSection = true;
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
          const term = item[1].match(/^(?:〔專有名詞〕\s*|\[term\]\s*)/i);
          result.items.push({
            text: item[1].replace(/^(?:〔專有名詞〕\s*|\[term\]\s*)/i, "").trim(),
            note: "",
            ...(term ? { kind: "term" } : {}),
          });
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
      if (section === "context") {
        contextLines.push(line);
        continue;
      }
      if (section === "personalNote") {
        personalNoteLines.push(line.replace(/^\s*> ?/, ""));
        continue;
      }
      if (section === "mindmap") {
        outlineLines.push(line);
      }
    }

    result.summary = summaryLines.join("\n").trim();
    result.context = contextLines.join("\n").trim();
    if (hasPersonalNoteSection) {
      result.personalNote = personalNoteLines.join("\n").trim();
    } else {
      delete result.personalNote;
    }
    result.outline = outlineLines.join("\n").trim();
    result.items = result.items.filter((item) => item.text);
    // 至少要有「重點」條列或摘要其中之一，且有標題或網址，才視為分享格式。
    if (!result.items.length && !result.summary && !result.context && !result.personalNote && !hasPersonalNoteSection) return null;
    if (!result.title && !result.url) return null;
    return result;
  };

  root.HkParsers = {
    isHttpUrl,
    unescapeAiMarkdown,
    breakNoteLines,
    AI_SECTION_ALIASES,
    splitAiSections,
    parsePageTags,
    parseMindmapOutline,
    looksLikeMindmapOutline,
    parseHighlightBlocks,
    parseReadingGuideResponse,
    getReadingGuideLimits,
    matchesReadingGuideSource,
    parseExtraBlocks,
    looseKey,
    pickTermText,
    normalizeBulkPages,
    exportNote,
    collectBackupPages,
    pagesToStorageUpdates,
    cleanStoredData,
    pageToMarkdown,
    parseSharedMarkdown,
  };
})(typeof window !== "undefined" ? window : globalThis);
