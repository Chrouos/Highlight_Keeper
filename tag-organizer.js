(function (root) {
  const prompt = (tags) => `請協助整理以下文章標籤，找出意思相同或高度重複的標籤。只合併確實同義的標籤，不要合併上下位概念、不同產品或不同技術。優先保留常用且清楚的既有名稱。每個來源和目標都必須來自清單，不要刪除標籤，不要讓目標同時是來源。沒有建議時回傳空陣列。

只回傳 JSON，不加說明：
{"merges":[{"from":"原標籤","to":"保留的標籤"}]}

標籤與使用頁數：
${JSON.stringify(tags, null, 2)}`;
  const parse = (text, tags) => {
    const raw = String(text).trim().replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "");
    const data = JSON.parse(raw);
    if (!Array.isArray(data?.merges)) throw new Error("invalid");
    const allowed = new Set(tags.map((tag) => tag.name));
    const sources = new Set();
    const merges = data.merges.map((item) => {
      if (!item || typeof item.from !== "string" || typeof item.to !== "string" || !allowed.has(item.from) || !allowed.has(item.to) || item.from === item.to || sources.has(item.from)) throw new Error("invalid");
      sources.add(item.from);
      return { from: item.from, to: item.to };
    });
    if (merges.some((item) => sources.has(item.to))) throw new Error("invalid");
    return merges;
  };
  const normalize = (text) => String(text || "").normalize("NFKC").toLowerCase();
  const matches = (text, query) => normalize(query).trim().split(/\s+/).every((word) => normalize(text).includes(word));
  const snippet = (text, query) => {
    const source = String(text || "").replace(/\s+/g, " ");
    const word = normalize(query).trim().split(/\s+/).find(Boolean);
    const index = word ? normalize(source).indexOf(word) : 0;
    const start = Math.max(0, index - 55);
    return (start ? "…" : "") + source.slice(start, start + 180) + (source.length > start + 180 ? "…" : "");
  };
  const mergeTags = (meta, sources, target) => {
    const selected = new Set(sources);
    return Object.fromEntries(Object.entries(meta).map(([url, entry]) => {
      if (!Array.isArray(entry?.tags) || !entry.tags.some((tag) => selected.has(tag))) return [url, entry];
      return [url, { ...entry, tags: [...new Set(entry.tags.map((tag) => selected.has(tag) ? target.trim() : tag).filter(Boolean))], updatedAt: Date.now() }];
    }));
  };
  const addTags = (meta, urls, tags) => {
    const next = { ...meta };
    urls.forEach((url) => {
      next[url] = { ...meta[url], tags: [...new Set([...(meta[url]?.tags || []), ...tags])], updatedAt: Date.now() };
    });
    return next;
  };
  const matchesTags = (tags, selected) => {
    const values = Array.isArray(tags) ? tags : [];
    return !selected.length || selected.some((tag) => tag === "__hk_untagged__" ? !values.length : values.includes(tag));
  };
  const paginate = (items, requestedPage, size) => {
    const pages = Math.max(1, Math.ceil(items.length / size));
    const page = Math.max(1, Math.min(requestedPage, pages));
    const offset = (page - 1) * size;
    return { items: items.slice(offset, offset + size), page, pages, total: items.length, start: items.length ? offset + 1 : 0, end: Math.min(offset + size, items.length) };
  };
  root.HkTagOrganizer = { prompt, parse, matches, snippet, mergeTags, addTags, matchesTags, paginate };
})(globalThis);
