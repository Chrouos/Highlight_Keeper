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
  root.HkTagOrganizer = { prompt, parse };
})(globalThis);
