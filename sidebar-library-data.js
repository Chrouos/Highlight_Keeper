(function (root) {
  const pages = (data) => {
    const meta = data.__hk_page_meta__ || {};
    const notes = data.hkPersonalNotes || {};
    const generated = data.hkGeneratedNotes || {};
    const maps = data.hkMindmaps || {};
    const desks = data.hkReadingDesks || {};
    const urls = new Set([...Object.keys(meta), ...Object.keys(notes), ...Object.keys(generated), ...Object.keys(maps), ...Object.keys(desks), ...Object.keys(data).filter((key) => Array.isArray(data[key]))]);
    return [...urls].filter((url) => /^https?:\/\//i.test(url)).flatMap((url) => {
      const entries = Array.isArray(data[url]) ? data[url] : [];
      const tags = meta[url]?.tags || [];
      if (!entries.length && !tags.length && !notes[url] && !generated[url] && !maps[url] && !desks[url]?.cards?.length) return [];
      const title = meta[url]?.title || url;
      const updatedAt = Math.max(0, ...entries.map((e) => Number(e.createdAt) || 0), Number(notes[url]?.updatedAt) || 0, Number(generated[url]?.generatedAt) || 0, Number(maps[url]?.generatedAt) || 0, Number(desks[url]?.updatedAt) || 0);
      const text = [title, url, ...tags, ...entries.flatMap((e) => [e.text, e.note]), notes[url]?.text, generated[url]?.note, generated[url]?.context, ...(desks[url]?.cards || []).map(c=>c.text)].filter(Boolean).join('\n').normalize('NFKC').toLowerCase();
      return [{ url, title, tags, updatedAt, text }];
    }).sort((a, b) => b.updatedAt - a.updatedAt);
  };
  const search = (items, query) => {
    const words = String(query).normalize('NFKC').trim().toLowerCase().split(/\s+/);
    return items.filter((item) => words.every((word) => item.text.includes(word)));
  };
  const articleNotes = (highlights, desk) => {
    const items = (Array.isArray(desk?.cards) ? desk.cards : []).filter(card=>card.type==='note' && typeof card.text==='string' && card.text.trim()).map(card=>({id:card.id,type:'note',text:card.text}));
    const seen=new Set();
    (Array.isArray(highlights) ? highlights : []).forEach(entry=>{
      if(!entry?.id || seen.has(entry.id) || typeof entry.note!=='string' || !entry.note.trim())return;
      seen.add(entry.id);items.push({id:'quote:'+entry.id,type:'quote',highlightId:entry.id,quote:entry.text || '',text:entry.note});
    });
    return items;
  };
  root.HkSidebarLibrary = { pages, search, articleNotes };
})(globalThis);
