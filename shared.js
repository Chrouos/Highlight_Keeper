/* 共用工具：網址正規化（側欄、manager 與 contentScript 共用，避免各複製一份）。
   以一般 script 在各情境最先載入，掛在 window.HkUrlKey。 */
(function (root) {
  // 「決定內容」的真參數白名單：只有這些會留下，其餘一律剝掉，
  // 讓 path?ref=、path?source=、不同順序的追蹤參數都收斂到同一頁，
  // 但 youtube ?v=、文章 ?id=、分頁 ?page= 仍視為不同頁，不會誤併。
  const CONTENT_PARAMS = new Set([
    "v",            // YouTube 影片
    "id", "pid", "tid", "p", "t",   // 通用 / 論壇 id、文章、分頁
    "page", "page_id", "paged",     // 分頁
    "post", "postid", "story_fbid", "fbid", "thread", "threadid", "article", "articleid",
    "q", "query", "s", "search", "keyword",  // 搜尋關鍵字
    "hl", "lang", "locale",          // 語言
    "sku", "product_id", "productid", "asin",  // 商品頁
    "tab", "view", "mode",           // 內容分頁／檢視
  ]);

  const normalizePageKey = (href) => {
    try {
      const url = new URL(href);
      url.hash = "";
      // 只保留白名單參數，並依鍵名排序，讓參數順序不同也算同一頁。
      const kept = [];
      url.searchParams.forEach((value, key) => {
        if (CONTENT_PARAMS.has(key.toLowerCase())) kept.push([key, value]);
      });
      kept.sort((a, b) =>
        a[0] === b[0] ? a[1].localeCompare(b[1]) : a[0].localeCompare(b[0])
      );
      const next = new URLSearchParams();
      kept.forEach(([key, value]) => next.append(key, value));
      const query = next.toString();
      url.search = query ? `?${query}` : "";
      return url.href;
    } catch (_e) {
      return href;
    }
  };

  root.HkUrlKey = { normalizePageKey, CONTENT_PARAMS };
})(typeof window !== "undefined" ? window : globalThis);
