/* 共用工具：網址正規化（popup 與 contentScript 共用，避免各複製一份）。
   以一般 script 在各情境最先載入，掛在 window.HkUrlKey。 */
(function (root) {
  // 只剝除已知追蹤／信件參數，保留 ?id=、?v= 等決定內容的真參數。
  const TRACKING_PARAMS = new Set([
    "fbclid", "gclid", "gclsrc", "dclid", "msclkid", "yclid", "ttclid",
    "twclid", "igshid", "mc_cid", "mc_eid", "_hsenc", "_hsmi", "spm",
    "scm", "vero_id", "oly_enc_id", "oly_anon_id", "_openstat",
  ]);

  const normalizePageKey = (href) => {
    try {
      const url = new URL(href);
      url.hash = "";
      const params = url.searchParams;
      const drop = [];
      params.forEach((_value, key) => {
        const lower = key.toLowerCase();
        if (lower.startsWith("utm_") || TRACKING_PARAMS.has(lower)) drop.push(key);
      });
      drop.forEach((key) => params.delete(key));
      const query = params.toString();
      url.search = query ? `?${query}` : "";
      return url.href;
    } catch (_e) {
      return href;
    }
  };

  root.HkUrlKey = { normalizePageKey, TRACKING_PARAMS };
})(typeof window !== "undefined" ? window : globalThis);
