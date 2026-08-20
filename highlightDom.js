(() => {
  const TRANSLATED_CLASS_NAMES = new Set(["translated-ltr", "translated-rtl"]);

  const hasClass = (element, className) => {
    if (!element) return false;
    if (element.classList?.contains?.(className)) return true;
    const rawClassName =
      typeof element.className === "string"
        ? element.className
        : element.className?.baseVal;
    return typeof rawClassName === "string"
      ? rawClassName.split(/\s+/).includes(className)
      : false;
  };

  const hasTranslatedClass = (element) =>
    [...TRANSLATED_CLASS_NAMES].some((name) => hasClass(element, name));

  const hasTranslatedFont = (element) => {
    try {
      return Boolean(
        element?.querySelector?.('font[style*="vertical-align"]')
      );
    } catch (_error) {
      return false;
    }
  };

  const isGoogleTranslatedDocument = (doc) => {
    const root = doc?.documentElement;
    const body = doc?.body;
    return Boolean(
      hasTranslatedClass(root) ||
        hasTranslatedClass(body) ||
        hasTranslatedFont(body) ||
        hasTranslatedFont(root)
    );
  };

  const isInsideTranslatedElement = (node) => {
    let current = node?.nodeType === 1 ? node : node?.parentElement;
    while (current) {
      if (hasTranslatedClass(current)) return true;
      if (current.tagName === "FONT") {
        const style = current.getAttribute?.("style") || "";
        if (/vertical-align\s*:\s*inherit/i.test(style)) return true;
      }
      current = current.parentElement;
    }
    return false;
  };

  const shouldUseTextNodeWrapping = (doc, range) => {
    if (isGoogleTranslatedDocument(doc)) return true;
    return [range?.startContainer, range?.endContainer].some(
      isInsideTranslatedElement
    );
  };

  const getHighlightMode = (doc, range) =>
    shouldUseTextNodeWrapping(doc, range) ? "css-highlight" : "inline";

  const api = {
    getHighlightMode,
    isGoogleTranslatedDocument,
    shouldUseTextNodeWrapping,
  };

  globalThis.HkHighlightDom = api;
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
})();
