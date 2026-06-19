/* Service worker：把鍵盤快捷鍵轉發給目前分頁的內容腳本。
   內容腳本已在 <all_urls> 注入並處理 APPLY_HIGHLIGHT／TOGGLE_PAGE_PANEL。 */
const sendToActiveTab = async (message) => {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) return;
    await chrome.tabs.sendMessage(tab.id, message);
  } catch (_e) {
    // 內容腳本不在此頁（如 chrome:// 或安裝前就開著的分頁）→ 忽略
  }
};

chrome.commands.onCommand.addListener((command) => {
  if (command === "highlight-selection") {
    sendToActiveTab({ type: "APPLY_HIGHLIGHT" });
  } else if (command === "toggle-panel") {
    sendToActiveTab({ type: "TOGGLE_PAGE_PANEL" });
  }
});
