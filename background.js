/* Service worker：點擴充功能圖示一律開啟閱讀導覽側欄、轉發鍵盤快捷鍵、定時自動同步。 */
importScripts("shared.js", "parsers.js", "share-link.js", "auto-sync.js");

const cleanStoredData = async () => {
  const updates = HkParsers.cleanStoredData(await chrome.storage.local.get(null));
  if (Object.keys(updates).length) await chrome.storage.local.set(updates);
};

const onExtensionReady = () => {
  cleanStoredData().catch((error) => console.debug("清理舊資料失敗", error));
  chrome.sidePanel
    .setPanelBehavior({ openPanelOnActionClick: true })
    .catch((error) => console.debug("設定側欄行為失敗", error));
  HkAutoSync.applySchedule().catch((error) => console.debug("設定自動同步失敗", error));
  HkAutoSync.readState().then(HkAutoSync.updateBadge).catch(() => {});
};
chrome.runtime.onInstalled.addListener(onExtensionReady);
chrome.runtime.onStartup.addListener(onExtensionReady);

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === HkAutoSync.ALARM_NAME) HkAutoSync.runCycle();
});
chrome.storage.onChanged.addListener((changes, area) => {
  if (area === "local" && changes[HkAutoSync.SETTINGS_KEY]) {
    HkAutoSync.applySchedule().catch((error) => console.debug("設定自動同步失敗", error));
  }
});

// sidePanel.open 必須在使用者操作的同一個呼叫堆疊內執行，前面不能有 await。
const openSidePanel = (tabId) => chrome.sidePanel.open({ tabId });

chrome.commands.onCommand.addListener((command, tab) => {
  if (!tab?.id) return;
  if (command === "open-reading-guide" || command === "toggle-panel") {
    openSidePanel(tab.id).catch((error) => console.debug("開啟側欄失敗", error));
    return;
  }
  const type = {
    "highlight-selection": "APPLY_HIGHLIGHT",
  }[command];
  // 內容腳本不在此頁（如 chrome:// 或安裝前就開著的分頁）時忽略。
  if (type) chrome.tabs.sendMessage(tab.id, { type }).catch(() => {});
});

// 頁面工具面板的「開啟閱讀導覽」按鈕。
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  // 同步操作只接受擴充功能自己的頁面（筆記總覽），網頁上的內容腳本不能觸發。
  const fromExtensionPage = sender.url?.startsWith(chrome.runtime.getURL(""));
  if (fromExtensionPage && (message?.type === "HK_SYNC_NOW" || message?.type === "HK_SYNC_RESOLVE")) {
    const run = message.type === "HK_SYNC_NOW"
      ? HkAutoSync.runCycle({ manual: true })
      : HkAutoSync.resolveConflicts(message.choices || {});
    run.then(sendResponse).catch((error) => sendResponse({ status: "error", errorMessage: error?.message }));
    return true;
  }
  if (message?.type !== "OPEN_SIDE_PANEL" || !sender.tab?.id) return undefined;
  openSidePanel(sender.tab.id)
    .then(() => sendResponse({ success: true }))
    .catch((error) => sendResponse({ success: false, error: error?.message }));
  return true;
});
