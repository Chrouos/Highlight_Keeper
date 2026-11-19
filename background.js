/* Service worker：點擴充功能圖示開啟閱讀工作桌、轉發鍵盤快捷鍵、定時自動同步。 */
importScripts("shared.js", "parsers.js", "share-link.js", "auto-sync.js", "reading-desk-model.js", "reading-desk-store.js");

const cleanStoredData = async () => {
  const updates = HkParsers.cleanStoredData(await chrome.storage.local.get(null));
  if (Object.keys(updates).length) await chrome.storage.local.set(updates);
};

const onExtensionReady = () => {
  cleanStoredData().catch((error) => console.debug("清理舊資料失敗", error));
  chrome.sidePanel
    .setPanelBehavior({ openPanelOnActionClick: false })
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
  if (command === "toggle-panel") { openReadingDesk(tab); return; }
  if (command === "open-reading-guide") {
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
  if(message.settings)chrome.storage.session.set({hkOpenSettingsAt:Date.now()});
  openSidePanel(sender.tab.id)
    .then(() => sendResponse({ success: true }))
    .catch((error) => sendResponse({ success: false, error: error?.message }));
  return true;
});


const deskStore = HkReadingDeskStore.createStore(chrome.storage.local);
globalThis.HkReadingDeskWriteLock = task => deskStore.exclusive(task);
const openReadingDesk = async (tab) => {
  if (!tab?.id) return;
  if (!/^https?:/i.test(tab.url || "")) { openSidePanel(tab.id).catch(()=>{}); return; }
  try {
    const result = await chrome.tabs.sendMessage(tab.id, { type: "TOGGLE_READING_DESK" });
    if (!result?.success) throw new Error("Reload this page to use the reading desk");
  } catch (error) {
    // A browser-owned page or a tab opened before extension reload may lack content scripts.
    // Creating a library tab does not depend on a still-active user gesture.
    chrome.tabs.create({url:chrome.runtime.getURL("manager.html")}).catch(error => console.debug("開啟筆記總覽失敗", error));
  }
};
chrome.action.onClicked.addListener(openReadingDesk);
chrome.sidePanel.setPanelBehavior({openPanelOnActionClick:false}).catch(()=>{});
chrome.runtime.onMessage.addListener((message, sender, respond) => {
  if(message?.type==="HK_RESTORE_DELETED"&&sender.tab?.id){deskStore.restoreDeleted(message.pageKey,message.undo).then(()=>respond({success:true})).catch(error=>respond({success:false,error:error.message}));return true;}
  if (message?.type === "HK_DELETE_PAGE_HIGHLIGHTS" && sender.tab?.id) {
    deskStore.deleteHighlights(message.pageKey,message.ids).then(result=>respond({success:true,...result})).catch(error=>respond({success:false,error:error.message}));
    return true;
  }
  if (message?.type === "HK_DESK_OPEN_AI" && sender.tab?.id) {
    const url=message.provider==="gemini"?"https://gemini.google.com/app":"https://chatgpt.com/";
    chrome.storage.session.set({hkReadingGuideSourceTabId:sender.tab.id}).then(()=>chrome.tabs.create({url})).then(()=>respond({success:true})).catch(error=>respond({success:false,error:error.message}));
    return true;
  }
  if (message?.type === "HK_DESK_PROMPT_SOURCE" && sender.tab?.id) {
    chrome.storage.session.set({hkReadingGuideSourceTabId:sender.tab.id}).then(()=>respond({success:true})).catch(error=>respond({success:false,error:error.message}));
    return true;
  }
  if (message?.type === "HK_DESK_BATCH" && (sender.tab?.id || sender.url?.startsWith(chrome.runtime.getURL("")))) {
    deskStore.batch(message.operations).then(()=>respond({success:true})).catch(error=>respond({success:false,error:error.message}));
    return true;
  }
  if (message?.type === "HK_OPEN_MANAGER" && sender.tab?.id) {
    chrome.tabs.create({url:chrome.runtime.getURL("manager.html")}).then(()=>respond({success:true})).catch(error=>respond({success:false,error:error.message}));
    return true;
  }
  if (message?.type !== "HK_DESK_PATCH" || !sender.tab?.id) return;
  deskStore.patch(message.pageKey, message.action).then(desk=>respond({success:true,desk})).catch(error=>respond({success:false,error:error.message}));
  return true;
});
