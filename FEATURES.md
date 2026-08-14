# Highlight Keeper — 功能總清單

盤點自 `manifest.json` / `popup.*` / `contentScript.js` / `manager.*` / `parsers.js` / `share-link.js` / `i18n.js`，作為 UI/UX 重做的依據。

---

## A. 標註核心（contentScript）

| # | 功能 | 說明 |
|---|---|---|
| A1 | 選取即標註 | 選取文字後浮出工具列，點 `HL` 用最近顏色標註 |
| A2 | 快捷鍵標註 | `Ctrl/Cmd+Shift+H` |
| A3 | 標註還原 | 以 canonical URL 正規化為 key，重開頁面自動復原 |
| A4 | 點擊標註開細節選單 | 換色、加註解、翻譯、刪除 |
| A5 | 顏色選擇 | 內建色票 + 自訂色碼，選單內即時挑色 |
| A6 | 選取翻譯 | 工具列翻譯鈕 → Google Translate 卡片，可把譯文寫回註解 |
| A7 | 失聯標註提示 | 原文變動找不到標註時，在面板常駐提示數量 |
| A8 | Toast 提示 | 存檔／匯入失敗等即時回饋 |

## B. 側邊面板（contentScript 內建，`Ctrl/Cmd+Shift+K` 開關）

| # | 功能 | 說明 |
|---|---|---|
| B1 | 可拖曳／左右側切換 | header 拖曳定位，記憶 `hkPanelSide` |
| B2 | 字級縮放 | A− / A+，記憶 `--hk-panel-font-scale` |
| B3 | AI 卡片（常駐） | 步驟 1 複製 Prompt → 貼到 ChatGPT → 步驟 2 貼回自動套用 |
| B4 | 直接產生 | 有 API Key 時一鍵呼叫 OpenAI / Gemini，不用複製貼上 |
| B5 | 心智圖 | 複製心智圖 Prompt、檢視已存心智圖、複製大綱 |
| B6 | 指引閱讀 | 依標註順序逐段引導閱讀 |
| B7 | 清除本頁標記 | 危險色按鈕 + 二次確認 |
| B8 | 分頁：本頁 | 本頁標註列表，點擊跳轉並閃爍定位，可單筆刪除 |
| B9 | 分頁：搜尋 | 關鍵字（標題／網址／標註／註解／標籤）+ 網址下拉 + 標籤 chips 篩選 |
| B10 | 分頁：摘要 | AI 筆記內容 + 複製 + 下方「跳到重點」清單 |
| B11 | 設定抽屜 | 顯示、本頁工具、本頁備份、全部備份、AI 服務 |
| B12 | 頁面標籤 | 新增／刪除標籤，含既有標籤建議下拉 |
| B13 | 本頁備份 | 複製純文字、下載本頁 JSON、匯入一份 JSON |
| B14 | 全部備份 | 下載全部 JSON、批次匯入多個 JSON（已存在頁面跳過） |
| B15 | AI 服務設定 | provider（OpenAI／Gemini／複製貼上）、model、API Key、Prompt 模板 |
| B16 | 標註分類 | 自訂分類名稱對應顏色，供 AI 畫重點使用；可重設 |
| B17 | 套用前預覽 | AI 回覆套用前先預覽數量，可確認／取消 |
| B18 | 逐段延伸 | 每個重點區塊多「延伸：」「來源：」兩行，直接存成該標註的註解；面板列表會把延伸分隔降階、來源做成 Google 搜尋連結。AI 設定可關；舊版獨立 `===延伸===` 區塊仍相容 |
| B19 | 舊標註重上色 | 設定 → 標註顏色 → 把舊的深色標註換淺，掃過所有頁面把舊飽和色換成新淺色 |

## C. Popup 快捷頁

| # | 功能 | 說明 |
|---|---|---|
| C1 | 本頁標記數 / AI 模式 | 直接讀 storage，不喚醒 content script |
| C2 | 在頁面開啟面板 | 必要時自動注入 content script |
| C3 | 複製 AI 畫重點 Prompt | 組出含全文＋色票＋標籤的 Prompt |
| C4 | 貼上並套用 | textarea 貼上即自動套用；補 macOS popup 缺少的 Cmd+A/C/X/V/Z |
| C5 | 複製分享連結 | 壓縮進 `原文網址#hk=…`；太長改走 GitHub raw |
| C6 | 複製本頁筆記 | 輸出 Markdown |
| C7 | 下載本頁 JSON | 匯出格式 `highlight-keeper-bulk` v2 |
| C8 | 清除本頁標記 | — |
| C9 | 色票管理 | 新增／編輯／刪除顏色，與面板即時同步 |
| C10 | 管理所有筆記 | 開 `manager.html` |
| C11 | 語言切換 | 繁中 / English |

## D. 筆記總覽 manager.html

| # | 功能 | 說明 |
|---|---|---|
| D1 | 全頁面清單 | 標題、網址、筆數、更新時間 |
| D2 | 關鍵字搜尋 | 標題／網址／標註／註解 |
| D3 | 排序 | 最後更新 / 筆記數 / 標題 |
| D4 | 網域篩選 + 標籤篩選 | 含「未加標籤」 |
| D5 | 多選批次操作 | 全選、匯出所選、刪除所選 |
| D6 | 下載全部 / 匯入多個 JSON | 衝突時可選「只加新的」或「覆蓋」 |
| D7 | 單頁操作 | 複製 Markdown、下載 JSON、複製分享連結、跳到原文、刪除 |
| D8 | AI 筆記 / 心智圖檢視 | 點擊複製筆記或心智圖大綱 |
| D9 | 標籤整理 | 改名、併入、刪除，一次套用到所有頁面 |
| D10 | 整理重複頁面 | 合併只差網址參數、或同站同標題的頁面 |
| D11 | GitHub 備份 | Token / repo / branch / path，上傳、下載、分類下載 |
| D12 | 匯出此分類 | 依目前篩選匯出 |
| D13 | 語言切換 | — |

## E. 分享／匯入（share-link.js）

| # | 功能 |
|---|---|
| E1 | 筆記壓縮進 URL fragment（`#hk=`）|
| E2 | 太長時 commit 到 GitHub 備份 repo，分享 raw 連結 |
| E3 | 對方開啟原文時偵測 fragment，跳出「匯入」提示（可合併／覆蓋／取消）|

## F. 資料格式（parsers.js / shared.js）

| # | 功能 |
|---|---|
| F1 | `normalizePageKey` 網址正規化（去追蹤參數、canonical 優先）|
| F2 | `pageToMarkdown` 轉 Markdown（標籤／重點／摘要／心智圖）|
| F3 | AI 回覆解析：重點 / 摘要 / 標籤 / 心智圖分段 |
| F4 | storage keys：`hkColorPalette`、`hkLastColor`、`hkAISettings`、`__hk_page_meta__`、`hkGeneratedNotes`、`hkMindmaps`、`hkGithubSyncSettings`、`hkPanelSide`、`hkLang` |

---

## 重做後的擺放位置

| 原本位置 | 新位置 |
|---|---|
| C3 複製 Prompt | **Popup 快捷區塊 ①** |
| C4 貼上並套用 | **Popup 快捷區塊 ②** |
| C5–C7 分享／複製／下載 | **Popup 快捷區塊 ③「分享此網頁」** |
| C2 開啟面板 | Popup 底部單一按鈕 |
| C9 色票管理 | 移到面板設定抽屜「標註顏色」 |
| C8 清除本頁 | 已在面板 AI 卡片下方（B7） |
| C10 / C11 | Popup 頁尾細字連結 |
| 其餘 | 維持在面板 / manager，僅換視覺 |
