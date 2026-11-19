# 全專案檢查（2026-11-18）

## 範圍與結果
檢查目前工作區的 Manifest V3 載入、GitHub 同步、備份解析與匯入匯出、資料刪除、SPA 導航、工作桌、旁註、側邊欄與設定。保留工作區原有修改，未提交、發布或使用真實 Token 進行遠端寫入。

修正 9 個可重現問題；新增或擴充回歸測試。同步／資料與 UI／導航另經獨立程式碼複查。

| 問題 | 修正 | 回歸證據 |
| --- | --- | --- |
| 遠端頁面資料損壞，被正規化忽略後視為刪除 | 同步前驗證整份備份、欄位及工作桌卡片；損壞時回報 remote-invalid，拒絕套用 | test/auto-sync.js、test/sync-remote-validation.js |
| 分類清理會刪除 categories 內的其他檔案 | 僅清理先前成功提交並記錄的分類檔；歷史不明檔案保留 | test/auto-sync.js：README.md、my-data.json 保留，舊管理分類仍正常移除 |
| 刪除整篇筆記時切換文章，可能刪到新文章 | 刪除目標在呼叫開始固定；儲存與移除均使用固定來源 | test/audit-spa-races.js |
| 舊文章還原較慢，將舊劃線插進新文章 | storage 讀取後檢查來源是否仍有效 | test/audit-spa-races.js |
| 舊版資料遷移中切換文章，資料可能搬錯位置 | 固定遷移目標與網址識別語意，完成後釋放遷移 Promise | test/audit-spa-races.js |
| 較早發起的側欄查詢晚到，覆蓋較新文章 | 查詢開始就配置序號，所有非同步階段驗證序號 | test/audit-spa-races.js |
| 舊文章還原結果更新新文章的標題或失聯提示 | 每個還原階段固定來源，過期結果不更新新頁狀態 | test/audit-spa-races.js |
| 刪除單一導讀項目時換頁，可能將舊資料寫入新頁 | 固定儲存來源，DOM／預覽只更新仍相同的文章 | test/audit-spa-races.js |
| 整篇刪除的二次確認可沿用到另一篇文章 | 確認狀態綁定文章；新文章需要重新確認 | test/audit-spa-races.js |

## 一致性整理
- Manifest 所有資源與 HTML 參照存在，HTML 無重複 ID。
- 中英文鍵完整，靜態引用無遺漏；加入 test/project-integrity.js 防止回歸。
- 移除 manager.js 中的實體 NUL 字元，保留相同分隔符語意。
- 隱私文件補齊 clipboardWrite、工作桌便條與卡片位置的儲存與備份說明。
- 最低 Chrome 版本調整為 123，對齊目前使用的 light-dark()。參考：[Chrome 123 release notes](https://developer.chrome.com/blog/new-in-chrome-123)。

## 驗證
- 根目錄全部 JS 的 node --check 通過。
- test/*.js 共 29 支測試腳本通過，含 7 個延遲／亂序的導航回歸案例。
- git diff --check 通過。
- 實際 Chrome headless DOM／排版 fixtures：
  - 工作桌：77 項
  - 頁邊註解：22 項
  - 側邊欄：25 項
  - 管理／側邊欄多寬度與語言排版：150 項
  - 選字選單視窗定位：通過
  - 按鈕標籤排版：36 項

重跑 Node 檢查：
```sh
for file in *.js; do node --check "$file" || exit 1; done
for file in test/*.js; do node "$file" || exit 1; done
git diff --check
```

## 驗證邊界
GitHub 使用記憶體內 API 模擬，已涵蓋雙裝置、衝突、刪除、失敗、讀取重試、提交讀取備援與復原；未驗證使用者實際 Token、代理、網路與 GitHub 儲存庫。瀏覽器 fixtures 使用真實 DOM/CSS 與模擬 extension API，不能取代載入實際擴充功能的操作驗收。Chrome 123 最低版本依官方功能支援對齊，未另外安裝歷史版本測試。

未知是否由程式建立的歷史分類檔會保留，避免誤刪其他儲存庫檔案。
