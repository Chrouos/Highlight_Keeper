/* Highlight Keeper i18n — runtime-switchable Traditional Chinese / English. */
(function (root) {
  const STORAGE_KEY = "hkLang";
  const SUPPORTED = ["zh", "en"];
  const DEFAULT_FALLBACK = "en";

  const MESSAGES = {
    zh: {
      // ── popup ────────────────────────────────────────────
      "popup.title": "頁面標註",
      "popup.colorLabel": "標註顏色",
      "popup.addToPalette": "加入到色票",
      "popup.availableColors": "可用顏色",
      "popup.paletteAria": "顏色色票",
      "popup.panelControlAria": "頁面標註面板",
      "popup.aiToolsAria": "AI 工具",
      "popup.dataToolsAria": "資料管理",
      "popup.openPanel": "在頁面開啟面板",
      "popup.aiAutoHighlight": "複製 AI 畫重點 Prompt",
      "popup.clearPage": "清除本頁標記",
      "popup.openManager": "管理所有筆記",
      "popup.contact": "若有任何問題，私信到 chrouodiu@gmail.com",
      "popup.langLabel": "顯示語言",
      "popup.langZh": "繁體中文",
      "popup.langEn": "English",
      "popup.emptyPalette": "目前沒有顏色，請新增一個色票。",
      "popup.useColor": "使用 {color}",
      "popup.editColor": "編輯顏色 {color}",
      "popup.deleteColor": "刪除",
      "popup.deleteColorTitle": "刪除顏色 {color}",
      "popup.statusSelected": "已選擇顏色 {color}",
      "popup.statusUpdated": "色票顏色已更新為 {color}",
      "popup.statusDeleted": "已刪除顏色 {color}",
      "popup.statusDuplicate": "此顏色已存在於色票中",
      "popup.statusAdded": "已新增顏色 {color}",
      "popup.errNoTab": "找不到目前的分頁，無法開啟面板。",
      "popup.errOpenPanel": "無法開啟頁面面板",
      "popup.statusPanelOpened": "已在頁面顯示面板",
      "popup.errOpenPanel2": "無法開啟頁面面板。",
      "popup.errNoTabSimple": "找不到分頁",
      "popup.errCannotTrigger": "無法觸發",
      "popup.statusAiTriggered": "已觸發 AI 自動畫重點",
      "popup.statusAiCopied": "已複製畫重點 Prompt，貼到 ChatGPT 後把回覆貼回頁面面板",
      "popup.errAiTrigger": "無法建立畫重點 Prompt",
      "popup.pageInfoAria": "頁面資訊",
      "popup.markCountLabel": "本頁標記",
      "popup.aiModeLabel": "AI 模式",
      "popup.aiModeOpenai": "OpenAI API",
      "popup.aiModeGemini": "Gemini API",
      "popup.aiModeCopy": "複製貼上（ChatGPT）",
      "popup.errCannotClear": "無法清除",
      "popup.statusCleared": "已清除本頁所有標記",
      "popup.errClear": "無法清除本頁標記",
      "popup.errCannotSendMessage": "無法傳送訊息",

      // ── manager ──────────────────────────────────────────
      "manager.pageTitle": "Highlight Keeper — 笔記總覽",
      "manager.heading": "筆記內容總覽",
      "manager.subheading": "一次匯出/匯入所有頁面的標註，集中管理你的資料。",
      "manager.close": "×",
      "manager.searchLabel": "搜尋筆記",
      "manager.searchPlaceholder": "輸入關鍵字或頁面標題",
      "manager.downloadAll": "下載全部筆記",
      "manager.importMulti": "匯入多個 JSON",
      "manager.githubBackup": "GitHub 備份",
      "manager.githubDesc": "填入 Personal Access Token 與儲存庫資訊，即可一鍵把全部筆記上傳到 GitHub。",
      "manager.githubTokenLabel": "GitHub Token",
      "manager.githubTokenPlaceholder": "ghp_xxx",
      "manager.githubRepoLabel": "儲存庫（owner/repo）",
      "manager.githubRepoPlaceholder": "username/highlight-backup",
      "manager.githubBranchLabel": "分支",
      "manager.githubBranchPlaceholder": "main",
      "manager.githubPathLabel": "檔案路徑",
      "manager.githubPathPlaceholder": "backups/highlight-keeper.json",
      "manager.githubDownload": "從 GitHub 下載",
      "manager.githubUpload": "上傳到 GitHub",
      "manager.pageListHeading": "頁面清單",
      "manager.errJsonInvalid": "JSON 格式不正確",
      "manager.statusPageCount": "共 {count} 個頁面",
      "manager.emptySearch": "沒有符合搜尋條件的頁面。",
      "manager.emptyAll": "目前尚未建立任何筆記。",
      "manager.pageMetaList": "筆記數：{total} · 最後更新：{updated}",
      "manager.pageMetaDetail": "筆記 {total} 則 · 最後更新：{updated}",
      "manager.unknownTime": "未知時間",
      "manager.errGithubBackupInvalid": "GitHub 備份檔案不是有效的 JSON 格式",
      "manager.errGithubBackupEmpty": "GitHub 備份中沒有頁面資料",
      "manager.errGithubBackupNoImport": "GitHub 備份裡沒有可匯入的筆記",
      "manager.errNoNotesExport": "沒有筆記可匯出",
      "manager.statusAllDownloaded": "已下載全部筆記",
      "manager.btnCancel": "取消",
      "manager.btnConfirm": "確定",
      "manager.dialogTitle": "確認動作",
      "manager.notesSection": "筆記",
      "manager.tagsLabel": "頁面標籤",
      "manager.tagsPlaceholder": "以逗號或空白分隔多個 Tags",
      "manager.applyTags": "套用標籤",
      "manager.aiRecord": "AI 紀錄",
      "manager.emptyNotes": "尚無筆記。",
      "manager.emptyContent": "(無內容)",
      "manager.clickToCopyNote": "點擊複製註解",
      "manager.statusNoteCopied": "已複製註解",
      "manager.errNoteCopy": "無法複製註解",
      "manager.emptyTags": "尚無標籤",
      "manager.emptyAiNote": "尚未產生 AI 紀錄。",
      "manager.statusTagsUpdated": "已更新頁面標籤",
      "manager.errTagsUpdate": "無法更新頁面標籤",
      "manager.errNoToken": "請輸入 GitHub Token",
      "manager.errNoRepo": "請輸入 owner/repo 格式的儲存庫",
      "manager.errNoPath": "請輸入檔案路徑",
      "manager.errRepoInvalid": "儲存庫格式不正確",
      "manager.errGithubRead": "GitHub 讀取檔案失敗：{error}",
      "manager.errGithubNotFound": "GitHub 上找不到指定的備份檔案",
      "manager.errGithubPathIsDir": "GitHub 路徑指向資料夾，請指定備份檔案",
      "manager.errGithubEmpty": "GitHub 回傳的檔案內容為空",
      "manager.statusUploading": "上傳中…",
      "manager.errNoNotesUpload": "目前沒有筆記可上傳",
      "manager.statusUploaded": "已成功上傳到 GitHub",
      "manager.errUpload": "上傳失敗",
      "manager.errGithubNoPages": "GitHub 備份中沒有可匯入的頁面",
      "manager.statusAlreadyLatest": "已是最新版本，沒有需要更新的筆記。",
      "manager.statusImportCancelled": "已取消匯入，本機筆記保持不變。",
      "manager.statusGithubImported": "已從 GitHub 匯入 {count} 個頁面",
      "manager.statusDownloading": "從 GitHub 下載中…",
      "manager.errDownload": "下載失敗",
      "manager.statusParsing": "解析匯入檔案中…",
      "manager.errNoImport": "沒有可匯入的筆記",
      "manager.errAllExist": "所有頁面皆已有筆記，已忽略匯入。",
      "manager.statusImportPartial": "成功匯入 {imported} 個頁面，跳過 {skipped} 個已存在的頁面。",
      "manager.errLoadNotes": "無法載入筆記",
      "manager.errImport": "匯入失敗",

      // ── floating toolbar / buttons ───────────────────────
      "toolbar.highlightTitle": "標註",
      "toolbar.highlightText": "HL",
      "toolbar.translateTitle": "翻譯",
      "toolbar.translateText": "譯",

      // ── translate card ───────────────────────────────────
      "translate.detecting": "偵測語言中…",
      "translate.arrow": "→",
      "translate.close": "×",
      "translate.translating": "翻譯中…",
      "translate.errRequest": "翻譯請求失敗 ({status})",
      "translate.errFallback": "翻譯失敗",

      // ── in-page panel (contentScript) ────────────────────
      "panel.title": "此頁標註",
      "panel.fontDecrease": "A-",
      "panel.fontIncrease": "A+",
      "panel.close": "×",
      "panel.tabHighlights": "標註",
      "panel.tabArchive": "存檔",
      "panel.tabSearch": "搜尋",
      "panel.tabAiNote": "AI 筆記",
      "panel.searchHeader": "搜尋筆記",
      "panel.searchPlaceholder": "輸入關鍵字或標籤",
      "panel.pageToolsHeader": "頁面工具",
      "panel.clearPage": "清除本頁標記",
      "panel.clearPageTitle": "刪除此頁所有標記，可重新 AI 標示",
      "panel.copyContent": "複製內容",
      "panel.downloadJson": "下載 JSON",
      "panel.importJson": "匯入 JSON",
      "panel.archivePageHint": "本頁筆記",
      "panel.archiveCopy": "複製文字",
      "panel.archiveDownload": "下載 JSON",
      "panel.archiveImport": "匯入 JSON",
      "panel.archiveAllHint": "全部筆記",
      "panel.archiveAllDownload": "下載全部筆記",
      "panel.archiveAllImport": "匯入全部筆記",
      "panel.providerLabel": "服務提供者",
      "panel.modelLabel": "模型",
      "panel.openaiKeyLabel": "OpenAI API Key",
      "panel.openaiKeyPlaceholder": "sk-...",
      "panel.geminiKeyLabel": "Gemini API Key",
      "panel.geminiKeyPlaceholder": "AIza...",
      "panel.chatgptInfo": "無需 API Key。點擊功能按鈕後，Prompt 會自動複製到剪貼簿並開啟 ChatGPT。取得回應後貼入各區塊下方即可套用。",
      "panel.chatgptCancel": "取消等待",
      "panel.autoHighlightTitle": "自動畫重點",
      "panel.autoHighlightDesc": "AI 分析頁面內容，自動找出重要段落並套上顏色標記",
      "panel.promptLabel": "Prompt",
      "panel.autoHighlightBtn": "自動畫重點",
      "panel.chatgptPasteHl": "貼上 ChatGPT 回應後點「套用重點」…",
      "panel.applyHighlights": "套用重點",
      "panel.categoryLabel": "標記分類",
      "panel.addCategory": "+ 新增分類",
      "panel.resetCategories": "重置為預設分類",
      "panel.categoryPlaceholder": "分類名稱",
      "panel.deleteCategory": "✕",
      "panel.noteSummaryTitle": "AI 摘要筆記",
      "panel.noteSummaryDesc": "整篇文章的大致摘要，整理成有脈絡的筆記",
      "panel.generateNote": "產生筆記",
      "panel.chatgptPasteNote": "貼上 ChatGPT 回應後點「套用筆記」…",
      "panel.applyNote": "套用筆記",
      "panel.aiNotesTitle": "AI 筆記",
      "panel.aiNotesCopy": "複製",
      "panel.aiNotesCopied": "已複製",
      "panel.aiNotesCopyFail": "複製失敗",
      "panel.emptyNote": "尚未產生筆記。",
      "panel.guidedReading": "指引閱讀",
      "panel.emptyPage": "本頁尚未建立標註。",
      "panel.emptySearch": "目前沒有符合的筆記。",
      "panel.tagsEmpty": "尚未建立標籤",
      "panel.tagsAll": "全部",
      "panel.tagInputHint": "輸入後按 Enter 或按新增；點擊標籤可移除",
      "panel.tagEmptyState": "尚未為此頁設定標籤",
      "panel.tagRemoveTitle": "移除此標籤",
      "panel.previewConfirm": "確認套用",
      "panel.previewCancel": "取消",
      "panel.previewMessage": "預覽 {count} 個重點，確認後正式套用",

      // ── highlight detail (popover) ───────────────────────
      "detail.heading": "標註設定",
      "detail.colorLabel": "顏色",
      "detail.noteLabel": "註解",
      "detail.translate": "翻譯",
      "detail.translateTitle": "翻譯此段標記並附加到備註",
      "detail.notePlaceholder": "輸入註解...",
      "detail.saveNote": "儲存註解",
      "detail.deleteHighlight": "刪除標註",

      // ── AI status / errors ───────────────────────────────
      "ai.statusClearedAndReady": "已清除本頁所有標記，可重新標記",
      "ai.statusPreviewCancelled": "已取消預覽",
      "ai.statusPreparing": "準備中…",
      "ai.statusChatgptOpened": "已開啟 ChatGPT，請在跳出的頁面點「複製 Prompt」並送出",
      "ai.errChatgptBridge": "無法啟動 ChatGPT 橋接",
      "ai.statusChatgptNote": "ChatGPT 筆記已匯入",
      "ai.statusApplyingHighlights": "套用重點中…",
      "ai.errNoHighlights": "ChatGPT 回傳沒有可用的重點",
      "ai.errNoMatchingText": "找不到可標註的文字",
      "ai.errImport": "匯入失敗",
      "ai.errNoApiKey": "請先輸入 API Key",
      "ai.statusGeneratingNote": "產生筆記中…",
      "ai.statusNoteDone": "已完成筆記產生",
      "ai.errGenerateNote": "無法產生筆記",
      "ai.statusAnalyzing": "分析重點中…",
      "ai.statusFallbackWholePage": "未選取文字，改為分析整頁",
      "ai.errNoUsableItems": "AI 沒有回傳可用的重點項目",
      "ai.errNoMatchAdjust": "找不到可標註的文字，請調整 Prompt 後再試",
      "ai.statusPreviewN": "預覽 {count} 個重點，請確認後套用",
      "ai.errAutoHighlight": "無法自動畫重點",
      "ai.errResponseEmpty": "AI 回傳內容為空",
      "ai.errInvalidJson": "AI 回傳不是有效 JSON",
      "ai.errOpenaiApi": "OpenAI API 錯誤：{error}",
      "ai.errOpenaiEmpty": "OpenAI 回傳內容為空",
      "ai.errGeminiApi": "Gemini API 錯誤：{error}",
      "ai.errGeminiEmpty": "Gemini 回傳內容為空",
      "ai.pasteFirst": "請先貼上 ChatGPT 回應",
      "ai.statusAppliedN": "已自動畫重點 {applied} 筆",
      "ai.statusAppliedWithSkip": "已自動畫重點 {applied} 筆，略過 {skipped} 筆",
      "ai.statusHighlightDeleted": "已刪除標註",
      "ai.errDeleteHighlight": "刪除失敗",
      "ai.errNoHighlightExport": "沒有可匯出的標註",
      "ai.statusTextCopied": "已複製純文字",
      "ai.errTextCopy": "複製失敗",
      "ai.statusJsonDownloaded": "已下載 JSON",
      "ai.errJsonInvalid": "JSON 格式不正確",
      "ai.errStorageAccess": "無法存取瀏覽器儲存空間",
      "ai.errNoHighlightImport": "沒有可匯入的標註",
      "ai.errNoHighlightInFile": "檔案中沒有標註資料",
      "ai.errImportFailed": "匯入失敗",
      "ai.errNoHighlightGuided": "本頁沒有標記可供閱讀",

      // ── guided reading ───────────────────────────────────
      "guided.exit": "✕ 結束",
      "guided.prev": "← 上一個",
      "guided.next": "下一個 →",

      // ── ChatGPT bridge UI ────────────────────────────────
      "bridge.copyPrompt": "複製 Prompt",
      "bridge.copied": "✓ 已複製",
      "bridge.copyFail": "複製失敗",
      "bridge.noResponse": "找不到回應，請確認已完成",
      "bridge.importing": "匯入中…",
      "bridge.importSuccess": "✓ 已匯入！請切回原頁面查看結果。",
      "bridge.importFail": "匯入失敗，請重試",
      "bridge.closeTitle": "關閉（取消此次橋接）",
      "bridge.close": "✕",

      // ── AI prompts (sent to model) ───────────────────────
      "prompt.systemAssistant": "你是一位筆記整理助手，根據提供的網頁全文與標註內容，整理出淺顯易懂的筆記。優先考慮使用者標注段落，將重點控制在五百字以內，輸出內容需要像說故事一樣有脈絡的說明。",
      "prompt.systemAnalyzer": "你是一位文章結構分析助手。請全面標註文章中各個關鍵面向的重要片段，確保覆蓋文章的完整論述架構，不要只挑少數幾句。",
      "prompt.systemJsonStrict": "You are a strict JSON extraction assistant. Return only valid JSON without markdown.",
      "prompt.systemNoteWriter": "你是一位筆記整理助手，輸出繁體中文。",
      "prompt.note.webInfo": "### 網頁資訊",
      "prompt.note.titleField": "- 標題：{title}",
      "prompt.note.urlField": "- URL：{url}",
      "prompt.note.original": "### 原文內容（可能已截斷）",
      "prompt.note.userHighlights": "### 使用者標註",
      "prompt.note.noHighlights": "（尚未加入標註）",
      "prompt.note.outputLang": "請以繁體中文輸出筆記。",
      "prompt.hl.categoriesHeader": "### 標記分類定義（必須嚴格依語義對應）",
      "prompt.hl.categoriesInstruction": "請根據文字語義選擇對應分類的顏色，只能使用以上顏色。",
      "prompt.hl.categoriesMinCount": "- 每個分類至少找 1 個片段（文章有提及的話），確保標註覆蓋完整論述架構。",
      "prompt.hl.categoriesTotalCount": "- 共可回傳 {count} 個以上的片段，不要因為「夠了」就停止。",
      "prompt.hl.colorsHeader": "### 可用顏色（依使用習慣排序）",
      "prompt.hl.colorsInstruction": "顏色分配請優先符合使用者習慣（可優先使用歷史使用次數較高的顏色）。",
      "prompt.hl.colorsCount": "- 盡量標註 10～20 個片段，確保覆蓋文章各個段落，不要只挑少數幾句。",
      "prompt.hl.taskHeader": "### 任務要求",
      "prompt.hl.taskNoLimit": "- 標註數量不要自我設限，文章夠長就多標。",
      "prompt.hl.taskContiguous": "- 每個重點必須是「原文中可直接找到」的連續文字。",
      "prompt.hl.taskNoDup": "- 不要回傳與「已有標註」重複的內容。",
      "prompt.hl.taskColorsOnly": "- 必須只使用下方提供的顏色。",
      "prompt.hl.taskReason": "- reason 欄位：用一句話直接摘要「這段原文的核心觀點或資訊」，不加任何前綴（不說「這段在說」「這裡提到」「作者指出」等），直接陳述內容本身。",
      "prompt.hl.outputHeader": "### 輸出格式（只能輸出 JSON，禁止 markdown 或額外說明）",
      "prompt.hl.outputFieldComment.text": "要標註的原文片段",
      "prompt.hl.outputFieldComment.reason": "用一句話直接摘要這段原文的核心觀點或資訊（不要加任何前綴）",
      "prompt.hl.existing": "### 已有標註",
      "prompt.hl.reasonLang": "reason 欄位請以繁體中文書寫。",
    },
    en: {
      // ── popup ────────────────────────────────────────────
      "popup.title": "Page Highlights",
      "popup.colorLabel": "Highlight color",
      "popup.addToPalette": "Add to palette",
      "popup.availableColors": "Available colors",
      "popup.paletteAria": "Color palette",
      "popup.panelControlAria": "Page highlight panel",
      "popup.aiToolsAria": "AI tools",
      "popup.dataToolsAria": "Data management",
      "popup.openPanel": "Open in-page panel",
      "popup.aiAutoHighlight": "Copy AI highlight prompt",
      "popup.clearPage": "Clear this page",
      "popup.openManager": "Manage all notes",
      "popup.contact": "Questions? Email chrouodiu@gmail.com",
      "popup.langLabel": "Language",
      "popup.langZh": "繁體中文",
      "popup.langEn": "English",
      "popup.emptyPalette": "No colors yet. Add one to the palette.",
      "popup.useColor": "Use {color}",
      "popup.editColor": "Edit color {color}",
      "popup.deleteColor": "Delete",
      "popup.deleteColorTitle": "Delete color {color}",
      "popup.statusSelected": "Selected color {color}",
      "popup.statusUpdated": "Palette color updated to {color}",
      "popup.statusDeleted": "Removed color {color}",
      "popup.statusDuplicate": "This color is already in the palette",
      "popup.statusAdded": "Added color {color}",
      "popup.errNoTab": "No active tab found — cannot open the panel.",
      "popup.errOpenPanel": "Cannot open the in-page panel",
      "popup.statusPanelOpened": "Panel shown on page",
      "popup.errOpenPanel2": "Cannot open the in-page panel.",
      "popup.errNoTabSimple": "No tab found",
      "popup.errCannotTrigger": "Cannot trigger",
      "popup.statusAiTriggered": "AI auto-highlight triggered",
      "popup.statusAiCopied": "Highlight prompt copied — paste it into ChatGPT, then paste the reply back into the in-page panel",
      "popup.errAiTrigger": "Cannot build highlight prompt",
      "popup.pageInfoAria": "Page info",
      "popup.markCountLabel": "Marks on page",
      "popup.aiModeLabel": "AI mode",
      "popup.aiModeOpenai": "OpenAI API",
      "popup.aiModeGemini": "Gemini API",
      "popup.aiModeCopy": "Copy & paste (ChatGPT)",
      "popup.errCannotClear": "Cannot clear",
      "popup.statusCleared": "Cleared all highlights on this page",
      "popup.errClear": "Cannot clear highlights on this page",
      "popup.errCannotSendMessage": "Cannot send message",

      // ── manager ──────────────────────────────────────────
      "manager.pageTitle": "Highlight Keeper — Notes overview",
      "manager.heading": "All notes",
      "manager.subheading": "Export/import all page highlights in one place.",
      "manager.close": "×",
      "manager.searchLabel": "Search notes",
      "manager.searchPlaceholder": "Keyword or page title",
      "manager.downloadAll": "Download all notes",
      "manager.importMulti": "Import multiple JSON",
      "manager.githubBackup": "GitHub backup",
      "manager.githubDesc": "Fill in a Personal Access Token and repo info to back up all notes to GitHub.",
      "manager.githubTokenLabel": "GitHub Token",
      "manager.githubTokenPlaceholder": "ghp_xxx",
      "manager.githubRepoLabel": "Repo (owner/repo)",
      "manager.githubRepoPlaceholder": "username/highlight-backup",
      "manager.githubBranchLabel": "Branch",
      "manager.githubBranchPlaceholder": "main",
      "manager.githubPathLabel": "File path",
      "manager.githubPathPlaceholder": "backups/highlight-keeper.json",
      "manager.githubDownload": "Download from GitHub",
      "manager.githubUpload": "Upload to GitHub",
      "manager.pageListHeading": "Pages",
      "manager.errJsonInvalid": "Invalid JSON format",
      "manager.statusPageCount": "{count} page(s)",
      "manager.emptySearch": "No pages match your search.",
      "manager.emptyAll": "No notes yet.",
      "manager.pageMetaList": "Notes: {total} · Updated: {updated}",
      "manager.pageMetaDetail": "{total} note(s) · Updated: {updated}",
      "manager.unknownTime": "unknown",
      "manager.errGithubBackupInvalid": "GitHub backup file is not valid JSON",
      "manager.errGithubBackupEmpty": "GitHub backup contains no pages",
      "manager.errGithubBackupNoImport": "GitHub backup has no importable notes",
      "manager.errNoNotesExport": "No notes to export",
      "manager.statusAllDownloaded": "Downloaded all notes",
      "manager.btnCancel": "Cancel",
      "manager.btnConfirm": "Confirm",
      "manager.dialogTitle": "Confirm",
      "manager.notesSection": "Notes",
      "manager.tagsLabel": "Page tags",
      "manager.tagsPlaceholder": "Separate tags with comma or space",
      "manager.applyTags": "Apply tags",
      "manager.aiRecord": "AI record",
      "manager.emptyNotes": "No notes yet.",
      "manager.emptyContent": "(empty)",
      "manager.clickToCopyNote": "Click to copy note",
      "manager.statusNoteCopied": "Note copied",
      "manager.errNoteCopy": "Cannot copy note",
      "manager.emptyTags": "No tags yet",
      "manager.emptyAiNote": "No AI notes yet.",
      "manager.statusTagsUpdated": "Tags updated",
      "manager.errTagsUpdate": "Cannot update page tags",
      "manager.errNoToken": "Please enter a GitHub Token",
      "manager.errNoRepo": "Please enter the repo as owner/repo",
      "manager.errNoPath": "Please enter the file path",
      "manager.errRepoInvalid": "Repo format is invalid",
      "manager.errGithubRead": "Failed to read from GitHub: {error}",
      "manager.errGithubNotFound": "Backup file not found on GitHub",
      "manager.errGithubPathIsDir": "The path points to a directory; specify a file",
      "manager.errGithubEmpty": "GitHub returned an empty file",
      "manager.statusUploading": "Uploading…",
      "manager.errNoNotesUpload": "No notes to upload",
      "manager.statusUploaded": "Uploaded to GitHub",
      "manager.errUpload": "Upload failed",
      "manager.errGithubNoPages": "GitHub backup has no importable pages",
      "manager.statusAlreadyLatest": "Already up to date — no notes to update.",
      "manager.statusImportCancelled": "Import cancelled; local notes unchanged.",
      "manager.statusGithubImported": "Imported {count} page(s) from GitHub",
      "manager.statusDownloading": "Downloading from GitHub…",
      "manager.errDownload": "Download failed",
      "manager.statusParsing": "Parsing import file…",
      "manager.errNoImport": "No notes to import",
      "manager.errAllExist": "All pages already have notes; nothing was imported.",
      "manager.statusImportPartial": "Imported {imported} page(s); skipped {skipped} already-existing page(s).",
      "manager.errLoadNotes": "Cannot load notes",
      "manager.errImport": "Import failed",

      // ── floating toolbar / buttons ───────────────────────
      "toolbar.highlightTitle": "Highlight",
      "toolbar.highlightText": "HL",
      "toolbar.translateTitle": "Translate",
      "toolbar.translateText": "Tr",

      // ── translate card ───────────────────────────────────
      "translate.detecting": "Detecting language…",
      "translate.arrow": "→",
      "translate.close": "×",
      "translate.translating": "Translating…",
      "translate.errRequest": "Translation request failed ({status})",
      "translate.errFallback": "Translation failed",

      // ── in-page panel ────────────────────────────────────
      "panel.title": "Highlights on this page",
      "panel.fontDecrease": "A-",
      "panel.fontIncrease": "A+",
      "panel.close": "×",
      "panel.tabHighlights": "Highlights",
      "panel.tabArchive": "Archive",
      "panel.tabSearch": "Search",
      "panel.tabAiNote": "AI Notes",
      "panel.searchHeader": "Search notes",
      "panel.searchPlaceholder": "Keyword or tag",
      "panel.pageToolsHeader": "Page tools",
      "panel.clearPage": "Clear this page",
      "panel.clearPageTitle": "Delete all highlights on this page; AI can re-highlight",
      "panel.copyContent": "Copy content",
      "panel.downloadJson": "Download JSON",
      "panel.importJson": "Import JSON",
      "panel.archivePageHint": "This page",
      "panel.archiveCopy": "Copy text",
      "panel.archiveDownload": "Download JSON",
      "panel.archiveImport": "Import JSON",
      "panel.archiveAllHint": "All notes",
      "panel.archiveAllDownload": "Download all notes",
      "panel.archiveAllImport": "Import all notes",
      "panel.providerLabel": "Provider",
      "panel.modelLabel": "Model",
      "panel.openaiKeyLabel": "OpenAI API Key",
      "panel.openaiKeyPlaceholder": "sk-...",
      "panel.geminiKeyLabel": "Gemini API Key",
      "panel.geminiKeyPlaceholder": "AIza...",
      "panel.chatgptInfo": "No API key required. After clicking a function button, the prompt is copied to your clipboard and ChatGPT opens. Paste the response into the box below to apply.",
      "panel.chatgptCancel": "Cancel wait",
      "panel.autoHighlightTitle": "Auto-highlight",
      "panel.autoHighlightDesc": "AI analyzes the page and highlights key segments with color tags",
      "panel.promptLabel": "Prompt",
      "panel.autoHighlightBtn": "Auto-highlight",
      "panel.chatgptPasteHl": "Paste ChatGPT response, then click \"Apply highlights\"…",
      "panel.applyHighlights": "Apply highlights",
      "panel.categoryLabel": "Categories",
      "panel.addCategory": "+ Add category",
      "panel.resetCategories": "Reset to defaults",
      "panel.categoryPlaceholder": "Category name",
      "panel.deleteCategory": "✕",
      "panel.noteSummaryTitle": "AI summary note",
      "panel.noteSummaryDesc": "Summarize the whole article into a coherent note",
      "panel.generateNote": "Generate note",
      "panel.chatgptPasteNote": "Paste ChatGPT response, then click \"Apply note\"…",
      "panel.applyNote": "Apply note",
      "panel.aiNotesTitle": "AI Notes",
      "panel.aiNotesCopy": "Copy",
      "panel.aiNotesCopied": "Copied",
      "panel.aiNotesCopyFail": "Copy failed",
      "panel.emptyNote": "No notes generated yet.",
      "panel.guidedReading": "Guided reading",
      "panel.emptyPage": "No highlights on this page yet.",
      "panel.emptySearch": "No matching notes.",
      "panel.tagsEmpty": "No tags yet",
      "panel.tagsAll": "All",
      "panel.tagInputHint": "Press Enter to add; click a tag to remove",
      "panel.tagEmptyState": "No tags set for this page",
      "panel.tagRemoveTitle": "Remove this tag",
      "panel.previewConfirm": "Apply",
      "panel.previewCancel": "Cancel",
      "panel.previewMessage": "Previewing {count} highlights — confirm to apply",

      // ── highlight detail ─────────────────────────────────
      "detail.heading": "Highlight settings",
      "detail.colorLabel": "Color",
      "detail.noteLabel": "Note",
      "detail.translate": "Translate",
      "detail.translateTitle": "Translate this highlighted text and append to note",
      "detail.notePlaceholder": "Type a note…",
      "detail.saveNote": "Save note",
      "detail.deleteHighlight": "Delete highlight",

      // ── AI status / errors ───────────────────────────────
      "ai.statusClearedAndReady": "Page cleared — ready to highlight again",
      "ai.statusPreviewCancelled": "Preview cancelled",
      "ai.statusPreparing": "Preparing…",
      "ai.statusChatgptOpened": "ChatGPT opened — click \"Copy Prompt\" on that tab and submit",
      "ai.errChatgptBridge": "Cannot start ChatGPT bridge",
      "ai.statusChatgptNote": "ChatGPT note imported",
      "ai.statusApplyingHighlights": "Applying highlights…",
      "ai.errNoHighlights": "ChatGPT response has no usable highlights",
      "ai.errNoMatchingText": "Could not find matching text to highlight",
      "ai.errImport": "Import failed",
      "ai.errNoApiKey": "Please enter an API Key first",
      "ai.statusGeneratingNote": "Generating note…",
      "ai.statusNoteDone": "Note generated",
      "ai.errGenerateNote": "Cannot generate note",
      "ai.statusAnalyzing": "Analyzing…",
      "ai.statusFallbackWholePage": "No selection — analyzing the whole page",
      "ai.errNoUsableItems": "AI returned no usable highlight items",
      "ai.errNoMatchAdjust": "No matching text found — adjust the prompt and try again",
      "ai.statusPreviewN": "Previewing {count} highlights — confirm to apply",
      "ai.errAutoHighlight": "Auto-highlight failed",
      "ai.errResponseEmpty": "AI returned empty content",
      "ai.errInvalidJson": "AI response is not valid JSON",
      "ai.errOpenaiApi": "OpenAI API error: {error}",
      "ai.errOpenaiEmpty": "OpenAI returned empty content",
      "ai.errGeminiApi": "Gemini API error: {error}",
      "ai.errGeminiEmpty": "Gemini returned empty content",
      "ai.pasteFirst": "Please paste the ChatGPT response first",
      "ai.statusAppliedN": "Auto-highlighted {applied}",
      "ai.statusAppliedWithSkip": "Auto-highlighted {applied}, skipped {skipped}",
      "ai.statusHighlightDeleted": "Highlight deleted",
      "ai.errDeleteHighlight": "Delete failed",
      "ai.errNoHighlightExport": "No highlights to export",
      "ai.statusTextCopied": "Text copied",
      "ai.errTextCopy": "Copy failed",
      "ai.statusJsonDownloaded": "JSON downloaded",
      "ai.errJsonInvalid": "Invalid JSON format",
      "ai.errStorageAccess": "Cannot access browser storage",
      "ai.errNoHighlightImport": "No highlights to import",
      "ai.errNoHighlightInFile": "File contains no highlight data",
      "ai.errImportFailed": "Import failed",
      "ai.errNoHighlightGuided": "No highlights on this page to read",

      // ── guided reading ───────────────────────────────────
      "guided.exit": "✕ Exit",
      "guided.prev": "← Prev",
      "guided.next": "Next →",

      // ── ChatGPT bridge UI ────────────────────────────────
      "bridge.copyPrompt": "Copy Prompt",
      "bridge.copied": "✓ Copied",
      "bridge.copyFail": "Copy failed",
      "bridge.noResponse": "No response found — finish ChatGPT first",
      "bridge.importing": "Importing…",
      "bridge.importSuccess": "✓ Imported! Switch back to the original page to see the result.",
      "bridge.importFail": "Import failed, please retry",
      "bridge.closeTitle": "Close (cancel this bridge)",
      "bridge.close": "✕",

      // ── AI prompts ───────────────────────────────────────
      "prompt.systemAssistant": "You are a note-taking assistant. Given the full webpage text and existing highlights, produce easy-to-read notes. Prioritize highlighted segments, keep within ~500 words, and write in a clear narrative.",
      "prompt.systemAnalyzer": "You are an article structure analyst. Comprehensively highlight key segments across all major aspects of the article — make sure the highlights cover the full argument structure, not just a few sentences.",
      "prompt.systemJsonStrict": "You are a strict JSON extraction assistant. Return only valid JSON without markdown.",
      "prompt.systemNoteWriter": "You are a note-taking assistant. Output in English.",
      "prompt.note.webInfo": "### Page info",
      "prompt.note.titleField": "- Title: {title}",
      "prompt.note.urlField": "- URL: {url}",
      "prompt.note.original": "### Original content (may be truncated)",
      "prompt.note.userHighlights": "### User highlights",
      "prompt.note.noHighlights": "(no highlights yet)",
      "prompt.note.outputLang": "Write the note in English.",
      "prompt.hl.categoriesHeader": "### Category definitions (match strictly by meaning)",
      "prompt.hl.categoriesInstruction": "Choose the color whose category matches the meaning of the text. Use only the colors above.",
      "prompt.hl.categoriesMinCount": "- Find at least 1 segment per category (when the article mentions it), so the highlights cover the full argument.",
      "prompt.hl.categoriesTotalCount": "- Return at least {count} segments; don't stop early because you think it's enough.",
      "prompt.hl.colorsHeader": "### Available colors (ordered by user preference)",
      "prompt.hl.colorsInstruction": "Prefer colors used more frequently by the user.",
      "prompt.hl.colorsCount": "- Aim for 10–20 segments to cover the whole article; don't pick just a few sentences.",
      "prompt.hl.taskHeader": "### Task requirements",
      "prompt.hl.taskNoLimit": "- Don't artificially limit the number of highlights; longer articles deserve more.",
      "prompt.hl.taskContiguous": "- Each highlight must be a contiguous span that appears verbatim in the original text.",
      "prompt.hl.taskNoDup": "- Don't return content that duplicates existing highlights.",
      "prompt.hl.taskColorsOnly": "- Use only the colors listed above.",
      "prompt.hl.taskReason": "- reason field: in one sentence, directly summarize the core point or information of this segment. No prefixes (don't say \"this paragraph says\", \"the author points out\", etc.); just state the content directly.",
      "prompt.hl.outputHeader": "### Output format (JSON only — no markdown or extra commentary)",
      "prompt.hl.outputFieldComment.text": "the exact span to highlight",
      "prompt.hl.outputFieldComment.reason": "one-sentence summary of the segment's core point (no prefix)",
      "prompt.hl.existing": "### Existing highlights",
      "prompt.hl.reasonLang": "Write the reason field in English.",
    },
  };

  const detectLang = () => {
    const nav = (root.navigator?.language || root.navigator?.userLanguage || "")
      .toLowerCase();
    if (nav.startsWith("zh")) return "zh";
    return "en";
  };

  let currentLang = detectLang();
  let ready = Promise.resolve(currentLang);
  const listeners = new Set();

  const interpolate = (template, params) => {
    if (!params) return template;
    return template.replace(/\{(\w+)\}/g, (m, key) =>
      Object.prototype.hasOwnProperty.call(params, key) ? String(params[key]) : m
    );
  };

  const t = (key, params) => {
    const table = MESSAGES[currentLang] || MESSAGES[DEFAULT_FALLBACK];
    const value = table[key] ?? MESSAGES[DEFAULT_FALLBACK][key] ?? key;
    return interpolate(value, params);
  };

  const getLang = () => currentLang;

  const notify = () => {
    listeners.forEach((fn) => {
      try { fn(currentLang); } catch (_e) {}
    });
  };

  const onLangChange = (fn) => {
    listeners.add(fn);
    return () => listeners.delete(fn);
  };

  const persist = async (lang) => {
    if (!root.chrome?.storage?.local) return;
    try { await root.chrome.storage.local.set({ [STORAGE_KEY]: lang }); } catch (_e) {}
  };

  const setLang = async (lang) => {
    if (!SUPPORTED.includes(lang)) return;
    if (lang === currentLang) return;
    currentLang = lang;
    await persist(lang);
    applyDOMTranslations();
    notify();
  };

  const applyDOMTranslations = (rootEl) => {
    const scope = rootEl || root.document;
    if (!scope?.querySelectorAll) return;
    scope.querySelectorAll("[data-i18n]").forEach((el) => {
      const key = el.getAttribute("data-i18n");
      if (key) el.textContent = t(key);
    });
    scope.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
      const key = el.getAttribute("data-i18n-placeholder");
      if (key) el.setAttribute("placeholder", t(key));
    });
    scope.querySelectorAll("[data-i18n-title]").forEach((el) => {
      const key = el.getAttribute("data-i18n-title");
      if (key) el.setAttribute("title", t(key));
    });
    scope.querySelectorAll("[data-i18n-aria-label]").forEach((el) => {
      const key = el.getAttribute("data-i18n-aria-label");
      if (key) el.setAttribute("aria-label", t(key));
    });
  };

  const initI18n = async () => {
    if (root.chrome?.storage?.local) {
      try {
        const stored = await root.chrome.storage.local.get(STORAGE_KEY);
        const lang = stored?.[STORAGE_KEY];
        if (SUPPORTED.includes(lang)) currentLang = lang;
      } catch (_e) {}
      root.chrome.storage.onChanged?.addListener?.((changes, area) => {
        if (area !== "local" || !changes[STORAGE_KEY]) return;
        const newLang = changes[STORAGE_KEY].newValue;
        if (SUPPORTED.includes(newLang) && newLang !== currentLang) {
          currentLang = newLang;
          applyDOMTranslations();
          notify();
        }
      });
    }
    ready = Promise.resolve(currentLang);
    return currentLang;
  };

  root.HkI18n = {
    t,
    getLang,
    setLang,
    initI18n,
    applyDOMTranslations,
    onLangChange,
    SUPPORTED,
    get ready() { return ready; },
  };
})(typeof window !== "undefined" ? window : globalThis);
