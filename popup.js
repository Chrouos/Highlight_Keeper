/* ══════════════════════════════════════════════════════
   Highlight Keeper — Popup
   三個快捷區塊：① 複製 Prompt ② 貼上 ③ 分享此網頁
   其餘設定（色票、標籤、備份、AI 服務）都在頁面面板裡。
   ════════════════════════════════════════════════════ */

const els = {
  markCount: document.getElementById("markCount"),
  aiMode: document.getElementById("aiMode"),
  status: document.getElementById("popStatus"),
  copyPromptBtn: document.getElementById("aiHighlightBtn"),
  mindmapPromptBtn: document.getElementById("mindmapPromptBtn"),
  pasteInput: document.getElementById("aiPasteInput"),
  pasteApplyBtn: document.getElementById("aiPasteApplyBtn"),
  shareLinkBtn: document.getElementById("shareLinkBtn"),
  copyNotesBtn: document.getElementById("copyNotesBtn"),
  downloadNotesBtn: document.getElementById("downloadNotesBtn"),
  openPanelBtn: document.getElementById("openPanelBtn"),
  openManagerBtn: document.getElementById("openManagerBtn"),
  langSelect: document.getElementById("langSelect"),
};

const t = (key, params) => HkI18n.t(key, params);

// 網址正規化共用自 shared.js（popup.html 已先載入）。
const normalizePageKey = (href) =>
  window.HkUrlKey ? window.HkUrlKey.normalizePageKey(href) : href;

// 與 manager.js / contentScript 一致的 storage 主鍵。
const PAGE_META_KEY = "__hk_page_meta__";
const GENERATED_NOTES_KEY = "hkGeneratedNotes";
const MINDMAP_KEY = "hkMindmaps";
const GITHUB_SETTINGS_KEY = "hkGithubSyncSettings";

const RECEIVER_MISSING_ERROR = "Receiving end does not exist.";

let panelSide = "right";
let statusTimer = 0;

/* ── 狀態列 ─────────────────────────────────────────── */
const setStatus = (message, isError = false) => {
  if (!els.status) return;
  window.clearTimeout(statusTimer);
  els.status.textContent = message || "";
  els.status.classList.toggle("is-error", Boolean(isError));
  els.status.classList.toggle("is-visible", Boolean(message));
  if (message && !isError) {
    statusTimer = window.setTimeout(() => {
      els.status.classList.remove("is-visible");
    }, 6000);
  }
};

// 按鈕短暫變成「已複製 ✓」，比只看狀態列更直覺。
const flashBtn = (btn, doneText) => {
  if (!btn) return;
  const original = btn.textContent;
  btn.textContent = doneText;
  btn.classList.add("is-done");
  window.setTimeout(() => {
    btn.textContent = original;
    btn.classList.remove("is-done");
  }, 1500);
};

/* ── 與 content script 溝通（必要時自動注入）─────────── */
const injectContentAssets = async (tabId) => {
  try {
    await chrome.scripting?.insertCSS({
      target: { tabId },
      files: ["contentStyles.css"],
    });
  } catch (error) {
    console.debug("注入面板樣式失敗", error);
  }
  await chrome.scripting?.executeScript({
    target: { tabId },
    files: ["shared.js", "parsers.js", "i18n.js", "contentScript.js"],
  });
};

const sendMessageToTab = async (tabId, payload) => {
  try {
    return await chrome.tabs.sendMessage(tabId, payload);
  } catch (error) {
    const message = chrome.runtime.lastError?.message || error?.message || "";
    if (message.includes(RECEIVER_MISSING_ERROR)) {
      await injectContentAssets(tabId);
      await new Promise((resolve) => setTimeout(resolve, 150));
      return chrome.tabs.sendMessage(tabId, payload);
    }
    throw new Error(message || t("popup.errCannotSendMessage"));
  }
};

const getActiveTab = async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab || null;
};

/* ── 頁面資訊 ───────────────────────────────────────── */
const aiModeLabel = (provider) => {
  switch (provider) {
    case "gemini":
      return t("popup.aiModeGemini");
    case "chatgpt":
      return t("popup.aiModeCopy");
    case "openai":
    default:
      return t("popup.aiModeOpenai");
  }
};

// 找出「當頁」的 storage 主鍵：contentScript 以 canonical 連結為主鍵，popup 不在
// 頁面內，先輕量讀取該頁 canonical（不觸發全文擷取），讓鍵與標註鍵一致。
const resolveCurrentPageKey = async () => {
  const tab = await getActiveTab();
  if (!tab?.url) return { tab, key: "" };
  let canonical = "";
  try {
    const [{ result } = {}] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => document.querySelector('link[rel="canonical"]')?.href || "",
    });
    if (typeof result === "string") canonical = result;
  } catch (_e) {
    // 受限頁面（chrome://、商店等）無法注入；退回網址。
  }
  return { tab, key: normalizePageKey(canonical || tab.url) };
};

const loadPageInfo = async () => {
  try {
    const stored = await chrome.storage?.local.get("hkAISettings");
    const provider = stored?.hkAISettings?.provider || "openai";
    if (els.aiMode) els.aiMode.textContent = aiModeLabel(provider);
  } catch (_e) {
    if (els.aiMode) els.aiMode.textContent = "—";
  }
  try {
    // 直接讀 storage：喚醒 content script 會強制全文擷取，popup 會變慢。
    const { key } = await resolveCurrentPageKey();
    if (!key) {
      if (els.markCount) els.markCount.textContent = "—";
      return;
    }
    const stored = await chrome.storage?.local.get(key);
    const entries = stored?.[key];
    if (els.markCount) {
      els.markCount.textContent = t("popup.markCountValue", {
        count: Array.isArray(entries) ? entries.length : 0,
      });
    }
  } catch (_e) {
    if (els.markCount) els.markCount.textContent = "—";
  }
};

/* ── 本頁匯出資料 ───────────────────────────────────── */
// 讀出「當頁」的匯出格式資料 {url,title,tags,entries,note,mindmap}；沒筆記回 error。
const collectCurrentPageEntry = async () => {
  const { key } = await resolveCurrentPageKey();
  if (!key) return { error: t("popup.errNoTabSimple") };
  const stored = await chrome.storage?.local.get([
    key,
    PAGE_META_KEY,
    GENERATED_NOTES_KEY,
    MINDMAP_KEY,
  ]);
  const entries = Array.isArray(stored?.[key]) ? stored[key] : [];
  const meta = stored?.[PAGE_META_KEY]?.[key] || {};
  const note = stored?.[GENERATED_NOTES_KEY]?.[key] || null;
  const mindmap = stored?.[MINDMAP_KEY]?.[key] || null;
  if (!entries.length && !note && !mindmap) {
    return { error: t("popup.errNoNotes") };
  }
  return {
    pageEntry: {
      url: key,
      title: meta.title || "",
      tags: Array.isArray(meta.tags) ? meta.tags : [],
      entries,
      note,
      mindmap,
    },
  };
};

/* ── ① 複製 Prompt ──────────────────────────────────── */
const copyPrompt = async (kind) => {
  const btn = kind === "mindmap" ? els.mindmapPromptBtn : els.copyPromptBtn;
  try {
    const tab = await getActiveTab();
    if (!tab?.id) {
      setStatus(t("popup.errNoTabSimple"), true);
      return;
    }
    const response = await sendMessageToTab(tab.id, {
      type: kind === "mindmap" ? "BUILD_MINDMAP_PROMPT" : "BUILD_AI_HIGHLIGHT_PROMPT",
    });
    if (!response?.success || !response.prompt) {
      throw new Error(response?.error ?? t("popup.errCannotTrigger"));
    }
    await navigator.clipboard.writeText(response.prompt);
    flashBtn(btn, t("popup.copied"));
    setStatus(
      kind === "mindmap" ? t("popup.statusMindmapCopied") : t("popup.statusAiCopied")
    );
  } catch (error) {
    setStatus(error?.message || t("popup.errAiTrigger"), true);
  }
};

/* ── ② 貼上並套用 ───────────────────────────────────── */
const applyPastedAiResponse = async () => {
  try {
    const tab = await getActiveTab();
    if (!tab?.id) {
      setStatus(t("popup.errNoTabSimple"), true);
      return;
    }
    let text = els.pasteInput?.value?.trim() || "";
    // 框內沒東西就直接讀剪貼簿（已要求 clipboardRead），按一下即可帶入。
    if (!text) {
      try {
        text = (await navigator.clipboard.readText())?.trim() || "";
        if (els.pasteInput && text) els.pasteInput.value = text;
      } catch (_e) {
        /* 剪貼簿空或被拒 → 維持空字串 */
      }
    }
    if (!text) {
      els.pasteInput?.focus();
      setStatus(t("popup.errPasteEmpty"), true);
      return;
    }
    setStatus(t("popup.statusApplyingPaste"));
    const response = await sendMessageToTab(tab.id, {
      type: "APPLY_AI_RESPONSE",
      text,
    });
    if (!response?.success) {
      throw new Error(response?.error ?? t("popup.errCannotTrigger"));
    }
    setStatus(response.summary || t("popup.statusPasteApplied"));
    if (els.pasteInput) els.pasteInput.value = "";
    loadPageInfo();
  } catch (error) {
    setStatus(error?.message || t("popup.errAiTrigger"), true);
  }
};

/* ── ③ 分享此網頁 ───────────────────────────────────── */
// 預設把筆記壓進「原文網址#hk=…」（零設定）；筆記太多放不下時，
// 改 commit 到 GitHub 備份 repo 分享 raw 連結（免開 Pages）。
const shareCurrentPageLink = async () => {
  if (els.shareLinkBtn) els.shareLinkBtn.disabled = true;
  try {
    const { pageEntry, error } = await collectCurrentPageEntry();
    if (error) {
      setStatus(error, true);
      return;
    }
    const fragmentLink = await window.HkShareLink.buildFragmentShareUrl(pageEntry);
    if (fragmentLink) {
      await navigator.clipboard.writeText(fragmentLink);
      flashBtn(els.shareLinkBtn, t("popup.copied"));
      setStatus(t("popup.statusLinkCopied"));
      return;
    }
    // fragment 放不下 → 走 GitHub raw（沿用備份設定）
    const stored = await chrome.storage?.local.get(GITHUB_SETTINGS_KEY);
    const settings = stored?.[GITHUB_SETTINGS_KEY] || {};
    if (window.HkShareLink.validateGithubSettings(settings)) {
      setStatus(t("popup.errLinkTooLong"), true);
      return;
    }
    setStatus(t("popup.statusLinkUploading"));
    const rawLink = await window.HkShareLink.commitPageToGithub(settings, pageEntry);
    await navigator.clipboard.writeText(rawLink);
    setStatus(t("popup.statusLinkCopiedGithub"));
  } catch (error) {
    console.debug("複製分享連結失敗", error);
    setStatus(error?.message || t("popup.errLinkShare"), true);
  } finally {
    if (els.shareLinkBtn) els.shareLinkBtn.disabled = false;
  }
};

// 複製本頁筆記成 Markdown（貼到聊天／筆記軟體都可讀）。
const copyCurrentPageNotes = async () => {
  try {
    const { pageEntry, error } = await collectCurrentPageEntry();
    if (error) {
      setStatus(error, true);
      return;
    }
    const markdown = window.HkParsers.pageToMarkdown(pageEntry, {
      tags: t("notesMd.tags"),
      highlights: t("notesMd.highlights"),
      summary: t("notesMd.summary"),
      mindmap: t("notesMd.mindmap"),
    });
    await navigator.clipboard.writeText(markdown);
    flashBtn(els.copyNotesBtn, t("popup.copied"));
    setStatus(t("popup.statusNotesCopied"));
  } catch (error) {
    console.debug("複製本頁筆記失敗", error);
    setStatus(error?.message || t("popup.errNotesCopy"), true);
  }
};

// 下載本頁筆記 JSON（給也用 Highlight Keeper 的人「匯入多個 JSON」）。
const downloadCurrentPageNotes = async () => {
  try {
    const { pageEntry, error } = await collectCurrentPageEntry();
    if (error) {
      setStatus(error, true);
      return;
    }
    const payload = {
      type: "highlight-keeper-bulk",
      version: 2,
      exportedAt: Date.now(),
      pages: [pageEntry],
    };
    const name =
      (pageEntry.title || pageEntry.url)
        .replace(/[\/\\:*?"<>|]+/g, "-")
        .replace(/\s+/g, "-")
        .replace(/^-+|-+$/g, "")
        .slice(0, 60) || "notes";
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: "application/json",
    });
    const href = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = href;
    link.download = `${name}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(href);
    setStatus(t("popup.statusNotesDownloaded"));
  } catch (error) {
    console.debug("下載本頁筆記失敗", error);
    setStatus(error?.message || t("popup.errNotesDownload"), true);
  }
};

/* ── 開啟面板 / 管理所有筆記 ────────────────────────── */
const openPagePanel = async () => {
  try {
    const tab = await getActiveTab();
    if (!tab?.id) {
      setStatus(t("popup.errNoTab"), true);
      return;
    }
    const response = await sendMessageToTab(tab.id, {
      type: "OPEN_PAGE_PANEL",
      side: panelSide,
    });
    if (!response?.success) {
      throw new Error(response?.error ?? t("popup.errOpenPanel"));
    }
    window.close();
  } catch (error) {
    setStatus(error?.message || t("popup.errOpenPanel2"), true);
  }
};

/* ── 初始化 ─────────────────────────────────────────── */
const loadInitialState = async () => {
  await HkI18n.initI18n();
  HkI18n.applyDOMTranslations();
  if (els.langSelect) {
    els.langSelect.value = HkI18n.getLang();
    els.langSelect.addEventListener("change", (event) => {
      HkI18n.setLang(event.target.value);
    });
  }
  HkI18n.onLangChange(() => {
    loadPageInfo();
    if (els.langSelect) els.langSelect.value = HkI18n.getLang();
  });

  try {
    const stored = await chrome.storage?.local.get("hkPanelSide");
    panelSide = stored?.hkPanelSide === "left" ? "left" : "right";
  } catch (_e) {
    panelSide = "right";
  }
};

els.copyPromptBtn?.addEventListener("click", () => copyPrompt("combined"));
els.mindmapPromptBtn?.addEventListener("click", () => copyPrompt("mindmap"));
els.pasteApplyBtn?.addEventListener("click", applyPastedAiResponse);
els.shareLinkBtn?.addEventListener("click", shareCurrentPageLink);
els.copyNotesBtn?.addEventListener("click", copyCurrentPageNotes);
els.downloadNotesBtn?.addEventListener("click", downloadCurrentPageNotes);
els.openPanelBtn?.addEventListener("click", openPagePanel);
els.openManagerBtn?.addEventListener("click", () => {
  chrome.tabs.create({ url: chrome.runtime.getURL("manager.html") });
});

// 貼上即自動套用（setTimeout 讓 textarea 先吃到剪貼簿內容），與頁面面板一致。
// 註：右鍵選單貼上會走這條；鍵盤 Cmd/Ctrl+V 走下方 keydown（popup 收不到原生快捷鍵）。
els.pasteInput?.addEventListener("paste", () => {
  window.setTimeout(() => {
    if (els.pasteInput.value.trim()) applyPastedAiResponse();
  }, 30);
});

// macOS 的擴充 popup 沒有「編輯」選單，收不到 Cmd+A/C/X/V/Z，這裡用 JS 補回來。
els.pasteInput?.addEventListener("keydown", (event) => {
  if (!(event.metaKey || event.ctrlKey) || event.altKey) return;
  const el = els.pasteInput;
  switch (event.key.toLowerCase()) {
    case "a":
      event.preventDefault();
      el.select();
      break;
    case "c":
      event.preventDefault();
      document.execCommand("copy");
      break;
    case "x":
      event.preventDefault();
      document.execCommand("cut");
      break;
    case "z":
      event.preventDefault();
      document.execCommand(event.shiftKey ? "redo" : "undo");
      break;
    case "v":
      event.preventDefault();
      navigator.clipboard
        .readText()
        .then((clip) => {
          const text = clip || "";
          if (!text) return;
          const start = el.selectionStart ?? el.value.length;
          const end = el.selectionEnd ?? el.value.length;
          el.value = el.value.slice(0, start) + text + el.value.slice(end);
          const caret = start + text.length;
          el.selectionStart = el.selectionEnd = caret;
          if (el.value.trim()) applyPastedAiResponse();
        })
        .catch(() => {
          /* 剪貼簿空或被拒 */
        });
      break;
    default:
      break;
  }
});

chrome.storage?.onChanged.addListener((changes, areaName) => {
  if (areaName !== "local") return;
  if (changes.hkPanelSide) {
    panelSide = changes.hkPanelSide.newValue === "left" ? "left" : "right";
  }
  if (changes.hkAISettings) loadPageInfo();
});

loadInitialState();
loadPageInfo();
