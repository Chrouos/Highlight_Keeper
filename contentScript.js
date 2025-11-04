const HIGHLIGHT_CLASS = "hk-highlight";
const HIGHLIGHT_ATTR = "data-highlight-id";

const pageKey = new URL(window.location.href).href;

const storage = chrome.storage?.local;
const DEFAULT_COLOR = "#ffeb3b";
const COLOR_PALETTE = ["#ffeb3b", "#ffa726", "#81c784", "#64b5f6", "#f48fb1"];
let currentColor = DEFAULT_COLOR;
let floatingButton = null;
const FLOATING_BUTTON_ID = "hk-floating-btn";
const FLOATING_BUTTON_MARGIN = 8;
let selectionDebounceTimer = null;
const HIGHLIGHT_MENU_ID = "hk-highlight-menu";
let highlightMenu = null;
let highlightMenuEls = null;
let activeHighlight = null;
let activeHighlightId = null;
let highlightMenuStatusTimer = null;

const ensureFloatingButton = () => {
  if (floatingButton) return floatingButton;
  floatingButton = document.createElement("button");
  floatingButton.type = "button";
  floatingButton.id = FLOATING_BUTTON_ID;
  floatingButton.className = "hk-floating-btn";
  floatingButton.textContent = "HL";
  floatingButton.title = "標註選取文字";
  floatingButton.style.display = "none";
  floatingButton.addEventListener("mousedown", (event) => {
    // Prevent losing selection before highlight is applied.
    event.preventDefault();
  });
  floatingButton.addEventListener("click", async (event) => {
    event.preventDefault();
    try {
      await applyHighlight(currentColor);
    } catch (error) {
      console.debug("無法套用 highlight 按鈕動作", error);
    } finally {
      hideFloatingButton();
    }
  });
  document.body.appendChild(floatingButton);
  return floatingButton;
};

const hideFloatingButton = () => {
  if (floatingButton) {
    floatingButton.style.display = "none";
  }
};

const positionFloatingButton = (rect) => {
  const button = ensureFloatingButton();
  const { innerWidth, innerHeight } = window;
  const buttonRect = button.getBoundingClientRect();
  const width = buttonRect.width || 32;
  const height = buttonRect.height || 32;

  let top = rect.top - height - FLOATING_BUTTON_MARGIN;
  let left = rect.right - width;

  if (top < FLOATING_BUTTON_MARGIN) {
    top = rect.bottom + FLOATING_BUTTON_MARGIN;
  }

  if (left < FLOATING_BUTTON_MARGIN) {
    left = rect.left;
  }

  if (left + width > innerWidth - FLOATING_BUTTON_MARGIN) {
    left = innerWidth - width - FLOATING_BUTTON_MARGIN;
  }

  if (top + height > innerHeight - FLOATING_BUTTON_MARGIN) {
    top = innerHeight - height - FLOATING_BUTTON_MARGIN;
  }

  button.style.top = `${Math.max(FLOATING_BUTTON_MARGIN, top)}px`;
  button.style.left = `${Math.max(FLOATING_BUTTON_MARGIN, left)}px`;
  button.style.display = "flex";
};

const showFloatingButton = (range) => {
  const rect = range.getBoundingClientRect();
  if (!rect || (rect.width === 0 && rect.height === 0)) {
    hideFloatingButton();
    return;
  }
  positionFloatingButton(rect);
};

const closeHighlightMenu = () => {
  if (highlightMenuStatusTimer) {
    window.clearTimeout(highlightMenuStatusTimer);
    highlightMenuStatusTimer = null;
  }
  if (highlightMenuEls?.status) {
    highlightMenuEls.status.textContent = "";
  }
  if (highlightMenu) {
    highlightMenu.style.display = "none";
  }
  activeHighlight = null;
  activeHighlightId = null;
};

const setHighlightMenuStatus = (message, isError = false) => {
  if (!highlightMenuEls?.status) return;
  if (highlightMenuStatusTimer) {
    window.clearTimeout(highlightMenuStatusTimer);
    highlightMenuStatusTimer = null;
  }
  highlightMenuEls.status.textContent = message;
  highlightMenuEls.status.style.color = isError ? "#d93025" : "#1a73e8";
  if (message) {
    highlightMenuStatusTimer = window.setTimeout(() => {
      if (highlightMenuEls?.status) {
        highlightMenuEls.status.textContent = "";
      }
    }, 2000);
  }
};

const toHexColor = (value) => {
  if (!value) return DEFAULT_COLOR;
  if (value.startsWith("#")) {
    return value.toLowerCase();
  }
  const match = value.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (!match) return DEFAULT_COLOR;
  const [_, r, g, b] = match;
  const toHex = (component) =>
    Number(component).toString(16).padStart(2, "0").toLowerCase();
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
};

const applyColorChange = async (color) => {
  if (!activeHighlight || !activeHighlightId) return;
  const nextColor = (color || DEFAULT_COLOR).toLowerCase();
  setHighlightMetadata(activeHighlight, {
    color: nextColor,
    note: activeHighlight.dataset.hkNote ?? "",
  });
  currentColor = nextColor;
  if (highlightMenuEls?.colorInput) {
    highlightMenuEls.colorInput.value = nextColor;
  }
  try {
    await updateHighlightEntry(activeHighlightId, { color: nextColor });
    try {
      await chrome.storage?.local.set({ hkLastColor: nextColor });
    } catch (storageError) {
      console.debug("更新預設 highlight 顏色失敗", storageError);
    }
    setHighlightMenuStatus("顏色已更新");
  } catch (error) {
    console.debug("更新 highlight 顏色失敗", error);
    setHighlightMenuStatus("無法更新顏色", true);
  }
};

const handleSaveNote = async () => {
  if (!activeHighlight || !activeHighlightId || !highlightMenuEls?.noteField) {
    return;
  }
  const note = highlightMenuEls.noteField.value ?? "";
  const trimmedNote = note.trim();
  const normalizedColor =
    activeHighlight.dataset.hkColor ||
    highlightMenuEls.colorInput?.value ||
    DEFAULT_COLOR;
  setHighlightMetadata(activeHighlight, { color: normalizedColor, note: trimmedNote });
  try {
    await updateHighlightEntry(activeHighlightId, { note: trimmedNote });
    if (highlightMenuEls?.noteField) {
      highlightMenuEls.noteField.value = trimmedNote;
    }
    setHighlightMenuStatus("註解已儲存");
  } catch (error) {
    console.debug("儲存 highlight 註解失敗", error);
    setHighlightMenuStatus("無法儲存註解", true);
  }
};

const handleDeleteHighlight = async () => {
  if (!activeHighlight || !activeHighlightId) return;
  const elementToRemove = activeHighlight;
  const idToRemove = activeHighlightId;
  closeHighlightMenu();
  unwrapHighlightElement(elementToRemove);
  try {
    await deleteHighlightEntry(idToRemove);
  } catch (error) {
    console.debug("刪除 highlight 失敗", error);
  }
};

const ensureHighlightMenu = () => {
  if (highlightMenu) return highlightMenu;

  const container = document.createElement("div");
  container.id = HIGHLIGHT_MENU_ID;
  container.className = "hk-highlight-menu";
  container.style.display = "none";

  const header = document.createElement("div");
  header.className = "hk-menu-header";
  header.textContent = "標註設定";
  container.appendChild(header);

  const colorWrapper = document.createElement("div");
  colorWrapper.className = "hk-menu-section hk-menu-colors";

  const colorLabel = document.createElement("label");
  colorLabel.className = "hk-menu-label";
  colorLabel.textContent = "顏色";

  const colorInput = document.createElement("input");
  colorInput.type = "color";
  colorInput.className = "hk-menu-color-input";
  colorInput.value = DEFAULT_COLOR;
  colorLabel.appendChild(colorInput);

  const swatchGroup = document.createElement("div");
  swatchGroup.className = "hk-menu-swatches";
  COLOR_PALETTE.forEach((color) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "hk-menu-swatch";
    button.style.backgroundColor = color;
    button.dataset.color = color;
    button.title = color;
    button.addEventListener("click", () => applyColorChange(color));
    swatchGroup.appendChild(button);
  });

  colorWrapper.appendChild(colorLabel);
  colorWrapper.appendChild(swatchGroup);
  container.appendChild(colorWrapper);

  const noteWrapper = document.createElement("div");
  noteWrapper.className = "hk-menu-section";
  const noteLabel = document.createElement("label");
  noteLabel.className = "hk-menu-label";
  noteLabel.textContent = "註解";
  const noteField = document.createElement("textarea");
  noteField.className = "hk-menu-note";
  noteField.rows = 3;
  noteField.placeholder = "輸入註解...";
  noteWrapper.appendChild(noteLabel);
  noteWrapper.appendChild(noteField);
  container.appendChild(noteWrapper);

  const actions = document.createElement("div");
  actions.className = "hk-menu-actions";
  const saveNoteBtn = document.createElement("button");
  saveNoteBtn.type = "button";
  saveNoteBtn.className = "hk-menu-btn hk-menu-btn-primary";
  saveNoteBtn.textContent = "儲存註解";
  saveNoteBtn.addEventListener("click", handleSaveNote);

  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "hk-menu-btn hk-menu-btn-danger";
  deleteBtn.textContent = "刪除標註";
  deleteBtn.addEventListener("click", handleDeleteHighlight);

  actions.appendChild(saveNoteBtn);
  actions.appendChild(deleteBtn);
  container.appendChild(actions);

  const status = document.createElement("div");
  status.className = "hk-menu-status";
  container.appendChild(status);

  colorInput.addEventListener("input", (event) => {
    const next = event.target.value;
    applyColorChange(next);
  });

  container.addEventListener("mousedown", (event) => {
    event.stopPropagation();
  });

  highlightMenu = container;
  highlightMenuEls = {
    container,
    colorInput,
    noteField,
    status,
  };

  document.body.appendChild(container);
  return container;
};

const positionHighlightMenu = (highlightEl) => {
  const menu = ensureHighlightMenu();
  if (!highlightEl) return;
  const margin = 12;
  menu.style.visibility = "hidden";
  menu.style.opacity = "0";
  menu.style.display = "flex";

  window.requestAnimationFrame(() => {
    const rect = highlightEl.getBoundingClientRect();
    const menuRect = menu.getBoundingClientRect();
    let top = rect.bottom + margin;
    let left = rect.left;

    if (left + menuRect.width > window.innerWidth - margin) {
      left = window.innerWidth - menuRect.width - margin;
    }
    if (left < margin) {
      left = margin;
    }

    if (top + menuRect.height > window.innerHeight - margin) {
      top = rect.top - menuRect.height - margin;
      if (top < margin) {
        top = window.innerHeight - menuRect.height - margin;
      }
    }

    menu.style.top = `${Math.max(margin, top)}px`;
    menu.style.left = `${Math.max(margin, left)}px`;
    menu.style.visibility = "visible";
    menu.style.opacity = "1";
  });
};

const openHighlightMenu = async (highlightEl) => {
  if (!highlightEl) return;
  const id = highlightEl.getAttribute(HIGHLIGHT_ATTR);
  if (!id) return;

  ensureHighlightMenu();

  activeHighlight = highlightEl;
  activeHighlightId = id;
  setHighlightMenuStatus("");

  try {
    const record = await findHighlightEntry(id);
    const color =
      record?.color ||
      highlightEl.dataset.hkColor ||
      highlightEl.getAttribute("data-highlight-color") ||
      DEFAULT_COLOR;
    const note =
      record?.note ??
      highlightEl.dataset.hkNote ??
      highlightEl.getAttribute("data-highlight-note") ??
      "";
    if (highlightMenuEls?.colorInput) {
      highlightMenuEls.colorInput.value = toHexColor(color);
    }
    if (highlightMenuEls?.noteField) {
      highlightMenuEls.noteField.value = note;
    }
  } catch (error) {
    console.debug("讀取 highlight 設定失敗", error);
  }

  if (highlightMenu) {
    highlightMenu.style.display = "flex";
    positionHighlightMenu(highlightEl);
  }
};

const updateCurrentColor = async () => {
  if (!storage) return;
  try {
    const stored = await storage.get("hkLastColor");
    const color = stored?.hkLastColor;
    if (typeof color === "string" && color.trim()) {
      currentColor = color;
    }
  } catch (error) {
    console.debug("讀取預設顏色失敗", error);
  }
};

updateCurrentColor();

chrome.storage?.onChanged.addListener((changes, areaName) => {
  if (areaName !== "local") return;
  if (changes.hkLastColor) {
    const nextColor = changes.hkLastColor.newValue;
    if (typeof nextColor === "string" && nextColor.trim()) {
      currentColor = nextColor;
    } else {
      currentColor = DEFAULT_COLOR;
    }
  }
});

const isEditableElement = (node) => {
  if (!(node instanceof HTMLElement)) return false;
  const tag = node.tagName.toLowerCase();
  return (
    node.isContentEditable ||
    tag === "input" ||
    tag === "textarea" ||
    tag === "select"
  );
};

const getNodeXPath = (node) => {
  if (!node || node === document) {
    return "";
  }

  const parts = [];
  let currentNode = node;

  while (currentNode && currentNode !== document) {
    let index = 0;
    let sibling = currentNode;

    if (currentNode.nodeType === Node.TEXT_NODE) {
      while (sibling.previousSibling) {
        sibling = sibling.previousSibling;
        if (sibling.nodeType === Node.TEXT_NODE) {
          index += 1;
        }
      }
      parts.unshift(`text()[${index + 1}]`);
      currentNode = currentNode.parentNode;
    } else if (currentNode.nodeType === Node.ELEMENT_NODE) {
      while (sibling.previousSibling) {
        sibling = sibling.previousSibling;
        if (
          sibling.nodeType === Node.ELEMENT_NODE &&
          sibling.nodeName === currentNode.nodeName
        ) {
          index += 1;
        }
      }
      const tagName = currentNode.nodeName.toLowerCase();
      parts.unshift(`${tagName}[${index + 1}]`);
      currentNode = currentNode.parentNode;
    } else {
      currentNode = currentNode.parentNode;
    }
  }

  return `/${parts.join("/")}`;
};

const getNodeFromXPath = (xpath) => {
  if (!xpath) return null;
  const result = document.evaluate(
    xpath,
    document,
    null,
    XPathResult.FIRST_ORDERED_NODE_TYPE,
    null
  );
  return result.singleNodeValue;
};

const getNodeLength = (node) => {
  if (!node) return 0;
  if (node.nodeType === Node.TEXT_NODE) {
    return node.nodeValue?.length ?? 0;
  }
  return node.childNodes.length;
};

const clampOffset = (node, offset) => {
  const max = getNodeLength(node);
  return Math.min(Math.max(offset, 0), max);
};

const serializeRange = (range) => {
  return {
    startXPath: getNodeXPath(range.startContainer),
    startOffset: range.startOffset,
    endXPath: getNodeXPath(range.endContainer),
    endOffset: range.endOffset,
    text: range.toString(),
  };
};

const deserializeRange = (data) => {
  if (!data) return null;
  const startNode = getNodeFromXPath(data.startXPath);
  const endNode = getNodeFromXPath(data.endXPath);

  if (!startNode || !endNode) return null;

  const range = document.createRange();
  range.setStart(startNode, clampOffset(startNode, data.startOffset));
  range.setEnd(endNode, clampOffset(endNode, data.endOffset));
  return range;
};

const createHighlightElement = (color, id) => {
  const mark = document.createElement("mark");
  mark.className = HIGHLIGHT_CLASS;
  mark.setAttribute(HIGHLIGHT_ATTR, id);
  mark.style.backgroundColor = color;
  mark.style.padding = "0";
  mark.style.margin = "0";
  mark.style.color = "inherit";
  return mark;
};

const setHighlightMetadata = (element, { color, note }) => {
  if (!element) return;
  if (color) {
    element.style.backgroundColor = color;
    if (element.dataset) {
      element.dataset.hkColor = color;
    }
    element.setAttribute("data-highlight-color", color);
  }
  if (typeof note === "string") {
    const trimmed = note.trim();
    if (trimmed) {
      element.dataset.hkNote = trimmed;
      element.setAttribute("data-highlight-note", trimmed);
      element.setAttribute("title", trimmed);
    } else {
      delete element.dataset.hkNote;
      element.removeAttribute("data-highlight-note");
      element.removeAttribute("title");
    }
  }
};

const unwrapHighlightElement = (highlightEl) => {
  if (!highlightEl) return;
  const parent = highlightEl.parentNode;
  if (!parent) {
    highlightEl.remove();
    return;
  }
  while (highlightEl.firstChild) {
    parent.insertBefore(highlightEl.firstChild, highlightEl);
  }
  parent.removeChild(highlightEl);
  parent.normalize();
};

const wrapRangeWithHighlight = (range, color, id) => {
  const highlightEl = createHighlightElement(color, id);
  const extracted = range.extractContents();
  highlightEl.appendChild(extracted);
  range.insertNode(highlightEl);
  return highlightEl;
};

const getStoredHighlights = async () => {
  if (!storage) return;
  const existing = await storage.get(pageKey);
  return existing[pageKey] ?? [];
};

const setStoredHighlights = async (highlights) => {
  if (!storage) return;
  await storage.set({ [pageKey]: highlights });
};

const saveHighlight = async (entry) => {
  if (!storage) return;
  const highlights = await getStoredHighlights();
  await setStoredHighlights([...highlights, entry]);
};

const updateHighlightEntry = async (id, partial) => {
  if (!storage) return false;
  const highlights = await getStoredHighlights();
  let updated = false;
  const next = highlights.map((item) => {
    if (item.id !== id) return item;
    updated = true;
    return { ...item, ...partial };
  });
  if (updated) {
    await setStoredHighlights(next);
  }
  return updated;
};

const deleteHighlightEntry = async (id) => {
  if (!storage) return false;
  const highlights = await getStoredHighlights();
  const next = highlights.filter((item) => item.id !== id);
  if (next.length !== highlights.length) {
    await setStoredHighlights(next);
    return true;
  }
  return false;
};

const findHighlightEntry = async (id) => {
  if (!storage) return null;
  const highlights = await getStoredHighlights();
  return highlights.find((item) => item.id === id) ?? null;
};

const restoreHighlights = async () => {
  if (!storage) return;
  try {
    const saved = await storage.get(pageKey);
    const highlights = saved[pageKey] ?? [];
    for (const highlight of highlights) {
      if (!highlight?.range) continue;
      const alreadyExists = document.querySelector(
        `[${HIGHLIGHT_ATTR}="${highlight.id}"]`
      );
      if (alreadyExists) continue;

      const range = deserializeRange(highlight.range);
      if (!range || range.collapsed) continue;

      try {
        const highlightEl = wrapRangeWithHighlight(range, highlight.color, highlight.id);
        setHighlightMetadata(highlightEl, {
          color: highlight.color,
          note: highlight.note ?? "",
        });
      } catch (error) {
        // log to console for debugging but do not interrupt other highlights
        console.debug("無法還原 highlight:", highlight, error);
      }
    }
  } catch (error) {
    console.debug("載入 highlight 失敗:", error);
  }
};

const applyHighlight = async (color) => {
  const selection = window.getSelection();

  if (!selection || selection.rangeCount === 0) {
    throw new Error("請先選取要標註的文字。");
  }

  const range = selection.getRangeAt(0);

  if (selection.isCollapsed || range.collapsed) {
    throw new Error("選取內容為空，請選擇文字後再試。");
  }

  if (isEditableElement(range.commonAncestorContainer)) {
    throw new Error("無法在可編輯欄位內標註。");
  }

  const snapshot = serializeRange(range.cloneRange());
  const highlightId = `hk-${Date.now()}-${Math.floor(Math.random() * 100000)}`;

  const highlightEl = wrapRangeWithHighlight(range, color, highlightId);
  setHighlightMetadata(highlightEl, { color, note: "" });
  selection.removeAllRanges();

  await saveHighlight({
    id: highlightId,
    color,
    text: snapshot.text,
    range: snapshot,
    url: pageKey,
    createdAt: Date.now(),
    note: "",
  });
};

const handleSelectionIntent = () => {
  if (selectionDebounceTimer) {
    window.clearTimeout(selectionDebounceTimer);
  }
  selectionDebounceTimer = window.setTimeout(() => {
    const selection = window.getSelection();
    const anchorNode = selection?.anchorNode;
    const anchorElement =
      anchorNode && anchorNode.nodeType === Node.TEXT_NODE
        ? anchorNode.parentNode
        : anchorNode;

    if (
      highlightMenu &&
      highlightMenu.style.display !== "none" &&
      anchorElement instanceof Node &&
      highlightMenu.contains(anchorElement)
    ) {
      return;
    }

    if (!selection || selection.rangeCount === 0) {
      hideFloatingButton();
      return;
    }

    const range = selection.getRangeAt(0);
    const ancestor =
      range.commonAncestorContainer instanceof HTMLElement
        ? range.commonAncestorContainer
        : range.commonAncestorContainer?.parentElement;

    if (
      selection.isCollapsed ||
      range.collapsed ||
      !ancestor ||
      isEditableElement(ancestor)
    ) {
      hideFloatingButton();
      return;
    }

    closeHighlightMenu();
    showFloatingButton(range);
  }, 60);
};

const handleHighlightClick = (event) => {
  const target = event.target;
  if (!(target instanceof Element)) return;
  const highlightEl = target.closest(`.${HIGHLIGHT_CLASS}`);
  if (!highlightEl) return;
  event.preventDefault();
  event.stopPropagation();
  hideFloatingButton();
  openHighlightMenu(highlightEl).catch((error) =>
    console.debug("開啟 highlight 面板失敗", error)
  );
};

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type === "APPLY_HIGHLIGHT") {
    const color = message.color || "#ffeb3b";
    applyHighlight(color)
      .then(() => {
        currentColor = color;
        hideFloatingButton();
        closeHighlightMenu();
        sendResponse({ success: true });
      })
      .catch((error) =>
        sendResponse({ success: false, error: error?.message })
      );
    return true;
  }
  return undefined;
});

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", restoreHighlights, {
    once: true,
  });
} else {
  restoreHighlights();
}

document.addEventListener("mouseup", handleSelectionIntent);
document.addEventListener("keyup", handleSelectionIntent);
document.addEventListener("selectionchange", handleSelectionIntent);
document.addEventListener("click", handleHighlightClick, true);
document.addEventListener(
  "mousedown",
  (event) => {
    const target = event.target;
    if (floatingButton && !floatingButton.contains(target)) {
      hideFloatingButton();
    }
    if (
      highlightMenu &&
      highlightMenu.style.display !== "none" &&
      !highlightMenu.contains(target) &&
      !(activeHighlight && activeHighlight.contains(target))
    ) {
      closeHighlightMenu();
    }
  },
  true
);
window.addEventListener(
  "scroll",
  () => {
    hideFloatingButton();
    closeHighlightMenu();
  },
  true
);
window.addEventListener(
  "resize",
  () => {
    closeHighlightMenu();
  },
  true
);
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeHighlightMenu();
    hideFloatingButton();
  }
});
