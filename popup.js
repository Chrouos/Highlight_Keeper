const colorInput = document.getElementById("color");
const highlightBtn = document.getElementById("highlightBtn");
const statusEl = document.querySelector(".status");
const swatches = [...document.querySelectorAll(".quick-swatches button")];
const DEFAULT_COLOR = "#ffeb3b";

const persistColor = async (color) => {
  try {
    await chrome.storage?.local.set({ hkLastColor: color });
  } catch (error) {
    console.debug("儲存顏色失敗", error);
  }
};

const loadInitialColor = async () => {
  try {
    const stored = await chrome.storage?.local.get("hkLastColor");
    const color = stored?.hkLastColor;
    if (color) {
      colorInput.value = color;
    } else {
      colorInput.value = DEFAULT_COLOR;
    }
  } catch (error) {
    colorInput.value = DEFAULT_COLOR;
  }
};

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "#d93025" : "#1a73e8";
}

const handleColorUpdate = (color) => {
  colorInput.value = color;
  persistColor(color);
  setStatus(`已選擇顏色 ${color}`);
};

colorInput.addEventListener("input", (event) => {
  handleColorUpdate(event.target.value);
});

swatches.forEach((btn) => {
  btn.addEventListener("click", () => {
    const color = btn.dataset.color;
    handleColorUpdate(color);
  });
});

highlightBtn.addEventListener("click", async () => {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab?.id) {
      setStatus("找不到目前的分頁。", true);
      return;
    }

    setStatus("正在標註...");
    highlightBtn.disabled = true;

    const response = await chrome.tabs.sendMessage(tab.id, {
      type: "APPLY_HIGHLIGHT",
      color: colorInput.value,
    });

    if (!response?.success) {
      throw new Error(response?.error ?? "標註失敗");
    }

    setStatus("標註完成！");
  } catch (error) {
    const fallbackMessage =
      chrome.runtime.lastError?.message || error?.message || "無法標註。";
    setStatus(fallbackMessage, true);
  } finally {
    highlightBtn.disabled = false;
  }
});

loadInitialColor();
