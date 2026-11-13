(function (root) {
  const mergeAiSettings = (current, changes) => ({ ...(current || {}), ...changes });
  const removePaletteColor = (palette, color) => palette.length > 1 ? palette.filter((item) => item !== color) : palette;
  root.HkSidebarSettings = { mergeAiSettings, removePaletteColor };
  if (!root.document || !root.chrome?.storage?.local) return;

  const $ = (id) => document.getElementById(id);
  const form = $("sidebarAiForm");
  if (!form) return;
  const fields = ["provider", "openaiKey", "openaiModel", "geminiKey", "geminiModel", "autoHighlightPrompt", "prompt"];
  const checks = ["usePreview", "selectionOnly", "deepResearch"];
  const input = (name) => $(`aiSetting-${name}`);
  const status = $("sidebarAiStatus");
  let categories = [];
  let categoriesTouched = false;
  const defaultCategories = () => [["cat.motivation", "#ece3fa"], ["cat.method", "#dfeefc"], ["cat.pros", "#dff2e0"], ["cat.cons", "#ffe6cc"], ["cat.conclusion", "#fff5b8"]].map(([name, color]) => ({ name: root.HkI18n.t(name), color }));
  const renderCategories = () => {
    const list = $("sidebarCategoryList");
    list.replaceChildren();
    categories.forEach((category, index) => {
      const row = document.createElement("div");
      row.className = "category-row";
      const color = document.createElement("input");
      color.type = "color";
      color.value = /^#[0-9a-f]{6}$/i.test(category.color || "") ? category.color : "#fff5b8";
      color.setAttribute("aria-label", root.HkI18n.t("sidebar.highlightColor"));
      color.addEventListener("input", () => { categories[index].color = color.value; categoriesTouched = true; });
      const name = document.createElement("input");
      name.type = "text";
      name.value = category.name || "";
      name.setAttribute("aria-label", root.HkI18n.t("sidebar.categoryName"));
      name.addEventListener("input", () => { categories[index].name = name.value; categoriesTouched = true; });
      const remove = document.createElement("button");
      remove.type = "button";
      remove.textContent = "×";
      remove.setAttribute("aria-label", root.HkI18n.t("sidebar.removeCategory"));
      remove.addEventListener("click", () => { categories.splice(index, 1); categoriesTouched = true; renderCategories(); });
      row.append(color, name, remove);
      list.appendChild(row);
    });
  };
  const updateVisibility = () => {
    const provider = input("provider").value;
    document.querySelectorAll("[data-ai-provider]").forEach((element) => { element.hidden = element.dataset.aiProvider !== provider; });
  };
  const load = async () => {
    const { hkAISettings = {} } = await chrome.storage.local.get("hkAISettings");
    const defaults = { provider: "chatgpt", openaiModel: "gpt-4o-mini", geminiModel: "gemini-2.0-flash" };
    fields.forEach((name) => { input(name).value = hkAISettings[name] ?? defaults[name] ?? ""; });
    checks.forEach((name) => { input(name).checked = hkAISettings[name] ?? (name === "usePreview"); });
    categories = Array.isArray(hkAISettings.categories) ? hkAISettings.categories.map((category) => ({ ...category })) : defaultCategories();
    categoriesTouched = false;
    renderCategories();
    updateVisibility();
  };
  input("provider").addEventListener("change", updateVisibility);
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    try {
      const { hkAISettings = {} } = await chrome.storage.local.get("hkAISettings");
      const changes = Object.fromEntries(fields.map((name) => [name, input(name).value]));
      if (!changes.autoHighlightPrompt && !Object.hasOwn(hkAISettings, "autoHighlightPrompt")) delete changes.autoHighlightPrompt;
      if (!changes.prompt && !Object.hasOwn(hkAISettings, "prompt")) delete changes.prompt;
      checks.forEach((name) => { changes[name] = input(name).checked; });
      if (categoriesTouched) {
        if (categories.some((category) => !category.name.trim() || !/^#[0-9a-f]{6}$/i.test(category.color || ""))) throw new Error("invalid-category");
        changes.categories = categories.map((category) => ({ ...category, name: category.name.trim() }));
      }
      await chrome.storage.local.set({ hkAISettings: mergeAiSettings(hkAISettings, changes) });
      status.textContent = root.HkI18n.t("sidebar.settingsSaved");
    } catch (error) {
      status.textContent = root.HkI18n.t(error.message === "invalid-category" ? "sidebar.invalidCategories" : "sidebar.settingsError");
    }
  });
  load().catch(() => { status.textContent = root.HkI18n.t("sidebar.settingsError"); });
  $("sidebarResetCategories").addEventListener("click", () => {
    categories = defaultCategories();
    categoriesTouched = true;
    renderCategories();
  });
  $("sidebarAddCategory").addEventListener("click", () => {
    categories.push({ name: "", color: "#fff5b8" });
    categoriesTouched = true;
    renderCategories();
  });

  const defaults = ["#fff5b8", "#ffe6cc", "#dff2e0", "#dfeefc", "#fce4ee", "#ece3fa"];
  const paletteList = $("sidebarPalette");
  const syncFrame = document.querySelector(".sync-frame");
  const resizeSyncFrame = () => {
    const content = syncFrame?.contentDocument?.querySelector(".hk-manager");
    if (content?.offsetHeight) syncFrame.style.height = `${content.scrollHeight + 8}px`;
  };
  syncFrame?.addEventListener("load", () => {
    const content = syncFrame.contentDocument?.querySelector(".hk-manager");
    if (content && root.ResizeObserver) new ResizeObserver(resizeSyncFrame).observe(content);
    resizeSyncFrame();
  });
  $("syncSettings").addEventListener("toggle", (event) => {
    if (event.currentTarget.open && syncFrame && !syncFrame.hasAttribute("src")) syncFrame.src = syncFrame.dataset.src;
    resizeSyncFrame();
  });
  const scaleInput = $("sidebarFontScale");
  const applyScale = (value) => {
    const scale = Math.min(1.35, Math.max(0.85, Number(value) || 1));
    document.documentElement.style.setProperty("--read-size", `${17 * scale}px`);
    document.documentElement.style.setProperty("--ui-size", `${13 * scale}px`);
    scaleInput.value = scale;
  };
  chrome.storage.local.get("hkPanelFontScale").then((stored) => applyScale(stored.hkPanelFontScale));
  scaleInput.addEventListener("input", () => applyScale(scaleInput.value));
  scaleInput.addEventListener("change", () => chrome.storage.local.set({ hkPanelFontScale: Number(scaleInput.value) }));
  const loadPalette = async () => {
    const saved = await chrome.storage.local.get("hkColorPalette");
    const palette = Array.isArray(saved.hkColorPalette) && saved.hkColorPalette.length ? saved.hkColorPalette : defaults;
    paletteList.replaceChildren();
    palette.forEach((color) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "palette-swatch";
      button.style.background = color;
      button.textContent = "×";
      button.title = `${root.HkI18n.t("sidebar.removeColor")} ${color}`;
      button.setAttribute("aria-label", button.title);
      button.disabled = palette.length <= 1;
      button.addEventListener("click", async () => {
        await chrome.storage.local.set({ hkColorPalette: removePaletteColor(palette, color) });
        await loadPalette();
      });
      paletteList.appendChild(button);
    });
  };
  $("sidebarAddColor").addEventListener("click", async () => {
    const saved = await chrome.storage.local.get("hkColorPalette");
    const palette = Array.isArray(saved.hkColorPalette) && saved.hkColorPalette.length ? saved.hkColorPalette : defaults;
    const color = $("sidebarNewColor").value.toLowerCase();
    if (!palette.includes(color)) await chrome.storage.local.set({ hkColorPalette: [...palette, color] });
    await loadPalette();
  });
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === "local" && changes.hkColorPalette) loadPalette();
    if (area === "local" && changes.hkPanelFontScale) applyScale(changes.hkPanelFontScale.newValue);
  });
  loadPalette().catch(() => { status.textContent = root.HkI18n.t("sidebar.settingsError"); });
})(typeof window !== "undefined" ? window : globalThis);
