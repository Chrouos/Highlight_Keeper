/* 自動同步：定時把本機筆記與 GitHub 備份做逐頁三方比對。
   service worker 與 manager 共用，掛在 HkAutoSync（依賴 parsers.js、share-link.js）。

   每頁比較三個雜湊：本機、遠端、上次同步後的基準（base）。
   - 只有一邊變動 → 自動帶到另一邊。
   - 兩邊都變動 → 衝突：這一頁兩邊都不動，等使用者在筆記總覽選擇。
   沒有任何變動時只打一個 API，不下載也不上傳。 */
(function (root) {
  const SETTINGS_KEY = "hkGithubSyncSettings";
  const STATE_KEY = "hkAutoSyncState";
  const PAGE_META_KEY = "__hk_page_meta__";
  const GENERATED_NOTES_KEY = "hkGeneratedNotes";
  const MINDMAP_KEY = "hkMindmaps";
  const PERSONAL_NOTES_KEY = "hkPersonalNotes";
  const ALARM_NAME = "hk-auto-sync";
  const DEFAULT_MINUTES = 15;
  const UNTAGGED_KEY = "__hk_untagged__";
  // 一次刪除超過這個量就當異常（例如遠端檔案被清空），不自動套用。
  const MASS_DELETE_MIN = 3;
  const MASS_DELETE_RATIO = 0.2;

  const clampMinutes = (value) => {
    const minutes = Math.round(Number(value));
    if (!Number.isFinite(minutes) || minutes < 1) return DEFAULT_MINUTES;
    return Math.min(minutes, 1440);
  };

  const isPageKey = (key) => root.HkParsers.isHttpUrl(key);

  const exportNote = (note) => root.HkParsers.exportNote(note);

  /* ── 資料轉換 ─────────────────────────────────────── */
  const collectLocalPages = (all) => root.HkParsers.collectBackupPages(all);

  const parseRemotePages = (rawText) => {
    const parsed = JSON.parse(rawText);
    const pages = Array.isArray(parsed?.pages) ? parsed.pages : Array.isArray(parsed) ? parsed : null;
    if (!pages) throw new Error("not a backup");
    const seen = new Set();
    return root.HkParsers.normalizeBulkPages(pages).filter((page) => {
      if (seen.has(page.url)) return false;
      seen.add(page.url);
      return true;
    });
  };

  const canonicalJson = (value) => {
    if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;
    if (value && typeof value === "object") {
      const keys = Object.keys(value).filter((key) => value[key] !== undefined).sort();
      return `{${keys.map((key) => `${JSON.stringify(key)}:${canonicalJson(value[key])}`).join(",")}}`;
    }
    return JSON.stringify(value ?? null);
  };

  const hexDigest = async (algorithm, bytes) => {
    const digest = await crypto.subtle.digest(algorithm, bytes);
    return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("");
  };

  const hashPage = (page) =>
    hexDigest(
      "SHA-256",
      new TextEncoder().encode(
        canonicalJson({
          entries: page.entries || [],
          title: page.title || "",
          tags: page.tags || [],
          note: exportNote(page.note),
          mindmap: page.mindmap || null,
          personalNote: page.personalNote || "",
        })
      )
    );

  const hashPages = async (pages) => {
    const hashes = {};
    await Promise.all(pages.map(async (page) => { hashes[page.url] = await hashPage(page); }));
    return hashes;
  };

  const digestHashes = (hashes) =>
    hexDigest(
      "SHA-256",
      new TextEncoder().encode(Object.keys(hashes).sort().map((url) => `${url} ${hashes[url]}`).join("\n"))
    );

  // 與 git hash-object 相同的算法：內容沒變就不用重傳這個檔案。
  const gitBlobSha = (text) => {
    const body = new TextEncoder().encode(text);
    const head = new TextEncoder().encode(`blob ${body.length}\0`);
    const bytes = new Uint8Array(head.length + body.length);
    bytes.set(head);
    bytes.set(body, head.length);
    return hexDigest("SHA-1", bytes);
  };

  /* ── 分類檔（手動上傳與自動同步共用） ──────────────── */
  const githubDir = (path) => {
    const parts = String(path || "").split("/").filter(Boolean);
    parts.pop();
    return parts.join("/");
  };
  const joinPath = (dir, rel) => (dir ? `${dir}/${rel}` : rel);
  const safeFileSlug = (name) =>
    String(name)
      .replace(/[\/\\:*?"<>|\u0000-\u001f]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 60) || "category";

  // 完整備份＋各標籤分類檔＋index.json。分類檔不放時間戳，沒變的頁面就不會產生新內容。
  const buildCategorizedFiles = (settings, payload) => {
    const dir = githubDir(settings.path);
    const fullName = String(settings.path).split("/").pop() || "highlight-keeper.json";
    const groups = new Map();
    const add = (key, page) => {
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(page);
    };
    (payload.pages || []).forEach((page) => {
      const tags = (Array.isArray(page.tags) ? page.tags : []).map((tag) => String(tag).trim()).filter(Boolean);
      if (!tags.length) add(UNTAGGED_KEY, page);
      else tags.forEach((tag) => add(tag, page));
    });
    const usedSlugs = new Set();
    const categories = [];
    const files = [{ path: settings.path, content: JSON.stringify(payload, null, 2) }];
    for (const [tag, pages] of groups) {
      const isUntagged = tag === UNTAGGED_KEY;
      let slug = isUntagged ? "_untagged" : safeFileSlug(tag);
      while (usedSlugs.has(slug)) slug += "-1";
      usedSlugs.add(slug);
      const rel = `categories/${slug}.json`;
      files.push({
        path: joinPath(dir, rel),
        content: JSON.stringify(
          { type: "highlight-keeper-bulk", version: 2, category: isUntagged ? null : tag, pages },
          null,
          2
        ),
      });
      categories.push({
        tag: isUntagged ? null : tag,
        untagged: isUntagged || undefined,
        file: rel,
        pageCount: pages.length,
      });
    }
    categories.sort((a, b) => (b.pageCount || 0) - (a.pageCount || 0) || a.file.localeCompare(b.file));
    const index = {
      type: "highlight-keeper-index",
      version: 1,
      totalPages: (payload.pages || []).length,
      fullBackup: fullName,
      categories,
    };
    files.push({ path: joinPath(dir, "index.json"), content: JSON.stringify(index, null, 2) });
    return { files, index };
  };

  /* ── 比對 ─────────────────────────────────────────── */
  const classify = (local, remote, base) => {
    if (local === remote) return "same";
    if (base === undefined) return local && remote ? "conflict" : local ? "push" : "pull";
    if (local === base) return remote ? "pull" : "delete-local";
    if (remote === base) return local ? "push" : "delete-remote";
    return "conflict";
  };

  const conflictKind = (local, remote) =>
    !local ? "local-deleted" : !remote ? "remote-deleted" : "both-changed";

  const buildPlan = ({ local, remote, base = {}, resolutions = {} }) => {
    const urls = [...new Set([...Object.keys(local), ...Object.keys(remote), ...Object.keys(base)])].sort();
    const actions = {};
    const conflicts = [];
    urls.forEach((url) => {
      const l = local[url] ?? null;
      const r = remote[url] ?? null;
      const choice = resolutions[url];
      if (l !== r && choice === "local") actions[url] = l ? "push" : "delete-remote";
      else if (l !== r && choice === "remote") actions[url] = r ? "pull" : "delete-local";
      else actions[url] = classify(l, r, base[url]);
      if (actions[url] === "conflict") conflicts.push({ url, kind: conflictKind(l, r) });
    });
    // 大量刪除多半是異常（遠端被清空、本機資料遺失），刪除不可逆，一律交給使用者決定。
    const baseCount = Object.keys(base).length;
    ["delete-local", "delete-remote"].forEach((action) => {
      const targets = urls.filter((url) => actions[url] === action && !resolutions[url]);
      if (targets.length > MASS_DELETE_MIN && targets.length > baseCount * MASS_DELETE_RATIO) {
        targets.forEach((url) => {
          actions[url] = "conflict";
          conflicts.push({ url, kind: "mass-delete" });
        });
      }
    });
    return { actions, conflicts };
  };

  /* ── 本機寫入 ─────────────────────────────────────── */
  const writeLocalChanges = async (pulled, deleted, all) => {
    if (!pulled.length && !deleted.length) return;
    const meta = { ...(all[PAGE_META_KEY] || {}) };
    const notes = { ...(all[GENERATED_NOTES_KEY] || {}) };
    const mindmaps = { ...(all[MINDMAP_KEY] || {}) };
    const personal = { ...(all[PERSONAL_NOTES_KEY] || {}) };
    const updates = {};
    const removals = [...deleted];
    deleted.forEach((url) => {
      delete meta[url];
      delete notes[url];
      delete mindmaps[url];
      delete personal[url];
    });
    pulled.forEach((page) => {
      if (page.entries.length) updates[page.url] = page.entries;
      else removals.push(page.url);
      meta[page.url] = { ...(meta[page.url] || {}), title: page.title, tags: page.tags };
      const note = exportNote(page.note);
      if (note) notes[page.url] = note;
      else delete notes[page.url];
      if (page.mindmap) mindmaps[page.url] = page.mindmap;
      else delete mindmaps[page.url];
      if (typeof page.personalNote === "string") {
        personal[page.url] = { text: page.personalNote, updatedAt: Date.now() };
      } else {
        delete personal[page.url];
      }
    });
    await chrome.storage.local.set({
      ...updates,
      [PAGE_META_KEY]: meta,
      [GENERATED_NOTES_KEY]: notes,
      [MINDMAP_KEY]: mindmaps,
      [PERSONAL_NOTES_KEY]: personal,
    });
    if (removals.length) await chrome.storage.local.remove(removals);
  };

  /* ── GitHub 讀取 ──────────────────────────────────── */
  const githubGet = async (settings, path) => {
    const github = root.HkShareLink;
    const base = github.buildRepoApiBase(settings.repo);
    if (!base) throw new Error(`Invalid repo: ${settings.repo}`);
    return github.githubApiRequest(base, path, settings.token);
  };

  const notFoundToNull = (error) => {
    if (error?.status === 404) return null;
    throw error;
  };

  const encodePath = (path) => String(path).split("/").filter(Boolean).map(encodeURIComponent).join("/");

  const decodeBase64 = (text) => {
    const binary = atob(String(text).replace(/\s/g, ""));
    return new TextDecoder().decode(Uint8Array.from(binary, (c) => c.charCodeAt(0)));
  };

  const fetchRemoteText = async (settings, commitSha) => {
    const file = await githubGet(
      settings,
      `/contents/${encodePath(settings.path)}?ref=${encodeURIComponent(commitSha)}`
    ).catch(notFoundToNull);
    if (!file) return null;
    if (Array.isArray(file)) throw new Error(`${settings.path} is a folder`);
    if (file.content) return decodeBase64(file.content);
    // 超過 1MB 時 contents API 不帶內容，改讀 blob。
    const blob = await githubGet(settings, `/git/blobs/${file.sha}`);
    return decodeBase64(blob.content);
  };

  // 只留下內容真的變動的檔案，並刪掉已經沒有對應標籤的分類檔。
  const filesToCommit = async (settings, files, commitSha) => {
    if (!commitSha) return { files, deletions: [] };
    const tree = await githubGet(settings, `/git/trees/${commitSha}?recursive=1`);
    if (tree.truncated) return { files, deletions: [] };
    const existing = new Map(tree.tree.filter((item) => item.type === "blob").map((item) => [item.path, item.sha]));
    const changed = [];
    for (const file of files) {
      if (existing.get(file.path) !== (await gitBlobSha(file.content))) changed.push(file);
    }
    const prefix = joinPath(githubDir(settings.path), "categories/");
    const keep = new Set(files.map((file) => file.path));
    const deletions = [...existing.keys()].filter((path) => path.startsWith(prefix) && !keep.has(path));
    return { files: changed, deletions };
  };

  /* ── 同步一輪 ─────────────────────────────────────── */
  const readState = async () => (await chrome.storage.local.get(STATE_KEY))[STATE_KEY] || {};

  const updateBadge = (state) => {
    if (!root.chrome?.action?.setBadgeText) return;
    const alert = state.conflicts?.length || (state.status === "error" && state.errorCode !== "remote-moved");
    chrome.action.setBadgeText({ text: alert ? "!" : "" }).catch(() => {});
    if (alert) chrome.action.setBadgeBackgroundColor?.({ color: "#c53030" }).catch(() => {});
  };

  const saveState = async (state) => {
    await chrome.storage.local.set({ [STATE_KEY]: state });
    updateBadge(state);
    return state;
  };

  let running = false;

  const runCycle = async ({ manual = false } = {}) => {
    if (running) return { status: "busy" };
    running = true;
    try {
      const stored = await chrome.storage.local.get([SETTINGS_KEY, STATE_KEY]);
      const settings = stored[SETTINGS_KEY] || {};
      const previous = stored[STATE_KEY] || {};
      if (!manual && !settings.autoSync) return { status: "disabled" };
      const invalid = root.HkShareLink.validateGithubSettings(settings);
      if (invalid) {
        return await saveState({ ...previous, status: "error", errorCode: invalid, errorMessage: "", lastRunAt: Date.now() });
      }
      const target = `${settings.repo}|${settings.branch}|${settings.path}`;
      // 換了備份位置，舊基準就沒有意義。
      const state = previous.target === target ? previous : { target };
      const base = state.base || {};
      const resolutions = state.resolutions || {};
      try {
        return await syncOnce(settings, state, base, resolutions);
      } catch (error) {
        console.debug("自動同步失敗", error);
        return await saveState({
          ...state,
          target,
          status: "error",
          // 其他裝置剛好同時推送：不算失敗，下一輪重新比對即可。
          errorCode: error?.status === 422 ? "remote-moved" : error?.code || "",
          errorMessage: error?.message || String(error),
          lastRunAt: Date.now(),
        });
      }
    } finally {
      running = false;
    }
  };

  const syncOnce = async (settings, state, base, resolutions) => {
    const now = Date.now();
    const firstAll = await chrome.storage.local.get(null);
    const localPages = collectLocalPages(firstAll);
    const localHashes = await hashPages(localPages);
    const localDigest = await digestHashes(localHashes);

    const ref = await githubGet(settings, `/git/ref/heads/${encodeURIComponent(settings.branch)}`).catch(notFoundToNull);
    const commitSha = ref?.object?.sha || null;

    // 兩邊都跟上一輪一樣：什麼都不用做，也不下載。
    if (
      state.commitSha === commitSha &&
      state.localDigest === localDigest &&
      state.status !== "error" &&
      !Object.keys(resolutions).length
    ) {
      return saveState({ ...state, lastRunAt: now, pulled: 0, pushed: 0, deleted: 0 });
    }

    const remoteText = commitSha ? await fetchRemoteText(settings, commitSha) : null;
    let remotePages = [];
    if (remoteText !== null) {
      try {
        remotePages = parseRemotePages(remoteText);
      } catch (_error) {
        // 看不懂遠端檔案時絕不動任何資料。
        const error = new Error(`${settings.path}: invalid backup file`);
        error.code = "remote-invalid";
        throw error;
      }
    }
    const remoteHashes = await hashPages(remotePages);
    const { actions, conflicts } = buildPlan({ local: localHashes, remote: remoteHashes, base, resolutions });
    const remoteByUrl = new Map(remotePages.map((page) => [page.url, page]));
    const localByUrl = new Map(localPages.map((page) => [page.url, page]));

    // 網路請求期間使用者可能又改了標註：這些頁留到下一輪，不覆蓋剛寫的內容。
    const writesLocally = Object.values(actions).some((action) => action === "pull" || action === "delete-local");
    const freshAll = writesLocally ? await chrome.storage.local.get(null) : null;
    const freshHashes = writesLocally ? await hashPages(collectLocalPages(freshAll)) : null;
    const unchangedLocally = (url) => (freshHashes[url] ?? null) === (localHashes[url] ?? null);
    const deferred = new Set(
      Object.keys(actions).filter((url) => ["pull", "delete-local"].includes(actions[url]) && !unchangedLocally(url))
    );
    const urlsWith = (action) => Object.keys(actions).filter((url) => actions[url] === action && !deferred.has(url));
    const pullUrls = urlsWith("pull");
    const deleteLocalUrls = urlsWith("delete-local");
    const pushUrls = urlsWith("push");
    const deleteRemoteUrls = urlsWith("delete-remote");

    await writeLocalChanges(pullUrls.map((url) => remoteByUrl.get(url)), deleteLocalUrls, freshAll);

    let nextCommitSha = commitSha;
    if (pushUrls.length || deleteRemoteUrls.length) {
      const merged = new Map(remoteByUrl);
      pushUrls.forEach((url) => merged.set(url, localByUrl.get(url)));
      deleteRemoteUrls.forEach((url) => merged.delete(url));
      const pages = [...merged.values()]
        .map((page) => ({ ...page, note: exportNote(page.note) }))
        .sort((a, b) => a.url.localeCompare(b.url));
      const payload = { type: "highlight-keeper-bulk", version: 2, exportedAt: now, pages };
      const { files, index } = buildCategorizedFiles(settings, payload);
      const commit = await filesToCommit(settings, files, commitSha);
      nextCommitSha = await root.HkShareLink.commitFilesToGithub(
        settings,
        commit.files,
        `sync: highlight-keeper（更新 ${pushUrls.length} 頁，移除 ${deleteRemoteUrls.length} 頁，共 ${index.totalPages} 頁）`,
        { parentSha: commitSha, deletions: commit.deletions }
      );
    }

    const nextBase = { ...base };
    Object.entries(actions).forEach(([url, action]) => {
      if (deferred.has(url) || action === "conflict") return;
      if (action === "delete-local" || action === "delete-remote") delete nextBase[url];
      else if (action === "pull") nextBase[url] = remoteHashes[url];
      else if (localHashes[url]) nextBase[url] = localHashes[url];
      else delete nextBase[url];
    });

    // 只記下這一輪處理過的內容：同步途中才改的頁面雜湊會不同，下一輪就不會被跳過。
    const handledHashes = { ...localHashes };
    pullUrls.forEach((url) => { handledHashes[url] = remoteHashes[url]; });
    deleteLocalUrls.forEach((url) => { delete handledHashes[url]; });
    return saveState({
      ...state,
      base: nextBase,
      commitSha: nextCommitSha,
      localDigest: await digestHashes(handledHashes),
      conflicts,
      resolutions: {},
      status: conflicts.length ? "conflict" : "ok",
      errorCode: "",
      errorMessage: "",
      pulled: pullUrls.length,
      pushed: pushUrls.length,
      deleted: deleteLocalUrls.length + deleteRemoteUrls.length,
      lastRunAt: now,
      lastSuccessAt: now,
    });
  };

  // 使用者在筆記總覽決定衝突怎麼處理：choices = { url: "local" | "remote" }。
  const resolveConflicts = async (choices) => {
    const state = await readState();
    await chrome.storage.local.set({
      [STATE_KEY]: { ...state, resolutions: { ...(state.resolutions || {}), ...choices } },
    });
    return runCycle({ manual: true });
  };

  /* ── 排程 ─────────────────────────────────────────── */
  const applySchedule = async () => {
    const stored = await chrome.storage.local.get(SETTINGS_KEY);
    const settings = stored[SETTINGS_KEY] || {};
    if (!settings.autoSync) {
      await chrome.alarms.clear(ALARM_NAME);
      return;
    }
    const period = clampMinutes(settings.autoSyncMinutes);
    const existing = await chrome.alarms.get(ALARM_NAME);
    if (existing?.periodInMinutes === period) return;
    await chrome.alarms.create(ALARM_NAME, { delayInMinutes: period, periodInMinutes: period });
  };

  root.HkAutoSync = {
    SETTINGS_KEY,
    STATE_KEY,
    ALARM_NAME,
    DEFAULT_MINUTES,
    clampMinutes,
    exportNote,
    collectLocalPages,
    buildCategorizedFiles,
    classify,
    buildPlan,
    runCycle,
    resolveConflicts,
    applySchedule,
    updateBadge,
    readState,
  };
})(typeof window !== "undefined" ? window : globalThis);
