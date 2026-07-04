/* 分享單頁筆記的共用邏輯（popup 與 manager 共用，避免各複製一份）。
   以一般 script 載入，掛在 window.HkShare。負責：把某一頁存成小檔上傳，
   一併帶上 viewer.html，回傳可公開瀏覽的 GitHub Pages 連結。
   低階 GitHub commit 沿用 manager.js 相同的 Git Data API 流程（blobs→tree→commit→ref）。 */
(function (root) {
  const GITHUB_API_VERSION = "2022-11-28";

  const githubHeaders = (token) => ({
    Authorization: `Bearer ${token}`,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": GITHUB_API_VERSION,
  });

  const buildRepoApiBase = (repo) => {
    if (typeof repo !== "string") return null;
    const [owner, name] = repo.split("/").map((part) => part?.trim());
    if (!owner || !name) return null;
    return `https://api.github.com/repos/${encodeURIComponent(
      owner
    )}/${encodeURIComponent(name)}`;
  };

  // 把 GitHub 錯誤回應整理成精簡可行動的訊息。
  const describeGithubError = (status, errorText) => {
    const raw = (errorText || "").trim();
    if (/^</.test(raw)) return `GitHub 回應異常（HTTP ${status || 400}）`;
    try {
      const json = JSON.parse(raw);
      const parts = [];
      if (json?.message) parts.push(json.message);
      if (Array.isArray(json?.errors) && json.errors.length) {
        const detail = json.errors
          .map((e) => e?.message || e?.code || e?.field)
          .filter(Boolean)
          .join("; ");
        if (detail) parts.push(detail);
      }
      if (parts.length) {
        return `${parts.join(" — ")}${status ? `（${status}）` : ""}`;
      }
    } catch (_e) {
      /* 非 JSON */
    }
    return raw.slice(0, 300) || `HTTP ${status}`;
  };

  const githubApiRequest = async (
    repoBase,
    path,
    token,
    { method = "GET", body } = {}
  ) => {
    const response = await fetch(`${repoBase}${path}`, {
      method,
      headers: {
        ...githubHeaders(token),
        ...(body ? { "Content-Type": "application/json" } : {}),
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (!response.ok) {
      const text = await response.text();
      const error = new Error(describeGithubError(response.status, text));
      error.status = response.status;
      throw error;
    }
    return response.json();
  };

  // UTF-8 → base64（相容中文等多位元組字元）。
  const encodeContentToBase64 = (text) => {
    const bytes = new TextEncoder().encode(text);
    let binary = "";
    bytes.forEach((byte) => {
      binary += String.fromCharCode(byte);
    });
    return btoa(binary);
  };

  // 經由 Git Data API 一次 commit 多個檔案：blobs → tree → commit → 更新 ref。
  const commitFilesToGithub = async (settings, files, commitMessage) => {
    const repoBase = buildRepoApiBase(settings.repo);
    if (!repoBase) throw new Error("INVALID_REPO");
    const { token, branch } = settings;
    const refPath = `/git/ref/heads/${encodeURIComponent(branch)}`;

    let latestSha = null;
    try {
      const refData = await githubApiRequest(repoBase, refPath, token);
      latestSha = refData?.object?.sha || null;
    } catch (error) {
      if (error?.status !== 404) throw error;
    }
    let baseTree = null;
    if (latestSha) {
      const commitData = await githubApiRequest(
        repoBase,
        `/git/commits/${latestSha}`,
        token
      );
      baseTree = commitData?.tree?.sha || null;
    }

    const treeEntries = [];
    for (const file of files) {
      const blob = await githubApiRequest(repoBase, "/git/blobs", token, {
        method: "POST",
        body: { content: encodeContentToBase64(file.content), encoding: "base64" },
      });
      treeEntries.push({
        path: file.path,
        mode: "100644",
        type: "blob",
        sha: blob.sha,
      });
    }

    const tree = await githubApiRequest(repoBase, "/git/trees", token, {
      method: "POST",
      body: {
        ...(baseTree ? { base_tree: baseTree } : {}),
        tree: treeEntries,
      },
    });

    const commit = await githubApiRequest(repoBase, "/git/commits", token, {
      method: "POST",
      body: {
        message:
          commitMessage || `share: highlight-keeper (${new Date().toISOString()})`,
        tree: tree.sha,
        parents: latestSha ? [latestSha] : [],
      },
    });

    if (latestSha) {
      await githubApiRequest(
        repoBase,
        `/git/refs/heads/${encodeURIComponent(branch)}`,
        token,
        { method: "PATCH", body: { sha: commit.sha } }
      );
    } else {
      await githubApiRequest(repoBase, "/git/refs", token, {
        method: "POST",
        body: { ref: `refs/heads/${branch}`, sha: commit.sha },
      });
    }
  };

  // ── 路徑 / slug 工具 ────────────────────────────────────────────
  const githubDir = (path) => {
    const parts = String(path || "").split("/").filter(Boolean);
    parts.pop();
    return parts.join("/");
  };
  const joinPath = (dir, rel) => (dir ? `${dir}/${rel}` : rel);

  const safeFileSlug = (name) =>
    String(name)
      .replace(/[\/\\:*?"<>| -]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 60) || "page";

  const normalizeUrlKey = (url) =>
    root.HkUrlKey ? root.HkUrlKey.normalizePageKey(url) : String(url || "");

  // FNV-1a 32-bit → 8 位 hex：確定性、跨裝置一致，讓檔名單由網址就能算出。
  const hashString = (str) => {
    let h = 0x811c9dc5;
    for (let i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
    }
    return ("0000000" + h.toString(16)).slice(-8);
  };

  const pageFileSlug = (url) => {
    const key = normalizeUrlKey(url);
    let host = "page";
    try {
      host = new URL(key).hostname.replace(/^www\./, "") || "page";
    } catch (_e) {}
    return `${safeFileSlug(host)}-${hashString(key)}`;
  };

  const pageBackupPath = (settings, url) =>
    joinPath(githubDir(settings.path), `pages/${pageFileSlug(url)}.json`);

  // GitHub Pages 上 viewer.html 的網址（含 <owner>.github.io repo 的根目錄特例）。
  const buildViewerShareUrl = (settings, filePath) => {
    const [owner, repo] = settings.repo.split("/").map((p) => p?.trim());
    const base =
      repo.toLowerCase() === `${owner.toLowerCase()}.github.io`
        ? `https://${owner}.github.io/viewer.html`
        : `https://${owner}.github.io/${repo}/viewer.html`;
    const hash =
      `o=${encodeURIComponent(owner)}` +
      `&r=${encodeURIComponent(repo)}` +
      `&b=${encodeURIComponent(settings.branch)}` +
      `&f=${encodeURIComponent(filePath)}`;
    return `${base}#${hash}`;
  };

  // 讀取擴充內封裝的 viewer.html 原始碼，一併 commit 確保 repo 有檢視器。
  const getBundledViewerHtml = async () => {
    const res = await fetch(chrome.runtime.getURL("viewer.html"));
    if (!res.ok) throw new Error(`viewer.html ${res.status}`);
    return res.text();
  };

  // settings 缺漏檢查，回傳錯誤碼字串或 null（讓呼叫端各自翻成 i18n 文案）。
  const validateSettings = (settings) => {
    if (!settings || !settings.token) return "no-token";
    if (!settings.repo || !settings.repo.includes("/")) return "no-repo";
    if (!settings.branch) return "no-branch";
    if (!settings.path) return "no-path";
    return null;
  };

  // 分享某一頁：只 commit 該頁小檔（+ viewer.html），回傳 { link, filePath }。
  // pageEntry 為匯出格式物件 { url, title, tags, entries, note, mindmap }。
  const sharePage = async ({ settings, pageEntry, commitMessage } = {}) => {
    const invalid = validateSettings(settings);
    if (invalid) {
      const err = new Error(invalid);
      err.code = invalid;
      throw err;
    }
    const payload = {
      type: "highlight-keeper-bulk",
      version: 2,
      exportedAt: Date.now(),
      shared: true,
      pages: [pageEntry],
    };
    const filePath = pageBackupPath(settings, pageEntry.url);
    const files = [{ path: filePath, content: JSON.stringify(payload, null, 2) }];
    try {
      files.push({ path: "viewer.html", content: await getBundledViewerHtml() });
    } catch (_e) {
      // 取不到封裝檔就略過 viewer.html（多半已存在於 repo）
    }
    await commitFilesToGithub(settings, files, commitMessage);
    return { link: buildViewerShareUrl(settings, filePath), filePath };
  };

  root.HkShare = {
    validateSettings,
    pageFileSlug,
    pageBackupPath,
    buildViewerShareUrl,
    getBundledViewerHtml,
    commitFilesToGithub,
    sharePage,
  };
})(typeof window !== "undefined" ? window : globalThis);
