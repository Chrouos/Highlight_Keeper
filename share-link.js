/* 分享連結（popup 與 manager 共用，掛 window.HkShareLink）。
   兩種形式，收件端都由 contentScript 偵測後一鍵匯入：
   1. fragment：筆記壓縮後藏在「原文網址#hk=…」——零設定、無需任何後端。
      payload 為輕量版（只留 text/note/color，不含 range），收件端靠文字
      重新定位，與 AI 貼上套用同一套規則。
   2. GitHub raw：筆記太多放不下 fragment 時，commit 單頁完整 JSON 到
      使用者現有的公開備份 repo，分享 raw.githubusercontent 連結（免 Pages）。 */
(function (root) {
  // fragment 連結總長上限：聊天軟體（Slack 等）約 4000 字會截斷。
  const FRAGMENT_LIMIT = 4000;
  const FRAGMENT_PREFIX = "#hk=";

  // ── fragment：壓縮／組連結 ──────────────────────────────────────
  const bytesToBase64Url = (bytes) => {
    let binary = "";
    bytes.forEach((b) => {
      binary += String.fromCharCode(b);
    });
    return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  };

  const deflateText = async (text) => {
    const stream = new Blob([new TextEncoder().encode(text)])
      .stream()
      .pipeThrough(new CompressionStream("deflate-raw"));
    const buffer = await new Response(stream).arrayBuffer();
    return new Uint8Array(buffer);
  };

  // 匯出格式 {url,title,tags,entries,note,mindmap} → 輕量分享 payload。
  const buildSharePayload = (pageEntry) => ({
    v: 1,
    url: pageEntry.url,
    title: pageEntry.title || "",
    tags: Array.isArray(pageEntry.tags) ? pageEntry.tags : [],
    hl: (Array.isArray(pageEntry.entries) ? pageEntry.entries : [])
      .map((entry) => ({
        text: entry?.text || "",
        note: entry?.note || "",
        color: entry?.color || "",
      }))
      .filter((item) => item.text),
    note: pageEntry.note?.note || "",
    mindmap: pageEntry.mindmap?.outline || "",
  });

  // 組 fragment 連結；超過長度上限回 null（呼叫端改走 GitHub 或提示）。
  const buildFragmentShareUrl = async (pageEntry) => {
    const payload = buildSharePayload(pageEntry);
    const token = bytesToBase64Url(await deflateText(JSON.stringify(payload)));
    let base = String(pageEntry.url || "");
    const hashIndex = base.indexOf("#");
    if (hashIndex >= 0) base = base.slice(0, hashIndex);
    const link = `${base}${FRAGMENT_PREFIX}${token}`;
    return link.length <= FRAGMENT_LIMIT ? link : null;
  };

  // ── GitHub raw：commit 單頁 JSON、回傳 raw 連結 ─────────────────
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

  const encodeContentToBase64 = (text) => {
    const bytes = new TextEncoder().encode(text);
    let binary = "";
    bytes.forEach((byte) => {
      binary += String.fromCharCode(byte);
    });
    return btoa(binary);
  };

  // Git Data API 一次 commit：blobs → tree → commit → 更新 ref。
  const commitFilesToGithub = async (settings, files, commitMessage) => {
    const repoBase = buildRepoApiBase(settings.repo);
    if (!repoBase) throw new Error("INVALID_REPO");
    const { token, branch } = settings;

    let latestSha = null;
    try {
      const refData = await githubApiRequest(
        repoBase,
        `/git/ref/heads/${encodeURIComponent(branch)}`,
        token
      );
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
      treeEntries.push({ path: file.path, mode: "100644", type: "blob", sha: blob.sha });
    }

    const tree = await githubApiRequest(repoBase, "/git/trees", token, {
      method: "POST",
      body: { ...(baseTree ? { base_tree: baseTree } : {}), tree: treeEntries },
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

  // ── 路徑／slug（確定性：同一頁永遠同一檔名） ─────────────────────
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

  // FNV-1a 32-bit → 8 位 hex
  const hashString = (str) => {
    let h = 0x811c9dc5;
    for (let i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
    }
    return ("0000000" + h.toString(16)).slice(-8);
  };

  const pageFileSlug = (url) => {
    const key = root.HkUrlKey ? root.HkUrlKey.normalizePageKey(url) : String(url || "");
    let host = "page";
    try {
      host = new URL(key).hostname.replace(/^www\./, "") || "page";
    } catch (_e) {}
    return `${safeFileSlug(host)}-${hashString(key)}`;
  };

  const validateGithubSettings = (settings) => {
    if (!settings || !settings.token) return "no-token";
    if (!settings.repo || !settings.repo.includes("/")) return "no-repo";
    if (!settings.branch) return "no-branch";
    if (!settings.path) return "no-path";
    return null;
  };

  // commit 單頁完整 JSON 到備份 repo，回傳 raw.githubusercontent 連結。
  const commitPageToGithub = async (settings, pageEntry) => {
    const invalid = validateGithubSettings(settings);
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
    const filePath = joinPath(
      githubDir(settings.path),
      `pages/${pageFileSlug(pageEntry.url)}.json`
    );
    await commitFilesToGithub(
      settings,
      [{ path: filePath, content: JSON.stringify(payload, null, 2) }],
      `share: ${pageEntry.title || pageEntry.url}`
    );
    const [owner, repo] = settings.repo.split("/").map((p) => p?.trim());
    const encodedPath = filePath
      .split("/")
      .map((seg) => encodeURIComponent(seg))
      .join("/");
    return `https://raw.githubusercontent.com/${encodeURIComponent(owner)}/${encodeURIComponent(
      repo
    )}/${encodeURIComponent(settings.branch)}/${encodedPath}`;
  };

  root.HkShareLink = {
    FRAGMENT_LIMIT,
    FRAGMENT_PREFIX,
    buildSharePayload,
    buildFragmentShareUrl,
    validateGithubSettings,
    commitPageToGithub,
  };
})(typeof window !== "undefined" ? window : globalThis);
