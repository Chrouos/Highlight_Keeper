// Auto-sync against an in-memory GitHub: two "devices" share one remote.
const assert = require("node:assert/strict");
const nodeCrypto = require("node:crypto");
require("../shared.js");
require("../parsers.js");
require("../share-link.js");
require("../auto-sync.js");
const S = globalThis.HkAutoSync;

const URL_A = "https://example.com/a";
const URL_B = "https://example.com/b";
const URL_C = "https://example.com/c";

/* ── fake GitHub ─────────────────────────────────────── */
const github = { blobs: new Map(), trees: new Map(), commits: new Map(), head: null, calls: [], uploads: [], fail: null };
const sha1 = (text) => nodeCrypto.createHash("sha1").update(text).digest("hex");
const gitSha = (content) => sha1(Buffer.concat([Buffer.from(`blob ${Buffer.byteLength(content)}\0`), Buffer.from(content)]));
const respond = (status, body) => ({ ok: status < 400, status, json: async () => body, text: async () => JSON.stringify(body) });
const headTree = (sha) => (sha ? github.trees.get(github.commits.get(sha).tree) : new Map());
globalThis.fetch = async (url, { method = "GET", body } = {}) => {
  const path = decodeURIComponent(new URL(url).pathname + new URL(url).search).replace("/repos/me/backup", "");
  github.calls.push(`${method} ${path}`);
  if (github.fail) { const f = github.fail; github.fail = null; return respond(f, { message: "boom" }); }
  const data = body ? JSON.parse(body) : null;
  if (method === "GET" && path.startsWith("/git/ref/")) return github.head ? respond(200, { object: { sha: github.head } }) : respond(404, {});
  if (method === "GET" && path.startsWith("/contents/")) {
    const [file, query] = path.slice("/contents/".length).split("?ref=");
    const blobSha = headTree(query).get(file);
    if (!blobSha) return respond(404, {});
    return respond(200, { sha: blobSha, content: Buffer.from(github.blobs.get(blobSha)).toString("base64") });
  }
  if (method === "GET" && path.startsWith("/git/trees/")) {
    const commit = path.slice("/git/trees/".length).split("?")[0];
    return respond(200, { truncated: false, tree: [...headTree(commit)].map(([p, sha]) => ({ path: p, sha, type: "blob" })) });
  }
  if (method === "GET" && path.startsWith("/git/commits/")) return respond(200, { tree: { sha: github.commits.get(path.slice("/git/commits/".length)).tree } });
  if (method === "POST" && path === "/git/blobs") {
    const content = Buffer.from(data.content, "base64").toString("utf8");
    const sha = gitSha(content);
    github.blobs.set(sha, content);
    github.uploads.push(content);
    return respond(201, { sha });
  }
  if (method === "POST" && path === "/git/trees") {
    const tree = new Map(data.base_tree ? github.trees.get(data.base_tree) : []);
    data.tree.forEach((item) => { if (item.sha === null) tree.delete(item.path); else tree.set(item.path, item.sha); });
    const sha = sha1(`tree${JSON.stringify([...tree])}${github.trees.size}`);
    github.trees.set(sha, tree);
    return respond(201, { sha });
  }
  if (method === "POST" && path === "/git/commits") {
    const sha = sha1(`commit${data.tree}${github.commits.size}`);
    github.commits.set(sha, { tree: data.tree, parents: data.parents });
    return respond(201, { sha });
  }
  if (method === "PATCH" && path.startsWith("/git/refs/heads/")) {
    if (!github.commits.get(data.sha).parents.includes(github.head)) return respond(422, { message: "Update is not a fast forward" });
    github.head = data.sha;
    return respond(200, {});
  }
  if (method === "POST" && path === "/git/refs") { github.head = data.sha; return respond(201, {}); }
  throw new Error(`unexpected request: ${method} ${path}`);
};
const remoteFile = () => JSON.parse(github.blobs.get(headTree(github.head).get("backups/hk.json")));
const remoteUrls = () => remoteFile().pages.map((p) => p.url);
const remotePaths = () => [...headTree(github.head).keys()];

/* ── fake devices ────────────────────────────────────── */
const SETTINGS = { token: "t", repo: "me/backup", branch: "main", path: "backups/hk.json", autoSync: true, autoSyncMinutes: 5 };
const makeDevice = (initial = {}) => ({ hkGithubSyncSettings: SETTINGS, ...initial });
let device = makeDevice();
globalThis.chrome = {
  storage: { local: {
    get: async (keys) => {
      if (keys === null) return structuredClone(device);
      const list = Array.isArray(keys) ? keys : [keys];
      return structuredClone(Object.fromEntries(list.filter((k) => k in device).map((k) => [k, device[k]])));
    },
    set: async (values) => { Object.assign(device, structuredClone(values)); },
    remove: async (keys) => { [].concat(keys).forEach((k) => delete device[k]); },
  } },
};
const on = async (d, fn) => { device = d; return fn(); };
const sync = (d, opts) => on(d, () => S.runCycle(opts));
const mark = (text) => [{ id: text, text, color: "#fff5b8", createdAt: 1 }];
const conflictUrls = (result) => result.conflicts.map((c) => c.url);

(async () => {
  // Pure decision table.
  assert.equal(S.classify("x", "x", "y"), "same");
  assert.equal(S.classify("x", null, undefined), "push", "new local page");
  assert.equal(S.classify(null, "x", undefined), "pull", "new remote page");
  assert.equal(S.classify("x", "y", undefined), "conflict", "first sync never overwrites a differing page");
  assert.equal(S.classify("x", "b", "b"), "push");
  assert.equal(S.classify("b", "x", "b"), "pull");
  assert.equal(S.classify("x", "y", "b"), "conflict");
  assert.equal(S.classify(null, "b", "b"), "delete-remote");
  assert.equal(S.classify("b", null, "b"), "delete-local");
  assert.equal(S.classify(null, "x", "b"), "conflict", "deleted here, edited there");
  assert.equal(S.classify("x", null, "b"), "conflict", "edited here, deleted there");
  assert.equal(S.clampMinutes("abc"), S.DEFAULT_MINUTES);
  assert.equal(S.clampMinutes(0), S.DEFAULT_MINUTES);
  assert.equal(S.clampMinutes(99999), 1440);

  // Off by default for the alarm, but "sync now" still works.
  const off = makeDevice({ hkGithubSyncSettings: { ...SETTINGS, autoSync: false } });
  assert.equal((await sync(off)).status, "disabled");
  assert.equal(github.calls.length, 0, "disabled sync makes no requests");

  // First push, without the article text stored in note.prompt.
  const laptop = makeDevice({
    [URL_A]: mark("alpha"),
    __hk_page_meta__: { [URL_A]: { title: "A", tags: ["AI"] } },
    hkGeneratedNotes: { [URL_A]: { context: "主線", prompt: "FULL ARTICLE TEXT" } },
  });
  let result = await sync(laptop);
  assert.equal(result.status, "ok");
  assert.equal(result.pushed, 1);
  assert.deepEqual(remoteUrls(), [URL_A]);
  assert.ok(![...github.blobs.values()].some((c) => c.includes("FULL ARTICLE TEXT")), "article text is not uploaded");
  assert.ok(remotePaths().includes("backups/categories/AI.json") && remotePaths().includes("backups/index.json"));

  // Nothing changed: one request, nothing written.
  github.calls.length = 0;
  result = await sync(laptop);
  assert.deepEqual(github.calls, ["GET /git/ref/heads/main"], "idle cycle only reads the ref");
  assert.equal(result.pushed, 0);

  // A second device pulls on its first sync; the note is restored.
  const desktop = makeDevice();
  result = await sync(desktop);
  assert.equal(result.pulled, 1);
  assert.deepEqual(desktop[URL_A], mark("alpha"));
  assert.equal(desktop.hkGeneratedNotes[URL_A].context, "主線");
  assert.deepEqual(desktop.__hk_page_meta__[URL_A].tags, ["AI"]);
  github.calls.length = 0;
  await sync(desktop);
  assert.deepEqual(github.calls, ["GET /git/ref/heads/main"], "a pull does not cause a follow-up push");

  // One side edits: it travels, and only changed files are uploaded.
  laptop[URL_B] = mark("bravo");
  laptop.__hk_page_meta__[URL_B] = { title: "B", tags: ["Other"] };
  await sync(laptop);
  github.uploads.length = 0;
  laptop[URL_B] = mark("bravo 2");
  result = await sync(laptop);
  assert.equal(result.pushed, 1);
  assert.ok(!github.uploads.some((c) => c.includes('"category": "AI"')), "untouched categories are not re-uploaded");
  result = await sync(desktop);
  assert.equal(result.pulled, 1);
  assert.deepEqual(desktop[URL_B], mark("bravo 2"));

  // Conflict: both edit B, and C is new. B stays put on both sides; C still syncs.
  await sync(laptop);
  laptop[URL_B] = mark("bravo laptop");
  desktop[URL_B] = mark("bravo desktop");
  desktop[URL_C] = mark("charlie");
  await sync(desktop);
  const remoteBefore = remoteFile().pages.find((p) => p.url === URL_B).entries;
  result = await sync(laptop);
  assert.equal(result.status, "conflict");
  assert.deepEqual(conflictUrls(result), [URL_B]);
  assert.equal(result.conflicts[0].kind, "both-changed");
  assert.deepEqual(laptop[URL_B], mark("bravo laptop"), "conflict leaves local untouched");
  assert.deepEqual(remoteFile().pages.find((p) => p.url === URL_B).entries, remoteBefore, "conflict leaves remote untouched");
  assert.deepEqual(laptop[URL_C], mark("charlie"), "other pages still sync");

  // Still a conflict later: unchanged, and no re-download while nothing moved.
  github.calls.length = 0;
  result = await sync(laptop);
  assert.deepEqual(conflictUrls(result), [URL_B]);
  assert.equal(github.calls.length, 1, "waiting on a conflict stays cheap");

  // Other edits keep flowing past a conflict.
  laptop[URL_A] = mark("alpha 2");
  result = await sync(laptop);
  assert.equal(result.pushed, 1);
  assert.deepEqual(conflictUrls(result), [URL_B]);
  assert.deepEqual(remoteFile().pages.find((p) => p.url === URL_B).entries, remoteBefore);

  // Choosing "keep local" pushes it; the other device then pulls it.
  result = await on(laptop, () => S.resolveConflicts({ [URL_B]: "local" }));
  assert.equal(result.status, "ok");
  assert.deepEqual(result.conflicts, []);
  assert.deepEqual(result.resolutions, {});
  assert.deepEqual(remoteFile().pages.find((p) => p.url === URL_B).entries, mark("bravo laptop"));
  // desktop's copy of B matches base, so it simply pulls.
  result = await sync(desktop);
  assert.deepEqual(desktop[URL_B], mark("bravo laptop"));
  assert.equal(result.status, "ok");

  // Choosing "use remote" overwrites local.
  laptop[URL_C] = mark("charlie laptop");
  desktop[URL_C] = mark("charlie desktop");
  await sync(desktop);
  result = await sync(laptop);
  assert.deepEqual(conflictUrls(result), [URL_C]);
  result = await on(laptop, () => S.resolveConflicts({ [URL_C]: "remote" }));
  assert.deepEqual(laptop[URL_C], mark("charlie desktop"));
  assert.equal(result.status, "ok");

  // A deletion travels both ways, and removes categories that no longer have pages.
  await sync(desktop);
  delete laptop[URL_C];
  delete laptop.__hk_page_meta__[URL_C];
  await sync(laptop);
  assert.ok(!remoteUrls().includes(URL_C));
  await sync(desktop);
  assert.ok(!(URL_C in desktop), "remote deletion reaches the other device");
  delete laptop[URL_B];
  delete laptop.__hk_page_meta__[URL_B];
  await sync(laptop);
  assert.ok(!remotePaths().includes("backups/categories/Other.json"), "empty category file is removed");
  await sync(desktop);

  // Another device pushes mid-cycle: the commit is refused rather than overwriting them.
  laptop[URL_B] = mark("late");
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    if (init?.method === "PATCH") github.head = sha1("someone else") && (() => {
      const sha = sha1("someone else commit");
      github.commits.set(sha, { tree: github.commits.get(github.head).tree, parents: [github.head] });
      return sha;
    })();
    return realFetch(url, init);
  };
  result = await sync(laptop);
  globalThis.fetch = realFetch;
  assert.equal(result.status, "error");
  assert.equal(result.errorCode, "remote-moved");
  assert.deepEqual(laptop[URL_B], mark("late"));
  result = await sync(laptop);
  assert.equal(result.status, "ok", "the next cycle recovers on its own");
  assert.ok(remoteUrls().includes(URL_B));
  await sync(desktop);

  // Network failures change nothing and are reported.
  github.fail = 401;
  result = await sync(laptop);
  assert.equal(result.status, "error");
  assert.ok(result.errorMessage.includes("boom"));
  result = await sync(laptop);
  assert.equal(result.status, "ok", "recovers after an error");

  // A remote wiped of its pages must not wipe this device.
  const many = makeDevice();
  ["1", "2", "3", "4", "5"].forEach((n) => { many[`https://example.com/p${n}`] = mark(n); });
  await sync(many);
  const wipe = await (async () => {
    const tree = new Map(headTree(github.head));
    const content = JSON.stringify({ type: "highlight-keeper-bulk", version: 2, pages: [] });
    const blob = gitSha(content);
    github.blobs.set(blob, content);
    tree.set("backups/hk.json", blob);
    const treeSha = sha1(`wiped${github.trees.size}`);
    github.trees.set(treeSha, tree);
    const commit = sha1(`wipe${github.commits.size}`);
    github.commits.set(commit, { tree: treeSha, parents: [github.head] });
    github.head = commit;
  })();
  void wipe;
  result = await sync(many);
  assert.ok(result.conflicts.length >= 5, "mass deletion becomes a conflict");
  assert.ok(result.conflicts.every((c) => c.kind === "mass-delete"));
  assert.deepEqual(many["https://example.com/p1"], mark("1"), "local data survives");

  // An unreadable remote file leaves everything alone.
  const content = "not json";
  const blob = gitSha(content);
  github.blobs.set(blob, content);
  const tree = new Map(headTree(github.head));
  tree.set("backups/hk.json", blob);
  const treeSha = sha1("broken");
  github.trees.set(treeSha, tree);
  const commit = sha1("broken commit");
  github.commits.set(commit, { tree: treeSha, parents: [github.head] });
  github.head = commit;
  const before = JSON.stringify(many["https://example.com/p1"]);
  result = await sync(many);
  assert.equal(result.errorCode, "remote-invalid");
  assert.equal(JSON.stringify(many["https://example.com/p1"]), before);
  assert.equal(github.head, commit, "nothing was pushed over it");

  // Missing settings are reported, and changing the target forgets the old base.
  const unset = makeDevice({ hkGithubSyncSettings: { autoSync: true } });
  assert.equal((await sync(unset)).errorCode, "no-token");
  const moved = makeDevice({ [URL_A]: mark("x"), hkGithubSyncSettings: { ...SETTINGS, path: "other/hk.json" } });
  moved.hkAutoSyncState = { target: "me/backup|main|backups/hk.json", base: { [URL_A]: "old" } };
  result = await sync(moved);
  assert.equal(result.pushed, 1, "a new backup location starts clean");

  // An edit made while a cycle is running is pushed by the next cycle.
  moved[URL_B] = mark("b");
  const fetchBeforeEdit = globalThis.fetch;
  let editedMidCycle = false;
  globalThis.fetch = async (...args) => {
    if (!editedMidCycle) { editedMidCycle = true; moved[URL_A] = mark("edited mid-cycle"); }
    return fetchBeforeEdit(...args);
  };
  result = await sync(moved);
  globalThis.fetch = fetchBeforeEdit;
  assert.equal(result.status, "ok");
  result = await sync(moved);
  assert.equal(result.pushed, 1, "the mid-cycle edit is not skipped");
  const movedFile = JSON.parse(github.blobs.get(headTree(github.head).get("other/hk.json")));
  assert.equal(movedFile.pages.find((p) => p.url === URL_A).entries[0].text, "edited mid-cycle");
  result = await sync(moved);
  assert.equal(result.pushed, 0, "and then it settles");

  // Pages that were only visited are not backed up, and old leftovers are cleaned.
  const P = globalThis.HkParsers;
  const stored = {
    "__hk_page_meta__": {
      "https://bank.example.com/1": { title: "Visited only" },
      [URL_A]: { title: "Has marks" },
      [URL_B]: { title: "Has tags", tags: ["x"] },
      [URL_C]: { title: "Has a personal note" },
    },
    [URL_A]: mark("x"),
    hkPersonalNotes: { [URL_C]: { text: "mine" } },
  };
  assert.deepEqual(P.collectBackupPages(stored).map((p) => p.url).sort(), [URL_A, URL_C].sort(), "visited-only pages stay out of backups");
  assert.deepEqual(Object.keys(P.cleanStoredData(stored)["__hk_page_meta__"]).sort(), [URL_A, URL_B, URL_C].sort());
  assert.deepEqual(P.cleanStoredData({ "__hk_page_meta__": { [URL_A]: { title: "t" } }, [URL_A]: mark("x") }), {}, "nothing to clean writes nothing");

  // The article text left in old notes is cleaned out; clean notes are not rewritten.
  const oldNotes = { hkGeneratedNotes: { [URL_A]: { context: "c", prompt: "article" }, [URL_B]: { context: "d" } } };
  assert.deepEqual(P.cleanStoredData(oldNotes).hkGeneratedNotes, { [URL_A]: { context: "c" }, [URL_B]: { context: "d" } });
  assert.deepEqual(P.cleanStoredData({ hkGeneratedNotes: { [URL_B]: { context: "d" } } }), {});

  // Importing restores everything a download contains, and never overwrites.
  const exported = P.collectBackupPages({
    "__hk_page_meta__": { [URL_A]: { title: "A", tags: ["t"] } },
    [URL_A]: mark("x"),
    hkGeneratedNotes: { [URL_A]: { context: "c", prompt: "article text" } },
    hkMindmaps: { [URL_A]: { outline: "# m" } },
    hkPersonalNotes: { [URL_A]: { text: "mine" } },
  });
  assert.equal(exported[0].note.prompt, undefined, "the article text is not exported");
  const restored = P.pagesToStorageUpdates(exported, {}, 5);
  assert.deepEqual(restored.imported, [URL_A]);
  assert.deepEqual(P.collectBackupPages(restored.updates), exported, "download then import round-trips");
  const kept = P.pagesToStorageUpdates(exported, { [URL_A]: mark("local") });
  assert.deepEqual(kept.skipped, [URL_A]);
  assert.deepEqual(kept.updates, {}, "an existing page is left alone");

  console.log("Auto-sync passed: push, pull, idle, conflict, resolve, deletion, races, failures and mass-delete guard.");
})().catch((error) => { console.error(error); process.exitCode = 1; });
