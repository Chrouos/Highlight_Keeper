/* 純函式單元測試（無依賴）。執行：node test/run.js
   載入 shared.js / parsers.js（IIFE，會掛到 globalThis）後直接驗證。 */
require("../shared.js");
require("../parsers.js");

const U = globalThis.HkUrlKey;
const P = globalThis.HkParsers;

let pass = 0;
let fail = 0;
const fails = [];
const eq = (actual, expected, msg) => {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    pass += 1;
  } else {
    fail += 1;
    fails.push(`✗ ${msg}\n    expected: ${e}\n    actual:   ${a}`);
  }
};
const ok = (cond, msg) => eq(Boolean(cond), true, msg);

// ── normalizePageKey ────────────────────────────────────
const npk = U.normalizePageKey;
eq(
  npk("https://x.com/p/a?utm_source=mail&id=7#sec"),
  "https://x.com/p/a?id=7",
  "normalizePageKey strips utm_/hash, keeps id"
);
eq(
  npk("https://x.com/a?fbclid=z&gclid=q"),
  "https://x.com/a",
  "normalizePageKey strips fbclid/gclid"
);
eq(
  npk("https://x.com/p/a?utm_source=a&r=1"),
  npk("https://x.com/p/a?utm_campaign=b&r=1"),
  "same article, different tracking → same key (r kept both)"
);
eq(npk("not a url"), "not a url", "invalid url returned as-is");

// ── splitAiSections（中英雙語）────────────────────────────
eq(
  P.splitAiSections("===重點===\nA\n\n===摘要===\nB"),
  { highlights: "A", note: "B" },
  "splitAiSections zh markers"
);
eq(
  P.splitAiSections("===Highlights===\nA\n\n===Summary===\nB"),
  { highlights: "A", note: "B" },
  "splitAiSections en markers"
);
eq(
  P.splitAiSections("## 心智圖\n- a"),
  { mindmap: "- a" },
  "splitAiSections markdown heading + mindmap alias"
);
eq(P.splitAiSections("just text no markers"), null, "splitAiSections no markers → null");

// ── parseHighlightBlocks（中英雙語 + 顏色）──────────────────
const deps = {
  categories: [
    { name: "動機／背景", color: "#c792ea" },
    { name: "Motivation", color: "#81c784" },
  ],
  toHexColor: (c) => (typeof c === "string" ? c.toLowerCase() : "#ffeb3b"),
  defaultColor: "#ffeb3b",
};
const zhBlocks = P.parseHighlightBlocks(
  "原文：人工智慧改變產業\n#動機／背景\n重點：AI 是核心動力",
  deps
);
eq(zhBlocks.length, 1, "parseHighlightBlocks zh → 1 item");
eq(zhBlocks[0].text, "人工智慧改變產業", "zh text parsed");
eq(zhBlocks[0].color, "#c792ea", "zh category color by name");
eq(zhBlocks[0].reason, "AI 是核心動力", "zh reason parsed");

const enBlocks = P.parseHighlightBlocks(
  "Source: AI changes industry\n#Motivation\nPoint: AI is the core driver",
  deps
);
eq(enBlocks.length, 1, "parseHighlightBlocks en → 1 item");
eq(enBlocks[0].text, "AI changes industry", "en text parsed (Source:)");
eq(enBlocks[0].color, "#81c784", "en category color by name (case-insensitive)");
eq(enBlocks[0].reason, "AI is the core driver", "en reason parsed (Point:)");

const hexBlock = P.parseHighlightBlocks("原文：x\n#AABBCC\n重點：y", deps);
eq(hexBlock[0].color, "#aabbcc", "raw hex tag resolved to color");

eq(
  P.parseHighlightBlocks("#重點\n重點：no source line", deps).length,
  0,
  "block without source text is skipped"
);

// ── parseMindmapOutline ─────────────────────────────────
const tree = P.parseMindmapOutline("# 主題\n- A\n  - A1\n- B", "fallback");
eq(tree.title, "主題", "mindmap title from heading");
eq(tree.children.length, 2, "mindmap top-level branches");
eq(tree.children[0].children[0].label, "A1", "mindmap nested child");
eq(P.parseMindmapOutline("", "fb"), null, "empty outline → null");
eq(
  P.parseMindmapOutline("- only\n- items", "FB").title,
  "FB",
  "no heading → fallback title"
);

// ── looksLikeMindmapOutline ─────────────────────────────
ok(P.looksLikeMindmapOutline("- a\n- b\n- c"), "list lines → looks like outline");
ok(
  !P.looksLikeMindmapOutline("原文：x\n#t\n重點：y"),
  "highlight blocks (原文：) → not outline"
);
ok(
  !P.looksLikeMindmapOutline("Source: x\n- a\n- b"),
  "english source line → not outline"
);

// ── normalizeBulkPages ──────────────────────────────────
const pages = P.normalizeBulkPages([
  {
    url: "https://x.com/a",
    title: "T",
    tags: [" t1 ", ""],
    entries: [{ id: 1 }],
    note: { note: "n" },
    mindmap: { outline: "# m" },
  },
  { url: "not-a-url", entries: [{ id: 2 }] },
  { url: "https://x.com/b", entries: [] },
  null,
]);
eq(pages.length, 1, "normalizeBulkPages drops invalid url / empty entries / null");
eq(pages[0].tags, ["t1"], "normalizeBulkPages trims/filters tags");
eq(pages[0].note, { note: "n" }, "normalizeBulkPages carries note");
eq(pages[0].mindmap, { outline: "# m" }, "normalizeBulkPages carries mindmap");

// ── 結果 ────────────────────────────────────────────────
console.log(`\n${pass} passed, ${fail} failed`);
if (fail) {
  console.log("\n" + fails.join("\n"));
  process.exit(1);
}
