# Reading Sidebar Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Chrome side panel with concise article context, term explanations, key highlights, and per-page personal notes.

**Architecture:** Keep AI and page text operations in the content script; add a side panel extension page for navigation, display, and manual notes. Use existing page highlight storage and generated-note storage, extending them with optional kind and context fields; keep personal notes in a separate per-page map.

**Tech Stack:** Chrome Manifest V3, vanilla JavaScript, HTML, CSS, `chrome.sidePanel`, `chrome.storage.local`.

**Spec:** `docs/superpowers/specs/2026-10-22-reading-sidebar-design.md`

## Global Constraints

- Minimum supported Chrome version: 116.
- AI output: 3–5 highlighted snippets by default, hard limit 7 total including terms.
- Context: 3–5 sentences and no more than 350 Chinese characters.
- Personal notes stay local and are not sent to an AI provider by default.
- Preserve existing popup, in-page panel, custom prompts, highlights, imports, and exports.
- Do not add or run tests unless the user asks.

## Review Focus

- Restricted pages and tabs without a content script must show an unavailable state without breaking the side panel.
- Switching active tabs must load the matching canonical page key and never show the previous page's notes.
- Legacy generated notes and highlights without a `kind` field must remain visible.
- Partial or malformed AI responses must not discard valid context, existing highlights, or personal notes.
- AI regeneration, import, and export must preserve custom prompts and handwritten notes.

---

### Task 1: Add a compact reading-guide format

**Files:**
- Modify: `parsers.js`
- Modify: `contentScript.js`
- Modify: `i18n.js`

**Interfaces:**
- Produces `HkParsers.parseReadingGuideResponse(rawText, deps) -> { context: string, terms: Array<{ text: string, explanation: string }>, highlights: Array<{ text: string, kind: "passage", reason: string, color: string }> }`.
- The term parser returns exact source text and a concise explanation; integration normalizes terms into `{ text, kind: "term", reason, color }` highlight tasks.
- Accepts existing bilingual section markers and legacy highlight formats without changing their existing parser behavior.

- [ ] Add `context` and `terms` section aliases and implement `parseReadingGuideResponse` using existing section-splitting and highlight-block helpers.
- [ ] Add the new default reading-guide and auto-highlight prompts: merge main line with narrative flow; cap context at 3–5 sentences/350 Chinese characters; target 3–5 marks and cap total marks at 7; require contiguous source text and no overlap between term and passage marks.
- [ ] Migrate only stored prompt values exactly equal to the previous built-in defaults. Disable the legacy default deep-research setting only when both stored prompts still equal their previous defaults; keep customized values intact.
- [ ] Normalize AI term and passage results to the optional `kind` field, preserving the current preview-before-apply setting and ignoring duplicate existing text.
- [ ] Add Traditional Chinese and English strings for guide generation, each tab, empty states, parsing errors, manual paste, and save status.

### Task 2: Persist guide data and personal notes through existing formats

**Files:**
- Modify: `contentScript.js`
- Modify: `popup.js`
- Modify: `parsers.js`
- Modify: `manager.js`
- Modify: `share-link.js` only if its page payload normalization drops unknown fields

**Interfaces:**
- Generated note payload keeps existing fields and adds `formatVersion` and `context`; legacy `note` remains readable.
- Personal notes use `hkPersonalNotes[pageKey] = { text: string, updatedAt: number }`.
- Exported page entries add optional `personalNote: string`; old imports without this property remain valid.

- [ ] Save and read `context` without overwriting legacy `note`, existing mind maps, or personal notes.
- [ ] Add optional `personalNote` to current-page export, Markdown copy/parse, JSON download/import, manager bulk import/export, and shared-page apply.
- [ ] Preserve `kind` on stored highlights through import, export, and share round trips; route legacy entries without `kind` to the regular highlight list.
- [ ] Keep AI regeneration isolated from `hkPersonalNotes` and keep personal-note text out of every AI prompt builder.

### Task 3: Add side-panel opening and page-context bridge

**Files:**
- Modify: `manifest.json`
- Modify: `popup.html`
- Modify: `popup.js`
- Modify: `contentScript.js`

**Interfaces:**
- Add popup control `openReadingSidebarBtn`; on click call `chrome.sidePanel.open({ tabId })` for the active supported tab.
- Add message `BUILD_READING_GUIDE_PROMPT` returning `{ success, prompt, pageKey }`.
- Add message `GENERATE_READING_GUIDE` for configured OpenAI/Gemini direct calls, returning `{ success, context, appliedCount, previewing }`.
- Add message `APPLY_READING_GUIDE_RESPONSE` accepting ChatGPT text, saving valid context and applying parsed marks through the existing preview flow.
- Add `FOCUS_SAVED_HIGHLIGHT` for `{ text, kind }` to scroll to the matching saved annotation and briefly emphasize it.

- [ ] Declare `minimum_chrome_version: "116"`, the `sidePanel` permission, and `side_panel.default_path: "sidepanel.html"`; retain the current toolbar popup and existing `toggle-panel` command.
- [ ] Add the popup button and localized label; open the side panel only from its explicit user click.
- [ ] Reuse `GET_PAGE_HIGHLIGHTS` data (`title`, canonical `url`, `pageText`, and highlights) for the side-panel context.
- [ ] Add the three reading-guide message paths and route them to the shared parser, existing provider calls, generated-note storage, and existing highlight preview/apply functions.
- [ ] Add `FOCUS_SAVED_HIGHLIGHT` to find matching saved text and use existing range/highlight navigation behavior; return a clear failure when no match exists.

### Task 4: Build the four-tab side-panel page

**Files:**
- Create: `sidepanel.html`
- Create: `sidepanel.js`
- Create: `sidepanel.css`
- Modify: `contentScript.js`
- Modify: `i18n.js`

**Interfaces:**
- Tabs are `context`, `terms`, `highlights`, and `personal-notes`; only the selected panel is visible.
- Side panel tracks `chrome.tabs.onActivated`, relevant `chrome.tabs.onUpdated` events, `chrome.storage.onChanged`, and the active page's canonical URL.
- Manual note input saves to `hkPersonalNotes[pageKey]` after a short debounce and displays saved/unsaved state.

- [ ] Build the compact header, four accessible tabs, separate count badges, refresh/generate action, per-tab empty states, and scrollable content areas to match `sidebar-wireframe.html`.
- [ ] On active-tab change, request `GET_PAGE_HIGHLIGHTS`, load generated context and personal note for that page, and clear stale view data before displaying the new page.
- [ ] Render terms and passage marks in their own tabs; clicking an item sends `FOCUS_SAVED_HIGHLIGHT` and retains the note if source text no longer matches.
- [ ] Add the personal-notes textarea, debounced local save, storage-change synchronization, and a visible save status. Do not trigger AI calls from edits.
- [ ] For ChatGPT provider, display copy/open/paste/apply controls; for missing AI configuration, link to existing settings. Show distinct unsupported-page, extraction, generation, and parse-error states.

### Task 5: Update user-facing documentation

**Files:**
- Modify: `README.md`
- Modify: `FEATURES.md`
- Modify: `PRIVACY.md`

- [ ] Document the four side-panel tabs, the compact output limits, personal-note autosave, and how to open the panel.
- [ ] State that personal notes stay in local storage and are excluded from AI requests by default.
- [ ] Document the Chrome 116 minimum required by `chrome.sidePanel.open()`.

## Plan Self-Review

- Spec coverage maps to Tasks 1–5: prompt and migration, AI output, storage compatibility, opening the panel, tab UI, errors, privacy, and docs.
- The side-panel page depends on the parser and message contract established in Tasks 1 and 3; the task order reflects those dependencies.
- `personalNote` is distinct from generated `note`; `kind` is optional so old highlight entries remain readable.
- No automated test files or test commands are included under the current instruction not to add or run tests unless requested.
