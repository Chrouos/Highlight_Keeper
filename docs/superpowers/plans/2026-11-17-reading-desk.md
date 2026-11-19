# Reading Desk Implementation Plan

> **For agentic workers:** Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Deliver the approved in-page desk with draggable quotations and auto-saved notes.
**Architecture:** A pure desk model validates and mutates records; a Shadow DOM controller handles cards and delegates highlight changes to the existing content script. Background serializes desk storage writes. Existing parsers and sync carry the desk record alongside page data.
**Tech Stack:** Vanilla JavaScript, Chrome MV3, Node assert tests.
**Spec:** docs/superpowers/specs/2026-11-17-reading-desk-design.md

## Global Constraints
- Preserve existing highlight IDs, notes and storage keys.
- HTTP(S) only, no new dependencies. Neutral/blue UI, isolated from site CSS.
- Never insert user content as HTML. Always save to the captured source page.

## Review Focus
- Navigation while a save or load is pending must not mix articles.
- Storage failure must report failure and preserve unsaved text.
- Hostile imported geometry and missing sources must remain usable.
- Backup round trips must retain free notes and layout.
- Typing or dragging must not trigger the host site's shortcuts.

### Task 1: Persistent desk model and backup
- [x] Write failing tests for normalization, immutable patching, viewport clamping, source IDs, backup round trip, and preservation of existing notes.
- [x] Implement reading-desk-model.js and desk fields in parsers/sync/manager.
- [x] Run node test/reading-desk.js and existing test scripts.

### Task 2: In-page controller and entry points
- [x] Integrate a bridge to canonical highlight operations.
- [x] Implement shadow-root dock, searchable source tray, draggable quote/note cards, collapse, close and save feedback.
- [x] Connect selection toolbar extraction and extension action, retain native sidebar shortcut.
- [x] Test bridge source binding, background writes and manifest integration.

### Task 3: Verify and review
- [x] Run all test/*.js, syntax checks and git diff --check.
- [x] Review races, storage failure, DOM isolation, selection preservation, and backup paths.
- [x] Report browser validation limits and how to use the desk.

## Execution notes
Work in place to preserve the user's current uncommitted implementation and installed extension path. No commits, merge, publishing, or external messaging required.


## Verification and review record
- Implemented in place; no commits or branch integration requested.
- Independent reviewer found stale retries, stale async reload snapshots, shared-map sync overwrite, and an expired-user-gesture fallback.
- Fixed stale retries with operation tickets and removed superseded callbacks. Fixed reloads with mutation versions and load guards. Added one background storage queue for desk patches/batch import/delete and GitHub sync cycles. Replaced the asynchronous side-panel fallback with a library tab.
- Node test suite passes. Headless Chrome fixture passes 18 controller integration checks, including canonical quote notes, free-note auto-save, movement, source focus, navigation flush, reopen, save failure/retry, stale reload, and superseded retry.
- Browser fixture uses mocked Chrome storage/runtime; a live installed-extension click-through remains to be performed by the user after extension reload.
- Ruling: on supported pages without current content scripts, open the library rather than silently attempting a side panel after the gesture expires. Cost: the user must refresh the article to use the desk after an extension reload.
