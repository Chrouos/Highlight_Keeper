# Reading desk
Approved through the draggable reading-desk preview and the user's request to implement it.

While highlighting an original web page, extract passages into draggable floating cards and write reactions directly beneath them. Add independent sticky notes. Auto-save both text and placement; restore on reopening. Keep the original page intact. Use neutral gray surfaces and restrained blue controls, with highlight colors confined to quotations. No green theme.

The extension icon opens the in-page desk on supported HTTP(S) pages. Preserve the existing reading/AI sidebar as a secondary action and the manager for later organization. Existing highlights and their notes remain canonical. A separate per-page desk record stores pinned highlight IDs, card geometry, collapsed state, and free notes. Preserve desk records in JSON backup and auto-sync. Returning to a source focuses the existing highlight, and reports a missing source without losing the excerpt.

Approaches: the old native sidebar is reliable but narrows the article; a detached window separates source and notes; the approved page overlay keeps excerpts beside their source and permits arranging them. Implement the overlay with a closed Shadow DOM to isolate site styles. The overlay receives no page-wide backdrop and cannot swallow clicks outside its controls. Close hides the desk without deleting data. Remove from desk unpins a quote and does not delete the original highlight.

Implementation calls: independent free notes, rather than replacing the existing personal-note document; per-page persistence; no new AI behavior. Unsupported pages open the existing sidebar instead. No external services or new dependencies.
