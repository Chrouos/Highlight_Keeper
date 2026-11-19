# Reading margin implementation record

Approved direction: show passage thoughts automatically near the visible source, fade away when the source leaves the reading band, and persist only after explicit pinning. Free notes remain draggable desk notes.

Architecture: reading-margin-layout.js is a pure placement algorithm; reading-annotations.js renders a read-only closed-shadow overlay and observes scrolling/storage/navigation. HkDeskBridge supplies canonical thoughts and source geometry. Pin delegates to the existing serialized desk store. No data migration or new storage key.

- [x] Add placement regressions for reading-band proximity, collision avoidance, pinned exclusions and narrow fallbacks; observed missing-module failure before implementation.
- [x] Implement transient notes, source connectors, fade, explicit pinning, and automatic refresh without scroll-triggered writes.
- [x] Fix quote-card/quote CSS class collision and use compatible 220px widths before/after pinning.
- [x] Include visible desk trays and AI panels in obstacles; continue source reconciliation while notes have focus.
- [x] Follow the source page's background theme and keep paragraph text clear.
- [x] Run existing Node suites, dedicated headless Chrome margin and desk fixtures, syntax checks and diff checks.

Independent review found width changes on pin, hover/focus blocking source reconciliation, and missing panel obstacles. All three addressed and covered by browser regressions.

Verification: pure placement tests; 17 margin browser checks; 47 desk browser checks. Fixtures mock extension storage/runtime but use real DOM/CSS. Static visual inspection used the actual annotation renderer and confirmed source linkage, readable spacing, and light-paper appearance. Live installed-extension behavior on the user's own article still requires extension reload and page refresh.
