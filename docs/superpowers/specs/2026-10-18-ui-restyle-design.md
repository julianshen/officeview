# officeview UI restyle: Fluent look, Tailwind, Lucide, motion, themes

Date: 2026-10-18. Status: design, awaiting approval.

## 1. Intent

The viewer chrome looks dated, "even uglier than Office 365." The owner wants a modern, stylish UI that keeps Google Sheets' familiarity but is cleaner, close to Office 365's Fluent look, with some of Canva's polish. It should feel alive through motion, support light and dark themes (auto-detected), and be themable by host apps. Inline styles everywhere are also hard to maintain, so the restyle moves styling into a real, token-driven system.

## 2. Goals

1. **One visual language across every viewer surface**: spreadsheet (interactive viewer), Word/DOCX, PowerPoint/PPTX and the shared shell (RTF/ODT/static XLSX use the shared shell too).
2. **Fluent / Office 365 direction** (owner picked mockup "B"):
   - crisp, businesslike; tight 4px control corners, 6px panels/menus;
   - Fluent-style soft layered shadows (`0 1.6px 3.6px` rest, `0 6.4px 14.4px` popovers);
   - system UI font stack;
   - Excel-green cell selection (distinct from blue text selection, as spec §9 already requires);
   - accent blue `#0f6cbd` (light) / `#479ef5` (dark) as the default accent.
   - Reference palette (light / dark): surface `#fff`/`#292827`, chrome `#faf9f8`/`#201f1e`, border `#e1dfdd`/`#3b3a39`, text `#201f1e`/`#f3f2f1`, muted `#605e5c`/`#a19f9d`, icon `#323130`/`#d2d0ce`, hover `#f3f2f1`/`#3b3a39`, header `#f3f2f1`/`#201f1e`, header-active `#deecf9`/`#1c3149`, gridline `#edebe9`/`#3b3a39`, selection `#107c41`/`#4cb87a`, page backdrop `#f0efee`/`#161514`. These are starting values; final values must pass the contrast rules in §5.
3. **Light and dark themes**, following `prefers-color-scheme` by default, with a way for the host to force light or dark.
4. **Themable by hosts without a rebuild**: every colour (and radii/shadows) is a named CSS custom property; hosts override a handful of tokens to re-brand.
5. **Tailwind CSS** replaces inline style objects for chrome styling; styles compile to one shipped stylesheet.
6. **Lucide icons** replace all Unicode glyph icons (↕ ↖ ⋯ + − ↑ ↓ … ✓ ⌘, ↕N/↔ badges) and the hand-drawn funnel SVG. The owner chose Lucide over Font Awesome (Lucide is per-icon tree-shaken SVG, no shared runtime, no font download).
7. **Motion (motion.dev)** makes the UI feel alive in *all* of these moments (owner: "all"): menus/popovers and dialogs opening and closing, the mobile filter bottom sheet, sheet switching and zoom changes, and click/hover/press feedback, including the "Copied" confirmation.
8. **Word & PowerPoint get a slim top toolbar** (Office Online style, owner picked layout "1"); nothing floats over the pages. Toolbar contents, all approved by the owner:
   - file name;
   - page navigation: previous/next with "3 / 12" (PowerPoint: "Slide 3 / 12");
   - zoom out / zoom % / zoom in, plus a new **fit-to-width** button;
   - Find (input with previous/next match);
   - copy and print.
   Below the pages, a small status line: page/slide position plus **word count** (Word). Page navigation, fit-to-width, file name and word count are new features.
9. **Spreadsheet viewer** gets the same treatment (mockup B): toolbar with sheet picker, sort, filter, Find, copy, print; headers with Lucide sort/filter icons; Fluent menus, sort/filter dialogs, mobile filter sheet; status bar with selected cell, summary and zoom.

## 3. Non-goals

- Changing how document *content* renders. Pages, slides, cells and text are canvas-painted and stay exactly as they are, apart from theme colours on chrome-like canvas elements (see §6.6).
- New editing features. The viewer stays view-only.
- Animating canvas content (cells, pages, slides).
- A toast/notification system (spec §9: the status line remains the only notice surface).

## 4. Anti-goals (failure even if it "works")

- Restyling or resetting the host page (leaking global CSS, Tailwind preflight, class collisions with a host's own Tailwind).
- Any accessibility regression: 44px coarse-pointer targets, keyboard paths, ARIA names/roles, focus visibility, WCAG contrast (now in both themes).
- Layout shift: the existing CLS ≤ 0.001 guarantee (`tests/browser/viewer-layout-shift.spec.ts`) must hold; animations use transform/opacity only.
- Growing the eager JS bundle (`dist/index.js` Node-zlib gzip must stay ≤ 38,720 B; owner's "lazy-load only" rule).
- Motion that ignores the user's reduce-motion setting.

## 5. Constraints

- **Library context.** officeview is a React library embedded in other people's pages. Today it ships no CSS: `package.json` exports only `"."` and `"./worker"`; styling is inline `style={{}}` objects plus runtime `<style>` tags appended to `document.head` (print guard, sheet print media/owner, `data-officeview-sheet-font`). Hosts customise via `className`/`style` (applied to the inner region), `background`, `pageGapPx`, `showZoomControls`, `showSearch`, and `OfficeFile`'s `loading`/`error` props. These props must keep working.
- **Shipped stylesheet is acceptable** (owner: self-containment is not necessary). Hosts add `import 'officeview/style.css'`. Upgrading without the import gives an unstyled viewer: this is an **accepted breaking change** (version bump, migration note in README/CHANGELOG). My call: in development builds, log a one-time console warning when the stylesheet is not detected (e.g. a sentinel custom property is missing on the root). Tests, the demo page and the Playwright harness must load the stylesheet.
- **Eager bundle.** Eager `dist/index.js` holds `OfficeDoc` (legacy viewer + copy button), `OfficeFile`, `viewerControls` (`SearchBar`, `ZoomControls`), `gridEvents`, `pagedDocument`. Current headroom is ~20–23 B. Nothing may make it larger than 38,720 B. The budget is currently a manual gate (not enforced in CI); measure with `node -e "console.log(require('zlib').gzipSync(require('fs').readFileSync('dist/index.js')).length)"` after `bun run build` (the checked-in `dist/` is stale; always rebuild).
- **Lazy chunk.** `components/xlsx/*` loads via `lazy(() => import('./xlsx/InteractiveSheetDoc'))` (`OfficeDoc.tsx` ~852).
- **Accessibility/test contracts that must keep passing** (behaviour, not necessarily exact assertions):
  - 44px coarse targets (spec §9, "§513" in `viewerControls.tsx`); touch header = label row + 44px actions row (88px); status bar fixed height; 20px rows mirror canvas geometry.
  - `HEADER_COLORS` contrast test (`tests/xlsx-viewer-visual-review.test.tsx`: text ≥ 4.5:1, track ≥ 3:1) — extend to both themes.
  - Playwright: `interactive-viewer-visual.spec.ts` (fills host, chrome font not host serif, filter popup placement, toolbar no-overlap, header labels never truncate, contrast, print), `interactive-viewer-a11y.spec.ts`, `interactive-viewer-mobile.spec.ts` (zoom never covers grid), `interactive-viewer.spec.ts`, `interactive-viewer-resize.spec.ts`, `viewer-layout-shift.spec.ts`, `interactive-viewer-dpr.spec.ts`.
  - aria-labels used by tests and assistive tech ("Zoom in", "Previous match", "Select worksheet", "Cycle sort column X", "Filter column X — …", etc.) stay; icons are decorative (`aria-hidden`).
- **Process rules** (project): TDD; every change reviewed by a different model family than its author; all findings fixed or disputed with accepted evidence before merge; never commit/merge/push to `main` (owner merges the PR); work on `feature/ui-restyle` and per-stream branches in `.worktrees/`.

## 6. Decided approach

### 6.1 Styling system
- Tailwind CSS (v4) compiled at build time into **`dist/style.css`**, exported as `"./style.css"` in `package.json` (add to `exports` and to `sideEffects` so bundlers keep it).
- **Class prefix** so library classes never collide with a host's own Tailwind (Tailwind v4 `prefix(ov)`, giving classes like `ov:flex`; exact prefix syntax is the builder's call within this rule).
- **Preflight/base reset disabled.** Any base styling is scoped under the library root (`[data-officeview-root]`), never global.
- Design tokens as CSS custom properties, namespaced (e.g. `--ov-color-surface`, `--ov-color-accent`, `--ov-radius-control`, `--ov-shadow-popover`), defined on the library root, with a light set, a dark set, and `prefers-color-scheme` switching. Tailwind's theme maps to these variables so components use semantic classes (surface, chrome, border, text, muted, icon, hover, accent, on-accent, header, header-active, gridline, selection, selection-fill, backdrop).
- Inline style objects for chrome are replaced by classes. Inline styles remain only for genuinely dynamic geometry (computed positions/sizes, transforms, canvas sizing).
- Existing runtime-injected `<style>` rules that are static (e.g. `data-officeview-sheet-font` control font, SearchBar overrides) move into the stylesheet; per-instance print rules that depend on runtime IDs may stay injected.

### 6.2 Theme API
- Default: follow the system (`prefers-color-scheme`).
- Host override: a `theme` prop on `OfficeDoc`/`OfficeFile` (`'light' | 'dark' | 'system'`, default `'system'`) that sets a data attribute on the root (e.g. `data-ov-theme`), plus documented CSS variables for re-branding. Document the token list in the README.

### 6.3 Icons
- `lucide-react`, imported per icon (named imports, tree-shaken), `aria-hidden` with the existing accessible labels on the buttons.
- Replace every glyph listed in the survey: search prev/next, zoom +/−, corner ↖, sort ↕/↑/↓/…, more ⋯ (and touch disclosure, lane "B⋯"), filter funnel, copy ✓/⌘ text, touch resize badges ↕N/↔.
- Eager-shell icons must fit the budget (see §6.7).

### 6.4 Motion
- Use motion.dev (`motion` package), loaded **only in lazy code**; never in the eager bundle. Prefer the lightweight APIs (`LazyMotion` + `domAnimation` features, or `motion/mini` animate) to keep the lazy chunk small.
- Simple feedback (hover, press, focus) uses Tailwind CSS transitions (zero JS).
- Richer moments use motion: menus/popovers/dialogs enter/exit, mobile filter sheet slide-up, sheet-switch crossfade, zoom change easing, "Copied" confirmation.
- transform/opacity only; honour `prefers-reduced-motion` (no/near-instant motion when set). CLS gate must stay ≤ 0.001.

### 6.5 Word & PowerPoint toolbar and status line
- Replace today's floating `SearchBar`/`ZoomControls`/copy pill in `LegacyOfficeDoc` with a top toolbar and a bottom status line, per §2 item 8.
- Respect existing props: `showSearch` hides Find, `showZoomControls` hides zoom; add `showToolbar` (default true) to hide the whole toolbar for hosts with their own UI. Status line visibility follows the toolbar unless the builder finds a reason to separate it.
- File name: from the opened file (`OfficeFile` knows it) with a prop to override (e.g. `title`).
- Page navigation scrolls to the previous/next page and shows the page in view; PowerPoint wording "Slide".
- Fit-to-width sets zoom so the page width fits the viewport.
- Word count from the document's text model (Word); PowerPoint shows slide count instead.
- Copy-blocked and copied states keep their meaning (today "Copied ✓" / "Copy blocked — press ⌘C"), now with icons and motion.

- **Print (owner decision):** a proper document print, not `window.print()` of the host page. Build page-image print output for the paged viewer (DOCX, PPTX, RTF, ODT, static XLSX) so Print outputs only the document's pages at page size, analogous to the spreadsheet viewer's existing print path (print container with per-page images, print media rules, the existing print guard). This is its own work stream.
- **Disallowed actions (owner decision):** when `allowPrint={false}` or `allowCopy={false}`, the Print / Copy buttons are **hidden**, not shown disabled. Existing blocking behaviour (Cmd/Ctrl-P intercept, copy blocking) is unchanged.
- **Formats (builder-check call, mine):** every format rendered by the paged viewer gets the toolbar and status line. RTF and ODT behave like Word (page navigation + word count, from their text models). Static (non-interactive) XLSX gets page navigation but no word count.
- **File name fallback (my call):** use `File.name`, the last URL path segment, or a `Content-Disposition` filename when available; the `title` prop overrides; with no name, the slot is hidden. The name travels from the loader (`OfficeFile`) to the doc.
- **Narrow / touch toolbar (my call):** on coarse pointers or narrow panes, the toolbar keeps 44px targets by collapsing Find to an icon button that expands the input, hiding the file name, and moving lower-priority actions (fit, print, copy) into a "More" menu. No control may overlap the pages (extend the mobile no-overlap Playwright checks to the paged viewer).
- **Lazy fallback (my call):** if the toolbar or status line must be lazy-loaded for bundle reasons, their height is reserved by a fixed-size placeholder at first paint so the CLS gate (≤ 0.001) holds. The "Copied" confirmation in the eager paged viewer uses a CSS-only animation (motion.dev stays lazy-only).

### 6.6 Canvas theming
- Chrome-like canvas elements read the theme tokens so dark mode is whole: sheet gridlines (`#d0d0d0` today), selection overlay (`src/core/overlay.ts`, green fill + 2px border), find highlights, freeze divider (`sheetTiles.ts`), tile "Loading…" placeholders, and the document page backdrop (`#888` today). Document content colours (cell fills, text, page paper) do not change.

### 6.7 Bundle strategy
- Moving chrome styles from inline JS objects to the stylesheet shrinks the eager JS; that freed room funds eager-shell icons and the new toolbar. Verify with a real build after every stream; the 38,720 B ceiling is hard. If the new Word/PowerPoint toolbar does not fit eagerly, lazy-load it (the pages can render first).

## 7. Left to the builder

- Exact Tailwind version/config details, prefix syntax, file layout of the stylesheet and token files.
- Exact final colour values within the Fluent direction, as long as contrast rules pass in both themes.
- Motion durations/easings (subtle, Fluent-like: ~100–250ms).
- Exact icon choice per action from the Lucide set.
- How to stream the work across agents and in what order, test strategy for visual changes (before/after screenshots can reuse `.worktrees/visual-shots/scripts/capture.mjs`, `compose.py`, `make_index.py`).
- Updating tests that read inline `style.*` values (about 15 files, listed in the survey) to assert behaviour or computed style instead, without weakening the guarantees they protect.
