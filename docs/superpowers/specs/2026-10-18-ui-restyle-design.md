# officeview UI restyle: Fluent look, Tailwind, Lucide, motion, themes

Date: 2026-10-18. Status: owner-approved 2026-10-18; revised the same day after the codex / agy / opencode review. The revision adds owner decisions A and B and builder calls R1–R12 (§8); sections that changed point to them.

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
   - *Revision (owner decision B, §6.6):* `gridline` themes only the DOM chrome dividers (header grid lines). The canvas gridlines inside the sheet are part of the paper and stay light. Selection is drawn on white paper, so its dark-theme value must stay legible on `#fff`. `#4cb87a` is only about 2.4:1 against white, so the starting dark value becomes `#107c41`, the same as light.
3. **Light and dark themes**, following `prefers-color-scheme` by default, with a way for the host to force light or dark.
4. **Themable by hosts without a rebuild**: every colour (and radii/shadows) is a named CSS custom property; hosts override a handful of tokens to re-brand.
5. **Tailwind CSS** replaces inline style objects for chrome styling; styles compile to one shipped stylesheet.
6. **Lucide icons** replace all Unicode glyph icons (↕ ↖ ⋯ + − ↑ ↓ … ✓ ⌘, ↕N/↔ badges) and the hand-drawn funnel SVG. The owner chose Lucide over Font Awesome (Lucide is per-icon tree-shaken SVG, no shared runtime, no font download). *Precision (review):* `lucide-react` icons do share a small React icon component, so "no shared runtime" means no font and no icon-set runtime. That shared component is still bytes, which is why Lucide stays out of eager code (§6.3, §6.7).
7. **Motion (motion.dev)** makes the UI feel alive in *all* of these moments (owner: "all"): menus/popovers and dialogs opening and closing, the mobile filter bottom sheet, sheet switching and zoom changes, and click/hover/press feedback, including the "Copied" confirmation. *Owner decision A (revision):* motion for sheet switching and zoom animates the chrome only: the sheet picker, toolbar, status line and zoom readout. The grid and page canvases swap instantly.
8. **Word & PowerPoint get a slim top toolbar** (Office Online style, owner picked layout "1"); nothing floats over the pages. Toolbar contents, all approved by the owner:
   - file name;
   - page navigation: previous/next with "3 / 12" (PowerPoint: "Slide 3 / 12");
   - zoom out / zoom % / zoom in, plus a new **fit-to-width** button;
   - Find (input with previous/next match);
   - copy and print.
   Below the pages, a small status line: page/slide position plus **word count** (Word). Page navigation, fit-to-width, file name and word count are new features.
9. **Spreadsheet viewer** gets the same treatment (mockup B): toolbar with sheet picker, sort, filter, Find, copy, print; headers with Lucide sort/filter icons; Fluent menus, sort/filter dialogs, mobile filter sheet; status bar with selected cell, summary and zoom.

## 3. Non-goals

- Changing how document *content* renders. Pages, slides, cells and text are canvas-painted and stay exactly as they are, apart from theme colours on chrome-like canvas elements (see §6.6). Under owner decision B, document content stays paper in both themes: white paper, authored fills and text, and light gridlines inside the sheet.
- New editing features. The viewer stays view-only.
- Animating canvas content (cells, pages, slides). This includes sheet-switch crossfades of the grid and eased canvas zoom (owner decision A). No canvas animation is authorised anywhere in this spec.
- A toast/notification system (spec §9: the status line remains the only notice surface).

## 4. Anti-goals (failure even if it "works")

- Restyling or resetting the host page (leaking global CSS, Tailwind preflight, class collisions with a host's own Tailwind).
- Any accessibility regression: 44px coarse-pointer targets, keyboard paths, ARIA names/roles, focus visibility, WCAG contrast (now in both themes).
- Layout shift: the existing CLS ≤ 0.001 guarantee (`tests/browser/viewer-layout-shift.spec.ts`) must hold; animations use transform/opacity only.
- Growing the eager JS bundle (`dist/index.js` Node-zlib gzip must stay ≤ 38,720 B; owner's "lazy-load only" rule).
- Motion that ignores the user's reduce-motion setting.

## 5. Constraints

- **Library context.** officeview is a React library embedded in other people's pages. Today it ships no CSS: `package.json` exports only `"."` and `"./worker"`; styling is inline `style={{}}` objects plus runtime `<style>` tags appended to `document.head` (print guard, sheet print media/owner, `data-officeview-sheet-font`). Hosts customise via `className`/`style` (applied to the inner region), `background`, `pageGapPx`, `showZoomControls`, `showSearch`, and `OfficeFile`'s `loading`/`error` props. These props must keep working.
- **Shipped stylesheet is acceptable** (owner: self-containment is not necessary). Hosts add `import 'officeview/style.css'`. Upgrading without the import gives an unstyled viewer: this is an **accepted breaking change** (version bump, migration note in README/CHANGELOG). My call: in development builds, log a one-time console warning when the stylesheet is not detected (e.g. a sentinel custom property is missing on the root). Tests, the demo page and the Playwright harness must load the stylesheet. *Revision (R10):* the warning is a runtime check and is not gated on `import.meta.env`. Vite lib mode would inline `PROD` as `true` when the library is built, which would silence the warning in every consumer's development build. The viewer warns once per page whenever the sentinel is missing.
- **Eager bundle.** Eager `dist/index.js` holds `OfficeDoc` (legacy viewer + copy button), `OfficeFile`, `viewerControls` (`SearchBar`, `ZoomControls`), `gridEvents`, `pagedDocument`. Current headroom is ~20–23 B (agy measured 38,697 B on the base, which leaves 23 B; the plan re-measures before any change). Nothing may make it larger than 38,720 B. The budget is currently a manual gate (not enforced in CI); measure with `node -e "console.log(require('zlib').gzipSync(require('fs').readFileSync('dist/index.js')).length)"` after `bun run build` (the checked-in `dist/` is stale; always rebuild).
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
- **Preflight/base reset disabled.** Any base styling is scoped under the library root (`[data-officeview-root]`), never global. *Revision (R2, R3):* the "library root" for styling is now the **theme root** (`[data-ov-theme-root]`, §6.8). `data-officeview-root` keeps its existing meaning: the inner region that receives the host's `className`/`style`.
- *Revision (R3), isolation mechanism:* the shipped CSS has no `:root`, `:host`, `@property`, `*` or other unscoped rule. Every emitted selector contains `[data-ov-theme-root]`. The details are in §6.8.
- Design tokens as CSS custom properties, namespaced (e.g. `--ov-color-surface`, `--ov-color-accent`, `--ov-radius-control`, `--ov-shadow-popover`), defined on the library root, with a light set, a dark set, and `prefers-color-scheme` switching. Tailwind's theme maps to these variables so components use semantic classes (surface, chrome, border, text, muted, icon, hover, accent, on-accent, header, header-active, gridline, selection, selection-fill, backdrop).
- Inline style objects for chrome are replaced by classes. Inline styles remain only for genuinely dynamic geometry (computed positions/sizes, transforms, canvas sizing).
- Existing runtime-injected `<style>` rules that are static (e.g. `data-officeview-sheet-font` control font, SearchBar overrides) move into the stylesheet; per-instance print rules that depend on runtime IDs may stay injected.

### 6.2 Theme API
- Default: follow the system (`prefers-color-scheme`).
- Host override: a `theme` prop on `OfficeDoc`/`OfficeFile` (`'light' | 'dark' | 'system'`, default `'system'`) that sets a data attribute on the root (e.g. `data-ov-theme`), plus documented CSS variables for re-branding. Document the token list in the README.
- *Revision (R2):* `data-ov-theme` sits on the theme root. The documented re-branding surface is any host rule that targets `[data-ov-theme-root]`, for example `.brand [data-ov-theme-root] { --ov-color-accent: #c50f1f }`. Library token rules use `:where()`, which has zero specificity, so any host selector wins.
- *Revision (review), live overrides:* DOM chrome follows token changes immediately. Canvas overlay colours are re-read on mount, on a `theme` change, on a system scheme change, and when the theme root's `style`, `class` or `data-ov-theme` attribute changes. A host that changes tokens only through an ancestor stylesheet must re-render with a `theme` change or remount for canvas overlays to follow. Document this in the README.

### 6.3 Icons
- `lucide-react`, imported per icon (named imports, tree-shaken), `aria-hidden` with the existing accessible labels on the buttons.
- Replace every glyph listed in the survey: search prev/next, zoom +/−, corner ↖, sort ↕/↑/↓/…, more ⋯ (and touch disclosure, lane "B⋯"), filter funnel, copy ✓/⌘ text, touch resize badges ↕N/↔.
- Eager-shell icons must fit the budget (see §6.7). *Revision (R5):* there are no eager-shell icons. All Lucide imports live in lazy chunks. A static import-graph test enforces this.

### 6.4 Motion
- Use motion.dev (`motion` package), loaded **only in lazy code**; never in the eager bundle. Prefer the lightweight APIs (`LazyMotion` + `domAnimation` features, or `motion/mini` animate) to keep the lazy chunk small.
- Simple feedback (hover, press, focus) uses Tailwind CSS transitions (zero JS).
- Richer moments use motion: menus/popovers/dialogs enter/exit, mobile filter sheet slide-up, ~~sheet-switch crossfade, zoom change easing~~, "Copied" confirmation. *Owner decision A:* on a sheet switch, the sheet picker's selected label and the status line animate. On a zoom change, the zoom readout and status line animate. The grid canvas and content transform change in one step. No outgoing pane is retained, so no second canvas is ever mounted or charged.
- transform/opacity only; honour `prefers-reduced-motion` (no/near-instant motion when set). CLS gate must stay ≤ 0.001.

### 6.5 Word & PowerPoint toolbar and status line
- Replace today's floating `SearchBar`/`ZoomControls`/copy pill in `LegacyOfficeDoc` with a top toolbar and a bottom status line, per §2 item 8.
- Respect existing props: `showSearch` hides Find, `showZoomControls` hides zoom; add `showToolbar` (default true) to hide the whole toolbar for hosts with their own UI. Status line visibility follows the toolbar unless the builder finds a reason to separate it.
- File name: from the opened file (`OfficeFile` knows it) with a prop to override (e.g. `title`).
- Page navigation scrolls to the previous/next page and shows the page in view; PowerPoint wording "Slide".
- Fit-to-width sets zoom so the page width fits the viewport. *Revision (R1):* fit-to-width is a **layout mode**, not a transform zoom. See §6.9.
- Word count from the document's text model (Word); PowerPoint shows slide count instead. *Revision (R8):* the counting policy and time slicing are in §6.11.
- Copy-blocked and copied states keep their meaning (today "Copied ✓" / "Copy blocked — press ⌘C"), now with icons and motion.

- **Print (owner decision):** a proper document print, not `window.print()` of the host page. Build page-image print output for the paged viewer (DOCX, PPTX, RTF, ODT, static XLSX) so Print outputs only the document's pages at page size, analogous to the spreadsheet viewer's existing print path (print container with per-page images, print media rules, the existing print guard). This is its own work stream. *Revision (R9):* the paged viewer prints through a hidden same-origin iframe that holds only the rendered page images (§6.12). Host content cannot reach the printout. The `allowPrint={false}` guard stays but is scoped to its own instance.
- **Disallowed actions (owner decision):** when `allowPrint={false}` or `allowCopy={false}`, the Print / Copy buttons are **hidden**, not shown disabled. Existing blocking behaviour (Cmd/Ctrl-P intercept, copy blocking) is unchanged.
- **Formats (builder-check call, mine):** every format rendered by the paged viewer gets the toolbar and status line. RTF and ODT behave like Word (page navigation + word count, from their text models). Static (non-interactive) XLSX gets page navigation but no word count.
- **File name fallback (my call):** use `File.name`, the last URL path segment, or a `Content-Disposition` filename when available; the `title` prop overrides; with no name, the slot is hidden. The name travels from the loader (`OfficeFile`) to the doc.
- **Narrow / touch toolbar (my call):** on coarse pointers or narrow panes, the toolbar keeps 44px targets by collapsing Find to an icon button that expands the input, hiding the file name, and moving lower-priority actions (fit, print, copy) into a "More" menu. No control may overlap the pages (extend the mobile no-overlap Playwright checks to the paged viewer).
- **Lazy fallback (my call):** if the toolbar or status line must be lazy-loaded for bundle reasons, their height is reserved by a fixed-size placeholder at first paint so the CLS gate (≤ 0.001) holds. The "Copied" confirmation in the eager paged viewer uses a CSS-only animation (motion.dev stays lazy-only).
  - *Revision (R5):* lazy loading is now mandatory, not a fallback. All new paged UI goes in a new lazy chunk: toolbar, status line, More menu, Find and zoom controls, the Copied confirmation, word count, page tracking and paged print. `LegacyOfficeDoc` renders fixed placeholders that reserve the toolbar's and status line's actual heights at both responsive sizes (regular and compact). One eager constants module holds those heights, and both the placeholders and the chunk read it (§6.10). Keyboard paths (Cmd/Ctrl-C copy, Cmd/Ctrl-F, Cmd/Ctrl-P) stay eager. The Copied confirmation renders in the lazy chrome, so the eager CSS-only animation is no longer needed.
- *Revision (review), narrow toolbar keyboard contract:* the More menu opens with focus on its first item. Arrow keys, Home and End move between items. Escape and outside clicks close the menu and return focus to the More button. Cmd/Ctrl-F expands the collapsed Find and focuses its input. Escape in Find collapses it in compact mode and returns focus to the Find button. On desktop, Escape returns focus to the document region. If the layout crosses the compact threshold while a control has focus, focus moves to that control's counterpart in the new layout, and the query and active match are kept.

### 6.6 Canvas theming
- ~~Chrome-like canvas elements read the theme tokens so dark mode is whole: sheet gridlines (`#d0d0d0` today), selection overlay (`src/core/overlay.ts`, green fill + 2px border), find highlights, freeze divider (`sheetTiles.ts`), tile "Loading…" placeholders, and the document page backdrop (`#888` today). Document content colours (cell fills, text, page paper) do not change.~~
- **Owner decision B (revision), dark mode keeps document content as paper.** The spreadsheet cell area and Word/PowerPoint pages keep their real document colours in both themes. That covers white paper (the worksheet's white fill in `renderViewport.ts` and `sheetTiles.ts`, the page canvas `#fff`), authored fills and text, and light gridlines inside the sheet (`#d0d0d0`). The chrome, headers, menus and the backdrop around the content go dark.
- **Themed canvas colours (R6).** The only themed canvas and backdrop colours are:
  - the backdrop around pages, which is the DOM background of the paged scroll container (`#888` today), and any area of the sheet pane outside the sheet canvas;
  - tile "Loading…" placeholders (`pending`, `pending-text`);
  - the freeze divider (`freeze`);
  - overlays drawn on paper: cell selection (`selection`, `selection-fill`), find highlights (`find`, `find-active`) and paged text selection (`text-selection`).

  Paper sits under every one of these except the backdrop. Their default values are therefore chosen to read on `#ffffff` in **both** themes: strokes ≥ 3:1 against white, and `pending-text` ≥ 4.5:1 against `pending`. Hosts can still override them. `surface` is never used for paper.
- Painters keep today's literal colours as defaults when no colours are passed, so the static `renderSheet` path and the golden corpus are unchanged.
- **Cache keying (R11).** Themed colours that are baked into cached bitmaps form a palette key. That key is part of the sheet tile cache identity and of the paged overlay key. A colour change misses the cache, and stale tiles or pages are never presented.
  - *Round-2 revision:* the palette key lives on the tile key (`TileKey` and its serialized `tileKey()`) and the renderer's frame, not on the shared `CacheIdentity`, which takes part in worker ownership checks.

### 6.7 Bundle strategy
- Moving chrome styles from inline JS objects to the stylesheet shrinks the eager JS; that freed room funds eager-shell icons and the new toolbar. Verify with a real build after every stream; the 38,720 B ceiling is hard. If the new Word/PowerPoint toolbar does not fit eagerly, lazy-load it (the pages can render first).
- *Revision (R5):* the headroom is about 23 B, so the strategy changes to lazy-first.
  1. Before anything else, the first task records a measured baseline. It then moves the paged viewer's existing floating Find, zoom and copy-pill presentation into a new lazy chunk. Those controls float over the pages and take no layout space, so loading them lazily cannot shift layout. After this move `viewerControls` is no longer in the eager graph, and the bytes it frees fund the foundation's eager additions.
  2. Lucide, motion and all new paged UI live only in lazy chunks. A static import-graph test with module provenance enforces this.
  3. Every task records eager gzip before and after. The shared-controls task must come out net ≤ 0.

### 6.8 Theme root and host isolation (R2, R3)
- **Theme root.** Each viewer's existing outer wrapper carries `data-ov-theme-root`, `data-ov-theme` and an instance id (`data-ov-instance`). That means the `LegacyOfficeDoc` wrapper, the `InteractiveSheetDoc` wrapper (already `data-officeview-sheet-instance`) and the lazy-sheet `Suspense` fallback. These wrappers already contain every control: the toolbar, SearchBar, ZoomControls (floating and docked), the copy pill, sheet popups and dialogs, the fixed-position mobile filter sheet (it is in the DOM tree, not portaled) and the status bar. Putting the attributes there adds no DOM layer and keeps the outer wrapper's flex-item contract with the host. Any future portaled or `position: fixed` popup must render inside the theme root or carry the same three attributes.
- `data-officeview-root` keeps its existing semantics: the inner region with the host's `className`/`style`, the print guard target and test hooks. That region carries **no** utility classes and is excluded from the scoped base rules, so a host's `className`/`style` behave exactly as before.
- **Tokens.** Tokens are defined on `:where([data-ov-theme-root])`. The dark set is defined on `:where([data-ov-theme-root][data-ov-theme=dark])` and, inside `@media (prefers-color-scheme: dark)`, on `:where([data-ov-theme-root][data-ov-theme=system])`. `--ov-sentinel: 1` lives in the same rule as the light set.
- **No global emission.** The stylesheet does not import Tailwind's default theme. A self-authored `@theme inline reference { --*: initial; … }` maps only the semantic names (colours, radii, shadows, a few font sizes and easings) to `var(--ov-*)`. `reference` suppresses the `:root, :host` block, and `inline` makes utilities read `var(--ov-…)` directly. Tailwind's own prefixed variable names (`--ov-color-*`), which would collide with the host-facing tokens, are therefore never emitted. Spacing is `--spacing: 4px`, so sizes are pixels and do not depend on the host's `html { font-size }`.
- **Post-processing.** A small build step handles what Tailwind v4 still emits for shadow, border and translate utilities.
  - *Round-2 revision:* the step is a **PostCSS plugin** registered in Vite's `css.postcss.plugins`. That stage runs after `@tailwindcss/vite` has generated the CSS and before Vite consumes it, in the library build, the dev/harness server, `?inline` imports and vitest.
  - A Vite `transform` hook with `enforce: 'post'` would run after `vite:css-post`, see only `""` or JS, and leave `dist/style.css` unscoped.
  - *Round-3 revision:* the plugin identifies the entry by its module id with any `?query` stripped, so `officeview.css?inline` and `?direct` are scoped too. Vite passes the query through to PostCSS.
  - A scratch build (vite 7.3.1, tailwindcss and `@tailwindcss/vite` 4.3.3) confirmed the PostCSS stage rewrites all of those outputs.

  The plugin:
  - it drops the `@property --tw-*` registrations;
  - it rewrites the `@layer properties { @supports … { *, ::before, ::after, ::backdrop { --tw-*: … } } }` fallback into an unconditional, **unlayered** rule placed before the utilities: `[data-ov-theme-root], [data-ov-theme-root] *, [data-ov-theme-root] ::before, [data-ov-theme-root] ::after { --tw-*: … }`;
  - it scopes every utility selector by appending `:is([data-ov-theme-root], [data-ov-theme-root] *)` to its leading class compound;
  - it fails the build if any rule remains whose selector lacks `[data-ov-theme-root]`.
- **Winning over host resets.** Utilities and the scoped base rules are **unlayered**. A scoped utility has specificity (0,2,0), which beats unlayered host type and universal resets (`button { all: unset }`, `* { box-sizing: content-box }`) and single class + type rules (`.app button`). The scoped base, placed before the utilities, restores what `all: unset` removes: `box-sizing: border-box`, font, colour, line height, button and input appearance, border, padding, background and cursor.
  - *Round-2 revision:* the base's form-control rule uses `:is(button, input, select, textarea)`, at specificity (0,1,1).
  - That only ties with a host rule like `.app button`, so it is not relied on. Every viewer control carries explicit `ov:` colour, background, padding and border utilities, which beat the host rule at (0,2,0).
  - A jsdom class check and the hostile-host Playwright leg enforce this in each restyle task.
- We **do not** use Tailwind's `important` flag. The viewer keeps dynamic inline geometry (widths, heights, transforms, canvas sizes), and `!important` utilities would override any inline value on the same property. Hosts customise through tokens, not by overriding utilities.
- **Residual risk, documented in the README.** Host rules with specificity ≥ (0,2,0), or with `!important`, that target viewer elements can still win.
- **Tests (plan Task 1).**
  - A Playwright page loads a hostile host sheet (`button { all: unset }`, `* { box-sizing: content-box }`, `.flex { display: block }`, `html { font-size: 10px }`, `.app button { padding: 20px }`) in both stylesheet orders. The viewer toolbar's control boxes and computed colours must be identical to the same page without the hostile sheet.
  - The same spec parses `dist/style.css` and asserts that every selector contains `[data-ov-theme-root]` and that there is no `@property`, `:root` or `:host`.
  - Two adjacent viewers with different `theme` values and different host overrides each resolve their own tokens, including their floating controls.

### 6.9 Fit-to-width layout mode (R1)
- Today `displayWidth = min(avail, naturalMaxPageWidth)`, where `avail = containerWidth − 2 × 16`, and zoom 1 already means "fit, capped at 1:1". Fit-to-width adds a `fitMode` flag:
  - `avail = max(0, containerWidth − 2 × 16)` (today's guard, kept); `displayWidth = fitMode && avail > 0 ? min(avail, natural × MAX_ZOOM) : min(avail || natural, natural)`. The width is never negative; containers 0–32 px wide use the natural width, as today.
  - Pages are laid out and re-rasterised at that width, which may exceed natural size. The zoom transform stays `{ zoom: 1, panX: 0, panY: 0 }`, so the container stays `overflow: auto`, native scrolling and momentum work, and a pointer drag selects text.
  - Page coordinates keep `scale = displayWidth / spec.widthPx`, so search and selection geometry are unchanged.
  - `zoom.ts` is not changed; its exported zoom semantics stay as they are.
- **Leaving fit mode.** *Round-2 revision: one rule.* Any zoom input turns fit mode off: zoom in/out buttons, Ctrl/Cmd-wheel, pinch and keyboard zoom.
  - The input is applied as if the viewer were on the natural layout at `z0 = fitWidth / naturalDisplayWidth`, with the content point under that input's anchor unchanged. Buttons step `z0` with `stepZoom` (×1.5 or ÷1.5); wheel and pinch apply their ratio to `z0`.
  - A result above 1 is a zoomed transform; a result clamped to 1 is the native natural layout.
  - Example: from `z0 ≈ 1.613`, zoom out gives 1.0756, and a second zoom out gives 1.
  - *Round-3 revision:* gaps between pages are fixed px, so the fit layout is not a uniform scale of the natural layout. The anchor therefore converts by page index plus page-local offset, the offset scaled by `naturalWidth / fitWidth`. A point inside a gap keeps its unscaled gap offset.
  - This replaces the earlier wording that zooming out from fit "lands at natural layout, zoom 1".
- **Fit state.** `fitActive = fitMode`, shown as `aria-pressed` on the Fit button. Container resizes re-lay out while fit mode is on. A zero-width container falls back to the non-fit formula.
- **Memory.** In fit mode only pages near the viewport are rasterised at the full fit width. Off-screen pages keep a backing store no larger than natural mode would give them, using the same visibility observer the zoomed path uses today. The per-page backing scale stays capped by `MAX_EFFECTIVE_SCALE`.

### 6.10 Paged chrome lazy chunk (R5)
- New lazy chunk `components/paged/PagedChrome` exports `PagedToolbar`, `PagedStatus` and, temporarily, `FloatingControls`. `FloatingControls` is the first task's pure move of today's floating controls, and the toolbar task replaces it. Find and zoom controls, the More menu, page tracking, word count and the Copied confirmation all live in this chunk. Paged print is its own lazy module, imported when the user prints.
- **Eager module `components/paged/metrics.ts`.** It holds `PAGED_TOOLBAR_HEIGHT = { regular: 40, compact: 48 }` (compact fits 44 px targets plus 2 px padding), `PAGED_STATUS_HEIGHT = 24`, `COMPACT_PANE_PX = 600` and `isCompact(width, coarse)`.
- **Placeholders.** `LegacyOfficeDoc` renders a top and a bottom placeholder from the first paint. Their heights come from these constants for the current responsive state, and they swap to the real components with identical heights. `showToolbar={false}` renders neither.
- **Tests.** The swap is CLS-tested with a delayed chunk at desktop and at coarse/narrow sizes, across document load and `showToolbar` toggles.
- Placeholder and status geometry is static (driven by constants), so CLS is 0 when the chunk resolves. A Cmd/Ctrl-F press before the chunk loads is queued and focuses Find on mount.

### 6.11 Page tracking and word count (R7, R8)
- **Page tracking.** Page tops and heights come analytically from page specs, `displayWidth` and the gap (the same numbers as `contentHeight`), with no layout reads.
  - On scroll, transform change, resize or page-list change, one rAF-throttled computation picks the current page. That is the page with the **largest visible pixel height** in the viewport, with ties going to the lower index. While zoomed, the viewport is mapped through the transform.
  - *Round-2 revision:* the mapping always includes the container's native `scrollTop`: content y = (view y + scrollTop − padding − panY) / zoom.
  - Mode changes move the offset atomically. Entering zoom turns the native scroll into pan and sets `scrollTop` to 0 in the same commit that sets `overflow: hidden`; leaving zoom turns pan back into `scrollTop`. Both keep the content point under the anchor fixed.
  - So `overflow: hidden` never hides a stale scroll offset.
  - *Round-3 revision:* `goTo` while zoomed uses the exact inverse of that mapping, `panY = scrollTop − padding − top·zoom`.
  - State updates only when the index changes, so fast momentum scrolling produces no extra React renders.
  - This replaces the sparse IntersectionObserver thresholds, which could miss the moment two tall pages swap dominance. It also replaces `intersectionRatio` comparisons, which favour short pages.
  - `goTo(i)` scrolls the viewer's own container (`container.scrollTo`), never `scrollIntoView`, so host ancestors never scroll. While zoomed (`overflow: hidden`), it sets `panY` through `clampPan` instead.
- **Word count** comes from the document text model, with no painting: `DocxDocument`, which DOCX, RTF and ODT all produce.
  - Justification: painting is what made the search-index route a long task. The text model has no line wrapping and no repeated page headers, and it avoids the watermark-muted painting span.
  - Policy, modelled on Word's default count (the expected test counts are derived from this policy, not measured in Word):
    - counted: body paragraphs, table cells (recursively), text-box paragraphs, footnotes and endnotes;
    - not counted: headers, footers and comments.
  - *Round-2 revision, tokens:*
    - A word is a maximal whitespace-delimited run of non-CJK code points containing at least one letter or digit. So `$50`, `100%`, `e-mail`, `3.14` and a URL each count once, and a lone `—` does not count.
    - Each Han, Hiragana or Katakana code point counts as one word. Combining marks continue a token.
    - Korean is counted by spaces.
    - Run boundaries are not word boundaries; paragraph boundaries are.
  - `Intl.Segmenter` is not used: its dictionary segmentation of CJK differs between engines.
  - *Round-2 revision, bounded work:*
    - The model is walked lazily. Runs are never concatenated, and each run's text is fed in chunks of at most 2,000 code units without splitting a surrogate pair.
    - A resumable counter carries the open-token state across chunks, runs and slices, so a giant token, URL or CJK paragraph is both bounded and counted exactly.
    - The slice clock and the cancellation signal are checked after every chunk. Slices take ≤ 8 ms of work and yield between them; the work is cancellable and tied to the document generation.
    - *Round-3 revision:* every node visit is a budgeted step. That includes empty runs, tables, cells and paragraph boundaries, and the traversal is iterative. A slice overshoots 8 ms by at most one step (one node visit or one ≤ 2,000-code-unit chunk), whatever the content.
  - PowerPoint and static XLSX show no word count.

### 6.12 Paged print isolation (R9)
- Print and Cmd/Ctrl-P (when allowed) lazy-load the print module and render each page **in the parent document**. The parent is where the registered fonts live, which a fresh iframe would not have. Pages are rendered one at a time on a single reused canvas, at one uniform scale for the whole job: `PRINT_SCALE` = 2, or else 1, within `PRINT_MAX_AXIS` = 16,384 and a workspace byte cap. A job that does not fit at scale 1 is refused (see the whole-job bound below) rather than printing per-page notices.
- Each page is turned into a PNG blob and an object URL. The module builds a hidden same-origin `<iframe>` inside the theme root whose document contains only those `<img>` elements. Each page size gets a named `@page` rule (`size: Wpx Hpx; margin: 0`), and the module calls `iframe.contentWindow.print()`. Host content is never in the printed document.
  - *Round-2 revision, page size:* where named pages are unsupported (`CSS.supports('page', 'p0')` is false, as in older WebKit), a single `@page` sized to the largest page is used, and each image is fitted inside it.
- *Round-2 revision, whole-job memory bound:*
  - Before anything is allocated, the job picks one uniform scale, 2 or else 1, so that every page fits the workspace and axis caps and the estimated decoded total (Σ W·H·s²·4) is ≤ 1 GiB. Otherwise the job is **refused** with a status-line notice, with no partial printout.
  - Retained PNG blobs are capped at 192 MiB in total. Going over aborts the job with the same notice.
  - The single raster canvas is cleared after each page and zeroed at the end.
  - Images are awaited with `load`, not force-decoded.
  - Every await point (lazy import, `toBlob`, image load) re-checks for abort. Late blobs are never turned into URLs, and failures before the job object exists still clean up everything.
- **Cleanup.** The iframe is removed and its URLs revoked when any of these happens:
  - the frame's or the window's `afterprint` fires;
  - the user starts a new print;
  - the source is replaced;
  - `allowPrint` becomes false;
  - the viewer unmounts.

  Cleanup does not depend on `afterprint` arriving.
- `allowPrint={false}` keeps the hidden Print button, the Cmd/Ctrl-P intercept and a print-media guard. The guard is now scoped to the denying instance's theme root, so an allowed sibling viewer is never hidden. This mirrors the sheet viewer's "hides only itself" rule.
- The sheet viewer's print path is unchanged.

## 7. Left to the builder

- Exact Tailwind version/config details, prefix syntax, file layout of the stylesheet and token files.
- Exact final colour values within the Fluent direction, as long as contrast rules pass in both themes.
- Motion durations/easings (subtle, Fluent-like: ~100–250ms).
- Exact icon choice per action from the Lucide set.
- How to stream the work across agents and in what order, test strategy for visual changes (before/after screenshots can reuse `.worktrees/visual-shots/scripts/capture.mjs`, `compose.py`, `make_index.py`).
- Updating tests that read inline `style.*` values (about 15 files, listed in the survey) to assert behaviour or computed style instead, without weakening the guarantees they protect. *Revision (R4):* jsdom unit tests assert classes, data attributes, ARIA and behaviour. Computed sizes (44 px targets, desktop sizes), colours, contrast and layout are asserted in Playwright against the real compiled stylesheet. Each jsdom geometry or colour assertion that is removed gets a named Playwright equivalent in the same task. Geometry that the canvas mirrors, or that is computed at runtime, stays inline, and its jsdom assertions stay. `vitest.config.ts` gets `@tailwindcss/vite` (and the scoping plugin) explicitly, because vitest has its own plugin list. A setup file injects the compiled stylesheet so the sentinel is present.

## 8. Review revision decisions (2026-10-18)

**Owner decisions**
- **A. Motion is chrome-only for sheet switching and zoom.** The sheet picker, toolbar, status line and zoom readout animate. The grid and page canvases swap instantly. There is no canvas crossfade and no eased canvas zoom. This is consistent with the §3 non-goal (§2.7, §6.4).
- **B. Dark mode keeps document content as paper.** Sheet cells and Word/PowerPoint pages keep white paper, authored fills and text, and light gridlines inside the sheet. Chrome, headers, menus and the backdrop go dark. Overlays drawn on paper (selection, find highlights) stay legible on white in both themes (§3, §6.6).

**Builder calls (design level)**
- **R1.** Fit-to-width is a layout mode (§6.9).
- **R2.** The theme root is the outer viewer wrapper, marked `data-ov-theme-root` (§6.8).
- **R3.** Isolation: no global emission; scoped, unlayered utilities without `important` (§6.8).
- **R4.** Test split between jsdom and Playwright (§7).
- **R5.** All new paged UI goes in a lazy chunk, with placeholders sized from shared constants. The baseline is measured and the existing floating controls move lazy first (§6.7, §6.10).
- **R6.** Canvas theming follows decision B; `surface` is never paper (§6.6).
- **R7.** rAF page tracking from cached offsets, by visible pixel height (§6.11).
- **R8.** Word count from the text model in ≤ 8 ms idle slices, with a Word-compatible policy (§6.11).
- **R9.** Paged print through a hidden iframe of page images; instance-scoped print guard (§6.12).
- **R10.** The dev warning is a runtime check, not env-gated (§5).
- **R11.** A palette key in the tile cache identity and the paged overlay key (§6.6).
- **R12.** Stream order: canvas theming after spreadsheet chrome and paged print, then motion. Shared-file ownership is tabled in the plan.

**Round-2 revisions (codex R2 and opencode R2 on 657b6f9)**

These refine existing decisions; no owner decision changes.
- The CSS scoping runs as a PostCSS stage (§6.8).
- Every viewer control carries explicit colour, background, padding and border utilities (§6.8).
- The fit formula keeps the zero/narrow-width guard, and fit exit follows one rule (§6.9).
- Native scroll and pan move atomically between modes (§6.11).
- Word count uses bounded chunks with a resumable tokenizer (§6.11).
- Print has a whole-job memory bound, in-flight cleanup and a single-size page fallback (§6.12).
- The palette key sits on the tile key, not the shared identity (§6.6).
