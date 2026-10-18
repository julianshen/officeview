# Fluent UI restyle Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Restyle every officeview viewer surface in a Fluent / Office 365 look with Tailwind, Lucide icons, motion, and light/dark/system themes driven by host-overridable CSS tokens; add a Word/PowerPoint top toolbar (page nav, fit-to-width, file name, word count) and a real paged-document print.

**Architecture:** A prefixed Tailwind v4 stylesheet (`dist/style.css`, no preflight) maps semantic utilities onto `--ov-*` CSS custom properties defined on `[data-officeview-root]`, switched by `data-ov-theme` or `prefers-color-scheme`. Components swap inline style objects for classes; canvas painters receive resolved theme colours. Motion lives only in lazy code.

**Tech Stack:** React 19, TypeScript, Vite 7 lib mode, Tailwind CSS v4 (`@tailwindcss/vite`), `lucide-react`, `motion` (motion.dev), Vitest (jsdom), Playwright.

**Spec:** `docs/superpowers/specs/2026-10-18-ui-restyle-design.md` (approved 2026-10-18). Read it with this plan.

## Global Constraints

- Eager bundle: `node -e "console.log(require('zlib').gzipSync(require('fs').readFileSync('dist/index.js')).length)"` after `bun run build` must stay **≤ 38,720**. Rebuild before measuring (no checked-in `dist/`).
- Never import `motion` from eager code (anything reachable from `src/index.ts` without crossing `lazy(() => import('./xlsx/InteractiveSheetDoc'))` or a new lazy boundary).
- Tailwind: class prefix `ov` (v4 `prefix(ov)` → `ov:flex`), **preflight off**, all base rules scoped under `[data-officeview-root]`; no global selectors.
- Tokens: CSS custom properties named `--ov-color-<name>`, `--ov-radius-<name>`, `--ov-shadow-<name>`. Color names exactly: `surface chrome field border text muted icon hover accent on-accent header header-text header-active gridline selection selection-fill backdrop find find-active text-selection pending pending-text freeze`.
- Light/dark starting values from spec §2 item 2; add `find #ffd600 @45%`, `find-active #ff8c00`, `text-selection rgba(64,128,255,.35)`, `pending #eceff3 / #2a2b2d`, `pending-text #666666 / #a19f9d`, `freeze #5f6368 / #a19f9d`. Radii: `control 4px`, `panel 6px`. Shadows: `rest 0 1.6px 3.6px rgba(0,0,0,.13)`, `popover 0 6.4px 14.4px rgba(0,0,0,.18)` (dark: alpha .5/.6).
- Theme prop: `theme?: 'light' | 'dark' | 'system'` (default `'system'`) on `OfficeDocProps` → `data-ov-theme` attribute on every `[data-officeview-root]`.
- New props: `showToolbar?: boolean` (default `true`), `title?: string`. Existing `showSearch`, `showZoomControls`, `allowCopy`, `allowPrint`, `className`, `style`, `background`, `pageGapPx` keep working; `background` overrides `--ov-color-backdrop`.
- `allowPrint={false}` / `allowCopy={false}` **hide** Print / Copy buttons.
- Accessibility invariants: 44px coarse-pointer targets; all existing aria-labels unchanged; icons `aria-hidden="true"`; WCAG text ≥ 4.5:1, UI/track ≥ 3:1 in both themes; CLS ≤ 0.001; honour `prefers-reduced-motion`.
- Process: TDD; every task reviewed by a different model family than its implementer; findings fixed before merge; work on stream branches off `feature/ui-restyle` in `.worktrees/`; never commit/push to `main`. Commit trailers per session.
- Breaking change: hosts must `import 'officeview/style.css'`; version `0.5.0`; CHANGELOG `[Unreleased]` migration note.

## Review Focus

1. **Host page that also uses Tailwind or global resets** — the viewer must look identical and must not restyle host elements. Test: Task 1 Playwright page with a host stylesheet (`* { box-sizing: content-box; } button { all: unset }` and an unprefixed `.flex{display:block}`) asserting viewer toolbar layout and a host `<button>` untouched.
2. **Theme switching while open** (system flips dark↔light, or host changes `theme` prop) — chrome and canvas colours update without remount, tiles/pages repaint, no layout shift. Test: Task 7 jsdom test toggling `theme` prop asserts painter receives new colours; Playwright `emulateMedia({colorScheme})` flip asserts CLS 0.
3. **Missing stylesheet import** — dev warning fires once, viewer still renders content (canvas pages visible). Test: Task 1 unit test.
4. **Tiny/narrow panes and touch** (phone 375px, coarse pointer) — toolbar collapses (Find icon, More menu), 44px targets, nothing overlaps pages/grid. Test: Task 3 and Task 5 mobile Playwright checks.
5. **Huge documents** (300-page DOCX, 1M-row sheet) — word count and page tracking never block first paint or scrolling. Test: Task 3 unit test that word count is computed via idle scheduling and status shows a placeholder until ready; page tracking uses IntersectionObserver, not per-scroll layout reads.

## Streams and ownership

| Stream | Tasks | Implementer | Reviewer (different family) | Depends on |
|---|---|---|---|---|
| A Foundation | 1 | Claude subagent | codex | — |
| B Shared controls + icons | 2 | Claude subagent | agy | A |
| C Paged toolbar + status | 3 | Claude subagent | codex | B |
| D Paged print | 4 | codex | Claude subagent | C |
| E Spreadsheet chrome | 5 | agy | Claude subagent | B |
| F Motion | 6 | Claude subagent | opencode | C, E |
| G Canvas theming | 7 | codex | Claude subagent | A |
| H Verify, docs, release gate | 8 | Claude subagent | codex + agy (visual) | all |

Merge each stream into `feature/ui-restyle` after its review approves, re-run the eager budget, then branch later streams from the updated `feature/ui-restyle`.

---

### Task 1: Foundation — Tailwind, tokens, theme prop, stylesheet export

**Files:**
- Create: `src/styles/officeview.css` (Tailwind entry: `@import "tailwindcss/theme" layer(theme) prefix(ov); @import "tailwindcss/utilities" layer(utilities) prefix(ov);` — no preflight; `@theme` mapping `--color-*`/`--radius-*`/`--shadow-*` to `--ov-*`; token blocks)
- Create: `src/styles/tokens.ts` (TS list of token names + `readThemeColors`)
- Create: `src/components/theme.ts` (`useThemeAttr`, stylesheet sentinel check)
- Modify: `vite.config.ts` (add `@tailwindcss/vite`; `build.lib.cssFileName: 'style'`), `package.json` (deps `tailwindcss`, `@tailwindcss/vite`; exports `"./style.css": "./dist/style.css"`; `sideEffects: ["./dist/worker.js", "*.css"]`; version `0.5.0`), `src/index.ts` (import `./styles/officeview.css` so lib mode emits it), `vitest.config.ts` (`test.css: true`), `src/debug/viewer-e2e.tsx`, `src/debug/viewer-perf.tsx`, `src/demo/main.tsx` (stylesheet import), `vite.harness.config.ts` (allow `src/styles`), `src/components/OfficeDoc.tsx` (`theme` prop, `data-ov-theme` on roots), `src/components/xlsx/InteractiveSheetDoc.tsx` (pass through to its root)
- Test: `tests/ui-theme-foundation.test.tsx`, `tests/browser/ui-host-isolation.spec.ts`

**Interfaces:**
- Produces:
  - `export type ThemeMode = 'light' | 'dark' | 'system'` (src/components/theme.ts)
  - `export const COLOR_TOKENS: readonly ColorToken[]` and `export type ColorToken = 'surface' | … | 'freeze'` (src/styles/tokens.ts, exact list from Global Constraints)
  - `export type ThemeColors = Record<ColorToken, string>`
  - `export function readThemeColors(root: Element): ThemeColors` — reads `getComputedStyle(root).getPropertyValue('--ov-color-' + name).trim()`, falling back to the light defaults when empty
  - `export function useThemeAttr(mode: ThemeMode | undefined): { 'data-ov-theme': ThemeMode }`
  - `export function warnIfStylesMissing(root: Element): void` — once per page, `console.warn('[officeview] style.css not loaded: import "officeview/style.css"')` when `--ov-sentinel` is not `1`; no-op when `import.meta.env.PROD`
  - CSS: `[data-officeview-root]{--ov-sentinel:1; …light tokens…}`, `[data-officeview-root][data-ov-theme=dark]{…dark…}`, `@media (prefers-color-scheme: dark){[data-officeview-root][data-ov-theme=system]{…dark…}}`; base scoped rules: `[data-officeview-root]{font-family: system-ui, -apple-system, "Segoe UI", sans-serif; color: var(--ov-color-text)} [data-officeview-root] :where(button,input,select){font: inherit}`
  - `OfficeDocProps.theme?: ThemeMode`

- [ ] **Step 1: Write failing tests** in `tests/ui-theme-foundation.test.tsx`:
  - `it('applies data-ov-theme=system by default and honours the prop')` — render `OfficeDoc` with a small DOCX fixture; `root.getAttribute('data-ov-theme')` is `'system'`; rerender with `theme="dark"` → `'dark'`; same for an interactive sheet root.
  - `it('readThemeColors returns every token, falling back to light defaults')` — on a bare element: `Object.keys(readThemeColors(el))` equals `COLOR_TOKENS`, `surface === '#ffffff'` (or `#fff` normalised), `accent === '#0f6cbd'`.
  - `it('dark tokens resolve when data-ov-theme=dark')` — inject `src/styles/officeview.css?inline` into a `<style>`, element with `data-officeview-root data-ov-theme="dark"`: `readThemeColors(el).surface === '#292827'`.
  - `it('warns once when the stylesheet is missing, never in production')` — spy `console.warn`; call twice on an element without the sentinel → 1 call; with sentinel → 0.
- [ ] **Step 2:** Run `bunx vitest run tests/ui-theme-foundation.test.tsx` → FAIL (module not found / attribute null).
- [ ] **Step 3:** Install deps (`bun add -d tailwindcss @tailwindcss/vite`), implement the files above, wire the prop.
- [ ] **Step 4:** Run the test file → PASS.
- [ ] **Step 5: Host isolation (Review Focus 1).** `tests/browser/ui-host-isolation.spec.ts`: harness page variant with a host stylesheet `button{all:unset;color:rgb(1,2,3)} .flex{display:block} *{box-sizing:content-box}`; assert a host `<button>` keeps `color: rgb(1, 2, 3)`; assert the viewer root exists and `getComputedStyle(root).getPropertyValue('--ov-sentinel').trim() === '1'`; assert no stylesheet rule in `dist/style.css` has a selector not containing `[data-officeview-root]` or an `ov:` class (parse with a small script in the spec). Run `bunx playwright test --config playwright.viewer.config.ts ui-host-isolation` → PASS.
- [ ] **Step 6: Build + budget.** `bun run build`; assert `dist/style.css` exists; eager gzip ≤ 38,720 (record number).
- [ ] **Step 7:** Full `bunx vitest run` and `bunx tsc --noEmit` green; commit `feat(ui): Tailwind foundation, theme tokens and stylesheet export`.

### Task 2: Shared controls — Lucide icons and classes for SearchBar / ZoomControls / copy pill

**Files:**
- Modify: `src/components/viewerControls.tsx` (classes, Lucide `ChevronUp`, `ChevronDown`, `Plus`, `Minus`, `Search`), `src/components/OfficeDoc.tsx` (copy pill → classes + `Copy`/`Check`/`Ban` icons), `package.json` (`lucide-react` dep)
- Create: `src/components/Icon.tsx` (wrapper: `size`, `aria-hidden`, `className`)
- Test: `tests/viewer-chrome.test.tsx` (update), `tests/xlsx-viewer-chrome-glyphs.test.tsx` (update), `tests/officeview-selection.test.tsx:77` (update)

**Interfaces:**
- Consumes: Task 1 tokens/classes.
- Produces: `export function Icon(props: { icon: LucideIcon; size?: 14 | 16 | 20; className?: string }): JSX.Element` (renders with `aria-hidden="true"`, `focusable="false"`, `strokeWidth={2}`); `export const coarse = (): boolean` exported from viewerControls (replaces private `T()`; `T()` becomes `(px) => coarse() ? Math.max(px, 44) : px`, still used for inline geometry only where needed).

- [ ] **Step 1: Update tests first** (assert behaviour, not inline style):
  - viewer-chrome: replace inline `style.*` checks with `getBoundingClientRect`-free equivalents usable in jsdom: under coarse stub each labelled control has class `ov:min-h-11 ov:min-w-11` (44px) *or* computed `min-height: 44px` with `test.css: true`; keep the "fine sizes unchanged" case by asserting the coarse classes are absent.
  - chrome-glyphs: assert search/zoom/copy controls contain an `svg[aria-hidden="true"]` and **no** text in `↑↓+−✓⌘`; keep the U+FE0E rule for any remaining glyph text.
  - officeview-selection L77: copied state is `button[data-state="copied"]` with accessible name `'Copied'` (aria-live polite region announces "Copied"); blocked state `data-state="blocked"` name `'Copy blocked — press ⌘C'` kept as accessible text.
- [ ] **Step 2:** Run those three files → FAIL.
- [ ] **Step 3:** Implement: Lucide icons via `Icon`, Tailwind classes using tokens (`ov:bg-surface ov:border ov:border-border ov:rounded-control ov:shadow-popover`), hover/press transitions (`ov:transition-colors ov:duration-150 motion-reduce:ov:transition-none`), `data-state` on the copy pill; remove inline colour literals. Keep all aria-labels.
- [ ] **Step 4:** Run the three files → PASS; run `tests/officeview-*.test.tsx` and `tests/xlsx-viewer-zoom-dock.test.tsx` → PASS.
- [ ] **Step 5:** `bun run build`; eager gzip ≤ 38,720 (icons are in eager code here — record before/after; if over, move icon imports behind `React.lazy` for the toolbar created in Task 3 and keep text-free CSS mask icons here as the fallback, noting it in the commit).
- [ ] **Step 6:** Full vitest + tsc green; commit `feat(ui): Lucide icons and token classes for shared viewer controls`.

### Task 3: Paged viewer toolbar and status line

**Files:**
- Create: `src/components/paged/PagedToolbar.tsx`, `src/components/paged/PagedStatus.tsx`, `src/components/paged/usePageTracking.ts`, `src/components/paged/useWordCount.ts`, `src/components/paged/fileName.ts`
- Modify: `src/components/OfficeDoc.tsx` (`LegacyOfficeDoc`: render toolbar above the scroll container and status below it; remove floating SearchBar/ZoomControls/copy pill placement in favour of toolbar slots; `title`, `showToolbar`; fit-to-width), `src/components/OfficeFile.tsx` (derive file name, pass `title` unless the host set one), `src/core/stream.ts` (expose `Content-Disposition` filename from the fetch path), `src/core/zoom.ts` (allow fit zoom ≥ 1 computed from width)
- Test: `tests/paged-toolbar.test.tsx`, `tests/paged-file-name.test.ts`, `tests/browser/paged-toolbar.spec.ts`

**Interfaces:**
- Consumes: Task 2 `Icon`, `coarse()`, restyled `SearchBar` (rendered inside the toolbar via a `variant: 'floating' | 'inline'` prop added in this task, default `'floating'` for back-compat).
- Produces:
  - `export function fileNameFrom(source: unknown, disposition?: string | null): string | undefined` — `File.name`; `Request`/`Response`/`HttpSource` URL last path segment (decoded, query/hash stripped); RFC 6266 `filename*=`/`filename=` from `disposition` wins over URL; `undefined` otherwise.
  - `export function usePageTracking(container: RefObject<HTMLElement | null>, pageCount: number): { current: number; goTo(index: number): void }` — IntersectionObserver on page canvases (threshold `[0, .5, 1]`), current = most-visible page; `goTo` scrolls the page canvas to `block: 'start'`.
  - `export function useWordCount(document: OfficeDocSource, ensureIndex: () => Promise<TextIndex>): number | null` — schedules via `requestIdleCallback` (fallback `setTimeout(…, 200)`) after first paint; counts `\p{L}[\p{L}\p{N}'’-]*` and CJK ideographs as 1 word each over all `IndexLine.text`; returns `null` until ready. Not used for PowerPoint.
  - `PagedToolbar` props: `{ title?: string; kind: 'document' | 'slides' | 'sheet-pages'; current: number; total: number; onPrev(): void; onNext(): void; zoom: number; onZoomIn(): void; onZoomOut(): void; onFit(): void; fitActive: boolean; search?: ReactNode; onCopy?: () => void; copyState: 'idle' | 'done' | 'failed'; onPrint?: () => void; compact: boolean }` — Copy/Print omitted when their callbacks are undefined (disallowed).
  - `PagedStatus` props: `{ kind; current: number; total: number; words: number | null }` — text: document `Page 3 of 12 · 2,418 words` (`…words` placeholder `— words` while null, fixed `min-width` so no shift), slides `Slide 3 of 12`, sheet-pages `Page 3 of 12`.
  - Fit-to-width: `fitZoom = (containerWidth - 2 * padding) / displayWidth` (≥ 1 by construction); `onFit` sets `{zoom: fitZoom, panX: 0, panY: 0}`; `fitActive = |zoom - fitZoom| < 0.01`.
  - Kind: `'slides' in doc → 'slides'`, `'sections' in doc → 'document'`, else `'sheet-pages'` (static XLSX).
  - Compact mode (`coarse() || containerWidth < 600`): file name hidden; Find collapses to a `Search` icon button that expands the inline search; Fit/Copy/Print move into a `More` (`Ellipsis` icon) menu button (`aria-haspopup="menu"`); all targets 44px.
  - Toolbar height fixed `40px` (`44px` compact) and rendered on first paint (even while pages load) so CLS stays 0.

- [ ] **Step 1: Write failing tests** (`tests/paged-toolbar.test.tsx`, `tests/paged-file-name.test.ts`):
  - `fileNameFrom`: `new File([], 'Q3.docx')` → `'Q3.docx'`; `{url:'https://x/a/b/Plan%20v2.pptx?sig=1'}` → `'Plan v2.pptx'`; disposition `attachment; filename*=UTF-8''r%C3%A9sum%C3%A9.docx` → `'résumé.docx'`; `new Uint8Array()` → `undefined`.
  - `it('renders title, page position and word count for a DOCX')` — 3-page DOCX fixture with `title="Report.docx"`: toolbar text contains `Report.docx` and `1 / 3`; status `Page 1 of 3 · — words`, then (after flushing idle) `Page 1 of 3 · N words` with N equal to a hand-counted fixture value.
  - `it('uses Slide wording and no word count for PPTX')` → `Slide 1 of N`, no "words".
  - `it('hides Print and Copy when disallowed')` — `allowPrint={false} allowCopy={false}`: no buttons named `Print` / `Copy selection`.
  - `it('showToolbar=false hides toolbar and status; showSearch/showZoomControls hide their groups')`.
  - `it('fit-to-width sets zoom so the page fills the pane')` — container 1000px, page natural width 600 → after `Fit to width` click, content transform contains `scale(1.6` (= (1000-32)/600 ≈ 1.613).
  - `it('next/prev move the current page and update the status')` — mock IntersectionObserver; `Next page` calls `scrollIntoView` on canvas index 1 and status shows `Page 2 of 3` after the observer reports it.
  - `it('compact mode collapses Find and moves Fit/Copy/Print into More')` — stub coarse: `More` button present with `aria-haspopup="menu"`; opening it lists `Fit to width`, `Copy selection`, `Print`.
  - Review Focus 5: `it('word count waits for idle and never runs during first render')` — spy `ensureIndex`; not called synchronously during render; called after idle flush.
- [ ] **Step 2:** Run → FAIL.
- [ ] **Step 3:** Implement the files and wiring. Keep the Ctrl-wheel/keyboard behaviour; Cmd/Ctrl-F focuses the toolbar search (expanding it in compact mode).
- [ ] **Step 4:** Run → PASS; run `tests/officeview-*.test.tsx`, `tests/viewer-chrome.test.tsx` → PASS (update floating-SearchBar expectations to the inline toolbar where the paged viewer is concerned; sheet viewer keeps its own placement).
- [ ] **Step 5: Browser (Review Focus 4).** `tests/browser/paged-toolbar.spec.ts` on `viewer-e2e.html` with a DOCX and a PPTX, at 1280×800 and 375×667 coarse, Chromium + WebKit: toolbar and status visible; no toolbar/status rect intersects any page canvas rect; every toolbar control ≥ 44px in coarse; CLS during open + scroll ≤ 0.001 (reuse `viewer-layout-shift` measurement helper).
- [ ] **Step 6:** Build; eager gzip ≤ 38,720. If over: put `PagedToolbar`/`PagedStatus` behind `lazy()` with a fixed-height placeholder `<div style={{height: 40}} data-officeview-toolbar-placeholder>` and re-measure.
- [ ] **Step 7:** Full vitest, tsc, browser suite green; commit `feat(ui): paged viewer toolbar, page navigation, fit-to-width, file name and word count`.

### Task 4: Paged document print

**Files:**
- Create: `src/components/print/printPages.ts` (shared print helpers extracted from the sheet viewer: `printCss(instance)`, `claimPrint`, `yieldPrint`, `sizePrintCanvas`, constants `PRINT_SCALE = 2`, `PRINT_MAX_AXIS = 16384`)
- Modify: `src/components/xlsx/InteractiveSheetDoc.tsx` (import the extracted helpers; behaviour unchanged), `src/components/OfficeDoc.tsx` (`LegacyOfficeDoc` print flow + print container; Cmd/Ctrl-P triggers it when allowed)
- Test: `tests/paged-print.test.tsx`, existing sheet print tests unchanged

**Interfaces:**
- Consumes: `Paintable[]` (`p.spec`, `p.paint`) already held by `LegacyOfficeDoc`; Task 3 `onPrint` slot.
- Produces: `export function printCss(instance: string, attr: 'sheet' | 'paged'): string`; `export function claimPrint(instance: string): void`; `export function yieldPrint(instance: string): void`; `export function sizePrintCanvas(canvas: HTMLCanvasElement, cssW: number, cssH: number, scale: number): number` (returns the effective scale after `PRINT_MAX_AXIS` clamping). Paged print DOM: `<div data-officeview-paged-print aria-hidden="true">` with one `<canvas>` per page, `break-before: page` from the 2nd page, page box size = page spec size; screen tree marked `data-officeview-paged-screen`.
- Behaviour: Print button / Cmd-P → render every page at `PRINT_SCALE` into print canvases (bitmap-budget admission like the sheet path: if a page cannot be admitted, print its notice `<p>` instead) → `claimPrint` → `window.print()`; `afterprint` → release canvases (zero their size) and `yieldPrint`.

- [ ] **Step 1: Failing tests** `tests/paged-print.test.tsx`: `it('Print renders one print canvas per page and calls window.print once')` (stub `window.print`, 3-page DOCX → 3 canvases under `[data-officeview-paged-print]`, `print` called 1×); `it('afterprint releases print canvases')` (dispatch `afterprint` → canvases width 0 or removed); `it('print media hides screen chrome and shows only pages')` (parse injected print `<style>`: `@media print` hides `[data-officeview-paged-screen]` and `[data-officeview-chrome]`, shows `[data-officeview-paged-print]` for this instance only); `it('allowPrint=false keeps the guard and has no Print button')`; `it('two viewers: only the printing one is shown')` (claim owner hides the other instance).
- [ ] **Step 2:** Run → FAIL.
- [ ] **Step 3:** Extract helpers (pure move first, run sheet print tests), then implement paged print.
- [ ] **Step 4:** Run new + all `tests/*print*` + `tests/viewer-chrome.test.tsx` → PASS.
- [ ] **Step 5:** Browser: add to `paged-toolbar.spec.ts` a print-media emulation check (`page.emulateMedia({media:'print'})` after triggering prepare; screenshot shows only pages, white background).
- [ ] **Step 6:** Build + budget; full suites; commit `feat(ui): proper paged-document print`.

### Task 5: Spreadsheet chrome restyle

**Files:**
- Modify: `src/components/xlsx/InteractiveSheetDoc.tsx` (toolbar, actions → icon buttons with labels, Print hidden when disallowed, remove `data-officeview-sheet-font` runtime rule except per-instance print rules, theme attr), `SheetNavigator.tsx`, `GridHeaders.tsx` (Lucide `ArrowUpDown`, `ArrowUp`, `ArrowDown`, `Loader2` for pending, `Filter`, `Ellipsis`, `SquareDashedMousePointer`/`Grid2x2` corner; `HEADER_COLORS` → token classes, keep exported constant as token-name map for tests), `GridSemantics.tsx` (popup + status bar), `ViewerStatus.tsx`, `FilterDropdown.tsx`, `SortDialog.tsx`, `ConditionEditor.tsx`, `controlUi.ts`, `sheetResize.tsx` (badges → `MoveVertical`/`MoveHorizontal` icons + number), `DockedZoomControls`
- Test: update the inline-style/glyph tests listed in the spec survey: `xlsx-viewer-policy`, `-sort-ui`, `-touch-header`, `-events`, `-resize`, `merge-ce-resize-panes`, `-sort-pending`, `-filter-ui`, `fix-c-scroll-into-view`, `-a11y`, `-visual-review`, `-zoom-dock`, `-chrome-glyphs`, `-frozen-panes-paint`

**Interfaces:**
- Consumes: Task 1 tokens/classes, Task 2 `Icon`/`coarse()`.
- Produces: unchanged component public behaviour and aria-labels; geometry that the canvas mirrors (row heights 20/36, header 88px touch, status bar `STATUS_BAR_HEIGHT = 24`, 44px targets) stays as inline style **only** where it is dynamic/computed; static chrome styling moves to classes. Test rule: geometry assertions on dynamic inline values (transforms, computed widths/heights) stay; assertions on static colours/positions move to class or `getComputedStyle` checks (`test.css: true`). Glyph assertions (`'↕'`, `'↖'`, `'…'`, `'↑'`, `'↕2'`) become icon assertions: `svg[data-lucide-icon="arrow-up"]` via a `data-icon` attribute that `Icon` sets from the icon's display name, plus unchanged aria-labels; touch resize badge text becomes `'2'` with `aria-label="Resize row 2"`.

- [ ] **Step 1:** Update the listed tests to the rule above (each file: replace glyph/colour literal checks; keep every geometry, target-size, contrast and aria assertion). Add `it('header and status colours meet contrast in light and dark')` in `xlsx-viewer-visual-review` computing contrast from `readThemeColors` for both themes (text ≥ 4.5, track ≥ 3).
- [ ] **Step 2:** Run the listed files → FAIL.
- [ ] **Step 3:** Implement the restyle (Fluent: toolbar `ov:bg-chrome ov:border-b ov:border-border`, controls `ov:rounded-control`, menus/dialogs `ov:rounded-panel ov:shadow-popover ov:bg-surface`, mobile filter sheet `ov:rounded-t-[12px]`), Print hidden when `!allowPrint`.
- [ ] **Step 4:** Run listed files + `tests/xlsx-viewer-*.test.ts*` → PASS.
- [ ] **Step 5:** Browser: `interactive-viewer-visual`, `-a11y`, `-mobile`, `-resize`, `viewer-layout-shift`, `-dpr` pass in Chromium + WebKit; add a dark-scheme run (`emulateMedia({colorScheme:'dark'})`) to the visual spec's contrast and no-overlap checks.
- [ ] **Step 6:** Build + budget; commit `feat(ui): Fluent restyle of the spreadsheet viewer chrome`.

### Task 6: Motion

**Files:**
- Create: `src/components/motion/presets.ts` (lazy-only module: `LazyMotion` + `domAnimation` feature loader, variants `popover`, `dialog`, `sheet`, `fade`, durations 120–220 ms, easing `[0.1, 0.9, 0.2, 1]` (Fluent decelerate))
- Modify: `FilterDropdown.tsx`, `SortDialog.tsx`, `ConditionEditor.tsx`, `GridSemantics.tsx` popup, `GridHeaders.tsx` disclosure group, `InteractiveSheetDoc.tsx` (sheet switch crossfade of the pane, zoom easing of the content transform), `PagedToolbar.tsx` More menu (only if `PagedToolbar` is lazy; otherwise CSS keyframes), `OfficeDoc.tsx` copy confirmation (CSS-only keyframes `ov-pop`)
- Test: `tests/ui-motion.test.tsx`, chunk-graph assertion

**Interfaces:**
- Consumes: Task 3/5 components.
- Produces: `export const popover: Variants`, `dialog`, `sheet`, `fade`; `export function MotionRoot({children}): JSX.Element` (wraps `LazyMotion features={domAnimation} strict`); all components use `m.div` (not `motion.div`) inside `MotionRoot`.
- Rules: transform/opacity only; `useReducedMotion()` → variants with `duration: 0`; exit animations via `AnimatePresence`; no animation on first mount of the viewer (avoid shift).

- [ ] **Step 1: Failing tests** `tests/ui-motion.test.tsx`: `it('motion is not reachable from the eager entry')` — use `tests/helpers/chunk-graph.ts` on a fresh build: no chunk statically imported by `dist/index.js` contains `motion` / `framer`; `it('reduced motion uses zero-duration variants')` — mock `matchMedia('(prefers-reduced-motion: reduce)')` true → filter dropdown opens with no `transform` transition (variant duration 0); `it('popover animates only transform and opacity')` — assert variant keys ⊆ `{opacity, scale, x, y, transition}`.
- [ ] **Step 2:** Run → FAIL. **Step 3:** `bun add motion`; implement. **Step 4:** Run → PASS.
- [ ] **Step 5:** Browser `viewer-layout-shift.spec.ts` (open, scroll, sort, filter, close) CLS ≤ 0.001 in light and dark; `paged-toolbar.spec.ts` CLS ≤ 0.001.
- [ ] **Step 6:** Build + budget (eager unchanged); record lazy-chunk gzip growth in the commit message; commit `feat(ui): motion for menus, dialogs, sheet switch, zoom and feedback`.

### Task 7: Canvas theming

**Files:**
- Modify: `src/xlsx/view/renderViewport.ts` (`ViewportPaintOptions.colors?: Pick<ThemeColors,'gridline'|'pending'|'pending-text'|'surface'>`; replace `#d0d0d0`, `COARSE_FILL`, `#666666`, white), `src/components/xlsx/sheetTiles.ts` (`SheetFrame.colors`; `PENDING_FILL`, `#666666`, `#ffffff`, freeze `#5F6368`), `src/core/overlay.ts` (`paintGridSelectionOverlay(ctx, geometry, scale = 1, colors?: {selection: string; 'selection-fill': string})`), `src/components/xlsx/selectionOverlay.ts` (`paintFindOverlay(ctx, highlights, active, colors?)`), `src/components/OfficeDoc.tsx` (`paintHighlights` colours; backdrop via `--ov-color-backdrop`, `background` prop overrides), `src/components/xlsx/InteractiveSheetDoc.tsx` (resolve colours with `readThemeColors(root)` once per theme change; invalidate tile cache + repaint on change)
- Create: `src/components/useThemeColors.ts`
- Test: `tests/ui-canvas-theme.test.tsx`

**Interfaces:**
- Consumes: Task 1 `readThemeColors`, `ThemeColors`.
- Produces: `export function useThemeColors(root: RefObject<Element | null>, mode: ThemeMode): ThemeColors` — recomputes on mount, on `mode` change, and on `matchMedia('(prefers-color-scheme: dark)')` `change`; returns a stable object identity until values change.
- Rule: document content colours (`src/xlsx/render.ts` static `renderSheet`, cell fills, text, page paper `#fff`) are **not** themed. Defaults (no `colors` passed) keep today's literal values exactly, so existing goldens and painter tests are unchanged.

- [ ] **Step 1: Failing tests** `tests/ui-canvas-theme.test.tsx`: `it('painters use passed theme colours and fall back to legacy literals')` — recording canvas context: `paintGridSelectionOverlay` with no colours strokes `#008000`, with `{selection:'#4cb87a'}` strokes `#4cb87a`; same for `paintFindOverlay` and `renderSheetViewport` gridline (`#d0d0d0` default). `it('switching theme repaints with new colours without remount')` (Review Focus 2) — render interactive sheet with `theme="light"`, spy the frame painter, rerender `theme="dark"` → next paint receives `colors.gridline === '#3b3a39'` and the root element is the same node.
- [ ] **Step 2:** FAIL. **Step 3:** Implement. **Step 4:** PASS; run `tests/golden.test.ts` locally on the Mac (unchanged) and all painter tests.
- [ ] **Step 5:** Browser: dark-scheme screenshot of the interactive and paged viewers shows dark backdrop and gridlines; CLS 0 on `emulateMedia` flip.
- [ ] **Step 6:** Build + budget; commit `feat(ui): theme-aware canvas chrome colours`.

### Task 8: Verification, docs and release gate

**Files:**
- Modify: `README.md` (Usage: `import 'officeview/style.css'` after the import snippet; props snippet adds `theme`, `showToolbar`, `title`; new `### Theming` section before "Interactive spreadsheet viewer" listing every `--ov-*` token with light/dark defaults and an override example; Copy and print section: hidden buttons, paged print), `CHANGELOG.md` (`## [Unreleased]` — Breaking: stylesheet import; Added: themes, toolbar, page nav, fit, word count, paged print, motion; Changed: Fluent look), `docs/validation/` (new `ui-restyle.md` with gate results and screenshots index)
- Use: `.worktrees/visual-shots/scripts/capture.mjs`, `compose.py`, `make_index.py` for before (`main`) / after (`feature/ui-restyle`) composites, light and dark, desktop and 375×667 touch, for DOCX, PPTX, static XLSX, interactive XLSX (filter, sort dialog, mobile filter sheet, zoom 50/200)

- [ ] **Step 1:** Capture before/after composites; publish a gallery for the owner.
- [ ] **Step 2:** Visual review by agy and codex on the composites (style consistency, Fluent fidelity, dark-mode legibility, overlap, truncation); fix findings in the owning stream's files.
- [ ] **Step 3:** Gates on the final branch: `bunx tsc --noEmit`, `bun run typecheck:scripts`, `bunx vitest run`, `bun run build` + eager gzip ≤ 38,720, `bun scripts/golden-corpus.ts compare --threshold 0 --max-ratio 0` (baseline-identical), `bunx playwright test --config playwright.viewer.config.ts` (Chromium + WebKit) — record numbers in `docs/validation/ui-restyle.md`.
- [ ] **Step 4:** Docs/CHANGELOG/README updates; commit `docs: theming, stylesheet import and UI restyle validation`.
- [ ] **Step 5:** Open PR `feature/ui-restyle` → `main` with the breaking-change note, gallery link, gate table and review trail; owner merges.
