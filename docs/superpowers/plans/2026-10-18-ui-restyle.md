# Fluent UI restyle Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Restyle every officeview viewer surface in a Fluent / Office 365 look, using Tailwind, Lucide icons, chrome-only motion and light/dark/system themes driven by host-overridable CSS tokens. Add a Word/PowerPoint top toolbar (page nav, fit-to-width, file name, word count), a status line and an isolated paged-document print.

**Architecture:**
- **Stylesheet.** A prefixed Tailwind v4 stylesheet, `dist/style.css`, has no preflight and no global rules. A build-time scoping step guarantees that every selector contains `[data-ov-theme-root]`.
- **Theme root.** Each viewer's outer wrapper is the theme root (`data-ov-theme-root`, `data-ov-theme`, `data-ov-instance`). It holds the `--ov-*` tokens and contains all chrome.
- **Paper.** Document content stays paper in both themes (owner decision B). Only the backdrop, placeholders, the freeze line and overlays on paper read canvas tokens.
- **Lazy-first.** New paged UI lives in a lazy `PagedChrome` chunk, with placeholders sized from shared eager constants. Lucide and motion are lazy-only, and a static import-graph check enforces it.

**Tech Stack:** React 19, TypeScript, Vite 7 lib mode, Tailwind CSS v4 (`@tailwindcss/vite`, `postcss` for the scoping step), `lucide-react`, `motion` (motion.dev, `LazyMotion` + `m`), Vitest (jsdom), Playwright (Chromium + WebKit).

**Spec:** `docs/superpowers/specs/2026-10-18-ui-restyle-design.md` (owner-approved 2026-10-18, revised after review: owner decisions A–B, builder calls R1–R12 in spec §8). Read it with this plan.

## Global Constraints

- **Eager budget.**
  - Run `bun run build && bun scripts/eager-budget.ts`. The script prints the Node-zlib gzip of `dist/index.js` and exits 1 above **38,720**.
  - Every task records **before and after** numbers in its commit message. Task 2 must come out net ≤ 0.
  - Always rebuild before measuring; the checked-in `dist/` is stale.
- **Eager graph.**
  - Run `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts`. It fails if the static-import closure of `index.js` contains any of: `node_modules/lucide-react/`, `node_modules/motion/`, `node_modules/framer-motion/`, `node_modules/motion-dom/`, `src/components/viewerControls.tsx`, `src/components/Icon.tsx`, `src/components/motion/`, `src/styles/tokens.ts`.
  - It also fails on anything under `src/components/paged/` except `metrics.ts`, the only eager file in that folder. The other new eager helpers (`pointer.ts`, `fileName.ts`, `scrollPan.ts`) live directly in `src/components/`, outside `paged/`. Introduced in Task 0.
- **CSS isolation (spec §6.8).**
  - `src/styles/officeview.css` imports `tailwindcss/theme.css` with `prefix(ov) reference` and `tailwindcss/utilities.css` with `prefix(ov)` and no `layer()`. A self-authored `@theme inline reference { --*: initial; … }` follows.
  - No preflight and no `important` flag. Spacing is `--spacing: 4px`, so sizes are in px.
  - **Scoping stage (R2-1).** `ovCssScope` is a **PostCSS plugin**, `scripts/postcss-ov-scope.ts`. It is registered once in a shared `ovCss` config object (`scripts/ov-css-config.ts`, exporting `{ plugins: [tailwindcss()], css: { postcss: { plugins: [ovCssScope()] } } }`), which `vite.config.ts` and `vitest.config.ts` spread.
    - `vite.harness.config.ts` inherits it through `mergeConfig(baseConfig, …)` and must **not** add a second copy.
    - Why this stage: `@tailwindcss/vite` generates the CSS first, then Vite's `vite:css` plugin runs `css.postcss.plugins` on the generated CSS before the CSS is consumed. A Vite `transform` hook with `enforce: 'post'` runs after `vite:css-post` and only ever sees `""` or JS (codex's R2 reproduction).
    - Verified with a scratch build (vite 7.3.1, tailwindcss and `@tailwindcss/vite` 4.3.3): a PostCSS plugin at this stage rewrote the emitted `dist/style.css`, the `?inline` string inside the JS bundle and the dev-server responses for `s.css`, `s.css?direct` and `s.css?inline`, and removed every `@property` rule in each of them.
    - The plugin only touches its input when `root.source.input.file` ends with `src/styles/officeview.css`.
  - The shipped CSS contains no `:root`, `:host`, `@property` or unscoped selector. `bun scripts/audit-css.ts dist/style.css` (built on `unscopedSelectors`) is part of every build gate.
- **Theme root.**
  - The theme root is the outer viewer wrapper; `data-officeview-root` keeps its existing semantics: the inner region that receives the host's `className`/`style`, with no `ov:` classes.
  - Tokens live under `:where([data-ov-theme-root])`, the dark set under `:where([data-ov-theme-root][data-ov-theme=dark])`, plus a `prefers-color-scheme` block for `data-ov-theme=system`.
  - `--ov-sentinel: 1` sits in the light block.
- **Tokens.** Names follow `--ov-color-<name>`, `--ov-radius-<name>` and `--ov-shadow-<name>`. The colour list is a superset of the spec's illustrative list (`field` and `header-text` are needed for contrast and Find coverage). `src/styles/tokens.ts` `COLOR_TOKENS` holds it, and the README table must match it exactly (Task 8 test).
  - Chrome: `surface chrome field border text muted icon hover accent on-accent header header-text header-active gridline backdrop` (`gridline` = DOM header dividers only).
  - On paper (canvas tokens, `CANVAS_TOKENS`): `selection selection-fill find find-active text-selection pending pending-text freeze`.
- **Default values.**
  - Chrome tokens: light/dark from spec §2 item 2, with `field #ffffff/#1b1a19` and `header-text #323130/#d2d0ce`.
  - Canvas tokens (legible on `#fff` in both themes):

    | Token | Light | Dark |
    |---|---|---|
    | `selection` | `#107c41` | same |
    | `selection-fill` | `rgba(16,124,65,.12)` | same |
    | `find` | `rgba(255,214,0,.45)` | same |
    | `find-active` | `#ca5010` | same |
    | `text-selection` | `rgba(64,128,255,.35)` | same |
    | `pending` | `#eceff3` | `#e1dfdd` |
    | `pending-text` | `#605e5c` | same |
    | `freeze` | `#5f6368` | same |

  - Radii: `control 4px`, `panel 6px`. Shadows: `rest 0 1.6px 3.6px rgba(0,0,0,.13)`, `popover 0 6.4px 14.4px rgba(0,0,0,.18)` (dark alpha .5/.6).
  - Painters keep today's literals as defaults when no colours are passed.
- **Paper rule (decision B).** Never theme the worksheet white (`renderViewport.ts` paintable fill, `sheetTiles.ts` frame clear), canvas gridlines `#d0d0d0`, the page canvas `#fff`, or authored fills and text. `surface` is never paper.
- **Props.**
  - New: `theme?: 'light' | 'dark' | 'system'` (default `'system'`), `showToolbar?: boolean` (default `true`) and `title?: string` on `OfficeDocProps`; `OfficeFile` passes them through.
  - Existing props keep working: `showSearch`, `showZoomControls`, `allowCopy`, `allowPrint`, `className`, `style`, `background` and `pageGapPx`. A host-provided `background` overrides `--ov-color-backdrop`.
- **Disallowed actions.** `allowPrint={false}` / `allowCopy={false}` **hide** the Print / Copy buttons; the blocking behaviour is unchanged.
- **Motion (decision A).** Motion is chrome-only. Grid and page canvases and the content transform change in one step, and no exiting canvas is ever retained.
- **Accessibility invariants:**
  - 44px coarse-pointer targets via the CSS `coarse` custom variant;
  - all existing aria-labels unchanged; icons `aria-hidden="true"`;
  - WCAG text ≥ 4.5:1 and UI/track ≥ 3:1 in both themes;
  - visible focus ring;
  - CLS ≤ 0.001;
  - `prefers-reduced-motion` honoured.
- **Test split (R4).**
  - jsdom asserts classes, data attributes, ARIA and behaviour. Computed sizes, colours, contrast and layout are asserted in Playwright with the real stylesheet.
  - A removed jsdom geometry or colour assertion gets a Playwright equivalent named `<original name> (browser)` in the same task. The commit message lists the mapping.
  - Geometry the canvas mirrors, or that is computed at runtime, stays inline, and its jsdom assertions stay.
  - `tests/setup/styles.ts` injects the compiled stylesheet in every jsdom test.
- **Process.**
  - TDD.
  - Every task is reviewed by a different model family than its implementer, and findings are fixed before merge.
  - Stream branches come off `feature/ui-restyle` in `.worktrees/`; never commit or push to `main`. Commit trailers per session.
- **Breaking change.** Hosts must `import 'officeview/style.css'`. Version `0.5.0`, with a CHANGELOG `[Unreleased]` migration note.

## Review Focus

1. **Hostile host CSS** (global resets, the host's own Tailwind, rem changes). The viewer must render identically and must not restyle the host. Tests (Task 1, `ui-host-isolation.spec.ts`):
   - the toolbar geometry and colours are identical with and without the hostile sheet, in both load orders;
   - parsing `dist/style.css` and the harness-served stylesheet finds every selector scoped, and so does the vitest `?inline` pipeline test, which fails if the scoping plugin sits at a stage that never sees the CSS;
   - host elements are untouched.
2. **Theme switching and host overrides while open.** Chrome and canvas overlays update without a remount, paper stays white and tile/page caches miss on a palette change. Tests:
   - Task 1: two adjacent viewers, and the host token override;
   - Task 7: `a palette change misses the tile cache`, `paged overlay repaints when canvas colours change without remount`, `paper stays paper`, and Playwright dark pixels plus CLS on an `emulateMedia` flip.
3. **Missing stylesheet import.** The warning fires once even in a library build, and content still renders. Test: Task 1 `warns once per page when the sentinel is missing` plus the dist string check.
4. **Narrow panes, touch and keyboard.** The toolbar collapses, targets are 44px, nothing overlaps the visible pages and keyboard/focus work. Tests: Task 3 jsdom keyboard contract and `paged-toolbar.spec.ts` (overlap of clipped regions, 44px, Tab order and visible focus); Task 5 mobile specs.
5. **Huge documents.** Word count, page tracking and fit mode never block. Tests (Task 3):
   - `never runs more than 8 ms of work per slice`;
   - `updates at most once per animation frame and only on change`;
   - `fit mode keeps off-screen page backing stores at natural-mode size`;
   - `a giant single token (2,000,000 chars, no whitespace) is bounded and counts once`;
   - `native scroll to a later page, then zoom: tracking and goTo stay correct`;
   - rapid-scroll browser check.
6. **Print isolation, memory and lifecycle.** Only document pages print, the whole job stays within budget, and cleanup never depends on `afterprint` or on the job having finished. Tests (Task 4):
   - `prints from a hidden same-origin iframe holding only page images`;
   - `planPrintJob picks one uniform scale and refuses beyond the whole-job budget`;
   - `exceeding the blob budget mid-job aborts, disposes and notifies`;
   - the in-flight abort tests;
   - `cleans up on …`;
   - `guard hides only this instance`.
7. **Bundle and lazy boundary.** The eager budget holds, and placeholders reserve the exact heights. Tests: Task 0 eager-graph mutation test and the budget script; Task 3 `placeholders reserve regular and compact heights` plus the delayed-chunk CLS browser test.
8. **Reduced motion and chrome-only motion.** Tests: Task 6 `sheet switch swaps the grid canvas in one commit and never retains an exiting canvas`, plus the Playwright reduced-motion leg.

## Streams, order and file ownership

| Stream | Tasks | Implementer | Reviewer (different family) | Depends on |
|---|---|---|---|---|
| A Foundation | 0, 1 | Claude subagent | codex | — |
| B Shared controls + icons | 2 | Claude subagent | agy | A |
| C Paged toolbar + status | 3 | Claude subagent | codex | B |
| E Spreadsheet chrome | 5 | agy | Claude subagent | B |
| D Paged print | 4 | codex | Claude subagent | C |
| G Canvas theming | 7 | codex | Claude subagent | D, E |
| F Motion (chrome only) | 6 | Claude subagent | opencode | G |
| H Verify, docs, release gate | 8 | Claude subagent | codex + agy (visual) | all |

**Merge order:** A → B → C → E (C and E may be developed in parallel; they share no source files, and C merges first) → D → G → F → H.
- Each stream branches from the latest `feature/ui-restyle` after its dependencies merge.
- After **every** merge, re-run the combined gates: `bunx tsc --noEmit`, `bunx vitest run`, budget, eager check, and `bunx playwright test --config playwright.viewer.config.ts`. Individual stream approvals do not prove the combination works.

**Shared files (owners in merge order; nobody else edits them):**

| File | Owners |
|---|---|
| `src/components/OfficeDoc.tsx` | A(0,1) → C → D → G |
| `src/components/paged/*` lazy files (`PagedChrome`, `PagedToolbar`, `PagedStatus`, `MoreMenu`) | A(0) → B → C → D → F |
| `src/components/viewerControls.tsx` | B → C → F (E must not edit it) |
| `src/components/xlsx/InteractiveSheetDoc.tsx` | A(1) → E → G → F |
| `FilterDropdown`, `SortDialog`, `ConditionEditor`, `GridSemantics`, `GridHeaders`, `SheetNavigator`, `ViewerStatus` | E → F |
| `sheetTiles.ts`, `renderViewport.ts`, `overlay.ts`, `selectionOverlay.ts`, `xlsx/view/tileCache.ts` | G only (`xlsx/view/contracts.ts` and `cacheIdentity.ts` are not edited) |
| `src/components/scrollPan.ts`, `pointer.ts`, `fileName.ts`, `paged/metrics.ts` | C |
| `src/styles/officeview.css` | A creates; B, C, E and F append only within their own `/* stream X */` section |
| `package.json` | A (tailwind, postcss), B (`lucide-react`), F (`motion`) |

---

### Task 0: Budget baseline, eager-graph check, lazy boundary for paged floating controls

Spec §6.7 R5. This is a pure move with no visual change. It frees eager bytes before anything is added.

**Files:**
- Create: `scripts/eager-budget.ts`, `scripts/vite-module-map.ts`, `scripts/check-eager.ts`, `tests/helpers/eager-graph.ts`, `tests/ui-eager-graph.test.ts`, `src/components/paged/PagedChrome.tsx`, `tests/paged-floating-lazy.test.tsx`.
- Modify `vite.config.ts`: add `moduleMap()` when `process.env.OV_MODULE_MAP`.
- Modify `src/components/OfficeDoc.tsx`: replace the inline floating `SearchBar`/`ZoomControls`/copy pill JSX with a lazy `FloatingControls`, and keep `viewerControls` to `import type` only.
- Modify the tests that expected synchronous floating controls (`tests/officeview-*.test.tsx`, `tests/viewer-chrome.test.tsx`) so they use `findBy…`.

**Interfaces:**
- `scripts/vite-module-map.ts`: `export function moduleMap(): Plugin`. In `generateBundle` it writes `dist/.module-map.json` as `Record<chunkFile, { modules: string[]; imports: string[]; dynamicImports: string[] }>`, using Rollup's `chunk.modules` keys, `chunk.imports` and `chunk.dynamicImports`.
- `tests/helpers/eager-graph.ts`:
  - `export type ModuleMap = Record<string, { modules: string[]; imports: string[]; dynamicImports: string[] }>`
  - `export function eagerModules(map: ModuleMap, entry = 'index.js'): Set<string>`, the closure over `imports` only (never `dynamicImports`).
  - `export function forbiddenEager(map: ModuleMap, rules: readonly RegExp[]): string[]`
  - `export const EAGER_FORBIDDEN: readonly RegExp[]`, matching the list in Global Constraints.
  - `scripts/check-eager.ts` imports these.
- `src/components/paged/PagedChrome.tsx`:
  - `export interface FloatingControlsProps { search: ComponentProps<typeof SearchBar> | null; zoom: ComponentProps<typeof ZoomControls> | null; copy: { state: 'idle' | 'done' | 'failed'; onCopy(): void } | null; focusSearchRequest: number }`
  - `copy` is non-null only when `allowCopy` is true and a non-empty selection exists, the same condition as `OfficeDoc.tsx:814` today. `onCopy` is the eager `doCopy`, so a click on the pill copies exactly as it does today (R2-8).
  - `export function FloatingControls(props: FloatingControlsProps): ReactElement`, which renders today's JSX verbatim.
- `OfficeDoc.tsx`: `const FloatingControls = lazy(() => import('./paged/PagedChrome').then(m => ({ default: m.FloatingControls })))`.
  - The copy, Cmd/Ctrl-F and Cmd/Ctrl-P key handlers stay eager.
  - A Cmd/Ctrl-F press before the chunk loads increments `focusSearchRequest`, which the chunk honours on mount.

- [ ] **Step 1: Baseline.** On the untouched branch, before creating any script, run `bun run build && node -e "console.log(require('zlib').gzipSync(require('fs').readFileSync('dist/index.js')).length)"` and record **B0**. Treat the spec's ~23 B headroom as stale until this measurement. Then create `scripts/eager-budget.ts`, which runs the same measurement and exits 1 above 38,720, and confirm it prints B0. No product code changes yet.
- [ ] **Step 2: Failing tests.**
  - `tests/ui-eager-graph.test.ts`:
    - `it('eagerModules follows static imports only')`: synthetic map; `index.js` statically imports `a.js` and dynamically imports `b.js`; the result contains `a.js`'s modules and not `b.js`'s.
    - `it('flags a deliberate eager motion import (mutation)')`: add `/node_modules/motion/dist/es/index.mjs` to `a.js`'s modules; `forbiddenEager` returns it; the same module under `b.js` returns `[]`.
  - `tests/paged-floating-lazy.test.tsx`:
    - `it('floating search and zoom controls appear after the lazy chunk resolves')`: `await findByRole('searchbox')`, and `findByLabelText('Zoom in')`.
    - `it('Cmd/Ctrl-F before the chunk loads focuses the search input on mount')`: hold the dynamic import with a deferred mock, press Ctrl-F, resolve, and expect `document.activeElement` to be the searchbox.
    - `it('copy confirmation still announces Copied')`.
    - `it('clicking the lazy copy pill with a selection copies it')`: spy on `navigator.clipboard.writeText`, make a selection, wait for the pill and click it; the spy is called once with the selected text and `data-state="copied"` follows.
    - `it('allowCopy=false renders no copy pill even with a selection')`.
- [ ] **Step 3:** Run `bunx vitest run tests/ui-eager-graph.test.ts tests/paged-floating-lazy.test.tsx`. Expect FAIL.
- [ ] **Step 4:** Implement. Run the two files and then all of `bunx vitest run`. Expect PASS.
- [ ] **Step 5:** Run `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts`. Expect PASS, with `viewerControls.tsx` absent from the eager set. Record **B1** and the delta (B1 < B0 expected). Run `bunx playwright test --config playwright.viewer.config.ts viewer-layout-shift`; CLS must be ≤ 0.001.
- [ ] **Step 6:** `bunx tsc --noEmit` is green. Commit `perf(ui): lazy paged floating controls, eager-graph check and budget baseline` with B0 and B1 in the body.

### Task 1: Foundation: Tailwind pipeline, scoping, tokens, theme root and prop, stylesheet export

**Files:**
- Create `src/styles/officeview.css`:
  ```css
  @import "tailwindcss/theme.css" prefix(ov) reference;
  @import "tailwindcss/utilities.css" prefix(ov);
  @custom-variant coarse (@media (pointer: coarse));
  @theme inline reference { --*: initial; --spacing: 4px; --color-<each COLOR_TOKEN>: var(--ov-color-<name>); --radius-control: var(--ov-radius-control); --radius-panel: …; --shadow-rest: …; --shadow-popover: …; --text-xs: 11px; --text-sm: 12px; --text-base: 13px; --ease-fluent: cubic-bezier(.1,.9,.2,1); }
  ```
  Then the token blocks and the scoped base rules (see Interfaces). This shape was verified with a scratch tailwindcss 4.3.3 compile: `reference` suppresses the `:root, :host` block, utilities read `var(--ov-…)` directly, spacing comes out as `calc(4px * n)`, and `coarse` emits `@media (pointer: coarse)`.
- Create: `scripts/postcss-ov-scope.ts`, `scripts/ov-css-config.ts`, `scripts/audit-css.ts`, `src/styles/tokens.ts`, `src/components/theme.ts`, `tests/setup/styles.ts`.
- Modify:
  - `vite.config.ts`: spread `ovCss` (adds `tailwindcss()` and `css.postcss.plugins: [ovCssScope()]`) and set `build.lib.cssFileName: 'style'`;
  - `vitest.config.ts`: spread the same `ovCss` explicitly, since vitest has its own config, and add `test.setupFiles: ['tests/setup/styles.ts']`;
  - `vite.harness.config.ts`: inherits `ovCss` through `mergeConfig(baseConfig, …)`; add **no** second copy of the plugins. Only allow `src/styles` in `server.fs`;
  - `package.json`: devDeps `tailwindcss`, `@tailwindcss/vite` and `postcss`; `exports["./style.css"] = "./dist/style.css"`; `sideEffects: ["./dist/worker.js", "*.css"]`; version `0.5.0`;
  - `src/index.ts`: `import './styles/officeview.css'`;
  - `src/demo/main.tsx`, `src/debug/viewer-e2e.tsx` and `src/debug/viewer-perf.tsx`: stylesheet import;
  - `src/components/OfficeDoc.tsx`: the `theme` prop; `themeRootProps` on the `LegacyOfficeDoc` wrapper and the `Suspense` fallback; pass `theme` to the sheet;
  - `src/components/xlsx/InteractiveSheetDoc.tsx`: `themeRootProps` on its outer wrapper; drop the inline `fontFamily`/`fontSize`/`color` there in favour of the scoped base rules;
  - `src/components/OfficeFile.tsx`: pass `theme` through.
- Test: `tests/ui-theme-foundation.test.tsx`, `tests/ui-css-scope.test.ts`, `tests/browser/ui-host-isolation.spec.ts`.

**Interfaces:**
- `src/components/theme.ts` (eager, small):
  - `export type ThemeMode = 'light' | 'dark' | 'system'`
  - `export function themeRootProps(mode: ThemeMode | undefined, instance: string): { 'data-ov-theme-root': ''; 'data-ov-theme': ThemeMode; 'data-ov-instance': string }`
  - `export function warnIfStylesMissing(themeRoot: Element): void`. It runs in an effect, with no `import.meta.env` gate. It checks `getComputedStyle(themeRoot).getPropertyValue('--ov-sentinel').trim() !== '1'` and, once per page (module flag), calls `console.warn('[officeview] style.css not loaded: import "officeview/style.css"')`.
  - `export function resetStyleWarningForTests(): void`
- `src/styles/tokens.ts` (not eager):
  - `export const COLOR_TOKENS`, `export type ColorToken`
  - `export const CANVAS_TOKENS: readonly CanvasToken[]`
  - `export const THEME_DEFAULTS: { light: Record<ColorToken, string>; dark: Record<ColorToken, string> }`
- `scripts/postcss-ov-scope.ts`:
  - `export function ovCssScope(): import('postcss').Plugin`. This is a PostCSS plugin (`postcssPlugin: 'ov-scope'`, `OnceExit`). It acts only when `root.source?.input.file` ends with `src/styles/officeview.css` and leaves any other CSS untouched. It runs through Vite's `css.postcss.plugins`, which is after `@tailwindcss/vite` has generated the CSS and before Vite consumes it, in the build, the dev server, `?inline` and vitest alike.
  - `export function scopeOvCss(css: string): string` is the same transform applied to a string, used by the unit tests. It does three things:
    - drops `@property` rules;
    - unwraps the `@layer properties`/`@supports` fallback into an unlayered rule `[data-ov-theme-root], [data-ov-theme-root] *, [data-ov-theme-root] ::before, [data-ov-theme-root] ::after { --tw-*: … }`, placed before the utilities;
    - appends `:is([data-ov-theme-root], [data-ov-theme-root] *)` to the leading class compound of each `.ov\:` utility selector, before any pseudo-element.
  - `export function unscopedSelectors(css: string): string[]`. The plugin throws if this returns anything non-empty. `scripts/audit-css.ts <file>` exits 1 on any unscoped selector or any `@property`, `:root` or `:host`.
- CSS blocks:
  - `:where([data-ov-theme-root]) { --ov-sentinel: 1; …light… }`
  - `:where([data-ov-theme-root][data-ov-theme=dark]) { …dark… }`
  - `@media (prefers-color-scheme: dark) { :where([data-ov-theme-root][data-ov-theme=system]) { …dark… } }`
  - Scoped base, unlayered and placed before the utilities:
    ```css
    [data-ov-theme-root] { font-family: system-ui, -apple-system, "Segoe UI", sans-serif; font-size: 13px; color: var(--ov-color-text) }
    [data-ov-theme-root] :where(*:not([data-officeview-root])) { box-sizing: border-box }
    [data-ov-theme-root] :is(button, input, select, textarea) { font: inherit; color: inherit; line-height: inherit; margin: 0; appearance: none; background: transparent; border: 0; padding: 0; cursor: pointer }
    ```
    `input` and `select` get `cursor: auto`. Using `:is()` here gives specificity (0,1,1), up from (0,1,0) with `:where()`.
  - **Control utility rule (opencode R2 MINOR 2).** A host rule with specificity (0,1,1), such as `.app button`, ties with the base, so the base cannot be relied on to beat it. Every viewer `button`, `input`, `select` and `textarea` therefore carries explicit `ov:` utilities for colour, background, padding and border. Those utilities are at (0,2,0) and beat (0,1,1) host rules. Tasks 2, 3 and 5 enforce this with `it('every viewer control declares colour, background, padding and border utilities')`, a jsdom class check over every such element inside `[data-ov-theme-root]`, and with the hostile-host Playwright leg.
- `OfficeDocProps.theme?: ThemeMode`.

- [ ] **Step 1: Failing tests.**
  - `tests/ui-css-scope.test.ts` (pure):
    - `it('drops @property and rescopes the properties fallback')`.
    - `it('appends the theme-root scope to utility selectors, variants and pseudo-elements')`: a table of real Tailwind output, including `.ov\:flex`, `.ov\:hover\:bg-hover:hover` inside `@media (hover: hover)`, `.ov\:placeholder\:text-muted::placeholder`, `.ov\:coarse\:min-h-11` inside `@media (pointer: coarse)` and `.ov\:motion-reduce\:transition-none`.
    - `it('unscopedSelectors reports any rule without [data-ov-theme-root] (mutation)')`: `.x{}` and `:root{}` are reported.
    - `it('the ?inline compiled stylesheet is scoped (pipeline placement)')`: `import css from '../src/styles/officeview.css?inline'` goes through the real vitest pipeline. `unscopedSelectors(css)` returns `[]`, `css` contains no `@property`, and it contains at least one `.ov\:` selector ending in `:is([data-ov-theme-root], [data-ov-theme-root] *)`. This fails if the plugin is registered at a stage that never sees the CSS.
  - `tests/ui-theme-foundation.test.tsx`:
    - `it('marks the outer wrapper as theme root, data-ov-theme=system by default, honouring the prop')`: for a DOCX, `[data-ov-theme-root]` contains `[data-officeview-root]` and, after the lazy chunk loads, the floating search box; rerender with `theme="dark"` gives `'dark'`. Do the same for the interactive sheet wrapper and the `Suspense` fallback.
    - `it('data-officeview-root keeps host className/style and carries no ov: classes')`.
    - `it('defines light tokens and the sentinel on the theme root')`: `--ov-color-accent` is `#0f6cbd` and `--ov-sentinel` is `1`.
    - `it('dark tokens apply for data-ov-theme=dark')`: `--ov-color-surface` is `#292827`.
    - `it('THEME_DEFAULTS mirror the stylesheet token blocks')`: parse `officeview.css?inline`; every `COLOR_TOKENS` entry in the light and dark blocks equals `THEME_DEFAULTS`.
    - `it('warns once per page when the sentinel is missing, regardless of build mode')`: remove the setup style and mount two viewers → 1 warn with the exact message; with the stylesheet present → 0.
- [ ] **Step 2:** Run `bunx vitest run tests/ui-css-scope.test.ts tests/ui-theme-foundation.test.tsx`. Expect FAIL.
- [ ] **Step 3:** Run `bun add -d tailwindcss @tailwindcss/vite postcss`, then implement the files and wiring.
- [ ] **Step 4:** Run the two files. Expect PASS. Then run all of `bunx vitest run`: no `[officeview] style.css` warnings in the output.
- [ ] **Step 5: Host isolation (Review Focus 1, 2).** Create `tests/browser/ui-host-isolation.spec.ts` on the harness with the sheet viewer and the paged viewer. The hostile sheet is `button{all:unset} *{box-sizing:content-box} .flex{display:block} html{font-size:10px} .app button{padding:20px;color:rgb(1,2,3)}`, and both viewers are mounted inside `<div class="app">` so that the `.app button` rule actually reaches viewer buttons.
  - `test('toolbar geometry and colours are identical with and without a hostile host reset (host sheet first)')` and `(… host sheet last)`: for every toolbar and floating control, the bounding box (±0.5px) and the computed `color`, `background-color`, `border-color`, `padding` and `box-sizing` are deep-equal to the clean page.
  - `test('host elements are untouched')`: a host button inside `.app` (outside the viewer) keeps `rgb(1, 2, 3)`; a host `.flex` element stays `display: block`.
  - `test('dist/style.css: every selector is scoped; no @property, :root or :host')`: read and parse with `postcss`, then use `unscopedSelectors` → `[]`.
  - `test('the harness-served stylesheet is scoped (dev pipeline)')`: in the page, gather the `cssText` of every `document.styleSheets` rule that comes from `officeview.css`; `unscopedSelectors` returns `[]` and there is no `@property`.
  - `test('two adjacent viewers resolve their own theme and host overrides, including floating controls')`: viewer 1 is light with `.brand-a [data-ov-theme-root]{--ov-color-accent:#c50f1f}`; viewer 2 is dark. Assert each one's chrome computed colours.
  - Run `bunx playwright test --config playwright.viewer.config.ts ui-host-isolation`. Expect PASS on Chromium and WebKit.
- [ ] **Step 6: Build + budget.** Run `bun run build`.
  - `dist/style.css` exists, and `bun scripts/audit-css.ts dist/style.css` exits 0. As a mutation check, temporarily remove `ovCssScope()` from `ovCss`, rebuild, confirm the audit exits 1 and the `?inline` test fails, then restore the plugin.
  - `grep -c "style.css not loaded" dist/index.js` is ≥ 1, which proves the warning check survives the library build.
  - Budget ≤ 38,720, and `check-eager` passes. Record before/after.
- [ ] **Step 7:** Full `bunx vitest run` and `bunx tsc --noEmit` are green. Commit `feat(ui): scoped Tailwind foundation, theme root, tokens and stylesheet export`.

### Task 2: Shared controls: Lucide icons and classes for SearchBar / ZoomControls / copy pill

`viewerControls.tsx` and the copy pill are lazy-only since Task 0. Task 2 must be **net ≤ 0 eager bytes**.

**Files:**
- Create: `src/components/Icon.tsx`, `tests/browser/ui-chrome-geometry.spec.ts`.
- Modify:
  - `src/components/viewerControls.tsx`: classes and icons; the private `T()` survives only for inline geometry that is truly computed;
  - `src/components/paged/PagedChrome.tsx`: copy pill to classes with `Copy`/`Check`/`Ban` icons and `data-state`;
  - `package.json`: `lucide-react`;
  - `officeview.css`: stream B section.
- Test: update `tests/viewer-chrome.test.tsx`, `tests/xlsx-viewer-chrome-glyphs.test.tsx` and `tests/officeview-selection.test.tsx` (copy confirmation test near :74–78).

**Interfaces:**
- `export type IconName = 'chevron-up' | 'chevron-down' | 'plus' | 'minus' | 'search' | 'copy' | 'check' | 'ban' | …`. Tasks 3 and 5 extend the union.
- `export function Icon(props: { icon: LucideIcon; name: IconName; size?: 14 | 16 | 20; className?: string }): ReactElement`. It renders the Lucide component with `aria-hidden="true"`, `focusable="false"`, `strokeWidth={2}` and **`data-icon={name}`**. This is the single icon contract; tests select `svg[data-icon="…"]`.
- Control classes: `ov:bg-surface ov:border ov:border-border ov:rounded-control ov:shadow-popover`. Targets use `ov:coarse:min-h-11 ov:coarse:min-w-11`. Transitions use `ov:transition-colors ov:duration-150 ov:motion-reduce:transition-none` (canonical v4 order, variant after the prefix).

- [ ] **Step 1: Update tests first** (R4 mapping; names quoted from the current files).
  - `viewer-chrome`:
    - `'coarse pointer: every control is at least 44px high and wide (the input is 148px wide)'` → jsdom `it('every labelled control carries the coarse 44px target classes; input keeps its 148px width class')`, plus Playwright `coarse pointer: every control is at least 44px high and wide (browser)`.
    - `'fine pointer: desktop sizes are unchanged'` → Playwright `fine pointer: desktop sizes are unchanged (browser)`. Each control's box must equal the fine-pointer sizes recorded from `main` into `tests/browser/fixtures/chrome-sizes.main.json` (±0.5px). This task restyles colours, radii and icons, not desktop sizes.
    - `'is bounded by the viewer width, lets the input shrink and wraps status/buttons'` → jsdom class assertions (`ov:max-w-[calc(100%-24px)] ov:box-border ov:flex-wrap`, input `ov:min-w-0`), plus Playwright `is bounded by the viewer width, lets the input shrink and wraps status/buttons (browser)`: in a 320px viewer the bar's right edge is ≤ viewer right − 12, there is no horizontal overflow, and the input is < 148px.
  - `chrome-glyphs`: search, zoom and copy controls contain `svg[data-icon][aria-hidden="true"]` and **no** text in `↑↓+−✓⌘`. The U+FE0E rule stays for any remaining glyph text.
  - `officeview-selection`: the copied state is `[data-state="copied"]` with an aria-live polite announcement of "Copied". The blocked state keeps `data-state="blocked"` and the accessible text `'Copy blocked — press ⌘C'`.
  - New in `viewer-chrome`: `it('every viewer control declares colour, background, padding and border utilities')`, the control utility rule from Task 1.
  - New in `ui-chrome-geometry.spec.ts`: `test('hostile host sheet (.app button, all:unset) leaves SearchBar/ZoomControls/copy pill geometry and colours unchanged (browser)')`. It reuses the Task 1 hostile sheet with the viewer inside `.app`, because the class swap lands in this task.
- [ ] **Step 2:** Run those three files. Expect FAIL.
- [ ] **Step 3:** Run `bun run build && bun scripts/eager-budget.ts` and record **Bb** (before). Implement with `bun add lucide-react`, keeping every aria-label.
- [ ] **Step 4:** Run the three files plus `tests/officeview-*.test.tsx` and `tests/xlsx-viewer-zoom-dock.test.tsx`. Expect PASS. Run `bunx playwright test --config playwright.viewer.config.ts ui-chrome-geometry`. Expect PASS.
- [ ] **Step 5: Budget.** Run `bun run build`, record **Ba**, and require **Ba ≤ Bb**. Run `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts`; `lucide-react` and `Icon.tsx` must not be eager.
  - **No alternate gate (R2-7).** Task 2 depends on Task 0 having merged. If Task 0 is rejected or reverted, Task 2 does not start, and Stream A is re-planned with the owner. There is no glyph fallback that would have to pass the eager check and fail it at the same time.
- [ ] **Step 6:** Full vitest and tsc are green. Commit `feat(ui): Lucide icons and token classes for shared viewer controls` with Bb and Ba.

### Task 3: Paged viewer toolbar, status line, fit-to-width layout mode, page tracking, word count, file name

**Files:**
- Create (eager): `src/components/paged/metrics.ts`, `src/components/pointer.ts`, `src/components/scrollPan.ts`.
- Create (lazy, inside the `PagedChrome` chunk): `src/components/paged/PagedToolbar.tsx`, `PagedStatus.tsx`, `MoreMenu.tsx`, `usePageTracking.ts`, `pageTracking.ts`, `wordCount.ts`, `useWordCount.ts`.
- Create (eager, small): `src/components/fileName.ts`, used by `OfficeFile`. It is deliberately outside `paged/` so the eager allowlist stays `paged/metrics.ts` only (R2-7). Record its eager gzip cost.
- Modify `src/components/OfficeDoc.tsx` (`LegacyOfficeDoc`):
  - placeholders from metrics; lazy `PagedToolbar` above the scroll container and `PagedStatus` below it;
  - remove `FloatingControls`;
  - `fitMode`, `title`, `showToolbar`;
  - name `PAGE_PADDING_PX = 16` (today's `padding = 16`).
- Modify:
  - `src/components/paged/PagedChrome.tsx`: export `PagedToolbar` and `PagedStatus`; delete `FloatingControls`;
  - `src/components/viewerControls.tsx`: `SearchBar` gets `variant?: 'floating' | 'inline'`, default `'floating'`;
  - `src/components/OfficeFile.tsx`: derive the name and pass `title` unless the host set one;
  - `src/core/stream.ts`: optional `onHeaders?: (headers: Headers) => void` on the fetch path, called once per response;
  - `officeview.css`: stream C section.
- Test: `tests/paged-toolbar.test.tsx`, `tests/paged-fit.test.tsx`, `tests/paged-tracking.test.ts`, `tests/paged-word-count.test.ts`, `tests/paged-file-name.test.ts`, `tests/paged-file-name-loader.test.tsx`, `tests/browser/paged-toolbar.spec.ts`.

**Interfaces:**
- `metrics.ts`:
  - `export const PAGED_TOOLBAR_HEIGHT = { regular: 40, compact: 48 } as const`
  - `export const PAGED_STATUS_HEIGHT = 24`
  - `export const COMPACT_PANE_PX = 600`
  - `export function isCompact(containerWidth: number, coarse: boolean): boolean` (`coarse || containerWidth < COMPACT_PANE_PX`)
  - The placeholders and the real toolbar/status both set `style={{ height }}` from these constants, so their heights match by construction.
- `pointer.ts`: `export function useCoarsePointer(): boolean`, which subscribes to `matchMedia('(pointer: coarse)')` `change`.
- `fileName.ts`: `export function fileNameFrom(source: unknown, disposition?: string | null): string | undefined`.
  - Accepts `File.name`; or the `Request`/`Response`/`HttpSource` URL's last path segment, with query and hash stripped and `decodeURIComponent` applied (a malformed escape keeps the raw segment).
  - RFC 6266: `filename*` wins over `filename`, and `filename` wins over the URL. A malformed `filename*` falls back to `filename`.
  - Returns `undefined` otherwise. The name is set from the resolved response for promised/one-shot sources and is discarded if the source changed meanwhile.
- `pageTracking.ts` (pure):
  - `export function pageOffsets(specs: readonly PageSpec[], displayWidth: number, gap: number): { top: number[]; height: number[] }`
  - `export function mostVisiblePage(o: { top: number[]; height: number[] }, viewTop: number, viewHeight: number): number`. It picks the largest visible pixel height, ties going to the lower index.
- `usePageTracking.ts`: `export function usePageTracking(container: RefObject<HTMLElement | null>, layout: { specs: readonly PageSpec[]; displayWidth: number; gap: number }, transform: Transform, setTransform: (t: Transform) => void): { current: number; goTo(index: number): void }`.
  - Recomputes once per rAF on scroll, transform, resize or layout change, and calls `setState` only when the index changes.
  - It maps the viewport into content coordinates **only** through `scrollPan.contentYAt`, which includes the container's native `scrollTop`. Visible content runs from `contentYAt(0, s)` to `contentYAt(clientHeight, s)`.
  - `goTo(i)`, not zoomed: `container.scrollTo({ top: PAGE_PADDING_PX + top[i] })`.
  - `goTo(i)`, zoomed: `setTransform(clampPan({ ...t, panY: PAGE_PADDING_PX - top[i] * zoom }, viewport))`. This relies on the invariant that `scrollTop === 0` whenever `zoom > 1`; it asserts the invariant in development and otherwise subtracts `scrollTop`.
  - Never `scrollIntoView`.
- `scrollPan.ts` (eager, pure; R2-5). Native scroll and pan move between modes atomically, so the retained `scrollTop` of an `overflow: hidden` container can never skew tracking or `goTo`. `zoom.ts` is unchanged.
  - `export interface ScrollPanState { scrollTop: number; panY: number; zoom: number; padding: number }`
  - `export function contentYAt(viewY: number, s: ScrollPanState): number`, which returns `(viewY + s.scrollTop - s.padding - s.panY) / s.zoom` in natural-layout content px.
  - `export function enterZoom(s: ScrollPanState, nextZoom: number, anchorY: number): { panY: number; scrollTop: 0 }`. It keeps `contentYAt(anchorY)` fixed. `LegacyOfficeDoc` applies `scrollTop = 0` in a layout effect in the same commit that switches `overflow` to `hidden`.
  - `export function exitZoom(s: ScrollPanState, anchorY: number, maxScrollTop: number): { panY: 0; scrollTop: number }`. It keeps `contentYAt(anchorY)` fixed, clamped to `[0, maxScrollTop]`, and is applied in the same commit that switches back to `overflow: auto`.
  - Every zoom input goes through these helpers: buttons, Ctrl/Cmd-wheel, pinch and keyboard.
- `wordCount.ts` (R2-3: bounded per unit, never concatenates, words stay whole across chunks):
  - Policy, modelled on Word's default count. The expected counts in the tests are hand-derived from this policy, not measured in Word.
    - A **word** is a maximal run of non-whitespace, non-CJK code points that contains at least one letter or digit (`\p{L}`/`\p{N}`). A lone `—` or `•` does not count; `$50`, `100%`, `e-mail`, `3.14` and `https://x.com/a-b?c=1` each count once.
    - Each Han, Hiragana or Katakana code point (plus U+30FC `ー`) counts as **one word** and also ends any open word. A combining mark (`\p{M}`) continues the current token and never starts a word.
    - Paragraph boundaries act as whitespace; run boundaries do not.
  - `export const UNIT_CHARS = 2000`
  - `export interface WordCounter { push(chunk: string): void; boundary(): void; total(): number }`
  - `export function createWordCounter(): WordCounter`. This is a resumable state machine. Its state is `inToken`, `tokenHasAlnum` and `pendingHigh` (a lone high surrogate carried over to the next `push`). It iterates by code point, so a token, surrogate pair or mark that spans any number of `push` calls is counted exactly once.
  - `export function countWords(text: string): number`, a convenience wrapper: `createWordCounter()`, then `push(text)`, then `total()`.
  - `export function* wordSourceChunks(doc: DocxDocument): Generator<string | typeof PARA>`.
    - It walks the model **lazily**: body blocks, then table cells (recursive), text-box paragraphs, footnotes and endnotes. Headers, footers and comments are excluded.
    - It yields each run's text in slices of at most `UNIT_CHARS` code units, never splitting a surrogate pair, and yields `PARA` after each paragraph.
    - It never joins runs and never collects arrays up front, so the work between two yields is bounded by `UNIT_CHARS` plus one node visit.
  - `export function countDocumentWords(doc: DocxDocument, opts: { signal: AbortSignal; now?: () => number; schedule?: (cb: () => void) => () => void; sliceMs?: number }): Promise<number>`.
    - Defaults: `now = performance.now`, `schedule = requestIdleCallback ?? setTimeout(0)`, `sliceMs = 8`.
    - Within a slice, it pulls chunks from the generator and pushes them into one counter, checking `signal.aborted` and `now()` after **every** chunk, and yields when the slice reaches `sliceMs`.
    - Abort rejects with an `AbortError`, and nothing is published for a superseded generation.
- `useWordCount.ts`: `export function useWordCount(document: OfficeDocSource): number | null`. Returns `null` for PPTX and static XLSX, and while counting. A fresh `AbortController` is created per document, and the previous one is aborted.
- `PagedToolbar` props:
  ```ts
  { title?: string; kind: 'document' | 'slides' | 'sheet-pages'; current: number; total: number;
    onPrev(): void; onNext(): void; zoom: number; onZoomIn(): void; onZoomOut(): void; onFit(): void;
    fitActive: boolean; search?: ReactNode; onCopy?: () => void; copyState: 'idle' | 'done' | 'failed';
    onPrint?: () => void; compact: boolean; showSearch: boolean; showZoom: boolean;
    focusSearchRequest: number; onEscapeToDocument(): void }
  ```
  - Copy and Print are omitted when their callbacks are undefined.
  - The title truncates with an ellipsis and has `title=` for the full name.
- `PagedStatus` props: `{ kind; current: number; total: number; words: number | null }`.
  - Document: `Page 3 of 12 · 2,418 words`, with `— words` while `null` (fixed min-width).
  - Slides: `Slide 3 of 12`. Sheet pages: `Page 3 of 12`.
- Fit-to-width (spec §6.9):
  - `avail = Math.max(0, containerWidth - 2 * PAGE_PADDING_PX)`. This keeps today's guard at `OfficeDoc.tsx:332` (R2-2).
  - `displayWidth = fitMode && avail > 0 ? Math.min(avail, naturalMax * MAX_ZOOM) : Math.min(avail || naturalMax, naturalMax)`. For container widths 0 to 32, `avail` is 0 and the natural width is used, exactly as today; it is never negative.
  - The transform stays at zoom 1 while fit mode is on.
  - **Fit exit, one rule.** A zoom input in fit mode is applied as if the viewer were on the natural layout at `z0 = fitDisplayWidth / naturalDisplayWidth`, with the same content point under that input's existing anchor. Fit mode turns off in the same commit.
    - Buttons give `stepZoom(z0, ±1)`; wheel and pinch apply their ratio to `z0`.
    - If the result is above 1, it goes through `enterZoom`. If it clamps to 1, the viewer lands on the native natural layout through `exitZoom`'s scroll mapping.
    - Example: `z0 = 968/600 ≈ 1.613`. Zoom in gives 2.42 (zoomed). Zoom out gives `stepZoom(1.613, -1)` = 1.0756 (still zoomed, not natural), and a second zoom out gives 1, the native natural layout.
    - The readout shows `round(zoom × 100)%` and shows `round(z0 × 100)%` while fit mode is on.
    - The spec's earlier wording ("zooming out from fit lands at natural layout, zoom 1") is replaced by this rule.
  - `fitActive = fitMode` (`aria-pressed`).
  - In fit mode, off-screen pages use the visibility path, with backing ≤ natural-mode size.
  - `zoom.ts` is unchanged.
- Kind: `'slides' in doc` → `'slides'`; `'sections' in doc` → `'document'`; otherwise `'sheet-pages'`.
- Compact mode (`isCompact`):
  - the file name is hidden;
  - Find collapses to a `Search` icon button that expands the inline search;
  - Fit, Copy and Print move into a `More` button (`Ellipsis`, `aria-haspopup="menu"`, `aria-expanded`) whose list has `role="menu"`;
  - all targets are 44px.
- Keyboard contract (spec §6.5 revision):
  - The menu opens focused on its first item. ArrowUp/ArrowDown/Home/End move between items. Escape and outside clicks close it and return focus to More.
  - Cmd/Ctrl-F expands Find and focuses it. In compact mode, Escape collapses Find and returns focus to the Find button; on desktop, Escape calls `onEscapeToDocument`.
  - Crossing the compact threshold keeps the query and active match, and moves focus to the counterpart control.

- [ ] **Step 1: Failing tests.**
  - `paged-file-name.test.ts`:
    - `File([], 'Q3.docx')` → `'Q3.docx'`;
    - `{url:'https://x/a/b/Plan%20v2.pptx?sig=1'}` → `'Plan v2.pptx'`;
    - `new Request('https://x/r.rtf#p')` → `'r.rtf'`;
    - `attachment; filename*=UTF-8''r%C3%A9sum%C3%A9.docx` → `'résumé.docx'`;
    - `filename*=UTF-8''%E0%A4; filename="x.docx"` → `'x.docx'`;
    - URL segment `bad%E0.docx` → `'bad%E0.docx'`;
    - `new Uint8Array()` → `undefined`.
  - `paged-file-name-loader.test.tsx`:
    - `it('OfficeFile passes the File name as title')`
    - `it('uses Content-Disposition when exposed, the URL when the header is absent or not exposed')`
    - `it('promised one-shot source: name from the resolved Response URL')`
    - `it('source replaced during fetch: the stale name is never shown')`
    - `it('host title wins')`
    - `it('no name hides the slot')`
  - `paged-tracking.test.ts`:
    - `it('two 2000px pages in an 800px viewport switch at the dominance crossover')`: page 0 visible 600 and page 1 visible 200 → 0; page 0 visible 200 and page 1 visible 600 → 1. Neither page crosses 0, 0.5 or 1 visibility in between.
    - `it('mixed page sizes compare visible pixels, not ratios')`: a fully visible 200px page loses to a 500px-visible tall page.
    - `it('maps the viewport through the zoom transform')`
    - `it('scrollPan: contentYAt includes native scrollTop; enterZoom and exitZoom keep the anchored content point fixed')` (pure). From `{ scrollTop: 1000, panY: 0, zoom: 1, padding: 16 }`, `enterZoom(…, 2, 400)` returns `scrollTop` 0 and a `panY` for which `contentYAt(400)` is unchanged (±0.5). `exitZoom` round-trips the same point.
    - `it('native scroll to a later page, then zoom: tracking and goTo stay correct')`, in `LegacyOfficeDoc` with a jsdom layout mock:
      - native-scroll to page 3, then press Zoom in: `container.scrollTop === 0` in the same commit, and the tracked page is still 3;
      - `goTo(5)` while zoomed brings page 5's top to the viewport top (±1px);
      - zoom back out to 1: native scrolling resumes with page 5 at the top.
    - `it('updates at most once per animation frame and only on change')`: rAF mock, 50 scroll events in one frame → 1 computation; render count rises only on index change.
    - `it('goTo scrolls the viewer container, never scrollIntoView')`: spy `Element.prototype.scrollIntoView` (0 calls); `container.scrollTop` equals the target offset.
    - `it('goTo while zoomed sets panY through clampPan')`
    - `it('a replacement document with the same page count resets to page 1')`
  - `paged-word-count.test.ts`:
    - `countWords` table:

      | Input | Count |
      |---|---|
      | `'Hello, world'` | 2 |
      | `"don't stop e-mail 3.14 2026"` | 5 |
      | `'$50 and 100%'` | 3 |
      | `'日本語のテキスト'` | 8 |
      | `'Excel表格'` | 3 |
      | `'한국어 텍스트'` | 2 |
      | `''` | 0 |

    - Additional table rows: `'— • —'` → 0, `'https://example.com/a-b/c.d?x=1 ok'` → 2, `'e\u0301te'` (combining mark) → 1, `'𠀀𠀁'` (astral Han) → 2.
    - `it('a word split across runs counts once')`: runs `['Hel', 'lo world']` → 2.
    - `it('a token crossing chunk, run and slice boundaries counts once')`: push `'ab'`, `'cd'` and `'ef gh'` across three slices → 2.
    - `it('surrogate pairs, CJK and combining marks at chunk edges')`: split `'x𠀀\u0301y'` at every code-unit offset into two `push` calls; the total always equals `countWords` of the unsplit string.
    - `it('counts body, tables, text boxes and notes; skips headers, footers and comments')`: hand-built `DocxDocument`.
    - `it('never runs more than 8 ms of work per slice')`: fake `now` advances 1 ms per chunk; 500 chunks → every slice ≤ 8 chunks; the total is correct.
    - `it('a giant single token (2,000,000 chars, no whitespace) is bounded and counts once')`: every yielded chunk is ≤ `UNIT_CHARS`, every slice is ≤ 8 chunks, and the total is 1.
    - `it('a long URL paragraph and a 100,000-char CJK paragraph are chunked and counted exactly')`: URL repeated without spaces → 1; CJK → 100,000.
    - `it('a paragraph with 100,000 runs is traversed lazily')`: the generator's first chunk arrives before the fixture's run-array iterator has advanced past a handful of runs (instrument with a Proxy).
    - `it('cancels on abort: no result after document change')`
    - `it('useWordCount returns null until counted, never counts during render, null for PPTX and static XLSX')`
  - `paged-fit.test.tsx`:
    - `it('fit-to-width lays pages out at the pane width without a transform')`: container 1000px, page natural 600 → canvas CSS width `968px`; the content transform has zoom 1 (no `scale(` above 1); container `overflow` is `auto`; `Fit to width` has `aria-pressed="true"`.
    - `it('a pointer drag in fit mode selects text, not pan')`
    - `it('zooming from fit mode starts at the equivalent zoom and turns fit off')`: Fit, then `Zoom in` → scale ≈ `stepZoom(968/600, 1)` (±0.01), `aria-pressed="false"`, and the content point under the anchor is unchanged (±1px, not merely the same page).
    - `it('zoom out from fit follows stepZoom: 1.613 → 1.0756 (zoomed), then 1 (native natural layout)')`: after the second step, canvas width is 600, `overflow` is `auto`, `scrollTop` keeps the anchored point, and the readout goes 161% → 108% → 100%. The same rule is checked once each for Ctrl-wheel and pinch.
    - `it('fit mode follows container resize; narrow widths keep the natural fallback')`: 1000→1400 gives 1368. Container widths 0, 1, 31 and 32 give 600 (never negative), in both fit and non-fit mode. 33 gives 1 in non-fit and 1 in fit (`avail = 1`), matching today's guard.
    - `it('fit mode caps at MAX_ZOOM × natural width')`: page 100px in a 1000px pane → 600.
    - `it('fit mode keeps off-screen page backing stores at natural-mode size')`: 30 pages with the visibility observer mocked.
  - `paged-toolbar.test.tsx`:
    - `it('renders title, page position and word count for a DOCX')`: 3-page DOCX fixture with `title="Report.docx"`. The toolbar shows `Report.docx` and `1 / 3`, and the status shows `Page 1 of 3 · — words`. After flushing the idle slices it shows `Page 1 of 3 · N words`, where N is the hand-counted fixture value stated in the test.
    - `it('uses Slide wording and no word count for PPTX')`
    - `it('hides Print and Copy when disallowed')`
    - `it('showToolbar=false renders no toolbar, status or placeholders; showSearch/showZoomControls hide their groups')`
    - `it('placeholders reserve regular and compact heights before the chunk resolves')`: with the import held, the top placeholder `style.height` is `40px`, or `48px` with a coarse stub or a 375px container, and the bottom is `24px`. After resolving, the toolbar and status roots have the same inline heights.
    - `it('compact mode collapses Find and moves Fit/Copy/Print into More')`
    - `it('More menu: opens focused on first item, arrows/Home/End move, Escape and outside click close and return focus')`
    - `it('Cmd/Ctrl-F expands compact Find and focuses it; Escape collapses and returns focus to the Find button')`
    - `it('desktop Find Escape returns focus to the document region')`
    - `it('crossing the compact threshold keeps query and active match and moves focus to the counterpart')`
    - `it('zoom readout follows Ctrl-wheel and pinch zoom synchronously')`
    - `it('next/prev move the current page and update the status')`
- [ ] **Step 2:** Run `bunx vitest run tests/paged-*.test.ts tests/paged-*.test.tsx`. Expect FAIL.
- [ ] **Step 3:** Implement. Keep the Ctrl-wheel/keyboard behaviour.
- [ ] **Step 4:** Run the new files, plus `tests/officeview-*.test.tsx` and `tests/viewer-chrome.test.tsx`. Expect PASS. Paged-viewer expectations move from the floating SearchBar to the inline toolbar; the sheet viewer keeps its placement.
- [ ] **Step 5: Browser (Review Focus 4, 5, 7).** `tests/browser/paged-toolbar.spec.ts` uses `viewer-e2e.html` with a DOCX and a PPTX, at 1280×800 and 375×667 coarse, on Chromium and WebKit.
  - `test('toolbar and status never overlap the visible part of any page')`: intersect each canvas rect with the scroll container rect first.
  - `test('coarse: every toolbar control is at least 44×44')`
  - `test('placeholder and real chrome heights match; CLS ≤ 0.001 with a delayed chunk')`: hold the `PagedChrome` chunk with `page.route` and release it after 1s. Cover desktop, compact, document load and a `showToolbar` toggle.
  - `test('keyboard: Tab order through toolbar groups with a visible focus ring in light and dark')`: outline width ≥ 2px and contrast ≥ 3:1.
  - `test('fit-to-width keeps native wheel scrolling and drag-to-select')`
  - `test('fit-to-width follows window resize')`
  - `test('two tall pages: status flips at the dominance crossover')`
  - `test('rapid scrolling a 100-page document: status updates at most once per frame')`
  - `test('native scroll to page 3, zoom in, goTo(5), zoom out: the tracked page and position stay correct')` (R2-5)
  - `test('hostile host sheet leaves paged toolbar and status geometry and colours unchanged')`, with the viewer inside `.app` (opencode R2 MINOR 2).
- [ ] **Step 6:** Build and budget. Before and after must stay ≤ 38,720, and `check-eager` passes (only `paged/metrics.ts`, `pointer.ts`, `scrollPan.ts` and `fileName.ts` are new eager code; record each one's cost).
- [ ] **Step 7:** Full vitest, tsc and the browser suite are green. Commit `feat(ui): paged toolbar and status (lazy), fit-to-width layout mode, page tracking, word count, file name`.

### Task 4: Paged document print (isolated iframe)

**Files:**
- Create: `src/components/paged/printDocument.ts` (its own lazy module).
- Modify `src/components/OfficeDoc.tsx`:
  - the Print handler and Cmd/Ctrl-P call `import('./paged/printDocument')`;
  - the `allowPrint={false}` guard is scoped to `[data-ov-instance="<id>"]` instead of every `[data-officeview-root]`;
  - cleanup on source change, permission withdrawal and unmount.
- Modify `src/components/paged/PagedToolbar.tsx`: wire `onPrint`.
- The sheet print path in `InteractiveSheetDoc.tsx` is **not** touched.
- Test: `tests/paged-print.test.tsx`; add to `tests/browser/paged-toolbar.spec.ts`.

**Interfaces:**
- Constants:
  - `PRINT_SCALE = 2` and `PRINT_MAX_AXIS = 16384`. These mirror `InteractiveSheetDoc.tsx:271–273`, and each copy carries a comment pointing at the other (opencode R2 NIT 4).
  - `PRINT_MAX_BYTES = 64 * 1024 * 1024`, the cap for the single raster workspace.
  - Whole-job budgets (R2-4): `PRINT_JOB_MAX_DECODED = 1024 * 1024 * 1024`, an estimate of Σ W·H·s²·4 over all pages at the chosen scale, and `PRINT_JOB_MAX_BLOB_BYTES = 192 * 1024 * 1024`, the actual sum of retained PNG blob sizes.
- `export type PrintPlan = { kind: 'ok'; scale: 1 | 2 } | { kind: 'refused'; reason: 'too-large' }`.
- `export function planPrintJob(specs: readonly PageSpec[], budgets = { decoded: PRINT_JOB_MAX_DECODED, workspace: PRINT_MAX_BYTES, maxAxis: PRINT_MAX_AXIS }): PrintPlan`.
  - It picks one uniform scale for the job: 2 if every page fits the workspace and axis caps at 2 and the decoded total fits, otherwise 1 under the same conditions, otherwise `refused`.
  - It runs **before** anything is allocated.
  - A refusal shows the status-line notice "This document is too large to print from the viewer." and never produces a partial printout.
- `export function pageRules(specs: readonly PageSpec[], namedPages: boolean): string`.
  - With `namedPages`, it emits one named `@page pN { size: Wpx Hpx; margin: 0 }` per distinct size, plus `img[data-page="pN"] { page: pN; display: block; width: Wpx; height: Hpx; break-after: page }`.
  - Without named pages (opencode R2 MINOR 3: `CSS.supports('page', 'p0')` is false, as in older WebKit), it emits a single `@page { size: <largest W>px <largest H>px; margin: 0 }`. Each image is fitted inside that box (`object-fit: contain`) with `break-after: page`. This fallback is documented.
- `export interface PrintJob { frame: HTMLIFrameElement; dispose(): void }`.
- `export async function printDocument(pages: readonly Paintable[], themeRoot: HTMLElement, opts: { signal: AbortSignal; title?: string; onNotice(text: string): void; rasterize?: (page: Paintable, scale: 1 | 2, canvas: HTMLCanvasElement) => Promise<Blob | null> }): Promise<PrintJob | null>`.
  - Calls `planPrintJob` first, then paints in the parent document (where the registered fonts are), one page at a time on one reused workspace canvas.
  - **Released per page.** Once a page's blob exists, the canvas is cleared before the next page. At the end of the job the canvas is zeroed (`width = height = 0`).
  - Each blob becomes `URL.createObjectURL` and an `<img>` in a hidden same-origin `<iframe aria-hidden tabindex=-1>` appended to `themeRoot`.
  - The running sum of blob sizes is checked against `PRINT_JOB_MAX_BLOB_BYTES` after each page. Going over aborts the job: it calls `dispose()`, shows the same notice and returns `null`.
  - It waits for each image's `load` event and does **not** call `img.decode()`, so the job never forces every bitmap to be decoded at once. The decoded memory the print pipeline may use is bounded by the `planPrintJob` estimate.
  - Then it calls `frame.contentWindow.print()`.
  - `dispose()` is idempotent: it removes the frame, revokes every URL and zeroes the canvas.
- **In-flight safety.** Every await point re-checks `signal.aborted` after it resumes: the lazy import, `rasterize`/`toBlob`, and image load.
  - A blob that arrives after disposal is never turned into a URL.
  - A URL created just before disposal is revoked immediately.
  - A failure (rasterize throws, an image `error` event) disposes everything, shows a notice and returns `null`. Nothing is left behind even when no `PrintJob` was ever returned.
- Cleanup: `dispose()` runs on the frame's or the window's `afterprint`, a new print, a source change, `allowPrint` turning false, unmount, and abort. It never depends on `afterprint` alone.

- [ ] **Step 1: Failing tests** in `tests/paged-print.test.tsx`:
  - `it('prints from a hidden same-origin iframe holding only page images')`: the host page has `<p>HOST-SECRET</p>`; 3-page DOCX; injected `rasterize`. The frame's body has exactly 3 `img` and `textContent` without `HOST-SECRET`. `frame.contentWindow.print` is called once and `window.print` 0 times.
  - `it('emits one named @page rule per distinct page size with zero margin')`: mixed page sizes.
  - `it('planPrintJob picks one uniform scale and refuses beyond the whole-job budget')`:
    - 20 Letter pages (816×1056) → scale 2;
    - 200 pages → scale 1 (scale 2 would be 2.76 GB);
    - 400 pages → refused (1.38 GB at scale 1 is over the 1 GiB budget);
    - one page with an axis over 16,384 at scale 2 forces scale 1 for the whole job.
  - `it('a refused job allocates nothing and shows the notice')`: `rasterize` is never called, no frame exists and `onNotice` is called once.
  - `it('exceeding the blob budget mid-job aborts, disposes and notifies')`: the injected `rasterize` returns 1 MiB blobs and the budget is set to 3 MiB; at page 4 the frame is removed, three URLs are revoked, `onNotice` is called and the result is `null`.
  - `it('the workspace canvas is cleared per page and zeroed at the end')`.
  - `it('abort during the lazy import creates no frame')`: hold the dynamic import, abort, then resolve.
  - `it('abort during rasterize/toBlob: a late blob never becomes a URL')`: gated `rasterize`; dispose; release the gate; `createObjectURL` is never called for that blob.
  - `it('abort during image load revokes everything and never calls print')`.
  - `it('a rasterize failure before PrintJob is returned leaves no frame or URL behind')`.
  - `it('renders pages one at a time on one reused canvas')`: `rasterize` concurrency is ≤ 1 and it always receives the same canvas.
  - `it('cleans up on afterprint, new print, source replacement, allowPrint=false and unmount')`: one sub-case each. The frame is removed and `URL.revokeObjectURL` is called 3 times.
  - `it('allowPrint=false: no Print button, Cmd/Ctrl-P prevented, guard hides only this instance')`: an allowed sibling viewer plus a sheet viewer stay printed (reuse the `printed()` helper pattern from `viewer-chrome.test.tsx`).
  - `it('Cmd/Ctrl-P when allowed prints the document frame, not the host page')`
- [ ] **Step 2:** Run → FAIL.
- [ ] **Step 3:** Implement.
- [ ] **Step 4:** Run the new file, all `tests/*print*` and `tests/viewer-chrome.test.tsx` → PASS.
- [ ] **Step 5:** Browser tests:
  - `test('paged print frame contains only page images at page size')`: `addInitScript` stubs `Window.prototype.print` in every frame. Find the frame; it has N `img` and no host text; a frame screenshot under `emulateMedia({ media: 'print' })` shows a white page.
  - `test('print cleanup without afterprint: next print or unmount removes the frame')`
  - `test('page box size: Chromium named pages give per-size PDF pages; WebKit uses named pages when supported, else the single-size fallback')`.
    - Chromium: `page.pdf()` of the frame for a mixed-size document has page boxes equal to each page's size (±1pt).
    - WebKit: assert which `pageRules` branch matches `CSS.supports('page', 'p0')` in that engine, and that every image's computed print box fits inside the declared `@page` size.
- [ ] **Step 6:** Build and budget (`printDocument.ts` must not be eager; `check-eager` covers it). Run the full suites, then commit `feat(ui): isolated paged-document print`.

### Task 5: Spreadsheet chrome restyle

**Files:**
- Modify `src/components/xlsx/InteractiveSheetDoc.tsx`:
  - toolbar actions become icon buttons with labels;
  - Print is hidden when `!allowPrint`;
  - the `data-officeview-sheet-font` runtime rule moves into the stylesheet, while per-instance print rules stay injected;
  - `DockedZoomControls` gets classes.
- Modify:
  - `SheetNavigator.tsx`;
  - `GridHeaders.tsx`: Lucide `ArrowUpDown`, `ArrowUp`, `ArrowDown`, `Loader2` (pending), `Filter`, `Ellipsis` and a corner icon. `HEADER_COLORS` becomes a token-name map that is still exported for tests;
  - `GridSemantics.tsx` (popup + status bar), `ViewerStatus.tsx`, `FilterDropdown.tsx`, `SortDialog.tsx`, `ConditionEditor.tsx`, `controlUi.ts`;
  - `sheetResize.tsx`: badges become `MoveVertical`/`MoveHorizontal` plus a number;
  - `officeview.css` stream E section.
- Must not edit `viewerControls.tsx`, `OfficeDoc.tsx` or any canvas painter.
- Test: update `xlsx-viewer-policy`, `-sort-ui`, `-touch-header`, `-events`, `-resize`, `xlsx-merge-ce-resize-panes`, `-sort-pending`, `-filter-ui`, `xlsx-fix-c-scroll-into-view`, `-a11y`, `-visual-review`, `-zoom-dock`, `-chrome-glyphs` and `-frozen-panes-paint`. Add to `tests/browser/ui-chrome-geometry.spec.ts` and `interactive-viewer-visual.spec.ts`.

**Interfaces:**
- Consumes Task 1 tokens and the `coarse` variant, and Task 2 `Icon`/`IconName` (extended with `arrow-up-down`, `arrow-up`, `arrow-down`, `loader`, `filter`, `ellipsis`, `move-vertical`, `move-horizontal`, `corner`).
- Public behaviour and aria-labels are unchanged.
- Geometry the canvas mirrors stays inline, and its jsdom assertions stay unchanged: row heights 20/36, the 88px touch header, the 44px actions row, `STATUS_BAR_HEIGHT = 24` and the computed widths. This covers all of `xlsx-viewer-touch-header`.
- Glyph assertions become icon assertions: `'↕'`, `'↖'`, `'…'` and `'↑'` become `svg[data-icon="arrow-up-down"|"corner"|"ellipsis"|"arrow-up"]`, with aria-labels unchanged. The touch badge text becomes `'2'` with `aria-label="Resize row 2"`.

- [ ] **Step 1:** Update the listed tests by the R4 rule. Every removed static colour or position assertion gets a `(browser)` twin in `ui-chrome-geometry.spec.ts` or `interactive-viewer-visual.spec.ts`; list the mapping in the commit. Then:
  - Add jsdom `it('header and status token colours meet contrast in light and dark')` to `xlsx-viewer-visual-review`. It computes from `THEME_DEFAULTS` (text ≥ 4.5, track ≥ 3).
  - Add the Playwright twin `header and status computed colours meet contrast in light and dark (browser)`.
  - Add `it('every viewer control declares colour, background, padding and border utilities')` for the sheet viewer's controls, popups and dialogs. Add the Playwright leg `hostile host sheet leaves sheet toolbar, header actions, filter popup and sort dialog geometry and colours unchanged (browser)`, with the viewer inside `.app`.
- [ ] **Step 2:** Run the listed files → FAIL.
- [ ] **Step 3:** Implement the Fluent look:
  - toolbar `ov:bg-chrome ov:border-b ov:border-border`;
  - controls `ov:rounded-control`;
  - menus and dialogs `ov:rounded-panel ov:shadow-popover ov:bg-surface`;
  - the mobile filter sheet `ov:rounded-t-[12px]`; it is position-fixed but stays inside the theme root, with no portal;
  - Print hidden when `!allowPrint`.
- [ ] **Step 4:** Run the listed files and `tests/xlsx-viewer-*.test.ts*` → PASS.
- [ ] **Step 5:** Browser: `interactive-viewer-visual`, `-a11y`, `-mobile`, `-resize`, `viewer-layout-shift` and `-dpr` pass on Chromium and WebKit. Add a dark-scheme run (`emulateMedia({ colorScheme: 'dark' })`) of the contrast, no-overlap and header-label checks.
- [ ] **Step 6:** Build, budget and `check-eager`. Commit `feat(ui): Fluent restyle of the spreadsheet viewer chrome`.

### Task 7: Canvas theming (paper rule, palette-keyed caches)

Runs after D and E (stream G).

**Files:**
- Create: `src/components/canvasColors.ts` (eager, small), `tests/ui-canvas-theme.test.tsx`.
- Modify:
  - `src/core/overlay.ts`: `paintGridSelectionOverlay(ctx, geometry, scale = 1, colors?: Pick<CanvasColors, 'selection' | 'selection-fill'>)` and `paintSearchOverlay(ctx, highlights, active, scale = 1, colors?: Pick<CanvasColors, 'find' | 'find-active'>)`;
  - `src/components/xlsx/selectionOverlay.ts`: `paintFindOverlay(ctx, highlights, active, colors?: Pick<CanvasColors, 'find' | 'find-active'>)`;
  - `src/xlsx/view/renderViewport.ts`: `ViewportPaintOptions.colors?: Pick<CanvasColors, 'pending' | 'pending-text'>`, replacing only `COARSE_FILL` and `#666666`. The white fill and `#d0d0d0` are unchanged;
  - `src/components/xlsx/sheetTiles.ts`: `SheetFrame.colors?: Pick<CanvasColors, 'pending' | 'pending-text' | 'freeze'>`, replacing `PENDING_FILL`, `#666666` and freeze `#5F6368`. The frame's `#ffffff` clear is unchanged;
  - Palette key (R2-6). It is renderer- and tile-specific; the shared `CacheIdentity` (`src/xlsx/view/contracts.ts:440`), which takes part in worker ownership checks in `src/xlsx/view/cacheIdentity.ts`, is **not** changed.
    - `src/xlsx/view/tileCache.ts`: `TileKey.paletteKey: string` (interface at :333), included in `tileKey()`'s serialized field array (:342–345). Every `TileKey` producer passes it.
    - `src/components/xlsx/sheetTiles.ts`: `SheetFrame.paletteKey: string`. Each `TileKey` built in the renderer gets `frame.paletteKey`. The `renderFrame` invalidation predicate (today `!sameIdentity(...) || key.zoom !== frame.zoom || key.dpr !== frame.dpr`) gains `|| key.paletteKey !== frame.paletteKey`. The local `sameIdentity` at :457 is unchanged.
    - Producer and consumer inventory, which the implementer confirms with `grep -n "TileKey\|tileKey(" src`: `tileCache.ts` (type, serialization, `invalidate`) and `sheetTiles.ts` (renderer key construction and prefetch). Any other hit is updated the same way;
  - `src/components/OfficeDoc.tsx`:
    - `paintHighlights(ctx, overlay, colors?: Pick<CanvasColors, 'find' | 'find-active' | 'text-selection'>)` (the existing name at `OfficeDoc.tsx:137`);
    - the overlay `key` includes `paletteKey`;
    - the backdrop becomes `background ?? 'var(--ov-color-backdrop)'`: the prop default moves from the `'#888888'` literal to the token, and a host value still wins;
  - `src/components/xlsx/InteractiveSheetDoc.tsx`: `useCanvasColors(themeRoot, theme)` feeds `frame.colors`, `frame.paletteKey` and the overlay painters.

**Interfaces:**
- `export type CanvasToken = 'find' | 'find-active' | 'text-selection' | 'selection' | 'selection-fill' | 'pending' | 'pending-text' | 'freeze'`
- `export type CanvasColors = Partial<Record<CanvasToken, string>>`
- `export function readCanvasColors(themeRoot: Element): CanvasColors`. It returns only non-empty resolved values, so painters fall back to their literals.
- `export function paletteKey(colors: CanvasColors): string`
- `export function useCanvasColors(themeRoot: RefObject<Element | null>, mode: ThemeMode): { colors: CanvasColors; key: string }`. It recomputes on mount, on `mode` change, on `matchMedia('(prefers-color-scheme: dark)')` `change`, and through a `MutationObserver` on the theme root's `style`, `class` and `data-ov-theme` attributes. The object identity stays stable until values change.
- Overrides made only through an ancestor stylesheet apply at the next theme change or remount; this is documented (spec §6.2).

- [ ] **Step 1: Failing tests** in `tests/ui-canvas-theme.test.tsx`:
  - `it('painters use passed colours and keep legacy literals by default')`: a recording context checks `paintGridSelectionOverlay` (`#008000` by default, `#107c41` when passed), `paintFindOverlay`, `paintSearchOverlay`, pending (`#eceff3`/`#666666` by default) and freeze (`#5F6368` by default).
  - `it('paper stays paper: worksheet white and gridlines are never themed')`: `renderSheetViewport` and `SheetTileRenderer.renderFrame` with dark colours still record the `#ffffff` base fill or clear and the `#d0d0d0` gridline stroke.
  - `it('TileCache keys differ by palette')`, called directly on `TileCache`: `put` with `paletteKey: 'a'`, then `get` with `'b'` misses and `get` with `'a'` hits.
  - `it('a palette change misses the tile cache')`: two frames identical except `paletteKey` → the second has `stats.tileHits === 0` and repaints its tiles; the same key → `tileHits > 0`. Tiles of the old palette are dropped only after the new frame is up.
  - `it('the shared CacheIdentity and worker ownership checks are unchanged')`: `CacheIdentity`'s keys are still exactly `docId sheetId valueGen fontGen viewGen`.
  - `it('paged overlay repaints when canvas colours change without remount')`: set `style="--ov-color-find: rgba(0,0,255,.4)"` on the theme root → the page canvas paint effect reruns and records that fill; the canvas node is the same.
  - `it('switching theme repaints the sheet with new overlay colours on the same root node')`
  - `it('canvas token defaults are legible on white paper in both themes')`: for light and dark `THEME_DEFAULTS`, `selection`, `freeze` and `find-active` are ≥ 3:1 against `#ffffff`, and `pending-text` is ≥ 4.5:1 against `pending`.
  - `it('readCanvasColors returns only resolved tokens so painters fall back')`
- [ ] **Step 2:** Run → FAIL.
- [ ] **Step 3:** Implement.
- [ ] **Step 4:** Run → PASS. Run `tests/golden.test.ts` locally on the Mac (unchanged) and all painter tests.
- [ ] **Step 5:** Browser tests (`interactive-viewer-visual.spec.ts` and `paged-toolbar.spec.ts`):
  - `test('dark: sheet paper and gridlines are pixel-identical to light; headers and backdrop are dark')`
  - `test('dark: paged page pixels are white and the backdrop is dark')`
  - `test('selection and find highlights are visible on paper in both themes')`
  - `test('host --ov-color-find override reaches the highlights')`
  - `test('host background prop overrides the backdrop in dark theme')`
  - `test('CLS 0 on an emulateMedia colour-scheme flip')`
- [ ] **Step 6:** Build, budget (`canvasColors.ts` is eager; record its cost) and `check-eager`. Commit `feat(ui): theme-aware canvas overlays with paper preserved and palette-keyed caches`.

### Task 6: Motion (chrome only)

Runs after G (stream F).

**Files:**
- Create: `src/components/motion/presets.ts` (lazy-only). It holds `LazyMotion` + `domAnimation`, the variants `popover`, `dialog`, `sheet`, `fade` and `tick`, durations of 120–220 ms and easing `[0.1, 0.9, 0.2, 1]`.
- Modify:
  - `FilterDropdown.tsx`, `SortDialog.tsx`, `ConditionEditor.tsx`, the `GridSemantics.tsx` popup and the `GridHeaders.tsx` disclosure group;
  - `SheetNavigator.tsx`: the selected sheet label, keyed by sheet;
  - `ViewerStatus.tsx` / the status bar text: keyed by sheet;
  - `DockedZoomControls` and the `ZoomControls` readout: keyed by zoom;
  - `PagedToolbar.tsx`: the More menu and the zoom readout;
  - `PagedStatus.tsx`: position text;
  - `PagedChrome.tsx`: the Copied confirmation;
  - `package.json`: `motion`.
- Do **not** animate the grid pane, page canvases or the content transform.
- Test: `tests/ui-motion.test.tsx`; browser additions.

**Interfaces:**
- `export const popover: Variants`, plus `dialog`, `sheet`, `fade` and `tick`.
- `export function MotionRoot({ children }: { children: ReactNode }): ReactElement` wraps `LazyMotion features={domAnimation} strict`. Components use `m.div`.
- `useReducedMotion()` switches to `duration: 0` variants. `AnimatePresence` is used for chrome only. There is no animation on the viewer's first mount.

- [ ] **Step 1: Failing tests** in `tests/ui-motion.test.tsx`:
  - `it('sheet switch swaps the grid canvas in one commit and never retains an exiting canvas')`: count the `canvas` elements under `[data-officeview-sheet-screen]` before, during the picker's exit animation and after. The count is constant and the new sheet's frame paints in the same commit.
  - `it('zoom change applies the content transform in one step')`: a style-mutation spy records exactly one transform value per zoom action.
  - `it('chrome animates on sheet switch and zoom: picker label, status line and zoom readout are m elements keyed by sheet/zoom')`
  - `it('reduced motion uses zero-duration variants')`: `matchMedia('(prefers-reduced-motion: reduce)')` → the filter dropdown's resolved transition duration is 0.
  - `it('variants animate only transform and opacity')`: the variant keys are a subset of `{opacity, scale, x, y, transition}`.
  - `it('no animation on first mount')`
- [ ] **Step 2:** Run → FAIL.
- [ ] **Step 3:** Run `bun add motion` and implement.
- [ ] **Step 4:** Run → PASS.
- [ ] **Step 5:** Browser tests:
  - `test('reduced motion: popups, More menu and mobile sheet open without animation')`: with `emulateMedia({ reducedMotion: 'reduce' })`, every `element.getAnimations()` has duration ≤ 1 ms, and `ov:motion-reduce:transition-none` elements compute `transition-property: none`.
  - `viewer-layout-shift.spec.ts` (open, scroll, sort, filter, sheet switch, zoom, close) has CLS ≤ 0.001 in light and dark.
  - `paged-toolbar.spec.ts` has CLS ≤ 0.001.
- [ ] **Step 6:** Build, budget (eager unchanged) and `check-eager` (motion is not eager). Record the lazy-chunk gzip growth, then commit `feat(ui): chrome-only motion for menus, dialogs, sheet switch, zoom readout and feedback`.

### Task 8: Verification, docs and release gate

**Files:**
- Modify `README.md`:
  - Usage: `import 'officeview/style.css'`.
  - Props: `theme`, `showToolbar`, `title`.
  - New `### Theming` section before "Interactive spreadsheet viewer":
    - the theme root and override selector (`.brand [data-ov-theme-root] { … }`);
    - a token table matching `COLOR_TOKENS`/`THEME_DEFAULTS`;
    - the paper rule;
    - the live-override limits;
    - the residual host-CSS risk (≥ (0,2,0) or `!important`).
  - Copy and print: hidden buttons, isolated paged print.
- Modify `CHANGELOG.md`: `## [Unreleased]`.
  - Breaking: the stylesheet import.
  - Added: themes, toolbar, page nav, fit-to-width, word count, isolated paged print, motion.
  - Changed: Fluent look; the print guard is per instance.
- Create: `docs/validation/ui-restyle.md`, `tests/readme-tokens.test.ts`.
- Use `.worktrees/visual-shots/scripts/capture.mjs`, `compose.py` and `make_index.py` for before (`main`) and after composites. Cover light and dark, desktop and 375×667 touch, for DOCX, PPTX, static XLSX and interactive XLSX (filter, sort dialog, mobile filter sheet, zoom 50/200).

- [ ] **Step 1:** Add `it('README token table lists exactly COLOR_TOKENS with THEME_DEFAULTS')` in `tests/readme-tokens.test.ts`. It FAILS until the README is written.
- [ ] **Step 2:** Capture before/after composites and publish a gallery for the owner.
- [ ] **Step 3:** agy and codex review the composites visually: style consistency, Fluent fidelity, dark-mode legibility (paper white), overlap and truncation. Fix findings in the owning stream's files.
- [ ] **Step 4:** Run the gates on the final branch and record the numbers in `docs/validation/ui-restyle.md`:
  - `bunx tsc --noEmit`, `bun run typecheck:scripts`, `bunx vitest run`;
  - `bun run build && bun scripts/eager-budget.ts` (≤ 38,720, with the B0→final table);
  - `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts`;
  - `bun scripts/golden-corpus.ts compare --threshold 0 --max-ratio 0`, which must be baseline-identical;
  - `bunx playwright test --config playwright.viewer.config.ts` on Chromium and WebKit.
- [ ] **Step 5:** Docs, CHANGELOG and README. `readme-tokens` passes. Commit `docs: theming, stylesheet import and UI restyle validation`.
- [ ] **Step 6:** Open the PR `feature/ui-restyle` → `main` with the breaking-change note, gallery link, gate table and review trail. The owner merges.

---

## Appendix: Review dispositions (codex / agy / opencode, 2026-10-18)

**codex**

| # | Disposition |
|---|---|
| 1 | Fixed: theme root on the outer wrapper (spec §6.8, R2). Task 1 two-adjacent-viewers test. The inner `className`/`style` contract is preserved. |
| 2 | Fixed: unlayered, scoped utilities and scoped base; `@property`/`*`/`:root` handled by `ovCssScope` (spec §6.8, R3). Task 1 hostile-host geometry/colour test in both load orders. |
| 3 | Fixed: R4 split; `vitest.config.ts` gets its own Tailwind plugin; geometry and colour move to Playwright twins (Tasks 1, 2, 5). |
| 4 | Fixed: decision B / R6. Paper stays white and `surface` is never paper. Task 7 `paper stays paper`. |
| 5 | Fixed: R7 rAF visible-pixel tracking; `goTo` via the container or `panY`. Task 3 tests, including the two-tall-pages case and same-page-count replacement. |
| 6 | Fixed: R8. Word count from the text model (no painting, no watermark-muted span), with ≤ 8 ms slices, abort/generation and a CJK/wrap/header policy (Task 3). |
| 7 | Fixed: R9 iframe isolation, so host and other viewers cannot leak. The guard is instance-scoped; cleanup paths are listed; named `@page` rules. `sizePrintCanvas` is dropped and the sheet path is untouched (Task 4). |
| 8 | Fixed: owner decision A. Chrome-only motion. Task 6 tests that no exiting canvas is retained and the transform changes in one step. |
| 9 | Fixed: shared height constants (regular 40 / compact 48, status 24). Delayed-chunk CLS test; overlap is checked on clipped visible regions (Task 3). |
| 10 | Fixed: `paletteKey` in caches; `MutationObserver` on the theme root; ancestor-only override limit documented (spec §6.2, Task 7). `@theme inline` adopted. |
| 11 | Fixed: fit is a layout mode (R1). Cap `MAX_ZOOM × natural`, zero-width fallback, resize, and zoom-from-fit tests (Task 3). |
| 12 | Fixed: Task 0 measures B0 and lands an independent lazy boundary first. Lucide is lazy-only, with the exact Task 2 fallback stated. Spec §2.6 wording corrected. |
| 13 | Fixed: static-only module map with provenance plus a mutation test (Task 0). The canonical `ov:motion-reduce:` order is adopted, although opencode verified both orders compile. Browser reduced-motion leg in Task 6. |
| 14 | Fixed: keyboard/focus contract (spec §6.5) and Task 3 tests; loader-level file-name tests added. |
| 15 | Fixed: spec status line; single `data-icon` contract (`Icon` takes `name`); merge order, shared-file table and combined gates after each merge. |

**agy**

| # | Disposition |
|---|---|
| 1 | Fixed: lazy-first (Task 0, R5); all new paged UI and print are lazy; budget measured per task. |
| 2 | Fixed: R1 fit layout mode keeps `overflow: auto` and text selection. |
| 3 | Fixed: decision B; worksheet white preserved. |
| 4 | Fixed: R4, with jsdom asserting classes; also moot because `--spacing: 4px` emits px, not rem. |
| 5 | Fixed: decision A, no canvas crossfade. |
| 6 | Fixed: R10 runtime check, not `import.meta.env`; a dist string check proves it survives the build. |
| 7 | Fixed: `paletteKey` in `CacheIdentity` rather than a public `invalidate()`, which keeps the "drop after the new frame is up" ordering. |
| 8 | Fixed: G depends on D and E; merge order and file table. |
| 9 | Fixed differently: prefix belongs on both imports, and the theme import is `reference` (verified with a scratch tailwindcss 4.3.3 compile). |
| 10 | **Rejected**: `OfficeDoc.tsx:137` defines `paintHighlights`, and there is no `paintOverlay`. The plan keeps the correct name. |
| 11 | Fixed: the regex counts numbers. `Intl.Segmenter` was not used because CJK dictionary segmentation differs between engines; per-character CJK counting matches Word. |
| 12 | Fixed: visible pixel height via cached offsets, which needs no IntersectionObserver (and so no jsdom fallback). |
| Missing RF 1 (keyboard) | Fixed in Task 3. |
| Missing RF 2 (resize during fit, threshold crossing) | Fixed in Task 3. |
| Missing RF 3 (wheel/pinch readout, fit off) | Fixed in Task 3. |
| Missing RF 4 (print cancel) | Fixed in Task 4; cleanup never depends on `afterprint`. |
| Missing RF 5 (rapid scrolling) | Fixed in Task 3. |

**opencode**

| # | Disposition |
|---|---|
| 1 | Fixed via R1. The proposed fix (`onFit` sets zoom 1) was rejected: zoom 1 caps at natural width, so it cannot fit a page narrower than the pane. A layout mode is used instead and `zoom.ts` is untouched. |
| 2 | Fixed: `@theme inline reference` plus `ovCssScope`, which gives the same result as the suggested `@utility` route (no `:root`) and also removes `@property`/`*`. |
| 3 | Fixed: G after E (and D). |
| 4 | Fixed: the paged overlay key includes `paletteKey`; Task 7 test. |
| 5 | Fixed: no helper extraction; the iframe print is new code and the sheet path is untouched. |
| 6 | Fixed: Task 0 records B0; every task records before/after. |
| 7 | Fixed: reduced-motion browser leg (Task 6), toolbar keyboard and focus (Task 3), host token override (Task 1), `background` + dark (Task 7). |
| 8 | Fixed: the token list is declared a superset; README lockstep test (Task 8). |
| 9 | Fixed: the Task 2 fallback is concrete, with no forward reference. |
| 10 | Fixed: `PAGE_PADDING_PX = 16` is named. |
| Verified-correct note (jsdom computed `min-height: 44px` feasible) | Superseded by R4. Codex's probe shows jsdom ignores layered rules, so computed sizes are asserted in Playwright. |

### Round 2 (review of 657b6f9): agy APPROVE, opencode APPROVE, codex REQUEST CHANGES

Every claim below was checked against the source and, where it concerns the CSS pipeline, against a scratch build (vite 7.3.1, tailwindcss and `@tailwindcss/vite` 4.3.3).

**codex**

| # | Disposition |
|---|---|
| R2-1 | **Accepted and fixed.** `ovCssScope` is now a PostCSS plugin in the shared `css.postcss.plugins` (Global Constraints, Task 1). A scratch build confirmed that this stage rewrites `dist/style.css`, `?inline` and the dev-server responses. The harness inherits the config through `mergeConfig` and adds no second copy. New checks: the vitest `?inline` pipeline test, `audit-css` on `dist/style.css` with a mutation step (remove the plugin → the audit fails), and the Playwright harness-served stylesheet test. |
| R2-2 | **Accepted and fixed.** `avail = Math.max(0, …)` is restored (`OfficeDoc.tsx:332`); tests cover widths 0, 1, 31, 32 and 33. Fit exit has one rule: inputs act on `z0` with the anchored content point preserved, using `stepZoom`, so 1.613 → 1.0756 → 1. The spec §6.9 wording is aligned, and the readout, wheel and pinch are tested. |
| R2-3 | **Accepted and fixed.** Lazy model walk with no run concatenation; ≤ 2,000-code-unit chunks that never split a surrogate pair; a resumable `createWordCounter` that carries token state across chunks, runs and slices; clock and abort checks after every chunk. The tokens are now whitespace-delimited, which counts a URL as one word. Tests: a 2,000,000-char token, a URL, a 100,000-char CJK paragraph, a 100,000-run paragraph, marks and surrogates at every split offset. The "matching Word" claim is softened to "modelled on". |
| R2-4 | **Accepted and fixed.** `planPrintJob` admits the whole job up front (uniform scale; estimated decoded total ≤ 1 GiB) or refuses it with a notice. Retained blobs are capped at 192 MiB, with an abort and notice when exceeded. The workspace canvas is cleared per page and zeroed at the end. Images are awaited with `load`, not `decode()`. Abort is re-checked at every await point. Tests cover the aggregate cap, the mid-job blob cap and in-flight aborts during the import, `toBlob` and image load, and failure before a `PrintJob` exists. |
| R2-5 | **Accepted and fixed.** New eager `scrollPan.ts`: `contentYAt` includes `scrollTop`, and `enterZoom`/`exitZoom` move the offset atomically (`scrollTop` 0 whenever zoomed). Today's code never resets `scrollTop` on zoom entry (verified: the only write is the search scroll at `OfficeDoc.tsx:634`). Tested as native scroll → zoom → `goTo` → zoom out, in jsdom and in the browser. |
| R2-6 | **Accepted and fixed.** The palette key moves to `TileKey` and its serialized `tileKey()` (`tileCache.ts:333,342–345`) and to `SheetFrame`, plus the renderer's invalidation predicate. The shared `CacheIdentity` (`contracts.ts:440`) and the worker ownership check are untouched. Added a direct `TileCache` put/get test and an identity-unchanged test. |
| R2-7 | **Accepted and fixed.** The eager helpers (`fileName.ts`, `pointer.ts`, `scrollPan.ts`) move to `src/components/`, so the `paged/` allowlist is only `metrics.ts`. The Task 2 glyph fallback is removed: Task 2 depends on Task 0, and if Task 0 is rejected Stream A is re-planned. |
| R2-8 | **Accepted and fixed.** `FloatingControlsProps.copy` gains `onCopy` (the eager `doCopy`, `OfficeDoc.tsx:817`), with click-to-copy and `allowCopy=false` tests. B0 is measured with the existing one-liner before `eager-budget.ts` is created. |

**opencode**

| # | Disposition |
|---|---|
| R2 MINOR 1 (eager allowlist vs Task 3) | Fixed, together with codex R2-7. (`pointer.ts` was already in `src/components/`; `fileName.ts` moved there.) |
| R2 MINOR 2 ((0,1,1) host resets) | Fixed. The base uses `:is()` (0,1,1). That is not relied on: every control must carry explicit colour, background, padding and border utilities at (0,2,0), enforced by a jsdom class check, and the hostile-host Playwright leg is added to Tasks 2, 3 and 5, where the class swap happens. Raising the base to (0,2,0) was rejected, because any selector trick that does it (for example `:not([x])`) ties with or beats the utilities. |
| R2 MINOR 3 (WebKit named pages) | Fixed. `pageRules(specs, namedPages)` has a single-size fallback chosen by `CSS.supports('page', 'p0')`. A Chromium PDF page-box test and a WebKit branch and fit test are added. |
| R2 NIT 4 (duplicated print constants) | Fixed. Cross-reference comments are added in both files; the sheet path stays untouched. |
