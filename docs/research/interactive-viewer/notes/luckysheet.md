# Luckysheet and Successors (FortuneSheet) as Open-Source Web Spreadsheets

## Is Luckysheet dead? Dates/evidence. Which fork should anyone evaluate in 2026?

### Takeaway
Luckysheet is definitively dead: archived, last npm release 2.1.13 in Jan 2021, with two explicit EOL notices (2021, 2024) redirecting to Univer. In 2026 the only community fork worth evaluating is `ruilisi/fortune-sheet` (FortuneSheet, latest v1.0.4 Nov 2025, MIT), while the official maintainer-blessed successor is Univer (Apache-2.0).

### Cited Findings
- Luckysheet originated at `mengshukeji/Luckysheet` and now lives at `dream-num/Luckysheet` (same DreamNum team); repo status is ARCHIVED with ~16.6k stars, ~2.6k forks, 685 open issues — [Source](https://github.com/dream-num/Luckysheet)
- Maintainer EOL notice #799 (Oct 14, 2021): "Luckysheet has stopped maintenance… core team is developing a new office suite Univer" — [Source](https://github.com/dream-num/Luckysheet/issues/799)
- Maintainer EOL notice #1454 (Jan 25, 2024): "Luckysheet is no longer maintained, please use Univer instead… Luckysheet has completed its mission"; README banner repeats this — [Source](https://github.com/dream-num/Luckysheet/issues/1454)
- Last npm release of `luckysheet` is 2.1.13, published Jan 19, 2021 (15 versions, Sep 2020–Jan 2021); package still draws ~5–6.4k weekly downloads on inertia — [Source](https://www.npmjs.com/package/luckysheet)
- Correct fork identity is `ruilisi/fortune-sheet` (NOT `fortuneshhet/FortuneSheet` — that org/name is a typo; no such repo). Created Mar 31, 2022, ~3.6–3.7k stars, ~307–319 forks, 75 open issues / 21 PRs — [Source](https://github.com/ruilisi/fortune-sheet)
- FortuneSheet releases are alive through late 2025: v1.0.4 (Nov 6, 2025, build fix), v1.0.3 (Nov 3, 2025), v1.0.2 (Feb 21, 2025), v1.0.1 (Feb 3, 2025), v0.22.0 (Jan 9, 2025); npm `@fortune-sheet/react` 1.0.4 (MIT) updated Nov 6, 2025 with ~55k weekly downloads — [Source](https://github.com/ruilisi/fortune-sheet/releases); [Source](https://www.npmjs.com/package/@fortune-sheet/react)
- Other `fortune-sheet` repos (`cnpap/`, `qqke/`, `x1028/` forks) are unmodified personal forks of `ruilisi/fortune-sheet`, not contenders — [Source](https://github.com/ruilisi/fortune-sheet/)
- Official successor Univer (`dream-num/univer`): TypeScript, Canvas rendering, Apache-2.0, "actively maintained, highly active" vs Luckysheet "archived"; fixes cited as motivation are large-data loading, chart styles, pivot tables, formula calculation — [Source](https://blog.univer.ai/posts/univer-vs-luckysheet-a-comprehensive-comparison-of-open-source-spreadsheet-solutions/)
- Velocity caveat: FortuneSheet's latest release is Nov 2025 and last-push snapshot was Dec 2025, i.e. slow/small-team cadence (maintainers `sanchit3008`, `Corbe30` plus dependabot) with many 2025 releases being dep-bumps and small bugfixes — [Source](https://github.com/ruilisi/fortune-sheet/releases); [Source](https://javascripts.com/packages/fortune-sheet__core/)

### Inferences
- For a 2026 evaluation: shortlist FortuneSheet if you need a MIT-licensed drop-in Luckysheet-compatible sheet, but treat Univer as the default for anything new/production since it has the original team and active development.
- The stale `luckysheet` download count is a trap signal (legacy embeds), not health.

### Gaps
- No 2026 (Jan–Oct) FortuneSheet commits/releases confirmed; release page shows v1.0.4 (Nov 2025) as latest as of research date.
- Maintainer funding/governance (who `ruilisi` is, bus factor, roadmap beyond 1.0.x) not verified from primary sources.

## How do its sort/filter UIs work and how faithful are they to Excel?

### Takeaway
Both implement Excel-style autofilter/sort (filter dropdowns per column, condition builder, multi-field sort) with color/numeric/date/text filtering, but fidelity is partial: no evidence of advanced Excel behaviors (custom lists, color-scale sorts, date-tree hierarchies, advanced filter dialog), and sort/filter settings are lost on .xlsx import.

### Cited Findings
- Luckysheet feature list claims "Tables: filter, sort"; controller inventory includes `filter.js` (filter ops), `orderByInitial()` (menu-bar sort init), `initialFilterHandler()` wired at workbook init — [Source](https://github.com/dream-num/Luckysheet); [Source](https://cloud.mo4tech.com/excel-luckysheet-source-code-analysis.html)
- FortuneSheet docs: "Filters (Support color, numerical, date, text filtering)" and "Sort (Sort multiple fields simultaneously)"; roadmap marks ✅ sort and ✅ filter — [Source](https://ruilisi.github.io/fortune-sheet-docs/guide/); [Source](https://github.com/ruilisi/fortune-sheet/blob/master/docs/guide/README.md)
- Filter state model (inherited from Luckysheet): `filter_select` defines the range, `filter` holds per-column conditions; each column entry has `caljs` (condition, e.g. `textinclude`, `morethan`, `dateequal`, `include` between with `value1/value2`), computed `rowhidden` map, `optionstate`, and range bounds (`str/edr/stc/edc`); `caljs` takes priority over `rowhidden` (value/color-picked rows) — [Source](https://ruilisi.github.io/fortune-sheet-docs/guide/sheet.html)
- Full `caljs` vocabulary (~20 types): empty/not-empty, text contains/not-contains/starts/ends/exact, date is/before/after, numeric comparisons, between/not-between — [Source](https://ruilisi.github.io/fortune-sheet-docs/guide/sheet.html)
- LuckyExcel (Luckysheet's .xlsx importer) adaptation list covers cell style, border, number/date/percent formats, formulas — while Sort, Filter, conditional formatting, pivot, chart, annotation, and export are explicitly under "Plan" (i.e. NOT preserved) — [Source](https://www.npmjs.com/package/luckyexcel); [Source](https://github.com/dream-num/Luckyexcel)
- FortuneSheet formula side note: docs list dynamic-array functions SORT/FILTER/UNIQUE (Excel-2019 era) as built-ins, i.e. formula-level sort/filter exists alongside UI autofilter — [Source](https://ruilisi.github.io/fortune-sheet-docs/guide/)

### Inferences
- The filter UX to study/borrow: per-column dropdown → condition-type picker → value inputs → hidden-row computation persisted in sheet JSON (`filter` + `rowhidden`). This maps cleanly onto a read-only "interactive viewing" layer (apply `rowhidden` without editing).
- Excel gaps to assume until proven otherwise: multi-level sort dialog parity, sort-by-color/icon, date-group tree, case-sensitive/custom-list sorts, and re-application of filters on data change.

### Gaps
- No primary source (video, spec, or code walkthrough) detailing the exact filter-dropdown UI strings/behaviors or sort-dialog options; demo at `ruilisi.github.io/fortune-sheet-demo` was not interactively tested.
- No citable bug reports quantifying sort/filter incorrectness (e.g. with merged cells, filteredCopy, cross-sheet refs).

## What is the rendering architecture (DOM vs canvas) and what does that imply for large sheets?

### Takeaway
Hybrid: canvas for the grid viewport, DOM for everything else (toolbar, formula bar, dialogs, menus built from HTML template strings). State is centralized in a global `Store` with full-viewport redraws per change — simple to learn from, but the maintainers themselves cite large-data loading and formula performance as the reasons for the Univer rewrite.

### Cited Findings
- Grid viewport is drawn on `<canvas>` from sheet data; surrounding chrome (menu, toolbar, formula bar, sheet bar) is generated as HTML template strings via `luckysheetCreateDOM`/`gridHTML`/`replaceHtml` in `global/createdom.js` — [Source](https://cloud.mo4tech.com/excel-luckysheet-source-code-analysis.html)
- Editing overlays the canvas: a transparent full-canvas div captures mouse events; coordinates map to row/col by binary search over visible-pixel arrays (`colLocationByIndex`, `Store.visibledatacolumn` in `global/location.js`); double-click positions an absolutely-placed input (`luckysheet-rich-text-editor` / `luckysheet-input-box`), keystrokes routed via `controllers/keyboard.js` → `functionInputHandler` → `updateCell` in `global/formula.js` — [Source](https://cloud.mo4tech.com/excel-luckysheet-source-code-analysis.html)
- Centralized `Store`: `Store.luckysheetfile[]` (per-sheet `data` 2D array + `config`), `Store.flowdata` (active sheet's data for fast render access); access via getters/setters (`global/getdata.js`, `global/setdata.js`) and public API (`global/api.js`, e.g. `setCellValue` handles formula-chain update + `jfrefreshgrid` + undo history) — [Source](https://readmex.com/en-US/dream-num/Luckysheet/page-768f7c4dd-5a0a-4eb8-91cc-a82af08554ab)
- Every refresh calls `jfrefreshgrid` → `luckysheetDrawMain`, which re-renders the whole visible area; contemporary code review flags this as unoptimized ("should render only the modified part later") — [Source](https://cloud.mo4tech.com/excel-luckysheet-source-code-analysis.html)
- Luckysheet's own README (via maintainers) says Univer was needed to solve "large data loading… formula calculations… table performance" — [Source](https://github.com/dream-num/Luckysheet)
- Third-party comparison scores large-sheet performance 35/100 for Luckysheet ("canvas keeps small/medium sheets responsive; maintainers list large data loading and formula calculation among problems the rewrite solves") — [Source](https://faun.dev/toolbox/handsontable-vs-luckysheet/) (aggregator; treat score as opinion, quoted rationale as corroborated by maintainers)
- Load-perf escape hatch: `forceCalculation` defaults false — formula cells reuse stored `v/m` on init without recalculation; enabling it "will [cause] performance problems when there are more formulas" — [Source](https://github.com/dream-num/Luckysheet/blob/master/docs/guide/config.md)
- FortuneSheet keeps the canvas grid (`Canvas` class: `drawMain/drawRowHeader/drawColumnHeader/cellRender/cellTextRender`) plus a single global `cellInput` DOM node, but manages state with React/Vue + `immer` (each context change redraws) — [Source](https://github.com/uptonking/note4yaoo/blob/main/lib-excel-luckysheet-codebase.md) (community notes; secondary source)
- FortuneSheet structural fixes vs Luckysheet: full TypeScript, `import/require` (`import { Workbook } from '@fortune-sheet/react'`), multiple instances per page, jQuery dropped, optimized DOM, SVG icons, nothing created outside the container, never touches `window` — [Source](https://github.com/ruilisi/fortune-sheet)
- Ongoing perf work in FortuneSheet releases: `fillText()` perf fix (#615/#578), "edit operation takes 4.5s in large sheets" optimization (#649), text-wrap/break optimizations (#646/#647) — [Source](https://github.com/ruilisi/fortune-sheet/releases)

### Inferences
- For a canvas-based viewer with sort/filter: borrow the coordinate→cell mapping, viewport-only canvas draw, and `filter`→`rowhidden` hidden-row model; do NOT copy full-redraw-per-keystroke or the global-Store singleton if multi-instance or large sheets matter.
- Expect perf cliffs from formula chains and text measurement on large sheets in both Luckysheet and (lesser extent) FortuneSheet; Univer's Web-Worker formula engine is the architecture answer if recalc is needed.

### Gaps
- No published row/column-count benchmarks or frame-time numbers for any of the three projects found.
- Whether FortuneSheet added true viewport virtualization beyond Luckysheet's visible-area draw is unconfirmed.

## Exact licenses. Any patent or attribution concerns?

### Takeaway
Luckysheet = MIT, FortuneSheet = MIT, Univer = Apache-2.0. The one real license incident: FortuneSheet merged HyperFormula (GPLv3) in Jan 2025 then reverted it in v1.0.1 and deprecated the tainted releases — so verify you are on ≥1.0.1/1.0.4 and keep HyperFormula out unless you accept GPLv3.

### Cited Findings
- `dream-num/Luckysheet` license field: MIT License — [Source](https://github.com/dream-num/Luckysheet)
- `ruilisi/fortune-sheet` README: "licensed under the MIT License"; npm `@fortune-sheet/react` lists License: MIT — [Source](https://github.com/ruilisi/fortune-sheet); [Source](https://www.npmjs.com/package/@fortune-sheet/react)
- Univer license: Apache-2.0 (vs Luckysheet MIT) — [Source](https://blog.univer.ai/posts/univer-vs-luckysheet-a-comprehensive-comparison-of-open-source-spreadsheet-solutions/)
- HyperFormula is GPLv3-or-proprietary; the free non-commercial license was removed (v2.1.0 license change; v3.x license-key gating) — [Source](https://hyperformula.handsontable.com/docs/guide/licensing.html); [Source](https://github.com/handsontable/hyperformula/issues/1015)
- FortuneSheet PR #655 "Replace formula-parser with Hyperformula" merged Jan 18, 2025; PR #662 "Revert 'Feature: Hyperformula'" shipped in v1.0.1 (Feb 3, 2025); CI then deprecated `@fortune-sheet/react@1.0.0` ("license conflict due to hyperformula") and `@fortune-sheet/react@1.0.3` ("yarn build issues"), making v1.0.4 the clean 1.0.x — [Source](https://github.com/ruilisi/fortune-sheet/pull/655); [Source](https://github.com/ruilisi/fortune-sheet/releases); [Source](https://github.com/ruilisi/fortune-sheet/compare/v1.0.3...v1.0.4)
- Formula lineage: FortuneSheet README states a forked `handsontable/formula-parser` handles calculations (post-revert status quo) — [Source](https://github.com/ruilisi/fortune-sheet)

### Inferences
- No project-specific patent disputes, patent grants, or attribution lawsuits surfaced; standard MIT practice applies (retain license/copyright notices when vendoring code). MIT has no express patent grant (unlike Apache-2.0) — relevant only if spreadsheet-rendering patents worry you, and no such claim against these projects was found.
- Practical risk is license hygiene, not patents: pin `@fortune-sheet/*` ≥1.0.4, audit transitive deps, and do not reintroduce HyperFormula without a commercial license or GPLv3 compliance.

### Gaps
- Full transitive-dependency license audit (both projects) not performed.
- Luckysheet's bundled third-party assets (iconfont, FontAwesome, ECharts chart plugin) license compatibility not verified file-by-file.

## What are the strongest reasons to adopt vs merely learn from it? (xlsx fidelity, bundle size, pitfalls)

### Takeaway
Learn from it for sort/filter viewing (filter schema, canvas interaction model, formula-chain concepts); adopt only via FortuneSheet (MIT, npm-ready, multi-instance) for internal/low-stakes tools — and only if you accept partial .xlsx fidelity, MB-scale bundles, missing pivot/charts, and a two-person maintenance bus.

### Cited Findings
- .xlsx import (Luckysheet via LuckyExcel): `.xlsx` only (no `.xls`); import covers style/border/number/date/percent formats + formulas; export was still "under development" with community exceljs workarounds — [Source](https://www.npmjs.com/package/luckyexcel); [Source](https://dream-num.github.io/LuckysheetDocs/guide/FAQ.html)
- .xlsx in FortuneSheet: external plugin `corbe30/fortuneexcel` handles import AND export (a gap Luckysheet never closed in-core) — [Source](https://github.com/ruilisi/fortune-sheet)
- Bundle/embedding cost (Luckysheet): npm unpacked 28.8MB / 67 files / 6 deps; `dist/luckysheet.umd.js` 2.91MB (+9.29MB sourcemap); plugin bundle `plugins/js/plugin.js` embeds jQuery; iconfont/FontAwesome/chart assets (`chartmix.umd.min.js` 457.6KB) ship in dist — [Source](https://registry.npmjs.org/luckysheet); [Source](https://cdn.jsdelivr.net/npm/luckysheet/dist/)
- Bundle (FortuneSheet): `@fortune-sheet/core` unpacked 6.5MB; no Bundlephobia aggregate published — [Source](https://javascripts.com/packages/fortune-sheet__core/)
- Extension model (Luckysheet): `initPlugins()` + `expendPlugins` (chart, print — chart noted to throw at registration in one code analysis), plus `hook` callbacks (`cellRenderBefore/After`, `cellAllRenderBefore`, `cellUpdated`, etc.) — [Source](https://cloud.mo4tech.com/excel-luckysheet-source-code-analysis.html); [Source](https://github.com/dream-num/Luckysheet/blob/master/docs/guide/config.md)
- Extension/collab model (FortuneSheet): ✅ hooks + custom tools; collaboration via `onOp` JSON ops (e.g. `{op:'replace', path:['data',1,0,'bl'], value:1}`) with Express+MongoDB `backend-demo` — [Source](https://github.com/ruilisi/fortune-sheet)
- Formula engine (Luckysheet): hand-rolled in `src/function/` — `functionlist.js` metadata (15 categories), `functionImplementation.js` (~each fn's `f` impl), `func.js` utils; custom functions require editing source + all locale packs; supports built-in, remote, and custom formulas — [Source](https://blog.csdn.net/u010593516/article/details/109603965); [Source](https://dream-num.github.io/LuckysheetDocs/guide/FAQ.html)
- Formula engine history: original cross-sheet formula chain had a perf bottleneck; chain was reconstructed (v2.x era) improving render + calc efficiency — [Source](https://dushusir.medium.com/the-backend-of-the-spreadsheet-luckysheet-is-also-open-source-25f43ea4d541)
- Known pitfalls: global jQuery bundled in `plugin.js` conflicts with host apps using their own jQuery (FAQ documents removal procedure) — [Source](https://dream-num.github.io/LuckysheetDocs/guide/FAQ.html); single-instance `window`-global design in Luckysheet vs FortuneSheet's multi-instance support — [Source](https://github.com/ruilisi/fortune-sheet)
- FortuneSheet pre-1.0 warning still relevant in spirit: "input data structure and APIs may change"; migration from Luckysheet renames `sheet.index→sheet.id`; docs/Storybook portions flagged outdated by maintainers themselves — [Source](https://github.com/ruilisi/fortune-sheet)
- Missing advanced surface in FortuneSheet roadmap (unchecked): pivot tables, charts, Vue support, location/search completeness — [Source](https://github.com/ruilisi/fortune-sheet)
- Vendor comparison (biased source — Univer blog) adds: Luckysheet collab is basic updateURL+WebSocket without robust conflict resolution, docs fragmented/partially Chinese, frontend-only protection — [Source](https://blog.univer.ai/posts/univer-vs-luckysheet-a-comprehensive-comparison-of-open-source-spreadsheet-solutions/)

### Inferences
- Strongest adopt reasons (FortuneSheet): MIT license; npm/React packaging; Luckysheet-compatible data model (cheap migration of existing Luckysheet JSON); Excel-plausible filter/sort UI out of the box; real .xlsx round-trip via fortuneexcel.
- Strongest learn-only reasons: dead upstream with 685 orphaned issues; full-redraw canvas loop and global Store don't suit large sheets or multi-instance embedding; formula engine is a forked `formula-parser`, far weaker than HyperFormula/Univer; chart/pivot/print remain weak or external.
- For "interactive sort/filter viewing on an existing canvas renderer": lift the `filter`/`filter_select`→`rowhidden` schema and the overlay/div coordinate-mapping pattern; skip wholesale adoption unless you also need editing.

### Gaps
- No verified production case studies with scale numbers for FortuneSheet; dependent count (26 on npm) is thin evidence.
- fortuneexcel round-trip fidelity (what survives: styles, merges, data validation, conditional formats) not quantified from primary sources.
- No Stack Overflow/Reddit/HN threads with substantive new signal surfaced during this pass.
