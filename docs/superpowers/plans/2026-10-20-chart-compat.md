# Chart Compatibility Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Status:** revision 2. It addresses the plan review from agy, codex and opencode, all of which returned REQUEST CHANGES with no Critical findings. The changes:

- a foundation task comes first;
- a resident/paged data-contract migration task is added;
- oversized tasks are split;
- tests discriminate native rendering from fallback;
- axis bindings are discriminated;
- the eager budget is measured against the project baseline (spec §9 amended);
- the native-fixture prerequisite for ChartEx is restored.

**Goal:** Render every chart a user is likely to meet natively in XLSX, DOCX and PPTX through one lazily loaded chart engine, with a fallback picture or a labelled card instead of blank space or wrong values.

**Architecture:**
1. Lay foundation contracts: workspace, context, point states, the format interface and axis bindings.
2. Move the lazy XLSX engine to `src/charts/`, give all formats one shared `parseChart` adapter, and delete the eager mini-reader.
3. Migrate resident and paged XLSX data to the new contracts.
4. Grow the engine: axes, layout, formats and labels, embedded data, styling, families, analytics and accessibility.
5. ChartEx (stream B) follows after a native-verified field map.

**Tech Stack:** TypeScript, React 19, Vite 7, vitest + jsdom, Playwright (Chromium + WebKit), bun 1.3.10, node-canvas in tests.

**Spec:** `docs/superpowers/specs/2026-10-20-chart-compat-design.md` (approved at `8beadb3`; §9 budget rule amended). It is the binding authority, and each task names its sections.

## Global Constraints

- **Eager budget** (spec §9, amended).
  - Measure with `bun run build && bun scripts/eager-budget.ts`: Node zlib gzip of `dist/index.js`, hard ceiling 38,720.
  - **B0** is the eager size at `3581768`. Task F0 records it in the ledger and the validation doc.
  - After every task, the eager size must be **≤ B0**. Bytes freed earlier (A0a) may be spent later.
  - Every task records before, after, the running total and the bytes it adds. Always rebuild before measuring.
  - New logic goes lazy. **Never add the embedded-zip inventory or uncached read to `src/core/zip.ts`**; at most a one-line entry accessor.
- **Eager graph.**
  - Run `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts`.
  - A0a adds `src/charts/` and `src/xlsx/chart.ts` to `EAGER_FORBIDDEN`.
  - `paintChartPlaceholder` stays eager.
- **Runtime tag.** The tag stays `kind:'xlsxChart'`; types and modules become neutral.
- **Representation contract** (spec §6.3). The adapter returns `{kind:'native',model} | {kind:'unsupported',reason} | {kind:'refused',reason}`. `unsupported` and `refused` never satisfy an `mc:Choice`, and a card is created only when no representation remains.
- **Point states** (spec §5.8): value, blank, error(code), unavailable. They apply to numeric roles, and errors keep their identity until plotting.
- **Analytic acquisition is separate from display truncation.**
  - Analytic inputs (trendlines, stdDev/stdErr, histogram, Pareto, box & whisker, treemap/sunburst) are acquired **before** any display cap, up to 65,536 per series. Above that, the chart (or the element) is refused.
  - Display caps (64 series / 4,096 points) apply only to what is drawn.
- **Ownership** (spec §6.4/§6.5). One per-document `ChartWorkspace` (64/32 MiB) charges every chart allocation before it is made.
  - It never runs as unlimited or no-op admission, even in early tasks.
  - It relates to the paged drawing admission by **transfer, not double-charge** (D0).
- **Two stages** (spec §6.5).
  - Stage 1 is async and size-independent, and produces an immutable model.
  - Stage 2 is synchronous layout in a realm-local WeakMap sidecar keyed by (w, h, dpr, fontGeneration), charged until it is evicted.
- **Paper rule.** Dark mode never changes chart pixels.
- **Native-success tests must reject fallbacks.** Any test of "renders natively" asserts:
  - the model kind is native (not a card, not a fallback picture);
  - the coverage representation is `chart`;
  - distinct native geometry, not just non-white pixels.
- **Goldens.**
  - A0a: existing XLSX chart goldens must be **pixel-identical (diff 0)**; no re-record.
  - Later tasks: identical except the enumerated, individually inspected spec §5 deltas, listed in the commit with reasons.
  - Pixel goldens are Mac-only, never run in CI.
- **Fixtures.**
  - Synthetic builder fixtures are **"Office-shaped synthetic"** and are never reported as Office evidence.
  - Real Office-authored files are evidence, with provenance (producer and version) recorded.
- **Tests per task:**
  - `bunx tsc --noEmit`;
  - the focused vitest files;
  - the full `bunx vitest run` once before reporting;
  - build plus budget plus `check-eager`;
  - affected Playwright specs with a unique `IV_HARNESS_PORT` (Chromium + WebKit).
  - codex cannot bind localhost or write git metadata; the controller runs those commands and commits for it.
- **CI is Linux and slower.** Drive clocks explicitly; never rely on wall time. `tests/setup/styles.ts` injects only the theme contract.
- **No new runtime dependencies.**
- **Review.** A different model family reviews each task. B1's field map gets two families.

## Review Focus

| # | Input | Expected | Pinned by |
|---|---|---|---|
| 1 | A paged chart whose **values role has no cache**, with the data only in the embedded workbook | Renders natively from the workbook | A3b `cacheless values role renders natively from the embedded workbook`, plus the A7 browser test |
| 2 | A ChartEx chart wrapped in `mc:AlternateContent` with a fallback picture, before stream B | The fallback picture keeps painting | A0b `chartex choice with fallback keeps the fallback picture` |
| 3 | Finance chart with `dispUnits` = millions | Tick labels divide | A1a `display units divide tick values` (with A2 formatting) |
| 4 | Line chart with #N/A points | The line spans | A2 `#N/A spans a line regardless of dispBlanksAs` |
| 5 | 200 charts or 50 embeddings | Bounded; later charts become cards; no throw | C1 robustness slice `200 charts stay within the chart workspace` |

---

## Order, branches and agents

Tasks run **sequentially**:

F0 → A0a → A0b → D0 → A1a → A1b → A2 → A3a → A3b → A4a → A4b → A4c → A4d → A5a → A5b → A5c → A5d → A6a → A6b → A7 → B1 → B2 → B3 → B4 → B5 → C1.

Each task gets a branch off `feature/chart-compat`, merged back `--no-ff` after review, followed by the combined gates.

| Task | Implementer | Reviewer |
|---|---|---|
| F0 foundation contracts | Claude | codex |
| A0a engine move, shared adapter, delete mini-reader | Claude | codex |
| A0b representation contract, containment, exports, coverage | Claude | opencode |
| D0 resident/paged XLSX data-contract migration | Claude | codex |
| A1a axis graph, scales, dates | Claude | codex |
| A1b layout, cache, text and fonts | codex | Claude |
| A2 display formats, data labels, point matrix | codex | Claude |
| A3a embedded zip admission, uncached read, cancellation | Claude | codex |
| A3b embedded cells, SST, formats, names; data precedence | Claude | codex |
| A4a markers, lines, fills, style precedence, invertIfNegative, explosion, dPt | agy | Claude |
| A4b titles, legend, multi-level categories, plotVisOnly | agy | Claude |
| A4c smoothing | Claude | opencode |
| A4d user shapes | Claude | codex |
| A5a bubble, radar (+ family conformance harness) | Claude | opencode |
| A5b stock | codex | Claude |
| A5c pie-of-pie, bar-of-pie | Claude | opencode |
| A5d contour, surface3D card | Claude | codex |
| A6a trendlines | codex | Claude |
| A6b error bars | Claude | codex |
| A7 accessibility, coverage, diagnostics, browser | Claude | agy |
| B1 ChartEx native parts, field map, parser, dispatch | codex | Claude + opencode |
| B2 waterfall, funnel | Claude | codex |
| B3 histogram, Pareto | Claude | codex |
| B4 box & whisker | agy | Claude |
| B5 treemap, sunburst | Claude | opencode |
| C1 robustness slice, gallery goldens, validation doc, README, CHANGELOG, PR | Claude | codex + agy |

**Step pattern for every task:**
1. Write the failing tests.
2. Run them and see them fail.
3. Implement.
4. Run them and see them pass, along with the listed regression files.
5. Run build, budget (running total ≤ B0) and `check-eager`.
6. Run the browser tests, when listed.
7. Commit with the given message.

Only the specifics are listed below.

---

### Task F0: Foundation contracts

**Spec:** §5.8 point states, §6.3 result union, §6.4 workspace/ownership, §6.5 two stages, §5.1 axis model, §5.2 format interface.

**Files:**
- Create:
  - `src/charts/model.ts`: types-only `ChartModel`, `ChartGroup` with **discriminated axis bindings**, `ChartPoint`, `ChartRole` and `ChartResult`;
  - `src/charts/workspace.ts`;
  - `src/charts/context.ts`;
  - `src/charts/format-api.ts` (interface only).
- Modify: `src/xlsx/chartModel.ts`. It **adds** `export type * from '../charts/model'` alongside its existing `Xlsx*` types, so the existing engine still compiles. The full swap happens in A0a. [plan r2: opencode N1]
- Test: `tests/charts-foundation.test.ts`.

**Interfaces (produced):**
- `type ChartPoint = {state:'value'; v:number; fmt?:string} | {state:'blank'} | {state:'error'; code:string} | {state:'unavailable'}`.
- `type ChartAxisBinding = {kind:'none'} | {kind:'cartesian'; category?:string; x?:string; value:string} | {kind:'xy'; x:string; y:string} | {kind:'radial'; value:string} | {kind:'surface'; category:string; series:string; value:string}`.
  - Pie and doughnut use `none`; radar uses `radial`; contour uses `surface`.
- `type ChartResult = {kind:'native'; model:ChartModel} | {kind:'unsupported'; reason:string} | {kind:'refused'; reason:string}`.
- `interface ChartWorkspace { readonly limit:number; reserve(bytes:number, tag:string): boolean; transfer(from:string, to:string, bytes:number): void; release(tag:string): void; hold<T>(p:Promise<T>, tag:string): Promise<T> }`.
  - `hold` keeps the charges until the promise settles, even after an abort.
  - `createChartWorkspace(profile:'desktop'|'mobile')` returns a limit of 64 MiB or 32 MiB.
- `interface ChartParseContext { theme?: ThemeContext; pkg: OfficePackage; partPath: string; signal?: AbortSignal; profile:'desktop'|'mobile'; workspace: ChartWorkspace; epoch?: '1900'|'1904' }`.
- `interface ChartFormatter { format(v:number, code:string|undefined, epoch:'1900'|'1904'): string }`. A2 implements it.
- `interface TextMeasure { width(text:string, font:string): number; lineHeight(font:string): number }`.
- **Production hooks:** `documentChartContext(profile, signal)`, created once per parse. It is threaded by `parseDocx`, `parsePptx` and `parseXlsx` options into `ContentLoadAdapters`. A standalone or transferred render uses its own workspace, created on demand.

- [ ] Tests:
  - `it('workspace reserves, transfers and releases; refuses past the limit; never unlimited')`;
  - `it('hold keeps charges until a gated producer settles after abort')`;
  - `it('axis bindings: pie none, cartesian named roles, xy, radial, surface; invalid references rejected at parse')`;
  - `it('documentChartContext is created once per parse and reaches the XLSX adapter')`. DOCX and PPTX get the shared adapter in A0a, and their reach is asserted there. [plan r2: opencode N3]
- [ ] Record B0 in `.superpowers/sdd/progress.md` and in `docs/validation/chart-compat.md`, which this task creates as a stub.
- [ ] Commit `feat(charts): foundation contracts for the chart engine`.

### Task A0a: Engine move, shared adapter, delete the mini-reader

**Spec:** §6.1, §6.2.

**Files:**
- Create: `src/charts/index.ts` (the lazy entry), plus `src/charts/parse.ts`, `paint.ts` and `fill.ts`, which move the code from `src/xlsx/chart.ts` **behaviour-identically**. Then delete `src/xlsx/chart.ts`.
- Modify:
  - `src/xlsx/chartLoader.ts`: imports `../charts`;
  - `src/drawing/content.ts`:
    - a shared default `parseChart` in `ContentLoadAdapters` (:370-376), with caller adapters spread over the defaults;
    - minimal `ChartResult` consumption in `follow()`: `native` → content; anything else → today's `unsupported-content` path. The full MC/card contract is A0b. [plan r2: opencode N2]
    - delete `chart()` (:251-293) and `kind:'chart'`;
  - `src/drawing/coverage.ts`: export `featureOfKind`; `src/docx/parse.ts:700` uses it. This moves from A0b, so A0a's coverage assertion holds.
  - `src/drawing/content-paint.ts`: delete `paintChart`;
  - `src/xlsx/drawing.ts`: its adapter becomes the default;
  - `tests/helpers/eager-graph.ts`: add `src/charts/` and `src/xlsx/chart.ts` to `EAGER_FORBIDDEN`.
- Test:
  - new `tests/charts-adapter.test.ts`;
  - update the `kind:'chart'` assertions (`tests/drawing-content.test.ts:277`, `tests/docx-complex-features.test.ts:294`) and `tests/drawing-coverage.test.ts:74-113`;
  - repoint the imports in `tests/xlsx-chart*.test.ts`, keeping every assertion.

- [ ] Tests:
  - `it('docx and pptx stacked bar, horizontal bar, line and pie parse natively as xlsxChart')`: assert the model, coverage `chart` and group geometry;
  - `it('a caller passing only text adapters still gets the default parseChart; a caller parseChart overrides')`;
  - `it('src/charts is not in the eager graph')`, from the module map.
- [ ] Gates:
  - all `tests/xlsx-chart*.test.ts` assertions pass unchanged;
  - `OFFICEVIEW_STRICT_GOLDEN=1 bunx vitest run tests/golden.test.ts` passes on the Mac, with **diff 0 and no re-record**;
  - the budget records the drop.
- [ ] Commit `feat(charts): one lazy chart engine for all formats; delete the eager column-chart reader`.

### Task A0b: Representation contract, containment, exports, coverage

**Spec:** §6.3, §6.5 (bounded chart XML, painter containment, `ensureChartEngine`, `preloadCharts`), §6.6 (`featureOfKind`).

**Files:**
- Modify:
  - `src/drawing/content.ts`: the `ChartResult` handling in `follow()` and `prepareCompatibleDrawingContent`; a bounded chart-part read before tree construction, through the bounded tokenizer path; chart dispatch by namespace (ChartEx still `unsupported` until B1);
  - `src/drawing/content-paint.ts`: save/try/finally/restore, with a card on throw;
  - `src/render/paint.ts`: `ensureChartEngine` in the DOCX and PPTX branches, catching failure;
  - `src/index.ts`: export `preloadCharts`.
- Test: `tests/charts-representation.test.ts`.

- [ ] Tests:
  - `it('unsupported and refused results do not satisfy an mc:Choice; the fallback picture is selected')`. Use **injected** adapter outcomes plus `surface3DChart`, which stays unsupported, in all three formats.
  - `it('chartex choice with fallback keeps the fallback picture')`.
  - `it('standalone unsupported chart becomes a card with the reason, the original feature and one representation')`.
  - `it('a broken fallback picture becomes a card recording chart-fallback-unusable')`.
  - `it('nested and alias-prefixed AlternateContent and unknown Requires')`.
  - `it('a chart part over the byte cap is refused before any XML parse')`: assert that no parser or tokenizer call happens and that the inflated-byte counter stays under the cap, including with a forged declared size.
  - `it('a throwing painter paints a card and restores canvas state for later shapes')`.
  - `it('getPaintables returns pages with cards when the chart chunk fails')`.
  - `it('renderSlide paints a card when the engine was never loaded in this realm')`.
  - `it('preloadCharts registers the painter')`.
  - `it('a card reports coverage chart-placeholder')`.
- [ ] Commit `feat(charts): representation contract, bounded chart parts and painter containment`.

### Task D0: Resident and paged XLSX data-contract migration

**Spec:** §6.4 (XLSX clause, per-role precedence), §5.8 point states, §7 analytic vs display limits, §5.8 `plotVisOnly`.

**Files:**
- Modify:
  - `src/charts/fill.ts`;
  - `src/charts/parse.ts`: the cache/literal readers emit `ChartPoint` and keep sparse indices;
  - `src/xlsx/parse.ts` (resident fill, :459-469);
  - `src/xlsx/view/pagedSource.ts`: `fillCharts` at :513-737 keeps error identity (today :628-629 collapses errors), keeps unavailable saved formula values, and keeps hidden-row/column provenance for `plotVisOnly`;
  - `src/xlsx/storage/pagedDrawings.ts`.
- Test:
  - `tests/charts-data-contract.test.ts`;
  - extend `tests/xlsx-viewer-paged-drawings.test.ts`;
  - keep `tests/xlsx-c-review-r3-chart-memory.test.ts` and the lease/stream tests green.

**Rules:**
- **Analytic acquisition before display caps.** Analytic roles are acquired up to 65,536 points per series; display caps apply at draw time.
- **Ownership.** The paged path's existing drawing admission stays authoritative for its backing storage. The chart workspace is charged by **transfer** when a filled model is retained, never as a second charge on the same bytes. Worker → main copies are charged once on the receiving side.

- [ ] Tests:
  - `it('resident and paged fills produce identical ChartPoint arrays for the same workbook')`, covering caches vs references, blanks, #N/A vs #DIV/0!, and formulas without saved values;
  - `it('analytic roles keep all points up to 65536: a 5000-point series is not truncated to 4096 for analysis')`, using a 4,097-point dataset where the last point changes a computed mean;
  - `it('65537 analytic points refuse; display still truncates at 4096 for non-analytic families')`;
  - `it('plotVisOnly skips hidden rows and columns; plotVisOnly=0 plots them')`;
  - `it('paged admission is not double-charged when a model is retained in the chart workspace')`;
  - `it('existing paged streaming and lease protections still hold')`.
- [ ] Commit `refactor(charts): xlsx resident and paged data use point states and analytic acquisition`.

### Task A1a: Axis graph, scales, dates

**Spec:** §5.1 (all), §6.5 rendering bounds.

**Files:**
- Create: `src/charts/axes.ts` and `src/charts/dates.ts`, which use `src/xlsx/formula/serial.ts`.
- Test: `tests/charts-axes.test.ts` and `tests/charts-dates.test.ts`. These are pure model and number tests; native family renders come in A5.

- [ ] Tests (numeric):
  - `it('groups bind axes by axId; a secondary value axis has its own scale')`;
  - `it('a deleted axis still controls the scale')`;
  - `it('crosses autoZero, min, max and crossesAt place the crossing')`;
  - `it('explicit min and max, majorUnit and minorUnit, minor gridlines and tick marks, tickLblPos nextTo, low, high and none')`;
  - `it('xy bindings keep independent x and y domains; a category group plus an XY group keep separate x domains')`;
  - `it('an unplaceable combination is chart-axis-unplaceable')`;
  - `it('crossBetween between vs midCat')`;
  - `it('orientation maxMin')`;
  - `it('logBase skips non-positive values with chart-invalid-log-value')`;
  - `it('display units divide tick values and name the unit label')`, covering builtInUnit and custUnit;
  - `it('date axis steps by calendar months and years and keeps unequal spacing')`;
  - `it('text categories on a date axis fall back to categorical with a diagnostic')`;
  - `it('chart epoch precedence and val-omitted date1904 is true')`;
  - `it('same-epoch serial 60 is preserved; cross-epoch conversion uses the civil model')`, covering serials 1, 59, 60, 60.5 and 61 in both directions, the unrepresentable cases, mixed cached and filled provenance, and fractions;
  - `it('ticks are bounded to 1000; a tiny majorUnit is enlarged with a diagnostic')`.
- [ ] Commit `feat(charts): axis graph, scales, date axes and display units`.

### Task A1b: Layout, cache, text and fonts

**Spec:** §6.5 stage 2, §5.10 manual layout, §6.7, §5.1 label rotation and thinning.

**Files:**
- Create: `src/charts/layout.ts`. It holds `layoutChart(model, w, h, dpr, fontGeneration, measure: TextMeasure): ChartLayout` and the WeakMap sidecar.
- Test: `tests/charts-layout.test.ts`.

- [ ] Tests:
  - `it('layout is cached per size, dpr and font generation; resize and font load invalidate; cached layouts stay charged; paint transients are refunded')`;
  - `it('a structuredClone-transferred model paints synchronously in a fresh realm with its own workspace')`;
  - `it('theme minor font via the font manager; txPr size clamped 6-72')`;
  - `it('rotation from txPr bodyPr rot and vert; tickLblSkip and automatic skip never overlap labels')`;
  - `it('manualLayout: edge and factor modes, defaults factor and outer, resolution order, inverted box falls back with a diagnostic')`.
- [ ] Commit `feat(charts): placement-keyed chart layout with charged caches`.

### Task A2: Display formats, data labels, point matrix

**Spec:** §5.2, §5.3, §5.8 matrix.

**Files:**
- Create: `src/charts/format.ts` (`ChartFormatter` over `src/xlsx/cellFormat.ts`) and `src/charts/labels.ts`.
- Modify: `src/charts/paint.ts`.
- Test: `tests/charts-format.test.ts`, `tests/charts-labels.test.ts`, `tests/charts-matrix.test.ts`.

- [ ] Tests:
  - `it('formats: scientific, fraction, elapsed, date, percent, locale prefix')`;
  - `it('sourceLinked per-point label formats; axis uses the first non-blank (acceptance, unverified until reference)')`;
  - `it('codes over 255 units are General; output capped at 256')`;
  - `it('label precedence group then series then point; explicit false overrides; delete hides')`;
  - `it('content flags, separator, rich text, leader lines; dLblPos defaults; bestFit for pie')`;
  - `it('field-reference and manual-layout labels fall back with chart-label-feature-unsupported')`;
  - `it('#N/A spans a line regardless of dispBlanksAs')`;
  - matrix cells for line and scatter × gap, zero and span; column; area × gap, zero and span; stacked overrides; pie; `dispNaAsBlank=1`. **Radar cells move to A5a.**
  - `it('a label on an #N/A point shows #N/A')`.
- [ ] Commit `feat(charts): display formats, data labels and blank and error plotting`.

### Task A3a: Embedded zip admission, uncached read, cancellation

**Spec:** §6.4 (resolve, admit, ownership, cancellation).

**Files:**
- Create: `src/charts/embeddedZip.ts` (lazy). It covers chart `.rels` → `externalData` resolution, the central-directory inventory peek, the uncached bounded inflate and the `hold` integration.
- Modify: `src/core/zip.ts`, **at most one line**: an accessor for the raw entry. Measure it.
- Test: `tests/charts-embedded-zip.test.ts`, plus the builder `tests/helpers/chart-fixtures.ts`, which makes Office-shaped synthetic DOCX/PPTX with `c:externalData` and an embedded xlsx.

- [ ] Tests:
  - `it('external targets and package-escaping paths are unavailable')`;
  - `it('admission: inflated cap 8 and 4 MiB, at most 256 entries, declared-size sum checked before load; a forged declared size is caught by the inflated counter')`;
  - `it('embedding bytes are not retained in the outer package cache')`;
  - `it('one embedded read at a time per document')`;
  - `it('cancellation stops publication immediately; charges are held until the gated inflate settles')`;
  - `it('a corrupt embedded zip is unavailable with a diagnostic')`.
- [ ] Commit `feat(charts): bounded embedded-workbook access for chart data`.

### Task A3b: Embedded cells, SST, formats, names; data precedence

**Spec:** §6.4 (read, values, budgets, per-role precedence, required roles).

**Files:**
- Create: `src/charts/embeddedWorkbook.ts` and `src/charts/dataPrecedence.ts`. Sheets are read with the streaming `XmlTokenizer` (`src/core/xml-stream.ts:140`).
- Test: `tests/charts-embedded-workbook.test.ts`, `tests/charts-precedence.test.ts`.

- [ ] Tests:
  - `it('cacheless values role renders natively from the embedded workbook')`, an Office-shaped synthetic fixture (Review Focus 1). Assert the native model, the filled values and coverage `chart`.
  - `it('cached-only chart never opens the embedding')`.
  - `it('saved data wins: caches and literals; sparse indices kept; ptCount=0 present-empty')`.
  - `it('literal-only and mixed literal-reference series')`.
  - `it('a missing required role makes the chart unsupported')`.
  - `it('formula cells without a saved value are unavailable points; with a saved value they use it')`.
  - `it('only referenced sheets and ranges are read; reading stops after the last referenced row')`. Assert with allocation counters, not only tokenizer calls.
  - `it('names resolved with bounds; a name pointing at an external workbook is unavailable')`.
  - `it('shared strings over the byte or count budget refuse and refund')`.
  - `it('50 distinct embeddings and repeated references stay within the workspace')`.
- [ ] Commit `feat(charts): embedded-workbook data and per-role precedence for docx and pptx charts`.

### Task A4a: Markers, lines, fills, style precedence, series options

**Spec:** §5.4, §5.11 colour precedence, §5.8 series and group options.

**Files:**
- Create: `src/charts/style.ts`, `src/charts/markers.ts` and `src/charts/fills.ts`.
- Test: `tests/charts-style.test.ts`.

- [ ] Tests:
  - `it('marker symbols, auto cycle, size clamp 2-72, marker spPr')`;
  - `it('line dash, width, cap and noFill')`;
  - `it('solid, gradient, pattern and noFill for series, points, plot area, chart area, legend and title')`;
  - `it('chart area noFill is transparent; default via c:style')`;
  - `it('colour precedence spPr, chartStyle, colorStyle, c:style, then accents with lumMod variants after six')`;
  - `it('invertIfNegative fills negative bars with the inverted fill')`;
  - `it('series and point explosion geometry')`;
  - `it('dPt overrides take precedence over series formatting')`;
  - `it('gapWidth, overlap and varyColors for existing families')`;
  - `it('dark mode leaves chart pixels identical')`.
- [ ] List the golden deltas.
- [ ] Commit `feat(charts): markers, fills, chart styles and series options`.

### Task A4b: Titles, legend, multi-level categories

**Spec:** §5.9, §5.10 (title, legend).

**Files:**
- Create: `src/charts/categories.ts`.
- Modify: `src/charts/layout.ts`.
- Test: `tests/charts-text.test.ts`.

- [ ] Tests:
  - `it('title content order: rich, strRef cache, automatic, autoTitleDeleted')`;
  - `it('run formatting precedence rPr, paragraph, txPr, style, theme')`;
  - `it('legend position, overlay and legendEntry delete')`;
  - `it('multi-level categories draw grouped tiers; noMultiLvlLbl shows the innermost only')`.
- [ ] Commit `feat(charts): chart titles, legends and multi-level categories`.

### Task A4c: Smoothing

**Spec:** §5.5.

- [ ] Tests (geometry):
  - `it('smooth passes through every point with Catmull-Rom tangents; end segments use the end point')`;
  - `it('blanks break the curve per dispBlanksAs')`;
  - `it('unequal x spacing for scatter')`.
- [ ] The validation doc marks smoothing as an **approximation, unverified** until a reference exists.
- [ ] Commit `feat(charts): smoothed lines`.

### Task A4d: User shapes

**Spec:** §5.11 user shapes.

**Files:**
- Create: `src/charts/userShapes.ts`. It parses `c:userShapes` through the existing relationship and depth budget, prepares the nodes and images with the existing drawing preparation, and paints them via `src/drawing/scene-paint.ts`.
- Test: `tests/charts-usershapes.test.ts`.

- [ ] Tests:
  - `it('relSizeAnchor and absSizeAnchor place shapes over the chart')`;
  - `it('user-shape images use the image budget; relationship cycles are refused')`;
  - `it('missing user-shapes part leaves the chart native with a diagnostic')`.
- [ ] Commit `feat(charts): chart user shapes`.

### Task A5a: Bubble, radar, family conformance harness

**Spec:** §4.1 bubble and radar, §7 every-family robustness.

**Files:**
- Create:
  - `src/charts/families/bubble.ts` and `radar.ts`;
  - `tests/helpers/chart-family-conformance.ts`, a parameterized harness covering empty series, all-blank values, NaN/∞, negative values and a single point. It asserts no throw, a native or card result, and the diagnostics. Every later family task runs it.
- Test: `tests/charts-families-bubble-radar.test.ts`.

- [ ] Tests:
  - `it('bubble: bubbleScale, sizeRepresents area and width, zero not drawn, showNegBubbles 0 and 1')`;
  - `it('radar: clockwise from 12, radial axis, standard, marker and filled styles')`;
  - the radar cells of the §5.8 matrix;
  - the conformance harness over every existing and new family.
- [ ] Commit `feat(charts): bubble and radar charts`.

### Task A5b: Stock

- [ ] **File:** `src/charts/families/stock.ts`.
- [ ] Tests:
  - `it('volume from the bar group; open, high, low, close by order; authored axes')`;
  - `it('HLC, OHLC, volume-HLC and volume-OHLC')`;
  - `it('hiLowLines and upDownBars colours')`;
  - `it('a role mismatch becomes a card')`;
  - the conformance harness.
- [ ] Commit `feat(charts): stock charts`.

### Task A5c: Pie-of-pie and bar-of-pie

- [ ] **File:** `src/charts/families/ofPie.ts`.
- [ ] Tests (numeric):
  - `it('auto splits ceil(n/3) last points; pos, val, percent and cust splits')`;
  - `it('the Other slice conserves totals and percent labels; original indices keep colours and labels')`;
  - `it('secondPieSize, serLines and bar-of-pie gapWidth')`;
  - the conformance harness.
- [ ] Commit `feat(charts): pie-of-pie and bar-of-pie charts`.

### Task A5d: Contour and surface3D card

- [ ] **File:** `src/charts/families/contour.ts`. In `src/charts/parse.ts`, `surface3DChart` becomes `unsupported`.
- [ ] Tests:
  - `it('contour bands from the major unit; grid at most 256x256')`;
  - `it('no interpolation across missing cells')`;
  - `it('wireframe draws boundaries only; bandFmts colours')`;
  - `it('surface3D uses the fallback picture, else a card')`;
  - the conformance harness.
- [ ] Commit `feat(charts): contour charts`.

### Task A6a: Trendlines

**Spec:** §5.6, §7 analytic.

- [ ] **File:** `src/charts/trendlines.ts`.
- [ ] Tests (numeric answers):
  - `it('linear, exponential, logarithmic, polynomial 2-6, power and moving average on known datasets')`;
  - `it('XY uses x values; category charts use 1..n')`;
  - `it('non-positive exclusions per type with diagnostics')`;
  - `it('a singular polynomial fit draws nothing with chart-singular-trendline')`;
  - `it('intercept for linear, exponential and polynomial only')`;
  - `it('extension bounded to 10x the data domain with a diagnostic')`;
  - `it('out-of-range order or period draws nothing, never clamped')`;
  - `it('equation and R-squared labels: authored numFmt, else 4 significant digits')`;
  - `it('at most 512 samples')`;
  - `it('a 4097-point series fits on all points')`, where the last point changes the slope;
  - `it('65537 points drop the trendline')`.
- [ ] Commit `feat(charts): trendlines`.

### Task A6b: Error bars

**Spec:** §5.7.

- [ ] **File:** `src/charts/errorBars.ts`. `src/charts/fill.ts` adds the custom plus/minus roles.
- [ ] Tests (geometry and numeric):
  - `it('x and y directions; plus, minus and both; noEndCap segment geometry')`;
  - `it('fixed, percentage, stdDev (n-1, centred on the series mean), stdErr')`;
  - `it('custom plus and minus from caches and the embedded workbook; a missing custom point draws no bar')`;
  - `it('analytic stdDev over 65536 drops the bars')`.
- [ ] Commit `feat(charts): error bars`.

### Task A7: Accessibility, coverage, diagnostics, browser

**Spec:** §6.6, §8 (browser).

**Files:**
- Create: `src/charts/describe.ts`, moved and extended from `src/xlsx/chartDescription.ts`.
- Modify:
  - `src/render/paint.ts`: paintables carry descriptions with placement, charged to the workspace;
  - `src/components/paged/PagedChrome.tsx` (lazy): a per-page offscreen `role="img"` list. **Do not edit eager `OfficeDoc.tsx` beyond passing data.**
- Test:
  - `tests/charts-a11y.test.tsx`;
  - `tests/charts-coverage.test.ts`;
  - `tests/browser/paged-charts.spec.ts`.

- [ ] Tests:
  - `it('two charts on one page expose two role=img entries in reading order')`;
  - `it('a header chart repeated on each page appears once per page')`;
  - `it('cards expose placeholder: reason')`;
  - `it('allowCopy false keeps descriptions readable, not selectable')`;
  - `it('coverage: native chart, card chart-placeholder, fallback keeps the original feature')`;
  - `it('every new diagnostic code is emitted with a reason')`.
- [ ] Browser tests (they must reject fallbacks):
  - `test('docx and pptx charts render natively in the paged viewer (browser)')`: assert native geometry plus the accessible description naming a native type;
  - `test('cacheless embedded-workbook chart renders natively (browser)')`;
  - `test('dark mode leaves chart pixels identical (browser)')`;
  - `test('CLS 0 while charts load (browser)')`.
- [ ] Commit `feat(charts): accessible chart descriptions and coverage for paged documents`.

### Task B1: ChartEx native parts, field map, parser, dispatch

**Spec:** §4.2, §4.3, §6.2, §6.3.

- [ ] **Step 0: Native parts (prerequisite, spec §4.3).**
  - Acquire at least one **Office-authored** part per family: waterfall, funnel, histogram, Pareto, box & whisker, treemap, sunburst, region map.
  - Sources: public test corpora such as Apache POI test-data, LibreOffice `qa` data and the Open-XML-SDK test files. Owner-supplied files are also welcome.
  - Record the provenance (file, producer and version, licence) in `tests/fixtures/chartex/PROVENANCE.md`.
  - **A family with no Office-authored part is blocked.** It stays on the fallback picture and is listed in the validation doc. It is not built from hand-written XML.
- [ ] **Step 1:** Write `docs/charts/chartex-map.md` against MS-ODRAWXML and the native parts, covering every spec §4.3 bullet and its defaults. **Gate:** two families approve it.
- [ ] **Files:**
  - `src/charts/chartex/parse.ts`, which extends the **one** `ChartModel` contract with a discriminated ChartEx group kind, not a second model;
  - `src/drawing/content.ts`: remove the unconditional chartex rejection (:869) and the structural rejection; dispatch `cx:`;
  - the consumer Choice capability lists (`src/docx/drawing.ts:102`, PPTX, XLSX);
  - **eager budget:** keep the dispatch to a namespace check; everything else stays lazy.
- [ ] Tests:
  - `it('chartex dimensions parse by id with levels and sparse points')`, using the native parts;
  - `it('series bind dimensions through dataId; layoutId selects the family')`;
  - `it('region map, blocked families and unreadable parts return unsupported so the fallback picture is kept')`;
  - `it('cx choices are attempted in docx, pptx and xlsx when the engine loads')`.
- [ ] Commit `feat(charts): chartex field map, parser and dispatch`.

### Tasks B2–B5: ChartEx families

Each task runs the conformance harness and refuses analytic inputs over the limit, never using a prefix. Its numeric tests use the native parts plus hand-computed datasets.

- **B2 waterfall and funnel** (`src/charts/chartex/waterfall.ts`, `funnel.ts`):
  - `it('subtotals show the running total anchored at zero; the accumulator continues')`;
  - `it('increase, decrease and total colours; connector visibility')`;
  - `it('funnel centred bars proportional to value in source order')`.
  - Commit `feat(charts): waterfall and funnel charts`.
- **B3 histogram and Pareto** (`src/charts/chartex/histogram.ts`):
  - `it('binSize, binCount and automatic Scott bins')`;
  - `it('intervalClosed right and left')`;
  - `it('underflow and overflow, explicit and auto')`;
  - `it('interval label notation')`;
  - `it('pre-aggregated data drawn as given')`;
  - `it('pareto descending with stable ties; cumulative percent on its own axis')`;
  - `it('a 4097-sample histogram uses every sample')`;
  - `it('65537 samples refuse')`.
  - Commit `feat(charts): histogram and pareto charts`.
- **B4 box & whisker** (`src/charts/chartex/boxWhisker.ts`):
  - `it('exclusive and inclusive split halves with or without the median; odd and even n')`;
  - `it('whiskers within 1.5 IQR; outliers when visible')`;
  - `it('mean marker, mean line and inner points')`.
  - Commit `feat(charts): box and whisker charts`.
- **B5 treemap and sunburst** (`src/charts/chartex/treemap.ts`, `sunburst.ts`):
  - `it('full-path identity; duplicates aggregate; missing level is a leaf')`;
  - `it('non-positive weights excluded')`;
  - `it('squarified areas proportional to weight; parentLabelLayout none, banner and overlapping')`;
  - `it('sunburst rings by depth with angles proportional to subtree weight')`;
  - `it('over 4096 nodes refuses')`.
  - Commit `feat(charts): treemap and sunburst charts`.

### Task C1: Robustness slice, gallery, validation, docs, PR

**Spec:** §7, §8, §9.

**Files:**
- Create:
  - `tests/charts-robustness.test.ts`;
  - `src/testdata/chart-gallery.ts`.
- Modify:
  - `tests/golden.test.ts` (gallery entries);
  - `docs/validation/chart-compat.md`;
  - `README.md` "Charts" section;
  - `CHANGELOG.md`.

- [ ] Robustness tests:
  - `it('200 charts stay within the chart workspace; later charts become cards')` (Review Focus 5);
  - Torture fixtures, split by damage class, per format [plan r2: opencode N4]. Each asserts no throw, the expected diagnostics, and the allocation/work counters:
    - (a) every family plus ChartEx, model3D and OMML;
    - (b) a huge embedded workbook and a corrupt embedded zip;
    - (c) external `externalData` and a name pointing at an external workbook;
    - (d) empty, NaN and ∞ data.
  - `it('repeated paints and cache churn release layout charges')`;
  - `it('path, label-byte and contour-grid bounds drop elements with diagnostics')`.
- [ ] `it('README chart family list matches the engine family registry')`.
- [ ] Gallery goldens: one DOCX, PPTX and XLSX each, recorded on the Mac, every one individually inspected.
- [ ] **Validation doc:**
  - the gate table;
  - B0 → final eager, with every task's running total;
  - a reference-status table **keyed by spec §8 case IDs**, each verified or unverified;
  - residual diffs;
  - blocked ChartEx families.
- [ ] Full gates:
  - tsc and typecheck:scripts;
  - full vitest;
  - budget;
  - check-eager;
  - golden-corpus compare, identical except the enumerated deltas;
  - full Playwright.
- [ ] Commit `docs: chart compatibility validation, README and CHANGELOG`.
- [ ] Open the PR `feature/chart-compat` → `main`, closing #13 and superseding #12. Merge only after CI and the bot review pass.
