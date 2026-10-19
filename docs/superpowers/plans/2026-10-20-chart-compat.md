# Chart Compatibility Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Render every chart a user is likely to meet natively in XLSX, DOCX and PPTX through one lazily loaded chart engine, with a fallback picture or a labelled card instead of blank space or wrong values.

**Architecture:**
- Move the lazy XLSX chart engine to `src/charts/`. Give DOCX and PPTX the same `parseChart` adapter XLSX uses, and delete the eager clustered-column mini-reader.
- Then grow the engine:
  - an axis model by ID;
  - the format adapter;
  - data labels and point states;
  - a bounded embedded-workbook reader;
  - fills, markers and layout;
  - new classic families;
  - trendlines and error bars;
  - accessibility.
- ChartEx (stream B) follows after a normative field map.

**Tech Stack:** TypeScript, React 19, Vite 7, vitest + jsdom, Playwright (Chromium + WebKit), bun 1.3.10, node-canvas for tests.

**Spec:** `docs/superpowers/specs/2026-10-20-chart-compat-design.md`. It is the binding authority. Every task's requirements include the spec sections it names.

## Global Constraints

- **Eager budget.**
  - Measure with `bun run build && bun scripts/eager-budget.ts`. The script reports the Node zlib gzip of `dist/index.js` and exits 1 above **38,720**. The baseline at `3581768` is about 38,631 B.
  - Every task records before and after numbers in its commit message, and must come out **net ≤ 0**.
  - Always rebuild before measuring.
- **Eager graph.**
  - Run `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts`.
  - Task A0 adds `src/charts/` and `src/xlsx/chart.ts` to `EAGER_FORBIDDEN` in `tests/helpers/eager-graph.ts`. After that, nothing under `src/charts/` may be statically reachable from `src/index.ts`.
  - `paintChartPlaceholder` stays eager (spec §6.1).
- **Runtime tag.** The tag stays `kind:'xlsxChart'` (spec §6.2). Types and modules become neutral (`ChartModel`, `src/charts/*`).
- **Representation contract** (spec §6.3). The adapter returns `{ kind:'native', model } | { kind:'unsupported', reason } | { kind:'refused', reason }`.
  - `unsupported` and `refused` must not satisfy an `mc:Choice`.
  - A card is created only when no representation remains.
- **Point states** (spec §5.8): value, blank, error(code), unavailable. Errors keep their identity until the family plotting rule.
- **Analytic inputs are never truncated** (spec §4.3, §7). If they exceed 65,536 points per series, refuse the chart.
- **Limits** are in spec §6.4 (embedded workbook and the chart workspace, 64/32 MiB) and §6.5 (work bounds). Charge before allocation. Chart XML is read with an 8/4 MiB cap before the tree is built.
- **Two stages** (spec §6.5):
  - stage 1 is async and size-independent, and produces an immutable model;
  - stage 2 is synchronous layout, kept in a realm-local WeakMap sidecar keyed by (w, h, dpr, fontGeneration) and charged until it is evicted.
- **Paper rule.** Dark mode never changes chart pixels.
- **Goldens.**
  - No golden is re-recorded without individual inspection and a stated reason.
  - Existing XLSX chart goldens must stay pixel-identical, unless a spec §5 element legitimately changes them; each such change is listed in the commit.
  - Pixel goldens are Mac-only and are never run in CI.
- **Tests per task:**
  - `bunx tsc --noEmit`;
  - the focused vitest files;
  - the full `bunx vitest run` once before reporting;
  - build plus budget plus `check-eager`;
  - affected Playwright specs with a unique `IV_HARNESS_PORT` (Chromium + WebKit).
  - Agents whose sandbox cannot bind localhost (codex) list the commands, and the controller runs them.
- **CI is Linux and slower than macOS.** Avoid wall-clock races in tests; drive clocks explicitly. Do not inject the full stylesheet into jsdom tests; `tests/setup/styles.ts` injects only the theme contract by default.
- **No new runtime dependencies** without owner approval.
- **Review.** Every task is reviewed by a model family other than its implementer's. B1 (field map) gets two reviewing families.

## Review Focus

These are the input classes most likely to bite users that no single task's tests fully exercise. Each one is pinned by a test in the task named.

1. **A real Office DOCX/PPTX chart with caches and an embedded workbook,** opened in the paged viewer: it must draw natively, never as a card. Owned by A0 (adapter) and A3 (data), as `office-shaped paged chart renders natively (browser)`.
2. **A ChartEx chart wrapped in `mc:AlternateContent` with a fallback picture,** before stream B lands: the fallback picture must keep painting, never a card. Owned by A0, as `chartex choice with fallback keeps the fallback picture`.
3. **Finance charts using `dispUnits` = millions:** tick labels must divide. Owned by A1, as `display units divide tick labels and draw the unit label`.
4. **A line chart with #N/A points:** the line spans the gap. Owned by A2, as `#N/A spans a line regardless of dispBlanksAs`.
5. **A document with many charts** (200 on one page, or 50 embeddings): bounded memory, later charts become cards, no throw. Owned by A3 and A7, as `200 charts stay within the chart workspace`.

---

## Streams and agents

Tasks run **sequentially** in this order: A0 → A1 → A2 → A3 → A4 → A5 → A6 → A7 → B1 → B2 → B3 → B4 → B5 → C1. Each task lives on its own branch off `feature/chart-compat` and is merged back with `--no-ff` after review, followed by the combined gates.

| Task | Implementer | Reviewer |
|---|---|---|
| A0 engine move + adapter seam + representation contract | Claude | codex |
| A1 axis model + two-stage layout | Claude | codex |
| A2 format adapter + data labels + point states | codex | Claude |
| A3 embedded workbook reader + data precedence + chart workspace | Claude | codex |
| A4 markers, lines, smoothing, fills, style, categories, layout, legend, user shapes | agy | Claude |
| A5 new classic families (bubble, radar, stock, ofPie, contour, surface3D card) | Claude | opencode |
| A6 trendlines + error bars | codex | Claude |
| A7 accessibility + coverage + diagnostics + browser | Claude | agy |
| B1 ChartEx field map + dispatch | codex | Claude + opencode |
| B2 waterfall + funnel | Claude | codex |
| B3 histogram + Pareto | Claude | codex |
| B4 box & whisker | agy | Claude |
| B5 treemap + sunburst | Claude | opencode |
| C1 gallery goldens, validation doc, README, CHANGELOG | Claude | codex + agy |

---

### Task A0: Engine move, shared adapter, representation contract

**Spec:** §6.1, §6.2, §6.3, §6.5 (bounded XML read, painter containment, `ensureChartEngine`, `preloadCharts`), §6.6 (coverage `featureOfKind` export).

**Files:**
- Create:
  - `src/charts/index.ts`, the lazy entry, which re-exports parse, paint, fill and register;
  - `src/charts/model.ts` (types; this replaces `src/xlsx/chartModel.ts`, and the old path re-exports it);
  - `src/charts/parse.ts`, `src/charts/paint.ts` and `src/charts/fill.ts`, which move the code out of `src/xlsx/chart.ts`. The old file is deleted, and its imports are repointed.
- Modify:
  - `src/xlsx/chartLoader.ts`: `loadChartModule()` imports `../charts`;
  - `src/drawing/content.ts`:
    - a shared default `parseChart` in `ContentLoadAdapters` (around :370-376), merged under caller adapters;
    - delete `chart()` (:251-293) and its `kind:'chart'` variant;
    - a bounded chart-part read;
    - chart dispatch by namespace;
    - the representation contract in `follow()` and `prepareCompatibleDrawingContent`;
    - keep ChartEx returning `unsupported` until B1;
  - `src/drawing/content-paint.ts`: delete `paintChart`; wrap the registered painter in save/try/finally/restore with a card on throw;
  - `src/docx/drawing.ts`, `src/pptx/parse.ts`: rely on the default adapter (remove any `parseChart` gaps);
  - `src/render/paint.ts`: `ensureChartEngine(contents: DrawingContent[])` in the DOCX and PPTX branches (it catches failure);
  - `src/index.ts`: `export { preloadCharts }`;
  - `src/drawing/coverage.ts`: export `featureOfKind`; `src/docx/parse.ts:700` uses it;
  - `tests/helpers/eager-graph.ts`: add `src/charts/` and `src/xlsx/chart.ts` to `EAGER_FORBIDDEN`.
- Test:
  - new `tests/charts-adapter.test.ts`, `tests/charts-representation.test.ts`;
  - update `tests/drawing-content.test.ts:277` and `tests/docx-complex-features.test.ts:294`: `kind:'chart'` → `kind:'xlsxChart'`;
  - update `tests/drawing-coverage.test.ts:74-113`: horizontal bar is now native.

**Interfaces:**
- Produces:
  - `type ChartResult = { kind:'native'; model: ChartModel } | { kind:'unsupported'; reason: string } | { kind:'refused'; reason: string }`;
  - `parseChartPart(root: XmlNode, ctx: ChartParseContext): Promise<ChartResult>` (lazy). `ChartParseContext = { theme: ThemeContext | undefined; pkg: OfficePackage; partPath: string; signal?: AbortSignal; profile: 'desktop'|'mobile'; workspace: ChartWorkspace }`;
  - `ensureChartEngine(contents: readonly DrawingContent[]): Promise<void>`;
  - `preloadCharts(): Promise<void>`;
  - `featureOfKind(kind: DrawingContent['kind']): string`.
- Later tasks extend `ChartModel` (A1+) and `ChartParseContext` (A3 adds `embedded`).

- [ ] **Step 1: Failing tests.**
  - `it('docx and pptx charts parse through the shared lazy adapter as xlsxChart')`: a stacked horizontal bar, a line and a pie in a DOCX and a PPTX built with `src/testdata/ooxml-builders.ts`, plus cached `c:` XML. Each yields `kind:'xlsxChart'`, not unsupported.
  - `it('a caller passing only text adapters still gets the default parseChart')`.
  - `it('unsupported and refused results do not satisfy an mc:Choice; the fallback picture is selected')`: radar Choice plus picture Fallback in all three formats.
  - `it('chartex choice with fallback keeps the fallback picture')` (Review Focus 2).
  - `it('standalone unsupported chart becomes a card with the reason and feature')`.
  - `it('a broken fallback picture becomes a card recording chart-fallback-unusable')`.
  - `it('a chart part over the byte cap is unsupported before the tree is built')`: spy on `pkg.xmlOrdered` and assert it is not called.
  - `it('a throwing chart painter paints a card and restores canvas state')`.
  - `it('getPaintables returns pages with cards when the chart chunk fails to load')`.
  - `it('renderSlide paints a card when the engine was never loaded in this realm')`.
  - `it('preloadCharts loads and registers the painter')`.
  - `it('docx coverage reports chart, not xlsxChart')`.
  - `it('existing xlsx chart behaviour is unchanged')`: run `tests/xlsx-charts*.test.ts` unchanged.
- [ ] **Step 2:** Run them and see them fail.
- [ ] **Step 3:** Implement. Move the code first and keep it behaviour-identical, then add the seam. Delete the mini-reader in the same task, so the eager bytes go down.
- [ ] **Step 4:** Run them and see them pass. Also pass `tests/xlsx-chart*.test.ts`, `tests/drawing-*.test.ts`, `tests/pptx-drawings.test.ts`, `tests/docx-complex-features.test.ts` and `OFFICEVIEW_STRICT_GOLDEN=1 bunx vitest run tests/golden.test.ts` (Mac).
- [ ] **Step 5:** Run build, budget (record the drop) and `check-eager`.
- [ ] **Step 6: Browser.** In `tests/browser/paged-toolbar.spec.ts` or a new `tests/browser/paged-charts.spec.ts`, add `test('docx and pptx charts render natively in the paged viewer (browser)')`, which counts non-white pixels in the chart box.
- [ ] **Step 7:** Commit `feat(charts): one lazy chart engine for xlsx, docx and pptx; delete the eager column-chart reader`.

### Task A1: Axis model by ID and two-stage layout

**Spec:** §5.1 (all), §5.10 manual layout, §6.5 (two stages, work bounds, rendering bounds).

**Files:**
- Create:
  - `src/charts/axes.ts`: axis graph, scales, ticks, date axes, display units;
  - `src/charts/layout.ts`: stage-2 layout and the sidecar cache;
  - `src/charts/dates.ts`: epoch resolution and conversion through `src/xlsx/formula/serial.ts`.
- Modify: `src/charts/parse.ts` and `src/charts/paint.ts`, to consume the axis graph and layout.
- Test: `tests/charts-axes.test.ts`, `tests/charts-layout.test.ts`, `tests/charts-dates.test.ts`.

**Interfaces:**
- Produces:
  - `ChartAxis = { id: string; kind:'cat'|'val'|'date'|'ser'; pos:'b'|'t'|'l'|'r'; crossAx: string; deleted: boolean; scale: AxisScale; ... }`;
  - `ChartModel.axes: ChartAxis[]`;
  - `ChartGroup.axIds: [string, string]`;
  - `layoutChart(model: ChartModel, w: number, h: number, dpr: number, fontGeneration: number, measure: TextMeasure): ChartLayout`, cached in a `WeakMap<ChartModel, Map<key, ChartLayout>>` and charged to the workspace.
- A2 consumes `ChartAxis.format`. A5 consumes the axis bindings for stock and bubble.

- [ ] **Step 1: Failing tests** (numeric oracles):
  - `it('groups bind axes by axId; secondary value axis has its own scale')`;
  - `it('scatter and bubble use independent x and y value axes')`;
  - `it('a category group plus an XY group keep separate x domains')`;
  - `it('an unplaceable axis combination becomes a card with chart-axis-unplaceable')`;
  - `it('crossBetween between vs midCat positions categories')`;
  - `it('orientation maxMin reverses the axis')`;
  - `it('logBase scales and skips non-positive values with chart-invalid-log-value')`;
  - `it('display units divide tick labels and draw the unit label')` (Review Focus 3): builtInUnit millions and custUnit;
  - `it('date axis steps by calendar months and years and keeps unequal spacing')`;
  - `it('date axis over text categories falls back to categorical with a diagnostic')`;
  - `it('chart epoch precedence: c:date1904, then workbook, then 1900')`;
  - `it('epoch conversion uses the civil model; serial 0, 60 and pre-1904 dates are unrepresentable')`: serials 1, 59, 60, 60.5 and 61, in both directions;
  - `it('tick generation is bounded to 1000 per axis; a tiny majorUnit is enlarged with a diagnostic')`;
  - `it('tickLblSkip and automatic skip prevent overlapping labels')`;
  - `it('label rotation from txPr bodyPr rot')`;
  - `it('manualLayout edge and factor modes, defaults factor and outer, inverted box falls back')`;
  - `it('layout is cached per size, dpr and font generation; resize and font load invalidate; cached layouts stay charged')`;
  - `it('a transferred model paints synchronously in a fresh realm')` (structuredClone).
- [ ] **Steps 2–5:** See them fail, implement, see them pass. Existing XLSX chart goldens stay identical; list any intended change. Run build, budget and `check-eager`.
- [ ] **Step 6:** Commit `feat(charts): axis model by id, date axes, display units and placement-keyed layout`.

### Task A2: Display formats, data labels, point states

**Spec:** §5.2, §5.3, §5.8.

**Files:**
- Create:
  - `src/charts/format.ts`, the chart display-format adapter over `src/xlsx/cellFormat.ts`;
  - `src/charts/labels.ts`;
  - `src/charts/points.ts`.
- Modify: `src/charts/parse.ts`, `src/charts/fill.ts` and `src/charts/paint.ts`.
- Test: `tests/charts-format.test.ts`, `tests/charts-labels.test.ts`, `tests/charts-points.test.ts`.

**Interfaces:**
- Produces:
  - `type ChartPoint = { state:'value'; v: number; fmt?: string } | { state:'blank' } | { state:'error'; code: string } | { state:'unavailable' }`;
  - `ChartSeries.values: ChartPoint[]`;
  - `formatChartValue(v: number, code: string | undefined, epoch: '1900'|'1904'): string`, capped at 256 units, with the 255-unit code cap.
- A3 fills `ChartPoint`, and A6 consumes it.

- [ ] **Step 1: Failing tests:**
  - `it('axis and label formats: scientific, fraction, elapsed, date, percent, locale prefix')`;
  - `it('sourceLinked uses per-point cache formats for labels and first non-blank for the axis')`;
  - `it('codes over 255 units render General; output capped at 256')`;
  - `it('data label precedence group then series then point; explicit false overrides; delete hides')`;
  - `it('label content flags, separator, authored rich text and leader lines')`;
  - `it('dLblPos defaults per family; bestFit for pie')`;
  - `it('field-reference labels and manual label layout fall back with chart-label-feature-unsupported')`;
  - `it('point states survive fill: blank, #N/A, other error, unavailable')`;
  - `it('#N/A spans a line regardless of dispBlanksAs')` (Review Focus 4);
  - one test per cell of the spec §5.8 matrix:
    - line/scatter/radar × gap, zero and span;
    - column;
    - area × gap, zero and span;
    - stacked;
    - pie;
    - `dispNaAsBlank=1`;
  - `it('a data label on an #N/A point shows #N/A')`.
- [ ] **Steps 2–6:** As in A1. Commit `feat(charts): display formats, data labels and point states`.

### Task A3: Embedded workbook reader, data precedence, chart workspace

**Spec:** §6.4 (all), §7.

**Files:**
- Create:
  - `src/charts/embeddedWorkbook.ts` (lazy);
  - `src/charts/workspace.ts`;
  - `src/charts/dataPrecedence.ts`.
- Modify:
  - `src/core/zip.ts`: add an uncached bounded read, `bytesUncached(path, max)`, and a central-directory inventory peek;
  - `src/charts/fill.ts` and `src/charts/parse.ts`.
- Test: `tests/charts-embedded.test.ts`, `tests/charts-precedence.test.ts`, `tests/charts-workspace.test.ts`, and a builder helper in `tests/helpers/chart-fixtures.ts` that makes a DOCX/PPTX with `c:externalData` and an embedded xlsx.

**Interfaces:**
- Produces:
  - `ChartWorkspace` with `reserve(bytes): boolean`, `transfer`, `release` and `settle(promise)`. It is owned per document, with a limit of 64/32 MiB.
  - `readEmbeddedRanges(pkg: OfficePackage, chartPath: string, refs: ChartRef[], ctx): Promise<EmbeddedCells | { refused: string } | { unavailable: string }>`.

- [ ] **Step 1: Failing tests:**
  - `it('saved data wins: caches and literals, sparse idx kept, ptCount=0 is present-empty')`;
  - `it('cache-absent references are read from the embedded workbook')`;
  - `it('office-shaped paged chart renders natively')`: caches and an embedded workbook (Review Focus 1);
  - `it('a missing required role makes the chart unsupported, never a silently missing series')`;
  - `it('mixed cached and workbook-filled roles are allowed')`;
  - `it('external targets and names pointing at external workbooks are unavailable')`;
  - `it('formula cells without a saved value are unavailable points')`;
  - `it('only referenced sheets and ranges are read; reading stops after the last referenced row')`, using a spy on the tokenizer;
  - `it('admission: inflated cap, central directory entries and declared sizes are checked before load')`;
  - `it('a sub-cap zip with an oversized inner sharedStrings is refused and fully refunded')`;
  - `it('50 distinct embeddings and repeated references stay within the workspace; one read at a time')`;
  - `it('cancellation stops publication immediately; charges are held until the gated producer settles')`;
  - `it('embedding bytes are not retained in the outer package cache')`;
  - `it('a corrupt embedded zip leaves caches or a card')`;
  - `it('200 charts stay within the chart workspace; later charts become cards')` (Review Focus 5).
- [ ] **Steps 2–6:** As in A1. Commit `feat(charts): bounded embedded-workbook data for docx and pptx charts`.

### Task A4: Markers, lines, smoothing, fills, style, categories, text, layout, legend, user shapes

**Spec:** §5.4, §5.5, §5.9, §5.10 (title, legend), §5.11.

**Files:**
- Create: `src/charts/style.ts`, `src/charts/markers.ts`, `src/charts/fills.ts` and `src/charts/categories.ts`.
- Modify: `src/charts/parse.ts` and `src/charts/paint.ts`. `src/charts/paint.ts` uses `src/drawing/scene-paint.ts` for user shapes through the existing relationship and depth budget.
- Test: `tests/charts-style.test.ts`, `tests/charts-markers.test.ts`, `tests/charts-categories.test.ts`, `tests/charts-usershapes.test.ts`.

- [ ] **Step 1: Failing tests:**
  - `it('marker symbols, size clamp 2-72, marker spPr')`;
  - `it('line dash, width, cap and noFill')`;
  - `it('smooth draws Catmull-Rom Bezier through every point; blanks break per dispBlanksAs')`;
  - `it('solid, gradient, pattern and noFill for series, points, plot area, chart area, legend and title')`;
  - `it('chart area noFill is transparent; default per producer via c:style')`;
  - `it('colour precedence spPr then chartStyle then colorStyle then c:style then accents with lumMod variants after six')`;
  - `it('multi-level categories draw grouped tiers; noMultiLvlLbl shows the innermost only')`;
  - `it('title content and run formatting precedence')`;
  - `it('legend position, overlay and legendEntry delete')`;
  - `it('userShapes draw through the scene painter within the relationship budget')`;
  - `it('dark mode leaves chart pixels identical')`.
- [ ] **Steps 2–6:** As in A1. List every existing XLSX golden that legitimately changes (for example, `noFill` chart area) with its reason. Commit `feat(charts): markers, smoothing, fills, chart styles, multi-level categories, layout and user shapes`.

### Task A5: New classic families

**Spec:** §4.1 rows for bubble, radar, stock, ofPie, contour and surface3D, plus §5.8 group options.

**Files:**
- Create: `src/charts/families/bubble.ts`, `radar.ts`, `stock.ts`, `ofPie.ts` and `contour.ts`.
- Modify: `src/charts/parse.ts` (remove these from the unsupported list; `surface3DChart` → `unsupported`).
- Test: `tests/charts-families-classic.test.ts`.

- [ ] **Step 1: Failing tests** (numeric oracles):
  - `it('bubble: bubbleScale, sizeRepresents area and width, zero not drawn, showNegBubbles')`;
  - `it('radar: clockwise from 12, radial value axis, filled and marker styles')`;
  - `it('stock: volume from the bar group; open, high, low, close by order; authored axes; hiLowLines and upDownBars; a role mismatch becomes a card')`;
  - `it('ofPie: auto, pos, val, percent and cust splits; the Other slice conserves totals and percentages; secondPieSize; serLines; bar-of-pie gapWidth')`;
  - `it('contour: bands from the major unit, no interpolation across missing cells, wireframe boundaries only, bandFmts colours')`;
  - `it('surface3D uses the fallback picture, else a card')`;
  - `it('dropLines and serLines')`.
- [ ] **Steps 2–6:** As in A1. Commit `feat(charts): bubble, radar, stock, pie-of-pie and contour charts`.

### Task A6: Trendlines and error bars

**Spec:** §5.6, §5.7, §7 analytic rows.

**Files:**
- Create: `src/charts/trendlines.ts` and `src/charts/errorBars.ts`.
- Modify: `src/charts/parse.ts`, `src/charts/fill.ts` (custom plus/minus are data roles) and `src/charts/paint.ts`.
- Test: `tests/charts-trendlines.test.ts` and `tests/charts-errorbars.test.ts`.

- [ ] **Step 1: Failing tests:** numeric answers on known datasets.
  - `it('linear, exponential, logarithmic, polynomial 2-6, power and moving average fit known datasets')`;
  - `it('XY charts fit x values; category charts fit indices 1..n')`;
  - `it('non-positive values are excluded per type with diagnostics')`;
  - `it('a singular polynomial fit draws nothing with chart-singular-trendline')`;
  - `it('intercept applies to linear, exponential and polynomial only')`;
  - `it('forward and backward extension is bounded to 10x the data domain with a diagnostic')`;
  - `it('an out-of-range order or period draws nothing with a diagnostic, never clamped')`;
  - `it('equation and R-squared labels use the authored numFmt, else 4 significant digits')`;
  - `it('trendline sampled to at most 512 points')`;
  - `it('error bars fixed, percentage, stdDev n-1 centred on the mean, stdErr and custom from data sources')`;
  - `it('analytic inputs over 65536 points drop the trendline or error bar, never a prefix')`.
- [ ] **Steps 2–6:** As in A1. Commit `feat(charts): trendlines and error bars`.

### Task A7: Accessibility, coverage and diagnostics finalization, browser

**Spec:** §6.6, §8 (browser).

**Files:**
- Modify:
  - `src/render/paint.ts`: paintables carry chart descriptions with placement;
  - `src/components/OfficeDoc.tsx` (eager, budget!) or the lazy `src/components/paged/PagedChrome.tsx`: a per-page offscreen list of `role="img"` entries. Prefer the lazy side;
  - `src/charts/describe.ts`, which moves `chartDescription` and extends it with the placeholder reason.
- Test:
  - `tests/charts-a11y.test.tsx`;
  - `tests/browser/paged-charts.spec.ts`;
  - `tests/charts-coverage.test.ts`.

- [ ] **Step 1: Failing tests:**
  - `it('two charts on one page expose two role=img entries in reading order')`;
  - `it('a header chart repeated on each page appears once per page')`;
  - `it('a card exposes placeholder: reason')`;
  - `it('allowCopy false keeps descriptions readable but not selectable')`;
  - `it('coverage native chart, card chart-placeholder, fallback picture keeps the original feature')`;
  - `it('all new diagnostic codes are emitted with reasons')`;
  - Browser:
    - `test('chart descriptions are exposed per page (browser)')`;
    - `test('dark mode leaves chart pixels identical to light (browser)')`;
    - `test('CLS 0 while charts load (browser)')`.
- [ ] **Steps 2–6:** As in A1. Commit `feat(charts): accessible chart descriptions and coverage for paged documents`.

### Task B1: ChartEx field map and dispatch

**Spec:** §4.2, §4.3, §6.2 (ChartEx dispatch, MC capability lists), §6.3.

**Files:**
- Create:
  - `docs/charts/chartex-map.md`, the normative map;
  - `src/charts/chartex/parse.ts`: data dimensions, `layoutId` and series binding, producing a ChartEx model;
  - `tests/fixtures/chartex/`: one Office-authored or hand-built-from-spec part per family, with provenance recorded.
- Modify:
  - `src/drawing/content.ts`: remove the unconditional chartex rejection (:869) and the structural rejection, and dispatch `cx:` to the engine;
  - `src/docx/drawing.ts:102` and the PPTX/XLSX Choice capability lists: accept `cx` when the engine can load.
- Test: `tests/charts-chartex-parse.test.ts`.

- [ ] **Step 1:** Write `docs/charts/chartex-map.md` against MS-ODRAWXML (CT_ChartSpace, CT_ChartData, CT_Data, CT_NumericDimension, CT_StringDimension, CT_Series with layoutId, CT_Binning, CT_SeriesLayoutProperties, CT_Subtotals, ST_QuartileMethod, ST_ParentLabelLayout). It must cover every spec §4.3 bullet, with defaults, and cite its sources. **Review gate:** two reviewing families approve the map before Step 2.
- [ ] **Step 2: Failing tests:**
  - `it('chartex data dimensions parse by id with levels and sparse points')`;
  - `it('series bind dimensions through dataId; layoutId selects the family')`;
  - `it('region map and unreadable parts return unsupported so the fallback picture is kept')`;
  - `it('cx choices are attempted in docx, pptx and xlsx when the engine loads')`;
  - `it('nested and alias-prefixed AlternateContent and unknown Requires')`.
- [ ] **Steps 3–6:** Implement, see the tests pass, run build, budget and `check-eager`. Families still return `unsupported` until B2–B5. Commit `feat(charts): chartex field map, parser and dispatch`.

### Task B2: Waterfall and funnel

**Spec:** §4.3 waterfall and funnel.

- [ ] **Files:** `src/charts/chartex/waterfall.ts` and `funnel.ts`; test `tests/charts-chartex-waterfall-funnel.test.ts`.
- [ ] **Tests (numeric):**
  - `it('waterfall subtotals show the running total anchored at zero and the accumulator continues')`;
  - `it('increase, decrease and total colours and connector visibility')`;
  - `it('funnel centred bars proportional to value in source order')`;
  - plus a gallery entry.
- [ ] Commit `feat(charts): waterfall and funnel charts`.

### Task B3: Histogram and Pareto

**Spec:** §4.3 histogram and Pareto; §7 analytic.

- [ ] **Files:** `src/charts/chartex/histogram.ts`; test `tests/charts-chartex-histogram.test.ts`.
- [ ] **Tests (numeric):**
  - `it('binSize, binCount and automatic Scott bins')`;
  - `it('intervalClosed right and left edges')`;
  - `it('underflow and overflow bins, explicit and auto')`;
  - `it('Office interval label notation')`;
  - `it('pre-aggregated data drawn as given')`;
  - `it('pareto sorts descending with stable ties and a cumulative percent on its own axis')`;
  - `it('over 65536 inputs refuses the chart, never a prefix')`.
- [ ] Commit `feat(charts): histogram and pareto charts`.

### Task B4: Box & whisker

**Spec:** §4.3 box & whisker.

- [ ] **Files:** `src/charts/chartex/boxWhisker.ts`; test `tests/charts-chartex-box.test.ts`.
- [ ] **Tests (numeric):**
  - `it('exclusive and inclusive quartile methods split halves with or without the median, odd and even n')`;
  - `it('whiskers to the furthest point within 1.5 IQR; outliers when visible')`;
  - `it('mean marker, mean line and inner points visibility')`.
- [ ] Commit `feat(charts): box and whisker charts`.

### Task B5: Treemap and sunburst

**Spec:** §4.3 treemap and sunburst; §6.5 node bound.

- [ ] **Files:** `src/charts/chartex/treemap.ts` and `sunburst.ts`; test `tests/charts-chartex-hierarchy.test.ts`.
- [ ] **Tests (numeric and geometry):**
  - `it('nodes identified by full path; duplicates aggregate; missing level is a leaf')`;
  - `it('non-positive weights excluded with a diagnostic')`;
  - `it('squarified treemap areas proportional to weight; parentLabelLayout none, banner and overlapping')`;
  - `it('sunburst rings by depth with angles proportional to subtree weight')`;
  - `it('over 4096 nodes refuses the chart')`.
- [ ] Commit `feat(charts): treemap and sunburst charts`.

### Task C1: Gallery goldens, validation doc, README, CHANGELOG

**Spec:** §8, §9.

**Files:**
- Create:
  - `src/testdata/chart-gallery.ts`, builders for one DOCX, PPTX and XLSX each, with one chart per family and major element;
  - `docs/validation/chart-compat.md`, covering the gate table, B0 → final eager, the reference status per case (verified or unverified) and residual diffs;
  - `tests/golden.test.ts` entries for the gallery (Mac-recorded).
- Modify:
  - `README.md` "Charts" section: supported families, data policy, limits and the `preloadCharts` API;
  - `CHANGELOG.md` `[Unreleased]`.
- Office references: add any owner-supplied files to `validation/office-reference/`; otherwise record the interim LibreOffice evidence and mark the cases unverified.

- [ ] **Step 1:** Add `it('README chart family list matches the engine family registry')`.
- [ ] **Step 2:** Build the gallery, record the goldens on the Mac, and inspect each one individually.
- [ ] **Step 3:** Run the full gates and record them in the validation doc:
  - tsc and typecheck:scripts;
  - full vitest;
  - eager budget;
  - check-eager;
  - golden-corpus compare (baseline-identical);
  - full Playwright, Chromium + WebKit.
- [ ] **Step 4:** Commit `docs: chart compatibility validation, README and CHANGELOG`.
- [ ] **Step 5:** Open the PR `feature/chart-compat` → `main`, closing #13 and superseding #12.
