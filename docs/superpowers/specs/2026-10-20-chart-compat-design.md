# Chart Compatibility — Design Spec

**Status:** draft for review (codex, agy, opencode). Supersedes the plan-only proposal in PR #12 (issue #13).
**Date:** 2026-10-20
**Owner decisions:** none recorded yet. Every choice below that a reviewer could reasonably dispute is marked **[call]**, meaning the author's call, which is open to review.

## 1. Intent

officeview renders Office documents read-only to canvas. Charts are among the most common objects in real DOCX, PPTX and XLSX files, and today they are the weakest:

- XLSX draws six chart families from a rich parser.
- DOCX and PPTX draw only clustered vertical column charts. Every other chart is left **blank**, with only a diagnostic.
- Office 2016+ charts (ChartEx) never render natively anywhere.

**Goal:** every chart a user is likely to meet in an Office file renders natively in all three formats. A chart that still cannot render shows a labelled placeholder card, never empty space. Output must be faithful enough that a reader gets the same information they would get in Office: series, values, axes, labels and legend.

**Success** means:

1. One chart engine serves XLSX, DOCX and PPTX, with no second chart model.
2. Every chart family in §4 renders natively in all three formats.
3. DOCX and PPTX read data the way Office does: chart caches first, then the embedded workbook when the caches are missing.
4. The eager bundle does not grow: it stays ≤ 38,720 B gzip (it is 38,631 B at `3581768`).
5. Unsupported or unreadable charts show a placeholder card and a diagnostic, in every format.

## 2. Non-goals and anti-goals

**Non-goals:**

- Chart interaction: tooltips, selection, drill-down.
- Editing.
- Animation.
- Linked external workbooks (`c:externalData` pointing outside the package).
- Legacy OLE `.bin` embeddings, which would need a compound-file reader.
- True 3D perspective rendering.
- Region-map charts (ChartEx `regionMap`), which need geographic data.

**Anti-goals:**

- Growing the eager bundle.
- A second chart model or painter for DOCX and PPTX.
- Claiming native support where a raster fallback paints. This repeats the rule in `2026-10-03-drawing-format-integration-design.md:15`.
- Unbounded work or memory from a malicious or huge chart or embedded workbook.

## 3. Current state (audited at `3581768`)

### DOCX and PPTX

- Both parse charts through the shared async `prepareDrawingContent` (`src/drawing/content.ts:497`).
- `follow()` loads the chart part and calls the eager mini-reader `chart()` (`content.ts:251-293`). That reader only accepts clustered vertical `c:barChart` and returns `kind:'chart'`, which `paintChart` draws (`content-paint.ts:120-200`).
- Anything else hits `unsupported-content`/`unsupported-chart` (`content.ts:600-606`) and paints nothing.
- The chart part's own `.rels` is never read, so `c:externalData` is never seen. No code opens an embedded package.

### XLSX

- XLSX uses the lazy `src/xlsx/chart.ts`, loaded through `chartLoader.ts` `loadChartModule()`. It supports:
  - bar and column (clustered, stacked and 100%);
  - line, area, pie, doughnut and scatter;
  - 3D variants drawn flat;
  - Cartesian combos on a single shared value scale.
- It has no secondary axis, no data labels, no smoothing, and only circle markers.
- The axis `formatCode` is parsed but unused.
- It rejects radar, bubble, stock, surface and ofPie. Pie combined with any other type becomes a placeholder.

### Shared plumbing that already exists

- `DrawingContent` has `kind:'xlsxChart'` (`content.ts:122`), and `paintDrawingContent` dispatches it to the registered lazy painter, or to `paintChartPlaceholder` (`content-paint.ts:102`).
- The XLSX adapter `contentAdapters(theme).parseChart` (`src/xlsx/drawing.ts:154-163`) awaits the lazy module at parse time.
- DOCX (`docx/drawing.ts:79`) and PPTX (`pptx/parse.ts:902`) call `prepareDrawingContent` with adapters, but not `parseChart`.
- Both paged formats already hold a `ThemeContext` at those call sites.

### Coverage and limits

- Coverage maps `chart` → `column-chart` and `xlsxChart` → `chart`/`chart-placeholder` (`coverage.ts:38-43`).
- DOCX coverage uses the raw kind (`docx/parse.ts:700`).
- Existing limits: 64 series, 4,096 points, 255-unit labels, 4 MiB of label bytes, 160 B per chart cell, label output cap 256, and the 255-unit number-format cap (`MAX_FORMAT_CODE_UNITS`).
- The interactive-viewer spec amendment (`2026-10-10-interactive-viewer.md:30`) governs XLSX chart data fill: Excel numeric-text rules, the value policy, and placeholder-plus-notice instead of an empty chart.

## 4. Scope: chart families

**Classic DrawingML (`c:`).** All of these render natively, in all three formats:

| Family | Today | Target |
|---|---|---|
| Bar and column: clustered, stacked, 100%, horizontal | XLSX ✓ / paged: clustered column only | ✓ all |
| Line: markers, smooth, stacked, 100% | XLSX partial | ✓, with smoothing and marker symbols |
| Area: standard, stacked, 100% | XLSX ✓ | ✓ |
| Pie, exploded pie, doughnut | XLSX ✓ | ✓, with `explosion` |
| Scatter: markers, lines, smooth | XLSX partial | ✓, with smoothing |
| Bubble (incl. `bubble3D` drawn flat) | ✗ | ✓ |
| Radar: standard, markers, filled | ✗ | ✓ |
| Stock: HLC, OHLC, volume-HLC, volume-OHLC | ✗ | ✓ (high-low lines, up/down bars, volume bars on the secondary axis) |
| Pie-of-pie, bar-of-pie (`ofPieChart`) | ✗ | ✓ (split by `splitType`/`splitPos`/`custSplit`) |
| Surface and surface3D | ✗ | **[call]** ✓ as a 2D contour (top view), the form Excel shows for "Contour". Wireframe draws contour lines only. 3D perspective is a non-goal. |
| 3D variants (bar3D, line3D, area3D, pie3D) | drawn flat | drawn flat (unchanged); 3D perspective is a non-goal |
| Combos: any Cartesian mix, including a secondary axis | one shared scale | ✓ with primary and secondary axes |

**Office 2016+ (`cx:`, ChartEx).** **[call]** These render natively from the ChartEx part's own data cache (`cx:chartData/cx:data`). They do this in all three formats:

| Family | Target |
|---|---|
| Waterfall | ✓ (increase, decrease and total bars; connector lines) |
| Histogram and Pareto | ✓ (binning from `cx:binning`, or precomputed; Pareto cumulative line on the secondary axis) |
| Box & whisker | ✓ (quartile method inclusive or exclusive, whiskers, mean marker or line, outliers) |
| Funnel | ✓ |
| Treemap | ✓ (squarified layout; parent labels banner or overlapping) |
| Sunburst | ✓ (rings by hierarchy level) |
| Region map | ✗: stays on the fallback (non-goal) |

If a ChartEx part cannot be read, or its family is region map, the existing `mc:AlternateContent` fallback picture keeps painting. The diagnostic keeps the feature `ChartEx`.

## 5. Scope: chart elements

The engine must honour these, wherever the family uses them:

1. **Axes.**
   - Match by `axId`/`crossAx`, so primary and secondary value and category axes are both drawn.
   - `c:dateAx` scales real dates using `baseTimeUnit` and `majorUnit`/`majorTimeUnit`.
   - Support `logBase`, `orientation=maxMin` (reversed), `crosses`/`crossesAt`, `delete`, `tickLblPos`, axis titles, `majorGridlines`/`minorGridlines`, and `majorTickMark`/`minorTickMark`.
2. **Number formats.**
   - Axis `numFmt@formatCode` and data-label `numFmt` are applied through the existing format engine (`src/xlsx/formula/format.ts`) in the chart's date system, unless `sourceLinked=1`; then the source cell or cache format applies.
   - Codes longer than 255 UTF-16 units render as General (the existing cap).
   - The output is capped at 256 units.
3. **Data labels** (`c:dLbls` and per-point `c:dLbl`):
   - content: `showVal`, `showCatName`, `showSerName`, `showPercent`, `showBubbleSize`, `separator`;
   - placement: `dLblPos` (center, inEnd, inBase, outEnd, l, r, t, b, bestFit);
   - per-point delete.
   - **[call]** Manual label layout (`c:layout` inside `dLbl`) is ignored, and the label is placed at its `dLblPos`.
4. **Markers.**
   - Symbols: circle, square, diamond, triangle, x, star, dash, dot, plus, none.
   - Marker size, plus marker fill and line from `spPr`.
5. **Lines.**
   - `c:smooth` uses Catmull-Rom, tension matching Excel visually.
   - Line dash (`prstDash`), width and `noFill`.
6. **Fills.**
   - Solid, gradient (linear and path), pattern and `noFill`, for series, points, plot area and chart area.
   - The chart area background respects `chartSpace/spPr`, including `noFill`. Today the painter always paints white (`chart.ts:637-638`); DOCX and PPTX charts often have no fill.
7. **Trendlines:** linear, exponential, logarithmic, polynomial (order 2–6), power and moving average, with `forward`/`backward`, `intercept` and `dispEq`/`dispRSqr` labels.
8. **Error bars:** fixed value, percentage, standard deviation, standard error and custom (`plus`/`minus` caches), in x and y, with the `noEndCap` style.
9. **Series-level options:**
   - `dispBlanksAs` (gap, zero, span) and `invertIfNegative`;
   - `explosion` (pie);
   - `gapWidth`/`overlap` (already supported);
   - `varyColors` (already supported);
   - up/down bars, high-low lines and drop lines (line and stock).
10. **Categories.** Multi-level categories (`multiLvlStrRef`, all `lvl`s) are drawn as grouped category tiers. Today only the first level is read.
11. **Text.**
    - Title and legend use `overlay`; the legend position supports `legendPos` r/l/t/b/tr.
    - Legend entries hidden by `c:legendEntry/c:delete` are hidden.
    - Rich-text title runs use the rich-text font, size, bold, italic and colour from `c:txPr`/`a:rPr`.
    - **[call]** Manual layout (`c:manualLayout`) for the plot area, title and legend is honoured. Its `x`/`y`/`w`/`h` are fractions of the chart area; `layoutTarget` is inner vs outer.
12. **Theme and style.**
    - Series colours follow `c:style`/the chart style part and the colour part (`chartStyle`/`colors` relationships) when present.
    - Otherwise they follow the accent cycle with lumMod/lumOff variants after six series, as Office does.
13. **User shapes** (`c:userShapes`, drawing overlays on a chart): **[call]** these render through the existing drawing scene painter, positioned by `relSizeAnchor`/`absSizeAnchor`.

## 6. Architecture

### 6.1 One engine, lazily loaded

- The chart engine lives in lazy modules under `src/charts/`: parser, layout and painter. The existing `src/xlsx/chart.ts` moves into it, split by family into files of reasonable size. `chartModel.ts` stays types-only.
- `loadChartModule()` becomes the single lazy entry, shared by all formats.
- Nothing in `src/charts/` may be statically reachable from `src/index.ts`. `EAGER_FORBIDDEN` (`tests/helpers/eager-graph.ts`) gains `src/charts/` and `src/xlsx/chart.ts` (while it exists), so `check-eager` enforces this instead of relying only on the gzip budget.

### 6.2 DOCX and PPTX use the same path as XLSX

- DOCX `loadDrawingParts` and PPTX `parsePptx` pass a `parseChart` adapter, as XLSX does. That adapter awaits `loadChartModule()` and produces `kind:'xlsxChart'`. **[call]** The kind is renamed to `kind:'chartModel'`, with a type alias kept for one release.
- Classic `c:` charts and ChartEx `cx:` charts both go through the adapter.
- The eager mini-reader `chart()` and `paintChart` are **deleted**. This is a net eager saving that the budget table must show.
- If the lazy module cannot load (chunk failure), the content becomes a placeholder model with the reason "Chart could not load". It never becomes blank, and it never falls back to a second parser.

### 6.3 Paint-time safety

- `getPaintables` DOCX and PPTX branches (`src/render/paint.ts:231, 281`) gain an `ensureChartPainter` equivalent before paint. It is typed over `DrawingContent`, not `XlsxDrawing`.
- The public synchronous `renderPages`/`renderSlide` keep working: parse already loaded the module, and loading it registers the painter.
- A missing painter still shows the "Chart not loaded" card.

### 6.4 Embedded-workbook data (DOCX and PPTX)

- Office stores a paged chart's data in an embedded workbook: chart part `.rels` → `c:externalData r:id` → `../embeddings/Microsoft_Excel_Worksheet*.xlsx`.
- **Data order:**
  1. chart caches (`numCache`/`strCache`, `numLit`/`strLit`);
  2. if a referenced range has **no cache**, the embedded workbook;
  3. if there is neither, a placeholder with the reason "This chart's data is not available".
- Caches win even when the workbook disagrees, because that is what Office displays until the user edits.
- **Reading the workbook.**
  - Read the chart part's `.rels`. Resolve the target inside the package only; `TargetMode="External"` is a non-goal, reported as a diagnostic.
  - Bounded read with `pkg.bytes(path, EMBEDDED_WORKBOOK_MAX_BYTES)`, **[call]** 16 MiB.
  - Open it with `OfficePackage.load`, then run the existing resident XLSX parser in a **values-only** mode: no drawings, no styles beyond number formats, and no formula evaluation beyond cached values.
  - Fill through the existing `fillChartRanges` and `scanChartNumber` rules, with the same limits (64 series, 4,096 points, 160 B per cell, label caps).
  - Parse each embedded workbook at most once per document and cache it per package path. Release it after all charts are filled.
- **Failure** of any kind (missing, external, too large, corrupt zip, malformed XML, budget refusal): the chart keeps its caches if it has them, otherwise it becomes a placeholder. Each failure records a diagnostic: `chart-data-unavailable` with a reason.

### 6.5 Coverage, diagnostics and accessibility

- Coverage: a native chart → `chart`, and a placeholder → `chart-placeholder`, in all three formats.
- DOCX coverage maps the kind through `featureOfKind`, not the raw kind.
- ChartEx native → feature `ChartEx`, representation `chart`. ChartEx on fallback → representation `fallback`, unchanged.
- The `unsupported-chart` diagnostic is kept for families still out of scope (region map) and for unreadable parts.
- Accessibility: DOCX and PPTX charts expose the `chartDescription` text (type, title, series count, value range) the same way XLSX sheet charts do, through the paged viewer's existing accessible-text path for drawings. **[call]** Where none exists, add one `aria-label` per chart canvas region.

### 6.6 Rendering rules

- **Paper rule:** charts are content, so they are never themed by the viewer's dark mode (owner decision B in the restyle spec). Chart colours come only from the file and its theme.
- **Fonts:** chart text uses the document theme's minor font through the existing font manager, the same as other drawing text, sized from `c:txPr` (6–72 pt clamp).
- **Determinism:** layout must be deterministic for a given input and font metrics, so goldens are stable.

## 7. Limits and robustness

These apply per chart, unless stated otherwise:

| Item | Limit | On breach |
|---|---|---|
| Series | 64 (existing) | truncate; `truncated` flag and diagnostic |
| Points per series | 4,096 (existing) | truncate |
| Label length | 255 units (existing) | truncate, marked |
| Label bytes per chart | 4 MiB (existing) | truncate |
| Chart part size | 8 MiB declared / 4 MiB actual (existing large-source cap) | placeholder |
| Embedded workbook | 16 MiB compressed **[call]**; parsed values bounded by the existing XLSX resident budgets | caches or placeholder |
| ChartEx data points | 4,096 per series; hierarchy depth 8 **[call]** | truncate, marked |
| Treemap/sunburst nodes | 4,096 **[call]** | truncate, marked |
| Polynomial trendline order | 2–6 | clamp |
| Histogram bins | 1–1,000 **[call]** | clamp |

- Every family must handle empty series, all-blank values, NaN/∞, negative values on log axes (skip point and record a diagnostic) and a single point, without throwing.
- Malformed XML or an unknown element is skipped (forward compatibility).
- Painting never throws; a painter error becomes the placeholder.

## 8. Testing and validation

- **Unit tests per family and element:** parse to model, then a recording-context paint. Use synthetic `c:`/`cx:` XML builders.
- **Shared goldens:** a synthetic chart gallery, one file per format with one chart per family, added to `tests/golden.test.ts` and recorded on the Mac.
- **Native references:**
  - Add real Office-authored DOCX, PPTX and XLSX files containing every family to `validation/office-reference/`, with Office screenshots.
  - Compare with per-unit justification for residual diffs.
  - **[call]** The owner supplies the Office-authored files and screenshots. If they are unavailable, files built with python-docx/python-pptx/openpyxl plus LibreOffice references are accepted as interim, labelled as such.
- **Corpus:** add chart-bearing files to the corpus lock if public sources exist; otherwise record the absence.
- **Eager and budget:** `check-eager` with `src/charts/` forbidden; the eager budget ≤ 38,720 B; a before/after table.
- **Robustness:** a torture file per format with every family, a ChartEx part, model3D, OMML, a huge embedded workbook, a corrupt embedded zip, external `externalData`, and empty or NaN data. It renders a valid PNG, records the expected diagnostics, and throws zero times.
- **Browser:** the paged viewer shows DOCX and PPTX charts; dark mode leaves chart pixels identical to light (paper rule); chart canvases have an accessible name.

## 9. Delivery constraints

- **Eager:** net eager change ≤ 0 per task. Deleting `chart()`/`paintChart` should free bytes; the plan records the numbers.
- **Goldens:** no golden is re-recorded without individual inspection and a stated reason. Existing XLSX chart goldens must be pixel-identical, except where a §5 element changes the picture (for example, `noFill` chart areas); each such change is listed.
- **Review:** every task is reviewed by a model family other than its implementer's.
- **CI:** CI must pass on Linux; the pixel goldens stay Mac-only per `ci.yml`.

## 10. Open questions for reviewers

1. Is surface-as-contour acceptable, or should surface stay a placeholder?
2. Is the ChartEx scope (all except region map) right for this project, or should ChartEx become a follow-up?
3. Is the 16 MiB embedded-workbook cap, with values-only parse, sound against the existing resident budgets?
4. Are trendlines and error bars in or out?
5. Is the `xlsxChart` → `chartModel` rename worth the churn?
