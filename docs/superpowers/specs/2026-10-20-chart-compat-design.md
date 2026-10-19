# Chart Compatibility — Design Spec

**Status:** revision 2. This addresses review round 1: codex, agy and opencode all returned REQUEST CHANGES, with no Critical findings. It supersedes the plan-only proposal in PR #12 (issue #13).
**Date:** 2026-10-20 · **Baseline:** `main` @ `3581768`.
**Decision marks:** **[call]** marks the author's call, open to review. **[rev]** marks a decision taken from a round-1 review finding. The reviewer and finding are given in brackets.

## 1. Intent

officeview renders Office documents read-only to canvas. Charts are common in real DOCX, PPTX and XLSX files, and they are the weakest drawing type:

- XLSX draws six classic families.
- DOCX and PPTX draw only clustered vertical column charts. Anything else is absent, unless a compatibility fallback picture happens to exist.
- Office 2016+ charts (ChartEx) never render natively.

**Goal:** every chart a user is likely to meet renders natively, with the same information Office shows (series, values, axes, labels, legend), in all three formats. A chart that cannot render shows its authored fallback picture if one exists, otherwise a labelled placeholder card. It never shows empty space and never shows plausible but wrong values.

**Success** means:

1. One lazily loaded chart engine serves XLSX, DOCX and PPTX.
2. Every family in §4 renders natively in all three formats.
3. DOCX and PPTX read chart data as this viewer's saved-state policy defines (§6.4): caches first, then the embedded workbook, under hard bounds.
4. The eager bundle does not grow. It stays ≤ 38,720 B gzip, measured as in §9.
5. A chart that cannot render follows the representation contract in §6.3: fallback picture, else card, plus a diagnostic.
6. Native-support claims are gated on reference evidence (§8). Cases without Office references are labelled *unverified* in the validation doc.

## 2. Non-goals and anti-goals

**Non-goals:**

- Chart interaction and editing.
- Animation.
- Linked external workbooks (`TargetMode="External"`).
- Legacy OLE `.bin` embeddings.
- Formula evaluation inside embedded workbooks.
- True 3D perspective.
- ChartEx region map.

**Anti-goals:**

- Growing the eager bundle.
- A second chart model.
- Claiming native support where a raster fallback paints (`2026-10-03-drawing-format-integration-design.md:15`).
- Unbounded work or memory from any chart, embedded workbook or derived computation.
- Silently changing a chart's meaning: no truncating the inputs of statistics, clamping authored parameters, or guessing axes.

## 3. Current state (audited at `3581768`; confirmed by all three reviewers)

### DOCX and PPTX

- They enter the shared async `prepareDrawingContent` (`src/drawing/content.ts:497`):
  - DOCX goes through `prepareCompatibleDrawingContent` (`src/docx/drawing.ts:79`), which owns `mc:AlternateContent` selection;
  - PPTX calls `prepareDrawingContent`, then the compatible wrapper on failure (`src/pptx/parse.ts:902-905`).
- `follow()` reads the chart part unbounded (`pkg.xmlOrdered`, `content.ts:515`) and calls the eager mini-reader `chart()` (`content.ts:251-293`).
- That reader finds a `c:barChart` group. It accepts only a clustered vertical group, and in a mixed plot it can draw just the bar portion. `paintChart` (`src/drawing/content-paint.ts:120-200`) draws it.
- Other charts produce `unsupported-content`/`unsupported-chart` (`content.ts:600-606`) and no content. An authorized compatibility fallback can still paint.
- ChartEx, model3D and OMML are rejected unconditionally (`content.ts:869`) and structurally (`content.ts:999-1045`). Consumers keep per-consumer Choice capability lists (`docx/drawing.ts:102`).
- The chart part's `.rels` is never read for charts, so `c:externalData` is unseen. No code opens an embedded package.

### XLSX

- XLSX uses the lazy `src/xlsx/chart.ts` (via `chartLoader.ts:6-9`), which supports:
  - bar and column (clustered, stacked and 100%);
  - line, area, pie, doughnut and scatter;
  - 3D variants drawn flat;
  - Cartesian combos on one shared value scale.
- It picks one category axis and one value axis, with no axId graph (`chart.ts:242-243`).
- It has no secondary axis, data labels or smoothing, and markers are circles only.
- Axis `formatCode` is parsed but unused (`chart.ts:135-144`), and only the first category level is read (`chart.ts:83`).
- It rejects radar, bubble, stock, surface and ofPie, and pie mixed with other types (`chart.ts:34-37, 233`).
- The background is always white (`chart.ts:637-638`).
- The registered painter is not wrapped in a try/catch (`content-paint.ts:101-102`).

### Shared plumbing

- `kind:'xlsxChart'` is already in `DrawingContent` (`content.ts:122`) and dispatched by `paintDrawingContent`.
- `featureOfKind`/`contentRepresentation` live in `src/drawing/coverage.ts:38-43`. `featureOfKind` is private.
- DOCX coverage uses the raw kind (`src/docx/parse.ts:700`).

### Limits

- Existing limits: 64 series, 4,096 points, 255-unit labels, 4 MiB label bytes and 160 B per chart cell (`chart.ts:18-23, 756`); the number-format cap is 255 units (`src/xlsx/parse.ts:86`).
- The drawing-part limit is a **desktop/mobile** pair (`src/xlsx/storage/importSession.ts:60-72, 100-108`; `admission.ts:87, 186`). It is not a "declared/actual" pair. [rev: codex audit]
- `OfficeZip.bytes(path, max)` bounds **inflated** bytes (`src/core/zip.ts:40-63`).
- There is no values-only XLSX parse mode. `ParseXlsxOptions` has only evaluation caps and `valuePolicy` (`parse.ts:166-172`).

## 4. Scope: chart families

### 4.1 Classic DrawingML (`c:`): stream A

| Family | Target | Semantic rules (normative) |
|---|---|---|
| Bar and column (clustered, stacked, 100%, horizontal) | native | `barDir`, `grouping`, `gapWidth`, `overlap`, `invertIfNegative` |
| Line (markers, smooth, stacked, 100%) | native | §5.5 smoothing; stacked lines accumulate per category |
| Area (standard, stacked, 100%) | native | |
| Pie, exploded pie, doughnut | native | `firstSliceAng`, `holeSize`, per-point and series `explosion` |
| Scatter (markers, lines, smooth) | native | independent x domain (§5.1) |
| Bubble (`bubble3D` drawn flat) | native | `bubbleScale` (default 100%). `sizeRepresents` = area (default) or w. Size ≤ 0 is not drawn and is diagnosed; it never draws as zero or negative. A series without `bubbleSize` data is unavailable. |
| Radar (standard, marker, filled) | native | `radarStyle`; categories clockwise from 12 o'clock; the value axis is radial |
| Stock (HLC, OHLC, volume-HLC, volume-OHLC) | native | Series roles are ordered: [volume], open?, high, low, close. Roles come from series order in the stock group; volume is the bar group in the same plot. **Axes come from each group's authored `axId` bindings, never assumed** [rev: codex I10]. Draw `hiLowLines` and `upDownBars` (up/down fills from `c:upBars`/`c:downBars`). A series-count/role mismatch becomes a card with a reason. |
| Pie-of-pie, bar-of-pie | native | `ofPieType`. `splitType` is one of: auto (Office default: last third of points, `ceil(n/3)`), pos (last `splitPos` points), val (points < `splitPos`), percent (points < `splitPos`% of total) or cust (`custSplit/secondPiePt` indices). Original point indices keep their colours and labels. The primary pie gains one synthetic "Other" slice whose value is the sum of the secondary points, so totals and percentages are conserved. Also `secondPieSize` (default 75%), bar-of-pie `gapWidth`, and `serLines` connectors. [rev: codex I10, opencode F6] |
| Contour (`surfaceChart`, the 2D surface) | native | z grid = series (rows) × categories (columns). Colour bands come from the value-axis major unit (Office band count). Missing cells are not interpolated across. `wireframe=1` draws band boundaries only. Band colours follow `c:bandFmts` when present, otherwise the accent cycle. |
| Surface 3D (`surface3DChart`) | **fallback/card** | Perspective is a non-goal. Show the authored fallback, else a card reading "3-D surface charts are shown as a placeholder". [rev: codex I11 / §10 Q1] |
| 3D bar/line/area/pie | drawn flat (unchanged) | Existing behaviour, kept and documented as an approximation; this is no new claim |
| Combos (any Cartesian mix) | native | Axes are modelled by ID (§5.1). A plot that mixes category groups with XY groups keeps independent x domains. An axis combination the model cannot place becomes a card with a diagnostic, never a guess [rev: codex I5]. Pie mixed with non-pie stays a card (Office cannot author it). |

### 4.2 Office 2016+ (`cx:`, ChartEx): stream B

**[call; rev: all reviewers — staged]** Stream B is part of this project, but it is a separately reviewed stream that starts after stream A's engine and adapter land. Its first task is a **normative field map** (§4.3), derived from real ChartEx parts. No ChartEx family is claimed native until its rules and numeric oracles pass.

| Family | Target |
|---|---|
| Waterfall | native |
| Funnel | native |
| Histogram, Pareto | native |
| Box & whisker | native |
| Treemap, sunburst | native |
| Region map | fallback picture, else card |

A ChartEx part that is unreadable, region map, or outside the limits **returns "unsupported"** to the representation contract (§6.3). The authored `mc:Fallback` picture is then selected; the card appears only if no fallback exists.

### 4.3 ChartEx normative rules (completed by stream B task 1; minimum content fixed here) [rev: codex I6, opencode F3]

The stream-B field-map task must produce `docs/charts/chartex-map.md`. Its element names, attributes and defaults are confirmed against MS-ODRAWXML and against at least one Office-authored part per family. It must state:

- **Data model.** `cx:chartData/cx:data[@id]` holds ID-keyed dimensions (`cx:numDim`/`cx:strDim`, each with a `@type`, levels `cx:lvl` and `ptCount`/`pt@idx`). A series binds dimensions through `cx:dataId`. Values are never assumed to be classic series arrays.
- **Layout selection.** Use `cx:series@layoutId` (clusteredColumn, paretoLine, waterfall, funnel, boxWhisker, treemap, sunburst, regionMap). Histogram is clusteredColumn with `cx:binning`. Pareto is the histogram plus a paretoLine series.
- **Waterfall.** Subtotal points come from `cx:subtotals/cx:idx`. A subtotal bar shows the running total, anchored at zero, and the accumulator continues from it. Increase, decrease and total colours come from data-point overrides, otherwise the defaults. Connector lines follow `cx:layoutPr/cx:visibility@connectorLines`.
- **Histogram.** Bins are one of `binSize`, `binCount` or automatic. Automatic follows Scott's rule, which is what Office documents for "Automatic". `intervalClosed` is r (default) or l. `underflow`/`overflow` set the bin edges (an explicit value or "auto"). Bin labels use the Office interval notation. Pre-aggregated (category) data is drawn as given.
- **Pareto.** Bars are sorted descending by value; ties keep source order. The cumulative line is the running sum ÷ total, plotted on its own secondary percent axis (0–100%).
- **Box & whisker.**
  - `quartileMethod` (exclusive, the default, or inclusive) defines whether the **median is included when splitting halves** (ST_QuartileMethod). It is not QUARTILE.INC/EXC interpolation.
  - Whiskers reach the furthest point within 1.5 × IQR. Points beyond are outliers, drawn only if `cx:visibility@outliers`.
  - Mean marker and mean line follow `meanMarker`/`meanLine`; inner points follow `nonoutliers`.
- **Treemap and sunburst.**
  - Each node is identified by its **full category path** (all levels). Duplicate paths are aggregated, and a missing level is treated as a leaf at the previous level.
  - Weights ≤ 0 or non-numeric are excluded and diagnosed.
  - Treemap uses squarified layout with `parentLabelLayout` (none, banner or overlapping). Sunburst draws rings by depth, with angles proportional to subtree weight.
- **Funnel.** Bars are centred, width proportional to value, in source order.

**Analytic inputs are never truncated** [rev: codex I7]. These families compute statistics over all points, using single-pass or bounded algorithms. If the inputs exceed the analytic limit in §7, the chart is refused: its fallback is selected, or a card reading "This chart has too many data points to summarize" appears. A prefix is never summarized.

## 5. Scope: chart elements (both streams where applicable)

### 5.1 Axes [rev: codex I5, agy 4, opencode F1]

- **Model.**
  - Axes are keyed by `axId`; each group binds its `axId` pair.
  - `crossAx` links an axis to the axis it crosses. `axPos` (b/t/l/r) places it, and `crosses` (autoZero, min, max) or `crossesAt` sets where it crosses.
  - Primary and secondary axes both render.
  - Scatter and bubble use two value axes, so x and y domains are independent.
- **Deleted axes.** A deleted axis (`delete=1`) still defines the scale; it is just not drawn.
- **Scale and ticks.**
  - `crossBetween` (between, the bar default; midCat, the line/area default) controls whether categories sit between ticks or on them.
  - `orientation=maxMin` reverses the axis.
  - `logBase` (2–1000) makes a log scale. Values ≤ 0 on a log axis are skipped, with the diagnostic `chart-invalid-log-value` [rev: agy 1].
  - Also: `min`/`max`, `majorUnit`/`minorUnit`, `majorGridlines`/`minorGridlines`, `majorTickMark`/`minorTickMark`, and `tickLblPos` (nextTo, low, high, none).
  - `tickLblSkip`/`tickMarkSkip` set label and tick thinning. When they are absent, Office's automatic skip applies, so labels never overlap.
  - Label rotation comes from `txPr/bodyPr@rot` and `vert`.
- **Display units.** `dispUnits`, either `builtInUnit` (hundreds … trillions) or `custUnit`, divides tick values. It draws the unit label (`dispUnitsLbl`, with the default unit text) beside the axis. Without this, finance charts show values off by factors of 10³–10⁹.
- **Date axes.** `c:dateAx` scales numeric date serials. `baseTimeUnit` and `majorTimeUnit` are days, months or years; **month and year steps follow the calendar, never a constant number of days**. Unequal date spacing is preserved. If the category data is text, the axis is treated as categorical (`auto=1` behaviour) and diagnosed. A date axis is never rebuilt from formatted labels.
- **Date system [rev: codex I4, agy 5].** The first of these that applies wins:
  1. `c:chartSpace/c:date1904`;
  2. the embedded workbook's `workbookPr@date1904` (DOCX/PPTX) or the containing workbook's (XLSX);
  3. the 1900 system.

  The 1900 system keeps serial 60 as 1900-02-29, as Excel does. Time fractions are preserved.

### 5.2 Number formats [rev: codex I4]

- A **chart display-format adapter** in the lazy engine formats axis labels, data labels and percentages.
  - It reuses the renderer's display layer, `src/xlsx/cellFormat.ts`, which handles scientific, fraction, elapsed and locale formats.
  - It does **not** use the formula TEXT grammar in `formula/format.ts`, which diagnoses those formats.
  - Codes over 255 units → General. Output is capped at 256 units.
- **Precedence:**
  1. explicit `numFmt` with `sourceLinked=0`;
  2. `sourceLinked=1` → the source format (the cache `formatCode` per point, or the embedded cell's format);
  3. no format → General, or percent for 100% axes.
- Mixed per-point source formats are applied per point.
- A raw numeric value and its format provenance are kept on the model; labels are formatted at paint time.

### 5.3 Data labels [rev: codex I10]

- **Precedence:** chart-group `dLbls` → series `dLbls` → point `dLbl`.
  - An explicit `0`/`false` overrides an inherited `1`.
  - `dLbl/delete` hides that label.
- **Content and placement:**
  - `showVal`, `showCatName`, `showSerName`, `showPercent`, `showBubbleSize`, `showLegendKey` and `separator`.
  - Authored rich text (`dLbl/tx/rich`) replaces generated text.
  - `dLblPos` per family (with Office defaults); `bestFit` is the pie default.
  - Leader lines (`showLeaderLines`) are drawn for pie labels moved outside.
- **Out of scope:** field-reference labels (`c15:` "value from cells" ext) and manual label positions (`dLbl/layout`). Each draws the label at its `dLblPos` with the diagnostic `chart-label-feature-unsupported`.

### 5.4 Markers, lines and fills

- **Markers.**
  - Symbols: circle, square, diamond, triangle, x, star, dash, dot, plus, none, auto (series cycle).
  - `size` is clamped to 2–72 per the schema.
  - Fill and line come from the marker's `spPr`.
- **Lines.** `prstDash`, width, `noFill`, and `cap`.
- **Fills.** Solid, gradient (linear and path), pattern (existing pattern renderer) and `noFill`, for series, points, plot area, chart area, legend and title.
- **Chart area.** The chart area honours `chartSpace/spPr`. A missing fill uses Office's default per producer: white for charts created in Excel. For DOCX and PPTX the default comes from `c:style`, else white. `noFill` is transparent.
- **Paper rule.** The viewer's dark mode never changes chart pixels (restyle decision B).

### 5.5 Smoothing [rev: codex I11]

**[call]** `c:smooth=1` draws a cubic Bézier through every point. Tangents come from neighbouring points (Catmull-Rom, tension ½); the end segments use the end point as its own neighbour.

- Missing points break the curve, following `dispBlanksAs`.
- This is an **approximation** of Office smoothing. It is labelled as such in the validation doc and compared against references for endpoints, unequal spacing and overshoot before being called matched.

### 5.6 Trendlines [rev: codex I7, agy 1]

- **Types:** linear, exponential, logarithmic, polynomial (order 2–6), power and moving average (period 2 to n−1).
- **Fitting:**
  - Fits use the **x values for XY charts** and 1..n category indices for category charts.
  - Least squares uses a numerically stable solve (QR, or normal equations with scaling). A singular or ill-conditioned fit draws no line and records `chart-singular-trendline`.
  - Exponential and power fits drop points with y ≤ 0. Logarithmic and power fits drop points with x ≤ 0. Each drops with a diagnostic.
  - `intercept` applies to linear, exponential and polynomial fits.
  - `forward`/`backward` extend the line in x units. The extension is bounded so the drawn domain is ≤ 10 × the data domain, otherwise the extension is clamped and diagnosed.
- **Out-of-range authored parameters** (order outside 2–6, a period outside its range) → no line and a diagnostic. They are never clamped silently.
- **Labels.** The equation and R² label use Office's formats: `y = 1.2345x + 0.5` and `R² = 0.9876`, with 4 significant digits.
- **Sampling.** The line is drawn by sampling ≤ 512 points.

### 5.7 Error bars [rev: codex I7]

- `errDir` (x/y; y only for non-XY), `errBarType` (both, plus, minus) and `noEndCap`.
- `errValType`:
  - `fixedVal` and `percentage` from `val`;
  - `stdDev` (n−1 sample standard deviation × `val`, centred on the series mean, as Office does);
  - `stdErr`;
  - `cust` from `c:plus`/`c:minus` data sources, which follow §6.4 data precedence like any series data.
- Missing custom values for a point → no bar for that point.

### 5.8 Chart-level and series options [rev: agy 4, opencode F2, codex I10]

- **Chart level.**
  - `c:dispBlanksAs` (gap, zero, span; Office default is gap) applies to line, scatter and radar. Bars treat blanks as gaps.
  - `c:plotVisOnly`.
  - Error cells (#N/A …) are treated as blanks for plotting. Line charts treat #N/A as span when `dispBlanksAs=gap`, matching Office.
- **Series level.** `invertIfNegative`, `explosion`, and per-point `dPt` overrides.
- **Group level.** `varyColors`, `gapWidth`, `overlap`, `upDownBars`, `hiLowLines`, `dropLines` and `serLines`.

### 5.9 Categories

Multi-level categories (`multiLvlStrRef`/`multiLvlStrCache`, all `lvl`s) are drawn as grouped tiers. A group label spans its children. `noMultiLvlLbl=1` shows only the innermost level.

### 5.10 Text, layout and legend [rev: codex I5, agy 6]

- **Title.** Content comes from `title/tx/rich` (authored runs), else `title/tx/strRef` (cache), else the automatic title (single-series name or "Chart Title"), unless `autoTitleDeleted`. Run formatting precedence: run `a:rPr` → paragraph default → `c:txPr` → style part → theme.
- **Legend.** `legendPos` (r, l, t, b, tr), `overlay`, and `legendEntry/delete`.
- **Manual layout.** `c:manualLayout` for the plot area, title and legend honours `ST_LayoutMode`:
  - `xMode`/`yMode` = edge: `x`/`y` is a fraction of the chart area.
  - `xMode`/`yMode` = factor: an offset from the element's default position, as a fraction of the chart area.
  - `wMode`/`hMode` = factor: `w`/`h` is the size as a fraction. `wMode`/`hMode` = edge: `w`/`h` is the right or bottom edge, so width = w − x.
  - `layoutTarget` = inner (the plot area excluding tick labels) or outer.
  - The result is clamped to the chart area. An empty or inverted box falls back to automatic layout, with a diagnostic.

### 5.11 Style, colour and user shapes [rev: codex minor 3]

- **Colour precedence:** explicit `spPr` → chart style part (`cs:chartStyle`) → colour part (`cs:colorStyle`) → `c:style` index → theme accent cycle, with lumMod/lumOff variants after six series, as Office does.
- **User shapes.** `c:userShapes` (relSizeAnchor/absSizeAnchor) are drawn with the existing drawing scene painter. They go through the existing relationship and depth budget (cycle-safe). Images inside them use the existing image budget.

## 6. Architecture

### 6.1 One lazy engine

- **Location.** The chart engine lives under `src/charts/`: parse, model, data, layout, format, paint, and families split one per file. `src/xlsx/chart.ts` moves there.
- **Model types.** `chartModel.ts` stays types-only.
- **Single entry.** `loadChartModule()` is the single lazy entry for all formats.
- **Eager guard.** `EAGER_FORBIDDEN` gains `src/charts/` and `src/xlsx/chart.ts`, so `check-eager` enforces the boundary.
- **Placeholder stays eager.** The placeholder painter (`paintChartPlaceholder`) is small and stays eager, so a chunk failure can still paint a card. [rev: codex]

### 6.2 Format adapters and the runtime tag

- One `parseChart` adapter is added to the shared `ContentLoadAdapters` defaults, so XLSX, DOCX (through the compatible wrapper) and PPTX all receive it. [rev: opencode F8]
- The dispatch in `follow()` recognises both classic `c:chartSpace` (with `c:chart`) and ChartEx `cx:chartSpace` roots by namespace, including alias prefixes. [rev: agy 3]
- The unconditional ChartEx rejections (`content.ts:869`, the structural check at `999-1045`) and each consumer's Choice capability list (`docx/drawing.ts:102`, the PPTX equivalents) are updated so that `cx` Choices are attempted when the engine can load.
- **Runtime tag.** **[rev: codex minor 1]** The runtime tag stays `kind:'xlsxChart'`. Renaming a runtime discriminant breaks transferred models and every switch. Module and type names become format-neutral (`ChartModel`); the tag string is unchanged. The eager mini-reader `chart()`, its `kind:'chart'` and `paintChart` are deleted.

### 6.3 Representation contract (MC, fallback, card) [rev: codex I3, agy 3]

- The adapter returns one of three results:
  - `native(model)`: the chart renders.
  - `unsupported(reason)`: unreadable, out of scope (region map, surface3D), over the limits, or chunk failure.
  - `refused(reason)`: a data budget was exceeded.
- `prepareCompatibleDrawingContent` treats `unsupported` and `refused` as **not satisfying** the Choice, so the `mc:Fallback` picture is selected.
- Only when **no** usable representation remains does the content become a placeholder card model. The card carries the reason, the original feature (`chart`/`ChartEx`) and exactly one selected representation.
- A standalone part (no AlternateContent) that is unsupported becomes a card directly.
- Nested and outer AlternateContent, alias prefixes and unknown `Requires` are pinned by tests in all three formats.

### 6.4 Data precedence and the embedded workbook (DOCX/PPTX) [rev: codex I1, I8; agy 2; opencode F4, F5]

**Saved-state policy.** This viewer shows a chart's **saved** values. It does not claim to reproduce every Office refresh behaviour.

**Per data role.** The roles are series name, categories, values, x values, bubble sizes and custom error plus/minus.

1. **Cache present:** the cache element exists. Use it, even when it is **partial**:
   - sparse `pt@idx` positions are kept;
   - missing indices are **blank**, never compacted and never zero;
   - `ptCount=0` is a present, empty cache.
2. **Cache absent, reference present:** read the referenced range from the embedded workbook.
3. **Neither:** the role is *unavailable*. This uses an explicit sentinel, distinct from blank.

**Required roles.**

- Values are required for every family. X values are required for scatter and bubble, and bubble sizes for bubble.
- Categories are optional; they default to 1..n.
- If any required role of any series is unavailable, the whole chart is `unsupported("This chart's data is not available")`. A chart is never drawn with a series silently missing.
- Mixing cached and workbook-filled roles is allowed.

**Embedded workbook reader.** This is a **new bounded values-only reader** in a lazy chunk: `src/charts/embeddedWorkbook.ts`. It is not an unrestricted `parseXlsx` call.

- **Resolve.** Resolve chart part `.rels` → `c:externalData@r:id`. Only a package-internal target is accepted; `External` gives a diagnostic and the data is unavailable.
- **Admit.** Admission condition: inflated size ≤ **8 MiB** (desktop) / **4 MiB** (mobile) via `pkg.bytes(path, max)` **[rev: agy 2; call]**. Real chart workbooks are 15–100 KB.
- **Read.** `OfficePackage.load` the bytes, then read only these:
  1. `xl/workbook.xml`: sheets, `date1904` and `definedNames`. External references are rejected even when reached through a name.
  2. `xl/sharedStrings.xml`, bounded by streaming SAX with a string-count and byte cap.
  3. `xl/styles.xml` `numFmts`/`cellXfs` only.
  4. **Only the sheets the chart's references name.** They are read with the existing streaming worksheet tokenizer, extracting cells **only inside the referenced ranges** and stopping after the last referenced row.
- **Never followed:** drawings, nested embeddings, external links, VBA or any other relationship.
- **Formulas.** Formula cells use their saved value (`<v>`). A formula without a saved value is unavailable for that point and diagnosed. Formulas are never evaluated.
- **Values.** Excel numeric-text rules apply via `scanChartNumber`. Error cells are treated as blanks (§5.8).
- **Budgets:**

  | Budget | Desktop / mobile |
  |---|---|
  | Inflated XML read per embedded package | 32 / 16 MiB |
  | Shared-string bytes | 8 / 4 MiB |
  | Cells extracted per chart | 262,144 (64 × 4,096) |
  | Cell bytes | 160 B each (existing) |
  | Retained embedded models per document | 32 MiB |

  Every budget is charged before allocation. A breach → `refused`, a diagnostic, and every partial allocation released.
- **Reuse and cancellation.**
  - Requests are deduplicated per normalized part path as a shared promise, then discarded once every consumer is filled.
  - Cancellation through the document `AbortSignal` releases all charges.
- **Tests:**
  - a sub-cap ZIP with an oversized inner sharedStrings or sheet;
  - 50 distinct embeddings;
  - repeated references to one embedding;
  - cancellation mid-read;
  - an external target;
  - a defined name pointing at an external workbook;
  - a formula without a saved value;
  - a corrupt ZIP.

**XLSX** keeps its existing fill path (`fillWorkbookCharts`, paged `fillCharts`) with the same per-role precedence, extended for the new roles (bubble size, error bars).

### 6.5 Paint and preparation [rev: codex I2, minor 2; agy 7]

- **Bounded chart XML.** The chart part is read with a byte cap before the tree is built: desktop/mobile 8/4 MiB, the same pair as the existing drawing-part limit, in all formats. Over the cap → `unsupported`.
- **All expensive work happens in async preparation (parse time).** This covers data fill, statistics, trendline fits, tick generation, label formatting and text measurement inputs. The model the painter receives is **paint-ready**, so synchronous paint is bounded.
- **Work bounds per chart:**

  | Item | Bound |
  |---|---|
  | Ticks per axis | ≤ 1,000 (Office would refuse more) |
  | Data-label strings | ≤ 4,096 |
  | Path segments painted | ≤ 262,144 |
  | Trendline samples | ≤ 512 per trendline |
  | Contour grid | ≤ 256 × 256 |
  | Treemap/sunburst nodes | ≤ 4,096 |

  On a breach, the affected element is dropped with a diagnostic; for analytic families the chart is refused (§4.3).
- **Document-wide retention.** Charts retain ≤ 64 MiB desktop / 32 MiB mobile, charged per model. Over that, later charts become cards.
- **Tiny units.** A tiny `majorUnit` or an extreme finite scale is handled by the tick bound: the unit is enlarged to fit ≤ 1,000 ticks, with a diagnostic.
- **Painter containment.** The painter dispatch in `content-paint.ts` wraps the engine in `ctx.save()`/`try`/`finally ctx.restore()`. A throw paints the card with "This chart could not be drawn" and never affects later shapes.
- **Preparation before paint.** `getPaintables` DOCX and PPTX branches call `ensureChartEngine(contents)`. It is typed over `DrawingContent`, and it **catches** a load failure, so pages still return, with cards.
- **Synchronous APIs.** The public synchronous `renderPages`/`renderSlide` stay synchronous. When the engine was never loaded in that realm (for example, a transferred model) they paint cards.
- **Preload export.** A new public export, `preloadCharts(): Promise<void>`, lets headless embedders load the engine first. It is a one-line eager wrapper over the existing eager `chartLoader`, and the budget table must show it.

### 6.6 Coverage, diagnostics and accessibility [rev: codex I9, agy 8]

- **Coverage.** `featureOfKind` becomes a shared export, and DOCX coverage uses it. Representations:
  - native → `chart`;
  - card → `chart-placeholder`;
  - an authored fallback picture → the existing fallback representation, with the original feature (`chart`/`ChartEx`).
- **Diagnostics.** New codes: `chart-data-unavailable`, `chart-refused`, `chart-invalid-log-value`, `chart-singular-trendline`, `chart-label-feature-unsupported` and `chart-axis-unplaceable`. Each has a reason string.
- **Accessibility.**
  - Each chart carries a bounded description: type, title, series names, value range, and "placeholder: reason" for cards.
  - The description travels with its page/slide placement in the paintables metadata (lazy, charged to the retention budget).
  - The paged viewer exposes a per-page list of `role="img"` elements with `aria-label` = description. They sit in an offscreen semantic layer in reading order (document order; in PPTX, slide z-order); this is the same pattern as `sheetCharts.tsx`.
  - A header or footer chart repeated on each page appears once per page.
  - The descriptions follow the copy permission: when `allowCopy={false}`, they are still read aloud but are not selectable.

### 6.7 Rendering rules

- **Fonts.** Chart text uses the theme's minor font through the font manager, sized from `txPr` and clamped to 6–72 pt.
- **Determinism.** Layout is deterministic for a given input and font metrics.

## 7. Limits summary

| Item | Limit | On breach |
|---|---|---|
| Series / points (display) | 64 / 4,096 (existing) | Truncate display, marked `truncated`, diagnostic. This applies to **non-analytic** families only. |
| Analytic inputs (histogram, Pareto, box & whisker, treemap/sunburst, trendlines, error stdDev/stdErr) | 65,536 points per series, processed fully | Refuse the chart (or drop the trendline / error bar), never use a prefix |
| Labels | 255 units; 4 MiB per chart (existing) | Truncate, marked |
| Chart part | 8 / 4 MiB desktop/mobile | Unsupported → fallback or card |
| Embedded workbook | §6.4 table | Refused |
| Work bounds | §6.5 table | Drop element, or refuse the analytic chart |
| Per-document chart retention | 64 / 32 MiB | Later charts become cards |
| Authored parameters out of range (polynomial order, period, bins > 1,000, `logBase` outside 2–1000) | none | Diagnose, and don't draw that element. The chart is refused if the parameter is essential (histogram bins). |

Every family handles empty series, all-blank values, NaN/∞ (treated as blank), negative values, and a single point without throwing.

## 8. Testing and validation [rev: codex I11]

- **Numeric oracles.** Hand-computed expected values (not only paint calls) for: tick generation, date axes (month/year boundaries, unequal spacing, both date systems), display units, every number-format class, ofPie splits, the stock role mapping, bubble scaling, every trendline type (with known datasets), error bars, histogram bins (edges, underflow/overflow, `intervalClosed`), Pareto cumulative values, box-whisker quartiles (both methods, odd and even n), waterfall subtotals, and treemap/sunburst aggregation.
- **Unit tests** cover parse → model, and paint with a recording context for every family and element.
- **Representation tests** cover §6.3 in all three formats, including no-fallback and chunk failure.
- **Data tests** cover §6.4.
- **Robustness tests** cover extreme scales, tiny units, huge forecasts, 200 charts in one document (retention), and a torture file per format. Each asserts allocation/work counters, not just that a valid PNG comes out.
- **Goldens.**
  - A synthetic chart gallery (one chart per family and major element), per format, is recorded on the Mac.
  - Existing XLSX chart goldens stay pixel-identical, except where a §5 element legitimately changes the picture. Each change is listed with a reason.
- **References.**
  - Office-authored files and Office screenshots go in `validation/office-reference/`, with producer and version recorded, one per promised semantic case.
  - The owner supplies them. Until then, files from python-docx/python-pptx/openpyxl plus LibreOffice renders are interim evidence. Those cases are marked **unverified** in `docs/validation/chart-compat.md`, and "Office-matched" is not claimed for them.
  - Smoothing is claimed matched only after reference comparison.
- **Browser tests:**
  - DOCX/PPTX charts render in the paged viewer;
  - dark mode leaves chart pixels identical to light;
  - the per-page chart list exposes accessible names;
  - CLS is 0 when a chart loads.

## 9. Delivery constraints

- **Eager budget.** Measured by `bun run build && bun scripts/eager-budget.ts`: Node zlib gzip of `dist/index.js`, ceiling 38,720 B. Every task records before and after.
  - Net change ≤ 0 per task.
  - Stream A's first task lands the adapter seam and deletes the mini-reader before anything else, so later tasks start with the freed bytes.
- **Goldens.** Changed only with individual inspection and a stated reason.
- **Review.** Every task is reviewed by a model family other than its implementer's. Stream B's field-map task is reviewed by two families.
- **CI.** CI passes on Linux; the pixel goldens stay Mac-only per `ci.yml`.

## 10. Decisions taken from review round 1

| Question | Decision |
|---|---|
| Surface-as-contour | `surfaceChart` (2D contour): native. `surface3DChart`: fallback picture, else card. |
| ChartEx scope | In this project as stream B, after stream A; field map first; region map stays fallback |
| Embedded-workbook cap | 8/4 MiB inflated admission, plus a new bounded values-only reader with the §6.4 budgets |
| Trendlines and error bars | In, with numeric oracles and explicit refusal rules |
| Rename | The runtime tag stays `xlsxChart`; type and module names become neutral |
