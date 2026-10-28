# XLSX Pivot Tables, Pivot Charts, Slicers and Timelines — Design Spec

**Status:** draft 2. Builder check folded in (§7.7–§7.10) and owner answers recorded (§11); two reviewers (different model families) next.
**Date:** 2026-10-29 · **Baseline:** `main` @ `ba8abf83` (v0.10.0).
**Decision marks:**
- **[owner]** is a decision the owner made.
- **[call]** is the author's call, open to review.
- **[rev]** is a decision taken from a review finding.
- **[measure]** is a behaviour stated as a hypothesis. It must be confirmed or corrected against Excel before it is coded as fact.

## 1. Intent

officeview renders Office documents to canvas. Excel pivot tables summarise a block of source rows into a cross-tab: rows, columns, filters and summarised values. Pivot charts plot a pivot table. Slicers and timelines are on-sheet filter controls that drive one or more pivots.

Today the viewer knows none of this.
- Excel saves a pivot's last output as ordinary cell values, so the numbers show.
- The pivot's look is lost: style, banding, bold subtotals, field buttons and expand/collapse buttons. It renders as a plain grid.
- Pivot charts draw from their saved caches but without field buttons.
- Slicers and timelines probably paint as Excel's fallback text box ("This shape represents a slicer…").
- None of it responds to the user.

**Why:** rounding out the XLSX viewer, the same motive as sparklines. **[owner]** The owner asked for the full feature: "work with it, full feature".

**Goal:**
- **Display:** every pivot table, pivot chart, slicer and timeline Excel draws, the viewer draws, matching Excel closely enough to pass a pixel comparison.
- **Interaction:** in the interactive viewer, the user can work with them the way Excel allows:
  - expand and collapse, filter, and sort;
  - rearrange fields, change how values are summarised and shown, and group;
  - click slicers and timelines;
  - pivot charts follow their pivot.
- **View-only. [owner]** Nothing is saved back to the file, and reopening shows the file as saved. This is the same contract as the interactive viewer's sort and filter.

**Success** means:
1. Pivot tables render with Excel's style, layout and buttons in page mode, export, and the interactive viewer, on the resident and paged paths. Each feature in §4 is verified against an Excel reference (§9).
2. Pivots are recomputed from the cache records saved in the file, and the recomputation reproduces Excel's saved output. This is proven per file on first interaction (§6.4) and per case against Excel (§9).
3. Every interaction in §4 changes the pivot exactly as Excel would for the same action, verified by driving Excel to make the same change (§9).
4. A pivot that cannot be recomputed faithfully is never shown with wrong numbers. It stays as saved, non-interactive, and the viewer says why (§7.6).
5. Workbooks without pivots pay nothing: the eager bundle and the hot path are unchanged (§8).

## 2. Non-goals and anti-goals

**Non-goals:**
- **Writing anything back to the file**, or exporting a changed pivot as a new file. **[owner]**
- **Refreshing from outside the file.** External connections, OLAP cubes and the Power Pivot data model are out of scope.
  - A pivot built on those is displayed as saved and is not interactive (§7.6).
  - "Refresh" means recomputing from the cache saved in the file, never reading the live source range or a connection. **[call]** The source range may have been edited since the last refresh; Excel's own Refresh would pick that up, but our pivot shows the cache, as Excel does until refreshed.
- Creating new pivots, slicers or timelines. Inserting, deleting or renaming calculated fields and items is out too, but existing ones are evaluated (§5.6).
- Drill-through, which in Excel double-clicks a value to open a new sheet of source rows.
- Pivot "Options" dialogs beyond the actions listed in §4. For example, the user cannot change the report layout, style, or grand-total settings, which are taken from the file. **[call]**
- Windows-only behaviour. Excel for Mac is the reference (§9).

**Anti-goals** (failure even if it technically works):
- **A wrong number.** A recomputed value that differs from what Excel would show is worse than no interaction. Unverifiable means view-only plus a stated reason.
- **Silent staleness.** After a pivot changes, a value elsewhere that depends on it must not keep its old number without a sign that it is stale (§7.5).
- **Slower open, scroll or zoom for workbooks without pivots.**
- **Pivot code in the eager bundle**, beyond a tiny detection stub.

## 3. Background: the file format

Pivot data is spread over several parts. A builder should read ECMA-376 Part 1 §18.10 (pivot tables) and the MS-XLSX extensions (x14 and x15) for slicers and timelines.

- **Workbook:** `<pivotCaches><pivotCache cacheId r:id/>` maps cache ids to `pivotCacheDefinition` parts. Slicer caches and timeline caches are listed in the workbook `extLst` (`x14:slicerCaches`, `x15:slicerCaches` for table slicers, and `x15:timelineCacheRefs`).
- **`pivotCacheDefinition`:**
  - `cacheSource` is `worksheet` (a ref plus sheet, or a name/table), `external`, `consolidation` or `scenario`.
  - `cacheFields`: each `cacheField` has a name, `numFmtId`, `sharedItems` (a typed item list with min/max and contains-flags), `fieldGroup` (date and number grouping: `rangePr`, `discretePr`, `groupItems`, `par`, `base`), and `formula` (a calculated field, with `databaseField="0"`).
  - Attributes: `refreshOnLoad`, `saveData`, `missingItemsLimit`, `refreshedVersion`.
  - Data-model caches carry x15 extensions.
- **`pivotCacheRecords`:** one `<r>` per source row. Each child is `x` (an index into `sharedItems`) or an inline value: `n` number, `s` string, `d` date, `b` boolean, `e` error, `m` missing. The part may be absent (`saveData="0"`), and it can be tens of MB.
- **`pivotTableDefinition`**, linked from the sheet's rels:
  - **Placement and layout:** `location` (ref, first header/data row/column, page-field rows and columns); `compact`, `outline`, `gridDropZones`, `indent`, `dataOnRows`, `dataPosition`; grand totals and captions.
  - **`pivotFields`**, one per cache field:
    - `axis`, `compact`, `outline`, `subtotalTop`, `insertBlankRow`, `defaultSubtotal` and named subtotal flags, `sortType`, `showAll`, `autoShow` (top 10), `multipleItemSelectionAllowed`;
    - `items`: `x`, `h` hidden, `sd` collapsed, `t` subtotal or blank type, `n` custom name, `f` calculated item, `m` missing;
    - `autoSortScope`.
  - **Axes:** `rowFields`, `colFields` (`x="-2"` is the Values pseudo-field), `pageFields`.
  - **`dataFields`:** `fld`, `subtotal`, `showDataAs`, `baseField`, `baseItem`, `numFmtId`. The x14 `pivotShowAs` adds % of parent, running-total %, and rank.
  - **Saved layout:** `rowItems` and `colItems`. These are the layout Excel computed, a free oracle for the engine.
  - **Formatting and filters:** `formats` and `conditionalFormats` (a `pivotArea` plus `dxfId`), `chartFormats`, `filters` (label, value, date, top 10, with `evalOrder`), `pivotTableStyleInfo` (name and the row/column header and stripe flags), and `pivotHierarchies` (OLAP).
  - x14 extensions such as `fillDownLabels` and `visualTotals`.
- **Sheet cells:** the pivot's last output is stored as ordinary cells inside `location`. Pivot *styles* are not stored in cells; Excel applies them when it draws. Cells may still carry `s=` formats, from number formats or explicit formatting. **[measure]** Find exactly what Excel writes into the cells' `s=`.
- **Styles (`styles.xml`):**
  - `dxfs`: the differential formats used by pivot `formats` and by custom styles.
  - `tableStyles`: custom table, pivot and slicer styles, plus `defaultPivotStyle`. The x14 `slicerStyles` extension holds custom slicer styles.
  - **Built-in pivot styles are not in the file.** PivotStyleLight1–28, Medium1–28 and Dark1–28 must be authored in code. They are defined relative to the theme, and ECMA-376 ships `presetTableStyles.xml` with their definitions. **[measure]** Check them against Excel renders.
- **Pivot charts:** a normal chart part plus `c:pivotSource` (pivot name, `fmtId`) and `c:pivotFmts`. Series refs point into the pivot's output range. Field buttons come from the x14 `pivotOptions` (`dropZoneFilter`, `dropZoneCategories`, `dropZoneData`, `dropZoneSeries`, `dropZonesVisible`).
- **Slicers:**
  - `xl/slicers/slicerN.xml`: caption, cache, `columnCount`, `rowHeight`, `style`, `showCaption`, `startItem`, `lockedPosition`.
  - `xl/slicerCaches/slicerCacheN.xml`: the pivot tables it drives; tabular data with sort order, `crossFilter` and `showMissing`; and items with `x`, `s` selected and `nd` no-data. A table slicer instead points at an Excel table column.
  - In the drawing: `mc:AlternateContent` with an `sle:slicer` Choice and a fallback shape.
- **Timelines:** x15 timeline parts and timeline caches (a date field, selection start/end, level), drawn in the same way through AlternateContent (`tsle:timeslicer`), with built-in TimeSlicerStyle definitions.

### 3.1 What the codebase already has (recon 2026-10-29, at `ba8abf83`)

- **Parsing:**
  - Nothing reads pivot, slicer or timeline parts, `dxfs` or `tableStyles`.
  - `parseStyles` (`src/xlsx/parse.ts:93`) reads only numFmts, fonts, fills, borders and cellXfs. `XlsxCellStyle` has no underline, font name or indent.
  - Excel table (ListObject) styles are not rendered.
- **Slicers today:** `XLSX_REQUIREMENTS` (`src/drawing/coverage.ts:110`) lacks the slicer and timeline namespaces, so the fallback shape paints.
- **Sparklines are the template:**
  - an eager parser plus resolver, behind an access interface shared by the resident and worker paths;
  - a lazy painter (`sparklineLoader.ts`);
  - seams at `render.ts:818`, `renderViewport.ts:262`, `pagedSource.ts:583`, and the streaming `sheetRecords.ts`.
- **The interactive viewer can permute and hide rows but cannot change cell contents.** `SourceStamp.valueGen` is fixed at 0 on the resident path (`src/xlsx/view/source.ts:693`), and tiles are keyed by `CacheIdentity` (`contracts.ts:467`). Sort and filter are a `RowOrder` plus a `RowVisibility` over a source range.
- **Controls are DOM.**
  - Header buttons: `GridHeaders.tsx:299`.
  - One anchored popup slot: `GridSemantics.tsx:409`.
  - A reusable checklist-with-search: `FilterDropdown.tsx:18`.
  - A modal: `SortDialog.tsx`.
  - Menus: `MoreMenu.tsx`.
  - Coarse pointers switch to a bottom sheet (`controlMedia.ts`).
  - Nothing is drawn in cells as a control, and there is no side panel.
- **Formula engine:**
  - It evaluates on load only (`parse.ts:464`). The paged path uses cached values only.
  - There is no `GETPIVOTDATA`, `STDEV`, `VAR`, `MEDIAN`, `SUBTOTAL` or `COUNTBLANK`.
  - `evaluateFormula(expr, ctx)` (`evaluator.ts:1145`) can evaluate against supplied names.
- **Charts:**
  - Cached series never re-read cells, and there is no refill API.
  - A refill means rebuilding the model with `pending` ranges, running `fillChartRanges` (`fill.ts:182`), passing ownership with `succeedChartModel`, and repainting.
  - `c:pivotSource` is unparsed.
- **Large files:**
  - Non-drawing parts over 4 MiB (2 MiB on mobile) route the workbook to the paged path (`admission.ts`).
  - `RangePackage` caps parts at 8 MiB, so cache records must be *streamed*, using `openPart` and `tokenizeBytes` with `sharedStringEvents` as the template.
  - Worker replies are grant-checked against the memory ledger.
  - With no worker route, a paged decision falls back to a full main-thread `parseXlsx`, so pivot loading must never live inside `parseXlsx`.
- **Eager budget:** a 38,720 B gzip ceiling. Usage was about 35.95 KB after sparklines, which leaves about 2.7 KB.
- **The interactive-viewer spec** (`docs/superpowers/specs/2026-10-10-interactive-viewer.md:26`) lists "pivot/chart interaction; slicers" as non-goals. This spec amends it. **[owner]**
- **Corpus:** none of the 782 xlsx packages in the repo contains a pivot, slicer or timeline. apache/poi has `ExcelPivotTableSample.xlsx`, and LibreOffice's `sc/qa/unit/data/xlsx/` very likely has more. Both are reachable through `scripts/fetch-corpus.ts`.
- **Excel automation:** Excel 16.114's AppleScript dictionary can create pivot tables, add row, column, page and data fields, auto-sort, and set a pivot style. It has no slicer or timeline commands.

## 4. Scope **[owner]**: the full feature, view-only

| Area | Display (as saved) | Interaction (view-only) |
|---|---|---|
| Layout | Compact, outline and tabular forms; repeat item labels (`fillDownLabels`); blank rows; subtotals top or bottom; grand totals on or off; Values on rows or columns; page-field area; empty-cell and error captions | — (taken from the file) |
| Style | Built-in pivot styles (84) and custom `tableStyles`; row/column header and stripe flags; `formats` dxf overrides; pivot conditional formats **[call]**; number formats per data field | — |
| Buttons | Field-header filter and sort buttons; expand/collapse (±) buttons; page-field dropdowns; honouring the "show buttons", "show headers" and "show drill" flags | Each button works (below) |
| Expand and collapse | Collapsed items as saved (`sd="0"`) | Expand or collapse an item, a whole field, or everything |
| Filter | Hidden items (`h`), page-field selections, label, value, date and top-10 filters as saved | Item checklist with search; label, value, date and top-10 filters; page-field single and multiple selection; clear filters |
| Sort | Saved order: ascending, descending, manual, by a value column (`autoSortScope`), custom lists | Sort ascending or descending by label or by a value column |
| Field list | — | A panel listing the fields: move a field between Rows, Columns, Values and Filters, reorder within an area, remove it |
| Values | Summary functions; show-values-as modes; number format | Change the summary function (sum, count, average, max, min, product, count numbers, stdev, stdevp, var, varp); change show-values-as (all modes incl. x14) with its base field and item |
| Grouping | Saved date groups (years, quarters, months, days, hours, minutes, seconds) and number range groups; manual (discrete) groups | Group or ungroup a date or number field; change the grouping levels or interval |
| Calculated | Saved output of calculated fields and items | Recomputed when the pivot changes (§5.6); no editing of their formulas |
| Pivot charts | Chart from the saved cache, with field buttons per `pivotOptions` | The chart follows its pivot; its field buttons open the same filter and sort menus |
| Slicers | Drawn per slicer style, columns, item states (selected, unselected, no data) and scroll position | Click to select, Ctrl/⌘-click for multi-select, clear filter; one slicer driving several pivots; cross-filter "no data" states |
| Timelines | Drawn per timeline style and level, with the selected period | Select a period by click or drag; change the level (years, quarters, months, days); clear |
| Table slicers | Slicers on Excel tables, drawn | Click filters the table's rows through the existing filter engine **[call]** |
| Dependent cells | — | Formulas that read the pivot (cell refs, `GETPIVOTDATA`) and charts on those cells update or are marked stale (§7.5) |

**Interaction lives in the interactive viewer only. [call]** Page mode, export and print show the pivot as saved, matching the existing split where sort and filter are interactive-only.

**Excel tables (ListObjects) get their table styles too. [owner]** They share the style engine and join Stage 1 (§11).

## 5. Engine behaviour

The engine is a pure function:

`(cache definition, cache records, pivot definition with the current view state) → layout`

The layout is the cells of the pivot region: values, labels, captions, per-cell style roles, button placements, and the rectangle's size.

### 5.1 Source data
- **Records become columns.** Records decode into one column per cache field: shared-item indices where items are shared, and typed inline values otherwise.
- **Order:** items keep the `sharedItems` order, then the `pivotField items` order, which is the display order before sorting.
- **Missing values:**
  - Missing (`m`) and blank values form the "(blank)" item.
  - Errors in values propagate as error results per Excel. **[measure]**
- **Grouping:**
  - Grouped fields (`fieldGroup`) derive their items from a base field with `rangePr` (by date part, or number interval with start and end) or `discretePr` and `groupItems`.
  - Excel's out-of-range buckets use labels such as "<1/1/2020" and ">12/31/2024". **[measure]**
  - Grouping changes made during interaction rebuild the derived items.

### 5.2 Axes and items
- **Visible items:** for a field on rows, columns or filters, an item is visible unless it is hidden (`h`), excluded by a filter, or produces no data under the other filters.
  - Items with no data are dropped unless `showAll`.
  - Whether filters on other fields restrict which items appear is **[measure]**.
- **Collapse:** a collapsed item (`sd="0"`) shows its subtotal row in place of its children.
- **Subtotals:**
  - `defaultSubtotal` uses the data field's own function.
  - Named subtotals (`sumSubtotal` and others) add one subtotal row per function.
  - Placement follows `subtotalTop`, honoured only in compact and outline forms. **[measure]**
- **Grand totals:** a grand total row and/or column, each with its caption.
- **Values pseudo-field (`x="-2"`):** places multiple data fields as an axis level, and its position matters.
- **Saved layout as oracle:** `rowItems` and `colItems` encode the saved layout as repeat counts plus item indices. The engine's layout for the *saved* state must equal them exactly (§6.4).

### 5.3 Values
- **Aggregation:** each data cell aggregates the matching records with the data field's function. `count` counts non-empty values and `countNums` counts numbers; the rest follow Excel's statistical definitions (stdev and var are sample, stdevp and varp population).
- **Totals are computed from the records, not by summing subtotals.** For example, an average's grand total is the mean of all records. **[measure]**
- **Filtered-out items are excluded from totals.** Visual totals apply to OLAP only, which is out of scope. **[measure]**
- **Show values as**, each with Excel's base-field and base-item rules:
  - % of grand total, % of column total, % of row total;
  - % of, % of parent row total, % of parent column total, % of parent total;
  - difference from, % difference from;
  - running total in, % running total in;
  - rank smallest to largest, rank largest to smallest;
  - index.

  Previous and next base items are encoded as `baseItem` 1048828 and 1048829. Division by zero gives `#DIV/0!`, shown as `errorCaption` if set. **[measure]**
- **Number format:** the data field's `numFmtId` first, then `formats` overrides, using the existing number-format engine.

### 5.4 Sorting
- **Field sort:** `sortType` ascending or descending by item label, with numbers before text and blanks last. **[measure]**
- **Custom lists:** by default Excel sorts its built-in custom lists (Sun–Sat, Jan–Dec) in calendar order. **[measure]**
- **Sort by values:** `autoSortScope` sorts by a data field's values within each parent group.
- **Manual order:** the `items` order is kept when `sortType` is manual or absent.

### 5.5 Filters
- **Label filters** (`filters/filter` with `type` caption*): string comparisons that are case-insensitive, with wildcards.
- **Value filters** (value*): test each item's aggregate on `iMeasureFld`.
- **Date filters** (date*, including dynamic ones such as "this month"): dynamic date filters are evaluated against the **saved file's reference date, not today**. **[call]** This reproduces the saved output; a "relative to today" option is left out.
- **Top 10** (`count`, `percent`, `sum`): ties are handled per Excel. **[measure]**
- **Order:** filters apply in `evalOrder`.
- **Page fields:** each filters by one item or a set of items.

### 5.6 Calculated fields and items
- **Calculated fields:** a calculated field's formula combines *sums* of other fields, evaluated per cell after aggregation, as Excel does. They use the existing formula engine through `evaluateFormula` with field names bound as names.
  - If the formula uses a function or construct the engine cannot evaluate, that pivot is not interactive (§7.6).
  - Quoted field names with spaces must lex correctly; this is untested today.
- **Calculated items:** a calculated item is a formula over items of one field, evaluated per record group. Its subtotal and grand-total interactions are subtle.
  - **[call]** Calculated items are recomputed only if the Stage 2 oracle tests pass on Excel-made references.
  - Until then, a pivot containing calculated items is not interactive.

### 5.7 Layout to cells
- **Forms:**
  - **Compact:** all row fields in one column, indented by `indent` characters per level.
  - **Outline:** one column per field, labels on their own row.
  - **Tabular:** one column per field, labels on the same row as the first value.
- **Optional rows:** "repeat item labels" and blank rows after items.
- **Captions:** the "Row Labels" and "Column Labels" header captions, or field names when headers are shown; the data caption.
- **Style roles:** each output cell gets one, such as header row, first column, subtotal row 1–3, grand-total row, row and column stripes, page field labels and values, or blank row. The style engine (§6.5) resolves roles to formats.
- **Placement:** the pivot keeps its top-left anchor from `location`. Its new size follows from the layout.

## 6. Architecture

### 6.1 Model (eager, tiny)
- **What the eager parse records:** for each sheet, the pivot locations (sheet, rectangle, part path) and each slicer and timeline anchor.
- **What it skips:** no definitions or records. That is enough to know a sheet has pivots and where they are.
- **Size:** the eager stub must stay within the remaining budget (§8).

### 6.2 Parsing (lazy chunk `pivots`)
- **The chunk** holds the full definition parsers (cache definition, pivot definition, slicer and timeline parts, styles `dxfs` and `tableStyles`), the built-in style tables, the engine, and the painters. It loads on the first sheet that has a pivot, slicer or timeline.
- **Resident path:** parts are read from `OfficePackage` after `parseXlsx` returns, not inside it.
- **Paged path:**
  - Definitions are read with `RangePackage`.
  - Records are **streamed** in the worker (`openPart` and `tokenizeBytes`) into a compact columnar store: typed arrays of item indices plus per-field value tables, charged to the worker memory pool.
  - The engine runs in the worker, and layouts come back as grant-checked replies.
  - A record store over its cap leaves the pivot as saved and non-interactive (§7.6). The cap is a builder figure, sized from realistic workbooks.

### 6.3 Engine placement and language **[owner] [call]**
- **Language:** the engine is TypeScript. The owner offered Rust/WebAssembly.
- **Why TypeScript:**
  - Most of the cost is decoding the XML into columns, not aggregating them.
  - Grouped aggregation over typed arrays is fast in JS.
  - WASM would add a `.wasm` asset that every host bundler must handle (Next.js, Vite), a Rust toolchain in CI and in the examples, and a second language for implementers and reviewers.
- **The gate:** Stage 2 includes a benchmark. On a realistic large cache, the worker must recompute a 1,000,000-record pivot (3 row fields, 1 column field, 2 data fields) in **≤ 200 ms** on the development machine, after decode. Decode time is reported separately.
- **If it misses:** the plan proposes Rust/WASM with the measured numbers and the costs listed above, and the owner decides.
- **Placement:** for resident workbooks the engine runs on the main thread inside the lazy chunk, which is fine because their caches are under 4 MiB. For paged workbooks it runs in the worker.

### 6.4 Faithfulness check (per pivot, on first interaction)
1. **Before** the first interaction with a pivot, the engine computes the layout for the saved state.
2. **It compares** that layout with the saved `rowItems` and `colItems`, and every value with the saved cell values inside `location`. Numbers must match within Excel's display precision; text must match exactly.
3. **On a mismatch** the pivot is not interactive, and the reason (the first differing cell) is reported as a diagnostic.
4. **Its purpose:** every real file proves the engine on its own data before any number we compute reaches the user.
5. **When it is skipped:** when the cache is stale (`refreshOnLoad="1"`, or the saved cells provably predate the records), the check cannot run. **[call]** The pivot is then still interactive, but the first interaction shows a one-line notice that values are recomputed from the saved data.

### 6.5 Style engine
- **Inputs:**
  - the built-in pivot style table, authored from `presetTableStyles.xml` and checked against Excel;
  - custom `tableStyles` and `dxfs`;
  - the style-info flags;
  - `formats` overrides.
- **Resolution:** each cell's style role is resolved to a final format with Excel's element precedence (ECMA-376 §18.8.40), using theme colours with tint.
- **Output:** the result is a `XlsxCellStyle` extended with underline, indent and font name, gaps the model has today (§3.1).
- **Cell formats:** the cell's own `s=` format layers over the pivot style as Excel does. **[measure]**

### 6.6 Painting
- **Pivot cells:** painted through the normal cell paint with the resolved styles. A *pivot value layer* overrides the saved cell contents and styles inside the pivot rectangle.
- **Interactive path:** a pivot change bumps a per-sheet value generation, so tiles re-key. Today `valueGen` is fixed at 0 for resident sheets; this spec makes it live for pivot sheets only.
- **Buttons:** filter arrows, ± buttons and page-field dropdowns are painted on canvas in every mode, so page mode and export match Excel. In the interactive viewer each one also gets a DOM hit target with an accessible name.
- **Slicers and timelines:** painted as drawings in their anchors (the AlternateContent Choice), replacing the fallback shape. In the interactive viewer they are DOM-backed controls positioned over the drawing.
- **Pivot charts:**
  - At display, `c:pivotSource` is linked to its pivot.
  - On a pivot change, the series are rebuilt from the new layout and the chart is refilled through `fillChartRanges` and `succeedChartModel`, then repainted.
  - Series are one per leaf column item, with categories from the row items (multi-level).
  - How Excel assigns formats to series that appear or disappear (`pivotFmts`) is **[measure]**.

### 6.7 Interaction UI
- **Filter, sort and value-filter menus:** they reuse `FilterDropdown`, `ConditionEditor` and the popup slot, opened from the pivot's buttons.
- **Field list:** a new side panel in the interactive viewer, which becomes a bottom sheet on coarse pointers through `controlMedia`.
  - It lists the fields and the four areas.
  - Drag and drop works, with a full keyboard and menu alternative ("Move to Rows/Columns/Values/Filters", reorder up and down).
  - It follows the existing control patterns (roving focus, Escape, 44 px targets).
  - It opens from a pivot cell's context action and from the toolbar when the active cell is in a pivot.
- **Value field settings:** a small dialog for the summary function and show-values-as with its base field and base item. It follows the `SortDialog` modal pattern.
- **Group dialog:** start, end, interval and date levels.
- **Shared view state:** all of it is view state, per sheet, discarded on close.

### 6.8 Worker protocol
- **New commands** join `viewerProtocol.ts`:
  - pivot open, which streams records and builds the store;
  - pivot apply, which takes a view-state delta and returns a layout under a grant;
  - pivot release.
  - Slicer clicks are applies to every pivot the slicer drives.
- **Lifetimes:** byte grants and release-on-every-exit follow the sparkline precedent (`RemoteSparklineCache`).

## 7. Behaviour rules

### 7.1 Size changes and overlap
- When a pivot grows or shrinks, its top-left stays fixed.
- **Collisions:** if the new rectangle would cover non-empty cells that are not part of this pivot (sheet cells, another pivot, a table), the change is refused. Excel refuses with "A PivotTable report cannot overlap…"; we show the same message, and the pivot stays as it was.
- **Shrinking:** cells that the pivot no longer covers show as empty.
- **[measure]** Find Excel's exact overlap rule, including whether formatting-only cells count.

### 7.2 Slicer and timeline linkage
- A slicer filters every pivot listed in its cache.
- **Cross-filtering:** with `crossFilter` on, items with no data under the other slicers show as "no data" and sort per `sortOrder` and `showMissing`. **[measure]**
- A timeline filters its date field across its pivots.
- **Page fields:** a slicer selection is a page-field-style multi-item filter, and it shows in the pivot's filter button state.

### 7.3 Pivot charts
- A pivot chart follows its pivot's filters, sort, expand/collapse, field moves and grouping.
- Its field buttons reflect and drive the same view state.

### 7.4 Multiple pivots on one cache
- Pivots that share a cache share the record store.
- Each pivot keeps its own view state, unless a slicer links them.

### 7.5 Dependent values **[call]**
- **Resident path:**
  - After a pivot changes, formulas that read cells inside the pivot rectangle, or call `GETPIVOTDATA` on it, are recomputed by the formula engine.
  - This needs a recompute entry point for a dependency set; the engine evaluates on load only today.
  - Charts on those cells are refilled.
  - `GETPIVOTDATA` is implemented against the live layout.
- **Paged path:** formulas are never evaluated there. Dependent formula cells are marked stale (a subtle indicator plus a tooltip), and a status-line notice appears.
- **Fallback:** if dependents cannot be found reliably, the notice appears for any pivot change on a sheet that has formulas outside pivots.

### 7.6 Fail closed: displayed as saved, not interactive
Each case below gets one diagnostic code (`pivot-*`) and a short, user-visible reason shown on the pivot's buttons or in the field list:
- the records part is absent (`saveData="0"`);
- an OLAP or data-model cache, or an external, consolidation or scenario source;
- an unknown construct in the definition;
- a calculated field the formula engine cannot evaluate;
- calculated items before they are verified (§5.6);
- the faithfulness check fails (§6.4);
- the record store is over its cap.

Display of the saved output and the style still happens in every case.

### 7.7 Stale caches and `refreshOnLoad` **[call]** (builder check 1)
- Excel recomputes a `refreshOnLoad="1"` pivot when it opens the file, so what Excel draws is the recomputed output, not the saved cells.
- **Before Stage 2:** the saved cells are shown.
- **From Stage 2:** such a pivot is displayed from the engine's output in every mode, including page mode and export. The engine therefore loads for display when a sheet has one.
- **If the engine cannot handle the pivot** (any §7.6 case), it is shown as saved, with the diagnostic.
- The §6.4 check cannot run for these pivots. Their correctness rests on the Excel-oracle tests (§9), which include `refreshOnLoad` cases.

### 7.8 Column widths after a change **[call]** (builder check 3)
- Excel autofits the pivot's columns on update when `useAutoFormatting` is on, which is the default.
- The viewer follows it: after a change it autofits the columns the pivot occupies, using the existing text measurement. **[measure]** Excel's exact autofit rule: padding, and whether it only grows.
- Columns the user resized in this session keep the user's width.
- This is a geometry change, so it invalidates layout on both paths like a user resize.

### 7.9 Pivot output is source data, not paint **[call]** (builder check 4)
- **A source change:** the pivot value layer changes the cell source that every consumer reads: painting, find, copy, grid ARIA, value queries, and the paged worker's view of the cells.
- **Why:** a paint-only overlay would let copy or search return the old numbers, which is a wrong number.
- **Growing past the sheet:** a pivot that grows past the sheet's used extent extends it.
- **Sheet sort and filter:** the interactive viewer's sheet-level sort and filter are disabled for a range that intersects a pivot, with a message. Excel does not apply sheet sort or filter across a pivot either. **[measure]**
- **Hidden rows** crossing a pivot stay hidden. The pivot lays out in its cells; hiding is the sheet's business.

### 7.10 Page mode and export follow print; the interactive viewer follows screen **[call]** (builder check 5)
- **Page mode, export and print** draw what Excel prints:
  - no filter or page-field dropdown arrows;
  - ± buttons only when `printDrill` is set.
- **The interactive viewer** draws what Excel shows on screen: all buttons, per the show-buttons flags.
- This matches today's convention, where page mode draws no autofilter arrows.
- References for page mode are captured from Excel's printed appearance (`copy picture` with appearance *printer*); interactive references use appearance *screen*. **[measure]**

## 8. Constraints
- **Eager budget:**
  - The eager graph grows only by the detection stub.
  - `tests/helpers/eager-graph.ts` gains a rule forbidding everything under `src/xlsx/pivots/` except the stub.
  - `check-eager.ts` and `eager-budget.ts` must pass.
- **No new runtime dependencies.** WASM would be the exception, and only through the §6.3 gate.
- **Hot path:** sheets without pivots add no per-tile work beyond an `undefined` check.
- **Determinism:** the layout is a pure function of its inputs, and painting is a pure function of the layout, zoom and DPR.
- **Memory:** the record store, layouts and style tables are charged to the existing ledgers.
- **Accessibility:** every pivot control is reachable by keyboard and has an accessible name. The pivot region keeps the grid's ARIA model, with expanded/collapsed state on row headers. **[call]**

## 9. Reference evidence **[owner]**: match Excel

Excel for Mac 16.114 is installed. The sparkline harness (`scripts/sparkline-refs/`) is reused where it is generic: `capture` (`copy picture` through AppleScript), `diffFrames`, crop and the sidecars.

- **Ground-truth workbooks:**
  - **Pivot tables.** A script generates the source data and drives Excel through AppleScript to create each case's pivot, using its create, add-field, data-field, sort and style commands, then saves. Excel's saved file is the ground truth for display *and* for the engine: its `rowItems`, `colItems` and cell values.
  - **Settings AppleScript cannot reach** (show-values-as, grouping, layout forms, filters): set them through further AppleScript properties if the dictionary allows. **[measure]**
  - **Otherwise:** hand-write the pivot definition with `refreshOnLoad="1"`, open it in Excel to recompute, and save under a new name. **[measure]** Check that Excel recomputes on open without a prompt.
- **Interaction ground truth:** where AppleScript cannot make the change, the §11 route applies (hand-written XML, then owner-authored recipes). Otherwise, for each interaction case, Excel applies the same change to the same pivot (an item's visibility, show detail, field orientation, the data-field function, sorting) and saves. The engine's layout after the same view-state change must equal Excel's saved output.
- **Slicers and timelines:** the dictionary has no slicer or timeline commands.
  - Reference files are hand-written as XML in the shape Excel writes, verified by Excel opening them without repair, and rendered by Excel for capture.
  - Interaction states (selected and unselected items, cross-filter) are pre-set in the files.
  - If hand-written files are rejected, the owner authors them in Excel's UI from agent-written recipes (§11).
- **Corpus:** fetch apache/poi `ExcelPivotTableSample.xlsx` and LibreOffice's pivot, slicer and timeline xlsx test files through `fetch-corpus.ts`. Each becomes a faithfulness-check fixture (§6.4); the saved output is their oracle.
- **Comparison:** pixel comparison of the pivot rectangle, slicers and timelines against Excel captures, with the existing meanAbs, rowMax and strong-pixel conventions. Tolerances are set from measured anti-aliasing noise and recorded in `docs/validation/xlsx-pivots.md`.
- **Engine tests:** every engine rule has unit tests (records in, layout out), plus the Excel-oracle tests above. Numbers are compared exactly or within Excel's 15-significant-digit display. **[measure]**
- **Verified:** a feature counts as verified when its references pass.
- **Unverified features:**
  - **Rendering fidelity:** an unverified feature ships drawn from the best-measured model, labelled *unverified*. **[owner]** This is the sparkline precedent.
  - **Numbers and interaction semantics never ship unverified.** An interaction whose Excel oracle does not pass is disabled, because a wrong number is an anti-goal.

## 10. Review focus: inputs most likely to bite
1. **Compact form with several row fields, collapsed items and subtotals at the top.** Indentation, subtotal placement and the oracle `rowItems` must all agree.
2. **Show values as with a base item of previous or next**, and % of parent total at several levels.
3. **Date grouping across years**, with blanks and out-of-range buckets, and a group field used together with its base field.
4. **A large cache on the paged path:** streaming, the cap, worker memory charging, and the 200 ms benchmark. A slicer click driving two pivots on different sheets.
5. **A pivot that grows into a non-empty cell** (refused), and one that shrinks (the cleared cells repaint across tile boundaries).
6. **A custom pivot style plus `formats` dxf overrides plus a cell `s=` format** on the same cell: the precedence rules.
7. **Pivot charts whose series count changes after a filter:** series formatting and the legend.
8. **Files from LibreOffice and other writers:** missing optional attributes, absent `rowItems`, absent records, inline record values instead of shared items, and `refreshOnLoad`.
9. **Calculated fields** with names that contain spaces or quotes, and division by zero inside them.
10. **Dependent formulas after a pivot change** on the resident path (recomputed) and the paged path (marked stale).

## 11. Owner answers (builder check, 2026-10-29)
- **Excel table (ListObject) styles join Stage 1. [owner]** The style engine is generic over the table and pivot element sets from the start. Built-in TableStyleLight1–21, Medium1–28 and Dark1–11 are authored alongside the pivot styles, and checked against Excel the same way. Table slicers (Stage 5) therefore show a styled table.
- **Oracle for interactions Excel cannot be scripted to make. [owner]**
  - First, agents hand-write the case as XML with `refreshOnLoad="1"`, open it in Excel and save; Excel's saved output is the oracle.
  - Cases Excel rejects, repairs or will not recompute go to the owner, who builds them in Excel's UI from a short recipe written by the agents. This applies to pivot interactions and to slicer and timeline states alike.
  - Each recipe states the source data, the pivot and the exact clicks. The owner's saved file is the oracle.
  - The plan batches these recipes so the owner gets them once per stage, not one at a time.

## 12. Delivery **[owner]**: five stages, each its own reviewed PR

1. **Faithful display.**
   - The eager detection stub; definition, style and slicer/timeline parsing; the style engine with the built-in pivot and table styles; Excel table (ListObject) styles rendered (§11).
   - Pivot cells painted with styles; buttons drawn per §7.10; pivot charts' field buttons.
   - Slicers and timelines drawn as saved (moved here from Stage 5 for display, because their fallback text boxes are the most visible defect today). **[call]**
   - Both paths. Includes the Excel reference harness for pivots.
2. **Engine.**
   - Records decode, both paths, with streaming in the worker.
   - Layout, aggregation, subtotals, totals, show-values-as, sorting, filters, grouping and calculated fields.
   - The faithfulness check, the Excel-oracle tests and the 200 ms benchmark gate (§6.3).
   - No user-visible change except diagnostics.
3. **In-table interaction.**
   - Expand and collapse, filter menus (item, label, value, date, top 10), page fields and sort.
   - Live value generation and tile invalidation, the overlap rule, pivot charts following, and dependent values (§7.5).
4. **Field list and values.** The field-list panel, moving fields, value field settings, and group and ungroup.
5. **Slicer and timeline interaction.** Clicks, multi-select, cross-filtering, timeline selection and levels, multiple pivots per slicer, and table slicers.

Each stage meets §8 and passes CI and review before the next starts. Before a stage merges, its interactions are absent and show no diagnostic (planned staging), and pivots display as saved.

## 13. Left to the builder
- The module layout under `src/xlsx/pivots/`, and the names of types and fields.
- Concrete caps (record store bytes, items per field, layout cells), sized from realistic workbooks and stated in the plan.
- The record-store encoding.
- How built-in pivot style definitions are encoded compactly in the lazy chunk.
- Per-reference tolerances, from measured noise.
- The exact DOM structure of the field list, within §6.7 and the existing control patterns.
