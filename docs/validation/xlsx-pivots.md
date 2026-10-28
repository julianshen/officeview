# XLSX pivot tables, tables, slicers and timelines — validation record

Spec: `docs/superpowers/specs/2026-10-29-xlsx-pivots-design.md`. Plan: `docs/superpowers/plans/2026-10-29-xlsx-pivots.md`.
Reference: Excel for Mac 16.114, UI language Traditional Chinese (zh-TW).

## Corpus (Task 1)

`scripts/fetch-corpus.ts` fetches these into `corpus/` (gitignored):
- **apache/poi** (Apache-2.0): `ExcelPivotTableSample.xlsx` (2 pivots), `ExcelTables.xlsx` and `tableStyle.xlsx`.
- **LibreOffice/core** at `6dd1649a` (MPL-2.0 test data, never committed), under `sc/qa/unit/data/xlsx/`:
  - `pivot/Pivot2_Row_Compact.xlsx`
  - `pivot/Pivot2_Row_Grand_Subtotals.xlsx`
  - `pivot/Pivot3_Column_Grand_Subtotals.xlsx`
  - `pivot/Pivot4_Column_Grand_Subtotals_SortDescending.xlsx`
  - `pivottable_outline_mode.xlsx`
  - `pivottable_tabular_mode.xlsx`
  - `pivottable_rowcolpage_field_filter.xlsx`
  - `pivottable_date_field_filter.xlsx`
  - `pivot_dark1.xlsx`
  - `pivot-table/pivotcompact.xlsx`
  - `pivot-table/two-data-fields.xlsx`
  - `pivot-table/calcfields.xlsx`
  - `pivot-table/shared-dategroup.xlsx`
  - `pivot-table/shared-numgroup.xlsx`
  - `pivot-table/PivotTableFormatsSubTotals.xlsx`
  - `pivot-table/PivotTableCellFormatsTest_7_TwoRowTwoColumnFields_DataColor.xlsx`
  - `pivot-table/PivotTableWithNoSourceData.xlsx`
  - `PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithCacheData.xlsx`
  - `TableStyleTest.xlsx`
  - `tableStyleInnerBorders.xlsx`
  - `tdf162963_TableWithTotalsEnabled.xlsx`

No public slicer or timeline test files were found (LibreOffice `sc/qa`, `oox/qa` and the POI test data were searched on 2026-10-29). Slicer and timeline references are therefore hand-written and checked by opening them in Excel (spec §9, §11).

## Stage 1 probes (Task 1, 2026-10-29)

Reproduced by `bun scripts/xlsx-refs/probe-pivots.ts <out-dir>`, except B.

### A. Creating and reshaping pivots
- **The prescribed creation probe was run.**
  - A 50-row source (Region, Product, Quarter, Sales; XlsxWriter, seed 7) and the specified pivot (Region and Product on rows, Quarter on columns, Sum of Sales, PivotStyleMedium9) were attempted through AppleScript.
  - It failed at the first step, creating the cache (below). `probe-pivots.ts` repeats that failing step on the POI sample.
- **A pivot cache cannot be created through AppleScript.** `make new pivot cache at <workbook>` returns parameter error -50 with a range object, an A1 string, an R1C1 string, or no properties. `create pivot table` needs a cache, so a new pivot cannot be made.
- **An existing pivot can be reshaped:**
  - the pivot style, through `table style2`. Setting `table style2` reads back as `missing value` but is what Excel saves in `pivotTableStyleInfo`. `table style` reads back the new name but is not saved (corrected 2026-10-29 during Task 3; the first probe only read the value back);
  - `layout row default`;
  - pivot item `visible`;
  - data field `calculation`, written as `pivot field calculation …`, not `percent of total`;
  - `refresh table`;
  - `save workbook as`.
- **The saved file holds Excel's recomputed output** in the definition (`showDataAs="percentOfTotal"`, `pivotTableStyleInfo`) and in the cells (0.25 / 0.75 / 1). It is the test answer.
- Address a pivot by worksheet name: `pivot table 1 of worksheet "<name>"`. A `whose` filter over worksheets fails.
- **Creating a pivot through UI scripting does work**, as found while building recipe R1. The steps:
  1. Select a cell in the source and click Data → 以樞紐分析表摘要 ("Summarize with PivotTable") through System Events.
  2. Confirm the dialog.
  3. Add fields with `add fields to pivot table … row fields {…}`.
  4. Make a data field with `set pivot field orientation of pivot field "<name>" to orient as data field`. `add data field` silently did nothing.
  5. Slicers and timelines are inserted through the PivotTable Analyze ribbon buttons (插入交叉分析篩選器 / 插入時間表) and their field-checkbox dialogs.

  This needs the macOS Accessibility and Screen Recording grants, which the owner gave on 2026-10-29.
- **Reference route for the engine and display cases:** a seed pivot, either Excel-written (corpus) or hand-written; then AppleScript reshapes it, refreshes it and saves.
  - **Grouping has no AppleScript command.** Grouping cases need hand-written XML or owner recipes.
  - The owner recipes stay with the plan's §11 route.

### B. `refreshOnLoad`
- `refreshOnLoad="1"` was added to both caches of the POI sample, and the source value B2 changed from 1 to 100. Opened through automation (`open -a`), **Excel did not refresh**: the saved output still came from the old records.
- An explicit `refresh table` then **did recompute from the source range** (the saved cells show 100 and 103).
- So the Excel-oracle cases for spec §7.7 use explicit `refresh table`. Whether Excel refreshes when opened interactively, perhaps after a security prompt, is not established.

### C. Capture appearances
- `copy picture … appearance screen` keeps the gridlines; `appearance printer` drops them. Both are 302×113 px for A1:E8, at 1 px per point.
- **Neither appearance includes the pivot field buttons** (filter arrows, ±). `copy picture` renders cell content only, so button geometry needs a window screenshot. The owner granted Screen Recording on 2026-10-29, and `screencapture` now works while the display is awake and unlocked.
- **Captions are localised:** after a refresh Excel rewrites the default captions in its UI language (列標籤 "Row Labels", 欄標籤 "Column Labels", 總計 "Grand Total"), replacing the German captions the file was saved with. The viewer shows saved cells, so display is unaffected. Captions that the engine generates (Stage 2+) need a locale rule, which is **[measure]** in Stage 2.

### D. Hand-written slicer and timeline workbooks
- Opening the first hand-written slicer workbook (`slicer-light1.xlsx`, from generator 604ef109) put up Excel's content-repair prompt. After repair, Excel opened it marked "Repaired" with the pivot table and the slicer **both removed**: the Report sheet had 0 pivot tables and 0 shapes, and only the cell data was left.
- **Re-test, 2026-10-29 (generator 5d95b657, Excel 16.114):** `slicer-light1`, `slicer-two-column`, `slicer-no-data`, `timeline-dark1-days` and `timeline-dark1-months` all open with no repair prompt, and the pivot, Product slicer and Date timeline load. A part-by-part bisect against recipe R1 (template parts swapped in with ids normalised, one group at a time) found what Excel needs from a hand-written cache:
  - Version attributes on the cache root (`createdVersion="8" refreshedVersion="8" minRefreshableVersion="3"`) and the pivot root (`createdVersion`/`updatedVersion="8"`, `minRefreshableVersion="5"`).
  - `sharedItems` type flags that agree with the items. A date field is `containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0"` with `minDate`/`maxDate`. Numeric fields carry no shared items (numbers inline as `<n>` in the records, `minValue`/`maxValue` on `sharedItems`).
  - No `slicerData` attribute on `x14:pivotCacheDefinition`. A single-variable test showed that this attribute alone makes Excel repair the file and drop the slicer.
  - No `refreshOnLoad`. This one was not isolated; it is dropped because R1 doesn't have it.
- A template pivot definition paired with a generated cache also makes Excel repair the file, because the field-item order differs. Bisects must swap the pivot definition and its cache together.
- The timeline cases shipped with the template's `scrollPosition` and bounds and without a `style` attribute (fix round 4). A hand fix (`style="TimeSlicerStyleDark1"`, scroll position and bounds taken from the data) opens clean.
- **Capture route:** still untested; Task 3 records it.
- **Recipe R1** (`tests/fixtures/xlsx-pivots/recipes/r1-slicer-timeline.xlsx`) is an Excel-saved pivot with a Product slicer (Apples and Cherries selected) and a Date timeline (all periods), built through UI scripting.
- An Excel-authored slicer and timeline template (recipe R1, built by the orchestrator through Excel UI scripting) is the structural reference for them.
