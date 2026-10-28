# XLSX pivot tables, tables, slicers and timelines — validation record

Spec: `docs/superpowers/specs/2026-10-29-xlsx-pivots-design.md`. Plan: `docs/superpowers/plans/2026-10-29-xlsx-pivots.md`.
Reference: Excel for Mac 16.114, UI language Traditional Chinese (zh-TW).

## Corpus (Task 1)

`scripts/fetch-corpus.ts` fetches these into `corpus/` (gitignored):
- **apache/poi** (Apache-2.0): `ExcelPivotTableSample.xlsx` (2 pivots), `ExcelTables.xlsx` and `tableStyle.xlsx`.
- **LibreOffice/core** at `6dd1649a` (MPL-2.0 test data, never committed), under `sc/qa/unit/data/xlsx/`: 18 pivot files and 3 table-style files:
  - compact, outline and tabular layouts;
  - row, column and page filters, and date filters;
  - PivotStyleDark1;
  - two data fields;
  - calculated fields;
  - date and number grouping;
  - `formats` dxfs;
  - no source data;
  - a stale cache;
  - table totals and inner borders.

No public slicer or timeline test files were found (LibreOffice `sc/qa`, `oox/qa` and the POI test data were searched on 2026-10-29). Slicer and timeline references are therefore hand-written and checked by opening them in Excel (spec §9, §11).

## Stage 1 probes (Task 1, 2026-10-29)

Reproduced by `bun scripts/xlsx-refs/probe-pivots.ts <out-dir>`, except B.

### A. Creating and reshaping pivots
- **A pivot cache cannot be created through AppleScript.** `make new pivot cache at <workbook>` returns parameter error -50 with a range object, an A1 string, an R1C1 string, or no properties. `create pivot table` needs a cache, so a new pivot cannot be made.
- **An existing pivot can be reshaped:**
  - `table style` (the property `table style2` does not apply);
  - `layout row default`;
  - pivot item `visible`;
  - data field `calculation`, written as `pivot field calculation …`, not `percent of total`;
  - `refresh table`;
  - `save workbook as`.
- **The saved file holds Excel's recomputed output** in the definition (`showDataAs="percentOfTotal"`, `pivotTableStyleInfo`) and in the cells (0.25 / 0.75 / 1). It is the test answer.
- Address a pivot by worksheet name: `pivot table 1 of worksheet "<name>"`. A `whose` filter over worksheets fails.
- **Reference route for the engine and display cases:** a seed pivot, either Excel-written (corpus) or hand-written; then AppleScript reshapes it, refreshes it and saves.
  - **Grouping has no AppleScript command.** Grouping cases need hand-written XML or owner recipes.
  - The owner recipes stay with the plan's §11 route.

### B. `refreshOnLoad`
- `refreshOnLoad="1"` was added to both caches of the POI sample, and the source value B2 changed from 1 to 100. Opened through automation (`open -a`), **Excel did not refresh**: the saved output still came from the old records.
- An explicit `refresh table` then **did recompute from the source range** (the saved cells show 100 and 103).
- So the Excel-oracle cases for spec §7.7 use explicit `refresh table`. Whether Excel refreshes when opened interactively, perhaps after a security prompt, is not established.

### C. Capture appearances
- `copy picture … appearance screen` keeps the gridlines; `appearance printer` drops them. Both are 302×113 px for A1:E8, at 1 px per point.
- **Neither appearance includes the pivot field buttons** (filter arrows, ±). `copy picture` renders cell content only. Button geometry therefore needs a window screenshot. This terminal has no Screen Recording permission: `screencapture` returns only the wallpaper. The owner must grant that permission (the agent must not), or the button geometry stays *unverified* (spec §9).
- **Captions are localised:** after a refresh Excel rewrites the default captions in its UI language (列標籤 "Row Labels", 欄標籤 "Column Labels", 總計 "Grand Total"), replacing the German captions the file was saved with. The viewer shows saved cells, so display is unaffected. Captions that the engine generates (Stage 2+) need a locale rule, which is **[measure]** in Stage 2.

### D. Hand-written slicer and timeline workbooks
- Opening the first hand-written slicer workbook (`slicer-light1.xlsx`) put up Excel's content-repair prompt. Generator fixes are in progress (Task 2, fix round 1); the result is recorded here when it is re-tested.
