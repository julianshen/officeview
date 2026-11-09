# XLSX pivots: Stage 5 owner recipes (slicers and timelines)

Owner-recorded Excel UI actions that cannot be established by changing saved XML state alone (spec §9, §11). Reference application: Microsoft Excel for Mac 16.114, zh-TW UI locale. These are procedures to record; they are not treated as measurements until the resulting workbook and observation record are saved.

## Start workbook

Use `tests/fixtures/xlsx-pivots/controls/controls-seed.xlsx`:

- `Data`: Date, Region, Product, Sales with rows (2023-12-31, East, Apples, 10), (2024-01-01, East, Bananas, 20), (2024-02-15, West, Apples, 30), (2024-03-31, West, Cherries, 40), (2024-04-01, South, Apples, 50).
- `工作表1`: `樞紐分析表1`, Region rows, Product columns, Sum of Sales, Product slicer `Slicer_Product`, Date timeline `NativeTimeline_Date`.

## R1 — Slicer click semantics and Multi-Select

Run the direct UI actions and record every intermediate selection, pivot total, and saved/reopened state. No keyboard shortcut is part of this recipe.

1. In All-selected state, click the visible Multi-Select button.
2. Click Apples, then Bananas. Record the selected items and pivot total after each click.
3. Click the Multi-Select button again, then click Cherries. Record the final selection and pivot total.
4. From a saved Apples-only state, click Apples again. Record whether it remains selected, clears, or selects all.
5. In a separate pass from Apples-only, Command-click Bananas; record the intermediate result. Command-click Bananas again and record the final result.
6. Save and reopen; record whether the toggle's pressed appearance persists. Do not infer click behavior from cache XML or an undocumented shortcut.

Cases: `s5-slicer-click-selected`, `s5-slicer-modifier-add-remove`, and `s5-slicer-multiselect-toggle`. Their XML-cache equivalents are not accepted as click evidence.

## R2 — Slicer on a manually filtered pivot field (measured XML/refresh case)

Run `bun scripts/xlsx-refs/build-control-interactions.ts` to rebuild `s5-slicer-manual-filter-vs-slicer.xlsx`. The prepared input hides Cherries (Product item x=2) in the pivot's column field while the slicer cache selects x=0,1,2. Excel for Mac 16.114 refresh/save removed the hidden flag from x=2, kept the slicer selection, and set `h="1"` on the no-data/missing x=3 item. Cherries remains in the saved pivot and the grand total is 150. The output and complete-cell oracle are checked in; the run log is `s5-build-run3.log`. This is a measured saved-XML reconciliation, not evidence of a manual UI gesture or a general filter-intersection rule.

## R3 — Report Connections with distinct layouts

1. Duplicate the pivot sheet twice. Keep three distinct layouts: Pivot 1 Region rows × Product columns; Pivot 2 Date rows × Product columns; Pivot 3 Product rows with Sales values and no column field.
2. Record each pivot's complete occupied cell map before the selection.
3. Open Report Connections for `Slicer_Product`. Connect Pivot 1 and Pivot 2; leave Pivot 3 unchecked.
4. Select Apples. Record the complete cell maps of all three pivots, the cache target list, and actual pivot names. Verify both linked layouts respond and the unlinked pivot's complete cell map is unchanged.

Case: `s5-linked-two-pivots`. A same-layout duplicate is not sufficient evidence.

## R4 — Table Slicer while retaining AutoFilter

No Excel-authored table-slicer template is checked in. A hand-written `tableSlicerCache` package produced a repair prompt in Excel 16.114, so this case is an owner recipe and is excluded from the XML builder.

1. Open `tests/fixtures/xlsx-pivots/controls/controls-seed.xlsx` in Excel.
2. On `Data`, select `A1:D6`, choose **Insert > Table**, and confirm that the table has headers. Verify the resulting table name is `Table1`.
3. With a Table1 cell selected, choose **Table Design > Insert Slicer**, select **Region**, and create the slicer.
4. Use Table1's Product AutoFilter to keep only **Apples**.
5. In the Region slicer, click **East**, then Command-click **South**. The intended intersection is East Apples and South Apples.
6. Save as `tests/fixtures/xlsx-pivots/controls/s5-table-slicer.xlsx`, close the workbook, then reopen that exact file.
7. Record the visible source rows and verify the saved AutoFilter still selects Apples. The expected rows are 2 and 6 visible (East/South Apples); rows 3, 4 and 5 are hidden. Also record the table-slicer cache's selected Region items and the Excel version.

If Excel cannot create or reopen the combined state, record the actual result and leave the case unmeasured; do not synthesize table-slicer parts or row-hidden flags in XML.

Case: `s5-table-slicer` (owner recipe; unmeasured until the saved workbook and observation are recorded).

## R5 — Timeline scroll

With All Periods selected, scroll to the 2023 range. Record the first visible period and whether the committed selection changes. Writing `scrollPosition` in XML does not prove UI scrolling.

Case: `s5-timeline-scroll`.

## R6 — Timeline forward and reverse drag

From All Periods, drag Jan→Mar, save and record normalized endpoints and the complete pivot cell map. Restore All Periods, drag Mar→Jan, and record the same fields. Compare both saved outcomes; a single XML date range does not establish drag normalization.

Case: `s5-timeline-drag-forward-reverse`.

## R7 — Cross-filter order

Create a real Region slicer connected to the pivot, select East, and let Excel refresh. Record the actual Region and Product cache parts, Product item order, selected/no-data flags, and complete pivot cell map. This recipe is pending: the current seed has no Region slicer cache, so adding a cache part without the real control connection or editing the Product cache is not evidence.

Case: `s5-slicer-crossfilter-order`.

## Automation boundary

The builder automates R2 by preparing package XML and asking Excel to recalculate and save. R4 is owner-recorded because the hand-written table-slicer package was rejected by Excel and there is no checked-in Excel-authored template. R1 and R3 require direct Excel UI actions because saved state alone cannot prove click semantics or Report Connections dialog actions. R5–R7 remain owner-recorded. Do not use coordinate-driven System Events as a substitute for a stable Excel object-model action.
