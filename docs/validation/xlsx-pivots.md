# XLSX pivot tables, tables, slicers and timelines — validation record

Spec: `docs/superpowers/specs/2026-10-29-xlsx-pivots-design.md`. Plan: `docs/superpowers/plans/2026-10-29-xlsx-pivots.md`.
Reference: Excel for Mac 16.114, UI language Traditional Chinese (zh-TW).

## Eager budget (Task 4)

`src/index.ts` statically imports the shared `office-worker` chunk containing `parse.ts` and `pivots/stub.ts`, so the stub is loaded eagerly. The repository's `eager-budget.ts` convention measures only `dist/index.js`; it does not include that static dependency. Using the same Node zlib gzip measurement for both artifacts, base `b7fff618` was **35,954 B** for `dist/index.js` and **365,128 B** for the `office-worker` chunk. At Task 4 HEAD `359eba83`, the figures were **35,955 B** and **366,084 B**. After fix round 1, they are **35,957 B** and **366,813 B**. Final deltas versus `b7fff618`: `dist/index.js` **+3 B**; `office-worker` **+1,685 B**. The existing index-only eager ceiling remains 38,720 B; Task 4's index delta remains within 1,200 B. Content-hashed chunk names can cause small run-to-run size variance.

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

## Stage 1 references and measurements (Task 3, 2026-10-29, Excel 16.114)

### How the references are made
- **Pivots (96 cases).** `recipes/pivot-seed.xlsx` was created once through Excel's Data menu (Summarize with PivotTable). It has rows Region › Product, column Year, values Sum of Sales and Sum of Units, compact form, Calibri 11, and English captions set through AppleScript. `scripts/xlsx-refs/build-pivots.ts` opens the seed for each case, reshapes it through the AppleScript pivot dictionary, refreshes and saves, then checks the saved `pivotTableStyleInfo` against the case.
  - The pivot style must be set through `table style2`; `table style` reads back but is not saved.
  - The first seed was made in Excel's zh-TW defaults (新細明體 12, Chinese captions) and was rebuilt.
  - Two strings stay localized because Excel ignores writes to them:
    - the column grand-total headers ("Sum of Sales 的加總");
    - the page field's "(全部)" (All).
- **Custom style, precedence and cell-override inputs.** `scripts/xlsx-refs/style-inputs.ts` derives these from the seed, and `build-pivots.ts --measure` has Excel re-save them.
- **Tables, slicers and timelines** come from the XML generator. Excel opens them without repair; the totals-row table needed `totalsRowLabel`/`totalsRowFunction` and `SUBTOTAL` cells.
- **Captures.** `scripts/xlsx-refs/capture-all.ts` uses `copy picture`, screen appearance, 1 px per pt.
  - A pivot is captured from A1 to the bottom-right corner of its saved `location`.
  - `copy picture` includes slicer and timeline shapes and the ± buttons. It leaves out field dropdown buttons, which were measured from a window screenshot at 100% zoom (`refs/measure-field-buttons.png`); the screen is also 1 px per pt.
  - Screen gridlines give the cell geometry, so pivots have no bordered twins.
- **Duplicate check.** Identical images across cases are expected only where a case equals the default (table flag defaults, `pivot-layout-compact`, `pivot-subtotals-top` and `pivot-light-16`, and light1/months).
  - The check caught timeline "Other1/2". These are not Excel timeline styles (Excel has Light1–6 and Dark1–6) and were removed.
  - It also caught slicer variants hidden under the timeline; the controls were re-laid out.

### What Excel writes into pivot cells' `s=` (spec §3)
- **Neutral formats only.** No pivot-style colour or font is ever written into cells. There are four formats:
  - `pivotButton="1"`, on the cells that show a field dropdown;
  - left-aligned, for item labels;
  - left-aligned with `indent=<level>`, for compact-form child items;
  - plain, for values.
- **Field buttons.** Compact form has them on the Row Labels and Column Labels cells; outline and tabular forms have one on each field header.
- **Indent.** The compact-form indent is stored as the cell's alignment `indent`; outline and tabular forms write none.
- **Formats override.** A format applied to a pivot cell is written twice: as a pivot `<format dxfId>` with a `pivotArea`, and as the cell's own `s=` (fill applied).

### A cell's own `s=` over the pivot style (spec §6.5 [measure] — confirmed)
- `measure-cell-override` gives three cells a fill and a bold red font through `s=` alone, with no `formats` entry: a subheading label (A6), a value (C7) and a grand-total cell (B22).
- Excel draws the cell's fill and font over the pivot style in all three. The style still supplies any property the cell's format does not set.

### `applyFill` / `applyFont` over a table style (Task 5; `measure-apply-flags`, TableStyleMedium2)
- `measure-apply-flags.xlsx` is `table-medium-2` with extra xfs: a green fill or a bold red font with the apply flag absent, `"0"` and `"1"`, plus a `none` fill with `applyFill="1"` on a banded row.
- Excel draws the own fill and font in every flag case, including the header row: **it ignores the apply flags when it layers a cell over a table or pivot style.**
- A `none` fill with `applyFill="1"` keeps the band fill, so an own fill wins only when its pattern is not `none`. Cells on xf 0 (fontId 0) keep the header's white bold, so an own font wins only when it is not the Normal font.
- The generated table fixtures do not link `xl/styles.xml` (no workbook relationship and no content-type override). Excel then ignores the stylesheet and repairs any cell with `s` above 0. This workbook adds both links.

### Element precedence (spec §6.5) — `measure-precedence`
- **Method.** Custom style `OfficeViewPrecedence` gives every element a distinct fill; the colours are in `recipes/pivot-precedence-colors.json`. All style-info flags are on. The winner is read at each cell; `>` means "wins where both apply".
- **Header rows:**
  - firstHeaderCell (A3:A4) > headerRow;
  - firstColumnSubheading and secondColumnSubheading (column-item rows 4 and 5) > headerRow;
  - on row 3, headerRow > lastColumn;
  - on rows 4–5, lastColumn > column subheadings.
- **Body:**
  - totalRow > lastColumn and firstColumn (on the whole grand-total row);
  - lastColumn > firstRowSubheading > firstColumn > row stripes;
  - firstColumn > row stripes;
  - row stripes > column stripes, which never show.
  - wholeTable is beaten everywhere.
- **Stripes.** Subheading rows count toward the stripe alternation: the first item row under a subheading takes secondRowStripe.
- **Not shown by this layout:** subtotal rows and columns, and blankRow. They need a subtotals-bottom or blank-line variant if Stage 2 relies on them.

### Compact indent and buttons (screen, 100%, Calibri 11)
- **Compact text position.** With ± buttons shown, label text starts at **15 + 10 × level px** from the cell's left edge: level 0 at 15, level 1 at 25, level 2 at 35 (`pivot-values-on-rows`).
- **± button.** A 7 × 7 px box at x = 6–12, vertically centred in a 15 px row, drawn for the item levels that have children.
- **Field dropdown button** (Row Labels, Column Labels and page field), measured from `refs/measure-field-buttons.png`. This is a window screenshot of `pivot-page-field.xlsx` at 100% on a 1× display, cropped to A1:I10; the sidecar records the method. In the crop the Column Labels button box is at x 154–168, y 30–43:
  - box: 15 × 14 px, flush with the cell's right and bottom edges, fill #FDFDFD, 1 px border #738A93;
  - arrow: a downward triangle 7 px wide and 4 rows tall (7, 5, 3, 1), centred, starting 5 px below the box top, colour ≈ #828282.

### Slicer geometry (`slicer-light1`, SlicerStyleLight1, 1 px per pt)
- **Frame.** 1 px border #4F81BD on all sides.
- **Header.** The caption ("Product", bold) sits on white. A 1 px #4F81BD separator lies 19 px below the top border, and the first item starts 4 px below it, so the top border to the first item's top is 23 px. The multi-select and clear-filter icons sit at the header's right.
- **Items.**
  - 17 px tall including a 1 px border, with a 2 px gap, so the pitch is 19 px;
  - inset 6 px from the frame on the left and right;
  - text starts 8 px inside the item.
- **Item states:**
  - selected with data: fill #B8CCE4, border #9A9A9A;
  - not selected with data: fill white, border #CCCCCC;
  - no data (`slicer-no-data`, `showNoDataItems="1"`): grey text, shown after the items with data.
- **Scrolled** (`slicer-scrolled`, 90 pt tall, 4 items, `startItem=2`). Excel draws a vertical scrollbar (a grey rounded thumb) at the right edge. It also **clamps `startItem`** so the last item fills the view: with 3 visible rows of 4 items, it shows items 1–3 (Dates, Cherries, Bananas), not from item 2.

### Timeline geometry (120 pt anchor; `timeline-light1-months`, `timeline-dark2-days`)
- **Frame.** 1 px border in the style's accent colour (Light1 #5887BF, Dark2 #C25956); the frame fills the anchor (121 px rows 105–225 in the capture).
- **Header.** The caption ("Date") and the clear-filter icon sit at the top. A 1 px accent separator lies 22 px below the top border.
- **Rows below the header, in order:**
  - the selection label ("All Periods") at left, with the level picker at right;
  - the year label;
  - the period labels (month or day numbers);
  - the time bar.
- **Time bar.** 13 px tall, filled with a vertical gradient from the style's light tint to its dark shade (Light1 #B9CEE5 → #5489C2; Dark2 #C3544B → #6A2521) and divided per period. It is inset 6 px from the frame on the left (x 456). With no selection (All Periods, as in every case) no selection handle is drawn; see "Slicer and timeline display details".
- **Scrollbar.** 7 px tall, 5 px below the bar; grey thumb #BCBCBC on a pale accent track.
- **Text.** Labels are drawn in Excel's UI language, not stored in the file: the zh-TW captures show 所有週期, 月 and 天, and 1月…. Comparisons must mask timeline text or treat it as locale-dependent.

### Slicer and timeline display details (Task 12, from the same captures)
- **The captures use the Office 2007–2010 palette.** The generator used to write a `theme1.xml` with one entry in each `fmtScheme` style list (the schema requires three) and all six accents set to #4472C4. Excel discarded it and drew with the default theme of `defaultThemeVersion="124226"`: accent1 #4F81BD, accent2 #C0504D, accent3 #9BBB59, accent4 #8064A2, accent5 #4BACC6, accent6 #F79646. The slicer frames of Light1–6 and Dark1–6 show exactly these, and so does the pivot in the same captures (PivotStyleLight16 header #DCE6F1). Since the generator batch (b3dfa74f, Ruling AJ), the slicer and timeline cases embed that theme, recipe R1's Excel-saved theme. The re-captures kept every slicer, timeline and pivot colour.
- **Slicer items** have rounded corners of about 3 px. The no-data item (`slicer-no-data`) has a #E0E0E0 border and #828282 text.
- **Slicer scrollbar** (`slicer-scrolled`, frame x 186–325, y 0–89): the items narrow to 192–305, 20 px inside the frame's right edge. The #BCBCBC thumb is 8 px wide at x 310–317 (8 px inside the right edge), rounded, and runs y 39–83. Its track is the item area, y 23 to 6 px above the bottom border; the thumb is the visible share (3 of 4 rows).
- **Timeline bar** (frame x 450–649, y 105–225): x 456–642 (6 px in on the left, 7 px on the right), y 183–195. The dividers are 1 px in accent tint 0.4 (#95B3D7 Light1, #DA9694 Dark2). Period pitch: months 39 px (dividers at 495, 534, 573 and 612), days 21 px (476, 497, 518…), quarters about 66 px (521, 588), a single year block 38 px (456–493).
- **Timeline selection handles: none for All Periods.** In the re-captures (generator b3dfa74f) none of the 48 timeline references, nor the timeline in the 17 slicer references, has a handle.
  - The handle strips are now background: x 451–455 left of the bar, and for years x 494–498 after the 2024 block. Across all 65 images nothing else differs from the earlier captures: not the bar (it still starts at x 456), labels, dividers, scrollbar, slicer or pivot.
  - The earlier captures drew 5 px grips just outside the whole range, as tall as the bar: left at 451–455, and for years the right one at 494–498. They were textured, not flat: a #C1C1C1 → #6C6C6C ramp with #DDDDDD grip lines and a #9E9E9E centre line (y 190). Those files are kept in `.superpowers/sdd/2026-10-29-xlsx-pivots/t12-evidence/*.old.png`.
  - Cause: in a control workbook only two parts changed between the two captures — `theme1.xml` (invalid → R1's valid theme) and the removal of the non-ECMA table-level `subtotalTop` from `pivotTable1.xml`. The worksheets, styles link, caches, slicer and timeline parts are byte-identical, so the date-format change in the batch does not reach these files. The timeline cache's `state` has no `<selection>` in either version. Which of the two changes Excel reacted to is not isolated; that needs two captures, one change each.
  - Rule: handles mark an explicit selection only (the cache state's `<selection>`). The renderer draws their size and place from the earlier captures [estimate]. No reference has an explicit selection.
- **Timeline scrollbar:** the track is 5 px of accent tint 0.8 (#DAE4F1 Light1, #F1DBDA Dark2) at y 202–206, under the whole bar. The thumb is 7 px (y 201–207): #AFB3B7 with #BCBCBC top and bottom rows, starting at the bar's left (456–528 for months). It is absent when every period fits (years).
- **Timeline label colours** (the darkest glyph pixel): year text #989898, which also covers the period labels at the years level; month, quarter and day labels #909090; the level picker about #A1A1A1. The selection label is an accent shade (#577AA4 Light1, #A75756 Dark2, #E47820 Dark6), which the thin zh-TW glyphs do not pin down.

### Cell indent (Task 5 input; `measure-indent`, Calibri 11, 1 px per pt)
- **Set-up.** Cells G1:G4 hold the same text with alignment `indent` 0, 1, 2 and 3, set through AppleScript `indent level`. Text starts at x = 4, 10, 16 and 22 px.
- **Result.** One indent level is **6 px**, one maximum-digit width at this scale, not 3 × MDW as the plan guessed. Ruling: the renderer indents `level × MDW` in its own pixel space.
- **Pivots.** The 10 px per level measured in compact pivots is the pivot's own layout, a separate measurement.

### Pivot field dropdowns, follow-up (Task 10; window screenshot of `pivot-layout-outline`, 100%, 1×)
- **Values header.** Outline form with Values on the column axis: Year, Region and Product show a dropdown; the Values header shows none, although Excel writes the `pivotButton` format on that cell too.
- **Box height follows the row.** Here the Year box is 15 × 15 px (border y 347–361, x 470–484; arrow rows 6 px below the top border). In `measure-field-buttons` it is 15 × 14. The box height is the smaller of 15 px and the row's pixel height.

### Pivot chart field buttons (Task 11; `pivot-chart-buttons-*`)
- **How the files are made.** `scripts/xlsx-refs/build-pivot-charts.ts`. The seed `recipes/pivot-chart-seed.xlsx` is `pivot-layout-compact.xlsx` with a clustered column chart over the pivot range A3:C10, made through AppleScript (`make new chart object`, `set source data`). Excel saves it as a PivotChart: `c:pivotSource`, `c14:pivotOptions` (`dropZoneFilter`, `dropZoneCategories`, `dropZoneData`, `dropZoneSeries`, `dropZonesVisible`) and `c16:pivotOptions16` (`showExpandCollapseFieldButtons`). Seven variants edit those flags and Excel re-saves each; Excel writes only the flags that are on.
- **Result.** Excel 16.114 for Mac draws no field buttons on a PivotChart, whatever the flags (window screenshot `refs/pivot-chart-buttons-all-window.png`, every flag on). copy picture omits them as well.
- **Ruling (owner, 2026-10-29).** Match Mac Excel: the viewer parses the pivot link and the flags, draws no chart field buttons and leaves the plot area as for any chart. Windows-style buttons wait for a Windows reference.
