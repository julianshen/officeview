# XLSX pivot tables, tables, slicers and timelines — validation record

Spec: `docs/superpowers/specs/2026-10-29-xlsx-pivots-design.md`. Plan: `docs/superpowers/plans/2026-10-29-xlsx-pivots.md`.
Reference: Excel for Mac 16.114, UI language Traditional Chinese (zh-TW).

## Eager budget (Task 4)

`src/index.ts` statically imports the shared `office-worker` chunk containing `parse.ts` and `pivots/stub.ts`, so the stub is loaded eagerly. The repository's `eager-budget.ts` convention measures only `dist/index.js`; it does not include that static dependency. Using the same Node zlib gzip measurement for both artifacts, base `b7fff618` was **35,954 B** for `dist/index.js` and **365,128 B** for the `office-worker` chunk. At Task 4 HEAD `359eba83`, the figures were **35,955 B** and **366,084 B**. After fix round 1, they are **35,957 B** and **366,813 B**. Final deltas versus `b7fff618`: `dist/index.js` **+3 B**; `office-worker` **+1,685 B**. The existing index-only eager ceiling remains 38,720 B; Task 4's index delta remains within 1,200 B. Content-hashed chunk names can cause small run-to-run size variance.

## Eager budget (Stage 1, final)

Measured after the final-review fix batch (2026-10-30), base `ba8abf83` (the branch base) against the fix-batch head, each built with `OV_MODULE_MAP=1 bun run build` in its own tree:

| Artifact | Measure | Base `ba8abf83` | Head | Growth |
|---|---|---|---|---|
| `dist/index.js` | Node zlib (`eager-budget.ts`) | 35,954 B | 35,969 B | +15 B |
| `office-worker-*.js` | Node zlib, default level | 365,128 B | 371,876 B | **+6,748 B** |
| `office-worker-*.js` | `gzip -9` | 361,679 B | 368,467 B | **+6,788 B** |
| `office-worker-*.js` | vite report | 365.15 kB | 371.89 kB | +6.74 kB |

Ruling N: `index.js` statically imports the shared `office-worker` chunk, so its growth is eager too, though `eager-budget.ts` measures `index.js` only (35,969 B, under the 38,720 B ceiling). The Stage 1 growth of the shared chunk is **+6.7 kB gzip** across the branch, not only Task 4's +1,685 B above; flag it in the Stage 1 PR. The table-style built-ins and the styles-extras parser stay out of it and out of the paged worker chunks' static closure (`check-eager.ts` checks both).

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
- **Field dropdown button** (Row Labels, Column Labels and page field), measured from `refs/measure-field-buttons.png`. This is a window screenshot of `pivot-page-field.xlsx` at 100% on a 1× display, cropped to A1:I10; the sidecar records the method. In the crop the Column Labels button box is at x 154–168, y 29–43 (border rows included):
  - box: 15 × 15 px, flush with the cell's right and bottom edges, fill #FDFDFD, 1 px border #738A93;
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
- **Timeline bar** (frame x 450–649, y 105–225): x 456–642 (6 px in on the left, 7 px on the right), y 183–195. The dividers are 1 px in accent tint 0.4 (#95B3D7 Light1, #DA9694 Dark2). Period pitch: months about 39 px (dividers at 495, 533, 573 and 612; Task 14 corrected 534 to 533), days 21 px (476, 497, 518…), quarters about 66 px (521, 588), a single year block 38 px (456–493). Task 14 found that the spacing follows calendar length (see "Stage 1 comparison").
- **Timeline selection handles: none for All Periods.** In the re-captures (generator b3dfa74f) none of the 48 timeline references, nor the timeline in the 17 slicer references, has a handle.
  - The handle strips are now background: x 451–455 left of the bar, and for years x 494–498 after the 2024 block. Across all 65 images nothing else differs from the earlier captures: not the bar (it still starts at x 456), labels, dividers, scrollbar, slicer or pivot.
  - The earlier captures drew 5 px grips just outside the whole range, as tall as the bar: left at 451–455, and for years the right one at 494–498. They were textured, not flat: a #C1C1C1 → #6C6C6C ramp with #DDDDDD grip lines and a #9E9E9E centre line (y 190). Those files are kept in `.superpowers/sdd/2026-10-29-xlsx-pivots/t12-evidence/*.old.png`.
  - Cause: capture-session state, not the file (see "Timeline selection handles" below: the old fixture re-captured, and theme-only and subtotalTop-only variants, all match the new references).
  - Rule (Ruling AK): static timelines never paint selection handles, with or without a selection.
- **Timeline scrollbar:** the track is 5 px of accent tint 0.8 (#DAE4F1 Light1, #F1DBDA Dark2) at y 202–206, under the whole bar. The thumb is 7 px (y 201–207): #AFB3B7 with #BCBCBC top and bottom rows, starting at the bar's left (456–528 for months). It is absent when every period fits (years).
- **Timeline label colours** (the darkest glyph pixel): year text #989898, which also covers the period labels at the years level; month, quarter and day labels #909090; the level picker about #A1A1A1. The selection label is an accent shade (#577AA4 Light1, #A75756 Dark2, #E47820 Dark6), which the thin zh-TW glyphs do not pin down.

### Cell indent (Task 5 input; `measure-indent`, Calibri 11, 1 px per pt)
- **Set-up.** Cells G1:G4 hold the same text with alignment `indent` 0, 1, 2 and 3, set through AppleScript `indent level`. Text starts at x = 4, 10, 16 and 22 px.
- **Result.** One indent level is **6 px**, one maximum-digit width at this scale, not 3 × MDW as the plan guessed. Ruling: the renderer indents `level × MDW` in its own pixel space.
- **Pivots.** The 10 px per level measured in compact pivots is the pivot's own layout, a separate measurement.

### Pivot field dropdowns, follow-up (Task 10; window screenshot of `pivot-layout-outline`, 100%, 1×)
- **Values header.** Outline form with Values on the column axis: Year, Region and Product show a dropdown; the Values header shows none, although Excel writes the `pivotButton` format on that cell too.
- **Box size.** Here the Year box is 15 × 15 px (border y 347–361, x 470–484; arrow rows 6 px below the top border), the same as in `measure-field-buttons` (re-counted: border y 29–43; the first reading of 15 × 14 left out the top border row). Both pivots use 15 pt rows, so the box fills the row. A box in a shorter row is not measured; the renderer uses the smaller of 15 px and the row's pixel height.

### Pivot chart field buttons (Task 11; `pivot-chart-buttons-*`)
- **How the files are made.** `scripts/xlsx-refs/build-pivot-charts.ts`. The seed `recipes/pivot-chart-seed.xlsx` is `pivot-layout-compact.xlsx` with a clustered column chart over the pivot range A3:C10, made through AppleScript (`make new chart object`, `set source data`). Excel saves it as a PivotChart: `c:pivotSource`, `c14:pivotOptions` (`dropZoneFilter`, `dropZoneCategories`, `dropZoneData`, `dropZoneSeries`, `dropZonesVisible`) and `c16:pivotOptions16` (`showExpandCollapseFieldButtons`). Seven variants edit those flags and Excel re-saves each; Excel writes only the flags that are on.
- **Result.** Excel 16.114 for Mac draws no field buttons on a PivotChart, whatever the flags (window screenshot `refs/pivot-chart-buttons-all-window.png`, every flag on). copy picture omits them as well.
- **Ruling (owner, 2026-10-29).** Match Mac Excel: the viewer parses the pivot link and the flags, draws no chart field buttons and leaves the plot area as for any chart. Windows-style buttons wait for a Windows reference.

### Timeline selection handles (Task 12 follow-up; `measure-timeline-selection`)
- **No handles in a static capture.** The original timeline references showed a grip at the selection edge, and later captures don't. Re-capturing the *old* fixture today gives the new image, and so do variants that only restore the old theme or the table-level `subtotalTop`. So the grip came from Excel's state during the first capture session, not from the file.
- **Explicit selection.** `measure-timeline-selection.xlsx` is `timeline-light1-months` with `<selection startDate="2024-03-01" endDate="2024-05-31"/>` (`filterType="dateBetween"`). Excel draws the selected months as the style's bar, the unselected months grey and the caption "2024 3月 - 5月", and still no handles.
- **Explicit-selection paint, measured** (`refs/measure-timeline-selection.png`, frame x 450–649, y 105–225):
  - Unselected months (x 456–494, 496–532) are flat #DADADA, y 183–195.
  - Selected months keep the style's vertical gradient; its centre row (y 190) is #8FB1D6, the same as All Periods. The renderer's flat [approx] fill for light timeline styles is accent tint 0.37, which gives #90B0D5 here and #C0D496 for Light3's #C1D396.
  - The 1 px divider closing a period follows that period: background (#FFFFFF) after an unselected month (x 495, 533) and accent tint 0.4 (#95B3D7) after a selected one (573, 612).
  - Label: the ink starts at x 456 (6 px in) and ends on row 142, so its baseline is 38 px below the top border; the text is year first. Its darkest pixel, #4B709D, is accent tint −0.25 (#376092) at one coverage of 90 % over white. The All Periods label's #577AA4 is the same colour at 84 %, and Dark6's #E47820 is #E36C09 at 90 %, so the label colour is accent tint −0.25.
  - The clear-filter icon turns red (filtered).
  - The scrollbar differs from the All Periods captures: track only at x 529–533, thumb 456–528 with its left part 456–486 in #BCBCBC. This is not modelled; it is the same capture-session caveat as the handles until a second capture agrees.

### Gridlines over pivots and styled tables (Task 14)
Found by the Stage 1 comparison (`scripts/pivot-compare.ts`); the viewer used to stroke every gridline over every cell. Sampled from the copy-picture references, where Excel's gridline is #D4D4D4:
- **No gridlines inside a pivot's location.** `pivot-light-16` (location A3:I22): in the white body rows (y 90–104, row 7) every inner column edge (x 77, 170, 239 … 620) is #FFFFFF, and so is every inner row edge (x 300 at y 105, 120 …). The outer edges keep theirs: x 0 and x 728 are #D4D4D4 beside the white body rows. `pivot-light-1`'s inner lines are the style's own #D9D9D9 and #808080 borders, not gridlines. Row 2 above the location (y 16–29) keeps its gridlines, so the rule covers the location only. The t10 window screenshot of `pivot-layout-outline` and the t10 copy picture of `poi-ExcelPivotTableSample` Tabelle2 (location A3:D7: x 112, 231, 265 are #FFFFFF at y 57–83, the right edge x 343 is #D4D4D4) agree.
- **A styled region cell's fill covers its gridlines.** `table-medium-2` x 53 (A|B): the band fill #C0E6F5 in the striped rows (y 16–29, 46–59), #D4D4D4 in the white rows (y 31–44). The fill also covers the line on its far side: `table-light-1` x 10 is the row-2 fill #D9D9D9 from y 16 through y 30, the row 2|3 line, and row 4's fill starts on its own top line at y 45. `slicer-light1` row 1: the pivot header fill #DCE6F1 runs from x 344 (its left line) through x 450 (the next column's left line).
- **Rule** (`src/xlsx/regionGrid.ts`, both the static painter and the viewport): a pivot location hides the gridlines between its cells; a table or pivot cell whose resolved style has a fill hides the line on each of its edges, and where the fill is on the left or top of an unfilled neighbour the line is painted in that fill (our gridline lies on the first pixel of the cell right of or below it). Own fills outside tables and pivots keep their gridlines, as before; Excel hides those too, which is a cross-format follow-up outside this feature.
- **Goldens.** `poi-ExcelPivotTableSample.xlsx.0–2`, `poi-ExcelTables.xlsx.0` and `poi-tableStyle.xlsx.0` were re-recorded for this rule (Tabelle1 of the POI sample holds a styled table).

### Stage 1 comparison (Task 14)
`scripts/pivot-compare.ts` renders every reference case through the interactive viewport painter at 75 %, which is 1 device px per pt like the captures, and compares the result with Excel's capture. It works as `scripts/sparkline-compare.ts` does, and its header gives the method in full:
- **Mode.** A copy picture shows the screen without the field dropdowns. The ± buttons and the compact gutter stay, and page mode (`print`) would drop both. These cases therefore render in screen mode with `showDataDropDown` turned off. The two window screenshots render in screen mode with the dropdowns.
- **Cell size.** Column and row boundaries are read from the capture's gridlines (#D4D4D4) where they show. Where fills hide them, as inside tables, they are interpolated by natural size. The boundaries are forced on the viewer through session sizes, so anchored drawings move with them.
- **Block.** Each case compares one block: a table's range, a pivot's whole capture, or a slicer's or timeline's frame.
- **Masks.** Every masked rectangle is the same in both images.
  - Cell text is flattened to the rectangle's most frequent colour in each image, so the fill under the text is still compared but the glyphs are not. Excel draws Calibri; the viewer draws whatever its font stack finds.
  - The control masks are listed under "Masked in the comparison".
- **Tolerances.** Each case's tolerance is its measured residual plus 0.002 / 0.010 / 3, rounded up, as for sparklines. The residual is the meanAbs, rowMax and strong columns, measured after the fixes below. It is anti-aliasing: the viewer strokes 1 CSS px borders and gridlines, which cover 0.75 of a device pixel at this scale, while Excel draws whole pixels.
- **New reference.** `refs/pivot-layout-outline-window.png` is cropped from Task 10's window screenshot (t10-excel/shot-pivot-layout-outline.png, x 277–1045, y 318–647). The crop is A1:J22, starting one pixel inside A1's corner as `measure-field-buttons` does. Nothing new was captured.

**Fixed in the owning tasks' files.** Each fix has a test; the evidence is the reference pixels named.
- **Gridlines** (Task 10/13 painters, `src/xlsx/regionGrid.ts`). No gridlines are drawn inside a pivot's location, and a styled cell's fill covers its gridlines. See "Gridlines over pivots and styled tables".
- **firstColumn spans the header rows** (Task 10, `savedLayout.ts`), under headerRow and firstHeaderCell there.
  - `pivot-light-9`: firstColumn's thin right line at x 77 runs from y 29 through header rows 3–5.
  - `pivot-medium-23`: A3:A5 have the firstColumn fill #83CCEB, because headerRow has no fill.
  - `measure-precedence`: headerRow still wins A5.
- **A shared edge takes the line of the higher-precedence element** (Task 7/10, `regionLayer.ts`, `resolve.ts`). Before, the cell painted last won.
  - `pivot-light-3` y 75: the header row's bottom line #E97132 runs under every column, over the next row's #FBE3D6 wholeTable inner line.
  - `pivot-dark-2` y 90: the East subheading's bottom line #44B3E1 wins over the stripe below.
  - Strong pixels went from about 1,100 to 1–43 for Light 1–7, and from 467 to 11 for Dark 2 and 4 (final residuals).
- **One totalRow band for the grand-total lines** (Task 10, `savedLayout.ts`). `pivot-values-on-rows` has one line above its two grand rows (y 780) and none between them. Strong went from 539 to 32.
- **Separators** (Task 12, `slicers/paint.ts`):
  - the slicer separator stops 7 pt inside the frame: x 193–318 in the frame x 186–325, in every slicer style;
  - the timeline separator spans the bar's columns: x 456–642 in the frame x 450–649.
- **Control frame far edges** (Task 12). The right and bottom borders lie on the last device pixel the anchor touches. `slicer-light1`'s anchor ends at y 197.25, and Excel's border is on row 197. The viewer drew it on row 196, which left a whole-row residual of rowMax 0.68.
- **Pivot buttons snap to device pixels** (Task 10, `buttons.ts`).
  - At 75 % the ± box used to be a resampled 9 CSS px square. It is now Excel's 7 × 7 px at x 6–12, and the dropdown's 15 × 15 box is crisp.
  - The ± box sits at rows 5–11 of a 15 px row, one row lower than before (`pivot-light-23`, `pivot-medium-9`: y 80–86 in the row starting at 75).

**Masked in the comparison.** The same rectangles are masked in both images, and each item is unverified for Stage 1. Owners are named.
- **Header icons.** The slicer's multi-select and clear-filter icons (the header's right 50 px) and the timeline's clear-filter icon (the right 30 px) are not drawn by the viewer. They become controls in Stage 5 (Tasks 34–35).
- **Timeline placement.**
  - The generator anchors every timeline at H8 + 62 pt, in a 53 pt column. Excel draws the frame at x 450–649, y 105–225, by a rule not established: neither "clamp the offset" nor "from + to" fits both edges.
  - The comparison places the viewer's timeline in that frame.
  - Excel-written files never carry an offset wider than its cell, so whether the viewer should clamp anchors is a base-drawing follow-up.
- **Timeline frame shade [approx].** Excel's 1 px frame is a partial-coverage accent (Light1: #5887BF left, #7095C2 bottom, #7BA0CD right). The accent itself is #4F81BD, which is what the viewer draws.
- **Timeline bar.** With no selection (every timeline case except `measure-timeline-selection`) the whole bar is masked.
  - The viewer fills the bar flat at the gradient's mid tone, an [approx] under Ruling R.
  - Excel also spaces periods by calendar length: January to April 2024 are 40, 38, 40 and 39 px, with dividers at 495, 533, 573 and 612. The viewer uses a fixed pitch of 39 px, so its dividers land within 1–2 px. The spacing is a Task 12 follow-up.
  - With a selection, only the selected periods are masked; unselected periods, dividers and the scrollbar are compared.
- **Timeline text.** The caption, selection label, level picker, years and periods are in Excel's UI language (zh-TW in the captures).

**Residuals worth naming:**
- **Double borders** (the `pivot-medium-2` to `7` grand total and the `table-totals-row` totals row).
  - Excel draws two crisp rows around the gridline, at y 314 and 316. The viewer strokes 1 CSS px line, gap and line, which blur over three device rows at 75 %.
  - The residual is about 1,470 strong pixels (rowMax 0.41) for the pivots and 523 (rowMax 0.26) for the table, so those seven cases have little `strong` sensitivity.
  - A device-snapped cell border painter in the base renderer would remove it.
- **Outer borders and gridlines** blur to 0.75 coverage. This is most of every case's rowMax; the worst row is a border row.
- **Table resize handle.** Excel's handle, a few pixels at a table's bottom-right corner, is not drawn.

## Cases

Status is `pass` (compared within tolerance) or `*unverified*`, with the reason in Note (spec §9). Tolerance is `meanAbs / rowMax / strong`. The measured columns are the residuals at Task 14's final renderer. "Masked" and "Residuals" in the notes refer to "Stage 1 comparison (Task 14)" above.

| Case id | Status | meanAbs | rowMax | strong | Tolerance | Note |
|---|---|---|---|---|---|---|
| table-light-1 | pass | 0.0064 | 0.2237 | 24 | 0.009 / 0.234 / 27 |  |
| table-light-2 | pass | 0.0082 | 0.1956 | 24 | 0.011 / 0.206 / 27 |  |
| table-light-3 | pass | 0.0073 | 0.1718 | 16 | 0.010 / 0.182 / 19 |  |
| table-light-4 | pass | 0.0082 | 0.1916 | 24 | 0.011 / 0.202 / 27 |  |
| table-light-5 | pass | 0.0082 | 0.1997 | 24 | 0.011 / 0.210 / 27 |  |
| table-light-6 | pass | 0.0075 | 0.1718 | 16 | 0.010 / 0.182 / 19 |  |
| table-light-7 | pass | 0.0074 | 0.1717 | 16 | 0.010 / 0.182 / 19 |  |
| table-light-8 | pass | 0.0074 | 0.2145 | 11 | 0.010 / 0.225 / 14 |  |
| table-light-9 | pass | 0.0068 | 0.1946 | 11 | 0.009 / 0.205 / 14 |  |
| table-light-10 | pass | 0.0060 | 0.1667 | 2 | 0.008 / 0.177 / 5 |  |
| table-light-11 | pass | 0.0067 | 0.1906 | 11 | 0.009 / 0.201 / 14 |  |
| table-light-12 | pass | 0.0069 | 0.1986 | 11 | 0.009 / 0.209 / 14 |  |
| table-light-13 | pass | 0.0059 | 0.1707 | 2 | 0.008 / 0.181 / 5 |  |
| table-light-14 | pass | 0.0061 | 0.1707 | 2 | 0.009 / 0.181 / 5 |  |
| table-light-15 | pass | 0.0095 | 0.2201 | 4 | 0.012 / 0.231 / 7 |  |
| table-light-16 | pass | 0.0100 | 0.1909 | 4 | 0.013 / 0.201 / 7 |  |
| table-light-17 | pass | 0.0092 | 0.1690 | 2 | 0.012 / 0.180 / 5 |  |
| table-light-18 | pass | 0.0099 | 0.1870 | 4 | 0.012 / 0.198 / 7 |  |
| table-light-19 | pass | 0.0102 | 0.1948 | 4 | 0.013 / 0.205 / 7 |  |
| table-light-20 | pass | 0.0091 | 0.1691 | 2 | 0.012 / 0.180 / 5 |  |
| table-light-21 | pass | 0.0092 | 0.1690 | 2 | 0.012 / 0.180 / 5 |  |
| table-medium-1 | pass | 0.0079 | 0.2225 | 11 | 0.010 / 0.233 / 14 |  |
| table-medium-2 | pass | 0.0073 | 0.1430 | 1 | 0.010 / 0.154 / 4 |  |
| table-medium-3 | pass | 0.0049 | 0.0913 | 0 | 0.007 / 0.102 / 3 |  |
| table-medium-4 | pass | 0.0066 | 0.1352 | 1 | 0.009 / 0.146 / 4 |  |
| table-medium-5 | pass | 0.0056 | 0.1153 | 0 | 0.008 / 0.126 / 3 |  |
| table-medium-6 | pass | 0.0058 | 0.1111 | 0 | 0.008 / 0.122 / 3 |  |
| table-medium-7 | pass | 0.0054 | 0.1032 | 0 | 0.008 / 0.114 / 3 |  |
| table-medium-8 | pass | 0.0061 | 0.2146 | 2 | 0.009 / 0.225 / 5 |  |
| table-medium-9 | pass | 0.0100 | 0.1951 | 2 | 0.012 / 0.206 / 5 |  |
| table-medium-10 | pass | 0.0082 | 0.1678 | 2 | 0.011 / 0.178 / 5 |  |
| table-medium-11 | pass | 0.0133 | 0.1912 | 2 | 0.016 / 0.202 / 5 |  |
| table-medium-12 | pass | 0.0112 | 0.1990 | 2 | 0.014 / 0.209 / 5 |  |
| table-medium-13 | pass | 0.0087 | 0.1717 | 2 | 0.011 / 0.182 / 5 |  |
| table-medium-14 | pass | 0.0086 | 0.1717 | 2 | 0.011 / 0.182 / 5 |  |
| table-medium-15 | pass | 0.0102 | 0.2150 | 4 | 0.013 / 0.225 / 7 |  |
| table-medium-16 | pass | 0.0087 | 0.2158 | 4 | 0.011 / 0.226 / 7 |  |
| table-medium-17 | pass | 0.0092 | 0.2158 | 4 | 0.012 / 0.226 / 7 |  |
| table-medium-18 | pass | 0.0085 | 0.2158 | 4 | 0.011 / 0.226 / 7 |  |
| table-medium-19 | pass | 0.0091 | 0.2158 | 4 | 0.012 / 0.226 / 7 |  |
| table-medium-20 | pass | 0.0089 | 0.2158 | 4 | 0.011 / 0.226 / 7 |  |
| table-medium-21 | pass | 0.0089 | 0.2158 | 4 | 0.011 / 0.226 / 7 |  |
| table-medium-22 | pass | 0.0078 | 0.2027 | 4 | 0.010 / 0.213 / 7 |  |
| table-medium-23 | pass | 0.0086 | 0.1404 | 0 | 0.011 / 0.151 / 3 |  |
| table-medium-24 | pass | 0.0065 | 0.0780 | 0 | 0.009 / 0.089 / 3 |  |
| table-medium-25 | pass | 0.0116 | 0.1286 | 0 | 0.014 / 0.139 / 3 |  |
| table-medium-26 | pass | 0.0088 | 0.1054 | 0 | 0.011 / 0.116 / 3 |  |
| table-medium-27 | pass | 0.0071 | 0.0976 | 0 | 0.010 / 0.108 / 3 |  |
| table-medium-28 | pass | 0.0067 | 0.0896 | 0 | 0.009 / 0.100 / 3 |  |
| table-dark-1 | pass | 0.0063 | 0.2162 | 4 | 0.009 / 0.227 / 7 |  |
| table-dark-2 | pass | 0.0091 | 0.2339 | 13 | 0.012 / 0.244 / 16 |  |
| table-dark-3 | pass | 0.0109 | 0.2338 | 10 | 0.013 / 0.244 / 13 |  |
| table-dark-4 | pass | 0.0090 | 0.2338 | 13 | 0.012 / 0.244 / 16 |  |
| table-dark-5 | pass | 0.0110 | 0.2339 | 13 | 0.013 / 0.244 / 16 |  |
| table-dark-6 | pass | 0.0107 | 0.2162 | 12 | 0.013 / 0.227 / 15 |  |
| table-dark-7 | pass | 0.0088 | 0.2179 | 12 | 0.011 / 0.228 / 15 |  |
| table-dark-8 | pass | 0.0038 | 0.2162 | 2 | 0.006 / 0.227 / 5 |  |
| table-dark-9 | pass | 0.0069 | 0.1690 | 2 | 0.009 / 0.179 / 5 |  |
| table-dark-10 | pass | 0.0112 | 0.2004 | 2 | 0.014 / 0.211 / 5 |  |
| table-dark-11 | pass | 0.0066 | 0.1729 | 2 | 0.009 / 0.183 / 5 |  |
| table-flag-first-column-on | pass | 0.0073 | 0.1430 | 1 | 0.010 / 0.154 / 4 |  |
| table-flag-first-column-off | pass | 0.0073 | 0.1430 | 1 | 0.010 / 0.154 / 4 |  |
| table-flag-last-column-on | pass | 0.0073 | 0.1430 | 1 | 0.010 / 0.154 / 4 |  |
| table-flag-last-column-off | pass | 0.0073 | 0.1430 | 1 | 0.010 / 0.154 / 4 |  |
| table-flag-row-stripes-on | pass | 0.0073 | 0.1430 | 1 | 0.010 / 0.154 / 4 |  |
| table-flag-row-stripes-off | pass | 0.0056 | 0.1547 | 2 | 0.008 / 0.165 / 5 |  |
| table-flag-column-stripes-on | pass | 0.0080 | 0.1430 | 1 | 0.010 / 0.154 / 4 |  |
| table-flag-column-stripes-off | pass | 0.0073 | 0.1430 | 1 | 0.010 / 0.154 / 4 |  |
| table-totals-row | pass | 0.0115 | 0.2577 | 523 | 0.014 / 0.268 / 526 | double top border blurs at 75 % (Residuals) |
| table-custom-dxf | pass | 0.0062 | 0.2497 | 16 | 0.009 / 0.260 / 19 |  |
| pivot-light-1 | pass | 0.0015 | 0.1226 | 1 | 0.004 / 0.133 / 4 |  |
| pivot-light-2 | pass | 0.0027 | 0.2295 | 43 | 0.005 / 0.240 / 46 |  |
| pivot-light-3 | pass | 0.0023 | 0.2017 | 40 | 0.005 / 0.212 / 43 |  |
| pivot-light-4 | pass | 0.0029 | 0.2255 | 43 | 0.005 / 0.236 / 46 |  |
| pivot-light-5 | pass | 0.0025 | 0.2335 | 43 | 0.005 / 0.244 / 46 |  |
| pivot-light-6 | pass | 0.0025 | 0.2057 | 43 | 0.005 / 0.216 / 46 |  |
| pivot-light-7 | pass | 0.0023 | 0.2057 | 43 | 0.005 / 0.216 / 46 |  |
| pivot-light-8 | pass | 0.0022 | 0.1197 | 8 | 0.005 / 0.130 / 11 |  |
| pivot-light-9 | pass | 0.0049 | 0.2242 | 21 | 0.007 / 0.235 / 24 |  |
| pivot-light-10 | pass | 0.0043 | 0.1971 | 21 | 0.007 / 0.208 / 24 |  |
| pivot-light-11 | pass | 0.0049 | 0.2204 | 21 | 0.007 / 0.231 / 24 |  |
| pivot-light-12 | pass | 0.0047 | 0.2280 | 21 | 0.007 / 0.239 / 24 |  |
| pivot-light-13 | pass | 0.0044 | 0.2009 | 21 | 0.007 / 0.211 / 24 |  |
| pivot-light-14 | pass | 0.0043 | 0.2010 | 21 | 0.007 / 0.211 / 24 |  |
| pivot-light-15 | pass | 0.0007 | 0.0351 | 0 | 0.003 / 0.046 / 3 |  |
| pivot-light-16 | pass | 0.0021 | 0.0628 | 10 | 0.005 / 0.073 / 13 |  |
| pivot-light-17 | pass | 0.0016 | 0.0431 | 0 | 0.004 / 0.054 / 3 |  |
| pivot-light-18 | pass | 0.0019 | 0.0627 | 10 | 0.004 / 0.073 / 13 |  |
| pivot-light-19 | pass | 0.0018 | 0.0510 | 0 | 0.004 / 0.062 / 3 |  |
| pivot-light-20 | pass | 0.0018 | 0.0509 | 0 | 0.004 / 0.061 / 3 |  |
| pivot-light-21 | pass | 0.0017 | 0.0470 | 0 | 0.004 / 0.058 / 3 |  |
| pivot-light-22 | pass | 0.0012 | 0.0862 | 0 | 0.004 / 0.097 / 3 |  |
| pivot-light-23 | pass | 0.0023 | 0.1920 | 53 | 0.005 / 0.203 / 56 |  |
| pivot-light-24 | pass | 0.0020 | 0.1646 | 18 | 0.005 / 0.175 / 21 |  |
| pivot-light-25 | pass | 0.0022 | 0.1881 | 53 | 0.005 / 0.199 / 56 |  |
| pivot-light-26 | pass | 0.0023 | 0.1960 | 53 | 0.005 / 0.206 / 56 |  |
| pivot-light-27 | pass | 0.0020 | 0.1685 | 19 | 0.005 / 0.179 / 22 |  |
| pivot-light-28 | pass | 0.0021 | 0.1685 | 19 | 0.005 / 0.179 / 22 |  |
| pivot-medium-1 | pass | 0.0026 | 0.2193 | 0 | 0.005 / 0.230 / 3 |  |
| pivot-medium-2 | pass | 0.0054 | 0.4111 | 1472 | 0.008 / 0.422 / 1475 | double grand-total border blurs at 75 % (Residuals) |
| pivot-medium-3 | pass | 0.0049 | 0.4033 | 1472 | 0.007 / 0.414 / 1475 | double grand-total border blurs at 75 % (Residuals) |
| pivot-medium-4 | pass | 0.0048 | 0.4033 | 1472 | 0.007 / 0.414 / 1475 | double grand-total border blurs at 75 % (Residuals) |
| pivot-medium-5 | pass | 0.0052 | 0.4189 | 1471 | 0.008 / 0.429 / 1474 | double grand-total border blurs at 75 % (Residuals) |
| pivot-medium-6 | pass | 0.0049 | 0.3837 | 1471 | 0.007 / 0.394 / 1474 | double grand-total border blurs at 75 % (Residuals) |
| pivot-medium-7 | pass | 0.0038 | 0.3760 | 1471 | 0.006 / 0.386 / 1474 | double grand-total border blurs at 75 % (Residuals) |
| pivot-medium-8 | pass | 0.0024 | 0.1199 | 2 | 0.005 / 0.130 / 5 |  |
| pivot-medium-9 | pass | 0.0041 | 0.2285 | 4 | 0.007 / 0.239 / 7 |  |
| pivot-medium-10 | pass | 0.0039 | 0.2285 | 4 | 0.006 / 0.239 / 7 |  |
| pivot-medium-11 | pass | 0.0040 | 0.2284 | 4 | 0.007 / 0.239 / 7 |  |
| pivot-medium-12 | pass | 0.0040 | 0.2283 | 4 | 0.006 / 0.239 / 7 |  |
| pivot-medium-13 | pass | 0.0037 | 0.2090 | 4 | 0.006 / 0.219 / 7 |  |
| pivot-medium-14 | pass | 0.0037 | 0.2130 | 4 | 0.006 / 0.223 / 7 |  |
| pivot-medium-15 | pass | 0.0011 | 0.2486 | 22 | 0.004 / 0.259 / 25 |  |
| pivot-medium-16 | pass | 0.0041 | 0.2486 | 22 | 0.007 / 0.259 / 25 |  |
| pivot-medium-17 | pass | 0.0040 | 0.2486 | 22 | 0.007 / 0.259 / 25 |  |
| pivot-medium-18 | pass | 0.0041 | 0.2486 | 22 | 0.007 / 0.259 / 25 |  |
| pivot-medium-19 | pass | 0.0041 | 0.2486 | 22 | 0.007 / 0.259 / 25 |  |
| pivot-medium-20 | pass | 0.0040 | 0.2486 | 22 | 0.007 / 0.259 / 25 |  |
| pivot-medium-21 | pass | 0.0040 | 0.2486 | 22 | 0.007 / 0.259 / 25 |  |
| pivot-medium-22 | pass | 0.0005 | 0.0384 | 0 | 0.003 / 0.049 / 3 |  |
| pivot-medium-23 | pass | 0.0042 | 0.0693 | 0 | 0.007 / 0.080 / 3 |  |
| pivot-medium-24 | pass | 0.0040 | 0.0476 | 0 | 0.007 / 0.058 / 3 |  |
| pivot-medium-25 | pass | 0.0049 | 0.0693 | 0 | 0.007 / 0.080 / 3 |  |
| pivot-medium-26 | pass | 0.0045 | 0.0576 | 0 | 0.007 / 0.068 / 3 |  |
| pivot-medium-27 | pass | 0.0041 | 0.0558 | 0 | 0.007 / 0.066 / 3 |  |
| pivot-medium-28 | pass | 0.0041 | 0.0519 | 0 | 0.007 / 0.062 / 3 |  |
| pivot-dark-1 | pass | 0.0018 | 0.1220 | 0 | 0.004 / 0.133 / 3 |  |
| pivot-dark-2 | pass | 0.0057 | 0.2367 | 11 | 0.008 / 0.247 / 14 |  |
| pivot-dark-3 | pass | 0.0059 | 0.2367 | 11 | 0.008 / 0.247 / 14 |  |
| pivot-dark-4 | pass | 0.0095 | 0.2367 | 11 | 0.012 / 0.247 / 14 |  |
| pivot-dark-5 | pass | 0.0080 | 0.2446 | 11 | 0.011 / 0.255 / 14 |  |
| pivot-dark-6 | pass | 0.0053 | 0.2327 | 11 | 0.008 / 0.243 / 14 |  |
| pivot-dark-7 | pass | 0.0053 | 0.2249 | 11 | 0.008 / 0.235 / 14 |  |
| pivot-dark-8 | pass | 0.0013 | 0.1197 | 16 | 0.004 / 0.130 / 19 |  |
| pivot-dark-9 | pass | 0.0043 | 0.1197 | 26 | 0.007 / 0.130 / 29 |  |
| pivot-dark-10 | pass | 0.0042 | 0.1197 | 26 | 0.007 / 0.130 / 29 |  |
| pivot-dark-11 | pass | 0.0056 | 0.1197 | 26 | 0.008 / 0.130 / 29 |  |
| pivot-dark-12 | pass | 0.0049 | 0.1197 | 26 | 0.007 / 0.130 / 29 |  |
| pivot-dark-13 | pass | 0.0042 | 0.1197 | 26 | 0.007 / 0.130 / 29 |  |
| pivot-dark-14 | pass | 0.0043 | 0.1197 | 26 | 0.007 / 0.130 / 29 |  |
| pivot-dark-15 | pass | 0.0023 | 0.2486 | 15 | 0.005 / 0.259 / 18 |  |
| pivot-dark-16 | pass | 0.0021 | 0.2486 | 32 | 0.005 / 0.259 / 35 |  |
| pivot-dark-17 | pass | 0.0028 | 0.2486 | 31 | 0.005 / 0.259 / 34 |  |
| pivot-dark-18 | pass | 0.0021 | 0.2486 | 32 | 0.005 / 0.259 / 35 |  |
| pivot-dark-19 | pass | 0.0028 | 0.2486 | 34 | 0.005 / 0.259 / 37 |  |
| pivot-dark-20 | pass | 0.0028 | 0.2486 | 31 | 0.005 / 0.259 / 34 |  |
| pivot-dark-21 | pass | 0.0021 | 0.2486 | 31 | 0.005 / 0.259 / 34 |  |
| pivot-dark-22 | pass | 0.0021 | 0.1223 | 0 | 0.005 / 0.133 / 3 |  |
| pivot-dark-23 | pass | 0.0030 | 0.2292 | 35 | 0.005 / 0.240 / 38 |  |
| pivot-dark-24 | pass | 0.0037 | 0.2045 | 35 | 0.006 / 0.215 / 38 |  |
| pivot-dark-25 | pass | 0.0030 | 0.2256 | 35 | 0.005 / 0.236 / 38 |  |
| pivot-dark-26 | pass | 0.0039 | 0.2327 | 35 | 0.006 / 0.243 / 38 |  |
| pivot-dark-27 | pass | 0.0037 | 0.2060 | 35 | 0.006 / 0.216 / 38 |  |
| pivot-dark-28 | pass | 0.0028 | 0.2063 | 35 | 0.005 / 0.217 / 38 |  |
| pivot-layout-compact | pass | 0.0021 | 0.0628 | 10 | 0.005 / 0.073 / 13 |  |
| pivot-layout-outline | pass | 0.0021 | 0.0628 | 11 | 0.005 / 0.073 / 14 |  |
| pivot-layout-tabular | pass | 0.0073 | 0.1682 | 17 | 0.010 / 0.179 / 20 |  |
| pivot-subtotals-top | pass | 0.0021 | 0.0628 | 10 | 0.005 / 0.073 / 13 |  |
| pivot-subtotals-bottom | pass | 0.0029 | 0.0628 | 24 | 0.005 / 0.073 / 27 |  |
| pivot-grand-totals-off | pass | 0.0018 | 0.0563 | 9 | 0.004 / 0.067 / 12 |  |
| pivot-values-on-rows | pass | 0.0018 | 0.0628 | 32 | 0.004 / 0.073 / 35 |  |
| pivot-collapsed-item | pass | 0.0024 | 0.0628 | 10 | 0.005 / 0.073 / 13 |  |
| pivot-page-field | pass | 0.0033 | 0.0628 | 5 | 0.006 / 0.073 / 8 |  |
| pivot-repeat-labels | pass | 0.0073 | 0.1682 | 17 | 0.010 / 0.179 / 20 |  |
| pivot-custom-style | pass | 0.0014 | 0.0987 | 0 | 0.004 / 0.109 / 3 |  |
| pivot-formats-override | pass | 0.0021 | 0.0628 | 10 | 0.005 / 0.073 / 13 |  |
| slicer-light1 | pass | 0.0009 | 0.0065 | 0 | 0.003 / 0.017 / 3 | header icons masked (Masked) |
| slicer-light2 | pass | 0.0009 | 0.0065 | 0 | 0.003 / 0.017 / 3 | header icons masked (Masked) |
| slicer-light3 | pass | 0.0009 | 0.0065 | 0 | 0.003 / 0.017 / 3 | header icons masked (Masked) |
| slicer-light4 | pass | 0.0009 | 0.0062 | 0 | 0.003 / 0.017 / 3 | header icons masked (Masked) |
| slicer-light5 | pass | 0.0004 | 0.0032 | 0 | 0.003 / 0.014 / 3 | header icons masked (Masked) |
| slicer-light6 | pass | 0.0010 | 0.0070 | 0 | 0.003 / 0.017 / 3 | header icons masked (Masked) |
| slicer-dark1 | pass | 0.0004 | 0.0028 | 0 | 0.003 / 0.013 / 3 | header icons masked (Masked) |
| slicer-dark2 | pass | 0.0004 | 0.0028 | 0 | 0.003 / 0.013 / 3 | header icons masked (Masked) |
| slicer-dark3 | pass | 0.0004 | 0.0026 | 0 | 0.003 / 0.013 / 3 | header icons masked (Masked) |
| slicer-dark4 | pass | 0.0003 | 0.0025 | 0 | 0.003 / 0.013 / 3 | header icons masked (Masked) |
| slicer-dark5 | pass | 0.0004 | 0.0028 | 0 | 0.003 / 0.013 / 3 | header icons masked (Masked) |
| slicer-dark6 | pass | 0.0004 | 0.0029 | 0 | 0.003 / 0.013 / 3 | header icons masked (Masked) |
| slicer-other1 | pass | 0.0004 | 0.0022 | 0 | 0.003 / 0.013 / 3 | header icons masked (Masked) |
| slicer-other2 | pass | 0.0056 | 0.0854 | 0 | 0.008 / 0.096 / 3 | header icons masked (Masked) |
| slicer-two-column | pass | 0.0007 | 0.0065 | 0 | 0.003 / 0.017 / 3 | header icons masked (Masked) |
| slicer-no-data | pass | 0.0010 | 0.0065 | 0 | 0.003 / 0.017 / 3 | header icons masked (Masked) |
| slicer-scrolled | pass | 0.0015 | 0.0308 | 0 | 0.004 / 0.041 / 3 | header icons masked (Masked) |
| timeline-light1-years | pass | 0.0007 | 0.0401 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light1-quarters | pass | 0.0007 | 0.0401 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light1-months | pass | 0.0008 | 0.0401 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light1-days | pass | 0.0011 | 0.0401 | 0 | 0.004 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light2-years | pass | 0.0005 | 0.0402 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light2-quarters | pass | 0.0015 | 0.0402 | 16 | 0.004 / 0.051 / 19 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light2-months | pass | 0.0011 | 0.0402 | 16 | 0.004 / 0.051 / 19 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light2-days | pass | 0.0010 | 0.0402 | 0 | 0.004 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light3-years | pass | 0.0008 | 0.0365 | 0 | 0.003 / 0.047 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light3-quarters | pass | 0.0015 | 0.0365 | 0 | 0.004 / 0.047 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light3-months | pass | 0.0013 | 0.0365 | 0 | 0.004 / 0.047 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light3-days | pass | 0.0013 | 0.0365 | 0 | 0.004 / 0.047 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light4-years | pass | 0.0006 | 0.0329 | 0 | 0.003 / 0.043 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light4-quarters | pass | 0.0010 | 0.0329 | 0 | 0.003 / 0.043 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light4-months | pass | 0.0009 | 0.0329 | 0 | 0.003 / 0.043 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light4-days | pass | 0.0011 | 0.0329 | 0 | 0.004 / 0.043 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light5-years | pass | 0.0007 | 0.0402 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light5-quarters | pass | 0.0010 | 0.0402 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light5-months | pass | 0.0009 | 0.0402 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light5-days | pass | 0.0011 | 0.0402 | 0 | 0.004 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light6-years | pass | 0.0021 | 0.0436 | 0 | 0.005 / 0.054 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light6-quarters | pass | 0.0024 | 0.0473 | 24 | 0.005 / 0.058 / 27 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light6-months | pass | 0.0024 | 0.0473 | 24 | 0.005 / 0.058 / 27 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-light6-days | pass | 0.0026 | 0.0509 | 65 | 0.005 / 0.061 / 68 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark1-years | pass | 0.0007 | 0.0401 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark1-quarters | pass | 0.0007 | 0.0401 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark1-months | pass | 0.0008 | 0.0401 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark1-days | pass | 0.0011 | 0.0401 | 0 | 0.004 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark2-years | pass | 0.0005 | 0.0402 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark2-quarters | pass | 0.0015 | 0.0402 | 16 | 0.004 / 0.051 / 19 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark2-months | pass | 0.0011 | 0.0402 | 16 | 0.004 / 0.051 / 19 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark2-days | pass | 0.0010 | 0.0402 | 0 | 0.004 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark3-years | pass | 0.0008 | 0.0365 | 0 | 0.003 / 0.047 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark3-quarters | pass | 0.0015 | 0.0365 | 0 | 0.004 / 0.047 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark3-months | pass | 0.0013 | 0.0365 | 0 | 0.004 / 0.047 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark3-days | pass | 0.0013 | 0.0365 | 0 | 0.004 / 0.047 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark4-years | pass | 0.0004 | 0.0329 | 0 | 0.003 / 0.043 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark4-quarters | pass | 0.0010 | 0.0329 | 0 | 0.003 / 0.043 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark4-months | pass | 0.0008 | 0.0329 | 0 | 0.003 / 0.043 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark4-days | pass | 0.0009 | 0.0329 | 0 | 0.003 / 0.043 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark5-years | pass | 0.0007 | 0.0402 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark5-quarters | pass | 0.0010 | 0.0402 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark5-months | pass | 0.0009 | 0.0402 | 0 | 0.003 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark5-days | pass | 0.0011 | 0.0402 | 0 | 0.004 / 0.051 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark6-years | pass | 0.0021 | 0.0436 | 0 | 0.005 / 0.054 / 3 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark6-quarters | pass | 0.0024 | 0.0473 | 24 | 0.005 / 0.058 / 27 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark6-months | pass | 0.0024 | 0.0473 | 24 | 0.005 / 0.058 / 27 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| timeline-dark6-days | pass | 0.0026 | 0.0509 | 65 | 0.005 / 0.061 / 68 | placed in Excel's frame; bar, frame shade, icon and labels masked (Masked) |
| measure-precedence | pass | 0.0008 | 0.1418 | 0 | 0.003 / 0.152 / 3 |  |
| measure-cell-override | pass | 0.0022 | 0.0765 | 14 | 0.005 / 0.087 / 17 |  |
| measure-apply-flags | *unverified* | 0.0972 | 0.3774 | 6782 | — | The workbook has no theme part. Excel draws its current default theme (Office 2023: header #156082, band #C0E6F5) and the viewer its no-theme fallback (Office 2007–2010 colours), a cross-format follow-up under Ruling AJ (refined). The case's subject, own fills and fonts over a table style (Ruling P), is pinned by the Task 5 tests. |
| measure-indent | pass | 0.0030 | 0.0309 | 0 | 0.005 / 0.041 / 3 | text masked: fills and gridlines only; indent pinned by Task 5 |
| measure-timeline-selection | pass | 0.0041 | 0.0853 | 0 | 0.007 / 0.096 / 3 | placed in Excel's frame; selected periods masked, unselected compared (Masked) |
| measure-field-buttons | pass | 0.0033 | 0.0627 | 20 | 0.006 / 0.073 / 23 | window screenshot: screen mode with dropdowns |
| pivot-layout-outline-window | pass | 0.0020 | 0.0627 | 11 | 0.005 / 0.073 / 14 | window screenshot: screen mode with dropdowns |
| pivot-chart-buttons-all-window | *unverified* | — | — | — | — | A window screenshot of the chart only, with no cell range. Its pixels belong to the chart renderer. What Stage 1 takes from it, that Excel for Mac draws no PivotChart field buttons (Ruling AI), is asserted by tests/charts-pivot-buttons.test.ts. |

## Stage 2 engine oracle and measurements (Task 15, 2026-10-31, Excel 16.114, zh-TW UI)

### How the oracle is made
- **Seed.** `recipes/engine-seed.xlsx` is built from `recipes/pivot-seed.xlsx` by `bun scripts/xlsx-refs/build-engine.ts --seed`.
  - AppleScript writes a 56-record, 12-column source (A1:L57) with blanks, ties, negatives, mixed-type Channel values, punctuation codes, month and weekday names, and a Stamp column of serials 0.5, 1, 59, 60 and 61. Row 58 holds a `#N/A` record outside the source; only the error cases widen the source to it.
  - The pivot is then pointed at the new source. `set source data` is a syntax error in Excel's dictionary (-2741), so the seed takes the XML route: the `worksheetSource` ref is rewritten, and Excel opens, refreshes and saves the file.
- **Cases.** `scripts/xlsx-refs/engine-cases.ts` defines 125 cases. `build-engine.ts` copies the seed for each one, applies the case, runs `refresh table`, saves, and checks markers in the saved parts, because Excel can ignore a setting without an error.
  - 96 cases go through AppleScript. 29 go through the XML route: an optional AppleScript prep, then a part edit, then Excel reopens, refreshes and saves.
  - Excel opened every XML-route file without a repair prompt, and re-applied filters on refresh. For example, `filter-label-begins-with`'s `rowItems` went from 19 to 9.
  - No case needed the owner, so there is no Stage 2 recipe batch.
- **Manifest.** `engine/manifest.json` has one entry per case: `{ case, file, features, task, pivot, route, status }`. `pivot` is the pivot table's `name` attribute, which Excel writes in the UI locale (`樞紐分析表2`).

### Dictionary probes (`--probes`)
| Reshape | Result |
|---|---|
| `set source data of <pivot> to "…"` | refused (syntax error -2741): XML route |
| `set value of range … to {{…}}`, `set formula`, `set number format` | reached |
| `set subtotals <field> subtotal index sum subtotal / count subtotal / average subtotal value true` | reached. `count subtotal` is Excel's Count, which writes `countASubtotal="1"` and `<item t="countA"/>` |
| `set pivot field orientation of data field … to orient as hidden` | reached |
| `data field N of <pivot>` for an added data field, then `set name` | reached |
| `set base item … to "(previous)"` / `"(next)"` | reached |
| `make new pivot filter at <field> with properties {…}` | refused (parameter error -50) for every filter type: XML route |
| `auto show <field> …` | ignored (no `autoShow` or top-10 filter saved): XML route |
| `group (label range …) start/end/by/periods` | reached |
| `make new calculated field` / `make new calculated item` | reached |
| `set enable multiple page items of <field> to true` | refused (-10006): XML route |
| `set column grand` / `set row grand` | reached, named the other way round from ECMA: `column grand` is `rowGrandTotals`, `row grand` is `colGrandTotals` |

### What Excel writes (measured from the saved parts)
- **Dates in the cache (ruling S2-H).** The ISO value is 1899-12-30 + serial, applied uniformly: serials 0.5, 1, 59, 60 and 61 are saved as `1899-12-30T12:00:00`, `1899-12-31T00:00:00`, `1900-02-27`, `1900-02-28` and `1900-03-01`. There is no 1900 leap-year correction in the cache. The fictitious 1900-02-29 is serial 60, written as `1900-02-28`.
- **Fields on no axis.** Their text stays inline in `pivotCacheRecords` (`<s v="…"/>`), and the definition keeps an empty `<sharedItems/>`.
  - When a field leaves every axis, Excel drops its hidden items and its shared items. An off-axis hidden item is therefore never written by Excel (`quirk-off-axis-hidden-item`). This was checked in two saves:
    - with Channel on rows and Retail hidden, Excel saved `<item h="1" x="1"/>` and four Channel shared items;
    - after Channel was taken off the axis, the same file saved as `<pivotField showAll="0"/>` with an empty `sharedItems`.
  - The engine still refuses that state in files from other writers.
- **`showAll`.** True is written by omitting the attribute (the schema default); false is `showAll="0"`.
- **Data field captions.** Setting a data field's function, even to Sum, renames it with Excel's localised default (`加總 - Sales`). Captions Excel wrote earlier stay as they are. The engine takes captions from the saved `name`.
- **Named subtotals.** Excel writes the `*Subtotal` flags and one `<item t=…>` per function after the item list. It writes no `defaultSubtotal` attribute and no `default` item.

### Rules read off the oracle (implemented and cited in the engine)
- **Totals arithmetic (ruling S2-AB, as refined).** Each leaf cell (leaf row × leaf column) is the naive Float64 sum of its records in record order. Totals are built from leaf cells, never from intermediate subtotals and never from the records directly:
  - each leaf row sums its leaf cells across the columns;
  - each total then sums those leaf-row cells down the rows, in display order.
  - This is bit-exact on every Sum cell of `baseline`, `values-fn-sum` and `values-addition-order`. Percent of total and index use the same totals; index is `((v × G) / R) / C`.
- **Average totals** sum the records leaf row by leaf row; column totals then sum those cells. This is fitted on `values-fn-average` and `values-average-hidden-item`.
- **Visible totals.** Records of hidden items, page selections and filtered-out items are left out of every total.
- **Blank and error cells.**
  - Sum and Count over no values are blank.
  - Count Numbers is 0 over text-only groups and blank with no values.
  - A sample statistic over one number is `#DIV/0!`.
  - For difference, % of and % difference, the blank cells, the base item's own cells, `#N/A`, `#NULL!` and `#DIV/0!` all follow the saved cells.
  - % of parent with a base field is blank on the grand line.
- **Rank** ties are dense.
- **Calculated fields** are evaluated on each cell's sums, unrounded, because Excel keeps full precision in pivot cells. In detail:
  - a cell without records uses zero sums;
  - the data field's function is ignored;
  - text values are left out of the sums;
  - a formula naming no field fills every cell;
  - a reference to another calculated field applies that formula to the same sums;
  - percent of total divides each cell by the field's grand-total cell.
- **Layout.**
  - The Values row is hidden in every form but classic.
  - Classic layout with one data field and no column field has two header rows.
  - A collapsed item above Values gets its own line, then one line per data field.
  - `mergeItem` merges the outer row labels, the total captions and the column labels.
  - A showAll item with no records lists the next field's visible items.
- **Sorting.**
  - Mixed types sort as number < text < blank, and descending reverses all of it.
  - Months and weekdays follow their custom lists unless `customListSort` is off.
  - Text uses Excel's word sort: apostrophes and hyphens are ignored, a difference only in hyphens sorts last, symbol weights are fixed, and only characters the oracle measured are accepted.
  - Value-sort ties keep item order.
- **Filters** apply in `evalOrder`.
  - Label filters match captions, including renamed items and wildcards.
  - Value and top-10 filters rank items by their aggregate over the records left by earlier filters, per parent when nested. Ties at a Count cut are kept.
  - Top **Percent** keeps the top items whose values reach N % of the total (`filter-top-percent-30`: West + South, 120099.08).
- **Groups.**
  - Number ranges are `[a, a + interval)`. Excel relabels the last bucket's upper bound as start + count × interval (`group-number-fractional`: `98-100.5`, `>100.5`).
  - Dates group by years, quarters and months. Blanks fall in the `<start` bucket.
  - A base field's `fieldGroup par` link is not a group definition (`group-manual`).

### Engine results per case
All 126 manifest cases go through `checkSavedLayout` (`tests/xlsx-pivots-*.test.ts`; the triage dump is `.superpowers/sdd/…/dump-cases.ts` during development).

| Task | Cases | ok | Refused (unverified) | Mismatch |
|---|---|---|---|---|
| 17 layout and data quirks | 33 | 33 | 0 | 0 |
| 18 values and show-as | 38 | 34 | 4 | 0 |
| 21 sort, filter, group | 44 | 43 | 1 | 0 |
| 22 calculated | 11 | 10 | 1 | 0 |
| **All** | **126** | **120** | **6** | **0** |

The refused cases, each asserted with its exact message in its test file:

| Case | Refusal | Why Excel's file does not settle it |
|---|---|---|
| `values-fn-stdev`, `values-fn-stdevp`, `values-fn-var`, `values-fn-varp` | `pivot-engine-unmeasured` | Every leaf cell matches. The total cells differ in the 15th digit under every accumulation tried: leaf-moment rollup, records in display order, row or column partials, Welford, two-pass, fused multiply-add. |
| `sort-value-showall-no-data` | `pivot-engine-unmeasured` | Several no-data items under one parent; their tie order is not measured. |
| `calc-item` | `pivot-calculated-item` | Excel lists a calculated item under every parent and counts it in totals. Spec §5.6 recomputes calculated items only once the oracle passes (ruling S2-AD). |

Beyond the manifest, the engine refuses with a named reason, rather than guessing, every construct no case measures, among them:
- Bottom-N filters;
- nested Top Percent;
- sorting a group field;
- custom lists mixed with other labels;
- labels with characters outside the measured collation set (other punctuation, accented letters, CJK);
- booleans and dates in mixed sorts;
- date filters in a 1904 workbook;
- date groups below serial 61;
- variance totals;
- Average, Product and variance on a collapsed item;
- error-valued records under any function but Sum and Count, under a named subtotal of another function, or under show-as;
- % of parent total with its base field on the columns, or on a cell without a parent line (grand totals off);
- error captions (`showError`) and missing-value captions (`showMissing` off or a `missingCaption`); the oracle files keep Excel's defaults, which compute;
- value filters, top-10 filters and value sorts on a data field with show-as;
- date grouping by days, hours, minutes or seconds;
- calculated items on any field, used or not;
- a collapsed item above Values with more than one unfilled level, and `mergeItem` beyond two tabular row fields and one column field;
- the record-less showAll listing outside its measured shape (rows, second level, last row field, compact parent), or when a listed item has no record past the hidden items and page selections.

Show-as modes that scan sibling or parent lines (previous/next and explicit base items, running totals, rank, % of parent) refuse as `pivot-values-limit` when cells × axis lines would exceed 5×10⁷. An indexed sibling lookup is a Stage 3 entry item.

Error-valued item labels are placed as text, as Excel saves them (`sort-error-item`, now ok).

### Speed and size (Task 23, 2026-10-31)
- **§6.3 gate.** `bun scripts/pivot-bench.ts --runs 3` runs 1,000,000 spec-shaped records (3 row fields of 10/50/200 items, 1 column field of 12, 2 sum data fields). `computeAxes + computeValues` gives a median-of-medians of **135.89 ms** (final-review fix round; 140.93 ms at Task 23), against a limit of 200 and an internal bar of 180. Decoding the records is reported separately.
- **Bundles.** The eager bundle is 35,968 B gzip, unchanged from `main` (budget 38,720). The engine lives in the lazy pivots chunk, which grows from 40,823 B raw / 12,702 B gzip on `main` to 161,560 / 46,267 at Task 23 (+33,565 B gzip), and to 164,065 B raw / 46.98 kB gzip (Vite's report) after the final-review fix round's refusals.

**LibreOffice and POI files** (`corpus/`, recorded, never tuned toward): see `CORPUS_STATUS` in `tests/xlsx-pivots-faithful.test.ts`. Collabora's `calcfields` reports a `colItems` mismatch, so its saved cells stay shown.
