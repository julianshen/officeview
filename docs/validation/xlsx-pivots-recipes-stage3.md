# XLSX pivots: Stage 3 owner recipes (draft)

One batch for the owner (pivots plan Task 24, step 4). Everything else in the Stage 3 interaction oracle is built by
`bun scripts/xlsx-refs/build-engine.ts --interaction`. These recipes cover only what an Excel alert blocks: an alert
stops an AppleScript-driven change, so automation cannot record what Excel says or does. They measure spec §7.1
(overlap) and §7.9 (sheet sort and filter across a pivot).

## Before you start

- Excel for Mac 16.114, zh-TW UI (the same install the oracle used).
- Work on copies. Never save over a file under `tests/fixtures/xlsx-pivots/engine/` or `recipes/`.
- Start files:
  - **C** = `tests/fixtures/xlsx-pivots/engine/layout-collapsed-item.xlsx`. This is the engine seed with East collapsed.
    - Sheet Pivot holds the pivot at `A3:I21`, compact form.
    - Expanding East (Region) adds 3 rows (East's Apples, Bananas and Cherries), so the pivot grows to `A3:I24`.
    - After the expand, the rows below East's line (A9 onward) move down 3.
  - **S** = `tests/fixtures/xlsx-pivots/recipes/engine-seed.xlsx`. This is the seed itself, pivot at `A3:I24`.
- To expand an item: click the **+** button left of "East" in A6, or right-click A6 → 展開/摺疊 (Expand/Collapse) → 展開 (Expand).
- **For every recipe, record:**
  1. the alert text, verbatim (copy it, or screenshot it into `docs/validation/assets/xlsx-pivots-stage3/<recipe id>.png`);
  2. the alert's buttons;
  3. what happened after you dismissed it: did the pivot change, and did the blocking cell keep its content and format?
- Then save the workbook (File → Save As, Excel Workbook .xlsx) under the name given, even when nothing changed. The
  saved file shows Excel's resulting state.

## O1 — expand into a non-empty cell (§7.1)

1. Open a copy of **C**.
2. On sheet Pivot, type `blocker` into **A23**. A23 is outside the pivot now (it ends at row 21), and inside it after the expand.
3. Expand East.
4. Record the alert. Answer it with the default button. If it offers a choice, take the one that keeps `blocker`, and note what the other choice says.
5. Save as `tests/fixtures/xlsx-pivots/interaction/owner-overlap-value.xlsx`.

Expected per spec §7.1: "A PivotTable report cannot overlap…" (the zh-TW text is what we need), with the pivot unchanged.

## O2 — expand into a formatting-only cell (§7.1 [measure])

1. Open a copy of **C**.
2. Select **A23** on sheet Pivot. Give it a yellow fill (常用 → 填滿色彩), and leave it empty.
3. Expand East.
4. Record whether an alert appears, and its text. Record whether the pivot grew over A23, and whether A23 keeps its yellow fill afterwards.
5. Save as `tests/fixtures/xlsx-pivots/interaction/owner-overlap-format.xlsx`.

This decides whether Task 29's overlap rule counts a formatting-only cell as covered.

## O3 — expand into another pivot (§7.1)

1. Open a copy of **C**.
2. Select **A23** on sheet Pivot → 插入 → 樞紐分析表 (Insert → PivotTable).
   - Source: `Data!$A$1:$L$57`.
   - Location: existing worksheet, `Pivot!$A$23`.
3. Put **Region** on Rows and **Sales** in Values. The second pivot then occupies about A23:B30.
4. Expand East in the first pivot.
5. Record the alert and its buttons, and whether either pivot changed.
6. Save as `tests/fixtures/xlsx-pivots/interaction/owner-overlap-pivot.xlsx`.

## O4 — shrink next to a non-empty cell (control for O1)

1. Open a copy of **S**.
2. On sheet Pivot, type `beside` into **J10**, the column just right of the pivot.
3. Collapse East. The pivot shrinks by 3 rows and does not widen.
4. Confirm that no alert appears, and that `beside` stays in J10.
5. Save as `tests/fixtures/xlsx-pivots/interaction/owner-shrink-beside.xlsx`.

This confirms that a shrink never alerts, and that cells outside the rectangle are untouched.

## O5 — sheet sort across a pivot (§7.9 [measure])

1. Open a copy of **S**.
2. On sheet Pivot, select **A1:K30**. That covers the pivot plus empty cells around it.
3. Choose 資料 → 排序 (Data → Sort). Pick any key (column B, ascending) → OK.
4. Record the alert text.
5. Repeat with a selection that lies only inside the pivot's value cells (**B6:B10**), and sort ascending. Record the alert text, or the resulting order if Excel allows it.
6. Save as `tests/fixtures/xlsx-pivots/interaction/owner-sheet-sort.xlsx`.

## O6 — sheet filter across a pivot (§7.9 [measure])

1. Open a copy of **S**.
2. Select **A5:I24** on sheet Pivot (the pivot's rows) → 資料 → 篩選 (Data → Filter).
3. Record the alert text, or what Excel does (no filter buttons, or filter buttons on the pivot's header row).
4. Select **K5:L10** (empty cells beside the pivot) → 資料 → 篩選, and record whether that is allowed.
5. Save as `tests/fixtures/xlsx-pivots/interaction/owner-sheet-filter.xlsx`.

## O7 (optional) — a value filter on a field that has a label filter

Task 25 refuses a second filter from another family on the same field. This recipe records what Excel's own dialog
does when `multipleFieldFilters="0"` (the seed's setting). XML cannot reproduce that dialog behaviour.

1. Open a copy of `tests/fixtures/xlsx-pivots/engine/filter-label-begins-with.xlsx` (Product: label begins with "B").
2. Open the Product filter in the Row Labels dropdown (選取欄位: Product) → 值篩選 (Value Filters) → 大於 (Greater Than).
   Use 加總 - Sales / Sum of Sales, greater than **3000** → OK.
3. Record whether the label filter is still listed (the dropdown shows a ✓ next to 標籤篩選), and which Products remain.
4. Save as `tests/fixtures/xlsx-pivots/interaction/owner-filter-two-families.xlsx`.

## Hand-back

Put the saved files in `tests/fixtures/xlsx-pivots/interaction/`, plus a short note giving, for each recipe id:
- the alert text;
- the buttons;
- the result.

The orchestrator writes these into `docs/validation/xlsx-pivots.md` ("Stage 3 interaction oracle and measurements")
as the overlap and sheet-sort rules that Tasks 28 and 29 cite.
