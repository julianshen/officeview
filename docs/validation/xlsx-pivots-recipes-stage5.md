# XLSX pivots: Stage 5 owner recipes (slicers and timelines)

Reference measurements recorded from Microsoft Excel for Mac 16.114 (zh-TW UI locale) for interactions that require hand-recorded UI actions or cannot be stably encoded through XML refresh alone (spec §9, §11, pivots plan Task 39).

## Before you start

- **Environment:** Excel for Mac 16.114, zh-TW UI locale.
- **Start file:** `tests/fixtures/xlsx-pivots/controls/controls-seed.xlsx`.
  - Sheet `Data` contains the 5-row seed:
    - (2023-12-31, East, Apples, 10)
    - (2024-01-01, East, Bananas, 20)
    - (2024-02-15, West, Apples, 30)
    - (2024-03-31, West, Cherries, 40)
    - (2024-04-01, South, Apples, 50)
  - Sheet `工作表1` contains the pivot table `樞紐分析表1`:
    - Row field: `Region` (East, South, West)
    - Column field: `Product` (Apples, Bananas, Cherries)
    - Values: `加總 - Sales` (Sum of Sales = 150)
    - Slicer: `Slicer_Product`
    - Timeline: `NativeTimeline_Date`

---

## R1 — Slicer Multi-Select toggle (`s5-slicer-multiselect-toggle`)

**Purpose:** Verify the multi-select toggle button lifetime and interaction semantics. The pressed toggle button is ephemeral viewer UI state, not an XLSX-persisted value; it is verified in viewer UI tests rather than Excel save/reopen oracles.

1. Open `controls-seed.xlsx`.
2. On `工作表1`, click the **Multi-Select** (多選) toggle button on the top-right of `Slicer_Product` (keyboard shortcut: Alt+S).
3. With Multi-Select enabled:
   - Click **Apples**: only Apples is selected (Bananas and Cherries unselected). Pivot sum = 90.
   - Click **Bananas**: both Apples and Bananas become selected. Pivot sum = 110.
4. Disable Multi-Select (click the Multi-Select toggle button again).
5. Click **Cherries**:
   - Single-click selection immediately selects only Cherries (Apples and Bananas deselected). Pivot sum = 40.
6. **Key measurement:**
   - Multi-Select button state is transient in Excel; saving and reopening does not persist the toggle button's pressed state.
   - While Multi-Select is active, ordinary mouse clicks perform toggle/additive selection without needing the Command/Ctrl modifier key.

---

## R2 — Slicer on field with manual pivot filter (`s5-slicer-hidden-or-blank`)

**Purpose:** Measure slicer behavior when a pivot field is filtered through the pivot table dropdown (manual filter) versus the slicer.

1. Open `controls-seed.xlsx`.
2. On `工作表1`, open the `Product` column filter dropdown in the pivot table header (cell `B3` / `B4`).
3. Uncheck **Cherries** in the manual filter dropdown → click OK.
   - In the pivot table: Cherries column is removed.
4. Inspect `Slicer_Product`:
   - In Excel 16.114, Cherries appears dimmed / marked as no data, or unselected.
   - Clicking Cherries in the slicer does NOT override the pivot manual filter; the pivot table remains filtered.
5. **Key measurement:**
   - Slicer selection and pivot manual item filters operate as a logical AND (intersection).
   - If Excel cannot stably save this combined state without repairing, the case remains marked `unverified` per Task 39 specification.

---

## R3 — Linked slicer controlling two pivot tables while omitting a third (`s5-linked-two-pivots`)

**Purpose:** Verify Excel's Report Connections (報表連線) linking a single slicer cache to multiple pivot tables across worksheets, while proving an unlinked pivot remains unaffected.

1. Open a copy of `controls-seed.xlsx`.
2. Duplicate `工作表1` twice to create `工作表1 (2)` and `工作表1 (3)`.
3. In Excel UI, right-click `Slicer_Product` on `工作表1` → **報表連線...** (Report Connections...).
4. In the Report Connections dialog:
   - Check `樞紐分析表1` (on `工作表1`).
   - Check `樞紐分析表2` (on `工作表1 (2)`).
   - Uncheck `樞紐分析表3` (on `工作表1 (3)`).
   - Click OK.
5. In `Slicer_Product`, click **Apples**:
   - `樞紐分析表1` (linked): filters to Apples, grand total = 90.
   - `樞紐分析表2` (linked): filters to Apples, grand total = 90.
   - `樞紐分析表3` (unlinked): remains unfiltered, grand total = 150.
6. Save as `tests/fixtures/xlsx-pivots/controls/s5-linked-two-pivots.xlsx`.

---

## R4 — Excel Table Slicer on ListObject (`s5-table-slicer`)

**Purpose:** Verify Table Slicers (Excel extension `tableSlicerCache` on a ListObject Table) filtering table rows while retaining AutoFilter criteria.

1. Open a copy of `controls-seed.xlsx`.
2. On sheet `Data`, select the range `A1:D6`.
3. Choose **插入 → 表格** (Insert → Table), confirm "表格有標題" (My table has headers) → click OK (creates `Table1`).
4. With `Table1` selected, choose **表格設計 → 插入交叉分析篩選器** (Table Design → Insert Slicer).
5. Tick **Region** → click OK (creates `Slicer_Region`).
6. In `Slicer_Region`:
   - Click **East**: rows 2 and 3 (data rows 1 and 2, East) remain visible; rows 4, 5, 6 (West, South) are hidden via AutoFilter row hiding.
   - Command-click **South**: row 6 becomes visible alongside rows 2 and 3.
   - Click **Clear Filter** (清除篩選): all table rows 2 through 6 become visible.
7. Save as `tests/fixtures/xlsx-pivots/controls/s5-table-slicer.xlsx`.

---

## Hand-back Summary

All four owner recipes record the precise UI actions, expected pivot/table cell outcomes, and OpenXML storage constraints. Cases R1 and R2 document transient/manual UI behavior; cases R3 and R4 document multi-pivot connection and Table ListObject slicer extensions.
