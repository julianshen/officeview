# Stage 1 owner recipes

These are cases that Excel cannot be scripted to make (spec §11). Save each result over its named file. Excel's UI language does not matter.

## R1: slicer and timeline template

**Purpose:** an Excel-written slicer and timeline, used as the structural reference for the hand-written cases. Excel removed our hand-written ones.

1. Open `tests/fixtures/xlsx-pivots/recipes/r1-source.xlsx`. The data is on sheet **Data**, A1:E41.
2. Click inside the data. Choose **Insert → PivotTable**, then **New Worksheet**, then OK.
3. In the field list, drag **Region** to Rows and **Sales** to Values (Sum of Sales).
4. With the pivot selected, choose **PivotTable Analyze → Insert Slicer**, tick **Product**, then OK.
5. Choose **PivotTable Analyze → Insert Timeline**, tick **Date**, then OK.
6. In the slicer, click **Apples**, then Ctrl/⌘-click **Cherries**, so two items are selected. In the timeline, select any two consecutive months.
7. Save as `tests/fixtures/xlsx-pivots/recipes/r1-slicer-timeline.xlsx`, in Excel Workbook (.xlsx) format, then close.
