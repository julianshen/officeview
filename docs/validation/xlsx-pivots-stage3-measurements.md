# XLSX pivots: Stage 3 raw measurements (generated)

Output of `bun scripts/xlsx-refs/measure-interaction.ts` over the Task 24 interaction oracle (2026-10-31, Excel 16.114, zh-TW). The rules drawn from it are in `xlsx-pivots.md`, "Stage 3 interaction oracle and measurements".


### (a) Column widths after a change (§7.8 autofit input)

Pivot sheet `<col>` widths (character units; `—` = no `<col>`, the sheet default) over the union of the start and saved pivot columns. "longest" is the longest saved cell text in that column of the case (numbers as stored, before number formatting).

**interaction-base-tabular-bottom** (start recipes/engine-seed, location A3:I24 → A3:J24; setup: tabular form, subtotals at the bottom (Stage 2 has no such file))

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A24) |
| B | 15.5 bf | 11 bf | -4.5 | Elderberries (12, B22) |
| C | 11.5 bf | 11.5 bf | 0 | 4217.1099999999997 (18, C7) |
| D | 11.33203125 bf | 11.5 bf | 0.168 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 11520.460000000001 (18, E7) |
| F | 11.33203125 bf | 11.5 bf | 0.168 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 10137.869999999999 (18, G12) |
| H | 17.83203125 bf | 11.5 bf | -6.332 | Sum of Units (12, H5) |
| I | 18 bf | 17.83203125 bf | -0.168 | 17787.900000000001 (18, I7) |
| J | — | 18 bf |  | Sum of Units 的加總 (16, J4) |

**interaction-collapse-item-compact-top** (start recipes/engine-seed, location A3:I24 → A3:I21) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A20) |
| B | 15.5 bf | 15.5 bf | 0 | 2413.6800000000003 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-collapse-field-compact-top** (start recipes/engine-seed, location A3:I24 → A3:I11)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 12.83203125 bf | -1.832 | Grand Total (11, A11) |
| B | 15.5 bf | 15.5 bf | 0 | 14898.060000000001 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-collapse-item-compact-bottom** (start engine/layout-subtotals-bottom-compact, location A3:I29 → A3:I25) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 2413.6800000000003 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 10137.869999999999 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-collapse-field-compact-bottom** (start engine/layout-subtotals-bottom-compact, location A3:I29 → A3:I11)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 12.83203125 bf | -1.832 | Grand Total (11, A11) |
| B | 15.5 bf | 15.5 bf | 0 | 14898.060000000001 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-collapse-item-outline-top** (start engine/layout-subtotals-top-outline, location A3:J24 → A3:J21) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A21) |
| B | 11 bf | 11 bf | 0 | Elderberries (12, B20) |
| C | 11.5 bf | 11.5 bf | 0 | 2413.6800000000003 (18, C9) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 23185.739999999998 (18, G7) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**interaction-collapse-field-outline-top** (start engine/layout-subtotals-top-outline, location A3:J24 → A3:J11)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A11) |
| B | 11 bf | 10.5 bf | -0.5 | Product (7, B5) |
| C | 11.5 bf | 11.5 bf | 0 | 14898.060000000001 (18, C9) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 23185.739999999998 (18, G7) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**interaction-collapse-item-outline-bottom** (start engine/layout-subtotals-bottom-outline, location A3:J29 → A3:J25) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A25) |
| B | 11 bf | 11 bf | 0 | Elderberries (12, B23) |
| C | 11.5 bf | 11.5 bf | 0 | 2413.6800000000003 (18, C9) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 10137.869999999999 (18, G10) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**interaction-collapse-field-outline-bottom** (start engine/layout-subtotals-bottom-outline, location A3:J29 → A3:J11)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A11) |
| B | 11 bf | 10.5 bf | -0.5 | Product (7, B5) |
| C | 11.5 bf | 11.5 bf | 0 | 14898.060000000001 (18, C9) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 23185.739999999998 (18, G7) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**interaction-collapse-item-tabular-top** (start engine/layout-subtotals-top-tabular, location A3:J24 → A3:J21) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A21) |
| B | 11 bf | 11 bf | 0 | Elderberries (12, B19) |
| C | 11.5 bf | 11.5 bf | 0 | 2413.6800000000003 (18, C8) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 10137.869999999999 (18, G9) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**interaction-collapse-field-tabular-top** (start engine/layout-subtotals-top-tabular, location A3:J24 → A3:J11)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A11) |
| B | 11 bf | 10.5 bf | -0.5 | Product (7, B5) |
| C | 11.5 bf | 11.5 bf | 0 | 14898.060000000001 (18, C9) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 23185.739999999998 (18, G7) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**interaction-collapse-item-tabular-bottom** (start interaction/interaction-base-tabular-bottom, location A3:J24 → A3:J21) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A21) |
| B | 11 bf | 11 bf | 0 | Elderberries (12, B19) |
| C | 11.5 bf | 11.5 bf | 0 | 2413.6800000000003 (18, C8) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 10137.869999999999 (18, G9) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**interaction-collapse-field-tabular-bottom** (start interaction/interaction-base-tabular-bottom, location A3:J24 → A3:J11)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A11) |
| B | 11 bf | 10.5 bf | -0.5 | Product (7, B5) |
| C | 11.5 bf | 11.5 bf | 0 | 14898.060000000001 (18, C9) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 23185.739999999998 (18, G7) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**interaction-collapse-all-compact-top** (start recipes/engine-seed, location A3:I24 → A3:I11)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 12.83203125 bf | -1.832 | Grand Total (11, A11) |
| B | 15.5 bf | 15.5 bf | 0 | 14898.060000000001 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-collapse-all-outline-bottom** (start engine/layout-subtotals-bottom-outline, location A3:J29 → A3:J11)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Grand Total (11, A11) |
| B | 11 bf | 10.5 bf | -0.5 | Product (7, B5) |
| C | 11.5 bf | 11.5 bf | 0 | 14898.060000000001 (18, C9) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 23185.739999999998 (18, G7) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**interaction-expand-item** (start engine/layout-collapsed-item, location A3:I21 → A3:I24) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-expand-field** (start interaction/interaction-collapse-field-compact-top, location A3:I11 → A3:I24)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 14.6640625 bf | 1.832 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-expand-all** (start interaction/interaction-collapse-all-compact-top, location A3:I11 → A3:I24)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 14.6640625 bf | 1.832 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-filter-all-hidden** (start recipes/engine-seed, location A3:I24 → A3:I10)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 12.83203125 bf | -1.832 | Grand Total (11, A10) |
| B | 15.5 bf | 15.5 bf | 0 | 2413.6800000000003 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 23775.739999999998 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F6) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 53459.319999999992 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-items-inner-all-but-one** (start recipes/engine-seed, location A3:I24 → A3:I14)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 12.83203125 bf | -1.832 | Grand Total (11, A14) |
| B | 15.5 bf | 15.5 bf | 0 | 1766.7599999999998 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 8332.880000000001 (17, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 10137.869999999999 (18, F8) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 20745.830000000002 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-items-show-again** (start engine/quirk-hidden-row-item, location A3:I20 → A3:I24) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-clear-filters-hidden-items** (start engine/quirk-hidden-row-item, location A3:I20 → A3:I24) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-page-select-other** (start engine/quirk-page-single, location A3:I10 → A3:I10) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A10) |
| B | 15.5 bf | 15.5 bf | 0 | 1766.7599999999998 (18, B7) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 8332.880000000001 (17, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 10137.869999999999 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 20745.830000000002 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-page-all** (start engine/quirk-page-single, location A3:I10 → A3:I11) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A11) |
| B | 15.5 bf | 15.5 bf | 0 | 14898.060000000001 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-page-multi** (start engine/quirk-page-single, location A3:I10 → A3:I10) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A10) |
| B | 15.5 bf | 15.5 bf | 0 | 4084.1599999999994 (18, B7) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 24313.839999999997 (18, D9) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 11481.789999999999 (18, F9) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 30331.059999999998 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-page-multi-to-all** (start engine/quirk-page-multiple, location A3:I10 → A3:I11) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A11) |
| B | 15.5 bf | 15.5 bf | 0 | 14898.060000000001 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-clear-label-filter** (start engine/filter-label-begins-with, location A3:I14 → A3:I24)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 14.6640625 bf | 1.832 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-clear-value-filter** (start engine/filter-value-greater-than, location A3:I14 → A3:I24)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 14.6640625 bf | 1.832 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-clear-top10-filter** (start engine/filter-top-count, location A3:I14 → A3:I24)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 14.6640625 bf | 1.832 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-filter-date-base** (start recipes/engine-seed, location A3:I24 → A3:I59; setup: Date alone on rows (Region and Product off the axis))

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 12.83203125 bf | -1.832 | Grand Total (11, A59) |
| B | 15.5 bf | 15.5 bf | 0 | 2101.2399999999998 (18, B47) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 4331.1400000000003 (18, D12) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 9266.3799999999992 (18, F58) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 4331.1400000000003 (18, H12) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-filter-date-between** (start interaction/interaction-filter-date-base, location A3:I59 → A3:I23) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A23) |
| B | 15.5 bf | 15.5 bf | 0 | Column Labels (13, B3) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 30416.870000000003 (18, D23) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | Sum of Sales (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 58609.090000000004 (18, H23) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-filter-date-clear** (start engine/filter-date-between, location A3:I23 → A3:I59) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A59) |
| B | 15.5 bf | 15.5 bf | 0 | 2101.2399999999998 (18, B47) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 4331.1400000000003 (18, D12) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 9266.3799999999992 (18, F58) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 4331.1400000000003 (18, H12) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-sort-value-ascending-ties** (start recipes/engine-seed, location A3:I24 → A3:I24) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A7) |
| B | 15.5 bf | 15.5 bf | 0 | 14898.060000000001 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 31949.339999999997 (18, D8) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 14651.759999999998 (18, F8) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 61499.159999999989 (18, H8) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**interaction-sort-then-collapse** (start recipes/engine-seed, location A3:I24 → A3:I21) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A20) |
| B | 15.5 bf | 15.5 bf | 0 | 14898.060000000001 (18, B6) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 4331.1400000000003 (18, D8) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 14651.759999999998 (18, F6) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 10353.620000000001 (18, H8) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**caption-defaults** (start recipes/engine-seed, location A3:I24 → A3:I24; setup: the seed's explicit grandTotalCaption, rowHeaderCaption and colHeaderCaption removed)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A23) |
| B | 15.5 bf | 11.33203125 bf | -4.168 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**caption-defaults-single-data** (start recipes/engine-seed, location A3:I24 → A3:E23; setup: Sum of Units removed from Values; explicit captions removed)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Sum of Sales (12, A3) |
| B | 15.5 bf | 10 bf | -5.5 | 4217.1099999999997 (18, B7) |
| C | 11.5 bf | 10 bf | -1.5 | 22019.200000000004 (18, C5) |
| D | 11.33203125 bf | 10 bf | -1.332 | 23185.739999999998 (18, D9) |
| E | 11.5 bf | 11 bf | -0.5 | 48118.960000000006 (18, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 |  |
| G | 11.5 bf | 11.5 bf | 0 |  |
| H | 17.83203125 bf | 17.83203125 bf | 0 |  |
| I | 18 bf | 18 bf | 0 |  |

**caption-defaults-outline-bottom** (start recipes/engine-seed, location A3:I24 → A3:J29; setup: outline form, subtotals at the bottom, explicit captions removed)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Central 合計 (10, A28) |
| B | 15.5 bf | 11 bf | -4.5 | Elderberries (12, B27) |
| C | 11.5 bf | 11.5 bf | 0 | 4217.1099999999997 (18, C8) |
| D | 11.33203125 bf | 11.5 bf | 0.168 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 11520.460000000001 (18, E8) |
| F | 11.33203125 bf | 11.5 bf | 0.168 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 10137.869999999999 (18, G14) |
| H | 17.83203125 bf | 11.5 bf | -6.332 | Sum of Units (12, H5) |
| I | 18 bf | 17.83203125 bf | -0.168 | 17787.900000000001 (18, I8) |
| J | — | 18 bf |  | Sum of Units 的加總 (16, J4) |

**caption-defaults-collapsed-bottom** (start interaction/caption-defaults-outline-bottom, location A3:J29 → A3:J25) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Central 合計 (10, A24) |
| B | 11 bf | 11 bf | 0 | Elderberries (12, B23) |
| C | 11.5 bf | 11.5 bf | 0 | 2413.6800000000003 (18, C9) |
| D | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | 22019.200000000001 (18, E6) |
| F | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | 10137.869999999999 (18, G10) |
| H | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, H5) |
| I | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, I6) |
| J | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, J4) |

**autofit-refresh-narrow** (start recipes/engine-seed, location A3:I24 → A3:I24; setup: columns A:K set to width 4 before the refresh) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**autofit-collapse-narrow** (start recipes/engine-seed, location A3:I24 → A3:I21; setup: columns A:K set to width 4 before the collapse) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A20) |
| B | 15.5 bf | 15.5 bf | 0 | 2413.6800000000003 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**autofit-collapse-wide** (start recipes/engine-seed, location A3:I24 → A3:I21; setup: columns A:K set to width 40 before the collapse) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A20) |
| B | 15.5 bf | 15.5 bf | 0 | 2413.6800000000003 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**autofit-hide-longest** (start recipes/engine-seed, location A3:I24 → A3:I22)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 12.83203125 bf | -1.832 | Grand Total (11, A22) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**chart-baseline** (start recipes/pivot-chart-seed, location A3:I22 → A3:I22) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A22) |
| B | 15.5 bf | 15.5 bf | 0 | 9068.8100000000013 (18, B6) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 8699.2800000000007 (18, D14) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 13600.380000000001 (18, F6) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 10168.450000000001 (18, H7) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**chart-hide-item** (start recipes/pivot-chart-seed, location A3:I22 → A3:I18) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A18) |
| B | 15.5 bf | 15.5 bf | 0 | 9068.8100000000013 (18, B6) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 8699.2800000000007 (18, D14) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 13600.380000000001 (18, F6) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 10168.450000000001 (18, H7) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**chart-collapse-item** (start recipes/pivot-chart-seed, location A3:I22 → A3:I19) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A19) |
| B | 15.5 bf | 15.5 bf | 0 | 9068.8100000000013 (18, B6) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 8699.2800000000007 (18, D11) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 13600.380000000001 (18, F6) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 36597.710000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**chart-value-filter** (start recipes/pivot-chart-seed, location A3:I22 → A3:I14) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A14) |
| B | 15.5 bf | 15.5 bf | 0 | 9068.8100000000013 (18, B6) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 8699.2800000000007 (18, D10) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 13600.380000000001 (18, F6) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 10168.450000000001 (18, H7) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**chart-sort** (start recipes/pivot-chart-seed, location A3:I22 → A3:I22) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A22) |
| B | 15.5 bf | 15.5 bf | 0 | 10519.900000000001 (18, B6) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 8699.2800000000007 (18, D10) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 9234.2000000000007 (18, F12) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 29762.620000000003 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**chart-values-on-rows** (start recipes/pivot-chart-seed, location A3:I22 → A3:E54; setup: Values moved to rows (a field move))

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 18 bf | 5.168 | North Sum of Sales (18, A27) |
| B | 15.5 bf | 15.5 bf | 0 | 4548.3500000000004 (18, B13) |
| C | 11.5 bf | 10 bf | -1.5 | 8699.2800000000007 (18, C34) |
| D | 11.33203125 bf | 10 bf | -1.332 | 13600.380000000001 (18, D15) |
| E | 11.5 bf | 10.83203125 bf | -0.668 | 10168.450000000001 (18, E7) |
| F | 11.33203125 bf | 11.33203125 bf | 0 |  |
| G | 11.5 bf | 11.5 bf | 0 |  |
| H | 17.83203125 bf | 17.83203125 bf | 0 |  |
| I | 18 bf | 18 bf | 0 |  |

**chart-page-base** (start recipes/pivot-chart-seed, location A3:I22 → A3:I10; setup: Product moved to a page field (a field move)) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A10) |
| B | 15.5 bf | 15.5 bf | 0 | 9068.8100000000013 (18, B6) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 8699.2800000000007 (18, D8) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 13600.380000000001 (18, F6) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 36597.710000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**chart-page-select** (start interaction/chart-page-base, location A3:I10 → A3:I10) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 12.83203125 bf | 12.83203125 bf | 0 | Grand Total (11, A10) |
| B | 15.5 bf | 15.5 bf | 0 | 21582.539999999997 (18, B10) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | Sum of Sales (12, D5) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 15766.630000000001 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 10168.450000000001 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**gpd-baseline** (start recipes/getpivotdata-seed, location A3:I24 → A3:I24) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**gpd-hide-item** (start recipes/getpivotdata-seed, location A3:I24 → A3:I20) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A19) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**gpd-collapse-item** (start recipes/getpivotdata-seed, location A3:I24 → A3:I21) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A20) |
| B | 15.5 bf | 15.5 bf | 0 | 2413.6800000000003 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 48118.960000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**gpd-filter-out** (start recipes/getpivotdata-seed, location A3:I24 → A3:I14)

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 12.83203125 bf | -1.832 | Grand Total (11, A14) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B6) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 11520.460000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | Sum of Sales (12, F5) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 17787.900000000001 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**refresh-on-load** (start recipes/refresh-seed, location A3:I24 → A3:I24) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 46884.460000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**refresh-on-load-collapse** (start recipes/refresh-seed, location A3:I24 → A3:I21) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A20) |
| B | 15.5 bf | 15.5 bf | 0 | 2413.6800000000003 (18, B9) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000001 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F7) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 46884.460000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**refresh-on-load-nodata** (start recipes/refresh-seed-nodata, location A3:I24 → A3:I24) — widths unchanged

| col | start | saved | Δ | longest saved text (chars, cell) |
| --- | --- | --- | --- | --- |
| A | 14.6640625 bf | 14.6640625 bf | 0 | Elderberries (12, A23) |
| B | 15.5 bf | 15.5 bf | 0 | 4217.1099999999997 (18, B8) |
| C | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, C5) |
| D | 11.33203125 bf | 11.33203125 bf | 0 | 22019.200000000004 (18, D6) |
| E | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, E5) |
| F | 11.33203125 bf | 11.33203125 bf | 0 | 23185.739999999998 (18, F10) |
| G | 11.5 bf | 11.5 bf | 0 | Sum of Units (12, G5) |
| H | 17.83203125 bf | 17.83203125 bf | 0 | 46884.460000000006 (18, H6) |
| I | 18 bf | 18 bf | 0 | Sum of Units 的加總 (16, I4) |

**Stage 2 cases against engine/baseline** (columns whose width differs only)

| case | col | baseline | saved | Δ | baseline longest | saved longest |
| --- | --- | --- | --- | --- | --- | --- |
| layout-compact-three-rows | A | 14.6640625 | 15.1640625 | 0.5 | Elderberries (12) | Elderberries (12) |
| layout-compact-three-rows | B | 15.5 | 11.33203125 | -4.168 | 4217.1099999999997 (18) | 17787.900000000001 (18) |
| layout-outline-three-rows | A | 14.6640625 | 15.1640625 | 0.5 | Elderberries (12) | Grand Total (11) |
| layout-outline-three-rows | B | 15.5 | 11.33203125 | -4.168 | 4217.1099999999997 (18) | Elderberries (12) |
| layout-outline-three-rows | C | 11.5 | 7.83203125 | -3.668 | Sum of Units (12) | Year (4) |
| layout-tabular-three-rows | A | 14.6640625 | 15.1640625 | 0.5 | Elderberries (12) | Grand Total (11) |
| layout-tabular-three-rows | B | 15.5 | 11.33203125 | -4.168 | 4217.1099999999997 (18) | Elderberries 合計 (15) |
| layout-tabular-three-rows | C | 11.5 | 7.83203125 | -3.668 | Sum of Units (12) | Year (4) |
| layout-collapsed-item-above-values | A | 14.6640625 | 18 | 3.336 | Elderberries (12) | Central Sum of Sales (20) |
| layout-collapsed-item-above-values | C | 11.5 | 10 | -1.5 | Sum of Units (12) | 22019.200000000001 (18) |
| layout-collapsed-item-above-values | D | 11.33203125 | 10 | -1.332 | 22019.200000000004 (18) | 10137.869999999999 (18) |
| layout-collapsed-item-above-values | E | 11.5 | 11 | -0.5 | Sum of Units (12) | 48118.960000000006 (18) |
| layout-subtotals-top-outline | B | 15.5 | 11 | -4.5 | 4217.1099999999997 (18) | Elderberries (12) |
| layout-subtotals-top-outline | D | 11.33203125 | 11.5 | 0.168 | 22019.200000000004 (18) | Sum of Units (12) |
| layout-subtotals-top-outline | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| layout-subtotals-top-outline | H | 17.83203125 | 11.5 | -6.332 | 48118.960000000006 (18) | Sum of Units (12) |
| layout-subtotals-top-outline | I | 18 | 17.83203125 | -0.168 | Sum of Units 的加總 (16) | 48118.960000000006 (18) |
| layout-subtotals-top-outline | J | — | 18 |  |  | Sum of Units 的加總 (16) |
| layout-subtotals-bottom-outline | B | 15.5 | 11 | -4.5 | 4217.1099999999997 (18) | Elderberries (12) |
| layout-subtotals-bottom-outline | D | 11.33203125 | 11.5 | 0.168 | 22019.200000000004 (18) | Sum of Units (12) |
| layout-subtotals-bottom-outline | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| layout-subtotals-bottom-outline | H | 17.83203125 | 11.5 | -6.332 | 48118.960000000006 (18) | Sum of Units (12) |
| layout-subtotals-bottom-outline | I | 18 | 17.83203125 | -0.168 | Sum of Units 的加總 (16) | 17787.900000000001 (18) |
| layout-subtotals-bottom-outline | J | — | 18 |  |  | Sum of Units 的加總 (16) |
| layout-subtotals-top-tabular | B | 15.5 | 11 | -4.5 | 4217.1099999999997 (18) | Elderberries (12) |
| layout-subtotals-top-tabular | D | 11.33203125 | 11.5 | 0.168 | 22019.200000000004 (18) | Sum of Units (12) |
| layout-subtotals-top-tabular | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| layout-subtotals-top-tabular | H | 17.83203125 | 11.5 | -6.332 | 48118.960000000006 (18) | Sum of Units (12) |
| layout-subtotals-top-tabular | I | 18 | 17.83203125 | -0.168 | Sum of Units 的加總 (16) | 17787.900000000001 (18) |
| layout-subtotals-top-tabular | J | — | 18 |  |  | Sum of Units 的加總 (16) |
| layout-values-on-rows-two-data | A | 14.6640625 | 18 | 3.336 | Elderberries (12) | Central Sum of Sales (20) |
| layout-values-on-rows-two-data | C | 11.5 | 10 | -1.5 | Sum of Units (12) | 11520.460000000001 (18) |
| layout-values-on-rows-two-data | D | 11.33203125 | 10 | -1.332 | 22019.200000000004 (18) | 10137.869999999999 (18) |
| layout-values-on-rows-two-data | E | 11.5 | 11 | -0.5 | Sum of Units (12) | 17787.900000000001 (18) |
| layout-values-on-columns-no-column-field | B | 15.5 | 11.33203125 | -4.168 | 4217.1099999999997 (18) | 17787.900000000001 (18) |
| layout-blank-rows-tabular | B | 15.5 | 11 | -4.5 | 4217.1099999999997 (18) | Elderberries (12) |
| layout-blank-rows-tabular | D | 11.33203125 | 11.5 | 0.168 | 22019.200000000004 (18) | Sum of Units (12) |
| layout-blank-rows-tabular | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| layout-blank-rows-tabular | H | 17.83203125 | 11.5 | -6.332 | 48118.960000000006 (18) | Sum of Units (12) |
| layout-blank-rows-tabular | I | 18 | 17.83203125 | -0.168 | Sum of Units 的加總 (16) | 17787.900000000001 (18) |
| layout-blank-rows-tabular | J | — | 18 |  |  | Sum of Units 的加總 (16) |
| layout-repeat-labels-tabular | B | 15.5 | 11 | -4.5 | 4217.1099999999997 (18) | Elderberries (12) |
| layout-repeat-labels-tabular | D | 11.33203125 | 11.5 | 0.168 | 22019.200000000004 (18) | Sum of Units (12) |
| layout-repeat-labels-tabular | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| layout-repeat-labels-tabular | H | 17.83203125 | 11.5 | -6.332 | 48118.960000000006 (18) | Sum of Units (12) |
| layout-repeat-labels-tabular | I | 18 | 17.83203125 | -0.168 | Sum of Units 的加總 (16) | 17787.900000000001 (18) |
| layout-repeat-labels-tabular | J | — | 18 |  |  | Sum of Units 的加總 (16) |
| layout-repeat-labels-outline | B | 15.5 | 11 | -4.5 | 4217.1099999999997 (18) | Elderberries (12) |
| layout-repeat-labels-outline | D | 11.33203125 | 11.5 | 0.168 | 22019.200000000004 (18) | Sum of Units (12) |
| layout-repeat-labels-outline | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| layout-repeat-labels-outline | H | 17.83203125 | 11.5 | -6.332 | 48118.960000000006 (18) | Sum of Units (12) |
| layout-repeat-labels-outline | I | 18 | 17.83203125 | -0.168 | Sum of Units 的加總 (16) | 48118.960000000006 (18) |
| layout-repeat-labels-outline | J | — | 18 |  |  | Sum of Units 的加總 (16) |
| layout-named-subtotals-compact | C | 11.5 | 13 | 1.5 | Sum of Units (12) | 69.666666666666671 (18) |
| layout-named-subtotals-compact | D | 11.33203125 | 13 | 1.668 | 22019.200000000004 (18) | 11520.460000000001 (18) |
| layout-named-subtotals-compact | E | 11.5 | 13 | 1.5 | Sum of Units (12) | 46.166666666666664 (18) |
| layout-named-subtotals-tabular | B | 15.5 | 11 | -4.5 | 4217.1099999999997 (18) | Elderberries (12) |
| layout-named-subtotals-tabular | C | 11.5 | 13 | 1.5 | Sum of Units (12) | 4217.1099999999997 (18) |
| layout-named-subtotals-tabular | D | 11.33203125 | 13 | 1.668 | 22019.200000000004 (18) | 69.666666666666671 (18) |
| layout-named-subtotals-tabular | E | 11.5 | 13 | 1.5 | Sum of Units (12) | 11520.460000000001 (18) |
| layout-named-subtotals-tabular | F | 11.33203125 | 13 | 1.668 | 23185.739999999998 (18) | 46.166666666666664 (18) |
| layout-named-subtotals-tabular | G | 11.5 | 13 | 1.5 | Sum of Units (12) | 10137.869999999999 (18) |
| layout-named-subtotals-tabular | H | 17.83203125 | 13 | -4.832 | 48118.960000000006 (18) | Sum of Units (12) |
| layout-named-subtotals-tabular | I | 18 | 17.83203125 | -0.168 | Sum of Units 的加總 (16) | 17787.900000000001 (18) |
| layout-named-subtotals-tabular | J | — | 18 |  |  | 43.285714285714285 (18) |
| layout-merge-labels | B | 15.5 | 12.83203125 | -2.668 | 4217.1099999999997 (18) | Elderberries (12) |
| layout-merge-labels | C | 11.5 | 11.83203125 | 0.332 | Sum of Units (12) | 4217.1099999999997 (18) |
| layout-merge-labels | D | 11.33203125 | 11.83203125 | 0.5 | 22019.200000000004 (18) | Sum of Units (12) |
| layout-merge-labels | E | 11.5 | 11.83203125 | 0.332 | Sum of Units (12) | 11520.460000000001 (18) |
| layout-merge-labels | F | 11.33203125 | 11.83203125 | 0.5 | 23185.739999999998 (18) | Sum of Units (12) |
| layout-merge-labels | G | 11.5 | 11.83203125 | 0.332 | Sum of Units (12) | 10137.869999999999 (18) |
| layout-merge-labels | H | 17.83203125 | 11.83203125 | -6 | 48118.960000000006 (18) | Sum of Units (12) |
| layout-merge-labels | I | 18 | 17.83203125 | -0.168 | Sum of Units 的加總 (16) | 17787.900000000001 (18) |
| layout-merge-labels | J | — | 18 |  |  | Sum of Units 的加總 (16) |
| layout-classic-single-data | B | 15.5 | 11 | -4.5 | 4217.1099999999997 (18) | 17787.900000000001 (18) |
| quirk-blank-and-mixed-row-items | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| quirk-error-value-count | D | 11.33203125 | 10.83203125 | -0.5 | 22019.200000000004 (18) | 計數 - Sales (10) |
| quirk-error-value-count | F | 11.33203125 | 10.83203125 | -0.5 | 23185.739999999998 (18) | 計數 - Sales (10) |
| quirk-error-value-count | H | 17.83203125 | 17.1640625 | -0.668 | 48118.960000000006 (18) | 計數 - Sales 的加總 (14) |
| quirk-page-single | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| quirk-page-single-showall | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| quirk-page-multiple | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| quirk-cross-hidden-item | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| quirk-cross-hidden-item-showall | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| quirk-dates-1900-rows | A | 14.6640625 | 14.33203125 | -0.332 | Elderberries (12) | Grand Total (11) |
| values-fn-sum | D | 11.33203125 | 10.83203125 | -0.5 | 22019.200000000004 (18) | 22019.200000000004 (18) |
| values-fn-sum | F | 11.33203125 | 10.83203125 | -0.5 | 23185.739999999998 (18) | 23185.739999999998 (18) |
| values-fn-sum | H | 17.83203125 | 17.1640625 | -0.668 | 48118.960000000006 (18) | 48118.960000000006 (18) |
| values-fn-count | D | 11.33203125 | 10.83203125 | -0.5 | 22019.200000000004 (18) | 計數 - Sales (10) |
| values-fn-count | F | 11.33203125 | 10.83203125 | -0.5 | 23185.739999999998 (18) | 計數 - Sales (10) |
| values-fn-count | H | 17.83203125 | 17.1640625 | -0.668 | 48118.960000000006 (18) | 計數 - Sales 的加總 (14) |
| values-fn-average | D | 11.33203125 | 13 | 1.668 | 22019.200000000004 (18) | 4403.8399999999992 (18) |
| values-fn-average | F | 11.33203125 | 12.6640625 | 1.332 | 23185.739999999998 (18) | 5796.4349999999995 (18) |
| values-fn-average | H | 17.83203125 | 19.1640625 | 1.332 | 48118.960000000006 (18) | 4009.9133333333339 (18) |
| values-fn-max | D | 11.33203125 | 10.83203125 | -0.5 | 22019.200000000004 (18) | 4331.1400000000003 (18) |
| values-fn-max | F | 11.33203125 | 10.83203125 | -0.5 | 23185.739999999998 (18) | 最大 - Sales (10) |
| values-fn-max | H | 17.83203125 | 17.1640625 | -0.668 | 48118.960000000006 (18) | 最大 - Sales 的加總 (14) |
| values-fn-min | D | 11.33203125 | 10.83203125 | -0.5 | 22019.200000000004 (18) | 4331.1400000000003 (18) |
| values-fn-min | F | 11.33203125 | 10.83203125 | -0.5 | 23185.739999999998 (18) | 最小 - Sales (10) |
| values-fn-min | H | 17.83203125 | 17.1640625 | -0.668 | 48118.960000000006 (18) | 最小 - Sales 的加總 (14) |
| values-fn-product | D | 11.33203125 | 13.6640625 | 2.332 | 22019.200000000004 (18) | -4.4620641884309103E+24 (23) |
| values-fn-product | F | 11.33203125 | 13 | 1.668 | 23185.739999999998 (18) | 1.9850284868621663E+60 (22) |
| values-fn-product | H | 17.83203125 | 17.1640625 | -0.668 | 48118.960000000006 (18) | -1.4130418834699652E+185 (24) |
| values-fn-countnums | D | 11.33203125 | 10.83203125 | -0.5 | 22019.200000000004 (18) | 計數 - Sales (10) |
| values-fn-countnums | F | 11.33203125 | 10.83203125 | -0.5 | 23185.739999999998 (18) | 計數 - Sales (10) |
| values-fn-countnums | H | 17.83203125 | 17.1640625 | -0.668 | 48118.960000000006 (18) | 計數 - Sales 的加總 (14) |
| values-fn-stdev | D | 11.33203125 | 13 | 1.668 | 22019.200000000004 (18) | 3158.7676829421953 (18) |
| values-fn-stdev | F | 11.33203125 | 13 | 1.668 | 23185.739999999998 (18) | 1902.2784517169589 (18) |
| values-fn-stdev | H | 17.83203125 | 19.1640625 | 1.332 | 48118.960000000006 (18) | 2386.5781018259481 (18) |
| values-fn-stdevp | B | 15.5 | 16.83203125 | 1.332 | 4217.1099999999997 (18) | 1875.7896938326774 (18) |
| values-fn-stdevp | D | 11.33203125 | 16.83203125 | 5.5 | 22019.200000000004 (18) | 3226.2299999999991 (18) |
| values-fn-stdevp | F | 11.33203125 | 16.83203125 | 5.5 | 23185.739999999998 (18) | 1337.1455715908444 (18) |
| values-fn-stdevp | H | 17.83203125 | 23.33203125 | 5.5 | 48118.960000000006 (18) | 2284.9745694658286 (18) |
| values-fn-var | D | 11.33203125 | 13 | 1.668 | 22019.200000000004 (18) | 9977813.2748000063 (18) |
| values-fn-var | F | 11.33203125 | 13 | 1.668 | 23185.739999999998 (18) | 3618663.3078666702 (18) |
| values-fn-var | H | 17.83203125 | 19.1640625 | 1.332 | 48118.960000000006 (18) | 5695755.0361151453 (18) |
| values-fn-varp | B | 15.5 | 16.83203125 | 1.332 | 4217.1099999999997 (18) | 3518586.9754888895 (18) |
| values-fn-varp | D | 11.33203125 | 16.83203125 | 5.5 | 22019.200000000004 (18) | 7982250.6198400017 (18) |
| values-fn-varp | F | 11.33203125 | 16.83203125 | 5.5 | 23185.739999999998 (18) | 2713997.4809000026 (18) |
| values-fn-varp | H | 17.83203125 | 23.33203125 | 5.5 | 48118.960000000006 (18) | 5221108.7831055485 (18) |
| values-fn-countnums-mixed | D | 11.33203125 | 13.1640625 | 1.832 | 22019.200000000004 (18) | 計數 - Channel (12) |
| values-fn-countnums-mixed | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 22019.199999999997 (18) |
| values-fn-countnums-mixed | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| values-fn-countnums-mixed | G | 11.5 | 13.1640625 | 1.664 | Sum of Units (12) | 計數 - Channel (12) |
| values-fn-countnums-mixed | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 23185.739999999998 (18) |
| values-fn-countnums-mixed | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| values-fn-countnums-mixed | J | — | 13.1640625 |  |  | 計數 - Channel (12) |
| values-fn-countnums-mixed | K | — | 17.83203125 |  |  | 48118.960000000006 (18) |
| values-fn-countnums-mixed | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| values-fn-countnums-mixed | M | — | 19.6640625 |  |  | 計數 - Channel 的加總 (16) |
| values-average-hidden-item | D | 11.33203125 | 13 | 1.668 | 22019.200000000004 (18) | 4403.8399999999992 (18) |
| values-average-hidden-item | F | 11.33203125 | 13 | 1.668 | 23185.739999999998 (18) | 5796.4349999999995 (18) |
| values-average-hidden-item | H | 17.83203125 | 19.1640625 | 1.332 | 48118.960000000006 (18) | 4009.9133333333339 (18) |
| values-fn-max-of-date | D | 11.33203125 | 10.6640625 | -0.668 | 22019.200000000004 (18) | 最大 - Date (9) |
| values-fn-max-of-date | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 22019.199999999997 (18) |
| values-fn-max-of-date | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| values-fn-max-of-date | G | 11.5 | 10.6640625 | -0.836 | Sum of Units (12) | 最大 - Date (9) |
| values-fn-max-of-date | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 23185.739999999998 (18) |
| values-fn-max-of-date | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| values-fn-max-of-date | J | — | 10.6640625 |  |  | 最大 - Date (9) |
| values-fn-max-of-date | K | — | 17.83203125 |  |  | 48118.960000000006 (18) |
| values-fn-max-of-date | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| values-fn-max-of-date | M | — | 17 |  |  | 最大 - Date 的加總 (13) |
| values-addition-order | D | 11.33203125 | 17.5 | 6.168 | 22019.200000000004 (18) | 784.29680000000008 (18) |
| values-addition-order | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 22019.200000000004 (18) |
| values-addition-order | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| values-addition-order | G | 11.5 | 17.5 | 6 | Sum of Units (12) | 0.60000000000000009 (19) |
| values-addition-order | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 23185.739999999998 (18) |
| values-addition-order | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| values-addition-order | J | — | 17.5 |  |  | 432.94090000000006 (18) |
| values-addition-order | K | — | 17.83203125 |  |  | 48118.960000000006 (18) |
| values-addition-order | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| values-addition-order | M | — | 23.83203125 |  |  | Sum of Net Amount 的加總 (21) |
| values-showas-index | D | 11.33203125 | 13 | 1.668 | 22019.200000000004 (18) | 0.52188519740023775 (19) |
| values-showas-index | F | 11.33203125 | 13 | 1.668 | 23185.739999999998 (18) | 0.93378835189524123 (19) |
| values-showas-index-zero-total | C | 11.5 | 13 | 1.5 | Sum of Units (12) | 0.90596590654339704 (19) |
| values-showas-index-zero-total | E | 11.5 | 13 | 1.5 | Sum of Units (12) | 4.4505140402025474E-2 (21) |
| values-showas-index-zero-total | G | 11.5 | 13 | 1.5 | Sum of Units (12) | 7.1748745806987854E-2 (21) |
| values-showas-percent-of-parent-row-two-levels | A | 14.6640625 | 15.1640625 | 0.5 | Elderberries (12) | Elderberries (12) |
| values-showas-percent-of-parent-row-two-levels | B | 15.5 | 13 | -2.5 | 4217.1099999999997 (18) | 9.7922399674101906E-2 (21) |
| values-showas-percent-of-parent-column | D | 11.33203125 | 13 | 1.668 | 22019.200000000004 (18) | 0.45759924985909922 (19) |
| values-showas-percent-of-parent-column | F | 11.33203125 | 13 | 1.668 | 23185.739999999998 (18) | 6.619230761801187E-2 (20) |
| values-showas-percent-of-parent-region | A | 14.6640625 | 15.1640625 | 0.5 | Elderberries (12) | Elderberries (12) |
| values-showas-percent-of-parent-region | B | 15.5 | 13 | -2.5 | 4217.1099999999997 (18) | 2.5655167942116787E-2 (21) |
| values-showas-percent-running-total | D | 11.33203125 | 13 | 1.668 | 22019.200000000004 (18) | 0.69044384999177033 (19) |
| values-showas-on-named-subtotals | C | 11.5 | 13 | 1.5 | Sum of Units (12) | 69.666666666666671 (18) |
| values-showas-on-named-subtotals | E | 11.5 | 13 | 1.5 | Sum of Units (12) | 46.166666666666664 (18) |
| values-named-subtotals-intersecting | H | 17.83203125 | 17.33203125 | -0.5 | 48118.960000000006 (18) | 4217.1099999999997 (18) |
| values-named-subtotals-intersecting | I | 18 | 17.5 | -0.5 | Sum of Units 的加總 (16) | 52.647058823529413 (18) |
| values-named-subtotals-intersecting | J | — | 11.33203125 |  |  | 15054.399999999998 (18) |
| values-named-subtotals-intersecting | K | — | 11.5 |  |  | Sum of Units (12) |
| values-named-subtotals-intersecting | L | — | 11.33203125 |  |  | 8332.880000000001 (17) |
| values-named-subtotals-intersecting | M | — | 11.5 |  |  | Sum of Units (12) |
| values-named-subtotals-intersecting | N | — | 11.33203125 |  |  | 4331.1400000000003 (18) |
| values-named-subtotals-intersecting | O | — | 11.5 |  |  | Sum of Units (12) |
| values-named-subtotals-intersecting | P | — | 17.33203125 |  |  | 5760.2300000000005 (18) |
| values-named-subtotals-intersecting | Q | — | 17.5 |  |  | 42.31818181818182 (17) |
| values-named-subtotals-intersecting | R | — | 11.33203125 |  |  | Sum of Sales (12) |
| values-named-subtotals-intersecting | S | — | 11.5 |  |  | Sum of Units (12) |
| values-named-subtotals-intersecting | T | — | 11.33203125 |  |  | 26559.520000000004 (18) |
| values-named-subtotals-intersecting | U | — | 11.5 |  |  | Sum of Units (12) |
| values-named-subtotals-intersecting | V | — | 11.33203125 |  |  | 21640.879999999997 (18) |
| values-named-subtotals-intersecting | W | — | 11.5 |  |  | Sum of Units (12) |
| values-named-subtotals-intersecting | X | — | 11.33203125 |  |  | Sum of Sales (12) |
| values-named-subtotals-intersecting | Y | — | 11.5 |  |  | Sum of Units (12) |
| values-named-subtotals-intersecting | Z | — | 17.33203125 |  |  | 5068.9349999999995 (18) |
| values-named-subtotals-intersecting | AA | — | 17.5 |  |  | 34.764705882352942 (18) |
| values-named-subtotals-intersecting | AB | — | 17.83203125 |  |  | 17787.900000000001 (18) |
| values-named-subtotals-intersecting | AC | — | 18 |  |  | Sum of Units 的加總 (16) |
| sort-mixed-blank-ascending | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| sort-mixed-blank-descending | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| sort-custom-list-months | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| sort-custom-list-weekdays | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| sort-custom-list-off | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| sort-collation-ascending | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| sort-collation-descending | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| sort-grouped-field | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-label-begins-with | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-label-wildcard | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-label-between | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-label-collation | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-label-mixed-items | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-label-mixed-items | F | 11.33203125 | 17.83203125 | 6.5 | 23185.739999999998 (18) | 19220.669999999998 (18) |
| filter-label-mixed-items | G | 11.5 | 18 | 6.5 | Sum of Units (12) | Sum of Units 的加總 (16) |
| filter-label-renamed-item | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-label-cross-field | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-value-greater-than | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-value-between | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-value-showall | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-value-nested | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-top-count | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-top-count-ties | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-top-percent | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-top-percent-30 | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-top-sum | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| filter-date-between | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| group-date-years | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| group-date-quarters | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| group-date-months | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| group-date-years-quarters-months | A | 14.6640625 | 16.5 | 1.836 | Elderberries (12) | Grand Total (11) |
| group-date-out-of-range | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | >2023/12/31 (11) |
| group-number-range | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| group-number-out-of-range | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| group-number-fractional | A | 14.6640625 | 12.83203125 | -1.832 | Elderberries (12) | Grand Total (11) |
| calc-cell-without-records | D | 11.33203125 | 17.83203125 | 6.5 | 22019.200000000004 (18) | Sum of Double Sales (19) |
| calc-cell-without-records | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 22019.200000000004 (18) |
| calc-cell-without-records | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| calc-cell-without-records | G | 11.5 | 17.83203125 | 6.332 | Sum of Units (12) | Sum of Double Sales (19) |
| calc-cell-without-records | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 23185.739999999998 (18) |
| calc-cell-without-records | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| calc-cell-without-records | J | — | 17.83203125 |  |  | Sum of Double Sales (19) |
| calc-cell-without-records | K | — | 17.83203125 |  |  | 48118.960000000006 (18) |
| calc-cell-without-records | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| calc-cell-without-records | M | — | 24.1640625 |  |  | Sum of Double Sales 的加總 (23) |
| calc-division-by-zero | B | 15.5 | 11.33203125 | -4.168 | 4217.1099999999997 (18) | 17787.900000000001 (18) |
| calc-division-by-zero | D | 11.33203125 | 18.83203125 | 7.5 | 22019.200000000004 (18) | Sum of Price per Unit (21) |
| calc-quoted-field-name | D | 11.33203125 | 14.83203125 | 3.5 | 22019.200000000004 (18) | 7.0000000000000007E-2 (21) |
| calc-quoted-field-name | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 22019.200000000004 (18) |
| calc-quoted-field-name | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| calc-quoted-field-name | G | 11.5 | 14.83203125 | 3.332 | Sum of Units (12) | 7.0000000000000021E-2 (21) |
| calc-quoted-field-name | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 23185.739999999998 (18) |
| calc-quoted-field-name | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| calc-quoted-field-name | J | — | 14.83203125 |  |  | 7.0000000000000007E-2 (21) |
| calc-quoted-field-name | K | — | 17.83203125 |  |  | 48118.960000000006 (18) |
| calc-quoted-field-name | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| calc-quoted-field-name | M | — | 21.33203125 |  |  | 7.0000000000000007E-2 (21) |
| calc-references-calculated | B | 15.5 | 11.33203125 | -4.168 | 4217.1099999999997 (18) | 17787.900000000001 (18) |
| calc-references-calculated | D | 11.33203125 | 18 | 6.668 | 22019.200000000004 (18) | Sum of Double Price (19) |
| calc-constant | D | 11.33203125 | 14 | 2.668 | 22019.200000000004 (18) | Sum of Flat One (15) |
| calc-constant | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 22019.200000000004 (18) |
| calc-constant | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| calc-constant | G | 11.5 | 14 | 2.5 | Sum of Units (12) | Sum of Flat One (15) |
| calc-constant | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 23185.739999999998 (18) |
| calc-constant | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| calc-constant | J | — | 14 |  |  | Sum of Flat One (15) |
| calc-constant | K | — | 17.83203125 |  |  | 48118.960000000006 (18) |
| calc-constant | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| calc-constant | M | — | 20.6640625 |  |  | Sum of Flat One 的加總 (19) |
| calc-text-field | D | 11.33203125 | 19.33203125 | 8 | 22019.200000000004 (18) | Sum of Channel Twice (20) |
| calc-text-field | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 22019.200000000004 (18) |
| calc-text-field | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| calc-text-field | G | 11.5 | 19.33203125 | 7.832 | Sum of Units (12) | Sum of Channel Twice (20) |
| calc-text-field | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 23185.739999999998 (18) |
| calc-text-field | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| calc-text-field | J | — | 19.33203125 |  |  | Sum of Channel Twice (20) |
| calc-text-field | K | — | 17.83203125 |  |  | 48118.960000000006 (18) |
| calc-text-field | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| calc-text-field | M | — | 25.6640625 |  |  | Sum of Channel Twice 的加總 (24) |
| calc-function-count | D | 11.33203125 | 17.1640625 | 5.832 | 22019.200000000004 (18) | 8434.2199999999993 (18) |
| calc-function-count | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 22019.199999999997 (18) |
| calc-function-count | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| calc-function-count | G | 11.5 | 17.1640625 | 5.664 | Sum of Units (12) | 44038.399999999994 (18) |
| calc-function-count | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 23185.739999999998 (18) |
| calc-function-count | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| calc-function-count | J | — | 17.1640625 |  |  | 29791.040000000001 (18) |
| calc-function-count | K | — | 17.83203125 |  |  | 48118.960000000006 (18) |
| calc-function-count | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| calc-function-count | M | — | 23.6640625 |  |  | 計數 - Double Sales 的加總 (21) |
| calc-show-as | D | 11.33203125 | 17.83203125 | 6.5 | 22019.200000000004 (18) | 5.0332753288180362E-2 (21) |
| calc-show-as | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 22019.200000000004 (18) |
| calc-show-as | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| calc-show-as | G | 11.5 | 17.83203125 | 6.332 | Sum of Units (12) | 9.8916745910753515E-2 (21) |
| calc-show-as | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 23185.739999999998 (18) |
| calc-show-as | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| calc-show-as | J | — | 17.83203125 |  |  | 6.6915072620646859E-2 (21) |
| calc-show-as | K | — | 17.83203125 |  |  | 48118.960000000006 (18) |
| calc-show-as | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| calc-show-as | M | — | 24.1640625 |  |  | Sum of Double Sales 的加總 (23) |
| calc-named-subtotals | D | 11.33203125 | 17.83203125 | 6.5 | 22019.200000000004 (18) | Sum of Double Sales (19) |
| calc-named-subtotals | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 11520.460000000001 (18) |
| calc-named-subtotals | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| calc-named-subtotals | G | 11.5 | 17.83203125 | 6.332 | Sum of Units (12) | Sum of Double Sales (19) |
| calc-named-subtotals | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 10137.869999999999 (18) |
| calc-named-subtotals | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| calc-named-subtotals | J | — | 17.83203125 |  |  | Sum of Double Sales (19) |
| calc-named-subtotals | K | — | 17.83203125 |  |  | 17787.900000000001 (18) |
| calc-named-subtotals | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| calc-named-subtotals | M | — | 24.1640625 |  |  | Sum of Double Sales 的加總 (23) |
| calc-value-sort | D | 11.33203125 | 17.83203125 | 6.5 | 22019.200000000004 (18) | Sum of Double Sales (19) |
| calc-value-sort | E | 11.5 | 11.33203125 | -0.168 | Sum of Units (12) | 31949.339999999997 (18) |
| calc-value-sort | F | 11.33203125 | 11.5 | 0.168 | 23185.739999999998 (18) | Sum of Units (12) |
| calc-value-sort | G | 11.5 | 17.83203125 | 6.332 | Sum of Units (12) | Sum of Double Sales (19) |
| calc-value-sort | H | 17.83203125 | 11.33203125 | -6.5 | 48118.960000000006 (18) | 14651.759999999998 (18) |
| calc-value-sort | I | 18 | 11.5 | -6.5 | Sum of Units 的加總 (16) | Sum of Units (12) |
| calc-value-sort | J | — | 17.83203125 |  |  | Sum of Double Sales (19) |
| calc-value-sort | K | — | 17.83203125 |  |  | 61499.159999999989 (18) |
| calc-value-sort | L | — | 18 |  |  | Sum of Units 的加總 (16) |
| calc-value-sort | M | — | 24.1640625 |  |  | Sum of Double Sales 的加總 (23) |
| calc-item | A | 14.6640625 | 20.83203125 | 6.168 | Elderberries (12) | Apples and Bananas (18) |

### (b) Implicit captions by kind (caption pattern table input)

Every cell the engine marks `caption: 'implicit'` (computed on the saved state with no saved-caption lookup), with the text Excel saved there. Pattern variables: `{item}` the item label, `{data}` the data field name, `{field}` the source field name.

| kind | engine default (English) | saved pattern | first example | cells |
| --- | --- | --- | --- | --- |
| (blank) item | (blank) | (空白) | engine/quirk-blank-and-mixed-row-items!A10 = "(空白)" | 11 |
| classic Total | Total | 合計 | engine/layout-classic-single-data!B4 = "合計" | 1 |
| column labels | Column Labels | 欄標籤 | interaction/caption-defaults!B3 = "欄標籤" | 2 |
| grand total | Grand Total | 總計 | interaction/caption-defaults!A24 = "總計" | 5 |
| grand total on the Values axis | Total Sum of Sales | {data} 的加總 | engine/baseline!H4 = "Sum of Sales 的加總" | 339 |
| named subtotal (Average) | East Average | {item} 平均值 | engine/layout-named-subtotals-compact!A12 = "East 平均值" | 15 |
| named subtotal (Average) | 2022 Average | {item} 平均值 - Sales | engine/values-named-subtotals-intersecting!H4 = "2022 平均值 - Sales" | 3 |
| named subtotal (Average) | 2022 Average | {item} 平均值 - Units | engine/values-named-subtotals-intersecting!I4 = "2022 平均值 - Units" | 3 |
| named subtotal (Count) | East Count | {item} 計數 | engine/layout-named-subtotals-compact!A11 = "East 計數" | 25 |
| named subtotal (Sum) | East Sum | {item} 加總 | engine/layout-named-subtotals-compact!A10 = "East 加總" | 25 |
| page (All) | (All) | (全部) | interaction/interaction-page-all!B1 = "(全部)" | 3 |
| page (Multiple Items) | (Multiple Items) | (多重項目) | engine/quirk-page-multiple!B1 = "(多重項目)" | 2 |
| row labels | Row Labels | 列標籤 | interaction/caption-defaults!A5 = "列標籤" | 2 |
| subtotal | Apples Total | {item} 合計 | engine/layout-tabular-three-rows!B7 = "Apples 合計" | 78 |

Data-field names Excel assigned (explicit `dataField@name`, not English defaults):

| name | first file |
| --- | --- |
| 計數 - Sales | engine/quirk-error-value-count |
| 加總 - Sales | engine/values-fn-sum |
| 平均值 - Sales | engine/values-fn-average |
| 最大 - Sales | engine/values-fn-max |
| 最小 - Sales | engine/values-fn-min |
| 乘積 - Sales | engine/values-fn-product |
| 標準差 - Sales | engine/values-fn-stdev |
| 母體標準差 - Sales | engine/values-fn-stdevp |
| 變異數 - Sales | engine/values-fn-var |
| 母體變異值 - Sales | engine/values-fn-varp |
| 計數 - Channel | engine/values-fn-countnums-mixed |
| 最大 - Date | engine/values-fn-max-of-date |
| 計數 - Double Sales | engine/calc-function-count |

Not measured (engine refused or no records):

- engine/values-fn-stdev: engine refuses (pivot-engine-unmeasured)
- engine/values-fn-stdevp: engine refuses (pivot-engine-unmeasured)
- engine/values-fn-var: engine refuses (pivot-engine-unmeasured)
- engine/values-fn-varp: engine refuses (pivot-engine-unmeasured)
- engine/sort-value-showall-no-data: engine refuses (pivot-engine-unmeasured)
- engine/calc-item: engine refuses (pivot-calculated-item)
- interaction/refresh-on-load-nodata: no records (saveData="0")

### (c) Pivot chart series and categories (§6.6, CODEX-3)

**chart-seed (input)** xl/charts/chart1.xml: pivotSource [.pc-saved.xlsx]Pivot!樞紐分析表2 fmtId 0; 6 series; 6 c:pivotFmt; pivot chartFormats 6; location A3:I22

| idx | title | tx ref | cat ref | cat pts | val ref | val pts |
| --- | --- | --- | --- | --- | --- | --- |
| 0 | 2022 - Sum of Sales | Pivot!$B$3:$B$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$B$6:$B$22 | 12 |
| 1 | 2022 - Sum of Units | Pivot!$C$3:$C$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$C$6:$C$22 | 12 |
| 2 | 2023 - Sum of Sales | Pivot!$D$3:$D$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$D$6:$D$22 | 12 |
| 3 | 2023 - Sum of Units | Pivot!$E$3:$E$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$E$6:$E$22 | 12 |
| 4 | 2024 - Sum of Sales | Pivot!$F$3:$F$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$F$6:$F$22 | 12 |
| 5 | 2024 - Sum of Units | Pivot!$G$3:$G$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$G$6:$G$22 | 12 |

**chart-baseline** xl/charts/chart1.xml: pivotSource [chart-baseline.xlsx]Pivot!樞紐分析表2 fmtId 0; 6 series; 6 c:pivotFmt; pivot chartFormats 6; location A3:I22

| idx | title | tx ref | cat ref | cat pts | val ref | val pts |
| --- | --- | --- | --- | --- | --- | --- |
| 0 | 2022 - Sum of Sales | Pivot!$B$3:$B$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$B$6:$B$22 | 12 |
| 1 | 2022 - Sum of Units | Pivot!$C$3:$C$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$C$6:$C$22 | 12 |
| 2 | 2023 - Sum of Sales | Pivot!$D$3:$D$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$D$6:$D$22 | 12 |
| 3 | 2023 - Sum of Units | Pivot!$E$3:$E$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$E$6:$E$22 | 12 |
| 4 | 2024 - Sum of Sales | Pivot!$F$3:$F$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$F$6:$F$22 | 12 |
| 5 | 2024 - Sum of Units | Pivot!$G$3:$G$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$G$6:$G$22 | 12 |

**chart-hide-item** xl/charts/chart1.xml: pivotSource [chart-hide-item.xlsx]Pivot!樞紐分析表2 fmtId 0; 6 series; 6 c:pivotFmt; pivot chartFormats 6; location A3:I18

| idx | title | tx ref | cat ref | cat pts | val ref | val pts |
| --- | --- | --- | --- | --- | --- | --- |
| 0 | 2022 - Sum of Sales | Pivot!$B$3:$B$5 | Pivot!$A$6:$A$18 | 9 | Pivot!$B$6:$B$18 | 9 |
| 1 | 2022 - Sum of Units | Pivot!$C$3:$C$5 | Pivot!$A$6:$A$18 | 9 | Pivot!$C$6:$C$18 | 9 |
| 2 | 2023 - Sum of Sales | Pivot!$D$3:$D$5 | Pivot!$A$6:$A$18 | 9 | Pivot!$D$6:$D$18 | 9 |
| 3 | 2023 - Sum of Units | Pivot!$E$3:$E$5 | Pivot!$A$6:$A$18 | 9 | Pivot!$E$6:$E$18 | 9 |
| 4 | 2024 - Sum of Sales | Pivot!$F$3:$F$5 | Pivot!$A$6:$A$18 | 9 | Pivot!$F$6:$F$18 | 9 |
| 5 | 2024 - Sum of Units | Pivot!$G$3:$G$5 | Pivot!$A$6:$A$18 | 9 | Pivot!$G$6:$G$18 | 9 |

**chart-collapse-item** xl/charts/chart1.xml: pivotSource [chart-collapse-item.xlsx]Pivot!樞紐分析表2 fmtId 0; 6 series; 6 c:pivotFmt; pivot chartFormats 6; location A3:I19

| idx | title | tx ref | cat ref | cat pts | val ref | val pts |
| --- | --- | --- | --- | --- | --- | --- |
| 0 | 2022 - Sum of Sales | Pivot!$B$3:$B$5 | Pivot!$A$6:$A$19 | 10 | Pivot!$B$6:$B$19 | 10 |
| 1 | 2022 - Sum of Units | Pivot!$C$3:$C$5 | Pivot!$A$6:$A$19 | 10 | Pivot!$C$6:$C$19 | 10 |
| 2 | 2023 - Sum of Sales | Pivot!$D$3:$D$5 | Pivot!$A$6:$A$19 | 10 | Pivot!$D$6:$D$19 | 10 |
| 3 | 2023 - Sum of Units | Pivot!$E$3:$E$5 | Pivot!$A$6:$A$19 | 10 | Pivot!$E$6:$E$19 | 10 |
| 4 | 2024 - Sum of Sales | Pivot!$F$3:$F$5 | Pivot!$A$6:$A$19 | 10 | Pivot!$F$6:$F$19 | 10 |
| 5 | 2024 - Sum of Units | Pivot!$G$3:$G$5 | Pivot!$A$6:$A$19 | 10 | Pivot!$G$6:$G$19 | 10 |

**chart-value-filter** xl/charts/chart1.xml: pivotSource [chart-value-filter.xlsx]Pivot!樞紐分析表2 fmtId 0; 6 series; 6 c:pivotFmt; pivot chartFormats 6; location A3:I14

| idx | title | tx ref | cat ref | cat pts | val ref | val pts |
| --- | --- | --- | --- | --- | --- | --- |
| 0 | 2022 - Sum of Sales | Pivot!$B$3:$B$5 | Pivot!$A$6:$A$14 | 6 | Pivot!$B$6:$B$14 | 6 |
| 1 | 2022 - Sum of Units | Pivot!$C$3:$C$5 | Pivot!$A$6:$A$14 | 6 | Pivot!$C$6:$C$14 | 6 |
| 2 | 2023 - Sum of Sales | Pivot!$D$3:$D$5 | Pivot!$A$6:$A$14 | 6 | Pivot!$D$6:$D$14 | 6 |
| 3 | 2023 - Sum of Units | Pivot!$E$3:$E$5 | Pivot!$A$6:$A$14 | 6 | Pivot!$E$6:$E$14 | 6 |
| 4 | 2024 - Sum of Sales | Pivot!$F$3:$F$5 | Pivot!$A$6:$A$14 | 6 | Pivot!$F$6:$F$14 | 6 |
| 5 | 2024 - Sum of Units | Pivot!$G$3:$G$5 | Pivot!$A$6:$A$14 | 6 | Pivot!$G$6:$G$14 | 6 |

**chart-sort** xl/charts/chart1.xml: pivotSource [chart-sort.xlsx]Pivot!樞紐分析表2 fmtId 0; 6 series; 6 c:pivotFmt; pivot chartFormats 6; location A3:I22

| idx | title | tx ref | cat ref | cat pts | val ref | val pts |
| --- | --- | --- | --- | --- | --- | --- |
| 0 | 2022 - Sum of Sales | Pivot!$B$3:$B$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$B$6:$B$22 | 12 |
| 1 | 2022 - Sum of Units | Pivot!$C$3:$C$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$C$6:$C$22 | 12 |
| 2 | 2023 - Sum of Sales | Pivot!$D$3:$D$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$D$6:$D$22 | 12 |
| 3 | 2023 - Sum of Units | Pivot!$E$3:$E$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$E$6:$E$22 | 12 |
| 4 | 2024 - Sum of Sales | Pivot!$F$3:$F$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$F$6:$F$22 | 12 |
| 5 | 2024 - Sum of Units | Pivot!$G$3:$G$5 | Pivot!$A$6:$A$22 | 12 | Pivot!$G$6:$G$22 | 12 |

**chart-values-on-rows** xl/charts/chart1.xml: pivotSource [chart-values-on-rows.xlsx]Pivot!樞紐分析表2 fmtId 0; 3 series; 6 c:pivotFmt; pivot chartFormats 6; location A3:E54

| idx | title | tx ref | cat ref | cat pts | val ref | val pts |
| --- | --- | --- | --- | --- | --- | --- |
| 0 | 2022 | Pivot!$B$3:$B$4 | Pivot!$A$5:$A$54 | 24 | Pivot!$B$5:$B$54 | 24 |
| 1 | 2023 | Pivot!$C$3:$C$4 | Pivot!$A$5:$A$54 | 24 | Pivot!$C$5:$C$54 | 24 |
| 2 | 2024 | Pivot!$D$3:$D$4 | Pivot!$A$5:$A$54 | 24 | Pivot!$D$5:$D$54 | 24 |

**chart-page-base** xl/charts/chart1.xml: pivotSource [chart-page-base.xlsx]Pivot!樞紐分析表2 fmtId 0; 6 series; 6 c:pivotFmt; pivot chartFormats 6; location A3:I10

| idx | title | tx ref | cat ref | cat pts | val ref | val pts |
| --- | --- | --- | --- | --- | --- | --- |
| 0 | 2022 - Sum of Sales | Pivot!$B$3:$B$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$B$6:$B$10 | 4 |
| 1 | 2022 - Sum of Units | Pivot!$C$3:$C$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$C$6:$C$10 | 4 |
| 2 | 2023 - Sum of Sales | Pivot!$D$3:$D$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$D$6:$D$10 | 4 |
| 3 | 2023 - Sum of Units | Pivot!$E$3:$E$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$E$6:$E$10 | 4 |
| 4 | 2024 - Sum of Sales | Pivot!$F$3:$F$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$F$6:$F$10 | 4 |
| 5 | 2024 - Sum of Units | Pivot!$G$3:$G$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$G$6:$G$10 | 4 |

**chart-page-select** xl/charts/chart1.xml: pivotSource [chart-page-select.xlsx]Pivot!樞紐分析表2 fmtId 0; 6 series; 6 c:pivotFmt; pivot chartFormats 6; location A3:I10

| idx | title | tx ref | cat ref | cat pts | val ref | val pts |
| --- | --- | --- | --- | --- | --- | --- |
| 0 | 2022 - Sum of Sales | Pivot!$B$3:$B$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$B$6:$B$10 | 4 |
| 1 | 2022 - Sum of Units | Pivot!$C$3:$C$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$C$6:$C$10 | 4 |
| 2 | 2023 - Sum of Sales | Pivot!$D$3:$D$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$D$6:$D$10 | 4 |
| 3 | 2023 - Sum of Units | Pivot!$E$3:$E$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$E$6:$E$10 | 4 |
| 4 | 2024 - Sum of Sales | Pivot!$F$3:$F$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$F$6:$F$10 | 4 |
| 5 | 2024 - Sum of Units | Pivot!$G$3:$G$5 | Pivot!$A$6:$A$10 | 4 | Pivot!$G$6:$G$10 | 4 |

### (d) GETPIVOTDATA and references after a change (§7.5)

Cached values Excel saved on sheet Formulas, column B.

| formula | text | seed | gpd-baseline | gpd-hide-item | gpd-collapse-item | gpd-filter-out |
| --- | --- | --- | --- | --- | --- | --- |
| total | =GETPIVOTDATA("Sum of Sales",Pivot!$A$3) | 222603.36000000004 | 222603.36000000004 | 161104.20000000001 | 222603.36000000004 | 70490.05 |
| item | =GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","East") | 48118.960000000006 | 48118.960000000006 | 48118.960000000006 | 48118.960000000006 | 17787.900000000001 |
| two fields | =GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","East","Product","Apples") | 9585.23 | 9585.23 | 9585.23 | error #REF! | error #REF! |
| column item | =GETPIVOTDATA("Sum of Units",Pivot!$A$3,"Year",2023) | 931 | 931 | 619 | 931 | 253 |
| row and column items | =GETPIVOTDATA("Sum of Units",Pivot!$A$3,"Region","North","Year",2023) | 242 | 242 | 242 | 242 | 78 |
| item hidden by gpd-hide-item (West) | =GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","West") | 61499.159999999989 | 61499.159999999989 | error #REF! | 61499.159999999989 | 17009.98 |
| detail under the item gpd-collapse-item collapses (East/Bananas) | =GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","East","Product","Bananas") | 17787.900000000001 | 17787.900000000001 | 17787.900000000001 | error #REF! | 17787.900000000001 |
| item gpd-filter-out filters out (South/Cherries) | =GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","South","Product","Cherries") | 27173.670000000002 | 27173.670000000002 | 27173.670000000002 | 27173.670000000002 | error #REF! |
| inner item alone, no subtotal line (Cherries) | =GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Product","Cherries") | error #REF! | error #REF! | error #REF! | error #REF! | error #REF! |
| item a Product filter leaves without records (Central) | =GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","Central") | 926 | 926 | 926 | 926 | error #REF! |
| cell reference into the pivot | =Pivot!B6 | 11204.24 | 11204.24 | 11204.24 | 11204.24 | 4217.1099999999997 |
| range over pivot cells | =SUM(Pivot!B6:B8) | 16655.849999999999 | 16655.849999999999 | 16655.849999999999 | 20019.480000000003 | 10847.9 |

### (e) refreshOnLoad seeds and cases (§7.7)

| file | location | refreshOnLoad | saveData | recordCount | records part |
| --- | --- | --- | --- | --- | --- |
| refresh-seed (input) | A3:I24 | 1 | — | 56 | yes |
| refresh-seed-nodata (input) | A3:I24 | 1 | 0 | 56 | no |
| refresh-on-load | A3:I24 | 1 | — | 56 | yes |
| refresh-on-load-collapse | A3:I21 | 1 | — | 56 | yes |
| refresh-on-load-nodata | A3:I24 | 1 | 0 | 56 | no |

