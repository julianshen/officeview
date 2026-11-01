# PLAN.md: Task 36 Value Field Settings Implementation

## Phase 1: Engine & Delta Implementation (`tests/xlsx-pivots-value-settings.test.ts`)
- [x] Test 1: valueSettings applies all 11 summary functions and all 15 showAs modes
- [x] Test 2: valueSettings caption rules: auto zh-TW, duplicate suffix, showAs preservation and fail-closed locale
- [x] Test 3: valueSettings rejects blank or duplicate custom names and out-of-scope custom numFmtIds with pivot-delta-invalid
- [x] Test 4: variance functions succeed without totals and fail closed with announcement when totals exist
- [x] Test 5: chart-linked pivot refuses non-zero numFmtId cleanly
- [x] Test 6: GETPIVOTDATA resolves renamed data field after valueSettings commit

## Phase 2: Oracle & Worker Parity (`tests/xlsx-pivots-interaction-oracle.test.ts`, `tests/xlsx-pivots-worker.test.ts`)
- [x] Test 7: Task 36 manifest cases pass on resident path
- [x] Test 8: Reclassified Stage 2 values-fn and values-showas cases pass on resident path
- [x] Test 9: Task 36 cases pass on worker/paged path with paged parity

## Phase 3: Dialog & Viewer UI (`src/components/xlsx/pivot/PivotValueSettingsDialog.tsx`, `tests/xlsx-viewer-pivot-value-settings.test.tsx`)
- [x] Test 10: opens from Values chip in field list and from value header cell, traps Tab and restores focus on Escape
- [x] Test 11: switching showAs to % Difference From enables Base field and Base item with (previous)/(next) and field items
- [x] Test 12: calculated data field restricts summary function and showAs options
- [x] Test 13: number format preset updates formatted value cells on resident and paged paths
- [x] Test 14: refused apply keeps dialog open with inline role="alert" and leaves committed cells unchanged
