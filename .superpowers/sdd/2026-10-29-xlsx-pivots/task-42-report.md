# Task 42 report — timeline controls

## Files

- `src/components/xlsx/InteractiveSheetDoc.tsx`
- `src/components/xlsx/pivot/PivotButtonLayer.tsx`
- `src/components/xlsx/pivot/TimelineLayer.tsx`
- `src/xlsx/pivots/slicers/controller.ts`
- `src/xlsx/pivots/slicers/model.ts`
- `tests/xlsx-pivots-control-oracle.test.ts`
- `tests/xlsx-pivots-timeline-controls.test.tsx`

## RED / GREEN

- Added controller tests for UTC calendar boundaries, inclusive final dates, range-preserving level changes, commit-only selection changes, and admission refusal retaining prior state.
- Added layer tests for period hits and picker, S5-C drag/scroll refusal, keyboard anchor/extend/Escape, focus cycling, and timeline-only sheet mounting. The worker-backed test checks the linked-control transaction and resulting pivot cells.
- Reproduced and fixed a worker linkage failure: timeline selection must link using the timeline control name, not the linked pivot's name. The mounted paged test now passes.
- Reproduced and fixed keyboard focus staying on the old period after Shift+Arrow; focus now follows the active period. Focused test passed after the fix.
- Extended the control oracle assertions with exact cells for `s5-timeline-click-month` and `s5-timeline-level-preserves-range`.

## Gates

- `npx tsc --noEmit -p tsconfig.json` — passed.
- `npx vitest run tests/xlsx-pivots-timeline-controls.test.tsx tests/xlsx-pivots-control-oracle.test.ts` — 2 files passed; 28 passed, 8 skipped.
- `npx vitest run tests/xlsx-viewer-construction-memory.test.ts -t 'each modelled structure charges'` — 1 passed, 16 skipped when isolated.
- Required broad run, `npx vitest run tests/xlsx-pivots* tests/xlsx-viewer-* tests/xlsx-formula*` — nonzero: 227 files passed, 2 failed; 6,298 tests passed, 9 failed, 59 skipped, 1 todo; 8 errors. The run emitted sandbox `listen EPERM 127.0.0.1` errors from localhost-based viewer harness/collector tests. Those two test files passed when rerun alone (30 passed); this broad-run result is environment-sensitive and remains a gate limitation.
- `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts` — passed; eager graph ok, paged worker chunks ok, eager gzip 35,976 B / 38,720 B.
- `git diff --check` — passed.

## Concerns / follow-up

- S5-B remains an open point for review: generated timeline text currently comes from the existing English-only `timelineLayout` strings (`MONTHS`, `LEVEL_LABEL`, and selection text). This task does not pass detected workbook language into that layout or provide measured localized strings. The ruling says generated labels use detected workbook language and otherwise retain saved labels or refuse generation; that behavior needs a focused follow-up before claiming locale parity.
- The task sequence document `s5-codex-rest.md` labels Task 42 as linked pivots/table slicers, while the explicit `task-42-brief.md` is Timeline controls and Task 41's report records slicer controls. This implementation follows the explicit Task 42 brief; reconcile the sequence text separately.
- Unmeasured timeline scroll/drag remain disabled/refused under S5-C.

## Commit

- Source and test commit: `7aba58cb` (`feat(pivots): add interactive timeline controls`).
- Report commit: recorded in the completion marker after creation.
