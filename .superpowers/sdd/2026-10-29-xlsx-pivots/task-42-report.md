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

## Fix round 1

### Files

- `src/components/xlsx/InteractiveSheetDoc.tsx`
- `src/xlsx/pivots/slicers/controller.ts`
- `tests/xlsx-viewer-pivot-ui.test.tsx`

### Finding 1 — RESOLVED

- RED: the mounted painter probe showed the controller committed `selectedItemIds: [0]` while the saved drawing passed to `paintControl` still contained `[0, 1, 2, 3]` (all four items selected).
- GREEN: `SlicerControllerHost.onChange` publishes the committed model only after a successful apply. The viewer retains that model by slicer name, bumps the sheet's drawing generation, and substitutes it into the next prepared drawing before tile paint. The controller stays alive across that drawing refresh and accepts only strictly newer registered revisions, avoiding a stale saved drawing overwriting its committed state.
- Regression evidence: the existing worker/paged click-one-and-clear test remains green; its linked-control transaction still updates the visible pivot.
- Test: `tests/xlsx-viewer-pivot-ui.test.tsx` / “repaints the canvas control from the committed slicer selection”.

### Finding 2 — VERIFIED, OPEN (owner decision required)

- The report's previous English-label concern is accurate: `timelineLayout` generates English month, quarter, level and selection labels; `TimelineModel` has no language or saved period-label fields. The current XLSX parser assigns `semantics.locale: 'en-US'` as a fixed default, while the measured timeline reference includes zh-TW strings. `dc:language` and equivalent workbook-language metadata are not parsed into the timeline model, and the timeline XML stores no generated period labels to preserve.
- This cannot be safely closed as a code-only correction: choosing a detected-language source and the zh-TW translations requires reconciling S5-B against the existing English painter baseline and deciding what to show when no saved labels exist. I have not silently selected a fallback or claimed the requirement is met. Finding 2 remains open pending the owner ruling; this fix round is not complete until that is resolved.

### Gates

- `npx tsc --noEmit -p tsconfig.json` — passed.
- Focused slicer/control/UI suites — 3 files passed; 68 passed, 8 skipped.
- Required broad gate `npx vitest run tests/xlsx-pivots* tests/xlsx-viewer-* tests/xlsx-formula*` — 227 files passed, 2 failed; 6,299 passed, 9 failed, 59 skipped, 1 todo. The failures were the sandbox network-bind restrictions: `listen EPERM 127.0.0.1` in `xlsx-viewer-harness-identity.test.ts` and `xlsx-viewer-bench-collector.test.ts`; the latter's eight dependent setup/hook errors timed out after its listener could not bind.
- `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts` — passed; eager graph ok, paged worker chunks ok, gzip 35,977 B / 38,720 B.
- `git diff --check` — passed.

### Commit

- Finding 1 source/test fix: `9c76f43f` (`fix(pivots): repaint committed slicer selection`). Finding 2 is still open, so no completion marker is written yet.

## Fix round 1

### Files

- `src/xlsx/pivots/slicers/model.ts`
- `src/xlsx/pivots/slicers/parseTimelines.ts`
- `src/xlsx/pivots/slicers/paint.ts`
- `tests/xlsx-pivots-timeline-controls.test.tsx`

### Finding 1 — RESOLVED before this round

- Verified the existing fix from `9c76f43f` by rerunning `tests/xlsx-viewer-pivot-ui.test.tsx -t 'repaints the canvas control from the committed slicer selection'`: 1 passed. The regression test observes the committed selection in the canvas drawing model after the controller publishes it. No duplicate production change was needed.

### Finding 2 — RESOLVED

- Verified the finding against the current code: timeline generation had no language property and hardcoded English month, level, group, and selection labels.
- RED: added a parsed-fixture test for the zh-TW-linked `controls-seed.xlsx` plus an unknown/mixed linked-pivot fallback test. Before the fix, both failed: the model had no `labelLocale`, and no locale resolver existed. The fixture's saved scroll position starts at March, so the final expected first period is `3月`.
- GREEN: `timelineLabelLocale()` uses the saved linked pivot names, following the existing pivot-caption locale evidence: all default `樞紐分析表N` targets select the measured zh-TW strings; all default `PivotTableN` targets preserve the English table. Unknown or mixed names yield no model and a `timeline-label-locale` diagnostic, leaving the drawing's AlternateContent fallback in place rather than guessing generated labels.
- Added locale-aware period, group, level, and selection labels. Tests pin `3月`, `月`, `所有週期`, and the measured selection string `2024 3月 - 5月`; the saved-paint baseline and control-oracle suites remain green.

### Gates

- `npx tsc --noEmit -p tsconfig.json` — passed.
- Focused `timeline-controls`, `control-oracle`, and `slicers-display` suites — 3 files passed; 127 passed, 8 skipped.
- Slicer repaint regression probe — 1 passed.
- Required broad `npx vitest run tests/xlsx-pivots* tests/xlsx-viewer-* tests/xlsx-formula*` — 227 files passed, 2 failed; 6,301 passed, 9 failed, 59 skipped, 1 todo; 8 errors. Failures were confined to `tests/xlsx-viewer-harness-identity.test.ts` and `tests/xlsx-viewer-bench-collector.test.ts`: both hit sandbox `listen EPERM 127.0.0.1`; the collector's eight dependent HTTP cases timed out after the listener could not bind.
- `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts` — passed; eager graph ok, paged worker chunks ok, gzip 35,978 B / 38,720 B.
- `git diff --check` — passed.

### Commit

- Source and test fix: `52184a461957ae9d2073f92e256cded5a452bee3` (`fix(pivots): follow measured timeline label locale`).

## Fix round 1 — OpenCode findings verified at HEAD

### Files

- No source or test edits were needed in this round; both findings are already fixed in `9c76f43f` and `52184a46`.
- This report is the only file changed for this verification round.

### Finding 1 — RESOLVED

- Verified the original failure mode against current code: after a committed selection, the canvas drawing must use the controller's committed slicer model, not the package's saved all-selected model.
- Test-first evidence is in `9c76f43f`: the mounted regression test `repaints the canvas control from the committed slicer selection` observes the actual `paintControl` input after selecting Apples and asserts selected ids `[0]` and item states `[true, false, false, false]`; it also waits for the filtered pivot result. The test would fail if the painter continued receiving the saved drawing.
- Current implementation publishes only committed models, stores them by slicer name, increments the owning sheet's drawing generation, and replaces the drawing's slicer model in `prepareEntry`; the regression test passes at HEAD.

### Finding 2 — RESOLVED

- The S5-B behavior is implemented in `52184a46`: label generation receives a locale only when all linked pivot names identify one of the two measured default-caption locales (`樞紐分析表N` → `zh-TW`, `PivotTableN` → `en-US`). Unknown, custom, or mixed names return no locale, refuse to build the generated timeline model, emit `timeline-label-locale`, and preserve the saved AlternateContent fallback. This deliberately fails closed rather than substituting the host locale or guessing.
- Tests exercise the measured zh-TW strings from the controls-seed workbook and the unknown/mixed-name refusal path. `timelineLayout` uses the model locale for period, group, level, all-period, and selection labels; no English generator is used for a model detected as zh-TW.

### Gates

- `npx tsc --noEmit -p tsconfig.json` — passed.
- Focused suites (`tests/xlsx-pivots-timeline-controls.test.tsx`, `tests/xlsx-pivots-control-oracle.test.ts`, `tests/xlsx-viewer-pivot-ui.test.tsx`) — 3 files passed; 70 passed, 8 skipped.
- Required broad gate — 227 files passed, 2 failed; 6,301 passed, 9 failed, 59 skipped, 1 todo; 8 errors. All failures/errors were the sandbox localhost bind restriction (`listen EPERM 127.0.0.1`) in `tests/xlsx-viewer-harness-identity.test.ts` and `tests/xlsx-viewer-bench-collector.test.ts` (the collector's eight HTTP tests/hooks depend on its blocked listener).
- `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts` — passed; eager graph and paged chunks validated; gzip 35,978 B / 38,720 B.
- `git diff --check` — passed.

### Commit

- Report-only verification: pending.

## Fix round 1 — OpenCode review findings, independent verification

### Files

- No source or test edits were needed in this verification round. The findings were already addressed in `9c76f43f` and `52184a46`; this section records the independent check at `0ac0210c`.
- This report is the only file changed for this round.

### Finding 1 — RESOLVED (committed slicer selection reaches canvas paint)

- Verified the failure mode: canvas painting consumes the drawing's slicer model, so a controller-only committed selection would leave the painted control at its saved selection.
- The test-first fix is in `9c76f43f`: `SlicerControllerHost.onChange` publishes only after successful apply; `InteractiveSheetDoc` retains the model by slicer name, bumps drawing generation, and substitutes the committed model for the next tile paint. The mounted regression test `repaints the canvas control from the committed slicer selection` asserts that selecting Apples changes the painter's actual model to selected ids `[0]` and item states `[true, false, false, false]`, while the filtered pivot result updates.
- Independent GREEN: `npx vitest run tests/xlsx-viewer-pivot-ui.test.tsx -t 'repaints the canvas control from the committed slicer selection'` — 1 passed. The broader focused run also passed.

### Finding 2 — RESOLVED (S5-B language rule)

- The reported gap was valid: timeline labels had English-only generators and no fail-closed locale decision.
- The test-first fix is in `52184a46`: generated labels are enabled only when all linked pivot captions identify a measured locale (`樞紐分析表N` → `zh-TW`, `PivotTableN` → `en-US`). Unknown, custom, or mixed captions produce `timeline-label-locale` and preserve the saved fallback rather than generating unmeasured labels. Tests cover the measured zh-TW labels and unknown/mixed refusal; the saved-paint baseline remains intact.
- Independent GREEN: `npx vitest run tests/xlsx-pivots-timeline-controls.test.tsx tests/xlsx-pivots-control-oracle.test.ts tests/xlsx-slicers-display.test.ts` passed as part of the four-file focused run below.

### Gates

- `npx tsc --noEmit -p tsconfig.json` — passed.
- Focused: `npx vitest run tests/xlsx-viewer-pivot-ui.test.tsx tests/xlsx-pivots-timeline-controls.test.tsx tests/xlsx-pivots-control-oracle.test.ts tests/xlsx-slicers-display.test.ts` — 4 files passed; 167 passed, 8 skipped.
- Required broad gate: `npx vitest run tests/xlsx-pivots* tests/xlsx-viewer-* tests/xlsx-formula*` — 226 files passed, 3 failed; 6,300 passed, 10 failed, 59 skipped, 1 todo; 8 unhandled errors. `tests/xlsx-viewer-harness-identity.test.ts` and eight dependent cases in `tests/xlsx-viewer-bench-collector.test.ts` failed on sandbox `listen EPERM 127.0.0.1`. `tests/xlsx-viewer-construction-memory.test.ts`'s heap-bound calibration failed once in the broad run (`builder-numeric x98304`: 9,437,184 B vs 6,282,246 B limit), then passed when rerun alone (1 passed, 16 skipped); no code in that test's accounting path changed in this round.
- `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts` — passed; eager graph and paged chunks validated; gzip 35,978 B / 38,720 B.
- `git diff --check` — passed before this report-only edit.

### Commit

- Finding 1 source/test fix: `9c76f43f` (`fix(pivots): repaint committed slicer selection`).
- Finding 2 source/test fix: `52184a46` (`fix(pivots): follow measured timeline label locale`).
- This round's report commit: recorded in the completion marker.

## Fix round 1 — OpenCode findings re-verified

### Files

- No source or test changes were needed: both findings are resolved by `9c76f43f` and `52184a46`.
- This section is the only change in this round.

### Finding 1 — RESOLVED

- Rechecked the mounted paint path: after a successful slicer transaction, `SlicerControllerHost.onChange` publishes the committed model; the viewer stores it by slicer name, increments drawing generation, and substitutes it for the saved drawing model before repaint.
- Re-ran `npx vitest run tests/xlsx-viewer-pivot-ui.test.tsx -t 'repaints the canvas control from the committed slicer selection'`: 1 passed, 39 skipped. The test asserts the painter receives selected ids `[0]` and item states `[true, false, false, false]` after Apples is selected, while the pivot result updates. This directly covers the stale saved-all-selected model failure.

### Finding 2 — RESOLVED

- Rechecked S5-B: only linked pivot names with measured default-caption forms produce a timeline locale (`樞紐分析表N` -> `zh-TW`, `PivotTableN` -> `en-US`). Unknown, custom, or mixed names refuse timeline model generation and retain the saved fallback. Locale-aware period, group, level, selection, and all-period strings are generated only for those measured locales.
- Re-ran `npx vitest run tests/xlsx-pivots-timeline-controls.test.tsx tests/xlsx-pivots-control-oracle.test.ts tests/xlsx-slicers-display.test.ts`: 3 files passed, 127 passed, 8 skipped. Tests pin the measured zh-TW labels and unknown/mixed-language refusal; the saved-paint baseline remains green.

### Gates

- `npx tsc --noEmit -p tsconfig.json` — passed.
- Required broad gate `npx vitest run tests/xlsx-pivots* tests/xlsx-viewer-* tests/xlsx-formula*` — 227 files passed, 2 failed; 6,301 passed, 9 failed, 59 skipped, 1 todo; 8 unhandled errors. The only failures were `tests/xlsx-viewer-harness-identity.test.ts` and 8 dependent cases in `tests/xlsx-viewer-bench-collector.test.ts`, all from sandbox `listen EPERM 127.0.0.1` (the HTTP hook timed out because its listener could not bind).
- `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts` — passed; eager graph and paged worker chunks validated; gzip 35,978 B / 38,720 B.
- `git diff --check` — passed before commit.

### Commit

- Report-only verification commit: `6dff036a6efd6eefee7fab8eabc965b7e039155e` (`docs(pivots): verify Task 42 review fixes`).
