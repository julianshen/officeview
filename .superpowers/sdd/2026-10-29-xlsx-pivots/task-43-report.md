# Task 43 report — Stage 5 verification subset

## Files

- `src/components/xlsx/InteractiveSheetDoc.tsx`
- `src/components/xlsx/pivot/PivotButtonLayer.tsx`
- `src/worker/viewerSession.ts`
- `src/xlsx/pivots/slicers/controller.ts`
- `src/xlsx/pivots/slicers/transaction.ts`
- `tests/xlsx-pivots-linked-controls.test.ts`
- `tests/xlsx-pivots-stage5-audit.test.ts` (new)
- `tests/xlsx-viewport-region-lines-subdetail.test.ts`
- `CHANGELOG.md`
- `docs/validation/xlsx-pivots.md`
- `.superpowers/sdd/2026-10-29-xlsx-pivots/task-43-audit.md`
- This report and the named completion marker (workflow evidence).

## RED / GREEN evidence

Base `416e5426`, clean `feature/pivots-s5`. No other branch/source files edited; no agents spawned.

1. Initial audit tests: **4 failed / 2 passed**. Cancel and close refunded a still-running producer; early commit accepted an incomplete set; denied slicer admission still read 10,000 ids. GREEN: these tests now retain 100 B until drain, reject early commit, and read zero ids on refusal. Abort also retains until drain and is covered. Logs `/tmp/s5-t43-red.log`, `/tmp/s5-t43-green.log`.
2. Extracted the existing mounted hidden-index comparison unchanged, then counter-test RED: **50,005,000 reads** vs a 10,000 bound. Set-based GREEN: **10,000 reads**. `/tmp/s5-t43-linear-red.log`, `/tmp/s5-t43-focused.log`.
3. Worker composite RED: a 512 B grant did not prevent a 10,000-rectangle construction. GREEN: length-only preflight refuses with **0 element reads**. Existing worker suite plus new audit suite: **148 passed**, `/tmp/s5-t43-composite-{red,green}.log`.
4. Added four-target/20-cycle model stress (400 B peak, 80 constructions, zero bytes after each cycle), and 100-year daily timeline bound (<100 periods for 600px).
5. Full-suite discovery: `xlsx-viewport-region-lines-subdetail` loaded the owner's ignored 165-byte `~$work-seed-step1.xlsx` as a ZIP. The same failure reproduced alone. The enumeration now excludes Excel owner-lock sidecars, without deleting them or weakening any real-workbook assertion. RED `/tmp/s5-t43-region-rerun.log`; GREEN **6 passed** `/tmp/s5-t43-region-green.log`.

## Gates

- `npx tsc --noEmit -p tsconfig.json`: exit 0, `/tmp/s5-t43-tsc-final.log`.
- Required `npx vitest run tests/xlsx-pivots* tests/xlsx-viewer-* tests/xlsx-formula*`: **230 files passed; 6,319 passed, 59 skipped, 1 todo**, exit 0. `/tmp/s5-t43-broad-stable.log`.
- `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts && bun scripts/eager-budget.ts`: exit 0; **35,977 B / 38,720 B**, 155 eager modules, paged-worker chunks valid. `/tmp/s5-t43-build-final.log`. Current lazy chunks: PivotButtonLayer 116.02 kB / 28.91 kB gzip; InteractiveSheetDoc 310.71 kB / 91.81 kB gzip. No clean-baseline lazy-byte delta was measured, so no claim of zero lazy growth.
- `IV_HARNESS_PORT=4193 npm run test:browser -- xlsx-pivots`: **10 passed**, Chromium and Playwright WebKit, `/tmp/s5-t43-browser.log`. Saved-paint/collapse parity only; not physical Safari or the full Stage 5 gesture matrix.
- `bun scripts/pivot-bench.ts`: **GATE PASS**, three runs, median-of-medians **138.68 ms <= 200 ms**, run medians 136.28/138.68/143.83 ms. One million records; shared-machine evidence. `/tmp/s5-t43-bench.log`.
- `npx vitest run` final rerun: **599 files passed; 12,127 passed, 61 skipped, 1 todo**, exit 0, 400.57 s. `/tmp/s5-t43-full-final.log`. The earlier full run had only the Excel lock-sidecar failure (12,126 passed); the final run includes its fix.
- `git diff --check`: passed.

Earlier runs launched during RED/GREEN work are not final validation: one cached the old composite helper; the original restricted run additionally hit `listen EPERM 127.0.0.1` in `xlsx-viewer-harness-identity.test.ts` and eight HTTP cases in `xlsx-viewer-bench-collector.test.ts`. The required suite was rerun with authorized localhost access on stable source.

## Concerns and scope

- This implements the subset in `s5-codex-rest.md`: audit/fixes, changelog, Stage 5 results and S5-A/B/C dispositions. It does not open a PR, push or merge. Independent review is still required.
- Browser smoke is not the full Task 43 gesture/focus/CLS matrix. No recipe was declared measured by executing synthetic tests. S5-C remains binding.
- Multi-target UI, table slicers and unmeasured cross-filter interactions remain refused. Lower-level fan-out tests do not imply four real mounted pivots share one records scan. The existing 100,000-row index sharing check is a helper test, not production scan-sharing evidence.
- Generated timeline locale recognition is limited to measured default pivot-caption forms, not general language detection.
- Per-control initial registration can revisit same-field control metadata; it is not certified at unbounded numbers of controls. Main control-state/map lifetime across document or sheet replacement and resetting live controls deserve the independent whole-feature review; these are not covered by the primitive ledger stress alone.
- Main-thread model bytes and deterministic counters are not physical heap peaks. Performance timing uses a shared machine, not quiet-machine acceptance. Actual Safari/VoiceOver/physical-device tests were not attempted.

## Commit

Implementation, tests, changelog, validation and audit: `63f8c22c909e2c97f91a4d456009c3767e482b1f`. All gates above passed before commit. This report is committed separately; its SHA is recorded in `s5-t43.done`.
