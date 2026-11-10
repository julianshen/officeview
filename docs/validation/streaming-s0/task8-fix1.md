# Task 8 fix round 1 — crop registration

Implementation is confined to Stage-0 prototype scripts/tests/results; no product `src/`, installs, commits, or protected harness/census edits.

## Files

- `scripts/stream-s0/mark/register.ts`: exhaustive coarse translation grid, top-64 block-average pyramid refinement, full-resolution translation minimum certification, and ambiguous-content synchronization using only page-keyed chip support.
- `scripts/stream-s0/mark/detect.ts`: passes page support to registration, without passing candidate symbols or scores.
- `scripts/stream-s0/t7-t9-battery.ts`: q30/q10 (reported, no target); per-survival-capture offset/support/score/cause records; full existing null path automatically includes registration.
- `tests/stream-s0/registration.test.ts`, `tests/stream-s0/battery.test.ts`: exact reviewer DOCX, flat crop at three scales, non-centred crops, capture fields and attack rows.
- `scripts/stream-s0/t5-showplan.ts`, `docs/validation/streaming-s0/t5.json`: gate uses import-graph test plus static tolerance; coverage gaps stay explicit and unproved.
- Measurement outputs: regenerated t7/t9 JSON, this report, fix-round logs and source hashes.

## Registration and evidence

The original reviewer probe was recovered from `/private/tmp/claude-501/t8rev/probe2.ts`. At `poi-55733.docx#0`, original code returns (33,100), full-resolution MSE 5.435, versus true (102,132), MSE 0.069. Its old sampled error was 0.068: a misleading point-sampled match. Original-code replay is in `task8-fix1-red.log`.

A coarse-to-fine beam alone still missed this sparse page. The implementation now certifies the full-resolution integer translation minimum: rich content uses nonnegative partial squared-error and integral-energy Cauchy-Schwarz lower bounds to reject poses; blank windows are evaluated exactly in O(1); sparse content uses the exact SSD identity about a constant background. The sparse page has only 188 non-background reference pixels and several equally good blank crop regions. The test therefore verifies both minimum content error and final exact offset after page-support synchronization. A centred candidate provides an initial upper bound, not the search result: all translations are considered by the certificate, and non-centred top-left/bottom-right tests pass.

Ambiguity is explicit: at least two evaluated content poses separated by >=2 capture pixels have errors within max(0.1 luminance-squared, 5% of the best error). Such cases scan page-keyed chip-support residual energy density on a half-cell grid, then refine the top eight poses to integer pixels. A support-selected pose must remain within that content-error tolerance. Candidate codewords and candidate Z are not inputs. Constant reference planes are inherently content-ambiguous. Existing crop-before/after-rescale hypotheses remain supported; this is not a rotation/arbitrary-projective registration claim.

RED/GREEN: original flat slide selected (200,130) rather than (160,90); fixed flat registration passes at 1x, 0.5x and 2x. The exact DOCX test also failed before the fix and passes with it. Missing per-capture diagnostics failed the battery test before their implementation. Final focused run: 3 files, 20 tests passed. The latest full Stage-0 run passed 12 files / 77 tests, with 3 skips. TypeScript was green after the harness fixes, then the concurrent report task introduced a separate `scripts/stream-s0/report.ts:119` ImportMeta.dir error. Final gate status is recorded below; no protected file was edited by this task.

The null calibration includes the identical registration function, including support-driven pose selection. Recalibration is mandatory; old thresholds are not reused. Crops still remove 25% of each axis. JPEG q30/q10 have no newly invented pass threshold.

## T5 ruling

The retained capture measurements have 14 static slides, all zero pixel difference (tolerance 0.1%), and the executable client import-graph test passes. T5 status is GREEN. `coverageGaps` lists the interactive trigger in `lo-bnc591147.pptx` and the unmeasured `lo-narration.pptx` step. Counts remain 14/58 proof subset, 44 unproved, 57 measured steps; no coverage gap was promoted to proof. Captures were not rerendered just to change status. The driver now runs the graph test before emitting its gate.

## Concurrent checkout incident

During an earlier battery attempt the tracked fix edits disappeared while another lane committed e44705c3. That run was stopped. Only this task's edits were restored, a hashed source backup was saved, focused tests rerun, and the full battery restarted. Other-lane Task 2/census changes were preserved. Results below must match the final source hashes, not the abandoned runs.

A later full run exposed a pathological low-contrast narration rectangle for token 23 at crop25-up2. The exact verifier now orders its early-rejection probes by local gradient rather than distance from the global mean, so it tests edges before flat rectangle interiors. Pyramid ranking samples at most roughly 4096 blocks per pose; it is only a seed, and the full-resolution certificate remains authoritative. The regression returns the exact (160,90) origin. Both changes preserve the squared-error objective; the full sweep/null/survival battery was restarted to bind final results to this source. Exact verification retains O(number of poses × capture pixels) worst-case cost; this offline prototype makes no real-time detector claim.

## Final source gates

- Focused registration/battery/T5 tests: 3 files, 20 passed.
- Full Stage-0 tests: 12 files, 77 passed, 3 skipped (80 total).
- `npx tsc --noEmit -p tsconfig.json`: PASS on the final rerun. The concurrent report task fixed its ImportMeta.dir error; this task did not edit that file.
- `git diff --check`: PASS. Source checksums are recorded in `task8-fix1-source.sha256` and were rechecked during the measurement run.
- The offline exhaustive registration is slow on some crop/upscale combinations, notably `poi-crop-to-0.pptx#2`; individual captures take several minutes. No real-time latency claim is made.

## Completed measurement run

T7 **GREEN**; T9 **GREEN**. Parameters: `{'nChips': 512, 'cellPx': 4, 'tilePx': 64, 'amplitude': 2, 'pedestal': 2}`.
Full-M null decisions 2000; attempts 2561; candidate scores 131072000; false attribution 0; suspected 0.
New tau: abs=7.991253712887693, margin=3.8492227566189174; extrapolated tau24=8.657395150546177.
Zero-false 95% upper bound over attempted decisions: 0.0011690670933816039. This is not empirical proof of 1e-6; fitted-tail and sub-Gaussian assumptions remain explicit.

| Attack | Correct / total | Rate | Previous rate | Gate |
|---|---:|---:|---:|---|
| default | 300/330 | 90.91% | 90.91% | GREEN |
| q95 | 300/330 | 90.91% | 90.91% | REPORTED |
| q90 | 300/330 | 90.91% | 90.91% | REPORTED |
| q80 | 300/330 | 90.91% | 90.91% | REPORTED |
| q70 | 300/330 | 90.91% | 90.91% | GREEN |
| q30 | 262/330 | 79.39% | not measured | REPORTED |
| q10 | 5/330 | 1.52% | not measured | REPORTED |
| up2 | 300/330 | 90.91% | 90.91% | REPORTED |
| down0.5 | 300/330 | 90.91% | 90.91% | GREEN |
| q70-down0.5 | 300/330 | 90.91% | 90.91% | GREEN |
| crop5 | 300/330 | 90.91% | 68.18% | REPORTED |
| crop25 | 300/330 | 90.91% | 60.61% | GREEN |
| crop25-down0.5 | 300/330 | 90.91% | 62.12% | REPORTED |
| crop25-up2 | 300/330 | 90.91% | 62.12% | REPORTED |

Failure split (diagnostic, does not change gates; offset mismatch >0.51 original pixels):

| Attack | Failure causes |
|---|---|
| default | capacity: 30 |
| q95 | capacity: 30 |
| q90 | capacity: 30 |
| q80 | capacity: 30 |
| q70 | capacity: 30 |
| q30 | score-or-margin: 30, flat-residual: 8, capacity: 30 |
| q10 | score-or-margin: 250, flat-residual: 45, capacity: 30 |
| up2 | capacity: 30 |
| down0.5 | capacity: 30 |
| q70-down0.5 | capacity: 30 |
| crop5 | capacity: 30 |
| crop25 | capacity: 30 |
| crop25-down0.5 | capacity: 30 |
| crop25-up2 | capacity: 30 |

Every survival capture carries page/token/attack, registered and true offsets, support, zTrue/zBest, decision and diagnostic cause in t9.json. Capacity failures remain in denominators.

Coverage remains {'units': 66, 'counts': {'docx': 12, 'xlsx': 11, 'pptx': 43}, 'imageHeavyVerified': 0, 'shortfall': {'units': 0, 'docx': 3, 'xlsx': 0, 'pptx': 0, 'imageHeavy': 5}, 'note': 'All corpus units retained, including insufficient-carrier failures. No synthetic padding. Image-heavy unit classification has not been independently verified; minimum is NOT established.'}; no synthetic padding or inferred image-heavy certification.

Collusion:
- average-2: 0 colluder, 60 suspected, 0 innocent, 6 none, of 66.
- average-4: 0 colluder, 60 suspected, 0 innocent, 6 none, of 66.
- average-6: 0 colluder, 60 suspected, 0 innocent, 6 none, of 66.
- median-3: 0 colluder, 60 suspected, 0 innocent, 6 none, of 66.
- median-4: 0 colluder, 60 suspected, 0 innocent, 6 none, of 66.
- median-5: 0 colluder, 58 suspected, 0 innocent, 8 none, of 66.
- median-6: 0 colluder, 59 suspected, 0 innocent, 7 none, of 66.
- difference-2: 23 colluder, 0 suspected, 0 innocent, 43 none, of 66.

Final artifact checks: 4,620 survival capture records (66 units × 5 tokens × 14 attacks), 2,561 null attempt records with 2,000 full-M scored decisions. Post-regeneration focused suite passed all 20 tests. All requested edits remain uncommitted; no product source or protected harness/census file was changed.
