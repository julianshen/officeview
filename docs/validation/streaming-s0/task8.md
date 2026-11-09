# Task 8 — keyed messages, T7/T9 measurement

Run: `bun run s0:napi scripts/stream-s0/t7-t9-battery.ts`.
Results: `t7.json`, `t9.json`; progress: `task8-run.log`.
This is an uncommitted Stage-0 prototype, not a production attribution guarantee.

## Implementation and RED/GREEN

Files: `scripts/stream-s0/mark/message.ts`, `scripts/stream-s0/mark/detect.ts`,
`scripts/stream-s0/t7-t9-battery.ts`, `tests/stream-s0/battery.test.ts`, and these measurement artifacts.
No `src/`, harness, census, or Task 2 files changed. No installs or commits.

The message encoding is uint32 big-endian keyGen followed by the five MSB-first
bytes token24/salt8/CRC8. The first 40 HMAC-SHA256 bits enter the existing
`extendedCodeword` function, which already invokes BCH encoding internally.
Audit insertion checks the generation-scoped used-message set and redraws salt;
256 unsuccessful draws fail closed, including a stuck random generator.
Candidate allocation uses that insertion path. Keys/seeds are public, fixed
reproducibility inputs, not production secret material.

Four Part A tests failed before implementation: keyed known vector, collision
redraw, multi-hit refusal, and K5 innocent framing. Then two battery tests failed
for missing calibration and row coverage. Final focused run: 3 files passed,
26 tests passed, 1 existing opt-in benchmark skipped. The battery smoke uses two
labelled synthetic pages and M=256; the full artifacts use only real corpus
units and M=65536. No synthetic page enters the full measurement.

The framing probe uses colluders 1/2/4/8/16 and innocent XOR token 31. Linear-map
innocent Z=9.396725; keyed-map innocent Z=-0.064316. Keyed colluder scores are
7.706–9.535. Multi-hit decisions return indices and scores for review, never one
viewer. This one regression is not a proof against all collusion.

## Method and limits

- All 24 parameter settings and all 66 available units are measured. Selection
  maximizes one-token-per-page q70→half survival at the full-M union floor,
  ties broken by lower maximum ΔE. Final five-token survival uses the subsequently
  calibrated thresholds. This is a pilot selection, not independent held-out
  optimization of every setting under its own calibration.
- The null loop varies the 32-byte basis key on every attempt, cycles real pages
  and all acquisition attacks, and includes unmarked, wrong-key marked, and
  same-basis/key out-of-table payloads. Every score call includes registration.
  Flat/failed-registration decisions are retained; extra attempts reach 2000
  nonflat full-M decisions (131072000 candidate scores).
- Calibration fits upper-decile Gaussian QQ lines for Z_best and the best-minus-
  second margin, both pooled and separately by null class. Thresholds take the
  largest class fit and the whole-M sub-Gaussian union floor. The JSON records
  empirical CDF points, fits and residuals. The union formula is a cross-check
  contingent on the score-tail model, not a theorem about structured BCH or
  registration-selected residuals. Zero false results in a few thousand trials
  cannot establish a 1e-6 empirical confidence bound; the 95% binomial upper
  bound is separately published. No Layer-2/consensus calibration is claimed.
- Survival retains insufficient-carrier pages in the denominator as failures.
  Crops are centered and remove 5% or 25% of each axis, following the existing
  attack helper; the composite rows crop before scaling.
- Collusion has one independently page-keyed cohort per unit, using six distinct
  table tokens; the appropriate prefix supplies K. Average K=2/4/6, median
  K=3/4/5/6, and encoded two-stream difference are reported separately. Counts
  distinguish colluder attribution, multi-hit suspicion, innocent attribution,
  and none/abstain. A suspect set may include innocent candidates: it is not an
  attribution and is deliberately left for manual review.
- Corpus coverage is 12 DOCX, 11 XLSX, 43 PPTX: short by 3 DOCX units. The five
  image-heavy-unit minimum is not independently classified/certified. This
  coverage gate is RED, not padded with synthetic pages.

TypeScript was run twice. The saved `task8-tsc.log` reports errors only in
protected harness/census/Task 2 files; no errors name this task's mapper,
detector, battery or tests. Those unrelated files were left untouched.

## Completed full-run results

Backend: `@napi-rs/canvas`; result-write SHA `709a8adfdf1f773ae54151d80d595648491063d2`; uncommitted source hashes in `task8-source.sha256`. Generated 2026-11-08T10:09:31.720Z. Run exited 0; RED scientific gates are results, not suppressed execution errors.

Chosen parameters: `{"nChips": 512, "cellPx": 4, "tilePx": 64, "amplitude": 2, "pedestal": 2}`. Pilot q70→half: 60/66; page maximum ΔE00 1.590280036. **T7 GREEN**. Every setting's per-page mean/p95/max, chip-pixel p95, carrier counts and pedestal variant are in t7.json.

### Calibration

2522 attempts; 2000 nonflat full-M decisions; 131,072,000 actual candidate scores. All use the same prepared 33,554,432-byte matrix and a distinct basis key per attempt. Classes: {'unmarked': 841, 'wrong-key': 841, 'out-of-table': 840}.

Chosen τ_abs=7.935362493, τ_margin=3.785665674. Pooled upper-tail fit controls both; union floor at M=65536 is 7.057742620. M=2²⁴ union floor=7.803978843; conservatively union-scaled fitted threshold=8.605831325.

| Fit population | Fitted abs threshold | Fitted margin threshold | Abs QQ RMS |
|---|---:|---:|---:|
| all | 7.935362 | 3.785666 | 0.166473 |
| unmarked | 5.995601 | 1.655519 | 0.020580 |
| wrong-key | 5.746855 | 1.384627 | 0.023301 |
| out-of-table | 6.589234 | 3.374926 | 0.082477 |

Empirical false attributions: **0**; null multi-hits: 0. Model target 1e-6 and union-floor checks: **GREEN**. The observed-zero check is GREEN; empirical proof of 1e-6 is **NOT ESTABLISHED**. A zero/2000 result has a one-sided 95% upper bound 0.00149674, not 1e-6. The JSON's attempt-based bound is 0.00118713; 228 attempts failed carrier capacity, so the stricter 2000-scored bound is the appropriate conservative headline. The other zero-score attempts are retained as non-attributions, not candidate-score evidence.

### Survival (330 captures per row)

| Attack | Correct | Rate | Target | Gate |
|---|---:|---:|---:|---|
| default | 300/330 | 90.91% | 90% | GREEN |
| q95 | 300/330 | 90.91% | reported only | REPORTED |
| q90 | 300/330 | 90.91% | reported only | REPORTED |
| q80 | 300/330 | 90.91% | reported only | REPORTED |
| q70 | 300/330 | 90.91% | 75% | GREEN |
| up2 | 300/330 | 90.91% | reported only | REPORTED |
| down0.5 | 300/330 | 90.91% | 75% | GREEN |
| q70-down0.5 | 300/330 | 90.91% | 75% | GREEN |
| crop5 | 225/330 | 68.18% | reported only | REPORTED |
| crop25 | 200/330 | 60.61% | 70% | RED |
| crop25-down0.5 | 205/330 | 62.12% | reported only | REPORTED |
| crop25-up2 | 205/330 | 62.12% | reported only | REPORTED |

**T9 RED:** 25% crop is below 70%. No threshold, target, denominator or test tolerance was changed to turn this green. Six small XLSX units cannot hold 512 carriers; all 30 of their five-token captures remain failures in each row. For the default row, DOCX=60/60, PPTX=215/215, XLSX=25/55. All row-by-class counts are in t9.json.

### Collusion (66 trials per attack)

| Attack | Colluder attributed | Collusion suspected | Innocent attributed | None/abstain |
|---|---:|---:|---:|---:|
| average-2 | 0 (0.00%) | 60 (90.91%) | 0 (0.00%) | 6 |
| average-4 | 0 (0.00%) | 60 (90.91%) | 0 (0.00%) | 6 |
| average-6 | 0 (0.00%) | 60 (90.91%) | 0 (0.00%) | 6 |
| median-3 | 0 (0.00%) | 60 (90.91%) | 0 (0.00%) | 6 |
| median-4 | 0 (0.00%) | 60 (90.91%) | 0 (0.00%) | 6 |
| median-5 | 0 (0.00%) | 58 (87.88%) | 0 (0.00%) | 8 |
| median-6 | 0 (0.00%) | 59 (89.39%) | 0 (0.00%) | 7 |
| difference-2 | 23 (34.85%) | 0 (0.00%) | 0 (0.00%) | 43 |

Collusion is **REPORTED**, not an attack-resistance pass claim. Zero observed innocent attributions on every row; multi-hit outputs name no single viewer. Two-stream differencing still attributes one of its colluders in 23/66 trials.

### Gates and deviations

- GREEN: keyed-message known-vector/generation tests; collision redraw/exhaustion; K5 framing regression; multi-hit refusal; 24-setting ΔE max sweep; ≥2000 full-M null decisions and zero false attribution; default/q70/half/q70→half survival targets; focused tests; diff whitespace check.
- RED: 25% crop survival (60.61% versus 70%); corpus coverage (12 DOCX versus 15, image-heavy minimum unverified). Extra crop→half/crop→2 rows are both 62.12%, reported with no separate numerical target.
- BLOCKED/RED repository TypeScript: protected census/harness/concurrent Task 2 errors, saved in task8-tsc.log. No errors in this task’s files; no unrelated fixes attempted.
- Corpus load failure: poi-crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx is a corrupt ZIP. The 66 loaded units are all included; one additional PPTX yielded zero units. No synthetic padding.
- Image-heavy metadata spot-check: only two PPTX slide parts had a direct picture rectangle at least half the slide area (lo-bnc591147 and lo-tdf156808). That is not rendered coverage verification and is not counted as satisfying five image-heavy units.
- Null class/page order is deterministic: pages cycle modulo 66 and classes modulo 3, giving each class a different 22-page subset; keys change per attempt and acquisition attacks cycle. This run is not a full Cartesian page×class experiment.
- Stage-0 Layer 1 only; no production secret, Layer-2 consensus, independent holdout calibration, physical capture or cryptographic robustness claim.

Focused command: `npx vitest run --config vitest.stream-s0-napi.config.ts tests/stream-s0/battery.test.ts tests/stream-s0/mark-embed.test.ts tests/stream-s0/mark-detect.test.ts` → **3 files passed; 26 passed, 1 existing opt-in benchmark skipped**. `git diff --check` passed; all four source hashes rechecked OK. No installs, source-tree edits or commits.

Concurrent-work provenance: measurement started at 8c0e6bfd; pi committed Task 2 as 709a8adf during the run. The harness writer records HEAD at result-write time. The four Task 8 source hashes stayed unchanged throughout; the saved TypeScript errors describe the earlier shared-tree checks, not a fresh check of pi’s final commit.
