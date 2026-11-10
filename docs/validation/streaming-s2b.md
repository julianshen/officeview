# Server streaming — Stage 2b validation (trace tool on served output; the Stage-2 gate)

Spec: `docs/superpowers/specs/2026-11-10-streaming-s2b-detection-design.md` (G4, G5, Capture geometry). Parent: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` §13.2–13.5. Stage-0 method: `docs/validation/streaming.md` (T9). Stage 2a: `docs/validation/streaming-s2.md`. Raw data: `streaming-s2b.json` (every trace's decision, residuals, band and timing; every null).

Run (code at 6e2129af plus `scripts/stream-s2b/calibrate.ts` as committed with this file; register.ts includes 7a's fix 12cac998), status **GREEN**:

```text
bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s2b/calibrate.ts \
  --nulls 2000 --viewer-every 4 --driver 4 --out-dir <workdir> \
  --write docs/validation/streaming-s2b.json --write-tau
```

## What it measures

- **Captures are served bytes.** For every corpus unit at bucket 1 (pptx also bucket 2) — 141 entries — an in-process stream server (`createStreamServer` through the test harness, real `AuditTable` over a fresh `AuditStore`) serves the unit over the wire to six marking sessions (fixed visible mark). The received `UNIT_RASTER` PNG is decoded by the tool's own `decodeImage`.
- **The population.** The store is padded through the server's own `AuditTable` (`issue` → `record` → `recordDelivery`) with issued rows and `l1` deliveries carrying the true canonical `refSha256`, until the (doc, unit) has exactly **65,536 candidates**. Every trace asserts the decision list has 65,536 scores (τ row 2¹⁶).
- **Every decision is `trace()` itself.** Its diagnostic `observe` hook (added in this task, no behaviour change) hands the score list, owners, registration residuals and ambiguity band to the script.
- **One process per document.** Each document runs in its own child process (fresh pool, fresh trace host), 4 at a time. See "Reference determinism" below for why.
- **τ refit, then re-decided.** Null decisions are fit with Stage 0's `calibrate()`. The batteries ran with the τ shipped at run time (Stage-0 7.99 / 3.85). They were then re-decided under the fitted τ by the same §13.2 rule (`decideScores`) over the recorded list. Every gate below uses the fitted τ.
  - Check, after `tau.ts` was rewritten: a native rerun over 4 documents agreed with the re-decided outcome on every survival, zoom, geometry and collusion trace. Numbers are in "Native check".

Entries: 141 planned. 135 marked. 6 shipped unmarked (capacity; listed in the JSON, same 6 as Stage 2a). 6 documents were refused at admission (§6.4 fail-closed strip, same as Stage 2a). Survival rates are over the 135 marked units (the §13.5 definition).

## Gates

| gate | value | status |
|---|---|---|
| survival default ≥ 90 % | 135/135 = 100.0 % | **GREEN** |
| survival q70 ≥ 75 % | 134/135 = 99.3 % | **GREEN** |
| survival down 0.5 ≥ 75 % | 135/135 = 100.0 % | **GREEN** |
| survival q70 → down 0.5 ≥ 75 % | 134/135 = 99.3 % | **GREEN** |
| survival crop 25 % ≥ 70 % | 135/135 = 100.0 % | **GREEN** |
| null decisions ≥ 2,000, fully scored | 2,160 (2,308 attempts) | **GREEN** |
| null false attributions = 0 | 0 at the fitted τ, 0 at the native τ; 0 collusion answers | **GREEN** |
| fitted tail ≤ 10⁻⁶ per decision | 9.60 × 10⁻⁷ (worst class and tail; union bound 7.2 × 10⁻¹¹) | **GREEN** |
| collusion average K = 2, 4, 6: innocent named = 0 | 0, 0, 0 | **GREEN** |
| collusion median K = 3, 4, 5, 6: innocent named = 0 | 0, 0, 0, 0 | **GREEN** |
| collusion two-stream differencing: innocent named = 0 | 0 | **GREEN** |
| geometry rotate 2°: 0 attributions | 0 | **GREEN** |
| geometry rotate 5°: 0 attributions | 0 | **GREEN** |
| geometry keystone 5 %: 0 attributions | 0 | **GREEN** |
| zoom and viewer batteries: innocent named = 0 (anti-goal check, not a §-gate) | 0 | **GREEN** |

"Innocent named" counts both kinds of naming:
- a `traced` answer naming a non-colluder;
- a `collusion` answer whose list contains a non-colluder.

## τ (shipped in `server/src/detect/tau.ts`, `source: 'stage-2b calibrate 6e2129af'`)

| M | τ_abs | τ_margin | how |
|---|---|---|---|
| 2¹⁶ | **7.93** | **3.45** | fitted: max over the pooled and per-class fits, rounded up to 0.01 |
| 2²⁰ | 8.27 | 3.45 | Stage-0 union rule, √(τ² + 2 ln(M / 2¹⁶)) |
| 2²⁴ | 8.60 | 3.45 | same rule |

For comparison, Stage 0 (provisional, shipped before this run): 7.99 / 3.85; 2²⁴ ≈ 8.66.

Fits (Stage-0 `calibrate`: an upper-decile Gaussian QQ fit, with the threshold at a 10⁻⁶ tail; the union floor is √(2 ln(M/10⁻⁶)) = 7.06):

| class | n | best: mean / σ / threshold | margin: mean / σ / threshold | observed max best / margin |
|---|---|---|---|---|
| all | 2,160 | 3.42 / 0.95 / **7.92** | −0.58 / 0.85 / **3.44** | 5.99 / 1.87 |
| unmarked | 699 | 3.78 / 0.45 / 5.94 (floor 7.06) | −0.13 / 0.38 / 1.69 | 5.19 / 0.96 |
| wrong-key | 728 | 3.86 / 0.43 / 5.90 (floor 7.06) | −0.14 / 0.39 / 1.69 | 5.35 / 1.15 |
| out-of-table | 733 | 5.06 / 0.31 / 6.54 (floor 7.06) | 0.12 / 0.68 / 3.35 | 5.99 / 1.87 |

Fully scored nulls by kind: 699 unmarked, 728 wrong-key, 733 out-of-table. 148 attempts abstained (97 geometry, mostly q30/q10; 51 flat) and are not counted toward 2,000.

**The 10⁻⁶ claim rests on the fitted tail and the union bound, not on observation.**
- 2,160 null decisions with zero false attributions only bound the per-decision rate at about 1.4 × 10⁻³ (one-sided 95 %).
- Below that, the claim is the Gaussian extrapolation of the upper decile: at the shipped τ, the worst class tail is 9.6 × 10⁻⁷.
- The union bound M·Q(τ_abs) = 7.2 × 10⁻¹¹ assumes sub-Gaussian per-candidate scores.
- The 2²⁰ and 2²⁴ rows are further extrapolations of the same fit.

The pooled fit sets τ. Its upper decile mixes the out-of-table class, whose best sits about 1.2 higher (5.06 vs 3.8): the Stage-0 `messageFor` nulls collide with keyed BCH structure. That mixture is what widens σ.

## Survival, zoom and anchor reuse (published; §13.5 rows above are the gates)

| attack | correct / 135 | no-signal (geometry) | no match | innocent |
|---|---|---|---|---|
| zoom 0.75 | 135 | 0 | 0 | 0 |
| zoom 0.75 + q70 | 133 | 1 | 1 | 0 |
| zoom 1.1 | 135 | 0 | 0 | 0 |
| zoom 1.1 + q70 | 134 | 1 | 0 | 0 |
| zoom 1.25 | 135 | 0 | 0 | 0 |
| zoom 1.25 + q70 | 134 | 1 | 0 | 0 |
| zoom 1.5 | 135 | 0 | 0 | 0 |
| zoom 1.5 + q70 | 134 | 1 | 0 | 0 |
| zoom 1.75 | 135 | 0 | 0 | 0 |
| zoom 1.75 + q70 | 134 | 1 | 0 | 0 |

The single q70 miss in every q70 row (and in the gated q70 and q70 → down 0.5 rows) is the same page, `poi-comment.docx#0@1`. Its JPEG residual is 0.18–0.50, above GEOMETRY_ERROR_MAX 0.0923, so it answers **no usable signal — geometry**. Nobody is named.

**Anchor reuse (`viewer` visible marks, distinct texts)** ran on every 4th entry (35 entries).
- Setup: four viewers with distinct e-mail texts, served first, so the first is the registration anchor. One fixed-mark session follows, then fixed padding to 65,536.

| capture | traced correctly | no-signal (geometry) | no match | innocent |
|---|---|---|---|---|
| anchor viewer's own capture, default | 35/35 | 0 | 0 | 0 |
| other viewers (3 per entry), default | 61/105 = 58.1 % | 44 | 0 | 0 |
| other viewers, q70 | 60/105 = 57.1 % | 45 | 0 | 0 |
| other viewers, down 0.5 | 58/105 = 55.2 % | 47 | 0 | 0 |
| other viewers, crop 25 % | 51/105 = 48.6 % | 53 | 1 | 0 |

**The register-once-per-(policy, scale, generation) assumption does not hold for distinct viewer texts.**
- A viewer's capture registered against the anchor viewer's text leaves a residual above GEOMETRY_ERROR_MAX on 12 of the 35 entries, for every attack. Further entries fail on some attacks.
- These are mostly pptx slides with large flat areas and the docx pages; the per-entry lists are in the JSON.
- The tool then answers "geometry not supported" for a genuine leak. It never names anyone.
- Fixed-mark deployments are unaffected (the gates above).
- This is a finding for the owner, not a gate.

## Collusion (published rates; innocent named is the gate)

| attack | colluder named alone | collusion answer (colluders listed) | no match | no-signal | innocent named |
|---|---|---|---|---|---|
| average K = 2 | 0 | 135 | 0 | 0 | 0 |
| average K = 4 | 0 | 135 | 0 | 0 | 0 |
| average K = 6 | 0 | 135 | 0 | 0 | 0 |
| median K = 3 | 0 | 135 | 0 | 0 | 0 |
| median K = 4 | 0 | 135 | 0 | 0 | 0 |
| median K = 5 | 0 | 133 | 2 | 0 | 0 |
| median K = 6 | 0 | 135 | 0 | 0 | 0 |
| difference of 2 streams | 0 | 0 | 0 | 135 (geometry) | 0 |

At the fitted τ, every averaged or median capture produced at least two colluders above τ_abs, giving `collusion`. Every listed sid was a colluder. No colluder was ever named alone: at K = 2 the two colluders' Z values are close, so both clear τ_abs. A difference image does not register (its residual is far out of scope), so the answer is "geometry not supported".

## Capture geometry (out of scope ⇒ 0 attributions)

Residuals are registration residual / GEOMETRY_ERROR_MAX = 0.0923 (unchanged), over 135 units each:

| attack | min | p50 | p90 | max | above max (abstain) | scored ⇒ answer |
|---|---|---|---|---|---|---|
| default | 0.0000 | 0.0003 | 0.0018 | 0.0071 | 0 | |
| q70 | 0.0001 | 0.0087 | 0.0230 | 0.4525 | 1 | |
| down 0.5 | 0.0000 | 0.0003 | 0.0020 | 0.0121 | 0 | |
| q70 → down 0.5 | 0.0001 | 0.0031 | 0.0125 | 0.1768 | 1 | |
| crop 25 % | 0.0000 | 0.0005 | 0.0020 | 0.0102 | 0 | |
| zoom 0.75 / 1.1 / 1.25 / 1.5 / 1.75 | 0.0000 | ≤ 0.0003 | ≤ 0.0018 | ≤ 0.0087 | 0 | |
| zoom × + q70 (worst row: 0.75) | 0.0002 | 0.0097 | 0.0318 | 0.4983 | 1 each (poi-comment) | |
| **rotate 2°** | 0.0342 | 1.047 | 3.51 | 33.4 | 131 | 4 no match |
| **rotate 5°** | 0.0887 | 1.377 | 9.95 | 78.0 | 133 | 2 no match |
| **keystone 5 %** | 0.0343 | 0.646 | 2.74 | 44.7 | 117 | 18 no match |

- 24 out-of-scope captures register under the threshold and are scored. All 24 answer **no match**, and none is attributed.
  - Pages: lo-tdf104788 (all six slides, keystone, both buckets), lo-tdf44223, lo-tdf104792-smart-art-animation and poi-table_test.
  - Residuals: 0.034–0.089.
  - Per-page residuals for every geometry and in-scope capture are in `streaming-s2b.json` → `residuals`.
- The in-scope and out-of-scope residual ranges overlap. In-scope p90 is ≤ 0.032, while the smallest out-of-scope residual is 0.034. So the threshold alone does not separate them. The safeguard is that a misregistered capture scores at null level; this run shows that.
- GEOMETRY_ERROR_MAX was not changed.

## 7a's ambiguity band (registration hypotheses within max(.1, 5 %) MSE, decided by keyed support)

`register()` now reports `band`. This includes the band size, whether keyed support decided it, and whether support overrode the lowest-MSE hypothesis. Decisions are at the fitted τ. "Near-flat" pages are those whose canonical has luminance σ < 10 (38 of 135 marked entries).

| decisions | n | band hit | support overrode MSE | correct when hit | wrong name when hit | no-signal when hit |
|---|---|---|---|---|---|---|
| survival, all pages | 675 | 2 | 0 | 1 | 0 | 1 |
| zoom, all pages | 1,350 | 15 | 5 | 10 (3 of the 5 overrides) | 0 | 5 |
| geometry, all pages | 405 | 122 | 71 | 0 | 0 | 122 |
| collusion, all pages | 1,080 | 102 | 84 | 0 | 0 | 102 (the difference images) |
| viewer, all pages | 455 | 89 | 42 | 3 | 0 | 86 |
| **null, all pages** | 2,308 | **37** | 18 | — | **0** | 31 (other 6: no match) |
| survival, near-flat | 190 | 2 | 0 | 1 | 0 | 1 |
| zoom, near-flat | 380 | 15 | 5 | 10 | 0 | 5 |
| **null, near-flat** | 691 | **31** | 15 | — | **0** | 25 |
| survival / zoom, other pages | 1,455 | 0 | 0 | — | — | — |
| null, other pages | 1,617 | 6 | 3 | — | 0 | 6 |

- On every in-scope capture of a non-flat page, the band never fired.
- On near-flat pages it fired on 17 of 570 survival and zoom captures. Of those, 11 traced correctly and 6 abstained. None named anyone else.
- In the null set it fired on 37 decisions (31 on near-flat pages) and produced **0 false attributions**: 31 abstained and 6 answered no match. On these pages the mark decides, and it never manufactured a name.

## Runtime (this machine, 4 documents at a time; published, target ≤ 5 min at 10,000)

| candidates | traces | p50 | p95 | max |
|---|---|---|---|---|
| 10,000 | 118 | 0.29 s | 0.68 s | 0.92 s |
| 65,536 (batteries) | 3,510 | 2.37 s | 9.11 s | 24.6 s |
| 65,536 (nulls) | 2,308 | 2.24 s | 6.25 s | 232 s (one q10/q30 null on a scale-2 page) |

Whole run: 97 min wall, 6.4 CPU-hours.

## Reference determinism (finding; why the run is one process per document)

A first full run used 4 long-lived processes, each handling many documents. It failed with `reference-mismatch` on several xlsx units: `lo-calc-pivotcompact`, `poi-ExcelTables` and `lo-calc-shared-dategroup`.

Reduced case: `lo-calc-pivotcompact.xlsx#0@1` was rendered by a fresh `DocumentHost`.
- Alone in a process, it hashes to the served canonical (2ebbeb86…).
- After the same process has opened and rendered `lo-bnc591147.pptx`, the same render hashes to ceae69e1….
- Process-global rendering state (font registration) changes later renders.

Consequences:
- The trace tool fails closed ("reference differs from what was served"), so it never mis-decodes.
- A long-lived server's pool workers render many documents. The served canonical can therefore depend on what the worker rendered before. A fresh trace tool would then refuse a genuine leak.

The calibration therefore runs one process per document, which matches a fresh server and a fresh tool. In that mode the run had 0 reference mismatches. Not fixed here: owner decision.

## Native check

After `tau.ts` was rewritten, a native rerun (`--nulls 0 --viewer-every 0`) covered 4 documents (lo-airplane-transition, lo-bnc591147, poi-comment, poi-ExcelTables; 8 entries).
- `trace()`'s own answers were compared with the re-decided outcomes above, per entry, battery and attack, as the class: correct, innocent, collusion, no match or no-signal.
- Result: **208/208 agree**, with no differences. This covers 8 entries × (5 survival + 10 zoom + 3 geometry + 8 collusion). It includes poi-comment's q70 abstentions and the lo-airplane band hits.
