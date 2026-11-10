# Server streaming — Stage 2a validation (Layer-1 forensic mark on served output)

Spec: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` §13.2–13.5, §16 T7/T9, §17 Stage 2. Plan: `docs/superpowers/plans/2026-11-09-streaming-s2a-forensic-mark.md`. Stage-0 numbers: `docs/validation/streaming.md`.

Run (at bc0d7228, status **GREEN**):

```text
bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s2/l1-integration.ts --write docs/validation/streaming-s2.json
```

What it measures: the **production** path. `DocumentHost` renders the canonical PNG (visible mark composited, as cached by the server), `markRaster` (the pool's encode-job body) embeds the session code, and the result is the PNG **bytes** a viewer receives. Every corpus unit at bucket 1, every pptx unit also at bucket 2. Codec `l1-bch512x40@1`, parameters `{"nChips": 512, "cellPx": 4, "tilePx": 64, "amplitude": 2, "pedestal": 2}`. Candidates: 65,536 rows issued by a fresh `AuditTable` (tokens 0…65,535, keyed `msg_40`); every unit is marked with row 0. Detector: Stage-0 `scoreCapture`/`decideScores` at the **Stage-0 τ** (τ_abs 7.99, τ_margin 3.85), not recalibrated.

## Units

| | rendered | marked | unmarked (capacity) | skipped (refused at admission) |
|---|---|---|---|---|
| all | 141 | 135 | 6 | 6 |
| docx | 11 | 11 | | |
| xlsx | 44 | 38 | | |
| pptx | 86 | 86 | | |

Per-unit carriers (T7 publication; 512 = no masked chip): min 418, p10 482, median 511, max 512. Per-unit rows (carriers, masked, pedestal cells, ΔE00, each attack's decision) are in `streaming-s2.json`.

Units that ship **unmarked** (visible mark only; the server sends `diag forensic-unmarked` with the reason, and counts them in `stats().forensic.unmarked`):

| unit | size | reason |
|---|---|---|
| `poi-47889.xlsx#1@1` | 64×20 | capacity |
| `poi-47889.xlsx#2@1` | 64×20 | capacity |
| `poi-62272.xlsx#0@1` | 81×17 | capacity |
| `poi-66365.xlsx#0@1` | 176×40 | capacity |
| `poi-link-external-workbook-a.xlsx#0@1` | 128×20 | capacity |
| `poi-xlsx-jdbc.xlsx#0@1` | 192×40 | capacity |

Skipped documents (the §6.4 fail-closed strip refuses them; no session can open them, so nothing is served):

| document | reason |
|---|---|
| `lo-calc-PivotTableFormatsSubTotals.xlsx` | unsupported-format: unsupported-format: unowned visibility flag definedName hidden=1 in xl/workbook.xml |
| `lo-calc-pivottable_tabular_mode.xlsx` | unsupported-format: unsupported-format: unowned visibility flag pane state=frozen in xl/worksheets/sheet2.xml |
| `poi-55733.docx` | refused:unsupported-format |
| `poi-59021.xlsx` | refused:unsupported-format |
| `poi-59746_NoRowNums.xlsx` | refused:unsupported-format |
| `poi-crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx` | refused:unsupported-format |

## ΔE00 (served bytes vs canonical)

Page max over all marked units: **1.59** (ceiling 3; Stage 0 measured 1.59). GREEN.

## Survival (§13.5 rows, Stage-0 τ, M = 65,536)

"marked" = attributed to row 0 / marked units (the §13.5 definition, image-only units). "served" counts the unmarked units as misses too, i.e. the rate over every rendered unit.

| attack | correct / marked | rate | served | target | innocent | status |
|---|---|---|---|---|---|---|
| default | 135/135 | 100.0 % | 135/141 = 95.7 % | 90 % | 0 | GREEN |
| q70 | 135/135 | 100.0 % | 135/141 = 95.7 % | 75 % | 0 | GREEN |
| down0.5 | 135/135 | 100.0 % | 135/141 = 95.7 % | 75 % | 0 | GREEN |
| q70-down0.5 | 135/135 | 100.0 % | 135/141 = 95.7 % | 75 % | 0 | GREEN |
| crop25 | 135/135 | 100.0 % | 135/141 = 95.7 % | 70 % | 0 | GREEN |

Stage 0 (same parameters and τ, prototype pipeline, 66 pages × 5 tokens): every one of these rows was 90.9 %, its misses being units below carrier capacity counted in the denominator. Here no marked unit misses under any row; the served-output rate is 95.7 % for the same reason Stage 0 lost its 9.1 % (capacity).

## Unmarked captures never attribute

For every marked unit, the canonical (unmarked) bytes, both as captured and after JPEG q70, were scored against the same 65,536 rows: **270 captures, 0 attributed, 0 collusion-suspected**. GREEN.

## Serving path (server suite)

`server/test/forensic-serve.test.ts`: two sessions on one unit receive different bytes, each detecting to its own audit row, while the render cache holds one canonical entry. Forensic-off sessions receive the canonical bytes and run no encode job. Relaxed sessions receive command units with no encode job. A 100×50 unit ships canonical bytes plus `diag forensic-unmarked`. The scratch lease admits at most 2 encodes per session. Closing mid-encode sends no frame, and the charge is refunded only after the worker replies. Encode-job timing in process: 816×1056 66 ms, 1632×2112 259 ms, 1920×1440 198 ms (the 2 s render deadline also covers encodes).

## What this does NOT claim

- No recalibrated τ: Stage-0 τ is reused. τ over the composed production decision and the ≤ 10⁻⁶ false-attribution gate on the production pipeline are **2b**.
- No collusion numbers on production output (2b).
- No detection CLI or audit export (`server/scripts/detect.ts`, 2b).
- No Layer 2: command units (`projection: 'relaxed'`) carry no forensic mark; the manifest says `forensic-commands-unmarked`.
- No audit persistence: rows and the token counter live in memory; a restart forgets them (2b: persistence, retention, key rotation).
- Detection here uses the exact canonical reference and the known page key, as Stage 0 did; registration against a re-rendered reference is 2b's detector work.

## Handoffs

- **2b:** detect CLI + audit export; τ over the composed decision and the 10⁻⁶ gate; collusion battery on production output; audit persistence, retention, key rotation (T10).
- **2c:** occlusion store (command culling) and the M-rule prototype with its shipping gate.
