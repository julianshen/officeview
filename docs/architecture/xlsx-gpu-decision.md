# XLSX renderer backend decision: keep Canvas2D, S12 not run

Status: decision record, no prototype. Scope: plan S12 (`docs/superpowers/plans/2026-10-10-interactive-viewer.md`),
spec §11.1, §11.2, §12.2, §12.4, §13.2, BACKEND-01–02 (`docs/superpowers/specs/2026-10-10-interactive-viewer.md`;
§13.2 makes any font/clip/rotation/grapheme mismatch an admission failure). Rule for every number below: **measured**
(source cited) or **unmeasured** (labelled). Historical paint figures cite merged S10.2 r1
(`docs/validation/interactive-viewer-performance.md`, receipt
`docs/validation/interactive-viewer-viewer-chrome-r1.json`); that historical run was headless Chrome 155
(receipt `environment.headless = true`) on a shared machine (load
3.53/2.72/2.53 at start, 1.99/2.38/2.48 at end) with a quiet-machine re-run pending. Canvas2D acceleration mode is not
recorded in the receipt (unmeasured; spec §12.2 requires the Canvas/GPU backend). Nothing here is a Safari, iPad, or
mobile result — no Safari/iPad receipts exist (actual Safari not run; no physical iPad, user decision).
S10.2 merged 2026-10-13; historical r16 (`docs/validation/interactive-viewer-viewer-chrome-r16.json`) admits desktop D
13/13, including sort acknowledgment p95 17.1 ms (max 37.0 ms). The final release receipt must use the final tree.

## 0. Summary

| Question | Answer |
|---|---|
| Decision | S12 not run: Canvas2D meets every paint row and budget; the one paint bottleneck was fixed in place |
| Evidence | Warm input-to-paint p95 ≤ 3.94 ms, frame interval 16.67 ms with 0 long tasks, bitmap 113.2 vs 128 MiB (headless Chrome 155) |
| Historical r1 bottleneck | ~12 s detection I/O and missing pending UI on sort ack; paint uninvolved (commit-and-present < 30 ms). r16 acknowledgment now passes |
| Revisit | A paint row misses with paint-attributed cause; Safari/mobile receipts arrive; the quiet re-run moves a paint row |

## 1. Hypothesis a WebGPU renderer would test

Viewport paint is the frame bottleneck in Canvas2D, and bounded WebGPU batches (fills/lines, glyph-atlas quads from
host-shaped records) beat it end to end while staying inside the §12.4 bitmap pool (128 MiB D) with text/image/warp
parity and a working fallback.

## 2. Paint measurements vs frozen budget

Frozen rows: `tests/fixtures/xlsx-viewer/performance-profile.json` (`warm-viewport-frame` 16.7 ms). S10.2 receipt
(historical r1), desktop D headless Chrome, 300 measured frames after 60 warmup:

| Mix | Input-to-paint p95 | Frame-interval p95 | Long tasks |
|---|---:|---:|---|
| Ordinary (F2 1M) | 3.94 ms | 16.67 ms | 0 |
| Wrapped ASCII (20k×12) | 1.03 ms | 16.67 ms | 0 |
| Styled (20k×12) | 2.35 ms | 16.67 ms | 0 |

Bitmap pool: accounted peak 113.2 MiB vs the 128 MiB budget — the peak is the F2 after-operations checkpoint (receipt
`raw.memory.F2.checkpoints.afterOps.frame.peakBitmapBytes` = 118,663,680 B). Cross-checked against user-agent `Canvas`
bytes 43.2 MiB (receipt `raw.memory.F2.checkpoints.afterOps.uasm.canvas` = 45,263,360 B) and DOM canvas backing stores
15.2 MiB (`canvas.bytes` = 15,903,232 B, the two pane surfaces). Wide-F3 horizontal scroll (informational):
input-to-paint p95 4.3 ms, frame interval 16.67 ms.

## 3. The only paint bottleneck found, and its TS fix

Wrapped-text tile prefetch constructed a new `Intl.Segmenter` per
`graphemes()` call (`src/core/text-recording.ts` in the S10.2 tree): 4.7 s of
a 10 s profile, frame-interval p95 83.3 ms with 70 long tasks. Fix (in the S10.2 tree): one segmenter per host
constructor plus an ASCII fast path (UAX #29 segments printable ASCII per character). Result: 16.67 ms, 0 long tasks,
pinned by `tests/text-recording-graphemes.test.ts` (S10.2 tree) comparing against the host segmenter. Non-ASCII
wrapped mix (informational): p95 16.67 ms, max 33.3 ms. No paint-side bottleneck remains open.

Real-wheel diagnostic (informational, S10.2 doc): `page.mouse.wheel` input-to-paint p95 20.4 ms, max 23.4 ms,
attributed to the non-passive wheel listener (every wheel scroll waits for the main thread), which a GPU backend would
not change. CDP wheel events also arrive faster than once per frame, so this is not a threshold input; it joins the
excluded causes in trigger 1.

## 4. Why WebGPU would not move the miss, and what it would cost (**unmeasured** — no prototype built)

The historical S10.2 r1 miss was sort acknowledgment (11,897–13,316 ms), attributed to main-thread range-detection
RPC/deserialization and missing pending UI. No paint phase was on that path; commit-and-present measured < 30 ms.
Worker range detection and pending UI addressed these causes in the merged TypeScript baseline; historical r16
acknowledgment passes (p95 17.1 ms, max 37.0 ms).
A GPU renderer would not address those causes and would add: device init plus cold shader compile on first paint;
atlas/texture memory (per §12.4 both backends count only at a switch peak, so the real pressure is the switch/fallback
peak on top of a 113.2 MiB Canvas2D peak, not steady-state addition); text parity re-proof (fallback fonts, grapheme
segmentation, warps, watermark, overlays) with the Canvas2D reference kept anyway as fallback. The Canvas2D fallback
story already works where it matters: device loss drops tiles and repaints both layers on
`contextrestored` (validation matrix G9, emulated events, pixel alpha 0 → 255
with growing `basePaints`). A GPU backend would add its own device-loss and atlas-exhaustion modes and still depend on
that fallback (§13.2) — G9 proves the reference recovers; it does not prove a second backend would.

## 5. Decision

**S12 not run**, under the plan's optional execution boundary: the Canvas2D baseline meets every paint row and budget,
and the single found paint bottleneck was fixed in place.

Re-open triggers (any one justifies running S12 later):

1. A paint row (`warm-viewport-frame`, or frame interval via the same receipt)
   exceeds its frozen value on a required profile after non-backend causes are excluded (including main-thread wheel
   handling and record preparation), with a measured profile attributing the excess to draw/composite.
2. The bitmap pool exceeds 128 MiB and analysis shows atlas/compositing — not
   record preparation — would fix it within budget.
3. Actual desktop Safari receipts miss a paint row with a paint-attributed
   cause.
4. Mobile paint measurements (once an iPad is available) miss with a
   paint-attributed cause.

This decision is desktop Chrome (headless) only — revisit when Safari and mobile receipts exist, and if the
quiet-machine S10.2 re-run moves any paint row. The r1 measurements above remain historical; the final release
receipt must use the final tree.
