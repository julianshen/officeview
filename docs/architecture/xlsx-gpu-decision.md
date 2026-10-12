# XLSX renderer backend decision: keep Canvas2D, S12 not run

Status: decision record, no prototype. Scope: plan S12
(`docs/superpowers/plans/2026-10-10-interactive-viewer.md`), spec §5.5, §12.2,
§12.4 (`docs/superpowers/specs/2026-10-10-interactive-viewer.md`).
Rule for every number below: **measured** (source cited) or **unmeasured**
(labelled). S10.2 numbers are from branch `feature/iv-s10-2`, file
`docs/validation/interactive-viewer-performance.md` and receipt
`docs/validation/interactive-viewer-viewer-chrome-r1.json`, which are **pending
merge**; the run was on a shared machine (load 3.53/2.72/2.53 at start,
1.99/2.38/2.48 at end) with a quiet-machine re-run pending. Nothing here is a
Safari, iPad, or mobile result.

## 1. Hypothesis a WebGPU renderer would test

Viewport paint is the frame bottleneck in Canvas2D, and bounded WebGPU batches
(fills/lines, glyph-atlas quads from host-shaped records) beat it end to end
while staying inside the §12.4 bitmap pool (128 MiB D) with text/image/warp
parity and a working fallback.

## 2. Paint measurements vs frozen budget

Frozen rows: `tests/fixtures/xlsx-viewer/performance-profile.json`
(`warm-viewport-frame` 16.7 ms, `input-acknowledgment` 50 ms). S10.2 receipt
(pending merge), desktop D Chrome, 300 measured frames after 60 warmup:

| Mix | Input-to-paint p95 | Frame-interval p95 | Long tasks |
|---|---:|---:|---|
| Ordinary (F2 1M) | 3.94 ms | 16.67 ms | 0 |
| Wrapped ASCII (20k×12) | 1.03 ms | 16.67 ms | 0 |
| Styled (20k×12) | 2.35 ms | 16.67 ms | 0 |

Bitmap pool: accounted peak 113.2 MiB vs the 128 MiB budget, cross-checked
against user-agent `Canvas` bytes (45.3 MiB on F2) and DOM canvas backing
stores (15.9 MiB). Wide-F3 horizontal scroll (informational): input-to-paint
p95 4.3 ms, frame interval 16.67 ms.

## 3. The only paint bottleneck found, and its TS fix

Wrapped-text tile prefetch constructed a new `Intl.Segmenter` per
`graphemes()` call (`src/core/text-recording.ts`): 4.7 s of a 10 s profile,
frame-interval p95 83.3 ms with 70 long tasks. Fix (in the S10.2 tree): one
segmenter per host constructor plus an ASCII fast path (UAX #29 segments
printable ASCII per character). Result: 16.67 ms, 0 long tasks, pinned by
`text-recording-graphemes.test.ts`. Non-ASCII wrapped mix (informational):
p95 16.67 ms, max 33.3 ms. No paint-side bottleneck remains open.

## 4. Why WebGPU would not move the miss, and what it would cost

The one S10.2 miss is sort acknowledgment (12.0–13.3 s), attributed to
main-thread range-detection RPC/deserialization plus missing pending UI — no
paint phase is on that path, and commit-and-present measures < 30 ms. A GPU
renderer leaves the miss untouched while adding: device init plus cold
shader compile on first paint; atlas/texture memory charged against the same
128 MiB pool that already peaks at 113.2 MiB; text parity re-proof (fallback
fonts, grapheme segmentation, warps, watermark, overlays) with the Canvas2D
reference kept anyway as fallback; and an unknown Safari/iPad coverage story
(no WebGPU receipts can exist there yet). The Canvas2D fallback story already
works where it matters: device loss drops tiles and repaints both layers on
`contextrestored` (validation matrix G9, emulated events, pixel alpha 0 → 255
with growing `basePaints`), so the failure mode a second backend would insure
against is already covered.

## 5. Decision

**S12 not run**, under the plan's optional execution boundary: the Canvas2D
baseline meets every paint row and budget, and the single found paint
bottleneck was fixed in place.

Re-open triggers (any one justifies running S12 later):

1. A paint row (`warm-viewport-frame`, or frame interval via the same receipt)
   exceeds its frozen value on a required profile after non-backend causes are
   excluded, with a measured profile attributing the excess to draw/composite.
2. The bitmap pool exceeds 128 MiB and analysis shows atlas/compositing — not
   record preparation — would fix it within budget.
3. Mobile paint measurements (once an iPad is available) miss with a
   paint-attributed cause.

This decision is desktop-evidenced only: mobile/iPad rows were not run (user
decision), so it must be revisited when mobile receipts exist, and if the
quiet-machine S10.2 re-run changes any paint row.
