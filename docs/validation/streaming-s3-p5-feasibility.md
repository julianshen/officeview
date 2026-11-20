# Streaming Stage 3 piece 5 — feasibility gate: forensic Layer 1 on H.264

**Verdict: RED.** No CRF in {18, 20, 23} passes the gate, so by the pinning rule (below) the verdict is RED at CRF 18. The plan stops here; the owner decides the next step. No gate, threshold or method was changed to look for a GREEN.

- Harness: `bun scripts/stream-p5/survival.ts` (`--quick` runs mandelbrot only). Numbers: `streaming-s3-p5-feasibility.json` next to this file.
- Date 2026-11-17, ffmpeg 8.1.1 (`/opt/homebrew/bin/ffmpeg`), libx264 `-preset veryfast -threads 2` (controller ruling X31), Apple Silicon, 14 cores, shared machine (1-minute load 1.3–2.2 during the committed run).
- Determinism: seed `p5-t1`; every random choice is SHA-256 of `p5-t1:<label>`. Two full runs (runs 2 and 4) produced identical results: every cell, carrier count, self-check, null fit, gate and even the browser tallies. A `--quick` run (run 3) reproduced run 2's mandelbrot subset exactly, including all 2,001 null decisions (see §Runs).

## Pinned values (for Tasks 6–8 and 12)

| Item | Value | Note |
|---|---|---|
| CRF (D8) | **none passes; RED at 18** | Rule (controller ruling 4): the largest of 23, 20, 18 whose gated cells all pass; if none, RED at 18 |
| Preset (D8) | `veryfast` | Encode is 246–611 fps at 1280×720 with `-threads 2`, far above the wall cap's need |
| Reference (D10) | **R1** | R2 is never better at a passing point (§R1 vs R2) |
| τ (D11) | 9.15, margin 3.45 | Union rule `sqrt(7.93² + 2 ln(2³¹/2¹⁶))` = 9.1476, computed in the script; `MAX_CANDIDATES · 301 ≤ 2³¹` asserted |
| Marking fps | PNG path 3.4–7.0 fps at 1280×720 (**fails** 15 fps); raw RGBA core 2,130–2,510 fps | D10's `markLayerRgba` is required |
| Null fit (D11) | fitted 5.20 / 5.48 / 5.27 (CRF 18/20/23) ≤ union 6.65 at M = 64·61; without testsrc2 5.30 / 5.54 / 5.33 | 2,000 decisions per CRF, 0 false attributions |

## Gates at CRF 18 (the verdict CRF)

Gated cells: the 640×360 box at (320,180) and the full-frame box, both at bucket 1, and the full-frame box at bucket 2 (crop scale 0.5, ruling 3). Three sources each, 60 captures per cell. Rates are pooled over the three sources; per-cell rows follow.

| Gate | Needed | 640×360 | 1280×720 | 1280×720 @ bucket 2 (scale 0.5) | Result |
|---|---|---|---|---|---|
| Lossless capture | ≥ 90 % | 40/180 = 22.2 % | 120/180 = 66.7 % | 54/180 = 30.0 % | **RED** |
| JPEG q70 capture | ≥ 75 % | 0/180 = 0 % | 120/180 = 66.7 % | 0/180 = 0 % | **RED** |
| Wrong attributions | 0 | 0 | 0 | 0 | GREEN |
| Fitted null threshold ≤ union at M = 64·61 = 3,904 | ≤ 6.65 | 5.20 (all gated bucket-1 cells together) | | | GREEN |
| 5-minute 1280×720 size, worst source | ≤ 64 MiB | | 293.5 MiB (mandelbrot, 1,009,899 B/s video + 16 kB/s AAC) | | **RED** |

Per cell at CRF 18 (lossless / q70 correct out of 60, R1, median z of the true code):

| Source | 640×360 | 1280×720 | 1280×720 @ b2 |
|---|---|---|---|
| testsrc2 | 0 / 0 (z 1.54) | 0 / 0 (z 4.31) | 0 / 0 (z 2.97) |
| mandelbrot | 0 / 0 (z 6.76) | 60 / 60 (z 14.06) | 0 / 0 (z 5.85) |
| Ken Burns | 40 / 0 (z 9.29) | 60 / 60 (z 19.28) | 54 / 0 (z 10.90) |

Not a gate, for reading only: without testsrc2 (see Concern 2), CRF 18 gives 640×360 40/120 lossless and 0/120 q70, full frame 120/120 and 120/120, bucket 2 54/120 and 0/120. That is still RED on two of the three gated boxes and on size.

## Every CRF

Pooled over the three sources, R1 (lossless / q70 rate):

| CRF | 640×360 | 1280×720 | 1280×720 @ b2 | Wrong | Null fit (≤ 6.65) | 5-min size | Passes |
|---|---|---|---|---|---|---|---|
| 18 | 0.22 / 0.00 | 0.67 / 0.67 | 0.30 / 0.00 | 0 | 5.20 | 293.5 MiB | no |
| 20 | 0.00 / 0.00 | 0.58 / 0.45 | 0.02 / 0.00 | 0 | 5.48 | 234.1 MiB | no |
| 23 | 0.00 / 0.00 | 0.34 / 0.19 | 0.00 / 0.00 | 0 | 5.27 | 160.8 MiB | no |

All cells (lossless / q70 out of 60, R1; video KiB/s):

| Source | Box | qp 0 (sanity) | CRF 18 | CRF 20 | CRF 23 |
|---|---|---|---|---|---|
| testsrc2 | 640×360 | — | 0/0, 124 | 0/0, 98 | 0/0, 65 |
| testsrc2 | 1280×720 | — | 0/0, 505 | 0/0, 421 | 0/0, 311 |
| testsrc2 | 1280×720 b2 | — | 0/0, 510 | 0/0, 424 | 0/0, 312 |
| mandelbrot | 640×360 | 60/44, 1992 | 0/0, 181 | 0/0, 131 | 0/0, 78 |
| mandelbrot | 1280×720 | 60/60, 7204 | 60/60, 986 | 45/21, 783 | 2/0, 533 |
| mandelbrot | 1280×720 b2 | 60/24, 7260 | 0/0, 992 | 0/0, 788 | 0/0, 537 |
| Ken Burns | 640×360 | 60/60, 686 | 40/0, 26 | 0/0, 18 | 0/0, 11 |
| Ken Burns | 1280×720 | 60/60, 1910 | 60/60, 88 | 60/60, 63 | 60/34, 38 |
| Ken Burns | 1280×720 b2 | 60/60, 1919 | 54/0, 88 | 3/0, 63 | 0/0, 40 |
| *320×180 (not gated)* | all sources | — | 0/0 everywhere | 0/0 | 0/0 |

The qp 0 rows are the harness sanity check: through the same RGBA → yuv420p → RGBA path with no quantisation, every source with enough carriers attributes 60/60 lossless. The loss at CRF 18–23 is the codec's, not the harness's. The unencoded self-check is 60/60 on every cell whose carrier ceiling clears τ (mandelbrot and Ken Burns at 640×360, full frame and bucket 2).

## Why it fails

1. **Carrier budget (structural).** The page carries 512 chips in total. A video box holds only its area's share of them, and the best possible z is about √carriers:

   | Box | Carriers (median frame) | Ceiling √carriers | τ |
   |---|---|---|---|
   | 640×360 | 109 (mandelbrot), 116 (Ken Burns), 3 (testsrc2) | 10.4–10.8 | 9.15 |
   | 1280×720 | 465, 512, 32 | 21.6–22.6 | 9.15 |
   | 1280×720 @ b2 | 487, 512, 47 | 22.1–22.6 | 9.15 |
   | 320×180 | 28, 31, 0 | 5.3–5.6 | 9.15 |

   At 640×360 the ceiling is 1.1× τ, so any codec noise sinks it. Even qp 0 fails q70 there (mandelbrot 44/60). This box cannot pass at the D11 window τ whatever the CRF. The 320×180 box cannot attribute even unencoded.
2. **H.264 on a ±2 4×4-pixel pattern.** Full frame keeps z ≈ 14–19 at CRF 18, which passes, but mandelbrot drops to 10.1 at CRF 20 and 5.6 at CRF 23. At bucket 2, the 4-px page cells become 2×2 frame px, and x264 removes them: z falls from about 22 to 5.9 (mandelbrot) and 10.9 (Ken Burns) at CRF 18, and q70 removes the rest.
3. **Saturated content.** `pedestalFor` masks every cell whose channels span more than 251. testsrc2 is saturated almost everywhere, so it keeps 3 of about 110 carriers at 640×360 and fails before any encode.
4. **Size.** High-entropy full-frame video (mandelbrot) is about 1 MB/s at CRF 18. A 5-minute 1280×720 encode would therefore be 293.5 MiB, against 64 MiB, and still 161 MiB at CRF 23. On the same basis (video plus the 16 kB/s AAC allowance), at CRF 18, testsrc2 is 152.5 MiB and Ken Burns is 30.5 MiB. Only Ken Burns fits.

## R1 vs R2 (D10)

R2 (the unmarked encode, decoded, used as the reference) was run on every bucket-1 gated cell at every CRF. It is never better where anything passes: mandelbrot full frame at CRF 18 gives R1 60/60 and 60/60, R2 60/60 and 47/60; Ken Burns 640×360 at CRF 18 gives R1 40/60, R2 9/60. The only R2 gain is Ken Burns full frame at CRF 23, lossless 60/60 against 60/60 with a higher median z (12.05 against 11.45), where both fail q70. **R1 stands.**

## Null window decisions (D11, step 8)

- **Setup.** 2,000 decisions per CRF, 6,000 in all, over every source and both bucket-1 gated boxes, alternating:
  - unmarked captures (visible mark only, encoded) scored against the 64 session codes;
  - marked captures scored against 64 out-of-table codes;
  - lossless and q70 captures.
  Each decision scores F = 61 reference frames (t ± 30) × 64 codes = 3,904 candidates. All 122,000 frame scorings per CRF registered and were scored.
- **Results:**

  | CRF | False attributions | Fitted abs threshold (`calibrate`) | `calibrate` union at 3,904 | D11 rule at 3,904 | Fitted margin |
  |---|---|---|---|---|---|
  | 18 | 0 | 5.20 | 6.65 | 7.57 | 1.38 |
  | 20 | 0 | 5.48 | 6.65 | 7.57 | 1.40 |
  | 23 | 0 | 5.27 | 6.65 | 7.57 | 1.38 |

- **Without testsrc2.** A third of the decisions sit on testsrc2 cells with only 3 or 32 carriers. Refitted without them (1,333 per CRF), the fitted threshold is 5.30, 5.54 and 5.33, and the fitted margin is 1.55–1.62. Both are still below 6.65, with 0 false attributions.
- **Reading.** Both readings of "union value" agree: the fitted threshold is below both. The tail fit supports carrying the union rule to F = 301 (τ = 9.15). The null side is GREEN; survival is what fails.

## Other measurements

- **Marking fps** (0.5× realtime needs ≥ 15 fps at 30 fps):

  | Box | PNG round trip | Raw RGBA core | Visible composite |
  |---|---|---|---|
  | 1280×720 | 3.4–7.0 | 2,130–2,510 | 600–680 |
  | 640×360 | 13.7–27.9 | 3,990–4,190 | 2,500–2,640 |

  The raw core was checked byte-equal to `markLayerRaster` on 30 frames of every cell. Load was 1.3–2.2 during these measurements.
- **Encode wall time** (`veryfast`, `-threads 2`, 180 frames): at most 0.73 s at 1280×720 (246–611 fps) and at most 0.19 s at 640×360 (946–1,941 fps). The 6-s cap is 42 s.
- **Browser-decoded cell (D20).** Chrome 157.0.8103.52 (`channel: 'chrome'`), the 640×360 MP4 at CRF 18, seek to (i + 0.5)/30, element screenshot at DPR 1 with `--force-color-profile=srgb`. `canPlayType` returned `probably`. Results:

  | Source | Correct | Median true z (vs ffmpeg-decoded) | Geometry abstains | Screenshots nearest a neighbour R1 frame |
  |---|---|---|---|---|
  | testsrc2 | 0/60 | −0.32 (1.54) | 0 | 0 |
  | mandelbrot | 0/60 | 2.51 (6.76) | 9 | 0 |
  | Ken Burns | 0/60 | 5.09 (9.29) | 19 | 18 |

  - **Colour shift, measured.** A follow-up compared Chrome screenshots with ffmpeg's decode of the same MP4 frame, at frames 31, 91 and 151, CRF 18, 640×360. Chrome's pixels are systematically brighter:

    | Source | Mean signed difference (R, G, B) | Mean absolute | MSE |
    |---|---|---|---|
    | mandelbrot | +4.6…5.3, +8.0…10.3, +2.4…4.1 | 6.3–6.7 | ≈ 63–65 |
    | Ken Burns | +2.2…3.1, +2.7…3.3, +2.7…3.5 | 2.5–3.3 | 8–16 |

  - **Seek is exact.** In that comparison every screenshot matched the requested frame of the ffmpeg decode best. The 18 Ken Burns "neighbour frame" offsets in the table above were measured against R1, where slow-zoom neighbours are ambiguous.
  - **Cause.** The encode's RGB → YUV matrix is ruled out: it is BT.709, matching the tags (RGB (200,50,50) encodes to Y = 86, the BT.709 value; BT.601 would give 97). The brightness shift is larger than rounding, which points to range, transfer or colour-management handling in Chrome's path, but the cause was not isolated.
  - **Risk.** R1, an ffmpeg bitexact decode, may not model what a viewer's browser displays. The detector would then need to rebuild its reference the browser's way, or tolerate this shift.

## Method notes (deviations, stated)

- **Detection path.** `detect(capture, page)` abstains on a 640×360 region of a 1280×720 page, because `register` admits crops of at least 75 % only, and a frame-sized reference cannot hold the page's chips. The harness therefore:
  1. registers the capture against the true-time reference frame, under the same `GEOMETRY_ERROR_MAX` gate;
  2. moves the pose by the audited crop offset;
  3. scores with `scoreRegistered` against a page-sized reference with the frame at the crop;
  4. decides with `decideScores` (τ 9.15, margin 3.45) over all 64 codes.

  This is what `trace.ts` does. The bucket-2 cell (scale 0.5) uses a minimal frame-space copy of `scoreAligned`, checked equal (difference 0) to `scoreRegistered` on every scale-1 self-check.
- **Null windows register once**, at the window's centre frame, and reuse that pose for all 61 frames, as `trace.ts` reuses one pose across references. Registering each of the 61 frames separately ran `register`'s zoom search (Skia resizes) on wrong-time frames, and the process deadlocked inside `@napi-rs/canvas` twice (threads in `__ulock_wait`). The cause was not investigated further.
- **Sources.**
  - testsrc2 and mandelbrot come from lavfi.
  - Ken Burns is a `zoompan` from 1× to 1.36× over `ppt/media/image1.jpeg` (720×540) of `corpus/lo-tdf104788.pptx`.
  - Each is 6 s at 30 fps, stored as FFV1, then decoded with the D8 decode argv.
  - The sources have no audio, so the encode has no input 1 and no AAC options.
  - The D7 input options are not applied.
- **Alignment check.** decoded[i] is nearest R1[i] for frames 10, 90 and 170 everywhere except Ken Burns at CRF 23 (all three gated boxes) and Ken Burns 320×180 at every CRF. There, slow-zoom neighbours are nearly identical and frame 170 (once also 90) is closest to frame i−1. qp 0 attributes 60/60 there, so the frames are not offset.

## Runs

- **Run 1** used the default x264 threads, before ruling X31. Its numbers are superseded and not committed. They differ only slightly; for example, CRF 18 full frame was 120/180 lossless and 118/180 q70.
- **Run 2** used `-threads 2` and covered the full matrix.
- **Run 3** was `--quick` with `-threads 2`: identical to run 2's mandelbrot subset, including every null decision.
- **Run 4**, the committed JSON, used `-threads 2` with a slimmer JSON: per-group digests replace the 6,000-row null list. It is identical to run 2 in every deterministic field compared (cells, carriers, self-checks, null fits, gates, browser tallies).

Every number in this report comes from the `-threads 2` runs.
