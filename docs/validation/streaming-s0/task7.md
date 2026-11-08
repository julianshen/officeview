# Task 6 fixes and Task 7 detector

Completed, uncommitted as requested. Base HEAD at measurement: `767e26ec88dec38550a5f755b2377861a3896c6c`. No commit/push, no source changes, no installs/network. Concurrent Task 1 edits are not mine and were not modified. No Task 8 battery or calibration implemented.

## Files owned by this round

Part A:
- scripts/stream-s0/mark/embed.ts
- scripts/stream-s0/mark/ciede2000.ts
- tests/stream-s0/mark-embed.test.ts
- docs/validation/streaming-s0/task6.json
- docs/validation/streaming-s0/task6.md (historical report appendix)

Part B:
- scripts/stream-s0/mark/detect.ts
- scripts/stream-s0/mark/register.ts
- scripts/stream-s0/mark/attacks.ts
- tests/stream-s0/mark-detect.test.ts
- docs/validation/streaming-s0/task7-null.json
- docs/validation/streaming-s0/task7.md

The requested marker `../_prompts/stream-s0-task7.done` contains this report.

## RED / GREEN

Part A tests were changed first: 3 failed / 4 passed (missing realised chip deltas, missing shared pedestal helper, missing chipPixelsP95). All 34 Sharma pairs already passed the existing math; it did not need a numerical fix. New membership checks use independent carryless GF multiplication and Horner evaluation at roots alpha^1..alpha^190 on 100 seeded random words, plus 100 linearity pairs. No random-pair-distance assertion remains. Designed BCH distance >=191 is justified by membership in the specified narrow-sense code, not a sampled minimum.

Part B RED: missing detect module. Implemented detector, registration and attacks. All planned checks pass, plus null statistics, arithmetic attack checks and 2x crop registration.

Final required command:
`bunx vitest run --config vitest.stream-s0.config.ts tests/stream-s0/mark-*.test.ts`

```
Test Files  2 passed (2)
     Tests  15 passed (15)
  Duration  1.29s (transform 210ms, setup 0ms, import 558ms, tests 1.61s, environment 0ms)
```

Focused TypeScript passed:
`bunx tsc --noEmit --target ES2022 --module ESNext --moduleResolution bundler --types node --skipLibCheck scripts/stream-s0/mark/*.ts tests/stream-s0/mark-*.test.ts`

Task-owned `git diff --check` passed. No whole-product build claimed.

## Measured results

- Full 34-pair Sharma–Wu–Dalal offline supplementary dataset: error <1e-4 each, both directions. Hue-wrap branches included.
- White-page **centred modulation**: 1.999969482421875 >= 0.8×2. This replaces mean absolute delta, which was confounded by pedestal DC.
- White-page ΔE00: mean 0.04953837075989174; page p95/max and chipPixelsP95 0.7957222220882249.
- Null scores, 12 captures/class × 1024 real extended BCH candidates:
  - Unmarked with small acquisition noise: mean **0.0011484085197206392**, std **0.9937983198499535**, N=12288.
  - Wrong-key marked: mean **0.0021140494138387656**, std **0.9982397572855789**, N=12288.
- Clean text-page, JPEG q70, half-scale + quarter crop at three anchors, and double-scale + quarter centre crop attribute candidate 73 among M=1024.
- Exact reference-identical capture abstains. Noisy unmarked capture returns none at the smoke thresholds.

`task6.json` regenerated with live `git rev-parse HEAD`; removed `uncommitted` and empirical-distance claims. `task7-null.json` is reproducible by prefixing the test command with `S0_MARK_WRITE_RESULTS=1`; ordinary tests do not write it. The JSON gate is only T9-null-smoke. No 10^-6 full-decision rate or corpus survival rate is claimed.

## Decisions and detector contracts

1. **Headroom:** `pedestalFor(referenceCell, params)` accepts a Bitmap cell and returns an equal-channel shift or null. Scan all RGB channel extrema and require both signs to fit: shift ∈ [amplitude−min, 255−amplitude−max]. Pick the shift closest to zero, only when its absolute magnitude <= pedestal; otherwise mask. Nonopaque cells mask. Endpoint 0/255 after modulation is valid because headroom was reserved; these chips are not dropped merely for having white reference pixels. This admits pale (255,255,200) with shift −2, black with +2, and refuses (255,0,0). Pedestal zero masks cells needing any shift.
2. **Recorded embedding:** `chipDelta[n]` is actual final-minus-reference Rec.709-weighted byte-sRGB luminance, including pedestal, measured from rounded output. `chipPedestal[n]` records the applied shift, NaN for masked cells; masked chipDelta is zero. `chipSupport` preserves positions for pageDeltaE. The legacy `mark.variant='white-pedestal'` label also covers newly admitted dark/pale headroom shifts; the actual per-chip array is authoritative. Input remains immutable; alpha preserved.
3. **ΔE:** identical RGBA pixels bypass Lab/ΔE work but still count as zeros in page statistics. Optional third `chips` argument selects exact carrier pixel support; default reads marked output's chipSupport, otherwise chipPixelsP95 is zero (no declared support). Support includes planned cells, including masked/unchanged pixels; max over the whole page remains the Task 8 gate. Exact nearest-rank p95, unchanged sRGB/D65/kL=kC=kH conventions. Measurement storage/sort remain offline O(pixels)/O(pixels log pixels).
4. **Registration:** search factors {1,0.5,2}. A candidate is dimensionally admissible when capture retains at least floor(75%) of each scaled reference axis and does not exceed it. Reference is resized with napi high-quality smoothing. Translation uses mean-subtracted, zero-padded 2D FFT phase correlation on a sampled grid (stride ceil(max reference axis/128), minimum 1), keeps 12 coarse peaks, then searches ±one stride at integer capture-pixel resolution using sampled luminance squared error (about 4096 samples per pose). Equal dimensions have only zero translation. Select the lowest-error content alignment, independently of candidate tokens. This is discrete scale/translation, not rotation/subpixel registration. Flat cropped content has ambiguous alignment; this prototype makes no recovery guarantee there. No extrapolated or out-of-bounds sampling.
5. **Coordinates/support:** a reference chip maps to [ceil(start×factor−translation), floor(end×factor−translation)) in capture pixels. Only whole surviving cells with nonempty support are used. Downscale averages capture cells directly against the correspondingly scaled reference. Preserve original n after masking/crop; template is sign0[n] × extendedCode[n%512]. Positive modulation is positive residual.
6. **Residual/DC:** subtract reference, subtract the shared per-chip pedestal, then centre r on surviving support. No 5×5 band-pass. Denominator sqrt(sum centred-r²) equals sigma_RMS×sqrt(L). Template centring is implicit since centred residual sums to zero. Exact unmarked reference is detected before pedestal subtraction to avoid manufacturing a signal. <2 usable cells or effectively zero residual energy abstains (`flat`). Invalid registration returns none; no candidates gives zBest=−Infinity. Tiny numerical energy cutoff is 1e-16, not a statistical attribution gate.
7. **Scoring:** O(M×L), real codeword40→extendedCodeword candidates in tests; sign-vector length/content validated at entry. Root membership is tested in the codec, not redundantly recomputed per detect call. Export scoreCapture for Task 8 calibration (scores, support count, registration and flat flag); detect applies abs and best-minus-second thresholds. Smoke tau={abs:6,margin:2} is explicitly **not** calibrated for deployment. Ties cannot meet the positive margin.
8. **Null experiment:** exact unmarked captures abstain and have no Z variance; the requested unmarked null distribution uses unmarked white pages with 0/−1 byte acquisition noise, no embed/pedestal. Wrong-key inputs use 12 other basis keys on the same white page. Fixed seeded real BCH candidates and deterministic noise; scores pooled over all candidates. This is marginal-Z smoke evidence, not the full alignment/crop/scale search maxima tail proof required by Task 8.
9. **Attack conventions:** JPEG quality is integer 1..100, napi encode('jpeg', q) then napi decode. Scale rounds dimensions, uses high-quality image smoothing. Crop frac removes that fraction of **each axis**, not area; anchors top-left/center/bottom-right retain the named part, integer-floor offsets. Compose is async and ordered left-to-right. Average and median operate in encoded byte RGB/alpha with final clamped rounding; even median averages the two middle values. Difference is signed a−b about byte 128 with opaque alpha (not a claimed best reconstruction attack). Dimension mismatch/empty image sets refuse.
10. **Fixture and scope:** a deterministic in-memory napi-rendered text page (14 rows plus a coloured rectangle) is the single Task 7 fixture. No corpus fidelity, all-white crop localisation, arbitrary rotation, attack survival percentages, collusion success or complete-decision false-attribution threshold is inferred from its pass. Existing xorshift/key conventions from Task 6 are unchanged. No Task 8 work performed.

## Fix round 1 — I1–I3 and minors

Uncommitted; no source/harness/census edits. Measurement HEAD `766abc6544b313b282656e9963124753cc3e575e` (concurrent harness commits may advance the branch).

### Files

- scripts/stream-s0/mark/register.ts
- scripts/stream-s0/mark/detect.ts
- tests/stream-s0/mark-detect.test.ts
- docs/validation/streaming-s0/task7.md
- docs/validation/streaming-s0/task7-fix1-before-0.25-0.5.json
- docs/validation/streaming-s0/task7-fix1-before-0.05-0.5.json
- docs/validation/streaming-s0/task7-fix1-before-0.25-2.json
- docs/validation/streaming-s0/task7-fix1-after-0.25-0.5.json
- docs/validation/streaming-s0/task7-fix1-after-0.05-0.5.json
- docs/validation/streaming-s0/task7-fix1-after-0.25-2.json
- docs/validation/streaming-s0/task7-fix1-null.json
- docs/validation/streaming-s0/task7-fix1-bench.json

### Verification and changes

Tests first: odd-size crop cases plus prepared matrix ownership. RED: 4 failed / 8 passed. Both crop-before-half-scale cases returned none; crop-before-double-scale returned none with no pose; prepareCandidates did not exist. Existing scale-before-crop tests retained unchanged.

I1 resolved: registration now supports both operation orders. Crop-then-scale searches integer offsets in ORIGINAL reference pixels (therefore half-pixel scaled offsets at f=0.5), considers the integer crop extents consistent with rounded capture size, and actually crops/resamples the reference using napi for final alignment. Coarse phase-correlation peaks seed the search. Approximate bilinear errors on ~256 samples rank each extent's candidates; the best eight are rendered with the real high-quality napi scaler and compared on ~4096 samples. Scale-then-crop retains its prior search. The returned reference is aligned to capture dimensions; scaleX/scaleY explicitly account for dimension rounding. Detector sampling uses original-index chips and this transform, without out-of-bounds reads. This remains a bounded heuristic registration search, not an exhaustive rotation/subpixel search.

I2 resolved: dimensional admission permits ceil(factor) slack below the nominal 75%-extent bound, covering crop-before-scale rounding. Captures larger than the full scaled reference remain rejected because there is no real support to sample. Extent inversion uses Math.round's half-up rule. Added 601×777, quarter-centre-crop → 2× regression.

I3 resolved: prepareCandidates validates and copies into one contiguous privately owned Int8Array (M×512), exposed through an immutable CandidateMatrix handle. WeakMap ownership prevents callers modifying the validated storage; invalid handles refuse. scoreCapture/detect accept that handle with O(1) lookup, keeping the old array signature as a wrapper that prepares each call. Tests mutate all caller arrays after preparation and check identical scores/attribution; repeated decisions never reread input codes. M=65,536 matrix owns 33,554,432 bytes and should be reused across decisions. Factory cost is paid once and storage is GC-released with the handle.

Minor resolved: removed the raw-reference-identical shortcut that tried to determine whether marking occurred. The shared pedestal is always subtracted on admitted support. For an unmarked page with spatially varying hypothetical pedestal, this can create structured residual and inflate sigma (or dilute genuine signal); the full null calibration must include it. Truly flat corrected residual still abstains. Added four odd-size crop/rescale captures per class with 1024 real BCH candidates, including white regions requiring pedestal, exercising registration rather than just same-size pages.

**Deviations:** the planned 5×5 box band-pass was dropped per the review-derived brief because filtering residual alone mismatches the unfiltered template; reference subtraction plus consistent pedestal subtraction/support centring is used instead. No unmatched band-pass is applied.

### Before / after evidence

On the committed test's 601×777 deterministic text page (different from the reviewer's fixture, so before Z differs from its 3.3/2.7):

| Attack order | Correct-candidate Z before | Correct-candidate Z after | Result |
|---|---:|---:|---|
| crop 25% centre → 0.5× | 1.8828767269 (best wrong score 3.2250240387) | 17.0399331581 | attributed 73 |
| crop 5% bottom-right → 0.5× | 2.9940391359 | 21.3486344327 | attributed 73 |
| crop 25% centre → 2× | no registration / zero support | 16.9404185353 | attributed 73 |

Registration-search nulls (4096 candidate scores per class): unmarked mean 0.0023574266, std **1.0006656168**; wrong-key mean 0.0042948453, std **1.0414416676**. Both std gates 0.8–1.25 pass. These are marginal-score checks, not Task 8's decision-maxima false-attribution calibration.

M=65,536 performance, Node v22.22.3, same-size text page, distinct real BCH candidates: one-time preparation 165.20 ms; 2 warmups + 7 measured decisions, **median 67.2155 ms**, min 66.9077 / max 67.4347 ms. Exactly 65,536 outer input reads during preparation and no further reads across all nine decisions. Shared-machine context only; no timing assertion. Registration-heavy crop decisions may cost more; no comparison to the reviewer's 350 ms is claimed as a controlled benchmark.

Reproduce:
`S0_MARK_FIX_RESULTS=after S0_MARK_FIX_BENCH=1 bunx vitest run --config vitest.stream-s0.config.ts tests/stream-s0/mark-*.test.ts`

```
Test Files  2 passed (2)
     Tests  21 passed (21)
  Duration  7.60s (transform 101ms, setup 0ms, import 321ms, tests 8.03s, environment 0ms)
```

The benchmark test is opt-in; the normal command runs 20 and skips that measurement. Focused TypeScript and task-owned diff whitespace checks pass. No Task 8 implementation, no new dependencies, no commits.
