# Task 6 — Layer-1 codec and ΔE00

Status: completed Task 6 steps 1–4; uncommitted as instructed. No commit SHA for this work. Measurement base HEAD: `f739bd1d901d5446fe7ea469c7cd1f6494bdc23c` (the concurrent harness agent may advance it). Task 7/8 detector and attack batteries are not started. `task6.json` is a codec-subset GREEN, not full T7/T9 approval.

## Exact files owned by this task

- scripts/stream-s0/mark/types.ts
- scripts/stream-s0/mark/params.ts
- scripts/stream-s0/mark/bch.ts
- scripts/stream-s0/mark/basis.ts
- scripts/stream-s0/mark/embed.ts
- scripts/stream-s0/mark/ciede2000.ts
- tests/stream-s0/mark-embed.test.ts
- docs/validation/streaming-s0/task6.json
- docs/validation/streaming-s0/task6.md

No source, dependency, configuration, corpus, harness, or other task files edited. No network or new dependencies. No commit or push.

## RED / GREEN and measurements

Initial required-test RED: one failed suite, missing `mark/ciede2000` module. Implemented all five required tests/functions; a syntax error in the exponent expression was corrected before tests ran. Initial GREEN: 1 file / 5 tests passed.

Additional genuine runtime RED: Bun 1.3.10 throws `ERR_CRYPTO_UNKNOWN_CIPHER: chacha20` despite Node's Vitest passing. Added `basis runs under the Bun probe runtime as well as Node`, observed 1 failed / 5 passed, then used the plan-permitted xorshift128+ implementation.

Final `bunx vitest run tests/stream-s0/mark-embed.test.ts`:

```
Test Files  1 passed (1)
     Tests  6 passed (6)
  Duration  708ms (transform 179ms, setup 152ms, import 45ms, tests 436ms, environment 0ms)
```

Focused type check passed, exit 0:
`bunx tsc --noEmit --target ES2022 --module ESNext --moduleResolution bundler --types node --skipLibCheck scripts/stream-s0/mark/*.ts tests/stream-s0/mark-embed.test.ts`

`git diff --check`: passed. No full product build claimed; no product code/import graph changed.

Sharma–Wu–Dalal pairs 1, 7, 17: 2.0424596801565738, 2.3668588191717523, 27.14923130074626, each within 1e-4 of the plan's published values. References were the offline values in the brief/plan, not a network lookup.

BCH: 1,000 seeded distinct random 40-bit message pairs, minimum observed distance **219 >= 191**. Also checked all 40 systematic unit-message weights >=191, zero word and even-parity extension. The empirical minimum is not a proof of the global minimum; generator construction has designed distance 191 (roots alpha^1 through alpha^190, cyclotomic closure degree 471, dimension 40).

White 256×256 page, defaults: mean absolute RGB chip delta **1.9921875 >= 1.6**, ΔE00 mean **0.04953837075989174**, p95/max **0.7957222220882249 <= 3**. All 512 cells used the pedestal, no masking. Corpus appearance and survival remain unmeasured.

## Choices left open by the plan / detector handoff

1. Local Bitmap is straight-alpha, byte sRGB RGBA. Embedding returns a new Bitmap plus `mark` metadata (variant, pedestalCells, maskedCells, parameter snapshot); input remains unchanged.
2. CRC-8/SMBUS: poly 0x07, init/xorout 0, non-reflected. Token's three bytes are big-endian, followed by salt, then CRC. `codeword40` exposes MSB-first bits. Check vector `123456789` = 0xF4; every single bit of the full 40-bit packet is detected by a nonzero CRC residue.
3. Primitive GF(512) polynomial 0x211 (`x^9+x^4+1`), roots 1–190 plus binary conjugates. Generator coefficients and output codeword are low-degree first. Systematic output has 471 parity bits then reversed input bits; bit 0 maps to sign +1, bit 1 to -1. `bchEncode` returns exactly 511 signs; `extendedCodeword` adds the parity sign as symbol 512 (product of all signs is +1).
4. `deriveBasisKey(master, docRef, unitId)` uses HKDF-SHA256 with empty salt and UTF-8 JSON tuple `["basis",docRef,unitId]` as unambiguous info; master must be 32 bytes. `basis` takes this already-derived 32-byte key, not the master. Neither path includes a session token, so carrier support is page-keyed.
5. PRNG: xorshift128+ (plan's allowed alternative), shifts 23/17/26, explicit unsigned 64-bit arithmetic. SHA-256 of all kBasis bytes seeds the first two LE uint64 words; all-zero state falls back to (0,1). Each draw returns the high 32 bits of the sum. Bounded selection uses rejection sampling. This is a feasibility PRNG, **not a cryptographic carrier-secrecy claim**. Bun lacks the attempted OpenSSL ChaCha20 cipher, pinned by a runtime test.
6. Carrier positions are full nonoverlapping `cellPx` blocks on a macroblock grid whose origin is (0,0), tiled every `tilePx`; leftover strips at tile/page edges are unused. Select nChips without replacement over the available grid using sparse Fisher–Yates, O(nChips) position state. Carrier signs consume separate PRNG draws. Default 4×4 blocks implement the specified block carrier; this is not a spectral filter or proof of JPEG/downscale survival. A tiny page with <nChips full cells refuses; no silent chip/cell reduction. For the 6px sweep, 64px tiles leave a 4px strip. nChips >512 cycles the extended code, recorded as a prototype convention; default and tested coding use exactly 512.
7. Embed multiplies carrier sign0 by the code sign and applies equal RGB delta. At cell mean RGB >=254 and pedestal>0, an equal-channel offset moves the mean to 255−pedestal before modulation. Uint8ClampedArray clamps/rounds the result. pedestal=0 uses direct modulation. A cell containing any nonopaque pixel is masked entirely; alpha is preserved. The variant/counts are recorded. Overlapping/out-of-bounds chips and malformed code/parameters refuse rather than partially painting.
8. Colour conversion uses the IEC sRGB transfer and D65 XYZ reference white (95.047,100,108.883); CIEDE2000 uses kL=kC=kH=1. pageDeltaE includes all pixels after white-background alpha composition, reports arithmetic mean, exact nearest-rank p95 and max. It uses O(pixels) measurement storage and O(pixels log pixels) sorting; this is an offline Stage-0 measurement, not a production admission claim.
9. sweepParams yields exactly 24 combinations, amplitude outermost 1/2/3/4, cellPx 2/4/6, pedestal 0/2; other defaults unchanged. No sweep outcome, detector calibration, false-attribution bound or attack-survival rate is asserted by Task 6.

## Review fixes (supersedes the original distance/pedestal checks)

See `task7.md` Part A for the full RED/GREEN record. The random-distance test is replaced by independent root-membership and linearity tests; all 34 Sharma pairs are checked. Pedestal now reserves both signs against every channel's extrema or masks the chip, with shared `pedestalFor`, per-chip actual luminance deltas and applied pedestal metadata. The white gate measures centred sign-correlated modulation, not absolute/DC power. pageDeltaE skips identical pixels and reports chipPixelsP95; task6.json is regenerated without the uncommitted field. These changes supersede original choices 7/8 where they differ; BCH encoding/key derivation remain unchanged.
