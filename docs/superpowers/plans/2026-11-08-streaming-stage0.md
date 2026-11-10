# Server Streaming — Stage 0 Feasibility Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce every Stage-0 number §17 row 0 requires — published in `docs/validation/streaming.md` — before any wire, recorder or server code exists.

**Architecture:** Stage 0 is measurement, not product. All code lives in `scripts/stream-s0/` (probes, the Layer-1 mark codec prototype, the show-plan client prototype) with tests in `tests/stream-s0/`. Nothing under `src/` changes, so the shipped bundle and its eager gates are untouched. Every probe writes a JSON result to `docs/validation/streaming-s0/<task>.json`; one report script renders them into `docs/validation/streaming.md`.

**Tech Stack:** TypeScript, bun (scripts), vitest (tests), `canvas` 3.2.3 (current node-canvas backend), `@napi-rs/canvas` 1.1.0 (the chosen server backend, §2 item 1 — already installed as a devDependency), pako (gzip).

**Spec:** `docs/superpowers/specs/2026-11-07-server-streaming-design.md` (draft 11, approved). Sections cited as §N.

## Global Constraints

- **Gate failures are reported, never worked around** (§17 row 0): if a number misses its target, record it, mark the gate RED in the report, and stop that gate. Do not switch backend, loosen a ceiling, or change a spec value.
- No file under `src/` is modified. `bun run build` output and `bun scripts/check-eager.ts` must be unchanged by this branch.
- Server backend is `@napi-rs/canvas` (§2 item 1). It is reached in Stage 0 **only** through `scripts/stream-s0/napi-shim.ts` + a vitest alias — the `CanvasBackend` seam (§7.3) is Stage 1.
- Corpus: `corpus/` (49 files, gitignored; copy from the main checkout if absent). Never commit corpus files.
- Zoom × DPR grid (T1): zoom `{0.5, 0.75, 1, 1.25, 1.5, 2, 3}` × DPR `{1, 2}`.
- Mark payload: `codeword_40 = token_24 | salt_8 | CRC8_8`; `N_chips ≥ 512`; macroblock cells ≥ 2×2 device px (§13.2).
- ΔE00 ceiling target ≤ 3 [call] (§13.2).
- False attribution ≤ 10⁻⁶ per **complete decision** (alignment + crop ≤ 25 % + scale search + all M candidates) (§13.2, §13.5).
- Declared audit-table size for Stage 0: **M = 65,536** [Stage-0 call]; report how τ scales to M = 2²⁴.
- Survival targets, image-only units (§13.5): default screenshot ≥ 90 %; JPEG q70 ≥ 75 %; 50 % downscale ≥ 75 %; q70 then 50 % downscale ≥ 75 %; 25 % crop ≥ 70 %.
- §14 caps for T12: per-session 8 outstanding units / queue 32; session transient 96 MiB; render reservation 512 MiB per process; process 256 sockets, 4,096 queued units; render deadline 2 s.
- T8 target: tile p95 ≤ 120 ms on a 500×100 sheet.
- T11 ceiling: transitive eager-graph gzip ≤ 38,720 B (§8.1).
- Results JSON always carries `{ gate, status: 'GREEN'|'RED'|'REPORTED', backend, gitSha, generatedAt, ...data }`. `REPORTED` = measurement with no pass/fail target at Stage 0.

## Stage-0 scope calls (made here; surfaced in the report)

1. **T11** — Stage 0 runs the §8.1 gate *scripts* against the current build and publishes the numbers plus a module-map baseline hash. Building `officeview/core` and its purity assert are Stage 1.
2. **T12** — there is no server at Stage 0, so T12 is a **synthetic admission probe**: an in-process model of the pool's reservation/queue rules driven by 1,000 concurrent requests with real `@napi-rs/canvas` paints. It proves the caps are enforceable and measures peak RSS; it does not test a server.
3. **T2** — Stage-0 subset = hidden-content fixture matrix + paint-leak check + model-exposure inventory (what content policy must strip). The Playwright WS harness and "local loads byte-for-byte unchanged" are Stage 1.
4. **T6** has no Stage-0 subset (wire is Stage 1). **T8** Stage-0 subset = the tile p95 timing under `@napi-rs/canvas`.
5. **T1** excludes "fractional zoom proves 1:1 replay" (needs the recorder, Stage 1). Stage-0 T1 = backend confirm: render success, construction-site matrix, cross-backend diff table.
6. **T7/T9** cover Layer 1 only. Layer 2 (glyph-advance) is command projection, Stage 2+ (§20).

## Review Focus

- A corpus unit that throws or returns an empty bitmap under `@napi-rs/canvas` but not node-canvas — must be counted as a render failure, never skipped (Task 2 test `counts napi-only failures`).
- Mostly-white office pages clip positive chips at 255 — the embed must keep measurable chip power on a pure `#FFFFFF` page (Task 6 test `white page keeps chip power`).
- A flat / reference-identical capture gives σ_r = 0 — detector must return `abstain`, never a score or NaN (Task 7 test `flat capture abstains`).
- Captures smaller than the reference after a crop + downscale — registration must not index out of bounds (Task 7 test `25% crop of 50% downscale registers`).
- Synthetic admission under bursty arrivals — a request arriving when the queue is exactly full must be refused, not silently queued (Task 10 test `33rd request refused`).

---

### Task 1: Shared harness — corpus units, napi shim, results writer

**Files:**
- Create: `scripts/stream-s0/napi-shim.ts`, `scripts/stream-s0/corpus.ts`, `scripts/stream-s0/results.ts`, `vitest.stream-s0.config.ts`, `vitest.stream-s0-napi.config.ts`
- Test: `tests/stream-s0/harness.test.ts`
- Modify: `package.json` scripts only (`"s0:test"`, `"s0:test:napi"`, `"s0:report"`); `.gitignore` (add `.s0-cache/`)

**Interfaces:**
- Produces:
  - `napi-shim.ts`: re-exports with node-canvas names — `createCanvas(w,h)`, `loadImage(src)`, `Image`, `registerFont(path, {family, weight?, style?})` (→ `GlobalFonts.registerFromPath`), `Canvas`.
  - `corpus.ts`: `listCorpus(): CorpusFile[]` (`{ name, path, kind: 'docx'|'xlsx'|'pptx' }`); `loadUnits(file): Promise<Unit[]>` where `Unit = { index, widthPx, heightPx, paint(ctx) }` via `loadOfficeFile → getPaintables`; `renderUnit(unit, scale): Promise<Bitmap>` (`Bitmap = { width, height, data: Uint8ClampedArray }`), painting at `scale = zoom × dpr` exactly like `src/test/pixel-diff.ts:98` but without the 8192 clamp silently: a unit above 8192 px on a side returns `{ skipped: 'oversize' }`.
  - `results.ts`: `writeResult(task: string, body: ResultBody): void` → `docs/validation/streaming-s0/<task>.json`; `ResultBody` per Global Constraints.
- `vitest.stream-s0-napi.config.ts` = the s0 config plus `resolve.alias: { canvas: <abs path to napi-shim.ts> }`.

- [ ] Step 1: tests — `shim exposes node-canvas names` (each export is a function/class; `createCanvas(4,4).getContext('2d').getImageData(0,0,4,4).data.length === 64`); `listCorpus finds 49 files` (skip with message when `corpus/` empty); `renderUnit scales` (a known docx unit at scale 2 has exactly 2× width/height of scale 1, ±1 px); `writeResult writes status+gitSha`.
- [ ] Step 2: run `bun run s0:test` and `bun run s0:test:napi` → FAIL (modules missing).
- [ ] Step 3: implement the three modules and two configs.
- [ ] Step 4: both commands PASS. Also run `bun run build && bun scripts/check-eager.ts` → unchanged output.
- [ ] Step 5: commit `test(stream-s0): harness, napi shim, results writer`.

### Task 2: T1 backend confirm

**Files:** Create `scripts/stream-s0/t1-backend.ts`; Test `tests/stream-s0/t1-backend.test.ts`

**Interfaces:** Consumes Task 1. Produces `docs/validation/streaming-s0/t1.json` with:
- `renderSuccess: { total, okNode, okNapi, napiOnlyFailures: [{file, unit, error}] }`
- `constructionSites: [{ site: 'src/render/paint.ts:199' | … (the 9 sites in §7.3), napi: 'ok'|'throws'|'undefined', evidence }]`
- `diff: [{ file, unit, zoom, dpr, pixelsDiffering, fracDiffering, maxChannelDelta, p99ChannelDelta }]` — "differing" = any channel |Δ| > 2.
- `tolerance: { byKind: {docx,xlsx,pptx}: { p95FracDiffering, maxChannelDelta } }`

Grid: every unit at (zoom 1, DPR 1) and (zoom 1, DPR 2); the **full 14-cell grid** on the painter-class subset — units whose T3 census (or, if Task 3 is not done yet, a grep of the corpus run) uses `shadowBlur>0`, `setTransform`, `createPattern`, or `drawImage` from a canvas source; at least one unit per class.

Construction sites: exercise each site under the napi config with a fixture that reaches it (gradient/warp/pattern fill, text metrics, svg raster, image decode) and record whether the result is a real canvas, `undefined` (silently dropped fill) or a throw.

- [ ] Step 1: tests — `counts napi-only failures` (inject a unit whose paint throws only when `ctx.constructor.name` is the napi context → appears in `napiOnlyFailures`); `diff of identical bitmaps is zero`; `site matrix has 9 rows`.
- [ ] Step 2: run → FAIL. Step 3: implement (run node-canvas renders under the plain config and napi renders under the napi config, each writing PNGs to `.s0-cache/` (gitignored); the diff step reads both).
- [ ] Step 4: tests PASS; run `bun scripts/stream-s0/t1-backend.ts` → writes t1.json. GREEN iff `napiOnlyFailures` empty **and** all 9 sites `ok`; otherwise RED with the list.
- [ ] Step 5: commit `test(stream-s0): T1 backend confirm`.

### Task 3: T3 recorder census

**Files:** Create `scripts/stream-s0/census-proxy.ts`, `scripts/stream-s0/t3-census.ts`; Test `tests/stream-s0/t3-census.test.ts`

**Interfaces:** Produces `wrapCensus(ctx): { ctx, report(): CensusReport }` — a `Proxy` over a real 2D context that forwards everything and counts method calls (by name + arg-shape, e.g. `drawImage/9`, `drawImage/source=canvas`) and property writes (by name + value class). `t3.json`: `{ methods: {name: count}, props: {name: {count, values: string[≤20]}}, notInSubset: string[], faceResolution: [...] }`.

`notInSubset` = every method/property used that §7.3's recorded subset / §7.5's op list does not cover (encode the subset as a constant `RECORDED_SUBSET` in `census-proxy.ts`, copied from §7.3). `RECORD_TEXT` access must be forwarded but not counted as a canvas op.

Face resolution: for each distinct `ctx.font` string seen, `measureText('Hxg 漢字 مرحبا').width` under node-canvas and napi, and the napi resolved family if the API exposes it; record `{font, widthNode, widthNapi, relDiff}`.

- [ ] Step 1: tests — `proxy forwards and counts` (fillRect twice → count 2, pixels painted); `unknown op lands in notInSubset` (call `ctx.filter = 'blur(1px)'` → listed); `instanceof and 'letterSpacing' in ctx behave as on the real ctx`.
- [ ] Step 2 FAIL → Step 3 implement → Step 4 PASS; run census over all corpus units at zoom 1 DPR 1 under napi. GREEN iff `notInSubset` empty; else RED listing the ops (these are spec gaps for the owner).
- [ ] Step 5: commit `test(stream-s0): T3 recorder census`.

### Task 4: T4 font probes

**Files:** Create `scripts/stream-s0/t4-fonts.ts`; Test `tests/stream-s0/t4-fonts.test.ts`

**Interfaces:** Produces `requiredFaces(file): Promise<Face[]>` where `Face = { id, family, source: 'embedded'|'bundled'|'system', bytes?: Uint8Array, hash96?: string }` (`hash96` = first 12 bytes of SHA-256 of face bytes, hex), derived from what `acquireFonts` (`src/core/fonts/register.ts:40`) registers plus the families the census saw. `t4.json`: per file `{ faces, verifiable, unverifiable, unitsThatWouldRaster }` and corpus totals; plus a fixture probe for CJK, ligature (`ffi`), and BiDi (Arabic + Hebrew) text comparing napi vs node-canvas advances.

A unit "would raster" when it uses any face with no bytes (`system`) — §7.4 rule 3.

- [ ] Step 1: tests — `hash96 is 24 hex chars and stable`; `system face is unverifiable`; `embedded font in a fixture docx is verifiable` (use an existing font fixture from `tests/fixtures/fonts/`).
- [ ] Step 2 FAIL → 3 implement → 4 PASS; run → t4.json, status REPORTED (the raster-fallback rate is the number the owner needs).
- [ ] Step 5: commit `test(stream-s0): T4 font probes`.

### Task 5: T2 hidden-content fixture matrix (Stage-0 subset)

**Files:** Create `scripts/stream-s0/hidden-fixtures.ts`, `scripts/stream-s0/t2-hidden.ts`; Test `tests/stream-s0/t2-hidden.test.ts`

**Interfaces:** `hiddenFixtures(): Fixture[]`, each `{ id, kind, withHidden: Uint8Array, withoutHidden: Uint8Array, marker: string }` built with `tests/helpers/zip-builder.ts`. Fixture ids (the §16 T2 matrix): `xlsx-hidden-sheet`, `xlsx-veryhidden-sheet`, `xlsx-hidden-rows`, `xlsx-hidden-cols`, `docx-vanish`, `docx-hidden-shape-docPr`, `docx-hidden-shape-in-group`, `pptx-hidden-shape-cNvPr`, `pptx-hidden-shape-in-group`, `pptx-hidden-slide`, `pptx-notes`, `xlsx-comment`, `docx-comment`. Each hidden payload contains a unique marker string.

`t2.json` per fixture: `{ paintLeak: boolean, modelExposure: string[] }` — `paintLeak` = rendered pixels of `withHidden` ≠ `withoutHidden` at zoom 1 (napi); `modelExposure` = JSON paths in the parsed document model where `marker` appears (what §6.4 content policy must strip before any recorder sees the model).

- [ ] Step 1: tests — `every fixture parses`; `marker found in model for docx-vanish` (proves the exposure walk works); `identical docs have no paint leak`.
- [ ] Step 2 FAIL → 3 implement → 4 PASS; run → t2.json. GREEN iff no `paintLeak` (exposure is REPORTED — it is the strip list).
- [ ] Step 5: commit `test(stream-s0): T2 hidden-content matrix`.

### Task 6: T7 Layer-1 mark codec + ΔE00

**Files:** Create `scripts/stream-s0/mark/ciede2000.ts`, `mark/bch.ts`, `mark/basis.ts`, `mark/embed.ts`, `mark/params.ts`; Test `tests/stream-s0/mark-embed.test.ts`

**Interfaces:**
- `deltaE00(lab1: [L,a,b], lab2): number`; `srgbToLab(r,g,b): [L,a,b]`; `pageDeltaE(a: Bitmap, b: Bitmap): { mean, p95, max }`.
- `bchEncode(msg40: Uint8Array /*40 bits*/): Int8Array /*511 ±1*/` for BCH(511, 40) (narrow-sense, primitive poly `x⁹+x⁴+1`); chip 512 = parity of the codeword.
- `codeword40(token24: number, salt8: number): Uint8Array` with CRC-8 (poly `0x07`) over token‖salt.
- `basis(kBasis: Uint8Array, unitW, unitH, params): Chip[]` — `Chip = { x, y, cell, sign0 }`: HKDF-SHA256(`kMaster`, `"basis" | docRef | unitId`) seeds a ChaCha20/xorshift128+ PRNG that places `nChips` chips on the macroblock grid (cell = `params.cellPx`, default 4 px at DPR 1) inside the low-mid band (each chip a `params.cellPx × params.cellPx` ±1 block, tiled every `params.tilePx`).
- `embed(bmp: Bitmap, chips: Chip[], code: Int8Array, params): Bitmap` — adds `±params.amplitude` to luminance (equal RGB delta) per chip; **white pedestal**: on a cell whose mean is ≥ 254, first shift to `255 − params.pedestal` (default 2) so positive chips survive; record which variant was used.
- `params.ts`: `DEFAULT_PARAMS = { nChips: 512, cellPx: 4, tilePx: 64, amplitude: 2, pedestal: 2 }` and `sweepParams()` yielding the grid {amplitude 1–4, cellPx 2/4/6, pedestal 0/2} for Task 8.

- [ ] Step 1: tests — `ciede2000 matches Sharma pairs` (Sharma–Wu–Dalal 2005 pairs 1, 7, 17: expected 2.0425, 2.3669, 27.1492, ±1e-4); `bch codeword has min distance ≥ 191` (random 1,000 pairs, Hamming ≥ 191); `crc8 detects single-bit flips`; `white page keeps chip power` (embed on an all-255 bitmap: mean |Δ| over chips ≥ 0.8 × amplitude); `embed is deterministic per key/unit`.
- [ ] Step 2 FAIL → 3 implement → 4 PASS.
- [ ] Step 5: commit `feat(stream-s0): Layer-1 mark codec prototype`.

### Task 7: T9 detector

**Files:** Create `scripts/stream-s0/mark/detect.ts`, `mark/attacks.ts`; Test `tests/stream-s0/mark-detect.test.ts`

**Interfaces:**
- `attacks.ts`: `jpeg(bmp, q)`, `scale(bmp, f)` (napi `drawImage` with `imageSmoothingQuality='high'`), `crop(bmp, frac, anchor)`, `compose(...fns)`; `average(bmps)`, `median(bmps)`, `difference(a,b)`.
- `detect(capture: Bitmap, reference: Bitmap, chips, candidates: Int8Array[] /*M codes*/, params, tau): Decision` where `Decision = { kind: 'attributed', index, z, margin } | { kind: 'none', zBest } | { kind: 'abstain', reason: 'flat' }`.
- Steps inside `detect`: (1) registration: search scale ∈ {1, 0.5, 2} × translation (phase correlation on luminance) to map capture into reference coordinates, tolerating ≤ 25 % crop; (2) residual `r = lum(capture') − lum(reference)` band-passed (subtract 5×5 box mean); (3) per chip soft value = mean residual over its surviving cell (cells outside the capture or clipped at 0/255 excluded → support L); (4) σ_r over the L surviving chip values; σ_r = 0 ⇒ abstain; (5) `Z_m = ⟨r, c_m⟩ / (σ_r √L)` for all M candidates — one matrix-vector pass, O(M·L); (6) attribute iff `Z_best ≥ tau.abs && Z_best − Z_second ≥ tau.margin`.

- [ ] Step 1: tests — `flat capture abstains`; `clean marked page attributes the right token among M=1024`; `unmarked page → none`; `25% crop of 50% downscale registers` (no throw; correct token on a text page); `jpeg q70 still attributes` on one fixture page.
- [ ] Step 2 FAIL → 3 implement → 4 PASS.
- [ ] Step 5: commit `feat(stream-s0): Layer-1 detector prototype`.

### Task 8: T7/T9 batteries, calibration, collusion

**Files:** Create `scripts/stream-s0/t7-t9-battery.ts`; Test `tests/stream-s0/battery.test.ts`

**Interfaces:** Consumes Tasks 6–7. Produces `t7.json` and `t9.json`.
- **Page set:** 40 corpus units at zoom 1 DPR 1 (napi): ≥ 15 docx text pages, ≥ 10 xlsx sheets, ≥ 10 pptx slides, ≥ 5 image-heavy; record the list.
- **T7 (`t7.json`):** for each `sweepParams()` setting: `pageDeltaE` mean/p95/max over the page set, chip carrier count per page (chips with non-zero support after pedestal/zeroing — the "per-page carrier-count publication"), pedestal variant used. **Choose** the setting with the best q70+½ survival whose page ΔE00 **max** ≤ 3 (chips cover < 1 % of a real page, so whole-page p95 is ~0 and would always pass; also publish p95 over chip pixels only); publish it as the codec parameters. GREEN iff such a setting exists.
- **T9 (`t9.json`):**
  - *Null calibration:* score unmarked pages and wrong-key marked pages against M = 65,536 random candidates through the **full** detect (registration search included) — ≥ 2,000 decisions → ≥ 1.3 × 10⁸ candidate scores. Fit the upper tail of `Z_best` (empirical CDF + Gaussian/GEV tail fit); set `tau.abs` so the fitted P(false attribution) ≤ 10⁻⁶ per decision, `tau.margin` from the null `Z_best − Z_second` distribution likewise. Publish the fitted tail, the empirical false-attribution count at the chosen τ (must be 0 of ≥ 2,000), and τ for M = 2²⁴ by the union-bound scaling of the fit.
  - *Survival* (chosen params, chosen τ): rate of correct attribution per attack row of §13.5 (default, q95/90/80/70, 2× up, 50 % down, q70→50 %, 5 % crop, 25 % crop) over the page set × 5 tokens each. GREEN iff every §13.5 image-only target is met.
  - *Collusion* (REPORTED): K-average K ∈ {2,4,6}; K-sample median K ∈ {4,5,6}; two-stream differencing — recoverability = fraction of colluders' tokens attributed (any of them) and fraction of false attributions.
- [ ] Step 1: tests — `battery runs on a 2-page set with M=256` (small-size smoke; all rows present in output); `tau chosen meets fitted 1e-6`.
- [ ] Step 2 FAIL → 3 implement → 4 PASS; run full battery (long — run in the background, log progress); write both JSONs.
- [ ] Step 5: commit `test(stream-s0): T7/T9 batteries and calibration`.

### Task 9: T5 show-plan client prototype + appearance-eligibility rate

**Files:** Create `scripts/stream-s0/showplan/serialize.ts`, `showplan/client.ts`, `scripts/stream-s0/t5-showplan.ts`; Test `tests/stream-s0/t5-showplan.test.ts`

**Interfaces:**
- `serializePlan(slide, plan /* from src/slideshow/render/plan.ts */): { plan: OvspJson, assets: Map<string, Bitmap> }` — layers rastered via `rasterLayer` (`src/slideshow/render/raster.ts:41`), JSON per §12a.2 (`layers[]: {id, assetId, bounds, z, partTag}`, `steps[]`), **no `PptxShape` data**.
- `client.ts`: `composeFrame(plan: OvspJson, assets, stepIndex, t): Bitmap` using only the engine/compositor pieces (`src/slideshow/render/canvas2d.ts`, transitions, masks). `client.ts` must not import anything under `src/pptx/` or the slideshow `read/` package readers.
- `classifyStep(step): 'proof-subset' | 'step-composited'` per §12a.3 rule 4 (visibility toggles / transitions over safe rasters; motion/Morph only with an analytic swept-area proof — at Stage 0, motion/Morph count as `step-composited` unless the mover's swept rect intersects no masked/withheld pixel).
- `t5.json`: per slideshow corpus deck: steps total, proof-subset count, step-composited count (the **appearance-eligibility rate**), and per-step final-frame pixel diff `composeFrame` vs the local player's frame (napi).
- [ ] Step 1: tests — `client.ts imports nothing from src/pptx` (static import-graph walk with `tests/helpers/eager-graph.ts` helpers or a regex walk of transitive imports); `static slide composes identical to local` (≤ 0.1 % pixels differ); `opacity-0 layer is classified step-composited or withheld`.
- [ ] Step 2 FAIL → 3 implement → 4 PASS; run over the pptx corpus → t5.json. GREEN iff the import test passes and static-slide diffs are within tolerance; the eligibility rate is REPORTED.
- [ ] Step 5: commit `test(stream-s0): T5 show-plan prototype`.

### Task 10: T12 synthetic admission probe + T8 tile timing + T11 gates

**Files:** Create `scripts/stream-s0/t12-admission.ts`, `scripts/stream-s0/t8-tiles.ts`, `scripts/stream-s0/t11-eager.ts`; Test `tests/stream-s0/t12-admission.test.ts`

**Interfaces:**
- `class AdmissionModel { constructor(caps: Caps); request(sessionId, bytes): 'admitted'|'queued'|'refused'; complete(id): void; peak(): {reserved, queued, outstandingMax} }` with `Caps` = §14 values in Global Constraints. Driver: 1,000 concurrent requests over 40 sessions + 5 rogue sessions bursting 100 each; admitted jobs run a real napi paint of a corpus unit with a 2 s deadline in `worker_threads` (4 workers). `t12.json`: peak reserved vs 512 MiB, per-session max outstanding (≤ 8), queue max (≤ 32), refusals, peak RSS, p50/p95 latency. GREEN iff no cap exceeded.
- `t8-tiles.ts`: build a 500×100 sheet with `tests/helpers/xlsx-grid-fixture.ts`, render viewport tiles through `src/xlsx/view/renderViewport.ts` under napi, 200 tiles, `t8.json` with p50/p95/max ms. GREEN iff p95 ≤ 120 ms.
- `t11-eager.ts`: run `OV_MODULE_MAP=1 bun run build`, compute the transitive eager-graph gzip sum (closure from `tests/helpers/eager-graph.ts`), and the sha256 of the `index.js` module-map entry; `t11.json` `{ transitiveGzipBytes, ceiling: 38720, moduleMapHash }`. GREEN iff ≤ ceiling. (Core purity: `"stage1"`.)
- [ ] Step 1: tests — `33rd request refused` (one session, 8 outstanding + 32 queued, next refused); `reservation released on complete`; `deadline exceeded job refunds only after drain`.
- [ ] Step 2 FAIL → 3 implement → 4 PASS; run all three scripts → JSONs.
- [ ] Step 5: commit `test(stream-s0): T12 admission, T8 tiles, T11 eager`.

### Task 11: Report

**Files:** Create `scripts/stream-s0/report.ts`; Create/overwrite `docs/validation/streaming.md`

**Interfaces:** Reads every `docs/validation/streaming-s0/*.json`; writes `streaming.md` with: a gate summary table (gate, status, headline number, target), the six Stage-0 scope calls verbatim from this plan, the chosen codec parameters, τ and tail fit, the survival table in §13.5's shape (measured vs target), collusion results, T1 tolerance table, T3 `notInSubset`, T4 raster-fallback rate, T2 strip list, T5 eligibility rate, T8/T11/T12 numbers, and a **"RED gates — owner decision needed"** section listing every RED with its number. A missing JSON is listed as `NOT RUN`, never omitted.

- [ ] Step 1: test — `report lists NOT RUN for a missing json` and `RED gates appear in the owner section`.
- [ ] Step 2 FAIL → 3 implement → 4 PASS; `bun run s0:report`.
- [ ] Step 5: commit `docs(stream-s0): Stage-0 validation report`.

## Execution order

1 → (2, 3, 4, 5 in any order; 2 may wait on 3 for the painter-class subset) → 6 → 7 → 8 → 9 → 10 → 11. Tasks 6–8 are the critical path (the mark math); assign them to the strongest implementer.
