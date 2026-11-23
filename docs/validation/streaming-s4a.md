# Server streaming — Stage 4a validation (server-painted XLSX sheet tiles)

Spec: `docs/superpowers/specs/2026-11-22-streaming-s4-design.md` (§9 4a; "As built (4a)" under §11.2). Plan: `docs/superpowers/plans/2026-11-22-streaming-s4a.md` (rulings R1–R16). Branch `feature/streaming-s4a`, off `docs/streaming-s4-design` at 2de7ae4a (main at the design: 53668e81). Measured 2026-11-23 on macOS, darwin-arm64, 14 cores, Bun 1.4.5.

## Status

Every 4a gate is green except three that were already red before 4a:
- the eager size gate (index closure), red at the base 2de7ae4a from the xlsx-cf work; 4a adds 45 B and no module to it (§Size gate);
- one vitest case, `tests/slideshow-screens.test.ts` "the drag hint works on a popup (other-realm) document and stays dismissed", which fails the same way at main 53668e81 (Node 26.13.1 exposes no `localStorage` without `--localstorage-file`); 4a touches neither the test nor `src/slideshow`;
- two manual-device collector tests of the unmodified `tests/browser/interactive-viewer.spec.ts`, which fail at the default ports because the stream harness holds 4181–4183, and fail the same way at 53668e81; all 25 pass with `OV_STREAM_PORT=4281` (§Known flakes).

T8 tile p95 is GREEN: 83.6 ms at dpr 1 and 56.0 ms at dpr 2 against 120 ms. The compatibility matrix passes all four legs.

## S4 §9 4a "done when"

| Line | Proven by |
|---|---|
| T8 tile pipeline: no too-large-sheet notice on any corpus sheet | `server/test/sheet-corpus.test.ts` "every corpus sheet tiles without a static unit": every corpus XLSX under `visible-only` and `'full'` opens in a tiles-mode host; every sheet's layout round-trips the wire codec and tile (0,0) at zoomStep 5, dpr 1 renders 512²; `paintablesBuilt === 0`, so neither `sheetPaintable` nor the too-large refusal paintable ran (33 files, 61 opens, 94 sheets). Three files are refused at open in both modes (below). Browser: `tests/browser/stream-sheet.spec.ts` "no too-large notice" |
| T8: a million-row fixture scrolls to its last row | `tests/browser/stream-sheet.spec.ts` "million-row sheet scrolls to row 1,048,576"; `server/test/sheet-tile.test.ts` "the million-row last row renders"; Task 1 leg (d) (open 196 ms, heap +6.45 MiB) |
| tile p95 ≤ 120 ms on the 500×100 fixture, measured as §6.8 states | `bun scripts/stream-s4/tile-p95.ts` (R3), tables below: GREEN |
| `stats.basePaints === 0` and no `renderSheetViewport` call on the client (spy) | `tests/xlsx-stream-tiles.test.ts` "a missing tile is requested, not painted", "an arriving tile is cached and a repaint scheduled", "no watermark on a stream frame"; `tests/xlsx-stream-viewer.test.tsx` (every frame `basePaints === 0`); `tests/browser/stream-sheet.spec.ts` "no local painting" |
| the wire audit for 4a: every byte server→client captured and decoded for a hostile fixture set; no sentinel, formula, hidden row/column/sheet string or drawing byte | `server/test/sheet-wire-audit.test.ts` "the audit is live: …", "no sentinel on the wire", "no drawing byte on the wire", "no hidden-fixture marker, control or hidden sheet name on the wire"; `server/test/sheet-layout.test.ts` "layout keys are on the allowlist", "hidden sizes never travel" |
| tile pixels equal those of the same fixture with the hidden constructs deleted, before the mark (parent T2 on tiles) | `server/test/sheet-t2.test.ts` "every withHidden tile equals the withoutHidden tile (each XLSX hidden fixture)", "base-covered paints empty under visible-only", "rotated box", "saved values" |
| the compatibility matrix of §6.10 passes | `bun scripts/stream-s4/compat-matrix.ts --base 53668e81` (output below), plus the Stage 3 XLSX suites unchanged |
| the per-session view-state key is in the cache key; same clock shares canonical tiles while marked bytes differ; another day or zone does not share | `server/test/sheet-tiles-wire.test.ts` "same clock shares canonical tiles; marked bytes differ", "a different day or zone never shares"; `server/test/cache.test.ts` (tile-key variants, `viewStateHash`, the unit-key golden) |
| opening an XLSX in tiles mode renders no unit until a print request | `server/test/sheet-handshake.test.ts` "tiles open renders no unit until print"; `server/test/sheet-open.test.ts` "tiles open builds no paintables and renders nothing", "a units path builds paintables once" |
| the visible mark is continuous across a seam | `server/test/sheet-tile.test.ts` "the visible mark is continuous across a seam"; `server/test/sheet-t2.test.ts` "wire tiles keep the mark continuous" |
| Layer-1 detects on a one-pane capture, and on a blank one (owner ruling 2026-11-23: never names a wrong session) | `server/test/trace-tiles.test.ts` "a one-pane capture traces to its sid", "a marked blank pane never names a wrong session"; `docs/validation/streaming-s4a-feasibility.md` leg (c): lossless, JPEG q70 and 50 % downscale 30/30 each, 0 wrong; blank pane 30/30 to the right session, 0 wrong |
| T11 gates green | `bun run gates:core`: core closure, engine modules and core purity ok; the index closure is red as at the base (below). The stream sheet viewer is lazy from `RemoteDocView` and reuses the `InteractiveSheetDoc` chunk (`RemoteSheetView` chunk, Task 14) |

**Corpus files refused at open in both modes.** `lo-calc-PivotTableFormatsSubTotals.xlsx` under `visible-only` (`unowned visibility flag definedName hidden=1`: a hidden defined name fails closed, parent §6.4 last row; it opens and tiles under `'full'`), and `poi-59021.xlsx` and `poi-59746_NoRowNums.xlsx` under both policies (`unsupported-format` at admission). Units mode refuses each the same way, so these are not tile-path failures; the test asserts that.

## Latency

### Task 1 (feasibility, in process, no socket)

From `docs/validation/streaming-s4a-feasibility.md`, run 3 (1-minute load 4.98 at start), per-tile total in ms:

| Stage | dpr 1 p50 | dpr 1 p95 | dpr 1 max | dpr 2 p50 | dpr 2 p95 | dpr 2 max |
|---|---|---|---|---|---|---|
| render | 36.45 | 53.33 | 76.19 | 17.79 | 23.81 | 30.08 |
| visible mark | 0.15 | 0.21 | 4.14 | 0.13 | 0.18 | 0.30 |
| PNG encode | 9.02 | 9.50 | 10.84 | 8.72 | 9.16 | 15.11 |
| markRaster (per-tile) | 22.98 | 25.91 | 30.92 | 22.81 | 25.03 | 31.44 |
| gzip framing | 1.45 | 1.57 | 4.19 | 1.31 | 1.43 | 4.65 |
| **total** | **70.85** | **86.79** | 117.88 | **50.94** | **57.39** | 64.47 |

Runs 2 and 3 agreed within 1 %; run 1, contended (load rising into the teens), read 184.4 ms at dpr 1.

### Task 17 (T8, R3, over the socket)

`bun scripts/stream-s4/tile-p95.ts`: `grid500x100`, `visibleMark: 'fixed'`, forensic on; one tile per request, serial; zoomStep 5 (zoom 1) at dpr 1 (13 × 20 = 260 tiles) and dpr 2 (25 × 40 = 1,000); warm document cache, cold render cache (fresh server and `Pool({})`); `'message'` and `'tile'` rates raised. t0/t1 from `StreamServerOptions.onTileTiming`. No suite of this task ran alongside, but the machine was not quiet: the 1-minute load of about 6 came from other processes (Preview, WindowServer, PerfPowerServices and another session's repeatedly restarting vitest worker, about one core each), and no idle window came during the task.

| Run | 1-min load (start / after serial / end) | dpr 1 p50 / **p95** / max (ms) | dpr 2 p50 / **p95** / max (ms) | all 1,260 p95 | recycles | client send→`tiles-done` p95, dpr 1 / dpr 2 |
|---|---|---|---|---|---|---|
| 0 (smoke) | 5.99 / 5.85 / 6.59 | 69.1 / **82.7** / 873.6 | 49.4 / **56.0** / 883.9 | 75.0 | 4 | 83.9 / 57.0 |
| 1 (recorded) | 5.93 / 5.80 / 5.80 | 70.3 / **83.6** / 883.5 | 49.3 / **56.0** / 886.1 | 75.7 | 4 | 84.8 / 57.1 |

Both runs are reported; neither was chosen over the other. Verdict: **GREEN** at dpr 1 and dpr 2. Task 9's interim measurement (a scratch script, t0 at the `pool.tile` call) read 82.8–83.6 ms and 56.2–56.6 ms at load 3.3–4.5.

| Not gated | Run 0 | Run 1 |
|---|---|---|
| `--batch 64`: 1,000 dpr 2 tiles | 27.7 s, 36.2 tiles/s, 3 recycles | 27.1 s, 37.0 tiles/s, 3 recycles |
| `--batch 64` per-tile t0→t1 p50 / p95 / max (ms) | 952.7 / 2,042.1 / 2,514.1 | 936.1 / 2,040.8 / 2,512.8 |
| loopback dpr 1 p50 / p95 / max (ms) | 72.0 / 85.4 / 885.6 | 71.1 / 84.6 / 890.2 |
| loopback dpr 2 p50 / p95 / max (ms) | 51.5 / 58.2 / 896.0 | 51.9 / 58.5 / 894.9 |

- **The max is the worker recycle.** Each tile counts one page toward the pool's 256-page recycle; 4 recycles in 1,260 tiles, and the tile waiting behind each re-open and snapshot rebuild takes about 0.87–0.89 s. That is above the p95 and is 4 tiles in 1,260; it stays as a follow-up.
- **Sanity.** Every R3 interval was at or under the client's own send → `tiles-done` interval (0 violations in both runs).
- **`--batch 64`** (reported, not gated): a fresh server, the 1,000 dpr 2 tiles in 16 batches of 64. Rendering is serial per document (`pool.tile` is exclusive per docRef slot), so the per-tile t0→t1, measured from the batch's admission, includes the queue wait behind up to 63 tiles: R3's alternative reading.
- **Loopback** (reported, not gated): the real client (`StreamSession.tiles`) over a socket, from the request to `drawImage` of the decoded PNG on an `@napi-rs/canvas` surface. It runs in Bun with napi canvas, not in a browser.

## Compatibility matrix

`bun scripts/stream-s4/compat-matrix.ts --base 53668e81 --tmp <scratchpad>`, exit 0, at a8aa1a8d:

```
PASS (a) base client x current server: no sheetMode on the raw manifest; unit (raster), 1 search hit(s), copy 22 chars, 1 print page(s) 336x20
PASS (b) current client x current server: sheetMode tiles, sheet-layout decoded (0 row / 0 col intervals), tile (0,0) decoded 512x512, manifest validated on grid500x100
PASS (c) current client (opted in, HELLO carried sheet-tiles@1) x base server: accepted, no sheetMode, unit arrives (raster)
PASS (d) 17 features (16 distinct unknowns + sheet-tiles@1) closed 4000 = unsupported-version; the 16-feature control was accepted
```

The script removed its base worktree afterwards (`git worktree list` unchanged). The old-client half also rests on the Stage 3 XLSX suites, which pass with no edit to their files (Task 16: 169 pass; `git diff --stat 53668e81` on them shows only the new `xlsx-units-baseline.test.ts`).

## Gates

Run one at a time at a8aa1a8d (code) on the same machine, with nothing else of this task running.

| Gate | Command | Result |
|---|---|---|
| server | `cd server && bun test` | **1,462 pass, 0 fail**, 96 files, 815 s |
| client | `bunx vitest run` | 15,690 pass, **1 fail**, 65 skipped, 1 todo; 762 of 763 files. The fail is `tests/slideshow-screens.test.ts` (pre-existing, above) |
| browser, stream | `IV_HARNESS_PORT=4391 npm run test:browser:stream` | **47 passed, 1 skipped** (the existing no-codec media case), 3.3 min |
| browser, viewer (unmodified) | `IV_HARNESS_PORT=4391 npx playwright test --config playwright.viewer.config.ts --project=chromium tests/browser/interactive-viewer` | 65 passed, **2 failed**: the two manual-device collector tests of `interactive-viewer.spec.ts` (port clash, pre-existing, below) |
| browser, viewer, moved stream port | `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npx playwright test … tests/browser/interactive-viewer.spec.ts` | **25 passed** |
| typecheck | `npx tsc --noEmit -p .` | clean |
| typecheck, server | `npm run -s typecheck:server` | clean |
| lint | `npm run -s lint` | clean |
| eager gate | `bun run gates:core` | core closure ok, engine modules ok, core purity ok; index closure FAIL as at the base (§Size gate) |
| compatibility | `bun scripts/stream-s4/compat-matrix.ts --base 53668e81` | 4 PASS, exit 0 |
| T8 | `bun scripts/stream-s4/tile-p95.ts` | GREEN, exit 0 |
| corpus | `bun scripts/fetch-corpus.ts`, then `cd server && bun test test/sheet-corpus.test.ts` | 70 verified; 1 pass |

## Size gate (T11)

```
ok   core closure 300 B over 1 chunk(s) ≤ 38720
ok   engine modules in core closure: none
ok   core purity: no react imports, no document global
FAIL index closure 426556 B (baseline 417436, Δ 9120, recorded 3193, limit 420645)
FAIL index module set differs from baseline ± recorded changes:
       + src/xlsx/rowFont.ts
```

- **The base was already red.** At 2de7ae4a, before any 4a code, Task 2 measured the index closure at 426,511 B against the 420,645 B limit, with `+ src/xlsx/rowFont.ts` in the module set (xlsx-cf, 8f85309c). The gate is not in CI. No limit or baseline was changed.
- **4a's delta is +45 B** (426,556 − 426,511), with no module added: the `rowFont.ts` line is the base's. Task 4 recorded +10 B and Task 11 +32 B (the plumbing its brief mandated), which with lazy-chunk hash noise accounts for it. The constraint expected growth within ±4 B; +45 B is outside that, and is reported here as a finding rather than absorbed.
- `dist/core.js` stays pure and engine-free. The stream sheet viewer (`RemoteSheetView`, a lazy chunk of about 9.3 kB from Task 14) is reached only lazily from `RemoteDocView` and reuses the `InteractiveSheetDoc` chunk.

## Known flakes and environment notes

- **`interactive-viewer.spec.ts`, 2 manual-device collector tests, at the default ports.** The viewer config always starts the stream harness on `OV_STREAM_PORT` (default 4181, identity on 4182), and the two collector tests pick 4181 and 4183 themselves, so they fail with "port 4181 is already in use and cannot be reused: the server on this port is not the officeview viewer harness (no identity marker)". **This predates 4a:** the same spec run at main 53668e81, in a detached worktree with the same command and default ports (nothing else listening on 4173/4181–4183/4391/4281 before either run), failed the same two tests the same way (23 passed, 2 failed). With `OV_STREAM_PORT=4281` all 25 pass at 4a's head. The config and the spec are unchanged by 4a.
- **`server/test/print-wire.test.ts` "two cancelled jobs"** timed out once under heavy parallel load (Task 5) and passes alone; it passed in this run.
- **`tests/xlsx-viewer-worker-overflow`** failed once under load (Task 13) and passes alone; it passed in this run.
- **`tests/slideshow-screens.test.ts`**, one case, fails deterministically under Node 26.13.1 here and at 53668e81 (no `localStorage` in that realm); not 4a's, not worked around.
- **Load.** The T8 runs were taken at a 1-minute load of about 6 on 14 cores, from other processes. Task 1 showed dpr 1 is load-sensitive (184 ms when the load rose into the teens).
- **F27.** The 4a edits to `tests/stream` are append-only: `git diff --numstat 2de7ae4a a8aa1a8d -- tests/stream` gives `38 0 manifest.test.ts`, `21 0 remote-session.test.ts`, `223 0 remote-sheet-view.test.tsx`, `441 0 remote-sheet.test.ts`, `138 0 sheet-wire.test.ts` (0 deletions); the other `tests/**` changes are new files.

## Rulings and findings as built

The full text is in the S4 design, §11.2 "As built (4a)"; each is tested where a test is named.
- **Blank panes (owner, 2026-11-23):** a marked blank pane never names a wrong session; blank tiles stay marked (d366d96c; `trace-tiles.test.ts`). This replaces "flat tiles may abstain" in parent §19.
- **B1 (predates 4a, open for the owner):** in units mode an XLSX unit over 8192 px gives a manifest the client rejects. Tiles mode advertises the refusal unit (5d6fc8d7; `sheet-layout.test.ts` "an oversize sheet opens on a validating client in tiles mode; its print refuses (B1)"). Units mode is unchanged.
- **B2:** a pane's state is layout, so frozen-pane workbooks open under `visible-only`, in units mode too (6a554adf).
- **Measured spill** (d4ece036) and the worst-case CF font (0b583d55) replace R8's whole-run spill (`sheet-snapshot.test.ts` "measured spill …"). The whole run is still taken for: no canvas backend; rotated or stacked text; a throwing layout; region styles; a CF sheet whose CF set or a rule is unresolved, or a rule writing an unknown set, a font size or name, or a number format, or an icon set.
- **Coverage blanking** stops covered cells painting but does not make their values underivable; it extends to every sheet's CF view and drops sparklines located in covered cells (36cb848e; `sheet-snapshot.test.ts` "cross-sheet CF …", "a sparkline located in a covered cell is absent (fix round 1)").
- **The visible mark** scales with zoom·dpr; it lives in sheet space, so seams stay continuous (eeae85cd).
- **Worker-recycle stall:** about 0.9 s on the tile after a recycle, about 1 tile in 256; a follow-up (Task 9, and the T8 max above).
- **Client:** a 5 s sheet request deadline (c4d5a79b); at most 3 retries in a row of a transient visible-tile failure (c445018a); a pacer at 8 per 1 s and 80 per 10 s; a `retryAfterMs` fallback of 1,000 ms, capped at 10 s (f414af01; `tests/xlsx-stream-tiles.test.ts`, `tests/stream/remote-sheet.test.ts`).
- **Units fallback per document:** more than 256 sheets, or a sheet name over 31 characters (5aa80083; `sheet-handshake.test.ts` "257 sheets", "a 32-character sheet name", "256 sheets and a 31-character name still get tiles").
- **F27:** the edits to `tests/stream` are append-only (§Known flakes).
- **`onTileTiming`** (`StreamServerOptions`, 96fdf4cf) is test and measure only and reads no clock when unset.

## Not machine-checked

- Real host themes: tiles use the fixed `light` and `dark` server palettes from the default theme tokens; no host's customised canvas tokens were tried, and by design they do not reach tiles.
- Non-Chromium browsers: the stream browser project runs Desktop Chrome only; WebKit and Firefox were not run on tiles.
- The loopback figure is Bun + napi canvas, not a browser's decode and `drawImage`.
- Look and feel of the visible mark, which scales with zoom·dpr (small at 0.25, large at 4×): an owner call, not a test.

## Follow-ups and open items

- **B1 (pre-existing, owner):** in units mode an XLSX unit over 8192 px gives a manifest the client rejects, so the workbook does not open. Tiles mode advertises the refusal unit (5d6fc8d7). Units mode is unchanged.
- **Worker-recycle stall:** about 0.9 s on the tile after each recycle (about 1 in 256 tiles). Options: lighter recycle weighting for tiles, or a pre-warm.
- **Spill room** is measured on the authored geometry while tiles paint the blanked one (no leak).
- **R10:** finding the tile window of a capture automatically.
- The deferred minors recorded in the ledger (`.superpowers/sdd/2026-11-22-streaming-s4a/progress.md`).
