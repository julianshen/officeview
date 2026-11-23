# Streaming Stage 4a: streamed spreadsheets as server tiles, read-only (implementation plan)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A host that opts in (`streamSheets: 'tiles'`) opens a streamed XLSX in the full interactive sheet viewer: headers, scrolling to row 1,048,576, sheet tabs, authored frozen panes and preset zoom steps. Every pixel is a server-painted, marked 512 device-px tile. No cell value, formula or hidden content reaches the client, and every other client keeps Stage 3 units unchanged.

**Architecture:**
- **Negotiation.** The HELLO gains a bounded `features` list. `sheet-tiles@1` puts an XLSX session in tiles mode, fixed for the sid. Without it the session stays in today's units mode.
- **Server.** In tiles mode the open builds no static units. The pool worker keeps a per-sheet servable snapshot over the projected parse and answers two requests:
  - `sheet-layout`: paint-facing geometry only;
  - `tiles` batches: one `XLSX_TILE` (kind 5) message per tile, painted by `renderSheetViewport`. Each tile carries the visible mark laid out in global sheet space and a per-tile Layer-1 basis, and is audited in `tile_deliveries` and `sheet_views`.
- **Client.** A new `StreamWorkbook` source (widened, not subclassed) holds geometry only, and its store rejects every content read. `SheetTileRenderer` gains a stream branch that requests missing tiles instead of painting them.

**Tech Stack:** TypeScript; Bun server (`server/`, `bun test`, `bun:sqlite`, pool `Worker`s, the server canvas backend); React client (`src/components/xlsx`, `src/components/remote`), vitest, Playwright `stream` project.

**Spec:** `docs/superpowers/specs/2026-11-22-streaming-s4-design.md` (binding; cited as "S4 §N"), above all §6, §8, §9 4a, §10 and §11. Parent: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` ("parent §N"): §6.2, §6.4, §7.1, §8.3, §9.3, §12, §13.5, §14, §15, §16 T2/T8/T9/T11 and §19. Stage 3 print: `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.2. Base: main `53668e81` plus spec commit `a84b2e2f`.

## Global Constraints

- **Owner decisions (S4 §2, §11.1):** full local viewer; approach A; three PRs, of which this is 4a; the parent's decisions stand; tiles chosen per client in the HELLO, with units kept indefinitely; print keeps static per-sheet units, rendered only on a print request.
- **Wire, verbatim:**
  - `features`: at most 16 entries, each at most 32 characters. The capability is `sheet-tiles@1`.
  - Tiles are 512 device px on the client's `tileRangeFor` grid.
  - `zoom` ∈ `SHEET_ZOOM_STEPS` (11 values) and `dpr` ∈ `{1, 2}` (dpr ≤ 1 → 1, otherwise 2): at most 22 pairs, keyed as a pair and never as their product.
  - `palette` ∈ `{light, dark}`.
  - At most 64 tiles per batch.
  - `XLSX_TILE` = kind 5: ≤ 8 MiB encoded and ≤ 64 MiB decoded. Every new text message is ≤ 64 KiB.
- **What never reaches the client:** a cell value, display text, formula, number behind a format, style, number format, CF rule, drawing content or bytes, chart model, table style, pivot definition, defined name or comment. Also hidden row, column or sheet content, `nonUniformFontRows` and `defaultRowPt`. The tests check transmitted payloads, not screenshots (parent §3).
- **What the client never does on a stream:** call `renderSheetViewport`, stamp its own watermark, or build a local text index. `stats.basePaints === 0` for a stream frame.
- **Rates (S4 §6.8):**
  - `'tile'` is counted per tile, refuse-only, never toward the 4008 breach window, and separate from `'unit'`. Sizing (R4): 600 tiles per 10 s, burst 120 per 1 s.
  - `sheet-layout` is charged to `'unit'` once per request; its later chunks charge nothing.
- **T8 (4a):** tile p95 ≤ 120 ms on the 500×100 fixture, measured as R3 states.
- **Snapshot (S4 §6.2):**
  - The tile import opens over the **projected package**, never over `DocState.bytes`.
  - The withholding steps take `contentPolicy` and are the identity under `'full'`.
  - Tiles paint saved values; nothing recomputes on a stream.
- **Determinism (parent §14):** a tile reads no clock and no randomness. The session clock `(day, timeZone)` is injected, and a client `nowMs` is never read.
- **Cache (parent §9.3, S4 §6.5):** the §9.3 key with `policyDetails.viewStateHash`, plus `sheetIndex ⊕ zoomStep ⊕ dpr ⊕ palette ⊕ col ⊕ row`. The sid is never in the hash. Canonical entries are pre-forensic and live under the 512 MiB render-cache ceiling.
- **Layer-1 basis label (S4 §6.5):** `"sheet:" + sheetId + "@" + zoom + "x" + dpr + ":" + col + "," + row`, where `zoom` is the `SHEET_ZOOM_STEPS` value printed by `String()` (e.g. `0.33`), unless Task 1 rules for the crop.
- **Units mode is untouched:**
  - an XLSX session without the capability gets today's manifest, key for key;
  - its search, copy, a11y and print are unchanged;
  - the Stage 3 suites pass with no edit to their files.
- **Opt-in:** the 4a client sends `features` only when the host passes `streamSheets: 'tiles'` (S4 §6.10).
- **Local mode:** every existing `tests/**` and `tests/browser/interactive-viewer*.spec.ts` suite passes unmodified.
- **Bundle (T11):** the stream sheet viewer loads lazily from `RemoteDocView` and reuses the `InteractiveSheetDoc` chunk; `dist/core.js` must not reach it.
  - Record `bun run gates:core` at the merge base in Task 2, and again after every client task.
  - Growth beyond ±4 B of lazy-chunk hash noise in the index closure is a finding. The baseline was already red on main (422,664 B against 420,540 B, `docs/validation/streaming-s3-p5.md`).
- **Reviewer:** each task is reviewed by a different model family from its implementer before the next starts.
- **Rulings:** the plan's calls are R1–R16 below. A task that finds one wrong stops and reports; it never deviates silently.
- **Commits:** by path (`git commit -m "<msg>" -- <paths>`), never `git commit -a` and never a bare `git stash`. Messages end with:
  ```
  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01B5aP5T5pG9oGcVPzuNpnmG
  ```
- **Commands:**
  - server: `cd server && bun test <file>`;
  - client: `bunx vitest run <file>`;
  - browser: `IV_HARNESS_PORT=4391 npm run test:browser:stream`;
  - typecheck: `npx tsc --noEmit -p .` and `npm run -s typecheck:server`;
  - eager gate: `bun run gates:core`;
  - corpus: `bun scripts/fetch-corpus.ts`, which must run before any corpus step.

## Rulings (planner's calls on spec silences; Task 17 writes the wire-visible ones into the specs)

- **R1. The gate is implied, not named.** The spec names no feasibility gate. Three of its statements could each change the design:
  - §4 says tile latency becomes a server cost;
  - §6.5 says either parse path must hold the p95;
  - §8 leaves the basis choice to the T9 battery on tiles.

  So Task 1 is a STOP-if-red spike, with no product code.
- **R2. Resident parse.** The worker paints from the resident projected parse: `detectAndParse`, then `buildSheetGeometry`, then `prepareSheetAssets`. It does not use `openImportSession`. Task 1 measures the million-row fixture; if that leg is red, Task 1 re-runs on `openImportSession` before reporting.
  - Task 1's thresholds for this leg (open ≤ 10 s, heap growth ≤ 256 MiB) are planner's calls, not spec values: a red result there is reported with the numbers and goes to the owner, rather than changing the design on its own.
  - Cost if wrong: the sheet source behind `DocumentHost.sheetSnapshot` changes; its interface does not.
- **R3. The p95 protocol.** "Tile-request admission" (S4 §6.8) is ambiguous for 64-tile batches, because later tiles would include queue wait.
  - **The clock.** t0 is the instant a tile passes validation, the token check and the `'tile'` rate check, which is when the driver would hand it to the pool. t1 is when the tile's last fragment has been passed to `sendBinary`.
  - **The driver.** One tile per request, sent serially: the next request goes out after the previous `tiles-done`.
  - **The run.** Zoom 1 at dpr 1 and at dpr 2 is what "scale 1 and scale 2" means. The document cache is warm (open and `sheet-layout` done) and the render cache is cold (a fresh server). The visible mark is `fixed` and forensic marking is on.
  - **Rates during the run.** 1,260 serial requests would trip the per-sid `'message'` budget (120 per 10 s, 20 per s) and then close 4008. The measurement therefore runs with `StreamServerOptions.rates` overrides for `'message'` and `'tile'`. Task 9 adds that option, TEST and measure only.
  - **Also reported, not gated:** 64-tile batch throughput, and the loopback figure from request to `drawImage`.
  - Cost if wrong: if the owner means arrival-to-last-fragment per batched tile, the script's `--batch 64` mode reports that figure and the concurrency in R4 is re-tuned.
- **R4. Batches, admission and rate.**
  - The pool's per-sid admission is `perSessionOutstanding 8` and `perSessionQueue 32` (`server/src/admission.ts:16–21`). A 64-tile batch is therefore never queued whole: the server driver keeps at most 8 tile jobs (render plus Layer-1) in flight per sid.
  - The client keeps at most 2 batches in flight, so at most 128 tiles are outstanding.
  - `'tile'` is charged per tile when a batch is admitted. Tiles beyond the window's remaining budget are refused individually in `tiles-done.refused` with code `rate-limited`, and `tiles-done.retryAfterMs` says when the budget frees.
- **R5. One OpenDoc, lazy paintables.** `DocumentHost.open` caches by `docRef|policy|fingerprint`, so units and tiles sessions share one entry.
  - An XLSX entry opened first in tiles mode builds no paintables. Any units path (`render`, `text`, `print`, `trace`) builds them on first use, once.
  - A units-mode open stays eager, so units mode is unchanged.
  - Manifest unit (print page) sizes come from `staticSheetSpec(sheet)`, which is the `computeMetrics` size or the refusal spec, without rendering. The "sized at print time" fallback is not needed.
- **R6. Stage 3 text ops in tiles mode.** For an XLSX tiles session, `unit`, `search`, `select`, `copy` and `a11y` are refused `unsupported-op`, because serving them would render units. `print` works.
  - `can.*` stays the policy's claim; the 4a client simply mounts no find or copy (S4 §6.11).
  - The a11y grid shows positions only, plus the parent §15 notice once.
- **R7. `sheet_views` is created in 4a,** with the identity-view row `{sort: null, filters: [], sizes: []}` and its clock, written before a sid's first tile of a sheet. The 4a tracer needs it; 4c extends the spec column.
- **R8. Snapshot steps in 4a.** The parts of S4 §6.2 that change tile pixels land here; the rest are under "Not in 4a".
  - **Lands in 4a:**
    - the policy parameter;
    - step 1, with its hard-constraint test;
    - step 4, the base coverage, and step 5, base-covered cells painting empty under `visible-only`;
    - step 7, saved values.
  - **The spill rect is conservative, as builder's call.** Text overflows into the contiguous run of empty cells on the side(s) its alignment allows:
    - general or left-aligned text spills right;
    - right-aligned text spills left;
    - centred text spills both ways;
    - wrap, fill, justify, shrink-to-fit and numbers never spill.

    A conservative rect is fail-closed.
  - Cost if wrong: if the owner defers step 5 to 4b, the change is one flag in `buildSheetSnapshot`.
- **R9. The session clock.**
  - `timeZone` is `AuthDecision.policy.timeZone`, an IANA zone validated with `Intl`. When it is absent, the server's `sheetTimeZone` option applies (default `'UTC'`); when it is invalid, the open closes `policy-refused`.
  - `day` is the calendar date of the server's `now()` at open, in that zone. Two sessions with the same `day` but different zones (e.g. `UTC` and `Europe/London` in November) still never share.
  - The pair is fixed for the sid: it survives resume, and a session that crosses midnight keeps its day. It is injected as the CF `dateContext`, with `epochNowMs` at 12:00 local time on `day`, and as the evaluation `timeZone`.
  - The zone is held on the session entry beside `features`, not in `EffectivePolicy`, so existing policy deep-equal tests are untouched. A resume whose decision resolves to another zone closes `auth-failed`.
- **R10. Tracing a pane capture.**
  - The investigator names the tile window the capture shows: `{sheetId, zoom, dpr, c0..c1, r0..r1}`.
  - `traceTiles` rebuilds each named tile through the host from `sheet_views` and the clock, and composites a window reference. It builds the window chips as each tile's own basis, offset by the tile's origin.
  - It then registers the capture against that composite. The candidates are the sids with a marked `tile_deliveries` row in the window, and `MAX_CANDIDATES` applies unchanged.
  - Finding the window automatically is a follow-up. Task 1 measures this approach; if it is red, the gate decides between it and the `markLayerRaster` crop (S4 §8).
- **R11. The compatibility matrix.**
  - **The pre-4a builds.** These are `53668e81`, checked out by `scripts/stream-s4/compat-matrix.ts` as a detached temporary worktree with `node_modules` symlinked. The server runs from source, and the old client is the base tree's `StreamSession`, imported by path.
  - **Old client against the new server:** the base client, plus the unchanged Stage 3 XLSX suites, whose harness HELLO is byte-for-byte the old client's.
- **R12. Small wire calls.**
  - **`features`:** a well-formed unknown entry is ignored. A list that is not an array, or that has more than 16 entries, a non-string entry, an empty entry, an entry over 32 characters or a duplicate, fails `validHello` and closes `unsupported-version`.
  - **`sheet-cancel {target}`** lands in 4a, for tile batches. An unknown target is ignored.
  - **Sheet ids.** The wire `sheetId` is the decimal index of the sheet among the **visible** sheets in tab order (`"0"`, `"1"`, …), and the payload's `sheetIndex` is the same number. The workbook's own `sheetId` attribute is never sent, because its gaps would reveal hidden sheets.
  - **Zoom** travels as `zoomStep`, an index from 0 to 10 into `SHEET_ZOOM_STEPS`, never as a float.
  - **Palette.** `light` and `dark` are fixed server palettes taken from the default theme tokens. Only the `pending`/`pending-text` colours reach a tile (`ViewportPaintOptions.colors`), and host-customised canvas tokens do not reach tiles.
  - **New refusal codes:** `unsupported-op`, `stale-view` and `layout-too-large`. A malformed request is refused `policy-refused`, as print does.
- **R13. The `XLSX_TILE` body.** It is an `encodeUnit` body with `unitId = sheetIndex` and `discriminator = 3`. The body is `u32 viewGen, u8 zoomStep, u8 dpr, u16 0, u32 col, u32 row`, followed by the `UNIT_RASTER` raster body (`u32 w, u32 h, u32 pngLen, PNG`) with `w = h = 512`.
- **R14. `sheet-layout` bounds.**
  - Each list may hold at most 65,536 entries (row intervals, column intervals, collapsed runs per axis, merges), and filter ranges are capped at 1,024.
  - The answer spans at most 64 messages, each ≤ 64 KiB, with `seq` and `last`.
  - Anything over a cap is refused `layout-too-large`.
- **R15. Partial edge tiles** are full 512² surfaces, blank beyond the sheet extent, so every tile has the full Layer-1 capacity (16,384 carrier cells). Flat tiles may abstain (parent §19 item 8).
- **R16. Disposition of the viewer's `remote.*` members.** `InteractiveSheetDoc` reaches 30 members through `remote` (`grep -o 'remote\.[A-Za-z]*' | sort -u`). `isRemoteWorkbook` has no user outside `InteractiveSheetDoc.tsx` and its definition, in `src` or `tests`.
  - **Shared by both sources** (the narrowed `ViewerSheetSource`): `layout`, `merges`, `store`, `identity`, `geometry`, `isClosed`, `subscribe`, `loadState`, `view`, `sort`, `filter`.
  - **Paged only**, reached under `kind === 'paged'`: `drawings`, `styleOf`, `acquireCfPreparation`, `invalidateCf`, `cfDependencyEpoch`, `takeCfDirtyRects`, `sparklines`, `subscribeSparklineEvictions`, `regionStyles`, `subscribeRegionStyleEvictions`, `pivotDisplay`, `pivots`, `retainOverflow`, `overflowFor`, `mainLedger`, `hasFormulaCellsOutsidePivots`, `formatting`, `findSession`, `retainCfDependencies`.
  - On a stream, `sort` and `filter` reject `stream-no-cell-data` until 4c, and `view` is the identity view.

## Review Focus

1. **A resume while tile batches are in flight.** Expected: every tile still pending is requested again once, under the new token. Late tiles from the old socket are dropped, and no tile stays pending. Pinned in Task 11.
2. **A `rate-limited` tile.** Expected: it is retried after `retryAfterMs` and never stays pending. Ten such refusals in a row never close the socket with 4008. Pinned in Tasks 9 and 13.
3. **A fractional dpr (1.25, 1.5) or a theme switch mid-session.** Expected:
   - the tiles are requested at bucket 2, and the pane is scaled once by CSS, with no single tile resampled;
   - a palette change requests new keys, and no old-palette tile is drawn under the new key.

   Pinned in Task 13.
4. **A units client and a tiles client on one document at once.** Expected:
   - each gets its own mode;
   - the tiles session never makes the shared entry render a unit;
   - closing either session leaves the other working.

   Pinned in Tasks 3 and 4.
5. **A hostile or buggy server on tiles.** Unsolicited, duplicate, wrong-`viewGen`, wrong-sheet and wrong-size tiles, and tiles after `tiles-done`, are all dropped before decode, and the session survives. Pinned in Task 11.

**Not in 4a** (S4 §9 4b and 4c):
- **4b:**
  - selection, keyboard selection and the client's own freeze choice;
  - `sheet-find`, `sheet-copy`, `sheet-a11y`, `sheet-fit` and `sheet-resize`, which bring non-identity sizes and `view-state-limit`;
  - step 3: the merge-covered purge in `projectXlsx` and the undrawn-cell blanking;
  - current-view coverage, and the per-cell ink recording oracle;
  - the S4 §6.11 notices other than the §15 one;
  - advertising `sheet-tiles@1` by default.
- **4c:** sort, filter, `filter-values`, `range-detect`, C1 (step 2), step 6, the typed-condition refusal, the `'sheet'` class, `pivot-refused` and `sheet_ops`.

---

### Task 1: Feasibility gate: headless tiles, tile latency, Layer-1 on tiles, million rows (STOP if red)

**Files:**
- Create: `scripts/stream-s4/tile-feasibility.ts`, `server/test/sheet-fixtures.ts`, `docs/validation/streaming-s4a-feasibility.md`, `docs/validation/streaming-s4a-feasibility.json`
- No product source changes. The script imports read-only from:
  - `src/xlsx/view/{viewport,renderViewport,sheetZoom,tileCache}.ts` and `src/components/xlsx/sheetAssets.ts`;
  - `server/src/{backend,policy,fonts}.ts`, `server/src/mark/*` and `server/src/detect/tau.ts`;
  - `scripts/stream-s0/mark/attacks.ts` (`jpeg`).

**Interfaces:**
- Produces, in `server/test/sheet-fixtures.ts` (used by every later task):
  - `grid500x100(): Promise<Uint8Array>`: `buildXlsx`, sheet `S`, cells `R{r}C{c}` in A1:CV500 with mixed text and numbers;
  - `millionRows()`: A1:D2000 and A1048576 populated, with a row-height band at rows 10–20;
  - `blankSheet()`;
  - `cfToday()`: a `timePeriod today` fill rule on A1:A10, which hold dates;
  - `unfilledShapeOverText()`: B2 holds `COVERED-B2`, under an unfilled rectangle;
  - `rotatedBox()`: S4 §6.2 step 4's `100×20` anchor at (0,0) rotated 90°, over cells holding `ROT-r-c`;
  - `unresolvedAnchor()`;
  - `savedFormula()`: `=1+1` with saved value `7`, and `plainSeven()`, which holds the value `7`.
- Produces, for Tasks 5–10: the pinned parse path (R2), the basis choice (per-tile or crop) and the in-process latency table, all stated in the report.

- [ ] **Step 1: Write the measurement.** `bun scripts/stream-s4/tile-feasibility.ts [--quick]` runs four legs.
  - **(a) Headless paint.** Over the projected parse of `grid500x100`, under `serverBackend()`, paint tile (0,0) at zoom 1, dpr 1 with `renderSheetViewport`, with `ctx.setTransform(scale…)` as `sheetTiles.ts:340–360` does. Record that the PNG is 512² and differs from a `blankSheet` tile, that two runs are byte-equal, and the `ViewportPaintResult.kind`.
  - **(b) Latency.** In process and single-threaded, for every tile of the sheet at (zoom 1, dpr 1) and (zoom 1, dpr 2), time:
    - the render;
    - the visible mark (`fixed`);
    - the PNG encode;
    - `markRaster` with the per-tile basis;
    - gzip framing.

    Record p50, p95 and max per scale.
  - **(c) Layer-1 on tiles.** Use 64 random session codes.
    - **Captures:** a lossless 2×2-tile pane crop (900×900 at offset (40,40)), the same through JPEG q70, and the same at a 50 % downscale.
    - **Decision:** register against the R10 composite reference, and decide with `tauFor(64)`.
    - **Recorded:** the attribution rate and wrong attributions per capture class, for 30 captures each.
    - **Blank pane:** a `blankSheet` pane, which must abstain (`flat`).
    - **The alternative:** the `markLayerRaster` crop of a sheet-wide basis, on the same captures.
  - **(d) Million rows.** Record the open (project, parse, `buildSheetGeometry`, `prepareSheetAssets`): wall time and heap growth. Then record the wall time of a render of the tile that holds row 1,048,576.
- [ ] **Step 2: Run** `bun scripts/stream-s4/tile-feasibility.ts`. Expected: the JSON is written, every leg has numbers, and a second run reproduces (b) within ±10 % and (a) and (c) exactly.
- [ ] **Step 3: Write the report.** Each gate gets its numbers and a verdict:
  - (a) the headless paint works and is deterministic;
  - (b) in-process p95 ≤ 100 ms at both scales is GREEN, which leaves 20 ms for the pool hop and the socket. 100–120 ms is AMBER: continue, and report. Over 120 ms is RED;
  - (c) lossless ≥ 90 %, q70 ≥ 75 % and 50 % downscale ≥ 75 % (the parent §13.5 image-only targets); zero wrong attributions; the blank pane abstains 30 times out of 30;
  - (d) open ≤ 10 s with heap growth ≤ 256 MiB (planner's calls, R2), and the last-row tile renders. If red, re-run (d) on `openImportSession` (R2) before the verdict.

  Then pin the parse path and the basis: per-tile, unless the crop's lossless rate is at least 5 points higher (S4 §8).
- [ ] **Step 4: Verify.** If any gate is RED: commit the report, STOP, and report to the controller and owner, changing no gate. If none is: continue.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "test(stream-s4): tile feasibility gate — headless paint, latency, Layer-1, million rows" -- scripts/stream-s4/tile-feasibility.ts server/test/sheet-fixtures.ts docs/validation/streaming-s4a-feasibility.md docs/validation/streaming-s4a-feasibility.json
  ```

### Task 2: Units-mode characterization, green before any 4a code

**Files:**
- Create: `server/test/xlsx-units-baseline.test.ts`
- Modify: none

**Interfaces:**
- Produces: the units-mode contract that Tasks 4 and 16 must keep. The existing XLSX Stage 3 suites are also part of it: `server/test/text-search-wire.test.ts`, `text-search-bound.test.ts`, `text-a11y-wire.test.ts`, `text-served.test.ts` and `print-wire.test.ts`.

- [ ] **Step 1: Write the characterization tests** on a 2-sheet `buildXlsx` with today's harness HELLO.
  - **`"units-mode xlsx manifest has exactly the Stage 3 keys"`:**
    - `Object.keys(manifest).sort()` deep-equals `['can', 'caps', 'docRef', 'kind', 'layoutFingerprint', 'limits', 'name', 'notices', 'policy', 'recorder', 'session', 'units', 'v']` (`ws.ts:420–434`);
    - `units.count === 2`;
    - `'sheetMode' in manifest`, `'sheets' in manifest` and `'tiles' in manifest` are all `false`.
  - **`"units-mode xlsx copy answers sheet text"`:** a `select` then `copy` over unit 0 returns a string containing the A1 text. This fills the one Stage 3 op without an XLSX test.
  - **`"units-mode xlsx unit request renders"`:** `unit 0` at scale 1 gives one kind-3 message.
- [ ] **Step 2: Run** `cd server && bun test test/xlsx-units-baseline.test.ts`. Expected: PASS on the base. These are characterization tests, not RED-first.
- [ ] **Step 3: Record** `bun run gates:core` at the base; Task 17 quotes the figure.
- [ ] **Step 4: Commit.**
  ```bash
  git commit -m "test(stream-s4): characterize units-mode XLSX before tiles" -- server/test/xlsx-units-baseline.test.ts
  ```

### Task 3: Lazy XLSX paintables, print sizes without rendering, visible sheet list

**Files:**
- Modify:
  - `server/src/render.ts`: `DocumentHost.open` at `:198–252`, plus `render`, `text`, `notes` and `drop`;
  - `server/src/pool.ts`: `open` at `:318`, `ensureOpen` at `:303`;
  - `server/src/worker/render-worker.ts:87`;
  - `src/render/paint.ts`: `sheetPaintable` at `:146`.
- Test: create `server/test/sheet-open.test.ts`

**Interfaces:**
- Produces:
  - `export function staticSheetSpec(sheet: XlsxSheet): PageSpec` in `src/render/paint.ts`. It returns `computeMetrics(sheet)`'s `{widthPx, heightPx}`, or `REFUSAL_SPEC` when `exportLimits` is non-empty. `sheetPaintable` uses it.
  - `DocumentHost.open(docRef, bytes, contentPolicy, o: { sheetMode?: 'units' | 'tiles' } = {})`. In tiles mode an XLSX entry builds no paintables, and `units` comes from `staticSheetSpec` with `sizeCeil` as today.
  - `OpenResult.sheets?: readonly string[]`: the projected (visible) sheet names in tab order, for XLSX.
  - `DocumentHost.paintablesBuilt: number` (TEST-visible), and a private `paintablesOf(doc): PaintableArray` that every units path calls.
  - `OpenDoc.parsed`: the projected `XlsxDocument`, retained for Task 5.
  - `Pool.open(docRef, bytes, contentPolicy, o?: { sheetMode?: 'units' | 'tiles' })`. Re-opens after a worker recycle use `'tiles'`, because the paintables come back lazily.

- [ ] **Step 1: Write the failing tests.**
  - `"tiles open builds no paintables and renders nothing"`:
    - `grid500x100` opened `{sheetMode: 'tiles'}` gives `paintablesBuilt === 0` and `pages === 0`;
    - its `units` deep-equals a fresh host's units-mode open;
    - `sheets` deep-equals `['S']`.
  - `"a units path builds paintables once"`: after a tiles open, `render(docKey, {unitIndex: 0, scale: 1, …})` twice leaves `paintablesBuilt === 1`.
  - `"units open stays eager and unchanged"`: `paintablesBuilt === 1` at open, and `fontDiagnostics` equals the pre-change value captured in the test.
  - `"units and tiles opens share one entry"` (Review Focus 4): a tiles open then a units open of the same key gives `parses === 1` and `paintablesBuilt === 1`. A `drop` disposes them once.
  - `"hidden sheets are not in sheets"`: on the `xlsx-hidden-sheet` and `xlsx-veryhidden-sheet` fixtures, `sheets` deep-equals `['S']`.
  - `"tiles-mode sizes equal units-mode sizes on every corpus xlsx"`: for each corpus XLSX, a tiles-mode open's `units` deep-equals a fresh host's units-mode open's `units`. The test asserts that it found at least one corpus XLSX (run `bun scripts/fetch-corpus.ts` first).
- [ ] **Step 2: Run** `cd server && bun test test/sheet-open.test.ts`. Expected: FAIL (no `sheetMode` option, no `paintablesBuilt`).
- [ ] **Step 3: Implement** the lazy builder as `paintablesOf`, guarded by the entry's single promise so that two first uses build once. `OpenResult.fontDiagnostics` reads 0 until the paintables exist.
- [ ] **Step 4: Run green:** the same file, then `cd server && bun test`. Expected: all pass, Task 2's file included.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4): XLSX opens without static units in tiles mode; sizes from metrics" -- server/src/render.ts server/src/pool.ts server/src/worker/render-worker.ts src/render/paint.ts server/test/sheet-open.test.ts
  ```

### Task 4: Handshake: `features`, sheet mode, manifest, resume, op gating, session clock

**Files:**
- Create: `server/src/sheet/mode.ts`, `server/src/sheet/clock.ts`, `src/stream/sheet-wire.ts` (constants only in this task)
- Modify:
  - `server/src/ws.ts`: `validHello` at `:1216`, `openSession` at `:303`, `resumeSession` at `:345`, `sendManifest` at `:420`, the dispatch at `:1247`, `RefusalCode` at `:54`;
  - `server/src/hooks.ts`: `AuthDecision['policy']` gains `timeZone?: string` (the type only; `EffectivePolicy` is unchanged, R9);
  - `server/src/audit.ts`: `AuditRowInput.sheetMode`;
  - `src/stream/manifest.ts`.
- Test: create `server/test/sheet-handshake.test.ts`; modify `tests/stream/manifest.test.ts`

**Interfaces:**
- Produces, in `src/stream/sheet-wire.ts`, which is React-free:
  - `export const SHEET_TILES_FEATURE = 'sheet-tiles@1'`, `TILE_PX = 512`, `TILE_BATCH_MAX = 64`, `TILE_DPRS = [1, 2] as const` and `TILE_PALETTES = ['light', 'dark'] as const`;
  - `export const TILE_ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4] as const`, a literal. `sheet-wire.ts` and `manifest.ts` import nothing from `src/xlsx/**`, so the stream runtime's module set and the eager gates do not move. A test asserts that it deep-equals `SHEET_ZOOM_STEPS`;
  - `export type TilePalette = typeof TILE_PALETTES[number]`;
  - `export const HELLO_FEATURES_MAX = 16, HELLO_FEATURE_CHARS = 32`.
- Produces, in `server/src/sheet/mode.ts`:
  - `export function parseFeatures(v: unknown): readonly string[] | 'invalid'`, where `undefined` gives `[]` and anything else follows R12;
  - `export type SheetMode = 'units' | 'tiles'`;
  - `export function sheetModeOf(format: 'docx' | 'xlsx' | 'pptx', features: readonly string[]): SheetMode`.
- Produces, in `server/src/sheet/clock.ts`:
  - `export interface SheetClock { day: string; timeZone: string }`, where `day` is `YYYY-MM-DD`;
  - `export function sheetClockAt(nowMs: number, timeZone: string): SheetClock`, which throws `RangeError` for a zone `Intl` rejects;
  - `export function clockEpochMs(c: SheetClock): number`, the instant of 12:00 on `day` in `timeZone`.
- `sessionDoc` entries gain `features: readonly string[]`, `sheetMode: SheetMode` and `clock: SheetClock` (its `timeZone` resolved per R9), set at open and kept on resume. `StreamServerOptions.sheetTimeZone?: string` (default `'UTC'`).
- Manifest, in tiles mode only:
  - `sheetMode: 'tiles'`;
  - `sheets: Array<{ sheetId: string; name: string }>`, with ids `"0".."n-1"` (R12);
  - `tiles: { px: 512; zoomSteps: number[]; dprs: [1, 2]; palettes: ['light', 'dark']; maxBatch: 64 }`.
- Client: `Manifest.sheetMode?: 'tiles'`, `Manifest.sheets?` and `Manifest.tiles?`.
  - Validated when present: `zoomSteps` must deep-equal `TILE_ZOOM_STEPS`; there are at most 256 sheets; each `sheetId` matches `/^\d{1,3}$/` and they run in sequence; each `name` is a string of at most 31 characters.
  - Absent, they mean units mode.
- `RefusalCode` gains `'unsupported-op' | 'stale-view' | 'layout-too-large'`. A test-only accessor `stream.sheetClockOf(sid)` follows the existing `canonicalOf` pattern.

- [ ] **Step 1: Write the failing server tests.**
  - `"no features → units mode, Stage 3 manifest"`: the same key list as Task 2.
  - `"sheet-tiles@1 on xlsx → tiles manifest"`:
    - `sheetMode === 'tiles'`;
    - on `xlsx-hidden-sheet` (withHidden), `sheets` deep-equals `[{sheetId: '0', name: 'S'}]`;
    - `tiles` deep-equals `{px: 512, zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4], dprs: [1, 2], palettes: ['light', 'dark'], maxBatch: 64}`.
  - `"features on docx and pptx stay units"`: no new keys.
  - `"unknown well-formed features are ignored"`: `['x@9', 'sheet-tiles@1']` gives tiles, and `['x@9']` gives units.
  - `"malformed features close the HELLO"`: each of these closes with `CLOSE['unsupported-version']`:
    - `'sheet-tiles@1'` as a bare string;
    - 17 entries;
    - one 33-character entry;
    - `[1]`;
    - `['', 'a']`;
    - `['a', 'a']`.
  - `"resume must present the same features"`:
    - a tiles sid resumed without features closes `CLOSE['auth-failed']`;
    - resumed with them, it gets `sheetMode: 'tiles'`.
  - `"op gating"`:
    - in a tiles session, `unit`, `search`, `select`, `copy` and `a11y` are each refused `{code: 'unsupported-op'}`, and `print 0..0` gives `print-start`;
    - in a units session, `tiles`, `sheet-layout` and `sheet-cancel` are each refused `unsupported-op`.
  - `"tiles open renders no unit until print"`: with a counting pool wrapper, `pool.render` calls stay at 0 after open and `ready`, and reach 1 after `print 0..0`.
  - `"units and tiles sessions on one document"` (Review Focus 4): open a units sid and a tiles sid on the same bytes.
    - The units sid's `unit 0` renders; the tiles sid's `unit 0` is refused `unsupported-op`.
    - After the units socket closes, the tiles sid's `print 0..0` still gives `print-start`.
  - `"the audit open row records the mode"`: `store.getRow(sid).sheetMode === 'tiles'`, and `'units'` for the other sid.
  - `"session clock"`:
    - `sheetClockAt(Date.UTC(2026, 10, 22, 23, 30), 'Asia/Tokyo')` deep-equals `{day: '2026-11-23', timeZone: 'Asia/Tokyo'}`, and with `'UTC'` gives `'2026-11-22'`;
    - `'Mars/Base'` throws;
    - a sid opened at 23:30 UTC and resumed after `now` passes midnight keeps `day: '2026-11-22'`;
    - `policy.timeZone: 'Asia/Tokyo'` overrides the server default, and `policy.timeZone: 'bad'` closes `CLOSE['policy-refused']`;
    - a resume whose decision gives another zone closes `CLOSE['auth-failed']`.
- [ ] **Step 2: Write the failing client tests** in `tests/stream/manifest.test.ts`:
  - `TILE_ZOOM_STEPS` deep-equals `SHEET_ZOOM_STEPS` (imported in the test only);
  - a valid tiles manifest passes, and `m.sheetMode === 'tiles'`;
  - it is rejected for each of: `sheets` missing, `zoomSteps` reordered, `sheetId: '07'`, ids out of sequence, and a 32-character name;
  - an absent `sheetMode` gives `m.sheetMode === undefined`.
- [ ] **Step 3: Run** `cd server && bun test test/sheet-handshake.test.ts`, then `bunx vitest run tests/stream/manifest.test.ts`. Expected: FAIL.
- [ ] **Step 4: Implement.**
  - The XLSX open passes `{sheetMode}` to `pool.open` (Task 3).
  - The units-mode manifest object is built exactly as today; the tiles fields are spread in only in tiles mode.
  - The resume check adds `JSON.stringify(features)` equality next to `resumeInvariant`.
- [ ] **Step 5: Run green:** the same commands, then `cd server && bun test` and `bunx vitest run tests/stream`. Expected: all pass, with Task 2's and the Stage 3 files unmodified.
- [ ] **Step 6: Commit.**
  ```bash
  git commit -m "feat(stream-s4): HELLO features select tiles mode for XLSX; fixed per sid, session clock" -- server/src/sheet/mode.ts server/src/sheet/clock.ts src/stream/sheet-wire.ts server/src/ws.ts server/src/hooks.ts server/src/audit.ts src/stream/manifest.ts server/test/sheet-handshake.test.ts tests/stream/manifest.test.ts
  ```

### Task 5: Servable snapshot for tiles: projected parse, base coverage, saved values, injected clock

**Files:**
- Create: `server/src/sheet/snapshot.ts`, `server/src/sheet/coverage.ts`
- Modify:
  - `server/src/render.ts`: `DocumentHost.sheetSnapshot`;
  - `src/components/xlsx/sheetAssets.ts`: `SheetAssetsOptions.cfClock`, threaded to `prepareCfPaint` as `dateContext` and as the evaluation `timeZone`.
- Test: create `server/test/sheet-snapshot.test.ts`

**Interfaces:**
- Consumes: `OpenDoc.parsed` (Task 3), `SheetClock`/`clockEpochMs` (Task 4) and the Task 1 fixtures.
- Produces:
  - **The coverage type:**
    ```ts
    export type BaseCoverage =
      | { kind: 'known'; covered(row: number, col: number): boolean; count: number }
      | { kind: 'unknown'; reason: 'unresolved-anchors' | 'unbounded-effect' | 'transform' }
    ```
  - **`export function baseCoverage(sheet: XlsxSheet, geometry: SheetGeometryResult): BaseCoverage`** in `coverage.ts`, which works in the author view (identity order, authored sizes). Two rules it applies:
    - **Effect padding.** Each drawing's extent is the transformed `drawingExtent`, padded by what the renderer actually paints outside it: half the line width, and the resolved `DrawingShadow` (`|offsetX| + blurPx`, `|offsetY| + blurPx`; `src/drawing/style.ts:73`). Effects the renderer never draws (the `EffectStyleDefinition.names`: glow, reflection, scene3d and so on) have no painted extent. A shadow with `scaledOrSkewed` gives `unbounded-effect`.
    - **Fail-closed nodes.** A node with `transformValid === false` gives `transform` (the S4 §6.2 step 4 pitfall), and `anchors.unresolved > 0` gives `unresolved-anchors`.

    A cell is covered iff its grid rect, or its R8 spill rect, intersects any extent.
  - **`export interface SheetSnapshot`:** `sheetIndex: number; sheet: XlsxSheet; geometry: SheetGeometryResult; assets: RenderSheetAssets; coverage: BaseCoverage | 'off'; clock: SheetClock; contentPolicy: 'visible-only' | 'full'; release(): void`.
  - **`export async function buildSheetSnapshot(doc: XlsxDocument, sheetIndex: number, contentPolicy, clock: SheetClock, prepare: (sheet: XlsxSheet) => Promise<SheetAssetsLease>): Promise<SheetSnapshot>`.** The order is fixed: coverage first, then blanking, then `prepare`.
    - Under `'visible-only'` with known coverage, `sheet` is a copy in which every covered cell has its value removed and keeps its style. The assets (CF batches above all) are prepared over this blanked copy, so a covered cell draws no data bar, colour scale or icon either.
    - Under `'full'`, coverage is `'off'` and the sheet is not copied.
  - **`DocumentHost.sheetSnapshot(docKey: string, sheetIndex: number, clock: SheetClock): Promise<SheetSnapshot>`**, cached per `(docKey, sheetIndex, day, timeZone)` (one promise per key) and released on `drop`.

- [ ] **Step 1: Write the failing tests.**
  - `"the snapshot opens over the projected package"`: for each XLSX hidden fixture (`augmentedFixtures()`), no cell of any snapshot sheet holds the fixture's `marker`, and the sheet count equals `OpenResult.sheets.length`.
  - `"unfilled shape covers"`: on `unfilledShapeOverText`, under visible-only:
    - `coverage.kind === 'known'`, `covered(1, 1) === true` and `count === 1`;
    - the snapshot's B2 has no value but keeps its style id;
    - under `'full'`, `coverage === 'off'` and B2 holds `COVERED-B2`.
  - `"covered cells draw no CF"`: on `unfilledShapeOverText` with a data-bar rule over B1:B3, the snapshot's CF batch has no data-bar entry for B2 under visible-only, and has one under `'full'`.
  - `"rotated box adversary"`: on `rotatedBox`, every cell whose grid rect meets `x 40..60, y 30..50` is covered. The cell at (0,0)'s anchor-only neighbour beyond the rotated extent is not.
  - `"spill counts toward coverage"`: a left-aligned long text in A3, with B3 and C3 empty and a shape over C3 only, is covered. The same text right-aligned is not.
  - `"unknown coverage paints as authored"`: `unresolvedAnchor` gives `{kind: 'unknown', reason: 'unresolved-anchors'}`, and the snapshot sheet is the projected sheet itself (`===`). Build a `transformValid: false` node in the test, which gives `reason: 'transform'`.
  - `"saved values, no recompute"`: the snapshot of `savedFormula` has A1's saved value `7`, and the parse's `valuePolicy` is the saved-values policy. Record which policy `detectAndParse` uses today in the test's comment (S4 §6.2 step 7, "builder verifies").
  - `"clock is injected"`: for `cfToday` with clocks `{day: '2026-11-22', 'UTC'}` and `{day: '2026-11-23', 'UTC'}`, the CF paint batch's `dateContext.epochNowMs` equals `clockEpochMs(clock)`, and the evaluation `timeZone` equals `clock.timeZone`.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-snapshot.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** `coverage.ts`, `snapshot.ts` and the host method.
- [ ] **Step 4: Run green:** the same file, `cd server && bun test` and `bunx vitest run xlsx- interactive`, so that the `sheetAssets` change leaves local CF unchanged. Expected: all pass.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4): servable sheet snapshot — projected parse, base coverage, saved values, session clock" -- server/src/sheet/snapshot.ts server/src/sheet/coverage.ts server/src/render.ts src/components/xlsx/sheetAssets.ts server/test/sheet-snapshot.test.ts
  ```

### Task 6: `sheet-layout`: paint-facing geometry, chunked, within the disclosure boundary

**Files:**
- Create: `server/src/sheet/layout.ts`, `src/xlsx/view/streamGeometry.ts`
- Modify:
  - `src/stream/sheet-wire.ts`: layout types, encode and decode;
  - `server/src/ws.ts`: the `sheet-layout` op;
  - `server/test/harness.ts`: `Client.sheetLayout(reqId, sheetId)`, returning every part.
- Test: create `server/test/sheet-layout.test.ts`, `tests/stream/sheet-wire.test.ts`

**Interfaces:**
- Consumes: `SheetSnapshot` (Task 5), and tiles-mode gating (Task 4).
- Produces, in `sheet-wire.ts`:
  - **The layout type:**
    ```ts
    export interface StreamSheetLayout {
      sheetId: string; viewGen: number
      defaults: { rowPx: number; colPx: number }
      rows: AxisSizeInterval[]; cols: AxisSizeInterval[]
      collapsedRows: Array<{ start: number; end: number }>; collapsedCols: Array<{ start: number; end: number }>
      frozen?: { rows: number; cols: number; topRow?: number; leftCol?: number }
      merges: SourceRect[]; used?: SourceRect
      filters: Array<{ kind: 'autoFilter' | 'table'; range: SourceRect; header: 0 | 1; totals: 0 | 1 }>
      sort: null; filterSummary: Array<{ col: number; kind: string }>
      authorHidden: boolean; counts: { visibleRows: number; totalRows: number }
    }
    ```
  - **The message:** `{reqId, type: 'sheet-layout', seq, last, head?, rows?, cols?, collapsedRows?, collapsedCols?, merges?, filters?}`. `head` carries every scalar field, and appears on `seq 0` only.
  - **`encodeSheetLayout(l: StreamSheetLayout): object[]`**, which throws `LayoutTooLarge` over an R14 cap.
  - **`decodeSheetLayout(parts: unknown[]): StreamSheetLayout`**, which throws `LayoutInvalid` for:
    - a seq gap;
    - an unknown key at any depth;
    - a cap breached;
    - a non-integer;
    - a row index outside `[0, MAX_ROWS)` or a column index outside `[0, MAX_COLS)`.
  - `LAYOUT_LIMITS = { intervals: 65_536, merges: 65_536, filters: 1_024, parts: 64 }`.
- Produces `export function streamSheetGeometry(l: StreamSheetLayout): SheetGeometry` in `streamGeometry.ts`:
  - it is built with `createAxisMetrics` and `createSheetMetrics`, as `pagedGeometry.ts:44–48` does;
  - it has no cells (`cellsOf` returns `undefined`) and no drawings;
  - its merges come from `l.merges`.
- Server: `export function buildStreamLayout(s: SheetSnapshot, sheetId: string): StreamSheetLayout` in `layout.ts`. Row and column intervals that meet a collapsed run are cut out, so a hidden row's authored height never travels. In 4a:
  - `viewGen` is 0, `sort` is `null` and `filterSummary` is `[]`;
  - `filters` lists the declared autoFilter and table ranges;
  - `authorHidden` is true iff a collapsed row lies inside a declared autoFilter or table range;
  - `counts.totalRows` is `used ? used.r1 + 1 : 0`, and `counts.visibleRows` is `totalRows` minus the collapsed rows in `[0, totalRows)`.

  The request is `{reqId, type: 'sheet-layout', token, sheetId}`.

- [ ] **Step 1: Write the failing server tests.**
  - **`"layout round-trips and matches the server geometry"`:** for `grid500x100`, `millionRows` and a frozen-pane plus merges fixture:
    - `streamSheetGeometry(decodeSheetLayout(parts)).metrics.rows.start(i)` equals the server `buildSheetGeometry`'s for `i` in `[0, 1, 2, 9, 10, 20, 499, 65_535, 1_048_575]`;
    - `cols.start(j)` is equal for `j` in `[0, 1, 99, 16_383]`, and `rows.totalCss()` and `cols.totalCss()` are equal;
    - `frozen` and `merges` are deep-equal.
  - **`"hidden sizes never travel"`:** on `xlsx-hidden-rows` with `ht="40"` added to the hidden row, and `xlsx-hidden-cols` with a hidden width of 30:
    - no `rows` or `cols` interval covers the hidden index;
    - `collapsedRows` and `collapsedCols` name it.
  - **`"layout keys are on the allowlist"`:** the recursive key set of every part is a subset of the `StreamSheetLayout` field names plus `reqId/type/seq/last/head`. No `MARKER-` or `VIS-` string occurs in any part, for every XLSX hidden fixture.
  - **`"4a view fields"`:** on an autoFilter fixture with one collapsed row inside its range and one outside, `authorHidden === true`, `counts` deep-equals `{visibleRows: totalRows - 2, totalRows}`, `sort === null` and `filterSummary` is `[]`.
  - **`"layout chunks under 64 KiB"`:** a fixture with 20,000 merges gives ≥ 2 parts, each ≤ 65,536 UTF-8 bytes, with consecutive `seq` and a single `last`.
  - **`"over a cap is refused"`:** 70,000 merges give `{type: 'refused', code: 'layout-too-large'}`.
  - **`"layout is charged to unit once"`:**
    - 60 requests in a minute are answered and the 61st is refused `rate-limited`;
    - a multi-part answer's later parts charge nothing: after 30 four-part answers, 30 more `sheet-layout` requests are still answered;
    - an unknown `sheetId` gives `policy-refused`.
- [ ] **Step 2: Write the failing client tests** in `tests/stream/sheet-wire.test.ts`. `decodeSheetLayout` rejects each of:
  - a missing `seq 1`;
  - a `head` on `seq 1`;
  - an extra key `values`;
  - `rows[0].start = -1`;
  - `merges` over the cap;
  - a second `last`.
- [ ] **Step 3: Run** `cd server && bun test test/sheet-layout.test.ts` and `bunx vitest run tests/stream/sheet-wire.test.ts`. Expected: FAIL.
- [ ] **Step 4: Implement** the builder, the codec, the op (in the order token → `'unit'` rate → build → encode → send) and the client geometry.
- [ ] **Step 5: Run green:** the same commands, then `cd server && bun test`. Expected: all pass.
- [ ] **Step 6: Commit.**
  ```bash
  git commit -m "feat(stream-s4): sheet-layout — paint-facing geometry only, chunked, capped" -- server/src/sheet/layout.ts src/xlsx/view/streamGeometry.ts src/stream/sheet-wire.ts server/src/ws.ts server/test/harness.ts server/test/sheet-layout.test.ts tests/stream/sheet-wire.test.ts
  ```

### Task 7: Tile render on the pool, with the visible mark in global sheet space

**Files:**
- Create: `server/src/sheet/tile.ts`
- Modify:
  - `server/src/mark/visible.ts`: `paintMarkRegion`;
  - `server/src/render.ts`: `DocumentHost.renderTile`;
  - `server/src/worker/render-worker.ts`: a `tile` message;
  - `server/src/pool.ts`: `Pool.tile`.
- Test: create `server/test/sheet-tile.test.ts`

**Interfaces:**
- Consumes: `DocumentHost.sheetSnapshot` (Task 5), `SheetClock` (Task 4), and the Task 1 pins.
- Produces:
  - `export interface TileJob { sheetIndex: number; viewGen: number; zoomStep: number; dpr: 1 | 2; palette: TilePalette; col: number; row: number; mark: MarkSpec; clock: SheetClock }`
  - `export function paintMarkRegion(ctx: CanvasRenderingContext2D, mark: WatermarkOptions, region: { x: number; y: number; w: number; h: number; scale: number }): void`. It paints the mark's periodic layout, anchored at the sheet's device origin, clipped to `region` (device px). A tile is the 512² region at `(col·512, row·512)`.
  - `export const TILE_PALETTE_COLORS: Record<TilePalette, Pick<CanvasColors, 'pending' | 'pending-text'>>`, taken from the default tokens in `src/styles/tokens`.
  - `DocumentHost.renderTile(docKey: string, job: TileJob): Promise<{ png: Uint8Array; w: 512; h: 512 }>` (async, because the snapshot is). It rejects `RenderRefusal('bad-job')` when:
    - `sheetIndex` is out of range;
    - `zoomStep` is outside 0..10, or `dpr` is outside {1, 2};
    - `col` or `row` is negative, or beyond `ceil(extentDevicePx / 512)`;
    - `viewGen !== 0` (4a).
  - **The tile pass.** Take the logical box `{x: col·512/s, y: row·512/s, w: 512/s, h: 512/s}`, with `s = zoom·dpr`. Plan it with `planViewport` (profile `'D'`, a per-tile `frameBudgetFor('D')`), paint with `renderSheetViewport` on the snapshot sheet and geometry, apply `paintMarkRegion`, then encode the PNG metadata-free.
  - `Pool.tile(sessionId: string, job: { docRef: string; contentPolicy: 'visible-only' | 'full' } & TileJob, bytesEstimate: number, signal?: AbortSignal): Promise<{ png: Uint8Array; w: number; h: number }>`, with the same affinity, admission, deadline and recycle as `render`. `TILE_BYTES_ESTIMATE = 512 * 512 * 4 * 2`.

- [ ] **Step 1: Write the failing tests.**
  - `"a 512² tile, deterministic"`: `grid500x100` tile (0,0), zoomStep 5, dpr 1:
    - `decodePng` gives 512×512;
    - two renders are byte-equal;
    - it differs from the `blankSheet` tile.
  - `"edge tiles are full surfaces"` (R15): the tile past the 100th column is 512², blank beyond the extent.
  - `"the visible mark is continuous across a seam"`: with `mark: {kind: 'fixed'}` on `blankSheet`, tiles (0,0) and (1,0) composited side by side are pixel-equal to `paintMarkRegion` over `{x: 0, y: 0, w: 1024, h: 512, scale: 2}` on a blank 1024×512 surface. The same holds vertically for (0,0) and (0,1).
  - `"bad jobs are refused"`: `zoomStep 11`, `dpr 3`, `col -1`, `row` one past the grid, `sheetIndex 1` on a one-sheet doc, and `viewGen 1` each reject `RenderRefusal` with code `bad-job`.
  - `"a tile reads no clock"`: with `spyOn(Date, 'now').mockImplementation(() => { throw new Error('clock') })`, `cfToday` tile (0,0) renders. Its bytes differ between `day 2026-11-22` and a day on which A1's date is not today.
  - `"pool.tile goes through admission"`: with `caps {perSessionOutstanding: 1, perSessionQueue: 1}` and a `hangMs` job, a third concurrent `Pool.tile` for the sid rejects `AdmissionRefused('session-queue')`.
  - `"the million-row last row renders"`: the tile that holds row 1,048,575 of `millionRows` is 512² and not blank.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-tile.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same file, then `cd server && bun test` (purity included: `server/test/purity.test.ts` must still pass). Expected: all pass.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4): pool renders 512 device-px sheet tiles with a sheet-aligned visible mark" -- server/src/sheet/tile.ts server/src/mark/visible.ts server/src/render.ts server/src/worker/render-worker.ts server/src/pool.ts server/test/sheet-tile.test.ts
  ```

### Task 8: Audit: `tile_deliveries` and `sheet_views`

**Files:**
- Modify: `server/src/audit-store.ts` (the `SCHEMA` at `:56`, `expire`), `server/src/audit.ts`
- Test: `server/test/audit-store.test.ts`

**Interfaces:**
- Produces:
  - `export interface TileDeliveryRecord { sid: string; sheetId: string; viewGen: number; zoomStep: number; dpr: 1 | 2; col: number; row: number; status: 'l1' | 'unmarked'; reason?: string; refSha256: string; at: number }`
  - `export interface SheetViewRecord { sid: string; sheetId: string; viewGen: number; spec: { sort: null; filters: []; sizes: [] }; clock: SheetClock; at: number }`
  - **Tables:**
    - `tile_deliveries (… UNIQUE (sid, sheet_id, view_gen, zoom_step, dpr, col, row, ref_sha256, status))`, deduplicated with `INSERT OR IGNORE` as `deliveries` is;
    - `sheet_views (… UNIQUE (sid, sheet_id, view_gen))`, with `spec` and `clock` as JSON.
  - **`AuditStore` methods:** `putTileDelivery`, `tileDeliveries(sid)`, `putSheetView`, `sheetView(sid, sheetId, viewGen)` and `tileCandidates(docRef, sheetId, zoomStep, dpr, window: { c0: number; c1: number; r0: number; r1: number })`. The last returns `Array<{ sid; row: StoredRow; delivery: TileDeliveryRecord }>`.
  - **`AuditTable` methods:** `recordTileDelivery(r)` and `recordSheetView(r)`. Both throw for an unknown or expired sid, as `recordCopy` does.

- [ ] **Step 1: Write the failing tests.**
  - Recording a tile delivery twice with equal fields stores one row, and a different `refSha256` adds a second.
  - `tileCandidates` returns only rows inside the window, for this docRef, sheet and pair.
  - A second `recordSheetView` for the same `(sid, sheetId, viewGen)` is ignored.
  - `sheetView` round-trips `clock`.
  - `expire(cutoff)` removes both tables' old rows.
  - An unknown sid throws.
- [ ] **Step 2: Run** `cd server && bun test test/audit-store.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement**, following `putDelivery`/`candidates`.
- [ ] **Step 4: Run green**, then `cd server && bun test`.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4): audit tile deliveries and the sheet view they were drawn under" -- server/src/audit-store.ts server/src/audit.ts server/test/audit-store.test.ts
  ```

### Task 9: The `tiles` op: `XLSX_TILE`, `tiles-done`, `'tile'` rate, cache key, Layer-1, cancel

**Files:**
- Create: `server/src/sheet/tiles-op.ts` (the per-sid batch driver)
- Modify:
  - `server/src/ws.ts`: dispatch; `KIND.XLSX_TILE = 5`; `ENCODED[5] = 8 << 20`; teardown and resume stop batches;
  - `server/src/rate.ts`: `'tile'`;
  - `server/src/cache.ts`: `RenderKeyFields.tile`, `policyDetails.viewStateHash`;
  - `src/stream/sheet-wire.ts`: `encodeTileBody`/`decodeTileBody`;
  - `server/test/harness.ts`: `Client.tiles(reqId, req)`, returning the decoded tiles and `tiles-done`; `CAPS` gains kind 5;
  - `server/src/ws.ts`, again: `StreamServerOptions.rates?: Partial<Record<'message' | 'tile', { perWindow: number; windowMs: number }>>`, TEST and measure only (Task 17 uses both members).
- Test: create `server/test/sheet-tiles-wire.test.ts`; modify `server/test/cache.test.ts`

**Interfaces:**
- Consumes: `Pool.tile` (Task 7), the audit methods (Task 8), `sessionDoc.clock`/`sheetMode` (Task 4), and the existing `markUnit` pattern (`ws.ts:1179`).
- Produces:
  - **The request:** `{reqId, type: 'tiles', token, sheetId, viewGen, zoomStep, dpr, palette, tiles: Array<{col, row}>}`, with 1 to 64 tiles and no duplicates.
  - **The reply:** one kind-5 message per delivered tile under `reqId`, then `{reqId, type: 'tiles-done', refused: Array<{col, row, code: 'rate-limited' | 'policy-refused'}>, retryAfterMs?: number, cancelled?: true}`.
    - A whole-request failure is one `refused` message instead: `policy-refused` for a malformed request, `stale-view` for an old `viewGen`.
  - **Cancel:** `{reqId, type: 'sheet-cancel', target}` stops starting tiles for `target`; its `tiles-done` carries `cancelled: true`.
  - **The tile body codec:** `encodeTileBody(t: { sheetIndex: number; viewGen: number; zoomStep: number; dpr: 1 | 2; col: number; row: number; png: Uint8Array; w: number; h: number }): Uint8Array` (R13), and `decodeTileBody(bytes: Uint8Array)`, which returns the same shape or throws `InvalidOperand`.
  - **The rate:** `SessionRates.takeTiles(n: number, now: number): { admitted: number; retryAfterMs: number }`, with `TILE_RATE = { perWindow: 600, windowMs: 10_000, burst: 120, burstMs: 1_000 }`. It never touches the breach window.
  - **The view-state hash:** `viewStateHash(v: { sort: null; filters: []; sizes: []; authorHiddenTier: 'none'; clock: SheetClock }): string`, the sha256 hex of the canonical JSON. It is put in `policyDetails`.
  - **The tile key field:** `RenderKeyFields.tile?: { sheetIndex: number; zoomStep: number; dpr: 1 | 2; palette: TilePalette; col: number; row: number }`, with `scaleBucket = dpr`. When `tile` is absent the field is not hashed, so unit keys keep their hashes.
  - **The driver:**
    1. validate, then check the token;
    2. `takeTiles`;
    3. write the `sheet_views` row once per `(sid, sheetId, viewGen)`;
    4. for each admitted tile, keeping at most `caps.perSessionOutstanding` in flight: cache, else `pool.tile`; Layer-1 on basis label `sheet:<sheetId>@<zoom>x<dpr>:<col>,<row>` (or the Task 1 crop); `recordTileDelivery`; frame; send;
    5. send `tiles-done`.

- [ ] **Step 1: Write the failing tests** in `sheet-tiles-wire.test.ts`, on a tiles session, forensic on, `visibleMark: 'fixed'`, `grid500x100`.
  - **`"happy path"`:** 4 tiles at zoomStep 5, dpr 1 give 4 kind-5 messages.
    - Each `decodeTileBody` has `sheetIndex 0`, `viewGen 0`, the requested `col`/`row`, and a 512² PNG.
    - `tiles-done` has `refused: []`.
    - `tileDeliveries(sid)` has 4 `l1` rows, and `sheetView(sid, '0', 0).clock` equals `sheetClockOf(sid)`.
  - **`"malformed batches are refused whole and charge nothing"`:** each of the following gives `policy-refused`:
    - 65 tiles, or a duplicate coordinate;
    - `zoomStep: 11`, `dpr: 3` or `palette: 'blue'`;
    - `col` past the grid, or `sheetId: '9'`;
    - a missing `token`.

    `viewGen: 1` gives `stale-view`. A following batch of 64 is fully admitted, with `refused: []`.
  - **`"tile rate refuses per tile, never closes"`:** with the real `TILE_RATE` and an injected `now`, batches of 64 then 64 inside 1 s leave 8 refused `rate-limited` in the second `tiles-done`, with `retryAfterMs > 0`.
    - Ten more such batches leave the socket open, with no 4008.
    - With `now` advanced by `retryAfterMs`, the refused 8 are admitted.
  - **`"cancel stops the batch"`:** after a 64-tile batch and `sheet-cancel {target}`, at most 8 kind-5 messages arrive after the cancel, and `tiles-done.cancelled === true`.
  - **`"same clock shares canonical tiles; marked bytes differ"`:** two sids on one server and one day request tile (0,0).
    - A counting `pool.tile` wrapper is called once.
    - The two delivered PNGs differ.
    - Both `refSha256` are equal.
  - **`"a different day or zone never shares"`:** each of these causes one more `pool.tile` call:
    - a third sid opened after `now` advances one day;
    - a fourth with `policy.timeZone: 'Asia/Tokyo'` at a UTC instant that falls on a different Tokyo date;
    - a fifth with `policy.timeZone: 'Europe/London'` on the same `day` as the UTC sids (November: same date, different zone).
  - **`"resume stops the old socket's batch"`:** a batch in flight when the sid resumes sends nothing more on the old socket.
- [ ] **Step 2: Add a unit-key golden** in `cache.test.ts`: `renderKey` of a fixed unit-key object equals the hex string printed before the change, pasted in.
- [ ] **Step 3: Run** `cd server && bun test test/sheet-tiles-wire.test.ts test/cache.test.ts`. Expected: FAIL.
- [ ] **Step 4: Implement.**
- [ ] **Step 5: Run green:** the same files, then `cd server && bun test` and `npm run -s typecheck:server`. Expected: all pass. Report the test count.
- [ ] **Step 6: Commit.**
  ```bash
  git commit -m "feat(stream-s4): tiles op — XLSX_TILE per tile, per-tile rate, view-state cache key, Layer-1 per tile" -- server/src/sheet/tiles-op.ts server/src/ws.ts server/src/rate.ts server/src/cache.ts src/stream/sheet-wire.ts server/test/harness.ts server/test/sheet-tiles-wire.test.ts server/test/cache.test.ts
  ```

### Task 10: Tracer and detection on tiles

**Files:**
- Modify: `server/src/detect/trace.ts` (adds `traceTiles`), `server/scripts/detect.ts` (adds `--tiles`)
- Test: create `server/test/trace-tiles.test.ts`; modify `server/test/detect-cli.test.ts`

**Interfaces:**
- Consumes: `tileCandidates`/`sheetView` (Task 8), `DocumentHost.renderTile` (Task 7), and the existing `register`/`scoreRegistered`/`decideScores`/`tauFor`.
- Produces:
  - `export interface TileTraceInput { capture: Bitmap; doc: Uint8Array; sheetId: string; zoomStep: number; dpr: 1 | 2; window: { c0: number; c1: number; r0: number; r1: number }; keys: ReadonlyMap<number, Keys>; store: AuditStore; host?: DocumentHost }`
  - `export async function traceTiles(input: TileTraceInput): Promise<TraceAnswer>`.
    - It follows the R10 composite, with one reference per `(contentPolicy, visibleMark, clock)` key.
    - It rebuilds through `host.renderTile` with the row's mark and the `sheet_views` clock.
    - A `refSha256` mismatch gives `no-signal: reference-mismatch`. The other answers are those of `trace`.
  - **CLI:** `bun scripts/detect.ts --images cap.png --key-file k.json --doc f.xlsx --tiles <sheetId>,<zoomStep>,<dpr>,<c0>-<c1>,<r0>-<r1> --ids ids.json`.

- [ ] **Step 1: Write the failing tests.**
  - `"a one-pane capture traces to its sid"`: three forensic sids fetch window (0..1, 0..1) of `grid500x100`. Sid B's four tiles, composited and cropped to 900×900 at (40,40), give `{kind: 'traced', sid: B}`. The same capture through JPEG q90 gives the same answer.
  - `"a blank pane abstains"`: the same, on `blankSheet`, gives `{kind: 'no-signal', reason: 'flat'}`.
  - `"the rebuild uses the stored clock"`: on `cfToday`, deliveries made under day D trace to their sid while the host's clock reads D+1.
  - `"a tampered reference fails closed"`: with one row's `ref_sha256` edited, the answer is `no-signal: reference-mismatch`.
  - CLI: `--tiles 0,5,1,0-1,0-1` prints the traced sid.
- [ ] **Step 2: Run** `cd server && bun test test/trace-tiles.test.ts test/detect-cli.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same files, plus `cd server && bun test test/trace.test.ts`, which must be unchanged.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4): trace a pane capture through rebuilt tile references" -- server/src/detect/trace.ts server/scripts/detect.ts server/test/trace-tiles.test.ts server/test/detect-cli.test.ts
  ```

### Task 11: Client session: opt-in HELLO, `sheetLayout`, `tiles`, hostile-server drops, resume

**Files:**
- Create: `src/stream/remote/sheet-channel.ts`
- Modify:
  - `src/stream/remote/session.ts`: HELLO at `:1436`; the per-kind caps at `:87` gain `[Kind.XLSX_TILE]: [8 << 20, 64 << 20]`; resume;
  - `src/stream/remote/remote-doc.ts`: adds `sheets`;
  - `src/stream/remote/open.ts`, `src/components/OfficeFile.tsx` (`:42`, `:122`), `src/stream/index.ts`.
- Test: create `tests/stream/remote-sheet.test.ts`; modify `tests/stream/remote-session.test.ts`

**Interfaces:**
- Consumes: the codecs (Tasks 6 and 9), and the manifest fields (Task 4).
- Produces:
  - **Opt-in plumbing:** `OfficeFileProps.streamSheets?: 'tiles'` becomes `OpenRemoteOptions.streamSheets`, then `SessionOptions.features?: readonly string[]`. The HELLO carries `features` only when the list is non-empty, and a resume sends the same list.
  - **`StreamSession.sheetLayout(sheetId: string, signal?: AbortSignal): Promise<StreamSheetLayout>`.**
  - **`StreamSession.tiles(req: { sheetId: string; viewGen: number; zoomStep: number; dpr: 1 | 2; palette: TilePalette; tiles: Array<{ col: number; row: number }> }, onTile: (t: DecodedTile) => void, signal?: AbortSignal): Promise<{ refused: Array<{ col: number; row: number; code: string }>; retryAfterMs?: number; cancelled?: boolean }>`.**
    - Aborting it sends `sheet-cancel`.
    - `DecodedTile = { sheetIndex; viewGen; zoomStep; dpr; col; row; png: Uint8Array }`.
  - **`RemoteOfficeDoc.sheets?: { list: ReadonlyArray<{ sheetId: string; name: string }>; layout: StreamSession['sheetLayout']; tiles: StreamSession['tiles'] }`.** It is present only when `manifest.sheetMode === 'tiles'`.
  - **React-free exports** from the lazy stream runtime `src/stream/index.ts`: `decodeTileBody`, `decodeSheetLayout` and the `sheet-wire` constants. This is the tile protocol client of S4 §6.11. `streamSheetGeometry` is not exported, because it would pull `src/xlsx/view` engine modules into the runtime.

- [ ] **Step 1: Write the failing tests**, using the fake-socket pattern of `tests/stream/remote-print.test.ts`.
  - `"default HELLO is Stage 3's"`: the HELLO's keys are exactly `['client', 'recVersions', 'reqId', 'type']`, plus `token` when given.
  - `"opt-in HELLO carries the feature"`: `features` deep-equals `['sheet-tiles@1']`.
  - `"no sheetMode from the server → units"`: opted in, against a manifest without `sheetMode`, gives `doc.sheets === undefined`.
  - `"tiles deliver and finish"`: two kind-5 messages then `tiles-done` call `onTile` twice and resolve `{refused: []}`.
  - **`"hostile tiles are dropped before decode"`** (Review Focus 5). Each of the following is dropped without throwing; `onTile` is not called for it, and the session state stays `'open'`:
    - a kind-5 on an unknown `reqId`;
    - an unrequested coordinate;
    - a duplicate;
    - `viewGen` off by one;
    - `sheetIndex` off by one;
    - `w = 256`;
    - a tile after `tiles-done`;
    - a kind-5 over the 8 MiB encoded cap.
  - `"abort sends sheet-cancel"`: the sent message is `{type: 'sheet-cancel', target: <reqId>}`.
  - `"resume mid-batch"` (Review Focus 1): with 3 of 6 tiles delivered when the socket drops, the client re-sends one `tiles` request for exactly the 3 missing coordinates under the new token. A late tile from the old socket is dropped, and the promise resolves with all 6 delivered once each.
- [ ] **Step 2: Run** `bunx vitest run tests/stream/remote-sheet.test.ts tests/stream/remote-session.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** `bunx vitest run tests/stream`, then `bun run gates:core`, recording the bytes.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4): client tile and layout channel, opt-in sheet-tiles feature" -- src/stream/remote/sheet-channel.ts src/stream/remote/session.ts src/stream/remote/remote-doc.ts src/stream/remote/open.ts src/components/OfficeFile.tsx src/stream/index.ts tests/stream/remote-sheet.test.ts tests/stream/remote-session.test.ts
  ```

### Task 12: `StreamWorkbook`: a widened source that holds no cell data

**Files:**
- Create: `src/components/xlsx/viewerSource.ts`, `src/components/xlsx/streamWorkbook.ts`
- Modify:
  - `src/worker/viewerRemoteSource.ts`: `RemoteWorkbook` gains `readonly kind = 'paged'`; `isRemoteWorkbook` at `:182` is deleted;
  - `src/components/xlsx/InteractiveSheetDoc.tsx`: `:99`, `:444` and every `remote.*` site, narrowed per R16.
- Test: create `tests/xlsx-stream-workbook.test.ts`

**Interfaces:**
- Consumes: `RemoteOfficeDoc.sheets` (Task 11) and `streamSheetGeometry` (Task 6).
- Produces:
  - `export type SourceKind = 'paged' | 'stream'`
  - `export interface ViewerSheetSource extends WorkbookSource { readonly kind: SourceKind }`, plus the R16 shared members, with their signatures copied from `RemoteWorkbook`.
  - `export function sheetSourceOf(source: WorkbookSource | undefined): { paged: RemoteWorkbook | null; stream: StreamWorkbook | null }`.
  - `export class StreamNoCellData extends Error { readonly code = 'stream-no-cell-data'; constructor(readonly method: string) }`
  - `export class StreamSheetStore implements SourceStore`. It carries `stamp` (valueGen 0, fontGen 0), `extent`, `completeness: 'complete'` and `valuePolicy`, and every content method rejects `StreamNoCellData`.
  - `export class StreamWorkbook implements ViewerSheetSource`:
    - `readonly kind = 'stream'`;
    - `constructor(sheets: NonNullable<RemoteOfficeDoc['sheets']>, docRef: string)`;
    - `activate(sheetId: string): Promise<void>`, which fetches and holds the layout;
    - `layout(sheetId): SheetLayout`, which maps `collapsed*` to `hidden`/`hiddenCols` and sets every `has*` false;
    - `identity(sheetId): CacheIdentity`, which is `{docId: docRef, sheetId, valueGen: 0, fontGen: 0, viewGen}`;
    - `tiles(): NonNullable<RemoteOfficeDoc['sheets']>['tiles']`.
    It has no `ViewerClient` and no row, overflow, CF, drawing or sparkline cache.
  - `export function streamShellDoc(list: ReadonlyArray<{ sheetId: string; name: string }>): XlsxDocument`: names only.

- [ ] **Step 1: Write the failing tests.**
  - `"every content method rejects stream-no-cell-data"`: each of `readCell`, `readRows`, `queryColumn`, `scanPopulated`, `scanOccupied`, `queryStructural`, `queryOverview` and `ensureComplete` rejects with `code === 'stream-no-cell-data'` and the right `method`.
  - `"stream reports no paged content"`:
    - `hasSparklines`, `hasConditionalFormatting`, `hasRegionStyles` and `hasPivotDisplay` are false for every sheet;
    - `pivotTableCount` is 0;
    - `store(id).stamp.valueGen === 0`;
    - `sort` and `filter` reject `stream-no-cell-data` (until 4c).
  - `"no ViewerClient, not a RemoteWorkbook"`: `expectTypeOf<StreamWorkbook>().not.toHaveProperty('client')`, and a `// @ts-expect-error` assignment to `RemoteWorkbook`.
  - `"zero-content shell"`: every sheet of `streamShellDoc` has empty rows, cells, styles and drawings.
  - `"layout maps collapsed runs"`: a decoded layout with `collapsedRows [{start: 1, end: 1}]` gives `layout().hidden` deep-equal to `[{r0: 1, r1: 1}]`.
  - `"isRemoteWorkbook is gone"`: a source scan of `src/` finds no `isRemoteWorkbook`.
- [ ] **Step 2: Run** `bunx vitest run tests/xlsx-stream-workbook.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement** the interface, the classes, and the viewer narrowing (paged-only members under `paged`, shared members under either).
- [ ] **Step 4: Run green:**
  - `bunx vitest run tests/xlsx-stream-workbook.test.ts`;
  - the local and paged viewer suites, unmodified: `bunx vitest run xlsx- interactive paged sheet viewer tests/components` (vitest name filters);
  - `npx tsc --noEmit -p .`;
  - `bun run gates:core`.

  Expected: all pass, and the gate bytes are recorded.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4): StreamWorkbook — a geometry-only viewer source whose store rejects cell reads" -- src/components/xlsx/viewerSource.ts src/components/xlsx/streamWorkbook.ts src/worker/viewerRemoteSource.ts src/components/xlsx/InteractiveSheetDoc.tsx tests/xlsx-stream-workbook.test.ts
  ```

### Task 13: Renderer swap: pending tiles, requests, preset zoom, dpr bucket

**Files:**
- Create: `src/components/xlsx/streamTiles.ts`
- Modify:
  - `src/components/xlsx/sheetTiles.ts`: `SheetFrame.stream`; `tile()` at `:359` gets a stream branch; `present()` at `:180–226` skips `paintWatermark`;
  - `src/xlsx/view/sheetZoom.ts`: `nearestZoomStep`;
  - `src/components/xlsx/InteractiveSheetDoc.tsx`: the stream frame (`frame.stream`), the zoom settle and the dpr bucket.
- Test: create `tests/xlsx-stream-tiles.test.ts`

**Interfaces:**
- Consumes: `StreamWorkbook.tiles()` (Task 12), and `TileKey`/`tileRangeFor` (`tileCache.ts:333`, `:439`).
- Produces:
  - `export function nearestZoomStep(z: number): SheetZoom`. It returns the nearest `SHEET_ZOOM_STEPS` value, ties going to the smaller step.
  - `export function dprBucket(dpr: number): 1 | 2`.
  - `export interface StreamTileSource { request(key: TileKey, visible: boolean): void; take(key: TileKey): CanvasTile | undefined; retain(visible: TileIndexRange, ring: TileIndexRange): void; onArrive(cb: () => void): () => void; readonly stats: { requested: number; outstanding: number; dropped: number } }`.
  - `export function createStreamTileSource(o: { sheetId: string; viewGen: () => number; palette: () => TilePalette; tiles: NonNullable<RemoteOfficeDoc['sheets']>['tiles']; decode: (png: Uint8Array) => Promise<CanvasTile>; now?: () => number; setTimer?: typeof setTimeout }): StreamTileSource`.
    - Requests are batched per frame, at most 64 per batch and at most 2 batches in flight, with visible tiles first.
    - `retain` cancels a batch none of whose tiles is visible or in the ring any more.
    - A `rate-limited` tile is re-queued after `retryAfterMs`.
  - `SheetFrame.stream?: StreamTileSource`. When it is set:
    - `tile()` never calls `renderSheetViewport`;
    - a missing tile is requested and painted with the pending fill, and is not cached;
    - an arrived tile is inserted under the same `TileKey` (`identity.viewGen` from the server);
    - `present()` never stamps the watermark.

- [ ] **Step 1: Write the failing tests.**
  - `"a missing tile is requested, not painted"`: one `present()` gives:
    - `stats.basePaints === 0`;
    - a `vi.spyOn` of `renderSheetViewport` with 0 calls;
    - `stats.pendingTiles` equal to the visible count;
    - each visible key requested once.
  - `"an arriving tile is cached and a repaint scheduled"`: `onArrive` fires once, and the next `present()` has `tileHits === 1`.
  - `"visible before the ring"`: the first batch's coordinates are all in the visible range.
  - `"bounded in flight"`: 300 keys give ≤ 2 outstanding batches, each ≤ 64.
  - `"scrolling away cancels"`: after `retain` with disjoint ranges, the earlier batch's signal is aborted.
  - `"rate-limited tiles are retried, none stuck pending"` (Review Focus 2): with a fake clock and a `tiles-done` refusing 3 with `retryAfterMs: 500`, the 3 are re-requested at 500 ms and arrive. After the run, `stats.outstanding === 0`.
  - `"no watermark on a stream frame"`: with `frame.watermark` set, `paintWatermark` (spied) has 0 calls.
  - `"zoom steps and dpr buckets"`:
    - `nearestZoomStep(0.8) === 0.75`, `nearestZoomStep(0.29) === 0.25` and `nearestZoomStep(9) === 4`;
    - `dprBucket(0.9) === 1`, `dprBucket(1) === 1`, `dprBucket(1.25) === 2` and `dprBucket(3) === 2`.
  - `"theme switch"` (Review Focus 3): after a palette change, the new keys are requested, and a late old-palette tile is never returned by `take` under a new-palette key.
- [ ] **Step 2: Run** `bunx vitest run tests/xlsx-stream-tiles.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
  - The viewer passes `frame.stream` when the source is a stream.
  - Wheel and pinch on a stream settle to `nearestZoomStep` at gesture end; during the gesture the last pane is CSS-scaled.
  - The pane surface uses `dprBucket(devicePixelRatio)`, and CSS scales the pane once.
- [ ] **Step 4: Run green:** the same file, the unmodified `bunx vitest run xlsx- interactive paged sheet viewer`, and `bun run gates:core` with bytes recorded.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4): sheet renderer requests server tiles; preset zoom and dpr buckets on streams" -- src/components/xlsx/streamTiles.ts src/components/xlsx/sheetTiles.ts src/xlsx/view/sheetZoom.ts src/components/xlsx/InteractiveSheetDoc.tsx tests/xlsx-stream-tiles.test.ts
  ```

### Task 14: Mount the viewer in tiles mode: tabs, frozen panes, positions-only a11y, print

**Files:**
- Create: `src/components/remote/RemoteSheetView.tsx`, `tests/browser/stream-sheet.spec.ts`
- Modify:
  - `src/components/remote/RemoteDocView.tsx`: lazy route at `:160`;
  - `src/components/xlsx/InteractiveSheetDoc.tsx`, for stream mode:
    - no selection, find, copy, or sort and filter menus;
    - the a11y grid has positions only and shows `TEXT_UNAVAILABLE` once;
    - print goes through `RemoteOfficeDoc.print`, with the picker in sheets;
  - `server/test/app.ts`: `?doc=million-rows|grid500x100&sheets=tiles`.
- Test: create `tests/stream/remote-sheet-view.test.tsx`

**Interfaces:**
- Consumes: `StreamWorkbook` (Task 12), the stream renderer (Task 13) and `RemoteOfficeDoc.print` (Stage 3b).
- Produces: `RemoteSheetView({ document: RemoteOfficeDoc, …OfficeDocProps })`, which mounts `InteractiveSheetDoc` with `doc = streamShellDoc(...)` and `source = new StreamWorkbook(...)`. `RemoteDocView` renders it, lazily, iff `document.sheets` is set.

- [ ] **Step 1: Write the failing component tests.**
  - `"tiles mode mounts the interactive viewer"`:
    - `[role=grid]` is present;
    - there is no `[data-unit]` canvas;
    - there is no Find or Copy control;
    - the status region text equals `'text unavailable in this viewing mode'` exactly once.
  - `"a11y grid has positions, no text"`: gridcells carry `aria-rowindex`/`aria-colindex`, and every `textContent === ''`.
  - `"units mode unchanged"`: a manifest without `sheetMode` renders the Stage 3 `[data-unit]` canvases.
  - `"sheet tabs fetch the layout once per activation"`: switching to sheet `'1'` calls `sheets.layout('1')` once.
- [ ] **Step 2: Write the failing browser tests** in `stream-sheet.spec.ts`, Chromium, against the harness.
  - `"million-row sheet scrolls to row 1,048,576"`: after setting the sheet scroller's `scrollTop` to its maximum (no keyboard: selection is 4b), the row header shows `1048576` and, after settle, no pending tiles remain (`window.__ovSheetStats.pendingTiles === 0`).
  - `"no too-large notice"`: the page text excludes `LARGE_EXPORT_MESSAGE`.
  - `"no local painting"`: `window.__ovSheetStats.basePaints === 0`.
  - `"authored frozen panes stay pinned"`: on a frozen fixture, the pinned header row stays at y 0 after scrolling 2,000 px.
  - `"wheel zoom settles to a preset step"`: Ctrl+wheel leaves the zoom label in `SHEET_ZOOM_STEPS`.
  - `"print prints the static unit"`: following `stream-print.spec.ts`, the print of sheet 1 gives one page.
- [ ] **Step 3: Run** `bunx vitest run tests/stream/remote-sheet-view.test.tsx` and `IV_HARNESS_PORT=4391 npm run test:browser:stream`. Expected: FAIL.
- [ ] **Step 4: Implement.** `window.__ovSheetStats` is set only by the harness page, through the existing test hook pattern.
- [ ] **Step 5: Run green:** the same commands, the unmodified `npx playwright test --config playwright.viewer.config.ts tests/browser/interactive-viewer.spec.ts`, and `bun run gates:core`, which must show the viewer only in a lazy chunk.
- [ ] **Step 6: Commit.**
  ```bash
  git commit -m "feat(stream-s4): streamed XLSX opens in the interactive viewer in tiles mode" -- src/components/remote/RemoteSheetView.tsx src/components/remote/RemoteDocView.tsx src/components/xlsx/InteractiveSheetDoc.tsx server/test/app.ts tests/stream/remote-sheet-view.test.tsx tests/browser/stream-sheet.spec.ts
  ```

### Task 15: Evidence: the 4a wire audit, T2 on tiles, coverage pixels, marks over the wire

**Files:**
- Create: `server/test/sheet-wire-audit.test.ts`, `server/test/sheet-t2.test.ts`
- Modify: `server/test/sheet-fixtures.ts`, adding `sentinels()`:
  - `SENTINEL_F` in a formula `=LEN("SENTINEL_F")` with a saved value;
  - `1234.5678` formatted `0`;
  - a defined name `SENTINEL_N`;
  - a comment;
  - a hidden drawing whose image is a unique 1 KiB PNG.

**Interfaces:**
- Consumes: the whole 4a server path through the harness `Client`.

- [ ] **Step 1: Write the wire audit.** Capture every server→client byte of a tiles session that does `sheet-layout` and fetches every tile of the used extent at (5, 1) and (5, 2). Decode the manifest, every text message and every kind-5 body. Run this for `sentinels()` and each XLSX hidden fixture.
  - `"no sentinel on the wire"`: no decoded text contains `SENTINEL_F`, `SENTINEL_N`, `1234.5678`, `=LEN`, any fixture `marker`, or the hidden sheet's name.
  - `"no drawing byte on the wire"`: the drawing PNG's bytes do not occur in any decoded payload (PNG `IDAT` included).
- [ ] **Step 2: Write T2 on tiles** (forensic off, `visibleMark: 'off'`). For each XLSX hidden fixture, every tile of `withHidden` at both pairs is PNG-byte-equal to the same tile of `withoutHidden`.
  - `"base-covered paints empty under visible-only"`: the `unfilledShapeOverText` tile equals the tile of the fixture with B2 deleted.
  - `"…and as authored under full"`: under `'full'` it differs from it.
  - `"rotated box"`: the `rotatedBox` tile equals the fixture with its covered cells deleted.
  - `"unknown coverage paints as authored"`: the `unresolvedAnchor` tile equals a local-path `renderSheetViewport` tile of the projected sheet.
  - `"saved values"`: the `savedFormula` tile equals the `plainSeven` tile.
- [ ] **Step 3: Write the marks-over-the-wire test.** `"wire tiles keep the mark continuous"`: with `visibleMark: 'fixed'` and forensic off, tiles (0,0) and (1,0) from the wire, composited, equal `paintMarkRegion` over 1024×512 on `blankSheet`.
- [ ] **Step 4: Run** `cd server && bun test test/sheet-wire-audit.test.ts test/sheet-t2.test.ts`. Expected: PASS, since Tasks 3–9 provide the behaviour. These are evidence, not RED-first. A failure is a defect in the owning task: fix it there, with a failing test first.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "test(stream-s4): wire audit and T2 on tiles; coverage and mark evidence" -- server/test/sheet-wire-audit.test.ts server/test/sheet-t2.test.ts server/test/sheet-fixtures.ts
  ```

### Task 16: Compatibility matrix over real builds

**Files:**
- Create: `scripts/stream-s4/compat-matrix.ts`

**Interfaces:**
- Consumes: R11, the Task 2 baseline and the Task 4 handshake.
- Produces: `bun scripts/stream-s4/compat-matrix.ts --base 53668e81 [--keep]`, which prints one PASS/FAIL line per leg and exits non-zero on any FAIL.

- [ ] **Step 1: Write the script.**
  1. Create the detached worktree for `--base` in the scratch temp dir, and symlink `node_modules`.
  2. Start the base server and the current server from source on free ports, with demo hooks serving `grid500x100` and a 2-sheet fixture.
  3. Run the legs:
     - **(a)** the base `StreamSession` against the current server: the manifest has no `sheetMode`; a unit, a search hit, a copy and a print page arrive;
     - **(b)** the current client with `features: ['sheet-tiles@1']` against the current server: `sheetMode === 'tiles'`, `sheet-layout` decodes, and one tile decodes;
     - **(c)** the current client, opted in, against the base server: no `sheetMode`, units arrive, and the base server accepted the HELLO;
     - **(d)** a raw HELLO with 17 features against the current server closes `unsupported-version`.
  4. Remove the worktree unless `--keep` is given.
- [ ] **Step 2: Run** `bun scripts/stream-s4/compat-matrix.ts --base 53668e81`. Expected: four PASS lines.
- [ ] **Step 3: Run the Stage 3 XLSX suites, unchanged:** `cd server && bun test test/xlsx-units-baseline.test.ts test/text-search-wire.test.ts test/text-search-bound.test.ts test/text-a11y-wire.test.ts test/text-served.test.ts test/print-wire.test.ts`. Expected: PASS, and `git diff --stat 53668e81 -- <those files>` is empty except for `xlsx-units-baseline.test.ts`, which is new.
- [ ] **Step 4: Commit.**
  ```bash
  git commit -m "test(stream-s4): compatibility matrix — old/new client × old/new server" -- scripts/stream-s4/compat-matrix.ts
  ```

### Task 17: T8 p95, corpus, gates, validation and spec amendments

**Files:**
- Create: `scripts/stream-s4/tile-p95.ts`, `server/test/sheet-corpus.test.ts`, `docs/validation/streaming-s4a.md`
- Modify:
  - `server/src/ws.ts`: `StreamServerOptions.onTileTiming?(t: { t0: number; t1: number; zoomStep: number; dpr: number })`, called with R3's instants. It is TEST and measure only;
  - the parent spec and S3 spec, as listed in Step 4;
  - the S4 spec: an "As built (4a)" note under §11.2.

**Interfaces:**
- Consumes: everything above.

- [ ] **Step 1: Write the corpus test.** `"every corpus sheet tiles without a static unit"`: for each corpus XLSX in a tiles-mode host:
  - every visible sheet's layout decodes;
  - tile (0,0) at (5, 1) renders 512²;
  - afterwards `paintablesBuilt === 0`, so neither `sheetPaintable` nor the too-large refusal paintable ran;
  - the test asserts that it visited at least one corpus XLSX.
- [ ] **Step 2: Write and run the measurement** `bun scripts/stream-s4/tile-p95.ts`. It follows R3 on `grid500x100` with `visibleMark: 'fixed'` and forensic on, with the `'message'` and `'tile'` rates raised so neither refuses. It reports p50, p95 and max at dpr 1 and dpr 2, the `--batch 64` throughput and the loopback figure.
  - Expected: p95 ≤ 120 ms at both. If it is red, stop and report; do not tune the gate.
- [ ] **Step 3: Run every gate:**
  - `cd server && bun test`;
  - `bunx vitest run`;
  - `IV_HARNESS_PORT=4391 npm run test:browser:stream`, and the unmodified viewer browser suites;
  - `npx tsc --noEmit -p .`, `npm run -s typecheck:server` and `npm run -s lint`;
  - `bun run gates:core`;
  - `bun scripts/stream-s4/compat-matrix.ts --base 53668e81`.

  Expected: green, except the size gate, which stays as red as at the Task 2 baseline within ±4 B. The `dist/core.js` purity and the module map are ok.
- [ ] **Step 4: Write the amendments.**
  - **Parent spec, the S4 §10 items 4a implements:**
    - §6.2 rule 1: HELLO `features` (item 9);
    - §6.4: the coverage clause of item 1, and item 10, the policy parameter;
    - §12: item 3 as it applies to tiles (base-covered cells paint empty);
    - §14: the `'tile'` class with R4's numbers, `XLSX_TILE` answering one request with several messages, and the R13 body (item 4, tile part);
    - §9.3: `viewStateHash` with the clock, the sid excluded, and the tile key fields (item 5);
    - §19: flat tiles may abstain (item 6, part);
    - item 8: the drifted anchors.
  - **S3 spec §4.2:** item 7, lazy print units in tiles mode. This amends the S3 spec, not the parent.
  - **S4 spec "As built (4a)":**
    - R3, R6, R9 and R10;
    - R12's wire-visible calls: `zoomStep`, the visible-index `sheetId`, the fixed palettes, and the per-session `timeZone` from the decision;
    - R8's conservative spill.
  - **Left to 4b and 4c:** items 1 (merge purge, undrawn), 2, 4 (`'sheet'`), 6 (sort order, fit) and 11.
- [ ] **Step 5: Write `docs/validation/streaming-s4a.md`:**
  - each S4 §9 4a "done when" line, with the test or script that proves it;
  - the Task 1 and Task 17 latency tables;
  - the compatibility matrix output;
  - the gate bytes;
  - what is not machine-checked: real host themes, and non-Chromium browsers.
- [ ] **Step 6: Commit.**
  ```bash
  git commit -m "docs(stream-s4): 4a validation, tile p95, spec amendments" -- scripts/stream-s4/tile-p95.ts server/test/sheet-corpus.test.ts server/src/ws.ts docs/validation/streaming-s4a.md docs/superpowers/specs/2026-11-07-server-streaming-design.md docs/superpowers/specs/2026-11-11-streaming-s3-design.md docs/superpowers/specs/2026-11-22-streaming-s4-design.md
  ```

## Spec coverage

| S4 spec | Task |
|---|---|
| §2 decisions; §3 anti-goals (no data on the wire, no local paint or watermark, units kept) | 15; 13, 14; 2, 4, 16 |
| §6.1 `StreamWorkbook`, widening, `stream-no-cell-data`, shell doc, 4a disposition rows | 12 (R16), 3 (`open`), 6 (`activateSheet`), 9 (`viewport` → tiles), 4 (gating) |
| §6.2 snapshot steps 1, 4, 5, 7 and the policy parameter | 5, 15 (R8); others: Not in 4a |
| §6.3 view state (identity), `viewStateHash`, clock | 9, 4 (R9) |
| §6.4 `sheet-layout` | 6 (R14) |
| §6.5 tiles, `XLSX_TILE`, buckets, render path, clock, cache key, marks, audit, renderer swap | 7, 9, 8, 10, 13 (R3, R4, R12, R13, R15) |
| §6.8 `'tile'` class, message caps, T8 p95 | 9, 17 (R3, R4) |
| §6.10 negotiation, opt-in, compatibility, no units at open, lazy print | 4, 11, 16, 3, 14 |
| §6.11 a11y positions and §15 notice, toolbar, bundle | 14, 17 |
| §6.12 items 3 and 8 (zoom steps, base-covered paint) | 13, 15 |
| §9 4a "done when" | 14 and 17 (no notice, million rows), 17 (p95), 13 and 14 (`basePaints`), 15 (wire audit, T2), 16 (matrix), 9 (view-state sharing), 4 (no unit until print), 7, 10 and 15 (seam, T9), 17 (T11) |
| §10 amendments 4a needs | 17 |
