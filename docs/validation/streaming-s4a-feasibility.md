# Stream S4a, Task 1: tile feasibility gate

**Verdict: RED, on one clause.** Leg (c)'s blank-pane clause failed: 0 of 30 abstained `flat`, because all 30 were attributed to the right session. Every other gate is GREEN. Per R1, Task 1 stops here and reports. No threshold was changed.

| Leg | Gate | Measured | Verdict |
|---|---|---|---|
| (a) headless paint | 512² PNG, differs from blank, byte-equal over two runs | 512×512, `kind: detail`, 1,040 records, differs, byte-equal; identical sha across 3 process runs | GREEN |
| (b) latency, zoom 1, dpr 1 | in-process p95 ≤ 100 ms (AMBER ≤ 120) | p95 **86.8 ms** (run 3), 87.3 (run 2); run 1, contended: 184.4 | GREEN (runs 2–3) |
| (b) latency, zoom 1, dpr 2 | same | p95 **57.4 ms** (run 3), 57.2 (run 2), 60.9 (run 1) | GREEN |
| (c) lossless | ≥ 90 % | 30/30 = 100 % | GREEN |
| (c) JPEG q70 | ≥ 75 % | 30/30 = 100 % | GREEN |
| (c) 50 % downscale | ≥ 75 % | 30/30 = 100 % | GREEN |
| (c) wrong attributions | 0 | 0 in 90 captures (plus 0 in 30 blank-pane captures) | GREEN |
| (c) blank pane | abstains `flat` 30/30 | **0/30 abstained; 30/30 attributed to the right session** (z p50 39.7) | **RED** |
| (d) million rows | open ≤ 10 s, heap growth ≤ 256 MiB, last-row tile renders | open 196 ms, heap +6.45 MiB, A1048576 painted | GREEN |

Machine: darwin-arm64 (14 cores), Bun 1.4.5. Other agents' test suites shared the CPU; the 1-minute load average was recorded with each run.

Run: `bun scripts/stream-s4/tile-feasibility.ts`. Raw numbers: `streaming-s4a-feasibility.json` (run 3, with its comparison against run 2).

## Pinned decisions (for Tasks 5–10)

### Parse path (R2, F1)

```
projectPackage(original, 'xlsx', 'visible-only')
  → detectAndParse(projected.pkg, pkg => parseXlsx(pkg, { valuePolicy: 'cached' }))
  → buildSheetGeometry(sheet) → prepareSheetAssets(doc, { sheetIndex: 0 })
```

- The parse is the resident one over the **projected** package, with saved values.
- `doc.valuePolicy === 'cached'` is checked in leg (a).
- The engine parse is recorded beside it in leg (d).
- `openImportSession` was not needed, since (d) is GREEN.

### Basis

The basis is **per-tile**, keyed `sheet:<id>@<zoom>x<dpr>:<c>,<r>`. The label is inlined here; Task 7 exports it as `tileBasisLabel`.

The crop alternative attributed 0 % against per-tile's 100 %, so per-tile wins by S4 §8's rule: crop would need to be at least 5 points higher on lossless.

## (a) Headless paint

The tile is `grid500x100` tile (0,0) at zoom 1, dpr 1, painted under `serverBackend()` with `ctx.setTransform(1,0,0,1,0,0)` and `renderSheetViewport`, as `sheetTiles.ts:340–360` paints.

| Measure | Value |
|---|---|
| PNG size | 512×512, 67,301 B |
| PNG sha256 | `0fcbd294…ccf990d` |
| `blankSheet` tile (0,0) sha256 | `1da64cc4…203e927` |
| Byte-equal in process | yes, with a fresh `LayoutCache` per run |
| Byte-equal across processes | yes, identical in all three runs |

## (b) In-process latency

**What was timed.** Every tile of the used range, measured serially in one process:

- **render:** `createCanvas`, `setTransform` and `renderSheetViewport`;
- **mark:** `paintMark` with the `fixed` text;
- **encode:** `toBuffer('image/png')` and `criticalChunksOnly`;
- **layer1:** `deriveBasisKey` and `markRaster` over the per-tile basis (decode, embed, re-encode);
- **frame:** the R13 body inside `encodeUnit` (discriminator 3), `gzipSync` at level 6, and `encodeMessage` with kind 5 and 1 MiB fragments.

**Tile counts.** The tile set is taken over the used range A1:CV500, which is 6,400 × 10,000 CSS px (F8: the geometry extent is the whole domain). That gives 13 × 20 = 260 tiles at dpr 1 and 25 × 40 = 1,000 at dpr 2, which is R3's 1,260 requests.

**Gate statistic.** The gate is on the p95 of the per-tile **total**.

### In-process latency table (run 3, ms)

| Stage | dpr 1 p50 | dpr 1 p95 | dpr 1 max | dpr 2 p50 | dpr 2 p95 | dpr 2 max |
|---|---|---|---|---|---|---|
| render | 36.45 | 53.33 | 76.19 | 17.79 | 23.81 | 30.08 |
| visible mark | 0.15 | 0.21 | 4.14 | 0.13 | 0.18 | 0.30 |
| PNG encode | 9.02 | 9.50 | 10.84 | 8.72 | 9.16 | 15.11 |
| markRaster (per-tile) | 22.98 | 25.91 | 30.92 | 22.81 | 25.03 | 31.44 |
| gzip framing | 1.45 | 1.57 | 4.19 | 1.31 | 1.43 | 4.65 |
| **total** | **70.85** | **86.79** | 117.88 | **50.94** | **57.39** | 64.47 |

**Tile output, run 3:**

| Scale | Kinds | Min Layer-1 carriers per tile | Unmarked tiles | Mean canonical PNG | Mean framed |
|---|---|---|---|---|---|
| dpr 1 | 260 `detail` | 445 / 512 | 0 | 122,592 B | 93,835 B |
| dpr 2 | 1,000 `detail` | 362 / 512 | 0 | 106,751 B | 85,339 B |

### Reproducibility

The p95 of the total, in ms:

| Run | 1-min load at start | dpr 1 p95 | dpr 2 p95 | Comparison |
|---|---|---|---|---|
| 1 | 7.96, rising during the run | 184.40 | 60.94 | — |
| 2 | 7.84 | 87.34 | 57.22 | dpr 1 −52.6 % vs run 1 (**outside ±10 %**); dpr 2 −6.1 % |
| 3 | 4.98 | 86.79 | 57.39 | vs run 2: dpr 1 −0.6 %, dpr 2 +0.3 % (both within ±10 %) |

**Run 1 was contended.** The load climbed back into the teens during its dpr 1 section, and every stage slowed together there: render max 310 ms, encode p50 14 ms against 9, layer-1 p50 37 ms against 23. Runs 2 and 3 agree within 1 %.

**(a) and (c)** were byte-identical (JSON-equal) in all three runs.

**Reading.** The dpr 1 total sits about 13 ms under the 100 ms GREEN line on an uncontended core. A loaded host pushes it past 120 ms, so T8's 120 ms in-socket budget is load-sensitive at dpr 1.

**Render cost of font pinning (diagnostic, not gated).** A side probe rendered the same 60 tiles (dpr 1) through the raw `napiBackend` and through `serverBackend()`:

| Backend | Render p50 | Fonts |
|---|---|---|
| raw `napiBackend` | 31.6 ms | unpinned (Calibri, a host face) |
| `serverBackend()` | 55.7 ms | pinned (Carlito, through the `pinFonts` Proxy) |

That is about 1.8×, measured under load. S0 T8 measured a 17 ms p50 on a raw backend.

## (c) Layer-1 on tiles

**Setup:**

- 64 session codes: `messageFor` with a fixed key and fixed tokens, then `extendedCodeword`.
- `tauFor(64)` gives abs 7.93 and margin 3.45; the smallest declared table is M = 2^16.
- 30 windows of 2×2 tiles, spread deterministically over `grid500x100`, each with its own session.

**Each capture:**

1. The four tiles are rendered with the visible mark and encoded; this is the canonical, pre-forensic tile.
2. Each is marked with its own basis (`markRaster`).
3. They are composited to 1,024², and a 900 × 900 crop is taken at (40, 40).
4. The crop is registered against the **R10 composite reference**: the unmarked canonical tiles composited, with each tile's full 512-chip list offset by its tile origin, concatenated in whole blocks.
5. It is scored with `scoreCapture`, then `decideScores`.

| Class | Attributed | Wrong | None | Abstain | z (true session) p50 / p95 / max |
|---|---|---|---|---|---|
| lossless | 30/30 (100 %) | 0 | 0 | 0 | 38.42 / 39.18 / 39.23 |
| JPEG q70 | 30/30 (100 %) | 0 | 0 | 0 | 27.21 / 28.32 / 28.75 |
| 50 % downscale | 30/30 (100 %) | 0 | 0 | 0 | 38.42 / 39.18 / 39.23 |
| **blank pane** (lossless) | **30/30 to the right sid** | 0 | 0 | **0 flat** | 39.72 / 40.20 / 40.47 |

### Why the blank pane does not abstain (the RED)

The gate assumes a blank pane gives flat evidence. It does not:

- **The white cells carry the mark.** `pedestalFor` gives an all-white cell a −2 pedestal, so it becomes a carrier, and `markRaster` marked every blank tile. The minimum was 445 carriers at dpr 1.
- **The gridlines and the visible mark let the capture register.** The capture therefore registers, the residual is far from σ_r = 0, and the right session is named with z ≈ 40.

`flat` is reached only for an **unmarked**, reference-identical capture (parent §13.2). Here it is never reached.

The plan's Task 10 test `"a blank pane abstains"` and R15's "flat tiles may abstain" rest on the same assumption. They will not hold as written. This is a ruling for the owner; it is not a detector defect. No wrong attribution was made.

### The `markLayerRaster` crop alternative

The alternative crops a sheet-wide basis over the used tile grid, 6,656 × 10,240 device px, with the same captures and the codes re-indexed to the kept chips.

| Measure | Value |
|---|---|
| Carriers per tile | p50 1, max 6 |
| Chips in a 2×2 window | p50 8, p95 11, max 13 |
| Attribution | 0 % in every class (lossless: 25 `none`, 5 `flat`) |

Over the whole 16,384 × 1,048,576 domain the basis cannot be keyed at all: `basis` throws `Insufficient carrier capacity` (capacity > 2³²). **Per-tile is pinned.**

## (d) Million rows

The fixture is `millionRows` (51,782 B): A1:D2000 and A1048576, with a 30 pt band on rows 10–20. The open was measured first in the process, with `Bun.gc(true)` before each sample and the opened document held through the second sample.

| Parse | Open wall | Project | Parse | Geometry | Assets | Heap growth | RSS growth |
|---|---|---|---|---|---|---|---|
| cached (pinned) | **196 ms** | 136 | 55 | 4.4 | 0.5 | **+6.45 MiB** | +71.5 MiB |
| engine (comparison) | 143 ms | 93 | 48 | 1.6 | 0.03 | +3.2 MiB | −10.0 MiB |

The engine row ran second, on a warm process.

**Last-row tile.** Row 1,048,576 sits at y = 20,971,720 CSS px, which is 1,048,575 × 20 plus 11 × 20 for the band. Its tile is row 40,960.

- It rendered in 16.95 ms, with `kind: detail`.
- The A1048576 record was painted.

## Fixtures (`server/test/sheet-fixtures.ts`)

**Construction.** `buildXlsx`, plus exported JSZip helpers (F6):

- `patchXlsx`, `finishXlsx`, `patchSheetXml` and `patchStylesXml`;
- `insertSheetElement`, which keeps CT_Worksheet order;
- `setRowAttrs`, `freezeXml`, `addDrawing`, `rectShapeXml`, `twoCellAnchorXml`, `absoluteAnchorXml` and `colName`.

**Deterministic bytes.** Every entry is rewritten at one fixed date.

**Sanity checks.** Each fixture is checked on the projected, cached-policy parse, and all checks pass (`fixtures.sanity` in the JSON):

| Fixture | Check |
|---|---|
| `grid500x100` | 500 × 100 cells, `R1C1`, CU500 = 500099 |
| `millionRows` | 2,001 rows, `LAST-ROW`; band rows 40 px, other rows 20 px |
| `blankSheet` | no rows |
| `cfToday` | the rule is present; A5 = 46349 (2026-11-23) |
| `unfilledShapeOverText` | 1 drawing, fill `none`, B2 = `COVERED-B2` |
| `rotatedBox` | `rotationDeg 90`, `transformValid` |
| `unresolvedAnchor` | `unresolved = 1` |

**`savedFormula`.** It paints the same PNG as `plainSeven` under the cached policy. It **also does so under the engine policy**: the engine leaves the saved 7 of `=1+1` in place. So this fixture does not tell the two policies apart (see the concerns below).

## Concerns

1. **The (c) blank-pane clause is RED**, as above. The owner decides among three options:
   - change the gate to "a marked blank pane never names a wrong session";
   - leave flat tiles unmarked, so that an unmarked flat capture abstains;
   - keep the gate, and require that blank tiles abstain.

   Task 10's blank-pane test inherits whichever is chosen.
2. **Leg (b) is load-sensitive at dpr 1.** The headroom is 13 ms on a quiet core, and the contended run read 184 ms. Two levers, neither tried here:
   - font pinning costs about 1.8× on render;
   - `markRaster` decodes and re-encodes the PNG, about 23 ms per tile, which is a quarter of the total.
3. **`savedFormula` does not discriminate.** The engine does not overwrite the saved 7. F1's T5 and T15 checks that `savedFormula == plainSeven` would pass under either policy. Task 5 needs a fixture the engine does recompute, for example `fullCalcOnLoad`, or a formula over a purged hidden row.
4. **The 50 % downscale is easy here.** The (40, 40) offset is even and the cells are 4 px, so a 2× box-average keeps every chip's mean. That is why its z equals the lossless z. An odd-offset downscale would be a harder test; it was not asked for.
5. **The visible mark is approximated.** It is laid out per tile, over the tile's CSS box at the tile's scale. The product lays it out in global sheet device space; the cost is of the same order, about 0.2 ms.
6. **`tauFor(64)`** resolves to the smallest declared table, M = 2^16, so τ is conservative for 64 candidates.
