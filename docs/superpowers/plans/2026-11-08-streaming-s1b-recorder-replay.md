# Streaming Stage 1b — OVSD frames, recorder, replayer, `CanvasBackend` — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Encode and decode the §7.1 `OVSD` frame grammar byte-exactly, record any unit's paint as a typed op list, replay it pixel-identically on the same backend, and give the paint pipeline one injectable `CanvasBackend` seam.

**Architecture:** Wire code lives in the lazy `src/stream/` runtime (React-free, engine-free, reachable from `officeview/core` only via `loadStreamRuntime()`): `crc32.ts`, `frame.ts` (encode/decode/reassemble), `ops.ts` (opcode table + LE writer/reader), `unit.ts` (unit body), `replay.ts`. The recorder (`src/stream/record/`) wraps a real backing context like `scripts/stream-s0/census-proxy.ts` does, and is used only by the server (1c) and tests — it may import engine types but must not be statically reachable from `core.js` (1a gate 1 checks this). `CanvasBackend` lives in `src/core/canvas-backend.ts` and threads through the nine construction sites T1 confirmed.

**Tech Stack:** TypeScript, Web Streams `DecompressionStream('gzip')` for decode (browser + Bun/Node), injected compressor for encode (`node:zlib` in tests/server), Vitest, `@napi-rs/canvas` for the zero-diff corpus run.

**Spec:** `docs/superpowers/specs/2026-11-07-server-streaming-design.md` — §7.1 (grammar, codes, rejection rules, caps), §7.3 (recorder contract, `CanvasBackend`, `registerFont`), §7.5 (op list), §7.6 (tile fallback), §9.2 (construction sites), §16 T3/T6 (census, byte vectors, **golden framing vectors**).

**Depends on:** 1a (`src/stream/index.ts`, core gates). **Feeds:** 1c (server encodes with these), 1d (client decodes/replays).

`src/` paths touched outside `src/stream/`: `src/core/canvas-backend.ts` (new), `src/render/paint.ts:199`, `src/core/search.ts:212`, `src/core/images.ts:65/:149`, `src/slideshow/render/svg-raster.ts:64`, `src/core/text-metrics.ts:31/:42`, `src/drawing/paint.ts:126`, `src/drawing/pattern.ts:155`. CF (`.worktrees/cf`) edits `src/xlsx/**` — check `src/drawing/paint.ts` for overlap before merge.

## Global Constraints

- All multi-byte integers and all recorded coordinates are **little-endian**; coordinates and advances are **IEEE-754 f64** (§7.3).
- Header layout, kind codes (1–8), flag bits (bit0 final, bit1 ext01, bit2 compressed, others 0), ext01 at offset 23 (`totalLen` u64, `fragCount` u16), trailer `wholeCrc` = CRC-32 of the complete **decompressed** payload — exactly as §7.1. Magic bytes `4F 56 53 44`.
- CRC-32 = IEEE 802.3 / zlib (`0xEDB88320`, init and final XOR `0xFFFFFFFF`).
- One compressor: the whole payload gzip-compressed once, before fragmentation.
- Fragment ≤ **1 MiB** including header + ext01; fragments ≤ **256** per message; reassembly window **5 s**.
- Rejection order is §7.1's list; nothing partially validated is returned.
- Recorder never reads browser globals (`devicePixelRatio`); scale comes in as a parameter. `RECORD_TEXT` is never exposed by the recorder facade.
- Replay is 1:1 at the recorded scale; the zero-diff claim is **same backend, same faces**.

## Review Focus

1. **Gzip "space bomb"** — a 26-byte fragment inflating to hundreds of MiB must abort mid-inflate at the decoded cap, not after. Test in Task 2.
2. **Interleaved first fragments for different `reqId`s** — reassembly is keyed by `reqId`; one message's fragments never complete another's. Test in Task 2.
3. **`drawImage` of a canvas that is later redrawn** (scratch surfaces in `text-paint.ts:712–765`) — the recorder must snapshot pixels at draw time, so replay shows the old content. Test in Task 4.
4. **A painter that measures text mid-paint and branches on the width** — the recorder forwards `measureText` to the backing context, so layout is identical to a direct paint. Test in Task 4.
5. **An op the recorder does not support** — the unit is reported as `tile` (raster fallback) with the op name, never silently dropped. Test in Task 4.

---

### Task 1: CRC-32

**Files:** Create `src/stream/crc32.ts`; Test `tests/stream/crc32.test.ts`.
**Interfaces:** Produces `crc32(bytes: Uint8Array, seed?: number): number` (unsigned; `seed` for incremental use, default 0 meaning a fresh CRC).

- [ ] **Step 1: Failing tests** — `crc32(ascii('123456789')) === 0xCBF43926`; `crc32(empty) === 0`; equals `zlib.crc32` on 1,000 random buffers; incremental `crc32(b, crc32(a)) === crc32(a‖b)`.
- [ ] **Step 2:** run — FAIL. **Step 3:** table-driven implementation. **Step 4:** PASS. **Step 5:** commit `feat(stream): CRC-32 (§7.1)`.

### Task 2: OVSD frame codec — encode, fragment, reassemble, decode

**Files:** Create `src/stream/frame.ts`; Test `tests/stream/frame.test.ts`, `tests/stream/frame-golden.test.ts`.

**Interfaces:**
- `export const enum Kind { MANIFEST = 1, ASSET_DEFINE = 2, UNIT_RASTER = 3, UNIT_COMMANDS = 4, XLSX_TILE = 5, SHOW_PLAN = 6, NOTES_RASTER = 7, PRINT_RASTER = 8 }`
- `export const FRAGMENT_CAP = 1 << 20, MAX_FRAGMENTS = 256, REASSEMBLY_MS = 5000, HEADER_BYTES = 23, EXT01_BYTES = 10`
- `encodeMessage(opts: { ver: number; kind: Kind; reqId: number; payload: Uint8Array; compress: (b: Uint8Array) => Uint8Array; fragmentBytes?: number; ext01?: boolean }): Uint8Array[]` — compress once, append `wholeCrc` (CRC of the uncompressed `payload`), split so the 4-byte trailer stays in the final fragment, each fragment ≤ `fragmentBytes` (default `FRAGMENT_CAP`); ext01 always on multi-fragment, optional (default false) on single.
- `class Reassembler { constructor(opts: { acceptedVersions: ReadonlySet<number>; expects: (reqId: number, kind: Kind) => boolean; decodedCap: (kind: Kind) => number; encodedCap: (kind: Kind) => number; now?: () => number }); push(fragment: Uint8Array): Promise<FrameResult | null>; sweep(): Rejection[] }`
- `type FrameResult = { ok: true; kind: Kind; reqId: number; payload: Uint8Array } | Rejection`; `type Rejection = { ok: false; reqId?: number; reason: RejectReason }`
- `type RejectReason = 'magic' | 'version' | 'kind' | 'unexpected-reqId' | 'flags' | 'invariant-mismatch' | 'sequence' | 'final-count' | 'missing-ext01' | 'fragment-cap' | 'length' | 'fragment-count' | 'fragment-crc' | 'total-len' | 'encoded-cap' | 'decoded-cap' | 'inflate' | 'whole-crc' | 'timeout'`
- `push` returns `null` while a message is incomplete; a rejection discards that `reqId`'s buffer. `sweep()` expires buffers older than 5 s with `'timeout'`. Inflation uses `DecompressionStream('gzip')`, counting output and cancelling the stream past `decodedCap(kind)`.

- [ ] **Step 1: Golden vectors (failing)** — the three hex vectors in spec §16 T6 verbatim: each decodes (accepted version {1}, `expects` true for reqId 7 / MANIFEST) to payload `{}` (bytes `7b 7d`); `encodeMessage({ver:1, kind:1, reqId:7, payload:'{}', compress: gzipVectors, ext01:false})` reproduces vector 1 byte-for-byte — where `gzipVectors = b => { const g = zlib.gzipSync(b, { level: 9 }); g[9] = 0xff; return g }`: the vectors were made with Python `gzip.compress(mtime=0)`, whose header OS byte is `0xff`, while Node writes a platform byte (`0x13` macOS, `0x03` Linux) — the OS byte is part of "the same gzip parameters", `ext01:true` reproduces vector 2, and `fragmentBytes: 23 + 10 + 10` reproduces vector 3's split.
- [ ] **Step 2: Battery (failing)** — one test per `RejectReason` built by mutating a valid encoding: wrong magic; ver 2 when {1} accepted; kind 0 and 9; reqId not expected (assert no buffer allocated — `Reassembler.bufferedBytes === 0`); reserved flag bit set; kind changed on fragment 1; frameSeq gap; final on fragment 0 of a 2-fragment message; multi-fragment without ext01; fragment of 1 MiB + 1; fragCount 257; flipped payload byte (fragment CRC); totalLen ≠ Σ payloadLen; compressed blob over `encodedCap`; **space bomb**: gzip of 512 MiB of `0x20` (built streaming, not in memory) with `decodedCap` 256 KiB rejects `'decoded-cap'` with peak RSS growth < 32 MiB; corrupted gzip → `'inflate'`; wrong trailer → `'whole-crc'`; timeout via injected `now`. Plus: every kind 1–8 round-trips; two `reqId`s with interleaved fragments both complete correctly; a 3 MiB payload of random bytes round-trips in ≥ 4 fragments, all ≤ 1 MiB.
- [ ] **Step 3:** run — FAIL. **Step 4:** implement. **Step 5:** PASS; commit `feat(stream): OVSD frame codec + golden vectors (§7.1, T6)`.

### Task 3: Op list + unit body

**Files:** Create `src/stream/ops.ts`, `src/stream/unit.ts`; Test `tests/stream/ops.test.ts`.

**Interfaces:**
- Opcode table (frozen here; add it to spec §7.5 in the same commit as an appendix row "assigned opcodes"):

| code | op | operands (LE) |
|---|---|---|
| 1 | STATE_PUSH | — |
| 2 | STATE_POP | — |
| 3 | TRANSFORM | u8 mode (0 set, 1 multiply) + 6 × f64 a b c d e f |
| 4 | STYLE_COLOR | u8 target (0 fill, 1 stroke, 2 shadow) + string |
| 5 | STYLE_OBJ | u8 target (0 fill, 1 stroke) + u32 assetId |
| 6 | STATE_SETUP | u32 field mask + present fields in mask-bit order (see `StateFields`) |
| 7 | PATH | u8 sub (0 begin, 1 moveTo, 2 lineTo, 3 quad, 4 bezier, 5 arc, 6 arcTo, 7 ellipse, 8 rect, 9 roundRect, 10 close) + f64 args (+ u8 ccw for arc/ellipse; roundRect radii as u8 count + f64s) |
| 8 | PATH_END | u8 mode (0 fill, 1 stroke, 2 fill_stroke, 3 clip) + u8 rule (0 nonzero, 1 evenodd) |
| 9 | RECT | u8 mode (0 fillRect, 1 strokeRect, 2 clearRect) + 4 × f64 |
| 10 | GLYPH_RUN | u8 mode (0 fill, 1 stroke) + string text + f64 x + f64 y + u8 hasMax + f64 maxWidth? + u32 n + n × f64 advances |
| 11 | IMAGE_DRAW | u32 assetId + u8 argc (3, 5, 9) + (argc − 1) × f64 |
| 12 | TILE | u32 w + u32 h + f64 dx + f64 dy + u32 byteLen + RGBA bytes |
| 13 | RASTER_BLOCK | reserved for §7.7; decoder rejects in 1b |

  Strings are `u32 byteLen + UTF-8`. `StateFields` (mask bit → field): 0 lineWidth f64, 1 lineCap u8, 2 lineJoin u8, 3 miterLimit f64, 4 lineDash (u16 n + n × f64), 5 lineDashOffset f64, 6 globalAlpha f64, 7 globalCompositeOperation string, 8 imageSmoothingEnabled u8, 9 imageSmoothingQuality u8, 10 letterSpacing string, 11 wordSpacing string, 12 direction u8, 13 lang string, 14 fontKerning u8, 15 fontVariantCaps string, 16 shadowBlur f64, 17 shadowOffsetX f64, 18 shadowOffsetY f64, 19 font string, 20 textAlign u8, 21 textBaseline u8, 22 fontStretch string, 23 textRendering u8. Enum fields map to the canonical CSS keyword lists in `ops.ts`.
- Assets: `type Asset = { id: number; hash: string } & ({ type: 'bitmap'; w: number; h: number; rgba: Uint8Array } | { type: 'linear'; x0,y0,x1,y1: number; stops: Array<[number, string]> } | { type: 'radial'; x0,y0,r0,x1,y1,r1: number; stops: Array<[number, string]> } | { type: 'pattern'; bitmap: number; repetition: string; transform?: [number,number,number,number,number,number] })`; `hash` = sha-256 hex of the canonical encoding (dedupe key).
- `class OpWriter { op(code, …) ; bytes(): Uint8Array; opCount: number }`, `class OpReader implements Iterable<Op>` — the reader throws `UnknownOpcode` / `Truncated` on bad input (callers map to unit refusal, §7.6).
- `encodeUnit({ unitId, discriminator: 1 | 2 | 3, body }): Uint8Array` / `decodeUnit(bytes): { unitId, byteLen, opCount, discriminator, body }` — `unitId u32 + byteLen u32 + opCount u32 + discriminator u8 + body` (§7.1).

- [ ] **Step 1: Failing tests** — every opcode round-trips with edge values (−0, NaN, ±Infinity, subnormal f64; empty and astral-plane strings); a truncated stream at every byte offset throws `Truncated`, never returns a partial op; opcode 0, 13, 255 → `UnknownOpcode`; unit body round-trips; `byteLen` mismatch rejected.
- [ ] **Steps 2–5:** FAIL → implement → PASS → commit `feat(stream): op list and unit body (§7.5)`.

### Task 4: Recorder facade

**Files:** Create `src/stream/record/recorder.ts`; Test `tests/stream/recorder.test.ts` (napi config: `vitest.stream-s0-napi.config.ts` alias; add `tests/stream/**` to its `include`).

**Interfaces:**
- `createRecorder(backing: CanvasRenderingContext2D, opts: { scale: number; assets?: AssetTable }): { ctx: CanvasRenderingContext2D; finish(): RecordedUnit }`
- `type RecordedUnit = { discriminator: 1; ops: Uint8Array; opCount: number; assets: Asset[] } | { discriminator: 3; reason: string /* first unsupported op, e.g. 'putImageData' */ }`
- `class AssetTable { add(asset): number /* dedupes by hash */; all(): Asset[] }`
- The facade forwards every call/property to `backing` **and** records it (so `measureText`, `getTransform`, `'letterSpacing' in ctx` behave as on the backing canvas). `drawImage` snapshots its source's pixels at call time (`getImageData` on canvas sources; decoded bitmaps via a backend-provided reader) into a bitmap asset. Gradients/patterns return wrapper objects that record `addColorStop`/`setTransform` and resolve to STYLE_OBJ assets when assigned. Any method or setter outside §7.3/§7.5's subset (reuse `RECORDED_SUBSET` from the census, moved to `src/stream/record/subset.ts`) latches `discriminator: 3` with that name. `RECORD_TEXT in ctx` is `false`.

- [ ] **Step 1: Failing tests** — (a) path + fill + transform + alpha sequence recorded in order; (b) scratch canvas drawn, redrawn red, drawn again → two distinct bitmap assets; (c) painter branches on `measureText('Hello').width > 30` → recorded ops identical to what the painter emits on a bare backing context (compare by replay pixels, Task 5); (d) `putImageData` → `{ discriminator: 3, reason: 'putImageData' }`; (e) `RECORD_TEXT in rec.ctx === false`; (f) no access to `globalThis.devicePixelRatio` (define a throwing getter in the test).
- [ ] **Steps 2–5:** FAIL → implement → PASS → commit `feat(stream): recorder facade (§7.3)`.

### Task 5: Replayer

**Files:** Create `src/stream/replay.ts`; Test `tests/stream/replay.test.ts` (napi config).

**Interfaces:** `replayUnit(ctx: CanvasRenderingContext2D, unit: { ops: Uint8Array; assets: ReadonlyMap<number, Asset> }, backend: Pick<CanvasBackend, 'createCanvas'>): void` — applies ops 1:1 (no scaling — scale was baked at record time); bitmap assets materialise once per replay via `backend.createCanvas` + `putImageData`; throws `ReplayError` on an unknown asset id or opcode before drawing anything (validate the whole op stream first, then draw).

- [ ] **Step 1: Failing tests** — for each recorder test scenario: direct paint on canvas A vs record on canvas B (backing) + replay on fresh canvas C → `A` and `C` RGBA identical; invalid asset id → `ReplayError` and canvas C untouched (all zero).
- [ ] **Steps 2–5:** FAIL → implement → PASS → commit `feat(stream): replayer (§7.3)`.

### Task 6: `CanvasBackend` + `registerFont` seams

**Files:** Create `src/core/canvas-backend.ts`; Modify the nine sites listed above, `src/render/paint.ts` (`PaintOptions.backend?`), `src/core/fonts/register.ts` (export the `registerFont` interface type); Test `tests/core/canvas-backend.test.ts`.

**Interfaces:**
- `export interface CanvasBackend { createCanvas(w: number, h: number): CanvasLike; decodeImage(bytes: Uint8Array, mime: string): Promise<ImageLike>; loadImage(bytes: Uint8Array): Promise<ImageLike>; rasteriseSvg(bytes: Uint8Array, w: number, h: number): Promise<ImageLike> }`
- `export const domBackend: CanvasBackend` — today's behaviour at each site, moved verbatim (OffscreenCanvas / `document.createElement('canvas')` / `createImageBitmap` / `Image`).
- `getPaintables(doc, { backend = domBackend, registerFont = browserRegisterFont, … })` threads both to every site; the four `canvas.constructor` clone sites call `backend.createCanvas` instead. The text-metrics sites keep their current throw-on-failure behaviour (§7.3 correctness note) — no new catch.
- `export type RegisterFont = (request: { family: string; bytes: Uint8Array; descriptors?: FontFaceDescriptors }) => Promise<{ release(): void }>`

- [ ] **Step 1: Failing tests** — a counting backend passed through `getPaintables` sees every construction for a docx, a pptx with a pattern fill + SVG picture, and an xlsx with a chart (assert ≥ 1 call at each of the nine sites via a site tag argument in test builds — or count by method and fixture); no site constructs a canvas outside the backend when one is supplied (spy on `OffscreenCanvas` / `document.createElement`).
- [ ] **Step 2:** FAIL. **Step 3:** implement. **Step 4:** PASS, **full `bun run test` green** (default `domBackend` = no behaviour change), `bun run gates:core` still green (seam is not in core's closure). **Step 5:** commit `feat(render): CanvasBackend and registerFont seams (§7.3, §9.2)`.

### Task 7: Acceptance — corpus zero local-backend diff + census on the real recorder

**Files:** Create `scripts/stream-s1/zero-diff.ts`; Results `docs/validation/streaming-s1/zero-diff.json`; Modify `docs/validation/streaming-s1.md`.

`zero-diff.ts` (run `bun run s0:napi scripts/stream-s1/zero-diff.ts`): for every corpus unit (57 files) at scale 1 and 2 with an `@napi-rs/canvas` `CanvasBackend`: paint directly (A); record on a backing canvas, `encodeUnit` → `encodeMessage` (zlib level 9, 64 KiB fragments) → `Reassembler` → `decodeUnit` → `replayUnit` (C); compare RGBA. Units the recorder marks `discriminator: 3` are counted as tiles with their reason.

- [ ] Gate: **every command unit has 0 differing pixels** at both scales; result lists tile units by reason (target: tiles ≤ the Stage-0 T3 expectation of 0 unsupported ops; any tile is reported to the owner, not fixed by widening the subset).
- [ ] `bun run test` full suite, `npx tsc --noEmit`, `bun run gates:core` green.
- [ ] Commit `test(stream-s1): corpus zero-diff record→frame→replay (§17 row 1)`.
