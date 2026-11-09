# Streaming Stage 2a — Layer-1 forensic mark live on raster units

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Every final-raster unit a forensic-on session receives carries its own keyed spread-spectrum pixel mark (§13.2 Layer 1), recorded in the audit table and recoverable from the real server pipeline's output.

**Architecture:** Stage 0's measured prototype (`scripts/stream-s0/mark/`) moves into `server/src/mark/` unchanged in behaviour (golden-pinned). The render cache keeps holding the **canonical** lossless PNG (visible mark composited, no forensic mark) exactly as in Stage 1. A forensic-on session then runs a per-session **encode job** in the pool (decode canonical PNG → embed the session's 512-sign code over the page-keyed basis → re-encode) under a ≤ 2-in-flight scratch lease (§9.3/§9.4/§14). Forensic-off sessions keep sending the canonical bytes. The audit table issues the 24-bit token, draws the salt, and stores the keyed `msg_40`.

**Tech Stack:** Bun server workspace (`server/`), Bun Workers pool, `node:crypto` / `node:zlib`, vitest + bun:test, Stage-0 detector/attacks for validation.

**Spec:** `docs/superpowers/specs/2026-11-07-server-streaming-design.md` — §2 items 10/11, §6.1, §9.3, §9.4, §13.1–13.4, §14 (scratch lease), §16 T7/T9, §17 Stage 2. Stage-0 numbers: `docs/validation/streaming.md` (T7/T9 rows).

**Stage 2 split (this plan is 2a):**
- **2a (this plan):** Layer-1 embed in the serving path; keys; token/salt/msg issuance; audit fields; policy + notices; scratch lease; integration proof at Stage-0 τ.
- **2b:** `server/scripts/detect.ts` CLI + audit export; τ recalibration over the **composed** decision and the ≤ 10⁻⁶ false-attribution gate; collusion battery on production output; audit persistence, retention and key rotation (T10 rows).
- **2c:** occlusion **store** (command culling) and the M-rule prototype with its shipping gate.

**What 2a's green does NOT claim** (say so in the PR and validation doc): no recalibrated τ, no 10⁻⁶ measurement on the production pipeline, no collusion numbers on production output, no detection CLI, no Layer 2 (command units carry no forensic mark), no audit persistence across restarts.

## Global Constraints

- Codec parameters are the Stage-0 choice, verbatim: `{ nChips: 512, cellPx: 4, tilePx: 64, amplitude: 2, pedestal: 2 }`; code = extended BCH(511,40) → 512 signs; codec id `l1-bch512x40@1`.
- `codeword_40 = token_24 | salt_8 | CRC8_8`; `msg_40 = HMAC-SHA256(kMsg, keyGen | codeword_40)` truncated to 40 bits; an insert colliding with a live row's `msg_40` re-draws the salt (§13.2).
- Token: 24-bit monotonic counter, never reused while its audit row lives; **exhaustion refuses sessions** rather than recycling (§13.2).
- `kBasis` is page-keyed (docRef, unitId — never session-keyed); `kMsg` is a further HKDF child (§13.4).
- ΔE00 page max on served output ≤ **3** (§13.2 ceiling; Stage 0 measured 1.59).
- Pipeline order: paint → occlusion pass → visible mark → Layer-1 → encode (§9.4).
- Canonical cache entries never hold per-sid pixels (§9.3); scratch lease ≤ **2** final-raster units in flight per session (§14).
- Survival targets (image-only units, §13.5): default ≥ 90 %, q70 ≥ 75 %, 50 % down ≥ 75 %, q70→50 % ≥ 75 %, 25 %/axis crop ≥ 70 %. Stage-0 τ: `τ_abs 7.99`, `τ_margin 3.85` at M = 65,536.
- Never claim a mark that is not applied: audit and notices record what was **applied**.
- `server/src/mark/` stays outside the T11 core closure (it is server-only); `gates:core` must stay green.

## Review Focus

1. **Tiny or flat units** — a unit with capacity < 512 cells (`basis` throws) or with every chip masked. Expected: the unit still ships (visible mark only), the server sends `diag` `forensic-unmarked`, and the per-unit carrier count is recorded — never a crash or a silent claim. Pinned in Task 5.
2. **Semi-transparent canonical pixels** — the PNG decode/re-encode must be bit-exact for every pixel the embed does not touch, including alpha < 255. Pinned in Task 2.
3. **A session closes or resumes mid-encode** — the lease is released exactly once, a late encode result is dropped, and the charge follows the pool's drain-then-refund rule. Pinned in Task 5.
4. **Salt collision / RNG stuck / token exhaustion** — re-draw, bounded refusal, and `policy-refused` at HELLO for a forensic-on session; forensic-off sessions still open. Pinned in Task 3.
5. **Two sessions, same unit** — different streamed bytes, each detecting to its own audit row; the cache still holds one canonical entry. Pinned in Task 5.

---

### Task 1: Production mark codec (`server/src/mark/`), golden-pinned to Stage 0

**Files:**
- Create: `server/src/mark/{types,params,bch,message,basis,embed}.ts` (moved from `scripts/stream-s0/mark/`)
- Modify: `scripts/stream-s0/mark/{types,params,bch,message,basis,embed}.ts` → one-line re-exports from `server/src/mark/`
- Test: `server/test/mark-golden.test.ts`

**Interfaces:**
- Produces: `DEFAULT_PARAMS`, `checkParams`, `extendedCodeword(msg40): Int8Array(512)`, `codeword40`, `crc8`, `messageFor(kMsg, gen, token, salt): Uint8Array(40)`, `insertMessage(kMsg, gen, token, used, draw?)`, `deriveBasisKey(root, docRef, unitId): Uint8Array(32)`, `basis(kBasis, w, h, params): Chip[]`, `embed(bmp, chips, code, params): MarkedBitmap`, plus `CODEC_ID = 'l1-bch512x40@1'` in `params.ts`.

- [ ] **Step 0: Setup** — `cd .worktrees/stream-s2 && bun install && (cd server && bun install)`; Expected: installs clean.
- [ ] **Step 1: Write the golden test against the Stage-0 modules first** (imports `scripts/stream-s0/mark/*`). Fixed inputs: key = bytes 0..31, gen 1, token `0x123456`, salt `0x5a`, docRef `'a'.repeat(24)`, unitId `'3@1'`, a 300×200 synthetic RGBA (white with a 40×40 grey block and a 20-px alpha-128 strip). Assertions (values captured on first run and pasted in as literals):
  - `sha256(messageFor(...))`, `sha256(extendedCodeword(msg))`, `sha256(JSON.stringify(basis(...)))`, `sha256(embed(...).data)` equal the literals;
  - `embed(...).mark` equals `{ variant, pedestalCells, maskedCells, params }` literal.
- [ ] **Step 2: Run** `cd server && bun test test/mark-golden.test.ts` — Expected: PASS against Stage-0 code (capture literals).
- [ ] **Step 3: Move** the six modules to `server/src/mark/`, switch the test imports to `../src/mark/*`, replace the Stage-0 files with re-exports. Add `CODEC_ID`.
- [ ] **Step 4: Run** the golden test + `bun run typecheck:scripts` + `npx tsc --noEmit -p server` — Expected: PASS, clean (identical literals prove no behaviour change).
- [ ] **Step 5: Commit** `refactor(stream): move the Stage-0 Layer-1 codec into server/src/mark (golden-pinned)`.

### Task 2: Lossless PNG codec + `markRaster` (the encode job body)

**Files:**
- Create: `server/src/mark/png.ts`, `server/src/mark/mark-raster.ts`
- Test: `server/test/mark-raster.test.ts`

**Interfaces:**
- Consumes: Task 1.
- Produces:
  - `decodePng(png: Uint8Array): { w: number; h: number; rgba: Uint8ClampedArray }` — 8-bit, non-interlaced, colour types 2 and 6, all five filters; anything else throws `PngUnsupported`.
  - `encodePng(w, h, rgba): Uint8Array` — colour type 6, per-row filter by minimum sum of absolute differences, `zlib` level 6, chunks IHDR/IDAT/IEND only.
  - `markRaster(input: { png: Uint8Array; w: number; h: number; kBasis: Uint8Array; code: Int8Array; params?: MarkParams }): { png: Uint8Array; carriers: number; masked: number; pedestalCells: number } | { unmarked: 'capacity' | 'all-masked'; png: Uint8Array }` — on `unmarked` the returned `png` is the input bytes unchanged.

- [ ] **Step 1: Write failing tests:**
  - `decodePng(skiaPng)` equals `getImageData` of the same `@napi-rs/canvas` surface for an opaque page and for a surface with alpha 0/128/255 regions (straight alpha);
  - `decodePng(encodePng(w,h,rgba))` is byte-identical to `rgba` for random RGBA (seeded), 1×1, 1×N and N×1;
  - `markRaster` on a rendered corpus unit (`poi-55733.docx#0` via `DocumentHost` at bucket 1): every pixel outside `chipSupport` is identical to the canonical decode (including alpha < 255); `carriers + masked === 512`;
  - `markRaster` on a 100×50 unit returns `{ unmarked: 'capacity' }` with the input bytes; on an all-transparent 512×512 surface returns `{ unmarked: 'all-masked' }`;
  - `markRaster` output detects to the embedded code with Stage-0 `scoreCapture`/`decideScores` at τ `{abs: 7.99, margin: 3.85}` over 64 candidates (index of the true code attributed).
- [ ] **Step 2: Run** `cd server && bun --preload ../scripts/stream-s0/napi-preload.ts test test/mark-raster.test.ts` — Expected: FAIL (modules missing).
- [ ] **Step 3: Implement** `png.ts` and `mark-raster.ts`. `markRaster` derives chips with `basis(kBasis, w, h, params)` (catch the capacity throw → `unmarked: 'capacity'`); `masked === nChips` → `unmarked: 'all-masked'`.
- [ ] **Step 4: Run** the test — Expected: PASS. Also record re-encoded vs Skia PNG size on the test unit in the test log (informational; no gate).
- [ ] **Step 5: Commit** `feat(stream): lossless PNG codec and markRaster for the Layer-1 encode job`.

### Task 3: Keys, token/salt/msg issuance, audit row

**Files:**
- Modify: `server/src/keys.ts`, `server/src/audit.ts`
- Test: `server/test/audit-issue.test.ts`

**Interfaces:**
- Produces:
  - `Keys` gains `kBasisRoot: Uint8Array` (HKDF info `'kBasis'`), `kMsg: Uint8Array` (info `'kMsg'`), `keyGeneration: number` (= 1 in 2a). Per-unit basis key: `deriveBasisKey(keys.kBasisRoot, docRef, \`${unitId}@${scaleBucket}\`)`.
  - `AuditRow` gains `forensic: { codec: 'l1-bch512x40@1'; token24: number; salt8: number; msg40: string /* hex, 10 chars */ } | null` and `forensicLayers: Array<'l1'>`; `markSettings.forensic` keeps meaning *applied*; `untraceable = visible === 'off' && forensicLayers.length === 0`.
  - `AuditTable.issue(kMsg, keyGeneration, draw?): { token24; salt8; msg40: Uint8Array; code: Int8Array }` — monotonic token; bounded salt re-draws via `insertMessage`; throws `ForensicExhausted` at 2²⁴ tokens or 256 failed draws. `expire()` frees the row's `msg40` from the used set; the token counter never rewinds.
- Ruling to ledger: per-unit `kBasis` is `HKDF(kMaster,'kBasis')` → `deriveBasisKey(root, docRef, unitId@bucket)` (two-step, equivalent secrecy; changes key bytes, not statistics). Bucket is part of the unit id because the basis depends on device-pixel size.
- Ruling to ledger: the audit table and token counter stay in memory in 2a; a restart resets the counter (attribution stays safe because prior rows are gone too, but the 90-day promise needs 2b persistence).

- [ ] **Step 1: Write failing tests:** tokens issue 0,1,2…; two issues never share `msg40`; a stuck `draw` (always 0) after a forced collision throws `ForensicExhausted`; a counter preset to `2**24 - 1` issues once then throws; `expire()` frees `msg40` but the next token is still `last + 1`; `record()` with `forensicLayers: []` and visible `off` stores `untraceable: true`, with `['l1']` stores false; `deriveKeys` returns distinct 32-byte `kBasisRoot`/`kMsg`.
- [ ] **Step 2: Run** `cd server && bun test test/audit-issue.test.ts` — Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run** — Expected: PASS.
- [ ] **Step 5: Commit** `feat(stream): issue forensic tokens, salts and keyed messages from the audit table`.

### Task 4: Policy, manifest notices, cache key

**Files:**
- Modify: `server/src/hooks.ts`, `server/src/cache.ts`, `server/src/ws.ts` (HELLO/manifest only), `server/README.md`
- Test: `server/test/handshake.test.ts`, `server/test/client-session.test.ts`, `server/test/cache.test.ts`, `tests/stream/manifest.test.ts`, `tests/stream/remote-session.test.ts` (fixture strings)

**Interfaces:**
- Produces:
  - `EffectivePolicy.forensicMark: boolean` (= requested, default true); `FORENSIC_NOTICE` removed. Notices: forensic on + `final-raster` → none; forensic on + `commands` → `forensic-commands-unmarked` (Layer 2 is later; raster fallbacks are still marked); forensic off → `forensic-off`. `relaxed-session` unchanged.
  - Session data gains `forensic?: { token24; salt8; msg40; code: Int8Array }` from `audit.issue` at HELLO when `forensicMark`; `ForensicExhausted` → close `policy-refused` (4009) before MANIFEST. Audit row: `forensicLayers: policy.forensicMark ? ['l1'] : []`.
  - A **resume** keeps the sid's existing `forensic` (same token, same code); it never issues a new row (the audit row is write-once per sid). Test: resume → no new token issued, units still detect to the original row.
  - `RenderKeyFields.markSettings` stays `{ visible; text; forensic: false }` — it keys the **canonical** tier only (comment updated).
- Ruling to ledger: exhaustion reuses `policy-refused` (no new close code; the client already maps 4009).

- [ ] **Step 1: Update/write failing tests:** default HELLO → manifest has no `forensic-*` notice, audit row `forensicLayers ['l1']`, `forensic.token24 === 0`; `?relaxed=1` → `forensic-commands-unmarked`; `forensicMark: false` + `visibleMark: 'off'` → `forensic-off`, `untraceable: true`; an audit table preset to exhaustion → close 4009 for default policy, MANIFEST for `forensicMark: false`. Replace `forensic-mark-stage-2` in the client fixtures with `forensic-off` (fixture strings only).
- [ ] **Step 2: Run** `cd server && bun --preload ../scripts/stream-s0/napi-preload.ts test test/handshake.test.ts test/client-session.test.ts test/cache.test.ts` — Expected: FAIL on the new assertions.
- [ ] **Step 3: Implement.** Update `server/README.md` (mark section) to describe what is applied in 2a.
- [ ] **Step 4: Run** the server suite + `npx vitest run tests/stream` — Expected: PASS.
- [ ] **Step 5: Commit** `feat(stream): forensic mark on by default; notices and audit record what is applied`.

### Task 5: Encode job, scratch lease, serving path

**Files:**
- Modify: `server/src/worker/render-worker.ts`, `server/src/pool.ts`, `server/src/ws.ts`
- Test: `server/test/forensic-serve.test.ts`

**Interfaces:**
- Consumes: Tasks 2–4.
- Produces:
  - `WorkerRequest` gains `{ type: 'encode'; id; png: Uint8Array; w; h; kBasis: Uint8Array; code: Int8Array; hangMs?; crash? }` → `markRaster` result (png transferred).
  - `Pool.encode(sessionId: string, job: EncodeJob, bytesEstimate: number): Promise<MarkRasterResult>` — same admission, deadline, crash and retire-lease paths as `render`; `bytesEstimate = w * h * 4 * 3` (decode + marked copy + encode buffers).
  - In `ws.ts`, after `obtain()` returns a raster for a session with `forensic`: acquire the session's scratch lease (FIFO semaphore of 2, per sid; released in `finally`, and on session teardown all waiters reject), run `pool.encode`, then frame. Command units and forensic-off sessions skip it. `{ unmarked }` → send the canonical bytes plus `diag` `{ code: 'forensic-unmarked', unitIndex, reason }`.
  - `stream.stats()` gains `forensic: { marked: number; unmarked: number; carriers: number[] /* last 256 */ }` for the validation script.

- [ ] **Step 1: Write failing tests** (harness `start()`, default policy):
  - two sessions request unit 0 of `pypptx-test.pptx`: their UNIT_RASTER PNGs differ; each decoded PNG detects (Stage-0 detector, reference = the canonical PNG from a forensic-off session, candidates = both sessions' `code`s + 62 decoys) to its own session; the render cache holds one entry for that unit (`stats().renderCache.entries` delta 1);
  - a forensic-off session's PNG is byte-identical to the canonical entry;
  - a `?relaxed=1` session receives UNIT_COMMANDS unchanged (no encode job: pool encode count 0);
  - a 100×50 unit (fixture deck with a tiny slide size) → canonical bytes + `diag forensic-unmarked reason capacity`;
  - with `hangMs` on encode, a third concurrent unit request of the same session is not posted to the pool until one of the first two finishes (lease = 2);
  - closing the socket mid-encode: no frame is sent, the lease count returns to 0, the pool's retire path refunds only after the worker replies.
- [ ] **Step 2: Run** `cd server && bun --preload ../scripts/stream-s0/napi-preload.ts test test/forensic-serve.test.ts` — Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run** the whole server suite, `npx vitest run`, `npm run test:browser:stream` — Expected: all green (browser specs stream marked rasters; `inked() > 100` unaffected).
- [ ] **Step 5: Commit** `feat(stream): per-session Layer-1 encode under a scratch lease`.

### Task 6: Integration proof on served output + docs

**Files:**
- Create: `scripts/stream-s2/l1-integration.ts`, `docs/validation/streaming-s2.md`
- Modify: `docs/validation/streaming.md` (Stage 2 row pointer), `server/README.md` (one paragraph)

**Interfaces:**
- Consumes: `DocumentHost` (canonical PNG per unit at bucket 1, default visible mark), `markRaster`, `AuditTable.issue`, Stage-0 `attacks`, `ciede2000.pageDeltaE`, `detect.scoreCapture/decideScores`.

- [ ] **Step 1: Write the script:** over every corpus unit at bucket 1 (plus every pptx unit at bucket 2), issue M = 65,536 candidate rows from a fresh `AuditTable`, mark each unit with row 0's code via `markRaster` on the canonical PNG, and report:
  - per-unit carriers / masked / pedestal cells / unmarked reason (the T7 per-unit carrier publication);
  - ΔE00 page max per unit (decoded canonical vs decoded marked), corpus max;
  - survival rows default / q70 / down0.5 / q70→down0.5 / crop25 (25 % per axis) on the marked **PNG bytes**, detected at Stage-0 τ against M candidates; rate = attributed-to-row-0 / marked units;
  - unmarked-capture abstain/none check: the canonical PNG scored against the same M never attributes.
  Exit 1 if ΔE00 max > 3, any survival row is below its target, or the unmarked check attributes.
- [ ] **Step 2: Run** `bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s2/l1-integration.ts --write docs/validation/streaming-s2.json` — Expected: exit 0. A red row is reported to the owner, not tuned around.
- [ ] **Step 3: Write** `docs/validation/streaming-s2.md`: the tables above, the Stage-0 comparison, the "does not claim" list from this plan's header, and the 2b/2c handoffs.
- [ ] **Step 4: Run** full gates: root suite, server suite, napi suite, `test:browser:stream`, `gates:core`, type checks — Expected: green.
- [ ] **Step 5: Commit** `docs(stream): Stage 2a Layer-1 integration proof on served output`.
