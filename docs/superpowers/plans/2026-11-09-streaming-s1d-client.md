# Streaming Stage 1d — client integration — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `<OfficeFile data={{ kind: 'stream', url }}>` opens a stream session against the 1c server, shows its units (rasters or replayed commands), survives a network drop without losing the reader's place, and leaves every local (byte-source) load exactly as it was; plus the Stage-1-mandatory §7.7 occlusion evidence on the server's paint pass.

**Architecture:** All client stream code lives in the lazy `src/stream/remote/` chunk. `useOfficeFile` branches on `isStreamSource` before any byte path and dynamically imports `openRemoteDocument`; the resulting `RemoteOfficeDoc` (§8.3) flows through `OfficeFileState` into `<OfficeDoc>`, whose dispatch sends `kind: 'remote'` to a lazy `RemoteDocView` (like `InteractiveSheetDoc`), never to `LegacyOfficeDoc`. The session protocol is driven by a `StreamSession` class that is engine- and React-free, so Bun tests run it against the real 1c server (`server/test/harness.ts`); Playwright covers the browser surface through `server/test/app.ts`. The server gains a detached-session grace window (resume after a real drop) and an occlusion observer on the existing `PaintObserver` seam.

**Tech Stack:** TypeScript, React 18 (lazy/Suspense), WebSocket, `DecompressionStream`, Web Crypto, `FontFace`, Bun test, Vitest (jsdom), Playwright (Chromium).

**Spec:** `docs/superpowers/specs/2026-11-07-server-streaming-design.md` (§6.1–6.3, §7.1 receipt guard, §7.4, §7.7, §8.1–8.3, §10, §11 disabled-service copy, §14 client caps, §16 T2/T4/T6/T10, §17 row 1). Previous: `docs/superpowers/plans/2026-11-09-streaming-s1c-server.md` ("Deferred to 1d").

**Branch:** continue on `feature/streaming-s1` (Stage 1 is one PR); one commit per task.

## Global Constraints

- `src/stream/remote/**` and `src/components/remote/**` must **never** enter `dist/index.js`'s or `dist/core.js`'s static closure; they are reached only through `import()`. `bun run gates:core` stays green; the routing edit is a **recorded change** in `scripts/eager-baseline.json` with its measured `gzipDelta` (target ≤ 256 B; over 512 B ⇒ stop and report to the owner).
- A `StreamSource` must never reach `fetch` or a byte reader (existing `StreamSourceMisuseError` guard stays).
- Client caps (§7.1/§14): every received WS message `byteLength` (or `Blob.size`) ≤ 1 MiB (`FRAGMENT_CAP`) checked **before** `Reassembler.push`; over ⇒ `ws.close(4007, 'oversized-frame')`. Per-kind encoded/decoded caps exactly the §14 table. Aggregate transient reservation ≤ 96 MiB per session; outstanding unit requests ≤ 8.
- `Reassembler.expects(reqId, kind)` is true only for a reqId the client issued and still awaits, and only for the kinds that request can yield (MANIFEST for the HELLO reqId; ASSET_DEFINE/UNIT_RASTER/UNIT_COMMANDS for a unit reqId).
- HELLO `{ reqId, type: 'hello', client: 'officeview/1.0', recVersions: [1], token? , resume? }`; READY `{ reqId, type: 'ready', tokenJti, missingFaces? }` sent ≤ 2 s after MANIFEST receipt **always** — face hashing never delays it (unfinished faces are reported missing).
- Close codes (client view): 4000–4006, 4008, 4009 and 4409 are terminal `StreamSessionError(code)`; 4007 is client-originated. Only an abnormal close (1001/1005/1006, or 1000 not initiated by the client) triggers resume.
- Stage 1 `can.search/copy/textA11y/present/notes/links` are false: the remote view has no text layer, no local text index, no search bar, no copy; it states **"text unavailable in this viewing mode"** once. `present()` rejects `SlideShowError('not-presentable')`. `allowPrint` = host prop ∧ `can.print`. The host `watermark` prop is ignored for a remote document (the baked mark is authoritative).
- Replay is 1:1 at a scale bucket from `manifest.recorder.scalesBuckets`; bucket = smallest bucket ≥ `zoom × devicePixelRatio`, else the largest. Zoom change re-requests; no client-side resampling of commands.
- Local byte-for-byte baseline: `ca952c50` (merge-base of `origin/main` and this branch).

## Decisions taken in this plan (owner may override)

- **D1 — font pack delivery (owner question, see handoff).** Protected sessions send `requiredFaces: []` (all rasters), so the pack matters only for relaxed `commands` sessions. Recommended default: **not in the npm package**; a `streamFonts?: { baseUrl: string }` prop names where the host serves `fontpack/` (`pack.json` + TTFs). Without it every required face is reported missing ⇒ those units arrive as rasters (§7.4 rule 4, correct by construction). The demo serves `fontpack/` itself.
- **D2 — `OfficeDocSource` arm.** The arm is `RemoteOfficeDoc` itself (it already carries `kind: 'remote'`), not a `{ kind: 'remote', session }` wrapper; same discriminant, one less object.
- **D3 — no app-level heartbeat.** The 1c server defines no heartbeat message and the §6.2 registry has none; liveness is the WS close event. `streamClient`'s default is `new WebSocket(url)`.
- **D4 — detached grace = 30 s [call]**, capped by the session token's own `exp`.

## Review Focus

1. **A server message over 1 MiB on receipt** — closed 4007 before any reassembly buffer exists; the view shows the typed error, not a hang. → Task 2 test.
2. **Source replaced or component unmounted mid-handshake (before MANIFEST)** — the socket closes exactly once, no state update after unmount, no second session for the old source. → Task 2 + Task 4 tests.
3. **RENEW arriving while units are in flight** — in-flight units complete; a `session-rotated` refusal re-issues the request once under the new token. → Task 2 test.
4. **A unit frame for a reqId after `close()` or after that request was abandoned** — dropped by `expects` before allocation; no paint into a released canvas. → Task 2 test.
5. **Network drop while scrolled deep into a document** — resume keeps `scrollTop` and repaints the visible units only; a drop after grace surfaces `auth-failed` once. → Task 1 + Task 7 tests.

---

### Task 1: Server detached-session grace (resume after a real drop)

**Files:**
- Modify: `server/src/ws.ts` (websocket `close` handler, `resumeSession`), `server/src/sessionTable.ts` (status `'detached'`, `detachedAt`)
- Test: `server/test/resume-detached.test.ts`

**Interfaces:**
- Produces: option `detachGraceMs` (default 30_000) on `createStreamServer`; session status `'detached'`.

- [ ] **Step 1: Failing tests.** (a) open + READY, `ws.close()` from the client side (abnormal: terminate the socket), then a new socket HELLO `{ resume: { sid, resumeProof: token } }` inside grace ⇒ MANIFEST with the **same sid**, a new jti, and `unit` for index 1 renders. (b) same after `detachGraceMs` (set 200 ms) ⇒ close 4003. (c) resume after the token's `exp` (ttl 2 s, grace 30 s) ⇒ 4003. (d) the doc ledger charge is still held during grace (`stats().documents` unchanged) and released after the reaper runs. (e) a server-initiated terminal close (revoke ⇒ 4005) leaves **no** detached session: resume ⇒ 4003.
- [ ] **Step 2:** `bun run test:server -- resume-detached` → FAIL.
- [ ] **Step 3:** In `close`: if the closing socket is active and the session is in `stream`/`manifest` and the close was not server-terminal, set `status = 'detached'`, `activeSocket = undefined`, schedule a reaper at `min(detachGraceMs, exp·1000 − now)` that performs today's teardown (table.close, rates, sessionDoc, releaseDoc) iff still detached. `resumeSession` accepts a detached session (no socket to supersede) and cancels the reaper. Terminal closes mark the session closed before the socket closes.
- [ ] **Step 4:** test file PASS; full `bun run test:server` PASS (ws-races unchanged).
- [ ] **Step 5:** Commit `feat(server): detached-session grace for resume after a drop`.

### Task 2: `StreamSession` — the client protocol state machine

**Files:**
- Create: `src/stream/remote/session.ts`, `src/stream/remote/errors.ts`, `src/stream/remote/reservation.ts`
- Test: `server/test/client-session.test.ts` (drives the real server via `start()`), `tests/stream/remote-session.test.ts` (Vitest, fake socket for guard/ordering cases)

**Interfaces:**
- Consumes: `Reassembler`, `Kind`, `FRAGMENT_CAP` (`src/stream/frame.ts`); `validateManifest`, `Manifest`; `decodeUnit`; `decodeAssets`; `decodeRasterBody`.
- Produces:
  ```ts
  type StreamClient = Pick<WebSocket, 'send' | 'close' | 'readyState' | 'addEventListener' | 'removeEventListener'> & { binaryType: BinaryType }
  type StreamClientFactory = (source: StreamSource) => StreamClient
  class StreamSessionError extends Error { code: StreamErrorCode }   // 'unsupported-version'|…|'superseded'|'oversized-frame'|'manifest-invalid'|'connect-failed'
  type UnitPayload = { kind: 'raster'; w: number; h: number; png: Uint8Array } | { kind: 'commands'; ops: Uint8Array; assets: ReadonlyMap<number, Asset> }
  class StreamSession {
    static open(source: StreamSource, o: { client?: StreamClientFactory; faces?: (m: Manifest) => Promise<string[]>; signal?: AbortSignal; onNotice?: (n: StreamNotice) => void; onState?: (s: 'open' | 'resuming' | 'closed') => void }): Promise<StreamSession>
    readonly manifest: Manifest
    unit(index: number, bucket: number, signal?: AbortSignal): Promise<UnitPayload>
    close(): void                       // idempotent
  }
  type StreamNotice = { code: string; unit?: number }  // manifest notices, 'face-missing', 'text-unavailable', 'resumed', 'unit-failed'
  ```
- `faces(manifest)` resolves to the `missingFaces` list; `open` races it against a 1.5 s timer and sends READY with whatever is unresolved counted missing.

- [ ] **Step 1: Failing tests (Bun, real server).** happy open ⇒ manifest validated, READY sent once with `tokenJti` = token jti; `unit(0,1)` on a protected pptx ⇒ `kind: 'raster'`, PNG chunk types ⊆ critical; relaxed session ⇒ `kind: 'commands'` with assets resolved from preceding ASSET_DEFINE; RENEW mid-flight (ttl 6 s, lead 5 s) ⇒ `renew-ack` sent, the unit promise issued before the RENEW still resolves, a later unit uses the new token; terminal close 4005 (`revoke`) rejects pending units with `StreamSessionError('revoked')`; abnormal drop ⇒ automatic resume (Task 1), pending unit re-issued once and resolves, `onNotice({code:'resumed'})`; second drop after grace ⇒ `auth-failed`.
- [ ] **Step 2: Failing tests (Vitest, fake socket).** a 1 MiB + 1 binary message ⇒ `close(4007)` and `Reassembler.push` never called (spy); a binary frame for a never-issued reqId ⇒ dropped (no reject of pending units); a frame for a unit whose `signal` aborted ⇒ dropped; `close()` during handshake ⇒ socket closed once, `open` rejects `AbortError`, no READY sent; READY is sent at 1.5 s when `faces` never resolves, listing all required ids; ninth concurrent `unit()` waits until one of eight settles; reservation (raster bucket² · w · h · 4 + 8 MiB encoded, or 8 + 9 MiB for commands, whichever is larger) never exceeds 96 MiB in flight; `session-rotated` refusal ⇒ one re-issue, a second ⇒ reject.
- [ ] **Step 3:** Implement. reqIds: monotonically increasing u32 from 1 (HELLO = 1, READY = 2, units from 3). One `Reassembler` per socket (a resumed socket gets a fresh one; pending units re-issue under new reqIds). `sweep()` every 1 s while any request is pending; a timed-out unit re-requests once then rejects `ReplayError`-class `unit-failed`. Exactly-once close through a generation counter: every socket listener checks `gen === this.gen`.
- [ ] **Step 4:** both test files PASS; `bun run test tests/stream` PASS.
- [ ] **Step 5:** Commit `feat(stream): client session state machine with receipt guard and resume`.

### Task 3: `RemoteOfficeDoc` + font pack hashing

**Files:**
- Create: `src/stream/remote/remote-doc.ts`, `src/stream/remote/faces.ts`, `src/stream/remote/open.ts` (`openRemoteDocument`), export from `src/stream/index.ts`
- Test: `server/test/client-remote-doc.test.ts` (napi backend installed via the preload), `tests/stream/remote-faces.test.ts`

**Interfaces:**
- Consumes: `StreamSession` (Task 2), `replayUnit` (`src/stream/replay.ts`), `canvasBackend()` (`src/core/canvas-backend.ts`).
- Produces:
  ```ts
  interface RemoteOfficeDoc {
    kind: 'remote'; manifest: Manifest
    units(): Promise<Array<{ widthPx: number; heightPx: number }>>
    paint(unit: number, scale: 'fit' | number, dpr: number, canvas: HTMLCanvasElement | OffscreenCanvas): Promise<void>
    search(q: string): Promise<never>; copy(range: unknown): Promise<never>   // reject StreamSessionError('policy-refused') + notice 'text-unavailable' once
    close(): void
  }
  function openRemoteDocument(source: StreamSource, o: { client?: StreamClientFactory; fonts?: { baseUrl: string }; signal?: AbortSignal; onNotice?: (n: StreamNotice) => void }): Promise<RemoteOfficeDoc>
  function resolveFaces(required: Manifest['recorder']['requiredFaces'], baseUrl: string | undefined, fetchImpl?: typeof fetch): Promise<string[]>   // returns missing ids
  function bucketFor(scale: number, dpr: number, buckets: readonly number[]): number
  ```

- [ ] **Step 1: Failing tests.** `bucketFor(1,1,[1,2])=1`, `(1,2)=2`, `(1.25,1)=2`, `(3,2)=2`. `resolveFaces`: no `baseUrl` ⇒ all ids missing; a served face whose SHA-256 first 12 bytes hex equals the manifest hash ⇒ registered via `FontFace(family, bytes, {weight, style})` from `pack.json` and not missing; a tampered byte ⇒ missing and **not** registered; a `pack.json` naming an unknown id ⇒ ignored. `paint` on a protected unit: canvas becomes `w·bucket × h·bucket` and its pixels equal the server PNG decoded (napi); on a relaxed unit with all faces present: pixels equal a direct `replayUnit` of the same payload; a corrupted ops payload ⇒ one re-request, then rejects and emits `{code:'unit-failed', unit}`; `search()` rejects and `text-unavailable` is emitted exactly once across two calls; `close()` closes the session once and later `paint` rejects.
- [ ] **Step 2:** FAIL.
- [ ] **Step 3:** Implement. Raster paint: `decodeRasterBody(body, { maxRgbaBytes: UNIT_RASTER_DECODED_CAP })` (admits w·h·4 before decode) ⇒ `canvasBackend().decodeImage(png, 'image/png')` ⇒ `drawImage` at 0,0 after `setTransform(1,0,0,1,0,0)` and `clearRect`. Commands: paint into an offscreen surface first, then copy, so a failing replay never half-paints the visible canvas. `units()` maps `manifest.units.list` (w,h are CSS px at scale 1).
- [ ] **Step 4:** PASS; `npx tsc --noEmit` clean.
- [ ] **Step 5:** Commit `feat(stream): RemoteOfficeDoc with raster/command paint and face hashing`.

### Task 4: Routing through `<OfficeFile>` / `<OfficeDoc>` + the remote view

**Files:**
- Modify: `src/components/OfficeFile.tsx` (`useOfficeFile` source type, `data` prop union, props `streamClient`, `streamFonts`, `onStreamNotice`), `src/components/loadOfficeSource.ts` (`AnyDoc` gains `RemoteOfficeDoc` as a type-only import), `src/components/OfficeDoc.tsx` (`OfficeDocSource` arm + dispatch), `scripts/eager-baseline.json` (recorded change)
- Create: `src/components/remote/RemoteDocView.tsx`
- Test: `tests/components/office-file-stream.test.tsx` (Vitest jsdom; `vi.mock('../../src/stream/remote/open')`)

**Interfaces:**
- Consumes: `openRemoteDocument`, `RemoteOfficeDoc` (Task 3); `isStreamSource`.
- Produces: `OfficeFileProps.data: ByteSource | StreamSource | Promise<ByteSource | StreamSource> | null | undefined`; `OfficeFileState` ready arm may carry a `RemoteOfficeDoc`; `RemoteDocView` props `{ document: RemoteOfficeDoc; theme?; className?; style?; pageGapPx?; allowPrint? }`.

- [ ] **Step 1: Failing tests.** stream source ⇒ `openRemoteDocument` called once with the source, `fetch` never called, `WebSocket` (global spy) not constructed by the byte path; ready ⇒ `RemoteDocView` rendered (role `document`), one `<canvas>` per unit, `[data-officeview-notice]` contains "text unavailable in this viewing mode" exactly once, no search input, no copy button; host `watermark` prop has no effect; replacing the source mid-open ⇒ the first document's `close()` called exactly once (even when it resolves after the replacement), and unmount ⇒ `close()` once; a **byte** source ⇒ `openRemoteDocument` never imported (module mock not invoked) and the rendered tree equals today's (`LegacyOfficeDoc` path); `ref.present()` on a remote doc rejects `not-presentable`; a `StreamSessionError('auth-failed')` ⇒ `status: 'error'` with `code: 'auth-failed'`.
- [ ] **Step 2:** FAIL.
- [ ] **Step 3:** Implement. `useOfficeFile`: `if (isStreamSource(src))` ⇒ `import('../stream/remote/open').then(m => m.openRemoteDocument(src, …))`, cleanup calls `close()` on the resolved doc (or on resolution if it lands after cleanup). `OfficeDoc`: `if ('kind' in document && document.kind === 'remote')` ⇒ `<Suspense><RemoteDocView …/></Suspense>` with `RemoteDocView = lazy(() => import('./remote/RemoteDocView'))`, handle `SHEET`-style `{ present: () => Promise.reject(new SlideShowError('not-presentable')) }`. `RemoteDocView`: vertical page column, an `IntersectionObserver` paints units entering the viewport (± 1 unit) at `bucketFor(zoom, devicePixelRatio)`; units are CSS-sized `w × h`; zoom controls reuse the existing zoom cluster props only if trivially available, otherwise none in Stage 1 (record which).
- [ ] **Step 4:** Test PASS; `bun run test` (root) PASS with no change in any existing test; `bun run gates:core`: record the measured delta as a second recorded change (`reason: "§6.3 stream routing (OfficeFile/OfficeDoc branch + lazy remote imports)"`, `added`/`removed` from the gate's diff output — only the two routing files' module ids may change, nothing under `src/stream/remote/` or `src/components/remote/`) and re-run ⇒ all ok.
- [ ] **Step 5:** Commit `feat(components): route StreamSource to a lazy remote viewer`.

### Task 5: §7.7 occlusion evidence on the server paint pass

**Files:**
- Create: `server/src/occlusion.ts`, `server/test/occlusion.test.ts`
- Modify: `server/src/render.ts` (`PaintObserver` gains optional `wrap?(ctx): CanvasRenderingContext2D`; `RenderedUnit` gains `evidence: OcclusionEvidence`), `server/src/worker/render-worker.ts` (always attach the observer), `docs/validation/streaming-s1.md`

**Interfaces:**
- Produces:
  ```ts
  interface OcclusionEvidence { runs: number; servable: number; withheld: number; reasons: Partial<Record<'covered' | 'clipped' | 'off-surface' | 'composite' | 'filter' | 'nonrect-clip' | 'transform', number>> }
  function occlusionObserver(cellPx?: number /* 4 */): PaintObserver & { evidence(): OcclusionEvidence }
  ```
- Rule (fail-closed, normative §7.7): each `fillText`/`strokeText` run records device-space ink bounds (transformed `measureText` actual bounding box, axis-aligned hull). The opaque-coverage grid (cells of `cellPx`) marks, **after** each later draw op whose effective `globalAlpha ≥ 0.95` under `source-over`, every cell its conservative device bounds touch (shadow offset + blur included; fill style of unknown opacity counts as opaque). A run is servable iff its ink lies inside the surface, inside the current clip (tracked exactly only while every clip is an axis-aligned rect under an axis-aligned transform; otherwise `nonrect-clip`), and no ink cell is covered by a later op. Any later non-`source-over` op or `filter` intersecting the ink ⇒ withheld (`composite`/`filter`). The visible mark is painted after evidence is taken. Evidence only — nothing culls or ships it in Stage 1.

- [ ] **Step 1: Failing tests** (napi canvas): text then opaque `fillRect` over it ⇒ withheld `covered`; text then the same rect at `globalAlpha 0.5` ⇒ servable; text partly outside a rect clip ⇒ `clipped`; text under an arc clip ⇒ `nonrect-clip`; text then `destination-out` rect elsewhere not intersecting ⇒ servable, intersecting ⇒ `composite`; text drawn past the right edge ⇒ `off-surface`; rotated text with no later ops ⇒ servable; the 13 T2 hidden fixtures (`visible-only` projection) report `withheld = 0` for their control text runs and every fixture's evidence `runs > 0`.
- [ ] **Step 2:** FAIL.
- [ ] **Step 3:** Implement as a `Proxy` over the context (the `census-proxy.ts` pattern); state tracked through `save`/`restore`.
- [ ] **Step 4:** PASS; run the corpus census (`bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s1/occlusion-census.ts`, new, prints per-kind runs/servable/withheld/reasons over `corpus/` s1 files at scale 1) and add the table to `docs/validation/streaming-s1.md` under "Stage 1d — occlusion evidence".
- [ ] **Step 5:** Commit `feat(server): §7.7 ink-bounds and opaque-coverage evidence`.

### Task 6: Local byte-for-byte identity

**Files:**
- Create: `scripts/stream-s1/local-identity.ts`, `docs/validation/streaming-s1/local-identity.json` (baseline hashes)
- Test: `tests/stream/local-unchanged.test.tsx` (Vitest jsdom)

- [ ] **Step 1:** `local-identity.ts --write <out>` paints every unit of every `corpus/` file through `loadOfficeFile` → `getPaintables` with the napi backend at scale 1 and writes `{ file: [sha256(rgba) per unit] }`; without `--write` it compares and exits 1 on any difference. Generate the baseline in a throwaway worktree at `ca952c50` (`git worktree add /tmp/ov-base ca952c50`, run with that tree's sources, then remove the worktree).
- [ ] **Step 2:** Run the compare on this branch ⇒ **0 differing units** (any difference: probe and report; never loosen).
- [ ] **Step 3: Test.** Mounting `<OfficeFile data={bytes}>` for a docx, xlsx and pptx: no `WebSocket` constructed, `src/stream/remote/open` never imported (vi.mock factory spy), no `[data-officeview-notice]` added by the stream path.
- [ ] **Step 4:** PASS. Record the result in `docs/validation/streaming-s1.md` (files, units, differing units, baseline commit).
- [ ] **Step 5:** Commit `test(stream): local loads byte-for-byte unchanged`.

### Task 7: Playwright WS harness (`server/test/app.ts`) + browser acceptance

**Files:**
- Create: `server/test/app.ts`, `src/debug/stream-viewer.html`, `src/debug/stream-viewer.tsx`, `tests/browser/stream-viewer.spec.ts`
- Modify: `playwright.viewer.config.ts` (project `stream`, `testMatch: /stream-.*\.spec\.ts/`; the existing projects ignore it; a second `webServer` entry runs `bun --preload ./scripts/stream-s0/napi-preload.ts server/test/app.ts` on `OV_STREAM_PORT` default 4181), `package.json` script `"test:browser:stream": "playwright test --config playwright.viewer.config.ts --project stream"`

**Interfaces:**
- `server/test/app.ts`: the 1c server in demo mode over `corpus/` with test-only controls on a separate port (`OV_STREAM_PORT + 1`, never part of `server/src`): `POST /drop?sid=` terminates the active socket (abnormal), `POST /revoke?sid=`, `POST /oversize?sid=` sends one 1 MiB + 1 binary message, query `?ttl=<s>` on the WS URL shortens the token lifetime; `GET /sessions` lists sids. Serves `fontpack/` at `/fontpack/` for `streamFonts`.
- `stream-viewer.tsx`: mounts `<OfficeFile data={{ kind: 'stream', url }}>` from query `?doc=&relaxed=&ttl=`, exposes `window.__ovStream = { state, notices, sid }` for assertions.

- [ ] **Step 1: Specs.** happy (protected pptx): N canvases = manifest count, first unit has non-background pixels, notice text present once, no element with selectable text inside the viewer; relaxed docx with `streamFonts`: units paint via commands (no `unit-failed` notice), and the published Chromium-vs-napi per-unit pixel difference is logged (not gated — §7.4 rule 5 tolerance is Stage-1 published only); error: unknown `doc` ⇒ error state `auth-failed` rendered; expiry: `ttl=8` and the page stays live for 20 s with ≥ 2 renews (server log) and no error; reconnect: scroll to unit 4 of a long docx, `POST /drop`, wait for notice `resumed`, `scrollTop` unchanged (± 1 px), the visible canvases repainted (new unit requests observed in `/sessions` stats), same sid; oversize: `POST /oversize` ⇒ error state `oversized-frame`; revoke ⇒ `revoked`.
- [ ] **Step 2:** `bun run test:browser:stream` → FAIL.
- [ ] **Step 3:** Implement the app and page.
- [ ] **Step 4:** PASS twice; existing `bun run test:browser` project counts unchanged.
- [ ] **Step 5:** Commit `test(browser): stream viewer harness — happy, error, expiry, reconnect`.

### Task 8: Close-out — validation, review, docs

- [ ] **Step 1:** Full verification: `bun run test`, `bun run test:server`, `bun run test tests/stream`, `bun run s0:test:napi`, `bun run test:browser:stream`, `npx tsc --noEmit` (root + server), `bun run gates:core`, `git diff --check`. Record counts in `docs/validation/streaming-s1.md` "Stage 1d".
- [ ] **Step 2:** Stage 1 exit check against §17 row 1: corpus zero diff (1b), T2 13/13 on the stream path (1c), close-code matrix (1c + Task 2), T11 gates, RemoteOfficeDoc branch, occlusion evidence, local identity — one table, each row linked to its evidence.
- [ ] **Step 3:** Codex review of the 1d diff (single checkpoint; quota). Fix confirmed findings with regression tests.
- [ ] **Step 4:** `server/README.md` + public docs: `StreamSource`, `streamClient`, `streamFonts`, `onStreamNotice`, Stage 1 limits (no search/copy/present, forensic mark Stage 2).
- [ ] **Step 5:** Commit, push; propose the Stage 1 PR to the owner (do not merge).

## Deferred (named so nothing is silently dropped)

Stage 2: forensic layer-1, command-culling store consuming Task 5's evidence, scratch lease, parse worker, audit retention/key rotation. Stage 3: search/copy/a11y services (they will consume the evidence), print jobs, notes. Stage 3b+: remote presentation (`present()` stays refused until then). Stage 4: XLSX tiles (oversized units remain refused). Cross-unit asset dedupe on the wire.
