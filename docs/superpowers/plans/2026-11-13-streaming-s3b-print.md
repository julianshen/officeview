# Streaming Stage 3b — print Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A streamed document prints from server-drawn, both-marked print rasters, within a developer-set page cap per job and per session, refused (never truncated) over the cap, and audited per job.

**Architecture:** A new `print` text op asks for a unit range. The server checks the whole job up front (range, caps, surface budgets), audits it, charges the session total, then drives at most 2 page renders at a time through the existing pool. Each page gets the session's Layer-1 mark and a `deliveries` row, and ships as a `PRINT_RASTER` (kind 8) frame in ascending page order, bracketed by `print-start` / `print-done` text messages. The client checks every page against the job, holds PNG blobs (not RGBA), and prints them from the local viewer's hidden-iframe surface, extracted so both paths share it. The on-screen remote view never prints.

**Tech Stack:** Bun server (`server/`, `bun test`), TypeScript client (`src/stream/remote`, `src/components/remote`, vitest), Playwright `stream` project.

**Spec:** `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.2 (parent `docs/superpowers/specs/2026-11-07-server-streaming-design.md` §2 item 9, §6.1, §6.2 rules 7–8, §7.1, §13.1, §14, §19). Piece 1 (`docs/superpowers/plans/2026-11-11-streaming-s3a-text-services.md`, merged 2db40c48) is the pattern to follow: copy caps → print caps, `copies` → `prints`, `'text'` rate class → `'print'`.

## Rulings made while planning (spec §4.2 amendment, Task 8 writes it into the spec)

These are the planner's calls, listed for the owner at handoff.

- **P1 — print resolution is scale 2, falling back to 1.** The fallback happens per job when any page's scale-2 surface fails `unitSurfaceAllowed` or the job budget. This matches the local viewer's `PRINT_SCALE = 2` / `planPrintJob`.
  - Why: the Stage 2b tracer (`server/src/detect/trace.ts:99-109`) re-renders the reference with `host.render` at the stored delivery scale, and keys chips on `unit@scale`. Any scale outside `SCALE_BUCKETS` would need render and tracer changes.
  - Why not higher: the forensic encode estimate (`w*h*16`) at scale 3 for a Letter page is about 124 MB, over the 96 MiB session budget.
  - Cost if wrong: a later "higher DPI print" adds a print-only scale to `DocumentHost` and the tracer.
- **P2 — "print units never reuse screen rasters" means the client never prints what it shows.** The server may serve print pages from the canonical render cache. Those bytes are exactly what the tracer reproduces.
  - Layer-1 is applied per print page, per request, as for screen.
  - A `deliveries` row deduplicates against a screen delivery of the same unit and scale (`INSERT OR IGNORE`). That is correct for tracing: same reference, same session code.
  - Cost if wrong: one `bypassCache` flag on the render call.
- **P3 — charging.** A job is validated whole before any frame.
  - The order is: validate, rate, caps and quota, audit row, charge, `print-start`, pages.
  - A job that fails mid-way (render error, socket loss, viewer cancel) prints nothing on the client, and its pages stay charged. This is the same rule as a lost copy answer (§4.1 item 11).
  - Jobs are serialised per sid. A `print` while a job runs is refused `rate-limited`.
- **P4 — print caps may change on resume**, as copy caps do: spec §4.1 says "the same rules apply to `maxSessionPages`".
  - `print` leaves `resumeInvariant`.
  - A change is recorded through `recordServices`.
  - The pages already printed are kept and measured against the new caps.
- **P5 — a page is a streamed unit, under every content policy.** Under `visible-only`, hidden slides are dropped before parsing (`server/src/policy/pptx.ts:12-43`), so they are never units and never pages.
  - Under audited `contentPolicy: 'full'`, the local paged engine draws `show="0"` slides as pages. Only the slideshow skips them (`src/slideshow/read/deck.ts:120`).
  - So under `'full'` they print, exactly as they show. This follows §4.1's "which text is visible follows the session".
  - This narrows §4.2's "hidden slides are not pages" to `visible-only` sessions.
  - Cost if wrong: Task 3 filters `show="0"` units out of print ranges under `'full'`.
- **P6 — a page Layer-1 cannot mark (`unmarked: 'capacity' | 'all-masked'`) prints as it shows.** It ships with the visible mark only and an `unmarked` delivery row, the same as on screen.
  - Refusing it would block printing a page the viewer can already see and screenshot.
  - Cost if wrong: refuse such pages with `policy-refused` before charging. That needs a pre-render pass, so it is not cheap.
- **P7 — print off is refused `policy-refused`.** No new code: the client never sends `print` when `can.print` is false.

## Global Constraints

- Policy: `print?: boolean | { maxPages: number; maxSessionPages?: number }`, default `true` (unchanged). Invalid caps fail closed to `false`, like `copyPolicy()`.
- Refusal codes: `print-too-large` and `print-quota-exhausted`, carrying `{limit, requested}`. Never truncate.
- Rate class `'print'`: 10 jobs/min per sid, refuse-only (never counts a breach, never closes), like `'text'`. At most 2 page renders in flight per sid.
- PRINT_RASTER caps: 8 MiB encoded, 64 MiB decoded, per parent §14. Every new text message is at most 48 KiB, like piece 1.
- Session caps are per sid in server memory. They survive resume and reset on a new open. The total is a running sum: printing a page twice counts twice.
- Fail closed: anything that does not match the job exactly fails the job, and nothing prints.
- The client has no relax surface. New client code lives only behind the existing lazy remote imports. `bun run gates:core` must not grow from this branch; report the number either way. It was already red on main at 421,066 B against 420,540 B.
- Local (non-remote) print behaviour is byte-for-byte unchanged. The existing `tests/**/printDocument*` and paged print tests must pass unmodified.
- Commits end with the session's attribution lines. Never `git commit -a`; never bare `git stash`.

## Review Focus

1. **Screen canvases print anyway** (browser menu Print, not Ctrl+P). Expected: the remote root is hidden in print media at all times, so the page prints blank and never shows the screen rasters. Pinned in Task 7.
2. **A hostile or buggy server sends a page that wasn't asked for**: wrong unit index, a duplicate, out of order, too big, too many, after `print-done`, or `print-done` early. Expected: the job fails, nothing prints, and the session survives. Pinned in Task 5.
3. **The viewer cancels mid-job, or the socket drops mid-job.** Expected: no print dialog. The server stops rendering after the in-flight pages. The quota stays charged. Pinned in Tasks 3 and 5.
4. **Resume with tighter caps after some pages were printed.** Expected: the next job is measured against the new caps using the kept total. Pinned in Task 1 (policy) and Task 3 (quota).
5. **A huge uncapped document** (hundreds of slides, or an XLSX sheet over the surface limit). Expected: the scale-1 fallback, or a `print-too-large` refusal with no `limit`, before anything is charged. The client's 192 MiB blob ceiling ends the job with the local "too large" notice. Pinned in Tasks 3 and 5.

---

### Task 1: Print policy, manifest caps, resume

**Files:**
- Modify: `server/src/hooks.ts` (policy type :16, `EffectivePolicy` :39, `resumeInvariant` :75-78, resolution :95)
- Modify: `server/src/ws.ts` (manifest `can`/`caps` :296-297; resume `recordServices` :251)
- Modify: `server/src/audit.ts` (`TextServices` gains `print`), `server/src/audit-store.ts` (`services` table gains a `print` JSON column)
- Modify: `src/stream/manifest.ts` (`Manifest.caps.print`, validation)
- Test: `server/test/sessions.test.ts`, `server/test/text-policy.test.ts` (or new `server/test/print-policy.test.ts`), `server/test/ws-races.test.ts`, `tests/stream/manifest.test.ts`

**Interfaces:**
- Produces:
  - `export type PrintPolicy = false | { maxPages: number | null; maxSessionPages: number | null }` and `export function printPolicy(p: unknown): PrintPolicy` in `server/src/hooks.ts`.
  - `EffectivePolicy.print: PrintPolicy`.
  - Manifest: `can.print = policy.print !== false`; `caps.print: null | { maxPages: number | null; maxSessionPages: number | null }`.
  - Client: `Manifest['caps']['print']`.

- [ ] **Step 1: Write failing server tests.** In `print-policy.test.ts`:
  - `printPolicy(undefined)` and `printPolicy(true)` → `{maxPages: null, maxSessionPages: null}`.
  - `printPolicy(false)` → `false`.
  - `printPolicy({maxPages: 3})` → `{maxPages: 3, maxSessionPages: null}`.
  - `printPolicy({maxPages: 3, maxSessionPages: 5})` → both.
  - Each of `{maxPages: 0}`, `{maxPages: 1.5}`, `{maxPages: -1}`, `{maxPages: 3, maxSessionPages: 0}`, `{}`, `'x'` → `false`.
  - A session opened with `policy: {print: {maxPages: 2, maxSessionPages: 4}}` gets a manifest with `can.print === true` and `caps.print` deep-equal to `{maxPages: 2, maxSessionPages: 4}`.
  - A `print: false` session gets `can.print === false` and `caps.print === null`.
  - Update `sessions.test.ts:44`'s expectation for `print:false` to the new resolved shape.
- [ ] **Step 2: Write failing resume tests** in `ws-races.test.ts`.
  - The existing "changed policy refused" case (:42-45) uses `print:false` as its change. Switch it to a field that stays invariant (e.g. `visibleMark: 'viewer'`), keeping its 4003 assertion.
  - Add: resuming with `print` changed from `true` to `{maxPages: 1}` is accepted. The resumed manifest carries the new `caps.print`, and `store.services(sid)` (or the test's equivalent read) has a row whose `print` is `{maxPages: 1, maxSessionPages: null}`.
- [ ] **Step 3: Write a failing client test** in `tests/stream/manifest.test.ts`.
  - `caps.print` null passes; `{maxPages: 2, maxSessionPages: null}` passes.
  - `{maxPages: 0}`, `{maxPages: 'x'}` and a missing `caps.print` are rejected, using the same `cap()` rule as `caps.copy`.
  - Update `exampleManifest()` in `tests/stream/remote-session.test.ts` and any other fixture to carry `caps.print`.
- [ ] **Step 4: Run and watch them fail.** Run `cd server && bun test test/print-policy.test.ts test/ws-races.test.ts test/sessions.test.ts`, then `bunx vitest run tests/stream/manifest.test.ts`. Expected: FAIL (no `printPolicy`; `caps.print` missing).
- [ ] **Step 5: Implement.**
  - `printPolicy` mirrors `copyPolicy`/`isCap`. `true`/`undefined` mean uncapped; anything not a valid cap object or boolean means `false`.
  - `resumeInvariant` omits `print` too.
  - Resume calls `recordServices(sid, {copy, textA11y, print})`.
  - Add a `print TEXT` column to `services` in the `CREATE TABLE`. It is an in-memory/new DB with no migration path today; if `AuditStore` has a schema version, bump it.
  - Client `validateManifest` validates `caps.print` like `caps.copy`. It is required, so an old server's manifest fails closed.
- [ ] **Step 6: Run green.** Run the same commands, then the full `cd server && bun test` and `bunx vitest run tests/stream`. Expected: all pass.
- [ ] **Step 7: Commit.** `feat(stream): print policy with page caps, carried in the manifest and changeable on resume`

### Task 2: Print audit — `prints` table and row totals

**Files:**
- Modify: `server/src/audit-store.ts` (SCHEMA :28-39, `expire` :148-154), `server/src/audit.ts` (`AuditRow`, `AuditTable`)
- Test: `server/test/audit-store.test.ts`

**Interfaces:**
- Produces:
  - `export interface PrintRecord { sid: string; pages: number; fromUnit: number; toUnit: number; scale: 1 | 2; at: number }`
  - `AuditStore.putPrint(r: PrintRecord): void` and `AuditStore.prints(sid: string): PrintRecord[]`
  - `AuditTable.recordPrint(r: Omit<PrintRecord, 'at'>): void`, which throws for an unknown or expired sid, like `recordCopy`.
  - `AuditRow.printed: { jobs: number; pages: number }`

- [ ] **Step 1: Write failing tests.**
  - `recordPrint` twice for one sid → `store.prints(sid)` returns both in order with `at` set, and `table.get(sid).printed` is `{jobs: 2, pages: <sum>}`.
  - `recordPrint` for an unknown sid throws.
  - `expire(cutoff)` removes `prints` rows older than the cutoff along with the other tables.
  - A fresh row has `printed: {jobs: 0, pages: 0}`.
- [ ] **Step 2: Run and watch them fail.** Run `cd server && bun test test/audit-store.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
  - Add the table `prints (sid TEXT NOT NULL REFERENCES rows (sid), pages INTEGER NOT NULL, from_unit INTEGER NOT NULL, to_unit INTEGER NOT NULL, scale REAL NOT NULL, at INTEGER NOT NULL)` with an index `prints_sid`.
  - Add the append and read methods, following `putCopy`/`copies`.
- [ ] **Step 4: Run green**, including the full `cd server && bun test`.
- [ ] **Step 5: Commit.** `feat(stream): audit every print job with its page count, range and scale`

### Task 3: Server `print` op — job check, quota, render driver, PRINT_RASTER

**Files:**
- Create: `server/src/print.ts` (job planning plus the per-sid driver; the parent §9 layout names `print.ts`)
- Modify: `server/src/ws.ts` (dispatch :580-587, `KIND` :41, `ENCODED` :132, `RefusalCode` :39, per-sid slot, teardown)
- Modify: `server/src/rate.ts` (`RateClass` gains `'print'`, window 10/60 s, refuse-only)
- Modify: `server/test/harness.ts` (`CAPS` gains kind 8; add `Client.print(reqId, from, to)` returning the `print-start`, the kind-8 bodies and the closing text message)
- Test: create `server/test/print-wire.test.ts`

**Interfaces:**
- Consumes: `printPolicy`/`EffectivePolicy.print` (Task 1); `AuditTable.recordPrint` (Task 2); existing `renderCache`, `keyFields`, `pool.render`, `markUnit`, `audit.recordDelivery`, `unitSurfaceAllowed`, `sendFrames`.
- Produces the wire, which Task 5 consumes verbatim:
  - Client → server: `{ reqId: u32, type: 'print', token: string, from: int, to: int }`. `from` and `to` are 0-based inclusive unit indices with `0 ≤ from ≤ to < units.count`.
  - Client → server: `{ reqId, type: 'print-cancel' }`, naming the running job's reqId. Unknown reqIds are ignored.
  - Server → client: `{ reqId, type: 'print-start', pages: n, scale: 1 | 2 }` where `n = to - from + 1`.
  - Server → client: `n` kind-8 messages on the same reqId, in ascending unit order. The body is `u32 LE unitIndex` followed by the existing raster body (`src/stream/raster.ts`: `u32 w, u32 h, u32 pngLen, PNG`). The raster is the renderer's `sizeCeil(raw*scale)`, which can be under `ceil(unit.w*scale)` for a fractional raw page size (the manifest rounds first). The contract (Task 3 fix round 1 amendment) is `(unit.w - 1) * scale < w ≤ unit.w * scale`, and the same for `h`, where `unit` is the manifest's size. The server fails the job `policy-refused` for a page outside it.
  - Server → client: `{ reqId, type: 'print-done', pages: n }`.
  - Failure at any point: `{ reqId, type: 'refused', code, limit?, requested? }`, and nothing further for that reqId.
  - `export const PRINT_RATE_PER_MIN = 10`, `PRINT_CONCURRENCY = 2`, `PRINT_ENCODED_CAP = 8 * 1024 * 1024`, `PRINT_JOB_MAX_DECODED = 1024 * 1024 * 1024` in `server/src/print.ts`.
  - `export function planPrint(units: readonly UnitSize[], from: number, to: number): { scale: 1 | 2 } | { refused: 'too-large' }`. It tries 2 then 1. Every page must pass `unitSurfaceAllowed` at that scale, and the sum of `w*h*4` must be ≤ `PRINT_JOB_MAX_DECODED`.
  - The session total lives on the per-sid slot as `printedPages: number`, next to `copiedParagraphs`, and is created at open, kept on resume and dropped in teardown.

- [ ] **Step 1: Write failing wire tests** in `print-wire.test.ts`, using the piece-1 spin-up pattern (`AuditStore.open(':memory:')`, `start({audit, policy})`, `Client`, `hello`, `ready`) on `corpus/pypptx-test.pptx`.
  - **Happy path:** with uncapped print, `print 0..1` gives `print-start {pages: 2, scale: 2}`, then two kind-8 bodies with unitIndex 0 then 1, each decodable by `decodeRasterBody` at a size within `(unit - 1) * 2 < w ≤ unit * 2` (and the same for `h`), then `print-done {pages: 2}`. `store.prints(sid)` is `[{pages: 2, fromUnit: 0, toUnit: 1, scale: 2}]`, and `deliveries` has a row for each page at scale 2.
  - **Per-job cap boundary:** with `{maxPages: 2}`, `0..1` is allowed and `0..2` is refused `{code: 'print-too-large', limit: 2, requested: 3}`. Nothing was charged: a following `0..1` is still allowed. The refused job left no `prints` row.
  - **Session cap boundary:** with `{maxPages: 5, maxSessionPages: 4}`, jobs of 2 then 2 are allowed. Then `0..0` is refused `{code: 'print-quota-exhausted', limit: 4, requested: 1}`.
  - **Running total:** with `{maxPages: 5, maxSessionPages: 2}`, printing page 0 twice is allowed and a third time is refused `print-quota-exhausted`.
  - **Resume keeps the total:** under `maxSessionPages: 4`, print 3 pages. Resume with `maxSessionPages: 3`, then `0..0` is refused `print-quota-exhausted {limit: 3, requested: 1}`.
  - **Malformed range:** each of `from > to`, `to ≥ count`, `from < 0`, `1.5`, a missing `to`, or a bad token → refused `policy-refused`. Nothing is charged and there is no `prints` row.
  - **Print off:** `print: false` → refused `policy-refused`.
  - **Rate:** 10 jobs inside a minute are allowed (use 1-page jobs, each awaited to `print-done`). The 11th is refused `rate-limited`, the socket stays open, and the next ordinary `unit` request still succeeds.
  - **One job at a time:** a second `print` sent before the first's `print-done` is refused `rate-limited`, and the first completes normally.
  - **Cancel:** with an uncapped job over all units of a multi-unit deck, send `print-cancel` right after `print-start`. The server sends no `print-done`, and at most `PRINT_CONCURRENCY` kind-8 messages arrive after the cancel. The job's pages stay charged in `prints`.
  - **Too large:** with an injected `planPrint` refusal (a test hook, or a doc whose unit at scale 1 exceeds `unitSurfaceAllowed`), the job is refused `{code: 'print-too-large'}` with no `limit`, and nothing is charged.
  - **Concurrency:** instrument the pool (a wrapping `render` that counts in-flight calls for this sid) so that a 6-page job never has more than 2 renders in flight.
  - **Teardown:** closing the socket mid-job stops further `pool.render` calls for that sid after the in-flight ones.
- [ ] **Step 2: Run and watch them fail.** Run `cd server && bun test test/print-wire.test.ts`. Expected: FAIL (`policy-refused` for the unknown op).
- [ ] **Step 3: Implement `planPrint` and the driver in `server/src/print.ts`.**
  - Order (P3): parse and validate, `rates.check('print')`, policy off → `policy-refused`, job already running → `rate-limited`, `planPrint`, per-job cap, session quota, `recordPrint` (a throw → `policy-refused`), charge `printedPages += n`, send `print-start`.
  - Render: at most 2 pages in flight through the same cache-then-pool-then-`markUnit`-then-`recordDelivery` sequence `onUnit` uses, factored into one shared helper rather than duplicated.
  - Send: completed pages go out strictly in ascending order. Buffer at most `PRINT_CONCURRENCY` finished pages.
  - Failures: any page failure (`AdmissionRefused` or `LeaseFull` → `rate-limited`, anything else → `policy-refused`) sends one `refused` and stops the job.
  - Cancel and teardown abort the driver.
  - Add `ENCODED[8] = PRINT_ENCODED_CAP` and `KIND.PRINT_RASTER = 8`. Every `print-start`/`print-done` passes the 48 KiB message check.
- [ ] **Step 4: Run green**, then the full `cd server && bun test` and `bun run typecheck:server`. Expected: all pass; report the counts.
- [ ] **Step 5: Commit.** `feat(stream): server print jobs — checked whole, audited, capped, two renders at a time`

### Task 4: Evidence — both marks and no hidden content on print rasters

**Files:**
- Test: create `server/test/print-evidence.test.ts` (patterns: `server/test/forensic-serve.test.ts:18-52`, `server/test/ops.test.ts:76` T2 over the wire, `server/test/trace.test.ts`)

**Interfaces:**
- Consumes: the Task 3 wire and harness `Client.print`.

- [ ] **Step 1: Write the forensic test.**
  - In a forensic-on session, print unit 0. Decode the kind-8 body, then `decodePng`.
  - Build chips with `basis(deriveBasisKey(keys.kBasisRoot, manifest.docRef, '0@2'), w, h, DEFAULT_PARAMS)`.
  - Use the reference `h.stream.canonicalOf(sid, 0, 2)`, with the session's code plus 62 decoys.
  - Expect `decideScores(scoreCapture(...), {abs: 7.99, margin: 3.85})` to be `{kind: 'attributed', index: <session>}`.
  - Also run the stored-delivery trace path once: `trace()` on the printed PNG resolves to this sid.
- [ ] **Step 2: Write the visible-mark test.** With forensic off, the printed page's PNG is byte-equal to `canonicalOf(sid, 0, 2)` (the visible-marked canonical render). This proves the visible mark is present, and that the page matches a screen render at scale 2 (ruling P2).
- [ ] **Step 3: Write the T2 test.**
  - For each of the 13 hidden-construct fixtures `ops.test.ts:76` uses, open a `visible-only` session and print every unit.
  - Each printed page (forensic off) must be byte-equal to the canonical render of the same unit at the same scale, which the existing T2 raster tests already establish as hidden-free.
  - No `prints` job may include a dropped hidden slide: the unit count equals the manifest's count, and the fixture's hidden slide is not among them.
- [ ] **Step 4: Run.** Run `cd server && bun test test/print-evidence.test.ts`. Expected: PASS. Tasks 1–3 already provide the behaviour, so these tests are evidence rather than RED-first. If any fails, that is a Task 3 defect: fix it there, with a failing test first.
- [ ] **Step 5: Commit.** `test(stream): print rasters carry both marks and no hidden content`

### Task 5: Client session — `print()` with strict job checks

**Files:**
- Modify: `src/stream/remote/session.ts` (`CAPS` :50-53 add kind 8 [8 MiB, 64 MiB]; `expects` :267 admits kind 8 only for the active print reqId; dispatch for `print-start`/`print-done`; refusal mapping)
- Modify: `src/stream/remote/errors.ts` (`StreamErrorCode` gains `'print-too-large' | 'print-quota-exhausted'`)
- Modify: `src/stream/remote/remote-doc.ts` (`RemoteOfficeDoc.print`)
- Create: `src/stream/remote/print-wire.ts` (page body decode and job checks, like `text-wire.ts`)
- Test: create `tests/stream/remote-print.test.ts` (scripted `FakeSocket`, as in `tests/stream/remote-session.test.ts`)

**Interfaces:**
- Consumes: the Task 3 wire, verbatim.
- Produces:
  - `export interface PrintedPage { unitIndex: number; widthPx: number; heightPx: number; scale: 1 | 2; png: Blob }`. `widthPx`/`heightPx` are the unit's CSS size from the manifest.
  - `StreamSession.print(from: number, to: number, opts: { signal?: AbortSignal; onProgress?(done: number, total: number): void }): Promise<PrintedPage[]>`
  - `RemoteOfficeDoc.print(range: { from: number; to: number }, opts): Promise<PrintedPage[]>`, a thin pass-through.
  - `export const PRINT_JOB_MAX_BLOB_BYTES = 192 * 1024 * 1024` is reused from `src/components/paged/printDocument.ts`. Do not import that module statically into `session.ts`. Duplicate the constant in `print-wire.ts` with a comment naming its twin.
  - Errors: `StreamSessionError` with `code` from `print-too-large`, `print-quota-exhausted`, `rate-limited`, `policy-refused`, or `'aborted'`/`'session-ended'` (use the existing names). Exceeding the blob ceiling gives `print-too-large` with no `limit`.

- [ ] **Step 1: Write failing tests.**
  - **Happy path:** a scripted `print-start {pages: 2, scale: 2}`, two well-formed kind-8 pages 0 and 1, then `print-done` → resolves 2 `PrintedPage`s in order with `png.type === 'image/png'`. `onProgress` is called (1,2), (2,2). The request sent was `{type: 'print', from: 0, to: 1}` with the token.
  - **Refusals:** `refused print-too-large {limit: 2, requested: 3}` rejects with exactly those fields; `print-quota-exhausted` likewise. An unknown refusal code maps to `policy-refused`.
  - **Hostile cases.** Each rejects `policy-refused`, leaves no pending job, and leaves the session usable (a following `unit()` succeeds):
    - `print-start.pages ≠ to-from+1`
    - a scale other than 1 or 2
    - a page whose unitIndex is not the next expected one (wrong index, duplicate, out of order)
    - a page whose raster is outside `(unit.w - 1) * scale < w ≤ unit.w * scale` or `(unit.h - 1) * scale < h ≤ unit.h * scale`
    - a page before `print-start`
    - `print-done` before all pages
    - more pages than announced
    - `print-done.pages ≠ n`
  - **Late frames:** kind-8 frames arriving after the job settled are rejected by `expects` (`unexpected-reqId`) and never revive it.
  - **Blob ceiling:** with `PRINT_JOB_MAX_BLOB_BYTES` injected small, the job rejects `print-too-large` with no `limit`, sends `print-cancel`, and has released the held blobs (count retained pages = 0).
  - **Abort:** aborting the signal after `print-start` sends `{type: 'print-cancel', reqId}`, rejects with the abort code, and removes the signal's listener. Count listeners as in the piece-1 abort test.
  - **Socket end:** a socket ending mid-job rejects with the session-ended code.
  - **One job at a time on the client:** a second `print()` while one is pending rejects immediately, without sending.
- [ ] **Step 2: Run and watch them fail.** Run `bunx vitest run tests/stream/remote-print.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.** Pages are checked in `print-wire.ts` before being kept. The PNG bytes are wrapped as a `Blob` immediately, and the decoded RGBA is never held. Settlement removes the abort listener on every path.
- [ ] **Step 4: Run green**, then `bunx vitest run tests/stream` and `bun run lint`.
- [ ] **Step 5: Commit.** `feat(stream): client print jobs with every page checked against the job`

### Task 6: Shared print surface — `printRasters` extracted from `printDocument`

**Files:**
- Modify: `src/components/paged/printDocument.ts`
- Test: the existing print tests (locate with `grep -rln printDocument tests`), plus a new `tests/components/print-rasters.test.ts` (jsdom)

**Interfaces:**
- Produces: `export async function printRasters(pages: readonly { widthPx: number; heightPx: number; png: Blob }[], themeRoot: HTMLElement, opts: { signal: AbortSignal; title?: string; onNotice(text: string): void }): Promise<PrintJob | null>`.
  - It builds the same hidden iframe, `pageRules`, per-page `<img data-page>`, the load waits, `win.print()`, and `afterprint`/abort dispose that `printDocument` uses today.
  - `printDocument` becomes: plan, rasterize each page to a blob, then hand off to the shared core. Its observable behaviour, notices and limits do not change.
  - The `PRINT_JOB_MAX_BLOB_BYTES` check stays in the shared core.

- [ ] **Step 1: Write failing tests** for `printRasters`.
  - Two blobs of different sizes produce an iframe with two `img[data-page]` and the named-page rules for both sizes.
  - `print` is called once on the frame's window (stub it).
  - Abort before the last image disposes the frame and revokes every object URL (spy `URL.revokeObjectURL`).
  - The total blob size over the ceiling gives the notice 'This document is too large to print from the viewer.' and no print.
  - An image `error` gives 'The document could not be printed. Please try again.'
- [ ] **Step 2: Run and watch them fail.** Expected: FAIL (`printRasters` is not exported).
- [ ] **Step 3: Refactor** `printDocument` onto `printRasters` with no behaviour change.
- [ ] **Step 4: Run green.** Run the new test and every existing print test, unmodified, plus `bunx vitest run` for `tests/components` and `bun run lint`.
- [ ] **Step 5: Commit.** `refactor(print): share the iframe print surface between local and streamed documents`

### Task 7: Remote print UI — button, range picker, progress, notices; old path removed

**Files:**
- Create: `src/components/remote/RemotePrint.tsx` (the print button, range dialog, progress line and notices; all elements `data-ov-remote-chrome`)
- Modify: `src/components/remote/RemoteDocView.tsx`:
  - The root's `data-ov-print` is always `"false"`.
  - Delete the `beforeprint` `paintAll` listener and the canvas-print Ctrl/Cmd+P path.
  - Ctrl/Cmd+P opens `RemotePrint`'s flow when allowed, and is always `preventDefault`ed.
- Test: create `tests/stream/remote-print-ui.test.tsx`; update `tests/stream/local-unchanged.test.tsx` only if it pins the removed path (say why in the commit).

**Interfaces:**
- Consumes: `RemoteOfficeDoc.print` (Task 5), `printRasters` (Task 6) via `await import('../paged/printDocument')` inside the print handler, never statically. Also `doc.manifest.caps.print` and `onManifestChange`.
- UI contract (exact copy):
  - Button: label "Print", shown only when `allowPrint && manifest.can.print`.
  - With no cap (`caps.print.maxPages === null && maxSessionPages === null`), Print requests the whole document (`0..count-1`) with no dialog.
  - With any cap, Print opens a dialog titled "Print pages", with "From" and "To" number inputs (1-based, `min=1`, `max=count`). The defaults are 1 and `min(count, maxPages ?? count)`.
    - When `maxPages` is set, the dialog shows "Up to {maxPages} pages per print."
    - When `maxSessionPages` is set, it shows "Up to {maxSessionPages} pages in this session."
    - It has Print and Cancel buttons. From > To disables Print.
  - Progress: "Preparing pages {done} of {total}…", with a Cancel button that aborts the job.
  - Notices:
    - `print-too-large` with a limit: "Too many pages to print at once (limit: {limit} pages)."
    - `print-too-large` without a limit: "This document is too large to print from the viewer."
    - `print-quota-exhausted`: "Print limit for this session reached ({limit} pages)."
    - `rate-limited`: "Too many requests. Try again shortly."
    - Anything else: "The document could not be printed. Please try again."
    - An abort shows no notice.
  - A manifest change on resume re-reads `caps.print`: an open dialog updates its text and bounds, and `can.print` turning false closes it.

- [ ] **Step 1: Write failing tests** (React Testing Library, fake `RemoteOfficeDoc`).
  - The uncapped click calls `doc.print({from: 0, to: count-1})` with no dialog. Then `printRasters` receives the pages (mock the dynamic import).
  - The capped click opens the dialog with defaults `1` and `min(count, maxPages)` and both cap lines, then submits `{from: 0, to: 1}` for inputs 1 and 2.
  - Each notice string appears for its error.
  - Cancel during progress aborts the signal passed to `doc.print`.
  - Ctrl+P with print allowed opens the same flow. Ctrl+P with `can.print=false` calls `preventDefault` and does nothing else, and no button is shown.
  - The root carries `data-ov-print="false"` at all times, including after every unit has painted.
  - No `beforeprint` listener is registered by the view (spy `window.addEventListener`).
  - A manifest change from `caps.print {maxPages: 5}` to `{maxPages: 2}` updates an open dialog's line to "Up to 2 pages per print.".
- [ ] **Step 2: Run and watch them fail.** Run `bunx vitest run tests/stream/remote-print-ui.test.tsx`. Expected: FAIL.
- [ ] **Step 3: Implement.** `RemotePrint` owns its `AbortController`, which is aborted on unmount. The dynamic import is inside the handler, so `gates:core` is unchanged.
- [ ] **Step 4: Run green.** Run `bunx vitest run tests/stream tests/components`, `bun run lint`, and `bun run gates:core`. Expected: tests pass, and the `dist/index.js` closure is not larger than main's 421,066 B; report the number.
- [ ] **Step 5: Commit.** `feat(stream): print a streamed document from server print rasters, with a page-range picker under a cap`

### Task 8: End-to-end, harness switch, docs

**Files:**
- Modify: `server/test/app.ts` (`?print=on|off|cap2|session3`, mapped to `true`, `false`, `{maxPages: 2}` and `{maxPages: 5, maxSessionPages: 3}`)
- Create: `tests/browser/stream-print.spec.ts` (Playwright project `stream`)
- Modify: `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` (§4.2 amendment with rulings P1–P7, dated 2026-11-13)
- Modify: `docs/superpowers/specs/2026-11-07-server-streaming-design.md`:
  - §6.1: the policy type.
  - §6.2 rule 7: print leaves the resume match.
  - §6.2 rule 8 registry: `print-too-large` and `print-quota-exhausted`.
  - §14: the print rate and concurrency as built, and the PRINT_RASTER caps as built.
  - §19 Honest limits: the `unit` op still serves screen rasters, so print caps limit printing, not capture. Print caps reset on re-open. A cancelled or failed job stays charged.
- Modify: `server/README.md` (a "Printing" section next to "Text services")
- Create: `docs/validation/streaming-s3b.md` (gates number; server/vitest/browser counts; print timing for the corpus decks at scale 2, from a small script or the e2e log)

- [ ] **Step 1: Write the browser spec.**
  - Install an init script that stubs `print` on any `[data-officeview-print-frame]` iframe's window as it is added (a MutationObserver), recording calls on `window.__ovPrints`.
  - With `?print=on` on `pypptx-test.pptx`, clicking Print leads to one recorded print, and the frame holds `img[data-page]` equal to the unit count, each `naturalWidth` within `(unit.w - 1) * 2 < naturalWidth ≤ unit.w * 2`.
  - With `?print=cap2`, the dialog appears. From 1 To 3 shows "Too many pages to print at once (limit: 2 pages)." and nothing is printed. From 1 To 2 prints 2 images.
  - With `?print=session3`, print 2 then 2: the second shows "Print limit for this session reached (3 pages)."
  - With `?print=off`, no Print button. Ctrl/Cmd+P opens nothing.
  - `page.emulateMedia({media: 'print'})` on the main page shows the remote root's computed `display` as `none`.
- [ ] **Step 2: Run.** Run `bun run test:browser:stream`. Expected: the new spec passes with the existing stream specs; report counts.
- [ ] **Step 3: Write the docs above.** Every number in them is the built value.
- [ ] **Step 4: Run the whole verification.**
  - `cd server && bun test`
  - `bun run typecheck:server`
  - `bun run typecheck` (or the repo's client typecheck)
  - `bun run lint`
  - `bunx vitest run`
  - `bun run test:browser:stream`
  - `bun run gates:core`
  - Record each result in `docs/validation/streaming-s3b.md`.
- [ ] **Step 5: Commit.** `docs(stream): print — spec amendment, registry, honest limits, validation`
