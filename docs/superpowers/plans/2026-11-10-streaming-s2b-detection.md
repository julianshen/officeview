# Streaming Stage 2b — Tracing a Leaked Capture: Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** An authorized investigator runs one command on a leaked screenshot,
a document and a page number, and gets a trustworthy answer: which session
leaked it, possible collusion, no match, or no usable signal.

**Architecture:** The Stage-0 detector moves into `server/src/mark/`, gains a
register-once/score-many API and a continuous zoom search that abstains on
unsupported geometry. The audit table is backed by a durable `bun:sqlite`
store (rows, per-unit delivery records with reference hashes, per-generation
token counters, a detection log). A keyring supplies key generations. A
`trace()` library assembles candidates, re-renders references with
`DocumentHost`, and makes one §13.2 decision; `server/scripts/detect.ts` is
its CLI. A calibration script over served bytes refits τ and runs the
survival, zoom, collusion and geometry batteries.

**Tech Stack:** Bun, TypeScript, `bun:sqlite` (built in), `@napi-rs/canvas`
(already a server dependency), `bun test`.

**Spec:** `docs/superpowers/specs/2026-11-10-streaming-s2b-detection-design.md`
(parent: `docs/superpowers/specs/2026-11-07-server-streaming-design.md`
§9.3, §13.2–13.5, §16 T9/T10).

## Global Constraints

- `server/` workspace (plus `scripts/stream-s0/mark/*` re-export shims and
  `scripts/stream-s2b/`); no change to the library's eager bundle —
  `bun run gates:core` stays green.
- No new dependencies. Store engine: `bun:sqlite`, WAL, `synchronous=FULL`,
  file mode 0600.
- Codec `l1-bch512x40@1` and `DEFAULT_PARAMS` unchanged.
- Audit rows write-once; deletion only by retention expiry (90 days,
  `AUDIT_RETENTION_SEC`).
- Issuance row durable before the first marked unit; a delivery record
  durable before its frame is sent.
- One decision per capture over every candidate; multi-hit ⇒ possible
  collusion, never a single name. Below threshold ⇒ no guess.
- Refuse above ⌊(2³¹ − 1) / 512⌋ = 4,194,303 candidates (the detector's
  matrix bound; PR #68 review — 2²² overflowed it by one). τ from the next declared
  M ∈ {2¹⁶, 2²⁰, 2²⁴} at or above the candidate count.
- Token counter and `msg40` uniqueness are per key generation.
- Key material never printed or logged.
- Out of scope: Layer 2, page/document identification, photos/rotation/
  perspective (answer "geometry not supported"), session persistence,
  audit export, M-rule, occlusion store.

## Review Focus

1. **Restart between viewing and tracing** (incl. a session detached in its
   grace window when the process dies): the row and deliveries survive and
   the trace names the session. → Task 4 test `restart keeps rows, deliveries, counters`, Task 5 test `traces a session served before a restart`.
2. **Bucket-2 unit captured at 50 % zoom** (same pixel size as bucket 1):
   both scales' candidates are scored in one decision and the right one is
   named. → Task 5 test `bucket-2 capture at half zoom traces`.
3. **Two sessions with identical visible-mark text** (same viewerId) share a
   reference group but stay distinct candidates. → Task 5 test `same viewer twice: both candidates, right one named`.
4. **Audit store write fails** (read-only file, disk full) at HELLO or before
   a unit: marked session refused 4009 / unit refused; nothing marked sent
   without a record. → Task 4 test `store failure refuses, never sends unrecorded`.
5. **JPEG capture file** handed to the CLI (screenshots are often saved as
   JPEG). → Task 6 test `accepts a JPEG capture`.

---

### Task 1: Detector into `server/src/mark`, register-once/score-many

**Files:**
- Create: `server/src/mark/detect.ts`, `server/src/mark/register.ts` (moved
  bodies of `scripts/stream-s0/mark/{detect,register}.ts`)
- Modify: `scripts/stream-s0/mark/detect.ts`, `scripts/stream-s0/mark/register.ts` → one-line `export * from '../../../server/src/mark/…'` (2a pattern)
- Test: `server/test/mark-detect.test.ts`

**Interfaces:**
- Produces (in `server/src/mark/register.ts`):
  - `type Pose = { factor: number; x: number; y: number; scaleX: number; scaleY: number }`
  - `interface Registration extends Pose { reference: Bitmap; error: number }` (unchanged shape)
  - `register(capture, reference, chips?): Registration | undefined` (unchanged)
  - `alignReference(reference: Bitmap, pose: Pose, width: number, height: number): Bitmap` — the same resample+crop `register` applies.
- Produces (in `server/src/mark/detect.ts`): existing `prepareCandidates`, `scoreCapture`, `detect`, `decideScores`, `Decision`, `Scores`, `CandidateMatrix`, plus
  - `scoreRegistered(capture: Bitmap, pose: Pose, reference: Bitmap, chips: Chip[], candidates: CandidateInput, params: MarkParams): Scores` — `scoreCapture` becomes `register` + `scoreRegistered`.

- [ ] **Step 1: Write the failing tests** in `server/test/mark-detect.test.ts`:
  - `alignReference reproduces register's aligned reference`: for a marked
    test page captured with `crop(b, 0.25)` and with `scale(b, 0.5)`,
    `alignReference(ref, reg, cap.width, cap.height).data` equals
    `reg.reference.data` byte for byte.
  - `scoreRegistered equals scoreCapture`: same inputs ⇒ identical `scores`
    arrays and `support`.
  - `pose reuse across visible-mark variants`: register against reference A
    (visible text "a"); `scoreRegistered` with that pose on reference B
    (text "b") and B's capture attributes B's token (index of B).
- [ ] **Step 2: Run** `cd server && bun test test/mark-detect.test.ts` — Expected: FAIL (module not found).
- [ ] **Step 3: Move and split.** Move the files, add `alignReference` and `scoreRegistered`; shims re-export.
- [ ] **Step 4: Run** `cd server && bun test test/mark-detect.test.ts` and the Stage-0 mark suites (`bun test scripts/stream-s0` from repo root, 77/77 and 340/340 as in 2a) — Expected: PASS, counts unchanged.
- [ ] **Step 5: Commit** `refactor(mark): detector in server/src/mark with register-once/score-many`.

### Task 2: Continuous zoom search and geometry abstention

**Files:**
- Modify: `server/src/mark/register.ts`, `server/src/mark/detect.ts`, `scripts/stream-s0/mark/attacks.ts` (add `rotate`, `keystone`)
- Test: `server/test/mark-geometry.test.ts`

**Interfaces:**
- Consumes: Task 1.
- Produces:
  - `register` searches factor continuously over `[ZOOM_MIN, ZOOM_MAX] = [0.45, 2.2]` (coarse multiplicative grid then refinement; algorithm is the builder's), crop ≤ 25 % per axis kept.
  - `export const GEOMETRY_ERROR_MAX: number` in `register.ts` — normalised registration residual above which the capture is out of scope.
  - `Scores` gains `geometry?: 'unsupported'`; `Decision` abstain becomes `{ kind: 'abstain'; reason: 'flat' | 'geometry' }`. No registration, or `error > GEOMETRY_ERROR_MAX` ⇒ `geometry: 'unsupported'` ⇒ `abstain geometry`.
  - `rotate(b: Bitmap, degrees: number): Bitmap`, `keystone(b: Bitmap, frac: number): Bitmap` in `attacks.ts`.

- [ ] **Step 1: Write the failing tests** in `server/test/mark-geometry.test.ts` (one marked corpus page, `pypptx-test.pptx` unit 0, M = 1,024 candidates, Stage-0 τ 7.99/3.85):
  - `registers browser zoom levels`: for f ∈ {0.5, 0.75, 1.1, 1.25, 1.5, 1.75, 2}, `register(scale(marked, f), ref)` has `|factor − f| ≤ 0.005` and the decision is `attributed` to the marking token.
  - `zoom plus jpeg`: `jpeg(scale(marked, 1.25), 70)` attributes correctly.
  - `rotation and perspective abstain`: `rotate(marked, 2)`, `rotate(marked, 5)`, `keystone(marked, 0.05)` ⇒ `{ kind: 'abstain', reason: 'geometry' }`.
  - `registration time`: one zoom-1.25 capture of that page registers in ≤ 15 s (report the measured time in the commit message).
- [ ] **Step 2: Run** `cd server && bun test test/mark-geometry.test.ts` — Expected: FAIL (factors 0.75/1.1/… not found; no `geometry` reason).
- [ ] **Step 3: Implement** the continuous search, `GEOMETRY_ERROR_MAX` (set to 1.5 × the largest in-scope residual measured over this test's in-scope cases; record both numbers in a comment), the abstain path, and the two attacks.
- [ ] **Step 4: Run** the new test, `server/test/mark-*.test.ts`, and the Stage-0 suites — Expected: PASS. A Stage-0 survival case that regresses is a finding to report, not a test to loosen.
- [ ] **Step 5: Commit** `feat(mark): continuous zoom registration; abstain on unsupported geometry`.

### Task 3: Keyring and key generations

**Files:**
- Create: `server/src/keyring.ts`
- Modify: `server/src/keys.ts`, `server/src/config.ts`, `server/src/main.ts`, `server/README.md`
- Test: `server/test/keyring.test.ts`, extend `server/test/config.test.ts` if present (else in keyring test)

**Interfaces:**
- Produces:
  - `interface Keyring { current: number; keys: ReadonlyMap<number, Uint8Array> }`
  - `parseKeyring(json: string): Keyring` — format `{"current": 2, "keys": {"1": "<base64 32 B>", "2": "<base64 32 B>"}}`; generations are integers 1..255; `current` must be present; errors are `ConfigError('bad-key', …)` whose message never contains key bytes.
  - `loadKeyring(env: Record<string, string | undefined>, demo: boolean): { keyring: Keyring; ephemeral: boolean }` — `OV_STREAM_KEYRING=<path>` reads the file; else `OV_STREAM_KEY` ⇒ `{ current: 1, keys: {1} }`; both set ⇒ `ConfigError('bad-arg')`; neither ⇒ demo generates gen 1, else `ConfigError('missing-key')`.
  - `deriveKeys(master: Uint8Array, generation = 1): Promise<Keys>` — `keyGeneration = generation`.
  - `deriveKeyring(ring: Keyring): Promise<Map<number, Keys>>`.
  - `Config.keyring: Keyring` replaces `Config.masterKey`; `main.ts` derives the current generation for the stream server.
- [ ] **Step 1: Write the failing tests** `server/test/keyring.test.ts`:
  - `single key is generation 1`; `keyring file selects current`; `rejects generation 0, 256, non-integer, missing current, wrong length` (each `ConfigError`, message has no base64 of the key); `both env vars is bad-arg`.
  - `sessions record the current generation`: harness started with `keys = await deriveKeys(k2, 2)` ⇒ `stream.audit.get(sid).keyGeneration === 2`, and the served unit decodes only with `deriveBasisKey((await deriveKeys(k2, 2)).kBasisRoot, …)`.
- [ ] **Step 2: Run** `cd server && bun test test/keyring.test.ts` — Expected: FAIL (module not found).
- [ ] **Step 3: Implement**; document `OV_STREAM_KEYRING` and rotation (add a generation, switch `current`, keep old ones until their rows expire) in `server/README.md`.
- [ ] **Step 4: Run** `cd server && bun test` — Expected: all pass.
- [ ] **Step 5: Commit** `feat(server): keyring with key generations (§13.4 rotation)`.

### Task 4: Durable audit store

**Files:**
- Create: `server/src/audit-store.ts`
- Modify: `server/src/audit.ts`, `server/src/ws.ts`, `server/src/config.ts`, `server/src/main.ts`, `server/README.md`
- Test: `server/test/audit-store.test.ts`; update `server/test/sessions.test.ts`, `server/test/audit-issue.test.ts`, `server/test/forensic-serve.test.ts`

**Interfaces:**
- Consumes: Task 3 (`keyGeneration` per row).
- Produces (in `audit-store.ts`):
  - `interface StoredRow extends AuditRowInput { docRef: string; contentPolicy: 'visible-only' | 'full'; visibleMark: MarkSpec }` (`MarkSpec` from `render.ts`, viewer text included).
  - `interface DeliveryRecord { sid: string; unitIndex: number; scale: number; status: 'l1' | 'unmarked'; reason?: string; refSha256: string; at: number }` — unique on `(sid, unitIndex, scale, refSha256, status)`.
  - `interface DetectionLog { at: number; osUser: string; investigator?: string; imageSha256: string; docRefs: string[]; unitIndex: number; answer: string; sid?: string }`
  - `class AuditStore` with `static open(path: string): AuditStore` (`':memory:'` allowed), `close()`, and:
    `putRow(sid, row: StoredRow)` (throws on duplicate sid), `getRow(sid)`, `putDelivery(d)`, `deliveries(sid)`,
    `candidates(docRef: string, unitIndex: number): Array<{ sid: string; row: StoredRow; delivery: DeliveryRecord }>` (all statuses),
    `takeToken(gen: number): number` (durably increments and returns the previous value; never rewinds),
    `liveMessages(gen: number): string[]`, `putDetection(r: DetectionLog)`, `detections(): DetectionLog[]`,
    `expire(nowSec: number): number` (deletes rows, deliveries and their messages older than `AUDIT_RETENTION_SEC`; counters stay).
  - `AuditStore.openReadOnlyForTrace(path)` — rows/deliveries read-only; only `putDetection` writes.
- Produces (in `audit.ts`): `new AuditTable({ store?: AuditStore })` (default `AuditStore.open(':memory:')`); `issue(kMsg, gen, draw?)` uses `store.takeToken(gen)` and the gen's live messages; `record(sid, row: StoredRow)`; `recordDelivery(d: Omit<DeliveryRecord, 'at'>)` replaces `noteDelivered`; `get(sid)` derives `forensicLayers`/`delivered`/`untraceable` from deliveries as today; `nextToken(gen)`.
- ws.ts: `record` failure ⇒ close 4009 `policy-refused` (before the manifest); `recordDelivery` runs **before** `sendFrames` with `refSha256 = sha256(canonical PNG)` and the unmarked reason; failure ⇒ refuse the unit `policy-refused`.
- config: `OV_AUDIT_DB=<path>` required outside `--demo` (`ConfigError('missing-audit')`, add the code); demo uses `':memory:'`. main.ts runs `audit.expire(now)` at startup and hourly.

- [ ] **Step 1: Write the failing tests** `server/test/audit-store.test.ts`:
  - `restart keeps rows, deliveries, counters`: file store in the scratch/tmp dir; issue 3 tokens on gen 1, record 2 rows, 2 deliveries; close; reopen ⇒ rows equal, deliveries equal, `nextToken(1) === 3`, `nextToken(2) === 0`, and an `insertMessage` draw that collides with a reloaded `msg40` re-draws.
  - `write-once survives reopen`: `record` of an existing sid after reopen throws.
  - `expiry deletes rows, deliveries and messages but not counters`.
  - `store failure refuses, never sends unrecorded`: harness with a store whose `putRow` throws ⇒ HELLO closes 4009 and no manifest; a store whose `putDelivery` throws ⇒ the unit request gets `refused policy-refused` and no `UNIT_RASTER` frame arrives.
  - `delivery recorded before send with reference hash`: after a served marked unit, `deliveries(sid)` has `{ unitIndex, scale, status: 'l1', refSha256: sha256(stream.canonicalOf(sid, unitIndex, scale).png) }`; a capacity-unmarked unit records `status 'unmarked'` with its reason (2a deferred minor).
  - `trace store is read-only`: `openReadOnlyForTrace` rejects `putRow`/`putDelivery`, accepts `putDetection`.
  - `file mode is 0600`.
- [ ] **Step 2: Run** `cd server && bun test test/audit-store.test.ts` — Expected: FAIL (module not found).
- [ ] **Step 3: Implement** store, table facade, ws ordering, config, expiry timer; update the three existing tests to the new row/delivery API (their assertions on `get()` stay).
- [ ] **Step 4: Run** `cd server && bun test` — Expected: all pass.
- [ ] **Step 5: Commit** `feat(server): durable audit store; delivery records before send`.

### Task 5: `trace()` — one decision over every candidate

**Files:**
- Create: `server/src/detect/trace.ts`, `server/src/detect/tau.ts`, `server/src/detect/image.ts`
- Test: `server/test/trace.test.ts`

**Interfaces:**
- Consumes: Tasks 1–4.
- Produces:
  - `tau.ts`: `export const TAU_TABLE: ReadonlyArray<{ m: number; abs: number; margin: number; source: string }>` for m ∈ {2¹⁶, 2²⁰, 2²⁴}; provisional values from Stage 0 — 2¹⁶: 7.99/3.85; 2²⁰: √(7.99² + 2 ln 16)/3.85; 2²⁴: 8.66/3.85 — `source: 'stage-0 provisional'` until Task 7 rewrites them. `tauFor(count: number): { abs; margin; m }`; `MAX_CANDIDATES = Math.floor(CANDIDATE_MATRIX_MAX_BYTES / 512)` (4,194,303; was 2 ** 22, which overflowed `prepareCandidates`' bound — PR #68 review).
  - `image.ts`: `decodeImage(bytes: Uint8Array): Promise<Bitmap>` (PNG and JPEG via `@napi-rs/canvas` `loadImage`).
  - `trace.ts`:
    ```ts
    export type TraceAnswer =
      | { kind: 'traced'; sid: string; viewerId?: string; ip: string; startedAt: number; keyGeneration: number; unitIndex: number; scale: number; z: number; margin: number }
      | { kind: 'collusion'; candidates: Array<{ sid: string; viewerId?: string; z: number }> }
      | { kind: 'no-match'; zBest: number }
      | { kind: 'no-signal'; reason: 'flat' | 'geometry' | 'never-marked' | 'no-candidates' | 'reference-mismatch' }
      | { kind: 'refused'; reason: 'too-many-candidates'; count: number }
    export interface TraceInput { capture: Bitmap; doc: Uint8Array; unitIndex: number; keys: ReadonlyMap<number, Keys>; store: AuditStore; host?: DocumentHost }
    export function trace(input: TraceInput): Promise<TraceAnswer>
    ```
  - Algorithm (pinned): per generation, `docRef = docRefOf(keys.kDoc, doc)`; candidates = `store.candidates(docRef, unitIndex)` with `status 'l1'`; none but some `unmarked` ⇒ `never-marked`; none at all ⇒ `no-candidates`; count > `MAX_CANDIDATES` ⇒ `refused`. Group by `(contentPolicy, visibleMark, scale)`; render each group's reference once with `DocumentHost.render(..., { projection: 'final-raster', mark })`; any group whose `sha256(png)` differs from its deliveries' `refSha256` ⇒ `reference-mismatch` (fail closed). Register once per `(contentPolicy, scale)` against one of its references; `scoreRegistered` each group with chips from `basis(deriveBasisKey(keys[gen].kBasisRoot, docRef_gen, \`${unitIndex}@${scale}\`), w, h, DEFAULT_PARAMS)` and codes `extendedCodeword(msg40)`. Concatenate all scores; `decideScores` with `tauFor(total)`; map indices back to sids. All registrations failed/out of range ⇒ `geometry`; all flat ⇒ `flat`.
- [ ] **Step 1: Write the failing tests** `server/test/trace.test.ts` (harness + file store; sessions fetch units over the wire; captures are the received `UNIT_RASTER` PNGs):
  - `traces the leaking session` (A, B viewing unit 0; A's capture ⇒ `traced` A with `keyGeneration`, `ip`).
  - `same viewer twice: both candidates, right one named` (two sessions, same viewerId).
  - `bucket-2 capture at half zoom traces` (`scale(capB2, 0.5)` with bucket-1 and bucket-2 candidates present).
  - `zoom 1.25 + jpeg q70 traces`.
  - `average of two is never a single innocent` (`average([A, B])` ⇒ `collusion` or `no-match` or `traced` ∈ {A, B}; never a third session C).
  - `unmarked canonical is not traced` (⇒ `no-match` or `no-signal`).
  - `never-marked page` (capacity-unmarked unit) ⇒ `no-signal never-marked`; `unrequested page` ⇒ `no-candidates`.
  - `reference mismatch fails closed` (tamper one delivery's `refSha256` in the store) ⇒ `no-signal reference-mismatch`.
  - `rotated capture` ⇒ `no-signal geometry`.
  - `traces across key generations` (session under gen 1, server restarted with ring current 2, second session; each capture traces to its own session).
  - `traces a session served before a restart` (store closed and reopened between serving and tracing).
- [ ] **Step 2: Run** `cd server && bun test test/trace.test.ts` — Expected: FAIL (module not found).
- [ ] **Step 3: Implement** `tau.ts`, `image.ts`, `trace.ts`.
- [ ] **Step 4: Run** `cd server && bun test` — Expected: all pass.
- [ ] **Step 5: Commit** `feat(server): trace() — one §13.2 decision over all candidates`.

### Task 6: `detect` CLI

**Files:**
- Create: `server/scripts/detect.ts`
- Modify: `server/README.md` (investigator section)
- Test: `server/test/detect-cli.test.ts`

**Interfaces:**
- Consumes: Task 5 `trace`, Task 4 `AuditStore.openReadOnlyForTrace`, Task 3 `loadKeyring`.
- Produces: `bun server/scripts/detect.ts --image <file> [--image <file> …] --doc <file> --page <N> --audit-db <path> [--keyring <path>] [--investigator <label>] [--json]`. `--page` is **1-based** (unitIndex = N − 1). Keys: `--keyring` file, else `OV_STREAM_KEYRING`/`OV_STREAM_KEY` env as the server reads them.
  - One answer per image, one line each, exact prefixes:
    `<image>: TRACED — session <sid>, viewer <id|unknown>, ip <ip>, started <ISO-8601 UTC>, key generation <g>, page <N> (z <z.2f>, margin <m.2f>)`
    `<image>: POSSIBLE COLLUSION — candidates <sid> (z …), …; no single viewer named`
    `<image>: NO MATCH — mark below threshold (best z <z.2f>); nothing attributed`
    `<image>: NO USABLE SIGNAL — <reason>` with reasons: flat `image carries no usable mark signal`, geometry `image geometry not supported (rotated, skewed or photographed)`, never-marked `this page was never delivered marked`, no-candidates `no audit record of this page being viewed`, reference-mismatch `re-rendered page differs from what was served (renderer or font drift)`
    `<image>: REFUSED — <count> candidates exceed the detector limit (4194303)`
  - `--json`: one JSON object per line `{ image, answer: TraceAnswer }`.
  - Exit codes: 0 every image answered; 2 usage/config error (message, no key bytes); 1 internal error.
  - Each image appends a `DetectionLog` (`osUser` = `os.userInfo().username`).
- [ ] **Step 1: Write the failing tests** `server/test/detect-cli.test.ts` (fixtures produced in-test via the harness, CLI run with `Bun.spawn`):
  - `prints TRACED for the leaking session` and exit 0; `--json` parses to `kind 'traced'`.
  - `accepts a JPEG capture` (capture saved as JPEG q90).
  - `two images, two answers, two log rows`.
  - `missing key is exit 2 without key material`; `bad --page is exit 2`.
  - `output and log never contain key material` (base64 and hex of the master key absent from stdout, stderr and `detections()`).
  - `does not modify audit rows` (row/delivery counts and contents equal before/after).
- [ ] **Step 2: Run** `cd server && bun test test/detect-cli.test.ts` — Expected: FAIL.
- [ ] **Step 3: Implement** the CLI and README section.
- [ ] **Step 4: Run** `cd server && bun test` — Expected: all pass.
- [ ] **Step 5: Commit** `feat(server): detect CLI for investigators`.

### Task 7: Calibration and batteries on served output (Stage-2 gate)

**Files:**
- Create: `scripts/stream-s2b/calibrate.ts`, `docs/validation/streaming-s2b.md`, `docs/validation/streaming-s2b.json`
- Modify: `server/src/detect/tau.ts` (values + `source`), `server/README.md` (link)

**Interfaces:**
- Consumes: Tasks 1–6; `calibrate()`/`tailFit` from `scripts/stream-s0/t7-t9-battery.ts` (import, do not copy).
- Produces: `bun --preload ./scripts/stream-s0/napi-preload.ts scripts/stream-s2b/calibrate.ts [--write docs/validation/streaming-s2b.json] [--nulls N]`; exit 1 on any RED gate.
  - Captures are **served bytes**: an in-process stream server (`createStreamServer` + harness `Client`) over the corpus at bucket 1 (pptx also bucket 2), one marking session per unit; a 65,536-row candidate population per (doc, page) is inserted into the store (issued rows + `l1` deliveries with the true `refSha256`).
  - Null decisions ≥ 2,000 (default 2,000; more if runtime allows), kinds unmarked / wrong-key / out-of-table as Stage 0, each through `trace()`; τ fit at 2¹⁶, extrapolated to 2²⁰ and 2²⁴ by the Stage-0 union rule; write `TAU_TABLE` with `source: 'stage-2b calibrate <git sha>'`.
  - Gates (re-run with the new τ): §13.5 Layer-1 survival default ≥ 90 %, q70 ≥ 75 %, down 0.5 ≥ 75 %, q70→down 0.5 ≥ 75 %, crop 25 % ≥ 70 %; null false attributions 0 and fitted tail ≤ 10⁻⁶; collusion K-average {2,4,6}, median {3,4,5,6}, differencing: **innocent named 0**; geometry battery (rotate 2°, 5°, keystone 5 %): **0 attributions**.
  - Published, not gated: zoom battery {0.75, 1.1, 1.25, 1.5, 1.75} (alone and with q70), colluder-named / suspected / none rates, runtime per trace at 65,536 and 10,000 candidates (spec target ≤ 5 min for 10,000).
  - `docs/validation/streaming-s2b.md` states plainly that 10⁻⁶ rests on the fitted tail and union bound, not observation.
- [ ] **Step 1: Write** the script with `--limit` for a smoke run; run `--limit 3` — Expected: completes, prints a gate table.
- [ ] **Step 2: Run** the full calibration with `--write` — Expected: all gates GREEN. Any RED gate is reported to the owner as-is; do not change parameters, targets or attacks to make it pass.
- [ ] **Step 3: Write** `docs/validation/streaming-s2b.md` from the JSON (exit table, τ table, survival, zoom, collusion, geometry, runtime).
- [ ] **Step 4: Run** `cd server && bun test` (trace tests now use the new τ) and `bun run gates:core` — Expected: all pass.
- [ ] **Step 5: Commit** `test(stream): Stage 2b calibration on served output; τ refit; collusion/geometry gates`.
