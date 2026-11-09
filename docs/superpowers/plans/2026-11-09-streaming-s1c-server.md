# Streaming Stage 1c — the demo stream server — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A WS-only `server/` Bun workspace that opens a document for a viewer through `authorize` + `resolveDocument`, strips hidden content before layout (fail-closed), and streams default-protected `final-raster` units (visible mark composited) — or, for an audited relaxed session, the 1b command units with their assets — under the §6.2 handshake, token contract, close-code registry and §14 caps.

**Architecture:** Bun main process = `Bun.serve` WebSocket endpoint + session/token/audit tables + caches + admission; a pool of Bun `Worker`s (napi preload, as T12) parses and paints. The content policy is a **pre-parse package projection**: per `contentPolicyId` the server rewrites the OOXML parts (removing hidden elements) into a new in-memory package and parses *that*, so the original bytes and every other policy's parse are untouched and **no library parser or painter changes** (local mode stays byte-for-byte unchanged by construction; the main-entry no-growth gate is unaffected). Wire codecs for the two new payload kinds live in `src/stream/` because the 1d client decodes them.

**Tech Stack:** Bun (workspace, `Bun.serve` websocket, `Worker`, `bun test`), `@napi-rs/canvas`, the 1b `src/stream/*` modules, Web Crypto (HMAC/HKDF/SHA-256).

**Spec:** `docs/superpowers/specs/2026-11-07-server-streaming-design.md` (draft 12.1; §6.1–6.4, §7.1, §7.4–7.7, §9, §10, §14, §16 T2/T4/T6/T10, §17 row 1).

**Branch:** continue on `feature/streaming-s1` (§17: Stage 1 is one PR); one commit per task.

## Global Constraints

- Endpoint: WS only. **No file-download route, no HTTP document endpoint**; any non-upgrade HTTP request → 404 (§6.1, §2 item 4).
- Server imports `../src/**` TS directly; **never** `src/components/OfficeFile.tsx`, `react`, `react-dom`, or `document` (purity gate, Task 1). Parsing goes through `detectAndParse(OfficePackage)` (`src/components/loadOfficeSource.ts:56`, React-free) and painting through `getPaintables` (`src/render/paint.ts`).
- Transport: `Bun.serve({ websocket: { maxPayloadLength: 1_048_576 } })` — the §7.1/§14 native receiver cap (Bun's equivalent of `ws` `maxPayload`); every server-emitted application fragment ≤ 1 MiB, one fragment per WS message.
- HELLO ≤ 4 KiB; any text control message ≤ 64 KiB; MANIFEST encoded ≤ 256 KiB and decoded ≤ 256 KiB; per-kind caps exactly the §14 table.
- `sid` = 48 random bits (hex, 12 chars); token = HS256 JWT under `kToken`, claims exactly `{ v: 1, docRef, sid, exp, iat, jti, scope: 'stream' }`, ≤ 512 chars; `docRef` = hex of HMAC-SHA256(kDoc, sha256(bytes)) truncated to 96 bits.
- Keys: one 32-byte master `OV_STREAM_KEY` (base64); `kDoc`, `kToken`, `kAudit` = HKDF-SHA256(master, salt `"officeview-stream"`, info `"kDoc"|"kToken"|"kAudit"`). `--demo` without a key generates an ephemeral one and logs `EPHEMERAL KEY — sessions die with this process`; non-demo without a key refuses to start (exit 2).
- Windows: ready window 2 s from MANIFEST; RENEW sent ≥ 5 s before `exp`; reassembly 5 s. Token lifetime 300 s [call].
- Close codes verbatim (§6.2 rule 8): 4000 `unsupported-version`, 4001 `unsupported-format`, 4002 `ready-timeout`, 4003 `auth-failed`, 4004 `session-expired`, 4005 `revoked`, 4006 `layout-changed`, 4007 `oversized-frame` (client only — never sent by the server), 4008 `rate-limited` (third breach in 60 s), 4009 `policy-refused` (client relax attempt), 4409 `superseded`. Op refusals `{ reqId, type: 'refused', code }` with code ∈ {`rate-limited`, `policy-refused`, `session-rotated`}. **No other code exists.**
- Version = set intersection of `recVersions` with `{1}`; empty ⇒ 4000.
- Defaults: projection `final-raster`, content policy `visible-only`, visible mark `fixed`, print `true`. **Forensic mark: Stage 2.** Stage 1 reports `forensicMark: false` in the effective policy and audit row, and the manifest carries the notice `forensic-mark-stage-2` (never claim a mark that is not applied).
- Rates (§14): per IP 30 sessions/min, 8 concurrent connects; per sid 120 msgs/10 s (burst 20/s), 60 unit requests/min; process 256 sockets, 4,096 queued units. Pool: 1 worker per 2 vCPU (cap 8), per-session outstanding 8 / queue 32, 2 s render deadline with drain-then-refund, recycle at ≥ 256 pages or ≥ 64 MiB registered embedded fonts. Retained ceilings: document cache 1 GiB, render cache 512 MiB per process; session transient 96 MiB.
- Determinism (§14): no `Date.now()`/`Math.random()` in the render path; the visible-mark text is an injected parameter.
- Scale buckets: `[1, 2]`; a unit request at any other scale is `policy-refused`.

## Review Focus

1. **READY twice / ops before READY** — a second READY is `policy-refused` (session stays open); any STREAM op before a verified READY is `policy-refused` and nothing renders. → Task 8 tests.
2. **Out-of-range `unitId`, non-bucket scale, NaN/negative fields in a unit request** — refused `policy-refused`, no pool job created. → Task 8 tests.
3. **Client→server binary frame** (no client binary kinds exist in Stage 1) — dropped before allocation, counted toward the rate gate; never parsed as OVSD. → Task 8 test.
4. **Text message under `maxPayloadLength` but over 64 KiB (or HELLO over 4 KiB)** — closed `unsupported-version` (the §6.2 transport-overage code) with no JSON parse. → Task 8 test.
5. **`renew-ack` carrying the old jti, then an op on the old jti** — the ack is refused; after a valid ack, an op pinned to the old jti gets `session-rotated` and the session stays open. → Task 8 test.

---

### Task 1: `server/` workspace, config/keys, purity gate, napi-in-Worker probe

**Files:**
- Create: `server/package.json` (`"name": "@officeview/stream-server"`, `"private": true`, deps `@napi-rs/canvas` same version as root devDeps), `server/tsconfig.json` (extends root, `include: ["src", "test"]`), `server/src/config.ts`, `server/src/keys.ts`, `server/test/config.test.ts`, `server/test/purity.test.ts`, `server/test/worker-probe.test.ts`, `server/src/worker/probe.ts`
- Modify: root `package.json` (`"workspaces": ["server"]`, script `"test:server": "cd server && bun test"`), `vitest.config.ts` exclude `server/**`

**Interfaces:**
- Produces: `loadConfig(env, argv): ServerConfig` with `{ port, demo, docDir, originAllowlist: string[], masterKey: Uint8Array }`; `deriveKeys(master): Promise<{ kDoc: CryptoKey; kToken: CryptoKey; kAudit: Uint8Array }>`; `docRefOf(kDoc, bytes): Promise<string>` (24 hex chars).

- [ ] **Step 1: Write failing tests.** `config.test.ts`: non-demo without `OV_STREAM_KEY` throws `ConfigError('missing-key')`; a key that is not 32 bytes after base64 throws `ConfigError('bad-key')`; demo without key yields a 32-byte key and `ephemeral: true`; `docRefOf` is 24 hex chars, stable for equal bytes, different under a different master. `purity.test.ts`: the static-import closure of `server/src/main.ts` (use `staticClosure` from `tests/helpers/core-gates.ts` over source TS — add a TS-source mode if it only walks dist) contains no `react`, `react-dom`, `src/components/OfficeFile.tsx`, `src/components/OfficeDoc.tsx`, and `purityViolations` reports no `document` reference. `worker-probe.test.ts`: a Bun `Worker` started with the napi preload (`scripts/stream-s0/napi-preload.ts`) paints a 4×4 red rect and posts back RGBA `[255,0,0,255]` ×16, and `@napi-rs/canvas` reports its version.
- [ ] **Step 2:** `cd server && bun test` → FAIL (modules missing).
- [ ] **Step 3:** Implement. HKDF via `crypto.subtle.deriveBits`. `main.ts` is a stub that imports config only (grows in later tasks; the purity test keeps guarding it).
- [ ] **Step 4:** `bun install && cd server && bun test` → PASS; `bun run test` (root) unchanged count; `npx tsc --noEmit` clean.
- [ ] **Step 5:** Commit `feat(server): workspace, keys, purity gate, napi worker probe`.

### Task 2: Admission + pre-parse content projection (§6.4) with the fail-closed rule

**Files:**
- Create: `server/src/policy.ts`, `server/src/policy/xlsx.ts`, `server/src/policy/docx.ts`, `server/src/policy/pptx.ts`, `server/src/policy/visibility-scan.ts`, `server/test/policy.test.ts`, `server/test/t2-stream.test.ts`, `server/test/policy-identity.test.ts`
- Read (fixtures): `scripts/stream-s0/hidden-fixtures.ts` (13 fixtures, `withHidden`/`withoutHidden`/`marker`/`control`)

**Interfaces:**
- Produces: `admit(bytes): { ok: true; format: 'docx'|'xlsx'|'pptx' } | { ok: false; code: 'unsupported-format' }` (ODF `mimetype`, RTF magic, non-zip, or a zip whose main part is not OOXML ⇒ refused; inspected on the original bytes); `projectPackage(pkg: OfficePackage, format, contentPolicy: 'visible-only'|'full'): Promise<{ pkg: OfficePackage; contentPolicyId: string; diagnostics: PolicyDiagnostic[] }>` where `PolicyDiagnostic = { code: 'hidden-stripped' | 'hidden-unrecognised'; part: string; element: string; attr?: string; value?: string }`. `contentPolicyId` = `'visible-only@1'` | `'full@1'`. `'full'` returns the input package unchanged (identity), with no diagnostics.

**Strip list (visible-only), as package rewrites — nothing is removed in a way that shifts the geometry of visible content:**
- XLSX: `<sheet state="hidden|veryHidden">` → remove the `<sheet>` from `xl/workbook.xml` (and its `definedName`s with that `localSheetId`, renumbering the rest); `<row hidden="1|true">` → keep the element and its `ht`/`hidden`, drop its `<c>` children; `<col hidden>` ranges → drop every `<c>` whose column falls in the range (geometry keeps the collapsed extents); comments parts (`xl/comments*.xml`, threaded comments, their rels) removed; `xdr:cNvPr hidden` shapes removed (anchor element whole, propagated through `xdr:grpSp`); `pivotCacheRecords*.xml` parts always removed (raw source rows, never drawn); pivot cache definitions, slicer caches and timeline caches whose source range lies on a removed sheet or covers hidden rows/cols have their item lists emptied (§6.4 pivot row; interactive pivots stay refused until 4c). The identity test below proves the pivot rendering is unchanged where nothing is hidden.
- DOCX: runs (`w:r`) whose effective run properties have `w:vanish` or `w:specVanish` or `w:webHidden` true — direct `rPr`, the `rStyle` chain, the paragraph style's `rPr`, and `w:docDefaults` — removed; `wp:docPr hidden` drawings removed (the `w:drawing` and its run; group children with `hidden` in `wpg:cNvPr`/`pic:cNvPr` removed); comments parts and `commentRangeStart/End`/`commentReference` removed.
- PPTX: `<p:sldId>` whose slide has `show="0"` removed from `ppt/presentation.xml` `sldIdLst`; `p:cNvPr hidden` shapes removed at their `p:sp|p:pic|p:grpSp|p:graphicFrame|p:cxnSp` element (propagated through groups); notes slides and comments parts removed (notes return in Stage 3 as presenter images).
- **Fail-closed scan (`visibility-scan.ts`, §2 item 14a):** after the rewrites, walk every XML part; any attribute named `hidden`, `vanish`, `state`, `show`, `visibility`, `visible` (any namespace), or any element named `vanish`/`specVanish`/`webHidden`, whose value is not an explicit "visible" value (`0|false|visible|show` as applicable) and that no rewrite above handled, removes the **nearest drawable ancestor** (run, shape, cell, row) and records `hidden-unrecognised`. Unparseable parts fail the whole projection (`unsupported-format`).

- [ ] **Step 1: Write failing tests.** `t2-stream.test.ts`: for each of the 13 fixture ids (listed in `hidden-fixtures.ts`) — project `withHidden` as `visible-only`, parse, `getPaintables`, paint every unit through a text-capturing proxy (as `scripts/stream-s0/t2-hidden.ts`) — assert the control string is drawn, the marker appears in **no** `fillText`/`strokeText` string nor their concatenation, and the JSON walk of the parsed projected doc contains no marker (the Stage-0 `modelExposure` walk must be 0 for all 13). `policy.test.ts`: `admit` refuses an ODT, an RTF, a plain zip and random bytes with `unsupported-format`; a synthetic `<p:cNvPr hidden="maybe">` and an `a14:hiddenFill`-style unknown `visibility="hidden"` attribute each strip their shape and emit `hidden-unrecognised`; `full` is identity (byte-equal parts); the input `OfficePackage` is not mutated (compare part bytes before/after). `policy-identity.test.ts` (napi, slow tag): for every corpus file with **zero** `hidden-stripped`/`hidden-unrecognised` diagnostics, projected-parse paint == original-parse paint, RGBA byte-identical at scale 1 (the projection is harmless where nothing is hidden); list the files with diagnostics in the test output.
- [ ] **Step 2:** `cd server && bun test test/policy.test.ts test/t2-stream.test.ts` → FAIL.
- [ ] **Step 3:** Implement using `src/core/xml.ts` `parseXml` + a minimal serializer in `policy/` (round-trip must preserve namespaces and unknown content; reuse an existing serializer if `src/core/xml.ts` has one). Effective-vanish resolution for DOCX reads `word/styles.xml` once.
- [ ] **Step 4:** tests PASS; run `policy-identity.test.ts` under napi and record its file list in `docs/validation/streaming-s1.md` §"Content projection".
- [ ] **Step 5:** Commit `feat(server): §6.4 pre-parse content projection, fail-closed scan, T2 stream path`.

### Task 3: Tokens, session table, audit table, `authorize` + `resolveDocument` hooks

**Files:**
- Create: `server/src/tokens.ts`, `server/src/sessionTable.ts`, `server/src/audit.ts`, `server/src/hooks.ts`, `server/test/tokens.test.ts`, `server/test/sessions.test.ts`

**Interfaces:**
- Produces:
  - `issueToken(kToken, { docRef, sid, ttlSec, now }): Promise<{ token: string; jti: string; exp: number }>`; `verifyToken(kToken, token, { now, docRef, sid }): Promise<{ ok: true; claims } | { ok: false; reason: 'signature'|'exp'|'docRef'|'sid'|'scope'|'malformed' }>`. `now` is injected (seconds).
  - `SessionTable`: `create(row): Session`, `get(sid)`, `rotate(sid): { oldJti, newJti }`, `acceptJti(sid, jti): 'current'|'rotating'|'rotated'|'unknown'`, `ack(sid, jti)`, `supersede(sid, ws)` (closes the prior socket with 4409 atomically), `close(sid)` (evicts). Row = `{ sid, docRef, viewerId?, role, effectivePolicy, activeSocket, jti, pendingJti?, startedAt, status }`.
  - `AuditTable`: write-once `record(sid, { viewerId?, ip, startedAt, keyGeneration, policyId, markSettings, untraceable })`; a second write for the same sid throws; `expire(now)` deletes rows older than 90 days.
  - `type Authorize = (req: { headers: Headers; ip: string; url: URL; token?: string; resume?: { sid: string } }) => Promise<AuthDecision | null>` with `AuthDecision = { viewerId?, role?: 'presenter'|'audience', policy: { contentPolicy?, projection?: 'protected'|'relaxed', visibleMark?: 'fixed'|'viewer'|'off', forensicMark?: boolean, print?: boolean } }`; `type ResolveDocument = (req, decision) => Promise<Uint8Array | null>`; `effectivePolicy(decision): EffectivePolicy` applying the Global-Constraints defaults and forcing `forensicMark: false` (Stage 1).

- [ ] **Step 1: Failing tests.** Round-trip; each `reason` produced by a targeted mutation (flipped signature byte, `now > exp`, other docRef, other sid, scope `x`, 3-part garbage); token length ≤ 512 for a 24-hex docRef; rotation: during rotation both jtis `current`/`rotating`, after `ack(new)` old ⇒ `rotated`, unknown ⇒ `unknown`; `supersede` calls the old socket's `close(4409, 'superseded')` exactly once; audit write-once throws on second write; `effectivePolicy({policy:{}})` = protected/visible-only/fixed/print true/forensicMark false; both marks off ⇒ `untraceable: true` in the audit row.
- [ ] **Step 2:** FAIL. **Step 3:** implement (Web Crypto HMAC; base64url). **Step 4:** PASS.
- [ ] **Step 5:** Commit `feat(server): token contract, session/audit tables, authorize/resolveDocument hooks`.

### Task 4: Pinned font pack (§7.4 rule 6) and `requiredFaces`

**Files:**
- Create: `fontpack/` at repo root — move the five Regular faces + licences from `scripts/stream-s0/fontpack/` (update `t4-fonts.ts` to read the new path) and add Bold, Italic, BoldItalic for Carlito, Caladea, Liberation Sans/Serif/Mono (20 faces total), `fontpack/pack.json`; `server/src/fonts.ts`; `server/test/fonts.test.ts`; `scripts/stream-s1/fontpack-lock.ts` (verifies `pack.json` hashes against bytes)
- Modify: root `package.json` `files` — **exclude** `fontpack/` for now (npm size decision belongs to 1d; flagged in handoff)

**Interfaces:**
- Produces: `pack.json` = `{ version: 1, faces: [{ id: 'carlito-regular', file, family: 'Carlito', weight: 400|700, style: 'normal'|'italic', hash96 }], aliases: { Calibri: 'Carlito', Cambria: 'Caladea', Arial: 'Liberation Sans', Helvetica: 'Liberation Sans', 'Times New Roman': 'Liberation Serif', 'Courier New': 'Liberation Mono' } }` (`hash96` = first 24 hex of SHA-256 of the file bytes); `registerPack(globalFonts)` registers every face under its own family **and** each alias family name (as T4 did); `facesUsed(ops: Uint8Array): string[]` — the face ids whose (family incl. alias, weight, style) match any `font` set in a STATE_SETUP before a GLYPH_RUN; `requiredFacesOf(units): Array<{ id, hash, alias? }>`; `unitNeedsRaster(facesUsed, missingFaces): boolean`. A GLYPH_RUN under a family not in the pack ⇒ that unit is not command-eligible (`face-missing` diagnostic) regardless of READY.
- **Executed note (T4):** alias registration alone is not enough — an installed system face of the alias name (Arial, Times New Roman on macOS) wins over `registerFromPath(path, alias)`. Pinning therefore rewrites the name: `pinFontString` / `pinFonts(ctx)` turn `12pt Calibri` into `12pt "Carlito"`. Task 5 applies `pinFonts` to the top-level paint context **outside** the recorder (so recorded fonts carry pack names) and to every context the server's `CanvasBackend` hands out (text measurement, scratch surfaces). T4 re-run on the s1 corpus: 59/259 units raster with the pack (22.8 %).

- [ ] **Step 1: Failing tests.** `pack.json` hashes match bytes (lock script exits 0; flip a byte in a temp copy ⇒ exit 1); 20 faces, 6 aliases; `facesUsed` on a recorded unit painting `bold 12pt Calibri` + `10pt "Courier New"` = `['carlito-bold', 'liberation-mono-regular']`; a unit using `Aptos` is not command-eligible; `unitNeedsRaster(['carlito-bold'], ['carlito-bold'])` true.
- [ ] **Step 2:** FAIL. **Step 3:** fetch the 15 new faces from the same upstream releases as the Regulars (record URL + SHA-256 in `pack.json`'s `source`), implement. **Step 4:** PASS; re-run `bun run s0:napi scripts/stream-s0/t4-fonts.ts` and confirm the raster rate is ≤ the recorded 28.8 %.
- [ ] **Step 5:** Commit `feat(server): pinned OFL font pack (20 faces), requiredFaces`.

### Task 5: Render pool (§9.2) with the occlusion seam

**Files:**
- Create: `server/src/pool.ts`, `server/src/admission.ts` (lifted from `scripts/stream-s0/t12-admission.ts`'s model — keep T12 importing it from the new location), `server/src/worker/render-worker.ts`, `server/src/render.ts`, `server/test/pool.test.ts`, `server/test/render.test.ts`

**Interfaces:**
- Consumes: Task 2 `projectPackage`, Task 4 `registerPack`/`facesUsed`, 1b `createRecorder`, `paintWatermark` (`src/core/watermark.ts:351`).
- Produces:
  - Worker job `{ jobId, docRef, bytes?: Uint8Array (first job for a docRef only), contentPolicy, unitIndex, scale, projection: 'final-raster'|'commands', mark: { kind: 'fixed'|'viewer'|'off'; text?: string } }` → reply `{ jobId, ok: true, unit: { kind: 'raster'; w; h; png: Uint8Array } | { kind: 'commands'; ops; opCount; assets: Asset[]; faces: string[] } | { kind: 'tile'; reason } , manifestUnits?: Array<{ w, h }>, layoutFingerprint, diagnostics }` or `{ jobId, ok: false, error }`.
  - `Pool.create({ workers, caps, now }): Pool`; `pool.render(sessionId, job, signal): Promise<Reply>` (admission per `admission.ts`: refused ⇒ rejects with `AdmissionRefused('session-queue'|'process-queue'|'too-large')`); `pool.dispose()`.
  - Paint pass hook: `type PaintObserver = { beforeUnit(ctx, unit): void; afterUnit(ctx, unit): void }` — `render.ts` accepts `observers: PaintObserver[]` and calls them around the unit paint (empty in 1c; 1d attaches §7.7 ink-bounds/coverage here).
- Rules: document affinity — a docRef is pinned to one worker (least-loaded at first use) which parses once per `(docRef, contentPolicyId, layoutFingerprint)` and holds the document cache entry [call: Stage 1 has no separate parse worker; §9.2's 1 parse worker is Stage 2]. `layoutFingerprint` = SHA-256 hex of `officeview` version ⊕ `@napi-rs/canvas` version ⊕ `pack.json` version + hashes ⊕ locale `en-US` ⊕ contentPolicyId. `final-raster`: paint at `scale`, then the visible mark through `paintWatermark` with the injected text (`fixed`: `"officeview · protected"` [call]; `viewer`: the viewerId; `off`: none), encode PNG with `canvas.encode('png')`. `commands`: record with `createRecorder(..., { partialAlphaSnapshots: 'allow', patternSources: 'snapshot' })` exactly as `scripts/stream-s1/zero-diff.ts`; the visible mark is painted **through the recorder** after the unit. Deadline 2 s: the job's reservation moves to a retire lease; the worker is terminated and replaced; the reservation is refunded only after `terminate()` resolves. Recycle a worker after 256 pages or 64 MiB of registered embedded fonts.

- [ ] **Step 1: Failing tests.** `render.test.ts` (napi): raster of the first corpus PPTX unit at scale 1 has PNG chunks exactly `IHDR, IDAT…, IEND` (metadata-free); its decoded RGBA equals a direct paint + `paintWatermark` with the same text; `mark: off` equals a direct paint; commands reply replays (via `replayUnit` + `napiBackend`) to the same RGBA as the direct paint + mark; a GLYPH_RUN under `Aptos` yields `faces` containing no pack id and a `face-missing` diagnostic. `pool.test.ts`: 9 concurrent jobs in one session ⇒ 8 running + 1 queued; 41st ⇒ `AdmissionRefused('session-queue')`; a worker forced to hang (test hook `__hangMs`) is retired at 2 s, its reservation still held until terminate resolves, then refunded (assert via `pool.stats()`); the 257th page on a worker recycles it; two docRefs land on different workers when two are idle; same `(docRef, policy)` parses once (worker reports `parses: 1` over 3 jobs).
- [ ] **Step 2:** FAIL. **Step 3:** implement. **Step 4:** PASS; T12 still runs from its new import (`bun run s0:napi scripts/stream-s0/t12-admission.ts` GREEN).
- [ ] **Step 5:** Commit `feat(server): render pool, final-raster + commands units, occlusion seam`.

### Task 6: Caches and ceilings (§9.3)

**Files:**
- Create: `server/src/cache.ts`, `server/test/cache.test.ts`

**Interfaces:**
- Produces: `renderKey({ docRef, contentPolicyId, layoutFingerprint, recorderVersion: 1, scaleBucket, markSettings, projectionMode, policyDetails, unitId }): string` (canonical JSON of the fields in this order, SHA-256 hex); `class RenderCache { get(key); put(key, entry, bytes); stats() }` with LRU eviction at 512 MiB; document-cache accounting `DocumentLedger { charge(docRef ⊕ policy ⊕ fp, bytes); evictLRU(): docKey[] }` at 1 GiB (the worker drops the evicted parse on the next message). `markSettings` = `{ visible: kind, text, forensic: false }`.
- Rules: canonical entries hold unmarked surfaces only when the mark is applied per session; in Stage 1 the visible mark is painted in the worker, so the render cache holds **marked** units keyed by `markSettings` (sessions with equal mark settings share; different settings never share). `viewer` marks therefore never share across viewers. Eviction never removes an entry a live send is streaming (`pin`/`unpin`).

- [ ] **Step 1: Failing tests.** Key changes when any single field changes (table-driven over all 9 fields); `relaxed` vs `protected` and `fixed` vs `viewer` produce disjoint keys in both fill orders (T6 pair); LRU evicts the oldest unpinned entry once 512 MiB is exceeded; a pinned entry survives eviction; ledger evicts LRU documents past 1 GiB.
- [ ] **Step 2:** FAIL. **Step 3:** implement. **Step 4:** PASS. **Step 5:** Commit `feat(server): render cache + document ledger with §9.3 keys and §14 ceilings`.

### Task 7: Wire codecs — MANIFEST, UNIT_RASTER body, ASSET_DEFINE (client-shared)

**Files:**
- Create: `src/stream/manifest.ts`, `src/stream/asset-wire.ts`, `src/stream/raster.ts`, `tests/stream/manifest.test.ts`, `tests/stream/asset-wire.test.ts`, `tests/stream/raster.test.ts`
- Modify: `src/stream/index.ts` (export them — the lazy runtime; the core closure must not grow: `bun run gates:core`)

**Interfaces:**
- Produces:
  - `Manifest` type = the §6.2 rule-3 JSON exactly; `encodeManifest(m): Uint8Array` (UTF-8 JSON); `validateManifest(json: unknown): { ok: true; manifest } | { ok: false; field: string }` — numbers range-checked, unit sizes 1..8192, enums closed, token ≤ 512, `requiredFaces` ≤ 64 entries, unknown fields ignored.
  - UNIT_RASTER unit body (discriminator 2, inside `encodeUnit`): `u32 w + u32 h + u32 pngLen + png bytes`; `encodeRasterBody`, `decodeRasterBody(body, { maxRgbaBytes }): { w, h, png }` — rejects when IHDR width/height ≠ w/h or `w×h×4 > maxRgbaBytes` (64 MiB) **before** decoding pixels.
  - ASSET_DEFINE payload (kind 2): `u32 count` then per asset `u32 id + string hash + u8 type (0 bitmap, 1 linear, 2 radial, 3 pattern)` + fields: bitmap `u8 source (0 image, 1 canvas) + u32 w + u32 h + RGBA (w×h×4)`; linear/radial `f64 coords… + u16 stopCount + (f64 offset + string colour)×n + u16 transformCount + TransformOperands…` (same encoding as TRANSFORM operands); pattern `u32 bitmapId + string repetition + u8 hasMatrix + 6×f64?`. Strings as §7.5. `encodeAssets(assets): Uint8Array[]` splits so each message's decoded payload ≤ 8 MiB (a single bitmap over 8 MiB ⇒ throws `AssetTooLarge` and the server ships the unit as raster instead); `decodeAssets(payload): Asset[]` validates every field and the `hash` against a recomputed hash of the asset's content (same hash function as the recorder).
  - Commands-unit streaming order: all ASSET_DEFINE messages for a unit's not-yet-sent assets precede its UNIT_COMMANDS message, same `reqId`.

- [ ] **Step 1: Failing tests.** Manifest: the spec example round-trips; each closed enum rejects an unknown value; unit `w: 8193` rejected; a 1 MiB-of-spaces gzip bomb through the 1b `Reassembler` with the MANIFEST caps rejects `decoded-cap` (T6 bomb) and an oversized **valid** JSON manifest (> 256 KiB) rejects `decoded-cap` too. Assets: every asset from a recorded corpus unit (the 1b zero-diff run's first command unit with ≥ 1 bitmap, 1 gradient and 1 pattern; build one synthetically if the corpus lacks a pattern) round-trips deep-equal; a flipped RGBA byte fails the hash check; truncation at every byte offset of a small payload throws (never returns a partial list). Raster: IHDR mismatch and `w×h×4 > cap` reject before decode.
- [ ] **Step 2:** FAIL. **Step 3:** implement. **Step 4:** `bun run test tests/stream` PASS; `bun run gates:core` unchanged (core closure 299 B; index closure within recorded change).
- [ ] **Step 5:** Commit `feat(stream): MANIFEST, UNIT_RASTER and ASSET_DEFINE wire codecs`.

### Task 8: `ws.ts` — the §6.2 state machine, per-op authZ, rate gates, close-code matrix

**Files:**
- Create: `server/src/ws.ts`, `server/src/rate.ts`, `server/src/ops.ts`, `server/test/handshake.test.ts`, `server/test/close-codes.test.ts`, `server/test/ops.test.ts`

**Interfaces:**
- Consumes: Tasks 2–7.
- Produces: `createStreamServer({ config, authorize, resolveDocument, pool, caches, now?: () => number }): { fetch, websocket }` for `Bun.serve`. Client text messages (all with `reqId: u32`): `hello`, `ready { tokenJti, missingFaces? }`, `renew-ack { jti }`, `unit { token, unitIndex, scale }`, `relax { … }` (always 4009). Server text messages: `renew { token, jti, exp }`, `refused { reqId, code }`, `diag { reqId?, code: 'face-missing'|'hidden-unrecognised', … }`. Binary out: MANIFEST (reqId = hello's), ASSET_DEFINE*, UNIT_RASTER/UNIT_COMMANDS (reqId = the unit request's) via `encodeMessage` (gzip level 6 [call], 1 MiB fragments).
- Flow: upgrade checks `Origin` ∈ allowlist (else 403, no upgrade) and the per-IP rates; HELLO must be the first message, ≤ 4 KiB, JSON object with `client` string, `recVersions` int array ≤ 16 entries; resume path per §6.2 rule 7; `authorize` → null ⇒ 4003; `resolveDocument` → null ⇒ 4003; `admit` refusal ⇒ 4001; the first worker reply provides unit sizes + fingerprint for the MANIFEST; ready timer 2 s ⇒ 4002; units whose faces the READY reports missing, or not command-eligible, ship as rasters (raster also for every unit in `final-raster`). `can` = `{ search: false, copy: false, print: effectivePolicy.print, textA11y: false, present: false, notes: false, links: false }` in Stage 1 (services land in Stage 3); every op re-validates the token (signature, exp, docRef, sid bound to this socket) and its `can.*` gate. RENEW is scheduled 5 s before `exp`; expiry without ack ⇒ old token valid until `exp`, then 4004. A fingerprint change reported by a worker mid-session ⇒ 4006. Revocation API `server.revoke(sid)` ⇒ 4005.

- [ ] **Step 1: Failing tests** (real `Bun.serve` on port 0 + a `WebSocket` client, injected `now`, a fixture authorize/resolve): `close-codes.test.ts` — one test per registry row reaching its code: 4000 (`recVersions: [2]`; and a 70 KiB text message), 4001 (ODT), 4002 (no READY, advance fake timer 2 s; zero binary frames after MANIFEST), 4003 (authorize → null; wrong-viewer resume; missing resumeProof), 4004 (no renew-ack, clock past exp), 4005 (`revoke`), 4006 (fingerprint change injected), 4008 (three rate breaches in 60 s; first two are op refusals), 4009 (`relax`), 4409 (resume supersedes the first socket), and `session-rotated` as a refusal with the socket still open; the server never sends 4007. `handshake.test.ts` — happy path HELLO→MANIFEST (validates with `validateManifest`, `policy.protection: 'final-raster'`, `notices` contains `forensic-mark-stage-2`)→READY→unit→UNIT_RASTER whose PNG decodes to the pool's raster; relaxed session (authorize projection `relaxed`) → ASSET_DEFINE then UNIT_COMMANDS that replay to the direct paint + mark; `missingFaces: ['carlito-regular']` turns a Calibri unit into a raster and emits `face-missing`. `ops.test.ts` — the five Review Focus items; non-upgrade GET `/` and `/doc` → 404; `Origin` not allowlisted → 403; a stale-docRef token is refused `policy-refused`; T2 over the wire: for each of the 13 hidden fixtures (relaxed `commands` session, so text is inspectable) the decoded UNIT_COMMANDS GLYPH_RUN strings and ASSET_DEFINE payloads contain no marker bytes, and protected sessions' PNG payloads contain no marker bytes (wire inspection, §16 T2).
- [ ] **Step 2:** FAIL. **Step 3:** implement. **Step 4:** `cd server && bun test` PASS.
- [ ] **Step 5:** Commit `feat(server): WS state machine, token-per-op authZ, rate gates, close-code registry`.

### Task 9: Demo server, acceptance record

**Files:**
- Create: `server/src/main.ts` (CLI: `bun server/src/main.ts --demo --doc-dir corpus --origin http://localhost:5173`), `server/src/demo-hooks.ts` (authorize: every request is viewer `demo`, default policy; `?relaxed=1` ⇒ projection relaxed **only when `--demo`**; resolveDocument: `?doc=<basename>` mapped to a file directly inside `docDir` — reject `/`, `\`, `..`, NUL, and anything not matching `^[\w.-]+\.(docx|xlsx|pptx)$`), `server/README.md`, `server/test/demo.test.ts`
- Modify: `docs/validation/streaming-s1.md` (§"Stage 1c" section), root `package.json` script `"stream:demo": "bun run s0:napi server/src/main.ts --demo --doc-dir corpus"` (or the napi preload equivalent used by Task 1)

- [ ] **Step 1: Failing tests.** `demo.test.ts`: path traversal attempts (`../package.json`, `%2e%2e%2fx.docx`, `a/b.docx`, `x.odt`) resolve to null ⇒ 4003/4001 as appropriate; `?relaxed=1` without `--demo` is ignored (protected); an end-to-end open of three corpus files (one per format) streams every unit at scale 1 with no refusal.
- [ ] **Step 2:** FAIL. **Step 3:** implement. **Step 4:** full verification — `cd server && bun test`; `bun run test`; `bun run s0:test:napi`; `npx tsc --noEmit`; `bun run gates:core`; `git diff --check`. Record in `streaming-s1.md`: the 13-fixture T2 stream-path table (clean ×13), the close-code matrix (row → test name), content-projection identity results, pool stats from the e2e run.
- [ ] **Step 5:** Commit `feat(server): demo server and Stage 1c acceptance record`; push.

## Deferred to 1d (named so nothing is silently dropped)

`RemoteOfficeDoc` + routing (§6.3/§8.3), the browser client (`streamClient`, receipt guard 4007, replay of UNIT_RASTER/UNIT_COMMANDS, font-pack lazy asset + hashing for READY), §7.7 ink-bounds/coverage evidence on the Task 5 seam, the Playwright WS harness (`server/test/app.ts`), reconnect-restores-viewport, local-vs-others byte-for-byte test. Stage 2: forensic layer-1 (until then `forensicMark: false` + notice), scratch lease with per-session marking of shared unmarked surfaces, separate parse worker, audit retention/key rotation. Stage 3: print jobs (Stage 1 only carries `can.print`).
