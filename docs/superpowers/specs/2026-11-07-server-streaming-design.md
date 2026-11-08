# Server-streamed Office rendering: design

Status: **draft 12.1 — T11 ruling corrected (§2 item 14b, §8.1); draft 12 — Stage-0 owner rulings folded (§2 item 14; §6.4, §7.4, §8.1, §17); draft 11 APPROVED by codex r10 (reviews: codex r1–r10, agy r1, claude r1)** — draft 11 folded codex r9 (byte-grammar codes, fragment cap, aggregate admission, close forms); — drafts 2–8
folded codex r1–r6, claude r1 and agy r1; draft 9 folded codex r7; **draft 10
folded codex r8 (narrow fixes: §6.2 close-code registry, §7.1 byte grammar and
transport caps, §14 decoded admission, §16 stage-labelled gates)** — nothing
else touched. Docs-only stage continues.
Branch: `spec/server-streaming`, from main `ca952c50`. Docs-only stage: no
source edits.

Notation: **"to verify in Stage 0"** marks a claim to be proven by Stage 0's
prototypes before production code lands. Code anchors follow the codex-r6
audit of this worktree (a0204250), which corrects several round-7 anchors.

## Table of contents

1. Intent
2. Owner decisions (all decided; 2026-11-07/08)
3. Goals, non-goals, anti-goals
4. Current state (recon; anchors per the codex-r6 audit)
5. Architecture overview
6. Stream source, session policy, handshake, and token contract (normative)
7. Draw-command format (`OVSD`), recorder, occlusion, and the projection rule
8. The client side
9. The server side (pool, caches, marks, isolation)
10. Behaviour: open, failure, resume, expiry, lifecycle
11. Search, copy, and semantic text services
12. XLSX in stream mode
12a. Presentation mode for streamed decks
13. Forensic marks
14. Performance, memory and limits
15. Accessibility
16. Testing and validation (suites `T1…T12`)
17. Staging
18. Boundaries
19. Honest limits
20. Appendix: command-projection extension (Stage 2+)
21. Review history pointer

## 1. Intent

officeview renders Office documents to canvas in the browser. For a
confidential document that means the bytes sit in the browser cache; every
client-side "protection" (`src/core/watermark.ts:6–13`) is documented as *a
deterrent, not DRM*.

This design adds a second mode with a different trust model: a server parses,
lays out, and paints the document, and **streams the client only what is
drawn** — Canvas 2D draw commands or rasters plus the minimum interaction
geometry. The original file never crosses the wire in this mode; search, copy
and semantic text are server calls. The mode a viewer runs in is **the
developer's choice, made by source type** (§6.1); the library does not decide
it and nothing can be inferred by inspecting a response.

## 2. Owner decisions (all decided — 2026-11-07/08)

Binding. Superseded branches (outlines mode, HTTP/`OVSM`/406/Accept
negotiation, per-document "stream-required" marking, §18-as-separate-section)
were deleted in round 7; the rules below are the single normative home of
everything they once covered.

1. **Mode = source type (2026-11-08).** "If a doc is supplied to the viewer,
   it goes file view. And if a stream, websocket, is supplied, go stream
   mode." `data` accepts `ByteSource | StreamSource`; the branch is at the
   `OfficeFile` layer; nothing is sniffed. No per-document marking; **the
   developer's backend is the boundary** (§6.1, §19.2).
2. **Same repo, open source.** Library at the root; the server is a `server/`
   Bun workspace on `officeview/core` (§8.1). Library dependencies stay zero.
3. **Routing.** Prop **`data`** (not `source`; `OfficeFile.tsx:154`,
   component at `:175`); a stream source enters via the widened `data` union
   guarded against the `HttpSource` structural collision (§6.3). `<OfficeDoc>`
   renders a remote document through a discriminated `OfficeDocSource` arm.
   The per-unit record point stays the `page.paint(ctx)` closure inside
   `getPaintables` (`src/render/paint.ts:354`, closures from `:389`) —
   **with the §12 exception**: interactive XLSX tiles are painted through the
   viewport pipeline (`renderSheetViewport`, §12), so the unit record point
   is "the paint-call seam the pool invokes", whichever provider serves it.
4. **Server surface.** WS-only; short-lived signed tokens; rate limits; render
   pool; caches; per-session marks; session/audit tables; detection tool.
   **`server/` has no file-download route and no HTTP document endpoint** (§6.1).
5. **Content policy and projection.** Every stream session is **protected by
   default**: `final-raster` (§7.7), content policy `visible-only` (§6.4), both
   marks on. The developer's server-side `authorize(request)` hook may relax
   **projection, content policy, or marks** per session (audited, §6.1) — the
   client has no relax surface. Full-content-policy relaxation is the owner
   decision this list records; §18's "never ship hidden content" is qualified
   by it (§18 Never).
6. **Text mode (2026-11-08).** Plain text + pinned server faces only (§7.4);
   no shaping/outlining dependency; missing faces ⇒ raster units. Applies to
   command projections; `final-raster` ships pixels.
7. **XLSX depth.** Server-side sort/filter (Stage 4c) with per-session view
   state, display-text-only filter lists, no cell data in replies, its own
   rate-limit class and a wire audit; interactive tiles through the viewport
   pipeline (§12), never `getPaintables`.
8. **Presentation.** Full mode with animation and Morph (§12a); speaker notes
   only for a presenter-role viewer, Presenter View only, as marked images
   (§12a.5); media poster frames by default, optional transcoded Stage 3d.
9. **Print (2026-11-08).** Per-session setting default on (`can.print` set by
   `authorize`); pool re-renders print units at print resolution with the
   session's marks; own rate class; each job audited (§14, §17 Stage 3).
10. **Forensic mark (2026-11-08).** The keyed **spread-spectrum pixel mark**
    is the primary mark **on every unit** (§13.2); keyed text displacement is
    a second, independent layer on command-projection units (§20). Rejected:
    hidden QR codes.
    **Collusion fix (2026-11-08, Stage 0):** the payload enters the code
    through a keyed non-linear map, and a multi-candidate hit is reported
    as `collusion-suspected`, never as one viewer (§13.2).
11. **Visible mark (2026-11-08).** Per session: fixed text (default), viewer
    identity fields, or off. Both marks off ⇒ an **untraceable** session,
    recorded in its audit row (§6.1, §19.2).
12. **A11y.** `can.textA11y` defaults on; turning it off ships the §15 notice.
    It alone never authorises notes (§12a.5).
13. **Canvas backend.** Skia via `@napi-rs/canvas` for the server pool; Stage
    0 confirms; **one `CanvasBackend` seam for every canvas-construction site**
    (§9.2).
14. **Stage-0 rulings (2026-11-08, after `docs/validation/streaming.md`).**
    (a) **T2:** the six constructs today's renderer draws (hidden/veryHidden
    sheets, `w:vanish`, a hidden shape inside a docx group, a hidden pptx
    shape, a hidden slide) are confirmed as the server strip list (§6.4),
    and the strip **fails closed**: an unrecognised hidden flag is not drawn
    (§6.4). The Stage-0 RED is the expected baseline; T2 conformance gates
    Stage 1. (b) **T11 (corrected 2026-11-08):** the 380 KB chunk is the
    shared parser/renderer engine that `<OfficeDoc>` imports eagerly, not
    the worker re-export. The 38,720 B transitive ceiling applies to the
    stream client entry (`officeview/core`); the main entry keeps the
    `index.js`-alone budget plus a no-growth gate on its closure (§8.1). (c) **T4:** a server-pinned open font pack
    of metric-compatible faces (§7.4 rule 6). (d) **Corpus:** top up DOCX to
    the minimum and hand-certify image-heavy units; T9 re-runs on those.

History (3 lines): drafts 1–4 routed streams through an HTTP-manifest sniff
with `Accept`/406 negotiation, then a per-document "stream-required" flag,
before the 2026-11-08 decision moved the mode choice to the source itself and
mandated default protection; an outlines text mode existed until the same
date. The full round dispositions live in the reviews file (§21).

## 3. Goals, non-goals, anti-goals

Goals: (1) stream mode is protected by default (§6.1); (2) same pixels on
both routes — one fidelity gate per text mode (§7.4, §16 T1); (3) a stream
session never leaks original media bytes (§7.7), hidden content (§6.4), or
covered content (§7.7) — and the developer's backend keeps the file itself
off every surface including `server/` (§6.1, §19.2); (4) both marks behave
per §13 and are switchable only by the developer's server; (5) search/copy/
semantic text are server calls over the conservative visible join (§11);
(6) XLSX streams tiles through the viewport pipeline (§12); (7) sessions are
governed end to end — token contract, cache identity, retained-memory
ceilings and lifecycle rules all normative (§6.2, §9.3, §9.4, §10, §14).

Non-goals: true DRM (§19.1); a general display protocol; rewriting the local
renderer; an officeview-branded product/server; client-side content
inspection as a trust boundary.

Anti-goals: covered or hidden content in any transmitted payload (the §16
tests inspect **transmitted payloads**, not screenshots); a stream server
that also serves document bytes anywhere; a client-side relaxation surface;
a replayer that silently mispaints (unknown ops ⇒ tile, §7.6); treating
client-submitted text as code (§18 Never).

## 4. Current state (recon; anchors per the codex-r6 audit)

- **Byte sources.** `loadOfficeFile(source: ByteSource, …)` starts at
  `src/components/OfficeFile.tsx:37` (options at `:39`); the component prop is
  **`data`** (`:154`); `<OfficeFile>` declares at `:175`; `useOfficeFile`
  starts `:59`; its route key is `:92`. `loadOfficeSource(source: ByteSource,
  …)` at `src/components/loadOfficeSource.ts:202`; the local `AnyDoc` union
  is `DocxDocument | PptxDocument | XlsxDocument` (`:32`). The loader also
  parses **ODT (`parseOdt`) / RTF (`parseRtf`)** into `DocxDocument`
  (`:67–68, :117`) — the server refuses those formats (§6.2).
- **Policy headers (local mode only).** `X-OfficeView-Protection` (no-copy,
  no-print; `src/core/stream.ts:51–75`) and `X-OfficeView-Watermark`
  (`src/core/watermark.ts:164`). Stream mode has no HTTP response, so policy
  is `manifest.can.*` server-enforced (§6.2, §9.3).
- **Paint surface.** `getPaintables` (`src/render/paint.ts:354`) returns
  per-unit `paint(ctx)` closures (`:389`) — the record point. Direct painter
  exports exist (`src/index.ts:49–56`); the server uses `getPaintables` or
  the §12 viewport path. `measurerFromDoc` (`paint.ts:184–211`) hard-codes
  its canvas at `paint.ts:199` — one of the five dynamic `import('canvas')`
  sites (written `import(/* @vite-ignore */ 'canvas')`, which is why a plain
  grep failed): `src/render/paint.ts:199`, `src/core/search.ts:212`,
  `src/core/images.ts:65/:149`, `src/slideshow/render/svg-raster.ts:64`.
  Four `ctx.canvas.constructor` clone sites exist:
  `src/core/text-metrics.ts:31/:42` (side measurement canvas),
  `src/drawing/paint.ts:126`, `src/drawing/pattern.ts:155` — the text-metrics
  sites have no catch and can throw; `@napi-rs/canvas` solves constructibility
  only through `createCanvas`. All nine sites funnel through one
  `CanvasBackend` (§9.2).
- **Device-space dependence.** `snapToPixels`
  (`src/drawing/scene-paint.ts:32–41`), CTM-scaled shadows
  (`src/drawing/paint.ts:363–397`), `pattern.setTransform` in the radial
  brush (`:178–196`), sparkline scale from the transform
  (`src/xlsx/render.ts:774–775`), charts keyed by document/global DPR
  (`src/charts/paint.ts:766–768`, `src/charts/layout.ts:253–263`), text laid
  out during paint (`src/drawing/text-paint.ts:972–989`,
  `src/core/text-metrics.ts:23–42`), `'letterSpacing' in ctx`
  (`text-paint.ts:113`), scratch surfaces drawn back under skew
  (`text-paint.ts:712–765`).
- **A text-only recording context exists.** `createRecordingContext`
  (`src/core/search.ts:92`) captures fill/stroke text calls;
  `buildTextIndex` (`:315`) records a **separate** pass with the watermark
  muted (`:322–334`; module-global mute at `src/core/watermark.ts:249`).
  Painters branch on `RECORD_TEXT` (`src/core/text-recording.ts:2`, a
  `Symbol(...)` — not `Symbol.for`; e.g. `src/pptx/render.ts:81` skips the
  master underlay) — so the visual recorder never exposes it (§7.3).
- **XLSX disclosure surface.** The local interactive viewer ships formulas
  (`AuthoritativeCell`, `src/xlsx/view/contracts.ts:74–80`), hidden rows/
  cols (`SheetLayout`, `:668–674`), and raw picture bytes
  (`contracts.ts:362–366`, `XlsxImage data: Uint8Array` at
  `src/xlsx/types.ts:27`, via an `op:'drawings'` read,
  `src/worker/viewerRemoteDrawings.ts:1–18`). Sheets parse with no hidden
  state (`src/xlsx/parse.ts:290–301` ignores `state`); `getPaintables`
  paints every sheet (`paint.ts:427–429`). **A filter can reveal
  author-hidden rows**: `isVisibleSource` (`src/xlsx/view/visibility.ts:149–155`)
  returns the filter match before `baseHidden` is ever consulted;
  `src/xlsx/view/viewport.ts:176–179` documents "a filter may reveal a
  snapshot-hidden body row". Sort/filter specs are client-typed
  (`contracts.ts:709–710`; command list `src/worker/viewerProtocol.ts:31`,
  which also lists `pivot*`/`control*`). §6.4 closes these.
- **Drawings ignore `hidden`; local hidden text unparsed.** `sourceOf` reads
  `cNvPr` id/name only (`src/pptx/parse.ts:714–717`; similarly
  `src/xlsx/drawing.ts:559ff`, `src/docx/parse.ts:580,620`); RTF
  `applyControlWord` (`src/rtf/parse.ts:1130ff`) has no `\v` case; ODT does
  not filter `text:display="none"` / `<text:hidden-*>` / annotations.
- **Interactive XLSX rendering lives in the viewport pipeline.**
  `openImportSession` → `SheetGeometry` → `renderSheetViewport`
  (`src/xlsx/view/renderViewport.ts`); the static `getPaintables` XLSX path
  renders the whole sheet and **paints a too-large-sheet notice inline**
  (`paint.ts:175ff`; the `StaticExportLimitError` class actually lives and
  throws in `src/xlsx/render.ts:54, 515, 778`) — §6.4-correcting wording vs
  the round-7 text (§12).
- **Eager budget.** Ceiling 38,720 B (`scripts/eager-budget.ts:14`) which
  today gzips **only `dist/index.js`** (`:16`); `vite.config.ts:59` sets
  `preserveEntrySignatures: 'allow-extension'` — §8.1 therefore measures the
  transitive graph and adds a `dist/core.js` purity assert.
- **Contract tests.** `src/test/pixel-diff.ts` (RGBA + `diffBitmaps` +
  `renderPaintables` `:98`), `scripts/golden.ts`, `scripts/golden-corpus.ts`.

## 5. Architecture overview

```
+--------------------------------------+        +----------------------------------------+
| HOST (embedding officeview)          |        | SERVER ("server/" Bun workspace)       |
|                                      |        |                                        |
|  <OfficeFile data={source}>          |        |  WS /stream (the only endpoint)        |
|      │  file source → local view     |        |      │ origin allowlist + authorize    |
|      │  (unchanged local path)       |        |      ▼ content policy (§6.4, pre-layout)|
|      ▼  stream source → stream session        |      ▼ parse+layout per (docRef,       |
|  <OfficeDoc document={RemoteDoc}>    |        |        contentPolicy, fingerprint) §9.3|
|      │  per unit → replay/raster at  |        |  pool: CanvasBackend render workers    |
|      │   the recorded scale (§7)     |        |  recorder + occlusion (§7.3/§7.7)      |
|  search/copy → WS op (§11)           |        |  projection (§7.7) + marks (§13)       |
|  a11y/print/present → remote impls   |        |  caches, session/audit, detection tool |
+--------------------------------------+        +----------------------------------------+
```

- **Client** (`officeview` at the root): the React-free core entry
  (`officeview/core`, §8.1); the stream loader/format code/remote provider
  live in a lazy `src/stream/`; the server depends on `officeview/core`,
  never React.
- **Server** (§9.1): CanvasBackend + pool + recorder + occlusion store +
  projection + policy gate + caches + marks. **No file route, no HTTP
  document endpoint.**
- **Wire**: one WebSocket per stream session; JSON-text control messages
  carry `reqId` and binary payload kinds share one little-endian frame
  header with the same `reqId` (§6.2/§7.1).

## 6. Stream source, session policy, handshake, and token contract (normative)

### 6.1 The stream source and per-session policy

```ts
export interface StreamSource {
  kind: 'stream'
  /** A WebSocket endpoint run by the developer's backend. `wss://` expected;
   * the origin must be in the deployment's allowlist (§9.3). */
  url: string
  /** Optional preauthorization proof for the open decision (§6.2). Never put
   * in a URL path; validated at HELLO and then discarded — it is NOT the
   * session token. */
  token?: string
}
```

- **The mode is the branch, not an inference.** `<OfficeFile data={…}>`
  accepts `ByteSource | StreamSource | Promise<ByteSource | StreamSource>`;
  a file source runs today's local path untouched; a stream source opens a
  stream session. Neither branch reads the other's content. `HttpSource`
  (`src/core/stream.ts:42–46: { url, headers?, credentials? }`) is
  structurally collidable, so it gains `kind?: never` and
  `readSource`/`openByteStream` throw a typed `StreamSourceMisuseError` if
  handed one — a `wss://` url must never reach `fetch`.
- **Per-session policy (the single normative home; §3/§19 summarize):** every
  stream session is **protected by default** — `final-raster` (§7.7), content
  policy `visible-only` (§6.4), visible mark fixed-text on, forensic pixel
  mark on (§13). The developer's server-side `authorize(request)` hook
  decides per session:
  ```ts
  authorize(req) → { viewerId?: string; role?: 'presenter' | 'audience';
                     policy: { contentPolicy?: 'visible-only' | 'full';
                               projection?: 'protected' | 'relaxed';
                               visibleMark?: 'fixed' | 'viewer' | 'off';
                               forensicMark?: boolean;   // default true
                               print?: boolean } }       // default true (§2 item 9)
  ```
  Relaxations are recorded in the audit table (§9.3/§13.4); **the client has
  no relax surface** — a client-requested relax closes with `policy-refused`
  (§16 T6). Turning both marks off makes the session **untraceable** and its
  audit row records that (§19.2). §18's blanket "never ship hidden content"
  carries exactly one owner-granted exception: an audited `contentPolicy:
  'full'` session (see §18 Never).
- **Who keeps the file safe is the developer's backend.** `server/` has no
  file-download route and no HTTP document endpoint; the library cannot stop
  a developer from exposing the file elsewhere (owner words in §2 item 1).

### 6.2 The handshake, the token contract, and one state machine (normative)

Terminology: a **message** is one protocol element (it may ride several WS
frames); a **frame** is the wire boundary. Text messages are JSON objects
with `{ reqId, type, … }` (round 8: **one** correlation field, `reqId`, in
text and binary frames alike). Binary payload kinds share the little-endian
§7.1 header whose `reqId` matches the requesting text message.

```
states: HELLO → MANIFEST → READY(ack) → STREAM → (RENEW → STREAM)* → CLOSE
        HELLO{resume} → (re)MANIFEST → READY(ack) → STREAM | CLOSE
```

1. **HELLO (client → server, first message; bounded `4 KiB`):** `{ client:
   'officeview/1.0', recVersions: number[] (a set), token?: string,
   resume?: { sid, resumeProof } }`.
2. **Version compatibility is set intersection** — the server picks a version
   it supports **from the client's set**; the empty intersection closes with
   `unsupported-version` (never a downgraded guess).
3. **MANIFEST (server → client):** a binary frame of kind `MANIFEST` (§7.1
   header; **one envelope story** — the same `'OVSD'` magic as every other
   binary frame; the old separate `OVSM` header is gone) carrying the
   gzip JSON manifest:
   ```json
   { v: 1,
     docRef: "<HMAC(kDoc, sha256(bytes)) truncated 96-bit, hex>",
     kind: 'docx'|'xlsx'|'pptx', name: "<label>",
     units: { count, list: [ { w, h } ] },
     layoutFingerprint: "<hash over (officeview version, backend version,
        font set, locale, parser inputs)>",          // §9.3 canonical key input
     session: { sid, token, expiresSec, resume: 'session' },  // token ISSUED here
     policy: { contentPolicy, protection, semantic: 'visible-only' | 'off',
               contentPolicyId, viewerMarkId },
     can: { search, copy, print, textA11y, present, notes, links },
     recorder: { commands: 'ovsd-1', textMode: 'text',
                 scalesBuckets: [1, 2],
                 requiredFaces: [ { id, hash, alias? } ],  // declared here (F5 fix)
                 projection: 'final-raster' | 'culled-commands' | 'commands' },
     limits: { maxUnitPx: 8192, maxAssetBytes, wsFrameBytes, tilePx },
     notices: ['...'] }
   ```
   **Token contract (restored round 8; one definition):** `sid` is a 48-bit
   random per-session value (never the viewer's identity); `token` is
   **HS256-signed under `kToken`** with claims `{ v, docRef, sid, exp, iat,
   jti, scope: 'stream' }` (§13.4 HKDF subkey). Every operation
   re-validates **signature, `exp`, `docRef`, and the `sid` bound to the open
   socket** before acting; the `can.*` flags are the server's claim and are
   **enforced per op** — a capability the server set `false` is refused at
   the op, never merely hidden in the client UI.
4. **READY (client → server): exactly one client readiness message**, merging
   ack-session and faces-ack: `{ reqId, type: 'ready', tokenJti,
   missingFaces?: id[] }`.
   **A valid READY is required to enter STREAM** — no STREAM op may run
   before one is verified, and a server that streams without one is a
   protocol bug. **Absent READY itself is a session timeout, not a face
   verdict:** if no valid READY arrives inside the §14 `ready window`
   (2 s [call] from the MANIFEST), the server **closes with the typed code
   `ready-timeout`** (a named 4xxx close, §6.2 close-code list) and streams
   nothing. Only missing/unverifiable face evidence **inside a valid
   READY** selects the raster fallback per §7.4: `missingFaces` comes from
   the §7.4 face registry; missing or unverifiable ⇒ those units are
   rasters (never treated as all-present). The same rule holds on a
   resumed handshake (§6.2 rule 7).
5. **STREAM:** normal ops (units/viewport/tiles/show plans/notes/copy/search/
   sort/filter/print — all governed by the token + `can.*` contract).
6. **RENEW (server → client) before expiry:** the server pushes a signed
   replacement `session.token` (only the server can sign); the client
   answers `renew-ack { reqId, jti }`. Rotation window (round-8 finding-1
   fix): while RENEW is outstanding **both** the old and the new `jti` are
   accepted; after `renew-ack` only the new; if the ack never arrives the
   old token remains valid until its own `exp`, then `session-expired`.
   Requests pinned to a rotated jti after the ack are refused with
   `session-rotated` and the client re-issues them under the new token.
7. **RESUME (HELLO with `resume`):** bound to the same viewer and policy.
   **`resumeProof` (one definition):** the still-current `session.token` of
   a live session, **or** a proof the `authorize` hook freshly issues for
   that viewer. The session table stores `{ sid → docRef, viewerId, role,
   effectivePolicy, activeSocket, jti, startedAt, status }`; resume closes
   any prior `activeSocket` atomically (WS close `4409 superseded`),
   **rotates `jti`**, issues a fresh `session.token`, and re-evaluates every
   capability including `can.notes` for the resumed viewer (§12a.5). A
   wrong-viewer proof, a proof/policy mismatch, or a missing `resumeProof`
   closes with `auth-failed` — never a silent switch.
8. **Code registry (round-10 F4; one normative registry, complete):**
   the only session CLOSE codes and op-refusal codes in the protocol are
   in this table — "no other code exists" means exactly this list, and
   the two required-but-unlisted ones (`unsupported-format`,
   `layout-changed`) are restored here:

   | name | class | raised when / where |
   |---|---|---|
   | `unsupported-version` | close (`4000`) | §6.2 rule 2: empty version-set intersection; also a client transferring an over-cap frame (§7.1) |
   | `unsupported-format` | close (`4001`) | stream-server admission refuses ODT/RTF inputs, §6.2/§6.4 |
   | `ready-timeout` | close (`4002`) | §6.2 rule 4: no valid READY inside the §14 ready window |
   | `auth-failed` | close (`4003`) | §6.2 rule 7: wrong viewer/policy proof, missing resumeProof |
   | `session-expired` | close (`4004`) | §6.2 rule 6: token `exp` reached without a successful RENEW |
   | `revoked` | close (`4005`) | §9.3 session write-off (permission, credential loss) |
   | `rate-limited` | op refusal; **close (`4008`)** on the third breach inside one 60 s window [call] | §14 rate gates breached |
   | `policy-refused` | op refusal for a disallowed op; **close (`4009`)** on any client attempt to relax policy | §6.1 relax-surface rule |
   | `session-rotated` | **op refusal, never a CLOSE** | §6.2 rule 6: a post-rotation op pinned to the superseded jti; the client re-issues under the new token |
   | `face-missing` | **unit diagnostic** (not close/refusal) | §7.4 registry verdict: units stream as rasters |
   | `layout-changed` | close (`4006`) | §9.3 `layoutFingerprint` mismatch mid-session — a new open, not a repaint of stale work |
   | `oversized-frame` | close (`4007`) | §7.1: the client-side guard when a trusted server's frame exceeds the transport cap (server→client direction only) |
   | `4409 superseded` | WS close on the superseded socket | §6.2 rule 7: resume closes the old socket |

   `session-rotated` is **only** an op refusal (a typed text-message
   reply; the session stays open) — never a CLOSE. `rate-limited` and
   `policy-refused` have both forms: an op refusal by default, and the
   CLOSE numbers above under exactly the stated conditions. Close numbers
   `4000–4009` and `4409` are unique. `unsupported-version` on receipt-side transport
   overage shares the close semantics of `oversized-frame`, which the
   F1 text defines as a client-side bug guard, not a hostile-server
   defense. Nothing outside this table participates in close-or-refusal
   decisions; any new gate added to §6/§9/§13 must join the table.
9. **Lifecycle (restored round 8):** close-on-unmount; **exactly-once close**
   on source replacement; buffered chunks discarded atomically (never a
   half-painted unit).

Manifest validation (client, fail-closed): numbers range-checked; unit
sizes 1..8192 (oversized units are tiled server-side, §14); enums closed;
token ≤ 512 chars, never in a URL path (bound at HELLO); unknown fields
ignored.

**Format scope (round-7 fix):** the stream server admits **OOXML only** —
`doc.format: 'docx' | 'xlsx' | 'pptx'`; the admission check inspects the
**original input format before any normalisation**
(`loadOfficeSource.ts:67–68` normalises ODT/RTF into DocxDocument, so by the
time a parsed doc exists the original format is gone). `.odt`/`.rtf` are
refused `unsupported-format` (round-7 fix) until hidden-text filters exist
locally — a deliberate scope limit (§19.3, agy C2c).

### 6.3 Routing a stream source (round-7 fix; anchors corrected per the round-8 audit)

Verified anchors (corrected values): `src/components/OfficeFile.tsx:37`
(`loadOfficeFile` declaration; `:39` is its options object), `:92` (route
key), `:154` (`data` prop), `:175` (`<OfficeFile>` forwardRef declaration),
`:127/:140` (`pagedWorkbookOf(…)?.release()` disposal precedent);
`src/components/loadOfficeSource.ts:32` (`AnyDoc`), `:202` (the loader);
`src/components/OfficeDoc.tsx:48–51` (`OfficeDocSource`), `:344/:672`
(`getPaintables` calls), `:947–962` (the dispatch neighbourhood).

- **`OfficeFile` layer:** the `data` union widens to
  `ByteSource | StreamSource | Promise<ByteSource | StreamSource>`; a branch
  checks `'kind' in src && src.kind === 'stream'` before any byte path; the
  result state is `OfficeFileState { status: 'ready',
  document: RemoteOfficeDoc }`.
- **`OfficeDoc` layer:** `OfficeDocSource` gains a fourth arm
  `{ kind: 'remote'; session: RemoteOfficeDoc }`; the dispatch at
  `OfficeDoc.tsx:947–962` goes remote → `src/stream/remote/` lazily, local
  sheet, legacy. Local search/present/a11y code never inspects a remote doc;
  `present()` reads the manifest synchronously (§8.2).
- **Eager gate:** §8.1 measures the **transitive static-import graph of
  `dist/index.js`** (`tests/helpers/eager-graph.ts` closure) and pins
  `dist/core.js` to zero React/document imports (the §8.1 gate list closes
  this parenthetical here, so the cross-reference stays well-formed).

### 6.4 The server content policy — pre-layout, per `contentPolicyId` (normative)

"Stream what is drawn" is bounded by "drawn **for this viewer**". The gate
runs **before layout**, keyed by a `contentPolicyId` recorded in the §9.3
canonical key. **Processing is per `contentPolicyId` over an immutable
parsed source** — no policy pass mutates the parse another session sees
(§16 T6 isolation pairs). The one owner-granted exception to
"hidden content never streams" is an audited `contentPolicy: 'full'`
session (§18 Never carries the same qualification).

| Content (as parsed, worktree-verified) | Local mode | Stream rule |
|---|---|---|
| `state="hidden"/"veryHidden"` XLSX sheets (unread by `src/xlsx/parse.ts:290–301`) | painted as extra pages (`paint.ts:427–429`) | not streamed; excluded from units/sheet-nav/all text services; an authorize relaxation may stream them per audited session (default off) |
| **Author-hidden rows/cols** (`row@hidden`, `col@hidden`) | painted collapsed (authored geometry) | **purged from the server's projected sheet snapshot before any `SheetView`/`FilterEngine` is built** — so the `visibility.ts:149–155` filter-over-`baseHidden` behaviour can never expose them to a protected stream; where a sheet declares `<autoFilter>` and a `spec.range` matches that declared block, removed rows may instead ride a separate immutable `authorHidden` tier that no client filter may flip (also §12) |
| `w:vanish` runs (unread) | painted | excluded from paint/index/copy/a11y; per-document notice; T2 assert |
| PPTX hidden slides (`show="0"`) | painted as units | not streamed as units |
| **Hidden shapes** with `p:cNvPr`, `xdr:cNvPr`, `wp:docPr` `hidden="1"`, propagated through `grpSp` ancestors (unread by `src/pptx/parse.ts:714–717`, `src/xlsx/drawing.ts:559ff`, `src/docx/parse.ts:580,620`) | painted (Selection Pane hides would leak) | excluded from paint, index, copy, a11y, and show plans; T2 fixture class |
| Speaker notes / comments | notes unread; comments unread | comments never streamed; notes only as §12a.5 marked images for a presenter-role viewer |
| Hidden rows/cols **shape** (collapsed extents) | painted collapsed | axis sizes + collapsed gaps only; numbering reveals a band's existence (§19.4) — accepted, or no streamed axis numbering |
| Chart text (`src/charts/describe.ts:40–62` reads model data) | model-based text | visible-projection only: painted names/labels/categories, never model value ranges (`policy.semantic`) |
| Pivot tables / slicers / timelines (`src/worker/viewerProtocol.ts:31` lists `pivot*`/`control*`) | interactive, pre-cache | **refused on protected 4c streams** until a scoped design lands; pivot caches and slicer items backed by hidden/veryHidden sheets or hidden rows/cols are **stripped before serving** |
| ODT / RTF inputs | parsed into DocxDocument locally (`loadOfficeSource.ts:67–68`) | **refused fail-closed at stream admission** (`unsupported-format`) until hidden-text filters exist (§19.3) |
| **Any other hidden/visibility flag** the parser sees but this table does not name (e.g. a new `hidden`/`vanish`/`state` attribute value, `visibility="hidden"` in DrawingML extensions) | as authored | **fails closed** (owner ruling §2 item 14a): the element is not drawn, indexed, copied or exposed to a11y, and a `hidden-unrecognised` diagnostic is recorded; recognising it later is a policy change, not a default |

The client never participates; it renders what arrives.

## 7. Draw-command format (`OVSD`), recorder, occlusion, and the projection rule

### 7.1 One binary frame header for every payload kind — one envelope story (normative; round-9 F2 completes it)

**Every binary WS payload uses exactly one envelope:** the little-endian
header below — one `magic` (`'OVSD'`) for **every** binary frame including
`MANIFEST` (the draft-7 `OVSM` header is gone; the manifest is a frame of
kind `MANIFEST`).

```
First-fragment byte grammar (all LE, offsets from the start of the
binary frame):

offset  size  field        meaning
0       4     magic        u32 LE, 'OVSD' (every kind — one magic)
4       1     ver          u8, negotiated format version (§6.2 set intersection)
5       1     kind         u8, one of: MANIFEST, ASSET_DEFINE, UNIT_RASTER,
                           UNIT_COMMANDS, XLSX_TILE, SHOW_PLAN (OVSP §12a.2),
                           NOTES_RASTER, PRINT_RASTER
6       4     reqId        u32 LE, the correlation id (matches the requesting
                           text message); must reference an expected/accepted
                           request — unsolicited or unmatched ids are dropped
                           before any allocation (round-10 F3)
10      4     frameSeq     u32 LE, zero-based fragment index within the message
14      4     crc32        u32 LE, CRC-32 of THIS fragment's `payload`
                           bytes only (never cross-fragment, never of headers)
18      1     flags        u8, bit flags (below)
19      4     payloadLen   u32 LE, this fragment's payload byte count
23      …     payload      payloadLen bytes
```

**Assigned codes (normative):** `kind` — `1` MANIFEST, `2` ASSET_DEFINE,
`3` UNIT_RASTER, `4` UNIT_COMMANDS, `5` XLSX_TILE, `6` SHOW_PLAN, `7`
NOTES_RASTER, `8` PRINT_RASTER; `0` and `9–255` reserved (⇒ drop). Unit
`discriminator` — `1` commands, `2` raster, `3` tile; others reserved.
**CRC-32** everywhere means IEEE 802.3 CRC-32 (reflected polynomial
`0xEDB88320`, init and final XOR `0xFFFFFFFF` — the zlib `crc32`). `magic`
is the four ASCII bytes `4F 56 53 44` on the wire.

flags bits: bit0 `final` (this is the last fragment); bit1 `ext01` (a
10-byte first-fragment extension sits between the header and the payload —
see below; set on fragment 0 only);
bit2 `compressed` (always 1 today, see the compression rule); the rest are
reserved and must be 0. `kind`/`reqId`/`compressed` are **invariant across
the fragments of one message**; `frameSeq` advances 0..fragCount−1; the
`final` bit must be set on exactly one fragment, whose frameSeq equals
fragCount−1 (a final bit contradicting fragCount/sequence fails the whole
message).

**ext01 extension (first fragment only, 10 bytes at fixed offset 23, its
presence is flag bit1 — round-10 F2 fix):**

```

offset  size  field
23      8     totalLen     u64 LE — the complete COMPRESSED message length
                           across all fragments incl. the wholeCrc trailer
31      2     fragCount    u16 LE — total fragment count
33      …     payload      begins here on fragment 0. `payloadLen` and
                           `crc32` cover the payload bytes only — never the
                           header or the ext01 extension; `totalLen` is the
                           sum of all fragments' `payloadLen`s
```

**Compression boundary — exactly one compressor (round-10 F2):** the
**complete message payload is gzip-compressed once, whole, before
fragmentation**; nothing else in the path compresses. The old per-unit
"uncompressed header + gzip body" description is deleted — `unitId`,
`byteLen`, `opCount` and the discriminator now travel as **plain fields
inside the once-compressed payload stream** (they are decoded fields, not
framing fields), so there is no second `gzip` compressor anywhere.

**Trailer (round-10 F2):** the message ends with a 4-byte trailer appended
**outside compression, inside the encoded message**: `wholeCrc u32 LE` =
CRC-32 of the **complete decompressed payload**. It rides as the final 4
bytes of the compressed blob, **lies entirely inside the final fragment**
(an encoder never splits the trailer across fragments), is covered by the
final fragment's `crc32`,
**is counted inside `totalLen`**, and is validated after inflation — end-
to-end proof of the decoded byte stream without re-hashing headers.

**`totalLen` == byte-accurate:** totalLen is the encoded size of the
compressed blob + trailer; it equals the **sum of every fragment's
`payloadLen`** with no exception — for a single fragment, `totalLen ==
payloadLen` whether or not `ext01` is present (absent `ext01` ⇒ inferred
`fragCount = 1`, `totalLen = payloadLen`) — a decoder that sums fragments and gets a different
number fails the whole message before inflation.

**Rejection rules (fail-closed, in order):** wrong magic ≠ 'OVSD' or wrong
verified `ver` ⇒ drop; mismatched invariant fields across fragments ⇒ drop
the message; final-bit/fragCount contradictions ⇒ drop; missing `ext01`
on a multi-fragment message ⇒ drop; summed payloadLens ≠ totalLen ⇒ drop
(pre-inflate); inflated size > decoded cap ⇒ abort mid-inflate (§14
streaming admission); wrong `wholeCrc` ⇒ drop after inflation, before
opcode/JSON validation; nothing partially validated or painted (§10).

**Single-fragment grammar:** a single-fragment message sets the `final`
bit, carries frameSeq 0, and MAY omit `ext01` when `payloadLen` alone
defines the whole message; no other special case.
- **Bounded reassembly (normative):** pre-allocation checks run before any
  buffer; fragment count ≤ 256; reassembly window 5 s [call] with abort on
  timeout (the partial buffer is **discarded**, the client re-requests
  once, nothing partially validated); see the rejection rules and §14's
  encoded/decoded caps above — these bound every kind alike.
  - **Application-fragment cap, both directions (normative):** every
    application fragment — header + optional `ext01` + payload — is
    **≤ 1 MiB** (`wsFrameBytes`) and is sent as **exactly one WebSocket
    message**; per-kind reassembled-message caps (§14) are separate and
    larger. A message above 1 MiB encoded is always fragmented.
  - **Transport caps are per endpoint (round-10 fix, codex r8 F1 — the round-9
    wording overstated the browser):** the **server** sets a native
    receiver cap (`ws` `maxPayload`, the largest encoded fragment = 1 MiB)
    on client→server messages, so a hostile client cannot deliver an
    oversized frame past the server's parser. **Server → browser**, the
    *trusted* side, bounds every message it sends: no server-emitted
    fragment exceeds 1 MiB and no reassembled message exceeds the §14
    per-kind encoded cap. **The browser `WebSocket`
    API has no pre-delivery size cap** (WHATWG: `new WebSocket(url,
    protocols)` takes only url+protocols), so the client checks each
    message's size (`ArrayBuffer.byteLength`, or `Blob.size` when
    `binaryType` is `blob`) against the 1 MiB fragment cap **on receipt** and closes with the typed transport
    code `oversized-frame` — a client-side guard against its own trusted
    server's bugs, **not** a security boundary against a malicious server
    (the server is trusted; the client is not, and a malicious server
    controls what the browser buffers before that guard runs).
    A custom `streamClient` factory returns the **same** browser API and
    therefore cannot add a pre-delivery cap either; the round-9 claim that
    `streamClient` passes a `maxPayload` is retracted.
- Per-unit body (inside the compressed message payload): `unitId u32 +
  byteLen u32 (decoded unit-body bytes) + opCount u32 + discriminator u8
  (commands | raster | tile) + raw unit record` — never a whole-document
  `unitIndex[]`.
- A `TILE` op carries its bitmap inline (within the compressed message),
  never as a data URI.

### 7.2 Why commands (short)

Command streams compress well, are resolution-independent *given the
recorded scale bucket* (§7.3), and make per-session forensic edits a mutate
pass on cached records (§13). Resolution dependence is why the pipeline
names its scale everywhere.

### 7.3 The recorder (normative contract) + `CanvasBackend`

- **Facade over a real backing 2-D context** (the generalisation of
  `src/core/search.ts:92`), so `'letterSpacing' in ctx` truthiness
  (`src/drawing/text-paint.ts:113`), `measureText` during paint
  (`text-paint.ts:972–989`, side canvases in `src/core/text-metrics.ts:23–42`),
  scratch-surface `drawImage` (`text-paint.ts:712–765` — snapshotted by
  content hash at draw time), and wrapper patterns with recorded
  `setTransform` (`src/drawing/paint.ts:178–196`) all behave as on a real
  canvas.
- **Recorded subset (typed, census-final; round-8 lost-rule-13 restore):**
  transforms (setTransform/translate/scale/rotate/transform); state —
  `alpha`, `composite`, `smoothingEnabled` **and `imageSmoothingQuality`**,
  `lineWidth/lineCap/lineJoin/miterLimit`, `setLineDash/lineDashOffset`,
  **`letterSpacing` and `wordSpacing`**, `direction`, `lang`,
  **font-feature/font-variant family state the painters set**; styles —
  solid colour, `createLinearGradient`/`createRadialGradient`,
  `createPattern` + recorded wrapper `setTransform`, the four shadow fields;
  geometry `beginPath/moveTo/lineTo/quadraticCurveTo/bezierCurveTo/arc/
  arcTo/ellipse/rect/roundRect/closePath`; `fill/stroke` (incl. the combined
  op), `fillRect/strokeRect/clearRect/clip`; text `font/textAlign/
  textBaseline` + `fillText` (with `maxWidth`) + `strokeText` + glyph runs;
  `drawImage` (9-arg forms); and the raster-inlining ops of §7.7.
- **Coordinate encoding (restored):** recorded coordinates are IEEE-754
  **little-endian doubles** by definition; length prefixes and integer
  fields are LE; the census (T3) is the closure proof.
- **`RECORD_TEXT` is never exposed by the facade** (`RECORD_TEXT` is
  `Symbol(...)`, `src/core/text-recording.ts:2`; painters branch on it,
  e.g. `src/pptx/render.ts:81`) — the text index stays its own second pass
  through `createRecordingContext` (`src/core/search.ts:92`) +
  `buildTextIndex` (`:315`) with the watermark muted
  (`src/core/watermark.ts:249`).
- **Scale discipline + server keys:** record at the session's scale bucket
  and replay 1:1 at it. **The pool injects the DPR/scale explicitly**
  (round-8 restore of draft 6's rule): the recorder never reads browser
  globals (`devicePixelRatio` etc.) — the charts DPr/PaintOptions seam
  (§8/§9.2) supplies the bucket, and `setWatermarkMuted` pass-interleaving
  guards stay (§9.2).
- **`registerFont(request) → lease`: the server-side font-registration
  seam, beside `CanvasBackend` (round-9 finding 6).** Registration is a
  pool-level injected interface — the browser/worker adapters exist at
  `src/core/fonts/register.ts:9/:21/:30/:40` (worker one exported), and
  the pool threads a conforming implementation into `getPaintables`
  exactly like `CanvasBackend`.
- **`CanvasBackend`:** one seam for every canvas construction the recorder
  pool uses — `{ createCanvas(w, h), decodeImage(bytes), loadImage(bytes),
  rasteriseSvg(bytes) }` — threaded into the five dynamic import sites
  (`src/render/paint.ts:199`, `src/core/search.ts:212`,
  `src/core/images.ts:65/:149`, `src/slideshow/render/svg-raster.ts:64`)
  and the four `canvas.constructor` clone sites
  (`src/core/text-metrics.ts:31/:42`, `src/drawing/paint.ts:126`,
  `src/drawing/pattern.ts:155`). **Correctness note (round-8, per audit):**
  the text-metrics sites have **no catch** and can throw (they do not
  silently return `undefined`); the other sites `catch → undefined`, which
  silently drops gradient/warp/pattern fills on the server — the backend
  removes both failure modes and Stage 0's matrix
  (backend-constructibility) proves it.
- **Render deadline:** 2 s [call] per unit ⇒ drain-then-refund worker
  retirement; a paint cannot be aborted mid-`paint()`; the unit retries
  once, then tiles.

### 7.4 Text mode — one mode (plain text + pinned faces)

Glyph runs travel as strings + advances; replay uses server-pinned faces
whose hashes the client confirmed:

1. **Identity list declared in the manifest** (`recorder.requiredFaces`,
   round-8 fix; hashes only — SHA-256 of face bytes truncated to 96 bits,
   with ids/aliases). Face bytes never ship.
2. **Bounded acknowledgement** inside the single READY message
   (`missingFaces?: id[]`), one-time per session.
3. **Match = hash equality.** A client that cannot hash the pinned bytes
   (an installed system face with no byte access) reports it **missing** —
   an unverifiable face is an absent face.
4. **Missing READY, or missing/mismatch faces:** those units stream as
   rasters (`face-missing` diagnostic). Substituted faces are never
   painted; the zero-diff claim is voided only for the affected units.
5. **Fidelity:** replay fidelity (recorder→replayer at the recorded scale,
   same backend, same pinned faces) is the zero-diff claim; renderer
   fidelity to the local hinted painter is a published tolerance per
   script family (T4 fixtures).
6. **Pinned open font pack (owner ruling §2 item 14c).** The server pins a
   set of open (OFL) faces and lays out with them; the client package
   carries the same bytes as a lazy asset (never in the eager closure,
   §8.1), so rule 3's hash match succeeds. A document family maps to a pack
   face **only when the pack face is metric-compatible** (same advances), so
   line breaks match the authored layout: Calibri → Carlito, Cambria →
   Caladea, Arial/Helvetica → Liberation Sans, Times New Roman → Liberation
   Serif, Courier New → Liberation Mono, plus those open families
   themselves. Families with no metric-compatible open face (Calibri Light,
   Aptos, Trebuchet MS, Century Gothic, …) stay rule 4 rasters unless the
   document embeds them. The alias table is versioned in the manifest with
   the face hashes. Stage-0 T4: raster rate 72.7 % (48/66 units) without the
   pack, 28.8 % (19/66) with it; Calibri Light alone is 14 of the 19.

### 7.5 Asset table and op list

Assets keyed by content hash: drawn-bitmap snapshots (§7.7), decoded
scratch surfaces, pattern/gradient wrappers (`setTransform` included),
pinned-font hashes (never bytes). Ops (census-final): STATE_PUSH/POP;
TRANSFORM; STYLE_COLOR; STYLE_OBJ; STATE_SETUP (dense dword:
width/join/cap/miter/dash/dashOffset/alpha/composite/**smoothing quality**/
letterSpacing/**wordSpacing**/direction/lang/**font-variant/feature
state**/shadow fields); PATH_*/PATH_END {FILL,STROKE,FILL_STROKE,CLIP};
GLYPH_RUN (f64 advances + string; **no outline opcodes**); IMAGE_DRAW;
TILE; RASTER_BLOCK.

### 7.6 Versioning, tiles, runtime fallback

Version compatibility is the §6.2 set intersection; the §7.1 header's `ver`
must fall inside the negotiated set or the frame fails closed. Units the
recorder cannot express stream as tiles (per-unit discriminator); mixing
command and tile units is expected. Replayed-frame validation failure
drops the unit and re-requests once.

### 7.7 The projection rule (P-rule) and occlusion store — normative

> **P-rule.** Every response of a stream session contains only content that
> belongs to the unit's **final composited image** — no covered text, no
> clipped-out media, no full z-order history. Protected sessions satisfy it
> **by default and without configuration**; the only departure is a
> server-side `authorize` relaxation to `commands` for a specific audited
> session (§6.1); the client is not a party to it.

- **`final-raster` (default):** whole units painted to the final frame
  server-side, shipped as rasters; occlusion is structural.
- **`culled-commands` (protected, opt-in; details in appendix §20, the
  M-rule's normative home):** ships commands after a **stored per-command
  occlusion proof**; the M-rule constraints and the mixed-unit rule are
  defined in §20, so a unit failing them downgrades to `final-raster`.
- **`commands` (relaxed sessions):** ordinary stream; an overlay mark is
  removable by a modified replayer — stated plainly, because a relaxed
  session already trusts the replayer (§13.2/§20).
- **Stage-1-mandatory semantic occlusion (mandatory from Stage 1 wherever a
  paint surface exists; the command-culling *store* is the only Stage-2
  item):** during the visual paint pass that serves the unit, the pool
  records (a) per-text-run **ink bounds** and (b) a conservative
  **opaque-coverage grid** (alpha ≥ 0.95 extents) updated after each
  retained draw op. A text run is servable iff **100 % of its ink bounds
  survive to the final composited surface**; anything unresolved — a state
  the census does not model, an unmodelled compositing op — is **withheld**
  (fail-closed). The instrumentation lives in the pool's record /
  projection pass, so it applies to **every provider that paints a
  unit**: the `getPaintables` unit closures (§2's record point) and
  the §12 XLSX viewport-tile path (below) alike.
- **No original media bytes (every projection):** `drawImage` ships the
  **drawn-bitmap snapshot** (as-painted), source-rect cropped
  (`src/drawing/scene-paint.ts:132–150` crops via `drawImage` source rects),
  re-encoded metadata-free at the recorded scale, keyed by pixel hash;
  other scales re-request.
- **Wire audits inspect transmitted payloads**, not screenshots (§16 T2).

## 8. The client side

### 8.1 The React-free core entry and honest eager gates

- Build `officeview/core` (`src/core-entry.ts`, name [call]) in a dedicated
  Vite **separate pass** with React externalised. Gates (honest because
  `eager-budget.ts:16` only gzips `dist/index.js` today):
  1. **Transitive-graph budget (owner ruling §2 item 14b):** sum gzipped
     bytes over **`dist/core.js`'s** static-import closure
     (`tests/helpers/eager-graph.ts` closure) — ceiling 38,720 B. The
     stream client must not statically reach any parser or painter
     module. The main entry keeps `eager-budget.ts`'s `dist/index.js`-alone
     ceiling (38,720 B) **plus** a no-growth gate on `dist/index.js`'s
     closure against the Stage-0 baseline (417,436 B; growth ≤ 16 B gzip
     recorded as a diff). Stage-0 T11 found that closure includes a 380 KB
     shared engine chunk (rollup names it `office-worker-*.js`; 56 xlsx,
     19 drawing, 17 core modules) that `<OfficeDoc>` imports eagerly; a
     lazy-engine refactor of the main entry is out of streaming scope.
  2. **`dist/core.js` purity assert:** zero `react`/`document` imports.
  3. **Main-bundle module-map identity:** `dist/index.js`'s map identical
     to the pre-branch build (± the branch's own discriminant); growth
     ≤ 16 B gzip recorded as a diff. The pre-branch baseline is taken
     **after** the worker-subpath move above, which is the one recorded
     intentional map change.

### 8.2 Component surfaces

- `<OfficeFile data={ByteSource | StreamSource | Promise<…>}>` (prop
  `data`, `OfficeFile.tsx:154`; `<OfficeFile>` at `:175`): the
  `HttpSource.kind?: never` guard plus the typed `StreamSourceMisuseError`
  make a stream source incapable of reaching `fetch`.
- `streamClient?: (source: StreamSource) => StreamClient` — the WS factory
  hook (mirrors the `viewerWorker` factory pattern,
  `OfficeFile.tsx:63–80`); default `WebSocket` + HELLO + heartbeat.
  `onStreamNotice(diag)` — deduplicated load notices.
- **`startPresent` stays synchronous** (round-7 fix): on a remote document
  the eager `present.ts:184–246` path reads the manifest synchronously
  (slide count, `can.present`, `can.notes`), refuses `presenterView: true`
  when `can.notes` is false (so a non-presenter never gets a blank
  `window.open`), and only then dispatches to the lazy remote show runner.

### 8.3 The remote document provider (normative shape)

```ts
RemoteOfficeDoc = {
  kind: 'remote', manifest,
  paint(unit, scale: 'fit'|number, dpr, canvas): Promise<void>, // replay or tile paste
  units(): Promise<UnitSpec[]>,
  search(query): Promise<HighlightRect[]>, copy(range): Promise<string>, // §11
  close(): void,
}
```

Dispatch at `OfficeDoc.tsx:947–962`; state wiring via `OfficeFileState`
(§10). Replay is 1:1 at the recorded bucket; zoom re-requests. The client
ignores the host watermark prop for a remote document (the baked mark is
authoritative). **Libraries: no local text selection, no local text-index
build, and no local `present` pipeline runs for a remote document** —
those consumers map to the remote interfaces above (§11/§12a, restored as
normative rules from draft 6).

## 9. The server side

### 9.1 Package layout [call]

```
server/  (bun workspace; deps: officeview workspace + @napi-rs/canvas)
  src/
    ws.ts             the only endpoint: upgrade (origin allowlist + authorize),
                      the §6.2 state machine, op routing, per-op authZ
    authorize.ts      authorize(req) → {viewerId, role, policy…} (§6.1)
    tokens.ts         token schema (§6.2), issuance/rotation/jti table,
                      per-op validation helpers
    sessionTable.ts   sid → {docRef, viewerId, role, effectivePolicy,
                      activeSocket, jti, startedAt, status}
    audit.ts          sid → {viewerId?, ip, startedAt, keyGeneration,
                      policyId, markSettings} (attribution; §13.4)
    pool.ts           CanvasBackend render pool + deadline retirement (§9.2)
    cache.ts          canonical cache + retained ceilings + eviction (§9.3)
    record.ts         recorder facade + scratch snapshot service (§7.3)
    projection.ts     §7.7 raster/projection + snapshot hashing
    occlusion.ts      §7.7 Stage-1 occlusion + §12 XLSX analytic occlusion
    policy.ts         §6.4 pre-layout gate
    forensic.ts       §13; detection.ts (CLI scripts/detect.ts)
    xlsxTiles.ts      §12 viewport tiles (renderViewport path)
    print.ts          §2 item 9 print jobs (per-session, audited)
    config.ts         keys/flags, fail-closed
  test/               (bun test; §16)
```

### 9.2 Render pool and `CanvasBackend`

- Pool shape mirrors `WorkerRenderEnv`
  (`src/worker/office-worker.ts:42–47`) with `createCanvas` bound to the
  chosen `@napi-rs/canvas` backend through the one `CanvasBackend` seam
  (all five `import('canvas')` sites + all four `canvas.constructor` clone
  sites, §4).
- Defaults [call]: 1 render worker per 2 vCPU (cap 8), 1 parse worker,
  parse-once-per-(docRef, contentPolicy, fingerprint), per-job decoded
  budget reserved up front, 2 s deadline → drain-then-refund retirement.
  **Worker and font recycling (restored round 8):** a worker that has
  painted ≥ 256 pages or whose registered embedded-font budget exceeds
  64 MiB is recycled before its next assignment (embedded fonts cannot
  unregister from a backend), and the §13 measurement device shares the
  same quota system.
- Pass interleaving: `setWatermarkMuted` is module-global until
  context-scoped, so a worker runs a paint pass **or** an index pass,
  never interleaved.
- Per-session outstanding units: 8 (queue 32); rate-limited by §14.

### 9.3 Caches, ceilings, and token/notes isolation (normative)

**Cache key (single form everywhere; full effective-policy binding
restored round 8):**

```
docRef ⊕ contentPolicyId ⊕ layoutFingerprint ⊕ recorderVersion ⊕
scaleBucket ⊕ markSettings (visible-mark kind + text/fields + marks
on/off; forensic mark on/off and its coding parameters) ⊕
projectionMode ⊕ policyDetails (the §6.4 content-projection id per
policy level; the XLSX per-session `viewStateHash` on interactive
tiles; the presenter notes sidecar namespace
`notes:<presenterId>` for notes-bearing units) ⊕ unitId (subkey)
```

- Canonical entries hold the **pre-forensic unit surface** (commands or
  lossless-raster/tiled-RGBA surface) + shared asset table + culling
  proof — shared across sessions, never per-sid pixels.
- **Per-session mark+encode** runs through a bounded **scratch lease**
  (≤ 2 final-raster units per session, pre-admitted against §9.2, streamed
  then released); mark-off sessions may share final encoded bytes.
- **Document cache** keyed by
  `(docRef, contentPolicyId, layoutFingerprint)` — parsed doc, projected
  visibility, and text index (join outputs of the §6.4 projection); a
  relaxed and a default session are different entries and nothing leaks
  across them. A fingerprint mismatch mid-session terminates the session
  with a typed `layout-changed` (restored round 8), never a repaint of
  stale work.
- **Notes namespace (restored round 8):** presenter notes assets live in
  a separate `notes:<presenterId>` sidecar namespace, never addressable
  through the page-asset hash table; **any `ASSET_DEFINE` frame or
  asset-id lookup** is authorized against (contentPolicy, notes
  namespace, marks) before release, so a guessed or reused notes id
  resolves to nothing for a non-presenter (T6). Two presenters with
  different notes rights therefore never share a cache tuple.
- **Ceilings:** the §14 retained-ceiling values (document cache 1 GiB,
  render cache 512 MiB — one normative home there) are enforced by the
  same admission that governs scratch leases; when a cache must evict,
  least-recently-used entries release first and a session's live leases
  never share another session's eviction.
- **Close vs drain (round-9 finding-4 fix; one normative home §10):** a
  session close ends **session ownership**; a running paint's charge
  moves to an explicit **retirement owner** (the pool's retire-lease) and
  is refunded **only after the worker drains** — it is never released at
  close (that would be exactly the premature refund T6/T10 forbid).
  Cache handoff on eviction also **precedes** the producer's reservation
  refund. The lifecycle rules live once in §10; §9.3 references them.
- Session table (run state) evicts **at session close**; the audit table
  is **write-once, deleted-on-expiry** (90-day default [call]) — both
  restored verbatim in substance (finding-3's lifecycle rules).
- Session run table fields: `sid → { docRef, viewerId, role,
  effectivePolicy, activeSocket, jti, startedAt, status }`.

### 9.4 Marks: visible + forensic, per session (normative)

- **Visible mark:** composited into the final raster server-side
  (`paintWatermark` machinery, `src/core/watermark.ts`), fixed-text
  default or viewer identity fields or off (§2 item 11); a fixed mark lets
  pre-forensic entries share, per-session settings still key `markSettings`.
  Foggy wording from draft 7 ("one cache entry serves every session") was
  already retracted: the surface is shared, the marked encode is
  streamed per session.
- **Forensic layer-1 pixel mark** (§13.2): per-sid perturbation of the
  page-keyed basis applied during the session's scratch-lease encode; never
  stored per-sid in the canonical tier. Fleet-wide relies on the shared
  pre-forensic surface + per-session encode.
- Ordering (one pipeline): paint → (protected) projection/occlusion pass →
  Layer-2 text shift only on command-unit records with glyph coordinates →
  visible + Layer-1 mark composite → flatten/encode.

## 10. Behaviour: open, failure, resume, expiry, lifecycle

**(Open)** developer passes source → §6.3 branch; stream arm opens WS
(`streamClient`), HELLO, origin allowlist + `authorize`, MANIFEST, READY,
unit/notes/print requests. Pool renders → projection + marks → frames;
client replays 1:1 at the bucket.

**(Failure/refusal/expiry)** typed `StreamSessionError` codes; a malformed
HELLO fails closed with no allocation beyond a bounded parse (truncated
envelope or malformed manifest alike); a token aging out mid-flow ends
with CLOSE `session-expired` unless `renew` lands first (§6.2 rule 6);
`revoked` closes once with atomic buffer discard; a presentation exits and
`ended` resolves; recovery is a developer-driven re-open. Replayer-side
unit validation failure drops the unit, re-requests once, then shows the
missing-unit diagnostic.

**(Lifecycle — the normative home, round-9 finding-4.)** close-on-unmount;
**exactly-once close** on source replacement (and close-on-unmount tied to
the same generation counter); a running paint's charge moves to the pool's
**retirement-owner lease** at close and is refunded only after the worker
drains — with the §9.3 cache handoff preceding the producer reservation
refund. Buffered chunks are discarded atomically, never a half-painted
unit. T6/T10 cover each path.

**(Isolation)** per-session caches/keys/marks, draft socket supersession on
resume, and per-`contentPolicyId` processing are all covered by §16 T6
pairs.

## 11. Search, copy, and semantic text services (normative)

- **Join rule (single home):** every semantic answer runs the §7.7
  conservative visible-projection join before release — Stage-1
  implementation independent of the deferred culling store (§17 gating
  matches). Text is servable iff (a) its range sits fully inside the
  painted clip band, (b) 100 % of its ink bounds survive the
  opaque-coverage grid to the final surface — or, for §§12 XLSX filter
  values and off-viewport cells, the **analytic occlusion rule** below —
  and (c) it is not another surface's private text. Any unresolved case ⇒
  **withhold**, never approximate.
- **Search** (`can.search`): server op against the index built in the
  §7.3 second pass with the visible mark muted (`src/core/search.ts:92,
  315`; mute at `:322–334`); answers are `HighlightRect[]`
  (`src/core/overlay.ts:7`).
- **Copy** (`can.copy`): display text of the selected range; XLSX serves
  formatted display text for selected cells, never values/formulas.
- **A11y/semantic** (`can.textA11y`): §15, from the same index, subject to
  §6.4. `can.textA11y` never authorises notes (§12a.5).
- **End-to-end semantic assertions (restored round 8, rule 16):** visible-
  mark text never appears in search/copy/a11y answers (an end-to-end T2
  assertion, not just the second-pass mute); a remote document **never**
  builds a local text index or uses local text-selection machinery; all
  text answers hold **exact locale/BiDi** ordering (`RECORD_TEXT`
  sources carry `scope`/`order` — see `src/core/text-recording.ts:9–16`);
  a disabled service states **"text unavailable in this viewing mode"**
  once.
- **Chart text:** visible-projection only, no model value ranges.

## 12. XLSX in stream mode

- **Tiles via the viewport pipeline (round-7 fix; anchors corrected
  round 8):** the server pool paints tiles with **`renderSheetViewport`**
  (`src/xlsx/view/renderViewport.ts`) over `openImportSession`/
  `SheetGeometry` — not `getPaintables`, whose XLSX path is a static
  full-sheet render that **paints a too-large-sheet notice** (`paint.ts:
  175ff`) and whose `StaticExportLimitError` class actually throws in
  `src/xlsx/render.ts:54, 515, 778`; neither path has a viewport, tile
  rect, `RowOrder`, or tileable visibility. A tile is keyed by the §9.3
  canonical key + `sheetId + tileRect` and ships `XLSX_TILE`.
- **Disclosure boundary:** no formulas (`AuthoritativeCell`,
  `contracts.ts:74–80`), no hidden rows/cols metadata (`SheetLayout`,
  `:668–674`), no drawing bytes (`SheetDrawings.images`, `:362–366`);
  paint-facing layout only; collapsed extent reveals a hidden band
  exists (§19.4); hidden sheets gated pre-layout (§6.4).
- **Staged interactions:** 4a static + read-only tiles; 4b
  selection/freeze/navigation parity; 4c server-side sort/filter
  (owner decision, §2 item 7) — with the C1 purge (§6.4: protected
  sessions never let `isVisibleSource` flip author-hidden rows),
  **server-side target authorization** (sort keys/filter ranges validated
  against the policy-approved view before engine-run, refused before
  order/counts/tiles can move), and **filter values through the §11
  join**. Pivot/slicer/timeline commands are **refused on protected 4c
  streams**; pivot caches and slicer items backed by hidden/veryHidden
  data are stripped before serving.
- **Analytic cell occlusion (round-6→8 rule, rewritten round 9 finding 3 for
  correctness):** filter values and any text about a cell must pass the §11
  join **where a visual pass exists**. For rows/tiles never rendered
  (off-viewport cells, whole-column filter lists), the server computes the
  conservative **sheet-coordinate occlusion check** on **transformed painted
  bounds**, not raw anchor rectangles: per `XlsxDrawing`, take the union of
  `drawingExtent`-style conservative painted extents
  (`src/xlsx/view/sheetCells.ts:23–73` — anchor placement transformed by the
  drawing's own rotation and group transforms, including custom geometry),
  **padded by its conservative effect extents** (shadow / glow / soft-edge
  margins) when the drawing declares them, and expanded to cover **cell text
  that spills beyond its grid rectangle** (spilled ink can reveal or
  obscure through the edge of the covering shape). A cell's grid rect
  (plus its own spill rect) intersecting any such painted extent means the
  cell **withholds** from filter values (fail-closed). `anchors.unresolved
  > 0` for that sheet's drawing set (`XlsxDrawingAnchors`,
  `src/xlsx/types.ts:149`), an effect the projection cannot bound, or a
  transform it cannot compute ⇒ **every cell withholds** — never a partial
  proof. `drawingExtent` alone is a geometry helper, **not** an
  effects/security proof. T8 adds codex's named rotated-box adversary: a
  `100×20` anchor-box drawing at (0,0) rotated 90°, whose covered region
  includes `x: 40..60, y: 30..50` (cells the anchor-rect-only test would
  leak) must withhold those cells' filter values.
- **Rate classes:** sort/filter ops and print jobs each carry their own
  §14 class.
- **Wire audit** covers the §6.4/§12 rows (T8).

## 12a. Presentation mode for streamed decks

`present()` works on a remote document with **animations, transitions and
Morph**, played back from server-drawn layer bitmaps + a bounded show plan
(`OVSP`, §12a.2). The file never reaches the client.

### 12a.1 The split (verified anchors)

Server: parse, content policy (§6.4), planner, layer raster
(`src/slideshow/render/plan.ts:17–29` `PlanLayer`, `PptxShape[]`
**retained server-side only**), `raster.ts:41` `rasterLayer`, Morph match,
appearance projection §12a.3, marks. Client: scheduler, behaviours,
easing, masks, transitions, compositor, input, presenter surfaces — the
same engine code the local player uses, minus any shape/package read.

### 12a.2 The show plan (`OVSP`), bounded schema (round-7, hardened round 8)

One per slide, framed with the §7.1 `SHOW_PLAN` kind:
- `layers[]`: `id, assetId, bounds, z, partTag` — resolved base geometry
  and per-part splits (`fill/line/text`), no shape data.
- `steps[]`: order, appearance-action eligibility flags (§12a.3), trigger
  kinds with **hit-region** rectangles only.
- `keyframes[]`: `{ t, per-layer delta }` for behaviours the client
  reproduces; per-paragraph/text-range targets resolved into layer ids
  server-side (text builds are a server expansion:
  `src/slideshow/player/text-builds.ts:57–106`;
  `render/scene.ts:37–45`'s per-target text metadata is server-held).
- `morph[]`: matched layer pairs with start/end plan geometry
  (`src/slideshow/morph/render.ts:31–65` consumes normalized frame
  geometry; the local match runs server-side).
- `transition[]`/`links[]`: kind, direction, duration, sound ref, and
  **link gating restored (round 8): internal jumps carry a target = a
  streamed slide index only; external URLs ride the jump record only
  when the manifest's `can.links` is true (default off), and a link can
  never carry data the page does not show.**
- **Admission limits [call] — one combined budget (round-8 finding-5
  restore):** per slide ≤ 128 layers, ≤ 64 steps, ≤ 512 keyframes
  (draft 6 allowed 1,024 — the stricter 512 stands, recorded in the
  reviews file); **per show ≤ 128 MiB aggregate decoded bytes covering
  layer assets AND show-plan records combined** (draft 7 split the two
  numbers; the combined limit is the norm). Pre-decode admission,
  `(session, slide, generation)` tagging, cancellation on slide exit and
  eviction follow §9.3's cache contract (restored, correctly
  cross-referenced); an over-limit slide surrenders animation — steps
  degrade to server-composed keyframes per §12a.3, never originals with
  content withheld.
- **Adapters (carried from round 6):** the client pipeline runs
  package-free (`src/slideshow/player/session.ts:228–232` today reopens
  the package); the "Next" preview composites planned layers
  (`src/slideshow/presenter/preview.ts:66,140–141,267–274`); the notes
  pane consumes server-drawn notes images
  (`presenter/PresenterView.tsx:52,315`); the no-package/no-`PptxShape`
  client prototype is a Stage-0/3c blocker (§16 **T5**; round-8
  cross-ref fix: it is a show-plan gate, not lifecycle).
- **Refused behaviours:** `cmd` (OLE) dropped with diagnostics; the rest
  plays.

### 12a.3 No hidden content in layers — eligibility on final observable appearance

A layer can hold content that never appears **observably** — including by
recolour, opacity or masking, not only geometry (verified:
`src/slideshow/render/canvas2d.ts:102–107,178–193,345–361,418`;
behaviour inputs ride separately in `src/slideshow/types.ts:118–173`). Rules:

1. Compute the layer's **observable appearance envelope** per build step
   and sampled motion time (rendered look with and without the layer,
   diffed: covering, cropping, alpha 0, recolour-to-background, full
   mask).
2. Anything never observable in appearance: **not sent**.
3. Partly observable: mask to the union and **re-raster at the
   observable appearance envelope** (a modified client cannot reveal
   more than the composed show showed).
4. **Proof subset (enumerated):** effects whose every observable state
   is determined by data the client already has — visibility toggles
   over safe rasters, transitions composited from safe rasters, and
   motion/Morph over safe rasters **only when either no lower layer had
   masked/withheld pixels inside the mover's swept-area trajectory, or
   the swept area is analytically proven at 100 % coverage for all
   `t ∈ [0, 1]`** (sampling alone is never the proof — the sample gap
   *is* the fallback trigger). Everything else plays as
   **server-composed keyframes** (`step-composited`), and **no original
   layers ship alongside fallback frames for the unsafe interval**.
5. Wire audit adds colour/opacity/mask/intermediate-time fixtures
   asserting transmitted decoded assets carry no recoverable
   unreleased ink (T5).

### 12a.4 Watermarks during a show (round-8 finding-7 fix)

Marks are composited into **every layer bitmap and keyframe** with
`globalCompositeOperation = 'source-atop'` (transparent padding stays
alpha-0; the `rasterLayer` layers are transparent RGBA boxes,
`src/slideshow/render/raster.ts:95`), and **no client-side duplicate
overlay** replays them. `source-atop` preserves alpha but does **not**
by itself prove that per-layer marks equal one composite-once mark:
overlapping translucent layers stack their marks. Therefore the approved
layer subset is validated against a **server-composited reference** under
the named tolerance gate published with §16 T7's §12a.4 mark-fidelity
battery (round-9 finding 7 names it here explicitly, replacing draft 6's
"unacceptably" wording), and any subset failing it (or whose
moving layer-local marks would stop being a page-anchored mark) plays as
**server-composed marked keyframes**. Measured in T5/T7.

### 12a.5 Speaker notes

Only a viewer with `role: 'presenter'` **and** `can.notes` gets notes;
Presenter View only; server-drawn marked images; notes
a11y/semantic/search answers re-check `can.notes` conjunctively after
resume (§6.2 rule 7); every delivery is audited. `can.textA11y` alone
never grants notes. Notes assets live in the §9.3
`notes:<presenterId>` namespace. **Limit:** a presenter can retype the
notes; marks trace screenshots, not retyping (§19.5).

### 12a.6 Media

Poster frames by default; optional Stage 3d transcodes (mark-burned,
metadata-free, via `src/slideshow/media/lease.ts`), never original bytes.

### 12a.7 Presentation behaviour (round-8 lost-rule-7 restore)

- Slide list follows the policy-filtered units (hidden slides not
  counted); `slideIndex`/`goTo()` count streamed slides only.
- **Resolution:** rendered at the audience screen's size × DPR rounded
  to the §7.3 scale bucket (part of the §9.3 cache key); a screen change
  or resize re-requests at the new bucket, showing the current bitmaps
  scaled meanwhile; **growth behaviours that enlarge a layer above its
  bucket request a larger bitmap for that layer, capped at `maxUnitPx`.**
- **Prefetch:** current slide + 2 next + 1 previous [call], under the §14
  unit rate limits (a show is not a bulk-export path); **`goTo()` fetches
  the target first**; a slide whose layers have not arrived shows its
  **composed final image** and starts animating when the layers land.
- **Gesture:** same synchronous start as local mode (§8.2) — the surface
  opens inside the gesture; content streams after.
- **Errors:** a slide that fails to stream shows a typed per-slide error
  frame **and the session stays alive**; revocation mid-show exits and
  `ended` resolves, and the show **never falls back to local rendering**.
- **Session lifetime:** token renew/revocation/expiry follow §10.

### 12a.8 Acceptance criteria (restored round 8; referenced by §17/T5)

- **3b (static show + notes):** present a streamed deck; forward, back,
  `goTo()`; hidden slides never appear and the counter excludes them;
  Presenter View shows notes **only** for a presenter viewer and never
  on the audience surface; a non-presenter gets no notes from any
  request; no PPTX, notes text outside the a11y service, or media bytes
  cross the wire; both marks on every slide and notes image; revocation
  mid-show exits cleanly.
- **3c (animation):** the corpus's animated decks play through the
  stream with pixel-diff against local playback at fixed timeline
  samples within the §7.4 tolerance; the §12a.3 wire audit is green; the
  `step-composited` fallback rate is recorded; build steps, transitions
  and Morph pairs match local mode's step count and order.
- **3d (media, optional):** transcoded playback carries the burned-in
  mark; no original media bytes on the wire.

## 13. Forensic marks

### 13.1 Two layers, one decision

Layer 1 (primary, §2 item 10) is a **keyed spread-spectrum pixel pattern
on every unit** (every page, layer bitmap, keyframe, notes image, print
page). Layer 2 is keyed glyph-advance displacement, an independent
second layer **only on command-projection units**. Both encode the same
per-session codeword. Marks off ⇒ unverifiable (§6.1).

### 13.2 Layer-1 algorithm (normative; round-7 redesign completed round 8, agy C4/codex-r3-1)

Payload: `codeword_40 = token_24 | salt_8 | CRC8_8` (the session token is
a 24-bit monotonic counter from the audit table, never reused while its
audit row lives; exhaustion refuses sessions rather than recycling).

**Keyed non-linear message map (owner decision 2026-11-08, Stage-0 review
finding):** the 40 bits that enter the code are **not** `codeword_40`
itself but `msg_40 = HMAC-SHA256(kMsg, keyGen | codeword_40)` truncated to
40 bits (`kMsg` a further HKDF child of §13.4), stored in the audit row and
recovered only by table lookup. Reason: BCH and CRC-8 are linear and the
salt was fixed, so the XOR of any odd number of colluders' messages was
itself an issued token's message, and a K = 5 sample median (per-chip
majority, +3/8 on the 5-way product) scored that **innocent** token as
high as the colluders (measured Z 9.19 vs 7.7–9.9). With the keyed map a
collusion lands on an issued message only by chance (≈ M / 2⁴⁰ per
collusion: 6 × 10⁻⁸ at M = 65,536, 1.5 × 10⁻⁵ at M = 2²⁴), and an audit
insert that collides with an existing row's `msg_40` re-draws the salt.
**Multi-hit refusal:** when two or more candidates pass `τ_abs`, the
decision is `collusion-suspected` — it names no single viewer and lists
the passing candidates for manual review; the `τ_margin` rule alone never
attributes in that case. T9 measures innocent-framing rate per collusion
attack (median K ∈ {3, 5}, K-average, differencing) against these rules.

- **Page-keyed tiled basis:** the carrier positions/sizes are
  **page-keyed**, not session-keyed: `kBasis = HKDF(kMaster, "basis" |
  docRef | unitId)` places `N_chips ≥ 512` chips on tiled macroblocks
  (≥ 2×2 device-pixel cells, constrained low-mid frequency band, sized
  for JPEG q70 + 50 % downscale). A **single `O(pixels + M × L)` decode**
  (one pooled projection per carrier plus per-candidate chip reads over
  the surviving `L` cells — the "whole detector is O(pixels)" wording of
  draft 7 is NOT claimed; per-candidate work is `O(L)`, and M is table
  size) yields per-cell soft LLRs; the session's codeword modulates
  chip **signs** through the `(N_chips, 40)` code, so every candidate's
  evidence is one pass away.
- **Soft-decision code:** a `(N_chips ≥ 512, 40)` ECC family (candidates:
  BCH / polar / short Reed-Muller; the concrete family and tables are
  Stage-0 deliverables) — **not** 3× repetition with a hard-decision
  40-bit Hamming-ball decode, whose per-candidate ball probability
  double-counted the CRC bits (see the decision rule and the r6
  arithmetic).
- **Detector gating (before any codeword match, complete decision):**
  decode is gated by a normalised correlation score — `Z = ⟨r, s⟩ /
  (σ_r √L)` per candidate — accepting only when `Z_best ≥ τ_abs` **and**
  `Z_best − Z_second ≥ τ_margin`, with τ calibrated by simulation over
  the **complete decision** (alignment, crop ≤ 25 %, scale search and
  all candidates included) so the measured false-attribution rate over
  the full audit-table size stays ≤ **10⁻⁶ per decision** (T9 batteries;
  the per-candidate binomial was never the bound).
- **Degenerate evidence handling (round-8 finding-6 fix):** flat or
  reference-identical evidence has `σ_r = 0`, which makes `Z` undefined —
  such a capture **abstains explicitly** ("no usable signal"), never
  scores; the normalisation defines `σ_r` over the surviving-chip support
  `L` only (cells zeroed by clamping or masking are excluded from both
  numerator and √L).
- **White-clipping handling:** office pages are predominantly `#FFFFFF`
  and clip positive mean-zero chips at 255, halving chip power — the
  embed uses a **white pedestal / content-adaptive masking** (per-chip
  pedestal toward 253 on flat regions) or zeroes the full-white cell;
  decided at Stage 0 per measurement and recorded with the codec
  parameters.
- **ΔE00 bound (Stage-0 output that blocks shipment):** the rendered-page
  diff must stay within a measured **ΔE00 ceiling** [call: target ≤ 3;
  mandatory Stage-0 output]; until it is set, no image-only unit ships a
  mark.
- **Survival gates:** JPEG q70, 50 % downscale, 25 % crop, at the §13.5
  table's rates.
- **Where it enters the pipeline:** composited into the final surface at
  the resident encode step (`final-raster`; show layers/keyframes; notes;
  print units; protected `culled-commands` content rasters under the
  M-rule). Sessions with forensic off skip it and are recorded
  untraceable.

### 13.3 Collusion (both layers; round-8 arithmetic fix, finding-6)

- **Two coordinated holders:** comparing two streams reveals which text
  positions changed (Layer 2) and, on flat regions, which pixel cells
  were perturbed (Layer 1 `I_1 − I_2` reveals carrier support) — what
  leaks is carrier **positions**, not the payload. Decoy jitter (Layer
  2) narrows the distinguishable set but remains a Stage-2 hypothesis to
  measure.
- **K-capture attacks (correct arithmetic, round-8 finding-6):** the old
  "attenuates ~1/K and rounds to 0" claim was wrong. Averaging K **independent ±2** chip values gives **RMS 2/√K** (exactly:
K=4 → RMS 1.0; K=5 → 0.894427; K=6 → 0.816497), and a plain sample
median over the ±2 perturbation signs keeps ±2 for **62.5 % (K=4, an
even sample where the median averages the middle two signs), 100 %
(K=5, odd) and 68.75 % (K=6)** of sign combinations (exhaustive
enumeration; stated explicitly with the even/odd distinction, round
9 finding 7). — `[2,2,2,2,−2]` has median +2. So averaging/
  medianing degrades, but does not remove, the mark; the real bound is a
  **Stage-0/2 measurement obligation**, not a claimed number: for each
  attack (K-average over K ∈ {2,4,6}, K-sample median for both even and odd K
  (K ∈ {4,5,6}), two-stream
  differencing) T9 reports the **measured per-attack recoverability**
  under the recorded ΔE00 ceiling and publishes it in
  `docs/validation/streaming.md`. Nothing claims "the attack fails".
- Attributions stay audit-table-only; measurement is joint per round.

### 13.4 Key law

`OFFICEVIEW_STREAM_KEY` (32 B, absent ⇒ fail-closed startup) → HKDF:
`kToken`, `kDoc`, `kCarrier`, `kPixel`, `kBasis`, `kJitter`, `kMsg` + a
key-generation byte per deployment era. `docRef` is
`HMAC(kDoc, sha256)` truncated; the plain hash never travels. `sid` is
random per session, never the viewer's identity; the audit table binds
`sid ⇄ viewerId` (when `authorize` supplied one). The run/audit
lifecycle (write-once audit, evict-on-expiry, 90-day default) lives
once in §9.3 — referenced here, not redefined.

### 13.5 Detection tool and gates (round-8 restored as a table)

`server/scripts/detect.ts`:

```
bun scripts/detect.ts --images page1.png […] --key-file ./ov-key.json \
  --doc <file.docx> [--page N] --ids ids.json   # audit-table export
```

Per image, deterministic; two decoders, one report:

1. **Reference frame** re-derived through the pool (§9.3 key; a cache miss
   renders rather than screenshotting).
2. **Pixel-pattern decoder (Layer 1, primary):** register the capture
   (similarity fit tolerant to ≤ 25 % crop), de-synchronise over the
   surviving area (scan the cell clock), correlate candidate chip
   sequences against the band-passed luminance residual, gate with the
   §13.2 z-score rule (σ_r handling included), apply §13.2's decision
   rule; per-page scores pool across the document set.
3. **Word-shift decoder (Layer 2, text-rich command units only):** coarse
   alignment, per-word normalised cross-correlation (with vs without
   ±δ); the local argmax gives the displacement sample; per-page codeword
   decode follows the §20 decision rule.
4. **Consensus + confidence:** the two decoders' per-page votes pool with
   **declared weights (recorded in the codec parameters)**; a
   below-threshold result reports "no token recoverable", never a guess;
   when the audit table resolves a token, the report carries the token's
   **key generation** (round-8 restored) and viewerId if resolvable.

**Survival gate table (measured targets; per-class degradation is
published, never silent):**

| Capture | Text-rich command units (Layer 2) | Image-only units (Layer 1) |
|---|---|---|
| default-zoom screenshot, single page | ≥ 95 % | ≥ 90 % |
| after JPEG q70 | ≥ 80 % | ≥ 75 % |
| after 50 % downscale | ≥ 80 % | ≥ 75 % |
| **after JPEG q70 followed by 50 % downscale (combined attack; pipeline order = JPEG then downscale)** | **≥ 75 %** | **≥ 75 %** |
| after 25 % crop (**25 % removed from each axis**, ≈ 44 % of the area; owner ruling 2026-11-08) | ≥ 70 % | ≥ 70 % |
| sparse-page spooling | cross-page reassembly, published rate | n/a |

The attribution calibration covers the **whole deployed decision**, not
per-layer alignment alone: per-page pooling, the **weighted consensus of
both decoders**, and the alignment/crop/scale search all sit inside the
same ≤ 10⁻⁶ measured false-attribution target (T9 asserts the composed
decision, not its parts).

The synthetic battery (T7/T9) covers **JPEG q95/90/80/70, upscale 2×,
downscale 50 % and crop 5 %** per class, with per-case published
outcomes.

## 14. Performance, memory and limits

- **Rates** [call]: per IP 30 new sessions/min (8 concurrent connects);
  per sid 120 WS messages/10s (burst 20/s), 60 unit requests/min, 10
  print jobs/min (2 concurrent renders), 30 sort/filter ops/min (1 in
  flight); process 256 sockets, 4,096 queued units.
- **Handshake windows [call]:** `ready window` 2 s from the MANIFEST
  (the one definition of the §6.2 rule-4 window; expiry CLOSEs with
  `ready-timeout`), **renew: a scheduling lead — the server sends RENEW
  ≥ 5 s before the token's own `exp`**; the old token remains valid until
  its own `exp` unless an ACK has already rotated it (§6.2 rule 6) and per-op expiry enforcement is unchanged (round-10 F4),
  and a reassembly window of 5 s per multi-fragment message (§7.1).
- **Sizes** [call]: `maxUnitPx 8192`; render deadline 2 s; outstanding
  units 8 (queue 32); scratch lease ≤ 2 final-raster units in flight
  (§9.3); assets/session aggregate 256 MiB streamed by hash; **show
  plans + layer assets ≤ 128 MiB aggregate decoded, combined** (§12a.2)
  — the aggregate is distinct from the per-message caps below.
  **Transport:** server native `ws` `maxPayload` on client→server
  messages = 1 MiB (the largest encoded fragment, §7.1).
- **Per-kind message admission — encoded AND decoded caps (round-10 F3;
  every kind, no implicit default):**

  | kind | encoded (`totalLen`) cap | decoded cap |
  |---|---|---|
  | MANIFEST | ≤ 256 KiB | ≤ 256 KiB |
  | ASSET_DEFINE | ≤ 1 MiB | ≤ 8 MiB |
  | UNIT_RASTER | ≤ 8 MiB | ≤ 64 MiB |
  | UNIT_COMMANDS | ≤ 8 MiB | ≤ 8 MiB |
  | XLSX_TILE | ≤ 8 MiB | ≤ 64 MiB |
  | SHOW_PLAN | ≤ 8 MiB | ≤ 32 MiB |
  | NOTES_RASTER | ≤ 8 MiB | ≤ 64 MiB |
  | PRINT_RASTER | ≤ 8 MiB | ≤ 64 MiB |
  | text control message | ≤ 64 KiB | ≤ 64 KiB |

  Admission is **streaming inflation** (round-10 F3): the compressed
  payload is inflated incrementally and the inflate/parse **aborts the
  moment decoded output exceeds the decoded cap** — a 1,052-byte gzip
  of 1 MiB of spaces never passes a 256 KiB manifest check (T6 bomb
  test; plus an oversized **valid**-JSON manifest, so the gate is the
  size bound, not a parse error). Raster kinds' decoded caps count the
  **image-decoder RGBA output**, not just gzip-expanded PNG bytes; the
  client admits `width × height × 4` from the PNG header before
  decoding. Encoded + decoded transient bytes are **reserved together
  before inflate** (`AllocationReservation`, the round-2 pattern), so a
  compliant implementation cannot eager-decode before discovering
  overage. **One aggregate reservation covers every in-flight receive,
  inflate, decode or parse job, fragmented or not** (manifest + assets +
  show plans + units + notes — not merely the eight unit requests),
  reserved before work starts and released on every exit path:
  aggregate transient ≤ 96 MiB per session [call]. The largest single
  admission (8 MiB encoded + 64 MiB decoded = 72 MiB) fits inside it;
  a job whose reservation does not fit waits or is refused, it is never
  over-allocated. **Unsolicited or unmatched `reqId`
  frames are dropped before any reassembly buffer is allocated.**
  Tests cover the compressed bomb, several concurrently admitted first
  fragments, and single-fragment compressed messages (§16 T6).
- **Retained ceilings (restored round 8):** document cache 1 GiB and
  render cache 512 MiB **per process, bounding the retained caches**
  (the 512 MiB render *reservation* bounds job admission; these two
  bound the retained tiers); eviction/recycling per §9.3/§9.2.
- Client: replay ≈ today's interactive paints; DOM pressure drops.
- **Determinism (restored statement):** same
  `docRef ⊕ layoutFingerprint ⊕ scaleBucket ⊕ policy/codec parameters`
  inputs ⇒ byte-identical recording. The recorder reads no **clock and
  no ambient randomness** — `Date.now()`, `Math.random()`, and any other
  unparameterised nondeterminism are absent from the render path; the
  optional visible-mark timestamp arrives as an injected paint
  parameter.

## 15. Accessibility

Server text service returns text for coordinates from the §7.3 second
pass, joined through the §11 rule — a11y text carries exactly what the
policy+join permit (possibly less than local), with **exact locale/BiDi
ordering** (`RECORD_TEXT` sources, not OCR). Visible-mark text never
appears in a11y/search/copy answers. `can.textA11y` never authorises
notes; `can.notes` re-evaluates per session and after resume. Focus and
keyboard surfaces unchanged; a disabled service states
**"text unavailable in this viewing mode"** once.

## 16. Testing and validation (suite IDs `T1…T12`; single home)

- **T1 contract:** corpus record→replay→pixel-diff zero local-backend
  diff at the **zoom × DPR grid {0.5, 0.75, 1, 1.25, 1.5, 2, 3} ×
  DPR{1, 2}**, including shadows, absolute/setTransform-headed painters,
  pixel-snap paths, transformed patterns and scratch-surface drawImage;
  fractional zoom proves 1:1 replay; per-script tolerance table;
  cross-backend published tolerances.
- **T2 disclosure (Stage-0 feasibility subset runs before any wire work; the full integration battery is the Stage-1 gate, matching §17):** wire-inspection of transmitted
  payloads; the **bounded truncated-manifest/hello allocation test**
  (truncated/malformed first message parses under its decoded cap — no
  first-look allocation); the **in-repo Playwright WS harness**
  (`server/test/app.ts`) asserting happy / error /
  expiry / **reconnect restores viewport and visible units**; the
  **cropped + duotoned transmitted-snapshot inspection** (every
  drawn-bitmap snapshot is decoded in the audit and inspected for crop /
  duotone correctness, not just byte-compared); fixture matrix: hidden
  sheet, **hidden rows/cols**, hidden painting/`w:vanish`, **hidden
  shapes (`cNvPr`/`docPr` through groups)**, notes, comments, covered
  cells, masked bands, unsampled intermediate-time intervals;
  per-endpoint (search/copy/a11y/filter values) partial-run and
  covered-run withholding; **visible-mark text absent from every
  semantic answer (end-to-end)**; presenter/non-presenter notes carve
  outs asserted separately; **local loads byte-for-byte unchanged when
  no stream source is supplied** (restored round 8, verified present in
  the body round 9).
- **T3 recorder census:** opcode subset closure over the corpus (per
  §7.5's restored state list); backend face-resolution measurements
  published.
- **T4 fonts handshake:** `requiredFaces` manifest probes; missing/mismatch
  and **missing/unverifiable-face** raster fallback **inside a valid READY**;
  absent READY at the ready window CLOSEs `ready-timeout` and streams
  nothing (round-9 finding-1); CJK/ligature/BiDi fixtures;
  scratch-rendered text case.
- **T5 show plans (full battery is the Stage-3c gate):** §12a.2 admission
  limits; the **mixed-page decoder pin** named by §20 (a mixed
  raster+command unit's command side must decode to the identical output
  the `final-raster` path paints and the M-rule pin resolves — the code
  path codex-r4 finding 2 requires); **no-package/no-`PptxShape` client
  prototype**
  (Stage-0/3c blocker — show-plan suite, not lifecycle); §12a.3
  appearance adversaries (colour/opacity/mask/Morph + intermediate
  times); remote presenter adapters; fallback-rate reporting; §12a.8
  acceptance list.
- **T6 session/isolation (adds the round-9 fragmentation battery):**
  first-fragment declares `totalLen`/`fragCount`; corrupted per-fragment
  CRC, missing final-fragment, over-cap fragment count or `totalLen`,
  and reassembly-window timeout each discard the buffer and refuse the
  unit (no partial validation); byte-vector tests for the §7.1 grammar
  (single fragment, every kind, contradictory final/count, altered
  flags/length, wrong magic, wrong `wholeCrc`, gzip space-bomb, several
  concurrent first fragments); **golden framing vectors** (codex r10,
  executed: ver 1, MANIFEST kind 1, reqId 7, decoded payload `{}`, gzip
  mtime 0 level 9) — decoders must accept each, and an encoder using the
  same gzip parameters must reproduce them; they pin the framing, not a
  canonical gzip encoding (ext01 is optional for one fragment and split
  points vary):
  ```text
  single, no ext01 (flags 5):
  4f56534401010700000000000000b948950d051a0000001f8b08000000000002ffabae050043bfa6a30200000043bfa6a3
  single, ext01 (flags 7, totalLen 26, fragCount 1):
  4f56534401010700000000000000b948950d071a0000001a0000000000000001001f8b08000000000002ffabae050043bfa6a30200000043bfa6a3
  two fragments (flags 6 then 5, payloadLen 10 + 16, totalLen 26):
  4f565344010107000000000000003bbafe1d060a0000001a0000000000000002001f8b08000000000002ff
  4f5653440101070000000100000018c5d5120510000000abae050043bfa6a30200000043bfa6a3
  ```
  transport caps per endpoint (server native `maxPayload`; browser
  receipt guard on `ArrayBuffer.byteLength`/`Blob.size` against the 1 MiB
  fragment cap, no pre-delivery cap — the F1 statement); handshake close-code matrix incl.
  version-intersection refusals; wrong-viewer resume, missing
  `resumeProof`, `renew` race and rotation windows; notes recheck after
  resume; client-relax refusal (`policy-refused`); paired relaxed/
  protected sessions over **both cache-fill orders** and both mark
  settings sharing zero entries; per-`contentPolicyId` processing;
  **guessed/reused notes asset ids resolve to nothing (authZ on
  ASset/hashes)**; fingerprint-change `layout-changed` termination;
  sort/filter target authorization, tautology-filter over
  author-hidden rows, hidden pivot/slicer refusal.
- **T7 marks (Stage-0 feasibility subset: codec parameters/ΔE00 measured before wire work; the full integration battery is the Stage-2 gate, matching §17):** ΔE00 ceiling measurement + codec
  parameters published; **per-page carrier-count publication per unit**
  (restored round-9 finding 6); the **mark-fidelity tolerance gate for
  the §12a.4 approved layer subset** (the named `source-atop` gate);
  pixel-mark injection/extraction on image-only units; survival
  batteries (q95/90/80/70, 2× upscale, 50 % downscale, **combined q70 +
  half-scale**, 5 % and 25 % crop); unmarked + flat-evidence abstain
  batteries (σ_r = 0); **K-average (2/4/6), even-K/odd-K median and
  2-stream differencing with measured per-attack recoverability**;
  text-layer round-trips incl. decoy distribution and sparse
  cross-page spooling.
- **T8 XLSX:** tile pipeline (no too-large-sheet notice path), tile p95
  ≤ 120 ms on a 500×100 sheet; sort/filter authorization incl.
  hostile/tautology filters; covered-cell and **off-viewport covered
  cell** filter values; per-session view-state key; **sort/filter
  row-order parity with local mode; selection parity** (restored
  round 8).
- **T9 detection (Stage-0 feasibility subset: the detector math/parameters measured before wire work; the full integration battery is the Stage-2 gate, matching §17):** §13.5 table gates incl.
  **joint z-score + margin gating and whole-search false-attribution
  ≤ 10⁻⁶ measured over the declared audit-table size** (the q95/90/80/
  70, upscale 2×, 5 % and 25 % crop battery); **all of it on top of a
  per-page carrier-count publication gate** (T7).
- **T10 lifecycle:** close-code matrix (`4xxx` surfacing), abort-honouring
  and abort-ignoring producers, drain-then-refund retirement,
  resume/supersede sockets, close-on-unmount / exactly-once
  source-replacement close / atomic discard, **audit retention + key
  rotation** (restored), print job audit.
- **T11 eager:** transitive-graph budget, `dist/core.js` purity,
  module-map identity — mandatory **before wire implementation**
  (restored ordering).
- **T12 perf probing:** 1,000 concurrent unit requests hold the §14
  caps — mandatory Stage-0-side gate (restored).

## 17. Staging

Each stage a PR; gates by §16 suite IDs; measured, not narrated.

| Stage | Scope | Done when |
|---|---|---|
| **0** | the **Stage-0 feasibility subsets** of T1–T9 as labelled in §16 (full T2/T7/T9 integration is gated at the stage each names) (including ΔE00 + coding parameters, §12a prototype, backend confirm, fonts probes, appearance-eligibility rate) **plus T11 (eager/core-entry gates) and T12 (resource probe) green before any wire implementation** (round-8 restore) | all gates green **with published numbers** in `docs/validation/streaming.md`; any gate failing reports to the owner rather than silently switching choices |
| **1** | core entry + eager gates (§8.1 as ruled in §2 item 14b); server hidden-content strip incl. the fail-closed row (§6.4); pinned font pack (§7.4 rule 6); recorder + replay; WS handshake + typed closes; RemoteOfficeDoc branch; **Stage-1-mandatory semantic occlusion evidence (§7.7/§11 — implemented in this stage; the command-culling store is the only Stage-2 deferral)**; demo WS-only server (pool, default-protected sessions, marks, per-session print wiring); local-vs-others byte-for-byte test | corpus zero local-backend diff; T2's local-vs-stream tests **including zero ink for all 13 Stage-0 hidden constructs on the stream path**; handshake close-code matrix; T11 gates green |
| **2** | marks + per-§13 policy, occlusion **store** (command-culling; the Stage-1 evidence remains mandatory), per-session scratch lease; audit table + retention/key rotation; detection tool; M-rule prototype (decided scope) **with its shipping gate** | T6/T7/T9 green incl. false-attribution ≤ 10⁻⁶ over the declared table; marks-off sessions audit-recorded; M-rule protected-command tests green before any `culled-commands` ships |
| **3** | §11 semantic services + print (§2 item 9); notes presenter surfaces | T2 notes variants; T8 partial; print audit |
| **3b/3c/3d** | stream presentation per §12a (static+notes; animation; optional media = §12a.6) | §12a.8 lists (restored round 8); T5 green; 3c further blocked on the T5 prototype gates |
| **4a/4b/4c** | XLSX tiles via the viewport pipeline; selection parity; server sort/filter with the C1 purge + authorization + hostile tests + rate class + pivot/slicer refusal strips | T8 green **including tile p95 ≤ 120 ms (500×100), sort/filter row-order parity with local mode, selection parity, and the off-viewport covered-cell withholding** (round-8 restore) |

## 18. Boundaries

**Always:** developer chooses mode by source type; once a session is open
its protections are server-side state (no client relaxation); fail-closed
on unknown op/version/opcode/frame/manifest; provenance on every artefact
(`docRef ⊕ contentPolicyId ⊕ fingerprint ⊕ scaleBucket ⊕ markSettings ⊕
projection ⊕ unitId`); wire audits over screenshot-only assertions;
honest docs language (screenshots are traceable, not preventable);
per-session policy binding incl. `can.notes` re-evaluation on resume.

**Ask first:** shipping font geometry for faces the document only
embedded (outline extraction — moot under the decided text mode, kept for
future layers); persisting sessions beyond process restart; audit-table
export beyond the operator; serving streams for documents that previously
came as bytes; retention beyond the owner's written policy.

**Never:** ship document bytes or any member/media/font bookkeeping bytes
over this protocol in any form (debug frames included; the `server/`
package has no file route and the library cannot stop a developer who
publishes the file elsewhere — §6.1/§19, not claimed away); ship
original encoded media/metadata-bearing replicas; answer a **protected**
session without the P-rule; **ship hidden content in any payload except
the single owner-granted exception — an audited `contentPolicy: 'full'`
session's `authorize` relaxation (§6.1/§2 item 5); the T2 hidden-content
suite asserts every protected session**; log `OFFICEVIEW_STREAM_KEY`,
tokens, sids or audit rows; pretend DRM; execute client input inside the
render pipeline.

## 19. Honest limits

1. **Screenshots/cameras:** traceability, not prevention; marks-off
   sessions are untraceable by construction (§6.1).
2. **Deployment boundary:** the library cannot stop a developer who also
   publishes the file elsewhere; `server/` has no route (§6.1).
3. **ODT/RTF refusal is a deliberate scope limit** until hidden-text
   filters exist (§6.2).
4. **Collapsed row/col extents reveal a hidden band exists** (§12);
   accepted, with a "no numbering" alternative named.
5. **Presenters can retype notes**; marks trace screenshots (§12a.5).
6. **The pixel-parity story is deployment-specific** (backend + font set +
   bucket); T1 measures what the deployment pins; §7.6's tile path
   guarantees a wrong pixel never paints as right.
7. **Pixel-mark collusion has an honest, measured bound** (§13.3):
   K-average/median attacks degrade but do not eliminate the mark on flat
   regions and 2-stream differencing reveals carrier support; the claim
   is the measured per-attack recoverability, not "no attack".
8. **Detection on flat/blank evidence is not promised:** σ_r = 0 capture
   abstain explicitly (§13.2); whole-search attribution bounds are
   simulation-verified over the declared table (§13.5), never analytic
   promises.

## 20. Appendix: command-projection extension (Stage 2+)

The default projection is `final-raster`. The command-projection extension
ships later, stage-gated, for deployments that want smaller payloads:


- **`culled-commands` (protected, opt-in):** per-command occlusion proof
  (recorded with the record, never the recorder's opinion); units failing
  proof downgrade to `final-raster`. The **M-rule** requires **every
  content-bearing component** of the unit to be a protected transmitted
  raster; units whose readable content rides commands answer
  `final-raster` (round-7 → round-9 fix: the mixed page's shipping gate is
  §17 Stage 2's M-rule gate, not T8's XLSX suite; a mixed-page decoder
  pin lives in T5).
- **`commands` (relaxed sessions):** ordinary stream; visible mark as a
  last-position `RASTER_BLOCK` overlay, removable by a modified replayer
  — stated plainly, because a relaxed session trusts the replayer.
- **Layer-2 text displacement (normative here, complete round-8
  restoration):** candidate words per page seeded
  `HMAC(kCarrier, docRef | unitId)` (length ≥ 4 grapheme clusters, body
  baseline, non-whitespace, ≥ 3 word boundaries apart); **classification
  carrier-rich (≥ 40 eligible words) / sparse (16–39) / none (< 16)** —
  a carrier-rich page hosts the whole 40-bit codeword alone; sparse
  pages spool across siblings with distance-k interleaving (detector
  reassembles and votes); `none` pages carry no text layer (Layer 1
  still marks them). Placements are `±δ` on a **1/256-px lattice**
  (δ < 0.5 CSS px [call]). Decoy jitter `±δ/2` on non-carriers is a
  **Stage-2-measured hypothesis**, never claimed undetectable.
  Decision rule: the observed bit string vs each candidate's expected
  codeword — CRC-8 checked over the **observed 32-bit payload** (sample
  integrity), Hamming distance ≤ **6 of 40** [call], unique-winner-or-
  abstain (runner-up within 2 ⇒ abstain), and a whole-search
  **≤ 10⁻⁶ false-attribution target** calibrated by simulation over the
  declared audit-table size (with unmarked-capture abstain batteries;
  the per-candidate binomial is raw material, never the bound).
- Text-layer capacity from §13.2 applies (token allocation, salt, page
  classification counting round-4 thresholds verbatim).

## 21. Review history

All round dispositions (codex r1–r6, claude r1, agy r1) live in the
sibling file `docs/superpowers/specs/2026-11-07-server-streaming-reviews.md`,
including the round-8 table (17 restored rules + findings 1–8).
