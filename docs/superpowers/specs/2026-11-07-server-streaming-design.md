# Server-streamed Office rendering: design

Status: **draft 7 (reviews: codex r1–r4, agy r1, claude r1)** — drafts 2–6 folded
codex r1–r4, claude r1 and the agy sweep (row tables live in the sibling file
`2026-11-07-server-streaming-reviews.md`); draft 7 folds the agy review and
consolidates the document. **Every owner decision is settled** (§2); branches and
"textModes" still marked open in older drafts no longer exist.
Branch: `spec/server-streaming`, from main `ca952c50`. Docs-only stage: no source
edits.

Notation: **"to verify in Stage 0"** marks a claim to be proven by Stage 0's
prototypes before production code lands. Code facts cite the worktree's HEAD;
anchors follow agy's re-verified table (§reviews).

## Table of contents

1. Intent
2. Owner decisions (all decided; 2026-11-07/08)
3. Goals, non-goals, anti-goals
4. Current state (recon, main `ca952c50`; anchors re-verified)
5. Architecture overview
6. Stream source, session policy and the handshake (normative)
7. Draw-command format (`OVSD`), recorder, and the projection rule (normative)
8. The client side
9. The server side
10. Behaviour: open, failure, resume, expiry
11. Search, copy, and semantic text services
12. XLSX in stream mode
12a. Presentation mode for streamed decks
13. Forensic marks
14. Performance, memory and limits
15. Accessibility
16. Testing and validation (suite IDs `T1…T12`; the single home of test gates)
17. Staging
18. Boundaries
19. Honest limits
20. Appendix: command-projection extension (Stage 2+)
21. Review history pointer

## 1. Intent

officeview renders Office documents to canvas in the browser. For a confidential
document that means the bytes sit in the browser cache; every client-side
"protection" (`src/core/watermark.ts:6–13`) is documented as *a deterrent, not
DRM*.

This design adds a second mode with a different trust model: a server parses,
lays out, and paints the Office document, and **streams the client only what is
drawn** — Canvas 2D draw commands or rasters, plus the minimum interaction
geometry. The original file never crosses the wire in this mode; search, copy
and semantic text are server calls. The mode a viewer runs in is **the
developer's choice, made by source type** (§6.1); the library does not decide it
and nothing can be inferred by inspecting a response.

## 2. Owner decisions (all decided — 2026-11-07/08)

These are binding. Superseded branches (outlines, HTTP/`OVSM`/406/Accept
negotiation, per-document "stream-required" marking, §18-as-separate-section)
were **deleted** in the round-7 consolidation; the rules below are the single
normative home of everything they once covered.

1. **Mode = source type (2026-11-08, replaces the round-1 item-1 rule).**
   "Client/Server does not need to determine the doc should be stream or not.
   That is developer's decision. If a doc is supplied to the viewer, it goes file
   view. And if a stream, websocket, is supplied, go stream mode." The component
   takes `ByteSource | StreamSource` (§6.1); the branch is at the `OfficeFile`
   layer; nothing is sniffed. There is no per-document marking; **the
   developer's backend is the boundary** (§6.1, §19).
2. **Same repo, open source.** The library stays at the root; the server is a
   `server/` Bun workspace package depending on a React-free `officeview/core`
   entry (§8.1). Library-side dependencies stay zero; server-side deps are Stage
   0 choices (originally: outlines-mode; **decided 2026-11-08** per item 6 — no
   outlines dependency at all).
3. **Routing (2026-11-08).** The component prop is **`data`** (not `source`; the
   real prop at `src/components/OfficeFile.tsx:154` is `data: ByteSource`, and
   `<OfficeFile>` declares it at `:219`); a stream source enters via the widened
   `data` union and is guarded against `HttpSource` structural collision (§6.3).
   `<OfficeDoc>` renders a remote document through a discriminated arm of
   `OfficeDocSource` dispatched inside `OfficeDoc` (§6.3). The record point stays
   the per-unit `page.paint(ctx)` closure inside `getPaintables`
   (`src/render/paint.ts:354`, closures at `:389`) — the server always paints
   through it.
4. **Server surface (2026-11-07/08).** HTTP/WebSocket service, short-lived
   signed tokens, rate limits, render pool, caching, per-session marks
   (switchable, see 8), session/audit tables, detection tool. **The `server/`
   package has no file-download route**; it serves streams only.
5. **Content policy and projection.** Every stream session is **protected by
   default**: final-raster projection (§7.7), pre-layout content policy (§6.4),
   marks on. The developer's server-side `authorize` hook may relax projection,
   content policy, or marks per session (audited, §6.1) — never the P-rule in a
   way that breaks final visibility (§7.7) and never from the client.
6. **Text mode (2026-11-08).** Plain text + pinned server faces only (§7.4); no
   shaping/outlining dependency; missing faces ⇒ raster units. Applies to command
   projections; `final-raster` ships pixels.
7. **XLSX depth (2026-11-07).** Server-side sort/filter ships as Stage 4c with
   per-session view state, display-text-only filter lists, no cell data in
   replies, its own rate-limit class and a wire audit; interactive tiles render
   through the paged viewport pipeline (§12), never `getPaintables`.
8. **Presentation (2026-11-07, revised same day).** Full presentation mode with
   animation and Morph (§12a); speaker notes only for a presenter-role viewer,
   Presenter View only, as marked images (§12a.5); media poster-frames by
   default with optional transcoded Stage 3d.
9. **Print (2026-11-08).** A per-session setting defaulting on (§18 old item 6):
   `can.print` is set by `authorize`, the pool re-renders print units at print
   resolution with the session's marks, requests carry their own rate class, and
   each job is audited.
10. **Forensic mark (2026-11-08).** The keyed **spread-spectrum pixel mark** is
    the primary mark **on every unit** (§13.2); keyed text displacement stays as
    a second, independent layer on command-projection units. Rejected: hidden
    QR codes (public reveal, single croppable region, JPEG-fragile).
11. **Visible mark (2026-11-08).** Per-session: fixed text (default, e.g.
    `CONFIDENTIAL`), or viewer identity fields, or off. Both marks off ⇒ an
    **untraceable** session, recorded in its audit row (§6.1, §19.2).
12. **A11y (2026-11-07).** `can.textA11y` defaults on; a deployment turning it
    off ships the §15 notice. It alone never authorises notes (§12a.5).
13. **Canvas backend (2026-11-07).** Skia via `@napi-rs/canvas` for the server
    pool; Stage 0 confirms rather than chooses (§9.2). Every canvas-construction
    site in the render pipeline goes through one **`CanvasBackend`** seam (§9.2).

History (3 lines): drafts 1–4 routed stream sessions through an HTTP-manifest
sniff with an `Accept`/406 negotiation, then through a per-document
"stream-required" flag, before the 2026-11-08 decision moved the mode choice to
the source itself and mandated default protection; an outlines text mode
existed above item 6's plain-text decision until the same date. The full round
dispositions live in the reviews file (§21).

## 3. Goals, non-goals, anti-goals

(Full rules are normative in §6/§7; this section only names the goals.)

Goals: (1) stream mode exists and is safe by construction (protected by
default, §6.1); (2) same pixels on both routes — a unit replayed equals the
same-backend direct render, with one fidelity gate per text mode (§7.4/§16);
(3) a stream session never leaks original media bytes, hidden content (§6.4),
or covered content (§7.7) — and the developer's backend keeps the file itself
off every surface including `server/` (§6.1); (4) both marks behave per §13 and
are switchable only by the developer's server; (5) search/copy/semantic text
are server calls over the visible-projection join (§11); (6) XLSX streams tiles
through the viewport pipeline (§12); (7) sessions are governed (§6.2, §9.3).

Non-goals: true DRM (§19.1); a general display protocol (only document ink
travels); rewriting the local renderer; an officeview-branded product/server;
client-side content inspection as a trust boundary (the server enforces, the
client just paints).

Anti-goals (failures even if they "work"): covered or hidden content in any
transmitted payload (the §16 tests inspect **transmitted payloads**, not
screenshots); a stream server that also serves document bytes anywhere; a
client-side relaxation surface; a replayer that silently mispaints (unknown ops
⇒ tile, §7.6); treating client-submitted text as code (§18 Never).

## 4. Current state (recon, main `ca952c50`; anchors re-verified against this worktree)

- **Byte sources.** `loadOfficeFile(data: ByteSource, …)`
  (`src/components/OfficeFile.tsx:39`; the prop is `data`, `OfficeFileProps` at
  `:152–156`, `<OfficeFile>` at `:219`) → `loadOfficeSource(source: ByteSource,
  …)` (`src/components/loadOfficeSource.ts:202`); local `AnyDoc` union is
  `DocxDocument | PptxDocument | XlsxDocument` (`:32`). The loader also
  normalises **ODT (`parseOdt`) and RTF (`parseRtf`) into `DocxDocument`**
  (`:67–68`, `:117`) — the server's format scope addresses this in §6.2.
- **Policy headers (local mode only).** `X-OfficeView-Protection: no-copy,
  no-print` (`src/core/stream.ts:51–75`) and `X-OfficeView-Watermark` (parser
  `src/core/watermark.ts:164`) still travel on local byte downloads; there is no
  HTTP response in stream mode, so stream policy is `manifest.can.*` as
  server-enforced (§6.2).
- **Paint surface.** `getPaintables` (`src/render/paint.ts:354`) returns
  per-unit `paint(ctx)` closures (`:389`) — the single record point. Public
  direct painters exist (`src/index.ts:49–56`); the server uses `getPaintables`
  only (§7.3). `measurerFromDoc` (`:184–211`) measures headless but
  **hard-codes `import('canvas')`** — as do **five** server-relevant sites
  total: `paint.ts:~202`, `src/core/search.ts:~212`, `src/core/images.ts:65,
  :149`, `src/slideshow/render/svg-raster.ts:64` (all written as
  `import(/* @vite-ignore */ 'canvas')`, which is why a plain `grep
  "import('canvas')"` misses them; verified). `ctx.canvas.constructor` surface
  cloning appears at `src/core/text-metrics.ts:31,42`,
  `src/drawing/paint.ts:126`, `src/drawing/pattern.ts:155` — all become
  `CanvasBackend` seams (§9.2, agy I3.4).
- **Device-space dependence (recorded, round-2's basis for scale buckets).**
  `snapToPixels` (`src/drawing/scene-paint.ts:32–41`), CTM-scaled shadows
  (`src/drawing/paint.ts:363–397`), `pattern.setTransform` in the radial brush
  (`:178–196`), sparkline scale from the transform (`src/xlsx/render.ts:774–775`),
  charts keyed by DPR (`src/charts/paint.ts:767–768`,
  `src/charts/layout.ts:253–263`), text laid out during paint
  (`src/drawing/text-paint.ts:972–989` + side canvases in
  `src/core/text-metrics.ts:23–42`), `'letterSpacing' in ctx`
  (`text-paint.ts:113`), scratch surfaces drawn back under skew
  (`text-paint.ts:712–765`).
- **A text-only recording context exists.** `createRecordingContext`
  (`src/core/search.ts:92`) captures `fillText`/`strokeText`;
  `buildTextIndex` (`:315`) drives a **separate** pass with the watermark muted
  (`:322–334`, module-global mute at `src/core/watermark.ts:244–249`). Painters
  branch on `RECORD_TEXT` (`src/core/text-recording.ts:2`; e.g.
  `src/pptx/render.ts:81` skips master underlay) — so the visual recorder never
  exposes `RECORD_TEXT` (§7.3).
- **XLSX disclosure surface.** The interactive viewer ships formulas
  (`AuthoritativeCell`, `src/xlsx/view/contracts.ts:74–80`), hidden rows/cols
  (`SheetLayout`, `:668–674`), and raw picture bytes
  (`contracts.ts:362–366`, `XlsxImage data: Uint8Array` at
  `src/xlsx/types.ts:27`, via an `op:'drawings'` read,
  `src/worker/viewerRemoteDrawings.ts:1–18`). Sheets parse with **no hidden
  state** (`src/xlsx/parse.ts:290–301` ignores `state`) and `getPaintables`
  paints every sheet (`paint.ts:427–429`). Sort/filter kinds are client-typed
  (`contracts.ts:709–710`; command list `src/worker/viewerProtocol.ts:31`, which
  also lists `pivot*`/`control*`). **A filter can reveal author-hidden rows:**
  `isVisibleSource` (`src/xlsx/view/visibility.ts:149–155`) lets a supplied
  filter override `baseHidden` inside its range —
  `src/xlsx/view/viewport.ts:176–179` documents "a filter may reveal a
  snapshot-hidden body row" (agy-confirmed; the local intent, and the reason
  stream-mode 4c must purge, §12).
- **Drawings ignore `hidden`.** `sourceOf` reads `cNvPr` id/name only
  (`src/pptx/parse.ts:714–717`; same pattern in `src/xlsx/drawing.ts:559ff` and
  `src/docx/parse.ts:580,620`) — `p:cNvPr`/`xdr:cNvPr`/`wp:docPr`
  `hidden="1"` is unread, and RTF's `applyControlWord`
  (`src/rtf/parse.ts:1130ff`) has no `\v` case; ODT
  (`src/odt/parse.ts`, `src/odf/styles.ts`) does not filter
  `text:display="none"` / `<text:hidden-*>` / annotations. §6.4 closes these.
- **Eager budget.** Ceiling 38,720 B (`scripts/eager-budget.ts:14`), last
  measured 38,582 B (`docs/validation/chart-compat.md:25`), re-measured fresh
  per stage. `eager-budget.ts:16` currently gzips **only `dist/index.js`** and
  ignores its static-import graph, while `vite.config.ts` sets
  `preserveEntrySignatures: 'allow-extension'` (`:59`) — the §8.1 gate
  therefore measures the **transitive graph** and asserts `dist/core.js` has no
  React/document imports (round-7 fix, agy I3.3).
- **Contract tests.** `src/test/pixel-diff.ts` (RGBA + `diffBitmaps` +
  `renderPaintables` at `:98`), `scripts/golden.ts` and
  `scripts/golden-corpus.ts` are the corpus contract loop (§16).

## 5. Architecture overview

```
+--------------------------------------+        +----------------------------------------+
| HOST (embedding officeview)          |        | SERVER ("server/" Bun workspace)       |
|                                      |        |                                        |
|  <OfficeFile data={source}>          |        |  WS /stream (the only endpoint)        |
|      │  file source → local view     |        |      │ origin allowlist + authorize    |
|      │  (unchanged local path)       |        |      ▼ content policy (§6.4, pre-layout)|
|      ▼  stream source → stream session        |      ▼ parse+layout per (docRef,       |
|  <OfficeDoc document={RemoteDoc}>    |        |        contentPolicy) (§9.3)           |
|      │  per unit → replay/raster at  |        |  pool: CanvasBackend render workers    |
|      │   the recorded scale (§7)     |        |  recorder over page.paint (§7.3)       |
|  search/copy → WS op (§11)           |        |  projection (§7.7) + marks (§13)       |
|  a11y/print/present → remote impls   |        |  caches, session/audit, detection tool |
+--------------------------------------+        +----------------------------------------+
```

- **Client** (`officeview` at the root): the React-free core entry
  (`officeview/core`, §8.1) exports parse/layout/paint/fonts/stream plumbing;
  the stream loader, format code and remote document source live in a lazy
  `src/stream/`; the server depends on `officeview/core`, never React.
- **Server** (`server/`, §9.1): CanvasBackend + pool + recorder + projection +
  policy gate + caches + marks. **No file route, no HTTP document endpoint.**
- **Wire**: one WebSocket per stream session. A JSON-text control frame
  multiplexing namespace with `reqId`s (§6.2) and one little-endian binary
  frame header for all payload kinds (§7.1).

## 6. Stream source, session policy and the handshake (normative)

### 6.1 The stream source and per-session policy

```ts
export interface StreamSource {
  kind: 'stream'
  /** A WebSocket endpoint run by the developer's backend. `wss://` expected;
   * the origin must be in the deployment's allowlist (§9.3). */
  url: string
  /** Optional preauthorization proof for the open decision (§6.2). Never put in
   * a URL path; unused after HELLO (it is NOT the session token). */
  token?: string
}
```

- **The mode is the branch, not an inference.** `<OfficeFile data={…}>` accepts
  `ByteSource | StreamSource | Promise<ByteSource | StreamSource>`; a file
  source runs today's local path untouched; a stream source opens a stream
  session. Neither branch reads the other's content. `HttpSource`
  (`src/core/stream.ts:42–46: { url, headers?, credentials? }`) is
  **structurally collidable** with a StreamSource, so `HttpSource` gains
  `kind?: never` (round 7, agy I3.1) and `readSource`/`openByteStream` throw a
  typed `StreamSourceMisuseError` if handed one — a spread `wss://` url must
  never reach `fetch`.
- **Per-session policy (the single normative home; §3/§19 summarize):** every
  stream session is **protected by default** — `final-raster` (§7.7), content
  policy `visible-only` (§6.4), visible mark fixed-text on, forensic pixel mark
  on (§13). The developer's server-side `authorize(request)` hook decides per
  session:
  ```ts
  authorize(req) → { viewerId?: string; role?: 'presenter' | 'audience';
                     policy: { contentPolicy?: 'visible-only' | 'full';
                               projection?: 'protected' | 'relaxed';
                               visibleMark?: 'fixed' | 'viewer' | 'off';
                               forensicMark?: boolean;      // default true
                               print?: boolean } }          // default true, §2 item 9
  ```
  Relaxations are recorded in the audit table (§13.4); **the client has no
  relax surface** — a client-requested relax is refused (§16 T6). Turning both
  marks off makes the session **untraceable** and the audit row says so (§19.2).
- **Who keeps the file safe is the developer's backend.** `server/` has no
  file-download route and no HTTP document endpoint; the library cannot stop a
  developer from exposing the file elsewhere (§19.2, the owner words in §2 item 1).

### 6.2 The session handshake, one state machine (normative)

Terminology: a **message** is one protocol element (it may ride several
WebSocket frames); a **frame** is the wire boundary. Text messages are JSON
objects with `{ id, type, … }`; binary payload kinds share one little-endian
frame header with a `reqId` (§7.1), so a control reply can reference its
request.

```
states: HELLO → MANIFEST → READY(ack) → STREAM → (RENEW → STREAM)* → CLOSE(code)
        HELLO{resume} → (re)MANIFEST → READY(ack) → STREAM | CLOSE
```

1. **HELLO (client → server, first message; bounded `4 KiB`):** `{ client:
   'officeview/1.0', recVersions: number[] (a set), token?: string,
   resume?: { sid, resumeProof } }`.
2. **Version compatibility is set intersection** — the server picks a version
   it supports **from the client's set**; the empty intersection closes with
   `unsupported-version` (never a downgraded guess).
3. **MANIFEST (server → client):** the `OVSM` envelope (§6.2 framing below)
   carrying the normative manifest JSON:
   ```json
   { v: 1,
     docRef: "<HMAC(kDoc, sha256(bytes)) truncated 96-bit, hex>",
     kind: 'docx'|'xlsx'|'pptx', name: "<label>",
     units: { count, list: [ { w, h } ] },
     layoutFingerprint: "<hash over (officeview version, backend version,
        font set, locale, parser inputs)>",          // §9.3 canonical key input
     session: { sid, token, expiresSec, resume: 'session' },   // token ISSUED here (I1 fix)
     policy: { contentPolicy, protection, semantic, contentPolicyId,
               viewerMarkId },                     // §6.1/§6.4; cache-key inputs
     can: { search, copy, print, textA11y, present, notes, links },
     recorder: { commands: 'ovsd-1', textMode: 'text',
                 scalesBuckets: [1, 2],
                 projection: 'final-raster' | 'culled-commands' | 'commands' },
     limits: { maxUnitPx: 8192, maxAssetBytes, wsFrameBytes, tilePx },
     notices: ['...'] }
   ```
4. **READY (client → server): exactly one client readiness message**, merging
   the old `ack-session` and `faces-ack`:
   `{ type: 'ready', tokenJti, missingFaces?: id[] }` (I1 fix). FACES-ACK is
   not a separate first/second frame debate any more; the §7.4 face-registry
   contract decides `missingFaces` and missing faces ⇒ raster units.
5. **STREAM:** normal ops (units/viewport/tiles/show plans/notes/copy/search).
6. **RENEW (server → client) before expiry:** the server pushes a signed
   replacement `session.token` (only the server can sign; the round-5
   "both-sides" race wording was impossible and is deleted) and the client
   answers `renew-ack { jti }`; on silent expiry the connection ends with
   CLOSE `session-expired`.
7. **RESUME (HELLO with `resume`):** bound **to the same viewer and policy**
   (verified against the session table, not the sid alone): the table stores
   `{ viewerId, role, effectivePolicy, activeSocket, jti }` (I1 fix); resume
   closes any prior `activeSocket` atomically (WS close `4409 superseded`),
   **rotates `jti`**, issues a fresh `session.token`, and re-evaluates
   `can.notes` for the resumed viewer (§12a.5). A wrong-viewer proof, a
   proof/policy mismatch or a missing `resumeProof` closes with
   `auth-failed` (never a silent switch).
8. **CLOSE codes (typed, all fail-closed):** `unsupported-version`,
   `auth-failed`, `session-expired`, `revoked`, `rate-limited`,
   `policy-refused` (client-relax attempts land here), `rotation` (see 7).
   The client maps each to a `StreamSessionError` surfaced as a load-error
   state (§10).

**OVSM frame envelope (still binary framing, no `+json`, never
transport-detection):** `magic 'OVSM', u16 LE version, u16 LE compression
(0 raw / 1 gzip), u32 LE decoded JSON length, JSON`.

Validation (client, fail-closed): numbers range-checked; unit sizes 1..8192
(oversized units are tiled server-side, §14); enums closed; token ≤ 512 chars,
never in a URL path, bound at HELLO; unknown fields ignored.

**Format scope (round-7 fix, agy C2/c):** the stream server admits **OOXML
only** — `doc.format: 'docx' | 'xlsx' | 'pptx'`. `.odt` / `.rtf` are refused
fail-closed at admission (`unsupported-format`) until their hidden-text
filters exist locally (RTF `\v` is unread today,
`src/rtf/parse.ts:1130ff`; ODF `text:display="none"` /
`<text:hidden-*>` / annotations are unfilted) — a deliberate scope limit, not
an oversight, recorded honestly in §19.3.

### 6.3 Routing a stream source (round-7 fix, agy I3.1/I3-5)

Verified anchors: `src/components/OfficeFile.tsx:39`
(`loadOfficeFile(data: ByteSource, …)`), `:152–156` (the `data` prop),
`:219` (`<OfficeFile>`), `:88ff` (`useOfficeFile`), `:92` (its route key),
`:127/:140` (`pagedWorkbookOf(…)?.release()` disposal pattern);
`src/components/loadOfficeSource.ts:32` (local `AnyDoc`),
`:202` (the loader); `src/components/OfficeDoc.tsx:48–51`
(`OfficeDocSource`), `:344/:672` (`getPaintables` calls),
`:947–962` (the XLSX structural dispatch).

- **`OfficeFile` layer:** the `data` union widens to
  `ByteSource | StreamSource | Promise<ByteSource | StreamSource>`. A branch
  checks `'kind' in src && src.kind === 'stream'` **before** any byte path;
  stream sources never call `loadOfficeSource`. The branch result state is
  `OfficeFileState { status: 'ready', document: RemoteOfficeDoc }`; disposal
  copies the `pagedWorkbookOf(...)?.release()` lifecycle.
- **`OfficeDoc` layer (round-6 correction stands, hardened):**
  `OfficeDocSource` gains a fourth arm `{ kind: 'remote'; session:
  RemoteOfficeDoc }`; the dispatch at `OfficeDoc.tsx:947–962` becomes
  remote → `src/stream/remote/` lazily, local-sheet, legacy. Local
  search/present/a11y code never inspects a remote doc (the union branch
  precedes them); `present()` reads the manifest synchronously (§8.2).
- **Eager gate:** the §8.1 gate measures the **transitive static-import graph
  of `dist/index.js`** (the same closure `tests/helpers/eager-graph.ts`
  computes), not just its own bytes, and a separate assert pins
  `dist/core.js` to zero React/document imports while keeping
  `preserveEntrySignatures` behaviour safe (§8.1.

### 6.4 The server content policy — pre-layout, per `contentPolicyId` (normative)

"Stream what is drawn" is bounded by "drawn **for this viewer**". The gate
runs **before layout**, keyed by a `contentPolicyId` recorded in the §9.3
canonical key. **Processing is per `contentPolicyId` over an immutable parsed
source** (a relaxed session never mutates the parse another session sees;
§16 isolation tests).

| Content (as parsed, worktree-verified) | Local mode | Stream rule |
|---|---|---|
| `state="hidden"/"veryHidden"` XLSX sheets (unread by `src/xlsx/parse.ts:290–301`) | painted as extra pages (`paint.ts:427–429`) | not streamed; excluded from units/sheet-nav/all text services; an authorize relaxation may stream them per audited session (default off) |
| **Author-hidden rows/cols** (`row@hidden`, `col@hidden`) | painted collapsed (authored geometry) | **purged from the server's projected sheet snapshot before any `SheetView`/`FilterEngine` is built on protected sessions** (round-7 fix, agy C1) — or, where a sheet declares `<autoFilter>` and the `spec.range` matches that block, carried as a separate immutable `authorHidden` tier that no filter may flip; either way `isVisibleSource` never sees an overrideable author-hidden tier |
| w:vanish runs (unread) | painted | excluded from paint/index/copy/a11y; per-document notice; T2 assert |
| PPTX hidden slides (`show="0"`) | painted as units | not streamed as units |
| **Hidden shapes** — `p:cNvPr`, `xdr:cNvPr`, `wp:docPr` `hidden="1"`, propagated through `grpSp` ancestors (unread by `src/pptx/parse.ts:714–717`, `src/xlsx/drawing.ts:559ff`, `src/docx/parse.ts:580,620`) | painted (Selection Pane hides would leak) | excluded from paint, index, copy, a11y, and show plans (round-7 fix, agy C2b); T2 fixture class |
| Speaker notes / comments | notes unread; comments unread | comments never streamed; notes only as §12a.5 marked images for a presenter-role viewer |
| Hidden rows/cols **shape** (collapsed extents) | painted collapsed | axis sizes + collapsed gaps only; numbering reveals a band's existence (§19.4) — the deployment accepts this or drops streamed axis numbering |
| Chart text (`src/charts/describe.ts:40–62` reads model data) | model-based text | visible-projection only: painted names/labels/categories, never model value ranges (`policy.semantic`) |
| Pivot tables / slicers / timelines (`src/worker/viewerProtocol.ts:31` lists `pivot*`/`control*`) | interactive, pre-cache | **refused on protected 4c streams** unless/until a scoped design lands; pivot caches and slicer items backed by hidden/veryHidden sheets or hidden rows/cols are **stripped before serving** (round-7 fix, agy C1c) |
| ODT / RTF inputs | parsed into DocxDocument locally (`loadOfficeSource.ts:67–68`) | **refused fail-closed at stream admission** (`unsupported-format`, §6.2) until hidden-text filters exist — deliberate scope limit (§19.3) |

The client never participates; it renders what arrives.

## 7. Draw-command format (`OVSD`), recorder, and the projection rule (normative)

### 7.1 One binary frame header for every payload kind (round-7 fix, agy I5)

All binary WS payloads share one little-endian header; the §7.1 body follows it:
`magic u32 ('OVSD') | frameSeq u32 | crc32 u32 (of payload) | kind u8 | reqId
u32 | flags u8 | payloadLen u32 | payload`. Kinds: `ASSET_DEFINE`,
`UNIT_RASTER`, `UNIT_COMMANDS`, `XLSX_TILE`, `SHOW_PLAN` (`OVSP` §12a.2),
`NOTES_RASTER`, `PRINT_RASTER`. `reqId` matches the JSON control frame that
requested the payload. **The old whole-document `unitIndex[]` offset table is
gone** — units stream on demand in any order at any scale bucket.

```
per unit: u32 LE unitId + u32 LE byteLen + u32 LE opCount +
          u8 per-unit discriminator (commands | raster | tile) +
          gzip-compressed body   // never a monolithic upfront index
```

A `TILE` op carries its bitmap **inline in the frame** (bounded by
`wsFrameBytes`), not as a data URI.

### 7.2 Why commands (short)

Command streams compress well, are resolution-independent *given the recorded
scale bucket* (§7.3), and make per-session forensic edits a mutate-pass on
cached records (§13). Resolution dependence is why the pipeline names its
scale everywhere.

### 7.3 The recorder (normative contract) + `CanvasBackend`

- **Facade over a real backing 2-D context** (the generalisation of
  `src/core/search.ts:92`), so `'letterSpacing' in ctx` truthiness
  (`src/drawing/text-paint.ts:113`), `measureText` during paint
  (`text-paint.ts:972–989`, side canvases in `src/core/text-metrics.ts:23–42`),
  scratch-surface `drawImage` (`text-paint.ts:712–765` — snapshotted by content
  hash at draw time), wrapper patterns with recorded `setTransform`
  (`src/drawing/paint.ts:178–196`) all behave as on a real canvas.
- **Recorded subset is typed and Stage-0-census-final** (§16 T3): transforms,
  state incl. dash/alpha/composite/smoothing/letterSpacing, gradients
  (`createPattern` wrappers) and shadows, path ops, `fill`/`stroke`/
  `fillRect`/`strokeRect`/`clip`, text (`font/textAlign/textBaseline`,
  `fillText` with `maxWidth`, `strokeText`), `drawImage` (all 9-arg forms),
  and the raster-inlining ops of §7.7. Unknown ops ⇒ promote that unit to
  `tile` (never a silent skip).
- **`RECORD_TEXT` is never exposed by the facade** (painters branch on it,
  `src/core/text-recording.ts:2`, e.g. `src/pptx/render.ts:81`) — the text
  index stays its own second pass through `createRecordingContext` +
  `buildTextIndex` (`src/core/search.ts:92,315`) with the watermark muted.
- **Scale discipline:** record at the session's scale bucket
  (`scalesBuckets`, §6.2) and replay 1:1 at it; device-space decisions above
  (`snapToPixels`, shadow CTMs, sparkline scale, charts' DPR) are captured as
  recorded values. Client-side zoom re-requests the bucket (§8.3).
- **`CanvasBackend` (round-7 fix, agy I3.4): one seam for every
  canvas-construction site the recorder pool uses:**
  `CanvasBackend { createCanvas(w, h); decodeImage(bytes); loadImage(bytes);
  rasteriseSvg(bytes) }` threaded into `paint.ts` (`measurerFromDoc`),
  `search.ts`, `images.ts:65/:149`, `svg-raster.ts:64`, plus the four
  `canvas.constructor` clone sites (`src/core/text-metrics.ts:31,42`,
  `src/drawing/paint.ts:126`, `src/drawing/pattern.ts:155`) — `@napi-rs/canvas`
  is a `createCanvas` factory whose `CanvasElement` is not
  `new`-constructible, and today those sites `catch → undefined`, silently
  dropping gradients/warp/patterns on the server (all five
  `import('canvas')` sites are verified, hidden behind
  `import(/* @vite-ignore */ 'canvas')`).
- **Render deadline = worker retirement** (drain-then-refund): a paint cannot
  be aborted mid-`paint()`; a 2 s unit deadline [call] terminates the worker,
  charges refund only after termination, the unit retries once, then tiles.

### 7.4 Text mode — one mode, decided (plain text + pinned faces)

Glyph runs travel as strings + advances; replay uses server-pinned faces
whose hashes the client confirmed:

1. **Identity list:** `requiredFaces: { id, hash (SHA-256 of face bytes,
   96-bit truncated), alias? }` — hashes only; face bytes never ship (§18).
2. **Bounded acknowledgement:** part of the **READY** message (`missingFaces`),
   one-time per session; absence of the message means `[]` (fail-closed).
3. **Match:** hash equality only. A client that cannot hash the pinned bytes
   (an installed system face) reports it as **missing** — an assertion without
   a registry is not an assertion (round-7 restatement of round-6's rule).
4. **Missing/mismatch:** the unit goes to raster (`face-missing` diagnostic);
   substituted faces are never painted, and the zero-diff fidelity claim is
   voided only for such units.
- **Fidelity gates per decided mode:** replay fidelity = recorder→replayer
  equality at the recorded scale (same backend, same pinned faces) is the
  zero-diff claim; *renderer fidelity to the local hinted painter* is a
  published tolerance per script family (CJK/Ligature/BiDi fixtures live in
  T4).

### 7.5 Asset table and op list

Asset table keyed by content hash: drawn-bitmap snapshots (§7.7), decoded
scratch surfaces, pattern/gradient wrappers (`setTransform` included),
pinned-font **hashes** (never bytes). Ops (normative, ~24 as per the round-2
census): STATE_PUSH/POP; TRANSFORM (setTransform/f64 pairs); STYLE_COLOR;
STYLE_OBJ (gradient/pattern); STATE_SETUP (dense dword: width/join/cap/
miter/dash/dashOffset/alpha/composite/smoothing/letterSpacing/direction/
lang/shadow); PATH_* and PATH_END {FILL,STROKE,FILL_STROKE,CLIP};
GLYPH_RUN (mode: f64 advances + string; shapes are text mode only — the
outlines `GLYPH_RUN_OUTLINES` opcode of older drafts is deleted); IMAGE_DRAW
(asset ref + transform + optional source rect); TILE (inline raster);
RASTER_BLOCK (final-frame block of a `final-raster` unit or the visible-mark
overlay on relaxed sessions).

### 7.6 Versioning, tiles, runtime fallback

Negotiation is the §6.2 set intersection; a frame version outside the agreed
set fails closed. Units the recorder cannot express stream as tiles with a
per-unit discriminator (§7.1); mixing command and tile units is expected;
replayed-frame validation failure drops the unit and re-requests once.

### 7.7 The projection rule (P-rule) — normative, one home

> **P-rule.** Every response of a stream session contains only content that
> belongs to the unit's **final composited image** — no covered text, no
> clipped-out media, no full z-order history. Protected sessions satisfy it
> **by default and without configuration**; the only departure is a
> server-side `authorize` relaxation to `commands` for a specific audited
> session (§6.1), and the client is not a party to it.

- **`final-raster` (default):** paint whole units to the final frame server-side
  and ship rasters; occlusion is structural.
- **`culled-commands` (protected, opt-in; details in appendix §20):** ships
  commands after a stored per-command occlusion proof; M-rule constraints in
  §9.4 else downgrade.
- **`commands` (relaxed sessions):** ordinary command stream; an overlay-mark
  is removable by a modified replayer — stated plainly, because a relaxed
  session already trusts the replayer (§13.2/§20).
- **Stage-1-mandatory semantic occlusion (round-7 fix, agy C3):** because
  default sessions are `final-raster` and the culling proof is deferred
  (§20), the §11 semantic join cannot depend on it. During the same visual
  paint pass that records the unit, the server keeps (a) per-text-run painted
  **ink bounds** and (b) a conservative **opaque-coverage grid** (alpha≥0.95
  extents) after each retained draw op; a text run is servable iff **100 % of
  its recorded ink bounds survive to the final raster** (per-cell, or
  per-sample at ≥ 2× the run height's resolution). Anything unresolved —
  including text whose localized region sits inside a graphics state the
  census did not model — is **withheld** (fail-closed), not shipped as
  trusted pixels (§11). This check runs on `final-raster` without any culling
  proof and is implemented once against the `getPaintables` paint surface.
- **No original media bytes (every projection):** `drawImage` ships the
  **drawn-bitmap snapshot** (as-painted, OCR-recap free), source-rect
  cropped (`src/drawing/scene-paint.ts:132–150` crops with `drawImage` source
  rects), re-encoded metadata-free at the recorded scale, keyed by pixel
  hash; other scales re-request.
- **Wire audits inspect transmitted payloads**, not screenshots (§16 T6).

## 8. The client side

### 8.1 The React-free core entry and honest eager gates

- Build `officeview/core` (`src/core-entry.ts`, name [call]) in a dedicated
  Vite **separate pass** with React externalised. Because
  `vite.config.ts:59` is `preserveEntrySignatures: 'allow-extension'` and
  `scripts/eager-budget.ts:16` gzips only `dist/index.js`, the eager gate is
  defined honestly (round-7 fix, agy I3.3):
  1. **Transitive-graph budget:** sum gzipped bytes over `dist/index.js`'s
     static-import closure (the same set `tests/helpers/eager-graph.ts`
     computes) — ceiling unchanged, 38,720 B.
  2. **`dist/core.js` purity assert:** zero `react`/`document` imports in its
     graph; `officeview/core` never imports React beyond erased types.
  3. **Main-bundle module-map check:** `dist/index.js`'s module map must be
     **identical** to the pre-branch build (± the branch's own discriminant);
     eager growth ≤ 16 B gzip recorded as a diff.

### 8.2 Component surfaces

- `<OfficeFile data={ByteSource | StreamSource | Promise<…>}>` (real prop
  name `data`, `src/components/OfficeFile.tsx:154`): the `HttpSource.kind?:
  never` guard plus the `readSource`/`openByteStream` typed misuse errors
  (§6.1) make a stream/wss source incapable of reaching `fetch`.
- `streamClient?: (source: StreamSource) => StreamClient` — WS factory hook
  (mirrors `viewerWorker`, `OfficeFile.tsx:63–80`); default is `WebSocket` +
  hello + heartbeat. `onStreamNotice(diag)` — deduplicated load notices.
- **`startPresent` stays synchronous** (round-7 fix, agy I3.2): on a remote
  document the eager `present.ts:184–246` path reads the manifest — slide
  count, `can.present`, `can.notes` — synchronously, refuses
  `presenterView: true` when `can.notes` is false (so a non-presenter never
  gets a blank `window.open`), and only then dispatches to the lazy remote
  show runner. All `present()` semantics (`presenterView`, gesture
  ordering) stay in the local-path shape.

### 8.3 The remote document provider (normative shape)

```ts
RemoteOfficeDoc = {
  kind: 'remote', manifest,
  paint(unit, scale: 'fit'|number, dpr, canvas): Promise<void>, // replay or tile paste
  units(): Promise<UnitSpec[]>,
  search(query): Promise<HighlightRect[]>; copy(range): Promise<string>;   // §11
  close(): void,
}
```

Dispatch at `OfficeDoc.tsx:947–962` (remote → `src/stream/remote/` lazily, then
sheet, then legacy) and state wiring via `OfficeFileState` (§6.2/§10). Replay
is 1:1 at the recorded bucket; zoom re-requests. The client ignores the host
watermark prop for a remote doc (the baked mark is authoritative).

## 9. The server side

### 9.1 Package layout [call]

```
server/  (bun workspace; deps: officeview workspace + @napi-rs/canvas)
  src/
    ws.ts             the only endpoint: upgrade (origin allowlist + authorize),
                      hello → manifest state machine (§6.2), op routing,
                      multiplexed text/binary frames (§7.1), per-op authZ
    authorize.ts      authorize(req) → {viewerId, role, policy…} (§6.1)
    tokens.ts         HS256 claims schema, issuance/rotation, jti table
    sessionTable.ts   sid → {docRef, viewerId, role, effectivePolicy,
                      activeSocket, jti, startedAt, status}   // I1 fix
    audit.ts          sid → {viewerId?, ip, startedAt, keyGeneration,
                      policyId, markSettings} (attribution; §13.4)
    pool.ts           CanvasBackend render pool + deadline retirement (§9.2)
    cache.ts          one canonical cache (§9.3)
    record.ts         recorder facade + scratch snapshot service (§7.3)
    projection.ts     §7.7 raster/projection + snapshot hashing
    occlusion.ts      §7.7 Stage-1 occlusion + §12 XLSX analytic occlusion
    policy.ts         §6.4 pre-layout gate
    forensic.ts       §13; detection.ts (CLI scripts/detect.ts)
    xlsxTiles.ts      §12 viewport tiles (renderViewport path, agy I4.1)
    print.ts          §2 item 9 print jobs (per-session, audited)
    config.ts         keys/flags, fail-closed
  test/               (bun test; §16)
```

### 9.2 Render pool and `CanvasBackend`

- Pool shape mirrors `WorkerRenderEnv`
  (`src/worker/office-worker.ts:42–47: createCanvas/toBitmap/parse/paintables`)
  with `createCanvas` bound to the chosen `@napi-rs/canvas` backend through
  the **one `CanvasBackend` seam** (covering all five `import('canvas')`
  sites and all four `canvas.constructor` clone sites verified in §4).
- Defaults [call]: 1 render worker per 2 vCPU (cap 8), 1 parse worker,
  parse-once-per-(docRef, contentPolicy, fingerprint), global render
  reservation 512 MiB per process, per-job decoded budget reserved up front,
  2 s deadline → drain-then-refund retirement, per-worker font budget 64 MiB.
- Pass interleaving: `setWatermarkMuted` is a module global
  (`src/core/watermark.ts:244–249`) until context-scoped, so a worker runs a
  paint pass **or** an index pass, never interleaved.
- Per-session outstanding units: 8 (queue 32); rate-limited by §14.

### 9.3 One canonical cache; session/audit storage (normative; round-7 fix agy I2)

**Cache key (single form everywhere):**

```
docRef ⊕ contentPolicyId ⊕ layoutFingerprint ⊕ recorderVersion ⊕
scaleBucket ⊕ markSettings ⊕ projectionMode ⊕ (policyDetails: per-session
view state hash for XLSX 4c) ⊕ unitId (subkey)
```

- Canonical entries hold the **pre-forensic unit surface** (commands or
  rendered lossless-raster/tiled-RGBA surface) + shared asset table + culling
  proof — **shared across sessions**, never per-sid pixels.
- **Per-session mark+encode** happens in a bounded per-session **scratch
  lease**: ≤ 2 final-raster units concurrently per session, explicitly
  admitted against the §9.2 512 MiB process reservation before allocation,
  streamed then released — never a per-sid mutated copy retained in the
  canonical tier, and never 8 × 64 MiB of unbounded surface. The old
  "8 MiB session edit tier" is renamed accordingly (I2 fix). Only sessions
  with forensic+visible marks off share final encoded bytes directly.
- **Document cache** (parsed doc + projected visibility + text index) is
  keyed **by `(docRef, contentPolicyId, layoutFingerprint)`** — a relaxed
  session's full-policy projection and a default session's visible-only one
  are different entries; nothing leaks across them (fixes the old single-
  per-docRef wording, which "leaked full content from relaxed to visible").'
- Text index entries are keyed the same way (they are join outputs of the
  §6.4 policy projection).
- Session table: `sid → { docRef, viewerId, role, effectivePolicy,
  activeSocket, jti, startedAt, status }` (used by resume, §6.2 rule 7).
- Audit table (survives render sessions): `sid → { viewerId?, ip, startedAt,
  keyGeneration, policyId, markSettings }`; 90-day default retention [call];
  never shipped, predicted or logged (§18).

### 9.4 Marks: visible + forensic, per session (normative; round-7 fix agy I2.1)

- **Visible mark** is composited into the final raster server-side
  (`paintWatermark` machinery, `src/core/watermark.ts`), fixed-text default
  or viewer identity fields, or off (§2 item 11). Per-session mark
  settings are part of `markSettings` in the canonical key.
- **Forensic layer-1 pixel mark** (§13.2) is applied to every raster per
  session (per-sid perturbation of the *page-keyed* basis) during the
  session's scratch-lease encode; it **cannot live in the shared canonical
  tier** and the "one cache entry serves every session" claim of older
  drafts is **retracted** — the pre-forensic surface is shared, the marked
  encode is per-session and streamed.
- Fleet-wide: fixed marks + same policy ⇒ one cache entry for the
  pre-forensic surface; every encode still goes through the per-session
  scratch lease.
- Ordering (one pipeline): paint → (protected) projection/occlusion pass →
  Layer-2 text shift only on command-unit records with glyph coordinates →
  visible + Layer-1 mark composite → flatten/encode.

## 10. Behaviour: open, failure, resume, expiry

**(Open)** developer passes source → §6.3 branch: file source keeps today's
loader; stream source opens WS (`streamClient`), HELLO (§6.2), origin
allowlist + `authorize` hook, `manifest`, READY (`missingFaces`), unit/notes/
print requests (§2 item 9 print). Pool renders → projection (§7.7) + marks
→ frames; client replays 1:1 at the bucket.

**(Failure/refusal/expiry)** typed `StreamSessionError` codes; a HANDSHAKE
malformed hello fails closed with no allocation beyond a bounded parse; a
token aging out mid-flow ends with CLOSE `session-expired` unless a `renew`
lands first (§6.2 rule 6); `revoked` closes once, buffers drop atomically; a
presentation exits and `ended` resolves; recovery is a developer-driven
re-open (§2 item 1). Replayer-side unit validation failure drops the
unit, re-requests once, then shows the missing-unit diagnostic.

**(Isolation)** per-session caches/keys/marks (§9.3), retiring sockets on
resume, and per-`contentPolicyId` processing (§6.4) are all covered by §16
pairs.

## 11. Search, copy, and semantic text services (normative)

- **Join rule (single home):** every semantic answer runs a conservative
  visible-projection join before release. Stage-1 implementation (round-7
  fix, agy C3): the server records each `RECORD_TEXT` run's paint-surface
  **ink bounds plus a conservative opaque-coverage grid** (collected during
  the same visual paint pass, §7.7 Stage-1 occlusion paragraph). Text is
  servable iff (a) its range sits fully inside the painted clip band,
  (b) 100 % of its ink bounds survive the opaque-coverage grid to the final
  raster (or per-§6.4 a drawing-free patched view), and (c) it is not part
  of another surface's private text. On **any unresolved or not-computed
  case** (including `final-raster` where the join composes structurally but
  the *logical range* is not exactly recoverable from pixels), **withhold** —
  never approximate. Per-unit diagnostic counting covers withheld ranges.
- **Search** (`can.search`): server op against the same (joint-projection)
  index built in the §7.3 second pass with the visible mark muted
  (`src/core/search.ts:92,315`, mute at `:322–334`); answers are
  `HighlightRect[]` (`src/core/overlay.ts:7`).
- **Copy** (`can.copy`): display text of the selected range; XLSX serves
  formatted display text for selected cells, never values/formulas.
- **A11y/semantic** (`can.textA11y`): §15, from the same index, subject to
  §6.4. `can.textA11y` never authorises notes (§12a.5).
- **Chart text:** visible-projection only (painted names//categories), no
  model value ranges.
- **Notes text** never appears in search/copy/a11y answers **unless**
  `can.notes` **and** presenter proof are both true (§12a.5).

## 12. XLSX in stream mode

- **Tiles via the viewport pipeline (round-7 fix, agy I4.1):** the server
  pool opens the sheet through `openImportSession`/`SheetGeometry` and
  paints tiles with **`renderSheetViewport`** (`src/xlsx/view/renderViewport.ts`)
  — the paged/resident pipeline — not `getPaintables` (whose XLSX path is a
  static full-sheet export with `StaticExportLimitError` at
  `src/render/paint.ts:175`, no viewport, no tile rect). A tile is keyed by
  the §9.3 canonical key + `sheetId + tileRect` and ships `XLSX_TILE`.
- **Disclosure boundary (unchanged from earlier rounds + C1 additions):**
  no formulas (`AuthoritativeCell`, `src/xlsx/view/contracts.ts:74–80`), no
  hidden rows/cols metadata (`SheetLayout`, `:668–674`), no drawing bytes
  (`SheetDrawings.images`, `:362–366`); paint-facing layout only; the
  collapsed row/col numbering limitation persists (§19.4). Hidden sheets are
  gated pre-layout (§6.4).
- **Staged interactions:** 4a static + read-only tiles; 4b
  selection/freeze/navigation parity; 4c server-side sort/filter (owner
  decision, §2 item 7) — **with the round-7 C1 purge** (§6.4: protected
  sessions never let `isVisibleSource` flip author-hidden rows; the
  `authorHidden`/`autoFilterHidden` split applies, `row@hidden` data is
  purged or masked), **server-side target authorization** (sort keys/filter
  ranges validated against the policy-approved view before engine-run,
  refused before order/counts/tiles can move), and **filter values through
  the §11 join** (values come from display text of cells that pass the
  full-visibility join; covered/occluded cells don't leak). Pivot/slicer/
  timeline commands are **refused on protected 4c streams** (§6.4), and
  PivotCache/Slicer items backed by hidden/veryHidden data are stripped.
- **Rate classes:** sort/filter ops and print sessions each carry their own
  §14 class.
- **Wire audit** covers the §6.4/§12 rows (T8).

## 12a. Presentation mode for streamed decks

`present()` works on a remote document with **animations, transitions and
Morph**, played back from server-drawn layer bitmaps + a bounded show plan
(`OVSP`, §12a.2). The file never reaches the client.

### 12a.1 The split (verified anchors)

Server: parse, content policy (§6.4), planner, layer raster
(`src/slideshow/render/plan.ts:17–29` `PlanLayer` (with `PptxShape[]`
**retained server-side only**), `raster.ts:41` `rasterLayer`, Morph match,
appearance projection §12a.3, marks). Client: scheduler, behaviours, easing,
masks, transitions, compositor, input, presenter surfaces — the same engine
code the local player uses, minus any shape/package read.

### 12a.2 The show plan (`OVSP`), bounded schema (round-7, agy I6)

One per slide, versioned with the §7.1 frame header:
- `layers[]`: `id, assetId, bounds, z, partTag` — resolved base geometry and
  per-part splits (`fill/line/text`), no shape data.
- `steps[]`: order, appearance-action eligibility flags (§12a.3), trigger
  kinds with **hit-region** rectangles only.
- `keyframes[]`: `{ t, per-layer delta }` for behaviours the client
  reproduces; per-paragraph/text-range targets already resolved into layer
  ids server-side (text builds are a server expansion:
  `src/slideshow/player/text-builds.ts:57–106` becomes a server pass;
  `render/scene.ts:37–45`'s per-target text metadata is server-held).
- `morph[]`: matched layer pairs with start/end plan geometry
  (`src/slideshow/morph/render.ts:31–65` consumes normalized frame
  geometry server-side; the local match runs on the server).
- `transition[]`/`links[]`: kind, direction, duration, sound ref, jumps.
- **Admission limits** [call]: per slide ≤ 128 layers, ≤ 64 steps,
  ≤ 512 keyframes; per show, aggregate decoded layer bytes ≤ 128 MiB;
  every record goes through §9.2 pre-admission on `(session, slide,
  generation)`; eviction follows §9.3.
- **Adapters (round-6-carry, round-7 restated):** the client's
  session pipeline runs package-free (`src/slideshow/player/session.ts:228–232`
  today reopens the package — the remote adapter takes the show plan instead);
  the "Next" preview composites planned layers
  (`src/slideshow/presenter/preview.ts:66,140–141,267–274` today paints from
  `PptxSlide`); the notes pane consumes server-drawn notes images
  (`src/slideshow/presenter/PresenterView.tsx:52,315` today eats parsed text);
  no-package/no-`PptxShape` client prototype is a Stage-0/3c blocker (§16
  T10). `render/scene.ts:37–45`'s per-target text metadata is server-held.
- **Refused behaviours:** `cmd` (OLE) dropped with diagnostics; the rest plays.

### 12a.3 No hidden content in layers — eligibility on final observable appearance

A layer can hold content that never appears **observably** — including by
recolour, opacity or masking, not only geometry (verified:
`src/slideshow/render/canvas2d.ts:102–107,178–193,345–361,418` recolour/opacity
apply over retained textures at draw time; behaviour inputs ride separately in
`src/slideshow/types.ts:118–173`). Rules:

1. Compute the layer's **observable appearance envelope** per build step
   and sampled motion time (rendered look **with and without** the layer,
   diffed: covering, cropping, alpha 0, recolour-to-background, full mask).
2. Anything never observable in appearance: **not sent**.
3. Partly observable: mask to the union and **re-raster at the observable
   appearance envelope** (the raster cannot contain ink that a modified
   client could reveal by skipping the effect).
4. **Proof subset (enumerated):** effects whose every observable state is
   determined by data the client already has — visibility toggles over
   safe rasters, transitions composited from safe rasters, motion/Morph
   of safe rasters **only when either no lower layer had masked/withheld
   pixels inside the mover's swept-area trajectory, or the swept area is
   analytically proven at 100 % coverage for all `t`** (round-7 fix, agy
   I4.3: sampling alone is never the proof — the sample gap *is* the
   fallback trigger). Everything else — colour/opacity animations whose
   appearance envelope isn't baked, mask parameterisations the sampler
   can't bound, blends depending on content below — plays as
   **server-composed keyframes** (`step-composited`), and **no original
   layers ship alongside fallback frames for the unsafe interval**.
5. Wire audit adds colour/opacity/mask/intermediate-time fixtures asserting
   the transmitted decoded assets carry no recoverable unreleased ink
   (T5).

### 12a.4 Watermarks during a show

Visible and forensic marks are composited into **every layer bitmap and
keyframe** exactly once: composited with
`globalCompositeOperation = 'source-atop'` so transparent padding stays
alpha-0 (the `rasterLayer` layers are transparent RGBA boxes,
`src/slideshow/render/raster.ts:95`) and the mark's opacity does **not**
stack across overlapping layers — where translucent layers overlap
unacceptably, the keyframe path is used instead. **No client-side second
overlay** of the same mark on already-marked content (round-7 fix, agy I4.2).
Presenter notes images carry both marks (§12a.5).

### 12a.5 Speaker notes

Only a viewer with `role: 'presenter'` **and** `can.notes` (both set by
`authorize`) gets notes; Presenter View only; server-drawn marked images;
the notes a11y/semantic/search answers re-check `can.notes` conjunctively
after resume (§6.2 rule 7); every delivery is audited. `can.textA11y`
alone never grants notes. **Limit:** a presenter can retype the notes;
marks trace screenshots, not retyping (§19.5).

### 12a.6 Media

Poster frames by default; optional Stage 3d transcodes (mark-burned,
metadata-free, via `src/slideshow/media/lease.ts`), never original bytes.

### 12a.7 Other presentation behaviour

Slide list follows the policy-filtered units (hidden slides not counted);
resolution is the audience bucket; prefetch (current + 2 next + 1 previous)
under normal rate limits; gesture-safe start (§8.2's synchronous
manifest reads); session lifetime per §10 (revocation exits the show).

## 13. Forensic marks

### 13.1 Two layers, one decision

Layer 1 (primary, §2 item 10) is a **keyed spread-spectrum pixel
pattern on every unit** (every page, layer bitmap, keyframe, notes image,
print page). Layer 2 is keyed glyph-advance displacement, an independent
second layer **only on command-projection units**. Both encode the same
per-session codeword. Marks off per session ⇒ unverifiable (§6.1).

### 13.2 Layer-1 algorithm (normative; round-7 redesign, agy C4)

Payload: `codeword_40 = token_24 | salt_8 | CRC8_8` (the session token is a
24-bit monotonic counter from the audit table, never reused; exhaustion
refuses sessions).

- **Page-keyed tiled basis (fixes the `O(M × pixels)` carrier contradiction):**
  the carrier positions/sizes are **page-keyed**, not session-keyed:
  `kBasis = HKDF(kMaster, "basis" | docRef | unitId)` places `N_chips` chips
  in tiles glued on tiled macroblocks (≥ 2×2 device-pixel cells in a
  constrained low-mid band, sized for JPEG q70 + 50 % downscale — chip power
  per cell published by Stage 0 measurements). A **single `O(pixels)`
  projection per candidate-*independent* carrier** yields per-cell soft
  LLRs; the session's codeword modulates those chips by a signed
  `(N_chips, 40)` code — so sessions on the same page differ only in
  coded chip *signs*, and every candidate's evidence is decodable in one
  pass.
- **Soft-decision code:** a `(N_chips ≥ 512, 40)` ECC family (candidate
  specifications: BCH / polar / short Reed-Muller; concrete choice and
  code tables are a Stage-0 deliverable) — **not** `3× repetition` + a
  40-bit Hamming-ball decode, which produced the double-counted bound
  below.
- **Detector gating (before any codeword match):** `Z = ⟨r, s⟩ / (σ_r √L)`
  per candidate; reject unless `Z_best ≥ τ_abs` **and**
  `Z_best − Z_second ≥ τ_margin`; τ values calibrated so the **empirical
  false-attribution rate over the full audit-table size stays ≤ 10⁻⁶ per
  decision**, measured with unmarked-capture batteries (T9) — the old
  per-candidate binomial printout was not a bound (it double-counted the
  same 8 CRC parity bits already inside the Hamming ball).
- **White-clipping handling:** uniform `#FFFFFF` is ~80–95 % of an office
  page and clips positive mean-zero chips at 255 — the mark embeds with a
  small **white pedestal / content-adaptive masking** (per-chip pedestal
  toward 253 on flat regions) or the chip gains zero for the full-white
  cell, decided at Stage 0 per measurement and recorded with the codec.
- **ΔE00 bound:** the rendered-page diff must stay within a measured
  **ΔE00 ceiling** [call: target ≤ 3, real number = a mandatory Stage-0
  output]; until it is set, no image-only unit ships a mark.
- **Survival gates:** JPEG q70, 50 % downscale, 25 % crop at §13.6 rates.
- **Where it enters the pipeline:** composited into the final surface at
  its resident encode step (final-raster; show layers/keyframes; notes;
  print units; protected `culled-commands` content rasters under the
  M-rule). Sessions with forensic off skip it and are recorded as
  untraceable.

**Layer-2 text displacement** (unchanged from the round-4 design, now
appendix §20): coin flips per eligible word; 1/256-px lattice; decision rule
in the appendix.

### 13.3 Collusion (both layers; round-7 fix, agy C4.4)

- **Two coordinated holders:** comparing two streams reveals which text
  positions changed (Layer 2) and, on flat regions, which pixel cells were
  perturbed (Layer 1 `I_1 − I_2` differencing reveals support) — this
  bounds what carrier *positions* leak, not the payload. Decoy jitter
  (Layer 2) narrows the distinguishable set but remains a Stage-2
  hypothesis to measure.
- **K-capture averaging/median (Layer 1):** averaging or pixelwise
  medianing `K = 4–6` session rasters attenuates chip amplitude by ~1/K;
  with 8-bit sRGB rounding that rounds `±2/K < 0.5` LSB to 0 and collapses
  the mark on flat regions — stated as a measured bound, with the attack
  verified at `ΔE00 ≤ 3` in T9 rather than claimed absent.
- Everything else from the round-2/3 collusion prose is unchanged
  (audit-table-only attribution; joint high-accuracy measurement per
  round).

### 13.4 Key law

`OFFICEVIEW_STREAM_KEY` (32 B, absent ⇒ fail-closed startup) → HKDF:
`kToken, kDoc, kCarrier, kPixel, kJitter` + a key-generation byte per era.
Session `docRef` is `HMAC(kDoc, sha256)` truncated; plain file hash never on
the wire. `sid` is random, never the viewer's identity; the audit table maps
`sid ⇄ viewerId` (when `authorize` supplied one). Audit and session tables
are distinct only in retention: audit rows survive process restarts per the
retention policy (default 90 days [call]).

### 13.5 Detection tool and gates

`server/scripts/detect.ts` re-derives the reference frame (same pool), aligns
(pixel + word co-registration) and runs §13.2's gated decode over the audit
table's candidate list; output either token + viewerId or "no token
recoverable" (never a guess). Gates (§16, T9): ≥ 95 % single-page recovery on
default-zoom screenshots of text-rich command units; ≥ 90 % on image-only
units; ≥ 80 % / ≥ 75 % after JPEG q70, 50 % downscale, 25 % crop;
**false-attribution ≤ 10⁻⁶ over the declared table size verified by
simulation + unmarked captures must abstain at 100 %**.

## 14. Performance, memory and limits

- **Rates** [call]: per IP 30 new sessions/min (8 concurrent connects);
  per sid 120 WS messages/10s (burst 20/s), 60 unit requests/min, 10 print
  jobs/min (2 concurrent renders), 30 sort/filter ops/min (1 in flight);
  process 256 sockets, 4,096 queued units.
- **Sizes** [call]: `maxUnitPx 8192`; decoded-image ≤ 64 MiB/unit;
  inflated commands ≤ 8 MiB/unit; WS frame chunks 1 MiB (binary frames per
  §7.1 header); assets/session 256 MiB; show plans ≤ 128 MiB aggregate
  decoded (§12a.2); manifest ≤ 256 KiB decoded; render deadline 2 s;
  outstanding units 8 (queue 32); session scratch lease ≤ 2 final-raster
  units in flight (§9.3).
- Client: replay ≈ today's interactive paints; DOM pressure drops (no file
  buffer).
- Determinism: same docRef + same layoutFingerprint inputs ⇒ same record;
  the recorder reads no clock.

## 15. Accessibility

Server text service returns text for coordinates from the §7.3 second-pass
index, subject to the §11 join — so a11y text carries exactly what the
policy+join permit (possibly less than local). `can.textA11y` never
authorises notes; `can.notes` re-evaluates per session and after resume.
Focus/keyboard surfaces unchanged; a disabled service states it once.

## 16. Testing and validation (suite IDs; single home)

- **T1 contract:** corpus record→replay→pixel-diff zero local-backend diff;
  per-script tolerance table; cross-backend published tolerances.
- **T2 disclosure:** wire-inspection over transmitted payloads; hidden
  sheet/painting/`w:vanish`/**hidden shapes (`cNvPr`/`docPr`, through
  groups)**/notes/comments/covered cells/masked bands/unsampled-times matrix;
  per-endpoint (search/copy/a11y/filter-values) partial-run and covered-run
  withholding.
- **T3 recorder census:** experimental opcode subset closure (per §7.3).
- **T4 fonts handshake:** `requiredFaces` probes; missing/mismatch → raster
  fallback; CJK/ligature/BiDi fixtures; scratch-rendered text case.
- **T5 show plans:** §12a.2 admission limits; no-package/no-shape prototype;
  appearance/intermediate adversaries; remote presenter adapters; fallback
  rate reporting.
- **T6 session/isolation:** handshake state machine incl. version-intersection
  refusals, wrong-viewer resume, missing `resumeProof`, `renew` race rule,
  notes recheck after resume; client-relax refusal; paired relaxed/protected
  sessions over both cache-fill orders (both marks settings) sharing zero
  entries; per-`contentPolicyId` processing; sort/filter authorized targets,
  tautology-filter (row-hidden purge) and hidden-slice/pivot-cache refusal.
- **T7 marks:** ΔE00 ceiling measurement + parameters published; pixel-mark
  injection/extraction on image-only units; survival batteries (q70/50 %
  downscale/25 % crop); unmarked-capture abstain batteries; K-capture
  averaging and 2-capture differencing measured bounds; text-layer
  round-trips incl. decoy distribution and sparse cross-page spooling.
- **T8 XLSX:** tile pipeline (no `StaticExportLimitError`), sort/filter
  authorization, covered-cell filter values, per-session view-state key.
- **T9 detection:** §13.5 gates incl. joint z-score gating.
- **T10 lifecycle:** handshake close-code matrix (`4xxx` surfacing),
  backpressure/cancellation, drain-then-refund retirement, resume/supersede
  sockets, print job audit.
- **T11 eager:** transitive-graph budget, `dist/core.js` purity, module-map
  identity.
- **T12 perf probing:** 1,000 concurrent unit requests hold caps.

## 17. Staging

Each stage a PR; gates by §16 suite IDs; measured, not narrated.

| Stage | Scope | Done when |
|---|---|---|
| **0** | T1–T7 batteries, incl. backend confirm (`Skia`), §12a client prototype, ΔE00 ceiling + coding params, fonts handshake probes, appearance-eligibility rate | gates green **with published numbers** in `docs/validation/streaming.md`; any gate failing reports to the owner rather than silently switching choices |
| **1** | core entry + honest eager gates; recorder + replay; WS handshake (§6.2) + typed closes; RemoteOfficeDoc branch; demo WS-only server (pool, default-protected sessions, marks, per-session print off/on wiring); T1/T2/T6/T11 green | corpus zero local-backend diff; handshake close-code matrix (T6); eager transitive + core purity |
| **2** | marks + policy per §13, occlusion store, per-session scratch lease; audit table; detection tool; M-rule prototype (decided scope) | T6/T7/T9 green incl. false-attribution ≤ 10⁻⁶ over the declared table; marks-off sessions audit-recorded |
| **3** | §11 semantic services + print (§2 item 9) + per-session marks; notes presenter surfaces | T2 notes variants; T8 partial; print audit |
| **3b/3c/3d** | stream presentation per §12a (static+notes; animation; optional media — 3d = §12a.6 transcoded playback) | §12a post-conditions verified in-browser from §16 T5 (see also the notes/3d items above; 3c further blocked on the no-package client prototype, §16 T10) |
| **4a/4b/4c** | XLSX tiles via the viewport pipeline; selection parity; server sort/filter with the C1 purge + authorization + hostile tests + rate class + pivot/slicer refusal strips | T8 green |

## 18. Boundaries

**Always:** developer chooses mode by source type; once a session is open its
protections are server-side state (no client relaxation); fail-closed on
unknown op/version/opcode/manifest; provenance on every artefact
(`docRef ⊕ contentPolicyId ⊕ fingerprint ⊕ scaleBucket ⊕ markSettings ⊕
projection ⊕ unitId`); wire audits over screenshot-only assertions; honest
docs language (screenshots are traceable, not preventable).

**Ask first:** shipping font geometry for faces the document only embedded
(outline extraction — moot while text mode is decided plain-text, but kept
for future layers); persisting sessions beyond process restart; audit-table
export beyond the operator; serving streams for documents that previously
came as bytes; retention beyond the owner's written policy.

**Never:** ship document bytes or any member/media/font bookkeeping bytes
over this protocol in any form (debug frames included; the `server/`
package has no file route, and the library cannot stop a developer who
publishes the file elsewhere — stated in §6.1/§19, not claimed away); ship
original encoded media/metadata-bearing replicas; answer a **protected**
session without the P-rule; ship hidden content in any payload
(§6.4/§12a.3/§16 T2 always); log `OFFICEVIEW_STREAM_KEY`, tokens, sids or
audit rows; pretend DRM; execute client input inside the render pipeline.

## 19. Honest limits

1. **Screenshots/cameras:** traceability, not prevention (§2 item 11);
   marks-off sessions are untraceable by construction (§6.1).
2. **Deployment boundary:** the library cannot stop a developer who also
   publishes the file elsewhere; `server/` has no route (§6.1, §2 item 4).
3. **ODT/RTF refusal is a deliberate scope limit** until hidden-text
   filters exist (§6.2).
4. **Collapsed row/col extents reveal a hidden band exists** (§12);
   accepted, with a "no numbering" alternative named.
5. **Presenters can retype notes**; marks trace screenshots (§12a.5).
6. **The pixel-parity story is deployment-specific** (backend + font set +
   bucket); §16 T1 measures what the deployment pins; the §7.6 tile path
   guarantees a wrong pixel never paints as right.
7. **Pixel-mark collusion has a measured, honest bound** (§13.3): K-capture
   averaging and 2-stream differencing *do* reduce recoverability on flat
   regions; the claim is limited to the measured value, not "no attack".

## 20. Appendix: command-projection extension (Stage 2+)

The default projection is `final-raster`. The command-projection extension
ships later, stage-gated, for deployments that want smaller payloads:

- **`culled-commands` (protected, opt-in):** per-command occlusion proof
  (recorded with the record, never the recorder's opinion); units failing
  proof downgrade to `final-raster`. The **M-rule** requires **every
  content-bearing component** of the unit to be a protected transmitted
  raster (a full-page raster or rasterised flatten group); a unit whose
  readable content rides commands cannot be guaranteed, so such units
  answer `final-raster` (round-7 fix, agy Codex-r4-2).
- **`commands` (relaxed sessions):** ordinary stream; visible mark as a
  last-position `RASTER_BLOCK` overlay, removable by a modified replayer —
  stated plainly, because a relaxed session trusts the replayer.
- **Layer-2 text displacement (normative in the appendix):** candidates per
  §13 HMAC-seeded (length ≥ 4 grapheme clusters, body baseline,
  non-whitespace, ≥ 3 word boundaries apart); placements `±δ` on a
  1/256-px lattice, δ calibrated (target < 0.5 CSS px); decoy jitter `±δ/2`
  keyed on non-carriers (a Stage-2 measured hypothesis, never claimed as
  undetectable); decision rule: observed bit string vs each candidate's
  expected codeword with CRC-8 as a **sample-integrity check over the
  observed 32-bit payload**, distance threshold and unique-winner or
  abstain, and a whole-search `≤ 10⁻⁶` false-attribution target calibrated
  by simulation over the audit-table size.
- Sparse pages spool codeword bits over sibling pages with distance-k
  interleaving; `none` pages carry no text layer (the layer-1 pixel mark
  covers them).

## 21. Review history

All round-by-round review dispositions (codex r1–r4, claude r1, agy r1)
live in the sibling file `docs/superpowers/specs/2026-11-07-server-streaming-reviews.md`,
including the round-7 table for this consolidation.
