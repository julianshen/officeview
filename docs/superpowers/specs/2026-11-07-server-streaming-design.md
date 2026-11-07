# Server-streamed Office rendering: design

Status: **draft 4 (reviews: codex r1–r3, claude r1)** — draft 2 folded codex r1 and
claude r1 (first disposition table); draft 3 folded codex r2 (second table);
draft 4 folds codex r3 (third table). Owner decided §18 items 1, 3, 4 and 9 on
2026-11-07. Docs-only stage continues.
Branch: `spec/server-streaming`, from main `ca952c50`. Docs-only stage: no source edits.

Notation: **"to verify in Stage 0"** marks a claim this spec could not prove while writing
— mostly about canvas backends, glyph shaping/outlining, disclosure projections and
protocol sizes. Stage 0's prototypes must confirm or correct each one before any
production code lands. Values are stated as facts only where the code was actually read,
at the paths shown; round-2 review corrections to round-1 citation anchors are applied
and the corrected anchor used.

## Table of contents

1. Intent
2. Owner decisions (2026-11-07)
3. Goals, non-goals, anti-goals
4. Current state (recon, main `ca952c50`, anchors re-verified)
5. Architecture overview
6. Open protocol: negotiation, manifest, and the server content policy
7. The draw-command format (`OVSD`), the recorder, and the protected projection
8. The client replayer and the library side
9. The server side
10. Behaviour detail: open, failure, resume, expiry
11. Search, copy and semantic services become server calls
12. XLSX in stream mode
13. Forensic watermark
14. Performance, memory and limits
15. Accessibility
16. Testing and validation
17. Staging with acceptance criteria
18. Open decisions for the owner
19. Boundaries: always / ask first / never
20. Honest limits
21. Review dispositions (round 1)

## 1. Intent

officeview is a client-side renderer: whoever wants to read a document must receive the
document. For a confidential corporate spreadsheet that is the wrong trust model — the
bytes sit in the browser cache, the inspector's network tab shows the original file, and
every "protection" the viewer offers (`src/core/watermark.ts:6–13` documents it) is a
*deterrent, not DRM*.

This design adds a second, server-driven mode: a server parses, lays out and paints the
Office document on its own machine, and **streams the client only what is drawn** —
recorded Canvas 2D draw commands (or pre-rasterised image tiles) plus the minimal
interaction geometry a purely visual viewer needs, all filtered by an explicit
**content policy** and a **protected projection** (§6.4, §7.7) so "what is drawn" is
also "nothing that is hidden". The original file never crosses the wire in this mode.
Search, copy and semantic services become server calls whose answers are
client-rendered exactly like every other response.

The client stays one library. There is **no client "secure mode"**: the loader looks at
what the server actually sent and behaves accordingly.

## 2. Owner decisions (2026-11-07)

These are binding; this draft implements them. Round-2 changes sharpen several of them
but revise none of the seven.

1. **No client-side secure-mode switch.** The client renders whatever the server's
   response is: office file bytes → local render exactly as today; a stream manifest
   (its own content type + magic bytes, §6.2) → remote session. The server decides.
   For a protected document the server must refuse (406) clients that cannot handle
   streams, never fall back to sending the file.
2. **Open source, same repo.** The library stays at the repo root; the server becomes a
   `server/` Bun workspace package depending on a new React-free `officeview/core` entry
   (§8.1). *(Round-2 scope note:* the *library* stays dependency-free; the server may
   adopt rendering-side dependencies (shaping, out) as Stage 0 choices (§18 item 2).)
3. **Library side:** React-free core entry (§8.1); versioned draw-command format +
   recorder, with the record point inside the single `getPaintables` provider —
   `page.paint(ctx)` closures (`src/render/paint.ts:354` declares `getPaintables`; the
   paint closures begin at `:389`) — + client replayer (§7, §8); manifest detection in
   the loader seam (§6.3; the actual choke point is `loadOfficeSource`, `src/components/
   loadOfficeSource.ts:199–224`, not only `openResponse`); a remote document source for
   `OfficeDoc` (§8.3); a Node `registerFont` adapter (§8.4).
4. **Server side:** HTTP/WebSocket service, short-lived tokens, rate limits, render
   worker pool, caching, per-viewer **visible** watermark, **invisible forensic
   watermark** (sub-pixel line/word shifts encoding a session id with error correction,
   HMAC-keyed), session-id table, and an internal detection tool (§9, §13).
5. **Search and copy become server calls** whose availability the manifest declares; the
   existing `X-OfficeView-Protection` / `X-OfficeView-Watermark` precedent
   (`src/core/stream.ts:51–75`, header parser at `src/core/watermark.ts:164`) is the
   model for "server is the authority" (§11, §9.4).
6. **XLSX:** the existing interactive viewer sends cell records including formulas
   (`src/xlsx/view/contracts.ts:74–80`), hidden rows/cols (`src/xlsx/view/
   contracts.ts:671–673`, declaration at `:668`), and raw picture bytes
   (`src/xlsx/types.ts:27`) to the client; stream mode needs server-painted tiles and
   display-only replies (§12).
7. **Honesty:** screenshots and cameras cannot be prevented, only traced.

Where this spec places a default that is not in the owner list, it is marked **[call]**
and sized to be changeable without rework. Owner-level decisions still open are in §18;
stages may not silently pick them.

## 3. Goals, non-goals, anti-goals

### 3.1 Goals

1. **Stream mode exists and is safe by construction.** A server can serve any DOCX /
   PPTX / static XLSX document as a manifest + draw-command stream, and officeview
   renders it without ever holding the file.
2. **Same pixels, two routes — with one fidelity gate per text mode.** Each text mode
   names exactly the reference its fidelity is measured against (§7.4): outline mode is
   graded against the server's own outline-rendered reference (never against the
   hinted original), text mode can be graded zero-diff only where the agreed client
   faces are present (a §7.4 face handshake otherwise picks the raster fallback).
   Within a mode, recorder→replayer equals the same-backend direct render (§16
   contract test), and cross-backend tolerance is published, not claimed.
3. **Protected documents never leak bytes — nor hidden or covered content.** A
   stream-required document is answered 406 for non-capable clients (§6.1), never
   carries original media bytes on the wire (§7.7), and every one of its responses
   ships under the **mandatory final-visibility projection** (the single rule of §7.7,
   P-rule) so covered or clipped-out content cannot be recovered even from a
   re-ordered replay; hidden content (§6.4) never streams regardless.
4. **Watermarks (§9.4, §13).**
   - **Visible:** per-viewer by default (identity text), deployable as a fixed
     deployment-wide mark, painted into the raster server-side.
   - **Forensic:** a per-session tracking pattern invisible at reading size — recorded
     text positions shift by a sub-pixel amount (§13.3) — encoding a session token with
     error correction, recoverable from a single screenshot page by the §13.6 tool.
5. **Search, copy and semantic text by server calls** over the visible projection
   (§11), constrained by the same content policy as the pixels (§6.4).
6. **XLSX stream tiles** (§12): interactive spreadsheet mode renders server-painted
   tiles; cell formulas, hidden rows/cols metadata and drawing picture bytes stay on
   the server.
7. **Session governance** (§9.3): short-lived signed tokens, typed rate limits with
   numbers, per-op authorisation checks, revocation, resume, an audit mapping that
   outlives the render session, and detection confidence rules (§13.6).

### 3.2 Non-goals

- True DRM. Anything a human can read can be photographed; §20 states this plainly.
- Preventing the *document owner* from doing anything — this is about restricting
  *distribution to the client*, not the reader's rendering fidelity.
- Converting stream mode into a general display protocol (no cursor/video streaming, no
  VNC in the browser). Only document ink travels.
- Rewriting the client's local renderer; local mode stays exactly what it is today.
- Publishing, multi-tenant account management, dashboards. The server is an embedded
  component for a host application, not a product.
- Animated slideshow playback in stream mode (§18 item 4, **owner-decided 2026-11-07**:
  no animation for server rendering for now; `present()` is refused for remote
  documents, and slides page as static units like any other page).
- Client-side content inspection as a boundary: the boundary is enforced **server-side
  before anything leaves** (§6.4, §7.7); the client is not trusted with it.

### 3.3 Anti-goals: failures even if they "work"

- Covered content in the transport. Commands that record later-covered z-order
  content both leak on a re-ordered replay and break the P-rule (§7.7), which every
  stream-required response must satisfy by construction (final-unit rasterisation or
  proof-culled commands); the §16 boundary tests inspect **transmitted payloads**, not
  screenshots.
- Original bytes anywhere in the stream: no ZIP member, no embedded media file, no font
  file, no EXIF/metadata-bearing image bytes (§7.7, §19).
- Hidden content on the wire: a stream of a document with `veryHidden` sheets or
  vanish-text that exposes them through pixels, the text index, copy, a11y or the
  manifest (§6.4).
- A "protected" document whose bytes reach the client through any endpoint (406 must be
  truthful; any debug protocol that serves bytes for a protected doc is a bug it
  implies — never ship one).
- A client-side switch that turns a stream into a file fetch. There is no such switch,
  and no prop that if mis-set would create one.
- A visible-in-source claim of "un-screenshot-able" content. The forensic watermark's
  value is traceability, stated as traceability (§20).
- A replayer that silently mis-paints. Every unrecordable canvas operation either
  downgrades that unit to the tile mode (§7.6) or refuses the unit with a diagnostic;
  nothing fakes it.
- A server that treats client-submitted text or markup as content, style or code inside
  the render pipeline; the only client inputs are validated commands and tokens, and
  neither ever reaches the painter.

## 4. Current state (recon, main `ca952c50`; anchors corrected in round 2)

- **Byte sources and policy headers today.** `loadOfficeFile` accepts any `ByteSource`
  (bytes, Blob, URL, Request, Response, ReadableStream, HttpSource and promises —
  the union is authoritative at `src/core/stream.ts:14–31`, the loader entry is
  `src/components/OfficeFile.tsx:37–49` which delegates to `loadOfficeSource`);
  every request/response source flows through `readSource`/`openByteStream`
  (`src/core/stream.ts:207`, `:368`) which fires the protection and watermark
  listeners **before the first byte**, from the download response headers:
  `X-OfficeView-Protection: no-copy, no-print` (`src/core/stream.ts:51–75`) and
  `X-OfficeView-Watermark: text=DRAFT; rotate=…` (parser at
  `src/core/watermark.ts:164`) — the "server is the authority" pattern this design
  extends. `openResponse` (`src/core/stream.ts:328–339`) throws the plain HTTP
  error on any non-OK status **before reading the body** — a 406 JSON body is
  therefore unreadable through today's seam (round-2 fix in §6.3). `pullStream`
  (`:279–325`) forwards chunks and has **no push-back buffer**; the loaded prefix
  that already exists lives one layer up, in `loadOfficeSource`'s sequential loader
  (`src/components/loadOfficeSource.ts:184–224`, the `prefixBytes` pull and the
  `bytes / blob / stream` arms).
- **Everything document-shaped reaches the canvas through one provider.**
  `getPaintables(doc, options)` (`src/render/paint.ts:354`) returns paintables whose
  per-unit `paint(ctx)` closures (first at `:389`) are the record point every
  renderer — layout-engine DOCX pages, `renderSheet` spreadsheets, `renderSlide` —
  funnels into. **Public direct painters also exist** (`renderSheet`, `renderSlide`,
  `renderPages` re-exported at `src/index.ts:49–56`); the server commits to
  `getPaintables` and never paints through the direct exports (§8.2), so one wrapper
  context genuinely wraps the paint surface. Paint options already carry the headless
  seams: `registerFont` (the `RegisterFont` type at `src/core/fonts/register.ts:9`;
  the browser default at `:21` is module-private and unexported; the exported worker
  one is `:30`) and `decodeImage` (`paint.ts:219–220`).
- **Text measurement is already headless — with a caveat.** The headless path in
  `measurerFromDoc` (`src/render/paint.ts:184–211`, not where draft 1 cited it) falls
  back to `import('canvas')`, **hard-coding node-canvas**; the same hard-coded import
  exists in the search-index builder (`src/core/search.ts:212–215`). A backend factory
  seam is required for Stage 0's backend matrix (§8.4, §18 item 1).
- **A Proxy recording context already exists — text-only.** `createRecordingContext`
  (`src/core/search.ts:92`) wraps a 2-D context and captures `fillText` /
  `strokeText` calls with document-space positions; `buildTextIndex`
  (`search.ts:315–335`) drives it as **a separate, second paint pass**, with
  `setWatermarkMuted(true)` around it (`search.ts:322–334`) so the search index never
  ingests watermark text. Round-2 rule (both reviewers): the visual recorder is
  **not** this context — painters branch on `RECORD_TEXT`
  (`src/core/text-recording.ts:2`; e.g. `src/pptx/render.ts:81` skips the master
  underlay in index mode), so the index must remain its own pass (§7.3).
- **Painting is device-space dependent — measured, not assumed.** `snapToPixels`
  (`src/drawing/scene-paint.ts:32–40`) reads the current transform and writes an
  absolute `setTransform`; text shadows (`src/drawing/paint.ts:363–397`) read the CTM
  and write device-space offsets and an absolute transform; the radial-gradient brush
  fallback builds a scratch `ImageData`, `putImageData`s it and calls
  `pattern.setTransform` during paint (`src/drawing/paint.ts:178–196`); sparklines
  scale off `ctx.getTransform()` (`src/xlsx/render.ts:774–775`); charts lay out with
  the document's DPR and key their layout cache on `dpr`
  (`src/charts/paint.ts:767–768`, `src/charts/layout.ts:253–263`). Text painting
  constructs measurers and performs layout **during** paint (`src/drawing/
  text-paint.ts:972–989`, side canvases via `src/core/text-metrics.ts:23–40`), draws
  cached scratch surfaces back into the main context under skew
  (`text-paint.ts:712–765`), and feature-detects with `'letterSpacing' in ctx`
  (`text-paint.ts:113`) — so any recorder must be a **facade over a real backing
  context**, not a description-only proxy (§7.3). Round-2 rule: record at a concrete
  resolved scale and replay only at it (§8.3).
- **Watermarks paint into the bitmap**, not the DOM, so a screenshot pipeline captures
  them (`src/core/watermark.ts:6–13`); defaults (`WATERMARK_DEFAULTS`,
  `src/core/watermark.ts:48–57`) cover placement/rotation/opacity/tile gap. Note
  `setWatermarkMuted` is **module-global** (`watermark.ts:249`) — pool workers must
  not interleave index and paint passes inside one isolate until the mute is
  context-scoped (§9.2).
- **The worker render protocol already has the server-shaped env seam.**
  `WorkerRenderEnv` (`src/worker/office-worker.ts:42–47`) abstracts `createCanvas` /
  `toBitmap` / `parse` / `paintables`; `handleWorkerRequest` (`:50–72`) validates and
  paints one unit. The server pool reuses the shape (§9.2).
- **XLSX today paints every sheet.** There is no `state`/hidden flag on `XlsxSheet`
  (`src/xlsx/types.ts:151–175` — verified: no hidden/veryHidden anywhere in `src/`).
  `getPaintables` maps `doc.sheets` in document order (`paint.ts:427–429`) with no
  hidden-sheet policy, and nothing in the parsers reads `w:vanish`, hidden PPTX
  slides or notes for the static path. Draft 1's "static XLSX streams with no special
  work" claim is **withdrawn**: hidden-content policy (§6.4) applies to every format
  and is a Stage 1 gate.
- **Eager budget.** Last recorded measurement: eager 38,582 B vs project ceiling
  38,647 B and hard limit 38,720 B (`scripts/eager-budget.ts:14`; ledger in
  `docs/validation/chart-compat.md:25,31` — re-measure fresh, never reuse stale
  numbers). The core entry (§8.1) is a *new build target*; the gate is a measured
  unchanged main bundle (§8.1, §16.1), not an assumption.
- **Test infrastructure the work reuses.** `src/test/pixel-diff.ts` gives RGBA +
  `diffBitmaps` + `renderPaintables` (`:98`) plumbing; `scripts/golden.ts` a per-file
  record/compare CLI; `scripts/golden-corpus.ts` renders the whole real-file
  `corpus/` through `loadOfficeFile → getPaintables → renderPaintables` — the
  corpus contract test (§16.1) is this loop with a recorder between the renders.

## 5. Architecture overview

Two new libraries in the same repo; zero new library-side dependencies (server-side
dependencies are Stage 0 choices, §18 item 2):

```
+--------------------------------------+        +----------------------------------------+
| HOST (embedding officeview)          |        | SERVER ("server/" Bun workspace)       |
|                                      |        |                                        |
|  <OfficeFile source={url}>           |        |  HTTP: /open (manifest | 406)          |
|      │  sniff at loadOfficeSource    | <----> |  WS:   /session (token, §9.3)          |
|      ▼    (§6.3)                     | HTTP/WS|      │ policy gate (§6.4)             |
|  <OfficeDoc document={RemoteDoc}>    |        |      ▼ parse+layout once per hash     |
|      │  unit → replay at its scale   |        |      ▼ (doc cache, §9.3)              |
|      │   (plays OVSD §7)             |        |  pool: render workers (§9.2)          |
|  search/copy → WS op (§11)           |        |  recorder over page.paint (§7.3)      |
|  a11y text → server text op (§15)    |        |  projection (§7.7) + marks (§13)      |
|                                      |        |  caches, audit table, detection tool  |
+--------------------------------------+        +----------------------------------------+
```

- **Client** stays `officeview` at the repo root. The React-free core entry
  (`officeview/core`, §8.1) exports parsing, paint, fonts and stream plumbing without
  React. The stream loader, format code, manifest types and the remote document
  source live in the library's lazy `src/stream/` plus shared format code exported
  from `officeview/core` — the server depends on React-free machinery only.
- **Server** is a new `server/` Bun workspace package (§9). It imports
  `officeview/core` for parse/layout/paint, runs its own canvas backend and render
  pool, and owns the policy gate, tokens, caches, marks and the detection tool.
- **The wire** is: one HTTP open call, one WebSocket per session. The open response
  is either the original bytes (transparent local path, §6.1) or a **manifest**
  (§6.2). The manifest carries unit geometry, a `docRef` (not the plain hash, §6.2),
  capability flags, a short-lived signed session token, and the recorder parameters.
- **Data on the page wire** is per-unit `OVSD` chunks (§7.1): a versioned,
  length-prefixed draw-command record — transforms, state, path geometry, glyph
  runs, image references — produced by wrapping each `page.paint(ctx)` closure
  (`paint.ts:354/389`) in a backend-backed recording facade (§7.3), filtered by the
  protected projection (§7.7), cached under the §9.3 canonical key, and post-processed
  per session for the marks (§13.3). Units the recorder cannot express become
  **tiles** at stream time (§7.6).

## 6. Open protocol: negotiation, manifest, and the server content policy

### 6.1 Who decides what — and the 406 contract

- **The client announces capability with safelisted headers only.** Round-2 fix
  (claude I6): the loader sets **no custom request header** (a custom header forces a
  CORS preflight on every cross-origin local load). It sets:
  - `Accept: application/vnd.officeview.stream` (with parameters of the media type,
    e.g. `application/vnd.officeview.stream;v=1`, all values of `Accept` are
    CORS-safelisted, so the local path stays a simple request);
  - the same media-type parameter carries the largest recorder version the loader can
    replay (`;rec=1`), which replaces draft 1's `X-OfficeView-Recorder` header.
  The loader sets these **additively** on every request it issues and never reads,
  filters or copies the host's own headers; a host supplying its own `Request` or
  `Response` can add whatever it wants — an important door for
  host-credentialled loads.
- **Server-side configuration decides per document** (the host's server config marks a
  document stream-required **by a rule the owner supplies** — path, content-kind or
  digest; the concrete config surface is owner decision §18 item 5). For a
  stream-required document the open endpoint answers a manifest **only when the
  client announced the capability** (`Accept` media type + a `rec=` parameter the
  server can replay).
  Otherwise it answers **`406 Not Acceptable`** with a bounded JSON body
  `{ error: 'stream-required', required: 'rec' }` (≤ 256 B). That body must be
  *readable*, which is why the loader's 406 handling is defined on a typed
  non-OK-response reader (§6.3), not inside today's `openResponse` throw. A server
  that cannot classify the client treats it as **capability-absent** — never as an
  invitation to send bytes.
- Un-required documents may be answered with either bytes or a manifest — the
  server's choice; the client is unaffected (the loader sniffs, §6.3) and **never
  upgrades a refusal into a second request for the file**; for a protected document
  that would be the leak.
- `X-OfficeView-Protection` still travels on manifest answers; in stream mode the
  **effective policy is the AND of the protection header and the manifest's `can.*`**
  (§8.3) — a server cannot widen policy through the manifest, only narrow it.
  `X-OfficeView-Watermark` is not sent on manifest answers: the visible mark is baked
  server-side (§9.4), and a remote document ignores the client's watermark prop and
  header entirely (§8.3).

### 6.2 The stream manifest (`OVSM`) — one normative schema

Content type **`application/vnd.officeview.stream`** (round-2 fix: the framing is
binary; no `+json` suffix). Magic bytes **`OVSM`**:

```
Offset Size
0      4   'O''V''S''M'                    ASCII magic
4      2   u16 LE  formatVersion = 1
6      2   u16 LE  compression flag: 0 = raw JSON, 1 = gzip (the single flag)
8      4   u32 LE  manifest JSON length — **decoded** length in bytes
12     n   manifest JSON (UTF-8), gzip-compressed when flagged (the wire length is
           recoverable by reading to the frame boundary; the field is never a mix of
           compressed/decoded semantics)
```

Manifest JSON (normative; fields marked `[call]` are first-pass defaults):

```
{
  v: 1,
  // Document identity. docRef is HMAC(kDoc, sha256(bytes)) truncated to 96 bits,
  // hex — NEVER the plain hash: a wire-visible plain hash would let a holder
  // confirm a byte-exact match against an offline guess list (round-2 fix).
  docRef: <hex>, kind: 'docx'|'xlsx'|'pptx', name: <label>,
  units: { count: N, list: [ per unit { w, h } ] },      // CSS px at 96 dpi
  layoutFingerprint: <hash over (officeview version, parser/layout inputs, backend
    version, canvas font set, locale)>,                  // one canonical key input (§9.3)
  session: {
    token, wsUrl,                                       // origin pinned by §9.3
    expiresSec,                                         // [call] 900
    resume: 'session'                                   // single scope, §9.3
  },
  can: { search, copy, print, textA11y },               // ANDed with protection header
  policy: { contentPolicyId, viewerMarkId,
            semantic: 'visible-only' | 'off' },       // §6.4 chart-text rule; cache-key input
  recorder: { commands: 'ovsd-1', textMode: 'outlines'|'text',
              scalesBuckets: [1, 2],                    // replay-only-at-bucket (§8.3)
              projection: 'final-raster' | 'culled-commands' }, // P-rule (§7.7)
  limits: { maxUnitPx: 8192, maxAssetBytes, wsFrameBytes, tilePx },
  notices: [ '...' ]                                    // viewer status-line string, dedup'd
}
```

Validation rules (client, fail-closed): numbers range-checked; unit pixel sizes in
`1..8192` (a unit larger than `maxUnitPx` cannot be rendered — oversized units
are down-scaled or tiled server-side, §14); enums closed; `token` ≤ 512 chars and **never**
placed in a URL path (transport in the first WS frame, §9.3); unknown properties
ignored (the same forward-compat stance the protection header already takes,
`stream.ts:65–75`).

**Session and token schema (normative, one definition):**

| Field | Meaning |
|---|---|
| `sid` | 48-bit random per render session (not the viewer's identity; §13.4) |
| `token` | HS256-signed claims `{v, docRef, sid, exp, iat, jti, scope:'stream'}` under `kToken` (`kMaster ⊃ HKDF` subkey §13.4) |
| `jti` | token instance id; every WS op re-checks signature, `docRef`, `exp`, and the `sid` bound to the connection |
| resume | one scope: 'session'; a resume re-attaches viewport + visible units. Renewal exists only for replayable-by-URL sources (§6.3): such a token carries the **optional** `refresh` claim naming the re-open grace; bytes/Blob/Response manifests carry no renewal and are terminal at expiry |

### 6.3 Where the sniff actually lives (loader seam) — round-2 rewrite, renewal classified round 3

Round-1 draft claimed `openResponse` peeks the first 12 bytes with an "appendable
reader". Both reviewers are right that this is false: `pullStream` (`stream.ts:279–325`)
forwards chunks with no push-back, `openResponse` throws on non-OK before reading any
body, and `loadOfficeSource` **already holds** the first chunk(s) of the stream in its
sequential-loader prefix (`loadOfficeSource.ts:184–200`) before choosing an arm
(`:210–224`), with `blob`/bytes arms beside it. Round-2 rule, one coherent seam:

1. **Sniff in `loadOfficeSource`**, on the prefix it already holds (`loadSequential`):
   if the first bytes (tolerating a 12-byte header that may span chunk boundaries)
   match `OVSM`, route to the **manifest arm** (`loadManifest`);
   `largeSource().continueSequential` (the existing large-source arm) gains the
   manifest outcome instead of spooling bytes.
2. The `bytes` / blob arms sniff their first bytes symmetrically (`loadBytes`, blob
   hint) — an office file never starts with `OVSM`, and the fail-closed default on any
   ambiguity is the existing (bytes) arm.
3. **406 handling**: `openResponse` (`stream.ts:328–339`) gains one early branch — on
   non-OK, read up to 256 B of the body (bounded, aborted-safe) and throw a typed
   `StreamRequiredError { status, code, required }` when the body parses to
   `stream-required`; other statuses keep today's error. **Nothing in this path ever
   re-fetches the document from a second endpoint.**
4. **Result union and lifecycle**: `loadOfficeSource` `AnyDoc` union gains a
   `RemoteDoc` member; `OfficeFileState` grows `{ status:'ready', document: RemoteDoc }`;
   OfficeDoc's remote path is §8.3. The remote arm owns: the held socket client, the
   buffers, and close-on-unmount — the existing `pagedWorkbookOf(document)?.release()`
   lifecycle pattern in `OfficeFile.tsx:127,140` is the disposal model to copy.
5. **Renewal classification (round-3 fix):** sources split decisively, because
   re-reading a Blob or byte array that carried a manifest replays the *same expired
   token*:
   - **Renewable**: URL, HttpSource. A **Request** is renewable only under a
     documented replayability condition — no body, or a body not yet consumed
     (`request.bodyUsed === false`); a body-bearing or already-consumed Request is
     terminal.
   - **Promises are classified by their resolved value**, not by being promises:
     the loader awaits such a promise once and dispatches the resolved object
     (`src/core/stream.ts:368–373`), so `Promise<HttpSource>` inherits
     renewability, while `Promise<Response>` yields the *same consumed Response*
     and is terminal.
   - **Terminal**: bytes, Blob, Response, ReadableStream, and any already-consumed
     or body-bearing Request — when the manifest's token expires the loader
     surfaces the typed terminal `stream-expired` state; recovery is the viewer's
     reload affordance, never a speculative re-open. A host that wants renewable
     in-memory input passes a live `HttpSource` (or a factory returning fresh
     Requests).

### 6.4 Server content policy — the server decides what exists (new; codex 2, claude C1/I2)

"Stream what is drawn" is bounded by "drawn **for this viewer**": a document can
contain content that the local renderer would paint but a stream deployment must not.
The policy gate runs **before layout**, once per document, under a
`contentPolicyId` recorded in the policy projection (§9.3) and asserted by Stage 1
wire tests (§16.2):

| Content | Local mode today | Stream-mode rule |
|---|---|---|
| Hidden / `veryHidden` XLSX sheets (`state` attribute, unread by `src/xlsx/parse.ts:290–300`) | painted as extra pages (`getPaintables` maps every sheet, `paint.ts:427–429`) | **not streamed** as units and excluded from the unit list, sheet navigation and all text services; restricted case: a deployment option may stream hidden sheets for a restricted class of viewer (owner choice, default off) |
| `w:vanish` runs (unread by the DOCX parser today) | painted (no handling exists) | excluded from painting, text index, copy and a11y. **A deliberate deviation from local mode**, diagnosed with a per-document notice; a Stage 1 test asserts absence everywhere (§16.2) |
| PPTX hidden slides (`show="0"`) | painted as units | not streamed as units |
| Speaker notes / comments / bookmarks panels | notes unread by the static pipeline; comments unread | never streamed; a slideshow read closes without them (§18 item 4) |
| Hidden rows / columns | painted collapsed (today's own geometry) | axis sizing intervals only (the client sees collapsed gaps); **row/column *numbering* alone while a hidden row has zero extent still reveals the presence of hidden bands — accepted limitation, stated in §20 and §12** |
| Chart text (`chartDescription`, `src/charts/describe.ts:40–62`) | emits model categories, per-series values (first 12 or min..max) — **model data, not paint** | stream-mode chart descriptions are **visible-projection only**: names/labels/categories that are painted; no value ranges. Server-declared via `policy.semantic: 'visible-only'` (§6.2) |
| PTX/DOCX comments, tracked-changes remnants | not parsed | never streamed |

The policy list is versioned by `contentPolicyId` so a policy change invalidates
caches (§9.3). The client never participates in the policy — it renders what arrives.

## 7. The draw-command format (`OVSD`), the recorder, and the protected projection

### 7.1 Container

```
Magic 'O''V''S''D', formatVersion u16 LE = 1
assetTable: counted, length-prefixed records (§7.5)
unitIndex[]: u32 offsets into the frame, one ordinal per unit
for each unit: u32 LE byteLen + u32 LE opCount + a gzip-compressed body
```

One `OVSD` frame holds a document's units; chunks can also be **per unit** (what the
WS sends: an asset-table preamble on the first frame of a session, then one chunk per
unit on demand). A `TILE` op carries its bitmap **inline in the frame** (bounded by
`wsFrameBytes` chunking), not as a data URI — round-2 fix.

### 7.2 Why commands, not video

A document unit is redrawn, not re-filmed: its draw state is resolution-independent
*given the recorded scale* (§7.3), command streams compress very well, and the client
paints any requested unit without fetching the rest. That independence is what makes
the per-session forensic edit (§13.3) a post-pass on cached records instead of a
re-film. The limit is recording fidelity at scale changes (§8.3), which is why the
whole pipeline always names its scale.

### 7.3 The recorder — a backend-backed recording facade (round-2 rewrite)

- **Where.** The server wraps one recording context around each `page.paint(ctx)`
  closure — the record point is inside `getPaintables`
  (`src/render/paint.ts:354` entry; per-unit closures at `:389` and on). The context
  is a **Proxy over a real backing 2-D context** (the generalisation of
  `createRecordingContext`, `src/core/search.ts:92`), not a description-only proxy:
  - **truthiness and `'in'` come from the backing context** (painters feature-detect
    with `'letterSpacing' in ctx`, `src/drawing/text-paint.ts:113` — the facade must
    forward gets/has/sets or feature detection lies);
  - **measureText works**: painters capture measurers and lay out **during** paint
    (`text-paint.ts:972–989`, `src/core/text-metrics.ts:23–40` create side 2-D
    contexts from `ctx.canvas`), so the facade answers deterministic server-side
    measurement synchronously;
  - **side surfaces**: painters build scratch surfaces and draw them back with
    `drawImage` (skewed text through a cached `PaintSurface`,
    `text-paint.ts:712–765`; radial-gradient brush patterns from scratch ImageData,
    `drawing/paint.ts:178–196`). Scratch sources are **mutable and cached**, so a
    `drawImage(surface, …)` on the facade records a **snapshot-by-consumer**: the
    facade snapshots the surface's pixels at draw time, keyed by content hash, and
    records the hash — never a reference that could later be repainted;
  - **patterns and gradients are wrapper objects**: recorded as asset-table entries
    with their own recorded `setTransform` (the `pattern.setTransform` call at
    `drawing/paint.ts:194–196` is real state, not free).
- **The recorded subset** is the painted Canvas 2-D surface the painters use today,
  with a Stage-0 census making it exhaustive (§16.1): state (`save/restore/setTransform
  /translate/scale/rotate/transform/getTransform-consumers/globalAlpha/
  globalCompositeOperation/setLineDash/lineDashOffset/lineCap/lineJoin/miterLimit/
  imageSmoothing*/letterSpacing/wordSpacing/direction/lang/fontFeature-family flags`),
  styles (solid colours, `createLinearGradient`/`createRadialGradient`,
  `createPattern` + wrapper `setTransform`, shadow fields), geometry
  (`beginPath/moveTo/lineTo/quadraticCurveTo/bezierCurveTo/arc/arcTo/ellipse/rect/
  roundRect/closePath`), drawing (`fill`/`stroke` incl. the combined
  fill-then-stroke op, `fillRect`, `strokeRect`, `clearRect`, `clip`), text
  (`font/textAlign/textBaseline`, `fillText` with `maxWidth`, `strokeText`,
  glyph runs §7.4), images (`drawImage`, all 9-arg forms), and the raster-inlining
  ops of §7.7. Ops **outside** the subset are unrecordable → tile (§7.6).
- **`RECORD_TEXT` is not part of the visual recorder** (round-2 merged rule): the
  facade must not expose `RECORD_TEXT` (`src/core/text-recording.ts:2`), so painters
  that branch on it (`src/pptx/render.ts:81` skips the master underlay in index mode;
  `text-paint.ts:994` uses it for records) behave exactly as they do for a real
  canvas — full-fidelity. The text index remains its own separate pass through the
  existing `createRecordingContext` + `buildTextIndex`
  (`search.ts:92,315`) with the watermark muting (§9.2 pass-interleaving, §11).
- **Determinism and scale.** Because the record captures post-decision device values
  (`snapToPixels`, shadow CTM math, `pattern.setTransform`, sparkline scale) at the
  paint-time transform, replay reproduces pixels **only under the same resolved
  scale**; the client replays 1:1 at the unit's scale bucket and re-requests on other
  buckets (§8.3). Recorded coordinates are IEEE-754 little-endian doubles by
  definition of the format (round-2 fix: no "varint coordinates" language; sizes
  stay bounded by the chunk budget, §14).
- **The render deadline is enforced by retirement, not by cooperation.** A paint
  cannot be aborted inside its synchronous `paint()` callback. The pool's 2 s unit deadline
  [call] ends by **worker retirement**: the worker is terminated (not asked to
  yield); the buffers it allocated refund to the pool ledger only after
  termination is confirmed — never before, and never by reusing a buffer the
  dead worker still holds (drain-then-refund, tightened round 3, tested §16.2);
  the in-flight unit retries once on a fresh worker and falls back to tile (§7.6)
  if it fails again.

### 7.4 Text mode (round-2 rewrite; default [call] selected by the Stage 0 matrix)

Draft 1 claimed glyph outlines "with zero new dependencies" arrive for free; both
reviews reject that. The recorder sees strings plus a CSS font string, not glyph ids;
joined-outlining needs a shaping/outlining library (HarfBuzz-wasm or fontkit-class)
and hinted text is not vector-reproducible bit-for-bit. Stage 0 must pick:

- **(a) `textMode:'outlines'`** — a named shaping+outlining dependency server-side
  (open decision, §18 item 2); outlines recorded into the §7.5 asset table; the
  parity gate becomes *(near)-zero-diff only against an **outline-rendered
  reference** generated the same way on the server* — never against the original
  page. Fallback if shaping/complex-script (CJK, Indic, BiDi) parity fails: per-cluster
  bitmap tiles of text (the §7.6 tile path applied to text bands), not per-glyph hope.
- **(b) `textMode:'text'`** — glyph runs recorded as strings plus advances; replay uses
  **server-pinned font data** (embedded-on-the-server faces that the client must
  already have, or the document's own embedded faces — the "ask-first" boundary,
  §19). When the pinned faces exist on the client, this mode is the only one under
  which the §16 zero-diff gate on a real browser is claimable.
- **Face-availability handshake (concrete protocol, round 4; draft 3 referenced it
  without defining it):**
  1. **Identity list.** For `textMode:'text'` the manifest carries
     `requiredFaces: readonly { id, hash, alias? }` where `hash` is SHA-256 of the
     pinned face file truncated to 96 bits [call] — hashes only, no face bytes ever
     travel (§19 unchanged); `alias` is a display name for diagnostics, not a match
     key.
  2. **Bounded acknowledgement.** The client answers in the **first WS frame** with
     `{ op: 'faces-ack', have: id[] }`, bounded to 128 ids and to a **2 s** window
     [call]; no answer in the window is treated as `have: []` (fail-closed). The ack
     is answered exactly once per session; late frames are refused with an
     already-answered diagnostic.
  3. **Match criterion.** A face is satisfied when its `hash` equality holds; a hash
     mismatch (a differing version) counts as missing, not as close.
  4. **Missing/mismatch transition.** Before any text unit streams, the server
     classifies each unit by the face ids its recorded text ops use; a unit needing
     an unsatisfied face is served as a per-unit raster (`unitMode:'tile'`, §7.6)
     with a typed `face-missing` diagnostic. Substituted faces are never painted,
     and the zero-diff fidelity claim is voided only for such units — the
     text-mode goal holds per satisfied unit, not document-wide.
- **One fidelity gate per mode (round-4 fix — round 3 referenced this wording but
  never landed it here):**
  - outline mode → **replay fidelity** is graded recorder→replayer vs the
    outline-rendered reference (server-generated, same shaping); **renderer
    fidelity** to the existing hinted painter is a *published tolerance* per script
    family, never a zero-diff claim;
  - text mode → for units whose required faces all pass the handshake **both**
    gates may be zero-diff; unsatisfied units degrade to raster before streaming.
- Either way, **a11y does not depend on the mode** (§15): the server is the text
  authority through the §15 service.
- Font **subsetting** and never shipping a font file remain hard rules (§19).

### 7.5 Asset table and op list (OVSD v1)

The asset table dedups by content hash: document images (**their drawn-bitmap
snapshots, see §7.7**), decoded scratch surfaces, patterns/gradients (with their
`setTransform`), glyph-outline groups (mode (a)), and pinned font *data hashes*
(mode (b) — hashes, never faces). Ops (normative, ~26):

| Op | Meaning | Replay notes |
|---|---|---|
| STATE_PUSH / STATE_POP | save / restore | full stack incl. clip |
| TRANSFORM | transform / setTransform | exact f64s at the recorded scale |
| STYLE_COLOR | solid fill or stroke target | hex + alpha |
| STYLE_OBJ(gradient / pattern) | asset-table ref | wrapper carries its own `setTransform` |
| STATE_SETUP | width/join/cap/miter/dash/dashOffset/alpha/composite/smoothing/letterSpacing/wordSpacing/direction/lang/shadow fields | dense dword packing decided in Stage 0 |
| PATH_BEGIN/…_MOVE/…_LINE/…_QUAD/…_CUBIC/…_ARC/…_ARC_TO/…_ELLIPSE/…_RECT/…_ROUND_RECT/…_CLOSE / PATH_END {FILL,STROKE,FILL_STROKE,CLIP} | geometry | paths over a budget are shared into the asset table by hash |
| GLYPH_RUN | font ref + glyph data (mode-dependent) + advances + paint style | mode (a) outlines; mode (b) strings + advances |
| IMAGE_DRAW | asset ref + transform + optional source rect | from the drawn snapshot, §7.7 |
| TILE | inline (already-rasterised) bitmap + dest box | fallback mode, §7.6 |
| RASTER_BLOCK | inline raster — the final-frame block of a `final-raster` unit or the visible-mark overlay (§7.7/§9.4) | painted straight |

### 7.6 Versioning, tiles and the runtime-fallback discriminator (round-2 fixes)

- Negotiation is one mechanism: `Accept` media-type parameters (`;v=1;rec=1`), the
  manifest's `recorder.commands`, and the frame header's version. Server picks the
  minimum of what it knows and what the client announced; a *higher* frame version
  than the client supports is a protocol bug and fails closed (the manifest-only
  pre-agreement makes version skew mismatch-free by construction).
- **Tiles are a runtime outcome, not a pre-declared list.** Units are known to be
  unrecordable only when the pool renders them; the manifest therefore **does not
  pre-list** draft 1's `fallbackTiles` — the server session reads each stream unit
  discriminator (`unitMode: 'tile'` in the per-unit preamble) when it streams.
  *(Net effect: no preflight is required — a manifest-declared tile promise would
  be a promise the pool cannot yet know is keepable.)* Chunk-1 preflight stays a
  server config option [call] reserved for status displays.
- Mixing commands and tiles in one document is fine and expected; the client's only
  obligation is to render whichever the unit says, fully, or refuse the unit with
  the missing-unit diagnostic.

### 7.7 The mandatory final-visibility projection (the P-rule, round-3: one rule, all sections reference it)

> **The P-rule.** Every response that serves a **stream-required** document contains
> only content that belongs to the unit's **final composited image** — no covered
> text, no clipped-out media, no full z-order history. It is not a configuration:
> a server cannot answer a stream-required client without it.

Enforcement, in one of two ways (the server owns the choice; `recorder.projection`
names it, §6.2):

- **`final-raster` (default)** — the pool paints the whole unit to the final frame
  server-side and ships it as a `RASTER_BLOCK`/`TILE` op. Final visibility is
  satisfied by construction: a group later covered by another group is cropped away
  *in the raster*, which is exactly what draft 2's per-group flattening failed to do
  (cross-group covers leaked lower groups' text).
- **`culled-commands` (opt-in)** — commands may ship only after the pool runs a
  server-side culling pass over the recorded op list, computing per-command the
  region occluded by every later op (clips, opaque fills, opaque images); any
  command that paints content not surviving with the final frame is dropped or the
  unit **downgrades to `final-raster`**. The proof, not the recorder's opinion, is
  stored with the record; an under-proven unit never sends commands. A protected
  `culled-commands` unit must additionally satisfy the **M-rule** of §9.4 (the
  visible mark composited into its transmitted content rasters, no removable
  mark op); a pure-vector unit that cannot meet M-rule answers `final-raster`.

Both projections also inherit the rule beneath; an *unproven* protected command is a
§16.1/§16.2 gate failure, not a configuration error.

- **No original media bytes (both projections).** Every `drawImage` of a document
  picture contributes the **drawn-bitmap snapshot**: the pixels as painted
  (duotone/recolour already applied at decode, `src/core/images.ts:80–92`),
  **source-rect-cropped** (scene-paint crops with `drawImage` source rectangles,
  `src/drawing/scene-paint.ts:132–149` — cropped-away regions stay on the server),
  **re-encoded metadata-free** at the recorded scale, and keyed by pixel hash.
  Session device scale determines the bitmap; other scales re-request (§8.3).
- **Wire audits inspect all transmitted content, not only commands.** Every §16.2
  boundary test decompresses and inspects **transmitted payloads** — frames, the
  asset table and raster blocks (a covered-text check must scan rasters too, since
  `final-raster` is the default) — for forbidden content: secret strings, pixels
  of the crop region outside the painted crop, EXIF-bearing originals. The
  screenshot-path assertion is never the evidence, because screenshots cannot
  witness the whole stream.
- What stays in scope intentionally: highlight rectangles, the visible mark
  (§9.4), and the visible-projection text data of §11 — nothing else is data the
  client needs for reading.

## 8. The client replayer and the library side

### 8.1 The React-free core entry (`officeview/core`), with a measured eager gate (round-2 fix)

- New build target: `src/core-entry.ts` (name [call]) built by Vite as a **separate
  rollup input** (`dist/core.js`) with React `external`, next to the existing
  `worker.js` entry (`vite.config.ts:41–43`).
- Exports (React-free, no canvas-API assumptions in the round-2 sense — the canvas
  objects flow through `getPaintables` either way): parse/layout/painters
  (`src/index.ts:49–56` equivalents) + `getPaintables` (`paint.ts:354`);
  fonts (`acquireFonts` at `register.ts:40`, adapters `:9/:21/:30`);
  stream plumbing (`readSource`/`openByteStream`, `stream.ts:207/:368`; header
  parsers at `stream.ts:51–75` and `watermark.ts:164`); the manifest/OVSM and
  OVSD format types and helpers; the recorder facade + text index for server use.
- **Eager gate (measurable, not assumed).** The gate is the measured claim:
  `dist/index.js` module map **identical apart from the sniff constant**, and eager
  bytes ≤ the pre-Stage baseline + 2 B gzip recorded as a diff; if the extra Vite
  entry shifts any shared chunk, Stage 1 fixes it before anything else.
- Documented boundary: `officeview/core` never imports React or DOM-only modules
  beyond the existing facade adapters.

### 8.2 How the loader decides; the client-side seam

- The server owns the mode (§6.1); the client's only mode logic is the sniff
  (§6.3) — in `loadOfficeSource` on the prefix it already holds, with the arms
  (`bytes` / `blob` / `stream`) each getting the same check, plus the typed 406.
- New surfaces: **none** for choosing a mode. Two plumbing-only surfaces:
  - `streamClient?: (manifest) => StreamClient` — the WS factory hook (mirrors
    `viewerWorker`, `OfficeFile.tsx:63–80`): hosts that need credentials, resume
    semantics or telemetry supply the client; the default uses `WebSocket` with
    tokens in the first frame (§9.3) and a heartbeat;
  - `onStreamNotice(diag)` — deduplicated per the §10 notice behaviour.

### 8.3 `RemoteDoc` — the remote document source for OfficeDoc (round-2 fixes)

A **lazy `src/stream/remote/` chunk** satisfies OfficeDoc's paintable-provider
contract (the local path keeps `getPaintables`):

```
RemoteOfficeDoc = {
  kind: 'remote', manifest,
  async paint(unit, scale: 'fit' | number, dpr, canvas): Promise<void>,
    // Replays at the unit's recorded scale bucket (the §7.3 rules) or pastes a TILE.
    // Zoom re-requests the server at the new bucket; the previous resolution
    // paints until the new one lands (no blur-up beyond the existing fallback).
  async units(): Promise<UnitSpec[]>,
  async search(query): Promise<overlay.Hits>,   // §11
  async copy(range): Promise<string>,           // §11
  close(): void,
}
```

- Replay is 1:1: the replayer builds the canvas at the recorded scale bucket
  (server-declared `scalesBuckets`, §6.2) and never re-roots; the visible resolution
  always names its bucket.
- **Policy conjunction**: the effective policy is the AND of the
  `X-OfficeView-Protection` header (if the server sent one on the manifest) and the
  manifest's `can.*`.
- **Visible watermark ownership**: the server baked it; the client **ignores** the
  host's `watermark` prop and never double-paints a header mark (§9.4).

### 8.4 Fonts, backends and measurement seams (round-2 fixes)

- `officeview/core` exports the `RegisterFont` seam (`register.ts:9`) and adds a
  **backend/measure-context factory** in `PaintOptions` for the server:
  `measureCanvas?: () => CanvasRenderingContext2D` replacing today's hard-coded
  `import('canvas')` in `measurerFromDoc` (`paint.ts:184–211`) and in the index
  builder (`search.ts:212–215`), plus a `PaintOptions.devicePixelRatio`
  (charts read a real DPR at `src/charts/paint.ts:767–768`; the server supplies the
  bucket instead of document DP). Stage 0 only needs the seam, not the choice
  (§18 item 1).
- `registerFont` registration is **process-global in node-canvas** and cannot
  unregister — pools keep a **per-worker embedded-font budget** (64 MiB [call]) and
  recycle a worker that exceeds it (paired with the §9.2 deadline retirement).
- Font faces actually in use per unit are declared server-side per §7.4; Stage 0's
  matrix records each backend's face resolution behaviour (fallback lists, hints,
  alternative face selection), published with the Stage 0 table (§16.1).

## 9. The server side

### 9.1 Package layout [call]

```
server/
  package.json          bun workspace member; depends on the root officeview
                        package (workspace protocol) + the canvas backend chosen
                        in §18 item 1 (declared as a direct dependency)
  src/
    http.ts             /open endpoint, Accept negotiation, typed 406 (§6.1)
    ws.ts               /session upgrade, op routing, backpressure, per-op authZ
    authorize.ts        deployment hook authorize(req) → {viewerId?, policy} (§13.4)
    tokens.ts           HS256 claims schema (§6.2), key loading, renewal
    audit.ts            sessionId → {viewerId?, ip, startedAt, keyGen, revoked?}
                        (attribution mapping; retention separate from §9.3)
    pool.ts             render worker pool + deadline retirement (§9.2)
    cache.ts            the one canonical cache (§9.3)
    record.ts           recorder facade, scratch snapshot service (§7.3)
    projection.ts       protected projection + asset snapshot hashing (§7.7)
    policy.ts           §6.4 content-policy gate (pre-layout)
    forensic.ts         §13
    detection.ts        §13.6 (CLI: server/scripts/detect.ts)
    xlsxTiles.ts        §12 tile painting via viewport reads
    limits.ts           §14 numbers
    config.ts           deployment key/flag loading, fail-closed
  test/                 (bun test) §16.3
```

Root `package.json` gains `"workspaces": ["server"]`; `server` runs `bun test` and is
not published.

### 9.2 The render pool

- Same seam shape as `WorkerRenderEnv` (`src/worker/office-worker.ts:42–47`) —
  `createCanvas/toBitmap/parse/paintables`; `toBitmap` is raster-or-record per job.
  **Round-2 fix:** the record job is *near-rasterless but not free* — it performs
  backing-context queries, holds scratch-surface snapshots (§7.3) and shaped text
  work; its memory is **reserved before allocation** like any other job.
- [call] defaults: one render worker per 2 vCPU, capped at 8; one parse worker;
  parse-once-per-hash; **global render reservation `512 MiB` per process** (not per
  worker); per-job decoded budget reserved up front; a unit deadline of 2 s handled
  by worker retirement (§7.3), not cooperative yield; per worker: 256 rendered
  pages or 64 MiB of registered fonts → recycle.
- **Pass interleaving (round-2 fix):** `setWatermarkMuted` is module-global
  (`src/core/watermark.ts:249`); until it is context-scoped [call], a worker runs a
  paint pass **or** an index pass, never interleaved; the scheduler serialises
  accordingly.
- Outstanding-unit policy (round-2 fix): at most 8 units in flight per session and a
  32-unit queue ceiling; further requests wait (backpressure surfaced by the
  existing progress UI). A request for units beyond the cap is refused with a typed
  rate message, not silently dropped.

### 9.3 The one canonical cache — and session/audit storage

**Cache key (round-2 fix; single form everywhere):**

```
docRef ⊕ layoutFingerprint ⊕ recorderVersion ⊕ scaleBucket ⊕ contentPolicyId
       ⊕ viewerMarkId ⊕ projectionMode ⊕ **unitId (subkey)**
```

- Per-unit records live here (pre-forensic), each entry holding one unit's
  commands/raster plus its shared asset table and culling proof — the `unitId`
  subkey (round-3 fix) makes the value per-unit explicit; the asset table dedups
  across units by hash inside the document layer of the same key.
- **XLSX tiles** share the same tuple plus `sheetId + tileRect` — no separate key
  shape (round-2 fix; draft 1 had diverging key formulas).
- `layoutFingerprint` covers **(officeview version, backend version, canvas font
  set, locale, parser inputs)** — the §14 byte-determinism claim (draft 1's
  under-specified wording) is now *always* conditioned on this fingerprint.
  A mid-session fingerprint mismatch ends the session as a typed
  `layout-changed` (not a repaint of stale work).

**Session tiers:**
- **Document cache** (per process): the decoded document, its text index, and the
  projection-relevant state (e.g. per-image painted-bitmap caches so cross-session
  hashing is not repeated). Single-parse admission; **nothing from this layer leaks
  into a session's wire bytes except filtered outputs** (§7.7).
- **Session edit tier**: the forensic-mutated command frames, bounded (≤ 8 MiB
  in flight [call]), dropped at session end.
- **Session table** (run state): `sid → {docRef, ip, startedAt, jti, status}` —
  lives for the session only, governs resume/revocation, evicts with the session.
- **Audit table** (round-2 fix, codex 7): the attribution mapping survives
  session teardown: `sid → {viewerId?, ip, startedAt, keyGeneration, policyId}`,
  governed by the owner's retention policy (default 90 days [call]); write-once,
  deleted-on-expiry, **never shipped, predicted or logged** (§19); present so the
  §13.6 recovery of a `sid` has somewhere to resolve months later.

**Tokens and per-op authorisation (round-2 fixes):**

- Token claims are the §6.2 schema; the **socket is validated once and the token is
  bound to the socket**, with `jti`/`sid`/`docRef`/`exp` re-checked on every op; the
  `can.*` flags are the server's claim — **the server enforces `can` on every op**
  (a capability the server set to `false` is refused at the op, never merely
  hidden in the client UI).
- A WS URL from the manifest is validated against a per-deployment **allowed-origin
  allowlist** [call] (default: same host). No open redirect surface into it.

### 9.4 Visible watermark (round-2 fix; owner's per-viewer read)

- **Default: per-viewer.** The audit table's viewer identity (§9.3) — or a
  session-derived label when the deployment has none — is rendered in the page
  raster server-side, in the same machinery as the existing mark
  (`paintWatermark`, `src/core/watermark.ts` painter + defaults at `:48–57`,
  header parser `:164`) with merge-style precedence merged server-side
  (`mergeWatermark` at `src/core/watermark.ts:114`).
- **Mark form depends on the projection (round-3 fix; one pipeline per projection,
  ordering fixed end-to-end):**
  - `final-raster`: paint the document frame, composite the mark into the final
    raster server-side, then assemble. The mark is pixels — a replayer cannot
    delete it without corrupting the only content it receives.
  - plain `commands` (unprotected only): the mark ships as a last-position,
    opaque `RASTER_BLOCK` overlay op (asset-table referenced). **Honest, and now
    stated plainly (round 4): the overlay is removable by a modified replayer —
    changing its representation from text to bitmap does not prevent dropping one
    op. Only `final-raster` guarantees the visible mark.** Command projections
    with an overlay mark are therefore never offered to protected documents.
  - `culled-commands` (protected, opt-in): the mark must be embedded by the
    **M-rule** — the mark bitmap is **composited into every transmitted content
    raster** (document image assets) before framing, and any transmitted asset
    that the pool computed unmarked is removed from the frame, so there is no
    separately removable mark op. A protected unit whose content has no such
    raster to composite into (pure vector commands) answers `final-raster`
    instead. The M-rule pipeline is stage-gated (§17); until its tests are green,
    `final-raster` is the only protected projection shipped.
  - the obvious false path — an overlay op as the only mark on a protected
    command projection — is **forbidden** by the null spot-check of §7.7 (P-rule
    consistency).
- **Ordering with the forensic pass (§13.2 pipeline)**, one order for every unit
  record: (1) document paint with text intact; (2) protected-projection culling
  proof over the command list; (3) forensic text-displacement mutate (needs glyph
  coordinates, so it runs where glyph runs survive — i.e. not on `final-raster`,
  which uses the image-domain mark of §13.2); (4) visible-mark compositing
  (raster-side on `final-raster`/M-rule units; an overlay op only on plain
  commands, unprotected only); (5) flatten/rasterise any remaining group, then
  frame. Nothing before step (3) may have baked pixels where the forensic mutation
  needs coordinates, and nothing after step (3) may exist as removable mark
  commands.
- A **fixed-text override** (e.g. `CONFIDENTIAL`) is an owner decision with a
  stated trade: it removes per-viewer *visible* identity, so it requires the §18
  item 8 approval, and the audit table stays the attribution record.
- A fixed mark lets one cache entry serve every session; per-viewer marks key the
  cache on `viewerMarkId` (§9.3). The **session-mutated frames live only in the
  §9.3 session edit tier** — the canonical tier holds the pre-mutation record and
  the culling proof.
- `X-OfficeView-Watermark` is not sent on manifest answers (the baked mark is
  authoritative; §8.3).

## 10. Behaviour detail: open, failure, resume, expiry

**(Open flow)**

1. Host navigates to `source` (or passes bytes / a Response — unchanged APIs).
2. `loadOfficeSource` → the prefix it holds is sniffed (§6.3): office bytes → the
   existing loader; `OVSM` → the manifest arm.
3. Manifest parses and validates (§6.2); the loader opens `session.wsUrl`, token
   delivery in the first WS frame (§9.3).
4. Server validates the token (signature, `jti`, `docRef`, `exp`, origin fence,
   optional `authorize` hook), creates or resumes the session, replies.
5. Client requests the units it needs (first visible page + next/prev prefetch
   within the §9.2 outstanding-units cap; scope the viewport budget on XLSX, §12).
6. Pool renders a cold unit → applies the protected projection (§7.7) and the
   marks → chunks over WS; the client replays 1:1 at the unit's scale bucket.

**(Failure, refusal, expiry)**

- 406 / bad token / refused upgrade / expired token while paginating → the client
  surfaces a load error state with the server-supplied notice (deduplicated, no
  retry loop on 406); the viewer is otherwise usable.
- **Renewal vs terminal** (round-4 fix for promises and Requests): token staleness
  follows the §6.3 classification — **renewable** only URL/HttpSource, or a
  body-unconsumed Request; a promise inherits its **resolved** source's class
  (`Promise<Response>` is terminal, §6.3). Renewable sources re-run open and link
  the fresh token's optional `refresh` claim into one audit history; **terminal**
  sources — bytes, Blob, Response, `Promise<Response>` and consumed Requests
  (re-reading replays the same expired token) — enter the typed `stream-expired`
  state and recover only through the viewer's reload affordance.
- Cancelled load / unmounted viewer / document changed → session `close` exactly
  once, with buffered chunks dropped atomically (never a half-painted unit).
- **Rendering failure on a unit** → tile fallback (§7.6), never hidden from the
  diagnostics; replayer-side frame validation failure (header/version/checksum)
  drops the unit and asks the server to re-send once; second failure → the
  missing-unit diagnostic, not a plausible paint.

## 11. Search, copy and semantic text services (round-2 merged)

- **Search** (`can.search`): client's query → server op against the **index built
  server-side in the §7.3 second pass** (existing `createRecordingContext` +
  `buildTextIndex`, `search.ts:92,315`) — with the visible mark muted
  (`search.ts:322–334`). The answer is highlight rectangles in the existing
  `HighlightRect` shape (`src/core/overlay.ts:7`) for the **visible projection**.
- **Copy** (`can.copy`): a server op for the selected range's *display* text —
  **visible-band text only** (round-2 rule): logical ranges that carry the whole
  source even when a band clip hides most of it (`src/core/text-recording.ts:11–14`)
  are intersectioned against the recorded band before serving; XLSX copy serves
  the formatted display text for the selected cells, not raw/formula values. The
  client's local text-selection machinery (`src/core/selection.ts`) is never used
  for a remote doc: selection is rects from the same index.
- **A11y/semantic text** (`can.textA11y`): the §15 service, from the same index —
  subject to the §6.4 policy (vanish runs and hidden content never appear).
- **The visual/index join (round-3 fix):** the index pass records the *full logical
  source* (`LogicalTextRange` retains it beyond the clip, `src/core/
  text-recording.ts:9–16`), so "what was recorded" is not "what is visible". Every
  semantic answer therefore runs one **join rule** before leaving the server: a
  text range is servable only if (a) it sits fully inside the unit's painted clip
  bands, (b) the §7.7 occlusion analysis from the culling proof shows no later op
  paints over any part of it — on `final-raster` the occlusion is structural in
  the raster, but final pixels identify neither which logical ranges were covered
  nor how much of a run remains, so the same conservative withhold applies there
  equally and by construction), and
  (c) it is not part of another surface's private text. A range that fails any
  condition — including *partially* clipped or overpainted runs and partially
  visible graphemes — is **withheld entirely** from search, copy and a11y, with a
  per-unit diagnostic count. When in doubt, withhold; tests assert covered and
  partially-clipped runs for every endpoint (§16.2).
- **Chart descriptions** (`describe.ts:40–62` reads model categories and numeric
  series values, not paint): in stream mode they are **visible-projection only**
  (§6.4): what's painted is available as text; nothing is emitted by reading the
  model. A Stage 1 test asserts a chart in a stream answer carries categories only
  when painted and never min/max ranges.
- Forensic watermark text never appears in search/copy/a11y answers (§13.4 rule).

## 12. XLSX in stream mode

- **Scope**: server-painted **tiles** of the *interactive* sheet viewport plus
  display-only metadata. This is the owner decision (§2 item 6) and the boundary:
  the local interactive path is where formulas (`contracts.ts:74–80`), hidden
  rows/cols geometry (`contracts.ts:668–673`) and raw picture bytes
  (`SheetDrawings.images` at `contracts.ts:362–366`, `XlsxImage data: Uint8Array`
  at `types.ts:27`, loaded by an `op:'drawings'` read,
  `src/worker/viewerRemoteDrawings.ts:1–18`) reach the client — **the stream mode
  ships none of these**:
  - Cell content stays server-side; copy/search/a11y are §11 server calls.
  - SheetLayout is reduced to a paint-facing subset (axis sizes, frozen panes,
    merged-rect and gridline extents only). **Round-2 honesty note:** the client
    still *sees the collapsed row/col numbering* of hidden bands — that reveals
    hidden bands exist (not their content). Stated as an accepted limitation,
    §20; if the deployment forbids it, the sheet renders without visible numbering.
  - Static pages: **draft-1's "already paint-only" claim is corrected** — the local
    path paints every sheet including hidden ones (`paint.ts:427–429`, §4 recon);
    in stream mode the §6.4 gate filters hidden sheets before layout.
- **Tiles** [call]: painted on the server pool by `getPaintables`' XLSX bring-up,
  keyed by the §9.3 canonical key + `sheetId + tileRect`; the tile bitmap is a
  drawn-bitmap snapshot (§7.7); the **viewport reply is display-only text-free** —
  a11y/copy/search are server calls on the recorded text index (§11).
- **Interactions are staged** (§17): 4a static + read-only sheet tiles;
  4b selection/freeze/sheet-navigation parity (client-side geometry, the server
  never sees selection except as requested ranges); **4c server-side sort/filter**
  (owner decision 2026-11-07, §18 item 3).
- **Server-side sort/filter (4c)**: the server adopts the existing engine surface,
  `op:'sort'`/`'filter'` (`src/worker/viewerProtocol.ts:31`, command kinds at
  `src/xlsx/view/contracts.ts:709–710`), running it in the pool against the
  server-held workbook. Rules:
  - **View state, not edits.** Sort/filter state is per session and never changes
    the source document or another viewer's view. The state lives in the session
    table (§9) and is dropped on session expiry.
  - **Cache key.** A tile painted under a sort/filter state adds a canonical
    `viewStateHash` (sort keys + filter criteria, normalised) to the §9.3 key;
    the unsorted, unfiltered view keeps the plain key so it stays shareable
    across viewers.
  - **Filter value lists are display text only.** The values a filter menu offers
    come from the formatted display text of cells in rows that are visible under
    the §6.4 content policy. Formulas, raw cell values, hidden rows/cols and
    hidden or veryHidden sheets never contribute values; a value that exists only
    in a policy-hidden row is not listed.
  - **Replies carry no cell data.** A sort/filter reply returns new axis geometry
    (row order and sizes for the paint-facing SheetLayout subset) and invalidates
    tiles; content arrives only as freshly painted tiles. Row-number labels follow
    the §20 hidden-band note.
  - **Rate limits.** Sort/filter ops are a separate rate-limit class in the §14
    table, because each one re-lays out the sheet.
  - **Wire audit.** 4c's acceptance adds the §16.2 wire-disclosure suite over sort
    and filter: a secret planted in a hidden row, a hidden column, a hidden sheet
    and a formula must not appear in any reply, filter value list or tile.

## 13. Forensic watermark — capacity, retention and collusion (round-2 rewrite)

### 13.1 What it is

A **per-session tracking layer baked into the recorded commands (or tiles)**
invisible at reading size: selected words in the page's recorded text are displaced
sideways by a sub-pixel amount. Recovering the token requires the detector's
re-render plus the candidate table (§13.6).

### 13.2 Per-page self-contained codeword (round-2 fix; replaces draft 1's broken 48-bit ×3 rule)

Capacity math draft 1 stated was incoherent (§21). Round 2 sized the codeword to a
page; round 3 completes it with range classification, a decision rule and the
projection-dependent pipeline placement (§9.4):

- **Where the token enters the pipeline (round-3 fix, per projection)**: text
  displacement needs glyph coordinates, so it runs on records that still carry
  glyph runs — plain `commands` (unprotected only) and protected
  `culled-commands` **after** the §7.7 culling proof (displacement cannot resurrect
  culled text; carriers are drawn only from servable, unclipped words of the §11
  join). On `final-raster` records there are no coordinates left, so the forensic
  mark there is the **image-domain mark**: a low-amplitude, keyed pixel-domain
  pattern (±1 on a keyed channel bitmask, tiled and interleaved) baked into the
  final raster, with its own encoded 40-bit codeword and its own honestly lower
  Stage-2 robustness targets (§18 item 7 decides whether it ships).
- **Token allocation** [call]: the 24-bit token is a monotonic counter drawn from
  the audit table, never reused while its audit row lives (not even after session
  end); exhaustion refuses new sessions rather than recycling.
- **Payload per page**: the 24-bit token + **8-bit page salt** (HMAC-derived, used
  to seed the carrier selection so different pages select different carriers) +
  **8-bit CRC-8** → a 40-bit codeword.
- **Page classification (round-4 fix — a page must hold the full codeword for the
  single-page gate to apply)**: **carrier-rich** (≥ **40** eligible words — a binary
  ±δ word carries one bit, so 40 bits need 40 carriers; the round-3 threshold of 32
  could not), **sparse** (16–39), **none** (< 16). Carrier-rich pages host the whole
  codeword themselves; sparse pages spool their bits across sibling pages with
  distance-k interleaving (the detector reassembles and votes across the document
  set); `none` pages fall back to the image-domain mark or carry nothing and are
  reported (§18 item 7).
- **Placement**: a carrier word's glyph-run x is displaced by `+δ` or `−δ`
  (binary), δ calibrated at Stage 0/2 (target < 0.5 CSS pixel). Carrier (and
  decoy) x values are written on a **1/256-px lattice** [call — round-3 fix, the
  quantisation now defined here instead of the undefined "§7.3 grid"] so decoder
  measurement noise is bounded.
- **Decision rule (round-4 fix — per-candidate distance alone is not an attribution
  bound)**: decode compares the *observed* carrier bit string against every
  candidate token's expected codeword for that page salt.
  1. **Sample integrity, not candidate self-checks**: the observed 40 bits contain
     the CRC-8 of the token+salt as data, so the detector recomputes the CRC over
     the *decoded sample's* 32-bit payload and requires it to match the observed
     CRC field. A candidate's own expected codeword trivially has valid CRC —
     checking it there proves nothing (round-3 defect).
  2. Distance threshold: candidate accepted only if its Hamming distance to the
     observed string is ≤ **6 of 40** [call].
  3. **Unique winner**: if two or more candidates pass at the minimal distance, or
     pass within 2 of each other, the page **abstains** — "ambiguous, no token
     recoverable" — never a tie-break guess.
  4. **Whole-search bound (the decision target)**: the claimed number is the
     *false-attribution probability per detection*, targeted ≤ **10⁻⁶ per page
     decision over the full declared audit-table size** [call]. The per-candidate
     uniform-random figure — Σ C(40,k) for k≤6 ÷ 2⁴⁰ ≈ 4.18×10⁻⁶, i.e. ≈ 70
     expected random matches among 2²⁴ nominal candidates — is printed here only
     as a *raw ingredient*: it demonstrates that the per-candidate binomial is NOT
     the decision bound, and the CRC-in-sample check, the unique-winner rule and
     the carrier-selection limits shrink it further. Stage 2 must **measure** the
     actual whole-search false-attribution rate over the declared table size,
     including **unmarked-fleet-page tests** (captures from pages that carry no
     session codeword; every one must abstain) — the per-candidate calculation is
     never presented as that rate.

### 13.3 Carrier eligibility and hardening (round-2 fix; merged from claude I7)

- **Candidate words** per page are `HMAC(kCarrier, docRef | unitId)`-seeded from the
  server-side recorded text index, restricted to: word length ≥ 4 grapheme
  clusters, on the body horizontal baseline, non-whitespace, and ≥ 3 word
  boundaries from any other carrier.
- **Decoy jitter** [call]: a second, keyed pseudo-random selection displaces
  non-carrier candidates by `±δ/2`. **Round-3 honesty:** with a reference or
  known base position the two amplitudes are distinguishable in principle, so
  single-session carrier concealment is a **Stage-2 hypothesis to measure** (mixin
  ratios, discovery rates), never a guarantee restated in product or docs; the
  protection against single-session readback comes from the carrier selection
  staying key-private (the §13.4 law), and the §13.5 collusion bound covers the
  two-stream case.
### 13.4 Key law (round-2 fixes for both M3 and codex 7)

- **One master, many subkeys** [call] (HKDF over `OFFICEVIEW_STREAM_KEY`, a
  32-byte key whose absence refuses startup, §9.1):
  `kToken` (§6.2), `kDoc` (docRef), `kCarrier` + `kJitter` (§13.3), plus a
  **key generation** byte per deployment era.
- `docRef` is `HMAC(kDoc, sha256)` (§6.2) — the plain hash is never on the wire.
- `sid` is random per session, **not** the viewer's identity; the audit table is the
  only place that binds a `sid` to a person (§9.3), and only when the
  `authorize` hook supplied one; otherwise it stays a collated-but-anonymous row.
- Only the **audit table** outlives a render session; the session *run* table
  still evicts at close (§9.3), so attribution survives teardown through the audit
  row, not the session row.

### 13.5 Collusion — stated as a real limitation (round-2 fix)

Draft 1's "collusion between two holders reveals nothing" was false: comparing two
sessions of the same document *does* reveal which word positions changed, and
averaging opposite-offset words cancels the signal. With §13.3's decoys the
distinguishing set shrinks to those that differ across streams, which is still a
leak — so the honest bound is:

- **Two coordinated holders can estimate the shifted word set**; that reveals which
  positions are candidates, not the payload bits (decoys sample deltas of both
  polarities).
- **The payload is not recoverable from carrier positions alone** across holders
  without the audit table (a token is only meaningful there).
- Stage 2 *measures* this (a two-stream diff harness with N colluding readers) and
  publishes real numbers instead of folklore; whatever the measured bound is,
  it is written into `docs/validation/streaming.md`.

### 13.6 Detection tool and single-page gate (round-2 fixes)

`server/scripts/detect.ts`:

```
bun scripts/detect.ts --images page1.png [page2.png …] --key-file ./ov-key.json \
  --doc <file.docx> [--page N] --ids ids.json   # or the audit-table export
```

Per image, deterministic:

1. Re-derive the **reference frame** through the same pool (§9.3 key; never a
   screenshot; cache miss renders on demand).
2. Coarse alignment (a page-border/frame similarity transform), then **per-word
   normalised cross-correlation** of candidate neighbourhoods (with vs without ±δ);
   the local argmax gives the per-word displacement sample.
3. Per-page codeword: decode by max-likelihood over the candidate session-token
   list (with/without the §13.3 decoy filter); consensus across pages wins; the
   audit table resolves the full `sid` if a list is supplied (never written outside
   the operator machine, §19).
4. Confidence: carrier hits, Δx residuals, CRC agreement; **a below-threshold
   detection reports "no token recoverable", never a guess**; if the audit table
   is present and a token resolves, output with the resolve's viewerId, plus
   the key generation.

**Stage-2 gates (numbers, not folklore; scoped round 3):** ≥ **95 %** single-page
recovery from a default-zoom screenshot **of a carrier-rich page** (§13.2
classification — sparse pages are excluded from this gate and are covered by the
cross-page reassembly test with its own published rate); ≥ **80 %** after JPEG
re-compression at quality 70; published per-class degradation below that, never
silent.

## 14. Performance, memory and limits

Wire budget [call, to verify in Stage 0 over the corpus]: text-heavy page ≈ tens of
KB gzip; image-heavy slide adds asset refs only (the **drawn-bitmap snapshots** of
§7.7, not original files); units the recorder cannot express fall to tiles
(measured, not assumed). Per-session/per-page caps [call]:

| Limit | Value [call] | Round-2 note |
|---|---|---|
| unit pixel size ceiling (`maxUnitPx`) | 8192 (manifest cap + validation) | closes the `2..20000`-vs-4096 mismatch |
| decoded-image bytes per unit | 64 MiB | *decoded* caps, reserved before allocation — encoded-asset caps do not bound memory (round-2 fix) |
| inflated command bytes per unit | 8 MiB | reserved before the pool builds, not after |
| WS frame chunk | 1 MiB | with one in-flight paint per unit backpressure per session |
| assets per session | 256 MiB total | streamed by hash once |
| outstanding units per session | 8 (queue 32) | §9.2 backpressure |
| manifest decoded size | 256 KiB | bounded parse before any allocation (round-3 fix) |
| unit-body checksum | CRC-32 over the gzip-deflated unit body + a 32-bit sequential frame id | the §10 reassembly/re-send drops a unit whose checksum or sequence mismatches (defining what draft 2 only referenced) |
| render deadline | 2 s → worker retirement | §7.3; sync paint cannot yield |

Pool / budget split [call]: `512 MiB` reservation is **global per process** (not per
worker); doc cache 1 GiB, render cache 512 MiB, session tier ≤ 8 MiB in flight.

Rate limits — named keys and numbers, not adjectives (round-3 fix; a per-session
queue does not bound a fleet of sessions):

| Key | Limit [call] | Burst |
|---|---|---|
| `ip` (unauthenticated) | 30 `/open` per rolling minute | 8 concurrent opens |
| `sid` (per session) | 120 WS messages per rolling 10 s | 20 messages/s |
| `sid` (unit paint requests) | 60 unit requests per rolling minute | in-caps only |
| process socket bound | 256 concurrent sockets, 4,096 queued units fleet-wide | new connections delayed, never silently dropped |
Client [call]: replay cost ≈ today's interactive paints; the loader gains one 12-byte
peek; DOM pressure *drops* (a 20 MB DOCX: local mode holds every byte resident, remote
mode holds at most a couple of unit-sized command buffers).

Determinism is the contract test's foundation, **conditioned**: same doc bytes +
same `layoutFingerprint` inputs (officeview version, backend version, canvas font
set, locale, scale bucket) ⇒ byte-identical recording. The recorder reads no clock
(the optional visible-mark timestamp arrives as a paint parameter, never read
in-pipe), and never reads `Date.now()`/random from the render path.

## 15. Accessibility

- **Server text service** (`can.textA11y`) returns text for coordinates from the
  `RECORD_TEXT` index (§7.3 second pass), joined through the §11 rule — so the
  offscreen page text carries exactly the text the content policy (§6.4) and the
  visibility join permit, **which may be less than local mode shows** (covered,
  clipped and vanish runs are withheld, not disclosed); it holds exact locale/BiDi,
  not OCR.
- Visible-mark text never appears in a11y/search/copy (§13.4). The forensic mark is
  geometric, and the §7.7 projection keeps it in the raster only — **not** in any
  text answer.
- Focus, keyboard and page-announcer surfaces are today's OfficeDoc surfaces (§4).
- `can.textA11y = false` (server policy): the client states
  **"text unavailable in this viewing mode"** in its existing a11y notice surface
  rather than pretending.

## 16. Testing and validation

### 16.1 Stage 0 gates (round-2's acceptance list, adopted verbatim as the gate)

Before any wire implementation is written, the Stage 0 matrix closes these:

1. **Backend matrix** (§18 item 1) — pixel/direct equivalence and resource behaviour.
2. **Direct-vs-record+replay at multiple scales** — zoom × DPR grid
   **{0.5, 0.75, 1, 1.25, 1.5, 2, 3} × {DPR 1, 2}**, including **shadows,
   absolute/setTransform-headed painters, pixel-snap paths, transformed patterns,
   and scratch-surface drawImage** (the §7.3 census is only true if these
   reproduce); fractional zoom proves the §8.3 1:1 replay rule.
3. **Proxy census over the corpus** — the recorded op subset is completed
   experimentally (round-2's census gate), giving the §7.5 op list final meaning.
4. **Glyph-extraction decision** — whichever of §7.4's (a)/(b) ships — including
   shaping/ligature/fallback/BiDi fixtures and a scratch-rendered text case.
5. **Per-page carrier count** measurement (§13.2, §14).
6. **Asset-snapshot correctness** — cropped/duotoned images inspectable
   (§7.7/§16.2), no original byte leakage in any sampled record.
7. **CORS simple-request check** — the `Accept`-only negotiation keeps local
   cross-origin loads simple (§16.3).
8. **Eager core-entry split measurement** (§8.1 gates).
9. **WS transport scaling probe** — 1,000 concurrent unit requests hold the
   per-process caps (§14).

### 16.2 Unit tests (vitest officeview; bun test server)

- **Wire inspection boundary tests** (round-2's rule): for each stream-mode answer,
  decompress all frames/assets/tiles and assert a **fixture secret string** (repeated
  in a hidden sheet, `w:vanish` run, comment, note) does not appear in any of:
  pixels, text index, copy, a11y answer, search, the manifest unit list or sheet nav.
- Recorder fixtures for **each op family**; unknown ctx method ⇒ tile promotion with
  a diagnostic, never a silent skip (§7.6).
- **Proxy census** corpus run validates the subset (round-2 gate).
- Forensic: injection/extraction round-trips on synthetic pages; interleaving across
  sparse/pure-faces/carrier-poor fixtures; decoy-jitter distribution checks;
  key-absence fail-closed; the two-stream collusion harness (§13.5) publishes the
  measured bound.
- Loader: **one-byte-chunked and truncated `OVSM`** parse fail-closed (round-2 gate);
  abort; source replacement; expiry on non-reopenable sources (typed
  `stream-expired`); ordinary file preservation when no magic (round-2 gate).
- **Per-endpoint visibility tests (round-3 fix):** a partially clipped logical run
  and a fully covered run are each asserted **absent** from search rects, copy
  output and a11y text on every format (the §11 join rule), and a partially
  visible grapheme is withheld — pixel-byte scanning is never the assertion.
- **Expiry tests for every source shape (round-4 fix):** manifest as bytes, as
  Blob, as URL, as Response, as `Promise<Response>` and inside a body-bearing
  Request — each observes its §6.3/§10 rule (terminal `stream-expired` for
  bytes/Blob/Response/promise/consumed Request; renewal only for URL/HttpSource and
  the body-unconsumed Request).
- Tokens/limits/backpressure/cancellation at the §14 numbers; the
  worker-retirement deadline path with a producer that ignores abort and one that
  honours it (session ownership retained until work drains, then released).
- Detection tool: synthetic draws (plain, JPEG q95/90/80/70, upscale 2×, crop 5 %)
  with published per-case outcomes.

### 16.3 Browser tests (Playwright, Chromium)

- Live against the **in-repo test server** (`server/test/app.ts`) wired by the
  Playwright harness (the existing `scripts/harness/fixture-server.ts` pattern).
- Manifest detection (mocked, then real), the §16.2 pipeline in-browser, server-call
  round-trips, print refusal under `no-print`, session expiry mid-view → open-recover,
  406 display (round-2 explicit), resume after network drop, **cross-origin local
  load stays a CORS simple request** (round-2 gate).
- Replayed-unit goldens only (no local baselines).

## 17. Staging with acceptance criteria (round-2 tightened; M4)

Each stage is its own PR; gates per §16; primarily measured, not narrated:

| Stage | Scope | Done when |
|---|---|---|
| **0: prototypes** | §16.1 gates, incl. the backend matrix + the projection/§13.2 carrier measurements + the §7.4 mode decision | §16.1 gates 1–9 resolved **with published numbers** in `docs/validation/streaming.md`; §18 items that depend on data are recorded |
| **1: core stream MVP** | core entry + gate; `OVSD` recorder+replayer; OVSM + sniff at `loadOfficeSource`; RemoteDoc rendering of DOCX/PPTX/*visible-sheet* XLSX pages; demo Bun server (open + WS + pool + per-viewer visible mark + token + 406); typed loader errors | §16.1 all gates green; corpus contract test green under the §7.4 wording for the chosen mode (zero-diff only where that mode allows the claim); Playwright happy path; 406 + **stale-version** + **byte-chunk** tests pass; eager diff ≤ 2 B gzip, module map identical apart from the sniff; **the P-rule projection (§7.7) and the wire-disclosure suite (§16.2) are green before any protected response ships** — protected streaming does not start in Stage 1 without them (round-3 fix: they were wrongly deferred to Stage 2) |
| **2: governance + forensics** | tokens/sessions, rate limits (§14 table), resume/revocation/TTL, cache tiers, §13 with decoys + detection tool; the M-rule mark-embedding pipeline for protected command output (§9.4) | rate-limit numbers asserted in tests; §13.6 gates measured and recorded (single-page recovery on carrier-rich pages ≥ 95 %; JPEG q70 ≥ 80 %; whole-search false-attribution bound ≤ 10⁻⁶ verified by simulation over the declared audit-table size, plus unmarked-page abstain tests); audit-table retention + revocation + key-rotation tests; per-class (sparse/none-page) fallback gates per §18 item 7; protected `culled-commands` may ship only after its M-rule tests are green (until then `final-raster` is the only protected projection shipped) |
| **3: server text services** | search/copy/a11y WS ops through overlays; visible-band copy and visible-projection chart text; per-viewer mark default on | local feature parity measured by the §11 service vs local query on the same corpus; zero local index building for remote docs; §6.4 policy tests (secrets absent from every service) |
| **4: XLSX stream mode** | 4a static + sheet tiles; 4b selection/freeze/sheet-nav parity; 4c server-side sort/filter (§12, §18 item 3) | hidden/veryHidden/vanish/token/bytes **never cross the wire** (wire audits); tile p95 render ≤ 120 ms pool-side at §14 budgets on a 500-row × 100-col sheet; selection parity passes against local mode; 4c: sort/filter results match local mode's row order on the corpus, filter value lists exclude policy-hidden values, and the §12 sort/filter wire audit is green |
| **5: optional extensions** | stream slideshow animation (§18 item 4: deferred by the owner) | only if the owner reopens it |

## 18. Open decisions for the owner (round-2 refresh)

1. **Canvas backend — decided 2026-11-07: Skia via `@napi-rs/canvas`** for the server
   pool. node-canvas stays where it is today (the test/golden path) and is not the
   server backend. Stage 0 still runs the matrix, now to *confirm* the pick rather than
   choose: pixel parity vs browser canvas, Bun compatibility, font registration scope
   and licence (to verify in Stage 0). If a gate fails, Stage 0 reports it back to the
   owner instead of switching backends silently.
2. **Shaping/outlining dependency** — if §7.4's (a) is picked (my expected default
   after Stage 0), name the lib (HarfBuzz-wasm vs fontkit-class) [call]; rejected
   alternative: staying dependency-free and guessing outlines — recorded as a
   round-2 rejection.
3. **XLSX interactivity depth — decided 2026-11-07: include server-side sort/filter.**
   It ships as Stage 4c under the §12 rules (per-session view state, display-text
   filter lists, no cell data in replies, own rate-limit class, wire audit).
4. **Slideshow in stream mode — decided 2026-11-07: no animation for server
   rendering for now.** `present()` is refused for remote documents; slides are
   viewed as static pages. The local engine drives its own clock and needs the
   local package handle, so a server variant would be new surface; it stays in
   Stage 5 until the owner reopens it.
5. **How documents are marked stream-required** [call]: by path, by content kind,
   by digest rule — one concrete config surface; my call: **path-glob + sha-by-list**
   (a document rarely changes identity within a release workflow).
6. **Print in stream mode**: pool re-renders print units through the recorder
   (costs capacity) or declares `no-print`; **my call: allow by manifest flag in
   Stage 3**, print-to-PDF from a stream is a real user need and cannot be faked
   client-side from raster tiles alone.
7. **Forensic fallback for carrier-less documents** (§13.2): (a) ship the image-domain
   invisible mark (low-amplitude pixel-domain pattern baked into the raster) — my
   expected default, and **the default for `final-raster`-projected units, which have
   no glyph coordinates for text displacement** (round-3 note); or (b) explicitly
   refuse/defer the forensic promise with an owner-agreed notice (the honest
   alternative). Stage 0's carrier measurement decides between them.
8. **Visible-mark override to fixed text** (round-3 fix): the default marks each page
   per viewer (§9.4). A deployment that wants a fixed `CONFIDENTIAL`-only mark is an
   owner decision with a stated trade — per-viewer identity moves entirely to the
   audit-aligned forensic grid, and for carrier-less documents the promise degrades
   per item 7 — and the audit table remains the attribution record; overriding the
   default and stripping the forensic mark together is not a supported configuration.
9. **A11y service default — decided 2026-11-07: on.** `can.textA11y` is on by
   default and any deployment turning it off ships the §15 notice.

## 19. Boundaries

**Always**

- The server owns the mode decision, from §6.1 negotiation through §6.4 content
  policy to §7.7 projection. The loader owns only the sniff; no client surface may
  set or override any of it.
- Fail closed: unknown opcode, unknown manifest version, unavailable backend,
  refused unit, unresolvable token → refuse the unit/session; never a
  plausible-but-wrong paint; never a fallback fetch.
- Every rendering artefact carries the exact
  `(docRef, layoutFingerprint, recorderVersion, scaleBucket, policy)` provenance
  the manifest promised (§9.3).
- Wire **audits inspect transmitted payloads**; screenshot-only assertions are
  never the evidence that a boundary holds (§16.2).
- Honesty in docs: screenshots possible; the module README and validation doc say
  so verbatim (§20).

**Ask first**

- Shipping **font geometry** derived from a face the document only *embedded*
  (outline extraction is paint input, not redistribution, but licences differ): when
  body faces are only runtime-licensed, §7.4's outlining falls back per unit; the
  document's own embedded faces carry no such question (to verify in Stage 0).
- Shipping **pinned font data** server-side (§7.4 mode (b)) — same ask first,
  per face.
- Persisting sessions beyond a process restart, and any export of the audit table
  beyond the operator's machine (logs included); making `textMode:'text'` the
  default.
- Serving a stream by default (without opt-in) for documents that previously came
  as bytes.
- Any change to `sessionTable`/audit retention beyond the owner's written retention.

**Never**

- Send a stream-required document's bytes, or any member/media/font bookkeeping
  bytes from it, to the client through any endpoint — including debug, internal or
  diagnostic responses (the ban is scoped to stream-required responses; an
  unprotected byte answer remains an ordinary server choice, §6.1).
- Ship original encoded image/media bytes or metadata-bearing replicas (§7.7).
- Answer a **stream-required** client **without** the mandatory final-visibility
  projection (the P-rule, §7.7): covered, clipped-out content and full z-order
  history never reach the wire — by final-unit rasterisation by default or by a
  stored culling proof, never by configuration alone.
- Ship hidden content in any transmitted payload (§6.4, always).
- Log or emit in telemetry `OFFICEVIEW_STREAM_KEY`, a session token, any `sid`, or
  any audit-table row beyond what the owner configured (§9.3, §19).
- Pretend DRM. No copy of "protects against screenshots" in product copy, README,
  or tests (§20); "screenshot-proof" claims are a docs bug.
- Execute anything a client posts inside the render pipeline — commands and tokens
  are the only inputs, and neither is ever treated as content, markup or code.

## 20. Honest limits

- **Screenshots, cameras, and determined readers**: any visible pixel can be
  captured; the forensic watermark's character is *traceability* after
  unauthorised resharing, not prevention. The promise is precise: a session token
  survives ordinary recompression and resize on most pages, is recoverable from a
  single default-zoom page screenshot at the published rate, and the deployment can
  then act on the audit row. It claims nothing about crop, manual redaction,
  deep laundering, or two-session collusion beyond the §13.5 bound.
- **XLSX hidden-band shape leaks**: collapsed row/col extents tell a reader that a
  hidden band exists. The deployment accepts this shape (§12) or gives up streamed
  axis numbering — the two options are named, not hidden.
- **The visible mark's guarantee is projection-specific.** On `final-raster` the
  mark is inseparable pixels. On any command projection the mark is either an
  overlay op (unprotected documents — removable by a modified replayer, stated as
  such, because the command projection already trusts the replayer) or, on
  protected `culled-commands`, embedded into content rasters by the §9.4 M-rule
  and made *stage-gated*: until that pipeline is tested, `final-raster` is the
  only projection that ships with a mark-guarantee for protected documents
  (round-4 statement; draft 2–3 overclaimed here).
- **The audit table needs a key generation**; rotating the master key invalidates
  captures, so the detection tool names which generation covered a capture and the
  rotation is recorded deploy-side (Stage 2). Two deployments sharing a key cannot
  distinguish their tokens' classes (not a security break; a traceability one).
- **Single-page documents** can carry the §13.2 codeword only when the page meets
  the carrier budget; otherwise the honest fallback is the image-domain mark or,
  if the owner refuses that (§18 item 7), an explicit "this document is not
  traceable" notice. Better to tell the deployer than to pretend.
- **Local mode is untouched by this design**: nothing is gated by stream code, and
  stream-only code is lazy so the eager graph cannot move it (§8.1's measurements
  make that a gate rather than a hope).
- The pixel-parity promise is specific to a (backend, font set, scale bucket)
  arrangement the deployment pins; §16.1's Stage 0 matrix measures what the
  owner's deployment actually ships, and §7.6 guarantees a wrong pixel never
  ships as right.

## 21. Review dispositions (round 1)

Codex 1–11 (its finding 1 is Critical, 2–10 Important, 11 Minor) and claude
C1–C5, I1–I10, M1–M7. "Fixed" = folded into this draft; "Rejected" = the finding
was verified and then not accepted, with reason.

| Finding | Disposition |
|---|---|
| codex 1 (original assets + reversible painting violate the visible-only boundary) | **Fixed (draft 2), superseded round 3** — draft 2 added drawn-bitmap snapshots, source-rect crops and an optional flattened projection; draft 3 makes the projection mandatory with the P-rule (see round-2 table, codex-r2 1) |
| codex 2 (XLSX + semantic services lack a common disclosure projection) | **Fixed** — §6.4 content policy + §11 visible-projection rules (chart text, visible-band copy, hidden-sheet exclusion) + §16.2 disclosure tests |
| codex 3 (one resolution-independent recording cannot preserve transform-dependent paint) | **Fixed** — §7.3/§8.3: record and replay at named scale buckets; §16.1 gate 2 multi-scale matrix incl. shadows/absolute transforms; charts get a PaintOptions DPR seam (§8.4) |
| codex 4 (context contract omits operations required during paint) | **Fixed** — §7.3 backend-backed facade: measureText via backing ctx, scratch-surface snapshots, wrapper patterns with recorded `setTransform`, `has`-trap fidelity, §16.1 gate 3 census |
| codex 5 / claude I1 (visual recording vs text indexing conflated) | **Fixed** — §7.3: the recorder does not expose `RECORD_TEXT`; the text index is a separate pass (existing buildTextIndex); painters' branching preserved; agreement tested (§16.2) |
| codex 6 (round-1; watermark defaults + forensic fallback) — partially related to claude C3/I7, **not** claude I6 (which is the CORS finding) | **Fixed** — §9.4 per-viewer by default; §13.2 per-page codeword + cross-page-sharing; §18 item 7 image-domain fallback; round-3 refinement in the second table (codex-r2 2,3) |
| codex 7 (collusion + retention claims unsound) | **Fixed** — §9.3 audit table kept separate from session state; §13.5 states the collusion limit and demands measurement; §13.6 includes confidence/abstention; `authorize` hook (§9.1) binds viewer identity where available; retention independent of cache eviction |
| codex 8 / claude I5 (loader migration + renewal; pullStream push-back) | **Fixed** — §6.3: sniff in `loadOfficeSource` on its existing prefix; typed 406 via a bounded non-OK reader; typed terminal `stream-expired` for non-reopenable sources; one-byte-chunk and truncated-`OVSM` tests |
| codex 9 (admission/cancellation insufficient) | **Fixed** — §14 table: decoded/inflation reservations before allocation, outstanding-unit/queue caps, worker-retirement deadline semantics; §16.2 probes |
| codex 10 (wire/auth/cache contract decisions) | **Fixed** — §6.2 normative schema + token claim table; negotiation via `Accept` parameters (no custom header, claude I6); unit-level tile discriminator instead of pre-declared `fallbackTiles`; §9.3 one canonical cache key incl. policy; per-op `can.*` enforcement; framed reassembly rules |
| codex 11 (misleading citations; qualify contradictions) | **Fixed** — §4/§19 anchors corrected (see M1); §19's byte ban is scoped to protected/stream-required responses; §8.2 names `getPaintables` as the single entry point |
| claude C1 (hidden content painted today) | **Fixed** — §6.4 policy + §16.2 secret-string-per-fixture tests across pixels/index/copy/a11y/manifest; draft 1's "XLSX no special work" claim withdrawn in §4 |
| claude C2 (original image bytes on the wire) | **Fixed** — same as codex 1 (merged: drawn-bitmap snapshot, session-scale, metadata-free, pixel-hash key) |
| claude C3 (forensic capacity math incoherent) | **Fixed** — §13.2 per-page codeword (initially draft 2's cross-page rule; **superseded in round 3** by the carrier-rich/sparse/none classification and the concrete decision rule — see round-2 table, codex-r2 3; threshold corrected round 4 per codex-r3 1) |
| claude C4 (outlines vs zero-diff gate; "zero dependencies" overclaim) | **Fixed** — §7.4 rewritten: mode (a) with a named shaping lib + outline-rendered reference + tolerance, or mode (b) pinned server fonts with the real zero-diff claim; "no new dependencies" scoped to the library (§3.2/§18 item 2) |
| claude C5 (device-space-dependent painting) | **Fixed** — same as codex 3 (merged: quantised scale buckets in every request + canonical key; low-res retained during zoom re-request) |
| claude I1 | **Fixed** — see codex 5 |
| claude I2 (semantic services exceed what's drawn) | **Fixed** — §6.4 + §11 visible-band text and visible-projection chart descriptions |
| claude I3 (op subset incomplete; scratch surfaces; pattern setTransform; `has` trap) | **Fixed** — §7.3/§7.5 (subset extended, real backing facade, Stage-0 census; `fillText` maxWidth and `strokeText` included) |
| claude I4 (backend seam; node-canvas font registration global) | **Fixed** — §8.4 `measureCanvas` factory + per-worker font budget/recycling; §9.2 |
| claude I5 | **Fixed** — see codex 8 (merged) |
| claude I6 (custom header forces CORS preflight) | **Fixed** — §6.1 `Accept` media-type parameters only; §16.3 simple-request test |
| claude I7 (replay can drop the visible mark; shifts detectable from advances) | **Fixed** — §13.3 keyed decoy jitter + quantised coordinates; mark baked into the raster before recording (§9.4), so replay cannot drop it; limits restated (§13.5, §20) |
| claude I8 (no viewer identity/auth hook/WS token transport; missing "how docs are marked" decision) | **Fixed** — §9.1 `authorize` hook; §9.3 token transport in the first WS frame + origin allowlist; §18 item 5 names the marking surface as an owner decision |
| claude I9 (cache keys + precedence contradictions) | **Fixed** — §9.3 single canonical key (incl. DPR/scale bucket, policy, mark) used identically for tiles; §9.4/§8.3 precedence: server-baked mark, `can.*` AND protection header |
| claude I10 (core entry could change the eager graph) | **Fixed** — §8.1 measured gate (module map unchanged apart from the sniff; ≤ 2 B gzip delta) |
| claude M1 (wrong citations) | **Fixed** — anchors corrected per the codex audit (§4); the cited `workerProtocol.ts` filename never existed (orchestrator-confirmed round 3; round 2 wrongly "rejected" this): the real symbol is `src/worker/viewerProtocol.ts:30–33` (`VIEWER_COMMAND_KINDS` incl. `'sort'`/`'filter'`; command kinds at `src/xlsx/view/contracts.ts:709–710`), cited correctly in §12; `RECORD_TEXT` is at `src/core/text-recording.ts:2` (citations updated) |
| claude M2 (format inconsistencies) | **Fixed** — §6.2/§7.1/§7.5: binary OVSM (no `+json`), `gzip` naming everywhere, `TILE` bytes inline (no data URIs), `1..8192` px bound aligned with `maxUnitPx`, f64 coordinates named as the encoding |
| claude M3 (plain sha256 hash lets a holder confirm a match; one key too many) | **Fixed** — §6.2 `docRef` (HMAC, truncated); §13.4 HKDF subkeys + key generation labelled in the audit table |
| claude M4 (untestable criteria) | **Fixed** — §14 numbers (rate limits, caps) + §17 done-when gates with concrete values/percentages |
| claude M5 (version negotiation unclear; redundant recorder header) | **Fixed** — §6.1/§7.6: negotiation via `Accept` parameters; the custom `X-OfficeView-Recorder` header removed |
| claude M6 (`setWatermarkMuted` is module-global) | **Fixed** — §9.2: workers never interleave a paint pass with an index pass until the mute becomes context-scoped [call] |
| claude M7 (row/col numbering reveals hidden bands; hidden sheet names) | **Accepted** — hidden/veryHidden sheet *names* are filtered from sheet nav by §6.4; numbered-axis reveal is an accepted limitation stated in §12/§20 with the render-without-numbering alternative offered |

### Round-2 dispositions (codex r2, findings 1–8, draft 3)

| Finding | Status in draft 2 | Disposition in draft 3 |
|---|---|---|
| codex-r2 1 (protected projection still optional; group flattening does not establish final visibility) | PARTIAL | **Fixed** — §7.7 rewritten as the mandatory P-rule, stated once: `final-raster` by default (whole-unit final frame, cross-group covers removed by construction) or proof-gated `culled-commands`; per-group flattening removed; §§3.1/3.3/19/16/17 reference the single rule |
| codex-r2 2 (watermark baking / rasterisation / post-record mutation ordering incoherent) | PARTIAL | **Fixed** — §9.4 defines one pipeline per projection with a fixed ordering (paint → culling proof → forensic mutate → mark composite → flatten); the mark is never mark-recorded-as-commands on protected documents; image-domain mark on `final-raster`; mutated frames live only in the §9.3 session tier; fixed-text override needs §18 item 8 approval |
| codex-r2 3 (forensic coding incomplete; single-page gate contradictory) | PARTIAL | **Fixed** — §13.2: page classification (carrier-rich ≥ 32 / sparse / none), token allocation (monotonic, never reused), decision rule (per-candidate compare, distance ≤ 6 of 40 + CRC-8, stated false-attribution bound calibrated in Stage 2), single-page gate scoped to carrier-rich pages; §13.3 quantisation defined (1/256-px lattice), decoy indistinguishability demoted to a measured hypothesis |
| codex-r2 4 (semantic visibility lacks the join; §15 overclaims local parity) | PARTIAL | **Fixed** — §11 visual/index join rule: servable only if fully inside clip bands AND unoccluded per the §7.7 proof; partial graphemes/clipped/covered runs withheld entirely with diagnostics; §15 wording corrected ("may be less than local mode shows"); §16.2 covered/partial tests per endpoint |
| codex-r2 5 (bytes/Blob renewal cannot refresh credentials; refresh claim absent from schema) | PARTIAL | **Fixed** — §6.3 rule 5 + §10: sources classified renewable (URL/HttpSource/Request) vs terminal (bytes, Blob, Response, ReadableStream); `refresh` claim added to the §6.2 table as optional/renewable-only; bytes/Blob/URL/Response expiry tests added |
| codex-r2 6 (retirement refunds live allocations; transport limits unspecified; OVSM length/checksum undefined) | PARTIAL | **Fixed** — §7.3 drain-then-refund retirement; §14: rate-limit table with named keys and numbers (ip/session/process bounds), manifest decoded-size cap, CRC-32 + frame-sequence checksum definition, OVSM length named as decoded length with a single compression-semantics flag |
| codex-r2 7 (fidelity gates contradict the chosen references; protected MVP shipped before boundary suite) | PARTIAL | **Fixed** — §3.1/§7.4: one fidelity gate per text mode (outlines graded vs outline-rendered reference with published tolerance; text mode zero-diff only over the §7.4 face handshake, else per-unit raster fallback); Stage 1 cannot ship a protected response before the P-rule projection + wire-disclosure suite are green (moved from Stage 2) |
| codex-r2 8 (workerProtocol rebuttal wrong; RECORD_TEXT anchor; `can.semantic` absent; cache key lacks unit identity; disposition pairing) | PARTIAL | **Fixed** — §21 round-1 M1 row corrected to "fixed" (orchestrator-confirmed: `workerProtocol.ts` never existed; `viewerProtocol.ts` cited in §12); `RECORD_TEXT` anchors moved to `text-recording.ts:2`; `can.semantic` became `policy.semantic` in the §6.2 schema; `unitId` subkey added to the §9.3 canonical key; codex-6/claude-I6 pairing re-labelled |

### Round-3 dispositions (codex r3, findings 1–5, draft 4)

| Finding | Status in draft 3 | Disposition in draft 4 |
|---|---|---|
| codex-r3 1 (40-bit codeword cannot fit the ≥32-carrier class; attribution rule conflates per-candidate and whole-search bounds) | PARTIAL (R2-3) | **Fixed** — carrier-rich threshold raised to ≥ 40 eligible words (a 40-bit binary codeword needs 40 carriers; 32 could not hold it), sparse 16–39, none < 16; the decision rule now states the whole-search target (false-attribution ≤ 10⁻⁶ per page decision over the full audit table, measured with unmarked-page tests) and prints the per-candidate Σ C(40,k≤6)/2⁴⁰ ≈ 4.18×10⁻⁶ figure only as a raw ingredient, never as the bound; the CRC moved to the *observed sample* (a candidate's own codeword trivially passes), unique-winner abstention added, ambiguity abstains |
| codex-r3 2 (command projection's overlay mark is separately removable; contradiction with protected culled-commands) | PARTIAL (R2-2) | **Fixed** — §9.4 now states plainly that an overlay mark is removable and only `final-raster` guarantees the visible mark; protected `culled-commands` output must satisfy the M-rule (mark composited into every transmitted content raster, unmarked assets removed) and pure-vector units answer `final-raster`; the M-rule pipeline is stage-gated, so until tested `final-raster` is the only protected projection shipped; §20 records the limit verbatim |
| codex-r3 3 (face handshake claimed but undefined) | PARTIAL (R2-7) | **Fixed** — §7.4 now defines the concrete protocol: `requiredFaces` identity list (hashes only, no face bytes), a bounded one-time `faces-ack` first-WS-frame message (≤ 128 ids, 2 s window, none-on-no-answer), hash-equality match (version mismatch counts as missing), per-unit raster transition with a `face-missing` diagnostic; the fidelity claim holds per satisfied unit, not document-wide |
| codex-r3 4 (promises and Requests treated as fetch factories) | PARTIAL (R2-5) | **Fixed** — §6.3/§10: promises classified by their *resolved* value (`Promise<Response>` terminal, matching `stream.ts:368–373`); Request renewable only when bodyless or body-not-consumed; expiry tests extend to `Promise<Response>` and body-bearing Requests |
| codex-r3 5 (stale wording: §19 "over"; §21 C3 row; §11 final-raster parenthetical) | PARTIAL (R2-4 clarifications) | **Fixed** — §19 now reads "without"; §21's round-1 claude C3 row points to the superseded/round-3/round-4 lineage; §11 clarifies that final pixels do not identify covered logical ranges and the conservative withhold applies equally on `final-raster` |
