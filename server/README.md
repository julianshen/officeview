# @officeview/stream-server (demo)

The WS-only stream server of the server-streaming design
(`docs/superpowers/specs/2026-11-07-server-streaming-design.md`, Stage 1c).
Private workspace; it imports the library's TypeScript sources directly and
never React.

```bash
bun run stream:demo                 # --demo, serves ./corpus, ephemeral key
# ws://localhost:8787/stream?doc=pypptx-test.pptx            protected (final-raster)
# ws://localhost:8787/stream?doc=pypptx-test.pptx&relaxed=1  audited commands session (--demo only)
bun run test:server                 # bun test (napi preload via bunfig)
```

Outside `--demo` the server refuses to start without `OV_STREAM_KEY` (32 bytes,
base64) or `OV_STREAM_KEYRING`, and an allowed origin (`--origin` /
`OV_STREAM_ORIGINS`). Setting both key variables is an error.

**Key rotation.** `OV_STREAM_KEYRING=<path>` names a JSON file
`{"current": 2, "keys": {"1": "<base64 32 B>", "2": "<base64 32 B>"}}`
(generations are integers 1..255). New sessions are issued under `current`;
each audit row records its `keyGeneration`. To rotate: add a new generation,
switch `current` to it, and keep every older generation until the audit rows
issued under it have expired, so leaked captures from that era stay traceable.
`OV_STREAM_KEY` alone is a keyring with generation 1.

**Audit store.** `OV_AUDIT_DB=<path>` (required outside `--demo`; `--demo`
keeps an in-memory store) names a SQLite file (`bun:sqlite`, WAL,
`synchronous=FULL`, created mode 0600; keep it on local disk and back it up
with its `-wal`/`-shm` files). It holds the write-once audit rows, one
delivery record per `(sid, unit, scale)` served to a forensic session (status
`l1` or `unmarked` with its reason, plus the SHA-256 of the canonical PNG the
mark was applied to), the per-generation token counters, and the trace tool's
detection log. Ordering: a session's row is durable before its manifest is
sent, and a delivery record before its frames; a failed write refuses (HELLO
closes 4009 `policy-refused`; a unit is refused `policy-refused`), never sends
unrecorded. Rows and their deliveries are deleted after 90 days (at startup
and hourly); token counters are never reset.

- **One endpoint.** `GET /stream` WebSocket upgrade from an allowlisted
  `Origin`; every other HTTP request is 404. There is no file route.
- **Your hooks.** `authorize(req)` decides the viewer and per-session policy;
  `resolveDocument(req, decision)` returns the bytes. `src/demo-hooks.ts` is
  the demo pair (bare file names inside one directory).
- **Protected by default.** Hidden content is removed before parsing
  (`src/policy.ts`, fail-closed); units ship as metadata-free PNGs with the
  visible mark composited. Each raster unit also carries the session's own
  invisible Layer-1 pixel mark (`l1-bch512x40@1`; `src/mark/`), added per
  session after the shared cache. The audit row records the token, salt and
  keyed message issued at HELLO, and separately the layers actually delivered
  (`forensicLayers`, counts of marked/unmarked raster units); `untraceable` is
  derived from what was delivered, never from issuance alone. Command units (`projection: 'relaxed'`) carry
  no forensic mark yet (`forensic-commands-unmarked` notice); `forensicMark:
  false` sends `forensic-off`. Audit rows live in memory until Stage 2b.
- **Fonts.** `fontpack/` (20 OFL faces, `pack.json` pins their hashes);
  Calibri/Cambria/Arial/Helvetica/Times New Roman/Courier New are painted
  with their metric-compatible pack faces. Verify: `bun scripts/stream-s1/fontpack-lock.ts`.

## Forensic mark (Stage 2a)

The render cache holds one canonical PNG per unit; a forensic-on session's
raster units then go through a per-session pool `encode` job (decode → embed
the session's 512-sign code over the page-keyed basis → lossless re-encode),
at most 2 in flight per session. Units too small to carry 512 cells ship with
the visible mark only and a `diag forensic-unmarked`. Integration proof on
served bytes (ΔE00, §13.5 survival at the Stage-0 τ, unmarked captures never
attributing): `docs/validation/streaming-s2.md`. Detection tooling, τ
recalibration and audit persistence are Stage 2b; the τ table the trace tool
ships (`server/src/detect/tau.ts`) and the Stage-2 gate (survival, null false
attribution, collusion and capture-geometry batteries through `trace()` on
served bytes) are in `docs/validation/streaming-s2b.md`.

## Investigator tool: `detect` (Stage 2b)

An authorized investigator gives the tool a leaked screenshot (PNG or JPEG),
the document and the page, and gets one plain line per image. Run it from the
repo root with the napi preload (the same backend the server uses):

```bash
OV_STREAM_KEY=<base64 master> bun --preload ./scripts/stream-s0/napi-preload.ts server/scripts/detect.ts \
  --image leak1.png [--image leak2.jpg …] --doc deck.pptx --page 3 \
  --audit-db /var/lib/officeview/audit.db [--keyring keyring.json] [--investigator "J. Doe"] [--json]
```

`--page` is 1-based. Keys come from `--keyring <file>` (same JSON as
`OV_STREAM_KEYRING`), else `OV_STREAM_KEYRING` / `OV_STREAM_KEY`; every
generation in the ring is tried. The audit store is opened in trace mode: rows
and deliveries are never modified, and the store must already exist.

One line per image, one of:

```
<image>: TRACED — session <sid>, viewer <id|unknown>, ip <ip>, started <UTC>, key generation <g>, page <N> (z 22.52, margin 21.49)
<image>: POSSIBLE COLLUSION — candidates <sid> (z …), …; no single viewer named
<image>: NO MATCH — mark below threshold (best z 5.68); nothing attributed
<image>: NO USABLE SIGNAL — <reason>
<image>: REFUSED — <count> candidates exceed the detector limit (4194303)
```

Reasons: no mark signal; geometry not supported (rotated, skewed,
photographed); page never delivered marked; no audit record of the page being
viewed; re-rendered page differs from what was served (renderer or font
drift). A margin that cannot be computed (a single scored candidate) prints as
`n/a`. `--json` prints `{ "image", "answer" }` per line, where `answer` is the
tool's structured answer.

Every image appends a detection log row (time, OS user, `--investigator`,
image SHA-256, the document's per-generation refs, page, answer text, session)
before its answer is shown. Key material is never printed or logged. Exit
codes: 0 every image answered; 2 usage or configuration error; 1 internal
error. The decision thresholds are the Stage-0 provisional values until the
Stage-2b calibration.

## Text services: search, copy, screen-reader text (Stage 3a)

Spec: `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.1;
measurements and test map: `docs/validation/streaming-s3a.md`.

**Only text the page shows is answered.** Each unit's text is recorded with
the visible mark muted. It is then joined against the occlusion evidence of
a paint pass (`server/src/text/join.ts`), and every span the join cannot
prove visible is withheld, never approximated. The served index holds only
the characters that passed. A copy can therefore contain an extra line break
where something in the paragraph was withheld, or at a page break. In a
`visible-only` session hidden content is never answered. In an audited
`contentPolicy: 'full'` session it is on the page and is answered like
other text. Speaker notes are never indexed. XLSX answers the painted
display text, never values or formulas. Text drawn as glyph outlines (warped
WordArt) is withheld.

**Turning services on** — in your `authorize` hook:

```ts
authorize: async (req) => ({
  viewerId: 'alice@example.com',
  policy: {
    copy: { maxParagraphs: 20, maxSessionParagraphs: 200 },  // or 'full', or false (default)
    textA11y: true,                                            // optional, see below
  },
})
```

- `copy`: `false` (default) is off. `'full'` is uncapped. `{ maxParagraphs,
  maxSessionParagraphs? }` caps each copy and, optionally, the session's
  total. A cap that is not a positive safe integer turns copy **off** (fail
  closed).
- `textA11y` (screen-reader text) resolves to one mode:
  - left out: follows `copy` — `'full'` ⇒ `unit`, capped ⇒ `focused`,
    off ⇒ `off`;
  - `true`: `unit` when copy is off or `'full'`, `focused` when copy is
    capped;
  - `false`: always `off`.
  `textA11y: true` with copy off puts the page text in the DOM for
  assistive technology, where any script on the page can read it. That is
  your choice to make.
- **Search is always on** (`can.search: true`). There is no switch. It is
  limited by rate only. Answers are highlight boxes, never text, and are not
  counted or audited per query. Repeated probing can still reconstruct text
  with copy off (parent spec §19).
- Turning copy or a11y on does not mark the session `relaxed`. The audit row
  records the resolved services instead.
- The manifest carries `can.copy`, `can.textA11y` and
  `caps: { copy: null | { maxParagraphs, maxSessionParagraphs }, a11y: 'off' | 'unit' | 'focused' }`
  ('full' ⇒ both caps `null`), so the viewer can explain a refusal.

**What counts as a paragraph.** One Word paragraph is 1, whether in the
body, a table cell, a header or footer, or a footnote. One spreadsheet cell
is 1, wrapped or merged. One paragraph of a text box or shape is 1. A
paragraph that crosses a page break counts once on each page. Empty
paragraphs count 0. Painters announce paragraphs through an optional
recording hook (`RECORD_GROUP`, `src/core/text-recording.ts`) that local
viewing does not define.

**Caps.** A copy over `maxParagraphs`, or one that would take the session
past `maxSessionParagraphs`, is **refused with the cap stated** and copies
nothing. A copy is never truncated. One copy is also at most 4 MiB of UTF-8,
even under `'full'`. Session totals:
- they are held per stream session (sid) in server memory, and they are a
  running total: copying the same paragraph twice counts twice;
- they survive a resume of the same session. If `authorize` returns
  different `copy`/`textA11y` on resume, the new values apply, the amount
  already used is kept and measured against the new caps, and the viewer
  follows the change (its screen-reader layer is removed and its selection
  cleared when the service turns off);
- they **reset when the viewer opens the document again** (a new session).
  A quota per viewer across opens would need your own storage; Stage 3
  has none;
- a copy whose answer is lost to a dropped connection still counts. It is
  never re-issued.

**Screen-reader modes.**
- `unit`: the text of each page, one paragraph per element, fetched when the
  page first comes into view. Audited per request; not counted against the
  caps.
- `focused`: one button per paragraph, "Paragraph k of n. Press Enter to
  read.". Enter fetches that one paragraph. Each paragraph read counts 1
  against `maxSessionParagraphs` (shared with copy) and is audited.
- `off`: no text layer.

When both copy and a11y are off, the view says "text unavailable in this
viewing mode" once.

**Wire** (client → server, each with `reqId` and `token`; server → client
answers are JSON text messages of ≤ 48 KiB of UTF-8 each, chunked):

| request | answer |
|---|---|
| `{ type: 'search', query, caseSensitive? }` | `search-result { matches: [{ unit, rects }] }` …, then `search-done { total, truncated }` |
| `{ type: 'select', range }` | `select-result { units: [{ unit, rects }], paragraphs }` …, `last: true` on the final one |
| `{ type: 'copy', range }` | `copy-result { seq, text, last }` … (concatenate in `seq` order) |
| `{ type: 'a11y', unit, paragraph? }` | `a11y-result { seq, paragraphs: [{ rect, text? }], last }` … |

`range` is `{ all: true }` or `{ anchor: { unit, x, y }, focus: { unit, x, y } }`,
in page px at scale 1. `rects` are flat `[x, y, w, h, …]` per unit.

**Refusals** (`{ reqId, type: 'refused', code, limit?, requested? }`; the
session stays open):

| code | when |
|---|---|
| `text-unavailable` | select or copy with copy off; a11y with a11y off |
| `copy-too-large` | over `maxParagraphs` (`limit` = the cap, `requested` = the paragraphs selected); over 4 MiB (`limit` = 4194304, `requested` = bytes when known); an a11y answer over 4 MiB of text (`limit` = 4194304, `requested` = bytes), over 20,000 paragraphs (`limit` = 20000, `requested` = the paragraph count), or with one paragraph over 48 KiB (`limit` = 49152) |
| `copy-quota-exhausted` | the session total would pass `maxSessionParagraphs` (`limit` = the cap, `requested` = this copy's paragraphs; 1 for a focused read) |
| `rate-limited` | over the `'text'` rate, or the pool is saturated |
| `policy-refused` | malformed arguments, or an internal failure |
| `session-rotated` | a request under a superseded token. The client re-issues a unit request, a search or a select once under the new token. A copy or an a11y read rejects and is not replayed (it may already have counted). |

**Rates and limits.**
- Rate class `'text'`: 60 requests per minute per session, covering search,
  copy and a11y (refused off-mode a11y requests included). An overflow is
  refused `rate-limited` and does not count toward the breach window that
  closes the socket (4008). Select uses only the general message budget
  (120 per 10 s), and the viewer sends at most one select every 200 ms.
- A query is at most 256 UTF-16 units. A search answers at most 5,000
  matches (`truncated: true` beyond). A single match too large for one
  message is left out and also marks the answer truncated. One search runs
  per session; a new one supersedes it. The worker stops matching once it
  has the matches still wanted (in unit, line, position order) and stops
  building highlight rects once 100,000 are reached in a job of 8 units
  (it may overshoot by the rects of the one match that passes the bound,
  which is dropped); such a job ends the search with what it has,
  `truncated: true`.
- The viewer rejects any answer of more than 4,096 messages, a select of more
  than 20,000 rects, and an a11y answer over 4 MiB of text or 20,000
  paragraphs. A single a11y paragraph over 48 KiB refuses the request.
- A page that records more than 3 MiB of text pieces (about 12,000) is
  refused for text services. The served index is cached up to 32 MiB per
  document per render worker.

**Audit.** The audit store gains two tables. Neither holds any text.
- `copies (sid, kind, paragraphs, from_unit, to_unit, at)`: one row per
  answered copy (`kind 'copy'`) and one per a11y request that returned text
  (`kind 'a11y'`: a `unit` page, or a `focused` paragraph). Refused requests,
  empty selections and rects-only a11y answers write none.
- `services (sid, copy, text_a11y, at)`: one row per resume that changed the
  resolved services. The open row is write-once and keeps the value at open.

`AuditTable.get(sid)` returns `services` (at open), `servicesChanged`, and
`copied: { events, paragraphs }` (copies only, not a11y). A copy's record is
durable before its text is sent; a failed write refuses it.

**Viewer messages** (exact): 'Selection too large to copy (limit: N
paragraphs).', 'Copy limit for this session reached (N paragraphs).',
'Selection too large to copy.' (the byte cap), 'Too many requests. Try again
shortly.', 'Copy failed.', and 'Text unavailable.' (a11y). The search bar
shows 'too many requests', 'search unavailable', 'search failed' or 'query
too long'.

**Yield.** `bun --preload ./scripts/stream-s0/napi-preload.ts server/scripts/text-yield.ts`
prepares every corpus unit and prints, per format, the spans served and the
withheld spans by cause (`docs/validation/streaming-s3a.md`).

## Printing (Stage 3b)

Spec: `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.2 and
its 2026-11-13 amendment; measurements and test map:
`docs/validation/streaming-s3b.md`.

**What prints.** The server obtains a print-resolution raster for each page
(scale 2, or scale 1 when a page or the whole job is too large at 2). It may
come from the shared canonical render cache, the same deterministic bytes a
screen render at that scale would have. The visible mark is part of that
raster, and the session's forensic mark is applied per page and per request,
exactly as on screen; a page too small to carry the forensic mark
(`capacity` / `all-masked`) prints with the visible mark only and is audited
as an `unmarked` delivery. The viewer prints those print
rasters through a hidden iframe. It never prints the canvases it shows: the
remote view is hidden in print media, so the browser's own Print menu
prints a blank page. A page is one streamed unit: a DOCX page, a PPTX slide
(no notes pages or handouts), an XLSX sheet unit as it streams. In a
`visible-only` session hidden slides are not units and never print; in an
audited `contentPolicy: 'full'` session they print as they show.

**Turning it on or capping it** — in your `authorize` hook:

```ts
authorize: async (req) => ({
  viewerId: 'alice@example.com',
  policy: {
    print: { maxPages: 20, maxSessionPages: 100 },  // or true (default, uncapped), or false
  },
})
```

- `print`: `true` (the default) prints the whole document with no dialog.
  `false` turns print off: no Print button, and Ctrl/Cmd+P inside the view
  does nothing.
  `{ maxPages, maxSessionPages? }` caps each job and, optionally, the
  session's total. A cap that is not a positive safe integer, or any other
  value, turns print **off** (fail closed).
- The manifest carries `can.print` and
  `caps.print: null | { maxPages, maxSessionPages }` (uncapped ⇒ both
  `null`; print off ⇒ `null`).
- With a cap, Print opens a "Print pages" dialog with From and To page
  numbers (defaults 1 and the smaller of the page count and `maxPages`) and
  the lines "Up to N pages per print." / "Up to N pages in this session.".

**Caps.** A job over `maxPages`, or one that would take the session past
`maxSessionPages`, is **refused with the cap stated** and prints nothing. A
job is never truncated. Session totals:
- are held per stream session (sid) in server memory, as a running total:
  printing a page twice counts twice;
- survive a resume. If `authorize` returns different print caps on resume,
  the new caps apply and the pages already printed are kept and measured
  against them;
- **reset when the viewer opens the document again** (a new session);
- are charged when the job is accepted, before its first page. A job that
  is cancelled or fails part-way (render error, lost connection) stays
  charged, and nothing prints on the viewer.

**Wire** (client → server, with `reqId` and `token`):

| message | meaning |
|---|---|
| `{ type: 'print', from, to }` | print units `from..to` (0-based, inclusive) |
| `{ type: 'print-cancel' }` | stop the running job with this `reqId`; anything else is ignored, never answered |

Server → client on the job's `reqId`: `print-start { pages, scale }`, then
exactly `pages` PRINT_RASTER (kind 8) messages in ascending unit order (body
`u32 LE unitIndex` + the UNIT_RASTER body: `u32 w, u32 h, u32 pngLen, PNG`),
then `print-done { pages }`. A page's raster size `w` satisfies
`(unit.w − 1) × scale < w ≤ unit.w × scale` (the same for `h`), where `unit`
is the manifest size: the renderer rounds after scaling and the manifest
before. A failure at any point is one `refused` and nothing more.

**Refusals** (`{ reqId, type: 'refused', code, limit?, requested? }`; the
session stays open):

| code | when |
|---|---|
| `print-too-large` | over `maxPages` (`limit` = the cap, `requested` = the job's pages); or the pages fit neither scale 2 nor 1 (each page ≤ 8192 px per axis and ≤ 64 MiB of RGBA, the job ≤ 1 GiB of RGBA in all): no `limit` |
| `print-quota-exhausted` | the session total would pass `maxSessionPages` (`limit` = the cap, `requested` = the job's pages) |
| `rate-limited` | over the `'print'` rate, a job already running on the session (or a cancelled job's last pages still settling), or the render pool is saturated |
| `policy-refused` | print off, a malformed range or token, a page outside the size contract, a dropped send, or an internal failure |
| `session-rotated` | a `print` under a superseded token: the viewer re-issues it once under the new token (nothing was charged). A job whose connection drops mid-way is not re-issued (it stays charged). |

**Rates and limits.**
- Rate class `'print'`: 10 jobs per minute per session, counting every
  `print` request that reaches the rate check (jobs then refused by a cap
  included). An overflow is refused `rate-limited` and never counts toward
  the breach window that closes the socket (4008).
- One job at a time per session; at most 2 pages rendering or waiting to be
  sent. A page is not rendered or sent while the socket holds more than
  4 MiB unsent (`PRINT_SEND_HIGH_WATER`); sockets keep Bun's default 16 MiB
  `backpressureLimit`.
- PRINT_RASTER: 8 MiB encoded, 64 MiB decoded. The viewer reserves
  `8 MiB + 2 × pageCap` from its 96 MiB per-session admission for a job,
  with `pageCap` at most 44 MiB, so **a page whose PNG is over about
  44 MiB cannot print** from a stream (the job fails). It holds at most
  192 MiB of PNG per job, and a job whose images would decode to more than
  1 GiB of RGBA at the announced scale (Σ `w × h × 4`, the server
  planner's own budget, checked at `print-start` from the manifest sizes
  and again per page) is failed before any page is kept (either way:
  "This document is too large to print from the viewer.").
- The viewer checks every page against the job (index, order, count, scale,
  size) and prints nothing when anything does not match.

**Audit.**
- `prints (sid, pages, from_unit, to_unit, scale, at)`: one row per
  accepted job, durable before `print-start` (a failed write refuses the
  job). Refused jobs write none.
- With the forensic mark on, each page also gets a `deliveries` row like a
  screen unit (the same unit and scale already delivered on screen is not
  repeated).
- The open row's services (`AuditRow.services`) include the resolved
  `print`, and `services` (one row per resume that changes the services)
  gains a `print` column. A store created before Stage 3b is migrated
  in place (`ALTER TABLE`). Its older rows have no print value and read
  `print: false`, which means "not recorded", not "print was off".
- `AuditTable.get(sid).printed` is `{ jobs, pages }`.

**Viewer messages** (exact): 'Preparing pages D of T…' (with Cancel), 'Too
many pages to print at once (limit: N pages).', 'This document is too large
to print from the viewer.', 'Print limit for this session reached (N
pages).', 'Too many requests. Try again shortly.', and 'The document could
not be printed. Please try again.'. A cancel shows nothing.

**Honest limits.** The `unit` op still serves screen rasters, so print caps
limit printing, not capture: what the viewer sees can be screenshotted, and
the marks trace it. See parent spec §19 item 11.

## The browser side (Stage 1d)

```tsx
<OfficeFile
  data={{ kind: 'stream', url: 'wss://viewer.example.com/stream?doc=…', token }}
  streamFonts={{ baseUrl: '/officeview-fontpack' }}   // optional: serve fontpack/ here
  onStreamNotice={(n) => console.log(n.code)}         // manifest notices, face-missing, resumed, unit-failed, session end codes
  streamClient={(s) => new WebSocket(s.url)}          // optional factory (the default)
/>
```

- **The mode is the branch.** A `StreamSource` (or a promise of one) never reaches `fetch` or a byte reader; the client code is a lazy chunk, so pages that only load files never download it, and local loads paint byte-for-byte as before (`docs/validation/streaming-s1.md`).
- **What a reader gets.** Rasters (protected) or replayed commands (audited relaxed sessions), painted at the scale bucket for the device pixel ratio. Since Stage 3a: a search bar (Ctrl/Cmd+F), drag-selection, select-all and copy (Ctrl/Cmd+A, Ctrl/Cmd+C) when the server grants `can.copy`, and a screen-reader text layer when it grants `can.textA11y` (see "Text services" above). All of it is answered by the server; the client builds no text index of its own. When both copy and a11y are off, the view says "text unavailable in this viewing mode" once. No zoom or presentation yet. The host `watermark` prop is ignored (the server bakes the mark); print follows the server's `can.print`.
- **Fonts.** The font pack is **not** in the npm package. Protected sessions need no fonts (every unit is a raster). For relaxed sessions, serve `fontpack/` (`pack.json` + TTFs) and pass `streamFonts.baseUrl`; each face is hash-checked against the manifest before it is registered, and any face that is missing, unverifiable or not hashed within 1.5 s is reported missing, so its units arrive as rasters.
- **Drops.** An abnormal drop (1001/1006) resumes the same session for up to 30 s (`detachGraceMs`, capped by the token's `exp`); the scroll position is untouched and pending units are re-requested. A terminal close (revoked, expired, oversized frame, …) shows "Viewing session ended (<code>)" and calls `onStreamNotice` with that code; recovery is a host re-open.
- **Test harness.** `bun run test:browser:stream` runs `tests/browser/stream-viewer.spec.ts` against `server/test/app.ts`, which adds test-only controls (drop / revoke / oversize) on a separate port.
