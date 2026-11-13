# Server streaming — Stage 3 design: search, copy, print, presenting

Date: 2026-11-11. Status: draft for owner review (builder check done; its three owner questions answered 2026-11-11).
Parent spec: `docs/superpowers/specs/2026-11-07-server-streaming-design.md`
(cited below as "parent §N"). The parent spec is normative wherever this
document is silent; this document records the owner's Stage 3 decisions
and the order the work ships in.

## 1. Intent

A viewer whom the developer's backend has already let in (authorization is
the developer's concern, not this library's) opens a protected document
over the stream. They mostly **read it and present it**. Stage 3 makes a
streamed document as usable as a local one for those two jobs, without the
file or any hidden content ever reaching the client:

- they can **search** it;
- they can **copy** and **print** from it **when, and as much as, the
  developer allows**;
- they can **present** a streamed deck exactly as they present a local
  one — transitions, builds, animations, Morph, media, and Presenter View
  with their notes.

Today (main 3f38b375) the server answers every manifest with
`can: { search: false, copy: false, print: <policy>, textA11y: false,
present: false, notes: false, links: false }`
(`server/src/ws.ts:280`); none of these services is implemented on the
stream path.

## 2. Owner decisions (2026-11-11)

1. **Search is allowed** for every stream session.
2. **Copy and print are controlled by the developer**, per session, in the
   server-side `authorize` hook (parent §6.1). The client has no surface
   to change them.
3. **Copy has a developer-set scope**: off, full, or a paragraph cap
   (owner example: "full, max n paragraph…").
4. **Print has a developer-set page cap**, in the same shape as copy.
5. **Presenting must be the same as local presenting**: transitions,
   click-by-click builds, animations and Morph — the parent §12a target.
6. **Embedded video and audio play** too (parent §12a.6's optional Stage
   3d becomes required).
7. **Presenting is a one-machine setup only**: one person's computer
   drives the audience screen and Presenter View, as locally today. No
   remote audience following along in their own browsers.
8. Order and packaging: five pieces, **each its own PR, shipped in the
   order of §4** (owner agreed to the proposed order).

## 3. Goals, non-goals, anti-goals

**Goals:** the five pieces of §4 meet their "done when" lines; every
service obeys the parent §11 join rule (only visible text is ever
answered) and parent §12a.8.

**Non-goals:** remote/broadcast audiences (decision 7); viewer
authorization; editing; XLSX tiles and server sort/filter (parent Stage
4); `culled-commands` (closed, parent §17 note).

**Anti-goals — a failure even if it "works":**
- any text the page does not show (hidden content, the visible mark's own
  text, notes for a non-presenter) appearing in a search, copy, a11y,
  print or show answer;
- a copy or print that exceeds the developer's cap by being **silently
  truncated** instead of refused — the viewer must know;
- the client building a local text index, using local selection
  machinery on a remote document, or ever falling back to local rendering
  (parent §11, §12a.7);
- original media bytes, PPTX parts or shape data on the wire.

## 4. The five pieces, in order

### 4.1 Piece 1 — search, copy, semantic text (parent §11)

**Policy surface** (extends `server/src/hooks.ts` `authorize` policy):

```ts
copy?: false | 'full' | { maxParagraphs: number; maxSessionParagraphs?: number }
// default false (my call: protected sessions; the owner said only that
// the developer controls copy)
```

- `can.search` is **true for every session** (decision 1). There is no
  developer switch. *My call*: search always on, as stated by the owner.
- `can.copy` is true iff `copy !== false`. The manifest carries the
  resolved caps so the client can explain a refusal without guessing.
- **`can.textA11y` follows the copy setting (owner, 2026-11-11 — this
  amends parent §2 item 12's "defaults on" for stream sessions).** A new
  `textA11y?: boolean` policy field defaults to `copy !== false`; the
  developer can turn it on separately. With `copy: 'full'` (or an
  explicit `textA11y: true` and no cap) a11y text is served per unit.
  With a paragraph cap, a11y text is sent **only for the focused
  element** and each paragraph sent counts toward the same
  `maxParagraphs` / `maxSessionParagraphs` totals. With a11y off, the
  parent §15 notice ships. It is served from the same index and join
  rule, so it ships in this piece.

**What counts as a paragraph** (the copy unit; *my call* on the owner's
"paragraph"): a text paragraph in a DOCX body, header/footer, text box or
table cell; a text paragraph in a slide shape or table cell; one cell in a
spreadsheet. Empty paragraphs count as zero.

**Behaviour:**
- Search: the client sends the query; the server answers
  `HighlightRect[]` (`src/core/overlay.ts:7`) per unit from the index
  built in the recorder's second pass with the visible mark muted
  (`src/core/search.ts`). Results only include text that passes the join
  rule.
- Copy: the viewer drag-selects on the remote view; the client sends the
  selection geometry (unit + rectangle/caret range in unit coordinates);
  the server resolves it against the index, applies the join rule, counts
  paragraphs, and answers display text in exact locale/BiDi order. XLSX
  answers formatted display text, never values or formulas. Select-all is
  a selection like any other.
- **Over a cap:** if the selection exceeds `maxParagraphs`, or the
  session's copied total would exceed `maxSessionParagraphs`, the server
  answers a typed refusal (`copy-too-large` / `copy-quota-exhausted`)
  carrying the cap; the client shows it and copies **nothing**. Never a
  truncated copy.
- Copy off: "text unavailable in this viewing mode" shown once (parent
  §11).
- **What "session" means for the caps (owner, 2026-11-11):** the totals
  are per stream session (sid), held in server memory. They survive a
  resume of the same sid (if `authorize` returns different caps on
  resume, the amount already used is kept and measured against the new
  caps), and **reset when the viewer opens the document again**. This is
  stated in the developer docs and in the parent's Honest limits. The
  count is a running total: copying the same paragraph twice counts
  twice. A per-viewer quota across opens would need a developer storage
  hook — not in Stage 3. The same rules apply to `maxSessionPages`.
- **Search is limited by rate only (owner, 2026-11-11).** Answers are
  highlight boxes, never text, but repeated probing can reconstruct text
  even with copy off. That is accepted and added to the parent's Honest
  limits; search does not count against the copy caps.
- Every copy is audited with its paragraph count; search queries are
  rate-limited, not audited per query (*my call*; parent §14 rates
  apply, a new rate class for search/copy is the builder's to size).
- **Notes are not indexed in piece 1.** Notes text never appears in
  search or copy answers for anyone; a presenter sees notes only as
  server-drawn marked images (parent §12a.5/§12a.8). *My call*: the
  parent's "no notes text outside the a11y service" leaves room for a
  presenter-only notes a11y answer; that is piece 3's to add, re-checking
  `can.notes` after resume.
- **Which text is visible follows the session.** In an audited
  `contentPolicy: 'full'` session the hidden content is drawn on the page,
  so search, copy and print answer it like any other visible text; in
  every `visible-only` session it is never answered. *My call*: the
  anti-goal is about text the viewer's page does not show.

**Done when:** parent §11 end-to-end assertions green (visible-mark text
never in answers; no local index; hidden constructs never matched or
copied — the 13 Stage-0 hidden constructs from parent T2); caps refuse at
n+1 and allow at n, per copy and per session; XLSX copies display text;
audit rows carry the counts.

> **Amendment (2026-11-12, Stage 3a as built).** Decisions taken while
> building piece 1 (plan `docs/superpowers/plans/2026-11-11-streaming-s3a-text-services.md`,
> "Execution notes"). Validation: `docs/validation/streaming-s3a.md`.
>
> 1. **Paragraph counting (supersedes the plan's R2).** One Word paragraph
>    counts 1, whether in the body, a table cell, a header or footer, or a
>    footnote. One spreadsheet cell counts 1, wrapped or merged. One
>    paragraph of a text box or shape counts 1. A paragraph that crosses a
>    page break counts once on each page. Empty paragraphs count 0. Painters
>    announce the paragraph or cell they draw through an optional recording
>    hook, `RECORD_GROUP` (`src/core/text-recording.ts`), which local
>    viewing does not define. The plan's per-painted-line count for wrapped
>    cells and per-line count for DOCX body text are gone.
> 2. **Resume (R3).** `copy` and `textA11y` may change on resume; they are
>    excluded from the policy-equality check, and every other field keeps
>    the exact-match rule (parent §6.2 rule 7). The new values replace the
>    old ones, the totals already used are kept and measured against the new
>    caps, and the resumed manifest carries the new caps. The viewer follows
>    the change: a11y turned off removes its text layer, and copy turned off
>    clears the selection and its handlers. The change is audited (item 8).
> 3. **The served index holds only characters proven visible.** After the
>    join, each paragraph's stored text is cut down to its served ranges
>    (the sanitise step, `server/src/text/served.ts`). Pieces are merged only
>    across whitespace. A copy can therefore contain an extra line break
>    where something in the paragraph was withheld, or at a page break.
> 4. **The join as built** (`server/src/text/join.ts`). A run is a
>    candidate for a span by geometry: its ink centre lies in the span box
>    inflated by 2 px, or the span box centre lies in its ink. A span is
>    served iff it has ≥ 1 candidate, every candidate is servable, and the
>    span's whitespace-stripped text occurs contiguously, at wholly
>    unconsumed positions, in the candidates' text in paint order. Each
>    painted character vouches for one span only. A whitespace-only span is
>    served iff at least one nearest same-line neighbour with ink exists
>    (before or after it) and every such neighbour is served.
> 5. **Limits.**
>    - A query is ≤ 256 UTF-16 units.
>    - A search sends ≤ 5,000 matches. A single match that cannot fit one
>      message is omitted, and the answer says `truncated: true`. A worker
>      search job (8 units) stops matching at the matches still wanted and
>      stops building highlight rects once 100,000 are reached (it may
>      overshoot by the rects of the one match that passes the bound,
>      which is dropped); the search then ends with what it has,
>      `truncated: true` (2026-11-12).
>    - An a11y answer over 20,000 paragraphs is refused `copy-too-large`
>      with `limit` 20000; over 4 MiB of text, with `limit` 4194304.
>    - One copy is ≤ 4 MiB of UTF-8 (`copy-too-large` beyond, even under
>      `'full'`).
>    - Every text answer message is ≤ 48 KiB of UTF-8.
>    - The client accepts ≤ 4,096 messages per request, a select of
>      ≤ 20,000 rects, and an a11y answer of ≤ 4 MiB and ≤ 20,000
>      paragraphs. A single a11y paragraph over 48 KiB refuses the request.
>    - A unit recording more than 3 MiB of text spans (about 12k spans) is
>      refused for text services.
>    - The served-index cache is 32 MiB per document per worker.
> 6. **Rates.** A new rate class `'text'`: 60 requests per minute per sid,
>    covering search, copy and a11y (refused off-mode a11y requests
>    included). An overflow is refused `rate-limited` and does not count
>    toward the breach window that closes the socket. Select uses only the
>    message budget, and the client sends at most 5 selects per second (one
>    per 200 ms). This settles "a new rate class … the builder's to size"
>    above.
> 7. **A11y.** The modes are `off`, `unit` and `focused`:
>    - `unit` (per-unit text, uncounted, audited): `copy: 'full'`, or
>      `textA11y: true` with copy off;
>    - `focused`: copy is capped;
>    - `off`: otherwise, and always for `textA11y: false`.
>
>    In `focused` mode each paragraph is a native button whose text is
>    "Paragraph k of n. Press Enter to read." One paragraph is fetched per
>    request; it counts 1 against `maxSessionParagraphs` (shared with copy)
>    and is audited. `textA11y: true` with copy off puts page text in the
>    DOM for assistive technology. That is the developer's choice.
> 8. **Audit.** New `copies` table: kind `copy` or `a11y`, paragraph count,
>    unit range and time, and no text. New `services` table: one row per
>    resume that changes the services. The open row records the services at
>    open (`AuditRow.services`). `AuditRow.copied` holds the copy events and
>    their paragraph total.
> 9. **Wire.** Requests `search {query, caseSensitive?}`, `select {range}`,
>    `copy {range}` and `a11y {unit, paragraph?}` carry `reqId` and `token`.
>    Answers are `search-result`/`search-done`, `select-result` (`last` on
>    the final chunk), `copy-result {seq, text, last}` and
>    `a11y-result {seq, paragraphs:[{rect, text?}], last}`. New refusal
>    codes are `text-unavailable`, `copy-too-large` and
>    `copy-quota-exhausted`; the last two carry `limit`. `requested` is
>    present on `copy-quota-exhausted`, and on `copy-too-large` only when
>    it is known (the paragraphs selected, the bytes of a built answer, an
>    a11y answer's paragraph count; not when the worker refused before
>    building the text) (parent §6.2 rule 8 amendment).
> 10. **UI strings (exact).** 'Selection too large to copy (limit: N
>     paragraphs).', 'Copy limit for this session reached (N paragraphs).',
>     'Selection too large to copy.', 'Too many requests. Try again
>     shortly.', 'Copy failed.', 'Text unavailable.'; search failures 'too
>     many requests', 'search unavailable', 'search failed', 'query too
>     long'. "text unavailable in this viewing mode" is shown once when
>     both copy and a11y are off.
> 11. **A copy whose answer is lost** to a dropped connection still counts
>     toward the session cap. It is never re-issued.
> 12. **Text decorations.** Underlined or struck text is served under a
>     narrow exemption to the parent's 100 %-ink rule (parent §7.7
>     amendment). Its residuals are listed in parent §19. As of 2026-11-12
>     a line of a neighbouring segment counts as the run's own line carried
>     on only when it is exactly collinear with it (both edges within
>     0.01 device px) and lies beside it, not on it (the two lines overlap
>     along the baseline by no more than their stroke-cap allowances). A
>     stroke has such an allowance only with butt caps and only when it is
>     drawn under an unsheared transform (an axis-aligned scale/translate
>     or a rotation by a multiple of 90°); otherwise its whole box counts
>     as painted. The review's residual A is closed.

### 4.2 Piece 2 — print (parent §2 item 9)

**Policy surface:** `print?: boolean | { maxPages: number; maxSessionPages?: number }`
(today `print?: boolean`, default true — unchanged default).

**Behaviour:**
- The pool re-renders the requested pages at print resolution with the
  session's marks (both marks, as on screen) and ships `PRINT_RASTER`
  frames; the client prints them from a print-only surface. Print units
  never reuse screen rasters.
- **With a cap:** the viewer picks a page range before printing (*my
  call* on how the cap is surfaced); a job over `maxPages`, or over the
  remaining `maxSessionPages`, is refused with the cap stated
  (`print-too-large` / `print-quota-exhausted`). Never truncated.
- **Without a cap:** the whole document prints, as locally.
- **What a page is** (*my call*): one streamed unit — a DOCX page, a
  PPTX slide (no notes pages or handouts), and for XLSX the sheet units as
  they stream today (parent Stage 4 tiles may change XLSX pagination
  later). Hidden slides are not pages.
- Own rate class (parent §14: 10 jobs/min, 2 concurrent renders); every
  job audited with page count and range.

**Done when:** print audit rows present; caps refuse/allow at the
boundary; printed rasters carry both marks (the forensic mark detects on
a print-resolution page via the Stage 2b detector); no hidden content in
print rasters (parent T2 constructs).

> **Amendment (2026-11-13, Stage 3b as built).** Rulings taken while
> planning (P1–P7, plan `docs/superpowers/plans/2026-11-13-streaming-s3b-print.md`)
> and while building. Validation: `docs/validation/streaming-s3b.md`.
>
> - **P1 — print resolution is scale 2, falling back to 1.** The fallback
>   is per job: it applies when any page's scale-2 surface fails
>   `unitSurfaceAllowed` (≤ 8192 px per axis and ≤ 64 MiB of RGBA), or the
>   job's Σ w·h·4 passes `PRINT_JOB_MAX_DECODED` (1 GiB). A job that fits
>   at neither scale is refused `print-too-large` with no `limit`, before
>   anything is charged. This matches the local viewer's `PRINT_SCALE = 2`.
>   Higher scales would need render and tracer changes (the Stage 2b
>   tracer keys chips on `unit@scale` within `SCALE_BUCKETS`).
> - **P2 — "print units never reuse screen rasters" means the client never
>   prints what it shows.** The server may serve print pages from the
>   canonical render cache, the bytes the tracer reproduces. Layer-1 is
>   applied per print page, per request, as on screen. A `deliveries` row
>   deduplicates against a screen delivery of the same unit and scale.
> - **P3 — charging.** A job is validated whole before any frame. The order
>   as built (`server/src/ws.ts` `onPrint`) is: token, range arguments,
>   the `'print'` rate, print off, a job already running, `planPrint`, the
>   per-job cap, the session quota, the layout fingerprint, the parse
>   ledger (re-charged on every page of a long job), the audit row
>   (`recordPrint`; a failed write refuses `policy-refused`), the charge,
>   `print-start`, then pages. Every request that reaches the rate check
>   takes a rate slot, refused ones included. A job that fails mid-way
>   (render error, dropped send, socket loss, viewer cancel) prints
>   nothing on the client, and its pages stay charged. Jobs are serialised
>   per sid: a `print` while a job runs, or while a cancelled job's
>   in-flight pages are still settling, is refused `rate-limited`.
> - **P4 — print caps may change on resume**, as copy caps do. `print`
>   leaves the resume policy match, the change is recorded through
>   `recordServices`, and the pages already printed are kept and measured
>   against the new caps.
> - **P5 — a page is a streamed unit, under every content policy.** Under
>   `visible-only`, hidden slides are dropped before parsing and are never
>   units or pages. Under audited `contentPolicy: 'full'` they print,
>   exactly as they show. This narrows "hidden slides are not pages" above
>   to `visible-only` sessions.
> - **P6 — a page Layer-1 cannot mark (`unmarked: 'capacity' |
>   'all-masked'`) prints as it shows**, with the visible mark only and an
>   `unmarked` delivery row, the same as on screen.
> - **P7 — print off is refused `policy-refused`.** The client never sends
>   `print` when `can.print` is false.
> - **Page size contract (built).** The renderer sizes a raster
>   `sizeCeil(raw × scale)` while the manifest rounds first
>   (`sizeCeil(raw)`), so for a fractional raw page size a scale-2 raster
>   can be 1 px under `ceil(unit.w × 2)`. The contract is
>   `(unit.w − 1) × scale < w ≤ unit.w × scale`, and the same for `h`,
>   where `unit` is the manifest size. The server fails the job
>   `policy-refused` for a page outside it (before mark, record or send);
>   the client fails the job for one too. The render size is unchanged, so
>   the tracer still reproduces canonical bytes.
> - **Wire (built).** Client → server `{reqId, type: 'print', token, from,
>   to}` (0-based inclusive, `0 ≤ from ≤ to < units.count`) and `{reqId,
>   type: 'print-cancel'}` (unknown reqIds ignored, never answered).
>   Server → client `print-start {pages, scale}`, then exactly `pages`
>   PRINT_RASTER (kind 8) messages in ascending unit order, each `u32 LE
>   unitIndex` + the UNIT_RASTER body, then `print-done {pages}`; or a
>   single `refused {code, limit?, requested?}` and nothing more for that
>   reqId. `print-too-large` (over `maxPages`) and `print-quota-exhausted`
>   carry `limit` and `requested`; `print-too-large` from `planPrint` has
>   neither.
> - **Server pacing (built).** At most `PRINT_CONCURRENCY` (2) pages are in
>   flight or buffered per job. A page is neither rendered nor sent while
>   the socket holds more than `PRINT_SEND_HIGH_WATER` (4 MiB) unsent. The
>   sockets keep Bun's default 16 MiB `backpressureLimit`; a load-time
>   assertion keeps `PRINT_SEND_HIGH_WATER` + one framed page (8 MiB +
>   fragment headers) under it, leaving about 4 MiB for interleaved unit
>   and text output. A send the transport drops fails the job
>   `policy-refused`.
> - **Client admission (built).** A job claims the client's one print slot
>   at once (a second `print()` rejects `rate-limited` without sending),
>   then takes a grant from the session's 96 MiB aggregate admission, the
>   same one units use, and sends `print` only once the grant is held. The
>   grant is `8 MiB + 2 × pageCap`, the inflate peak (encoded fragments,
>   inflated chunks and their contiguous copy). `pageCap` is the range's
>   largest unit cost at scale 2 plus the 16-byte page header, clamped to
>   64 MiB and to
>   `PRINT_PAGE_BODY_MAX = (96 MiB − 8 MiB) / 2 = 44 MiB`. A page whose body
>   is larger fails the job (honest limit, parent §19). The client holds
>   PNG blobs only, never RGBA, up to `PRINT_JOB_MAX_BLOB_BYTES` (192 MiB;
>   beyond, `print-too-large` with no `limit`, a `print-cancel`, and the
>   blobs released). The client also holds the job's decoded images to
>   `PRINT_JOB_MAX_DECODED` (1 GiB, the server planner's twin): at
>   `print-start` it sums `(unit.w × scale) × (unit.h × scale) × 4` over
>   the range from the manifest sizes and the announced scale, and beyond
>   the budget fails the job `print-too-large` with no `limit` (a
>   `print-cancel`, everything released) before any page is kept; a
>   running Σ `w × h × 4` of the checked page sizes repeats the check.
>   Small, highly compressible PNGs therefore cannot hand the print
>   surface more decoded images than the planner allows. Every page is checked against the job: index, order,
>   count, scale, size and `print-done.pages`, and `print-start` /
>   `print-done` are held to 48 KiB like text answers. Anything else fails
>   the job `policy-refused` and nothing prints; the session survives. On
>   every settlement the job drops its partial reassembly
>   (`Reassembler.drop`), releases its grant and abort listener, and
>   settles once even if a send throws; a cancel sends `print-cancel` best
>   effort. A `session-rotated` refusal before `print-start` (a RENEW race,
>   nothing charged) re-issues the job once; a socket drop mid-job fails
>   it (`session-rotated`, shown as "could not be printed") and it is
>   never re-issued, since the server kept it charged.
> - **The remote view never prints its canvases.** Its root carries
>   `data-ov-print="false"` at all times and an inline
>   `<style data-ov-print-guard>` that hides this instance's root in print
>   media (it does not depend on `officeview.css`). The `beforeprint`
>   paint path is gone, and Ctrl/Cmd+P pressed inside the view is always
>   swallowed (with focus outside the view the browser's own print runs
>   and prints the view blank). Print rasters
>   go through the local viewer's hidden-iframe surface (`printRasters`,
>   shared with `printDocument`, loaded on demand), whose frame is hosted
>   on `document.body`, outside the hidden root. A manifest with
>   `can.print: true` and `caps.print: null` is contradictory and fails
>   closed (no Print button). `printRasters` also refuses page sizes that
>   are not finite, positive and ≤ `PRINT_MAX_AXIS` (16384) with "The
>   document could not be printed. Please try again." (Task 6 ruling).
> - **How the cap is surfaced (the "my call" above).** Uncapped, Print
>   prints the whole document with no dialog. With any cap, Print opens a
>   "Print pages" dialog with 1-based "From"/"To" inputs (defaults 1 and
>   `min(count, maxPages ?? count)`) and the lines "Up to N pages per
>   print." / "Up to N pages in this session.". Progress reads "Preparing
>   pages D of T…" with a Cancel button. Notices (exact): "Too many pages
>   to print at once (limit: N pages).", "This document is too large to
>   print from the viewer.", "Print limit for this session reached (N
>   pages).", "Too many requests. Try again shortly.", "The document could
>   not be printed. Please try again."; an abort shows none. The dialog
>   contains keyboard focus (Tab/Shift+Tab wrap; Escape closes and returns
>   focus to Print) but does not make the page behind it inert: a click or
>   programmatic focus can still reach background controls (deferred, as in
>   the pivot group dialog it follows). A replaced
>   document aborts the old job and starts with no progress, range or
>   notice. A session end (`doc.ended`, terminal or a host close) turns
>   print off for that document: it aborts the network job and a surface
>   still importing or loading its images (the frame never calls
>   `print()`), closes the dialog, clears progress and notice (the
>   viewer's "Viewing session ended" banner is the only message) and
>   hides Print; an old document's late end changes nothing.
> - **Evidence (Task 4 ruling).** The no-hidden-content check compares each
>   printed page of the 13 hidden-construct fixtures with the print of a
>   hidden-free copy by decoded pixels, with a mutation self-check that a
>   planted marker is caught, instead of comparing a print to its own
>   canonical render (which a contaminated renderer would pass).
> - **Audit.** New `prints (sid, pages, from_unit, to_unit, scale, at)`
>   table, one row per accepted job, written before `print-start`.
>   `services` gains a `print` column (a store from before Stage 3b is
>   migrated with `ALTER TABLE`; its old rows read `print: false`, meaning
>   "not recorded", not "print was off"). `AuditTable.get(sid).printed` is
>   `{jobs, pages}`.

### 4.3 Piece 3 — presenting, slide by slide, with notes (parent 3b)

The parent §12a.8 3b list, verbatim in intent: present a streamed deck;
forward, back, `goTo()`; hidden slides never appear and the counter
excludes them; Presenter View shows notes **only** for a `role:
'presenter'` viewer with `can.notes`, and never on the audience window; a
non-presenter gets no notes from any request; both marks on every slide
and notes image; revocation mid-show exits cleanly and never falls back
to local rendering. Gesture, prefetch and resolution rules: parent
§12a.7. One machine only (decision 7): the audience window and Presenter
View are both surfaces of the same client session.

**Where `can.present` and `can.notes` come from** (*my call*):
`can.present` is true for every PPTX session (presenting is the normal
use; there is no developer switch, as with search). `can.notes` is true
iff the developer's `authorize` returns `role: 'presenter'` and does not
set a new `notes?: false`; the default for a presenter is true.

**Done when:** parent §12a.8 3b list green.

### 4.4 Piece 4 — animations, transitions, Morph (parent 3c)

Show plans (`OVSP`, parent §12a.2) and server-drawn layer bitmaps drive
the client's existing slideshow engine, package-free. Appearance
eligibility per parent §12a.3; over-budget slides degrade to
server-composed keyframes, never to originals with content withheld.

**Entry gate:** the full parent T5 battery (the Stage-0 run was a
feasibility subset). **Done when:** parent §12a.8 3c list — the corpus's
animated decks match local playback by pixel diff at fixed timeline
samples within parent §7.4 tolerance; build/transition/Morph step count
and order match local; the `step-composited` fallback rate is recorded.

### 4.5 Piece 5 — video and audio (parent 3d, now required)

Media plays during a show via server transcodes: mark-burned video,
metadata-free, through `src/slideshow/media/lease.ts`; never original
media bytes. Audio carries no visible mark; *my call*: audio is
transcoded metadata-free and audited, with no forensic audio mark (none
is specified anywhere; adding one is a new owner decision).

**Done when:** parent §12a.8 3d list; corpus media slides play with the
burned-in mark; a wire audit shows no original media bytes.

## 5. Constraints (from the parent, restated because they bite here)

- Fail-closed: anything the join rule cannot resolve is withheld.
- Protected by default; relaxations are server-side and audited; the
  client has no relax surface.
- The eager/core-entry gates (parent T11, `bun run gates:core`) hold:
  new client code for these services loads lazily on the remote path.
- Every new message kind or field gets an encoded and decoded cap
  (parent §14 table) and a hostile test.
- Reviewer for each piece's PR is a different model family from its
  implementer.

## 6. Left to the builder

- Wire shapes for search/copy/print requests and answers (text control
  messages, within the 64 KiB cap), and refusal codes beyond those named.
- How selection geometry maps to index ranges; how the selection is drawn
  on the remote view.
- The page-range picker's UI, and the print-only surface.
- Rate numbers for search/copy (a new class), within parent §14's style.
- Transcode formats and codecs for piece 5.
- Whether pieces 4 and 5 need their own detailed design before planning
  (likely for 4, given the T5 battery).
