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
