# Server streaming — Stage 3c validation (presenting, with notes)

Spec: `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.3 and its
2026-11-14 amendment (parent `docs/superpowers/specs/2026-11-07-server-streaming-design.md`
§6.1, §6.2 rule 7, §8.2, §12a.5, §12a.7, §12a.8 3b, §14, §19 item 12).
Plan: `docs/superpowers/plans/2026-11-14-streaming-s3c-present.md`.

What was built: `present()` on a streamed PPTX. The show is full screen,
one slide at a time, from the server's marked slide bitmaps. It has
forward, back and `goTo()`, and hidden slides are skipped. Presenter View
shows speaker notes as server-drawn, marked images, for a presenter viewer
only. There are no animations, transitions or media yet (pieces 4–5):
every slide shows its final composed image.

## Verification (2026-11-14)

Product code at `a7fba779`; tests and docs at the Task 8 commit. Machine:
Apple-silicon Mac.

| check | result |
|---|---|
| `bun run test:server` | **907 pass, 0 fail** (50 files, 7,347 expects) |
| `bun run typecheck:server` | clean |
| `bun run lint` (`tsc --noEmit`, the client typecheck; covers `tests/`, so the browser specs too) | clean |
| `bun run test` (`vitest run`) | full run: **13,409 passed, 1 failed**, 64 skipped, 1 todo (677 files). The failure is `tests/fc-t2-drawing-reply-grant.test.ts` "keeps an empty sheet ready under admissible pressure…" (an XLSX drawings grant assertion, `expected +0 to be 1884`). Re-run alone 4 times: **15/15 passed** each time. No XLSX code changed in this piece (`git diff b222a7ce..HEAD -- src/xlsx src/components/xlsx` is empty), so this reads as a load-dependent flake, not a regression; reported, not hidden |
| `bun run test:browser:stream` | **29 passed** (6 new in `stream-present.spec.ts`; fix round 1 re-run) |
| `bun run gates:core` | **FAIL, red as on `main`**: index closure 421,141 B (see below) |

The first browser run on this branch had 2 failures. Both were in existing
specs that still pinned the view as it was before 3c, which Task 7 changed
on purpose by adding a Present button to the remote view:

- **The "no document text" check** in `stream-viewer.spec.ts` "happy". The
  Present button's label is chrome, so the check now strips
  `[data-ov-present-button]` as it already stripped the Print button.
- **The focused-a11y tab order** in `stream-text.spec.ts`. Present now comes
  first in the tab order. That test runs with `?present=off`, which the
  harness page maps to `allowPresent={false}`.

The harness dev server (`vite.harness.config.ts`) also warms
`src/stream/remote/show.ts`. It is reached only through
`present-start.ts`'s dynamic import, and the locked-down server refuses
modules it has not seen ("outside of Vite serving allow list"). The other
lazy chunks are warmed the same way. The run above is after these changes.

## Size gate (`bun run gates:core`)

Index closure **421,141 B** on this branch (`FAIL index closure 421141 B (baseline 417436, Δ 3705, recorded 3088, limit 420540)`). The piece started at 421,068 B (Task 1 baseline), already over the **420,540 B** limit. During the branch it read 421,068 B for Tasks 1–4, 421,069 B for Tasks 5–6 (+1 B hash noise) and 421,235 B at Task 7's first commit, before the eager remote logic moved into the lazy `present-start.ts`.

Disposition (ruling R26, Task 7): the index closure grew by 72 B, which is
accepted as the irreducible eager cost of a remote `present()` entry:

- 40 B is the launch hook in `components/present.ts`, through which a
  caller-prepared start is injected;
- 32 B is the `OfficeDoc` ref bridge, which gives a typed refusal before
  the lazy view mounts.

Everything else is lazy:

- the remote checks (`stream/remote/present-start.ts`, imported only by the
  lazy `RemoteDocView` chunk);
- the show (`stream/remote/show.ts`, a dynamic import after the surface is
  up);
- the fetcher and the notes client.

The gate is reported red, as on `main`. The baseline is untouched.

## §12a.8 3b checklist

| line | proved by |
|---|---|
| present a streamed deck | browser: `stream-present.spec.ts` "audience: present, step all 6 slides…" (started from a real click; full screen asked for in the gesture); `tests/stream/remote-present-start.test.tsx` (gesture-required without activation; host shown and `requestFullscreen` called **before** the lazy show module resolves; never the local chunk) |
| forward, back, `goTo()` | browser: → through all 6 slides with the counter "Slide n of 6", the end slide, digits + Enter to slide 3, Esc exits (`ended`); ← in the hold test; `tests/stream/remote-show.test.tsx` "order and navigation" |
| hidden slides never appear and the counter excludes them | `tests/stream/remote-show.test.tsx` "the order skips hidden units and the counter excludes them", "goTo(2) with unit 1 hidden shows unit 3"; `remote-present-start.test.tsx` "from is a streamed index", "a deck whose every unit is hidden is not-presentable"; `server/test/present-policy.test.ts` "hidden slides" (`'full'` flags them, `'visible-only'` drops them); `tests/slideshow-presenter-remote-adapters.test.tsx` (the grid leaves hidden slides out). Not in the browser: no corpus deck has a hidden slide. |
| Presenter View shows notes only for a presenter viewer | browser: "presenter: Presenter View swaps in single-screen with a painted notes image…" (the pane canvas is 720 px wide; slide 1's empty-notes image is fully opaque, over 80 % near-white paper, with no dark ink in the body region (marks only); slide 5's image is taller, fully opaque, over 80 % near-white, with more than 500 opaque dark (R, G, B < 128) pixels inside the 24 px-padded body region, polled until painted. A predicate test shows that a transparent canvas and a white canvas with a light marks-like tint both fail, and drawn text passes); "audience: …" and "notes=off presenter: …" (Alt+F5 Presenter View shows "No notes for this slide" and sends no `notes`); `remote-present-start.test.tsx` (`presenterView: true` with `can.notes` false is `not-presentable`: no `window.open`, no full screen) |
| … and never on the audience surface | browser: the audience surface holds exactly its one canvas, never a notes canvas, before and after Alt+F5; `remote-show.test.tsx` "a presenter: notes only for the settled current slide, never on the audience canvas" (audience pixels unchanged by a notes delivery) |
| a non-presenter gets no notes from any request | `server/test/notes-wire.test.ts` "refusals send no frame and write no audit row" (audience, presenter with `notes: false`, presenter resumed to `notes: false`, docx, hidden unit); `tests/stream/remote-notes.test.ts` "can.notes false → policy-refused before any send"; `remote-show.test.tsx` "an audience session (can.notes false) sends no notes request across a full run"; browser: zero `notes` messages sent for an audience or `notes=off` viewer |
| no PPTX, notes text outside the a11y service, or media bytes cross the wire | `server/test/notes-wire.test.ts` "the kind-7 bodies, the manifest and every text message carry none of the notes text"; browser: slide 5's notes phrases appear in no text message received (below) |
| both marks on every slide and notes image | `server/test/notes-evidence.test.ts` (the image decodes the session's Layer-1 codeword at scales 1 and 2; the visible mark differs from a mark-off draw; `trace({target: 'notes'})` and `detect --notes` attribute a leaked notes image); an image that cannot carry Layer-1 is refused, not sent (R8). Slides are the existing marked units (Stages 2a/2b). |
| revocation mid-show exits cleanly | browser: "revocation mid-show exits the show cleanly…" (`/revoke` → close 4005, `ended`, no surface, no Presenter View); `remote-show.test.tsx` "doc.ended exits the show; a later ready paints nothing"; `remote-notes.test.ts` "a terminal end (revoked) rejects a pending request…"; `remote-present-start.test.tsx` "a session that ends while the show module is still loading cancels the start at once" |

## Wire audit

- **Server side** (`notes-wire.test.ts`, deck `lo-tdf104788.pptx`). For
  each notes paragraph, the first 12 characters are searched for in every
  kind-7 body, the manifest and every text message the session receives.
  None is found.
- **Browser side** (`stream-present.spec.ts` "presenter: …"). Slide 5's
  notes phrases ("Small Basic", "different elements of the environment",
  "multiple Editor") appear in none of the text messages received. Binary
  frames are the marked PNG rasters.
- **Client side.** The client never sends notes text and has none to send:
  notes arrive as a kind-7 image only. A non-presenter's client never sends
  `{type: 'notes'}`, and a hand-written one is refused by the server
  (`policy-refused`, no frame, no audit row).
- **No PPTX or media bytes.** The show requests only `unit` rasters and
  `notes` images, and media is out of this piece.

## Rate measurement (the hold-→ test)

`stream-present.spec.ts` "holding the arrow keys…". Audience viewer,
`lo-tdf104788.pptx`, about 30 keydowns/s, with `repeat` set while a key is
held. A held → past the last slide would end the show, because this deck
does not loop. So → is held to slide 6, then ← back to slide 1, alternating
for 10 s. The test then stops on slide 4 and asserts:

- `data-slide=3` and `data-unit=3`;
- the frame is an image and the counter reads "Slide 4 of 6";
- the show is still `started`, with no stream error;
- no close code was recorded by the WebSocket wrapper (so no 4008), and no
  `rate-limited` refusal was received.

Measured (the test's log line): `hold: 264 key events in 10 s; unit requests from page load 11 (peak in any 60 s 11); rate refusals 0; close codes none` (earlier runs: 261–264 key events, 9–11 unit requests)

**How much this proves.** On a 6-slide deck it shows little about the rate
bound: the show's 8-bitmap cache holds the whole deck, so a long hold
issues only the first sweep's requests. The bound itself is proved by
`tests/stream/show-fetch.test.ts`:

- "200 want() calls of fast stepping over 30 s on a 40-unit deck stay
  ≤ 48 requests in any 60 s window";
- "R14 pacing (a)": a held → for 5 s, then a stop;
- "reviewer model: every paint charges twice …, real SessionRates — no
  close over 60 s of 400 ms navigation".

Those cases run one fetcher alone. The final review found that the
server's `'unit'` window lives as long as the session (sid), resume
included, while each show had its own budget: restarting a show inside a
minute, or viewer paints before it, could reach 4008 (R28, below). The
session-level bound is proved by `tests/stream/unit-budget.test.ts`. It runs
a real `StreamSession` and the real show fetcher against a fake server that
applies the real `SessionRates`:

- the reviewer's probe: 40 slides at 400 ms, dispose, restart the show on
  the same session — no refusal, no close, ≤ 48 sends in any 60 s;
- 30 viewer paints, then a 40-slide show — no refusal;
- 70 viewer paints at once — 48 go out, the rest wait for room; no refusal;
- resume mid-show with a request in flight, then a restarted show — the
  log survives the resume, no refusal;
- a server window already full at open (history the client cannot know):
  one refusal, then no unit send for 62 s — no second breach, no close.

The corpus has no deck with more than 6 slides.

## Owner-facing calls (plan C1–C4)

- **C1 — hidden slides under `contentPolicy: 'full'`.** Under
  `visible-only`, hidden slides are not units at all. Under `'full'` they
  are units (they print, Stage 3b P5), and the manifest flags each one
  `units.list[i].hidden: true` so the show skips it. The flag leaks
  nothing, because under `'full'` the slide is already viewable as a unit.
- **C2 — the notes look.** The image is the notes body text, the same text
  local Presenter View shows, drawn by the server as plain text on white.
  It is not the notes-page layout.
  - Width: 720 CSS px × scale.
  - Height: fits the text, up to 8,192 device px.
  - Longer text ends with the drawn line "Notes continue — too long to show
    in full.", so nothing is cut silently.
  - A slide with no notes gets a valid white image carrying the marks only.
- **C3 — `presenterId` is the session id.** No notes are shared across
  sessions.
- **C4 — the server never caches notes across requests.** Every request is
  rendered and audited again.

## Rulings that change behaviour or limits

In plain words; the ledger has the full list (R1–R26).

- **R3 — notes on/off.** Only `notes: undefined` or `true` turns notes on
  for a presenter. Any other value (`false`, `null`, `'yes'`, `1`) turns
  them off (fail closed).
- **R4, R5 — refusal shapes.**
  - A notes refusal is `{reqId, type: 'error', code}`, where `code` is
    `policy-refused`, `invalid-unit`, `rate-limited` or `session-rotated`.
    `session-rotated` means a stale token; the client re-issues once.
  - The generic per-message gate still answers `{type: 'refused', code:
    'rate-limited'}`.
- **R6 — no `truncated` flag on the wire.** The continuation line is drawn
  into the image.
- **R7, R8 — audit and marks.**
  - The `notes_deliveries` row is written before the draw. A draw that
    then fails leaves a row with nothing sent (over-audit only).
  - A notes image that cannot carry Layer-1 is refused, which is stricter
    than slides (Stage 3b P6).
- **R9 — an unsolicited or late kind-7 frame** is dropped as
  `unexpected-reqId`: nothing paints and the session stays alive.
- **R10 — notes admission is sized from the declared image.** `notes-start`
  declares `w × h`. The grant is `min(encoded bound, 8 MiB) + 2 × (16 +
  w·h·4)`, and the frame must match the declaration.
- **R11 — no admission room.** A notes request refused for lack of
  admission room is retried once, after the show's slide requests have
  drained. A second failure shows "Notes could not be loaded".
- **R14 — pacing.**
  - A navigation target goes out at once, unless another target went out
    under 400 ms ago; then it waits for 250 ms of stability.
  - Prefetch waits until the target is ready or has failed, and needs 1 s
    of stability.
  - Fast keying adds at most about 650 ms of latency.
- **R16 — every server-charged send takes budget** (validation retries
  included).
- **R28 — the budget is the session's** (final review, Important 1).
  - One sliding send log per `StreamSession`. It lives as long as the
    server sid, so it survives show restarts and resume.
  - The session stamps every server-charged unit send, from any caller:
    viewer paints, show fetches, validation and replay re-sends, and resume
    re-sends. It does not stamp the stale-jti re-send, which the server
    refuses before its rate gate. The show's `charge()` only marks its
    request as sent, so nothing is counted twice.
  - The client window is 62 s: the server's 60 s plus 2 s, because the
    server stamps a request when it arrives.
  - A first send waits, first in first out, until fewer than 48 sends are
    in the window. A re-send of a request already sent waits for fewer
    than 56. A viewer paint over budget therefore waits for room. It is
    never sent into a full window.
  - The show reads the same budget (`unitBudget`) and keeps its R14 pacing
    and its 8-request prefetch reserve against it. Viewer paints are not
    held back by the reserve, but they count.
  - After a `rate-limited` unit refusal the session sends no charged unit
    request, from any caller, until the log shows room and 2 s have
    passed. The log never holds 60 sends, so a refusal means the server
    counted sends the log never saw. That history is unknown, so the
    window is treated as full for a whole 62 s span. After that, every send
    the server counts is one the log stamped. Resume does not clear the
    log: the server keeps the sid's window across resume, and the client
    stamped every send it made, including the ones in flight at the drop.
- **R17, R19 — hidden slides in the counts.**
  - Presenter View's grid leaves hidden slides out.
  - `slideIndex`, `goTo()`, `from` and the counter count shown slides only.
- **R20 — notes are asked for only after a slide has been current for
  300 ms.** The server has no notes-cancel, so fast stepping sends no notes
  for slides passed over.
- **R21 — the slide cache** is sized at start from half the free show
  workspace at the largest bucket (at most 8 bitmaps; fewer neighbours are
  prefetched when that is less).
- **R22 — workspace reservation.** Each notes request reserves the
  worst-case image in the show workspace first. At scale 2 that is
  1,440 × 8,192 × 4 B = 47.2 MB (45 MiB). With less free, the notes are
  refused.
- **R24 — `presenterView: 'auto'` without `can.notes`** opens no presenter
  window. A non-presenter never gets a blank `window.open`.
- **R25 — `present()` before the remote view has mounted** is refused
  `not-presentable`.
- **R26 — eager growth.** +72 B eager (above).

## Honest limits (recorded)

- **R12 — maximal notes at scale 2.** The tallest scale-2 notes image has
  an inflate peak of about 98 MiB. Admission clamps it to the 96 MiB
  session aggregate, so it is admitted only when nothing else is in flight.
  Otherwise it is refused, then retried once (R11).
- **R15 — long non-stop keying.** More than about 19 s of non-stop target
  changes can use up the session's 48/min budget. Viewer paints and an
  earlier show in the same minute use it up sooner. The show then waits
  for the window to free a slot ("Loading slide n" meanwhile).
- **R28 — what "no 4008" rests on.** The session never sends a unit
  request into its own full window, and after a refusal it holds every unit
  send for a whole window. The tests above show no close with the real
  rate gate. What they cannot show: the client's clock is not the
  server's. The 2 s receipt margin and the 8-send gap between 48 and 56
  absorb normal latency. A delay longer than that, followed by a burst,
  can still produce one refusal. That refusal starts the 62 s hold, so a
  second and third breach would need the server to count sends the client
  never made.
- **R28 — a refusal costs a minute.** One `rate-limited` unit refusal
  stalls all unit painting (viewer and show) for 62 s.
- **R16 — accounting.** Charge accounting is conservative, not exact.
  Every server-charged send is counted, retries included, so the show may
  hold back when the server would still accept.
- **R20 — notes latency.** Notes appear at least 300 ms after a slide
  becomes current, plus the render.
- **R22 — notes need room.** Notes need about 45 MiB free in the show
  workspace at scale 2 (about 22.5 MiB at scale 1).
- **R26 — the gate.** `gates:core` is red by 72 B more than before this
  piece.
- **Notes accessibility is out.** The notes image has
  `aria-label="Speaker notes (image)"` and no text alternative. Its text is
  not served as text to anyone in this piece (§19 item 12).
- **Hidden flags are fixed at show start.** A resume does not change which
  slides are hidden: `hidden` comes from the package at session open, and
  `role` is a resume invariant.
- **Per-kind buffering (deferred minor, pre-existing for units too).** The
  client buffers up to the 8 MiB per-kind encoded cap before the
  per-request decoded cap rejects a frame.

## Known limits not machine-checked

- **Two-screen popup placement.** The browser spec stubs `window.open` to
  return `null` (a blocked popup), so Presenter View runs single-screen,
  swapped into the show host. The presenter window on a second screen
  (Window Management placement, `moveTo`/`resizeTo`, a real popup's
  lifetime) is covered only by the local slideshow specs and the unit tests
  in `remote-present-start.test.tsx`; it is not run against a stream.
- **Real full screen and a real projector.** `requestFullscreen` is stubbed
  to reject, so the show runs in the window, and the spec checks only that
  it was requested inside the click. An actual full-screen show, a
  projector or display change mid-show, and DPR changes on a real second
  screen are not checked. The bucket change on resize is unit-tested
  (`remote-show.test.tsx` "scale bucket").
- **Rate on a long deck in a browser.** See the rate measurement above.
- **Hidden slides in a browser.** No corpus deck has one; the server and
  client unit tests cover it.

## Test map

| what | where |
|---|---|
| `can.present`, `can.notes`, R3 notes values, resume (notes mutable, role not), hidden flags, hidden detection failure | `server/test/present-policy.test.ts`, `tests/stream/manifest.test.ts` |
| the `notes` op: presenter path, empty notes, every request re-audited (C4), refusals with no frame and no row, check order, token / stale jti, rate 30/min (31st refused, socket open), encoded cap, wire audit, long notes (8,192 px with the continuation line), reader-cap continuation, bounded measurement | `server/test/notes-wire.test.ts` |
| both marks on a notes image; production trace and `detect --notes` | `server/test/notes-evidence.test.ts` |
| `services.notes` (and its `ALTER TABLE` migration; old rows read notes off) | `server/test/audit-store.test.ts` |
| `notes_deliveries` rows (one per request that passes the checks and the plan, written before the draw; a refusal before that step writes none, a failure after it keeps the row, which records the request, not a delivery, R7) and `notes_marks` rows (one per marked image, written just before its send) | `server/test/notes-wire.test.ts`, `server/test/notes-evidence.test.ts` |
| client notes request: kind-7 admission (R10), hostile frames, unsolicited frames (R9), caps, abort, deadlines, rotation, resume, end barriers and image close (R13) | `tests/stream/remote-notes.test.ts` |
| show fetcher: target first, the 48/min sliding window, reserve, pacing (R14), charge per send (R16), LRU, bucket change, dispose | `tests/stream/show-fetch.test.ts` |
| the session's unit send log (R28): one per sid, every charged send stamped, viewer paints wait for room, show restart and resume, hold after a refusal, the real `SessionRates` | `tests/stream/unit-budget.test.ts`, `tests/stream/unit-send-log.test.ts`, `tests/stream/remote-show.test.tsx` ("R28") |
| Presenter View over an image-backed source (counts, Next, grid, go-to, image notes) | `tests/slideshow-presenter-remote-adapters.test.tsx` |
| remote show session: order, errors, end, notes, bucket, workspace leases, exit inside dispatch, refresh, pane copy | `tests/stream/remote-show.test.tsx` |
| `present()` wiring: gesture, refusals, popup and full-screen order, lazy loading, Present button, F5, ref | `tests/stream/remote-present-start.test.tsx` |
| browser end to end | `tests/browser/stream-present.spec.ts` (below) |

**Browser end to end** (`tests/browser/stream-present.spec.ts`, Playwright
project `stream`, against `server/test/app.ts` with `?role=presenter` and
`?notes=off`). Deck: `lo-tdf104788.pptx`, 6 slides; the notes bodies of
slides 1–4 are empty and slides 5–6 have text. An init script makes
three browser parts deterministic:

- `window.open` is recorded and returns `null`;
- `Element.prototype.requestFullscreen` is recorded and rejects;
- every WebSocket's close code is recorded.

The tests:

- **Audience.** Present opens with one full-screen request and no
  `window.open`. → steps slides 1–6 ("Slide n of 6", an image frame each),
  then the end slide. `3` + Enter goes to slide 3. Alt+F5 Presenter View
  shows "Slide 3 of 6" and "No notes for this slide". Esc ends the show.
  Zero `notes` sent, no close.
- **Presenter.** Presenter View (`presenterView: true`) opens the
  provisional window in the click; it is blocked, so Presenter View swaps
  in and the audience surface is hidden. Slide 1's notes pane canvas is
  720 px wide, fully opaque and mostly white paper, with no dark ink in the
  body region (its notes body is empty: marks only). `5` + Enter in
  Presenter View requests slide 5's notes; the test polls until that image
  is taller, fully opaque, over 80 % near-white, and has more than 500
  opaque dark pixels inside the body region. The audience surface holds only
  its own canvas throughout. Alt+F5 back shows slide 5 with no notes canvas
  anywhere. No refusal, no stream error, no close; the notes phrases are in
  no text message.
- **Hold.** As in the rate measurement above.
- **`notes=off` presenter.** Presenter View is refused `not-presentable`,
  with no `window.open`, no full-screen request and no surface. Present
  then works: → to slide 2, and Alt+F5 shows "No notes for this slide".
  Zero `notes` sent.
- **Paint predicate.** The same in-page measurement rejects a transparent
  720 × 200 canvas and a white one carrying only a light marks-like tint,
  and accepts one with drawn text (the negative controls for the check
  above).
- **Revocation.** `/revoke` mid-show: close 4005, `ended`, and no surface
  or Presenter View left.
