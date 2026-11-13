# Streaming Stage 3 piece 3 — present a streamed deck, with notes — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `present()` works on a streamed PPTX: full-screen slide-by-slide show from server-marked slide bitmaps, forward/back/`goTo()`, hidden slides skipped, and Presenter View with server-drawn, marked speaker-notes images for a presenter viewer only.

**Architecture:** The server already serves a marked bitmap per slide (unit). This piece adds (a) manifest facts the show needs (`can.present`, `can.notes`, per-unit `hidden`), (b) a `notes` request answered by a `NOTES_RASTER` (kind 7) image drawn from the notes slide's body text in the original package, gated by role + policy and audited, and (c) a lazily loaded client **remote show** that drives the existing audience surface and the existing Presenter View through additive adapters, fetching slides through a cache that stays inside the server's unit rate limit. No animations, transitions or media (pieces 4–5): every slide shows its final composed image.

**Tech Stack:** TypeScript, React 18, Bun server (`server/`), vitest, Playwright.

**Spec:** `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.3 (parent: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` §8.2, §9.3, §12a.5, §12a.7, §12a.8 3b, §13.1, §14). Builds on main 283940f5 (pieces 1–2 merged).

## Global Constraints

- Notes only when `role === 'presenter'` **and** `can.notes`; Presenter View only; never on the audience surface; a non-presenter gets no notes from any request (§12a.5, §12a.8).
- `can.present` is true for every PPTX session, false otherwise; `can.notes` = presenter role ∧ `policy.notes !== false` (design §4.3). `can.notes` is re-evaluated on resume (§6.2 rule 7); `role` stays a resume invariant.
- No notes text, PPTX bytes or media bytes on the wire; notes travel as images only (§12a.8). Notes a11y text is **out** of this piece (honest limit, §19).
- Both marks (visible + Layer-1) on every slide and notes image (§13.1).
- Hidden slides never appear and are not counted; `slideIndex`/`goTo()` count streamed slides only (§12a.7).
- Show traffic stays within the `'unit'` rate class (60/min; third breach in 60 s closes 4008). A show is not a bulk-export path (§12a.7).
- Prefetch: current + 2 next + 1 previous; `goTo()` fetches its target first (§12a.7).
- Gesture: the surface opens synchronously inside the user gesture; content streams after (§8.2, §12a.7). `startPresent` reads the manifest synchronously and refuses `presenterView: true` when `can.notes` is false (§8.2).
- A slide that fails to stream shows a typed per-slide error frame and the session stays alive; revocation/expiry exits the show and `ended` resolves; never local rendering (§12a.7).
- NOTES_RASTER caps: 8 MiB encoded / 64 MiB decoded (§14). Every new message/field gets an encoded and decoded cap and a hostile test (design §5).
- New client code for the remote show loads lazily; record `bun run gates:core` per task (baseline 421,068 B, already over its 420,540 B limit; CI does not run it). Growth beyond lazy-chunk hash noise (±a few bytes) is a finding.
- Local presenting is unchanged: existing slideshow/presenter tests and goldens stay green.
- Reviewer for each task is a different model family from its implementer.

## Owner-facing calls in this plan (*my call*, flagged at handoff)

- **C1 — hidden slides under `contentPolicy: 'full'`.** Under `visible-only` hidden slides are not units at all. Under `'full'` they are units (and print prints them, ruling P5). The show must still skip them, so the manifest flags them: `units.list[i].hidden?: true`. Flagging leaks nothing: under `'full'` the slide is already viewable as a unit.
- **C2 — notes look.** The notes image is the body-placeholder text (what local Presenter View shows, `readNotes`) drawn by the server as plain text on white, not the notes-page layout. Width 720 CSS px × scale; height fits the text up to 8,192 device px; text beyond that ends with the line "Notes continue — too long to show in full." (no silent cut).
- **C3 — `presenterId` = the session id.** Notes cache namespace `notes:<sid>`; no notes cache sharing across sessions (§9.3 "two presenters never share a tuple").
- **C4 — notes are not cached across requests on the server**; each request renders and audits.

## Review Focus

1. Fast navigation (holding →, typing a slide number and Enter repeatedly) on a 40-slide deck must never trip the 4008 rate close; the show must slow its fetching, not die.
2. A non-presenter (or presenter with `notes:false`, or after a resume that turns notes off) sending a hand-written `{type:'notes'}` gets a refusal and no frame.
3. Revocation or expiry mid-show (and mid-notes-request) exits the show cleanly; nothing paints after exit; no local rendering.
4. A slide whose raster fails shows a typed error frame; next/previous still work.
5. Resize / screen change mid-show re-requests at the new scale bucket and shows the old bitmap scaled meanwhile.

---

### Task 1: Server policy — `can.present`, `can.notes`, hidden flags

**Files:**
- Modify: `server/src/hooks.ts` (AuthDecision.policy `notes?: boolean`; EffectivePolicy `notes: boolean`; `resumeInvariant` excludes `notes`)
- Modify: `server/src/ws.ts` (`can` at :327; resume change recording)
- Modify: `server/src/render.ts` or `server/src/policy/pptx.ts` (per-unit hidden flag under `'full'` from the slide root `show="0"`)
- Modify: `server/src/audit.ts`, `server/src/audit-store.ts` (`services.notes` column, ALTER-if-missing like `print`)
- Modify: `src/stream/manifest.ts` (validate `units.list[i].hidden` as optional `true`; existing `can.present/notes` bools)
- Modify: fixtures hardcoding `present: false, notes: false` (tests/stream/*.test.ts*, server/test/handshake.test.ts:29)
- Test: `server/test/present-policy.test.ts`, `tests/stream/manifest.test.ts`
- Test fixture: synthetic PPTX with one `show="0"` slide (build with the existing testdata builders; no corpus deck has one)

**Interfaces:**
- Produces: `EffectivePolicy.notes: boolean`; `canNotes(p: EffectivePolicy): boolean` (= `p.role === 'presenter' && p.notes`) exported from hooks.ts; manifest `can.present = format === 'pptx'`, `can.notes = canNotes(p)`; `ManifestUnit.hidden?: true`.

- [ ] **Step 1: Failing tests**
  - `present-policy.test.ts`: pptx session → `can.present === true`; docx → false. Audience role → `can.notes === false`; presenter → true; presenter + `notes:false` → false; non-boolean `notes` ('yes', 1, null) → treated as default (true for presenter) only when `undefined`, otherwise false (fail closed: anything but `undefined`/`true` is off).
  - Resume: presenter resumes with `notes:false` → accepted, new manifest `can.notes === false`, `services` row records `notes:false`; role change on resume → still 4003 (`auth-failed`).
  - Hidden: `'full'` policy on the synthetic deck → manifest unit 1 has `hidden: true`, others have no `hidden` key; `'visible-only'` → hidden slide absent from units, no `hidden` keys.
  - `manifest.test.ts`: `hidden: true` accepted; `hidden: false`/`'x'`/`1` → manifest rejected (fail closed).
- [ ] **Step 2: Run** `bun run test:server -- present-policy` and `bunx vitest run tests/stream/manifest.test.ts` → FAIL.
- [ ] **Step 3: Implement** as in Interfaces. Hidden detection reads `p:sld/@show` from each slide part of the original package once at session open (not from the projection).
- [ ] **Step 4: Run** both suites plus `bun run test:server` and `bunx vitest run tests/stream` → PASS.
- [ ] **Step 5: Commit** `feat(stream): manifest can.present, can.notes and hidden slide flags`

### Task 2: Server notes op — `NOTES_RASTER`

**Files:**
- Create: `server/src/notes.ts` (notes text read, layout, draw)
- Modify: `server/src/render.ts` / worker job table (a `notes` job given `{unitIndex, scale}` → raster)
- Modify: `server/src/ws.ts` (`KIND` adds 7, `ENCODED[7] = 8 MiB`; `{reqId,type:'notes',token,unit,scale}` op)
- Modify: `server/src/rate.ts` (notes ride a new `'notes'` class: 30/min, refuse-only, no close)
- Modify: `server/src/audit-store.ts`, `server/src/audit.ts` (`notes_deliveries(sid, unit_index, scale, chars_drawn, at)`; `AuditRow.notesShown`)
- Modify: `server/test/harness.ts` (CAPS kind 7; `Client.notes(unit, scale)`)
- Test: `server/test/notes-wire.test.ts`, `server/test/notes-evidence.test.ts`

**Interfaces:**
- Consumes: `canNotes` (Task 1), `obtainUnit`/`deliverRaster` pattern and `paintMark` (existing).
- Produces: wire reply `notes-start {reqId, w, h, scale}` then one kind-7 frame (`u32 LE unitIndex` + raster body, same body codec as UNIT_RASTER) — or a refusal `{reqId, type:'error', code}` with `code ∈ 'policy-refused' | 'invalid-unit' | 'rate-limited'`. `drawNotes(paragraphs: NoteParagraph[], scale: 1|2): {w,h,truncated}` in notes.ts.

Rules (in this order, all before any render): token valid → `canNotes(policy)` else `policy-refused` → unit in range and not hidden else `invalid-unit` → rate → audit row written → render → Layer-1 + visible mark → send. Slides with no notes part or empty body answer a valid empty-notes image (white, mark only) so absence is not distinguishable from a refusal by size of work, and is audited with `chars_drawn: 0`.

- [ ] **Step 1: Failing tests** (`notes-wire.test.ts`, deck `lo-tdf104788.pptx`, 6 notes slides):
  - presenter gets `notes-start` + one kind-7 frame for unit 0; decoded width = 720×scale; audit `notes_deliveries` has one row.
  - audience role: `{type:'notes'}` → `policy-refused`, no frame, no audit row. Presenter with `notes:false` → same. Presenter resumed to `notes:false` → same.
  - hidden unit (synthetic deck, `'full'`) → `invalid-unit`. Out-of-range / non-integer unit → `invalid-unit`.
  - notes rate: 31st request in a minute → `rate-limited`, socket stays open.
  - wire audit: the kind-7 body and every text message for the session contain none of the notes text (search the raw bytes for each notes paragraph's first 12 chars).
  - long notes (synthetic 2,000 paragraphs) → height 8,192, `truncated` and the continuation line present (pixel probe: last 60 rows not white).
- [ ] **Step 2: Failing evidence test** (`notes-evidence.test.ts`): the notes image decodes the session's Layer-1 codeword (existing detector) and carries the visible mark (pixel diff against an unmarked draw is non-zero in the mark region).
- [ ] **Step 3: Run** → FAIL.
- [ ] **Step 4: Implement.** Read notes server-side with the same rule as `src/slideshow/read/notes.ts` (body placeholders only, 65,536-char cap; import it if it runs under Bun, otherwise share the pure part).
- [ ] **Step 5: Run** both tests + `bun run test:server` → PASS.
- [ ] **Step 6: Commit** `feat(stream): server-drawn marked speaker notes for presenters`

### Task 3: Client notes request

**Files:**
- Modify: `src/stream/remote/session.ts` (`notes(unit, scale, signal)`; kind-7 admission only for a live notes request, grant = encoded cap + decoded body; `CAPS` kind 7: 8 MiB / 64 MiB)
- Modify: `src/stream/remote/remote-doc.ts` (`SessionLike.notes?`; `RemoteOfficeDoc.paintNotes(unit, dpr, canvas, signal): Promise<{truncated:boolean}>` — throws `RemoteError('policy-refused')` when `!manifest.can.notes`, before any send)
- Test: `tests/stream/remote-notes.test.ts`

- [ ] **Step 1: Failing tests:** happy path paints the canvas at the decoded size; `can.notes` false → rejects without sending; unsolicited kind-7 frame (no live request) → session error `protocol` and no paint; kind-7 for a different unit than asked → rejected; frame over decoded cap → rejected, admission released; abort mid-flight → no paint, reservation released; `doc.ended` → pending notes reject `ended`.
- [ ] **Step 2: Run** `bunx vitest run tests/stream/remote-notes.test.ts` → FAIL.
- [ ] **Step 3: Implement** mirroring `print()` (pending map, `Reassembler.drop` on settle).
- [ ] **Step 4: Run** + `bunx vitest run tests/stream` → PASS.
- [ ] **Step 5: Commit** `feat(stream): client notes request with kind-7 admission`

### Task 4: Show slide cache and request budget (pure)

**Files:**
- Create: `src/stream/remote/show-fetch.ts`
- Test: `tests/stream/show-fetch.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export interface ShowFetcher {
    /** Bitmap for `unit` at the current bucket, or the best cached one at another bucket (scaled), or undefined. */
    cached(unit: number): ImageBitmapLike | undefined
    /** Wanted set changes on every navigation: `[target, ...prefetch]`, target first. */
    want(units: readonly number[]): void
    setBucket(scale: number, dpr: number): void
    onReady(cb: (unit: number) => void): () => void
    onError(cb: (unit: number, code: RemoteErrorCode) => void): () => void
    dispose(): void
  }
  export function createShowFetcher(io: { paint(unit: number, scale: number, dpr: number, signal: AbortSignal): Promise<ImageBitmapLike> }, clock: { now(): number }, opts?: { budgetPerMin?: number /* default 48 */, maxCached?: number /* default 8 */ }): ShowFetcher
  ```
  - Token bucket of 48 requests / 60 s (headroom under the server's 60). The first entry of `want` is always next in line; prefetch entries are only issued while budget remains above 8 (reserve for navigation targets). Units no longer wanted are aborted if not started. LRU of 8 bitmaps; a bucket change keeps old bitmaps as fallbacks until replaced.

- [ ] **Step 1: Failing tests** (fake clock, fake io): target issued before prefetch; 200 `want()` calls over 30 s of simulated fast stepping on a 40-unit deck issue ≤ 48 requests in any 60 s window; budget at reserve → prefetch withheld, target still issued; dropped-from-want unstarted requests aborted; bucket change re-requests the current unit and `cached()` returns the old bitmap meanwhile; error on a unit → `onError` once, no retry storm (retry only on the next `want` that includes it, after ≥ 2 s).
- [ ] **Step 2–4:** run → FAIL, implement, run → PASS.
- [ ] **Step 5: Commit** `feat(stream): show fetcher within the unit rate budget`

### Task 5: Presenter View adapters (additive)

**Files:**
- Modify: `src/slideshow/presenter/PresenterView.tsx` (PresenterSource gains optional `slideCount?: number`, `slideImage?(index: number, canvas: HTMLCanvasElement): boolean`, `hidden?(index: number): boolean`, and `notes()` may return `{ kind: 'image'; paint(canvas: HTMLCanvasElement): void; truncated: boolean } | readonly NoteParagraph[]`)
- Modify: `src/slideshow/presenter/notes.tsx` (image variant: a canvas in the same pane, `aria-label="Speaker notes (image)"`, not selectable)
- Modify: `src/slideshow/presenter/preview.ts` / grid only where they read `deck`/`model`/`previewOf` (5 sites: PresenterView.tsx:135,137,194,315,332)
- Test: `tests/slideshow/presenter-remote-adapters.test.tsx`

Rule: when `slideImage` is present, the Next preview and grid thumbnails call it instead of `previewOf/paintPreview`; when `slideCount` is present it replaces `deck.slides.length`; `hidden` replaces `model.index.slides[i].hidden`. `deck` and `model` become optional in the type; local sessions keep supplying them, so local behaviour is byte-identical.

- [ ] **Step 1: Failing tests:** a fake source with `slideImage` + image notes renders Next preview via `slideImage(after)`, grid via `slideImage(i)` skipping `hidden(i)`, counter "Slide n of slideCount", notes canvas painted, no `previewOf` call.
- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run** the new test **and** `bunx vitest run tests/slideshow` (all local presenter tests unchanged) → PASS.
- [ ] **Step 5: Commit** `feat(slideshow): Presenter View accepts image-backed slide sources`

### Task 6: Remote show session (lazy)

**Files:**
- Create: `src/stream/remote/show.ts` (`runRemoteShow(req): Promise<SessionController>`)
- Test: `tests/stream/remote-show.test.tsx`

**Interfaces:**
- Consumes: `createAudienceSurface` (src/slideshow/player/surface.tsx:35), `order.ts` (`visibleOrder/nextVisible/previousVisible`), the keys/input mapping, `mountPresenterView`, timer, placement; Task 3 `paintNotes`; Task 4 fetcher; Task 5 adapters.
- Produces: a `SessionController` whose `state.order` = non-hidden unit indices, `stage`/`cursor` undefined, `build` always 0, `presenterSource` with `slideCount`, `slideImage`, `hidden`, image notes (only when `manifest.can.notes`; otherwise `notes()` returns `[]` and **no notes request is ever sent**). `next/previous/goTo/exit/ended`, black/white screens, end-of-show slide, loop/kiosk inputs as local.

Behaviour: the audience canvas draws the cached bitmap for the current unit fitted to the surface (letterboxed); missing → the best cached bucket scaled, else a neutral "Loading slide n" frame; an `onError` for the current unit → typed error frame ("Slide n could not be shown (code)") and navigation keeps working. `goTo(i)` validates against `order.length` and maps streamed index → unit. `doc.ended` or session revocation → `exit()`, `ended` resolves, and no further paint. Bucket = audience backing size × DPR via the existing `'fit'` scale rule; resize → `fetcher.setBucket`.

- [ ] **Step 1: Failing tests** (fake RemoteOfficeDoc + fake fetcher clock): order skips hidden units and counter excludes them; `goTo(2)` on a deck with unit 1 hidden shows unit 3; next past the end → end slide; error frame on failed unit then `next()` still advances; `doc.ended` resolves `ended` and later `ready` events paint nothing; audience role (`can.notes` false) → zero notes calls across a full run; presenter → notes requested only for the current slide and never drawn on the audience canvas (audience canvas pixels unchanged by notes delivery); resize → bucket change.
- [ ] **Step 2–4:** run → FAIL, implement, run → PASS.
- [ ] **Step 5: Commit** `feat(stream): remote slide show session`

### Task 7: Wiring — `present()` on streamed documents

**Files:**
- Modify: `src/components/present.ts` (`startPresent` remote branch: synchronous checks from the manifest — `can.present`, unit count after hidden filtering > 0, `from` in range, `presenterView === true && !can.notes` → `not-presentable`; then the same popup/fullscreen/token sequence; loads `src/stream/remote/show.ts` lazily)
- Modify: `src/components/OfficeDoc.tsx` (remote handle: replace the SHEET handle with one whose `present` calls `startPresent` on the remote doc; show host div as for local)
- Modify: `src/components/remote/RemoteDocView.tsx` (Present button when `can.present`, F5 shortcut; host)
- Modify: `src/debug/stream-viewer.tsx` (ref + Present button), `server/test/app.ts` (`?role=presenter`, `?notes=off`)
- Test: `tests/stream/remote-present-start.test.tsx`

- [ ] **Step 1: Failing tests:** remote `present()` without activation → `gesture-required`; docx stream → `not-presentable`; `presenterView:true` with `can.notes` false → `not-presentable`; with activation → host visible and `requestFullscreen` called **before** the lazy import resolves (gesture rule); Present button absent when `can.present` false.
- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run** + `bunx vitest run tests/stream tests/components` + `bun run gates:core` (record bytes; remote show must be in a lazy chunk) → PASS / recorded.
- [ ] **Step 5: Commit** `feat(stream): present() on streamed decks`

### Task 8: End-to-end, validation and docs

**Files:**
- Create: `tests/browser/stream-present.spec.ts`
- Create: `docs/validation/streaming-s3c.md`
- Modify: parent spec §12a.5/§12a.7/§14/§19 (notes caps, `notes` rate class, notes a11y out, calls C1–C4), design §4.3 amendment dated 2026-11-14, `server/README.md` (Presenting section)

- [ ] **Step 1: Browser spec** (`lo-tdf104788.pptx`): audience role presents, steps 6 slides with →, counter text, `goTo` via digits+Enter, Esc exits; presenter role opens swap Presenter View (single screen) and the notes canvas is painted; audience surface never contains a notes canvas; holding → for 10 s keeps the socket open (no 4008); `?notes=off` presenter → `presenterView:true` refused, plain show works.
- [ ] **Step 2: Run** `bun run test:browser:stream` → all pass.
- [ ] **Step 3: Validation doc**: §12a.8 3b checklist with the test proving each line, wire audit result, rate measurement from the hold-→ test, gates:core bytes, and "Known limits not machine-checked" (two-screen popup placement, real projector).
- [ ] **Step 4: Full gates:** `bun run test:server`, `bun run test`, `bun run lint`, `bun run typecheck:server`, `bun run gates:core` → record.
- [ ] **Step 5: Commit** `docs(stream): Stage 3 piece 3 validation and spec amendments`
