# Streaming Stage 3a — Search, Copy and A11y Text Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A streamed (remote) document can be searched by every viewer, and copied from / read by a screen reader exactly as far as the developer's `authorize` policy allows, with only visible text ever answered.

**Architecture:** The server builds a per-unit **served text index** inside the render worker: the existing text-recording pass (`buildTextIndex`) joined against the visual pass's occlusion runs, keeping only spans whose ink survives. All text operations (search, select, copy, a11y) run in the worker against that index and answer plain JSON over new WebSocket text messages; the main thread only enforces policy, caps, rates and audit. The client never builds an index: `StreamSession` gains a text-request class, `RemoteOfficeDoc` exposes `search/select/copy/a11y`, and `RemoteDocView` draws highlight boxes over the page canvases.

**Tech Stack:** TypeScript, Bun (`server/` workspace, `bun test`), bun:sqlite audit store, vitest (root), Playwright (`stream` project), React (lazy remote view).

**Spec:** `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.1 (piece 1 only). Parent: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` §11, §15, §7.7, §14.

**Worktree:** `.worktrees/stream-s3a`, branch `feature/streaming-s3a`, from `origin/main` once PR #71 is merged (else stack on `spec/streaming-s3`). Own `bun install`; copy `corpus/` and `corpus-s0/` from a sibling worktree if absent.

## Global Constraints

- Fail closed: a span the join cannot resolve is **withheld**, never approximated.
- Only visible text is answered: hidden content (in `visible-only` sessions), the visible mark's text and speaker notes never appear in any search, select, copy or a11y answer. Notes are **not indexed** in this piece.
- In a `contentPolicy: 'full'` session, hidden content is on the page and is answered like other visible text.
- The client **never** builds a local text index and never imports `src/core/search.ts` `buildTextIndex`/`findMatches` or `src/core/selection.ts` for a remote document.
- A copy over a cap is **refused with the cap stated**, never truncated.
- XLSX answers the painted display text, never values or formulas.
- Client request fields must not be named `policy`, `projection` or `contentPolicy` (`server/src/ws.ts:444` closes on them).
- Every server→client text message is ≤ 48 KiB of UTF-8 (under the 64 KiB text cap); larger answers are chunked.
- Every new client→server op follows the `onUnit` template (`server/src/ws.ts:301`): `phase === 'stream'`, `isU32(reqId) && reqId !== 0`, `verifyToken`, `owns()` re-check after every `await`, `table.acceptJti`, argument validation, rate check, fingerprint check.
- New client code loads only on the lazy remote path; `bun run gates:core` must not grow. It is **already RED on main** (index closure over its limit since PR #67): record the before/after numbers; do not edit the baseline or work around it. A gate that is newly red because of this work is reported to the owner.
- Never `commit -a`. Commits end with the trailers in the session's attribution note.
- Reviewer of each task is a different model family from its implementer.

## Plan rulings (decisions the spec leaves open — each is the plan's call)

- **R1 Join rule.** See Task 2. Spans with no matching visual run are withheld.
- **R2 Paragraph count.** A paragraph is one distinct `span.logical.source` object whose `source.text.trim() !== ''`. A span with **no** `logical` (every XLSX cell line, chart text) counts as **one paragraph per span**. XLSX wrapped cells therefore over-count (one per painted line) — the safe direction for a cap; the spec's "one cell" is met for unwrapped cells. A paragraph that crosses a page break counts once per page.
- **R3 Resume.** `copy` and `textA11y` are excluded from the resume policy-equality check; on resume the new values replace the old, the used totals are kept, and the resumed manifest carries the new caps. All other policy fields keep today's exact-match rule.
- **R4 `relaxed`.** Turning copy or a11y on does not set `EffectivePolicy.relaxed`; the audit row records the resolved services instead.
- **R5 Rates.** New rate class `'text'`: 60 ops/min per sid covering search, copy and a11y requests; `select` is charged only the existing message budget. A `'text'` overflow answers `refused: rate-limited` and does **not** count toward the breach window that closes the socket.
- **R6 Sizes.** Query ≤ 256 UTF-16 units; ≤ 5,000 matches per search (answer carries `truncated: true` beyond); one copy ≤ 4 MiB of UTF-8 (`copy-too-large` beyond, even under `'full'`); served-index cache ≤ 32 MiB per document per worker; one search in flight per sid (a new one supersedes).
- **R7 a11y modes.** `'unit'` (text per unit, uncounted) when copy is `'full'`, or `textA11y: true` with copy off; `'focused'` (one paragraph per request, counted against the copy caps) when copy is capped; `'off'` otherwise.

## Review Focus

1. **A selection that starts on one page and ends on another** — copies the text of both in reading order, counted once per paragraph per page. (Task 5 test.)
2. **Right-to-left and mixed-direction text** — copied text is in logical order, identical to local `textForRange` on the same document. (Task 5 test.)
3. **Worker recycle between select and copy** — the served index is rebuilt and the copy still answers. (Task 3 test.)
4. **A session resumed mid-search** — the in-flight search is re-issued or cleanly rejected; it never hangs. (Task 6 test.)
5. **Text drawn as glyph paths (warped WordArt), which has no visual run** — withheld, and reported in the yield table rather than silently served. (Task 9 measurement.)

## File structure

| File | Responsibility |
|---|---|
| `server/src/hooks.ts` (modify) | `copy` / `textA11y` policy input and resolution |
| `server/src/occlusion.ts` (modify) | export per-run ink + verdict (`runDetails()`) |
| `server/src/text/join.ts` (create) | pure join: spans × runs → served spans |
| `server/src/text/served.ts` (create) | per-unit served index build + paragraph ids; worker-side ops |
| `server/src/render.ts`, `server/src/worker/render-worker.ts`, `server/src/pool.ts` (modify) | `text` job plumbing |
| `server/src/text/ops.ts` (create) | main-thread op handlers: validation, caps, quota, audit, chunking |
| `server/src/ws.ts` (modify) | dispatch, manifest `can`/`caps`, resume rule |
| `server/src/rate.ts` (modify) | `'text'` class |
| `server/src/audit-store.ts`, `server/src/audit.ts` (modify) | `copies` table |
| `src/core/search.ts` (modify) | export `spansToLines` |
| `src/stream/manifest.ts` (modify) | `caps` in type + validator |
| `src/stream/remote/session.ts`, `errors.ts`, `remote-doc.ts` (modify) | text request class, typed refusals, public API |
| `src/components/remote/RemoteDocView.tsx` (modify), `src/components/remote/RemoteText.tsx` (create) | search bar, highlights, selection, copy, a11y layer |

---

### Task 1: Policy, manifest caps, resume rule

**Files:** Modify `server/src/hooks.ts`, `server/src/ws.ts:226-251,271-288`, `server/src/sessionTable.ts`, `src/stream/manifest.ts`. Test: `server/test/text-policy.test.ts` (create), `tests/stream/manifest.test.ts` (extend existing manifest test file).

**Interfaces — Produces:**
```ts
// hooks.ts — AuthDecision.policy gains:
copy?: false | 'full' | { maxParagraphs: number; maxSessionParagraphs?: number }
textA11y?: boolean
// EffectivePolicy gains:
copy: false | { maxParagraphs: number | null; maxSessionParagraphs: number | null }   // 'full' ⇒ both null
textA11y: 'off' | 'unit' | 'focused'                                                  // R7
// manifest.ts — Manifest gains:
caps: { copy: null | { maxParagraphs: number | null; maxSessionParagraphs: number | null }; a11y: 'off' | 'unit' | 'focused' }
```
`can.search` is `true` always; `can.copy = policy.copy !== false`; `can.textA11y = policy.textA11y !== 'off'`. A cap that is not a positive safe integer resolves to `copy: false` (fail closed).

- [ ] **Step 1: Failing tests** in `server/test/text-policy.test.ts`:
```ts
test('defaults: search on, copy off, a11y off', async () => { /* start({}) → hello() */
  expect(m.can).toMatchObject({ search: true, copy: false, textA11y: false }); expect(m.caps).toEqual({ copy: null, a11y: 'off' }) })
test("copy 'full' ⇒ a11y unit", ...)            // caps.copy = {maxParagraphs:null,maxSessionParagraphs:null}, a11y 'unit'
test('copy cap ⇒ a11y focused', ...)            // policy {copy:{maxParagraphs:3}} → {maxParagraphs:3,maxSessionParagraphs:null}, 'focused'
test('textA11y:true with copy off ⇒ unit', ...)
test('textA11y:false with copy full ⇒ off', ...)
test.each([0, -1, 1.5, NaN, '3'])('bad cap %p ⇒ copy off', ...)
test('resume with a changed copy cap succeeds and the new manifest carries it', ...)   // authorize returns {maxParagraphs:3} then {maxParagraphs:1}
test('resume with a changed visibleMark still closes auth-failed', ...)
```
and in the manifest test: `validateManifest` keeps `caps`, rejects a manifest whose `caps.copy.maxParagraphs` is negative or whose `caps.a11y` is unknown, and rejects a missing `caps`.
- [ ] **Step 2:** `cd server && bun test test/text-policy.test.ts` → FAIL (`caps` undefined).
- [ ] **Step 3:** Implement. Resume (R3): compare `effectivePolicy(d)` to the stored one with `copy`/`textA11y` removed from both; on success assign `sd.policy = effectivePolicy(d)` and update the session-table row's `effectivePolicy`.
- [ ] **Step 4:** `bun run test:server` and `bun run test -- tests/stream` → PASS. Update every manifest literal the validator now rejects (e.g. `server/test/client-remote-doc.test.ts:49`).
- [ ] **Step 5:** Commit `feat(stream): copy/a11y policy, manifest caps, resume keeps services mutable`.

### Task 2: Occlusion run details and the join

**Files:** Modify `server/src/occlusion.ts`; create `server/src/text/join.ts`; modify `src/core/search.ts` (export `spansToLines`). Test: `server/test/text-join.test.ts` (create), extend `server/test/occlusion.test.ts`.

**Interfaces — Produces:**
```ts
// occlusion.ts: the observer object gains
runDetails(): RunDetail[]
export interface RunDetail { text: string; ink: { x: number; y: number; width: number; height: number } /* device px */; servable: boolean; reason?: OcclusionReason }
// join.ts
export interface JoinResult { served: TextSpan[]; withheld: number }
export function joinSpans(spans: TextSpan[], runs: RunDetail[], scale: number): JoinResult
```
**R1, the join rule (normative for this plan).** Convert run ink to page px (÷ `scale`). A span's box is the bounds of `placement.visual.polygon` when present, else `{x, y: placement?.top ?? y - fontSize, width, height}` from the span. A run is a **candidate** for a span when the run's ink centre lies inside the span box inflated by 2 px, or the span box centre lies inside the run ink. The span is **served** iff (a) it has ≥ 1 candidate, (b) every candidate is `servable`, and (c) every non-whitespace code point of `span.text` is covered by the multiset of code points in the candidates' texts. Otherwise it is withheld. Whitespace-only spans are served iff a neighbouring span on the same line is.

- [ ] **Step 1: Failing tests** (`text-join.test.ts`; synthetic spans/runs):
```ts
test('one span, one servable run at the same box ⇒ served')
test('span whose run is covered ⇒ withheld')
test('span with no run ⇒ withheld')                              // glyph-path text
test('span over two runs, one withheld ⇒ withheld')
test('per-glyph runs covering every character ⇒ served')        // 'abc' vs runs 'a','b','c'
test('run text missing a character of the span ⇒ withheld')
test('scale 2 ink maps onto page-px spans')
```
and in `occlusion.test.ts`: `runDetails()` returns one entry per `fillText`, `servable:false, reason:'covered'` after an opaque `fillRect` over it.
- [ ] **Step 2:** `cd server && bun test test/text-join.test.ts` → FAIL (module missing).
- [ ] **Step 3:** Implement `runDetails()` (keep `evidence()` unchanged), `joinSpans`, and add `export` to `spansToLines` in `src/core/search.ts`.
- [ ] **Step 4:** `bun run test:server && bun run test -- tests/search` → PASS; `bun run gates:core` — record the index-closure bytes before and after in the task report.
- [ ] **Step 5:** Commit `feat(stream): occlusion run details and visible-text join`.

### Task 3: Served index in the worker; `text` jobs

**Files:** Create `server/src/text/served.ts`; modify `server/src/render.ts` (`DocumentHost`), `server/src/worker/render-worker.ts`, `server/src/pool.ts`. Test: `server/test/text-served.test.ts` (create).

**Interfaces — Consumes:** `joinSpans`, `runDetails()`, `buildTextIndex`, `spansToLines`, `hitTest`/`normalizeRange`/`rectsForSelectionOnPage`/`textForRange` (`src/core/selection.ts`), `findMatches`.
**Produces:**
```ts
export interface UnitPoint { unit: number; x: number; y: number }                 // page px, scale 1
export type TextRange = { all: true } | { anchor: UnitPoint; focus: UnitPoint }
export type TextJob =
  | { op: 'prepare'; unit: number }
  | { op: 'search'; query: string; caseSensitive: boolean; fromUnit: number; toUnit: number; maxMatches: number }
  | { op: 'select'; range: TextRange }
  | { op: 'copy'; range: TextRange; maxParagraphs: number | null; maxBytes: number }
  | { op: 'a11y'; unit: number; paragraph?: number }
export type TextReply =
  | { op: 'prepare'; served: number; withheld: number; bytes: number }
  | { op: 'search'; matches: Array<{ unit: number; rects: number[] /* x,y,w,h … */ }> }
  | { op: 'select'; units: Array<{ unit: number; rects: number[] }>; paragraphs: number }
  | { op: 'copy'; text: string; paragraphs: number } | { op: 'copy'; refused: 'too-large'; paragraphs: number; bytes: number }
  | { op: 'a11y'; paragraphs: Array<{ rect: number[]; text?: string }> }
  | { missing: number[] }                                                         // units not prepared in this worker
// DocumentHost
text(docKey: string, job: TextJob): Promise<TextReply>
// Pool
text(sessionId: string, job: { docRef: string; contentPolicy: 'visible-only' | 'full' } & TextJob): Promise<TextReply>
```
`prepare` builds one unit: `buildTextIndex([paintables[unit]])`, one visual pass at scale 1 through `occlusionObserver(...).wrap` to get `runDetails()`, `joinSpans`, then `spansToLines(served)`; it is cached per `(docKey, unit)` under the R6 32 MiB bound (evict least recently used). Every other op answers `{ missing }` when a unit it needs is not cached; `Pool.text` then runs `prepare` for each missing unit (one job each, existing 2 s deadline) and retries the op once. Paragraph counting follows R2. `a11y` without `paragraph` returns rect + text for every paragraph in reading order; with `paragraph: k` returns only the k-th with text and the rest rect-only. Notes are never passed to `buildTextIndex`.

- [ ] **Step 1: Failing tests** (real corpus documents through `pool()`):
```ts
test('search finds visible text and answers page-px rects inside the unit', ...)   // pypptx-test.pptx, a word known to be on slide 1
test('select between two points answers rects and a paragraph count ≥ 1', ...)
test('copy answers the same string as local textForRange for the same range', ...)  // build the local index with buildTextIndex in the test
test('copy over maxParagraphs is refused with the count, no text', ...)             // maxParagraphs: 1 on a 2-paragraph range
test('copy over maxBytes is refused', ...)
test('a11y paragraph:k carries text only for k', ...)
test('after a worker recycle, select then copy still answers', ...)                 // Review Focus 3: Pool with recyclePages: 1
test.each(hiddenIds)('hidden construct %s: marker never in any text reply; control is', ...)  // 13 fixtures, contentPolicy visible-only
test('contentPolicy full: the docx-vanish marker is answered', ...)
test('xlsx copy answers painted display text, not the formula', ...)
test('FIXED_MARK_TEXT is not searchable', ...)
```
- [ ] **Step 2:** `cd server && bun test test/text-served.test.ts` → FAIL.
- [ ] **Step 3:** Implement. Model `Pool.text` on `Pool.render`/`renderOn` (`ensureOpen`, doc affinity, admission, deadline).
- [ ] **Step 4:** `bun run test:server` → PASS.
- [ ] **Step 5:** Commit `feat(stream): served text index and text jobs in the render worker`.

### Task 4: Search over the wire

**Files:** Create `server/src/text/ops.ts`; modify `server/src/ws.ts` (dispatch, `RefusalCode`), `server/src/rate.ts`. Test: `server/test/text-search-wire.test.ts`; extend `server/test/harness.ts` `Client` with `text(msg)`.

**Interfaces — Produces (wire):**
```
C→S {type:'search', reqId, token, query, caseSensitive?}
S→C {reqId, type:'search-result', matches:[{unit, rects:[x,y,w,h,…]}]}      // 0..n messages, each ≤ 48 KiB
S→C {reqId, type:'search-done', total, truncated}
S→C {reqId, type:'refused', code, limit?, requested?}                      // refused gains optional numbers
RefusalCode += 'text-unavailable' | 'copy-too-large' | 'copy-quota-exhausted'
RateClass  += 'text'
```
The server walks units in order in batches (`fromUnit..toUnit`, 8 units per job), sending a `search-result` after each batch that has matches; a new `search` on the same sid supersedes the running one (the old reqId gets `search-done` with what was sent so far and `truncated: true`). Empty or over-long query (R6) ⇒ `refused: policy-refused`.

- [ ] **Step 1: Failing tests:**
```ts
test('search answers rects then search-done with the total', ...)
test('a second search supersedes the first', ...)
test('query of 257 chars is refused; 256 is accepted', ...)
test('61st text op in a minute is refused rate-limited and the socket stays open', ...)   // injected clock
test('search with a stale jti is refused session-rotated', ...)
test('no search-result message exceeds 48 KiB', ...)                                      // a query matching every line of a large docx
test('T2 over the wire: marker never in any search-result for the 13 constructs', ...)
```
- [ ] **Step 2:** run → FAIL (`policy-refused` for unknown type).
- [ ] **Step 3:** Implement per the `onUnit` template (Global Constraints).
- [ ] **Step 4:** `bun run test:server` → PASS.
- [ ] **Step 5:** Commit `feat(stream): search op`.

### Task 5: Select and copy over the wire; caps, quota, audit

**Files:** Modify `server/src/text/ops.ts`, `server/src/ws.ts`, `server/src/audit-store.ts`, `server/src/audit.ts`. Test: `server/test/text-copy-wire.test.ts`, extend `server/test/audit-store.test.ts`.

**Interfaces — Produces:**
```
C→S {type:'select', reqId, token, range}            // range = {all:true} | {anchor:{unit,x,y}, focus:{unit,x,y}}
S→C {reqId, type:'select-result', units:[{unit, rects}], paragraphs}        // chunked ≤ 48 KiB; last chunk carries last:true
C→S {type:'copy', reqId, token, range}
S→C {reqId, type:'copy-result', seq, text, last}                            // chunked on code-point boundaries, ≤ 48 KiB each
S→C {reqId, type:'refused', code:'copy-too-large', limit, requested} | {… 'copy-quota-exhausted', limit, requested} | {… 'text-unavailable'}
// audit-store.ts
export interface CopyRecord { sid: string; kind: 'copy' | 'a11y'; paragraphs: number; fromUnit: number; toUnit: number; at: number }
putCopy(r: CopyRecord): void;  copies(sid: string): CopyRecord[]
// audit.ts
recordCopy(r: Omit<CopyRecord, 'at'>): void          // throws for an unknown sid
// AuditRow gains  services: { copy: EffectivePolicy['copy']; textA11y: EffectivePolicy['textA11y'] };  copied: { events: number; paragraphs: number }
```
`copies` is its own table (`sid REFERENCES rows`, no UNIQUE key — repeat copies must not dedupe); `expire()` deletes `copies` before `rows`. Per-sid counter `copiedParagraphs` lives on the `sessionDoc` entry (survives resume, gone on teardown). Order for a copy: `can.copy` (else `text-unavailable`) → worker `copy` job with `maxParagraphs`, `maxBytes` 4 MiB → per-copy cap → session cap (`used + n > maxSessionParagraphs` ⇒ refuse with `limit` = cap, `requested` = n) → `audit.recordCopy` (failure ⇒ `refused: policy-refused`, nothing sent) → add to counter → send chunks. Copies on one sid are handled serially. `select` requires `can.copy`.

- [ ] **Step 1: Failing tests:**
```ts
test('copy off: select and copy are refused text-unavailable', ...)
test('maxParagraphs n: a selection of n is copied, n+1 is refused copy-too-large with limit n and requested n+1', ...)
test('maxSessionParagraphs: copies up to the total succeed, the next is refused copy-quota-exhausted', ...)
test('copying the same paragraph twice counts twice', ...)
test('the session total survives a resume and is measured against the resumed cap', ...)
test('a selection across two pages copies both in order and counts per page', ...)     // Review Focus 1
test('rtl/mixed text copies in logical order, equal to local textForRange', ...)         // Review Focus 2
test('a refused copy sends no copy-result and writes no audit row', ...)
test('each successful copy writes one audit copies row with its paragraph count', ...)
test('{all:true} on a full-copy session returns the whole document text chunked ≤ 48 KiB', ...)
test('expire() removes copies with their row', ...)
```
- [ ] **Step 2:** run → FAIL.
- [ ] **Step 3:** Implement.
- [ ] **Step 4:** `bun run test:server` → PASS.
- [ ] **Step 5:** Commit `feat(stream): select and copy ops with caps, session quota and audit`.

### Task 6: Client session and `RemoteOfficeDoc` API

**Files:** Modify `src/stream/remote/session.ts`, `errors.ts`, `remote-doc.ts`, `src/stream/index.ts`. Test: `server/test/client-text.test.ts` (create; real server), update `server/test/client-remote-doc.test.ts:74-79`.

**Interfaces — Produces:**
```ts
// errors.ts
StreamErrorCode += 'text-unavailable' | 'copy-too-large' | 'copy-quota-exhausted'
class StreamSessionError { readonly limit?: number; readonly requested?: number }
// session.ts (StreamSession and SessionLike)
text(request: TextRequest, on: (m: TextMessage) => void, signal?: AbortSignal): Promise<void>   // resolves on the op's terminal message
// remote-doc.ts — replaces the Promise<never> stubs
export interface RemoteSearchHit { unit: number; rects: HighlightRect[] }
search(query: string, o?: { caseSensitive?: boolean; signal?: AbortSignal; onHits?: (h: RemoteSearchHit[]) => void }): Promise<{ total: number; truncated: boolean }>
select(range: RemoteRange, signal?: AbortSignal): Promise<{ units: RemoteSearchHit[]; paragraphs: number }>
copy(range: RemoteRange): Promise<string>                      // rejects StreamSessionError with limit/requested on a cap
a11y(unit: number, paragraph?: number): Promise<Array<{ rect: HighlightRect; text?: string }>>
```
Text requests get their own pending map; on resume in-flight `search`/`select` are re-issued, in-flight `copy`/`a11y` are **rejected** with `session-rotated` (never re-issued — a re-issue would double-count quota); all are rejected on session end. When `can.copy`/`can.textA11y` is false the method rejects locally with `text-unavailable` and emits the `text-unavailable` notice once, as today.

- [ ] **Step 1: Failing tests:** search resolves with hits delivered through `onHits`; `copy` rejects with `code:'copy-too-large', limit, requested`; `copy` rejects `text-unavailable` once-notice when off; a resume during a search still resolves it (Review Focus 4); a resume during a copy rejects `session-rotated`; aborting a search stops `onHits`.
- [ ] **Step 2:** `cd server && bun test test/client-text.test.ts` → FAIL.
- [ ] **Step 3:** Implement.
- [ ] **Step 4:** `bun run test:server && bun run test && bun run lint` → PASS.
- [ ] **Step 5:** Commit `feat(stream): client text requests and RemoteOfficeDoc search/select/copy/a11y`.

### Task 7: Remote view — search bar, highlights, selection, copy

**Files:** Create `src/components/remote/RemoteText.tsx`; modify `src/components/remote/RemoteDocView.tsx`. Test: `tests/office-file-stream.test.tsx` (update), `tests/browser/stream-text.spec.ts` (create), update `tests/browser/stream-viewer.spec.ts:31-35`.

**Behaviour to build:**
- Each unit canvas gets a positioned wrapper `data-unit-wrap={i}` holding an overlay of percent-positioned boxes: `data-hit` (search), `data-hit-active`, `data-sel` (selection). No text is ever placed in the overlay.
- Ctrl/Cmd+F opens the existing `SearchBar` (`src/components/viewerControls.tsx:20`); typing debounces 200 ms then calls `document.search`; next/prev scroll the active hit into view; the count shows `total` (with "+" when `truncated`).
- With `can.copy`: pointer drag sends `select` (throttled to ≤ 5 per second, last position always sent) and draws `data-sel` boxes; Ctrl/Cmd+A selects `{all:true}`; Ctrl/Cmd+C calls `document.copy` then `copyText` (`src/core/clipboard.ts`).
- A refused copy shows, in the existing `role="status"` region, exactly: `Selection too large to copy (limit: N paragraphs).` or `Copy limit for this session reached (N paragraphs).` and copies nothing.
- `TEXT_UNAVAILABLE` is shown only when `can.copy` is false **and** `can.textA11y` is false, still once.
- Pointer coordinates map to page px by `x * unit.w / rect.width`.

- [ ] **Step 1: Failing tests.** vitest: notice absent when `can.copy` is true; present once when both are false. Playwright (`stream-text.spec.ts`, harness `server/test/app.ts` extended with `?copy=full|cap3|off`):
```ts
test('Ctrl+F then a query draws hit boxes and a count', ...)
test('drag-select then Ctrl+C puts the selected text on the clipboard (copy=full)', ...)
test('with cap3, selecting four paragraphs shows the limit message and the clipboard is unchanged', ...)
test('with copy off, dragging draws no selection', ...)
test('the page DOM contains no document text when a11y is off', ...)   // innerText has no word from the document
```
- [ ] **Step 2:** `bun run test:browser:stream` → FAIL.
- [ ] **Step 3:** Implement.
- [ ] **Step 4:** `bun run test && bun run test:browser:stream && bun run lint && bun run gates:core` → PASS (gates:core: no growth versus Task 2's recorded number).
- [ ] **Step 5:** Commit `feat(stream): search, selection and copy in the remote view`.

### Task 8: A11y text

**Files:** Modify `server/src/text/ops.ts`, `server/src/ws.ts`, `src/components/remote/RemoteText.tsx`. Test: `server/test/text-a11y-wire.test.ts`, extend `tests/browser/stream-text.spec.ts`.

**Interfaces — Produces (wire):**
```
C→S {type:'a11y', reqId, token, unit, paragraph?}
S→C {reqId, type:'a11y-result', seq, paragraphs:[{rect, text?}], last}     // chunked ≤ 48 KiB
```
Server rules: `textA11y 'off'` ⇒ `text-unavailable`. `'unit'` ⇒ `paragraph` must be absent; all text returned; audited as one `CopyRecord` `kind:'a11y'` per request, not counted against caps. `'focused'` ⇒ a request **without** `paragraph` returns rects only (no text, uncounted, unaudited); a request **with** `paragraph` returns that one paragraph's text, counts 1 against `maxSessionParagraphs` (refused `copy-quota-exhausted` when exhausted) and is audited `kind:'a11y'`.

Client: each unit wrapper gets a visually-hidden `role="group" aria-label="Page N"` layer. `'unit'`: filled with one `<p>` per paragraph when the unit first becomes visible. `'focused'`: one focusable `<p tabindex="0" aria-label="Paragraph k of n. Press Enter to read.">` per paragraph; Enter requests it once and caches the text for the session (re-focus does not re-request). `'off'`: no layer. The canvases stay `aria-hidden`.

- [ ] **Step 1: Failing tests:** unit mode returns every paragraph's text; focused mode without `paragraph` returns no `text` field anywhere; focused with `paragraph` counts 1 and exhausts with copy; off is refused; the 13 hidden constructs never appear; browser: in `'unit'` mode the accessibility tree of page 1 contains the control text; in `'focused'` mode it contains none until Enter.
- [ ] **Step 2:** run → FAIL.
- [ ] **Step 3:** Implement.
- [ ] **Step 4:** `bun run test:server && bun run test:browser:stream` → PASS.
- [ ] **Step 5:** Commit `feat(stream): a11y text service`.

### Task 9: Yield measurement, docs, spec amendments

**Files:** Create `server/scripts/text-yield.ts`, `docs/validation/streaming-s3a.md`; modify `server/README.md`, both spec files.

- [ ] **Step 1:** Write `server/scripts/text-yield.ts`: for every corpus document the server admits, run `prepare` on every unit and print per format: units, spans served, spans withheld, and withheld spans grouped by cause (no candidate run / candidate withheld by occlusion reason / character mismatch). Run it; expected: completes, non-zero served for every format.
- [ ] **Step 2:** Write `docs/validation/streaming-s3a.md` with that table, the `gates:core` before/after numbers, and the list of test files that cover each "done when" line of spec §4.1. A served rate below 90 % for any format, or any `no candidate run` class above 5 % of spans, is **reported to the owner in the final message** with the top causes (Review Focus 5) — not tuned away by loosening R1.
- [ ] **Step 3:** `server/README.md`: document `copy`, `textA11y`, the caps, the refusal codes, that session totals reset when the document is opened again, and that search is rate-limited only.
- [ ] **Step 4:** Spec amendments: design §4.1 gains R2 (XLSX wrapped cells count per painted line; paragraphs count per page) and R3; parent §14 gains the `'text'` rate class and the 48 KiB chunk rule; parent §6.2 resume rule notes the `copy`/`textA11y` exception.
- [ ] **Step 5:** Full gates: `bun run test && bun run test:server && bun run test:browser:stream && bun run lint && bun run gates:core`; commit `docs(stream): Stage 3a validation, README and spec amendments`.
