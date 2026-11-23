# Streaming Stage 4b: selection, resize and fit, find, copy and a11y text on streamed sheets (implementation plan)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** On a tiles-mode XLSX stream the viewer selects, navigates, resizes and auto-fits as local mode does, and offers find, copy and a11y text. Every text answer comes from the server's servable snapshot, and a covered, undrawn or merge-covered cell never answers. From 4b the default client asks for tiles.

**Architecture:**
- **Server.** One accessor, `sheetTextView`, turns a sheet snapshot plus the session's sizes into the only sheet any text answer may read. Covered, undrawn and merge-tail cells are blank in it, and so is every cell of an unknown-coverage sheet.
  - The pool worker runs the **local engines** over it: `WorkbookFind`, `serializeSelection`, `formatCellPresentation` and `autoFitSize`. Parity with local mode therefore holds by construction, and the tests check it against fixtures with the withheld cells physically deleted.
  - Sizes become per-session view state (`viewGen`, `viewStateHash`), so tiles, layout and the tracer stop assuming the identity view.
- **Client.** A new capability `sheet-view@1` selects tiles mode with every 4b behaviour. A 4a server ignores it and answers units mode, so Stage 3 search and copy keep working there. `StreamWorkbook` gains find, copy, a11y, fit and resize over the session. `InteractiveSheetDoc` keeps its own selection and keyboard code on streams: they are geometry-only.

**Tech Stack:** TypeScript. Server: Bun (`server/`, `bun test`, `bun:sqlite`, pool `Worker`s, the server canvas backend). Client: React (`src/components/xlsx`, `src/components/remote`), vitest, the Playwright `stream` project.

**Spec:** `docs/superpowers/specs/2026-11-22-streaming-s4-design.md` (binding; cited as "S4 §N"): §6.1 (the 4b rows), §6.2 steps 3–5, §6.3, §6.7, §6.8, §6.10, §6.11, §6.12, §9 4b, §10, and "As built (4a)", whose rulings bind here. Parent: `docs/superpowers/specs/2026-11-07-server-streaming-design.md` ("parent §N"): §6.2, §9.3, §11, §12, §14, §15, §19. Stage 3 text: `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.1 and its 2026-11-12 amendment ("S3 item N"). Format exemplar: `docs/superpowers/plans/2026-11-22-streaming-s4a.md` ("4a R*n*"). Base: main `1e2fbf4d` (4a merged).

## Open for owner

1. **Where merge-tail values are blanked.** S4 §10 item 1 says merge-covered values are purged "before layout, in `projectXlsx`". But `projectXlsx` also feeds units mode, whose formulas and CF would then read the purged cells. That breaks "units mode unchanged" and S4 §6.12 item 2's "merge-covered … cells paint as locally".
   - **Recommendation:** blank merge tails in `sheetTextView` only, for every text and key answer, and leave tiles and units untouched. A non-anchor merge cell paints no text, so nothing visible changes and no answer can carry it. Task 15 amends §10 item 1 to say so.
   - The plan is written to this recommendation (P3).
2. **When focused-mode a11y fetches.** In `focused` mode each fetched cell counts 1 against `maxSessionParagraphs` (S3 item 7). Fetching on every active-cell move lets arrow keys exhaust the quota. S3 avoided that with an explicit "Press Enter to read", but Enter moves the cursor in a grid.
   - **Recommendation:** fetch once the active cell has rested 400 ms, cache the text per `(sheetId, viewGen, cell)`, and never fetch a cached cell again. A cell counts once per view, and a blank or withheld cell counts 0.
   - The plan is written to this recommendation (P10).

## Global Constraints

- **Owner decisions (S4 §2, §11.1, as built):**
  - full local viewer; approach A; this PR is 4b;
  - search always on; copy and print set by the developer; `can.textA11y` follows copy;
  - merge-covered, undrawn and covered cells are hidden content;
  - covered stays covered;
  - a marked blank pane never names a wrong session.
- **4a as built binds:**
  - wire `sheetId` = the decimal index among the projected sheets in tab order;
  - `zoomStep` 0..10, never a float;
  - fixed `light`/`dark` palettes; the session clock `(day, timeZone)` (R9);
  - measured spill, and coverage blanking that also covers CF views and sparklines;
  - B2 (a pane's state is layout); the visible mark scales with zoom·dpr;
  - `viewStateHash = sha256(canonical({sort, filters, sizes, authorHiddenTier, clock}))`, never including the sid (parent §9.3 amendment);
  - the client pacing numbers (parent §14 amendment) stay unchanged.
- **Text answers apply 4a's coverage, always.** Every find hit, copy field, a11y text and fit measurement reads `sheetTextView` (Task 1) and nothing else:
  - base-covered cells (measured spill; the rotated-box adversary `100×20` at (0,0) rotated 90°, covering `x 40..60, y 30..50`);
  - current-view covered cells (S4 §6.2 step 5);
  - merge tails, and cells undrawn by number format or by CF `showValue="0"` (step 3);
  - `anchors.unresolved > 0`, an unbounded effect or an uncomputable transform: **every cell blank** for text answers, and `sheet-resize` is refused `coverage-unknown`;
  - under `'full'` the view is the identity (S4 §6.2 policy parameter).
- **Answers carry no hit text.** `sheet-find` sends positions only (S4 §6.7). Copy is display text as TSV, with one empty field per withheld cell. There is no local text index on a stream (parent §8.3): the client never constructs `WorkbookFind` or calls `copySelection` for a stream source.
- **Copy caps (S3 §4.1, verbatim):** `copy?: false | 'full' | { maxParagraphs: number; maxSessionParagraphs?: number }`.
  - "Caps refuse at n+1 and allow at n, per copy and per session." n is the policy's `maxParagraphs` (per copy) and `maxSessionParagraphs` (per session).
  - One cell counts 1, and an empty or withheld cell counts 0 (S3 item 1).
  - The refusals are `copy-too-large {limit, requested}` and `copy-quota-exhausted {limit, requested}`. Never truncated.
  - The counter is `TextSlot.copiedParagraphs` (`server/src/text/ops.ts:166`), shared with focused a11y. It survives a resume and resets when the document is opened again.
- **Limits:**
  - query ≤ 256 UTF-16 units; ≤ 5,000 hits per find;
  - one copy ≤ 4 MiB UTF-8; text answer messages ≤ 48 KiB (`TEXT_MESSAGE_BYTES`); every new request message ≤ 64 KiB;
  - a11y ≤ 2,000 cells per request (S4 §6.7).
- **Rates:**
  - `'text'` (60/min, refuse-only) covers `sheet-find`, `sheet-copy` and `sheet-a11y` (refused off-mode requests included).
  - New `'sheet'` (parent §14: 30 ops/min, 1 in flight, refuse-only, never toward the 4008 window) covers `sheet-fit` and `sheet-resize`. 4c completes the class.
- **Units mode and local mode unchanged:**
  - a session with neither `sheet-view@1` nor `sheet-tiles@1` gets Stage 3's manifest; one with `sheet-tiles@1` only gets 4a's, key for key;
  - `server/test/xlsx-units-baseline.test.ts` and the Stage 3 text and print suites pass unedited;
  - every `tests/**` local and viewer suite passes unmodified.
- **`tests/stream/**` edits are append-only.** New cases only; no existing assertion changes. `git diff --numstat 1e2fbf4d -- tests/stream` must show 0 deletions.
  - 4a's `remote-sheet-view.test.tsx` cases ("no Find", "no selection") stay true, because their fake manifests carry no `sheetView` (P1).
- **No runtime import from `src/xlsx/**` in `src/stream/**`.** `sheet-wire.ts` may take type-only imports, as it already does (`:7–8`). Selection encoding (`GridSelection` to the wire form) lives in `src/components/xlsx/streamText.ts`.
- **Bun 1.4.5 tests:**
  - no `expect(...).rejects`/`.resolves`; use `server/test/fixtures/rejection.ts` (`rejection`, `rejectionMessage`);
  - `await pool.dispose()` in every after-hook;
  - a worker sets `onmessage` before any top-level await;
  - new Workers start only through `startWorker` (`server/src/worker-start.ts:59`). 4b adds a message type, not a worker.
- **Heavy-test budget, decided here:**
  - `}, 30_000)` on every server test that builds more than one sheet snapshot or renders more than 4 tiles: Tasks 4, 6–8, 14 and the Task 1 battery;
  - `}, 120_000)` on the ink-oracle corpus test (Task 14).
- **Browser:**
  - stream specs run with their own stream port: `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npm run test:browser:stream`;
  - the unmodified viewer suites run with `OV_STREAM_PORT=4281`.
- **Size gate (T11):**
  - `bun run gates:core` is red at the base (index closure). Record it at the base in Task 3 and after every client task, and report the delta; never re-baseline.
  - New client code is reachable only from the lazy `InteractiveSheetDoc` and stream runtime chunks.
- **Determinism:** text answers and tiles read no clock. The session clock is injected, and every answer under one `(view state, clock)` is a pure function of it.
- **T8 p95** is not re-gated: 4b changes only the tile geometry's sizes. Task 15 re-runs `scripts/stream-s4/tile-p95.ts` once and records the figure.
- **Reviewer:** a different model family from the implementer, for each task.
- **Commits:** by path (`git commit -m "<msg>" -- <paths>`); never `-a`, never amend, never a bare `git stash`. Messages end with:
  ```
  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01B5aP5T5pG9oGcVPzuNpnmG
  ```
- **Commands:**
  - server: `cd server && bun test <file>`; client: `bunx vitest run <file>`;
  - typecheck: `npx tsc --noEmit -p .` and `npm run -s typecheck:server`; lint: `npm run -s lint`;
  - corpus: `bun scripts/fetch-corpus.ts`, run before any corpus step.

## Plan calls (spec silences and conflicts, decided here; Task 15 writes the wire-visible ones into the specs)

- **P1. A 4b capability.** The spec does not say what a 4b client does against a 4a server, which refuses every `sheet-*` text op.
  - Plan call: a new HELLO feature `sheet-view@1`.
    - A 4b server answers an XLSX in tiles mode iff the features hold `sheet-tiles@1` or `sheet-view@1`. The manifest carries `sheetView: 1` iff they hold `sheet-view@1`.
    - The default 4b client sends `['sheet-view@1']` only. A 4a server ignores it (4a R12) and answers **units** mode, with full Stage 3 search, copy and a11y.
    - A host that passes `streamSheets: 'tiles'` sends both features and gets tiles, with or without `sheetView`.
    - Without `sheetView`, the server refuses `sheet-find`, `sheet-copy`, `sheet-a11y`, `sheet-fit` and `sheet-resize` `unsupported-op`, and the client shows 4a's viewer.
  - **Why:** a default client in a mixed deployment never loses Stage 3 text (the reason for §6.10's 4a opt-in), and 4a's tests stay valid unedited.
- **P2. `sheetTextView` is the one text accessor** (S4 §6.2, §6.7).
  - Plan call: no text op reads `SheetSnapshot.sheet` directly. Under `unknown` coverage, `servableSheet` keeps the authored sheet (`server/src/sheet/snapshot.ts:83–90`), so a direct read would leak exactly that case.
  - **Why:** one proof point.
- **P3. Merge tails are blanked in answers only** (Open item 1). Tiles and `projectXlsx` are untouched.
- **P4. Undrawn cells.**
  - **Number format.** A cell is blank when its formatted display text is `''` while its value is non-empty; this is what `;;;` produces. The display text is formatted under the **CF-effective style**, so a CF rule that writes `;;;` counts too.
  - **The CF text scope.** It comes from `prepareCfTextScope` (`src/xlsx/cf/text.ts:375`, React-free; the local viewer reaches it through `InteractiveSheetDoc.tsx:1173`).
    - It is prepared once per text view, over each CF rule's sqref ∩ the used range, with the session clock: `dateContext {dateSystem, timeZone: clock.timeZone, epochNowMs: clockEpochMs(clock)}` and `now: () => clockEpochMs(clock)`, as the tile's `cfClock` is.
    - It is held on the `TextView` and is the `prepareScope`/`textOf` scope of every op (find, copy, a11y, fit).
    - When the sheet has CF rules and the scope is unavailable (refused, budget or throw), every cell of every CF sqref is blank (fail closed).
  - **CF.** A cell is blank when the snapshot's CF batch gives `barAt(r,c)?.showValue === false` or `iconAt(r,c)?.showValue === false`. When the sheet has a `showValue="0"` rule and its CF batch is unavailable, every cell of that rule's sqref is blank.
  - Plan call: `hidesTextAt` is not used, because it depends on host and scale (`src/xlsx/cf/render.ts:39–41`).
  - **Why:** analytic, scale-free and fail-closed.
- **P5. Current-view coverage** is `baseCoverage` over the projected sheet, with its drawings placed by session sizes (`sizedDrawings`, `src/xlsx/view/viewport.ts:470`) and `buildSheetGeometry(projected, {sizes})`. Base coverage stays on authored geometry (4a). The union of the two is blank in text answers; tiles stay base-blanked only.
- **P6. Find on the wire.**
  - **Request:** `sheet-find {query, matchCase?, scope?}`. `scope` is `'workbook'` (the default, as local find is workbook-wide) or a `sheetId`.
  - **Answer:** `sheet-find-result {hits: [[sheetIndex, row, col], …]}` chunks, then `sheet-find-done {total, truncated}`, ordered as `WorkbookFind` orders them.
  - Plan call: no `wholeCell` (local find has none) and no `cursor` (one answer holds up to 5,000 hits; past that, `truncated: true`). A `hidden: 'filter'` hit is the 4-tuple `[s, r, c, 1]`. Local find matches visible cells only (`src/xlsx/view/find.ts:563`), so 4b never emits it; the codec and the viewer message are built and tested with a fake server.
  - **Why:** parity with the local engine, and spec gaps reported rather than invented.
- **P7. The selection on the wire** is interval-only: `{active: [r, c] | null, anchor: [r, c] | null, ranges: [{rows: [[r0, r1], …], cols: [[c0, c1], …]}]}`.
  - Caps: ≤ 256 ranges; intervals ascending, disjoint, inclusive and in range; message ≤ 64 KiB.
  - The client normalises each part's `CompactSourceRowSet` with `rowSetIntervals` and refuses locally with 'Selection too large to copy.' before sending. A worker `ref` part cannot occur on a stream; it refuses the same way. The server rebuilds the rows as `{kind: 'intervals'}` and refuses any other shape `policy-refused`.
  - **Why:** `CompactSourceRowSet` has typed-array, bitset and `ref` variants (`src/xlsx/view/contracts.ts:162`).
- **P8. Copy count and audit.**
  - n = the number of non-empty fields in the TSV.
  - Order: resolve → serialize (4 MiB cap) → count → per-copy cap → session cap → audit → counter → chunks, as `copyTurn` (`server/src/text/ops.ts:319`) does, on the sid's `copyQueue`.
  - `copies` gains nullable `sheet_id TEXT` and `rect TEXT` (JSON of the selection's bounding source rect), added with a guarded `ALTER TABLE`. `from_unit` and `to_unit` hold the sheet index.
- **P9. Fit carries the pane.** Local `autoFitSize` measures visible rows first (`src/components/xlsx/sheetResize.tsx:225–255`), and S4 §6.7's request has no pane.
  - Plan call: `sheet-fit {sheetId, viewGen, axis, index, pane: {x, y, width, height}, regions?: [≤ 4 boxes]}`, finite and non-negative, clamped to the sheet extent. `knownRows` are the text view's first rows, as on the local resident path. The answer is `sheet-fit-result {size}`.
  - `autoFitSize`, `AutoFitInput` and `createFitMeasure` move to a React-free `src/xlsx/view/autoFit.ts`, and `sheetResize.tsx` re-exports them.
  - **Why:** fit-size parity is defined only with the pane.
- **P10. A11y.**
  - **Request:** `sheet-a11y {sheetId, viewGen, rows: [[r0, r1], …], cols: [c0, c1]}`, in source rows (so 4c's sorted windows fit), ≤ 2,000 cells.
  - **Answer:** `sheet-a11y-result {seq, cells: [[row, col, text], …], last}`, listing only cells with servable non-empty text.
  - **Text:** `formatCellPresentation(textView, cell, budget).text` (`src/components/xlsx/formatPresentation.ts:44`). The CF descriptor is omitted (S4 §6.7's default).
  - **Modes:** `unit` serves the window, uncounted, with one audit row per request; `focused` requires a 1×1 window and counts 1 when the text is non-empty; `off` is refused `text-unavailable`, and the client never sends it.
  - **Client:** unit mode fetches a resting window (250 ms) and only its uncached cells, at most 30 requests per minute; focused mode follows Open item 2.
- **P11. Resize on the wire.**
  - `sheet-resize {sheetId, viewGen, axis: 'row' | 'col', index, size}`, or `{sheetId, viewGen, reset: true}`, which resets the whole sheet, as local "Reset sizes" does. Rows are keyed by source row.
  - The size is clamped by `clampSize` (`RESIZE_LIMITS`, `src/xlsx/view/axisSizes.ts:28`).
  - The answer is `sheet-view {sheetId, viewGen, change: {axis, index, size} | {reset: true}}`, carrying the applied size.
  - `sheet-layout` gains `sessionRows` and `sessionCols` (`[index, px]` lists, chunked like `rows`), so the full sizes reach the client once.
  - **Why:** a full list in every reply would pass 64 KiB at 4,096 entries per axis.
- **P12. Refusals.**
  - **Order:** `stale-view` for an old `viewGen`; `hidden-target` for a collapsed row or column (local: "is hidden; it cannot be resized", `InteractiveSheetDoc.tsx:3123`); `coverage-unknown` under `visible-only` on an unknown-coverage sheet; `view-state-limit`. None advances `viewGen`.
  - **The `view-state-limit` caps:** at most `MAX_SESSION_SIZES` = 4,096 per axis per sheet (the local cap), and at most 16 MiB of view state per process, at `SESSION_SIZE_ENTRY_BYTES` = 128 B per entry.
  - Plan call: S4 §6.3's "two non-identity sheets per sid" is sized for 8 MiB sort bands, so it counts sort and filter state only (4c). Sizes are capped per axis instead, so a viewer may resize on any number of tabs.
  - Charging the document-cache admission (§6.3 says "required") waits for 4c, whose bands need it. 4b's view state is at most 1 MiB per sheet.
- **P12a. Client pacing for `'sheet'` and `'text'`.**
  - One `sheet-resize` is in flight per viewer. Key-repeat steps made meanwhile fold into one pending target size (the latest wins), which is sent when the reply lands.
  - A fit is two ops (`sheet-fit`, then `sheet-resize`).
  - The find query is sent 250 ms after the last keystroke, and a newer one supersedes it.
  - **Announcements:**
    - `rate-limited`: 'Too many requests. Try again shortly.' (S3 item 10);
    - `view-state-limit`: the local "{name} was not resized: this sheet already has the maximum number of resized {columns|rows}. Reset sizes to start over.";
    - `coverage-unknown`: "{name} can't be resized in this viewing mode.";
    - `hidden-target`: the local "{name} is hidden; it cannot be resized.".
- **P13. Lost replies.**
  - On resume, `sheet-find` and `sheet-fit` are re-issued; `sheet-copy`, `sheet-a11y` and `sheet-resize` reject `session-rotated` and are never re-issued (S3 item 11).
  - After a rejected resize the client re-fetches `sheet-layout` for the authoritative `viewGen` and sizes.
- **P14. A resize stops the sid's tile batches of that sheet** (`batch.stopped = true`). Their late tiles carry the old `viewGen` and the client drops them, as 4a already does.
- **P15. The notices** (S4 §6.11, exact): "Hidden cells can't be shown in this viewing mode.", "Some rows are hidden by the author and can't be shown." (once per sheet, when `layout.authorHidden`), "This match is in a row hidden by the filter.", and 'text unavailable in this viewing mode' once when copy or a11y is off (S4 §6.11; this widens S3 item 10's "both").
- **P16. The ink oracle** (S4 §6.7) is pixel-based and does not reuse coverage.
  - A cell's ink is the set of pixels that differ between two renders of the projected sheet with its drawings removed: one as is, and one with that cell blank.
  - The cell would be withheld if the authored render (drawings on) differs from the drawings-removed render at any of those pixels.
  - Each such cell must already be blank in `sheetTextView`. A counterexample is a bug, fixed by widening the analytic rule (fail closed).
- **P17. No user freeze command.** The viewer has authored frozen panes only; there is no freeze toggle. "Client freeze choice" therefore means selection and navigation parity under authored panes. This is a spec gap, reported.

## Review Focus

1. **A resize lost to a resume, or crossing tile batches in flight.** Expected: the client re-fetches `sheet-layout`, its `viewGen` and sizes equal the server's, no old-`viewGen` tile is drawn, and no tile stays pending. Pinned in Tasks 4 and 9.
2. **`copy` or `textA11y` changing on resume** (S3 item 2). Expected:
   - the used total is kept and measured against the new caps;
   - a11y turned off clears the grid text and fetches nothing more;
   - copy turned off disables Copy and Ctrl+C.

   Pinned in Tasks 7 and 11.
3. **A covered cell far off the viewport, never tiled.** Expected: find, copy, a11y and fit treat it as blank from any scroll history, including after a resize that uncovers its neighbour. Pinned in Tasks 1 and 14.
4. **A select-all copy on a million-row sheet.** Expected: the local answer in one copy (the fixture is about 8k cells), never truncated, and no 4008 close. Over 4 MiB, the answer is `copy-too-large {limit: 4194304}`. Pinned in Task 7.
5. **Mixed versions.**
   - A default 4b client against a 4a server gets units mode with Stage 3 text.
   - An opted-in 4b client against a 4a server gets tiles with 4a behaviour.
   - A 4a client against a 4b server gets an unchanged manifest.

   Pinned in Tasks 3 and 13.

**Not in 4b** (4c): sort, filter, `filter-values`, `range-detect`, C1, step 6's key use, typed-condition refusal, `pivot-refused`, `sheet_ops`, and the rest of `'sheet'`.

---

### Task 1: `sheetTextView`: the one sheet text answers read

**Files:**
- Create: `server/src/sheet/text-view.ts`
- Modify: `server/test/sheet-fixtures.ts` (append fixtures)
- Test: create `server/test/sheet-text-view.test.ts`

**Interfaces:**
- Consumes:
  - `SheetSnapshot`, `servableSheet` and `blankCovered` (`server/src/sheet/snapshot.ts`);
  - `baseCoverage` (`server/src/sheet/coverage.ts:208`);
  - `buildSheetGeometry` and `sizedDrawings` (`src/xlsx/view/viewport.ts:470`);
  - `SessionSizes` (`src/xlsx/view/axisSizes.ts`).
- Produces:
  ```ts
  export interface TextView { sheetIndex: number; sheet: XlsxSheet; geometry: SheetGeometryResult; blankAll: boolean; withheld: number; scope: EffectiveFormatScope | undefined; release(): void }
  export async function sheetTextView(snap: SheetSnapshot, projected: XlsxSheet, sizes: SessionSizes, o: { workbook: XlsxDocument; clock: SheetClock; budget: FormatBudget }): Promise<TextView>
  export function textDocument(base: XlsxDocument, views: ReadonlyMap<number, TextView>): XlsxDocument   // the sheets replaced by their views
  ```
  - Under `'full'`, `sheet === projected` and `withheld === 0`.
  - Under `visible-only` with unknown coverage, `blankAll` is true and every cell is a value-less copy that keeps its style.
  - Otherwise the view is `snap.sheet` with P3, P4 and P5 blanking applied, `geometry` is built with `sizes`, and `scope` is P4's CF text scope.
  - A view needs the full `SheetSnapshot`, because P4 reads its CF batch (`barAt`/`iconAt`). The host builds views from cached snapshots (Task 5).
- Fixtures appended to `sheet-fixtures.ts`:
  - `offViewportCovered()`: `OFFVIEW-Z5000` at Z5000 under an unfilled rect, and an uncovered `OFFVIEW-A1`;
  - `slidUnderAnchor()`: a `oneCellAnchor` 150×20 px shape at B1 and `SLID-E1` in E1. Shrinking B to 8 px covers E1;
  - `mergeTail()`: A1:C1 merged, with `MERGE-A1` and `TAIL-B1`;
  - `hiddenFormat()`: B2 `HIDDENFMT` formatted `;;;`;
  - `cfHiddenValue()`: a data bar with `showValue="0"` on B1:B4, numbers in B1:B3 and `CFTEXT-B4` in B4;
  - `cfHiddenFormat()`: B1:B3 hold 4, 7 and 9, with a CF `cellIs > 5` rule whose dxf writes `numFmt ";;;"`. B2 and B3 paint nothing; B1 paints `4`;
  - `absoluteOverB2()`: an `absoluteAnchor` unfilled rect over B2, which holds `ABS-B2`. Widening column A to 300 px slides B2 out from under it;
  - `textGrid()`: 2 sheets of 30×8 mixed text, numbers, `#,##0.0` and dates. It has `needle`/`NEEDLE` in 6 cells, hidden row 7, hidden column F, merge A5:C6, frozen 1×1, and an image over D20 that covers `NEEDLE-D20`.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)` on each).
  - `"engines import under Bun"`: `createObjectWorkbookSource` (`src/xlsx/view/source.ts:975`), `WorkbookFind` (`find.ts:248`), `serializeSelection` (`copy.ts:176`), `prepareCfTextScope` (`cf/text.ts:375`) and `formatCellPresentation` all run once on `textGrid` under `bun test`. If one cannot, STOP and report.
  - `"covered cells are blank in the view"`:
    - `unfilledShapeOverText`: B2 has no value;
    - `rotatedBox`: every cell whose grid rect meets `x 40..60, y 30..50` has no value;
    - `offViewportCovered`: Z5000 has none and A1 keeps `OFFVIEW-A1`.
  - `"unknown coverage blanks every cell"`: on `unresolvedAnchor`, `blankAll === true`, no cell holds a value, and styles are kept. The snapshot's own sheet still holds values (tiles paint as authored).
  - `"current-view coverage"`: on `slidUnderAnchor`, E1 holds `SLID-E1` with no sizes, and has no value with `cols {1: 8}`. The snapshot's tile sheet still holds it.
  - `"merge tails and undrawn cells are blank"`:
    - `TAIL-B1` is gone and `MERGE-A1` stays;
    - `HIDDENFMT` is gone;
    - B1:B3 of `cfHiddenValue` are blank and `CFTEXT-B4` stays (P4 precision);
    - with the CF batch stubbed unavailable, B1:B4 are all blank;
    - `cfHiddenFormat`: B2 and B3 are blank and B1 keeps `4`; with the CF scope stubbed unavailable, B1:B3 are all blank;
    - the scope's clock is the session's: `prepareCfTextScope` (spied) receives `epochNowMs === clockEpochMs(clock)`.
  - `"covered stays covered in answers"`: on `absoluteOverB2` with `cols {0: 300}`, B2 is still blank, both in the view and in the snapshot's tile sheet.
  - `"full is the identity"`: under `'full'`, `sheet === projected` for each fixture above.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-text-view.test.ts`. Expected: FAIL (no module).
- [ ] **Step 3: Implement** `text-view.ts`. P3, P4 and P5 blank by `blankCell` semantics: value removed, style kept. Whether the CF batch is unavailable is read from the snapshot's assets (`RenderSheetAssets` CF batch absent or `undefined`).
- [ ] **Step 4: Run green:** the same file, then `cd server && bun test test/sheet-snapshot.test.ts test/sheet-t2.test.ts`, both unchanged.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): sheetTextView — the one sheet every text answer reads" -- server/src/sheet/text-view.ts server/test/sheet-fixtures.ts server/test/sheet-text-view.test.ts
  ```

### Task 2: Wire codecs for 4b (React-free)

**Files:**
- Modify: `src/stream/sheet-wire.ts`
- Test: modify `tests/stream/sheet-wire.test.ts` (append only)

**Interfaces:**
- Produces, in `sheet-wire.ts`:
  - `SHEET_VIEW_FEATURE = 'sheet-view@1'`;
  - `SELECTION_LIMITS = { ranges: 256 }`, `A11Y_MAX_CELLS = 2_000`, `FIND_MAX_HITS = 5_000` and `QUERY_MAX_UNITS = 256`;
  - `WireSelection` (P7), `encodeWireSelection(s: WireSelection): object` and `decodeWireSelection(v: unknown): WireSelection | undefined` (strict);
  - request validators returning the typed args or `undefined` (⇒ `policy-refused`):
    - `findArgs(m, sheetCount)`;
    - `copyArgs(m, sheetCount)`;
    - `a11yArgs(m, sheetCount)`, which checks ≤ 2,000 cells and `cols` within `[0, SHEET_MAX_COLS)`;
    - `fitArgs(m, sheetCount)` (P9);
    - `resizeArgs(m, sheetCount)` (P11).
  - answer decoders (client):
    - `decodeFindHits(m): Array<{ sheetIndex: number; row: number; col: number; hidden?: 'filter' }>`;
    - `decodeA11yCells(m): Array<{ row: number; col: number; text: string }>`;
    - `decodeFitResult(m): number`;
    - `decodeSheetView(m): SheetViewChange`.
  - the shared types:
    ```ts
    export interface SheetFindHit { sheetIndex: number; row: number; col: number; hidden?: 'filter' }
    export interface SheetFindAnswer { hits: SheetFindHit[]; total: number; truncated: boolean }
    export interface Box { x: number; y: number; width: number; height: number }
    export interface SheetFitRequest { sheetId: string; viewGen: number; axis: 'row' | 'col'; index: number; pane: Box; regions?: Box[] }
    export type SheetResizeRequest = { sheetId: string; viewGen: number } & ({ axis: 'row' | 'col'; index: number; size: number } | { reset: true })
    export interface SheetViewChange { sheetId: string; viewGen: number; change: { axis: 'row' | 'col'; index: number; size: number } | { reset: true } }
    ```
  - `StreamSheetLayout` gains `sessionRows: Array<[number, number]>` and `sessionCols: Array<[number, number]>`, chunked under `LAYOUT_LIMITS.intervals`. A 4a server's layout has neither, and decoding gives `[]`.

- [ ] **Step 1: Write the failing tests** (append):
  - `"selection round-trips; hostile selections are refused"`: a 3-range selection round-trips. `decodeWireSelection` returns `undefined` for each of:
    - 257 ranges;
    - an interval with `r1 < r0`;
    - overlapping intervals;
    - row 1,048,576;
    - col 16,384;
    - a `kind` key;
    - a non-integer.
  - `"find args"`: an empty query, 257 UTF-16 units, `scope: '9'` on a 2-sheet book and `wholeCell: true` are each refused. `{query: 'x'}` gives scope `'workbook'` and `matchCase: false`.
  - `"a11y args"`: 2,001 cells refused, 2,000 accepted, and `cols [5, 2]` refused.
  - `"fit and resize args"`:
    - a NaN pane, a negative width and `axis: 'z'` are refused;
    - `size: 1e9` is accepted and clamped later by the server;
    - `{reset: true, axis: 'row'}` is refused (mixed forms).
  - `"hits decode"`: `[[0, 4, 2], [1, 0, 0, 1]]` gives the second with `hidden: 'filter'`; `[[0, -1, 0]]` throws `InvalidOperand`.
  - `"layout session sizes"`: a layout with 5,000 `sessionCols` chunks and round-trips. A 4a-shaped layout decodes with `sessionRows` deep-equal to `[]`.
- [ ] **Step 2: Run** `bunx vitest run tests/stream/sheet-wire.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.** Only type imports from `src/xlsx/**`.
- [ ] **Step 4: Run green:** the same file, and `npx tsc --noEmit -p .`.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): sheet-view wire — selection, find, copy, a11y, fit, resize codecs" -- src/stream/sheet-wire.ts tests/stream/sheet-wire.test.ts
  ```

### Task 3: Handshake `sheet-view@1`, op gating, the `'sheet'` class, refusal codes

**Files:**
- Modify:
  - `server/src/ws.ts`: `TILES_REFUSED` at `:88`, the manifest at `:492`, the dispatch at `:1469–1476`, `RefusalCode` at `:60`;
  - `server/src/rate.ts`: `RateClass` at `:41`, `SessionRates`;
  - `server/src/sheet/mode.ts:24`: `sheetModeOf`;
  - `src/stream/manifest.ts`: `sheetView?: 1`.
- Test: create `server/test/sheet-view-handshake.test.ts`; modify `tests/stream/manifest.test.ts` (append)

**Interfaces:**
- Produces:
  - `sheetModeOf` (`server/src/sheet/mode.ts:24`) gives `'tiles'` for an XLSX when `features` include `sheet-tiles@1` or `sheet-view@1` (P1).
  - `sessionDoc` entries gain `sheetView: boolean`, which is true iff tiles mode and `features` include `sheet-view@1`. It is fixed per sid, and resume already compares `features`.
  - The manifest gains `sheetView: 1` only when that holds.
  - `SessionRates.check('sheet', now)` at 30 per 60 s, plus `SessionRates.sheetBusy: boolean`, the one-in-flight flag, set and cleared by the op.
  - `SHEET_RATE_PER_MIN = 30`.
  - `RefusalCode` gains `'hidden-target' | 'coverage-unknown' | 'view-state-limit'`.
  - `SHEET_VIEW_OPS = new Set(['sheet-find', 'sheet-copy', 'sheet-a11y', 'sheet-fit', 'sheet-resize'])`. Each is refused `unsupported-op` unless `sheetView` (units mode included).
  - Client: `Manifest.sheetView?: 1`, valid only with `sheetMode: 'tiles'`; any other value rejects the manifest.

- [ ] **Step 1: Write the failing server tests:**
  - `"sheet-view@1 adds sheetView to a tiles manifest only"`:
    - `['sheet-tiles@1', 'sheet-view@1']` on XLSX gives `sheetView === 1`;
    - `['sheet-tiles@1']` gives the 4a tiles manifest with `'sheetView' in m === false`;
    - `['sheet-view@1']` alone gives tiles mode with `sheetView === 1` (P1);
    - a DOCX gets no new keys.
  - `"the 4b ops need the feature"`: with `['sheet-tiles@1']`, each of the five ops is refused `unsupported-op`, and so is each in units mode.
  - `"resume keeps the feature"`: a `sheet-view` sid resumed without `sheet-view@1` closes `CLOSE['auth-failed']`.
  - `"the sheet class"`: with an injected `now`, 30 `check('sheet')` calls are `'ok'` and the 31st is `'refuse'`. Ten refusals never close 4008.
- [ ] **Step 2: Write the failing client tests** (append to `manifest.test.ts`): `sheetView: 1` with tiles passes; `sheetView: 2` rejects; `sheetView: 1` without `sheetMode` rejects; absent gives `undefined`.
- [ ] **Step 3: Run** `cd server && bun test test/sheet-view-handshake.test.ts` and `bunx vitest run tests/stream/manifest.test.ts`. Expected: FAIL.
- [ ] **Step 4: Implement.** Record `bun run gates:core` at the base before the client edit, and again after.
- [ ] **Step 5: Run green:** the same commands, then `cd server && bun test test/sheet-handshake.test.ts test/xlsx-units-baseline.test.ts` (unchanged).
- [ ] **Step 6: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): sheet-view@1 capability, sheet rate class, 4b refusal codes" -- server/src/ws.ts server/src/rate.ts server/src/sheet/mode.ts src/stream/manifest.ts server/test/sheet-view-handshake.test.ts tests/stream/manifest.test.ts
  ```

### Task 4: Session view state, `sheet-resize`, and sized tiles

**Files:**
- Create: `server/src/sheet/view-state.ts`
- Modify:
  - `server/src/cache.ts:51–56`: `SheetViewState.sizes`;
  - `server/src/sheet/tiles-op.ts:42`: `tilesArgs` takes the current viewGen;
  - `server/src/sheet/tile.ts:13`, `:53`: `TileJob.sizes`, with no `viewGen !== 0` check;
  - `server/src/render.ts:329`, `:369`, `:392`: snapshot geometry per sizes, `sheetLayout` and `renderTile`;
  - `server/src/sheet/layout.ts`: `sessionRows`/`sessionCols` and the current `viewGen`;
  - `server/src/ws.ts`: `onSheetLayout`, `tilesOp`, `extentKey`, `tileKeyFields` at `:836`, the `sheet_views` row and the new `onSheetResize`;
  - `server/src/audit-store.ts:31`: `SheetViewRecord.spec.sizes`;
  - `server/src/detect/trace.ts:206`: `traceTiles` rebuilds each tile with its delivery's `viewGen` spec;
  - `src/xlsx/view/streamGeometry.ts:17`: `streamSheetGeometry(l)` merges `l.sessionRows`/`l.sessionCols` with `mergeAxisOverrides`, exactly as `buildSheetGeometry` merges session sizes. The layout is the one source of sizes, and the server's `extentOf` (ws.ts) reads it too.
- Test: create `server/test/sheet-resize.test.ts`; modify `server/test/trace-tiles.test.ts` (append)

**Interfaces:**
- Produces, in `view-state.ts`:
  ```ts
  export type WireSizes = Array<{ axis: 'row' | 'col'; index: number; px: number }>   // canonical: axis 'col' < 'row', index ascending
  export class SessionViews {
    constructor(ledger: ViewStateLedger)
    viewGen(sheetIndex: number): number
    sizes(sheetIndex: number): SessionSizes
    canonical(sheetIndex: number): WireSizes
    apply(sheetIndex: number, change: { axis: 'row' | 'col'; index: number; size: number } | { reset: true }): { viewGen: number; applied?: number } | 'view-state-limit'
    dispose(): void                                   // returns its bytes to the ledger
  }
  export class ViewStateLedger { constructor(readonly maxBytes = 16 * 1024 * 1024) }
  export function sessionSizesOf(w: WireSizes): SessionSizes
  export function wireSizesOf(s: SessionSizes): WireSizes          // canonical order
  ```
  - `sessionDoc` entries hold `views: SessionViews`. It survives resume and is disposed at teardown.
  - `viewStateHash` takes `sizes: WireSizes`, so the identity view hashes exactly as in 4a (a golden pins it).
- `TileJob.sizes: WireSizes`. `DocumentHost` caches the sized geometry per `(snapshot key, sha256(sizes))`, keeping at most 8 per document (LRU) and releasing on `drop`. `renderTile` paints `snap.sheet` with that geometry and with the drawings it places.
- `onSheetResize`, in order:
  - the token;
  - `resizeArgs`;
  - `'sheet'` rate and busy;
  - `stale-view`;
  - `hidden-target`, when the authored size of the row or column is 0;
  - `coverage-unknown`, under `visible-only` with `coverage.kind === 'unknown'`;
  - `views.apply`, which can refuse `view-state-limit`;
  - stopping the sid's batches of that sheet (P14);
  - `sheet-view`.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)`):
  - `"resize applies, clamps and advances viewGen"`: on `grid500x100`, a col 2 resize to 2,000 gives `sheet-view {viewGen: 1, change: {axis: 'col', index: 2, size: 1790}}`. The next `sheet-layout` has `viewGen 1` and `sessionCols` deep-equal to `[[2, 1790]]`, and `reset: true` gives `viewGen 2` and `sessionCols []`.
  - `"refusals never move viewGen"`: each of these is refused, and a following layout still says `viewGen 0`:
    - `viewGen 5` (`stale-view`);
    - a hidden row of `xlsx-hidden-rows` under `'full'` (`hidden-target`);
    - `unresolvedAnchor` under visible-only (`coverage-unknown`);
    - 4,097 distinct column indices (`view-state-limit` on the last);
    - with `new ViewStateLedger(1024)` injected, the 9th entry in the process (`view-state-limit`).
  - `"three tabs may be resized"`: resizes on sheets 0, 1 and 2 of a 3-sheet book all succeed (P12).
  - `"client geometry equals the server's under sizes"`: for `grid500x100` with `{cols {1: 8, 3: 200}, rows {4: 60}}`, `streamSheetGeometry(layout).metrics.cols.start(j)` and `rows.start(i)` equal `buildSheetGeometry(sheet, {sizes})`'s for `j` in [0, 1, 2, 3, 4, 16383] and `i` in [0, 4, 5, 499, 1048575].
  - `"the sheet class refuses at 31"`: 30 resizes in a minute pass, the 31st is refused `rate-limited`, and a second resize sent while one is in flight is refused `rate-limited`.
  - `"sized tiles"`:
    - after a resize, `tiles` at `viewGen 0` gives `stale-view`;
    - at `viewGen 1`, tile (0,0) equals `renderSheetViewport` of the snapshot sheet over `buildSheetGeometry(sheet, {sizes})`;
    - the sheet extent grows, so a column past the old grid is admitted.
  - `"same sizes share canonical tiles; different sizes never do"`: two sids with equal sizes give one `pool.tile` call and different marked bytes; a third sid with another size gives a second call.
  - `"sheet_views records the sizes"`: `sheetView(sid, '0', 1).spec.sizes` deep-equals `[{axis: 'col', index: 2, px: 1790}]`.
  - `"a resize stops in-flight batches"`: a 64-tile batch, then a resize, gives that batch's `tiles-done.cancelled === true`.
  - (append to `trace-tiles.test.ts`) `"a capture of a resized view traces"`: deliveries at `viewGen 1` trace to their sid.
  - `"identity hash unchanged"`: `viewStateHash` of the identity view and a fixed clock equals the 4a hex, pasted from base.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-resize.test.ts test/trace-tiles.test.ts test/cache.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.** `tileJobProblem` keeps every check except `viewGen`.
- [ ] **Step 4: Run green:** the same files, then `cd server && bun test` (report the count).
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): per-session sizes — sheet-resize, view-state caps, sized tiles, tracer" -- src/xlsx/view/streamGeometry.ts server/src/sheet/view-state.ts server/src/cache.ts server/src/sheet/tiles-op.ts server/src/sheet/tile.ts server/src/render.ts server/src/sheet/layout.ts server/src/ws.ts server/src/audit-store.ts server/src/detect/trace.ts server/test/sheet-resize.test.ts server/test/trace-tiles.test.ts server/test/cache.test.ts
  ```

### Task 5: The `sheet-text` pool job and `sheet-fit`

**Files:**
- Create: `src/xlsx/view/autoFit.ts`, `server/src/sheet/text-ops.ts`
- Modify:
  - `src/components/xlsx/sheetResize.tsx:225–330`: re-export from `autoFit.ts`;
  - `server/src/render.ts`: `DocumentHost.sheetText`;
  - `server/src/worker/render-worker.ts:46–49`, `:108`: a `sheet-text` message;
  - `server/src/pool.ts:377`: `Pool.sheetText`;
  - `server/src/ws.ts`: `onSheetFit`.
- Test: create `server/test/sheet-fit.test.ts`

**Interfaces:**
- Produces:
  - `autoFit.ts` exports `autoFitSize`, `AutoFitInput`, `createFitMeasure` and `FitMeasure`, moved verbatim. `sheetResize.tsx` re-exports them, so local imports are unchanged.
  - The job type:
    ```ts
    export type SheetTextJob =
      | { op: 'find'; query: string; matchCase: boolean; sheetIndex: number; sizes: WireSizes; limit: number }
      | { op: 'copy'; sheetIndex: number; selection: WireSelection; sizes: WireSizes; maxBytes: number }
      | { op: 'a11y'; sheetIndex: number; rows: Array<[number, number]>; cols: [number, number]; sizes: WireSizes }
      | { op: 'fit'; sheetIndex: number; axis: 'row' | 'col'; index: number; pane: Box; regions: Box[]; sizes: WireSizes }
    export type SheetTextReply =
      | { op: 'find'; hits: Array<[number, number]>; more: boolean }   // [row, col] in local order; more: hits past `limit`
      | { op: 'copy'; text: string; paragraphs: number } | { op: 'copy'; refused: 'too-large'; bytes: number }
      | { op: 'a11y'; cells: Array<[number, number, string]> }
      | { op: 'fit'; size: number }
    ```
  - `DocumentHost.sheetText(docKey, clock, job): Promise<SheetTextReply>` builds the text views it needs (Task 1) and caches them per `(sheetIndex, clock, sha256(sizes))`, keeping at most 16 per document and releasing on `drop`.
  - `Pool.sheetText(sid, job & {docRef, contentPolicy, clock}, bytesEstimate, signal)`, with the same affinity and admission as `sheetLayout`.
  - `onSheetFit`, in order: the token, `fitArgs`, `'sheet'` rate and busy, `stale-view`, `hidden-target`, then the job.
    - The fit result is `autoFitSize(kind, index, {geometry: view.geometry, sheet: view.sheet, pane, regions, knownRows: limit => view.sheet.rows.slice(0, limit).map(r => r.index), measure})`, with `createFitMeasure` on the server backend. Cell text is measured under `view.scope`'s effective style, as the local resident fit does.
    - The answer is `sheet-fit-result {size}`. Fit never changes view state.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)`):
  - `"fit equals local autoFitSize on the deleted fixture"`: on `textGrid`, columns 0..7 and rows 0..29 under a fixed pane. Each answered size equals local `autoFitSize` run (in the test) over a copy of `textGrid` with its covered and merge-tail cells deleted, with the same pane.
  - `"fit ignores withheld text"`: column D of `unfilledShapeOverText`, whose only long text is covered, fits to the same size as the fixture with B2 deleted. Every column of `unresolvedAnchor` fits to the blank-column size.
  - `"fit is on without copy"`: `copy: false` gets an answer.
  - `"fit is a sheet op"`: the 31st in a minute is refused `rate-limited`, and a collapsed column is refused `hidden-target`.
  - `"local fit unchanged"`: `bunx vitest run xlsx-merge-ce-viewer-autofit-union paged-fit` passes unmodified.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-fit.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.** The move is mechanical. If `autoFit.ts` needs something from `src/components`, only React-free modules may come with it (`measureContext.ts` is React-free).
- [ ] **Step 4: Run green:** the same file; `bunx vitest run xlsx- interactive paged`; and `bun run gates:core` (report the delta).
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): sheet-text pool job and sheet-fit over the text view" -- src/xlsx/view/autoFit.ts src/components/xlsx/sheetResize.tsx server/src/sheet/text-ops.ts server/src/render.ts server/src/worker/render-worker.ts server/src/pool.ts server/src/ws.ts server/test/sheet-fit.test.ts
  ```

### Task 6: `sheet-find`: positions only, local order

**Files:**
- Modify: `server/src/sheet/text-ops.ts` (op `find`), `server/src/render.ts`, `server/src/ws.ts` (`onSheetFind`)
- Test: create `server/test/sheet-find.test.ts`

**Interfaces:**
- Consumes: `textDocument` (Task 1), `createObjectWorkbookSource`, and `WorkbookFind` pulled page by page (`FIND_PAGE_HITS`) until complete or the job's `limit`.
- **Batched by sheet, as `runSearch` batches units** (`ops.ts:109`, `SEARCH_BATCH_UNITS`). The ws side sends one `find` job per sheet in workbook order, with `limit` = 5,000 minus the hits so far. Between jobs it re-checks ownership, the stream phase and the session entry, and re-touches the document ledger. Concatenating per-sheet answers gives `WorkbookFind`'s workbook order, which the parity test checks.
- Produces:
  - Find input: the identity order, `visibility` from the view's axes, and `isColumnVisible: col => size > 0`. `textOf` is `cellDisplayText` under `view.scope?.effectiveStyle(row, col, base) ?? base`, and `prepareScope` returns `view.scope`, as `InteractiveSheetDoc.tsx:1839–1857` builds them locally.
  - `onSheetFind`, in order: the token, `findArgs`, `'text'` rate, fingerprint, ledger, the job, then the chunks (`sheet-find-result` ≤ 48 KiB each, then `sheet-find-done`).
  - Hits never carry text. One find runs per sid: a newer one supersedes it, as `supersede` does (`ops.ts:92`).

- [ ] **Step 1: Write the failing tests** (`}, 30_000)`):
  - `"hit set and order equal local find on the deleted fixture"`: for queries `needle`, `NEEDLE`, `1,234.5` and `2026`, under visible-only, the hits deep-equal local `WorkbookFind` over `textGrid` with `NEEDLE-D20` and the merge tail deleted, mapped to `[sheetIndex, row, col]`.
  - `"withheld cells never match"`: none of these is found:
    - `COVERED-B2`;
    - every covered `ROT-r-c` (an uncovered one is);
    - `OFFVIEW-Z5000` (`OFFVIEW-A1` is);
    - `TAIL-B1`;
    - `HIDDENFMT`;
    - `cfHiddenFormat`'s `7` and `9` (`4` is found);
    - any text on `unresolvedAnchor`.

    After a `cols {1: 8}` resize on `slidUnderAnchor`, `SLID-E1` is not found; before it, it is.
  - `"full finds as local"`: under `'full'`, `COVERED-B2` is found.
  - `"CF-effective display text"`: on a CF fixture whose rule writes `0.0`, the find for `7.0` hits as local find does on the same fixture with the scope.
  - `"batched by sheet"`: on a 3-sheet `textGrid`, `pool.sheetText` is called once per sheet. A socket closed after the first job leads to no second job.
  - `"no text on the wire"`: no message of a find for `needle` contains `needle` (case-insensitive) outside the echoed request.
  - `"limits"`: 6,000 matching cells give 5,000 hits and `truncated: true`; 257 units give `policy-refused`; the 61st `'text'` op in a minute gives `rate-limited`, with no 4008.
  - `"search works without copy"`: `copy: false` answers.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-find.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same file, then `cd server && bun test test/text-search-wire.test.ts` (unchanged).
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): sheet-find — local find over the text view, positions only" -- server/src/sheet/text-ops.ts server/src/render.ts server/src/ws.ts server/test/sheet-find.test.ts
  ```

### Task 7: `sheet-copy`: TSV over servable cells, caps, audit

**Files:**
- Modify:
  - `server/src/sheet/text-ops.ts` (op `copy`; `runSheetCopy`);
  - `server/src/ws.ts` (`onSheetCopy`);
  - `server/src/audit-store.ts:33`, `:75`, `:201–206`, and `server/src/audit.ts:111`.
- Test: create `server/test/sheet-copy.test.ts`

**Interfaces:**
- Consumes: `serializeSelection` with the view's store, the identity order, row visibility from the view's axes and column visibility from zero widths, as `InteractiveSheetDoc.tsx:2611–2659` builds them. `format` is `cellDisplayText` under `scope?.effectiveStyle(...)`, and `prepareScope` returns `view.scope` (P4).
- Produces:
  - `export function runSheetCopy(c: SheetCopyContext): Promise<void>`. It queues on `TextSlot.copyQueue` and follows the P8 order.
  - `CopyRecord` gains `sheetId?: string` and `rect?: {r0, c0, r1, c1}`. `putCopy` writes them, and the migration is a guarded `ALTER TABLE copies ADD COLUMN`.
  - The answer is `copy-result {seq, text, last}`, the Stage 3 chunks.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)`):
  - `"copy TSV equals local on the deleted fixture"`: for a 3-range selection and select-all on `textGrid`, the text equals local `serializeSelection` over the deleted fixture, byte for byte. Hidden row 7 and hidden column F are skipped, as locally.
  - `"withheld cells are empty fields"`:
    - a B1:B3 copy of `unfilledShapeOverText` has three lines and an empty B2 line;
    - `rotatedBox`'s covered cells, `OFFVIEW-Z5000` (in a Z4999:Z5001 copy), `TAIL-B1` and every cell of `unresolvedAnchor` copy as empty fields, and the TSV shape (line and tab counts) equals the local copy's on the deleted fixture.
  - `"caps refuse at n+1 and allow at n, per copy"`: with `copy: {maxParagraphs: 3}`, a 3-text-cell copy succeeds. A 4-text-cell copy is refused `{code: 'copy-too-large', limit: 3, requested: 4}` and sends no text. A 6-cell copy where 3 cells are withheld counts 3 and succeeds.
  - `"… and per session"`: with `copy: {maxParagraphs: 10, maxSessionParagraphs: 5}`, copies of 3 then 2 succeed. The next 1-cell copy is refused `{code: 'copy-quota-exhausted', limit: 5, requested: 1}`.
    - After a resume with `maxSessionParagraphs: 6`, a 1-cell copy succeeds (Review Focus 2).
    - A focused a11y read after that is refused `copy-quota-exhausted` (shared counter).
  - `"copy off"`: `copy: false` is refused `text-unavailable`.
  - `"audit"`: `store.copies(sid)` has `{kind: 'copy', paragraphs: 3, sheetId: '0', rect: {r0, c0, r1, c1}}` and no text column. A database created before the change opens and is migrated.
  - `"select-all on a million rows"` (Review Focus 4): on `millionRows` with `copy: 'full'`:
    - the test first asserts the precondition that local `serializeSelection` of the fixture is ≤ 4 MiB;
    - the select-all copy then equals it in one answer, and the socket stays open.

    A 5 MiB-text fixture gives `copy-too-large {limit: 4194304, requested: <bytes>}`.
  - `"cfHiddenFormat copies as painted"`: B1:B3 copy as `'4\n\n'`, as local does with the scope.
  - `"hostile selections"`: each P7 violation is refused `policy-refused` and charges no paragraphs.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-copy.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same file, then `cd server && bun test test/text-copy-wire.test.ts test/xlsx-units-baseline.test.ts` (unchanged).
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): sheet-copy — servable TSV, copy caps and quota, sheet audit" -- server/src/sheet/text-ops.ts server/src/ws.ts server/src/audit-store.ts server/src/audit.ts server/test/sheet-copy.test.ts
  ```

### Task 8: `sheet-a11y`: unit, focused and off

**Files:**
- Modify: `server/src/sheet/text-ops.ts` (op `a11y`; `runSheetA11y`), `server/src/ws.ts` (`onSheetA11y`)
- Test: create `server/test/sheet-a11y.test.ts`

**Interfaces:**
- Produces: `runSheetA11y(c)` on the sid's `copyQueue`. Its order follows `a11yTurn` (`ops.ts:436`), with P10's modes. The text is `formatCellPresentation(view.sheet, cell, budget, false, cf)` with `cf` from `view.scope` (P4). `'text'` is charged first, off sessions included. The answer is `sheet-a11y-result` chunks, and the audit row is `{kind: 'a11y', sheetId, rect}`.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)`):
  - `"unit mode serves the window, uncounted"`: with `copy: 'full'`, the window A1:H30 of `textGrid` gives, for each cell, the text `formatCellPresentation` gives on the deleted fixture. There is one `a11y` audit row and `copiedParagraphs` is unchanged.
  - `"withheld cells have no text"`: no entry for `COVERED-B2`, a covered `ROT` cell, `OFFVIEW-Z5000` (window Z4990:Z5010), `TAIL-B1`, `cfHiddenFormat`'s B2 and B3, or any `unresolvedAnchor` cell.
  - `"focused counts 1 per non-empty cell"`: with `copy: {maxParagraphs: 3, maxSessionParagraphs: 2}` (focused):
    - a 2×1 window gives `policy-refused`;
    - B1 succeeds (count 1), an empty cell succeeds (count 0), and B3 succeeds (count 2);
    - B5 gives `copy-quota-exhausted {limit: 2, requested: 1}`.
  - `"off"`: with `copy: false`, the request is refused `text-unavailable`, and it is still charged `'text'` (the 61st gives `rate-limited`).
  - `"limits"`: 2,001 cells give `policy-refused`.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-a11y.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same file, then `cd server && bun test test/text-a11y-wire.test.ts` (unchanged).
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): sheet-a11y — window text in unit mode, active cell in focused mode" -- server/src/sheet/text-ops.ts server/src/ws.ts server/test/sheet-a11y.test.ts
  ```

### Task 9: Client session and `StreamWorkbook`: text ops, sizes, viewGen

**Files:**
- Create: `src/components/xlsx/streamText.ts`
- Modify:
  - `src/stream/remote/text-wire.ts:13–38`, `:117`: the sheet request and answer types;
  - `src/stream/remote/session.ts:505`: resume policy per P13;
  - `src/stream/remote/remote-doc.ts:57`: `RemoteSheets.view?`;
  - `src/components/xlsx/streamWorkbook.ts`: sizes, viewGen and the text methods.
- Test: create `tests/stream/remote-sheet-text.test.ts`, `tests/xlsx-stream-text.test.ts`

**Interfaces:**
- Produces:
  - `RemoteSheets.view`, present iff `manifest.sheetView === 1`, with the Task 2 types:
    ```ts
    view?: {
      find(q: { query: string; matchCase?: boolean }, signal?: AbortSignal): Promise<SheetFindAnswer>
      copy(sheetId: string, viewGen: number, sel: WireSelection): Promise<string>
      a11y(sheetId: string, viewGen: number, rows: Array<[number, number]>, cols: [number, number]): Promise<Array<{ row: number; col: number; text: string }>>
      fit(req: SheetFitRequest): Promise<number>
      resize(req: SheetResizeRequest): Promise<SheetViewChange>
    }
    ```
    A refusal rejects `StreamSessionError` with its code and `limit`/`requested`.
  - In `streamText.ts`:
    - `export function wireSelection(s: GridSelection): WireSelection | 'too-large'`, through `rowSetIntervals`;
    - `export function streamFindSession(view, owner: WorkbookOwner, sheetIds: string[], query: SearchQuery): FindSession`, which pages the answer at `FIND_PAGE_HITS` as `FindHit`s with `text: ''`, `start: 0`, `end: 0` and `displaySlot` from the order.
  - `StreamWorkbook` gains:
    - `sizes(sheetId): SessionSizes`;
    - `applyView(sheetId, change: SheetViewChange): void`, which sets `viewGen` and the sizes and notifies;
    - `findSession(owner: WorkbookOwner, query: SearchQuery, options: FindOptions): FindSession`;
    - `copy(sheetId: string, selection: GridSelection): Promise<string>`, which rejects `StreamSessionError('copy-too-large')` with no `limit` when `wireSelection` gives `'too-large'`;
    - `a11y(sheetId: string, rows: Array<[number, number]>, cols: [number, number]): Promise<Array<{ row: number; col: number; text: string }>>`;
    - `fit(req: Omit<SheetFitRequest, 'viewGen'>): Promise<number>`;
    - `resize(req: Omit<SheetResizeRequest, 'viewGen'>): Promise<SheetViewChange>`;
    - `textOps: boolean`, true iff `view` is present.

    Each method fills in the sheet's current `viewGen`. `applyView` updates the held `wireLayout`'s `sessionRows`/`sessionCols`, so `streamSheetGeometry` (Task 4) is the one geometry path.

- [ ] **Step 1: Write the failing tests:**
  - (`tests/stream/remote-sheet-text.test.ts`, fake socket):
    - `"view present only with sheetView"`;
    - `"find is re-issued on resume; copy, a11y and resize reject session-rotated and are never re-sent"` (P13);
    - `"a rejected resize re-fetches sheet-layout"`: the next `layout()` call is made and its `viewGen` adopted (Review Focus 1);
    - `"refusals carry limit and requested"`: `copy-too-large {limit: 3, requested: 4}` rejects with them;
    - `"hostile answers drop the request"`: a hit with row −1, an a11y cell outside the requested window, or a `sheet-view` for another sheet rejects `policy-refused`, and the session stays open.
  - (`tests/xlsx-stream-text.test.ts`):
    - `"applyView moves the geometry"`: after `applyView({sheetId: '0', viewGen: 1, change: {axis: 'col', index: 1, size: 8}})`, the entry's `cols.size(1) === 8`, and `identity('0').viewGen === 1`;
    - `"wireSelection"`: a bitset part becomes intervals; 257 ranges give `'too-large'`;
    - `"stream find pages without text"`: 450 hits give 3 pages (200, 200, 50), and every `text === ''`.
- [ ] **Step 2: Run** `bunx vitest run tests/stream/remote-sheet-text.test.ts tests/xlsx-stream-text.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** `bunx vitest run tests/stream tests/xlsx-stream-workbook.test.ts`, and `bun run gates:core` (report the delta).
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): client sheet text ops, session sizes and viewGen on StreamWorkbook" -- src/components/xlsx/streamText.ts src/stream/remote/text-wire.ts src/stream/remote/session.ts src/stream/remote/remote-doc.ts src/components/xlsx/streamWorkbook.ts tests/stream/remote-sheet-text.test.ts tests/xlsx-stream-text.test.ts
  ```

### Task 10: Viewer: selection, keyboard, resize and fit, find and copy on streams

**Files:**
- Modify:
  - `src/components/xlsx/GridSemantics.tsx:75`, `:284`, `:347`, `:375`, `:395`, `:443`: split `positionsOnly`;
  - `src/components/xlsx/InteractiveSheetDoc.tsx`: `:345–349`, `:1811`, `:2004`, `:2641`, `:3071–3072`, `:3092–3101`, `:3116–3160`, `:3241`, `:3301`, `:3318`, `:3359`;
  - `src/components/remote/RemoteSheetView.tsx`.
- Test: create `tests/xlsx-stream-selection.test.tsx`; modify `tests/stream/remote-sheet-view.test.tsx` (append)

**Interfaces:**
- Consumes: `StreamWorkbook.textOps` and its methods (Task 9).
- Produces:
  - **The `positionsOnly` split.** It now means "gridcells carry no text" only. A new `selectable?: boolean` prop (default `true`) gates click, header, keyboard and select-all selection. 4a's stream passes `selectable: false`.
  - **The 4b stream**, when `textOps` is set:
    - selection, keyboard, Go To and find navigation run the local code unchanged;
    - `resizeLine` on a stream goes to `fit`, then `resize`, then `applyView`, then `publishEntry`, keeping the local announcement strings and P12a's pacing and refusal strings; "Reset sizes" sends `reset: true`;
    - Copy and Ctrl+C go to `copy(sheetId, selection)`, then `copyText`, with the S3 item 10 strings;
    - `FindController.createSession` is `stream.findSession`, and `goToHit` selects the cell, with no hit text rects;
    - a hit with `hidden: 'filter'` announces "This match is in a row hidden by the filter." and opens no inspection (P15).
  - `InteractiveSheetDocProps.onSelection?: (sheetId: string, selection: GridSelection) => void` (`@internal`), called after each committed selection change.
  - `RemoteSheetView`: `allowCopy = can.copy && sheetView`, and `showSearch = sheetView` (search always on).

- [ ] **Step 1: Write the failing tests** (`tests/xlsx-stream-selection.test.tsx`):
  - **`"selection parity: keyboard"`.** Mount a local viewer on `textGrid`'s visible-only text-view document (local mode on the snapshot) and a stream viewer whose fake `RemoteSheets` serves `buildStreamLayout` of the same snapshot. Script on both:
    - Arrow ×3 → Shift+ArrowRight ×2 → Ctrl+ArrowDown → Shift+Ctrl+ArrowRight → PageDown → Home → Ctrl+End → Ctrl+A;
    - a run over hidden row 7, hidden column F and merge A5:C6, under the frozen 1×1.

    The `onSelection` sequences deep-equal, and so do the `aria-activedescendant` values.
  - **`"selection parity: mouse"`:** click C3, Shift+click E8, Ctrl+click G2, a column-header click on B, a Shift+header click on D, a row-header click on 10 and a drag C3→E5. The sequences deep-equal.
  - **`"resize routes to the server"`:** Alt+Shift+ArrowRight on column B calls `view.resize({sheetId: '0', viewGen: 0, axis: 'col', index: 1, size: 72})`. After the fake `sheet-view`, the tiles are requested at `viewGen 1` and the status says 'Column B width 72 pixels. View only; the file is unchanged.' Alt+Shift+Enter calls `fit`, then `resize` with the fitted size.
  - **`"key-repeat folds into one pending resize"`:** five Alt+Shift+ArrowRight presses while the first is in flight lead to exactly 2 `resize` calls, the second with `size: 104`. A `rate-limited` refusal shows 'Too many requests. Try again shortly.'
  - **`"find waits for typing to stop"`:** typing 'needle' over 100 ms makes one `view.find` call, 250 ms after the last key.
  - **`"copy off disables Copy and Ctrl+C"`** (Review Focus 2): with `can.copy: false`, Copy is absent and Ctrl+C makes no `copy` call. After an `onManifestChange` to copy off mid-session, the same holds.
  - **`"no local text index"`:** with `WorkbookFind` and `copySelection` spied (`vi.mock`), a find for 'needle' and a Copy each make 0 calls, and `stream.view.find` and `copy` are called once each.
  - **`"copy refusals speak S3's strings"`:** `copy-too-large {limit: 3}` shows 'Selection too large to copy (limit: 3 paragraphs).', and `copy-quota-exhausted {limit: 5}` shows 'Copy limit for this session reached (5 paragraphs).'
  - **`"4a streams are unchanged"`:** without `sheetView`, a click and arrows select nothing, and there is no Find or Copy. (4a's own cases in `remote-sheet-view.test.tsx` also pass unedited.)
  - (append to `remote-sheet-view.test.tsx`) **`"4b stream mounts Find and Copy"`:** with `sheetView: 1` and `can.copy`, the searchbox and 'Copy selection' are present, and a click selects a cell.
- [ ] **Step 2: Run** `bunx vitest run tests/xlsx-stream-selection.test.tsx tests/stream/remote-sheet-view.test.tsx`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:**
  - the same files;
  - the unmodified `bunx vitest run xlsx- interactive paged sheet viewer tests/components officeview-selection technical-review-e-find-ui`;
  - `npx tsc --noEmit -p .`;
  - `bun run gates:core` (report the delta).
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): selection, keyboard, resize, fit, find and copy on streamed sheets" -- src/components/xlsx/GridSemantics.tsx src/components/xlsx/InteractiveSheetDoc.tsx src/components/remote/RemoteSheetView.tsx tests/xlsx-stream-selection.test.tsx tests/stream/remote-sheet-view.test.tsx
  ```

### Task 11: The a11y grid's text on streams, and the §6.11 notices

**Files:**
- Create: `src/components/xlsx/streamA11y.ts`
- Modify: `src/components/xlsx/InteractiveSheetDoc.tsx` (`cellPresentation` at `:3427`, the notices at `:3352`, inspection at `:1807` and `:3092`)
- Test: create `tests/xlsx-stream-a11y.test.tsx`

**Interfaces:**
- Produces:
  ```ts
  export interface StreamA11yText { textAt(sheetId: string, viewGen: number, at: SourceCell): string | undefined; showWindow(sheetId: string, viewGen: number, rows: Array<[number, number]>, cols: [number, number]): void; focus(sheetId: string, viewGen: number, at: SourceCell): void; onChange(cb: () => void): () => void; dispose(): void }
  export function createStreamA11yText(o: { mode: 'unit' | 'focused' | 'off'; a11y: RemoteSheets['view']['a11y']; now?: () => number; setTimer?: typeof setTimeout }): StreamA11yText
  ```
  - **unit:** a 250 ms rest, uncached cells only, windows split into rects of ≤ 2,000 cells, at most 30 requests per 60 s, and a `rate-limited` retry after 1 s.
  - **focused:** a 400 ms rest, then one 1×1 fetch per uncached cell (Open item 2).
  - **off:** nothing is fetched.
  - `cellPresentation` on a stream returns the cached text, `'Loading…'` while a fetch is pending (the paged string), or `''`.
- Notices (P15), each once per sheet in the status region:
  - `TEXT_UNAVAILABLE` when copy or a11y is off;
  - "Some rows are hidden by the author and can't be shown." when `layout.authorHidden`;
  - "Hidden cells can't be shown in this viewing mode." on any hidden-cell inspection attempt.

- [ ] **Step 1: Write the failing tests** (fake timers):
  - `"unit mode fills the visible window once"`: the first rest sends one request covering the visible cells. A scroll back over the same cells sends nothing, and each gridcell's hidden text equals the served text.
  - `"pacing"`: 40 settled scrolls to new windows in 60 s send ≤ 30 requests.
  - `"focused fetches the rested active cell once"`: arrows over 5 cells within 400 ms each send 0 requests, and a rest sends 1. Returning to that cell later sends 0.
  - `"off fetches nothing and says so once"`: no `a11y` call, and the status region shows `TEXT_UNAVAILABLE` exactly once.
  - `"resize invalidates by viewGen"`: after `applyView`, the next rest refetches the window.
  - `"a11y turned off on resume clears the text"` (Review Focus 2): a manifest change to `caps.a11y: 'off'` clears the gridcells' text, and nothing more is fetched.
  - `"notices"`: an `authorHidden` layout shows the author-hidden string once per sheet. A Go To a collapsed row shows the hidden-cells string and opens no inspection.
- [ ] **Step 2: Run** `bunx vitest run tests/xlsx-stream-a11y.test.tsx`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same file, `bunx vitest run tests/stream interactive xlsx-viewer`, and `bun run gates:core`.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): a11y grid text on streams by mode, paced; S4 §6.11 notices" -- src/components/xlsx/streamA11y.ts src/components/xlsx/InteractiveSheetDoc.tsx tests/xlsx-stream-a11y.test.tsx
  ```

### Task 12: Tiles by default; the harness and browser checks

**Files:**
- Modify:
  - `src/stream/remote/open.ts:20–28`;
  - `src/components/OfficeFile.tsx:46`, `:125`, `:262`;
  - `src/debug/stream-viewer.tsx`;
  - `server/test/app.ts` (`?doc=text-grid|slid-under-anchor`).
- Test: modify `tests/stream/remote-sheet.test.ts` (append), `tests/browser/stream-sheet.spec.ts` (append)

**Interfaces:**
- Produces: `streamSheets?: 'tiles' | 'units'`.
  - Unset (the default) sends `features: ['sheet-view@1']`: tiles with 4b on a 4b server, and units on a 4a server (P1).
  - `'tiles'` sends `['sheet-tiles@1', 'sheet-view@1']`.
  - `'units'` sends no `features`, as Stage 3 did. `StreamSession.open` with no `features` stays Stage 3's HELLO (`remote-session.test.ts:341` unchanged).

- [ ] **Step 1: Write the failing tests:**
  - (append to `remote-sheet.test.ts`):
    - `"openRemoteDocument sends sheet-view@1 by default"`: `features` deep-equals `['sheet-view@1']`;
    - `"explicit tiles sends both features"`;
    - `"streamSheets: 'units' sends Stage 3's HELLO"`: the keys are exactly `['client', 'recVersions', 'reqId', 'type']`.
  - (append to `stream-sheet.spec.ts`, against the default client):
    - `"drag selects a range"`;
    - `"find moves to a match and shows no hit text"`: no element's text contains the query except the search box;
    - `"copy puts servable TSV on the clipboard"` (clipboard permission; the covered cell's field is empty);
    - `"resizing a column moves the tiles"`: after a header drag, `__ovSheetStats.pendingTiles` reaches 0 and the column header has the new width;
    - `"a slid-under cell stops answering after resize"`: `SLID-E1` is found before shrinking B, and not after.
- [ ] **Step 2: Run** `bunx vitest run tests/stream/remote-sheet.test.ts` and `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npm run test:browser:stream`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:**
  - the same commands;
  - unmodified: `bunx vitest run tests/stream tests/office-file-stream.test.tsx tests/officeview-streaming.test.tsx`;
  - the unmodified Stage 3 browser stream specs. If any existing test or spec needs an edit to pass under the new default, STOP and report, without editing it;
  - `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npx playwright test --config playwright.viewer.config.ts --project=chromium tests/browser/interactive-viewer`, unmodified.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "feat(stream-s4b): streamed spreadsheets open in tiles mode with sheet-view by default" -- src/stream/remote/open.ts src/components/OfficeFile.tsx src/debug/stream-viewer.tsx server/test/app.ts tests/stream/remote-sheet.test.ts tests/browser/stream-sheet.spec.ts
  ```

### Task 13: Compatibility legs for 4b

**Files:**
- Modify: `scripts/stream-s4/compat-matrix.ts`

**Interfaces:**
- Produces: `bun scripts/stream-s4/compat-matrix.ts --base 53668e81 --base4a 1e2fbf4d`, whose new legs are:
  - **(e)** the 4a client (`1e2fbf4d`'s `StreamSession`, `features: ['sheet-tiles@1']`) against the current server: the 4a client's `validateManifest` accepts the manifest, `'sheetView' in raw === false`, and one tile decodes;
  - **(f)** the current default client against the 4a server: units mode (no `sheetMode`), and a Stage 3 search hit and copy arrive;
  - **(f′)** the current client with `streamSheets: 'tiles'` against the 4a server: tiles mode, `doc.sheets.view === undefined`, and one tile decodes;
  - **(g)** the current client with `streamSheets: 'units'` against the current server: units, and a Stage 3 copy works.

- [ ] **Step 1: Add the legs**, with a second detached worktree for `--base4a`.
- [ ] **Step 2: Run** the command. Expected: eight PASS lines, and `git worktree list` unchanged after.
- [ ] **Step 3: Commit.**
  ```bash
  git commit -m "test(stream-s4b): compatibility legs — 4a client × 4b server, 4b client × 4a server" -- scripts/stream-s4/compat-matrix.ts
  ```

### Task 14: Evidence: wire audit on text answers, the withholding battery, and the ink oracle

**Files:**
- Create: `server/test/sheet-text-audit.test.ts`, `server/test/sheet-ink-oracle.test.ts`

**Interfaces:**
- Consumes: the whole 4b server path through the harness `Client`.

- [ ] **Step 1: Extend the wire audit** to text answers (`}, 30_000)`). Capture every server→client byte of a `sheet-view` session that does each of the following:
  - `sheet-layout`;
  - a find for each sentinel and marker string;
  - a select-all copy (`copy: 'full'`);
  - a unit-mode a11y over the used range in ≤ 2,000-cell windows;
  - a fit of every used column.

  Run it for `sentinels()`, every XLSX hidden fixture, and the Task 1 withholding fixtures. Then check:
  - `"no sentinel or withheld text in any text answer"`: no decoded message contains `SENTINEL_F`, `SENTINEL_N`, `1234.5678`, `=LEN`, any `marker`, a hidden sheet's name, `COVERED-B2`, a covered `ROT-` text, `OFFVIEW-Z5000`, `TAIL-B1`, `HIDDENFMT` or `ABS-B2`. `SLID-E1` is absent after the shrink, and `cfHiddenFormat`'s `7` and `9` never occur in a text answer.
  - `"find answers carry positions only"`: every `sheet-find-result` has only the keys `reqId`, `type` and `hits`, and its hits are number tuples.
- [ ] **Step 2: Write the withholding battery** (S4 §9 4b) as one table-driven test over `{offViewportCovered, rotatedBox, unresolvedAnchor, cfHiddenFormat, absoluteOverB2 resized}` × `{find, copy, a11y}`. Each case asserts the withheld set is empty in the answer, from three scroll histories:
  - no tiles fetched;
  - every tile fetched;
  - only the tiles that hold the cell fetched.
- [ ] **Step 3: Write the ink oracle** (P16).
  - `"ink-withheld cells are blank in the text view"` (`}, 30_000)`) runs on every fixture of Task 1 plus `labelBesideShape` and `spillUnderShape('left' | 'right')`.
  - `"… on the corpus"` (`}, 120_000)`) runs on each corpus XLSX: for each sheet, the text cells in tile (0,0) at zoomStep 5, dpr 1, at most 64 cells per sheet and at most 20 sheets. It asserts that it visited at least one corpus file.

  A failure is a fix to `coverage.ts` or `text-view.ts`, made with a failing case first, never a change to the oracle.
- [ ] **Step 4: Run** `cd server && bun test test/sheet-text-audit.test.ts test/sheet-ink-oracle.test.ts`. Expected: PASS, since Tasks 1–8 supply the behaviour. These are evidence, not RED-first.
- [ ] **Step 5: Commit.**
  ```bash
  git commit -m "test(stream-s4b): text wire audit, withholding battery, ink oracle" -- server/test/sheet-text-audit.test.ts server/test/sheet-ink-oracle.test.ts
  ```

### Task 15: Gates, validation and spec amendments

**Files:**
- Create: `docs/validation/streaming-s4b.md`
- Modify:
  - `docs/superpowers/specs/2026-11-22-streaming-s4-design.md`: an "As built (4b)" block under §11.2;
  - `docs/superpowers/specs/2026-11-07-server-streaming-design.md`: §6.2, §9.3, §11, §12, §14 and §19, as dated amendments;
  - `docs/superpowers/specs/2026-11-11-streaming-s3-design.md` §4.1: the sheet copy unit.

- [ ] **Step 1: Run every gate, one at a time.**
  - `cd server && bun test` and `bunx vitest run`;
  - `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npm run test:browser:stream`, and the unmodified viewer browser suites with `OV_STREAM_PORT=4281`;
  - `npx tsc --noEmit -p .`, `npm run -s typecheck:server` and `npm run -s lint`;
  - `bun run gates:core`: red as at the base, with the delta reported;
  - the compatibility matrix;
  - `bun scripts/stream-s4/tile-p95.ts`, recorded and not gated;
  - `git diff --numstat 1e2fbf4d -- tests/stream`: 0 deletions.

  Pre-existing failures (4a §Known flakes) are listed, not fixed.
- [ ] **Step 2: Write the amendments.**
  - **S4 "As built (4b)":** P1, P3–P14 and P16 as built.
  - **Parent:**
    - §6.2: the `sheet-view@1` feature;
    - §9.3: `viewStateHash` sizes as `WireSizes`;
    - §11/§12: item 2 (the analytic rule as the sole rule; the ink rule as a test oracle) and item 3's answers half (current-view coverage);
    - §14: the `'sheet'` class (fit, resize) and the 4b ops under `'text'`;
    - §19: item 6's fit limit ("fit reveals a width bound of servable text").
  - **S4 §10 item 1, if the owner confirms Open item 1:** merge tails are blanked in answers, not in `projectXlsx`.
  - **Spec gaps and deviations:**
    - `wholeCell`, the find `cursor`, `hidden: 'filter'` before 4c, the client freeze choice and fit's missing pane;
    - `sheet-a11y`'s `rect`, which became `rows`/`cols`;
    - the notice condition (S4 §6.11 against S3 item 10);
    - §6.3's admission, deferred to 4c, and its two-sheet cap applied to sort and filter only;
    - P1's mixed-version behaviour.
- [ ] **Step 3: Write `docs/validation/streaming-s4b.md`:**
  - each S4 §9 4b "done when" line with the test that proves it;
  - the gate table, the size-gate delta, the compatibility output and the p95 record;
  - what is not machine-checked: screen readers (VoiceOver/NVDA) on the stream grid, non-Chromium browsers, and real clipboards beyond Chromium.
- [ ] **Step 4: Commit.**
  ```bash
  git commit -m "docs(stream-s4b): validation, spec amendments" -- docs/validation/streaming-s4b.md docs/superpowers/specs/2026-11-22-streaming-s4-design.md docs/superpowers/specs/2026-11-07-server-streaming-design.md docs/superpowers/specs/2026-11-11-streaming-s3-design.md
  ```

## Spec coverage

| S4 spec | Task |
|---|---|
| §6.1 rows marked 4b: `search`/`findSession` → `sheet-find`; `styleOf`/`formatting` → server answers; copy; a11y; auto-fit; manual resize | 6, 9, 10; 1, 8; 7, 10; 8, 11; 5, 10; 4, 10 |
| §6.2 step 3 (merge tails, `;;;`, CF `showValue="0"`), step 4 (all-blank on unknown coverage), step 5 (current-view coverage) | 1 (P3–P5) |
| §6.3 view state: sizes, `viewStateHash`, lifetime, `view-state-limit` | 4 (P12) |
| §6.5 tiles under sizes, cache key, `sheet_views`, tracer | 4 |
| §6.7 join for sheets (analytic sole rule, ink oracle), find, copy, a11y, fit, resize | 1, 14; 6; 7; 8, 11; 5; 4 |
| §6.8 `'sheet'` (fit, resize), `'text'` (find, copy, a11y), message caps | 3, 4, 5, 6–8; 2 |
| §6.10 advertise tiles by default; compatibility | 12, 13 (P1) |
| §6.11 notices; Find and Copy in the toolbar | 11, 10 (P15) |
| §6.12 items 2 and 4 | 1, 6, 14 |
| §9 4b "done when": selection parity | 10 |
| find-hit sets and copy TSV equal local on the snapshot | 6, 7 |
| copy caps at n / n+1, per copy and per session | 7 |
| off-viewport covered cell and rotated box withheld from find, copy and a11y | 1, 6–8, 14 |
| wire audit extended to text answers | 14 |
| no local text index (spy) | 10 |
| §10 items 1 (P3), 2, 3, 4 (`'sheet'`), 6 (fit limit) | 15 |
