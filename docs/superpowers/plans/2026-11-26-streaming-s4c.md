# Streaming Stage 4c: server sort and filter, value lists, range detection, the pivot refusal and the withholding tests (implementation plan)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** On a tiles-mode XLSX stream the viewer sorts, filters, lists filter values and detects ranges as local mode does. The server computes every row order, visibility mask and value list from the sheet's servable snapshot. An author-hidden row never shows (C1), a withheld cell sorts and lists as blank, typed number and date conditions are refused on protected streams, and every pivot or control command is refused.

**Architecture:**
- **Server.** One more accessor, `sheetKeyView`, is the sheet every key answer reads: sort keys, filter evaluation, value lists and range-detect occupancy. It is 4b's text view at the identity view (base coverage only, S4 §6.2 step 5), plus blank zero-width columns and masked strings (P5).
  - The pool worker runs the **local resident engine** over it: `viewerControls` (`src/components/xlsx/viewerControlsClient.ts:127`), which stages each op in a private `DocumentSession`, exactly as local mode does. Parity therefore holds by construction. The tests check it against fixtures with the withheld cells physically deleted.
  - A session's view is a small canonical spec `{sort, sortFilter, filter}` (P2). The worker rebuilds the row order and mask from it by replay, as the local session rebuilds an evicted order from its recipe (`src/worker/viewerSession.ts:1264–1275`). Tiles, text ops, `viewStateHash`, `sheet_views` and the tracer all carry that one spec, never the bands.
- **Client.** A new manifest key `sheetSort: 1` adds sort, filter, value lists and range detection to a `sheet-view@1` session (no HELLO change). `StreamWorkbook` holds the order and mask from server deltas. A stream `ControlsClient` (`streamControls.ts`) backs the existing sort and filter menus.

**Tech Stack:** TypeScript. Server: Bun (`server/`, `bun test`, `bun:sqlite`, pool `Worker`s). Client: React (`src/components/xlsx`, `src/components/remote`), vitest, Playwright (a new `stream-4c` project).

**Spec:** `docs/superpowers/specs/2026-11-22-streaming-s4-design.md`, cited as "S4 §N" (binding). The relevant sections:
- §6.6 and §9 4c (scope and "Done when");
- §6.1 (the 4c rows), §6.2 (steps 2–6, C1, the parity oracle), §6.3 (view state, `viewStateHash`, the ceiling and admission), §6.4, §6.8, §6.9, §6.11, §6.12;
- §10 items 1, 3, 4, 6 and 11;
- "As built (4a)" and "As built (4b)", whose rulings bind here, and `docs/validation/streaming-s4b.md` (F17: the admission was deferred to 4c).

Parent: `docs/superpowers/specs/2026-11-07-server-streaming-design.md`, cited as "parent §N": §6.4, §9.3, §12 (C1), §14 and §16 T6/T8. Format exemplar: `docs/superpowers/plans/2026-11-23-streaming-s4b.md` (cited as "4b P*n*"). Base: main `fb624989` (4a and 4b merged).

## Open items for the owner

1. **How a 4c client learns that sort is served.** Both obvious routes touch an existing assertion.
   - A new default HELLO feature would edit `tests/stream/remote-sheet.test.ts:457–458`, which pins the default to exactly `['sheet-view@1']`. That breaks the append-only rule.
   - Advertising on the manifest needs no HELLO change. A 4b client ignores unknown manifest fields (`src/stream/manifest.ts:3`). But it widens `server/test/sheet-view-handshake.test.ts:56`, which pins a `sheetView` manifest's exact key set.
   - **Recommendation:** advertise. A 4c server adds `sheetSort: 1` to every `sheetView` manifest, plus `typedFilters: true` under `'full'`. Permit the one widening of that server assertion to `[...TILES_KEYS, 'sheetView', 'sheetSort']` (outside `tests/stream`). Mixed versions stay correct both ways, and no HELLO, `tests/stream` or browser spec changes.
   - The plan is written to this recommendation (P1). If the owner refuses, Task 3 STOPs: the fallback is an opt-in `streamSheets: 'sort'` that sends `['sheet-view@1', 'sheet-sort@1']`, which leaves the default client without sort.
2. **Find hits in rows hidden by the session's filter.** S4 §6.7 says such a hit is reported as `hidden: 'filter'`. Local find never matches a filter-hidden row (`src/xlsx/view/find.ts:563`), so a parity answer has no such hit. 4b ruled for local parity and built the flag unused.
   - **Recommendation:** keep local parity. 4c's find runs over the session's mask and never sends `hidden: 'filter'`. The client keeps the 4b message for a hit that a later filter hides.
   - The plan is written to this recommendation (P9).
3. **Masked strings, a new withholding class.** A text cell whose format's text section has no `@` (`;;;"****"`) paints other text than its value. S4 §6.6 keeps text conditions and sorts by value, so a text rule over the raw string would bisect a string the viewer cannot see. That is the same oracle the owner closed for typed number conditions.
   - **Recommendation:** under `visible-only`, the key view reads such a cell as its painted text, with no number format. It sorts, lists and matches as `****`. A format that contains the string (`"ID-"@`) is untouched, so ordinary prefixes keep local parity.
   - The plan is written to this recommendation (P5). If the owner declines, drop P5's second bullet and the `MASKED` assertions, and record the limit in parent §19.

## Global Constraints

- **Owner decisions (S4 §2, §11.1, as built):**
  - full local viewer; approach A; this PR is 4c;
  - filter lists are display text only; view state is per session; sort and filter have their own rate class;
  - typed number and date conditions are refused on protected streams and allowed on `'full'`;
  - merge-covered, undrawn and covered cells are hidden content; covered stays covered;
  - pivot, slicer and timeline commands are refused on every stream (S4 §6.9 design call);
  - everything fails closed.
- **4a and 4b as built bind:**
  - wire `sheetId` = the decimal index among the projected sheets;
  - the session clock `(day, timeZone)` (4a R9);
  - `sheetTextView` is the one text accessor (4b P2), and `TextView.styleAt` is the painted pre-CF style (final review R1);
  - fit runs the viewer's own `layoutHeight` (4b R2);
  - `sheet-view@1` and `sheet-tiles@1` behave as 4b built them (F1, F2);
  - `MAX_SESSION_SIZES` = 4,096 per axis per sheet; `viewStateHash` never includes the sid.
- **Every key answer under `visible-only` decides "undrawn" from what is painted** (4b's lesson). That covers:
  - the cell's own style plus its table or pivot region style (`TextView.styleAt`), then CF;
  - zero-height rows and zero-width columns.

  It never decides from `hidden` flags alone. A withheld cell sorts as blank, matches only a blanks filter or a blank-matching condition, and lists only as "(Blanks)" (S4 §6.2 step 6).
- **C1** (parent §12, S4 §6.2 step 2). Under `visible-only`, a row in the body of a filter range is visible iff `matches(row) && !authorHidden(row)`, never `matches(row)` alone.
  - `authorHidden` is the authored geometry's collapsed rows (`ServerSheetLayout.collapsedRows`): `hidden="1"`, `ht="0"` and rows a saved autoFilter hid.
  - The tier is immutable for the session.
  - Under `'full'` the local rule holds: the filter decides alone (`src/xlsx/view/visibility.ts:149–154`).
- **Limits (S4 §6.6, §6.8, plan calls):**
  - at most `SORT_FILTER_MAX_RULES` = 64 sort keys and 64 filter columns (P4);
  - text operands ≤ 256 UTF-16 units, run under the local `WorkBudget`;
  - value lists ≤ 10,000 distinct entries (`VALUE_DISPLAY_LIMIT`), pages ≤ 500 entries (`VALUE_PAGE_LIMIT`) and ≤ 48 KiB (`TEXT_MESSAGE_BYTES`);
  - every request message ≤ 64 KiB (`TEXT_CAP`); every answer message ≤ 64 KiB.
- **Rates:**
  - `'sheet'` (`SHEET_RATE_PER_MIN` = 30 per sliding 60 s, one in flight through `SessionRates.sheetBusy`, refuse-only, never toward 4008) now also covers `sheet-sort`, `sheet-filter`, the first page of each `filter-values` request and `range-detect`. Later cursor pages charge only `'message'`.
  - `sheet-order` (the view fetch, P7) is charged to `'unit'`, as `sheet-layout` is.
  - The class is charged before the busy check (4b).
- **Never change limits, timeouts or baselines.** This covers the per-IP gate (30 new sessions per minute, `server/src/rate.ts:104`), `LAYOUT_LIMITS`, the copy caps, the size-gate baseline and every test timeout already in the tree.
- **Units mode, 4a and 4b sessions are unchanged.**
  - A `sheetView` session's manifest is 4b's plus `sheetSort` (and `typedFilters` under `'full'`); its `sheet-layout` keeps 4b's bytes until it sorts or filters. Every other session's manifest and layout are 4b's, key for key.
  - `server/test/xlsx-units-baseline.test.ts` and every Stage 3 and 4b suite pass unedited.
  - Every `tests/**` local and viewer suite passes unmodified.
- **`tests/stream/**` edits are append-only:** `git diff --numstat fb624989 -- tests/stream` shows 0 deletions. `tests/browser/stream-sheet.spec.ts`, `stream-sheet-4b.spec.ts` and every other existing browser spec stay byte-identical.
- **Browser.**
  - The `stream` project sits at exactly the per-IP limit. New specs go in `tests/browser/stream-sheet-4c.spec.ts`, run as their own project `stream-4c` against a third `server/test/app.ts` on `OV_STREAM_4C_PORT` (default `OV_STREAM_PORT + 20`).
  - `STREAM_SPECS` matches every `stream-*.spec.ts`, so the `stream` project's `testIgnore` gains `STREAM_4C_SPECS`.
  - Run: `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npm run test:browser:stream`.
- **No runtime import from `src/xlsx/**` in `src/stream/**`.** `sheet-wire.ts` takes type-only imports, as it already does.
- **Bun 1.4.5 tests:**
  - no `expect(...).rejects`/`.resolves` under `server/test` (`no-expect-rejects.test.ts`); use `rejection`/`rejectionMessage` (`server/test/fixtures/rejection.ts`);
  - `await pool.dispose()` (or `await disposePool()`) in every after-hook (`no-unawaited-dispose.test.ts`);
  - new Workers only through `startWorker` (`server/src/worker-start.ts`). 4c adds a message type, not a worker.
- **Heavy tests:** `}, 30_000)` on every server test that builds more than one snapshot, sorts more than 10,000 rows or renders more than 4 tiles; `}, 120_000)` on the corpus parity test (Task 13).
- **Determinism:** no answer reads a clock or randomness except the opaque list ids (P6). Date filters on `'full'` use the session clock (`clockEpochMs(clock)`, `clock.timeZone`), never `Date.now()`.
- **Size gate (T11):** `bun run gates:core` is red at the base (index closure, 426,556 B). Record it in Task 3 and after every client task, report the delta, never re-baseline. New client code is reachable only from the lazy `RemoteSheetView` and `InteractiveSheetDoc` chunks.
- **Reviewer:** a different model family from the implementer, for each task.
- **Commits:** by path, `git commit -F <msg> -- <paths>`; never `-a`, never amend, never a bare `git stash`. Messages end with:
  ```
  Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01B5aP5T5pG9oGcVPzuNpnmG
  ```
- **Commands:**
  - server: `cd server && bun test <file>`; client: `bunx vitest run <file>`;
  - typecheck: `npx tsc --noEmit -p .` and `npm run -s typecheck:server`; lint: `npm run -s lint`;
  - corpus: `bun scripts/fetch-corpus.ts`, run before any corpus step.

## Plan calls (spec silences and conflicts, decided here; Task 14 writes the wire-visible ones into the specs)

- **P1. The 4c capability is a manifest key** (Open item 1). A 4b client ignores unknown manifest fields, and a 4b server never sends the key.
  - A 4c server serves `sheet-sort`, `sheet-filter`, `filter-values`, `range-detect` and `sheet-order` to every `sheetView` session. Its manifest carries `sheetSort: 1`, plus `typedFilters: true` under `'full'` only.
  - The HELLO is unchanged: there is no new feature, and `streamSheets` behaves as 4b built it.
  - A session without `sheetView` (units mode, or `sheet-tiles@1` only) refuses the five ops `unsupported-op`, uncharged, as 4b leaves off-mode fit and resize uncharged (`server/src/ws.ts:1784`).
  - `sheet-layout` keeps 4b's bytes (`sort: null`, `filterSummary: []`) until the session itself sorts or filters, which no 4b client does.
  - **Why:** mixed deployments keep working in both directions, and no `tests/stream` assertion changes.
- **P2. The canonical view is `{sort, sortFilter, filter}`, not `{sort, filter}`.** S4 §6.3 hashes "the canonical spec: sort keys, filter rules". But a local sort pins the rows hidden by the filter in force when it ran (`src/xlsx/view/sort.ts:689`), and a later filter keeps that order. So the order depends on history.
  - The local session records exactly this as `recipe.sortFilter` and rebuilds an evicted order from it (`viewerSession.ts:1264–1275`).
  - Plan call: the view state carries `sortFilter`, the filter committed when the current sort ran, and the derived view is a pure function of `(key view, sort, sortFilter, filter)`. `sortFilter` is `null` whenever `sort` is.
  - **Why:** otherwise two sids with equal `{sort, filter}` but different histories would share wrong tiles (T6).
- **P3. Replay.** The worker derives a view by replaying three ops through the local engine from the identity view: filter(`sortFilter`), then sort(`sort`), then filter(`filter`). Each op is skipped when its spec is `null`.
  - **Assumption, checked in Task 5 (STOP if false):** a replay passes the guards the original ops passed. The local guards depend on the spec and the structure, except `drawing-impact`. If every accepted op leaves every cell-anchored drawing's display span unchanged (read `guardFilter` and `guardSort` in `src/xlsx/view/operationGuard.ts` to confirm), then every reachable view has the identity view's spans.
  - A replay that refuses anyway fails closed: the job errors, and its tiles or answer are refused `policy-refused`.
  - How the spec moves: a committed sort sets `sortFilter := base.filter`; a committed filter keeps `sort` and `sortFilter`; clearing the sort sets both to `null`.
  - Derived views are cached per `(sheetIndex, clock, canonical hash)`, at most 8 per document (LRU), and rebuilt after a worker recycle.
- **P4. Target authorization** (S4 §6.6), before any engine run.
  - **First the shape check** (`*Args`, Task 2), refused `policy-refused`, uncharged and with no audit row. It rejects:
    - a range outside the sheet grid (`r1 < 1,048,576`, `c1 < 16,384`);
    - header or totals counts other than 0 or 1, or a body of fewer than 1 row;
    - more than 64 keys or columns, or a key or rule column outside the range;
    - a text operand over 256 units, or a malformed rule.
  - **Then the `'sheet'` rate and busy check** (`rate-limited`).
  - **Then, in order,** each failure one refusal that moves nothing, with its `sheet_ops` row:
    - `stale-view` for an old `viewGen`;
    - `hidden-target` when a key column, a filter column, or the range's `c0` or `c1` is a collapsed column of the authored layout;
    - `coverage-unknown` under `visible-only` on an unknown-coverage sheet (S4 §6.2 step 4);
    - `filter-condition-refused` under `visible-only` for a `number`, `date` or `datePreset` rule, or a `compound` rule with such a child. A compound of two text children is allowed;
    - `stale-filter-values` for a values rule whose list was not issued to this sid for this `(sheetId, range, col)` at the current `viewGen`, or whose positions are out of range;
    - `view-state-limit` (P11).

  `N` = 64 because it is Excel's sort-level limit; the same number bounds filter columns.
- **P5. The key view, `sheetKeyView(snap, projected, o)`** is `sheetTextView(snap, projected, NO_SESSION_SIZES, o)`: base coverage only, never the session's sizes or order. Under `visible-only` it adds:
  - **painted styles.** Each content cell's style becomes its painted, CF-effective style: `view.scope?.effectiveStyle(r, c, view.styleAt(r, c, own)) ?? view.styleAt(r, c, own)`. It gets a fresh `styleIndex` in the key view's own style table, so the engines' index-based lookups (`viewerControlsClient.ts:183`, `:187–211`) format labels, display lanes and the date classification with what the tile paints, not with the own style. This is 4b's R1 on the key path: a table DXF `0,"K"` over a General `1234` lists `1K`, never `1234`;
  - every cell of a collapsed (zero-width) authored column, blank;
  - **masked strings** (Open item 3). A string cell whose painted display text does not contain its raw string is read as that display text, with no number format.

  Under `'full'` the key view is the identity (local parity, own styles). Value lists, sort keys, filter evaluation and range detection all read it, so a list entry and a filter over that entry agree (S4 §6.12 item 7).
- **P6. Value lists.**
  - **Request:** `filter-values {sheetId, viewGen, range, col, search?: {text, wildcards}}` for a first page, or `{sheetId, viewGen, cursor}` for a later one.
  - **Answer:** `filter-values-result {list, items: [[id, text, flags], …], blanks: {present, selected}, total, truncated, cursor?, classification?}`. `flags` is a bit set: 1 = selected, 2 = every merged value is an error.
  - Entries are local's value entries (`ValuesDraftPort.read`, `VALUE_PAGE_LIMIT` per read) merged by full display text: `label`, or `inspect(ordinal)` when `labelTruncated`. They keep local order (first occurrence), so positions are deterministic.
  - `id` is the entry's position in the merged unsearched list. `text` is local's label (≤ 256 units). A search answer lists matching entries with their full-list ids.
  - `list` is an opaque id (`crypto.randomUUID()`), one per `(sid, sheetIndex, range, col, viewGen)`, at most 16 per sid (LRU). An evicted list's ids are `stale-filter-values`.
  - `classification` (first page only) is the local `ColumnClassification` over the key view. The menus need it, and it says only what kinds of servable values the list already shows.
  - `selected` reflects the committed rule of that column, else all selected.
  - Cap: 10,000 entries (local `VALUE_DISPLAY_LIMIT`, `src/xlsx/view/valueIndex.ts:29`). S4 §6.6's "no local cap was found" is wrong; this is reported.
  - `inspect` on a stream returns the label as sent, with no further request.
- **P7. The view delta on the wire.** A sorted 1,048,576-row band is 4 MiB of `sourceAtOffset`, so it cannot fit one message.
  - `sheet-sort` and `sheet-filter` are answered by `sheet-order` parts: `{type: 'sheet-order', seq, last, head?, band?, mask?}`.
    - `head` (seq 0) is `{sheetId, viewGen, counts: {visible, total}, sort: {range, keys} | null, filterSummary: [{col, kind}]}`.
    - Each `band` part carries `{r0, r1, from, data}`: base64 of little-endian `Uint32` `sourceAtOffset` entries from offset `from`.
    - Each `mask` part carries `{r0, r1, from, data}`: base64 words of the display-slot bitset.
  - Every part is ≤ 64 KiB. Caps: ≤ 160 parts, ≤ 1 band (`sortRowOrder` returns one band over the sort range on both of its paths, `src/xlsx/view/sort.ts:724–749`), band rows ≤ 1,048,576. A decoder rejects a gap, an overlap, a non-permutation, or anything over a cap.
  - Size metrics are not sent: the client builds them with `createSheetMetrics({order, visibility})` from its layout, and a test pins equality with the server's geometry.
  - `sheet-order {sheetId}` (a request, `'unit'` rate) answers the same parts for the current view. The client sends it after a layout whose `sort` or `filterSummary` is not the identity view.
  - 4b's `sheet-view` resize answer is unchanged.
- **P8. Progress, cancel and lost replies.**
  - The server sends no intermediate progress; the worker runs a whole op as one job. The client reports `completed 0 / 1` when it sends and `1 / 1` on the answer, with the local messages ('Sorting view', 'Filtering view').
  - `sheet-cancel {target}` with a sort or filter `reqId` aborts its pool job. A cancel that lands before the commit point (the synchronous apply after the job returns) commits nothing, and the op is refused `cancelled`. After that point the commit stands and the answer is sent; the client adopts any `sheet-order` it receives.
  - On resume, `sheet-sort` and `sheet-filter` reject `session-rotated` and are never re-sent. The client then re-fetches `sheet-layout` and `sheet-order`. `filter-values`, `range-detect` and `sheet-order` are re-issued.
- **P9. Text ops under the view.**
  - Find, copy, a11y and fit read the session's order and mask: `textInputs` takes the derived view in place of the identity order and authored visibility. Fit's geometry is `buildSheetGeometry(view.sheet, {sizes, order, visibility})`.
  - The text view's current-view coverage (4b P5) runs whenever sizes are non-empty or the view is not the identity view, over that geometry, so a cell a sort slides under a fixed drawing is blank to text answers.
  - Find follows local find and never matches a filter-hidden row (Open item 2).
- **P10. Drawings under a sort** (S4 §8): as locally. Cell anchors map through `sourceToDisplay`. A sort whose band meets a cell-anchored drawing, or a filter that would move one, is refused by the local guard (`drawing-impact`, `src/xlsx/view/operationGuard.ts:151`, `:285`). Recorded with the `sortUnderShape` fixture.
- **P11. View-state ceiling and admission** (S4 §6.3; 4b F17 deferred it here).
  - A sheet is "ordered" when its view has a sort or a filter. At most `MAX_ORDERED_SHEETS` = 2 per sid (S4 §6.3's suggestion). A third is `view-state-limit`, and resizes stay unlimited by tab (4b P12).
  - Admission: each distinct derived view `(docRef, contentPolicy, sheetIndex, clock, canonical hash)` is charged `orderBytes = 8 × bodyRows + 160 KiB` to the `DocumentLedger`. It uses a new pinned, reference-counted hold (`DocumentLedger.hold(key, bytes): boolean`, `release(key)`). Unpinned parsed documents may be evicted to make room, as any parse is.
  - A hold that cannot fit is `view-state-limit`. It is released when no sid holds that view. `DOCUMENT_CACHE_BYTES` is unchanged.
- **P12. Audit.**
  - `sheet_views.spec` becomes `{sort, sortFilter, filter, sizes}`. Values rules are stored as `{list: {range, col}, positions, mode, blanks}`. Typed operands are stored in the clear (S4 §6.5 leaves it to the builder), because a hashed view cannot be traced; they never go into `sheet_ops`. A record without `sortFilter`/`filter` (4b) reads as `null`.
  - New table `sheet_ops (sid, sheet_id, kind, view_gen, range, rules, outcome, at)`:
    - `kind` ∈ `sort | filter | filter-values | range-detect`;
    - `rules` is JSON `[[col, ruleKind | 'asc' | 'desc'], …]` and never holds an operand;
    - `outcome` is `ok` or the refusal code.
  - A row is written for every request that passes its `*Args` check, before the answer is sent. A failed write refuses the op `policy-refused`.
- **P13. Engine refusals on the wire.** A local `OperationRefusal` is sent as `refused {code: 'view-refused', reason, range?, count?, at?}`.
  - `reason` is the local reason. `range` is a source rect for `merge`, `fixed-array`, `dynamic-spill`, `outside-column-data` and `pivot`. `count` and `at` (a source cell) are sent for `formula-value-unavailable`.
  - `drawingId` and every `message` are dropped; the client rebuilds local strings from `reason`.
  - `incomplete` cannot occur on a resident key view; if it does, it is `policy-refused`.
- **P14. The pivot refusal** (S4 §6.9, §9 4c). The ten worker command kinds are refused before the mode checks, in every mode, after the `'message'` charge and with no token:
  - the eight pivot and control kinds of `VIEWER_COMMAND_KINDS` (`src/worker/viewerProtocol.ts:30–34`) are refused `pivot-refused`: `pivotOpen`, `pivotApply`, `pivotCommit`, `pivotFormulaStatus`, `pivotRelease`, `controlApply`, `controlCommit` and `controlCancel`;
  - `prepareColumns` is refused `server-prepared`;
  - `read` is refused `stream-no-cell-data`.

  The set is derived from `VIEWER_COMMAND_KINDS` by prefix, so a ninth pivot kind is refused too. No client sends them, so there is no client notice: the layout holds no pivot definition (S4 §6.4) to show it against. A slicer click already does nothing (4a).
- **P15. Client values selection.** The stream `ValuesDraftPort` keeps `{base: 'all' | 'none', flipped: Set<id>, blanks}`.
  - "Select all" under a search fetches every page of that search (≤ 10,000) and flips those ids.
  - The filter request sends the smaller of include and exclude, as `{kind: 'values', list, mode, ids: [[i0, i1], …], blanks}` with ascending, disjoint intervals: at most 4,096 intervals per request. Past that the client refuses locally: "Too many values selected for this viewing mode."
  - A column the user did not edit is sent as `{kind: 'committed'}`, which the server resolves to its committed rule. After a refetch the client holds placeholder rules built from `filterSummary`, which it re-sends as `committed`.
- **P16. Client menus.** Sort and filter menu items appear iff `sheetSort` (S4 §6.11: absent before 4c).
  - `FilterDropdown` gains `typedConditions?: boolean` (default `true`). When it is `false` (no `typedFilters` in the manifest):
    - the "Number filters" and "Date filters" choice is rendered `disabled`, titled "This filter isn't available in this viewing mode.";
    - the text/mixed panel offers no numeric operator.
  - Pacing: one `'sheet'` op in flight per viewer, queued in order. The value-list search is sent 250 ms after the last keystroke, and a newer one supersedes it.

## 4b follow-ups in 4c's area (`docs/validation/streaming-s4b.md`)

- **`blankCell`/`holdsContent` duplicated in `text-view.ts`:** fixed in Task 1, which exports them once from `snapshot.ts`; `key-view.ts` imports them.
- **Other painted-style sources** (row and column default styles; not searched by the 4b review): Task 1's `paintedFormats` fixture checks both in the text and key views.
- **The corpus oracle visits nothing to withhold:** Task 13 counts the corpus sheets with a withheld key cell and reports a 0. A drawing-bearing corpus file stays the owner's call.
- **Tile-side CF reads merge-tail values (F20), copy's pool admission estimate, and the `wireSelection` headroom:** not in 4c's area; left open.

## Review Focus

1. **A sort or filter answer lost to a resume, a cancel or a `stale-view`.** Expected:
   - the client re-fetches `sheet-layout` and `sheet-order`, and its order, mask and counts equal the server's;
   - no tile of the old `viewGen` is drawn, and no tile stays pending.

   Pinned in Tasks 8 and 9.
2. **A masked or region-styled cell in a key column.** Examples: a table style's `;;;` or `0,"K"` over General numbers (4b R1), a CF rule that writes `;;;` or a scaled format, a zero-height row, a zero-width column, a `;;;"****"` text section. Expected: a withheld cell sorts as blank, matches only a blanks filter and lists only as "(Blanks)"; a painted cell lists under its painted text (`1K`), never its value; a text rule never matches a masked raw string. Pinned in Tasks 1 and 13.
3. **Two value-list columns, and a list issued before a resize.**
   - Expected: a filter that keeps column B's committed values rule while editing column C applies both.
   - Ids of a list issued at an older `viewGen` are refused `stale-filter-values`, and the client re-fetches the list.

   Pinned in Tasks 8 and 10.
4. **A sort over a band with merges, a pivot, or a cell-anchored drawing.** Expected: the local refusal reaches the viewer with local's string, and `viewGen`, tiles and counts do not move. Pinned in Tasks 5 and 8.
5. **A full-height sort of a 1,048,576-row sheet.** Expected:
   - the delta arrives in ≤ 160 parts of ≤ 64 KiB and decodes, with no 4008 close;
   - the admission hold is charged once for two sids with the same view, and released at the last close.

   Pinned in Tasks 4, 8 and 9.

**Not in 4c:** pivot interaction (refused), a React-free sort UI for `officeview/core`, and per-viewer view state beyond the sid.

---

### Task 1: `sheetKeyView`, the engines under Bun, and the 4c fixtures

**Files:**
- Create: `server/src/sheet/key-view.ts`
- Modify:
  - `server/test/sheet-fixtures.ts` (append fixtures);
  - `server/src/sheet/snapshot.ts:33`, `server/src/sheet/text-view.ts:55–60`: export `blankCell` and `holdsContent` from `snapshot.ts`, and import them in `text-view.ts` and `key-view.ts` (the 4b Task 1 duplicate, so there is no third copy);
  - `src/components/xlsx/viewerControlsClient.ts:96`, `:130`, `:399`: an optional `context?: () => FilterContext` option. Its `nowMs` replaces `Date.now()` at `:399`. Absent, the local default is unchanged.
- Test: create `server/test/sheet-key-view.test.ts`

**Interfaces:**
- Consumes: `sheetTextView`, `TextView` (`server/src/sheet/text-view.ts`); `NO_SESSION_SIZES` (`src/xlsx/view/axisSizes.ts`).
- Produces:
  ```ts
  export interface KeyView { sheetIndex: number; sheet: XlsxSheet; blankAll: boolean; withheld: number; collapsedCols: ReadonlyArray<{ start: number; end: number }>; release(): void }
  export async function sheetKeyView(snap: SheetSnapshot, projected: XlsxSheet, o: { workbook: XlsxDocument; clock: SheetClock; budget: FormatBudget }): Promise<KeyView>
  export function keyDocument(base: XlsxDocument, views: ReadonlyMap<number, KeyView>): XlsxDocument
  ```
  - Under `'full'`, `sheet === projected`.
  - Otherwise `sheet` is the text view's sheet at the identity view, with P5's painted styles and its two blankings. The view's `scope` is released when the key view is built (the styles are resolved by then).
- Fixtures appended (each a `SheetFixture` builder, as in 4b):
  - `sortGrid()`: 40×6, header row; mixed numbers, `#,##0.0`, dates and text; hidden row 9; `ht="0"` row 12; a zero-width column E; merge-free.
  - `sortWithheld()`: `sortGrid` plus an unfilled rect over B20 (`COVER-B20` = 1), `;;;` in B21 (5), a CF `cellIs > 90 → ;;;` rule over B (B22 = 95), a wholeTable `;;;` table style over C30:C34 (General numbers), and a text cell `MASKED` formatted `;;;"****"` in D23.
  - `sortUnderShape()`: a `oneCellAnchor` picture at A15 over rows 15–16, inside the sort range.
  - `sortMerges()`: a merge B5:B6 inside the range.
  - `oneMillionSort()`: 1,048,576 rows in column A, numbers descending.
  - `hiddenPivotCache()`: a pivot cache over a hidden-row source with item `PIVOT-SECRET`, and a slicer cache with item `SLICER-SECRET`, plus a visible slicer drawing.
  - `filterAutoHidden()`: a saved autoFilter on A1:C20 that hides rows 4 and 7.
  - `paintedFormats()`: General `1234` in B2:B4 under a wholeTable DXF `0,"K"` (paints `1K`); `5678` in C2 under a CF rule writing `0.0,"K"` (paints `5.7K`); a column default style `;;;` on D (`<col style>`) and a row default style `;;;` on row 6 (`<row s customFormat="1">`) over cells with no own style.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)` on each).
  - `"engines import under Bun"`: `viewerControls` over `createObjectWorkbookSource(keyDocument(...))` sorts `sortGrid` by column B ascending, filters column D `contains "a"`, reads one `values` page and resolves one `range-detect`, all under `bun test`. If one cannot, STOP and report.
  - `"withheld cells are blank keys"`: in the `sortWithheld` key view, B20, B21, B22, C30:C34, every cell of column E and row 12 hold no value; D23 holds `'****'` with no number format; `sortGrid`'s other cells are unchanged.
  - `"painted styles"` (4b R1 on the key path): in the `paintedFormats` key view, `cellDisplayText` of B2 is `'1K'` and of C2 is `'5.7K'`. A `values` page for B through `viewerControls` lists exactly `['1K']` and never `'1234'`, and C lists `['5.7K']`. D's cells and row 6's cells are blank in both the text view and the key view (the 4b validation's "other painted-style sources"). If either default style is not painted by the tile, record it and keep the assertion on the tile's behaviour.
  - `"the key view ignores session sizes"`: the key view is the same object for any sizes the caller holds. It takes none, and a test spies that `sheetTextView` was called with `NO_SESSION_SIZES`.
  - `"unknown coverage"`: on `unresolvedAnchor`, `blankAll === true` and no cell holds a value.
  - `"full is the identity"`: under `'full'`, `sheet === projected` for every fixture above.
  - `"the injected clock"`: a `datePreset: 'today'` filter run through `viewerControls` with `context` from `clockEpochMs(clock)` keeps exactly the rows dated on `clock.day`. With `Date.now` stubbed to a different day, the result is the same.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-key-view.test.ts`. Expected: FAIL (no module).
- [ ] **Step 3: Implement** `key-view.ts`, the dedupe and the `context` option. Masked-string detection formats under the painted, CF-effective style (P5) and uses `String.prototype.includes`.
- [ ] **Step 4: Run green:** the same file; `cd server && bun test test/sheet-text-view.test.ts test/sheet-text-final.test.ts` (unchanged); `bunx vitest run xlsx-viewer-main-controls interactive-viewer-filter` (unmodified).
- [ ] **Step 5: Commit.**
  ```bash
  git commit -F msg -- server/src/sheet/key-view.ts server/test/sheet-fixtures.ts server/test/sheet-key-view.test.ts src/components/xlsx/viewerControlsClient.ts
  ```
  Message: `feat(stream-s4c): sheetKeyView — the one sheet every key answer reads`.

### Task 2: Wire codecs for 4c (React-free)

**Files:**
- Modify: `src/stream/sheet-wire.ts`
- Test: modify `tests/stream/sheet-wire.test.ts` (append only)

**Interfaces:**
- Produces, in `sheet-wire.ts`:
  - `SORT_FILTER_MAX_RULES = 64`, `FILTER_TEXT_UNITS = 256`, `VALUE_LIST_LIMITS = { entries: 10_000, page: 500, intervals: 4_096 }` and `ORDER_LIMITS = { parts: 160, bands: 1, rows: 1_048_576 }`.
  - The wire shapes:
    ```ts
    export interface WireRange { r0: number; r1: number; c0: number; c1: number; header: 0 | 1; totals: 0 | 1 }
    export interface WireSortSpec { range: WireRange; keys: Array<{ col: number; order: 'asc' | 'desc' }>; caseSensitive: boolean }
    export type WireTyped =
      | { kind: 'text'; op: TextOp; first: string; wildcards: boolean }
      | { kind: 'number'; op: NumberOp; first: number; second?: number }
      | { kind: 'date'; op: DateOp; first: CalendarDay; second?: CalendarDay }
      | { kind: 'datePreset'; preset: DatePreset }
    export type WireRule = WireTyped | { kind: 'compound'; connective: 'and' | 'or'; first: WireTyped; second: WireTyped }
      | { kind: 'values'; list: string; mode: 'include' | 'exclude'; ids: Array<[number, number]>; blanks: boolean } | { kind: 'committed' }
    export interface WireFilterSpec { range: WireRange; columns: Array<[number, WireRule]> }   // ascending, distinct columns
    export interface SheetOrderHead { sheetId: string; viewGen: number; counts: { visible: number; total: number }; sort: { range: WireRange; keys: WireSortSpec['keys'] } | null; filterSummary: Array<{ col: number; kind: string }> }
    export interface SheetOrder { head: SheetOrderHead; bands: Array<{ r0: number; r1: number; sourceAtOffset: Uint32Array }>; mask: { r0: number; r1: number; words: Uint32Array } | null }
    export interface ValuesPage { list: string; items: Array<{ id: number; text: string; selected: boolean; error: boolean }>; blanks: { present: boolean; selected: boolean }; total: number; truncated: boolean; cursor?: string; classification?: { kind: 'number' | 'date' | 'text' | 'empty'; hasNumeric: boolean } }
    export interface WireDetection { range: WireRange; source: 'table' | 'region' | 'used' | 'explicit'; header: 'metadata' | 'heuristic' | 'override'; overrideIgnored: boolean }
    ```
  - Request validators returning the typed args or `undefined` (⇒ `policy-refused`):
    - `sortArgs(m, sheetCount)`;
    - `filterArgs(m, sheetCount)`;
    - `valuesArgs(m, sheetCount)`;
    - `detectArgs(m, sheetCount)`: `{kind: 'sort' | 'filter', selection: WireSelection, override?: {header: 0 | 1}}`;
    - `orderArgs(m, sheetCount)`.

    These check P4's shape rules only. Hidden targets, coverage, typed kinds and ids are the server's.
  - `encodeSheetOrder(o: SheetOrder): object[]` (P7 parts) and `decodeSheetOrder(parts: unknown[]): SheetOrder` (strict: a non-permutation, an overlap or a cap breach throws `InvalidOperand`).
  - `decodeValuesPage(m): ValuesPage` and `decodeDetection(m): WireDetection`.

- [ ] **Step 1: Write the failing tests** (append):
  - `"sort and filter args"`: a 3-key sort and a 4-column filter round-trip. Each of these is refused:
    - 65 keys;
    - a key column outside the range;
    - `header: 2`;
    - a body of 0 rows;
    - a 257-unit `first`;
    - an `ids` interval with `i1 < i0`;
    - overlapping intervals;
    - 4,097 intervals;
    - duplicate columns;
    - a `compound` with a `values` child;
    - an unknown `kind`;
    - a NaN `first`.
  - `"values args"`: a first page and a cursor page are accepted. Both forms at once, and a 257-unit search, are refused.
  - `"order parts"`: a 1,048,576-row reverse band encodes in ≤ 160 parts, each ≤ 65,536 B with a 10-digit reqId added, and decodes equal. A part dropped, a duplicate entry, `seq` out of order and 161 parts each throw.
  - `"values page"`: an item with `flags 3` decodes `{selected: true, error: true}`. An item `id` outside `[0, total)` throws.
- [ ] **Step 2: Run** `bunx vitest run tests/stream/sheet-wire.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.** Only type imports from `src/xlsx/**`. base64 goes through `btoa`/`atob` over a `Uint8Array` view of the LE words.
- [ ] **Step 4: Run green:** the same file, and `npx tsc --noEmit -p .`.
- [ ] **Step 5: Commit** `src/stream/sheet-wire.ts tests/stream/sheet-wire.test.ts`: `feat(stream-s4c): sort, filter, value-list, range-detect and order codecs`.

### Task 3: The `sheetSort` manifest key, op gating, refusal codes, and the pivot refusal

**Files:**
- Modify:
  - `server/src/ws.ts`: `RefusalCode` at `:65`, `SHEET_VIEW_OPS` at `:95`, the manifest at `:546`, the dispatch at `:1778–1800`;
  - `src/stream/manifest.ts`: `sheetSort?: 1`, `typedFilters?: true`;
  - `server/test/sheet-view-handshake.test.ts:56`: the one widening Open item 1 permits, `[...TILES_KEYS, 'sheetView', 'sheetSort'].sort()`. No other line of it changes.
- Test: create `server/test/sheet-sort-handshake.test.ts`; modify `tests/stream/manifest.test.ts` (append)

**Interfaces:**
- Produces:
  - the manifest of every `sheetView` session gains `sheetSort: 1`, and `typedFilters: true` under `'full'` (P1);
  - `SHEET_SORT_OPS = new Set(['sheet-sort', 'sheet-filter', 'filter-values', 'range-detect', 'sheet-order'])`, each refused `unsupported-op` without `sheetView`, **uncharged** (4b leaves off-mode fit and resize uncharged, `ws.ts:1784`);
  - `RefusalCode` gains `'stale-filter-values' | 'filter-condition-refused' | 'pivot-refused' | 'server-prepared' | 'stream-no-cell-data' | 'view-refused' | 'cancelled'`;
  - `WORKER_KIND_REFUSALS: ReadonlyMap<string, RefusalCode>`, built from `VIEWER_COMMAND_KINDS` (P14). It is checked right after `ready` in the dispatch, before the mode checks.
  - Client: `Manifest.sheetSort?: 1` is valid only with `sheetView: 1`, and `typedFilters?: true` only with `sheetSort`; any other value rejects the manifest.

- [ ] **Step 0: Check Open item 1.** If the owner has not permitted the widening at `sheet-view-handshake.test.ts:56`, STOP and report.
- [ ] **Step 1: Write the failing server tests:**
  - `"sheetSort on every sheet-view manifest"`:
    - `['sheet-view@1']` on XLSX under visible-only gives `sheetSort === 1` and no `typedFilters`;
    - under `'full'` it gives `typedFilters === true`;
    - `['sheet-tiles@1']` gives 4a's manifest key for key, and units mode gives Stage 3's;
    - a DOCX gets no new keys.
  - `"the 4c ops need sheetView"`: on a `sheet-tiles@1` session and in units mode, each of the five ops is refused `unsupported-op`. Forty of them in a minute are all `unsupported-op`, never `rate-limited`, and none closes 4008.
  - `"4b layout bytes unchanged"`: a `sheet-view@1` session's `sheet-layout` of `grid500x100` equals, byte for byte, the bytes pasted from the base `fb624989`.
  - `"pivot and control kinds are refused"`: on a tiles session and on a units session, each of the eight kinds is refused `pivot-refused`, `prepareColumns` is refused `server-prepared` and `read` is refused `stream-no-cell-data`. None needs a token, none closes, and each charges `'message'` once.
- [ ] **Step 2: Write the failing client tests** (append to `manifest.test.ts`):
  - `sheetSort: 1` with `sheetView` passes;
  - `sheetSort: 1` without `sheetView` rejects;
  - `typedFilters: true` without `sheetSort` rejects;
  - `sheetSort: 2` rejects.
- [ ] **Step 3: Run** `cd server && bun test test/sheet-sort-handshake.test.ts` and `bunx vitest run tests/stream/manifest.test.ts`. Expected: FAIL.
- [ ] **Step 4: Implement.** Record `bun run gates:core` before the client edit and after it.
- [ ] **Step 5: Run green:** the same commands, then `cd server && bun test test/sheet-view-handshake.test.ts test/sheet-handshake.test.ts test/xlsx-units-baseline.test.ts` (only the permitted line changed).
- [ ] **Step 6: Commit** `server/src/ws.ts src/stream/manifest.ts server/test/sheet-view-handshake.test.ts server/test/sheet-sort-handshake.test.ts tests/stream/manifest.test.ts`: `feat(stream-s4c): sheetSort manifest key, 4c refusal codes, pivot and worker-kind refusals`.

### Task 4: The canonical view, `viewStateHash`, session view state, admission and the audit tables

**Files:**
- Create: `server/src/sheet/view-spec.ts`
- Modify:
  - `server/src/sheet/view-state.ts`: `SessionViews` sort and filter state;
  - `server/src/cache.ts:51–57`: `SheetViewState`, `viewStateHash`; `:107–122`: `DocumentLedger.hold`/`release`;
  - `server/src/audit-store.ts:32`, `:77`: `SheetViewRecord.spec`, the `sheet_ops` table and `putSheetOp`/`sheetOps(sid)`;
  - `server/src/audit.ts:106`: `recordSheetOp`.
- Test: create `server/test/sheet-view-spec.test.ts`; modify `server/test/cache.test.ts` and `server/test/audit-store.test.ts` (append)

**Interfaces:**
- Produces, in `view-spec.ts`:
  ```ts
  export type ResolvedRule = Exclude<WireRule, { kind: 'values' } | { kind: 'committed' }> | { kind: 'values'; list: { range: WireRange; col: number }; positions: number[]; mode: 'include' | 'exclude'; blanks: boolean }
  export interface ResolvedFilter { range: WireRange; columns: Array<[number, ResolvedRule]> }
  export interface CanonicalView { sort: WireSortSpec | null; sortFilter: ResolvedFilter | null; filter: ResolvedFilter | null }
  export const IDENTITY_VIEW: CanonicalView
  export function canonicalView(v: CanonicalView): CanonicalView      // sorted columns, sortFilter null when sort is
  export function canonicalViewHash(v: CanonicalView): string        // sha256 of canonical JSON
  export function isIdentityView(v: CanonicalView): boolean
  export function orderBytes(v: CanonicalView): number               // P11: 8 × body rows of the sort/filter ranges + 160 KiB; 0 for identity
  ```
- `SheetViewState` becomes `{ view: CanonicalView; sizes: WireSizes; authorHiddenTier: 'none' | 'c1'; clock: SheetClock }`.
  - `viewStateHash` hashes `{sort: v.sort && {spec: v.sort, under: v.sortFilter}, filters: v.filter ? [v.filter] : [], sizes, authorHiddenTier, clock}`. The identity and sizes-only hashes are byte-identical to 4b's: `sort: null`, `filters: []`, tier `'none'`.
  - The tier is `'c1'` iff `visible-only` and a filter is set.
- `SessionViews` gains:
  - `view(sheetIndex): CanonicalView`;
  - `commit(sheetIndex, next: CanonicalView, hold: () => boolean): number | 'view-state-limit'`. It enforces `MAX_ORDERED_SHEETS` = 2. `hold` charges the admission and is skipped for identity. The result is the new `viewGen`.
  - `orderedSheets(): number`.
  - `dispose()` also releases holds.
- `DocumentLedger.hold(key: string, bytes: number): boolean` puts a pinned entry, or increments its count. `release(key: string): void` decrements, unpinning and deleting at 0.
- `SheetOpRecord { sid; sheetId; kind: 'sort' | 'filter' | 'filter-values' | 'range-detect'; viewGen; range: WireRange; rules: Array<[number, string]>; outcome: string; at }`.

- [ ] **Step 1: Write the failing tests:**
  - `"identity hashes unchanged"` (append to `cache.test.ts`): `viewStateHash` of the identity view, and of a sizes-only view, equals the 4b hex pasted from base `fb624989`.
  - `"sortFilter keys the hash"`: `{sort S, sortFilter F1, filter F1}` and `{sort S, sortFilter null, filter F1}` hash differently. Column order in the input does not change the hash, and `sortFilter` is dropped when `sort` is `null`.
  - `"two ordered sheets per sid"`: commits on sheets 0 and 1 succeed and sheet 2 gets `'view-state-limit'`. Clearing sheet 0 back to identity frees a slot, and resizes on sheet 2 still work.
  - `"admission"`: with `new DocumentLedger(drop, 10 << 20)`, a held 8 MiB view evicts an unpinned 4 MiB parse (`drop` called). A second hold of the same key charges nothing more. A third distinct 8 MiB hold returns false. Releasing both holders frees it.
  - `"sheet_ops"` (append to `audit-store.test.ts`): a row round-trips, its columns hold no operand text, and a pre-4c database opens and gains the table.
  - `"sheet_views spec"`: a 4b record reads `sortFilter: null, filter: null`, and a 4c record round-trips its values positions and text operands.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-view-spec.test.ts test/cache.test.ts test/audit-store.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same files, then `cd server && bun test test/sheet-resize.test.ts test/trace-tiles.test.ts` (unchanged).
- [ ] **Step 5: Commit** the five source files and three tests: `feat(stream-s4c): canonical sheet view, view-state hash, ordered-sheet cap and admission, sheet_ops`.

### Task 5: The `sheet-view` pool job: replay, sort, filter, value lists and range detection over the key view

**Files:**
- Create: `server/src/sheet/view-ops.ts`
- Modify:
  - `server/src/render.ts`: `DocumentHost.sheetView`, key-view and derived-view caches, released on `drop`;
  - `server/src/worker/render-worker.ts:52`, `:112`: a `sheet-view` message;
  - `server/src/pool.ts:403`: `Pool.sheetView`.
- Test: create `server/test/sheet-view-ops.test.ts`

**Interfaces:**
- Consumes:
  - `sheetKeyView` and `keyDocument` (Task 1);
  - `CanonicalView` (Task 4);
  - `viewerControls` with `context` (`viewerControlsClient.ts`);
  - `ControlsClient.sort`/`filter`/`values`/`resolve` (`controlsClient.ts:38–47`).
- Produces:
  ```ts
  export type SheetViewJob =
    | { op: 'derive'; sheetIndex: number; view: CanonicalView }
    | { op: 'sort'; sheetIndex: number; base: CanonicalView; sort: WireSortSpec | null }
    | { op: 'filter'; sheetIndex: number; base: CanonicalView; filter: ResolvedFilter | null }
    | { op: 'values'; sheetIndex: number; view: CanonicalView; range: WireRange; col: number; search?: { text: string; wildcards: boolean }; offset: number }
    | { op: 'detect'; sheetIndex: number; kind: 'sort' | 'filter'; selection: WireSelection; override?: { header: 0 | 1 } }
  export type SheetViewReply =
    | { op: 'view'; view: CanonicalView; order: SheetOrder }                 // derive, or a committed sort/filter
    | { op: 'view'; refused: WireOperationRefusal }
    | { op: 'values'; page: Omit<ValuesPage, 'list' | 'cursor'>; next?: number; count: number }   // count: merged entries in the unsearched list
    | { op: 'detect'; detection: WireDetection } | { op: 'detect'; refused: WireOperationRefusal }
  export interface WireOperationRefusal { reason: string; range?: SourceRect; count?: number; at?: [number, number] }
  ```
  - `DocumentHost.sheetView(docKey, clock, contentPolicy, job)`:
    - one `viewerControls` port per `(sheetIndex, clock)` over `keyDocument`, at most 4 per document (LRU). Its `view`, `snapshot` and `publish` read and write the job's current base: the worker runs one job per slot at a time;
    - derived views per P3, at most 8 per document (LRU).
  - **C1 under `visible-only`:** the port's `publish` (and replay) replaces a filter's visibility with `matches ∧ ¬collapsed(source row)` over the authored collapsed rows before it is kept, applied to the body rows of the filter range only. Under `'full'` the local mask is kept.
  - `values`:
    - merges local entries by full display text (P6), keeps 500 entries per page (`next` is the next offset, or absent);
    - `selected` comes from `view.filter`'s rule for `(range, col)`;
    - expands a values rule's positions back into local ordinals through the same merge, then drives a fresh draft (`change` … `candidate`) to build the local `ValuesFilter`.
  - `Pool.sheetView(sid, job & {docRef, contentPolicy, clock}, bytesEstimate, signal)`, with the same affinity, admission, deadline and recycle as `sheetText`. Its `bytesEstimate` is `orderBytes(next view) + 16 MiB`.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)`). The local oracle in each is `viewerControls` run in the test over a copy of the fixture with its withheld cells physically deleted (`sortWithheld` minus B20, B21, B22, C30:C34, column E's cells and row 12's cells, with D23 replaced by `'****'`). It is built apart from the server inputs.
  - `"sort parity"`: for keys B asc, B desc, D asc case-sensitive, and (C asc, B desc), the derived `sourceAtOffset` equals the oracle's order. Under `'full'`, the oracle is plain local mode on the projected `sortWithheld`.
  - `"filter parity and C1"`: on `filterAutoHidden` under visible-only:
    - `text contains ""`, `values all` and the cleared filter each leave rows 4 and 7 hidden;
    - every other row's visibility equals the oracle's;
    - under `'full'`, clearing the filter shows rows 4 and 7, as locally.
  - `"withheld cells sort and list as blank"`: sorting B ascending puts B20, B21 and B22 among the blanks at the end, as the oracle does. `values` for B lists no `1`, `5` or `95`, and `blanks.present === true`. `values` for D lists `****` and never `MASKED`. A text rule `contains "MASK"` on D matches nothing.
  - `"replay equals incremental"`: commit filter F1 (`values`), sort S, then filter F2 (`text`), then filter F3 (`blanks`). After each commit, `derive` of the canonical view deep-equals the incremental result; also on `sortUnderShape` with S outside the drawing's band. `derive({S, null, F2})` differs on `sortGrid` (hidden-row pins). If any replay is refused where the incremental op was accepted, STOP and report (P3's assumption failed).
  - `"the spec moves as P3 says"`: after F1 then S, the view is `{S, F1, F1}`; after F2, `{S, F1, F2}`; after clearing the sort, `{null, null, F2}`.
  - `"display-text merge"`: B holding `1200` and `1204`, both formatted `0.0,"K"`, lists one `1.2K` entry. Selecting it keeps both rows, and the oracle with values {1200, 1204} gives the same rows (S4 §6.12 item 7).
  - `"two values columns"`: rules on B and D together equal the oracle's rows.
  - `"local guards refuse"`: a sort over `sortUnderShape`'s band gives `{reason: 'drawing-impact'}` with no `drawingId`, `sortMerges` gives `{reason: 'merge', range}`, and neither returns a view.
  - `"range detection over servable occupancy"`: on `sortWithheld`, a filter detect from C31 gives the oracle's range over the deleted fixture.
  - `"a million rows"`: `oneMillionSort` sorted A ascending gives `sourceAtOffset[0] === 1048575`, and the job's heap high-water stays under `orderBytes + 32 MiB`.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-view-ops.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.** `SheetOrder.bands` comes from the committed `RowOrder.bands`, and `mask` from the committed `CompactRowVisibility` words.
- [ ] **Step 4: Run green:** the same file; `cd server && bun test test/pool.test.ts test/sheet-fit.test.ts` (unchanged).
- [ ] **Step 5: Commit** `server/src/sheet/view-ops.ts server/src/render.ts server/src/worker/render-worker.ts server/src/pool.ts server/test/sheet-view-ops.test.ts`: `feat(stream-s4c): sheet-view pool job — local engine over the key view, C1, replay`.

### Task 6: Tiles, text ops and the tracer under the view

**Files:**
- Modify:
  - `server/src/sheet/tile.ts:16`, `:72`: `TileJob.view?`, validated by `tileJobProblem`;
  - `server/src/render.ts:488`, `:527`: `renderTile` and `sizedGeometry` with order and visibility (cache key `+ canonicalViewHash`);
  - `server/src/sheet/text-ops.ts:32–42`, `:73`: `view?` on every `SheetTextJob`, and `textInputs(view, source, derived?)`;
  - `server/src/sheet/text-view.ts`: the current-view coverage condition (P9);
  - `server/src/detect/trace.ts:210`: the delivery's spec view goes into each rebuilt tile.
- Test: create `server/test/sheet-view-text.test.ts`; modify `server/test/trace-tiles.test.ts` (append)

**Interfaces:**
- Consumes: the derived-view cache (Task 5).
- Produces:
  - `TileJob.view?: CanonicalView`. Absent is identity. With a view, the tile paints `snap.sheet` over `buildSheetGeometry(snap.sheet, {sizes, order, visibility})`.
  - `textInputs(view, source, derived?: { order: RowOrder; visibility: RowVisibility })`: without `derived` it is 4b's identity behaviour, unchanged.
  - `sheetTextView(snap, projected, sizes, o & { order?: RowOrder; visibility?: RowVisibility })`.
  - Text-view cache key `+ canonicalViewHash(view)`.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)`):
  - `"tiles follow the view"`: tile (0,0) of `sortGrid` sorted B desc equals `renderSheetViewport` of the snapshot sheet over `buildSheetGeometry(sheet, {order, visibility})`. Two sids with equal views share one `pool.tile` call, and a different filter gives a second call.
  - `"copy, find and a11y follow the view"`: after filter D `contains "a"`:
    - a select-all copy equals local `serializeSelection` over the deleted fixture with the oracle's order and mask;
    - find for a term only in a filtered-out row gives no hit (Open item 2);
    - a11y of a window lists no filtered-out row.
  - `"fit follows the view"`: a column fit after the filter equals local `autoFitSize` with the oracle's geometry.
  - `"a sort that slides a cell under a fixed drawing blanks its text"`: on a fixture with an `absoluteAnchor` rect over display rows 3–4 and `UNDER` in source row 30, find for `UNDER` hits before the sort and not after one that brings row 30 to display row 3.
  - (append to `trace-tiles.test.ts`) `"a capture of a sorted view traces"`: deliveries at the sorted `viewGen` trace to their sid.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-view-text.test.ts test/trace-tiles.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same files, then `cd server && bun test test/sheet-find.test.ts test/sheet-copy.test.ts test/sheet-a11y.test.ts test/sheet-fit.test.ts test/sheet-text-audit.test.ts` (unchanged).
- [ ] **Step 5: Commit** the five sources and two tests: `feat(stream-s4c): tiles, find, copy, a11y, fit and the tracer under the session's view`.

### Task 7: `filter-values` and `range-detect` on the socket

**Files:**
- Create: `server/src/sheet/value-lists.ts`
- Modify: `server/src/ws.ts` (`onFilterValues`, `onRangeDetect`)
- Test: create `server/test/sheet-filter-values.test.ts`

**Interfaces:**
- Produces, in `value-lists.ts`:
  ```ts
  export class ListLedger {                                   // per sid; at most 16 lists (LRU)
    issue(key: { sheetIndex: number; range: WireRange; col: number; viewGen: number }, count: number): string   // the same id for the same key
    get(id: string): { sheetIndex: number; range: WireRange; col: number; viewGen: number; count: number } | undefined
    resolve(rule: Extract<WireRule, { kind: 'values' }>, at: { sheetIndex: number; range: WireRange; col: number; viewGen: number }): ResolvedRule | 'stale-filter-values'
  }
  ```
- `onFilterValues`, in order:
  1. the token;
  2. `valuesArgs`, else `policy-refused`;
  3. for a first page, `'sheet'` rate and busy; for a cursor page, nothing more;
  4. `stale-view`;
  5. `hidden-target` for a collapsed column;
  6. the pool `values` job;
  7. `ListLedger.issue`, the `sheet_ops` row (first pages only), then `filter-values-result` (≤ 48 KiB; a page over it is split by entries, and the cursor continues from there).

  The cursor is `${list}:${offset}`; an unknown list or a bad offset is `stale-filter-values`.
- `onRangeDetect`: token, `detectArgs`, `'sheet'` rate and busy, the pool `detect` job, the `sheet_ops` row, then `range-detect-result {detection}` or `view-refused`.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)`):
  - `"lists are display text, withheld cells only as blanks"`: on `sortWithheld` B:
    - the items equal the oracle's merged labels over the deleted fixture, in order;
    - `blanks.present` is true;
    - no item text is `1`, `5`, `95` or any C30:C34 value;
    - D lists `****` and no `MASKED`;
    - on `paintedFormats`, B lists exactly `1K`, and no item text anywhere is `1234` or `5678`.
  - `"rotated box and unknown coverage"` (S4 §9 4c):
    - on `rotatedBox`, every cell meeting `x 40..60, y 30..50` is absent from its column's list, and the list has `blanks.present`;
    - on `unresolvedAnchor`, every column's list is empty with `blanks.present`.
  - `"off-viewport covered cell"`: `offViewportCovered` column Z lists no `OFFVIEW-Z5000`, with no tile ever fetched.
  - `"paging and search"`: 1,200 distinct texts give pages of 500, 500 and 200, each ≤ 48 KiB, with stable ids. A search for `"7"` gives full-list ids. 12,000 distinct texts give `truncated: true` and 10,000 entries in total.
  - `"rates"`: the first page charges `'sheet'` and cursor pages do not (thirty cursor pages within the minute all succeed). The 31st first page is `rate-limited`.
  - `"values on without copy"`: `copy: false` answers (S4 §6.6 availability).
  - `"ids are sid- and column-bound"`: `ListLedger.resolve` gives `stale-filter-values` for an id of another sid's ledger, for a list of column C resolved at column D, for position 99 of a 3-entry list, for a descending interval, and for a list issued at `viewGen 0` resolved at `viewGen 1`. A 17th issued key evicts the first.
  - `"range-detect"`: equals Task 5's detection, and it is `rate-limited` at the 31st `'sheet'` op.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-filter-values.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same file, then `cd server && bun test test/sheet-view-ops.test.ts`.
- [ ] **Step 5: Commit** `server/src/sheet/value-lists.ts server/src/ws.ts server/test/sheet-filter-values.test.ts`: `feat(stream-s4c): filter-values — display-text lists with sid-bound ids; range-detect`.

### Task 8: `sheet-sort`, `sheet-filter`, `sheet-order` and cancel on the socket

**Files:**
- Modify:
  - `server/src/ws.ts`: `onSheetSort`, `onSheetFilter`, `onSheetOrder`; `onSheetCancel` at `:957` (sort and filter targets); `tilesOp`/`deliverTile` pass `sd.views.view(i)`; the text ops pass it too; `sessionLayout` callers pass the sort and filter summary;
  - `server/src/sheet/layout.ts:132`: `sessionLayout(base, viewGen, sizes, summary?)`.
- Test: create `server/test/sheet-sort-filter.test.ts`

**Interfaces:**
- Consumes: `sortArgs`, `filterArgs`, `orderArgs`, `encodeSheetOrder` (Task 2); `SessionViews.commit`, `DocumentLedger.hold`, `recordSheetOp`, `recordSheetView` (Task 4); `Pool.sheetView` (Task 5).
- Produces, `onSheetSort` and `onSheetFilter`, in order:
  1. the token (`textOp`);
  2. `*Args`, else `policy-refused` (no row, no charge);
  3. `'sheet'` rate, then busy (`rate-limited`; set before the next await, and cleared in `finally` only while this socket owns the sid);
  4. P4's checks against the base layout (`baseLayout`) and the issued lists (Task 7's `ListLedger.resolve`). A `committed` rule becomes the column's committed `ResolvedRule`, or `policy-refused` when there is none;
  5. the fingerprint and the document ledger;
  6. the pool job;
  7. re-check `viewGen` and cancellation, synchronously, then `views.commit` (with the admission hold of the new view and the release of the old) and stop the sid's tile batches of the sheet (4b P14);
  8. the `sheet_ops` row, then the `sheet-order` parts.
- An engine refusal is `view-refused` (P13) and moves nothing. Every refusal after step 2 writes its `sheet_ops` row with the code.
- `onSheetOrder`: token, `orderArgs`, `'unit'` rate, then `derive` of the current view and its parts.
- `sheet-layout` to a `sheetSort` session carries `sort` and `filterSummary` of the current view, only when they are not the identity view. A `sheetView`-only session's layout is 4b's.

- [ ] **Step 1: Write the failing tests** (`}, 30_000)`):
  - `"sort round trip"`: on `sortGrid`, `sheet-sort` B asc at `viewGen 0` gives `sheet-order` parts. Their decoded head says `viewGen 1` and `counts` equal the oracle's, and their band equals Task 5's. The next `sheet-layout` has `viewGen 1` and `sort.keys` `[{col: 1, order: 'asc'}]`. `sheet-order` re-answers the same parts.
  - `"authorization refuses before anything moves"`: each of these is refused, and a following `sheet-order` still says `viewGen 0` with no `sheet_ops` row of outcome `ok`:
    - `viewGen 5` (`stale-view`);
    - a key on collapsed column E (`hidden-target`);
    - range edge `c1` = E (`hidden-target`);
    - an `unresolvedAnchor` sort under visible-only (`coverage-unknown`);
    - a `number gt` rule under visible-only (`filter-condition-refused`);
    - a `compound(text, datePreset)` (same);
    - a values rule with another sid's list, with a list for column C used on D, with position 99 of a 3-entry list, and with a list issued at `viewGen 0` sent at `viewGen 1` (each `stale-filter-values`);
    - a third ordered sheet (`view-state-limit`).
  - `"full allows typed conditions"`: under `'full'`, `number gt 50` on B keeps the oracle's rows.
  - `"tautology filters"` (S4 §6.6): `text neq ""`, under visible-only and under `'full'`, on `filterAutoHidden` over A1:C20. Under visible-only, rows 4 and 7 stay hidden and the layout's collapsed runs are unchanged. `number gte -1e308` is refused `filter-condition-refused` under visible-only.
  - `"T6 pair"`: sids X and Y on one document, X filtered D `contains "a"` and Y unfiltered:
    - X's tiles and Y's tiles of the same coordinate come from different canonical keys;
    - Y receives no `sheet-order` part;
    - each socket's captured frames carry only its own `viewGen` sequence.
  - `"the sheet class"`: thirty sorts in a minute are accepted, the 31st is `rate-limited`, and a sort sent while one is in flight is `rate-limited`. None closes 4008.
  - `"cancel"`: a `sheet-cancel` sent before the pool job settles (an injected slow pool) gives `refused cancelled` and `viewGen` unchanged. A cancel after the answer changes nothing.
  - `"a million-row sort"`: `oneMillionSort` gives ≤ 160 parts, each ≤ 64 KiB; the ledger holds `orderBytes` once for two sids with the same view and 0 after both close.
  - `"audit"`: after a sort and a refused filter, `store.sheetOps(sid)` has `{kind: 'sort', rules: [[1, 'asc']], outcome: 'ok'}` and `{kind: 'filter', outcome: 'filter-condition-refused'}`, and no row holds the operand text. `sheet_views` for `viewGen 1` holds the sort.
- [ ] **Step 2: Run** `cd server && bun test test/sheet-sort-filter.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** the same file and `test/sheet-filter-values.test.ts`, then `cd server && bun test` (report the count).
- [ ] **Step 5: Commit** `server/src/ws.ts server/src/sheet/layout.ts server/test/sheet-sort-filter.test.ts`: `feat(stream-s4c): sheet-sort and sheet-filter — target authorization, C1, typed refusal, order delta, cancel`.

### Task 9: Client session and `StreamWorkbook`: sort, filter, value lists, detection and the order

**Files:**
- Modify:
  - `src/stream/remote/text-wire.ts`: the 4c request and answer types;
  - `src/stream/remote/session.ts:505–515`: resume policy per P8; `sheet-order` part assembly;
  - `src/stream/remote/remote-doc.ts:57–63`, `:313`: `RemoteSheets.sortOps?`;
  - `src/components/xlsx/streamWorkbook.ts`: `:183` `viewMoved` keeps the order across a resize; `:293`, `:330–332`: `view`, `sort` and `filter`; new `values`, `detect`, `refreshOrder`.
- Test: create `tests/stream/remote-sheet-sort.test.ts`, `tests/xlsx-stream-sort.test.ts`

**Interfaces:**
- Produces:
  ```ts
  export interface RemoteSortOps {
    sort(sheetId: string, viewGen: number, spec: WireSortSpec | null, signal?: AbortSignal): Promise<SheetOrder>
    filter(sheetId: string, viewGen: number, spec: WireFilterSpec | null, signal?: AbortSignal): Promise<SheetOrder>
    values(req: { sheetId: string; viewGen: number; range: WireRange; col: number; search?: { text: string; wildcards: boolean } } | { sheetId: string; viewGen: number; cursor: string }, signal?: AbortSignal): Promise<ValuesPage>
    detect(sheetId: string, kind: 'sort' | 'filter', selection: WireSelection, override?: { header: 0 | 1 }, signal?: AbortSignal): Promise<WireDetection>
    order(sheetId: string, signal?: AbortSignal): Promise<SheetOrder>
  }
  ```
  - `RemoteSheets.sortOps` is present iff `manifest.sheetSort === 1`. An abort sends `sheet-cancel {target}`. A refusal rejects `StreamSessionError` with its code, and `view-refused` carries `refusal: OperationRefusal` rebuilt with local message strings.
  - `StreamWorkbook` gains:
    - `sortOps: boolean`;
    - `typedFilters: boolean`;
    - `applyOrder(o: SheetOrder): void`, which ignores a `viewGen` at or below the held one, and otherwise sets `view(sheetId)` to `{revision: viewGen, order: createRowOrder(bands), visibility: CompactRowVisibility(mask), sort, filter}` and notifies;
    - `refreshOrder(sheetId)`, which fetches the layout and then `order` when the layout's `sort`/`filterSummary` is not the identity view.
  - `sort(owner, spec, signal, progress)` and `filter(owner, spec, keySets, signal, progress)` implement P8's progress and return `RemoteCommit`. `keySets` is ignored on a stream; values rules carry P15's wire form through `streamControls` (Task 10).
  - A resize (`applyView`) bumps `revision` and keeps the order and mask.
  - Without `sortOps` (a 4a or 4b server), `sort`, `filter`, `values` and `detect` still reject `stream-no-cell-data`, as they do today (`streamWorkbook.ts:331–332`).

- [ ] **Step 1: Write the failing tests:**
  - (`tests/stream/remote-sheet-sort.test.ts`, fake socket):
    - `"sortOps present only with sheetSort"`;
    - `"order parts assemble"`: 90 parts out of order are refused, and in order they assemble;
    - `"resume: sort and filter reject session-rotated, values and detect are re-issued"` (P8);
    - `"abort sends sheet-cancel"`;
    - `"view-refused carries the local refusal"`: `{reason: 'merge', range}` rejects with `refusal.reason === 'merge'`;
    - `"hostile answers"`: an order whose band is not a permutation, a values item `id` ≥ `total`, or a detection outside the grid each rejects `policy-refused`, and the session stays open.
  - (`tests/xlsx-stream-sort.test.ts`):
    - `"applyOrder moves the view"`: `view('0').order.displayToSource(0) === 39` after a reversed band, and `identity('0').viewGen === 1`;
    - `"client geometry equals the server's under the view"`: `createSheetMetrics` over the layout with the applied order and mask gives `rows.start(i)` equal to `buildSheetGeometry(sheet, {order, visibility})` for i in [0, 3, 9, 12, 39, 1048575];
    - `"a resize keeps the order"`;
    - `"no sortOps, no sort"`: without `sheetSort`, `sort` rejects `StreamNoCellData` with `method === 'sort'`;
    - `"a lost sort answer refetches"` (Review Focus 1): after `session-rotated`, `layout` then `order` are called and adopted.
- [ ] **Step 2: Run** `bunx vitest run tests/stream/remote-sheet-sort.test.ts tests/xlsx-stream-sort.test.ts`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:** `bunx vitest run tests/stream tests/xlsx-stream-workbook.test.ts tests/xlsx-stream-text.test.ts` (existing files unedited), and `bun run gates:core` (report the delta).
- [ ] **Step 5: Commit** the four sources and two tests: `feat(stream-s4c): client sort, filter, value-list and detect ops; order deltas on StreamWorkbook`.

### Task 10: The viewer: stream controls, menus and the typed-condition greying

**Files:**
- Create: `src/components/xlsx/streamControls.ts`
- Modify:
  - `src/components/xlsx/InteractiveSheetDoc.tsx:2133–2178`: a `sortOps` stream uses `streamControls` in place of `viewerControls`; filter buttons and the sort and filter menu items are mounted iff `sortOps`;
  - `src/components/xlsx/FilterDropdown.tsx:128`: `typedConditions` (P16).
- Test: create `tests/xlsx-stream-controls.test.tsx`; modify `tests/stream/remote-sheet-view.test.tsx` (append)

**Interfaces:**
- Consumes: `StreamWorkbook.sortOps`, `typedFilters`, `sort`, `filter`, `values`, `detect` (Task 9).
- Produces:
  ```ts
  export function streamControls(o: { stream: StreamWorkbook; sheetId: string; snapshot(): ControlSnapshot; view(): SheetView; publish(view: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>): ControlSnapshot; pace?: { setTimer?: typeof setTimeout } }): ViewerControls
  ```
  - `client.resolve` calls `detect`.
  - `client.classify` returns the classification of the last first page for `(range, col)`, else `{kind: 'loading', hasNumeric: false}`.
  - `client.values` returns a `ValuesDraftPort` over `values` pages (P15). `read` is a 250 ms debounced search, and `inspect` returns the label.
  - `client.filter` and `client.sort` build `WireFilterSpec` and `WireSortSpec`: unchanged columns become `committed`, and the draft gives the values rule.
  - One `'sheet'` op is in flight per viewer, the rest queued.

- [ ] **Step 1: Write the failing tests** (fake `RemoteSortOps` over Task 5's server answers, recorded):
  - `"sort from the column menu"`: "Sort ascending" on column B sends `sort({range, keys: [{col: 1, order: 'asc'}]})` at `viewGen 0`. The grid's first data row shows source row 39 (`aria-rowindex`), and the status says 'View-only sort applied; formulas unchanged.'
  - `"filter by values"`: the dropdown on D lists the fake's items. Deselecting two and pressing OK sends `{kind: 'values', mode: 'exclude', ids: [[1, 2]]}`, with every other filtered column sent as `committed`.
  - `"typed conditions greyed"`: with `typedFilters` absent, the 'Number filters' option is `disabled` with title "This filter isn't available in this viewing mode.", and the text/mixed panel offers no numeric operator. With `typedFilters: true`, it is enabled.
  - `"refusals speak local strings"`: `view-refused {reason: 'merge'}` shows the local merge refusal text, and `stale-filter-values` re-fetches the list and shows 'Values changed; review the list and try again.'
  - `"pacing"`: opening a dropdown sends `detect` then `values`, never both in flight. Typing 'abc' in the search sends one search request, 250 ms after the last key.
  - `"author-hidden notice"`: a filter on `filterAutoHidden` shows "Some rows are hidden by the author and can't be shown." once (4b P15 string), and clearing the filter keeps rows 4 and 7 collapsed.
  - `"no local value index"`: with `ValueIndex` and `DocumentSession` spied (`vi.mock`), a values filter makes 0 local constructions.
  - (append to `remote-sheet-view.test.tsx`) `"4c stream mounts the filter buttons"`: with `sheetSort: 1` the column header filter button is present; without it, absent (4b unchanged).
- [ ] **Step 2: Run** `bunx vitest run tests/xlsx-stream-controls.test.tsx tests/stream/remote-sheet-view.test.tsx`. Expected: FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:**
  - the same files;
  - unmodified: `bunx vitest run xlsx- interactive paged sheet viewer filter tests/components`;
  - `npx tsc --noEmit -p .`;
  - `bun run gates:core` (report the delta).
- [ ] **Step 5: Commit** `src/components/xlsx/streamControls.ts src/components/xlsx/InteractiveSheetDoc.tsx src/components/xlsx/FilterDropdown.tsx tests/xlsx-stream-controls.test.tsx tests/stream/remote-sheet-view.test.tsx`: `feat(stream-s4c): sort and filter menus on streamed sheets; typed conditions greyed when protected`.

### Task 11: The harness and the `stream-4c` browser project

**Files:**
- Modify:
  - `server/test/app.ts` (`?doc=sort-grid|filter-auto-hidden`);
  - `playwright.viewer.config.ts:62–123`: `STREAM_4C_SPECS`, the `stream-4c` project, `stream`'s `testIgnore`, a third webServer on `OV_STREAM_4C_PORT`;
  - `package.json:58`: `test:browser:stream` adds `--project stream-4c`.
- Create: `tests/browser/stream-sheet-4c.spec.ts`

**Interfaces:**
- Produces: the default client (unchanged HELLO `['sheet-view@1']`) gets sort and filter from a 4c server through `sheetSort` (P1).

- [ ] **Step 1: Write the failing specs** (`stream-sheet-4c.spec.ts`, against `?doc=sort-grid` and `filter-auto-hidden`):
  - `"sort descending moves the tiles"`: after the menu sort, `__ovSheetStats.pendingTiles` reaches 0 and the first row header reads the source row of the largest B;
  - `"filter by values hides rows and keeps author-hidden rows hidden"`;
  - `"number filters are greyed under protection"`;
  - `"a resume restores the sorted view"`: the page's socket is dropped (`page.context().setOffline(true)`, then `false`), the session resumes, and `sheet-order` restores the order. If the harness cannot drop the socket this way, delete the case and record that Task 9's fake-socket test covers it. A page reload is a new open, which resets view state (S4 §6.3), so it is not this test.
- [ ] **Step 2: Run** `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npm run test:browser:stream`. Expected: the new cases FAIL.
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run green:**
  - the same command; the `stream` and `stream-4b` projects pass with their spec files byte-identical (`git diff --stat fb624989 -- tests/browser` lists only the new file);
  - unmodified: `bunx vitest run tests/stream tests/office-file-stream.test.tsx tests/officeview-streaming.test.tsx`;
  - `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npx playwright test --config playwright.viewer.config.ts --project=chromium tests/browser/interactive-viewer`, unmodified.

  If an existing test or spec needs an edit to pass, STOP and report, without editing it.
- [ ] **Step 5: Commit** `server/test/app.ts playwright.viewer.config.ts package.json tests/browser/stream-sheet-4c.spec.ts`: `test(stream-s4c): stream-4c browser project with its own harness`.

### Task 12: Compatibility legs for 4c

**Files:**
- Modify: `scripts/stream-s4/compat-matrix.ts`

**Interfaces:**
- Produces: `bun scripts/stream-s4/compat-matrix.ts --base 53668e81 --base4a 1e2fbf4d --base4b fb624989`, whose new legs are:
  - **(h)** the 4b client (`fb624989`'s default, `['sheet-view@1']`) against the current server: 4b's `validateManifest` accepts the manifest (its `sheetSort: 1` is an unknown field there and ignored), 4b's `decodeSheetLayout` decodes a `sheet-layout` byte-identical to the 4b server's, and a find and a resize work;
  - **(i)** the current default client against the 4b server: `sheetView === 1`, no `sheetSort`, `doc.sheets.sortOps === undefined`, and a find still works;
  - **(j)** the current client against the current server: a sort round-trips and one tile at the new `viewGen` decodes.

- [ ] **Step 1: Add the legs**, with a third detached worktree for `--base4b`.
- [ ] **Step 2: Run** the command. Expected: eleven PASS lines, and `git worktree list` unchanged after.
- [ ] **Step 3: Commit** `scripts/stream-s4/compat-matrix.ts`: `test(stream-s4c): compatibility legs — 4b client × 4c server, 4c client × 4b server`.

### Task 13: Evidence: the wire audit, the hostile battery, pivot strips and corpus parity

**Files:**
- Create: `server/test/sheet-sort-audit.test.ts`, `server/test/sheet-sort-corpus.test.ts`

**Interfaces:**
- Consumes: the whole 4c server path through the harness `Client`.

- [ ] **Step 1: Extend the wire audit** (`}, 30_000)`). Capture every server→client byte of a `sheetSort` session that:
  - lists every used column's values (every page);
  - range-detects from A1;
  - sorts by each used column;
  - filters each column by its first two list entries and by `text contains "e"`;
  - fetches `sheet-order` and every tile of the used extent at (zoomStep 5, dpr 1).

  Run it for `sentinels()`, every XLSX hidden fixture, `sortWithheld`, `paintedFormats`, `rotatedBox`, `offViewportCovered` and `unresolvedAnchor`. Then check:
  - `"no sentinel or withheld text in any 4c answer"`: no decoded message or payload contains `SENTINEL_F`, `SENTINEL_N`, `1234.5678`, `=LEN`, any `marker`, a hidden sheet's name, `COVER-B20`, `MASKED`, `OFFVIEW-Z5000` or a covered `ROT-` text. No value-list item text equals B21's `5`, B22's `95`, or `paintedFormats`' `1234` or `5678`.
  - `"order answers carry positions only"`: every `sheet-order` part has only the keys `reqId`, `type`, `seq`, `last`, `head`, `band` and `mask`.
- [ ] **Step 2: The pivot strip on the wire.** On `hiddenPivotCache` under visible-only:
  - no frame contains `PIVOT-SECRET` or `SLICER-SECRET`;
  - every tile equals the same fixture's tile with those items physically removed from the cache parts (the parent T2 method).

  Non-vacuity: under `'full'`, the slicer's tile differs from the stripped render. If it does not (slicers paint no item text), record that, keep the frame check, and say so in Task 14.
- [ ] **Step 3: The hostile battery** (S4 §6.6, §9 4c), as one table over `{sortGrid, filterAutoHidden}` × `{forged list id, another sid's id, another column's id, out-of-grid range, hidden key column, stale viewGen, typed rule under visible-only, 65 rules}`. Each refuses with its P4 code, and afterwards:
  - `viewGen`, the layout bytes, `counts` and one probe tile are byte-identical to before;
  - the `sheet_ops` row holds the code.
- [ ] **Step 4: Corpus parity** (`}, 120_000)`, `bun scripts/fetch-corpus.ts` first). For each corpus XLSX, and for each sheet with a used range (at most 20 sheets, first 3 used columns), under `'full'` and under visible-only:
  - sort asc and desc by each column;
  - filter `values` (the first list entry excluded), `text contains` the first entry's first character, and blanks only.

  Each `sheet-order` must equal Task 5's oracle over the deleted fixture: the key view's withheld cells deleted, then C1. Assert at least one corpus file was visited. Count the corpus sheets that have a withheld key cell; if it is 0, the visible-only half is not evidence of withholding (as 4b's corpus ink oracle was not), and Task 14 says so.
- [ ] **Step 5: Run** `cd server && bun test test/sheet-sort-audit.test.ts test/sheet-sort-corpus.test.ts`. Expected: PASS, since Tasks 1–8 supply the behaviour. These are evidence, not RED-first. A failure is fixed in `key-view.ts`, `view-ops.ts` or `ws.ts` with a failing case first, never by changing the oracle.
- [ ] **Step 6: Commit** `server/test/sheet-sort-audit.test.ts server/test/sheet-sort-corpus.test.ts`: `test(stream-s4c): 4c wire audit, pivot strip, hostile battery, corpus row-order parity`.

### Task 14: Gates, validation and spec amendments

**Files:**
- Create: `docs/validation/streaming-s4c.md`
- Modify:
  - `docs/superpowers/specs/2026-11-22-streaming-s4-design.md`: an "As built (4c)" block under §11.2, and §10 amendments;
  - `docs/superpowers/specs/2026-11-07-server-streaming-design.md`: §6.2, §6.4, §9.3, §12, §14 and the refusal table, as dated amendments.

- [ ] **Step 1: Run every gate, one at a time.**
  - `cd server && bun test` and `bunx vitest run`;
  - `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npm run test:browser:stream` (projects `stream`, `stream-4b`, `stream-4c`), and the unmodified viewer browser suites;
  - `npx tsc --noEmit -p .`, `npm run -s typecheck:server` and `npm run -s lint`;
  - `bun run gates:core`: red as at the base, with the delta;
  - the compatibility matrix;
  - `bun scripts/stream-s4/tile-p95.ts`, recorded and not gated (an identity view; 4c's tiles add only order and visibility);
  - `git diff --numstat fb624989 -- tests/stream`: 0 deletions.

  Pre-existing failures (the 4b §Known flakes) are listed, not fixed.
- [ ] **Step 2: Write the amendments.**
  - **S4 "As built (4c)":** P1–P16, with any ruling that changed them.
  - **Parent:**
    - §6.2: the `sheetSort` manifest key;
    - §6.4: the key view and masked strings;
    - §9.3: the canonical view with `sortFilter`, and the `'c1'` tier;
    - §12: C1 as built (collapsed rows);
    - §14: the `'sheet'` class completed, admission through `DocumentLedger.hold`, `MAX_ORDERED_SHEETS`;
    - the refusal table: the seven new codes;
    - §19: item 6's sort-order limit.
  - **Spec gaps and deviations:**
    - §6.6's "no local cap" (there is `VALUE_DISPLAY_LIMIT`);
    - §6.3's canonical spec lacking `sortFilter`;
    - §6.6's reply as one `ViewCommitDelta` (chunked `sheet-order` instead; metrics derived client-side);
    - §6.6 progress (two steps client-side);
    - §6.7 `hidden: 'filter'` (Open item 2);
    - the masked-string withholding (Open item 3) and the manifest capability (Open item 1);
    - §6.9/§6.11's pivot notice (none: the client cannot know);
    - the `classification` on the first values page.
- [ ] **Step 3: Write `docs/validation/streaming-s4c.md`:**
  - each S4 §9 4c "done when" line with the test that proves it;
  - the gate table, the size-gate delta, the compatibility output and the p95 record;
  - what is not machine-checked: screen readers on the menus, non-Chromium browsers, and the corpus case if no corpus sheet carries a drawing or a withheld cell.
- [ ] **Step 4: Commit** `docs/validation/streaming-s4c.md docs/superpowers/specs/2026-11-22-streaming-s4-design.md docs/superpowers/specs/2026-11-07-server-streaming-design.md`: `docs(stream-s4c): validation, spec amendments`.

## Spec coverage

| S4 spec | Task |
|---|---|
| §6.1 rows marked 4c: `sort`/`filter` → `sheet-sort`/`sheet-filter`; `valueQuery` → `filter-values`; `read`/`detectRange` → `range-detect`; `pivot*`/`control*` refused; `prepareColumns` refused | 8; 7; 7; 3 (P14) |
| §6.2 step 2 (C1), steps 3–6 as keys (blank, not skipped), step 4 (`coverage-unknown` for sort and filter), step 5 (keys on base coverage) | 5, 8; 1 (P5); 8 (P4); 1 |
| §6.2 parity oracle for row order | 5, 13 |
| §6.3 view state: sort/filter spec, `viewStateHash`, lifetime, ceiling, admission | 4 (P2, P11) |
| §6.6 commands, target authorization, value lists, oracles (typed refusal), hostile and tautology filters, range detection | 7, 8 (P4, P6, P7) |
| §6.8 `'sheet'` completed; message caps | 3, 7, 8; 2 |
| §6.9 pivot refusal on every stream; strip asserted on the wire | 3, 13 |
| §6.11 menus present from 4c; typed items greyed with the tooltip | 10 (P16) |
| §6.12 items 1, 2, 6, 7 | 5, 7, 8, 10 |
| §9 4c: row-order parity over a corpus | 5, 13 |
| hostile and tautology filters; author-hidden rows stay hidden | 8, 13 |
| off-viewport covered-cell filter values, rotated box, unresolved anchors | 8, 13 |
| forged or stale value ids refused | 7, 8, 13 |
| eight pivot and control kinds, `prepareColumns`, `read`; stripped caches on the wire | 3, 13 |
| T6 pair: per-session view state isolated | 8 |
| rate class: 31st refused, one in flight | 7, 8 |
| typed conditions refused on `visible-only` before state moves; `'full'` parity | 8 |
| `sheet_ops`, `sheet_views` | 4, 7, 8 (P12) |
| §10 items 1 (key use), 3 (keys on base coverage), 4, 6 (sort-order limit), 11 | 1, 5, 14 |
