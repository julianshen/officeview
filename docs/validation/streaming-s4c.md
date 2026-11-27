# Server streaming — Stage 4c validation (server sort and filter, value lists, range detection, the pivot refusal and the withholding tests)

Spec: `docs/superpowers/specs/2026-11-22-streaming-s4-design.md` (§9 4c; "As built (4c)" under §11.2). Plan: `docs/superpowers/plans/2026-11-26-streaming-s4c.md` (plan calls P1–P16; pre-flight findings F1–F5, the inserted Task 11a and the ledger rulings in `.superpowers/sdd/2026-11-26-streaming-s4c/`). Branch `feature/streaming-s4c`, off main fb624989 (4a and 4b merged). **Code head eca7be34.** Every gate below was measured there on 2026-11-28, one at a time, on macOS, darwin-arm64, Bun 1.4.5, with no other job of this task running. The docs commits after eca7be34 change only this file and the two spec files (no code, no test).

## Status

Every 4c gate is green at eca7be34 except these. The first is unexplained; the others are pre-existing or environmental. None was worked around.
- **One server case failed in the full run, and passes alone:** `server/test/sheet-sort-filter.test.ts` "a million-row sort" (in `describe('rates, cancel, admission and audit (P8, P11, P12)')`). Its first sort was answered `refused {code: 'policy-refused'}`; the test took 15.1 s. The only load reading during the suite, about 8 minutes in with the suite running, was 7.93 / 15.95 / 19.10 (1 / 5 / 15 min). Alone, the file passed 15 of 15 three times in a row (11.0–11.5 s each). Because this was the first full run to include Task 13's files, and `sheet-sort-audit` ran just before it in the log, the two files were also run together: 35 of 35 three times in a row (about 18 s each). The test is 4c's (Task 8). Earlier full runs on this branch passed it (at f9872fc5, 8e1f0d80 and 7513a97d). Tasks 5, 7 and 8 each first surfaced a real engine leak as an intermittent refusal of a large sort or list, so this failure is not assumed to be load alone. **Root cause not found.** The sort handler in `server/src/ws.ts` answers `policy-refused` on several paths: the document ledger refusing a `touch`, the pool's `sheetView` job throwing anything but `AdmissionRefused`, a reply of the wrong kind or view, an order that fails to encode, and a failed audit write. Which one fired was not traced. It is listed as an unexplained failure, not fixed; no timeout or limit was changed.
- **The eager size gate** (index closure) is red, as it was at the base fb624989. 4c ends at 426,561 B against Task 3's 426,556 B at the base, Δ +5 B (§Size gate).
- **One vitest case fails, as at the base:** `tests/slideshow-screens.test.ts` "the drag hint works on a popup (other-realm) document and stays dismissed" (Node has no `localStorage` in that realm). 4c does not change the test or `src/slideshow`: `git diff --stat fb624989 -- tests/slideshow-screens.test.ts src/slideshow` is empty.
- **The `stream` browser project still sits at the per-IP session limit.** In one of the two runs, `tests/browser/stream-viewer.spec.ts:53` "error: an unknown document ends in the typed auth-failed error" failed. The trace shows its HELLO refused with `Unexpected response code: 429` (the per-IP gate, `server/src/rate.ts`, 30 new sessions per minute). This is the same case 4b saw at 91fb7237. The `stream` project's spec files are byte-identical to the base, and 4c's specs run in their own `stream-4c` project.

The compatibility matrix passes all eleven legs, and `tests/stream` has 0 deleted lines since fb624989. T8 tile p95, recorded and not gated, is 94.3 ms at dpr 1 and 55.7 ms at dpr 2 (GREEN against 120 ms).

## S4 §9 4c "done when"

| Line | Proven by |
|---|---|
| **sort/filter row-order parity** with local mode on the snapshot, across a corpus of sort keys and filter rules | `server/test/sheet-sort-corpus.test.ts` "every sheet-order equals the local engine over the deleted fixture, then C1": every corpus XLSX under `'full'` and under `visible-only`, over the whole wire path, against `viewerControls` over an independent parse (saved values) with the key view's withheld cells deleted and its rewritten cells taken from `sheetKeyView`, then C1. Sorts asc and desc, a values rule, `text contains`, blanks only. This record's server run (eca7be34) printed: 33 xlsx, 61 (file, policy) sessions, 86 sheets, 1,200 comparisons (645 committed, 555 refused by both with the same local reason), 0 `visible-only` sheets with a withheld key cell, 0 with an authored collapsed body row. **The `visible-only` half is parity evidence only** (§Not machine-checked). Fixtures: `server/test/sheet-view-ops.test.ts` "sort parity", "filter parity and C1", "replay equals incremental", "display-text merge", "two values columns"; `server/test/sheet-sort-filter.test.ts` "sort round trip", "full allows typed conditions" |
| hostile and tautology filters are refused or neutral, as §6.6 lists, and author-hidden rows stay hidden under every rule | **Neutral:** `server/test/sheet-sort-filter.test.ts` "tautology filters": `text neq ""` over `filterAutoHidden` commits and leaves the saved-autoFilter rows hidden, `collapsedRows` unchanged (C1); under `'full'` the same filter equals the local engine. **Refused, not neutral:** §6.6's other tautology, `gte −1e308`, is a typed condition, so under `visible-only` the same test asserts it is refused `filter-condition-refused`. **Hostile:** `server/test/sheet-sort-audit.test.ts` "`${f.name}: each hostile op refuses with its P4 code and moves nothing`" (on `sortGrid` and a `filterAutoHidden` variant with column B `width="0"`: forged, other-sid, other-column and older-`viewGen` list ids `stale-filter-values`; out-of-grid range and 65 rules `policy-refused`, uncharged, no audit row; hidden key column `hidden-target`; stale `viewGen` `stale-view`; `number gt` `filter-condition-refused`. After each, the `sheet-order` parts, the `sheet-layout` bytes and a probe tile equal the snapshot taken before; each valid twin commits on its own session). `server/test/sheet-sort-filter.test.ts` "authorization refuses before anything moves" (11 cases; no engine sort or filter ran for any). **C1 in every delta:** `server/test/sheet-sort-audit.test.ts` "C1 and withheld keys in every order delta"; `server/test/sheet-view-ops.test.ts` "filter parity and C1"; `server/test/sheet-detect-see-through.test.ts` "a header sort and a filter over the detected range commit; author-hidden rows stay hidden". Browser: `tests/browser/stream-sheet-4c.spec.ts` "filter by values hides rows and keeps author-hidden rows hidden" |
| **off-viewport covered-cell filter values** withhold, including the rotated box at `x 40..60, y 30..50`; an `anchors.unresolved > 0` sheet withholds every cell | `server/test/sheet-filter-values.test.ts` "off-viewport covered cell" (`OFFVIEW-Z5000` absent, blanks present, no tile ever requested), "rotated box and unknown coverage" (`ROT-2-1`, `ROT-3-1` absent, `ROT-4-1` and `ROT-2-2` listed; under unresolved anchors every column lists `items: []` with blanks), "lists are display text, withheld cells only as blanks"; `server/test/sheet-key-view.test.ts` "withheld cells are blank keys", "unknown coverage"; `server/test/sheet-sort-audit.test.ts` "no sentinel or withheld text in any 4c answer" (every 4c answer over `rotatedBox`, `offViewportCovered`, `unresolvedAnchor` and the other fixtures, by needle and by exact field) |
| forged or stale value ids are refused | `server/test/sheet-filter-values.test.ts` "ids are sid- and column-bound" (another sid, column, range, sheet, `viewGen`, an evicted list, out-of-range and malformed intervals: all `stale-filter-values`); the hostile battery above; `server/test/sheet-sort-filter.test.ts` "authorization refuses before anything moves" (a list issued at `viewGen` 0 is stale at 1) |
| the eight pivot and control kinds get `pivot-refused`; `prepareColumns` and raw `read` get their own typed refusals; stripped caches are asserted on the wire | `server/test/sheet-sort-handshake.test.ts` "pivot and control kinds are refused" (each kind in tiles and units mode, no token, charged `'message'` only: `pivot-refused`, `server-prepared`, `stream-no-cell-data`; the socket stays open); `server/test/sheet-sort-audit.test.ts` "no frame carries a pivot or slicer cache item; every tile equals the stripped render" (`hiddenPivotCache` under `visible-only`: no frame carries `PIVOT-SECRET` or `SLICER-SECRET`; all 4 tiles equal a twin with the items physically removed; under `'full'` the slicer tile differs, so the check is not vacuous) |
| per-session view state is isolated: two sids with different filters never receive each other's tiles or deltas (T6 pair) | `server/test/sheet-sort-filter.test.ts` "T6 pair" (a filtered sid and an unfiltered sid: separate canonical tiles that differ; a second sid with the same view shares the canonical tile; the unfiltered sid's view never moves and its `viewGen`-1 tile request is `stale-view`), "text ops follow the view"; `server/test/sheet-view-spec.test.ts` "sortFilter keys the hash" |
| the rate class refuses at 31/min and allows one in flight | `server/test/sheet-sort-filter.test.ts` "the sheet class" (30 sorts commit, the 31st is `rate-limited` with its `sheet_ops` row; a second op while one is in flight is `rate-limited`; neither closes the socket); `server/test/sheet-filter-values.test.ts` "rates" (first pages charge `'sheet'`, cursor pages only `'message'`, the 31st first page is refused, 31 audit rows) |
| a typed number, date, date-preset or compound-with-typed-child condition is refused on a `visible-only` stream before any state moves; on `'full'` it gives local row-order parity, and so do C1 and the coverage rules there | **Refused:** `server/test/sheet-sort-filter.test.ts` "authorization refuses before anything moves" (`number gt`, `compound(text, datePreset)`, `date eq`: `filter-condition-refused`, `viewGen` stays 0, no engine op); the hostile battery. **`'full'`:** "full allows typed conditions" (`number gt 50` commits; shown rows and counts equal the local engine. **Wire parity of a typed condition under `'full'` is shown for `number` only**: no test runs a `date`, `datePreset` or typed `compound` rule over the wire under `'full'`, and the corpus ops include no typed rule. The session clock for date presets is checked at the engine level in `server/test/sheet-key-view.test.ts` "the injected clock"; `typedFilters: true` on the manifest per `server/test/sheet-sort-handshake.test.ts` "sheetSort on every sheet-view manifest"); C1 off under `'full'`: "tautology filters" (the `'full'` half equals the local engine), `server/test/sheet-view-ops.test.ts` "filter parity and C1" (under `'full'`, `contains ""` shows the authored-hidden rows 4 and 7, as locally) and `server/test/sheet-sort-filter.test.ts` "a full filter that shows authored-hidden rows gets its bottom tiles"; coverage off under `'full'`: `server/test/sheet-key-view.test.ts` `describe('full is the identity')` (one `test(name, …)` per fixture: the key view is the projected sheet, nothing withheld) and the corpus test's `'full'` sessions |
| full T8 green | T8's parts (parent §16): authorization with hostile and tautology filters, covered and off-viewport covered filter values, the per-session view-state key and row-order parity are the rows above. **Tile p95 is recorded, not gated** (the brief: 4c's tiles add only order and visibility to an identity view); see §T8 p95 |

Further 4c evidence by plan call:

| Call | Tests |
|---|---|
| P1 capability | `server/test/sheet-sort-handshake.test.ts` (all 5: "sheetSort on every sheet-view manifest", "a DOCX gets no new keys", "the 4c ops need sheetView", "4b layout bytes unchanged", "pivot and control kinds are refused"); `tests/stream/manifest.test.ts` `describe('MANIFEST sheetSort and typedFilters (Stage 4c)')` (5 cases); `tests/stream/remote-sheet-sort.test.ts` "sortOps present only with sheetSort"; compatibility legs (h)–(j) |
| P2, P11 canonical view, ceiling, admission | `server/test/sheet-view-spec.test.ts` (6 tests, incl. "identity, the C1 tier and orderBytes (P11)", "two ordered sheets per sid", "a refused hold or an unchanged view changes nothing"); `server/test/cache.test.ts` "identity hashes unchanged", "admission"; `server/test/sheet-sort-filter.test.ts` "a million-row sort" (≤ 160 parts of ≤ 64 KiB; the hold is charged once for two sids and released at the last close) |
| P3 replay | `server/test/sheet-view-ops.test.ts` "replay equals incremental", "the spec moves as P3 says" |
| P4 shape checks | `tests/stream/sheet-wire.test.ts` `describe('sort and filter args')`, `describe('values args')`, `describe('range-detect and order args')`; `server/test/sheet-sort-filter.test.ts` "authorization refuses before anything moves" |
| P5 key view | `server/test/sheet-key-view.test.ts` (all: "engines import under Bun", "withheld cells are blank keys", "painted styles", "the key view ignores session sizes", "unknown coverage", `describe('full is the identity')`, "the injected clock"); `server/test/sheet-view-ops.test.ts` "withheld cells sort and list as blank", "painted keys list as the tile paints them (Task 1 style note)" |
| P6 value lists | `server/test/sheet-filter-values.test.ts` (8 tests in `describe('filter-values (S4 §6.6; plan P6)')`, incl. "paging and search", "values on without copy"); `tests/stream/sheet-wire.test.ts` `describe('values page')`; `server/test/sheet-view-values-memory.test.ts` "60 sequential values jobs over a 10,000-entry list, no forced GC" |
| P7 the order on the wire | `tests/stream/sheet-wire.test.ts` `describe('order parts')` (incl. "encodes a 1,048,576-row reverse band in at most 160 parts of at most 64 KiB, and decodes equal", "throws for a dropped part, a duplicate entry, a gap, an overlap, a reorder and 161 parts"); `tests/xlsx-stream-sort.test.ts` "applyOrder moves the view", "client geometry equals the server's under the view"; `server/test/sheet-sort-audit.test.ts` "order answers carry positions only" |
| P8 progress, cancel, lost replies | `server/test/sheet-sort-filter.test.ts` "cancel", "resume during an in-flight sort"; `tests/stream/remote-sheet-sort.test.ts` "resume: sort and filter reject session-rotated, values and detect are re-issued", "abort sends sheet-cancel"; `tests/xlsx-stream-sort.test.ts` "a sort commits the order it answers, with progress 0 / 1 then 1 / 1", "a lost sort answer refetches", "every refusal refetches; view-refused, stale-view and a cancel map to their outcomes" |
| P9 text ops and tiles under the view | `server/test/sheet-view-text.test.ts` (5 tests, incl. "a sort that slides a cell under a fixed drawing blanks its text", "a job under a view is admitted with its order bytes"); `server/test/sheet-sort-filter.test.ts` "a full filter that shows authored-hidden rows gets its bottom tiles"; `server/test/sheet-sort-audit.test.ts` "a sorted view paints no withheld cell" (T2 method: tiles equal a twin with the withheld cells deleted); `server/test/trace-tiles.test.ts` "a capture of a sorted view traces" |
| P10, P13 drawings and engine refusals | `server/test/sheet-view-ops.test.ts` "local guards refuse"; `server/test/sheet-sort-filter.test.ts` "a local refusal is view-refused and moves nothing"; `tests/stream/remote-sheet-sort.test.ts` "view-refused carries the local refusal", "pins the local strings the refusals are rebuilt with"; `tests/xlsx-stream-controls.test.tsx` "refusals speak local strings" |
| P12 audit | `server/test/audit-store.test.ts` "sheet_ops", "sheet_views spec"; `server/test/sheet-sort-filter.test.ts` "audit" (typed operand and text operand never in `sheet_ops`; stored in `sheet_views`) |
| P14 pivot refusal | as in the "done when" table |
| P15, P16 client | `tests/xlsx-stream-controls.test.tsx` (9 tests: "sort from the column menu", "filter by values", "typed conditions greyed", "pacing", "author-hidden notice", "no local value index", …); `tests/stream/remote-sheet-view.test.tsx` "4c stream mounts the filter buttons"; browser `tests/browser/stream-sheet-4c.spec.ts` "sort descending moves the tiles", "number filters are greyed under protection" |
| Range detection, see-through (owner, Task 11a) | `server/test/sheet-detect-see-through.test.ts` (11 tests in 4 describes, incl. "content-withheld cells still end a region (Task 1 ruling stands for them)", "rows: text, number and empty hidden rows (inside and directly under the data) answer identically"); `tests/xlsx-viewer-range-see-through.test.ts` (7 tests); `server/test/sheet-view-ops.test.ts` "range detection over servable occupancy"; `server/test/sheet-filter-values.test.ts` "range-detect" |
| Clearing and the refetched view (Task 9 rulings) | `server/test/sheet-sort-filter.test.ts` "sort then clear: the identity order, the filter kept, a sheet_ops row; the head carries caseSensitive and the filter range"; `tests/stream/sheet-wire.test.ts` `describe('sort clear (Task 9 fix round 1)')`; `tests/xlsx-stream-sort.test.ts` `describe('Task 9 fix round 1: clears and a refetched view')` |
| Engine leak fixes (Tasks 5, 7, 8) | `tests/xlsx-viewer-guard-page-settle.test.ts` "a full-column sort guard over 1,048,576 rows stays within the main-thread ledger"; `server/test/sheet-view-values-memory.test.ts` (above); `server/test/sheet-view-sort-memory.test.ts` "14 sequential million-row sorts, no forced GC" |

## Gates

All measured at eca7be34 on 2026-11-28, one at a time, in this order. No other job of this task ran beside any of them. The load (1 / 5 / 15 min) read 7.93 / 15.95 / 19.10 about 8 minutes into the server suite, with the suite running. It was 5.76 at the start of vitest, 8.22 at its end, about 6–7 during the browser runs and 5.9 at the start of `tile-p95`. The docs commits after eca7be34 touch only `docs/validation/streaming-s4c.md` and the two spec files, so they change no code or test that a gate measured.

| Gate | Command | Result at eca7be34 |
|---|---|---|
| server | `cd server && bun test` | 1,655 pass, **1 fail**, 1,656 tests in 118 files, 982 s. The fail is "a million-row sort" (see Status). Alone, `bun test test/sheet-sort-filter.test.ts`: 15 pass, 0 fail, three runs in a row |
| client | `bun scripts/fetch-corpus.ts` (70 verified), then `bunx vitest run` | 795 of 796 files; 16,173 pass, **1 fail**, 65 skipped, 1 todo (16,240), 1,096 s. The fail is `tests/slideshow-screens.test.ts` (pre-existing) |
| browser, stream | `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npm run test:browser:stream` (projects `stream`, `stream-4b`, `stream-4c`) | run 1: **55 passed**, 1 skipped, 3.7 min. Run 2: 54 passed, **1 failed** (`stream-viewer.spec.ts:53`, HTTP 429, see Status), 1 skipped, 3.8 min. All three `stream-4c` cases passed in both runs. The skip is the existing no-codec media case. A first attempt at run 2 exited 134 before Playwright printed anything; it was re-run unchanged, and the run above is that re-run. The abort was not diagnosed |
| browser, viewer (unmodified) | `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npx playwright test --config playwright.viewer.config.ts --project=chromium tests/browser/interactive-viewer` | **67 passed**, 2.5 min |
| typecheck | `npx tsc --noEmit -p .` | clean (exit 0) |
| typecheck, server | `npm run -s typecheck:server` | clean (exit 0) |
| lint | `npm run -s lint` | clean (exit 0) |
| eager gate | `bun run gates:core` | core closure ok (299 B), engine modules ok, core purity ok; index closure FAIL as at the base (§Size gate) |
| compatibility | `bun scripts/stream-s4/compat-matrix.ts --base 53668e81 --base4a 1e2fbf4d --base4b fb624989 --tmp <scratchpad>/cm` | **11 PASS**, exit 0; `git worktree list` byte-identical before and after (diffed) |
| T8 (not gated) | `bun scripts/stream-s4/tile-p95.ts` | GREEN, exit 0 (§T8 p95) |
| append-only | `git diff --numstat fb624989 -- tests/stream` | 0 deletions: `manifest.test.ts` +39, `remote-sheet-sort.test.ts` +332 (new), `remote-sheet-view.test.tsx` +37, `sheet-wire.test.ts` +304. `tests/browser/stream-sheet.spec.ts` and `stream-sheet-4b.spec.ts` are byte-identical to fb624989 |

Outside `tests/stream`, one existing assertion was widened, as the owner permitted (Open item 1): `server/test/sheet-view-handshake.test.ts`, 1 line changed, the `sheetView` manifest's key set gains `sheetSort`.

## Size gate (T11)

`bun run gates:core` at eca7be34 (the 64 module-set lines naming `node_modules` files are omitted here. They are 32 pairs: each `+` entry is printed as an absolute path under the main checkout, because the worktree resolves `node_modules` from there, and matches a `−` entry of the same relative path. Diffing the two sets leaves `src/xlsx/rowFont.ts` as the only unpaired entry):

```
ok   core closure 299 B over 1 chunk(s) ≤ 38720
ok   engine modules in core closure: none
ok   core purity: no react imports, no document global
FAIL index closure 426561 B (baseline 417436, Δ 9125, recorded 3193, limit 420645)
FAIL index module set differs from baseline ± recorded changes:
       …
       + src/xlsx/rowFont.ts
```

- **The base was already red.** Task 3 measured 426,556 B at the base (with the same `src/xlsx/rowFont.ts` line, from xlsx-cf). That figure is taken from the Task 3 report; it was not re-measured here. The gate is not in CI. No limit or baseline was changed.
- **4c's delta is +5 B** (426,561 − 426,556). The task reports record per-task moves of a few bytes each (0 at Task 3, +2 and +1 at Task 9, −6 at Task 10, +5 at Task 11a), and they disagree on the running figure (Task 9 gives its base as 426,561 B), so only the endpoints above are claimed. The moves are hashed lazy-chunk names in the index's `dynamicImports`. The only `src/` module in the differing set is `rowFont.ts`, as at the base, so no 4c module joined the index closure.
- New client code is reachable only from the lazy chunks: `streamControls.ts`, `streamWorkbook.ts` and the `InteractiveSheetDoc.tsx` / `FilterDropdown.tsx` edits sit behind `RemoteSheetView` and `InteractiveSheetDoc`. `dist/core.js` stays pure and engine-free (the three `ok` lines).

## Compatibility matrix

`bun scripts/stream-s4/compat-matrix.ts --base 53668e81 --base4a 1e2fbf4d --base4b fb624989 --tmp <scratchpad>/cm`, exit 0, at eca7be34:

```
PASS (a) base client x current server: no sheetMode on the raw manifest; unit (raster), 1 search hit(s), copy 22 chars, 1 print page(s) 336x20
PASS (b) current client x current server: sheetMode tiles, sheet-layout decoded (0 row / 0 col intervals), tile (0,0) decoded 512x512, manifest validated on grid500x100
PASS (c) current client (opted in, HELLO carried sheet-tiles@1) x base server: accepted, no sheetMode, unit arrives (raster)
PASS (d) 17 features (16 distinct unknowns + sheet-tiles@1) closed 4000 = unsupported-version; the 16-feature control was accepted
PASS (e) 4a client x current server: 4a validateManifest accepted it, no sheetView on the raw manifest, 4a decodeSheetLayout decoded sheet-layout (1 part(s), no session-size keys), tile (0,0) decoded 512x512
PASS (f) current default client (HELLO sheet-view@1) x 4a server: units mode (no sheetMode), 2 unit(s), 1 search hit(s), copy 22 chars
PASS (f') current client (streamSheets 'tiles', HELLO sheet-tiles@1) x 4a server: tiles mode, doc.sheets.view undefined, tile (0,0) decoded 512x512
PASS (g) current client (streamSheets 'units', no HELLO features) x current server: units mode, Stage 3 manifest keys, 2 unit(s), copy 22 chars
PASS (h) 4b client (HELLO sheet-view@1) x current server: 4b validateManifest accepted sheetSort:1 and dropped it, 4b decodeSheetLayout decoded 1 sheet-layout part(s) byte-identical to the 4b server's, find 6 hit(s), resize row 1 -> 48 at viewGen 0 -> 1
PASS (i) current default client x 4b server: sheetView 1, no sheetSort on the raw or parsed manifest, sortOps undefined, find 6 hit(s)
PASS (j) current client x current server: sheetSort 1 (no typedFilters on visible-only), sort B desc answered at viewGen 0 -> 1 (display row 1 = source 499), layout sort summary matches, tile (0,0) at the new viewGen decoded 512x512, clear by keys [] -> viewGen 2
```

Legs (a)–(g) are 4a's and 4b's, unchanged. (h)–(j) are 4c's (Task 12): (h) a 4b client gets 4b's `sheet-layout` bytes from a 4c server, compared with what the same client gets from the 4b server; (i) a 4c client against a 4b server has no `sortOps`; (j) a sort, a tile at the new `viewGen` and the clear. The script removed its three base worktrees afterwards.

## T8 p95 (recorded, not gated)

`bun scripts/stream-s4/tile-p95.ts` at eca7be34, R3 protocol as in 4a and 4b (grid500x100, an identity view). 4c's tiles add only order and visibility, so T8 is not re-gated; this is the one recorded run.

| Measure | dpr 1 (260 tiles) | dpr 2 (1,000 tiles) |
|---|---|---|
| R3 p50 / **p95** / max (ms) | 70.5 / **94.3** / 860.6 | 49.1 / **55.7** / 853.2 |
| client send → `tiles-done` p50 / p95 / max (ms) | 71.6 / 95.8 / 861.6 | 50.2 / 56.7 / 854.4 |
| loopback (request → `drawImage`, Bun + napi) p50 / p95 / max (ms) | 71.3 / 83.3 / 867.2 | 50.7 / 56.8 / 866.6 |

- All 1,260: R3 p95 78.5 ms; 4 worker recycles; 0 R3 intervals above the client's own interval. Verdict line: `T8 (p95 ≤ 120 ms at dpr 1 and dpr 2): GREEN`.
- `--batch 64` (dpr 2, 1,000 tiles, cold): 26,899 ms, 37.2 tiles/s, 3 recycles; per-tile t0→t1 p50 927.4, p95 1,959.5, max 2,388.9 ms.
- Load (1/5/15 min): start 5.90/6.78/8.14, after the serial run 5.44/6.60/7.98, end 5.23/6.19/7.66. 4b recorded 83.8 ms (dpr 1) at a load of about 4; on this one run the dpr 1 p95 is 94.3 ms. 4c edited the tile path (`server/src/sheet/tile.ts`, `render.ts`, `pool.ts`, the tile-grid bound), so the rise is not attributed to load; it was not investigated, since T8 p95 is recorded and not gated for 4c.
- The max is still the worker-recycle stall (about 0.86 s).

## Known flakes and environment notes

Seen during 4c; listed, not fixed. The server and vitest cases passed when their files ran alone, except `slideshow-screens`, which fails alone too.
- `server/test/sheet-sort-filter.test.ts` "a million-row sort": refused `policy-refused` in the full run at eca7be34; passes alone and beside `sheet-sort-audit` (Status). **New in 4c; unexplained; not assumed to be load.**
- `server/test/client-text.test.ts` "a newer select settles the older one": timed out (60 s) in the full run at 8a1750e9; 30/30 alone three times. The file and its code are untouched by 4c.
- `tests/xlsx-viewer-worker-overflow.test.ts` `it.each(…)('SPEC guard %s preserves old rows and raw ownership')`, the `epoch-after-stage` case: failed in the full vitest runs at b46272c1 and 8e1f0d80 (4b-known, root cause not found in 4b either).
- `tests/slideshow-screens.test.ts`, one case, fails under the current Node (no `localStorage` in that realm), as on main.
- `tests/browser/stream-viewer.spec.ts:53` in the `stream` project: the per-IP 429 (Status). The project has no headroom; whether the base also crosses 31 sessions in some runs was not measured.
- An invalid run, not a flake: the full server suite at 2664737d read Task 4's half-written files (`audit-store.test.ts` "Export sheetViewSpecOf not found"), because a task agent was active. From then on, full suites ran only between tasks.

## Rulings and findings as built

The full text is in the S4 design, §11.2 "As built (4c)", and the dated parent amendments (2026-11-28).
- **Owner, 2026-11-27:** plan Open items 1–3 approved as recommended: the manifest key `sheetSort: 1` (with `typedFilters: true` under `'full'`); find never matches a session-filter-hidden row (local parity, no `hidden: 'filter'`); masked strings sort, list and match as their painted text under `visible-only`.
- **Owner, 2026-11-27 (Task 11):** under `visible-only`, author-hidden rows are stepped over by range detection, and a shift-click rectangle across them is one rectangle. This supersedes the Task 1 ruling for author-hidden rows only. Ruling: the same holds for author-hidden columns. Content-withheld cells still end a region (Task 11a inserted to build it).
- **Pre-flight F1–F5:** the layout carries the sort; the 4b view state and hash stay readable (F2); `InteractiveSheetDoc` wires the stream menus (F3); `SessionViews.commit` takes a releasable hold (F4); sorted or filtered tiles without sizes are accepted (F5).
- **Wire shapes:** flat requests (Task 2); a sort clear is `keys: []` (Task 9); the order head carries optional `caseSensitive` and `filterRange` (Task 9); the search cursor is `${list}~${gen}:${offset}` and `total` is the list's entry count (Task 7).
- **Admission:** `viewTier` is `'c1'` when a filter or a `sortFilter` is set (Task 4); `orderBytes` counts the union of the ranges' body rows (Task 4); a job that re-derives a view is admitted with its order bytes (Task 6).
- **Audit:** a rate-limited first values page has its row; cursor pages and shape refusals have none (Tasks 7, 13); a failed `ok` write after commit answers `policy-refused` with the view moved, and the client re-fetches (Task 8).
- **Engine leaks fixed at the cause** in shared local code (Tasks 5, 7, 8), each with a regression; a 6,304 B per-sort residual released by GC is accepted and asserted bounded (Task 8).
- **`ht="0"` rows** moved by a sort are not pinned; their slot draws blank, and their text never reaches the wire (Task 11a).
- **64 sort keys on the wire, 8 committable** (Task 13): the local engine sorts at most 8 distinct keys and refuses a repeated key column (`view-refused`, reason `source-error`), as locally. The 64-column filter commits.
- **Client:** a refetched filter column is a `committed` placeholder (the dropdown cannot show a refetched typed rule's operands); the date-preset preview uses the browser clock (Task 10).

## Not machine-checked, and what is not evidence

- **Screen readers on the menus.** VoiceOver and NVDA were not run on the stream sort and filter menus; the tests check the DOM (`disabled`, the tooltip title) only.
- **Non-Chromium browsers.** The stream browser projects run Desktop Chrome only.
- **The corpus `visible-only` half is parity evidence only, not withholding evidence.** Task 13 counted 0 corpus sheets with a withheld key cell and 0 with an authored collapsed body row under `visible-only`, and no corpus XLSX has a drawing over text (4b's finding). So the corpus case shows engine and wire parity; withholding, C1 and the drawing cases are proven on the hand-built fixtures only. A drawing-bearing or withheld-cell corpus file is the owner's call.
- **The `COVER-B20` needle is vacuous.** It appears only in the plan; the fixture's B20 is the number `1`. The real check in `server/test/sheet-sort-audit.test.ts` is exact: no answer field equals `1` (with the other short values of `sortWithheld`).
- **Column and row default styles.** A cell under a column or row default `;;;` style with no `s` of its own is painted and served (the parser keeps no column or row style). Tiles and answers agree with each other, so nothing withheld leaks from what is painted; the gap is a divergence from Excel, which would paint such a cell empty. `paintedFormats` D2:D4, A6 and C6 carry it, and the audit does not ban them.
- **Mutation proofs are unrecorded.** Task 13 ran 15 mutations by hand (all caught; M7e only once eca7be34 added the older-`viewGen` list case) and reverted them, and Tasks 5, 7, 8, 10 and 11a showed RED by mutation or by restoring base source. No artifact or committed negative test records them; the logs were in an ephemeral scratchpad. Three mutants were equivalent and are not tested: key-view blanking of zero-width columns (every op on such a column is refused first), the unknown-coverage early return (the text view is already blank), and either half of the pivot-and-slicer strip alone (each covers the other).
- **Tests written after their code** (accepted by ledger rulings on RED shown by mutation or by restoring the base): Task 5's main `sheet-view-ops` file, Task 7's `value-lists.ts`, Task 8's view-refused test, Task 10's client tests, Task 11a's engine tests.
- **No browser resume case.** `setOffline` does not close Chromium's socket, so a resume during a sort is tested with a fake socket (`tests/stream/remote-sheet-sort.test.ts` "resume: sort and filter reject session-rotated, values and detect are re-issued") and at the server (`server/test/sheet-sort-filter.test.ts` "resume during an in-flight sort"), not in a real browser.
- **The million-row sort under load** is not explained (Status).
- **T8 p95** was not re-gated for 4c (recorded only).

## Follow-ups and open items

- **`server/README.md`** still describes 4b's sheet features only (`streamSheets`, line 784); it does not say a 4c server adds sort, filter and value lists. Not in this task's file list.
- **Root cause:** "a million-row sort" refused in the full server run (new, unexplained); `xlsx-viewer-worker-overflow` `SPEC guard %s` with `epoch-after-stage` (from 4b).
- **A withholding-bearing corpus XLSX** (a drawing over text, a withheld key cell, a collapsed body row), for the 4b ink oracle and the 4c corpus parity alike.
- **Column and row default styles** are not kept by the parser (above).
- **A harness `POST /drop`** would let a browser spec prove resume (Task 11).
- **The `stream` browser project's per-IP headroom.**
- **Search match counts:** `total` is the list's entry count, so the client cannot show a search's match count (Task 7).
- **A 6,304 B per-sort residual** is released only by GC (Task 8).
- From 4b, still open: tile-side CF reads merge-tail values (F20); copy's pool admission estimate; the `wireSelection` headroom; the worker-recycle stall.
