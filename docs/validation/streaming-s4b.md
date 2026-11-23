# Server streaming — Stage 4b validation (selection, resize and fit, find, copy and a11y text on streamed sheets)

Spec: `docs/superpowers/specs/2026-11-22-streaming-s4-design.md` (§9 4b; "As built (4b)" under §11.2). Plan: `docs/superpowers/plans/2026-11-23-streaming-s4b.md` (plan calls P1–P17; pre-flight findings F1–F22 and the ledger rulings in `.superpowers/sdd/2026-11-23-streaming-s4b/`). Branch `feature/streaming-s4b`, off main 1e2fbf4d (4a merged). Code head d30394a3. Measured 2026-11-24 on macOS, darwin-arm64, 14 cores, Bun 1.4.5.

## Status

Every 4b gate is green except three that were already red, or already known, before 4b:
- the eager size gate (index closure), red at the base 1e2fbf4d from the xlsx-cf work; 4b ends at the base figure, 426,556 B, Δ 0 (§Size gate);
- one vitest case, `tests/slideshow-screens.test.ts` "the drag hint works on a popup (other-realm) document and stays dismissed", which fails alone the same way (Node 26.13.1 exposes no `localStorage` without `--localstorage-file`), as 4a recorded at 53668e81; 4b touches neither the test nor `src/slideshow`;
- one stream browser case, `tests/browser/stream-text.spec.ts` "a11y=off: no screen-reader layer and no a11y request" (PPTX, a 15 s paint timeout), in the full stream-project run only; alone the file passes 10/10 and the case 5/5 with `--repeat-each=5`. It is the flake Task 12 saw (§Known flakes).

The compatibility matrix passes all eight legs, and `tests/stream` has 0 deleted lines since 1e2fbf4d. T8 tile p95, recorded and not gated, is 83.8 ms at dpr 1 and 56.0 ms at dpr 2 (GREEN against 120 ms).

## S4 §9 4b "done when"

| Line | Proven by |
|---|---|
| **selection parity:** scripted mouse and keyboard sequences give the same `GridSelection` as local mode on the snapshot | `tests/xlsx-stream-selection.test.tsx` "selection parity: keyboard" and "selection parity: mouse": the same scripts drive a local viewer over the servable snapshot (`sheetTextView` → `textDocument`) and a stream viewer over the same snapshot's wire layout, and compare every committed selection and `aria-activedescendant` (by cell part), across the merge A5:C6, hidden row 7, hidden column F and the authored 1 × 1 frozen panes. Browser: `tests/browser/stream-sheet.spec.ts` "shift+click selects a range". There is no freeze command (P17) and no drag-select, locally or on streams |
| find-hit sets equal local mode on the snapshot | `server/test/sheet-find.test.ts` "hit set and order equal local find on the deleted fixture" (local `WorkbookFind`, written apart from the server inputs, over `textGrid` with D20 and the merge tails deleted; it first shows the undeleted find hits D20), "full finds as local", "CF-effective display text". Browser: "find moves to a match and shows no hit text" |
| copy TSV equals local mode on the snapshot | `server/test/sheet-copy.test.ts` "copy TSV equals local on the deleted fixture" (a 3-range selection, byte for byte), "withheld cells are empty fields", "cfHiddenFormat copies as painted". Browser: "copy puts servable TSV on the clipboard" (`'1900\tE19\n\tE20\n2100\tE21'`, D20 covered) |
| copy caps refuse at n+1 and allow at n, per copy and per session | `server/test/sheet-copy.test.ts` "caps refuse at n+1 and allow at n, per copy" (3 ok, 4 refused `{copy-too-large, limit: 3, requested: 4}`, 6 cells with 3 withheld count 3) and "… and per session, measured against the resumed cap"; `server/test/sheet-a11y.test.ts` "focused counts 1 per non-empty cell", "the quota is shared with copy" |
| the off-viewport covered-cell fixture withholds from find, copy and a11y, and so does the rotated-box adversary | `server/test/sheet-text-audit.test.ts` "find, copy and a11y withhold the set from every scroll history" (`offViewportCovered`, `rotatedBox`, `unresolvedAnchor`, `cfHiddenFormat`, resized `absoluteOverB2`, each from three tile histories: none, every tile, only the covering tiles; answers deep-equal across histories); `server/test/sheet-text-view.test.ts` "covered cells are blank in the view" (rotatedBox A1:A3 blank, B1:C6 and A4:A6 kept); the find, copy and a11y "withheld cells …" tests |
| the wire audit is extended to text answers | `server/test/sheet-text-audit.test.ts` "no sentinel or withheld text in any text answer" (every text frame, payload and raw frame of find, select-all-used-range copy, unit a11y and fit over the hidden and covered fixtures), "find answers carry positions only"; `server/test/sheet-find.test.ts` "no text on the wire" |
| no local text index is built (spy) | `tests/xlsx-stream-selection.test.tsx` "no local text index" (0 `WorkbookFind` pages and 0 `copySelection` calls on a stream; a local control shows the spy sees a real page) |

Further 4b evidence by plan call:

| Call | Tests |
|---|---|
| P1 capability, F1 | `server/test/sheet-view-handshake.test.ts` (all); `tests/stream/remote-sheet.test.ts` "4b default HELLO (P1, F1)"; `tests/stream/manifest.test.ts` (`sheetView`); `tests/stream/remote-sheet-text.test.ts` "view present only with sheetView", "a resume that changes sheetView ends the session layout-changed"; compatibility legs (e)–(g) |
| P2–P5 text view | `server/test/sheet-text-view.test.ts` (9 tests, including "current-view coverage", "merge tails and undrawn cells are blank", "CF over a long column (F3) …", "full is the identity") |
| P9 fit | `server/test/sheet-fit.test.ts` (8 tests) |
| P11–P14 resize and view state | `server/test/sheet-resize.test.ts` (15 tests), `server/test/trace-tiles.test.ts` "a capture of a resized view traces", `server/test/cache.test.ts` (identity hash golden, sizes hashed); `tests/xlsx-stream-text.test.ts`; `tests/xlsx-stream-selection.test.tsx` "resize routes to the server", "key-repeat folds into one pending resize", "resize refusals speak P12a's strings"; browser "resizing a column moves the tiles", "a slid-under cell stops answering after resize" |
| P10 a11y, Open item 2 | `server/test/sheet-a11y.test.ts` (9 tests); `tests/xlsx-stream-a11y.test.tsx` (18 tests) |
| P13 lost replies | `tests/stream/remote-sheet-text.test.ts` "find is re-issued on resume; copy, a11y and resize reject session-rotated and are never re-sent"; `tests/xlsx-stream-text.test.ts` "a rejected resize re-fetches sheet-layout", "stale-view on copy, a11y or fit re-fetches sheet-layout and rejects stale-view" |
| P15 notices, F15 | `tests/xlsx-stream-a11y.test.tsx` "notices: …", "copy off says text unavailable once; a11y on still fills the grid"; `tests/xlsx-stream-selection.test.tsx` 'a truncated answer says "Showing the first 5,000 matches."', "a hit hidden by the filter is said, never inspected (P15)" |
| P16 ink oracle | `server/test/sheet-ink-oracle.test.ts` "ink-withheld cells are blank in the text view" (fixtures; see §Not machine-checked for the corpus case) |

**Mutation proofs (Task 14).** Each mutation was applied, run and reverted, and none was committed: an identity text view fails the audit, the battery and the fixture ink oracle; a `text` key or a non-numeric hit on find results fails "find answers carry positions only"; unknown coverage not blanked fails the audit and the battery; no measured spill fails the fixture ink oracle. The corpus ink oracle passed under both the identity-view and the no-spill mutation (below).

## Gates

Run one at a time at d30394a3 (code; the docs commit after it changes no code) on the same machine, with nothing else of this task running. The 1-minute load from other processes was about 3.5–4.4 throughout.

| Gate | Command | Result |
|---|---|---|
| server | `cd server && bun test` | **1,541 pass, 0 fail**, 105 files, 646,223 `expect()` calls, 855 s |
| client | `bun scripts/fetch-corpus.ts`, then `bunx vitest run` | 773 of 774 files; 15,872 pass, **1 fail**, 65 skipped, 1 todo (15,939), 892 s. The fail is `tests/slideshow-screens.test.ts` (pre-existing; alone: 9 pass, 1 fail, the same case) |
| browser, stream | `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npm run test:browser:stream` | 51 passed, **1 failed**, 1 skipped (the existing no-codec media case), 3.7 min. The fail is `stream-text.spec.ts:223` "a11y=off …" (flake). Alone: `npx playwright test --config playwright.viewer.config.ts --project=stream tests/browser/stream-text.spec.ts` 10 passed; `… stream-text.spec.ts:223 --repeat-each=5` 5 passed |
| browser, viewer (unmodified) | `IV_HARNESS_PORT=4391 OV_STREAM_PORT=4281 npx playwright test --config playwright.viewer.config.ts --project=chromium tests/browser/interactive-viewer` | **67 passed**, 2.4 min |
| typecheck | `npx tsc --noEmit -p .` | clean (exit 0) |
| typecheck, server | `npm run -s typecheck:server` | clean (exit 0) |
| lint | `npm run -s lint` | clean (exit 0) |
| eager gate | `bun run gates:core` | core closure ok (300 B), engine modules ok, core purity ok; index closure FAIL as at the base (§Size gate) |
| compatibility | `bun scripts/stream-s4/compat-matrix.ts --base 53668e81 --base4a 1e2fbf4d --tmp <scratchpad>/runs` | 8 PASS, exit 0; `git worktree list` byte-identical before and after |
| T8 (not gated) | `bun scripts/stream-s4/tile-p95.ts` | GREEN, exit 0 (§T8 p95) |
| append-only | `git diff --numstat 1e2fbf4d -- tests/stream` | 0 deletions: `manifest.test.ts` +24, `remote-sheet-text.test.ts` +261 (new), `remote-sheet-view.test.tsx` +26, `remote-sheet.test.ts` +27, `sheet-wire.test.ts` +233 |

## Size gate (T11)

```
ok   core closure 300 B over 1 chunk(s) ≤ 38720
ok   engine modules in core closure: none
ok   core purity: no react imports, no document global
FAIL index closure 426556 B (baseline 417436, Δ 9120, recorded 3193, limit 420645)
FAIL index module set differs from baseline ± recorded changes:
       + src/xlsx/rowFont.ts
```

- **The base was already red.** Task 3 measured 426,556 B at the base (1e2fbf4d, with the same `rowFont.ts` module-set line, from xlsx-cf). The gate is not in CI. No limit or baseline was changed.
- **4b's delta is 0 B** (426,556 − 426,556). Per task it moved +1, −3, +2, +1, +1, −1 and −1 B; these are hashed lazy-chunk names in the index's `dynamicImports`, and no module was added to the index closure.
- New client code is in lazy chunks only: `streamText.ts` and `RemoteSheetView.tsx` in `RemoteSheetView-*.js`; `InteractiveSheetDoc.tsx` and `streamA11y.ts` in `InteractiveSheetDoc-*.js` (`dist/.module-map.json`). `dist/core.js` stays pure and engine-free.

## Compatibility matrix

`bun scripts/stream-s4/compat-matrix.ts --base 53668e81 --base4a 1e2fbf4d --tmp <scratchpad>/runs`, exit 0, at d30394a3:

```
PASS (a) base client x current server: no sheetMode on the raw manifest; unit (raster), 1 search hit(s), copy 22 chars, 1 print page(s) 336x20
PASS (b) current client x current server: sheetMode tiles, sheet-layout decoded (0 row / 0 col intervals), tile (0,0) decoded 512x512, manifest validated on grid500x100
PASS (c) current client (opted in, HELLO carried sheet-tiles@1) x base server: accepted, no sheetMode, unit arrives (raster)
PASS (d) 17 features (16 distinct unknowns + sheet-tiles@1) closed 4000 = unsupported-version; the 16-feature control was accepted
PASS (e) 4a client x current server: 4a validateManifest accepted it, no sheetView on the raw manifest, 4a decodeSheetLayout decoded sheet-layout (1 part(s), no session-size keys), tile (0,0) decoded 512x512
PASS (f) current default client (HELLO sheet-view@1) x 4a server: units mode (no sheetMode), 2 unit(s), 1 search hit(s), copy 22 chars
PASS (f') current client (streamSheets 'tiles', HELLO sheet-tiles@1) x 4a server: tiles mode, doc.sheets.view undefined, tile (0,0) decoded 512x512
PASS (g) current client (streamSheets 'units', no HELLO features) x current server: units mode, Stage 3 manifest keys, 2 unit(s), copy 22 chars
```

Legs (a)–(d) are 4a's, unchanged. (e)–(g) are 4b's (Task 13): (e) decodes a real `sheet-layout` with the 4a client's own decoder (F2). The script removed both base worktrees afterwards; `git worktree list` was the same before and after.

## T8 p95 (recorded, not gated)

`bun scripts/stream-s4/tile-p95.ts`, R3 protocol as in 4a (grid500x100, `visibleMark: 'fixed'`, forensic on, one tile per request, serial, zoomStep 5, warm document cache, cold render cache). 4b changes only the tile geometry's sizes, so T8 is not re-gated; this is the one recorded run.

| Measure | dpr 1 (260 tiles) | dpr 2 (1,000 tiles) |
|---|---|---|
| R3 p50 / **p95** / max (ms) | 70.1 / **83.8** / 859.3 | 50.0 / **56.0** / 859.1 |
| client send → `tiles-done` p50 / p95 / max (ms) | 71.4 / 85.3 / 860.3 | 51.0 / 57.1 / 860.3 |
| loopback (request → `drawImage`, Bun + napi) p50 / p95 / max (ms) | 72.4 / 86.5 / 846.9 | 52.0 / 58.8 / 868.5 |

- All 1,260: R3 p95 75.4 ms; 4 worker recycles; 0 R3 intervals above the client's own interval. Verdict line: `T8 (p95 ≤ 120 ms at dpr 1 and dpr 2): GREEN`.
- `--batch 64` (dpr 2, 1,000 tiles, cold): 26,965 ms, 37.1 tiles/s, 3 recycles; per-tile t0→t1 p50 927.9, p95 1,961.6, max 2,382.1 ms (queue wait included).
- Load (1/5/15 min): start 3.95/4.09/4.14, after the serial run 3.21/3.83/4.04, end 4.35/4.07/4.11. 4a recorded 83.6 and 56.0 ms at a load of about 6.
- The max is still the worker-recycle stall (about 0.86 s, 4 tiles in 1,260).

## Known flakes and environment notes

Seen during 4b; listed, not fixed. Each passed when its file ran alone.
- `server/test/mark-geometry.test.ts` "registration time: uncropped zoom 1.5 registers in <= 15 s" and "zoom then a 10 % crop": wall-clock bounds under parallel load (Tasks 3, 5).
- `server/test/print-wire.test.ts` "print then print-cancel back to back …" and "two cancelled jobs back to back …": cancel races under load (Task 7).
- `tests/browser/stream-text.spec.ts` "a11y=off: no screen-reader layer and no a11y request" (PPTX): a 15 s paint timeout in a full stream-project run (Task 12).
- `tests/xlsx-viewer-main-controls-accounting.test.ts` "keeps membership and independent reader copies charged through close": a heap-delta assertion under parallel load (Task 10).
- `tests/xlsx-viewer-worker-overflow.test.ts` "SPEC guard epoch-after-stage preserves old rows and raw ownership": failed twice alone right after a full run, then passed at the base and three times after (Task 11). **Root cause not found.**
- `tests/slideshow-screens.test.ts`, one case, fails under Node 26.13.1 (no `localStorage` in that realm), as at 4a's base; not 4b's.
- In this task's runs only `stream-text` "a11y=off" recurred (above). Every other listed flake passed in the full runs.

## Rulings and findings as built

The full text is in the S4 design, §11.2 "As built (4b)", and the dated parent and S3 amendments.
- **Owner, 2026-11-23:** merge tails are blanked in text answers only (`sheetTextView`), not in `projectXlsx` (S4 §10 item 1 amended); tile-side CF still reads merge-tail values (F20). Focused a11y fetches after a 400 ms rest, cached per `(sheetId, viewGen, cell)`.
- **F1:** `streamSheets: 'tiles'` sends exactly `['sheet-tiles@1']`, so it gets 4a and no 4b text; unset sends `['sheet-view@1']`; `'units'` sends none. This deviates from plan P1's "tiles sends both".
- **F2:** session-size layout keys go only to `sheetView` sessions, only when non-empty.
- **F3 (amended):** the CF text scope is prepared as the local viewer prepares it; a refused scope blanks every CF range (fail closed).
- **F4:** fit uses the base style (local parity), with no CF scope.
- **F8:** `sheet-copy` and `sheet-a11y` carry `viewGen` and are refused `stale-view` before any count or audit.
- **F15:** find `total` = hits sent, `truncated` = more existed, the viewer says "Showing the first 5,000 matches.", and there is no cursor (spec deviation).
- **F17:** §6.3 admission is deferred to 4c (deviation from "required"); the per-axis 4,096 and 16 MiB process caps stand in.
- **F19:** `sheet-view@1` extends §6.10.
- **Select-all:** a whole-sheet select-all copy is refused `copy-too-large`, as locally (parity ruling); a full-height column copies in one answer.
- **Drag-select:** none exists locally or on streams; the browser spec is named "shift+click selects a range".
- **Audit:** a unit a11y request with no servable text, a focused empty cell, and an n = 0 copy write no audit row.
- **Off-mode text ops** are charged one `'text'` check before `unsupported-op`; `'sheet'` is charged before the one-in-flight check.

## Not machine-checked

- **Screen readers.** VoiceOver and NVDA were not run on the stream grid; the a11y tests check the DOM text and the request pattern only.
- **Non-Chromium browsers.** The stream browser project runs Desktop Chrome only.
- **Real clipboards beyond Chromium.** The browser copy spec reads Chromium's clipboard; other browsers and OS clipboards were not tried.
- **The corpus ink oracle is not evidence.** `server/test/sheet-ink-oracle.test.ts` "… on the corpus" visits 30 files and 46 sheets, but no corpus XLSX has a drawing over text (the one drawing part, `poi-66365.xlsx` `drawing1.xml`, is an empty `wsDr`), so it checks 0 cells and passes even with an identity text view. A drawing-bearing corpus XLSX is needed; adding one is the owner's call. The fixture oracle is real evidence.
- **Wrapped and rotated row fits.** Server fit has no browser layout height and uses the local wrapped-lines estimate; parity was checked against the same estimate, not a browser fit.

## Follow-ups and open items

- **A drawing-bearing corpus XLSX** for the ink oracle (above).
- **`streamSheets: 'tiles'` hosts** get no 4b text until they drop the option (F1); stated in `server/README.md` and the option's doc comments.
- **Tile-side CF reads merge-tail values** (F20); only the owner can widen.
- **Flake root cause:** `xlsx-viewer-worker-overflow` "SPEC guard epoch-after-stage".
- **Task 4 review was thin** (ledger note): the session view state, `sheet-resize` and sized tiles deserve extra attention in the final review.
- Deferred minors from the ledger:
  - `blankCell`/`holdsContent` duplicated from `server/src/sheet/snapshot.ts` in `text-view.ts` (Task 1);
  - `decodeSheetView`'s key check allows absent `reqId`/`type` when used standalone (Task 2);
  - server fit has no `layoutHeight`: wrapped or rotated rows use the estimate (Task 5);
  - copy's pool admission estimate (16 MiB + layout) is below `serializeSelection`'s 24 MiB working budget, so it can under-admit by up to 8 MiB (Task 7);
  - `copy: false` + `textA11y: true` (unit mode) has no test in `sheet-a11y.test.ts` (Task 8);
  - `wireSelection`'s 1 KiB envelope headroom assumes token + envelope < 1 KiB (Task 9);
  - the fallback resize-refusal wording for unlisted codes is implementer-chosen (Task 10);
  - `STREAM_A11Y_COARSE_CELLS` duplicates `GridSemantics`' coarse threshold (Task 11);
  - the corpus ink oracle visits 0 cells (Task 14; above);
  - closed: `textInputs`' store and `withUnavailableFormulas` (Task 5) — find uses the wrapped `findStore()`, copy the plain store as locally, and the function moved to one React-free module (8c331229).
- From 4a, still open: B1 (units-mode XLSX unit over 8192 px), the worker-recycle stall, R10 (finding a capture's tile window automatically).
