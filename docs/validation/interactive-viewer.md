# Interactive viewer: functional and complexity evidence (S10.1)

**Branch:** `feature/iv-s10-1`, cut from `feature/interactive-viewer` at `cfa5cb5` (everything through S5.4, S9.2 and S8.4 merged). **Run date:** 2026-10-12. **Machine:** Apple M4 Pro, macOS 26.5.2. **Engines:** Playwright 1.64.0 with Chromium headless shell build 1248 and Playwright WebKit build 2370. WebKit results are labeled `webkit-playwright` and are **not** Safari evidence (spec §12.2).

This page is the S10.1 functional/complexity matrix only. Performance and memory receipts belong to S10.2. Screen-reader, device and capability-fallback receipts belong to S10.3. Their acceptance cases are listed here as **not-run**, with the reason. They stay blocking until those packages deliver them. A passing vitest suite counts as code-test evidence. Watching the UI is not a receipt.

## How to reproduce

```sh
bun install --frozen-lockfile
bunx tsc --noEmit
bunx vitest run                                                    # whole suite
bunx vitest run tests/xlsx-viewer-complexity.test.ts               # complexity audit
bunx playwright test --config playwright.viewer.config.ts          # real browsers, both source paths
bun run build                                                      # eager dist/index.js size
bun scripts/fetch-corpus.ts && bun scripts/golden-corpus.ts compare --threshold 0 --max-ratio 0
```

The browser suite mounts the shipped `<OfficeFile interactiveSheets showSearch>` on the dev-only page `src/debug/viewer-e2e.html`. With `?worker=1` it also supplies a real module `Worker` running `src/worker/worker-entry.ts` (through `src/debug/viewerWorker.ts`), so a paged admission opens worker-backed, exactly as a host's `viewerWorker` factory would. The instruments sit at public boundaries only:

- the worker port: commands by kind and read op, and transferred bytes;
- the clipboard: the exact text written;
- the `onFrame` stats: bitmap bytes, budget and record counts;
- the DOM.

Packages are true XLSX files, routed to the page as `/__fixtures/<name>`.

## Gate results (this branch head)

| Gate | Result |
|---|---|
| `bunx tsc --noEmit` | pass (no output) |
| `bunx vitest run` | **186 files passed; 5,362 tests passed, 2 skipped, 1 todo** at the branch head (review round 1 included). An earlier run reported 184 files / 5,279 tests because two corpus-dependent files skipped before `corpus/` was populated. Base `cfa5cb5` without corpus: 180 files / 5,265 tests, all passing |
| `bunx playwright test --config playwright.viewer.config.ts` | **27 passed, 1 skipped** (the skip is the Chromium-only long-task case on WebKit, by design). This includes 20 S10.1 scenarios, 10 each on Chromium and WebKit |
| `bun run build` | pass; eager `dist/index.js` **38.32 kB gzip** (vite report), unchanged from base 38.32 kB, ceiling 38.72 kB |
| Golden compare `--threshold 0 --max-ratio 0` | **every XLSX unit passes** (poi-47889 ×3, 56295, 59021, 59746_NoRowNums, 62272, 64508, 66365, link-external-workbook-a, xlsx-jdbc). 6 DOCX/PPTX units fail identically on base `cfa5cb5`, checked in a temporary base worktree: poi-55733.docx, poi-picture-transparency.pptx ×3, pydocx-blk-inner-content.docx, pypptx-test_slides.pptx. They are pre-existing and environmental, and no DOCX/PPTX/render file changed on this branch |

## Reproduced gaps and fixes

Each gap was first reproduced in a real browser, then pinned by a failing test (RED for the intended reason), then fixed.

| # | Gap (where reproduced) | Requirement | Fix | Regression tests |
|---|---|---|---|---|
| G1 | Worker-backed F2 (100,010 rows): every header sort/filter refused "Source range is incomplete." forever, and no worker command was issued. The controls port snapshots structure once when the sheet opens. The worker answered `completeness: 'loading'` while its import ran, and the snapshot was never refreshed. | STORE-04, LIFE-03, SEM-01 | `viewerControlsClient.ts` re-queries a structural snapshot that was still loading before resolving a range. | `tests/xlsx-viewer-controls-structure.test.ts` (loading-first snapshot), browser F2 worker-backed scenario |
| G2 | F2 resident (1,000 rows) and worker-backed (100,010 rows): a header sort with no selection resolved A1:H6, so the sort and record counts covered **5 body rows**. The shipped object/paged/remote stores continue `scanPopulated` by iteration without `nextCursor`. Range detection read one page and treated the missing cursor as "exhausted", so a full strip page proved a row blank. The S7.1 suite's synthetic store pages by cursor, which masked the bug. | SEM-01, STORE-04 (§6.1 current region) | `rangeDetection.ts` Scanner reads one page ahead and issues its own continuation token; gallop strips are sized in cells (lines × span, capped at 65,536). Observed during the investigation (not pinned; the committed test pins < 200 pages): **60,006 → 66** store pages on a 20,001-row table. On worker-backed F2 at 100,010 rows the investigation counted about 597 scan reads, instead of one probe per row. | `tests/xlsx-viewer-range-continuation.test.ts` (whole-table range, body active cell, page-count bound), browser F2 scenarios |
| G3 | F2 (both paths): after Go To A1, the funnel on column D refused "Contiguous selection does not contain the active cell". The header gesture moves the active cell to its column but kept the old single-cell selection. S7.1 detection deliberately refuses that combination; its own test is unchanged. | SEM-02, A11Y-02 | `viewerControlsClient.ts` `rangeRequest` drops a single-cell selection for header gestures. | `tests/xlsx-viewer-controls-structure.test.ts` (header gesture after A1), browser F2 scenarios |
| G4 | Chromium/WebKit, 400 px viewer pane: the column popup is anchored to the pane bottom and grew past the pane top. There `overflow:hidden` clipped Sort ascending/descending, Clear column and Clear all filters under the navigator, so they could not be clicked. | A11Y-02, SEM-09 | `GridSemantics.tsx` popup is bounded to `max-height:100%` of the pane and scrolls. | browser "column filter popup stays inside a short viewer pane…" |
| G5 | Chromium, F7-shaped sheet: **one** authored 0.75 pt row (2 device px at DPR 2) in an ordinary 41-row viewport made the whole owned grid coarse: no column headers, rows or gridcells, only "Zoom in or inspect a cell for detail". The painter meanwhile kept every other row in detail. | A11Y-01, RENDER-01 | `GridSemantics.tsx`: only lines under 4 device px go unowned. The whole tree is coarse only when the window caps are exceeded or no detail line remains. A navigated cell on the tiny row stays the exact active descendant. | `tests/xlsx-viewer-a11y-subdetail.test.tsx`, browser F7 scenario (Row 11 owned, Row 10 summarized) |

No reproduced gap fell in `src/xlsx/view/sort.ts` or `scheduler.ts`. Both files are owned by the concurrent S8.6 work and were not edited.

## Fixture register

| Name | Source | Scale (labelled) | Path used |
|---|---|---|---|
| `f1-small.xlsx` | committed S0.2 | F1 corners + 200 sampled cells | resident |
| `f1-paged.xlsx` | `writeFixturePackage('F1', { nnz: 120000 })` | F1 at 120,000 NNZ (crosses the 100,000-cell object cap) | worker-backed |
| `f2-1k.xlsx` | `writeFixturePackage('F2', { rowCount: 1000 })` | reduced F2 | resident |
| `f2-paged.xlsx` | `writeFixturePackage('F2', { rowCount: 100010 })` | reduced F2 (crosses the 100,000-row cap) | worker-backed |
| `f3sparse-small.xlsx` | committed S0.2 | F3 sparse-wide 256 × 512 | resident |
| `f3-paged.xlsx` | `writeFixturePackage('F3', { rowCount: 8 })` | dense 8 × 16,384 = 131,072 cells | worker-backed |
| `f4-12k.xlsx` | `writeFixturePackage('F4', { rowCount: 12000 })` | reduced F4, > 10,000 unique values | resident |
| `f5-structure.xlsx` | hand-built in `tests/browser/viewer-e2e.ts` | giant merge D1:D1048576, fixed array C30:C32 | resident |
| `f7-two-sheets.xlsx` | `writeSheetsPackage` | sheets A/B, hidden row 8, 0.75 pt row 10 | resident |
| F6 full-axis provider | `createFixture('F6')` store | all 2^34 addresses, analytic | vitest only (never enumerated) |

Full-scale F2 (1,048,576 × 8), F3 (256 × 16,384) and F4 (1M strings) browser runs are **not-run** here. They belong to the S10.2 release job. No scenario on this page claims full-scale behavior or any timing.

## Evidence matrix by acceptance case

Status legend: **pass** = executed on this branch and passing; **not-run** = not executed, with the reason. No case failed at the end of the run. The cases named in a "not-run" row are listed again under *Blocking*.

Browser shorthand: **[B:C+W]** = Playwright scenario passing on Chromium and Playwright WebKit. **res** = resident (object) path, **wkr** = worker-backed paged path. Suite files are under `tests/`.

| Case | Fixture(s) | Evidence | Status |
|---|---|---|---|
| AC-CAP-01 | F1, F6, synthetic boundary rows | `xlsx-viewer-contracts.test.ts`; `xlsx-viewer-complexity.test.ts`: DOM active-descendant ids 2^31, 2^32 and 2^34−1 end to end. Browser F1 res+wkr: `aria-activedescendant` id 17179869183 **[B:C+W]** | pass |
| AC-CAP-02 | F1, F3, F7 | `xlsx-viewer-scroll-window.test.ts`, `xlsx-viewer-navigation.test.tsx`, `xlsx-viewer-selection.test.ts`. Browser: F1 Go To XFD1048576 with header/row label agreement; F3 recenter, header/cell `aria-colindex` agreement and edge ArrowLeft after recenter; F7 Original-cell Go To on a hidden row opens the read-only inspector without moving the active descendant **[B:C+W]** | pass (zoom/DPR browser leg partial: zoom 125% exercised, DPR change not exercised) |
| AC-CAP-03 | F2/F3/F4 quota variants | `xlsx-viewer-import-limits.test.ts`, `xlsx-viewer-source-pages.test.ts` (typed quota/incomplete/cancel) | pass (unit); real-browser quota-denial run not-run (S10.3) |
| AC-STORE-01 | F8 | `xlsx-viewer-adapter-parity.test.ts`, `xlsx-viewer-values.test.ts`, `xlsx-viewer-source.test.ts`: object and paged adapters agree on values/errors/cache/formula/date/raw UTF-16; pending ≠ blank. Browser wkr: unknown rows paint "Loading…" (`xlsx-viewer-paged-viewer.test.tsx`) | pass; **native Excel checkpoint not-run** (see below) |
| AC-STORE-02 | F1–F4 malformed variants | `xlsx-viewer-import.test.ts`, `xlsx-viewer-import-limits.test.ts` | pass (unit); peak/cleanup ledger in a browser → S10.2 |
| AC-STORE-03 | F5 | `xlsx-viewer-structure.test.ts` (pure index). Browser F5: giant narrow-column merge and fixed array queried before consumers, child selection allowed **[B:C+W]** | **partial**: parser/drawing/dynamic-spill wiring (S3.1, unticked) is not delivered. Any sheet with a drawing or dynamic array reports incomplete structure (`source.ts` coverage diagnostics), so sort/filter there always refuse as incomplete. Blocking |
| AC-STORE-04 | F8, F7 | `xlsx-viewer-values.test.ts`, `xlsx-viewer-worker-races.test.ts`, `xlsx-viewer-saved-notices.test.tsx`; G1 regression | pass |
| AC-VIEW-01 | F2, F7 | `xlsx-viewer-projection.test.ts`, `xlsx-viewer-sort.test.ts` | pass |
| AC-VIEW-02 | F1, F2 | `xlsx-viewer-cap-regression.test.ts`, `xlsx-viewer-viewport-integration.test.tsx`. Browser F2 wkr: last body row sorted to display row 2 shows "original G100009" with source id preserved **[B:C+W]** | pass |
| AC-VIEW-03 | F5, F6, F7 | `xlsx-viewer-axes.test.ts`, `xlsx-viewer-scroll-window.test.ts`; `xlsx-viewer-complexity.test.ts` (no row lane, bounded canvases). Browser F1: every canvas ≤ viewport × DPR, bounded scroller extent **[B:C+W]** | pass |
| AC-VIEW-04 | F1, F8 | `xlsx-viewer-sheet-zoom.test.ts`, existing zoom/legacy suites. Browser F7: zoom retained across a committed sort, reset by reload (document replacement) **[B:C+W]** | pass |
| AC-SEM-01 | F2, F5 | `xlsx-viewer-filter.test.ts`, `xlsx-viewer-projection.test.ts`. Browser F2 res+wkr: header protected, value filter, Cancel publishes nothing, Clear all restores **[B:C+W]** | pass |
| AC-SEM-02 | F2, F4, F7 | `xlsx-viewer-sort.test.ts` (truth table, 1–8 keys, host collator) | pass (unit); 8-key browser leg not-run |
| AC-SEM-03 | F5, F8 | `xlsx-viewer-sort-range.test.ts`, `xlsx-viewer-range-continuation.test.ts` (G2) | pass |
| AC-SEM-04 | F5 | `xlsx-viewer-operation-guard.test.ts`. Browser F5: merge D1:D1048576 refuses both a region sort and an explicit A2:B7 sort ("Operation refused: merge"); view unchanged **[B:C+W]** | pass for merge/fixed array; drawing legs depend on S3.1 wiring (see AC-STORE-03) |
| AC-SEM-05 | F5 | `xlsx-viewer-operation-guard.test.ts` (pure guard) | pass (unit); end-to-end drawing-anchor leg not-run: needs S3.1 drawing wiring. Blocking |
| AC-SEM-06 | F2, F4, F8 | `xlsx-viewer-filter.test.ts`, `xlsx-viewer-canonical-text.test.ts`, `xlsx-viewer-filter-ui.test.tsx` | pass |
| AC-SEM-07 | F8 | `xlsx-viewer-date-filter.test.ts` | pass (unit); native 1900/1904 workbook checkpoint not-run |
| AC-SEM-08 | F4 | `xlsx-viewer-value-index.test.ts`, `xlsx-viewer-value-search-keys.test.ts`. Browser F4 (12,000 values): "Showing first 10,000", search finds item-11000 beyond 10k, searched Select All selects 1, first Select All covers all 12,000 **[B:C+W]** | pass (100k/1M universe at scale → S10.2) |
| AC-SEM-09 | F4, F7 | `xlsx-viewer-filter-ui.test.tsx`. Browser F4/F2: reopen retains applied choice, Cancel publishes no generation (worker `filter` count unchanged) **[B:C+W]** | pass |
| AC-SEL-01 | F7 | `xlsx-viewer-selection.test.ts` ([0,3,1,2] Shift cases), `xlsx-viewer-copy.test.ts` | pass (unit) |
| AC-SEL-02 | F7 | `xlsx-viewer-selection.test.ts`, `xlsx-viewer-a11y.test.tsx`. Browser F7 hidden-row inspection keeps the active descendant **[B:C+W]** | pass |
| AC-COPY-01 | F7, F8 | `xlsx-viewer-copy.test.ts`, `xlsx-viewer-policy.test.tsx`. Browser F1 res+wkr: clipboard receives exactly "A1048576" after find **[B:C+W]** | pass |
| AC-A11Y-01 | F1, F3, F7 | `xlsx-viewer-a11y.test.tsx`, `xlsx-viewer-a11y-subdetail.test.tsx` (G5). Browser: `aria-rowcount` 1048577 / `aria-colcount` 16385, bounded owned tree (< 2,000 cells), header/cell index agreement **[B:C+W]** | pass (code); **desktop screen-reader and iPad VoiceOver receipts not-run** (S10.3). Blocking |
| AC-A11Y-02 | F7 | `xlsx-viewer-events.test.tsx`, `xlsx-viewer-filter-ui.test.tsx`, `xlsx-viewer-sort-ui.test.tsx`; G4 browser regression | pass (code); actual keyboard/touch/focus receipts on Safari/iPad not-run (S10.3). Blocking |
| AC-LIFE-01 | F7 | `xlsx-viewer-find-races.test.tsx`, `xlsx-viewer-worker-races.test.ts`, `xlsx-viewer-paged-viewer.test.tsx`. Browser F7: sheets A/B both at their first revision keep their own sort and data across switches; views never bump fontGen; sort/filter keep zoom **[B:C+W]** | pass |
| AC-LIFE-02 | F7 multi-sheet | `xlsx-viewer-lifecycle.test.tsx` (F7 plateau, close returns ledgers to baseline) | pass (unit); real retained-memory/storage receipts → S10.2. Blocking |
| AC-SEARCH-01 | F1, F7, F8 | `xlsx-viewer-find.test.ts`, `xlsx-viewer-find-races.test.tsx`. Browser F1 res+wkr: find reaches the never-painted A1048576 with source id 17179852800 **[B:C+W]** | pass |
| AC-RENDER-01 | F1, F3, F5, F6 | `xlsx-viewer-layout.test.ts`, `xlsx-viewer-overview.test.ts`, `xlsx-viewer-complexity.test.ts` (F6 window visits ≤ 2 × visible cells; 25 % tiny-size frame is analytic, 0 per-cell visits); G5 | pass |
| AC-RENDER-02 | F7, F8 | `xlsx-viewer-selection-overlay.test.ts`, `xlsx-viewer-viewport-integration.test.tsx`, static print/export suites | pass (unit) |
| AC-PERF-01 | F1–F8 | S10.2: controlled Chrome/Safari/iPad p50/p95 receipts | **not-run** (out of S10.1 scope; Safari/iPad also blocked on hardware/automation). Blocking |
| AC-PERF-02 | F2–F7 | complexity part: `xlsx-viewer-complexity.test.ts` (see below); per-pool memory receipts → S10.2 | complexity **pass**; memory **not-run**. Blocking |
| AC-BACKEND-01 | — | TS resident baseline is the only backend. S11/S12 were not executed, and the S11 allowance ("record not-run when the baseline already meets targets") needs the S10.2 receipts first | **not-run** (pending S10.2 and S11). Blocking |
| AC-REG-01 | F8, corpus | typecheck, full vitest, build, golden compare (all XLSX units pass; 6 non-XLSX failures pre-existing on base) | pass (code); native source/checkpoint evidence not-run |

**Counts (34 acceptance cases):**

| Status | Cases | Which |
|---|---:|---|
| pass | 17 | AC-CAP-01, STORE-04, VIEW-01–04, SEM-01, SEM-03, SEM-06, SEM-09, SEL-01, SEL-02, COPY-01, LIFE-01, SEARCH-01, RENDER-01, RENDER-02 |
| pass with a not-run leg | 9 | AC-CAP-02, CAP-03, STORE-01, STORE-02, SEM-02, SEM-04, SEM-07, SEM-08, REG-01 |
| partial, blocking | 6 | AC-STORE-03, SEM-05, A11Y-01, A11Y-02, LIFE-02, PERF-02 |
| not-run, blocking | 2 | AC-PERF-01, BACKEND-01 |
| fail | 0 | |

## Complexity audit (spec §3 invariants)

The instruments live in `tests/xlsx-viewer-complexity.test.ts`.

What is instrumented, and what is not:

- **Allocation ledger:** proxies over every global typed-array constructor and `Array(n)`, plus `Map#set` and `Set#add` on collections constructed by officeview modules. jsdom's and React's own maps are excluded, because a real browser does not run them in JS.
- **Retained heap:** the V8 heap after a forced full collection, measured before and after each warmed-up window while the viewer stays mounted. Plain objects and literal arrays have no constructor to proxy, so this is what bounds them. Any per-row/per-cell object table or per-row array set that the window keeps reachable exceeds the 4 MiB bound.
- **Self-checks:**
  - The ledger sees a row-sized `Uint32Array(1048576)` and 60,000 inserts into an `axes.ts`-built `Set`.
  - The review probe (a 65,536-entry plain-object cell table plus 1,024 plain row arrays) records 0 inserts and 0 typed elements, but retains about 7 MiB, which is over the bound.
- **Residual (not bounded here):** transient garbage created and freed inside a window, and per-coordinate work that allocates nothing (pure loops). Visit counts cover those only at the instrumented boundaries: worker read rects and F6 provider visits.

| Anti-pattern | Instrument | Observed | Bound (why) |
|---|---|---|---|
| Dense R×C allocation / row lanes | largest typed array during corner navigation, giant-merge paint, worker-backed sort, F6 window | 0 / 32 / 32,768 elements (the 32,768 is the visibility bitset words) | < 1,048,576 (one row lane); a cell lane would be 2^34 |
| Per-blank-row/cell objects | officeview Map/Set inserts; retained heap after warmed-up windows | 34 – 10,231 inserts; retained 0 – ~2.1 MiB (the largest is the F6 25 % result holding up to 2,048 coarse groups) | < 50,000 inserts (one per row would be ≥ 1,048,575); < 4 MiB retained (the review probe retains ~7 MiB) |
| All-column scans per row | every worker `read rows` rect | ≤ 1,024 columns × ≤ 4,096 rows per read | never 16,384 columns or the whole axis |
| Full-sheet canvas / CSS | largest canvas backing store; scroller extent (browser) | ≤ 4096² px; canvases ≤ viewport × DPR; scroller < 1M × row height | a full-sheet canvas would be 2^34 px tall |
| Merge coverage maps | inserts while painting and sort-guarding D1:D1048576 | < 50,000 | a coverage map holds 1,048,576 entries |
| Per-command clone/reparse | worker `open` commands, kernel `parse` spy | 1 open, 0 parses for open + navigation + sort | one per document |
| Lossy bitwise coordinate IDs | DOM `aria-activedescendant` cell ids | 2147483648, 4294967296, 17179869183 exact | signed/32-bit packing gives negative or 0 |
| Range detection cost (G2) | store page count for a 20,001-row region | pinned bound only (66 pages observed in the investigation, 60,006 before) | < 200 (strip pages, not per-row probes) |
| F6 geometry window | provider `cellAt` + `isPopulated` visits | ≤ 2 × visible cells at XFD1048576; 0 per-cell visits for the 25 % tiny-size frame | never proportional to 2^34 |

## Native Excel, Safari and devices

- **Native Excel (AC-STORE-01/SEM-07/REG-01 native legs): not-run.** One automation attempt was made: `osascript` drove Microsoft Excel to create a workbook (number, `DATE()+0.5` with a datetime format, `=1/0`, a cached formula, a hidden row, Unicode text, XFD1048576) and to save it as `.xlsx`. The script reached the save step and then blocked for 120 s on a dialog that needed a user, most likely Excel's sandbox file-access grant: an Excel lock file appeared, but no workbook was written. A `quit saving no` was then sent to Excel. Excel process 699 stayed running afterwards, possibly the user's own session, so the user should check Excel for a leftover dialog or an untitled workbook (the coordinator has told the user). No native receipt is claimed. **Any future native automation must run under a dedicated macOS user account or in a VM, never against the user's live Office session.** No further Excel attempts were made.
  - **Manual steps:**
    1. In Excel, build the cells above, plus a table (Insert ▸ Table over D1:E3, header row on). Save as `native-corner.xlsx` and close it.
    2. Reopen the file and confirm Excel shows no repair prompt. Take a screenshot.
    3. Copy the file to `tests/fixtures/xlsx-viewer/native-corner.xlsx`, record its SHA-256, and add a parity case like `xlsx-viewer-adapter-parity.test.ts`: object and paged adapters agree, cache/error/date/hidden/table values match the screenshot.
    4. Open it in the e2e page (`?src=`), with and without `?worker=1` (the latter needs a cap-crossing copy).
- **Safari desktop: not-run.** `safaridriver` exists, but remote automation requires a one-time `safaridriver --enable` (admin prompt) plus Develop ▸ Allow Remote Automation, which needs the user. The Playwright WebKit runs above are not Safari evidence.
- **iPad / physical devices: not-run** by user decision.

## Blocking for release

These are mandatory cases whose evidence is incomplete. Nothing here is hidden as a pass.

1. **AC-STORE-03 / AC-SEM-04 drawing legs / AC-SEM-05:** the S3.1 parser/drawing/dynamic-spill wiring is not delivered. A sheet with any drawing or dynamic array has permanently incomplete structure, so its sort/filter refuse.
2. **AC-A11Y-01 / AC-A11Y-02:** the desktop screen-reader, iPad VoiceOver and real keyboard/touch/focus receipts (S10.3).
3. **AC-PERF-01 / AC-PERF-02 (memory) / AC-LIFE-02 (real retained memory):** the controlled S10.2 Chrome, Safari and iPad receipts.
4. **Native Excel checkpoint** for AC-STORE-01, AC-SEM-07 and AC-REG-01. The manual steps are above.
5. **AC-BACKEND-01:** the baseline status is pending the S10.2 receipts and the S11 decision.
6. **Full-scale browser runs** for F2 (1M × 8), F3 dense (256 × 16,384) and F4 (1M strings). Only reduced, labelled scales ran here.

## Observations (not fixed; not reproduced as failures)

- G2's read-ahead costs one extra `next()` per complete cursor-less page. On the worker path that is one more worker round trip; it is part of the roughly 597 reads observed above. Detection keeps at most one read-ahead continuation open, because the worker caps open scans at 4 (`MAX_OPEN_SCANS`), shared with find. This is a cost note for the S8.6 owner, not a bug.
- Generated browser packages are cached under `$TMPDIR/officeview-viewer-e2e-v1`, keyed by file name. If the S0.2 builders change, bump the `-v1` suffix in `tests/browser/viewer-e2e.ts`.
- `fromStoreScan` (`src/xlsx/view/structure.ts`) has the same first-page-only pattern as G2. It is used only by tests today. Any future production caller must follow iteration-continued pages.
- On worker-backed F2, the saved-values status reads "0 formulas not recalculated (file too large)". The count follows the documented S1.3 definition (formulas the engine would have evaluated at load), and a cached `SUM` is not one of them. The wording is correct but may confuse users.
- `queryViewport` documents itself as not yet wired to a worker `viewport` command (a TODO in `viewport.ts`). The worker-backed viewer paints from bounded `read rows` windows instead, and those windows are what the complexity audit measured.
