# Interactive viewer: functional and complexity evidence (S10.1)

**Branch:** `feature/iv-s10-1`, cut from `feature/interactive-viewer` at `cfa5cb5` (everything through S5.4, S9.2 and S8.4 merged). **Run date:** 2026-10-12. **Machine:** Apple M4 Pro, macOS 26.5.2. **Engines:** Playwright 1.64.0 with Chromium headless shell build 1248 and Playwright WebKit build 2370. WebKit results are labeled `webkit-playwright` and are **not** Safari evidence (spec §12.2).

This page is the S10.1 functional/complexity matrix, extended by S10.3 (branch `feature/iv-s10-3`, cut at `65a112b`, run 2026-10-12, same machine and engines) with keyboard, accessibility-tree, touch-emulation and capability-failure receipts: see [S10.3 support and accessibility receipts](#s103-support-and-accessibility-receipts). Performance and memory receipts belong to S10.2. Acceptance legs that were not executed are listed as **not-run**, with the reason, and stay blocking. A passing vitest suite counts as code-test evidence. Watching the UI is not a receipt.

## How to reproduce

```sh
bun install --frozen-lockfile
bunx tsc --noEmit
bunx vitest run                                                    # whole suite
bunx vitest run tests/xlsx-viewer-complexity.test.ts               # complexity audit
bunx playwright test --config playwright.viewer.config.ts          # real browsers, both source paths (+ S10.3 a11y/capability specs)
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
| **S10.3 head** (`feature/iv-s10-3`) | `tsc` pass; `vitest run` **189 files, 5,388 passed, 2 skipped, 1 todo** (adds `xlsx-viewer-capability.test.tsx`, `xlsx-viewer-support.test.tsx`); Playwright **53 passed, 3 skipped** (skips: Chromium-only long-task, CDP AX tree and touch emulation on WebKit, by design); build eager `dist/index.js` **38.34 kB gzip** (ceiling 38.72); golden compare: every XLSX unit passes, the same 6 pre-existing DOCX/PPTX failures |
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
| `f2-storage.xlsx` (S10.3) | `writeFixturePackage('F2', { rowCount: 160000 })` | reduced F2; page bytes exceed the 32 MiB memory page store | worker-backed (storage/quota scenarios) |
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
| AC-CAP-03 | F2/F3/F4 quota variants | `xlsx-viewer-import-limits.test.ts`, `xlsx-viewer-source-pages.test.ts` (typed quota/incomplete/cancel). S10.3 browser: no OPFS/IndexedDB and zero quota are visible typed refusals with no grid; quota denied mid-import states the failure with an incomplete body (G7) **[B:C+W]**; `xlsx-viewer-capability.test.tsx` | pass (S10.3; was "pass (unit); browser leg not-run"). Quota is emulated by stubbing the worker's storage APIs, not by filling a real disk quota |
| AC-STORE-01 | F8 | `xlsx-viewer-adapter-parity.test.ts`, `xlsx-viewer-values.test.ts`, `xlsx-viewer-source.test.ts`: object and paged adapters agree on values/errors/cache/formula/date/raw UTF-16; pending ≠ blank. Browser wkr: unknown rows paint "Loading…" (`xlsx-viewer-paged-viewer.test.tsx`) | pass; **native Excel checkpoint not-run** (see below) |
| AC-STORE-02 | F1–F4 malformed variants | `xlsx-viewer-import.test.ts`, `xlsx-viewer-import-limits.test.ts` | pass (unit); peak/cleanup ledger in a browser → S10.2 |
| AC-STORE-03 | F5 | `xlsx-viewer-structure.test.ts` (pure index). Browser F5: giant narrow-column merge and fixed array queried before consumers, child selection allowed **[B:C+W]** | pass (S3.3): `xlsx-viewer-structure-metadata.test.ts` — drawing anchors (oneCell/twoCell/absolute), saved and engine spills, tables, giant merge; resident = paged structural results; fail-closed cases. Browser drawing/spill leg not-run |
| AC-STORE-04 | F8, F7 | `xlsx-viewer-values.test.ts`, `xlsx-viewer-worker-races.test.ts`, `xlsx-viewer-saved-notices.test.tsx`; G1 regression | pass |
| AC-VIEW-01 | F2, F7 | `xlsx-viewer-projection.test.ts`, `xlsx-viewer-sort.test.ts` | pass |
| AC-VIEW-02 | F1, F2 | `xlsx-viewer-cap-regression.test.ts`, `xlsx-viewer-viewport-integration.test.tsx`. Browser F2 wkr: last body row sorted to display row 2 shows "original G100009" with source id preserved **[B:C+W]** | pass |
| AC-VIEW-03 | F5, F6, F7 | `xlsx-viewer-axes.test.ts`, `xlsx-viewer-scroll-window.test.ts`; `xlsx-viewer-complexity.test.ts` (no row lane, bounded canvases). Browser F1: every canvas ≤ viewport × DPR, bounded scroller extent **[B:C+W]** | pass |
| AC-VIEW-04 | F1, F8 | `xlsx-viewer-sheet-zoom.test.ts`, existing zoom/legacy suites. Browser F7: zoom retained across a committed sort, reset by reload (document replacement) **[B:C+W]** | pass |
| AC-SEM-01 | F2, F5 | `xlsx-viewer-filter.test.ts`, `xlsx-viewer-projection.test.ts`. Browser F2 res+wkr: header protected, value filter, Cancel publishes nothing, Clear all restores **[B:C+W]** | pass |
| AC-SEM-02 | F2, F4, F7 | `xlsx-viewer-sort.test.ts` (truth table, 1–8 keys, host collator) | pass (unit); 8-key browser leg not-run |
| AC-SEM-03 | F5, F8 | `xlsx-viewer-sort-range.test.ts`, `xlsx-viewer-range-continuation.test.ts` (G2) | pass |
| AC-SEM-04 | F5 | `xlsx-viewer-operation-guard.test.ts`. Browser F5: merge D1:D1048576 refuses both a region sort and an explicit A2:B7 sort ("Operation refused: merge"); view unchanged **[B:C+W]** | pass for merge/fixed array (browser); dynamic-spill and drawing refusals pass through `ViewerKernel` in both adapters (S3.3 `xlsx-viewer-structure-metadata.test.ts`); browser drawing leg not-run |
| AC-SEM-05 | F5 | `xlsx-viewer-operation-guard.test.ts` (pure guard) | pass (S3.3): filter hiding a row above a drawing anchor refused, filters below anchors commit, absolute anchors never block — through `ViewerKernel`, object and paged. Browser leg not-run |
| AC-SEM-06 | F2, F4, F8 | `xlsx-viewer-filter.test.ts`, `xlsx-viewer-canonical-text.test.ts`, `xlsx-viewer-filter-ui.test.tsx` | pass |
| AC-SEM-07 | F8 | `xlsx-viewer-date-filter.test.ts` | pass (unit); native 1900/1904 workbook checkpoint not-run |
| AC-SEM-08 | F4 | `xlsx-viewer-value-index.test.ts`, `xlsx-viewer-value-search-keys.test.ts`. Browser F4 (12,000 values): "Showing first 10,000", search finds item-11000 beyond 10k, searched Select All selects 1, first Select All covers all 12,000 **[B:C+W]** | pass (100k/1M universe at scale → S10.2) |
| AC-SEM-09 | F4, F7 | `xlsx-viewer-filter-ui.test.tsx`. Browser F4/F2: reopen retains applied choice, Cancel publishes no generation (worker `filter` count unchanged) **[B:C+W]** | pass |
| AC-SEL-01 | F7 | `xlsx-viewer-selection.test.ts` ([0,3,1,2] Shift cases), `xlsx-viewer-copy.test.ts` | pass (unit) |
| AC-SEL-02 | F7 | `xlsx-viewer-selection.test.ts`, `xlsx-viewer-a11y.test.tsx`. Browser F7 hidden-row inspection keeps the active descendant **[B:C+W]** | pass |
| AC-COPY-01 | F7, F8 | `xlsx-viewer-copy.test.ts`, `xlsx-viewer-policy.test.tsx`. Browser F1 res+wkr: clipboard receives exactly "A1048576" after find **[B:C+W]** | pass |
| AC-A11Y-01 | F1, F2, F3, F7 | `xlsx-viewer-a11y.test.tsx`, `xlsx-viewer-a11y-subdetail.test.tsx` (G5), `xlsx-viewer-support.test.tsx` (G6). Browser: `aria-rowcount` 1048577 / `aria-colcount` 16385, bounded owned tree (< 2,000 cells), header/cell index agreement **[B:C+W]**. S10.3: one Tab stop (G6 fixed) and edge navigation after recenter **[B:C+W]**; Chromium AX tree over CDP — grid → row → rowheader/columnheader/gridcell hierarchy, valid active-descendant relation to a named owned gridcell, `selected` on 4 cells, polite/atomic live status; DOM `aria-rowindex` = displaySlot+2 after a sort with a hidden row, `aria-colindex` = sourceCol+2, `aria-sort` on the priority-1 header **[B:C]** | pass with not-run legs: **macOS VoiceOver receipt not-run** (cannot be driven unattended; manual steps below); **iPad VoiceOver not-run** (user decision). Blocking only on the screen-reader receipt |
| AC-A11Y-02 | F2, F7 | `xlsx-viewer-events.test.tsx`, `xlsx-viewer-filter-ui.test.tsx`, `xlsx-viewer-sort-ui.test.tsx`; G4 browser regression; `xlsx-viewer-capability.test.tsx` (G10 Esc). S10.3 browser: arrows/Enter/Home/Ctrl edges, Shift and Ctrl/Cmd+click multi-selection, Esc order (popup → selection), Alt+Down filter with focus return to the grid, F2 control mode with Enter/Space on native header controls and focus return to the popup invoker, Ctrl/Cmd+F find, no zoom/navigation hijack of the search and Go To inputs **[B:C+W]**; coarse-pointer emulation at default column widths: every fully visible control ≥ 44 px, tap selects, row-header tap owned by its row **[B:C, emulation]**. That browser run never exercised columns narrower than 44 px: the S13 review found legal 20 px columns gave 20 px touch sort/filter buttons. S13.2: spec §9.3's 44 px minimum applies to every header control, select included, so for the active cell's column narrower than 44 px a 44 × 44 corner-lane disclosure ("Actions for column B", in the lane under "Select worksheet" that no column owns) opens that column's select, sort, filter and more-sort buttons, each 44 × 44, with no inline sub-44 px action left in narrow headers. `xlsx-viewer-sort-ui.test.tsx` (S13.2 item 5) checks the 20 px case: target sizes, no overlap with any column header, aria-expanded/controls, handlers, Escape with focus return in control mode, tap-outside close, and that the lane's select selects that column in `GridSemantics` (unit, jsdom) | pass (unit) with not-run legs: **narrow-column coarse-pointer browser re-run not-run** (fix after S10.3); **Safari desktop not-run** (needs admin `safaridriver --enable`; WebKit runs are Playwright emulation); **iPad touch not-run** (user decision). Remaining limit: the narrow column's own strip button (column-wide, selection by hit-test) is how a column first becomes active from its header; tapping any of its cells also does. The open lane group overlays the row-header strip below the header until it closes. The real-device legs stay listed under *Blocking* |
| AC-LIFE-01 | F7 | `xlsx-viewer-find-races.test.tsx`, `xlsx-viewer-worker-races.test.ts`, `xlsx-viewer-paged-viewer.test.tsx`. Browser F7: sheets A/B both at their first revision keep their own sort and data across switches; views never bump fontGen; sort/filter keep zoom **[B:C+W]** | pass |
| AC-LIFE-02 | F7 multi-sheet | `xlsx-viewer-lifecycle.test.tsx` (F7 plateau, close returns ledgers to baseline) | pass (unit); real retained-memory/storage receipts → S10.2. Blocking |
| AC-SEARCH-01 | F1, F7, F8 | `xlsx-viewer-find.test.ts`, `xlsx-viewer-find-races.test.tsx`. Browser F1 res+wkr: find reaches the never-painted A1048576 with source id 17179852800 **[B:C+W]** | pass |
| AC-RENDER-01 | F1, F3, F5, F6 | `xlsx-viewer-layout.test.ts`, `xlsx-viewer-overview.test.ts`, `xlsx-viewer-complexity.test.ts` (F6 window visits ≤ 2 × visible cells; 25 % tiny-size frame is analytic, 0 per-cell visits); G5. S13.2 (the S13 review found paged sheets painted no drawings): `xlsx-viewer-paged-drawings.test.ts` (same workbook resident vs worker-backed: drawable count, placement, extent, geometry, fill, line, text, theme, VML, picture bytes; budget diagnostics), `xlsx-viewer-paged-drawings-viewer.test.tsx` (picture lease and status notice through `OfficeFile`), `xlsx-viewer-remote-drawings.test.ts` (owner checks, retries, byte budget). The worker read is a transient share of the document pool (refused with a notice when it does not fit), parsed payloads are pooled and reused, and the drawing node budget is per document. Paged print is refused with a statement that cells and drawings are not printed | pass (unit); no browser pixel receipt for paged drawings |
| AC-RENDER-02 | F7, F8 | `xlsx-viewer-selection-overlay.test.ts`, `xlsx-viewer-viewport-integration.test.tsx`, static print/export suites | pass (unit) |
| AC-PERF-01 | F1–F8 | S10.2: controlled Chrome/Safari/iPad p50/p95 receipts | **not-run** (out of S10.1 scope; Safari/iPad also blocked on hardware/automation). Blocking |
| AC-PERF-02 | F2–F7 | complexity part: `xlsx-viewer-complexity.test.ts` (see below); per-pool memory receipts → S10.2 | complexity **pass**; memory **not-run**. Blocking |
| AC-BACKEND-01 | — | TS resident baseline is the only backend. S11/S12 are decided not-run on desktop Chrome (headless)-only evidence: `docs/architecture/xlsx-wasm-decision.md` (nine S8.6 scenarios against the five frozen kernel thresholds pass; the S10.2 ack miss is detection I/O, not kernel CPU) and `docs/architecture/xlsx-gpu-decision.md` (every paint row and the 128 MiB bitmap budget pass; the one paint bottleneck was fixed in place). Revisit when Safari and mobile receipts exist. S10.3 fallback leg for the Canvas2D reference: `contextrestored` drops tiles and repaints both layers (G9, emulated events) **[B:C+W]** | **not-run** (blocking: baseline S10.2 receipt pending merge; S11/S12 decisions recorded as not-run); device-loss leg pass (emulated). Blocking |
| AC-REG-01 | F8, corpus | typecheck, full vitest, build, golden compare (all XLSX units pass; 6 non-XLSX failures pre-existing on base) | pass (code); native source/checkpoint evidence not-run |

**Counts (34 acceptance cases):**

| Status | Cases | Which |
|---|---:|---|
| pass | 18 | AC-CAP-01, CAP-03, STORE-04, VIEW-01–04, SEM-01, SEM-03, SEM-06, SEM-09, SEL-01, SEL-02, COPY-01, LIFE-01, SEARCH-01, RENDER-01, RENDER-02 |
| pass with a not-run leg | 12 | AC-CAP-02, STORE-01, STORE-02, STORE-03, SEM-02, SEM-04, SEM-05, SEM-07, SEM-08, REG-01, A11Y-02 (Safari/iPad legs), A11Y-01 (screen-reader legs, still blocking) |
| partial, blocking | 2 | AC-LIFE-02, PERF-02 |
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
- **Safari desktop: not-run** (also in S10.3). `safaridriver` exists, but remote automation requires a one-time `safaridriver --enable` (admin prompt) plus Develop ▸ Allow Remote Automation, which needs the user. The Playwright WebKit runs above are not Safari evidence.
- **iPad / physical devices: not-run** by user decision.

## S10.3 support and accessibility receipts

**Branch:** `feature/iv-s10-3` (cut at `65a112b`). **Specs:** `tests/browser/interactive-viewer-a11y.spec.ts`, `tests/browser/interactive-viewer-capability.spec.ts`; vitest `tests/xlsx-viewer-capability.test.tsx`, `tests/xlsx-viewer-support.test.tsx`. Same page (`src/debug/viewer-e2e.html`), engines and labels as S10.1: **[C]** Chromium headless shell, **[W]** Playwright WebKit (`webkit-playwright`, never Safari evidence).

**Capability denial.** Main-thread capabilities are removed with Playwright init scripts. Init scripts never reach a dedicated worker, so the page passes worker-side denials in the Worker `name` (`?wcaps=` → `caps:<flags>`), and `src/debug/workerCaps.ts` applies them before the worker entry evaluates: `no-offscreen`, `no-fonts`, `no-storage` (OPFS `getDirectory` rejects, no IndexedDB), `quota-zero` (`estimate()` reports 0 free), `idb-quota` (no OPFS; every IndexedDB `put` throws `QuotaExceededError`). Storage scenarios use `f2-storage.xlsx` (F2 at 160,000 rows, 6.5 MB compressed; its expected page bytes (6 × compressed, 38,944,548 B) exceed the 32 MiB in-memory page store, so it needs OPFS or IndexedDB).

### Reproduced gaps and fixes (S10.3)

Each gap was reproduced in a real browser (G10 by reading the handler), pinned by a test that failed for the intended reason, then fixed. For G6 the browser spec itself was RED first. For G7–G9 the browser reproduction was an exploratory probe on the e2e page ("Worker is not defined"; "Visible body 0 / complete body 0" with cells "Loading…"; a pane left blank), the vitest regressions were RED first, and the browser scenarios were committed with the fixes.

| # | Gap (where reproduced) | Requirement | Fix | Regression tests |
|---|---|---|---|---|
| G6 | **Chromium keyboard trap.** Chromium makes an `overflow:auto` element without a tabindex keyboard-focusable. The pan scroller sits inside the grid, whose focus capture sends inner focus back to the grid wrapper, so Tab cycled grid → scroller → grid forever: focus could not leave the viewer by keyboard. Observed focus sequence: `grid, scroller, grid, scroller, …`. WebKit was unaffected. | A11Y-01 (one Tab entry), A11Y-02 | `InteractiveSheetDoc.tsx`: the scroller has `tabIndex={-1}`. | browser "one Tab stop…" **[C+W]** (RED on Chromium: focus stayed in the grid after Tab); `xlsx-viewer-support.test.tsx` (every scroll container inside the grid is out of the tab order; RED: `expected null to be '-1'`) |
| G7 | **Quota denied mid-import masqueraded as an empty sheet.** With `idb-quota`, every worker read answered `incomplete` with `{code:'quota', message:'browser storage quota exceeded while writing a source page'}`. The status read "Visible body 0 / complete body 0", cells stayed "Loading…" for good, and no notice said why. | AC-CAP-03 (no empty-cell lie), LIFE-03 | `viewerRemoteSource.ts`: `RemoteWorkbook` records the failure reported on structural reads and exposes `loadState()` (complete / loading / incomplete + message). `InteractiveSheetDoc.tsx`: the status uses it ("body N (loading)" until the import completes; "(incomplete) — This sheet could not be fully loaded: … Free some browser storage for this site (or use a smaller file), then reopen the file." on a storage failure). Review round 1: a storage failure is terminal — `pagedSource.ts` reports reads after an ended import as not retryable (the import's own flag still means "reopening may succeed"), and `RemoteWorkbook` stops structural polling (it observed 11 reads in 5 s before); every load-state change notifies subscribers, including loading → complete with unchanged geometry. | browser "quota denied while the worker writes pages…" **[C+W]**; `xlsx-viewer-capability.test.tsx` (the observed reply shape; RED: status "complete body 2" with no failure text). Round 1: `xlsx-viewer-capability.test.tsx` RemoteWorkbook cases (RED: 11 structural reads, expected 1; 0 notifications on completion), `xlsx-viewer-source-pages.test.ts` "reads after an import ended by a quota failure are not retryable" (RED: `retryable: true`), browser structural-read count flat for 3 s **[C+W]** |
| G8 | **Worker unavailable failed the document.** With no `Worker` global, the host `viewerWorker` factory threw and a large file failed with "Worker is not defined". A worker that dies, or no worker at all, keeps compatibility mode. | §4.2 transition policy, LIFE-03 | `OfficeFile.tsx`: a factory that throws returns "no paged route", so the file keeps compatibility mode with its notice. | browser "Worker unavailable…" **[C+W]**; `xlsx-viewer-capability.test.tsx` (RED: received "Worker is not defined") |
| G9 | **Canvas context loss left a blank pane.** Nothing listened for `contextrestored`: a restored 2D context (and possibly its cached tile canvases) stayed blank until the next scroll. | AC-BACKEND-01 (device-loss fallback), RENDER-01 | `InteractiveSheetDoc.tsx`: on `contextrestored` of either canvas, drop all cached tiles and repaint both layers. | browser "canvas context loss (emulated)…" **[C+W]** (pixel alpha 0 → 255, `basePaints` grows); `xlsx-viewer-support.test.tsx` (RED: `expected 6 to be greater than 6` basePaints) |
| G10 | **Esc never cancelled a pending copy** outside the grid: the root handler compared the lower-cased key with `'Escape'`. | A11Y-02 (§9.3 Esc fall-through) | `InteractiveSheetDoc.tsx`: compare with `'escape'`. | `xlsx-viewer-capability.test.tsx` (copy waiting on held worker reads; Esc on "Cancel copy"; RED: the button stayed) |

Not a gap (designed): committing a sort from a header control ends control mode and returns focus to the grid with a valid active descendant, because the committed view is a new owner (GridSemantics resets control mode on owner change). Focus is never lost to `<body>`; re-enter with F2.

### Keyboard and focus receipts **[C+W]**

| Scenario | Result |
|---|---|
| One Tab stop: Print → grid → out of the viewer; Shift+Tab returns | pass (after G6) |
| ArrowDown, ArrowRight, Enter (moves down), Home, Ctrl/Cmd+Home; live status names each new active cell; descendant valid each step | pass |
| Edge navigation across recentered segments: Ctrl/Cmd+ArrowDown → A1048576, ArrowUp → A1048575, Ctrl/Cmd+ArrowRight → XFD1048575, ArrowLeft → XFC1048575 with matching row/column headers | pass |
| Shift+Arrow range B2:C4 (6 `aria-selected` cells) plus Ctrl/Cmd+click E6 (7) with grid focus kept | pass |
| Alt+Down opens the active column's filter with focus inside; Esc closes only the popup (selection kept, focus to grid, valid descendant); second Esc clears selection and the descendant ("Selection cleared.") | pass |
| F2 → "Select Column D" (Enter selects the column: only its header has `aria-selected=true`); Tab → "Cycle sort column D"; Enter → ascending; F2/Tab/Space → descending (`aria-label` "Column D; sort priority 1, descending"); Tab → filter button, Enter opens, Esc returns focus to that invoker, Esc returns to the grid. Popup names: the filter popup is "Column filter", the sort popup "Sort worksheet" (round 1; was "Column filter" for both; `xlsx-viewer-support.test.tsx`) | pass |
| Ctrl/Cmd+F focuses find; typing `-0+=` and ArrowDown/Home in find, and `0-`/ArrowUp in Go To, edit the input only (zoom stays 100%, active cell unchanged); on the grid `+` zooms and `0` resets | pass |
| Coarse pointer (Chromium `isMobile`/`hasTouch` emulation, 1024×768 DPR 2): every fully visible grid control ≥ 44 × 44 px (pane-clipped edge columns are clamped by design), tap selects C5, a tap 2 px inside row 6's header selects only row 6. Default column widths only: narrow (< 44 px) columns were not exercised; see AC-A11Y-02 for the S13.2 fix and its unit evidence | pass **[C, emulation]** |

### Accessibility tree receipt **[C]**

Chromium's AX layer read with `Accessibility.getFullAXTree` (what a screen reader is given), F7 sheet A after a score sort (hidden source row 8, so display slots differ from source rows) and a Shift selection:

```
1 grid "Worksheet A" (focused, multiselectable), rows=19, columnheaders=23, rowheaders=18,
gridcells=396, selected=4, activedescendant="B3, value 1; original B40", live=polite/atomic;
DOM rowcount/colcount=1048577/16385, sorted=[["Column B; sort priority 1, ascending","ascending"]]
```

Asserted: exactly one grid; every row's semantic parent is the grid; every gridcell/rowheader/columnheader's parent is a row; one header row holding the corner and column headers; each body row owns exactly one rowheader; the `activedescendant` relation resolves to an owned gridcell whose AX name equals the DOM label and carries "original …"; 4 cells expose `selected`; the status is a polite, atomic live region. CDP exposes no AX property for ARIA indices, so they were read from the owned DOM: header row `aria-rowindex=1`, corner `aria-colindex=1`, columns A–C at 2–4, rowheaders at column 1, first cells at column 2. Round 1: body rows are checked against fixture-derived values, not self-consistency — after the sort, slots 1–6 hold ids 40..35 and slots ≥ 8 hold 34, 33, …, slot 7 is the pinned hidden source row 8, so `aria-rowindex` 8 and 10 are present and 9 is absent, each row is labelled "Row slot+1", and column A reads "A<slot+1>, value <id>; original A<id+1>" (VisibleRank indexing would close the gap and fail).

### Capability matrix

| Capability removed | Small admitted file | Large (paged) file | Result |
|---|---|---|---|
| `Worker` (main thread) | `f1-small`: resident, Go To XFD1048576 exact, no worker created **[C+W]** | `f1-paged`: compatibility mode with "Large file — opened in compatibility mode", corner reachable, no alert (after G8) **[C+W]** | pass |
| `OffscreenCanvas` (main + worker) | — | `f1-paged`: worker-backed (1 worker, 1 open), corner exact, no compatibility notice **[C+W]** | pass |
| Fonts: no `FontFace`/`document.fonts` (main) and no `FontFace`/`self.fonts` (worker) | — | `f1-paged`: corner exact, owned grid, base painted **[C+W]**; vitest: a face whose registration throws yields a `font-load-failed` diagnostic, the viewer paints and keeps its grid | pass. Round 1: the status shows "Font not available, shown with a fallback: Brand Sans" (deduplicated by family, at most 3 named + "and N more"; `xlsx-viewer-support.test.tsx`, RED: status had counts only) |
| OPFS + IndexedDB, and zero quota | `f1-paged` (page bytes within the 32 MiB memory store): worker-backed from bounded memory pages **[C+W]** | `f2-storage`: visible typed refusal "…needs browser storage (OPFS or IndexedDB)…" / "browser storage has 0 bytes free; about 38944548 are needed…", no grid, worker terminated **[C+W]** | pass |
| IndexedDB quota during import (OPFS denied) | — | `f2-storage`: "(incomplete) — This sheet could not be fully loaded: browser storage quota exceeded while writing a source page"; never "complete body" (after G7) **[C+W]** | pass |
| *(none denied)* Playwright WebKit, large file | — | `f2-storage` worker-backed on default Playwright WebKit (build 2370): OPFS is chosen, its page writes fail, and the sheet shows "Visible body 0 / body 0 (incomplete) — This sheet could not be fully loaded: browser storage failed while writing a source page. Free some browser storage for this site (or use a smaller file), then reopen the file." No IndexedDB fallback is attempted (`sourcePages.ts` `choosePageBackend`). Files whose pages fit the 32 MiB memory store (every S10.1 fixture) are unaffected; real Safari not-run | **documented capacity restriction** (Playwright WebKit, worker-backed files > ~5.6 MB compressed): visibly diagnosed, no capacity claim |
| No viewer worker supplied | — | `f2-storage`: compatibility mode with its notice **[C+W]** | pass |
| Canvas 2D context loss | `f2-1k`: visible canvases cleared, `contextlost`/`contextrestored` dispatched; the pane is repainted after restore (pixel alpha 0 → 255, `basePaints` grows) and the viewer keeps working (after G9) **[C+W]**. Tile-cache invalidation itself is pinned by the vitest (detached tile canvases are not reachable from the page) | — | pass (emulated: a 2D context loss cannot be triggered from page script; real GPU-process loss not-run) |

No unsupported profile receives a capacity claim: refused files show no grid; compatibility mode is labelled; an incomplete import says so.

### Manual macOS VoiceOver steps (not-run)

VoiceOver cannot be driven unattended (it needs System Settings/TCC grants and a user session), so no screen-reader receipt is claimed. To produce it:

1. `bun install --frozen-lockfile`, then `bunx vite --config vite.harness.config.ts --port 4173 --strictPort --host 127.0.0.1`.
2. Serve a fixture to the page: copy `tests/fixtures/xlsx-viewer/f1-small.xlsx` (or a generated one from `$TMPDIR/officeview-viewer-e2e-v1/`) into `src/debug/` and open `http://127.0.0.1:4173/src/debug/viewer-e2e.html?src=/src/debug/f1-small.xlsx` in Safari and in Chrome.
3. Start VoiceOver (Cmd+F5). Tab to the grid. Expect: "Worksheet F1, grid"; the capacity description; the active cell announced as "A1, value …".
4. Arrow Right/Down: each move announces the new cell name. Shift+Arrow: the extension target is announced and VO reports "selected".
5. Go To `XFD1048576`, focus the grid: VO reads "XFD1048576, value 17179869183"; VO+Right/Left within the row reads neighbouring owned cells and their column headers ("Column XFD"); row header "Row 1048576".
6. On `f7-two-sheets.xlsx` (generated), sort column B: the column header reads "Column B; sort priority 1, ascending" and VO reports sort state; cell names append "original A41"-style source addresses.
7. F2: VO lands on "Select Column …, button"; Tab through sort/filter buttons; Enter opens "Column filter, dialog"; Esc returns to the invoker; Esc returns to the grid and the active cell is re-announced.
8. Esc with a selection: VO announces "Selection cleared." from the status live region.
9. Record VO output (VoiceOver Utility ▸ Verbosity, or a screen recording with captions), macOS/Safari/Chrome builds, and the commit, and attach them here as the AC-A11Y-01 screen-reader receipt.

### Not-run (S10.3)

- **macOS VoiceOver** (both browsers): not automatable unattended; steps above. AC-A11Y-01 stays blocking on it.
- **Safari desktop**: needs admin `safaridriver --enable`; not attempted. WebKit results are Playwright emulation.
- **iPad, physical devices, iPad VoiceOver, real touch**: not-run by user decision. Touch rows above are Chromium emulation only.
- **Real GPU/device loss**: not triggerable for a 2D canvas from page script; emulated with events.
- **Real disk-quota exhaustion**: quota was emulated by stubbing the worker's storage APIs.

### Observations (S10.3, not fixed)

- **Playwright WebKit cannot write OPFS pages from the worker** (recorded as a capacity restriction in the capability matrix). `choosePageBackend` (`sourcePages.ts`) picks OPFS when `getDirectory` works and does not fall back to IndexedDB when a write fails; an injected OPFS write failure confirms IndexedDB is never probed. Automatic write-failure failover is not required by §4.2 and is left to the storage owner.
- A recorded load failure is cleared by a later complete structural read, but after a terminal storage failure no further read is issued (the import has ended), so that path is only reachable for non-storage failures.

## Blocking for release

These are mandatory cases whose evidence is incomplete. Nothing here is hidden as a pass.

1. ~~**AC-STORE-03 / AC-SEM-04 drawing legs / AC-SEM-05**~~ — resolved by S3.3 (merged 2026-10-12); only their browser legs remain not-run.
2. **AC-A11Y-01 / AC-A11Y-02:** S10.3 delivered the Chromium/WebKit keyboard and focus receipts and the Chromium AX-tree receipt. Still open: the **macOS VoiceOver** receipt (manual steps in the S10.3 section), **Safari desktop** (needs admin `safaridriver --enable`), and **iPad touch/VoiceOver** (not-run by user decision).
3. **AC-PERF-01 / AC-PERF-02 (memory) / AC-LIFE-02 (real retained memory):** the controlled S10.2 Chrome, Safari and iPad receipts.
4. **Native Excel checkpoint** for AC-STORE-01, AC-SEM-07 and AC-REG-01. The manual steps are above.
5. **AC-BACKEND-01:** S11/S12 decision records delivered (`docs/architecture/xlsx-wasm-decision.md`, `xlsx-gpu-decision.md`, both not-run on desktop Chrome (headless)-only evidence); the shipped TS/Canvas2D backends' S10.2 receipt is still pending merge; revisit when Safari and mobile receipts exist.
6. **Full-scale browser runs** for F2 (1M × 8), F3 dense (256 × 16,384) and F4 (1M strings). Only reduced, labelled scales ran here.

## Observations (not fixed; not reproduced as failures)

- G2's read-ahead costs one extra `next()` per complete cursor-less page. On the worker path that is one more worker round trip; it is part of the roughly 597 reads observed above. Detection keeps at most one read-ahead continuation open, because the worker caps open scans at 4 (`MAX_OPEN_SCANS`), shared with find. This is a cost note for the S8.6 owner, not a bug.
- Generated browser packages are cached under `$TMPDIR/officeview-viewer-e2e-v1`, keyed by file name. If the S0.2 builders change, bump the `-v1` suffix in `tests/browser/viewer-e2e.ts`.
- `fromStoreScan` (`src/xlsx/view/structure.ts`) has the same first-page-only pattern as G2. It is used only by tests today. Any future production caller must follow iteration-continued pages.
- On worker-backed F2, the saved-values status reads "0 formulas not recalculated (file too large)". The count follows the documented S1.3 definition (formulas the engine would have evaluated at load), and a cached `SUM` is not one of them. The wording is correct but may confuse users.
- `queryViewport` documents itself as not yet wired to a worker `viewport` command (a TODO in `viewport.ts`). The worker-backed viewer paints from bounded `read rows` windows instead, and those windows are what the complexity audit measured.
