# Interactive viewer: functional and complexity evidence (S10.1; S10.2 performance rows)

> **Current status (2026-10-16, integration fb0ddc7):** S13.5, evaluator/all-leaf seams, worker-overflow bridge, reporter replay, final cleanup and the release-regression ownership fix are merged. Final code candidate reviews: agy + opencode APPROVE (7e1bb8b), and codex APPROVE of the ownership fix. Final-tree evidence on tree 7effb20 (= benchmarked 9cfb090):
> - desktop Chrome release benchmark ADMITTED, 13/13 rows (see `interactive-viewer-performance.md` § Final-tree release receipt);
> - Chromium/WebKit browser suite 70 passed / 8 skipped;
> - vitest 6581 passed / 0 failed.
>
> The receipt tables below keep their named historical scopes. `report-plan-final-codex.md` records the checkbox audit. README's copy/bitmap profile sentence has been corrected (main-thread pools follow the source's admitted profile; bitmap follows the viewport).

> **User decision, still in force:** Safari and VoiceOver final re-checks are deferred while the Mac is locked; do not attempt them. Earlier Safari 7/7 and guided VoiceOver V1/V2/V5 checks at 0c65bff remain historical receipts. Physical iPad/device checks and Excel for Windows remain not-run. These deferrals do not block package integration merges and are not a final release-readiness claim.

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

## Gate results (historical package heads)

| Gate | Result |
|---|---|
| `bunx tsc --noEmit` | pass (no output) |
| `bunx vitest run` | **186 files passed; 5,362 tests passed, 2 skipped, 1 todo** at the branch head (review round 1 included). An earlier run reported 184 files / 5,279 tests because two corpus-dependent files skipped before `corpus/` was populated. Base `cfa5cb5` without corpus: 180 files / 5,265 tests, all passing |
| `bunx playwright test --config playwright.viewer.config.ts` | **27 passed, 1 skipped** (the skip is the Chromium-only long-task case on WebKit, by design). This includes 20 S10.1 scenarios, 10 each on Chromium and WebKit |
| `bun run build` | pass; eager `dist/index.js` **38.32 kB gzip** (vite report), unchanged from base 38.32 kB, ceiling 38.72 kB |
| **S10.3 head** (`feature/iv-s10-3`) | `tsc` pass; `vitest run` **189 files, 5,388 passed, 2 skipped, 1 todo** (adds `xlsx-viewer-capability.test.tsx`, `xlsx-viewer-support.test.tsx`); Playwright **53 passed, 3 skipped** (skips: Chromium-only long-task, CDP AX tree and touch emulation on WebKit, by design); build eager `dist/index.js` **38.34 kB gzip** (ceiling 38.72); golden compare: every XLSX unit passes, the same 6 pre-existing DOCX/PPTX failures |
| **S13.5 head** (`feature/iv-s13-5`, after merging S10.2 `d40cefe`) | `tsc --noEmit`: no errors outside `tests/browser/`; the 58 `tests/browser/*.spec.ts` errors come from `@playwright/test` not being installed in this worktree. `vitest run`: 212 files, 5,530 passed, 10 skipped, 1 todo, 3 failed. The 3 failures (`drawing-cached-content-regression` ×2, `drawing-geometry` ×1, subprocess module resolution) also fail on base `4e4f0e1`. `bun run build` stops at its `tsc` step for the same Playwright reason; `vite build` plus `tsc -p tsconfig.build.json` pass. Eager `dist/index.js` is **38.37 kB gzip**, identical to the integration head `d40cefe` (ceiling 38.72) |
| **S13.5 r4 fixes** (`4d39757`, codex r3 #1/#2) | Worktree `tsc --noEmit`: only the 58 `tests/browser/` Playwright-not-installed errors. On `git archive 4d39757` with matching installed dependencies: `tsc --noEmit`, `typecheck:scripts` and `bun run build` pass; eager `dist/index.js` **38.37 kB gzip** (ceiling 38.72). `vitest run`: 217 files, 5,543 passed, 10 skipped, 1 todo, 3 failed (`drawing-cached-content-regression` ×2, `drawing-geometry` ×1), a subset of the 4 that fail on exact base `b4c447a` in the same environment (base also timed out `xlsx-viewer-complexity` Go To XFD1048576). Focused 15 suites 505/505 |
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
| AC-CAP-02 | F1, F3, F7 | `xlsx-viewer-scroll-window.test.ts`, `xlsx-viewer-navigation.test.tsx`, `xlsx-viewer-selection.test.ts`. Browser: F1 Go To XFD1048576 with header/row label agreement; F3 recenter, header/cell `aria-colindex` agreement and edge ArrowLeft after recenter; F7 Original-cell Go To on a hidden row opens the read-only inspector without moving the active descendant **[B:C+W]** | pass (zoom/DPR browser leg partial: zoom 125% exercised; **S13 DPR change — emulated (CDP) in installed Chrome, not a physical display move:** header/cell/gridline agreement holds and, after the D-DPR fix, a DPR-only change re-rasterizes by itself — see [S13](#s13-runtime-and-manual-checks)) |
| AC-CAP-03 | F2/F3/F4 quota variants | `xlsx-viewer-import-limits.test.ts`, `xlsx-viewer-source-pages.test.ts` (typed quota/incomplete/cancel). S10.3 browser: no OPFS/IndexedDB and zero quota are visible typed refusals with no grid; quota denied mid-import states the failure with an incomplete body (G7) **[B:C+W]**; `xlsx-viewer-capability.test.tsx` | pass (S10.3; was "pass (unit); browser leg not-run"). Quota is emulated by stubbing the worker's storage APIs, not by filling a real disk quota |
| AC-STORE-01 | F8 | `xlsx-viewer-adapter-parity.test.ts`, `xlsx-viewer-values.test.ts`, `xlsx-viewer-source.test.ts`: object and paged adapters agree on values/errors/cache/formula/date/raw UTF-16; pending ≠ blank. Browser wkr: unknown rows paint "Loading…" (`xlsx-viewer-paged-viewer.test.tsx`) | pass; **native Excel checkpoint pass** for three Excel 16 (macOS) workbooks: adapters agree with each other and with Excel (`xlsx-viewer-native-excel*.test.ts`, see [Native Excel](#native-excel-safari-and-devices)). Worker-backed browser legs pass [C+W]: `native-large.xlsx` (200,001 cells) is admitted paged with the shipped limits, and `native-corner.xlsx` through the debug-only `?paged=1` override |
| AC-STORE-02 | F1–F4 malformed variants; S13.4 oversized-value and construction-memory fixtures | `xlsx-viewer-import.test.ts`, `xlsx-viewer-import-limits.test.ts`, `xlsx-viewer-blob-records.test.ts`, `xlsx-viewer-construction-memory.test.ts`, oversized parity case in `xlsx-viewer-source.test.ts` | pass (unit, fakes for OPFS/IndexedDB). S13.4: values over the 64 KiB inline soft threshold (2 B per UTF-16 unit) are blob records in the page backend referenced from pages, charged to the storage-backed `blobRecords` pool, removed on cancel/quota/close, and a page that fails after writing some deletes them record by record, so storage, ledger and pool are back to the pre-page state; object/paged parity incl. lone surrogates, find and copy read the full value. Builder buffers, pending pages (with their formula metadata) and shared strings, encode scratch (transient), page summaries and every buffer a memory backend retains (persistent) are charged to the document MemoryPool before allocation; a too-small pool ends the import `resource-limit` / `import-memory-budget`, and binding a pool that cannot hold the live charges refuses the open typed. Rejected commits (overlap, repeated id or chunk) are refused before any write. The estimate model is pinned against a Bun heap probe (Bun/JSC, not Chrome). Limits: a >64 KiB shared string stays inline in its 4,096-string SST chunk (§4.3.5 SST paging, not spilled); charges before the viewer binds the document pool sit on a provisional pool of the worker budget and move on bind; byte figures are conservative estimates checked only on Bun; crash orphans are swept at session start only for owners whose Web Lock nobody holds (plus a 24 h age guard; no Web Locks means no cross-owner sweep) and at store open only within an owner the store holds; real-browser Web Locks/OPFS behaviour is not yet exercised (S10.2). Peak/cleanup in a real browser → S10.2 |
| AC-STORE-03 | F5 | `xlsx-viewer-structure.test.ts` (pure index). Browser F5: giant narrow-column merge, fixed array, dynamic spill, and drawing anchors queried before consumers, child selection allowed **[B:C+W]** | pass (S3.3): `xlsx-viewer-structure-metadata.test.ts` — drawing anchors (oneCell/twoCell/absolute), saved and engine spills, tables, giant merge; resident = paged structural results; fail-closed cases. Browser drawing/spill leg pass **[B:C+W]** |
| AC-STORE-04 | F8, F7 | `xlsx-viewer-values.test.ts`, `xlsx-viewer-worker-races.test.ts`, `xlsx-viewer-saved-notices.test.tsx`; G1 regression | pass |
| AC-VIEW-01 | F2, F7 | `xlsx-viewer-projection.test.ts`, `xlsx-viewer-sort.test.ts` | pass |
| AC-VIEW-02 | F1, F2 | `xlsx-viewer-cap-regression.test.ts`, `xlsx-viewer-viewport-integration.test.tsx`. Browser F2 wkr: last body row sorted to display row 2 shows "original G100009" with source id preserved **[B:C+W]** | pass |
| AC-VIEW-03 | F5, F6, F7 | `xlsx-viewer-axes.test.ts`, `xlsx-viewer-scroll-window.test.ts`; `xlsx-viewer-complexity.test.ts` (no row lane, bounded canvases). Browser F1: every canvas ≤ viewport × DPR, bounded scroller extent **[B:C+W]** | pass |
| AC-VIEW-04 | F1, F8 | `xlsx-viewer-sheet-zoom.test.ts`, existing zoom/legacy suites. Browser F7: zoom retained across a committed sort, reset by reload (document replacement) **[B:C+W]** | pass |
| AC-SEM-01 | F2, F5 | `xlsx-viewer-filter.test.ts`, `xlsx-viewer-projection.test.ts`. Browser F2 res+wkr: header protected, value filter, Cancel publishes nothing, Clear all restores **[B:C+W]** | pass |
| AC-SEM-02 | F2, F4, F7 | `xlsx-viewer-sort.test.ts` (truth table, 1–8 keys, host collator). Browser F2: interactive modal adds 8 levels across columns A–H with tied prefix groups and alternating asc/desc; asserts exact multi-key row order and header updates **[B:C+W]** | pass (unit); 8-key browser modal multi-level sort leg pass **[B:C+W]** |
| AC-SEM-03 | F5, F8 | `xlsx-viewer-sort-range.test.ts`, `xlsx-viewer-range-continuation.test.ts` (G2) | pass |
| AC-SEM-04 | F5 | `xlsx-viewer-operation-guard.test.ts`. Browser F5: merge D1:D1048576 refuses both a region sort and an explicit A2:B7 sort ("Operation refused: merge"); dynamic-spill ("Operation refused: dynamic-spill") and drawing anchor ("Operation refused: drawing-impact") refuse sort **[B:C+W]**; view unchanged | pass (S3.3): `xlsx-viewer-structure-metadata.test.ts` — drawing anchors (oneCell/twoCell/absolute), saved and engine spills, tables, giant merge; resident = paged structural results; fail-closed cases; browser merge/fixed array/drawing/dynamic-spill refusal legs pass **[B:C+W]** |
| AC-SEM-05 | F5 | `xlsx-viewer-operation-guard.test.ts` (pure guard). Browser F5: filter hiding row above drawing anchor refused ("Operation refused: drawing-impact"), filter below anchor allowed and commits **[B:C+W]** | pass (S3.3): filter hiding a row above a drawing anchor refused, filters below anchors commit, absolute anchors never block — through `ViewerKernel`, object and paged. Browser leg pass **[B:C+W]** |
| AC-SEM-06 | F2, F4, F8 | `xlsx-viewer-filter.test.ts`, `xlsx-viewer-canonical-text.test.ts`, `xlsx-viewer-filter-ui.test.tsx`; S13.5 `xlsx-viewer-display-lane.test.ts` (on-demand formatted display-text lane, worker and main-thread paths) | pass (unit). Semantics unchanged: parity with the former eager preparation over 24 rule shapes × numeric/date/mixed columns. The formatted lane is now built only for a text condition on a text/mixed column. It is keyed by valueGen and style owner and reused across applies. **Worker:** it is admitted against the session's single prepared-column budget and counter (`preparedBytes`, 64 MiB D / 32 MiB M), shared with sort/filter prepared columns; room is made by one shared recency across both kinds, a lane also holds one `maxColumns` slot, leased lanes and pinned columns are never evicted, raw prepare staging and lane builds are admitted together, and a lane still being read (retired, or across session close) stays charged until its last lease ends. Full bounded/lifetime clause subsequently approved by opencode (net delta + NIT-1) and agy/opencode final candidate reviews; merged by 485a57c. The staging/idle/leased-close cases are pinned by `xlsx-viewer-display-lane-lifecycle.test.ts`. **Main thread:** it is charged to the controls caches' share (16 MiB D / 8 MiB M, reported in the main ledger's `controls` pool) for its retained lifetime. Exhaustion is a typed `resource-limit` refusal. Bun-only timing ([S13.5](#s135-on-demand-display-text-lane)); no browser leg |
| AC-SEM-07 | F8 | `xlsx-viewer-date-filter.test.ts` | pass (unit); native leg **pass** for 1900 and 1904 workbooks saved by Excel (`xlsx-viewer-native-excel-dates.test.ts`): serials and display match Excel on both adapters, and the worker filter evaluator lowers eq/before/after/between (incl. the leap day and 23:59:59)/thisMonth/thisYear to [start, nextStart) over the native cells of both adapters. Midday/zone presets on a native file were checked only with an injected clock |
| AC-SEM-08 | F4 | `xlsx-viewer-value-index.test.ts`, `xlsx-viewer-value-search-keys.test.ts`. Browser F4 (12,000 values): "Showing first 10,000", search finds item-11000 beyond 10k, searched Select All selects 1, first Select All covers all 12,000 **[B:C+W]** | pass (100k/1M universe at scale → S10.2) |
| AC-SEM-09 | F4, F7 | `xlsx-viewer-filter-ui.test.tsx`. Browser F4/F2: reopen retains applied choice, Cancel publishes no generation (worker `filter` count unchanged) **[B:C+W]** | pass |
| AC-SEL-01 | F7 | `xlsx-viewer-selection.test.ts` ([0,3,1,2] Shift cases), `xlsx-viewer-copy.test.ts` | pass (unit) |
| AC-SEL-02 | F7 | `xlsx-viewer-selection.test.ts`, `xlsx-viewer-a11y.test.tsx`. Browser F7 hidden-row inspection keeps the active descendant **[B:C+W]** | pass |
| AC-COPY-01 | F7, F8 | `xlsx-viewer-copy.test.ts`, `xlsx-viewer-policy.test.tsx`. Browser F1 res+wkr: clipboard receives exactly "A1048576" after find **[B:C+W]** | pass |
| AC-A11Y-01 | F1, F2, F3, F7 | `xlsx-viewer-a11y.test.tsx`, `xlsx-viewer-a11y-subdetail.test.tsx` (G5), `xlsx-viewer-support.test.tsx` (G6). Browser: `aria-rowcount` 1048577 / `aria-colcount` 16385, bounded owned tree (< 2,000 cells), header/cell index agreement **[B:C+W]**. S10.3: one Tab stop (G6 fixed) and edge navigation after recenter **[B:C+W]**; Chromium AX tree over CDP — grid → row → rowheader/columnheader/gridcell hierarchy, valid active-descendant relation to a named owned gridcell, `selected` on 4 cells, polite/atomic live status; DOM `aria-rowindex` = displaySlot+2 after a sort with a hidden row, `aria-colindex` = sourceCol+2, `aria-sort` on the priority-1 header **[B:C]** | pass with not-run legs: **macOS VoiceOver receipt not-run** (cannot be driven unattended; manual steps below); **iPad VoiceOver not-run** (user decision). Blocking only on the screen-reader receipt |
| AC-A11Y-02 | F2, F7 | `xlsx-viewer-events.test.tsx`, `xlsx-viewer-filter-ui.test.tsx`, `xlsx-viewer-sort-ui.test.tsx`; G4 browser regression; `xlsx-viewer-capability.test.tsx` (G10 Esc). S10.3 browser: arrows/Enter/Home/Ctrl edges, Shift and Ctrl/Cmd+click multi-selection, Esc order (popup → selection), Alt+Down filter with focus return to the grid, F2 control mode with Enter/Space on native header controls and focus return to the popup invoker, Ctrl/Cmd+F find, no zoom/navigation hijack of the search and Go To inputs **[B:C+W]**; coarse-pointer emulation at default column widths: every fully visible control ≥ 44 px, tap selects, row-header tap owned by its row **[B:C, emulation]**. That browser run never exercised columns narrower than 44 px: the S13 review found legal 20 px columns gave 20 px touch sort/filter buttons. S13.2: spec §9.3's 44 px minimum applies to every header control, select included, so for the active cell's column narrower than 44 px a 44 × 44 corner-lane disclosure ("Actions for column B", in the lane under "Select worksheet" that no column owns) opens that column's select, sort, filter and more-sort buttons, each 44 × 44, with no inline sub-44 px action left in narrow headers. `xlsx-viewer-sort-ui.test.tsx` (S13.2 item 5) checks the 20 px case: target sizes, no overlap with any column header, aria-expanded/controls, handlers, Escape with focus return in control mode, tap-outside close, and that the lane's select selects that column in `GridSemantics` (unit, jsdom) | pass (unit) with not-run legs: **narrow-column coarse-pointer browser re-run not-run** (fix after S10.3); **Safari desktop not-run** (needs admin `safaridriver --enable`; WebKit runs are Playwright emulation); **iPad touch not-run** (user decision). Remaining limit: the narrow column's own strip button (column-wide, selection by hit-test) is how a column first becomes active from its header; tapping any of its cells also does. The open lane group overlays the row-header strip below the header until it closes. The real-device legs stay listed under *Blocking* |
| AC-LIFE-01 | F7 | `xlsx-viewer-find-races.test.tsx`, `xlsx-viewer-worker-races.test.ts`, `xlsx-viewer-paged-viewer.test.tsx`. Browser F7: sheets A/B both at their first revision keep their own sort and data across switches; views never bump fontGen; sort/filter keep zoom **[B:C+W]** | pass |
| AC-LIFE-02 | F7 multi-sheet; F2/F4 browser | `xlsx-viewer-lifecycle.test.tsx` (F7 plateau; close returns the ledgers to baseline). Browser F7: retained memory plateaus across repeated multi-sheet switches, ledger and bitmap back near baseline **[B:C+W]**. S10.2 Chrome receipts r8 (`interactive-viewer-viewer-chrome-r8.json`) and r14 (closed: ledger 0; F4 within the gate, F2 0.06 MiB of Blink/V8 table capacity grown by the instrument's own snapshots): after a close with the host left idle (no re-render, no input; three GCs), every worker terminated (2/2 incl. the warmed host), worker memory 0, main ledger 0 (`close()` drops its buffers; the lifetime-charged pools reach zero through collection), DOM canvases 0, page storage 221.2 MB / 159.1 MB → 0. Main heap owned after close 0.38 / 0.16 MiB; ten F2 open/scroll/sort/close cycles leave 0.62–1.16 MiB (growth 0.44 MiB, under 2 MiB). The React-held update found in r5 is fixed in the viewer (layout-effect observers, liveness guards); `xlsx-viewer-unmount-gc.test.tsx` checks collection under production React and real GC with no further render | pass with not-run legs: F7 multi-sheet switching browser leg pass **[B:C+W]**; Safari/iPad not-run |
| AC-SEARCH-01 | F1, F7, F8 | `xlsx-viewer-find.test.ts`, `xlsx-viewer-find-races.test.tsx`. Browser F1 res+wkr: find reaches the never-painted A1048576 with source id 17179852800 **[B:C+W]** | pass |
| AC-RENDER-01 | F1, F3, F5, F6 | `xlsx-viewer-layout.test.ts`, `xlsx-viewer-overview.test.ts`, `xlsx-viewer-complexity.test.ts` (F6 window visits ≤ 2 × visible cells; 25 % tiny-size frame is analytic, 0 per-cell visits); G5. S13.2 (the S13 review found paged sheets painted no drawings): `xlsx-viewer-paged-drawings.test.ts` (same workbook resident vs worker-backed: drawable count, placement, extent, geometry, fill, line, text, theme, VML, picture bytes; budget diagnostics), `xlsx-viewer-paged-drawings-viewer.test.tsx` (picture lease and status notice through `OfficeFile`), `xlsx-viewer-remote-drawings.test.ts` (owner checks, retries, byte budget). The worker read is a transient share of the document pool (refused with a notice when it does not fit), parsed payloads are pooled and reused, and the drawing node budget is per document. Paged print is refused with a statement that cells and drawings are not printed | pass (unit); no browser pixel receipt for paged drawings |
| AC-RENDER-02 | F7, F8 | `xlsx-viewer-selection-overlay.test.ts`, `xlsx-viewer-viewport-integration.test.tsx`, static print/export suites | pass (unit) |
| AC-PERF-01 | F1–F4 full scale + scroll mixes | S10.2 desktop Chrome receipt r16 (`interactive-viewer-performance.md` § S10.2): **desktop D admitted**, 13/13 rows. Every row is derived from validated raw records; fixtures are verified against the content manifest. The viewer adds no layout shift to its host (`viewer-layout-shift.spec.ts`, CLS ≤ 0.001). F2 filter apply end to end 1.3 s first, 0.66–0.70 s after (was 2.85 s). Kernel rows reused from S8.6 | pass with not-run legs: recorded on a shared machine, so a quiet re-run is pending; actual Safari not-run (admin `safaridriver`); iPad/mobile not-run (user decision); F5–F8 have no browser performance rows |
| AC-PERF-02 | F2–F7 | Complexity: `xlsx-viewer-complexity.test.ts` (see below). Per-pool memory, S10.2 Chrome r16. Main: application ledger peak 41.4 / 64 MiB (budget unchanged; F2), live peak covered; per-checkpoint gate (quiet, sampler-free checkpoints; measured, bounded code/layout-shift exclusions; independent noise ~0.003 MiB) plus an explicit 0.25 MiB engine-cache allowance per checkpoint and for the warm-up residue (**user decision 2026-10-13**; owners: V8 smi_string_cache, Blink IdToNodeMap capacity, lazily loaded module state, untraced remainder disclosed): largest checkpoint excess −0.15 MiB, warm-up objects/arrays/DOM 0.13 MiB. `xlsx-viewer-main-ledger.test.tsx`. Worker: 228.7 / 512 MiB (ledger + measureUserAgentSpecificMemory). Bitmap: 113.2 / 128 MiB S13.5: the formatted display-text lane is built on demand and shares the worker prepared-column budget and the controls main-ledger share ([S13.5](#s135-on-demand-display-text-lane); Bun only). | complexity **pass**; desktop pools **pass** (main under the user-approved engine-cache allowance); mobile M pools not-run (no device); GPU backend absent (Canvas2D only) |
| AC-BACKEND-01 | — | TS resident baseline is the only backend. S11/S12 are decided not-run on desktop Chrome (headless)-only evidence: `docs/architecture/xlsx-wasm-decision.md` (nine S8.6 scenarios against the five frozen kernel thresholds pass; the historical S10.2 r1 ack miss was detection I/O and missing pending UI, not kernel CPU; merged r16 passes acknowledgment) and `docs/architecture/xlsx-gpu-decision.md` (every paint row and the 128 MiB bitmap budget pass; the one paint bottleneck was fixed in place). Revisit when Safari and mobile receipts exist. S10.3 fallback leg for the Canvas2D reference: `contextrestored` drops tiles and repaints both layers (G9, emulated events) **[B:C+W]** | **not-run** (S11/S12 decisions recorded as not-run); device-loss leg pass (emulated). Blocking |
| AC-REG-01 | F8, corpus | typecheck, full vitest, build, golden compare (all XLSX units pass; 6 non-XLSX failures pre-existing on base) | pass (code); native formula/date/Unicode oracle **pass** on three Excel 16 (macOS) workbooks (`xlsx-viewer-native-excel.test.ts` 10/10, `-dates` 5/5, `-large` 5/5; browser native legs 10/10 [C+W]); no wider native corpus, no Excel for Windows |

**Counts (34 acceptance cases):**

| Status | Cases | Which |
|---|---:|---|
| pass | 18 | AC-CAP-01, CAP-03, STORE-04, VIEW-01–04, SEM-01, SEM-03, SEM-06, SEM-09, SEL-01, SEL-02, COPY-01, LIFE-01, SEARCH-01, RENDER-01, RENDER-02 |
| pass with a not-run leg | 15 | AC-CAP-02, STORE-01, STORE-02, STORE-03, SEM-02, SEM-04, SEM-05, SEM-07, SEM-08, REG-01, A11Y-02 (Safari/iPad legs), A11Y-01 (screen-reader legs, still blocking), LIFE-02, PERF-01 (Safari/iPad legs, quiet re-run), PERF-02 (mobile pools; main pool under the 0.25 MiB engine-cache allowance, user decision 2026-10-13) |
| partial, blocking | 0 | |
| not-run, blocking | 1 | BACKEND-01 |
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

- **Native Excel (AC-STORE-01/SEM-07/REG-01 native legs): checkpoint ran on one workbook; see what passed and what is still not-run below.**
  - **Produced:** on 2026-10-12, `osascript` drove the user's own Microsoft Excel 16 for macOS (zh-TW UI) to build and save `native-corner.xlsx`. The user explicitly approved automating their live Excel for this run, which overrides the dedicated-account/VM policy recorded below (kept as history). Excel then reopened the file with **no repair prompt**. The only banner was Excel's AutoRecover offer "開啟已復原的活頁簿？ 您最近所做的變更已儲存. 您想要在上次離開的地方繼續工作嗎?" (as shown in the committed window crop) (open the recovered workbook?), which belongs to an earlier session and is not a repair prompt.
  - **Fixture:** `tests/fixtures/xlsx-viewer/native-corner.xlsx`, 10,441 bytes, SHA-256 `d263264a8648e8e6d4932bc94ee6dc63a7b69bc3978ea245f6a538b1364e6715` (pinned by the test). It is a real Excel package: sheet `工作表1` (zh-TW default name), `sharedStrings`, `theme`, `calcChain`, a table part, `<dimension ref="A1:XFD1048576"/>`.
  - **Privacy edit to the fixture:** one optional element was removed from `xl/workbook.xml`: Excel's `<mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath url="…"/></mc:Choice></mc:AlternateContent>`, which recorded the local save folder. Nothing else was touched: `xl/workbook.xml` was re-deflated with the element gone, and its CRC and sizes were updated in the local and central headers. The other 12 entries keep their local headers and raw compressed bytes, in Excel's order, and the timestamps and compression methods are unchanged. As Excel wrote it, the file was 10,603 bytes with SHA-256 `09104b6d0123eb53c9cd1a7ec900fce780d2c1520abc357f3049dc8631ede82c`. Run against the edited file, the parity test failed only its hash check (8 of 9 passed), so the edit changed nothing the viewer reads. A new test asserts that `absPath` is gone and the sheet entry is intact. `docProps/core.xml` still names the creator ("Julian Shen", the repository's git author).
  - **Screenshot:** `docs/validation/native/native-corner-excel-reopen-window.png` is the Excel window cropped from the full-screen capture (SHA-256 `c027ee46fff8f95bdc8b6c3c98ca61350cfc37822b45f8d244e02f53ac17994b`). The full-screen original showed other applications, so the integrator keeps it privately; it is not in the repository.
  - **Excel's displayed values after reopening** (AppleScript `string value`), which the tests treat as the oracle: A1 `42`; A2 `=DATE(2024,1,15)+0.5`, format `yyyy\-mm\-dd\ hh:mm`, shown `2024-01-15 12:00` (serial 45306.5); A3 `=1/0` → `#DIV/0!`; A4 `=A1*2` → `84` (cached); A5 `hidden row` with row 5 hidden; A6 `Ünïcødé 漢字 🎉` (NFC letters plus U+1F389 as the surrogate pair D83C DF89); XFD1048576 `corner`; table `表格1` over D1:E3 with a header row (Name, Qty; a, 1; b, 2).
  - **vitest `tests/xlsx-viewer-native-excel.test.ts`: 10/10 pass (run 2026-10-13).** The object adapter (`parseXlsx` with the cached policy) and the paged adapter (`openImportSession` with an in-memory page store, 4-row pages) are built from the same bytes. Every check asserts that both adapters agree with each other and with Excel:
    - every authored cell has the same value, error flag, formula and cached provenance on both adapters;
    - `#DIV/0!` is the error lane class (tag 4), not text, in typed column lanes;
    - A2 keeps serial 45306.5 and formats to `2024-01-15 12:00` twice: through the resident sheet/style path and through the worker's style table plus `semanticsFor('1900')` (the `viewerFilterEvaluator` path);
    - A6 has Excel's exact UTF-16 code units;
    - `baseHidden` is exactly row 5 on both adapters;
    - the used extent reaches XFD1048576, and a row read there returns `corner`;
    - `queryStructural` returns one table D1:E3 with `headerRows: 1`, `totalsRows: 0`, and the resident metadata keeps the name `表格1` and the columns Name/Qty;
    - sheet name `工作表1`;
    - the default engine policy also reproduces A2/A3/A4;
    - the fixture hash is pinned, and the privacy edit removed only `absPath`.
    **No viewer/Excel disagreement was found.**
  - **Browser `tests/browser/interactive-viewer.spec.ts`, "native Excel native-corner.xlsx …": 6/6 pass [C+W] (run 2026-10-13).** The file is opened on `src/debug/viewer-e2e.html?src=…` in three legs:
    1. **resident**;
    2. **`?worker=1`**, which still admits this small file resident (asserted: no worker `open`);
    3. **worker-backed paged, `?worker=1&paged=1`** (asserted: exactly one worker `open`, and the saved-values status of a paged source is shown).
    Each leg checks:
    - the grid is labelled `Worksheet 工作表1`;
    - the Go To labels are `A1, value 42`, `A2, value 2024-01-15 12:00`, `A3, value #DIV/0!`, `A4, value 84` and `E3, value 2`;
    - there is no `Row 5` header, and ArrowDown from A4 lands on `A6, value Ünïcødé 漢字 🎉`;
    - Original-cell Go To A5 opens the read-only inspector;
    - `XFD1048576, value corner` has cell id 17179869183.
  - **`?paged=1` is a debug-only admission override** (`src/debug/forcePaged.ts`). Only with `worker=1`, the page loads through `loadOfficeSource` with `limits: { object: { cells: 1 } }` and the same paged route `useOfficeFile` uses (`openPagedWorkbook` on the counting worker), then renders `<OfficeDoc interactiveSheets showSearch>`. The worker learns `paged` from its `name` (`;`-separated segments, next to the S10.3 `caps:` flags). `src/debug/viewerWorker.ts` then replaces the entry's handler with one around a `ViewerKernel` that gets the same `importLimits`, because the worker repeats admission on the blob. No shipped code changed:
    - `src/debug` is excluded from the library build;
    - `bun run build` produced 216 byte-identical `dist/` files with and without the change (`dist/index.js` 133,767 B, `dist/worker.js` 27,134 B);
    - control run: with the worker-side override disabled, the paged leg fails because the saved-values status never appears (the worker opens an object source), so the leg does exercise the paged adapter in the worker.
  - **Second batch (2026-10-13): `native-1904.xlsx` and `native-large.xlsx`.** Both were made the same way: the user's Excel 16 for Mac was automated with their approval, each file was saved in Excel's container, and Excel reopened both with no repair prompt. Each got the same privacy edit, with only `xl/workbook.xml` changed (verified entry by entry).

    | File | Excel's original (bytes, SHA-256) | Committed (bytes, SHA-256, pinned) |
    |---|---|---|
    | `native-1904.xlsx` | 9,658, `12ccf94d54aec7bb85e781524fc8c9241b668ad727c23edef8a6ffd9c051fc84` | 9,499, `eebaea60f40ba35bef3e4941ad862317f71f8abfd68af2109f2d3a34942a81dc` |
    | `large/native-large.xlsx` | 1,800,677, `61a30b0bf2227f94a71849a625c7690335d4ce5ba9184dbc937d339414af8852` | 1,800,516, `32122deb5454d7666bed63867d15395094336b398d7c7b7ef027d8a74e5f7857` |

    - **`native-1904.xlsx`** has `date1904="1"`. A2:A7 hold Jan 15 noon, Jan 16 00:00, Jan 16 23:59:59, Feb 29 06:00, the epoch 1904-01-01 (serial 0) and Dec 31 2023 noon, formatted `yyyy-mm-dd hh:mm:ss`.
      - **vitest `tests/xlsx-viewer-native-excel-dates.test.ts`: 5/5.** The 1904 serials, formulas and labels agree on both adapters and with Excel. Display text equals Excel's for all six dates, including `2024-01-16 23:59:59` from `+0.999988426`, through both the resident and the worker formatting paths.
      - **AC-SEM-07 filters.** The worker filter evaluator, the production rule path for worker-owned sheets, ran over both adapters' stores. In source rows:
        - eq 2024-01-16 → Jan 16 00:00 and 23:59:59, not Jan 15 noon;
        - before 2024-01-16 → Jan 15, the epoch and Dec 31;
        - after 2024-01-16 → Feb 29;
        - between Feb 1 and the leap day Feb 29 → Feb 29 06:00;
        - eq 1904-01-01 → the epoch;
        - thisMonth / thisYear, with a clock injected at 2024-02-15 → Feb 29 / the four 2024 cells.
        Separately, `native-corner.xlsx`'s 1900 A2 (noon) matches eq 2024-01-15, between Jan 1 and Jan 15, and thisYear, and does not match eq 2024-01-16 or before 2024-01-15. Control run: with the date system forced to 1900, the 1904 case fails.
      - **Browser:** resident Go To labels show the 1904 dates as Excel displayed them [C+W].
      - Not an assertion: B4 was meant to hold a label, but Excel stored it as a number in built-in format 22 (it read the typed label as a date). The tests do not use B4.
    - **`native-large.xlsx`** (`tests/fixtures/xlsx-viewer/large/`, kept out of the directory-wide small-document sweeps of `tests/fixtures/xlsx-viewer/`, where it correctly opens in compatibility mode with the notice "Large file — opened in compatibility mode"): A1:C50000 hold `=ROW()*10+COLUMN()` and D1:D50000 hold `=DATE(2024,1,1)+ROW()+0.5`, which Excel saved as shared formulas. E50000 is `last`. That is 200,001 cells.
      - **vitest `tests/xlsx-viewer-native-excel-large.test.ts`: 5/5.** `admitSpreadsheet` with the shipped profile-D limits picks **paged**. The reason it records is the declared sheet-part size (`13179894 bytes (> 4194304)`, gate 1), which trips before the 100,000-cell cap is ever counted. The paged adapter reads Excel's values at A1/B1/C1 (11/12/13), A50000/B50000/C50000 (500001/500002/500003), E50000 `last`, sampled rows (A2, B12345, C25000, A33333, C49999) and D1/D2/D25000/D50000. D2 and D50000 display as `2024-01-03 12:00` and `2160-11-23 12:00`, and the object (compatibility) parse agrees at the sampled cells.
      - **Browser `native-large.xlsx` with plain `?worker=1` (no override): pass [C+W].** One worker and one `open` command; the saved-values notice is shown. Go To labels are `A50000, value 500001`, `D50000, value 2160-11-23 12:00`, `E50000, value last` and `D2, value 2024-01-03 12:00`. Header sort on A: the first click shows "View-only sort applied; formulas unchanged."; the second makes it descending, and display A1 reads `A1, value 500001; original A50000` (row 1 is numeric, so no header row is inferred). The transient pending state was not captured: the sort finishes too quickly to observe it reliably. The test asserts the commit acknowledgment instead.
      - **Observation, not an Excel disagreement:** under the cached value policy (paged and cached-object adapters alike), shared-formula children carry only their shared group (`sharedFormula.si`) and no expanded `formula` text. The master, A1 here, has the text. Both adapters agree, `hasFormula` is true, and the viewer shows no formula text. Excel's formula bar would show `=ROW()*10+COLUMN()` for B12345, so a future formula display would need to expand children.
  - **Still not-run:** Excel for Windows, and a native workbook beyond these three. The live-clock zone/DST presets were checked only with an injected clock.
  - **History (first attempt, superseded by the run above):** One automation attempt was made: `osascript` drove Microsoft Excel to create a workbook (number, `DATE()+0.5` with a datetime format, `=1/0`, a cached formula, a hidden row, Unicode text, XFD1048576) and to save it as `.xlsx`. The script reached the save step and then blocked for 120 s on a dialog that needed a user, most likely Excel's sandbox file-access grant: an Excel lock file appeared, but no workbook was written. A `quit saving no` was then sent to Excel. Excel process 699 stayed running afterwards, possibly the user's own session, so the user should check Excel for a leftover dialog or an untitled workbook (the coordinator has told the user). No native receipt is claimed. **Any future native automation must run under a dedicated macOS user account or in a VM, never against the user's live Office session.** No further Excel attempts were made.
  - **Manual steps (status):**
    1. ~~Build the cells and the table in Excel, save as `native-corner.xlsx`.~~ Done 2026-10-12 (automated with the user's approval).
    2. ~~Reopen, confirm no repair prompt, take a screenshot.~~ Done (see above).
    3. ~~Copy into `tests/fixtures/xlsx-viewer/`, record the SHA-256, add the adapter parity case.~~ Done (`xlsx-viewer-native-excel.test.ts`).
    4. ~~Open it in the e2e page (`?src=`), with and without `?worker=1`.~~ Done: resident, `?worker=1` (stays resident) and worker-backed paged through the debug-only `?paged=1` override, [C+W]. The cap-crossing copy saved by Excel came later: `native-large.xlsx` (second batch) opens worker-backed with no override.
- **Safari desktop: pass (S13), automated in real Safari 26.5.2 through safaridriver.** All 7 flows passed at `6dd5240`; see [S13](#s13-runtime-and-manual-checks), which also keeps the earlier locked-screen and harness-race runs. The Playwright WebKit runs above are not Safari evidence.
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

## S13 runtime and manual checks

**Branch:** `feature/iv-s13-manual` (cut from `feature/interactive-viewer` at `d4acf78`). **Run date:** 2026-10-12/13. **Machine:** macOS 26.5.2 (25F84), one external 1920×1080 display at DPR 1.

### Runtime DPR change — emulated (CDP) in desktop Chrome, not a physical display move

**Spec:** `tests/browser/interactive-viewer-dpr.spec.ts`. **Browser:** installed Google Chrome 155.0.8059.39 (`channel: 'chrome'`). The spec checks that CDP `Browser.getVersion` equals the `/Applications/Google Chrome.app` bundle version. **Run:** `IV_HARNESS_PORT=4191 bunx playwright test --config playwright.viewer.config.ts tests/browser/interactive-viewer-dpr.spec.ts --project chromium`. **Method:** `Emulation.setDeviceMetricsOverride` at a constant 1440×900 CSS viewport, `deviceScaleFactor` 1 → 2 → 1.5 → 1. Before the change each run sets up find (`A1048576`), Go To C3, a Shift selection to D4 and zoom 125%. Fixtures: `f1-small.xlsx` (resident) and `f1-paged.xlsx` (worker-backed). The machine has one DPR 1 display, so no physical display move was possible.

| Assertion | F1 resident | F1 worker-backed |
|---|---|---|
| **A DPR-only change re-rasterizes the viewer with no other input** | FAIL before the fix; **pass after** (2792×712, 2094×534, 1396×356) | FAIL before the fix; **pass after** (same) |
| Backing store = CSS size × DPR (±1 px), base and overlay, at 2 / 1.5 / 1 (after a re-render) | pass: 2792×712, 2094×534, 1396×356 for CSS 1396×356 | pass (same) |
| Tiles at the new DPR: each requested tile is 512 device px at zoom × DPR (logical box 204.8 / 273.07 / 409.6 at zoom 1.25) | pass | pass |
| No stale-DPR tile presented once re-rendered: the first frame at the new DPR reuses 0 tiles, and every later frame paints only new-DPR tiles | pass | pass |
| Gridline alignment: the darkest device column/row lies within ±1.5 device px of the empty cell G12's DOM (hit-test) right and bottom edges (expected x 450 / 337.5 / 225, found 450 / 337 / 225; cell centre stays white) | pass | pass |
| State kept: active cell, selected cells, first row/column header, zoom text, find text, status text | pass | pass |
| Bitmap budget: every frame's `bitmapBytes` and `peakBitmapBytes` ≤ 134,217,728 B; tile bytes = tiles × 1 MiB (peak observed 41,069,056 B at DPR 2) | pass | pass |
| No console errors or page errors | pass | pass |

**D-DPR (found by the FAIL row, now fixed).** After each override, `devicePixelRatio` and `matchMedia('(resolution: …)').matches` changed at once. Before the fix, the viewer then reported no frame for the 5 s window, and the canvases kept the old backing size (for example 1396×356 at DPR 2). The cause was in `InteractiveSheetDoc.tsx`:

- the viewer read `window.devicePixelRatio` only while rendering;
- its only size input was a content-box `ResizeObserver` on the pane;
- nothing listened for a `(resolution)` media-query change.

A DPR-only change keeps the CSS size, so nothing woke the viewer, and the old raster stayed on screen, scaled.

**Fix:** a new `src/components/xlsx/devicePixelRatio.ts` exports `useDevicePixelRatio`. The `InteractiveSheetDoc.tsx` change is the hook call plus its import. The hook re-reads the ratio when any of these fire:

- a `(resolution: Xdppx)` media query, re-armed for each new ratio;
- window `resize` and `visualViewport` `resize`;
- a 1 s poll while the page is visible.

Setting an unchanged value does not re-render. Unmounting removes every listener and clears the poll. A new ratio changes the pane surface bytes, so the tile pool is re-admitted (old-DPR tiles are released). `present()` also drops any tile keyed to another DPR. Unit tests: `tests/xlsx-viewer-dpr.test.tsx` (5 tests). Before the hook was implemented, a deliberately wrong boundary failed 4 of them with the intended assertions, for example `expected [] to deeply equal [ '(resolution: 1dppx)' ]` and `expected 1 to be 2`. All 4 browser tests now pass with unchanged bodies. The only spec edit is a file-level skip, so the webkit project no longer tries to launch with `channel: 'chrome'`.

The other rows are checked after a pane width change of 100 px and back, so they also cover the resize path.

*Emulation note:* under CDP, Chrome changed the DPR and media-query state but fired no `resize`, no media-query `change`, and no `device-pixel-content-box` ResizeObserver callback (checked with a probe). So in this spec it is the hook's visible-page poll that fires (logged page signals: `[]`). A real display move normally fires the media-query change. That path is covered by the unit tests only, because this machine has a single DPR 1 display.

**The 1 px nudge observation: not reproducible as a pane-width bug.** I tried to reproduce it, before and after the fix. Before the fix, 72 cycles of DPR 2 / 1.5 / 1, each with a 1 px pane nudge and back, produced 3 stale states. All 3 were old-DPR backing stores with the correct CSS width, which is D-DPR (or the two resizes coalescing into one observation). The earlier "CSS width stuck at 1397" state did not recur. After the fix, 60 cycles that waited for each leg plus 60 cycles with 0–57 ms gaps produced 0 stale states (canvas CSS width = pane `clientWidth` − 44, backing = `ceil(CSS × DPR)`). No `devicePixelContentBoxSize` change was made, and no regression test was added, since there is no reproduction. The probe was a throwaway and is not committed.

### Actual Safari via safaridriver — pass (7/7 at `6dd5240`)

**Script:** `bun scripts/harness/safari-check.ts` (W3C WebDriver over HTTP to `safaridriver -p 4451`; the harness vite server on 4191; a loopback fixture server on 4193). It writes the receipt and screenshots to `/tmp/officeview-safari-check/`. **Browser:** Safari 26.5.2 (safaridriver "Included with Safari 26.5.2 (21624.2.5.11.8)"), session `platformName: macOS`, `browserVersion: 26.5.2`. This is real Safari, not Playwright WebKit.

**Early runs (screen locked).** The script first ran twice while the macOS screen was **locked** (`CGSSessionScreenIsLocked = true`). The first run failed on a harness bug before it reached the page: the W3C element-reference key was wrong. The bug is fixed, and that run's results are discarded. The second run's receipt is kept as `/tmp/officeview-safari-check/receipt-locked-screen-run.json`. In that run every WebDriver screenshot was black, and WebDriver clicks reported success but had no effect (Go To never moved the active cell). The likely cause is the locked screen; an unlocked rerun has to confirm it. Those interaction results say nothing about the viewer. The script now refuses to run on a locked screen.

| Step | Result |
|---|---|
| Capability detection, main thread: `Worker`, `OffscreenCanvas` with a 2D context, IndexedDB open, `navigator.storage.getDirectory`, estimate (quota 82,463,372,084 B); **OPFS `createWritable` on the main thread wrote 3 bytes** | ran: all available |
| Capability detection, dedicated worker: `OffscreenCanvas`, IndexedDB, OPFS `createSyncAccessHandle` (wrote and flushed 1 MiB) and `createWritable` both present | ran: all available |
| F2 storage package (160,000 rows; pages exceed the 32 MiB memory store), worker-backed: the WebKit OPFS-write restriction | ran: the import **completed** ("Visible body 160000 / complete body 160000"). The page-write failure seen in Playwright WebKit (S10.3 capability matrix) did not happen in real Safari 26.5.2. The locked-run script waited for the wrong row count (159,999), so the receipt marks the step as a timeout, but its `statusAtTimeout` holds the line above. The count is fixed. |
| Opens: F1 resident (no worker), F1/F2/F3 worker-backed (1 worker, 1 `open`), each reaching "complete body" | ran: opened |
| Go To XFD1048576, find A1048576, copy payload, keyboard navigation, zoom, header sort, value filter apply/cancel/clear, F3 recenter | **not-run (invalid):** the clicks returned success but had no effect, most likely because of the locked screen. Rerun with the screen unlocked |

**Unlocked run (integrator, at `da9eac8`; receipt in `/tmp/officeview-safari-check-unlocked/`): 5 pass, 2 fail.**

Passed:
- capabilities;
- F1 resident and worker-backed: Go To XFD1048576, find the never-painted corner, copy payload;
- F3 dense recenter;
- the F2 storage import of 160,000 rows.

Both failures were harness bugs, confirmed by re-running those steps in real Safari:
- **"waiting for active cell A1":** Safari reserves Cmd+Home. The WebDriver Meta+Home navigated the tab to the Safari home page, so the page never saw the key. Ctrl+Home moved the active cell to A1 at once. Cmd+Down and Cmd+Right reached A1048576 and XFD1048576.
- **"alpha checkbox not found":** the filter popup's buttons stay disabled until its value list loads (about 1.7 s for 100,010 rows), and a WebDriver click on a disabled button does nothing. The value checkboxes are named correctly (`aria-label="alpha"`).

Fixed in `606db4c`: Ctrl+Home, and clicks wait for an enabled target.

**Integrator re-runs at `1143e8b` (the user had cleared a stray macOS permission dialog first):**
- **Run 1: 6 pass, 1 fail.** "F1 worker-backed paged … corner label XFD1048576, value Loading…" was a harness race: the step read the corner's name while the worker read was still in flight, and "Loading…" is the correct pending state.
- **Run 2: 7 pass.**

The harness fix (`6dd5240`) waits up to 30 s for a settled name and asserts that the labels went pending → value, never through a blank value.

**Final run at `6dd5240`: 7/7 pass** (2026-10-12T18:32:21Z–18:32:39Z, Safari 26.5.2, safaridriver 21624.2.5.11.8, ports 4485–4487, receipt `/tmp/officeview-safari-check-final/receipt.json`). Every click was a real WebDriver click; no synthetic fallback was used. Key values:

| Step | Result |
|---|---|
| Capabilities | Worker, OffscreenCanvas 2D, IndexedDB, OPFS (main-thread `createWritable`; worker `createSyncAccessHandle`): all present |
| F1 resident | corner "XFD1048576, value 17179869183.00"; find reached A1048576 (id …17179852800); copy payload "A1048576" |
| F1 worker-backed | same corner and copy; 1 worker, 1 `open` |
| Keyboard | B3 → C3 → C4 (Enter) → A4 (Home) → A1 (Ctrl+Home) → A1048576 (Cmd+Down) → XFD1048576 (Cmd+Right) → XFC1048576; zoom 100% → 125% (`+`) → 100% (`0`) |
| F2 worker-backed, 100,010 rows | sort G ascending then descending; G2 = "G2, value 200017.00; original G100009". In-page click → first status/`aria-sort` change: 1,538 ms (ascending) and 2,975 ms (descending); this is a WebDriver observation, not an S10.2 receipt. Filter D = alpha: "Visible body 24972 / complete body 100009"; Cancel left it unchanged; Clear all filters gave "Visible body 100009". Worker commands: open 1, sort 2, filter 2, prepareColumns 2 |
| F3 worker-backed dense | recenter at XFD1/A1/XFD8/H3 with no orphan cells; ArrowLeft from XFD8 → "XFC8, value 1984.20" |
| F2 storage, 160,000 rows | "Visible body 160000 / complete body 160000"; OPFS pages persisted, no write restriction |

### macOS VoiceOver — guided manual pass (real VoiceOver, real Safari 26.5.2)

**Automated script:** `bun scripts/harness/voiceover-check.ts` (`--preflight` only checks and never starts VoiceOver). It is still not usable: VoiceOver starts, but `tell application "VoiceOver" to return content of last phrase` hangs until the osascript timeout. The user is re-checking the AppleScript/Automation permissions. No automated VoiceOver receipt is claimed.

**Guided pass (run by the integrator with the user at the machine).** Fixture `f2-1k` (resident), real keyboard, VoiceOver on, macOS VoiceOver UI in Traditional Chinese. Evidence is the VoiceOver caption panel, cropped to the panel only, in [`voiceover-s13/`](voiceover-s13/). In the captions, 第N欄 = column N and 第N列 = row N, both chrome-inclusive. These ran **before** the V1/V2 fixes below.

| Step | VoiceOver said | Result |
|---|---|---|
| Down onto A2 ([01](voiceover-s13/01-down-a2.png)) | "A2, value 1.00 1.00, column 2, row 3" | pass for name and chrome-inclusive indices; **V1**: value spoken twice |
| B3 date cell ([09](voiceover-s13/09-b3-date.png)) | "B3, value 2023-06-09 2023-06-09, column 3, row 4" | pass for indices; **V1** |
| Right ([02](voiceover-s13/02-right.png)) | "C3, value TRUE TRUE, column 4, row 4" | pass for movement; **V1** |
| Shift+Right ([03](voiceover-s13/03-shift-right-d3.png)) | "D3, value beta beta, column 5, row 4" | movement passes; **V2**: no selection announced |
| Ctrl+Home ([04](voiceover-s13/04-ctrl-home-status.png)) | live status "Visible body 1000 / complete body 1000 — A1, value id" | pass |
| Option+Down on a cell ([05](voiceover-s13/05-option-down-no-popup.png)) | status only, no "Column filter" popup | **V3**, see below |
| Escape ([06](voiceover-s13/06-escape.png)) | "A2, value 1.00 1.00" | pass |
| Tab out of the grid ([07](voiceover-s13/07-tab-1.png), [08](voiceover-s13/08-tab-2.png)) | "Search document, Search…, search text field", then the buttons | pass (one Tab stop) |

**Findings, resolved on this branch:**

- **V1 — every value spoken twice: real defect, fixed.**
  - Cause: the gridcell's `aria-label` already carries the value, and its visually hidden text child was still exposed, so VoiceOver read the name and then the contents.
  - Fix (at the time; superseded by V5, where the name now comes from the exposed text itself): the text child was made `aria-hidden` (`GridSemantics.tsx`). The accessible name is unchanged ("B3, value …"), so the S10.3 name contract and Chromium AX tests still pass.
  - Test: `tests/xlsx-viewer-a11y-announce.test.tsx`, which failed first with `expected [ '10' ] to deeply equal []`.
- **V2 — no selection announced on extension: real gap against spec §9.2, fixed.** §9.2 says cell names carry selection descriptions and the live region reports selection.
  - Shift+Arrow and Shift+click now announce the range, e.g. "D3, value beta; selected B3:D3".
  - Cells of a multi-cell selection get an `aria-description` of "selected". A single active cell does not, to avoid noise.
  - Tests are in the same file. They failed first with `expected 'C2, value 10; original C4' to be '…; selected A2:C2'`.
  - Not yet re-heard with VoiceOver.
- **V3 — Option+Down did not open the filter under VoiceOver: not reproduced as a viewer defect.**
  - With VoiceOver **off**, I sent real OS key events (System Events `key code 125 using option down`) to a normal Safari 26.5.2 window. A key probe page recorded `key: "ArrowDown", code: "ArrowDown", altKey: true`.
  - On the viewer, with the grid focused via the e2e page's new `focus=grid`, the same keystroke opened "Filter column B".
  - So the handler works in Safari. Under VoiceOver the keystroke did not reach the page as Option+Down, most likely because VoiceOver handles it. That is not confirmed: no VoiceOver key trace was possible.
  - Keyboard alternative that VoiceOver users can reach: F2 (header control mode) → Tab to "Filter column X" → Enter (S10.3 receipt). Re-test Option+Down with VoiceOver's pass-through (pass next keystroke) on the next VoiceOver pass.
  - Separately, safaridriver's synthetic modified arrows arrive with control-character `key` values (Option+Down → `"\u001f"`, Shift+Right → `"\u001d"`; `code` is correct). That is a WebDriver synthesis artifact; real keys were correct. The Safari check therefore uses no modified arrows except Cmd+Arrow, which arrives correctly.
- **V4 — "1.002023-06-02" in the date column: not Safari-specific; fixed for Excel fidelity (`3b759df`).**
  - Playwright Chromium and WebKit rendered it identically. With Calibri missing, "2023-06-02" measures about 68.2 px in a 64 px column. The renderer clipped it at the cell edge, so it read as joined to column A's "1.00".
  - Excel instead shows `#` repeated across the cell for a number, date or time that does not fit. `layoutCellBody` (`src/xlsx/render.ts`), shared by the static and viewer renderers, now does the same:
    - it measures with the same font string the cell paints with;
    - the cell gets ⌊interior / width("#")⌋ `#`s, and at least one;
    - General keeps its order (fewer decimals, then scientific), and its "nothing fits" result now fills the cell too.
  - Unchanged: text cells, booleans, errors, wrapped and rotated cells, and numbers shown through a text (`@`) section. The accessible and find value (`cellDisplayText`) stays the real value.
  - At `3b759df` shrink-to-fit was not yet parsed. It is now (see below), and a shrink-to-fit number reaches `#` only when even the floor size does not fit.
  - Since review round 3, `#` eligibility comes from the value type and the **selected** format section (`numericSectionIsText` in `formula/format.ts`), not from the output string. General's "nothing fits" outcome is a separate signal, so an authored literal `"#"` or `\#` paints as written.
  - Tests: `tests/xlsx-number-overflow.test.ts`, covering dates, times, negative numbers, percentages, General, a 1-glyph column, text and accessible value. It failed first with, e.g., `expected [ '2023-06-02' ] to deeply equal [ '########' ]`.
  - Two format tests now render with wide columns or narrower fake glyphs, so they keep asserting formats rather than overflow.
  - **Golden:** corpus compare after hydrating with `bun scripts/fetch-corpus.ts` changed exactly one XLSX golden: `poi-64508.xlsx.0` (2,720 / 44,800 px). Its `dd.mm.yyyy` date column used to spill garbled digits into the next column and now shows `######` in the cell. Re-recorded.
  - Six docx/pptx goldens also differ (`poi-55733.docx.0`, `poi-picture-transparency.pptx.1–3`, `pydocx-blk-inner-content.docx.0`, `pypptx-test_slides.pptx.0`). They fail identically on the base commit without this change, so they are unrelated and not re-recorded.
  - With real Calibri metrics, a column that fits in Excel would also fit here. The harness has no Calibri, so this date column shows `#`s.
- **V5 — "blank" (空白) after every cell after the V1 fix: real defect, fixed; historical guided re-check at 0c65bff, final-tree re-check deferred.**
  - Heard on the 2026-10-15 re-check at `9b35131` (Safari 26.5.2, user present): V1 is fixed (value read once) and V2 is confirmed ("selected" on Shift+Arrow), but every cell ended with "blank", e.g. "B2, value 2023-06-02，空白，第3欄，第3列".
  - Cause: with the inner text `aria-hidden`, the gridcell had an `aria-label` but no accessible content, so VoiceOver announced it as blank.
  - Fix: the gridcell has no `aria-label`; its visually hidden text is exposed and carries the full name ("B2, value …; original …"), so the name comes from the cell's own contents. The name contract is unchanged and the value occurs once. Browser specs and harness scripts read the cell name from `textContent`.
  - Test: `tests/xlsx-viewer-a11y-announce.test.tsx`, which failed first with `expected '' to match /^[A-Z]+\d+(, value .*)?/` (no exposed text).
  - **Subsequent receipt:** the guided pass at 0c65bff re-heard V5, as recorded in the blocking-list update below. No automated VoiceOver receipt is claimed; the final-tree re-check remains deferred by user decision.
- **Text overflow into non-empty neighbours: fixed for Excel fidelity.** Excel lets text overflow only into empty adjacent cells: rightward for left/General alignment, leftward for right alignment, both ways for center. It clips at the first non-empty cell.
  - What blocks: a value, a formula (even one returning ""), an empty-string constant, and any merged cell, including a merge anchored in another row. The sheet edge stops it too.
  - Hidden (zero-width) columns are skipped: they neither block nor add room. Merged cells never overflow. Formatting-only cells count as empty.
  - Rule and placement:
    - `src/xlsx/view/textOverflow.ts` holds the pure rule, `textOverflowAllowance`.
    - The static renderer computes each row's occupancy once.
    - The viewer planner (`viewport.ts`) computes it inside its fetched overflow envelope. Beyond the envelope the room is unbounded, because the pane never reaches there.
    - `layoutCellBody` turns the allowance into a horizontal clip: `clipX` on plain text, a widened `clip` on unwrapped explicit layouts.
    - The allowance is part of the viewer's layout-cache key.
  - Tests:
    - `tests/xlsx-text-overflow.test.ts` covers left/right/center; blocking by a value, by `""` (formula and constant) and by a merge; the sheet and left edges; hidden columns; formatting-only cells. It failed first with `expected null to deeply equal { x0: 0, x1: 100 }` and similar.
    - A viewer/static parity crop in `tests/xlsx-viewer-layout.test.ts` gives diff 0, with clips at the expected column edges.
- **Shrink-to-fit: implemented.** `alignment@shrinkToFit` is now parsed (`parse.ts`, `pagedFormatting.ts`; `XlsxCellStyle.shrinkToFit`).
  - For unwrapped, unrotated cells, the font scales down in 0.1 pt steps until the formatted text fits the interior. General keeps full precision rather than dropping decimals.
  - A number shows `#` only when even the 4 pt minimum does not fit; text stays at the minimum.
  - Find and accessible values (`cellDisplayText`) are unchanged.
  - Tests (same file): numbers, dates, text, text that already fits, the below-minimum `#`, and the accessible value.
- **Goldens for these two rules:** after the fixes, the corpus compare changed three XLSX goldens, all re-recorded. In each, overlapping text from a cell is now clipped at its non-empty neighbour, as Excel does:
  - `poi-59021.xlsx.0` (1,004 / 20,480 px): "Abkhazia - Fixed" over "7840", and "10/02/2016" over F2, which holds an empty-string constant.
  - `poi-64508.xlsx.0` (3,519 / 44,800 px): headers and "SI33738548"/"04656-2020"/barcodes over their neighbours.
  - `poi-xlsx-jdbc.xlsx.0` (1,080 / 7,680 px): "A Very large string in a column" over "foo"/"bar".
  - The same six docx/pptx goldens still differ exactly as on the base commit (unrelated).
- **Review round 3 fixes (Codex):**
  - **One overflow geometry.**
    - A text record's clip is now the painted clip, translated from the box-local body; centered and right-aligned text included.
    - Find highlight geometry rebuilds the record with the same allowance the painter used (`viewport.cellTextOverflow`, which uses `rowOverflowInput` on the row's real cells and merges). It then intersects highlight rectangles with that clip.
    - A match whose ink is fully clipped stays a hit and highlights its own cell.
    - The viewer no longer uses an envelope: allowances end at the real worksheet edges (A and XFD).
  - **Shrink-to-fit** never enlarges an authored size. At the 4 pt floor, numbers show `#` and text is held to its own cell; it never spills.
  - **Fill, justify, distributed and center-across-selection alignments** never overflow (`alignmentOverflows`).
  - **Paint order.** The static renderer and the viewer paint every fill first, then text and borders. Overflow into an empty formatted neighbour is drawn over its fill, as Excel does. The viewer test checks for text ink inside a yellow empty B1.
  - **Selection wording and cardinality.**
    - Extending a header selection announces "selected columns A:C" or "selected rows 2:3", plus any additional ranges.
    - Cardinality uses each row-set's own count (intervals, list, bitset, ref), so a one-cell bitset stays quiet.
  - **Implicit cell positions** (found while checking the paint-order change). The resident parser placed every `<c>` without `r` at A1 and every `<row>` without `r` at row 1, stacking the cells on top of each other. It now follows Excel: a ref-less row follows the previous row, and a ref-less cell follows the previous cell of its row (`tests/xlsx-implicit-refs.test.ts`).
    - One existing assertion that pinned the legacy row index (0) for a ref-less row moved to Excel's position (1); its cell's explicit reference and the inconsistency diagnostic are unchanged.
    - The streaming (worker) parser already refuses a ref-less row (`a worksheet row has no valid row number`) or cell with a visible "unsupported" outcome.
    - **Golden re-recorded:** `poi-59746_NoRowNums.xlsx.0`, the POI fixture for exactly this case. It was 64×20 px showing one garbled A1 cell and is now 704×300 px showing its 15×11 table.
    - The three goldens above are unchanged by the round-3 fixes. The same six docx/pptx goldens still differ exactly as on base.
- **Review round 4 fixes (pi):**
  - **Implicit positions are bounded.** An implicit position past the worksheet is a typed `PackageError('unsupported', …, part)` refusal, like the paged reader's, never an uncaught `CoordinateError`. That covers a 16,385th ref-less cell in a row, a ref-less row after row 1,048,576, or a ref-less cell in a row beyond the sheet. One full row of 16,384 ref-less cells still parses. Excel reports such a file as needing repair, so no clamp is attempted. Regression: `tests/xlsx-implicit-refs-bounds.test.ts`, the reviewer's repro.
  - **Duplicate explicit rows.** A repeated `<row r>` is merged into its first record, with later cells replacing same-column ones, and the sheet gets a `duplicate-row` diagnostic. The paged reader refuses the same file (`worksheet rows are out of order or duplicated`). This divergence is intentional: refusing would break resident files that open today.
  - **Allowance cost.**
    - Occupancy is reduced once per row to the sorted, visible blocking columns (`rowOverflowFromCells`).
    - The viewer now uses one document-owned byte LRU across all sheets and current/retired geometries. Its retained 2 MiB subcap is charged to the real shared main ledger (layout owner), within the unchanged D 64 MiB / M 32 MiB budget. Construction and coexisting copies are reserved in that same budget before allocations.
    - Blockers and merge endpoints use lossless Uint16 storage. Weak snapshot/geometry/row ownership and weak eviction callbacks do not retain dead cells, text or projections; document/source/profile replacement and close invalidate memos and release their charge. Unbound geometry computes without retaining a global memo.
    - Admission failure produces a coarse memory notice; find skips unavailable highlight geometry with that notice. Refusal tiles are not retained, so later admission can rebuild exact allowances. The ownership, pressure, snapshot/reload, close, eviction parity and corrected Bun/JSC job-boundary GC regressions are maintained in `xlsx-overflow-owned-cache`, `xlsx-overflow-ledger` and `xlsx-text-overflow-cache`. This code received independent Codex subagent SPEC and QUALITY technical APPROVE (OpenAI GPT family). The different-model-family review is done: Claude (claude-opus-5-5, a different family) approved round 7 on 2026-10-15 at `9b35131` with NITs only (`findings-s13-manual-claude-r7.md`), and the user's 2026-10-15 decision keeps the strict different-family rule. No new browser performance or native receipt is claimed.
    - Each allowance is a binary search plus the row's merges.
    - Index reads for one allowance no longer grow with the formatting-only cells before the blocker. The reviewer's 16,384-wide styled row case is now bounded (`text-overflow allowance cost on a wide styled row`).
    - **Claude r7 NITs (2026-10-15).** Pinned by tests that fail on a mutant: a resident (no pending) refused tile is never retained and repaints detail once pressure ends (`xlsx-overflow-refusal-render`); an over-wide row repeating one column 16,384 times keeps a later blocker (`xlsx-overflow-owned-cache`); an unsupported-format diagnostic is published only inside an admitted `retain` (`xlsx-viewer-copy`). Fixed: a closed workbook hands out one closed, uncharged overflow owner, so the viewer refuses instead of throwing from render/memo/layout effects; present and find keep separate overflow notices, so an unrelated base present no longer clears a find refusal (`xlsx-overflow-owner-lifecycle`, `xlsx-overflow-ledger`). The frame-time rebind in `frameOf()` is kept for sibling-unmount/reset recovery. **Subsequent landing-seam fix (merged by 485a57c):** every shared-ledger consumer now uses `remote?.profile ?? profile`; differently configured sibling viewers no longer alternate the remote overflow owner's profile each frame. Closed-owner rebinds and mixed-profile caps are pinned in `xlsx-overflow-owner-lifecycle.test.tsx`.
  - **Copy ownership and aggregate admission.** A copy job charges its own `copy` owner in the same document main ledger as the overflow owner (`mainLedger = remote?.mainLedger ?? new MainMemoryLedger()`). Growth is admitted only while the job stays within its session cap (`copySessionBudget`, 3/8 of the profile's main budget) **and** the ledger's aggregate total (all pools, minus this job's held bytes, plus the new size) stays within `MAIN_THREAD_BUDGETS[profile]`; shrinking skips the aggregate check (a closed job or one over its session cap is still refused). Column reservation uses the same admission. Every temporary formatter and serializer byte is refunded; only per-block metadata survives until the result is built. A new canonical format diagnostic is adopted from the copy owner into a reachability-bound `document` charge (the copy charge stays column bytes plus live memory; duplicates net zero). A ready result stays charged until the native clipboard write settles (about twice the text while the write is pending, deliberately conservative). Overlapping jobs have independent owners; unmount aborts the current job, and each owner is released when its job settles. Tests: `xlsx-viewer-copy`, `xlsx-viewer-copy-lifecycle`.
  - **Merges per source row.** The static renderer keys each cell's occupancy on that cell's own source row, so a malformed row whose cells name other rows uses the right merges.
  - **Reviewer's unconfirmed suspicions:**
    - *Active-record z-order (confirmed and fixed).* The exact active record kept inside a coarse region was replayed after gridlines and drawings, and its border was never painted. It now joins the records' layer, and a coarse frame paints its border. Test: `review round 4: the exact active record…` in `tests/xlsx-viewer-layout.test.ts`.
    - *`laidOut` retention (not changed).* The static renderer keeps one layout per painted cell between its fill and content passes. That is bounded by the cells of the admitted raster (static admission already caps the raster), and avoiding it would mean laying every cell out twice.
    - *Lowercase or alternate-content refs (not changed).* A lowercase `r` (`a1`) is treated as missing by both the resident and paged readers (same `[A-Z]+` rule), so they agree. Excel writes upper case.

## S13.5 on-demand display-text lane

Plan S8.1 ("Prepare the formatted display-text lane … only when a text condition targets a text/mixed column"), spec §7 and §12.4. Branch `feature/iv-s13-5`, cut at `4e4f0e1`.

- **Before:** both filter preparers (`src/worker/viewerFilterEvaluator.ts` and `src/components/xlsx/viewerControlsClient.ts`) called `cellDisplayText` for every numeric/date/boolean value. They did so in classification and in every rule evaluation, whatever the rule kind.
- **After:** classification, number/date conditions and value lists read raw values. A text condition (or a compound with a text child) on a text/mixed column reads display text from `src/worker/displayTextLane.ts`:
  - the lane is keyed by (docId, sheetId, valueGen, style owner, column, body band);
  - a lane of a dead valueGen or style owner is dropped on the next request;
  - storage is one Uint32 id per row plus unique strings, charged like a prepared column;
  - a lane an evaluation is reading is leased: it is never evicted for room, and its worker-pool entry is pinned;
  - exhaustion is `ValueIndexError('quota')`, which surfaces as a typed `resource-limit` refusal.
- **Worker budget: one prepared-column budget and counter.** Lane bytes, both building and retained, join the session's `preparedBytes`. This is the same counter and the same `PREPARED_BUDGETS[profile].maxBytes` (64 MiB D / 32 MiB M) that sort/filter prepared columns use.
  - Lane admission and prepared-column staging both make room by evicting unpinned columns and unleased lanes, least recently used first across both kinds.
  - Either refuses without evicting anything when the pinned working set leaves no room.
  - A lane is a dense per-row lane, so it also holds one `maxColumns` slot from the start of its build until it is uncharged; admission evicts by the same recency until both `maxBytes` and `maxColumns` fit.
  - Raw staging of a running `prepareColumns` batch (bytes and one slot per column being or already staged) is part of the same admission. A lane build sees it, and the batch re-reads the unevictable set (pinned/batch columns, leased, retired-but-leased or building lanes) at every staging step, so a lane and raw staging never overlap past `maxBytes`/`maxColumns` in either order, across sheets of one session (codex r3 #1; `xlsx-viewer-display-lane-lifecycle.test.ts`). Each staging step also evicts idle retained columns and unleased lanes (least recently used, never the batch's own columns) before the staging grows, so retained plus staged bytes and slots stay within `maxBytes`/`maxColumns` for the whole batch. A later cache hit therefore pins a lane that already fits beside the staging, or, once evicted, the rebuild is refused (codex r4 #1). An idle entry evicted for a batch that later stops (quota/cancel/stale) is not restored; it is rebuildable.
  - A lane invalidated (valueGen/style owner) or released by the cache while a filter still reads it leaves lookup at once but stays charged (bytes, slot, main-thread controls account) until its last lease ends.
  - Worker session close (kernel `close`, after aborting in-flight work) refuses new lane reservations at once. A lane a filter still leases moves its bytes from the closing pool entry to a transient charge and stays in `preparedBytes` until its last lease ends. A lane still being built keeps its per-filter transient share until the build ends. A closed pool cannot shrink a share, so a filter's share is released whole when its builds end (codex r3 #2). Likewise a raw prepared column a sort reader still leases moves its bytes to a transient charge and stays in `preparedBytes` until that lease's release (normal, error or cancelled sort); close itself never waits for the reader (codex r4 #2, inherited from before S13.5).
  - Building bytes also hold a transient worker-pool share. A completed lane moves its bytes onto a `rebuildable` pool entry, its single uncharge point: owner drop, session or pool eviction, close (no lease), or its last lease after close.
- **Main-thread budget: the controls share of the main ledger.** In the in-process (resident document) path the lane reserves its bytes from the controls value-index cache. That cache is 16 MiB D / 8 MiB M (`createViewerControlResources`), shared with value indexes and committed key runs, and its usage reports into the main ledger's `controls` pool (S10.2). So the lane is charged for exactly its retained lifetime, inside the controls caches' existing share of the §12.4 main budget (64 / 32 MiB), never the whole pool.
  - When the share is full, idle value indexes are shed first, then unleased lanes; otherwise the filter refuses.
  - The port's lanes are released on dispose.
  - Worker-backed sheets filter in the worker and build no main-thread lane.
- **Not changed:** the dropdown's value-search labels were already formatted per entry on demand (`domain.label`) and stay so. Value-list membership uses raw keys and never needs the lane.

**Bun microbenchmark (non-browser; not an SLA receipt).** Run with `bun scripts/bench-display-lane.ts [root] [label]` on the same machine, with a 1M-row object source and a cold evaluator per sample (median of 4). The base was measured with the same harness on `4e4f0e1`.

| Case | Base `4e4f0e1` | S13.5 |
|---|---:|---:|
| Number `gt` on a 1M-row `0.00` column | 1,298 ms | 117–139 ms |
| Date `between` on a 1M-row date column | 363 ms | 93–116 ms |
| Text `contains` on a 1M-row mixed column, apply 1 / 2 / 3 (one evaluator) | 1,531 / 1,434 / 1,450 ms | 1,000–1,253 / 194–228 / 193–215 ms |

The text lane for that mixed column charges 9.52 MiB (9,978,256 B). That fits the desktop prepared budget (64 MiB) and the desktop controls share (16 MiB); on the mobile main-thread controls share (8 MiB) it would refuse, but a 1M-row sheet is worker-backed and uses the worker budget. The bench drives the evaluator standalone, with a local 64 MiB ledger and no session. Numeric/date filters charge nothing and make zero `cellDisplayText` calls; the tests pin this with a spy. Heap deltas around one apply were noise-dominated (transient page objects), so no retained-memory saving is claimed beyond the lane ledger.

## PR #11 visual review (2026-10-16/17)

PR #11 opens from `feature/interactive-viewer` (user decision 2026-10-16). Its visual-review work is recorded here and in the plan §0.4; the performance side (iOS Simulator directional baseline, refused-by-design text sort) is in `interactive-viewer-performance.md`.

**Audit.** Claude's screenshot audit (142 PNGs across resident/worker-backed/static/print/WebKit at `06e9137`, `/tmp/iv-visual/`) and agy's candidate list produced the 25 findings in `.worktrees/_prompts/visual-findings-claude.md` and `findings-visual-agy.md` (numbered 1–21 plus NEW-1..4). User decision 4 **requires all 25 to be addressed** in PR #11 under an eager lazy-load-only rule (it did not itself fix them): eager `dist/index.js` gzip stays within the 38,720 B spec cap (base 38,699 B), each stream reporting exact bytes per commit; frozen panes required a spec amendment because §1.2 lists them as a non-goal.

**Merged in PR #11 (with the reviewer who merged each stream):**

| Work | Merge | Reviewer |
|---|---|---|
| Bot/source fixes round 1 (through base `06e9137`) and round 2 (`03e6342`: source-removal reset, one-shot route pinning, outer-wrapper zoom suppression, scoped/owned print rules, narrow search row, 44px coarse-pointer targets, print-media test; §5.5 spec amendment `e8d9ca6`) | in base + `03e6342` | agy (round 1, APPROVE) and codex (round 2, r1 REQUEST CHANGES → r2 APPROVE) |
| CI first-paint fixes (test-only `tests/helpers/viewer-timing.ts`; first-gridcell waits 10 s after mount/open in 9 files, timing/complexity windows untouched) | in base `56a8ee5..` line | opencode APPROVE (`findings-ci-fix-opencode.md`) |
| Copy-lifecycle CI flake pin (deterministic copy-capacity pin in the budget-boundary tests; test-only) | `7ee282d` | opencode APPROVE (`findings-copy-flake-opencode.md`) |
| Stream D — RTF/ODT/DOCX/PPTX visual fixes (#15, #16, NEW-1..3: windows-1252 table decode on every runtime, text measured in the painting canvas's language, `w:tblInd` table indent, `jc start/end` bidi resolution, Writer Heading 1, inline-picture descent; NEW-4 PPTX `a:highlight` (`1cbf514`, warped/rotated highlighting excluded and disclosed in the commit); #17: the empty-paragraph bullet removed (`ccfc75d`), the picture grey frames verified authored in the fixture's own `<a:ln>` outline and intentionally retained, the golden's omission accepted as a reference limitation) | `dad4b8d` | opencode APPROVE (`findings-vis-docs-opencode.md`; accepted re-record of poi-rtl/poi-table-indent/poi-text-highlight goldens, and the authored-grey-frame and excluded-warped/rotated-highlight dispositions) |
| Stream A — interactive viewer UX (#6 viewer fills its flex host, #9 header glyphs, #12 print chrome under scoped/owned rules, #13 filter-panel/search layout, #18 evaluator-limit notice on small files, #19 control font, #20 find highlights the matched span, #21; horizontal scroll moves grid rows by transform only) | `72034e8` | agy APPROVE (`findings-vis-viewer-agy.md`; tsc, vitest 6,709, build eager 38,701 B, Playwright suite 90 passed) |
| Stream B — XLSX cell-rendering correctness (#1 SpreadsheetML ARGB alpha ignored for fills/fonts, #2/#14 Excel `####` and wrap decisions over one Calibri fallback stack, #4 Excel border styles/colours incl. the lower-clamp fix, #7 scientific/fraction/colour-section number formats and built-in time ids 45–47, #8 gridlines hidden under spilling text, #10 stored row height without `customHeight`, #11 worker-served object documents keep fills/borders) | `32a845c` | opencode APPROVE with 2 NITs (`findings-vis-xlsx-opencode.md`) |
| Stream B follow-up (`b66c7ea` golden re-records; NIT-2 pin of native-1904 B4's exact `#` count in the branded-Chrome spec) | `2c02e37` (`6161e0d`) | opencode NIT-2 confirmed, second-look APPROVE |
| Mobile text-sort refusal-by-design (harness records `environment.scenario.refusals`) | `541f212` (`9b5840b`) | opencode APPROVE (`findings-mobile-cap-opencode.md`) |

Of the 25 audit findings, 23 are addressed by the merged work above. The mapping is the implementers' per-commit claims plus the reviewers' reports, and "addressed" includes findings whose disposition is a verified no-fix: #17's picture grey frames are authored in the fixture (the golden's omission is a reference limitation, retained intentionally), findings with "sample/reference" causes were resolved as authored or reference behaviour, and PPTX a:highlight warping/rotated highlighting remains explicitly unsupported (disclosed in `1cbf514`). Findings #3 (worksheet charts) and #5 (authored frozen panes) are delivered by stream C, which is in review and not merged.

**Spec amendments in PR #11.** `docs/superpowers/specs/2026-10-10-interactive-viewer.md`:
- **§5.5 worker-shell flag-off + one-shot routes** — merged (`e8d9ca6`), in force. A worker-backed (paged) shell stays in sheet mode if the interactive flag later turns off (the legacy renderer has no resident rows); one-shot sources keep the route they loaded with, bytes/Blob/URL sources re-open under the new route.
- **§1.2 authored frozen panes** and **chart data rules** (finding #3/#5, user decision 4) — **pending**, on `feature/iv-vis-features` (stream C, in review, not merged). Authored `<pane state="frozen|frozenSplit">` pins leading rows/columns in sheet mode; `state="split"` is not a freeze; one logical scroll origin; charts draw read-only from series caches or their single-area reference after evaluation, numeric text follows Excel not `Number()`, labels cap at 255 characters, unsupported/unreadable charts show a placeholder card.
- **§5.6 view-only row/column resize** (+ its §1.2 amendment; user decision 5) — **pending**, on `feature/iv-vis-resize` (stream E, codex-APPROVED at `a53045f`, unmerged; final frozen/resize composition is a separate check still open). Session-only sparse overlay on the axis metadata, never saved to the file; rows keyed by source row under sort/filter; pointer drag with 2 px slop, coarse 44 px handles off the worksheet corner; bounded auto-fit; announcement through the status line. Admission order per §5.6 (as amended at `a53045f`): an input-derived upper bound of what building would allocate is **reserved against the live ledger before anything is built**; the unpublished candidate is then built **under that reservation**, off to the side; its exact charge is admitted while the current entry is still held; and the provisional reservation is released before the exact (and DOM) charge is taken, so the ledger never holds provisional and replacement charges at the same time. "Never holds both" refers to those two charge kinds during handoff, not old versus new geometry.
- **§12.4 repaint admission** (per-record layout/canvas/shaping admission amendment, part of the §5.6 stream's text) — **pending**, same branch as §5.6.

**In review, not merged (do not count as delivered):** stream C (`feature/iv-vis-features`; latest verdict codex confirmation r11 at `04f6b10`: **REQUEST CHANGES**, with the accepted-lease lifetime gap across close (R1 MAJOR, reproduced also on base) and the per-admission metadata rescan (R2 MINOR) outstanding; earlier rounds' MAJOR fixes have landed), stream E (`feature/iv-vis-resize`; codex APPROVE at `a53045f` after r9 resolved all four prior findings, **unmerged**; baseline gate limitations preserved in the reviewer's report, final frozen/resize composition a separate open check), stream F (`feature/iv-vis-mobile`: the viewer presents with the admitted device profile instead of fixed D, and header glyphs use text presentation rather than iOS emoji tiles; no stream review yet). **The iOS-Simulator memory overshoots listed in the performance doc (main-owned 36.17–41.48 MiB vs the 32 MiB M cap; bitmap-retained 112.96–113.79 MiB vs the 64 MiB M cap) are recorded observations whose cause is not established: F's presentation-profile fix supports a bitmap-admission hypothesis, but the source base already sized the worker-backed main budget from `remote?.profile`, so it does not explain the transient import-transfer main-owned peak. Both overshoots are under investigation in stream F; no before/after F measurements exist, and the overshoots stand as current merged behaviour.**

## Blocking for release

These are mandatory cases whose evidence is incomplete. Nothing here is hidden as a pass.

1. ~~**AC-STORE-03 / AC-SEM-04 drawing legs / AC-SEM-05**~~ — resolved by S3.3 (merged 2026-10-12); only their browser legs remain not-run.
2. **AC-A11Y-01 / AC-A11Y-02:** S10.3 delivered the Chromium/WebKit keyboard and focus receipts and the Chromium AX-tree receipt. Still open: the **macOS VoiceOver** receipt (S13: guided manual passes in Safari: V1/V2 fixed and re-heard on 2026-10-15, and V5 (a trailing “blank” after each cell) found and fixed, then re-heard at 0c65bff; the automated `scripts/harness/voiceover-check.ts` is blocked by the `content of last phrase` AppleScript hang), ~~Safari desktop~~ (S13: 7/7 in real Safari 26.5.2), and **iPad touch/VoiceOver** (not-run by user decision).
3. **AC-PERF-01 / AC-PERF-02:** S10.2 receipt r16 admits desktop Chrome (13/13 rows; main memory under the round-5 gate plus the user-approved 0.25 MiB engine-cache allowance, 2026-10-13). r8 was admitted under the round-3 rules. The user’s 2026-10-13 allowance decision, with the untraced remainder disclosed, satisfies the prior clause “attribute or charge that remainder (or an explicit decision on it)” through its explicit-decision option; no further remainder decision is pending. Final-tree desktop re-run on 2026-10-16 (`9cfb090` = integration `fb0ddc7`): ADMITTED, 13/13 rows. Still not run: actual Safari performance and the iPad/mobile receipts (not run, by user decision).
4. ~~**Native Excel checkpoint** for AC-STORE-01, AC-SEM-07 and AC-REG-01~~: passed on three Excel 16 (macOS) workbooks (S13, see above): `native-corner`, `native-1904` and the cap-crossing `native-large`, which opens worker-backed without any override. Still open: Excel for Windows.
5. **AC-BACKEND-01:** S11/S12 decision records delivered (`docs/architecture/xlsx-wasm-decision.md`, `xlsx-gpu-decision.md`, both not-run on desktop Chrome (headless)-only evidence); revisit when Safari and mobile receipts exist.
6. ~~**Full-scale browser runs** for F2 (1M × 8), F3 dense (256 × 16,384) and F4 (1M strings)~~: S10.2 ran all three at full scale in Chrome (performance receipt). The S10.1 functional browser suite still uses reduced, labelled scales.
7. ~~**D-DPR (S13, defect)**~~ — fixed on `feature/iv-s13-manual` (`useDevicePixelRatio`); was: a runtime device-pixel-ratio change with an unchanged CSS size (a window moved between a DPR 1 and a DPR 2 display) does not re-rasterize the viewer. The old backing store stays on screen, scaled, until some unrelated input re-renders it. Found with emulation (CDP) in installed Chrome; see [S13](#s13-runtime-and-manual-checks).

## Observations (not fixed; not reproduced as failures)

- (S13) A 1 px pane-width nudge at DPR 1.5 once seemed to leave the pane width stale (canvas CSS 1397 at `innerWidth` 1440). It was investigated and did not reproduce in 192 cycles; every stale state seen was D-DPR, now fixed. Details in the S13 section.
- G2's read-ahead costs one extra `next()` per complete cursor-less page. On the worker path that is one more worker round trip; it is part of the roughly 597 reads observed above. Detection keeps at most one read-ahead continuation open, because the worker caps open scans at 4 (`MAX_OPEN_SCANS`), shared with find. This is a cost note for the S8.6 owner, not a bug.
- Generated browser packages are cached under `$TMPDIR/officeview-viewer-e2e-v1`, keyed by file name. If the S0.2 builders change, bump the `-v1` suffix in `tests/browser/viewer-e2e.ts`.
- `fromStoreScan` (`src/xlsx/view/structure.ts`) has the same first-page-only pattern as G2. It is used only by tests today. Any future production caller must follow iteration-continued pages.
- On worker-backed F2, the saved-values status reads "0 formulas not recalculated (file too large)". The count follows the documented S1.3 definition (formulas the engine would have evaluated at load), and a cached `SUM` is not one of them. The wording is correct but may confuse users.
- `queryViewport` documents itself as not yet wired to a worker `viewport` command (a TODO in `viewport.ts`). The worker-backed viewer paints from bounded `read rows` windows instead, and those windows are what the complexity audit measured.
