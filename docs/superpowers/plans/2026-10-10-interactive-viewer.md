# Full-capacity interactive spreadsheet viewer implementation plan

> **For agentic workers:** Use `superpowers:subagent-driven-development` when execution is delegated, or `superpowers:executing-plans` for sequential execution. Deliver the work packages below in dependency order. Checkboxes record implementation work; writing this plan does not complete any checkbox.

**Date:** 2026-10-10. **Status:** implementation-ready plan, pending specification/plan review; no runtime implementation or measured SLA claim is implied.

**Specification:** [`../specs/2026-10-10-interactive-viewer.md`](../specs/2026-10-10-interactive-viewer.md). The specification owns behavior, limits and acceptance IDs; this plan owns decomposition, paths, test counterexamples and delivery gates. Resolve contradictions in both documents before implementing the affected package.

**Goal:** Provide immutable spreadsheet viewing, navigation, selection, sort, filter, search and copy over the complete Excel coordinate domain of 1,048,576 rows × 16,384 columns, with bounded loading, memory and visible rendering work on desktop and mobile.

**Architecture:** An asynchronous `SourceStore` owns sparse/paged source values and metadata. A resident worker owns workbook sessions, compact view projections and cancellable queries; an identity projection allocates no per-row or per-cell object. The viewer maps complete logical axes into a bounded physical scroll window and paints only visible canonical layout records, with Canvas2D as the behavioral reference and optional interchangeable Wasm/GPU backends admitted only after measured comparisons.

**Tech stack:** Existing TypeScript, React, Vite, Bun, Vitest/jsdom, Canvas2D, formula/format/date engine and font/image leases. Browser performance/interaction verification requires a new, explicit dev-only Playwright harness. Streaming ZIP support is a separately selected implementation dependency, not an assumption that current JSZip/full XML trees satisfy the large mode contract.

## 0. Delivery rules and observed starting point

The complete domain contains **17,179,869,184 logical cells (2^34)**. Capacity is a coordinate and query contract; it is not a claim that that many populated values can fit RAM or open instantly. `capacity`, declared dimension, populated/used extent and viewport extent are separate fields. Small and large source adapters have identical values and structural semantics; large mode must not depend on first building the current object workbook.

At the reviewed checkout, `src/xlsx/render.ts` caps metric arrays at 16,384 rows and 4,096 columns and falls back to Y=0 for higher row indices. `renderSheet` has an optional viewport, but `OfficeDoc` does not provide it, and the renderer still scans source rows and expands every merged cell into a map. `src/worker/office-worker.ts` reparses source for each request. `src/core/zip.ts` uses JSZip and caches full part strings/XML trees. Those are required foundation changes, not deferred renderer work.

- Complete S0–S5 before connecting interactive controls. S6–S9 consume these contracts, not alternative ad hoc maps.
- S10, including physical-device controlled performance results, is required before claiming the feature ready. S11/S12 define bounded optional experiments and require an explicit status record (not-run, not-admitted or admitted); a prototype is not mandatory. An unsuccessful optional backend does not block a compliant TypeScript/Canvas2D release, but a successful prototype does not waive S10.
- Parsed source cells, formulas, styles, table metadata and source order stay immutable. The view owns reorder/visibility and source-identity selection only.
- A red test must fail for the intended behavior, not only a missing import. When a new module is absent, first create its typed boundary with a deliberately incorrect minimal implementation, then record the substantive assertion failure. Remove all deliberate defects before GREEN.
- Checkbox granularity: a RED, run, GREEN, verify or deliver checkbox is one reviewable work item. A GREEN item that bundles several concerns is written as an indented sub-checklist, one concern per sub-item, and the delegating agent dispatches sub-items, never the whole paragraph. The largest packages (S1.1, S1.2, S3.2, S5.1) are pre-split below; before dispatching any other multi-concern GREEN item, split it the same way. Work-package estimates include integration, review and test iteration; a multi-day architecture package must not be presented as a two-minute step.
- No direct commits to `main`. Implement each package on a feature branch/isolated worktree, commit only its owned paths and review before integrating. Preserve unrelated work. This document update does not authorize code changes, dependency installations or commits in the current turn.
- A reviewer independently reproduces the named counterexample after a fix. Passing unit summaries alone do not establish pipeline correctness, browser performance or native parity.

### 0.1 Standard commands and evidence

Run focused suites first. After they pass, every independently deliverable PR runs these existing repository commands from the repository root:

```sh
bunx tsc --noEmit
bun run test
bun run build
bun scripts/corpus-report.ts
bun scripts/golden-corpus.ts compare --threshold 0 --max-ratio 0
```

Expected: all commands exit 0; no newly failed/empty/degraded corpus cases; exact corpus comparison passes for unchanged output. Existing tracked, known degradations must be recorded with baseline evidence, never silently relabeled. For an intentional authored-hidden correction, attach a named fixture list, source evidence and before/after PNGs; update only reviewed affected goldens using `bun scripts/golden-corpus.ts record --only <fixture-name>` and rerun exact comparison. Do not lower thresholds or regenerate the whole corpus to hide a regression. Node-canvas goldens do not prove browser font or GPU parity.

Component verification extends existing jsdom conventions in `tests/components.test.tsx` and `tests/officeview-*.test.tsx`. Numeric timing assertions do not belong in Vitest: performance gates live in the controlled real-browser harness. Test commands named below refer to planned new files and become runnable only in their package. Record failing and passing logs in the PR, not committed generated artifacts.

### 0.2 Interface freeze and file ownership

| Contract / responsibility | Proposed focused files | Existing integration files | Owner |
|---|---|---|---|
| Domain/address checks and typed boundaries | `src/xlsx/view/coordinates.ts`, `src/xlsx/view/contracts.ts` | `src/xlsx/types.ts` | S0 foundation |
| Async source adapters/value authority | `src/xlsx/view/source.ts`, `src/xlsx/view/values.ts` | `src/xlsx/parse.ts`, `src/xlsx/strings.ts`, `src/xlsx/formula/workbook.ts` | S1 source |
| Bounded import/pages/quota/external runs | `src/xlsx/storage/importSession.ts`, `sourcePages.ts`, `overviewIndex.ts`, `quota.ts`, `externalRuns.ts`; `src/core/zip-stream.ts`, `xml-stream.ts` | `src/core/zip.ts`, `xml.ts`, `stream.ts`, `src/components/OfficeFile.tsx` | S1 source |
| Order, visibility, axes, scroll mapping | `src/xlsx/view/order.ts`, `visibility.ts`, `axes.ts`, `projection.ts`, `scrollWindow.ts` | `src/xlsx/render.ts` | S2 projection |
| Range queries and operation safety | `src/xlsx/view/structure.ts`, `operationGuard.ts` | `src/xlsx/drawing.ts`, `parse.ts`, `src/xlsx/formula/spills.ts`, `saved-spills.ts` | S3 structure |
| Worker protocol/session/kernel | `src/worker/viewerProtocol.ts`, `viewerSession.ts`, `viewerKernel.ts`, `viewerClient.ts` | `src/worker/office-worker.ts`, `worker-entry.ts` | S4 session |
| Shared layout/find/base tiles | `src/xlsx/view/textLayout.ts`, `find.ts`, `tileCache.ts`, `viewport.ts` | `src/xlsx/render.ts`, `src/render/paint.ts`, `src/core/search.ts`, `text-recording.ts`, `text-metrics.ts`, `images.ts`, `fonts/*` | S5 render |
| Grid selection, virtual semantics/headers | `src/xlsx/view/selection.ts`; `src/components/xlsx/GridSemantics.tsx`, `GridHeaders.tsx`, `SheetNavigator.tsx`, `ViewerStatus.tsx` | `src/core/overlay.ts`, `src/components/OfficeDoc.tsx` | S6 UI shell |
| Sort/filter/value index/date/copy | `src/xlsx/view/sort.ts`, `rangeDetection.ts`, `filter.ts`, `valueIndex.ts`, `datePredicates.ts`, `copy.ts` | `src/xlsx/formula/serial.ts`, `format.ts`, `src/xlsx/formula/tables.ts`, `src/core/clipboard.ts` | S7–S9 engines |
| Popups/dialog UI | `src/components/xlsx/FilterDropdown.tsx`, `SortDialog.tsx` | `src/components/OfficeDoc.tsx` | S9 controls |
| Fixtures/browser/performance | `tests/helpers/xlsx-viewer-fixtures.ts`, `tests/fixtures/xlsx-viewer/*`, `tests/browser/interactive-viewer.spec.ts`, `scripts/bench-interactive-viewer.ts`, `playwright.viewer.config.ts`, new `src/debug/*` (directory does not exist today; S0.3 creates it and must keep it out of the library build) | `package.json`, `bun.lock`, `vite.config.ts` (dev page input only) | S0/S10 verification |
| Sheet scale/zoom model (viewport mode) | `src/xlsx/view/sheetZoom.ts` | `src/core/zoom.ts` (read-only reuse of `clampZoom` with explicit bounds; `MIN_ZOOM`/`MAX_ZOOM` unchanged), `src/components/OfficeDoc.tsx` | S2.4 |
| Optional Wasm kernel / WebGPU reference comparison | `src/xlsx/backends/*`, `experiments/xlsx-wasm/*`, `experiments/xlsx-webgpu/*` | Vite worker/backend build seams only | S11/S12 experiment owners |

`SourceStore` defines asynchronous `readCell`, `readRows`, `queryColumn`, `queryStructural`, `queryOverview`, `scanPopulated`, `ensureComplete` and `release`; names and signatures are frozen in S0 against the specification. Small source is `ObjectXlsxSheetAdapter`; large source is `PagedXlsxSheetAdapter`. `RowOrder` is lazy identity or a band `Uint32Array` permutation plus inverse. `RowVisibility` uses compact masks and rank/select. `AxisMetrics` uses default dimensions, sparse interval overrides and `Float64` prefixes/segments. `ScrollWindow` maps logical offsets to a physical surface of at most 1,000,000 CSS pixels per axis. `CacheIdentity` includes `docId`, `sheetId`, `valueGen`, `fontGen`, `viewGen`; shared workbook publication additionally uses `WorkbookOwner.globalViewGen` and a complete ordered sheet revision vector.

The integration owner serializes changes to `OfficeDoc.tsx`, `xlsx/render.ts`, `xlsx/parse.ts`, `xlsx/types.ts`, `worker-entry.ts`, `render/paint.ts`, `core/search.ts`, `core/overlay.ts`, `package.json` and `bun.lock`. Other owners land focused modules first and supply a small reviewed integration patch; two agents never edit those shared files concurrently. Before accepting that patch, rebase onto the current integration branch, rerun its counterexample and inspect the staged diff. Module boundaries must not break DOCX/PPTX rendering, search, selection, zoom, protection or worker exports.

### 0.3 Dependency and delivery map

| Group | Independent deliverable | Prerequisites | Estimate, engineering days including review | Required for release |
|---|---|---|---:|---|
| S0 | frozen contracts, fixtures, real-browser harness scaffold | specification agreement | 2–4 | yes |
| S1 | bounded import, semantic source adapters and the `cached`/`engine` value policy | S0 | 12–22 (re-estimated: streaming ZIP reader selection, XML tokenizer, OPFS/IndexedDB pages, overview hierarchy, external runs, two adapters, value policy) | yes |
| S2 | full axes/order/visibility/scroll mapping and sheet scale model | S0; adapter contract, not importer completion | 5–8 | yes |
| S3 | indexed structural metadata and safe operation queries | S0, S1 metadata; S2 impact math | 3–5 | yes |
| S4 | resident worker, cancellation, resource lifetime | S1, S2, S3 | 4–7 | yes |
| S5 | bounded viewport/canonical records/global find | S1–S4 | 6–10 | yes |
| S6 | navigator, source selection, headers/a11y shell | S3–S5 | 4–7 | yes |
| S7 | stable whole-row sorting and range detection | S1–S4; S6 only for controls | 3–5 | yes |
| S8 | typed filtering and complete value universe | S1–S4; S7 parity fixtures | 4–7 | yes |
| S9 | popup/dialog/copy/mobile interaction | S6–S8 | 4–6 | yes |
| S10 | end-to-end hardening, physical-device performance gate | S0–S9 | 5–10 | yes |
| S11 | bounded Rust/Wasm comparison and decision | S4, S7, S8; S10 measurements | 3–5 | status decision yes; prototype/adoption optional |
| S12 | bounded WebGPU comparison and decision | S5; S10 measurements | 4–7 | status decision yes; prototype/adoption optional |
| S13 | independent release review/docs/rollout checklist | all mandatory groups and backend decisions | 2–3 | yes |

S1 and S2 focused modules may proceed concurrently after S0; shared integrations remain serial. S7 and S8 may proceed concurrently after S4 using synthetic adapters; neither may bypass S3. S11 and S12 are independent experiments and must not delay a compliant baseline with an unbounded rewrite. Estimates are planning ranges, not schedule promises.

### 0.4 Progress log (updated when a package merges)

Every package is implemented on its own branch, reviewed by at least one agent from a different model family than its implementer, and merged into `feature/interactive-viewer` only after every finding is fixed. Checkboxes below in each section are ticked only for merged work.

| Package | State (2026-10-10) |
|---|---|
| S0.1 contracts, S0.2 fixtures, S0.3 harness scaffold | merged |
| S2.1 order/visibility/axes, S2.2 scroll window, S2.4 sheet zoom model (UI wiring open) | merged |
| S8.1 typed filters, date lowering (`filter.ts`, `datePredicates.ts`, `textFilter.ts`; worker/mask publication open) | merged; its collator text matching is replaced by S8.5 canonical matching |
| S8.5 canonical text matching (`canonicalText.ts`, pinned Unicode 16.0 tables under `src/xlsx/view/unicode/`, `textFilter.ts`) | merged (codex reviews over two rounds; Unicode 16.0 pinned data incl. combining classes; key strategy `canonical-text/v2`) |
| S7.2 pure sort engine (`src/xlsx/view/sort.ts`; worker kernel/session/header integration open) | merged |
| S1.1 streaming ZIP/XML | design note merged (`docs/architecture/xlsx-large-source.md`, reviewed by codex, pi and opencode over five rounds; user approved `pako`); implementation split into S1.1a (ZIP/inflate/XML tokenizer core) and S1.1b (spool, admission gates, entry wiring) |
| S3.1/S3.2 structural index and operation guard (pure parts: `structure.ts`, `operationGuard.ts`; parser/drawing/spill wiring open) | merged |
| S4.1 resident worker protocol/session (`src/worker/viewer*.ts`; temporary object-backed source until S1.2) | merged |
| S7.1 range detection/header policy/admission (`src/xlsx/view/rangeDetection.ts`) | merged |
| S4.2 cooperative scheduling/cancellation/atomic publication (`src/xlsx/view/scheduler.ts`, worker session/kernel/client) | merged |
| S1.1a bounded ZIP reader, pull inflater, incremental XML tokenizer (`src/core/{zip-stream,inflater,xml-stream}.ts`, `pako` 1.0.11) | merged (opencode review; one disputed finding accepted by the reviewer) |
| S2.3 full-coordinate renderer seam and explicit static budgets (`src/xlsx/view/viewport.ts`, `render.ts`, `paint.ts`) | merged |
| S8.3 declared filter diagnostics (`src/xlsx/view/declaredFilters.ts`, parse.ts) | merged |
| S8.2 semantic value index and searched Select All (`valueIndex.ts`, `valueRuns.ts`, `searchKeys.ts`) | merged |
| S4.3 lifecycle, document-wide memory pool, LRU tiers, tile cache and leases (`tileCache.ts`, worker session/kernel) | merged |
| S1.1b spool, admission gates, quota ledger, entry wiring (`spool.ts`, `storage/{quota,admission,importSession,sheetRecords,imageAdmission}.ts`, `largeSource.ts` lazy chunk) | merged; until S1.2 every 'paged' decision opens in compatibility mode with a notice |
| S6.1 selection state machine, navigator and status (`selection.ts`, `SheetNavigator.tsx`, `ViewerStatus.tsx`) | merged |
| S5.1 canonical text/layout records and viewport queries (`textLayout.ts`, `viewport.ts`, render layout/paint split) | merged |
| S6.2 virtual owned grid semantics, headers, event isolation (`GridSemantics.tsx`, `GridHeaders.tsx`, OfficeDoc guards) | merged; Ctrl/Cmd+wheel over any viewer now actually blocks browser page zoom (behaviour change, flagged) |
| S5.3 global workbook find engine (`src/xlsx/view/find.ts`) | merged; OfficeDoc wiring and worker search command wait for S5.2 |
| S1.2 sparse pages, external page storage, small/large adapter parity (`sourcePages.ts`, `pagedImport.ts`, `pagedSource.ts`, `overviewIndex.ts`, `viewerLargeSource.ts`) | merged; worker routes admitted-paged blobs to paged sources; exported entry keeps compatibility mode until the viewer is enabled |
| S9.1 atomic filter dropdown and sort dialog (`FilterDropdown.tsx`, `SortDialog.tsx`, `controlsClient.ts`, header sort/funnel buttons) | merged; mounted by S5.2 |
| S1.3 authoritative formula/cache/error/format value bridge (`values.ts`, `value-policy.ts`) | merged |
| S5.2 | in progress (behind `interactiveSheets` flag) | (implementation or review rounds on their own branches) |
| S1.2, S1.3, S4.4–S4.5, S5, S6, S9, S10 (physical-device receipts skipped for now by user decision), S11/S12 status, S13 | not started |

## S0 — Freeze contracts, counterexamples and fixture/harness infrastructure

### S0.1 Coordinate/source/view contracts

> **Status (2026-10-10): done.** Merged into `feature/interactive-viewer` (codex implementation; three review rounds by Claude and opencode; every finding fixed). Review follow-ups recorded below in S1.2 and S4.1.

**Files:** create `src/xlsx/view/coordinates.ts`, `contracts.ts`; test `tests/xlsx-viewer-contracts.test.ts`; narrowly extend `src/xlsx/types.ts` only for additive metadata/source descriptors.

- [x] RED: add boundaries for A1, XFD1048576, row 16385, XFE1, row 1048577, negative/fractional/NaN/Infinity coordinates, malicious dimensions and a declared full extent with one populated cell. Pin zero-based row/column identity and maximum cell ID. Include row boundaries 16,384/16,385/131,072/262,144/1,048,575 and column boundaries 4,095/4,096/16,383: row 131,072 crosses signed 2^31 and row 262,144 crosses 2^32 under forbidden packing.
- [x] Run `bunx vitest run tests/xlsx-viewer-contracts.test.ts`; expected assertion failure for an invalid accepted coordinate or aliased ID, not only missing symbols.
- [x] GREEN: implement finite safe-integer domain checks, address parsing/formatting and `cellId = row * 16384 + col`. Do not use 32-bit shifts, `row << 14`, bitwise `|0`, or pack the total ID into Uint32. Carry row in Uint32 and col in Uint16; maximum ID is 17,179,869,183 and remains a safe JS Number.
- [x] Freeze `SourceStore`, `OverviewQuery`/`OverviewDomain`/`OverviewGroup`, read status (`known value`, `known blank`, `pending/unknown`, `error`), `RowOrder`, `RowVisibility`, `AxisMetrics`, `ScrollWindow`, structural results, cancellable requests and cache/session IDs. Define pending as distinct from blank, including sort/filter/count eligibility. `queryOverview` is mandatory and source-space only: domains carry groupId, exact `CompactSourceRowSet` and column intervals; raw responses carry source-owned count/countKind/style/content and no viewport bounds. Freeze maxGroups≤2048 desktop/≤1024 mobile, bounded resident membership handles, cancellation/cursor/source-stamp validity and terminal complete/resource/error outcomes. Never replace disjoint rows with their enclosing rectangle, sample text as complete, or keep a complete source indefinitely loading.
- [x] Run the focused command and `bunx tsc --noEmit`; expected all contract assertions pass and additive types compile without changing existing format consumers.
- [x] Deliver a reviewed contract PR with versioned DTO examples and paths; no returning-identity sort/filter stubs are considered completed engines.

Meaningful boundary assertion to preserve:

```ts
expect(cellId(1_048_575, 16_383)).toBe(17_179_869_183)
expect(cellId(262_144, 0)).not.toBe(cellId(0, 0))
expect(() => validateCoordinate(1_048_576, 0)).toThrow()
expect(await source.readCell({ row: 1_048_575, col: 16_383 }, readOptions)).toMatchObject({ kind: 'cell' })
```

### S0.2 Deterministic fixture family

> **Status (2026-10-10): done.** Merged (opencode implementation; three review rounds by Claude; every finding fixed). The `it.todo` for `source.readCell` in the contracts test is satisfied in `tests/xlsx-viewer-fixtures.test.ts`.

**Files:** create `tests/helpers/xlsx-viewer-fixtures.ts`, `tests/fixtures/xlsx-viewer/manifest.json`; generated small native packages under `tests/fixtures/xlsx-viewer/`; test `tests/xlsx-viewer-fixtures.test.ts`.

- [x] RED: assert builder output domain, populated count, key checksums, formula/error/style metadata and repeatability. A full-capacity synthetic provider must report every cell readable while generating values by coordinates and materializing only requested pages.
- [x] Run `bunx vitest run tests/xlsx-viewer-fixtures.test.ts`; expected fixture metadata/checksum assertions fail on deliberately incomplete fixture builders.
- [x] GREEN: implement these deterministic families (names per specification §12.5): F1 corner (full extent, 100k NNZ, edge XFD1048576 and legacy-cap row), F2 tall (1,048,576×8 numeric/mixed columns), F3 wide (dense 256×16,384 true XLSX = 4,194,304 populated records; separate sparse-wide variant), F4 string domain (1M unique natural-sort strings, accented/case/lone-surrogate examples), F5 structural (giant merge/style intervals/spill child/drawing below band), F6 full-axis synthetic (synthetic fully populated 1,048,576×16,384 provider, geometry/window queries only), F7 view stress (many max-dimension sparse sheets, equal local revisions, eight-key sorts, slow jobs, variable/tiny dimensions/zoom/DPR, fonts/assets), F8 compatibility (cache/error/date/table/Unicode parity).
- [x] Keep committed binary fixtures small. Generate tall/text/wide packages from deterministic seeds into a declared temporary output directory. Record encoded/uncompressed bytes, NNZ, row/column bounds, seed, source hash and expected query hashes in the manifest. Do not allocate `rows × cols` during generation.
- [x] Run the focused suite; expected no giant in-memory workbook, F6 only requests bounded windows, generated/package and adapter queries agree.
- [x] Deliver builders reusable by units, browser scripts, native comparison and both backend experiments; document native fixture provenance and permission-free local generation. F2 and F3 benchmarks require generated true XLSX input, not only an arithmetic provider. Fixture IDs F1–F8 are the single canonical names from the specification catalog (§12.5); the manifest, harness flags and tests use only these IDs.

### S0.3 Real-browser harness and performance result schema

> **Status (2026-10-10): done as a scaffold.** Merged (pi/Claude/opencode implementation, Claude reviews over three rounds, all findings fixed). Corner-navigation remains `test.fixme` until S5; `--mode release` is refused until real fixtures flow through the viewer (`REAL_FIXTURES_AVAILABLE` in `scripts/harness/args.ts`). No viewer SLA is claimed.

**Files:** create `playwright.viewer.config.ts`, `tests/browser/interactive-viewer.spec.ts`, `src/debug/interactive-viewer.html`, `src/debug/interactive-viewer.tsx`, `scripts/bench-interactive-viewer.ts`; modify `package.json` and `bun.lock` in the execution PR.

- [x] RED (corner-navigation scenario is `test.fixme` until S5; the harness smoke itself is green): harness smoke scenario opens a source fixture and asks for corner navigation, keyboard input acknowledgement and a structured result; expected missing interactive path or result fields fail. Unit harness tests may check accounting, never timing SLA values.
- [x] Add explicit dev dependency through `bun add -d @playwright/test` in this future work package; record the resolved version in the lockfile. Install only required engines with `bunx playwright install chromium webkit`. Safari and iPad results still require actual Safari/device runs; Playwright WebKit is not relabeled Safari.
- [x] GREEN: implement fixture loading, deterministic clock, input markers, frame/present markers, long-task collection, application allocation ledger and JSON/CSV artifacts. Browser-produced event/layout metrics remain separate from Bun/model microbenchmarks.
- [x] Implement CLI flags for `--browser chromium|webkit|safari-manual`, `--profile desktop|mobile`, `--fixture`, `--repeat`, `--output`, `--mode smoke|release`, `--baseline`, and manual-device server flags `--serve --port 4173`. Manual-device mode serves the same deterministic harness, emits a LAN/local URL, and ingests/export-verifies device result JSON without relabeling emulation as physical evidence. Record hardware, OS, browser/build, DPR, viewport, font resolution, cache/cold state, source bytes and backend.
- [x] Run `bunx playwright test --config playwright.viewer.config.ts --grep 'harness smoke'` and `bun scripts/bench-interactive-viewer.ts --browser chromium --profile desktop --fixture F1 --repeat 3 --mode smoke --output /tmp/officeview-viewer-smoke`; expected harness works and emits complete measured fields, with unsupported metrics labeled unavailable rather than zero.
- [x] Deliver a runnable harness scaffold and device-run instructions. Do not claim any full viewer SLA passed from this scaffold.

## S1 — Mandatory bounded import, source access and value fidelity

### S1.1 Package/XML streaming and admission

> **Status (2026-10-11):** S1.1a (ZIP/inflate/XML core) and S1.1b (spool, admission G1/G2, quota ledger, evaluator materialization cap, entry wiring) merged after codex/pi/opencode reviews. Orchestrator rule applied: anything that opened before still opens — paged decisions use the legacy path with a 'compatibility mode' notice until S1.2; evaluator cap hits reopen with saved values; damaged-but-readable archives open with a 'may be damaged' notice; typed refusals only for produced≠declared size, truncated deflate and spool storage/quota. Eager bundle +2.7 KB gz; large path in a lazy chunk never loaded for DOCX/PPTX. Not delivered: RED 14a, 14b render wiring, 21, import progress (22), session-level 23 (S1.2).

**Files:** create `src/core/zip-stream.ts`, `xml-stream.ts`, `src/xlsx/storage/importSession.ts`, `quota.ts`; modify `src/core/zip.ts`, `stream.ts`, `src/components/OfficeFile.tsx`; test `tests/xlsx-viewer-import.test.ts`, `tests/xlsx-viewer-import-limits.test.ts`, retain `tests/streaming.test.ts`/`officeview-streaming.test.tsx`.

- [ ] RED: use chunked local Blob/stream inputs whose worksheet/sharedStrings parts exceed the small-object threshold. Assert no whole-file ArrayBuffer or full worksheet/sharedStrings XML string/tree is retained in large mode, bounded chunk memory, progressive first rows after package inventory, actual inflated-byte accounting and cancellation cleanup. Network input is first downloaded/spooled to a seekable Blob/OPFS-backed source under a bounded admission profile because ZIP central directory occurs at the archive end; do not promise OOXML render from the first network byte. Test download/spool pause/resume, backpressure, cancellation and cleanup separately from part parsing.
- [ ] Add ZIP central-directory/part validation counterexamples: bogus declared sizes, inflated limit crossings, duplicate/escaping paths, unsupported compression, truncated tokens and huge attributes/string records. Include a single giant XML token/string/drawing element and a compressed image with huge decoded dimensions. Limits apply while decoding/tokenizing, before retaining excess output; an incomplete import cannot publish a complete blank source.
- [ ] Run `bunx vitest run tests/xlsx-viewer-import.test.ts tests/xlsx-viewer-import-limits.test.ts tests/streaming.test.ts`; expected memory/unknown-state/cleanup assertions fail with the current full-buffer path.
- [ ] Compare current JSZip capabilities against the required bounded random part/chunk reads. Select a maintained streaming inflater/ZIP reader, or implement a bounded local reader using supported primitives. Record exact dependency/version/license/API and a reviewed design note at `docs/architecture/xlsx-large-source.md`; do not assume JSZip full-part `.async('string')` meets the contract. Dependency installation belongs only to this implementation PR.
- [ ] GREEN: bounded streaming import (dispatch each sub-item separately):
  - [ ] Preserve a source Blob or page-backed stream handle; never retain a whole-file `ArrayBuffer`.
  - [ ] Add bounded compressed/uncompressed queues with producer pause/resume/backpressure and admission counters.
  - [ ] Stream worksheet and sharedString records. Initial profile matches spec: 256KiB input chunks, at most 8MiB queued expanded XML, 4MiB target cell pages, 64KiB inline-string soft threshold; oversized values become bounded referenced blob records.
  - [ ] Decode image dimensions/pixel area before full raster allocation and charge decoded buffers to owner budgets.
  - [ ] Bound or page token, rich-run, relationship and drawing metadata.
  - [ ] Keep small documents on the current parser through an explicit threshold and adapter; large documents import source pages directly and never pass through a huge `XlsxDocument.rows` materialization.
- [ ] Wire progress/cancel/protection callbacks with existing stale-owner guards; preserve source HTTP headers/credentials, protection/watermark policy and `allowCopy`/`allowPrint` through both load profiles. A denied/limited resource emits a typed diagnostic and keeps previous loaded document intact where applicable.
- [ ] Run the focused command plus `bunx vitest run tests/officeview-streaming.test.tsx tests/officeview-protection.test.tsx`; expected bounded importer passes and protected/replaced loads cannot publish stale permissive state.
- [ ] Deliver an import-only PR with byte/allocation evidence for F2, first-page readiness vs full-import throughput separately, and explicit browser-storage support/failure behavior.

### S1.2 Sparse pages, external storage and small/large adapter parity

> **Status (2026-10-12):** merged (Claude implementation; opencode reviews over two rounds plus pi notes; every finding fixed). Typed ColumnLanes, object and paged adapters behind one SourceStore with parity tests, OPFS/IndexedDB/memory page backends (fakes in vitest), decoded-page cache in the MemoryPool source tier, one-level overview index, RED 14a/21/22/23. Proposed-until-S10 values: pages pool 1024/512 MiB, `paged.structuralRecords` 1,048,576, quota precheck 6 page bytes per compressed byte. Deferred: `externalRuns.ts` (sorted external runs), `styles.xml` on the paged path (styleIndex only), real-browser storage checks (S10).

**Files:** create `src/xlsx/storage/sourcePages.ts`, `overviewIndex.ts`, `externalRuns.ts`, `src/xlsx/view/source.ts`; test `tests/xlsx-viewer-source.test.ts`, `tests/xlsx-viewer-source-pages.test.ts`, `tests/xlsx-viewer-overview.test.ts`.

- [ ] RED: query far-corner known values, missing records, not-yet-imported pages, whole-column sparse chunks, empty wide bands, hidden/style interval ranges and 1M×8 scan batches. Assert unknown never equals blank and counts never pretend completeness before the queried domain is imported. Add `queryOverview` RED cases: exact noncontiguous source rows {3,1} exclude row2; blank/default analytic spans, mixed style/content, bounded groups, incomplete pages, stale SourceStamp, cancel during hierarchy construction and injected quota/error. A complete source terminates with complete summaries or an explicit typed failure.
- [ ] Run `bunx vitest run tests/xlsx-viewer-source.test.ts tests/xlsx-viewer-source-pages.test.ts tests/xlsx-viewer-overview.test.ts`; expected pending/blank or bounded query failures.
- [ ] GREEN: source adapters and overview queries (dispatch each sub-item separately):
  - [ ] Implement `ObjectXlsxSheetAdapter` and `PagedXlsxSheetAdapter` behind the same async `SourceStore` interface; store populated cells/styles/shared strings by compact page/chunk indices, never objects for blank rows/cells.
  - [ ] Coalesce concurrent page loads and cache under a measured LRU budget; preserve source identity through eviction/reload.
  - [ ] Maintain/build a bounded sparse occupancy/default-style hierarchy during ingestion or in cancellable page batches (hierarchy pages, construction scratch and temporary copies count against aggregate budgets).
  - [ ] Implement source-only `queryOverview` for exact compact source row sets plus column intervals without dense enumeration; raw aggregates use SourceStamp/value ownership, `countKind` distinguishes exact/estimated, style uniform/mixed and content blank/populated/mixed are explicit.
  - [ ] Make the analytical F6 provider return bounded aggregate groups directly.
  - [ ] Guarantee completion, error/quota and cancellation release ownership deterministically.
  - [ ] **Mandatory before S7/S8 consume `queryColumn`:** add the typed lanes variant to `ColumnEntry`/`queryColumn` (row ids `Uint32Array`, type tags `Uint8Array`, numbers `Float64Array`, strings by offset/index preserving raw UTF-16, plus a default-blank run) next to the S0.1 extension-point note in `contracts.ts`; a late union member would break exhaustive switches written earlier.
- [ ] Add OPFS/IndexedDB-backed pages and sorted external runs with quota admission, transaction completion, cancellation, cleanup on close and deterministic memory fallback for small admitted documents. If storage is unavailable/denied or quota is insufficient, refuse the required large operation or open mode with an actionable diagnostic; do not silently truncate, drop columns or fallback to unbounded RAM.
- [ ] Limit dense query preparation to at most eight requested columns in a batch. Outsiderange checks use nonempty chunk/row-column indexes. A sparse full-shape query scales with logical axes plus queried NNZ; F6 viewport geometry performs no populated scan of its 17B cells.
- [ ] Run focused suites with injected quota denial/storage failure and adapter parity over identical fixtures, including incomplete structural metadata and blank/nonblank coverage; expected query values, range metadata, metadataCoverage/pending transitions and checksums agree, no full-grid allocations, bounded buffer accounting passes.
- [ ] Deliver source-store PR with eviction/reload and external-run receipts. Explicitly record bytes charged to source pages, temporary runs, worker heap and transfer buffers.

### S1.3 Authoritative formula/cache/error/format bridge

> **Status (2026-10-12):** merged (Claude implementation; opencode review; every finding fixed). One value authority for object and paged adapters; `engine` documents unchanged and byte-identical; `cached` documents never invoke the evaluator (evaluateWorkbookFormulas throws), uncached formulas are `unavailable`; sort and filter refuse `formula-value-unavailable` with count + first address in key/rule domains; saved-values notice counts formulas the engine would have evaluated at load (orchestrator decision); fixed the worker wrapping cached documents as engine. Policy and evaluation-bridge upgrade path in `docs/architecture/xlsx-large-source.md` §9.

**Files:** create `src/xlsx/view/values.ts`; modify `src/xlsx/parse.ts`, `strings.ts`, narrowly `src/xlsx/formula/workbook.ts` only to expose the per-document value-policy flag (`engine` for small, `cached` for paged; no evaluator is wired to paged pages); test `tests/xlsx-viewer-values.test.ts`, `tests/xlsx-viewer-adapter-parity.test.ts`; retain native/formula suites.

- [x] RED: pin numbers/booleans/blanks/error tokens vs literal '#VALUE!' text, published calculated values (`engine` policy), file caches, absent formula caches, unsupported formula cache preservation, saved spill children, date styles and lone UTF-16 surrogates on both adapters, each under one declared value policy.
- [x] RED: value-policy cases (spec §4.4): a `cached` document never invokes the evaluator (spy on the engine); formula cells carry provenance `cached`; a cache-less formula in a key column refuses with `formula-value-unavailable` (count + first address) while the same formula outside the key/value domain does not; `engine` output for small documents is byte-identical to today; the “Showing saved values; formulas are not recalculated” status appears only for `cached` documents; cross-policy differences are enumerated, never asserted equal.
- [x] Run `bunx vitest run tests/xlsx-viewer-values.test.ts tests/xlsx-viewer-adapter-parity.test.ts`; expected a bad raw/formula fallback or altered string fails.
- [x] GREEN: value authority (dispatch each sub-item separately):
  - [x] `engine` documents: keep reading the existing engine's published `cell.value`/cache/error status and formatting/date semantics, unchanged.
  - [x] `cached` documents: store per-cell authority fields (value, provenance, `valueIsError`, `unavailable`) in pages; run no evaluator and add no new formula evaluator or function catalog.
  - [x] Add the typed `formula-value-unavailable` refusal reason and the saved-values status/accessible text.
  - [x] Record the policy decision, the enumerated engine-versus-cache differences and the out-of-scope evaluation-bridge upgrade path in `docs/architecture/xlsx-large-source.md`.
- [x] Preserve lossless UTF-16 in string tables and transport. Reuse existing format/date serial helpers, font/value generations and immutable source versioning.
- [x] Run focused tests plus `bunx vitest run tests/xlsx-native-display.test.ts tests/xlsx-native-strings.test.ts tests/xlsx-native-metadata.test.ts tests/xlsx-formula-calc-policy.test.ts tests/xlsx-formula-spills.test.ts`; expected authoritative engine parity, unchanged string/error distinctions and no source mutation.
- [x] Deliver value bridge and native evidence matrix. Import/value publication completion bumps `valueGen`; view operations never recalculate or rewrite formulas merely because display order changes.

## S2 — Full logical axes and compact projection before any UI

### S2.1 Order, visibility and axis data structures

> **Status (2026-10-10): done.** Merged (codex implementation; reviews by Claude and opencode; every finding fixed). Typed `AxisMetricsError` mapping to diagnosed fallbacks is tracked in S2.3.

**Files:** create `src/xlsx/view/order.ts`, `visibility.ts`, `axes.ts`, `projection.ts`; test `tests/xlsx-viewer-projection.test.ts`, `tests/xlsx-viewer-axes.test.ts`.

- [x] RED: construct full row/column capacity without row/cell objects, row 16385 and XFD widths, hidden row/column intervals, default dimension overrides across whole axes, a 1M-row sorted band and all-body-hidden rank/select. Assert identity outside a sorted band and invert source↔display slot exactly.
- [x] Run `bunx vitest run tests/xlsx-viewer-projection.test.ts tests/xlsx-viewer-axes.test.ts`; expected wrong mapping, overflow or object allocation assertions fail.
- [x] GREEN: implement lazy identity `RowOrder`; allocate a Uint32 permutation/inverse only for a reordered band. Compact bitsets/rank/select govern visibility and pinned rows. Store width/height defaults plus sparse interval overlays and Float64 prefix/segment navigation; no object per blank row/cell and no per-cell merge expansion.
- [x] Define source row, display slot and visible rank distinctly. Worksheet row labels use display slot +1, never visible rank +1; body ARIA row indices use display slot +2 and body column indices use sourceCol +2 because header chrome occupies semantic row1/col1 (S6.2); hit testing uses visible rank→display slot→source. Source dimensions move with reordered source rows. Collapsed rows/cols have zero extent and produce no text/layout records.
- [x] Implement one body-only visibility predicate: filter body excludes protected header/totals; protected/outside rows retain base snapshot visibility; matching authored-hidden body may reveal; range replacement recomputes from base; Clear restores the snapshot exactly.
- [x] Run focused suites; expected full capacity/roundtrip and body/header/totals cases pass, identity construction has no million-object allocation and prefix precision remains stable at the edge. For explicit full-band accounting, row order+inverse≈8MiB, Float64 prefix≈8MiB, Float64 height lane≈8MiB, one visibility mask≈128KiB; all scratch lanes and up-to-eight numeric key lanes are charged, never treated as free.
- [x] Deliver projection-only PR with allocation ledger and unit complexity invariants (query visits counts), no UI wiring.

### S2.2 Bounded scroll window and navigation math

> **Status (2026-10-10): done.** Merged with S2.4 (codex implementation; reviews by opencode and Claude; every finding fixed).

**Files:** create `src/xlsx/view/scrollWindow.ts`; test `tests/xlsx-viewer-scroll-window.test.ts`.

- [x] RED: navigate directly to XFD1048576, recenter across both axis window edges, `SheetZoom` (S2.4)/DPR changes (25%/200%/400% and fractional zoom), hidden zero-height runs, giant custom intervals and offsets near the end. Assert physical scroll surface, including wide default 64px columns and applied zoom, never exceeds 1,000,000 CSS pixels per axis and logical coordinates never jump after recentering.
- [x] Run `bunx vitest run tests/xlsx-viewer-scroll-window.test.ts`; expected unbounded surface or rounding/recenter mismatch fails.
- [x] GREEN: implement Float64 logical offsets plus local physical scroll offset, segment recentering and viewport anchor preservation. Coordinate navigation, keyboard navigation and programmatic find use the same navigator; do not depend on browser maximum scrollHeight/width. UI Go To defaults to a displayed address (row display slot→source via RowOrder); explicit Original cell navigation and source-coordinate API accept SourceCell and map through inverse order. Formula/search refs remain source identities. Add order [0,2,1] with no header: Go To A2 resolves source A3; Original A2 resolves displayed A3. An exact hidden target opens the read-only value inspector with hidden reason without unhiding or creating an invalid grid active descendant; existing nearest-visible/null grid active state remains valid.
- [x] Run focused command; expected exact source hit tests at first/last cells before/after recentering, clamped legal edges, preserved zoom anchor and no allocation proportional to `R×C`.
- [x] Deliver scroll math independently; browser scroll integration is S5/S6 after resident sessions.

### S2.3 Remove cap aliasing from interactive path; explicit static budgets

> **Status (2026-10-11):** merged (Claude implementation; codex, pi and opencode reviews over four rounds; every finding fixed; xlsx goldens unchanged). Interactive geometry uses full logical axes with per-frame admission from bounded metadata before any cell read (record/gridline budgets, four-device-pixel detail rule, bounded drawing queries, indexed merge lookup); `AxisMetricsError` codes map to diagnosed fallbacks. Static path: ordinary sheets inside the legacy 16,384×4,096 grid render exactly as before at any pixel size; only output the old path silently truncated gets a typed `large-export-limit` refusal; grid-truncated merges lay out at their true extent in bounded requests. Coarse layer is a placeholder until S5.1; new API (`renderSheetViewport`, `buildSheetGeometry`, `StaticExportLimitError`) not yet re-exported from `src/index.ts`; `parse.ts` row numbering for rows without `r` remains open.

**Files:** modify `src/xlsx/render.ts`, `src/render/paint.ts`; create `src/xlsx/view/viewport.ts` boundary; test `tests/xlsx-viewer-cap-regression.test.ts`, extend `tests/xlsx-viewport.test.ts`, `tests/xlsx-print.test.ts`.

- [x] RED: current-row regression with source rows 0 and 16385, viewport at origin and at row 16385; assert 'BeyondCap' is absent at row-zero baseline, then visible only at its true offset or its sorted display slot. Repeat edge-column XFD and edge-row XFD1048576.
- [x] Run `bunx vitest run tests/xlsx-viewer-cap-regression.test.ts tests/xlsx-viewport.test.ts`; expected current fallback paints BeyondCap at Y≈16.67 and fails.
- [x] GREEN: interactive metrics use full logical axis/projection methods, not capped arrays; remove every `rowY[row.index] ?? 0`/`colX[col] ?? 0` alias in that path. Viewport reads only admitted logical bounds and fetched rows. If no geometry exists, return pending/diagnostic, never origin.
- [x] Map the typed `AxisMetricsError` codes thrown by S2.1 (contradictory overlapping overrides, unrepresentable Float64 geometry) to a diagnosed fallback in every caller (S2.3 renderer seam, S3 structure/guards); a thrown error must never reach the user as an unhandled exception or blank sheet.
- [x] Preserve small ordinary-sheet layout behavior; honor authored-hidden rows/cols unconditionally rather than drawing-gated. Keep static print/export separate with explicit pixel/page/operation budgets and diagnostic/refusal; it must not quietly clamp logical coordinates or emit a deceptively complete partial bitmap.
- [x] Run focused suites plus `bunx vitest run tests/xlsx.test.ts tests/xlsx-drawings.test.ts tests/xlsx-print.test.ts` and standard corpus gates; expected no-origin-alias passes, small default goldens unchanged, authored-hidden deltas named and reviewed.
- [x] Deliver full-coordinate renderer seam. This package is not complete if only model sorting bypasses `MAX_GRID_*` while painting still aliases high coordinates.

### S2.4 Sheet scale/zoom model and legacy-path isolation

> **Status (2026-10-10): model done, UI wiring open.** The pure sheet-zoom model, retained per-sheet store and `ScrollWindow` integration are merged; the `OfficeDoc.tsx` mode switch behind the flag, the §11.1 detail-admission wiring (S5.1) and the corpus gates for the flag-off path remain unchecked below.

**Files:** create `src/xlsx/view/sheetZoom.ts`; integrate `src/components/OfficeDoc.tsx` behind the interactive feature flag; reuse `src/core/zoom.ts` read-only; test `tests/xlsx-viewer-sheet-zoom.test.ts`, extend `tests/officeview-zoom.test.tsx`.

- [x] RED: encode spec §5.5 and AC-VIEW-04. `SHEET_MIN_ZOOM = 0.25`, `SHEET_MAX_ZOOM = 4`, initial `1.0`, absolute CSS scale (not relative to container width). A 1000px-wide sheet in a 375px container stays at 100% and scrolls horizontally in sheet mode. Pinch/Ctrl+wheel/keyboard zoom keeps the pointer/viewport anchor through `ScrollWindow`. Per-sheet zoom survives sheet switches and sort/filter and resets only on document replacement. With the flag off, gestures and ordinary admitted unhidden-sheet rendering are byte-identical to the baseline. Preserve the S2.3 named authored-hidden golden exception and explicit oversized static-output failures; the flag does not suppress those required correctness fixes. Pin that `MIN_ZOOM === 1` and `MAX_ZOOM === 6` and every DOCX/PPTX/legacy-XLSX gesture are unchanged.
- [x] Run `bunx vitest run tests/xlsx-viewer-sheet-zoom.test.ts tests/officeview-zoom.test.tsx`; expected the sheet-mode assertions fail while the page-document assertions already pass.
- [ ] GREEN: implement sheet zoom as a scalar next to `ScrollWindow`, not as the translate/scale pan transform.
  - [x] Clamp with `clampZoom(z, SHEET_MIN_ZOOM, SHEET_MAX_ZOOM)`; do not share helpers that hard-code `MIN_ZOOM` (for example pan clamping at `z <= MIN_ZOOM`).
  - [ ] Add a mode switch in `OfficeDoc.tsx` selecting sheet mode only for XLSX with the interactive flag on; the flag defaults off until S13.
  - [ ] Feed `SheetZoom × DPR` to the §11.1 detail-admission math.
- [ ] Run focused command plus `bunx vitest run tests/officeview-selection.test.tsx tests/xlsx-viewport.test.ts` and standard corpus gates; expected ordinary admitted unhidden-sheet flag-off goldens remain unchanged; only the S2.3 named source-evidenced authored-hidden corrections may move, and oversized-output admission returns an explicit resource error.
- [x] Deliver the scale model as a small, independent PR before S5.2 wires real scrolling.

## S3 — Shared structural range index and refusal policy

### S3.1 Structural metadata without per-cell coverage

> **Status (2026-10-11):** pure structural index merged (`src/xlsx/view/structure.ts`; opencode fixes after codex-era review, Claude and opencode re-reviews, every finding fixed). Parser/`drawing.ts`/`types.ts` metadata extraction and the spill bridges are NOT delivered; only the checkboxes for the pure index are ticked.

**Files:** create `src/xlsx/view/structure.ts`; modify `src/xlsx/parse.ts`, `drawing.ts`, `types.ts`; bridge `src/xlsx/formula/spills.ts`, `saved-spills.ts`; test `tests/xlsx-viewer-structure.test.ts`.

- [x] RED: a full-domain merge/style interval, spill child without formula, saved/dynamic spill bounds, table headerRows=0/totals rows, marker and absolute drawings. Assert range queries return anchors/children/overlap in bounded memory, never a map entry for every covered cell.
- [x] Run `bunx vitest run tests/xlsx-viewer-structure.test.ts`; expected giant merge memory/child classification or missing anchor metadata fails.
- [x] GREEN: build rectangle/interval indexes for merges, spills, tables, styled/hidden dimension overlays and cell-anchored drawing markers. Rectangles remain true two-dimensional row/column ranges; a flattened cell-ID interval hull must not introduce phantom covered cells for narrow-column giant rectangles. `queryStructural` returns immutable range records, metadataCoverage/completeness and source anchor identities; preserve from/to offsets/anchor kind and resolve one-cell extents sufficient for complete occupied geometry projection. Absolute drawings have explicit independent placement.
- [x] Add nonempty-row/column-chunk indexes supporting current-region and outsiderange safety queries; do not repeatedly scan all 16,384 columns per candidate row.
- [ ] Run focused command plus existing merge/drawing/spill tests; expected exact hits, no per-cell coverage allocation and parser metadata remains compatible.
- [x] Deliver the shared structural API before selection, sorting or filtering consumes it.

### S3.2 Operation guard and drawing prefix impact

> **Status (2026-10-11):** guard merged (`src/xlsx/view/operationGuard.ts`, 76 tests across both S3 files; Claude and opencode reviews, every finding fixed). Integration into worker commands/UI is owned by S4/S7/S8.

**Files:** create `src/xlsx/view/operationGuard.ts`; test `tests/xlsx-viewer-operation-guard.test.ts`.

- [x] RED: allow read-only selection touching a spill child without its anchor; refuse sort×merge/spill and filter×merge/spill, including a distant blank merge/spill outside A:B within the affected row band and a cell-anchored drawing at source row 2 below filtered body rows 0–1. Pin operation refusal atomically: order/visibility/revisions/caches unchanged, status diagnostic emitted once.
- [x] Run `bunx vitest run tests/xlsx-viewer-operation-guard.test.ts`; expected naive row-span intersection accepts filter below-anchor displacement and fails.
- [x] GREEN: operation guard (dispatch each sub-item separately):
  - [x] Share one structural safety query among all consumers. Selection/highlighting is read-only and stays allowed on spill/merge children; deriving a sort/filter range from a selection queries the structural anchor/extent and refuses unsupported child-only ranges.
  - [x] Account for changing prefix geometry before/below retained cell anchors, not only operation-range intersection. v1 **refuses** every affected operation with a typed reason; drawing remapping is out of scope (spec §6.3 requires a separate reviewed contract). Absolute anchors follow the spec independently.
- [x] Structural guards query the entire affected body-row band across all 16,384 columns, even when keys/selected range are only A:B, because whole-row permutation/hiding affects distant blank/style-only merges and spills. Add a far-column blank merge/spill regression that cannot be found by the nonblank-data guard. Query all nonblank outsiderange columns anywhere in a sort band, including D values separated by blank C from selected A:B. Safe whole-row permutation may proceed only after this query completes authoritatively.
- [x] Run focused suite; expected anchored Y40→Y20 counterexample is refused with no view change, generated spill child is detected and isolated adjacent data never moves silently.
- [x] Deliver a reusable operation guard with typed refusal reasons and generation-aware query completion.

## S4 — Required resident worker, publication identity and lifecycle

### S4.1 Protocol and session ownership

> **Status (2026-10-11):** merged (Claude implementation; codex, pi and opencode reviews over three rounds; every finding fixed). Interim limits: the temporary object-backed `SourceStore` (`src/worker/viewerObjectSource.ts`, replaced in S1.2) reports drawings, dynamic-array spill extents and unreadable table lists as `incomplete` coverage with a `coverage-unresolvable` diagnostic, so sort/filter on those sheets wait as incomplete until S3.1 parser wiring; `viewport`/`search` return `unsupported` until S5; typed filters need the S8 evaluator wiring; scheduling/cancellation beyond request abort is S4.2.

**Files:** create `src/worker/viewerProtocol.ts`, `viewerSession.ts`, `viewerKernel.ts`, `viewerClient.ts`; modify `src/worker/office-worker.ts`, `worker-entry.ts`; test `tests/xlsx-viewer-worker.test.ts`, retain `tests/worker-render.test.ts`.

- [x] RED: open once, switch sheets, prepare/query repeatedly and assert parse count=1, no workbook clone per sort/filter and no detached source needed again. Two sheets reach local revision 1; assert their aggregate workbook index identity differs. Cross-document same-sheet IDs must not collide.
- [x] Run `bunx vitest run tests/xlsx-viewer-worker.test.ts tests/worker-render.test.ts`; expected current per-request parse or insufficient identity fails.
- [x] GREEN: resident document session owns source/pages, authoritative values, per-sheet views, prepared columns and cache handles. Send compact commands/ranges/query IDs and return typed permutations/masks/window records/diagnostics; make a single admitted initial compact transfer/copy, then reuse handles. Record actual ArrayBuffer copies and transfer ownership; 'typed' does not mean zero-copy.
- [x] Populate `commandKind` on every worker reply (S0.1 freeze addition to spec §10.2 that correlates the reply owner type with the command kind; refusals ride an `error` reply with code `operation-refused` and `retryable:false`).
- [x] Populate the S0.1 `space` tag (`source` or `display`) on every `BoundedRowBitset` (source membership versus `ViewCommitDelta.visibility`); the two spaces are deliberately not assignable to each other. Design note: `operation-refused` lives in the shared `SOURCE_ERROR_CODES` registry (reviewer-requested single registry); the generic error reply excludes it, so only sort/filter replies may carry it, with `retryable:false` and a `refusal` payload.
- [x] Add monotonically increasing workbook-wide `WorkbookOwner.globalViewGen` for every sheet view commit; per-sheet `viewGen` remains for local caches. Cache/publication identity includes docId, sheetId, valueGen, fontGen and viewGen; aggregate results capture `WorkbookOwner.globalViewGen` and the ordered `sheetVector` of every contributing sheet revision.
- [x] Run focused suite; expected once-only preparation, compact command payloads, two-sheet revision collision prevented, legacy one-shot DOCX/PPTX worker still passes.
- [x] Deliver a resident session PR; never call legacy full reparse for ordinary interactive commands.

Representative race assertion:

```ts
const before = session.globalViewGen
await session.commitView('A', aView) // A local revision becomes 1
const staleA = session.beginWorkbookIndex()
await session.commitView('B', bView) // B local revision also becomes 1
expect(session.globalViewGen).toBeGreaterThan(before + 1)
expect(session.publishWorkbookIndex(staleA)).toBe(false)
```

### S4.2 Cooperative scheduling, cancellation and atomic publication

> **Status (2026-10-11):** merged (Claude implementation; pi and opencode reviews over two rounds; every finding fixed). Shared cooperative scheduler (≤4,096 units or 10 ms) used by sort, preparation, guard scans and visibility projection; epoch/owner rechecks before every publication; one-transaction `commitDraft`; cancel reasons on replies; in-thread fallback for small documents, typed `worker-required`/`storage-required` refusal in large mode. Deferred with TODO markers: §10.2 viewport/active-cell priority (S5), 120 s query timeout, wire commands for value/font publication (S1.3/S4.3).

**Files:** extend `viewerProtocol.ts`, `viewerSession.ts`, `viewerKernel.ts`, `viewerClient.ts`; test `tests/xlsx-viewer-worker-races.test.ts`.

- [x] RED: delayed query/sort/index result arrives after another view commit, sheet switch, value publication, font change, close or reopen. Assert stale results never replace current order, UI counts, search geometry or status; cancelling a command preserves the previous committed view.
- [x] Run `bunx vitest run tests/xlsx-viewer-worker-races.test.ts`; expected stale publication or uncancellable work fails.
- [x] GREEN: request IDs, document epoch, workbook/per-sheet generations and cancellation tokens guard every publication. Long jobs yield at least every 10ms or 4,096 rows, whichever arrives first, including preparation, scans, external runs and value indexes. Provide progress from completed authoritative batches, not an optimistic timer.
- [x] Build drafts privately, validate resource/structural/source completion, then commit order/visibility/metrics/revisions as one transaction. Cancel on range replacement, source close or superseding draft; acknowledge input immediately while worker work continues. Define no-worker support separately: small admitted fallback may share chunked TS kernels; large mode refuses unsupported mandatory worker/storage features visibly.
- [x] Run focused suite with deterministic scheduler; expected bounded yield counts, cancellation reclaim, no half-applied views and no stale status/progress publication.
- [x] Deliver scheduler/publication patch with reproducible race logs, not timing-based unit assertions.

### S4.3 LRU/resource close and font/image leases

> **Status (2026-10-11):** merged (Claude implementation; pi reviews over two rounds; every finding fixed). One `MemoryPool` per document (D 512 / M 256 MiB, minimum one sheet mask) with no force path; LRU tiers rebuildable→source→view; evicted inactive views rebuilt from recipes and verified against the committed-order digest; value publication re-derives sorted orders on all sheets (`view-reapplied`); closed sessions never revive; `TileCache` closes surfaces exactly once (not yet wired to paint, S5.2); image/font lease leak in `getPaintables` fixed. F7 ledger plateaus (12 sheets, 4 MiB budget). Browser retained-memory verification remains S10.

**Files:** extend `viewerSession.ts`; create `src/xlsx/view/tileCache.ts`; integrate `src/core/images.ts`, `src/core/fonts/register.ts`, `src/render/paint.ts`; test `tests/xlsx-viewer-lifecycle.test.tsx`, retain `tests/fonts-index-lease.test.tsx`, `tests/cached-picture-pipeline.test.ts`.

- [x] RED: repeatedly open/close documents, change sheet/font/backend, cancel jobs, evict/reload pages and dispose while bitmap/font/image resources are in flight. Assert lease counts and retained owned bytes return to baseline, closed sessions cannot revive and previously active sheet views survive LRU eviction by source identity. F7 switches among many full-capacity sparse sheets with local sort/filter states; aggregate memory must plateau across sessions/sheets, not merely satisfy a per-sheet cap.
- [x] Run `bunx vitest run tests/xlsx-viewer-lifecycle.test.tsx tests/fonts-index-lease.test.tsx tests/cached-picture-pipeline.test.ts`; expected retained lease/session or stale publish fails.
- [x] GREEN: close removes listeners/pending entries/pages/external runs/column caches; close ImageBitmaps and GPU objects; retain/release existing image/font leases exactly once. LRU separates persisted source pages, retained compact sheet views and rebuildable layout/tile data. Inactive permutations/prefixes/prepared columns also count toward the single document-wide budget and may be externalized/evicted/rebuilt; never retain one 8MiB prefix per sheet without admission. Charge Wasm heaps, pending transfers and temporary buffers to the same worker budget.
- [x] Run focused command; expected cancellation/close recover allocations and cached/offscreen search records do not release fonts still needed by active paint.
- [ ] Deliver lifecycle PR with repeated-open memory ledger; browser retained-memory verification remains S10.

## S5 — Visible rendering, shared canonical records and global search

### S5.1 Canonical text/layout records and viewport queries

> **Status (2026-10-11):** merged (Claude implementation; opencode reviews over two rounds; every finding fixed or disputed with accepted evidence; xlsx goldens unchanged). Canonical records with Float64 logical bounds and tile-local paint coordinates, byte-bounded layout cache keyed by full identity and fontGen, §11.1 per-frame/per-tile caps with per-run coarse groups, exact overview source sets, unknown=loading, stale viewGen rejected, active-cell record carried beside `ViewportResult`. Not delivered: worker `viewport` command dispatch, store-path overflow/drawings (TODO S5.2), drawings still painted under a global translation.

**Files:** create `src/xlsx/view/textLayout.ts`, finish `viewport.ts`; modify `src/xlsx/render.ts`, `src/render/paint.ts`, `src/core/text-recording.ts`, `text-metrics.ts`; test `tests/xlsx-viewer-layout.test.ts`, extend `tests/xlsx-viewport.test.ts`.

- [x] RED: viewport at edge of F1/F3/F6; wrap, rotation, stacked text, General width formatting, merge anchor offscreen but merged area visible, hidden rows, fonts changing after initial layout and lone surrogates. Assert visited rows/cells/layout records scale with admitted viewport detail/pixel summaries, not all source rows or merge area. Include tiny positive row/column dimensions at 25% zoom so millions of coordinates could geometrically fit one viewport; a pixel-bounded canvas alone is insufficient. Instrument and assert maximum source pages/visited records/layouts/gridline spans/semantic nodes per viewport; a nonblank F6 provider must answer required source-only `queryOverview` analytically and terminate with complete coarse summaries or explicit resource/error/cancel outcome, not arbitrary samples marked complete or permanent loading on complete source. Add coarse projection RED: displayed [0,3,1,2] grouping rows {3,1} excludes row2, filtering changes exact membership, and late aggregates/layouts from old viewGen never publish.
- [x] Run `bunx vitest run tests/xlsx-viewer-layout.test.ts tests/xlsx-viewport.test.ts`; expected full scan/per-paint merge expansion/repeated wrap layout fails.
- [x] GREEN: canonical records and viewport queries (dispatch each sub-item separately):
  - [x] Source/window query plus structural intersection feeds canonical records containing source ID, displayed geometry, text runs/grapheme boundaries, clips, style/font identity, shaped/measurement data and logical search/copy information.
  - [x] Keep logical positions Float64; paint/GPU vertices and pointer math use tile-local coordinates. Test last-row global Y≈25Mpx at fractional zoom so Float32 global positions cannot introduce 2px quantization/jitter.
  - [x] Reuse existing host measurement/formatting behavior; the pixel backend consumes records instead of recomputing semantic layout.
- [x] GREEN: caching and bounded detail modes (dispatch each sub-item separately):
  - [x] Cache formatted values, text layout and merge queries across frames by complete identity and dimension/zoom-relevant keys; only changed records rebuild; `fontGen` invalidates dependent layout.
  - [x] Apply the specification §11.1 caps: desktop per-frame≤20k layouts/32 pages/4096 gridline instances and per-tile≤5000 layouts/8 pages; mobile per-frame≤10k/16/2048 and per-tile≤2500/4. Coalesce below-pixel gridlines and default/style spans.
  - [x] Detail admission uses `SheetZoom × DPR` (spec §5.5): individual detail requires both projected axes ≥4 device pixels. There is **no separate text font-size floor**: within caps and the geometry rule text paints at any projected size like the legacy renderer; text is omitted only in the declared coarse layer. Never apply these thresholds to the legacy fit-width path.
  - [x] Before enumeration, any density/size/cap crossing selects coarse mode and shows “Zoom in or inspect a cell for detail”. Use analytical/default summaries for F6 rather than enumerating 2^34 values; unknown summaries show loading, never an asserted blank; complete-source summaries finish or return an explicit resource/error/cancel.
  - [x] The worker forms bounded display pixel groups under immutable RowOrder/visibility/AxisMetrics, maps them to exact source sets/column intervals for `queryOverview`, then assigns viewport bounds under CacheIdentity/viewGen. Never summarize sorted disjoint rows through a source bounding box.
  - [x] Prioritize one exact active-cell detail/value inspector/accessibility record within budgets. Exact Go To/hit testing/active semantic cell/search/filter/copy stay available and correct; the display mode is declared, not a semantic truncation. Keep shape/rich-text/fallback font behavior and non-XLSX pipeline compatibility.
  - [x] RED counterpart: a small sheet at SheetZoom 0.25 (default 11pt text ≈ 3.7 CSS px) still paints its text within caps, and the legacy fit-width path at an effective 0.375 scale is unchanged.
- [x] Run focused and native text/rotation/merge suites plus exact goldens; expected visible work is bounded, zero-height rows create no text, canonical extraction and Canvas2D text agree.
- [x] Deliver reusable record/window layer before a GPU experiment; bitmap-only output is insufficient.

### S5.2 Canvas tiles, separate overlays and OfficeDoc integration

**Prerequisite:** S2.4 sheet scale model. The viewport integration below uses absolute `SheetZoom` and `ScrollWindow`; it never reuses the fit-to-width pan transform for sheet mode.

**Files:** extend `tileCache.ts`, `viewport.ts`; modify `src/components/OfficeDoc.tsx`, `src/render/paint.ts`, `src/core/overlay.ts`; test `tests/xlsx-viewer-viewport-integration.test.tsx`, extend `tests/officeview-zoom.test.tsx`.

- [ ] RED: scroll/zoom/pan produce bounded viewport requests; moving selection/search highlight does not increment base-cell paint count; view commit updates affected tiles while preserving zoom; source change resets zoom through existing ownership rules. Assert physical canvas/tile allocations stay below budgets.
- [ ] Run `bunx vitest run tests/xlsx-viewer-viewport-integration.test.tsx tests/officeview-zoom.test.tsx`; expected current all-base repaint/no viewport fails.
- [ ] GREEN: connect logical navigator to real scroll/zoom/DPR and active-sheet canvas tiles. Base tile cache is distinct from selection/search/watermark overlays as appropriate; overlay movement never repaints all cells. Prefetch bounded adjacent tiles, LRU reclaim bitmaps and keep last committed viewport while new records arrive.
- [ ] View generations invalidate geometry/search/base tiles explicitly; document ownership alone controls zoom reset. Watermark/protection remain applied to every presented region. Static export uses its explicit budgeted path, never stitching a full 17B-cell bitmap.
- [ ] Run focused command and browser smoke; expected no base repaint on overlay-only updates, correct edge recentering and retained zoom after sort/filter draft commits.
- [ ] Deliver interactive viewport integration behind a feature flag until S6–S10 complete.

### S5.3 Global workbook find without a full bitmap

> **Status (2026-10-11):** engine merged (Claude implementation; codex and opencode reviews over two rounds; every finding fixed). Find follows existing `core/search.ts` semantics (spec §11.4), scans visible authoritative cells in bounded cancellable batches, orders hits by display slot with a 262,144-hit buffer budget, pages one at a time, rejects stale/closed/cancelled publication, and reports typed source failures. Not delivered: OfficeDoc integration and worker `search` command (S5.2), browser far-corner smoke (S10).

**Files:** create `src/xlsx/view/find.ts`; modify `src/core/search.ts`, `src/components/OfficeDoc.tsx`, `src/render/paint.ts`; test `tests/xlsx-viewer-find.test.ts`, `tests/xlsx-viewer-find-races.test.tsx`; retain `tests/officeview-search.test.tsx`.

- [ ] RED: offscreen and other-sheet matches, filtered hidden match exclusion, sort changes match display order, source/value/font generation changes and A/B same-local-revision collision. Searching F1 must find edge text without painting all cells or building a huge bitmap.
- [ ] Run `bunx vitest run tests/xlsx-viewer-find.test.ts tests/xlsx-viewer-find-races.test.tsx tests/officeview-search.test.tsx`; expected viewport-only/global stale index failure.
- [ ] GREEN: paged source text scan/index over authoritative nonempty visible cells with bounded batches. Store logical source IDs and lightweight hit metadata; lazily obtain shared canonical geometry for focused/offscreen matches. Aggregate workbook cache key captures every contributing view and value/font generation; stale async publication is rejected.
- [ ] Navigate to match through active-sheet navigator/ScrollWindow, preserving per-sheet state and zoom. Text highlighting/copy uses the same layout records and source identities. Search must not trigger rendering/index materialization of every cell in a fully populated synthetic provider; expensive global source scans are cancellable and admitted as such.
- [ ] Run focused command and browser smoke for far-corner find; expected result order/geometry matches displayed view, offscreen result appears after navigation and no zoom reset or old workbook index overwrite.
- [ ] Deliver workbook find PR with scan/record allocation counts and DOCX/PPTX search parity.

## S6 — Active sheet, source-identity selection, virtual headers and accessibility

### S6.1 Navigator and selection state machine

> **Status (2026-10-11):** merged (codex implementation; Claude fixes for the last MINORs; Claude and pi reviews over four rounds; every finding fixed). Source-owned selection with display-space extension, serializable gesture state, bounded parts (4,096 parts / 16,384 column intervals / 1 MiB per selection, no document-wide cap), operation ranges only via S7.1, S2 navigation reuse. OfficeDoc/S5 viewport wiring remains a documented seam.

**Files:** create `src/xlsx/view/selection.ts`, `src/components/xlsx/SheetNavigator.tsx`, `ViewerStatus.tsx`; modify `src/components/OfficeDoc.tsx`, `src/core/overlay.ts`; test `tests/xlsx-viewer-selection.test.ts`, `tests/xlsx-viewer-navigation.test.tsx`.

- [x] RED: display order [0,3,1,2], Shift from source row 3 (display slot 1) to source row 1 (display slot 2) selects source membership {3,1}, not source interval [1,3]; Shift from source row 0 (slot 0) to source row 1 (slot 2) selects {0,3,1} and excludes source row 2. Anchors/targets are always source cells and slots are always written as slots, never as "displayed row N". Later sort preserves that membership. Test active cell survives sort by source ID, filter moves focus to nearest eligible displayed candidate, all body rows hidden sets active=null and anchor=null; Clear restores visibility but retains null active/anchor until an explicit navigation/Go To/pointer action creates them. Test nearest eligible visible body relocation when one exists, with authoritative source/display identity and no hidden candidate or unsolicited focus resurrection.
- [x] Run `bunx vitest run tests/xlsx-viewer-selection.test.ts tests/xlsx-viewer-navigation.test.tsx`; expected rectangle-only/visible-rank identity fails.
- [x] GREEN: active/anchor source coordinates, selection as disjoint source row/column intervals or compact membership sets, display-space extension converted through projection, Ctrl/Cmd toggles, row/column header selection and visible-order arrow traversal. All operation consumers share S3 structural guards; read-only selection may highlight structural children. A disjoint source selection must not be converted to its bounding rectangle for sort/filter: use an explicit contiguous SourceRect/table/current region, or refuse an ambiguous multi-range operation with notice.
- [x] Implement active-sheet navigation for large mode, per-sheet retained view/selection/zoom, edge coordinate entry and visible status counts. Go To outside a filtered body may still address a legal blank coordinate; distinguish displayed-address UI from explicit Original cell/source API through S2.2 mappings. Announce displayed address/value and original address when different (for example “A2, value 10; original A3”). Hidden targets use an exact read-only inspector/hidden reason without changing visibility or invalidating current nearest-visible/null grid active state. Add display/source navigation and hidden-target tests for [0,2,1], with source formula/search refs unaffected; default record counts use resolved used-data body, never all implicit million blank rows. Selection overlay clips to viewport records and stays distinct from blue text selection.
- [x] Run focused tests with hidden-pinned interleavings and structural fixtures; expected membership invariance, nullable active, source-safe selection and correct active-sheet resource lifecycle.
- [x] Deliver shell/state machine independently; no sort/filter dialog needed yet.

Preserve this counterexample:

```ts
const projection = fixtureProjection([0, 3, 1, 2])
// extendDisplayRange(projection, anchor: SourceCell, target: SourceCell); both arguments are SOURCE cells
const selection = extendDisplayRange(projection, { row: 3, col: 0 }, { row: 1, col: 0 }) // source 3 = slot 1, source 1 = slot 2
expect([...selection.sourceRows()].sort((a, b) => a - b)).toEqual([1, 3])
expect(selection.contains({ row: 2, col: 0 })).toBe(false)
const wide = extendDisplayRange(projection, { row: 0, col: 0 }, { row: 1, col: 0 }) // source 0 = slot 0, source 1 = slot 2
expect([...wide.sourceRows()].sort((a, b) => a - b)).toEqual([0, 1, 3])
expect(wide.contains({ row: 2, col: 0 })).toBe(false)
```

### S6.2 Virtual owned grid semantics and event isolation

> **Status (2026-10-11):** merged (codex implementation; Claude fixes; Claude and opencode reviews over two rounds; every finding fixed). ARIA grid with header-inclusive indices (1,048,577 × 16,385), bounded DOM, one Tab entry, control mode only via F2, copy/print protection and Escape fall-through preserved under the viewer UI guard. Not mounted until S5.2; desktop/mobile screen-reader checks are S10.

**Files:** create `src/components/xlsx/GridSemantics.tsx`, `GridHeaders.tsx`; modify `src/components/OfficeDoc.tsx`; test `tests/xlsx-viewer-a11y.test.tsx`, `tests/xlsx-viewer-events.test.tsx`; extend existing selection/zoom suites.

- [x] RED: assert grid owns a bounded row→gridcell/header DOM tree, total semantic row/column counts, header-inclusive `aria-rowindex`/`aria-colindex`, `aria-selected`, `aria-sort`, valid `aria-activedescendant` and live status. Navigation to XFD1048576 exposes correct indices without 17B DOM nodes. Semantic grid counts are MAX_ROWS+1=1,048,577 and MAX_COLS+1=16,385 to include header chrome: corner row1/col1, letter headers row1/col(c+2), data rows row(displaySlot+2), rowheader col1, data cells col(c+2). Last data cell must exactly equal both counts.
- [x] Run `bunx vitest run tests/xlsx-viewer-a11y.test.tsx tests/xlsx-viewer-events.test.tsx`; expected count-only accessibility/independent header ownership fails; header/body indices must not collide.
- [x] GREEN: virtual viewport/active semantic rows and cells with unique generation-safe IDs, columnheader/rowheader roles owned through rows, one managed Tab entry and grid-owned aria-activedescendant. Explicit nested header-control mode restores valid active-grid focus on Escape. Keep active cell mounted or move its semantic record atomically; nullable active removes invalid descendant. Worksheet labels use display slot+1 and A..XFD, while semantic data indices include header offset +2 and values identify source cells; do not label compressed visible rank as row number. Expose full worksheet capacity separately from semantic chrome counts.
- [x] Sticky headers share ScrollWindow/zoom/DPR transforms, viewport-only elements and native button click zones. Grid arrows stop container scrolling only when grid owns focus; Esc priority popup→grid selection→existing handler, Ctrl/Cmd+F unchanged.
- [x] Ignore viewer gestures from editable targets and viewer UI subtrees; negative/zero filter typing must not reset zoom, header double-click must not pan/reset, checkbox Space and native Enter work. Mobile targets/hit-slop meet 44px without changing row geometry.
- [ ] Run focused command plus `bunx vitest run tests/officeview-selection.test.tsx tests/officeview-zoom.test.tsx` and browser keyboard smoke; expected ownership/focus/actions and event isolation pass. Record desktop/mobile screen-reader manual checks in S10.
- [x] Deliver header/a11y PR; DOM shell should not move default canvas goldens.

## S7 — Stable whole-row sort and authoritative range detection

### S7.1 Range detection/header policy/structural admission

> **Status (2026-10-11):** merged (opencode implementation; Claude fixes; reviews by Claude, codex and pi over three rounds; every finding fixed). Region detection is a bounded async ring-growth probe (full-capacity dense column: ~93 MB peak vs 1.8 GB before) with a work budget; orchestrator decisions: a blank seed resolves to the sole current region in the used extent or refuses `ambiguous-selection`; any derived range intersecting a table it does not exactly equal, overlapping tables at the active cell, and a contiguous selection not containing the active cell refuse `ambiguous-selection`.

**Files:** create `src/xlsx/view/rangeDetection.ts`; consume `structure.ts`, `operationGuard.ts`; test `tests/xlsx-viewer-sort-range.test.ts`.

- [x] RED: table→current-region→used-range precedence, table headerRows=0/1, totals rows, all-text data, explicit override, blank gap before D outsiderange values and a pending outsiderange page. Assert no commit before complete safety query.
- [x] Run `bunx vitest run tests/xlsx-viewer-sort-range.test.ts`; expected naive adjacent-only/header-always-true acceptance fails.
- [x] GREEN: derive source-coordinate whole-row band with explicit protected header/totals, allowed override and authoritative used extent. Query indexed nonblank content in every outside column anywhere in the band; refuse conflicting data, merges/spills or geometrically affected drawings through the shared guard.
- [x] Run focused suite; expected table metadata precedence, source ranges immutable, separated D data refuses atomically and pending checks wait/cancel without fake safe state.
- [x] Deliver range detector independently for engine/UI reuse.

### S7.2 Worker sort kernel and cycle commands

> **Status (2026-10-11):** pure sort engine merged (`src/xlsx/view/sort.ts`, 82 tests; reviewed by Claude and opencode over three rounds, codex implementer). Worker kernel/session commands, header cycles, external-run merge and browser smoke are NOT delivered and stay unchecked.

**Files:** create `src/xlsx/view/sort.ts`; extend `src/worker/viewerKernel.ts`, `viewerSession.ts`; integrate `GridHeaders.tsx`; test `tests/xlsx-viewer-sort.test.ts`, `tests/xlsx-viewer-sort-integration.test.tsx`.

- [x] RED: comparator truth table, numeric/date serial, host `Intl.Collator` accent/variant sensitivity with `numeric:true`, caseFirst upper, booleans, mixed types, errors before blanks, both directions blanks last, source-ID tiebreak. Test 8 keys, repeated sort/unsort, filtered-pinned rows, source formula/string immutability and source high-row entering top viewport.
- [ ] Run `bunx vitest run tests/xlsx-viewer-sort.test.ts tests/xlsx-viewer-sort-integration.test.tsx`; expected comparator/whole-row/visibility/stability assertion fails.
- [ ] GREEN: prepare at most eight columns in typed batches; stable permutation by source-row ID ties, sorted visible eligible slots only, pin hidden/protected/outside rows. Multi-key/case options and unsorted→asc→desc→unsorted commands publish one view transaction; Shift+sort appends under documented 8-key UI limit.
- [ ] Large admitted numeric sorts use resident prepared columns/external runs as required; comparators must never asynchronously load OPFS/IndexedDB pages or create NlogN random page churn. Prepare bounded key runs first, stable-sort admitted runs, then merge sequentially with cancellation/backpressure and exact host semantics. Count key lanes, row/index/scratch/string buffers and run readers for eight-key sorts toward aggregate budgets; natural text compares through the host collator. Do not reparse, recalculate formulas, clone the workbook or scan every column. Unknown source values prevent a complete sort commit until resolved/admitted.
- [ ] Run focused command plus formula/native strings suites and browser top/edge sort smoke; expected row values/styles/source refs stay paired and source intact, no legacy-cap alias, search/selection/zoom retained.
- [ ] Deliver sort engine/header cycles with RED/GREEN logs and measured microbench evidence clearly labeled non-browser; S10 owns SLA acceptance.

## S8 — Typed predicates, snapshot visibility and complete value universe

### S8.1 Typed filters and day/period boundaries

> **Status (2026-10-11):** pure filter/date engine merged (codex implementation; Claude round-3 fixes; reviews by Claude and opencode over four rounds; every finding fixed; 0 mismatches in ~1.85M fast-vs-reference comparisons across 34 locales). The collator certificate and its `resource-limit` residual for uncertifiable text (`ß`, NFD marks, Danish `aa`+`å`) are removed by S8.5, which replaces collator matching with the canonical form of the amended spec §7.1. Mask publication through S3 guards, the formatted display-text lane and worker integration are NOT delivered.

**Files:** create `src/xlsx/view/filter.ts`, `datePredicates.ts`; extend worker kernel/session; consume existing `src/xlsx/formula/serial.ts`, `functions/wildcard.ts`; test `tests/xlsx-viewer-filter.test.ts`, `tests/xlsx-viewer-date-filter.test.ts`.

- [x] RED: number/text/date conditions, AND across columns, explicit two-predicate AND/OR union, blanks/errors table, wildcard escapes and lone surrogates. Text conditions use host Intl.Collator with sensitivity=accent, numeric=false, usage=search; test filter eq numeric strings 2 vs02 distinct, A/a equal, e/é distinct, prefixes/suffixes/subsequences, emoji/lone surrogate '?' matching one code point and invalid trailing '~' escape. Header 'Amount' with body 5 and numeric >3 retains the header; matching authored-hidden body reveals; outside hidden/totals keep base flags; replacement/Clear restores snapshots.
- [x] RED: text-evaluable semantics (spec §7.1). In a mixed column, text `contains "2"` matches the number 20 through its formatted display text and does not match an error or blank; number `gt 3` never matches a string cell; `neq`/`not-contains` complement only within the evaluable set; date-typed numeric cells also evaluate number conditions on their serial values, and date cells in text/mixed columns evaluate text conditions on formatted display text; a pure number column offers no Text Filters and a text/mixed column offers both. Pin each case so the viewer-specific decision cannot drift toward raw-value matching.
- [x] Add date counterexamples: serial D+0.5 equals day D, before excludes day start, after begins at next day, inclusive end-day between includes its fractional times; standalone datePreset lowers today/week/month/year to [start,nextStart). Pin injected nowMs/locale/timeZone, Sunday week, month/year/DST boundaries, 1900 serial 60 and 1904 date systems using existing calendar helpers.
- [x] Run `bunx vitest run tests/xlsx-viewer-filter.test.ts tests/xlsx-viewer-date-filter.test.ts`; expected literal in-range predicate hides header or scalar date equality excludes midday and fails.
- [x] GREEN: implement discriminated typed predicates; datePreset is a standalone kind, never RelativeDate accepted as any arbitrary operator operand. Text literals use host accent-sensitive/case-insensitive collator, numeric=false (sorting remains numeric=true); contains/begins/ends compare candidate subsequences/prefixes/suffixes. Wildcards '*' match zero-or-more code points, '?' exactly one including a lone surrogate; '~' escapes and invalid trailing escape is diagnosed. Use memoized DP or bounded equivalent rather than exponential backtracking and retain raw UTF-16 semantic keys. Respect source date formatting/classification and existing serial semantics; resolve periods at query time with captured nowMs and generations.
- [ ] Compute compact mask/rank-select over filter body; protect header/totals visibility and apply S3 safety queries before publication, including drawing prefix effects. Counts include full authoritative filter body beyond legacy caps; no hidden text records leak.
- [ ] Prepare the formatted display-text lane for numeric/date/boolean cells only when a text condition targets a text/mixed column; bound it to the aggregate budget, key it by `valueGen` and style owner, and drop it on invalidation.
- [ ] Run focused suites plus native/date/formula suites; expected type/compound/calendar/visibility cases pass, no source mutation or contradictory hidden predicate remains.
- [x] Deliver filters/date engine independently of dropdowns.

### S8.2 Semantic value index and searched Select All

> **Status (2026-10-11):** merged (Claude implementation; codex and opencode reviews over seven rounds; every finding fixed). Complete query-backed value universe with typed semantic keys, all-minus-exclusions/include compaction, searched-subset Select All, async draft preparation, charged memory with typed quota errors, cancellable/resumable builds. Search uses the S8.5 canonical matcher; default `EXACT_SEARCH` (no resident keys; ≈0.12–0.2 µs/value, Bun, non-browser), canonical keys opt-in; no persistence in v1 (value-search spike). Worker/session wiring and `FilterEngine` values rules are not delivered (S4/S9).

**Files:** create `src/xlsx/view/valueIndex.ts`; extend `storage/externalRuns.ts`, worker kernel/session; test `tests/xlsx-viewer-value-index.test.ts`.

- [x] RED: 100k and 1M distinct values, tagged numeric/text/boolean/error/date keys, same display with different serials, search match after displayed first 10k, unchanged OK, all-minus-exclusions, searched-subset Select All, unchecking one entry and cancel/reopen. Unknown pages must not make the value universe complete.
- [x] Run `bunx vitest run tests/xlsx-viewer-value-index.test.ts`; expected 10k construction cap or selected-string-array omission fails.
- [x] GREEN: lazy query-backed complete value universe with paged index, semantic key membership or all/query-minus-exclusions representation. 10k limits displayed entries only; dropdown search queries all values and Select All affects full universe or exact searched subset. Display dedupe never loses semantic keys. Represent blank separately and error identity explicitly.
- [x] Build in cancellable batches; cache by sheet/source values/range/column/classification generations; charged storage/worker memory with quota admission. Exclude display-only transforms from key identity. Reopen uses complete cache or resumes valid pages; cancel closes draft without changing committed rules.
- [x] Run focused suite and injected quota/cancel races; expected no undisplayed-value exclusion on unchanged OK, complete counts/search membership and no excessive strings/objects retained on main.
- [x] Deliver value-index PR with cold/warm allocation and source-read counts, not an implementation of a capped 10k semantic universe.

### S8.3 Declared filter diagnostic parsing

> **Status (2026-10-11):** merged (opencode implementation; Claude fixes; codex, Claude and pi reviews over four rounds; every finding fixed or disputed with primary-source evidence accepted by the reviewer). Declared autoFilter/sortState (worksheet and table, direct and nested) parsed as diagnostics only, bounded by a 256 KiB per-document byte budget, adapter-neutral builders ready for the large adapter. Small/large adapter parity is pending the large adapter (S1.2); Clear restores from the existing file-hidden snapshot.

**Files:** modify `src/xlsx/parse.ts`, page importer metadata; test `tests/xlsx-viewer-declared-filters.test.ts`.

- [x] RED: package autoFilter refs/kinds, unsupported color/Top-10/average filters, supported filters saved only as hidden rows, sortState source order and unsupported dynamic preset. Small/large adapters must expose identical diagnostics.
- [x] Run `bunx vitest run tests/xlsx-viewer-declared-filters.test.ts`; expected missing/mismatched diagnostic or mistaken authored-view import fails.
- [x] GREEN: parse declared range and rule kinds for diagnostics only per spec; preserve file-hidden snapshot and original stored order. Do not silently import unsupported filter state or invent funnel state from hidden rows. Clear restores load snapshot, including file-declared filtering effects.
- [ ] Run focused command plus adapter parity and standard gates; expected supported saved effects render correctly through S2 hidden fix, no formula/table parser regression.
- [x] Deliver parser delta through the shared integration owner.

### S8.5 Canonical text matching (spec §7.1 amendment, approved 2026-10-11)

> **Status (2026-10-11):** merged (Claude implementation; codex reviews over two rounds; every finding fixed: Turkic After_I/Before_Dot over the post-removal sequence with pinned combining classes, single raw-length admission on raw and key paths, independent test oracle). 0 mismatches in 600k differential comparisons; linear cost (≈2 µs per Latin value at 2,000 chars, Bun, non-browser).

**Files:** modify `src/xlsx/view/textFilter.ts`, `filter.ts`; create `src/xlsx/view/canonicalText.ts` and generated pinned Unicode data (case folding incl. Turkic, default-ignorable set, extended-grapheme segmentation rules) under `src/xlsx/view/unicode/` with its generator script; test `tests/xlsx-viewer-canonical-text.test.ts`, update `tests/xlsx-viewer-filter.test.ts`.

- [x] RED: every row of the "verified differences" table in `docs/validation/s8-1-text-match-amendment-proposal.md` (Danish `aa`/`aA`, `å`/`aa`, `å`/`a\u200ba`, `a`/U+0363, `ß`/`ss`, `ﬁ`/`fi` with `?`, `contains('e','é')` vs `e\u0301`, `?` vs `é`/`e\u0301`, `?` vs lone U+200B), Turkish/Azeri dotted/dotless I incl. decomposed `I`+U+0307, emoji ZWJ sequences and flags as single graphemes, lone surrogates as single units, default-ignorables removed, `~` escapes unchanged.
- [x] GREEN: canonical pipeline (NFD → remove default-ignorables → locale-aware full case fold → NFD) over pinned data independent of the host's ICU; grapheme segmentation of the canonical form; equality/begins/ends/contains on grapheme boundaries with linear-time matching (e.g. KMP over grapheme-unit IDs) and wildcard `*`/`?` over grapheme units with memoized/DP or linear glob matching; remove the collator certificate and the cubic reference path from the filter (keep a slow test-only oracle built from the spec text for differential tests).
- [x] Derived canonical key per value is cacheable (S8.2 search-key port supplies `searchKeyOf`); raw values/semantic keys unchanged.
- [x] Differential/property tests: ≥100k seeded cases comparing the optimized matcher with the spec-text oracle over the S8.1 hostile pools; no `resource-limit` for ordinary text up to the §12 text-length targets; cost linear in text length.
- [x] Deliver with measured (non-browser, labelled) per-value cost for contains/equals/wildcard at 100–2000 characters, and remove the S8.1 residual from §0.4.

### S8.4 Performance threshold freeze (exit gate for S8)

**Files:** create `tests/fixtures/xlsx-viewer/performance-profile.json`; extend `scripts/bench-interactive-viewer.ts`; write `docs/validation/interactive-viewer-performance.md` (baseline section).

- [ ] Run the real-browser harness (S0.3) on the D profile, Chrome, against the delivered S7/S8 kernels for the F2 numeric sort/filter, F4 natural text sort and cold 100k value-index operations, with the sample counts required by spec §12.6. Record kernel, query-completion and input-to-present times separately.
- [ ] Open a reviewed PR that freezes each threshold in `performance-profile.json` as either the spec §12.2 value or a measured, justified replacement. Loosening any value more than 2× needs a specification amendment citing the browser measurements. Record the F4 headroom explicitly (spec estimates about 3.1 s against the 4 s ceiling).
- [ ] If a threshold is unmet, apply only the spec-permitted levers (ASCII/Latin-1 fast path proven identical to the host collator, run-merge sort, worker-parallel runs); never relax collation or semantics. Do not start S11/S12 as a substitute.
- [ ] From this point S10.2 enforces the frozen values; before it, a miss is a recorded finding, not a failure.

## S9 — Controls, copy/print policy and mobile interactions

> **User-approved behaviour changes (2026-10-11), to be surfaced by S9 copy and listed in the final PR:**
> 1. A remote spreadsheet whose type cannot be determined from its first 64 KiB, too large for the memory spool, with neither OPFS nor IndexedDB available, is refused with: "This file is too large to open from the web in this browser. Download it and open it locally."
> 2. Workbooks whose load-time evaluation would exceed the memory budget open with the file's cached values; show a status line such as "Showing saved values — N formulas not recalculated (file too large)"; uncached formula cells show as unavailable (a "Recalculate anyway" action is a possible later addition).
> 3. Whole-sheet static export/print of sheets beyond the legacy 16,384×4,096 grid (or drawing sheets beyond the old drawing clamp) is refused instead of silently truncated, with: "This sheet is too large to export as one image. Choose a range or print selected pages." Ordinary sheets within the legacy grid render exactly as before.

### S9.1 Atomic filter dropdown and sort dialog

> **Status (2026-10-12):** merged (codex implementation; Claude fixes for the last MINORs; Claude reviews over two rounds plus opencode delta check; every finding fixed). Atomic OK/Cancel drafts over the S8.2 value universe, typed AND/OR/datePreset builder, ≤8-level focus-trapped sort dialog, always-visible header sort-arrow and funnel buttons under the S6.2 focus model, mobile bottom sheet, the three user-approved notices in ViewerStatus. Mounting into OfficeDoc happens in S5.2; real-device touch/screen-reader checks are S10.

**Files:** create `src/components/xlsx/FilterDropdown.tsx`, `SortDialog.tsx`; integrate `GridHeaders.tsx`, `ViewerStatus.tsx`, `OfficeDoc.tsx`; test `tests/xlsx-viewer-filter-ui.test.tsx`, `tests/xlsx-viewer-sort-ui.test.tsx`.

- [ ] RED: open→search beyond 10k→Select All→exclude→OK, Cancel/Esc→reopen, worker result arrives after sheet/source/value generation changes, replace range, refused operation, custom typed AND/OR/datePreset builder and no-header table. Assert draft/committed state and generations change only atomically on successful OK.
- [ ] Run `bunx vitest run tests/xlsx-viewer-filter-ui.test.tsx tests/xlsx-viewer-sort-ui.test.tsx`; expected incomplete value selection/draft leakage/focus failures.
- [ ] GREEN: dropdown uses paged universe and bounded list DOM, sort/clear shortcuts, typed condition/date presets, blanks/errors, inline display-cap note and progress/cancel. Sort dialog adds/removes/reorders up to eight levels, case/header override and formula-view-only note. Disable/diagnose pending source or unsafe commits; do not display successful counts before authoritative transaction publication.
- [ ] Esc cancels/focus returns to invoker; Enter applies only the correct form action; sort dialog is focus-trapped, dropdown is not. Render file text via React text nodes/textContent, never innerHTML. Separate letter-select, sort-cycle and funnel-open buttons with accessible names; funnel/arrow/title/status track committed state.
- [ ] Mobile dropdown uses bottom-sheet style and 44px targets; tap provides tooltip-equivalent status. Counts are visible body N / complete body M with no-filter used-body fallback; removing the final column rule clears FilterSpec to null and restores snapshot visibility; stale jobs cannot overwrite a current notice.
- [ ] Run focused command and browser keyboard/touch scenarios; expected generation-safe draft behavior, complete Select All, no viewer gesture while typing 0/-1 and maintained header/grid semantics.
- [ ] Deliver controls PR with a11y/mobile manual record and standard gates.

### S9.2 Copy in displayed view order and budgeted static print

**Files:** create `src/xlsx/view/copy.ts`; modify `src/core/clipboard.ts`, `src/components/OfficeDoc.tsx`, `src/render/paint.ts` budget integration; test `tests/xlsx-viewer-copy.test.ts`, `tests/xlsx-viewer-policy.test.tsx`.

- [ ] RED: disjoint source membership after Shift/re-sort, filtered hidden rows, multiple selected ranges/columns, far-corner cells, formatted values, literal errors, tabs/newlines and surrogate strings. Clipboard limit is 8MiB of raw UTF-16 code units×2 bytes: a 5MiB ASCII string is 10MiB UTF-16 and must reject; lone surrogates count their original code units without replacement. Copy >8MiB must reject the entire operation, never silently truncate; denied allowCopy/allowPrint does not access clipboard/print.
- [ ] RED: a one-million-row single-column short-value selection fitting the complete byte/session budget is accepted without an arbitrary 100k-cell cap. Selecting the fully visible/default-unhidden full worksheet fails its separator-byte lower bound before any source-cell reads; assert bounded metadata/count queries and zero selected-cell enumeration. Selecting the full logical sheet with only one visible row and column may fit and is accepted; preflight always uses actual visible counts.
- [ ] Run `bunx vitest run tests/xlsx-viewer-copy.test.ts tests/xlsx-viewer-policy.test.tsx tests/clipboard.test.ts tests/officeview-protection.test.tsx`; expected source-order copy/partial truncation/policy violation fails.
- [ ] GREEN: project selected source membership into current visible display order, query authoritative values in bounded batches without allocating the selected row×column product, dedupe overlapping membership, assemble bounded TSV/text with declared escaping, and write once only after full budget check. Preserve source selection after later sorts; unknown pages keep copy pending/cancellable, not blank. Preflight byte accounting uses the final clipboard string’s raw UTF-16 code-unit length×2, including delimiters/escaping, under the 8MiB ceiling. Preserve lone surrogates; do not budget via UTF-8 encoding or replacement.
- [ ] Implement the rectangular separator lower bound `2 * (R * (C - 1) + (R - 1))` bytes from compact visible-membership counts, plus conservative disjoint-block bounds. Refuse provably oversized output before reads, then enforce exact byte/session reservations during bounded worker serialization; capacity alone never creates a lower fixed row-count limit.
- [ ] Enforce existing copy/print permissions across keyboard, UI and programmatic path. Static print/export reports selected extent and budget; if full output exceeds admitted pages/pixels/memory/time, refuse with status/diagnostic or require an explicit bounded selection rather than outputting an unlabeled subset.
- [ ] Run focused tests and browser clipboard/print mock scenarios; expected whole-operation budget refusal, correct visible order and protected content blocked.
- [ ] Deliver policy/copy PR with readable output examples and no changes to source content.

## S10 — Mandatory full-scale hardening and controlled release performance

### S10.1 End-to-end functional and complexity audit

**Files:** expand `tests/browser/interactive-viewer.spec.ts`, `tests/xlsx-viewer-complexity.test.ts`; maintain fixture manifest; create `docs/validation/interactive-viewer.md`.

- [ ] RED: encode F1 navigation/find/copy→sort/filter→clear, F2 sort/filter/cancel, F3 horizontal recenter/headers, F4 first/reopen/searched Select All, F5 giant merge/spill/drawing refusal, F7 same revision races/reload/fonts and F6 geometry window. The F6 test is not a whole-data sort performance claim.
- [ ] Run `bunx vitest run tests/xlsx-viewer-complexity.test.ts` and `bunx playwright test --config playwright.viewer.config.ts`; expected failures surface unfinished full-domain paths, stale records or unbounded visit counts.
- [ ] GREEN: fix only reproduced gaps. Audit absence of dense `R×C` allocation, per-blank-row/cell objects, all-column scans per row, full-sheet canvases, merge coverage maps, per-command workbook clone/reparse and lossy bitwise coordinate IDs. Instrument query/copy/transfer/layout counts meaningfully rather than tests mirroring implementation.
- [ ] Verify real XLSX corner fixture plus synthetic fully populated provider; compare both source adapters and native Excel cache/error/date/hidden/table examples. Record native workbook source/checkpoint/screenshots and absence of repair warnings; observational UI is not a code-test receipt.
- [ ] Run focused failures again, then standard gates and browser suite; expected all required cases pass with reviewed goldens and retained protections/non-XLSX behavior.
- [ ] Deliver independent functional/complexity evidence matrix by requirement ID and fixture; incomplete mandatory cases block release.

### S10.2 Real-browser performance, memory and device gate

**Files:** finish `scripts/bench-interactive-viewer.ts`, browser performance scenarios, result schema/thresholds under `tests/fixtures/xlsx-viewer/performance-profile.json`; write `docs/validation/interactive-viewer-performance.md`.

- [ ] Define controlled desktop profile: Apple M4 Pro, ≥8GiB available RAM, current Chrome and actual Safari. Mobile profile: iPad A14 with 4GiB RAM, actual Safari. Pin fixture seed, local-source bytes, viewport (desktop 1440×900, mobile 1024×768 CSS px), DPR=2, fonts, zoom=1 and cache states; record actual device/browser versions on each run. Alternative hardware cannot silently inherit these thresholds.
- [ ] Measure at least 300 sustained-scroll frames and 30 operation samples where feasible after stated warmup; cold-open/index/module runs use fresh sessions and cache/reset evidence. Expensive 1M cold operations may use fewer samples only with explicit limitation, never present three warmed medians as a p95 distribution. Record p50/p95, scroll input-to-present, main long tasks, import bytes/throughput, first viewport time, worker operation and end-to-end latency separately. Report mobile thermal/background conditions and failures, not only best trials.
- [ ] Measure memory with application buffer/cache/lease ledger on every browser; include main heap owned by viewer, resident worker including source working pages/Wasm/temp/copies, GPU textures/buffers and persistent storage separately. Use browser/OS retained-memory instruments where available; unsupported native probes are labeled unavailable, not zero. Capture baseline, peak, quiescent-after-cancel and post-close.
- [ ] Run desktop automation: `bun scripts/bench-interactive-viewer.ts --browser chromium --profile desktop --fixture all --repeat 30 --mode release --output /tmp/officeview-viewer-desktop-chrome`. Repeat browser fixtures in Playwright WebKit for automated regression, then run actual desktop Safari collection with `bun scripts/bench-interactive-viewer.ts --browser safari-manual --profile desktop --fixture all --repeat 30 --mode release --serve --port 4173 --output /tmp/officeview-viewer-desktop-safari` and physical iPad collection with `bun scripts/bench-interactive-viewer.ts --browser safari-manual --profile mobile --fixture all --repeat 30 --mode release --serve --port 4173 --output /tmp/officeview-viewer-ipad-safari`. Open the emitted harness URL in the actual target browser/device, import the manifest fixtures, execute the prompted cold/warm sequences and export the hash-tagged JSON receipt back to the collector. The receipt carries a SHA-256 over its canonical JSON (build commit, device/browser identity, fixture hashes, samples) so accidental or casual modification is detectable; no cryptographic signing or key management is claimed, and attestation is the named collector recording the run in the PR. Expected: complete device/browser/profile identity and sample schema; unsupported memory instruments remain explicit. Actual Safari/iPad receipts are mandatory for the named target profile.
- [ ] Apply the thresholds frozen in S8.4 (the table below lists their initial values) only in this controlled release job. Missing samples/metrics/physical profiles or a exceeded threshold fails release admission; ordinary unit tests never fail on wall-clock budgets. File first-viewport deadlines start when the local seekable Blob/complete compressed package is available and cover F1/F2/F3 with their manifest byte sizes, not arbitrary input size or entire-file completion. Network download/spool/central-directory inventory is reported separately.
- [ ] Investigate failed thresholds by separating import/decode/storage, source scan/prepare, worker scheduling/transfer, projection, layout and paint/present. Rerun only changed/failing cases after fixes, then one complete release run. Do not select Rust/GPU as a substitute for a reproduced pipeline bottleneck.
- [ ] Deliver hash-tagged result artifacts and a summary with all target thresholds, regressions from baseline and memory recovery. Claims about performance remain proposals until these receipts exist.

| Acceptance measurement | Desktop ceiling | Mobile ceiling | Required fixture/meaning |
|---|---:|---:|---|
| Warm viewport input-to-present p95 | 16.7ms | 33.3ms | ordinary/wrapped/styled scrolling; 60fps target when feasible |
| Input acknowledgement | 50ms | 50ms | sort/filter/find/cancel UI acknowledgement while worker runs |
| 100k numeric sort, 1–8 keys | 100ms | 300ms | admitted prepared/query workload, end-to-end separately recorded |
| 1M numeric sort, 1–8 keys | 1,000ms | 3,000ms | F2 complete authoritative band; cold preparation recorded separately |
| 1M numeric filter | 100ms | 300ms | F2 one-rule scan/mask, end-to-end separately recorded |
| 1M natural text sort | 4,000ms | 10,000ms | F4, same Intl.Collator semantics |
| 100k unique cold value index | 750ms | 2,000ms | complete semantic universe, paged display |
| First usable viewport | 1,500ms | 3,000ms | local seekable F1/F2/F3 package available → labeled usable viewport; progressive readiness |
| Steady-state expanded XML import throughput | ≥20MiB/s | ≥8MiB/s | after inventory/startup, separately from full-source download |
| Steady-state populated cell ingestion | ≥100k cells/s | ≥30k cells/s | F1/F2/F3, includes page persistence/backpressure; report both bytes/cells |
| Main viewer owned resident working memory | 64MiB | 32MiB | source assets/caches/buffers charged, baseline host app excluded explicitly; canvas backing stores are charged to the bitmap pool, not here |
| Worker owned resident working memory | 512MiB | 256MiB | including heap/Wasm/temp/cached pages/transfers |
| Bitmap/GPU retained allocation budget | 128MiB | 64MiB | Canvas2D base/overlay backing stores or GPU textures/buffers, including in-flight replacement tiles (reference Canvas2D layout ≈76MiB D / ≈48MiB M, spec §12.4) |

Import file-size, compressed/inflated-byte and storage quotas are explicit in the specification/resource profile. Throughput and full-import completion grow with bytes; record them separately. Backend memory admission must include fixed source storage overhead and maximum simultaneous operation buffers. Verify tiny-axis coarse mode bounds layout/page/DOM calls and frame work under F6/F7, while exact navigated cells remain accessible. Resource excess triggers clear admission refusal/progress cancellation, not a smaller hidden sort/filter range.

### S10.3 Accessibility/mobile/browser support verification

**Files:** extend browser suite; update `docs/validation/interactive-viewer.md`.

- [ ] Run full keyboard interaction (active navigation, headers, Alt+Down, Enter, Space, Esc, find, multi-selection) in Chrome/Safari and touch sequences on iPad, including edge navigation after recenter.
- [ ] Verify bounded owned grid semantics with desktop screen reader and iPad VoiceOver: role hierarchy, display-slot counts/indices, current active descendant, selected/sorted headers, announcements and focus return. Missing gridcell ownership blocks readiness even if canvas appears correct.
- [ ] Test unavailable Worker, OffscreenCanvas, OPFS/IndexedDB quota denial, font failure and backend/device loss. Small admitted fallback retains correct behavior; unsupported mandatory large mode is visibly diagnosed and cannot claim full capacity/performance.
- [ ] Run protection/event-isolation/zoom/non-XLSX component suites and actual browser scenarios after any fix; expected no hijacked editable input or lost focus, and all supported profiles have documented capability/admission behavior.
- [ ] Deliver support/a11y receipts with known limitations tied to exact acceptance IDs; no unsupported browser receives an unqualified capacity claim.

## S11 — Bounded Rust/Wasm kernel experiment and go/no-go

### S11.1 Interchangeable kernel proof, not a rewrite

**Optional execution boundary:** Record `not-run` with current TS measurements when baseline already meets targets and no measured CPU bottleneck warrants a port. If invoked, timebox to 3–5 engineering days, at most two optimization iterations, numeric sort/filter/metrics plus text-interop parity only; stop/admit against the documented gate, not an expanding rewrite.

**Files:** create `src/xlsx/backends/kernel.ts`, `tsKernel.ts`; experiment `experiments/xlsx-wasm/Cargo.toml`, `src/lib.rs`, `README.md`, `bench.ts`; only if admitted add `src/xlsx/backends/wasmKernel.ts`; tests `tests/xlsx-viewer-kernel-parity.test.ts`.

- [ ] Freeze batch API `prepareColumns` → `sortIndices` / `filterMask` / `valueIndex` / `projectMetrics`, with typed buffers, ownership, handles, cancellation, diagnostics and lossless string dictionary representation. The TS resident implementation remains oracle and default throughout the experiment.
- [ ] RED: identical parity vectors for numeric sort/filter/metrics, blanks/errors/source ties, hidden slots, natural text accented/case keys, lone surrogates and cancellation. Deliberate Rust byte/Unicode sorting or UTF-8 replacement must fail parity.
- [ ] Run `bunx vitest run tests/xlsx-viewer-kernel-parity.test.ts`; expected incompatible kernel fails intended semantics. Future Rust toolchain installation/build is scoped to experiment PR and recorded; do not install during planning.
- [ ] Build bounded Rust/Wasm numeric/index/metric kernels only. Count wasm-bindgen/linear-memory copies, UTF-16 dictionary handling, host Intl.Collator bridge cost, cold module/download/compile/initialization and retained memory. Keep text comparison host-compatible; repeated per-cell FFI calls are rejected in favor of batch operations.
- [ ] Run same controlled fixtures through TS and Wasm using `bun scripts/bench-interactive-viewer.ts --browser chromium --profile desktop --fixture kernel --repeat 30 --mode release --output /tmp/officeview-viewer-kernel-comparison`; repeat mobile and cold paths before admission.
- [ ] Publish `docs/architecture/xlsx-wasm-decision.md`: bounded improvement hypothesis, compatibility results, total bundle/compressed cold cost, memory/copy ledger and actual end-to-end advantage. Go only if all goldens/parity pass, memory/cold budgets remain within S10 and target end-to-end improvement is demonstrated across required profiles; otherwise keep TS and archive prototype clearly.
- [ ] If admitted, add backend loader/version/capability detection and fallback tests, rerun S10 for the shipped backend. Do not make Rust a parser/formula rewrite or claim GPU compute is required.

## S12 — Bounded WebGPU rendering experiment and go/no-go

### S12.1 Canonical-record GPU prototype with Canvas2D fallback

**Optional execution boundary:** Record `not-run` when the compliant Canvas2D baseline has no measured paint bottleneck justifying another backend. If invoked, timebox to 4–7 engineering days and at most two optimization iterations; one viewport pipeline with fills/lines/text/image parity, no parser/formula migration. Any unsupported drawing/font/warp coverage retains reference fallback and cannot silently become default.

**Files:** experiment `experiments/xlsx-webgpu/README.md`, `renderer.ts`, `atlas.ts`, `bench.ts`; potential shipped `src/xlsx/backends/gpuRenderer.ts`, `canvasRenderer.ts`, `renderer.ts`; test `tests/xlsx-viewer-renderer-parity.test.ts`; browser `tests/browser/interactive-viewer-backends.spec.ts`.

- [ ] Freeze renderer consumer contract: canonical cell fill/grid/border/text/image records and clipping, source IDs and logical search/copy records. GPU cannot replace semantics with bitmap-only output. Canvas2D remains reference/fallback.
- [ ] RED: parity cases for ordinary/wrapped/rotated/stacked text, fallback fonts/graphemes, hidden cells, merge clips, borders, drawings/images/warps, watermark and overlays. Force unavailable WebGPU/device loss/atlas exhaustion; fallback must retain view/selection/search and zoom without publishing stale frames.
- [ ] Run `bunx vitest run tests/xlsx-viewer-renderer-parity.test.ts` and `bunx playwright test --config playwright.viewer.config.ts --grep 'backend parity'`; expected an incomplete GPU prototype fails record/pixel/capability cases.
- [ ] Implement bounded WebGPU batches for fills/lines and glyph-atlas quads using canonical host-shaped/layout records. Include atlas upload/eviction, image compositing, device detection/loss, output scaling and resource close. Do not introduce WebGPU compute for sort/filter or require Wasm to render.
- [ ] Measure identical viewports via `bun scripts/bench-interactive-viewer.ts --browser chromium --profile desktop --fixture render --repeat 30 --mode release --output /tmp/officeview-viewer-renderer-comparison`, then actual Safari/iPad. Charge textures/buffers/uploads/CPU record prep and cold shader/device init; microbench GPU draw time alone is insufficient.
- [ ] Publish `docs/architecture/xlsx-gpu-decision.md` with coverage, text/image/warp parity, fallbacks, memory/cold costs and end-to-end frame advantage. Go only if compatibility gates and S10 budgets pass on shipped targets; otherwise retain Canvas2D. Add WebGL2 only after a measured target-browser coverage gap justifies a third backend and its own parity/lifecycle/perf work; do not preselect it from presumed support.
- [ ] If admitted, ship feature-detected fallback and rerun S10/default corpus/backend comparisons. A backend cannot become default before its decision receipt and complete fallback verification.

## S13 — Independent release review, documentation and readiness decision

**Files:** update `README.md`, `PLAN.md`, `docs/validation/interactive-viewer.md`, `interactive-viewer-performance.md`; optional approved decision docs from S11/S12; no new runtime scope.

- [ ] Have an independent reviewer inspect the final spec↔plan↔implementation requirement matrix, then reproduce the header visibility, A/B revision collision, anchored drawing below band, Shift membership, no-owned-gridcell, date midday, >10k Select All and row16385 origin-alias regressions from a fresh checkout.
- [ ] Run final existing gates plus `bunx playwright test --config playwright.viewer.config.ts`; expected all tests/build/corpus comparisons pass. Attach actual Safari/iPad S10 artifacts and source/checkpoint native evidence; UI observations alone do not mark tests verified.
- [ ] Confirm all S0–S10 mandatory boxes complete, S11/S12 statuses explicit (not-run is acceptable when no prototype is justified), no retained stub/fake fixture/timing claim and every accepted backend reran complete release performance. Mark remaining gaps with requirement IDs and hold release if they are mandatory.
- [ ] Document full coordinate capacity vs admitted resource limits, large/small profiles, immutable view-only formula behavior, supported/refused structures, hidden snapshot semantics, filter date/collation rules, copy/print budgets and browser capability fallback. Do not promise instantaneous fully populated 17B-cell load.
- [ ] Verify artifacts' hardware/browser/fixture/version hashes match the release commit, record risk/rollback and feature flag default; keep existing static/export behavior budgeted and honest.
- [ ] Land only reviewed feature branch/PR after required approval/review policy, then update PLAN.md status with commit and receipts. A planning document edit itself is not feature readiness.

## Acceptance-to-delivery matrix

Every stable acceptance case in specification §15 has an owner and evidence below. Implementation PR descriptions cite the exact IDs and attach the corresponding RED/GREEN and independent verification receipts. Families in parentheses are specification requirement IDs, not additional local requirements.

| Acceptance case / requirement IDs | Owning packages | Fixture + exact test/evidence | Release gate |
|---|---|---|---|
| AC-CAP-01 (CAP-01) | S0.1 | F1/F6; `tests/xlsx-viewer-contracts.test.ts`: first/last address, 2^31/2^32/2^34 ID boundaries | mandatory |
| AC-CAP-02 (CAP-02, VIEW-02) | S2.1–S2.2, S6.1 | F1/F7; scroll-window/navigation suites, Go To A2→source A3 vs Original A2→display A3 under [0,2,1], hidden exact inspector/no-unhide, actual edge recenter/zoom/DPR browser scenario | mandatory |
| AC-CAP-03 (CAP-03, STORE-01, LIFE-03) | S1.1–S1.2, S10 | F2/F3/F4 quota variants; import-limits/source-pages suites, typed incomplete/cancel/limit UI | mandatory |
| AC-STORE-01 (STORE-01–03) | S1.2–S1.3 | F8; source/adapter-parity/values suites and native checkpoint, pending≠blank | mandatory |
| AC-STORE-02 (STORE-02, PERF-02) | S1.1–S1.2, S10.2 | F1/F2/F3/F4 malformed variants; import-limits, actual byte/queue/peak/cleanup ledger | mandatory |
| AC-STORE-03 (STORE-04) | S3.1 | F5; structure suite: giant narrow-column merges, headerRows0, saved/dynamic child, complete anchors | mandatory |
| AC-STORE-04 (STORE-03, LIFE-01) | S1.3, S4.2 | F8/F7; values + worker-races suites, valueGen invalidation, incomplete-evaluation admission, `cached` vs `engine` value policy, `formula-value-unavailable` refusal and no-evaluator spy | mandatory |
| AC-VIEW-01 (VIEW-01) | S2.1, S7.2 | F2/F7; projection/order suite with bijection/pinned slots and measured allocations | mandatory |
| AC-VIEW-02 (VIEW-02, RENDER-01) | S2.3, S5.1 | F1/F2; cap-regression/viewport suites: row16385 and max source row sorted to viewport | mandatory |
| AC-VIEW-03 (VIEW-02) | S2.1–S2.2 | F5/F6/F7; axes/scroll suites: Float64, full hidden spans, precision, bounded CSS/canvas | mandatory |
| AC-VIEW-04 (VIEW-02, RENDER-03) | S2.4, S5.1 | F1/F8; sheet-zoom + zoom + cap/legacy suites: absolute 0.25–4 zoom, initial 1.0, `MIN_ZOOM`/`MAX_ZOOM` and page-document gestures unchanged, ordinary unhidden flag-off goldens identical with named S2.3 hidden-correction exceptions and explicit oversized-output failures; small sheet at 0.25 keeps text | mandatory |
| AC-SEM-01 (SEM-02) | S2.1, S8.1 | F2/F5; filter suite: Amount header, hidden body, totals/outside, replacement/Clear | mandatory |
| AC-SEM-02 (SEM-01) | S7.2 | F2/F4/F7; sort truth table/stability/host collator for 1–8 keys, errors/blanks terminal | mandatory |
| AC-SEM-03 (SEM-01) | S7.1–S7.2 | F5/F8; sort-range/integration suites: metadata/no-header/far D outsiderange/formulas intact | mandatory |
| AC-SEM-04 (SEM-01, STORE-04) | S3.1–S3.2, S7–S8 | F5; operation-guard suite: full-column band even for A:B keys, distant blank merge/spill; readonly child selection allowed | mandatory |
| AC-SEM-05 (SEM-01, STORE-04) | S3.2, S8.1 | F5; operation-guard suite: filter0–1 before anchor2, within/below/absolute anchors | mandatory |
| AC-SEM-06 (SEM-03) | S8.1 | F2/F4/F8; filter suite: typed/wildcard/blank/error/compound AND/OR serialization/replay, text-evaluable semantics in mixed columns | mandatory |
| AC-SEM-07 (SEM-03) | S8.1 | F8; date-filter + native/date suites: D+.5, period/day bounds, zones/DST, 1900/1904 | mandatory |
| AC-SEM-08 (SEM-04) | S8.2, S9.1 | F4; value-index/filter-ui: 100k+/1M universe, beyond10k search, all/search-subset selections | mandatory |
| AC-SEM-09 (SEM-02, SEM-04, LIFE-01) | S8.1–S8.2, S9.1 | F4/F7; filter-ui/index/races: last-column Clear→null, Cancel no publication, stale draft rejection | mandatory |
| AC-SEL-01 (VIEW-04) | S6.1 | F7; selection suite: Shift [0,3,1,2] source 3→source 1 = {3,1} and source 0→source 1 = {0,3,1}, later sort membership, explicit display-vs-source Go To, compact full-row/column | mandatory |
| AC-SEL-02 (VIEW-04, A11Y-01) | S6.1–S6.2 | F7; selection/a11y suites: hidden relocation, active+anchor null, no stale descendant | mandatory |
| AC-COPY-01 (SEM-05) | S9.2 | F7/F8; copy/policy suites: visible current order, overlap dedupe, >8MiB atomic refusal, permissions | mandatory |
| AC-A11Y-01 (A11Y-01) | S6.2, S10.3 | F1/F3/F7; a11y owned rows/cells, header-inclusive indices, activedescendant + actual screen-reader/browser receipts | mandatory |
| AC-A11Y-02 (A11Y-02) | S6.2, S9.1, S10.3 | F7; events/filter-ui/sort-ui + actual keyboard/touch/focus receipts | mandatory |
| AC-LIFE-01 (LIFE-01, VIEW-03) | S4.1–S4.2, S5.3 | F7; worker/find-races/zoom suites: A1/B1 local collision, font/value/document owners, atomic view publication/source immutability and zoom retained | mandatory |
| AC-LIFE-02 (LIFE-02, PERF-02) | S4.3, S10.2 | F7 multi-sheet; lifecycle + real retained peak/plateau/close memory/storage ledger | mandatory |
| AC-SEARCH-01 (SEM-05, RENDER-02) | S5.3 | F1/F7/F8; find/browser: never-painted corner/inactive sheet, hidden exclusion, bounded hit pages/lazy geometry | mandatory |
| AC-RENDER-01 (RENDER-01–02) | S5.1–S5.2, S10.1 | F1/F3/F5/F6/F8; layout/viewport/complexity/native suites, overview exact-source-set/cancel/terminal + [0,3,1,2]/filter/stale-view projection tests, full per-frame/per-tile/detail caps, exact goldens, host font/rotation/merge parity | mandatory |
| AC-RENDER-02 (RENDER-03) | S2.3, S5.2, S9.2 | F7/F8; paint counters, no overlay-triggered global rebuild, static print/export typed limit | mandatory |
| AC-PERF-01 (PERF-01–03) | S0.3, S10.2 | F1–F8; actual Chrome/Safari/iPad p50/p95/longtasks/first-view/import/cold/warm result JSON and traces | mandatory |
| AC-PERF-02 (PERF-02–03) | S1–S5, S10.2 | F2/F3/F4/F5/F6/F7; actual peak/retained copies/strings/8-key/temp budgets per pool (main, worker, bitmap/GPU), bounded complexity | mandatory |
| AC-BACKEND-01 (BACKEND-01–02) | S4, S11–S12 if executed | TS resident baseline required; optional prototype parity/UTF16/IntlCollator/cold/copy/memory/fallback receipts or explicit not-run status | baseline/status mandatory; prototype/adoption optional |
| AC-REG-01 (STORE-03, RENDER-03) | each PR, S10.1, S13 | F8/corpus; targeted/full tests/typecheck/build/exact golden comparison and source-evidenced native checkpoints | mandatory |

## Risk register and bounded response

| Risk | Trigger / evidence | Owning response |
|---|---|---|
| Full-object import undermines all later memory work | full worksheet/sharedStrings tree or whole source clone before first view | S1 stops large-mode readiness; implement direct paged import, do not raise RAM budget to conceal it |
| Logical identity aliases through caps/bitwise packing | row16385 at origin or 2^34 ID wraps | S0/S2 boundary tests and full interactive mapping, explicit static budget |
| Unknown pages silently become blanks | incomplete sort/filter/count/value universe commits | S1/S4 authoritative completion state and atomic admission |
| Prefix displacement moves cells but not drawing | drawing below filter band retains old Y | S3 range+prefix impact guard that refuses (no remapping in v1) before any consumer |
| Per-sheet revisions collide globally | A=1/B=1 rebuild reuses prior workbook index | S4 complete identity and generation rejection, S5 aggregate find |
| Natural text/UTF-16 behavior changes in Wasm | byte order/accent/case/lone surrogate parity failures | S11 retain host semantics/oracle; no backend admission |
| GPU preserves pixels but loses text/focus/search semantics | bitmap-only output or unsupported fallback | S5/S12 shared records, Canvas reference, DOM a11y and lifecycle |
| Browser memory/performance metric unavailable or flaky | missing mobile retained-memory/long-task receipts | S10 explicit unsupported fields plus allocation/OS evidence; no fabricated zeros or unit timing gate |
| Storage quota/device capability cannot support admitted operation | OPFS/IndexedDB denied, worker/GPU failure, temp-run exhaustion | typed refusal/cancel and coherent fallback, no truncated 'success' |
| Paged documents cannot recalculate formulas | large workbook with formulas whose caches are missing or stale | S1.3 `cached` value policy (spec §4.4): saved values with provenance, `formula-value-unavailable` refusal limited to key/value domains, visible status; an evaluation bridge is a separate spec amendment, never an implicit v1 fallback |
| Sheet viewer reuses the fit-to-width page zoom model | grid shrunk to fit, `MIN_ZOOM` change leaks into DOCX/PPTX | S2.4 separate sheet scale model behind a flag, page-document zoom constants pinned by tests |
| Canvas backing stores blow the main-thread budget | 76MiB D / 48MiB M of bitmaps charged to a 64/32MiB pool | spec §12.4 dedicated bitmap/GPU pool with worked arithmetic; S10.2 measures every pool |
| Shared integration files collide with other work | simultaneous parse/render/OfficeDoc edits | one integration owner, focused module PRs, rebase and independently rerun affected counterexamples |

This plan's completion condition is a compliant full-coordinate, resource-admitted viewer with verified end-to-end behavior and controlled performance receipts. UI controls alone, model-only 100k timing, synthetic provider geometry alone, or a faster optional kernel cannot satisfy that condition.
