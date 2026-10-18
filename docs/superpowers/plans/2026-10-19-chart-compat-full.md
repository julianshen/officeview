# Chart Compatibility — Plan to Full Support

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development, one implementer at a time, with spec review followed by quality review for each phase. No commits until both reviews approve a phase gate. Frozen candidate + guard per phase (task3-browser-prep/task3-candidate-guard.py style); temp-only diagnostics; do not modify main/inputs/goldens without individual inspection.

**Goal:** Reach full chart compatibility across all document types (XLSX, DOCX, PPTX) — every OOXML chart type the viewer can encounter renders natively with correct geometry, axes, legend, colors, and data fill, instead of placeholders or unsupported diagnostics. ChartEx/model3D/OMML stay as documented fallbacks.

**Baseline facts (audited, do not re-derive):**
- XLSX is the only type with rich native charts (`src/xlsx/chart.ts`): bar/column (clustered/stacked/percentStacked, col+bar), line, area, pie, doughnut, scatter, incl. 3D→2D. Unsupported placeholders: radar, bubble, stock, surface/surface3D, ofPie, combination-pie, unreadable refs, no-series.
- DOCX/PPTX share a generic drawing `chart` kind (`src/drawing/content.ts` `chart()`) that renders **only clustered vertical bar charts** (`barDir=col`, grouping clustered/undefined). Every other embedded chart returns `undefined` → `unsupported-content` diagnostic → coverage `chart` unsupported. ChartEx recognized but not painted in all formats.
- `src/drawing/content-paint.ts` has two painters: generic `paintChart` (bar-only model) and `paintChartPlaceholder` (title + reason). XLSX charts use `xlsxChartPainter` or placeholder.
- Data fill: XLSX resolves uncached refs via worksheet cells (`fillWorkbookCharts`); DOCX/PPTX embedded charts carry cached values only (c:numCache/strCache) — no cell-fill path exists for them.
- Coverage system already tracks `chart`/`ChartEx` features and statuses (`src/drawing/coverage.ts`); diagnostics taxonomy exists.

## Phase 0: Impact audit gate (no source changes)

- [ ] Enumerate every corpus fixture exercising embedded (DOCX/PPTX) and worksheet (XLSX) charts; record which render natively vs placeholder today.
- [ ] Record baseline golden hashes for fixtures that will move; record their absence as accepted risk with root sign-off where no native reference exists.
- [ ] Confirm chart relationship resolution actually loads chart parts in DOCX/PPTX (trace `prepareDrawingContent` → `c:chart` r:id); document any short-circuit.
- [ ] Run full gates (`bunx tsc --noEmit`, `OFFICEVIEW_STRICT_GOLDEN=1 bun run test`, `bun run build`) as baseline. Complete spec then quality review.

## Phase 1: Extend generic `chart()` to all XLSX-supported types

**Files:** Modify `src/drawing/content.ts` (`chart()` + `DrawingContent` `chart` kind), `src/drawing/content-paint.ts` (`paintChart`), `src/drawing/content-model` types; tests in `tests/drawing-content.test.ts` / new `tests/drawing-chart.test.ts`.

- [ ] Test: Parse line, area, pie, doughnut, scatter from cached DrawingML chart (c:lineChart, c:areaChart, c:pieChart, c:doughnutChart, c:scatterChart) into the generic chart model, mirroring XLSX semantics (grouping, barDir, gapWidth, overlap, holeSize, firstSliceAngle, axes, gridlines).
- [ ] Test: Parse stacked/percentStacked/standard groupings and horizontal bars (`barDir=bar`) — not just clustered vertical.
- [ ] Test: Paint each newly-parsed type natively (extend `paintChart` or delegate to the XLSX painter); no placeholder for supported types.
- [ ] Test: Unsupported types (radar, bubble, stock, surface, ofPie) still diagnose `unsupported-chart`/`unsupported-content` and render placeholder — no silent behavior change on reject path.
- [ ] Test: Chart title, legend, axes min/max/majorUnit, gridlines, series colors, category labels parsed from caches (strCache/numCache/strLit/numLit).
- [ ] Run full gates. Complete spec then quality review.

## Phase 2: Chart data fill for embedded charts (DOCX/PPTX)

**Files:** Modify chart resolution + data fill path; reuse XLSX range parsing where possible; new tests.

- [ ] Test: Embedded charts with only references (no cache) resolve values from the document's own cells/tables where a containing sheet/table exists (DOCX tables, PPTX table shapes).
- [ ] Test: When no cell source exists, chart keeps its cached values; if neither cache nor source, placeholder with reason (never blank).
- [ ] Test: Multi-series, scatter xValues, per-point colors (pie varyColors) preserved through fill.
- [ ] Run full gates. Complete spec then quality review.

## Phase 3: ChartEx decision & fallback hardening

**Files:** `src/drawing/coverage.ts`, content.ts, content-paint.ts; tests.

- [ ] Test: ChartEx remains an explicit fallback (`ChartEx` feature) with diagnostic; document in plan limitations whether native ChartEx support is deferred (recommended: deferred — distinct schema, separate phase).
- [ ] Test: Torture file (ChartEx + model3D + OMML + unsupported chart types) renders valid PNG, diagnostics present, zero throws.
- [ ] Test: Coverage reports reflect actual selection (native vs fallback) on owning frame (parity with drawing-coverage tests).
- [ ] Run full gates. Complete spec then quality review.

## Phase 4: Native validation pairs

**Files:** `validation/office-reference/` additions + `tests/office-reference-regressions.test.ts`.

- [ ] Test: Native pair with a spreadsheet containing all supported XLSX chart types.
- [ ] Test: Native pair with a DOCX/PPTX containing embedded charts of each supported generic type.
- [ ] Test: Re-run all pairs after Phases 1–3; record residual diffs with per-unit justification.
- [ ] Complete spec then quality review of the validation set itself.

## Out of scope (deferred, not denied)

ChartEx native rendering, radar/bubble/stock/surface/ofPie (remain placeholders), 3D scene effects, full axis formatting (number formats, tick labels beyond caches), animation. Each retains its current diagnostic/placeholder.

## Acceptance

SPEC + QUALITY approved per phase, strict suite green with only individually-inspected golden updates, tsc/build clean, no commits until both reviews approve a phase gate.
