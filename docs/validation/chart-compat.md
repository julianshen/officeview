# Chart compatibility validation

Status of every promised validation case of the chart-compatibility project
(spec `docs/superpowers/specs/2026-10-20-chart-compat-design.md`, plan
`docs/superpowers/plans/2026-10-20-chart-compat.md`). Created by task F0 as a stub;
later tasks update the status, evidence and provenance of each case. Case IDs are
stable: `CH-<spec section>-<n>`, never renumbered.

Statuses: `pending`, `passing` (synthetic, Office-shaped), `unverified` (interim
non-Office evidence), `matched` (Office reference compared). Synthetic builder
fixtures are "Office-shaped synthetic" and are never reported as Office evidence.

## Final gate table (C1 phase 1, `chart/c1`)

Measured on the C1 branch (base `feature/chart-compat` `10c4f1e`). A7, B3 and B4 are still in review on their own
branches; their rows below are **pending merge** and the controller re-runs every gate after they land (C1 phase 2).

| Gate | Command | Result |
|---|---|---|
| Typecheck | `bunx tsc --noEmit`; `bun run typecheck:scripts` | 0 errors (both) |
| Focused chart tests | `bunx vitest run tests/charts- tests/chart-gallery.test.ts tests/xlsx-chart tests/golden.test.ts` | 38 files, 729 tests passed |
| Full suite | `bunx vitest run` | 431 files: 8,251 passed, 10 skipped, 1 todo, 0 failed (no flakes this run) |
| Eager budget | `bun run build && bun scripts/eager-budget.ts` | 38,639 B (B0 38,631; +8 B; ceiling 38,647, hard 38,720) |
| Eager graph | `OV_MODULE_MAP=1 bun run build && bun scripts/check-eager.ts` | ok (150 eager modules) |
| Goldens | `OFFICEVIEW_STRICT_GOLDEN=1 bunx vitest run tests/golden.test.ts` | 13/13 units identical at threshold 0 (3 existing goldens unchanged; 10 new gallery goldens) |
| Browser (Chromium only: WebKit cannot launch on the build host, ruling at A3b) | `IV_HARNESS_PORT=4300 bunx playwright test --config playwright.viewer.config.ts tests/browser/interactive-viewer-visual.spec.ts tests/browser/chart-embedded-data.spec.ts --project chromium` | 20/20 passed (Chromium; vite on 4300 stopped afterwards) |
| A7 accessibility, B3 histogram/Pareto, B4 box & whisker | (their branches) | pending merge |

## Eager budget

Measured with `bun run build && bun scripts/eager-budget.ts`: the gzip of
`dist/index.js` by `node:zlib` **as implemented by Bun** (the script runs under Bun,
so this is Bun's `node:zlib`, not Node's), Bun 1.3.10, hard ceiling 38,720 B. Every
figure in this table was measured this way; a Node-run gzip may differ by a few bytes.

| Point | Commit | Eager gzip (B) | Note |
|---|---|---|---|
| **B0** | `ea03a34` (`chart/f0` before F0; no src/scripts/build-config change since `3581768`) | **38,631** | Every later task must stay at or below B0 unless earlier tasks freed bytes |
| After F0 | F0 commit | 38,632 | `dist/index.js` identical to B0 except six content-hashed chunk filenames; +1 B is hash entropy (see F0 report) |
| After A0a | A0a commit | 38,635 | `dist/index.js` byte-identical to the `f4563ba` build after normalising chunk-hash names (B0 + 4 B hash churn, within the 16 B tolerance; the lazy engine chunk is now `charts-*.js`). The deleted mini-reader and `paintChart` lived in the statically imported `office-worker` chunk, which dropped 350,840 → 349,510 B gzip (−1,330 B); the eager closure (index + importSession + office-worker) went 390,631 → 389,304 B. Running total vs B0: +4 B on `dist/index.js` (hash churn only) |
| After A0b | A0b commit | 38,638 | Before (A0a, rebuilt): 38,635. The `preloadCharts` export adds its one import binding and export name to `dist/index.js` (+13 B, measured 38,648 = B0 + 17, over B0 + 16). Recovered without behaviour change by deleting one comment that `dist/index.js` preserves verbatim (the restating `/** Decode and emit any buffered \'hh bytes. */` on `RtfParser.flushHex`, −10 B). `ensureChartEngine`, the painter containment and the representation contract live in the `office-worker` chunk, not `dist/index.js`. Bytes added by A0b: +3 B. Running total vs B0: +7 B (within the 16 B hash-churn tolerance) |
| After A0b pre-review fix | A0b fix commit | 38,637 | DOCX no-fallback card and DOCX/PPTX broken-fallback card; no `dist/index.js` code change (−1 B hash churn). Running total vs B0: +6 B |
| After A0b fix round 1 | A0b r1 commit | 38,637 | Bounded admitted chart reader (`src/charts/read.ts`), namespace dispatch, chunk/decode-failure cards, whole-fallback evaluation, failure propagation and balanced engine saves: all outside `dist/index.js` (lazy `charts-*` chunk or the `office-worker` chunk). `dist/index.js` unchanged. Running total vs B0: +6 B |
| After A0b fix round 2 | A0b r2 commit | 38,639 | Exhausted-wrapper propagation, persistent priced namespace scopes, streamed comment discard and token-limit classification; outside `dist/index.js` (+2 B hash churn only). Running total vs B0: +8 B |
| After A0b fix round 3 | A0b r3 commit | 38,635 | Exhaustion from the fallback's actual resolution (PPTX, XLSX); outside `dist/index.js` (hash churn only). Running total vs B0: +4 B |
| After A0b fix round 4 | A0b r4 commit | 38,637 | PPTX fallback exhaustion from remaining usable content; outside `dist/index.js` (hash churn only). Running total vs B0: +6 B |
| After D0 | D0 commit | 38,637 | Before (fdd1117, rebuilt): 38,637. The point-state model, analytic acquisition, display model, paged parent ledger and largeSource context reuse live in the lazy `charts-*`, `office-worker`, `largeSource-*` and paged-drawings chunks. `dist/index.js` is byte-identical to the base build once chunk hashes are normalised (intermediate builds measured 38,639: hash churn only). Bytes added by D0: 0. Running total vs B0: +6 B |
| After A1a | A1a commit | 38,638 | Before (`22576dd`, rebuilt from a clean `git archive`): 38,635. The axis graph (`src/charts/axes.ts`), dates (`src/charts/dates.ts`), axis parsing and admission, date-role fill and the graph-driven painter live in the lazy `charts-*` chunk; the paged fill's epoch option in the `office-worker` chunk. `dist/index.js` is byte-identical to the base build once chunk hashes are normalised (intermediate build 38,637). Bytes added by A1a: 0 (hash churn only). Running total vs B0: +7 B |
| After A1a fix round 1 | A1a r1 commit | 38,637 | Workbook-epoch stamping, text category source type, admitted graph construction, screen-order date labels, minor x/date elements and overflow-safe normalisation: all in the lazy `charts-*` chunk and the `office-worker` chunk. `dist/index.js` byte-identical to the `22576dd` build after normalising chunk hashes. Running total vs B0: +6 B |
| After A1a fix round 2 | A1a r2 commit | 38,636 | Streamed line/area/scatter construction, receiving-document chart contexts (`receiveCharts`, `PaintOptions.chartContext`/`chartProfile`, `ContentPaintAssets.chartContext`) and the finite-scale fallback: lazy `charts-*`, `office-worker` and `InteractiveSheetDoc` chunks. `dist/index.js` byte-identical to the `22576dd` build after normalising chunk hashes. Running total vs B0: +5 B |
| After A1a fix round 3 | A1a r3 commit | 38,639 | The interactive sheet preparation forwards the chart context and parents the receiving chart facet to the viewer's main-thread ledger (`mainChartParent`): `InteractiveSheetDoc`, `office-worker` and `charts-*` chunks only. `dist/index.js` byte-identical to the `22576dd` build after normalising chunk hashes (hash churn only). Running total vs B0: +8 B |
| After A1a fix round 4 | A1a r4 commit | 38,637 | Receiving chart contexts per (document, parent, profile), held by preparation/asset leases and carried per paint (`acquireChartContext`, `ContentPaintAssets.chartContext`): `charts-*`, `office-worker` and `InteractiveSheetDoc` chunks only. `dist/index.js` byte-identical to the `22576dd` build after normalising chunk hashes. Running total vs B0: +6 B |
| After A1b | `634dfd4` | 38,639 | Placement-keyed layout sidecar and transferable stage-1 snapshot (lazy `charts-*`); base `416117d` builds 38,637, index.js byte-identical after normalising chunk hashes (ledger). Running total vs B0: +8 B |
| After A1b fix round 1 | `7c91179` | 38,638 | Pressure-evicted layout LRU, worst-case charge fix (lazy chunk only; ledger). Running total vs B0: +7 B |
| After A2 | `356c19f`, fixes `cc02f2d`, `36f7eb7` | not recorded | The A2 ledger entries carry no eager figure; the next measured point (A3a, built on the A2 merge) is 38,641 before its fixes. Lazy `charts-*`/`cellFormat` only |
| After A3a | `51c026d` | 38,641 | Bounded embedded-workbook access (`src/charts/embeddedZip.ts`, lazy; nothing added to `src/core/zip.ts`). Running total vs B0: +10 B |
| After A3a fix round 1 | `bfc343b` | 38,638 | Bounded `.rels` read and abort handling (lazy). Running total vs B0: +7 B |
| After A3a fix round 2 | `db3be11` | 38,636 | Narrow relationship grammar (lazy). Running total vs B0: +5 B |
| After A3b | `a63b6e6`, `484e88b` | 38,637 | Values-only embedded workbook reader (lazy). Running total vs B0: +6 B |
| After A3b fix round 2 | `3c146a5` | 38,639 | Batch discovery after an uncovered request (lazy; hash churn). Running total vs B0: +8 B |
| After A3b fix round 3 | `854a84b` | 38,637 | Concurrent waiters share one batch (lazy). Running total vs B0: +6 B |
| After A4a fix round 1 | `c7fe3de`, `9648201` | 38,637 | Colour chain (spPr, style part, colour part, c:style, theme) in the lazy chunk; the first (agy) report claimed 38,639. Running total vs B0: +6 B |
| After A4a fix round 2 | `bf6fdb8` | 38,636 | Gradient stops, alpha, marker noFill (lazy). Running total vs B0: +5 B |
| After A4a fix round 4 | `654e46a` | 38,639 | Variation-expanded total colour set (lazy; hash churn). Combined gates at `f04481f`: 38,639. Running total vs B0: +8 B |
| After A4b | A4b commit | 38,640 | Before (`f04481f`, rebuilt from a clean `git archive`): 38,636. Title runs, legend overlay/legendEntry, multi-level category tiers (`src/charts/titles.ts`, `src/charts/categories.ts`, layout, paint, parse, style part) are all in the lazy `charts-*` chunk. `dist/index.js` byte-identical to the base build after normalising chunk hashes. Bytes added by A4b: 0 (hash churn only). Running total vs B0: +9 B |
| After A4c / A4d (combined gates after the A4d merge) | `12559dc`, `6fcc4e6` | 38,636 | Smoothing (A4c) and user shapes (A4d), both lazy. Running total vs B0: +5 B |
| After A5a | `f82ad82`, `2ce5a24` | 38,639 | Bubble and radar families, conformance harness (lazy; implementer reported ≤ 38,639). Running total vs B0: +8 B |
| Combined gates at `8be9dce` (A4d + A4c + A5a) | `8be9dce` | 38,638 | tsc 0, vitest 7,978/0. Running total vs B0: +7 B |
| After A5c | `b5587ff` | 38,639 | ofPie family (lazy). Fix round 2 + feature merge `53d7897`/`0ee83cf`: 38,639. Running total vs B0: +8 B |
| After A5d fix round 2 | `3388b90` | 38,637 | Contour and wireframe contour (lazy). Running total vs B0: +6 B |
| Combined gates at `5578e55` (A5d) | `5578e55` | 38,638 | A5b stock merged after this point without a separate figure (lazy only). Running total vs B0: +7 B |
| After A6a | `b2287e7` | 38,636 | Trendlines (lazy). Merged tree `abec1ae` gated 38,637. Running total vs B0: +5 B |
| After A6b | `fa6f60f` | 38,635 | Error bars (lazy). Merged tree `b71a970` gated 38,636. Running total vs B0: +4 B |
| After B1 parser | `2fbeb7f` | 38,639 | ChartEx parser and field map (lazy). Running total vs B0: +8 B |
| Combined gates at `7c67e33` (B1) | `7c67e33` | 38,640 | tsc 0, vitest 8,098/0. Running total vs B0: +9 B |
| Combined gates after A4b + A5c | feature | 38,639 | vitest 8,144/0. Running total vs B0: +8 B |
| Combined gates after A6b + B2 | feature | 38,635 | vitest 8,186/0. Running total vs B0: +4 B |
| After B2 | B2 commit | 38,636 | Before (`7c67e33` tree plus B2's type-only `chartEx.palette` field, rebuilt): 38,640. Waterfall and funnel families (`src/charts/chartex/waterfall.ts`, `funnel.ts`) and their hooks in axes, display, prepare, layout, labels and paint live in the lazy `charts-*` chunk; `dist/index.js` changes by hash churn only. Bytes added by B2: −4 (hash churn). Running total vs B0: +5 B. Ceiling 38,647 |
| After B2 fix round 1 | B2 r1 commit | 38,635 | Subtotal state gate, finite funnel ratio, hidden series out of the display view, funnel label pricing: lazy `charts-*` chunk only (hash churn). Running total vs B0: +4 B |
| After B2 merge with feature | merge commit | 38,635 | `feature/chart-compat` `cb5a0b5` merged into `chart/b2`; conflicts resolved keeping both sides (lazy `charts-*` chunk only). Running total vs B0: +4 B. Ceiling 38,647 |
| After B5 | B5 commit | 38,639 | Before (`7c67e33`, rebuilt): 38,640. Treemap and sunburst (`src/charts/chartex/hierarchy.ts`, `hierarchyStyle.ts`, `treemap.ts`, `sunburst.ts`) and the hierarchy hooks in fill, precedence, embedded reader, stage 1, layout and paint all live in the lazy `charts-*` chunk; the two description names in the `imageAdmission` chunk. No eager source file was touched; measured −1 B (hash churn, not normalised). Bytes added by B5: 0. |
| After B5 merge with feature | merge commit | 38,637 | `feature/chart-compat` merged into `chart/b5` (B2, A6b, A5c, A4b and the rest); conflicts resolved keeping both sides; B5 review minors (label numFmt, exclusion order) applied. Lazy `charts-*` chunk only. Before (feature ledger, B2 merge): 38,635; +2 B hash churn. Ceiling 38,647 |
| Task B5 merged (feature `10c4f1e`) | `10c4f1e` | 38,637 | Gated at merge: full 8,218/0. Running total vs B0: +6 B |
| After A7 | `chart/a7` | pending merge | Accessibility and paged chart list (implementer reported 38,639 at `88bbdfe`); the controller records the merged figure in phase 2 |
| After B3 | `chart/b3` | pending merge | Histogram and Pareto (implementer reported 38,635 at `f2c8ee3`); controller records the merged figure |
| After B4 | `chart/b4` | pending merge | Box & whisker; controller records the merged figure |
| **After C1 (phase 1)** | `chart/c1` | 38,639 | Robustness slice, gallery, README/CHANGELOG, validation doc. Source changes are all lazy (`charts-*`: stock charge name, ofPie sweep, hierarchy legend bound, XLSX external-reference diagnostic, exported family registry); `src/testdata/chart-gallery.ts` is test data, never bundled. +2 B vs the `10c4f1e` figure (hash churn). Ceiling 38,647 (B0 + 16). Running total vs B0: +8 B |

## Spec §8 testing and validation

One case per §8 bullet (leaf level: sub-bullets of Goldens, References and Browser tests are separate cases).

**Reference status (C1).** `Reference` says whether the case is compared with an Office-authored reference
rendering: **verified** or **unverified**. No owner-supplied Office reference exists yet (CH-8-8), so every case is
**unverified**; none is claimed Office-matched. The ChartEx inputs in `tests/fixtures/chartex/` are Office-authored
files (provenance and MPL-2.0 licence in their PROVENANCE.md), but no Office *output* for them is compared, so they
are input evidence only. `Status` is the test evidence: `passing` (Office-shaped synthetic tests pass),
`pending merge` (the evidence lives on a branch still in review), `pending` (no evidence yet).

| ID | Case | Status | Reference | Evidence |
|---|---|---|---|---|
| CH-8-1 | Numeric oracles: hand-computed values for ticks, date axes (both systems), display units, number-format classes, ofPie splits, stock roles, bubble scaling, trendlines, error bars, histogram bins, Pareto, box-whisker quartiles, waterfall subtotals, treemap/sunburst aggregation | passing (histogram bins, Pareto and box-whisker quartiles: pending merge, B3/B4) | unverified | A1a (ticks, date axes in both systems, display units; Office-shaped synthetic, not Office-matched): `tests/charts-axes.test.ts` (scales, crossings, log, dispUnits, tick bound) and `tests/charts-dates.test.ts` (epoch precedence, civil-model conversion of serials 0/1/59/60/60.5/61 both ways, calendar month/year steps, unequal spacing). Default unit labels other than hundreds/thousands/millions/billions/trillions are unverified; A5c ofPie splits (`charts-families-ofpie`), A5a bubble scaling (`charts-families-bubble-radar`), A5b stock roles (C1: exact painted coordinates for HLC, OHLC, volume-HLC and volume-OHLC in `charts-families-stock`), A2 number-format classes (`charts-format`), A6a every trendline type (`charts-trendlines`), A6b error bars (`charts-error-bars`), B2 waterfall subtotals (`charts-chartex-waterfall-funnel`), B5 treemap/sunburst aggregation and squarified/angle oracles (`charts-chartex-treemap`) |
| CH-8-2 | Unit tests: parse to model, and paint with a recording context, for every family and element | passing (B3/B4 families pending merge) | unverified | Family conformance harness (`tests/helpers/chart-family-conformance.ts`: empty, all-blank, NaN/∞, negative, single point, unavailable point, each through the adapter, stage 1, layout and paint) registered by every family suite; per-family recording-paint suites `charts-families-*`, `charts-chartex-*`, element suites `charts-axes`, `-dates`, `-format`, `-labels*`, `-style*`, `-smooth`, `-trendlines`, `-error-bars`, `-text`, `-usershapes`, `-layout` |
| CH-8-3 | Representation tests: spec 6.3 in all three formats, including no-fallback and chunk failure | passing | unverified | A0b: `tests/charts-representation.test.ts` (Office-shaped synthetic): MC with injected unsupported/refused and surface3DChart, ChartEx with fallback, nested/alias/unknown Requires and the standalone card in all three formats; no-fallback card and broken fallback picture (all three formats); chunk failure (DOCX, PPTX) |
| CH-8-4 | Data tests: spec 6.4 (embedded workbook, precedence, budgets, cancellation) | passing | unverified | D0 (XLSX clause only; Office-shaped synthetic): `tests/charts-data-contract.test.ts` resident vs paged point arrays (sparse caches, references, blanks, #N/A vs #DIV/0!, formulas without saved values), analytic acquisition (65,536 / refusal), plotVisOnly provenance, paged parent-ledger ownership and largeSource single workspace. The embedded workbook is a later task; A3a/A3b: `charts-embedded-zip`, `charts-embedded-workbook`, `charts-precedence` (inflated and declared caps, inventory, one read per document, cancellation, values-only reader, external links); C1: torture classes (b) and (c) in `charts-robustness` at document level in all three formats |
| CH-8-5 | Robustness tests: extreme scales, tiny units, huge forecasts, 200 charts per document, a torture file per format, asserting allocation/work counters | passing | unverified | A1a: tiny majorUnit/minorUnit and a 0..1e300 scale stay within 1,000 ticks per axis with `chart-tick-limit` (`tests/charts-axes.test.ts`); A6a forecasts clamp (`chart-trendline-extension-clamped`); C1 `tests/charts-robustness.test.ts`: 200 charts per document per format within the 32 MiB mobile facet (11 native, 189 memory cards, peak ≤ limit, fully refunded), torture documents per format by damage class (a)–(d), layout churn, and the path (exact admission costs 1/2/3 per point), label-byte (4 MiB, 255 units) and contour-grid (64 × 256) bounds |
| CH-8-6 | Goldens: synthetic chart gallery (one chart per family and major element), per format, recorded on the Mac | passing (histogram, Pareto, box & whisker: pending merge) | unverified | C1: `src/testdata/chart-gallery.ts`, goldens `tests/goldens/chart-gallery-{docx.0-3,pptx.0-4,xlsx.0}.png` (20 charts per format, each PNG inspected; contents listed under "C1 gallery" below); `tests/chart-gallery.test.ts` checks in CI that every gallery chart is native |
| CH-8-7 | Goldens: existing XLSX chart goldens pixel-identical, except listed spec 5 deltas with reasons | passing | unverified | D0: strict goldens 3/3, `tests/goldens` untouched (no golden contains a chart). Listed delta (spec 5.8 `plotVisOnly`, Office default on): a reference-filled point in a hidden row or column is no longer drawn, and a series whose every point is hidden is not drawn (`tests/charts-data-contract.test.ts` pixel check). Non-value point states still draw as gaps (the matrix is A2). A1a: the painter draws on the axis graph; every existing chart fixture repaints pixel-identically (hashes of all 41 chart paints in the chart/drawing/docx/pptx/golden suites; of 48 synthetic variant renders, all identical except the two deleted-category-axis renders, which lose the axis line row: delta 1 below). The strict golden suite contains no chart and is not independent evidence of chart parity. Listed deltas (spec 5.1): a deleted category axis no longer draws its axis line; a group on a secondary value axis is scaled by that axis; a deleted scatter x axis still draws its authored gridlines (all three apply only to charts that author them). C1: the three pre-existing goldens (`docx-report`, `xlsx-budget`, `pptx-slides`) compare identical under strict mode; no golden re-recorded. C1 deltas: a pie-of-pie secondary pie is now a full circle (it swept share fractions as radians, a ~57° fan), and (phase 1.5) the primary pie is turned so its Other slice faces the secondary plot; no pre-existing golden contains an ofPie chart |
| CH-8-8 | References: Office-authored files and screenshots in validation/office-reference/ with producer and version, one per promised semantic case | pending | unverified | The owner supplies them; none yet |
| CH-8-9 | References: interim python-docx/python-pptx/openpyxl + LibreOffice evidence marked unverified; Office-matched not claimed | passing | unverified | Every task section of this document marks its unverified rules; the ChartEx inputs are Office-authored LibreOffice test files (13 parts, all 8 families) with provenance; no LibreOffice render is used as a reference |
| CH-8-10 | References: smoothing claimed matched only after reference comparison | passing (not claimed) | unverified | A4c: smoothing implemented via Catmull-Rom cubic Bézier (tension ½, end point tangents), marked approximation and unverified until Office reference comparison (including absent c:smooth default on c:ser vs group). C1 tightened the span oracle to exact control points (Catmull-Rom, tension ½) |
| CH-8-11 | Browser: DOCX/PPTX charts render in the paged viewer | passing (Chromium only) | unverified | A3b `tests/browser/chart-embedded-data.spec.ts` (Chromium; WebKit unavailable on the host, ruling at A3b); A7's paged chart list pending merge |
| CH-8-12 | Browser: dark mode leaves chart pixels identical to light | pending merge (A7) | unverified | |
| CH-8-13 | Browser: the per-page chart list exposes accessible names | pending merge (A7) | unverified | |
| CH-8-14 | Browser: CLS is 0 when a chart loads | pending merge (A7) | unverified | A7 notes a pre-existing paged-chrome CLS shift that is not chart-caused (follow-up) |

## Spec §5.1 compatibility policies and unverified defaults (A1a)

Each row is implemented as stated but has no Office reference yet; none is claimed Office-matched.

| ID | Behaviour | Disposition | Status | Evidence |
|---|---|---|---|---|
| CH-5.1-1 | `crossBetween` absent: the legacy per-family geometry (bar and line between ticks, area on ticks), not the spec's midCat line default | Compatibility amendment (controller ruling A1a Q1; spec §5.1 amended in the feature branch) | unverified | `tests/charts-axes.test.ts` crossBetween; synthetic pixel identity only |
| CH-5.1-2 | `tickLblPos` absent: labels at the near edge (category labels at the low edge even when the axis crosses mid-plot); `crosses=max` or a reversed crossing axis takes the far edge | Deferral (controller ruling A1a Q2); the schema default is nextTo | unverified | painter `labelPlace` |
| CH-5.1-3 | A plot area without axis elements (or a group without axIds) binds an implied primary pair instead of a card | Recovery policy (controller ruling A1a Q3), not Office evidence | unverified | `tests/charts-axes.test.ts` unplaceable/implied pair |
| CH-5.1-4 | Default display-unit labels: Hundreds, Thousands, Millions, Billions, Trillions; "Ten Thousands", "Hundred Thousands", "Ten Millions", "Hundred Millions"; the custom-unit label `x<n>` | Default text pending an Office reference | unverified | `tests/charts-axes.test.ts` display units |
| CH-5.1-5 | Automatic date base unit (years when every date shares month and day, months when they share the day, else days) | Inferred rule pending an Office reference | unverified | `tests/charts-dates.test.ts` |
| CH-5.1-6 | Automatic date major unit: one base unit, labels thinned so they never overlap | Inferred rule pending an Office reference | unverified | `tests/charts-axes.test.ts` I4 |

## Spec §5.8 plotting matrix

One case per matrix cell (Office default: `dispNaAsBlank` absent or false). Each is a
numeric oracle and needs its Office reference before it is claimed matched (C1: every cell has passing synthetic evidence; none is verified against Office); standard
area blank/#N/A cells are checked under each `dispBlanksAs` value (gap, zero, span).

| ID | Family | Point state | Specified behaviour | Status | Reference | Evidence |
|---|---|---|---|---|---|---|
| CH-5.8-1 | Line, scatter, radar | blank | per dispBlanksAs (gap: break; zero: 0; span: connect) | passing | unverified | `charts-matrix` (line, scatter: `plotPoint` all three dispBlanksAs modes, native stroke oracles), `charts-families-bubble-radar` (radar, all modes, unavailable diagnostic) |
| CH-5.8-2 | Line, scatter, radar | #N/A | no marker; line connects across it whatever dispBlanksAs says | passing | unverified | `charts-matrix` (line, scatter: `plotPoint` all three dispBlanksAs modes, native stroke oracles), `charts-families-bubble-radar` (radar, all modes, unavailable diagnostic) |
| CH-5.8-3 | Line, scatter, radar | other error | as #N/A | passing | unverified | `charts-matrix` (line, scatter: `plotPoint` all three dispBlanksAs modes, native stroke oracles), `charts-families-bubble-radar` (radar, all modes, unavailable diagnostic) |
| CH-5.8-4 | Line, scatter, radar | unavailable | unconditional gap, plus chart-point-unavailable | passing | unverified | `charts-matrix` (line, scatter: `plotPoint` all three dispBlanksAs modes, native stroke oracles), `charts-families-bubble-radar` (radar, all modes, unavailable diagnostic) |
| CH-5.8-5 | Column, bar | blank | no bar | passing | unverified | `charts-matrix` (bar: no bar for every non-value state, all modes); unavailable diagnostic: `charts-matrix` "unavailable plot inputs publish the required diagnostic" |
| CH-5.8-6 | Column, bar | #N/A | no bar | passing | unverified | `charts-matrix` (bar: no bar for every non-value state, all modes); unavailable diagnostic: `charts-matrix` "unavailable plot inputs publish the required diagnostic" |
| CH-5.8-7 | Column, bar | other error | no bar | passing | unverified | `charts-matrix` (bar: no bar for every non-value state, all modes); unavailable diagnostic: `charts-matrix` "unavailable plot inputs publish the required diagnostic" |
| CH-5.8-8 | Column, bar | unavailable | no bar, plus diagnostic | passing | unverified | `charts-matrix` (bar: no bar for every non-value state, all modes); unavailable diagnostic: `charts-matrix` "unavailable plot inputs publish the required diagnostic" |
| CH-5.8-9 | Area (standard) | blank | gap if dispBlanksAs=gap, else 0 | passing | unverified | `charts-matrix` (area: gap under gap, else 0); unavailable diagnostic as above |
| CH-5.8-10 | Area (standard) | #N/A | gap if dispBlanksAs=gap, else 0 | passing | unverified | `charts-matrix` (area: gap under gap, else 0); unavailable diagnostic as above |
| CH-5.8-11 | Area (standard) | other error | as #N/A | passing | unverified | `charts-matrix` (area: gap under gap, else 0); unavailable diagnostic as above |
| CH-5.8-12 | Area (standard) | unavailable | as blank, plus diagnostic | passing | unverified | `charts-matrix` (area: gap under gap, else 0); unavailable diagnostic as above |
| CH-5.8-13 | Stacked or 100% (any family) | blank | 0 | passing | unverified | `charts-matrix` "stacked and 100% override every non-value state" (bar, line, area) |
| CH-5.8-14 | Stacked or 100% (any family) | #N/A | 0 | passing | unverified | `charts-matrix` "stacked and 100% override every non-value state" (bar, line, area) |
| CH-5.8-15 | Stacked or 100% (any family) | other error | 0 | passing | unverified | `charts-matrix` "stacked and 100% override every non-value state" (bar, line, area) |
| CH-5.8-16 | Stacked or 100% (any family) | unavailable | 0, plus diagnostic | passing | unverified | `charts-matrix` "stacked and 100% override every non-value state" (bar, line, area) |
| CH-5.8-17 | Pie, doughnut | blank | no slice | passing | unverified | `charts-matrix` (pie, doughnut: no slice for every non-value state) |
| CH-5.8-18 | Pie, doughnut | #N/A | no slice | passing | unverified | `charts-matrix` (pie, doughnut: no slice for every non-value state) |
| CH-5.8-19 | Pie, doughnut | other error | no slice | passing | unverified | `charts-matrix` (pie, doughnut: no slice for every non-value state) |
| CH-5.8-20 | Pie, doughnut | unavailable | no slice, plus diagnostic | passing | unverified | `charts-matrix` (pie, doughnut: no slice for every non-value state) |

## A1b placement and typography validation

A1b adds a transferable stage-1 axis/tick/label snapshot and a realm-local,
placement-keyed layout sidecar. Retained layouts belong to the receiving chart
workspace; paint scratch is released on return. Font publication, model replacement and receiver disposal invalidate the sidecar.
Multiple sizes share a pressure-evicted LRU; revisiting a size reuses its placement. An active paint pins
its invalidated entry until it returns. The synchronous API without an explicit
context retains its existing bounded per-call default. Unowned stage-1 preparation
uses an isolated mobile workspace: it is bounded locally but is not aggregated
with a receiving document ledger. Callers needing aggregate admission must supply
an owner/context.

`tests/charts-layout.test.ts` covers placement keys, warm measurement reuse,
refusal before measurement, transfer, receiver/model disposal, reentrant disposal,
font publication and registered family resolution, 6–72 pt bounds, axis rotation,
label thinning, separate tickMarkSkip, manual edge/factor/default/inner layout,
refill ownership and the path work cap. Existing accounting tests keep their
point/string byte assertions after deducting the separately owned preparation
snapshot. These are synthetic evidence, **not Office-matched references**.

The deleted A1a 41-paint scratch harness could not be reused. A reconstructed
node-canvas harness captured 38 paints / 15 distinct cases from the existing
chart, drawing and document suites. Fourteen matched initially; a 100×100 chart
exposed unnecessary value-label thinning. After its regression test/fix, its PNG
also matches exactly (SHA256
`9535280548247157ebd54e205ed5f7a6e0065b6ec315af252ab9a3cb3e25718d`).
This is narrower coverage than A1a's original 41 paints. No golden was re-recorded.
Intentional changes are authored text size/rotation and manual placement,
nonoverlapping label thinning, tickMarkSkip with hidden labels, and dropping later
series when the path work bound would be exceeded.

A clean archive of base `416117d` builds an eager bundle of 38,637 bytes; the
candidate build measured 38,639 bytes with bun/node:zlib. Normalizing lazy chunk
hashes makes index.js byte-identical: +2 B from the base, +8 B from B0 (38,631),
below the task ceiling of 38,647. That measurement belongs to the initial A1b implementation. Browser and
independent Office-reference validation remain pending; no acceptance box is
advanced on the strength of the node-canvas checks.

Fix-round-1 regressions additionally cover 200 native mobile charts, tight retained
charges, workspace LRU pressure, one-byte construction refusal, A→B→A reuse,
a fresh VM realm with its own receiving workspace, collection of the original
painting canvas, uniform automatic label intervals, authored skip and vertical
text, shared display-unit bands, worker font publication, and refused snapshot
release. These tests establish accounting and placement invariants, not independent
Office-reference fidelity.

Fix-round-1 gates: typecheck, 161 chart tests, mapped production build and eager
import graph pass; eager gzip is 38,638 B (task ceiling 38,647 B). The full suite
with the three socket suites excluded has 7,638 passing and five failing tests.
Three failures pass isolated; the policy eviction/recreation timeout and mounted
font-generation asset-loader assertion reproduce on unchanged base 634dfd4.
The full-suite gate therefore remains non-green.

## A2 display formats, data labels and point-state matrix

A2 uses the lazy `cellFormat` display layer for prepared axis and data-label
strings, with a chart-workspace allocation observer, a 255-unit format-code
limit and a 256-unit output cap. Explicit formats override source formats;
sourceLinked retains cache and resident/paged cell provenance per point. The
mixed-source axis rule selects the first nonblank point of the first bound
series. This rule remains **unverified against Office**.

`charts-format`, `charts-labels` and `charts-matrix` are Office-shaped synthetic
unit/native-paint checks, not Office-authored reference evidence. They cover
format classes and epochs, inherited/point label controls, rich runs, separator,
legend keys and leader lines, unsupported field/manual-position fallback, the
non-radar point-state matrix, unavailable diagnostics, and line/scatter stroke
oracles. The data-contract suite pins resident/paged source-format parity.

A scratch comparison against A2 base `32d4f50` exercised 34 paints / 12 distinct
cases. Eleven cases remain pixel-identical. The secondary-axis combination
case changes `1,000` to `1000` under General; the narrower measured label also
reduces its axis band and expands the plot. Both images and label traces were
inspected. This is narrower than the deleted A1a 41-paint harness. No golden was
re-recorded. Intentional deltas are display-formatted ticks, new data labels,
and the specified blank/error gap/zero/span behavior. Pie bestFit uses a bounded
small-slice outside-placement heuristic; independent Office matching of label
placement remains pending. No acceptance checkbox is advanced by these tests.

## A4a colour precedence: implemented and unverified behaviour (spec 5.11)

- **Implemented.** Series colours (and varyColors point colours) follow spPr, then the chart style part (`cs:chartStyle` `dataPoint`/`dataPointLine`), then the colour part (`cs:colorStyle`), then `c:style`, then the theme accent cycle. The cycle uses the variation table of Office's default colour part (colors1.xml, `id="10"`) for every cycle after the sixth series.
- **Colour part methods.** MS-ODRAWXML §2.8.4.2 fixes how colours are chosen and in what order, and the engine follows it:
  - both methods select from the **total set**, the part's colours expanded by its variations (§2.8.3.2: colour fastest, then variation); the set is indexed, never materialized;
  - `withinLinear` gives every object the **first** colour of the total set, which includes its first variation;
  - `acrossLinear` picks the colour by index modulo the size of the total set;
  - transforms apply in this order: the colour's own, then the variation, then the brightness ramp;
  - in both, brightness progresses with index/count; the `…Reversed` methods run that progression backwards;
  - `cycle`, and any unknown method (§2.8.4.1), cycles through the colours, applying the part's variations.

  The spec gives no brightness formula. The engine uses a symmetric lumMod/lumOff ramp from −50% (darker) to +50% (lighter) across the objects (`LINEAR_SPREAD`, `linearBrightness` in `src/charts/style.ts`). This ramp is an **approximation** and is **unverified against Office**.
- **`c:style` single-colour columns: unverified.** Greyscale (column 0) and single-accent (columns 2-7) styles spread a shade or tint over (-70%, 70%) across the series. This follows LibreOffice's `oox` import (`DetailFormatterBase::getPhColor`), not Office output.
- **Explicit `c:symbol val="auto"`: unverified.** The marker cycle order (diamond, square, triangle, x, star, circle, plus, dot, dash) is recalled, not measured. A series without `c:marker` keeps the engine's default circle.
- **Inverted negative bars: unverified.** Without `c14:invertSolidFillFmt`, an inverted negative is painted white.
- **Admission.** A style or colour part, or its relationships, that cannot be admitted refuses the chart (spec 6.4). A missing, malformed or unsupported part falls to the next precedence level.
- **Evidence.** All fixtures are Office-shaped synthetic. None is Office evidence.

## A4b titles, legends and multi-level categories (spec 5.9, 5.10)

- **Title content.** `title/tx/rich`, else the `title/tx/strRef` cache, else the automatic title (a single series' name, else "Chart Title"; 圖表標題 for a `zh` chart), never when `autoTitleDeleted` is set. As before A4b, the automatic title needs a `c:title` element; an uncached `strRef` title falls to the automatic title (§5.10 names the cache; no worksheet fill of title references).
- **Run formatting.** Each attribute (size, bold, italic, colour, typeface, underline) comes from the first level that states it: run `a:rPr`, the paragraph's `a:pPr/a:defRPr`, the title's `c:txPr`, the chart style part's `cs:title` (`cs:defRPr`, its `a:latin`, and the `cs:fontRef` colour and `idx` major/minor face), then the theme defaults (the theme minor font at 1.4 x the chart size, `#595959`). Theme typefaces (`+mn-lt`, `+mj-lt`) resolve to the theme's latin faces. Run typefaces are measured and painted through the viewer's font-family mapping and are part of the placement key. Underline is drawn as a single line below the baseline (all `u` values but `none`; double/wavy styles are not distinguished). Mixed-size runs share one baseline (the tallest ascent, measured where the canvas reports it, else 0.8 of the size); a single-run line keeps the plain title's drawing. Paragraphs and `a:br` are lines; the title band grows with its lines (1.2 x line spacing, unverified) but never beyond half the chart area: lines that do not fit are omitted with `chart-title-truncated`, the first line too (the title is then not drawn, plain titles included), so the plot never loses its height to the title. At most 64 records (text runs and line breaks alike), 16 lines and 1,024 text units per title are kept, charged before they are built; a title cut by these bounds marks the chart truncated, carries `chart-title-truncated`, and its shown (painted and described) title is what the runs keep. A run-less plain title draws exactly as before.
- **Legend.** `legendPos` r, l, t, b and tr as before; `overlay=1` keeps the legend's place but takes no space from the plot and paints it over the plot (likewise a title `overlay=1`). A point legend lists the points (categories, each with its point colour): pie and doughnut, and any chart whose single series has varyColors on (bar, line and the rest). `legendEntry/delete` removes a point entry by its point index, and otherwise a series entry by the series' `c:idx` (any unsigned index; never its document position). Combined with A6a and A5d: trendline entries follow the series (or a single varyColors series' points); a trendline's entry index is the highest series `c:idx` + 1 + its ordinal in series order (in a point legend: the point count + its ordinal), so it never shares a series or point key; a contour's band entries are deleted by band index. This numbering scheme is **unverified against Office**. Office's legend for combination charts and trendline entries is **unverified**.
- **Multi-level categories.** Every `c:lvl` of `multiLvlStrCache` is read: the first is the innermost (the axis labels as before), each further level a tier of group labels. A group spans from its `pt@idx` to the next label of its level or of any outer level; the cache's `ptCount` now sets the category count. Each level is sized by a walk over the parsed nodes in place (no copy of the source lists), charged (with a sort workspace of one key and position per valid label for out-of-order indices), then built from valid entries only; the walk stops at the first level over the bounds and reads no level outside it. An **uncached** `multiLvlStrRef` reads its outer levels from the worksheet through the same fill (resident and paged): the reference's lines outside the innermost (columns leftward of the last for a vertical range, rows above the last for a horizontal one), a non-empty cell starting a group; only the eight outer lines the fill can keep are acquired (resident and paged), the excess reported from the full range; an outer cell without a saved value makes that level and those outside it undrawn, with `chart-category-level-unreadable`; levels over the bounds give `chart-category-levels-truncated`. Tiers are drawn beyond the innermost labels (rows below a horizontal axis, columns beside a vertical one), each label centred on its span and compressed to it, with separators at the group edges from the axis through the tier (the look is **unverified against Office**; Office also wraps long group labels). Spans follow the display view (plotVisOnly, the 4,096-point display bound). `noMultiLvlLbl=1` draws the innermost level only, pixel-identical to a single-level axis. At most 4,096 outer labels and 8 outer levels are kept; a level that does not fit is dropped with every level outside it, and a cut label or dropped level marks the chart truncated.
- **Validation status.** Pixel identity is from a scratch hash harness (node-canvas) over the chart test files; browser evidence is Chromium only (WebKit cannot launch on the build host). No Office-rendered reference is compared.
- **Evidence.** All fixtures are Office-shaped synthetic. None is Office evidence.

## A6b error bars: implemented and unverified behaviour (spec 5.7)

Fixtures are Office-shaped synthetic (`tests/charts-error-bars.test.ts`); none of this is matched against an Office reference yet.

- **100% stacked groups.** Bars are drawn in the normalised (share) space. The anchor is the stacked share. Fixed and custom amounts are divided by the slot's total of |values|, the same normalisation as the point. A percentage is relative to the normalised value. stdDev and stdErr run over the normalised values. This is an **approximation** and is **unverified against Office**.
- **stdErr** is s/√n, with s the n−1 sample deviation. `val` has no role. Unverified.
- **stdDev on stacked groups** is centred on the series mean plus the point's stack offset. Unverified.
- **Absent `val`** is 1 (the schema has no default). Unverified.
- **Custom amounts** are signed as authored (a negative plus amount draws downward). A point missing an amount its `errBarType` needs draws no bar. Unverified.
- **Hidden points** (plotVisOnly) leave the stdDev/stdErr computation. Unverified.
- **Default rendering** without `a:ln` is #595959 at 1 px, with 3 px end caps either side. Unverified constants.
- **Automatic axes** include the bar ends (100% stacked axes stay 0–100%). Unverified.

## B2 waterfall and funnel (ChartEx display families)

Waterfall and funnel are painted natively (`CHARTEX_PAINTED`). They are display families: at most 4,096 points are
drawn, and the model is capped and marked `truncated` at parse (cache) or fill (reference). This was already
implemented and tested in B1 (`charts-chartex-review.test.ts` I1).

Native evidence: WF1, WF2 and FN1 (Excel) and FN2 (PowerPoint) render natively end to end. The `mc:Choice` is
selected, and FN2's coverage is `native` / `chart`. The subtotal, override, connector-off and several-series cases
use Office-shaped synthetic parts, which are not Office evidence.

Rules implemented from the map but **unverified** against Office output:

- **Subtotal height** (map §6, §12.3): a subtotal bar is drawn from 0 to the running total `acc`, not to the
  point's own cell value. The accumulator continues from it. No native part lists a `subtotals/idx`. The data label
  of a subtotal still shows the point's own value.
- **Absent `connectorLines`** defaults to on (map §4.1). WF1 omits it and draws 11 connectors.
- **Connector geometry and style:** a horizontal line at the level after bar *i*, from its far edge to the near edge
  of bar *i+1*. It is drawn only between two adjacent drawn bars, as a thin grey line (`#7F7F7F`, 0.75 pt).
- **Kind colours:** increase, decrease and total use style palette entries 0, 1 and 2 (accents 1-3), resolved at
  parse. The legend shows Increase, Decrease and Total. A `dataPt` fill or colour overrides one bar. A series-level
  `spPr` fill is not applied to the bars.
- **Waterfall point states** (map §2.4 [P]): blank, error, unavailable and hidden points draw no bar and contribute 0.
  An error is diagnosed (`chart-waterfall-error-point`); an unavailable point raises `chart-point-unavailable`.
- **Ordinal category labels** (map §2.3 [P]): a waterfall or funnel with no category dimension is labelled 1..n. A
  category role filled per series (XLSX `_xlchart` names) is used for the labels.
- **Funnel geometry** (map §9 [U]): bars run top to bottom in source order. Each bar is centred, its width is
  value / max of the plot width, and its thickness is band / (1 + gapWidth). Values of zero or below draw a
  zero-width bar but keep their label. Category labels are drawn right-aligned to the left of the plot.
- **Funnel label positions:** `ctr` (the default) is the bar centre. `inEnd`, `inBase` and `outEnd` are placed at
  the bar ends, and `t`, `b`, `l` and `r` are offset 8 px. The waterfall default is `outEnd`.
- **Several visible series** are unsupported, and the fallback picture is kept. For funnel this is map §9 [U]. For
  waterfall it is a **[P] project rule**, recorded in map §6 by controller ruling at B2 review round 1: no native
  file shows a multi-series waterfall. Hidden series never reach the display view, so they change neither the
  slots, the geometry, the labels nor the legend.
- **Error bars:** waterfall and funnel never draw error bars. They stay outside A6b's drawable families (the
  ChartEx families are on its never-draw list), and the ChartEx parser reads no error bars; a set placed on such a
  model is dropped with `chart-error-bars-dropped` (pinned in the B2 tests).
- **Subtotal point states [P]:** a subtotal index on a hidden, blank, error or unavailable point draws no bar and
  leaves the running total unchanged; an error is diagnosed like any other waterfall point.
## B5 treemap and sunburst: project rules and unverified behaviour (chartex-map §7)

Evidence: TM and SB (Office-authored inputs; their parts carry no derived tree and their XLSX fallback is a text box,
so no Office output confirms any rule below). Tests: `tests/charts-chartex-treemap.test.ts` (hand-computed trees,
the Bruls squarified oracle, sunburst angles) are Office-shaped synthetic beyond the native inputs.

- **Hierarchy levels [P]: unverified.** A multi-column category reference is one level per column (`f@dir="col"`, the
  default; one per row for `row`), the leftmost root. Every level is read on the resident and paged XLSX paths and from
  the DOCX/PPTX embedded workbook. More than 64 levels, or a level over 65,536 points, refuses the chart before any read.
- **Multi-level caches [P].** Ignored; the referenced range is read; with no readable reference the chart is unsupported.
- **Carry-forward [P]: unverified.** A blank cell inherits the current path at its level; a changed non-blank ancestor
  clears the deeper levels, an identical one keeps them; levels below a row's deepest label are ignored (a leaf); a
  level cleared by an ancestor change is an unnamed node `""` (diagnosed); a row with no label is excluded (diagnosed);
  a row excluded for its weight still moves the path.
- **Identity and aggregation [P]: unverified.** Full-path identity; duplicate paths summed; a parent is the sum of its
  descendants; a path that is both a leaf and a prefix keeps its own weight as an unnamed child `""`.
- **Value labels.** Node values shown by `dataLabels` use the label `numFmt` (unless source-linked; else General), formatted at stage 1.
- **Order [P].** A row with no label at any level is excluded before its weight is read (an unavailable weight on such a row does not refuse).
- **Weights [P].** Positive finite values only; zero, negative, blank and error weights are excluded with a diagnostic;
  an unavailable weight refuses the chart. Error bars are never drawn (outside A6b's drawable families).
- **Node bound.** More than 4,096 nodes (internal and unnamed nodes counted) refuses the chart with "This chart has too
  many data points to summarize".
- **Label budget [P].** A hierarchy label dropped past the chart's 4 MiB label budget refuses the chart (a dropped label
  would read as a blank cell and re-parent its row); a label cut to 255 units is diagnosed and may merge two paths.
- **parentLabelLayout: unverified.** Absent means `overlapping` (the proposed default). `none`: no parent labels;
  `banner`: a one-line strip on top of each parent holds its label and the children fill the rest; `overlapping`: the
  label is drawn over the children at the parent's top-left. Parent labels are drawn whatever `dataLabels` says.
- **Layout: unverified.** Squarified treemap (descending weight, ties in source order); sunburst rings of R/levels from a
  hole-less centre, clockwise from 12 o'clock in source order. Only area and angle proportions are oracles.
- **Colours: unverified.** Each root-level branch takes the series colour cycle by its root ordinal (style and colour
  parts, then theme); an authored series fill paints every node. `dataPt` overrides are not applied (row or node
  identity unknown). Only the first visible series is drawn (a further one is diagnosed).


## C1 gallery (spec 8, Goldens)

`src/testdata/chart-gallery.ts` writes one Office-shaped synthetic document per format with the same 20 charts, each
titled by what it shows. Every PNG was inspected on the Mac when recorded:

- `chart-gallery-docx.0.png`: clustered column with rich title runs (bold + red italic) and outside-end data labels,
  legend on top; stacked bar; stacked area; line with markers, a dashed linear trendline (legend entry "Linear
  (Actual)") and a smoothed series; pie with best-fit percentage labels; doughnut.
- `chart-gallery-docx.1.png`: pie of pie (Other slice centred at 3 o'clock facing a full secondary pie, series lines from its top and bottom edges); scatter with fixed
  ±1.5 error bars and a polynomial (order 2) trendline; bubble; radar (two series, marker style); contour with four
  bands; stock (volume bars on the left axis, open-high-low-close up/down bars and high-low lines on a right 30–60 axis).
- `chart-gallery-docx.2.png`: column + line combination on a secondary percent axis; line on a month date axis
  (`mmm-yy`, labels thinned); column on a base-10 logarithmic axis (1–10000); column with two-level categories
  (months under 2025 / 2026); waterfall (Increase / Decrease / Total, subtotals at "Subtotal" and "End", connectors);
  funnel with centred values.
- `chart-gallery-docx.3.png`: treemap (banner parent labels Fruit / Veg / Nuts, leaf labels with values where they
  fit); sunburst (inner ring of groups, outer ring of items).
- `chart-gallery-pptx.0-4.png`: the same 20 charts, four per slide, at 400 × 260 px (all category labels fit, so the
  date axis shows every month and the waterfall every step label).
- `chart-gallery-xlsx.0.png`: the hierarchy cells (A1:C9, the treemap/sunburst source through `_xlchart.v1.*` names)
  and the 20 charts in a 3-column grid of absolute anchors.

No gallery chart is a card or a fallback (checked in CI by `tests/chart-gallery.test.ts`). The histogram, Pareto and
box & whisker gallery entries are added when B3 and B4 merge (pending merge).

## C1 residual differences

Observed while inspecting the gallery and the torture documents; none is claimed against an Office reference.

- **Pie of pie, fixed in C1:** the secondary pie swept each slice's share as radians (a ~57° fan, A5c). It is now a
  full circle (`ofPieSlices`; test in `charts-families-ofpie`). The bar-of-pie fractions were already right.
- **Pie of pie orientation, fixed in C1 phase 1.5 (visual review, agy, Issue 1):** the composite Other slice used to
  follow `firstSliceAng` (in the gallery it sat at 198°–270°), so a series line cut through the primary pie. The
  primary pie is now turned so the Other slice is centred on 3 o'clock (angle 0), facing the secondary plot, and
  `firstSliceAng` is ignored for ofPie, as Office does; the series lines leave the slice's top and bottom rim points
  without crossing the pie (geometry test in `charts-families-ofpie`). The secondary pie still starts at 12 o'clock:
  **unverified**. Goldens re-recorded for this delta only: `chart-gallery-docx.1`, `chart-gallery-pptx.1`,
  `chart-gallery-xlsx.0` (each inspected; nothing else in them changed).
- **Edge bubbles are clipped** at the plot edge (no Office axis padding; known limit from A5a), visible in the gallery.
- **Inside-label contrast (visual review Issues 2 and 3, unverified):** data labels drawn inside a fill (funnel `ctr`
  labels on the accent bars; pie labels inside the accent-1 slice) use the default dark label colour `#595959`, so
  contrast is low. Office appears to switch to a light label colour on dark fills, but its automatic label-contrast
  rule is not confirmed against a reference, so the engine does not guess one. No code change.
- **Waterfall category labels in narrow layouts (visual review Issue 4, known):** at the gallery's 336 px DOCX width
  the category axis thins its labels to every other step (Start, Returns, Costs), the engine's general non-overlap
  rule. Office's handling of crowded waterfall step labels (wrapping, shrinking or angling) is unverified. No code
  change.
- **Document-level diagnostics:** a DOCX/PPTX chart that becomes a card keeps only its reason text and the generic
  coverage reason `unsupported-chart`; the adapter's specific code (`chart-data-external`, `chart-data-corrupt`,
  `embedded-workbook-limit`, …) is on the discarded model only. The robustness tests read it through the adapter.
- **Coverage consistency:** for a card, XLSX coverage reports `status: native` (DOCX/PPTX: `unsupported`), and a DOCX
  ChartEx/model3D/OMML fallback reports feature `picture` (PPTX names `ChartEx`, `model3D`, `OMML`).
- **PPTX admission order:** PPTX reads its slides' charts in its own order, so when the chart workspace fills, the
  carded charts are not a document-order suffix (DOCX and XLSX: natives first). The count is the same in all formats.
- **XLSX external references, changed in C1:** a reference into another workbook (`[1]Sheet1!…`, directly or through
  a defined name) still reads blank and is now diagnosed `chart-data-external-reference`.

## C1 blocked ChartEx families

None. All eight ChartEx families have native Office-authored parts (13 files from LibreOffice core
`chart2/qa/extras/data`, provenance in `tests/fixtures/chartex/PROVENANCE.md`). Waterfall, funnel, treemap and sunburst
are painted (B2, B5); histogram and Pareto (B3) and box & whisker (B4) are pending merge; region map stays
fallback-only by decision (spec 10), never parsed for geography.

## Known limits

- **A5a:** edge bubbles are clipped (no Office axis padding); radar data labels are not drawn; the 25% default bubble
  size is unverified.
- **A5b:** the HLC close tick is drawn as a right-hand tick half a bar wide (M4): **unverified** against Office.
- **A5d:** the contour is a filled-cell approximation (each cell takes its band colour), not interpolated isolines.
- **A4d:** nested graphic frames inside user shapes are only diagnosed, not drawn.
- **A6b:** the unverified choices listed in its section (100% stacked normalisation, stdErr, absent `val`, custom
  amount signs, hidden points, default rendering, automatic axes).
- **B5:** `hierarchyBytes` prices a shared label once per occurrence (conservative double pricing, accepted); `dataPt`
  overrides are not applied to treemap/sunburst nodes (node identity unknown) [U]. C1 made the hierarchy legend bound
  capacity-aware.
- **A7 (pending merge):** one shared 1.5 s picture-decode deadline per page; the pre-existing paged-chrome CLS shift is
  not chart-caused (follow-up outside this project).
- **Browser evidence is Chromium only:** WebKit cannot launch on the build host (CVDisplayLink; reproduces on `main`),
  ruling at A3b.
- **Fixture licence:** the ChartEx fixtures are used under **MPL-2.0, confirmed** (PROVENANCE.md, feature branch
  `5be9dd1`).
- **Not Office evidence:** every synthetic builder fixture, the gallery and the torture documents are Office-shaped
  synthetic.

## C1 dispositions of parked review items

| Item | Disposition |
|---|---|
| A4c N1: span oracle asserts nothing exact | Fixed: `charts-smooth` asserts exact anchor spacing and every Catmull-Rom control point |
| A4c N2: scatter admission cost changed for non-smooth | Verified intended: a straight scatter series without markers is charged one segment per point (its line only, as line charts); with markers two. Pinned by the path-bound test |
| A4c N3: the 3× smoothing charge is untested | Covered: the path-bound test draws exactly 20 smoothed series (61 straight) under the same remaining budget |
| A5b: exact-coordinate stock role oracle | Added for HLC, OHLC, volume-HLC and volume-OHLC (hi-low lines, bodies, close ticks, volume bars) |
| A5b: unnamed `charge(64)` after build | Replaced by `chargeChartUpDownBars` (`CHART_STYLE_BYTES.upDownBars`), charged before either bar is built; tested including refusal |
| A5b M4: close tick | Listed as unverified (known limits) |
| B5: double-priced shared labels | Documented (accepted, conservative) |
| B5: legend bound not capacity-aware | Fixed: `hierarchyLayoutBytes(model, legendRows)` prices the rows the legend can show |
| B5: `dataPt` overrides not applied [U] | Documented (known limits) |
| Known limits list | Above, licence as confirmed |
| §8 reference-status table | Above (`Reference` column on CH-8-* and CH-5.8-*) |
