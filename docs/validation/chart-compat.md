# Chart compatibility validation

Status of every promised validation case of the chart-compatibility project
(spec `docs/superpowers/specs/2026-10-20-chart-compat-design.md`, plan
`docs/superpowers/plans/2026-10-20-chart-compat.md`). Created by task F0 as a stub;
later tasks update the status, evidence and provenance of each case. Case IDs are
stable: `CH-<spec section>-<n>`, never renumbered.

Statuses: `pending`, `passing` (synthetic, Office-shaped), `unverified` (interim
non-Office evidence), `matched` (Office reference compared). Synthetic builder
fixtures are "Office-shaped synthetic" and are never reported as Office evidence.

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
| After A4b | A4b commit | 38,640 | Before (`f04481f`, rebuilt from a clean `git archive`): 38,636. Title runs, legend overlay/legendEntry, multi-level category tiers (`src/charts/titles.ts`, `src/charts/categories.ts`, layout, paint, parse, style part) are all in the lazy `charts-*` chunk. `dist/index.js` byte-identical to the base build after normalising chunk hashes. Bytes added by A4b: 0 (hash churn only). Running total vs B0: +9 B |
| After B2 | B2 commit | 38,636 | Before (`7c67e33` tree plus B2's type-only `chartEx.palette` field, rebuilt): 38,640. Waterfall and funnel families (`src/charts/chartex/waterfall.ts`, `funnel.ts`) and their hooks in axes, display, prepare, layout, labels and paint live in the lazy `charts-*` chunk; `dist/index.js` changes by hash churn only. Bytes added by B2: −4 (hash churn). Running total vs B0: +5 B. Ceiling 38,647 |
| After B2 fix round 1 | B2 r1 commit | 38,635 | Subtotal state gate, finite funnel ratio, hidden series out of the display view, funnel label pricing: lazy `charts-*` chunk only (hash churn). Running total vs B0: +4 B |
| After B2 merge with feature | merge commit | 38,635 | `feature/chart-compat` `cb5a0b5` merged into `chart/b2`; conflicts resolved keeping both sides (lazy `charts-*` chunk only). Running total vs B0: +4 B. Ceiling 38,647 |
| After B5 | B5 commit | 38,639 | Before (`7c67e33`, rebuilt): 38,640. Treemap and sunburst (`src/charts/chartex/hierarchy.ts`, `hierarchyStyle.ts`, `treemap.ts`, `sunburst.ts`) and the hierarchy hooks in fill, precedence, embedded reader, stage 1, layout and paint all live in the lazy `charts-*` chunk; the two description names in the `imageAdmission` chunk. No eager source file was touched; measured −1 B (hash churn, not normalised). Bytes added by B5: 0. |
| After B5 merge with feature | merge commit | 38,637 | `feature/chart-compat` merged into `chart/b5` (B2, A6b, A5c, A4b and the rest); conflicts resolved keeping both sides; B5 review minors (label numFmt, exclusion order) applied. Lazy `charts-*` chunk only. Before (feature ledger, B2 merge): 38,635; +2 B hash churn. Ceiling 38,647 |
| After A7 fix round 1 | A7 r1 commit | 38,584 | Before (`db78d06`, rebuilt): 38,637. A7's descriptions, paintables metadata and chart list live in the lazy `charts-*`, `office-worker` and `PagedChrome` chunks. Review I1 needs `OfficeDoc` to abort the chart picture wait on unmount: one `AbortController` and its `signal` passed to `getPaintables` (+25 B, measured 38,664). Recovered by condensing three verbose style comments inside `OfficeDoc`'s JSX that `dist/index.js` preserves verbatim (−80 B), with no behaviour change. Running total vs B0: −47 B |
| After B3 | B3 commit (`f2c8ee3`) | 38,635 | Before (`7c67e33`, rebuilt): 38,640. Histogram and Pareto (`src/charts/chartex/histogram.ts`, the stage-1 hook, the ChartEx adapter): all in the lazy `charts-*` chunk. `dist/index.js` unchanged apart from hash churn (−5 B). Bytes added by B3: 0. Running total vs B0 (38,631): +4 B |
| After B3 fix round 1 | B3 r1 commit | 38,638 | Before (`d5eaff5`, B3 merged with `feature/chart-compat` at `e7e64f5`, rebuilt): 38,639. Leaf reason constants (`src/charts/limits.ts`), range-safe binning and override remapping (lazy `charts-*`); deferred XLSX chart fallbacks (`settleChartFallbacks` in `src/xlsx/drawing.ts`, in the worker chunks, not `dist/index.js`). `dist/index.js` unchanged apart from hash churn (−1 B). Bytes added: 0. Running total vs B0: +7 B |
| After B3 fix round 2 | B3 r2 commit | 38,638 | Before (`0cdda65`, rebuilt): 38,638. Branch-unit fallback descriptors and post-settlement media indexing (`src/xlsx/drawing.ts`, `src/xlsx/parse.ts`, worker chunks); plain-first arithmetic and per-slot overrides (lazy `charts-*`). `dist/index.js` unchanged. Bytes added: 0. Running total vs B0: +7 B |
| After B3 fix round 3 | B3 r3 commit | 38,638 | Before (`74670f5`): 38,638. Nested fallback-wrapper hierarchy in `src/xlsx/drawing.ts` (worker chunks). `dist/index.js` unchanged. Bytes added: 0. Running total vs B0: +7 B |
| After B3 merge with feature | merge commit | 38,638 | `feature/chart-compat` `5be9dd1` (B2, B5, A6b, A5c, A4b, licence) merged into `chart/b3`; conflicts resolved keeping every family and section. Bytes added: 0. Running total vs B0: +7 B |
| After B3 fix round 4 + merge with feature (B4) | merge commit | 38,639 | Round 4 (per-pass wrapper settlement, ancestor membership) in worker chunks; `feature/chart-compat` with B4 box & whisker merged (CHARTEX_PAINTED keeps all seven painted families). +1 B hash churn. Running total vs B0: +8 B |
| After B3 merge with feature (A7) | merge commit | 38,582 | `feature/chart-compat` with A7 (descriptions in `src/charts/describe.ts`, which freed eager bytes) merged into `chart/b3`; B3 adds `derivedFamily` naming (lazy). Running total vs B0: −49 B |

## Spec §8 testing and validation

One case per §8 bullet (leaf level: sub-bullets of Goldens, References and Browser tests are separate cases).

| ID | Case | Status | Evidence |
|---|---|---|---|
| CH-8-1 | Numeric oracles: hand-computed values for ticks, date axes (both systems), display units, number-format classes, ofPie splits, stock roles, bubble scaling, trendlines, error bars, histogram bins, Pareto, box-whisker quartiles, waterfall subtotals, treemap/sunburst aggregation | pending | A1a (ticks, date axes in both systems, display units; Office-shaped synthetic, not Office-matched): `tests/charts-axes.test.ts` (scales, crossings, log, dispUnits, tick bound) and `tests/charts-dates.test.ts` (epoch precedence, civil-model conversion of serials 0/1/59/60/60.5/61 both ways, calendar month/year steps, unequal spacing). Default unit labels other than hundreds/thousands/millions/billions/trillions are unverified |
| CH-8-2 | Unit tests: parse to model, and paint with a recording context, for every family and element | pending | |
| CH-8-3 | Representation tests: spec 6.3 in all three formats, including no-fallback and chunk failure | pending | A0b: `tests/charts-representation.test.ts` (Office-shaped synthetic): MC with injected unsupported/refused and surface3DChart, ChartEx with fallback, nested/alias/unknown Requires and the standalone card in all three formats; no-fallback card and broken fallback picture (all three formats); chunk failure (DOCX, PPTX) |
| CH-8-4 | Data tests: spec 6.4 (embedded workbook, precedence, budgets, cancellation) | pending | D0 (XLSX clause only; Office-shaped synthetic): `tests/charts-data-contract.test.ts` resident vs paged point arrays (sparse caches, references, blanks, #N/A vs #DIV/0!, formulas without saved values), analytic acquisition (65,536 / refusal), plotVisOnly provenance, paged parent-ledger ownership and largeSource single workspace. The embedded workbook is a later task |
| CH-8-5 | Robustness tests: extreme scales, tiny units, huge forecasts, 200 charts per document, a torture file per format, asserting allocation/work counters | pending | A1a: tiny majorUnit/minorUnit and a 0..1e300 scale stay within 1,000 ticks per axis with `chart-tick-limit` (`tests/charts-axes.test.ts`) |
| CH-8-6 | Goldens: synthetic chart gallery (one chart per family and major element), per format, recorded on the Mac | pending | |
| CH-8-7 | Goldens: existing XLSX chart goldens pixel-identical, except listed spec 5 deltas with reasons | pending | D0: strict goldens 3/3, `tests/goldens` untouched (no golden contains a chart). Listed delta (spec 5.8 `plotVisOnly`, Office default on): a reference-filled point in a hidden row or column is no longer drawn, and a series whose every point is hidden is not drawn (`tests/charts-data-contract.test.ts` pixel check). Non-value point states still draw as gaps (the matrix is A2). A1a: the painter draws on the axis graph; every existing chart fixture repaints pixel-identically (hashes of all 41 chart paints in the chart/drawing/docx/pptx/golden suites; of 48 synthetic variant renders, all identical except the two deleted-category-axis renders, which lose the axis line row: delta 1 below). The strict golden suite contains no chart and is not independent evidence of chart parity. Listed deltas (spec 5.1): a deleted category axis no longer draws its axis line; a group on a secondary value axis is scaled by that axis; a deleted scatter x axis still draws its authored gridlines (all three apply only to charts that author them) |
| CH-8-8 | References: Office-authored files and screenshots in validation/office-reference/ with producer and version, one per promised semantic case | pending | |
| CH-8-9 | References: interim python-docx/python-pptx/openpyxl + LibreOffice evidence marked unverified; Office-matched not claimed | pending | |
| CH-8-10 | References: smoothing claimed matched only after reference comparison | pending | A4c: smoothing implemented via Catmull-Rom cubic Bézier (tension ½, end point tangents), marked approximation and unverified until Office reference comparison (including absent c:smooth default on c:ser vs group) |
| CH-8-11 | Browser: DOCX/PPTX charts render in the paged viewer | pending | |
| CH-8-12 | Browser: dark mode leaves chart pixels identical to light | pending | |
| CH-8-13 | Browser: the per-page chart list exposes accessible names | pending | |
| CH-8-14 | Browser: CLS is 0 when a chart loads | pending | |

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
numeric oracle and needs its Office reference before it is claimed matched; standard
area blank/#N/A cells are checked under each `dispBlanksAs` value (gap, zero, span).

| ID | Family | Point state | Specified behaviour | Status | Evidence |
|---|---|---|---|---|---|
| CH-5.8-1 | Line, scatter, radar | blank | per dispBlanksAs (gap: break; zero: 0; span: connect) | pending | |
| CH-5.8-2 | Line, scatter, radar | #N/A | no marker; line connects across it whatever dispBlanksAs says | pending | |
| CH-5.8-3 | Line, scatter, radar | other error | as #N/A | pending | |
| CH-5.8-4 | Line, scatter, radar | unavailable | unconditional gap, plus chart-point-unavailable | pending | |
| CH-5.8-5 | Column, bar | blank | no bar | pending | |
| CH-5.8-6 | Column, bar | #N/A | no bar | pending | |
| CH-5.8-7 | Column, bar | other error | no bar | pending | |
| CH-5.8-8 | Column, bar | unavailable | no bar, plus diagnostic | pending | |
| CH-5.8-9 | Area (standard) | blank | gap if dispBlanksAs=gap, else 0 | pending | |
| CH-5.8-10 | Area (standard) | #N/A | gap if dispBlanksAs=gap, else 0 | pending | |
| CH-5.8-11 | Area (standard) | other error | as #N/A | pending | |
| CH-5.8-12 | Area (standard) | unavailable | as blank, plus diagnostic | pending | |
| CH-5.8-13 | Stacked or 100% (any family) | blank | 0 | pending | |
| CH-5.8-14 | Stacked or 100% (any family) | #N/A | 0 | pending | |
| CH-5.8-15 | Stacked or 100% (any family) | other error | 0 | pending | |
| CH-5.8-16 | Stacked or 100% (any family) | unavailable | 0, plus diagnostic | pending | |
| CH-5.8-17 | Pie, doughnut | blank | no slice | pending | |
| CH-5.8-18 | Pie, doughnut | #N/A | no slice | pending | |
| CH-5.8-19 | Pie, doughnut | other error | no slice | pending | |
| CH-5.8-20 | Pie, doughnut | unavailable | no slice, plus diagnostic | pending | |

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

## B3 histogram and Pareto: implemented and unverified rules (chartex-map §5)

Implemented in `src/charts/chartex/histogram.ts`. Stage 1 (`prepareChartAxes`) turns the ChartEx group into a plain clustered column group (and a line group for the Pareto), labelled with the bin or category texts. Tests are in `tests/charts-chartex-histogram.test.ts`. The native evidence is HI1, HI2, PA1 and PA2, read end to end. Every other part is Office-shaped synthetic.

The following rules are **unverified** ([U]/[P]). No Office-rendered reference has been compared for any of them.
- **Scott variant:** `h = 3.49·s·n^(−1/3)` with the sample s of the inner set. Bins start at `min` (or U), with no rounding of `h` or of the start.
- **Automatic bins, binCount and closure `l`:** no native example exists for `binCount`, and none for automatic binning with a count. Closure `l` and its mirrored outer bins (`<U`, `≥O`) also have no native example.
- **Degenerate inputs:**
  - `hi = lo` gives one bin `[lo, lo]` in every mode.
  - An empty inner set draws only the outer bins.
  - A Scott width that cannot be computed (fewer than two inner values) while `hi > lo` is refused as "invalid binning".
- **Bin labels:**
  - The glyphs are `(a, b]`, `[a, b)`, `≤U`, `>O`, `<U` and `≥O`.
  - Edges use the category axis's authored format, else the first value's source format, else General (15 significant digits). This makes long labels for Scott widths.
  - An empty trailing inner bin is drawn (HI2 `(13, 15]`), and its upper edge is not clipped to O.
- **Hidden cells:** under plotVisOnly, hidden source cells are excluded from the analytic input.
- **Several column series:** a histogram pools one set of bins over every visible series, drawn clustered, and requires every series to use the same binning. A Pareto with more than one column series is unsupported.
- **By Category:**
  - Categories are grouped by exact text in first-occurrence order and their values summed.
  - A category whose points are all blank or error still appears, at 0.
  - A missing category dimension is unsupported.
  - Controller ruling: the brief's "pre-aggregated data drawn as given" is tested as unique categories drawn as given. Duplicates are summed.
- **Pareto:**
  - The owner is the series `ownerIdx` names (document position).
  - Ties keep their prior order (bin order, or the first occurrence of the category).
  - The line passes through the band centres. An all-zero total draws a zero line.
  - The percent axis is the paretoLine's value axis (min 0, max 1 and `0%` unless authored), placed on the right. The line's colour comes from its `formatIdx` (accent 2 by default).
  - The line's legend name is the owner's name.
- **Gap width:** an absent `gapWidth` is 0 for binned columns and 150 % for By Category.
- **Per-point overrides ([P], fix rounds 1 and 2):**
  - The rule is decided per category slot. A By Category column whose slot holds exactly one source point of its series keeps that point's `dataPt` fill, colour and line, and its `dataLabel`/`dataLabelHidden`. A Pareto sort moves them with their column.
  - Binned columns, and slots grouped from several points, drop only the overrides addressing those points, with one `chart-histogram-point-overrides` diagnostic carrying the count.
  - Series-level labels are always kept, so PA2's counts are drawn.
- **Hidden series (fix round 1):**
  - A hidden column series is out of layout: it is not drawn, not binned, and not counted by the structural checks.
  - A hidden paretoLine draws no line, but the columns are still sorted by their owner.
  - `ownerIdx` counts every `cx:series` in document position, hidden ones included.
- **Range safety (fix rounds 1 and 2):**
  - Ordinary arithmetic is used whenever it stays representable.
  - A bin edge `lo + i·h` is halved and doubled only when the plain edge overflows, so a subnormal width keeps its value.
  - Differences are taken from quotients only when the plain difference overflows.
  - A By Category slot is summed again, with its own power-of-two scale, only when its plain sum overflows, so small sums in other slots are never scaled.
  - Pareto fractions use plain running sums unless those overflow.
  - Edges that are non-finite or not strictly increasing are refused as "invalid binning".
  - A column whose value passes the double range is refused ("This chart's values are out of range").
- **Refusal after the data is read (XLSX, fix rounds 1 and 2; every ChartEx family):**
  - A chart Choice is selected before the workbook fill reads its data. If the fill or stage 1 then refuses any chart in the selected branch, `settleChartFallbacks` replaces the whole branch with its `mc:Fallback`, parsed only at that point. "The whole branch" includes a shape beside the chart and a chart inside a group. This applies to the resident and the paged XLSX readers.
  - Nested AlternateContent (fix round 3): each wrapper is a node in a hierarchy, and a drawing maps to its innermost wrapper. Settlement runs innermost first, so the first usable inner fallback wins, as it does for a chart refused at selection. An inner replacement takes the replaced objects' place in the enclosing branch, so the object count may change. An unusable inner fallback leaves the refusal to the enclosing wrapper.
  - Fix round 4: an inner replacement is carried into every enclosing wrapper whose flattened branch held the replaced objects, so a grandparent with a refused sibling chart still gives way as a unit. Each settlement pass handles an open wrapper at most once, so a chart still waiting on data returns promptly and keeps its route for a later call.
  - The branch's descriptor lives until its charts are final. A paged chart that is still waiting on an importing sheet keeps its Choice, and the paged reader reads the sheet's drawings again once the sheet completes.
  - A selected fallback's pictures are indexed into the document's media (`collectXlsxImages`), so they paint.
  - The coverage becomes `fallback`/`fallback`, keeping the chart's feature (`ChartEx`), with the refusal as the reason.
  - With no usable fallback, the card stays and its stale `chart` coverage entry is recomputed.
  - DOCX/PPTX need no change: their data is read in the adapter, before selection.
- **Accessible description (after the A7 merge):** a derived model carries `derivedFamily`, so `chartDescription` names it "Histogram chart" or "Pareto chart" (the lazy `src/charts/describe.ts`).

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

## B4 Box & whisker (ChartEx): implemented and unverified behaviour (spec 4.3, chartex-map §8)

- **Implemented.** Box & whisker marks are drawn natively per series. Points in `value` state form the sorted distribution; blank and error points are excluded; an unavailable point or points exceeding 65,536 refuse the chart. Medians, quartiles, whiskers within 1.5×IQR, outlier markers, inner points (non-outliers), mean markers (×), and connected mean lines are rendered according to the series layout properties and visibility flags.
- **`inclusive` quartile method: unverified.** When $n$ is odd, the `inclusive` method includes the median in both halves when calculating $Q_1$ and $Q_3$, following the ST_QuartileMethod specification and `chartex-map.md` §8. This behaviour is unconfirmed against native Office ground truth because available native Office fixtures (BW1, BW2) use `exclusive`.
- **Absence defaults for box visibility flags: unverified.** Defaults (`connectorLines: true, meanLine: false, meanMarker: true, nonoutliers: false, outliers: true`) follow `chartex-map.md` §4.1 / §8 when `visibility` is omitted (as in BW2b).
- **1.5×IQR whisker factor: unverified.** Standard Tukey box plot factor $1.5 \times \text{IQR}$ is used to determine whisker reach and outlier classification. Microsoft documentation states that points beyond whiskers are outliers, but does not explicitly document the 1.5 multiplier.
- **Category axis slot mapping: unverified.** When a `cat` dimension is present, boxes are clustered per distinct category with series side-by-side inside each category; when absent, one box is drawn per series along the category axis with spacing governed by `gapWidth`. Box labels default to distinct category name when present, else series name, else ordinal.
- **Empty and missing categories: project rule, unverified.** A category dimension whose every entry is empty counts as absent (one box per series). An empty category among named ones is its own group with an empty label. A point with no category entry at all (the dimension is shorter than the values) is not drawn and is diagnosed (`chart-point-uncategorized`).
- **Category limit: project rule.** More than 4,096 distinct categories refuse the chart (`This chart has too many data points to summarize`); a prefix is never drawn.
- **Mean line with categories: unverified.** With a `cat` dimension the mean line joins each series' means across the categories, in the series colour; without one it joins the means of consecutive series (map §8), in the later series' colour. Each series' own `meanLine` flag decides whether it takes part. A series with no values in a category is joined across that empty slot (from its previous drawn box).
- **Stage 1 on every path.** XLSX charts are prepared by the workbook fill; DOCX/PPTX charts are prepared by the ChartEx adapter after their data precedence. Paint reads only the prepared records and draws no box for an unprepared model.
- **Data labels and data-point overrides: unverified.** Data labels (`cx:dataLabels`) and per-point shape properties (`cx:dataPt`) are not painted (e.g. BW2b carries `dataLabels/visibility value="1"`).
- **Whisker end caps: unverified.** Horizontal end caps at whisker endpoints are omitted; unverified whether Office renders end caps natively.

## A7 accessibility, coverage and paged chart metadata (spec 6.6)

- **Descriptions.** Each chart's description names its type, title, series names and value range. A card reads `placeholder: <reason>`.
- **Bounds.** Names are cut at 120 characters, at most 8 series are described, and each description is capped at 1,024 characters.
- **Placement.** Descriptions travel with each page's or slide's placement in the paintables metadata. That metadata is admitted to the chart workspace before it is built, per page, and refunded with the pages.
- **Per-page list.** It is capped at 64 entries. Above that, the list ends with an "N more charts not listed" note and is labelled with the true total; diagnostic `chart-description-limit`.
- **Refused admission.** A page whose admission is refused gets a readable notice and diagnostic `chart-description-refused`. Later pages are unaffected.
- **Known limitation: chart picture wait.** User-shape pictures of a model transferred from another realm are decoded before the pages are published, so that their first paint shows them. The wait is bounded and cancellable:
  - **Deadline.** One shared 1.5 s deadline, driven by an injectable clock; tests never rely on wall time.
  - **Cancellation.** It honours the caller's abort signal (OfficeDoc aborts on unmount or document change) and the chart context's signal.
  - **Missed deadline.** A page whose pictures miss the deadline is published without them and gets diagnostic `chart-user-shapes-deadline`. The decode continues, its bytes stay charged until it settles, and the picture appears the next time that page is painted. Nothing triggers that repaint (zoom, resize or a document change do).
  - **Delay to other pages.** The paint pipeline publishes all pages of a document together. A slow picture can therefore delay the first publication of every page by at most the deadline, never by more.
- **Layout shift.** The Chromium CLS test proves zero layout shift outside the viewer chrome while charts load. It does not prove whole-load CLS. A pre-existing mount shift of the status slot (seen also in documents with no charts) is excluded as chrome-only.
- **Fixtures.** All are Office-shaped synthetic; none is Office evidence.
