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
| After B5 | B5 commit | 38,639 | Before (`7c67e33`, rebuilt): 38,640. Treemap and sunburst (`src/charts/chartex/hierarchy.ts`, `hierarchyStyle.ts`, `treemap.ts`, `sunburst.ts`) and the hierarchy hooks in fill, precedence, embedded reader, stage 1, layout and paint all live in the lazy `charts-*` chunk; the two description names in the `imageAdmission` chunk. No eager source file was touched; measured −1 B (hash churn, not normalised). Bytes added by B5: 0. |

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
- **Weights [P].** Positive finite values only; zero, negative, blank and error weights are excluded with a diagnostic;
  an unavailable weight refuses the chart.
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

