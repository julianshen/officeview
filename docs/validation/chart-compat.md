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

## Spec §8 testing and validation

One case per §8 bullet (leaf level: sub-bullets of Goldens, References and Browser tests are separate cases).

| ID | Case | Status | Evidence |
|---|---|---|---|
| CH-8-1 | Numeric oracles: hand-computed values for ticks, date axes (both systems), display units, number-format classes, ofPie splits, stock roles, bubble scaling, trendlines, error bars, histogram bins, Pareto, box-whisker quartiles, waterfall subtotals, treemap/sunburst aggregation | pending | |
| CH-8-2 | Unit tests: parse to model, and paint with a recording context, for every family and element | pending | |
| CH-8-3 | Representation tests: spec 6.3 in all three formats, including no-fallback and chunk failure | pending | A0b: `tests/charts-representation.test.ts` (Office-shaped synthetic): MC with injected unsupported/refused and surface3DChart, ChartEx with fallback, nested/alias/unknown Requires and the standalone card in all three formats; no-fallback card and broken fallback picture (all three formats); chunk failure (DOCX, PPTX) |
| CH-8-4 | Data tests: spec 6.4 (embedded workbook, precedence, budgets, cancellation) | pending | |
| CH-8-5 | Robustness tests: extreme scales, tiny units, huge forecasts, 200 charts per document, a torture file per format, asserting allocation/work counters | pending | |
| CH-8-6 | Goldens: synthetic chart gallery (one chart per family and major element), per format, recorded on the Mac | pending | |
| CH-8-7 | Goldens: existing XLSX chart goldens pixel-identical, except listed spec 5 deltas with reasons | pending | |
| CH-8-8 | References: Office-authored files and screenshots in validation/office-reference/ with producer and version, one per promised semantic case | pending | |
| CH-8-9 | References: interim python-docx/python-pptx/openpyxl + LibreOffice evidence marked unverified; Office-matched not claimed | pending | |
| CH-8-10 | References: smoothing claimed matched only after reference comparison | pending | |
| CH-8-11 | Browser: DOCX/PPTX charts render in the paged viewer | pending | |
| CH-8-12 | Browser: dark mode leaves chart pixels identical to light | pending | |
| CH-8-13 | Browser: the per-page chart list exposes accessible names | pending | |
| CH-8-14 | Browser: CLS is 0 when a chart loads | pending | |

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
