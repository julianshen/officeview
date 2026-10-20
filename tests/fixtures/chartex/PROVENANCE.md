# ChartEx native fixtures: provenance

These files are Office-authored ChartEx (`cx:`, Office 2016+) documents. They
are the native evidence for `docs/charts/chartex-map.md` (task B1, step 0).
They are untrusted test data. Nothing in them is executed.

## Source and licence

- **Source:** LibreOffice core, `chart2/qa/extras/data/`, on GitHub at
  <https://github.com/LibreOffice/core>. Each file is pinned to a full commit
  SHA (see the table), never to `master`.
- **Licence:** MPL-2.0 is recorded, under LibreOffice project policy. Contributors
  license their work under MPL-2.0 or dual MPL-2.0/LGPLv3+, and the repository
  carries `COPYING.MPL`. The test documents have no per-file header. All of them
  were committed in 2025–2026 by Kurt Nordback (tdf#165742 chartex series) or
  Devansh Varshney (histogram import/export).
  **Needs owner confirmation:** GitHub's licence detector reports the repo as
  GPL-3.0, because the top-level `COPYING` is the GPLv3 text that covers the
  combined binary. MPL-2.0 is a file-level licence, so redistributing these
  files unmodified next to MIT code is permitted, provided this notice and the
  MPL-2.0 terms (<https://mozilla.org/MPL/2.0/>) stay with them. The files must
  not be modified in place.

## How "Office-authored" was checked

1. `docProps/app.xml` `<Application>` names Microsoft Excel, Microsoft Excel
   Online or Microsoft Office PowerPoint.
2. **Pristine-serialisation check.** Office's ChartEx serialiser writes every
   empty element with a space before `/>` (for example `<cx:subtotals />`). An
   element written as `…"/>` with no space marks a hand edit made at zip level
   with `app.xml` left intact. Every `chartEx*.xml` part committed here has zero
   such elements. To reproduce: `unzip -p F xl/charts/chartEx1.xml | grep -c '[^ ]/>'`
   should print 0.
3. **Git history.** Where a file was modified after its first commit, the
   first-commit version was compared with the later one. If the later version
   added hand-edited elements, the pristine version was kept.

Rejected candidates from the same folder:

- `funnel1.xlsx`: hand-edited `cx:spPr` (lower-case `c55a11`, no space before `/>`).
- `sunburst.xlsx` and `regionMap.xlsx` at master (`03ecbbadd4`, "Support more
  children of layoutPr"): these add hand-written `<cx:parentLabelLayout val="overlapping"/>`
  and `<cx:regionLabelLayout val="bestFitOnly"/>`. The pristine `2dba43e317`
  versions are kept instead.
- `color_funnel.xlsx` and `tdf163727_histogram_roundtrip_multi_chart.xlsx`:
  saved by Collabora Office/LibreOffice, so not Office-authored.

## Files

Raw URL pattern: `https://raw.githubusercontent.com/LibreOffice/core/<sha>/chart2/qa/extras/data/<original path>`.

| Fixture | Original path | Commit (full SHA) | Producer (app.xml) | Families | SHA-256 |
|---|---|---|---|---|---|
| `waterfall/waterfall.xlsx` | `xlsx/waterfall.xlsx` | `2dba43e31707b183d9b7cb7975209bd334900268` | Microsoft Excel 16.0300 | waterfall | `f23626e52b75c7e57a6b3a6c69361bdacdbf2648a77c198963daacc4cba2e660` |
| `waterfall2/waterfall2.xlsx` | `xlsx/waterfall2.xlsx` | `996bf7400802ca79f99d590dea2f5aa8762977b7` | Microsoft Excel 16.0300 | waterfall (connectorLines, plot-surface and chart-space fills, axis titles) | `33a8b7b438870cdfe4296c34d2c6af7780ff7fa70f3502f86ce51cb0b54fa851` |
| `funnel-label-options/funnel-label-options.xlsx` | `xlsx/funnel-label-options.xlsx` | `94a1e0c05cb03c53dd3331727d0da463cd0ad9f5` | Microsoft Excel 16.0300 | funnel (dataPt, dataLabel, dataLabelHidden) | `479745fef9364b3755c37fdb13b7f75a0dd3f22390f4a29b9af9a4afcf2e71d3` |
| `funnel-pp1/funnel-pp1.pptx` | `pptx/funnel-pp1.pptx` | `c84c7b3c81308fbafa273c76a75f4a71596666b8` | Microsoft Office PowerPoint 16.0000 | funnel (cached `cx:lvl`/`cx:pt`, `externalData`, hidden series, picture fallback) | `8f971346a21010dfdb22799be79bbb883497cc260129cba41d5c9b06967fc3e3` |
| `SimpleHistogram/SimpleHistogram.xlsx` | `xlsx/SimpleHistogram.xlsx` | `5e836dfec517894e0b41d08d1234ca734366cb2b` | Microsoft Excel **Online** 16.0300 | histogram (automatic binning) | `dfe40e742cf92aad408273d17bb0733e1571c309678ec3a3148e10e45a946e56` |
| `tdf163727_histogram_underflow_overflow/…xlsx` | `xlsx/tdf163727_histogram_underflow_overflow.xlsx` | `88550ce19c0d13af868d21624892018b573996a5` | Microsoft Excel **Online** 16.0300 | histogram (binSize, underflow, overflow) | `eae06f9cf985eb1c3e6bc70fbc1aef1fc4d012ec6a3de0f9a272c310d0a393c3` |
| `paretoLine/paretoLine.xlsx` | `xlsx/paretoLine.xlsx` | `2dba43e31707b183d9b7cb7975209bd334900268` | Microsoft Excel 16.0300 | Pareto | `2bf21b804186d6d796a772ad7db46f2d7518658786ed2075de551ffb347238ab` |
| `pareto-gapwidth/pareto-gapwidth.xlsx` | `xlsx/pareto-gapwidth.xlsx` | `8cbd028c0b0fbf24ba1847d4ef06fd15529c916c` | Microsoft Excel 16.0300 | Pareto (data labels, parentLabelLayout on the bar series, axis ids swapped) | `9835a85a0003b0e74a1c1c2a6044a1297c9695c33ecc09f2edac4daf1be767d4` |
| `boxWhisker/boxWhisker.xlsx` | `xlsx/boxWhisker.xlsx` | `2dba43e31707b183d9b7cb7975209bd334900268` | Microsoft Excel 16.0300 | box & whisker (3 series) | `7c725dda9da6f571ea58824f603bf6eff7bc6c49d122ac4322cc0d6129fa7f92` |
| `box-whisker-gapwidth/box-whisker-gapwidth.xlsx` | `xlsx/box-whisker-gapwidth.xlsx` | `8cbd028c0b0fbf24ba1847d4ef06fd15529c916c` | Microsoft Excel 16.0300 | box & whisker (2 charts; chartEx2 has an empty data range and no visibility element) | `e128108b5957cd17f1d6f3a5bfdf4e4e83e3cb70095b19604487b9197248273f` |
| `treemap/treemap.xlsx` | `xlsx/treemap.xlsx` | `2dba43e31707b183d9b7cb7975209bd334900268` | Microsoft Excel 16.0300 | treemap (3-level hierarchy) | `2f08a53fb19ddd8468ec5b52d8428503f0ede06dfb9afc3c4cd02d7cfe2be3eb` |
| `sunburst/sunburst.xlsx` | `xlsx/sunburst.xlsx` | `2dba43e31707b183d9b7cb7975209bd334900268` (pristine, pre-`03ecbbadd4`) | Microsoft Excel 16.0300 | sunburst (3-level hierarchy) | `b2ffb6e101ff71e34d3cdb6b098555ae3b5dac75e2cedf025018bcd8d73367fe` |
| `regionMap/regionMap.xlsx` | `xlsx/regionMap.xlsx` | `2dba43e31707b183d9b7cb7975209bd334900268` (pristine, pre-`03ecbbadd4`) | Microsoft Excel 16.0300 | region map (`cx:geography`, `cx:geoCache`, `cx:nf`) | `fdae8ebde4239aa30f668ad1e88bf7f7e3fe5a147bf570af6ca9e129b66a9a56` |

Notes:

- The two "Excel Online" files have `uniqueId` values that begin `{00000000-…}`
  and a leaner `txPr`. They are accepted as Office-authored.
- All XLSX parts reference data only through hidden `_xlchart.vN.M` defined
  names in `xl/workbook.xml`. They carry **no** cached `cx:lvl`. The PPTX part is
  the only one with cached points.
- In XLSX the `mc:Fallback` is an `xdr:sp` text box, not a picture. In PPTX it is
  a `p:pic`. The `mc:Choice` requirement is `cx1`
  (`http://schemas.microsoft.com/office/drawing/2015/9/8/chartex`) in XLSX and
  `cx2` (`…/2015/10/21/chartex`) in PPTX. The part namespace is always
  `http://schemas.microsoft.com/office/drawing/2014/chartex`.

## Families

All eight families have at least one Office-authored part: waterfall, funnel,
histogram, Pareto, box & whisker, treemap, sunburst and region map. **No family
is blocked.** Specific attributes that no native part exhibits are marked
*unconfirmed* in the map.
