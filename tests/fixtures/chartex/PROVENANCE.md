# ChartEx native fixtures: provenance

These files are Office-authored ChartEx (`cx:`, Office 2016+) documents. They
are the native evidence for `docs/charts/chartex-map.md` (task B1, step 0).
They are untrusted test data. Nothing in them is executed.

## Source and licence

- **Source:** LibreOffice core, `chart2/qa/extras/data/`, on GitHub at
  <https://github.com/LibreOffice/core>. Each file is pinned to a full commit
  SHA (see the table), never to `master`.
- **Licence: MPL-2.0** (contributed under the MPL-2.0 / LGPLv3+ dual licence;
  we use them under MPL-2.0). Checked against primary sources on 2026-10-21:
  - LibreOffice's product licence (`readlicense_oo/license/license.xml`) says the
    product is made available under MPL-2.0; the repository carries
    `COPYING.MPL` and `COPYING.LGPL`. GitHub reports the repo as GPL-3.0 only
    because its detector reads the top-level `COPYING`, which is the GPLv3 text
    for bundled third-party components. It is not the licence of contributed
    files.
  - The files carry no per-file header, so their licence follows the
    contributors' standing licence statements (TDF wiki,
    `Development/Developers`):
    - Kurt Nordback (commits `2dba43e317`, `8cbd028c0b`, `94a1e0c05c`,
      `996bf74008`, `c84c7b3c81`) commits as `kurt.nordback@collabora.com`. The
      documents' `docProps/core.xml` creator is Andras Timar, also
      `@collabora.com`. Both are covered by Collabora's company statement:
      "All contributions past, present, and future … made to LibreOffice from
      Collabora are available under the terms the MPL / LGPLv3+"
      (<https://lists.freedesktop.org/archives/libreoffice/2012-April/030629.html>).
    - Devansh Varshney (commits `5e836dfec5`, `88550ce19c`; the two histogram
      files) has an individual statement: "All of my past & future
      contributions to LibreOffice may be licensed under the MPLv2/LGPLv3+ dual
      license"
      (<https://lists.freedesktop.org/archives/libreoffice/2024-January/091510.html>).
  - None of the fixture commits mentions a third-party (for example Bugzilla)
    attachment as the source.
  - MPL-2.0 is a file-level licence. These files are redistributed unmodified
    next to MIT code, with this notice and the MPL-2.0 terms
    (<https://mozilla.org/MPL/2.0/>). The files must not be modified in place.

## How "Office-authored" was checked

This is a screening process, not proof of authorship.

1. **Metadata.** `docProps/app.xml` `<Application>`/`<AppVersion>` reports
   Microsoft Excel, Microsoft Excel Online or Microsoft Office PowerPoint.
   - This is a **metadata-reported** producer: a later zip-level edit can leave
     `app.xml` untouched.
   - No file here was independently re-saved in Office for this project.
   - The AppVersion values (16.0300, 16.0000) identify the product line, not a
     precise build.
2. **Serialisation screening heuristic.** Office's ChartEx serialiser, as seen
   in these files, writes every empty element with a space before `/>` (for
   example `<cx:subtotals />`). An element written as `…"/>` **flags a
   possible** hand edit for closer inspection; it does not prove one.
   Whitespace and hex case are not authorship evidence on their own.
   - Every `chartEx*.xml` part committed here has zero flagged elements.
   - To reproduce: `unzip -p F xl/charts/chartEx1.xml | grep -c '[^ ]/>'`
     prints 0.
3. **Git history and XML comparison (the deciding check).** Where a file was
   modified after its first commit, the first-commit part was diffed against
   the later one, and the earliest version that matches the Office metadata was
   kept.

Rejected candidates from the same folder. The reasons below are this
project's assessment; reviewers could not independently retrieve the rejected
bytes, so they are not independently confirmed.

- `sunburst.xlsx` and `regionMap.xlsx` at master (`03ecbbadd4`, "Support more
  children of layoutPr"). **Confirmed by XML diff** against `2dba43e317`: the
  later commit adds `<cx:parentLabelLayout val="overlapping"/>` (sunburst) and
  `<cx:regionLabelLayout val="bestFitOnly"/>` (region map), and the commit's
  purpose is to test those elements. The pristine `2dba43e317` versions are
  kept.
- `funnel1.xlsx`: its series `cx:spPr` was flagged by the heuristic
  (lower-case `c55a11`, no space before `/>`). It has a single commit, so there
  is no earlier version to diff against. It was excluded as **likely
  modified**, not proven modified; FN1 and FN2 cover funnel without it.
- `color_funnel.xlsx` and `tdf163727_histogram_roundtrip_multi_chart.xlsx`:
  `app.xml` names Collabora Office/LibreOffice, so they are not Office-authored
  by their own metadata.

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
- In XLSX the `mc:Fallback` is an `xdr:sp` text box, not a picture. In PPTX it
  is a `p:pic`.
- `mc:Choice@Requires` tokens and their bound URIs. The prefix is lexical; the
  URI is what matters.
  - `cx1` = `http://schemas.microsoft.com/office/drawing/2015/9/8/chartex`:
    WF1, WF2, HI1, HI2, PA1, PA2, BW1 (and both BW2 drawings), TM, SB.
  - `cx2` = `http://schemas.microsoft.com/office/drawing/2015/10/21/chartex`:
    **FN1 (xlsx)** and FN2 (pptx).
  - `cx4` = `http://schemas.microsoft.com/office/drawing/2016/5/10/chartex`:
    RM (xlsx).
- The `a:graphicData@uri`, and the root namespace of all 14 chart parts, is
  always `http://schemas.microsoft.com/office/drawing/2014/chartex`.

## Families

All eight families have at least one Office-authored part: waterfall, funnel,
histogram, Pareto, box & whisker, treemap, sunburst and region map. **No family
is blocked.** Specific attributes that no native part exhibits are marked
*unconfirmed* in the map.
