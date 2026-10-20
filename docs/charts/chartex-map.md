# ChartEx field map (normative)

Status: **draft for review** (task B1 step 1). It needs approval from two
reviewer families before the parser is written.
Scope: spec `docs/superpowers/specs/2026-10-20-chart-compat-design.md` §4.2 and §4.3.

This document fixes what the ChartEx parser reads, what each field means and
what applies when a field is absent. Where §4.3 says more than the sources
support, this map says so, and the parser follows this map.

## 0. Conventions

### Namespaces

| Prefix | URI | Where |
|---|---|---|
| `cx` | `http://schemas.microsoft.com/office/drawing/2014/chartex` | chartEx part root and all elements below |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | `spPr`, `txPr`, `rich` |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | `externalData@r:id` |
| `cx1` | `http://schemas.microsoft.com/office/drawing/2015/9/8/chartex` | `mc:Choice Requires="cx1"` in XLSX drawings [N: all xlsx] |
| `cx2` | `http://schemas.microsoft.com/office/drawing/2015/10/21/chartex` | `mc:Choice Requires="cx2"` in PPTX slides [N: FN2] |

The graphic frame's `a:graphicData@uri` is the 2014 `cx` URI in every native
file. The dispatch check is `cx1` or `cx2` in `Requires` plus that `uri`.

### Source tags

Every row carries one or more source tags:

- **[S]** MS-ODRAWXML, the schema and prose on the live Microsoft Learn pages
  (`updated_at` 2026-10-08). The spec gives structure, types and a few prose
  meanings. Every Appendix A product-behaviour note for §2.24 says only "not
  available in Office 2013 [or 2016] and earlier". **MS-ODRAWXML defines no
  rendering algorithms and almost no defaults.**
- **[D]** Microsoft Office support documentation. This is normative for
  behaviour, but it is not a file-format spec.
- **[N:xx]** Exhibited by a native fixture (codes in §0.4).
- **[U] unconfirmed.** Neither MS-ODRAWXML nor a native part confirms it. The
  rule stated is a project decision, and it is flagged for a numeric oracle or
  for owner review.

"Default" means the value the parser uses when the attribute or element is
absent. "none (schema)" means that MS-ODRAWXML gives no default.

### MS-ODRAWXML section index

Section numbers come from the live Learn TOC
(`https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/toc.json`).
URLs are `https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/<guid>`.
Older revisions number these sections differently, and Appendix A section
references are offset. Cite by name plus GUID.

| § | Type | GUID |
|---|---|---|
| 2.24.3.5 | CT_Axis | f2235d6d-370d-49b9-8a74-2232ffc8ab76 |
| 2.24.3.11 | CT_Binning | 2e5dfc5a-d3b7-4390-bc00-188a2fefcb3a |
| 2.24.3.12 | CT_CategoryAxisScaling | a630f348-bc49-4045-b2aa-d6b1fb6f22b6 |
| 2.24.3.14 | CT_ChartData | 8a963284-245d-4c2a-935d-cc6aef448b09 |
| 2.24.3.20 | CT_Data | b895f5cf-d6a5-4fc1-b566-7b594b2a4e60 |
| 2.24.3.21 | CT_DataId | ed0f546d-ab0a-48dd-b2fb-bd39f67b488a |
| 2.24.3.22 | CT_DataLabel | b35bbd84-1663-401d-addf-b1cabe61a54f |
| 2.24.3.23 | CT_DataLabelHidden | 2adc7fd4-a611-4180-8928-203c24813b0b |
| 2.24.3.24 | CT_DataLabels | ab25525c-374c-46c9-997e-cee7cc73436e |
| 2.24.3.25 | CT_DataLabelVisibilities | 09f308ed-c45c-4da0-8295-2200e68604db |
| 2.24.3.26 | CT_DataPoint | d0552245-43df-497b-8dc9-a107b260b20e |
| 2.24.3.29 | CT_ExternalData | 9a9d715b-dcae-4b03-89e9-a6362994e1a2 |
| 2.24.3.34 | CT_Formula | 174c9b33-29d8-4764-a3c8-7ad71abf4424 |
| 2.24.3.50 | CT_Geography | 5e722ee2-d933-454f-8198-8490cbf17e77 |
| 2.24.3.68 | CT_NumericDimension | 37fbbd78-c3e6-4f33-9d0b-3fe428ad1c09 |
| 2.24.3.69 | CT_NumericLevel | 06393495-73ef-4826-af89-4e119dc35e42 |
| 2.24.3.70 | CT_NumericValue | e3ae6c9c-786a-4b7d-920d-41e9a0f5a045 |
| 2.24.3.74 | CT_ParentLabelLayout | 748b6764-5358-450d-b0d9-4a07525915fa |
| 2.24.3.77 | CT_PlotAreaRegion | 953cac64-c127-438a-bfc4-3e1b17b3f9a4 |
| 2.24.3.80 | CT_RegionLabelLayout | 0164520d-dbce-4994-8c59-5a6d9d8af63e |
| 2.24.3.82 | CT_Series | 86df19ea-8e94-49fc-9160-377c1b40b985 |
| 2.24.3.83 | CT_SeriesElementVisibilities | dd94d472-6968-4372-9f25-0fc8fa56ffc8 |
| 2.24.3.84 | CT_SeriesLayoutProperties | 1ff5c9ad-828a-4d0c-be72-07754eba3588 |
| 2.24.3.85 | CT_Statistics | 2f21033a-29c6-4283-a9b5-63d465544799 |
| 2.24.3.86 | CT_StringDimension | ab157fdd-8e5a-4b42-8205-aaf41384d457 |
| 2.24.3.87 | CT_StringLevel | 83fa07e1-6090-4757-acc8-8dd35dac0b63 |
| 2.24.3.88 | CT_StringValue | 50e379c7-84c2-4506-a36e-4ded3d30a5ba |
| 2.24.3.89 | CT_SubtotalIndex | cd69de6d-a670-45fe-b6b3-b8b4384350d9 |
| 2.24.3.90 | CT_Subtotals | 4fc38c3f-c55e-4453-8a06-21f92056554e |
| 2.24.3.95 | CT_ValueAxisScaling | bd42958d-efa4-4400-9647-a0b3cb3d1beb |
| 2.24.4.2 | ST_AxisId | 80d24e4a-c3a3-42c0-8a20-60cf964b4341 |
| 2.24.4.3 | ST_AxisUnit | bc84901b-0b43-41cf-bf2e-a74bd1b33847 |
| 2.24.4.5 | ST_DataId | 0be78a69-1ff8-419a-bc93-320e94867198 |
| 2.24.4.6 | ST_DataLabelPos | a8ac0b74-aed1-4bcf-912a-2065d316d90b |
| 2.24.4.7 | ST_DoubleOrAutomatic | 7d4f4fcd-1ae3-4d1c-a090-c0f956d3a490 |
| 2.24.4.10 | ST_FormulaDirection | dc4c0e17-c0a4-4ceb-ab6c-e3be112af187 |
| 2.24.4.11 | ST_GapWidthRatio | 55b11e8e-31c5-49e0-85ee-90a9c7fb8e72 |
| 2.24.4.14 | ST_IntervalClosedSide | 21a3a14c-7dde-4b9f-9812-a2f9a96fc5fe |
| 2.24.4.15 | ST_NumericDimensionType | bef43d72-bbb8-44aa-ae26-107f84b88cfa |
| 2.24.4.17 | ST_ParentLabelLayout | 5ec636c6-6309-4c6d-a912-d192832a981b |
| 2.24.4.19 | ST_QuartileMethod | a1463864-276e-4cef-b772-d9b2a046e5a7 |
| 2.24.4.21 | ST_SeriesLayout | 2ea3f228-fe39-4f55-b8ec-cee89596e926 |
| 2.24.4.23 | ST_StringDimensionType | 0bad2d53-91bb-4c50-a5af-b38831ee7123 |

The local XSD cross-check (docx4j `xsd/odrawxml/office_drawing_2014_chartex.xsd`)
is MS-ODRAWXML revision 29.0 (2021). It lacks `geography`, `CT_FormatOverrides`
and the dashboard and binding types. The live pages are primary.

Office support pages cited as [D]:

- D-hist: "Create a histogram", support.microsoft.com/office/85680173-064b-4024-b39d-80f17ff2f4e8
- D-pareto: "Create a Pareto chart", support.microsoft.com/office/a1512496-6dba-4743-9ab1-df5012972856
- D-wf: "Create a waterfall chart", support.microsoft.com/office/8de1ece4-ff21-4d37-acd7-546f5527f185
- D-box: "Create a box and whisker chart", support.microsoft.com/office/62f4219f-db4b-4754-aca8-4743f6190f0d

### Native fixture codes

All files are in `tests/fixtures/chartex/` (provenance in `PROVENANCE.md`).

| Code | File | Part |
|---|---|---|
| WF1 | waterfall/waterfall.xlsx | xl/charts/chartEx1.xml |
| WF2 | waterfall2/waterfall2.xlsx | xl/charts/chartEx1.xml |
| FN1 | funnel-label-options/funnel-label-options.xlsx | xl/charts/chartEx1.xml |
| FN2 | funnel-pp1/funnel-pp1.pptx | ppt/charts/chartEx1.xml |
| HI1 | SimpleHistogram/SimpleHistogram.xlsx (Excel Online) | xl/charts/chartEx1.xml |
| HI2 | tdf163727_histogram_underflow_overflow/….xlsx (Excel Online) | xl/charts/chartEx1.xml |
| PA1 | paretoLine/paretoLine.xlsx | xl/charts/chartEx1.xml |
| PA2 | pareto-gapwidth/pareto-gapwidth.xlsx | xl/charts/chartEx1.xml |
| BW1 | boxWhisker/boxWhisker.xlsx | xl/charts/chartEx1.xml |
| BW2 | box-whisker-gapwidth/box-whisker-gapwidth.xlsx | chartEx1.xml (BW2a), chartEx2.xml (BW2b) |
| TM | treemap/treemap.xlsx | xl/charts/chartEx1.xml |
| SB | sunburst/sunburst.xlsx | xl/charts/chartEx1.xml |
| RM | regionMap/regionMap.xlsx | xl/charts/chartEx1.xml |

---

## 1. Part skeleton

`cx:chartSpace` (root, 2.24.1.2 / CT_ChartSpace 2.24.3.15) has the sequence
`chartData` (1), `chart` (1), `spPr` (0..1), `txPr` (0..1), `clrMapOvr` (0..1),
`fmtOvrs` (0..1), `printSettings` (0..1) and `extLst` (0..1). [S][N: all;
chart-space `spPr` in WF2]

`cx:chart` (CT_Chart 2.24.3.13) contains `title` (0..1), `plotArea` (1),
`legend` (0..1) and `extLst`.

`cx:plotArea` (CT_PlotArea 2.24.3.76) contains `plotAreaRegion` (1), `axis`
(0..n), `spPr` and `extLst`.

`cx:plotAreaRegion` (CT_PlotAreaRegion 2.24.3.77) contains `plotSurface`
(0..1) [N: WF2], `series` (0..n) and `extLst`.

Title, legend and axis elements are mapped onto the shared chart model by
stream A's adapter. Only the ChartEx-specific defaults are listed here:

| Element / attribute | Card. | Default | Meaning | Source |
|---|---|---|---|---|
| `title@pos` | opt | `t` | side | [S] XSD default; [N: WF1, BW1…] |
| `title@align` | opt | `ctr` | alignment along the side | [S] |
| `title@overlay` | opt | `0` | title overlaps the plot | [S] |
| `title/tx` | 0..1 | none: **no title text**. Excel writes the text only into `txPr` runs when the title is auto-generated [N: TM, SB, FN1 have `txPr/a:p/a:r/a:t` but no `tx`]. | title text, as `txData` (`f`?, `v`) or `rich` | [S][N: WF1 `txData/v`; FN2 `rich`] |
| `legend@pos` | opt | **`r`** | side | [S] XSD default; [N: WF1 `t`, WF2 `b`] |
| `legend@align`, `@overlay` | opt | `ctr`, `0` | | [S] |
| `axis@id` | req | n/a | axis key that series `axisId` refers to | [S] |
| `axis@hidden` | opt | `0` | the axis is not drawn | [S] |
| `axis/catScaling` or `axis/valScaling` | exactly 1 | n/a | **the axis kind is defined by this child, never by `id` or order.** PA2 and BW2b put `valScaling` on id 0 and `catScaling` on id 1, the reverse of WF1, BW1 and PA1. | [S][N: PA2, BW2b] |
| `catScaling@gapWidth` | opt | none (schema); `auto` allowed | gap as a ratio of the bar width (0.5 = 50 %) | [S][N: WF1 `0.5`, HI1 `0`, FN1 `0.0599999987`] |
| `valScaling@min`/`@max` | opt | `auto` | ST_DoubleOrAutomatic | [S][N: PA1 axis 2 `min="0" max="1"`] |
| `valScaling@majorUnit`/`@minorUnit` | opt | `auto` | | [S] |
| `axis/units@unit` | opt | none | display unit. `percentage` means values are fractions shown ×100 %. | [S][N: PA1, PA2 axis 2] |
| `axis/majorGridlines`, `minorGridlines`, `tickLabels`, `majorTickMarks@type` | 0..1 | absent = not drawn. Tick marks absent = none. | presence toggles | [S][N: WF1 `majorGridlines` + `tickLabels`] |
| `axis/numFmt@formatCode`, `@sourceLinked` | 0..1 | source format | | [S] |

A **chart with no `cx:axis`** (TM, SB, FN1, FN2 has only a cat axis) draws no
axes. [N]

---

## 2. Data model (§4.3 "Data model")

### 2.1 Elements

| Element / attribute | NS | Card. | Default | Meaning | Source |
|---|---|---|---|---|---|
| `chartData` | cx | 1 | n/a | container | [S] 2.24.3.14 |
| `chartData/externalData@r:id` | cx, r | 0..1 | n/a | relationship to the embedded workbook (PPTX/DOCX) | [S] 2.24.3.29; [N: FN2] |
| `externalData@cx:autoUpdate` | cx | opt | none (schema) | refresh from source on open. **Ignore.** The viewer never refreshes. Note that FN2 writes it namespace-qualified, `cx:autoUpdate="0"`, although the XSD declares it unqualified. Accept both. | [S][N: FN2] |
| `chartData/data` | cx | 1..n | n/a | one data set | [S] 2.24.3.20 |
| `data@id` | — | **req** | n/a | ST_DataId (unsignedInt). The key that `series/dataId@val` refers to. Ids are not necessarily dense or ordered: resolve by value, not by position. | [S] 2.24.4.5; [N: BW1 ids 0,1,2; FN2 ids 0,1,2] |
| `data/numDim` \| `data/strDim` | cx | 1..n, any order | n/a | dimensions of this data set | [S] |
| `numDim@type` | — | **req** | n/a | `val` \| `x` \| `y` \| `size` \| `colorVal` (2.24.4.15) | [S][N: `val` WF1, HI1, BW1, FN1; `size` TM, SB; `colorVal` RM] |
| `strDim@type` | — | **req** | n/a | `cat` \| `colorStr` \| `entityId` (2.24.4.23) | [S][N: `cat` TM, SB, FN1, FN2, RM] |
| `*Dim/f` | cx | 1 in form A | n/a | data reference formula (text) | [S] 2.24.3.34 |
| `f@dir` | — | opt | **`col`** | `col` \| `row` orientation of a multi-cell reference | [S] XSD default |
| `*Dim/nf` | cx | 0..1 | n/a | reference to the dimension **name** (header cell) | [S][N: RM] |
| `*Dim/lvl` | cx | 0..n (form A, cache); 1..n (form B, literal) | n/a | one level of points. Form A is `f, nf?, lvl*` (`lvl` is optional cached data). Form B is `lvl+` (literal data, no formula). | [S] 2.24.3.86 / .68 |
| `lvl@ptCount` | — | **req** | n/a | number of points in the level. Points `0..ptCount-1` exist. A missing `pt` is a **blank** (sparse). | [S] 2.24.3.87 / .69; [N: FN2 `ptCount="4"`] |
| `lvl@formatCode` (numeric only) | — | opt | none, so `General` | number format of the cached values | [S][N: FN2 `General`] |
| `lvl@name` | — | opt | none | level name | [S] |
| `lvl/pt@idx` | — | **req** | n/a | 0-based point index. Text is the value: xsd:double for numeric (FN2 `4.2999999999999998`), string for string. | [S] 2.24.3.88 / .70; [N: FN2] |

### 2.2 Value resolution: XLSX has no cache

**Native finding [N: every xlsx fixture]:** Excel writes XLSX dimensions as
`<cx:f>_xlchart.vN.M</cx:f>` with **no `cx:lvl` cache**. `_xlchart.vN.M` is a
hidden workbook defined name (`xl/workbook.xml` `definedName hidden="1"`) whose
value is a sheet range, for example `Sheet1!$C$4:$C$15` (WF1) or
`Table1!$A$2:$A$10` (BW2). `N` varies: `v1` in most files, `v2` in FN1 and `v5`
in RM. The parser must therefore:

1. If the dimension has `lvl` children, use them. This is form B or a form A
   cache [N: FN2].
2. Otherwise, in XLSX, resolve `f` via the workbook's defined names. The value
   is either the name or a direct range, and the cells are read from the
   already-parsed sheet model. A range that crosses sheets, or one that cannot
   be resolved, makes the chart **unsupported** (§4.2), so the fallback is kept.
3. Otherwise (DOCX/PPTX with no cache, or an unreadable embedded workbook),
   the chart is **unsupported**. The embedded workbook is never parsed. [U:
   project decision. FN2 shows that PowerPoint does write the cache.]

A multi-column `f` range (TM/SB `Sheet1!$B$4:$D$14`, under the `col`
direction) yields **one level per column**, with the leftmost column as the
root level (§7). When the cache is used, the order of multiple `lvl` elements
(leaf first or root first) **is unconfirmed: no native part has a multi-level
cache.** [U] The proposed rule follows the classic `c:multiLvlStrRef`
convention: the first `lvl` is the **leaf** level. It needs a native PPTX
treemap or sunburst before it becomes normative.

### 2.3 Series binding

| Element / attribute | Card. | Default | Meaning | Source |
|---|---|---|---|---|
| `series/dataId@val` | 0..1 | none | `data@id` of this series' data set. **Absent is valid**: a `paretoLine` series has no dataId and takes its data from its owner (§5). Any other series with no dataId, or with a dangling id, is unsupported. | [S] 2.24.3.21; [N: PA1, PA2 paretoLine has no dataId] |
| role of each dim | — | — | Within a data set, the dimension is chosen by `@type`, never by position: `val`/`size`/`colorVal` give the values, `cat`/`entityId` give the categories. If there is no `cat` dimension, categories are the 1-based point ordinals [N: WF1, HI1, BW1 have only `numDim val`]. If several dims share a type, the first wins and a diagnostic is raised [U]. | [S][N] |
| point count | — | — | `max(ptCount)` over the bound dims, or the resolved range length. A value point that is blank, non-numeric or missing is **excluded from statistics** (histogram, box, Pareto) and drawn as a gap or zero per family (§§3–8). | [U] |
| `series@layoutId` | **req** | n/a | family selector (§3) | [S] |
| `series@hidden` | opt | `0` | the series is hidden from layout | [S][N: FN2 two series `hidden="1"`] |
| `series@ownerIdx` | opt | none | index of the owning series (paretoLine → its histogram) | [S][N: PA1 `ownerIdx="0"`] |
| `series@uniqueId` | opt | none | GUID string. Not used for rendering. | [S][N: all] |
| `series@formatIdx` | opt | none, so series ordinal | index into the default style palette. Absent: use the 0-based position among series. | [S][N: FN2 0,1,2]; [U] absence rule |
| `series/tx` | 0..1 | none | series name (`txData` with `f`+`v`, or `v`) | [S][N: FN2] |
| `series/axisId` | 0..n | the first cat and first val axis | Axis bindings. **XSD types it as element text (ST_AxisId), but Office writes `<cx:axisId val="1" />`** [N: PA1, PA2]. Accept `@val` first, then the text. | [S][N] |

`ownerIdx` counts `cx:series` elements in document order within the
`plotAreaRegion`, 0-based. [N: PA1 owner is the first series; U as a general
rule]

---

## 3. Layout selection (§4.3 "Layout selection")

`series@layoutId` is ST_SeriesLayout (2.24.4.21). All eight values are confirmed
natively.

| layoutId | Family | Discriminator | Native |
|---|---|---|---|
| `clusteredColumn` | **histogram** when `layoutPr/binning` is present; **pre-aggregated column** (Excel "By Category") when `layoutPr/aggregation` is present or no binning | `binning` / `aggregation` choice (2.24.3.84) | HI1, HI2, PA1, PA2 (binning). `aggregation` is unconfirmed. |
| `paretoLine` | **Pareto** cumulative line. Only valid with `ownerIdx` pointing at a `clusteredColumn` series. | `ownerIdx` | PA1, PA2 |
| `waterfall` | waterfall | | WF1, WF2 |
| `funnel` | funnel | | FN1, FN2 |
| `boxWhisker` | box & whisker | | BW1, BW2 |
| `treemap` | treemap | | TM |
| `sunburst` | sunburst | | SB |
| `regionMap` | region map: **always unsupported** (fallback picture, else card) | | RM |

Family rules:

- A chart's family is the layoutId of its **first non-paretoLine series**.
- A `clusteredColumn` series with a `paretoLine` series owned by it is
  **Pareto**. Without one, it is a **histogram**.
- Mixed families in one plot area (other than clusteredColumn + paretoLine),
  or any `regionMap` series, make the chart unsupported. [U]
- Multiple series of the same family are allowed: BW1 has 3 boxWhisker series
  and FN2 has 3 funnel series (2 of them hidden).

---

## 4. Series layout properties: `cx:layoutPr` (2.24.3.84)

The sequence is `parentLabelLayout?`, `regionLabelLayout?`, `visibility?`,
(`aggregation` | `binning`)?, `geography?`, `statistics?`, `subtotals?` and
`extLst?`. An **absent or empty `layoutPr`** is valid [N: TM `<cx:layoutPr />`;
SB and FN1 have no layoutPr].

### 4.1 Visibility flags: `cx:visibility` (CT_SeriesElementVisibilities, 2.24.3.83)

All attributes are optional xsd:boolean, and **none has a schema default**.

| Attribute | Applies to | Parser default when absent | Meaning [S] | Native |
|---|---|---|---|---|
| `connectorLines` | waterfall | **`1`** [U] | "visibility of connector lines between data points" | WF2 `1`. WF1 has no visibility element. |
| `meanLine` | boxWhisker | **`0`** [U] | "visibility of the line connecting all mean points" | BW1, BW2a `0` |
| `meanMarker` | boxWhisker | **`1`** [U] | "visibility of markers denoting the mean" | BW1, BW2a `1` |
| `nonoutliers` | boxWhisker | **`0`** [U] | "visibility of non-outlier data points" (inner points) | BW1, BW2a `0` |
| `outliers` | boxWhisker | **`1`** [U] | "visibility of outlier data points" | BW1, BW2a `1` |

The absence defaults equal the values Excel writes for a newly inserted chart
(every Excel-written box and whisker writes all four flags explicitly). BW2b
omits `visibility` entirely, and Excel's rendering of that case is
**unconfirmed**. [U]

### 4.2 Data-point and label overrides (§4.3 "data-point overrides")

| Element / attribute | Card. | Default | Meaning | Source |
|---|---|---|---|---|
| `series/spPr` | 0..1 | style palette by `formatIdx` | series fill and line | [S] |
| `series/dataPt@idx` | 0..n, req idx | n/a | per-point `spPr` override. `idx` is the **point index in the bound data**. For waterfall it addresses points, so this is how increase, decrease and total colours are overridden per point. | [S] 2.24.3.26; [N: FN1 idx 2 and 5, effects only] |
| `series/dataLabels` | 0..1 | absent = **no labels** | series label defaults | [S] 2.24.3.24; [N: WF1, FN1, TM, SB] |
| `dataLabels@pos` | opt | none (schema), so family default: `outEnd` for waterfall columns, `ctr` for others [U] | ST_DataLabelPos: `bestFit b ctr inBase inEnd l outEnd r t` | [S][N: WF1 `outEnd`, TM/SB/FN1 `ctr`, PA2 `inEnd`] |
| `dataLabels/visibility@seriesName/@categoryName/@value` | 0..1 | absent attribute = `0` [U] | which fields the label shows | [S] 2.24.3.25; [N: WF1 `0,0,1`; TM `0,1,0`] |
| `dataLabels/separator` | 0..1 | `", "` [U] | text between fields | [S][N: FN1 `", "`] |
| `dataLabels/numFmt`, `spPr`, `txPr` | 0..1 | | label format and text | [S] |
| `dataLabels/dataLabel@idx` (+`@pos`, numFmt, spPr, txPr, visibility, separator) | 0..n | inherit from `dataLabels` field by field | per-point label override | [S] 2.24.3.22; [N: FN1 idx 0,1,3,7,9,11] |
| `dataLabels/dataLabelHidden@idx` | 0..n | n/a | label for point idx is not drawn | [S] 2.24.3.23; [N: FN1 idx 4, 8] |
| `series/valueColors`, `valueColorPositions@count` | 0..1 | count `2` | gradient for `colorVal` (region map only) | [S]; region map is unsupported, so these are ignored |

---

## 5. Histogram and Pareto (§4.3 "Histogram", "Pareto")

### 5.1 `cx:binning` (CT_Binning, 2.24.3.11)

| Element / attribute | Card. | Default | Meaning | Source |
|---|---|---|---|---|
| `binning` | 0..1 (choice with `aggregation`) | n/a | numeric binning of the series values | [S][N: HI1, HI2, PA1, PA2] |
| `binning/binSize` | 0..1 (choice) | n/a | bin width. **The XSD types it as element text (xsd:double), but Office writes `<cx:binSize val="2" />`.** Accept `@val` first, then the text. | [S][N: HI2 `val="2"`] |
| `binning/binCount` | 0..1 (choice) | n/a | number of bins (unsignedInt). Same `@val`/text rule. Per [D-hist], the count **includes** the underflow and overflow bins. | [S][D-hist]; **unconfirmed natively** |
| neither child | — | **automatic** | automatic bin width: Scott's normal reference rule [D-hist "Automatic … calculated using Scott's normal reference rule"] | [N: HI1, PA1, PA2 `<cx:binning intervalClosed="r" />`] |
| `binning@intervalClosed` | opt | none (schema); parser **`r`** | `r` gives bins `(a, b]`, `l` gives bins `[a, b)` (2.24.4.14) | [S][N: `r` in HI1, HI2, PA1, PA2]. `l` is unconfirmed natively. The absence default is [U] (every native writes it). |
| `binning@underflow` | opt | absent = **no underflow bin** | ST_DoubleOrAutomatic. A number creates a bin for values `≤ underflow` [D-hist "below or equal"]. `auto` is an automatic cut-off. | [S][N: HI2 `11`]; `auto` is unconfirmed |
| `binning@overflow` | opt | absent = **no overflow bin** | a bin for values `> overflow` [D-hist "above"]. `auto` is automatic. | [S][N: HI2 `14`]; `auto` is unconfirmed |

Binning algorithm:

1. Take the finite numeric values `v`, excluding blanks and text.
2. Let `lo` be `underflow` if it is set, else `min(v)`. Let `hi` be `overflow`
   if it is set, else `max(v)`.
3. Width: `binSize`, or `(hi−lo)/k` with `k = binCount − (#under/overflow bins
   present)`, or Scott's rule `h = 3.49·σ·n^(−1/3)` (σ is the sample standard
   deviation). Bins start at `lo`. **The exact Scott variant (sample or
   population σ, rounding of `h` and of the bin start) is unconfirmed and needs
   a numeric oracle** [U].
4. Interval closure follows `intervalClosed`. The first inner bin is closed on
   both ends when `r` (`[lo, lo+h]`), matching Excel's label `[a, b]`. [U]
5. Labels use Office interval notation: inner bins `(a, b]` (or `[a, b)`),
   underflow `≤lo`, overflow `>hi`. [D-hist shows these only as images: U on
   exact glyphs and number format]
6. Counts are drawn as columns with `catScaling@gapWidth` (natively `0` for
   histograms).

HI2 worked example (values 12, 10, 11, 13, 15; binSize 2; underflow 11;
overflow 14; `r`): bins `≤11` → {10, 11} = 2, `(11,13]` → {12, 13} = 2,
`>14` → {15} = 1. Whether an empty `(13,15]` bin appears before the overflow bin
is unconfirmed [U: oracle].

Pre-aggregated data ("By Category"): a `clusteredColumn` series with
`aggregation` (or with no binning) is drawn as given, one column per category
point. [S CT_Aggregation 2.24.3.2][D-pareto "By Category"]. **No native part.**
[U]

### 5.2 Pareto

| Element / attribute | Card. | Default | Meaning | Source |
|---|---|---|---|---|
| owner series | 1 | n/a | `clusteredColumn` with `binning`, or by category | [N: PA1, PA2] |
| `paretoLine` series | 1 | n/a | `ownerIdx` points at the owner. **No `dataId`** (the data comes from the owner). Has its own `axisId`. | [S][N: PA1 `<cx:series layoutId="paretoLine" ownerIdx="0">…<cx:axisId val="2" />`] |
| percent axis | 1 | n/a | `valScaling min="0" max="1"` + `units unit="percentage"`. The line is plotted as a **fraction 0..1** and labelled 0 %–100 %. | [N: PA1, PA2 axis 2] |

Rules:

- Bars (bins, or categories) are sorted **descending** by height [D-pareto
  "columns sorted in descending order"].
- Ties keep their pre-sort order (the order of bins, or the source order for
  categories). [U]
- The line point *i* is `Σ_{j≤i} bar_j / Σ bar` [D-pareto "cumulative total
  percentage"].
- The secondary axis is the axis that `paretoLine/axisId` names. [N]
- PA2's `clusteredColumn` series carries `<cx:parentLabelLayout val="banner" />`,
  which has no meaning for columns. Ignore it. [N]

---

## 6. Waterfall (§4.3 "Waterfall")

| Element / attribute | Card. | Default | Meaning | Source |
|---|---|---|---|---|
| `subtotals` | 0..1 | absent or empty = no subtotal points | list of subtotal points | [S] 2.24.3.90; [N: WF1, WF2 **empty** `<cx:subtotals />`] |
| `subtotals/idx@val` | 0..n | n/a | 0-based index of a point that is a "total". **Unconfirmed natively: both natives have an empty list.** | [S] 2.24.3.89 |
| `visibility@connectorLines` | opt | `1` [U] | connector lines (§4.1) | [S][N: WF2] |

Rules:

- Running total: `acc₀ = 0`. A regular point *i* is drawn as a floating bar
  from `acc` to `acc + vᵢ`, then `acc += vᵢ`. It is an increase if `vᵢ ≥ 0`,
  else a decrease. [D-wf]
- A **subtotal point** (its idx is in `subtotals`) is drawn **from 0 to the
  bar's value, anchored on the horizontal axis** [D-wf "start on the horizontal
  axis at zero and don't float"]. **Open question:** whether the bar height is
  the running total `acc` (§4.3) or the point's own cell value `vᵢ`. Excel's
  "Set as total" is normally applied to a cell that already holds the running
  total, so the two coincide for well-formed sheets. The parser draws `acc` and
  then sets `acc` to the drawn value. [U: needs an oracle where `vᵢ ≠ acc`]
- Colours: increase, decrease and total use the style palette (default
  accents 1, 2 and 3 for the three legend entries Increase, Decrease and Total
  [U]). A `dataPt@idx` `spPr` overrides one bar.
- Connector lines join the end of bar *i* to the start of bar *i+1* when
  `connectorLines` is true [D-wf].
- Categories: the `cat` dimension if present, else ordinals [N: WF1, WF2 have
  no `cat` dimension].

---

## 7. Treemap and sunburst (§4.3 "Treemap and sunburst")

| Element / attribute | Card. | Default | Meaning | Source |
|---|---|---|---|---|
| `strDim type="cat"` (multi-level) | 1 | n/a | category path, one level per column (root = leftmost) | [N: TM, SB `_xlchart.v1.0` = `Sheet1!$B$4:$D$14`] |
| `numDim type="size"` | 1 | n/a | leaf weight | [N: TM, SB] |
| `parentLabelLayout@val` | 0..1 | **unconfirmed** [U]. The proposed treemap default is `overlapping` (the Excel UI default) and is unverified. Sunburst ignores it. | `none` (no parent labels), `banner` (banner above the category), `overlapping` (label laid out within the category) (2.24.4.17) | [S]; TM writes `<cx:layoutPr />` (absent); SB has no layoutPr |

### Hierarchy derivation (from the TM/SB sheet data; native)

TM/SB rows (B, C, D | E):

```
Best  First  A | 3        Worst Third  E | 4
      .      B | 4              Fourth . | 3
      .      C | 2              Fifth  . | 4
      .      D | 3              Sixth  F | 5
      .      E | 4
      Second C | 5
      .      D | 6
```

1. **Carry-forward rule:** a blank cell in a non-leaf column, followed by a
   non-blank cell further right in the same row, inherits the value from the
   row above in that column. (Row 5: blank B and C, then D=B, gives the path
   Best/First/B.)
2. **Short-path rule:** a row whose trailing columns are blank is a **leaf at
   its last non-blank level**. (Row 12: B blank → Worst, C=Fourth, D blank, so
   the leaf is Worst/Fourth with weight 3.) Carry-forward never fills the
   trailing blanks.
3. Node identity is the **full path** (all levels), so `C` under First and `C`
   under Second are different nodes.
4. Duplicate full paths are summed. [U: there is no native example]
5. A weight that is ≤ 0, blank or non-numeric excludes its row, and a
   diagnostic is raised. [U]
6. A parent's weight is the sum of its descendants.

Layout: the treemap uses squarified rectangles. The sunburst draws rings by
depth (the root level is the inner ring), with angle proportional to subtree
weight in source order. Office does not document either algorithm, so pixel
layout is not claimed to match; only the area and angle proportions are
oracles. [U]

---

## 8. Box & whisker (§4.3 "Box & whisker")

| Element / attribute | Card. | Default | Meaning | Source |
|---|---|---|---|---|
| `statistics@quartileMethod` | 0..1 | none (schema); parser **`exclusive`** [U] | `inclusive`: "includes the median when splitting the dataset into quartiles". `exclusive`: "excludes the median when splitting". (2.24.4.19) | [S][D-box "Inclusive median … included … if N is odd"][N: BW1, BW2 `exclusive`]; `inclusive` is unconfirmed natively |
| `visibility@meanLine/meanMarker/nonoutliers/outliers` | opt | §4.1 | | [S][N: BW1] |
| data | 1 per series | n/a | one box per series. BW1 has 3 series with `numDim val` each and **no `cat` dimension**, so the box labels come from ordinals or series names. | [N: BW1, BW2a] |

Statistics per series (values: finite numbers only, sorted, n = count):

- **Median**: the usual median.
- **Q1, Q3**: per ST_QuartileMethod, when *n* is odd the median is
  **excluded** (`exclusive`) or **included** (`inclusive`) in both halves. Q1 is
  the median of the lower half and Q3 the median of the upper half. When *n* is
  even, the halves are the lower and upper *n/2*, and both methods agree. [S][D-box]
- **Open question:** whether Excel computes the half-medians as above (§4.3's
  "not QUARTILE.INC/EXC interpolation") or uses QUARTILE.EXC/INC interpolation.
  **BW1 cannot discriminate.** For series 1 (3,4,4,5,6,6,7,8,9), the
  exclusive half-median and QUARTILE.EXC both give Q1 = 4 and Q3 = 7.5. The
  inclusive half-median and QUARTILE.INC both give Q1 = 4 and Q3 = 7. A native oracle
  with, for example, n = 6 or 10 and an Excel-rendered fallback picture is
  required. [U]
- **Whiskers** reach the furthest data point within `[Q1 − 1.5·IQR, Q3 + 1.5·IQR]`.
  Points beyond are **outliers**, drawn only when `outliers` is set [D-box
  "any point outside those lines or whiskers is considered an outlier"]. The
  1.5 factor is **not stated by Microsoft** [U].
- The **mean marker** (×) is drawn when `meanMarker` is set. The **mean line**
  connects the means of consecutive series when `meanLine` is set [S wording].
  **Inner points** are drawn when `nonoutliers` is set.
- A series with no numeric values draws no box [N: BW2b references an empty
  range, `Sheet1!$D$5:$D$13`], and the chart is still supported. [U render]

---

## 9. Funnel (§4.3 "Funnel")

| Item | Rule | Source |
|---|---|---|
| data | `strDim cat` + `numDim val` | [N: FN1, FN2] |
| order | source order, from top to bottom | [U] |
| bar | horizontally **centred**. Width is proportional to value / max(value). Gap is `catScaling@gapWidth`. | [U]; [N: FN1 gap `0.06`, FN2 `2.19`] |
| hidden series | a `series@hidden="1"` series is not drawn | [S][N: FN2] |
| multiple visible series | unsupported (Office draws only one funnel) | [U] |
| negative or zero values | width 0, and the label is still drawn | [U] |
| axis | only a category axis (FN2 has `axis id="1"` with `catScaling`; FN1 likewise) | [N] |

---

## 10. Region map

`layoutId="regionMap"` (RM: `strDim cat` + `numDim colorVal`, both with `nf`;
`layoutPr/geography` with `cultureLanguage`, `cultureRegion`, `attribution` and
`geoCache provider` holding base64 `binary`). It always returns
**unsupported**, so the `mc:Fallback` picture is selected, else the card. The
parser must recognise it **before** reading geography and must never decode
`cx:binary`. [S 2.24.3.50][N: RM]

---

## 11. Unsupported conditions (all return "unsupported" under §6.3)

- The part cannot be parsed, its root is not `cx:chartSpace`, or it has no
  `plotAreaRegion/series`.
- An unknown `layoutId`, mixed families, or any `regionMap`.
- A `dataId` that resolves to no `data`, or a required dimension (§§5–9) that
  is missing.
- XLSX: an `f` that does not resolve to a single-sheet range. DOCX/PPTX: no
  `lvl` cache.
- Analytic inputs over the §7 limit (spec: never truncated).

## 12. Open questions for reviewers

1. **Multi-level `lvl` cache order** (§2.2): no native part has one. A PPTX
   treemap or sunburst saved by PowerPoint is needed.
2. **Quartile computation** (§8): half-medians or interpolation. BW1 cannot
   discriminate, so an oracle with even *n* is needed.
3. **Waterfall subtotal height** (§6): `acc` or the cell value. There is no
   native `subtotals/idx`.
4. **Histogram**: the Scott's-rule variant and bin alignment. `binCount`,
   `intervalClosed="l"` and `underflow/overflow="auto"` have no native example.
   Bin-label glyphs are unconfirmed.
5. **Absence defaults** for `connectorLines`, the box visibility flags,
   `quartileMethod`, `intervalClosed` and `parentLabelLayout`: MS-ODRAWXML gives
   none, and every native writes them explicitly. The proposed defaults are
   Excel's new-chart values.
6. **By-category histogram / Pareto (`aggregation`)**: no native part.
7. **Schema/native mismatches** to accept: `binSize`/`binCount`/`axisId` are
   written as `@val` attributes rather than text, and `externalData@cx:autoUpdate`
   is written namespace-qualified.
8. **Fixture licence** (PROVENANCE.md): MPL-2.0 per LibreOffice policy. Owner
   confirmation is needed.
