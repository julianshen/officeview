/**
 * Format-neutral chart model (types only; spec 5.1, 5.8, 6.3).
 *
 * The engine's one model (D0): every producer (the chart part reader, the
 * resident and paged XLSX fills) and consumer (painter, description, cards)
 * uses it; the legacy `Xlsx*` names in `src/xlsx/chartModel.ts` are aliases.
 * D0 records each group's axis binding structurally from its `c:axId`s; A1a
 * validates it and drives the axis graph from it.
 * The runtime tag of a chart drawing stays `kind:'xlsxChart'` (spec 6.2).
 */

/**
 * One numeric plotting input (values, x values, bubble sizes, error amounts;
 * spec 5.8). `error` keeps its code (`#N/A` distinct from the rest) until the
 * family's plotting rule; `unavailable` is a saved value that does not exist
 * (a formula cell with no cached value), distinct from `blank`.
 */
export type ChartPoint =
  | { state: 'value'; v: number; fmt?: string }
  | { state: 'blank' }
  | { state: 'error'; code: string }
  | { state: 'unavailable' }

/**
 * One numeric role's points, stored compactly (no object per point; see
 * `charts/points.ts`): `states[i]` is 0 blank, 1 value, 2 error, 3 unavailable;
 * `values[i]` is the number of a value point, or an error point's code index
 * (into CHART_ERROR_CODES, then `codes`). `formats`/`formatIndex` keep
 * per-point `pt@formatCode` (index 0: none). `hidden` lists, sorted, the point
 * indices whose cell sits in a hidden row or column of this role's own source
 * range (plotVisOnly provenance; spec 5.8). Sparse cache indices are kept: a
 * missing `pt@idx` is a blank point at its index (spec 6.4).
 */
export interface ChartPoints {
  length: number
  states: Uint8Array
  values: Float64Array
  codes?: string[]
  formats?: string[]
  formatIndex?: Uint16Array
  hidden?: Uint32Array
}

/** The data roles of a series (spec 6.4). Names and categories are strings; the rest are `ChartPoint`s. */
export type ChartRole = 'name' | 'categories' | 'values' | 'xValues' | 'bubbleSizes' | 'errorPlus' | 'errorMinus'

/** Chart families of streams A and B (spec 4). */
export type ChartFamily =
  | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'ofPie' | 'scatter' | 'bubble' | 'radar' | 'stock' | 'contour'
  | 'waterfall' | 'histogram' | 'pareto' | 'boxWhisker' | 'funnel' | 'treemap' | 'sunburst'

export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard'

/**
 * How a group binds its axes, by `axId` (spec 5.1). Pie and doughnut use
 * `none`; radar `radial`; contour (the 2D surface) `surface`; scatter and bubble
 * `xy` (two value axes). A `cartesian` group has exactly one of `category`
 * (a category or date axis) or `x` (a value axis); see `validateAxisBinding`.
 */
export type ChartAxisBinding =
  | { kind: 'none' }
  | { kind: 'cartesian'; category?: string; x?: string; value: string }
  | { kind: 'xy'; x: string; y: string }
  | { kind: 'radial'; value: string }
  | { kind: 'surface'; category: string; series: string; value: string }

/** Placement-only text properties (DrawingML txPr). Rotation is degrees clockwise. */
export interface ChartTextStyle {
  sizePt?: number; rotation?: number; vertical?: string
  /** Run properties (ChartEx `txPr` defRPr or first run; CSS hex without `#`). */
  color?: string; bold?: boolean; italic?: boolean
}
export interface ChartManualLayout {
  x?: number; y?: number; w?: number; h?: number
  xMode?: 'edge' | 'factor'; yMode?: 'edge' | 'factor'; wMode?: 'edge' | 'factor'; hMode?: 'edge' | 'factor'
  target?: 'inner' | 'outer'
}

/** An axis keyed by its `axId` (spec 5.1). A deleted axis still defines the scale. */
export interface ChartAxis {
  id: string
  kind: 'category' | 'value' | 'date' | 'series'
  /** The `axId` of the axis this one crosses. */
  crossAx?: string
  position?: 'b' | 't' | 'l' | 'r'
  crosses?: 'autoZero' | 'min' | 'max'
  crossesAt?: number
  deleted: boolean
  gridlines: boolean
  minorGridlines?: boolean
  min?: number
  max?: number
  majorUnit?: number
  minorUnit?: number
  logBase?: number
  reversed?: boolean
  crossBetween?: 'between' | 'midCat'
  title?: string
  /** Number format and whether it follows the source (spec 5.2). */
  formatCode?: string
  sourceLinked?: boolean
  /** Tick marks and label placement as authored (absent: the legacy drawing, no marks, labels at the plot edge). */
  majorTickMark?: ChartTickMark
  minorTickMark?: ChartTickMark
  tickLblPos?: 'nextTo' | 'low' | 'high' | 'none'
  text?: ChartTextStyle
  tickLblSkip?: number
  tickMarkSkip?: number
  /** `c:dispUnits`: tick values are divided by `divisor`; `label` is the unit label drawn beside the axis (spec 5.1). */
  dispUnits?: { divisor: number; unit: string; label?: string }
  /** `c:dateAx` time units (spec 5.1); `auto` (`c:auto`) lets text categories fall back to a category axis. */
  baseTimeUnit?: ChartTimeUnit
  majorTimeUnit?: ChartTimeUnit
  minorTimeUnit?: ChartTimeUnit
  auto?: boolean
  /** No axis element: the plot area had none and its groups no `c:axId`, so a primary pair is implied (A1a). */
  implicit?: true
  /** The axis title's text style (ChartEx `axis/title/txPr`). */
  titleText?: ChartTextStyle
  /** ChartEx `cx:units@unit` (chartex-map §1): `percentage` means the values are fractions shown x100 %. */
  unit?: string
}
export type ChartTickMark = 'cross' | 'in' | 'out' | 'none'
export type ChartTimeUnit = 'days' | 'months' | 'years'
/** A chart diagnostic (spec 6.6): a code and its reason. */
export interface ChartDiagnostic { code: string; reason: string }

/**
 * A single-area worksheet reference, 0-based and inclusive, that a series role
 * reads when its saved data (cache or literal) is absent (spec 6.4 precedence 2).
 */
export interface ChartRange {
  sheet?: string; r0: number; c0: number; r1: number; c1: number
  /** A workbook defined name still to resolve (the coordinates are then 0). */
  definedName?: string
}

export interface ChartSeries {
  labels?: import('./labels').LabelControls
  name: string
  /** CSS hex without `#`. */
  color: string
  /** Per-series categories when they differ from the model's shared list. */
  categories?: string[]
  values: ChartPoints
  xValues?: ChartPoints
  bubbleSizes?: ChartPoints
  /** Per-point colours (pie slices, varyColors). */
  pointColors?: string[]
  line?: boolean
  markers?: boolean
  fill?: import('./fills').ChartFillStyle
  lineStyle?: import('./style').ChartLineStyle
  markerStyle?: import('./markers').ChartMarkerStyle
  invertIfNegative?: boolean
  /** The fill of inverted negative points (c14 `invertSolidFillFmt`); absent: white, as Office. */
  invertColor?: string
  /** Percent of the radius, clamped 0-400. */
  explosion?: number
  dPt?: Map<number, import('./style').ChartPointOverride>
  pointFills?: Array<import('./fills').ChartFillStyle | undefined>
  /** References still to be filled from worksheet cells (no saved data for that role). */
  pending?: { name?: ChartRange; values?: ChartRange; categories?: ChartRange; xValues?: ChartRange }
  /**
   * Each role's worksheet reference, kept even when saved data (a cache) supplies
   * its values: the cache decides the values, the reference still decides which
   * points sit in hidden rows or columns (plotVisOnly provenance; spec 5.8, 6.4).
   */
  refs?: { values?: ChartRange; xValues?: ChartRange; categories?: ChartRange }
  /** Roles with neither saved data nor a reference (spec 6.4 precedence 3): unavailable, never blank. */
  absent?: ChartRole[]
  /**
   * The series feeds an analytic element (a trendline, stdDev/stdErr error bars):
   * its numeric roles are acquired in full up to CHART_ANALYTIC_POINT_LIMIT (spec 7),
   * and only the drawing is capped at CHART_POINT_LIMIT. `refused`: an input
   * exceeded that bound, so the analytic element is dropped (never computed from
   * a prefix) and the series is acquired for display only. `chart`: the series is the input of an analytic
   * ChartEx family (histogram, Pareto, box & whisker, treemap, sunburst): every role, categories included, is
   * acquired in full up to CHART_ANALYTIC_POINT_LIMIT on every path (cache, defined name, host sheet, embedded
   * workbook); above it, or with a required value point unavailable, the whole chart is refused.
   */
  analytic?: true | 'refused' | 'chart'
}

export interface ChartGroup {
  /** The group's kind: absent for classic (`c:`) groups; `chartEx` for a ChartEx group (`ChartExGroup`). */
  kind?: 'chartEx'
  labels?: import('./labels').LabelControls
  family: ChartFamily
  grouping: ChartGrouping
  barDir?: 'col' | 'bar'
  series: ChartSeries[]
  binding: ChartAxisBinding
  gapWidth: number
  overlap: number
  holeSize?: number
  firstSliceAngle?: number
  varyColors?: boolean
  dropLines?: import('./style').ChartLineStyle
  serLines?: import('./style').ChartLineStyle
}

/**
 * ChartEx (`cx:`, Office 2016+) groups, read against docs/charts/chartex-map.md (spec 4.2, 4.3). A ChartEx chart is
 * the same `ChartModel`: its one group is a `ChartExGroup`, whose `series` carry the bound roles (values,
 * categories, names; cached, or pending references for the workbook) and whose `chartEx` payload keeps the
 * ID-keyed data sets (all levels) and the per-series layout properties the family painters need.
 */
export type ChartExLayout = 'clusteredColumn' | 'paretoLine' | 'waterfall' | 'funnel' | 'boxWhisker' | 'treemap' | 'sunburst' | 'regionMap'
export type ChartExFamily = 'waterfall' | 'histogram' | 'pareto' | 'boxWhisker' | 'funnel' | 'treemap' | 'sunburst'
/** One `cx:lvl` (map §2.1): points `0..ptCount-1`; a missing `pt@idx` is blank (numeric) or `undefined` (string). */
export type ChartExLevel = { ptCount: number; name?: string; formatCode?: string } &
  ({ kind: 'num'; points: ChartPoints } | { kind: 'str'; values: Array<string | undefined> })
/** A `cx:numDim`/`cx:strDim`: its `@type`, its reference (`cx:f`; a defined name or a single-area range) and its levels. */
export interface ChartExDimension { kind: 'num' | 'str'; type: string; ref?: ChartRange; dir: 'col' | 'row'; levels: ChartExLevel[] }
/** A `cx:data` set, keyed by `@id` (resolved by value, never by position). */
export interface ChartExData { id: number; dimensions: ChartExDimension[] }
export interface ChartExBinning { intervalClosed: 'l' | 'r'; binSize?: number; binCount?: number; underflow?: number | 'auto'; overflow?: number | 'auto' }
/** `cx:layoutPr` (map §4), with the map's absence defaults applied. */
export interface ChartExLayoutProperties {
  visibility: { connectorLines: boolean; meanLine: boolean; meanMarker: boolean; nonoutliers: boolean; outliers: boolean }
  binning?: ChartExBinning
  /** `cx:aggregation`: by-category columns (equal categories grouped, values summed). */
  aggregation?: true
  /** `cx:subtotals/cx:idx@val`: 0-based subtotal point indices. */
  subtotals?: number[]
  quartileMethod?: 'inclusive' | 'exclusive'
  parentLabelLayout?: 'none' | 'banner' | 'overlapping'
}
/** One `cx:series` as authored (map §2.3). */
export interface ChartExSeriesInfo {
  layout: ChartExLayout
  dataId?: number
  ownerIdx?: number
  hidden: boolean
  formatIdx: number
  axisIds: string[]
  layoutPr: ChartExLayoutProperties
}
export interface ChartExGroup extends ChartGroup {
  kind: 'chartEx'
  family: ChartExFamily
  chartEx: {
    /** The data sets the series bind (by `dataId`). */
    data: ChartExData[]
    /** Parallel to the group's `series`. */
    series: ChartExSeriesInfo[]
    /** Pareto: the `paretoLine` series (no data of its own; bound through `ownerIdx`). */
    paretoLine?: ChartExSeriesInfo
    /** By-category histogram (Excel "By Category"): the column series carries `cx:aggregation`; equal categories
     * are grouped and their values summed (map §3, §5.1). */
    byCategory?: true
    /** `cx:plotArea/cx:spPr` (the whole plot area; the inner `cx:plotSurface` is the model's `plotAreaFill`). */
    plotAreaFill?: import('./fills').ChartFillStyle
    plotAreaLine?: import('./style').ChartLineStyle
    /** The category axis `catScaling@gapWidth` (a ratio of the bar width; absent: not authored). */
    gapWidth?: number
  }
}

export interface ChartModel {
  /** Immutable, transferable stage-1 domains, ticks and label strings. */
  preparedAxes?: import('./prepare').PreparedAxes
  plotLayout?: ChartManualLayout
  titleLayout?: ChartManualLayout
  legendLayout?: ChartManualLayout
  text?: ChartTextStyle
  title?: string
  /** The title's text style (ChartEx `title/txPr`). */
  titleText?: ChartTextStyle
  chartAreaFill?: import('./fills').ChartFillStyle
  chartAreaLine?: import('./style').ChartLineStyle
  plotAreaFill?: import('./fills').ChartFillStyle
  plotAreaLine?: import('./style').ChartLineStyle
  legendFill?: import('./fills').ChartFillStyle
  legendLine?: import('./style').ChartLineStyle
  titleFill?: import('./fills').ChartFillStyle
  titleLine?: import('./style').ChartLineStyle
  styleId?: number
  /** The chart-area fill when `chartSpace/spPr` has none and `c:style` gives a non-white default (DOCX/PPTX; spec 5.4). */
  chartAreaDefault?: string
  groups: ChartGroup[]
  /** Every axis of the plot area, in document order; group bindings reference them by `id`. */
  axes: ChartAxis[]
  categories: string[]
  /**
   * The numeric reading of the shared categories when they sit on a date axis (spec 5.1): serials in the chart
   * epoch (a filled serial from another epoch is converted, or unavailable when it has no counterpart). A label
   * without a number here is text: the date axis then falls back to a category axis.
   */
  categoryValues?: ChartPoints
  /** The shared category reference still to be filled from worksheet cells. */
  categoryPending?: ChartRange
  /** The shared category reference (kept beside a cache, for provenance only). */
  categoryRef?: ChartRange
  /** Sorted category indices whose cell sits in a hidden row or column of the category reference. */
  categoryHidden?: Uint32Array
  legend?: 'r' | 'l' | 't' | 'b' | 'tr'
  fontFamily: string
  fontSizePt: number
  /** Chart-level options (spec 5.8); Office defaults: gap, plot visible only. */
  dispBlanksAs?: 'gap' | 'zero' | 'span'
  plotVisOnly?: boolean
  /** `c16r3:dataDisplayOptions16/dispNaAsBlank`. */
  dispNaAsBlank?: boolean
  /** The chart epoch as authored (`c:date1904`; spec 5.1). */
  epoch?: '1900' | '1904'
  /**
   * The containing workbook's epoch (spec 5.1 precedence 2), carried on the transferable model so every paint
   * realm resolves the same chart epoch. Stamped by the XLSX fills for a 1904 workbook; absent means 1900.
   */
  workbookEpoch?: '1904'
  /** Display points dropped by the per-chart bound. */
  truncated?: boolean
  /** Diagnostics of the parse and fill (spec 6.6), e.g. `chart-axis-unplaceable`, `chart-date-unrepresentable`. */
  diagnostics?: ChartDiagnostic[]
  /** Set when the chart cannot be drawn: the reason shown on the placeholder card. */
  unsupported?: string
  /** A card's original feature when it is not a classic chart (spec 6.3). */
  feature?: 'ChartEx'
  /** The chart's storage could not be admitted (a typed refusal before allocation; spec 6.4/6.5). */
  refused?: true
}

/**
 * The representation an adapter returns (spec 6.3). `unsupported` and
 * `refused` never satisfy an `mc:Choice`; a card is made only when no
 * representation remains. The runtime tag stays `xlsxChart`.
 */
export type ChartResult<Model = ChartModel> =
  | { kind: 'native'; model: Model }
  | { kind: 'unsupported'; reason: string; title?: string; feature?: 'ChartEx' }
  | { kind: 'refused'; reason: string; title?: string; feature?: 'ChartEx' }
