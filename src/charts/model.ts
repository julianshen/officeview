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
  | { kind: 'radial'; value: string; category?: string }
  | { kind: 'surface'; category: string; series: string; value: string }

/** Placement-only text properties (DrawingML txPr). Rotation is degrees clockwise. */
export interface ChartTextStyle { sizePt?: number; rotation?: number; vertical?: string }
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
  smooth?: boolean
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
  /** `c:trendline`s as authored (spec 5.6); fitted in stage 1 (`PreparedAxes.trendlines`). */
  trendlines?: import('./trendlines').ChartTrendline[]
  /** References still to be filled from worksheet cells (no saved data for that role). */
  pending?: { name?: ChartRange; values?: ChartRange; categories?: ChartRange; xValues?: ChartRange; bubbleSizes?: ChartRange }
  /**
   * Each role's worksheet reference, kept even when saved data (a cache) supplies
   * its values: the cache decides the values, the reference still decides which
   * points sit in hidden rows or columns (plotVisOnly provenance; spec 5.8, 6.4).
   */
  refs?: { values?: ChartRange; xValues?: ChartRange; categories?: ChartRange; bubbleSizes?: ChartRange }
  /** Roles with neither saved data nor a reference (spec 6.4 precedence 3): unavailable, never blank. */
  absent?: ChartRole[]
  /**
   * The series feeds an analytic element (a trendline, stdDev/stdErr error bars):
   * its numeric roles are acquired in full up to CHART_ANALYTIC_POINT_LIMIT (spec 7),
   * and only the drawing is capped at CHART_POINT_LIMIT. `refused`: an input
   * exceeded that bound, so the analytic element is dropped (never computed from
   * a prefix) and the series is acquired for display only.
   */
  analytic?: true | 'refused'
}

export interface ChartUpDownBar {
  fill?: import('./fills').ChartFillStyle
  line?: import('./style').ChartLineStyle
}

export interface ChartUpDownBars {
  upBars?: ChartUpDownBar
  downBars?: ChartUpDownBar
}

export interface ChartGroup {
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
  /** Contour: `c:wireframe` (band boundaries only). */
  wireframe?: boolean
  /** Contour `c:bandFmts` by band index: the fill colour (hex), null for noFill; holes take the accent cycle. */
  bandColors?: Array<string | null>
  /** Contour `c:bandFmts` by band index: the line colour (hex) a wireframe strokes, null for noFill. */
  bandLines?: Array<string | null>
  /** Contour: the document's accent cycle (hex), resolved at parse time, for bands without an authored colour. */
  bandAccents?: string[]
  varyColors?: boolean
  smooth?: boolean
  dropLines?: import('./style').ChartLineStyle
  serLines?: import('./style').ChartLineStyle
  /** Bubble (spec 4.1): `c:bubbleScale` percent (0-300, default 100), `c:sizeRepresents`, `c:showNegBubbles`. */
  bubbleScale?: number
  sizeRepresents?: 'area' | 'w'
  showNegBubbles?: boolean
  /** Radar (spec 4.1): `c:radarStyle`. */
  radarStyle?: 'standard' | 'marker' | 'filled'
  /** Stock (spec 4.1): `c:hiLowLines` and `c:upDownBars`. */
  hiLowLines?: import('./style').ChartLineStyle
  upDownBars?: ChartUpDownBars
}

export interface ChartModel {
  /** Immutable, transferable stage-1 domains, ticks and label strings. */
  preparedAxes?: import('./prepare').PreparedAxes
  plotLayout?: ChartManualLayout
  titleLayout?: ChartManualLayout
  legendLayout?: ChartManualLayout
  text?: ChartTextStyle
  title?: string
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
  /** Drawing overlays from `c:userShapes` (spec 5.11; charts/userShapes.ts), prepared at stage 1. */
  userShapes?: import('./userShapes').ChartUserShapes
}

/**
 * The representation an adapter returns (spec 6.3). `unsupported` and
 * `refused` never satisfy an `mc:Choice`; a card is made only when no
 * representation remains. The runtime tag stays `xlsxChart`.
 */
export type ChartResult<Model = ChartModel> =
  | { kind: 'native'; model: Model }
  | { kind: 'unsupported'; reason: string; title?: string }
  | { kind: 'refused'; reason: string; title?: string }
