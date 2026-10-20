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
}

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
  name: string
  /** CSS hex without `#`. */
  color: string
  /** Per-series categories when they differ from the model's shared list. */
  categories?: string[]
  /** Sparse cache indices are kept: a missing `pt@idx` is a `blank` point at its index (spec 6.4). */
  values: ChartPoint[]
  xValues?: ChartPoint[]
  bubbleSizes?: ChartPoint[]
  /** Per-point colours (pie slices, varyColors). */
  pointColors?: string[]
  line?: boolean
  markers?: boolean
  /** References still to be filled from worksheet cells (no saved data for that role). */
  pending?: { name?: ChartRange; values?: ChartRange; categories?: ChartRange; xValues?: ChartRange }
  /**
   * Point indices whose worksheet cell sits in a hidden row or column (provenance
   * kept at fill for `plotVisOnly`; applied at draw by `displayModel`). Cached data has none.
   */
  hidden?: number[]
  /**
   * The series feeds an analytic element (a trendline, stdDev/stdErr error bars):
   * its numeric roles are acquired in full up to CHART_ANALYTIC_POINT_LIMIT (spec 7),
   * and only the drawing is capped at CHART_POINT_LIMIT. `refused`: an input
   * exceeded that bound, so the analytic element is dropped (never computed from
   * a prefix) and the series is acquired for display only.
   */
  analytic?: true | 'refused'
}

export interface ChartGroup {
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
}

export interface ChartModel {
  title?: string
  groups: ChartGroup[]
  /** Every axis of the plot area, in document order; group bindings reference them by `id`. */
  axes: ChartAxis[]
  categories: string[]
  /** The shared category reference still to be filled from worksheet cells. */
  categoryPending?: ChartRange
  /** Category indices whose cell sits in a hidden row or column (see `ChartSeries.hidden`). */
  categoryHidden?: number[]
  legend?: 'r' | 'l' | 't' | 'b' | 'tr'
  fontFamily: string
  fontSizePt: number
  /** Chart-level options (spec 5.8); Office defaults: gap, plot visible only. */
  dispBlanksAs?: 'gap' | 'zero' | 'span'
  plotVisOnly?: boolean
  /** `c16r3:dataDisplayOptions16/dispNaAsBlank`. */
  dispNaAsBlank?: boolean
  /** The chart epoch (spec 5.1). */
  epoch?: '1900' | '1904'
  /** Display points dropped by the per-chart bound. */
  truncated?: boolean
  /** Set when the chart cannot be drawn: the reason shown on the placeholder card. */
  unsupported?: string
  /** A card's original feature when it is not a classic chart (spec 6.3). */
  feature?: 'ChartEx'
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
