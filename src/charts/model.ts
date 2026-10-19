/**
 * Format-neutral chart model (types only; spec 5.1, 5.8, 6.3).
 *
 * F0 adds these beside the legacy `Xlsx*` model in `src/xlsx/chartModel.ts`,
 * which still drives the engine. D0 migrates producers and consumers to this
 * shape in one step; A1a populates and validates the axis bindings.
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

export interface ChartSeries {
  name: string
  /** CSS hex without `#`. */
  color?: string
  /** Per-series categories when they differ from the model's shared list. */
  categories?: string[]
  values: ChartPoint[]
  xValues?: ChartPoint[]
  bubbleSizes?: ChartPoint[]
  /** Per-point colours (pie slices, varyColors). */
  pointColors?: string[]
  line?: boolean
  markers?: boolean
}

export interface ChartGroup {
  family: ChartFamily
  grouping: ChartGrouping
  barDir?: 'col' | 'bar'
  series: ChartSeries[]
  binding: ChartAxisBinding
  gapWidth?: number
  overlap?: number
  holeSize?: number
  firstSliceAngle?: number
  varyColors?: boolean
}

export interface ChartModel {
  title?: string
  groups: ChartGroup[]
  /** Every axis of the plot area; group bindings reference them by `id`. */
  axes: ChartAxis[]
  categories: string[]
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
}

/**
 * The representation an adapter returns (spec 6.3). `unsupported` and
 * `refused` never satisfy an `mc:Choice`; a card is made only when no
 * representation remains. `Model` defaults to `ChartModel`; until D0 the
 * engine still returns the legacy `XlsxChartModel` (runtime tag `xlsxChart`).
 */
export type ChartResult<Model = ChartModel> =
  | { kind: 'native'; model: Model }
  | { kind: 'unsupported'; reason: string; title?: string }
  | { kind: 'refused'; reason: string; title?: string }
