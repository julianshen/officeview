/**
 * Worksheet chart model (types only; the parser and painter live in the lazy
 * `chart.ts` module, loaded only when a sheet has a chart).
 *
 * Values come from the chart part's caches (c:numCache/c:strCache/literals).
 * A series whose reference carries no cache keeps the reference so the
 * workbook can fill it from its cells (openpyxl writes no caches; Excel
 * recomputes from the sheet).
 *
 * The format-neutral contracts (`ChartModel`, `ChartPoint`, axis bindings,
 * `ChartResult`) are re-exported beside the legacy `Xlsx*` types; D0 moves
 * the engine onto them.
 */
export type * from '../charts/model'
export type XlsxChartType = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter'
export type XlsxChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard'

/** A single-area worksheet reference, 0-based and inclusive. */
export interface XlsxChartRange {
  sheet?: string; r0: number; c0: number; r1: number; c1: number
  /** A workbook defined name still to resolve (the coordinates are then 0). */
  definedName?: string
}

export interface XlsxChartSeries {
  name: string
  /** CSS hex without `#`. */
  color: string
  values: Array<number | undefined>
  /** Scatter x values. */
  xValues?: Array<number | undefined>
  /** Per-point colours (pie slices, varyColors). */
  pointColors?: string[]
  /** References still to be filled from worksheet cells. */
  pending?: { name?: XlsxChartRange; values?: XlsxChartRange; categories?: XlsxChartRange; xValues?: XlsxChartRange }
  /** Draw a connecting line (line/scatter) and markers. */
  line?: boolean
  markers?: boolean
}

export interface XlsxChartGroup {
  type: XlsxChartType
  /** Bar direction: `col` = vertical columns, `bar` = horizontal bars. */
  barDir?: 'col' | 'bar'
  grouping: XlsxChartGrouping
  series: XlsxChartSeries[]
  gapWidth: number
  overlap: number
  holeSize?: number
  firstSliceAngle?: number
}

export interface XlsxChartAxis {
  min?: number
  max?: number
  majorUnit?: number
  gridlines: boolean
  deleted: boolean
  title?: string
  formatCode?: string
}

export interface XlsxChartModel {
  title?: string
  groups: XlsxChartGroup[]
  categories: string[]
  categoryPending?: XlsxChartRange
  legend?: 'r' | 'l' | 't' | 'b' | 'tr'
  categoryAxis: XlsxChartAxis
  valueAxis: XlsxChartAxis
  fontFamily: string
  fontSizePt: number
  /** Set when the chart cannot be drawn: the reason shown on the placeholder card. */
  unsupported?: string
  /** Data points dropped by the per-chart bound. */
  truncated?: boolean
}
