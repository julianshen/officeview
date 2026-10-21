/**
 * Legacy names of the chart model (types only). D0 moved the engine onto the
 * format-neutral model in `src/charts/model.ts`; these aliases keep existing
 * imports compiling. New code imports from `src/charts/model` directly.
 */
import type { ChartAxis, ChartFamily, ChartGroup, ChartGrouping, ChartModel, ChartRange, ChartSeries } from '../charts/model'
export type * from '../charts/model'
export type XlsxChartType = ChartFamily
export type XlsxChartGrouping = ChartGrouping
export type XlsxChartRange = ChartRange
export type XlsxChartSeries = ChartSeries
export type XlsxChartGroup = ChartGroup
export type XlsxChartAxis = ChartAxis
export type XlsxChartModel = ChartModel
