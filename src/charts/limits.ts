/**
 * Chart bounds (spec 6.5/7). A leaf module: family code needs the point bound
 * without depending on the parser (no import cycle through labels.ts).
 */
/** Bounds: series per chart, points per series drawn (display, spec 7). */
export const CHART_SERIES_LIMIT = 64
export const CHART_POINT_LIMIT = 4096
/** Points per series acquired for an analytic element (trendline, stdDev/stdErr error bars; spec 7). */
export const CHART_ANALYTIC_POINT_LIMIT = 65536
/** Characters kept per chart label (category, series name); longer labels are cut and the chart marked truncated. */
export const CHART_LABEL_LIMIT = 255
/** Hard cap on one fill's label storage (UTF-16 bytes). */
export const CHART_TEXT_BYTES_LIMIT = 4 * 1024 * 1024
/** Refusal reasons shared by the fill, the ChartEx adapter and the analytic families (spec 7, 5.8). */
export const CHART_TOO_MANY_POINTS = 'This chart has too many data points to summarize'
export const CHART_DATA_NOT_AVAILABLE = "This chart's data is not available"
