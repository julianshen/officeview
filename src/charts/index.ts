/**
 * The chart engine (lazy chunk), one for every format (spec 6.1): reads
 * c:chartSpace into a chart model, fills uncached series from worksheet
 * cells and paints the common chart types. Anything else becomes a
 * placeholder card carrying the chart title and an accessible description;
 * nothing is ever left as silent blank space.
 *
 * Loaded through `xlsx/chartLoader.ts` (`loadChartModule`) only when a
 * document has a chart; importing it registers the painter with the shared
 * content painter. Never import this module eagerly (EAGER_FORBIDDEN).
 */
import { registerXlsxChartPainter } from '../drawing/content-paint'
import { chartDescription, chartTypeName } from './describe'
import { paintXlsxChart } from './paint'

export { chartParseContext, chartParseContextFor } from './context'
export { readChartPart, type ChartRead } from './read'
export { parseChartPart, applyDataPrecedence } from './dataPrecedence'
export { openEmbeddedWorkbook, resolveExternalData, type EmbeddedWorkbook, type EmbeddedWorkbookResult } from './embeddedZip'
export * from './parse'
export * from './fill'
export * from './points'
export { displayView, type DisplayView } from './display'
export { releaseChartModel, chartModelOwner, releaseChartRetained } from './own'
export { paintXlsxChart }
export { layoutChart, releaseChartLayouts, releaseWorkspaceLayouts, type ChartLayout } from './layout'
export { acquireChartContext, CHART_STANDALONE_PAINT_PROFILE, type ChartReceiver, type ChartReceiverLease } from './paint'
export { chartDescription, chartTypeName }
registerXlsxChartPainter(paintXlsxChart)

export { sourceFormat } from './format'

export { prepareChartPaintables, CHART_DESCRIPTIONS_PER_PAGE, type PaintableChart, type PlacedChart } from './paintables'
