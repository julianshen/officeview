/** Entry to the lazy chart engine (src/charts), for every format: nothing here pulls chart code in eagerly. */
import type { XlsxDrawing } from './types'
import type { DrawingContent } from '../drawing/content'

export type ChartModule = typeof import('../charts')
let pending: Promise<ChartModule> | undefined
/** Loads (once) the chart parser/painter; loading registers the painter. */
export function loadChartModule(): Promise<ChartModule> {
  return pending ??= import('../charts').catch(error => { pending = undefined; throw error })
}

/** Headless embedders: load the chart engine (and register its painter) before a synchronous paint. */
export const preloadCharts = (): Promise<void> => loadChartModule().then(() => {})
type ChartCarrier = { content?: DrawingContent<unknown, unknown>; children?: readonly ChartCarrier[] }
/** Whether any drawing (or group child) carries a worksheet chart. */
export function hasXlsxCharts(drawings: readonly ChartCarrier[] | undefined): boolean {
  for (const node of drawings ?? []) {
    if (node.content?.kind === 'xlsxChart') return true
    if (node.children?.length && hasXlsxCharts(node.children)) return true
  }
  return false
}

/** Preparation before paint (spec 6.5): a load failure is swallowed, so pages still return, with cards. */
export const ensureChartEngine = (contents: readonly ChartCarrier[]): Promise<unknown> => hasXlsxCharts(contents) ? loadChartModule().catch(() => {}) : Promise.resolve()
/** Make the chart painter available before a sheet with charts is painted. */
export async function ensureChartPainter(drawings: readonly XlsxDrawing[] | undefined): Promise<void> {
  if (hasXlsxCharts(drawings)) await loadChartModule()
}
