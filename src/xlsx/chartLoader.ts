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

/** The receiving document (spec 6.4) and the context lease its paints carry. */
export type ChartReceiver = import('../charts/paint').ChartReceiver
export type ChartReceiverLease = import('../charts/paint').ChartReceiverLease
/** Preparation before paint (spec 6.5): a load failure is swallowed, so pages still return, with cards. With a
 * receiver, the receiving chart context this preparation's paints carry (release it with the preparation). */
export const ensureChartEngine = (contents: readonly ChartCarrier[], receiver?: ChartReceiver): Promise<ChartReceiverLease | undefined> =>
  hasXlsxCharts(contents) ? loadChartModule().then(charts => receiver ? charts.acquireChartContext(receiver) : undefined).catch(() => undefined) : Promise.resolve(undefined)
/** Make the chart painter available before a sheet with charts is painted (and acquire its receiving context). */
export async function ensureChartPainter(drawings: readonly XlsxDrawing[] | undefined, receiver?: ChartReceiver): Promise<ChartReceiverLease | undefined> {
  if (!hasXlsxCharts(drawings)) return undefined
  const charts = await loadChartModule()
  return receiver ? charts.acquireChartContext(receiver) : undefined
}
