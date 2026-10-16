/** Entry to the lazy worksheet chart chunk: nothing here pulls chart code in eagerly. */
import type { XlsxDrawing } from './types'

export type ChartModule = typeof import('./chart')
let pending: Promise<ChartModule> | undefined
/** Loads (once) the chart parser/painter; loading registers the painter. */
export function loadChartModule(): Promise<ChartModule> {
  return pending ??= import('./chart').catch(error => { pending = undefined; throw error })
}

/** Whether any drawing (or group child) carries a worksheet chart. */
export function hasXlsxCharts(drawings: readonly XlsxDrawing[] | undefined): boolean {
  for (const node of drawings ?? []) {
    if (node.content?.kind === 'xlsxChart') return true
    if (node.children?.length && hasXlsxCharts(node.children)) return true
  }
  return false
}

/** Make the chart painter available before a sheet with charts is painted. */
export async function ensureChartPainter(drawings: readonly XlsxDrawing[] | undefined): Promise<void> {
  if (hasXlsxCharts(drawings)) await loadChartModule()
}
