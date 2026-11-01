import type { ChartModel } from '../../../charts/model'
import type { XlsxDrawing } from '../../types'
import type { XlsxSheet } from '../../types'
import { cellDisplayText } from '../../render'
import type { ChartCell } from '../../../charts/fill'

/** Keep the original resident chart model until the pivot session is restored. */
export function rememberDependentChart(
  snapshots: Map<XlsxDrawing, ChartModel>, node: XlsxDrawing, current: ChartModel,
): void {
  if (!snapshots.has(node)) snapshots.set(node, current)
}

/** Restore charts changed by dependent-value refills and release the snapshots. */
export function restoreDependentCharts(snapshots: Map<XlsxDrawing, ChartModel>): XlsxDrawing[] {
  const restored: XlsxDrawing[] = []
  for (const [node, chart] of snapshots) {
    if (node.content?.kind === 'xlsxChart') { node.content.chart = chart; restored.push(node) }
  }
  snapshots.clear()
  return restored
}

/** The dependent refill path reads committed overlay values but keeps the host cell's number format. */
export function overlayChartCell(sheet: XlsxSheet, row: number, col: number, value: string | number | boolean | null, error: string | undefined, needsText: boolean): ChartCell {
  const cell = sheet.rows.find(candidate => candidate.index === row)?.cells.find(candidate => candidate.col === col)
  const style = cell?.style
  return {
    value,
    ...(error === undefined ? {} : { error }),
    ...(style?.formatCode ? { formatCode: style.formatCode } : {}),
    ...(needsText ? { text: cellDisplayText(sheet, { value, style }) } : {}),
  }
}
