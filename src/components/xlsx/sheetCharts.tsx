/** Accessible names and notices for a sheet's charts (the canvas carries the pixels). */
import type { ReactElement } from 'react'
import type { XlsxDrawing } from '../../xlsx/types'
import type { XlsxChartModel } from '../../xlsx/chartModel'
import { chartDescription } from '../../charts/describe'

/** Charts of a drawing tree, in drawing order (bounded by the drawing node budget). */
export function sheetCharts(drawings: readonly XlsxDrawing[] | undefined, out: XlsxChartModel[] = []): XlsxChartModel[] {
  for (const node of drawings ?? []) {
    if (node.content?.kind === 'xlsxChart') out.push(node.content.chart)
    if (node.children?.length) sheetCharts(node.children, out)
  }
  return out
}

/** One notice per chart shown as a placeholder card. */
export function chartNotices(charts: readonly XlsxChartModel[]): string[] {
  return charts.filter(chart => chart.unsupported).map(chart => `${chart.title ? `Chart “${chart.title}”` : 'A chart'} is shown as a placeholder: ${chart.unsupported}.`)
}

const HIDDEN = { position: 'absolute', width: 1, height: 1, margin: -1, padding: 0, overflow: 'hidden', clipPath: 'inset(50%)', whiteSpace: 'nowrap', border: 0 } as const

/** A screen-reader list describing every chart on the active sheet. */
export function SheetChartDescriptions({ charts }: { charts: readonly XlsxChartModel[] }): ReactElement | null {
  if (!charts.length) return null
  return (
    <ul aria-label={`Charts on this sheet (${charts.length})`} data-testid="officeview-sheet-charts" style={HIDDEN}>
      {charts.map((chart, index) => <li key={index} role="img" aria-label={chartDescription(chart)} />)}
    </ul>
  )
}
