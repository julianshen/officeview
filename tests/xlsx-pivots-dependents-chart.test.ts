// Non-pivot charts on dependent formula cells (Task 31, spec §7.5): a commit refills the charts whose ranges hold a
// recomputed cell (a cached chart, what Excel writes, is re-read from its cells), keeps the host cells' number formats,
// and gives them their saved models back when the pivots go.
import { expect, it } from 'vitest'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'
import { parseXlsx } from '../src/xlsx/parse'
import { fillChartRanges, refillChartModel } from '../src/charts/fill'
import { pointList } from '../src/charts/points'
import { overlayChartCell, rememberDependentChart, restoreDependentCharts } from '../src/xlsx/pivots/session/dependentCharts'

const CACHED_REF = '<chartSpace><chart><plotArea><barChart><barDir val="col"/><grouping val="clustered"/><ser><idx val="0"/><val><numRef><f>Styled!$A$1:$A$2</f><numCache><ptCount val="2"/><pt idx="0"><v>3</v></pt><pt idx="1"><v>5</v></pt></numCache></numRef></val></ser></barChart><catAx/><valAx/></plotArea></chart></chartSpace>'

it('after recomputeDependents reports a chart, the refill re-reads its cells (a cached chart has no pending ranges)', async () => {
  const doc = await parseXlsx(await chartWorkbook(CACHED_REF, { cells: [{ ref: 'A1', v: 3 }, { ref: 'A2', v: 5 }] }))
  const node = doc.sheets[0]!.drawings!.find(d => d.content?.kind === 'xlsxChart')!
  const chart = (node.content as { chart: Parameters<typeof fillChartRanges>[0] }).chart
  expect(chart.groups[0]!.series[0]!.refs?.values).toBeTruthy()
  // The cells now hold recomputed values (as recomputeDependents writes them); the viewer's lookup returns them.
  const pending = { ...chart, groups: chart.groups.map(group => ({ ...group, series: group.series.map(series => ({ ...series, pending: series.refs })) })) }
  const next = refillChartModel(chart, () => fillChartRanges(pending, 'Styled', (_name, row) => ({ value: row === 0 ? 100 : 200 })))
  expect(pointList(next.groups[0]!.series[0]!.values).map(p => (p as { v?: number }).v)).toEqual([100, 200])
})

it('restores a dependent-refilled non-pivot chart when the pivot client is torn down', async () => {
  const doc = await parseXlsx(await chartWorkbook(CACHED_REF, { cells: [{ ref: 'A1', v: 3 }, { ref: 'A2', v: 5 }] }))
  const node = doc.sheets[0]!.drawings!.find(d => d.content?.kind === 'xlsxChart')!
  if (node.content?.kind !== 'xlsxChart' || !node.content.chart) throw new Error('fixture chart missing')
  const original = node.content.chart, snapshots = new Map<typeof node, typeof original>()
  const pending = { ...original, groups: original.groups.map(group => ({ ...group, series: group.series.map(series => ({ ...series, pending: series.refs ?? series.pending })) })) }
  const replacement = refillChartModel(original, () => fillChartRanges(pending, 'Styled', (_name, row) => ({ value: row === 0 ? 100 : 200 })))
  expect(pointList(replacement.groups[0]!.series[0]!.values).map(p => (p as { v?: number }).v)).toEqual([100, 200])
  rememberDependentChart(snapshots, node, original)
  node.content.chart = replacement
  // A second refill keeps the first source model, not an intermediate model.
  rememberDependentChart(snapshots, node, replacement)
  const invalidated = restoreDependentCharts(snapshots)
  expect(invalidated).toEqual([node])
  expect(node.content.chart).toBe(original)
  expect(snapshots.size).toBe(0)
})

it('keeps a category cell number format when the committed pivot overlay supplies its value', async () => {
  const doc = await parseXlsx(await chartWorkbook(CACHED_REF, { cells: [{ ref: 'A1', v: 3 }, { ref: 'A2', v: 5 }] }))
  const sheet = doc.sheets[0]!
  const cell = sheet.rows[0]!.cells[0]!
  cell.style = { numFmtId: 10, formatCode: '0%' }
  expect(overlayChartCell(sheet, 0, cell.col, 0.25, undefined, true)).toMatchObject({ value: 0.25, formatCode: '0%', text: '25%' })
})
