// F3: the viewer's chart "refill" (InteractiveSheetDoc.tsx installPivot) calls fillChartRanges on an already-filled
// model. A cached chart (numRef + numCache, what Excel always writes) has no `pending`, so nothing is re-read.
import { expect, it } from 'vitest'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'
import { parseXlsx } from '../src/xlsx/parse'
import { fillChartRanges, refillChartModel } from '../src/charts/fill'
import { pointList } from '../src/charts/points'

const CACHED_REF = '<chartSpace><chart><plotArea><barChart><barDir val="col"/><grouping val="clustered"/><ser><idx val="0"/><val><numRef><f>Styled!$A$1:$A$2</f><numCache><ptCount val="2"/><pt idx="0"><v>3</v></pt><pt idx="1"><v>5</v></pt></numCache></numRef></val></ser></barChart><catAx/><valAx/></plotArea></chart></chartSpace>'

it('F3 (Critical): after recomputeDependents reports a chart, the viewer refill re-reads its cells', async () => {
  const doc = await parseXlsx(await chartWorkbook(CACHED_REF, { cells: [{ ref: 'A1', v: 3 }, { ref: 'A2', v: 5 }] }))
  const node = doc.sheets[0]!.drawings!.find(d => d.content?.kind === 'xlsxChart')!
  const chart = (node.content as { chart: Parameters<typeof fillChartRanges>[0] }).chart
  expect(chart.groups[0]!.series[0]!.refs?.values).toBeTruthy()
  // The cells now hold recomputed values (as recomputeDependents writes them); the viewer's lookup returns them.
  const pending = { ...chart, groups: chart.groups.map(group => ({ ...group, series: group.series.map(series => ({ ...series, pending: series.refs })) })) }
  const next = refillChartModel(chart, () => fillChartRanges(pending, 'Styled', (_name, row) => ({ value: row === 0 ? 100 : 200 })))
  expect(pointList(next.groups[0]!.series[0]!.values).map(p => (p as { v?: number }).v)).toEqual([100, 200])
})
