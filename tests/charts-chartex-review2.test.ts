/**
 * Task B1 review round 2 regressions (codex re-review of 46197aa): the paged reader keeps the analytic
 * too-many-points refusal (I1), paragraph separators count toward the text cap and the walk stops there (I4), one
 * 4 MiB label budget covers every ChartEx family and series, references included (I5), and point/label overrides
 * survive for analytic source indices up to 65,536 (I6). Parts are Office-shaped synthetic.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseChartEx } from '../src/charts/chartex/parse'
import { createChartWorkspace } from '../src/charts/workspace'
import { fillChartRanges } from '../src/charts/fill'
import { PagedXlsxSheetAdapter } from '../src/xlsx/view/pagedSource'
import type { ChartModel } from '../src/charts/model'
import type { ChartParseContext } from '../src/charts/context'

const context = (): ChartParseContext => ({ partPath: 'xl/charts/chartEx1.xml', profile: 'desktop', workspace: createChartWorkspace('desktop'), pkg: {} as ChartParseContext['pkg'] })
const root = (data: string, series: string, title = '') => parseXmlOrdered(`<chartSpace><chartData>${data}</chartData><chart>${title}<plotArea><plotAreaRegion>${series}</plotAreaRegion></plotArea></chart></chartSpace>`)
const cache = '<numDim type="val"><lvl ptCount="1"><pt idx="0">1</pt></lvl></numDim>'
const ser = (id = 0, layout = 'funnel', extra = '') => `<series layoutId="${layout}"><dataId val="${id}"/>${layout === 'clusteredColumn' ? '<layoutPr><binning/></layoutPr>' : ''}${extra}</series>`
const slice = String.prototype.slice
afterEach(() => { String.prototype.slice = slice })

describe('B1 review round 2', () => {
  it('I1: the paged reader refuses an over-limit analytic chart with the point-limit reason before memory admission', async () => {
    const c = context()
    const over = parseChartEx(root('<data id="0"><numDim type="val"><f>Sheet1!A1:A65537</f></numDim></data>', ser(0, 'clusteredColumn')), undefined, c)
    const drawings = [{ content: { kind: 'xlsxChart', chart: over } }]
    const adapter = Object.create(PagedXlsxSheetAdapter.prototype) as { fillCharts: (...args: unknown[]) => Promise<unknown> }
    Object.assign(adapter, { workbook: { formatting: { dateSystem: '1900' }, sheets: [{ name: 'Sheet1', sheetId: '1', state: 'complete', importedThroughRow: 65536 }] } })
    const requested: number[] = []
    await adapter.fillCharts(drawings, 'Sheet1', new AbortController().signal, [], (bytes: number) => { requested.push(bytes); return false }, c.workspace)
    const chart = (drawings[0].content as { chart: ChartModel }).chart
    expect({ unsupported: chart.unsupported, refused: chart.refused }).toEqual({ unsupported: 'This chart has too many data points to summarize', refused: true })
    // Nothing was admitted for the refused chart's points.
    expect(requested.every(bytes => bytes === 0)).toBe(true)
  })

  it('I4: paragraph separators count toward the text cap and the walk stops at it (no unbounded slice)', () => {
    const title = `<title><tx><rich><p><r><t>X</t></r></p>${'<p/>'.repeat(1000)}<p><r><t>${'Z'.repeat(100000)}</t></r></p></rich></tx></title>`
    const c = context()
    const observed: number[] = []
    String.prototype.slice = function (this: string, ...args: [number?, number?]) {
      const result = slice.apply(this, args)
      if (this.length === 100000) observed.push(result.length)
      return result
    }
    const model = parseChartEx(root(`<data id="0">${cache}</data>`, ser(), title), undefined, c)
    String.prototype.slice = slice
    expect(observed.every(length => length <= 256), `slices of the 100,000-unit run: ${observed}`).toBe(true)
    expect(model.title?.length).toBeLessThanOrEqual(255)
    // The separators fill the cap; trailing whitespace is trimmed.
    expect(model.title).toBe('X')
    expect(model.truncated).toBe(true)
  })

  it('I5: one 4 MiB label budget covers every ChartEx family and series, reference-filled display categories included', () => {
    const data = Array.from({ length: 3 }, (_, i) => `<data id="${i}">${cache}<strDim type="cat"><f>Sheet1!A1:A4096</f></strDim></data>`).join('')
    const c = context()
    const model = parseChartEx(root(data, Array.from({ length: 3 }, (_, i) => ser(i)).join('')), undefined, c)
    const filled = fillChartRanges(model, 'Sheet1', (_sheet, _row, _col, text) => ({ value: text ? 'a'.repeat(300) : 1 }), { workspace: c.workspace })
    const bytes = filled.groups.flatMap(g => g.series).reduce((n, s) => n + (s.categories ?? []).reduce((m, t) => m + (t ? t.length * 2 + 32 : 0), 0), 0)
    expect(bytes).toBeLessThanOrEqual(4 * 1024 * 1024)
    expect(filled.truncated).toBe(true)
    expect(filled.diagnostics?.map(d => d.code)).toEqual(expect.arrayContaining(['chart-label-budget', 'chart-label-truncated']))
    // The budget is shared: cached labels and the title count before the references are filled.
    const cachedData = `<data id="0">${cache}<strDim type="cat"><lvl ptCount="7000">${Array.from({ length: 7000 }, (_, i) => `<pt idx="${i}">${'c'.repeat(255)}</pt>`).join('')}</lvl></strDim></data><data id="1">${cache}<strDim type="cat"><f>Sheet1!A1:A4096</f></strDim></data>`
    const c2 = context()
    const mixed = parseChartEx(root(cachedData, ser(0) + ser(1)), undefined, c2)
    const filled2 = fillChartRanges(mixed, 'Sheet1', (_s, _r, _c, text) => ({ value: text ? 'b'.repeat(255) : 1 }), { workspace: c2.workspace })
    const all = filled2.groups.flatMap(g => g.series).flatMap(s => s.categories ?? [])
    expect(all.reduce((n, t) => n + (t ? t.length * 2 + 32 : 0), 0)).toBeLessThanOrEqual(4 * 1024 * 1024)
    expect(filled2.diagnostics?.map(d => d.code)).toContain('chart-label-budget')
  })

  it('I6: point and label overrides are kept for analytic source indices up to 65,536 (display families stay at 4,096)', () => {
    const overrides = (i: number) => `<dataPt idx="${i}"><spPr><solidFill><srgbClr val="FF0000"/></solidFill></spPr></dataPt><dataLabels><dataLabel idx="${i}"><visibility value="1"/></dataLabel><dataLabelHidden idx="${i + 1}"/></dataLabels>`
    const c = context()
    const treemap = parseChartEx(root('<data id="0"><numDim type="size"><lvl ptCount="65536"><pt idx="65534">1</pt></lvl></numDim></data>', ser(0, 'treemap', overrides(65534))), undefined, c)
    const s = treemap.groups[0].series[0]
    expect(s.values.length).toBe(65536)
    expect(s.dPt?.get(65534)).toMatchObject({ color: 'FF0000' })
    expect(s.labels?.points?.[65534]).toMatchObject({ showVal: true })
    expect(s.labels?.points?.[65535]).toEqual({ delete: true })
    const c2 = context()
    const funnel = parseChartEx(root(`<data id="0">${cache}</data>`, ser(0, 'funnel', overrides(5000))), undefined, c2)
    expect(funnel.groups[0].series[0].dPt?.get(5000)).toBeUndefined()
  })
})
