/**
 * Task B1 review round 1: analytic acquisition on every data path (I1), authored binning validation (I2), the
 * By-Category discriminator (I3), bounded and admitted text (I4), the chart-wide label budget (I5), styling and
 * label metadata in the shared model (I6) and the duplicate-dimension diagnostic (M1).
 *
 * Inline parts and workbooks are Office-shaped synthetic; native parts come from tests/fixtures/chartex.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseThemeContext } from '../src/drawing/style'
import { chartParseContextFor, documentChartContext, chartParseContext } from '../src/charts/context'
import { readChartPart } from '../src/charts/read'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { chartExFamily, isChartExGroup, parseChartEx, parseChartExPart } from '../src/charts/chartex/parse'
import { chartModelOwner, releaseChartModel } from '../src/charts/own'
import { fillWorkbookCharts } from '../src/charts/fill'
import { pointList } from '../src/charts/points'
import type { ChartExGroup, ChartModel } from '../src/charts/model'
import { buildChartPptx, cell, embeddedWorkbook, row } from './helpers/chart-fixtures'

const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const theme = parseThemeContext()
const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, 'fixtures', 'chartex', name)))
const space = (body: string) => `<cx:chartSpace xmlns:cx="${CX}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${body}</cx:chartSpace>`
const plot = (series: string) => `<cx:chart><cx:plotArea><cx:plotAreaRegion>${series}</cx:plotAreaRegion></cx:plotArea></cx:chart>`

async function readPart(pkg: OfficePackage, path: string, context = chartParseContextFor(pkg, path)) {
  const read = await readChartPart(pkg, path, context)
  if (read?.kind !== 'tree') throw new Error(`${path}: ${read ? read.reason : 'missing'}`)
  return { root: read.root, context, release: read.release }
}
async function synthetic(body: string, path = 'ppt/charts/chartEx1.xml') {
  const zip = new JSZip()
  zip.file(path, space(body))
  return readPart(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })), path)
}
const group = (model: ChartModel): ChartExGroup => {
  const g = model.groups[0]
  if (!g || !isChartExGroup(g)) throw new Error(`not a ChartEx group: ${model.unsupported}`)
  return g
}
/** A sparse numeric cache of `count` points (two written), bound by one series of `layout`. */
const cachedPart = (layout: string, count: number, layoutPr = layout === 'clusteredColumn' ? '<cx:layoutPr><cx:binning /></cx:layoutPr>' : '') =>
  `<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="${count}"><cx:pt idx="0">1</cx:pt><cx:pt idx="${count - 1}">2</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>${plot(`<cx:series layoutId="${layout}"><cx:dataId val="0" />${layoutPr}</cx:series>`)}`
/** A reference-only histogram part. */
const referencePart = (formula: string, externalData = false) =>
  `<cx:chartData>${externalData ? '<cx:externalData r:id="rId1" />' : ''}<cx:data id="0"><cx:numDim type="val"><cx:f>${formula}</cx:f></cx:numDim></cx:data></cx:chartData>${plot('<cx:series layoutId="clusteredColumn"><cx:dataId val="0" /><cx:layoutPr><cx:binning /></cx:layoutPr></cx:series>')}`

describe('B1 review round 1', () => {
  it('I1: analytic families acquire every point to 65,536 on the cache path; above it the chart is refused; display families cap', async () => {
    for (const [count, expected] of [[4097, 4097], [65536, 65536], [65537, 'refused']] as const) {
      const p = await synthetic(cachedPart('clusteredColumn', count))
      const result = await parseChartExPart(p.root, theme, p.context, new Set(['histogram']))
      if (expected === 'refused') expect(result, `${count}`).toMatchObject({ kind: 'refused', reason: 'This chart has too many data points to summarize', feature: 'ChartEx' })
      else {
        expect(result.kind, `${count}`).toBe('native')
        if (result.kind === 'native') {
          expect(result.model.groups[0].series[0].values.length, `${count}`).toBe(expected)
          expect(result.model.groups[0].series[0].analytic).toBe('chart')
          expect(result.model.truncated, `${count}: never truncated`).toBeUndefined()
          releaseChartModel(result.model)
        }
      }
      p.release()
    }
    // Waterfall is a display family: the display cap applies (truncated, never refused).
    const wf = await synthetic(cachedPart('waterfall', 4097))
    const waterfall = parseChartEx(wf.root, theme, wf.context)
    expect(group(waterfall).series[0].values.length).toBe(4096)
    expect(group(waterfall).series[0].analytic).toBeUndefined()
    expect(waterfall.truncated).toBe(true)
    wf.release()
  })

  it('I1: the defined-name / host-sheet path acquires every point to 65,536 and refuses above it or on an unavailable point', async () => {
    for (const [count, expected] of [[4097, 4097], [65536, 65536], [65537, 'refused']] as const) {
      const p = await synthetic(referencePart('_xlchart.v1.0'), 'xl/charts/chartEx1.xml')
      const model = parseChartEx(p.root, theme, p.context)
      expect(model.groups[0].series[0].pending?.values?.definedName).toBe('_xlchart.v1.0')
      const rows = Array.from({ length: count }, (_, i) => ({ index: i, cells: [{ row: i, col: 0, value: i + 1 }] }))
      const sheet = { name: 'Sheet1', workbookIndex: 0, rows, drawings: [{ content: { kind: 'xlsxChart', chart: model } }] }
      fillWorkbookCharts([sheet], { names: [{ name: '_xlchart.v1.0', source: `Sheet1!$A$1:$A$${count}` }], workspace: p.context.workspace })
      const filled = (sheet.drawings[0].content as { chart: ChartModel }).chart
      if (expected === 'refused') expect(filled, `${count}`).toMatchObject({ unsupported: 'This chart has too many data points to summarize', refused: true })
      else {
        expect(filled.unsupported, `${count}`).toBeUndefined()
        const values = filled.groups[0].series[0].values
        expect(values.length, `${count}`).toBe(expected)
        expect(pointList(values).at(-1), `${count}`).toEqual({ state: 'value', v: count })
        expect(filled.truncated, `${count}`).toBeUndefined()
      }
      releaseChartModel(filled)
      p.release()
    }
    // A required value point that is unavailable (a formula with no saved value) refuses an analytic chart.
    const p = await synthetic(referencePart('Sheet1!$A$1:$A$3'), 'xl/charts/chartEx1.xml')
    const model = parseChartEx(p.root, theme, p.context)
    const sheet = { name: 'Sheet1', workbookIndex: 0, rows: [0, 1, 2].map(i => ({ index: i, cells: [{ row: i, col: 0, value: i }] })), drawings: [{ content: { kind: 'xlsxChart', chart: model } }] }
    fillWorkbookCharts([sheet], { available: c => c.row !== 1, workspace: p.context.workspace })
    expect((sheet.drawings[0].content as { chart: ChartModel }).chart).toMatchObject({ unsupported: "This chart's data is not available", refused: true })
    p.release()
  })

  it('I1: the embedded-workbook path (DOCX/PPTX) acquires every point to 65,536 and refuses above it', async () => {
    for (const [count, expected] of [[4097, 4097], [65536, 65536], [65537, 'refused']] as const) {
      const data = Array.from({ length: count }, (_, i) => row(i + 1, cell(`A${i + 1}`, i + 1))).join('')
      const built = buildChartPptx([{ chartXml: space(referencePart(`Sheet1!$A$1:$A$${count}`, true)), workbook: embeddedWorkbook({ sheets: [{ name: 'Sheet1', data }] }) }])
      const pkg = await OfficePackage.load(built.bytes)
      const doc = documentChartContext('desktop')
      const p = await readPart(pkg, built.chartPaths[0], chartParseContext(doc, pkg, built.chartPaths[0]))
      const result = await parseChartExPart(p.root, theme, p.context, new Set(['histogram']))
      if (expected === 'refused') expect(result, `${count}`).toMatchObject({ kind: 'refused', reason: 'This chart has too many data points to summarize' })
      else {
        expect(result.kind, `${count}`).toBe('native')
        if (result.kind === 'native') {
          expect(result.model.groups[0].series[0].values.length, `${count}`).toBe(expected)
          expect(pointList(result.model.groups[0].series[0].values).at(-1)).toEqual({ state: 'value', v: count })
          releaseChartModel(result.model)
        }
      }
      p.release()
    }
  }, 60_000)

  it('I2: invalid authored binning is refused at parse; automatic outer bins are unsupported', async () => {
    const histogram = (binning: string) => `<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>${plot(`<cx:series layoutId="clusteredColumn"><cx:dataId val="0" /><cx:layoutPr>${binning}</cx:layoutPr></cx:series>`)}`
    for (const binning of [
      '<cx:binning><cx:binSize val="0" /></cx:binning>', '<cx:binning><cx:binSize val="-1" /></cx:binning>', '<cx:binning><cx:binSize val="NaN" /></cx:binning>',
      '<cx:binning><cx:binCount val="0" /></cx:binning>', '<cx:binning><cx:binCount val="bad" /></cx:binning>', '<cx:binning><cx:binCount val="1001" /></cx:binning>',
      '<cx:binning underflow="1" overflow="5"><cx:binCount val="2" /></cx:binning>',
      '<cx:binning underflow="5" overflow="5" />', '<cx:binning intervalClosed="l" underflow="6" overflow="5" />',
      '<cx:binning underflow="x" />', '<cx:binning intervalClosed="m" />', '<cx:binning><cx:binSize val="1" /><cx:binCount val="3" /></cx:binning>',
    ]) {
      const p = await synthetic(histogram(binning))
      expect(await parseChartPart(p.root, theme, p.context), binning).toMatchObject({ kind: 'refused', reason: 'invalid binning', feature: 'ChartEx' })
      p.release()
    }
    // Valid edge cases still parse: equal bounds with l, a count that leaves one inner bin, exactly 1,000 bins.
    for (const [binning, expected] of [
      ['<cx:binning intervalClosed="l" underflow="5" overflow="5" />', { intervalClosed: 'l', underflow: 5, overflow: 5 }],
      ['<cx:binning underflow="1" overflow="5"><cx:binCount val="3" /></cx:binning>', { intervalClosed: 'r', underflow: 1, overflow: 5, binCount: 3 }],
      ['<cx:binning><cx:binCount val="1000" /></cx:binning>', { intervalClosed: 'r', binCount: 1000 }],
    ] as const) {
      const p = await synthetic(histogram(binning))
      expect(group(parseChartEx(p.root, theme, p.context)).chartEx.series[0].layoutPr.binning, binning).toEqual(expected)
      p.release()
    }
    const auto = await synthetic(histogram('<cx:binning underflow="auto" />'))
    expect(await parseChartPart(auto.root, theme, auto.context)).toMatchObject({ kind: 'unsupported', reason: 'Automatic histogram overflow and underflow bins are not supported' })
    auto.release()
  })

  it('I3: By Category needs an aggregation element; neither binning nor aggregation is unsupported', async () => {
    const column = (layoutPr: string, extra = '') => `<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>${plot(`<cx:series layoutId="clusteredColumn"><cx:dataId val="0" />${layoutPr}</cx:series>${extra}`)}`
    for (const [body, expected] of [
      [column(''), { unsupported: 'This chart could not be read' }],
      [column('<cx:layoutPr />'), { unsupported: 'This chart could not be read' }],
      [column('<cx:layoutPr><cx:aggregation /></cx:layoutPr>'), { family: 'histogram', byCategory: true }],
      [column('<cx:layoutPr><cx:aggregation /></cx:layoutPr>', '<cx:series layoutId="paretoLine" ownerIdx="0" />'), { family: 'pareto', byCategory: true }],
      [column('<cx:layoutPr><cx:binning /></cx:layoutPr>', '<cx:series layoutId="clusteredColumn"><cx:dataId val="0" /><cx:layoutPr><cx:aggregation /></cx:layoutPr></cx:series>'), { unsupported: 'This chart could not be read' }],
    ] as const) {
      const p = await synthetic(body)
      expect(chartExFamily(p.root), body).toEqual(expected)
      p.release()
    }
  })

  it('I4: text is bounded and admitted before it is built; every model is owned', async () => {
    // A title of 100 runs x 1,000 units: 255 units kept, the chart truncated and diagnosed.
    const runs = Array.from({ length: 100 }, () => `<a:r><a:t>${'x'.repeat(1000)}</a:t></a:r>`).join('')
    const body = cachedPart('funnel', 2).replace('<cx:chart>', `<cx:chart><cx:title><cx:tx><cx:rich><a:p>${runs}</a:p></cx:rich></cx:tx></cx:title>`)
    const p = await synthetic(body)
    const model = parseChartEx(p.root, theme, p.context)
    expect(model.title?.length).toBe(255)
    expect(model.truncated).toBe(true)
    expect(model.diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-label-truncated' }))
    expect(chartModelOwner(model)).toBeDefined()
    releaseChartModel(model)
    // A workspace that refuses the first reservation refuses the chart before any text is built (no title kept).
    const refusing = { ...p.context, workspace: { ...p.context.workspace, reserve: () => false } as typeof p.context.workspace }
    const refused = parseChartEx(p.root, theme, refusing)
    expect(refused).toMatchObject({ refused: true, unsupported: 'This chart needs more memory than is available for charts', feature: 'ChartEx' })
    expect(refused.title).toBeUndefined()
    // The adapter's unsupported result carries the admitted title (no unadmitted extraction of its own).
    expect(await parseChartPart(p.root, theme, p.context)).toMatchObject({ kind: 'unsupported', title: 'x'.repeat(255) })
    expect(p.context.workspace.used).toBeGreaterThan(0)
    p.release()
  })

  it('I5: cached labels share the chart-wide 4 MiB budget; past it they are dropped, the chart truncated and diagnosed', async () => {
    const n = 10_000, label = 'y'.repeat(300)
    const pts = Array.from({ length: n }, (_, i) => `<cx:pt idx="${i}">${label}</cx:pt>`).join('')
    const body = `<cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="${n}">${pts}</cx:lvl></cx:strDim><cx:numDim type="size"><cx:lvl ptCount="${n}"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>${plot('<cx:series layoutId="treemap"><cx:dataId val="0" /></cx:series>')}`
    const p = await synthetic(body)
    const model = parseChartEx(p.root, theme, p.context)
    const level = group(model).chartEx.data[0].dimensions[0].levels[0]
    if (level.kind !== 'str') throw new Error('string level')
    const kept = level.values.filter((v): v is string => v !== undefined)
    expect(kept.every(v => v.length === 255)).toBe(true)
    const bytes = kept.reduce((sum, v) => sum + v.length * 2 + 32, 0)
    expect(bytes).toBeLessThanOrEqual(4 * 1024 * 1024)
    expect(kept.length).toBeLessThan(n)
    expect(model.truncated).toBe(true)
    expect(model.diagnostics?.map(d => d.code)).toEqual(expect.arrayContaining(['chart-label-truncated', 'chart-label-budget']))
    releaseChartModel(model)
    p.release()
  }, 60_000)

  it('I6: chart-space, plot, title, legend, axis and series styling, point overrides and data labels reach the model', async () => {
    // WF2 (native): chart-space gradient fill and line, title fill and dashed line, plot surface noFill, axis titles.
    const wf2Pkg = await OfficePackage.load(fixture('waterfall2/waterfall2.xlsx'))
    const wf2 = await readPart(wf2Pkg, 'xl/charts/chartEx1.xml')
    // Scheme colours resolve through the workbook's own theme.
    const w = parseChartEx(wf2.root, parseThemeContext(await wf2Pkg.text('xl/theme/theme1.xml')), wf2.context)
    expect(w.chartAreaFill?.kind).toBe('gradient')
    expect(w.chartAreaLine).toBeDefined()
    expect(w.titleFill?.kind).toBe('solid')
    expect(w.titleLine).toMatchObject({ dash: 'dash', color: expect.any(String) })
    expect(w.plotAreaFill).toEqual({ kind: 'none' })
    expect(w.axes.map(a => a.title)).toEqual(['Entry', 'Cumulative change'])
    // The title run: 000000 with lumMod 65 % / lumOff 35 % (Office's dark grey), 14 pt, not bold.
    expect(w.titleText).toMatchObject({ color: '595959', sizePt: 14, bold: false })
    expect(group(w).series[0].labels).toMatchObject({ position: 'outEnd', showSerName: false, showCatName: false, showVal: true })
    wf2.release()
    // FN1 (native): series labels, per-point overrides (field by field), hidden labels.
    const fn1 = await readPart(await OfficePackage.load(fixture('funnel-label-options/funnel-label-options.xlsx')), 'xl/charts/chartEx1.xml')
    const labels = group(parseChartEx(fn1.root, theme, fn1.context)).series[0].labels!
    expect(labels).toMatchObject({ position: 'ctr', showSerName: false, showCatName: true, showVal: true, separator: ', ' })
    expect(labels.points?.[0]).toMatchObject({ position: 't', formatCode: '0.00%', sourceLinked: false, showVal: true, separator: ';' })
    expect(labels.points?.[1]).toMatchObject({ showSerName: true, showVal: false, formatCode: '# ?/?' })
    expect(labels.points?.[4]).toEqual({ delete: true })
    expect(labels.points?.[8]).toEqual({ delete: true })
    fn1.release()
    // Office-shaped synthetic: series spPr, dataPt fill/line overrides, label spPr and txPr.
    const styled = await synthetic(`<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="3"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">-1</cx:pt><cx:pt idx="2">2</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>
      <cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:spPr><a:solidFill><a:srgbClr val="112233" /></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="445566" /></a:solidFill></a:ln></cx:spPr>
      <cx:dataPt idx="1"><cx:spPr><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill></cx:spPr></cx:dataPt>
      <cx:dataLabels><cx:spPr><a:solidFill><a:srgbClr val="EEEEEE" /></a:solidFill></cx:spPr><cx:txPr><a:bodyPr /><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="0000FF" /></a:solidFill></a:defRPr></a:pPr></a:p></cx:txPr><cx:visibility value="1" /></cx:dataLabels>
      <cx:dataId val="0" /></cx:series></cx:plotAreaRegion><cx:spPr><a:solidFill><a:srgbClr val="ABCDEF" /></a:solidFill></cx:spPr></cx:plotArea></cx:chart>
      <cx:txPr><a:bodyPr /><a:p><a:pPr><a:defRPr sz="1200" /></a:pPr></a:p></cx:txPr>`)
    const m = parseChartEx(styled.root, theme, styled.context)
    const s = group(m).series[0]
    expect(s.color).toBe('112233')
    expect(s.fill).toMatchObject({ kind: 'solid' })
    expect(s.lineStyle).toMatchObject({ color: '445566' })
    expect(s.dPt?.get(1)).toMatchObject({ color: 'FF0000' })
    expect(s.pointColors?.[1]).toBe('FF0000')
    expect(s.labels).toMatchObject({ showVal: true, showSerName: false, showCatName: false, fill: { kind: 'solid' }, text: { sizePt: 9, bold: true, color: '0000FF' } })
    expect(group(m).chartEx.plotAreaFill).toMatchObject({ kind: 'solid' })
    expect(m.text?.sizePt).toBe(12)
    expect(m.fontSizePt).toBe(12)
    styled.release()
  })

  it('M1: a repeated dimension type keeps the first and is diagnosed', async () => {
    const p = await synthetic(`<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">9</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>${plot('<cx:series layoutId="funnel"><cx:dataId val="0" /></cx:series>')}`)
    const model = parseChartEx(p.root, theme, p.context)
    expect(pointList(group(model).series[0].values)).toEqual([{ state: 'value', v: 1 }])
    expect(model.diagnostics).toContainEqual(expect.objectContaining({ code: 'chartex-duplicate-dimension' }))
    p.release()
  })
})
