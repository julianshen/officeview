/**
 * Task B1: the ChartEx (cx:) parser and dispatch, against docs/charts/chartex-map.md.
 *
 * Native evidence: the Office-authored parts in tests/fixtures/chartex (PROVENANCE.md).
 * Inline parts built here are Office-shaped synthetic, used only for cases no native part exhibits
 * (non-dense data ids, sparse points, dangling ids, unknown layouts); they are never Office evidence.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseThemeContext } from '../src/drawing/style'
import { chartParseContextFor } from '../src/charts/context'
import { isChartExRoot, readChartPart } from '../src/charts/read'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { CHARTEX_PAINTED, isChartExGroup, parseChartEx, parseChartExPart } from '../src/charts/chartex/parse'
import type { ChartExGroup, ChartModel } from '../src/charts/model'
import { pointList } from '../src/charts/points'
import { chartNameResolver, resolveChartNames } from '../src/charts/fill'
import { parseDefinedNames } from '../src/xlsx/formula/names'
import { parseXml } from '../src/core/xml'
import { parseXlsx } from '../src/xlsx/parse'

const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const FIXTURES = join(__dirname, 'fixtures', 'chartex')
const fixture = (name: string) => new Uint8Array(readFileSync(join(FIXTURES, name)))
const theme = parseThemeContext()

/** One chartEx part read through the production bounded reader. */
async function readPart(pkg: OfficePackage, path: string) {
  const context = chartParseContextFor(pkg, path)
  const read = await readChartPart(pkg, path, context)
  if (read?.kind !== 'tree') throw new Error(`${path}: ${read ? read.reason : 'missing'}`)
  return { root: read.root, context, release: read.release, pkg }
}
async function nativePart(file: string, path = 'xl/charts/chartEx1.xml') {
  return readPart(await OfficePackage.load(fixture(file)), path)
}
/** An Office-shaped synthetic part, placed at `path` in a minimal package. */
async function syntheticPart(body: string, path = 'ppt/charts/chartEx1.xml') {
  const zip = new JSZip()
  zip.file(path, `<cx:chartSpace xmlns:cx="${CX}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</cx:chartSpace>`)
  return readPart(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })), path)
}
function exGroup(model: ChartModel): ChartExGroup {
  expect(model.unsupported, 'parsed').toBeUndefined()
  const group = model.groups[0]
  if (!group || !isChartExGroup(group)) throw new Error('not a ChartEx group')
  return group
}
const values = (points: ChartModel['groups'][number]['series'][number]['values']) => pointList(points).map(p => p.state === 'value' ? p.v : p.state)

describe('ChartEx parser (chartex-map §§1-9)', () => {
  it('chartex dimensions parse by id with levels and sparse points', async () => {
    // FN2 (PowerPoint): three data sets keyed 0,1,2, each a cached strDim cat + numDim val.
    const fn2 = await nativePart('funnel-pp1/funnel-pp1.pptx', 'ppt/charts/chartEx1.xml')
    expect(isChartExRoot(fn2.root)).toBe(true)
    const funnel = exGroup(parseChartEx(fn2.root, theme, fn2.context))
    expect(funnel.chartEx.data.map(d => d.id)).toEqual([0, 1, 2])
    const d0 = funnel.chartEx.data[0]
    expect(d0.dimensions.map(d => [d.kind, d.type])).toEqual([['str', 'cat'], ['num', 'val']])
    const [cat, val] = d0.dimensions
    expect(cat.ref).toEqual({ sheet: 'Sheet1', r0: 1, c0: 0, r1: 4, c1: 0 })
    expect(cat.dir).toBe('col')
    expect(cat.levels).toHaveLength(1)
    expect(cat.levels[0]).toMatchObject({ kind: 'str', ptCount: 4, values: ['Thing 1', 'Thing 2', 'Thing 3', 'Thing 4'] })
    expect(val.levels[0]).toMatchObject({ kind: 'num', ptCount: 4, formatCode: 'General' })
    const level = val.levels[0]
    if (level.kind !== 'num') throw new Error('numeric level')
    expect(values(level.points)).toEqual([4.3, 2.5, 3.5, 4.5])
    // The series roles come from the bound data set's cache (saved data wins, its reference kept for provenance).
    expect(values(funnel.series[0].values)).toEqual([4.3, 2.5, 3.5, 4.5])
    expect(funnel.series[0].categories).toEqual(['Thing 1', 'Thing 2', 'Thing 3', 'Thing 4'])
    expect(funnel.series[0].pending).toBeUndefined()
    expect(funnel.series[0].refs?.values).toEqual({ sheet: 'Sheet1', r0: 1, c0: 1, r1: 4, c1: 1 })
    fn2.release()

    // XLSX natives carry no cache: each dimension is the hidden _xlchart defined name, left for the
    // workbook's defined names (map §2.2). TM: a multi-column category path and a size dimension.
    const tm = await nativePart('treemap/treemap.xlsx')
    const treemap = exGroup(parseChartEx(tm.root, theme, tm.context))
    expect(treemap.chartEx.data[0].dimensions.map(d => [d.kind, d.type, d.ref?.definedName, d.levels.length]))
      .toEqual([['str', 'cat', '_xlchart.v1.0', 0], ['num', 'size', '_xlchart.v1.1', 0]])
    expect(treemap.series[0].pending).toEqual({ values: expect.objectContaining({ definedName: '_xlchart.v1.1' }), categories: expect.objectContaining({ definedName: '_xlchart.v1.0' }) })
    tm.release()

    // Office-shaped synthetic: non-dense ids out of order resolve by value; missing pt@idx are blank at
    // their position; ptCount pads; a string level keeps its gaps.
    const sparse = await syntheticPart(`<cx:chartData>
      <cx:data id="7"><cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">70</cx:pt><cx:pt idx="1">71</cx:pt></cx:lvl></cx:numDim></cx:data>
      <cx:data id="3"><cx:strDim type="cat"><cx:lvl ptCount="5"><cx:pt idx="1">b</cx:pt><cx:pt idx="4">e</cx:pt></cx:lvl></cx:strDim>
        <cx:numDim type="val"><cx:lvl ptCount="5" formatCode="0.0"><cx:pt idx="0">1</cx:pt><cx:pt idx="3">4</cx:pt><cx:pt idx="4">#N/A</cx:pt></cx:lvl></cx:numDim></cx:data>
      </cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>
      <cx:series layoutId="funnel"><cx:dataId val="3" /></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>`)
    const synthetic = exGroup(parseChartEx(sparse.root, theme, sparse.context))
    // Only the bound data set is acquired.
    expect(synthetic.chartEx.data.map(d => d.id)).toEqual([3])
    expect(synthetic.chartEx.series[0].dataId).toBe(3)
    expect(values(synthetic.series[0].values)).toEqual([1, 'blank', 'blank', 4, 'error'])
    expect(pointList(synthetic.series[0].values)[4]).toEqual({ state: 'error', code: '#N/A' })
    expect(synthetic.series[0].categories).toEqual(['', 'b', '', '', 'e'])
    const strLevel = synthetic.chartEx.data[0].dimensions[0].levels[0]
    expect(strLevel.kind === 'str' && strLevel.values).toEqual([undefined, 'b', undefined, undefined, 'e'])
    sparse.release()

    // A multi-level cache is ignored (map §2.2, level order unconfirmed): the referenced range is read instead;
    // with no reference the chart is unsupported.
    const multi = (f: string) => `<cx:chartData><cx:data id="0"><cx:strDim type="cat">${f}<cx:lvl ptCount="1"><cx:pt idx="0">leaf</cx:pt></cx:lvl><cx:lvl ptCount="1"><cx:pt idx="0">root</cx:pt></cx:lvl></cx:strDim>
      <cx:numDim type="size"><cx:lvl ptCount="1"><cx:pt idx="0">5</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="treemap"><cx:dataId val="0" /></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>`
    const withRef = await syntheticPart(multi('<cx:f>Sheet1!$A$2:$B$2</cx:f>'))
    const tree = exGroup(parseChartEx(withRef.root, theme, withRef.context))
    expect(tree.chartEx.data[0].dimensions[0].levels).toEqual([])
    expect(tree.series[0].categories).toBeUndefined()
    expect(tree.series[0].pending?.categories).toEqual({ sheet: 'Sheet1', r0: 1, c0: 0, r1: 1, c1: 1 })
    withRef.release()
    const noRef = await syntheticPart(multi(''))
    expect(await parseChartPart(noRef.root, theme, noRef.context)).toMatchObject({ kind: 'unsupported', reason: 'multi-level cache order unconfirmed', feature: 'ChartEx' })
    noRef.release()
  })

  it('series bind dimensions through dataId; layoutId selects the family', async () => {
    const parse = async (file: string, path?: string) => { const p = await nativePart(file, path); try { return parseChartEx(p.root, theme, p.context) } finally { p.release() } }

    // BW1: three boxWhisker series bound to data 0, 1, 2 (no cat dimension), per-series layout properties.
    const box = exGroup(await parse('boxWhisker/boxWhisker.xlsx'))
    expect(box.family).toBe('boxWhisker')
    expect(box.chartEx.series.map(s => [s.layout, s.dataId])).toEqual([['boxWhisker', 0], ['boxWhisker', 1], ['boxWhisker', 2]])
    expect(box.series.map(s => s.pending?.values?.definedName)).toEqual(['_xlchart.v1.0', '_xlchart.v1.1', '_xlchart.v1.2'])
    expect(box.series.every(s => s.pending?.categories === undefined)).toBe(true)
    expect(box.chartEx.series[0].layoutPr).toMatchObject({ quartileMethod: 'exclusive', visibility: { meanLine: false, meanMarker: true, nonoutliers: false, outliers: true } })
    // BW2b omits visibility: the map's absence defaults (§4.1, [U]).
    const bw2b = exGroup(await parse('box-whisker-gapwidth/box-whisker-gapwidth.xlsx', 'xl/charts/chartEx2.xml'))
    expect(bw2b.chartEx.series[0].layoutPr.visibility).toEqual({ connectorLines: true, meanLine: false, meanMarker: true, nonoutliers: false, outliers: true })

    // PA1: clusteredColumn + paretoLine (no dataId; bound through ownerIdx) is Pareto, with a 0..1 percent axis.
    const pa1 = await parse('paretoLine/paretoLine.xlsx')
    const pareto = exGroup(pa1)
    expect(pareto.family).toBe('pareto')
    expect(pareto.series).toHaveLength(1)
    expect(pareto.chartEx.series[0]).toMatchObject({ layout: 'clusteredColumn', dataId: 0, axisIds: ['1'] })
    expect(pareto.chartEx.series[0].layoutPr.binning).toEqual({ intervalClosed: 'r' })
    expect(pareto.chartEx.paretoLine).toMatchObject({ layout: 'paretoLine', ownerIdx: 0, axisIds: ['2'] })
    expect(pareto.chartEx.paretoLine?.dataId).toBeUndefined()
    expect(pa1.axes.map(a => [a.id, a.kind])).toEqual([['0', 'category'], ['1', 'value'], ['2', 'value']])
    expect(pa1.axes[2]).toMatchObject({ min: 0, max: 1, unit: 'percentage' })
    expect(pareto.binding).toEqual({ kind: 'cartesian', category: '0', value: '1' })
    // PA2: the axis kind comes from the scaling child, not the id (ids swapped).
    const pa2 = await parse('pareto-gapwidth/pareto-gapwidth.xlsx')
    expect(pa2.axes.map(a => [a.id, a.kind])).toEqual([['0', 'value'], ['1', 'category'], ['2', 'value']])
    expect(exGroup(pa2).binding).toEqual({ kind: 'cartesian', category: '1', value: '0' })
    expect(exGroup(pa2).chartEx.gapWidth).toBeCloseTo(2.55, 5)

    // HI1 (automatic bins) and HI2 (binSize/underflow/overflow written as @val and attributes): histogram.
    const hi1 = exGroup(await parse('SimpleHistogram/SimpleHistogram.xlsx'))
    expect(hi1.family).toBe('histogram')
    expect(hi1.chartEx.series[0].layoutPr.binning).toEqual({ intervalClosed: 'r' })
    const hi2 = exGroup(await parse('tdf163727_histogram_underflow_overflow/tdf163727_histogram_underflow_overflow.xlsx'))
    expect(hi2.chartEx.series[0].layoutPr.binning).toEqual({ intervalClosed: 'r', binSize: 2, underflow: 11, overflow: 14 })

    // WF1: waterfall, empty subtotals, connector lines by default; WF2 writes connectorLines explicitly.
    const wf1 = await parse('waterfall/waterfall.xlsx')
    expect(wf1.title).toBe('Waterfall')
    expect(wf1.legend).toBe('t')
    expect(exGroup(wf1).family).toBe('waterfall')
    expect(exGroup(wf1).chartEx.series[0].layoutPr).toMatchObject({ subtotals: [], visibility: { connectorLines: true } })
    expect(exGroup(await parse('waterfall2/waterfall2.xlsx')).chartEx.series[0].layoutPr.visibility.connectorLines).toBe(true)

    // FN1 / FN2: funnel; FN2 keeps its hidden series and their names; title from rich text.
    const fn2 = await parse('funnel-pp1/funnel-pp1.pptx', 'ppt/charts/chartEx1.xml')
    expect(fn2.title).toBe('PowerPoint funnel chart')
    expect(exGroup(fn2).family).toBe('funnel')
    expect(exGroup(fn2).series.map(s => s.name)).toEqual(['Legend Entry 1', 'Series 2', 'Series 3'])
    expect(exGroup(fn2).chartEx.series.map(s => [s.dataId, s.hidden, s.formatIdx])).toEqual([[0, false, 0], [1, true, 1], [2, true, 2]])
    const fn1 = await parse('funnel-label-options/funnel-label-options.xlsx')
    expect(exGroup(fn1).family).toBe('funnel')
    // An auto title is the txPr run text (map §1).
    expect(fn1.title).toBe('Chartex label option test')

    // TM / SB: treemap and sunburst (no axes).
    const tm = await parse('treemap/treemap.xlsx')
    expect(exGroup(tm).family).toBe('treemap')
    expect(tm.title).toBe('Treemap')
    expect(tm.axes).toEqual([])
    expect(exGroup(await parse('sunburst/sunburst.xlsx')).family).toBe('sunburst')

    // The XLSX defined names resolve against the workbook through the existing name machinery (map §2.2).
    const wfPkg = await OfficePackage.load(fixture('waterfall/waterfall.xlsx'))
    const names = parseDefinedNames(parseXml((await wfPkg.text('xl/workbook.xml'))!))
    const resolved = resolveChartNames(wf1, 'Sheet1', chartNameResolver(names, [{ name: 'Sheet1', workbookIndex: 0 }]))
    expect(resolved.unsupported).toBeUndefined()
    expect(resolved.groups[0].series[0].pending?.values).toEqual({ sheet: 'Sheet1', r0: 3, c0: 2, r1: 14, c1: 2 })
    expect(isChartExGroup(resolved.groups[0])).toBe(true)

    // Office-shaped synthetic: a dangling dataId, a non-Pareto series without dataId and a paretoLine whose
    // owner is not a clusteredColumn are unreadable.
    for (const series of [
      '<cx:series layoutId="funnel"><cx:dataId val="9" /></cx:series>',
      '<cx:series layoutId="funnel" />',
      '<cx:series layoutId="funnel"><cx:dataId val="0" /></cx:series><cx:series layoutId="paretoLine" ownerIdx="0" />',
      '<cx:series layoutId="clusteredColumn"><cx:dataId val="0" /></cx:series><cx:series layoutId="paretoLine" ownerIdx="3" />',
    ]) {
      const p = await syntheticPart(`<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>${series}</cx:plotAreaRegion></cx:plotArea></cx:chart>`)
      const model = parseChartEx(p.root, theme, p.context)
      expect(model.unsupported, series).toBe('This chart could not be read')
      expect(model.feature, series).toBe('ChartEx')
      p.release()
    }
  })

  it('region map, blocked families and unreadable parts return unsupported so the fallback picture is kept', async () => {
    // A family is unsupported until its family task lands (B2-B5) and registers it as painted.
    const files: Array<[ChartExGroup['family'], string, string?]> = [
      ['waterfall', 'waterfall/waterfall.xlsx'], ['funnel', 'funnel-pp1/funnel-pp1.pptx', 'ppt/charts/chartEx1.xml'], ['histogram', 'SimpleHistogram/SimpleHistogram.xlsx'],
      ['pareto', 'paretoLine/paretoLine.xlsx'], ['boxWhisker', 'boxWhisker/boxWhisker.xlsx'], ['treemap', 'treemap/treemap.xlsx'], ['sunburst', 'sunburst/sunburst.xlsx'],
    ]
    for (const [family, file, path] of files) {
      if (CHARTEX_PAINTED.has(family)) continue
      const p = await nativePart(file, path)
      // The default adapter dispatches the cx root to the ChartEx parser (never the classic one).
      const result = await parseChartPart(p.root, theme, p.context)
      expect(result, file).toMatchObject({ kind: 'unsupported', reason: 'ChartEx charts are not supported yet', feature: 'ChartEx' })
      p.release()
    }
    // Region map: always unsupported, recognised before its geography (cx:binary is never decoded).
    const rm = await nativePart('regionMap/regionMap.xlsx')
    expect(await parseChartPart(rm.root, theme, rm.context)).toMatchObject({ kind: 'unsupported', reason: 'Region map charts are not supported', feature: 'ChartEx', title: 'RegionMap' })
    rm.release()
    // Unreadable parts (Office-shaped synthetic): no series, a dashboard root, an unknown layout, mixed families.
    for (const body of [
      '<cx:chartData /><cx:chart><cx:plotArea /></cx:chart>',
      '<cx:chartData /><cx:dashboard />',
      '<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="spiral"><cx:dataId val="0" /></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>',
      '<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="1"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel"><cx:dataId val="0" /></cx:series><cx:series layoutId="waterfall"><cx:dataId val="0" /></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>',
    ]) {
      const p = await syntheticPart(body)
      expect(await parseChartPart(p.root, theme, p.context), body).toMatchObject({ kind: 'unsupported', reason: 'This chart could not be read', feature: 'ChartEx' })
      p.release()
    }
    // Refusals (Office-shaped synthetic): an analytic level over 65,536 points, an authored binCount over 1,000.
    const col = (layoutPr: string, ptCount = 1) => `<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="${ptCount}"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn"><cx:dataId val="0" /><cx:layoutPr>${layoutPr}</cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>`
    for (const [body, reason] of [
      [col('<cx:binning />', 65537), 'This chart has too many data points to summarize'],
      [col('<cx:binning><cx:binCount val="1001" /></cx:binning>'), 'invalid binning'],
    ] as const) {
      const p = await syntheticPart(body)
      expect(await parseChartPart(p.root, theme, p.context), reason).toMatchObject({ kind: 'refused', reason, feature: 'ChartEx' })
      p.release()
    }
    // A painted family (as later tasks will register it) is native, its PPTX data from the part's own cache.
    const fn2 = await nativePart('funnel-pp1/funnel-pp1.pptx', 'ppt/charts/chartEx1.xml')
    const native = await parseChartExPart(fn2.root, theme, fn2.context, new Set(['funnel']))
    expect(native.kind).toBe('native')
    if (native.kind === 'native') expect(values(native.model.groups[0].series[0].values)).toEqual([4.3, 2.5, 3.5, 4.5])
    fn2.release()

    // End to end: an unsupported native XLSX (region map) keeps its fallback text box. (Waterfall and funnel are
    // painted since B2: tests/charts-chartex-waterfall-funnel.test.ts asserts their native end-to-end results.)
    for (const file of ['regionMap/regionMap.xlsx']) {
      const xlsx = await parseXlsx(await OfficePackage.load(fixture(file)))
      const drawings = xlsx.sheets[0].drawings ?? []
      expect(drawings, file).toHaveLength(1)
      expect(drawings[0].content, `${file}: no chart and no card`).toBeUndefined()
      expect(drawings[0].source.representation, file).toBe('fallback')
    }
  })

  it('the PPTX data falls back to the embedded workbook when the part has no cache', async () => {
    // FN2 with its cx:lvl caches removed (a derived variant of a native part): the values come from its embedded
    // workbook through cx:chartData/cx:externalData (spec 6.4 precedence 2).
    const pkg = await OfficePackage.load(fixture('funnel-pp1/funnel-pp1.pptx'))
    const zip = await JSZip.loadAsync(fixture('funnel-pp1/funnel-pp1.pptx'))
    const xml = (await zip.file('ppt/charts/chartEx1.xml')!.async('string')).replace(/<cx:lvl\b[^>]*>.*?<\/cx:lvl>/g, '')
    zip.file('ppt/charts/chartEx1.xml', xml)
    const stripped = await readPart(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })), 'ppt/charts/chartEx1.xml')
    void pkg
    const parsed = exGroup(parseChartEx(stripped.root, theme, stripped.context))
    expect(parsed.series[0].pending?.values).toEqual({ sheet: 'Sheet1', r0: 1, c0: 1, r1: 4, c1: 1 })
    const result = await parseChartExPart(stripped.root, theme, stripped.context, new Set(['funnel']))
    expect(result.kind).toBe('native')
    if (result.kind === 'native') {
      expect(values(result.model.groups[0].series[0].values)).toEqual([4.3, 2.5, 3.5, 4.5])
      expect(result.model.groups[0].series[0].categories ?? result.model.categories).toEqual(['Thing 1', 'Thing 2', 'Thing 3', 'Thing 4'])
    }
    stripped.release()
  })
})
