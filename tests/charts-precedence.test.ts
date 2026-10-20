/**
 * A3b: per-role data precedence for DOCX/PPTX charts (spec 6.4). Saved data
 * (a cache or a literal, even partial) wins; a role without saved data is
 * read from the chart's embedded workbook; a role with neither is
 * unavailable, and a missing required role makes the chart unsupported.
 * Fixtures are Office-shaped synthetic, never Office evidence.
 */
import { describe, expect, it, vi } from 'vitest'
import type JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { documentChartContext } from '../src/charts/context'
import { pointList } from '../src/charts/points'
import { paintXlsxChart } from '../src/charts/paint'
import type { ChartModel } from '../src/charts/model'
import type { DrawingCoverageEntry } from '../src/drawing/coverage'
import { buildChartDocx, buildChartPptx, cell, embeddedWorkbook, officeChartWorkbook, officeChartXml, row, type ChartEmbedSpec, type ChartFixture } from './helpers/chart-fixtures'

/** Every inflate of an entry, by name: the embedding (outer) and its inner parts. */
const streams = vi.hoisted(() => ({ names: [] as string[] }))
vi.mock('../src/charts/read', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/charts/read')>()
  return { ...actual, streamEntry: (file: JSZip.JSZipObject, ...rest: Parameters<typeof actual.streamEntry> extends [unknown, ...infer R] ? R : never) => { streams.names.push(file.name); return actual.streamEntry(file, ...rest) } }
})

const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
/** A bar chart part (Office's layout) around the given series, with `c:externalData` unless `external` is false. */
function chartXml(series: string, options: { family?: 'bar' | 'scatter' | 'line'; externalData?: boolean } = {}): string {
  const family = options.family ?? 'bar'
  const group = family === 'bar' ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}<c:gapWidth val="219"/><c:axId val="1"/><c:axId val="2"/></c:barChart>`
    : family === 'line' ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>`
      : `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>`
  const axes = family === 'scatter'
    ? '<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>'
    : '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>'
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace ${NS}><c:date1904 val="0"/><c:lang val="en-US"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${group}${axes}</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>${options.externalData === false ? '' : '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'}</c:chartSpace>`
}
const numCache = (count: number, points: Array<[number, number | string]>) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${count}"/>${points.map(([i, v]) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`
const strCache = (labels: string[]) => `<c:strCache><c:ptCount val="${labels.length}"/>${labels.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache>`
const numRef = (f: string, cache = '') => `<c:numRef><c:f>${f}</c:f>${cache}</c:numRef>`
const strRef = (f: string, cache = '') => `<c:strRef><c:f>${f}</c:f>${cache}</c:strRef>`
const ser = (i: number, parts: { tx?: string; cat?: string; val?: string; xVal?: string; yVal?: string }) =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${parts.tx ? `<c:tx>${parts.tx}</c:tx>` : ''}${parts.cat ? `<c:cat>${parts.cat}</c:cat>` : ''}${parts.val ? `<c:val>${parts.val}</c:val>` : ''}${parts.xVal ? `<c:xVal>${parts.xVal}</c:xVal>` : ''}${parts.yVal ? `<c:yVal>${parts.yVal}</c:yVal>` : ''}</c:ser>`

const values = (chart: ChartModel, series = 0, group = 0) => pointList(chart.groups[group].series[series].values).map(p => p.state === 'value' ? p.v : p.state === 'error' ? p.code : p.state)
const embeddingStreams = () => streams.names.filter(name => name.includes('embeddings/') || name.startsWith('xl/'))

interface Parsed { charts: ChartModel[]; coverage: DrawingCoverageEntry[]; workspace: ReturnType<typeof documentChartContext>['workspace'] }
async function parse(format: 'docx' | 'pptx', fixture: ChartFixture): Promise<Parsed> {
  const pkg = await OfficePackage.load(fixture.bytes)
  const doc = documentChartContext('desktop')
  streams.names.length = 0
  if (format === 'docx') {
    const parsed = await parseDocx(pkg, { chartContext: doc })
    const charts = parsed.sections.flatMap(s => s.paragraphs).flatMap(p => p.images ?? []).flatMap(i => i.drawing?.kind === 'xlsxChart' ? [i.drawing.chart] : [])
    return { charts, coverage: (parsed.drawingCoverage ?? []).filter(e => e.scope === 'original'), workspace: doc.workspace }
  }
  const parsed = await parsePptx(pkg, { chartContext: doc })
  const charts = parsed.slides.flatMap(s => s.shapes).flatMap(s => s.content?.kind === 'xlsxChart' ? [s.content.chart] : [])
  return { charts, coverage: (parsed.drawingCoverage ?? []).filter(e => e.scope === 'original'), workspace: doc.workspace }
}
const build = (format: 'docx' | 'pptx', charts: ChartEmbedSpec[]) => format === 'docx' ? buildChartDocx(charts) : buildChartPptx(charts)

/** The series-colour bars a native paint draws (accent 1), as [width, height] rounded. */
function bars(chart: ChartModel): Array<[number, number]> {
  const ctx = createCanvas(400, 300).getContext('2d'), out: Array<[number, number]> = []
  const fillRect = ctx.fillRect.bind(ctx)
  ctx.fillRect = (x: number, y: number, w: number, h: number) => { if (ctx.fillStyle === '#4472c4') out.push([Math.round(w), Math.round(h)]); fillRect(x, y, w, h) }
  paintXlsxChart(chart, ctx as unknown as CanvasRenderingContext2D, 400, 300)
  return out
}

describe('data precedence (spec 6.4)', () => {
  it('cacheless values role renders natively from the embedded workbook', async () => {
    for (const format of ['docx', 'pptx'] as const) {
      const { charts, coverage } = await parse(format, build(format, [{ workbook: officeChartWorkbook([4.3, 2.5]) }]))
      expect(charts, format).toHaveLength(1)
      const chart = charts[0]
      // A native model: not a card, not a fallback picture; every role filled from the workbook.
      expect(chart.unsupported, format).toBeUndefined()
      expect(chart.refused, format).toBeUndefined()
      expect(chart.groups[0].series[0].pending, format).toBeUndefined()
      expect(chart.categoryPending, format).toBeUndefined()
      expect(values(chart), format).toEqual([4.3, 2.5])
      expect(chart.categories, format).toEqual(['A', 'B'])
      expect(chart.groups[0].series[0].name, format).toBe('Series 1')
      expect(coverage, format).toHaveLength(1)
      expect(coverage[0], format).toMatchObject({ selectedRepresentation: 'chart' })
      // Distinct native geometry: two bars of one width whose heights keep the filled values' ratio.
      const drawn = bars(chart)
      expect(drawn, format).toHaveLength(2)
      expect(drawn[0][0], format).toBe(drawn[1][0])
      expect(drawn[0][1] / drawn[1][1], format).toBeCloseTo(4.3 / 2.5, 1)
      expect(embeddingStreams().some(name => name === 'xl/worksheets/sheet1.xml'), format).toBe(true)
    }
  })

  it('cached-only chart never opens the embedding', async () => {
    const cached = chartXml(ser(0, { tx: strRef('Sheet1!$B$1', strCache(['Saved'])), cat: strRef('Sheet1!$A$2:$A$3', strCache(['x', 'y'])), val: numRef('Sheet1!$B$2:$B$3', numCache(2, [[0, 7], [1, 9]])) }))
    for (const format of ['docx', 'pptx'] as const) {
      const { charts, workspace } = await parse(format, build(format, [{ chartXml: cached, workbook: officeChartWorkbook([1, 2]) }]))
      expect(charts[0].unsupported, format).toBeUndefined()
      expect(values(charts[0]), format).toEqual([7, 9])
      expect(charts[0].groups[0].series[0].name, format).toBe('Saved')
      expect(embeddingStreams(), format).toEqual([])
      expect(workspace.used, format).toBeGreaterThan(0)
    }
  })

  it('saved data wins: caches and literals; sparse indices kept; ptCount=0 present-empty', async () => {
    // Series 0: a partial cache (idx 1 missing) beside a reference whose cells differ: the cache wins, the gap is blank.
    // Series 1: a present, empty cache (ptCount=0): present, never filled from the workbook.
    // Series 2: a literal. Series 3: cacheless, filled from the workbook.
    const xml = chartXml([
      ser(0, { val: numRef('Sheet1!$B$2:$B$4', numCache(3, [[0, 10], [2, 30]])) }),
      ser(1, { val: numRef('Sheet1!$B$2:$B$4', numCache(0, [])) }),
      ser(2, { val: `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="1"><c:v>5</c:v></c:pt></c:numLit>` }),
      ser(3, { val: numRef('Sheet1!$C$2:$C$4') }),
    ].join(''))
    const workbook = embeddedWorkbook({ sheets: [{ name: 'Sheet1', data: [row(2, cell('B2', 1), cell('C2', 100)), row(3, cell('B3', 2), cell('C3', 200)), row(4, cell('B4', 3), cell('C4', 300))].join('') }] })
    for (const format of ['docx', 'pptx'] as const) {
      const { charts } = await parse(format, build(format, [{ chartXml: xml, workbook }]))
      const chart = charts[0]
      expect(chart.unsupported, format).toBeUndefined()
      expect(values(chart, 0), format).toEqual([10, 'blank', 30])
      expect(chart.groups[0].series[1].values.length, format).toBe(0)
      expect(chart.groups[0].series[1].pending, format).toBeUndefined()
      expect(values(chart, 2), format).toEqual(['blank', 5])
      expect(values(chart, 3), format).toEqual([100, 200, 300])
    }
  })

  it('literal-only and mixed literal-reference series', async () => {
    const literal = chartXml(ser(0, { tx: '<c:v>Literal</c:v>', cat: '<c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>p</c:v></c:pt><c:pt idx="1"><c:v>q</c:v></c:pt></c:strLit>', val: '<c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit>' }), { externalData: false })
    // Literal values and name, categories and a second series' values by reference (no caches).
    const mixed = chartXml([
      ser(0, { tx: '<c:v>Lit</c:v>', cat: strRef('Sheet1!$A$2:$A$3'), val: '<c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit>' }),
      ser(1, { tx: strRef('Sheet1!$B$1'), cat: strRef('Sheet1!$A$2:$A$3'), val: numRef('Sheet1!$B$2:$B$3') }),
    ].join(''))
    for (const format of ['docx', 'pptx'] as const) {
      const lit = await parse(format, build(format, [{ chartXml: literal }]))
      expect(lit.charts[0].unsupported, format).toBeUndefined()
      expect(values(lit.charts[0]), format).toEqual([3, 4])
      expect(lit.charts[0].categories, format).toEqual(['p', 'q'])
      expect(embeddingStreams(), format).toEqual([])

      const mix = await parse(format, build(format, [{ chartXml: mixed, workbook: officeChartWorkbook([8, 9], ['m', 'n']) }]))
      const chart = mix.charts[0]
      expect(chart.unsupported, format).toBeUndefined()
      expect(chart.categories, format).toEqual(['m', 'n'])
      expect(values(chart, 0), format).toEqual([3, 4])
      expect(chart.groups[0].series[0].name, format).toBe('Lit')
      expect(values(chart, 1), format).toEqual([8, 9])
      expect(chart.groups[0].series[1].name, format).toBe('Series 1')
    }
  })

  it('a missing required role makes the chart unsupported', async () => {
    const reason = "This chart's data is not available"
    const cases: Array<[string, ChartEmbedSpec, string?]> = [
      // No embedding part (the relationship dangles).
      ['dangling embedding', { chartXml: officeChartXml() }],
      // No c:externalData at all.
      ['no externalData', { chartXml: chartXml(ser(0, { val: numRef('Sheet1!$B$2:$B$3') }), { externalData: false }) }],
      // An external target.
      ['external target', { chartXml: officeChartXml(), external: true, target: 'file:///C:/book.xlsx' }],
      // The referenced sheet is not in the workbook.
      ['missing sheet', { chartXml: chartXml(ser(0, { val: numRef('Other!$B$2:$B$3') })), workbook: officeChartWorkbook() }],
      // A values element with neither saved data nor a reference.
      ['no source', { chartXml: chartXml(ser(0, { val: '<c:numRef/>' })), workbook: officeChartWorkbook() }],
      // Scatter x values referenced but unreadable (sheet missing): required for scatter.
      ['scatter x', { chartXml: chartXml(ser(0, { xVal: numRef('Gone!$A$2:$A$3'), yVal: numRef('Sheet1!$B$2:$B$3') }), { family: 'scatter' }), workbook: officeChartWorkbook() }],
    ]
    for (const format of ['docx', 'pptx'] as const) {
      for (const [label, spec] of cases) {
        const { charts, coverage } = await parse(format, build(format, [spec]))
        expect(charts, `${format} ${label}`).toHaveLength(1)
        expect(charts[0].unsupported, `${format} ${label}`).toBe(reason)
        expect(coverage[0], `${format} ${label}`).toMatchObject({ selectedRepresentation: 'chart-placeholder' })
      }
      // An optional role that is unavailable (categories on a missing sheet) keeps the chart: categories default.
      const optional = await parse(format, build(format, [{ chartXml: chartXml(ser(0, { cat: strRef('Gone!$A$2:$A$3'), val: numRef('Sheet1!$B$2:$B$3') })), workbook: officeChartWorkbook([1, 2]) }]))
      expect(optional.charts[0].unsupported, format).toBeUndefined()
      expect(values(optional.charts[0]), format).toEqual([1, 2])
      expect(optional.charts[0].categories, format).toEqual([])
      expect(optional.charts[0].diagnostics?.map(d => d.code), format).toContain('chart-data-sheet-missing')
      // Mixed sources: one series cached, one from the workbook, one unavailable required role -> unsupported.
      const mixed = chartXml([ser(0, { val: numRef('Sheet1!$B$2:$B$3', numCache(2, [[0, 1], [1, 2]])) }), ser(1, { val: numRef('Nope!$B$2:$B$3') })].join(''))
      const failed = await parse(format, build(format, [{ chartXml: mixed, workbook: officeChartWorkbook() }]))
      expect(failed.charts[0].unsupported, format).toBe(reason)
    }
  })

  it('the embedded workbook epoch is the chart epoch when c:date1904 is absent (D4)', async () => {
    const xml = (date1904: string) => chartXml(ser(0, { cat: numRef('Sheet1!$A$2:$A$3'), val: numRef('Sheet1!$B$2:$B$3') })).replace('<c:date1904 val="0"/>', date1904)
    const workbook = embeddedWorkbook({ date1904: true, cellXfs: [0, 14], sheets: [{ name: 'Sheet1', data: [row(2, cell('A2', 0, { s: 1 }), cell('B2', 1)), row(3, cell('A3', 1, { s: 1 }), cell('B3', 2))].join('') }] })
    const absent = await parse('docx', buildChartDocx([{ chartXml: xml(''), workbook }]))
    expect(absent.charts[0].unsupported).toBeUndefined()
    expect(absent.charts[0].workbookEpoch).toBe('1904')
    // Serial 0 in the 1904 system is 1904-01-01.
    expect(absent.charts[0].categories).toEqual(['1/1/1904', '1/2/1904'])
  })
})
