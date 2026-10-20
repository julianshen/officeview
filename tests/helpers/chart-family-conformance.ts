/**
 * The family conformance harness (spec 7): every chart family handles an empty
 * series, all-blank values, NaN/∞ (treated as blank), negative values and a
 * single point without throwing. Each case is run through the chart adapter
 * (`parseChartPart`, as an XLSX part and as a DOCX part), stage 1, layout and
 * paint, and must give either a native model drawn without the card or a card
 * with its reason; its diagnostics are collected from the parse and the
 * prepared axes, and must be well formed. An extra case marks one
 * point unavailable and requires the `chart-point-unavailable` diagnostic.
 *
 * A family task registers its family with `describeFamilyConformance`; the
 * family's `chart` builder wraps the given numeric cache in a chart part.
 * Office-shaped synthetic fixtures, never Office evidence.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../../src/core/zip'
import { parseXmlOrdered } from '../../src/core/xml'
import { documentChartContext, chartParseContext } from '../../src/charts/context'
import { parseChartPart } from '../../src/charts/dataPrecedence'
import { parseXlsxChart, placeholderChart } from '../../src/charts/parse'
import { paintXlsxChart } from '../../src/charts/paint'
import { prepareChartAxes } from '../../src/charts/prepare'
import { PointWriter } from '../../src/charts/points'
import type { ChartDiagnostic, ChartModel, ChartPoints } from '../../src/charts/model'

export const CONFORMANCE_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/** A numeric cache (`c:numCache` content) of the given point texts; `undefined` leaves the index out (blank). */
export const conformanceCache = (count: number, points: Array<string | undefined>) =>
  `<c:numRef><c:f>Sheet1!$B$2:$B$${count + 1}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${count}"/>${points.map((v, i) => v === undefined ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`

export type ConformanceCase = 'no series' | 'empty series' | 'all blank' | 'NaN and infinity' | 'negative values' | 'single point' | 'unavailable point'
/** The numeric cache of each case's series; `null`: the group has no series. */
export const CONFORMANCE_CASES: Readonly<Record<Exclude<ConformanceCase, 'unavailable point'>, string | null>> = {
  'no series': null,
  'empty series': conformanceCache(0, []),
  'all blank': conformanceCache(3, [undefined, '', undefined]),
  'NaN and infinity': conformanceCache(4, ['NaN', 'Infinity', '-Infinity', '1e999']),
  'negative values': conformanceCache(3, ['-3', '-1', '-2']),
  'single point': conformanceCache(1, ['5']),
}

export interface ConformanceFamily {
  /** The family's name in test titles. */
  name: string
  /**
   * A whole `c:chartSpace` (namespaces `CONFORMANCE_NS`) with one group of this family. `cache` is the numeric
   * source (a `c:numRef` with its cache) for every numeric role of its one series; `null` means no series.
   */
  chart(cache: string | null): string
  /** Per case, the representation the family must give (default: either) and diagnostic codes it must emit. */
  expect?: Partial<Record<ConformanceCase, { kind?: 'native' | 'card'; diagnostics?: string[] }>>
}

export interface ConformanceRun { kind: 'native' | 'card'; reason?: string; diagnostics: ChartDiagnostic[]; card: boolean }

let emptyPackage: Promise<OfficePackage> | undefined
const pkg = () => emptyPackage ??= new JSZip().generateAsync({ type: 'uint8array' }).then(bytes => OfficePackage.load(bytes))

/** Paints a model (or its card) on a canvas; true when the placeholder card was painted. */
function paint(model: ChartModel, w = 400, h = 300): { card: boolean } {
  const ctx = createCanvas(w, h).getContext('2d')
  let card = false
  const fillRect = ctx.fillRect.bind(ctx)
  ctx.fillRect = (x: number, y: number, rw: number, rh: number) => { if (String(ctx.fillStyle).toLowerCase() === '#f7f7f7' && rw === w && rh === h) card = true; fillRect(x, y, rw, rh) }
  paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, w, h)
  return { card }
}
const collect = (model: ChartModel): ChartDiagnostic[] => [...model.diagnostics ?? [], ...model.preparedAxes?.diagnostics ?? []]

/** Every numeric role of every series with point 0 made unavailable (a formula cell with no saved value). */
function withUnavailable(model: ChartModel): ChartModel {
  const mark = (points: ChartPoints | undefined): ChartPoints | undefined => {
    if (!points) return points
    const out = new PointWriter(Math.max(points.length, 1))
    for (let i = 0; i < points.length; i++) if (points.states[i] === 1) out.value(i, points.values[i])
    out.unavailable(0)
    return out.done()
  }
  return { ...model, preparedAxes: undefined, groups: model.groups.map(g => ({ ...g, series: g.series.map(s => ({ ...s, values: mark(s.values)!, ...(s.xValues ? { xValues: mark(s.xValues) } : {}), ...(s.bubbleSizes ? { bubbleSizes: mark(s.bubbleSizes) } : {}) })) })) }
}

/** Runs one case of a family through the adapter (as `partPath`), stage 1, layout and paint. */
export async function runConformanceCase(family: ConformanceFamily, name: ConformanceCase, partPath = '/xl/charts/chart1.xml'): Promise<ConformanceRun> {
  if (name === 'unavailable point') {
    const parsed = parseXlsxChart(parseXmlOrdered(family.chart(CONFORMANCE_CASES['single point'])), undefined)
    if (!parsed || parsed.unsupported) return { kind: 'card', reason: parsed?.unsupported, diagnostics: parsed ? collect(parsed) : [], card: true }
    const model = prepareChartAxes(withUnavailable(parsed))
    const painted = paint(model)
    return { kind: 'native', diagnostics: collect(model), card: painted.card }
  }
  const doc = documentChartContext('desktop')
  const context = chartParseContext(doc, await pkg(), partPath)
  const result = await parseChartPart(parseXmlOrdered(family.chart(CONFORMANCE_CASES[name])), undefined, context)
  if (result.kind === 'native') {
    const painted = paint(result.model)
    return { kind: 'native', diagnostics: collect(result.model), card: painted.card }
  }
  const painted = paint(placeholderChart(result.title, result.reason))
  return { kind: 'card', reason: result.reason, diagnostics: [], card: painted.card }
}

/** Registers the conformance cases of one family (spec 7). */
export function describeFamilyConformance(family: ConformanceFamily): void {
  describe(`${family.name} family conformance (spec 7)`, () => {
    const cases: ConformanceCase[] = [...Object.keys(CONFORMANCE_CASES) as Array<Exclude<ConformanceCase, 'unavailable point'>>, 'unavailable point']
    for (const name of cases) for (const part of name === 'unavailable point' ? ['/xl/charts/chart1.xml'] : ['/xl/charts/chart1.xml', '/word/charts/chart1.xml']) {
      it(`${family.name}: ${name} (${part.split('/')[1]}) does not throw and gives a native chart or a card with its diagnostics`, async () => {
        const run = await runConformanceCase(family, name, part)
        const want = family.expect?.[name]
        if (want?.kind) expect(run.kind, `${family.name} ${name}`).toBe(want.kind)
        if (run.kind === 'native') expect(run.card, 'a native result is drawn without the card').toBe(false)
        else {
          expect(run.reason, 'a card carries its reason').toMatch(/\S/)
          expect(run.card, 'a card result paints the card').toBe(true)
        }
        for (const d of run.diagnostics) { expect(d.code).toMatch(/^chart-[a-z0-9-]+$/); expect(d.reason).toMatch(/\S/) }
        const codes = run.diagnostics.map(d => d.code)
        for (const code of want?.diagnostics ?? []) expect(codes, `${family.name} ${name}`).toContain(code)
        // A family that is a card in every case (surface3D) has no point to mark; every drawn family must diagnose it.
        if (name === 'unavailable point' && want?.kind !== 'card') expect(codes, `${family.name} unavailable`).toContain('chart-point-unavailable')
      })
    }
  })
}

/** Chart-space wrapper shared by the family builders. */
export const conformanceSpace = (plot: string) => `<c:chartSpace ${CONFORMANCE_NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}</c:plotArea><c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`
const CAT_VAL = '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>'
const VAL_VAL = '<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>'
/** Office's contour wiring: category and series axes cross the value axis. */
const CAT_VAL_SER = '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx><c:serAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:serAx>'
const ser = (inner: string) => `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>S1</c:v></c:tx>${inner}</c:ser>`
const categorySeries = (cache: string | null) => cache === null ? '' : ser(`<c:val>${cache}</c:val>`)

/** The families the engine draws natively at A5a; later family tasks add theirs. */
export const CLASSIC_FAMILIES: ConformanceFamily[] = [
  { name: 'column', chart: c => conformanceSpace(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${categorySeries(c)}<c:axId val="1"/><c:axId val="2"/></c:barChart>${CAT_VAL}`) },
  { name: 'stacked bar', chart: c => conformanceSpace(`<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>${categorySeries(c)}<c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${CAT_VAL.replace('"b"', '"l"').replace('<c:axPos val="l"/><c:majorGridlines/>', '<c:axPos val="b"/><c:majorGridlines/>')}`) },
  { name: 'line', chart: c => conformanceSpace(`<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${categorySeries(c)}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${CAT_VAL}`) },
  { name: 'area', chart: c => conformanceSpace(`<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${categorySeries(c)}<c:axId val="1"/><c:axId val="2"/></c:areaChart>${CAT_VAL}`) },
  { name: 'pie', chart: c => conformanceSpace(`<c:pieChart><c:varyColors val="1"/>${categorySeries(c)}<c:firstSliceAng val="0"/></c:pieChart>`) },
  { name: 'doughnut', chart: c => conformanceSpace(`<c:doughnutChart><c:varyColors val="1"/>${categorySeries(c)}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`) },
  { name: 'scatter', chart: c => conformanceSpace(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${c === null ? '' : ser(`<c:xVal>${c}</c:xVal><c:yVal>${c}</c:yVal>`)}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>${VAL_VAL}`) },
  { name: 'bubble', chart: c => conformanceSpace(`<c:bubbleChart><c:varyColors val="0"/>${c === null ? '' : ser(`<c:xVal>${c}</c:xVal><c:yVal>${c}</c:yVal><c:bubbleSize>${c}</c:bubbleSize><c:bubble3D val="0"/>`)}<c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>${VAL_VAL}`) },
  { name: 'radar', chart: c => conformanceSpace(`<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${categorySeries(c)}<c:axId val="1"/><c:axId val="2"/></c:radarChart>${CAT_VAL}`) },
  { name: 'contour', chart: c => conformanceSpace(`<c:surfaceChart><c:wireframe val="0"/>${categorySeries(c)}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:surfaceChart>${CAT_VAL_SER}`) },
  { name: 'wireframe contour', chart: c => conformanceSpace(`<c:surfaceChart><c:wireframe val="1"/>${categorySeries(c)}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:surfaceChart>${CAT_VAL_SER}`) },
  { name: 'filled radar', chart: c => conformanceSpace(`<c:radarChart><c:radarStyle val="filled"/><c:varyColors val="0"/>${categorySeries(c)}<c:axId val="1"/><c:axId val="2"/></c:radarChart>${CAT_VAL}`) },
]

/** surface3D (spec 4.1: perspective is a non-goal): a card in every case, never a native chart. */
export const SURFACE3D_FAMILY: ConformanceFamily = {
  name: 'surface3D',
  chart: c => conformanceSpace(`<c:surface3DChart><c:wireframe val="0"/>${categorySeries(c)}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:surface3DChart>${CAT_VAL_SER}`),
  expect: { 'no series': { kind: 'card' }, 'empty series': { kind: 'card' }, 'all blank': { kind: 'card' }, 'NaN and infinity': { kind: 'card' }, 'negative values': { kind: 'card' }, 'single point': { kind: 'card' }, 'unavailable point': { kind: 'card' } },
}
