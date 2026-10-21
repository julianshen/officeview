// @vitest-environment node
/**
 * Task C1: the chart robustness slice (spec 7, 8 "Robustness tests"). Every test asserts that nothing throws, the
 * expected diagnostics, and the allocation/work counters (chart workspace use and peak, path segments), not just that
 * something was painted:
 *
 * - retention: 200 charts in one document stay within the 32 MiB mobile chart workspace; later charts become cards;
 * - torture documents per format, split by damage class: (a) every family plus ChartEx, model3D and OMML; (b) a huge
 *   embedded workbook and a corrupt embedded zip (XLSX, which has no embedding: an oversized chart part, a malformed
 *   one and a whole-column reference); (c) an external `externalData` link and names/references into another
 *   workbook; (d) empty, NaN and infinite data;
 * - repeated paints and cache churn release their layout charges;
 * - the path, label-byte and contour-grid bounds drop elements with diagnostics (with A4c's parked admission-cost
 *   checks: a straight scatter series without markers costs one segment per point, markers two, smoothing three).
 *
 * Every document here is Office-shaped synthetic (src/testdata/chart-gallery.ts, tests/helpers), never Office evidence.
 * CI is Linux on 2 CPUs: no wall-clock assertion anywhere; sparse caches keep the 200-chart documents small.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { chartParseContext, documentChartContext, type ChartDocumentContext } from '../src/charts/context'
import { parseXlsxChart } from '../src/charts/parse'
import { applyDataPrecedence } from '../src/charts/dataPrecedence'
import { readChartPart } from '../src/charts/read'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart, releaseChartLayouts, releaseWorkspaceLayouts } from '../src/charts/layout'
import { releaseChartModel } from '../src/charts/own'
import { prepareChartAxes } from '../src/charts/prepare'
import { PointWriter } from '../src/charts/points'
import { parseChartEx } from '../src/charts/chartex/parse'
import { CHART_WORKSPACE_LIMIT } from '../src/charts/workspace'
import { chartDescription } from '../src/xlsx/chartDescription'
import type { ChartExGroup, ChartModel } from '../src/charts/model'
import type { DrawingCoverageEntry } from '../src/drawing/coverage'
import { buildChartDocument, chartExGallery, chartSpace, galleryHierarchyWorkbook, GALLERY_HIERARCHY, series as gallerySeries, type ChartDocFormat, type ChartDocObject } from '../src/testdata/chart-gallery'
import { CLASSIC_FAMILIES, CONFORMANCE_CASES, CONFORMANCE_NS, OF_PIE_FAMILY, SURFACE3D_FAMILY, conformanceCache, type ConformanceFamily } from './helpers/chart-family-conformance'
import { cell, embeddedWorkbook, officeChartWorkbook, officeChartXml, row } from './helpers/chart-fixtures'

const FORMATS: ChartDocFormat[] = ['docx', 'pptx', 'xlsx']
const MiB = 1024 * 1024
const MEMORY = 'This chart needs more memory than is available for charts'
const UNAVAILABLE = "This chart's data is not available"
const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'

/** Deterministic incompressible bytes. */
function noise(length: number, seed = 1): Uint8Array {
  const out = new Uint8Array(length)
  let x = seed >>> 0 || 1
  for (let i = 0; i < length; i++) { x ^= x << 13; x ^= x >>> 17; x ^= x << 5; out[i] = x & 0xff }
  return out
}
const png = () => { const c = createCanvas(4, 4); const g = c.getContext('2d'); g.fillStyle = '#00ff00'; g.fillRect(0, 0, 4, 4); return new Uint8Array(c.toBuffer('image/png')) }

/** A document context whose workspace records its peak use (every reservation, admitted or not). */
function meteredContext(profile: 'desktop' | 'mobile'): ChartDocumentContext & { peak(): number } {
  const context = documentChartContext(profile)
  let peak = 0
  const reserve = context.workspace.reserve.bind(context.workspace)
  context.workspace.reserve = (bytes, tag) => { const ok = reserve(bytes, tag); peak = Math.max(peak, context.workspace.used); return ok }
  return Object.assign(context, { peak: () => peak })
}

interface Opened { charts: Array<ChartModel | undefined>; coverage: DrawingCoverageEntry[]; pkg: OfficePackage }
/** Parses a document with `context` and lists, per placed object in document order, its chart model (or card). */
async function open(format: ChartDocFormat, bytes: Uint8Array, context: ChartDocumentContext): Promise<Opened> {
  const pkg = await OfficePackage.load(bytes)
  if (format === 'docx') {
    const doc = await parseDocx(pkg, { chartContext: context })
    const images = doc.sections.flatMap(s => s.paragraphs).flatMap(p => p.images ?? [])
    return { charts: images.map(i => i.drawing?.kind === 'xlsxChart' ? i.drawing.chart : undefined), coverage: (doc.drawingCoverage ?? []).filter(e => e.scope === 'original'), pkg }
  }
  if (format === 'pptx') {
    const doc = await parsePptx(pkg, { chartContext: context })
    const shapes = doc.slides.flatMap(s => s.shapes)
    return { charts: shapes.map(s => s.content?.kind === 'xlsxChart' ? s.content.chart : undefined), coverage: (doc.drawingCoverage ?? []).filter(e => e.scope === 'original'), pkg }
  }
  const doc = await parseXlsx(pkg, { chartContext: context })
  return { charts: (doc.sheets[0].drawings ?? []).map(d => d.content?.kind === 'xlsxChart' ? d.content.chart : undefined), coverage: (doc.drawingCoverage ?? []).filter(e => e.scope === 'original'), pkg }
}
/** Paints a model at w x h on a real canvas; true when the placeholder card was painted. */
function paint(model: ChartModel, context?: ChartDocumentContext, w = 320, h = 220): boolean {
  const ctx = createCanvas(w, h).getContext('2d')
  let card = false
  const fillRect = ctx.fillRect.bind(ctx)
  ctx.fillRect = (x: number, y: number, rw: number, rh: number) => { if (String(ctx.fillStyle).toLowerCase() === '#f7f7f7' && rw === w && rh === h) card = true; fillRect(x, y, rw, rh) }
  paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, w, h, undefined, context)
  return card
}
/** Releases every model and the workspace's layouts; the document facet must then hold nothing. */
function releaseAll(charts: Array<ChartModel | undefined>, context: ChartDocumentContext): void {
  releaseWorkspaceLayouts(context.workspace)
  for (const chart of charts) if (chart) releaseChartModel(chart)
}
/** The adapter's own diagnostics for one DOCX/PPTX chart part (the document keeps only the card's reason). */
async function adapterDiagnostics(pkg: OfficePackage, partPath: string, profile: 'desktop' | 'mobile' = 'mobile'): Promise<string[]> {
  const context = chartParseContext(documentChartContext(profile), pkg, partPath)
  const read = await readChartPart(pkg, partPath, context)
  if (read?.kind !== 'tree') return [`read:${read?.kind}`]
  try {
    const parsed = parseXlsxChart(read.root, undefined, context)
    if (!parsed) return ['not-a-chart']
    const model = await applyDataPrecedence(parsed, read.root, context)
    const codes = (model.diagnostics ?? []).map(d => d.code)
    releaseChartModel(model)
    return codes
  } finally {
    read.release()
    expect(context.workspace.used, `${partPath}: the adapter refunds everything`).toBe(0)
  }
}

afterEach(() => { /* every test releases its own models; nothing is shared */ })

// ---- Retention ---------------------------------------------------------------------------------------------------

describe('retention (spec 6.4, 7: per-document chart retention 64 / 32 MiB)', () => {
  /** A line chart of 64 series whose caches declare 4,096 points but store two (sparse, so the XML stays small):
   * the point storage of a full display-sized chart is charged, about 2.4 MB, before it is built. */
  const heavy = (k: number) => {
    const cache = `<c:numRef><c:f>Sheet1!$B$2:$B$4097</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4096"/><c:pt idx="0"><c:v>${k}</c:v></c:pt><c:pt idx="4095"><c:v>2</c:v></c:pt></c:numCache></c:numRef>`
    const ser = Array.from({ length: 64 }, (_, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:val>${cache}</c:val></c:ser>`).join('')
    return `<c:chartSpace ${CONFORMANCE_NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`
  }

  it.each(FORMATS)('200 charts stay within the chart workspace; later charts become cards (%s)', async format => {
    const context = meteredContext('mobile')
    const bytes = await buildChartDocument(format, Array.from({ length: 200 }, (_, i) => ({ name: `Chart ${i + 1}`, part: heavy(i) })), { columns: 4, rows: 5, width: 160, height: 100 })
    const { charts, coverage } = await open(format, bytes, context)
    expect(charts).toHaveLength(200)
    const limit = CHART_WORKSPACE_LIMIT.mobile
    // Every chart is either native (all 64 series) or a memory card: once the facet is full, a chart is refused before
    // its storage is built. DOCX and XLSX admit in document order (natives first); PPTX reads its slides' charts in
    // its own order, so only the count is pinned there.
    const native = charts.filter(c => c && !c.unsupported).length
    expect(native).toBeGreaterThan(0)
    expect(charts.every(c => c && (c.unsupported === MEMORY || (!c.unsupported && c.groups[0].series.length === 64)))).toBe(true)
    if (format !== 'pptx') expect(charts.findIndex(c => c?.unsupported !== undefined)).toBe(native)
    // Each native model retains the same charge; together they fit, one more would not.
    const retained = context.workspace.used, perChart = retained / native
    expect(perChart).toBeGreaterThan(64 * 4096 * 9)
    expect(retained).toBeLessThanOrEqual(limit)
    expect(retained + perChart).toBeGreaterThan(limit)
    expect(native).toBe(Math.floor(limit / perChart))
    expect(context.peak()).toBeLessThanOrEqual(limit)
    // The cards are diagnosed as refusals in the coverage, never silently blank.
    const refused = coverage.filter(e => e.reason === 'chart-refused')
    expect(refused).toHaveLength(200 - native)
    expect(refused.every(e => e.selectedRepresentation === 'chart-placeholder')).toBe(true)
    // Painting every chart in the same document facet stays within the limit; a layout the facet cannot admit is
    // drawn as a card instead of being built.
    let cards = 0
    for (const chart of charts) if (paint(chart!, context, 160, 100)) cards++
    expect(cards).toBeGreaterThanOrEqual(200 - native)
    expect(context.peak()).toBeLessThanOrEqual(limit)
    releaseAll(charts, context)
    expect(context.workspace.used).toBe(0)
  }, 60_000)
})

// ---- Torture documents by damage class -----------------------------------------------------------------------------

const HIERARCHY_CELLS = [{ r: 1, cells: [] }, ...GALLERY_HIERARCHY.map(([a, b, v], i) => ({ r: i + 2, cells: [{ ref: `A${i + 2}`, t: 'str' as const, v: a }, { ref: `B${i + 2}`, t: 'str' as const, v: b }, { ref: `C${i + 2}`, v }] }))]
const HIERARCHY_NAMES = [{ name: '_xlchart.v1.0', formula: 'Sheet1!$A$2:$B$9' }, { name: '_xlchart.v1.1', formula: 'Sheet1!$C$2:$C$9' }]
/** A minimal ChartEx part of `layout` with one cached value dimension. */
const chartEx = (layout: string, extra = '') => `<cx:chartSpace xmlns:cx="${CX}"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:lvl ptCount="3"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt><cx:pt idx="2">3</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${layout}"><cx:dataId val="0"/>${extra}</cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`
const CLASSIC: ConformanceFamily[] = [...CLASSIC_FAMILIES, OF_PIE_FAMILY]

describe('torture (a): every family plus ChartEx, model3D and OMML, per format', () => {
  it.each(FORMATS)('%s: natives, cards and fallbacks as specified, bounded and fully refunded', async format => {
    const workbook = format === 'xlsx' ? undefined : await galleryHierarchyWorkbook()
    const objects: ChartDocObject[] = [
      ...CLASSIC.map(family => ({ name: family.name, part: family.chart(conformanceCache(3, ['3', '1', '2'])) })),
      { name: 'surface3D', part: SURFACE3D_FAMILY.chart(conformanceCache(3, ['3', '1', '2'])) },
      ...chartExGallery(format !== 'xlsx').map(c => ({ name: c.name, part: c.part, ...(c.hierarchy && workbook ? { workbook } : {}) })),
      // A ChartEx family the engine does not paint yet (B3 pending merge) keeps its fallback picture.
      { name: 'histogram', part: chartEx('clusteredColumn', '<cx:layoutPr><cx:binning intervalClosed="r"/></cx:layoutPr>'), requires: 'cx1', fallback: png() },
      // Region map is always unsupported; without a fallback it is a card with its reason.
      { name: 'region map', part: chartEx('regionMap') },
      { name: 'model', choice: 'model3D', fallback: png() },
      { name: 'equation', choice: 'omml', fallback: png() },
    ]
    const options = format === 'xlsx' ? { cells: HIERARCHY_CELLS, definedNames: HIERARCHY_NAMES } : format === 'docx' ? { omml: true, heading: 'Torture' } : {}
    const context = meteredContext('desktop')
    let opened: Opened | undefined
    await expect((async () => { opened = await open(format, await buildChartDocument(format, objects, options), context) })()).resolves.toBeUndefined()
    const { charts, coverage } = opened!
    const nativeNames = [...CLASSIC.map(f => f.name), 'Waterfall', 'Funnel', 'Treemap', 'Sunburst']
    // Chart contents in document order: every classic family and painted ChartEx family native; surface3D and
    // region map are cards with their reasons; the histogram, model and equation are fallback pictures.
    const drawn = charts.filter((c): c is ChartModel => c !== undefined)
    expect(drawn).toHaveLength(nativeNames.length + 2)
    const natives = drawn.filter(c => !c.unsupported), cards = drawn.filter(c => c.unsupported)
    expect(natives).toHaveLength(nativeNames.length)
    expect(cards.map(c => c.unsupported)).toEqual(['3-D surface charts are shown as a placeholder', 'Region map charts are not supported'])
    for (const model of natives) expect(paint(model, context), 'a native chart is drawn without the card').toBe(false)
    for (const model of cards) expect(paint(model, context)).toBe(true)
    // Coverage: one entry per object; natives select `chart`, cards `chart-placeholder`, the rest `raster-fallback`.
    expect(coverage).toHaveLength(objects.length)
    expect(coverage.filter(e => e.selectedRepresentation === 'chart')).toHaveLength(nativeNames.length)
    expect(coverage.filter(e => e.selectedRepresentation === 'chart-placeholder').map(e => e.reason)).toEqual(['unsupported-chart', 'unsupported-chart'])
    const fallbacks = coverage.filter(e => e.selectedRepresentation === 'raster-fallback')
    expect(fallbacks).toHaveLength(3)
    if (format === 'pptx') expect(fallbacks.map(e => [e.feature, e.status])).toEqual([['ChartEx', 'fallback'], ['model3D', 'fallback'], ['OMML', 'fallback']])
    else expect(fallbacks.every(e => e.status === 'fallback')).toBe(true)
    // Counters: what the natives retain after the parse (and after painting, their layouts too) stays far inside
    // the 64 MiB desktop facet; releasing every model and layout refunds the facet to zero.
    expect(context.peak()).toBeLessThan(8 * MiB)
    expect(context.workspace.used).toBeGreaterThan(0)
    releaseAll(charts, context)
    expect(context.workspace.used).toBe(0)
  })
})

describe('torture (b): a huge embedded workbook and a corrupt embedded zip', () => {
  const cached = chartSpace('Cached', `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${gallerySeries(0, 'S', ['a', 'b'], [1, 2])}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>`)
    .replace('</c:chart>', '</c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>')
  const huge = noise(4 * MiB + 1)

  it.each(['docx', 'pptx'] as const)('%s: an embedding over the mobile 4 MiB inflated cap refuses the chart; saved data never opens it; corrupt zips are unavailable', async format => {
    const context = meteredContext('mobile')
    const objects: ChartDocObject[] = [
      { name: 'huge', part: officeChartXml(), workbook: huge },
      { name: 'corrupt', part: officeChartXml(), workbook: noise(4096) },
      { name: 'truncated', part: officeChartXml(), workbook: officeChartWorkbook().subarray(0, 200) },
    ]
    const { charts, coverage, pkg } = await open(format, await buildChartDocument(format, objects), context)
    expect(charts.map(c => c?.unsupported)).toEqual([MEMORY, UNAVAILABLE, UNAVAILABLE])
    expect(coverage.map(e => e.reason)).toEqual(['chart-refused', 'unsupported-chart', 'unsupported-chart'])
    // The inflate is metered: the peak stays near the cap (never the whole forged/huge payload twice), and nothing is retained.
    expect(context.peak()).toBeLessThanOrEqual(4 * MiB + 512 * 1024)
    expect(context.workspace.used).toBe(0)
    const root = format === 'docx' ? 'word' : 'ppt'
    expect(await adapterDiagnostics(pkg, `${root}/charts/chart1.xml`)).toContain('embedded-workbook-limit')
    expect(await adapterDiagnostics(pkg, `${root}/charts/chart2.xml`)).toContain('chart-data-corrupt')
    expect(await adapterDiagnostics(pkg, `${root}/charts/chart3.xml`)).toContain('chart-data-corrupt')
    for (const chart of charts) paint(chart!, context)
    releaseAll(charts, context)
    expect(context.workspace.used).toBe(0)

    // A chart with saved data and the same huge embedding is native from its caches; the embedding is never read.
    const lean = meteredContext('mobile')
    const savedOnly = await open(format, await buildChartDocument(format, [{ name: 'cached', part: cached, workbook: huge }]), lean)
    expect(savedOnly.charts[0]?.unsupported).toBeUndefined()
    expect(lean.peak()).toBeLessThan(MiB)
    releaseAll(savedOnly.charts, lean)
    expect(lean.workspace.used).toBe(0)
  }, 30_000)

  it('xlsx (no embedding): an oversized chart part and a malformed one are cards; a whole-column reference is display-truncated within bounds', async () => {
    const context = meteredContext('mobile')
    const pad = ' '.repeat(4 * MiB + 16)
    const objects: ChartDocObject[] = [
      { name: 'huge part', part: cached.replace('<c:chart>', `<c:chart>${pad}`) },
      { name: 'malformed', part: `<c:chartSpace ${CONFORMANCE_NS}><` },
    ]
    const { charts, coverage } = await open('xlsx', await buildChartDocument('xlsx', objects), context)
    expect(charts.map(c => c?.unsupported)).toEqual(['This chart is too large to display', 'Chart data could not be read'])
    expect(coverage.map(e => e.reason)).toEqual(['chart-part-limit', 'invalid-content-xml'])
    expect(context.workspace.used).toBe(0)

    const wide = meteredContext('mobile')
    const column = await open('xlsx', await buildChartDocument('xlsx', [{ name: 'whole column', part: officeChartXml('rId1', 'Sheet1!$B$1:$B$1048576') }], { cells: [{ r: 2, cells: [{ ref: 'B2', v: 5 }] }] }), wide)
    const model = column.charts[0]!
    expect(model.unsupported).toBeUndefined()
    expect(model.truncated).toBe(true)
    expect(model.groups[0].series[0].values).toHaveLength(4096)
    expect(model.groups[0].series[0].values.values[1]).toBe(5)
    expect(chartDescription(model)).toMatch(/data truncated\.$/)
    // Display-capped: the retained storage is the 4,096-point prefix, not a million rows.
    expect(wide.peak()).toBeLessThan(MiB)
    expect(paint(model, wide)).toBe(false)
    releaseAll(column.charts, wide)
    expect(wide.workspace.used).toBe(0)
  }, 30_000)
})

describe('torture (c): external links and names pointing at another workbook', () => {
  const cached = chartSpace('Cached', `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${gallerySeries(0, 'S', ['a', 'b'], [1, 2])}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>`)
    .replace('</c:chart>', '</c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>')
  const linked = embeddedWorkbook({ sheets: [{ name: 'Sheet1', data: row(2, cell('B2', 1)) }], definedNames: [{ name: 'Ext', formula: '[1]Sheet1!$B$2:$B$3' }] })

  it.each(['docx', 'pptx'] as const)('%s: an external externalData is never fetched; saved data still draws; external names and references are unavailable', async format => {
    const context = meteredContext('mobile')
    const objects: ChartDocObject[] = [
      { name: 'external', part: officeChartXml(), externalWorkbook: 'file:///C:/Users/me/Book1.xlsx' },
      { name: 'external, saved data', part: cached, externalWorkbook: 'https://example.com/book.xlsx' },
      { name: 'external name', part: officeChartXml('rId1', 'Ext'), workbook: linked },
      { name: 'external reference', part: officeChartXml('rId1', '[1]Sheet1!$B$2:$B$3'), workbook: linked },
    ]
    const { charts, coverage, pkg } = await open(format, await buildChartDocument(format, objects), context)
    expect(charts.map(c => c?.unsupported)).toEqual([UNAVAILABLE, undefined, UNAVAILABLE, UNAVAILABLE])
    expect(coverage.map(e => e.selectedRepresentation)).toEqual(['chart-placeholder', 'chart', 'chart-placeholder', 'chart-placeholder'])
    const values = charts[1]!.groups[0].series[0].values
    expect([...values.values]).toEqual([1, 2])
    const root = format === 'docx' ? 'word' : 'ppt'
    expect(await adapterDiagnostics(pkg, `${root}/charts/chart1.xml`)).toContain('chart-data-external')
    expect(await adapterDiagnostics(pkg, `${root}/charts/chart3.xml`)).toContain('chart-data-external-name')
    expect(await adapterDiagnostics(pkg, `${root}/charts/chart4.xml`)).toContain('chart-data-external-reference')
    // Only the two local embeddings are ever read (their admitted, uncached reads); nothing scales with a link.
    expect(context.peak()).toBeLessThan(2 * MiB)
    releaseAll(charts, context)
    expect(context.workspace.used).toBe(0)
  })

  it('xlsx: a reference into another workbook, direct or through a defined name, reads blank and is diagnosed', async () => {
    const context = meteredContext('mobile')
    const objects: ChartDocObject[] = [
      { name: 'external reference', part: officeChartXml('rId1', '[1]Sheet1!$B$2:$B$3') },
      { name: 'external name', part: officeChartXml('rId1', 'Ext') },
    ]
    const { charts } = await open('xlsx', await buildChartDocument('xlsx', objects, { cells: [{ r: 2, cells: [{ ref: 'B2', v: 5 }] }], definedNames: [{ name: 'Ext', formula: '[1]Sheet1!$B$2:$B$3' }] }), context)
    for (const chart of charts) {
      expect(chart?.unsupported).toBeUndefined()
      // Never the local Sheet1!B2 (5): the other workbook's cells are blank.
      expect([...chart!.groups[0].series[0].values.states]).toEqual([0, 0])
      expect(chart!.diagnostics?.map(d => d.code)).toContain('chart-data-external-reference')
      expect(paint(chart!, context)).toBe(false)
    }
    releaseAll(charts, context)
    expect(context.workspace.used).toBe(0)
  })
})

describe('torture (d): empty, NaN and infinite data, per format', () => {
  const families = CLASSIC.filter(f => ['column', 'line', 'pie', 'scatter', 'bubble', 'stock', 'contour'].includes(f.name))
  const cases = Object.keys(CONFORMANCE_CASES) as Array<keyof typeof CONFORMANCE_CASES>

  it.each(FORMATS)('%s: every family and data case parses, paints natively or as a card with a reason, and refunds', async format => {
    const objects = families.flatMap(family => cases.map(name => ({ name: `${family.name}: ${name}`, part: family.chart(CONFORMANCE_CASES[name]) })))
    const context = meteredContext('mobile')
    const { charts } = await open(format, await buildChartDocument(format, objects, { columns: 6, rows: 7, width: 120, height: 90 }), context)
    expect(charts).toHaveLength(objects.length)
    charts.forEach((chart, i) => {
      expect(chart, objects[i].name).toBeDefined()
      const card = paint(chart!, context, 120, 90)
      if (chart!.unsupported) expect(chart!.unsupported, objects[i].name).toMatch(/\S/)
      else expect(card, `${objects[i].name}: native without the card`).toBe(false)
      for (const d of chart!.diagnostics ?? []) expect(d.code).toMatch(/^chart-[a-z0-9-]+$/)
    })
    // NaN and ±∞ (and 1e999) read as blank, never as a plotted value.
    for (const [i, chart] of charts.entries()) {
      if (!objects[i].name.endsWith('NaN and infinity') || chart!.unsupported) continue
      for (const group of chart!.groups) for (const s of group.series) expect([...s.values.states].every(state => state !== 1), objects[i].name).toBe(true)
    }
    expect(context.peak()).toBeLessThan(4 * MiB)
    releaseAll(charts, context)
    expect(context.workspace.used).toBe(0)
  })
})

// ---- Layout churn ------------------------------------------------------------------------------------------------

describe('layout retention churn (spec 6.5 stage 2)', () => {
  it('repeated paints and cache churn release layout charges', async () => {
    const context = meteredContext('mobile')
    const { charts } = await open('xlsx', await buildChartDocument('xlsx', [{ name: 'column', part: CLASSIC_FAMILIES[0].chart(conformanceCache(4, ['3', '1', '2', '5'])) }]), context)
    const model = charts[0]!
    const modelBytes = context.workspace.used
    expect(modelBytes).toBeGreaterThan(0)
    const sizes = Array.from({ length: 12 }, (_, i) => [200 + i * 37, 140 + i * 23] as const)
    // First cycle: each new placement builds and retains a layout.
    for (const [w, h] of sizes) expect(paint(model, context, w, h)).toBe(false)
    const warm = context.workspace.used
    expect(warm).toBeGreaterThan(modelBytes)
    // Revisiting the same placements reuses them: no new charge, however often.
    for (let round = 0; round < 5; round++) for (const [w, h] of sizes) paint(model, context, w, h)
    expect(context.workspace.used).toBe(warm)
    // Releasing the model's layouts returns exactly their charge; the model keeps its own.
    releaseChartLayouts(model)
    expect(context.workspace.used).toBe(modelBytes)
    // Churn: 40 short-lived models painted and released in turn never accumulate.
    for (let i = 0; i < 40; i++) {
      const { charts: more } = await open('xlsx', await buildChartDocument('xlsx', [{ name: 'tmp', part: CLASSIC_FAMILIES[2].chart(conformanceCache(3, [String(i), '1', '2'])) }]), context)
      paint(more[0]!, context, 300 + i, 200)
      releaseChartLayouts(more[0]!)
      releaseChartModel(more[0]!)
      expect(context.workspace.used).toBe(modelBytes)
    }
    expect(context.peak()).toBeLessThan(2 * MiB)
    releaseChartModel(model)
    expect(context.workspace.used).toBe(0)
  }, 30_000)
})

// ---- Work bounds -------------------------------------------------------------------------------------------------

/** Counts path operations on a painting stub (no rasterisation, so 262,144 segments stay cheap). */
function countingContext() {
  const counts = { moveTo: 0, lineTo: 0, bezierCurveTo: 0, arc: 0 }
  const target: Record<string, unknown> = { font: '', measureText: (t: string) => ({ width: t.length * 6, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 }), getTransform: () => ({ a: 1, d: 1 }) }
  const ctx = new Proxy(target, {
    get: (t, k: string) => k in t ? t[k] : (..._args: unknown[]) => { if (k in counts) counts[k as keyof typeof counts]++; return undefined },
    set: (t, k: string, v) => { t[k] = v; return true },
  })
  return { ctx: ctx as unknown as CanvasRenderingContext2D, counts }
}
const measure = { width: (s: string) => s.length * 6, lineHeight: () => 14 }

describe('work bounds drop elements with diagnostics (spec 6.5, 7)', () => {
  /** A chart of `count` series × 4,096 points of `family`, built in memory from one small parsed series. */
  function wide(family: 'line' | 'scatter', count: number, options: { markers?: boolean; smooth?: boolean } = {}): ChartModel {
    const plot = family === 'line'
      ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:marker><c:symbol val="none"/></c:marker><c:val>${conformanceCache(2, ['1', '2'])}</c:val></c:ser><c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`
      : `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:marker><c:symbol val="none"/></c:marker><c:xVal>${conformanceCache(2, ['1', '2'])}</c:xVal><c:yVal>${conformanceCache(2, ['1', '2'])}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`
    const model = parseXlsxChart(parseXmlOrdered(`<c:chartSpace ${CONFORMANCE_NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`), undefined)!
    const y = new PointWriter(4096), x = new PointWriter(4096)
    for (let i = 0; i < 4096; i++) { y.value(i, i % 97); x.value(i, i) }
    const values = y.done(), xValues = x.done(), template = model.groups[0].series[0]
    model.groups[0].series = Array.from({ length: count }, (_, k) => ({ ...template, name: `S${k}`, values, ...(family === 'scatter' ? { xValues } : {}), markers: options.markers ?? false, ...(options.markers ? { markerStyle: { symbol: 'circle', size: 5 } } : {}), smooth: options.smooth ?? false }) as typeof template)
    if (family === 'line') model.categories = []
    return prepareChartAxes({ ...model, preparedAxes: undefined })
  }
  /** Paints at 400 x 260 and returns the drawn series count (by their one-per-point path calls) and the diagnostics. */
  function drawnSeries(model: ChartModel, per: 'lineTo' | 'bezierCurveTo' | 'arc') {
    const context = documentChartContext('desktop')
    const { ctx, counts } = countingContext()
    paintXlsxChart(model, ctx, 400, 260, undefined, context)
    const layout = layoutChart(model, 400, 260, 1, 0, measure, context)
    const codes = layout.diagnostics.map(d => d.code)
    const segments = counts.moveTo + counts.lineTo + counts.bezierCurveTo + counts.arc
    releaseChartLayouts(model)
    return { series: Math.floor(counts[per] / 4095), counts, segments, codes }
  }

  it('path: admission cost per point is 1 (straight, no markers), 2 (markers) and 3 (smoothing); later series are dropped with chart-path-limit', () => {
    // The bound is 262,144 segments. The axes are admitted first (a 4,096-slot axis prices 16 + 4 × (ticks + 1,000)),
    // leaving R for the series; a series costs 4,096 × perPoint + 1. The counts below are mutually consistent for one
    // R in [249,917, 253,983): straight floor(R / 4,097) = 61, markers floor(R / 8,193) = 30, smoothing
    // floor(R / 12,289) = 20. A series that does not fit is skipped and chart-path-limit is diagnosed.
    const straight = drawnSeries(wide('line', 40), 'lineTo')
    expect(straight.series).toBe(40)
    expect(straight.codes).not.toContain('chart-path-limit')
    const crowded = drawnSeries(wide('line', 64), 'lineTo')
    expect(crowded.series).toBe(61)
    expect(crowded.codes).toContain('chart-path-limit')
    // A4c N3: smoothing is admitted at three segments per point (each Bézier stands for three).
    const smooth = drawnSeries(wide('line', 40, { smooth: true }), 'bezierCurveTo')
    expect(smooth.series).toBe(20)
    expect(smooth.counts.bezierCurveTo).toBe(20 * 4095)
    expect(smooth.codes).toContain('chart-path-limit')
    // A4c N2: a scatter series without markers is admitted at one segment per point (its line only, as for line
    // charts; before A4c every scatter series paid two). With markers it pays two.
    const plain = drawnSeries(wide('scatter', 40), 'lineTo')
    expect(plain.series).toBe(40)
    expect(plain.codes).not.toContain('chart-path-limit')
    expect(drawnSeries(wide('scatter', 64), 'lineTo').series).toBe(61)
    const marked = drawnSeries(wide('scatter', 40, { markers: true }), 'lineTo')
    expect(marked.series).toBe(30)
    expect(marked.counts.arc).toBe(30 * 4096)
    expect(marked.codes).toContain('chart-path-limit')
    for (const run of [straight, crowded, smooth, plain, marked]) expect(run.segments).toBeLessThanOrEqual(262144)
  }, 30_000)

  it('label bytes: ChartEx labels share one 4 MiB budget; long labels are cut to 255 units; both are diagnosed', async () => {
    // A treemap (analytic: no 4,096-point display cap) with 8,000 cached labels of 300 units: each is cut to 255
    // (542 B charged), and the chart-wide 4 MiB budget keeps 7,739 of them.
    const n = 8000, label = 'y'.repeat(300)
    const levels = `<cx:strDim type="cat"><cx:lvl ptCount="${n}">${Array.from({ length: n }, (_, i) => `<cx:pt idx="${i}">${label}</cx:pt>`).join('')}</cx:lvl></cx:strDim><cx:numDim type="size"><cx:lvl ptCount="${n}"><cx:pt idx="0">1</cx:pt></cx:lvl></cx:numDim>`
    const part = `<cx:chartSpace xmlns:cx="${CX}"><cx:chartData><cx:data id="0">${levels}</cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="treemap"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`
    const bytes = await buildChartDocument('pptx', [{ name: 'labels', part }])
    const pkg = await OfficePackage.load(bytes)
    const context = chartParseContext(meteredContext('desktop'), pkg, 'ppt/charts/chartEx1.xml')
    const read = await readChartPart(pkg, 'ppt/charts/chartEx1.xml', context)
    if (read?.kind !== 'tree') throw new Error('read')
    const model = parseChartEx(read.root, undefined, context)
    read.release()
    expect(model.truncated).toBe(true)
    expect(model.diagnostics?.map(d => d.code)).toEqual(expect.arrayContaining(['chart-label-truncated', 'chart-label-budget']))
    const group = model.groups[0] as ChartExGroup
    const level = group.chartEx.data[0].dimensions[0].levels[0]
    if (level.kind !== 'str') throw new Error('string level')
    const kept = level.values.filter((v): v is string => v !== undefined)
    expect(kept.every(v => v.length === 255)).toBe(true)
    expect(kept.length).toBe(Math.floor(4 * MiB / (255 * 2 + 32)))
    expect(context.workspace.used).toBeLessThan(6 * MiB)
    releaseChartModel(model)
    expect(context.workspace.used).toBe(0)
  }, 30_000)

  it('contour grid: more than 64 series rows or 256 category columns are clipped, diagnosed and never drawn', () => {
    const rows = (count: number, columns: number) => Array.from({ length: count }, (_, k) => `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:val>${conformanceCache(columns, Array.from({ length: columns }, (_, j) => String((j + k) % 10)))}</c:val></c:ser>`).join('')
    const contour = (count: number, columns: number) => parseXlsxChart(parseXmlOrdered(`<c:chartSpace ${CONFORMANCE_NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:surfaceChart><c:wireframe val="0"/>${rows(count, columns)}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:surfaceChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx><c:serAx><c:axId val="3"/><c:crossAx val="2"/></c:serAx></c:plotArea></c:chart></c:chartSpace>`), undefined)!
    const fills = (model: ChartModel) => { const { ctx, counts } = countingContext(); let rects = 0; (ctx as unknown as Record<string, unknown>).fillRect = () => { rects++ }; paintXlsxChart(model, ctx, 600, 400, undefined, documentChartContext('desktop')); return { rects, counts } }
    const fits = contour(4, 256)
    expect(fits.diagnostics?.map(d => d.code) ?? []).not.toContain('chart-contour-grid-limit')
    const tall = contour(70, 8), widened = contour(2, 300)
    for (const model of [tall, widened]) {
      expect(model.truncated).toBe(true)
      expect(model.diagnostics?.map(d => d.code)).toContain('chart-contour-grid-limit')
    }
    // The clipped grids draw exactly what the largest admitted grids draw: 70 rows paint as 64, 300 columns as 256.
    const atRows = contour(64, 8), atColumns = contour(2, 256)
    expect(fills(tall).rects).toBe(fills(atRows).rects)
    expect(fills(widened).rects).toBe(fills(atColumns).rects)
    for (const model of [atRows, atColumns]) releaseChartModel(model)
    for (const model of [fits, tall, widened]) releaseChartModel(model)
  })
})
