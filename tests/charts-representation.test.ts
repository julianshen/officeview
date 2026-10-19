/**
 * Task A0b: the representation contract (spec 6.3), bounded chart parts,
 * painter containment, `ensureChartEngine` and `preloadCharts` (spec 6.5),
 * and card coverage (spec 6.6). Fixtures are Office-shaped synthetic.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { prepareCompatibleDrawingContent, type DrawingContent } from '../src/drawing/content'
import { drawingPartContext } from '../src/drawing/parts'
import { parseThemeContext } from '../src/drawing/style'
import { documentChartContext } from '../src/charts/context'
import { parseXlsxChart } from '../src/charts/parse'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import type { DrawingCoverageEntry } from '../src/drawing/coverage'

// Every XML tree the code builds is observable (the bounded chart-part test).
vi.mock('../src/core/xml', async importOriginal => {
  const original = await importOriginal<typeof import('../src/core/xml')>()
  return { ...original, parseXmlOrdered: vi.fn(original.parseXmlOrdered) }
})

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const CX1 = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex'
const RED = 'FF0000', BLUE = '0000FF'
const EMU = 9525
const W_PX = 200, H_PX = 160

const lit = (values: number[]) => `<c:numLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`
const ser = (i: number, color: string, values: number[]) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>S${i + 1}</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr><c:cat><c:strLit><c:ptCount val="${values.length}"/>${values.map((_, k) => `<c:pt idx="${k}"><c:v>C${k}</c:v></c:pt>`).join('')}</c:strLit></c:cat><c:val>${lit(values)}</c:val></c:ser>`
const axes = '<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>'
const space = (chart: string, extra = '') => `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}">${extra}<c:chart>${chart}</c:chart></c:chartSpace>`
/** A stacked column chart: blue segments above red ones. */
const STACKED = space(`<c:autoTitleDeleted val="1"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>${ser(0, RED, [2, 2])}${ser(1, BLUE, [2, 2])}<c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}</c:plotArea>`)
/** surface3DChart stays unsupported in the engine (spec 4.1). */
const SURFACE = space(`<c:title><c:tx><c:rich><a:p><a:r><a:t>Surf</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:surface3DChart>${ser(0, RED, [1, 2])}</c:surface3DChart>${axes}</c:plotArea>`)
const SURFACE_REASON = 'Surface chart is not supported'
const png = () => { const c = createCanvas(2, 2); const g = c.getContext('2d'); g.fillStyle = '#00ff00'; g.fillRect(0, 0, 2, 2); return new Uint8Array(c.toBuffer('image/png')) }

async function packageOf(parts: Record<string, string | Uint8Array>): Promise<OfficePackage> {
  const zip = new JSZip()
  for (const [path, value] of Object.entries(parts)) zip.file(path, value)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
const rels = (...items: Array<[string, string, string]>) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`

// ---- One MC/standalone body per format ---------------------------------------------------------
type Format = 'docx' | 'pptx' | 'xlsx'
const FORMATS: Format[] = ['docx', 'pptx', 'xlsx']
const chartPayload = (rid = 'rC') => `<c:chart r:id="${rid}"/>`
const chartExPayload = '<cx:chart r:id="rX"/>'
/** One placed graphic object in the format's own carrier. */
function graphicObject(format: Format, id: number, payload: string, uri = C): string {
  const graphic = `<a:graphic><a:graphicData uri="${uri}">${payload}</a:graphicData></a:graphic>`
  if (format === 'docx') return `<w:drawing><wp:inline><wp:extent cx="${W_PX * EMU}" cy="${H_PX * EMU}"/><wp:docPr id="${id}" name="Object ${id}"/>${graphic}</wp:inline></w:drawing>`
  const xfrm = `<a:off x="0" y="0"/><a:ext cx="${W_PX * EMU}" cy="${H_PX * EMU}"/>`
  if (format === 'pptx') return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Object ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm>${xfrm}</p:xfrm>${graphic}</p:graphicFrame>`
  return `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Object ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm>${xfrm}</xdr:xfrm>${graphic}</xdr:graphicFrame>`
}
/** A fallback picture in the format's own carrier. */
function picture(format: Format, id: number, rid = 'im'): string {
  if (format === 'docx') return graphicObject('docx', id, `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill><pic:spPr/></pic:pic>`, PIC)
  const xfrm = `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${W_PX * EMU}" cy="${H_PX * EMU}"/></a:xfrm>`
  if (format === 'pptx') return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/></p:blipFill><p:spPr>${xfrm}<a:prstGeom prst="rect"/></p:spPr></p:pic>`
  return `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"/></xdr:blipFill><xdr:spPr>${xfrm}<a:prstGeom prst="rect"/></xdr:spPr></xdr:pic>`
}
const alternate = (choice: string, fallback: string | undefined, requires = 'c', prefix = 'mc') =>
  `<${prefix}:AlternateContent${prefix === 'mc' ? '' : ` xmlns:${prefix}="${MC}"`}><${prefix}:Choice Requires="${requires}">${choice}</${prefix}:Choice>${fallback === undefined ? '' : `<${prefix}:Fallback>${fallback}</${prefix}:Fallback>`}</${prefix}:AlternateContent>`
const NAMESPACES = `xmlns:mc="${MC}" xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}" xmlns:cx="${CX}" xmlns:cx1="${CX1}" xmlns:pic="${PIC}" xmlns:ch="${C}"`

interface Built { pkg: OfficePackage; chartPath: string }
/** A one-object document whose body is `body`; `chart` (and `chartEx`) are its chart parts. */
async function build(format: Format, body: string, options: { chart?: string; chartEx?: boolean; image?: Uint8Array | null } = {}): Promise<Built> {
  const image = options.image === undefined ? png() : options.image
  const chart = options.chart ?? SURFACE
  const chartEx = `<cx:chartSpace xmlns:cx="${CX}"><cx:chart><cx:plotArea/></cx:chart></cx:chartSpace>`
  const relItems = (prefix: string): Array<[string, string, string]> => [
    ['rC', 'chart', `${prefix}charts/chart1.xml`], ['rX', 'chartEx', `${prefix}charts/chartEx1.xml`], ...(image ? [['im', 'image', `${prefix}media/image1.png`] as [string, string, string]] : []),
  ]
  if (format === 'docx') {
    const pkg = await packageOf({
      'word/document.xml': `<w:document xmlns:w="${W}" xmlns:wp="${WP}" ${NAMESPACES}><w:body><w:p><w:r>${body}</w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(...relItems('')),
      'word/charts/chart1.xml': chart,
      ...(options.chartEx ? { 'word/charts/chartEx1.xml': chartEx } : {}),
      ...(image ? { 'word/media/image1.png': image } : {}),
    })
    return { pkg, chartPath: 'word/charts/chart1.xml' }
  }
  if (format === 'pptx') {
    const pkg = await packageOf({
      'ppt/presentation.xml': `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldSz cx="${300 * EMU}" cy="${200 * EMU}"/><p:sldIdLst><p:sldId id="256" r:id="s1"/></p:sldIdLst></p:presentation>`,
      'ppt/_rels/presentation.xml.rels': rels(['s1', 'slide', 'slides/slide1.xml']),
      'ppt/slides/slide1.xml': `<p:sld xmlns:p="${P}" ${NAMESPACES}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': rels(...relItems('../')),
      'ppt/charts/chart1.xml': chart,
      ...(options.chartEx ? { 'ppt/charts/chartEx1.xml': chartEx } : {}),
      ...(image ? { 'ppt/media/image1.png': image } : {}),
    })
    return { pkg, chartPath: 'ppt/charts/chart1.xml' }
  }
  const pkg = await packageOf({
    'xl/workbook.xml': `<workbook xmlns:r="${R}"><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels(['s1', 'worksheet', 'worksheets/sheet1.xml']),
    'xl/worksheets/sheet1.xml': `<worksheet xmlns:r="${R}"><sheetData/><drawing r:id="dr"/></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels(['dr', 'drawing', '../drawings/drawing1.xml']),
    'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="${XDR}" ${NAMESPACES}><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="${W_PX * EMU}" cy="${H_PX * EMU}"/>${body}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
    'xl/drawings/_rels/drawing1.xml.rels': rels(...relItems('../')),
    'xl/charts/chart1.xml': chart,
    ...(options.chartEx ? { 'xl/charts/chartEx1.xml': chartEx } : {}),
    ...(image ? { 'xl/media/image1.png': image } : {}),
  })
  return { pkg, chartPath: 'xl/charts/chart1.xml' }
}

interface Outcome { contents: Array<DrawingContent<unknown, unknown>>; pictures: number; coverage: DrawingCoverageEntry[]; diagnostics: Array<{ reason?: string }> }
/** Parse and collect what the object resolved to: its drawn contents, pictures and coverage. */
async function resolve(format: Format, built: Built): Promise<Outcome> {
  if (format === 'docx') {
    const doc = await parseDocx(built.pkg)
    const images = doc.sections[0].paragraphs[0]?.images ?? []
    return { contents: images.flatMap(i => i.drawing ? [i.drawing] : []), pictures: images.filter(i => !i.drawing && i.data.length).length, coverage: (doc.drawingCoverage ?? []).filter(e => e.scope === 'original'), diagnostics: drawingPartContext(built.pkg).diagnostics }
  }
  if (format === 'pptx') {
    const doc = await parsePptx(built.pkg)
    const shapes = doc.slides[0].shapes
    return { contents: shapes.flatMap(s => s.content ? [s.content] : []), pictures: shapes.filter(s => s.imageIndex !== undefined).length, coverage: (doc.drawingCoverage ?? []).filter(e => e.scope === 'original'), diagnostics: drawingPartContext(built.pkg).diagnostics }
  }
  const doc = await parseXlsx(built.pkg)
  const drawings = doc.sheets[0].drawings ?? []
  return { contents: drawings.flatMap(d => d.content ? [d.content] : []), pictures: drawings.filter(d => d.image).length, coverage: (doc.drawingCoverage ?? []).filter(e => e.scope === 'original'), diagnostics: drawingPartContext(built.pkg).diagnostics }
}
function expectFallbackPicture(format: Format, outcome: Outcome, id: string, label: string): void {
  expect(outcome.contents, `${format} ${label}: no chart or card`).toEqual([])
  expect(outcome.pictures, `${format} ${label}: the fallback picture`).toBe(1)
  expect(outcome.coverage, `${format} ${label}`).toHaveLength(1)
  expect(outcome.coverage[0], `${format} ${label}`).toMatchObject({ id, status: 'fallback', representation: 'fallback', selectedRepresentation: 'raster-fallback' })
}
function expectCard(format: Format, outcome: Outcome, reason: string, feature = 'chart'): Extract<DrawingContent<unknown, unknown>, { kind: 'xlsxChart' }> {
  expect(outcome.contents, format).toHaveLength(1)
  const card = outcome.contents[0]
  if (card.kind !== 'xlsxChart') throw new Error(`${format}: not a chart card`)
  expect(card.chart.unsupported, format).toBe(reason)
  expect(card.chart.groups, format).toEqual([])
  expect(outcome.pictures, format).toBe(0)
  // Exactly one selected representation, carrying the original feature.
  expect(outcome.coverage, format).toHaveLength(1)
  expect(outcome.coverage[0], format).toMatchObject({ feature, selectedRepresentation: 'chart-placeholder' })
  return card
}

/** Bounding box and count of the pixels close to one colour. */
function colourBox(data: Uint8ClampedArray, w: number, h: number, hex: string, tolerance = 24) {
  const [r, g, b] = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16))
  let count = 0, x0 = w, y0 = h, x1 = -1, y1 = -1
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const o = (y * w + x) * 4
    if (Math.abs(data[o] - r) < tolerance && Math.abs(data[o + 1] - g) < tolerance && Math.abs(data[o + 2] - b) < tolerance) { count++; x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y) }
  }
  return { count, x0, y0, x1, y1 }
}
/** Native stacked geometry (blue above red, same columns), never the #F7F7F7 card. */
function expectNativeStacked(data: Uint8ClampedArray, w: number, h: number, label: string): void {
  const red = colourBox(data, w, h, RED), blue = colourBox(data, w, h, BLUE), card = colourBox(data, w, h, 'F7F7F7', 2)
  expect(card.count, `${label}: no card`).toBeLessThan(500)
  expect(red.count, label).toBeGreaterThan(200)
  expect(blue.count, label).toBeGreaterThan(200)
  expect(blue.y1, `${label}: blue stacked above red`).toBeLessThan(red.y0 + 2)
  expect(Math.abs(blue.x0 - red.x0), label).toBeLessThanOrEqual(1)
}
function expectCardPixels(data: Uint8ClampedArray, w: number, h: number, label: string): void {
  expect(colourBox(data, w, h, 'F7F7F7', 2).count, `${label}: card`).toBeGreaterThan(W_PX * H_PX * 0.6)
  expect(colourBox(data, w, h, RED).count, `${label}: no native bars`).toBe(0)
}

afterEach(() => { vi.mocked(parseXmlOrdered).mockClear(); vi.doUnmock('../src/charts') })

describe('representation contract (spec 6.3)', () => {
  it('unsupported and refused results do not satisfy an mc:Choice; the fallback picture is selected', async () => {
    // surface3DChart stays unsupported in the engine: end to end, in all three formats.
    for (const format of FORMATS) {
      const body = alternate(graphicObject(format, 10, chartPayload()), picture(format, 11))
      expectFallbackPicture(format, await resolve(format, await build(format, body)), '11', 'surface3DChart')
    }
    // Injected adapter outcomes, through the shared resolver, on each format's own carrier.
    const theme = parseThemeContext()
    const contentTheme = { colors: new Map<string, string>(), fonts: new Map<string, string>() }
    const native = parseXlsxChart(parseXmlOrdered(STACKED), theme)!
    for (const format of FORMATS) {
      const inner = alternate(chartPayload(), picture(format, 21))
      const wrapper = format === 'docx' ? `<wp:inline xmlns:wp="${WP}" ${NAMESPACES}><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData uri="${C}">${inner}</a:graphicData></a:graphic></wp:inline>`
        : format === 'pptx' ? `<a:graphicData uri="${C}" ${NAMESPACES}>${inner}</a:graphicData>`
        : `<xdr:graphicFrame xmlns:xdr="${XDR}" ${NAMESPACES}><a:graphic><a:graphicData uri="${C}">${inner}</a:graphicData></a:graphic></xdr:graphicFrame>`
      const owner = format === 'docx' ? 'word/document.xml' : format === 'pptx' ? 'ppt/slides/slide1.xml' : 'xl/drawings/drawing1.xml'
      for (const outcome of [{ kind: 'unsupported', reason: 'injected unsupported' }, { kind: 'refused', reason: 'injected refusal' }, { kind: 'native', model: native }] as const) {
        const { pkg } = await build(format, '')
        const node = parseXmlOrdered(wrapper)
        const parseChart = vi.fn(async () => outcome)
        const prepared = await prepareCompatibleDrawingContent(pkg, node, owner, contentTheme, theme, { parseChart })
        expect(parseChart, `${format} ${outcome.kind}: the Choice was attempted`).toHaveBeenCalledTimes(1)
        if (outcome.kind === 'native') {
          expect(prepared?.content, format).toEqual({ kind: 'xlsxChart', chart: native })
          expect(prepared?.selections.at(-1)?.representation, format).toBe('choice')
        } else {
          expect(prepared?.content, `${format} ${outcome.kind}`).toBeUndefined()
          expect(prepared?.selections.at(-1)?.representation, `${format} ${outcome.kind}`).toBe('fallback')
        }
      }
    }
  })

  it('chartex choice with fallback keeps the fallback picture', async () => {
    for (const format of FORMATS) {
      const body = alternate(graphicObject(format, 30, chartExPayload, CX), picture(format, 31), 'cx1')
      expectFallbackPicture(format, await resolve(format, await build(format, body, { chartEx: true })), '31', 'ChartEx')
    }
  })

  it('standalone unsupported chart becomes a card with the reason, the original feature and one representation', async () => {
    for (const format of FORMATS) {
      const card = expectCard(format, await resolve(format, await build(format, graphicObject(format, 40, chartPayload()))), SURFACE_REASON)
      expect(card.chart.title, format).toBe('Surf')
    }
    // A standalone ChartEx frame keeps its original feature.
    const ex = await resolve('pptx', await build('pptx', graphicObject('pptx', 41, chartExPayload, CX), { chartEx: true }))
    expect(expectCard('pptx', ex, 'ChartEx charts are not supported yet', 'ChartEx').chart.unsupported).toBe('ChartEx charts are not supported yet')
  })

  it('a chart Choice with no fallback becomes its card (standalone and no-fallback)', async () => {
    for (const format of FORMATS) {
      const outcome = await resolve(format, await build(format, alternate(graphicObject(format, 45, chartPayload()), undefined)))
      expect(expectCard(format, outcome, SURFACE_REASON).chart.title, format).toBe('Surf')
      expect(outcome.coverage[0], format).toMatchObject({ id: '45', representation: 'choice' })
    }
  })

  it('a broken fallback picture becomes a card recording chart-fallback-unusable', async () => {
    for (const format of FORMATS) {
      const body = alternate(graphicObject(format, 50, chartPayload()), picture(format, 51, 'im'))
      const outcome = await resolve(format, await build(format, body, { image: null }))
      expectCard(format, outcome, SURFACE_REASON)
      expect(outcome.coverage[0], format).toMatchObject({ id: '50', representation: 'choice' })
      // Recorded beside the original reason.
      expect(outcome.diagnostics.some(d => d.reason === 'chart-fallback-unusable'), format).toBe(true)
      // A usable fallback picture is still selected over the card.
      expectFallbackPicture(format, await resolve(format, await build(format, body)), '51', 'usable fallback')
    }
    const drawing = (await parseXlsx((await build('xlsx', alternate(graphicObject('xlsx', 50, chartPayload()), picture('xlsx', 51, 'im')), { image: null })).pkg)).sheets[0].drawings![0]
    expect(drawing.source.reason).toBe(`chart-fallback-unusable: ${SURFACE_REASON}`)
  })

  it('nested and alias-prefixed AlternateContent and unknown Requires', async () => {
    for (const format of FORMATS) {
      // Nested: the inner unsupported chart Choice falls to the inner fallback picture.
      const nested = alternate(alternate(graphicObject(format, 60, chartPayload()), picture(format, 61)), picture(format, 62))
      // PPTX coverage keeps the attempted outer object's identity for a nested selection (existing convention).
      expectFallbackPicture(format, await resolve(format, await build(format, nested)), format === 'pptx' ? '60' : '61', 'nested')
      // An alias-prefixed wrapper and an aliased Requires prefix behave as mc: and c:.
      const aliased = alternate(graphicObject(format, 63, chartPayload()), picture(format, 64), 'ch', 'mc2')
      expectFallbackPicture(format, await resolve(format, await build(format, aliased)), '64', 'alias')
      // An unknown Requires is never attempted, even when its chart would draw natively.
      const unknown = alternate(graphicObject(format, 65, chartPayload()), picture(format, 66), 'future')
      expectFallbackPicture(format, await resolve(format, await build(format, unknown, { chart: STACKED })), '66', 'unknown Requires')
    }
  })

  it('a card reports coverage chart-placeholder', async () => {
    for (const format of FORMATS) {
      const outcome = await resolve(format, await build(format, graphicObject(format, 70, chartPayload())))
      expect(outcome.coverage[0], format).toMatchObject({ id: '70', feature: 'chart', selectedRepresentation: 'chart-placeholder' })
      // The card paints as the card.
      const ctx = createCanvas(W_PX, H_PX).getContext('2d')
      const { paintDrawingContent } = await import('../src/drawing/content-paint')
      paintDrawingContent(outcome.contents[0] as DrawingContent, ctx as unknown as CanvasRenderingContext2D, W_PX, H_PX)
      expectCardPixels(ctx.getImageData(0, 0, W_PX, H_PX).data, W_PX, H_PX, format)
    }
  })
})

describe('bounded chart parts (spec 6.5)', () => {
  /** Rewrites one entry's declared uncompressed size (local header and central directory). */
  function forgeDeclaredSize(zip: Uint8Array, name: string, size: number): Uint8Array {
    const out = zip.slice(), view = new DataView(out.buffer), bytes = new TextEncoder().encode(name)
    const matches = (at: number) => bytes.every((b, i) => out[at + i] === b)
    for (let i = 0; i + 46 < out.length; i++) {
      const sig = view.getUint32(i, true)
      if (sig === 0x04034b50 && view.getUint16(i + 26, true) === bytes.length && matches(i + 30)) view.setUint32(i + 22, size, true)
      if (sig === 0x02014b50 && view.getUint16(i + 28, true) === bytes.length && matches(i + 46)) view.setUint32(i + 24, size, true)
    }
    return out
  }

  it('a chart part over the byte cap is refused before any XML tree is built', async () => {
    const cap = 4 * 1024 * 1024 // mobile
    const marker = 'oversized-chart-marker'
    const huge = space(`<c:plotArea><c:barChart><c:barDir val="col"/>${ser(0, RED, [1])}</c:barChart></c:plotArea>`, `<!-- ${marker} ${'x'.repeat(cap)} -->`)
    for (const forged of [false, true]) {
      const zip = new JSZip()
      zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:wp="${WP}" ${NAMESPACES}><w:body><w:p><w:r>${graphicObject('docx', 80, chartPayload())}</w:r></w:p></w:body></w:document>`)
      zip.file('word/_rels/document.xml.rels', rels(['rC', 'chart', 'charts/chart1.xml']))
      zip.file('word/charts/chart1.xml', huge)
      let bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
      if (forged) bytes = forgeDeclaredSize(bytes, 'word/charts/chart1.xml', 100)
      const pkg = await OfficePackage.load(bytes)
      // Observe every inflated chunk the bounded read is handed.
      const accepted: number[] = [], refused: number[] = [], late: number[] = []
      const entry = (pkg as unknown as { zip: JSZip }).zip.file('word/charts/chart1.xml') as unknown as { internalStream: (type: string) => { on: (event: string, fn: (chunk: Uint8Array) => void) => unknown } }
      const internal = entry.internalStream.bind(entry)
      entry.internalStream = type => {
        const stream = internal(type), on = stream.on.bind(stream)
        stream.on = (event, fn) => on(event, event !== 'data' ? fn : (chunk: Uint8Array) => {
          if (refused.length) { late.push(chunk.length); return }
          try { fn(chunk); accepted.push(chunk.length) } catch (error) { refused.push(chunk.length); throw error }
        })
        return stream
      }
      const xmlOrdered = vi.spyOn(pkg, 'xmlOrdered')
      vi.mocked(parseXmlOrdered).mockClear()

      const doc = await parseDocx(pkg, { chartContext: documentChartContext('mobile') })
      await new Promise(resolve => setTimeout(resolve, 20))

      // No tree of the chart part was ever built.
      expect(xmlOrdered.mock.calls.some(([path]) => path === 'word/charts/chart1.xml'), `forged=${forged}`).toBe(false)
      expect(vi.mocked(parseXmlOrdered).mock.calls.some(([xml]) => xml.includes(marker)), `forged=${forged}`).toBe(false)
      // Accepted bytes stay within the cap; the refusal is the first crossing chunk, and nothing follows it.
      expect(accepted.reduce((a, b) => a + b, 0), `forged=${forged}`).toBeLessThanOrEqual(cap)
      expect(refused, `forged=${forged}`).toHaveLength(1)
      expect(accepted.reduce((a, b) => a + b, 0) + refused[0], `forged=${forged}`).toBeGreaterThan(cap)
      expect(late, `forged=${forged}`).toEqual([])
      // The chart is unsupported: a standalone frame becomes a card.
      const image = doc.sections[0].paragraphs[0].images[0]
      expect(image?.drawing?.kind, `forged=${forged}`).toBe('xlsxChart')
      if (image?.drawing?.kind === 'xlsxChart') expect(image.drawing.chart.unsupported).toBe('This chart is too large to display')
    }
  })
})

describe('locale-aware automatic title', () => {
  it('an empty title element takes the c:lang default; the single-series name and authored titles are preserved', () => {
    const plot = (n: number) => `<c:plotArea><c:barChart><c:barDir val="col"/>${Array.from({ length: n }, (_, i) => ser(i, RED, [1])).join('')}</c:barChart>${axes}</c:plotArea>`
    const title = (lang: string | undefined, chart: string) => parseXlsxChart(parseXmlOrdered(space(chart, lang ? `<c:lang val="${lang}"/>` : '')), undefined)?.title
    expect(title('zh-TW', `<c:title/>${plot(2)}`)).toBe('圖表標題')
    expect(title('zh-CN', `<c:title/><c:autoTitleDeleted val="0"/>${plot(2)}`)).toBe('圖表標題')
    expect(title('en-US', `<c:title/>${plot(2)}`)).toBe('Chart Title')
    expect(title(undefined, `<c:title/>${plot(2)}`)).toBe('Chart Title')
    // The single-series rule, autoTitleDeleted and an authored title win.
    expect(title('zh-TW', `<c:title/>${plot(1)}`)).toBe('S1')
    expect(title('zh-TW', `<c:title/><c:autoTitleDeleted val="1"/>${plot(2)}`)).toBeUndefined()
    expect(title('zh-TW', `<c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>${plot(2)}`)).toBe('Sales')
    expect(title('zh-TW', plot(2))).toBeUndefined()
  })
})

describe('painter containment and the engine across realms (spec 6.5)', () => {
  /** A worker-shaped document: parsed in this realm, then structured-cloned (transferred). */
  async function transferred(format: 'docx' | 'pptx') {
    const { pkg } = await build(format, graphicObject(format, 90, chartPayload()), { chart: STACKED })
    const doc = format === 'docx' ? await parseDocx(pkg) : await parsePptx(pkg)
    return structuredClone(doc)
  }
  /** A fresh module realm: nothing has loaded the chart engine (or registered its painter). */
  async function freshRealm() {
    vi.resetModules()
    return {
      paint: await import('../src/render/paint'),
      contentPaint: await import('../src/drawing/content-paint'),
      render: await import('../src/pptx/render'),
    }
  }
  async function paintFirst(pages: Array<{ spec: { widthPx: number; heightPx: number }; paint: (ctx: CanvasRenderingContext2D) => void }>) {
    const { widthPx: w, heightPx: h } = pages[0].spec
    const ctx = createCanvas(w, h).getContext('2d')
    pages[0].paint(ctx as unknown as CanvasRenderingContext2D)
    return { data: ctx.getImageData(0, 0, w, h).data, w, h }
  }

  it('worker-shaped (transferred) DOCX and PPTX charts paint natively after getPaintables', async () => {
    for (const format of ['docx', 'pptx'] as const) {
      const doc = await transferred(format)
      const realm = await freshRealm()
      const pages = await realm.paint.getPaintables(doc)
      const { data, w, h } = await paintFirst(pages)
      expectNativeStacked(data, w, h, format)
      pages.dispose()
    }
  })

  it('getPaintables returns pages with cards when the chart chunk fails', async () => {
    for (const format of ['docx', 'pptx'] as const) {
      const doc = await transferred(format)
      vi.resetModules()
      let attempts = 0
      vi.doMock('../src/charts', () => { attempts++; throw new Error('chunk failed to load') })
      const { getPaintables } = await import('../src/render/paint')
      const pages = await getPaintables(doc)
      // The engine load was attempted and its failure caught: the pages still return.
      expect(attempts, format).toBeGreaterThan(0)
      expect(pages.length, format).toBe(1)
      const { data, w, h } = await paintFirst(pages)
      expectCardPixels(data, w, h, format)
      pages.dispose()
      vi.doUnmock('../src/charts')
    }
  })

  it('renderSlide paints a card when the engine was never loaded in this realm', async () => {
    const doc = await transferred('pptx') as Awaited<ReturnType<typeof parsePptx>>
    const realm = await freshRealm()
    const metrics = realm.render.slideMetrics(doc)
    const ctx = createCanvas(metrics.widthPx, metrics.heightPx).getContext('2d')
    realm.render.renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, metrics, [])
    expectCardPixels(ctx.getImageData(0, 0, metrics.widthPx, metrics.heightPx).data, metrics.widthPx, metrics.heightPx, 'renderSlide')
  })

  it('preloadCharts registers the painter', async () => {
    const doc = await transferred('pptx') as Awaited<ReturnType<typeof parsePptx>>
    const content = doc.slides[0].shapes[0].content as DrawingContent
    const realm = await freshRealm()
    const before = createCanvas(W_PX, H_PX).getContext('2d')
    realm.contentPaint.paintDrawingContent(content, before as unknown as CanvasRenderingContext2D, W_PX, H_PX)
    expectCardPixels(before.getImageData(0, 0, W_PX, H_PX).data, W_PX, H_PX, 'before preloadCharts')
    const { preloadCharts } = await import('../src/index')
    await expect(preloadCharts()).resolves.toBeUndefined()
    const after = createCanvas(W_PX, H_PX).getContext('2d')
    realm.contentPaint.paintDrawingContent(content, after as unknown as CanvasRenderingContext2D, W_PX, H_PX)
    expectNativeStacked(after.getImageData(0, 0, W_PX, H_PX).data, W_PX, H_PX, 'after preloadCharts')
  })

  it('a throwing painter paints a card and restores canvas state for later shapes', async () => {
    const realm = await freshRealm()
    const doc = await transferred('pptx') as Awaited<ReturnType<typeof parsePptx>>
    const content = doc.slides[0].shapes[0].content as DrawingContent
    realm.contentPaint.registerXlsxChartPainter((_chart, ctx) => {
      ctx.translate(37, 41); ctx.fillStyle = '#ff0000'; ctx.lineWidth = 9; ctx.globalAlpha = 0.5
      ctx.beginPath(); ctx.rect(0, 0, 10, 10); ctx.clip()
      throw new Error('engine bug')
    })
    const ctx = createCanvas(W_PX, H_PX).getContext('2d')
    ctx.fillStyle = '#123456'
    expect(() => realm.contentPaint.paintDrawingContent(content, ctx as unknown as CanvasRenderingContext2D, W_PX, H_PX)).not.toThrow()
    expectCardPixels(ctx.getImageData(0, 0, W_PX, H_PX).data, W_PX, H_PX, 'card on throw')
    // State is back: transform, style, alpha, width and clip do not leak into a later shape.
    const m = ctx.getTransform()
    expect([m.a, m.b, m.c, m.d, m.e, m.f]).toEqual([1, 0, 0, 1, 0, 0])
    expect([ctx.fillStyle, ctx.lineWidth, ctx.globalAlpha]).toEqual(['#123456', 1, 1])
    ctx.fillRect(150, 120, 20, 20)
    expect(Array.from(ctx.getImageData(160, 130, 1, 1).data)).toEqual([0x12, 0x34, 0x56, 255])
  })
})
