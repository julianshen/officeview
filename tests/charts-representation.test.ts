/**
 * Task A0b: the representation contract (spec 6.3), bounded chart parts,
 * painter containment, `ensureChartEngine` and `preloadCharts` (spec 6.5),
 * and card coverage (spec 6.6). Fixtures are Office-shaped synthetic.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { namespaceScope, namespaceUri, parseXmlOrdered } from '../src/core/xml'
import { chartCard, prepareCompatibleDrawingContent, prepareDrawingContent, type DrawingContent } from '../src/drawing/content'
import { drawingPartContext } from '../src/drawing/parts'
import { parseThemeContext } from '../src/drawing/style'
import { documentChartContext } from '../src/charts/context'
import { parseXlsxChart } from '../src/charts/parse'
import { CHART_READ_CROSSING, ELEMENT_BYTES, NAMESPACE_BINDING_BYTES, NAMESPACE_SCOPE_BYTES, NamespaceScope, readChartPart } from '../src/charts/read'
import { paintDrawingContent } from '../src/drawing/content-paint'
import { getPaintables } from '../src/render/paint'
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
const UNRELATED = { docx: '<w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>', pptx: `<p:sp><p:nvSpPr><p:cNvPr id="900" name="Unrelated"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Unrelated</a:t></a:r></a:p></p:txBody></p:sp>`, xlsx: '<row r="1"><c r="A1" t="inlineStr"><is><t>Unrelated</t></is></c></row>' }
/** `chart: null` omits the chart part (a missing reference); `unrelated` adds content that must survive. */
async function build(format: Format, body: string, options: { chart?: string | Uint8Array | null; chartEx?: boolean; image?: Uint8Array | null; unrelated?: boolean } = {}): Promise<Built> {
  const image = options.image === undefined ? png() : options.image
  const chart = options.chart === undefined ? SURFACE : options.chart
  const extra = (f: Format) => options.unrelated ? UNRELATED[f] : ''
  const chartEx = `<cx:chartSpace xmlns:cx="${CX}"><cx:chart><cx:plotArea/></cx:chart></cx:chartSpace>`
  const relItems = (prefix: string): Array<[string, string, string]> => [
    ['rC', 'chart', `${prefix}charts/chart1.xml`], ['rX', 'chartEx', `${prefix}charts/chartEx1.xml`], ...(image ? [['im', 'image', `${prefix}media/image1.png`] as [string, string, string]] : []),
  ]
  if (format === 'docx') {
    const pkg = await packageOf({
      'word/document.xml': `<w:document xmlns:w="${W}" xmlns:wp="${WP}" ${NAMESPACES}><w:body><w:p><w:r>${body}</w:r></w:p>${extra('docx')}</w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(...relItems('')),
      ...(chart !== null ? { 'word/charts/chart1.xml': chart } : {}),
      ...(options.chartEx ? { 'word/charts/chartEx1.xml': chartEx } : {}),
      ...(image ? { 'word/media/image1.png': image } : {}),
    })
    return { pkg, chartPath: 'word/charts/chart1.xml' }
  }
  if (format === 'pptx') {
    const pkg = await packageOf({
      'ppt/presentation.xml': `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldSz cx="${300 * EMU}" cy="${200 * EMU}"/><p:sldIdLst><p:sldId id="256" r:id="s1"/></p:sldIdLst></p:presentation>`,
      'ppt/_rels/presentation.xml.rels': rels(['s1', 'slide', 'slides/slide1.xml']),
      'ppt/slides/slide1.xml': `<p:sld xmlns:p="${P}" ${NAMESPACES}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${body}${extra('pptx')}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': rels(...relItems('../')),
      ...(chart !== null ? { 'ppt/charts/chart1.xml': chart } : {}),
      ...(options.chartEx ? { 'ppt/charts/chartEx1.xml': chartEx } : {}),
      ...(image ? { 'ppt/media/image1.png': image } : {}),
    })
    return { pkg, chartPath: 'ppt/charts/chart1.xml' }
  }
  const pkg = await packageOf({
    'xl/workbook.xml': `<workbook xmlns:r="${R}"><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels(['s1', 'worksheet', 'worksheets/sheet1.xml']),
    'xl/worksheets/sheet1.xml': `<worksheet xmlns:r="${R}"><sheetData>${extra('xlsx')}</sheetData><drawing r:id="dr"/></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels(['dr', 'drawing', '../drawings/drawing1.xml']),
    'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="${XDR}" ${NAMESPACES}><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="${W_PX * EMU}" cy="${H_PX * EMU}"/>${body}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
    'xl/drawings/_rels/drawing1.xml.rels': rels(...relItems('../')),
    ...(chart !== null ? { 'xl/charts/chart1.xml': chart } : {}),
    ...(options.chartEx ? { 'xl/charts/chartEx1.xml': chartEx } : {}),
    ...(image ? { 'xl/media/image1.png': image } : {}),
  })
  return { pkg, chartPath: 'xl/charts/chart1.xml' }
}

interface Outcome { contents: Array<DrawingContent<unknown, unknown>>; pictures: number; coverage: DrawingCoverageEntry[]; diagnostics: Array<{ reason?: string }> }
/** Parse and collect what the object resolved to: its drawn contents, pictures and coverage. */
type Parsers = { parseDocx: typeof parseDocx; parsePptx: typeof parsePptx; parseXlsx: typeof parseXlsx }
async function resolve(format: Format, built: Built, parsers: Parsers = { parseDocx, parsePptx, parseXlsx }): Promise<Outcome> {
  if (format === 'docx') {
    const doc = await parsers.parseDocx(built.pkg)
    const images = doc.sections[0].paragraphs[0]?.images ?? []
    return { contents: images.flatMap(i => i.drawing ? [i.drawing] : []), pictures: images.filter(i => !i.drawing && i.data.length).length, coverage: (doc.drawingCoverage ?? []).filter(e => e.scope === 'original'), diagnostics: drawingPartContext(built.pkg).diagnostics }
  }
  if (format === 'pptx') {
    const doc = await parsers.parsePptx(built.pkg)
    const shapes = doc.slides[0].shapes
    return { contents: shapes.flatMap(s => s.content ? [s.content] : []), pictures: shapes.filter(s => s.imageIndex !== undefined).length, coverage: (doc.drawingCoverage ?? []).filter(e => e.scope === 'original'), diagnostics: drawingPartContext(built.pkg).diagnostics }
  }
  const doc = await parsers.parseXlsx(built.pkg)
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
    const huge = space(`<c:plotArea><c:barChart><c:barDir val="col"/>${ser(0, RED, [1])}</c:barChart></c:plotArea>`, `<!-- ${marker} -->${`<!--${'x'.repeat(1000)}-->`.repeat(Math.ceil(cap / 1000) + 8)}`)
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

// ---- Fix round 1 (codex review I1-I7) ------------------------------------------------------------
const GREEN_RECT = (format: 'pptx' | 'xlsx', id: number) => {
  const p = format === 'pptx' ? 'p' : 'xdr'
  return `<${p}:sp><${p}:nvSpPr><${p}:cNvPr id="${id}" name="Rect ${id}"/><${p}:cNvSpPr/>${p === 'p' ? '<p:nvPr/>' : ''}</${p}:nvSpPr><${p}:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W_PX * EMU}" cy="${H_PX * EMU}"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></${p}:spPr></${p}:sp>`
}
const group = (format: 'pptx' | 'xlsx', id: number, children: string) => {
  const p = format === 'pptx' ? 'p' : 'xdr'
  const box = `<a:off x="0" y="0"/><a:ext cx="${W_PX * EMU}" cy="${H_PX * EMU}"/>`
  return `<${p}:grpSp><${p}:nvGrpSpPr><${p}:cNvPr id="${id}" name="Group ${id}"/><${p}:cNvGrpSpPr/>${p === 'p' ? '<p:nvPr/>' : ''}</${p}:nvGrpSpPr><${p}:grpSpPr><a:xfrm>${box}<a:chOff x="0" y="0"/><a:chExt cx="${W_PX * EMU}" cy="${H_PX * EMU}"/></a:xfrm></${p}:grpSpPr>${children}</${p}:grpSp>`
}
type Node = { source?: { id?: string }; children?: Node[] }
const ids = (nodes: Node[]): string[] => nodes.flatMap(node => [...(node.source?.id ? [node.source.id] : []), ...ids(node.children ?? [])])
async function sceneIds(format: 'pptx' | 'xlsx', built: Built): Promise<string[]> {
  if (format === 'pptx') return ids((await parsePptx(built.pkg)).slides[0].shapes as Node[])
  return ids(((await parseXlsx(built.pkg)).sheets[0].drawings ?? []) as Node[])
}
const CORRUPT = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4, 5, 6, 7, 8])

describe('fix round 1: chunk failure, decode failure, whole fallbacks, propagation (spec 6.3)', () => {
  it('I1: a chart chunk that fails at parse time is a chart failure; the fallback, the card and unrelated content survive', async () => {
    for (const format of FORMATS) {
      vi.resetModules()
      let attempts = 0
      vi.doMock('../src/charts', () => { attempts++; throw new Error('chunk failed to load') })
      const parsers: Parsers = { parseDocx: (await import('../src/docx/parse')).parseDocx, parsePptx: (await import('../src/pptx/parse')).parsePptx, parseXlsx: (await import('../src/xlsx/parse')).parseXlsx }
      // A valid fallback picture is selected.
      const mc = alternate(graphicObject(format, 10, chartPayload()), picture(format, 11))
      expectFallbackPicture(format, await resolve(format, await build(format, mc), parsers), '11', 'chunk failure')
      // With no fallback, the card says why; unrelated content survives beside it.
      const built = await build(format, graphicObject(format, 12, chartPayload()), { unrelated: true })
      const outcome = await resolve(format, built, parsers)
      expect(attempts, format).toBeGreaterThan(0)
      const card = outcome.contents.find(content => content.kind === 'xlsxChart')
      expect(card?.kind === 'xlsxChart' && card.chart.unsupported, format).toBe('The chart engine could not be loaded')
      expect(outcome.coverage.find(entry => entry.id === '12'), format).toMatchObject({ selectedRepresentation: 'chart-placeholder' })
      if (format === 'docx') expect((await parsers.parseDocx((await build('docx', graphicObject('docx', 12, chartPayload()), { unrelated: true })).pkg)).sections[0].paragraphs[1].runs.map(r => r.text).join('')).toBe('Unrelated')
      if (format === 'pptx') expect(outcome.coverage.some(entry => entry.id === '900' && entry.status === 'native')).toBe(true)
      if (format === 'xlsx') {
        const doc = await parsers.parseXlsx((await build('xlsx', graphicObject('xlsx', 12, chartPayload()), { unrelated: true })).pkg)
        expect(doc.sheets[0].rows.flatMap(row => row.cells).map(cell => cell.value)).toContain('Unrelated')
      }
      vi.doUnmock('../src/charts')
    }
    vi.resetModules()
  })

  it('I2: a selected fallback picture that fails to decode becomes the card and records chart-fallback-unusable, after transfer', async () => {
    for (const format of FORMATS) {
      const body = alternate(graphicObject(format, 20, chartPayload()), picture(format, 21))
      for (const corrupt of [true, false]) {
        const doc = structuredClone((format === 'docx' ? await parseDocx((await build(format, body, { image: corrupt ? CORRUPT : undefined })).pkg)
          : format === 'pptx' ? await parsePptx((await build(format, body, { image: corrupt ? CORRUPT : undefined })).pkg)
          : await parseXlsx((await build(format, body, { image: corrupt ? CORRUPT : undefined })).pkg)) as never) as Parameters<typeof getPaintables>[0]
        const pages = await getPaintables(doc)
        const { widthPx: w, heightPx: h } = pages[0].spec
        const ctx = createCanvas(w, h).getContext('2d')
        pages[0].paint(ctx as unknown as CanvasRenderingContext2D)
        const data = ctx.getImageData(0, 0, w, h).data
        const originals = ((doc as { drawingCoverage?: DrawingCoverageEntry[] }).drawingCoverage ?? []).filter(entry => entry.scope === 'original')
        expect(originals, `${format} corrupt=${corrupt}`).toHaveLength(1)
        if (corrupt) {
          expectCardPixels(data, w, h, `${format} decode failure`)
          expect(originals[0], format).toMatchObject({ selectedRepresentation: 'chart-placeholder', reason: `chart-fallback-unusable: ${SURFACE_REASON}` })
        } else {
          expect(colourBox(data, w, h, '00FF00').count, `${format} the picture paints`).toBeGreaterThan(W_PX * H_PX * 0.5)
          expect(originals[0], format).toMatchObject({ selectedRepresentation: 'raster-fallback' })
        }
        pages.dispose()
      }
    }
  })

  it('I3: one unusable picture never discards usable shapes or groups in the same fallback', async () => {
    for (const format of ['pptx', 'xlsx'] as const) {
      const broken = picture(format, 31, 'gone')
      // Mixed: a broken picture beside a valid rectangle keeps the fallback (the rectangle draws).
      const mixed = await build(format, alternate(graphicObject(format, 30, chartPayload()), broken + GREEN_RECT(format, 99)))
      expect(await sceneIds(format, mixed), `${format} mixed`).toContain('99')
      expect((await resolve(format, await build(format, alternate(graphicObject(format, 30, chartPayload()), broken + GREEN_RECT(format, 99))))).contents, `${format} mixed: no card`).toEqual([])
      // A group of only broken pictures is unusable: the card.
      const onlyBroken = await resolve(format, await build(format, alternate(graphicObject(format, 30, chartPayload()), group(format, 98, broken))))
      expectCard(format, onlyBroken, SURFACE_REASON)
      // A group with a usable shape keeps the fallback.
      const grouped = await build(format, alternate(graphicObject(format, 30, chartPayload()), group(format, 98, broken + GREEN_RECT(format, 97))))
      expect(await sceneIds(format, grouped), `${format} grouped`).toContain('97')
    }
    // DOCX: a broken picture beside a valid one keeps the fallback picture.
    const docx = await resolve('docx', await build('docx', alternate(graphicObject('docx', 30, chartPayload()), picture('docx', 31, 'gone') + picture('docx', 32))))
    expect(docx.contents).toEqual([])
    expect(docx.pictures).toBe(1)
  })

  it('I4: a terminal chart failure propagates through nested wrappers and missing references', async () => {
    for (const format of FORMATS) {
      // Outer supported Choice, inner supported Choice, unsupported chart; no fallback anywhere.
      const nested = alternate(alternate(graphicObject(format, 40, chartPayload()), undefined), undefined)
      expectCard(format, await resolve(format, await build(format, nested)), SURFACE_REASON)
      // A standalone missing chart part.
      expectCard(format, await resolve(format, await build(format, graphicObject(format, 41, chartPayload()), { chart: null })), 'Chart data could not be read')
      // A missing chart in a Choice with a valid fallback still selects the fallback.
      expectFallbackPicture(format, await resolve(format, await build(format, alternate(graphicObject(format, 42, chartPayload()), picture(format, 43)), { chart: null })), '43', 'missing chart')
    }
  })
})

describe('fix round 1: the real engine balances every save (spec 6.5, I5)', () => {
  const titled = space(`<c:autoTitleDeleted val="1"/><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>${ser(0, RED, [3, 5])}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Category</a:t></a:r></a:p></c:rich></c:tx></c:title><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Value</a:t></a:r></a:p></c:rich></c:tx></c:title><c:crossAx val="1"/></c:valAx></c:plotArea>`)
  it.each(['rotated axis title', 'plot area'] as const)('a throw inside the %s save leaves the caller\'s depth and transform', async where => {
    await import('../src/charts') // the real engine registers its painter
    const chart = parseXlsxChart(parseXmlOrdered(titled), undefined)!
    expect(chart.categoryAxis.title).toBe('Category')
    const canvas = createCanvas(W_PX, H_PX), real = canvas.getContext('2d')
    let depth = 0, thrown = false
    const ctx = new Proxy(real, {
      get(target, prop) {
        if (prop === 'save') return () => { depth++; target.save() }
        if (prop === 'restore') return () => { depth--; target.restore() }
        if (prop === 'fillText' && where === 'rotated axis title') return (...args: [string, number, number]) => {
          if (!thrown && target.getTransform().b !== 0) { thrown = true; throw new Error('engine bug in the rotated title') }
          target.fillText(...args)
        }
        if (prop === 'fillRect' && where === 'plot area') return (...args: [number, number, number, number]) => {
          if (!thrown && String(target.fillStyle).toLowerCase() === '#ff0000') { thrown = true; throw new Error('engine bug in the plot') }
          target.fillRect(...args)
        }
        const value = Reflect.get(target, prop)
        return typeof value === 'function' ? value.bind(target) : value
      },
      set(target, prop, value) { return Reflect.set(target, prop, value) },
    })
    ctx.save(); ctx.translate(17, 19)
    paintDrawingContent({ kind: 'xlsxChart', chart }, ctx as unknown as CanvasRenderingContext2D, W_PX, H_PX)
    expect(thrown, where).toBe(true)
    expect(depth, where).toBe(1)
    const m = real.getTransform()
    expect([m.a, m.b, m.c, m.d, m.e, m.f], where).toEqual([1, 0, 0, 1, 17, 19])
    ctx.restore()
    expect(depth).toBe(0)
    const after = real.getTransform()
    expect([after.e, after.f]).toEqual([0, 0])
  })
})

describe('fix round 1: admitted bounded read and namespace dispatch (spec 6.5, I6, I7)', () => {
  const graphicData = () => parseXmlOrdered(`<a:graphicData xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}" uri="${C}"><c:chart r:id="rC"/></a:graphicData>`)
  const contentTheme = { colors: new Map<string, string>(), fonts: new Map<string, string>() }
  async function attempt(part: string | Uint8Array, chartContext = documentChartContext('desktop')) {
    const pkg = await packageOf({ 'word/_rels/document.xml.rels': rels(['rC', 'chart', 'charts/chart1.xml']), 'word/charts/chart1.xml': part })
    const parseChart = vi.fn(async (root: Parameters<typeof parseXlsxChart>[0]) => ({ kind: 'native' as const, model: parseXlsxChart(root, undefined)! }))
    const node = graphicData()
    const content = await prepareDrawingContent(pkg, node, 'word/document.xml', contentTheme, parseThemeContext(), { parseChart, chartContext })
    return { pkg, content, parseChart, card: chartCard(node)?.chart }
  }
  const utf16 = (text: string, bigEndian: boolean) => {
    const out = new Uint8Array(2 + text.length * 2)
    out[0] = bigEndian ? 0xfe : 0xff; out[1] = bigEndian ? 0xff : 0xfe
    for (let i = 0; i < text.length; i++) { const code = text.charCodeAt(i); out[2 + 2 * i + (bigEndian ? 1 : 0)] = code & 255; out[2 + 2 * i + (bigEndian ? 0 : 1)] = code >> 8 }
    return out
  }
  const chartEx = (open: string) => `${open}<cx:chart xmlns:cx="${CX}"><cx:plotArea/></cx:chart></${open.slice(1, open.indexOf(' '))}>`

  it('I7: dispatch uses the real root namespace (default, alias, UTF-8, UTF-16, long preambles)', async () => {
    const preamble = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<!--${'p'.repeat(20000)}-->\n<?office-pi data?>${' '.repeat(5000)}`
    const classicAlias = STACKED.replace(/c:/g, 'k:').replace('xmlns:k=', 'xmlns:k=')
    for (const [label, part, expected] of [
      ['default ChartEx namespace', chartEx(`<chartSpace xmlns="${CX}">`), 'ChartEx'],
      ['alias ChartEx prefix', chartEx(`<q:chartSpace xmlns:q="${CX}">`), 'ChartEx'],
      ['long preamble ChartEx', preamble + chartEx(`<q:chartSpace xmlns:q="${CX}">`), 'ChartEx'],
      ['UTF-16BE ChartEx', utf16(chartEx(`<q:chartSpace xmlns:q="${CX}">`), true), 'ChartEx'],
      ['other namespace', `<chartSpace xmlns="urn:example:not-a-chart"><chart><plotArea><barChart>${ser(0, RED, [1])}</barChart></plotArea></chart></chartSpace>`.replace(/c:/g, ''), 'other'],
      ['classic alias prefix', classicAlias, 'classic'],
      ['long preamble classic', preamble + STACKED, 'classic'],
      ['UTF-8 BOM classic', new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(STACKED)]), 'classic'],
      ['UTF-16LE classic', utf16(STACKED, false), 'classic'],
      ['UTF-16BE classic', utf16(STACKED, true), 'classic'],
    ] as const) {
      const { content, parseChart, card } = await attempt(part)
      if (expected === 'classic') {
        expect(parseChart, label).toHaveBeenCalledTimes(1)
        expect(content?.kind === 'xlsxChart' && content.chart.groups[0]?.series.length, label).toBe(2)
      } else {
        // A non-classic part never reaches the classic parser.
        expect(parseChart, label).not.toHaveBeenCalled()
        expect(content, label).toBeUndefined()
        expect(card, label).toMatchObject(expected === 'ChartEx' ? { unsupported: 'ChartEx charts are not supported yet', feature: 'ChartEx' } : { unsupported: 'This chart format is not supported' })
      }
    }
  })

  it('I6: a refusing parent refuses the read before anything is allocated', async () => {
    let parentCalls = 0
    const context = documentChartContext('mobile', undefined, { reserve: () => { parentCalls++; return false }, release: () => {} })
    const part = space(`<c:plotArea><c:barChart><c:barDir val="col"/>${ser(0, RED, [1, 2])}</c:barChart></c:plotArea>`, `<!-- refused-marker -->${`<!--${'x'.repeat(1000)}-->`.repeat(256)}`)
    const pkg = await packageOf({ 'word/_rels/document.xml.rels': rels(['rC', 'chart', 'charts/chart1.xml']), 'word/charts/chart1.xml': part })
    let streamed = 0
    const entry = (pkg as unknown as { zip: JSZip }).zip.file('word/charts/chart1.xml') as unknown as { internalStream: (type: string) => unknown }
    const internal = entry.internalStream.bind(entry)
    entry.internalStream = type => { streamed++; return internal(type) }
    vi.mocked(parseXmlOrdered).mockClear()
    const node = graphicData()
    const content = await prepareDrawingContent(pkg, node, 'word/document.xml', contentTheme, parseThemeContext(), { chartContext: context })
    expect(content).toBeUndefined()
    expect(chartCard(node)?.chart.unsupported).toBe('This chart needs more memory than is available for charts')
    expect(parentCalls).toBeGreaterThan(0)
    expect(streamed, 'nothing inflated').toBe(0)
    expect(vi.mocked(parseXmlOrdered).mock.calls.some(([xml]) => xml.includes('refused-marker'))).toBe(false)
    expect(context.workspace.used).toBe(0)
    expect(drawingPartContext(pkg).diagnostics.some(d => d.reason === 'chart-refused')).toBe(true)
  })

  it('I6: a paged package keeps owning the chart tree through its own ledger', async () => {
    // The paged RangePackage shape: no zip entry, its own bounded bytes() and a chartLedger (its pool).
    const bytes = new TextEncoder().encode(STACKED)
    const pool = { charged: 0, refuseAfter: Infinity }
    const paged = (refuseAfter = Infinity) => {
      pool.charged = 0; pool.refuseAfter = refuseAfter
      return { has: () => true, bytes: async (_path: string, max?: number) => { if (max !== undefined && bytes.length > max) throw new Error('Part byte limit exceeded'); return bytes },
        chartLedger: { reserve: (n: number) => { if (pool.charged + n > pool.refuseAfter) return false; pool.charged += n; return true }, release: () => {} } } as unknown as OfficePackage
    }
    const context = documentChartContext('desktop')
    const read = await readChartPart(paged(), 'xl/charts/chart1.xml', context)
    expect(read?.kind).toBe('tree')
    // The pool was charged for every tree node; the workspace holds the tree (and the crossing), not the pool's raw bytes.
    expect(pool.charged).toBeGreaterThan(0)
    expect(context.workspace.used).toBe(CHART_READ_CROSSING + pool.charged)
    if (read?.kind === 'tree') read.release()
    expect(context.workspace.used).toBe(0)
    // The pool refusing part way refuses the read and refunds the workspace.
    const refused = await readChartPart(paged(500), 'xl/charts/chart1.xml', context)
    expect(refused).toMatchObject({ kind: 'refused', code: 'chart-refused' })
    expect(context.workspace.used).toBe(0)
  })

  it('I6: every allocation of the read is charged first, peaks within the limit and is released', async () => {
    const context = documentChartContext('desktop')
    const workspace = context.workspace
    let peak = 0, charges = 0
    const reserve = workspace.reserve.bind(workspace)
    ;(workspace as { reserve: typeof reserve }).reserve = (bytes, tag) => { const ok = reserve(bytes, tag); if (ok) { charges++; peak = Math.max(peak, workspace.used) } return ok }
    const padding = `<!--${'x'.repeat(1000)}-->`.repeat(256) // ~256 KiB streamed, never retained
    const before = workspace.used
    const { content, parseChart } = await attempt(STACKED.replace('<c:chart>', `${padding}<c:chart>`), context)
    expect(parseChart).toHaveBeenCalledTimes(1)
    expect(content?.kind).toBe('xlsxChart')
    expect(charges, 'crossing, chunks and tree nodes').toBeGreaterThan(20)
    expect(peak).toBeGreaterThan(before + CHART_READ_CROSSING)
    expect(peak).toBeLessThanOrEqual(workspace.limit)
    // Streamed bytes are refunded chunk by chunk: the peak stays far below the part size plus its text.
    expect(peak - before).toBeLessThan(CHART_READ_CROSSING + 3 * 64 * 1024 + 256 * 1024)
    expect(workspace.used, 'released after the parse').toBe(before)
  })
})

// ---- Fix round 2 (codex re-review R1, R2, comment limit) ---------------------------------------
describe('fix round 2: propagation follows the selection result (R1)', () => {
  it('a failed inner fallback is not a representation: the card propagates; a usable outer fallback still wins', async () => {
    for (const format of FORMATS) {
      // Outer supported Choice → inner unsupported chart Choice + missing-picture Fallback; no outer fallback.
      const failedInner = alternate(alternate(graphicObject(format, 60, chartPayload()), picture(format, 61, 'gone')), undefined)
      expectCard(format, await resolve(format, await build(format, failedInner)), SURFACE_REASON)
      // Control: a usable outer fallback wins over the propagated failure.
      const control = alternate(alternate(graphicObject(format, 60, chartPayload()), picture(format, 61, 'gone')), picture(format, 62))
      expectFallbackPicture(format, await resolve(format, await build(format, control)), '62', 'usable outer fallback')
    }
  })

  it('seven nested supported Choices with no fallbacks give the card', async () => {
    for (const format of FORMATS) {
      let body = graphicObject(format, 70, chartPayload())
      for (let i = 0; i < 7; i++) body = alternate(body, undefined)
      expectCard(format, await resolve(format, await build(format, body)), SURFACE_REASON)
    }
  })
})

describe('fix round 2: namespace ownership is admitted before construction (R2)', () => {
  // 100 root declarations and 200 nested elements, each declaring one more namespace.
  const nested = `<chartSpace ${Array.from({ length: 100 }, (_, i) => `xmlns:n${i}="urn:n${i}"`).join(' ')}>${Array.from({ length: 200 }, (_, i) => `<x xmlns:z${i}="urn:z${i}">`).join('')}${'</x>'.repeat(200)}</chartSpace>`
  const pkgOf = () => packageOf({ 'chart.xml': nested })

  it('every namespace scope and binding is charged, inherited bindings are shared (never copied)', async () => {
    const context = documentChartContext('desktop')
    const read = await readChartPart(await pkgOf(), 'chart.xml', context)
    expect(read?.kind).toBe('tree')
    if (read?.kind !== 'tree') return
    // Deepest node: it sees all 300 bindings, but owns only its own declaration.
    let node = read.root
    for (let i = 0; i < 200; i++) node = (node as Record<string, unknown>).x as typeof node
    expect(namespaceUri(node, 'n0')).toBe('urn:n0')
    expect(namespaceUri(node, 'z199')).toBe('urn:z199')
    expect(namespaceUri(node, 'z0')).toBe('urn:z0')
    const scope = namespaceScope(node) as NamespaceScope
    expect(scope.size).toBe(300)
    expect(scope.ownSize).toBe(1)
    // The charge covers 201 scopes and 300 owned bindings, on top of the elements and the crossing.
    expect(context.workspace.used - CHART_READ_CROSSING).toBeGreaterThanOrEqual(201 * (NAMESPACE_SCOPE_BYTES) + 300 * NAMESPACE_BINDING_BYTES + 201 * ELEMENT_BYTES)
    read.release()
    expect(context.workspace.used).toBe(0)
  })

  it.each([3, 4, 60])('a refusal at reservation %i happens before the refused node or its scope is built', async refuseAt => {
    let calls = 0
    const context = documentChartContext('desktop', undefined, { reserve: () => ++calls < refuseAt, release: () => {} })
    const built = NamespaceScope.built
    const read = await readChartPart(await pkgOf(), 'chart.xml', context)
    expect(read).toMatchObject({ kind: 'refused', code: 'chart-refused' })
    // Reservations: 1 crossing, 2 the chunk, then one per element (each before its node and scope exist).
    expect(NamespaceScope.built - built).toBe(Math.max(0, refuseAt - 3))
    expect(context.workspace.used).toBe(0)
  })
})

describe('fix round 2: token limits are resource limits; comments stream (comment limit)', () => {
  const contentTheme = { colors: new Map<string, string>(), fonts: new Map<string, string>() }
  async function attempt(part: string) {
    const pkg = await packageOf({ 'word/_rels/document.xml.rels': rels(['rC', 'chart', 'charts/chart1.xml']), 'word/charts/chart1.xml': part })
    const node = parseXmlOrdered(`<a:graphicData xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}" uri="${C}"><c:chart r:id="rC"/></a:graphicData>`)
    const content = await prepareDrawingContent(pkg, node, 'word/document.xml', contentTheme, parseThemeContext(), {})
    return { content, card: chartCard(node)?.chart, diagnostics: drawingPartContext(pkg).diagnostics }
  }
  it('a valid 66,000-character comment (across chunks, also in the prolog) parses natively', async () => {
    for (const part of [STACKED.replace('<c:chart>', `<!--${'c'.repeat(66000)}--><c:chart>`), `<?xml version="1.0"?><!--${'- c'.repeat(30000)}-->${STACKED}`]) {
      const { content } = await attempt(part)
      expect(content?.kind === 'xlsxChart' && content.chart.groups[0]?.series.length).toBe(2)
    }
  })
  it('markup over the token limit is a resource refusal with its limit, not malformed XML', async () => {
    const { content, card, diagnostics } = await attempt(STACKED.replace('<c:chart>', `<c:chart big="${'v'.repeat(70000)}">`))
    expect(content).toBeUndefined()
    expect(card?.unsupported).toBe('This chart contains markup larger than the viewer reads')
    expect(diagnostics.find(d => d.reason === 'chart-token-limit')).toMatchObject({ kind: 'unsupported-content', limit: 65536 })
    expect(diagnostics.some(d => d.reason === 'invalid-content-xml')).toBe(false)
  })
})
