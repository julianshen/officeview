/**
 * Task A0a: one lazy chart engine for every format (spec 6.1, 6.2).
 * DOCX and PPTX charts parse through the shared default `parseChart`
 * (the former XLSX adapter) into the `xlsxChart` runtime tag; the eager
 * column-chart mini-reader is gone. Fixtures are Office-shaped synthetic.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { getChildren, parseXmlOrdered } from '../src/core/xml'
import { prepareDrawingContent, type DrawingContent } from '../src/drawing/content'
import { paintDrawingContent } from '../src/drawing/content-paint'
import { parseThemeContext } from '../src/drawing/style'
import { documentChartContext } from '../src/charts/context'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import type { XlsxChartModel } from '../src/xlsx/chartModel'
import { EAGER_FORBIDDEN, forbiddenEager, eagerModules, type ModuleMap } from './helpers/eager-graph'

// The lazy engine is wrapped so the test sees the context each chart parse receives.
vi.mock('../src/charts', async importOriginal => {
  const original = await importOriginal<typeof import('../src/charts')>()
  return { ...original, parseXlsxChart: vi.fn(original.parseXlsxChart) }
})

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const RED = 'FF0000', BLUE = '0000FF'

const lit = (values: number[]) => `<c:numLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`
const cats = (n: number) => `<c:cat><c:strLit><c:ptCount val="${n}"/>${Array.from({ length: n }, (_, i) => `<c:pt idx="${i}"><c:v>C${i + 1}</c:v></c:pt>`).join('')}</c:strLit></c:cat>`
const fill = (color: string) => `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></c:spPr>`
const ser = (i: number, color: string, values: number[]) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>S${i + 1}</c:v></c:tx>${fill(color)}${cats(values.length)}<c:val>${lit(values)}</c:val></c:ser>`
const axes = '<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>'
const space = (plot: string) => `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>${plot}</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`

const CHARTS = {
  stackedColumn: space(`<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>${ser(0, RED, [2, 2])}${ser(1, BLUE, [2, 2])}<c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}`),
  horizontalBar: space(`<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser(0, RED, [4, 2])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}`),
  line: space(`<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser(0, RED, [1, 4, 2])}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes}`),
  pie: space(`<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="${RED}"/></a:solidFill></c:spPr></c:dPt><c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="${BLUE}"/></a:solidFill></c:spPr></c:dPt>${cats(2)}<c:val>${lit([1, 1])}</c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart>`),
} as const
type ChartName = keyof typeof CHARTS

const EXT = { cx: 200 * 9525, cy: 160 * 9525 }
const OFF = { x: 12 * 9525, y: 24 * 9525 }

async function zip(parts: Record<string, string>): Promise<OfficePackage> {
  const archive = new JSZip()
  for (const [path, xml] of Object.entries(parts)) archive.file(path, xml)
  return OfficePackage.load(await archive.generateAsync({ type: 'uint8array' }))
}
const chartRels = (target: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${target}"/></Relationships>`
const graphic = `<a:graphic xmlns:a="${A}"><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rC"/></a:graphicData></a:graphic>`

function docxPackage(name: ChartName): Promise<OfficePackage> {
  return zip({
    'word/document.xml': `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="${R}"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="${EXT.cx}" cy="${EXT.cy}"/><wp:docPr id="7" name="Chart 7"/>${graphic}</wp:inline></w:drawing></w:r></w:p></w:body></w:document>`,
    'word/_rels/document.xml.rels': chartRels('charts/chart1.xml'),
    'word/charts/chart1.xml': CHARTS[name],
  })
}
function pptxPackage(name: ChartName): Promise<OfficePackage> {
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
  return zip({
    'ppt/presentation.xml': `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldSz cx="${300 * 9525}" cy="${200 * 9525}"/><p:sldIdLst><p:sldId id="256" r:id="s1"/></p:sldIdLst></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
    'ppt/slides/slide1.xml': `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Chart 7"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${OFF.x}" y="${OFF.y}"/><a:ext cx="${EXT.cx}" cy="${EXT.cy}"/></p:xfrm>${graphic}</p:graphicFrame></p:spTree></p:cSld></p:sld>`,
    'ppt/slides/_rels/slide1.xml.rels': chartRels('../charts/chart1.xml'),
    'ppt/charts/chart1.xml': CHARTS[name],
  })
}

/** Bounding box and count of the pixels close to one colour. */
function colourBox(ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>, w: number, h: number, hex: string, tolerance = 24) {
  const [r, g, b] = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16))
  const data = ctx.getImageData(0, 0, w, h).data
  let count = 0, x0 = w, y0 = h, x1 = -1, y1 = -1
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const o = (y * w + x) * 4
    if (Math.abs(data[o] - r) < tolerance && Math.abs(data[o + 1] - g) < tolerance && Math.abs(data[o + 2] - b) < tolerance) {
      count++; x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y)
    }
  }
  return { count, x0, y0, x1, y1, w: x1 - x0 + 1, h: y1 - y0 + 1 }
}

/** Paint the content natively and assert the family's distinct geometry (never a card). */
function expectNativeGeometry(content: DrawingContent<unknown, unknown>, name: ChartName): void {
  const w = 200, h = 160
  const ctx = createCanvas(w, h).getContext('2d')
  paintDrawingContent(content as DrawingContent, ctx as unknown as CanvasRenderingContext2D, w, h)
  // A native chart paints a white chart area; the placeholder card would paint #F7F7F7.
  expect(Array.from(ctx.getImageData(2, 2, 1, 1).data).slice(0, 3)).toEqual([255, 255, 255])
  expect(colourBox(ctx, w, h, 'F7F7F7', 2).count).toBeLessThan(w * h * 0.01)
  const red = colourBox(ctx, w, h, RED), blue = colourBox(ctx, w, h, BLUE)
  expect(red.count).toBeGreaterThan(0)
  if (name === 'stackedColumn') {
    // Two stacked segments per category: blue sits on top of red, same columns.
    expect(blue.count).toBeGreaterThan(0)
    expect(blue.y1).toBeLessThan(red.y0 + 2)
    expect(Math.abs(blue.x0 - red.x0)).toBeLessThanOrEqual(1)
    expect(red.h).toBeGreaterThan(red.w / 4)
  } else if (name === 'horizontalBar') {
    // Bars run horizontally: wider than tall, two bars of different lengths.
    expect(red.w).toBeGreaterThan(red.h)
    expect(blue.count).toBe(0)
  } else if (name === 'line') {
    // A stroked polyline: sparse within its box, spanning most of the plot width.
    expect(red.count).toBeLessThan(red.w * red.h * 0.5)
    expect(red.w).toBeGreaterThan(w / 2)
  } else {
    // Two equal half-discs side by side, about as tall as wide.
    expect(blue.count).toBeGreaterThan(0)
    expect(Math.abs(red.count - blue.count) / red.count).toBeLessThan(0.1)
    expect(red.x0).toBeGreaterThanOrEqual(blue.x1 - 2)
    expect(Math.abs(red.h - 2 * red.w) / red.h).toBeLessThan(0.15)
  }
}

function expectModel(chart: XlsxChartModel, name: ChartName): void {
  expect(chart.unsupported).toBeUndefined()
  const group = chart.groups[0]
  expect(chart.groups).toHaveLength(1)
  if (name === 'stackedColumn') expect(group).toMatchObject({ family: 'bar', barDir: 'col', grouping: 'stacked' })
  if (name === 'horizontalBar') expect(group).toMatchObject({ family: 'bar', barDir: 'bar', grouping: 'clustered' })
  if (name === 'line') expect(group).toMatchObject({ family: 'line', grouping: 'standard' })
  if (name === 'pie') expect(group).toMatchObject({ family: 'pie' })
  expect(group.series.map(s => s.color)).toEqual(name === 'stackedColumn' ? [RED, BLUE] : [name === 'pie' ? group.series[0].color : RED])
}

afterEach(() => { vi.clearAllMocks() })

describe('one lazy chart engine (spec 6.1, 6.2)', () => {
  it.each(Object.keys(CHARTS) as ChartName[])('docx and pptx stacked bar, horizontal bar, line and pie parse natively as xlsxChart (%s)', async name => {
    // DOCX
    const docx = await parseDocx(await docxPackage(name))
    const image = docx.sections[0].paragraphs[0].images[0]
    expect(image.drawing?.kind).toBe('xlsxChart')
    if (image.drawing?.kind !== 'xlsxChart') throw new Error('not native')
    expectModel(image.drawing.chart, name)
    expect([image.widthEmu, image.heightEmu]).toEqual([EXT.cx, EXT.cy])
    expect(docx.drawingCoverage?.find(e => e.scope === 'original')).toMatchObject({ feature: 'chart', status: 'native', selectedRepresentation: 'chart' })
    expectNativeGeometry(image.drawing, name)

    // PPTX
    const pptx = await parsePptx(await pptxPackage(name))
    const shape = pptx.slides[0].shapes[0]
    expect(shape.content?.kind).toBe('xlsxChart')
    if (shape.content?.kind !== 'xlsxChart') throw new Error('not native')
    expectModel(shape.content.chart, name)
    expect([shape.xEmu, shape.yEmu, shape.widthEmu, shape.heightEmu]).toEqual([OFF.x, OFF.y, EXT.cx, EXT.cy])
    expect(pptx.drawingCoverage?.find(e => e.scope === 'original')).toMatchObject({ feature: 'chart', status: 'native', selectedRepresentation: 'chart' })
    expectNativeGeometry(shape.content, name)
  })

  it('a caller passing only text adapters still gets the default parseChart; a caller parseChart overrides', async () => {
    const pkg = await zip({ 'word/charts/chart1.xml': CHARTS.horizontalBar, 'word/_rels/document.xml.rels': chartRels('charts/chart1.xml') })
    const node = getChildren(getChildren(parseXmlOrdered(graphic), 'graphic')[0] ?? parseXmlOrdered(graphic), 'graphicData')[0]
    expect(node).toBeDefined()
    const theme = parseThemeContext()
    const contentTheme = { colors: new Map<string, string>(), fonts: new Map<string, string>() }
    const textOnly = { parseParagraph: () => ({ runs: [], align: 'left' as const }), parseDiagramText: () => 'text' }
    const defaulted = await prepareDrawingContent(pkg, node, 'word/document.xml', contentTheme, theme, textOnly)
    expect(defaulted?.kind).toBe('xlsxChart')
    if (defaulted?.kind === 'xlsxChart') expect(defaulted.chart.groups[0]).toMatchObject({ family: 'bar', barDir: 'bar' })

    const override: XlsxChartModel = { groups: [], axes: [], categories: ['override'], fontFamily: 'Calibri', fontSizePt: 10 }
    const parseChart = vi.fn(async () => ({ kind: 'native' as const, model: override }))
    const overridden = await prepareDrawingContent(await zip({ 'word/charts/chart1.xml': CHARTS.horizontalBar, 'word/_rels/document.xml.rels': chartRels('charts/chart1.xml') }), node, 'word/document.xml', contentTheme, theme, { ...textOnly, parseChart })
    expect(parseChart).toHaveBeenCalledWith(expect.anything(), 'word/charts/chart1.xml')
    expect(overridden).toEqual({ kind: 'xlsxChart', chart: override })

    // Anything but native follows today's unsupported-content path.
    const unsupported = await prepareDrawingContent(await zip({ 'word/charts/chart1.xml': CHARTS.horizontalBar, 'word/_rels/document.xml.rels': chartRels('charts/chart1.xml') }), node, 'word/document.xml', contentTheme, theme, { ...textOnly, parseChart: async () => ({ kind: 'unsupported' as const, reason: 'injected' }) })
    expect(unsupported).toBeUndefined()
    const refused = await prepareDrawingContent(await zip({ 'word/charts/chart1.xml': CHARTS.horizontalBar, 'word/_rels/document.xml.rels': chartRels('charts/chart1.xml') }), node, 'word/document.xml', contentTheme, theme, { ...textOnly, parseChart: async () => ({ kind: 'refused' as const, reason: 'injected' }) })
    expect(refused).toBeUndefined()
  })

  it('documentChartContext reaches DOCX and PPTX charts; without one, each parse creates its own', async () => {
    const charts = await import('../src/charts')
    const parse = vi.mocked(charts.parseXlsxChart)
    const doc = documentChartContext('mobile')
    await parseDocx(await docxPackage('line'), { chartContext: doc })
    await parsePptx(await pptxPackage('pie'), { chartContext: doc })
    expect(parse).toHaveBeenCalledTimes(2)
    for (const [, , context] of parse.mock.calls) expect(context).toMatchObject({ workspace: doc.workspace, profile: 'mobile' })
    expect(parse.mock.calls.map(call => call[2]?.partPath)).toEqual(['word/charts/chart1.xml', 'ppt/charts/chart1.xml'])

    parse.mockClear()
    await parseDocx(await docxPackage('line'))
    await parsePptx(await pptxPackage('pie'))
    const [first, second] = parse.mock.calls.map(call => call[2])
    expect(first?.profile).toBe('desktop')
    expect(first?.workspace).not.toBe(doc.workspace)
    expect(second?.workspace).not.toBe(first?.workspace)
  })

  it('src/charts is not in the eager graph', () => {
    const hit = (id: string) => forbiddenEager({ 'index.js': { modules: [id], imports: [], dynamicImports: [] } }, EAGER_FORBIDDEN).length > 0
    for (const id of ['/r/src/charts/index.ts', '/r/src/charts/parse.ts', '/r/src/charts/paint.ts', '/r/src/charts/fill.ts', '/r/src/charts/model.ts', '/r/src/xlsx/chart.ts', '/r/src/xlsx/chart.ts?worker']) expect(hit(id), id).toBe(true)
    // The eager loader, the types and the placeholder painter stay allowed.
    for (const id of ['/r/src/xlsx/chartLoader.ts', '/r/src/xlsx/chartModel.ts', '/r/src/xlsx/chartDescription.ts', '/r/src/drawing/content-paint.ts']) expect(hit(id), id).toBe(false)
    // The built module map (OV_MODULE_MAP=1 bun run build), when present.
    const file = join(__dirname, '..', 'dist', '.module-map.json')
    if (!existsSync(file)) return
    const map = JSON.parse(readFileSync(file, 'utf8')) as ModuleMap
    if (!map['index.js']) return
    const eager = [...eagerModules(map)].map(id => id.replace(/\\/g, '/'))
    expect(eager.filter(id => /src\/charts\/|src\/xlsx\/chart\.ts/.test(id))).toEqual([])
    expect(forbiddenEager(map, EAGER_FORBIDDEN)).toEqual([])
    // The engine is built, lazily.
    expect(Object.values(map).some(chunk => chunk.modules.some(id => id.replace(/\\/g, '/').includes('src/charts/index.ts')))).toBe(true)
  })
})
