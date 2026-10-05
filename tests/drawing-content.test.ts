import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { drawingPartContext, reserveDrawingNode } from '../src/drawing/parts'
import { loadDrawingContent, reserveDrawingContent } from '../src/drawing/content'
import { paintDrawingContent } from '../src/drawing/content-paint'
import { parseThemeContext } from '../src/drawing/style'
import { resolveGeometry } from '../src/drawing/geometry'
import { parseDocx } from '../src/docx/parse'
import { getChildren, parseXmlOrdered } from '../src/core/xml'
import { loadDrawingParts, paintDrawing } from '../src/docx/drawing'
import { parsePptx } from '../src/pptx/parse'
import { renderSlide } from '../src/pptx/render'

declare global { interface ImportMeta { glob: (pattern: string) => Record<string, unknown> } }

const emu = (n: number) => n * 9525
const rel = (id: string, target: string, external = false) => `<Relationship Id="${id}" Target="${target}" Type="x/content"${external ? ' TargetMode="External"' : ''}/>`
const rels = (value: string) => `<Relationships>${value}</Relationships>`
const transform = `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(100)}"/></a:xfrm>`
const diagram = '<a:graphicData uri="diagram"><dgm:relIds r:dm="dm"/></a:graphicData>'
const chart = '<a:graphicData uri="chart"><c:chart r:id="chart"/></a:graphicData>'
const ink = '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"><p:contentPart r:id="ink"/></a:graphicData>'
const leaf = `<dsp:sp><dsp:spPr>${transform}<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></dsp:spPr></dsp:sp>`
const data = '<dgm:dataModel><dsp:dataModelExt relId="cached"/></dgm:dataModel>'
const cached = `<dsp:drawing><dsp:spTree>${leaf}</dsp:spTree></dsp:drawing>`
const columns = '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
const inkml = '<ink><traceFormat><channel name="X" units="px"/><channel name="Y" units="px"/></traceFormat><brush xml:id="b"><brushProperty name="width" value="5" units="px"/><brushProperty name="color" value="#0000FF"/></brush><trace brushRef="#b">0 0, 100 100</trace></ink>'
const parts = {
  'assets/data.xml': data, 'assets/cached.xml': cached,
  'assets/_rels/data.xml.rels': rels(rel('cached', './cached.xml')),
  'assets/chart.xml': columns, 'assets/ink.xml': inkml,
}
// MC fixtures name real understood namespaces. The parser resolves Requires
// through URI bindings, so attach those bindings to the fixture root.
const fixtureNamespaces: Record<string, string> = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  dsp: 'http://schemas.microsoft.com/office/drawing/2008/diagram',
  ink: 'http://www.w3.org/2003/InkML',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
}
function bindFixtureNamespaces(xml: string): string {
  if (!xml.includes('mc:AlternateContent')) return xml
  const end = xml.indexOf('>')
  const opening = xml.slice(0, end)
  const declarations = Object.entries(fixtureNamespaces).filter(([prefix]) => !opening.includes(`xmlns:${prefix}=`))
    .map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`).join('')
  return `${xml.slice(0, end)}${declarations}${xml.slice(end)}`
}
async function pkg(extra: Record<string, string> = {}) {
  const zip = new JSZip()
  for (const [path, xml] of Object.entries({ ...parts, ...extra })) zip.file(path, bindFixtureNamespaces(xml))
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}
const ownerRels = rels(rel('dm', '../../assets/data.xml') + rel('chart', '../../assets/chart.xml') + rel('ink', '../../assets/ink.xml'))
const frame = (content: string, id = 1) => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}"/></p:nvGraphicFramePr>${transform}<a:graphic>${content}</a:graphic></p:graphicFrame>`
async function pptx(body: string, extra: Record<string, string> = {}) {
  const package_ = await pkg({
    'ppt/presentation.xml': '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': rels(rel('s', 'slides/slide1.xml')),
    'ppt/slides/slide1.xml': `<p:sld><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`,
    'ppt/slides/_rels/slide1.xml.rels': ownerRels, ...extra,
  })
  return { package_, doc: await parsePptx(package_) }
}
const pixel = (ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>, x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data).slice(0, 3)

describe('shared embedded drawing content', () => {
  test('Word cached labels paint direct DrawingML alpha over the prior canvas color', async () => {
    const loadLabel = async (alpha: string) => {
      const cache = `<dsp:drawing><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></dsp:spPr><dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="FF0000">${alpha}</a:srgbClr></a:solidFill></a:rPr><a:t>Red</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`
      const package_ = await pkg({
        'assets/cached.xml': cache,
        'word/_rels/document.xml.rels': rels(rel('dm', '../assets/cached.xml')),
      })
      const root = parseXmlOrdered(`<w:document><w:body><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(100)}"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing></w:body></w:document>`)
      const drawings = await loadDrawingParts(package_, root, 'word/document.xml', { colors: new Map(), fonts: new Map() }, () => ({ runs: [], images: [], align: 'left' }))
      const drawing = [...drawings.values()][0]?.drawing
      expect(drawing?.kind).toBe('diagram')
      const ctx = createCanvas(200, 100).getContext('2d')
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, 200, 100)
      ctx.fillStyle = '#00ff00'
      const painted: string[] = []
      const fillText = ctx.fillText.bind(ctx)
      ctx.fillText = (text, x, y, maxWidth) => {
        painted.push(ctx.fillStyle.toString())
        fillText(text, x, y, maxWidth)
      }
      paintDrawing(drawing!, ctx as unknown as CanvasRenderingContext2D, 200, 100)
      expect(painted).toHaveLength(1)
      expect(ctx.fillStyle).toBe('#00ff00')
      return { ctx, color: painted[0], run: drawing?.kind === 'diagram' ? drawing.shapes[0].paragraphs[0].runs[0] : undefined }
    }
    const opaque = await loadLabel('')
    const translucent = await loadLabel('<a:alpha val="50000"/>')
    expect(opaque.color).toBe('#ff0000')
    expect(translucent.run?.color).toBe('rgba(255,0,0,0.5)')
    expect(translucent.color).toBe('rgba(255, 0, 0, 0.50)')
    const opaquePixels = opaque.ctx.getImageData(0, 0, 200, 100).data
    const translucentPixels = translucent.ctx.getImageData(0, 0, 200, 100).data
    const solidRed = Array.from({ length: 200 * 100 }, (_, i) => i).find(i => opaquePixels[i * 4 + 1] < 40)
    expect(solidRed).toBeDefined()
    const green = translucentPixels[solidRed! * 4 + 1]
    expect(green).toBeGreaterThan(110)
    expect(green).toBeLessThan(150)
  })
  test('cached diagram labels inherit body/list/paragraph/fontRef fields with direct false and zero', async () => {
    const label = `<dsp:sp><dsp:spPr>${transform}<a:prstGeom prst="rect"/></dsp:spPr><dsp:style><a:fontRef idx="major"><a:srgbClr val="112233"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr vert="eaVert" lIns="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="2400" b="1" spc="0"/></a:lvl1pPr></a:lstStyle><a:p><a:pPr><a:defRPr i="1"/></a:pPr><a:r><a:rPr b="0" sz="0"/><a:t>off</a:t></a:r><a:r><a:t>on</a:t></a:r><a:endParaRPr sz="4200"/></a:p></dsp:txBody></dsp:sp>`
    const cache = `<dsp:drawing><dsp:spTree>${label}</dsp:spTree></dsp:drawing>`
    const { doc } = await pptx(frame(diagram), { 'assets/cached.xml': cache })
    const shape = doc.slides[0].shapes[0].content!.kind === 'diagram' ? doc.slides[0].shapes[0].content!.shapes[0] : undefined
    expect(shape?.textBody).toMatchObject({ direction: 'eaVert', insetLeftEmu: 0, anchor: 'ctr' })
    expect(shape?.textBody?.paragraphs[0].runs).toMatchObject([
      { text: 'off', bold: false, fontSizePt: 0, italic: true, color: '#112233', fontFamily: '+mj-lt' },
      { text: 'on', bold: true, fontSizePt: 24, italic: true, color: '#112233', fontFamily: '+mj-lt' },
    ])
    const drawing = await loadDrawingContent(await pkg({ 'assets/cached.xml': cache, 'owner/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml')) }), parseXmlOrdered(`<a:graphicData uri="diagram"><dgm:relIds r:dm="dm"/></a:graphicData>`), 'owner/document.xml', { colors: new Map(), fonts: new Map() }, parseThemeContext())
    const legacy = drawing?.kind === 'diagram' ? drawing.shapes[0] : undefined
    expect(legacy?.paragraphs[0].runs).toMatchObject([{ bold: false, fontSizePt: 0 }, { bold: true, fontSizePt: 24 }])
    const wordPackage = await pkg({ 'assets/cached.xml': cache, 'word/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml')) })
    const wordRoot = parseXmlOrdered(`<w:document><w:body><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing></w:body></w:document>`)
    const wordDrawings = await loadDrawingParts(wordPackage, wordRoot, 'word/document.xml', { colors: new Map(), fonts: new Map() }, () => ({ runs: [], images: [], align: 'left' }))
    const wordDiagram = [...wordDrawings.values()][0]?.drawing
    expect(wordDiagram?.kind === 'diagram' ? wordDiagram.shapes[0].paragraphs[0].runs : []).toMatchObject([{ bold: false, fontSizePt: 0 }, { bold: true, fontSizePt: 24 }])
  })

describe('word drawing text keeps generic font fallback', () => {
  // Unknown families must retain their CSS generic fallback (", sans-serif")
  // in both measurement and actual paint, never a quoted whole-stack family.
  const captureFonts = (ctx: CanvasRenderingContext2D) => {
    const fonts: string[] = []
    const proxy = new Proxy(ctx, {
      get(t, p) {
        const v = Reflect.get(t, p, t)
        return typeof v === 'function' ? (v as (...a: never[]) => unknown).bind(t) : v
      },
      set(t, p, v) {
        if (p === 'font') fonts.push(String(v))
        return Reflect.set(t, p, v)
      },
    })
    return { fonts, proxy: proxy as unknown as CanvasRenderingContext2D }
  }
  const expectGenericFallback = (fonts: string[], family: string) => {
    const matching = fonts.filter(f => f.includes(family))
    expect(matching.length).toBeGreaterThan(0)
    for (const f of matching) expect(f).toContain(', sans-serif')
    expect(fonts.some(f => f.includes('\\"'))).toBe(false)
  }
  test('textbox runs keep generic fallback in measure and paint', async () => {
    const { CT_TYPES, ROOT_RELS } = await import('../src/testdata/ooxml-builders')
    const { parseDocx } = await import('../src/docx/parse')
    const { collectDocImages } = await import('../src/docx/layout')
    const { OfficePackage } = await import('../src/core/zip')
    const WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
    const inner = `<w:p><w:r><w:rPr><w:rFonts w:ascii="Liter" w:hAnsi="Liter"/><w:sz w:val="24"/></w:rPr><w:t>Hi</w:t></w:r></w:p>`
    const box = `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CT_TYPES)
    zip.file('_rels/.rels', ROOT_RELS)
    zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(100)}"/><wp:docPr id="1" name="b1"/><a:graphic><a:graphicData>${box}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`)
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = collectDocImages(doc)[0]?.drawing
    if (drawing?.kind !== 'textbox') throw Error(`expected textbox, got ${drawing?.kind}`)
    expect(drawing.paragraphs[0].runs[0].fontFamily).toBe('Liter')
    const ctx = createCanvas(200, 100).getContext('2d')
    const { fonts, proxy } = captureFonts(ctx as unknown as CanvasRenderingContext2D)
    paintDrawing(drawing, proxy, 200, 100)
    expectGenericFallback(fonts, 'Liter')
  })
  test('cached diagram labels keep generic fallback in measure and paint', async () => {
    const cache = `<dsp:drawing><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></dsp:spPr><dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Liter"/></a:rPr><a:t>Red</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`
    const package_ = await pkg({
      'assets/cached.xml': cache,
      'word/_rels/document.xml.rels': rels(rel('dm', '../assets/cached.xml')),
    })
    const root = parseXmlOrdered(`<w:document><w:body><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(100)}"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing></w:body></w:document>`)
    const drawings = await loadDrawingParts(package_, root, 'word/document.xml', { colors: new Map(), fonts: new Map() }, () => ({ runs: [], images: [], align: 'left' }))
    const drawing = [...drawings.values()][0]?.drawing
    if (drawing?.kind !== 'diagram') throw Error(`expected diagram, got ${drawing?.kind}`)
    expect(drawing.shapes[0].textBody?.paragraphs[0].runs[0].fontFamily).toBe('Liter')
    const ctx = createCanvas(200, 100).getContext('2d')
    const { fonts, proxy } = captureFonts(ctx as unknown as CanvasRenderingContext2D)
    paintDrawing(drawing!, proxy, 200, 100)
    expectGenericFallback(fonts, 'Liter')
  })
})
  test('shared loading and payload dispatch modules are available without a format runtime dependency', () => {
    const modules = import.meta.glob('../src/drawing/{parts,content,content-paint}.ts')
    expect(Object.keys(modules).sort()).toEqual(['../src/drawing/content-paint.ts', '../src/drawing/content.ts', '../src/drawing/parts.ts'])
  })
  test.each([[diagram, 'diagram', [255, 0, 0], 50, 50], [chart, 'chart', [0, 255, 0], 50, 60], [ink, 'ink', [0, 0, 255], 50, 50]] as const)('Word and PPTX paint the same supported %s payload', async (content, kind, color, x, y) => {
    const package_ = await pkg({ 'word/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml') + rel('chart', '../assets/chart.xml') + rel('ink', '../assets/ink.xml')) })
    const root = parseXmlOrdered(`<w:document><w:body><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${content}</a:graphic></wp:inline></w:drawing></w:body></w:document>`)
    const drawings = await loadDrawingParts(package_, root, 'word/document.xml', { colors: new Map(), fonts: new Map() }, () => ({ runs: [], images: [], align: 'left' }))
    const drawing = [...drawings.values()][0]?.drawing
    expect(drawing?.kind).toBe(kind)
    const wordContext = createCanvas(100, 100).getContext('2d')
    wordContext.fillStyle = '#ffffff'; wordContext.fillRect(0, 0, 100, 100)
    paintDrawing(drawing!, wordContext as unknown as CanvasRenderingContext2D, 100, 100)
    expect(pixel(wordContext, x, y)).toEqual(color)
    const { doc } = await pptx(frame(content))
    expect((doc.slides[0].shapes[0] as { content?: { kind: string } })?.content?.kind).toBe(kind)
    const pptContext = createCanvas(100, 100).getContext('2d')
    renderSlide(doc.slides[0], pptContext as unknown as CanvasRenderingContext2D)
    expect(pixel(pptContext, x, y)).toEqual(color)
  })
  test.each([
    ['chart Choice', chart, diagram, 'c', 'chart', [0, 255, 0], 50, 60, 1],
    ['diagram Choice', diagram, ink, 'dgm', 'diagram', [255, 0, 0], 50, 50, 2],
    ['InkML contentPart Choice', ink, chart, 'ink', 'ink', [0, 0, 255], 50, 50, 1],
    ['chart Fallback', ink, chart, 'future', 'chart', [0, 255, 0], 50, 60, 1],
    ['InkML contentPart Fallback', diagram, ink, 'future', 'ink', [0, 0, 255], 50, 50, 1],
  ] as const)('PPTX paints nested selected %s from its original content node', async (_label, choice, fallback, requires, kind, color, x, y, nodes) => {
    const alternate = `<mc:AlternateContent><mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    const { doc, package_ } = await pptx(frame(alternate))
    expect(doc.slides[0].shapes[0]?.content?.kind).toBe(kind)
    expect(drawingPartContext(package_).nodes).toBe(nodes)
    const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, x, y)).toEqual(color)
  })
  test.each([
    ['chart Choice', chart, ink, 'c', 'chart', [0, 255, 0], 50, 60, 1],
    ['diagram Choice', diagram, ink, 'dgm', 'diagram', [255, 0, 0], 50, 50, 2],
    ['InkML contentPart Fallback', chart, ink, 'future', 'ink', [0, 0, 255], 50, 50, 1],
    ['chart Fallback', diagram, chart, 'future', 'chart', [0, 255, 0], 50, 60, 1],
  ] as const)('PPTX paints selected %s when the graphic itself is wrapped', async (_label, choice, fallback, requires, kind, color, x, y, nodes) => {
    const alternate = `<mc:AlternateContent><mc:Choice Requires="${requires}"><a:graphic>${choice}</a:graphic></mc:Choice><mc:Fallback><a:graphic>${fallback}</a:graphic></mc:Fallback></mc:AlternateContent>`
    const body = frame(choice).replace(`<a:graphic>${choice}</a:graphic>`, alternate)
    const { doc, package_ } = await pptx(body)
    expect(doc.slides[0].shapes[0]?.content?.kind).toBe(kind)
    expect(drawingPartContext(package_).nodes).toBe(nodes)
    const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, x, y)).toEqual(color)
  })
  test.each([
    ['chart Choice', '<c:chart r:id="chart"/>', '<p:contentPart r:id="ink"/>', 'c', 'chart', [0, 255, 0], 50, 60, 1],
    ['diagram Choice', '<dgm:relIds r:dm="dm"/>', '<c:chart r:id="chart"/>', 'dgm', 'diagram', [255, 0, 0], 50, 50, 2],
    ['InkML contentPart Choice', '<p:contentPart r:id="ink"/>', '<c:chart r:id="chart"/>', 'ink', 'ink', [0, 0, 255], 50, 50, 1],
    ['chart Fallback', '<p:contentPart r:id="ink"/>', '<c:chart r:id="chart"/>', 'future', 'chart', [0, 255, 0], 50, 60, 1],
    ['InkML contentPart Fallback', '<c:chart r:id="chart"/>', '<p:contentPart r:id="ink"/>', 'future', 'ink', [0, 0, 255], 50, 50, 1],
    ['InkML Fallback after missing chart', '<c:chart r:id="missing"/>', '<p:contentPart r:id="ink"/>', 'c', 'ink', [0, 0, 255], 50, 50, 1],
    ['InkML Fallback after unsupported OMML', '<m:oMath/><c:chart r:id="chart"/>', '<p:contentPart r:id="ink"/>', 'c', 'ink', [0, 0, 255], 50, 50, 1],
  ] as const)('PPTX paints selected %s when the payload child is wrapped', async (_label, choice, fallback, requires, kind, color, x, y, nodes) => {
    const alternate = `<mc:AlternateContent><mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    const { doc, package_ } = await pptx(frame(`<a:graphicData uri="chart">${alternate}</a:graphicData>`))
    expect(doc.slides[0].shapes[0]?.content?.kind).toBe(kind)
    expect(drawingPartContext(package_).nodes).toBe(nodes)
    const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, x, y)).toEqual(color)
  })
  test('nested payload Choice leaves an unselected cached diagram unmaterialized', async () => {
    const alternate = '<a:graphicData uri="chart"><mc:AlternateContent><mc:Choice Requires="c"><c:chart r:id="chart"/></mc:Choice><mc:Fallback><dgm:relIds r:dm="dm"/></mc:Fallback></mc:AlternateContent></a:graphicData>'
    const package_ = await pkg({
      'assets/cached.xml': `<dsp:drawing><dsp:spTree>${leaf.repeat(1001)}</dsp:spTree></dsp:drawing>`,
      'ppt/presentation.xml': '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s', 'slides/slide1.xml')),
      'ppt/slides/slide1.xml': `<p:sld><p:cSld><p:spTree>${frame(alternate)}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': ownerRels,
    })
    const cache = await package_.xmlOrdered('assets/cached.xml')
    let materialized = 0
    for (const node of getChildren(getChildren(cache, 'spTree')[0], 'sp')) {
      const pr = node.spPr
      Object.defineProperty(node, 'spPr', { get() { materialized++; return pr }, configurable: true, enumerable: true })
    }
    const doc = await parsePptx(package_)
    expect(doc.slides[0].shapes[0]?.content?.kind).toBe('chart')
    expect(materialized).toBe(0)
    expect(drawingPartContext(package_).nodes).toBe(1)
  })
  test.each([
    ['double supported InkML Choice', '<mc:AlternateContent><mc:Choice Requires="a"><mc:AlternateContent><mc:Choice Requires="ink"><p:contentPart r:id="ink"/></mc:Choice><mc:Fallback><c:chart r:id="chart"/></mc:Fallback></mc:AlternateContent></mc:Choice><mc:Fallback><c:chart r:id="chart"/></mc:Fallback></mc:AlternateContent>', 'ink', [0, 0, 255], 50, 50],
    ['double missing chart with inner InkML Fallback', '<mc:AlternateContent><mc:Choice Requires="a"><mc:AlternateContent><mc:Choice Requires="c"><c:chart r:id="missing"/></mc:Choice><mc:Fallback><p:contentPart r:id="ink"/></mc:Fallback></mc:AlternateContent></mc:Choice><mc:Fallback><c:chart r:id="chart"/></mc:Fallback></mc:AlternateContent>', 'ink', [0, 0, 255], 50, 50],
    ['OMML before contentPart', '<mc:AlternateContent><mc:Choice Requires="a"><m:oMath/><p:contentPart r:id="ink"/></mc:Choice><mc:Fallback><c:chart r:id="chart"/></mc:Fallback></mc:AlternateContent>', 'chart', [0, 255, 0], 50, 60],
    ['OMML after contentPart', '<mc:AlternateContent><mc:Choice Requires="a"><p:contentPart r:id="ink"/><m:oMath/></mc:Choice><mc:Fallback><c:chart r:id="chart"/></mc:Fallback></mc:AlternateContent>', 'chart', [0, 255, 0], 50, 60],
  ] as const)('PPTX resolves recursive payload %s before selecting a branch', async (_label, xml, kind, color, x, y) => {
    const { doc, package_ } = await pptx(frame(`<a:graphicData uri="chart">${xml}</a:graphicData>`))
    expect(doc.slides[0].shapes[0]?.content?.kind).toBe(kind)
    expect(drawingPartContext(package_).nodes).toBe(1)
    const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, x, y)).toEqual(color)
  })
  test.each([
    ['chart Choice', `<mc:AlternateContent><mc:Choice Requires="c">${chart}</mc:Choice><mc:Fallback>${ink}</mc:Fallback></mc:AlternateContent>`, 'chart', [0, 255, 0], 50, 60],
    ['missing chart with InkML Fallback', `<mc:AlternateContent><mc:Choice Requires="c">${chart.replace('id="chart"', 'id="missing"')}</mc:Choice><mc:Fallback>${ink}</mc:Fallback></mc:AlternateContent>`, 'ink', [0, 0, 255], 50, 50],
    ['double nested InkML Choice', `<mc:AlternateContent><mc:Choice Requires="a"><mc:AlternateContent><mc:Choice Requires="ink">${ink}</mc:Choice><mc:Fallback>${chart}</mc:Fallback></mc:AlternateContent></mc:Choice><mc:Fallback>${chart}</mc:Fallback></mc:AlternateContent>`, 'ink', [0, 0, 255], 50, 50],
    ['payload child chart Choice', '<a:graphicData uri="chart"><mc:AlternateContent><mc:Choice Requires="c"><c:chart r:id="chart"/></mc:Choice><mc:Fallback><p:contentPart r:id="ink"/></mc:Fallback></mc:AlternateContent></a:graphicData>', 'chart', [0, 255, 0], 50, 60],
    ['payload child double missing chart with InkML Fallback', '<a:graphicData uri="chart"><mc:AlternateContent><mc:Choice Requires="a"><mc:AlternateContent><mc:Choice Requires="c"><c:chart r:id="missing"/></mc:Choice><mc:Fallback><p:contentPart r:id="ink"/></mc:Fallback></mc:AlternateContent></mc:Choice><mc:Fallback><c:chart r:id="chart"/></mc:Fallback></mc:AlternateContent></a:graphicData>', 'ink', [0, 0, 255], 50, 50],
    ['OMML before contentPart', '<mc:AlternateContent><mc:Choice Requires="a"><a:graphicData><m:oMath/><p:contentPart r:id="ink"/></a:graphicData></mc:Choice><mc:Fallback><a:graphicData uri="chart"><c:chart r:id="chart"/></a:graphicData></mc:Fallback></mc:AlternateContent>', 'chart', [0, 255, 0], 50, 60],
    ['OMML after contentPart', `<mc:AlternateContent><mc:Choice Requires="a"><a:graphicData>${'<p:contentPart r:id="ink"/><m:oMath/>'}</a:graphicData></mc:Choice><mc:Fallback>${chart}</mc:Fallback></mc:AlternateContent>`, 'chart', [0, 255, 0], 50, 60],
  ] as const)('Word paints selected nested %s and retains surrounding text', async (_label, content, kind, color, x, y) => {
    const package_ = await pkg({
      'word/document.xml': `<w:document><w:body><w:p><w:r><w:t>before</w:t><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${content}</a:graphic></wp:inline></w:drawing><w:t>after</w:t></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('chart', '../assets/chart.xml') + rel('ink', '../assets/ink.xml')),
    })
    const doc = await parseDocx(package_)
    const paragraph = doc.sections[0].paragraphs[0]
    expect(paragraph.runs.map(run => run.text).join('')).toBe('beforeafter')
    expect(paragraph.images[0]?.drawing?.kind).toBe(kind)
    expect(drawingPartContext(package_).nodes).toBe(1)
    const ctx = createCanvas(100, 100).getContext('2d'); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, 100, 100)
    paintDrawing(paragraph.images[0].drawing!, ctx as unknown as CanvasRenderingContext2D, 100, 100)
    expect(pixel(ctx, x, y)).toEqual(color)
  })
  test('Word selects a chart when compatibility wraps the graphic child itself', async () => {
    const wrapped = `<mc:AlternateContent><mc:Choice Requires="c"><a:graphic>${chart}</a:graphic></mc:Choice><mc:Fallback><a:graphic>${ink}</a:graphic></mc:Fallback></mc:AlternateContent>`
    const package_ = await pkg({
      'word/document.xml': `<w:document><w:body><w:p><w:r><w:t>before</w:t><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/>${wrapped}</wp:inline></w:drawing><w:t>after</w:t></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('chart', '../assets/chart.xml') + rel('ink', '../assets/ink.xml')),
    })
    const doc = await parseDocx(package_)
    const paragraph = doc.sections[0].paragraphs[0]
    expect(paragraph.runs.map(run => run.text).join('')).toBe('beforeafter')
    expect(paragraph.images[0]?.drawing?.kind).toBe('chart')
    expect(drawingPartContext(package_).nodes).toBe(1)
    const ctx = createCanvas(100, 100).getContext('2d'); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, 100, 100)
    paintDrawing(paragraph.images[0].drawing!, ctx as unknown as CanvasRenderingContext2D, 100, 100)
    expect(pixel(ctx, 50, 60)).toEqual([0, 255, 0])
  })
  test.each([
    ['metadata Choice before graphic', 'a', false],
    ['metadata Choice after graphic', 'a', true],
    ['metadata Fallback before graphic', 'future', false],
  ] as const)('Word keeps chart content with %s', async (_label, requires, after) => {
    const metadata = '<wp:docPr id="1" name="Chart"/>'
    const alternate = `<mc:AlternateContent><mc:Choice Requires="${requires}">${metadata}</mc:Choice><mc:Fallback>${metadata}</mc:Fallback></mc:AlternateContent>`
    const graphic = `<a:graphic>${chart}</a:graphic>`
    const placement = after ? graphic + alternate : alternate + graphic
    const package_ = await pkg({
      'word/document.xml': `<w:document><w:body><w:p><w:r><w:t>before</w:t><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/>${placement}</wp:inline></w:drawing><w:t>after</w:t></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('chart', '../assets/chart.xml')),
    })
    const doc = await parseDocx(package_)
    const paragraph = doc.sections[0].paragraphs[0]
    expect(paragraph.runs.map(run => run.text).join('')).toBe('beforeafter')
    expect(paragraph.images[0]?.drawing?.kind).toBe('chart')
    expect(drawingPartContext(package_).nodes).toBe(1)
    const ctx = createCanvas(100, 100).getContext('2d'); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, 100, 100)
    paintDrawing(paragraph.images[0].drawing!, ctx as unknown as CanvasRenderingContext2D, 100, 100)
    expect(pixel(ctx, 50, 60)).toEqual([0, 255, 0])
  })
  test('recursive selected chart keeps a valid neighbor and defers unselected cached leaves', async () => {
    const inner = '<mc:AlternateContent><mc:Choice Requires="c"><c:chart r:id="chart"/></mc:Choice><mc:Fallback><dgm:relIds r:dm="dm"/></mc:Fallback></mc:AlternateContent>'
    const outer = `<mc:AlternateContent><mc:Choice Requires="a">${inner}</mc:Choice><mc:Fallback><p:contentPart r:id="ink"/></mc:Fallback></mc:AlternateContent>`
    const neighbor = `<p:sp><p:spPr>${transform.replace('<a:off x="0" y="0"/>', `<a:off x="${emu(80)}" y="0"/>`).replace(`cx="${emu(100)}" cy="${emu(100)}"`, `cx="${emu(20)}" cy="${emu(20)}"`)}<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></p:spPr></p:sp>`
    const package_ = await pkg({
      'assets/cached.xml': `<dsp:drawing><dsp:spTree>${leaf.repeat(1001)}</dsp:spTree></dsp:drawing>`,
      'ppt/presentation.xml': '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s', 'slides/slide1.xml')),
      'ppt/slides/slide1.xml': `<p:sld><p:cSld><p:spTree>${frame(`<a:graphicData uri="chart">${outer}</a:graphicData>`)}${neighbor}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': ownerRels,
    })
    const cache = await package_.xmlOrdered('assets/cached.xml')
    let materialized = 0
    for (const node of getChildren(getChildren(cache, 'spTree')[0], 'sp')) {
      const pr = node.spPr
      Object.defineProperty(node, 'spPr', { get() { materialized++; return pr }, configurable: true, enumerable: true })
    }
    const doc = await parsePptx(package_)
    expect(doc.slides[0].shapes.map(shape => shape.content?.kind ?? shape.geometry)).toEqual(['chart', 'rect'])
    expect(materialized).toBe(0)
    expect(drawingPartContext(package_).nodes).toBe(2)
    const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, 50, 60)).toEqual([0, 255, 0])
    expect(pixel(ctx, 90, 10)).toEqual([255, 255, 0])
  })
  test('missing/external charts select one native fallback and keep a valid neighbor', async () => {
    const fallback = `<p:sp><p:spPr>${transform}<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></p:spPr></p:sp>`
    const alternate = (id: string) => `<mc:AlternateContent><mc:Choice Requires="a">${frame(chart.replace('id="chart"', `id="${id}"`))}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    const { doc } = await pptx(alternate('missing') + alternate('external') + frame(diagram, 3), { 'ppt/slides/_rels/slide1.xml.rels': ownerRels.replace('</Relationships>', rel('external', '../../assets/chart.xml', true) + '</Relationships>') })
    expect(doc.slides[0].shapes).toHaveLength(3)
    expect(doc.slides[0].shapes.map(s => s.source?.representation)).toEqual(['fallback', 'fallback', 'native'])
    expect((doc.slides[0].shapes[2] as { content?: { kind: string } }).content?.kind).toBe('diagram')
  })
})

const contentTheme = { colors: new Map<string, string>(), fonts: new Map<string, string>() }
const diagramGroup = (children: string) => `<dsp:grpSp><dsp:grpSpPr>${transform.replace('</a:xfrm>', '<a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm>')}</dsp:grpSpPr>${children}</dsp:grpSp>`
async function shared(package_: OfficePackage, xml: string, owner = 'owner/document.xml') {
  return loadDrawingContent(package_, parseXmlOrdered(xml), owner, contentTheme, parseThemeContext())
}
function chain(count: number, cycle = false): Record<string, string> {
  const out: Record<string, string> = { 'owner/_rels/document.xml.rels': rels(rel('first', '../chain/0.xml')) }
  for (let i = 0; i < count; i++) {
    out[`chain/${i}.xml`] = i === count - 1 && !cycle ? inkml : '<root><p:contentPart r:id="next"/></root>'
    out[`chain/_rels/${i}.xml.rels`] = rels(rel('next', `${i === count - 1 ? 0 : i + 1}.xml`))
  }
  return out
}
const contentRef = '<root><p:contentPart r:id="first"/></root>'

describe('embedded content traversal boundaries', () => {
  test.each([
    ['empty', '<dsp:drawing><dsp:spTree/></dsp:drawing>'],
    ['picture only', '<dsp:drawing><dsp:spTree><dsp:pic><dsp:blipFill><a:blip r:embed="missing"/></dsp:blipFill></dsp:pic></dsp:spTree></dsp:drawing>'],
    ['empty group', `<dsp:drawing><dsp:spTree>${diagramGroup('')}</dsp:spTree></dsp:drawing>`],
    ['invalid group', `<dsp:drawing><dsp:spTree>${diagramGroup(leaf).replace('<a:chExt cx="952500" cy="952500"/>', '<a:chExt cx="0" cy="0"/>')}</dsp:spTree></dsp:drawing>`],
  ])('PPTX selects its usable raster fallback for %s cached content', async (_label, cache) => {
    const fallback = leaf.replace(/dsp:sp/g, 'p:sp').replace('FF0000', 'FFFF00')
    const alternate = `<mc:AlternateContent><mc:Choice Requires="a">${frame(diagram)}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    const { doc, package_ } = await pptx(alternate, { 'assets/cached.xml': cache })
    expect(doc.slides[0].shapes[0].source?.representation).toBe('fallback')
    const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, 50, 50)).toEqual([255, 255, 0])
    if (_label === 'picture only') {
      expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'unsupported-content', feature: 'pic' }))
      expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'missing-part', identity: 'missing' }))
    }
  })
  test('a supported transparent cached leaf still selects the diagram Choice', async () => {
    const transparent = `<dsp:drawing><dsp:spTree><dsp:sp><dsp:spPr>${transform}<a:prstGeom prst="rect"/><a:noFill/></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`
    const fallback = `<p:sp><p:spPr>${transform}<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></p:spPr></p:sp>`
    const { doc } = await pptx(`<mc:AlternateContent><mc:Choice Requires="a">${frame(diagram)}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`, { 'assets/cached.xml': transparent })
    expect(doc.slides[0].shapes[0].source?.representation).toBe('choice')
    expect(doc.slides[0].shapes[0].content).toMatchObject({ kind: 'diagram', shapes: [expect.any(Object)] })
  })
  test('malformed cached XML and relationship XML are isolated from PPTX fallback and Word neighbors', async () => {
    const fallback = `<p:sp><p:spPr>${transform}<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></p:spPr></p:sp>`
    const alternate = `<mc:AlternateContent><mc:Choice Requires="a">${frame(diagram)}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    const { doc, package_ } = await pptx(alternate + frame(ink, 2), { 'assets/cached.xml': '<dsp:drawing><dsp:spTree>' })
    expect(doc.slides[0].shapes.map(shape => shape.source?.representation)).toEqual(['fallback', 'native'])
    expect(doc.slides[0].shapes[1].content?.kind).toBe('ink')
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'malformed-part', partPath: 'assets/cached.xml' }))
    const drawing = (graphic: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${graphic}</a:graphic></wp:inline></w:drawing>`
    const wordPackage = await pkg({
      'word/document.xml': `<w:document><w:body><w:p><w:r><w:t>before</w:t>${drawing(diagram)}${drawing(ink)}<w:t>after</w:t></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml') + rel('ink', '../assets/ink.xml')),
      'assets/cached.xml': '<dsp:drawing><dsp:spTree>',
    })
    const word = await parseDocx(wordPackage)
    expect(word.sections[0].paragraphs[0].runs.map(run => run.text).join('')).toBe('beforeafter')
    expect(word.sections[0].paragraphs[0].images.map(image => image.drawing?.kind)).toEqual(['ink'])
    const malformedRels = await pkg({ 'owner/_rels/document.xml.rels': '<Relationships><Relationship' })
    expect(await shared(malformedRels, contentRef)).toBeUndefined()
    expect(drawingPartContext(malformedRels).diagnostics).toContainEqual(expect.objectContaining({ kind: 'malformed-part', partPath: 'owner/_rels/document.xml.rels' }))
    const brokenDiagramRels = await pptx(alternate + frame(ink, 2), { 'assets/_rels/data.xml.rels': '<Relationships><Relationship' })
    expect(brokenDiagramRels.doc.slides[0].shapes.map(shape => shape.source?.representation)).toEqual(['fallback', 'native'])
    expect(brokenDiagramRels.doc.slides[0].shapes[1].content?.kind).toBe('ink')
    expect(drawingPartContext(brokenDiagramRels.package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'malformed-part', partPath: 'assets/_rels/data.xml.rels' }))
  })
  test('malformed Word owner relationships retain body text and independent owner drawings', async () => {
    const drawing = `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${ink}</a:graphic></wp:inline></w:drawing>`
    const package_ = await pkg({
      'word/document.xml': `<w:document><w:body><w:p><w:r><w:t>before</w:t>${drawing}<w:t>after</w:t></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': '<Relationships><Relationship',
      'word/header1.xml': `<w:hdr><w:p><w:r>${drawing}</w:r></w:p></w:hdr>`,
      'word/_rels/header1.xml.rels': rels(rel('ink', '../assets/ink.xml')),
    })
    const doc = await parseDocx(package_)
    expect(doc.sections[0].paragraphs[0].runs.map(run => run.text).join('')).toBe('beforeafter')
    const header = parseXmlOrdered(`<w:hdr><w:p><w:r>${drawing}</w:r></w:p></w:hdr>`)
    const independent = await loadDrawingParts(package_, header, 'word/header1.xml', contentTheme, () => ({ runs: [], images: [], align: 'left' }))
    expect([...independent.values()][0].drawing?.kind).toBe('ink')
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'malformed-part', partPath: 'word/_rels/document.xml.rels' }))
  })
  test('malformed Word header relationships leave body drawing and header text', async () => {
    const drawing = `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${ink}</a:graphic></wp:inline></w:drawing>`
    const package_ = await pkg({
      'word/document.xml': `<w:document><w:body><w:p><w:r>${drawing}</w:r></w:p><w:sectPr><w:headerReference r:id="header"/></w:sectPr></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('ink', '../assets/ink.xml') + rel('header', 'header1.xml')),
      'word/header1.xml': `<w:hdr><w:p><w:r><w:t>Header</w:t>${drawing}</w:r></w:p></w:hdr>`,
      'word/_rels/header1.xml.rels': '<Relationships><Relationship',
    })
    const doc = await parseDocx(package_)
    expect(doc.sections[0].paragraphs[0].images[0].drawing?.kind).toBe('ink')
    expect(doc.sections[0].header![0].runs.map(run => run.text).join('')).toBe('Header')
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'malformed-part', partPath: 'word/_rels/header1.xml.rels' }))
  })
  test('malformed PPTX slide relationships retain native shape and a valid next slide', async () => {
    const native = `<p:sp><p:spPr>${transform}<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></p:spPr></p:sp>`
    const package_ = await pkg({
      'ppt/presentation.xml': '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s1"/><p:sldId r:id="s2"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s1', 'slides/slide1.xml') + rel('s2', 'slides/slide2.xml')),
      'ppt/slides/slide1.xml': `<p:sld><p:cSld><p:spTree>${native}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': '<Relationships><Relationship',
      'ppt/slides/slide2.xml': `<p:sld><p:cSld><p:spTree>${frame(ink)}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide2.xml.rels': ownerRels,
    })
    const doc = await parsePptx(package_)
    expect(doc.slides[0].shapes).toHaveLength(1)
    expect(doc.slides[1].shapes[0].content?.kind).toBe('ink')
    const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, 50, 50)).toEqual([255, 255, 0])
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'malformed-part', partPath: 'ppt/slides/_rels/slide1.xml.rels' }))
  })
  test('an invalid group and throwing text callback preserve neighboring paint and caller state', () => {
    const base = { xEmu: 0, yEmu: 0, widthEmu: emu(20), heightEmu: emu(20), geometry: 'rect', paragraphs: [], fontFamily: '', fill: 'FF0000' }
    const content = { kind: 'diagram' as const, shapes: [
      { ...base, group: { off: { x: 0, y: 0 }, ext: { width: emu(20), height: emu(20) }, chOff: { x: 0, y: 0 }, chExt: { width: 0, height: emu(20) } }, children: [base] },
      { ...base, xEmu: emu(30), paragraphs: [{ runs: [{ text: 'throw' }], align: 'left' as const }] },
      { ...base, xEmu: emu(60) },
    ] }
    const ctx = createCanvas(100, 30).getContext('2d'); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, 100, 30)
    paintDrawingContent(content, ctx as unknown as CanvasRenderingContext2D, 100, 30, { paintDiagramText(shape) { if (shape.paragraphs.length) throw Error('text failure') } })
    expect(pixel(ctx, 70, 10)).toEqual([255, 0, 0])
    expect(pixel(ctx, 10, 10)).toEqual([255, 255, 255])
    expect(ctx.getTransform().a).toBe(1)
    expect(ctx.getTransform().e).toBe(0)
  })
  test('a repeated legitimate reference survives a cyclic sibling and records ancestry diagnostics', async () => {
    const package_ = await pkg({ ...chain(2, true), 'owner/_rels/document.xml.rels': rels(rel('first', '../chain/0.xml') + rel('good', '../assets/ink.xml')) })
    expect(await shared(package_, contentRef)).toBeUndefined()
    const first = await shared(package_, contentRef.replace('first', 'good'))
    const second = await shared(package_, contentRef.replace('first', 'good'))
    expect(first?.kind).toBe('ink'); expect(second).toEqual(first)
    expect(second).not.toBe(first)
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'content-cycle' }))
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'content-cycle', partPath: 'chain/1.xml', identity: 'chain/0.xml', reason: 'ancestry' }))
    expect(drawingPartContext(package_).coverage.repeated).toBeGreaterThan(0)
  })
  test('32 acyclic content references load, a deeper branch stops, and a valid neighbor survives', async () => {
    const atLimit = await pkg(chain(32))
    expect((await shared(atLimit, contentRef))?.kind).toBe('ink')
    const overLimit = await pkg({ ...chain(34), 'owner/_rels/document.xml.rels': rels(rel('first', '../chain/0.xml') + rel('good', '../assets/chart.xml')) })
    expect(await shared(overLimit, contentRef)).toBeUndefined()
    expect((await shared(overLimit, contentRef.replace('first', 'good')))?.kind).toBe('chart')
    expect(drawingPartContext(overLimit).diagnostics).toContainEqual(expect.objectContaining({ kind: 'content-depth' }))
    expect(drawingPartContext(overLimit).diagnostics).toContainEqual(expect.objectContaining({ kind: 'content-depth', limit: 32, identity: 'chain/32.xml' }))
  })
  test('a deep Word branch leaves surrounding text and a supported neighbor intact', async () => {
    const drawing = (graphic: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${graphic}</a:graphic></wp:inline></w:drawing>`
    const package_ = await pkg({
      ...chain(34),
      'word/document.xml': `<w:document><w:body><w:p><w:r><w:t>before</w:t>${drawing('<a:graphicData><p:contentPart r:id="first"/></a:graphicData>')}${drawing(chart)}<w:t>after</w:t></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('first', '../chain/0.xml') + rel('chart', '../assets/chart.xml')),
    })
    const doc = await parseDocx(package_)
    expect(doc.sections[0].paragraphs[0].runs.map(run => run.text).join('')).toBe('beforeafter')
    expect(doc.sections[0].paragraphs[0].images.map(image => image.drawing?.kind)).toEqual(['chart'])
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'content-depth', limit: 32 }))
  })
  test('64 cached groups survive and an excessive group keeps its valid sibling', async () => {
    const nested = (depth: number) => Array.from({ length: depth }).reduce<string>(xml => diagramGroup(xml), leaf)
    const { doc, package_ } = await pptx(frame(diagram), { 'assets/cached.xml': `<dsp:drawing><dsp:spTree>${nested(65)}${leaf}</dsp:spTree></dsp:drawing>` })
    const content = doc.slides[0].shapes[0].content
    expect(content?.kind).toBe('diagram')
    if (content?.kind !== 'diagram') throw Error('diagram missing')
    let node = content.shapes[0], depth = 0
    while (node.group) { depth++; if (!node.children?.length) break; node = node.children[0] }
    expect(depth).toBe(64)
    expect(content.shapes).toHaveLength(2)
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'group-depth' }))
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'group-depth', limit: 64 }))
    const ctx = createCanvas(100, 100).getContext('2d')
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, 50, 50)).toEqual([255, 0, 0])
  })
  test('mixed cached paths and group transforms retain order in both adapters', async () => {
    const custom = leaf.replace('<a:prstGeom prst="rect"/>', '<a:custGeom><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:moveTo><a:pt x="100" y="100"/></a:moveTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo></a:path></a:pathLst></a:custGeom>').replace('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>', '<a:ln w="38100"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>')
    const cache = `<dsp:drawing><dsp:spTree>${diagramGroup(custom).replace('cx="952500" cy="952500"/></a:xfrm>', 'cx="1905000" cy="1905000"/></a:xfrm>')}${leaf.replace('FF0000', '00FF00').replace('cx="952500" cy="952500"', 'cx="238125" cy="238125"')}</dsp:spTree></dsp:drawing>`
    const { doc } = await pptx(frame(diagram), { 'assets/cached.xml': cache })
    const content = doc.slides[0].shapes[0].content
    if (content?.kind !== 'diagram') throw Error('diagram missing')
    expect(resolveGeometry(content.shapes[0].children![0].drawingGeometry!, 100, 100).paths[0].commands.map(c => c[0])).toEqual(['moveTo', 'lnTo', 'moveTo', 'lnTo'])
    const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, 10, 10)).toEqual([0, 255, 0]); expect(pixel(ctx, 40, 0)).toEqual([0, 0, 255]); expect(pixel(ctx, 40, 50)).toEqual([0, 0, 255]); expect(pixel(ctx, 40, 25)).toEqual([255, 255, 255])
    const wordPackage = await pkg({ 'assets/cached.xml': cache, 'word/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml')) })
    const word = parseXmlOrdered(`<w:document><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing></w:document>`)
    const model = [...(await loadDrawingParts(wordPackage, word, 'word/document.xml', contentTheme, () => ({ runs: [], images: [], align: 'left' }))).values()][0].drawing!
    const wordCtx = createCanvas(100, 100).getContext('2d'); wordCtx.fillStyle = '#ffffff'; wordCtx.fillRect(0, 0, 100, 100)
    paintDrawing(model, wordCtx as unknown as CanvasRenderingContext2D, 100, 100)
    expect(Buffer.from(wordCtx.getImageData(0, 0, 100, 100).data)).toEqual(Buffer.from(ctx.getImageData(0, 0, 100, 100).data))
  })
  test('cached source XML is reused while slide theme payloads remain independent', async () => {
    const package_ = await pkg({ 'owner/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml')), 'assets/cached.xml': cached.replace('<a:srgbClr val="FF0000"/>', '<a:schemeClr val="accent1"/>') })
    const context = (value: string) => parseThemeContext(`<a:theme><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="${value}"/></a:accent1></a:clrScheme></a:themeElements></a:theme>`)
    const firstXml = await package_.xmlOrdered('assets/cached.xml')
    const red = await loadDrawingContent(package_, parseXmlOrdered(diagram), 'owner/document.xml', contentTheme, context('FF0000'))
    const blue = await loadDrawingContent(package_, parseXmlOrdered(diagram), 'owner/document.xml', contentTheme, context('0000FF'))
    expect(await package_.xmlOrdered('assets/cached.xml')).toBe(firstXml)
    if (red?.kind !== 'diagram' || blue?.kind !== 'diagram') throw Error('diagram missing')
    expect(red.shapes[0].drawingStyle?.fill).toMatchObject({ color: { r: 255, b: 0 } })
    expect(blue.shapes[0].drawingStyle?.fill).toMatchObject({ color: { r: 0, b: 255 } })
  })
  test('image-bearing cached content diagnoses unsupported drawing without losing a vector neighbor', async () => {
    const { doc, package_ } = await pptx(frame(diagram), { 'assets/cached.xml': `<dsp:drawing><dsp:spTree><dsp:pic><dsp:blipFill><a:blip r:embed="missing"/></dsp:blipFill></dsp:pic>${leaf}</dsp:spTree></dsp:drawing>` })
    expect(doc.slides[0].shapes[0].content?.kind).toBe('diagram')
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'unsupported-content', feature: 'pic' }))
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'missing-part', feature: 'missing' }))
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'missing-part', identity: 'missing', reason: 'relationship-not-found' }))
  })
  test('Word compatibility charges only the selected representation across body and header loads', async () => {
    const drawing = (content: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${content}</a:graphic></wp:inline></w:drawing>`
    const package_ = await pkg({
      'word/document.xml': `<w:document><w:body><w:p><w:r><mc:AlternateContent><mc:Choice>${drawing(diagram)}</mc:Choice><mc:Fallback>${drawing(chart)}</mc:Fallback></mc:AlternateContent></w:r></w:p><w:sectPr><w:headerReference r:id="header"/></w:sectPr></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml') + rel('chart', '../assets/chart.xml') + rel('header', 'header1.xml')),
      'word/header1.xml': `<w:hdr><w:p><w:r>${drawing(diagram)}</w:r></w:p></w:hdr>`,
      'word/_rels/header1.xml.rels': rels(rel('dm', '../assets/data.xml')),
    })
    const doc = await parseDocx(package_)
    expect(doc.sections[0].paragraphs[0].images[0].drawing?.kind).toBe('diagram')
    expect(doc.sections[0].header![0].images[0].drawing?.kind).toBe('diagram')
    expect(drawingPartContext(package_).nodes).toBe(4)
    expect(drawingPartContext(package_).coverage.repeated).toBeGreaterThan(0)
  })
  test('Word charges selected raster and vector placements in body order before a header', async () => {
    const raster = '<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="image"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
    const vector = `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing>`
    const package_ = await pkg({
      'assets/pic.png': 'image bytes',
      'word/document.xml': `<w:document><w:body><w:p><w:r>${raster}${vector}<w:t>tail</w:t></w:r></w:p><w:sectPr><w:headerReference r:id="header"/></w:sectPr></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels('<Relationship Id="image" Target="../assets/pic.png" Type="x/image"/>' + rel('dm', '../assets/data.xml') + rel('header', 'header1.xml')),
      'word/header1.xml': `<w:hdr><w:p><w:r>${raster}</w:r></w:p></w:hdr>`,
      'word/_rels/header1.xml.rels': rels('<Relationship Id="image" Target="../assets/pic.png" Type="x/image"/>'),
    })
    for (let i = 0; i < 9999; i++) reserveDrawingNode(drawingPartContext(package_), 'earlier/part.xml')
    const doc = await parseDocx(package_)
    expect(doc.sections[0].paragraphs[0].images).toHaveLength(1)
    expect(doc.sections[0].paragraphs[0].images[0].drawing).toBeUndefined()
    expect(doc.sections[0].paragraphs[0].runs.map(run => run.text).join('')).toContain('tail')
    expect(doc.sections[0].header![0].images).toHaveLength(0)
    expect(drawingPartContext(package_).nodes).toBe(10000)
    expect(drawingPartContext(package_).diagnostics.filter(d => d.kind === 'node-budget')).toHaveLength(1)
  })
  test('Word does not charge an inspected Choice that produces no drawing', async () => {
    const broken = '<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="missing"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
    const vector = `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing>`
    const package_ = await pkg({
      'word/document.xml': `<w:document><w:body><w:p><w:r><mc:AlternateContent><mc:Choice>${broken}</mc:Choice><mc:Choice>${vector}</mc:Choice></mc:AlternateContent></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml')),
    })
    for (let i = 0; i < 9998; i++) reserveDrawingNode(drawingPartContext(package_), 'earlier/part.xml')
    const doc = await parseDocx(package_)
    expect(doc.sections[0].paragraphs[0].images[0].drawing?.kind).toBe('diagram')
    expect(doc.sections[0].paragraphs[0].images[0].drawing?.kind === 'diagram' && doc.sections[0].paragraphs[0].images[0].drawing.shapes).toHaveLength(1)
    expect(drawingPartContext(package_).nodes).toBe(10000)
  })
  test('Word selects a supported vector Fallback in body and header after a missing Choice', async () => {
    const drawing = (graphic: string) => `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${graphic}</a:graphic></wp:inline></w:drawing>`
    const alternate = `<mc:AlternateContent><mc:Choice>${drawing(chart.replace('id="chart"', 'id="missing"'))}</mc:Choice><mc:Fallback>${drawing(ink)}</mc:Fallback></mc:AlternateContent>`
    const package_ = await pkg({
      'word/document.xml': `<w:document><w:body><w:p><w:r><w:t>before</w:t>${alternate}<w:t>after</w:t></w:r></w:p><w:sectPr><w:headerReference r:id="header"/></w:sectPr></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('ink', '../assets/ink.xml') + rel('header', 'header1.xml')),
      'word/header1.xml': `<w:hdr><w:p><w:r>${alternate}</w:r></w:p></w:hdr>`,
      'word/_rels/header1.xml.rels': rels(rel('ink', '../assets/ink.xml')),
    })
    const doc = await parseDocx(package_)
    expect(doc.sections[0].paragraphs[0].runs.map(run => run.text).join('')).toBe('beforeafter')
    expect(doc.sections[0].paragraphs[0].images.map(image => image.drawing?.kind)).toEqual(['ink'])
    expect(doc.sections[0].header![0].images.map(image => image.drawing?.kind)).toEqual(['ink'])
    expect(drawingPartContext(package_).nodes).toBe(2)
    const ctx = createCanvas(100, 100).getContext('2d'); paintDrawing(doc.sections[0].paragraphs[0].images[0].drawing!, ctx as unknown as CanvasRenderingContext2D, 100, 100)
    expect(pixel(ctx, 50, 50)).toEqual([0, 0, 255])
  })
  test('PPTX counts only actual groups through compatibility wrappers', async () => {
    const native = `<p:sp><p:spPr>${transform}<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>`
    const group = (children: string) => `<p:grpSp><p:grpSpPr>${transform.replace('</a:xfrm>', '<a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm>')}</p:grpSpPr>${children}</p:grpSp>`
    const wrapped = `<mc:AlternateContent><mc:Choice Requires="p">${group(native)}</mc:Choice><mc:Fallback>${native}</mc:Fallback></mc:AlternateContent>`
    const atLimit = Array.from({ length: 63 }).reduce<string>(inner => group(inner), wrapped)
    const neighbor = native.replace('FF0000', '00FF00').replace('<a:off x="0" y="0"/>', `<a:off x="${emu(80)}" y="0"/>`).replace('cx="952500" cy="952500"', `cx="${emu(20)}" cy="${emu(20)}"`)
    const { doc, package_ } = await pptx(atLimit + neighbor)
    let node = doc.slides[0].shapes[0], count = 0
    while (node?.group) { count++; node = node.children?.[0]! }
    expect(count).toBe(64)
    expect(node?.geometry).toBe('rect')
    const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(pixel(ctx, 50, 50)).toEqual([255, 0, 0]); expect(pixel(ctx, 90, 10)).toEqual([0, 255, 0])
    expect(ctx.getTransform().a).toBe(1)
    const overLimit = await pptx(Array.from({ length: 64 }).reduce<string>(inner => group(inner), wrapped) + neighbor)
    expect(overLimit.doc.slides[0].shapes.at(-1)?.geometry).toBe('rect')
    expect(drawingPartContext(overLimit.package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'group-depth', limit: 64, reason: 'group-depth', identity: expect.any(String) }))
    expect(drawingPartContext(package_).diagnostics.filter(issue => issue.kind === 'group-depth')).toHaveLength(0)
  })
  test('Word charges a selected VML picture fallback once', async () => {
    const empty = '<w:drawing><wp:anchor><wp:extent cx="952500" cy="952500"/></wp:anchor></w:drawing>'
    const package_ = await pkg({
      'assets/pic.png': 'image bytes',
      'word/document.xml': `<w:document><w:body><w:p><w:r><mc:AlternateContent><mc:Choice>${empty}</mc:Choice><mc:Fallback><v:imagedata r:id="image"/></mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels('<Relationship Id="image" Target="../assets/pic.png" Type="x/image"/>'),
    })
    for (let i = 0; i < 10000; i++) reserveDrawingNode(drawingPartContext(package_), 'earlier/part.xml')
    const doc = await parseDocx(package_)
    expect(doc.sections[0].paragraphs[0].images).toHaveLength(0)
  })
  test('document node budget spans Word body and separate header owners', async () => {
    const drawing = `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing>`
    const package_ = await pkg({ 'word/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml')), 'word/_rels/header1.xml.rels': rels(rel('dm', '../assets/data.xml')) })
    const body = await loadDrawingParts(package_, parseXmlOrdered(`<w:document>${drawing.repeat(5000)}</w:document>`), 'word/document.xml', contentTheme, () => ({ runs: [], images: [], align: 'left' }))
    const header = await loadDrawingParts(package_, parseXmlOrdered(`<w:hdr>${drawing}</w:hdr>`), 'word/header1.xml', contentTheme, () => ({ runs: [], images: [], align: 'left' }))
    expect(body.size).toBe(5000); expect(header.size).toBe(0)
    expect(drawingPartContext(package_).nodes).toBe(10000)
    expect(drawingPartContext(package_).diagnostics.filter(d => d.kind === 'node-budget')).toHaveLength(1)
  })
  test('PPTX charges selected pictures and ordinary shapes across slides', async () => {
    const package_ = await pkg({
      'ppt/presentation.xml': '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s1"/><p:sldId r:id="s2"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s1', 'slides/slide1.xml') + rel('s2', 'slides/slide2.xml')),
      'ppt/slides/slide1.xml': `<p:sld><p:cSld><p:spTree><p:pic><p:spPr>${transform}</p:spPr></p:pic></p:spTree></p:cSld></p:sld>`,
      'ppt/slides/slide2.xml': `<p:sld><p:cSld><p:spTree><p:sp><p:spPr>${transform}<a:prstGeom prst="rect"/></p:spPr></p:sp></p:spTree></p:cSld></p:sld>`,
    })
    for (let i = 0; i < 9999; i++) reserveDrawingNode(drawingPartContext(package_), 'earlier/part.xml')
    const doc = await parsePptx(package_)
    expect(doc.slides.map(slide => slide.shapes.length)).toEqual([1, 0])
    expect(drawingPartContext(package_).nodes).toBe(10000)
    expect(drawingPartContext(package_).diagnostics.filter(d => d.kind === 'node-budget')).toHaveLength(1)
  })
  test('selected repeated cache leaves materialize only within the document budget', async () => {
    const leaves = 1001, placements = 20
    const cache = `<dsp:drawing><dsp:spTree>${leaf.repeat(leaves)}</dsp:spTree></dsp:drawing>`
    const drawing = `<w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing>`
    const word = await pkg({
      'assets/cached.xml': cache,
      'word/document.xml': `<w:document><w:body><w:p><w:r>${drawing.repeat(placements)}<w:t>tail</w:t></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml')),
    })
    const instrument = async (package_: OfficePackage) => {
      const root = await package_.xmlOrdered('assets/cached.xml')
      let count = 0
      for (const node of getChildren(getChildren(root, 'spTree')[0], 'sp')) {
        const pr = node.spPr
        Object.defineProperty(node, 'spPr', { get() { count++; return pr }, configurable: true, enumerable: true })
      }
      return () => count
    }
    const wordCount = await instrument(word)
    const doc = await parseDocx(word)
    expect(wordCount()).toBeLessThanOrEqual(10000)
    expect(drawingPartContext(word).nodes).toBe(10000)
    expect(doc.sections[0].paragraphs[0].runs.map(run => run.text).join('')).toContain('tail')
    const power = await pkg({
      'assets/cached.xml': cache,
      'ppt/presentation.xml': '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s', 'slides/slide1.xml')),
      'ppt/slides/slide1.xml': `<p:sld><p:cSld><p:spTree>${frame(diagram).repeat(placements)}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': ownerRels,
    })
    const powerCount = await instrument(power)
    await parsePptx(power)
    expect(powerCount()).toBeLessThanOrEqual(10000)
    expect(drawingPartContext(power).nodes).toBe(10000)
  })
  test('unselected PPTX cached fallbacks do not materialize leaves or spend their budget', async () => {
    const cache = `<dsp:drawing><dsp:spTree>${leaf.repeat(1001)}</dsp:spTree></dsp:drawing>`
    const native = `<p:sp><p:spPr>${transform}<a:prstGeom prst="rect"/></p:spPr></p:sp>`
    const alternate = `<mc:AlternateContent><mc:Choice Requires="p">${native}</mc:Choice><mc:Fallback>${frame(diagram)}</mc:Fallback></mc:AlternateContent>`
    const package_ = await pkg({
      'assets/cached.xml': cache,
      'ppt/presentation.xml': '<p:presentation><p:sldSz cx="952500" cy="952500"/><p:sldIdLst><p:sldId r:id="s"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': rels(rel('s', 'slides/slide1.xml')),
      'ppt/slides/slide1.xml': `<p:sld><p:cSld><p:spTree>${alternate.repeat(20)}</p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': ownerRels,
    })
    const root = await package_.xmlOrdered('assets/cached.xml')
    let materialized = 0
    for (const node of getChildren(getChildren(root, 'spTree')[0], 'sp')) {
      const pr = node.spPr
      Object.defineProperty(node, 'spPr', { get() { materialized++; return pr }, configurable: true, enumerable: true })
    }
    const doc = await parsePptx(package_)
    expect(materialized).toBe(0)
    expect(doc.slides[0].shapes).toHaveLength(20)
    expect(drawingPartContext(package_).nodes).toBe(20)
  })
  test('selected unsupported cached pictures consume bounded source nodes and one budget diagnostic', async () => {
    const cache = `<dsp:drawing><dsp:spTree>${'<dsp:pic/>'.repeat(10001)}${leaf}</dsp:spTree></dsp:drawing>`
    const { package_, doc } = await pptx(frame(diagram), { 'assets/cached.xml': cache })
    expect(doc.slides[0].shapes[0].content?.kind).toBe('diagram')
    expect(drawingPartContext(package_).nodes).toBe(10000)
    expect(drawingPartContext(package_).diagnostics.filter(issue => issue.kind === 'node-budget')).toEqual([expect.objectContaining({ reason: 'document-budget', limit: 10000 })])
    expect(drawingPartContext(package_).diagnostics.filter(issue => issue.feature === 'pic').length).toBeLessThanOrEqual(64)
  })
  test('direct eager load still charges selected unsupported cached source pictures', async () => {
    const package_ = await pkg({
      'assets/cached.xml': `<dsp:drawing><dsp:spTree>${'<dsp:pic/>'.repeat(10001)}${leaf}</dsp:spTree></dsp:drawing>`,
      'owner/_rels/document.xml.rels': rels(rel('dm', '../assets/data.xml')),
    })
    const content = await shared(package_, diagram)
    expect(content?.kind).toBe('diagram')
    expect(content?.kind === 'diagram' ? content.shapes : []).toHaveLength(0)
    expect(reserveDrawingContent(package_, content!, 'owner/document.xml')).toBe(true)
    expect(drawingPartContext(package_).nodes).toBe(10000)
    expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'node-budget', reason: 'document-budget', limit: 10000 }))
  })
  test('cached capability inspection follows source order at its probe limit', async () => {
    const fallback = leaf.replace(/dsp:sp/g, 'p:sp').replace('FF0000', 'FFFF00')
    const alternate = `<mc:AlternateContent><mc:Choice Requires="a">${frame(diagram)}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    for (const [label, source, expected] of [
      ['leaf first', leaf + '<dsp:pic/>'.repeat(20001), 'choice'],
      ['leaf last', '<dsp:pic/>'.repeat(20001) + leaf, 'fallback'],
    ] as const) {
      const { doc, package_ } = await pptx(alternate, { 'assets/cached.xml': `<dsp:drawing><dsp:spTree>${source}</dsp:spTree></dsp:drawing>` })
      expect(doc.slides[0].shapes[0].source?.representation, label).toBe(expected)
      const ctx = createCanvas(100, 100).getContext('2d'); renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
      expect(pixel(ctx, 50, 50), label).toEqual(expected === 'choice' ? [255, 0, 0] : [255, 255, 0])
      if (expected === 'choice') {
        expect(drawingPartContext(package_).nodes).toBe(10000)
        expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ kind: 'unsupported-content', feature: 'pic' }))
      } else expect(drawingPartContext(package_).diagnostics).toContainEqual(expect.objectContaining({ reason: 'source-probe-limit', limit: 20000 }))
    }
  })
  test('over-depth cached groups diagnose an unselected branch and a later selected branch', async () => {
    const fallback = leaf.replace(/dsp:sp/g, 'p:sp').replace('FF0000', 'FFFF00')
    const alternate = `<mc:AlternateContent><mc:Choice Requires="a">${frame(diagram)}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    const deepest = diagramGroup(leaf).replace('<dsp:grpSp>', '<dsp:grpSp><dsp:nvGrpSpPr><dsp:cNvPr id="deepest" name="Depth boundary"/></dsp:nvGrpSpPr>')
    const chain = Array.from({ length: 64 }).reduce<string>(inner => diagramGroup(inner), deepest)
    const legacy = Array.from({ length: 64 }).reduce<string>(inner => diagramGroup(inner), diagramGroup(leaf).replace('<dsp:grpSp>', '<dsp:grpSp id="legacy">'))
    for (const [label, source, expected, identity] of [
      ['only over-depth branch', chain, 'fallback', 'deepest'],
      ['valid leaf before over-depth branch', leaf + chain, 'choice', 'deepest'],
      ['direct-attribute compatibility', legacy, 'fallback', 'legacy'],
    ] as const) {
      const { doc, package_ } = await pptx(alternate, { 'assets/cached.xml': `<dsp:drawing><dsp:spTree>${source}</dsp:spTree></dsp:drawing>` })
      expect(doc.slides[0].shapes[0].source?.representation, label).toBe(expected)
      expect(drawingPartContext(package_).diagnostics, label).toContainEqual(expect.objectContaining({ kind: 'group-depth', partPath: 'assets/cached.xml', identity, reason: 'group-depth', limit: 64 }))
    }
  })
})

const wpsNs = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
const textboxDoc = async (vert: string | undefined, paras: string) => {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ${wpsNs}><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="1905000"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${paras}</w:txbxContent></wps:txbx><wps:bodyPr${vert === undefined ? '' : ` vert="${vert}"`}/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`
  )
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}
const textboxRun = (text: string, color: string, extra = '') =>
  `<w:p><w:r><w:rPr><w:sz w:val="48"/><w:color w:val="${color}"/>${extra}</w:rPr><w:t>${text}</w:t></w:r></w:p>`
async function textboxDrawing(vert: string | undefined, paras: string) {
  const doc = await parseDocx(await textboxDoc(vert, paras))
  const { collectDocImages } = await import('../src/docx/layout')
  const drawing = collectDocImages(doc)[0]?.drawing
  if (drawing?.kind !== 'textbox') throw Error(`expected textbox, got ${drawing?.kind}`)
  return drawing
}
function paintTextboxModel(drawing: { kind: 'textbox' }, w = 200, h = 200) {
  const ctx = createCanvas(w, h).getContext('2d')
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, w, h)
  paintDrawing(drawing as never, ctx as never, w, h)
  return ctx
}
const inkIn = (ctx: ReturnType<ReturnType<typeof createCanvas>['getContext']>, w: number, h: number) => {
  const data = ctx.getImageData(0, 0, w, h).data
  const pts: Array<[number, number, number, number, number]> = []
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = (y * w + x) * 4
    if (data[i] < 128 || data[i + 1] < 128 || data[i + 2] < 128) pts.push([x, y, data[i], data[i + 1], data[i + 2]])
  }
  return pts
}

describe('word textbox shared direction routing', () => {
  test.each(['vert', 'eaVert', 'wordArtVert', 'wordArtVertRtl'] as const)('parses bodyPr vert=%s onto the model', async (vert) => {
    const drawing = await textboxDrawing(vert, textboxRun('AB', 'FF0000'))
    expect(drawing.direction).toBe(vert)
  })

  test('absent bodyPr direction stays horizontal', async () => {
    const drawing = await textboxDrawing(undefined, textboxRun('AB', 'FF0000'))
    expect(drawing.direction ?? 'horz').toBe('horz')
  })

  test('vertical textbox stacks paragraphs as right-to-left columns', async () => {
    const drawing = await textboxDrawing('vert', textboxRun('AB', 'FF0000') + textboxRun('CD', '0000FF'))
    const ctx = paintTextboxModel(drawing)
    const pts = inkIn(ctx, 200, 200)
    expect(pts.length).toBeGreaterThan(0)
    const red = pts.filter(([, , r, g, b]) => r > 200 && g < 128 && b < 128).map(([x]) => x)
    const blue = pts.filter(([, , r, g, b]) => b > 200 && r < 128 && g < 128).map(([x]) => x)
    expect(red.length).toBeGreaterThan(0)
    expect(blue.length).toBeGreaterThan(0)
    // First paragraph (red AB) sits in the rightmost column.
    expect(Math.min(...red)).toBeGreaterThan(Math.max(...blue))
  })

  test('wordArtVertRtl mirrors wordArtVert column order', async () => {
    const columnOrder = async (vert: string) => {
      const drawing = await textboxDrawing(vert, textboxRun('AB', 'FF0000') + textboxRun('CD', '0000FF'))
      const pts = inkIn(paintTextboxModel(drawing), 200, 200)
      const red = pts.filter(([, , r, g, b]) => r > 200 && g < 128 && b < 128).map(([x]) => x)
      const blue = pts.filter(([, , r, g, b]) => b > 200 && r < 128 && g < 128).map(([x]) => x)
      return Math.min(...red) - Math.max(...blue)
    }
    expect(await columnOrder('wordArtVert')).toBeLessThan(0)
    expect(await columnOrder('wordArtVertRtl')).toBeGreaterThan(0)
  })

  test('actual-paint search, copy and selection work on transformed textbox text', async () => {
    const drawing = await textboxDrawing('vert', textboxRun('AB', 'FF0000') + textboxRun('CD', '0000FF'))
    const { buildTextIndex, findMatches } = await import('../src/core/search')
    const { hitTest, textForRange } = await import('../src/core/selection')
    const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 200 }, paint: (ctx) => paintDrawing(drawing as never, ctx as never, 200, 200) }])
    const matches = findMatches(index, 'AB')
    expect(matches.length).toBeGreaterThan(0)
    expect(matches[0].rects.length).toBeGreaterThan(0)
    const rect = matches[0].rects[0]
    const hit = hitTest(index, 0, rect.x + rect.width / 2, rect.y + rect.height / 2, { strict: true })
    expect(hit).toBeDefined()
    const end = { pageIndex: 0, lineIndex: 0, charIndex: 2 }
    expect(textForRange(index, { start: { ...end, charIndex: 0 }, end })).toContain('AB')
  })

  test('underline decoration survives shared direction paint', async () => {
    const plain = await textboxDrawing('vert', textboxRun('AB', '000000'))
    const under = await textboxDrawing('vert', textboxRun('AB', '000000', '<w:u w:val="single"/>'))
    const count = (drawing: { kind: 'textbox' }) => inkIn(paintTextboxModel(drawing), 200, 200).length
    expect(count(under)).toBeGreaterThan(count(plain))
  })

  test('horizontal textbox keeps legacy single-line paint', async () => {
    const drawing = await textboxDrawing(undefined, textboxRun('AB', '000000'))
    const pts = inkIn(paintTextboxModel(drawing), 200, 200)
    expect(pts.length).toBeGreaterThan(0)
    const ys = pts.map(([, y]) => y)
    // Single horizontal line: shallow vertical band.
    expect(Math.max(...ys) - Math.min(...ys)).toBeLessThan(40)
  })
})

const labelLeaf = (bodyPr: string, runs: string) =>
  `<dsp:sp><dsp:spPr>${transform}<a:prstGeom prst="rect"/><a:noFill/></dsp:spPr><dsp:txBody><a:bodyPr${bodyPr}/><a:p><a:r>${runs}</a:r></a:p></dsp:txBody></dsp:sp>`
const labelRun = (text: string, extra = '') =>
  `<a:rPr sz="2400"${extra}/><a:t>${text}</a:t>`
async function labelDrawing(bodyPr: string, runs: string) {
  const cache = `<dsp:drawing><dsp:spTree>${labelLeaf(bodyPr, runs)}</dsp:spTree></dsp:drawing>`
  const package_ = await pkg({
    'assets/cached.xml': cache,
    'word/_rels/document.xml.rels': rels(rel('dm', '../assets/cached.xml')),
  })
  const root = parseXmlOrdered(`<w:document><w:body><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(100)}"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing></w:body></w:document>`)
  const drawings = await loadDrawingParts(package_, root, 'word/document.xml', { colors: new Map(), fonts: new Map() }, () => ({ runs: [], images: [], align: 'left' }))
  const drawing = [...drawings.values()][0]?.drawing
  if (drawing?.kind !== 'diagram') throw Error(`expected diagram, got ${drawing?.kind}`)
  return drawing.shapes[0]
}

describe('word cached labels route through shared direction layout', () => {
  test('vertical label keeps bodyPr direction, insets and anchor on the model', async () => {
    const shape = await labelDrawing(' vert="vert" lIns="0" anchor="ctr"', labelRun('AB'))
    expect(shape.textBody?.direction).toBe('vert')
    expect(shape.textBody).toMatchObject({ insetLeftEmu: 0, anchor: 'ctr' })
  })

  test('vertical cached label paints a stacked ink column', async () => {
    const shape = await labelDrawing(' vert="vert"', labelRun('AB'))
    const ctx = createCanvas(200, 100).getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 200, 100)
    paintDrawing({ kind: 'diagram', shapes: [shape] } as never, ctx as never, 200, 100)
    const data = ctx.getImageData(0, 0, 200, 100).data
    const rows: number[] = []
    for (let y = 0; y < 100; y++) for (let x = 0; x < 200; x++) {
      if (data[(y * 200 + x) * 4] < 128) { rows.push(y); break }
    }
    // Two stacked glyphs produce ink in two separated row bands.
    expect(rows.length).toBeGreaterThan(10)
    expect(rows[rows.length - 1] - rows[0]).toBeGreaterThan(20)
  })

  test('actual-paint search, copy and selection work on cached label text', async () => {
    const shape = await labelDrawing(' vert="vert"', labelRun('AB'))
    const { buildTextIndex, findMatches } = await import('../src/core/search')
    const { hitTest, textForRange } = await import('../src/core/selection')
    const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 100 }, paint: (ctx) => paintDrawing({ kind: 'diagram', shapes: [shape] } as never, ctx as never, 200, 100) }])
    const matches = findMatches(index, 'AB')
    expect(matches.length).toBeGreaterThan(0)
    expect(matches[0].rects.length).toBeGreaterThan(0)
    const rect = matches[0].rects[0]
    expect(hitTest(index, 0, rect.x + rect.width / 2, rect.y + rect.height / 2, { strict: true })).toBeDefined()
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 2 } })).toContain('AB')
  })

  test('translucent cached label alpha blends like the horizontal control', async () => {
    const paint = async (alpha: string) => {
      const shape = await labelDrawing(' vert="vert"', `<a:rPr sz="2400"><a:solidFill><a:srgbClr val="FF0000">${alpha}</a:srgbClr></a:solidFill></a:rPr><a:t>Red</a:t>`)
      const ctx = createCanvas(200, 100).getContext('2d')
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, 200, 100)
      paintDrawing({ kind: 'diagram', shapes: [shape] } as never, ctx as never, 200, 100)
      return ctx.getImageData(0, 0, 200, 100).data
    }
    const opaque = await paint('')
    const translucent = await paint('<a:alpha val="50000"/>')
    const firstInk = (data: Uint8ClampedArray) => Array.from({ length: 200 * 100 }, (_, i) => i).find(i => data[i * 4] > 200 && data[i * 4 + 1] < 150)
    const o = firstInk(opaque)!, t = firstInk(translucent)!
    expect(o).toBeDefined()
    expect(t).toBeDefined()
    // Half alpha over white lifts the green channel toward the background.
    expect(translucent[t * 4 + 1]).toBeGreaterThan(opaque[o * 4 + 1] + 40)
  })

  test('noFill cached label run paints nothing but stays searchable', async () => {
    const shape = await labelDrawing('', `<a:rPr sz="2400"><a:noFill/></a:rPr><a:t>Hidden</a:t>`)
    const ctx = createCanvas(200, 100).getContext('2d')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 200, 100)
    paintDrawing({ kind: 'diagram', shapes: [shape] } as never, ctx as never, 200, 100)
    const data = ctx.getImageData(0, 0, 200, 100).data
    let count = 0
    for (let i = 0; i < data.length; i += 4) if (data[i] < 128 || data[i + 1] < 128 || data[i + 2] < 128) count++
    expect(count).toBe(0)
    const { buildTextIndex, findMatches } = await import('../src/core/search')
    const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 100 }, paint: (ctx2) => paintDrawing({ kind: 'diagram', shapes: [shape] } as never, ctx2 as never, 200, 100) }])
    expect(findMatches(index, 'Hidden').length).toBeGreaterThan(0)
  })
})

describe('word cached labels keep Task2 font inheritance', () => {
  test('fontRef major/minor resolve with direct overrides and noFill retained', async () => {
    const label = `<dsp:sp><dsp:spPr>${transform}<a:prstGeom prst="rect"/></dsp:spPr><dsp:style><a:fontRef idx="major"><a:srgbClr val="112233"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>direct</a:t></a:r><a:r><a:rPr><a:noFill/></a:rPr><a:t>hidden</a:t></a:r><a:r><a:t>plain</a:t></a:r></a:p></dsp:txBody></dsp:sp>`
    const cache = `<dsp:drawing><dsp:spTree>${label}</dsp:spTree></dsp:drawing>`
    const package_ = await pkg({
      'assets/cached.xml': cache,
      'word/_rels/document.xml.rels': rels(rel('dm', '../assets/cached.xml')),
    })
    const root = parseXmlOrdered(`<w:document><w:body><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(100)}"/><a:graphic>${diagram}</a:graphic></wp:inline></w:drawing></w:body></w:document>`)
    const drawings = await loadDrawingParts(package_, root, 'word/document.xml', { colors: new Map(), fonts: new Map() }, () => ({ runs: [], images: [], align: 'left' }))
    const drawing = [...drawings.values()][0]?.drawing
    if (drawing?.kind !== 'diagram') throw Error(`expected diagram, got ${drawing?.kind}`)
    const runs = drawing.shapes[0].textBody?.paragraphs[0].runs
    expect(runs).toMatchObject([
      { text: 'direct', color: '#FF0000' },
      { text: 'hidden', noFill: true },
      { text: 'plain', color: '#112233' },
    ])
  })
})
