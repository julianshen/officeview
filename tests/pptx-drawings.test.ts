import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas, loadImage } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { attrs, getChildren, orderedChildren, parseXmlOrdered } from '../src/core/xml'
import { parsePptx } from '../src/pptx/parse'
import { renderSlide } from '../src/pptx/render'
import { resolveGeometry } from '../src/drawing/geometry'

const emu = (px: number) => px * 9525
const xfrm = (x = 10, y = 10, w = 80, h = 60, attributes = '') => `<a:xfrm ${attributes}><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>`
const nv = (id: number, name = `Shape ${id}`) => `<p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:nvPr/></p:nvSpPr>`
const solid = (color: string) => `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`
const shape = (id: number, properties = '', preset = 'rect', transform = xfrm(), extra = '', tag = 'sp') => `<p:${tag}>${nv(id)}<p:spPr>${transform}<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>${properties}</p:spPr>${extra}</p:${tag}>`
const picture = (id: number, rid: string, transform = xfrm(), fillRect = '', srcRect = '') => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/>${srcRect}<a:stretch><a:fillRect ${fillRect}/></a:stretch></p:blipFill><p:spPr>${transform}</p:spPr></p:pic>`
const table = (id: number) => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="95250"/></a:tblGrid><a:tr h="95250"><a:tc><a:txBody><a:p><a:r><a:t>cell</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
const group = (id: number, children: string, transform: string) => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/></p:nvGrpSpPr><p:grpSpPr>${transform}</p:grpSpPr>${children}</p:grpSp>`
const groupXfrm = (x: number, y: number, w: number, h: number, cx: number, cy: number, cw: number, ch: number, attributes = '') => xfrm(x, y, w, h, attributes).replace('</a:xfrm>', `<a:chOff x="${emu(cx)}" y="${emu(cy)}"/><a:chExt cx="${emu(cw)}" cy="${emu(ch)}"/></a:xfrm>`)
const alternate = (choice: string, fallback = '', requires = 'a14') => `<mc:AlternateContent><mc:Choice Requires="${requires}">${choice}</mc:Choice>${fallback ? `<mc:Fallback>${fallback}</mc:Fallback>` : ''}</mc:AlternateContent>`
const relations = (items: string) => `<Relationships>${items}</Relationships>`
const rel = (id: string, kind: string, target: string, external = false) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`

async function fixture(body: string, parts: Record<string, string | Uint8Array> = {}, slideExtra = '', beforeParse?: (pkg: OfficePackage) => Promise<void>) {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldSz cx="2857500" cy="1905000"/><p:sldIdLst><p:sldId r:id="s1"/></p:sldIdLst></p:presentation>')
  zip.file('ppt/_rels/presentation.xml.rels', relations(rel('s1', 'slide', 'slides/slide1.xml')))
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="r" xmlns:mc="mc" xmlns:a14="a14"><p:cSld><p:spTree>${body}</p:spTree></p:cSld>${slideExtra}</p:sld>`)
  for (const [path, data] of Object.entries(parts)) zip.file(path, data)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
  await beforeParse?.(pkg)
  return parsePptx(pkg)
}
async function render(doc: Awaited<ReturnType<typeof fixture>>) {
  const ctx = createCanvas(300, 200).getContext('2d')
  const images = await Promise.all(doc.images.map(image => loadImage(Buffer.from(image.data))))
  renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, undefined, images as unknown as CanvasImageSource[])
  return { ctx, pixel: (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data).slice(0, 3) }
}
function image(color: string) {
  const canvas = createCanvas(40, 20)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = color
  ctx.fillRect(0, 0, 40, 20)
  return new Uint8Array(canvas.toBuffer('image/png'))
}
const theme = `<a:theme><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="0000FF"/></a:accent2></a:clrScheme><a:fmtScheme><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"><a:alphaMod val="50000"/><a:alphaOff val="25000"/><a:alphaMod val="50000"/></a:schemeClr></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="dash"/></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>`
const themeParts = {
  'ppt/slides/_rels/slide1.xml.rels': relations(rel('layout', 'slideLayout', '../slideLayouts/slideLayout1.xml') + rel('external', 'theme', 'https://example.invalid/theme.xml', true)),
  'ppt/slideLayouts/slideLayout1.xml': '<p:sldLayout><p:clrMapOvr><a:overrideClrMapping accent1="accent2"/></p:clrMapOvr></p:sldLayout>',
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels': relations(rel('master', 'slideMaster', '../slideMasters/slideMaster1.xml')),
  'ppt/slideMasters/slideMaster1.xml': '<p:sldMaster><p:clrMap accent1="accent1"/></p:sldMaster>',
  'ppt/slideMasters/_rels/slideMaster1.xml.rels': relations(rel('theme', 'theme', '../theme/theme7.xml')),
  'ppt/theme/theme7.xml': theme,
}
const style = '<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></p:style>'

describe('PPTX shared static drawings', () => {
  test('preserves alternating shapes, connectors, pictures and tables in source order', async () => {
    const doc = await fixture(shape(2, solid('FF0000')) + picture(3, 'im') + shape(4, '<a:ln w="38100">' + solid('00FF00') + '</a:ln>', 'line', xfrm(10, 40, 80, 0), '', 'cxnSp') + table(5) + shape(6, solid('0000FF')), {
      'ppt/slides/_rels/slide1.xml.rels': relations(rel('im', 'image', '../media/p.png')),
      'ppt/media/p.png': image('#ffff00'),
    })
    expect(doc.slides[0].shapes.map(s => s.source?.id)).toEqual(['2', '3', '4', '5', '6'])
    expect(doc.slides[0].shapes[2].drawingStyle?.fill).toEqual({ kind: 'none' })
    expect(doc.slides[0].shapes[3].table?.rows[0].cells[0].paragraphs[0].runs[0].text).toBe('cell')
    expect((await render(doc)).pixel(50, 40)).toEqual([0, 0, 255])
  })

  test('resolves theme-only refs, nearest color maps and repeated theme transforms', async () => {
    const doc = await fixture(shape(2, '', 'rect', xfrm(), style), themeParts)
    const drawing = doc.slides[0].shapes[0].drawingStyle!
    expect(drawing.fill).toEqual({ kind: 'solid', color: { r: 0, g: 0, b: 255, a: .375 } })
    expect(drawing.line).toMatchObject({ width: 3, cap: 'round', dash: 'dash', fill: { color: { b: 255 } } })
    const pixel = (await render(doc)).pixel(50, 40)
    // Canvas quantizes .375 alpha into an 8-bit channel before compositing.
    expect(pixel[0]).toBeCloseTo(160, 0)
    expect(pixel[1]).toBeCloseTo(160, 0)
    expect(pixel[2]).toBe(255)
  })

  test('master relationships back to layouts cannot divert the slide theme chain', async () => {
    const doc = await fixture(shape(2, '', 'rect', xfrm(), style), {
      ...themeParts,
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': relations(rel('backlink', 'slideLayout', '../slideLayouts/other.xml') + rel('cycle', 'slideLayout', '../slideLayouts/slideLayout1.xml') + rel('theme', 'theme', '../theme/theme7.xml')),
      'ppt/slideLayouts/other.xml': '<p:sldLayout/>',
    })
    expect(doc.slides[0].shapes[0].drawingStyle?.fill).toMatchObject({ kind: 'solid', color: { b: 255 } })
  })

  test('explicit master color mapping resets layout colors for shape, table and background', async () => {
    const styledTable = table(3).replace('<a:tbl>', '<a:tbl><a:tblPr><a:tableStyleId>mapped</a:tableStyleId></a:tblPr>')
    const doc = await fixture(shape(2, '', 'rect', xfrm(), style) + styledTable, {
      ...themeParts,
      'ppt/slideLayouts/slideLayout1.xml': themeParts['ppt/slideLayouts/slideLayout1.xml'].replace('<p:sldLayout>', '<p:sldLayout><p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:bgPr></p:bg></p:cSld>'),
      'ppt/tableStyles.xml': '<a:tblStyleLst><a:tblStyle styleId="mapped"><a:wholeTbl><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>',
    }, '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>')
    expect(doc.slides[0].shapes[0].drawingStyle?.fill).toEqual({ kind: 'solid', color: { r: 255, g: 0, b: 0, a: .375 } })
    expect(doc.slides[0].shapes[1].table?.styleFills?.wholeTable).toBe('#FF0000')
    expect(doc.slides[0].background).toBe('#FF0000')
  })

  test('layout master mapping applies unless the slide supplies a closer override', async () => {
    const parts = { ...themeParts, 'ppt/slideLayouts/slideLayout1.xml': '<p:sldLayout><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>' }
    const masterMapped = await fixture(shape(2, '', 'rect', xfrm(), style), parts)
    const slideMapped = await fixture(shape(2, '', 'rect', xfrm(), style), parts, '<p:clrMapOvr><a:overrideClrMapping accent1="accent2"/></p:clrMapOvr>')
    expect(masterMapped.slides[0].shapes[0].drawingStyle?.fill).toMatchObject({ color: { r: 255, b: 0 } })
    expect(slideMapped.slides[0].shapes[0].drawingStyle?.fill).toMatchObject({ color: { r: 0, b: 255 } })
  })

  test('direct fill and partial line override referenced fields independently', async () => {
    const doc = await fixture(shape(2, solid('00FF00') + '<a:ln w="57150"/>', 'rect', xfrm(), style), themeParts, '<p:clrMapOvr><a:overrideClrMapping accent1="accent1"/></p:clrMapOvr>')
    expect(doc.slides[0].shapes[0].drawingStyle).toMatchObject({ fill: { color: { g: 255, a: 1 } }, line: { width: 6, cap: 'round', dash: 'dash', fill: { color: { r: 255, b: 0 } } } })
    expect((await render(doc)).pixel(50, 40)).toEqual([0, 255, 0])
  })

  test('adjusted roundRect changes its corner while custom paths keep command order', async () => {
    const adjusted = shape(2, solid('FF0000'), 'roundRect', xfrm(10, 10, 80, 80)).replace('<a:avLst/>', '<a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst>')
    const custom = `<p:sp>${nv(3)}<p:spPr>${xfrm(110, 10, 80, 80)}<a:custGeom><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:moveTo><a:pt x="100" y="100"/></a:moveTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:ln w="38100">${solid('0000FF')}</a:ln></p:spPr></p:sp>`
    const doc = await fixture(adjusted + custom)
    expect(doc.slides[0].shapes[0].drawingGeometry?.adjustments).toContainEqual(['adj', 'val 50000'])
    expect(resolveGeometry(doc.slides[0].shapes[1].drawingGeometry!, 80, 80).paths[0].commands.map(c => c[0])).toEqual(['moveTo', 'lnTo', 'moveTo', 'lnTo'])
    const { pixel } = await render(doc)
    expect(pixel(15, 15)).toEqual([255, 255, 255])
    expect(pixel(50, 50)).toEqual([255, 0, 0])
    expect(pixel(150, 11)).toEqual([0, 0, 255])
    expect(pixel(150, 50)).toEqual([255, 255, 255])
  })

  test('composes shape flips and rotation about the local extent center', async () => {
    const doc = await fixture(shape(2, solid('FF0000'), 'rtTriangle', xfrm(100, 40, 80, 40, 'flipH="1" rot="5400000"')))
    expect(doc.slides[0].shapes[0]).toMatchObject({ flipH: true, rotationDeg: 90, presetName: 'rtTriangle' })
    const { pixel } = await render(doc)
    expect(pixel(125, 40)).toEqual([255, 0, 0])
    expect(pixel(150, 40)).toEqual([255, 255, 255])
    expect(pixel(150, 85)).toEqual([255, 0, 0])
    expect(pixel(105, 50)).toEqual([255, 255, 255])
  })

  test('retains nested nonuniform group coordinates and composes chOff/chExt normalization', async () => {
    const inner = group(3, shape(4, solid('00FF00'), 'rect', xfrm(15, 10, 20, 10)), groupXfrm(20, 10, 40, 20, 5, 5, 40, 10))
    const doc = await fixture(group(2, inner, groupXfrm(100, 50, 120, 80, 10, 0, 60, 40)))
    const outer = doc.slides[0].shapes[0]
    expect(outer.children?.[0].children?.[0].xEmu).toBe(emu(15))
    expect(outer.group).toMatchObject({ chOff: { x: emu(10), y: 0 }, chExt: { width: emu(60), height: emu(40) } })
    const { pixel } = await render(doc)
    expect(pixel(145, 95)).toEqual([0, 255, 0])
    expect(pixel(125, 95)).toEqual([255, 255, 255])
    expect(pixel(180, 95)).toEqual([255, 255, 255])
  })

  test('invalid shape and group transforms are diagnosed without poisoning neighboring paints', async () => {
    const broken = shape(2, solid('FF0000'), 'rect', xfrm().replace('x="95250"', 'x="NaN"'))
    const badGroup = group(3, shape(4, solid('FF0000')), groupXfrm(0, 0, 80, 80, 0, 0, 0, 80))
    const doc = await fixture(broken + badGroup + shape(5, solid('0000FF'), 'rect', xfrm(120, 20, 40, 40)))
    expect(doc.slides[0].shapes[0].diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'invalid-transform' })]))
    expect(doc.slides[0].shapes[1].diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'invalid-transform' })]))
    const { pixel } = await render(doc)
    expect(pixel(20, 20)).toEqual([255, 255, 255])
    expect(pixel(130, 30)).toEqual([0, 0, 255])
  })

  test('unknown preset keeps its identity and text without painting a fake rectangle', async () => {
    const text = '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Still here</a:t></a:r></a:p></p:txBody>'
    const doc = await fixture(shape(2, solid('FF0000'), 'futurePreset', xfrm(), text))
    expect(doc.slides[0].shapes[0]).toMatchObject({ geometry: 'other', presetName: 'futurePreset' })
    expect(doc.slides[0].shapes[0].textBody?.paragraphs[0].runs[0].text).toBe('Still here')
    expect(doc.slides[0].shapes[0].diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'unknown-preset' })]))
    expect((await render(doc)).pixel(50, 60)).toEqual([255, 255, 255])
  })

  test('supported Choice paints once and nested AlternateContent selects once', async () => {
    const nested = alternate(shape(3, solid('00FF00')), shape(4, solid('0000FF')), 'a')
    const doc = await fixture(alternate(group(2, nested, groupXfrm(0, 0, 100, 100, 0, 0, 100, 100)), shape(5, solid('FF0000')), 'a'))
    expect(doc.slides[0].shapes).toHaveLength(1)
    expect(doc.slides[0].shapes[0].source?.representation).toBe('choice')
    expect(doc.slides[0].shapes[0].children?.map(s => s.source?.id)).toEqual(['3'])
    expect((await render(doc)).pixel(50, 40)).toEqual([0, 255, 0])
  })

  test.each([12, 40])('%i nested supported Choices are inspected linearly and paint once', async nesting => {
    let body = shape(2, solid('00FF00'))
    for (let i = 0; i < nesting; i++) body = alternate(body, '', 'a')
    let requirementsRead = 0
    const doc = await fixture(body + shape(7, solid('0000FF'), 'rect', xfrm(120, 20, 40, 40)), {}, '', async pkg => {
      // Count real branch decisions, with a deterministic guard against exponential work.
      const pending = [await pkg.xmlOrdered('ppt/slides/slide1.xml')]
      while (pending.length) {
        for (const [name, child] of orderedChildren(pending.pop())) {
          pending.push(child)
          if (name !== 'Choice') continue
          Object.defineProperty(attrs(child), 'Requires', { get() {
            if (++requirementsRead > nesting * 8) throw new Error('Compatibility traversal exceeded the linear decision budget')
            return 'a'
          } })
        }
      }
    })
    expect(requirementsRead).toBeLessThanOrEqual(nesting * 2)
    expect(doc.slides[0].shapes.map(s => s.source?.id)).toEqual(['2', '7'])
    const { pixel } = await render(doc)
    expect(pixel(50, 40)).toEqual([0, 255, 0])
    expect(pixel(130, 30)).toEqual([0, 0, 255])
  })

  test('excessively nested alternatives are diagnosed while valid neighbors survive', async () => {
    const doc = await fixture(alternate(shape(2), '', 'a') + shape(7, solid('0000FF'), 'rect', xfrm(120, 20, 40, 40)), {}, '', async pkg => {
      const root = await pkg.xmlOrdered('ppt/slides/slide1.xml')
      const tree = getChildren(getChildren(root, 'cSld')[0], 'spTree')[0]
      let nested = getChildren(tree, 'AlternateContent')[0]
      // Build an already-parsed acyclic tree beyond the XML reader's independent 100-tag limit.
      for (let i = 0; i < 256; i++) {
        const outer = parseXmlOrdered(alternate('', '', 'a'))
        const choice = getChildren(outer, 'Choice')[0]
        choice.AlternateContent = nested
        orderedChildren(choice).push(['AlternateContent', nested])
        nested = outer
      }
      tree.AlternateContent = nested
      orderedChildren(tree)[0] = ['AlternateContent', nested]
      let requirementsRead = 0
      const pending = [root]
      while (pending.length) {
        for (const [name, child] of orderedChildren(pending.pop())) {
          pending.push(child)
          if (name === 'Choice') Object.defineProperty(attrs(child), 'Requires', { get() {
            if (++requirementsRead > 2048) throw new Error('Compatibility traversal exceeded the linear decision budget')
            return 'a'
          } })
        }
      }
    })
    expect(doc.slides[0].shapes.map(s => s.source?.id)).toEqual(['7'])
    expect(doc.slides[0].diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'missing-representation', feature: 'drawing-depth' })]))
    expect((await render(doc)).pixel(130, 30)).toEqual([0, 0, 255])
  })

  test('unsupported OMML uses the native shape image fill despite noFill and honors fillRect/crop', async () => {
    const pixels = createCanvas(40, 20)
    const c = pixels.getContext('2d'); c.fillStyle = '#ff0000'; c.fillRect(0, 0, 20, 20); c.fillStyle = '#00ff00'; c.fillRect(20, 0, 20, 20)
    const omml = shape(2, '', 'rect', xfrm(), '<p:txBody><a:p><a:r><a:t>Original body</a:t></a:r><a14:m><m:oMath xmlns:m="m"/></a14:m></a:p></p:txBody>')
    const fallback = shape(2, '<a:noFill/><a:blipFill><a:blip r:embed="im"/><a:srcRect l="50000"/><a:stretch><a:fillRect l="-25000" t="-50000"/></a:stretch></a:blipFill>')
    const doc = await fixture(alternate(omml, fallback), {
      'ppt/slides/_rels/slide1.xml.rels': relations(rel('im', 'image', '../media/p.png')),
      'ppt/media/p.png': new Uint8Array(pixels.toBuffer('image/png')),
    })
    expect(doc.slides[0].shapes).toHaveLength(1)
    expect(doc.slides[0].shapes[0].source).toMatchObject({ id: '2', representation: 'fallback', feature: 'OMML' })
    expect(doc.slides[0].shapes[0].sourceTextBody?.paragraphs[0].runs[0].text).toBe('Original body')
    expect(doc.slides[0].shapes[0].image).toMatchObject({ srcRect: { l: .5, t: 0, r: 0, b: 0 }, fillRect: { l: -.25, t: -.5, r: 0, b: 0 } })
    expect(doc.slides[0].shapes[0].drawingStyle?.issues.some(i => i.feature === 'blipFill')).toBe(false)
    const { pixel } = await render(doc)
    expect(pixel(1, 1)).toEqual([0, 255, 0])
    expect(pixel(50, 40)).toEqual([0, 255, 0])
    expect(pixel(95, 40)).toEqual([255, 255, 255])
  })

  test('nested duplicate IDs retain the first source text while painting only the outer fallback', async () => {
    const textShape = (text: string, color: string) => shape(2, solid(color), 'rect', xfrm(), `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>`)
    const inner = alternate(textShape('Original source', 'FF0000'), textShape('Inner fallback', '00FF00'), 'a')
    const doc = await fixture(alternate(inner, textShape('Outer fallback', '0000FF')))
    expect(doc.slides[0].shapes).toHaveLength(1)
    const selected = doc.slides[0].shapes[0]
    expect(selected.source).toMatchObject({ id: '2', representation: 'fallback' })
    expect(selected.sourceTextBody?.paragraphs[0].runs[0].text).toBe('Original source')
    expect(selected.textBody?.paragraphs[0].runs[0].text).toBe('Outer fallback')
    const ctx = createCanvas(300, 200).getContext('2d')
    const painted: string[] = []
    const fillText = ctx.fillText.bind(ctx)
    ctx.fillText = (text, x, y) => { painted.push(text); fillText(text, x, y) }
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
    expect(painted.join('')).toBe('Outer fallback')
    expect(Array.from(ctx.getImageData(50, 60, 1, 1).data)).toEqual([0, 0, 255, 255])
  })

  test('3D and ChartEx use fallback identity with feature diagnostics and stable recursive images', async () => {
    const model = '<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d"><am3d:model3d xmlns:am3d="am3d"/></a:graphicData></a:graphic></p:graphicFrame>'
    const chart = '<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="cx"/></a:graphicData></a:graphic></p:graphicFrame>'
    const doc = await fixture(group(2, alternate(model, picture(3, 'second'), 'am3d') + picture(4, 'first'), groupXfrm(0, 0, 100, 100, 0, 0, 100, 100)) + alternate(chart, picture(5, 'second'), 'cx1'), {
      'ppt/slides/_rels/slide1.xml.rels': relations(rel('first', 'image', '../media/first.png') + rel('second', 'image', '../media/second.png')),
      'ppt/media/first.png': image('#ff0000'), 'ppt/media/second.png': image('#0000ff'),
    })
    expect(doc.slides[0].shapes[0].children?.[0].source).toMatchObject({ id: '3', representation: 'fallback', feature: 'model3D' })
    expect(doc.slides[0].shapes[1].source).toMatchObject({ id: '5', representation: 'fallback', feature: 'ChartEx' })
    expect(doc.images).toHaveLength(2)
    expect(doc.slides[0].shapes[0].children?.map(s => s.imageIndex)).toEqual([0, 1])
    expect(doc.slides[0].shapes[1].imageIndex).toBe(0)
  })

  test('records missing and unusable alternate representations', async () => {
    const doc = await fixture(alternate('<p:graphicFrame><a:graphic><a:graphicData uri="chartex"/></a:graphic></p:graphicFrame>') + alternate('<p:sp><p:spPr/><p:txBody><a:p><a14:m/></a:p></p:txBody></p:sp>', picture(2, 'missing')))
    expect(doc.slides[0].diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: 'missing-representation' }),
      expect.objectContaining({ kind: 'missing-image' }),
    ]))
  })

  test('an unusable decoded fallback image is recorded without duplicate paint', async () => {
    const doc = await fixture(alternate('<p:sp><p:spPr/><p:txBody><a:p><a14:m/></a:p></p:txBody></p:sp>', picture(2, 'im')), {
      'ppt/slides/_rels/slide1.xml.rels': relations(rel('im', 'image', '../media/broken.png')),
      'ppt/media/broken.png': new Uint8Array([0x89, 0x50]),
    })
    const ctx = createCanvas(300, 200).getContext('2d')
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, undefined, [undefined])
    expect(doc.slides[0].diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'missing-image', source: expect.objectContaining({ id: '2', representation: 'fallback' }) })]))
    expect(doc.slides[0].shapes[0].diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'missing-image' })]))
    expect(Array.from(ctx.getImageData(50, 40, 1, 1).data)).toEqual([255, 255, 255, 255])
  })

  test('nested alternatives in a shape text body keep the supported text only', async () => {
    const text = `<p:txBody><a:p>${alternate('<a:r><a:t>chosen</a:t></a:r>', '<a:r><a:t>duplicate</a:t></a:r>', 'a')}</a:p></p:txBody>`
    const doc = await fixture(shape(2, '', 'rect', xfrm(), text))
    expect(doc.slides[0].shapes[0].textBody?.paragraphs[0].runs.map(r => r.text)).toEqual(['chosen'])
  })

  test('placeholder geometry prefers layout over its master', async () => {
    const placeholder = (transform: string) => shape(2, '', 'rect', transform).replace('<p:nvPr/>', '<p:nvPr><p:ph type="title"/></p:nvPr>')
    const doc = await fixture(placeholder(''), {
      ...themeParts,
      'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout><p:cSld><p:spTree>${placeholder(xfrm(10, 20, 100, 60))}</p:spTree></p:cSld></p:sldLayout>`,
      'ppt/slideMasters/slideMaster1.xml': `<p:sldMaster><p:cSld><p:spTree>${placeholder(xfrm(50, 50, 200, 100))}</p:spTree></p:cSld></p:sldMaster>`,
    })
    expect(doc.slides[0].shapes[0]).toMatchObject({ xEmu: emu(10), yEmu: emu(20), widthEmu: emu(100), heightEmu: emu(60) })
  })
})

describe('PPTX group text uses local layout metrics', () => {
  test.each([
    { align: 'left' as const, width: 100, text: 'Hello world' },
    { align: 'left' as const, width: 100, text: 'Hello world again' },
    { align: 'center' as const, width: 200, text: 'Hello world' },
    { align: 'right' as const, width: 200, text: 'Hello world' },
  ])('$align text at $width px keeps local wrapping/alignment under scale(2,3)', ({ align, width, text }) => {
    const child = {
      xEmu: 0, yEmu: 0, widthEmu: emu(width), heightEmu: emu(100), geometry: 'rect' as const,
      textBody: { anchor: 't' as const, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, wrap: true,
        paragraphs: [{ align, level: 0, runs: [{ text, fontSizePt: 12, fontFamily: 'Arial' }] }],
      },
    }
    const grouped = { xEmu: emu(40), yEmu: emu(20), widthEmu: emu(width * 2), heightEmu: emu(300), geometry: 'other' as const,
      group: { off: { x: emu(40), y: emu(20) }, ext: { width: emu(width * 2), height: emu(300) }, chOff: { x: 0, y: 0 }, chExt: { width: emu(width), height: emu(100) } }, children: [child],
    }
    const capture = (shapes: import('../src/pptx/types').PptxShape[]) => {
      const ctx = createCanvas(700, 400).getContext('2d')
      ctx.font = '14px serif'; ctx.translate(7, 11)
      const originalTransform = ctx.getTransform(), originalFont = ctx.font
      const calls: Array<{ text: string; x: number; y: number; pageX: number; pageY: number }> = []
      const fillText = ctx.fillText.bind(ctx)
      ctx.fillText = (value, x, y) => {
        const m = ctx.getTransform()
        calls.push({ text: value, x, y, pageX: m.a * x + m.c * y + m.e, pageY: m.b * x + m.d * y + m.f })
        fillText(value, x, y)
      }
      renderSlide({ index: 0, widthEmu: emu(700), heightEmu: emu(400), shapes }, ctx as unknown as CanvasRenderingContext2D)
      expect(ctx.getTransform()).toEqual(originalTransform)
      expect(ctx.font).toBe(originalFont)
      return calls
    }
    const normal = capture([child]), scaled = capture([grouped])
    expect(normal.map(call => call.text)).toEqual(text === 'Hello world' ? ['Hello world'] : ['Hello world ', 'again'])
    expect(scaled.map(call => ({ text: call.text, x: call.x, y: call.y }))).toEqual(normal.map(call => ({ text: call.text, x: call.x, y: call.y })))
    const identity = createCanvas(4, 4).getContext('2d'); identity.font = '12pt "Arial"'
    const expectedX = align === 'center' ? (width - identity.measureText('Hello world').width) / 2 : align === 'right' ? width - identity.measureText('Hello world').width : 0
    expect(normal[0].x).toBeCloseTo(expectedX)
    expect(scaled[0].pageX).toBeCloseTo(47 + 2 * expectedX)
    expect(normal[0].y).toBeGreaterThan(0)
    expect(normal[0].y).toBeLessThan(16)
    expect(scaled[0].pageY).toBeCloseTo(31 + 3 * normal[0].y)
  })
})
