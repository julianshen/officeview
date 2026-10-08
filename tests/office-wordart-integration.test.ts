import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { renderSlide } from '../src/pptx/render'


/** Test-only OffscreenCanvas host for the WordArt warp raster painter: the
 * warp pipeline resolves gradients/outlines/transforms and fillText INSIDE the
 * shaped run raster (not on the caller's context), so this observer proxies
 * EVERY non-8x8 scratch surface's 2d context to record the real paint calls. */
const rasterObs = (): { grads: Array<{ offset: number; color: string }>; strokes: string[]; fills: string[]; transforms: number[][]; restore: () => void } => {
  const obs = { grads: [] as Array<{ offset: number; color: string }>, strokes: [] as string[], fills: [] as string[], transforms: [] as number[][] }
  const realCreateCanvas = createCanvas as unknown as (w: number, h: number) => any
  ;(globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = function (this: unknown, w: number, h: number) {
    const surface = realCreateCanvas.call(null, w, h)
    if (w !== 8 || h !== 8) {
      const original = surface.getContext.bind(surface)
      surface.getContext = (kind: string): unknown => {
        const ctx = original(kind) as any
        return new Proxy(ctx, {
          get(target, prop) {
            if (prop === 'fillText') return (text: string, ...rest: unknown[]) => { obs.fills.push(String(text)); return target.fillText(text, ...rest) }
            if (prop === 'strokeText') return (text: string, ...rest: unknown[]) => { obs.strokes.push(String(text)); return target.strokeText(text, ...rest) }
            if (prop === 'createLinearGradient') return (...args: unknown[]) => {
              const grad = target.createLinearGradient(...(args as [number, number, number, number]))
              const origAdd = grad.addColorStop.bind(grad)
              grad.addColorStop = (offset: number, color: string) => { obs.grads.push({ offset, color }); return origAdd(offset, color) }
              return grad
            }
            if (prop === 'strokeText') return (text: string, ...rest: unknown[]) => { obs.strokes.push(String(text)); return target.strokeText(text, ...rest) }
            if (prop === 'transform') return (a: number, b: number, c: number, d: number, e: number, f: number) => { obs.transforms.push([a, b, c, d, e, f]); return target.transform(a, b, c, d, e, f) }
            const value = target[prop]
            return typeof value === 'function' ? value.bind(target) : value
          },
          set(target, prop, value) { target[prop as string] = value; return true },
        })
      }
    }
    return surface
  } as unknown as typeof OffscreenCanvas
  return {
    ...obs,
    restore() { delete (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas },
  }
}


interface PaintPixels { data: Uint8ClampedArray; width: number; height: number }
/** Painted-pixel evidence helpers for the shaped-raster warp pipeline. */
function pixelDominator(p: { r: number; g: number; b: number; a: number }): 'red' | 'green' | 'blue' | 'strokeGreen' | null {
  if (p.a < 60 || (p.r > 235 && p.g > 235 && p.b > 235)) return null
  // Dominance with antialiasing headroom: the candidate channel must EXCEED the
  // others by the margin (neutral grays never classify as a hue).
  const M = 1.12
  if (p.b > 30 && p.b > p.r * M && p.b > p.g * M) return 'blue'
  if (p.r > 30 && p.r > p.g * M && p.r > p.b * M) return 'red'
  if (p.g > 30 && p.g > p.r * 1.35 && p.g > p.b * M) return 'green'
  if (p.g > 30 && p.g > p.r * 1.05 && p.g > p.b * 1.05) return 'strokeGreen'
  return null
}
type PaintDrawingFn = (drawing: unknown, ctx: CanvasRenderingContext2D | unknown, width: number, height: number, assets?: unknown) => void
/** Paint a docx drawing into a small canvas and return pixel statistics. */
function pixelPaint(paint: PaintDrawingFn, drawing: any, width: number, height: number) {
  const cv = createCanvas(width, height)
  const c = cv.getContext('2d')
  c.fillStyle = '#ffffff'
  c.fillRect(0, 0, width, height)
  paint(drawing, c as unknown as CanvasRenderingContext2D, width, height)
  return paintStats({ data: c.getImageData(0, 0, width, height).data, width, height })
}

function paintStats(img: PaintPixels): { ink: number; red: number; green: number; blue: number; strokeGreen: number; redMeanY: number; greenMeanY: number; bbox: { x0: number; y0: number; x1: number; y1: number }; rows: Array<{ r: number; b: number }> } {
  let ink = 0, red = 0, green = 0, blue = 0, strokeGreen = 0
  let redY = 0, greenY = 0
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
  for (let y = 0; y < img.height; y++) {
    for (let x = 0; x < img.width; x++) {
      const i = (y * img.width + x) * 4
      const a = img.data[i + 3]
      const r = img.data[i], g = img.data[i + 1], b = img.data[i + 2]
      // Bounding box from NON-BACKGROUND ink only: a white-filled canvas has
      // alpha=255 everywhere, so classify by dominance rather than alpha.
      const classified = pixelDominator({ r, g, b, a }) || (a > 40 && (r < 230 || g < 230 || b < 230) ? 'ink' : null)
      if (a > 40 && classified !== null && (classified !== 'ink' || !(r > 235 && g > 235 && b > 235))) {
        ink++
        x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y)
      }
      const kind = pixelDominator({ r, g, b, a })
      if (kind === 'red') { red++; redY += y }
      else if (kind === 'green') { green++; greenY += y }
      else if (kind === 'blue') { blue++ }
      else if (kind === 'strokeGreen') { strokeGreen++ }
    }
  }
  // Per-ROW and per-COLUMN mean channels across non-background ink pixels.
  // Under an outer affine frame (vert270 etc.) the run's advance axis maps to
  // the ROW axis, so a faithful directional-gradient walk is measured per row;
  // unrotated runs read per column. Rows with <8 ink pixels are noise-skipped.
  const rowRed = new Array(img.height).fill(0), rowBlue = new Array(img.height).fill(0)
  const rowCount = new Array(img.height).fill(0)
  for (let y = 0; y < img.height; y++) {
    for (let x = 0; x < img.width; x++) {
      const i = (y * img.width + x) * 4
      const r = img.data[i], g = img.data[i + 1], b = img.data[i + 2], a = img.data[i + 3]
      if (a > 60 && !(r > 235 && g > 235 && b > 235)) {
        rowCount[y]++; rowRed[y] += r; rowBlue[y] += b
      }
    }
  }
  const rows: Array<{ r: number; b: number }> = []
  for (let y = 0; y < img.height; y++) {
    if (rowCount[y] >= 8) rows.push({ r: rowRed[y] / rowCount[y], b: rowBlue[y] / rowCount[y] })
  }
  return {
    ink, red, green, blue, strokeGreen,
    redMeanY: red ? redY / red : -1, greenMeanY: green ? greenY / green : -1,
    bbox: { x0: x0 === Infinity ? 0 : x0, y0: y0 === Infinity ? 0 : y0, x1, y1 },
    rows,
  }
}

const emu = (px: number) => px * 9525
const rels = (items: string) => `<Relationships>${items}</Relationships>`
const rel = (id: string, kind: string, target: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${target}"/>`

const THEME_XML = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="111111"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="FF2200"/></a:accent1>
      <a:accent2><a:srgbClr val="0033CC"/></a:accent2>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`

async function createPptxFixture(spBody: string) {
  const zip = new JSZip()
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="2857500" cy="1905000"/><p:sldIdLst><p:sldId r:id="s1"/></p:sldIdLst></p:presentation>',
  )
  zip.file('ppt/_rels/presentation.xml.rels', rels(rel('s1', 'slide', 'slides/slide1.xml')))
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>${spBody}</p:spTree></p:cSld></p:sld>`,
  )
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    rels(rel('layout', 'slideLayout', '../slideLayouts/slideLayout1.xml')),
  )
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:clrMapOvr><a:masterClrMapping xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></p:clrMapOvr></p:sldLayout>',
  )
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    rels(rel('master', 'slideMaster', '../slideMasters/slideMaster1.xml')),
  )
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    '<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:clrMap xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"/></p:sldMaster>',
  )
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    rels(rel('theme', 'theme', '../theme/theme1.xml')),
  )
  zip.file('ppt/theme/theme1.xml', THEME_XML)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
  return parsePptx(pkg)
}

describe('Phase 3: Format Adapters Integration', () => {
  test('PPTX shapes parse and render WordArt text runs inheriting theme colors', async () => {
    const sp = `<p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="WordArt Shape"/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="${emu(10)}" y="${emu(10)}"/><a:ext cx="${emu(200)}" cy="${emu(80)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr sz="2400">
              <a:gradFill>
                <a:gsLst>
                  <a:gs pos="0"><a:schemeClr val="accent1"/></a:gs>
                  <a:gs pos="100000"><a:schemeClr val="accent2"/></a:gs>
                </a:gsLst>
                <a:lin ang="5400000"/>
              </a:gradFill>
              <a:ln w="25400">
                <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
              </a:ln>
              <a:effectLst>
                <a:outerShdw blurRad="38100" dist="25400" dir="5400000">
                  <a:schemeClr val="dk1"/>
                </a:outerShdw>
              </a:effectLst>
            </a:rPr>
            <a:t>Theme WordArt</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:sp>`

    const doc = await createPptxFixture(sp)
    const shape = doc.slides[0].shapes[0]
    expect(shape).toBeDefined()
    expect(shape.textBody).toBeDefined()

    const run = shape.textBody!.paragraphs[0].runs[0]
    expect(run.text).toBe('Theme WordArt')
    // Theme scheme colors must be resolved:
    expect(run.textFill).toEqual({
      kind: 'gradient',
      angle: Math.PI / 2,
      stops: [
        { position: 0, color: '#FF2200' },
        { position: 1, color: '#0033CC' },
      ],
    })
    expect(run.textOutline).toBeDefined()
    expect(run.textOutline?.color).toBe('#FF2200')
    expect(run.textOutline?.widthPx).toBeCloseTo(2.67, 1)

    expect(run.textShadow).toBeDefined()
    expect(run.textShadow?.color).toBe('#111111')
    expect(run.textShadow?.blurPx).toBeCloseTo(4, 1)

    // Render onto canvas and verify stroke and fill executions
    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    const gradientStops: Array<{ offset: number; color: string }> = []
    const strokes: Array<{ text: string; lineWidth: number; strokeStyle: string }> = []
    const fills: Array<{ text: string; shadowBlur: number; shadowColor: string }> = []

    const origCreateLinearGradient = ctx.createLinearGradient.bind(ctx)
    ctx.createLinearGradient = (x0, y0, x1, y1) => {
      const grad = origCreateLinearGradient(x0, y0, x1, y1)
      const origAddColorStop = grad.addColorStop.bind(grad)
      grad.addColorStop = (offset, color) => {
        gradientStops.push({ offset, color })
        return origAddColorStop(offset, color)
      }
      return grad
    }

    const origStrokeText = ctx.strokeText.bind(ctx)
    ctx.strokeText = (text, x, y) => {
      strokes.push({ text: String(text), lineWidth: ctx.lineWidth, strokeStyle: String(ctx.strokeStyle) })
      return origStrokeText(text, x, y)
    }

    const origFillText = ctx.fillText.bind(ctx)
    ctx.fillText = (text, x, y) => {
      fills.push({ text: String(text), shadowBlur: ctx.shadowBlur, shadowColor: String(ctx.shadowColor) })
      return origFillText(text, x, y)
    }

    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)

    // Gradient was constructed with resolved theme colors for both segments:
    expect(gradientStops).toEqual([
      { offset: 0, color: '#FF2200' },
      { offset: 1, color: '#0033CC' },
      { offset: 0, color: '#FF2200' },
      { offset: 1, color: '#0033CC' },
    ])
    // Outline stroke was called with resolved theme outline:
    expect(strokes.map(s => s.text)).toEqual(['Theme ', 'WordArt'])
    expect(strokes[0].strokeStyle).toBe('#ff2200')
    // Fill text was called with shadow parameters:
    expect(fills.map(f => f.text)).toEqual(['Theme ', 'WordArt'])
    expect(fills[0].shadowColor).toBe('#111111')
    expect(fills[0].shadowBlur).toBeCloseTo(4, 1)
  })

  test('DOCX DrawingML shapes (<wps:wsp>) parse and render WordArt gradient and outline text', async () => {
    const zip = new JSZip()
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body>
        <w:p>
          <w:r>
            <w:drawing>
              <wp:inline>
                <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
                <wp:docPr id="1" name="WordArt Shape 1"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                    <wps:wsp>
                      <wps:cNvSpPr txBox="1"/>
                      <wps:spPr>
                        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm>
                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                        <a:noFill/><a:ln><a:noFill/></a:ln>
                      </wps:spPr>
                      <wps:txbx id="1">
                        <w:txbxContent>
                          <w:p>
                            <w:r>
                              <w:rPr>
                                <w:sz w:val="48"/>
                                <w14:textFill>
                                  <w14:gradFill>
                                    <w14:gsLst>
                                      <w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs>
                                      <w14:gs w14:pos="100000"><w14:srgbClr w14:val="00FF00"/></w14:gs>
                                    </w14:gsLst>
                                    <w14:lin w14:ang="5400000" w14:scaled="1"/>
                                  </w14:gradFill>
                                </w14:textFill>
                                <w14:textOutline w14:w="25400">
                                  <w14:solidFill><w14:srgbClr w14:val="0000FF"/></w14:solidFill>
                                  <w14:round/>
                                </w14:textOutline>
                              </w:rPr>
                              <w:t>Docx WordArt</w:t>
                            </w:r>
                          </w:p>
                        </w:txbxContent>
                      </wps:txbx>
                      <wps:bodyPr wrap="none" anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0">
                        <a:prstTxWarp prst="textWave1">
                          <a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst>
                        </a:prstTxWarp>
                        <a:noAutofit/>
                      </wps:bodyPr>
                    </wps:wsp>
                  </a:graphicData>
                </a:graphic>
              </wp:inline>
            </w:drawing>
          </w:r>
        </w:p>
      </w:body>
    </w:document>`

    zip.file('word/document.xml', docxXml)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const doc = await parseDocx(pkg)

    const para = doc.sections[0].paragraphs[0]
    expect(para.images.length).toBe(1)
    const img = para.images[0]
    expect(img.drawing).toBeDefined()
    expect(img.drawing?.kind).toBe('textbox')
    if (img.drawing?.kind !== 'textbox') throw new Error('Expected textbox drawing')

    const p = img.drawing.paragraphs[0]
    expect(p).toBeDefined()
    const run = p.runs[0]
    expect(run.text).toBe('Docx WordArt')
    expect(run.textFill).toEqual({
      kind: 'gradient',
      angle: Math.PI / 2,
      stops: [
        { position: 0, color: '#FF0000' },
        { position: 1, color: '#00FF00' },
      ],
    })
    expect(run.textOutline).toBeDefined()
    expect(run.textOutline?.color).toBe('#0000FF')
    expect(img.drawing.textWarp?.preset).toBe('textWave1')

    // Render canvas
    const canvas = createCanvas(200, 100)
    const ctx = canvas.getContext('2d')
    const strokes: string[] = []
    const gradStops: Array<{ offset: number; color: string }> = []

    const origCreateGrad = ctx.createLinearGradient.bind(ctx)
    ctx.createLinearGradient = (x0, y0, x1, y1) => {
      const grad = origCreateGrad(x0, y0, x1, y1)
      const origAdd = grad.addColorStop.bind(grad)
      grad.addColorStop = (o, c) => {
        gradStops.push({ offset: o, color: c })
        return origAdd(o, c)
      }
      return grad
    }
    const origStroke = ctx.strokeText.bind(ctx)
    ctx.strokeText = (t, x, y) => {
      strokes.push(String(t))
      return origStroke(t, x, y)
    }

    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, 200, 100)
    // REAL PAINTED PIXELS plus staged source evidence. This original fixture
    // (normal 24pt, 2pt blue outline, red→green fill, Wave) is a STROKE-
    // OCCLUSION CONTROL: the 2pt centered stroke legitimately covers the
    // thin 24pt stems, so final red/green fill pixels are near-zero AND THAT
    // IS VALID. Matched-case evidence on this same 24pt fixture: source
    // immediately after fill holds red729/green871 (both gradient stops
    // present); after the blue stroke it holds red6/green7/blue3036, with
    // near-occlusion in the final image. It must not be read as gradient
    // loss. Both stops are required as separate genuine populations below —
    // a warm aggregate alone cannot catch a missing green stop.
    const observer = rasterObs()
    paintDrawing(img.drawing, ctx as unknown as CanvasRenderingContext2D, 200, 100)
    observer.restore()
    expect(observer.grads.length).toBeGreaterThanOrEqual(2)
    expect(observer.grads[0].color).toBe('#FF0000')
    expect(observer.grads[1].color).toBe('#00FF00')
    expect(observer.strokes.length).toBeGreaterThan(0)
    expect(observer.fills.length).toBeGreaterThan(0)
    // Fill-before-stroke: hook the shaped source raster and count warm fill
    // ink right after its fillText (before the outline stroke occludes it).
    // Strong warm ink here + blue-dominated final proves VALID OCCLUSION,
    // not a gradient/frame/opacity regression in the warp path. Both stops
    // are additionally required as SEPARATE red and green populations: a
    // fault redirecting the green stop to red keeps warm ink while green
    // drops to zero, so the aggregate alone must not pass.
    const { createCanvas: rasterCanvas } = await import('canvas')
    const realCreate = rasterCanvas as unknown as (w: number, h: number) => any
    let fillWarm = 0, fillRed = 0, fillGreen = 0
    ;(globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas = function (this: unknown, w: number, h: number) {
      const surface = realCreate.call(null, w, h)
      if (w !== 8 || h !== 8) {
        const original = surface.getContext.bind(surface)
        surface.getContext = (kind: string): unknown => {
          const rc = original(kind) as any
          return new Proxy(rc, {
            get(target, prop) {
              if (prop === 'fillText') return (text: string, ...rest: unknown[]) => {
                const result = target.fillText(text, ...rest)
                try {
                  const d = surface.getContext('2d').getImageData(0, 0, surface.width, surface.height).data
                  let warm = 0, red = 0, green = 0
                  for (let i = 0; i < d.length; i += 4) {
                    if (d[i + 3] > 60 && d[i] > d[i + 2] + 30) warm++
                    if (d[i + 3] > 60 && d[i] > 30 && d[i] > d[i + 1] * 1.12 && d[i] > d[i + 2] * 1.12) red++
                    if (d[i + 3] > 60 && d[i + 1] > 30 && d[i + 1] > d[i] * 1.12 && d[i + 1] > d[i + 2] * 1.12) green++
                  }
                  fillWarm = Math.max(fillWarm, warm)
                  fillRed = Math.max(fillRed, red)
                  fillGreen = Math.max(fillGreen, green)
                } catch { /* noop */ }
                return result
              }
              const value = target[prop]
              return typeof value === 'function' ? value.bind(target) : value
            },
            set(target, prop, value) { target[prop as string] = value; return true },
          })
        }
      }
      return surface
    } as unknown as typeof OffscreenCanvas
    try {
      const canvas2 = createCanvas(200, 100)
      const ctx2 = canvas2.getContext('2d')
      ctx2.fillStyle = '#ffffff'; ctx2.fillRect(0, 0, 200, 100)
      paintDrawing(img.drawing, ctx2 as unknown as CanvasRenderingContext2D, 200, 100)
    } finally {
      delete (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas
    }
    expect(fillWarm, 'shaped source raster holds warm gradient fill before stroke').toBeGreaterThan(500)
    expect(fillRed, 'shaped source raster holds genuine red stop ink before stroke').toBeGreaterThan(200)
    expect(fillGreen, 'shaped source raster holds genuine green stop ink before stroke').toBeGreaterThan(200)
    const stats = paintStats({ data: ctx.getImageData(0, 0, 200, 100).data, width: 200, height: 100 })
    expect(stats.blue, 'painted #0000FF outline pixels exist').toBeGreaterThan(500)
    expect(stats.ink, 'warped outline ink painted').toBeGreaterThan(500)
  })

  test('normal24 gradient negative control: a redirected green stop fails the dual-stop gate', async () => {
    // Same valid normal24pt/red→green/2pt-blue-outline/Wave package as above
    // (compact equivalent XML). A host fault redirects the actual green stop
    // to red downstream of the authored model and stop-call trace: the old
    // warm aggregate and the blue outline still pass, but the genuine green
    // source population must drop to zero and the dual-stop gate must fail.
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(100)}"/><wp:docPr id="1" name="WordArt Shape 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx id="1"><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="48"/><w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs><w14:gs w14:pos="100000"><w14:srgbClr w14:val="00FF00"/></w14:gs></w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="1"/></w14:gradFill></w14:textFill><w14:textOutline w14:w="25400"><w14:solidFill><w14:srgbClr w14:val="0000FF"/></w14:solidFill><w14:round/></w14:textOutline></w:rPr><w:t>Docx WordArt</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr wrap="none" anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"><a:prstTxWarp prst="textWave1"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstTxWarp><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`
    const zip = new JSZip()
    zip.file('word/document.xml', docxXml)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    if (drawing.kind !== 'textbox') throw new Error('Expected textbox drawing')
    const { createCanvas: faultCanvas } = await import('canvas')
    const realCreate = faultCanvas as unknown as (w: number, h: number) => any
    const NodeCanvas = realCreate(1, 1).constructor as unknown as { prototype: { getContext: (...args: unknown[]) => unknown } }
    const originalGetContext = NodeCanvas.prototype.getContext
    const seen = new WeakSet<object>()
    const stops: Array<{ at: number; color: string }> = []
    let warm = 0, sourceGreen = 0, sourceRed = 0
    ;(NodeCanvas.prototype as any).getContext = function (kind: string, ...args: unknown[]) {
      const target = originalGetContext.call(this, kind, ...args) as any
      if (!target || seen.has(target)) return target
      seen.add(target)
      const createGradient = target.createLinearGradient.bind(target)
      target.createLinearGradient = (...gargs: unknown[]) => {
        const grad = createGradient(...(gargs as [number, number, number, number]))
        const add = grad.addColorStop.bind(grad)
        grad.addColorStop = (at: number, color: string) => {
          stops.push({ at, color })
          return add(at, String(color).toUpperCase() === '#00FF00' ? '#FF0000' : color)
        }
        return grad
      }
      const fillText = target.fillText.bind(target)
      target.fillText = (...fargs: unknown[]) => {
        const out = fillText(...(fargs as [string, number, number]))
        try {
          if ((this as unknown as { width: number }).width !== 200) {
            const d = target.getImageData(0, 0, target.canvas.width, target.canvas.height).data
            let w = 0, g = 0, r = 0
            for (let i = 0; i < d.length; i += 4) {
              if (d[i + 3] <= 60) continue
              if (d[i] > d[i + 2] + 30) w++
              if (d[i + 1] > 30 && d[i + 1] > d[i] * 1.12 && d[i + 1] > d[i + 2] * 1.12) g++
              if (d[i] > 30 && d[i] > d[i + 1] * 1.12 && d[i] > d[i + 2] * 1.12) r++
            }
            warm = Math.max(warm, w); sourceGreen = Math.max(sourceGreen, g); sourceRed = Math.max(sourceRed, r)
          }
        } catch { /* noop */ }
        return out
      }
      return target
    }
    try {
      const canvas = faultCanvas(200, 100)
      const ctx = canvas.getContext('2d')
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, 200, 100)
      paintDrawing(drawing, ctx as unknown as CanvasRenderingContext2D, 200, 100)
      const dd = ctx.getImageData(0, 0, 200, 100).data
      let finalBlue = 0, finalGreen = 0
      for (let i = 0; i < dd.length; i += 4) {
        if (dd[i + 3] < 60) continue
        if (dd[i + 2] > 30 && dd[i + 2] > dd[i] * 1.12 && dd[i + 2] > dd[i + 1] * 1.12) finalBlue++
        if (dd[i + 1] > 30 && dd[i + 1] > dd[i] * 1.12 && dd[i + 1] > dd[i + 2] * 1.12) finalGreen++
      }
      // Authored model and stop-call trace are preserved by the fault.
      expect(stops.map(s => s.color)).toEqual(['#FF0000', '#00FF00'])
      // Old aggregate + outline gates still pass under the fault (the gap).
      expect(warm, 'fault keeps warm aggregate passing').toBeGreaterThan(500)
      expect(finalBlue, 'fault keeps blue outline passing').toBeGreaterThan(500)
      // The dual-stop gate has teeth: green vanishes, gate evaluates false.
      expect(sourceRed, 'fault keeps red stop').toBeGreaterThan(200)
      expect(sourceGreen, 'fault removes green stop').toBe(0)
      expect(finalGreen, 'fault removes final green').toBe(0)
      expect(sourceRed > 200 && sourceGreen > 200, 'dual-stop gate fails under the fault').toBe(false)
    } finally {
      ;(NodeCanvas.prototype as any).getContext = originalGetContext
    }
  })

  test('regression gate: plain textboxes preserve resolved styles and render without WordArt appearance', async () => {
    const zip = new JSZip()
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body>
        <w:p>
          <w:r>
            <w:drawing>
              <wp:inline>
                <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
                <wp:docPr id="2" name="Plain Box"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                    <wps:wsp>
                      <wps:cNvSpPr txBox="1"/>
                      <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
                      <wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="24"/><w:color w:val="112233"/></w:rPr><w:t>Plain Text</w:t></w:r></w:p></w:txbxContent></wps:txbx>
                      <wps:bodyPr wrap="none" anchor="ctr"/>
                    </wps:wsp>
                  </a:graphicData>
                </a:graphic>
              </wp:inline>
            </w:drawing>
          </w:r>
        </w:p>
      </w:body>
    </w:document>`
    zip.file('word/document.xml', docxXml)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    expect(drawing.kind).toBe('textbox')
    if (drawing.kind !== 'textbox') throw new Error('Expected textbox')
    const run = drawing.paragraphs[0].runs[0]
    expect(run.text).toBe('Plain Text')
    expect(run.color).toBe('112233')
    expect(run.textFill).toBeUndefined()
    expect(run.textOutline).toBeUndefined()
    expect(drawing.textWarp).toBeUndefined()

    const canvas = createCanvas(200, 100)
    const ctx = canvas.getContext('2d')
    let gradientCount = 0, strokeCount = 0
    ctx.createLinearGradient = () => { gradientCount++; return { addColorStop: () => {} } as any }
    ctx.strokeText = () => { strokeCount++ }
    paintDrawing(drawing, ctx as unknown as CanvasRenderingContext2D, 200, 100)
    expect(gradientCount).toBe(0)
    expect(strokeCount).toBe(0)
  })

  test('regression gate: vertical textboxes (vert270) compose with layout and paint', async () => {
    // Authored XML pair: identical 12pt vert270 textWave1 packages differing
    // only by <w:b/>. Both go through real parsing (no model mutation), so
    // paired fill/outline evidence reflects the actual pipeline.
    const buildVertDocx = (bold: boolean): string => `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
        <wp:docPr id="3" name="Vertical Box"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
            <wps:txbx><w:txbxContent><w:p><w:r><w:rPr>${bold ? '<w:b/>' : ''}<w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs><w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs></w14:gsLst><w14:lin w14:ang="0"/></w14:gradFill></w14:textFill><w14:textOutline w14:w="12700"><w14:solidFill><w14:srgbClr w14:val="008800"/></w14:solidFill></w14:textOutline></w:rPr><w:t>Vertical Art</w:t></w:r></w:p></w:txbxContent></wps:txbx>
            <wps:bodyPr vert="vert270" wrap="none" anchor="ctr">
              <a:prstTxWarp prst="textWave1"/>
            </wps:bodyPr>
          </wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body>
    </w:document>`
    const loadVertDrawing = async (bold: boolean) => {
      const zip = new JSZip()
      zip.file('word/document.xml', buildVertDocx(bold))
      zip.file('word/_rels/document.xml.rels', '<Relationships/>')
      const { parseDocx } = await import('../src/docx/parse')
      const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
      const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
      if (drawing.kind !== 'textbox') throw new Error('Expected textbox')
      return drawing
    }
    const { paintDrawing } = await import('../src/docx/drawing')
    const { RECORD_TEXT } = await import('../src/core/text-recording')
    const drawing = await loadVertDrawing(false)
    expect(drawing.kind).toBe('textbox')
    if (drawing.kind !== 'textbox') throw new Error('Expected textbox')
    expect(drawing.direction).toBe('vert270')
    expect(drawing.textWarp?.preset).toBe('textWave1')
    const run = drawing.paragraphs[0].runs[0]
    expect(run.text).toBe('Vertical Art')
    expect(run.bold ?? false).toBe(false)

    // Visual paint observation: vert270 rotation composes with the local warp
    // (destinations map through the member affine), gradient + stroke resolve
    // inside the shaped source raster. Rotation composition is verified
    // GEOMETRICALLY below (vertical band + advance-mapped rows), not by
    // sniffing raster transform calls: the local-frame design composes the
    // member affine into destination coordinates directly.
    const canvas = createCanvas(200, 100)
    const ctx = canvas.getContext('2d') as any
    const observer = rasterObs()
    paintDrawing(drawing, ctx, 200, 100)
    observer.restore()
    expect(observer.fills.length).toBeGreaterThan(0)
    expect(observer.grads.length).toBeGreaterThan(0)
    expect(observer.strokes.length).toBeGreaterThan(0)
    // Painted-pixel evidence on the ORIGINAL REGULAR fixture, kept as a
    // STROKE-OCCLUSION/BASELINE control WITHOUT brittle exact-zero RGB
    // assertions (antialiasing across fonts/rasters is not a fixed contract):
    // vertical band geometry + outline display remain hard criteria here, and
    // gradient-stop VISIBILITY is asserted via the AUTHOR-explicit bold pair
    // below (fill must strengthen, stroke stays).
    const paint = paintDrawing as unknown as PaintDrawingFn
    const occlude = pixelPaint(paint, drawing, 200, 100)
    expect(occlude.ink, 'vert270 regular ink painted').toBeGreaterThan(60)
    const ocTall = occlude.bbox.y1 - occlude.bbox.y0, ocWide = occlude.bbox.x1 - occlude.bbox.x0
    expect(ocTall, 'vert270 regular ink occupies a vertical band (rotated column)').toBeGreaterThan(ocWide)
    expect(ocTall + ocWide, 'vert270 regular ink spans the canvas column').toBeGreaterThan(80)
    expect(occlude.green, 'REGULAR CONTROL: #008800 outline displays').toBeGreaterThan(0)

    // PAIRED BOLD package: identical XML except authored <w:b/>. Bold runs are
    // wider so the red→blue gradient fill shows past the green stroke:
    // both gradient stops appear while the outline still displays.
    const boldDrawing = await loadVertDrawing(true)
    expect(boldDrawing.kind).toBe('textbox')
    if (boldDrawing.kind !== 'textbox') throw new Error('Expected textbox')
    expect(boldDrawing.paragraphs[0].runs[0].bold).toBe(true)
    const bold = pixelPaint(paint, boldDrawing, 200, 100)
    expect(bold.red, 'BOLD: painted red gradient-stop ink exists (fill visible)').toBeGreaterThan(40)
    expect(bold.blue, 'BOLD: painted blue gradient-stop ink exists').toBeGreaterThan(40)
    expect(bold.green, 'BOLD: #008800 outline still displays').toBeGreaterThan(100)
    expect(bold.ink, 'BOLD pixels painted').toBeGreaterThan(400)
    // Directional gradient walk under the outer vert270 affine — the run
    // advance maps onto page ROWS: red weight wins at the mapped START
    // (page bottom), blue weight at the mapped tail (page top).
    const rows = bold.rows
    expect(rows.length, 'gradient samples across mapped ink rows').toBeGreaterThanOrEqual(3)
    expect(rows[rows.length - 1].r, 'run-advance start carries more red').toBeGreaterThanOrEqual(rows[0].r - 40)
    expect(rows[0].b, 'run-advance tail carries more blue').toBeGreaterThanOrEqual(rows[rows.length - 1].b - 40)

    // RECORD_TEXT observation: once-only source text recording
    const recCanvas = createCanvas(200, 100)
    const recCtx = recCanvas.getContext('2d') as any
    const recorded: any[] = []
    recCtx[RECORD_TEXT] = (text: string, _x: number, _y: number, _w: number, logical: any) => {
      recorded.push({ text, start: logical.start, end: logical.end, source: logical.source.text })
    }
    paintDrawing(drawing, recCtx, 200, 100)
    expect(recorded.length).toBeGreaterThan(0)
    const sources = Array.from(new Set(recorded.map(r => r.source)))
    expect(sources).toEqual(['Vertical Art'])
    const combined = recorded.map(r => r.source.slice(r.start, r.end)).join('')
    expect(combined).toBe('Vertical Art')
  })

  test('regression gate: once-only logical text recording with RECORD_TEXT', async () => {
    const zip = new JSZip()
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
        <wp:docPr id="4" name="Recorded Box"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
            <wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w14:textOutline w14:w="12700"><w14:solidFill><w14:srgbClr w14:val="008800"/></w14:solidFill></w14:textOutline></w:rPr><w:t>Recorded Art</w:t></w:r></w:p></w:txbxContent></wps:txbx>
            <wps:bodyPr wrap="none" anchor="ctr"><a:prstTxWarp prst="textWave1"/></wps:bodyPr>
          </wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body>
    </w:document>`
    zip.file('word/document.xml', docxXml)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const { RECORD_TEXT } = await import('../src/core/text-recording')
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!

    const canvas = createCanvas(200, 100)
    const ctx = canvas.getContext('2d') as any
    const recorded: any[] = []
    ctx[RECORD_TEXT] = (text: string, _x: number, _y: number, _w: number, logical: any) => {
      recorded.push({ text, start: logical.start, end: logical.end, source: logical.source.text })
    }
    paintDrawing(drawing, ctx, 200, 100)
    expect(recorded.length).toBeGreaterThan(0)
    const combined = recorded.map(r => r.source.slice(r.start, r.end)).join('')
    expect(combined).toBe('Recorded Art')
  })

  test('regression gate: XML br in ordinary DOCX layout emits exact lines without leading blanks and handles multiple, leading, trailing, and programmatic breakBefore', async () => {
    const { layoutDocx } = await import('../src/docx/layout')
    const { parseDocx } = await import('../src/docx/parse')
    const measureFixed = (text: string, style: { fontSizePt: number }) => text.length * style.fontSizePt * 0.6 * (96 / 72)

    // Subcase 1: A <w:br/> B => exactly 2 physical lines, no leading blank
    const zip1 = new JSZip()
    zip1.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r></w:p></w:body></w:document>`)
    zip1.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc1 = await parseDocx(await OfficePackage.load(await zip1.generateAsync({ type: 'uint8array' })))
    expect(doc1.sections[0].paragraphs[0].runs[0].text).toBe('A\nB')
    expect(doc1.sections[0].paragraphs[0].runs[0].breakBefore).toBeUndefined()
    const pages1 = layoutDocx(doc1, measureFixed)
    expect(pages1[0].lines).toHaveLength(2)
    expect(pages1[0].lines[0].segs[0].text).toBe('A')
    expect(pages1[0].lines[1].segs[0].text).toBe('B')

    // Subcase 2: two <w:br/> => A <w:br/><w:br/> B => exactly 3 physical lines
    const zip2 = new JSZip()
    zip2.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t><w:br/><w:br/><w:t>B</w:t></w:r></w:p></w:body></w:document>`)
    zip2.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc2 = await parseDocx(await OfficePackage.load(await zip2.generateAsync({ type: 'uint8array' })))
    expect(doc2.sections[0].paragraphs[0].runs[0].text).toBe('A\n\nB')
    const pages2 = layoutDocx(doc2, measureFixed)
    expect(pages2[0].lines).toHaveLength(3)
    expect(pages2[0].lines[0].segs[0].text).toBe('A')
    expect(pages2[0].lines[1].segs).toHaveLength(0) // empty authored blank line
    expect(pages2[0].lines[2].segs[0].text).toBe('B')

    // Subcase 3: leading <w:br/> => <w:br/> B => exactly 2 physical lines
    const zip3 = new JSZip()
    zip3.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:br/><w:t>B</w:t></w:r></w:p></w:body></w:document>`)
    zip3.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc3 = await parseDocx(await OfficePackage.load(await zip3.generateAsync({ type: 'uint8array' })))
    const pages3 = layoutDocx(doc3, measureFixed)
    expect(pages3[0].lines).toHaveLength(2)
    expect(pages3[0].lines[0].segs).toHaveLength(0) // leading blank
    expect(pages3[0].lines[1].segs[0].text).toBe('B')

    // Subcase 4: trailing <w:br/> => A <w:br/> => exactly 2 physical lines
    const zip4 = new JSZip()
    zip4.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t><w:br/></w:r></w:p></w:body></w:document>`)
    zip4.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc4 = await parseDocx(await OfficePackage.load(await zip4.generateAsync({ type: 'uint8array' })))
    const pages4 = layoutDocx(doc4, measureFixed)
    expect(pages4[0].lines).toHaveLength(2)
    expect(pages4[0].lines[0].segs[0].text).toBe('A')
    expect(pages4[0].lines[1].segs).toHaveLength(0) // trailing blank

    // Subcase 5: programmatic breakBefore on second run => exactly 2 lines
    const para5 = {
      runs: [
        { text: 'A', fontSizePt: 12 },
        { text: 'B', fontSizePt: 12, breakBefore: true },
      ],
      images: [],
      align: 'left' as const,
    }
    const doc5 = {
      sections: [{
        pageSize: { widthTwips: 12240, heightTwips: 15840, orientation: 'portrait' as const },
        margins: { topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, headerTwips: 720, footerTwips: 720, gutterTwips: 0 },
        paragraphs: [para5],
        blocks: [{ kind: 'p' as const, paragraph: para5 }],
      }],
    }
    const pages5 = layoutDocx(doc5 as any, measureFixed)
    expect(pages5[0].lines).toHaveLength(2)
    expect(pages5[0].lines[0].segs[0].text).toBe('A')
    expect(pages5[0].lines[1].segs[0].text).toBe('B')
  })

  test('regression gate: textbox with XML br handles source order, lines and once-only text recording', async () => {
    const zip = new JSZip()
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
        <wp:docPr id="5" name="Textbox BR"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
            <wps:txbx><w:txbxContent><w:p><w:r><w:t>First</w:t><w:br/><w:t>Second</w:t></w:r></w:p></w:txbxContent></wps:txbx>
            <wps:bodyPr wrap="none" anchor="ctr"/>
          </wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body>
    </w:document>`
    zip.file('word/document.xml', docxXml)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const { RECORD_TEXT } = await import('../src/core/text-recording')
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!

    // 1. Visual paint observation: observe fillText baselines
    const canvas = createCanvas(200, 100)
    const ctx = canvas.getContext('2d') as any
    const fills: Array<{ text: string; x: number; y: number }> = []
    const origFill = ctx.fillText.bind(ctx)
    ctx.fillText = (text: string, x: number, y: number) => {
      fills.push({ text, x, y })
      origFill(text, x, y)
    }
    paintDrawing(drawing, ctx, 200, 100)

    // Exactly 2 lines painted: First and Second, no leading blank
    expect(fills).toHaveLength(2)
    expect(fills[0].text).toBe('First')
    expect(fills[1].text).toBe('Second')
    expect(fills[1].y).toBeGreaterThan(fills[0].y) // Second line baseline below first line

    // 2. RECORD_TEXT observation: accumulate start/end coverage of 'First\nSecond'
    const recCanvas = createCanvas(200, 100)
    const recCtx = recCanvas.getContext('2d') as any
    const recorded: Array<{ text: string; start: number; end: number; source: string }> = []
    recCtx[RECORD_TEXT] = (text: string, _x: number, _y: number, _w: number, logical: any) => {
      recorded.push({ text, start: logical.start, end: logical.end, source: logical.source.text })
    }
    paintDrawing(drawing, recCtx, 200, 100)
    expect(recorded.length).toBeGreaterThan(0)
    const sourceText = 'First\nSecond'
    expect(recorded[0].source).toBe(sourceText)

    // Monotonic ranges check
    for (let i = 1; i < recorded.length; i++) {
      expect(recorded[i].start).toBeGreaterThanOrEqual(recorded[i - 1].end)
    }

    // Every UTF-16 position from 0 to sourceText.length is covered exactly once
    const covered = new Array<number>(sourceText.length).fill(0)
    for (const r of recorded) {
      expect(r.start).toBeLessThan(r.end)
      for (let pos = r.start; pos < r.end; pos++) {
        covered[pos]++
      }
    }
    for (let pos = 0; pos < sourceText.length; pos++) {
      expect(covered[pos]).toBe(1)
    }
  })

  test('regression gate: empty w14:textFill solid black contract vs absent fill and unsupported fill', async () => {
    const { readRunProperties } = await import('../src/docx/styles')
    const { parseXmlOrdered } = await import('../src/core/xml')

    // Truly empty <w14:textFill/> element => solid black '000000', clear noFill & textFill
    const emptyFillXml = parseXmlOrdered('<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:color w:val="FF00FF"/><w14:textFill/></w:rPr>')
    const emptyResult = readRunProperties(emptyFillXml)
    expect(emptyResult.color).toBe('000000')
    expect(emptyResult.noFill).toBe(false)
    expect(emptyResult.textFill).toBeUndefined()

    // Absent textFill => preserves existing color 'FF00FF'
    const absentFillXml = parseXmlOrdered('<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:color w:val="FF00FF"/></w:rPr>')
    const absentResult = readRunProperties(absentFillXml)
    expect(absentResult.color).toBe('FF00FF')

    // Unsupported child element in textFill => diagnostic emitted and does not force black
    const issues: any[] = []
    const unsupportedFillXml = parseXmlOrdered('<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:color w:val="0000FF"/><w14:textFill><w14:grpFill/></w14:textFill></w:rPr>')
    const unsupportedResult = readRunProperties(unsupportedFillXml, undefined, issues)
    expect(issues.some(i => i.kind === 'unsupported-text-appearance' && i.feature.includes('grpFill'))).toBe(true)
    expect(unsupportedResult.color).toBe('0000FF')
  })

  test('regression gate: w14:shadow parses with finite checks and robust diagnostics', async () => {
    const { readRunProperties } = await import('../src/docx/styles')
    const { parseXmlOrdered } = await import('../src/core/xml')

    const shadowXml = parseXmlOrdered('<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:shadow blurRad="25400" dist="38100" dir="5400000"><w14:srgbClr w14:val="000088"/></w14:shadow></w:rPr>')
    const res = readRunProperties(shadowXml)
    expect(res.textShadow).toBeDefined()
    expect(res.textShadow?.color).toBe('#000088')
    expect(Number.isFinite(res.textShadow?.offsetX)).toBe(true)
    expect(Number.isFinite(res.textShadow?.offsetY)).toBe(true)
    expect(Number.isFinite(res.textShadow?.blurPx)).toBe(true)
  })

  test('regression gate: w14:gradFill defaults - two stops with no lin/path retains gradient angle 0 and paints linear gradient without diagnostic', async () => {
    const { readRunProperties } = await import('../src/docx/styles')
    const { parseXmlOrdered } = await import('../src/core/xml')
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')

    // 1. readRunProperties with two stops, no lin, no path
    const gradNoLinXml = parseXmlOrdered(`
      <w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w14:textFill>
          <w14:gradFill>
            <w14:gsLst>
              <w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs>
              <w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs>
            </w14:gsLst>
          </w14:gradFill>
        </w14:textFill>
      </w:rPr>
    `)
    const issues: any[] = []
    const runProps = readRunProperties(gradNoLinXml, undefined, issues)
    expect(issues).toHaveLength(0)
    expect(runProps.textFill).toBeDefined()
    expect(runProps.textFill).toEqual({
      kind: 'gradient',
      stops: [
        { position: 0, color: '#FF0000' },
        { position: 1, color: '#0000FF' },
      ],
      angle: 0,
    })

    // 2. Full textbox paint observation: createLinearGradient is called
    const zip = new JSZip()
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
        <wp:docPr id="101" name="Default Gradient"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
            <wps:txbx><w:txbxContent>
              <w:p><w:r>
                <w:rPr>
                  <w14:textFill>
                    <w14:gradFill>
                      <w14:gsLst>
                        <w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs>
                        <w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs>
                      </w14:gsLst>
                    </w14:gradFill>
                  </w14:textFill>
                </w:rPr>
                <w:t>LinearDefault</w:t>
              </w:r></w:p>
            </w:txbxContent></wps:txbx>
            <wps:bodyPr wrap="none" anchor="ctr"/>
          </wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body>
    </w:document>`
    zip.file('word/document.xml', docxXml)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing as any).diagnostics as any[])?.filter((d: any) => d.feature.startsWith('gradFill')) ?? []).toHaveLength(0)

    const canvas = createCanvas(200, 100)
    const ctx = canvas.getContext('2d') as any
    let gradientCalls = 0
    const origCreateLinear = ctx.createLinearGradient.bind(ctx)
    ctx.createLinearGradient = (x0: number, y0: number, x1: number, y1: number) => {
      gradientCalls++
      return origCreateLinear(x0, y0, x1, y1)
    }
    paintDrawing(drawing, ctx, 200, 100)
    expect(gradientCalls).toBeGreaterThan(0)
  })

  test('regression gate: w14:gradFill defaults - absent gsLst paints solid black over inherited nonblack color without diagnostic', async () => {
    const { readRunProperties } = await import('../src/docx/styles')
    const { parseXmlOrdered } = await import('../src/core/xml')
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')

    // 1. readRunProperties with empty gradFill (<w14:gradFill/>) over color 112233
    const emptyGradXml = parseXmlOrdered(`
      <w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:color w:val="112233"/>
        <w14:textFill>
          <w14:gradFill/>
        </w14:textFill>
      </w:rPr>
    `)
    const issues: any[] = []
    const runProps = readRunProperties(emptyGradXml, undefined, issues)
    expect(issues).toHaveLength(0)
    expect(runProps.color).toBe('000000')
    expect(runProps.textFill).toBeUndefined()
    expect(runProps.noFill).toBe(false)

    // 2. Full textbox paint observation: inherits 112233 from style, but <w14:gradFill/> defaults to solid black
    const zip = new JSZip()
    const stylesXml = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="NonBlackStyle">
        <w:name w:val="NonBlackStyle"/>
        <w:rPr><w:color w:val="112233"/></w:rPr>
      </w:style>
    </w:styles>`
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
        <wp:docPr id="102" name="Black Default"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
            <wps:txbx><w:txbxContent>
              <w:p>
                <w:pPr><w:pStyle w:val="NonBlackStyle"/></w:pPr>
                <w:r>
                  <w:rPr>
                    <w14:textFill><w14:gradFill/></w14:textFill>
                  </w:rPr>
                  <w:t>BlackText</w:t>
                </w:r>
              </w:p>
            </w:txbxContent></wps:txbx>
            <wps:bodyPr wrap="none" anchor="ctr"/>
          </wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body>
    </w:document>`
    zip.file('word/styles.xml', stylesXml)
    zip.file('word/document.xml', docxXml)
    zip.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = doc.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing as any).diagnostics as any[])?.filter((d: any) => d.feature.startsWith('gradFill')) ?? []).toHaveLength(0)

    const canvas = createCanvas(200, 100)
    const ctx = canvas.getContext('2d') as any
    const fillStyles: string[] = []
    const origFillText = ctx.fillText.bind(ctx)
    ctx.fillText = (text: string, x: number, y: number) => {
      fillStyles.push(String(ctx.fillStyle))
      origFillText(text, x, y)
    }
    paintDrawing(drawing, ctx, 200, 100)
    expect(fillStyles.length).toBeGreaterThan(0)
    for (const style of fillStyles) {
      // Must be solid black, NOT #112233
      expect(style.toLowerCase()).toBe('#000000')
    }
  })

  test('regression gate: inherited effective diagnostics, source attribution, and override controls', async () => {
    const { parseDocx } = await import('../src/docx/parse')

    // 1. Style-inherited unsupported glow is diagnosed on paragraph and drawing
    const zip1 = new JSZip()
    zip1.file('word/styles.xml', `
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:style w:type="paragraph" w:styleId="GlowStyle">
          <w:name w:val="GlowStyle"/>
          <w:rPr>
            <w14:glow w14:rad="63500"><w14:srgbClr w14:val="FF8800"/></w14:glow>
          </w:rPr>
        </w:style>
      </w:styles>
    `)
    zip1.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="201" name="Inherited Glow"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:pPr><w:pStyle w:val="GlowStyle"/></w:pPr>
                  <w:r><w:t>InheritedGlow</w:t></w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip1.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc1 = await parseDocx(await OfficePackage.load(await zip1.generateAsync({ type: 'uint8array' })))
    const drawing1 = doc1.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing1 as any).diagnostics as any[])?.some((d: any) => d.feature === 'glow')).toBe(true)
    expect((drawing1 as any).paragraphs[0].diagnostics?.some((d: any) => d.feature === 'glow')).toBe(true)

    // 2. Ordinary w:color alone retains inherited unsupported w14 fill warning; explicit same-color w14 solid suppresses it
    const zip2a = new JSZip()
    zip2a.file('word/styles.xml', `
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:style w:type="paragraph" w:styleId="BlipStyle">
          <w:name w:val="BlipStyle"/>
          <w:rPr>
            <w:color w:val="0000FF"/>
            <w14:textFill><w14:blipFill/></w14:textFill>
          </w:rPr>
        </w:style>
      </w:styles>
    `)
    zip2a.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="202a" name="Color Alone Retains Fill Issue"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:pPr><w:pStyle w:val="BlipStyle"/></w:pPr>
                  <w:r>
                    <w:rPr>
                      <w:color w:val="0000FF"/>
                    </w:rPr>
                    <w:t>ColorAlone</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip2a.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc2a = await parseDocx(await OfficePackage.load(await zip2a.generateAsync({ type: 'uint8array' })))
    const drawing2a = doc2a.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing2a as any).diagnostics as any[])?.some((d: any) => d.feature === 'blipFill')).toBe(true)

    const zip2b = new JSZip()
    zip2b.file('word/styles.xml', `
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:style w:type="paragraph" w:styleId="BlipStyle">
          <w:name w:val="BlipStyle"/>
          <w:rPr>
            <w:color w:val="0000FF"/>
            <w14:textFill><w14:blipFill/></w14:textFill>
          </w:rPr>
        </w:style>
      </w:styles>
    `)
    zip2b.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="202b" name="Same Color Override"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:pPr><w:pStyle w:val="BlipStyle"/></w:pPr>
                  <w:r>
                    <w:rPr>
                      <w14:textFill><w14:solidFill><w14:srgbClr w14:val="0000FF"/></w14:solidFill></w14:textFill>
                    </w:rPr>
                    <w:t>OverriddenFill</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip2b.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc2b = await parseDocx(await OfficePackage.load(await zip2b.generateAsync({ type: 'uint8array' })))
    const drawing2b = doc2b.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing2b as any).diagnostics as any[])?.some((d: any) => d.feature === 'blipFill')).toBeFalsy()

    // 3. Explicit outline noFill clears outline and suppresses inherited outline issue
    const zip3 = new JSZip()
    zip3.file('word/styles.xml', `
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:style w:type="paragraph" w:styleId="BadOutlineStyle">
          <w:name w:val="BadOutlineStyle"/>
          <w:rPr>
            <w14:textOutline w:w="0"/>
          </w:rPr>
        </w:style>
      </w:styles>
    `)
    zip3.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="203" name="Outline Clear"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:pPr><w:pStyle w:val="BadOutlineStyle"/></w:pPr>
                  <w:r>
                    <w:rPr>
                      <w14:textOutline><w14:noFill/></w14:textOutline>
                    </w:rPr>
                    <w:t>ClearedOutline</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip3.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc3 = await parseDocx(await OfficePackage.load(await zip3.generateAsync({ type: 'uint8array' })))
    const drawing3 = doc3.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing3 as any).diagnostics as any[])?.some((d: any) => d.feature === 'textOutline')).toBeFalsy()

    // 4. Mixed-run override control: run 1 overrides fill, run 2 does not -> inherited fill issue remains effective
    const zip4 = new JSZip()
    zip4.file('word/styles.xml', `
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:style w:type="paragraph" w:styleId="BlipStyle2">
          <w:name w:val="BlipStyle2"/>
          <w:rPr>
            <w14:textFill><w14:blipFill/></w14:textFill>
          </w:rPr>
        </w:style>
      </w:styles>
    `)
    zip4.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="204" name="Mixed Runs"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:pPr><w:pStyle w:val="BlipStyle2"/></w:pPr>
                  <w:r>
                    <w:rPr><w14:textFill><w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill></w14:textFill></w:rPr>
                    <w:t>Overridden</w:t>
                  </w:r>
                  <w:r>
                    <w:t>Inherited</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip4.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc4 = await parseDocx(await OfficePackage.load(await zip4.generateAsync({ type: 'uint8array' })))
    const drawing4 = doc4.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing4 as any).diagnostics as any[])?.some((d: any) => d.feature === 'blipFill')).toBe(true)

    // 5. Attribution: two paragraphs; p0 plain, p1 glow -> p0 has 0 diagnostics, p1 has glow, drawing has glow once
    const zip5 = new JSZip()
    zip5.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="205" name="Two Paragraph Attribution"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:r><w:t>Plain Paragraph Zero</w:t></w:r>
                </w:p>
                <w:p>
                  <w:r>
                    <w:rPr><w14:glow w14:rad="63500"><w14:srgbClr w14:val="FF8800"/></w14:glow></w:rPr>
                    <w:t>Glow Paragraph One</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip5.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc5 = await parseDocx(await OfficePackage.load(await zip5.generateAsync({ type: 'uint8array' })))
    const drawing5 = doc5.sections[0].paragraphs[0].images[0].drawing!
    expect((drawing5 as any).paragraphs[0].diagnostics ?? []).toHaveLength(0)
    expect((drawing5 as any).paragraphs[1].diagnostics?.filter((d: any) => d.feature === 'glow')).toHaveLength(1)
    expect(((drawing5 as any).diagnostics as any[])?.filter((d: any) => d.feature === 'glow')).toHaveLength(1)

    // 6. Inherited malformed shadow dir/dist warning survives supported direct fill and uses shadow category
    const { issueCategory } = await import('../src/docx/styles')
    expect(issueCategory('shadow:dir')).toBe('shadow')
    expect(issueCategory('shadow:dist')).toBe('shadow')
    const zip6 = new JSZip()
    zip6.file('word/styles.xml', `
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:style w:type="paragraph" w:styleId="MalformedShadowStyle">
          <w:name w:val="MalformedShadowStyle"/>
          <w:rPr>
            <w14:shadow w14:dist="38100" w14:dir="1e308">
              <w14:srgbClr w14:val="000088"/>
            </w14:shadow>
          </w:rPr>
        </w:style>
      </w:styles>
    `)
    zip6.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="206" name="Shadow Survives Direct Fill"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:pPr><w:pStyle w:val="MalformedShadowStyle"/></w:pPr>
                  <w:r>
                    <w:rPr>
                      <w14:textFill><w14:solidFill><w14:srgbClr w14:val="00AA00"/></w14:solidFill></w14:textFill>
                    </w:rPr>
                    <w:t>DirectFill</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip6.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc6 = await parseDocx(await OfficePackage.load(await zip6.generateAsync({ type: 'uint8array' })))
    const drawing6 = doc6.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing6 as any).diagnostics as any[])?.some((d: any) => d.feature === 'shadow:dir')).toBe(true)
    expect((drawing6 as any).paragraphs[0].diagnostics?.some((d: any) => d.feature === 'shadow:dir')).toBe(true)

    // 7. Non-schema w14 pattFill produces one warning and readable fallback
    const { paintDrawing: paintDocxDrawing } = await import('../src/docx/drawing')
    const zip7 = new JSZip()
    zip7.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="207" name="Pattern Fallback"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w14:textFill>
                        <w14:pattFill prst="pct5">
                          <w14:fgClr><w14:srgbClr w14:val="008800"/></w14:fgClr>
                          <w14:bgClr><w14:srgbClr w14:val="FFFFFF"/></w14:bgClr>
                        </w14:pattFill>
                      </w14:textFill>
                    </w:rPr>
                    <w:t>PatternFallbackText</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip7.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc7 = await parseDocx(await OfficePackage.load(await zip7.generateAsync({ type: 'uint8array' })))
    const drawing7 = doc7.sections[0].paragraphs[0].images[0].drawing!
    const pattDiags = ((drawing7 as any).diagnostics as any[])?.filter((d: any) => d.feature === 'pattFill') ?? []
    expect(pattDiags).toHaveLength(1)
    expect(pattDiags[0].kind).toBe('unsupported-text-appearance')
    const run7 = (drawing7 as any).paragraphs[0].runs[0]
    expect(run7.textFill).toBeUndefined()
    expect(run7.color).toBe('#008800')
    const canvas7 = createCanvas(200, 100)
    const ctx7 = canvas7.getContext('2d') as any
    let paintFillStyle7: string | undefined
    const origFillText7 = ctx7.fillText.bind(ctx7)
    ctx7.fillText = (text: string, x: number, y: number) => {
      paintFillStyle7 = ctx7.fillStyle
      origFillText7(text, x, y)
    }
    paintDocxDrawing(drawing7, ctx7, 200, 100)
    expect(paintFillStyle7).toBe('#008800')
  })

  test('regression gate: extreme numeric overflow validation and actual canvas gates', async () => {
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')

    // 1. Gradient angle overflow: <w14:lin w14:ang="1e308"/> diagnoses gradFill:ang, falls back legibly, canvas does not receive NaN
    const zip1 = new JSZip()
    zip1.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="301" name="Gradient Overflow"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w14:textFill>
                        <w14:gradFill>
                          <w14:gsLst>
                            <w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs>
                            <w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs>
                          </w14:gsLst>
                          <w14:lin w14:ang="1e308"/>
                        </w14:gradFill>
                      </w14:textFill>
                    </w:rPr>
                    <w:t>ExtremeAngle</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip1.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc1 = await parseDocx(await OfficePackage.load(await zip1.generateAsync({ type: 'uint8array' })))
    const drawing1 = doc1.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing1 as any).diagnostics as any[])?.some((d: any) => d.feature.startsWith('gradFill'))).toBe(true)

    const canvas1 = createCanvas(200, 100)
    const ctx1 = canvas1.getContext('2d') as any
    const linearGradientArgs: number[][] = []
    const origCreateLinear = ctx1.createLinearGradient.bind(ctx1)
    ctx1.createLinearGradient = (...args: number[]) => {
      linearGradientArgs.push(args)
      return origCreateLinear(...args)
    }
    paintDrawing(drawing1, ctx1, 200, 100)
    for (const args of linearGradientArgs) {
      for (const arg of args) {
        expect(Number.isFinite(arg)).toBe(true)
      }
    }

    // 2. Shadow angle overflow: <w14:shadow w14:dist="38100" w14:dir="1e308"> diagnoses shadow:dir, canvas shadow offset never NaN
    const zip2 = new JSZip()
    zip2.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="302" name="Shadow Overflow"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w14:shadow w14:dist="38100" w14:dir="1e308">
                        <w14:srgbClr w14:val="000088"/>
                      </w14:shadow>
                    </w:rPr>
                    <w:t>ExtremeShadowDir</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip2.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc2 = await parseDocx(await OfficePackage.load(await zip2.generateAsync({ type: 'uint8array' })))
    const drawing2 = doc2.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing2 as any).diagnostics as any[])?.some((d: any) => d.feature.startsWith('shadow'))).toBe(true)

    const canvas2 = createCanvas(200, 100)
    const ctx2 = canvas2.getContext('2d') as any
    const shadowOffsets: Array<{ x: number; y: number }> = []
    const origFill2 = ctx2.fillText.bind(ctx2)
    ctx2.fillText = (text: string, x: number, y: number) => {
      shadowOffsets.push({ x: ctx2.shadowOffsetX, y: ctx2.shadowOffsetY })
      origFill2(text, x, y)
    }
    paintDrawing(drawing2, ctx2, 200, 100)
    for (const offset of shadowOffsets) {
      expect(Number.isFinite(offset.x)).toBe(true)
      expect(Number.isFinite(offset.y)).toBe(true)
    }

    // 3. Qualified w14:shadow attributes parse accurately and pass finite numbers to canvas
    const zip3 = new JSZip()
    zip3.file('word/document.xml', `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p><w:r><w:drawing><wp:inline>
          <wp:extent cx="${emu(200)}" cy="${emu(100)}"/>
          <wp:docPr id="303" name="Qualified Shadow"/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(200)}" cy="${emu(100)}"/></a:xfrm></wps:spPr>
              <wps:txbx><w:txbxContent>
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w14:shadow w14:blurRad="25400" w14:dist="38100" w14:dir="5400000">
                        <w14:srgbClr w14:val="005588"/>
                      </w14:shadow>
                    </w:rPr>
                    <w:t>QualifiedShadow</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr wrap="none" anchor="ctr"/>
            </wps:wsp>
          </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p></w:body>
      </w:document>
    `)
    zip3.file('word/_rels/document.xml.rels', '<Relationships/>')
    const doc3 = await parseDocx(await OfficePackage.load(await zip3.generateAsync({ type: 'uint8array' })))
    const drawing3 = doc3.sections[0].paragraphs[0].images[0].drawing!
    expect(((drawing3 as any).diagnostics as any[])?.filter((d: any) => d.feature.startsWith('shadow')) ?? []).toHaveLength(0)

    const canvas3 = createCanvas(200, 100)
    const ctx3 = canvas3.getContext('2d') as any
    let capturedShadowColor = ''
    let capturedOffsetY = 0
    const origFill3 = ctx3.fillText.bind(ctx3)
    ctx3.fillText = (text: string, x: number, y: number) => {
      capturedShadowColor = String(ctx3.shadowColor)
      capturedOffsetY = ctx3.shadowOffsetY
      origFill3(text, x, y)
    }
    paintDrawing(drawing3, ctx3, 200, 100)
    expect(capturedShadowColor.toLowerCase()).toBe('#005588')
    expect(Number.isFinite(capturedOffsetY)).toBe(true)
    expect(capturedOffsetY).toBeGreaterThan(0)
  })

  test('XLSX DrawingML shapes (<xdr:sp>) parse and render WordArt styled text runs', async () => {
    const { buildXlsx } = await import('../src/testdata/ooxml-builders')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const { renderSheet } = await import('../src/xlsx/render')

    const xdrSp = `<xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="10" name="WordArt Shape 10"/></xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(180)}" cy="${emu(80)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr sz="2200">
              <a:gradFill>
                <a:gsLst>
                  <a:gs pos="0"><a:srgbClr val="00AAFF"/></a:gs>
                  <a:gs pos="100000"><a:srgbClr val="FFAA00"/></a:gs>
                </a:gsLst>
                <a:lin ang="10800000"/>
              </a:gradFill>
              <a:ln w="12700">
                <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
              </a:ln>
              <a:effectLst>
                <a:outerShdw blurRad="19050" dist="19050" dir="5400000">
                  <a:srgbClr val="445566"/>
                </a:outerShdw>
              </a:effectLst>
            </a:rPr>
            <a:t>Xlsx WordArt</a:t>
          </a:r>
        </a:p>
      </xdr:txBody>
    </xdr:sp>`

    const drawingXml = `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="${emu(180)}" cy="${emu(80)}"/>${xdrSp}<xdr:clientData/></xdr:absoluteAnchor>`

    const base = await buildXlsx([{ name: 'Sheet1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }])
    const zip = await JSZip.loadAsync(base)
    let sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    sheetXml = sheetXml.replace('</worksheet>', '<drawing r:id="rDraw"/></worksheet>')
    zip.file('xl/worksheets/sheet1.xml', sheetXml)
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships><Relationship Id="rDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
    )
    zip.file(
      'xl/drawings/drawing1.xml',
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${drawingXml}</xdr:wsDr>`,
    )

    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets[0]
    expect(sheet.drawings).toBeDefined()
    expect(sheet.drawings!.length).toBe(1)

    const shape = sheet.drawings![0]
    expect(shape.textBody).toBeDefined()
    const run = shape.textBody!.paragraphs[0].runs[0]
    expect(run.text).toBe('Xlsx WordArt')
    expect(run.textFill).toEqual({
      kind: 'gradient',
      angle: Math.PI,
      stops: [
        { position: 0, color: '#00AAFF' },
        { position: 1, color: '#FFAA00' },
      ],
    })
    expect(run.textOutline).toBeDefined()
    expect(run.textOutline?.color).toBe('#112233')
    expect(run.textOutline?.widthPx).toBeCloseTo(1.33, 1)
    expect(run.textShadow).toBeDefined()
    expect(run.textShadow?.color).toBe('#445566')
    expect(run.textShadow?.blurPx).toBeCloseTo(2, 1)

    // Render canvas
    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d')
    const strokes: string[] = []
    const gradStops: Array<{ offset: number; color: string }> = []
    const shadows: string[] = []

    const origCreateGrad = ctx.createLinearGradient.bind(ctx)
    ctx.createLinearGradient = (x0, y0, x1, y1) => {
      const grad = origCreateGrad(x0, y0, x1, y1)
      const origAdd = grad.addColorStop.bind(grad)
      grad.addColorStop = (o, c) => {
        gradStops.push({ offset: o, color: c })
        return origAdd(o, c)
      }
      return grad
    }
    const origStroke = ctx.strokeText.bind(ctx)
    ctx.strokeText = (t, x, y) => {
      strokes.push(String(t))
      return origStroke(t, x, y)
    }
    const origFill = ctx.fillText.bind(ctx)
    ctx.fillText = (t, x, y) => {
      shadows.push(String(ctx.shadowColor))
      return origFill(t, x, y)
    }

    renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D)
    expect(gradStops.length).toBeGreaterThanOrEqual(2)
    expect(gradStops[0].color).toBe('#00AAFF')
    expect(gradStops[1].color).toBe('#FFAA00')
    expect(strokes.length).toBeGreaterThan(0)
    expect(shadows).toContain('#445566')
  })

  test('WordArt diagnostics flow into single body diagnostic channel across all three formats', async () => {
    const unsupportedPattXml = `<a:rPr>
      <a:pattFill prst="unsupportedPreset">
        <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
        <a:bgClr><a:srgbClr val="0000FF"/></a:bgClr>
      </a:pattFill>
    </a:rPr><a:t>Bad Pattern</a:t>`

    // 1. PPTX
    const pptxSp = `<p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Bad Patt"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(50)}"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:r>${unsupportedPattXml}</a:r></a:p></p:txBody>
    </p:sp>`
    const pptxDoc = await createPptxFixture(pptxSp)
    const pptxBody = pptxDoc.slides[0].shapes[0].textBody
    const expectedDiag = {
      kind: 'unsupported-text-appearance',
      feature: 'pattFill:unsupportedPreset',
      message: 'WordArt pattern preset is deferred; using flat color',
    }
    expect(pptxBody?.diagnostics).toContainEqual(expectedDiag)

    // 2. DOCX
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="${emu(100)}" cy="${emu(50)}"/>
        <wp:docPr id="1" name="Docx Bad Glow"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(50)}"/></a:xfrm></wps:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w14:glow w14:rad="63500"><w14:srgbClr w14:val="FFFF00"/></w14:glow>
                    </w:rPr>
                    <w:t>Bad Glow</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent>
            </wps:txbx>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body>
    </w:document>`
    const zipDocx = new JSZip()
    zipDocx.file('word/document.xml', docxXml)
    zipDocx.file('word/_rels/document.xml.rels', '<Relationships/>')
    const { parseDocx } = await import('../src/docx/parse')
    const docxDoc = await parseDocx(await OfficePackage.load(await zipDocx.generateAsync({ type: 'uint8array' })))
    const docxDrawing = docxDoc.sections[0].paragraphs[0].images[0].drawing
    expect(docxDrawing?.kind).toBe('textbox')
    if (docxDrawing?.kind !== 'textbox') throw new Error('Expected textbox')
    const docxDiag = (docxDrawing as any).diagnostics ?? (docxDrawing as any).paragraphs[0].diagnostics
    expect(docxDiag).toContainEqual({
      kind: 'unsupported-text-appearance',
      feature: 'glow',
      message: 'WordArt text glow effect is unsupported; falling back to plain run',
    })

    // 3. XLSX
    const { buildXlsx } = await import('../src/testdata/ooxml-builders')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const xlsxSp = `<xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="2" name="Xlsx Bad Patt"/></xdr:nvSpPr>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(50)}"/></a:xfrm></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r>${unsupportedPattXml}</a:r></a:p></xdr:txBody>
    </xdr:sp>`
    const drawingXml = `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="${emu(100)}" cy="${emu(50)}"/>${xlsxSp}<xdr:clientData/></xdr:absoluteAnchor>`
    const baseXlsx = await buildXlsx([{ name: 'Sheet1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }])
    const zipXlsx = await JSZip.loadAsync(baseXlsx)
    let sheetXml = await zipXlsx.file('xl/worksheets/sheet1.xml')!.async('string')
    sheetXml = sheetXml.replace('</worksheet>', '<drawing r:id="rDraw"/></worksheet>')
    zipXlsx.file('xl/worksheets/sheet1.xml', sheetXml)
    zipXlsx.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships><Relationship Id="rDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
    )
    zipXlsx.file(
      'xl/drawings/drawing1.xml',
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${drawingXml}</xdr:wsDr>`,
    )
    const xlsxDoc = await parseXlsx(await OfficePackage.load(await zipXlsx.generateAsync({ type: 'uint8array' })))
    const xlsxBody = xlsxDoc.sheets[0].drawings![0].textBody
    expect(xlsxBody?.diagnostics).toContainEqual(expectedDiag)
  })

  test('End-to-end multi-format fixture test parsing and painting WordArt shapes', async () => {
    // 1. PPTX E2E: WordArt shape + unstyled shape on same slide
    const pptxTree = `
      <p:sp>
        <p:nvSpPr><p:cNvPr id="1" name="WordArt 1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(150)}" cy="${emu(50)}"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r>
          <a:rPr sz="2000">
            <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>
            <a:effectLst><a:outerShdw blurRad="25400" dist="19050"><a:srgbClr val="111111"/></a:outerShdw></a:effectLst>
          </a:rPr>
          <a:t>PPTX WordArt</a:t>
        </a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Plain Text"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="${emu(60)}"/><a:ext cx="${emu(150)}" cy="${emu(50)}"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1600"/><a:t>Plain Body</a:t></a:r></a:p></p:txBody>
      </p:sp>
    `
    const pptxDoc = await createPptxFixture(pptxTree)
    const pptxCanvas = createCanvas(300, 200)
    const pptxCtx = pptxCanvas.getContext('2d')
    const pptxFillShadows: number[] = []
    const origPptxFill = pptxCtx.fillText.bind(pptxCtx)
    pptxCtx.fillText = (t, x, y) => {
      pptxFillShadows.push(pptxCtx.shadowBlur)
      return origPptxFill(t, x, y)
    }
    renderSlide(pptxDoc.slides[0], pptxCtx as unknown as CanvasRenderingContext2D)
    // First run (WordArt) had shadow blur > 0, second run (Plain) had shadow blur == 0
    expect(pptxFillShadows.length).toBeGreaterThanOrEqual(2)
    expect(pptxFillShadows[0]).toBeGreaterThan(0)
    expect(pptxFillShadows[pptxFillShadows.length - 1]).toBe(0)

    // 2. DOCX E2E: WordArt drawing
    const docxXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <w:body><w:p><w:r><w:drawing><wp:inline>
        <wp:extent cx="${emu(160)}" cy="${emu(80)}"/>
        <wp:docPr id="1" name="Docx WordArt E2E"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(160)}" cy="${emu(80)}"/></a:xfrm></wps:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w:sz w:val="44"/>
                      <w14:textFill>
                        <w14:gradFill>
                          <w14:gsLst>
                            <w14:gs w14:pos="0"><w14:srgbClr w14:val="FF00FF"/></w14:gs>
                            <w14:gs w14:pos="100000"><w14:srgbClr w14:val="00FFFF"/></w14:gs>
                          </w14:gsLst>
                        </w14:gradFill>
                      </w14:textFill>
                      <w14:textOutline w14:w="12700">
                        <w14:solidFill><w14:srgbClr w14:val="000000"/></w14:solidFill>
                      </w14:textOutline>
                    </w:rPr>
                    <w:t>Docx E2E</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent>
            </wps:txbx>
            <wps:bodyPr wrap="none" anchor="ctr"/>
          </wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body>
    </w:document>`
    const zipDocx = new JSZip()
    zipDocx.file('word/document.xml', docxXml)
    zipDocx.file('word/_rels/document.xml.rels', '<Relationships/>')
    const { parseDocx } = await import('../src/docx/parse')
    const { paintDrawing } = await import('../src/docx/drawing')
    const docxDoc = await parseDocx(await OfficePackage.load(await zipDocx.generateAsync({ type: 'uint8array' })))
    const docxDrawing = docxDoc.sections[0].paragraphs[0].images[0].drawing!
    const docxCanvas = createCanvas(160, 80)
    const docxCtx = docxCanvas.getContext('2d')
    let docxStroked = false
    const origDocxStroke = docxCtx.strokeText.bind(docxCtx)
    docxCtx.strokeText = (t, x, y) => {
      docxStroked = true
      return origDocxStroke(t, x, y)
    }
    paintDrawing(docxDrawing, docxCtx as unknown as CanvasRenderingContext2D, 160, 80)
    expect(docxStroked).toBe(true)

    // 3. XLSX E2E: WordArt drawing
    const { buildXlsx } = await import('../src/testdata/ooxml-builders')
    const { parseXlsx } = await import('../src/xlsx/parse')
    const { renderSheet } = await import('../src/xlsx/render')
    const xlsxSp = `<xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="1" name="Xlsx WordArt E2E"/></xdr:nvSpPr>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(160)}" cy="${emu(80)}"/></a:xfrm></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r>
        <a:rPr sz="2200">
          <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF5500"/></a:gs><a:gs pos="100000"><a:srgbClr val="0055FF"/></a:gs></a:gsLst></a:gradFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
        </a:rPr>
        <a:t>Xlsx E2E</a:t>
      </a:r></a:p></xdr:txBody>
    </xdr:sp>`
    const xlsxDrXml = `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="${emu(160)}" cy="${emu(80)}"/>${xlsxSp}<xdr:clientData/></xdr:absoluteAnchor>`
    const baseXlsx = await buildXlsx([{ name: 'Sheet1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 42 }] }] }])
    const zipXlsx = await JSZip.loadAsync(baseXlsx)
    let sheetXml = await zipXlsx.file('xl/worksheets/sheet1.xml')!.async('string')
    sheetXml = sheetXml.replace('</worksheet>', '<drawing r:id="rDraw"/></worksheet>')
    zipXlsx.file('xl/worksheets/sheet1.xml', sheetXml)
    zipXlsx.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="rDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
    zipXlsx.file('xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xlsxDrXml}</xdr:wsDr>`)
    const xlsxDoc = await parseXlsx(await OfficePackage.load(await zipXlsx.generateAsync({ type: 'uint8array' })))
    const xlsxCanvas = createCanvas(300, 200)
    const xlsxCtx = xlsxCanvas.getContext('2d')
    let xlsxStroked = false
    const origXlsxStroke = xlsxCtx.strokeText.bind(xlsxCtx)
    xlsxCtx.strokeText = (t, x, y) => {
      xlsxStroked = true
      return origXlsxStroke(t, x, y)
    }
    renderSheet(xlsxDoc.sheets[0], xlsxCtx as unknown as CanvasRenderingContext2D)
    expect(xlsxStroked).toBe(true)
  })
})
