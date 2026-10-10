import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { resolveDrawingStyle, computeShadowOffset, shadowAlignmentOrigin, type DrawingStyle, type ThemeContext } from '../src/drawing/style'
import { paintGeometry } from '../src/drawing/paint'
import { resolvePreset } from '../src/drawing/geometry'

const dummyTheme: ThemeContext = {
  colors: {},
  palette: {},
  colorMap: {},
  fonts: { major: { supplemental: {} }, minor: { supplemental: {} } },
  fillStyles: [],
  bgFillStyles: [],
  lineStyles: [],
  effectStyles: [],
  issues: [],
}

describe('Shape outer shadow (<a:outerShdw>)', () => {
  test.each(['translation', 'rotation', 'minification'] as const)('never paints the displaced shadow source under large %s', transform => {
    const rotated = transform === 'rotation'
    const canvas = createCanvas(rotated ? 100 : 3000, rotated ? 3000 : 100)
    const ctx = canvas.getContext('2d')
    if (rotated) { ctx.translate(60, 2500); ctx.rotate(Math.PI / 2) }
    else { ctx.translate(2500, 0); if (transform === 'minification') ctx.scale(.5, .5) }
    const before = ctx.getTransform()
    const style: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } },
      shadow: { color: { r: 0, g: 0, b: 0, a: 1 }, blurPx: 0, offsetX: 40, offsetY: 40 },
      issues: [],
    }
    expect(paintGeometry(ctx as never, resolvePreset('rect', 20, 20), style, 20, 20)).toEqual([])
    expect(ctx.getTransform()).toEqual(before)
    const pixel = (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data)
    const minified = transform === 'minification'
    expect(pixel(rotated ? 50 : minified ? 505 : 510, rotated ? 510 : minified ? 5 : 10)).toEqual([0, 0, 0, 0])
    expect(pixel(rotated ? 50 : minified ? 2505 : 2510, rotated ? 2510 : minified ? 5 : 10)).toEqual([255, 0, 0, 255])
    expect(pixel(rotated ? 10 : minified ? 2525 : 2550, rotated ? 2550 : minified ? 25 : 50)).toEqual([0, 0, 0, 255])
    // Only the actual shape and its offset shadow may contribute opaque pixels.
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data
    let opaque = 0
    for (let i = 3; i < data.length; i += 4) if (data[i] === 255) opaque++
    expect(opaque).toBe(minified ? 200 : 800)
  })

  test('parses outerShdw dist, dir, blurRad, and color into DrawingStyle shadow', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="190500" dir="5400000" blurRad="38100">` +
      `      <a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const style = resolveDrawingStyle(xml, undefined, dummyTheme)
    expect(style.shadow).toBeDefined()
    expect(style.shadow?.blurPx).toBeCloseTo(4, 1)
    expect(style.shadow?.offsetX).toBeCloseTo(0, 1)
    expect(style.shadow?.offsetY).toBeCloseTo(20, 1)
    expect(style.shadow?.color).toEqual({ r: 0, g: 0, b: 0, a: 0.5 })
    expect(style.shadow?.rotWithShape).toBe(true)
    expect(style.issues.some(i => i.kind === 'unsupported-effect')).toBe(false)
  })

  test('parses outerShdw algn and rotWithShape attributes', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="100000" dir="0" blurRad="20000" algn="tl" rotWithShape="0">` +
      `      <a:srgbClr val="000000"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const style = resolveDrawingStyle(xml, undefined, dummyTheme)
    expect(style.shadow).toBeDefined()
    expect(style.shadow?.algn).toBe('tl')
    expect(style.shadow?.rotWithShape).toBe(false)

    // Also verify rotWithShape="1" and explicit algn="ctr"
    const xml2 = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="100000" dir="0" blurRad="20000" algn="ctr" rotWithShape="1">` +
      `      <a:srgbClr val="000000"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const style2 = resolveDrawingStyle(xml2, undefined, dummyTheme)
    expect(style2.shadow?.algn).toBe('ctr')
    expect(style2.shadow?.rotWithShape).toBe(true)
  })

  test('inner shadows, glow, and reflection continue to diagnose and degrade gracefully alongside outerShdw', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>` +
      `  <a:effectLst>` +
      `    <a:glow rad="50000"><a:srgbClr val="FFFF00"/></a:glow>` +
      `    <a:innerShdw dist="50000"><a:srgbClr val="000000"/></a:innerShdw>` +
      `    <a:outerShdw dist="95250" dir="0" blurRad="19050">` +
      `      <a:srgbClr val="333333"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const style = resolveDrawingStyle(xml, undefined, dummyTheme)
    expect(style.shadow).toBeDefined()
    expect(style.shadow?.offsetX).toBeCloseTo(10, 1)
    expect(style.shadow?.offsetY).toBeCloseTo(0, 1)
    expect(style.issues.some(i => i.kind === 'unsupported-effect' && i.feature === 'glow')).toBe(true)
    expect(style.issues.some(i => i.kind === 'unsupported-effect' && i.feature === 'innerShdw')).toBe(true)
    expect(style.issues.some(i => i.kind === 'unsupported-effect' && i.feature === 'outerShdw')).toBe(false)
  })

  test('paints shape outer shadow offset under fill and stroke', () => {
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('rect', 50, 50)
    const style: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } },
      shadow: {
        color: { r: 0, g: 0, b: 0, a: 1 },
        blurPx: 0,
        offsetX: 20,
        offsetY: 20,
      },
      issues: [],
    }
    const issues = paintGeometry(ctx as never, geom, style, 50, 50)
    expect(issues).toEqual([])

    // Shape is at 0..50, 0..50.
    // Shadow is offset by +20, +20, so shadow ink extends to x=60, y=60.
    const data = ctx.getImageData(0, 0, 100, 100).data
    const pixelAt = (x: number, y: number) => {
      const idx = (y * 100 + x) * 4
      return [data[idx], data[idx + 1], data[idx + 2], data[idx + 3]]
    }

    // Inside shape: red fill
    const center = pixelAt(25, 25)
    expect(center[0]).toBe(255)

    // In shadow region (x=60, y=60): black shadow
    const shadowPixel = pixelAt(60, 60)
    expect(shadowPixel[3]).toBeGreaterThan(100)
    expect(shadowPixel[0]).toBeLessThan(50)
  })

  test('shape shadow state is isolated and does not bleed into subsequent drawings', () => {
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('rect', 20, 20)
    const styleWithShadow: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } },
      shadow: {
        color: { r: 0, g: 0, b: 0, a: 1 },
        blurPx: 10,
        offsetX: 30,
        offsetY: 30,
      },
      issues: [],
    }
    paintGeometry(ctx as never, geom, styleWithShadow, 20, 20)

    // Canvas state outside must have zero shadow
    expect(ctx.shadowColor.startsWith('rgba(0, 0, 0, 0')).toBe(true)
    expect(ctx.shadowBlur).toBe(0)
    expect(ctx.shadowOffsetX).toBe(0)
    expect(ctx.shadowOffsetY).toBe(0)

    // Subsequent shape without shadow
    const stylePlain: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 0, g: 255, b: 0, a: 1 } },
      issues: [],
    }
    paintGeometry(ctx as never, geom, stylePlain, 20, 20)
    expect(ctx.shadowColor.startsWith('rgba(0, 0, 0, 0')).toBe(true)
    expect(ctx.shadowBlur).toBe(0)
  })

  test('stroke-only geometry with shadow casts shadow from stroke', () => {
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('line', 40, 0)
    const style: DrawingStyle = {
      line: { width: 4, fill: { kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } } },
      shadow: {
        color: { r: 0, g: 0, b: 0, a: 1 },
        blurPx: 0,
        offsetX: 0,
        offsetY: 20,
      },
      issues: [],
    }
    paintGeometry(ctx as never, geom, style, 40, 40)
    const data = ctx.getImageData(0, 0, 100, 100).data
    const pixelAt = (x: number, y: number) => {
      const idx = (y * 100 + x) * 4
      return [data[idx], data[idx + 1], data[idx + 2], data[idx + 3]]
    }
    // Shadow region around y=20
    const shadowPixel = pixelAt(20, 20)
    expect(shadowPixel[3]).toBeGreaterThan(100)
  })

  test('reports invalid-paint on nonfinite shadow color or metrics without crashing', () => {
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('rect', 20, 20)
    const badColor: DrawingStyle = {
      shadow: { color: { r: NaN, g: 0, b: 0, a: 1 }, blurPx: 0, offsetX: 10, offsetY: 10 },
      issues: [],
    }
    const issuesColor = paintGeometry(ctx as never, geom, badColor, 20, 20)
    expect(issuesColor.some(i => i.kind === 'invalid-paint' && i.message.includes('shadow'))).toBe(true)

    const badMetrics: DrawingStyle = {
      shadow: { color: { r: 0, g: 0, b: 0, a: 1 }, blurPx: NaN, offsetX: 10, offsetY: 10 },
      issues: [],
    }
    const issuesMetrics = paintGeometry(ctx as never, geom, badMetrics, 20, 20)
    expect(issuesMetrics.some(i => i.kind === 'invalid-paint' && i.message.includes('shadow'))).toBe(true)
  })

  test('paints shape outer shadow correctly under canvas rotation', () => {
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    // Rotate 90 degrees around center (50, 50)
    ctx.translate(50, 50)
    ctx.rotate(Math.PI / 2)
    ctx.translate(-50, -50)

    const geom = resolvePreset('rect', 20, 20)
    const style: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } },
      shadow: {
        color: { r: 0, g: 0, b: 0, a: 1 },
        blurPx: 0,
        offsetX: 10,
        offsetY: 10,
      },
      issues: [],
    }
    // Shape is at local (20..40, 20..40), shadow at local (30..50, 30..50).
    // Under 90 deg rotation around (50, 50):
    // Local (40, 40) -> device (60, 40).
    ctx.translate(20, 20)
    paintGeometry(ctx as never, geom, style, 20, 20)

    const data = ctx.getImageData(0, 0, 100, 100).data
    const pixelAt = (x: number, y: number) => {
      const idx = (y * 100 + x) * 4
      return [data[idx], data[idx + 1], data[idx + 2], data[idx + 3]]
    }
    const shadowPixel = pixelAt(60, 40)
    expect(shadowPixel[3]).toBeGreaterThan(100)
    // True shadow-only ink: under the 90-degree rotation the local (10, 10)
    // offset maps to device (-10, +10), so the shadow rect spans device
    // x 50..70, y 30..50; (55, 45) lies inside the shadow but outside the
    // shape itself (device x 60..80, y 20..40).
    const rotatedShadow = pixelAt(55, 45)
    expect(rotatedShadow[3]).toBeGreaterThan(100)
    expect(rotatedShadow[0]).toBeLessThan(50)
  })

  test('hostile shadow distance clamps to text-side offset posture', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="99999999999" dir="0" blurRad="0">` +
      `      <a:srgbClr val="000000"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const style = resolveDrawingStyle(xml, undefined, dummyTheme)
    expect(style.shadow).toBeDefined()
    expect(Math.abs(style.shadow?.offsetX ?? 0)).toBeLessThanOrEqual(200)
    expect(Math.abs(style.shadow?.offsetY ?? 0)).toBeLessThanOrEqual(200)
    // Still paints without crashing.
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('rect', 20, 20)
    const issues = paintGeometry(ctx as never, geom, style, 20, 20)
    expect(issues).toEqual([])
  })

  test('shadow alignment anchor only positions scale/skew, never the offset (ECMA-376)', () => {
    // 1. Normalized alignment anchors
    expect(shadowAlignmentOrigin('tl')).toEqual({ x: 0, y: 0 })
    expect(shadowAlignmentOrigin('t')).toEqual({ x: 0.5, y: 0 })
    expect(shadowAlignmentOrigin('tr')).toEqual({ x: 1, y: 0 })
    expect(shadowAlignmentOrigin('l')).toEqual({ x: 0, y: 0.5 })
    expect(shadowAlignmentOrigin('ctr')).toEqual({ x: 0.5, y: 0.5 })
    expect(shadowAlignmentOrigin('r')).toEqual({ x: 1, y: 0.5 })
    expect(shadowAlignmentOrigin('bl')).toEqual({ x: 0, y: 1 })
    expect(shadowAlignmentOrigin('b')).toEqual({ x: 0.5, y: 1 })
    expect(shadowAlignmentOrigin('br')).toEqual({ x: 1, y: 1 })
    expect(shadowAlignmentOrigin(undefined)).toEqual({ x: 0.5, y: 1 })

    // 2. Per ECMA-376 (CT_OuterShadowEffect), "alignment happens first,
    // effectively setting the origin for scale, skew, and offset". The dist/dir
    // translation is anchor-independent, so at 100% scale / zero skew every
    // algn yields the authored offset unchanged.
    const baseShadow = {
      color: { r: 0, g: 0, b: 0, a: 1 },
      blurPx: 0,
      offsetX: 10,
      offsetY: 20,
    }
    for (const algn of [undefined, 'tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'] as const) {
      expect(computeShadowOffset({ ...baseShadow, algn })).toEqual({ offsetX: 10, offsetY: 20 })
    }

    // 3. Canvas rendering: a PowerPoint-style algn="ctr" offset shadow renders
    // BELOW the shape (pure dist/dir translation), not centered on it.
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const geom = resolvePreset('rect', 40, 40)
    // Shape is at 30..70, 30..70. Shadow offset is (0, 20), so shadow ink
    // spans y=50..90 regardless of algn="ctr".
    ctx.translate(30, 30)
    const styleCtr: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } },
      shadow: {
        color: { r: 0, g: 0, b: 0, a: 1 },
        blurPx: 0,
        offsetX: 0,
        offsetY: 20,
        algn: 'ctr',
      },
      issues: [],
    }
    paintGeometry(ctx as never, geom, styleCtr, 40, 40)

    const data = ctx.getImageData(0, 0, 100, 100).data
    const pixelAt = (x: number, y: number) => {
      const idx = (y * 100 + x) * 4
      return [data[idx], data[idx + 1], data[idx + 2], data[idx + 3]]
    }

    // At (50, 80): below the shape (ends y=70), inside the translated shadow.
    const belowShape = pixelAt(50, 80)
    expect(belowShape[3]).toBeGreaterThan(100)
    expect(belowShape[0]).toBeLessThan(50)
  })

  test('omitted algn and rotWithShape resolve to schema defaults', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="95250" dir="0" blurRad="0">` +
      `      <a:srgbClr val="000000"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const style = resolveDrawingStyle(xml, undefined, dummyTheme)
    expect(style.shadow).toBeDefined()
    expect(style.shadow?.algn).toBeUndefined()
    expect(style.shadow?.rotWithShape).toBe(true)
    expect(style.shadow?.offsetX).toBeCloseTo(10, 1)
    expect(style.shadow?.offsetY).toBeCloseTo(0, 1)
  })

  test('invalid outerShdw algn falls back to the default anchor without crashing', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="95250" dir="0" blurRad="0" algn="bogus">` +
      `      <a:srgbClr val="000000"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const style = resolveDrawingStyle(xml, undefined, dummyTheme)
    expect(style.shadow).toBeDefined()
    expect(style.shadow?.algn).toBeUndefined()
    expect(style.shadow?.offsetX).toBeCloseTo(10, 1)
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const issues = paintGeometry(ctx as never, resolvePreset('rect', 20, 20), style, 20, 20)
    expect(issues).toEqual([])
  })

  test('non-default outerShdw sx/sy/kx/ky diagnose and fall back to pure dist/dir offset', () => {
    const xml = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="95250" dir="0" blurRad="0" sx="50000" sy="50000" kx="60000" ky="0">` +
      `      <a:srgbClr val="000000"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const style = resolveDrawingStyle(xml, undefined, dummyTheme)
    expect(style.shadow).toBeDefined()
    // Scale/skew are deferred with an explicit diagnostic ...
    expect(style.issues.some(i => i.kind === 'unsupported-effect' && i.feature === 'outerShdw-scale')).toBe(true)
    // ... while the shadow still renders as the pure dist/dir translation.
    expect(style.shadow?.offsetX).toBeCloseTo(10, 1)
    expect(style.shadow?.offsetY).toBeCloseTo(0, 1)
    expect(computeShadowOffset(style.shadow!)).toEqual({ offsetX: style.shadow!.offsetX, offsetY: style.shadow!.offsetY })
    const canvas = createCanvas(100, 100)
    const ctx = canvas.getContext('2d')
    const issues = paintGeometry(ctx as never, resolvePreset('rect', 20, 20), style, 20, 20)
    expect(issues).toEqual([])

    // Default 100% / 0 scale+skew diagnose nothing.
    const xmlClean = parseXmlOrdered(
      `<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `  <a:effectLst>` +
      `    <a:outerShdw dist="95250" dir="0" blurRad="0" sx="100000" sy="100000" kx="0" ky="0">` +
      `      <a:srgbClr val="000000"/>` +
      `    </a:outerShdw>` +
      `  </a:effectLst>` +
      `</a:spPr>`
    )
    const clean = resolveDrawingStyle(xmlClean, undefined, dummyTheme)
    expect(clean.issues.some(i => i.kind === 'unsupported-effect' && i.feature === 'outerShdw-scale')).toBe(false)
  })

  test('when rotWithShape="0", shape rotation does not alter shadow vector angle in CTM', () => {
    const geom = resolvePreset('rect', 20, 20)
    const pixelAt = (ctx: CanvasRenderingContext2D, x: number, y: number) => {
      const data = ctx.getImageData(0, 0, 100, 100).data
      const idx = (y * 100 + x) * 4
      return [data[idx], data[idx + 1], data[idx + 2], data[idx + 3]]
    }

    // 1. With rotWithShape: false
    // Shape is 20x20 at (40..60, 40..60).
    // Rotated 90 degrees around center (50, 50).
    // Local offset is (0, 20) (downward).
    // Because rotWithShape is false, shadow stays downward in device coords (40..60, 60..80).
    const canvasNoRot = createCanvas(100, 100)
    const ctxNoRot = canvasNoRot.getContext('2d')
    ctxNoRot.translate(50, 50)
    ctxNoRot.rotate(Math.PI / 2)
    ctxNoRot.translate(-50, -50)
    ctxNoRot.translate(40, 40)

    const styleNoRot: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } },
      shadow: {
        color: { r: 0, g: 0, b: 0, a: 1 },
        blurPx: 0,
        offsetX: 0,
        offsetY: 20,
        rotWithShape: false,
      },
      issues: [],
    }
    paintGeometry(ctxNoRot as never, geom, styleNoRot, 20, 20)

    // Downward in device space (50, 70) HAS shadow ink
    const shadowBelow = pixelAt(ctxNoRot as never, 50, 70)
    expect(shadowBelow[3]).toBeGreaterThan(100)
    expect(shadowBelow[0]).toBeLessThan(50)

    // Leftward in device space (30, 50) has ZERO ink
    const shadowLeft = pixelAt(ctxNoRot as never, 30, 50)
    expect(shadowLeft[3]).toBe(0)

    // 2. By contrast, with rotWithShape: true (default)
    // 90 deg rotation rotates local downward (0, 20) to device leftward (-20, 0).
    const canvasWithRot = createCanvas(100, 100)
    const ctxWithRot = canvasWithRot.getContext('2d')
    ctxWithRot.translate(50, 50)
    ctxWithRot.rotate(Math.PI / 2)
    ctxWithRot.translate(-50, -50)
    ctxWithRot.translate(40, 40)

    const styleWithRot: DrawingStyle = {
      fill: { kind: 'solid', color: { r: 255, g: 0, b: 0, a: 1 } },
      shadow: {
        color: { r: 0, g: 0, b: 0, a: 1 },
        blurPx: 0,
        offsetX: 0,
        offsetY: 20,
        rotWithShape: true,
      },
      issues: [],
    }
    paintGeometry(ctxWithRot as never, geom, styleWithRot, 20, 20)

    // Leftward in device space (30, 50) HAS shadow ink
    const shadowLeftWithRot = pixelAt(ctxWithRot as never, 30, 50)
    expect(shadowLeftWithRot[3]).toBeGreaterThan(100)
    expect(shadowLeftWithRot[0]).toBeLessThan(50)

    // Downward in device space (50, 70) has ZERO ink
    const shadowBelowWithRot = pixelAt(ctxWithRot as never, 50, 70)
    expect(shadowBelowWithRot[3]).toBe(0)
  })
})
