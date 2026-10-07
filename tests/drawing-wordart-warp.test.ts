import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'

function createSpyContext(w = 300, h = 100, overrides: Record<string | symbol, unknown> = {}) {
  const canvas = createCanvas(w, h)
  const ctx = canvas.getContext('2d')
  return new Proxy(ctx as unknown as object, {
    get(target, prop) {
      if (prop in overrides) {
        return overrides[prop]
      }
      const val = (target as Record<string | symbol, unknown>)[prop]
      return typeof val === 'function' ? val.bind(target) : val
    },
    set(target, prop, value) {
      if (prop in overrides) {
        overrides[prop] = value
        return true
      }
      ;(target as Record<string | symbol, unknown>)[prop] = value
      return true
    },
  }) as unknown as CanvasRenderingContext2D
}

const txBodyXml = (bodyPrInner: string, text = 'WordArt Warp') => `
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr>
    ${bodyPrInner}
  </a:bodyPr>
  <a:p>
    <a:r>
      <a:t>${text}</a:t>
    </a:r>
  </a:p>
</a:txBody>`

describe('WordArt Preset Text Warp Parsing & Modeling (<a:prstTxWarp>) - Phase 4', () => {
  it('parses preset text warp type (prst) from <a:bodyPr><a:prstTxWarp> into DrawingTextBody', () => {
    const xml = txBodyXml('<a:prstTxWarp prst="textArchUp"/>')
    const parsed = parseTextBody(parseXmlOrdered(xml))

    expect(parsed.textWarp).toBeDefined()
    expect(parsed.textWarp?.preset).toBe('textArchUp')
  })

  it('parses adjustment values (<a:avLst><a:gd>) for preset text warps with appropriate units (angles vs percentage)', () => {
    // Angle adjustment: 5400000 (90 degrees in 60000ths of a degree)
    const archXml = txBodyXml(`
      <a:prstTxWarp prst="textArchUp">
        <a:avLst>
          <a:gd name="adj" fmla="val 5400000"/>
        </a:avLst>
      </a:prstTxWarp>
    `)
    const parsedArch = parseTextBody(parseXmlOrdered(archXml))
    expect(parsedArch.textWarp?.preset).toBe('textArchUp')
    expect(parsedArch.textWarp?.adjustments?.adj).toBe(5400000)

    // Percentage adjustment: 25000 (25% in 100000ths)
    const inflateXml = txBodyXml(`
      <a:prstTxWarp prst="textInflate">
        <a:avLst>
          <a:gd name="adj" fmla="val 25000"/>
        </a:avLst>
      </a:prstTxWarp>
    `)
    const parsedInflate = parseTextBody(parseXmlOrdered(inflateXml))
    expect(parsedInflate.textWarp?.preset).toBe('textInflate')
    expect(parsedInflate.textWarp?.adjustments?.adj).toBe(25000)

    // Multi-guide adjustment (e.g. wave presets with adj1 and adj2)
    const waveXml = txBodyXml(`
      <a:prstTxWarp prst="textWave1">
        <a:avLst>
          <a:gd name="adj1" fmla="val 10000"/>
          <a:gd name="adj2" fmla="val 35000"/>
        </a:avLst>
      </a:prstTxWarp>
    `)
    const parsedWave = parseTextBody(parseXmlOrdered(waveXml))
    expect(parsedWave.textWarp?.preset).toBe('textWave1')
    expect(parsedWave.textWarp?.adjustments?.adj1).toBe(10000)
    expect(parsedWave.textWarp?.adjustments?.adj2).toBe(35000)
  })

  it('applies default adjustment table for presets when <a:avLst> is omitted', () => {
    // textArchUp default adj is 10800000 (180 degrees)
    const archXml = txBodyXml('<a:prstTxWarp prst="textArchUp"/>')
    const parsedArch = parseTextBody(parseXmlOrdered(archXml))
    expect(parsedArch.textWarp?.adjustments?.adj).toBe(10800000)

    // textCircle default adj is 10800000 (180 degrees)
    const circleXml = txBodyXml('<a:prstTxWarp prst="textCircle"/>')
    const parsedCircle = parseTextBody(parseXmlOrdered(circleXml))
    expect(parsedCircle.textWarp?.adjustments?.adj).toBe(10800000)

    // textWave1 defaults: adj1 = 0, adj2 = 50000 (50%)
    const waveXml = txBodyXml('<a:prstTxWarp prst="textWave1"/>')
    const parsedWave = parseTextBody(parseXmlOrdered(waveXml))
    expect(parsedWave.textWarp?.adjustments?.adj1).toBe(0)
    expect(parsedWave.textWarp?.adjustments?.adj2).toBe(50000)

    // textInflate default adj is 50000 (50%)
    const inflateXml = txBodyXml('<a:prstTxWarp prst="textInflate"/>')
    const parsedInflate = parseTextBody(parseXmlOrdered(inflateXml))
    expect(parsedInflate.textWarp?.adjustments?.adj).toBe(50000)

    // textSlantUp default adj is 25000
    const slantXml = txBodyXml('<a:prstTxWarp prst="textSlantUp"/>')
    const parsedSlant = parseTextBody(parseXmlOrdered(slantXml))
    expect(parsedSlant.textWarp?.adjustments?.adj).toBe(25000)
  })

  it('handles textNoShape and textPlain as unwarped standard text', () => {
    const noShapeXml = txBodyXml('<a:prstTxWarp prst="textNoShape"/>', 'Standard Plain Text')
    const parsedNoShape = parseTextBody(parseXmlOrdered(noShapeXml))
    expect(parsedNoShape.textWarp).toBeUndefined()
    expect(parsedNoShape.diagnostics).toBeUndefined()

    const plainXml = txBodyXml('<a:prstTxWarp prst="textPlain"/>', 'Standard Plain Text')
    const parsedPlain = parseTextBody(parseXmlOrdered(plainXml))
    expect(parsedPlain.textWarp).toBeUndefined()
    expect(parsedPlain.diagnostics).toBeUndefined()
  })

  it('emits diagnostic for unknown warp presets and falls back to unwarped text rendering without throwing', () => {
    const unknownXml = txBodyXml('<a:prstTxWarp prst="textSuperUnknownPreset"/>', 'Fallback Text')
    const parsed = parseTextBody(parseXmlOrdered(unknownXml))

    // Must not crash and fall back to unwarped text
    expect(parsed.textWarp).toBeUndefined()
    expect(parsed.paragraphs[0].runs[0].text).toBe('Fallback Text')

    // Must emit diagnostic with kind 'unsupported-text-warp'
    expect(parsed.diagnostics).toBeDefined()
    expect(parsed.diagnostics).toContainEqual({
      kind: 'unsupported-text-warp',
      feature: 'textSuperUnknownPreset',
      message: expect.stringContaining('textSuperUnknownPreset'),
    })

    // Unmodeled preset (such as textTriangle) must also emit unsupported-text-warp diagnostic
    const triangleXml = txBodyXml('<a:prstTxWarp prst="textTriangle"/>', 'Triangle Text')
    const parsedTriangle = parseTextBody(parseXmlOrdered(triangleXml))
    expect(parsedTriangle.textWarp).toBeUndefined()
    expect(parsedTriangle.diagnostics).toContainEqual(
      expect.objectContaining({
        kind: 'unsupported-text-warp',
        feature: 'textTriangle',
      })
    )
  })
})

describe('WordArt Text Warp Geometry Engine & Canvas Deformation - Phase 5', () => {
  it('computes arc curve transformation for textArchUp and textArchDown', async () => {
    const { computeWarpTransform } = await import('../src/drawing/text-warp')
    const box = { x: 0, y: 0, width: 200, height: 100 }

    // textArchUp: default 180 degrees (pi radians)
    const archUpCenter = computeWarpTransform({ preset: 'textArchUp', adjustments: { adj: 10800000 } }, box, { x: 100, y: 50 })
    const archUpLeft = computeWarpTransform({ preset: 'textArchUp', adjustments: { adj: 10800000 } }, box, { x: 50, y: 50 })
    const archUpRight = computeWarpTransform({ preset: 'textArchUp', adjustments: { adj: 10800000 } }, box, { x: 150, y: 50 })

    // Center is the apex: rotation is 0
    expect(archUpCenter.rotation).toBeCloseTo(0, 4)
    // Left tilts counter-clockwise (negative angle)
    expect(archUpLeft.rotation).toBeLessThan(0)
    // Right tilts clockwise (positive angle)
    expect(archUpRight.rotation).toBeGreaterThan(0)
    // Center y is highest (smallest y value in canvas); left and right dip down (larger y)
    expect(archUpCenter.y).toBeLessThan(archUpLeft.y)
    expect(archUpCenter.y).toBeLessThan(archUpRight.y)

    // textArchDown: curves downward
    const archDownCenter = computeWarpTransform({ preset: 'textArchDown', adjustments: { adj: 10800000 } }, box, { x: 100, y: 50 })
    const archDownLeft = computeWarpTransform({ preset: 'textArchDown', adjustments: { adj: 10800000 } }, box, { x: 50, y: 50 })
    const archDownRight = computeWarpTransform({ preset: 'textArchDown', adjustments: { adj: 10800000 } }, box, { x: 150, y: 50 })

    expect(archDownCenter.rotation).toBeCloseTo(0, 4)
    expect(archDownLeft.rotation).toBeGreaterThan(0)
    expect(archDownRight.rotation).toBeLessThan(0)
    // Center y is lowest (largest y value in canvas); left and right are higher up (smaller y)
    expect(archDownCenter.y).toBeGreaterThan(archDownLeft.y)
    expect(archDownCenter.y).toBeGreaterThan(archDownRight.y)
  })

  it('computes circular envelope transformation for textCircle', async () => {
    const { computeWarpTransform } = await import('../src/drawing/text-warp')
    const box = { x: 0, y: 0, width: 200, height: 200 }

    // Full 360 degree circle (21600000 in 60000ths of a degree)
    const warp = { preset: 'textCircle' as const, adjustments: { adj: 21600000 } }

    const topPoint = computeWarpTransform(warp, box, { x: 0, y: 100 }) // u = 0 (top: angle = -pi/2)
    const rightPoint = computeWarpTransform(warp, box, { x: 50, y: 100 }) // u = 0.25 (right: angle = 0)
    const bottomPoint = computeWarpTransform(warp, box, { x: 100, y: 100 }) // u = 0.5 (bottom: angle = pi/2)
    const leftPoint = computeWarpTransform(warp, box, { x: 150, y: 100 }) // u = 0.75 (left: angle = pi)

    // Center of circle is at (100, 100), radius is 100
    // Top point (u = 0): x ~ 100, y ~ 0
    expect(topPoint.x).toBeCloseTo(100, 1)
    expect(topPoint.y).toBeCloseTo(0, 1)

    // Right point (u = 0.25): x ~ 200, y ~ 100
    expect(rightPoint.x).toBeCloseTo(200, 1)
    expect(rightPoint.y).toBeCloseTo(100, 1)

    // Bottom point (u = 0.5): x ~ 100, y ~ 200
    expect(bottomPoint.x).toBeCloseTo(100, 1)
    expect(bottomPoint.y).toBeCloseTo(200, 1)

    // Left point (u = 0.75): x ~ 0, y ~ 100
    expect(leftPoint.x).toBeCloseTo(0, 1)
    expect(leftPoint.y).toBeCloseTo(100, 1)
  })

  it('maintains seam continuity where start meets end for textCircle', async () => {
    const { computeWarpTransform } = await import('../src/drawing/text-warp')
    const box = { x: 0, y: 0, width: 200, height: 200 }
    const warp = { preset: 'textCircle' as const, adjustments: { adj: 21600000 } }

    // Start of circle (u = 0) vs end of circle (u = 1)
    const start = computeWarpTransform(warp, box, { x: 0, y: 100 })
    const end = computeWarpTransform(warp, box, { x: 200, y: 100 })

    // Distance between start and end position must be 0 (continuous seam)
    expect(Math.abs(start.x - end.x)).toBeLessThan(1e-4)
    expect(Math.abs(start.y - end.y)).toBeLessThan(1e-4)

    // Rotation at end is 2*pi, which is angularly identical to start rotation (0)
    const rotationDiff = Math.abs(end.rotation - (start.rotation + 2 * Math.PI))
    expect(rotationDiff).toBeLessThan(1e-4)
  })

  it('computes vertical sine wave baseline displacement for textWave1 and textWave2', async () => {
    const { computeWarpTransform } = await import('../src/drawing/text-warp')
    const box = { x: 0, y: 0, width: 200, height: 100 }

    // Wave 1 with adj2 = 50000 (50% amplitude)
    const wave1 = { preset: 'textWave1' as const, adjustments: { adj2: 50000 } }

    const p0 = computeWarpTransform(wave1, box, { x: 0, y: 50 }) // u = 0 -> sin(0) = 0
    const pQuarter = computeWarpTransform(wave1, box, { x: 50, y: 50 }) // u = 0.25 -> sin(pi/2) = 1 (positive peak)
    const pHalf = computeWarpTransform(wave1, box, { x: 100, y: 50 }) // u = 0.5 -> sin(pi) = 0
    const pThreeQuarter = computeWarpTransform(wave1, box, { x: 150, y: 50 }) // u = 0.75 -> sin(3pi/2) = -1 (negative peak)

    expect(p0.y).toBeCloseTo(50, 2)
    expect(pQuarter.y).toBeGreaterThan(50)
    expect(pHalf.y).toBeCloseTo(50, 2)
    expect(pThreeQuarter.y).toBeLessThan(50)

    // Wave 2 is inverted phase
    const wave2 = { preset: 'textWave2' as const, adjustments: { adj2: 50000 } }
    const pQuarter2 = computeWarpTransform(wave2, box, { x: 50, y: 50 })
    const pThreeQuarter2 = computeWarpTransform(wave2, box, { x: 150, y: 50 })

    expect(pQuarter2.y).toBeLessThan(50)
    expect(pThreeQuarter2.y).toBeGreaterThan(50)
  })

  it('computes envelope height scaling for textInflate and textDeflate', async () => {
    const { computeWarpTransform } = await import('../src/drawing/text-warp')
    const box = { x: 0, y: 0, width: 200, height: 100 }

    const inflate = { preset: 'textInflate' as const, adjustments: { adj: 50000 } }
    const deflate = { preset: 'textDeflate' as const, adjustments: { adj: 50000 } }

    // Middle point (u = 0.5)
    const inflateMid = computeWarpTransform(inflate, box, { x: 100, y: 80 }) // 30 units below midY (50)
    const deflateMid = computeWarpTransform(deflate, box, { x: 100, y: 80 })

    // For inflate, scaleY at mid is > 1 (envelope expanded)
    expect(inflateMid.scaleY).toBeGreaterThan(1.0)
    expect(inflateMid.y).toBeGreaterThan(80)

    // For deflate, scaleY at mid is < 1 (envelope compressed)
    expect(deflateMid.scaleY).toBeLessThan(1.0)
    expect(deflateMid.y).toBeLessThan(80)
  })

  it('computes affine shear transformation for textSlantUp and textSlantDown', async () => {
    const { computeWarpTransform } = await import('../src/drawing/text-warp')
    const box = { x: 0, y: 0, width: 200, height: 100 }

    const slantUp = { preset: 'textSlantUp' as const, adjustments: { adj: 25000 } }
    const slantDown = { preset: 'textSlantDown' as const, adjustments: { adj: 25000 } }

    const slantUpLeft = computeWarpTransform(slantUp, box, { x: 0, y: 50 })
    const slantUpRight = computeWarpTransform(slantUp, box, { x: 200, y: 50 })

    // textSlantUp slants up from left to right (y decreases as x increases)
    expect(slantUpRight.y).toBeLessThan(slantUpLeft.y)
    expect(slantUpLeft.rotation).toBeLessThan(0)

    const slantDownLeft = computeWarpTransform(slantDown, box, { x: 0, y: 50 })
    const slantDownRight = computeWarpTransform(slantDown, box, { x: 200, y: 50 })

    // textSlantDown slants down from left to right (y increases as x increases)
    expect(slantDownRight.y).toBeGreaterThan(slantDownLeft.y)
    expect(slantDownLeft.rotation).toBeGreaterThan(0)
  })

  it('text layout advances and line boxes remain strictly invariant under text warp', async () => {
    const { layoutTextBody } = await import('../src/drawing/text-layout')
    const measure = (text: string) => ({ width: text.length * 10, ascent: 12, descent: 4, normalHeight: 16 })

    const unwarpedBody = {
      paragraphs: [{
        runs: [{ text: 'Hello WordArt Warp Architecture', fontSizePt: 16 }],
        align: 'left' as const,
        level: 0,
      }],
      anchor: 't' as const,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      wrap: true,
    }

    const warpedBody = {
      ...unwarpedBody,
      textWarp: { preset: 'textArchUp' as const, adjustments: { adj: 10800000 } },
    }

    const layoutUnwarped = layoutTextBody(unwarpedBody, 200, 100, measure)
    const layoutWarped = layoutTextBody(warpedBody, 200, 100, measure)

    expect(layoutWarped.lines.length).toBe(layoutUnwarped.lines.length)
    expect(layoutWarped.height).toBe(layoutUnwarped.height)

    for (let i = 0; i < layoutUnwarped.lines.length; i++) {
      const uLine = layoutUnwarped.lines[i]
      const wLine = layoutWarped.lines[i]
      expect(wLine.x).toBe(uLine.x)
      expect(wLine.baseline).toBe(uLine.baseline)
      expect(wLine.segments.length).toBe(uLine.segments.length)
      for (let j = 0; j < uLine.segments.length; j++) {
        expect(wLine.segments[j].width).toBe(uLine.segments[j].width)
        expect(wLine.segments[j].text).toBe(uLine.segments[j].text)
        expect(wLine.segments[j].sourceStart).toBe(uLine.segments[j].sourceStart)
        expect(wLine.segments[j].sourceEnd).toBe(uLine.segments[j].sourceEnd)
      }
    }
  })

  it('canvas paints warped text along transform curves while maintaining stroke and fill styling', async () => {
    const { paintTextBody } = await import('../src/drawing/text-paint')

    const rotations: number[] = []
    const filledTexts: string[] = []
    const strokedTexts: string[] = []

    const mockCtx = createSpyContext(200, 100, {
      rotate: (rad: number) => { rotations.push(rad) },
      fillText: (text: string) => { filledTexts.push(text) },
      strokeText: (text: string) => { strokedTexts.push(text) },
    })

    const body = {
      paragraphs: [{
        runs: [{
          text: 'WARP',
          fontSizePt: 20,
          color: '#ff0000',
          textOutline: { color: '#000000', widthPx: 2 },
        }],
        align: 'center' as const,
        level: 0,
      }],
      anchor: 't' as const,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      wrap: true,
      textWarp: { preset: 'textArchUp' as const, adjustments: { adj: 10800000 } },
    }

    paintTextBody(body, mockCtx, 0, 0, 200, 100, f => f)

    // Glyphs 'W', 'A', 'R', 'P' should each be filled and stroked
    expect(filledTexts).toEqual(['W', 'A', 'R', 'P'])
    expect(strokedTexts).toEqual(['W', 'A', 'R', 'P'])
    // Rotations should be recorded along the arc
    expect(rotations.length).toBe(4)
    // Leftmost glyph rotates negatively, rightmost glyph rotates positively
    expect(rotations[0]).toBeLessThan(0)
    expect(rotations[3]).toBeGreaterThan(0)
  })

  it('warped text transforms gradient and pattern fill coordinate spaces with glyph bounds', async () => {
    const { paintTextBody } = await import('../src/drawing/text-paint')

    const gradientBoxes: Array<{ x0: number; y0: number; x1: number; y1: number; extent: number }> = []

    const mockCtx = createSpyContext(200, 100, {
      createLinearGradient: (x0: number, y0: number, x1: number, y1: number) => {
        gradientBoxes.push({ x0, y0, x1, y1, extent: Math.hypot(x1 - x0, y1 - y0) })
        return { addColorStop: () => {} }
      },
    })

    const body = {
      paragraphs: [{
        runs: [{
          text: 'GRAD',
          fontSizePt: 20,
          textFill: {
            kind: 'gradient' as const,
            angle: 90,
            stops: [{ position: 0, color: '#ff0000' }, { position: 1, color: '#0000ff' }],
          },
        }],
        align: 'left' as const,
        level: 0,
      }],
      anchor: 't' as const,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      wrap: true,
      textWarp: { preset: 'textArchUp' as const, adjustments: { adj: 10800000 } },
    }

    paintTextBody(body, mockCtx, 0, 0, 200, 100, f => f)

    expect(gradientBoxes.length).toBe(4)
    // Each gradient box is sized to its glyph bounds in transformed coordinate space
    for (const box of gradientBoxes) {
      expect(box.extent).toBeGreaterThanOrEqual(0.5)
    }
  })


  it('text warp on vertical text (vert/wordArtVert) applies in local rotated frame preserving column progression', async () => {
    const { paintTextBody } = await import('../src/drawing/text-paint')

    const paintedGlyphs: string[] = []
    const transforms: number[][] = []
    const translates: number[][] = []

    const mockCtx = createSpyContext(100, 200, {
      transform: (a: number, b: number, c: number, d: number, e: number, f: number) => {
        transforms.push([a, b, c, d, e, f])
      },
      translate: (x: number, y: number) => {
        translates.push([x, y])
      },
      fillText: (text: string) => { paintedGlyphs.push(text) },
    })

    const body = {
      paragraphs: [{
        runs: [{ text: 'AB', fontSizePt: 16, color: '#000000' }],
        align: 'left' as const,
        level: 0,
      }],
      direction: 'vert' as const,
      anchor: 't' as const,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      wrap: true,
      textWarp: { preset: 'textArchUp' as const, adjustments: { adj: 10800000 } },
    }

    paintTextBody(body, mockCtx, 0, 0, 100, 200, f => f)

    expect(paintedGlyphs).toEqual(['A', 'B'])
    // Vertical text segment transforms establish rotated frame
    expect(transforms.length).toBeGreaterThan(0)
    // Per-glyph warp translations follow reading flow order down the column
    expect(translates.length).toBe(2)
    // Discriminator: bounded transverse displacement (old broken code had negative x ~ -94 and inflated y ~ 218)
    expect(translates[0][0]).toBeGreaterThan(0)
    expect(translates[0][1]).toBeLessThan(100)
    expect(translates[1][0]).toBeGreaterThan(translates[0][0])
    expect(translates[0][1] - translates[1][1]).toBeGreaterThan(10)

    // Multi-segment flow offset discriminator: two spaced runs down the column must incorporate
    // their flow offset into warp coordinates, whereas buggy code dropping segment.x produced near-identical offsets
    const spacedTranslates: number[][] = []
    const spacedCtx = createSpyContext(100, 200, {
      translate: (x: number, y: number) => { spacedTranslates.push([x, y]) },
    })
    const spacedBody = {
      paragraphs: [{
        runs: [
          { text: 'A', fontSizePt: 16, color: '#000000' },
          { text: '     ', fontSizePt: 16, color: '#000000' },
          { text: 'B', fontSizePt: 16, color: '#000000' },
        ],
        align: 'left' as const,
        level: 0,
      }],
      direction: 'vert' as const,
      anchor: 't' as const,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      wrap: true,
      textWarp: { preset: 'textArchUp' as const, adjustments: { adj: 10800000 } },
    }
    paintTextBody(spacedBody, spacedCtx, 0, 0, 100, 200, f => f)
    expect(spacedTranslates.length).toBe(7)
    // Flow offset between A (glyph 0) and B (glyph 6) across spaces must reflect transverse progression (>20px shift)
    expect(Math.abs(spacedTranslates[6][0] - spacedTranslates[0][0])).toBeGreaterThan(20)
    expect(Math.abs(spacedTranslates[6][1] - spacedTranslates[0][1])).toBeGreaterThan(30)
  })

  it('clampAdjustment clamps angle, percentage, and handles NaN defensively', async () => {
    const { clampAdjustment } = await import('../src/drawing/text-warp')
    // Angle family [0, 21600000]
    expect(clampAdjustment('textArchUp', 'adj', 99999999)).toBe(21600000)
    expect(clampAdjustment('textCircle', 'adj', -500)).toBe(0)
    expect(clampAdjustment('textCircle', 'adj', 10800000)).toBe(10800000)

    // Percentage family [0, 100000]
    expect(clampAdjustment('textSlantUp', 'adj', 500000)).toBe(100000)
    expect(clampAdjustment('textWave1', 'adj2', -100)).toBe(0)
    expect(clampAdjustment('textInflate', 'adj', 25000)).toBe(25000)

    // NaN defensive guard
    expect(clampAdjustment('textArchUp', 'adj', NaN)).toBe(0)
    expect(clampAdjustment('textWave1', 'adj2', NaN)).toBe(0)
  })

  it('modeled but ungeometrized presets safely fall back to unwarped rendering', async () => {
    const { paintTextBody } = await import('../src/drawing/text-paint')
    const painted: string[] = []
    const mockCtx = createSpyContext(200, 100, {
      fillText: (text: string) => { painted.push(text) },
    })
    const body = {
      paragraphs: [{
        runs: [{ text: 'FALLBACK', fontSizePt: 16 }],
        align: 'left' as const,
        level: 0,
      }],
      anchor: 't' as const,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      wrap: true,
      textWarp: { preset: 'textWave4' as const, adjustments: {} },
    }

    paintTextBody(body, mockCtx, 0, 0, 200, 100, f => f)
    expect(painted.length).toBeGreaterThan(0)
    expect(painted.join('')).toBe('FALLBACK')
  })


  it('warped glyphs exceeding line bounding boxes clip deterministically', async () => {
    const { paintTextBody } = await import('../src/drawing/text-paint')

    const clipRects: Array<{ x: number; y: number; w: number; h: number }> = []

    const mockCtx = createSpyContext(200, 100, {
      rect: (x: number, y: number, w: number, h: number) => { clipRects.push({ x, y, w, h }) },
    })

    const body = {
      paragraphs: [{
        runs: [{ text: 'CLIPPED', fontSizePt: 20, color: '#000000' }],
        align: 'left' as const,
        level: 0,
      }],
      anchor: 't' as const,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      wrap: true,
      textWarp: { preset: 'textArchUp' as const, adjustments: { adj: 10800000 } },
    }

    // Pass clipToLineBox option
    paintTextBody(body, mockCtx, 10, 20, 200, 100, f => f, undefined, { clipToLineBox: true })

    expect(clipRects.length).toBeGreaterThan(0)
    // Clip rectangle bounds match line box
    expect(clipRects[0].w).toBeGreaterThan(0)
    expect(clipRects[0].h).toBeGreaterThan(0)
  })

  it('search indexing emits unwarped layout coordinates preserving logical reading order and selection stability', async () => {
    const { paintTextBody } = await import('../src/drawing/text-paint')
    const { RECORD_TEXT } = await import('../src/core/text-recording')

    const recorded: Array<{ text: string; x: number; y: number; width: number }> = []

    const mockCtx = createSpyContext(300, 100, {
      [RECORD_TEXT]: (text: string, x: number, y: number, width: number) => {
        recorded.push({ text, x, y, width })
      },
    })

    const unwarpedBody = {
      paragraphs: [{
        runs: [{ text: 'Search Order Test', fontSizePt: 16 }],
        align: 'left' as const,
        level: 0,
      }],
      anchor: 't' as const,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      wrap: true,
    }

    const warpedBody = {
      ...unwarpedBody,
      textWarp: { preset: 'textArchUp' as const, adjustments: { adj: 10800000 } },
    }

    // Paint warped text with recording hook
    paintTextBody(warpedBody, mockCtx, 0, 0, 300, 100, f => f)

    expect(recorded.length).toBe(1)
    expect(recorded[0].text).toBe('Search Order Test')
    // Layout coordinates are linear and unwarped
    expect(recorded[0].x).toBe(0)
    expect(recorded[0].width).toBeGreaterThan(0)
  })

  it('text hit-testing along warped curves produces monotonically non-decreasing character offsets', async () => {
    const { paintTextBody } = await import('../src/drawing/text-paint')
    const { RECORD_TEXT } = await import('../src/core/text-recording')

    const offsets: number[] = []

    const mockCtx = createSpyContext(300, 100, {
      [RECORD_TEXT]: (_text: string, _x: number, _y: number, _width: number, logical: { start: number; end: number }) => {
        offsets.push(logical.start, logical.end)
      },
    })

    const body = {
      paragraphs: [{
        runs: [
          { text: 'First ', fontSizePt: 14 },
          { text: 'Second ', fontSizePt: 14 },
          { text: 'Third', fontSizePt: 14 },
        ],
        align: 'left' as const,
        level: 0,
      }],

      anchor: 't' as const,
      insetLeftEmu: 0,
      insetRightEmu: 0,
      insetTopEmu: 0,
      insetBottomEmu: 0,
      wrap: true,
      textWarp: { preset: 'textCircle' as const, adjustments: { adj: 10800000 } },
    }

    paintTextBody(body, mockCtx, 0, 0, 300, 100, f => f)

    // Character offsets must be monotonically non-decreasing
    expect(offsets.length).toBeGreaterThan(0)
    for (let i = 1; i < offsets.length; i++) {
      expect(offsets[i]).toBeGreaterThanOrEqual(offsets[i - 1])
    }
  })
})






