import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'
import { paintTextBody, type WarpVisualCell } from '../src/drawing/text-paint'
import { RECORD_TEXT, type LogicalTextRange } from '../src/core/text-recording'
import type { DrawingTextBody } from '../src/drawing/text'

const identityFont = (family: string): string => family

function parseWarp(preset: string, adjVal = 55555, text = 'Official Slant'): DrawingTextBody {
  const xml = `
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr><a:prstTxWarp prst="${preset}"><a:avLst><a:gd name="adj" fmla="val ${adjVal}"/></a:avLst></a:prstTxWarp></a:bodyPr>
  <a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="200000"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r></a:p>
</a:txBody>`
  const body = parseTextBody(parseXmlOrdered(xml))
  body.insetLeftEmu = 0
  body.insetRightEmu = 0
  body.insetTopEmu = 0
  body.insetBottomEmu = 0
  return body
}

function segmentsCross(
  p1: [number, number],
  p2: [number, number],
  p3: [number, number],
  p4: [number, number],
): boolean {
  const ccw = (a: [number, number], b: [number, number], c: [number, number]): number =>
    (c[1] - a[1]) * (b[0] - a[0]) - (b[1] - a[1]) * (c[0] - a[0])
  return (
    ccw(p1, p3, p4) * ccw(p2, p3, p4) < 0 &&
    ccw(p1, p2, p3) * ccw(p1, p2, p4) < 0
  )
}

function isConvexQuad(cell: WarpVisualCell): boolean {
  let sign = 0
  for (let i = 0; i < 4; i++) {
    const p = cell[i]
    const q = cell[(i + 1) % 4]
    const next = cell[(i + 2) % 4]
    const cross = (q[0] - p[0]) * (next[1] - q[1]) - (q[1] - p[1]) * (next[0] - q[0])
    if (Math.abs(cross) <= 1e-9) continue
    const s = Math.sign(cross)
    if (sign !== 0 && s !== sign) return false
    sign = s
  }
  return true
}

describe('Warp R3 mesh cell ordering and bowtie prevention', () => {
  it('textSlantUp public visual cells preserve x-correspondence and form non-bowtie quads', () => {
    const body = parseWarp('textSlantUp', 55555, 'Official Slant')
    const canvas = createCanvas(300, 150)
    const ctx = canvas.getContext('2d') as any
    let visual: LogicalTextRange['visual']
    ctx[RECORD_TEXT] = (_text: string, _x: number, _y: number, _w: number, logical: LogicalTextRange) => {
      if (logical?.visual) visual = logical.visual
    }
    paintTextBody(body, ctx as never, 10, 10, 200, 80, identityFont)

    expect(visual).toBeDefined()
    const cells = visual?.cells
    expect(cells).toBeDefined()
    expect(cells!.length).toBeGreaterThan(0)

    for (let i = 0; i < cells!.length; i++) {
      const [topLeft, topRight, bottomRight, bottomLeft] = cells![i]

      // 1. Top and bottom boundaries must both progress left-to-right
      expect(topRight[0]).toBeGreaterThan(topLeft[0])
      expect(bottomRight[0]).toBeGreaterThan(bottomLeft[0])

      // 2. Exact x-correspondence: official Slant formula preserves x exactly
      // (x is strictly a function of parameter u, independent of v).
      expect(Math.abs(topLeft[0] - bottomLeft[0])).toBeLessThanOrEqual(1e-6)
      expect(Math.abs(topRight[0] - bottomRight[0])).toBeLessThanOrEqual(1e-6)

      // 3. Left and right quad boundaries must not cross each other (no bowtie)
      const bowtied = segmentsCross(topLeft, bottomLeft, topRight, bottomRight)
      expect(bowtied).toBe(false)

      // 4. Each quad must be strictly convex
      expect(isConvexQuad(cells![i])).toBe(true)
    }
  })
})
