import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'
import { paintTextBody } from '../src/drawing/text-paint'
import { buildTextIndex } from '../src/core/search'
import { hitTest } from '../src/core/selection'

function makeBody(preset: string, text = 'MMMM', sz = 2400) {
  return parseTextBody(
    parseXmlOrdered(
      `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:bodyPr anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0" wrap="none">` +
      `<a:prstTxWarp prst="${preset}"/>` +
      `</a:bodyPr>` +
      `<a:p><a:r><a:rPr sz="${sz}"/><a:t>${text}</a:t></a:r></a:p>` +
      `</a:txBody>`
    )
  )
}

describe('U2: visual warped spans must not enter charAtX fast path during drag', () => {
  it('U2.1: centered textCircle ABCD opaque ink gives identical carets for strict and drag', async () => {
    const body = makeBody('textCircle', 'ABCD', 2400)
    const W = 400, H = 300
    const canvas = createCanvas(W, H)
    const ctx = canvas.getContext('2d') as any
    paintTextBody(body, ctx, 40, 100, 200, 100, (x: string) => x)

    const index = await buildTextIndex([
      {
        spec: { widthPx: W, heightPx: H },
        paint: (replayCtx: any) => paintTextBody(body, replayCtx, 40, 100, 200, 100, (x: string) => x),
      },
    ])

    const pixels = ctx.getImageData(0, 0, W, H).data
    const sampledInk: Array<{ x: number; y: number }> = []
    for (let yy = 100; yy < 200; yy++) {
      for (let xx = 40; xx < 240; xx++) {
        const alpha = pixels[(yy * W + xx) * 4 + 3]
        if (alpha >= 210) {
          sampledInk.push({ x: xx + 0.5, y: yy + 0.5 })
        }
      }
    }
    expect(sampledInk.length).toBeGreaterThan(0)

    for (const pt of sampledInk) {
      const strictHit = hitTest(index, 0, pt.x, pt.y, { strict: true })
      const dragHit = hitTest(index, 0, pt.x, pt.y, { strict: false })
      expect(strictHit).toBeDefined()
      expect(dragHit).toBeDefined()
      expect(dragHit!.charIndex).toBe(strictHit!.charIndex)
    }
  })

  it('U2.2: mixed page retains ordinary drag on plain line and curved hit on warped line', async () => {
    const plainBody = parseTextBody(
      parseXmlOrdered(
        `<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/>` +
        `<a:p><a:r><a:rPr sz="2400"/><a:t>Plain Text Here</a:t></a:r></a:p>` +
        `</a:txBody>`
      )
    )
    const warpedBody = makeBody('textCircle', 'ABCD', 2400)

    const W = 400, H = 400
    const probeCanvas = createCanvas(W, H)
    const probeCtx = probeCanvas.getContext('2d') as any
    paintTextBody(warpedBody, probeCtx, 40, 150, 200, 100, (x: string) => x)
    const probeData = probeCtx.getImageData(0, 0, W, H).data

    let samplePt: { x: number; y: number } | undefined
    for (let yy = 150; yy < 250; yy++) {
      for (let xx = 40; xx < 240; xx++) {
        if (probeData[(yy * W + xx) * 4 + 3] >= 210) {
          samplePt = { x: xx + 0.5, y: yy + 0.5 }
          break
        }
      }
      if (samplePt) break
    }
    expect(samplePt).toBeDefined()

    const index = await buildTextIndex([
      {
        spec: { widthPx: W, heightPx: H },
        paint: (ctx: any) => {
          paintTextBody(plainBody, ctx, 20, 20, 300, 40, (x: string) => x)
          paintTextBody(warpedBody, ctx, 40, 150, 200, 100, (x: string) => x)
        },
      },
    ])

    // Ordinary drag near plain text line
    const plainDrag = hitTest(index, 0, 50, 30, { strict: false })
    expect(plainDrag).toBeDefined()
    expect(plainDrag?.lineIndex).toBe(0)

    // Drag on visible warped ink
    const warpedStrict = hitTest(index, 0, samplePt!.x, samplePt!.y, { strict: true })
    const warpedDrag = hitTest(index, 0, samplePt!.x, samplePt!.y, { strict: false })
    expect(warpedStrict).toBeDefined()
    expect(warpedDrag).toBeDefined()
    expect(warpedDrag!.charIndex).toBe(warpedStrict!.charIndex)
  })

  it('U2.3: off-ink drag on warped text within clip returns defined nearest caret', async () => {
    const body = makeBody('textCircle', 'ABCD', 2400)
    const index = await buildTextIndex([
      {
        spec: { widthPx: 400, heightPx: 300 },
        paint: (ctx: any) => paintTextBody(body, ctx, 40, 100, 200, 100, (x: string) => x, undefined, { clip: { x: 40, y: 100, width: 200, height: 100 } }),
      },
    ])
    // Point inside clip box (40..240, 100..200) but in the middle blank opening of the circle
    const dragInBlank = hitTest(index, 0, 140, 150, { strict: false })
    expect(dragInBlank).toBeDefined()
  })

  it('U2.4: overlapping visual spans rank by nearest cluster', () => {
    const q = [{ x: 0, y: 0 }, { x: 100, y: 0 }, { x: 100, y: 10 }, { x: 0, y: 10 }]
    const source = { text: 'AB', scope: {}, order: 0 }
    const span = (text: string, start: number, cx: number) => ({
      text, x: start * 20, y: 5, width: 20, fontSize: 10,
      logical: { source, start, end: start + 1, graphemeBoundaries: [start, start + 1] },
      placement: {
        x: start * 20, y: 5, width: 20, top: 0, bottom: 10,
        transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
        visual: { polygon: q, cells: [q as any], clusters: [{ x: cx, y: 5 }] },
      },
    })
    const index: any = {
      pages: [{ index: 0, lines: [{ text: 'AB', top: 0, bottom: 10, y: 5, spans: [span('A', 0, 1), span('B', 1, 49)] }] }],
    }
    const hit = hitTest(index, 0, 50, 5, { strict: true })
    expect(hit).toBeDefined()
    expect(hit?.charIndex).toBe(2)
  })
})

describe('U3: collapsed mesh triangles must not select unbounded blank page points', () => {
  it('U3.1: textSlantUp 20 Ms collapsed tris at x=140 reject distant point (140, 500)', async () => {
    const body = makeBody('textSlantUp', 'M'.repeat(20), 2400)
    const index = await buildTextIndex([
      {
        spec: { widthPx: 600, heightPx: 600 },
        paint: (ctx: any) => paintTextBody(body, ctx, 40, 40, 100, 100, (x: string) => x),
      },
    ])

    // Far blank point along collapsed line at x=140, y=500
    const farHit = hitTest(index, 0, 140, 500, { strict: true })
    expect(farHit).toBeUndefined()
  })

  it('U3.2: valid triangles retain positive hits on actual painted ink', async () => {
    const body = makeBody('textSlantUp', 'MMMM', 2400)
    const W = 400, H = 300
    const canvas = createCanvas(W, H)
    const ctx = canvas.getContext('2d') as any
    paintTextBody(body, ctx, 40, 40, 200, 100, (x: string) => x)
    const data = ctx.getImageData(0, 0, W, H).data

    let sampleInk: { x: number; y: number } | undefined
    for (let yy = 40; yy < 140; yy++) {
      for (let xx = 40; xx < 240; xx++) {
        if (data[(yy * W + xx) * 4 + 3] >= 210) {
          sampleInk = { x: xx + 0.5, y: yy + 0.5 }
          break
        }
      }
      if (sampleInk) break
    }
    expect(sampleInk).toBeDefined()

    const index = await buildTextIndex([
      {
        spec: { widthPx: W, heightPx: H },
        paint: (replayCtx: any) => paintTextBody(body, replayCtx, 40, 40, 200, 100, (x: string) => x),
      },
    ])
    const hit = hitTest(index, 0, sampleInk!.x, sampleInk!.y, { strict: true })
    expect(hit).toBeDefined()
  })
})

describe('U4: scratch surface allocation fallback checks returned surface and usable 2D context', () => {
  it('U4.1: null-context allocation failure in recording preflight falls back without visual metadata', async () => {
    const { RECORD_TEXT } = await import('../src/core/text-recording')
    const body = makeBody('textSlantUp', 'MMMM', 2400)
    const c = createCanvas(300, 200)
    class NullContextCanvas {
      constructor(w: number, h: number) {
        if (w > 4 || h > 4) {
          return { getContext: () => null } as any
        }
        return createCanvas(w, h) as any
      }
    }
    Object.defineProperty(c, 'constructor', { value: NullContextCanvas })
    const ctx: any = c.getContext('2d')
    const records: any[] = []
    ctx[RECORD_TEXT] = (text: string, x: number, y: number, width: number, logical: any) =>
      records.push({ text, hasVisual: !!logical.visual, x, y, width })

    paintTextBody(body, ctx, 0, 0, 200, 100, (f: string) => f)
    expect(records.length).toBeGreaterThan(0)
    expect(records[0].hasVisual).toBe(false)
    expect(body.diagnostics?.some(d => d.feature === 'warp-alloc-failure')).toBe(true)
    expect(body.diagnostics?.some(d => d.feature === 'warp-budget')).toBe(false)
  })

  it('U4.2: thrown-context allocation failure in recording preflight falls back without visual metadata', async () => {
    const { RECORD_TEXT } = await import('../src/core/text-recording')
    const body = makeBody('textSlantUp', 'MMMM', 2400)
    const c = createCanvas(300, 200)
    class ThrowContextCanvas {
      constructor(w: number, h: number) {
        if (w > 4 || h > 4) {
          return { getContext: () => { throw new Error('context unavailable') } } as any
        }
        return createCanvas(w, h) as any
      }
    }
    Object.defineProperty(c, 'constructor', { value: ThrowContextCanvas })
    const ctx: any = c.getContext('2d')
    const records: any[] = []
    ctx[RECORD_TEXT] = (text: string, x: number, y: number, width: number, logical: any) =>
      records.push({ text, hasVisual: !!logical.visual, x, y, width })

    paintTextBody(body, ctx, 0, 0, 200, 100, (f: string) => f)
    expect(records.length).toBeGreaterThan(0)
    expect(records[0].hasVisual).toBe(false)
    expect(body.diagnostics?.some(d => d.feature === 'warp-alloc-failure')).toBe(true)
  })

  it('U4.3: paint-first and index-first modes both emit truthful single allocation diagnostic', async () => {
    const { RECORD_TEXT } = await import('../src/core/text-recording')
    const body = makeBody('textSlantUp', 'MMMM', 2400)
    const c = createCanvas(300, 200)
    class NullContextCanvas {
      constructor(w: number, h: number) {
        if (w > 4 || h > 4) {
          return { getContext: () => null } as any
        }
        return createCanvas(w, h) as any
      }
    }
    Object.defineProperty(c, 'constructor', { value: NullContextCanvas })
    const ctx: any = c.getContext('2d')

    // Paint first
    paintTextBody(body, ctx, 0, 0, 200, 100, (f: string) => f)
    const paintDiags = (body.diagnostics ?? []).filter(d => d.feature === 'warp-alloc-failure')
    expect(paintDiags.length).toBe(1)

    // Then record
    const records: any[] = []
    ctx[RECORD_TEXT] = (text: string, x: number, y: number, width: number, logical: any) =>
      records.push({ text, hasVisual: !!logical.visual, x, y, width })
    paintTextBody(body, ctx, 0, 0, 200, 100, (f: string) => f)
    expect(records[0].hasVisual).toBe(false)
    const totalDiags = (body.diagnostics ?? []).filter(d => d.feature === 'warp-alloc-failure')
    expect(totalDiags.length).toBe(1)
  })
})
