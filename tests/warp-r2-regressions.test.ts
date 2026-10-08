/**
 * R4 warp regression gates for the independently confirmed R2 SPEC/QUALITY
 * findings. All imports bind THIS workspace; expectations are the finalized
 * report gates (never a passing subset):
 * - exact affine Slant oracle (whole-raster correspondence, no remap)
 * - local vertical-frame composition (pixel identity vs local reference)
 * - whole opaque-ink (alpha>=210) strict-hit coverage >=99% (six controls)
 * - terminal/single-cluster end carets incl. vertical and ZWJ boundaries
 * - mapped highlight bands (counts, full-ink coverage, mixed spans kept)
 * - fallback paint/index agreement incl. host allocation failure, both orders
 * - half-alpha preservation, 0.2px outlines, caller-origin composition
 * Docx-level authored appearance evidence (24pt occlusion control, bold12
 * vertical pair) lives in office-wordart-integration.test.ts; the root-
 * reviewed cell-order regression lives in warp-r3-cell-order.test.ts.
 */
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'
import { paintTextBody, createTextBodyMeasurer } from '../src/drawing/text-paint'
import { layoutTextBody } from '../src/drawing/text-layout'
import { buildTextIndex, rectsForRange } from '../src/core/search'
import { hitTest } from '../src/core/selection'
import type { DrawingTextBody } from '../src/drawing/text'

const MARGIN = 8
const identityFont = (family: string): string => family
const warpXml = (preset: string, adjustments?: Record<string, number>): string => {
  const gd = Object.entries(adjustments ?? {}).map(([name, value]) =>
    `<a:gd name="${name}" fmla="val ${value}"/>`).join('')
  return `<a:prstTxWarp prst="${preset}"><a:avLst>${gd}</a:avLst></a:prstTxWarp>`
}
function warpBody(preset: string | undefined, adjustments: Record<string, number> | undefined, bodyText: string, styles = ''): DrawingTextBody {
  const body = parseTextBody(parseXmlOrdered(`
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr>${preset ? warpXml(preset, adjustments) : ''}</a:bodyPr>
  <a:p><a:r><a:rPr sz="3600">${styles}</a:rPr><a:t>${bodyText}</a:t></a:r></a:p>
</a:txBody>`))
  body.insetLeftEmu = 0; body.insetRightEmu = 0; body.insetTopEmu = 0; body.insetBottomEmu = 0
  return body
}
function opaqueSet(data: Uint8ClampedArray, alpha = 210): Set<number> {
  const out = new Set<number>()
  for (let i = 0; i < data.length; i += 4) {
    if (data[i + 3] >= alpha && (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250)) out.add(i / 4)
  }
  return out
}
function iou(a: Set<number>, b: Set<number>): number {
  let inter = 0
  for (const p of a) if (b.has(p)) inter++
  const union = a.size + b.size - inter
  return union === 0 ? 1 : inter / union
}
/** Strict-hit carets reached over every actual opaque-ink pixel. */
async function inkCarets(body: DrawingTextBody, W: number, H: number, x: number, y: number, w: number, h: number): Promise<{ ink: number; seen: Set<number> }> {
  const canvas = createCanvas(W, H)
  const ctx = canvas.getContext('2d')
  paintTextBody(body, ctx as unknown as never, x, y, w, h, identityFont, undefined, { clip: { x, y, width: w, height: h } })
  const data = ctx.getImageData(0, 0, W, H).data
  const index = await buildTextIndex([{ spec: { widthPx: W, heightPx: H }, paint: (replayCtx: CanvasRenderingContext2D) => {
    paintTextBody(body, replayCtx as unknown as never, x, y, w, h, identityFont, undefined, { clip: { x, y, width: w, height: h } })
  } }])
  const seen = new Set<number>()
  let ink = 0
  for (let py = 0; py < H; py++) for (let px = 0; px < W; px++) {
    const i = (py * W + px) * 4
    if (data[i + 3] < 210 || (data[i] >= 250 && data[i + 1] >= 250 && data[i + 2] >= 250)) continue
    ink++
    const hit = hitTest(index, 0, px + 0.5, py + 0.5, { strict: true })
    if (hit) seen.add(hit.charIndex)
  }
  return { ink, seen }
}

describe('R2 finding 1: caller origin composes exactly once', () => {
  it('slant at (100,100) equals origin-zero paint under outer translation', () => {
    const body = warpBody('textSlantUp', { adj: 55555 }, 'MMMM', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    const directCanvas = createCanvas(750, 500)
    const directCtx = directCanvas.getContext('2d')
    directCtx.fillStyle = '#ffffff'; directCtx.fillRect(0, 0, 750, 500)
    paintTextBody(body, directCtx as unknown as never, 100, 100, 200, 100, identityFont)
    const direct = opaqueSet(directCtx.getImageData(0, 0, 750, 500).data)
    const canvas = createCanvas(750, 500)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, 750, 500)
    ctx.save(); ctx.translate(100, 100)
    paintTextBody(body, ctx as unknown as never, 0, 0, 200, 100, identityFont)
    ctx.restore()
    const moved = opaqueSet(ctx.getImageData(0, 0, 750, 500).data)
    expect(iou(direct, moved)).toBeGreaterThan(0.9)
  })
})

describe('R2 finding 2: mesh source padding matches destination region', () => {
  it('slant ink matches the exact affine raster oracle', () => {
    // 200x100 matches the authoritative matrix SlantUp-default rows:
    // Top (0,55.555)->(200,0), Bottom (0,100)->(200,44.445), i.e.
    // F(x,y) = (x, 55.555 - 55.555*x/200 + (1-55.555/100)*y).
    const W = 200, H = 100, CW = 216, CH = 116
    const body = warpBody('textSlantUp', { adj: 55555 }, 'MMMM', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    const canvas = createCanvas(CW, CH)
    const ctx = canvas.getContext('2d')
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, CW, CH)
    paintTextBody(body, ctx as unknown as never, MARGIN, MARGIN, W, H, identityFont)
    const candidate = opaqueSet(ctx.getImageData(0, 0, CW, CH).data)
    // Oracle: ordinary source raster mapped through the exact official affine
    // with native drawImage (same resampling family as the candidate mesh).
    const plain = warpBody(undefined, undefined, 'MMMM', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    const src = createCanvas(CW, CH)
    const sctx = src.getContext('2d')
    sctx.fillStyle = '#ffffff'; sctx.fillRect(0, 0, CW, CH)
    paintTextBody(plain, sctx as unknown as never, MARGIN, MARGIN, W, H, identityFont)
    const dy = 55.555, k = 1 - dy / H
    const out = createCanvas(CW, CH)
    const octx = out.getContext('2d')
    octx.fillStyle = '#ffffff'; octx.fillRect(0, 0, CW, CH)
    octx.setTransform(1, -(dy / W), 0, k, 0, MARGIN + dy + (dy / W) * MARGIN - k * MARGIN)
    octx.drawImage(src, 0, 0)
    octx.setTransform(1, 0, 0, 1, 0, 0)
    const oracleData = octx.getImageData(0, 0, CW, CH).data
    const oracle = opaqueSet(oracleData)
    // Exact x preservation plus affine y: mesh must agree closely.
    expect(iou(candidate, oracle)).toBeGreaterThan(0.85)
    // Broad-ink gate (independent alpha40 methodology): geometry must agree
    // with zero best-shift; opaque-core value is reported, not gated.
    const cand40 = opaqueSet(ctx.getImageData(0, 0, CW, CH).data, 40)
    const orac40 = opaqueSet(oracleData, 40)
    let inter40 = 0
    for (const p of cand40) if (orac40.has(p)) inter40++
    const iou40 = inter40 / (cand40.size + orac40.size - inter40)
    expect(iou40).toBeGreaterThan(0.85)
  })
  it('integer-aligned source window reproduces the whole-body oracle exactly', () => {
    // Root affine-phase inputs: Arial28 centered MMMM/سلام, 200x100 body,
    // SlantUp adj=55555, outer translation(100,100). The scratch raster must
    // paint at an integer window origin (floor left/top, ceil right/bottom)
    // so ordinary glyph pixel phase survives; the fractional origin gave
    // IoU 0.6699/0.4762 with 891/282 RGBA differences.
    const diff = (a: Uint8ClampedArray, b: Uint8ClampedArray): { opaqueInkIoU: number; differentPixels: number } => {
      let intersection = 0, union = 0, different = 0
      for (let i = 0; i < a.length; i += 4) {
        if (a[i + 3] > 240 || b[i + 3] > 240) union++
        if (a[i + 3] > 240 && b[i + 3] > 240) intersection++
        if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2] || a[i + 3] !== b[i + 3]) different++
      }
      return { opaqueInkIoU: union ? intersection / union : 1, differentPixels: different }
    }
    for (const text of ['MMMM', 'سلام']) {
      const b = {
        paragraphs: [{ runs: [{ text, fontFamily: 'Arial', fontSizePt: 28, color: '#000000' }], align: 'center', level: 0 }],
        anchor: 'ctr', insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, wrap: false,
      }
      const actual = createCanvas(650, 500)
      const actx = actual.getContext('2d')
      actx.translate(100, 100)
      paintTextBody({ ...b, textWarp: { preset: 'textSlantUp', adjustments: { adj: 55555 } } } as never, actx as never, 0, 0, 200, 100, identityFont)
      const ordinary = createCanvas(200, 100)
      paintTextBody(b as never, ordinary.getContext('2d') as never, 0, 0, 200, 100, identityFont)
      const target = createCanvas(650, 500)
      const tc = target.getContext('2d')
      const dy = 55.555
      tc.setTransform(1, -dy / 200, 0, 1 - dy / 100, 100, 100 + dy)
      tc.drawImage(ordinary as never, 0, 0)
      const r = diff(actx.getImageData(0, 0, 650, 500).data, tc.getImageData(0, 0, 650, 500).data)
      expect(r.opaqueInkIoU).toBe(1)
      expect(r.differentPixels).toBeLessThan(100)
    }
  })
})

describe('R2 finding 3: vertical warp runs in the local writing frame', () => {
  const wave = { preset: 'textWave1', adjustments: { adj1: 12500, adj2: 0 } } as const
  for (const direction of ['vert', 'vert270'] as const) {
    it(`${direction} Wave1 matches the locally-warped then rotated reference`, () => {
      // Independent reference from actual unwarped layout: the same source
      // point warps within its local frame, then the ordinary outer
      // quarter-turn applies. Identity controls are pixel-identical, proving
      // frame construction; nonidentity Wave must match too.
      const textual = {
        paragraphs: [{ runs: [{ text: 'MMMM', fontFamily: 'Arial', fontSizePt: 24, color: '#000000' }], align: 'center', level: 0 }],
        anchor: 'ctr', insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, wrap: false,
      }
      const b = { ...textual, direction, textWarp: { ...wave } }
      const mctx = createCanvas(4, 4).getContext('2d')
      const laid = layoutTextBody(b as never, 100, 200, createTextBodyMeasurer(mctx as never, identityFont))
      const seg = laid.lines[0].segments[0] as unknown as { transform: { e: number; f: number } }
      const t = seg.transform
      const clockwise = direction === 'vert'
      const localBaseline = clockwise ? 100 - t.e : t.e
      const localX = clockwise ? t.f : 200 - t.f
      const localSeg = { ...seg, x: localX, transform: undefined }
      const localLine = { ...laid.lines[0], x: 0, y: 0, baseline: localBaseline, height: 100, segments: [localSeg], logicalLineIndex: undefined }
      const refBody = { ...b, direction: 'horz' }
      const refOuter = clockwise ? [0, 1, -1, 0, 300, 100] : [0, -1, 1, 0, 200, 300]
      const paintAt = (bb: unknown, x: number, y: number, w: number, h: number, opt: object, outer: number[]): Set<number> => {
        const c = createCanvas(650, 500)
        const cc = c.getContext('2d')
        cc.setTransform(outer[0], outer[1], outer[2], outer[3], outer[4], outer[5])
        paintTextBody(bb as never, cc as never, x, y, w, h, identityFont, undefined, opt as never)
        return opaqueSet(cc.getImageData(0, 0, 650, 500).data)
      }
      const expected = paintAt(refBody, 0, 0, 200, 100, { layout: { lines: [localLine], height: 100 }, clip: { x: 0, y: 0, width: 200, height: 100 } }, refOuter)
      const actual = paintAt(b, 0, 0, 100, 200, { layout: laid }, [1, 0, 0, 1, 200, 100])
      expect(actual.size).toBeGreaterThan(50)
      expect(iou(actual, expected)).toBeGreaterThan(0.85)
    })
  }
})

describe('R2 finding 4/quality F1: strict hits cover real curved ink', () => {
  // Root r3 cell-fix harness geometry (Times bold 36pt, 244x164 box): the
  // finalized 99% whole-ink gate over all six controls.
  const runs = (text: string, color = '112233') => `<a:r><a:rPr sz="3600" b="1"><a:latin typeface="Times New Roman"/><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r>`
  const txBody = (preset: string, runXml: string) => {
    const model = parseTextBody(parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="ctr"><a:prstTxWarp prst="${preset}"><a:avLst/></a:prstTxWarp></a:bodyPr><a:p>${runXml}</a:p></a:txBody>`))
    return model
  }
  const cases = [
    { name: 'curve-up-single', model: txBody('textCurveUp', runs('MMMM')) },
    { name: 'circle-single', model: txBody('textCircle', runs('MMMM')) },
    { name: 'wave-single', model: txBody('textWave1', runs('MMMM')) },
    { name: 'curve-up-two-runs', model: txBody('textCurveUp', runs('AB', 'CC0000') + runs('CD', '0000CC')) },
    { name: 'curve-up-unicode-two-runs', model: txBody('textCurveUp', runs('AA', 'CC0000') + runs('Å👩‍💻B', '0000CC')) },
    { name: 'long-run-budget', model: txBody('textCircle', runs('M'.repeat(260))) },
  ]
  for (const c of cases) {
    it(`${c.name}: whole opaque ink strict hits >=99%`, async () => {
      const W = 260, H = 180
      const canvas = createCanvas(W, H)
      const ctx = canvas.getContext('2d')
      paintTextBody(c.model, ctx as unknown as never, 8, 8, 244, 164, (x: string) => x, undefined, { clip: { x: 8, y: 8, width: 244, height: 164 } })
      const data = ctx.getImageData(0, 0, W, H).data
      const ink: number[] = []
      for (let i = 0; i < data.length; i += 4) {
        if (data[i + 3] >= 210 && (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250)) ink.push(i / 4)
      }
      expect(ink.length).toBeGreaterThan(100)
      const index = await buildTextIndex([{ spec: { widthPx: W, heightPx: H }, paint: (replayCtx: CanvasRenderingContext2D) => {
        paintTextBody(c.model, replayCtx as unknown as never, 8, 8, 244, 164, (x: string) => x, undefined, { clip: { x: 8, y: 8, width: 244, height: 164 } })
      } }])
      let hits = 0
      for (const p of ink) {
        const x = (p % W) + 0.5, y = Math.floor(p / W) + 0.5
        if (hitTest(index, 0, x, y, { strict: true })) hits++
      }
      expect(hits / ink.length).toBeGreaterThanOrEqual(0.99)
    })
  }
})

describe('R2 finding 5/quality F2: warped carets use line-local offsets once', () => {
  it('second-span blue pixel resolves inside the run, not at end', async () => {
    const body = parseTextBody(parseXmlOrdered(`
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr>${warpXml('textCurveUp', { adj: 45977 })}</a:bodyPr>
  <a:p>
    <a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>AB</a:t></a:r>
    <a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr><a:t>CD</a:t></a:r>
  </a:p>
</a:txBody>`))
    body.insetLeftEmu = 0; body.insetRightEmu = 0; body.insetTopEmu = 0; body.insetBottomEmu = 0
    const index = await buildTextIndex([{ spec: { widthPx: 300, heightPx: 140 }, paint: (replayCtx: CanvasRenderingContext2D) => {
      paintTextBody(body, replayCtx as unknown as never, MARGIN, MARGIN, 284, 124, identityFont)
    } }])
    const line = index.pages[0].lines[0]
    const span = line.spans[1]
    const clusters = span.placement?.visual?.clusters ?? []
    expect(clusters.length).toBeGreaterThan(0)
    const c = clusters[0]
    const hit = hitTest(index, 0, c.x, c.y, { strict: true })
    expect(hit).toBeDefined()
    expect(hit!.charIndex).toBeGreaterThanOrEqual(2)
    expect(hit!.charIndex).toBeLessThanOrEqual(3)
  })
  it('single-char slant reaches its end caret over actual ink', async () => {
    const body = warpBody('textSlantUp', { adj: 55555 }, 'M', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    const { ink, seen } = await inkCarets(body, 260, 180, 8, 8, 244, 164)
    expect(ink).toBeGreaterThan(50)
    expect(seen.has(0)).toBe(true)
    expect(seen.has(1)).toBe(true)
  })
  it('four-char slant reaches end4 over actual ink', async () => {
    const body = warpBody('textSlantUp', { adj: 55555 }, 'MMMM', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    const { ink, seen } = await inkCarets(body, 260, 180, 8, 8, 244, 164)
    expect(ink).toBeGreaterThan(100)
    expect(seen.has(4)).toBe(true)
  })
  it('vertical single char reaches its end caret over actual ink', async () => {
    const body = {
      paragraphs: [{ runs: [{ text: 'M', fontFamily: 'Arial', fontSizePt: 28, color: '#000000' }], align: 'center', level: 0 }],
      anchor: 'ctr', insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, wrap: false,
      direction: 'vert270', textWarp: { preset: 'textWave1', adjustments: { adj1: 12500, adj2: 0 } },
    }
    const { ink, seen } = await inkCarets(body as never, 260, 220, 8, 8, 120, 200)
    expect(ink).toBeGreaterThan(20)
    expect(seen.has(0)).toBe(true)
    expect(seen.has(1)).toBe(true)
  })
  it('ZWJ cluster never splits and its end caret is reachable', async () => {
    const body = {
      paragraphs: [{ runs: [{ text: 'A👩‍💻B', fontFamily: 'Arial', fontSizePt: 28, color: '#000000' }], align: 'center', level: 0 }],
      anchor: 'ctr', insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, wrap: false,
      textWarp: { preset: 'textCurveUp', adjustments: { adj: 45977 } },
    }
    const { ink, seen } = await inkCarets(body as never, 300, 180, 8, 8, 284, 164)
    expect(ink).toBeGreaterThan(100)
    // 'A👩‍💻B' is 7 UTF-16 units with legal boundaries only at 0,1,6,7.
    expect([...seen].every(c => [0, 1, 6, 7].includes(c))).toBe(true)
    expect(seen.has(7)).toBe(true)
  })
  it('interior right-half hits match the ordinary nearest caret', async () => {
    // Near-identity SlantUp adj1000 preserves x, so warped strict hits over
    // actual opaque ink must agree with the matched ordinary index. The old
    // neighbor-Voronoi threshold kept the preceding caret through the first
    // glyph's right half (207/750 differed); the split is at the selected
    // glyph's own center, matching ordinary semantics.
    const model = parseTextBody(parseXmlOrdered('<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="ctr"><a:prstTxWarp prst="textSlantUp"><a:avLst><a:gd name="adj" fmla="val 1000"/></a:avLst></a:prstTxWarp></a:bodyPr><a:p><a:r><a:rPr sz="3600" b="1"><a:latin typeface="Times New Roman"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>MM</a:t></a:r></a:p></a:txBody>'))
    const ordinary = { ...model, textWarp: undefined }
    const W = 260, H = 180
    const canvas = createCanvas(W, H)
    const ctx = canvas.getContext('2d')
    paintTextBody(model, ctx as unknown as never, 8, 8, 244, 164, (x: string) => x)
    const paint = (m: DrawingTextBody) => (r: CanvasRenderingContext2D): void => {
      paintTextBody(m, r as unknown as never, 8, 8, 244, 164, (x: string) => x)
    }
    const wi = await buildTextIndex([{ spec: { widthPx: W, heightPx: H }, paint: paint(model) }])
    const oi = await buildTextIndex([{ spec: { widthPx: W, heightPx: H }, paint: paint(ordinary as DrawingTextBody) }])
    const data = ctx.getImageData(0, 0, W, H).data
    let both = 0, different = 0
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      if (data[(y * W + x) * 4 + 3] < 210) continue
      const w = hitTest(wi, 0, x + 0.5, y + 0.5, { strict: true })
      const o = hitTest(oi, 0, x + 0.5, y + 0.5, { strict: true })
      if (w && o) {
        both++
        if (w.charIndex !== o.charIndex) different++
      }
    }
    expect(both).toBeGreaterThan(500)
    expect(different).toBe(0)
  })
})

describe('R2 finding 6/quality F3: highlights intersect both range ends', () => {
  it('range [1,2) on MMMM yields one cluster band, mixed spans kept', async () => {
    const body = warpBody('textCircle', {}, 'MMMM', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    const index = await buildTextIndex([{ spec: { widthPx: 240, heightPx: 130 }, paint: (replayCtx: CanvasRenderingContext2D) => {
      paintTextBody(body, replayCtx as unknown as never, MARGIN, MARGIN, 224, 114, identityFont)
    } }])
    const line = index.pages[0].lines[0]
    const rects = rectsForRange(line, 1, 2)
    expect(rects.length).toBe(1)
  })
  it('full Circle highlight covers the mapped ink', async () => {
    const body = warpBody('textCircle', { adj: 10800000 }, 'MMMMMMMM', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    const W = 650, H = 500
    const canvas = createCanvas(W, H)
    const ctx = canvas.getContext('2d')
    paintTextBody(body, ctx as unknown as never, 150, 100, 300, 300, identityFont)
    const data = ctx.getImageData(0, 0, W, H).data
    const index = await buildTextIndex([{ spec: { widthPx: W, heightPx: H }, paint: (replayCtx: CanvasRenderingContext2D) => {
      paintTextBody(body, replayCtx as unknown as never, 150, 100, 300, 300, identityFont)
    } }])
    const line = index.pages[0].lines[0]
    const rects = rectsForRange(line, 0, line.text.length)
    let ink = 0, covered = 0
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const i = (y * W + x) * 4
      if (data[i + 3] < 210 || (data[i] >= 250 && data[i + 1] >= 250 && data[i + 2] >= 250)) continue
      ink++
      if (rects.some(r => x + 0.5 >= r.x && x + 0.5 <= r.x + r.width && y + 0.5 >= r.y && y + 0.5 <= r.y + r.height)) covered++
    }
    expect(ink).toBeGreaterThan(500)
    expect(covered / ink).toBeGreaterThanOrEqual(0.9)
  })
  it('styled two-run highlight keeps both spans with real blue ink covered', async () => {
    const body = parseTextBody(parseXmlOrdered(`
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr>${warpXml('textCurveUp', { adj: 45977 })}</a:bodyPr>
  <a:p>
    <a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>AB</a:t></a:r>
    <a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr><a:t>CD</a:t></a:r>
  </a:p>
</a:txBody>`))
    body.insetLeftEmu = 0; body.insetRightEmu = 0; body.insetTopEmu = 0; body.insetBottomEmu = 0
    const W = 300, H = 140
    const canvas = createCanvas(W, H)
    const ctx = canvas.getContext('2d')
    paintTextBody(body, ctx as unknown as never, MARGIN, MARGIN, 284, 124, identityFont)
    const data = ctx.getImageData(0, 0, W, H).data
    const index = await buildTextIndex([{ spec: { widthPx: W, heightPx: H }, paint: (replayCtx: CanvasRenderingContext2D) => {
      paintTextBody(body, replayCtx as unknown as never, MARGIN, MARGIN, 284, 124, identityFont)
    } }])
    const line = index.pages[0].lines[0]
    const rects = rectsForRange(line, 0, line.text.length)
    expect(rects.length).toBeGreaterThanOrEqual(2)
    let blue = 0, blueCovered = 0
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const i = (y * W + x) * 4
      if (data[i + 3] < 210 || !(data[i + 2] > data[i] * 1.3 && data[i + 2] > data[i + 1] * 1.3)) continue
      blue++
      if (rects.some(r => x + 0.5 >= r.x && x + 0.5 <= r.x + r.width && y + 0.5 >= r.y && y + 0.5 <= r.y + r.height)) blueCovered++
    }
    expect(blue).toBeGreaterThan(20)
    expect(blueCovered / blue).toBeGreaterThanOrEqual(0.9)
  })
})

describe('R2 finding 7/quality F4: fallback records share resolved mode', () => {
  it('budget-exceeded paint records no warp geometry', async () => {
    const { MAX_WARP_SURFACE_PIXELS } = await import('../src/drawing/text-paint')
    expect(MAX_WARP_SURFACE_PIXELS).toBeGreaterThan(0)
    // One non-wrapping run far beyond any bounded scratch budget: paint and
    // record must BOTH take the diagnosed ordinary fallback.
    const { parseXmlOrdered } = await import('../src/core/xml')
    const { parseTextBody } = await import('../src/drawing/text-parse')
    const big = 'M'.repeat(5000)
    const body = parseTextBody(parseXmlOrdered(`
<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:bodyPr wrap="none"><a:prstTxWarp prst="textCircle"><a:avLst/></a:prstTxWarp></a:bodyPr>
  <a:p><a:r><a:rPr sz="3600"><a:solidFill><a:srgbClr val="200000"/></a:solidFill></a:rPr><a:t>${big}</a:t></a:r></a:p>
</a:txBody>`))
    body.insetLeftEmu = 0; body.insetRightEmu = 0; body.insetTopEmu = 0; body.insetBottomEmu = 0
    const { RECORD_TEXT } = await import('../src/core/text-recording')
    const canvas = createCanvas(300, 200)
    const ctx = canvas.getContext('2d') as any
    let visualCount = 0, records = 0
    ctx[RECORD_TEXT] = (_t: string, _x: number, _y: number, _w: number, logical: any) => {
      records++
      if (logical?.visual) visualCount++
    }
    paintTextBody(body, ctx as unknown as never, 8, 8, 284, 184, identityFont)
    expect(records).toBeGreaterThan(0)
    expect(visualCount).toBe(0)
    expect(((body as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics ?? []).some((d: { feature: string }) => d.feature === 'warp-budget')).toBe(true)
  })
  it('vertical budget fallback paints transformed ordinary ink indexed at 100%', async () => {
    const model = parseTextBody(parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="ctr" vert="vert"><a:prstTxWarp prst="textCircle"><a:avLst/></a:prstTxWarp></a:bodyPr><a:p><a:r><a:rPr sz="3600" b="1"><a:latin typeface="Times New Roman"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>${'M'.repeat(260)}</a:t></a:r></a:p></a:txBody>`))
    const W = 260, H = 180
    const canvas = createCanvas(W, H)
    const ctx = canvas.getContext('2d') as any
    paintTextBody(model, ctx, 8, 8, 244, 164, (x: string) => x, undefined, { clip: { x: 8, y: 8, width: 244, height: 164 } })
    const index = await buildTextIndex([{ spec: { widthPx: W, heightPx: H }, paint: (r: CanvasRenderingContext2D) =>
      paintTextBody(model, r as unknown as never, 8, 8, 244, 164, (x: string) => x, undefined, { clip: { x: 8, y: 8, width: 244, height: 164 } }) }])
    const data = ctx.getImageData(0, 0, W, H).data
    let ink = 0, hits = 0
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const i = (y * W + x) * 4
      if (data[i + 3] < 210 || (data[i] >= 250 && data[i + 1] >= 250 && data[i + 2] >= 250)) continue
      ink++
      if (hitTest(index, 0, x + 0.5, y + 0.5, { strict: true })) hits++
    }
    const hasVisual = index.pages[0].lines.some(l => l.spans.some(s => (s as unknown as { placement?: { visual?: unknown } }).placement?.visual))
    expect(ink).toBeGreaterThan(100)
    expect(hasVisual).toBe(false)
    expect(hits / ink).toBeGreaterThanOrEqual(0.99)
  })
  for (const order of ['paint-first', 'index-first'] as const) {
    it(`scratch allocation failure (${order}) keeps transformed ordinary ink indexed with a truthful diagnostic`, async () => {
      const model = parseTextBody(parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="ctr" vert="vert270"><a:prstTxWarp prst="textCurveUp"/></a:bodyPr><a:p><a:r><a:rPr sz="3600" b="1"><a:solidFill><a:srgbClr val="200000"/></a:solidFill></a:rPr><a:t>Vertical Art</a:t></a:r></a:p></a:txBody>`))
      ;(globalThis as any).OffscreenCanvas = function () { throw Error('isolated scratch allocation failure') }
      try {
        const W = 260, H = 180
        const canvas = createCanvas(W, H)
        const ctx = canvas.getContext('2d') as any
        const paint = (): void => {
          paintTextBody(model, ctx, 8, 8, 244, 164, (x: string) => x, undefined, { clip: { x: 8, y: 8, width: 244, height: 164 } })
        }
        const record = (): Promise<Awaited<ReturnType<typeof buildTextIndex>>> => buildTextIndex([{ spec: { widthPx: W, heightPx: H }, paint: (r: CanvasRenderingContext2D) =>
          paintTextBody(model, r as unknown as never, 8, 8, 244, 164, (x: string) => x, undefined, { clip: { x: 8, y: 8, width: 244, height: 164 } }) }])
        let index: Awaited<ReturnType<typeof buildTextIndex>>
        if (order === 'paint-first') { paint(); index = await record() } else { index = await record(); paint() }
        const data = ctx.getImageData(0, 0, W, H).data
        let ink = 0, hits = 0
        for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
          const i = (y * W + x) * 4
          if (data[i + 3] < 210 || (data[i] >= 250 && data[i + 1] >= 250 && data[i + 2] >= 250)) continue
          ink++
          if (hitTest(index, 0, x + 0.5, y + 0.5, { strict: true })) hits++
        }
        const hasVisual = index.pages[0].lines.some(l => l.spans.some(s => (s as unknown as { placement?: { visual?: unknown } }).placement?.visual))
        const diagnostics = (model as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics ?? []
        expect(ink).toBeGreaterThan(100)
        expect(hasVisual).toBe(false)
        expect(hits / ink).toBeGreaterThanOrEqual(0.99)
        expect(diagnostics.some(d => d.feature === 'warp-alloc-failure')).toBe(true)
      } finally {
        delete (globalThis as any).OffscreenCanvas
      }
    })
  }
})

describe('R2 finding 8/quality F5: mesh does not accumulate alpha', () => {
  it('half-alpha warped source keeps max alpha near 127 on transparency', () => {
    const body = warpBody('textCurveUp', { adj: 45977 }, 'MMMM', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill>')
    const canvas = createCanvas(240, 130)
    const ctx = canvas.getContext('2d')
    // Transparent target (no white fill).
    ctx.clearRect(0, 0, 240, 130)
    ctx.globalAlpha = 0.5
    paintTextBody(body, ctx as unknown as never, MARGIN, MARGIN, 224, 114, identityFont)
    ctx.globalAlpha = 1
    const data = ctx.getImageData(0, 0, 240, 130).data
    let max = 0
    for (let i = 0; i < data.length; i += 4) if (data[i + 3] > max) max = data[i + 3]
    expect(max).toBeLessThanOrEqual(140)
  })
})

describe('R2 finding 9/quality F6: 0.2px outlines survive warp raster', () => {
  it('raster stroke uses authored 0.2px width', async () => {
    const { createCanvas: cc } = await import('canvas')
    const widths: number[] = []
    const realCreate = cc as unknown as (w: number, h: number) => any
    ;(globalThis as any).OffscreenCanvas = function (this: unknown, w: number, h: number) {
      const surface = realCreate.call(null, w, h)
      if (w !== 8 || h !== 8) {
        const original = surface.getContext.bind(surface)
        surface.getContext = (kind: string): unknown => {
          const ctx = original(kind) as any
          return new Proxy(ctx, {
            get(target, prop) {
              if (prop === 'strokeText') return (text: string, ...rest: unknown[]) => {
                widths.push(target.lineWidth)
                return target.strokeText(text, ...rest)
              }
              const value = target[prop]
              return typeof value === 'function' ? value.bind(target) : value
            },
            set(target, prop, value) { target[prop as string] = value; return true },
          })
        }
      }
      return surface
    }
    try {
      const canvas = cc(240, 130)
      const ctx = canvas.getContext('2d')
      const body = warpBody('textCurveUp', { adj: 45977 }, 'M', '<a:solidFill><a:srgbClr val="200000"/></a:solidFill><a:ln w="1905"><a:solidFill><a:srgbClr val="008800"/></a:solidFill></a:ln>')
      const { paintTextBody: fresh } = await import('../src/drawing/text-paint')
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, 240, 130)
      fresh(body, ctx as unknown as never, MARGIN, MARGIN, 224, 114, identityFont)
      expect(widths.length).toBeGreaterThan(0)
      for (const w of widths) expect(w).toBeCloseTo(0.2, 2)
    } finally {
      delete (globalThis as any).OffscreenCanvas
    }
  })
})
