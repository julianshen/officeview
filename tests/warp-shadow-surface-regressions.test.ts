/**
 * R2: a failing SHADOW-composite surface must never flip paint to ordinary
 * while the index keeps warped visuals. The source-raster preflight is a
 * separate policy (unchanged): here the source surface succeeds and only the
 * later composite context is null/throwing, in both index/paint orders. The
 * intended direct warped path must continue: warped records kept, warped ink
 * painted, zero strict misses over actual ink, no fallback diagnostics,
 * caller state kept.
 *
 * Branch note: both original left-aligned fixtures (Wave1 AND SlantUp)
 * execute the MESH branch — left-aligned SlantUp's padded source crosses u<0
 * where the path clamps, so affinity is rejected. Genuine affine-branch
 * coverage (the changed paintAffineSource guard) comes only from the separate
 * centered-SlantUp block below, which asserts its operation counts.
 *
 * Allocation tracing records the paint/index phase explicitly: the composed
 * size is taken from healthy PAINT-phase surfaces (the 4x4 measurer never
 * goes through the scratch factory), so a green run cannot pass without ever
 * exercising the composite. Failed allocations are asserted to have really
 * triggered while source allocations stayed healthy.
 */
import { createCanvas } from 'canvas'
import { describe, expect, it } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseTextBody } from '../src/drawing/text-parse'
import { paintTextBody } from '../src/drawing/text-paint'
import { buildTextIndex } from '../src/core/search'
import { hitTest } from '../src/core/selection'
import type { DrawingTextBody } from '../src/drawing/text'

function shadowBody(preset: string, align: 'left' | 'center' = 'left'): DrawingTextBody {
  const pPr = align === 'center' ? '<a:pPr algn="ctr"/>' : ''
  return parseTextBody(parseXmlOrdered(`<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"><a:prstTxWarp prst="${preset}"/></a:bodyPr><a:p>${pPr}<a:r><a:rPr sz="2400"><a:effectLst><a:outerShdw blurRad="0" dist="38100" dir="0"><a:srgbClr val="555555"/></a:outerShdw></a:effectLst></a:rPr><a:t>MMMM</a:t></a:r></a:p></a:txBody>`))
}

interface ScratchSize { w: number; h: number }
interface Allocation extends ScratchSize { phase: string; fail: boolean }

async function runCase(preset: string, failure: 'healthy' | 'null' | 'throw', order: 'index-first' | 'paint-first', composedSize?: ScratchSize, align: 'left' | 'center' = 'left', countOpsFor?: ScratchSize): Promise<{
  hasVisual: boolean; ordinaryEqual: boolean; changed: number; targetFillCalls: string[];
  sampledHits: number; sampledTotal: number; faultInk: number; faultMisses: number;
  diagnostics: Array<{ feature: string }>; allocations: Allocation[];
  stateKept: boolean; composedOps: { clips: number; transforms: number; draws: number };
}> {
  const allocations: Allocation[] = []
  const composedOps = { clips: 0, transforms: 0, draws: 0 }
  let phase = ''
  ;(globalThis as any).OffscreenCanvas = function (w: number, h: number) {
    const fail = composedSize !== undefined && w === composedSize.w && h === composedSize.h && failure !== 'healthy'
    allocations.push({ phase, w, h, fail })
    if (fail) {
      return {
        width: w, height: h,
        getContext(): unknown {
          if (failure === 'throw') throw new Error('composed 2D context unavailable')
          return null
        },
      }
    }
    const surface: any = createCanvas(w, h)
    if (countOpsFor !== undefined && w === countOpsFor.w && h === countOpsFor.h) {
      const originalGet = surface.getContext.bind(surface)
      surface.getContext = (kind: string): unknown => {
        const ctx = originalGet(kind) as any
        const clip = ctx.clip.bind(ctx)
        ctx.clip = (...args: unknown[]) => { composedOps.clips++; return clip(...args) }
        const transform = ctx.transform.bind(ctx)
        ctx.transform = (...args: unknown[]) => { composedOps.transforms++; return transform(...args) }
        const drawImage = ctx.drawImage.bind(ctx)
        ctx.drawImage = (...args: unknown[]) => { composedOps.draws++; return drawImage(...args) }
        return ctx
      }
    }
    return surface
  }
  try {
    const b = shadowBody(preset, align)
    const c = createCanvas(400, 300)
    const ctx = c.getContext('2d') as any
    const targetFillCalls: string[] = []
    const fill = ctx.fillText.bind(ctx)
    ctx.fillText = (text: string, x: number, y: number) => { targetFillCalls.push(text); return fill(text, x, y) }
    const before = { font: ctx.font, lineWidth: ctx.lineWidth, shadow: String(ctx.shadowColor), alpha: ctx.globalAlpha, align: ctx.textAlign, baseline: ctx.textBaseline }
    const paint = (): void => { phase = 'paint'; paintTextBody(b, ctx, 40, 100, 200, 100, (x: string) => x) }
    let index: Awaited<ReturnType<typeof buildTextIndex>> | undefined
    const record = async (): Promise<void> => {
      phase = 'index'
      index = await buildTextIndex([{ spec: { widthPx: 400, heightPx: 300 }, paint: (r: CanvasRenderingContext2D) => paintTextBody(b, r as unknown as never, 40, 100, 200, 100, (x: string) => x) }])
    }
    if (order === 'index-first') { await record(); paint() } else { paint(); await record() }
    const after = { font: ctx.font, lineWidth: ctx.lineWidth, shadow: String(ctx.shadowColor), alpha: ctx.globalAlpha, align: ctx.textAlign, baseline: ctx.textBaseline }
    const plain = shadowBody(preset, align)
    plain.textWarp = undefined
    const ref = createCanvas(400, 300)
    paintTextBody(plain, ref.getContext('2d') as any, 40, 100, 200, 100, (x: string) => x)
    const pixels = ctx.getImageData(0, 0, 400, 300).data
    const ordinaryEqual = Buffer.from(pixels).equals(Buffer.from(ref.getContext('2d').getImageData(0, 0, 400, 300).data))
    // Changed-pixel buffer: any visible deposit (shadow/AA-compressed glyph
    // ink has no fixed alpha floor on the healthy direct path).
    let changed = 0
    const candidates: Array<[number, number]> = []
    for (let y = 0; y < 300; y++) for (let x = 0; x < 400; x++) {
      const at = (y * 400 + x) * 4
      if (pixels[at + 3] >= 64) {
        changed++
        // Glyph-dominant samples only: the gray drop shadow is painted offset
        // outside the glyph band by design, so it must not count as missed
        // glyph ink. Opaque near-black pixels are the warped glyphs.
        if (pixels[at + 3] >= 210 && pixels[at] < 40 && pixels[at + 1] < 40 && pixels[at + 2] < 40 && (x + y * 400) % 7 === 0) candidates.push([x + 0.5, y + 0.5])
      }
    }
    let sampledHits = 0
    for (const [x, y] of candidates) {
      if (hitTest(index!, 0, x, y, { strict: true })) sampledHits++
    }
    // Original ordinary-fallback fault signature: opaque black ink missed by
    // strict hits (kept exactly — the fault really does paint ordinary).
    let faultInk = 0, faultMisses = 0
    for (let y = 0; y < 300; y++) for (let x = 0; x < 400; x++) {
      const at = (y * 400 + x) * 4
      if (pixels[at + 3] >= 210 && pixels[at] < 20 && pixels[at + 1] < 20 && pixels[at + 2] < 20) {
        faultInk++
        if (!hitTest(index!, 0, x + 0.5, y + 0.5, { strict: true })) faultMisses++
      }
    }
    return {
      hasVisual: index!.pages[0].lines.some(l => l.spans.some(s => (s as unknown as { placement?: { visual?: unknown } }).placement?.visual)),
      ordinaryEqual, changed, targetFillCalls,
      sampledHits, sampledTotal: candidates.length, faultInk, faultMisses,
      diagnostics: ((b as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics ?? []),
      allocations, stateKept: JSON.stringify(before) === JSON.stringify(after),
      composedOps,
    }
  } finally {
    delete (globalThis as any).OffscreenCanvas
  }
}

/** Composed size from healthy PAINT-phase surfaces (second paint scratch). */
async function paintComposedSize(preset: string, align: 'left' | 'center' = 'left'): Promise<ScratchSize> {
  const r = await runCase(preset, 'healthy', 'paint-first', undefined, align)
  const paintSizes = r.allocations.filter(a => a.phase === 'paint')
  expect(paintSizes.length).toBeGreaterThanOrEqual(2)
  return { w: paintSizes[1].w, h: paintSizes[1].h }
}

describe('R2 shadow-composite failure keeps the warped path', () => {
  for (const preset of ['textWave1', 'textSlantUp']) {
    it(`${preset}: healthy composition paints warped ink hit at sampled positions`, async () => {
      const r = await runCase(preset, 'healthy', 'paint-first')
      expect(r.targetFillCalls).toEqual([])
      expect(r.changed).toBeGreaterThan(100)
      expect(r.hasVisual).toBe(true)
      expect(r.ordinaryEqual).toBe(false)
      expect(r.sampledTotal).toBeGreaterThan(20)
      expect(r.sampledHits).toBe(r.sampledTotal)
      expect(r.diagnostics.length).toBe(0)
      expect(r.stateKept).toBe(true)
    })
    for (const failure of ['null', 'throw'] as const) {
      for (const order of ['index-first', 'paint-first'] as const) {
        it(`${preset}: composed ${failure} context (${order}) keeps warped paint/index agreement`, async () => {
          const composedSize = await paintComposedSize(preset)
          const r = await runCase(preset, failure, order, composedSize)
          // The composite allocation really failed; every source allocation
          // stayed healthy.
          const failed = r.allocations.filter(a => a.fail)
          expect(failed.length).toBeGreaterThanOrEqual(1)
          expect(failed.every(a => a.w === composedSize.w && a.h === composedSize.h)).toBe(true)
          expect(r.allocations.filter(a => !a.fail && (a.w !== composedSize.w || a.h !== composedSize.h)).length).toBeGreaterThan(0)
          expect(r.targetFillCalls).toEqual([])
          expect(r.changed).toBeGreaterThan(100)
          expect(r.hasVisual).toBe(true)
          expect(r.ordinaryEqual).toBe(false)
          expect(r.sampledTotal).toBeGreaterThan(20)
          expect(r.sampledHits).toBe(r.sampledTotal)
          // Original fault metric stays live: opaque black ink (if any) is
          // fully hit — the 175/791 ordinary-fallback miss pattern is gone.
          expect(r.faultMisses).toBe(0)
          expect(r.diagnostics.some(d => d.feature === 'warp-budget' || d.feature === 'warp-alloc-failure' || d.feature === 'warp-paint-failure')).toBe(false)
          expect(r.stateKept).toBe(true)
        })
      }
    }
  }
})

describe('R2 genuine affine shadow-composite (centered SlantUp)', () => {
  // Centered paragraph alignment keeps the padded source inside u>=0, so the
  // combined map stays affine and the changed paintAffineSource guard (not
  // the mesh path) executes. Operation counts prove the branch: the affine
  // fastpath performs zero clips, one transform, one drawImage on the
  // composed context.
  it('healthy composition takes the affine branch with warped ink hit', async () => {
    for (const order of ['index-first', 'paint-first'] as const) {
      const composedSize = await paintComposedSize('textSlantUp', 'center')
      const r = await runCase('textSlantUp', 'healthy', order, composedSize, 'center', composedSize)
      expect(r.composedOps).toEqual({ clips: 0, transforms: 1, draws: 1 })
      expect(r.targetFillCalls).toEqual([])
      expect(r.changed).toBeGreaterThan(100)
      expect(r.hasVisual).toBe(true)
      expect(r.ordinaryEqual).toBe(false)
      expect(r.sampledTotal).toBeGreaterThan(20)
      expect(r.sampledHits).toBe(r.sampledTotal)
      expect(r.diagnostics.length).toBe(0)
      expect(r.stateKept).toBe(true)
    }
  })
  for (const failure of ['null', 'throw'] as const) {
    for (const order of ['index-first', 'paint-first'] as const) {
      it(`centered SlantUp: composed ${failure} context (${order}) keeps warped paint/index agreement`, async () => {
        const composedSize = await paintComposedSize('textSlantUp', 'center')
        const r = await runCase('textSlantUp', failure, order, composedSize, 'center')
        const failed = r.allocations.filter(a => a.fail)
        expect(failed.length).toBeGreaterThanOrEqual(1)
        expect(failed.every(a => a.w === composedSize.w && a.h === composedSize.h)).toBe(true)
        expect(r.allocations.filter(a => !a.fail && (a.w !== composedSize.w || a.h !== composedSize.h)).length).toBeGreaterThan(0)
        expect(r.targetFillCalls).toEqual([])
        expect(r.changed).toBeGreaterThan(100)
        expect(r.hasVisual).toBe(true)
        expect(r.ordinaryEqual).toBe(false)
        expect(r.sampledTotal).toBeGreaterThan(20)
        expect(r.sampledHits).toBe(r.sampledTotal)
        expect(r.faultMisses).toBe(0)
        expect(r.diagnostics.some(d => d.feature === 'warp-budget' || d.feature === 'warp-alloc-failure' || d.feature === 'warp-paint-failure')).toBe(false)
        expect(r.stateKept).toBe(true)
      })
    }
  }
})
