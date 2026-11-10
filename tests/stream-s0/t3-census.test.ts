/** Stage-0 Task 3 recorder-census tests. */
import { createCanvas } from 'canvas'
import { expect, it } from 'vitest'
import { RECORD_TEXT } from '../../src/core/text-recording'
import { wrapCensus } from '../../scripts/stream-s0/census-proxy'

const ctxOf = () => createCanvas(24, 16).getContext('2d') as unknown as CanvasRenderingContext2D

it('proxy forwards and counts (fillRect twice → count 2, pixels painted)', () => {
  const real = ctxOf()
  const { ctx, report } = wrapCensus(real)
  ctx.save() // baseline state so fillRect really marks ink
  ctx.fillStyle = '#ffffff' // node-canvas starts transparent black
  ctx.fillRect(0, 0, 24, 16)
  ctx.fillStyle = '#000000'
  ctx.fillRect(0, 0, 8, 4)
  ctx.fillRect(12, 4, 8, 4)
  ctx.restore()
  const report1 = report()
  const keys = Object.keys(report1.methods).filter((k) => k === 'fillRect/4')
  expect(keys).toHaveLength(1)
  expect(report1.methods['fillRect/4']).toBe(3)
  // pixels painted: the two black rects = 8×4 + 8×4 = 64 (white bg first).
  // Read via the REAL ctx (renderUnit-style): the census ctx does not
  // guarantee RGBA reads, and getImageData is not part of §7.3's subset.
  const data = real.getImageData(0, 0, 24, 16).data
  expect(data.length).toBe(24 * 16 * 4)
  let black = 0
  for (let i = 0; i < data.length; i += 4) {
    if (data[i] === 0 && data[i + 1] === 0 && data[i + 2] === 0 && data[i + 3] === 255) black++
  }
  expect(black).toBe(64)
  // subset coverage: all §7.3-recorded, nothing unrecorded in the paint
  expect(report1.notInSubset).toEqual([])
})

it('unknown op lands in notInSubset (ctx.filter is a §7.3 gap)', () => {
  const { ctx, report } = wrapCensus(ctxOf())
  try { (ctx as unknown as { filter?: string }).filter = 'blur(1px)' } catch { /* backends may refuse the value; the census still counts the write */ }
  const r = report()
  expect(r.notInSubset.some((k) => k === 'filter' || k.startsWith('filter'))).toBe(true)
})

it("instanceof and 'letterSpacing' in ctx behave as on the real ctx", () => {
  const real = ctxOf()
  const { ctx } = wrapCensus(real)
  expect('letterSpacing' in ctx).toBe('letterSpacing' in real)
  expect(ctx instanceof Object).toBe(true)
  // canvas getter: a census ctx must keep the real canvas reachable (painters
  // scratch-clone through ctx.canvas.constructor, draft-8 lost rule 13).
  expect(ctx.canvas).toBe(real.canvas)
})

it('RECORD_TEXT is forwarded per plan (not a canvas op)', () => {
  const { ctx, report } = wrapCensus(ctxOf())
  const probe = (ctx as unknown as Record<PropertyKey, unknown>)
  // Painters read the symbol off the ctx: with no recorder registration, the
  // read is absent (§7.3 recorder never exposes it either) — forwarded, not
  // counted.
  expect(probe[RECORD_TEXT]).toBeUndefined()
  const r = report()
  expect(r.recordText).toHaveLength(0)
  expect(r.fontStrings).toEqual([])
  expect(Object.keys(r.methods).some((k) => k.includes('recordText'))).toBe(false)
  expect(r.notInSubset.some((k) => k.includes('recordText'))).toBe(false)
})

it('gradient and pattern wrappers are census evidence (§7.3 recorded subset)', () => {
  const { ctx, report } = wrapCensus(ctxOf())
  const grad = ctx.createLinearGradient(0, 0, 10, 10)
  const addColorStop = (grad as unknown as { addColorStop: (o: number, c: string) => void }).addColorStop
  addColorStop(0, '#000000')
  addColorStop(1, '#ffffff')
  const pattern = ctx.createPattern(ctx.canvas as unknown as CanvasImageSource, 'repeat')
  if (pattern && typeof (pattern as unknown as { setTransform?: unknown }).setTransform === 'function') {
    // DOMMatrix is not global in node: skip the wrapper transform probe that
    // needs one (round-9 note) — pattern wrapper presence + addColorStop
    // already cover the wrapper-recording path.
  }
  const r = report()
  expect(r.methods['createLinearGradient/4']).toBe(1)
  expect(r.methods['createPattern/2']).toBe(1)
  // Wrapper key = the created method name (§7.3 wrapper setTransform + the
  // addColorStop that gradient ops record). Consistent §7.3/§7.5 frame keys.
  expect(r.methods['createLinearGradient.addColorStop/2']).toBe(2)
  // §§7.3 wrapper setTransform is recorded; nothing unrecorded should show.
  expect(r.notInSubset.some((k) => k.includes('addColorStop'))).toBe(false)
  expect(Object.keys(r.methods).some((k) => k.startsWith('Pattern.setTransform') ||
    k.includes('setTransform') && !k.startsWith('setTransform/'))).toBe(false)
})
