/** Stage 2b Task 2: continuous zoom registration; abstain on unsupported geometry. */
import { describe, expect, it } from 'bun:test'
import { createCanvas } from '@napi-rs/canvas'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { basis, deriveBasisKey } from '../src/mark/basis'
import { extendedCodeword } from '../src/mark/bch'
import { decideScores, prepareCandidates, scoreCapture } from '../src/mark/detect'
import { messageFor } from '../src/mark/message'
import { markRaster } from '../src/mark/mark-raster'
import { DEFAULT_PARAMS } from '../src/mark/params'
import { decodePng } from '../src/mark/png'
import { embed } from '../src/mark/embed'
import { GEOMETRY_ERROR_MAX, register } from '../src/mark/register'
import { tauFor } from '../src/detect/tau'
import type { Bitmap } from '../src/mark/types'
import { DocumentHost } from '../src/render'
import { crop, jpeg, keystone, rotate, scale } from '../../scripts/stream-s0/mark/attacks'

const root = resolve(import.meta.dir, '../..')
const TAU = { abs: 7.99, margin: 3.85 }
const TOKEN = 417
const key = Uint8Array.from({ length: 32 }, (_, i) => i + 1)
const codes = Array.from({ length: 1024 }, (_, token) => extendedCodeword(messageFor(key, 1, token, 17)))
const matrix = prepareCandidates(codes)
const bmp = (d: { w: number; h: number; rgba: Uint8ClampedArray }): Bitmap => ({ width: d.w, height: d.h, data: d.rgba })

const host = new DocumentHost()
const doc = await host.open('pypptx', new Uint8Array(readFileSync(resolve(root, 'corpus/pypptx-test.pptx'))), 'visible-only')
const kBasis = deriveBasisKey(key, 'docref-pypptx', '0@1')
const unit = host.render(doc.docKey, { unitIndex: 0, scale: 1, projection: 'final-raster', mark: { kind: 'viewer', text: 'viewer' } }) as { png: Uint8Array; w: number; h: number }
const marked0 = markRaster({ png: unit.png, w: unit.w, h: unit.h, kBasis, code: codes[TOKEN]! })
if ('unmarked' in marked0) throw new Error(`unexpected unmarked: ${marked0.unmarked}`)
const reference = bmp(decodePng(unit.png)), marked = bmp(decodePng(marked0.png))
const chips = basis(kBasis, unit.w, unit.h, DEFAULT_PARAMS)
const decide = (capture: Bitmap) => decideScores(scoreCapture(capture, reference, chips, matrix, DEFAULT_PARAMS), TAU)

describe('continuous zoom', () => {
  for (const f of [0.5, 0.75, 1.1, 1.25, 1.5, 1.75, 2]) {
    it(`registers browser zoom ${f}`, () => {
      const capture = scale(marked, f), reg = register(capture, reference, chips)
      if (!reg) throw new Error('registration failed')
      expect(Math.abs(reg.factor - f)).toBeLessThanOrEqual(0.005)
      expect(reg.residual).toBeLessThanOrEqual(GEOMETRY_ERROR_MAX)
      expect(decide(capture)).toMatchObject({ kind: 'attributed', index: TOKEN })
    }, 30_000)
  }
  it('zoom plus jpeg', async () => {
    const capture = await jpeg(scale(marked, 1.25), 70)
    expect(register(capture, reference, chips)!.residual).toBeLessThanOrEqual(GEOMETRY_ERROR_MAX)
    expect(decide(capture)).toMatchObject({ kind: 'attributed', index: TOKEN })
  }, 30_000)
  it('zoom then a 10 % crop (search within the crop range)', () => {
    const capture = crop(scale(marked, 1.3), 0.1, 'bottom-right'), reg = register(capture, reference, chips)
    if (!reg) throw new Error('registration failed')
    expect(Math.abs(reg.factor - 1.3)).toBeLessThanOrEqual(0.005)
    expect(decide(capture)).toMatchObject({ kind: 'attributed', index: TOKEN })
  }, 30_000)
  it('registration time: one zoom-1.25 capture registers in <= 15 s', () => {
    const capture = scale(marked, 1.25), start = performance.now()
    expect(register(capture, reference, chips)).toBeDefined()
    const ms = performance.now() - start
    console.log(`zoom-1.25 registration ${ms.toFixed(0)} ms`)
    expect(ms).toBeLessThanOrEqual(15_000)
  }, 30_000)
})

// A served scale-2 page (2560x1440) that is nearly flat: a uniform fill under
// the tiled, almost isoluminant visible mark. An uncropped zoom of 0.75 (1.5)
// has the extents of a 25 % crop at factor 1 (2), and on this page that wrong
// crop is in scope too; registration must pick the better fit, not the first
// in-scope one, and must not spend minutes certifying the losing hypothesis.
const lo = await host.open('lo-airplane', new Uint8Array(readFileSync(resolve(root, 'corpus/lo-airplane-transition.pptx'))), 'visible-only')
const page = host.render(lo.docKey, { unitIndex: 0, scale: 2, projection: 'final-raster', mark: { kind: 'viewer', text: 'viewer' } }) as { png: Uint8Array; w: number; h: number }
const loBasis = deriveBasisKey(key, 'docref-lo-airplane', '0@2')
const loMarked0 = markRaster({ png: page.png, w: page.w, h: page.h, kBasis: loBasis, code: codes[TOKEN]! })
if ('unmarked' in loMarked0) throw new Error(`unexpected unmarked: ${loMarked0.unmarked}`)
const loReference = bmp(decodePng(page.png)), loMarked = bmp(decodePng(loMarked0.png))
const loChips = basis(loBasis, page.w, page.h, DEFAULT_PARAMS)
const tau = tauFor(2 ** 16)
describe('scale-2 sparse page (lo-airplane-transition)', () => {
  it('is a scale-2 page', () => expect([page.w, page.h]).toEqual([2560, 1440]))
  const zooms: [string, number, (b: Bitmap) => Bitmap | Promise<Bitmap>][] = [
    ['zoom 0.75', 0.75, b => scale(b, 0.75)], ['zoom 0.75 + q70', 0.75, b => jpeg(scale(b, 0.75), 70)],
    ['zoom 1.5', 1.5, b => scale(b, 1.5)], ['zoom 1.5 + q70', 1.5, b => jpeg(scale(b, 1.5), 70)],
  ]
  for (const [name, f, attack] of zooms) {
    it(`uncropped ${name} registers to the true factor and attributes`, async () => {
      const capture = await attack(loMarked), reg = register(capture, loReference, loChips)
      if (!reg) throw new Error('registration failed')
      expect(Math.abs(reg.factor - f)).toBeLessThanOrEqual(0.005)
      expect(reg.residual).toBeLessThanOrEqual(GEOMETRY_ERROR_MAX)
      expect(decideScores(scoreCapture(capture, loReference, loChips, matrix, DEFAULT_PARAMS), tau)).toMatchObject({ kind: 'attributed', index: TOKEN })
    }, 60_000)
    it(`registration time: uncropped ${name} registers in <= 15 s`, async () => {
      const capture = await attack(loMarked), start = performance.now()
      expect(register(capture, loReference, loChips)).toBeDefined()
      const ms = performance.now() - start
      console.log(`scale-2 ${name} registration ${ms.toFixed(0)} ms`)
      expect(ms).toBeLessThanOrEqual(15_000)
    }, 60_000)
  }
  // Same extents as zoom 0.75: a true 25 % crop must still register as the crop.
  for (const [name, attack] of [['crop25', (b: Bitmap) => crop(b, 0.25)], ['crop25 + q70', (b: Bitmap) => jpeg(crop(b, 0.25), 70)]] as const) {
    it(`a true ${name} registers as a factor-1 crop and attributes`, async () => {
      const capture = await attack(loMarked), reg = register(capture, loReference, loChips)
      if (!reg) throw new Error('registration failed')
      expect(reg.factor).toBe(1)
      expect([reg.x, reg.y]).toEqual([320, 180])
      expect(decideScores(scoreCapture(capture, loReference, loChips, matrix, DEFAULT_PARAMS), tau)).toMatchObject({ kind: 'attributed', index: TOKEN })
    }, 120_000)
  }
  // No uncropped zoom has these extents, so factor 1 is certified exhaustively;
  // the block lower bounds keep that to seconds on a flat page (was ~83 s).
  it('an anisotropic crop (2000x1080 at 280,180) certifies at factor 1 in <= 15 s', () => {
    const W = 2000, H = 1080, capture: Bitmap = { width: W, height: H, data: new Uint8ClampedArray(W * H * 4) }
    for (let y = 0; y < H; y++) capture.data.set(loMarked.data.subarray(((y + 180) * page.w + 280) * 4, ((y + 180) * page.w + 280 + W) * 4), y * W * 4)
    const start = performance.now(), reg = register(capture, loReference, loChips), ms = performance.now() - start
    console.log(`scale-2 anisotropic crop registration ${ms.toFixed(0)} ms`)
    expect(reg).toMatchObject({ factor: 1, x: 280, y: 180 })
    expect(ms).toBeLessThanOrEqual(15_000)
  }, 120_000)
})

// Stage-0 survival captures stay in scope; with the zoom cases above they are
// the in-scope basis of GEOMETRY_ERROR_MAX. A blank page's residual is its
// own mark energy.
describe('Stage-0 survival stays in scope', () => {
  const survival: [string, (b: Bitmap) => Bitmap | Promise<Bitmap>][] = [
    ['q70', b => jpeg(b, 70)], ['crop5', b => crop(b, 0.05)], ['crop25', b => crop(b, 0.25)],
    ['q70-down0.5', async b => scale(await jpeg(b, 70), 0.5)], ['crop25-down0.5', b => scale(crop(b, 0.25), 0.5)], ['crop25-up2', b => scale(crop(b, 0.25), 2)],
  ]
  for (const [name, attack] of survival) {
    it(`${name} attributes in scope`, async () => {
      const capture = await attack(marked), result = scoreCapture(capture, reference, chips, matrix, DEFAULT_PARAMS)
      expect(result.registration!.residual).toBeLessThanOrEqual(GEOMETRY_ERROR_MAX)
      expect(decideScores(result, TAU)).toMatchObject({ kind: 'attributed', index: TOKEN })
    }, 30_000)
  }
  // The Stage-0 detector fixtures (tests/stream-s0/mark-detect.test.ts): the
  // small blank page carries the densest mark energy per pixel and is the
  // largest in-scope residual, so it sets GEOMETRY_ERROR_MAX.
  const blank = (w: number, h: number): Bitmap => ({ width: w, height: h, data: new Uint8ClampedArray(w * h * 4).fill(255) })
  for (const [w, h] of [[unit.w, unit.h], [256, 256]] as const) {
    it(`a marked blank ${w}x${h} page is in scope`, () => {
      const white = blank(w, h), carrier = basis(new Uint8Array(32).fill(9), w, h, DEFAULT_PARAMS)
      const result = scoreCapture(embed(white, carrier, codes[TOKEN]!, DEFAULT_PARAMS), white, carrier, matrix, DEFAULT_PARAMS)
      expect(result.registration!.residual).toBeLessThanOrEqual(GEOMETRY_ERROR_MAX)
      expect(result.geometry).toBeUndefined()
    }, 30_000)
  }
  it('the Stage-0 synthetic text page at q70 attributes in scope', async () => {
    const c = createCanvas(384, 320), ctx = c.getContext('2d')
    ctx.fillStyle = '#eeeeee'; ctx.fillRect(0, 0, 384, 320)
    ctx.fillStyle = '#333333'; ctx.font = '14px sans-serif'
    for (let i = 0; i < 14; i++) ctx.fillText(`Office page ${i}: unique row ${i * 73}.`, 12 + i % 3 * 9, 28 + i * 20)
    ctx.fillStyle = '#8c9eab'; ctx.fillRect(280, 140, 63, 87)
    const page: Bitmap = { width: 384, height: 320, data: new Uint8ClampedArray(ctx.getImageData(0, 0, 384, 320).data) }
    const carrier = basis(new Uint8Array(32).fill(7), 384, 320, DEFAULT_PARAMS)
    const result = scoreCapture(await jpeg(embed(page, carrier, codes[TOKEN]!, DEFAULT_PARAMS), 70), page, carrier, matrix, DEFAULT_PARAMS)
    expect(result.registration!.residual).toBeLessThanOrEqual(GEOMETRY_ERROR_MAX)
    expect(decideScores(result, { abs: 6, margin: 2 })).toMatchObject({ kind: 'attributed', index: TOKEN }) // Stage-0 smoke tau
  }, 30_000)
})

describe('unsupported geometry', () => {
  for (const [name, attack] of [['rotate 2', (b: Bitmap) => rotate(b, 2)], ['rotate 5', (b: Bitmap) => rotate(b, 5)], ['keystone 0.05', (b: Bitmap) => keystone(b, 0.05)]] as const) {
    it(`${name} abstains on geometry`, () => {
      const result = scoreCapture(attack(marked), reference, chips, matrix, DEFAULT_PARAMS)
      expect(result.geometry).toBe('unsupported')
      expect(decideScores(result, TAU)).toEqual({ kind: 'abstain', reason: 'geometry' })
    }, 60_000)
  }
  it('no registration abstains on geometry', () => {
    const tiny: Bitmap = { width: 8, height: 8, data: new Uint8ClampedArray(8 * 8 * 4).fill(255) }
    expect(decideScores(scoreCapture(tiny, reference, chips, matrix, DEFAULT_PARAMS), TAU)).toEqual({ kind: 'abstain', reason: 'geometry' })
  })
})
