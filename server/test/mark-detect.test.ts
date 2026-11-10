/** Stage 2b Task 1: the detector in server/src/mark, register once / score many. */
import { describe, expect, it } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { basis, deriveBasisKey } from '../src/mark/basis'
import { extendedCodeword } from '../src/mark/bch'
import { decideScores, prepareCandidates, scoreCapture, scoreRegistered } from '../src/mark/detect'
import { messageFor } from '../src/mark/message'
import { markRaster } from '../src/mark/mark-raster'
import { DEFAULT_PARAMS } from '../src/mark/params'
import { decodePng } from '../src/mark/png'
import { alignReference, register } from '../src/mark/register'
import type { Bitmap } from '../src/mark/types'
import { DocumentHost } from '../src/render'
import { crop, scale } from '../../scripts/stream-s0/mark/attacks'

const root = resolve(import.meta.dir, '../..')
const TAU = { abs: 7.99, margin: 3.85 }
const key = Uint8Array.from({ length: 32 }, (_, i) => i + 1)
const codes = Array.from({ length: 64 }, (_, token) => extendedCodeword(messageFor(key, 1, token, 17)))
const matrix = prepareCandidates(codes)
const bmp = (d: { w: number; h: number; rgba: Uint8ClampedArray }): Bitmap => ({ width: d.w, height: d.h, data: d.rgba })

const host = new DocumentHost()
const doc = await host.open('pydocx', new Uint8Array(readFileSync(resolve(root, 'corpus/pydocx-test.docx'))), 'visible-only')
const kBasis = deriveBasisKey(key, 'docref-poi', '0@1')
function page(text: string, token: number) {
  const unit = host.render(doc.docKey, { unitIndex: 0, scale: 1, projection: 'final-raster', mark: { kind: 'viewer', text } }) as { png: Uint8Array; w: number; h: number }
  const out = markRaster({ png: unit.png, w: unit.w, h: unit.h, kBasis, code: codes[token]! })
  if ('unmarked' in out) throw new Error(`unexpected unmarked: ${out.unmarked}`)
  return { reference: bmp(decodePng(unit.png)), marked: bmp(decodePng(out.png)), chips: basis(kBasis, unit.w, unit.h, DEFAULT_PARAMS) }
}
const A = page('a', 3), B = page('b', 5)

describe('alignReference', () => {
  for (const [name, attack] of [['crop25', (b: Bitmap) => crop(b, 0.25)], ['down0.5', (b: Bitmap) => scale(b, 0.5)], ['identity', (b: Bitmap) => b], ['crop10-down0.5', (b: Bitmap) => scale(crop(b, 0.1), 0.5)], ['crop15-br-down0.5', (b: Bitmap) => scale(crop(b, 0.15, 'bottom-right'), 0.5)]] as const) {
    it(`reproduces register's aligned reference byte for byte (${name})`, () => {
      const cap = attack(A.marked), reg = register(cap, A.reference, A.chips)
      if (!reg) throw new Error('registration failed')
      const aligned = alignReference(A.reference, reg, cap.width, cap.height)
      expect([aligned.width, aligned.height]).toEqual([cap.width, cap.height])
      expect(Buffer.from(aligned.data).equals(Buffer.from(reg.reference.data))).toBe(true)
    })
  }
})

describe('Pose.order', () => {
  it('names the construction when a pose is valid both ways', () => {
    // Crop then exactly 2x: an integer origin at scale 2 fits both orders, but
    // the two resamples differ at the crop edges.
    const cap = scale(crop(A.marked, 0.25), 2), reg = register(cap, A.reference, A.chips)
    if (!reg) throw new Error('registration failed')
    expect([reg.scaleX, reg.scaleY, reg.order]).toEqual([2, 2, 'crop-scale'])
    expect(Number.isInteger(reg.x) && Number.isInteger(reg.y)).toBe(true)
    const same = alignReference(A.reference, reg, cap.width, cap.height)
    const other = alignReference(A.reference, { ...reg, order: reg.order === 'scale-crop' ? 'crop-scale' : 'scale-crop' }, cap.width, cap.height)
    expect(Buffer.from(same.data).equals(Buffer.from(reg.reference.data))).toBe(true)
    expect(Buffer.from(other.data).equals(Buffer.from(reg.reference.data))).toBe(false)
  })
  it('rejects a scale-crop pose whose scale is not the resampled extent', () => {
    expect(() => alignReference(A.reference, { factor: 0.5, x: 0, y: 0, scaleX: 0.4, scaleY: 0.5, order: 'scale-crop' }, 100, 100)).toThrow(RangeError)
  })
})

describe('scoreRegistered', () => {
  for (const [name, attack] of [['crop25', (b: Bitmap) => crop(b, 0.25)], ['down0.5', (b: Bitmap) => scale(b, 0.5)]] as const) {
    it(`equals scoreCapture (${name})`, () => {
      const cap = attack(A.marked), reg = register(cap, A.reference, A.chips)
      if (!reg) throw new Error('registration failed')
      const whole = scoreCapture(cap, A.reference, A.chips, matrix, DEFAULT_PARAMS)
      const split = scoreRegistered(cap, reg, A.reference, A.chips, matrix, DEFAULT_PARAMS)
      expect(Array.from(split.scores)).toEqual(Array.from(whole.scores))
      expect(split.support).toBe(whole.support)
      expect(split.flat).toBe(whole.flat)
      expect(decideScores(split, TAU)).toMatchObject({ kind: 'attributed', index: 3 })
    })
  }
  it('reuses one pose across visible-mark variants', () => {
    const cap = crop(B.marked, 0.25)
    const reg = register(cap, A.reference, A.chips) // the pose is found against the WRONG visible mark
    if (!reg) throw new Error('registration failed')
    const decision = decideScores(scoreRegistered(cap, reg, B.reference, B.chips, matrix, DEFAULT_PARAMS), TAU)
    expect(decision).toMatchObject({ kind: 'attributed', index: 5 })
  })
  it('a capture with no registration still reports an empty score set through scoreCapture', () => {
    const tiny: Bitmap = { width: 8, height: 8, data: new Uint8ClampedArray(8 * 8 * 4).fill(255) }
    const r = scoreCapture(tiny, A.reference, A.chips, matrix, DEFAULT_PARAMS)
    expect(r.scores.length).toBe(0)
    expect(r.registration).toBeUndefined()
  })
})
