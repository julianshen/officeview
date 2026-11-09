// @vitest-environment node
import { expect, it } from 'vitest'
import { deltaE00, srgbToLab, pageDeltaE } from '../../scripts/stream-s0/mark/ciede2000'
import { bchEncode, codeword40, crc8, extendedCodeword } from '../../scripts/stream-s0/mark/bch'
import { basis, deriveBasisKey } from '../../scripts/stream-s0/mark/basis'
import { embed, pedestalFor } from '../../scripts/stream-s0/mark/embed'
import { DEFAULT_PARAMS, sweepParams } from '../../scripts/stream-s0/mark/params'
import type { Bitmap } from '../../scripts/stream-s0/mark/types'

const white = (): Bitmap => ({ width: 256, height: 256, data: new Uint8ClampedArray(256 * 256 * 4).fill(255) })
const key = new Uint8Array(32).fill(7)
it('ciede2000 matches Sharma pairs', () => {
  // Sharma–Wu–Dalal supplementary 34-pair dataset, offline transcription.
  const pairs = `50 2.6772 -79.7751 50 0 -82.7485 2.0425
50 3.1571 -77.2803 50 0 -82.7485 2.8615
50 2.8361 -74.0200 50 0 -82.7485 3.4412
50 -1.3802 -84.2814 50 0 -82.7485 1.0000
50 -1.1848 -84.8006 50 0 -82.7485 1.0000
50 -0.9009 -85.5211 50 0 -82.7485 1.0000
50 0 0 50 -1 2 2.3669
50 -1 2 50 0 0 2.3669
50 2.49 -0.001 50 -2.49 0.0009 7.1792
50 2.49 -0.001 50 -2.49 0.0010 7.1792
50 2.49 -0.001 50 -2.49 0.0011 7.2195
50 2.49 -0.001 50 -2.49 0.0012 7.2195
50 -0.001 2.49 50 0.0009 -2.49 4.8045
50 -0.001 2.49 50 0.0010 -2.49 4.8045
50 -0.001 2.49 50 0.0011 -2.49 4.7461
50 2.5 0 50 0 -2.5 4.3065
50 2.5 0 73 25 -18 27.1492
50 2.5 0 61 -5 29 22.8977
50 2.5 0 56 -27 -3 31.9030
50 2.5 0 58 24 15 19.4535
50 2.5 0 50 3.1736 0.5854 1
50 2.5 0 50 3.2972 0 1
50 2.5 0 50 1.8634 0.5757 1
50 2.5 0 50 3.2592 0.3350 1
60.2574 -34.0099 36.2677 60.4626 -34.1751 39.4387 1.2644
63.0109 -31.0961 -5.8663 62.8187 -29.7946 -4.0864 1.2630
61.2901 3.7196 -5.3901 61.4292 2.2480 -4.9620 1.8731
35.0831 -44.1164 3.7933 35.0232 -40.0716 1.5901 1.8645
22.7233 20.0904 -46.6940 23.0331 14.9730 -42.5619 2.0373
36.4612 47.8580 18.3852 36.2715 50.5065 21.2231 1.4146
90.8027 -2.0831 1.4410 91.1528 -1.6435 0.0447 1.4441
90.9257 -0.5406 -0.9208 88.6381 -0.8985 -0.7239 1.5381
6.7747 -0.2908 -2.4247 5.8714 -0.0985 -2.2286 0.6377
2.0776 0.0795 -1.1350 0.9033 -0.0636 -0.5514 0.9082`.split('\n').map(line => { const v = line.split(' ').map(Number); return [v.slice(0, 3), v.slice(3, 6), v[6]!] as [number[], number[], number] })
  for (const [a, b, expected] of pairs) {
    expect(Math.abs(deltaE00(a as [number, number, number], b as [number, number, number]) - expected)).toBeLessThan(1e-4)
    expect(deltaE00(a as never, b as never)).toBeCloseTo(deltaE00(b as never, a as never), 10)
  }
  expect(srgbToLab(0, 0, 0)).toEqual([0, 0, 0])
  expect(srgbToLab(255, 255, 255)[0]).toBeCloseTo(100, 3)
  expect(pageDeltaE(white(), white())).toEqual({ mean: 0, p95: 0, max: 0, chipPixelsP95: 0 })
})
it('BCH membership and linearity establish the designed distance bound', () => {
  // Primitive narrow-sense BCH roots 1..190 imply d >= 191, not a
  // minimum inferred from random distances concentrated near n/2.
  const mul = (a: number, b: number) => {
    let r = 0
    while (b) { if (b & 1) r ^= a; b >>>= 1; a <<= 1; if (a & 512) a ^= 0x211 }
    return r
  }
  let state = 0x12345678
  const next = () => { state ^= state << 13; state ^= state >>> 17; state ^= state << 5; return state >>> 0 }
  for (let trial = 0; trial < 100; trial++) {
    const a = Uint8Array.from({ length: 40 }, () => next() & 1), b = Uint8Array.from({ length: 40 }, () => next() & 1)
    const x = bchEncode(a), y = bchEncode(b), z = bchEncode(a.map((v, i) => v ^ b[i]!))
    expect([...z]).toEqual([...x].map((v, i) => v * y[i]!))
    let root = 1
    for (let exponent = 1; exponent <= 190; exponent++) {
      root = mul(root, 2)
      let value = 0
      for (let j = 510; j >= 0; j--) value = mul(value, root) ^ Number(x[j] === -1)
      expect(value).toBe(0)
    }
  }
  expect([...extendedCodeword(codeword40(1, 2))].reduce((a, b) => a * b, 1)).toBe(1)
})
it('crc8 detects single-bit flips', () => {
  expect(crc8(new TextEncoder().encode('123456789'))).toBe(0xf4)
  const bits = codeword40(0x123456, 0x78)
  const bytes = (b: Uint8Array) => Uint8Array.from({ length: 5 }, (_, i) => [...b.slice(i * 8, i * 8 + 8)].reduce((v, bit) => v * 2 + bit, 0))
  expect([...bytes(bits).slice(0, 4)]).toEqual([0x12, 0x34, 0x56, 0x78])
  expect(crc8(bytes(bits))).toBe(0)
  for (let i = 0; i < 40; i++) { const bad = bits.slice(); bad[i] ^= 1; expect(crc8(bytes(bad))).not.toBe(0) }
  expect(() => codeword40(2 ** 24, 0)).toThrow()
  expect(() => bchEncode(new Uint8Array(39))).toThrow()
})
it('white page keeps chip power', () => {
  const input = white(), chips = basis(key, input.width, input.height, DEFAULT_PARAMS)
  const output = embed(input, chips, extendedCodeword(codeword40(1, 2)), DEFAULT_PARAMS)
  const signs = chips.map((c, i) => c.sign0 * extendedCodeword(codeword40(1, 2))[i]!)
  const deltas = Array.from(output.chipDelta)
  const mean = deltas.reduce((a, b) => a + b, 0) / deltas.length
  const modulation = deltas.reduce((sum, d, i) => sum + (d - mean) * signs[i]!, 0) / deltas.length
  expect(modulation).toBeGreaterThanOrEqual(0.8 * DEFAULT_PARAMS.amplitude)
  expect(input.data.every(x => x === 255)).toBe(true)
  expect(pageDeltaE(input, output).max).toBeLessThanOrEqual(3)
  expect(output.mark?.pedestalCells).toBe(chips.length)
})
it('embed is deterministic per key/unit', () => {
  const k = deriveBasisKey(key, 'doc', 'unit'), other = deriveBasisKey(key, 'doc', 'other')
  const chips = basis(k, 256, 256, DEFAULT_PARAMS)
  expect(chips).toEqual(basis(k, 256, 256, DEFAULT_PARAMS))
  expect(chips).not.toEqual(basis(other, 256, 256, DEFAULT_PARAMS))
  expect(new Set(chips.map(c => `${c.x},${c.y}`)).size).toBe(512)
  const code = extendedCodeword(codeword40(123, 42))
  expect(embed(white(), chips, code, DEFAULT_PARAMS)).toEqual(embed(white(), chips, code, DEFAULT_PARAMS))
  expect(() => basis(k, 8, 8, DEFAULT_PARAMS)).toThrow()
  expect([...sweepParams()]).toHaveLength(24)
})

it('basis runs under the Bun probe runtime as well as Node', async () => {
  const { execFileSync } = await import('node:child_process')
  const root = new URL('../../', import.meta.url)
  expect(() => execFileSync('bun', ['-e', 'import {basis} from "./scripts/stream-s0/mark/basis"; import {DEFAULT_PARAMS} from "./scripts/stream-s0/mark/params"; if(basis(new Uint8Array(32).fill(7),256,256,DEFAULT_PARAMS).length!==512) throw Error("bad basis")'], { cwd: root, stdio: 'pipe' })).not.toThrow()
})

it('pedestal reserves headroom in every channel or masks the cell', () => {
  const cell = (rgb: number[]) => ({ width: 1, height: 1, data: Uint8ClampedArray.from([...rgb, 255]) })
  expect(pedestalFor(cell([255, 255, 200]), DEFAULT_PARAMS)).toBe(-2)
  expect(pedestalFor(cell([0, 0, 20]), DEFAULT_PARAMS)).toBe(2)
  expect(pedestalFor(cell([255, 0, 0]), DEFAULT_PARAMS)).toBeNull()
  expect(pedestalFor(cell([255, 255, 255]), { ...DEFAULT_PARAMS, pedestal: 0 })).toBeNull()
})
