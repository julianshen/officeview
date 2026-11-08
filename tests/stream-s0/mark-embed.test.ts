// @vitest-environment node
import { expect, it } from 'vitest'
import { deltaE00, srgbToLab, pageDeltaE } from '../../scripts/stream-s0/mark/ciede2000'
import { bchEncode, codeword40, crc8, extendedCodeword } from '../../scripts/stream-s0/mark/bch'
import { basis, deriveBasisKey } from '../../scripts/stream-s0/mark/basis'
import { embed } from '../../scripts/stream-s0/mark/embed'
import { DEFAULT_PARAMS, sweepParams } from '../../scripts/stream-s0/mark/params'
import type { Bitmap } from '../../scripts/stream-s0/mark/types'

const white = (): Bitmap => ({ width: 256, height: 256, data: new Uint8ClampedArray(256 * 256 * 4).fill(255) })
const key = new Uint8Array(32).fill(7)
it('ciede2000 matches Sharma pairs', () => {
  const pairs: [number[], number[], number][] = [
    [[50, 2.6772, -79.7751], [50, 0, -82.7485], 2.0425],
    [[50, 0, 0], [50, -1, 2], 2.3669],
    [[50, 2.5, 0], [73, 25, -18], 27.1492],
  ]
  for (const [a, b, expected] of pairs) {
    expect(Math.abs(deltaE00(a as [number, number, number], b as [number, number, number]) - expected)).toBeLessThan(1e-4)
    expect(deltaE00(a as never, b as never)).toBeCloseTo(deltaE00(b as never, a as never), 10)
  }
  expect(srgbToLab(0, 0, 0)).toEqual([0, 0, 0])
  expect(srgbToLab(255, 255, 255)[0]).toBeCloseTo(100, 3)
  expect(pageDeltaE(white(), white())).toEqual({ mean: 0, p95: 0, max: 0 })
})
it('bch codeword has min distance ≥ 191', () => {
  let state = 0x12345678
  const next = () => { state ^= state << 13; state ^= state >>> 17; state ^= state << 5; return state >>> 0 }
  let minimum = 511
  for (let i = 0; i < 1000; i++) {
    const a = Uint8Array.from({ length: 40 }, () => next() & 1)
    const b = Uint8Array.from({ length: 40 }, () => next() & 1)
    const x = bchEncode(a), y = bchEncode(b)
    expect(x.length).toBe(511)
    const distance = x.reduce((n, bit, j) => n + Number(bit !== y[j]), 0)
    minimum = Math.min(minimum, distance)
    expect(distance).toBeGreaterThanOrEqual(191)
  }
  console.log('BCH 1000 seeded independent message pairs: minimum distance', minimum)
  const zero = bchEncode(new Uint8Array(40))
  expect([...zero].every(x => x === 1)).toBe(true)
  for (let j = 0; j < 40; j++) {
    const msg = new Uint8Array(40); msg[j] = 1
    const word = bchEncode(msg)
    expect([...word].filter(x => x === -1).length).toBeGreaterThanOrEqual(191)
    expect([...word.slice(471)]).toEqual([...msg].reverse().map(x => x ? -1 : 1))
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
  let delta = 0
  for (const chip of chips) for (let y = chip.y; y < chip.y + chip.cell; y++) for (let x = chip.x; x < chip.x + chip.cell; x++) delta += Math.abs(output.data[(y * input.width + x) * 4]! - 255)
  expect(delta / (chips.length * DEFAULT_PARAMS.cellPx ** 2)).toBeGreaterThanOrEqual(0.8 * DEFAULT_PARAMS.amplitude)
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
