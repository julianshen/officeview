// @vitest-environment node
import { expect, it } from 'vitest'
import { writeFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { createCanvas } from '../../scripts/stream-s0/napi-shim'
import { basis } from '../../scripts/stream-s0/mark/basis'
import { codeword40, extendedCodeword } from '../../scripts/stream-s0/mark/bch'
import { embed } from '../../scripts/stream-s0/mark/embed'
import { DEFAULT_PARAMS as p } from '../../scripts/stream-s0/mark/params'
import { detect, scoreCapture } from '../../scripts/stream-s0/mark/detect'
import { jpeg, scale, crop, compose, average, median, difference } from '../../scripts/stream-s0/mark/attacks'
import type { Bitmap } from '../../scripts/stream-s0/mark/types'
function page(): Bitmap {
  const c = createCanvas(384, 320), ctx = c.getContext('2d')
  ctx.fillStyle = '#eeeeee'; ctx.fillRect(0, 0, 384, 320)
  ctx.fillStyle = '#333333'; ctx.font = '14px sans-serif'
  for (let i = 0; i < 14; i++) ctx.fillText(`Office page ${i}: unique row ${i * 73}.`, 12 + i % 3 * 9, 28 + i * 20)
  ctx.fillStyle = '#8c9eab'; ctx.fillRect(280, 140, 63, 87)
  return { width: 384, height: 320, data: new Uint8ClampedArray(ctx.getImageData(0, 0, 384, 320).data) }
}
let seed = 0x12345678
const random = () => { seed ^= seed << 13; seed ^= seed >>> 17; seed ^= seed << 5; return seed >>> 0 }
const candidates = Array.from({ length: 1024 }, () => extendedCodeword(codeword40(random() % 2 ** 24, random() % 256)))
const reference = page(), chips = basis(new Uint8Array(32).fill(7), reference.width, reference.height, p)
const marked = embed(reference, chips, candidates[73]!, p)
const tau = { abs: 6, margin: 2 } // smoke thresholds only: Task 8 calibrates full-search attribution.
it('flat capture abstains', () => expect(detect(reference, reference, chips, candidates, p, tau)).toMatchObject({ kind: 'abstain', reason: 'flat' }))
it('clean marked page attributes the right token among M=1024', () => expect(detect(marked, reference, chips, candidates, p, tau)).toMatchObject({ kind: 'attributed', index: 73 }))
it('unmarked page → none (small acquisition noise)', () => {
  const noisy = { ...reference, data: reference.data.slice() }
  for (let i = 0; i < noisy.data.length; i += 4) { const n = random() % 3 - 1; for (let c = 0; c < 3; c++) noisy.data[i + c] += n }
  expect(detect(noisy, reference, chips, candidates, p, tau).kind).toBe('none')
})
it('25% crop of 50% downscale registers', () => {
  for (const anchor of ['top-left', 'center', 'bottom-right'] as const) {
    const capture = crop(scale(marked, 0.5), 0.25, anchor)
    expect(detect(capture, reference, chips, candidates, p, tau)).toMatchObject({ kind: 'attributed', index: 73 })
  }
}, 30_000)
it('jpeg q70 still attributes on a synthetic text fixture page', async () => {
  expect(detect(await jpeg(marked, 70), reference, chips, candidates, p, tau)).toMatchObject({ kind: 'attributed', index: 73 })
})
it('pedestal DC does not inflate null Z', () => {
  const white: Bitmap = { width: 256, height: 256, data: new Uint8ClampedArray(256 * 256 * 4).fill(255) }
  const carrier = basis(new Uint8Array(32).fill(9), 256, 256, p)
  const groups: number[][] = [[], []]
  for (let k = 0; k < 12; k++) {
    const noisy = { ...white, data: white.data.slice() }
    for (let i = 0; i < noisy.data.length; i += 4) { const n = random() & 1; for (let c = 0; c < 3; c++) noisy.data[i + c] -= n }
    const wrong = embed(white, basis(new Uint8Array(32).fill(20 + k), 256, 256, p), candidates[k]!, p)
    groups[0]!.push(...scoreCapture(noisy, white, carrier, candidates, p).scores)
    groups[1]!.push(...scoreCapture(wrong, white, carrier, candidates, p).scores)
  }
  const measurements: { kind: string; samples: number; mean: number; std: number }[] = []
  for (const [i, z] of groups.entries()) {
    const mean = z.reduce((a, b) => a + b, 0) / z.length
    const std = Math.sqrt(z.reduce((a, b) => a + (b - mean) ** 2, 0) / z.length)
    measurements.push({ kind: i === 0 ? 'unmarked' : 'wrong-key', samples: z.length, mean, std })
    console.log(`null ${i === 0 ? 'unmarked' : 'wrong-key'} N=${z.length} mean=${mean} std=${std}`)
    expect(std).toBeGreaterThanOrEqual(0.8); expect(std).toBeLessThanOrEqual(1.25)
  }
  if (process.env.S0_MARK_WRITE_RESULTS === '1') writeFileSync(new URL('../../docs/validation/streaming-s0/task7-null.json', import.meta.url), JSON.stringify({ gate: 'T9-null-smoke', status: 'GREEN', backend: '@napi-rs/canvas', gitSha: execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim(), generatedAt: new Date().toISOString(), measurements, note: '12 captures per null class, 1024 real BCH candidates; no full-search false-attribution calibration claimed' }, null, 2) + '\n')
}, 30_000)
it('attack combinators preserve shape and exact arithmetic', async () => {
  const pixel = (n: number): Bitmap => ({ width: 1, height: 1, data: Uint8ClampedArray.of(n, n, n, 255) })
  expect(average([pixel(10), pixel(20)]).data[0]).toBe(15)
  expect(median([pixel(10), pixel(20), pixel(90), pixel(100)]).data[0]).toBe(55)
  expect(median([pixel(10), pixel(20), pixel(90)]).data[0]).toBe(20)
  expect(difference(pixel(10), pixel(20)).data[0]).toBe(118)
  const result = await compose(b => scale(b, 0.5), b => crop(b, 0.25, 'center'))(marked)
  expect(result.width).toBe(144); expect(result.height).toBe(120)
  expect(() => average([])).toThrow()
})

it('2x scale and crop retain original carrier indexing', () => {
  const capture = crop(scale(marked, 2), 0.25, 'center')
  expect(detect(capture, reference, chips, candidates, p, tau)).toMatchObject({ kind: 'attributed', index: 73 })
})
