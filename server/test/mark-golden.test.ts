/**
 * Stage 2a Task 1: the production Layer-1 codec is byte-identical to the
 * Stage-0 prototype it was measured as (docs/validation/streaming.md T7/T9).
 * Literals were captured from scripts/stream-s0/mark before the move.
 */
import { expect, it } from 'bun:test'
import { createHash } from 'node:crypto'
import { messageFor, insertMessage } from '../src/mark/message'
import { extendedCodeword } from '../src/mark/bch'
import { basis, deriveBasisKey } from '../src/mark/basis'
import { embed } from '../src/mark/embed'
import { DEFAULT_PARAMS } from '../src/mark/params'

const sha = (b: Uint8Array | Uint8ClampedArray | string) => createHash('sha256').update(b).digest('hex')
const key = Uint8Array.from({ length: 32 }, (_, i) => i)

function synthetic(): { width: number; height: number; data: Uint8ClampedArray } {
  const width = 300, height = 200, data = new Uint8ClampedArray(width * height * 4).fill(255)
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const i = (y * width + x) * 4
    if (x >= 100 && x < 140 && y >= 60 && y < 100) { data[i] = 128; data[i + 1] = 128; data[i + 2] = 128 }
    if (y >= 150 && y < 170) data[i + 3] = 128
  }
  return { width, height, data }
}

it('golden: message, code, basis and embed match the Stage-0 prototype', () => {
  const msg = messageFor(key, 1, 0x123456, 0x5a)
  const code = extendedCodeword(msg)
  const kb = deriveBasisKey(key, 'a'.repeat(24), '3@1')
  const chips = basis(kb, 300, 200, DEFAULT_PARAMS)
  const marked = embed(synthetic(), chips, code, DEFAULT_PARAMS)
  const got = {
    msg: sha(msg), code: sha(new Uint8Array(code.buffer)), kb: sha(kb), chips: sha(JSON.stringify(chips)), data: sha(marked.data),
    mark: marked.mark,
  }
  expect(got).toEqual(GOLDEN)
})

it('golden: insertMessage re-draws a colliding salt', () => {
  const used = new Set<string>(), draws = [7, 7, 9]
  const a = insertMessage(key, 1, 5, used, () => draws.shift()!)
  const b = insertMessage(key, 1, 5, used, () => draws.shift()!)
  expect([a.salt8, b.salt8]).toEqual([7, 9])
})

const GOLDEN = {
  msg: '9ee175fcaf0bddc0706f0797f1243be7a38b78ad23065e1e96f4103fb3f3f7af',
  code: '540d8d78a80d6827f412f566a2db582f277ca809801a4a8b6f1abe2b9531a7c5',
  kb: '5afb16c132f6e49b413f07b4615d5182466135c3d03d37e9cfc247f831da4103',
  chips: '66a0b86c0e3e2fe7996e4a67ce9604818b76f13a2733792ab4f762a0f623f8e3',
  data: 'd5f0cfc9957b03a4241be2361727732596d68aae26627809b505cccadcd16837',
  mark: { variant: 'white-pedestal' as const, pedestalCells: 441, maskedCells: 55, params: { nChips: 512, cellPx: 4, tilePx: 64, amplitude: 2, pedestal: 2 } },
}

it('the Stage-0 paths re-export the production modules', async () => {
  const s0 = await import('../../scripts/stream-s0/mark/embed'), prod = await import('../src/mark/embed')
  expect(s0.embed).toBe(prod.embed)
})
