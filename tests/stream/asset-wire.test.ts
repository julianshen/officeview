import { describe, expect, it } from 'vitest'
import { AssetTable } from '../../src/stream/record/recorder'
import { ASSET_DEFINE_DECODED_CAP, AssetTooLarge, AssetWireError, decodeAssets, encodeAssets } from '../../src/stream/asset-wire'
import type { Asset } from '../../src/stream/ops'

function assets(): Asset[] {
  const t = new AssetTable()
  const rgba = new Uint8Array(3 * 2 * 4).map((_, i) => (i * 37) & 255)
  const bmp = t.add({ type: 'bitmap', source: 'image', w: 3, h: 2, rgba })
  t.add({ type: 'bitmap', source: 'canvas', w: 1, h: 1, rgba: Uint8Array.of(1, 2, 3, 4) })
  t.add({ type: 'linear', x0: 0, y0: 1.5, x1: 10, y1: -2, stops: [[0, '#ff0000'], [1, 'rgba(0, 0, 255, 0.5)']], transforms: [{ mode: 2, matrix: [3, 4] }, { mode: 4, matrix: [0.25] }] })
  t.add({ type: 'radial', x0: 1, y0: 2, r0: 0, x1: 1, y1: 2, r1: 9, stops: [[0.5, 'red']] })
  t.add({ type: 'pattern', bitmap: bmp, repetition: 'repeat-x', transform: [1, 0, 0, 1, 5, 6] })
  t.add({ type: 'pattern', bitmap: bmp, repetition: 'repeat' })
  return t.all()
}

describe('ASSET_DEFINE codec', () => {
  it('round-trips every asset kind with the recorder hash', async () => {
    const a = assets()
    const payloads = encodeAssets(a)
    expect(payloads.length).toBe(1)
    expect(await decodeAssets(payloads[0]!)).toEqual(a)
  })

  it('splits at the decoded cap and resolves pattern references across payloads', async () => {
    const a = assets()
    const payloads = encodeAssets(a, 256)
    expect(payloads.length).toBeGreaterThan(1)
    const known = new Map<number, Asset>()
    for (const p of payloads) for (const x of await decodeAssets(p, known)) known.set(x.id, x)
    expect([...known.values()]).toEqual(a)
  })

  it('a single asset over the cap throws AssetTooLarge', () => {
    expect(() => encodeAssets(assets(), 64)).toThrow(AssetTooLarge)
  })

  it('a flipped content byte fails the hash check', async () => {
    const [p] = encodeAssets(assets().slice(0, 1))
    const bad = p!.slice(); bad[bad.length - 1] ^= 1
    await expect(decodeAssets(bad)).rejects.toThrow(/hash mismatch/)
  })

  it('truncation at every offset throws and never returns a partial list', async () => {
    const [p] = encodeAssets(assets().slice(2, 4))
    for (let n = 0; n < p!.length; n++) await expect(decodeAssets(p!.subarray(0, n))).rejects.toBeInstanceOf(AssetWireError)
    const extra = new Uint8Array(p!.length + 1); extra.set(p!)
    await expect(decodeAssets(extra)).rejects.toThrow(/trailing/)
  })

  it('a pattern whose bitmap is unknown is refused', async () => {
    const a = assets()
    const [p] = encodeAssets([a[4]!])
    await expect(decodeAssets(p!)).rejects.toThrow(/pattern bitmap/)
  })
})

describe('ASSET_DEFINE hardening (review r1-5/6/7)', () => {
  it.each([
    ['NaN coordinate', { type: 'radial', x0: Number.NaN, y0: 0, r0: 0, x1: 0, y1: 0, r1: 1, stops: [[0, 'red']] }],
    ['negative radius', { type: 'radial', x0: 0, y0: 0, r0: -1, x1: 0, y1: 0, r1: 1, stops: [[0, 'red']] }],
    ['stop offset > 1', { type: 'linear', x0: 0, y0: 0, x1: 1, y1: 1, stops: [[2, 'red']] }],
    ['infinite transform', { type: 'linear', x0: 0, y0: 0, x1: 1, y1: 1, stops: [[0, 'red']], transforms: [{ mode: 2, matrix: [Number.POSITIVE_INFINITY, 0] }] }],
  ])('a correctly hashed asset with %s is refused', async (_n, input) => {
    const t = new AssetTable(); t.add(input as never)
    const [p] = encodeAssets(t.all())
    await expect(decodeAssets(p!)).rejects.toBeInstanceOf(AssetWireError)
  })
  it('an unknown pattern repetition is refused', async () => {
    const t = new AssetTable(); const b = t.add({ type: 'bitmap', source: 'image', w: 1, h: 1, rgba: Uint8Array.of(1, 2, 3, 4) })
    t.add({ type: 'pattern', bitmap: b, repetition: 'sideways' })
    await expect(decodeAssets(encodeAssets(t.all())[0]!)).rejects.toThrow(/repetition/)
  })
  it('a payload over the 8 MiB decoded cap is refused before parsing', async () => {
    await expect(decodeAssets(new Uint8Array(ASSET_DEFINE_DECODED_CAP + 1))).rejects.toThrow(/payload-cap/)
  })
  it('an oversized bitmap is refused before its serialised copy is made', () => {
    const rgba = new Uint8Array(1500 * 1500 * 4)
    const a = { id: 0, hash: 'f'.repeat(64), type: 'bitmap' as const, source: 'image' as const, w: 1500, h: 1500, rgba }
    expect(() => encodeAssets([a])).toThrow(AssetTooLarge)
  })
})
