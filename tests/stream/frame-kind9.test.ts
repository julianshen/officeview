// @vitest-environment node
/**
 * Piece 5 Task 9 (plan D13, D14; spec §4.3): MEDIA_CHUNK (kind 9) is sent stored, and is the only kind that may be.
 * Its hostile tests: a compressed kind 9 and a stored kind 6 are `flags`; kind 10 is `kind`; the 4 MiB encoded cap
 * holds before any copy; an unsolicited chunk is dropped before allocation.
 */
import { gzipSync } from 'node:zlib'
import { describe, expect, it, vi } from 'vitest'
import { crc32 } from '../../src/stream/crc32'
import { encodeMessage, EXT01_BYTES, HEADER_BYTES, Kind, Reassembler, type RejectReason } from '../../src/stream/frame'

const view = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength)
const start = (b: Uint8Array) => HEADER_BYTES + ((b[18]! & 2) ? EXT01_BYTES : 0)
const repair = (b: Uint8Array) => { view(b).setUint32(14, crc32(b.subarray(start(b))), true); return b }
const CAP = 4 << 20
const make = (extra: Partial<ConstructorParameters<typeof Reassembler>[0]> = {}) => new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, decodedCap: () => CAP, encodedCap: () => CAP, ...extra })
const noise = (n: number) => { const b = new Uint8Array(n); let x = 7; for (let i = 0; i < n; i++) { x = (x * 1103515245 + 12345) >>> 0; b[i] = x >>> 24 } return b }
const stored = (payload: Uint8Array, extra = {}) => encodeMessage({ ver: 1, kind: Kind.MEDIA_CHUNK, reqId: 41, payload, compress: () => { throw new Error('a stored message is never compressed') }, stored: true, ...extra })
async function refused(r: Reassembler, b: Uint8Array, reason: RejectReason) {
  expect(await r.push(b)).toMatchObject({ ok: false, reason })
  expect(r.bufferedBytes).toBe(0)
}

describe('MEDIA_CHUNK (kind 9) is stored only (D14)', () => {
  it('kind 9 stored round-trips through Reassembler', async () => {
    const payload = noise(1048512) // D13: one chunk is 1 MiB − 64 B, one fragment
    const frames = stored(payload)
    expect(frames).toHaveLength(1)
    expect(frames[0]![18]! & 4).toBe(0) // the compressed flag is clear
    expect(frames[0]![5]).toBe(9)
    const r = make()
    const res = await r.push(frames[0]!)
    expect(res).toMatchObject({ ok: true, kind: 9, reqId: 41 })
    expect((res as { payload: Uint8Array }).payload).toEqual(payload)
    // A multi-fragment stored message reassembles too (the cap, not the fragment, bounds it).
    const big = noise(3 << 20)
    const many = stored(big)
    expect(many.length).toBeGreaterThan(1)
    let out: Awaited<ReturnType<Reassembler['push']>> = null
    for (const f of many) out = await r.push(f)
    expect((out as { payload: Uint8Array }).payload).toEqual(big)
    expect(r.bufferedBytes).toBe(0)
  })

  it('compressed kind 9 → flags', async () => {
    const [b] = encodeMessage({ ver: 1, kind: Kind.SHOW_PLAN, reqId: 41, payload: noise(64), compress: gzipSync })
    b![5] = 9
    await refused(make(), repair(b!), 'flags')
    // The encoder never frames one.
    expect(() => encodeMessage({ ver: 1, kind: Kind.MEDIA_CHUNK, reqId: 41, payload: noise(64), compress: gzipSync })).toThrow(RangeError)
  })

  it('stored kind 6 → flags', async () => {
    const [b] = stored(noise(64))
    b![5] = 6
    await refused(make(), repair(b!), 'flags')
    expect(() => encodeMessage({ ver: 1, kind: Kind.SHOW_PLAN, reqId: 41, payload: noise(64), compress: gzipSync, stored: true })).toThrow(RangeError)
  })

  it('kind 10 → kind', async () => {
    const [b] = stored(noise(64))
    b![5] = 10
    await refused(make(), repair(b!), 'kind')
    expect(() => encodeMessage({ ver: 1, kind: 10 as Kind, reqId: 41, payload: noise(64), compress: gzipSync })).toThrow(RangeError)
  })

  it('kind 9 over 4 MiB encoded → encoded-cap', async () => {
    const frames = stored(noise(CAP)) // CAP + 4 bytes of whole CRC: one byte… four over the cap
    const r = make()
    const pending = (r as unknown as { pending: Map<number, unknown> }).pending
    const insert = vi.spyOn(pending, 'set')
    try {
      await refused(r, frames[0]!, 'encoded-cap')
      expect(insert).not.toHaveBeenCalled()
    } finally { insert.mockRestore() }
    // Exactly at the cap (payload + its CRC) is admitted.
    const at = stored(noise(CAP - 4))
    const ok = make()
    let out: Awaited<ReturnType<Reassembler['push']>> = null
    for (const f of at) out = await ok.push(f)
    expect(out).toMatchObject({ ok: true, kind: 9 })
  })

  it('kind 9 over its decoded cap → decoded-cap', async () => {
    await refused(make({ decodedCap: () => 100 }), stored(noise(101))[0]!, 'decoded-cap')
  })

  it('a stored kind 9 with a bad whole CRC → whole-crc', async () => {
    const [b] = stored(noise(64))
    view(b!).setUint32(b!.length - 4, 0xdeadbeef, true)
    await refused(make(), repair(b!), 'whole-crc')
  })

  it('unsolicited kind 9 → unexpected-reqId before allocation', async () => {
    const r = make({ expects: (_id, kind) => kind !== Kind.MEDIA_CHUNK })
    const pending = (r as unknown as { pending: Map<number, unknown> }).pending
    const insert = vi.spyOn(pending, 'set')
    try {
      await refused(r, stored(noise(1048512))[0]!, 'unexpected-reqId')
      expect(insert).not.toHaveBeenCalled()
    } finally { insert.mockRestore() }
  })
})
