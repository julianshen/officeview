// @vitest-environment node
import { randomBytes } from 'node:crypto'
import { createGzip, gzipSync } from 'node:zlib'
import { once } from 'node:events'
import { describe, expect, it, vi } from 'vitest'
import { crc32 } from '../../src/stream/crc32'
import { encodeMessage, Reassembler, FRAGMENT_CAP, MAX_FRAGMENTS, REASSEMBLY_MS, HEADER_BYTES, EXT01_BYTES, type Kind, type RejectReason } from '../../src/stream/frame'

const bytes = new TextEncoder().encode('{}')
const encode = (payload = bytes, extra = {}) => encodeMessage({ ver: 1, kind: 1, reqId: 7, payload, compress: gzipSync, ...extra })
const make = (extra = {}) => new Reassembler({ acceptedVersions: new Set([1]), expects: () => true, decodedCap: () => 8 << 20, encodedCap: () => 8 << 20, ...extra })
const view = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength)
const start = (b: Uint8Array) => HEADER_BYTES + ((b[18] & 2) ? EXT01_BYTES : 0)
const repair = (b: Uint8Array) => { view(b).setUint32(14, crc32(b.subarray(start(b))), true); return b }
async function refused(r: Reassembler, b: Uint8Array, reason: RejectReason) {
  expect(await r.push(b)).toMatchObject({ ok: false, reason })
  expect(r.bufferedBytes).toBe(0)
}

describe('OVSD rejection battery', () => {
  it('pins constants', () => expect([FRAGMENT_CAP, MAX_FRAGMENTS, REASSEMBLY_MS, HEADER_BYTES, EXT01_BYTES]).toEqual([1048576, 256, 5000, 23, 10]))
  it('magic', async () => { const [b] = encode(); b[0] ^= 1; await refused(make(), b, 'magic') })
  it('version', async () => { const [b] = encode(); b[4] = 2; await refused(make(), b, 'version') })
  it.each([0, 9])('kind %i', async kind => { const [b] = encode(); b[5] = kind; await refused(make(), b, 'kind') })
  it('unexpected-reqId never creates a pending entry', async () => {
    const r = make({ expects: () => false })
    // Observe the actual insertion, not only the post-rejection byte count.
    const pending = (r as unknown as { pending: Map<number, unknown> }).pending
    const insert = vi.spyOn(pending, 'set')
    try { await refused(r, encode(bytes, { fragmentBytes: 43 })[0], 'unexpected-reqId'); expect(insert).not.toHaveBeenCalled() }
    finally { insert.mockRestore() }
  })
  it('flags', async () => { const [b] = encode(); b[18] |= 8; await refused(make(), b, 'flags') })
  it('requires gzip even in a single frame', async () => { const [b] = encode(); b[18] &= ~4; await refused(make(), b, 'flags') })
  it('invariant-mismatch', async () => { const [a, b] = encode(bytes, { fragmentBytes: 43 }); const r = make(); expect(await r.push(a)).toBeNull(); b[5] = 2; await refused(r, b, 'invariant-mismatch') })
  it('compressed invariant precedes flags', async () => { const [a, b] = encode(bytes, { fragmentBytes: 43 }); const r = make(); await r.push(a); b[18] &= ~4; await refused(r, b, 'invariant-mismatch') })
  it('sequence', async () => { const [a, b] = encode(bytes, { fragmentBytes: 43 }); const r = make(); await r.push(a); view(b).setUint32(10, 2, true); await refused(r, b, 'sequence') })
  it('rejects an orphan continuation', async () => { await refused(make(), encode(bytes, { fragmentBytes: 43 })[1], 'sequence') })
  it('final-count', async () => { const [a] = encode(bytes, { fragmentBytes: 43 }); a[18] |= 1; await refused(make(), a, 'final-count') })
  it('rejects missing final on declared final index', async () => { const [a, b] = encode(bytes, { fragmentBytes: 43 }); const r = make(); await r.push(a); b[18] &= ~1; await refused(r, b, 'final-count') })
  it('final-count precedes missing-ext01', async () => { const [b] = encode(); b[18] &= ~1; await refused(make(), b, 'final-count') })
  it('fragment-cap', async () => { const [b] = encode(); const huge = new Uint8Array(FRAGMENT_CAP + 1); huge.set(b); await refused(make(), huge, 'fragment-cap') })
  it('rejects truncated headers, extensions and payload lengths', async () => {
    const [b] = encode(bytes, { ext01: true })
    for (const length of [0, 1, 3, 4, 10, 22, 23, 32, b.length - 1]) await refused(make(), b.subarray(0, length), 'length')
  })
  it('fragment-count', async () => { const [b] = encode(bytes, { ext01: true }); view(b).setUint16(31, 257, true); await refused(make(), b, 'fragment-count') })
  it('fragment-crc', async () => { const [b] = encode(); b[start(b)] ^= 1; await refused(make(), b, 'fragment-crc') })
  it('total-len', async () => { const [a, b] = encode(bytes, { fragmentBytes: 43 }); view(a).setBigUint64(23, 27n, true); const r = make(); await r.push(a); await refused(r, b, 'total-len') })
  it('encoded-cap before allocating the declared buffer', async () => { const [b] = encode(bytes, { ext01: true }); view(b).setBigUint64(23, 2n ** 63n, true); await refused(make(), b, 'encoded-cap') })
  it('encoded-cap without ext01', async () => { await refused(make({ encodedCap: () => 4 }), encode()[0], 'encoded-cap') })
  it('decoded-cap', async () => { await refused(make({ decodedCap: () => 1 }), encode()[0], 'decoded-cap') })
  it('inflate', async () => { const [b] = encode(); b[start(b)] = 0; await refused(make(), repair(b), 'inflate') })
  it('whole-crc', async () => { const [b] = encode(); b[b.length - 1] ^= 1; await refused(make(), repair(b), 'whole-crc') })
  it('timeout discards retained fragments', async () => {
    let now = 0; const r = make({ now: () => now }); await r.push(encode(bytes, { fragmentBytes: 43 })[0]); expect(r.bufferedBytes).toBeGreaterThan(0)
    now = 5001; expect(r.sweep()).toEqual([{ ok: false, reqId: 7, reason: 'timeout' }]); expect(r.bufferedBytes).toBe(0); expect(r.sweep()).toEqual([])
  })
  it('checks expiry when a final fragment arrives without sweep', async () => {
    let now = 0; const r = make({ now: () => now }); const [a, b] = encode(bytes, { fragmentBytes: 43 }); await r.push(a); now = 5001; await refused(r, b, 'timeout')
  })
  it('aborts a streamed 512 MiB space bomb below 32 MiB RSS growth', async () => {
    // The input is generated in reusable chunks, never as a 512 MiB allocation.
    const gzip = createGzip(); const chunks: Buffer[] = []; gzip.on('data', (b: Buffer) => chunks.push(b))
    const end = once(gzip, 'end'); const chunk = Buffer.alloc(64 * 1024, 0x20)
    for (let n = 0; n < 8192; n++) if (!gzip.write(chunk)) await once(gzip, 'drain')
    gzip.end(); await end
    const compressed = Buffer.concat(chunks)
    const frames = encode(bytes, { compress: () => compressed })
    const r = make({ decodedCap: () => 256 * 1024 })
    const before = process.memoryUsage().rss; let peak = before
    const timer = setInterval(() => { peak = Math.max(peak, process.memoryUsage().rss) }, 1)
    try { for (const b of frames) await refused(r, b, 'decoded-cap') } finally { clearInterval(timer) }
    peak = Math.max(peak, process.memoryUsage().rss)
    expect(peak - before).toBeLessThan(32 * 1024 * 1024)
  }, 60_000)
})

it.each([1, 2, 3, 4, 5, 6, 7, 8])('round-trips kind %i', async kind => {
  const [b] = encode(bytes, { kind: kind as Kind }); expect(await make().push(b)).toEqual({ ok: true, reqId: 7, kind, payload: bytes })
})
it('keeps interleaved requests independent and owns received bytes', async () => {
  const a = encode(bytes, { fragmentBytes: 43 }); const b = encode(bytes, { reqId: 8, fragmentBytes: 43 }); const r = make()
  await r.push(a[0]); a[0].fill(0); await r.push(b[0]); expect(r.bufferedBytes).toBe(20)
  expect(await r.push(a[1])).toMatchObject({ ok: true, reqId: 7, payload: bytes }); expect(await r.push(b[1])).toMatchObject({ ok: true, reqId: 8, payload: bytes }); expect(r.bufferedBytes).toBe(0)
})
it('round-trips 3 MiB incompressible data through four bounded frames', async () => {
  const payload = randomBytes(3 << 20); const frames = encode(payload); expect(frames.length).toBeGreaterThanOrEqual(4)
  const r = make(); for (let i = 0; i < frames.length; i++) { expect(frames[i].length).toBeLessThanOrEqual(FRAGMENT_CAP); const result = await r.push(frames[i]); if (i < frames.length - 1) expect(result).toBeNull(); else { expect(result).toMatchObject({ ok: true, reqId: 7, kind: 1 }); if (result?.ok) expect(Buffer.compare(Buffer.from(result.payload), payload)).toBe(0) } }
})
it('keeps all four trailer bytes in the final fragment for small split sizes', async () => {
  for (let size = 34; size < 65; size++) {
    const frames = encode(bytes, { fragmentBytes: size }); const last = frames.at(-1)!; expect(view(last).getUint32(19, true)).toBeGreaterThanOrEqual(4)
    const r = make(); for (const b of frames) { expect(b.length).toBeLessThanOrEqual(size); const result = await r.push(b); if (b === last) expect(result).toMatchObject({ ok: true, payload: bytes }) }
  }
})

it('reports invariant corruption before rechecking a kind-specific expectation', async () => {
  const r = make({ expects: (id: number, kind: Kind) => id === 7 && kind === 1 })
  const [a, b] = encode(bytes, { fragmentBytes: 43 }); await r.push(a); b[5] = 2
  await refused(r, b, 'invariant-mismatch')
})
it('a rejection during asynchronous inflation prevents stale success', async () => {
  const r = make(); const [b] = encode(); const first = r.push(b)
  await refused(r, b, 'sequence')
  expect(await first).toMatchObject({ ok: false, reason: 'sequence' })
  expect(r.bufferedBytes).toBe(0)
  expect(await r.push(b)).toMatchObject({ ok: true, payload: bytes })
})


describe('fix round 1 frame isolation', () => {
  it.each(['magic', 'short', 'version', 'kind', 'reserved-flags', 'ext-continuation', 'ext-truncated', 'payload-length'])(
    'drops only the malformed %s frame, never its forged reqId', async failure => {
      const [a, b] = encode(bytes, { fragmentBytes: 43 }); const r = make()
      await r.push(a); const held = r.bufferedBytes
      let bad = b.slice(); let reason = failure
      switch (failure) {
        case 'magic': bad = new Uint8Array(23); bad.set([7, 0, 0, 0], 6); bad[0] = 1; break
        case 'short': bad = bad.slice(0, 10); reason = 'length'; break
        case 'version': bad[4] = 2; break
        case 'kind': bad[5] = 0; break
        // Even a changed invariant cannot authorize deletion when reserved flags are present.
        case 'reserved-flags': bad[18] |= 8; bad[5] = 2; reason = 'flags'; break
        case 'ext-continuation': bad[18] |= 2; reason = 'flags'; break
        case 'ext-truncated': bad = a.slice(0, 32); reason = 'length'; break
        case 'payload-length': view(bad).setUint32(19, 999, true); reason = 'length'; break
      }
      expect(await r.push(bad)).toEqual({ ok: false, reason })
      expect(r.bufferedBytes).toBe(held)
      expect(await r.push(b)).toEqual({ ok: true, reqId: 7, kind: 1, payload: bytes })
      expect(r.bufferedBytes).toBe(0)
    },
  )
  it('stops the clock at final admission, even if inflate crosses the deadline', async () => {
    let now = 0; const r = make({ now: () => now }); const [a, b] = encode(bytes, { fragmentBytes: 43 })
    await r.push(a); now = 4999; const result = r.push(b); now = 9000
    expect(await result).toEqual({ ok: true, reqId: 7, kind: 1, payload: bytes })
    expect(r.bufferedBytes).toBe(0)
  })
  it('sweep expires only incomplete messages, not an inflating complete message', async () => {
    let now = 0; const r = make({ now: () => now })
    const [a, b] = encode(bytes, { fragmentBytes: 43 })
    await r.push(a); await r.push(encode(bytes, { reqId: 8, fragmentBytes: 43 })[0])
    now = 4999; const result = r.push(b); now = 5000
    expect(r.sweep()).toEqual([{ ok: false, reqId: 8, reason: 'timeout' }])
    expect(await result).toEqual({ ok: true, reqId: 7, kind: 1, payload: bytes })
    expect(r.sweep()).toEqual([]); expect(r.bufferedBytes).toBe(0)
  })
  it('rejects zero fragment count', async () => {
    const [a] = encode(bytes, { ext01: true }); view(a).setUint16(31, 0, true)
    await refused(make(), a, 'fragment-count')
  })
  it('rejects a non-final cumulative length overflow immediately', async () => {
    const [a, b] = encode(randomBytes(100), { fragmentBytes: 43 }); const r = make()
    view(a).setBigUint64(23, 15n, true)
    expect(await r.push(a)).toBeNull(); expect(b[18] & 1).toBe(0)
    await refused(r, b, 'total-len')
  })
  it('rejects duplicate fragments and drops the affected message', async () => {
    const [a] = encode(bytes, { fragmentBytes: 43 }); const r = make()
    await r.push(a); await refused(r, a, 'sequence')
  })
})
