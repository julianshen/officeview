// @vitest-environment node
import { gzipSync } from 'node:zlib'
import { expect, it } from 'vitest'
import { encodeMessage, Reassembler } from '../../src/stream/frame'

// Verbatim §16 T6 vectors, including gzip OS byte ff.
const vectors = [
  ['4f56534401010700000000000000b948950d051a0000001f8b08000000000002ffabae050043bfa6a30200000043bfa6a3'],
  ['4f56534401010700000000000000b948950d071a0000001a0000000000000001001f8b08000000000002ffabae050043bfa6a30200000043bfa6a3'],
  ['4f565344010107000000000000003bbafe1d060a0000001a0000000000000002001f8b08000000000002ff',
    '4f5653440101070000000100000018c5d5120510000000abae050043bfa6a30200000043bfa6a3'],
]
const compress = (b: Uint8Array) => { const g = gzipSync(b, { level: 9 }); g[9] = 0xff; return g }
for (const [i, hexes] of vectors.entries()) {
  it(`decodes and reproduces golden vector ${i + 1}`, async () => {
    const r = new Reassembler({ acceptedVersions: new Set([1]), expects: (id, k) => id === 7 && k === 1, decodedCap: () => 256 * 1024, encodedCap: () => 256 * 1024 })
    for (let n = 0; n < hexes.length; n++) {
      const result = await r.push(Buffer.from(hexes[n], 'hex'))
      expect(result).toEqual(n < hexes.length - 1 ? null : { ok: true, kind: 1, reqId: 7, payload: new TextEncoder().encode('{}') })
    }
    const encoded = encodeMessage({ ver: 1, kind: 1, reqId: 7, payload: new TextEncoder().encode('{}'), compress, ext01: i === 1, ...(i === 2 ? { fragmentBytes: 43 } : {}) })
    expect(encoded.map(b => Buffer.from(b).toString('hex'))).toEqual(hexes)
    expect(r.bufferedBytes).toBe(0)
  })
}
