/** S1.1b image admission before raster decode (design note §4.3 "Image parts"; RED test 14b). */
import { describe, expect, it } from 'vitest'
import { bytesRangeSource, readZipInventory, type RangeSource } from '../src/core/zip-stream'
import { admitImagePart, imageLimitsFor, IMAGE_PROBE_BYTES, probeImageHeader } from '../src/xlsx/storage/imageAdmission'
import { createQuotaLedger, quotaLimitsFor } from '../src/xlsx/storage/quota'
import { buildZip } from './helpers/zip-builder'

function png(width: number, height: number, tail = 4096): Uint8Array {
  const out = new Uint8Array(33 + tail)
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52])
  const v = new DataView(out.buffer)
  v.setUint32(16, width)
  v.setUint32(20, height)
  out[24] = 8
  out[25] = 6
  return out
}

function jpeg(width: number, height: number): Uint8Array {
  // SOI, APP0 (16 bytes), SOF0
  const app0 = [0xff, 0xe0, 0, 16, 0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0]
  const sof = [0xff, 0xc0, 0, 17, 8, (height >> 8) & 0xff, height & 0xff, (width >> 8) & 0xff, width & 0xff, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1]
  return new Uint8Array([0xff, 0xd8, ...app0, ...sof, ...new Array(2000).fill(0)])
}

async function part(name: string, data: Uint8Array, method = 8): Promise<{ src: RangeSource & { read: RangeSource['read'] }; entry: Awaited<ReturnType<typeof readZipInventory>>['entries'][number]; reads: number[] }> {
  const bytes = buildZip([{ name, data, method }]).bytes
  const inner = bytesRangeSource(bytes)
  const reads: number[] = []
  const src = { size: inner.size, read(a: number, b: number, s?: AbortSignal) { reads.push(b - a); return inner.read(a, b, s) } }
  const inv = await readZipInventory(src, { maxEntries: 9, maxCentralDirectoryBytes: 1e6, maxNameBytes: 999, maxPartExpandedBytes: 2 ** 40, maxTotalExpandedBytes: 2 ** 40, maxRatio: 1e6 })
  return { src, entry: inv.entries[0], reads }
}

const ledger = () => createQuotaLedger(quotaLimitsFor('D'), { maxTotalExpandedBytes: 1e9 })
const signal = () => new AbortController().signal

describe('header probe', () => {
  it('reads dimensions from PNG, JPEG, GIF, BMP, WebP and TIFF headers', () => {
    expect(probeImageHeader(png(640, 480))).toEqual({ format: 'png', width: 640, height: 480 })
    expect(probeImageHeader(jpeg(1024, 768))).toEqual({ format: 'jpeg', width: 1024, height: 768 })
    const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x20, 0x03, 0x58, 0x02, 0, 0])
    expect(probeImageHeader(gif)).toEqual({ format: 'gif', width: 800, height: 600 })
    const bmp = new Uint8Array(54)
    bmp.set([0x42, 0x4d])
    new DataView(bmp.buffer).setUint32(14, 40, true)
    new DataView(bmp.buffer).setInt32(18, 300, true)
    new DataView(bmp.buffer).setInt32(22, -200, true)
    expect(probeImageHeader(bmp)).toEqual({ format: 'bmp', width: 300, height: 200 })
    const webp = new Uint8Array(32)
    webp.set([...'RIFF'].map((c) => c.charCodeAt(0)), 0)
    webp.set([...'WEBPVP8X'].map((c) => c.charCodeAt(0)), 8)
    webp.set([0x3f, 0x1f, 0x00, 0xdf, 0x0e, 0x00], 24) // 8000 x 3808
    expect(probeImageHeader(webp)).toEqual({ format: 'webp', width: 8000, height: 3808 })
    const tif = new Uint8Array(64)
    const t = new DataView(tif.buffer)
    tif.set([0x49, 0x49, 0x2a, 0])
    t.setUint32(4, 8, true)
    t.setUint16(8, 2, true)
    t.setUint16(10, 256, true); t.setUint16(12, 4, true); t.setUint32(14, 1, true); t.setUint32(18, 5000, true)
    t.setUint16(22, 257, true); t.setUint16(24, 3, true); t.setUint32(26, 1, true); t.setUint16(30, 4000, true)
    expect(probeImageHeader(tif)).toEqual({ format: 'tiff', width: 5000, height: 4000 })
  })

  it('review MINOR 3: BMP width is unsigned (no 32-bit wrap); height sign only flips orientation', () => {
    const bmp = (w: number, h: number) => {
      const b = new Uint8Array(54)
      b.set([0x42, 0x4d])
      const v = new DataView(b.buffer)
      v.setUint32(14, 40, true)
      v.setUint32(18, w >>> 0, true)
      v.setInt32(22, h, true)
      return b
    }
    expect(probeImageHeader(bmp(0xffffff00, 10))).toEqual({ format: 'bmp', width: 0xffffff00, height: 10 })
    expect(probeImageHeader(bmp(0xffffffff, 10))).toEqual({ format: 'bmp', width: 0xffffffff, height: 10 })
    expect(probeImageHeader(bmp(300, -2_147_483_648))).toEqual({ format: 'bmp', width: 300, height: 2_147_483_648 })
  })

  it('unknown and truncated headers are reported, never guessed', () => {
    expect(probeImageHeader(new Uint8Array(100).fill(7))).toEqual({ error: 'unknown-format' })
    expect(probeImageHeader(png(10, 10).subarray(0, 12))).toEqual({ error: 'truncated' })
    expect(probeImageHeader(jpeg(10, 10).subarray(0, 10))).toEqual({ error: 'truncated' })
    expect(probeImageHeader(png(0, 10))).toEqual({ error: 'invalid' })
  })
})

describe('admission before decode (RED test 14b)', () => {
  it('a PNG declaring 100,000 x 100,000 pixels in a few KB is never decoded: <= 64 KiB probed, placeholder, nothing reserved', async () => {
    const { src, entry, reads } = await part('xl/media/image1.png', png(100_000, 100_000, 500_000))
    expect(entry.compressedSize).toBeLessThan(8 * 1024)
    const l = ledger()
    const res = await admitImagePart(src, entry, l, imageLimitsFor('D'), signal())
    expect(res.kind).toBe('placeholder')
    expect(res.kind === 'placeholder' && res.diagnostic.code).toBe('image-too-large')
    expect(res.probedBytes).toBeLessThanOrEqual(IMAGE_PROBE_BYTES)
    expect(Math.max(...reads)).toBeLessThanOrEqual(IMAGE_PROBE_BYTES)
    expect(l.snapshot().peak.decodedBuffers).toBe(0)
  })

  it('a JPEG with an equivalent SOF0 is a placeholder too (stored entry)', async () => {
    const { src, entry } = await part('xl/media/image2.jpeg', jpeg(60_000, 60_000), 0)
    const l = ledger()
    const res = await admitImagePart(src, entry, l, imageLimitsFor('M'), signal())
    expect(res.kind === 'placeholder' && res.header).toEqual({ format: 'jpeg', width: 60_000, height: 60_000 })
    expect(l.snapshot().peak.decodedBuffers).toBe(0)
  })

  it('an image just under the cap reserves 2 x decodedBytes and releases the transient half after decode', async () => {
    const { src, entry } = await part('xl/media/image3.png', png(4096, 4095))
    const l = ledger()
    const res = await admitImagePart(src, entry, l, imageLimitsFor('D'), signal())
    expect(res.kind).toBe('decode')
    if (res.kind !== 'decode') return
    expect(res.decodedBytes).toBe(4096 * 4095 * 4)
    expect(l.snapshot().current.decodedBuffers).toBe(2 * res.decodedBytes)
    res.releaseTransient()
    res.releaseTransient()
    expect(l.snapshot().current.decodedBuffers).toBe(res.decodedBytes)
    res.surface.release()
    expect(l.snapshot().current.decodedBuffers).toBe(0)
  })

  it('the pixel cap is exact: 16,777,216 pixels decode on D, one more row does not', async () => {
    const at = await part('a.png', png(4096, 4096))
    expect((await admitImagePart(at.src, at.entry, ledger(), imageLimitsFor('D'), signal())).kind).toBe('decode')
    const over = await part('b.png', png(4096, 4097))
    expect((await admitImagePart(over.src, over.entry, ledger(), imageLimitsFor('D'), signal())).kind).toBe('placeholder')
  })

  it('unknown or truncated headers become placeholders', async () => {
    const unknown = await part('xl/media/x.bin', new Uint8Array(5000).fill(3))
    const u = await admitImagePart(unknown.src, unknown.entry, ledger(), imageLimitsFor('D'), signal())
    expect(u.kind === 'placeholder' && u.diagnostic.code).toBe('image-unknown-format')
    const cut = await part('xl/media/y.png', png(10, 10).subarray(0, 14), 0)
    const c = await admitImagePart(cut.src, cut.entry, ledger(), imageLimitsFor('D'), signal())
    expect(c.kind === 'placeholder' && c.diagnostic.code).toBe('image-truncated')
  })
})
