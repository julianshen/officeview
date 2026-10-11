/**
 * Image admission before any raster decode (design note §4.3 "Image parts",
 * spec §12.4): read at most 64 KiB of the part, parse the dimensions from the
 * format header, and admit the decode only if its pixel area and RGBA size
 * fit; the decode reserves 2 x decodedBytes (surface plus transient) from
 * the ledger's decodedBuffers pool. Anything else is a diagnosed placeholder.
 * A compressed byte cap alone does not bound decoder memory.
 */
import { throwIfAborted } from '../../core/package-error'
import { openPart, type RangeSource, type ZipEntryInfo } from '../../core/zip-stream'
import type { ProfileId, WorkerDiagnostic } from '../view/contracts'
import type { QuotaLedger, Reservation } from './quota'

export type ImageFormat = 'png' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tiff'
export interface ImageHeader { format: ImageFormat; width: number; height: number }
export type ProbeResult = ImageHeader | { error: 'unknown-format' | 'truncated' | 'invalid' }

export const IMAGE_PROBE_BYTES = 64 * 1024

export interface ImageLimits { maxPixels: number; maxDecodedBytes: number }
const MiB = 1024 * 1024
/** Proposed, frozen in S10. */
export function imageLimitsFor(profile: ProfileId): ImageLimits {
  return profile === 'M' ? { maxPixels: 8_388_608, maxDecodedBytes: 32 * MiB } : { maxPixels: 16_777_216, maxDecodedBytes: 64 * MiB }
}

const be16 = (b: Uint8Array, i: number) => (b[i] << 8) | b[i + 1]
const le16 = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8)
const be32 = (b: Uint8Array, i: number) => b[i] * 0x1000000 + ((b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3])
const le32 = (b: Uint8Array, i: number) => b[i + 3] * 0x1000000 + ((b[i + 2] << 16) | (b[i + 1] << 8) | b[i])
const le24 = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8) | (b[i + 2] << 16)
const ascii = (b: Uint8Array, i: number, s: string) => {
  for (let k = 0; k < s.length; k++) if (b[i + k] !== s.charCodeAt(k)) return false
  return true
}

function dims(format: ImageFormat, width: number, height: number): ProbeResult {
  if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width <= 0 || height <= 0) return { error: 'invalid' }
  return { format, width, height }
}

function jpeg(b: Uint8Array): ProbeResult {
  let p = 2
  for (;;) {
    while (p < b.length && b[p] === 0xff && b[p + 1] === 0xff) p++ // fill bytes
    if (p + 4 > b.length) return { error: 'truncated' }
    if (b[p] !== 0xff) return { error: 'invalid' }
    const marker = b[p + 1]
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { p += 2; continue }
    if (marker === 0xd9 || marker === 0xda) return { error: 'invalid' } // EOI / scan before any frame header
    const length = be16(b, p + 2)
    if (length < 2) return { error: 'invalid' }
    const sof = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc
    if (sof) {
      if (p + 9 > b.length) return { error: 'truncated' }
      return dims('jpeg', be16(b, p + 7), be16(b, p + 5))
    }
    p += 2 + length
  }
}

function tiff(b: Uint8Array): ProbeResult {
  const little = b[0] === 0x49
  const u16 = little ? le16 : be16
  const u32 = little ? le32 : be32
  const ifd = u32(b, 4)
  if (ifd + 2 > b.length) return { error: 'truncated' }
  const count = u16(b, ifd)
  let width = -1
  let height = -1
  for (let k = 0; k < count; k++) {
    const e = ifd + 2 + 12 * k
    if (e + 12 > b.length) return { error: 'truncated' }
    const tag = u16(b, e)
    const type = u16(b, e + 2)
    const value = type === 3 ? u16(b, e + 8) : type === 4 ? u32(b, e + 8) : -1
    if (tag === 256) width = value
    else if (tag === 257) height = value
  }
  return dims('tiff', width, height)
}

function webp(b: Uint8Array): ProbeResult {
  if (b.length < 30) return { error: 'truncated' }
  if (ascii(b, 12, 'VP8X')) return dims('webp', le24(b, 24) + 1, le24(b, 27) + 1)
  if (ascii(b, 12, 'VP8L')) {
    if (b[20] !== 0x2f) return { error: 'invalid' }
    const bits = le32(b, 21)
    return dims('webp', (bits & 0x3fff) + 1, ((bits >>> 14) & 0x3fff) + 1)
  }
  if (ascii(b, 12, 'VP8 ')) {
    if (b[23] !== 0x9d || b[24] !== 0x01 || b[25] !== 0x2a) return { error: 'invalid' }
    return dims('webp', le16(b, 26) & 0x3fff, le16(b, 28) & 0x3fff)
  }
  return { error: 'unknown-format' }
}

/** Dimensions from the format header only: PNG IHDR, JPEG SOFn, GIF screen descriptor, BMP info header, WebP VP8/VP8L/VP8X, TIFF IFD. */
export function probeImageHeader(b: Uint8Array): ProbeResult {
  if (b.length >= 8 && b[0] === 0x89 && ascii(b, 1, 'PNG\r\n\x1a\n')) {
    if (b.length < 24) return { error: 'truncated' }
    if (!ascii(b, 12, 'IHDR')) return { error: 'invalid' }
    return dims('png', be32(b, 16), be32(b, 20))
  }
  if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return jpeg(b)
  if (b.length >= 6 && (ascii(b, 0, 'GIF87a') || ascii(b, 0, 'GIF89a'))) {
    if (b.length < 10) return { error: 'truncated' }
    return dims('gif', le16(b, 6), le16(b, 8))
  }
  if (b.length >= 2 && ascii(b, 0, 'BM')) {
    if (b.length < 26) return { error: 'truncated' }
    const headerSize = le32(b, 14)
    if (headerSize === 12) return dims('bmp', le16(b, 18), le16(b, 20))
    const w = le32(b, 18) | 0
    const h = le32(b, 22) | 0
    return dims('bmp', Math.abs(w), Math.abs(h))
  }
  if (b.length >= 12 && ascii(b, 0, 'RIFF') && ascii(b, 8, 'WEBP')) return webp(b)
  if (b.length >= 8 && (ascii(b, 0, 'II*\0') || ascii(b, 0, 'MM\0*'))) return tiff(b)
  if (b.length < 12) return { error: 'truncated' }
  return { error: 'unknown-format' }
}

export type ImageAdmission =
  | { kind: 'decode'; header: ImageHeader; decodedBytes: number; surface: Reservation; releaseTransient(): void; probedBytes: number }
  | { kind: 'placeholder'; header?: ImageHeader; diagnostic: WorkerDiagnostic; probedBytes: number }

/**
 * Probe and admit one image part. At most IMAGE_PROBE_BYTES are inflated
 * (deflated images go through the same bounded output buffer) and the
 * compressed reads are capped to the same size.
 */
export async function admitImagePart(src: RangeSource, entry: ZipEntryInfo, ledger: QuotaLedger, limits: ImageLimits, signal: AbortSignal): Promise<ImageAdmission> {
  const head = new Uint8Array(Math.min(IMAGE_PROBE_BYTES, entry.uncompressedSize))
  let filled = 0
  const part = openPart(src, entry, {
    signal,
    budget: { charge() {} },
    rangeReadBytes: IMAGE_PROBE_BYTES,
    outputBufferBytes: Math.min(65_536, Math.max(1, head.length)),
  })
  try {
    for await (const chunk of part) {
      const n = Math.min(chunk.length, head.length - filled)
      head.set(chunk.subarray(0, n), filled)
      filled += n
      if (filled >= head.length) break // return() disposes the inflater and stops reads
    }
  } finally {
    await part.return()
  }
  throwIfAborted(signal)
  const probe = probeImageHeader(head.subarray(0, filled))
  const placeholder = (code: string, message: string, header?: ImageHeader): ImageAdmission => ({
    kind: 'placeholder', ...(header ? { header } : {}), probedBytes: filled,
    diagnostic: { code, message, partId: entry.name },
  })
  if ('error' in probe) return placeholder(`image-${probe.error}`, `image header ${probe.error === 'unknown-format' ? 'not recognised' : probe.error}; shown as a placeholder`)
  const pixels = probe.width * probe.height
  const decodedBytes = pixels * 4
  if (pixels > limits.maxPixels || decodedBytes > limits.maxDecodedBytes) {
    return placeholder('image-too-large', `${probe.width}x${probe.height} ${probe.format} exceeds ${limits.maxPixels} pixels / ${limits.maxDecodedBytes} decoded bytes; shown as a placeholder`, probe)
  }
  const surface = ledger.reserve('decodedBuffers', decodedBytes)
  let transient: Reservation
  try {
    transient = ledger.reserve('decodedBuffers', decodedBytes)
  } catch (error) {
    surface.release()
    throw error
  }
  return { kind: 'decode', header: probe, decodedBytes, surface, releaseTransient: () => transient.release(), probedBytes: filled }
}
