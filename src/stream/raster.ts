/** UNIT_RASTER unit body (discriminator 2): `u32 w + u32 h + u32 pngLen + PNG`.
 * The decoder admits `w × h × 4` from the PNG IHDR before any pixel decode (§14). */
export class RasterBodyError extends Error { constructor(m: string) { super(m); this.name = 'RasterBodyError' } }
export const UNIT_RASTER_DECODED_CAP = 64 * 1024 * 1024
const SIG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]

export function encodeRasterBody(w: number, h: number, png: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + png.length), v = new DataView(out.buffer)
  v.setUint32(0, w, true); v.setUint32(4, h, true); v.setUint32(8, png.length, true); out.set(png, 12)
  return out
}

export function decodeRasterBody(body: Uint8Array, opts: { maxRgbaBytes?: number } = {}): { w: number; h: number; png: Uint8Array } {
  if (body.length < 12) throw new RasterBodyError('truncated')
  const v = new DataView(body.buffer, body.byteOffset, body.byteLength)
  const w = v.getUint32(0, true), h = v.getUint32(4, true), len = v.getUint32(8, true)
  if (len !== body.length - 12) throw new RasterBodyError('length')
  const png = body.subarray(12)
  if (png.length < 33 || SIG.some((b, i) => png[i] !== b)) throw new RasterBodyError('signature')
  const p = new DataView(png.buffer, png.byteOffset, png.byteLength)
  if (p.getUint32(8) !== 13 || String.fromCharCode(...png.subarray(12, 16)) !== 'IHDR') throw new RasterBodyError('ihdr')
  const iw = p.getUint32(16), ih = p.getUint32(20)
  if (iw !== w || ih !== h || w < 1 || h < 1) throw new RasterBodyError('ihdr-size')
  if (w * h * 4 > (opts.maxRgbaBytes ?? UNIT_RASTER_DECODED_CAP)) throw new RasterBodyError('decoded-cap')
  return { w, h, png }
}

/** PNG chunk types in order (tests assert metadata-free output). */
export function pngChunkTypes(png: Uint8Array): string[] {
  const out: string[] = [], v = new DataView(png.buffer, png.byteOffset, png.byteLength)
  for (let at = 8; at + 8 <= png.length;) { const len = v.getUint32(at); out.push(String.fromCharCode(...png.subarray(at + 4, at + 8))); at += 12 + len }
  return out
}
