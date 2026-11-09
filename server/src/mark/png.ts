/**
 * Minimal lossless PNG codec for the Layer-1 encode job (§9.4): 8-bit,
 * non-interlaced, colour types 2 (RGB) and 6 (RGBA, straight alpha).
 * Decoding then re-encoding is bit-exact for every pixel the embed leaves alone.
 */
import { deflateSync, inflateSync, crc32 } from 'node:zlib'

export class PngUnsupported extends Error { constructor(why: string) { super(`unsupported PNG: ${why}`); this.name = 'PngUnsupported' } }

const SIG = [137, 80, 78, 71, 13, 10, 26, 10]

export function decodePng(png: Uint8Array): { w: number; h: number; rgba: Uint8ClampedArray } {
  if (png.length < 8 + 25 || SIG.some((b, i) => png[i] !== b)) throw new PngUnsupported('signature')
  const dv = new DataView(png.buffer, png.byteOffset, png.byteLength)
  let at = 8, w = 0, h = 0, type = -1
  const idat: Uint8Array[] = []
  while (at + 12 <= png.length) {
    const len = dv.getUint32(at), name = String.fromCharCode(...png.subarray(at + 4, at + 8)), body = png.subarray(at + 8, at + 8 + len)
    if (at + 12 + len > png.length) throw new PngUnsupported('truncated chunk')
    if (name === 'IHDR') {
      if (len !== 13) throw new PngUnsupported('IHDR')
      w = dv.getUint32(at + 8); h = dv.getUint32(at + 12)
      const depth = body[8]!, interlace = body[12]!
      type = body[9]!
      if (depth !== 8 || (type !== 2 && type !== 6) || body[10] !== 0 || body[11] !== 0 || interlace !== 0) throw new PngUnsupported(`depth ${depth} type ${type} interlace ${interlace}`)
    } else if (name === 'IDAT') idat.push(body)
    else if (name === 'IEND') break
    at += 12 + len
  }
  if (type < 0 || w <= 0 || h <= 0 || idat.length === 0) throw new PngUnsupported('missing IHDR/IDAT')
  const bpp = type === 6 ? 4 : 3, stride = w * bpp
  let raw: Uint8Array
  try { raw = inflateSync(Buffer.concat(idat)) } catch { throw new PngUnsupported('zlib') }
  if (raw.length < h * (stride + 1)) throw new PngUnsupported('short image data')
  const rows = new Uint8Array(h * stride)
  for (let y = 0; y < h; y++) {
    const f = raw[y * (stride + 1)]!, src = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1))
    const cur = rows.subarray(y * stride, (y + 1) * stride), prev = y ? rows.subarray((y - 1) * stride, y * stride) : undefined
    for (let i = 0; i < stride; i++) {
      const a = i >= bpp ? cur[i - bpp]! : 0, b = prev ? prev[i]! : 0, c = prev && i >= bpp ? prev[i - bpp]! : 0
      let p: number
      switch (f) {
        case 0: p = 0; break
        case 1: p = a; break
        case 2: p = b; break
        case 3: p = (a + b) >> 1; break
        case 4: p = paeth(a, b, c); break
        default: throw new PngUnsupported(`filter ${f}`)
      }
      cur[i] = (src[i]! + p) & 255
    }
  }
  if (bpp === 4) return { w, h, rgba: new Uint8ClampedArray(rows.buffer) }
  const rgba = new Uint8ClampedArray(w * h * 4)
  for (let p = 0; p < w * h; p++) { rgba[p * 4] = rows[p * 3]!; rgba[p * 4 + 1] = rows[p * 3 + 1]!; rgba[p * 4 + 2] = rows[p * 3 + 2]!; rgba[p * 4 + 3] = 255 }
  return { w, h, rgba }
}

function paeth(a: number, b: number, c: number) {
  const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c)
  return pa <= pb && pa <= pc ? a : pb <= pc ? b : c
}

function chunk(name: string, body: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + body.length), dv = new DataView(out.buffer)
  dv.setUint32(0, body.length)
  for (let i = 0; i < 4; i++) out[4 + i] = name.charCodeAt(i)
  out.set(body, 8)
  dv.setUint32(8 + body.length, crc32(out.subarray(4, 8 + body.length)))
  return out
}

/** RGBA colour type 6; per-row filter by minimum sum of absolute (signed) residuals; zlib level 6. */
export function encodePng(w: number, h: number, rgba: Uint8ClampedArray | Uint8Array): Uint8Array {
  if (!Number.isSafeInteger(w) || !Number.isSafeInteger(h) || w <= 0 || h <= 0 || rgba.length !== w * h * 4) throw new Error('Invalid RGBA for encodePng')
  const stride = w * 4, raw = new Uint8Array(h * (stride + 1)), trial = new Uint8Array(stride)
  for (let y = 0; y < h; y++) {
    const cur = rgba.subarray(y * stride, (y + 1) * stride), prev = y ? rgba.subarray((y - 1) * stride, y * stride) : undefined
    let best = Infinity, bestF = 0
    const out = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1))
    for (let f = 0; f < 5; f++) {
      let sum = 0
      for (let i = 0; i < stride; i++) {
        const a = i >= 4 ? cur[i - 4]! : 0, b = prev ? prev[i]! : 0, c = prev && i >= 4 ? prev[i - 4]! : 0
        const p = f === 0 ? 0 : f === 1 ? a : f === 2 ? b : f === 3 ? (a + b) >> 1 : paeth(a, b, c)
        const v = (cur[i]! - p) & 255
        trial[i] = v; sum += v < 128 ? v : 256 - v
      }
      if (sum < best) { best = sum; bestF = f; out.set(trial) }
    }
    raw[y * (stride + 1)] = bestF
  }
  const ihdr = new Uint8Array(13), dv = new DataView(ihdr.buffer)
  dv.setUint32(0, w); dv.setUint32(4, h); ihdr[8] = 8; ihdr[9] = 6
  const parts = [Uint8Array.from(SIG), chunk('IHDR', ihdr), chunk('IDAT', new Uint8Array(deflateSync(raw, { level: 6 }))), chunk('IEND', new Uint8Array())]
  const png = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0
  for (const p of parts) { png.set(p, o); o += p.length }
  return png
}
