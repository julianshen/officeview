import { createCanvas, loadImage } from '../napi-shim'
import { checkBitmap, type Bitmap } from './types'
function canvas(b: Bitmap) {
  checkBitmap(b)
  const c = createCanvas(b.width, b.height), ctx = c.getContext('2d'), pixels = ctx.createImageData(b.width, b.height)
  pixels.data.set(b.data); ctx.putImageData(pixels, 0, 0)
  return c
}
function read(c: ReturnType<typeof createCanvas>): Bitmap {
  return { width: c.width, height: c.height, data: new Uint8ClampedArray(c.getContext('2d').getImageData(0, 0, c.width, c.height).data) }
}
export async function jpeg(b: Bitmap, q: number): Promise<Bitmap> {
  if (!Number.isInteger(q) || q < 1 || q > 100) throw new Error('JPEG quality must be 1..100')
  const c = canvas(b), bytes = await c.encode('jpeg', q), image = await loadImage(bytes)
  const decoded = createCanvas(b.width, b.height); decoded.getContext('2d').drawImage(image, 0, 0)
  return read(decoded)
}
export function scale(b: Bitmap, factor: number): Bitmap {
  if (!Number.isFinite(factor) || factor <= 0) throw new Error('Invalid scale')
  const src = canvas(b), w = Math.round(b.width * factor), h = Math.round(b.height * factor)
  if (!Number.isSafeInteger(w) || !Number.isSafeInteger(h) || w < 1 || h < 1) throw new Error('Invalid scaled size')
  const c = createCanvas(w, h), ctx = c.getContext('2d')
  ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high'; ctx.drawImage(src, 0, 0, w, h)
  return read(c)
}
/** frac removes this fraction of each axis (not area); anchor retains the
 * named corner/centre. Dimensions floor once; no extrapolated pixels. */
export function crop(b: Bitmap, frac: number, anchor: 'top-left' | 'center' | 'bottom-right' = 'center'): Bitmap {
  checkBitmap(b)
  if (!Number.isFinite(frac) || frac < 0 || frac >= 1 || !['top-left', 'center', 'bottom-right'].includes(anchor)) throw new Error('Invalid crop')
  const width = Math.max(1, Math.floor(b.width * (1 - frac))), height = Math.max(1, Math.floor(b.height * (1 - frac)))
  const f = anchor === 'top-left' ? 0 : anchor === 'center' ? 0.5 : 1
  const x = Math.floor((b.width - width) * f), y = Math.floor((b.height - height) * f), data = new Uint8ClampedArray(width * height * 4)
  for (let row = 0; row < height; row++) data.set(b.data.subarray(((row + y) * b.width + x) * 4, ((row + y) * b.width + x + width) * 4), row * width * 4)
  return { width, height, data }
}
type Attack = (b: Bitmap) => Bitmap | Promise<Bitmap>
export const compose = (...attacks: Attack[]): Attack => async b => { for (const attack of attacks) b = await attack(b); return b }
function combine(bmps: Bitmap[], fn: (values: number[]) => number): Bitmap {
  if (!bmps.length) throw new Error('Empty image set')
  bmps.forEach(checkBitmap)
  const { width, height } = bmps[0]!
  if (bmps.some(b => b.width !== width || b.height !== height)) throw new Error('Dimensions differ')
  const data = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < data.length; i++) data[i] = fn(bmps.map(b => b.data[i]!))
  return { width, height, data }
}
export const average = (bmps: Bitmap[]): Bitmap => combine(bmps, v => v.reduce((a, b) => a + b, 0) / v.length)
export const median = (bmps: Bitmap[]): Bitmap => combine(bmps, v => { v.sort((a, b) => a - b); return (v[Math.floor((v.length - 1) / 2)]! + v[Math.floor(v.length / 2)]!) / 2 })
/** Signed difference encoded about 128, clamped to bytes; alpha opaque. */
export function difference(a: Bitmap, b: Bitmap): Bitmap {
  const result = combine([a, b], v => 128 + v[0]! - v[1]!)
  for (let i = 3; i < result.data.length; i += 4) result.data[i] = 255
  return result
}
