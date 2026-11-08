import { scale as resize } from './attacks'
import type { Bitmap } from './types'
export const lum = (b: Bitmap, x: number, y: number): number => {
  const i = (y * b.width + x) * 4, a = b.data[i + 3]! / 255
  return a * (0.2126 * b.data[i]! + 0.7152 * b.data[i + 1]! + 0.0722 * b.data[i + 2]!) + (1 - a) * 255
}
function fft(r: Float64Array, im: Float64Array, inverse: boolean): void {
  const n = r.length
  for (let i = 1, j = 0; i < n; i++) {
    let bit = n >> 1; for (; j & bit; bit >>= 1) j ^= bit; j ^= bit
    if (i < j) { [r[i], r[j]] = [r[j]!, r[i]!]; [im[i], im[j]] = [im[j]!, im[i]!] }
  }
  for (let len = 2; len <= n; len *= 2) {
    const angle = (inverse ? 2 : -2) * Math.PI / len, wr = Math.cos(angle), wi = Math.sin(angle)
    for (let i = 0; i < n; i += len) {
      let ar = 1, ai = 0
      for (let j = 0; j < len / 2; j++) {
        const u = i + j, v = u + len / 2, vr = r[v]! * ar - im[v]! * ai, vi = r[v]! * ai + im[v]! * ar
        r[v] = r[u]! - vr; im[v] = im[u]! - vi; r[u] += vr; im[u] += vi
        const next = ar * wr - ai * wi; ai = ar * wi + ai * wr; ar = next
      }
    }
  }
  if (inverse) for (let i = 0; i < n; i++) { r[i] /= n; im[i] /= n }
}
function fft2(r: Float64Array, im: Float64Array, n: number, inverse: boolean): void {
  for (let y = 0; y < n; y++) fft(r.subarray(y * n, (y + 1) * n), im.subarray(y * n, (y + 1) * n), inverse)
  const a = new Float64Array(n), b = new Float64Array(n)
  for (let x = 0; x < n; x++) {
    for (let y = 0; y < n; y++) { a[y] = r[y * n + x]!; b[y] = im[y * n + x]! }
    fft(a, b, inverse)
    for (let y = 0; y < n; y++) { r[y * n + x] = a[y]!; im[y * n + x] = b[y]! }
  }
}
/** Coarse phase-correlation peaks, zero padded to prevent cyclic wrap. */
function peaks(ref: Bitmap, cap: Bitmap): { x: number; y: number; stride: number }[] {
  const stride = Math.max(1, Math.ceil(Math.max(ref.width, ref.height) / 128))
  let n = 1; while (n < 2 * Math.ceil(Math.max(ref.width, ref.height) / stride)) n *= 2
  const rr = new Float64Array(n * n), ri = new Float64Array(n * n), cr = new Float64Array(n * n), ci = new Float64Array(n * n)
  const fill = (b: Bitmap, a: Float64Array) => {
    let sum = 0, count = 0
    for (let y = 0; y < b.height; y += stride) for (let x = 0; x < b.width; x += stride) { sum += lum(b, x, y); count++ }
    for (let y = 0; y < b.height; y += stride) for (let x = 0; x < b.width; x += stride) a[(y / stride) * n + x / stride] = lum(b, x, y) - sum / count
  }
  fill(ref, rr); fill(cap, cr); fft2(rr, ri, n, false); fft2(cr, ci, n, false)
  for (let i = 0; i < rr.length; i++) {
    const re = rr[i]! * cr[i]! + ri[i]! * ci[i]!, im = ri[i]! * cr[i]! - rr[i]! * ci[i]!, norm = Math.hypot(re, im)
    rr[i] = norm > 1e-12 ? re / norm : 0; ri[i] = norm > 1e-12 ? im / norm : 0
  }
  fft2(rr, ri, n, true)
  const candidates: { x: number; y: number; value: number }[] = []
  for (let y = 0; y <= Math.ceil((ref.height - cap.height) / stride); y++) for (let x = 0; x <= Math.ceil((ref.width - cap.width) / stride); x++) candidates.push({ x: x * stride, y: y * stride, value: rr[y * n + x]! })
  return candidates.sort((a, b) => b.value - a.value).slice(0, 12).map(c => ({ ...c, stride }))
}
export interface Registration {
  factor: number
  /** Crop origin in ORIGINAL reference pixels; scaled-first crops may be fractional. */
  x: number; y: number; scaleX: number; scaleY: number
  /** Already aligned/resampled to capture dimensions, never sampled out of bounds. */
  reference: Bitmap; error: number
}
function rectangle(b: Bitmap, x: number, y: number, width: number, height: number): Bitmap {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let row = 0; row < height; row++) data.set(b.data.subarray(((y + row) * b.width + x) * 4, ((y + row) * b.width + x + width) * 4), row * width * 4)
  return { width, height, data }
}
function sample(b: Bitmap, x: number, y: number): number {
  x = Math.max(0, Math.min(b.width - 1, x)); y = Math.max(0, Math.min(b.height - 1, y))
  const x0 = Math.floor(x), y0 = Math.floor(y), tx = x - x0, ty = y - y0
  return (1 - ty) * ((1 - tx) * lum(b, x0, y0) + tx * lum(b, Math.min(x0 + 1, b.width - 1), y0)) + ty * ((1 - tx) * lum(b, x0, Math.min(y0 + 1, b.height - 1)) + tx * lum(b, Math.min(x0 + 1, b.width - 1), Math.min(y0 + 1, b.height - 1)))
}
export function register(capture: Bitmap, reference: Bitmap): Registration | undefined {
  let best: Registration | undefined
  const step = Math.max(1, Math.floor(Math.sqrt(capture.width * capture.height / 4096)))
  const errorOf = (read: (x: number, y: number) => number, stride = step): number => {
    let sum = 0, count = 0
    for (let y = 0; y < capture.height; y += stride) for (let x = 0; x < capture.width; x += stride) { sum += (lum(capture, x, y) - read(x, y)) ** 2; count++ }
    return sum / count
  }
  for (const factor of [1, 0.5, 2]) {
    const w = Math.round(reference.width * factor), h = Math.round(reference.height * factor), slack = Math.ceil(factor)
    if (capture.width > w || capture.height > h || capture.width < Math.floor(w * 0.75) - slack || capture.height < Math.floor(h * 0.75) - slack) continue
    const ref = factor === 1 ? reference : resize(reference, factor)
    const seeds = w === capture.width && h === capture.height ? [{ x: 0, y: 0, stride: 1 }] : peaks(ref, capture)
    // Preserve scale-then-crop support, where origins are scaled pixel integers.
    let scaledBest: { x: number; y: number; error: number } | undefined
    const visited = new Set<number>()
    for (const seed of seeds) for (let y = Math.max(0, seed.y - seed.stride); y <= Math.min(h - capture.height, seed.y + seed.stride); y++) for (let x = Math.max(0, seed.x - seed.stride); x <= Math.min(w - capture.width, seed.x + seed.stride); x++) {
      const id = y * w + x; if (visited.has(id)) continue; visited.add(id)
      const error = errorOf((cx, cy) => lum(ref, cx + x, cy + y))
      if (!scaledBest || error < scaledBest.error) scaledBest = { x, y, error }
    }
    if (scaledBest && (!best || scaledBest.error < best.error)) {
      const { x, y, error } = scaledBest
      best = { factor, x: x * reference.width / w, y: y * reference.height / h, scaleX: w / reference.width, scaleY: h / reference.height, reference: rectangle(ref, x, y, capture.width, capture.height), error }
    }
    if (factor === 1) continue // both orders coincide
    // Invert round(sourceExtent*factor): half-scale has two possible integer
    // extents. Search integer ORIGINAL offsets, including half-pixel phases
    // after scaling. Coarse interpolation ranks candidates; actual napi
    // crop-then-resize determines the final content error and reference.
    const extents = (size: number, max: number) => {
      const result: number[] = []
      for (let n = Math.ceil((size - 0.5) / factor); n < (size + 0.5) / factor; n++) if (n > 0 && n <= max) result.push(n)
      return result
    }
    for (const cw of extents(capture.width, reference.width)) for (const ch of extents(capture.height, reference.height)) {
      const sx = capture.width / cw, sy = capture.height / ch
      const shortlist: { x: number; y: number; error: number }[] = [], seen = new Set<number>()
      for (const seed of seeds) {
        const radius = Math.ceil(seed.stride / factor) + 1, ox = Math.round(seed.x / factor), oy = Math.round(seed.y / factor)
        for (let y = Math.max(0, oy - radius); y <= Math.min(reference.height - ch, oy + radius); y++) for (let x = Math.max(0, ox - radius); x <= Math.min(reference.width - cw, ox + radius); x++) {
          const id = y * reference.width + x; if (seen.has(id)) continue; seen.add(id)
          const error = errorOf((cx, cy) => sample(reference, x + (cx + 0.5) / sx - 0.5, y + (cy + 0.5) / sy - 0.5), Math.max(step, Math.floor(Math.sqrt(capture.width * capture.height / 256))))
          if (shortlist.length < 8 || error < shortlist[shortlist.length - 1]!.error) { shortlist.push({ x, y, error }); shortlist.sort((a, b) => a.error - b.error); if (shortlist.length > 8) shortlist.pop() }
        }
      }
      for (const { x, y } of shortlist) {
        const aligned = resize(rectangle(reference, x, y, cw, ch), factor)
        const error = errorOf((cx, cy) => lum(aligned, cx, cy))
        if (!best || error < best.error) best = { factor, x, y, scaleX: sx, scaleY: sy, reference: aligned, error }
      }
    }
  }
  return best
}
