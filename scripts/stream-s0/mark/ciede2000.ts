import { checkBitmap, type Chip, type Bitmap } from './types'
type Lab = [number, number, number]
const rad = Math.PI / 180
const cos = (d: number) => Math.cos(d * rad)
const sin = (d: number) => Math.sin(d * rad)
const hue = (a: number, b: number) => (Math.atan2(b, a) / rad + 360) % 360
/** CIEDE2000, Sharma/Wu/Dalal, kL=kC=kH=1; angles in degrees. */
export function deltaE00(lab1: Lab, lab2: Lab): number {
  if (![...lab1, ...lab2].every(Number.isFinite)) throw new Error('Nonfinite Lab')
  const [l1, a1, b1] = lab1, [l2, a2, b2] = lab2
  const c1 = Math.hypot(a1, b1), c2 = Math.hypot(a2, b2), cm = (c1 + c2) / 2
  const g = 0.5 * (1 - Math.sqrt(cm ** 7 / (cm ** 7 + 25 ** 7)))
  const ap1 = (1 + g) * a1, ap2 = (1 + g) * a2
  const cp1 = Math.hypot(ap1, b1), cp2 = Math.hypot(ap2, b2)
  const hp1 = cp1 === 0 ? 0 : hue(ap1, b1), hp2 = cp2 === 0 ? 0 : hue(ap2, b2)
  const dl = l2 - l1, dc = cp2 - cp1
  let dh = hp2 - hp1
  if (cp1 * cp2 === 0) dh = 0
  else if (dh > 180) dh -= 360
  else if (dh < -180) dh += 360
  const dH = 2 * Math.sqrt(cp1 * cp2) * sin(dh / 2)
  const lm = (l1 + l2) / 2, cpm = (cp1 + cp2) / 2
  let hm = hp1 + hp2
  if (cp1 * cp2 !== 0) {
    if (Math.abs(hp1 - hp2) <= 180) hm /= 2
    else hm = (hm + (hm < 360 ? 360 : -360)) / 2
  }
  const t = 1 - 0.17 * cos(hm - 30) + 0.24 * cos(2 * hm) + 0.32 * cos(3 * hm + 6) - 0.20 * cos(4 * hm - 63)
  const sl = 1 + 0.015 * (lm - 50) ** 2 / Math.sqrt(20 + (lm - 50) ** 2)
  const sc = 1 + 0.045 * cpm, sh = 1 + 0.015 * cpm * t
  const rt = -2 * Math.sqrt(cpm ** 7 / (cpm ** 7 + 25 ** 7)) * sin(60 * Math.exp(-(((hm - 275) / 25) ** 2)))
  const l = dl / sl, c = dc / sc, h = dH / sh
  return Math.sqrt(Math.max(0, l * l + c * c + h * h + rt * c * h))
}
/** IEC sRGB transfer, D65 XYZ reference (95.047,100,108.883). */
export function srgbToLab(r: number, g: number, b: number): Lab {
  if (![r, g, b].every(x => Number.isFinite(x) && x >= 0 && x <= 255)) throw new Error('Invalid sRGB')
  const linear = (v: number) => { v /= 255; return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4 }
  const R = linear(r), G = linear(g), B = linear(b)
  const f = (v: number) => v > (6 / 29) ** 3 ? Math.cbrt(v) : v / (3 * (6 / 29) ** 2) + 4 / 29
  const x = f((0.4124564 * R + 0.3575761 * G + 0.1804375 * B) / 0.95047)
  const y = f(0.2126729 * R + 0.7151522 * G + 0.0721750 * B)
  const z = f((0.0193339 * R + 0.1191920 * G + 0.9503041 * B) / 1.08883)
  return [116 * y - 16, 500 * (x - y), 200 * (y - z)]
}
/** Exact nearest-rank p95 over all pixels, composited onto white for alpha.
 * Stage-0 measurement: O(pixels) storage, O(pixels log pixels) sort. */
export function pageDeltaE(a: Bitmap, b: Bitmap, chips: readonly Chip[] = (b as Bitmap & { chipSupport?: Chip[] }).chipSupport ?? []): { mean: number; p95: number; max: number; chipPixelsP95: number } {
  checkBitmap(a); checkBitmap(b)
  if (a.width !== b.width || a.height !== b.height) throw new Error('Bitmap dimensions differ')
  const values = new Float64Array(a.width * a.height)
  const lab = (data: Uint8ClampedArray, i: number): Lab => {
    const alpha = data[i + 3]! / 255
    return srgbToLab(data[i]! * alpha + 255 * (1 - alpha), data[i + 1]! * alpha + 255 * (1 - alpha), data[i + 2]! * alpha + 255 * (1 - alpha))
  }
  let sum = 0, max = 0
  for (let i = 0; i < values.length; i++) { const at = 4 * i; const same = a.data[at] === b.data[at] && a.data[at + 1] === b.data[at + 1] && a.data[at + 2] === b.data[at + 2] && a.data[at + 3] === b.data[at + 3]; const d = same ? 0 : deltaE00(lab(a.data, at), lab(b.data, at)); values[i] = d; sum += d; max = Math.max(max, d) }
  const support = new Set<number>()
  for (const chip of chips) for (let y = chip.y; y < chip.y + chip.cell; y++) for (let x = chip.x; x < chip.x + chip.cell; x++) { if (x < 0 || y < 0 || x >= a.width || y >= a.height) throw new Error("Invalid chip support"); support.add(y * a.width + x) }
  const selected = [...support].map(i => values[i]!).sort((a, b) => a - b)
  const chipPixelsP95 = selected.length ? selected[Math.ceil(selected.length * 0.95) - 1]! : 0
  values.sort()
  return { mean: sum / values.length, p95: values[Math.ceil(values.length * 0.95) - 1]!, max, chipPixelsP95 }
}
