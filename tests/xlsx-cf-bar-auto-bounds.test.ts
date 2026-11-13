// @vitest-environment node
/**
 * Task 22c (owner ruling S3-AUTO-BOUNDS): stage3-bar-auto-bounds, Excel's preset bounds (x14 autoMin/autoMax).
 *
 * The case has no capture yet, so its hypothesis families are run end to end through the real harness on a synthetic
 * capture: the generated workbook, Excel's measured geometry convention (180 pt columns, 30 pt rows, 1 px per pt) and
 * bars planted under one automatic-bound rule by an independent painter (solid colours, the axis column, the verified
 * gradient and bordered-gradient pixels). The planted rule must come back as the one match. This checks the families,
 * not Excel.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { CfManifest } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import {
  computeWorksheetRangeGeometry,
  evaluateBarCapabilities,
  measureCaseBars,
  predictBarGradientBorderPixel,
  predictBarGradientPixel,
  readBarSheetFacts,
  readZipPartsSync,
  type DecodedPng,
} from '../scripts/cf-refs/verify'

const REPO_ROOT = resolve(__dirname, '..')
const TIMEOUT_MS = 60_000
const c = (JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest).cases.find(x => x.id === 'stage3-bar-auto-bounds')!
const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
const facts = readBarSheetFacts(parts)
const rows = Number(/\d+$/.exec(c.range)![0])
const colWidthsPt = [84, ...Array.from({ length: 10 }, () => 180)]
const geom = computeWorksheetRangeGeometry(parts, c.range, { colWidthsPt, rowHeightsPt: Array.from({ length: rows }, () => 30) })!
const half = (x: number): number => Math.floor(x + 0.5 + 1e-9)
const BLUE = [0, 0, 255], RED = [255, 0, 0], BLACK = [0, 0, 0]
// B..G all-positive (F, G at 10/90; D, E, G gradient; C, D, G bordered), H, I all-negative, J, K mixed-sign.
const kind = (col: number) => ({ gradient: [3, 4, 6].includes(col), border: [2, 3, 6, 8, 10].includes(col), defaultLengths: [5, 6].includes(col) })
const scopeOf = (col: number): number[] => [2, 3, 4, 5, 6].map(r => (facts.values.get(`${String.fromCharCode(65 + col)}${r}`) as { v: number }).v)

function synthetic(rule: 'zero-anchored' | 'min-max' | 'outside'): DecodedPng {
  const width = colWidthsPt.reduce((a, b) => a + b, 0) + 1, height = rows * 30 + 1
  const data = new Uint8ClampedArray(width * height * 4).fill(255)
  const put = (x: number, y: number, rgb: readonly number[]) => { const j = (y * width + x) * 4; data[j] = rgb[0]!; data[j + 1] = rgb[1]!; data[j + 2] = rgb[2]! }
  for (const cell of c.expected.cells.filter(x => x.matched)) {
    const col = cell.ref.charCodeAt(0) - 65, row = Number(cell.ref.slice(1)) - 1
    const v = (facts.values.get(cell.ref) as { v: number }).v
    const scope = scopeOf(col), vmin = Math.min(...scope), vmax = Math.max(...scope)
    const lo = rule === 'zero-anchored' ? Math.min(0, vmin) : vmin
    const hi = rule === 'zero-anchored' ? Math.max(0, vmax) : rule === 'outside' && vmax > 0 ? vmax + 50 : vmax
    const x0 = colWidthsPt.slice(0, col).reduce((a, b) => a + b, 0) + 2, S = 177, top = row * 30 + 2, bottom = row * 30 + 27
    const k = kind(col)
    let start = 0, len = 0, rgb = BLUE
    if (lo < 0 && hi > 0) {
      // Automatic axis at the data's zero, each side its own span (the verified axis/side layout).
      const sn = half((S * -lo) / (hi - lo))
      for (let y = row * 30; y < row * 30 + 30; y++) put(x0 + sn - 1, y, BLACK)
      if (v >= 0) { len = half((v / hi) * (S - sn)); start = sn } else { len = half((v / lo) * (sn - 1)); start = sn - 1 - len; rgb = RED }
    } else if (hi <= 0) {
      // All-negative: from the right edge, longest at the minimum.
      len = lo === hi ? 0 : half(((v - hi) / (lo - hi)) * S); start = S - len; rgb = RED
    } else {
      const t = (v - lo) / (hi - lo)
      len = half((k.defaultLengths ? 0.05 + 0.8 * t : t) * S)
    }
    const sample = { fill: rgb as [number, number, number], border: rgb as [number, number, number], length: len, span: S, leftPx: x0 + start, topPx: top, bottomPx: bottom, ref: cell.ref, pixels: [] }
    for (let y = top; y <= bottom; y++) {
      for (let i = 0; i < len; i++) {
        const px = !k.gradient ? rgb
          : k.border ? predictBarGradientBorderPixel('frame-over-ramp', sample, i, y, 0.847).map(Math.round)
            : predictBarGradientPixel('bar-length', sample, i, 0.847, []).map(Math.round)
        put(x0 + start + i, y, px)
      }
    }
  }
  return { width, height, data }
}

const judge = (rule: 'zero-anchored' | 'min-max' | 'outside') => {
  const png = synthetic(rule)
  const measured = measureCaseBars(c, png, geom, parts)
  expect(measured instanceof Map, JSON.stringify(measured).slice(0, 300)).toBe(true)
  const ev = evaluateBarCapabilities(c, parts, png, geom, measured as never)
  return ev
}

describe('stage3-bar-auto-bounds (synthetic capture, Task 22c)', () => {
  it('designates 50 cells of ten paired x14 bars with autoMin/autoMax cfvo', () => {
    expect(c.status).toBe('recipe-pending')
    expect(c.expected.cells.filter(x => x.matched)).toHaveLength(50)
    expect(c.markers.filter(m => m.extension === 'x14' && m.bar?.cfvo.map(v => v.type).join() === 'autoMin,autoMax')).toHaveLength(10)
  })

  for (const rule of ['zero-anchored', 'min-max'] as const) {
    it(`planted ${rule}: both automatic bounds come back as the one match`, () => {
      const ev = judge(rule)
      for (const cap of ['bar:bounds-automin', 'bar:bounds-automax']) expect([ev[cap]!.status, ev[cap]!.matched], `${cap}: ${ev[cap]!.reason}`).toEqual(['discriminating', rule])
    }, TIMEOUT_MS)
  }

  it('bounds outside both rules match none (refused, never fitted)', () => {
    const ev = judge('outside')
    expect(ev['bar:bounds-automin']!.status).toBe('none')
  }, TIMEOUT_MS)

  it('the two rules are far apart where they differ: at least 28 px, or a bar against no bar', () => {
    const ev = judge('zero-anchored')
    const sep = (cap: string, ref: string) => {
      const [a, b] = ev[cap]!.hypotheses![ref]!
      return a!.noBar || b!.noBar ? (a!.noBar && b!.noBar ? 0 : Infinity) : Math.max(Math.abs(a!.leftPx - b!.leftPx), Math.abs(a!.rightPx - b!.rightPx))
    }
    // autoMin = 0 vs v_min = 20 on {20..100}: the 20 draws 35 px at 0/100 or nothing; at 10/90, 37 px against 9 px.
    expect(sep('bar:bounds-automin', 'B2')).toBe(Infinity)
    expect(sep('bar:bounds-automin', 'F2')).toBe(28)
    // autoMax = 0 vs v_max = -20 on {-100..-20}: the -20 draws 35 px from the right edge or nothing.
    expect(sep('bar:bounds-automax', 'H6')).toBe(Infinity)
    // On mixed-sign data both rules agree (the axis cross-check).
    expect(sep('bar:bounds-automax', 'J2')).toBe(0)
  }, TIMEOUT_MS)
})
