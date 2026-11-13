// @vitest-environment node
/**
 * Task 22b: Excel's gradient bar fill (`bar:render-gradient-stops`, Ruling S3-M), judged from the Stage 3 captures.
 *
 * The model: linear in sRGB over the bar's own length, from the fill toward white by a blend fraction t fitted to the
 * pixels (f + (255 - f)·t·(i + ½)/L). Every alternative (the ramp over the whole span, a white end, linear light, a
 * solid fill, one fixed end colour) is measurably different on the same captures, and a mutated capture shows it.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { auditStage3Manifest, barGradientsAcrossCases, renderBarGradientBlock } from '../scripts/cf-refs/audit'
import type { CfManifest, CfOracleCase } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import {
  BAR_COLOR_TOLERANCE,
  BAR_LENGTH_RULES,
  computeWorksheetRangeGeometry,
  evaluateBarCapabilities,
  judgeBarGradient,
  measureCaseBars,
  readZipPartsSync,
  loadBarScreenCapture,
  predictBarGradientPixel,
  sampleBarGradients,
  type BarGradientObservation,
  type DecodedPng,
} from '../scripts/cf-refs/verify'

const REPO_ROOT = resolve(__dirname, '..')
const manifest = (): CfManifest => JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
const caseOf = (id: string): CfOracleCase => manifest().cases.find(c => c.id === id)!
const capture = (id: string) => loadBarScreenCapture(caseOf(id), REPO_ROOT)!
const C_REFS = new Set(['C2', 'C3', 'C4', 'C5', 'C6'])

/** A copy of the capture with every sampled bar of `o` repainted column by column (`paint` returns RGB). */
function repaint(png: DecodedPng, o: BarGradientObservation, paint: Paint): DecodedPng {
  const data = new Uint8ClampedArray(png.data)
  for (const s of o.samples) {
    for (let y = s.topPx; y <= s.bottomPx; y++) {
      for (let i = 0; i < s.length; i++) {
        const j = (y * png.width + s.leftPx + i) * 4
        const rgb = paint(i, s.length, s.span, s.fill)
        data[j] = rgb[0]!; data[j + 1] = rgb[1]!; data[j + 2] = rgb[2]!
      }
    }
  }
  return { width: png.width, height: png.height, data }
}
type Paint = (i: number, length: number, span: number, fill: [number, number, number]) => number[]

describe('gradient sampling (Task 22b)', () => {
  it('samples rendering C3:C6 at the solid-length model extents; C2 draws no bar', () => {
    const { parts, png, geom } = capture('stage3-bar-rendering')
    const o = sampleBarGradients(caseOf('stage3-bar-rendering'), parts, png, geom, C_REFS)
    expect(o.samples.map(s => [s.ref, s.length, s.predictedLength, s.span])).toEqual([
      ['C3', 18, 18, 177], ['C4', 89, 89, 177], ['C5', 159, 159, 177], ['C6', 177, 177, 177],
    ])
    expect(o.unsampled).toEqual(['C2: no bar is drawn'])
    expect(o.samples.every(s => s.leftPx === 266 && s.fill.join() === '0,0,255')).toBe(true)
    // Non-vacuous: every bar row of every sampled bar contributes pixels (text ink is masked, never sampled).
    expect(o.samples.reduce((n, s) => n + s.pixels.length, 0)).toBeGreaterThan(11_000)
  })

  it('samples only the measured shape: the solid, bordered, right-to-left and axis bars of the captures are not gradients', () => {
    for (const id of ['stage3-bar-lengths', 'stage3-bar-axis-negative', 'stage3-bar-geometry', 'stage3-bar-auto-colors']) {
      const { parts, png, geom } = capture(id)
      const o = sampleBarGradients(caseOf(id), parts, png, geom)
      expect(o.samples, id).toEqual([])
      expect(o.unsampled, id).toEqual([])
    }
  })
})

describe('gradient model judgement (Task 22b)', () => {
  it('the owning capture discriminates: bar-length matches within the tolerance, every alternative misses', () => {
    const { parts, png, geom } = capture('stage3-bar-rendering')
    const j = judgeBarGradient(sampleBarGradients(caseOf('stage3-bar-rendering'), parts, png, geom, C_REFS))
    expect(j.status).toBe('discriminating')
    expect(j.matched).toBe('bar-length')
    expect(j.endBlend).toBe(0.847)
    expect(j.maxError!['bar-length']).toBeLessThanOrEqual(BAR_COLOR_TOLERANCE)
    for (const h of ['cell-span', 'white-end', 'linear-light', 'solid'] as const) expect(j.maxError![h], h).toBeGreaterThan(2 * BAR_COLOR_TOLERANCE)
    // One fill cannot tell a fixed end colour from a blend toward white, so that alternative is not listed here.
    expect(j.maxError!['fixed-end']).toBeUndefined()
  })

  it('pooled over every capture (two fills, lengths 18..177): one blend fraction, the fixed end refuted', () => {
    const pooled = barGradientsAcrossCases(manifest(), REPO_ROOT)
    expect(new Set(pooled.samples.map(s => s.ref.split(':')[0]))).toEqual(new Set(['stage3-bar-settings-probe', 'stage3-bar-degenerate', 'stage3-bar-rendering', 'stage3-bar-pairing']))
    expect(pooled.samples).toHaveLength(24)
    expect(new Set(pooled.samples.map(s => s.fill.join()))).toEqual(new Set(['0,0,255', '99,142,198']))
    const j = judgeBarGradient(pooled)
    expect(j.status).toBe('discriminating')
    expect(j.matched).toBe('bar-length')
    expect(j.endBlend).toBe(0.847)
    expect(j.maxError!['bar-length']).toBeLessThanOrEqual(6)
    expect(j.maxError!['fixed-end']).toBeGreaterThan(2 * BAR_COLOR_TOLERANCE)
  }, 30_000)

  it('mutants: a solid repaint, a span-wide ramp and a white end never verify bar-length', () => {
    const c = caseOf('stage3-bar-rendering')
    const { parts, png, geom } = capture('stage3-bar-rendering')
    const o = sampleBarGradients(c, parts, png, geom, C_REFS)
    const judgeRepainted = (paint: Paint) =>
      judgeBarGradient(sampleBarGradients(c, parts, repaint(png, o, paint), geom, C_REFS))
    const solid = judgeRepainted((_i, _l, _s, f) => f)
    expect(solid.matched).not.toBe('bar-length')
    const span = judgeRepainted((i, l, s, f) => predictBarGradientPixel('cell-span', { fill: f, length: l, span: s }, i, 0.847, []).map(Math.round))
    // The blend is fitted under bar-length, so a span-wide ramp fits no hypothesis: refused, never verified.
    expect(span.status).toBe('none')
    expect(span.matched).toBeUndefined()
    const white = judgeRepainted((i, l, s, f) => predictBarGradientPixel('white-end', { fill: f, length: l, span: s }, i, 0.847, []).map(Math.round))
    expect(white.matched).not.toBe('bar-length')
  })

  it('a gradient extent off the solid-length model by more than 1 px refuses (Ruling S3-M)', () => {
    const { parts, png, geom } = capture('stage3-bar-rendering')
    const o = sampleBarGradients(caseOf('stage3-bar-rendering'), parts, png, geom, C_REFS)
    const shifted: BarGradientObservation = { ...o, samples: o.samples.map(s => (s.ref === 'C4' ? { ...s, predictedLength: s.length + 2 } : s)) }
    expect(judgeBarGradient(shifted).status).toBe('none')
    const within: BarGradientObservation = { ...o, samples: o.samples.map(s => (s.ref === 'C4' ? { ...s, predictedLength: s.length + 1 } : s)) }
    expect(judgeBarGradient(within).status).toBe('discriminating')
  })

  it('a full-span bar alone cannot tell the bar-length ramp from the span-wide one (Lesson 3)', () => {
    const { parts, png, geom } = capture('stage3-bar-rendering')
    const j = judgeBarGradient(sampleBarGradients(caseOf('stage3-bar-rendering'), parts, png, geom, new Set(['C6'])))
    expect(j.status).toBe('non-discriminating')
    expect(j.reason).toMatch(/bar-length\/cell-span/)
    expect(judgeBarGradient({ samples: [], unsampled: [] }).status).toBe('no-hypotheses')
  })
})

describe('gradient export (Task 22b)', () => {
  it('the audit verifies bar:render-gradient-stops and exports the pooled blend; the block renders idempotently', () => {
    const report = auditStage3Manifest(manifest(), { rootDir: REPO_ROOT })
    expect(report.perCapability.find(p => p.capabilityId === 'bar:render-gradient-stops')?.disposition).toBe('verified')
    expect(report.barGradient).toEqual({ 'bar-length': { endBlend: 0.847, settledBy: ['bar:render-gradient-stops'] } })
    const block = renderBarGradientBlock(report.barGradient)
    expect(block).toBe("export const VERIFIED_STAGE3_BAR_GRADIENT: Readonly<Record<string, { endBlend: number; settledBy: readonly string[] }>> = {\n  'bar-length': { endBlend: 0.847, settledBy: ['bar:render-gradient-stops'] },\n}\n")
    expect(renderBarGradientBlock({})).toBe('export const VERIFIED_STAGE3_BAR_GRADIENT: Readonly<Record<string, { endBlend: number; settledBy: readonly string[] }>> = {\n}\n')
  }, 30_000)
})

/**
 * The new case has no capture yet, so it is run through the real harness on a synthetic capture: the generated workbook,
 * Excel's measured geometry convention (6 pt per character, 30 pt rows, 1 px per pt) and bars planted by one rule (solid
 * blue, or the verified gradient). The planted rule must come back as the one match, and a rule outside the family must
 * match none. This checks the family and the gradient-run override end to end before the Excel run, not Excel.
 */
describe('stage3-bar-lengths-default discriminates the 10/90 rules (synthetic capture, Task 22b)', () => {
  const c = caseOf('stage3-bar-lengths-default')
  const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
  const colWidthsPt = [84, 180, 270, 360, 180, 270, 360, 180, 270, 360]
  const rowHeightsPt = Array.from({ length: 24 }, () => 30)
  const geom = computeWorksheetRangeGeometry(parts, c.range, { colWidthsPt, rowHeightsPt })!
  const cls = (col: number): { x14: boolean; solid: boolean } => ({ x14: col <= 6, solid: col <= 3 })
  const half = (x: number): number => Math.floor(x + 0.5 + 1e-9)
  const planted = (rule: string, col: number, v: number, S: number): number => {
    const t = v / 100, [form, scope] = rule.split('@'), k = cls(col)
    const applies = form !== 'linear' && (scope === 'all' || (scope === 'x14' && k.x14) || (scope === 'solid' && k.solid))
    if (rule === 'outside') return Math.max(0, half((0.1 + 0.8 * t) * S) - 20)
    if (applies && form === 'half-min') return half((0.05 + 0.8 * t) * S)
    const L = half((0.1 + 0.8 * t) * S)
    return applies && form === 'offset-9pt' ? Math.max(0, L - 9) : L
  }
  const synthetic = (rule: string): DecodedPng => {
    const width = colWidthsPt.reduce((a, b) => a + b, 0) + 1, height = 24 * 30 + 1
    const data = new Uint8ClampedArray(width * height * 4).fill(255)
    for (const cell of c.expected.cells.filter(x => x.matched)) {
      const col = cell.ref.charCodeAt(0) - 65, row = Number(cell.ref.slice(1)) - 1
      const v = (Number(cell.ref.slice(1)) - 2) * 10
      const x0 = colWidthsPt.slice(0, col).reduce((a, b) => a + b, 0), S = colWidthsPt[col]! - 3
      const L = planted(rule, col, v, S)
      for (let y = row * 30 + 2; y <= row * 30 + 27; y++) {
        for (let i = 0; i < L; i++) {
          const rgb = cls(col).solid ? [0, 0, 255] : predictBarGradientPixel('bar-length', { fill: [0, 0, 255], length: L, span: S }, i, 0.847, []).map(Math.round)
          const j = (y * width + x0 + 2 + i) * 4
          data[j] = rgb[0]!; data[j + 1] = rgb[1]!; data[j + 2] = rgb[2]!
        }
      }
    }
    return { width, height, data }
  }
  const judge = (rule: string) => {
    const png = synthetic(rule)
    const measured = measureCaseBars(c, png, geom, parts)
    expect(measured instanceof Map, `${rule}: ${JSON.stringify(measured).slice(0, 200)}`).toBe(true)
    return evaluateBarCapabilities(c, parts, png, geom, measured as Exclude<typeof measured, { kind: 'refusal' }> as never)['bar:lengths-default-10-90']!
  }

  it('the case designates 99 cells: three fill/pairing classes at three widths, eleven values each', () => {
    expect(c.status).toBe('recipe-pending')
    expect(c.expected.cells.filter(x => x.matched)).toHaveLength(99)
  })

  for (const rule of BAR_LENGTH_RULES) {
    it(`planted ${rule}: the one match`, () => {
      const ev = judge(rule)
      expect([ev.status, ev.matched], ev.reason).toEqual(['discriminating', rule])
    })
  }

  it('a rule outside the family matches none (refused, never fitted)', () => {
    expect(judge('outside').status).toBe('none')
  })
})
