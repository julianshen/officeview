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
import { auditStage3Manifest, barGradientsAcrossCases, renderBarGradientBlock, renderBarGradientBorderBlock } from '../scripts/cf-refs/audit'
import type { CfManifest, CfOracleCase } from '../scripts/cf-refs/cases'
import { CF_MANIFEST_PATH } from '../scripts/cf-refs/generate'
import {
  BAR_COLOR_TOLERANCE,
  BAR_GRADIENT_BORDER_HYPOTHESES,
  BAR_LENGTH_RULES,
  computeWorksheetRangeGeometry,
  fitBarGradientBlend,
  judgeBarGradientBorder,
  predictBarGradientBorderPixel,
  readBarSheetFacts,
  type BarGradientBorderHypothesis,
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
/** Synthetic captures of the 510-cell case: about a second each alone, far more beside the full parallel suite. */
const SYNTHETIC_TIMEOUT_MS = 60_000
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
  }, 60_000)

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
    // Ruling S3-GRAD-BORDER: the bordered shape stays blocked (its case is uncaptured) and exports nothing.
    expect(report.perCapability.find(p => p.capabilityId === 'bar:render-gradient-border')?.disposition).toBe('supported-structural')
    expect(report.barGradientBorder).toEqual({})
    expect(renderBarGradientBorderBlock({})).toBe('export const VERIFIED_STAGE3_BAR_GRADIENT_BORDER: Readonly<Record<string, { endBlend?: number; settledBy: readonly string[] }>> = {\n}\n')
    expect(renderBarGradientBorderBlock({ 'frame-over-ramp': { endBlend: 0.847, settledBy: ['bar:render-gradient-border'] } }))
      .toBe("export const VERIFIED_STAGE3_BAR_GRADIENT_BORDER: Readonly<Record<string, { endBlend?: number; settledBy: readonly string[] }>> = {\n  'frame-over-ramp': { endBlend: 0.847, settledBy: ['bar:render-gradient-border'] },\n}\n")
    expect(renderBarGradientBorderBlock({ 'frame-solid-fill': { settledBy: ['bar:render-gradient-border'] } })).toContain("'frame-solid-fill': { settledBy: ['bar:render-gradient-border'] },")
  }, 60_000)
})

/**
 * The new case has no capture yet, so it is run through the real harness on a synthetic capture: the generated workbook,
 * Excel's measured geometry convention (6 pt per character, 30 pt rows, 1 px per pt) and bars planted by one rule (solid
 * blue, the verified gradient, or one bordered-gradient shape). The planted answer must come back as the one match, and
 * an answer outside the family must match none. This checks the families end to end before the Excel run, not Excel.
 */
describe('stage3-bar-lengths-default discriminates the 10/90 rules and the bordered gradient (synthetic capture)', () => {
  const c = caseOf('stage3-bar-lengths-default')
  const parts = readZipPartsSync(new Uint8Array(readFileSync(resolve(REPO_ROOT, c.saved))))!
  const facts = readBarSheetFacts(parts)
  const colWidthsPt = [84, ...Array.from({ length: 7 }, () => [180, 270, 360]).flat()]
  const rows = Number(/\d+$/.exec(c.range)![0])
  const rowHeightsPt = Array.from({ length: rows }, () => 30)
  const geom = computeWorksheetRangeGeometry(parts, c.range, { colWidthsPt, rowHeightsPt })!
  // B:D paired solid, E:G paired gradient, H:J base-only; K:P bordered 0/100, Q:V bordered 10/90.
  const cls = (col: number): { x14: boolean; solid: boolean; bordered: boolean; full: boolean; green: boolean } =>
    ({ x14: col <= 6 || col >= 10, solid: col <= 3, bordered: col >= 10, full: col >= 10 && col <= 15, green: (col >= 13 && col <= 15) || col >= 19 })
  const half = (x: number): number => Math.floor(x + 0.5 + 1e-9)
  const plantedLength = (rule: string, col: number, v: number, S: number): number => {
    const t = v / 100, [form, scope] = rule.split('@'), k = cls(col)
    if (k.full) return half(t * S)
    if (k.bordered) return half((0.1 + 0.8 * t) * S)
    const applies = form !== 'linear' && (scope === 'all' || (scope === 'x14' && k.x14) || (scope === 'solid' && k.solid))
    if (rule === 'outside') return Math.max(0, half((0.1 + 0.8 * t) * S) - 20)
    if (applies && form === 'half-min') return half((0.05 + 0.8 * t) * S)
    const L = half((0.1 + 0.8 * t) * S)
    return applies && form === 'offset-9pt' ? Math.max(0, L - 9) : L
  }
  /** `rule` plants the 10/90 columns; `border` plants the bordered ones ('outside': a 2 px frame). */
  const synthetic = (rule: string, border: string): DecodedPng => {
    const width = colWidthsPt.reduce((a, b) => a + b, 0) + 1, height = rows * 30 + 1
    const data = new Uint8ClampedArray(width * height * 4).fill(255)
    for (const cell of c.expected.cells.filter(x => x.matched)) {
      const col = cell.ref.charCodeAt(0) - 65, row = Number(cell.ref.slice(1)) - 1
      const v = (facts.values.get(cell.ref) as { v: number }).v
      const x0 = colWidthsPt.slice(0, col).reduce((a, b) => a + b, 0), S = colWidthsPt[col]! - 3
      const L = plantedLength(rule, col, v, S), k = cls(col)
      const top = row * 30 + 2, bottom = row * 30 + 27
      const sample = { ref: cell.ref, fill: [0, 0, 255] as [number, number, number], border: (k.green ? [0, 0xb0, 0x50] : [0, 0, 255]) as [number, number, number], leftPx: x0 + 2, topPx: top, bottomPx: bottom, length: L, span: S, pixels: [] }
      for (let y = top; y <= bottom; y++) {
        for (let i = 0; i < L; i++) {
          let rgb: number[]
          if (k.solid) rgb = [0, 0, 255]
          else if (!k.bordered) rgb = predictBarGradientPixel('bar-length', sample, i, 0.847, []).map(Math.round)
          else if (border === 'outside') {
            const frame2 = i <= 1 || i >= L - 2 || y <= top + 1 || y >= bottom - 1
            rgb = frame2 ? sample.border : predictBarGradientPixel('bar-length', sample, i, 0.847, []).map(Math.round)
          } else rgb = predictBarGradientBorderPixel(border as BarGradientBorderHypothesis, sample, i, y, 0.847).map(Math.round)
          const j = (y * width + x0 + 2 + i) * 4
          data[j] = rgb[0]!; data[j + 1] = rgb[1]!; data[j + 2] = rgb[2]!
        }
      }
    }
    return { width, height, data }
  }
  const judge = (rule: string, border: string, cap: string) => {
    const png = synthetic(rule, border)
    const measured = measureCaseBars(c, png, geom, parts)
    expect(measured instanceof Map, `${rule}/${border}: ${JSON.stringify(measured).slice(0, 200)}`).toBe(true)
    return evaluateBarCapabilities(c, parts, png, geom, measured as never)[cap]!
  }

  it('the case designates 255 cells: 99 for the 10/90 rules, 156 bordered gradients (twelve columns of thirteen values)', () => {
    expect(c.status).toBe('recipe-pending')
    expect(c.expected.cells.filter(x => x.matched)).toHaveLength(255)
  })

  for (const rule of BAR_LENGTH_RULES) {
    it(`planted ${rule}: the one 10/90 match`, () => {
      const ev = judge(rule, 'frame-over-ramp', 'bar:lengths-default-10-90')
      expect([ev.status, ev.matched], ev.reason).toEqual(['discriminating', rule])
    }, SYNTHETIC_TIMEOUT_MS)
  }

  it('a 10/90 rule outside the family matches none (refused, never fitted)', () => {
    expect(judge('outside', 'frame-over-ramp', 'bar:lengths-default-10-90').status).toBe('none')
  }, SYNTHETIC_TIMEOUT_MS)

  /** The bordered judgement exactly as evaluateBarCapabilities makes it, without re-measuring every other capability. */
  const judgeBordered = (shape: string) => {
    const png = synthetic('linear', shape)
    const refs = new Set(c.expected.cells.filter(x => x.matched && x.ref.charCodeAt(0) >= 75).map(x => x.ref))
    return judgeBarGradientBorder(sampleBarGradients(c, parts, png, geom, refs, 'bordered'), fitBarGradientBlend(sampleBarGradients(c, parts, png, geom).samples))
  }

  for (const shape of BAR_GRADIENT_BORDER_HYPOTHESES) {
    it(`planted bordered ${shape}: the one bordered-gradient match`, () => {
      const ev = judgeBordered(shape)
      expect([ev.status, ev.matched], ev.reason).toEqual(['discriminating', shape])
    }, SYNTHETIC_TIMEOUT_MS)
  }

  it('a bordered shape outside the family (a 2 px frame) matches none', () => {
    expect(judgeBordered('outside').status).toBe('none')
  }, SYNTHETIC_TIMEOUT_MS)

  it('evaluateBarCapabilities gives the same bordered verdict (one end-to-end run)', () => {
    const ev = judge('linear', 'frame-inner-ramp', 'bar:render-gradient-border')
    expect([ev.status, ev.matched], ev.reason).toEqual(['discriminating', 'frame-inner-ramp'])
  }, SYNTHETIC_TIMEOUT_MS)

  it('the bordered judge fits the planted blend and checks the 0/100 extents against the solid model (Ruling S3-M)', () => {
    const png = synthetic('linear', 'frame-over-ramp')
    const refs = new Set(c.expected.cells.filter(x => x.matched && x.ref.charCodeAt(0) >= 75).map(x => x.ref))
    const o = sampleBarGradients(c, parts, png, geom, refs, 'bordered')
    expect(o.samples.length).toBe(150) // 156 cells less the six zeros of the 0/100 columns, which draw no bar
    const j = judgeBarGradientBorder(o, 0.847)
    expect([j.matched, j.endBlend]).toEqual(['frame-over-ramp', 0.847])
    const shifted = { ...o, samples: o.samples.map(s => (s.predictedLength !== undefined && s.length > 20 ? { ...s, predictedLength: s.length + 2 } : s)) }
    expect(judgeBarGradientBorder(shifted, 0.847).status).toBe('none')
    // Without the capture's borderless blend there is nothing to predict the ramps with.
    expect(judgeBarGradientBorder(o, undefined).status).toBe('no-hypotheses')
    // A ramp of a different blend than the borderless bars' matches no shape: refused, never refitted.
    expect(judgeBarGradientBorder(o, 0.6).status).toBe('none')
    // The plain gradient sampler leaves every bordered bar to this judge.
    expect(sampleBarGradients(c, parts, png, geom, refs).samples).toEqual([])
  }, SYNTHETIC_TIMEOUT_MS)
})
