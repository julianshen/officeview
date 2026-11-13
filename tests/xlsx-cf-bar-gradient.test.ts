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
  judgeBarGradient,
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
