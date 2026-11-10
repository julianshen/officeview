import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import {
  percentileOrderStatistic,
  resolveScaleThresholds,
  scaleColorAt,
  interpolateScaleColor,
  interpolateChannel,
} from '../src/xlsx/cf/colorScale'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfRangeStatistics } from '../src/xlsx/cf/statistics'
import type { CfScalePoint, CfLiteral } from '../src/xlsx/cf/model'

function readFixture(name: string) {
  const filePath = resolve(process.cwd(), `tests/fixtures/xlsx-cf/refs/${name}.json`)
  return JSON.parse(readFileSync(filePath, 'utf8'))
}

function makeStats(numbers: number[]): CfRangeStatistics {
  const sorted = [...numbers].sort((a, b) => a - b)
  const min = sorted.length > 0 ? sorted[0] : undefined
  const max = sorted.length > 0 ? sorted[sorted.length - 1] : undefined
  const sum = numbers.reduce((a, b) => a + b, 0)
  const mean = numbers.length > 0 ? sum / numbers.length : undefined
  const frequencies = new Map<string, number>()
  for (const n of numbers) {
    const k = `n:${n}`
    frequencies.set(k, (frequencies.get(k) ?? 0) + 1)
  }
  return {
    scopeKey: 'test-scope',
    numericCount: numbers.length,
    sum,
    mean,
    min,
    max,
    stddevPop: undefined,
    stddevSample: undefined,
    sortedAsc: new Float64Array(sorted),
    frequencies,
    declaredCells: numbers.length,
    spanCells: numbers.length,
    storedCells: numbers.length,
  }
}

describe('CF Stage 2 Task 14 — Scale evaluation, percentile and interpolation', () => {
  it('ruling T13-A: a 2-point rgb min/max scale is unresolved with only scale:thresholds-min-max', () => {
    const points: readonly [CfScalePoint, CfScalePoint] = [
      { bound: { type: 'min' }, argb: 0xffff0000 },
      { bound: { type: 'max' }, argb: 0xff00ff00 },
    ]
    const stats = makeStats([10, 20, 30, 40, 50, 60, 70])

    // With only scale:thresholds-min-max injected
    withExactCfEvidence(['scale:thresholds-min-max'], () => {
      const thresholds = resolveScaleThresholds(points, stats)
      expect(thresholds.kind).toBe('ready')
      if (thresholds.kind === 'ready') {
        expect(thresholds.value).toEqual([10, 70])
      }
      // Evaluation requires scale:interpolation-color-space and scale:interpolation-rounding-ladder
      const outcome = scaleColorAt(10, points, [10, 70])
      expect(outcome.kind).toBe('unresolved')
    })

    // With full required capabilities
    const fullRequired = [
      'scale:thresholds-min-max',
      'scale:interpolation-color-space',
      'scale:interpolation-rounding-ladder',
    ]
    withExactCfEvidence(fullRequired, () => {
      const outcome = scaleColorAt(10, points, [10, 70])
      expect(outcome.kind).toBe('ready')
      if (outcome.kind === 'ready') {
        expect(outcome.value).toBe(0xffff0000)
      }
    })
  })

  it('percentile order statistic matches its definition', () => {
    // Pure-function test of the closed variant set over {1, 2, 3, 4, 10}
    const data = new Float64Array([1, 2, 3, 4, 10])

    // Over {1,2,3,4,10}, the 30th percentile is 2.2 (INC) and 1.8 (EXC)
    expect(percentileOrderStatistic(data, 30, 'inc')).toBeCloseTo(2.2, 8)
    expect(percentileOrderStatistic(data, 30, 'exc')).toBeCloseTo(1.8, 8)

    // and the 50th is 3 under both.
    expect(percentileOrderStatistic(data, 50, 'inc')).toBeCloseTo(3.0, 8)
    expect(percentileOrderStatistic(data, 50, 'exc')).toBeCloseTo(3.0, 8)

    // Boundaries
    expect(percentileOrderStatistic(data, 0, 'inc')).toBeCloseTo(1.0, 8)
    expect(percentileOrderStatistic(data, 100, 'inc')).toBeCloseTo(10.0, 8)
    expect(percentileOrderStatistic(data, 0, 'exc')).toBeUndefined()
    expect(percentileOrderStatistic(data, 100, 'exc')).toBeUndefined()

    // Second test pins the evidence-selected variant to stage2-scale-thresholds manifest cells:
    // Column E in stage2-scale-thresholds uses 3-point min / percentile 30 / max over {1,2,3,4,10}.
    const points: readonly [CfScalePoint, CfScalePoint, CfScalePoint] = [
      { bound: { type: 'min' }, argb: 0xfff8696b },
      { bound: { type: 'percentile', value: 30 }, argb: 0xffffeb84 },
      { bound: { type: 'max' }, argb: 0xff63be7b },
    ]
    const stats = makeStats([1, 2, 3, 4, 10])

    // Without evidence the result is unresolved
    withExactCfEvidence(['scale:thresholds-min-max', 'scale:interpolation-spec-palette', 'scale:interpolation-color-space', 'scale:interpolation-rounding-ladder'], () => {
      const thresholds = resolveScaleThresholds(points, stats)
      expect(thresholds.kind).toBe('unresolved')
    })

    // With evidence: selects INC (2.2), resolving E3 (value 2) exactly
    const fullEvidence = [
      'scale:thresholds-min-max',
      'scale:thresholds-percentile-inc-exc',
      'scale:interpolation-spec-palette',
      'scale:interpolation-color-space',
      'scale:interpolation-rounding-ladder',
    ]
    withExactCfEvidence(fullEvidence, () => {
      const thresholds = resolveScaleThresholds(points, stats)
      expect(thresholds.kind).toBe('ready')
      if (thresholds.kind === 'ready') {
        expect(thresholds.value[0]).toBe(1)
        expect(thresholds.value[1]).toBeCloseTo(2.2, 8)
        expect(thresholds.value[2]).toBe(10)

        // Cell E3 (value 2) fill under INC matches observed Excel display format FFFDD57F
        const color = scaleColorAt(2, points, thresholds.value)
        expect(color.kind).toBe('ready')
        if (color.kind === 'ready') {
          const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
          expect(hex).toBe('FFFDD57F')
        }
      }
    })
  })

  it('closed interpolation-variant set (S2-B)', () => {
    // Pure-function test of rounding ladder variants
    // Truncate (floor towards 0 when increasing, floor towards start when decreasing)
    expect(interpolateChannel(0, 255, 1 / 510, 'truncate')).toBe(0) // 0.5 -> 0
    expect(interpolateChannel(0, 255, 3 / 510, 'truncate')).toBe(1) // 1.5 -> 1
    expect(interpolateChannel(0, 255, 5 / 510, 'truncate')).toBe(2) // 2.5 -> 2
    expect(interpolateChannel(0, 255, 255 / 510, 'truncate')).toBe(127) // 127.5 -> 127
    expect(interpolateChannel(255, 0, 255 / 510, 'truncate')).toBe(128) // 255 - 127 = 128

    // Half-up
    expect(interpolateChannel(0, 255, 1 / 510, 'half-up')).toBe(1) // 0.5 -> 1
    expect(interpolateChannel(0, 255, 3 / 510, 'half-up')).toBe(2) // 1.5 -> 2
    expect(interpolateChannel(0, 255, 5 / 510, 'half-up')).toBe(3) // 2.5 -> 3
    expect(interpolateChannel(0, 255, 255 / 510, 'half-up')).toBe(128)

    // Half-even
    expect(interpolateChannel(0, 255, 1 / 510, 'half-even')).toBe(0) // 0.5 -> 0 (even)
    expect(interpolateChannel(0, 255, 3 / 510, 'half-even')).toBe(2) // 1.5 -> 2 (even)
    expect(interpolateChannel(0, 255, 5 / 510, 'half-even')).toBe(2) // 2.5 -> 2 (even)

    // Pure-function test of colour space variants on (255,0,0) -> (0,0,255) at 0.5
    const cRed = 0xffff0000
    const cBlue = 0xff0000ff
    const srgbColor = interpolateScaleColor(cRed, cBlue, 0.5, 'sRGB', 'truncate')
    expect(srgbColor.toString(16).toUpperCase().padStart(8, '0')).toBe('FF80007F')

    const linearColor = interpolateScaleColor(cRed, cBlue, 0.5, 'linear-light', 'truncate')
    expect(linearColor).not.toBe(srgbColor)

    const hslColor = interpolateScaleColor(cRed, cBlue, 0.5, 'hsl', 'truncate')
    expect(hslColor).not.toBe(srgbColor)
  })

  it('no division by zero on degenerate scopes', () => {
    const points: readonly [CfScalePoint, CfScalePoint] = [
      { bound: { type: 'min' }, argb: 0xfff8696b },
      { bound: { type: 'max' }, argb: 0xff63be7b },
    ]

    // 1. v_min = v_max (all values equal, e.g. [5, 5, 5, 5, 5])
    const equalStats = makeStats([5, 5, 5, 5, 5])
    withExactCfEvidence(['scale:thresholds-min-max', 'scale:interpolation-color-space', 'scale:interpolation-rounding-ladder'], () => {
      // Without scale:degenerate-equal-values evidence: unresolved
      const res = resolveScaleThresholds(points, equalStats)
      expect(res.kind).toBe('unresolved')
    })

    withExactCfEvidence([
      'scale:thresholds-min-max',
      'scale:degenerate-equal-values',
      'scale:interpolation-color-space',
      'scale:interpolation-rounding-ladder',
    ], () => {
      const res = resolveScaleThresholds(points, equalStats)
      expect(res.kind).toBe('ready')
      if (res.kind === 'ready') {
        const color = scaleColorAt(5, points, res.value)
        expect(color.kind).toBe('ready')
        if (color.kind === 'ready') {
          expect(Number.isNaN(color.value)).toBe(false)
          expect(Number.isFinite(color.value)).toBe(true)
          expect(color.value.toString(16).toUpperCase()).toBe('FF63BE7B')
        }
      }
    })

    // 2. Single numeric cell (e.g. [42])
    const singleStats = makeStats([42])
    withExactCfEvidence(['scale:thresholds-min-max', 'scale:interpolation-color-space', 'scale:interpolation-rounding-ladder'], () => {
      const res = resolveScaleThresholds(points, singleStats)
      expect(res.kind).toBe('unresolved')
    })

    withExactCfEvidence([
      'scale:thresholds-min-max',
      'scale:degenerate-single-cell',
      'scale:interpolation-color-space',
      'scale:interpolation-rounding-ladder',
    ], () => {
      const res = resolveScaleThresholds(points, singleStats)
      expect(res.kind).toBe('ready')
      if (res.kind === 'ready') {
        const color = scaleColorAt(42, points, res.value)
        expect(color.kind).toBe('ready')
        if (color.kind === 'ready') {
          expect(Number.isNaN(color.value)).toBe(false)
          expect(Number.isFinite(color.value)).toBe(true)
          expect(color.value.toString(16).toUpperCase()).toBe('FF63BE7B')
        }
      }
    })

    // 3. t_low = t_high (e.g. num 50 to num 50)
    const fixedPoints: readonly [CfScalePoint, CfScalePoint] = [
      { bound: { type: 'num', value: 50 }, argb: 0xfff8696b },
      { bound: { type: 'num', value: 50 }, argb: 0xff63be7b },
    ]
    withExactCfEvidence([
      'scale:thresholds-num-bounds',
      'scale:degenerate-equal-values',
      'scale:interpolation-color-space',
      'scale:interpolation-rounding-ladder',
    ], () => {
      const res = resolveScaleThresholds(fixedPoints, equalStats)
      expect(res.kind).toBe('ready')
      if (res.kind === 'ready') {
        const color = scaleColorAt(50, fixedPoints, res.value)
        expect(color.kind).toBe('ready')
        if (color.kind === 'ready') {
          expect(Number.isNaN(color.value)).toBe(false)
          expect(Number.isFinite(color.value)).toBe(true)
        }
      }
    })
  })

  it('oracle cells reproduce exactly', () => {
    // Test that for verified Task 12 variants, the evaluator's ARGB equals
    // the manifest's trusted display format ARGB exactly.
    // Expectations are read directly from tests/fixtures/xlsx-cf/refs/stage2-scale-*.json observedCells[].fillRgb
    // where fillSource is 'excel-display-format' and hidden is not true.
    // Never compute an expected colour with the function under test!

    // Case 1: stage2-scale-interpolation
    const interpFixture = readFixture('stage2-scale-interpolation')
    const interpEvidence = [
      'scale:interpolation-precondition-probe',
      'scale:interpolation-rounding-ladder',
      'scale:interpolation-color-space',
      'scale:interpolation-spec-palette',
      'scale:thresholds-min-max',
      'scale:thresholds-num-bounds',
      'scale:thresholds-percentile-inc-exc',
    ]

    withExactCfEvidence(interpEvidence, () => {
      // Column B: 2-colour min/max over 10..70
      const bPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'min' }, argb: 0xffff0000 },
        { bound: { type: 'max' }, argb: 0xff00ff00 },
      ]
      const bStats = makeStats([10, 20, 30, 40, 50, 60, 70])
      const bThresholds = resolveScaleThresholds(bPoints, bStats)
      expect(bThresholds.kind).toBe('ready')
      if (bThresholds.kind === 'ready') {
        for (const cell of interpFixture.observedCells) {
          if (!cell.ref.match(/^B[2-8]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, bPoints, bThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }

      // Column C: num 0 to num 510 rounding ladder
      const cPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'num', value: 0 }, argb: 0xff000000 },
        { bound: { type: 'num', value: 510 }, argb: 0xffffffff },
      ]
      const cThresholds = resolveScaleThresholds(cPoints, bStats)
      expect(cThresholds.kind).toBe('ready')
      if (cThresholds.kind === 'ready') {
        for (const cell of interpFixture.observedCells) {
          if (!cell.ref.match(/^C[2-8]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, cPoints, cThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }

      // Column D: color space over 0, 50, 100
      const dPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'min' }, argb: 0xffff0000 },
        { bound: { type: 'max' }, argb: 0xff0000ff },
      ]
      const dStats = makeStats([0, 50, 100])
      const dThresholds = resolveScaleThresholds(dPoints, dStats)
      expect(dThresholds.kind).toBe('ready')
      if (dThresholds.kind === 'ready') {
        for (const cell of interpFixture.observedCells) {
          if (!cell.ref.match(/^D[2-4]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, dPoints, dThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }

      // Column E: spec palette over 1..9 with midpoint 5
      const ePoints: readonly [CfScalePoint, CfScalePoint, CfScalePoint] = [
        { bound: { type: 'min' }, argb: 0xfff8696b },
        { bound: { type: 'percentile', value: 50 }, argb: 0xffffeb84 },
        { bound: { type: 'max' }, argb: 0xff63be7b },
      ]
      const eStats = makeStats([1, 2, 3, 4, 5, 5, 6, 7, 8, 9])
      const eThresholds = resolveScaleThresholds(ePoints, eStats)
      expect(eThresholds.kind).toBe('ready')
      if (eThresholds.kind === 'ready') {
        for (const cell of interpFixture.observedCells) {
          if (!cell.ref.match(/^E([2-9]|1[0-1])$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, ePoints, eThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }
    })

    // Case 2: stage2-scale-thresholds
    const threshFixture = readFixture('stage2-scale-thresholds')
    const threshEvidence = [
      'scale:thresholds-min-max',
      'scale:thresholds-num-bounds',
      'scale:thresholds-percent',
      'scale:thresholds-percentile-inc-exc',
      'scale:degenerate-values-beyond-bounds',
      'scale:interpolation-spec-palette',
      'scale:interpolation-color-space',
      'scale:interpolation-rounding-ladder',
    ]

    withExactCfEvidence(threshEvidence, () => {
      // Column B: min/max over 1, 2, 3, 4, 10
      const bPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'min' }, argb: 0xfff8696b },
        { bound: { type: 'max' }, argb: 0xff63be7b },
      ]
      const bStats = makeStats([1, 2, 3, 4, 10])
      const bThresholds = resolveScaleThresholds(bPoints, bStats)
      expect(bThresholds.kind).toBe('ready')
      if (bThresholds.kind === 'ready') {
        for (const cell of threshFixture.observedCells) {
          if (!cell.ref.match(/^B[2-6]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, bPoints, bThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }

      // Column C: num 20 to num 80 over 10, 30, 50, 70, 90
      const cPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'num', value: 20 }, argb: 0xfff8696b },
        { bound: { type: 'num', value: 80 }, argb: 0xff63be7b },
      ]
      const cThresholds = resolveScaleThresholds(cPoints, bStats)
      expect(cThresholds.kind).toBe('ready')
      if (cThresholds.kind === 'ready') {
        for (const cell of threshFixture.observedCells) {
          if (!cell.ref.match(/^C[2-6]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, cPoints, cThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }

      // Column E: min / percentile 30 / max
      const ePoints: readonly [CfScalePoint, CfScalePoint, CfScalePoint] = [
        { bound: { type: 'min' }, argb: 0xfff8696b },
        { bound: { type: 'percentile', value: 30 }, argb: 0xffffeb84 },
        { bound: { type: 'max' }, argb: 0xff63be7b },
      ]
      const eThresholds = resolveScaleThresholds(ePoints, bStats)
      expect(eThresholds.kind).toBe('ready')
      if (eThresholds.kind === 'ready') {
        for (const cell of threshFixture.observedCells) {
          if (!cell.ref.match(/^E[2-6]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, ePoints, eThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }
    })

    // Case 3: stage2-scale-degenerate
    const degenFixture = readFixture('stage2-scale-degenerate')
    const degenEvidence = [
      'scale:thresholds-min-max',
      'scale:thresholds-num-bounds',
      'scale:degenerate-equal-values',
      'scale:degenerate-single-cell',
      'scale:degenerate-inverted-thresholds',
      'scale:degenerate-midpoint-above-max',
      'scale:degenerate-values-beyond-bounds',
      'scale:interpolation-spec-palette',
      'scale:interpolation-color-space',
      'scale:interpolation-rounding-ladder',
    ]

    withExactCfEvidence(degenEvidence, () => {
      // Column B: all values equal (5)
      const bPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'min' }, argb: 0xfff8696b },
        { bound: { type: 'max' }, argb: 0xff63be7b },
      ]
      const bStats = makeStats([5, 5, 5, 5, 5])
      const bThresholds = resolveScaleThresholds(bPoints, bStats)
      expect(bThresholds.kind).toBe('ready')
      if (bThresholds.kind === 'ready') {
        for (const cell of degenFixture.observedCells) {
          if (!cell.ref.match(/^B[2-6]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, bPoints, bThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }

      // Column C: single cell C2 (42)
      const cStats = makeStats([42])
      const cThresholds = resolveScaleThresholds(bPoints, cStats)
      expect(cThresholds.kind).toBe('ready')
      if (cThresholds.kind === 'ready') {
        const c2Cell = degenFixture.observedCells.find((c: any) => c.ref === 'C2')
        if (c2Cell && c2Cell.fillSource === 'excel-display-format') {
          const color = scaleColorAt(42, bPoints, cThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, 'Cell C2').toBe(c2Cell.fillRgb)
          }
        }
      }

      // Column D: inverted thresholds 100 -> 10 over values 10, 30, 50, 70, 100
      const dPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'num', value: 100 }, argb: 0xfff8696b },
        { bound: { type: 'num', value: 10 }, argb: 0xff63be7b },
      ]
      const dStats = makeStats([10, 30, 50, 70, 100])
      const dThresholds = resolveScaleThresholds(dPoints, dStats)
      expect(dThresholds.kind).toBe('ready')
      if (dThresholds.kind === 'ready') {
        for (const cell of degenFixture.observedCells) {
          if (!cell.ref.match(/^D[2-6]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, dPoints, dThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }

      // Column E: midpoint (200) above max (100) over values 10, 30, 50, 70, 100
      const ePoints: readonly [CfScalePoint, CfScalePoint, CfScalePoint] = [
        { bound: { type: 'min' }, argb: 0xfff8696b },
        { bound: { type: 'num', value: 200 }, argb: 0xffffeb84 },
        { bound: { type: 'num', value: 100 }, argb: 0xff63be7b },
      ]
      const eThresholds = resolveScaleThresholds(ePoints, dStats)
      expect(eThresholds.kind).toBe('ready')
      if (eThresholds.kind === 'ready') {
        for (const cell of degenFixture.observedCells) {
          if (!cell.ref.match(/^E[2-6]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, ePoints, eThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }

      // Column F: values beyond bounds (num 20 to num 80, values 5, 20, 50, 80, 95)
      const fPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'num', value: 20 }, argb: 0xfff8696b },
        { bound: { type: 'num', value: 80 }, argb: 0xff63be7b },
      ]
      const fThresholds = resolveScaleThresholds(fPoints, dStats)
      expect(fThresholds.kind).toBe('ready')
      if (fThresholds.kind === 'ready') {
        for (const cell of degenFixture.observedCells) {
          if (!cell.ref.match(/^F[2-6]$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue
          const val = Number(cell.formattedText)
          const color = scaleColorAt(val, fPoints, fThresholds.value)
          expect(color.kind).toBe('ready')
          if (color.kind === 'ready') {
            const hex = color.value.toString(16).toUpperCase().padStart(8, '0')
            expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
          }
        }
      }
    })

    // Case 4: stage2-scale-participation
    const partFixture = readFixture('stage2-scale-participation')
    const partEvidence = [
      'scale:thresholds-min-max',
      'scale:participation-types',
      'scale:participation-error-presence',
      'scale:interpolation-color-space',
      'scale:interpolation-rounding-ladder',
    ]

    withExactCfEvidence(partEvidence, () => {
      // Column B: mixed types (numbers, date serial, TRUE, FALSE, text "5", #N/A, #DIV/0!, "")
      const bPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'min' }, argb: 0xfff8696b },
        { bound: { type: 'max' }, argb: 0xff63be7b },
      ]
      // Numbers in Column B are 10, 50, 90, 45000.0
      const bStats = makeStats([10, 50, 90, 45000])
      const bThresholds = resolveScaleThresholds(bPoints, bStats)
      expect(bThresholds.kind).toBe('ready')
      if (bThresholds.kind === 'ready') {
        for (const cell of partFixture.observedCells) {
          if (!cell.ref.match(/^B([2-9]|1[0-1])$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue

          let inputLiteral: CfLiteral | number
          if (cell.ref === 'B6') inputLiteral = { kind: 'boolean', value: true }
          else if (cell.ref === 'B7') inputLiteral = { kind: 'boolean', value: false }
          else if (cell.ref === 'B8') inputLiteral = { kind: 'string', value: '5' }
          else if (cell.ref === 'B9') inputLiteral = { kind: 'error', value: '#N/A' }
          else if (cell.ref === 'B10') inputLiteral = { kind: 'error', value: '#DIV/0!' }
          else if (cell.ref === 'B11') inputLiteral = { kind: 'blank' }
          else inputLiteral = Number(cell.formattedText)

          const outcome = scaleColorAt(inputLiteral, bPoints, bThresholds.value)
          if (cell.fillRgb === 'FFFFFFFF') {
            // Uncoloured / no match
            expect(outcome.kind === 'no-match' || outcome.kind === 'ready' && outcome.value === 0xffffffff, `Cell ${cell.ref}`).toBe(true)
          } else {
            expect(outcome.kind).toBe('ready')
            if (outcome.kind === 'ready') {
              const hex = outcome.value.toString(16).toUpperCase().padStart(8, '0')
              expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
            }
          }
        }
      }

      // Column C: error presence (#N/A at C9) does not suppress other cells
      const cNumbers = [10, 20, 30, 40, 50, 60, 70, 80, 90]
      const cStats = makeStats(cNumbers)
      const cThresholds = resolveScaleThresholds(bPoints, cStats)
      expect(cThresholds.kind).toBe('ready')
      if (cThresholds.kind === 'ready') {
        for (const cell of partFixture.observedCells) {
          if (!cell.ref.match(/^C([2-9]|1[0-1])$/)) continue
          if (cell.fillSource !== 'excel-display-format' || cell.hidden === true) continue

          let inputLiteral: CfLiteral | number
          if (cell.ref === 'C9') inputLiteral = { kind: 'error', value: '#N/A' }
          else inputLiteral = Number(cell.formattedText)

          const outcome = scaleColorAt(inputLiteral, bPoints, cThresholds.value)
          if (cell.fillRgb === 'FFFFFFFF') {
            expect(outcome.kind === 'no-match' || outcome.kind === 'ready' && outcome.value === 0xffffffff, `Cell ${cell.ref}`).toBe(true)
          } else {
            expect(outcome.kind).toBe('ready')
            if (outcome.kind === 'ready') {
              const hex = outcome.value.toString(16).toUpperCase().padStart(8, '0')
              expect(hex, `Cell ${cell.ref}`).toBe(cell.fillRgb)
            }
          }
        }
      }
    })

    // Without evidence: assert unresolved
    withExactCfEvidence([], () => {
      const bPoints: readonly [CfScalePoint, CfScalePoint] = [
        { bound: { type: 'min' }, argb: 0xffff0000 },
        { bound: { type: 'max' }, argb: 0xff00ff00 },
      ]
      const bStats = makeStats([10, 20, 30])
      expect(resolveScaleThresholds(bPoints, bStats).kind).toBe('unresolved')
    })
  })
})
