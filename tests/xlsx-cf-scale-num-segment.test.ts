// @vitest-environment node
/**
 * Finding (capture e448d831, stage2-scale-scope-disjoint twins T4 and V3): on a 3-colour scale with literal num bounds
 * 0/30/60, values 40 and 50 (the upper segment at f = 1/3 and 2/3) render one unit HIGHER per RGB channel in Excel 16.115
 * (T4 CC/DD/82 against CB/DC/81; V3 98/CE/7F against 97/CD/7E: the descending step one unit smaller in magnitude) than
 * production's truncating interpolation. Excel's rule is unexplained; the rounding is not changed (that would fit the
 * capture). User ruling: withhold the shape (`cf-unmeasured-num-segment-fraction`) where ALL hold:
 * - the scale has 3 colours;
 * - both cfvo bounding the segment the value falls in are literal `num` (a formula bound resolved to a number does not
 *   count: no capture shows it in this shape);
 * - f·2^20 is not an integer (a finite binary-resolution test, not exact non-dyadicity);
 * - the value lies strictly inside the segment.
 * Endpoints, fractions with f·2^20 an integer, 2-colour scales and segments bounded by min/max/percent/percentile paint exactly as before.
 */
import { describe, expect, it } from 'vitest'
import { scaleColorAt, interpolateScaleColor } from '../src/xlsx/cf/colorScale'
import type { CfScalePoint } from '../src/xlsx/cf/model'

const RED = 0xfff8696b, YELLOW = 0xffffeb84, GREEN = 0xff63be7b
const num = (value: number, argb: number, resolved?: 'formula'): CfScalePoint => ({ bound: { type: 'num', value, ...(resolved ? { resolved } : {}) }, argb })
const all = { hasEvidence: () => true }
const NUM3 = [num(0, RED), num(30, YELLOW), num(60, GREEN)]
const WITHHELD = { kind: 'unresolved', reason: 'cf-unmeasured-num-segment-fraction' }

describe('3-colour scale, literal num segment, f·2^20 not an integer: withheld', () => {
  it('T4 and V3 of the capture (num 0/30/60, values 40 and 50) are withheld', () => {
    expect(scaleColorAt(40, NUM3, [0, 30, 60], all)).toEqual(WITHHELD)
    expect(scaleColorAt(50, NUM3, [0, 30, 60], all)).toEqual(WITHHELD)
  })

  it('the lower segment of the same rule (values 10 and 20, f = 1/3 and 2/3) is withheld too', () => {
    expect(scaleColorAt(10, NUM3, [0, 30, 60], all)).toEqual(WITHHELD)
    expect(scaleColorAt(20, NUM3, [0, 30, 60], all)).toEqual(WITHHELD)
  })

  it('fractions with f·2^20 an integer and endpoints are not this shape (U4 of the capture: num 30/60/100, value 70, f = 1/4, is the upper-segment whole-product shape since stage2-scale-num-segment)', () => {
    const P = [num(30, RED), num(60, YELLOW), num(100, GREEN)]
    expect(scaleColorAt(70, P, [30, 60, 100], all)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-upper-segment-whole-product' })
    expect(scaleColorAt(45, P, [30, 60, 100], all)).toEqual({ kind: 'ready', value: interpolateScaleColor(RED, YELLOW, 0.5) })
    for (const v of [0, 30, 60]) expect(scaleColorAt(v, NUM3, [0, 30, 60], all).kind, String(v)).toBe('ready')
    expect(scaleColorAt(15, NUM3, [0, 30, 60], all).kind).toBe('ready') // f = 1/2, lower segment
    expect(scaleColorAt(45, NUM3, [0, 30, 60], all)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-upper-segment-whole-product' }) // f = 1/2, upper segment: 78 whole
  })

  it('values beyond the bounds (clamped to an endpoint) paint as before', () => {
    expect(scaleColorAt(80, NUM3, [0, 30, 60], all)).toEqual({ kind: 'ready', value: GREEN })
  })

  it('a segment with a non-num bound paints as before (min/percentile/max: stage2-scale-thresholds E4 shape)', () => {
    const P: CfScalePoint[] = [{ bound: { type: 'min' }, argb: RED }, { bound: { type: 'percentile', value: 30 }, argb: YELLOW }, { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(3, P, [1, 2.2, 10], all)).toEqual({ kind: 'ready', value: 0xffefe784 })
    // Only one of the two bounds of the segment is num: not the shape.
    const mixed: CfScalePoint[] = [num(0, RED), num(30, YELLOW), { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, mixed, [0, 30, 60], all)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-upper-segment-whole-product' }) // not this shape, but the upper-segment whole-product one
    expect(scaleColorAt(41, mixed, [0, 30, 60], all).kind).toBe('ready')
    expect(scaleColorAt(10, mixed, [0, 30, 60], all)).toEqual(WITHHELD)
  })

  it('a formula bound resolved to a number is not the shape (no capture shows it)', () => {
    const P = [num(0, RED), num(30, YELLOW, 'formula'), num(60, GREEN)]
    expect(scaleColorAt(40, P, [0, 30, 60], all)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-upper-segment-whole-product' }) // not this shape, but the upper-segment whole-product one
    expect(scaleColorAt(41, P, [0, 30, 60], all).kind).toBe('ready')
  })

  it('2-colour num scales at fractions with f·2^20 not an integer paint as before', () => {
    expect(scaleColorAt(10, [num(0, RED), num(30, GREEN)], [0, 30], all).kind).toBe('ready')
  })

  it('a midpoint above the maximum (thresholds re-sorted) follows the re-sorted segment bounds', () => {
    // num 10 / num 200 / num 100: segments 10..100 (cfvo 0 and 2) and 100..200 (cfvo 2 and 1), all literal num.
    const P = [num(10, RED), num(200, YELLOW), num(100, GREEN)]
    expect(scaleColorAt(40, P, [10, 200, 100], all)).toEqual(WITHHELD)
    expect(scaleColorAt(55, P, [10, 200, 100], all).kind).toBe('ready') // f = 1/2, lower segment
  })
})
