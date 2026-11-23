// @vitest-environment node
/**
 * The bar all-equal follow-up (brief cf-bar-allequal; owner rulings 2026-11-23, docs/validation/xlsx-cf.md §14.z.19).
 *
 * Task 2a, a registration of EXISTING verified evidence: `bar:degenerate-equal-values` (stage3-bar-degenerate B, Excel
 * 86..174 = 89 of 177 px, matched `half`) paints the half-length bar on exactly the shape its capture covers: min/max
 * bounds, the x14 paired form, minLength 0 / maxLength 100, solid, no border, `context` direction on a left-to-right
 * sheet, automatic axis, two or more numeric cells all equal and positive. Owner ruling Q1: the bar announces NO
 * percentage (the value has no position in a zero-width range), so the descriptor carries no `valueFraction`. Owner
 * ruling Q2: every other all-equal shape stays withheld (`cf-degenerate-scope`), as on main.
 */
import { readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { evaluateDataBar, resolveBarThresholds } from '../src/xlsx/cf/dataBar'
import { injectCfBarAutoBounds, injectCfBarLayout, injectCfBarLengths, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { describeCfWin } from '../src/xlsx/cf/describe'
import { barPaintable } from '../src/xlsx/cf/dataBarPaint'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { measureCfCellBytes } from '../src/xlsx/view/contracts'
import type { CfRangeStatistics } from '../src/xlsx/cf/statistics'
import type { CfBarDescriptor, CfDefinition, CfOutcome } from '../src/xlsx/cf/model'

type BarDef = Extract<CfDefinition, { kind: 'dataBar' }>
const REPO_ROOT = resolve(__dirname, '..')
const BLUE = 0xff0000ff, RED = 0xffff0000

/** The evidence shape of stage3-bar-degenerate B (its saved workbook's definition). */
function evidenceShape(overrides: Partial<BarDef> = {}): BarDef {
  return {
    kind: 'dataBar', dialect: 'base', formulas: [],
    bounds: [{ type: 'min' }, { type: 'max' }],
    fillArgb: BLUE, negativeFillArgb: RED,
    minLength: 0, maxLength: 100, gradient: false, border: false,
    negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: true,
    axisPosition: 'automatic', direction: 'context', showValue: true, sheetRightToLeft: false, pairing: 'x14',
    ...overrides,
  }
}
function stats(numbers: number[]): CfRangeStatistics {
  const sorted = [...numbers].sort((a, b) => a - b)
  return {
    scopeKey: 'all-equal', numericCount: numbers.length, sum: 0, mean: undefined, min: sorted[0], max: sorted.at(-1),
    stddevPop: undefined, stddevSample: undefined, sortedAsc: new Float64Array(sorted), frequencies: new Map(),
    declaredCells: numbers.length, spanCells: numbers.length, storedCells: numbers.length,
  } as CfRangeStatistics
}
const EVIDENCE = [
  'bar:bounds-min-max', 'bar:bounds-num', 'bar:bounds-percent', 'bar:bounds-percentile', 'bar:lengths-explicit-0-100', 'bar:lengths-default-10-90',
  'bar:render-solid', 'bar:render-gradient-stops', 'bar:render-border', 'bar:axis-automatic', 'bar:axis-middle', 'bar:axis-none',
  'bar:direction-ltr', 'bar:direction-rtl', 'bar:direction-context-ltr-sheet', 'bar:negative-fill-color', 'bar:participation-types',
  'bar:degenerate-equal-values',
]
const SPAN = { origin: 'span', lengthBasis: 'cell' } as const
const evalAt = (v: number, def: BarDef, scope: number[]): CfOutcome<Omit<CfBarDescriptor, 'ruleId'>> => {
  const s = stats(scope)
  const t = resolveBarThresholds(def, s)
  return t.kind !== 'ready' ? t : evaluateDataBar(v, def, t.value, s, { variant: SPAN })
}
const DEGENERATE = { kind: 'unresolved', reason: 'cf-degenerate-scope' }

beforeEach(() => { injectCfBarLayout({}); injectCfBarLengths({}); injectCfBarAutoBounds({}) })
afterEach(() => { injectCfBarLayout(null); injectCfBarLengths(null); injectCfBarAutoBounds(null) })

describe('2a: bar:degenerate-equal-values paints its verified half length on the evidence shape only', () => {
  it('the evidence shape: a half-length bar (0..0.5 of the span) with no announced position', () => {
    withExactCfEvidence(EVIDENCE, () => {
      const out = evalAt(42, evidenceShape(), [42, 42, 42, 42, 42])
      expect(out.kind, JSON.stringify(out)).toBe('ready')
      const bar = (out as { value: Omit<CfBarDescriptor, 'ruleId'> }).value
      expect([bar.startFraction, bar.endFraction, bar.axisFraction, bar.direction, bar.isNegative, bar.isGradient]).toEqual([0, 0.5, 0, 'leftToRight', false, false])
      expect('valueFraction' in bar).toBe(false)
      expect(bar.axisColor).toBeUndefined()
    })
  })

  it('without bar:degenerate-equal-values the shape stays withheld, as on main', () => {
    withExactCfEvidence(EVIDENCE.filter(c => c !== 'bar:degenerate-equal-values'), () => {
      expect(evalAt(42, evidenceShape(), [42, 42, 42])).toEqual(DEGENERATE)
    })
  })

  it('every shape outside the evidence stays withheld (owner ruling Q2)', () => {
    withExactCfEvidence(EVIDENCE, () => {
      const outside: Array<[string, BarDef, number[]]> = [
        ['explicit leftToRight', evidenceShape({ direction: 'leftToRight' }), [42, 42, 42]],
        ['rightToLeft', evidenceShape({ direction: 'rightToLeft' }), [42, 42, 42]],
        ['context on a right-to-left sheet', evidenceShape({ sheetRightToLeft: true }), [42, 42, 42]],
        ['10/90 lengths', evidenceShape({ minLength: 10, maxLength: 90 }), [42, 42, 42]],
        ['gradient', evidenceShape({ gradient: true }), [42, 42, 42]],
        ['border', evidenceShape({ border: true, borderArgb: RED }), [42, 42, 42]],
        ['middle axis', evidenceShape({ axisPosition: 'middle', axisArgb: RED }), [42, 42, 42]],
        ['no axis', evidenceShape({ axisPosition: 'none' }), [42, 42, 42]],
        ['base-only pairing', evidenceShape({ pairing: 'base-only' }), [42, 42, 42]],
        ['unknown pairing', evidenceShape({ pairing: undefined }), [42, 42, 42]],
        ['percent bounds', evidenceShape({ bounds: [{ type: 'percent', value: 0 }, { type: 'percent', value: 100 }] }), [42, 42, 42]],
        ['percentile bounds', evidenceShape({ bounds: [{ type: 'percentile', value: 10 }, { type: 'percentile', value: 90 }] }), [42, 42, 42]],
        ['min / num', evidenceShape({ bounds: [{ type: 'min' }, { type: 'num', value: 42 }] }), [42, 42, 42]],
        ['all zero', evidenceShape(), [0, 0, 0]],
        ['all negative', evidenceShape(), [-5, -5, -5]],
        ['single numeric cell (2c)', evidenceShape(), [42]],
      ]
      for (const [what, def, scope] of outside) {
        const out = evalAt(scope[0]!, def, scope)
        expect(out.kind, what).toBe('unresolved')
      }
      // Distinct literal bounds over all-equal data (stage3-bar-geometry H and I) and num 0/0 over distinct data stay withheld.
      expect(evalAt(100, evidenceShape({ bounds: [{ type: 'num', value: 0 }, { type: 'num', value: 100 }] }), [100, 100, 100])).toEqual(DEGENERATE)
      expect(evalAt(5, evidenceShape({ bounds: [{ type: 'num', value: 0 }, { type: 'num', value: 0 }] }), [5, 10, 20]).kind).toBe('unresolved')
      // Direct evaluation against equal thresholds without the scope's count stays withheld (defence in depth).
      expect(evaluateDataBar(42, evidenceShape(), [42, 42], { min: 42, max: 42 }, { variant: SPAN })).toEqual(DEGENERATE)
    })
  })

  it('a bar with no valueFraction paints, announces nothing about its position, and is priced exactly on the wire', () => {
    const bar: CfBarDescriptor = { ruleId: 1, axisFraction: 0, startFraction: 0, endFraction: 0.5, fillColor: BLUE, isGradient: false, isNegative: false, showValue: true, direction: 'leftToRight' }
    expect(barPaintable(bar, undefined, () => true)).toBe(true)
    // A present but invalid valueFraction still refuses (fail closed).
    expect(barPaintable({ ...bar, valueFraction: Number.NaN }, undefined, () => true)).toBe(false)
    expect(describeCfWin({}, undefined, undefined, 'en', undefined, {}, undefined)).toBeUndefined()
    expect(describeCfWin({}, undefined, undefined, 'en', undefined, { fraction: 0.5 }, undefined)?.text).toBe('Data bar, 50% of range')
    const wire = { ruleId: 1, axisFraction: 0, startFraction: 0, endFraction: 0.5, fillColor: BLUE, isGradient: false, isNegative: false, showValue: true, rtl: false }
    expect(measureCfCellBytes({ properties: {}, bar: { ...wire, valueFraction: 0.5 } }) - measureCfCellBytes({ properties: {}, bar: wire })).toBe(16 + 8)
  })

  it('production on Excel\'s saved stage3-bar-degenerate: B2:B6 paint half, C2 (single cell) stays withheld', async () => {
    // The published exports, not the evaluator tests' empty ones.
    injectCfBarLayout(null); injectCfBarLengths(null); injectCfBarAutoBounds(null)
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/stage3-bar-degenerate.xlsx')))))
    const batch = (await prepareCfPaint(doc.sheets[0]!, { evidence: cfRuntimeEvidence }))!
    try {
      for (let row = 1; row <= 5; row++) {
        const bar = batch.barAt(row, 1)
        expect(bar, `B${row + 1}`).toBeDefined()
        expect([bar!.startFraction, bar!.endFraction, bar!.direction, bar!.valueFraction]).toEqual([0, 0.5, 'leftToRight', undefined])
      }
      expect(batch.barAt(1, 2)).toBeUndefined()
    } finally { batch.dispose() }
  }, 300_000)
})
