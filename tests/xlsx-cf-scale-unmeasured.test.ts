/**
 * Final review I3 / M2 / M6 (Ruling F-2): a scale shape that no committed Excel observation covers must be WITHHELD,
 * never painted with a computed colour, even when every coarse capability id is verified. The probes below
 * run under the production evidence (every verified id allowed).
 */
import { describe, expect, it } from 'vitest'
import { composeCell } from '../src/xlsx/cf/compose'
import { resolveScaleThresholds, scaleColorAt } from '../src/xlsx/cf/colorScale'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfRangeStatistics } from '../src/xlsx/cf/statistics'
import type { CfRule, CfScalePoint } from '../src/xlsx/cf/model'

function stats(numbers: number[]): CfRangeStatistics {
  const sorted = [...numbers].sort((a, b) => a - b)
  return {
    scopeKey: 'k', numericCount: numbers.length, sum: numbers.reduce((a, b) => a + b, 0), mean: undefined,
    min: sorted[0], max: sorted[sorted.length - 1], stddevPop: undefined, stddevSample: undefined,
    sortedAsc: new Float64Array(sorted), frequencies: new Map(), declaredCells: numbers.length, spanCells: numbers.length, storedCells: numbers.length,
  }
}
const C = [0xfff8696b, 0xffffeb84, 0xff63be7b]
const two = (a: CfScalePoint['bound'], b: CfScalePoint['bound']): CfScalePoint[] => [{ bound: a, argb: C[0]! }, { bound: b, argb: C[2]! }]
const three = (a: CfScalePoint['bound'], b: CfScalePoint['bound'], c: CfScalePoint['bound']): CfScalePoint[] => [
  { bound: a, argb: C[0]! }, { bound: b, argb: C[1]! }, { bound: c, argb: C[2]! },
]
const num = (value: number) => ({ type: 'num', value }) as const
const production = (cap: string) => cfRuntimeEvidence.allows(cap)

describe('I3 unmeasured scale shapes withhold under production evidence', () => {
  it('sanity: the production evidence allows every capability the probes lean on', () => {
    for (const cap of ['scale:degenerate-inverted-thresholds', 'scale:degenerate-equal-values', 'scale:thresholds-percent', 'scale:composition-equal-priority', 'scale:thresholds-percent-percentile-out-of-bounds']) {
      expect(production(cap), cap).toBe(true)
    }
  })

  it('A: a 3-colour inverted scale (num 100 / num 50 / num 200) is withheld', () => {
    const out = resolveScaleThresholds(three(num(100), num(50), num(200)), stats([10, 50, 75, 100, 200]), { hasEvidence: production })
    expect(out.kind).toBe('unresolved')
  })

  it('A: every 3-colour ordering other than the two measured ones is withheld', () => {
    const data = stats([10, 50, 100, 200])
    for (const [a, b, c] of [[100, 50, 200], [50, 10, 100], [200, 100, 10], [100, 200, 50], [50, 50, 100], [50, 100, 100], [50, 100, 50], [70, 70, 70]] as const) {
      expect(resolveScaleThresholds(three(num(a), num(b), num(c)), data, { hasEvidence: production }).kind, `${a}/${b}/${c}`).toBe('unresolved')
    }
    // measured: ascending, and midpoint above max (min / num 200 / num 100 over 10..200)
    expect(resolveScaleThresholds(three(num(10), num(50), num(100)), data, { hasEvidence: production }).kind).toBe('ready')
    expect(resolveScaleThresholds(three({ type: 'min' }, num(200), num(100)), data, { hasEvidence: production }).kind).toBe('ready')
  })

  it('A: scaleColorAt itself refuses a 3-point inversion even when handed the thresholds', () => {
    const out = withExactCfEvidence(['scale:interpolation-color-space', 'scale:interpolation-rounding-ladder', 'scale:degenerate-values-beyond-bounds'], () =>
      scaleColorAt(75, three(num(100), num(50), num(200)), [100, 50, 200]))
    expect(out.kind).toBe('unresolved')
  })

  it('A: equal thresholds (num 50 / num 50) over DISTINCT data is withheld; all-equal data stays measured', () => {
    expect(resolveScaleThresholds(two(num(50), num(50)), stats([10, 20, 90]), { hasEvidence: production }).kind).toBe('unresolved')
    expect(resolveScaleThresholds(two({ type: 'min' }, { type: 'max' }), stats([7, 7, 7]), { hasEvidence: production }).kind).toBe('ready')
    expect(resolveScaleThresholds(two({ type: 'min' }, { type: 'max' }), stats([7]), { hasEvidence: production }).kind).toBe('ready')
  })

  it('A: the measured 2-colour inversion (num 100 -> num 10) stays resolvable', () => {
    expect(resolveScaleThresholds(two(num(100), num(10)), stats([10, 40, 70, 100]), { hasEvidence: production }).kind).toBe('ready')
  })

  it('B: an out-of-range percent or percentile alone is withheld', () => {
    const data = stats([0, 25, 50, 75, 100])
    expect(resolveScaleThresholds(two({ type: 'percent', value: -10 }, { type: 'max' }), data, { hasEvidence: production }).kind).toBe('unresolved')
    expect(resolveScaleThresholds(two({ type: 'min' }, { type: 'percent', value: 150 }), data, { hasEvidence: production }).kind).toBe('unresolved')
    expect(resolveScaleThresholds(two({ type: 'min' }, { type: 'percentile', value: 150 }), data, { hasEvidence: production }).kind).toBe('unresolved')
    expect(resolveScaleThresholds(two({ type: 'percent', value: -10 }, { type: 'percentile', value: 150 }), data, { hasEvidence: production }).kind).toBe('unresolved')
    expect(resolveScaleThresholds(two({ type: 'percent', value: 10 }, { type: 'percent', value: 90 }), data, { hasEvidence: production }).kind).toBe('ready')
  })
})

describe('I3-C / M2 composition withholds what was not measured', () => {
  const all = { hasEvidence: () => true }
  const dxfFill = (documentOrder: number, priority: number): CfRule => ({
    id: 1, priority, documentOrder, ranges: [{ firstRow: 0, lastRow: 3, firstCol: 0, lastCol: 0 }], stopIfTrue: false,
    writes: new Set(['fill']), dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE' },
    definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
  })
  const scale = (documentOrder: number, priority: number): CfRule => ({
    id: 2, priority, documentOrder, ranges: [{ firstRow: 0, lastRow: 3, firstCol: 0, lastCol: 0 }], stopIfTrue: false,
    writes: new Set(['colorScale', 'fill']), definition: { kind: 'colorScale', points: two({ type: 'min' }, { type: 'max' }) as [CfScalePoint, CfScalePoint], dialect: 'base', formulas: [] },
  })
  const yes = { kind: 'ready', value: true } as const
  const scaled = { kind: 'ready', value: { argb: 0xff63be7b } } as const

  it('C: equal priority with the SCALE first in document order is withheld (only DXF-first was measured)', () => {
    const state = composeCell([scale(0, 5), dxfFill(1, 5)], [scaled, yes], all)
    expect(state.visual?.colorScale).toBeUndefined()
    expect(state.properties['fill.bg']).toBeUndefined()
    expect(state.unresolvedWrites?.has('fill')).toBe(true)
  })

  it('C: equal priority with the DXF first (the measured order) still lets the scale win', () => {
    const state = composeCell([dxfFill(0, 5), scale(1, 5)], [yes, scaled], all)
    expect(state.visual?.colorScale?.argb).toBe(0xff63be7b)
    expect(state.properties['fill.bg']).toBeUndefined()
  })

  it('M2: a bare true scale outcome has no evaluated colour and is withheld, not painted as endpoint 0', () => {
    const state = composeCell([scale(0, 1)], [yes], all)
    expect(state.visual?.colorScale).toBeUndefined()
    expect(state.unresolvedWrites?.has('fill')).toBe(true)
  })
})

describe('M6: a scale positive does not switch off the Stage 1 multi-positive guards for later rules', () => {
  it('equal-priority DXF rules after a scale are still refused when priority:equal-priority-first is unevidenced', async () => {
    const { prepareCfPaint } = await import('../src/xlsx/cf/render')
    const { createCfMemory } = await import('../src/xlsx/cf/limits')
    const range = [{ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }]
    const sheet = {
      name: 'M6', sheetId: 'm6', cols: [], merges: [], mergeRanges: [],
      rows: [0, 1, 2].map(r => ({ index: r, cells: [{ row: r, col: 0, ref: `A${r + 1}`, styleIndex: 0, value: r + 1 }] })),
      conditionalFormatting: { diagnostics: [], rules: [
        { id: 1, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: range, writes: new Set(['colorScale', 'fill']),
          definition: { kind: 'colorScale', points: two({ type: 'min' }, { type: 'max' }), dialect: 'base', formulas: [] } },
        { id: 2, priority: 2, documentOrder: 1, stopIfTrue: false, ranges: range, writes: new Set(['font.bold']), dxf: { 'font.bold': true },
          definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] } },
        { id: 3, priority: 2, documentOrder: 2, stopIfTrue: false, ranges: range, writes: new Set(['font.italic']), dxf: { 'font.italic': true },
          definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] } },
      ] },
    } as never
    const evidence = { allows: (cap: string) => cap !== 'priority:equal-priority-first' }
    const memory = createCfMemory()
    const batch = await prepareCfPaint(sheet, { memory, evidence })
    const style = batch?.layer.styleAt(0, 0, undefined)
    expect(style?.fillColor, 'the equal-priority DXF tie is unevidenced: nothing of this cell may resolve').toBeUndefined()
    expect(style?.bold).toBeUndefined()
    batch?.dispose()
  })
})

describe('bot P2: finite thresholds whose span overflows interpolate correctly, never a wrong colour', () => {
  const ev = ['scale:interpolation-color-space', 'scale:interpolation-rounding-ladder', 'scale:degenerate-values-beyond-bounds']
  const at = (v: number) => withExactCfEvidence(ev, () => scaleColorAt(v, two(num(-1e308), num(1e308)), [-1e308, 1e308]))
  it('2-point [-1e308, 1e308]: endpoints are the endpoint colours, the midpoint is a blend', () => {
    expect(at(-1e308)).toEqual({ kind: 'ready', value: C[0] })
    expect(at(1e308)).toEqual({ kind: 'ready', value: C[2] })
    const mid = at(0)
    expect(mid.kind).toBe('ready')
    expect(mid.kind === 'ready' && mid.value).not.toBe(C[0])
    expect(mid.kind === 'ready' && mid.value).not.toBe(C[2])
    expect(mid.kind === 'ready' && mid.value).not.toBe(0)
  })
  it('3-point with an overflowing outer span keeps its segments', () => {
    const out = (v: number) => withExactCfEvidence(ev, () => scaleColorAt(v, three(num(-1e308), num(0), num(1e308)), [-1e308, 0, 1e308]))
    expect(out(-1e308)).toEqual({ kind: 'ready', value: C[0] })
    expect(out(0)).toEqual({ kind: 'ready', value: C[1] })
    expect(out(1e308)).toEqual({ kind: 'ready', value: C[2] })
  })
  it('ordinary spans still interpolate', () => {
    expect(withExactCfEvidence(ev, () => scaleColorAt(5, two(num(0), num(10)), [0, 10])).kind).toBe('ready')
  })
})
