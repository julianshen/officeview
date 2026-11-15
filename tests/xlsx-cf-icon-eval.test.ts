import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import type { CfCellState, CfDefinition, CfIconThreshold, CfOutcome, CfRule, CfStamp, SourceRect } from '../src/xlsx/cf/model'
import type { XlsxCell } from '../src/xlsx/types'
import { injectCfIconSelections, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { percentileOrderStatistic } from '../src/xlsx/cf/colorScale'
import { evaluateIconSet, iconBucket, iconEvaluationRequiredCapabilities, resolveIconThresholds } from '../src/xlsx/cf/iconSet'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { createCfService } from '../src/xlsx/cf/service'
import { createResidentCfSource, type CfRangeStatistics } from '../src/xlsx/cf/statistics'

// Stage 4 Task 26: icon thresholds, buckets and participation. Every bucket below is the S4-D/S4-E HYPOTHESIS run under
// injected evidence (withExactCfEvidence, injectCfIconSelections), never an Excel fact: the facts come from Task 24's
// sidecars, and production evidence stays closed.

type IconDef = Extract<CfDefinition, { kind: 'iconSet' }>
const th = (type: 'num' | 'percent' | 'percentile', value: number, gte = true): CfIconThreshold => ({ bound: { type, value }, gte })
const PCT3 = [th('percent', 0), th('percent', 33), th('percent', 67)]
const NUM3 = [th('num', 0), th('num', 20), th('num', 40)]
const iconDef = (over: Partial<IconDef> = {}): IconDef => ({
  kind: 'iconSet', dialect: 'base', formulas: [], set: '3Arrows', thresholds: PCT3, reverse: false, showValue: true,
  percentAttr: true, sheetRightToLeft: false, ...over,
})
function statsOf(values: readonly number[]): CfRangeStatistics {
  const sorted = Float64Array.from([...values].sort((a, b) => a - b))
  const sum = values.reduce((a, b) => a + b, 0)
  return {
    scopeKey: 't', numericCount: values.length, sum, mean: values.length ? sum / values.length : undefined,
    min: values.length ? sorted[0] : undefined, max: values.length ? sorted[sorted.length - 1] : undefined,
    stddevPop: undefined, sortedAsc: sorted, frequencies: new Map(), declaredCells: values.length, spanCells: values.length, storedCells: values.length,
  }
}
const range = (lo: number, hi: number): number[] => Array.from({ length: hi - lo + 1 }, (_, i) => lo + i)
const num = (value: number) => ({ kind: 'number' as const, value })

/** Every evaluation-layer gate of a built-in or custom set (stubs only). */
const EVAL = [
  'icon:families-base', 'icon:families-x14', 'icon:thresholds-num', 'icon:thresholds-percent', 'icon:thresholds-percentile-inc-exc',
  'icon:thresholds-gte-true', 'icon:thresholds-gte-false', 'icon:thresholds-mixed', 'icon:reverse', 'icon:reverse-custom', 'icon:participation-types',
]
afterEach(() => { injectCfIconSelections(null) })

function bucketOf(out: CfOutcome<{ bucket: number }>): number | string {
  return out.kind === 'ready' ? out.value.bucket : out.kind === 'unresolved' ? out.reason : 'no-match'
}

describe('iconEvaluationRequiredCapabilities names each feature', () => {
  it('names the family, threshold types, gte, reverse and percentile ids', () => {
    expect(iconEvaluationRequiredCapabilities(iconDef())).toEqual(new Set(['icon:families-base', 'icon:thresholds-percent']))
    expect(iconEvaluationRequiredCapabilities(iconDef({ set: '3Stars' })).has('icon:families-x14')).toBe(true)
    expect(iconEvaluationRequiredCapabilities(iconDef({ set: '3Stars' })).has('icon:families-base')).toBe(false)
    expect(iconEvaluationRequiredCapabilities(iconDef({ thresholds: [PCT3[0]!, th('percent', 33, false), PCT3[2]!] })).has('icon:thresholds-gte-false')).toBe(true)
    expect(iconEvaluationRequiredCapabilities(iconDef({ reverse: true })).has('icon:reverse')).toBe(true)
    expect(iconEvaluationRequiredCapabilities(iconDef({ reverse: true })).has('icon:reverse-custom')).toBe(false)
    const custom = iconDef({ custom: [{ set: '3Flags', iconId: 0 }, { set: 'NoIcons', iconId: 0 }, { set: '5Boxes', iconId: 4 }] })
    const customReverse = iconEvaluationRequiredCapabilities({ ...custom, reverse: true })
    expect(customReverse.has('icon:reverse') && customReverse.has('icon:reverse-custom')).toBe(true)
    expect(iconEvaluationRequiredCapabilities(iconDef({ thresholds: [th('percentile', 0), th('percentile', 33), th('percentile', 67)] })).has('icon:thresholds-percentile-inc-exc')).toBe(true)
    // A custom set needs no icon:families-* id.
    expect([...iconEvaluationRequiredCapabilities(custom)].some(c => c.startsWith('icon:families-'))).toBe(false)
    expect(iconEvaluationRequiredCapabilities(iconDef({ thresholds: NUM3 })).has('icon:thresholds-num')).toBe(true)
    // Out-of-domain percent and percentile values are their own shape (S4-D).
    expect(iconEvaluationRequiredCapabilities(iconDef({ thresholds: [PCT3[0]!, th('percent', 150), PCT3[2]!] })).has('icon:bounds-out-of-domain')).toBe(true)
  })

  it('evaluation still requires gte-true for a default gte and thresholds-mixed for a mixed list (S4-D)', () => {
    const stats = statsOf(range(0, 100))
    const allBut = (cap: string) => EVAL.filter(c => c !== cap)
    withExactCfEvidence(allBut('icon:thresholds-gte-true'), () => {
      const t = resolveIconThresholds(iconDef(), stats)
      expect(t).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
      expect(evaluateIconSet(num(50), iconDef(), [0, 33, 67])).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
    })
    const mixed = iconDef({ thresholds: [th('percent', 0), th('num', 20), th('percent', 67)] })
    withExactCfEvidence(allBut('icon:thresholds-mixed'), () => {
      expect(resolveIconThresholds(mixed, stats)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
    })
    withExactCfEvidence(EVAL, () => {
      expect(resolveIconThresholds(mixed, stats)).toEqual({ kind: 'ready', value: [0, 20, 67] })
    })
  })
})

describe('no threshold evidence ⇒ unresolved', () => {
  it('every icon evaluates to cf-unmeasured-capability and the cell state carries no iconSet', async () => {
    const defs = [iconDef(), iconDef({ thresholds: NUM3 }), iconDef({ set: '3Stars' }), iconDef({ custom: [{ set: '3Flags', iconId: 0 }, { set: 'NoIcons', iconId: 0 }, { set: '3Flags', iconId: 2 }] })]
    await withExactCfEvidence([], async () => {
      for (const def of defs) {
        expect(resolveIconThresholds(def, statsOf(range(0, 100)))).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
        expect(evaluateIconSet(num(50), def, [0, 33, 67])).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
      }
      const run = await serviceRun([iconRule(1, iconDef())], range(0, 9).map(v => cellAt(v, v)), [])
      for (let r = 0; r < 10; r++) {
        const st = run.at(r)
        expect(st.kind).toBe('ready')
        if (st.kind === 'ready') { expect(st.value.visual?.iconSet).toBeUndefined(); expect(st.value.unresolvedWrites?.has('iconSet')).toBe(true) }
      }
    })
  })
})

describe('bucket boundaries under the S4-D hypothesis', () => {
  it('num 20/40 over {0..100}', () => {
    withExactCfEvidence(EVAL, () => {
      const def = iconDef({ thresholds: NUM3 })
      const t = resolveIconThresholds(def, statsOf(range(0, 100)))
      expect(t).toEqual({ kind: 'ready', value: [0, 20, 40] })
      const at = (v: number, d = def) => bucketOf(evaluateIconSet(num(v), d, [0, 20, 40]))
      expect([19, 20, 39, 40, 100].map(v => at(v))).toEqual([0, 1, 1, 2, 2])
      const gteFalse = iconDef({ thresholds: [NUM3[0]!, th('num', 20, false), NUM3[2]!] })
      expect(at(20, gteFalse)).toBe(0)
      expect(at(21, gteFalse)).toBe(1)
    })
    expect(iconBucket(20, [0, 20, 40], [true, false, true])).toBe(0)
    expect(iconBucket(40, [0, 20, 40], [true, true, false])).toBe(1)
    expect(iconBucket(-5, [0, 20, 40], [true, true, true])).toBe(0)
  })

  it('a value below a non-trivial floor is unresolved until icon:thresholds-floor selects a reading (S4-C)', () => {
    const def = iconDef({ thresholds: [th('num', 10), th('num', 20), th('num', 40)] })
    withExactCfEvidence(EVAL, () => {
      expect(bucketOf(evaluateIconSet(num(5), def, [10, 20, 40]))).toBe('cf-unmeasured-capability')
      expect(bucketOf(evaluateIconSet(num(10), def, [10, 20, 40]))).toBe(0)
      // The bare id never picks a reading: it names two (bucket 0, no icon).
      expect(bucketOf(evaluateIconSet(num(5), def, [10, 20, 40]))).toBe('cf-unmeasured-capability')
    })
    withExactCfEvidence([...EVAL, 'icon:thresholds-floor'], () => {
      expect(bucketOf(evaluateIconSet(num(5), def, [10, 20, 40]))).toBe('cf-unmeasured-capability')
      injectCfIconSelections({ 'icon:thresholds-floor': 'below-floor-bucket0' })
      expect(bucketOf(evaluateIconSet(num(5), def, [10, 20, 40]))).toBe(0)
      injectCfIconSelections({ 'icon:thresholds-floor': 'below-floor-no-icon' })
      expect(evaluateIconSet(num(5), def, [10, 20, 40])).toEqual({ kind: 'no-match' })
    })
    // A selection without its capability lifts nothing.
    withExactCfEvidence(EVAL, () => {
      injectCfIconSelections({ 'icon:thresholds-floor': 'below-floor-bucket0' })
      expect(bucketOf(evaluateIconSet(num(5), def, [10, 20, 40]))).toBe('cf-unmeasured-capability')
    })
  })
})

describe('percent and percentile thresholds reuse the shared primitives', () => {
  it('percent 25/50/75 resolves exactly through lerp', () => {
    const def = iconDef({ set: '4Arrows', thresholds: [th('percent', 0), th('percent', 25), th('percent', 50), th('percent', 75)] })
    withExactCfEvidence(EVAL, () => {
      expect(resolveIconThresholds(def, statsOf(range(0, 100)))).toEqual({ kind: 'ready', value: [0, 25, 50, 75] })
      expect(resolveIconThresholds(def, statsOf(range(10, 110)))).toEqual({ kind: 'ready', value: [10, 35, 60, 85] })
    })
  })

  it('percentile 33/67 equals percentileOrderStatistic under the selected variant only', () => {
    const def = iconDef({ thresholds: [th('percentile', 0), th('percentile', 33), th('percentile', 67)] })
    const data = [1, 2, 3, 4, 10]
    withExactCfEvidence(EVAL, () => {
      // The bare capability names two readings (inc, exc): without a selection nothing resolves.
      expect(resolveIconThresholds(def, statsOf(data))).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
      injectCfIconSelections({ 'icon:thresholds-percentile-inc-exc': 'inc' })
      const out = resolveIconThresholds(def, statsOf(data))
      expect(out.kind).toBe('ready')
      if (out.kind !== 'ready') return
      const sorted = Float64Array.from(data)
      expect(out.value).toEqual([0, 33, 67].map(p => percentileOrderStatistic(sorted, p, 'inc')))
      expect(out.value[0]).toBe(1)
      expect(Math.abs(out.value[1]! - 2.32)).toBeLessThan(1e-12)
      expect(Math.abs(out.value[2]! - 3.68)).toBeLessThan(1e-12)
      injectCfIconSelections({ 'icon:thresholds-percentile-inc-exc': 'exc' })
      // EXC is undefined at p = 0, so the hypothesis run is unresolved, never a substitute value.
      expect(resolveIconThresholds(def, statsOf(data)).kind).toBe('unresolved')
      injectCfIconSelections({ 'icon:thresholds-percentile-inc-exc': 'median' })
      expect(resolveIconThresholds(def, statsOf(data))).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
    })
  })

  it('iconSet.ts contains no sort and no percentile implementation', () => {
    const src = readFileSync(resolve(process.cwd(), 'src/xlsx/cf/iconSet.ts'), 'utf8')
    expect(src).not.toMatch(/sort\(/)
    expect(src).not.toMatch(/rank\s*=|Math\.floor\(/)
    expect(src).toMatch(/percentileOrderStatistic\(/)
    expect(src).toMatch(/lerp\(/)
  })
})

describe('reverse and custom mapping', () => {
  it('reverse maps buckets to mirrored ids; bucket and buckets stay semantic', () => {
    withExactCfEvidence(EVAL, () => {
      const def = iconDef({ thresholds: NUM3, reverse: true })
      const ids = [5, 25, 45].map(v => evaluateIconSet(num(v), def, [0, 20, 40]))
      expect(ids.map(o => o.kind === 'ready' ? [o.value.iconId, o.value.bucket, o.value.buckets, o.value.set] : o)).toEqual([[2, 0, 3, '3Arrows'], [1, 1, 3, '3Arrows'], [0, 2, 3, '3Arrows']])
      const plain = [5, 25, 45].map(v => evaluateIconSet(num(v), iconDef({ thresholds: NUM3 }), [0, 20, 40]))
      expect(plain.map(o => o.kind === 'ready' ? o.value.iconId : o)).toEqual([0, 1, 2])
    })
    withExactCfEvidence(EVAL.filter(c => c !== 'icon:reverse'), () => {
      expect(bucketOf(evaluateIconSet(num(5), iconDef({ thresholds: NUM3, reverse: true }), [0, 20, 40]))).toBe('cf-unmeasured-capability')
    })
  })

  it('a custom mapping returns its (set, iconId) per bucket, NoIcons as id 0', () => {
    const custom = [{ set: '3Flags', iconId: 2 }, { set: 'NoIcons', iconId: 0 }, { set: '5Boxes', iconId: 4 }] as const
    const def = iconDef({ thresholds: NUM3, custom })
    withExactCfEvidence(EVAL, () => {
      const out = [5, 25, 45].map(v => evaluateIconSet(num(v), def, [0, 20, 40]))
      expect(out).toEqual([
        { kind: 'ready', value: { set: '3Flags', iconId: 2, bucket: 0, buckets: 3, showValue: true } },
        { kind: 'ready', value: { set: 'NoIcons', iconId: 0, bucket: 1, buckets: 3, showValue: true } },
        { kind: 'ready', value: { set: '5Boxes', iconId: 4, bucket: 2, buckets: 3, showValue: true } },
      ])
    })
    // Custom + reverse: unresolved without icon:reverse-custom, and with it only under a selected reading.
    withExactCfEvidence(EVAL.filter(c => c !== 'icon:reverse-custom'), () => {
      expect(bucketOf(evaluateIconSet(num(5), { ...def, reverse: true }, [0, 20, 40]))).toBe('cf-unmeasured-capability')
    })
    withExactCfEvidence(EVAL, () => {
      expect(bucketOf(evaluateIconSet(num(5), { ...def, reverse: true }, [0, 20, 40]))).toBe('cf-unmeasured-capability')
      injectCfIconSelections({ 'icon:reverse-custom': 'reverse-then-map' })
      const r = evaluateIconSet(num(5), { ...def, reverse: true }, [0, 20, 40])
      expect(r).toEqual({ kind: 'ready', value: { set: '5Boxes', iconId: 4, bucket: 0, buckets: 3, showValue: true } })
      injectCfIconSelections({ 'icon:reverse-custom': 'ignored' })
      expect(evaluateIconSet(num(5), { ...def, reverse: true }, [0, 20, 40])).toEqual({ kind: 'ready', value: { set: '3Flags', iconId: 2, bucket: 0, buckets: 3, showValue: true } })
      injectCfIconSelections({ 'icon:reverse-custom': 'map-then-mirror' })
      expect(evaluateIconSet(num(5), { ...def, reverse: true }, [0, 20, 40])).toEqual({ kind: 'ready', value: { set: '3Flags', iconId: 0, bucket: 0, buckets: 3, showValue: true } })
      expect(evaluateIconSet(num(25), { ...def, reverse: true }, [0, 20, 40])).toEqual({ kind: 'ready', value: { set: 'NoIcons', iconId: 0, bucket: 1, buckets: 3, showValue: true } })
    })
  })
})

describe('icon evaluation stays finite at ±1e308 (Lesson 6)', () => {
  it('percent 50 over {-1e308, 1e308} resolves finite through lerp; every field is finite', () => {
    withExactCfEvidence(EVAL, () => {
      const def = iconDef({ thresholds: [th('percent', 0), th('percent', 50), th('percent', 100)] })
      const t = resolveIconThresholds(def, statsOf([-1e308, 1e308]))
      expect(t.kind).toBe('ready')
      if (t.kind !== 'ready') return
      expect(t.value.every(Number.isFinite)).toBe(true)
      expect(t.value[1]).toBe(0)
      for (const v of [-1e308, -1, 0, 1, 1e308]) {
        const o = evaluateIconSet(num(v), def, t.value)
        expect(o.kind).toBe('ready')
        if (o.kind === 'ready') for (const n of [o.value.bucket, o.value.buckets, o.value.iconId]) expect(Number.isFinite(n)).toBe(true)
      }
      expect(bucketOf(evaluateIconSet(num(Infinity), def, t.value))).toBe('cf-non-finite-value')
      expect(bucketOf(evaluateIconSet(num(5), def, [0, NaN, 1]))).toBe('cf-non-finite-value')
    })
  })

  it('equal values, a single numeric cell, non-monotonic and out-of-domain thresholds are unresolved, never NaN', () => {
    withExactCfEvidence(EVAL, () => {
      expect(resolveIconThresholds(iconDef(), statsOf([7, 7, 7]))).toEqual({ kind: 'unresolved', reason: 'cf-degenerate-scope' })
      expect(resolveIconThresholds(iconDef(), statsOf([7]))).toEqual({ kind: 'unresolved', reason: 'cf-degenerate-scope' })
      expect(resolveIconThresholds(iconDef(), statsOf([]))).toEqual({ kind: 'unresolved', reason: 'cf-no-numeric-values' })
      const inverted = iconDef({ thresholds: [th('num', 0), th('num', 40), th('num', 20)] })
      expect(resolveIconThresholds(inverted, statsOf(range(0, 100)))).toEqual({ kind: 'unresolved', reason: 'cf-degenerate-scope' })
      const outOfDomain = iconDef({ thresholds: [th('percent', 0), th('percent', 150), th('percent', 67)] })
      expect(resolveIconThresholds(outOfDomain, statsOf(range(0, 100)))).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
    })
    // A capability alone never picks a degenerate reading; with its selection 'as-model' the hypothesis resolves.
    withExactCfEvidence([...EVAL, 'icon:degenerate-equal-values', 'icon:degenerate-non-monotonic'], () => {
      expect(resolveIconThresholds(iconDef(), statsOf([7, 7, 7])).kind).toBe('unresolved')
      injectCfIconSelections({ 'icon:degenerate-equal-values': 'as-model', 'icon:degenerate-non-monotonic': 'as-model' })
      expect(resolveIconThresholds(iconDef(), statsOf([7, 7, 7]))).toEqual({ kind: 'ready', value: [7, 7, 7] })
      const t = resolveIconThresholds(iconDef({ thresholds: [th('num', 0), th('num', 40), th('num', 20)] }), statsOf(range(0, 100)))
      expect(t).toEqual({ kind: 'ready', value: [0, 40, 20] })
      injectCfIconSelections({ 'icon:degenerate-equal-values': 'lowest' })
      expect(resolveIconThresholds(iconDef(), statsOf([7, 7, 7])).kind).toBe('unresolved')
    })
  })
})

/* ---------------------------------------------------------------------------------------------------------------- */
/* Service-level participation and scope gates (S4-G, S4-H).                                                          */
/* ---------------------------------------------------------------------------------------------------------------- */

const SHEET = 'icons'
const stamp: CfStamp = { owner: 'icon', sheetId: SHEET, workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1 }
const cellAt = (row: number, value: XlsxCell['value'], extra?: Partial<XlsxCell>): XlsxCell => ({ ref: `A${row + 1}`, row, col: 0, value, styleIndex: 0, ...extra })
const R = (firstRow: number, lastRow: number, firstCol = 0, lastCol = 0): SourceRect => ({ firstRow, lastRow, firstCol, lastCol })
function iconRule(id: number, definition: IconDef, ranges: readonly SourceRect[] = [R(0, 9)], extra: Partial<CfRule> = {}): CfRule {
  return { id, priority: id, documentOrder: id, ranges, stopIfTrue: false, writes: new Set(['iconSet']), definition, ...extra }
}
async function serviceRun(rules: CfRule[], cells: XlsxCell[], caps: readonly string[] | null, options: { merges?: SourceRect[]; hiddenRows?: SourceRect[]; numberFormats?: 'unknown'; rows?: number } = {}) {
  const memory = createCfMemory({})
  const set = caps === null ? undefined : new Set(caps)
  const service = createCfService({ rules, diagnostics: [] }, memory, {
    ...(set ? { evidence: { allows: (cap: string) => set.has(cap) } } : {}),
    ...(options.merges ? { merges: options.merges } : {}),
    ...(options.hiddenRows ? { statisticsOptions: { hiddenRows: options.hiddenRows } } : {}),
  })
  const inner = createResidentCfSource(cells, SHEET, stamp, memory)
  const source = options.numberFormats === 'unknown' ? { bounds: inner.bounds, stamp: inner.stamp, pages: inner.pages } : inner
  const rows = options.rows ?? 10
  const lease = await service.prepareCf({ stamp, sourceRects: [R(0, rows - 1)], signal: new AbortController().signal, grantBytes: 8 * 1024 * 1024 }, source)
  const out = Array.from({ length: rows }, (_, r) => lease.value.at(r, 0))
  lease.release(); service.dispose()
  return { at: (r: number) => out[r]!, all: out }
}
const iconOf = (o: CfOutcome<CfCellState>) => o.kind === 'ready' ? o.value.visual?.iconSet : undefined
const taintsIcon = (o: CfOutcome<CfCellState>) => o.kind === 'ready' && o.value.unresolvedWrites?.has('iconSet') === true

describe('icon participation and scope gates', () => {
  const NUM_DEF = iconDef({ thresholds: NUM3 })
  const values = (): XlsxCell[] => range(0, 9).map(r => cellAt(r, r * 5))

  it('S4-G: numbers evaluate; other kinds draw no icon only with icon:participation-types; date serials need their id', async () => {
    const cells = [cellAt(0, 5), cellAt(1, 'text'), cellAt(2, true), cellAt(3, '#DIV/0!', { valueIsError: true }), cellAt(4, ''), cellAt(6, 45)]
    const noTypes = EVAL.filter(c => c !== 'icon:participation-types')
    const without = await serviceRun([iconRule(1, NUM_DEF, [R(0, 6)])], cells, noTypes, { rows: 7 })
    expect(iconOf(without.at(0))?.bucket).toBe(0)
    expect(iconOf(without.at(6))?.bucket).toBe(2)
    for (const r of [1, 2, 3, 4, 5]) { expect(iconOf(without.at(r)), `row ${r}`).toBeUndefined(); expect(taintsIcon(without.at(r)), `row ${r}`).toBe(true) }
    const withTypes = await serviceRun([iconRule(1, NUM_DEF, [R(0, 6)])], cells, EVAL, { rows: 7 })
    for (const r of [1, 2, 3, 4, 5]) { expect(iconOf(withTypes.at(r)), `row ${r}`).toBeUndefined(); expect(taintsIcon(withTypes.at(r)), `row ${r}`).toBe(false) }

    // A text cell taints the icon slot for a lower icon rule without evidence.
    const lower = await serviceRun([iconRule(1, NUM_DEF, [R(1, 1)]), iconRule(2, NUM_DEF, [R(0, 6)])], cells, [...noTypes, 'icon:composition-overlapping-icons'], { rows: 7 })
    expect(iconOf(lower.at(1))).toBeUndefined()
    expect(iconOf(lower.at(0))?.ruleId).toBe(2)

    // A date serial (a date-formatted number) evaluates only with icon:participation-dateserial.
    const dated = [cellAt(0, 45000, { style: { numFmtId: 14 } as XlsxCell['style'] }), cellAt(1, 25)]
    const noDate = await serviceRun([iconRule(1, NUM_DEF, [R(0, 1)])], dated, EVAL, { rows: 2 })
    expect(iconOf(noDate.at(0))).toBeUndefined()
    expect(taintsIcon(noDate.at(0))).toBe(true)
    expect(iconOf(noDate.at(1))?.bucket).toBe(1)
    const withDate = await serviceRun([iconRule(1, NUM_DEF, [R(0, 1)])], dated, [...EVAL, 'icon:participation-dateserial'], { rows: 2 })
    expect(iconOf(withDate.at(0))?.bucket).toBe(2)
    // A source that cannot tell dates from numbers withholds every icon until then.
    const unknownFormats = await serviceRun([iconRule(1, NUM_DEF, [R(0, 1)])], [cellAt(1, 25)], EVAL, { rows: 2, numberFormats: 'unknown' })
    expect(iconOf(unknownFormats.at(1))).toBeUndefined()
  })

  it('S4-G: an error makes a percent rule unresolved but not a num-only rule', async () => {
    const cells = [...values().slice(0, 9), cellAt(9, '#N/A', { valueIsError: true })]
    const pct = await serviceRun([iconRule(1, iconDef())], cells, EVAL)
    for (let r = 0; r < 9; r++) expect(iconOf(pct.at(r)), `row ${r}`).toBeUndefined()
    expect(taintsIcon(pct.at(3))).toBe(true)
    const lifted = await serviceRun([iconRule(1, iconDef())], cells, [...EVAL, 'icon:participation-error-presence'])
    expect(iconOf(lifted.at(8))?.bucket).toBe(2)
    const numOnly = await serviceRun([iconRule(1, NUM_DEF)], cells, EVAL)
    expect(iconOf(numOnly.at(8))?.bucket).toBe(2)
    expect(iconOf(numOnly.at(9))).toBeUndefined()
    // The error cell itself follows the per-cell table: no icon, slot untainted under participation-types.
    expect(taintsIcon(numOnly.at(9))).toBe(false)
  })

  it('S4-H: each scope shape is unresolved until its icon:scope-* evidence, and bar:scope-* never lifts it', async () => {
    const shapes: Array<{ cap: string; rules: () => CfRule[]; options?: Parameters<typeof serviceRun>[3] }> = [
      { cap: 'icon:scope-disjoint', rules: () => [iconRule(1, iconDef(), [R(0, 4), R(5, 9)])] },
      { cap: 'icon:scope-duplicate-blocks', rules: () => [iconRule(1, iconDef()), iconRule(2, iconDef({ set: '3Flags' }))] },
      { cap: 'icon:scope-whole-column', rules: () => [iconRule(1, iconDef(), [R(0, 1048575)])] },
      { cap: 'icon:scope-merged-cell', rules: () => [iconRule(1, iconDef())], options: { merges: [R(2, 3)] } },
      { cap: 'icon:scope-hidden-row', rules: () => [iconRule(1, iconDef())], options: { hiddenRows: [R(4, 4)] } },
    ]
    const overlap = ['icon:composition-overlapping-icons']
    for (const shape of shapes) {
      const extra = shape.cap === 'icon:scope-hidden-row' ? ['icon:scope-autofilter-hidden-row'] : []
      const barTwin = [...EVAL, ...overlap, ...extra.map(c => c.replace('icon:', 'bar:')), shape.cap.replace('icon:', 'bar:')]
      const blocked = await serviceRun(shape.rules(), values(), barTwin, shape.options)
      expect(iconOf(blocked.at(9)), shape.cap).toBeUndefined()
      expect(taintsIcon(blocked.at(9)), shape.cap).toBe(true)
      const lifted = await serviceRun(shape.rules(), values(), [...EVAL, ...overlap, ...extra, shape.cap], shape.options)
      expect(iconOf(lifted.at(9))?.bucket, shape.cap).toBe(2)
    }
  })

  it('an icon rule carrying a DXF is withheld whole: icon and DXF writes tainted (fail closed)', async () => {
    const rule = iconRule(1, NUM_DEF, [R(0, 9)], { writes: new Set(['iconSet', 'fill']), dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FFFFC7CE' } })
    const run = await serviceRun([rule], values(), EVAL)
    const st = run.at(9)
    expect(iconOf(st)).toBeUndefined()
    if (st.kind === 'ready') {
      expect(st.value.properties).toEqual({})
      expect(new Set(st.value.unresolvedWrites)).toEqual(new Set(['iconSet', 'fill']))
    }
  })
})
