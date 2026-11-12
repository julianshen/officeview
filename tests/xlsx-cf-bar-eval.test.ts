import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import {
  barEvaluationRequiredCapabilities,
  evaluateDataBar,
  resolveBarThresholds,
  type CfBarLayoutVariant,
} from '../src/xlsx/cf/dataBar'
import { resolveScaleThresholds } from '../src/xlsx/cf/colorScale'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { createCfService } from '../src/xlsx/cf/service'
import type { CfRangeStatistics } from '../src/xlsx/cf/statistics'
import type { CfDefinition, CfLease, CfLiteral, CfOutcome, CfRule, CfSheetRules, CfStamp, CfValuePage, CfValueSource, SourceRect } from '../src/xlsx/cf/model'

// Every numeric expectation in this file is a HYPOTHESIS run under a stubbed variant and stubbed evidence, never an
// Excel fact (Task 20 brief). In particular the linear 10/90 mapping contradicts the Stage 3 capture (9/44/115/150 px
// for 0/25/75/100 % of a 177 px span, docs/validation/xlsx-cf.md §11), which is why `bar:lengths-default-10-90` stays
// blocked and gates it.

type BarDef = Extract<CfDefinition, { kind: 'dataBar' }>
const BLUE = 0xff0000ff, RED = 0xffff0000, BLACK = 0xff000000, GREEN = 0xff00b050

function bar(overrides: Partial<BarDef> = {}): BarDef {
  return {
    kind: 'dataBar', dialect: 'base', formulas: [],
    bounds: [{ type: 'min' }, { type: 'max' }],
    fillArgb: BLUE, negativeFillArgb: RED, axisArgb: BLACK,
    minLength: 0, maxLength: 100, gradient: false, border: false,
    negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: true,
    axisPosition: 'automatic', direction: 'leftToRight', showValue: true, sheetRightToLeft: false,
    ...overrides,
  }
}

function makeStats(numbers: number[], extra: Partial<CfRangeStatistics> = {}): CfRangeStatistics {
  // Test-side oracle data only (the source assertion below keeps sorting out of dataBar.ts).
  const sorted = [...numbers].sort((a, b) => a - b)
  return {
    scopeKey: 'bar-scope', numericCount: numbers.length, sum: 0, mean: undefined,
    min: sorted[0], max: sorted[sorted.length - 1], stddevPop: undefined, stddevSample: undefined,
    sortedAsc: new Float64Array(sorted), frequencies: new Map(),
    declaredCells: numbers.length, spanCells: numbers.length, storedCells: numbers.length, ...extra,
  }
}

const SPAN_CELL: CfBarLayoutVariant = { origin: 'span', lengthBasis: 'cell' }
const AXIS_SIDE: CfBarLayoutVariant = { origin: 'axis', lengthBasis: 'side' }

/** Every evaluation capability a bar of these settings could need, as a stub (never production evidence). */
const STUB_EVIDENCE = [
  'bar:bounds-min-max', 'bar:bounds-num', 'bar:bounds-percent', 'bar:bounds-percentile', 'bar:bounds-automin', 'bar:bounds-automax',
  'bar:lengths-default-10-90', 'bar:lengths-explicit-0-100', 'bar:render-solid', 'bar:render-border',
  'bar:axis-automatic', 'bar:axis-middle', 'bar:axis-none', 'bar:direction-ltr', 'bar:direction-rtl', 'bar:direction-context-ltr-sheet',
  'bar:negative-fill-color', 'bar:negative-fill-same-as-positive', 'bar:negative-border-color', 'bar:negative-border-same-as-positive',
  'bar:participation-types',
]

const num = (value: number): CfLiteral => ({ kind: 'number', value })

/** Threshold resolution then evaluation, the service's per-cell path, as one pure call. */
function evalBar(v: CfLiteral | number, def: BarDef, scope: number[], variant?: CfBarLayoutVariant) {
  const stats = makeStats(scope)
  const thresholds = resolveBarThresholds(def, stats)
  if (thresholds.kind !== 'ready') return thresholds
  return evaluateDataBar(v, def, thresholds.value, stats, variant ? { variant } : {})
}

function ready<T>(out: CfOutcome<T>): T {
  expect(out.kind, JSON.stringify(out)).toBe('ready')
  return (out as { kind: 'ready'; value: T }).value
}

const range = (lo: number, hi: number): number[] => Array.from({ length: hi - lo + 1 }, (_, i) => lo + i)

describe('CF Stage 3 Task 20 — data-bar evaluation (hypotheses under stubbed variants)', () => {
  it('no layout variant ⇒ unresolved (S3-B)', () => {
    const scope = range(0, 100)
    // With no evidence at all, every bar is unresolved.
    withExactCfEvidence([], () => {
      for (const def of [bar(), bar({ axisPosition: 'middle' }), bar({ gradient: true }), bar({ minLength: 10, maxLength: 90 })]) {
        for (const v of [0, 25, 100]) {
          expect(evalBar(v, def, scope, SPAN_CELL)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
          expect(evaluateDataBar(v, def, [0, 100], makeStats(scope))).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
        }
      }
    })
    // Every capability stubbed, but no variant: still unresolved. Exactly one variant must be enabled.
    withExactCfEvidence(STUB_EVIDENCE, () => {
      expect(evalBar(25, bar(), scope)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
    })
  })

  it('no layout variant ⇒ CfCellState.visual carries no dataBar (S3-B, service path)', async () => {
    const out = await prepareBars([barRule(0, bar())], [0, 50, 100], [], undefined)
    for (const row of [0, 1, 2]) {
      const st = out.at(row)
      expect(st.kind === 'ready' ? st.value.visual?.dataBar : undefined).toBeUndefined()
    }
    out.release()
  })

  it('all-positive lengths under stubbed variants', () => {
    const scope = range(0, 100)
    withExactCfEvidence(STUB_EVIDENCE, () => {
      const d1090 = bar({ minLength: 10, maxLength: 90 })
      const expected: Array<[number, number]> = [[0, 0.10], [25, 0.30], [50, 0.50], [100, 0.90]]
      for (const [v, end] of expected) {
        const b = ready(evalBar(v, d1090, scope, SPAN_CELL))
        expect(b.startFraction, `v=${v}`).toBe(0)
        expect(b.endFraction, `v=${v}`).toBeCloseTo(end, 12)
        expect(b.isNegative).toBe(false)
        expect(b.axisColor).toBeUndefined()
        expect(b.fillColor).toBe(BLUE)
      }
      const b0100 = ready(evalBar(25, bar(), scope, SPAN_CELL))
      expect([b0100.startFraction, b0100.endFraction]).toEqual([0, 0.25])
    })
    // 10/90 is gated by its own (blocked) capability: the capture shows Excel's default lengths are not linear.
    withExactCfEvidence(STUB_EVIDENCE.filter(c => c !== 'bar:lengths-default-10-90'), () => {
      expect(evalBar(25, bar({ minLength: 10, maxLength: 90 }), scope, SPAN_CELL).kind).toBe('unresolved')
      expect(evalBar(25, bar(), scope, SPAN_CELL).kind).toBe('ready')
    })
    // A length pair with no capability at all never resolves.
    withExactCfEvidence(STUB_EVIDENCE, () => {
      expect(evalBar(25, bar({ minLength: 5, maxLength: 95 }), scope, SPAN_CELL).kind).toBe('unresolved')
    })
  })

  it('autoMin and autoMax bounds', () => {
    withExactCfEvidence(STUB_EVIDENCE, () => {
      const scope = [20, 40, 60]
      // Hypothesis: autoMin = min(0, v_min), autoMax = max(0, v_max).
      const auto = ready(evalBar(40, bar({ bounds: [{ type: 'autoMin' }, { type: 'max' }] }), scope, SPAN_CELL))
      expect(Math.abs(auto.endFraction - 2 / 3)).toBeLessThanOrEqual(1e-9)
      expect(auto.endFraction).toBeCloseTo(0.6667, 4)
      const min = ready(evalBar(40, bar(), scope, SPAN_CELL))
      expect(min.endFraction).toBe(0.5)
      // All-positive autoMax equals max.
      const s = makeStats(scope)
      expect(ready(resolveBarThresholds(bar({ bounds: [{ type: 'min' }, { type: 'autoMax' }] }), s))).toEqual(ready(resolveBarThresholds(bar(), s)))
      expect(ready(resolveBarThresholds(bar({ bounds: [{ type: 'autoMin' }, { type: 'autoMax' }] }), s))).toEqual([0, 60])
    })
    // Each auto bound needs its own capability.
    withExactCfEvidence(STUB_EVIDENCE.filter(c => c !== 'bar:bounds-automin'), () => {
      expect(resolveBarThresholds(bar({ bounds: [{ type: 'autoMin' }, { type: 'max' }] }), makeStats([20, 40, 60])).kind).toBe('unresolved')
    })
  })

  it('mixed-sign axis under stubbed variants', () => {
    const scope = range(-50, 100)
    withExactCfEvidence(STUB_EVIDENCE, () => {
      // automatic, { axis, side }: the spec formula axis = -v_min / (v_max - v_min) = 1/3.
      const p = ready(evalBar(50, bar(), scope, AXIS_SIDE))
      expect(p.axisFraction).toBeCloseTo(1 / 3, 12)
      expect(p.startFraction).toBeCloseTo(1 / 3, 12)
      expect(p.endFraction).toBeCloseTo(2 / 3, 12)
      expect(p.isNegative).toBe(false)
      expect(p.fillColor).toBe(BLUE)
      expect(p.axisColor).toBe(BLACK)
      const n = ready(evalBar(-25, bar(), scope, AXIS_SIDE))
      expect(n.startFraction).toBeCloseTo(1 / 6, 12)
      expect(n.endFraction).toBeCloseTo(1 / 3, 12)
      expect(n.isNegative).toBe(true)
      expect(n.fillColor).toBe(RED)

      // middle: axis 0.5, both sides share max(|v_min|, |v_max|).
      const mp = ready(evalBar(50, bar({ axisPosition: 'middle' }), scope, AXIS_SIDE))
      expect([mp.axisFraction, mp.startFraction, mp.endFraction]).toEqual([0.5, 0.5, 0.75])
      expect(mp.axisColor).toBe(BLACK)
      const mn = ready(evalBar(-25, bar({ axisPosition: 'middle' }), scope, AXIS_SIDE))
      expect([mn.startFraction, mn.endFraction]).toEqual([0.375, 0.5])

      // none: no axis line, and the bar runs from the leading edge.
      const nn = ready(evalBar(-25, bar({ axisPosition: 'none' }), scope, AXIS_SIDE))
      expect(nn.axisColor).toBeUndefined()
      expect(nn.startFraction).toBe(0)
      expect(nn.endFraction).toBeCloseTo(1 / 6, 12)

      // Same-as-positive, with its evidence: the negative bar takes the positive fill.
      const same = ready(evalBar(-25, bar({ negativeBarColorSameAsPositive: true, negativeFillArgb: undefined }), scope, AXIS_SIDE))
      expect(same.fillColor).toBe(BLUE)
      expect(same.isNegative).toBe(true)

      // An axis needs the axis-origin variant; the span variant draws no axis layout.
      expect(evalBar(50, bar(), scope, SPAN_CELL).kind).toBe('unresolved')
      // A colour the cell draws but the definition lacks is never defaulted (S3-F).
      expect(evalBar(-25, bar({ negativeFillArgb: undefined }), scope, AXIS_SIDE).kind).toBe('unresolved')
      expect(evalBar(50, bar({ axisArgb: undefined }), scope, AXIS_SIDE).kind).toBe('unresolved')
      // Border colours come from the definition too.
      const bordered = ready(evalBar(-25, bar({ border: true, borderArgb: GREEN }), scope, AXIS_SIDE))
      expect(bordered.borderColor).toBe(GREEN)
      expect(evalBar(-25, bar({ border: true }), scope, AXIS_SIDE).kind).toBe('unresolved')
    })
    // Negative settings and the automatic axis are each gated.
    withExactCfEvidence(STUB_EVIDENCE.filter(c => c !== 'bar:negative-fill-color'), () => {
      expect(evalBar(-25, bar(), scope, AXIS_SIDE).kind).toBe('unresolved')
      expect(evalBar(50, bar(), scope, AXIS_SIDE).kind).toBe('ready')
    })
    withExactCfEvidence(STUB_EVIDENCE.filter(c => c !== 'bar:negative-fill-same-as-positive'), () => {
      expect(evalBar(-25, bar({ negativeBarColorSameAsPositive: true }), scope, AXIS_SIDE).kind).toBe('unresolved')
    })
    withExactCfEvidence(STUB_EVIDENCE.filter(c => c !== 'bar:axis-automatic'), () => {
      expect(evalBar(50, bar(), scope, AXIS_SIDE).kind).toBe('unresolved')
      // All-positive data under automatic draws no axis, so it does not need the axis capability.
      expect(evalBar(50, bar(), range(0, 100), AXIS_SIDE).kind).toBe('ready')
    })
    // All-negative data has no hypothesis yet: withheld (bar:lengths-all-negative).
    withExactCfEvidence([...STUB_EVIDENCE, 'bar:lengths-all-negative'], () => {
      expect(evalBar(-25, bar(), range(-100, -1), AXIS_SIDE).kind).toBe('unresolved')
    })
  })

  it('bar evaluation stays finite at ±1e308 (Lesson 6)', () => {
    const allFinite = (out: unknown) => {
      const walk = (v: unknown): void => {
        if (typeof v === 'number') expect(Number.isFinite(v), String(v)).toBe(true)
        else if (v && typeof v === 'object') for (const x of Object.values(v)) walk(x)
      }
      walk(out)
    }
    withExactCfEvidence(STUB_EVIDENCE, () => {
      const def = bar({ bounds: [{ type: 'num', value: -1e308 }, { type: 'num', value: 1e308 }] })
      const scope = [-1e308, 0, 1e308]
      expect(1e308 - -1e308).toBe(Infinity) // the span itself overflows
      for (const variant of [SPAN_CELL, AXIS_SIDE, { origin: 'axis', lengthBasis: 'cell' } as const]) {
        for (const d of [def, bar({ ...def, axisPosition: 'none' }), bar({ ...def, axisPosition: 'middle' })]) {
          const out = evalBar(0, d, scope, variant)
          allFinite(out)
          if (d.axisPosition !== 'none' && variant.origin === 'span') { expect(out.kind).toBe('unresolved'); continue } // no axis layout
          const b = ready(out)
          expect(b.endFraction).toBe(0.5)
          for (const v of [-1e308, -1e307, 1e307, 1e308]) allFinite(ready(evalBar(v, d, scope, variant)))
        }
      }
      // Degenerate scopes (S3-E): v_min = v_max, a single numeric cell, inverted bounds — unresolved, never NaN.
      const degenerate: Array<[BarDef, number[]]> = [
        [bar(), [5, 5, 5]],
        [bar(), [5]],
        [bar({ bounds: [{ type: 'num', value: 100 }, { type: 'num', value: 10 }] }), [10, 50, 100]],
        [bar({ bounds: [{ type: 'num', value: 1e308 }, { type: 'num', value: -1e308 }] }), [-1e308, 0, 1e308]],
      ]
      for (const [d, scope2] of degenerate) {
        const out = evalBar(5, d, scope2, SPAN_CELL)
        expect(out.kind).toBe('unresolved')
        allFinite(out)
        // Direct evaluation against degenerate thresholds is unresolved too (defence in depth).
        const direct = evaluateDataBar(5, d, [5, 5], makeStats(scope2), { variant: SPAN_CELL })
        expect(direct.kind).toBe('unresolved')
        allFinite(direct)
      }
      // Cells beyond the num bounds (S3-E): the whole rule is unresolved, never clamped.
      const narrow = bar({ bounds: [{ type: 'num', value: 20 }, { type: 'num', value: 80 }] })
      expect(evalBar(50, narrow, range(0, 100), SPAN_CELL).kind).toBe('unresolved')
      expect(evalBar(50, narrow, range(20, 80), SPAN_CELL).kind).toBe('ready')
      // Non-finite cell values withhold.
      expect(evaluateDataBar(Number.NaN, bar(), [0, 100], makeStats(range(0, 100)), { variant: SPAN_CELL }).kind).toBe('unresolved')
    })
  })

  it('bars reuse the shared threshold code, and dataBar.ts never sorts', () => {
    withExactCfEvidence([...STUB_EVIDENCE, 'scale:thresholds-percentile-inc-exc'], () => {
      const s = makeStats([3, 1, 4, 1, 5, 9, 2, 6, 5, 3, 5, 8, 9, 7])
      const bounds = [{ type: 'percentile', value: 25 }, { type: 'percentile', value: 75 }] as const
      expect(ready(resolveBarThresholds(bar({ bounds }), s))).toEqual(ready(resolveScaleThresholds(bounds.map(bound => ({ bound })), s)))
    })
    // Bar gates are bar:* evidence, never inherited from scale:* (production verifies scale:*, not bar:*).
    withExactCfEvidence(['scale:thresholds-min-max', 'scale:thresholds-percentile-inc-exc', 'scale:degenerate-single-cell'], () => {
      expect(resolveBarThresholds(bar(), makeStats([1, 2])).kind).toBe('unresolved')
      expect(resolveBarThresholds(bar(), makeStats([1])).kind).toBe('unresolved')
    })
    const source = readFileSync(resolve(process.cwd(), 'src/xlsx/cf/dataBar.ts'), 'utf8')
    expect(source).not.toMatch(/sort\(/)
    expect(source).not.toMatch(/function\s+percentile|rank\s*=|Math\.floor\(rank/)
    expect(source).toMatch(/resolveScaleThresholds/)
    expect(source).toMatch(/safeFraction/)
  })

  it('barEvaluationRequiredCapabilities lists the settings gates of the definition', () => {
    const req = barEvaluationRequiredCapabilities(bar({ gradient: true, border: true, axisPosition: 'middle', minLength: 10, maxLength: 90, bounds: [{ type: 'autoMin' }, { type: 'percentile', value: 90 }] }))
    for (const cap of ['bar:render-gradient-stops', 'bar:render-border', 'bar:axis-middle', 'bar:lengths-default-10-90', 'bar:bounds-automin', 'bar:bounds-percentile', 'bar:direction-ltr']) {
      expect(req.has(cap), cap).toBe(true)
    }
    expect(barEvaluationRequiredCapabilities(bar()).has('bar:render-solid')).toBe(true)
    expect(barEvaluationRequiredCapabilities(bar({ direction: 'context', sheetRightToLeft: undefined })).has('bar:direction-context-ltr-sheet')).toBe(false)
  })
})

// --- service path -------------------------------------------------------------------------------------------------

const SHEET = 'bars'
const stamp: CfStamp = { owner: 't20', sheetId: SHEET, workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1 }

function barRule(order: number, def: BarDef, ranges: readonly SourceRect[] = [{ firstRow: 0, lastRow: 99, firstCol: 0, lastCol: 0 }], extra: Partial<CfRule> = {}): CfRule {
  return { id: order, priority: order + 1, documentOrder: order, ranges, stopIfTrue: false, writes: new Set(['dataBar']), definition: def, ...extra }
}

/** Column A values (row i holds values[i]); a literal or a number. */
function columnSource(values: readonly (CfLiteral | number)[]): CfValueSource {
  const lits = values.map(v => typeof v === 'number' ? num(v) : v)
  return {
    bounds: async () => ({ firstRow: 0, lastRow: Math.max(0, lits.length - 1), firstCol: 0, lastCol: 0 }),
    stamp: () => stamp,
    pages: async function* (_id: string, rects: readonly SourceRect[]) {
      const rows: number[] = [], vals: CfLiteral[] = []
      for (const r of rects) for (let row = r.firstRow; row <= Math.min(r.lastRow, lits.length - 1); row++) {
        if (r.firstCol <= 0 && r.lastCol >= 0) { rows.push(row); vals.push(lits[row]!) }
      }
      yield { value: { rows: Uint32Array.from(rows), cols: new Uint16Array(rows.length), values: vals }, bytes: 0, release() {} } as unknown as CfLease<CfValuePage>
    },
  }
}

async function prepareBars(rules: CfRule[], values: readonly (CfLiteral | number)[], evidence: readonly string[], variant: CfBarLayoutVariant | undefined,
  opts: { merges?: readonly SourceRect[]; hiddenRows?: readonly SourceRect[] } = {}) {
  const memory = createCfMemory()
  const allowed = new Set(evidence)
  const sheet: CfSheetRules = { rules, diagnostics: [] }
  const service = createCfService(sheet, memory, {
    evidence: { allows: id => allowed.has(id) },
    ...(variant ? { barLayoutVariant: variant } : {}),
    ...(opts.merges ? { merges: opts.merges } : {}),
    ...(opts.hiddenRows ? { statisticsOptions: { hiddenRows: opts.hiddenRows } } : {}),
  })
  const rect = { firstRow: 0, lastRow: Math.max(0, values.length - 1), firstCol: 0, lastCol: 0 }
  const lease = await service.prepareCf({ stamp, sourceRects: [rect], signal: new AbortController().signal, grantBytes: 8 * 1024 * 1024 }, columnSource(values))
  return {
    at: (row: number) => lease.value.at(row, 0),
    release: () => { lease.release(); service.dispose(); expect(memory.currentBytes).toBe(0) },
  }
}

const barAt = (o: CfOutcome<import('../src/xlsx/cf/model').CfCellState>) => o.kind === 'ready' ? o.value.visual?.dataBar : undefined

describe('CF Stage 3 Task 20 — bar participation and scope gates (S3-G, S3-H)', () => {
  it('bar participation and scope gates', async () => {
    const values: Array<CfLiteral | number> = [0, 50, 100, { kind: 'string', value: 'x' }]
    // Under stubbed evidence and a variant, the numeric cells carry bars with the rule id.
    const ok = await prepareBars([barRule(0, bar())], values, STUB_EVIDENCE, SPAN_CELL)
    expect(barAt(ok.at(1))).toMatchObject({ ruleId: 0, startFraction: 0, endFraction: 0.5, isNegative: false, direction: 'leftToRight' })
    // Text with bar:participation-types: no bar, no taint.
    expect(barAt(ok.at(3))).toBeUndefined()
    const textOk = ok.at(3)
    expect(textOk.kind).toBe('ready')
    if (textOk.kind === 'ready') expect(textOk.value.unresolvedWrites?.has('dataBar') ?? false).toBe(false)
    ok.release()

    // Text without bar:participation-types taints the bar slot for a lower bar.
    const twoBars = [barRule(0, bar(), [{ firstRow: 3, lastRow: 3, firstCol: 0, lastCol: 0 }]), barRule(1, bar())]
    const noTypes = await prepareBars(twoBars, values, STUB_EVIDENCE.filter(c => c !== 'bar:participation-types'), SPAN_CELL)
    const text = noTypes.at(3)
    expect(text.kind).toBe('ready')
    if (text.kind === 'ready') {
      expect(text.value.visual?.dataBar).toBeUndefined()
      expect(text.value.unresolvedWrites?.has('dataBar')).toBe(true)
    }
    expect(barAt(noTypes.at(1))?.endFraction).toBe(0.5)
    noTypes.release()

    // An error anywhere in the scope makes the whole rule unresolved, unless error presence is evidenced.
    const withError: Array<CfLiteral | number> = [0, 50, 100, { kind: 'error', value: '#DIV/0!' }]
    const err = await prepareBars([barRule(0, bar())], withError, STUB_EVIDENCE, SPAN_CELL)
    for (const row of [0, 1, 2]) expect(barAt(err.at(row)), `row ${row}`).toBeUndefined()
    err.release()
    const errOk = await prepareBars([barRule(0, bar())], withError, [...STUB_EVIDENCE, 'bar:participation-error-presence'], SPAN_CELL)
    expect(barAt(errOk.at(1))?.endFraction).toBe(0.5)
    errOk.release()

    // S3-H shapes: each unresolved until its own bar:scope-* evidence (never scale:scope-*).
    const nums = [0, 50, 100]
    const a = (r0: number, r1: number) => ({ firstRow: r0, lastRow: r1, firstCol: 0, lastCol: 0 })
    const shapes: Array<[string, CfRule[], { merges?: SourceRect[]; hiddenRows?: SourceRect[] }]> = [
      ['bar:scope-disjoint', [barRule(0, bar(), [a(0, 0), a(1, 2)])], {}],
      ['bar:scope-whole-column', [barRule(0, bar(), [a(0, 1048575)])], {}],
      ['bar:scope-duplicate-blocks', [barRule(0, bar(), [a(0, 2)]), barRule(1, bar(), [a(0, 2)])], {}],
      ['bar:scope-merged-cell', [barRule(0, bar(), [a(0, 2)])], { merges: [{ firstRow: 1, lastRow: 1, firstCol: 0, lastCol: 1 }] }],
      ['bar:scope-hidden-row', [barRule(0, bar(), [a(0, 2)])], { hiddenRows: [a(2, 2)] }],
    ]
    for (const [cap, rules, opts] of shapes) {
      const scaleCap = cap.replace('bar:', 'scale:')
      const extra = cap === 'bar:scope-hidden-row' ? ['bar:scope-autofilter-hidden-row'] : []
      const blocked = await prepareBars(rules, nums, [...STUB_EVIDENCE, scaleCap, ...extra], SPAN_CELL, opts)
      expect(barAt(blocked.at(0)), `${cap} withholds without its evidence`).toBeUndefined()
      blocked.release()
      if (cap === 'bar:scope-duplicate-blocks') continue // two bars on one cell stay withheld (exclusivity, S3-I)
      const lifted = await prepareBars(rules, nums, [...STUB_EVIDENCE, cap, ...extra], SPAN_CELL, opts)
      expect(barAt(lifted.at(0)), `${cap} lifts with its evidence`).toBeDefined()
      lifted.release()
    }
  })
})
