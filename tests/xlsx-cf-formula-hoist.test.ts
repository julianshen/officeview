// @vitest-environment node
/**
 * Stage 5 Task 34, Ruling S5-HOIST: absolute-range aggregates are hoisted out of the per-cell call into a chunked scan
 * with engine-ordered reducers, and the COUNTIF frequency map reproduces the engine's plain-equality arm only.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
// functions.ts first: importing cond.ts before it evaluates FUNCTIONS before CONDITIONAL_FUNCTIONS exists (an engine
// import cycle, functions.ts <-> evaluator.ts <-> cond.ts).
import '../src/xlsx/formula/functions'
import { CONDITIONAL_MAX_CRITERION_LENGTH } from '../src/xlsx/formula/functions/cond'
import { DEFAULT_MAX_DP_CELLS, matchWildcard } from '../src/xlsx/formula/functions/wildcard'
import * as schedulerModule from '../src/xlsx/view/scheduler'
import { cfCountifDpRefused, cfCountifKey, cfTruth, planCfHoist, type CfCallWorkBreakdown, type CfHoisted } from '../src/xlsx/cf/formula'
import { createCfMemory, createCfUnitCeiling, createCfWorkBudget, CF_FREQUENCY_ENTRY_BYTES, CF_LIMITS, type CfWorkBudget } from '../src/xlsx/cf/limits'
import { createCfStatistics } from '../src/xlsx/cf/statistics'
import { createCfService } from '../src/xlsx/cf/service'
import type { CfLiteral, CfRule } from '../src/xlsx/cf/model'
import {
  EVALUATION, a1, batchOf, compile, engineEval, engineOf, fakeSource, formulaOf, formulaRule, literalOf, rectOf, refOf, type CellInput, type Cells, type FakeSource,
} from './helpers/cf-formula'

const hostHooks: { onHost?: () => void; count: number } = { count: 0 }
vi.mock('../src/xlsx/view/scheduler', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/view/scheduler')>()
  return { ...actual, hostTask: vi.fn(async () => { hostHooks.count++; hostHooks.onHost?.() }) }
})
void schedulerModule
afterEach(() => { hostHooks.onHost = undefined; hostHooks.count = 0 })
const frozen = () => 0
const scan = (signal = new AbortController().signal): CfWorkBudget => createCfWorkBudget({ now: frozen, signal, scan: true })
/** Deterministic PRNG (mulberry32). */
function rng(seed: number): () => number {
  let a = seed >>> 0
  return () => { a = (a + 0x6d2b79f5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296 }
}
const pick = <T>(r: () => number, xs: readonly T[]): T => xs[Math.floor(r() * xs.length)]!
async function hoistOf(text: string, cells: Cells, options: { memory?: ReturnType<typeof createCfMemory>; source?: FakeSource; signal?: AbortSignal } = {}) {
  const rule = formulaRule(text, 'Z1')
  const plan = planCfHoist(formulaOf(rule))
  const memory = options.memory ?? createCfMemory()
  const session = createCfStatistics(memory)
  const source = options.source ?? fakeSource(cells)
  const out = await session.prepareHoists(rule, plan, source.stamp('s1'), source, scan(options.signal), options.signal ? { signal: options.signal } : {})
  return { rule, plan, out, session, memory, source }
}
const hoisted = (o: { kind: string; value?: unknown }): CfHoisted => { expect(o.kind).toBe('ready'); return (o as { value: { value: CfHoisted } }).value.value }

describe('S5-HOIST: the engine limits the hoist mirrors are the engine\'s own named constants', () => {
  it('CONDITIONAL_MAX_CRITERION_LENGTH and DEFAULT_MAX_DP_CELLS drive the engine refusals unchanged', () => {
    expect(CONDITIONAL_MAX_CRITERION_LENGTH).toBe(255) // cond.ts criterionOf, `value.length > 255` before Task 34
    expect(DEFAULT_MAX_DP_CELLS).toBe(4_000_000) // wildcard.ts, unchanged
    const at = { row: 0, col: 2 }
    const cells = { A1: 'x'.repeat(CONDITIONAL_MAX_CRITERION_LENGTH), B1: 'x'.repeat(CONDITIONAL_MAX_CRITERION_LENGTH + 1) }
    expect(engineEval('COUNTIF(A1,A1)', cells, at).features).toEqual([])
    expect(engineEval('COUNTIF(A1,B1)', cells, at).features).toEqual(['conditional-long-criterion-unverified'])
    // (tokens + 1) x (length + 1) at the limit matches; one DP cell over refuses.
    const n = DEFAULT_MAX_DP_CELLS / 2 - 1 // pattern 'a' (1 token): 2 x (n + 1) = DEFAULT_MAX_DP_CELLS
    expect(matchWildcard('a', 'b'.repeat(n)).status).toBe('not-matched')
    expect(matchWildcard('a', 'b'.repeat(n + 1)).status).toBe('resource-limit')
  })
})


describe('absolute-range aggregates are hoisted and paint on a 10,000-row column (S5-HOIST)', () => {
  const ROWS = 10_000
  const valueAt = (r: number): CellInput => (r === 1 ? 'Value' : r % 40 === 0 ? pick(rng(r), ['abc', 'ABC', 'Abc', 'zz']) : (r * 37) % 997)
  const cells: Cells = Object.fromEntries(Array.from({ length: ROWS }, (_, i) => [`A${i + 1}`, valueAt(i + 1)]))
  const engine = engineOf(cells)
  const texts = ['COUNTIF($A:$A,A2)>1', 'A2>AVERAGE($A:$A)']
  // The engine's own per-cell evaluation, computed once per distinct value (each formula reads only A2's value).
  const expectation = texts.map(text => {
    const memo = new Map<string, unknown>()
    return (row: number): boolean => {
      const key = JSON.stringify(cells[`A${row + 1}`])
      if (!memo.has(key)) memo.set(key, engine(text.replace(/A2/g, `A${row + 1}`), { row, col: 1 }).value)
      return memo.get(key) === true
    }
  })
  const batches = [rectOf('A2:A257'), rectOf('A5000:A5255'), rectOf('A9745:A10000')]
  async function paint(hoist: boolean) {
    const source = fakeSource(cells)
    const memory = createCfMemory()
    const session = createCfStatistics(memory)
    const rules = texts.map(t => formulaRule(t, `A2:A${ROWS}`))
    const calls: CfCallWorkBreakdown[] = []
    const scanBudgets: CfWorkBudget[] = []
    const scanBudget = () => { const b = createCfWorkBudget({ now: frozen, scan: true }); scanBudgets.push(b); return b }
    const painted: boolean[][] = [[], []], refused: string[] = []
    let snapshotEntries = 0, ceilingUsed = 0, callUnits = 0, hostDuringPrepare = 0
    for (const target of batches) {
      const ceiling = createCfUnitCeiling()
      hostHooks.count = 0
      const batch = await batchOf(rules, source, { hoist, statistics: session, targets: [target], ceiling, scanBudget, onCall: b => { calls.push(b); callUnits += b.total } })
      hostDuringPrepare += hostHooks.count
      snapshotEntries = Math.max(snapshotEntries, batch.snapshotEntries)
      for (let row = target.firstRow; row <= target.lastRow; row++) for (let t = 0; t < 2; t++) {
        const o = batch.evaluate(t, row, 0)
        if (o.kind === 'ready') { const truth = cfTruth(o.value, { row, col: 0 }, new Set(['boolean'])); painted[t]![row] = truth.kind === 'ready' && truth.value === true }
        else refused.push(o.kind === 'unresolved' ? o.reason : o.kind)
      }
      ceilingUsed += batch.ceiling.used
      batch.dispose()
    }
    session.dispose()
    return { source, calls, scanBudgets, painted, refused, snapshotEntries, ceilingUsed, callUnits, hostDuringPrepare, memory }
  }
  it('paints exactly the engine\'s per-cell marks, every call <= 128 units, the scan on its own budget, the column scanned once per rule', async () => {
    const run = await paint(true)
    expect(run.refused).toEqual([])
    for (let t = 0; t < 2; t++) for (const target of batches) for (let row = target.firstRow; row <= target.lastRow; row++) {
      expect(run.painted[t]![row], `${texts[t]} at A${row + 1}`).toBe(expectation[t]!(row))
    }
    expect(run.painted[0]!.filter(Boolean).length).toBeGreaterThan(0)
    expect(run.painted[1]!.filter(Boolean).length).toBeGreaterThan(0)
    expect(Math.max(...run.calls.map(c => c.total))).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits)
    // The ceiling pays the snapshot and the calls only; the hoist scans pay from scan budgets (and yield at their checkpoints).
    expect(run.ceilingUsed).toBe(run.callUnits + 3 * 256)
    expect(run.scanBudgets.some(b => b.units >= ROWS)).toBe(true)
    expect(run.hostDuringPrepare).toBeGreaterThan(0)
    // The snapshot holds the per-cell footprint (256 criterion cells), never the 10,000-cell column.
    expect(run.snapshotEntries).toBe(256)
    const columnScans = run.source.pageCalls.filter(c => c.rects.length === 1 && c.rects[0]!.firstRow === 0 && c.rects[0]!.lastRow === ROWS - 1)
    expect(columnScans).toHaveLength(2) // one per rule, across all three batches (the session memoises the hoists)
    expect(run.memory.currentBytes).toBe(0)
  }, 60_000)
  it('the COUNTIF rule alone snapshots its own criterion cells (no other rule lends it coverage)', async () => {
    const source = fakeSource(cells)
    const batch = await batchOf([formulaRule(texts[0]!, `A2:A${ROWS}`)], source, { targets: [batches[1]!] })
    expect(batch.snapshotEntries).toBe(256)
    for (let row = batches[1]!.firstRow; row <= batches[1]!.lastRow; row++) {
      const o = batch.evaluate(0, row, 0)
      expect(o.kind, `A${row + 1}`).toBe('ready')
      if (o.kind === 'ready') expect('value' in o.value && o.value.value === true, `A${row + 1}`).toBe(expectation[0]!(row))
    }
    batch.dispose()
  }, 60_000)
  it('a hoist-free mutant refuses both rules (cf-resource-limit)', async () => {
    const run = await paint(false)
    expect(run.refused.length).toBe(2 * 3 * 256)
    expect(new Set(run.refused)).toEqual(new Set(['cf-resource-limit']))
  }, 60_000)
})

describe('the COUNTIF frequency map equals the engine\'s COUNTIF on admitted inputs (S5-HOIST property test, round-3 #5)', () => {
  const NUMBERS = [-2, 0, 1, 1.5, 3, 7, 10]
  const TEXTS = ['a', 'A', 'ab', 'AB', 'aB', 'hello', 'HELLO', 'x y', 'q', 'Zz']
  const ADMITTED: readonly CellInput[] = [...NUMBERS, ...TEXTS, null, { error: '#N/A' }, { error: '#DIV/0!' }]
  const TEXT = 'COUNTIF($A$1:$A$30,B1)'
  const compare = async (cells: Cells, criteria: readonly CellInput[]) => {
    const { plan, out, session } = await hoistOf(TEXT, cells)
    const map = hoisted(out).frequencies.get(plan[0]!.node)!
    const results: Array<{ kind: 'number' | 'text'; ok: boolean }> = []
    for (const criterion of criteria) {
      const key = cfCountifKey(literalOf(criterion))
      if (key === 'refuse') continue
      const engine = engineEval(TEXT, { ...cells, B1: criterion }, a1('C1'))
      expect(engine.features).toEqual([])
      results.push({ kind: typeof criterion === 'number' ? 'number' : 'text', ok: (map.get(key) ?? 0) === engine.value })
    }
    session.dispose()
    return results
  }
  it('at least 500 successful comparisons each for number and for text criteria, and fixed cases', async () => {
    const r = rng(34)
    let numbers = 0, texts = 0
    for (let trial = 0; trial < 120; trial++) {
      const cells: Cells = Object.fromEntries(Array.from({ length: 30 }, (_, i) => [`A${i + 1}`, pick(r, ADMITTED)]))
      const criteria = Array.from({ length: 10 }, () => (r() < 0.5 ? pick(r, NUMBERS) : pick(r, TEXTS)))
      for (const result of await compare(cells, criteria)) {
        expect(result.ok).toBe(true)
        if (result.kind === 'number') numbers++; else texts++
      }
    }
    expect(numbers).toBeGreaterThanOrEqual(500)
    expect(texts).toBeGreaterThanOrEqual(500)
    // Fixed: an empty range, a range of only errors, duplicates differing only in case.
    for (const cells of [{}, { A1: { error: '#N/A' }, A2: { error: '#DIV/0!' } }, { A1: 'abc', A2: 'ABC', A3: 'Abc', A4: 'abd' }] as Cells[]) {
      for (const result of await compare(cells, ['abc', 'ABC', 1, 'zz'])) expect(result.ok).toBe(true)
    }
    const dup = await hoistOf(TEXT, { A1: 'abc', A2: 'ABC', A3: 'Abc' })
    expect(hoisted(dup.out).frequencies.get(dup.plan[0]!.node)!.get(cfCountifKey({ kind: 'string', value: 'aBc' }) as string)).toBe(3)
  }, 60_000)
  it('a mutant map that admits a wildcard criterion fails: "a*" counts prefixes in the engine, a key lookup counts none', async () => {
    const cells: Cells = { A1: 'abc', A2: 'ab', A3: 'b' }
    expect(cfCountifKey({ kind: 'string', value: 'a*' })).toBe('refuse')
    const { plan, out } = await hoistOf(TEXT, cells)
    const mutantCount = hoisted(out).frequencies.get(plan[0]!.node)!.get('s:a*') ?? 0
    expect(engineEval(TEXT, { ...cells, B1: 'a*' }, a1('C1')).value).toBe(2)
    expect(mutantCount).not.toBe(2)
  })
})

describe('COUNTIF hoist policy refusals (S5-HOIST, round-3 #1, #5)', () => {
  const lit = (value: string): CfLiteral => ({ kind: 'string', value })
  it('criteria refused per cell: "#N/A" over a range holding #N/A, 256 characters, the matcher DP policy, " 1 " and "1"', async () => {
    expect(cfCountifKey(lit('#N/A'))).toBe('refuse')
    expect(cfCountifKey(lit('a'.repeat(CONDITIONAL_MAX_CRITERION_LENGTH + 1)))).toBe('refuse') // conditional-long-criterion-unverified
    expect(cfCountifKey(lit('a'.repeat(CONDITIONAL_MAX_CRITERION_LENGTH)))).not.toBe('refuse')
    const c123 = lit('a'.repeat(123))
    expect(cfCountifKey(c123)).not.toBe('refuse')
    expect(cfCountifDpRefused(c123, 32_767)).toBe(true) // (123 + 1) x (32,767 + 1) > DEFAULT_MAX_DP_CELLS: conditional-wildcard-resource-limit
    expect(cfCountifDpRefused(c123, 32_000)).toBe(false)
    for (const value of [' 1 ', '1', '', '>3', 'a?', '~a']) expect(cfCountifKey(lit(value)), JSON.stringify(value)).toBe('refuse')
    for (const literal of [{ kind: 'blank' }, { kind: 'boolean', value: true }, { kind: 'error', value: '#N/A' }] as CfLiteral[]) expect(cfCountifKey(literal)).toBe('refuse')
    // Through the evaluator, each refusal is the cell's own outcome: never a lookup.
    const cells: Cells = { A1: { error: '#N/A' }, A2: 'x'.repeat(32_767), A3: 'b', B1: '#N/A', B2: 'a'.repeat(123), B3: ' 1 ', B4: 'b' }
    const batch = await batchOf([formulaRule('COUNTIF($A$1:$A$3,B1)>0', 'C1:C4')], fakeSource(cells))
    for (const row of [0, 2]) expect(batch.evaluate(0, row, 2), `C${row + 1}`).toEqual({ kind: 'unresolved', reason: 'cf-formula-unsupported' })
    // The 123-unit criterion is refused either way: its text units alone push the call past 128 before the DP check.
    expect(batch.evaluate(0, 1, 2)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(batch.evaluate(0, 3, 2)).toEqual({ kind: 'ready', value: { kind: 'boolean', value: true } })
    batch.dispose()
    // Where the bound admits the call, the DP policy is what refuses: 50 characters against a 100,000-character string.
    const dp = await batchOf([formulaRule('COUNTIF($A$1:$A$2,B1)>0', 'C1')], fakeSource({ A1: 'y'.repeat(100_000), A2: 'b', B1: 'a'.repeat(50) }))
    expect(cfCountifDpRefused(lit('a'.repeat(50)), 100_000)).toBe(true)
    expect(dp.evaluate(0, 0, 2)).toEqual({ kind: 'unresolved', reason: 'cf-formula-unsupported' })
    dp.dispose()
  })
  it('a range holding " 1 ", a boolean or non-ASCII text refuses the whole rule', async () => {
    for (const bad of [' 1 ', true, 'café'] as CellInput[]) {
      const { out } = await hoistOf('COUNTIF($A$1:$A$3,B1)', { A1: 'a', A2: bad, A3: 2 })
      expect(out, JSON.stringify(bad)).toEqual({ kind: 'unresolved', reason: 'cf-formula-unsupported' })
    }
  })
})

describe('hoisted scalar aggregates equal the engine\'s own evaluation on admitted inputs (S5-HOIST property test, round-3 #3, #5)', () => {
  const MIXED: readonly CellInput[] = [1, -2.5, 1e15, 0, 7, '3', ' 4 ', 'ab', 'AB', '', true, false, null, { error: '#N/A' }, { error: '#DIV/0!' }, { date: 46332 }]
  const ADMITTED: readonly CellInput[] = [1, -2.5, 0, 3, 7, 'ab', 'AB', 'q', '', null, { error: '#N/A' }]
  const profiles: Array<{ name: string; text: string; pool: readonly CellInput[] }> = [
    ...['SUM', 'COUNT', 'COUNTA', 'AVERAGE', 'MIN', 'MAX'].map(fn => ({ name: fn, text: `${fn}($A$1:$A$12)`, pool: MIXED })),
    { name: 'COUNTIF literal', text: 'COUNTIF($A$1:$A$12,"ab")', pool: ADMITTED },
    { name: 'COUNTIF absolute ref', text: 'COUNTIF($A$1:$A$12,$C$1)', pool: ADMITTED },
    { name: 'SUMIF literal', text: 'SUMIF($A$1:$A$12,3)', pool: ADMITTED },
    { name: 'SUMIF absolute ref', text: 'SUMIF($A$1:$A$12,$C$1,$A$1:$A$12)', pool: ADMITTED },
  ]
  it('at least 300 successful comparisons per profile, mixed kinds, empty and all-error ranges included', async () => {
    const r = rng(3403)
    for (const profile of profiles) {
      let ok = 0
      for (let trial = 0; trial < 360; trial++) {
        const shape = trial % 12
        const cells: Cells = Object.fromEntries(Array.from({ length: 12 }, (_, i) => [`A${i + 1}`,
          shape === 0 ? null : shape === 1 ? { error: '#DIV/0!' } : pick(r, profile.pool)]))
        const withCriterion: Cells = { ...cells, C1: pick(r, [3, 7, 'ab', 'q', 'AB']) }
        const { plan, out, session } = await hoistOf(profile.text, withCriterion)
        expect(plan, profile.name).toHaveLength(1)
        if (out.kind !== 'ready') continue // refused inputs are covered by the policy tests
        const value = hoisted(out).scalars.get(plan[0]!.node)
        const engine = engineEval(profile.text, withCriterion, a1('Z1'))
        expect(engine.features, profile.name).toEqual([])
        expect(value, `${profile.name} ${JSON.stringify(cells)}`).toEqual(engine.value)
        ok++
        session.dispose()
      }
      expect(ok, profile.name).toBeGreaterThanOrEqual(300)
    }
  }, 120_000)
  it('fixed cases: [1e16, 1, -1e16] sums to 0 (the compensated statistics sum is 1), overflow, the first error, empty ranges', async () => {
    const seq: Cells = { A1: 1e16, A2: 1, A3: -1e16 }
    const sum = await hoistOf('SUM($A$1:$A$3)', seq)
    expect(hoisted(sum.out).scalars.get(sum.plan[0]!.node)).toBe(0)
    expect(engineEval('SUM($A$1:$A$3)', seq, a1('Z1')).value).toBe(0)
    const stats = await createCfStatistics(createCfMemory()).prepareStatistics([[rectOf('A1:A3')]], fakeSource(seq).stamp('s1'), fakeSource(seq), scan())
    expect(stats[0]!.kind === 'ready' && stats[0]!.value.value.sum).toBe(1) // a mutant reusing it would publish 1
    const over: Cells = { A1: 1e308, A2: 1e308 }
    const o = await hoistOf('SUM($A$1:$A$2)', over)
    expect(hoisted(o.out).scalars.get(o.plan[0]!.node)).toEqual(engineEval('SUM($A$1:$A$2)', over, a1('Z1')).value)
    expect(hoisted(o.out).scalars.get(o.plan[0]!.node)).toEqual({ kind: 'formula-error', code: '#NUM!' })
    const errs: Cells = { A1: 1, A2: { error: '#DIV/0!' }, A3: 2, A4: { error: '#N/A' } }
    const e = await hoistOf('SUM($A$1:$A$4)', errs)
    expect(hoisted(e.out).scalars.get(e.plan[0]!.node)).toEqual({ kind: 'formula-error', code: '#DIV/0!' })
    for (const fn of ['SUM', 'COUNT', 'COUNTA', 'AVERAGE', 'MIN', 'MAX']) {
      const empty = await hoistOf(`${fn}($A$1:$A$5)`, { B9: 1 })
      expect(hoisted(empty.out).scalars.get(empty.plan[0]!.node), fn).toEqual(engineEval(`${fn}($A$1:$A$5)`, { B9: 1 }, a1('Z1')).value)
    }
  })
  it('a mutant that ignores booleans, and one that ignores errors, each fail', async () => {
    const cells: Cells = { A1: true, A2: 1, A3: { error: '#N/A' } }
    const counta = await hoistOf('COUNTA($A$1:$A$3)', cells)
    expect(hoisted(counta.out).scalars.get(counta.plan[0]!.node)).toBe(engineEval('COUNTA($A$1:$A$3)', cells, a1('Z1')).value)
    expect(engineEval('COUNTA($A$1:$A$3)', cells, a1('Z1')).value).toBe(3) // ignoring the boolean would give 2
    const sum = await hoistOf('SUM($A$1:$A$3)', cells)
    expect(hoisted(sum.out).scalars.get(sum.plan[0]!.node)).toEqual({ kind: 'formula-error', code: '#N/A' }) // ignoring errors would give 1
  })
})

describe('a cancelled or refused hoist leaves nothing charged (S5-HOIST)', () => {
  const column = (n: number, value: (r: number) => CellInput): Cells => Object.fromEntries(Array.from({ length: n }, (_, i) => [`A${i + 1}`, value(i + 1)]))
  it('an abort mid-scan publishes nothing, returns the ledger to baseline, and a later scan starts over', async () => {
    const cells = column(2000, r => r % 50)
    const controller = new AbortController(), memory = createCfMemory()
    const source = fakeSource(cells)
    hostHooks.onHost = () => controller.abort()
    const rule = formulaRule('COUNTIF($A$1:$A$2000,B1)', 'B1')
    const plan = planCfHoist(formulaOf(rule))
    const session = createCfStatistics(memory)
    await expect(session.prepareHoists(rule, plan, source.stamp('s1'), source, scan(controller.signal), { signal: controller.signal })).rejects.toMatchObject({ code: 'cancelled' })
    expect(memory.currentBytes).toBe(0)
    hostHooks.onHost = undefined
    const calls = source.pageCalls.length
    const again = await session.prepareHoists(rule, plan, source.stamp('s1'), source, scan(), {})
    expect(again.kind).toBe('ready')
    expect(source.pageCalls.length).toBe(calls + 1)
    session.dispose()
    expect(memory.currentBytes).toBe(0)
  })
  it('a frequency-map growth one byte over its reservation (old plus new table at the peak) refuses with nothing held', async () => {
    const keys = (n: number) => column(n, r => `k${String.fromCharCode(96 + r)}`) // n distinct ASCII text keys
    const sixteen = await hoistOf('COUNTIF($A$1:$A$17,B1)', keys(16))
    const peak16 = sixteen.memory.peakBytes
    sixteen.session.dispose()
    // The 17th distinct key grows the 16-entry table to 32: old plus new are held at once.
    const atGrowth = peak16 + 32 * CF_FREQUENCY_ENTRY_BYTES
    const short = await hoistOf('COUNTIF($A$1:$A$17,B1)', keys(17), { memory: createCfMemory({ limitBytes: atGrowth - 1 }) })
    expect(short.out).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(short.memory.currentBytes).toBe(0)
    const fits = await hoistOf('COUNTIF($A$1:$A$17,B1)', keys(17), { memory: createCfMemory({ limitBytes: atGrowth + 2 * 'kq'.length + 2 }) })
    expect(fits.out.kind).toBe('ready')
    expect(fits.memory.peakBytes).toBeGreaterThanOrEqual(atGrowth)
    fits.session.dispose()
    expect(fits.memory.currentBytes).toBe(0)
  })
  it('a host-sheet commit and a referenced-sheet commit each invalidate the cached hoist and every tile of the sheet', async () => {
    const cells = column(20, r => r % 4)
    const source = fakeSource(cells)
    const memory = createCfMemory()
    const session = createCfStatistics(memory)
    const rule = formulaRule('COUNTIF($A$1:$A$20,B1)', 'B1')
    const plan = planCfHoist(formulaOf(rule))
    const hoistNow = () => session.prepareHoists(rule, plan, source.stamp('s1'), source, scan(), {})
    await hoistNow()
    const scans = () => source.pageCalls.length
    const first = scans()
    await hoistNow()
    expect(scans()).toBe(first) // memoised
    const held = memory.currentBytes
    source.valueGen++
    await hoistNow()
    expect(scans()).toBe(first + 1) // host-sheet commit: rescanned
    expect(memory.currentBytes).toBe(held) // the old generation's entry was released
    source.dependencyGen = 5
    await hoistNow()
    expect(scans()).toBe(first + 2) // referenced-sheet commit: rescanned
    session.dispose()
    // Tiles: a prepared batch goes stale synchronously at either commit.
    const tileSource = fakeSource({ A1: 5 })
    const cellIs: CfRule = { id: 1, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [rectOf('A1')], writes: new Set(['font']), dxf: { 'font.bold': true },
      definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 1 }], dialect: 'base', formulas: [] } }
    for (const commit of [() => { tileSource.valueGen++ }, () => { tileSource.dependencyGen++ }]) {
      const service = createCfService({ rules: [cellIs], diagnostics: [] }, createCfMemory(), { evidence: { allows: () => true } })
      const lease = await service.prepareCf({ stamp: tileSource.stamp('s1'), sourceRects: [rectOf('A1')], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, tileSource)
      expect(lease.value.at(0, 0).kind).toBe('ready')
      commit()
      expect(lease.value.at(0, 0)).toEqual({ kind: 'unresolved', reason: 'cf-stale-generation' })
      lease.release(); service.dispose()
    }
  })
})

describe('the service statistics loop runs hoists beside statistical rules only with their evidence (S5-HOIST)', () => {
  const cells: Cells = Object.fromEntries(Array.from({ length: 50 }, (_, i) => [`A${i + 1}`, i % 9]))
  async function run(evidence: (cap: string) => boolean, evaluation = true) {
    const source = fakeSource(cells)
    const rule = formulaRule('COUNTIF($A:$A,A2)>1', 'A2:A50')
    const service = createCfService({ rules: [rule], diagnostics: [] }, createCfMemory(), { evidence: { allows: evidence }, ...(evaluation ? { evaluation: EVALUATION } : {}) })
    const lease = await service.prepareCf({ stamp: source.stamp('s1'), sourceRects: [rectOf('A2:A50')], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
    const out = lease.value.at(1, 0)
    lease.release(); service.dispose()
    return { columnScans: source.pageCalls.filter(c => c.rects.length === 1 && c.rects[0]!.firstRow === 0 && c.rects[0]!.lastRow === 49).length, out }
  }
  it('scans the column once with the capability and a context; production (no capability) and a missing context scan nothing', async () => {
    expect((await run(() => true)).columnScans).toBe(1)
    expect((await run(cap => !cap.startsWith('formula:hoist'))).columnScans).toBe(0)
    expect((await run(() => true, false)).columnScans).toBe(0)
    // The per-cell formula arm is still Task 35's: the rule stays withheld whatever was hoisted.
    expect((await run(() => true)).out).toMatchObject({ kind: 'ready' })
  })
})
void refOf; void compile; void createCfUnitCeiling
