// @vitest-environment node
/**
 * Stage 5 Task 34, Ruling S5-N (round-1 #2, #3; round-2 #1, #2; round-3 #6; round-4 #1, #2): every per-cell call's work
 * is bounded statically before invocation (<= CF_LIMITS.checkpointUnits) and charged before it happens, traversal per
 * occurrence; evaluator scratch and snapshot growth are reserved before use; the loop yields between cells only.
 * Every number below is the brief's (task-34-brief.md), and each is checked component by component against the
 * breakdown the evaluator reports. Each mutant is built in the test (a non-charging seam), never by editing source.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import '../src/xlsx/formula/functions'
import * as evaluatorModule from '../src/xlsx/formula/evaluator'
import * as schedulerModule from '../src/xlsx/view/scheduler'
import {
  cfCallShape, cfCallWorkBreakdown, cfEvaluatorScratchBytes, cfFormulaSemantics, createCfInvariantMemo, evaluateCfFormula,
  type CfCallWorkBreakdown, type CfFormulaContext,
} from '../src/xlsx/cf/formula'
import { CF_DEPENDENCY_COLUMN_BYTES, prepareCfDependencies, type CfDependencySnapshot } from '../src/xlsx/cf/dependencies'
import { CF_DEPENDENCY_CELL_BYTES, CF_FORMULA_NODE_BYTES, CF_LIMITS, createCfMemory, createCfUnitCeiling, createCfWorkBudget, type CfUnitCeiling, type CfWorkBudget } from '../src/xlsx/cf/limits'
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { EVALUATION, a1, batchOf, compile, fakeSource, formulaRule, rectOf, refOf, type Cells } from './helpers/cf-formula'

vi.mock('../src/xlsx/formula/evaluator', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/formula/evaluator')>()
  return { ...actual, evaluateNode: vi.fn(actual.evaluateNode), textChars: vi.fn(actual.textChars) }
})
const hostHooks: { onHost?: () => void } = {}
vi.mock('../src/xlsx/view/scheduler', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/view/scheduler')>()
  return { ...actual, hostTask: vi.fn(async () => { hostHooks.onHost?.() }) }
})
const evaluateNodeSpy = vi.mocked(evaluatorModule.evaluateNode)
const textCharsSpy = vi.mocked(evaluatorModule.textChars)
const hostTaskSpy = vi.mocked(schedulerModule.hostTask)
afterEach(() => { hostHooks.onHost = undefined; evaluateNodeSpy.mockClear(); textCharsSpy.mockClear(); hostTaskSpy.mockClear() })

const frozen = () => 0
const bigBudget = (signal = new AbortController().signal): CfWorkBudget => createCfWorkBudget({ now: frozen, signal, scan: true, maxUnits: 1e9 })
/** A ceiling that never charges: the mutant "charging removed". */
const freeCeiling = (): CfUnitCeiling => ({ limit: Infinity, used: 0, remaining: Infinity, take: () => true })
const column = (prefix: string, from: number, to: number, value: (r: number) => number | string): Cells =>
  Object.fromEntries(Array.from({ length: to - from + 1 }, (_, i) => [`${prefix}${from + i}`, value(from + i)]))

describe('relative-range traversal work is charged per occurrence before it happens (S5-N, round-1 #2, round-3 #6, round-4 #2)', () => {
  const cells: Cells = { A1: 1, A2: 2, A3: 3, A4: 4 }
  async function sumTwice(ceiling: (snapshotUnits: number) => CfUnitCeiling) {
    const calls: CfCallWorkBreakdown[] = []
    const sum = formulaRule('SUM(A1:A4)+SUM(A1:A4)', 'B1'), cheap = formulaRule('A1>0', 'C1')
    const batch = await batchOf([sum, cheap], fakeSource(cells), { ceiling: ceiling(4), onCall: b => calls.push(b) })
    evaluateNodeSpy.mockClear()
    const out = [batch.evaluate(0, 0, 1), batch.evaluate(1, 0, 2)]
    const sumRan = evaluateNodeSpy.mock.calls.length > 1
    batch.dispose()
    return { out, calls, sumRan }
  }
  it('SUM(A1:A4)+SUM(A1:A4) on a 4-cell view: 5 nodes + 4 view build + 2 x (4 enumeration + 4 reads + 8 sort) = 41', async () => {
    const { out, calls } = await sumTwice(units => createCfUnitCeiling(units + 1000))
    expect(out[0]).toEqual({ kind: 'ready', value: { kind: 'number', value: 20 } })
    expect(calls[0]).toEqual({ nodes: 5, viewBuild: 4, enumeration: 8, reads: 8, sort: 16, comparisons: 0, matcher: 0, text: 0, lookups: 0, total: 41 })
  })
  it('with 40 units left the call never runs and the rule is unresolved, while a second rule in the batch still evaluates', async () => {
    const { out, calls, sumRan } = await sumTwice(units => createCfUnitCeiling(units + 40)) // the snapshot build takes 4 (A1:A4)
    expect(out[0]).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(sumRan).toBe(false)
    expect(out[1]).toEqual({ kind: 'ready', value: { kind: 'boolean', value: true } })
    expect(calls.map(c => c.total)).toEqual([5]) // only the second rule was admitted: 3 nodes + 1 view + 1 read
    // Mutant (charging removed): the 41-unit call runs on 40, so the refusal above would fail.
    const mutant = await sumTwice(() => freeCeiling())
    expect(mutant.out[0]).toEqual({ kind: 'ready', value: { kind: 'number', value: 20 } })
  })
  it('COUNTIF(A$1:A$4,B2) over 1,000 targets: a 1,004-cell snapshot, 37 per call, so 108 targets evaluate under a 5,000 ceiling', async () => {
    const data: Cells = { ...column('A', 1, 4, r => r), ...column('B', 2, 1001, r => r % 7) }
    const run = async (ceiling: CfUnitCeiling) => {
      const calls: CfCallWorkBreakdown[] = []
      const batch = await batchOf([formulaRule('COUNTIF(A$1:A$4,B2)', 'B2:B1001')], fakeSource(data), { ceiling, onCall: b => calls.push(b) })
      const snapshotUnits = batch.ceiling.used, entries = batch.snapshotEntries
      const outcomes: Array<{ kind: string }> = []
      await batch.run(Array.from({ length: 1000 }, (_, i) => ({ row: i + 1, col: 1 })), (_t, _r, _c, o) => outcomes.push(o))
      batch.dispose()
      return { calls, snapshotUnits, entries, outcomes }
    }
    const { calls, snapshotUnits, entries, outcomes } = await run(createCfUnitCeiling(5000))
    expect(entries).toBe(1004)
    expect(snapshotUnits).toBe(1004) // built once, one unit per entry
    expect(calls[0]).toEqual({ nodes: 3, viewBuild: 5, enumeration: 5, reads: 5, sort: 15, comparisons: 4, matcher: 0, text: 0, lookups: 0, total: 37 })
    expect(Math.floor((5000 - 1004) / 37)).toBe(108)
    const readyRows = outcomes.map((o, i) => (o.kind === 'ready' ? i : -1)).filter(i => i >= 0)
    expect(readyRows).toEqual(Array.from({ length: 108 }, (_, i) => i)) // B2:B109, in row-major order
    expect(outcomes.slice(108).every(o => o.kind === 'unresolved' && (o as unknown as { reason: string }).reason === 'cf-resource-limit')).toBe(true)
    // Mutant (charging removed): every target evaluates.
    const mutant = await run(freeCeiling())
    expect(mutant.outcomes.filter(o => o.kind === 'ready')).toHaveLength(1000)
  })
})

describe('full-store enumeration is counted per traversal occurrence (round-4 #1)', () => {
  it('4 x SUM(A1:A6) on a 6-cell view is 11 + 24 + 24 + 72 + 6 = 137 > 128: refused before invocation', async () => {
    const text = 'SUM(A1:A6)+SUM(A1:A6)+SUM(A1:A6)+SUM(A1:A6)'
    const f = compile(text, 'B1')
    const metrics = { viewCells: 6, atoms: Array.from({ length: 4 }, () => ({ cells: 6, stringUnits: 0, longestString: 0 })) }
    const b = cfCallWorkBreakdown(cfCallShape(f), metrics, []) as CfCallWorkBreakdown
    expect(b).toMatchObject({ nodes: 11, reads: 24, enumeration: 24, sort: 72, viewBuild: 6, total: 137 })
    // The mutant that omits the enumeration term computes 113 and would admit it.
    expect(b.total - b.enumeration).toBe(113)
    expect(b.total - b.enumeration).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits)
    const batch = await batchOf([formulaRule(text, 'B1')], fakeSource(column('A', 1, 6, r => r)))
    evaluateNodeSpy.mockClear()
    expect(batch.evaluate(0, 0, 1)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled() // no engine call, so no stored() visit (only the engine enumerates)
    batch.dispose()
  })
})

describe('each synchronous call stays within one checkpoint and the loop yields between cells (round-2 #2)', () => {
  const data = column('A', 1, 16, r => r)
  it('a per-call bound of 128 evaluates; 129 is refused and evaluateNode is never called', async () => {
    const at128 = 'SUM(A1:A16)+1+1+1+1+1+1+1', at129 = 'SUM(A1:A16)+1+1+1+1+1+1+-1'
    const calls: CfCallWorkBreakdown[] = []
    const batch = await batchOf([formulaRule(at128, 'B1'), formulaRule(at129, 'C1')], fakeSource(data), { onCall: c => calls.push(c) })
    expect(batch.evaluate(0, 0, 1)).toEqual({ kind: 'ready', value: { kind: 'number', value: 143 } })
    expect(calls.map(c => c.total)).toEqual([128])
    evaluateNodeSpy.mockClear()
    expect(batch.evaluate(1, 0, 2)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    batch.dispose()
  })
  it('over 1,000 targets hostTask is awaited between cells on every yield, never inside a call', async () => {
    const cells = column('A', 1, 1000, r => r)
    let depth = 0, insideCall = 0, yields = 0
    evaluateNodeSpy.mockImplementation(((node: never, ctx: never) => {
      depth++
      try { return evaluatorActual.evaluateNode(node, ctx) } finally { depth-- }
    }) as never)
    try {
      const signal = new AbortController().signal
      const base = bigBudget(signal)
      const budget: CfWorkBudget = { get units() { return base.units }, spend: n => { const r = base.spend(n); if (r === 'yield') yields++; return r }, resume: () => base.resume() }
      const batch = await batchOf([formulaRule('A1>0', 'B1:B1000')], fakeSource(cells), { budget, signal })
      hostTaskSpy.mockClear(); yields = 0
      hostHooks.onHost = () => { if (depth > 0) insideCall++ }
      let results = 0
      await batch.run(Array.from({ length: 1000 }, (_, i) => ({ row: i, col: 1 })), () => { results++ })
      batch.dispose()
      expect(results).toBe(1000)
      expect(yields).toBeGreaterThan(10)
      expect(hostTaskSpy).toHaveBeenCalledTimes(yields)
      expect(insideCall).toBe(0)
    } finally { evaluateNodeSpy.mockImplementation(evaluatorActual.evaluateNode as never) }
  })
  it('an abort fired from the 10th host task stops the preparation before the next call, throws cancelled, and frees the ledger', async () => {
    const cells = column('A', 1, 1000, r => r)
    const controller = new AbortController(), memory = createCfMemory()
    const batch = await batchOf([formulaRule('A1>0', 'B1:B1000')], fakeSource(cells), { budget: bigBudget(controller.signal), signal: controller.signal, memory })
    let hosts = 0, callsAfterAbort = 0
    hostTaskSpy.mockClear()
    hostHooks.onHost = () => { if (++hosts === 10) controller.abort() }
    const before = evaluateNodeSpy.mock.calls.length
    let afterAbortStart = -1
    const run = batch.run(Array.from({ length: 1000 }, (_, i) => ({ row: i, col: 1 })), () => { if (controller.signal.aborted) callsAfterAbort++ })
    await expect(run).rejects.toMatchObject({ code: 'cancelled' })
    afterAbortStart = evaluateNodeSpy.mock.calls.length
    expect(hosts).toBe(10)
    expect(callsAfterAbort).toBe(0)
    expect(afterAbortStart).toBeGreaterThan(before)
    expect(memory.currentBytes).toBe(0)
  })
})

describe('text work is bounded by measured payload before invocation (round-2 #1, round-3 #6)', () => {
  async function snapshotOf(cells: Cells, rects: string[], memory = createCfMemory()) {
    const signal = new AbortController().signal
    const out = await prepareCfDependencies(rects.map(r => ({ rect: rectOf(r) })), fakeSource(cells), undefined, bigBudget(signal), memory, signal, { sheetId: 's1' })
    if (out.kind !== 'ready') throw new Error('snapshot refused')
    return out.value.value
  }
  const context = (memory = createCfMemory(), calls?: CfCallWorkBreakdown[]): CfFormulaContext => ({
    sheetName: 'Sheet1', evaluation: EVALUATION, semantics: cfFormulaSemantics(EVALUATION), edge: 'ref-error', materialization: { limit: 1e6, used: 0 },
    budget: bigBudget(), memory, invariants: createCfInvariantMemo(createCfMemory()), ceiling: createCfUnitCeiling(), ...(calls ? { onCall: b => calls.push(b) } : {}),
  })
  const metricsOf = (snapshot: CfDependencySnapshot, rects: string[]) => {
    const view = snapshot.viewFor(rects.map(r => ({ rect: rectOf(r) })), bigBudget())
    if (view.kind !== 'ready') throw new Error('view refused')
    return view.value.metrics
  }
  it('LEN(A1&A2) over two 32,767-character strings is refused from its payload whatever memory is available: no engine or textChars call', async () => {
    const big = 'x'.repeat(32_767)
    const snapshot = await snapshotOf({ A1: big, A2: big }, ['A1', 'A2'])
    const f = compile('LEN(A1&A2)', 'B1')
    const b = cfCallWorkBreakdown(cfCallShape(f), metricsOf(snapshot, ['A1', 'A2']), []) as CfCallWorkBreakdown
    expect(b.text).toBe(65_534)
    expect(b.total - b.text).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits) // a mutant without the text term admits it
    evaluateNodeSpy.mockClear(); textCharsSpy.mockClear()
    expect(evaluateCfFormula(f, a1('B1'), snapshot, context(createCfMemory({ limitBytes: Number.MAX_SAFE_INTEGER })))).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    expect(textCharsSpy).not.toHaveBeenCalled()
    snapshot.dispose()
  })
  it('arithmetic coercion and AND/OR/NOT charge the string they consume: over 32,767 characters, and just over the cap, each withholds (Task 34 fix 3)', async () => {
    for (const text of ['A1+1>0', 'AND(A1)', 'OR(A1,TRUE)', 'NOT(A1)', '-A1<0', 'ABS(A1)>0']) {
      const f = compile(text, 'B1')
      for (const length of [32_767, 125]) {
        const snapshot = await snapshotOf({ A1: 'x'.repeat(length) }, ['A1'])
        const b = cfCallWorkBreakdown(cfCallShape(f), metricsOf(snapshot, ['A1']), []) as CfCallWorkBreakdown
        expect(b.text, `${text} over ${length}`).toBe(length)
        expect(b.total, `${text} over ${length}`).toBeGreaterThan(CF_LIMITS.checkpointUnits)
        // A mutant that charges text only for text functions and & would admit it.
        expect(b.total - b.text, `${text} over ${length}`).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits)
        evaluateNodeSpy.mockClear()
        expect(evaluateCfFormula(f, a1('B1'), snapshot, context()), `${text} over ${length}`).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
        expect(evaluateNodeSpy).not.toHaveBeenCalled()
        snapshot.dispose()
      }
    }
  })
  it('LEN(A1&A2) over two 10-character strings is 4 + 2 + 2 + 20 = 28, and runs at its exact scratch bound but never one byte below', async () => {
    const snapshot = await snapshotOf({ A1: 'abcdefghij', A2: 'klmnopqrst' }, ['A1', 'A2'])
    const f = compile('LEN(A1&A2)', 'B1')
    const metrics = metricsOf(snapshot, ['A1', 'A2'])
    const scratch = cfEvaluatorScratchBytes(f, metrics)
    const tree = f.nodes * CF_FORMULA_NODE_BYTES // the translated tree, reserved separately before it is built
    const calls: CfCallWorkBreakdown[] = []
    evaluateNodeSpy.mockClear()
    expect(evaluateCfFormula(f, a1('B1'), snapshot, context(createCfMemory({ limitBytes: scratch - 1 })))).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    const exact = createCfMemory({ limitBytes: scratch + tree })
    expect(evaluateCfFormula(f, a1('B1'), snapshot, context(exact, calls))).toEqual({ kind: 'ready', value: { kind: 'number', value: 20 } })
    expect(calls[0]).toEqual({ nodes: 4, viewBuild: 2, enumeration: 0, reads: 2, sort: 0, comparisons: 0, matcher: 0, text: 20, lookups: 0, total: 28 })
    expect(exact.peakBytes).toBe(scratch + tree)
    expect(exact.currentBytes).toBe(0)
    snapshot.dispose()
  })
  it('COUNTIF(A1:A2,B1): 2 x 11 x 11 = 242 matcher transitions over 10-character strings refuses; over 3 characters it is 50', async () => {
    const f = compile('COUNTIF(A1:A2,B1)', 'C1')
    const ten = await snapshotOf({ A1: 'abcdefghij', A2: 'abcdefghik', B1: 'abcdefghij' }, ['A1:A2', 'B1'])
    const b10 = cfCallWorkBreakdown(cfCallShape(f), metricsOf(ten, ['A1:A2', 'B1']), [10]) as CfCallWorkBreakdown
    expect(b10.matcher).toBe(242)
    expect(b10.total - b10.matcher).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits) // a mutant without the matcher term admits it
    evaluateNodeSpy.mockClear()
    expect(evaluateCfFormula(f, a1('C1'), ten, context())).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    const three = await snapshotOf({ A1: 'abc', A2: 'abd', B1: 'abc' }, ['A1:A2', 'B1'])
    const calls: CfCallWorkBreakdown[] = []
    expect(evaluateCfFormula(f, a1('C1'), three, context(createCfMemory(), calls))).toEqual({ kind: 'ready', value: { kind: 'number', value: 1 } })
    expect(calls[0]).toEqual({ nodes: 3, viewBuild: 3, enumeration: 3, reads: 3, sort: 6, comparisons: 0, matcher: 32, text: 0, lookups: 0, total: 50 })
    ten.dispose(); three.dispose()
  })
})

describe('evaluator scratch and snapshot growth are reserved before use (Lessons 9, 15, round-1 #3)', () => {
  it('one byte below cfEvaluatorScratchBytes(f, viewFor metrics): evaluateNode is never called and nothing is constructed', async () => {
    const signal = new AbortController().signal, memory = createCfMemory()
    const snap = await prepareCfDependencies([{ rect: rectOf('A1:A4') }], fakeSource({ A1: 1, A2: 2, A3: 3, A4: 4 }), undefined, bigBudget(signal), memory, signal, { sheetId: 's1' })
    if (snap.kind !== 'ready') throw new Error('snapshot')
    const snapshot = snap.value.value
    const f = compile('SUM(A1:A4)>2', 'B1')
    const view = snapshot.viewFor([{ rect: rectOf('A1:A4') }], bigBudget())
    if (view.kind !== 'ready') throw new Error('view')
    const scratch = cfEvaluatorScratchBytes(f, view.value.metrics)
    const short = createCfMemory({ limitBytes: scratch - 1 })
    evaluateNodeSpy.mockClear()
    const out = evaluateCfFormula(f, a1('B1'), snapshot, {
      sheetName: 'Sheet1', evaluation: EVALUATION, semantics: cfFormulaSemantics(EVALUATION), edge: 'ref-error', materialization: { limit: 1e6, used: 0 },
      budget: bigBudget(), memory: short, invariants: createCfInvariantMemo(createCfMemory()), ceiling: createCfUnitCeiling(),
    })
    expect(out).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    expect(short.peakBytes).toBe(0)
    snapshot.dispose()
    expect(memory.currentBytes).toBe(0)
  })
  it('a snapshot growth from 1,024 to 2,048 entries charges old plus new at the peak; one byte less refuses with nothing held', async () => {
    const cells = column('A', 1, 1025, r => r)
    const rect = rectOf('A1:A1025')
    const bounds = rectOf('A1:A1025')
    const base = measurePayloadBytes({ declared: [Object.freeze({ ...rect })], bounds })!
    const peak = base + CF_DEPENDENCY_COLUMN_BYTES + (1024 + 2048) * CF_DEPENDENCY_CELL_BYTES
    const build = async (limitBytes?: number) => {
      const memory = createCfMemory(limitBytes === undefined ? {} : { limitBytes }), signal = new AbortController().signal
      const out = await prepareCfDependencies([{ rect }], fakeSource(cells), undefined, bigBudget(signal), memory, signal, { sheetId: 's1' })
      return { out, memory }
    }
    const ok = await build()
    expect(ok.out.kind).toBe('ready')
    expect(ok.memory.peakBytes).toBe(peak)
    if (ok.out.kind === 'ready') {
      expect(ok.out.value.bytes).toBe(base + CF_DEPENDENCY_COLUMN_BYTES + 2048 * CF_DEPENDENCY_CELL_BYTES) // the new table only
      ok.out.value.value.dispose()
    }
    expect(ok.memory.currentBytes).toBe(0)
    const short = await build(peak - 1)
    expect(short.out).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(short.memory.currentBytes).toBe(0)
  })
  it('a refused shared snapshot is retried per rule, each attempt on a fresh budget, all drawn from one ceiling', async () => {
    const source = fakeSource({ A2: 1, C2: 1 })
    source.unavailable.add('A2') // only the first rule's footprint holds the unavailable cell
    const budgets: CfWorkBudget[] = []
    const ceiling = createCfUnitCeiling(10_000)
    const batch = await batchOf([formulaRule('A2>0', 'B2'), formulaRule('C2>0', 'D2')], source, {
      ceiling, attemptBudget: max => { const b = createCfWorkBudget({ now: frozen, scan: true, maxUnits: max }); budgets.push(b); return b },
    })
    expect(batch.attempts).toBe(3) // shared, then one per rule
    expect(new Set(budgets).size).toBe(3)
    expect(ceiling.used).toBe(budgets.reduce((n, b) => n + b.units, 0))
    expect(batch.evaluate(0, 1, 1)).toEqual({ kind: 'unresolved', reason: 'cf-formula-dependency' })
    expect(batch.evaluate(1, 1, 3)).toEqual({ kind: 'ready', value: { kind: 'boolean', value: true } })
    batch.dispose()
  })
  it('cancellation during the asynchronous snapshot build, and dispose, each return the ledger to baseline', async () => {
    const cells = column('A', 1, 2000, r => r)
    const controller = new AbortController(), memory = createCfMemory()
    let hosts = 0
    hostHooks.onHost = () => { if (++hosts === 3) controller.abort() }
    await expect(batchOf([formulaRule('A1>0', 'B1:B2000')], fakeSource(cells), { signal: controller.signal, memory })).rejects.toMatchObject({ code: 'cancelled' })
    expect(hosts).toBe(3)
    expect(memory.currentBytes).toBe(0)
    hostHooks.onHost = undefined
    const kept = createCfMemory()
    const batch = await batchOf([formulaRule('A1>0', 'B1:B2000')], fakeSource(cells), { memory: kept })
    expect(kept.currentBytes).toBeGreaterThan(0)
    batch.dispose()
    expect(kept.currentBytes).toBe(0)
  })
})

const evaluatorActual = await vi.importActual<typeof import('../src/xlsx/formula/evaluator')>('../src/xlsx/formula/evaluator')

describe('the static bound covers the store work the engine actually performs (S5-N, fail closed past it)', () => {
  const cells: Cells = { A1: 1, A2: 'abc', A3: 3, A4: { error: '#N/A' }, B1: 2, B2: 'abd', B3: 0, B4: 7, C1: 'abc', C2: null }
  // Formulas whose every store access is one of the bound's dynamic terms: the meter equals the dynamic part exactly.
  const exact = ['SUM(B1:B4)+SUM(B1:B4)', 'COUNTIF(A$1:A$4,B1)', 'COUNTIF(A1:A2,C1)', 'AVERAGE(A1:A3)+B1', 'LEN(A2&B2)', 'A2', 'A1:A3', 'COUNT(A1:A4)>B4']
  // Formulas the bound over-estimates (lazy branches, a reads bound of ranges x positions): the meter stays under it.
  // (SUM(A1:A4)+SUM(A1:A4): the left SUM meets #N/A in A4, so the engine never evaluates the right one.)
  const under = ['SUM(A1:A4)+SUM(A1:A4)', 'IF(A1>0,SUM(B1:B4),COUNT(A1:A4))', 'SUMIF(A1:A3,"abc",B1:B3)', 'COUNTIFS(A1:A4,">0",B1:B4,">0")', 'IFERROR(MAX(A1:A4),MIN(B1:B4))', 'SUMIF(C1:C2,"",B1:B2)']
  it('metered spend equals view + enumeration + sort + reads (+ lookups) on exact shapes, and stays under it otherwise', async () => {
    for (const text of [...exact, ...under]) {
      const settled: Array<[number, CfCallWorkBreakdown]> = []
      const batch = await batchOf([formulaRule(text, 'E1')], fakeSource(cells), { hoist: false, onSettled: (spent, b) => settled.push([spent, b]) })
      const out = batch.evaluate(0, 0, 4)
      batch.dispose()
      expect(out.kind, text).not.toBe('no-match')
      expect(settled, text).toHaveLength(1)
      const [spent, b] = settled[0]!
      const dynamic = b.viewBuild + b.enumeration + b.sort + b.reads + b.lookups
      if (exact.includes(text)) expect(spent, text).toBe(dynamic)
      else expect(spent, text).toBeLessThanOrEqual(dynamic)
      expect(spent, text).toBeGreaterThan(0)
    }
  })
})
void refOf
