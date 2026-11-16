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
import * as wildcardModule from '../src/xlsx/formula/functions/wildcard'
import {
  cfCallShape, cfCallWorkBreakdown, cfEvaluatorScratchBytes, cfFormulaSemantics, createCfInvariantMemo, evaluateCfFormula,
  type CfCallWorkBreakdown, type CfFormulaContext, type TextValue, CF_TEXT_RENDER_UNITS, CF_TEXT_UNIT_CHARS, cfTextUnits,
  CF_FREQUENCY_KEY_PASSES, cfCriterionUnits,
} from '../src/xlsx/cf/formula'
import { CF_DEPENDENCY_COLUMN_BYTES, prepareCfDependencies, type CfDependencySnapshot } from '../src/xlsx/cf/dependencies'
import { CF_DEPENDENCY_CELL_BYTES, CF_FORMULA_CALL_BYTES, CF_MATCHER_CELL_BYTES, CF_FORMULA_NODE_BYTES, CF_LIMITS, CF_VIEW_WORKSPACE_BYTES, type CfReservation, createCfMemory, createCfUnitCeiling, createCfWorkBudget, type CfUnitCeiling, type CfWorkBudget } from '../src/xlsx/cf/limits'
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { EVALUATION, a1, batchOf, compile, engineEval, fakeSource, formulaRule, rectOf, refOf, type CellInput, type Cells } from './helpers/cf-formula'
import type { AstNode } from '../src/xlsx/formula/types'

vi.mock('../src/xlsx/formula/evaluator', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/formula/evaluator')>()
  return { ...actual, evaluateNode: vi.fn(actual.evaluateNode), textChars: vi.fn(actual.textChars) }
})
const hostHooks: { onHost?: () => void } = {}
vi.mock('../src/xlsx/view/scheduler', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/view/scheduler')>()
  return { ...actual, hostTask: vi.fn(async () => { hostHooks.onHost?.() }) }
})
vi.mock('../src/xlsx/formula/functions/wildcard', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/formula/functions/wildcard')>()
  return { ...actual, tokenizeWildcardPattern: vi.fn(actual.tokenizeWildcardPattern) }
})
const evaluateNodeSpy = vi.mocked(evaluatorModule.evaluateNode)
const textCharsSpy = vi.mocked(evaluatorModule.textChars)
const hostTaskSpy = vi.mocked(schedulerModule.hostTask)
afterEach(() => { hostHooks.onHost = undefined; evaluateNodeSpy.mockClear(); textCharsSpy.mockClear(); hostTaskSpy.mockClear() })

const frozen = () => 0
/** The per-call bookkeeping reservation a direct viewFor call needs (the evaluator makes it before building a view). */
type FakeSourceLike = ReturnType<typeof fakeSource>
const workspace = () => createCfMemory().reserve(CF_FORMULA_CALL_BYTES + CF_VIEW_WORKSPACE_BYTES) as CfReservation
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
    expect(calls[0]).toEqual({ nodes: 5, viewBuild: 4, index: 0, enumeration: 8, reads: 8, sort: 16, comparisons: 0, matcher: 0, text: 0, criterion: 0, lookups: 0, total: 41 })
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
    expect(calls[0]).toEqual({ nodes: 3, viewBuild: 5, index: 0, enumeration: 5, reads: 5, sort: 15, comparisons: 4, matcher: 0, text: 0, criterion: 0, lookups: 0, total: 37 })
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
      const budget: CfWorkBudget = { get units() { return base.units }, get remaining() { return base.remaining }, spend: n => { const r = base.spend(n); if (r === 'yield') yields++; return r }, resume: () => base.resume() }
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
    const view = snapshot.viewFor(rects.map(r => ({ rect: rectOf(r) })), bigBudget(), workspace())
    if (view.kind !== 'ready') throw new Error('view refused')
    return view.value.metrics
  }
  it('LEN(A1&A2) over two 32,767-character strings is refused from its payload whatever memory is available: no engine or textChars call', async () => {
    const big = 'x'.repeat(32_767)
    const snapshot = await snapshotOf({ A1: big, A2: big }, ['A1', 'A2'])
    const f = compile('LEN(A1&A2)', 'B1')
    const b = cfCallWorkBreakdown(cfCallShape(f), metricsOf(snapshot, ['A1', 'A2']), []) as CfCallWorkBreakdown
    // & renders both (65,534) and writes the result (65,534); LEN renders and splits it (2 x 65,534): 262,136 characters.
    expect(b.text).toBe(262_136 / CF_TEXT_UNIT_CHARS)
    expect(b.total - b.text).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits) // a mutant without the text term admits it
    evaluateNodeSpy.mockClear(); textCharsSpy.mockClear()
    expect(evaluateCfFormula(f, a1('B1'), snapshot, context(createCfMemory({ limitBytes: Number.MAX_SAFE_INTEGER })))).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    expect(textCharsSpy).not.toHaveBeenCalled()
    snapshot.dispose()
  })
  it('every operator and every string-consuming call (AND/OR/NOT/IF, math, date, scalar aggregates over a scalar) charges the string it consumes: over 32,767 characters, and just over the cap, each withholds (Task 34 fix 3)', async () => {
    for (const text of ['A1+1>0', 'AND(A1)', 'OR(A1,TRUE)', 'NOT(A1)', '-A1<0', 'ABS(A1)>0', 'MAX(IFERROR(A1,0))', 'SUM(IFNA(A1,0))', 'IF(A1,1,0)', 'YEAR(A1)>0', 'MOD(A1,2)']) {
      const f = compile(text, 'B1')
      for (const length of [32_767]) {
        const snapshot = await snapshotOf({ A1: 'x'.repeat(length) }, ['A1'])
        const b = cfCallWorkBreakdown(cfCallShape(f), metricsOf(snapshot, ['A1']), []) as CfCallWorkBreakdown
        // At least one pass over the string at the text rate (round 3 per-operation model).
        expect(b.text, `${text} over ${length}`).toBeGreaterThanOrEqual(Math.ceil(length / CF_TEXT_UNIT_CHARS))
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
  it('LEN(A1&A2) over two 10-character strings is 4 + 2 + 2 + (40 + 40) / 4 = 28, and runs at its exact scratch bound but never one byte below', async () => {
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
    expect(calls[0]).toEqual({ nodes: 4, viewBuild: 2, index: 0, enumeration: 0, reads: 2, sort: 0, comparisons: 0, matcher: 0, text: 20, criterion: 0, lookups: 0, total: 28 })
    expect(exact.peakBytes).toBe(scratch + tree)
    expect(exact.currentBytes).toBe(0)
    snapshot.dispose()
  })
  it('COUNTIF(A1:A2,B1): 2 x 11 x 11 = 242 matcher transitions (+ 11 for the criterion\'s validation, round 4) over 10-character strings refuses; over 3 characters it is 50 plus the 3-character criterion parse and its 4-row validation, 57', async () => {
    const f = compile('COUNTIF(A1:A2,B1)', 'C1')
    const ten = await snapshotOf({ A1: 'abcdefghij', A2: 'abcdefghik', B1: 'abcdefghij' }, ['A1:A2', 'B1'])
    const b10 = cfCallWorkBreakdown(cfCallShape(f), metricsOf(ten, ['A1:A2', 'B1']), [10]) as CfCallWorkBreakdown
    expect(b10.matcher).toBe(242 + 11)
    expect(b10.total - b10.matcher).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits) // a mutant without the matcher term admits it
    evaluateNodeSpy.mockClear()
    expect(evaluateCfFormula(f, a1('C1'), ten, context())).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    const three = await snapshotOf({ A1: 'abc', A2: 'abd', B1: 'abc' }, ['A1:A2', 'B1'])
    const calls: CfCallWorkBreakdown[] = []
    expect(evaluateCfFormula(f, a1('C1'), three, context(createCfMemory(), calls))).toEqual({ kind: 'ready', value: { kind: 'number', value: 1 } })
    expect(calls[0]).toEqual({ nodes: 3, viewBuild: 3, index: 0, enumeration: 3, reads: 3, sort: 6, comparisons: 0, matcher: 32 + 4, text: 0, criterion: 3, lookups: 0, total: 57 })
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
    const view = snapshot.viewFor([{ rect: rectOf('A1:A4') }], bigBudget(), workspace())
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
    // Only the per-call bookkeeping (the view's workspace) was ever held: the scratch growth was refused before use.
    expect(short.peakBytes).toBe(CF_FORMULA_CALL_BYTES + CF_VIEW_WORKSPACE_BYTES)
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
      const dynamic = b.viewBuild + b.index + b.enumeration + b.sort + b.reads + b.lookups + b.criterion
      if (exact.includes(text)) expect(spent, text).toBe(dynamic)
      else expect(spent, text).toBeLessThanOrEqual(dynamic)
      expect(spent, text).toBeGreaterThan(0)
    }
  })
})
void refOf

describe('a refused call never poisons the shared production budget (Task 34 fix round 2, finding 6)', () => {
  it('with 6 units left, the 7-unit A1+1>0 refuses, paying only its discovery; the 5-unit A1>0 that follows still evaluates', async () => {
    const signal = new AbortController().signal
    const budget = createCfWorkBudget({ now: () => 0, signal }) // the production per-call budget (activeUnits)
    const calls: number[] = []
    const batch = await batchOf([formulaRule('A1+1>0', 'B1'), formulaRule('A1>0', 'B1')], fakeSource({ A1: 3 }), { budget, signal, onCall: b => calls.push(b.total) })
    expect(budget.spend(budget.remaining - 6)).not.toBe('refused')
    expect(budget.remaining).toBe(6)
    expect(batch.evaluate(0, 0, 1)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    // The refusal latched nothing; it paid only its 1-unit view build (round 3, finding 1: discovery is charged).
    expect(budget.remaining).toBe(5)
    expect(batch.evaluate(1, 0, 1)).toEqual({ kind: 'ready', value: { kind: 'boolean', value: true } })
    expect(calls).toEqual([5])
    expect(budget.remaining).toBe(0)
    batch.dispose()
  })
})

describe('the view index walk is bounded, charged and reserved before it builds anything (Task 34 fix round 2, finding 1)', () => {
  const ref = (row: number, col: number) => refOf(row, col)
  /** A rule `SUM(<rect>)>0` anchored at `at`, beside a second rule whose footprint puts `others` in the shared snapshot. */
  async function walk(rect: string, others: Cells, coverRects: readonly string[], at: string, options: { memory?: ReturnType<typeof createCfMemory> } = {}) {
    const calls: CfCallWorkBreakdown[] = []
    const settled: Array<{ spent: number; total: number }> = []
    const rule = formulaRule(`SUM(${rect})>0`, at)
    const cover = coverRects.map(r => formulaRule(`COUNTA(${r})>0`, at))
    const batch = await batchOf([rule, ...cover], fakeSource(others), {
      hoist: false, onCall: b => calls.push(b), onSettled: (spent, b) => settled.push({ spent, total: b.total }), ...(options.memory ? { memory: options.memory } : {}),
    })
    const t = a1(at)
    evaluateNodeSpy.mockClear()
    const out = batch.evaluate(0, t.row, t.col)
    batch.dispose()
    return { out, calls, settled }
  }
  it('a one-row call over a snapshot whose other row fills every column walks the row, not 16,384 columns', async () => {
    // Row 2 holds A2:XFD2 (16,384 columns); row 1 holds only XFD1. SUM(A1:XFD1) reads one cell.
    const others: Cells = Object.fromEntries(Array.from({ length: 16384 }, (_, c) => [ref(1, c), 1]))
    const cells: Cells = { ...others, XFD1: 5 }
    const calls: CfCallWorkBreakdown[] = []
    const rule = formulaRule('SUM(A1:XFD1)>0', 'A3'), cover = formulaRule('SUM(A2:XFD2)>0', 'A3')
    const batch = await batchOf([rule, cover], fakeSource(cells), { hoist: false, onCall: b => calls.push(b) })
    expect(batch.evaluate(0, 2, 0)).toEqual({ kind: 'ready', value: { kind: 'boolean', value: true } })
    expect(calls[0]!.index).toBeLessThanOrEqual(2)
    expect(calls[0]!.total).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits)
    batch.dispose()
  }, 60_000)
  it('empty steps are charged as index work: 20 rows each holding a key past the rect cost 20, in the bound and metered', async () => {
    // A1:T20 is empty; U1:U20 (a key per row, outside the rect) and A21:T21 (a key per column, outside the rect).
    const others: Cells = { ...Object.fromEntries(Array.from({ length: 20 }, (_, r) => [ref(r, 20), 1])), ...Object.fromEntries(Array.from({ length: 20 }, (_, c) => [ref(20, c), 1])) }
    const { out, calls, settled } = await walk('A1:T20', others, ['U1:U20', 'A21:T21'], 'Z1')
    expect(out).toEqual({ kind: 'ready', value: { kind: 'boolean', value: false } })
    expect(calls[0]!.index).toBe(20)
    expect(calls[0]!.total).toBe(calls[0]!.nodes + calls[0]!.viewBuild + calls[0]!.enumeration + calls[0]!.reads + calls[0]!.sort + calls[0]!.index + calls[0]!.text)
    expect(settled[0]!.spent).toBeLessThanOrEqual(settled[0]!.total)
  })
  it('an empty rect whose index walk passes 128 empty steps is refused, though it reads no cell', async () => {
    const others: Cells = { ...Object.fromEntries(Array.from({ length: 200 }, (_, r) => [ref(r, 200), 1])), ...Object.fromEntries(Array.from({ length: 200 }, (_, c) => [ref(200, c), 1])) }
    const { out } = await walk(`A1:${ref(199, 199)}`, others, [`${ref(0, 200)}:${ref(199, 200)}`, `A201:${ref(200, 199)}`], 'HA1')
    expect(out).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
  })
  it('viewFor itself stops the walk at the cap (each empty step spent as taken) and refuses without the workspace reservation', async () => {
    const signal = new AbortController().signal, memory = createCfMemory()
    const atoms = ['A1:GR200', 'GS1:GS200', 'A201:GR201'].map(r => ({ rect: rectOf(r) }))
    const cells: Cells = { ...Object.fromEntries(Array.from({ length: 200 }, (_, r) => [ref(r, 200), 1])), ...Object.fromEntries(Array.from({ length: 200 }, (_, c) => [ref(200, c), 1])) }
    const snap = await prepareCfDependencies(atoms, fakeSource(cells), undefined, bigBudget(signal), memory, signal, { sheetId: 's1' })
    if (snap.kind !== 'ready') throw new Error('snapshot')
    const snapshot = snap.value.value
    let limit: number = CF_LIMITS.checkpointUnits, used = 0
    const spend = vi.fn((n: number) => { if (n > limit - used) return 'refused' as const; used += n; return 'continue' as const })
    const meter: CfWorkBudget = { spend, resume() {}, get units() { return used }, get remaining() { return limit - used } }
    expect(snapshot.viewFor([{ rect: rectOf('A1:GR200') }], meter, workspace())).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(spend).toHaveBeenCalledTimes(CF_LIMITS.checkpointUnits + 1) // 128 steps taken, the 129th refused, none after
    spend.mockClear(); used = 0; limit = 1e9
    const none = createCfMemory().reserve(0) as CfReservation
    expect(snapshot.viewFor([{ rect: rectOf('A1:GR200') }], meter, none)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(spend).not.toHaveBeenCalled()
    snapshot.dispose()
  })
  it('an out-of-order source rebuilds the row-major index (charged): the row walk still finds every cell once', async () => {
    const cells: Cells = { ...Object.fromEntries(Array.from({ length: 26 }, (_, c) => [ref(1, c), 1])), Z1: 5, Y1: 'ab' }
    const base = fakeSource(cells)
    const reversed: FakeSourceLike = { ...base, async *pages(id, rects, budget, signal) { const all = []; for await (const p of base.pages(id, rects, budget, signal)) all.push(p); yield* all.reverse() } }
    const signal = new AbortController().signal, memory = createCfMemory()
    const inOrder = bigBudget(signal), outOfOrder = bigBudget(signal)
    const a = await prepareCfDependencies([{ rect: rectOf('A1:Z2') }], base, undefined, inOrder, memory, signal, { sheetId: 's1' })
    const b = await prepareCfDependencies([{ rect: rectOf('A1:Z2') }], reversed, undefined, outOfOrder, memory, signal, { sheetId: 's1' })
    if (a.kind !== 'ready' || b.kind !== 'ready') throw new Error('snapshot')
    expect(outOfOrder.units).toBeGreaterThan(inOrder.units) // the rebuild and its sort were charged
    for (const snap of [a.value.value, b.value.value]) {
      const view = snap.viewFor([{ rect: rectOf('A1:Z1') }], bigBudget(), workspace())
      if (view.kind !== 'ready') throw new Error('view')
      expect(view.value.metrics).toMatchObject({ viewCells: 2, index: 0, atoms: [{ cells: 2, stringUnits: 2, longestString: 2 }] })
      snap.dispose()
    }
    expect(memory.currentBytes).toBe(0)
  })
  it('the per-call workspace is reserved before the view is built: a grant below it never calls viewFor', async () => {
    const signal = new AbortController().signal, memory = createCfMemory()
    const snap = await prepareCfDependencies([{ rect: rectOf('A1:A4') }], fakeSource({ A1: 1, A2: 2, A3: 3, A4: 4 }), undefined, bigBudget(signal), memory, signal, { sheetId: 's1' })
    if (snap.kind !== 'ready') throw new Error('snapshot')
    const snapshot = snap.value.value
    const viewSpy = vi.spyOn(snapshot, 'viewFor')
    const budget = bigBudget()
    const out = evaluateCfFormula(compile('SUM(A1:A4)>2', 'B1'), a1('B1'), snapshot, {
      sheetName: 'Sheet1', evaluation: EVALUATION, semantics: cfFormulaSemantics(EVALUATION), edge: 'ref-error', materialization: { limit: 1e6, used: 0 },
      budget, memory: createCfMemory({ limitBytes: 0 }), invariants: createCfInvariantMemo(createCfMemory()), ceiling: createCfUnitCeiling(),
    })
    expect(out).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(viewSpy).not.toHaveBeenCalled()
    expect(budget.units).toBe(0)
    snapshot.dispose()
  })
})

describe('text work is bounded per operation, generated strings and criteria included (Task 34 fix round 2, finding 2)', () => {
  async function callOf(text: string, cells: Cells, at = 'Z1') {
    const calls: CfCallWorkBreakdown[] = []
    const batch = await batchOf([formulaRule(text, at)], fakeSource(cells), { onCall: b => calls.push(b), hoist: true })
    const t = a1(at)
    evaluateNodeSpy.mockClear()
    const out = batch.evaluate(0, t.row, t.col)
    const ran = evaluateNodeSpy.mock.calls.length > 0
    batch.dispose()
    return { out, calls, ran }
  }
  it('30 nested UPPER calls over 50 characters are charged per call and refused before invocation', async () => {
    const text = `LEN(${'UPPER('.repeat(30)}A1${')'.repeat(30)})>0`
    const { out, ran } = await callOf(text, { A1: 'x'.repeat(50) })
    expect(out).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(ran).toBe(false)
    // The work bound itself (not only the scratch cap) walks the string once per UPPER and once for LEN.
    const b = cfCallWorkBreakdown(cfCallShape(compile(text, 'Z1')), { viewCells: 1, atoms: [{ cells: 1, stringUnits: 50, longestString: 50 }] }, []) as CfCallWorkBreakdown
    expect(b.text).toBeGreaterThanOrEqual(50 * 31)
    // One UPPER over 3 characters still evaluates.
    expect((await callOf('LEN(UPPER(A1))>0', { A1: 'abc' })).out).toEqual({ kind: 'ready', value: { kind: 'boolean', value: true } })
  })
  it('a 255-character numeric-text criterion is charged for its parse: COUNTIF(A1:A2,B1) is refused', async () => {
    const { out, ran } = await callOf('COUNTIF(A1:A2,B1)', { A1: 1, A2: 2, B1: '0'.repeat(255) })
    expect(out).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(ran).toBe(false)
    expect((await callOf('COUNTIF(A1:A2,B1)', { A1: 1, A2: 2, B1: '1' })).out).toEqual({ kind: 'ready', value: { kind: 'number', value: 1 } })
  })
  it('a frequency hoist\'s criterion text is charged for its key: a 200-character criterion is refused', async () => {
    const cells: Cells = { A1: 'a', A2: 'x'.repeat(200) }
    expect((await callOf('COUNTIF($A$1:$A$2,A2)', cells, 'B2')).out).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect((await callOf('COUNTIF($A$1:$A$2,A1)', cells, 'B1')).out).toEqual({ kind: 'ready', value: { kind: 'number', value: 1 } })
  })
  it('numbers rendered as text are charged: LEN(CONCAT(1e100 x 60)) is refused, and the render bound covers the engine\'s widest', async () => {
    const { out, ran } = await callOf(`LEN(CONCAT(${Array.from({ length: 60 }, () => '1e100').join(',')}))>0`, {})
    expect(out).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(ran).toBe(false)
    for (const v of [-0.00000123456789012345, -1.23456789012345e-300, -1.79769313486231e308, -123456789012345e6, 2 ** 53 + 1]) {
      expect(evaluatorModule.coerceToString(v).length).toBeLessThanOrEqual(CF_TEXT_RENDER_UNITS)
    }
    for (const v of [true, false, { kind: 'formula-error', code: '#GETTING_DATA' }] as const) expect(evaluatorModule.coerceToString(v as never).length).toBeLessThanOrEqual(CF_TEXT_RENDER_UNITS)
  })
  it('sibling case mappings do not compound: UPPER(A2)=UPPER(B2) over two 5-character strings evaluates', async () => {
    const f = compile('UPPER(A2)=UPPER(B2)', 'C2')
    const five = { cells: 1, stringUnits: 5, longestString: 5 } // unmeasured ASCII: taken as non-ASCII
    const asciiFive = { ...five, strings: 1, nonAscii: 0 }
    // ASCII: each UPPER reads 5 and writes 5; the comparison folds and compares each 5-unit side (2 x 5 each): 40 / 4.
    expect((cfCallWorkBreakdown(cfCallShape(f), { viewCells: 2, atoms: [asciiFive, asciiFive] }, []) as CfCallWorkBreakdown).text).toBe(10)
    // Non-ASCII: each UPPER may write 15, and each folded side may grow again: (5 + 15) x 2 + 15 x 4 x 2 = 160 / 4.
    const b = cfCallWorkBreakdown(cfCallShape(f), { viewCells: 2, atoms: [five, five] }, []) as CfCallWorkBreakdown
    expect(b.text).toBe(40)
    expect((await callOf('UPPER(A2)=UPPER(B2)', { A2: 'abcde', B2: 'ABCDE' }, 'C2')).out).toEqual({ kind: 'ready', value: { kind: 'boolean', value: true } })
    // Nested mappings do compound: UPPER(LOWER(A2)) is bounded at 9x.
    expect(cfTextUnits(cfCallShape(compile('LEN(UPPER(LOWER(A2)))>0', 'C2')), { viewCells: 1, atoms: [five] }).intermediate).toBe(45)
  })
  it('case mapping can lengthen a string: UPPER over "ß" x 40 is bounded by its output (80), not its input', async () => {
    const f = compile('LEN(UPPER(A1))=80', 'B1')
    const shape = cfCallShape(f)
    const b = cfCallWorkBreakdown(shape, { viewCells: 1, atoms: [{ cells: 1, stringUnits: 40, longestString: 40 }] }, []) as CfCallWorkBreakdown
    expect('ß'.repeat(40).toUpperCase().length).toBe(80)
    expect(cfTextUnits(shape, { viewCells: 1, atoms: [{ cells: 1, stringUnits: 40, longestString: 40 }] }).chars).toBeGreaterThanOrEqual(40 + 80)
    expect(b.text).toBeGreaterThan(0)
    // The longest intermediate (what scratch is sized from, and what each later walk may meet) is the 80-unit output.
    expect(cfTextUnits(shape, { viewCells: 1, atoms: [{ cells: 1, stringUnits: 40, longestString: 40 }] }).intermediate).toBeGreaterThanOrEqual(80)
  })
})

describe('a refused call\'s discovery work is charged to the preparation (Task 34 fix round 3, finding 1)', () => {
  it('an index walk refused at the cap charges its steps to the production budget and ceiling, and asks for a yield', async () => {
    const ref = (row: number, col: number) => refOf(row, col)
    const others: Cells = { ...Object.fromEntries(Array.from({ length: 200 }, (_, r) => [ref(r, 200), 1])), ...Object.fromEntries(Array.from({ length: 200 }, (_, c) => [ref(200, c), 1])) }
    const signal = new AbortController().signal
    const budget = createCfWorkBudget({ now: frozen, signal })
    const ceiling = createCfUnitCeiling()
    const rule = formulaRule(`SUM(A1:${ref(199, 199)})>0`, 'HA1')
    const cover = [formulaRule(`COUNTA(${ref(0, 200)}:${ref(199, 200)})>0`, 'HA1'), formulaRule(`COUNTA(A201:${ref(200, 199)})>0`, 'HA1')]
    const batch = await batchOf([rule, ...cover], fakeSource(others), { hoist: false, budget, signal, ceiling })
    const before = ceiling.used
    for (let i = 0; i < 3; i++) expect(batch.evaluate(0, 0, a1('HA1').col)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(budget.units).toBe(3 * CF_LIMITS.checkpointUnits)
    expect(ceiling.used - before).toBe(3 * CF_LIMITS.checkpointUnits)
    // The pacing yield the charged work earns is honoured between cells.
    hostTaskSpy.mockClear()
    await batch.run([{ row: 0, col: a1('HA1').col }], () => {})
    expect(hostTaskSpy).toHaveBeenCalled()
    batch.dispose()
  })
  it('cell hits walked before a refusal are charged: an un-hoisted SUM(A1:A200) refused at hit 129, three times', async () => {
    const signal = new AbortController().signal
    const budget = createCfWorkBudget({ now: frozen, signal })
    const ceiling = createCfUnitCeiling()
    const batch = await batchOf([formulaRule('SUM(A1:A200)>0', 'B1')], fakeSource(Object.fromEntries(Array.from({ length: 200 }, (_, r) => [`A${r + 1}`, r]))), { budget, signal, ceiling, hoist: false })
    const before = ceiling.used
    for (let i = 0; i < 3; i++) expect(batch.evaluate(0, 0, 1)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(budget.units).toBe(3 * CF_LIMITS.checkpointUnits)
    expect(ceiling.used - before).toBe(3 * CF_LIMITS.checkpointUnits)
    batch.dispose()
  })
  it('discovery is bounded by the budget\'s remaining units before it runs', async () => {
    const signal = new AbortController().signal
    const budget = createCfWorkBudget({ now: frozen, signal })
    expect(budget.spend(budget.remaining - 3)).not.toBe('refused')
    const batch = await batchOf([formulaRule('SUM(A1:A10)>0', 'B1')], fakeSource(Object.fromEntries(Array.from({ length: 10 }, (_, r) => [`A${r + 1}`, r]))), { budget, signal, hoist: false })
    expect(batch.evaluate(0, 0, 1)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(budget.remaining).toBe(0) // the view build spent its first 3 keys, then refused: those 3 are charged
    expect(budget.spend(0)).not.toBe('refused') // never latched
    batch.dispose()
  })
})

describe('criterion preprocessing is charged before it runs (Task 34 fix round 3, finding 2)', () => {
  const twenty = Array.from({ length: 20 }, () => 'COUNTIF(A1:A2,B1)').join('+') + '>0'
  it('twenty COUNTIF terms over a 255-character criterion never tokenise it: the first preflight refuses', async () => {
    const tokenize = vi.mocked(wildcardModule.tokenizeWildcardPattern)
    for (const B1 of ['a'.repeat(255), '0'.repeat(255)]) {
      const signal = new AbortController().signal
      const budget = createCfWorkBudget({ now: frozen, signal })
      const batch = await batchOf([formulaRule(twenty, 'C1')], fakeSource({ A1: 1, A2: 2, B1 }), { budget, signal })
      tokenize.mockClear(); evaluateNodeSpy.mockClear()
      expect(batch.evaluate(0, 0, 2)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
      expect(tokenize).not.toHaveBeenCalled()
      expect(evaluateNodeSpy).not.toHaveBeenCalled()
      // What the refused call did is charged, and it is bounded by the per-call cap.
      expect(budget.units).toBeGreaterThan(0)
      expect(budget.units).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits)
      batch.dispose()
    }
  })
  it('a short criterion is still tokenised and evaluates', async () => {
    const tokenize = vi.mocked(wildcardModule.tokenizeWildcardPattern)
    tokenize.mockClear()
    const batch = await batchOf([formulaRule('COUNTIF(A1:A2,B1)>0', 'C1')], fakeSource({ A1: 'ab', A2: 'x', B1: 'ab' }))
    expect(batch.evaluate(0, 0, 2)).toEqual({ kind: 'ready', value: { kind: 'boolean', value: true } })
    expect(tokenize).toHaveBeenCalled()
    batch.dispose()
  })
})

describe('the per-node text model: own passes per operation, ASCII from the snapshot (Task 34 fix round 3, finding 4)', () => {
  /** One call through the evaluator: its outcome, its admitted breakdown, and its bound from the measured view. */
  async function call(text: string, cells: Cells, at: string) {
    const calls: CfCallWorkBreakdown[] = []
    const batch = await batchOf([formulaRule(text, at)], fakeSource(cells), { onCall: b => calls.push(b) })
    const t = a1(at)
    evaluateNodeSpy.mockClear()
    const out = batch.evaluate(0, t.row, t.col)
    const ran = evaluateNodeSpy.mock.calls.length > 0
    batch.dispose()
    return { out, calls, ran }
  }
  /** The static bound over the real snapshot view (ASCII flags measured), whether or not the call is admitted. */
  async function boundOf(text: string, cells: Cells, at: string, rects: string[]) {
    const signal = new AbortController().signal
    const snap = await prepareCfDependencies(rects.map(r => ({ rect: rectOf(r) })), fakeSource(cells), undefined, bigBudget(signal), createCfMemory(), signal, { sheetId: 's1' })
    if (snap.kind !== 'ready') throw new Error('snapshot')
    const view = snap.value.value.viewFor(rects.map(r => ({ rect: rectOf(r) })), bigBudget(), workspace())
    if (view.kind !== 'ready') throw new Error('view')
    const shape = cfCallShape(compile(text, at))
    const b = cfCallWorkBreakdown(shape, view.value.metrics, shape.conditionals.flatMap(c => c.pairs.map(() => 0))) as CfCallWorkBreakdown
    snap.value.value.dispose()
    return b
  }
  const asText = (v: unknown) => (typeof v === 'boolean' ? { kind: 'boolean', value: v } : typeof v === 'string' ? { kind: 'text', value: v } : typeof v === 'number' ? { kind: 'number', value: v } : { kind: 'error', value: v })
  const ascii = (n: number) => Array.from({ length: n }, (_, i) => String.fromCharCode(97 + (i % 26))).join('')
  it('the brief\'s ordinary formulas evaluate and paint as the engine', async () => {
    const cases: Array<[string, Cells, string]> = [
      ['UPPER(A2)="ABC"', { A2: 'abc' + ascii(15) }, 'C2'],
      ['UPPER(A2)="ABC"', { A2: ascii(60) }, 'C2'],
      ['UPPER(A2)="ABC"', { A2: 'abc' }, 'C2'],
      ['LOWER(A2)=B2', { A2: ascii(11).toUpperCase(), B2: ascii(11) }, 'C2'],
      ['UPPER(A2)=""', { A2: '' }, 'C2'],
      ['B1&UPPER($A$2)', { B1: -0.00000123456789012345, A2: 'ab' }, 'C1'],
    ]
    for (const [text, cells, at] of cases) {
      const { out, calls } = await call(text, cells, at)
      expect(out, `${text} ${JSON.stringify(cells)}`).toEqual({ kind: 'ready', value: asText(engineEval(text, cells, a1(at)).value) })
      expect(calls[0]!.total).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits)
      // An empty-string cell is a string of length 0, never a rendered non-string: no text work at all.
      if (text === 'UPPER(A2)=""') expect(calls[0]!.text).toBe(0)
    }
  })
  it('adversarial shapes still refuse, their bounds over the cap', async () => {
    const x = 'x'.repeat(32_767)
    const cases: Array<[string, Cells, string, string[]]> = [
      [`LEN(${'UPPER('.repeat(30)}A1${')'.repeat(30)})>0`, { A1: 'ß'.repeat(50) }, 'Z1', ['A1']],
      ['UPPER(A2)="ABC"', { A2: 'é'.repeat(60) }, 'C2', ['A2']], // non-ASCII keeps the x3 expansion
      ['A1+1>0', { A1: x }, 'B1', ['A1']],
      ['AND(A1)', { A1: x }, 'B1', ['A1']],
      ['COUNTIF(A1:A2,B1)', { A1: 1, A2: 2, B1: '0'.repeat(255) }, 'C1', ['A1:A2', 'B1']],
      [`LEN(CONCAT(${Array.from({ length: 60 }, () => '1e100').join(',')}))>0`, {}, 'Z1', []],
    ]
    for (const [text, cells, at, rects] of cases) {
      const { out, ran } = await call(text, cells, at)
      expect(out, text).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
      expect(ran, text).toBe(false)
      if (rects.length) expect((await boundOf(text, cells, at, rects)).total, text).toBeGreaterThan(CF_LIMITS.checkpointUnits)
    }
  })
  it('the 60-character case needs a rate: at one unit per character its own passes alone exceed the cap', async () => {
    const b = await boundOf('UPPER(A2)="ABC"', { A2: ascii(60) }, 'C2', ['A2'])
    const chars = cfTextUnits(cfCallShape(compile('UPPER(A2)="ABC"', 'C2')), { viewCells: 1, atoms: [{ cells: 1, stringUnits: 60, longestString: 60, strings: 1, nonAscii: 0 }] }).chars
    expect(chars).toBe(60 + 60 + 2 * 60 + 2 * 3) // UPPER reads and writes 60; the comparison folds and compares each side
    expect(b.total - b.text + chars).toBeGreaterThan(CF_LIMITS.checkpointUnits)
    expect(b.total).toBeLessThanOrEqual(CF_LIMITS.checkpointUnits)
  })
  it('soundness: the model\'s length bound covers the engine\'s rendered result on random text formulas', () => {
    const r = (() => { let a = 341; return () => { a = (a + 0x6d2b79f5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296 } })()
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(r() * xs.length)]!
    const values: CellInput[] = ['', 'ab', 'Hello World', '  a  b  ', 'ß'.repeat(4), 'ΐ'.repeat(3), 'İi', 'x'.repeat(30), 7, -0.00000123456789012345, 1e21, true, false, null]
    const leaves = ['A1', 'A2', 'A3', 'A4', 'A5', '"ab"', '"ẞ"', '1.5', 'TRUE', '-12345.678']
    const ops: Array<(a: string, b: string) => string> = [
      a => `UPPER(${a})`, a => `LOWER(${a})`, a => `TRIM(${a})`, a => `LEFT(${a},2)`, a => `RIGHT(${a},3)`, a => `MID(${a},2,3)`,
      (a, b) => `${a}&${b}`, (a, b) => `CONCAT(${a},${b})`, (a, b) => `IF(TRUE,${a},${b})`, (a, b) => `IFERROR(${a},${b})`, a => `LEN(${a})`,
      (_a, b) => `IF(TRUE,,${b})`, (a, _b) => `IF(FALSE,${a},)`, // omitted arguments
      a => `IF(TRUE,${a},"")`, a => `IF(FALSE,"",${a})`, // a selected blank reference reads as 0 (round 4, finding 2)
    ]
    const gen = (depth: number): string => (depth === 0 || r() < 0.25 ? pick(leaves) : pick(ops)(gen(depth - 1), gen(depth - 1)))
    let checked = 0
    for (let trial = 0; trial < 400; trial++) {
      const cells: Cells = Object.fromEntries(['A1', 'A2', 'A3', 'A4'].map(k => [k, pick(values)]))
      const text = gen(4)
      const f = compile(text, 'B1')
      const shape = cfCallShape(f)
      const atoms = shape.refs.map(node => {
        const v = cells[refOf((node as { ref: { row: number } }).ref.row, (node as { ref: { col: number } }).ref.col)]
        if (v === null || v === undefined) return { cells: 0, stringUnits: 0, longestString: 0, strings: 0, nonAscii: 0 }
        if (typeof v !== 'string') return { cells: 1, stringUnits: 0, longestString: 0, strings: 0, nonAscii: 0 }
        return { cells: 1, stringUnits: v.length, longestString: v.length, strings: 1, nonAscii: /[^\x00-\x7f]/.test(v) ? 1 : 0 }
      })
      const trace = new Map<AstNode, TextValue>()
      cfTextUnits(shape, { viewCells: atoms.length, atoms }, trace)
      const engine = engineEval(text, cells, a1('B1'))
      const rendered = evaluatorModule.coerceToString(engine.value as never).length
      expect(rendered, `${text} over ${JSON.stringify(cells)}`).toBeLessThanOrEqual(trace.get(f.ast)!.len)
      checked++
    }
    expect(checked).toBe(400)
  })
})

describe('Task 34 fix round 4 (codex round-3 findings 1, 2, 4, 5)', () => {
  it('finding 4: a text criterion\'s validation (matchWildcard(body, \'\')) is in the bound whatever the range holds', async () => {
    // Codex counterexample: COUNTIF(A1:A2,B1) over an empty A range, B1 = 140 ASCII 'a'.
    const cells: Cells = { B1: 'a'.repeat(140) }
    const calls: CfCallWorkBreakdown[] = []
    const batch = await batchOf([formulaRule('COUNTIF(A1:A2,B1)', 'C1')], fakeSource(cells), { onCall: b => calls.push(b) })
    evaluateNodeSpy.mockClear()
    expect(batch.evaluate(0, 0, 2)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    expect(calls).toEqual([])
    batch.dispose()
    // The validation walks one DP row per token plus one over the empty text: 141 cells, independent of the range.
    const shape = cfCallShape(compile('COUNTIF(A1:A2,B1)', 'C1'))
    const payload = { viewCells: 1, atoms: [{ cells: 0, stringUnits: 0, longestString: 0 }, { cells: 1, stringUnits: 140, longestString: 140, strings: 1, nonAscii: 0 }] }
    const b = cfCallWorkBreakdown(shape, payload, [140]) as CfCallWorkBreakdown
    expect(b.matcher).toBe(141)
    expect(b.total).toBeGreaterThan(CF_LIMITS.checkpointUnits)
    // Its DP rows are in the scratch bound.
    const f = compile('COUNTIF(A1:A2,B1)', 'C1')
    expect(cfEvaluatorScratchBytes(f, payload, [], [140]) - cfEvaluatorScratchBytes(f, payload, [], [0])).toBeGreaterThanOrEqual(141 * CF_MATCHER_CELL_BYTES)
    const none = cfCallWorkBreakdown(shape, payload, [0]) as CfCallWorkBreakdown
    expect(none.matcher).toBe(0) // a numeric criterion never reaches the matcher
    // A short text criterion over an empty range still evaluates: 3 nodes + 1 view + 1 read + criterion + 4 matcher.
    const short = await batchOf([formulaRule('COUNTIF(A1:A2,B1)', 'C1')], fakeSource({ B1: 'abc' }), { onCall: x => calls.push(x) })
    expect(short.evaluate(0, 0, 2)).toEqual({ kind: 'ready', value: { kind: 'number', value: 0 } })
    expect(calls[0]!.matcher).toBe(4)
    short.dispose()
  })
  it('finding 2: IF converts a selected blank to 0, so the length it propagates covers "0"', async () => {
    // Codex counterexample: 120 nested UPPER over IF(TRUE,A1,"") with A1 blank renders "0" through every UPPER.
    const text = `${'UPPER('.repeat(120)}IF(TRUE,A1,"")${')'.repeat(120)}`
    const calls: CfCallWorkBreakdown[] = []
    const batch = await batchOf([formulaRule(text, 'B1')], fakeSource({}), { onCall: b => calls.push(b) })
    evaluateNodeSpy.mockClear()
    expect(batch.evaluate(0, 0, 1)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    batch.dispose()
    expect(engineEval(text, {}, a1('B1')).value).toBe('0')
    // The model: the IF node renders at least "0", and each UPPER reads and writes it (2 characters each).
    const f = compile(text, 'B1')
    const trace = new Map<AstNode, TextValue>()
    const units = cfTextUnits(cfCallShape(f), { viewCells: 0, atoms: [{ cells: 0, stringUnits: 0, longestString: 0 }] }, trace)
    expect(trace.get(f.ast)!.len).toBeGreaterThanOrEqual(1)
    expect(units.chars).toBeGreaterThanOrEqual(120 * 2)
    // IF(FALSE,"",A1) and an omitted branch convert too; an IFERROR passes a blank through unconverted (renders "").
    for (const [t, want] of [['IF(FALSE,"",A1)', '0'], ['IFERROR(A1,"x")', '']] as const) {
      const g = compile(t, 'B1')
      const tr = new Map<AstNode, TextValue>()
      cfTextUnits(cfCallShape(g), { viewCells: 0, atoms: [{ cells: 0, stringUnits: 0, longestString: 0 }] }, tr)
      expect(evaluatorModule.coerceToString(engineEval(t, {}, a1('B1')).value as never)).toBe(want)
      expect(tr.get(g.ast)!.len, t).toBeGreaterThanOrEqual(want.length)
    }
  })
  it('finding 5: a refused frequency key pays its key-processing allowance to the budget and the ceiling; an admitted one pays it once', async () => {
    const run = async (B1: string) => {
      const signal = new AbortController().signal
      const budget = createCfWorkBudget({ now: frozen, signal })
      const ceiling = createCfUnitCeiling()
      const calls: CfCallWorkBreakdown[] = []
      const batch = await batchOf([formulaRule('COUNTIF($A$1:$A$2,B1)>0', 'C1')], fakeSource({ A1: 1, A2: 2, B1 }), { budget, signal, ceiling, hoist: true, onCall: b => calls.push(b) })
      const units = budget.units, used = ceiling.used
      const out = batch.evaluate(0, 0, 2)
      batch.dispose()
      return { out, budget: budget.units - units, ceiling: ceiling.used - used, calls }
    }
    // Codex counterexample: 50 zeros are processed (ASCII, wildcard and numeric checks), then refused.
    const refused = await run('0'.repeat(50))
    expect(refused.out).toEqual({ kind: 'unresolved', reason: 'cf-formula-unsupported' })
    const key = cfCriterionUnits(50, CF_FREQUENCY_KEY_PASSES)
    expect(key).toBe(113)
    // View build (1) + the criterion read (1) + the key's processing, spent before cfCountifKey runs.
    expect(refused.budget).toBe(2 + key)
    expect(refused.ceiling).toBe(2 + key)
    // An admitted key of the same length is charged once, inside the admitted bound.
    const admitted = await run('a'.repeat(50))
    expect(admitted.out).toEqual({ kind: 'ready', value: { kind: 'boolean', value: false } })
    expect(admitted.calls[0]!.criterion).toBe(key)
    expect(admitted.budget).toBe(admitted.calls[0]!.total)
    expect(admitted.ceiling).toBe(admitted.calls[0]!.total)
  })
})
