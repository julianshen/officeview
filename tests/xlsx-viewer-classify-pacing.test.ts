/** Column classification passes yield like every other long job (spec §10.2;
 * plan S4.2 LIFE-02: "Long jobs yield at least every 10ms or 4,096 rows,
 * whichever arrives first"). Pages resolve without a macrotask turn, so only a
 * real yield lets a cancel posted as a task run. */
import { describe, expect, it } from 'vitest'
import { createWorkerFilterEvaluator } from '../src/worker/viewerFilterEvaluator'
import { viewerControls } from '../src/components/xlsx/viewerControlsClient'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { emptyGridSelection } from '../src/xlsx/view/selection'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { CheckpointBudget, CooperativeSlice } from '../src/xlsx/view/scheduler'
import { SOURCE_NUMBER, classifyLanePages, type ValueLanePage } from '../src/xlsx/view/valueIndex'
import type { FilterSpec, SheetView, SourceStore } from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'

function workbook(rows: number): XlsxDocument {
  return { sheets: [{ name: 's', sheetId: 's', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900'),
    rows: Array.from({ length: rows + 1 }, (_, row) => ({ index: row, cells: [{ row, col: 0, ref: '', value: row ? row : 'Header', styleIndex: 0 }] })) }], tables: [] } as never
}
const range = (rows: number): FilterSpec['range'] => ({ r0: 0, r1: rows, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 })
/** Counts column rows requested from the store (a page is read just before it is processed). */
function counted(store: SourceStore, read: { rows: number }, owner?: { gen: number }): SourceStore {
  return new Proxy(store, {
    get(target, prop) {
      if (prop === 'stamp' && owner) return { ...target.stamp, valueGen: target.stamp.valueGen + owner.gen }
      if (prop === 'queryColumn') return (request: { rows: { intervals: { r0: number; r1: number }[] } }, options: never) => {
        for (const { r0, r1 } of request.rows.intervals) read.rows += r1 - r0 + 1
        return target.queryColumn(request as never, options)
      }
      const value = Reflect.get(target, prop)
      return typeof value === 'function' ? value.bind(target) : value
    },
  })
}
/** Records rows read at every macrotask turn until stopped. */
function turns(read: { rows: number }): { marks: number[]; stop(): void } {
  const marks: number[] = []
  let live = true
  const tick = () => { marks.push(read.rows); if (live) setImmediate(tick) }
  setImmediate(tick)
  return { marks, stop: () => { live = false } }
}
const gaps = (marks: readonly number[]): number[] => marks.slice(1).map((at, i) => at - marks[i])
// A values rule whose keys were not shipped refuses right after classification,
// so classification is the only column pass this evaluation runs.
const unshipped = (rows: number): FilterSpec => ({ range: range(rows), columns: new Map([[0, {
  kind: 'values', showBlanks: false, selection: { mode: 'include', keys: { refId: 'missing', count: 1, owner: { docId: 'd', sheetId: 's', valueGen: 0 } } },
}]]) })

describe('worker filter evaluator classification pacing', () => {
  it('yields after at most 4,096 classified rows while the injected clock never advances', async () => {
    const rows = 10_240, read = { rows: 0 }, slices: number[] = []
    let since = 0
    const store = counted(createObjectWorkbookSource(workbook(rows), 'd').store('s'), read)
    const cooperate = new CooperativeSlice<string>({ now: () => 0, yield: async () => { slices.push(read.rows - since); since = read.rows } })
    const evaluate = createWorkerFilterEvaluator(() => new Map())
    const outcome = await evaluate({ store, spec: unshipped(rows), nowMs: 0, signal: new AbortController().signal, cooperate })
    expect(outcome.kind).toBe('refused')
    slices.push(read.rows - since)
    expect(read.rows).toBe(rows)
    expect(slices.length - 1).toBeGreaterThanOrEqual(2)
    expect(Math.max(...slices)).toBeLessThanOrEqual(4096)
  })
  it('observes a cancel posted as a macrotask within one 4,096-row slice without a session slice', async () => {
    const rows = 40_960, read = { rows: 0 }, controller = new AbortController()
    const store = counted(createObjectWorkbookSource(workbook(rows), 'd').store('s'), read)
    const evaluate = createWorkerFilterEvaluator(() => new Map())
    setImmediate(() => controller.abort())
    const outcome = await evaluate({ store, spec: { range: range(rows), columns: new Map([[0, { kind: 'number', op: 'gt', first: 10 }]]) }, nowMs: 0, signal: controller.signal })
    expect(outcome.kind).toBe('cancelled')
    expect(read.rows).toBeLessThanOrEqual(4096)
  })
})

describe('main-thread controls classification pacing', () => {
  function open(rows: number, read: { rows: number }) {
    const doc = workbook(rows), base = createObjectWorkbookSource(doc, 'd')
    const source = { ...base, store: (id: string) => counted(base.store(id), read) } as typeof base
    let view: SheetView = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null, selection: emptyGridSelection() }
    const owner = () => ({ docId: 'd', sheetId: 's', valueGen: base.store('s').stamp.valueGen, fontGen: 0, viewGen: 0 })
    const snapshot = () => ({ owner: owner(), filter: null, sort: null, records: { visible: rows + 1, total: rows + 1, completeness: 'complete' as const } })
    const publish = (next: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>) => { view = { ...view, ...next }; return snapshot() }
    return { owner, open: () => viewerControls({ doc, source, sheetId: 's', snapshot, view: () => view, publish }, new AbortController().signal) }
  }
  it('gives the host a macrotask turn at least every 4,096 classified rows', async () => {
    const rows = 10_240, read = { rows: 0 }, harness = open(rows, read)
    const controls = await harness.open()
    read.rows = 0
    const probe = turns(read)
    const draft = await controls.client.values(harness.owner(), range(rows), 0, null, new AbortController().signal)
    probe.stop()
    draft.cancel()
    controls.dispose()
    expect(read.rows).toBeGreaterThanOrEqual(rows)
    expect(probe.marks.length).toBeGreaterThanOrEqual(2)
    expect(Math.max(...gaps([0, ...probe.marks, read.rows]))).toBeLessThanOrEqual(4096)
  })
  it('observes a cancel posted as a macrotask within one 4,096-row slice', async () => {
    const rows = 40_960, read = { rows: 0 }, harness = open(rows, read), controller = new AbortController()
    const controls = await harness.open()
    read.rows = 0
    setImmediate(() => controller.abort())
    const error = await controls.client.values(harness.owner(), range(rows), 0, null, controller.signal).then(() => null, cause => cause)
    controls.dispose()
    expect(error).not.toBeNull()
    expect(read.rows).toBeLessThanOrEqual(4096)
  })
})

describe('shared column-pass pacing', () => {
  it('keeps uneven pages and work carried from an earlier pass within 4,096 rows per slice', async () => {
    const slices: number[] = []
    let classified = 0, since = 0
    const cooperate = new CooperativeSlice<'filter'>({ now: () => 0, yield: async () => { slices.push(classified - since); since = classified } })
    // An earlier pass left 3,000 rows unyielded on the shared slice.
    expect(cooperate.due(3000)).toBe(false)
    const page = (rows: number): ValueLanePage => ({ kind: 'values', pageId: 'p', count: rows, tags: new Uint8Array(rows).fill(SOURCE_NUMBER), numbers: new Float64Array(rows), texts: [] })
    const classification = await classifyLanePages(async accept => { for (let p = 0; p < 11; p++) await accept(page(1000)) }, (lane, index) => {
      classified++
      return { kind: 'number', value: lane.numbers[index] }
    }, () => {}, cooperate)
    expect(classification).toEqual({ kind: 'number', hasNumeric: true })
    slices.push(classified - since)
    expect(slices[0]).toBe(0) // the carried work yielded before this pass started
    expect(slices.length - 1).toBeGreaterThanOrEqual(3)
    expect(Math.max(...slices)).toBeLessThanOrEqual(4096)
  })
})

describe('review round (codex R1, R3, R4)', () => {
  const numberRule = (rows: number): FilterSpec => ({ range: range(rows), columns: new Map([[0, { kind: 'number', op: 'gt', first: 10 }]]) })
  it('splits one 8,192-entry page: the first yield comes before 8,192 entries are classified (R1)', async () => {
    let classified = 0
    const firsts: number[] = []
    const cooperate = new CooperativeSlice<'filter'>({ now: () => 0, yield: async () => { firsts.push(classified) } })
    const page: ValueLanePage = { kind: 'values', pageId: 'p', count: 8192, tags: new Uint8Array(8192).fill(SOURCE_NUMBER), numbers: new Float64Array(8192), texts: [] }
    await classifyLanePages(async accept => { await accept(page) }, (lane, index) => { classified++; return { kind: 'number', value: lane.numbers[index] } }, () => {}, cooperate)
    expect(classified).toBe(8192)
    expect(firsts[0]).toBeLessThanOrEqual(4096)
  })
  it('cancels a warm-cache filter with no session slice through a queued macrotask abort (R3)', async () => {
    const rows = 10_240, read = { rows: 0 }, controller = new AbortController()
    const store = counted(createObjectWorkbookSource(workbook(rows), 'd').store('s'), read)
    const evaluate = createWorkerFilterEvaluator(() => new Map())
    // Warm the classification cache (the values rule then refuses for missing keys).
    expect((await evaluate({ store, spec: unshipped(rows), nowMs: 0, signal: new AbortController().signal })).kind).toBe('refused')
    read.rows = 0
    setImmediate(() => controller.abort())
    expect((await evaluate({ store, spec: numberRule(rows), nowMs: 0, signal: controller.signal })).kind).toBe('cancelled')
    expect(read.rows).toBeLessThanOrEqual(4096)
  })
  it('reports stale-source when the owner changes during a classification pause, and caches nothing (R4)', async () => {
    // One page: no later page read would notice the change.
    const rows = 4096, read = { rows: 0 }, owner = { gen: 0 }
    const store = counted(createObjectWorkbookSource(workbook(rows), 'd').store('s'), read, owner)
    const evaluate = createWorkerFilterEvaluator(() => new Map())
    let pauses = 0
    const cooperate = new CooperativeSlice<string>({ now: () => 0, yield: async () => { if (++pauses === 1) owner.gen = 1 } })
    expect(await evaluate({ store, spec: unshipped(rows), nowMs: 0, signal: new AbortController().signal, cooperate })).toMatchObject({ kind: 'error', error: { code: 'stale-source' } })
    owner.gen = 0
    read.rows = 0
    expect((await evaluate({ store, spec: unshipped(rows), nowMs: 0, signal: new AbortController().signal })).kind).toBe('refused')
    // The column was classified again: no stale classification was cached.
    expect(read.rows).toBe(rows)
  })
  it.each(['stale', 'abort'] as const)('re-checks the owner and cancellation after the final match pause (R4, %s)', async change => {
    const rows = 4096, owner = { gen: 0 }, controller = new AbortController()
    const store = counted(createObjectWorkbookSource(workbook(rows), 'd').store('s'), { rows: 0 }, owner)
    const evaluate = createWorkerFilterEvaluator(() => new Map())
    let pauses = 0
    // Pause 1 ends classification; pause 2 follows the last matched row.
    const cooperate = new CooperativeSlice<string>({ now: () => 0, yield: async () => {
      if (++pauses === 2) { if (change === 'stale') owner.gen = 1; else controller.abort() }
    } })
    const outcome = await evaluate({ store, spec: numberRule(rows), nowMs: 0, signal: controller.signal, cooperate })
    expect(pauses).toBe(2)
    if (change === 'stale') expect(outcome).toMatchObject({ kind: 'error', error: { code: 'stale-source' } })
    else expect(outcome.kind).toBe('cancelled')
  })
  it('main thread: an owner change during a classification yield is stale-source and caches nothing (R4)', async () => {
    const rows = 4096, read = { rows: 0 }, owner = { gen: 0 }
    const doc = workbook(rows), base = createObjectWorkbookSource(doc, 'd')
    const source = { ...base, store: (id: string) => counted(base.store(id), read, owner) } as typeof base
    let view: SheetView = { sheetId: 's', revision: 0, order: createIdentityOrder(), baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null, selection: emptyGridSelection() }
    const ownerOf = () => ({ docId: 'd', sheetId: 's', valueGen: base.store('s').stamp.valueGen, fontGen: 0, viewGen: 0 })
    const snapshot = () => ({ owner: ownerOf(), filter: null, sort: null, records: { visible: rows + 1, total: rows + 1, completeness: 'complete' as const } })
    const controls = await viewerControls({ doc, source, sheetId: 's', snapshot, view: () => view, publish: next => { view = { ...view, ...next }; return snapshot() } }, new AbortController().signal)
    read.rows = 0
    setImmediate(() => { owner.gen = 1 })
    const error = await controls.client.values(ownerOf(), range(rows), 0, null, new AbortController().signal).then(() => null, cause => cause)
    expect(error).toMatchObject({ code: 'stale-source' })
    owner.gen = 0
    read.rows = 0
    const draft = await controls.client.values(ownerOf(), range(rows), 0, null, new AbortController().signal)
    expect(read.rows).toBeGreaterThanOrEqual(rows)
    draft.cancel()
    controls.dispose()
  })
})

describe('checkpoints within 256 classification units (spec §12.1, codex C2)', () => {
  const page = (count: number): ValueLanePage => ({ kind: 'values', pageId: 'p', count, tags: new Uint8Array(count).fill(SOURCE_NUMBER), numbers: new Float64Array(count), texts: [] })
  it('checks cancellation within 256 classified entries', async () => {
    const controller = new AbortController()
    let processed = 0
    const cooperate = new CooperativeSlice<'filter'>({ now: () => 0, yield: async () => {} })
    await expect(classifyLanePages(async accept => accept(page(4096)), () => { processed++; if (processed === 1) controller.abort(); return { kind: 'number', value: 1 } }, () => controller.signal.throwIfAborted(), cooperate)).rejects.toBeDefined()
    expect(processed).toBeLessThanOrEqual(256)
  })
  it('samples the time budget within 256 entries of one page', async () => {
    let clock = 0, at = 0
    const marks: number[] = []
    const cooperate = new CooperativeSlice<'filter'>({ now: () => clock, yield: async () => { marks.push(clock - at); at = clock } })
    await classifyLanePages(async accept => accept(page(4096)), () => { clock++; return { kind: 'number', value: 1 } }, () => {}, cooperate)
    expect(marks[0]).toBeLessThanOrEqual(256)
  })
  it('carries a short tail into the next pass sharing the slice (codex D2)', async () => {
    let processed = 0, last = 0
    const gaps: number[] = []
    const cooperate = new CooperativeSlice<'filter'>({ now: () => 0, yield: async () => { gaps.push(processed - last); last = processed } })
    await classifyLanePages(async accept => accept(page(4095)), () => { processed++; return { kind: 'number', value: 1 } }, () => {}, cooperate)
    for (let i = 0; i < 512; i++) { processed++; if (cooperate.due(1)) await cooperate.pause('filter') }
    expect(gaps[0]).toBeLessThanOrEqual(4096)
  })
  it('keeps the operation start as the owned slice clock origin (codex D2)', () => {
    let now = 0
    const budget = new CheckpointBudget(() => {}, undefined, { now: () => now, yield: async () => {} })
    try { now = 256; expect(budget.spend(256)).toBe(true) } finally { budget.dispose() }
  })
})
