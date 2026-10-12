/**
 * S13.5 (plan S8.1 line "Prepare the formatted display-text lane …"; spec §7,
 * AC-SEM-06, §12.4): the formatted display text of numeric/date/boolean cells is
 * prepared only when a text condition targets a text/mixed column, keyed by
 * valueGen and style owner, charged against the prepared-column budget with a
 * typed refusal, reused across applies and dropped on invalidation/eviction/close.
 * Results stay identical to the former eager preparation.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import * as render from '../src/xlsx/render'
import { createWorkerFilterEvaluator } from '../src/worker/viewerFilterEvaluator'
import { DisplayTextLaneCache, needsDisplayText } from '../src/worker/displayTextLane'
import { DocumentSession, PREPARED_COLUMN_OVERHEAD, PREPARED_STRING_OVERHEAD, type FilterEvaluation } from '../src/worker/viewerSession'
import { viewerControls } from '../src/components/xlsx/viewerControlsClient'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { emptyGridSelection } from '../src/xlsx/view/selection'
import { semanticsFor } from '../src/xlsx/view/semanticsFor'
import { classifyColumn, FilterEngine, type PreparedValue } from '../src/xlsx/view/filter'
import type { ColumnRule, FilterSpec, SheetView, SourceStore, TypedCondition, WorkbookFormatting } from '../src/xlsx/view/contracts'
import type { XlsxCellStyle, XlsxDocument, XlsxSheet } from '../src/xlsx/types'

vi.mock('../src/xlsx/render', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/render')>()
  return { ...actual, cellDisplayText: vi.fn(actual.cellDisplayText) }
})
const displayCalls = vi.mocked(render.cellDisplayText)

const styles: XlsxCellStyle[] = [{ numFmtId: 0 }, { numFmtId: 2 }, { numFmtId: 14 }, { numFmtId: 9 }, { numFmtId: 0, formatCode: '"n"0' }]
const formatting: WorkbookFormatting = { dateSystem: '1900', styles }
type Value = number | string | boolean | { error: string } | null
/** Body rows 1..N; column 0 numeric (0.00), 1 dates (numFmt 14), 2 mixed. */
const body: [Value, number, Value, number][] = [
  [1.5, 45306, 20, 0],
  [20, 45306.5, 0.2, 3],
  [300, 46089, 46089, 2],
  [-4, 46089.75, true, 0],
  [0, 45000, { error: '#DIV/0!' }, 0],
  [12.25, 45307, null, 0],
  [2, 45305.999, 'Straße', 0],
  [7, 46090, 'ss', 0],
  [20.5, 45306, 'café', 0],
  [99, 45306, 'café', 0],
  [3, 45306, '2x', 0],
  [4, 45306, 'TRUE', 0],
  [5, 45306, -1, 0],
  [6, 45306, 12, 1],
  [8, 45306, false, 0],
  [9, 45306, 21, 4],
]
const N = body.length
function cell(row: number, col: number, value: Value, styleIndex: number) {
  const error = value !== null && typeof value === 'object'
  return { ref: '', row, col, value: error ? value.error : value, valueIsError: error || undefined, styleIndex, style: styles[styleIndex] }
}
function makeDoc(): XlsxDocument {
  const rows = [{ index: 0, cells: [cell(0, 0, 'amount', 0), cell(0, 1, 'day', 0), cell(0, 2, 'mixed', 0)] }]
  body.forEach(([a, d, m, ms], i) => {
    const r = i + 1
    rows.push({ index: r, cells: [cell(r, 0, a, 1), cell(r, 1, d, 2), ...(m === null ? [] : [cell(r, 2, m, ms)])] })
  })
  return { sheets: [{ name: 'S', sheetId: '1', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows, semantics: semanticsFor('1900') } as never], tables: [] } as XlsxDocument
}
const range: FilterSpec['range'] = { r0: 0, r1: N, c0: 0, c1: 2, headerRows: 1, totalsRows: 0 }
const spec = (col: number, rule: ColumnRule): FilterSpec => ({ range, columns: new Map([[col, rule]]) })
const nowMs = Date.UTC(2024, 0, 15, 12)
const text = (op: 'contains' | 'eq' | 'neq' | 'not-contains' | 'begins' | 'ends', first: string, wildcards = false): TypedCondition => ({ kind: 'text', op, first, wildcards } as TypedCondition)
const mixedNonText = body.filter(([, , m]) => typeof m === 'number' || typeof m === 'boolean').length

function rowsOf(result: FilterEvaluation, filter: FilterSpec): number[] {
  if (result.kind !== 'complete') throw new Error(`not complete: ${JSON.stringify(result)}`)
  const rows: number[] = []
  for (let row = range.r0 + range.headerRows; row <= range.r1; row++) if (result.matches(row, filter)) rows.push(row)
  return rows
}
/** A store whose published valueGen can be advanced (a value publication). */
function bumpable(store: SourceStore): { store: SourceStore; bump(): void } {
  let gen = store.stamp.valueGen
  const stamp = () => ({ ...store.stamp, valueGen: gen })
  const proxy = new Proxy(store, {
    get(target, prop) {
      if (prop === 'stamp') return stamp()
      if (prop === 'queryColumn') return async (request: never, options: { signal: AbortSignal }) => ({ ...await target.queryColumn(request, { ...options, stamp: target.stamp }), stamp: stamp() })
      if (prop === 'ensureComplete') return (rect: never, options: { signal: AbortSignal }) => target.ensureComplete(rect, { ...options, stamp: target.stamp })
      const value = Reflect.get(target, prop)
      return typeof value === 'function' ? value.bind(target) : value
    },
  })
  return { store: proxy, bump: () => { gen++ } }
}
function controlsHarness(doc: XlsxDocument, lanes?: DisplayTextLaneCache) {
  const source = createObjectWorkbookSource(doc, 'doc-1')
  const order = createIdentityOrder()
  let view: SheetView = { sheetId: '1', revision: 0, order, baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null, selection: emptyGridSelection() }
  let viewGen = 0
  let filter: FilterSpec | null = null
  const owner = () => ({ docId: 'doc-1', sheetId: '1', valueGen: source.store('1').stamp.valueGen, fontGen: 0, viewGen })
  const snapshot = () => ({ owner: owner(), filter, sort: null, records: { visible: N + 1, total: N + 1, completeness: 'complete' as const } })
  const publish = (next: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>) => {
    view = { ...view, ...next }
    filter = next.filter
    viewGen++
    return snapshot()
  }
  return { source, snapshot, owner, view: () => view, open: () => viewerControls({ doc, source, sheetId: '1', snapshot, view: () => view, publish, ...(lanes ? { displayLanes: lanes } : {}) }, new AbortController().signal) }
}

beforeEach(() => { displayCalls.mockClear() })

describe('S13.5 formatted display-text lane is prepared on demand', () => {
  it('needsDisplayText: only text conditions (also inside compounds) on text/mixed columns', () => {
    const mixed = { kind: 'text' as const, hasNumeric: true }, numeric = { kind: 'number' as const, hasNumeric: true }
    expect(needsDisplayText(text('contains', '2'), mixed)).toBe(true)
    expect(needsDisplayText({ kind: 'compound', connective: 'or', first: { kind: 'number', op: 'gt', first: 3 }, second: text('eq', 'a') }, mixed)).toBe(true)
    expect(needsDisplayText({ kind: 'number', op: 'gt', first: 3 }, mixed)).toBe(false)
    expect(needsDisplayText(text('contains', '2'), numeric)).toBe(false)
    expect(needsDisplayText(text('contains', '2'), { kind: 'date', hasNumeric: true })).toBe(false)
    expect(needsDisplayText({ kind: 'values', selection: { mode: 'include', keys: { refId: 'r', count: 0, owner: { docId: 'd', sheetId: '1', valueGen: 0 } } }, showBlanks: false }, mixed)).toBe(false)
  })

  it('(1) worker: numeric and date filters over numeric/date columns never format display text', async () => {
    const source = createObjectWorkbookSource(makeDoc(), 'doc-1')
    const lanes = new DisplayTextLaneCache()
    const evaluate = createWorkerFilterEvaluator(() => new Map(), formatting, 'en-US', { lanes })
    const store = source.store('1')
    const numberSpec = spec(0, { kind: 'number', op: 'gt', first: 10 })
    expect(rowsOf(await evaluate({ store, spec: numberSpec, nowMs, signal: new AbortController().signal }), numberSpec)).toEqual([2, 3, 6, 9, 10])
    const dateSpec = spec(1, { kind: 'date', op: 'eq', first: { year: 2024, month: 1, day: 15 } })
    expect(rowsOf(await evaluate({ store, spec: dateSpec, nowMs, signal: new AbortController().signal }), dateSpec)).toEqual([1, 2, 9, 10, 11, 12, 13, 14, 15, 16])
    // A number rule on the mixed column needs no text either.
    const mixedNumber = spec(2, { kind: 'number', op: 'gte', first: 20 })
    expect(rowsOf(await evaluate({ store, spec: mixedNumber, nowMs, signal: new AbortController().signal }), mixedNumber)).toEqual([1, 3, 16])
    // A text rule on a number column cannot match numbers (§7): no lane either.
    const numericText = spec(0, text('contains', '2'))
    expect(rowsOf(await evaluate({ store, spec: numericText, nowMs, signal: new AbortController().signal }), numericText)).toEqual([])
    expect(displayCalls).not.toHaveBeenCalled()
    expect(lanes.builds).toBe(0)
    expect(lanes.bytes).toBe(0)
  })

  it('(1) main thread: numeric and date filters never format display text', async () => {
    const harness = controlsHarness(makeDoc())
    const controls = await harness.open()
    for (const [col, rule] of [[0, { kind: 'number', op: 'gt', first: 10 }], [1, { kind: 'datePreset', preset: 'today' }]] as const) {
      const outcome = await controls.client.filter(harness.owner(), spec(col, rule as ColumnRule), new AbortController().signal)
      expect(outcome.kind).toBe('committed')
    }
    expect(displayCalls).not.toHaveBeenCalled()
    controls.dispose()
  })

  it('(2) worker: a text contains rule on a mixed column builds the lane once and reuses it', async () => {
    const source = createObjectWorkbookSource(makeDoc(), 'doc-1')
    const lanes = new DisplayTextLaneCache()
    const evaluate = createWorkerFilterEvaluator(() => new Map(), formatting, 'en-US', { lanes })
    const store = source.store('1')
    const contains = spec(2, text('contains', '2'))
    const first = rowsOf(await evaluate({ store, spec: contains, nowMs, signal: new AbortController().signal }), contains)
    expect(displayCalls).toHaveBeenCalledTimes(mixedNonText)
    expect(lanes.builds).toBe(1)
    // A different text rule on the same column reuses the same lane.
    const second = rowsOf(await evaluate({ store, spec: contains, nowMs, signal: new AbortController().signal }), contains)
    const ends = spec(2, text('ends', 'E'))
    rowsOf(await evaluate({ store, spec: ends, nowMs, signal: new AbortController().signal }), ends)
    expect(second).toEqual(first)
    expect(displayCalls).toHaveBeenCalledTimes(mixedNonText)
    expect(lanes.builds).toBe(1)
    expect(lanes.size).toBe(1)
  })

  it('(2) main thread: a text rule builds the lane once across applies and dispose drops it', async () => {
    const lanes = new DisplayTextLaneCache()
    const harness = controlsHarness(makeDoc(), lanes)
    const controls = await harness.open()
    expect((await controls.client.filter(harness.owner(), spec(2, text('contains', '2')), new AbortController().signal)).kind).toBe('committed')
    expect(displayCalls).toHaveBeenCalledTimes(mixedNonText)
    expect((await controls.client.filter(harness.owner(), spec(2, text('begins', '2')), new AbortController().signal)).kind).toBe('committed')
    expect(displayCalls).toHaveBeenCalledTimes(mixedNonText)
    expect(lanes.builds).toBe(1)
    expect(lanes.bytes).toBeGreaterThan(0)
    controls.dispose()
    expect(lanes.bytes).toBe(0)
    expect(lanes.size).toBe(0)
  })

  it('(3) a valueGen bump invalidates and rebuilds the lane, releasing the dead one', async () => {
    const source = createObjectWorkbookSource(makeDoc(), 'doc-1')
    const lanes = new DisplayTextLaneCache()
    const evaluate = createWorkerFilterEvaluator(() => new Map(), formatting, 'en-US', { lanes })
    const { store, bump } = bumpable(source.store('1'))
    const contains = spec(2, text('contains', '2'))
    const before = rowsOf(await evaluate({ store, spec: contains, nowMs, signal: new AbortController().signal }), contains)
    const bytes = lanes.bytes
    bump()
    const after = rowsOf(await evaluate({ store, spec: contains, nowMs, signal: new AbortController().signal }), contains)
    expect(after).toEqual(before)
    expect(lanes.builds).toBe(2)
    expect(displayCalls).toHaveBeenCalledTimes(2 * mixedNonText)
    expect(lanes.size).toBe(1)
    expect(lanes.bytes).toBe(bytes)
  })

  it('(3) a style owner change invalidates and rebuilds the lane', async () => {
    const lanes = new DisplayTextLaneCache()
    const key = { docId: 'd', sheetId: '1', valueGen: 1, col: 2, r0: 1, r1: 4 }
    const owner1 = {}, owner2 = {}
    let fills = 0
    const fill = async (sink: { add(offset: number, text: string): void }) => { fills++; sink.add(0, 'abc'); sink.add(2, 'abc'); sink.add(3, 'de') }
    const { lane, release } = await lanes.lease({ ...key, styleOwner: owner1 }, fill)
    release()
    expect([lane.text(0), lane.text(1), lane.text(2), lane.text(3)]).toEqual(['abc', undefined, 'abc', 'de'])
    void (await lanes.lease({ ...key, styleOwner: owner1 }, fill)).release()
    expect(fills).toBe(1)
    void (await lanes.lease({ ...key, styleOwner: owner2 }, fill)).release()
    expect(fills).toBe(2)
    expect(lanes.size).toBe(1)
    // The worker evaluator's style owner is its style table: another table rebuilds.
    const source = createObjectWorkbookSource(makeDoc(), 'doc-1')
    const shared = new DisplayTextLaneCache()
    const contains = spec(2, text('contains', '2'))
    for (const table of [formatting, { ...formatting, styles: [...styles] }]) {
      await createWorkerFilterEvaluator(() => new Map(), table, 'en-US', { lanes: shared })({ store: source.store('1'), spec: contains, nowMs, signal: new AbortController().signal })
    }
    expect(shared.builds).toBe(2)
    expect(shared.size).toBe(1)
  })

  it('(4) lane bytes are charged exactly and released; the cache never exceeds its limit', async () => {
    const lanes = new DisplayTextLaneCache()
    const key = { docId: 'd', sheetId: '1', valueGen: 1, styleOwner: {}, col: 2, r0: 1, r1: 10 }
    void (await lanes.lease(key, async sink => { sink.add(0, 'abc'); sink.add(1, 'abc'); sink.add(5, 'xy') })).release()
    const expected = PREPARED_COLUMN_OVERHEAD + 10 * 4 + (3 * 2 + PREPARED_STRING_OVERHEAD) + (2 * 2 + PREPARED_STRING_OVERHEAD)
    expect(lanes.bytes).toBe(expected)
    lanes.invalidate()
    expect(lanes.bytes).toBe(0)
    // Typed refusal: a lane that cannot fit refuses with quota and holds nothing.
    const tiny = new DisplayTextLaneCache({ limitBytes: PREPARED_COLUMN_OVERHEAD + 40 + 10 })
    const refusal = await tiny.lease(key, async sink => { sink.add(0, 'abcdefgh') }).catch(error => error)
    expect(refusal).toMatchObject({ name: 'ValueIndexError', code: 'quota', detail: { resource: 'memory', limitBytes: PREPARED_COLUMN_OVERHEAD + 50 } })
    expect(refusal.detail.requiredBytes).toBeGreaterThan(PREPARED_COLUMN_OVERHEAD + 50)
    expect(tiny.bytes).toBe(0)
    // A second lane evicts the least recently used one rather than exceeding the limit.
    const two = new DisplayTextLaneCache({ limitBytes: 2 * (PREPARED_COLUMN_OVERHEAD + 40) + 60 })
    void (await two.lease({ ...key, col: 1 }, async sink => { sink.add(0, 'a') })).release()
    const held = await two.lease({ ...key, col: 2 }, async sink => { sink.add(0, 'b') })
    expect(two.size).toBe(1)
    expect(two.bytes).toBeLessThanOrEqual(two.limitBytes)
    // A leased lane is never evicted for room: the next build refuses instead.
    const blocked = await two.lease({ ...key, col: 3 }, async sink => { sink.add(0, 'c') }).catch(error => error)
    expect(blocked).toMatchObject({ code: 'quota' })
    expect(held.lane.text(0)).toBe('b')
    expect(two.size).toBe(1)
    held.release()
    void (await two.lease({ ...key, col: 3 }, async sink => { sink.add(0, 'c') })).release()
    expect(two.size).toBe(1)
    expect(two.bytes).toBeLessThanOrEqual(two.limitBytes)
  })

  it('(4) concurrent builds of one key keep a single charged lane', async () => {
    const lanes = new DisplayTextLaneCache()
    const key = { docId: 'd', sheetId: '1', valueGen: 1, styleOwner: {}, col: 2, r0: 1, r1: 10 }
    let open!: () => void
    const gate = new Promise<void>(resolve => { open = resolve })
    const fill = async (sink: { add(offset: number, text: string): void }) => { await gate; sink.add(0, 'abc') }
    const first = lanes.lease(key, fill), second = lanes.lease(key, fill)
    open()
    const [a, b] = await Promise.all([first, second])
    expect(lanes.builds).toBe(1)
    expect(lanes.size).toBe(1)
    expect(lanes.bytes).toBe(PREPARED_COLUMN_OVERHEAD + 10 * 4 + 3 * 2 + PREPARED_STRING_OVERHEAD)
    expect(a.lane).toBe(b.lane)
    a.release(); b.release()
  })

  it('(4) worker session: the lane is retained in the worker pool, refused by the prepared budget, dropped at close', async () => {
    const source = createObjectWorkbookSource(makeDoc(), 'doc-1')
    const lanes = new DisplayTextLaneCache()
    const session = new DocumentSession('doc-1', source, { filterEvaluator: createWorkerFilterEvaluator(() => new Map(), formatting, 'en-US', { lanes }) })
    const before = session.memoryStats().retainedBytes
    const outcome = await session.filter(session.identity('1'), spec(2, text('contains', '2')), nowMs, new AbortController().signal)
    expect(outcome.kind).toBe('committed')
    expect(lanes.bytes).toBeGreaterThan(0)
    const viewBytes = session.memoryStats().viewBytes
    expect(session.memoryStats().retainedBytes - before - viewBytes).toBeGreaterThanOrEqual(lanes.bytes)
    await session.release()
    expect(lanes.bytes).toBe(0)
    expect(lanes.size).toBe(0)

    // The prepared-column budget bounds the lane: exhaustion is a typed refusal.
    const small = new DisplayTextLaneCache()
    const tight = new DocumentSession('doc-2', createObjectWorkbookSource(makeDoc(), 'doc-2'), {
      preparedBudget: { maxColumns: 8, maxBytes: PREPARED_COLUMN_OVERHEAD + 8 },
      filterEvaluator: createWorkerFilterEvaluator(() => new Map(), formatting, 'en-US', { lanes: small }),
    })
    const refused = await tight.filter(tight.identity('1'), spec(2, text('contains', '2')), nowMs, new AbortController().signal)
    expect(refused).toMatchObject({ kind: 'refused', refusal: { reason: 'resource-limit' } })
    expect(small.bytes).toBe(0)
    // Numeric rules never need the lane, so the same budget still admits them.
    expect((await tight.filter(tight.identity('1'), spec(0, { kind: 'number', op: 'gt', first: 10 }), nowMs, new AbortController().signal)).kind).toBe('committed')
    await tight.release()
  })

  it('(4) main thread: lane budget exhaustion is a typed resource-limit refusal', async () => {
    const harness = controlsHarness(makeDoc(), new DisplayTextLaneCache({ limitBytes: PREPARED_COLUMN_OVERHEAD + 8 }))
    const controls = await harness.open()
    const outcome = await controls.client.filter(harness.owner(), spec(2, text('contains', '2')), new AbortController().signal)
    expect(outcome).toMatchObject({ kind: 'refused', refusal: { reason: 'resource-limit' } })
    controls.dispose()
  })

  it('(5) results equal the former eager preparation for every typed rule kind', async () => {
    const doc = makeDoc()
    const sheet = { name: '', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor('1900') } as XlsxSheet
    /** The pre-S13.5 preparer: display text for every numeric/boolean cell. */
    const eager = (col: number): PreparedValue[] => body.map((row): PreparedValue => {
      const value = row[col === 0 ? 0 : col === 1 ? 1 : 2], style = styles[col === 0 ? 1 : col === 1 ? 2 : row[3]]
      if (value === null) return { kind: 'blank' }
      if (typeof value === 'object') return { kind: 'error', value: value.error }
      if (typeof value === 'string') return { kind: 'string', value }
      const displayText = render.cellDisplayText(sheet, { value, style })
      return typeof value === 'boolean' ? { kind: 'boolean', value, displayText } : { kind: 'number', value, displayText, numFmtId: style.numFmtId, formatCode: style.formatCode }
    })
    const rules: ColumnRule[] = [
      text('contains', '2'), text('eq', 'ss'), text('eq', 'STRASSE'), text('neq', 'true'), text('not-contains', '2'),
      text('begins', 'caf'), text('ends', 'é'), text('eq', 'café'), text('contains', '%'), text('eq', '2?', true), text('eq', '*E', true), text('contains', '~*', true),
      text('eq', 'n21'), text('eq', '12.00'), text('contains', '/'),
      { kind: 'number', op: 'gt', first: 3 }, { kind: 'number', op: 'neq', first: 20 }, { kind: 'number', op: 'between', first: -1, second: 20 },
      { kind: 'date', op: 'eq', first: { year: 2024, month: 1, day: 15 } }, { kind: 'date', op: 'before', first: { year: 2024, month: 1, day: 16 } },
      { kind: 'datePreset', preset: 'today' }, { kind: 'datePreset', preset: 'thisYear' },
      { kind: 'compound', connective: 'or', first: { kind: 'number', op: 'lt', first: 0 }, second: text('contains', 'TRUE') },
      { kind: 'compound', connective: 'and', first: text('contains', '2'), second: { kind: 'number', op: 'gt', first: 1 } },
    ]
    const source = createObjectWorkbookSource(doc, 'doc-1')
    const evaluate = createWorkerFilterEvaluator(() => new Map(), formatting, 'en-US')
    const context = { locale: 'en-US', dateSystem: '1900' as const, timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone, nowMs }
    let compared = 0
    for (const col of [0, 1, 2]) {
      const values = eager(col)
      const classification = classifyColumn(values, true)
      for (const rule of rules) {
        let expected: number[] | string
        try {
          const match = new FilterEngine(context).compileRule(rule, classification)
          expected = values.flatMap((value, i) => match(value) ? [i + 1] : [])
        } catch (error) { expected = String((error as Error).message) }
        const filter = spec(col, rule)
        const result = await evaluate({ store: source.store('1'), spec: filter, nowMs, signal: new AbortController().signal })
        const actual = typeof expected === 'string' ? (result.kind === 'refused' ? (result.refusal as { message?: string }).message ?? '' : `unexpected ${result.kind}`) : rowsOf(result, filter)
        expect(actual, `col ${col} ${JSON.stringify(rule)}`).toEqual(expected)
        compared++
      }
    }
    expect(compared).toBe(3 * rules.length)
  })
})
