import { describe, expect, it, vi } from 'vitest'
import type {
  CacheIdentity, FilterSpec, SortSpec, SourceStore, ViewCommitDelta, WorkbookOwner, WorkerCommand, WorkerReply,
} from '../src/xlsx/view/contracts'
import type { XlsxDocument, XlsxMergeRange } from '../src/xlsx/types'
import { formatA1 } from '../src/xlsx/view/coordinates'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import {
  bitsetsOf, initialIdentity, initialWorkbookOwner, isViewerReply, sourceBitset, transferablesOf, validateBitset,
  visibilityFromDisplayBitsets,
} from '../src/worker/viewerProtocol'
import {
  cacheKey, DocumentSession, PREPARED_COLUMN_OVERHEAD, PREPARED_ENTRY_LANE_BYTES, workbookIndexKey, type FilterEvaluator,
  type ViewDraft, type WorkbookSource,
} from '../src/worker/viewerSession'
import { createObjectWorkbookSource, READ_ROWS_MAX_CELLS } from '../src/worker/viewerObjectSource'
import { connectInProcess, ViewerKernel } from '../src/worker/viewerKernel'
import { createViewerClient } from '../src/worker/viewerClient'
import { createWorkerMessageHandler } from '../src/worker/worker-entry'

const quiet = { yield: async () => {} }
type Value = number | string | null
interface SheetInput { sheetId: string; name?: string; rows: Value[][]; hidden?: number[]; merges?: XlsxMergeRange[] }
function objectDoc(sheets: SheetInput[]): XlsxDocument {
  return {
    sheets: sheets.map(sheet => ({
      name: sheet.name ?? `Sheet ${sheet.sheetId}`,
      sheetId: sheet.sheetId,
      rows: sheet.rows.map((values, row) => ({
        index: row,
        ...(sheet.hidden?.includes(row) ? { hidden: true } : {}),
        cells: values.flatMap((value, col) => value === null ? [] : [{
          ref: formatA1({ row, col }), row, col, value, styleIndex: 0, hasCachedValue: true,
        }]),
      })),
      cols: [], merges: [], mergeRanges: sheet.merges ?? [],
    })),
    // A fully inspected inventory with no table parts (undefined = unavailable).
    tables: [],
  }
}
function deepFreeze<T>(value: T): T {
  if (typeof value === 'object' && value !== null && !Object.isFrozen(value)) {
    Object.freeze(value)
    for (const entry of Object.values(value)) deepFreeze(entry)
  }
  return value
}
const column = (values: Value[]): Value[][] => values.map(value => [value])
const sortSpec = (r1: number, c1: number, col = 0, order: 'asc' | 'desc' = 'asc'): SortSpec => ({
  range: { r0: 0, r1, c0: 0, c1, headerRows: 0, totalsRows: 0 }, keys: [{ col, order }], caseSensitive: false,
})
let nextId = 0
const rid = () => `t-${++nextId}`
async function run<K extends WorkerCommand['kind']>(kernel: ViewerKernel, command: Extract<WorkerCommand, { kind: K }>, log?: WorkerReply[]): Promise<WorkerReply<K>> {
  const replies: WorkerReply[] = []
  await kernel.handle(command, out => { replies.push(out); log?.push(out) })
  expect(replies.length).toBeGreaterThan(0)
  return replies[replies.length - 1] as WorkerReply<K>
}
function completed<T extends WorkerReply>(out: T): Extract<T, { kind: 'complete' }>['result'] {
  if (out.kind !== 'complete') throw new Error(`expected complete, got ${JSON.stringify(out)}`)
  return (out as Extract<T, { kind: 'complete' }>).result
}
function deltaOf(out: WorkerReply): ViewCommitDelta {
  const result = completed(out)
  if (result.kind !== 'view') throw new Error(`expected view, got ${result.kind}`)
  return result.delta
}
async function workbookBytes(): Promise<Uint8Array<ArrayBuffer>> {
  return new Uint8Array(await buildXlsx([
    { name: 'Data', rows: [3, 1, 4, 1, 5].map((v, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, v }, { ref: `B${i + 1}`, v: i, t: 's' as const }] })) },
    { name: 'Other', rows: [9, 7, 8].map((v, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, v }] })) },
  ], ['e', 'd', 'c', 'b', 'a']))
}
function countingParse() {
  const counter = { parses: 0 }
  return { counter, parse: async (bytes: Uint8Array) => { counter.parses++; return parseXlsx(await OfficePackage.load(bytes)) } }
}
/** Wraps a workbook source so tests can count store reads per sheet. */
function instrumented(source: WorkbookSource) {
  const calls = { queryColumn: 0 }
  const stores = new Map<string, SourceStore>()
  const wrapped: WorkbookSource = {
    valuePolicy: source.valuePolicy, sheets: source.sheets, geometry: id => source.geometry(id), release: () => source.release(),
    store(sheetId) {
      let store = stores.get(sheetId)
      if (!store) {
        const inner = source.store(sheetId)
        store = Object.create(inner) as SourceStore
        store.queryColumn = (request, options) => { calls.queryColumn++; return inner.queryColumn(request, options) }
        stores.set(sheetId, store)
      }
      return store
    },
  }
  return { calls, wrapped }
}

describe('resident viewer worker (S4.1)', () => {
  it('opens once: sheet switches, prepares and queries never reparse or reread the source', async () => {
    const { counter, parse } = countingParse()
    const kernel = new ViewerKernel({ parse }, quiet)
    const blob = new Blob([await workbookBytes()])
    const read = vi.spyOn(blob, 'arrayBuffer')
    const open = completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'blob', blob }, profile: 'D' }))
    expect(open).toEqual({ kind: 'open', valuePolicy: 'engine', sheets: [{ sheetId: '1', name: 'Data' }, { sheetId: '2', name: 'Other' }] })
    let workbook: WorkbookOwner = initialWorkbookOwner('doc', open.kind === 'open' ? open.sheets : [])
    let owner: CacheIdentity = initialIdentity('doc', '1')
    completed(await run(kernel, { kind: 'activateSheet', requestId: rid(), owner: workbook, sheetId: '2' }))
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0, 1]) }))
    for (let i = 0; i < 2; i++) {
      const page = completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, limit: 10 } }))
      expect(page.kind === 'valueQuery' && page.page.items.map(item => item.kind === 'cell' ? item.cell.value : item.kind)).toEqual([3, 1, 4, 1, 5])
    }
    const asc = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: sortSpec(4, 1) }))
    expect([...asc.bands[0].sourceAtOffset]).toEqual([1, 3, 0, 2, 4])
    owner = asc.owner
    const desc = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: sortSpec(4, 1, 0, 'desc') }))
    expect([...desc.bands[0].sourceAtOffset]).toEqual([4, 2, 0, 1, 3])
    owner = desc.owner
    workbook = desc.workbookOwner
    completed(await run(kernel, { kind: 'activateSheet', requestId: rid(), owner: workbook, sheetId: '1' }))
    owner = deltaOf(await run(kernel, { kind: 'filter', requestId: rid(), owner, spec: null, nowMs: 0 })).owner
    completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 1, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, limit: 10 } }))
    expect(counter.parses).toBe(1)
    expect(kernel.stats).toEqual({ opens: 1, parses: 1, sourceReads: 1 })
    expect(read).toHaveBeenCalledTimes(1)
    expect(owner).toEqual({ docId: 'doc', sheetId: '1', valueGen: 0, fontGen: 0, viewGen: 3 })
  })

  it('prepared columns stay resident: repeated queries and sorts do not reread the store', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: [[5, 'e'], [2, 'b'], [4, 'd'], [1, 'a'], [3, 'c']] }])
    const tracked = instrumented(createObjectWorkbookSource(doc, 'doc'))
    const kernel = new ViewerKernel({ parse: async () => { throw new Error('no parse expected') }, openSource: async () => tracked.wrapped }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    let owner = initialIdentity('doc', 'A')
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) }))
    const afterPrepare = tracked.calls.queryColumn
    expect(afterPrepare).toBeGreaterThan(0)
    for (let i = 0; i < 3; i++) {
      completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 0, rows: { kind: 'list', r0: 0, r1: 4, rows: new Uint32Array([1, 3]) }, limit: 10 } }))
    }
    owner = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: sortSpec(4, 1) })).owner
    owner = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: sortSpec(4, 1, 0, 'desc') })).owner
    expect(tracked.calls.queryColumn).toBe(afterPrepare)
    // An unprepared key column does read the store, so the counter is live.
    deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: sortSpec(4, 1, 1) }))
    expect(tracked.calls.queryColumn).toBeGreaterThan(afterPrepare)
  })

  it('sort and filter run on the resident workbook without cloning it', async () => {
    const doc = deepFreeze(objectDoc([{ sheetId: 'A', rows: column([3, 1, 2]) }]))
    const clone = vi.spyOn(globalThis, 'structuredClone')
    try {
      const filterEvaluator: FilterEvaluator = async () => ({ kind: 'complete', matches: row => row !== 1 })
      const kernel = new ViewerKernel({ parse: async () => { throw new Error('no parse expected') } }, { ...quiet, filterEvaluator })
      completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
      let owner = initialIdentity('doc', 'A')
      owner = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: sortSpec(2, 0) })).owner
      const filter: FilterSpec = { range: { r0: 0, r1: 2, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
      const filtered = deltaOf(await run(kernel, { kind: 'filter', requestId: rid(), owner, spec: filter, nowMs: 0 }))
      expect(filtered.counts.visible).toBe(filtered.counts.total - 1)
      owner = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner: filtered.owner, spec: null })).owner
      expect(owner.viewGen).toBe(3)
      expect(kernel.stats.parses).toBe(0)
      expect(clone).not.toHaveBeenCalled()
    } finally {
      clone.mockRestore()
    }
  })

  it('two sheets at local revision 1 have different aggregate workbook identities', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([2, 1]) }, { sheetId: 'B', rows: column([2, 1]) }])
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const a = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc', 'A'), spec: sortSpec(1, 0) }))
    const b = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc', 'B'), spec: sortSpec(1, 0) }))
    expect(a.owner.viewGen).toBe(1)
    expect(b.owner.viewGen).toBe(1)
    expect(a.workbookOwner.globalViewGen).toBe(1)
    expect(b.workbookOwner.globalViewGen).toBe(2)
    expect(a.workbookOwner.sheetVector).toEqual([{ sheetId: 'A', revision: 1, valueGen: 0 }, { sheetId: 'B', revision: 0, valueGen: 0 }])
    expect(b.workbookOwner.sheetVector).toEqual([{ sheetId: 'A', revision: 1, valueGen: 0 }, { sheetId: 'B', revision: 1, valueGen: 0 }])
    expect(workbookIndexKey(a.workbookOwner)).not.toBe(workbookIndexKey(b.workbookOwner))
    // Separate histories with equal global generations: only the vector separates "A=1" from "B=1".
    const onlyA = new DocumentSession('doc', createObjectWorkbookSource(doc, 'doc'))
    const onlyB = new DocumentSession('doc', createObjectWorkbookSource(doc, 'doc'))
    await onlyA.commitView('A', draftOf(onlyA, 'A'))
    await onlyB.commitView('B', draftOf(onlyB, 'B'))
    expect(onlyA.globalViewGen).toBe(onlyB.globalViewGen)
    expect(workbookIndexKey(onlyA.workbookOwner())).not.toBe(workbookIndexKey(onlyB.workbookOwner()))
    expect(onlyB.publishWorkbookIndex(onlyA.workbookOwner())).toBe(false)
  })

  it('rejects a stale whole-workbook index publication (plan race assertion)', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }, { sheetId: 'B', rows: column([1]) }])
    const session = new DocumentSession('doc', createObjectWorkbookSource(doc, 'doc'))
    const aView = draftOf(session, 'A'), bView = draftOf(session, 'B')
    const before = session.globalViewGen
    await session.commitView('A', aView) // A local revision becomes 1
    const staleA = session.beginWorkbookIndex()
    await session.commitView('B', bView) // B local revision also becomes 1
    expect(session.view('A').revision).toBe(1)
    expect(session.view('B').revision).toBe(1)
    expect(session.globalViewGen).toBeGreaterThan(before + 1)
    expect(session.publishWorkbookIndex(staleA)).toBe(false)
    expect(session.publishedWorkbookIndex).toBeNull()
    const fresh = session.beginWorkbookIndex()
    expect(fresh.sheetVector).toEqual([{ sheetId: 'A', revision: 1, valueGen: 0 }, { sheetId: 'B', revision: 1, valueGen: 0 }])
    expect(session.publishWorkbookIndex(fresh)).toBe(true)
    expect(session.publishedWorkbookIndex).toEqual(fresh)
  })

  it('never collides same-sheet ids across documents', async () => {
    expect(cacheKey({ docId: 'a:b', sheetId: 'c', valueGen: 0, fontGen: 0, viewGen: 0 }))
      .not.toBe(cacheKey({ docId: 'a', sheetId: 'b:c', valueGen: 0, fontGen: 0, viewGen: 0 }))
    const { parse } = countingParse()
    const kernel = new ViewerKernel({ parse }, quiet)
    const bytes = await workbookBytes()
    for (const docId of ['doc-a', 'doc-b']) {
      completed(await run(kernel, { kind: 'open', requestId: rid(), docId, source: { kind: 'blob', blob: new Blob([bytes]) }, profile: 'D' }))
    }
    const docA = kernel.session('doc-a')!, docB = kernel.session('doc-b')!
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner: initialIdentity('doc-a', '1'), cols: new Uint16Array([0]) }))
    const a = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc-a', '1'), spec: sortSpec(4, 1) }))
    expect(a.owner).toEqual({ docId: 'doc-a', sheetId: '1', valueGen: 0, fontGen: 0, viewGen: 1 })
    expect(docB.identity('1').viewGen).toBe(0)
    expect(docB.view('1').order.kind).toBe('identity')
    expect(docB.isPrepared('1', 0)).toBe(false)
    expect(docA.isPrepared('1', 0)).toBe(true)
    expect(cacheKey(docA.identity('1'))).not.toBe(cacheKey(docB.identity('1')))
    expect(cacheKey({ ...a.owner, docId: 'doc-b' })).not.toBe(cacheKey(a.owner))
    // doc-a's generation never authorizes a command on doc-b's same sheet id.
    const stale = await run(kernel, { kind: 'sort', requestId: rid(), owner: { ...a.owner, docId: 'doc-b' }, spec: null })
    expect(stale).toMatchObject({ kind: 'error', code: 'stale-source', retryable: true })
    expect(docB.globalViewGen).toBe(0)
  })

  it('populates commandKind on every reply kind, including refusals and cancellation', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: [[2, null, null, null, 'x'], [1]], merges: [{ minRow: 0, minCol: 3, maxRow: 1, maxCol: 3 }] }])
    let release: () => void = () => {}
    const filterEvaluator: FilterEvaluator = ({ signal }) => new Promise(resolve => {
      release = () => resolve({ kind: 'complete', matches: () => true })
      signal.addEventListener('abort', () => resolve({ kind: 'cancelled' }))
    })
    const kernel = new ViewerKernel({ parse: async () => doc }, { ...quiet, filterEvaluator })
    const log: WorkerReply[] = []
    const kinds = new Map<string, string>()
    const send = <K extends WorkerCommand['kind']>(command: Extract<WorkerCommand, { kind: K }>) => {
      kinds.set(command.requestId, command.kind)
      return run(kernel, command, log)
    }
    const owner = initialIdentity('doc', 'A')
    const filter: FilterSpec = { range: { r0: 0, r1: 1, c0: 0, c1: 4, headerRows: 0, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
    await send({ kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
    await send({ kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
    await send({ kind: 'activateSheet', requestId: rid(), owner: initialWorkbookOwner('doc', [{ sheetId: 'A' }]), sheetId: 'A' })
    await send({ kind: 'activateSheet', requestId: rid(), owner: initialWorkbookOwner('doc', [{ sheetId: 'A' }]), sheetId: 'nope' })
    await send({ kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) })
    const refused = await send({ kind: 'sort', requestId: rid(), owner, spec: sortSpec(1, 4) })
    expect(refused).toMatchObject({ kind: 'error', code: 'operation-refused', retryable: false, refusal: { reason: 'merge' } })
    await send({ kind: 'viewport', requestId: rid(), owner, window: { box: { x: 0, y: 0, width: 10, height: 10 }, zoom: 1, deviceScale: 1 } })
    await send({ kind: 'search', requestId: rid(), owner: initialWorkbookOwner('doc', [{ sheetId: 'A' }]), query: { text: 'x', caseSensitive: false } })
    await send({ kind: 'valueQuery', requestId: rid(), owner, query: { col: 0, rows: sourceBitset(0, 1, [0, 1]), limit: 4 } })
    await send({ kind: 'valueQuery', requestId: rid(), owner: { ...owner, viewGen: 9 }, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1 }] }, limit: 4 } })
    // A blocked filter is cancelled while in flight: both replies keep their own kinds.
    const filterId = rid()
    kinds.set(filterId, 'filter')
    const pending = run(kernel, { kind: 'filter', requestId: filterId, owner, spec: filter, nowMs: 0 }, log)
    await Promise.resolve()
    const ack = await send({ kind: 'cancel', requestId: rid(), targetRequestId: filterId })
    expect(ack).toMatchObject({ kind: 'cancelled', commandKind: 'cancel', owner })
    expect(await pending).toMatchObject({ kind: 'cancelled', commandKind: 'filter' })
    release()
    await send({ kind: 'cancel', requestId: rid(), targetRequestId: 'missing' })
    await send({ kind: 'close', requestId: rid(), docId: 'doc' })
    await send({ kind: 'sort', requestId: rid(), owner, spec: null })
    const seen = new Set(log.map(out => `${out.commandKind}:${out.kind}`))
    for (const kind of ['open', 'activateSheet', 'prepareColumns', 'sort', 'filter', 'viewport', 'valueQuery', 'search', 'cancel', 'close']) {
      expect([...seen].some(entry => entry.startsWith(`${kind}:`))).toBe(true)
    }
    for (const out of log) {
      expect(isViewerReply(out)).toBe(true)
      expect(out.commandKind).toBe(kinds.get(out.requestId))
      expect(out).toHaveProperty('owner')
      expect(Array.isArray(out.diagnostics)).toBe(true)
    }
  })

  it('tags every bitset with its space and round-trips the display mask', async () => {
    // Row 2 is authored hidden: it stays pinned and hidden through the sort.
    const doc = objectDoc([{ sheetId: 'A', rows: column([5, 4, 3, 2, 1]), hidden: [2] }])
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    const log: WorkerReply[] = []
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }, log))
    const out = await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc', 'A'), spec: sortSpec(4, 0) }, log)
    const delta = deltaOf(out)
    expect([...delta.bands[0].sourceAtOffset]).toEqual([4, 3, 2, 1, 0])
    expect(delta.visibility).toHaveLength(1)
    expect(delta.visibility[0]).toMatchObject({ kind: 'bitset', space: 'display', r0: 0, r1: 2 })
    const page = completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner: delta.owner, query: { col: 0, rows: sourceBitset(0, 4, [0, 4]), limit: 8 } }, log))
    expect(page.kind === 'valueQuery' && page.page.items.filter(item => item.kind === 'cell').map(item => item.kind === 'cell' && item.cell.value)).toEqual([5, 1])
    const displayRows = { ...sourceBitset(0, 4, [0]), space: 'display' } as unknown as ReturnType<typeof sourceBitset>
    const wrong = completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner: delta.owner, query: { col: 0, rows: displayRows, limit: 8 } }, log))
    expect(wrong.kind === 'valueQuery' && wrong.page.error?.code).toBe('invalid-source-rect')
    const bitsets = bitsetsOf(log)
    expect(bitsets.length).toBeGreaterThan(0)
    for (const bitset of bitsets) {
      expect(['source', 'display']).toContain(bitset.space)
      validateBitset(bitset, bitset.space)
    }
    const resident = kernel.session('doc')!.view('A').visibility
    const decoded = visibilityFromDisplayBitsets(delta.visibility)
    for (let slot = 0; slot < 70; slot++) expect(decoded.visible(slot)).toBe(resident.visible(slot))
    expect(decoded.rank(decoded.logicalCount)).toBe(delta.counts.visible)
    expect(delta.counts).toEqual({ visible: 1_048_575, total: 1_048_576 })
  })

  it('records actual copies and transfers and never detaches resident buffers', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([3, 1, 2, 5, 4]), hidden: [3] }])
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    const port = connectInProcess(kernel)
    const client = createViewerClient(port, 'c')
    const opened = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
    expect(opened.kind).toBe('complete')
    expect(kernel.ledger()).toEqual({ copies: 0, copiedBytes: 0, transfers: 0, transferredBytes: 0, estimatedCopiedBytes: 0 })
    const owner = initialIdentity('doc', 'A')
    const cols = new Uint16Array([0])
    const prepared = await client.send({ kind: 'prepareColumns', requestId: client.nextRequestId(), owner, cols })
    expect(cols.byteLength).toBe(0) // ownership moved to the worker
    // One estimated copy: the object source document structured-cloned at open.
    expect(client.ledger()).toMatchObject({ copies: 1, transfers: 1, transferredBytes: 2 })
    expect(client.ledger().copiedBytes).toBe(client.ledger().estimatedCopiedBytes)
    expect(completed(prepared)).toMatchObject({ kind: 'prepareColumns' })
    expect(kernel.ledger()).toEqual({ copies: 1, copiedBytes: 2, transfers: 1, transferredBytes: 2, estimatedCopiedBytes: 0 })
    const before = kernel.ledger()
    const sorted = await client.send({ kind: 'sort', requestId: client.nextRequestId(), owner, spec: sortSpec(4, 0) })
    const delta = deltaOf(sorted)
    const arrays = [...delta.bands.flatMap(band => [band.sourceAtOffset, band.displayOffsetOfSource]), ...delta.visibility.map(bitset => bitset.words)]
    const bytes = arrays.reduce((sum, array) => sum + array.byteLength, 0)
    expect(arrays.length).toBe(3)
    expect(kernel.ledger()).toEqual({
      copies: before.copies + arrays.length, copiedBytes: before.copiedBytes + bytes,
      transfers: before.transfers + arrays.length, transferredBytes: before.transferredBytes + bytes, estimatedCopiedBytes: 0,
    })
    const resident = kernel.session('doc')!.view('A').order.bands[0]
    expect(resident.sourceAtOffset.byteLength).toBe(20)
    expect([...delta.bands[0].sourceAtOffset]).toEqual([...resident.sourceAtOffset])
    // The resident order/mask stay usable for the next replacement sort.
    const again = deltaOf(await client.send({ kind: 'sort', requestId: client.nextRequestId(), owner: delta.owner, spec: sortSpec(4, 0, 0, 'desc') }))
    expect([...again.bands[0].sourceAtOffset]).toEqual([4, 0, 2, 3, 1]) // hidden row 3 stays pinned
    expect(transferablesOf(again)).toHaveLength(3)
    await port.idle()
    expect(client.pendingCount).toBe(0)
    client.dispose()
  })

  it('copies blob source bytes into the worker exactly once', async () => {
    const { parse } = countingParse()
    const kernel = new ViewerKernel({ parse }, quiet)
    const blob = new Blob([await workbookBytes()])
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'blob', blob }, profile: 'D' }))
    expect(kernel.ledger()).toEqual({ copies: 1, copiedBytes: blob.size, transfers: 0, transferredBytes: 0, estimatedCopiedBytes: 0 })
    deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc', '1'), spec: sortSpec(4, 1) }))
    // Plus the two 5-row Uint32 band copies (2 x 20B); no hidden rows, so no bitset.
    expect(kernel.ledger()).toEqual({ copies: 3, copiedBytes: blob.size + 40, transfers: 2, transferredBytes: 40, estimatedCopiedBytes: 0 })
  })

  it('refuses guarded operations as operation-refused errors with the owner unchanged', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([2, 1, 3]), merges: [{ minRow: 1, minCol: 6, maxRow: 2, maxCol: 7 }] }])
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const owner = initialIdentity('doc', 'A')
    const session = kernel.session('doc')!
    const merge = await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: sortSpec(2, 0) })
    expect(merge).toMatchObject({
      kind: 'error', commandKind: 'sort', code: 'operation-refused', retryable: false, owner,
      refusal: { reason: 'merge', range: { r0: 1, r1: 2, c0: 6, c1: 7 } },
    })
    expect(merge.diagnostics).toHaveLength(1)
    expect(merge).not.toHaveProperty('result')
    const filter: FilterSpec = { range: { r0: 0, r1: 2, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 1 }]]) }
    const unsupported = await run(kernel, { kind: 'filter', requestId: rid(), owner, spec: filter, nowMs: 0 })
    expect(unsupported).toMatchObject({ kind: 'error', code: 'operation-refused', retryable: false, refusal: { reason: 'unsupported' } })
    expect(session.globalViewGen).toBe(0)
    expect(session.identity('A')).toEqual(owner)
    expect(session.view('A').order.kind).toBe('identity')
  })

  it('rejects stale owners and closed documents without publishing', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([2, 1]) }])
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const first = deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc', 'A'), spec: sortSpec(1, 0) }))
    const stale = await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc', 'A'), spec: null })
    expect(stale).toMatchObject({ kind: 'error', commandKind: 'sort', code: 'stale-source', retryable: true })
    const session = kernel.session('doc')!
    expect(session.globalViewGen).toBe(1)
    expect(completed(await run(kernel, { kind: 'close', requestId: rid(), docId: 'doc' }))).toEqual({ kind: 'close' })
    expect(kernel.session('doc')).toBeUndefined()
    expect(session.isClosed).toBe(true)
    expect(await run(kernel, { kind: 'sort', requestId: rid(), owner: first.owner, spec: null })).toMatchObject({ kind: 'error', code: 'stale-source' })
    expect(await run(kernel, { kind: 'close', requestId: rid(), docId: 'doc' })).toMatchObject({ kind: 'error', commandKind: 'close', code: 'stale-source' })
  })

  it('routes viewer commands to the resident kernel and legacy requests to the one-shot path', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }])
    const posted: unknown[] = []
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    const handler = createWorkerMessageHandler(message => { posted.push(message) }, kernel, {
      createCanvas: async () => { throw new Error('unused') }, toBitmap: async () => { throw new Error('unused') },
      parse: async () => ({}), paintables: async () => { throw new Error('unused') },
    } as never)
    handler({ data: { kind: 'open', requestId: 'r1', docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' } })
    handler({ data: { id: 5, kind: 'nope', format: 'docx', source: new ArrayBuffer(0), index: 0 } })
    await vi.waitFor(() => expect(posted).toHaveLength(2))
    expect(posted).toContainEqual(expect.objectContaining({ commandKind: 'open', requestId: 'r1', kind: 'complete' }))
    expect(posted).toContainEqual(expect.objectContaining({ id: 5, ok: false }))
    expect(kernel.session('doc')).toBeDefined()
  })
})

function draftOf(session: DocumentSession, sheetId: string): ViewDraft {
  const view = session.view(sheetId)
  return { order: view.order, visibility: view.visibility, sort: null, filter: null }
}

function deferred<T = void>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>(done => { resolve = done })
  return { promise, resolve }
}
const tick = () => new Promise(resolve => setTimeout(resolve, 0))

describe('S4.1 review findings (codex)', () => {
  it('#1 refuses a sort over generated dynamic-spill children whose evaluated extent is not resolved', async () => {
    // SEQUENCE(3,1) at A1 with saved footprint A1; A2:A3 are engine-generated children.
    const doc = objectDoc([{ sheetId: 'A', rows: column([1, 2, 3]) }])
    Object.assign(doc.sheets[0].rows[0].cells[0], { formula: '_xlfn.SEQUENCE(3,1)', dynamicArray: true, arrayRef: 'A1' })
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const spec: SortSpec = { range: { r0: 1, r1: 2, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'desc' }], caseSensitive: false }
    const out = await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc', 'A'), spec })
    // Incomplete coverage is a wait (spec §6.3/§10.4), not a refusal.
    expect(out).toMatchObject({ kind: 'incomplete', commandKind: 'sort', missing: [{ r0: 0, r1: 1_048_575, c0: 0, c1: 16_383 }] })
    expect(kernel.session('doc')!.globalViewGen).toBe(0)
  })

  it('#2 copies resident selection buffers into the delta and accounts them', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([3, 1, 2]) }])
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    const port = connectInProcess(kernel)
    const client = createViewerClient(port, 'sel')
    completed(await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const session = kernel.session('doc')!
    const selection = { active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 }, ranges: [{ rows: { kind: 'list' as const, r0: 0, r1: 2, rows: new Uint32Array([0, 2]) }, cols: [{ c0: 0, c1: 0 }] }] }
    const committed = await session.commitView('A', { ...draftOf(session, 'A'), selection })
    const before = kernel.ledger()
    const first = deltaOf(await client.send({ kind: 'sort', requestId: client.nextRequestId(), owner: committed.identity, spec: sortSpec(2, 0) }))
    const resident = session.view('A').selection.ranges[0].rows
    expect(resident.kind === 'list' && resident.rows.length).toBe(2)
    const sent = first.selection.ranges[0].rows
    expect(sent.kind === 'list' && [...sent.rows]).toEqual([0, 2])
    // 2 band arrays (12B each) + 8B selection rows; no hidden rows, so no bitset.
    expect(kernel.ledger().copies - before.copies).toBe(3)
    expect(kernel.ledger().copiedBytes - before.copiedBytes).toBe(32)
    const second = deltaOf(await client.send({ kind: 'sort', requestId: client.nextRequestId(), owner: first.owner, spec: sortSpec(2, 0, 0, 'desc') }))
    const again = second.selection.ranges[0].rows
    expect(again.kind === 'list' && [...again.rows]).toEqual([0, 2])
    client.dispose()
  })

  it('#3 a query that finishes after a newer commit replies stale, never complete for the old owner', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: [[2, 'b'], [1, 'a']] }])
    const tracked = instrumented(createObjectWorkbookSource(doc, 'doc'))
    const gate = deferred()
    const inner = tracked.wrapped.store('A')
    const delayed = inner.queryColumn
    inner.queryColumn = async (request, options) => {
      if (request.col === 1) await gate.promise
      return delayed(request, options)
    }
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => tracked.wrapped }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const owner = initialIdentity('doc', 'A')
    const query = run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 1, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1 }] }, limit: 4 } })
    await tick()
    expect(deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: sortSpec(1, 1) })).owner.viewGen).toBe(1)
    gate.resolve()
    expect(await query).toMatchObject({ kind: 'error', commandKind: 'valueQuery', code: 'stale-source', owner })
  })

  it('#4 the client never resolves a request from a reply with another command kind or owner', async () => {
    const posted: { requestId: string }[] = []
    let deliver: (event: { data: unknown }) => void = () => {}
    const port = {
      postMessage: (message: unknown) => { posted.push(message as { requestId: string }) },
      addEventListener: (_type: 'message', listener: (event: { data: unknown }) => void) => { deliver = listener },
      removeEventListener: () => {},
    }
    const client = createViewerClient(port)
    const owner = initialIdentity('doc-a', 'A')
    let settled: unknown = 'pending'
    const pending = client.send({ kind: 'prepareColumns', requestId: 'x', owner, cols: new Uint16Array([0]) })
    void pending.then(out => { settled = out })
    deliver({ data: { kind: 'complete', commandKind: 'close', requestId: 'x', owner: { docId: 'other' }, diagnostics: [], result: { kind: 'close' } } })
    deliver({ data: { kind: 'complete', commandKind: 'prepareColumns', requestId: 'x', owner: { ...owner, viewGen: 7 }, diagnostics: [], result: { kind: 'prepareColumns', cols: new Uint16Array([0]) } } })
    await tick()
    expect(settled).toBe('pending')
    expect(client.rejectedReplies).toBe(2)
    deliver({ data: { kind: 'complete', commandKind: 'prepareColumns', requestId: 'x', owner, diagnostics: [], result: { kind: 'prepareColumns', cols: new Uint16Array([0]) } } })
    expect(await pending).toMatchObject({ kind: 'complete', commandKind: 'prepareColumns' })
  })

  it('#5 close during a pending open abandons the open instead of installing a session', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }])
    const gate = deferred()
    let released = false
    const source = createObjectWorkbookSource(doc, 'doc')
    const kernel = new ViewerKernel({
      parse: async () => doc,
      openSource: async () => { await gate.promise; return { ...source, sheets: source.sheets, store: id => source.store(id), geometry: id => source.geometry(id), release: async () => { released = true } } },
    }, quiet)
    const opening = run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
    await tick()
    expect(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })).toMatchObject({ kind: 'error' })
    expect(await run(kernel, { kind: 'close', requestId: rid(), docId: 'doc' })).toMatchObject({ kind: 'complete', commandKind: 'close' })
    gate.resolve()
    expect(await opening).toMatchObject({ kind: 'cancelled', commandKind: 'open' })
    expect(kernel.session('doc')).toBeUndefined()
    expect(released).toBe(true)
  })

  it('#6 an unavailable table inventory blocks guarded operations instead of reporting complete', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([2, 1]) }])
    delete (doc as { tables?: unknown }).tables
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const out = await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc', 'A'), spec: sortSpec(1, 0) })
    expect(out).toMatchObject({ kind: 'incomplete', commandKind: 'sort' })
    const structural = await createObjectWorkbookSource(doc, 'doc').store('A').queryStructural({ r0: 0, r1: 1, c0: 0, c1: 0 }, { signal: new AbortController().signal, stamp: { docId: 'doc', sheetId: 'A', valueGen: 0 } })
    expect(structural.completeness).toBe('incomplete')
  })

  it('#7 prepared columns are admitted against the profile working set', async () => {
    const row = Array.from({ length: 10 }, (_v, col) => col)
    const doc = objectDoc([{ sheetId: 'A', rows: [row, row] }])
    const cols = (...list: number[]) => new Uint16Array(list)
    const open = async (profile: 'D' | 'M', options = {}) => {
      const kernel = new ViewerKernel({ parse: async () => doc }, { ...quiet, ...options })
      completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile }))
      return kernel
    }
    const owner = initialIdentity('doc', 'A')
    const mobile = await open('M')
    expect(await run(mobile, { kind: 'prepareColumns', requestId: rid(), owner, cols: cols(0, 1, 2, 3, 4, 5, 6, 7, 8) })).toMatchObject({ kind: 'error', code: 'resource-limit', retryable: false })
    expect(await run(mobile, { kind: 'prepareColumns', requestId: rid(), owner, cols: cols(0, 1, 2, 3, 4) })).toMatchObject({ kind: 'quota', resource: 'memory' })
    expect(mobile.session('doc')!.preparedStats().columns).toBe(0)
    completed(await run(mobile, { kind: 'prepareColumns', requestId: rid(), owner, cols: cols(0, 1, 2, 3) }))
    completed(await run(mobile, { kind: 'prepareColumns', requestId: rid(), owner, cols: cols(4, 5, 6, 7) }))
    expect(mobile.session('doc')!.preparedStats().columns).toBe(4)
    expect(mobile.session('doc')!.isPrepared('A', 7)).toBe(true)
    expect(mobile.session('doc')!.isPrepared('A', 0)).toBe(false)
    const desktop = await open('D')
    completed(await run(desktop, { kind: 'prepareColumns', requestId: rid(), owner, cols: cols(0, 1, 2, 3, 4, 5, 6, 7) }))
    expect(desktop.session('doc')!.preparedStats().columns).toBe(8)
    const tiny = await open('D', { preparedBudget: { maxColumns: 8, maxBytes: 64 } })
    expect(await run(tiny, { kind: 'prepareColumns', requestId: rid(), owner, cols: cols(0) })).toMatchObject({ kind: 'quota', resource: 'memory', limitBytes: 64 })
    expect(tiny.session('doc')!.preparedStats()).toEqual({ columns: 0, bytes: 0 })
  })

  it('#8 the client accounts whole cloned backing buffers and the cloned source document', async () => {
    const port = { postMessage: () => {}, addEventListener: () => {}, removeEventListener: () => {} }
    const client = createViewerClient(port)
    const big = new ArrayBuffer(1 << 20)
    client.send({ kind: 'valueQuery', requestId: 'q', owner: initialIdentity('doc', 'A'), query: { col: 0, rows: { kind: 'list', r0: 0, r1: 1, rows: new Uint32Array(big, 0, 1) }, limit: 1 } }).catch(() => {})
    expect(client.ledger()).toMatchObject({ copies: 1, copiedBytes: 1 << 20, transfers: 0 })
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }])
    client.send({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }).catch(() => {})
    const ledger = client.ledger()
    expect(ledger.copies).toBe(2)
    expect(ledger.estimatedCopiedBytes).toBeGreaterThan(0)
    expect(ledger.copiedBytes).toBe((1 << 20) + ledger.estimatedCopiedBytes)
    client.dispose()
  })

  it('#9 column cursors stay valid at the last row and across preparation', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: [] }])
    doc.sheets[0].rows = [{ index: 1_048_575, cells: [{ ref: 'A1048576', row: 1_048_575, col: 0, value: 9, styleIndex: 0, hasCachedValue: true }] }]
    const store = createObjectWorkbookSource(doc, 'doc').store('A')
    const options = { signal: new AbortController().signal, stamp: store.stamp }
    const edge = await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 1_048_574, r1: 1_048_575 }] }, limit: 2 }, options)
    expect(edge.error).toBeUndefined()
    expect(edge.items.map(item => item.kind)).toEqual(['blank', 'cell'])
    expect(edge.nextCursor).toBeUndefined()
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const owner = initialIdentity('doc', 'A')
    const query = (cursor?: string) => ({ col: 0, rows: { kind: 'intervals' as const, intervals: [{ r0: 1_048_570, r1: 1_048_575 }] }, limit: 1, cursor })
    const first = completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: query() }))
    const cursor = first.kind === 'valueQuery' ? first.page.nextCursor : undefined
    expect(cursor).toBeDefined()
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) }))
    const second = completed(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: query(cursor) }))
    expect(second.kind === 'valueQuery' && second.page.error).toBeUndefined()
    expect(second.kind === 'valueQuery' && second.page.items).toMatchObject([{ kind: 'cell', row: 1_048_575 }])
  })

  it('#10 malformed source bitsets are rejected before iteration', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }])
    const store = createObjectWorkbookSource(doc, 'doc').store('A')
    const options = { signal: new AbortController().signal, stamp: store.stamp }
    const beyond = { kind: 'bitset' as const, space: 'source' as const, r0: 1_048_545, r1: 1_048_576, words: new Uint32Array(1), count: 0 }
    expect((await store.queryColumn({ col: 0, rows: beyond, limit: 4 }, options)).error?.code).toBe('invalid-source-rect')
    const badCount = { kind: 'bitset' as const, space: 'source' as const, r0: 0, r1: 3, words: new Uint32Array([1]), count: 2 }
    expect((await store.queryColumn({ col: 0, rows: badCount, limit: 4 }, options)).error?.code).toBe('invalid-source-rect')
    const badList = { kind: 'list' as const, r0: 0, r1: 1_048_576, rows: new Uint32Array([0]) }
    expect((await store.queryColumn({ col: 0, rows: badList, limit: 4 }, options)).error?.code).toBe('invalid-source-rect')
  })
})

/** Store whose column cursors are opaque tokens (a conforming §4.1 backend). */
function opaqueCursors(source: WorkbookSource): WorkbookSource {
  const stores = new Map<string, SourceStore>()
  let counter = 0
  return {
    valuePolicy: source.valuePolicy, sheets: source.sheets, geometry: id => source.geometry(id), release: () => source.release(),
    store(sheetId) {
      let store = stores.get(sheetId)
      if (!store) {
        const inner = source.store(sheetId)
        const tokens = new Map<string, string>()
        store = Object.create(inner) as SourceStore
        store.queryColumn = async (request, options) => {
          let cursor: string | undefined
          if (request.cursor !== undefined) {
            cursor = tokens.get(request.cursor)
            if (cursor === undefined) return { items: [], completeness: 'incomplete', stamp: inner.stamp, error: { code: 'invalid-cursor', retryable: false, message: 'unknown token' } }
          }
          const page = await inner.queryColumn({ ...request, cursor }, options)
          if (page.nextCursor === undefined) return page
          const token = `opaque-${++counter}`
          tokens.set(token, page.nextCursor)
          return { ...page, nextCursor: token }
        }
        stores.set(sheetId, store)
      }
      return store
    },
  }
}
const cellValues = (out: WorkerReply) => {
  const result = completed(out)
  if (result.kind !== 'valueQuery') throw new Error(result.kind)
  expect(result.page.error).toBeUndefined()
  return result.page.items.map(item => item.kind === 'cell' ? item.cell.value : item.kind)
}

describe('S4.1 round-2 findings (codex C / pi P)', () => {
  it('P1 structural incomplete coverage replies incomplete with missing domains and an unresolvable diagnostic', async () => {
    const drawing = objectDoc([{ sheetId: 'A', rows: column([2, 1]) }])
    drawing.sheets[0].drawings = [{} as never]
    const sequence = objectDoc([{ sheetId: 'A', rows: column([1, 2, 3]) }])
    Object.assign(sequence.sheets[0].rows[0].cells[0], { formula: '_xlfn.SEQUENCE(3,1)', dynamicArray: true, arrayRef: 'A1' })
    for (const [doc, missing] of [[drawing, { r0: 0, r1: 1_048_575, c0: 0, c1: 16_383 }], [sequence, { r0: 0, r1: 1_048_575, c0: 0, c1: 16_383 }]] as const) {
      const evaluator: FilterEvaluator = async () => ({ kind: 'complete', matches: () => true })
      const kernel = new ViewerKernel({ parse: async () => doc }, { ...quiet, filterEvaluator: evaluator })
      completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
      const owner = initialIdentity('doc', 'A')
      const sorted = await run(kernel, { kind: 'sort', requestId: rid(), owner, spec: sortSpec(1, 0) })
      expect(sorted).toMatchObject({ kind: 'incomplete', commandKind: 'sort', owner })
      expect(sorted.kind === 'incomplete' && sorted.missing).toContainEqual(missing)
      expect(sorted.diagnostics.map(d => d.code)).toContain('coverage-unresolvable')
      const filter: FilterSpec = { range: { r0: 0, r1: 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
      const filtered = await run(kernel, { kind: 'filter', requestId: rid(), owner, spec: filter, nowMs: 0 })
      expect(filtered).toMatchObject({ kind: 'incomplete', commandKind: 'filter' })
      expect(filtered.diagnostics.map(d => d.code)).toContain('coverage-unresolvable')
      expect(kernel.session('doc')!.globalViewGen).toBe(0)
    }
  })

  it('C1 concurrent preparations are serialized: no double charge, every completed batch was resident', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: [[1, 2, 3, 4], [5, 6, 7, 8]] }])
    const tracked = instrumented(createObjectWorkbookSource(doc, 'doc'))
    const gate = deferred()
    const inner = tracked.wrapped.store('A')
    const base = inner.queryColumn
    inner.queryColumn = async (request, options) => { if (request.col === 0) await gate.promise; return base(request, options) }
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => tracked.wrapped }, { ...quiet, preparedBudget: { maxColumns: 2, maxBytes: 1 << 20 } })
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const owner = initialIdentity('doc', 'A')
    const session = kernel.session('doc')!
    const same = [run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) }), run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) })]
    await tick()
    gate.resolve()
    for (const out of await Promise.all(same)) completed(out)
    expect(session.preparedStats()).toEqual({ columns: 1, bytes: session.preparedFootprint('A', 0)!.bytes })
    const gate2 = deferred()
    inner.queryColumn = async (request, options) => { if (request.col === 1) await gate2.promise; return base(request, options) }
    const first = run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0, 1]) })
    const second = run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([2, 3]) })
    await tick()
    gate2.resolve()
    completed(await first)
    completed(await second)
    // Commit order follows request order; the later batch wins the working set.
    expect(session.preparedColumns('A')).toEqual([2, 3])
    expect(session.preparedStats().bytes).toBe(session.preparedFootprint('A', 2)!.bytes + session.preparedFootprint('A', 3)!.bytes)
  })

  it('C2 the source keeps no whole-column caches behind the prepared working set', async () => {
    const row = Array.from({ length: 8 }, (_v, col) => col)
    const doc = objectDoc([{ sheetId: 'A', rows: [row, row] }])
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'M' }))
    const owner = initialIdentity('doc', 'A')
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0, 1, 2, 3]) }))
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([4, 5, 6, 7]) }))
    const session = kernel.session('doc')!
    expect(session.preparedStats().columns).toBe(4)
    const store = session.source.store('A') as unknown as { cachedColumnCount: number }
    expect(store.cachedColumnCount).toBe(0)
  })

  it('C3 a completion is rechecked against close immediately before emission', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }])
    const order: string[] = []
    let closing: Promise<void> | undefined
    const kernel = new ViewerKernel({
      parse: async () => doc,
      afterDispatch: async command => {
        if (command.kind === 'valueQuery') closing = kernel.handle({ kind: 'close', requestId: 'close-1', docId: 'doc' }, out => order.push(`close:${out.kind}`))
      },
    }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    await kernel.handle({ kind: 'valueQuery', requestId: 'q', owner: initialIdentity('doc', 'A'), query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, limit: 1 } }, out => order.push(`query:${out.kind}`))
    await closing
    expect(order).not.toContain('query:complete')
    expect(order).toContain('close:complete')
  })

  it('C4 a cancel racing a finished target settles with the target owner instead of hanging', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }])
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    const queue: { data: unknown }[] = []
    let listener: (event: { data: unknown }) => void = () => {}
    const tasks: Promise<void>[] = []
    const port = {
      postMessage: (message: unknown) => { tasks.push(kernel.handle(structuredClone(message) as WorkerCommand, out => queue.push({ data: structuredClone(out) }))) },
      addEventListener: (_type: 'message', l: (event: { data: unknown }) => void) => { listener = l },
      removeEventListener: () => {},
    }
    const flush = async () => { await Promise.all(tasks); while (queue.length) listener(queue.shift()!) }
    const client = createViewerClient(port)
    const opened = client.send({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
    await flush()
    completed(await opened)
    const owner = initialIdentity('doc', 'A')
    const query = client.send({ kind: 'valueQuery', requestId: 'q', owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, limit: 1 } })
    await Promise.all(tasks) // worker finished; its reply is still queued
    const cancel = client.send({ kind: 'cancel', requestId: 'c', targetRequestId: 'q' })
    await flush()
    expect(await cancel).toMatchObject({ commandKind: 'cancel', kind: 'error', owner })
    expect((await query).kind).toBe('complete')
    expect(client.rejectedReplies).toBe(0)
    // Cancelling something the worker never saw still settles (document-less owner).
    const unknown = client.send({ kind: 'cancel', requestId: 'c2', targetRequestId: 'never' })
    await flush()
    expect(await unknown).toMatchObject({ commandKind: 'cancel', kind: 'error', owner: { docId: '' } })
  })

  it('C5 paging survives preparation and eviction over a store with opaque cursors', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: [[1, 'x'], [2, 'y'], [3, 'z']] }])
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => opaqueCursors(createObjectWorkbookSource(doc, 'doc')) }, { ...quiet, preparedBudget: { maxColumns: 1, maxBytes: 1 << 20 } })
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const owner = initialIdentity('doc', 'A')
    const query = (cursor?: string) => ({ col: 0, rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 2 }] }, limit: 1, cursor })
    const page = async (cursor?: string) => {
      const out = await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: query(cursor) })
      const result = completed(out)
      return result.kind === 'valueQuery' ? result.page : undefined
    }
    const p1 = (await page())!
    expect(p1.error).toBeUndefined()
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) }))
    const p2 = (await page(p1.nextCursor))!
    expect(p2.error).toBeUndefined()
    // Evict column 0 (budget 1) and continue from a prepared-view cursor via the store.
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([1]) }))
    const p3 = (await page(p2.nextCursor))!
    expect(p3.error).toBeUndefined()
    expect([...p1.items, ...p2.items, ...p3.items].map(item => item.kind === 'cell' ? item.cell.value : item.kind)).toEqual([1, 2, 3])
    expect(p3.nextCursor).toBeUndefined()
  })

  it('C6 empty-column overhead is reserved before staging', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }])
    const kernel = new ViewerKernel({ parse: async () => doc }, { ...quiet, preparedBudget: { maxColumns: 8, maxBytes: Math.floor(PREPARED_COLUMN_OVERHEAD * 2.5) } })
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
    const out = await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner: initialIdentity('doc', 'A'), cols: new Uint16Array([5, 6, 7]) })
    expect(out).toMatchObject({ kind: 'quota', resource: 'memory' })
    expect(kernel.session('doc')!.preparedStats()).toEqual({ columns: 0, bytes: 0 })
  })

  it('C7 a failed session construction releases the acquired source', async () => {
    let releases = 0
    const kernel = new ViewerKernel({
      parse: async () => objectDoc([]),
      openSource: async () => ({ valuePolicy: 'engine', sheets: [], store: () => { throw new Error('none') }, geometry: () => { throw new Error('none') }, release: async () => { releases++ } }),
    }, quiet)
    const out = await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: objectDoc([]) }, profile: 'D' })
    expect(out.kind).toBe('error')
    expect(releases).toBe(1)
    expect(kernel.session('doc')).toBeUndefined()
  })

  it('P2 prepared columns retain compact lanes whose charged bytes equal measured retention', async () => {
    const n = 5000
    const numbers = objectDoc([{ sheetId: 'A', rows: Array.from({ length: n }, (_v, row) => [row * 1.5, `s${row % 10}`]) }])
    const kernel = new ViewerKernel({ parse: async () => numbers }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: numbers }, profile: 'D' }))
    const owner = initialIdentity('doc', 'A')
    completed(await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0, 1]) }))
    const session = kernel.session('doc')!
    const numeric = session.preparedFootprint('A', 0)!
    expect(numeric).toMatchObject({ entries: n, extraObjects: 0, stringBytes: 0 })
    expect(numeric.laneBytes).toBe(n * PREPARED_ENTRY_LANE_BYTES)
    expect(numeric.bytes).toBe(PREPARED_COLUMN_OVERHEAD + numeric.laneBytes)
    const text = session.preparedFootprint('A', 1)!
    expect(text.extraObjects).toBe(0)
    expect(text.uniqueStrings).toBe(10)
    expect(text.bytes).toBe(PREPARED_COLUMN_OVERHEAD + text.laneBytes + text.stringBytes)
    expect(session.preparedStats().bytes).toBe(numeric.bytes + text.bytes)
    // Order of magnitude: compact lanes, not per-row payload objects (~tens of bytes per row).
    expect(numeric.bytes / n).toBeLessThan(32)
    // Served pages are rebuilt from lanes with identical values.
    expect(cellValues(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 2, r1: 3 }] }, limit: 4 } }))).toEqual([3, 4.5])
    expect(cellValues(await run(kernel, { kind: 'valueQuery', requestId: rid(), owner, query: { col: 1, rows: { kind: 'intervals', intervals: [{ r0: 11, r1: 11 }] }, limit: 4 } }))).toEqual(['s1'])
  })

  it('P3 readRows pages are bounded by cells as well as rows', async () => {
    expect(READ_ROWS_MAX_CELLS).toBe(262_144)
    const doc = objectDoc([{ sheetId: 'A', rows: Array.from({ length: 5 }, () => [1, 2, 3, 4]) }])
    const store = createObjectWorkbookSource(doc, 'doc', { readRowsMaxCells: 10 }).store('A')
    const options = { signal: new AbortController().signal, stamp: store.stamp }
    const sizes: number[] = []
    let cursor: string | undefined
    do {
      const page = await store.readRows({ r0: 0, r1: 1_048_575, c0: 0, c1: 16_383 }, cursor, options)
      expect(page.error).toBeUndefined()
      sizes.push(page.items.reduce((sum, row) => sum + row.cells.length, 0))
      cursor = page.nextCursor
    } while (cursor !== undefined)
    expect(sizes).toEqual([8, 8, 4])
  })

  it('P4 the populated guard feed honors the caller page limit', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: [Array.from({ length: 10 }, (_v, col) => col + 1)] }])
    const session = new DocumentSession('doc', createObjectWorkbookSource(doc, 'doc'))
    const feed = session.populatedSource('A', new AbortController().signal)
    const range = { r0: 0, r1: 1_048_575, c0: 0, c1: 16_383 }
    const pages: number[] = []
    let cursor: string | undefined
    do {
      const page = await feed.iterate(range, cursor, 3)
      pages.push(page.items.length)
      cursor = page.nextCursor
    } while (cursor !== undefined)
    expect(pages).toEqual([3, 3, 3, 1])
  })

  it('P5 an unknown command kind gets a typed unsupported error', async () => {
    const kernel = new ViewerKernel({ parse: async () => objectDoc([]) }, quiet)
    const replies: WorkerReply[] = []
    await kernel.handle({ kind: 'bogus', requestId: 'b' } as never, out => replies.push(out))
    expect(replies).toHaveLength(1)
    expect(replies[0]).toMatchObject({ kind: 'error', code: 'unsupported', requestId: 'b' })
  })

  it('P7 progress replies never settle a request (progress is a separate onProgress subscription since S4.2)', async () => {
    let listener: (event: { data: unknown }) => void = () => {}
    const client = createViewerClient({ postMessage: () => {}, addEventListener: (_t: 'message', l: (event: { data: unknown }) => void) => { listener = l }, removeEventListener: () => {} })
    const owner = initialIdentity('doc', 'A')
    let settled = false
    const pending = client.send({ kind: 'valueQuery', requestId: 'v', owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, limit: 1 } })
    void pending.then(() => { settled = true })
    listener({ data: { kind: 'progress', commandKind: 'valueQuery', requestId: 'v', owner, diagnostics: [], completed: 1 } })
    await tick()
    expect(settled).toBe(false)
    listener({ data: { kind: 'cancelled', commandKind: 'valueQuery', requestId: 'v', owner, diagnostics: [] } })
    expect((await pending).kind).toBe('cancelled')
    expect(client.send.length).toBe(1) // send(command) only: progress uses client.onProgress
  })

  it('P9 view snapshots are frozen copies', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }])
    const session = new DocumentSession('doc', createObjectWorkbookSource(doc, 'doc'))
    const view = session.view('A')
    expect(Object.isFrozen(view)).toBe(true)
    expect(Object.isFrozen(view.selection)).toBe(true)
    expect(() => { (view as { revision: number }).revision = 9 }).toThrow()
    expect(session.view('A').revision).toBe(0)
  })

  it('P10 prepareColumns validates its input before use and quota replies carry requiredBytes', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([1]) }])
    const kernel = new ViewerKernel({ parse: async () => doc }, quiet)
    completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'M' }))
    const owner = initialIdentity('doc', 'A')
    const session = kernel.session('doc')!
    expect(await session.prepareColumns(owner, null as never, new AbortController().signal)).toMatchObject({ kind: 'error', error: { code: 'invalid-coordinate' } })
    const bad = await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: [0] as never })
    expect(bad).toMatchObject({ kind: 'error', code: 'invalid-coordinate' })
    expect(bad.diagnostics.length).toBe(1)
    const quota = await run(kernel, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0, 1, 2, 3, 4]) })
    expect(quota).toMatchObject({ kind: 'quota', resource: 'memory' })
    expect(quota.kind === 'quota' && quota.requiredBytes).toBeGreaterThan(0)
  })
})

describe('S4.1 round-3 findings (opencode)', () => {
  it('MINOR lane growth charges the transient peak of old plus new buffers', async () => {
    // 32 numeric rows: capacity grows 16 -> 32 at the 17th entry and needs no trim.
    // Final retention is 256 + 32*22 = 960 B, but the growth peak holds the old
    // 16-slot lanes (352 B) and the new 32-slot lanes (704 B) at once: 1,312 B.
    const doc = objectDoc([{ sheetId: 'A', rows: column(Array.from({ length: 32 }, (_v, i) => i)) }])
    const owner = initialIdentity('doc', 'A')
    const open = async (maxBytes: number) => {
      const kernel = new ViewerKernel({ parse: async () => doc }, { ...quiet, preparedBudget: { maxColumns: 8, maxBytes } })
      completed(await run(kernel, { kind: 'open', requestId: rid(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }))
      return kernel
    }
    const tight = await open(1000)
    expect(await run(tight, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) }))
      .toMatchObject({ kind: 'quota', limitBytes: 1000, requiredBytes: 256 + 16 * 22 + 32 * 22 })
    expect(tight.session('doc')!.preparedStats()).toEqual({ columns: 0, bytes: 0 })
    const roomy = await open(1312)
    completed(await run(roomy, { kind: 'prepareColumns', requestId: rid(), owner, cols: new Uint16Array([0]) }))
    expect(roomy.session('doc')!.preparedStats()).toEqual({ columns: 1, bytes: 256 + 32 * 22 })
  })

  it('NIT structural handles keep one entry per sheet and drop stale value generations', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: column([2, 1]) }])
    const inner = createObjectWorkbookSource(doc, 'doc')
    let valueGen = 0
    const base = inner.store('A')
    const store = Object.create(base) as SourceStore
    Object.defineProperty(store, 'stamp', { get: () => ({ docId: 'doc', sheetId: 'A', valueGen }) })
    store.queryStructural = async (rect, options) => ({ ...await base.queryStructural(rect, { ...options, stamp: base.stamp }), stamp: { docId: 'doc', sheetId: 'A', valueGen } })
    const session = new DocumentSession('doc', { ...inner, sheets: inner.sheets, store: () => store, geometry: id => inner.geometry(id), release: () => inner.release() }, quiet)
    for (valueGen = 0; valueGen < 3; valueGen++) {
      // Each generation builds (and caches) its structural index before the kernels run.
      await session.sort(session.identity('A'), sortSpec(1, 0), new AbortController().signal)
    }
    expect(session.structureCacheSize()).toBe(1)
  })
})
