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
  cacheKey, createObjectWorkbookSource, DocumentSession, workbookIndexKey, type FilterEvaluator, type ViewDraft,
  type WorkbookSource,
} from '../src/worker/viewerSession'
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
    expect(kernel.ledger()).toEqual({ copies: 0, copiedBytes: 0, transfers: 0, transferredBytes: 0 })
    const owner = initialIdentity('doc', 'A')
    const cols = new Uint16Array([0])
    const prepared = await client.send({ kind: 'prepareColumns', requestId: client.nextRequestId(), owner, cols })
    expect(cols.byteLength).toBe(0) // ownership moved to the worker
    expect(client.ledger()).toEqual({ copies: 0, copiedBytes: 0, transfers: 1, transferredBytes: 2 })
    expect(completed(prepared)).toMatchObject({ kind: 'prepareColumns' })
    expect(kernel.ledger()).toEqual({ copies: 1, copiedBytes: 2, transfers: 1, transferredBytes: 2 })
    const before = kernel.ledger()
    const sorted = await client.send({ kind: 'sort', requestId: client.nextRequestId(), owner, spec: sortSpec(4, 0) })
    const delta = deltaOf(sorted)
    const arrays = [...delta.bands.flatMap(band => [band.sourceAtOffset, band.displayOffsetOfSource]), ...delta.visibility.map(bitset => bitset.words)]
    const bytes = arrays.reduce((sum, array) => sum + array.byteLength, 0)
    expect(arrays.length).toBe(3)
    expect(kernel.ledger()).toEqual({
      copies: before.copies + arrays.length, copiedBytes: before.copiedBytes + bytes,
      transfers: before.transfers + arrays.length, transferredBytes: before.transferredBytes + bytes,
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
    expect(kernel.ledger()).toEqual({ copies: 1, copiedBytes: blob.size, transfers: 0, transferredBytes: 0 })
    deltaOf(await run(kernel, { kind: 'sort', requestId: rid(), owner: initialIdentity('doc', '1'), spec: sortSpec(4, 1) }))
    // Plus the two 5-row Uint32 band copies (2 x 20B); no hidden rows, so no bitset.
    expect(kernel.ledger()).toEqual({ copies: 3, copiedBytes: blob.size + 40, transfers: 2, transferredBytes: 40 })
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
