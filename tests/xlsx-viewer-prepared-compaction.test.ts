import { expect, it, vi } from 'vitest'
import { DocumentSession, PREPARED_BUDGETS, type WorkbookSource } from '../src/worker/viewerSession'
import { initialIdentity } from '../src/worker/viewerProtocol'
import { buildGeometry } from '../src/xlsx/view/source'
import { sortRowOrder } from '../src/xlsx/view/sort'
import { PREPARED_SORT_READER, type PreparedSortStore } from '../src/xlsx/view/sortPrepared'
import { sortRowOrder as baseline } from './helpers/sort-baseline'
import { createFixture, type FixtureCell } from './helpers/xlsx-viewer-fixtures'
import type { SortSpec, SourceStore } from '../src/xlsx/view/contracts'

function sessionFor(store: SourceStore) {
  const source: WorkbookSource = {
    valuePolicy: store.valuePolicy, sheets: [{ sheetId: store.stamp.sheetId, name: 'Prepared' }], store: () => store,
    geometry: () => buildGeometry([], []), coverageDiagnostics: () => [], release: async () => { await store.release() },
  }
  return new DocumentSession(store.stamp.docId, source, { profile: 'D', now: () => 0, yield: async () => {} })
}

it('admits every full-size F2 key in one 64 MiB batch and preserves the frozen permutation', async () => {
  const raw = createFixture('F2').createStore()
  let reads = 0
  const store = { ...raw, queryColumn: (...args: Parameters<SourceStore['queryColumn']>) => { reads++; return raw.queryColumn(...args) } }
  const session = sessionFor(store), cols = [2, 3, 7, 1, 5, 4, 6, 0]
  try {
    const result = await session.prepareColumns(initialIdentity(store.stamp.docId, store.stamp.sheetId), cols, new AbortController().signal)
    expect(result).toEqual({ kind: 'complete' })
    expect(session.preparedColumns(store.stamp.sheetId)).toEqual([0, 1, 2, 3, 4, 5, 6, 7])
    expect(session.preparedStats().bytes).toBeLessThanOrEqual(PREPARED_BUDGETS.D.maxBytes)
    expect(session.memoryStats().transientBytes).toBe(0)
    const spec: SortSpec = { range: { r0: 0, r1: 1048575, c0: 0, c1: 7, headerRows: 1, totalsRows: 1 },
      keys: cols.map(col => ({ col, order: 'asc' })), caseSensitive: false }
    reads = 0
    const current = await sortRowOrder(session.preparedStore(store.stamp.sheetId), spec, { now: () => 0, yield: async () => {} })
    expect(reads).toBe(0)
    const old = await baseline(raw, spec, { now: () => 0, yield: async () => {} })
    if (current.kind !== 'complete' || old.kind !== 'complete') throw new Error(JSON.stringify({ current: current.kind, old: old.kind }))
    expect(Buffer.from(Uint32Array.from(current.order.bands[0].sourceAtOffset).buffer).equals(Buffer.from(Uint32Array.from(old.order.bands[0].sourceAtOffset).buffer))).toBe(true)
    expect(Buffer.from(Uint32Array.from(current.order.bands[0].displayOffsetOfSource).buffer).equals(Buffer.from(Uint32Array.from(old.order.bands[0].displayOffsetOfSource).buffer))).toBe(true)
    console.info('Prepared F2:', JSON.stringify({ retained: session.preparedStats(), memory: session.memoryStats(), columns: cols.map(col => ({ col, ...session.preparedFootprint(store.stamp.sheetId, col) })) }))
  } finally { await session.release() }
}, 120_000)

it('reconstructs mixed payloads exactly across page boundaries and sparse row gaps', async () => {
  const n = 5003
  const values = [-0, 0, 0.1, Math.PI, 2 ** 32 + 1, -2147483649, Infinity, NaN, true, false, null, '', '\ud800', 'é', 'e\u0301', '#N/A']
  const cells: FixtureCell[] = Array.from({ length: n }, (_, row) => ({
    value: values[row % values.length], styleIndex: [0, 256, 65536, 4294967295][row % 4],
    hasCachedValue: row % 3 !== 0, valueIsError: row % 13 === 0,
    ...(row % 11 === 0 ? { formula: 'X()', sharedFormula: { si: row, ref: 'A1:A3' }, arrayRef: 'A1:A3', dynamicArray: true, formulaType: 'array', ca: false } : {}),
  }))
  const store = createFixture('F2', { rowCount: n }).createStore({ read: { cellAt: (row, col) => col === 0 && row % 19 !== 0 ? cells[row] ?? null : null } })
  const session = sessionFor(store), owner = initialIdentity(store.stamp.docId, store.stamp.sheetId)
  try {
    expect(await session.prepareColumns(owner, [0], new AbortController().signal)).toEqual({ kind: 'complete' })
    const prepared = session.preparedStore(owner.sheetId)
    const options = { signal: new AbortController().signal, stamp: store.stamp }
    for (const [r0, r1] of [[0, 40], [4090, 4130], [4300, 4350], [4950, 5002]]) {
      const request = { col: 0, rows: { kind: 'intervals' as const, intervals: [{ r0, r1 }] }, limit: 200 }
      expect((await prepared.queryColumn(request, options)).items).toEqual((await store.queryColumn(request, options)).items)
    }
  } finally { await session.release() }
})

it('serves bounded typed key pages, with cell fallback for repeated giant strings', async () => {
  for (const large of [false, true]) {
    const value = large ? 'x'.repeat(100_000) : 'item-2'
    const store = createFixture('F2', { rowCount: 100 }).createStore({ read: { cellAt: (row, col) => col === 0 ? { value: row % 2 ? value : row, hasCachedValue: true, valueIsError: row % 2 !== 0, styleIndex: 0 } : null } })
    const session = sessionFor(store), owner = initialIdentity(store.stamp.docId, store.stamp.sheetId)
    try {
      expect(await session.prepareColumns(owner, [0], new AbortController().signal)).toEqual({ kind: 'complete' })
      const response = await session.preparedStore(owner.sheetId).queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 99 }] }, limit: 100, encoding: 'lanes' }, { signal: new AbortController().signal, stamp: store.stamp })
      expect(response.items.some(item => item.kind === 'lanes')).toBe(!large)
      if (!large) {
        expect(response.items[0]).toMatchObject({ kind: 'lanes', r0: 0, r1: 99 })
        const entries = await session.preparedStore(owner.sheetId).queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 99 }] }, limit: 100 }, { signal: new AbortController().signal, stamp: store.stamp })
        expect(entries.items.some(item => item.kind === 'lanes')).toBe(false)
      }
    } finally { await session.release() }
  }
})

it('keeps typed cursors positional through eviction while rejecting encoding changes', async () => {
  const store = createFixture('F2', { rowCount: 20 }).createStore({ read: { cellAt: (row, col) => ({ value: row + col, styleIndex: 0, hasCachedValue: true, valueIsError: false }) } })
  const source: WorkbookSource = { valuePolicy: store.valuePolicy, sheets: [{ sheetId: store.stamp.sheetId, name: 'Cursor' }], store: () => store,
    geometry: () => buildGeometry([], []), coverageDiagnostics: () => [], release: async () => { await store.release() } }
  const session = new DocumentSession(store.stamp.docId, source, { now: () => 0, yield: async () => {}, preparedBudget: { maxColumns: 1, maxBytes: 64 * 1024 * 1024 } })
  const owner = initialIdentity(store.stamp.docId, store.stamp.sheetId), opts = { signal: new AbortController().signal, stamp: store.stamp }
  try {
    await session.prepareColumns(owner, [0], opts.signal)
    const query = { col: 0, rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 10 }] }, limit: 2, encoding: 'lanes' as const }
    const first = await session.preparedStore(owner.sheetId).queryColumn(query, opts)
    expect(first.items[0]).toMatchObject({ kind: 'lanes', r0: 0, r1: 1 })
    expect((await session.preparedStore(owner.sheetId).queryColumn({ ...query, cursor: first.nextCursor, encoding: 'entries' }, opts)).error?.code).toBe('invalid-cursor')
    await session.prepareColumns(owner, [1], opts.signal)
    expect(session.isPrepared(owner.sheetId, 0)).toBe(false)
    const next = await session.preparedStore(owner.sheetId).queryColumn({ ...query, cursor: first.nextCursor }, opts)
    expect(next.error).toBeUndefined()
    expect(next.items[0]).toMatchObject({ kind: 'cell', row: 2 })
  } finally { await session.release() }
})

it('keeps repeated ordinary text dictionary-backed instead of encoding it on every read', async () => {
  const store = createFixture('F2', { rowCount: 20 }).createStore({ read: { cellAt: (_row, col) => col === 0 ? { value: 'ordinary-text', styleIndex: 0, hasCachedValue: true, valueIsError: false } : null } })
  const session = sessionFor(store), owner = initialIdentity(store.stamp.docId, store.stamp.sheetId)
  try {
    await session.prepareColumns(owner, [0], new AbortController().signal)
    const page = await session.preparedStore(owner.sheetId).queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 19 }] }, limit: 20, encoding: 'lanes' }, { signal: new AbortController().signal, stamp: store.stamp })
    expect(page.items.every(item => item.kind === 'cell')).toBe(true)
  } finally { await session.release() }
})

it('reserves the raw old-plus-new growth peak in the worker pool before allocating', async () => {
  const store = createFixture('F2', { rowCount: 32 }).createStore({ read: { cellAt: (row, col) => col === 0 ? { value: row, styleIndex: 0, hasCachedValue: true, valueIsError: false } : null } })
  const session = sessionFor(store), owner = initialIdentity(store.stamp.docId, store.stamp.sheetId)
  const Native = Float64Array
  let chargedAtGrowth = 0
  const ctor = vi.spyOn(globalThis, 'Float64Array').mockImplementation(function (length: number) {
    if (length === 32) chargedAtGrowth = session.memoryStats().transientBytes
    return new Native(length)
  } as unknown as typeof Float64Array)
  try {
    expect(await session.prepareColumns(owner, [0], new AbortController().signal)).toEqual({ kind: 'complete' })
    expect(chargedAtGrowth).toBeGreaterThanOrEqual(256 + (16 + 32) * 22)
  } finally { ctor.mockRestore(); await session.release() }
})

it('reads pinned prepared scalars directly without rebuilding source pages', async () => {
  const values = [null, false, true, -0, 0.1, Math.PI, 'A', 'a', 'é', 'e\u0301', '\ud800', '#N/A']
  const n = 5003
  const store = createFixture('F2', { rowCount: n }).createStore({ read: { cellAt: (row, col) => ({ value: values[(row + col) % values.length], styleIndex: 0, valueIsError: col === 1 && row % 13 === 0, hasCachedValue: true }) } })
  const session = sessionFor(store), owner = initialIdentity(store.stamp.docId, store.stamp.sheetId)
  try {
    await session.prepareColumns(owner, [0, 1], new AbortController().signal)
    const prepared = session.preparedStore(owner.sheetId), query = prepared.queryColumn.bind(prepared)
    let reads = 0
    prepared.queryColumn = (...args) => { reads++; return query(...args) }
    const spec: SortSpec = { range: { r0: 0, r1: n - 1, c0: 0, c1: 1, headerRows: 1, totalsRows: 1 }, keys: [{ col: 0, order: 'desc' }, { col: 1, order: 'asc' }], caseSensitive: false }
    const opts = { now: () => 0, yield: async () => {} }, current = await sortRowOrder(prepared, spec, opts), old = await baseline(store, spec, opts)
    if (current.kind !== 'complete' || old.kind !== 'complete') throw new Error(JSON.stringify({ current: current.kind, old: old.kind }))
    expect(current.order.bands[0].sourceAtOffset).toEqual(old.order.bands[0].sourceAtOffset)
    expect(reads).toBe(0)
  } finally { await session.release() }
})

it('pins prepared readers against explicit LRU eviction and releases the pin', async () => {
  const store = createFixture('F2', { rowCount: 20 }).createStore()
  const source: WorkbookSource = { valuePolicy: store.valuePolicy, sheets: [{ sheetId: store.stamp.sheetId, name: 'Pins' }], store: () => store,
    geometry: () => buildGeometry([], []), coverageDiagnostics: () => [], release: async () => { await store.release() } }
  const session = new DocumentSession(store.stamp.docId, source, { now: () => 0, yield: async () => {}, preparedBudget: { maxColumns: 1, maxBytes: 64 * 1024 * 1024 } })
  const owner = initialIdentity(store.stamp.docId, store.stamp.sheetId), signal = new AbortController().signal
  try {
    await session.prepareColumns(owner, [0], signal)
    const prepared = session.preparedStore(owner.sheetId) as PreparedSortStore
    const lease = prepared[PREPARED_SORT_READER]?.(0, store.stamp)
    expect(lease).toBeDefined()
    if (!lease) throw new Error('missing lease')
    expect((await session.prepareColumns(owner, [1], signal)).kind).toBe('quota')
    expect(session.isPrepared(owner.sheetId, 0)).toBe(true)
    lease.release(); lease.release()
    expect(await session.prepareColumns(owner, [1], signal)).toEqual({ kind: 'complete' })
    expect(session.isPrepared(owner.sheetId, 0)).toBe(false)
  } finally { await session.release() }
})
it('rejects malformed cell identities before making a trusted prepared reader', async () => {
  const raw = createFixture('F2', { rowCount: 3 }).createStore(), query = raw.queryColumn.bind(raw)
  const store: SourceStore = { ...raw, queryColumn: async (req, opts) => {
    const page = await query(req, opts)
    return { ...page, items: page.items.map(item => item.kind === 'cell' ? { ...item, cell: { ...item.cell, id: item.cell.id + 1 } } : item) }
  } }
  const session = sessionFor(store)
  try {
    expect(await session.prepareColumns(initialIdentity(store.stamp.docId, store.stamp.sheetId), [0], new AbortController().signal)).toMatchObject({ kind: 'error', error: { code: 'page-error' } })
    expect(session.preparedStats().columns).toBe(0)
  } finally { await session.release() }
})

it('rejects a prepared scalar read when its session closes during a read pause', async () => {
  const store = createFixture('F2', { rowCount: 9003 }).createStore(), session = sessionFor(store)
  const owner = initialIdentity(store.stamp.docId, store.stamp.sheetId)
  await session.prepareColumns(owner, [0], new AbortController().signal)
  let pauses = 0
  const spec: SortSpec = { range: { r0: 0, r1: 9002, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }
  try {
    expect(await sortRowOrder(session.preparedStore(owner.sheetId), spec, { now: () => 0, yield: async phase => { if (phase === 'read' && ++pauses === 3) await session.release() } })).toMatchObject({ kind: 'error', error: { code: 'stale-source' } })
  } finally { await session.release() }
})

it('keeps invalidated leased storage charged until its last reader releases', async () => {
  const raw = createFixture('F2', { rowCount: 20 }).createStore()
  let stamp = raw.stamp
  const store: SourceStore = { ...raw, get stamp() { return stamp } }
  const session = sessionFor(store), owner = initialIdentity(stamp.docId, stamp.sheetId)
  try {
    await session.prepareColumns(owner, [0], new AbortController().signal)
    const prepared = session.preparedStore(owner.sheetId) as PreparedSortStore
    const first = prepared[PREPARED_SORT_READER]!(0, stamp)!, second = prepared[PREPARED_SORT_READER]!(0, stamp)!
    const bytes = session.memoryStats().retainedBytes
    stamp = { ...stamp, valueGen: stamp.valueGen + 1 }
    session.publishValues(owner.sheetId)
    expect(session.isPrepared(owner.sheetId, 0)).toBe(false)
    expect(session.memoryStats().retainedBytes).toBe(bytes)
    expect(await first.read(0, 19, () => { throw new Error('stale reader emitted a value') }, { due: () => false, pause: async () => {} })).toBe('stale')
    first.release()
    expect(session.memoryStats().retainedBytes).toBe(bytes)
    second.release()
    expect(session.memoryStats().retainedBytes).toBe(0)
    expect(session.preparedStats().columns).toBe(0)
  } finally { await session.release() }
})
