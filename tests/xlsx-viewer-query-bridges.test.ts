// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess, ViewerKernel } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { initialIdentity } from '../src/worker/viewerProtocol'
import { WorkbookFind } from '../src/xlsx/view/find'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { queryViewport } from '../src/xlsx/view/viewport'
import type { OverviewQuery } from '../src/xlsx/view/contracts'
import { findInputOf, workbookOf } from './helpers/xlsx-find'
import type { XlsxDocument } from '../src/xlsx/types'
import type { WorkerCommand } from '../src/xlsx/view/contracts'
import { cellDisplayText } from '../src/xlsx/render'
import type { WorkbookSource } from '../src/worker/viewerSession'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { xlsxBytes } from './helpers/xlsx-package'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createIntervalSet } from '../src/xlsx/view/visibility'
import { createRowOrder } from '../src/xlsx/view/order'

const quiet = { yield: async () => {} }
const document = () => workbookOf([
  { sheetId: 's', cells: [0, 1, 2, 3, 4].map(row => ({ row, col: 0, value: `alpha-${row}` })) },
  { sheetId: 'other', cells: [{ row: 1_048_575, col: 16_383, value: 'alpha-corner' }] },
])

async function harness(doc: XlsxDocument = document(), override?: WorkbookSource) {
  const source = override ?? createObjectWorkbookSource(doc, 'doc')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, quiet)
  const port = connectInProcess(kernel)
  const client = new ViewerClient(port)
  const open = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
  expect(open.kind).toBe('complete')
  const session = kernel.session('doc')!
  const remote = new RemoteWorkbook(client, { docId: 'doc', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], layouts: new Map(source.sheets.map(s => [s.sheetId, session.layout(s.sheetId)])) })
  return { doc, source, kernel, port, client, session, remote }
}

it('serves the same exact five-cell overview through the remote store without detaching borrowed membership', async () => {
  const h = await harness()
  try {
    const rows = Uint32Array.of(0, 1, 2, 3, 4)
    const query: OverviewQuery = { groups: [{ groupId: 'five', rows: { kind: 'list', r0: 0, r1: 4, rows }, cols: [{ c0: 0, c1: 0 }] }], maxGroups: 8 }
    const store = h.source.store('s'), remote = h.remote.store('s')
    const options = { signal: new AbortController().signal, stamp: store.stamp }
    const resident = await store.queryOverview(query, options)
    expect(resident).toMatchObject({ completeness: 'complete', items: [{ populatedCount: 5, countKind: 'exact' }] })
    expect(await remote.queryOverview(query, options)).toEqual(resident)
    expect([...rows]).toEqual([0, 1, 2, 3, 4])
  } finally { await h.remote.release() }
})

it('serves all three bridges over actual paged XLSX storage with resident parity', async () => {
  const bytes = xlsxBytes({ sheetData: Array.from({ length: 5 }, (_, r) => `<row r="${r + 1}"${r === 1 ? ' hidden="1"' : ''}><c r="A${r + 1}" t="inlineStr"><is><t>alpha-${r}</t></is></c></row>`).join('')
    + '<row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>alpha-corner</t></is></c></row>' })
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const imported = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'doc', profile: 'D', store: new SourcePageStore({ docId: 'doc', backend: memoryPageBackend() }) })
  expect(await imported.done).toMatchObject({ kind: 'complete' })
  const h = await harness(doc, createPagedWorkbookSource(imported))
  const resident = createObjectWorkbookSource(doc, 'doc')
  try {
    const query: OverviewQuery = { groups: [{ groupId: 'five', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, cols: [{ c0: 0, c1: 0 }] }], maxGroups: 1 }
    const options = { signal: new AbortController().signal, stamp: resident.store('1').stamp }
    expect(await h.remote.store('1').queryOverview(query, options)).toEqual(await resident.store('1').queryOverview(query, options))
    const owner = h.session.identity('1'), window = { box: { x: 0, y: 0, width: 128, height: 100 }, zoom: 1, deviceScale: 1 }
    await h.session.ensureView('1', options.signal)
    const view = h.session.view('1'), geometry = resident.geometry('1')
    const expected = await queryViewport(resident.store('1'), createSheetMetrics({ rows: geometry.rows, cols: geometry.cols, order: view.order, visibility: view.visibility }), window, owner, options)
    const viewport = await h.remote.viewport(owner, window, options.signal)
    expect(viewport).toMatchObject({ kind: 'complete' })
    if (typeof viewport === 'object' && viewport.kind === 'complete' && expected.kind === 'viewport') expect(viewport.result.content).toEqual({ kind: 'detail', cells: expected.cells })
    const engine = new WorkbookFind(findInputOf(doc).input, { text: 'alpha', caseSensitive: false }, quiet)
    const page = await engine.nextPage()
    engine.close()
    const search = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text: 'alpha', caseSensitive: false } })
    expect(search).toMatchObject({ kind: 'complete', result: { items: page.hits, page: { complete: true, totalKnown: 5 } } })
    expect(page.hits.at(-1)?.source).toEqual({ row: 1_048_575, col: 16_383 })
  } finally { await h.remote.release(); await resident.release() }
})

it('answers coarse viewport groups analytically without any detail read over tiny full-capacity axes', async () => {
  const h = await harness()
  const original = h.source.geometry('s')
  const tiny = { ...original,
    rows: createAxisMetrics({ logicalCount: 1_048_576, defaultSizeCss: 0.01 }),
    cols: createAxisMetrics({ logicalCount: 16_384, defaultSizeCss: 0.01 }),
  }
  vi.spyOn(h.source, 'geometry').mockImplementation(() => tiny)
  const reads = vi.spyOn(h.source.store('s'), 'readRows')
  try {
    const out = await h.remote.viewport(initialIdentity('doc', 's'), { box: { x: 0, y: 0, width: 100, height: 100 }, zoom: 1, deviceScale: 1 }, new AbortController().signal)
    expect(out).toMatchObject({ kind: 'complete', result: { completeness: 'complete', content: { kind: 'overview' } } })
    if (typeof out === 'object' && out.kind === 'complete' && out.result.content.kind === 'overview') {
      expect(out.result.content.groups.length).toBeLessThanOrEqual(2048)
      expect(out.result.content.groups.some(group => 'populatedCount' in group && group.populatedCount === 5)).toBe(true)
    }
    expect(reads).not.toHaveBeenCalled()
  } finally { vi.restoreAllMocks(); await h.remote.release() }
})

it('coalesces concurrent viewport projection admission instead of retaining duplicate caches', async () => {
  const h = await harness(), owner = initialIdentity('doc', 's')
  const window = { box: { x: 0, y: 0, width: 128, height: 100 }, zoom: 1, deviceScale: 1 }
  try {
    await h.remote.viewport(owner, window, new AbortController().signal)
    const expected = h.session.memoryStats().retainedBytes
    h.session.memory.trim(0)
    await Promise.all([h.remote.viewport(owner, window, new AbortController().signal), h.remote.viewport(owner, window, new AbortController().signal)])
    expect(h.session.memoryStats().retainedBytes, 'The same projection must remain charged and retained once').toBe(expected)
  } finally { await h.remote.release() }
})

it('searches canonical formatted display text rather than only raw numbers', async () => {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 0.25 }] }])
  doc.sheets[0].rows[0].cells[0].style = { numFmtId: 10 }
  const h = await harness(doc)
  const text = cellDisplayText(doc.sheets[0], doc.sheets[0].rows[0].cells[0])
  try {
    const out = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text, caseSensitive: false } })
    expect(out.kind).toBe('complete')
    if (out.kind === 'complete') expect(out.result.items, 'The worker must search the same formatted text as the viewer').toMatchObject([{ text, source: { row: 0, col: 0 } }])
  } finally { await h.remote.release() }
})

it('pages a 450-hit search and releases an abandoned continuation through a document-owned read', async () => {
  const h = await harness(workbookOf([{ sheetId: 's', cells: Array.from({ length: 450 }, (_, row) => ({ row, col: 0, value: 'hit' })) }]))
  try {
    const owner = h.session.workbookOwner(), query = { text: 'hit', caseSensitive: false }
    const first = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner, query })
    expect(first.kind).toBe('complete')
    if (first.kind !== 'complete') return
    expect(first.result.items).toHaveLength(200)
    expect(first.result.page).toMatchObject({ complete: false })
    const cursor = first.result.nextCursor!
    const next = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner, query, cursor })
    expect(next.kind).toBe('complete')
    if (next.kind !== 'complete') return
    expect(next.result.items).toHaveLength(200)
    expect(next.result.items[0].source.row).toBe(200)
    const replay = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner, query, cursor })
    expect(replay).toMatchObject({ kind: 'error', code: 'invalid-cursor' })
    expect(h.session.memoryStats().transientBytes).toBeGreaterThan(0)
    const cleaned = await h.client.send({ kind: 'read', requestId: h.client.nextRequestId(), owner: initialIdentity('doc', 's'), query: { op: 'releaseSearch', cursor: next.result.nextCursor! } })
    expect(cleaned).toMatchObject({ kind: 'complete', result: { result: { op: 'releaseSearch', released: true } } })
    expect(h.session.memoryStats().transientBytes).toBe(0)
  } finally { await h.remote.release() }
})

it('reclaims a retained search immediately when its workbook generation changes', async () => {
  const h = await harness(workbookOf([{ sheetId: 's', cells: Array.from({ length: 450 }, (_, row) => ({ row, col: 0, value: 'hit' })) }]))
  try {
    const out = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text: 'hit', caseSensitive: false } })
    expect(out).toMatchObject({ kind: 'complete' })
    expect(h.session.memoryStats().transientBytes).toBeGreaterThan(0)
    h.session.publishFonts()
    expect(h.session.memoryStats().transientBytes, 'A stale search must release its hit workspace and view leases').toBe(0)
  } finally { await h.remote.release() }
})

it('admits overview membership before cloning or posting an oversized caller buffer', async () => {
  const h = await harness()
  try {
    const groups = [0, 1].map(i => ({ groupId: String(i), rows: { kind: 'list' as const, r0: 0, r1: 1_048_575, rows: new Uint32Array(1_048_576) }, cols: [{ c0: 0, c1: 0 }] }))
    const before = h.remote.stats.commands['read:overview'] ?? 0
    const out = await h.remote.store('s').queryOverview({ groups, maxGroups: 2 }, { signal: new AbortController().signal, stamp: h.remote.store('s').stamp })
    expect(out).toMatchObject({ completeness: 'incomplete', error: { code: 'resource-limit' } })
    expect(h.remote.stats.commands['read:overview'] ?? 0).toBe(before)
    expect(groups[0].rows.rows.byteLength).toBe(4_194_304)
  } finally { await h.remote.release() }
})

it('does not finalize search over the used-range prefix of a progressively importing sheet', async () => {
  const h = await harness()
  const store = h.source.store('s'), original = store.scanPopulated.bind(store)
  vi.spyOn(store, 'scanPopulated').mockImplementation(async function* (query, options) {
    for await (const page of original(query, options)) yield { ...page, completeness: query.range.r1 === 1_048_575 ? 'loading' as const : page.completeness }
  })
  try {
    const out = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text: 'alpha', caseSensitive: false } })
    expect(out.kind).toBe('complete')
    if (out.kind === 'complete') {
      expect(out.result.page?.incomplete, 'An importing sheet cannot report a complete prefix as the whole search domain').toContain('s')
      expect(out.result.page?.totalKnown).toBeUndefined()
    }
  } finally { vi.restoreAllMocks(); await h.remote.release() }
})

it('uses authored hidden rows learned after a view was first materialized', async () => {
  const h = await harness()
  await h.session.ensureView('s', new AbortController().signal)
  const original = h.source.geometry('s')
  vi.spyOn(h.source, 'geometry').mockImplementation(() => ({ ...original, baseHidden: createIntervalSet([{ r0: 2, r1: 2 }]) }))
  try {
    const out = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text: 'alpha', caseSensitive: false }, })
    expect(out.kind).toBe('complete')
    if (out.kind === 'complete') expect(out.result.items.some(item => item.sheetId === 's' && item.source.row === 2), 'New authored-hidden metadata must be respected by the worker query').toBe(false)
  } finally { vi.restoreAllMocks(); await h.remote.release() }
})

it('maps authored hidden source membership through a non-identity committed order', async () => {
  const h = await harness(workbookOf([{ sheetId: 's', hiddenRows: [2], cells: [0, 1, 2, 3].map(row => ({ row, col: 0, value: `alpha-${row}` })) }]))
  const order = createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: Uint32Array.of(0, 3, 1, 2) }])
  try {
    await h.session.ensureView('s', new AbortController().signal)
    const { projectVisibility } = await import('../src/xlsx/view/visibility')
    await h.session.commitView('s', { order, visibility: projectVisibility(h.source.geometry('s').baseHidden, null, order, () => true).visibility, filter: null, sort: null })
    const out = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text: 'alpha', caseSensitive: false } })
    expect(out.kind).toBe('complete')
    if (out.kind === 'complete') expect(out.result.items.map(item => item.source.row)).toEqual([0, 3, 1])
  } finally { await h.remote.release() }
})

it.each(['overview', 'viewport', 'search'] as const)('aborts %s reads and withholds late publication', async kind => {
  const h = await harness()
  let release!: () => void, began!: () => void, readSignal: AbortSignal | undefined
  const blocked = new Promise<void>(resolve => { release = resolve })
  const started = new Promise<void>(resolve => { began = resolve })
  const store = h.source.store('s'), owner = initialIdentity('doc', 's')
  let command: WorkerCommand
  if (kind === 'overview') {
    const original = store.queryOverview.bind(store)
    vi.spyOn(store, 'queryOverview').mockImplementation(async (q, o) => { readSignal = o.signal; began(); await blocked; return original(q, { ...o, signal: new AbortController().signal }) })
    command = { kind: 'read', requestId: h.client.nextRequestId(), owner, query: { op: 'overview', request: { groups: [{ groupId: 'g', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, cols: [{ c0: 0, c1: 0 }] }], maxGroups: 1 } } }
  } else if (kind === 'viewport') {
    const original = store.readRows.bind(store)
    vi.spyOn(store, 'readRows').mockImplementation(async (r, c, o) => { readSignal = o.signal; began(); await blocked; return original(r, c, { ...o, signal: new AbortController().signal }) })
    command = { kind: 'viewport', requestId: h.client.nextRequestId(), owner, window: { box: { x: 0, y: 0, width: 100, height: 100 }, zoom: 1, deviceScale: 1 } }
  } else {
    const original = store.scanPopulated.bind(store)
    vi.spyOn(store, 'scanPopulated').mockImplementation(async function* (q, o) { readSignal = o.signal; began(); await blocked; yield* original(q, { ...o, signal: new AbortController().signal }) })
    command = { kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text: 'alpha', caseSensitive: false } }
  }
  const pending = h.client.send(command)
  try {
    await started
    await h.client.send({ kind: 'cancel', requestId: h.client.nextRequestId(), targetRequestId: command.requestId })
    expect(readSignal!.aborted).toBe(true)
    release()
    expect(await pending).toMatchObject({ kind: 'cancelled' })
    await h.port.idle()
    expect(h.session.memoryStats().transientBytes).toBe(0)
  } finally { release(); vi.restoreAllMocks(); await h.remote.release() }
})

it('surfaces a timed-out remote query as a visible retry notice without marking the import failed', async () => {
  vi.useFakeTimers()
  const h = await harness()
  vi.spyOn(h.source.store('s'), 'queryOverview').mockImplementation(() => new Promise(() => {}))
  try {
    const pending = h.remote.store('s').queryOverview({ groups: [{ groupId: 'g', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, cols: [{ c0: 0, c1: 0 }] }], maxGroups: 1 }, { signal: new AbortController().signal, stamp: h.source.store('s').stamp })
    await vi.advanceTimersByTimeAsync(120_001)
    expect(await pending).toMatchObject({ completeness: 'incomplete', error: { code: 'cancelled', retryable: true } })
    expect(h.remote, 'The UI must receive the terminal query timeout notice').toMatchObject({ queryNotice: expect.stringMatching(/timed out after 120 seconds/) })
    expect(h.remote.valueStatus().ended).not.toBe(true)
  } finally { vi.restoreAllMocks(); await h.remote.release(); vi.useRealTimers() }
})

it.each(['replaced', 'caller-cancelled'] as const)('does not publish an old timeout notice after the query was %s', async action => {
  vi.useFakeTimers()
  const h = await harness(), controller = new AbortController()
  vi.spyOn(h.source.store('s'), 'queryOverview').mockImplementationOnce(() => new Promise(() => {}))
  const store = h.remote.store('s'), query: OverviewQuery = { groups: [{ groupId: 'g', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, cols: [{ c0: 0, c1: 0 }] }], maxGroups: 1 }
  try {
    const old = store.queryOverview(query, { signal: controller.signal, stamp: store.stamp })
    for (let i = 0; i < 30; i++) await Promise.resolve()
    if (action === 'replaced') expect(await store.queryOverview(query, { signal: new AbortController().signal, stamp: store.stamp })).toMatchObject({ completeness: 'complete' })
    else controller.abort()
    await vi.advanceTimersByTimeAsync(120_001)
    await old
    expect(h.remote.queryNotice, 'An obsolete request must not replace the current visible status').toBeUndefined()
  } finally { vi.restoreAllMocks(); await h.remote.release(); vi.useRealTimers() }
})

it('returns resident-equivalent bounded viewport source records over the worker command', async () => {
  const h = await harness()
  try {
    const owner = initialIdentity('doc', 's'), window = { box: { x: 0, y: 0, width: 128, height: 100 }, zoom: 1, deviceScale: 1 }
    await h.session.ensureView('s', new AbortController().signal)
    const view = h.session.view('s'), geometry = h.source.geometry('s')
    const metrics = createSheetMetrics({ rows: geometry.rows, cols: geometry.cols, order: view.order, visibility: view.visibility })
    const resident = await queryViewport(h.source.store('s'), metrics, window, owner, { signal: new AbortController().signal })
    expect(resident.kind).toBe('viewport')
    const out = await h.client.send({ kind: 'viewport', requestId: h.client.nextRequestId(), owner, window })
    expect(out.kind, 'Worker viewport must resolve admitted detail records').toBe('complete')
    if (out.kind !== 'complete' || resident.kind !== 'viewport') return
    expect(out.result).toMatchObject({ kind: 'viewport', completeness: 'complete', content: { kind: 'detail', cells: resident.cells } })
    expect(resident.cells.length).toBeLessThanOrEqual(20_000)
  } finally { await h.remote.release() }
})

it('returns resident-equivalent search hits including the inactive far corner', async () => {
  const h = await harness()
  const query = { text: 'alpha', caseSensitive: false }
  const resident = new WorkbookFind(findInputOf(h.doc).input, query, quiet)
  try {
    const page = await resident.nextPage()
    const out = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query })
    expect(out.kind, 'Worker search must scan the complete visible workbook').toBe('complete')
    if (out.kind !== 'complete') return
    expect(out.result.items).toEqual(page.hits)
    expect(out.result.items.at(-1)?.source).toEqual({ row: 1_048_575, col: 16_383 })
    expect(out.result.unavailableCount).toBe(page.unavailableCount)
  } finally { resident.close(); await h.remote.release() }
})

it('rejects stale sheet and workbook generations before viewport or search reads', async () => {
  const h = await harness()
  try {
    h.session.publishFonts()
    const viewport = await h.client.send({ kind: 'viewport', requestId: h.client.nextRequestId(), owner: initialIdentity('doc', 's'), window: { box: { x: 0, y: 0, width: 100, height: 100 }, zoom: 1, deviceScale: 1 } })
    const search = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: { ...h.session.workbookOwner(), fontGen: 0 }, query: { text: 'alpha', caseSensitive: false } })
    expect(viewport).toMatchObject({ kind: 'error', code: 'stale-source' })
    expect(search).toMatchObject({ kind: 'error', code: 'stale-source' })
  } finally { await h.remote.release() }
})
