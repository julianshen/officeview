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
import { createCanvas } from 'canvas'
import { WorkerViewportPlans } from '../src/components/xlsx/workerViewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { boundedQueryStore } from '../src/worker/viewerQueryBudget'
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
  const remote = new RemoteWorkbook(client, { docId: 'doc', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(s => [s.sheetId, session.layout(s.sheetId)])) })
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
  let releaseCleanup!: () => void, cleanupBegan!: () => void, cleanupEnded!: () => void
  const cleanupGate = new Promise<void>(resolve => { releaseCleanup = resolve })
  const cleanupStarted = new Promise<void>(resolve => { cleanupBegan = resolve })
  const cleanupFinished = new Promise<void>(resolve => { cleanupEnded = resolve })
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
    vi.spyOn(store, 'scanPopulated').mockImplementation(async function* (q, o) {
      readSignal = o.signal; began()
      try { await blocked; yield* original(q, { ...o, signal: new AbortController().signal }) }
      finally { cleanupBegan(); await cleanupGate; cleanupEnded() }
    })
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
    if (kind === 'search') {
      // Public cancellation settles before an abort-ignoring provider's
      // queued iterator.return(). Its workspace stays owned through cleanup.
      await cleanupStarted
      expect(h.session.memoryStats().transientBytes).toBeGreaterThan(0)
      releaseCleanup()
      await cleanupFinished
      await vi.waitFor(() => expect(h.session.memoryStats().transientBytes).toBe(0))
    }
    expect(h.session.memoryStats().transientBytes).toBe(0)
  } finally { release(); releaseCleanup(); vi.restoreAllMocks(); await h.remote.release() }
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
    else {
      controller.abort()
      expect(await old).toMatchObject({ completeness: 'incomplete', error: { code: 'cancelled' } })
      for (let i = 0; i < 100; i++) await Promise.resolve()
      expect(h.kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
    }
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
    expect(out.result).toMatchObject({ kind: 'viewport', completeness: 'complete', content: { kind: 'detail', cells: resident.cells }, merges: resident.merges, coarse: resident.coarse, rowLines: resident.rowLines, colLines: resident.colLines, work: resident.work })
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


it('clears a same-scope timeout notice after a successful overview retry', async () => {
  vi.useFakeTimers()
  const h = await harness(), store = h.remote.store('s')
  vi.spyOn(h.source.store('s'), 'queryOverview').mockImplementationOnce(() => new Promise(() => {}))
  const query: OverviewQuery = { groups: [{ groupId: 'g', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, cols: [{ c0: 0, c1: 0 }] }], maxGroups: 1 }
  const options = { signal: new AbortController().signal, stamp: store.stamp }
  try {
    const pending = store.queryOverview(query, options)
    await vi.advanceTimersByTimeAsync(120_000)
    await pending
    expect(h.remote.queryNotice).toMatch(/timed out/)
    expect(await store.queryOverview(query, options)).toMatchObject({ completeness: 'complete' })
    expect(h.remote.queryNotice).toBeUndefined()
  } finally { await h.remote.release(); vi.useRealTimers() }
})

it.each([false, true])('keeps drawing timeouts out of the range-query notice (prior overview timeout: %s)', async priorTimeout => {
  vi.useFakeTimers()
  const doc = document(), source = createObjectWorkbookSource(doc, 'doc')
  const query: OverviewQuery = { groups: [{ groupId: 'g', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, cols: [{ c0: 0, c1: 0 }] }], maxGroups: 1 }
  let release!: () => void
  const held = new Promise<void>(resolve => { release = resolve })
  const provider = vi.spyOn(source.store('s'), 'readDrawings').mockImplementation(async () => {
    await held
    return { completeness: 'complete', drawings: [], images: [], diagnostics: [], bytes: 0 }
  })
  const h = await harness(doc, source), store = h.remote.store('s')
  const options = { signal: new AbortController().signal, stamp: store.stamp }
  try {
    if (priorTimeout) {
      vi.spyOn(h.source.store('s'), 'queryOverview').mockImplementationOnce(() => new Promise(() => {}))
      const pending = store.queryOverview(query, options)
      await vi.advanceTimersByTimeAsync(120_001)
      await pending
      expect(h.remote.queryNotice).toMatch(/timed out after 120 seconds/)
    }
    const notice = h.remote.queryNotice
    const pending = h.remote.request({ kind: 'read', requestId: h.remote.nextRequestId(), owner: h.remote.identity('s'), query: { op: 'drawings' } })
    await vi.advanceTimersByTimeAsync(120_001)
    expect(provider).toHaveBeenCalledTimes(1)
    expect(await pending).toMatchObject({ kind: 'cancelled', diagnostics: [{ code: 'query-timeout' }] })
    expect(h.remote.queryNotice, 'Drawing failures belong to the drawing cache, not the range-query notice').toBe(notice)
    expect(await store.queryOverview(query, options)).toMatchObject({ completeness: 'complete' })
    expect(h.remote.queryNotice, 'Drawing traffic must not steal the overview notice scope').toBeUndefined()
  } finally { release(); await h.remote.release(); vi.restoreAllMocks(); vi.useRealTimers() }
})

it('diagnoses exhausted drawing timeout retries without leaving a smaller-range notice', async () => {
  vi.useFakeTimers()
  const doc = document(), source = createObjectWorkbookSource(doc, 'doc')
  let release!: () => void
  const held = new Promise<void>(resolve => { release = resolve })
  const provider = vi.spyOn(source.store('s'), 'readDrawings').mockImplementation(async () => {
    await held
    return { completeness: 'complete', drawings: [], images: [], diagnostics: [], bytes: 0 }
  })
  const h = await harness(doc, source)
  try {
    h.remote.merges('s')
    await vi.advanceTimersByTimeAsync(0)
    expect(h.remote.loadState('s')).toEqual({ completeness: 'complete' })
    const content = h.remote.drawings('s')
    await vi.advanceTimersByTimeAsync(20 * 120_000 + 19 * 250 + 1)
    expect(provider).toHaveBeenCalledTimes(20)
    expect(content).toMatchObject({ state: 'failed', diagnostics: [{ code: 'drawings-unavailable', message: 'Drawings could not be loaded' }] })
    expect(h.remote.queryNotice).toBeUndefined()
    expect(h.kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
  } finally { release(); await h.remote.release(); vi.restoreAllMocks(); vi.useRealTimers() }
})

it('counts unavailable uncached engine formulas and excludes authored hidden columns in worker search', async () => {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'alpha' }, { row: 1, col: 0, value: { unavailable: true } }, { row: 0, col: 1, value: 'alpha-hidden' }] }])
  doc.sheets[0].cols = [{ min: 1, max: 1, hidden: true }]
  const h = await harness(doc)
  try {
    expect(h.source.valuePolicy).toBe('engine')
    const out = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text: 'alpha', caseSensitive: false } })
    expect(out).toMatchObject({ kind: 'complete', result: { unavailableCount: 1, items: [{ source: { row: 0, col: 0 } }] } })
    if (out.kind === 'complete') expect(out.result.items).toHaveLength(1)
  } finally { await h.remote.release() }
})

it('preserves optional occupancy and settlement capabilities in bounded stores', async () => {
  const source = createObjectWorkbookSource(document(), 'doc'), store = source.store('s')
  const settle = vi.fn(), scanOccupied = vi.fn(async function* () { yield { items: [{ row: 2, col: 0 }], stamp: store.stamp, completeness: 'complete' as const } })
  const bounded = boundedQueryStore({ ...store, scanOccupied, settle }, 1024)
  expect(bounded.scanOccupied).toBeTypeOf('function')
  const pages = []
  for await (const page of bounded.scanOccupied!({ range: { r0: 0, r1: 4, c0: 0, c1: 0 }, limit: 5 }, { signal: new AbortController().signal, stamp: store.stamp })) pages.push(page)
  expect(pages[0].items).toEqual([{ row: 2, col: 0 }])
  bounded.settle!(pages[0])
  expect(settle).toHaveBeenCalledWith(pages[0])
  await source.release()
})

it('charges the main-owned overview clone while its worker read is held', async () => {
  const h = await harness()
  let release!: () => void, began!: () => void
  const held = new Promise<void>(r => { release = r }), started = new Promise<void>(r => { began = r })
  const store = h.source.store('s'), original = store.queryOverview.bind(store)
  vi.spyOn(store, 'queryOverview').mockImplementationOnce(async (q, o) => { began(); await held; return original(q, o) })
  try {
    const before = h.remote.mainLedger.snapshot().byPool.replies
    const pending = h.remote.store('s').queryOverview({ groups: [{ groupId: 'g', rows: { kind: 'list', r0: 0, r1: 4, rows: Uint32Array.of(0, 1, 2, 3, 4) }, cols: [{ c0: 0, c1: 0 }] }], maxGroups: 1 }, { signal: new AbortController().signal, stamp: store.stamp })
    await started
    expect(h.remote.mainLedger.snapshot().byPool.replies).toBeGreaterThan(before)
    release()
    const page = await pending
    h.remote.mainLedger.settle(page)
    expect(h.remote.mainLedger.snapshot().byPool.replies).toBe(before)
  } finally { release(); await h.remote.release() }
})


it('returns loaded hits promptly from a real paged import with a held backend page', async () => {
  vi.useFakeTimers()
  const bytes = xlsxBytes({ sheetData: [1, 2, 3].map(r => `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>alpha-${r}</t></is></c></row>`).join('') })
  const backend = memoryPageBackend(), put = backend.put.bind(backend)
  let release!: () => void, began!: () => void
  const held = new Promise<void>(r => { release = r }), started = new Promise<void>(r => { began = r })
  backend.put = async (key, value) => { if (key.endsWith('/p1')) { began(); await held } await put(key, value) }
  const imported = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'doc', profile: 'D', pageRows: 1, store: new SourcePageStore({ docId: 'doc', backend }) })
  await started
  const h = await harness(document(), createPagedWorkbookSource(imported))
  try {
    const replies: unknown[] = []
    const pending = h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text: 'alpha', caseSensitive: false } }).then(out => { replies.push(out); return out })
    await vi.advanceTimersByTimeAsync(60_000)
    expect(replies, 'Loaded search hits must arrive before the pending import page is released').toHaveLength(1)
    expect(await pending).toMatchObject({ kind: 'complete', result: { items: [{ source: { row: 0, col: 0 }, text: 'alpha-1' }], page: { incomplete: ['1'] } } })
    const out = await pending
    if (out.kind === 'complete') expect(out.result.page?.totalKnown).toBeUndefined()
    release()
    expect(await imported.done).toMatchObject({ kind: 'complete' })
    const complete = await h.client.send({ kind: 'search', requestId: h.client.nextRequestId(), owner: h.session.workbookOwner(), query: { text: 'alpha', caseSensitive: false } })
    expect(complete).toMatchObject({ kind: 'complete', result: { page: { incomplete: [], totalKnown: 3 } } })
  } finally { release(); await imported.done; await h.remote.release(); vi.useRealTimers() }
})


it('includes bounded offscreen text dependencies in worker viewport plans', async () => {
  const h = await harness(workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'alpha overflowing toward the viewport' }, { row: 0, col: 8, value: 'right-aligned dependency' }] }]))
  try {
    const out = await h.remote.viewport(initialIdentity('doc', 's'), { box: { x: 256, y: 0, width: 64, height: 20 }, zoom: 1, deviceScale: 1 }, new AbortController().signal, 'tile')
    expect(out).toMatchObject({ kind: 'complete', result: { content: { kind: 'detail' } } })
    if (typeof out === 'object' && out.kind === 'complete' && out.result.content.kind === 'detail') expect(out.result.content.cells.map(cell => cell.source.col)).toEqual([0, 8])
  } finally { await h.remote.release() }
})


it('paints worker-planned merged cells and overflowing text identically through canonical host layout', async () => {
  const bytes = xlsxBytes({ sheetData: '<row r="1"><c r="A1" t="inlineStr"><is><t>alpha merged</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>long alpha overflowing from outside the viewport</t></is></c></row>', afterSheetData: '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>' })
  const doc = await parseXlsx(await OfficePackage.load(bytes))
  const h = await harness(doc), plans = new WorkerViewportPlans(h.remote)
  const window = { box: { x: 64, y: 0, width: 192, height: 80 }, zoom: 1, deviceScale: 1 }
  const geometry = buildSheetGeometry(doc.sheets[0])
  expect(geometry.kind).toBe('ready')
  if (geometry.kind !== 'ready') return
  const extent = { width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() }
  try {
    h.remote.merges('1')
    await h.port.idle()
    expect(plans.get(h.session.identity('1'), window, extent, true).kind).toBe('pending')
    await h.port.idle()
    for (let i = 0; i < 30; i++) await Promise.resolve()
    // Each poll is a repaint, i.e. a new frame with its own page allowance (S13.3 r4).
    await vi.waitFor(() => { plans.beginFrame(h.session.identity('1'), window); expect(plans.get(h.session.identity('1'), window, extent, true).kind).toBe('detail') })
    const plan = plans.get(h.session.identity('1'), window, extent, true)
    if (plan.kind === 'detail') expect(plan.merges).toHaveLength(1)
    const a = createCanvas(192, 80), b = createCanvas(192, 80)
    const resident = renderSheetViewport(doc.sheets[0], a.getContext('2d') as unknown as CanvasRenderingContext2D, geometry, window)
    const worker = renderSheetViewport(doc.sheets[0], b.getContext('2d') as unknown as CanvasRenderingContext2D, undefined, window, undefined, { preparedPlan: plan })
    expect(worker.records).toEqual(resident.records)
    expect([...b.getContext('2d').getImageData(0, 0, 192, 80).data]).toEqual([...a.getContext('2d').getImageData(0, 0, 192, 80).data])
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(0)
    plans.dispose()
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
  } finally { plans.dispose(); await h.remote.release() }
})


it('bounds production tile requests and retains actual stalled producers until true drain after disposal', async () => {
  const h = await harness(), plans = new WorkerViewportPlans(h.remote)
  const source = h.source.store('s'), actualRows = source.readRows.bind(source)
  let began!: () => void, drain!: () => void, held = 0
  const started = new Promise<void>(resolve => { began = resolve })
  const gate = new Promise<void>(resolve => { drain = resolve })
  const reads = vi.spyOn(source, 'readRows').mockImplementation(async (...args) => {
    const page = await actualRows(...args)
    if (++held === 4) began()
    await gate
    return page
  })
  try {
    for (let i = 0; i < 10; i++) plans.get(initialIdentity('doc', 's'), { box: { x: i * 512, y: 0, width: 512, height: 512 }, zoom: 1, deviceScale: 1 }, { width: 1_048_576, height: 20_971_520 }, true)
    await started
    expect(h.remote.stats.commands.viewport).toBe(4)
    expect(reads).toHaveBeenCalledTimes(4)
    plans.dispose()
    await h.port.idle()
    expect(h.kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 4 * 2 * 8 * 1024 * 1024 })
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    drain()
    await Promise.all(reads.mock.results.map(result => result.value))
    await h.port.idle()
    expect(h.kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
    expect(h.remote.pagesInFlight).toBe(0)
    expect(h.remote.mainLedger.snapshot().byPool.layout).toBe(0)
    expect(h.remote.stats.commands.viewport).toBe(4)
  } finally {
    drain()
    await Promise.all(reads.mock.results.map(result => result.value))
    await h.port.idle()
    reads.mockRestore()
    plans.dispose()
    await h.remote.release()
  }
})

it('adopts known viewport cells without claiming unknown remainder columns are blank', async () => {
  const h = await harness(), store = h.remote.store('s')
  try {
    const cell = await h.source.store('s').readCell({ row: 0, col: 0 }, { signal: new AbortController().signal, stamp: store.stamp })
    expect(cell.kind).toBe('cell')
    if (cell.kind !== 'cell') return
    store.adoptViewport([{ ...cell, source: { row: 0, col: 0 }, bounds: { x: 0, y: 0, width: 64, height: 20 } }])
    expect(store.cellAt({ row: 0, col: 0 })?.value).toBe('alpha-0')
    const before = h.remote.stats.commands['read:rows'] ?? 0
    await store.load([0], 1, 1)
    expect(h.remote.stats.commands['read:rows']).toBe(before + 1)
  } finally { await h.remote.release() }
})


it('cancels held viewport tiles that scrolling moved out of the current frame', async () => {
  const h = await harness(), plans = new WorkerViewportPlans(h.remote), signals: AbortSignal[] = []
  let began!: () => void
  const started = new Promise<void>(r => { began = r })
  vi.spyOn(h.source.store('s'), 'readRows').mockImplementation((_r, _c, options) => { signals.push(options.signal); if (signals.length === 4) began(); return new Promise(() => {}) })
  const owner = initialIdentity('doc', 's'), extent = { width: 1_048_576, height: 20_971_520 }
  try {
    for (let i = 0; i < 4; i++) plans.get(owner, { box: { x: i * 512, y: 0, width: 512, height: 512 }, zoom: 1, deviceScale: 1 }, extent, true)
    await started
    plans.beginFrame(owner, { box: { x: 4096, y: 0, width: 512, height: 512 }, zoom: 1, deviceScale: 1 })
    for (let i = 0; i < 60; i++) await Promise.resolve()
    expect(signals.every(signal => signal.aborted), 'A current viewport must not wait behind four invisible stalled tiles').toBe(true)
    await h.port.idle()
    plans.get(owner, { box: { x: 4096, y: 0, width: 512, height: 512 }, zoom: 1, deviceScale: 1 }, extent, true)
    expect(h.remote.stats.commands.viewport).toBe(5)
  } finally { plans.dispose(); await h.port.idle(); await h.remote.release() }
})


it('keeps loaded viewport cells paintable beside explicit unknown rows', async () => {
  vi.useFakeTimers()
  const h = await harness(), plans = new WorkerViewportPlans(h.remote), store = h.source.store('s')
  const original = store.readRows.bind(store)
  vi.spyOn(store, 'readRows').mockImplementation(async (r, c, o) => { const page = await original(r, c, o); return { ...page, items: page.items.slice(0, 1), completeness: 'loading' } })
  const owner = initialIdentity('doc', 's'), window = { box: { x: 0, y: 0, width: 64, height: 40 }, zoom: 1, deviceScale: 1 }, extent = { width: 1_048_576, height: 20_971_520 }
  try {
    plans.get(owner, window, extent, true)
    await vi.advanceTimersByTimeAsync(0)
    expect(plans.get(owner, window, extent, true)).toMatchObject({ kind: 'detail', cells: [{ cell: { value: 'alpha-0' } }], pending: [{ bounds: { y: 20 } }] })
    const plan = plans.get(owner, window, extent, true), canvas = createCanvas(64, 40)
    const painted = renderSheetViewport(h.doc.sheets[0], canvas.getContext('2d') as unknown as CanvasRenderingContext2D, undefined, window, undefined, { preparedPlan: plan })
    expect(painted).toMatchObject({ pending: true, records: [{ body: { formattedText: 'alpha-0' } }] })
    expect([...canvas.getContext('2d').getImageData(2, 35, 1, 1).data], 'Unknown rows must be shaded rather than asserted blank').toEqual([236, 239, 243, 255])
  } finally { plans.dispose(); await h.remote.release(); vi.useRealTimers() }
})

it('stops automatic viewport retries after a nonretryable source failure', async () => {
  vi.useFakeTimers()
  const h = await harness(), plans = new WorkerViewportPlans(h.remote), store = h.source.store('s')
  vi.spyOn(store, 'readRows').mockResolvedValue({ items: [], completeness: 'incomplete', stamp: store.stamp, error: { code: 'parse-error', retryable: false, message: 'Import failed' } })
  const owner = initialIdentity('doc', 's'), window = { box: { x: 0, y: 0, width: 64, height: 40 }, zoom: 1, deviceScale: 1 }, extent = { width: 1_048_576, height: 20_971_520 }
  const unsubscribe = h.remote.subscribe(() => plans.get(owner, window, extent, true))
  try {
    plans.get(owner, window, extent, true)
    await vi.advanceTimersByTimeAsync(0)
    const before = h.remote.stats.commands.viewport
    await vi.advanceTimersByTimeAsync(1000)
    expect(h.remote.stats.commands.viewport, 'A terminal failure must not create a new 100ms query loop').toBe(before)
  } finally { unsubscribe(); plans.dispose(); await h.remote.release(); vi.useRealTimers() }
})
