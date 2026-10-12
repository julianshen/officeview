// @vitest-environment node
import { afterEach, expect, it, vi } from 'vitest'
import { FOREGROUND_YIELD_WAIT_MS, QUERY_DEADLINE_MS, ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { initialIdentity } from '../src/worker/viewerProtocol'
import type { WorkerReply } from '../src/xlsx/view/contracts'
import { workbookOf } from './helpers/xlsx-find'

afterEach(() => vi.useRealTimers())

it('times out a never-resolving preparation read at 120 seconds and releases its operation charge', async () => {
  vi.useFakeTimers()
  const document = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 5 }] }])
  const source = createObjectWorkbookSource(document, 'doc')
  let readSignal: AbortSignal | undefined
  vi.spyOn(source.store('s'), 'queryColumn').mockImplementation((_query, options) => {
    readSignal = options.signal
    return new Promise(() => {})
  })
  const kernel = new ViewerKernel({ parse: async () => document, openSource: async () => source }, { yield: async () => {} })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const replies: WorkerReply[] = []
  const pending = kernel.handle({ kind: 'prepareColumns', requestId: 'stalled', owner: initialIdentity('doc', 's'), cols: Uint16Array.of(0) }, out => replies.push(out))
  for (let i = 0; i < 30; i++) await Promise.resolve()
  expect(readSignal).toBeDefined()
  expect(kernel.lifecycleStats().transientBytes).toBeGreaterThan(0)
  try {
    await vi.advanceTimersByTimeAsync(119_999)
    expect(replies.filter(reply => reply.kind !== 'progress')).toHaveLength(0)
    await vi.advanceTimersByTimeAsync(1)
    expect(replies.filter(reply => reply.kind !== 'progress'), 'A stalled query must emit one terminal timeout reply').toMatchObject([
      { kind: 'cancelled', diagnostics: [{ code: 'query-timeout', message: expect.stringMatching(/120 seconds/) }] },
    ])
    await pending
    expect(readSignal!.aborted).toBe(true)
    expect(kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
    expect(kernel.session('doc')!.isPrepared('s', 0)).toBe(false)
  } finally {
    await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {})
  }
})

it('keeps the deadline armed until a staged search reply is emitted and releases an unpublished cursor', async () => {
  vi.useFakeTimers()
  const document = workbookOf([{ sheetId: 's', cells: Array.from({ length: 450 }, (_, row) => ({ row, col: 0, value: 'hit' })) }])
  let began!: () => void, release!: () => void
  const started = new Promise<void>(resolve => { began = resolve })
  const held = new Promise<void>(resolve => { release = resolve })
  const kernel = new ViewerKernel({ parse: async () => document, afterDispatch: async command => { if (command.kind === 'search') { began(); await held } } }, { yield: async () => {} })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const replies: WorkerReply[] = []
  const pending = kernel.handle({ kind: 'search', requestId: 'held', owner: kernel.session('doc')!.workbookOwner(), query: { text: 'hit', caseSensitive: false } }, out => replies.push(out))
  try {
    await started
    expect(kernel.lifecycleStats().transientBytes).toBeGreaterThan(0)
    await vi.advanceTimersByTimeAsync(120_001)
    expect(replies, 'A staged but unpublished query must still time out').toMatchObject([{ kind: 'cancelled', diagnostics: [{ code: 'query-timeout' }] }])
    await pending
    expect(kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
    release()
    for (let i = 0; i < 30; i++) await Promise.resolve()
    expect(replies).toHaveLength(1)
  } finally { release(); await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}); await pending }
})

it('suppresses a late page after timeout and accepts another request without waiting for that provider', async () => {
  vi.useFakeTimers()
  const document = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 5 }] }])
  const source = createObjectWorkbookSource(document, 'doc'), store = source.store('s')
  const original = store.queryColumn.bind(store)
  let resolve!: () => void
  const blocked = new Promise<void>(done => { resolve = done })
  const spy = vi.spyOn(store, 'queryColumn').mockImplementationOnce(async (query, options) => {
    await blocked
    return original(query, { ...options, signal: new AbortController().signal })
  })
  const kernel = new ViewerKernel({ parse: async () => document, openSource: async () => source }, { yield: async () => {} })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const replies: WorkerReply[] = [], owner = initialIdentity('doc', 's')
  const pending = kernel.handle({ kind: 'prepareColumns', requestId: 'first', owner, cols: Uint16Array.of(0) }, out => replies.push(out))
  try {
    await vi.advanceTimersByTimeAsync(120_001)
    expect(replies.filter(out => out.kind !== 'progress')).toMatchObject([{ kind: 'cancelled' }])
    await pending
    const again: WorkerReply[] = []
    await kernel.handle({ kind: 'prepareColumns', requestId: 'second', owner, cols: Uint16Array.of(0) }, out => again.push(out))
    expect(again.at(-1)).toMatchObject({ kind: 'complete' })
    resolve()
    for (let i = 0; i < 30; i++) await Promise.resolve()
    expect(replies.filter(out => out.kind !== 'progress')).toHaveLength(1)
    expect(kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
  } finally { resolve(); spy.mockRestore(); await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
})

it.each(['evaluator', 'yield'] as const)('reclaims session leases when the timed-out %s ignores cancellation', async blocked => {
  vi.useFakeTimers()
  const document = workbookOf([
    { sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] },
    { sheetId: 'other', cells: Array.from({ length: blocked === 'yield' ? 5000 : 1 }, (_, row) => ({ row, col: 0, value: row })) },
  ])
  let began!: () => void
  const started = new Promise<void>(resolve => { began = resolve })
  const kernel = new ViewerKernel({ parse: async () => document }, {
    yield: async () => { if (blocked === 'yield') { began(); await new Promise(() => {}) } },
    filterEvaluator: async () => { began(); return new Promise(() => {}) },
  })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const owner = initialIdentity('doc', 'other'), replies: WorkerReply[] = []
  const pending = blocked === 'yield'
    ? kernel.handle({ kind: 'prepareColumns', requestId: 'blocked', owner, cols: Uint16Array.of(0) }, out => replies.push(out))
    : kernel.handle({ kind: 'filter', requestId: 'blocked', owner, nowMs: 0, spec: { range: { r0: 0, r1: 0, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }, keySets: [{ refId: 'keys', scope: 'domain', run: { count: 1, kinds: Uint8Array.of(1), numbers: Float64Array.of(1), texts: [undefined] } }] }, out => replies.push(out))
  try {
    await started
    if (blocked === 'evaluator') expect(kernel.lifecycleStats().transientBytes, 'Shipped filter keys must be charged while their evaluator is stalled').toBeGreaterThan(0)
    await vi.advanceTimersByTimeAsync(120_001)
    await pending
    expect(replies.at(-1)).toMatchObject({ kind: 'cancelled', diagnostics: [{ code: 'query-timeout' }] })
    expect(kernel.lifecycleStats().transientBytes, 'A timeout must unwind operation charges and shipped-key leases').toBe(0)
    kernel.session('doc')!.memory.trim(0)
    expect(kernel.session('doc')!.isResident('other'), 'A timed-out evaluator must unpin its inactive view').toBe(false)
  } finally { await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
})


it('cancels a signal-ignoring read promptly without relabeling it as a timeout', async () => {
  vi.useFakeTimers()
  const document = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 5 }] }])
  const source = createObjectWorkbookSource(document, 'doc')
  vi.spyOn(source.store('s'), 'queryColumn').mockImplementation(() => new Promise(() => {}))
  const kernel = new ViewerKernel({ parse: async () => document, openSource: async () => source }, { yield: async () => {} })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const replies: WorkerReply[] = []
  const pending = kernel.handle({ kind: 'prepareColumns', requestId: 'held', owner: initialIdentity('doc', 's'), cols: Uint16Array.of(0) }, out => replies.push(out))
  try {
    for (let i = 0; i < 30; i++) await Promise.resolve()
    expect(kernel.lifecycleStats().transientBytes).toBeGreaterThan(0)
    await kernel.handle({ kind: 'cancel', requestId: 'cancel', targetRequestId: 'held' }, () => {})
    for (let i = 0; i < 60; i++) await Promise.resolve()
    expect(replies.at(-1)).toMatchObject({ kind: 'cancelled', diagnostics: [{ code: 'cancel-requested' }] })
    expect(kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
    await vi.advanceTimersByTimeAsync(120_000)
    expect(replies.filter(r => r.kind !== 'progress')).toHaveLength(1)
    expect(replies.at(-1)?.diagnostics.some(d => d.code === 'query-timeout')).toBe(false)
    await pending
  } finally { await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
})


it('gives a current viewport priority over a background search at its next cooperative yield', async () => {
  const document = workbookOf([{ sheetId: 's', cells: Array.from({ length: 4500 }, (_, row) => ({ row, col: 0, value: 'alpha' })) }])
  const source = createObjectWorkbookSource(document, 'doc')
  let resume!: () => void, began!: () => void, releaseViewport!: () => void, viewportBegan!: () => void
  const yielded = new Promise<void>(r => { began = r }), held = new Promise<void>(r => { resume = r })
  const viewportStarted = new Promise<void>(r => { viewportBegan = r }), viewportHeld = new Promise<void>(r => { releaseViewport = r })
  const original = source.store('s').readRows.bind(source.store('s'))
  vi.spyOn(source.store('s'), 'readRows').mockImplementationOnce(async (q, c, o) => { viewportBegan(); await viewportHeld; return original(q, c, o) })
  const formatted = vi.spyOn(source, 'displayText')
  let first = true, ticks = 0
  const kernel = new ViewerKernel({ parse: async () => document, openSource: async () => source }, { now: () => ticks += 11, yield: async phase => { if (phase === 'find-scan' && first) { first = false; began(); await held } } })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const searchReplies: WorkerReply[] = [], viewportReplies: WorkerReply[] = []
  const search = kernel.handle({ kind: 'search', requestId: 'search', owner: kernel.session('doc')!.workbookOwner(), query: { text: 'alpha', caseSensitive: false } }, r => searchReplies.push(r))
  await yielded
  const viewport = kernel.handle({ kind: 'viewport', requestId: 'viewport', owner: initialIdentity('doc', 's'), window: { box: { x: 0, y: 0, width: 64, height: 100 }, zoom: 1, deviceScale: 1 } }, r => viewportReplies.push(r))
  try {
    await viewportStarted
    const before = formatted.mock.calls.length
    resume()
    for (let i = 0; i < 200; i++) await Promise.resolve()
    expect(formatted.mock.calls.length, 'Background work cannot resume past its yield while viewport is pending').toBe(before)
    expect(searchReplies, 'Background search must remain paused while the foreground viewport is pending').toHaveLength(0)
    releaseViewport()
    await viewport
    expect(viewportReplies.at(-1)).toMatchObject({ kind: 'complete' })
    await search
    expect(searchReplies.at(-1)).toMatchObject({ kind: 'complete' })
  } finally { resume(); releaseViewport(); await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}); await search; await viewport }
})


it('cancels a retained search continuation while foreground work is pending', async () => {
  const document = workbookOf([{ sheetId: 's', cells: Array.from({ length: 4500 }, (_, row) => ({ row, col: 0, value: 'hit' })) }])
  const source = createObjectWorkbookSource(document, 'doc')
  let continuation = false, ticks = 0, began!: () => void, viewportBegan!: () => void, release!: () => void
  const started = new Promise<void>(r => { began = r }), viewportStarted = new Promise<void>(r => { viewportBegan = r }), held = new Promise<void>(r => { release = r })
  const kernel = new ViewerKernel({ parse: async () => document, openSource: async () => source }, { now: () => ticks += 11, yield: async () => { if (continuation) began() } })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const owner = kernel.session('doc')!.workbookOwner(), query = { text: 'hit', caseSensitive: false }
  let cursor: string | undefined
  for (let i = 0; i < 20; i++) {
    const replies: WorkerReply[] = []
    await kernel.handle({ kind: 'search', requestId: `page-${i}`, owner, query, ...(cursor ? { cursor } : {}) }, r => replies.push(r))
    const out = replies.at(-1)!
    expect(out.kind).toBe('complete')
    if (out.kind !== 'complete' || out.result.kind !== 'search') throw new Error('Search failed')
    cursor = out.result.nextCursor
  }
  const searchBytes = kernel.lifecycleStats().transientBytes
  const original = source.store('s').readRows.bind(source.store('s'))
  vi.spyOn(source.store('s'), 'readRows').mockImplementationOnce(async (q, c, o) => { viewportBegan(); await held; return original(q, c, o) })
  const viewport = kernel.handle({ kind: 'viewport', requestId: 'viewport', owner: initialIdentity('doc', 's'), window: { box: { x: 0, y: 0, width: 64, height: 100 }, zoom: 1, deviceScale: 1 } }, () => {})
  await viewportStarted
  const viewportBytes = kernel.lifecycleStats().transientBytes - searchBytes
  continuation = true
  const replies: WorkerReply[] = []
  const pending = kernel.handle({ kind: 'search', requestId: 'next', owner, query, cursor }, r => replies.push(r))
  try {
    await started
    for (let i = 0; i < 60; i++) await Promise.resolve()
    await kernel.handle({ kind: 'cancel', requestId: 'cancel', targetRequestId: 'next' }, () => {})
    for (let i = 0; i < 100; i++) await Promise.resolve()
    expect(replies.at(-1)).toMatchObject({ kind: 'cancelled', diagnostics: [{ code: 'cancel-requested' }] })
    await pending
    expect(kernel.lifecycleStats()).toMatchObject({ inflight: 1, transientBytes: viewportBytes })
    release()
    await viewport
    expect(kernel.lifecycleStats()).toMatchObject({ inflight: 0, transientBytes: 0 })
  } finally { release(); await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}); await pending; await viewport }
})

it('S13.3 MINOR-1: a sort completes while overlapping viewport commands keep arriving', async () => {
  const rows = 4500
  const document = workbookOf([{ sheetId: 's', cells: Array.from({ length: rows }, (_, row) => ({ row, col: 0, value: (row * 7919) % rows })) }])
  let ticks = 0, yields = 0
  // A fast virtual clock: the sort yields many times on the way.
  const kernel = new ViewerKernel({ parse: async () => document }, { now: () => ticks += 3, yield: async () => { yields++ } })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const session = kernel.session('doc')!
  const viewport = session.viewport.bind(session)
  // Each viewport takes a few milliseconds; a new one starts before the last
  // ends, so foreground work never drains while "scrolling" continues.
  vi.spyOn(session, 'viewport').mockImplementation(async (...args) => { await new Promise(resolve => setTimeout(resolve, 8)); return viewport(...args) })
  const owner = initialIdentity('doc', 's')
  let scrolling = true, open = 0, overlapped = 0, serial = 0
  const scroll = (async () => {
    for (let i = 0; scrolling && i < 1500; i++) {
      open++
      void kernel.handle({ kind: 'viewport', requestId: `v${serial++}`, owner, window: { box: { x: 0, y: (i % 50) * 20, width: 64, height: 100 }, zoom: 1, deviceScale: 1 } }, out => { if (out.kind !== 'progress') open-- })
      await new Promise(resolve => setTimeout(resolve, 4))
    }
  })()
  const replies: WorkerReply[] = []
  try {
    await kernel.handle({ kind: 'sort', requestId: 'sort', owner, spec: { range: { r0: 0, r1: rows - 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false } }, out => {
      if (out.kind !== 'progress' && scrolling && open > 0) overlapped++
      replies.push(out)
    })
    expect(yields, 'the sort yielded to foreground work repeatedly').toBeGreaterThan(5)
    expect(replies.at(-1)).toMatchObject({ kind: 'complete' })
    expect(overlapped, 'the sort finished while viewport commands were still in flight').toBe(1)
  } finally {
    scrolling = false
    await scroll
    await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {})
  }
}, 20_000)

it('S13.3 MINOR-1: a slow foreground command parks background work for a bounded share per yield only', async () => {
  vi.useFakeTimers()
  const rows = 2000
  const document = workbookOf([{ sheetId: 's', cells: Array.from({ length: rows }, (_, row) => ({ row, col: 0, value: (row * 7919) % rows })) }])
  let ticks = 0, yields = 0
  const kernel = new ViewerKernel({ parse: async () => document }, { now: () => ticks += 3, yield: async () => { yields++ } })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  const session = kernel.session('doc')!
  let release!: () => void
  const held = new Promise<void>(resolve => { release = resolve })
  const viewport = session.viewport.bind(session)
  vi.spyOn(session, 'viewport').mockImplementation(async (...args) => { await held; return viewport(...args) })
  const owner = initialIdentity('doc', 's')
  const slow = kernel.handle({ kind: 'viewport', requestId: 'slow', owner, window: { box: { x: 0, y: 0, width: 64, height: 100 }, zoom: 1, deviceScale: 1 } }, () => {})
  const replies: WorkerReply[] = []
  let done = false
  const sort = kernel.handle({ kind: 'sort', requestId: 'sort', owner, spec: { range: { r0: 0, r1: rows - 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false } }, out => replies.push(out)).then(() => { done = true })
  try {
    let waited = 0
    while (!done && waited < QUERY_DEADLINE_MS / 2) { await vi.advanceTimersByTimeAsync(FOREGROUND_YIELD_WAIT_MS); waited += FOREGROUND_YIELD_WAIT_MS }
    expect(done, 'the sort progresses although the foreground command never finishes').toBe(true)
    expect(yields).toBeGreaterThan(1)
    expect(replies.at(-1)).toMatchObject({ kind: 'complete' })
    expect(waited, 'each yield waited at most one bounded share').toBeLessThanOrEqual(yields * FOREGROUND_YIELD_WAIT_MS)
  } finally {
    release()
    await vi.advanceTimersByTimeAsync(0)
    await slow; await sort
    await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {})
  }
})
