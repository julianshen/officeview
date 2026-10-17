// @vitest-environment node
// FC-T2: a drawings read carries a receiving grant main reserved before it was
// sent; the worker never posts a reply over that grant.
import { describe, expect, it, vi } from 'vitest'
import { ViewerKernel, connectInProcess, QUERY_DEADLINE_MS } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { drawingOwnerMetadataBytes } from '../src/xlsx/view/drawingOwners'
import { DRAWING_MEMORY_REFUSED, drawingPayloadBudget } from '../src/xlsx/view/drawingPayload'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { initialIdentity, transferablesOf, type CommandOf } from '../src/worker/viewerProtocol'
import { replyLimitEnvelope } from '../src/worker/viewerReplyLimit'
import { estimatePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { estimatePayloadBytes as ledgerEstimate } from '../src/xlsx/view/mainLedger'
import type { WorkerReply } from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'
import { workbookOf } from './helpers/xlsx-find'

const E = 9525
/** A sheet with one grouped chart and, optionally, one picture of `imageBytes`. */
function drawingDoc(sheetIds: readonly string[] = ['s'], imageBytes = 0): XlsxDocument {
  const doc = workbookOf(sheetIds.map(sheetId => ({ sheetId, cells: [{ row: 0, col: 0, value: 42 }] })))
  for (const sheet of doc.sheets) {
    sheet.columnMetrics = { maxDigitWidthPx: 7, defaultWidthPx: 80 }
    sheet.drawings = [{
      xEmu: 80 * E, yEmu: 0, widthEmu: 80 * E, heightEmu: 20 * E,
      source: { partPath: 'd', treePath: 'twoCellAnchor[0]/graphicFrame[0]', element: 'graphicFrame', representation: 'native' },
      content: { kind: 'xlsxChart', chart: { groups: [{ series: [{ values: Array.from({ length: 1000 }, (_, i) => i + 42) }] }] } } as never,
      ...(imageBytes ? { image: { data: new Uint8Array(imageBytes), mime: 'image/png', partPath: 'p.png' } } : {}),
    }]
    sheet.drawingAnchors = { unresolved: 0, records: [{ id: 'd', kind: 'twoCell', from: { col: 1, row: 0, colOffEmu: 0, rowOffEmu: 0 }, to: { col: 2, row: 1, colOffEmu: 0, rowOffEmu: 0 } }] }
  }
  return doc
}
async function kernelFixture(imageBytes = 4096) {
  const doc = drawingDoc(['s'], imageBytes), source = createObjectWorkbookSource(doc, 'k')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  await kernel.handle({ kind: 'open', requestId: 'open', docId: 'k', source: { kind: 'object', document: doc }, profile: 'D' }, () => {})
  // The raw sink: what the kernel would hand to postMessage / structuredClone.
  const read = async (requestId: string, replyByteLimit?: number) => {
    const posted: WorkerReply[] = []
    const command: CommandOf<'read'> = { kind: 'read', requestId, owner: initialIdentity('k', 's'), query: { op: 'drawings', ...(replyByteLimit === undefined ? {} : { replyByteLimit }) } }
    await kernel.handle(command, out => posted.push(out))
    return { posted, command }
  }
  return { kernel, source, read }
}
const drawingsOf = (out: WorkerReply) => out.kind === 'complete' && out.result.kind === 'read' && out.result.result.op === 'drawings' ? out.result.result.result : undefined

describe('shared reply measure', () => {
  it('is equal on the sender object and on its transferred structured clone, counting whole backing buffers once', () => {
    const padded = new Uint8Array(64 * 1024).subarray(100, 200)
    const shared = new ArrayBuffer(8192)
    const out = {
      kind: 'complete', requestId: 'r-1', commandKind: 'read', owner: initialIdentity('d', 's'),
      result: { kind: 'read', result: { op: 'drawings', result: {
        completeness: 'complete', bytes: 0, diagnostics: [], drawings: [{ image: { data: padded } }, { a: new Uint8Array(shared, 0, 16) }, { b: new Float64Array(shared, 8, 4) }],
        images: [], theme: new Map([['accent1', 'ff0000']]),
      } } },
      // The kernel appends cached-value diagnostics to the final reply.
      diagnostics: [{ code: 'cached-value-policy', message: 'Values are the cached values saved with the file' }],
    }
    const before = estimatePayloadBytes(out)
    expect(before).toBeGreaterThanOrEqual(64 * 1024 + 8192)
    const clone = structuredClone(out, { transfer: transferablesOf(out) })
    expect(estimatePayloadBytes(clone)).toBe(before)
    expect(ledgerEstimate).toBe(estimatePayloadBytes)
  })
})

describe('kernel reply grant', () => {
  it('posts a final reply that measures exactly the grant and replaces one byte less with a payload-free quota reply', async () => {
    const { read } = await kernelFixture()
    const { posted: [full] } = await read('d0')
    const D = estimatePayloadBytes(full)
    expect(drawingsOf(full)?.completeness).toBe('complete')

    const exact = await read('d1', D)
    expect(exact.posted).toHaveLength(1)
    expect(drawingsOf(exact.posted[0])?.completeness).toBe('complete')
    expect(estimatePayloadBytes(exact.posted[0])).toBe(D)

    const under = await read('d2', D - 1)
    expect(under.posted).toHaveLength(1)
    expect(under.posted[0]).toMatchObject({ kind: 'quota', requestId: 'd2', resource: 'memory', requiredBytes: D, limitBytes: D - 1, diagnostics: [] })
    expect(transferablesOf(under.posted[0])).toEqual([])
    expect(estimatePayloadBytes(under.posted[0])).toBeLessThanOrEqual(replyLimitEnvelope(under.command))
  })

  it('checks the final reply, after the cached-value diagnostics were added', async () => {
    const { source, read } = await kernelFixture()
    const { posted: [plain] } = await read('d0')
    const D = estimatePayloadBytes(plain)
    vi.spyOn(source as unknown as { valueStatus(): unknown }, 'valueStatus').mockReturnValue({ policy: 'cached', message: 'Cached values are shown' })
    const { posted: [diagnosed] } = await read('d1')
    expect(diagnosed.diagnostics.some(d => d.code === 'cached-value-policy')).toBe(true)
    const final = estimatePayloadBytes(diagnosed)
    expect(final).toBeGreaterThan(D)
    // The provider's payload alone fits D; the diagnosed reply does not.
    const { posted: [out] } = await read('d2', final - 1)
    expect(out).toMatchObject({ kind: 'quota', requiredBytes: final })
  })

  it('refuses a grant below the payload-free envelope before it invokes the provider', async () => {
    const { source, read } = await kernelFixture()
    const provider = vi.spyOn(source.store('s'), 'readDrawings')
    const { posted, command } = await read('d0', 10)
    expect(provider).not.toHaveBeenCalled()
    expect(posted).toMatchObject([{ kind: 'quota', requestId: 'd0', diagnostics: [] }])
    const envelope = replyLimitEnvelope(command)
    await read('d1', envelope)
    expect(provider).toHaveBeenCalledTimes(1)
  })
})

/** A worker-backed workbook whose drawing provider can be held per call and
 * whose reply delivery can be queued (a reply posted but not yet consumed). */
async function remoteFixture(options: { sheets?: string[]; imageBytes?: number; empty?: boolean; hold?: boolean; queue?: boolean } = {}) {
  const sheets = options.sheets ?? ['s']
  const doc = drawingDoc(sheets, options.imageBytes ?? 0)
  if (options.empty) for (const sheet of doc.sheets) { sheet.drawings = []; sheet.drawingAnchors = { unresolved: 0, records: [] } }
  const source = createObjectWorkbookSource(doc, 'r')
  const gates: (() => void)[] = []
  if (options.hold) for (const id of sheets) {
    const store = source.store(id), read = store.readDrawings!.bind(store)
    // Ignores its signal: resumes only when the test opens the gate.
    store.readDrawings = (...args) => new Promise<void>(resolve => gates.push(resolve)).then(() => read(...args))
  }
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const inner = connectInProcess(kernel)
  const queued: { data: unknown; deliver: () => void }[] = []
  let delay = false
  const sent: unknown[] = []
  const port = {
    ...inner,
    postMessage(message: unknown, transfer: Transferable[]) { sent.push(message); inner.postMessage(message, transfer) },
    addEventListener(type: 'message', listener: (event: { data: unknown }) => void) {
      inner.addEventListener(type, event => { if (delay) queued.push({ data: event.data, deliver: () => listener(event) }); else listener(event) })
    },
  }
  const client = new ViewerClient(port as never)
  expect((await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'r', source: { kind: 'object', document: doc }, profile: 'D' })).kind).toBe('complete')
  const session = kernel.session('r')!
  const remote = new RemoteWorkbook(client, { docId: 'r', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(sheets.map(id => [id, session.layout(id)])) })
  for (const id of sheets) {
    remote.merges(id)
    for (let i = 0; i < 200 && remote.loadState(id).completeness !== 'complete'; i++) await new Promise(r => setTimeout(r, 2))
  }
  const ledger = remote.mainLedger, cap = MAIN_THREAD_BUDGETS.D
  const command = (requestId: string, sheetId = sheets[0]): CommandOf<'read'> => ({ kind: 'read', requestId, owner: remote.identity(sheetId), query: { op: 'drawings' } })
  const drawingsSent = () => sent.filter(m => (m as { kind?: string; query?: { op?: string } }).query?.op === 'drawings') as CommandOf<'read'>[]
  const settle = async () => { await inner.idle(); for (let i = 0; i < 20; i++) await Promise.resolve() }
  return {
    remote, client, kernel, source, ledger, cap, command, drawingsSent, settle, gates,
    pressure: (headroom: number) => ledger.set('dom', 'pressure', Math.max(0, cap - (ledger.bytes - ledger.ownerBytes('dom', 'pressure')) - headroom)),
    queue: (on: boolean) => { delay = on; if (!on) for (const { deliver } of queued.splice(0)) deliver() },
    /** Delivers the queued replies `pick` selects, keeping the rest queued. */
    deliver: (pick: (data: { commandKind?: string; requestId?: string }) => boolean) => {
      for (const item of [...queued]) if (pick(item.data as never)) { queued.splice(queued.indexOf(item), 1); item.deliver() }
    },
    grant: (requestId: string) => ledger.ownerBytes('replies', `drawing-grant:${requestId}`),
  }
}
const until = async (check: () => boolean, ms = 4000) => { const end = Date.now() + ms; while (!check() && Date.now() < end) await new Promise(r => setTimeout(r, 5)) }

describe('main receiving grant', () => {
  it('reserves before dispatch, bounded by the drawing budget, and returns the unused rest before adopting', async () => {
    const f = await remoteFixture({ hold: true })
    try {
      const replies = f.ledger.snapshot().byPool.replies
      const pending = f.remote.request(f.command('a-1'))
      expect(f.drawingsSent()).toHaveLength(1)
      expect(f.drawingsSent()[0].query).toEqual({ op: 'drawings', replyByteLimit: drawingPayloadBudget('D') })
      expect(f.grant('a-1')).toBe(drawingPayloadBudget('D'))
      await until(() => f.gates.length > 0)
      f.gates.shift()!()
      const out = await pending
      expect(typeof out === 'object' && out.kind).toBe('complete')
      expect(f.grant('a-1')).toBe(0)
      expect(f.ledger.snapshot().byPool.replies).toBe(replies)
      // The reply's data moved from the grant; owner metadata was admitted after the refund.
      const payload = drawingsOf(out as WorkerReply)!
      expect(f.ledger.snapshot().byPool.drawings).toBe(estimatePayloadBytes(out) + drawingOwnerMetadataBytes(payload.drawings))
    } finally { await f.remote.close() }
  })

  it('sends nothing when the headroom is below the payload-free envelope, and the refusal is retryable', async () => {
    const f = await remoteFixture()
    try {
      const command = f.command('e-1')
      f.pressure(replyLimitEnvelope({ ...command, query: { op: 'drawings', replyByteLimit: 0 } }) - 1)
      const before = f.ledger.snapshot()
      const out = await f.remote.request(command)
      expect(f.drawingsSent()).toHaveLength(0)
      expect(out).toMatchObject({ kind: 'complete', result: { result: { result: { completeness: 'incomplete', error: { code: 'resource-limit', retryable: true }, diagnostics: [{ code: DRAWING_MEMORY_REFUSED }] } } } })
      expect(f.ledger.snapshot()).toEqual(before)
    } finally { await f.remote.close() }
  })

  it('gives two concurrent reads two credits and settles each one alone', async () => {
    const f = await remoteFixture({ hold: true })
    try {
      f.pressure(24 * 1024 * 1024)
      const first = f.remote.request(f.command('c-1'))
      const second = f.remote.request(f.command('c-2'))
      const budget = drawingPayloadBudget('D')
      expect(f.grant('c-1')).toBe(budget)
      expect(f.grant('c-2')).toBe(24 * 1024 * 1024 - budget)
      expect(f.drawingsSent().map(c => c.query)).toEqual([{ op: 'drawings', replyByteLimit: budget }, { op: 'drawings', replyByteLimit: 24 * 1024 * 1024 - budget }])
      expect(f.ledger.bytes).toBeLessThanOrEqual(f.cap)
      await until(() => f.gates.length === 2)
      f.gates.shift()!()
      await first
      expect(f.grant('c-1')).toBe(0)
      expect(f.grant('c-2')).toBe(24 * 1024 * 1024 - budget)
      f.gates.shift()!()
      await second
      expect(f.grant('c-2')).toBe(0)
      expect(f.ledger.peakBytes).toBeLessThanOrEqual(f.cap)
    } finally { await f.remote.close() }
  })

  it('never leaves concurrent sheet loads permanently failed when only one grant fits at a time', async () => {
    const f = await remoteFixture({ sheets: ['s', 't'] })
    try {
      // Room for one sheet's grant and owner workspace, not two at once.
      f.pressure(1024 * 1024)
      f.remote.drawings('s'); f.remote.drawings('t')
      expect(f.drawingsSent(), 'the second sheet found no credit left').toHaveLength(1)
      await until(() => f.remote.drawings('s').state === 'ready' && f.remote.drawings('t').state === 'ready')
      expect([f.remote.drawings('s').state, f.remote.drawings('t').state]).toEqual(['ready', 'ready'])
      expect(f.ledger.peakBytes).toBeLessThanOrEqual(f.cap)
    } finally { await f.remote.close() }
  })

  it('holds the credit after a public abort and the cancel ACK, until the target terminal is consumed', async () => {
    const f = await remoteFixture({ hold: true })
    try {
      const replies = f.ledger.snapshot().byPool.replies
      f.queue(true)
      const controller = new AbortController()
      const pending = f.remote.request(f.command('h-1'), controller.signal)
      const grant = f.grant('h-1')
      expect(grant).toBeGreaterThan(0)
      await until(() => f.gates.length > 0)
      controller.abort()
      expect(typeof await pending).toBe('symbol')
      await f.settle()
      f.deliver(data => data.commandKind === 'cancel')
      await f.settle()
      expect(f.grant('h-1'), 'cancel ACK and public abort do not refund').toBe(grant)
      f.deliver(data => data.requestId === 'h-1')
      await f.settle()
      expect(f.grant('h-1')).toBe(0)
      expect(f.ledger.snapshot().byPool.replies).toBe(replies)
    } finally { f.queue(false); f.gates.forEach(open => open()); await f.remote.close() }
  })

  it('adopts a full reply that was posted before the abort and consumed after it', async () => {
    const f = await remoteFixture({ imageBytes: 64 * 1024 })
    try {
      f.queue(true)
      const controller = new AbortController()
      const pending = f.remote.request(f.command('q-1'), controller.signal)
      await f.settle()
      controller.abort()
      expect(typeof await pending).toBe('symbol')
      const grant = f.grant('q-1')
      expect(grant).toBe(drawingPayloadBudget('D'))
      const peak = f.ledger.peakBytes
      f.queue(false)
      await f.settle()
      expect(f.grant('q-1')).toBe(0)
      expect(f.ledger.snapshot().byPool.drawings).toBeGreaterThan(64 * 1024)
      expect(f.ledger.peakBytes).toBe(peak)
    } finally { f.queue(false); await f.remote.close() }
  })

  it('refunds at once only when nothing was sent; a disposed client keeps the credit', async () => {
    const f = await remoteFixture({ hold: true })
    try {
      const original = f.client['port'].postMessage
      f.client['port'].postMessage = () => { throw new DOMException('cannot clone', 'DataCloneError') }
      expect(await f.remote.request(f.command('n-1'))).toBeTypeOf('symbol')
      expect(f.grant('n-1')).toBe(0)
      f.client['port'].postMessage = original
      const pending = f.remote.request(f.command('n-2'))
      expect(f.grant('n-2')).toBeGreaterThan(0)
      f.client.dispose()
      expect(await pending).toBeTypeOf('symbol')
      expect(f.grant('n-2'), 'listener detachment does not prove the reply cannot arrive').toBeGreaterThan(0)
    } finally { f.gates.forEach(open => open()); await f.remote.close() }
  })

  it('keeps an empty sheet ready under admissible pressure and refunds the whole grant', async () => {
    const f = await remoteFixture({ empty: true })
    try {
      f.pressure(64 * 1024)
      const before = f.ledger.snapshot()
      await until(() => f.remote.drawings('s').state !== 'waiting' && f.remote.drawings('s').state !== 'loading')
      expect(f.remote.drawings('s').state).toBe('ready')
      const after = f.ledger.snapshot()
      expect(after.byPool.replies).toBe(before.byPool.replies)
      expect(after.byPool.drawings - before.byPool.drawings).toBeLessThan(4096)
      expect(after.peakBytes).toBeLessThanOrEqual(f.cap)
    } finally { await f.remote.close() }
  })

  it('a main-cache warm hit sends and reserves nothing; a worker-cache hit still needs a grant', async () => {
    const f = await remoteFixture()
    try {
      await until(() => f.remote.drawings('s').state === 'ready')
      const sent = f.drawingsSent().length, peak = f.ledger.peakBytes, bytes = f.ledger.bytes
      expect(f.remote.drawings('s').state).toBe('ready')
      expect(f.drawingsSent()).toHaveLength(sent)
      expect([f.ledger.peakBytes, f.ledger.bytes]).toEqual([peak, bytes])
      const again = f.remote.request(f.command('w-1'))
      expect(f.grant('w-1')).toBeGreaterThan(0)
      await again
    } finally { await f.remote.close() }
  })

  it('asks once more with the measured size when a dense reply exceeds the initial grant', async () => {
    const f = await remoteFixture({ imageBytes: drawingPayloadBudget('D') })
    try {
      const replies = f.ledger.snapshot().byPool.replies
      const out = await f.remote.request(f.command('d-1'))
      expect(f.drawingsSent()).toHaveLength(2)
      const [first, second] = f.drawingsSent()
      expect(first.query).toEqual({ op: 'drawings', replyByteLimit: drawingPayloadBudget('D') })
      expect((second.query as { replyByteLimit: number }).replyByteLimit).toBeGreaterThan(drawingPayloadBudget('D'))
      expect(out).toMatchObject({ kind: 'complete', result: { result: { result: { completeness: 'complete' } } } })
      expect(f.ledger.snapshot().byPool.replies).toBe(replies)
      expect(f.ledger.peakBytes).toBeLessThanOrEqual(f.cap)
    } finally { await f.remote.close() }
  })
})

describe('worker lifetime of a signal-ignoring drawing provider under a grant', () => {
  it('times out with exactly one bounded terminal, never posts late, and keeps the worker share until the provider settles', async () => {
    vi.useFakeTimers()
    try {
      const { kernel, source } = await kernelFixture()
      let open!: () => void
      const held = new Promise<void>(resolve => { open = resolve })
      vi.spyOn(source.store('s'), 'readDrawings').mockImplementation(async (_options, _profile, memory) => {
        memory?.set(4096)
        await held
        return { completeness: 'complete', drawings: [], images: [], diagnostics: [], bytes: 0 }
      })
      const posted: WorkerReply[] = []
      const command: CommandOf<'read'> = { kind: 'read', requestId: 't-1', owner: initialIdentity('k', 's'), query: { op: 'drawings', replyByteLimit: 1 << 20 } }
      const pending = kernel.handle(command, out => posted.push(out))
      await vi.advanceTimersByTimeAsync(QUERY_DEADLINE_MS + 1)
      await pending
      expect(posted).toMatchObject([{ kind: 'cancelled', diagnostics: [{ code: 'query-timeout' }] }])
      expect(estimatePayloadBytes(posted[0])).toBeLessThanOrEqual(replyLimitEnvelope(command))
      expect(kernel.lifecycleStats().transientBytes, 'the provider still holds its working set').toBeGreaterThanOrEqual(4096)
      open()
      await vi.advanceTimersByTimeAsync(1)
      expect(posted).toHaveLength(1)
      expect(kernel.lifecycleStats().transientBytes).toBe(0)
    } finally { vi.useRealTimers() }
  })
})
