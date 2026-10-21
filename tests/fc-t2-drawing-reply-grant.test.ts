// @vitest-environment node
// FC-T2: a drawings read carries a receiving grant main reserved before it was
// sent; the worker never posts a reply over that grant.
import { describe, expect, it, vi } from 'vitest'
import { QUERY_DEADLINE_MS } from '../src/worker/viewerKernel'
import { drawingOwnerMetadataBytes } from '../src/xlsx/view/drawingOwners'
import { DRAWING_MEMORY_REFUSED, drawingPayloadBudget } from '../src/xlsx/view/drawingPayload'
import { initialIdentity, transferablesOf, type CommandOf } from '../src/worker/viewerProtocol'
import { replyLimitEnvelope } from '../src/worker/viewerReplyLimit'
import { estimatePayloadBytes, measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { estimatePayloadBytes as ledgerEstimate } from '../src/xlsx/view/mainLedger'
import type { WorkerReply } from '../src/xlsx/view/contracts'
import { drawingsOf, kernelFixture, remoteFixture, until } from './helpers/drawing-grant'

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
    const before = measurePayloadBytes(out)
    expect(before).toBeGreaterThanOrEqual(64 * 1024 + 8192)
    const clone = structuredClone(out, { transfer: transferablesOf(out) })
    expect(measurePayloadBytes(clone)).toBe(before)
    expect(ledgerEstimate).toBe(estimatePayloadBytes)
  })
})

describe('kernel reply grant', () => {
  it('posts a final reply that measures exactly the grant and replaces one byte less with a payload-free quota reply', async () => {
    const { read } = await kernelFixture()
    // A granted reply carries its own size (replyBytes), counted in it.
    const { posted: [full] } = await read('d0', 1 << 30)
    const D = full.replyBytes!
    expect(D).toBe(measurePayloadBytes(full))
    expect(drawingsOf(full)?.completeness).toBe('complete')

    const exact = await read('d1', D)
    expect(exact.posted).toHaveLength(1)
    expect(drawingsOf(exact.posted[0])?.completeness).toBe('complete')
    expect(measurePayloadBytes(exact.posted[0])).toBe(D)

    const under = await read('d2', D - 1)
    expect(under.posted).toHaveLength(1)
    expect(under.posted[0]).toMatchObject({ kind: 'quota', requestId: 'd2', resource: 'memory', requiredBytes: D, limitBytes: D - 1, diagnostics: [] })
    expect(transferablesOf(under.posted[0])).toEqual([])
    expect(measurePayloadBytes(under.posted[0])).toBeLessThanOrEqual(replyLimitEnvelope(under.command))
  })

  it('checks the final reply, after the cached-value diagnostics were added', async () => {
    const { source, read } = await kernelFixture()
    const { posted: [plain] } = await read('d0', 1 << 30)
    const D = plain.replyBytes!
    vi.spyOn(source as unknown as { valueStatus(): unknown }, 'valueStatus').mockReturnValue({ policy: 'cached', message: 'Cached values are shown' })
    const { posted: [diagnosed] } = await read('d1', 1 << 30)
    expect(diagnosed.diagnostics.some(d => d.code === 'cached-value-policy')).toBe(true)
    const final = diagnosed.replyBytes!
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
    expect(posted).toMatchObject([{ kind: 'error', code: 'resource-limit', requestId: 'd0', diagnostics: [] }])
    const envelope = replyLimitEnvelope(command)
    await read('d1', envelope)
    expect(provider).toHaveBeenCalledTimes(1)
  })
})

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
      expect(f.ledger.snapshot().byPool.drawings).toBe(measurePayloadBytes(out)! + drawingOwnerMetadataBytes(payload.drawings))
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
      // The baseline is taken quiescent (no reply in flight), so the comparison sees only this read's grant.
      await f.settle()
      const before = f.ledger.snapshot()
      await until(() => f.remote.drawings('s').state !== 'waiting' && f.remote.drawings('s').state !== 'loading')
      await f.settle()
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
      expect(measurePayloadBytes(posted[0])).toBeLessThanOrEqual(replyLimitEnvelope(command))
      expect(kernel.lifecycleStats().transientBytes, 'the provider still holds its working set').toBeGreaterThanOrEqual(4096)
      open()
      await vi.advanceTimersByTimeAsync(1)
      expect(posted).toHaveLength(1)
      expect(kernel.lifecycleStats().transientBytes).toBe(0)
    } finally { vi.useRealTimers() }
  })
})
