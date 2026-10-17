// @vitest-environment node
// FC-T2 review (Codex, findings-fc-t2-codex.md): counterexamples ported from
// review-fc-t2-counterexamples.test.ts (sha256 94106b51...).
import { describe, expect, it, vi } from 'vitest'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { ViewerClient, ViewerNotSentError } from '../src/worker/viewerClient'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { initialIdentity, MAX_REQUEST_ID_UNITS, type CommandOf } from '../src/worker/viewerProtocol'
import { replyLimitEnvelope } from '../src/worker/viewerReplyLimit'
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import type { WorkerReply } from '../src/xlsx/view/contracts'
import { drawingDoc, drawingsOf, E, kernelFixture, remoteFixture, until } from './helpers/drawing-grant'

const MiB = 1024 * 1024
async function kernelOver(doc: ReturnType<typeof drawingDoc>, options: { workerBudgetBytes?: number } = {}, afterDispatch?: (command: CommandOf<'read'>) => Promise<void> | undefined) {
  const source = createObjectWorkbookSource(doc, 'k')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source, ...(afterDispatch ? { afterDispatch: c => c.kind === 'read' ? afterDispatch(c) : undefined } : {}) }, { yield: async () => {}, ...options })
  await kernel.handle({ kind: 'open', requestId: 'o', docId: 'k', source: { kind: 'object', document: doc }, profile: 'D' }, () => {})
  const cmd = (requestId: string, replyByteLimit = 2 * MiB): CommandOf<'read'> => ({ kind: 'read', requestId, owner: initialIdentity('k', 's'), query: { op: 'drawings', replyByteLimit } })
  return { kernel, source, cmd, close: () => kernel.handle({ kind: 'close', requestId: 'c', docId: 'k' }, () => {}) }
}

describe('finding 1: the transport gate measures exactly, never by extrapolation', () => {
  it('refuses a real chart whose 4,194,304-unit title alone exceeds a 6 MiB grant', async () => {
    const doc = drawingDoc(['s'])
    const title = '界'.repeat(4 * MiB)
    const series = Array.from({ length: 64 }, (_, s) => ({ name: String(s), color: '4472C4', values: Array.from({ length: 4096 }, (_, i) => i), line: true, markers: false }))
    doc.sheets[0].drawings![0].content = { kind: 'xlsxChart', chart: { title, groups: [{ type: 'line', grouping: 'standard', series, gapWidth: 150, overlap: 0 }], categories: [], categoryAxis: { gridlines: false, deleted: false }, valueAxis: { gridlines: true, deleted: false }, fontFamily: 'Calibri', fontSizePt: 10 } } as never
    const f = await kernelOver(doc)
    const posted: WorkerReply[] = []
    try {
      await f.kernel.handle(f.cmd('valid', 6 * MiB), out => posted.push(out))
      expect(posted).toHaveLength(1)
      expect(posted[0]).toMatchObject({ kind: 'quota', limitBytes: 6 * MiB })
      expect((posted[0] as { requiredBytes?: number }).requiredBytes).toBeGreaterThan(title.length * 2 + 64 * 4096 * 8)
    } finally { await f.close() }
  })

  it('never posts an 8 MiB buffer placed after a long primitive array under a 6 MiB grant', async () => {
    const f = await kernelOver(drawingDoc(['s']))
    const image = new Uint8Array(8 * MiB), values = Array.from({ length: 300_000 }, (_, i) => i)
    const payload = { completeness: 'complete' as const, images: [], diagnostics: [], bytes: image.byteLength, drawings: [{
      image: { data: image, mime: 'image/png', partPath: 'p.png' }, xEmu: 0, yEmu: 0, widthEmu: 80 * E, heightEmu: 20 * E,
      content: { kind: 'xlsxChart', chart: { groups: [{ series: [{ values }] }] } },
    }] }
    vi.spyOn(f.source.store('s'), 'readDrawings').mockImplementation(async (_o, _p, memory) => { memory?.set(image.byteLength + values.length * 8); return payload as never })
    const sunk: { kind: string; transferred: number }[] = []
    try {
      await f.kernel.handle(f.cmd('large', 6 * MiB), (out, transfer) => sunk.push({ kind: out.kind, transferred: transfer.reduce((n, b) => n + b.byteLength, 0) }))
      expect(sunk).toEqual([{ kind: 'quota', transferred: 0 }])
    } finally { await f.close() }
  })

  // Confirmation review R2: no object-count cap; the walk's workspace is admitted instead.
  it('measures and delivers a 300,000-object reply that fits its grant and the worker pool', async () => {
    const f = await kernelOver(drawingDoc(['s']))
    const nodes = Array.from({ length: 300_000 }, () => ({}))
    vi.spyOn(f.source.store('s'), 'readDrawings').mockResolvedValue({ completeness: 'complete', images: [], diagnostics: [], bytes: 0, drawings: [{ xEmu: 0, yEmu: 0, widthEmu: 1, heightEmu: 1, content: { kind: 'xlsxChart', chart: { nodes } } }] } as never)
    const posted: WorkerReply[] = []
    try {
      await f.kernel.handle(f.cmd('many', 60 * MiB), out => posted.push(out))
      expect(posted).toHaveLength(1)
      expect(drawingsOf(posted[0])?.completeness).toBe('complete')
      expect((posted[0] as { replyBytes?: number }).replyBytes).toBe(measurePayloadBytes(posted[0]))
    } finally { await f.close() }
  })

  it('counts strings, arrays of primitives and shared objects exactly, and reports a refused workspace as unmeasurable', () => {
    const shared = { a: 1 }
    expect(measurePayloadBytes('abc')).toBe(16 + 6)
    expect(measurePayloadBytes([1, 2, 'xy'])).toBe(16 + 3 * 8 + 8 + 8 + 16 + 4)
    expect(measurePayloadBytes([shared, shared])).toBe(16 + 2 * 8 + 24 + 16 + 8)
    expect(measurePayloadBytes(Array.from({ length: 64 }, () => ({})), () => false)).toBeUndefined()
    expect(measurePayloadBytes([{}, {}, {}])).toBe(16 + 3 * 8 + 3 * 24)
  })
})

describe('finding 2: a duplicate request id never touches the original credit', () => {
  it('keeps the original read credit and the main peak within the cap', async () => {
    const f = await remoteFixture({ hold: true, imageBytes: MiB })
    let first: Promise<unknown> | undefined
    try {
      first = f.remote.request(f.command('dup'))
      const before = f.grant('dup')
      expect(before).toBe(16 * MiB)
      expect(typeof await f.remote.request(f.command('dup'))).toBe('symbol')
      expect(f.grant('dup')).toBe(before)
      f.pressure(64 * 1024)
      await until(() => f.gates.length > 0)
      f.gates.forEach(open => open())
      await first
      expect(f.ledger.peakBytes).toBeLessThanOrEqual(f.cap)
    } finally { f.gates.forEach(open => open()); await first; f.ledger.release('dom', 'pressure'); await f.remote.close() }
  })
})

describe('finding 3: the outgoing reply stays charged until it was posted', () => {
  it('holds the provider share while the finished reply waits to be posted', async () => {
    let release!: () => void, arrived!: () => void
    const gate = new Promise<void>(r => { release = r }), at = new Promise<void>(r => { arrived = r })
    const f = await kernelOver(drawingDoc(['s'], MiB), {}, c => { if (c.query.op !== 'drawings') return undefined; arrived(); return gate })
    const pending = f.kernel.handle(f.cmd('read'), () => {})
    try {
      await at
      expect(f.kernel.session('k')!.memoryStats().transientBytes).toBeGreaterThanOrEqual(MiB)
    } finally { release(); await pending; await f.close() }
    expect(f.kernel.lifecycleStats().transientBytes).toBe(0)
  })

  it('leaves no microtask gap in which a second read can clone beside the unposted first (no afterDispatch)', async () => {
    // Fits one reply plus its admitted walk workspace, not two 1 MiB clones.
    const doc = drawingDoc(['s'], MiB), limit = Math.floor(1.9 * MiB)
    const f = await kernelOver(doc, { workerBudgetBytes: limit })
    const store = f.source.store('s'), raw = store.readDrawings!.bind(store), buffers: WeakRef<ArrayBufferLike>[] = []
    store.readDrawings = async (...args) => { const p = await raw(...args); for (const image of p.images) buffers.push(new WeakRef(image.data.buffer)); return p }
    const live = () => buffers.reduce((sum, ref) => sum + (ref.deref()?.byteLength ?? 0), 0)
    let posted = 0, maxLive = 0
    const kinds: (string | undefined)[] = []
    const sink = (_out: WorkerReply, transfer: ArrayBuffer[]) => { maxLive = Math.max(maxLive, live()); posted++; kinds.push(drawingsOf(_out)?.completeness); structuredClone(_out, { transfer }) }
    const first = f.kernel.handle(f.cmd('one'), sink)
    let gap = false
    for (let i = 0; i < 100 && !posted; i++) {
      await Promise.resolve()
      if (!posted && live() > 0 && f.kernel.lifecycleStats().transientBytes === 0) { gap = true; break }
    }
    const second = gap ? f.kernel.handle(f.cmd('two'), sink) : undefined
    try {
      await Promise.all([first, second])
      expect(kinds[0], 'the first read is delivered').toBe('complete')
      expect(gap, 'a finished result must not be live outside the worker pool').toBe(false)
      expect(maxLive).toBeLessThanOrEqual(limit)
    } finally { await Promise.all([first, second]); await f.close() }
  })

  it('refuses the second of two staggered clones that only fit one at a time', async () => {
    // Fits one reply plus its admitted walk workspace, not two 1 MiB clones.
    const limit = Math.floor(1.9 * MiB)
    let release!: () => void, firstAt!: () => void, secondAt!: () => void, count = 0
    const gate = new Promise<void>(r => { release = r }), firstReady = new Promise<void>(r => { firstAt = r }), secondReady = new Promise<void>(r => { secondAt = r })
    const f = await kernelOver(drawingDoc(['s'], MiB), { workerBudgetBytes: limit }, c => {
      if (c.query.op !== 'drawings') return undefined
      if (++count === 1) firstAt(); else secondAt()
      return gate
    })
    const posted: WorkerReply[] = []
    const first = f.kernel.handle(f.cmd('one'), out => posted.push(out))
    await firstReady
    const second = f.kernel.handle(f.cmd('two'), out => posted.push(out))
    try {
      await Promise.race([secondReady, second])
      release()
      await Promise.all([first, second])
      const payloads = posted.flatMap(out => { const p = drawingsOf(out); return p?.completeness === 'complete' ? [p] : [] })
      const actual = payloads.reduce((sum, p) => sum + p.images.reduce((n, image) => n + image.data.byteLength, 0), 0)
      expect(payloads.length, 'one of the two is delivered').toBe(1)
      expect(actual).toBeLessThanOrEqual(limit)
    } finally { release(); await Promise.all([first, second]); await f.close() }
  })
})

describe('finding 4: malformed grants and request ids', () => {
  it.each([1, Number.NaN, -1, Infinity, 1.5])('answers grant %s with one constant-size error reply, never a larger substitute, without the provider', async limit => {
    const f = await kernelFixture()
    const provider = vi.spyOn(f.source.store('s'), 'readDrawings')
    const { posted, command } = await f.read('bad', limit)
    expect(provider).not.toHaveBeenCalled()
    expect(posted).toMatchObject([{ kind: 'error', code: 'resource-limit', retryable: false, diagnostics: [] }])
    const reference = await f.read('bad', 0)
    expect(measurePayloadBytes(posted[0])).toBe(measurePayloadBytes(reference.posted[0]))
    expect(measurePayloadBytes(posted[0])!).toBeLessThanOrEqual(replyLimitEnvelope(command))
  })

  it('posts nothing for a drawings read whose request id exceeds the protocol maximum', async () => {
    const f = await kernelFixture()
    const provider = vi.spyOn(f.source.store('s'), 'readDrawings')
    const { posted } = await f.read('x'.repeat(MAX_REQUEST_ID_UNITS + 1), 1)
    expect(posted).toEqual([])
    expect(provider).not.toHaveBeenCalled()
  })

  it('refuses to send an over-long request id from the client', async () => {
    const port = { postMessage: vi.fn(), addEventListener: () => {}, removeEventListener: () => {} }
    const client = new ViewerClient(port)
    await expect(client.send({ kind: 'read', requestId: 'x'.repeat(MAX_REQUEST_ID_UNITS + 1), owner: initialIdentity('d', 's'), query: { op: 'drawings' } })).rejects.toBeInstanceOf(ViewerNotSentError)
    expect(port.postMessage).not.toHaveBeenCalled()
  })
})
