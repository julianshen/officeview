// @vitest-environment node
// FC-T2 confirmation review (Codex, findings-fc-t2-codex.md "Confirmation
// (8db7b1a)"), ported from review-fc-t2-confirm-new.test.ts (sha256 1f06b2f1...):
// main never walks a granted drawings reply; the worker walk admits its
// workspace in chunks, without a fixed object cap.
import { describe, expect, it, vi } from 'vitest'
import { parseXlsxChart } from '../src/charts'
import { parseXmlOrdered } from '../src/core/xml'
import * as bytes from '../src/xlsx/view/payloadBytes'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { ViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { initialIdentity, type CommandOf } from '../src/worker/viewerProtocol'
import { admitReply } from '../src/worker/viewerReplyLimit'
import { MemoryPool } from '../src/xlsx/view/tileCache'
import type { WorkerReply } from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'
import { drawingDoc, drawingsOf, kernelFixture } from './helpers/drawing-grant'

/** A cached line chart with 64 one-point series, as the parser produces it. */
function chart() {
  const xml = '<chartSpace><chart><plotArea><lineChart>' + Array.from({ length: 64 }, (_, i) => `<ser><idx val="${i}"/><tx><v>S${i}</v></tx><val><numLit><ptCount val="1"/><pt idx="0"><v>1</v></pt></numLit></val></ser>`).join('') + '</lineChart></plotArea></chart></chartSpace>'
  const model = parseXlsxChart(parseXmlOrdered(xml), undefined)!
  expect(model.groups[0].series).toHaveLength(64)
  return model
}
function docWithCharts(n: number): XlsxDocument {
  const doc = drawingDoc(['s']), frame = doc.sheets[0].drawings![0], model = chart()
  doc.sheets[0].drawings = Array.from({ length: n }, (_, i) => ({ ...structuredClone(frame), source: { ...frame.source!, treePath: `twoCellAnchor[${i}]/graphicFrame[0]` }, content: { kind: 'xlsxChart', chart: structuredClone(model) } })) as never
  doc.sheets[0].drawingAnchors = { unresolved: 0, records: Array.from({ length: n }, (_, i) => ({ id: 'd' + i, kind: 'twoCell' as const, from: { col: 1, row: 0, colOffEmu: 0, rowOffEmu: 0 }, to: { col: 2, row: 1, colOffEmu: 0, rowOffEmu: 0 } })) }
  return doc
}
async function openRemote(doc: XlsxDocument, workerBudgetBytes?: number) {
  const source = createObjectWorkbookSource(doc, 'k'), kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {}, ...(workerBudgetBytes ? { workerBudgetBytes } : {}) })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'k', profile: 'D', source: { kind: 'object', document: doc } })
  const remote = new RemoteWorkbook(client, { docId: 'k', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map([['s', kernel.session('k')!.layout('s')]]) })
  return { kernel, source, client, remote }
}
const read = (requestId: string, replyByteLimit?: number): CommandOf<'read'> => ({ kind: 'read', requestId, owner: initialIdentity('k', 's'), query: { op: 'drawings', ...(replyByteLimit === undefined ? {} : { replyByteLimit }) } })

describe('R2: legal content that fits the budgets is measured, not refused by an object count', () => {
  it('delivers a 1,500-chart reply (over 200,000 objects) under the D budgets', async () => {
    const f = await openRemote(docWithCharts(1500))
    try {
      const out = await f.remote.request(read('admitted'))
      expect(typeof out === 'object' && drawingsOf(out as WorkerReply)?.completeness).toBe('complete')
      expect(drawingsOf(out as WorkerReply)!.drawings).toHaveLength(1500)
      expect(f.remote.mainLedger.peakBytes).toBeLessThanOrEqual(64 * 1024 * 1024)
    } finally { await f.remote.close() }
  }, 30000)
})

describe('R1: no unadmitted measurement workspace', () => {
  it('main adopts the size the worker stamped on the reply and never walks it', async () => {
    const f = await openRemote(docWithCharts(300))
    // A walk of the reply after the worker posted it (no longer pending) is main's.
    let onMain = 0
    const record = <F extends (...args: never[]) => unknown>(original: F) => ((...args: Parameters<F>) => {
      if ((args[0] as { requestId?: string; kind?: string })?.requestId === 'stamped' && (args[0] as { kind?: string }).kind === 'complete' && !f.client.isPending('stamped')) onMain++
      return original(...args)
    }) as F
    const originalMeasure = bytes.measurePayloadBytes, originalEstimate = bytes.estimatePayloadBytes
    const measure = vi.spyOn(bytes, 'measurePayloadBytes').mockImplementation(record(originalMeasure))
    const estimate = vi.spyOn(bytes, 'estimatePayloadBytes').mockImplementation(record(originalEstimate))
    try {
      const out = await f.remote.request(read('stamped')) as WorkerReply
      measure.mockRestore(); estimate.mockRestore()
      expect(onMain, 'main builds no visited set on receipt').toBe(0)
      expect(Number.isSafeInteger(out.replyBytes)).toBe(true)
      expect(out.replyBytes).toBe(bytes.measurePayloadBytes(out))
      expect(f.remote.mainLedger.ownerBytes('replies', 'drawing-grant:stamped')).toBe(0)
    } finally { measure.mockRestore(); estimate.mockRestore(); await f.remote.close() }
  }, 30000)

  it('charges the whole grant when a granted reply carries no valid size', async () => {
    const f = await openRemote(drawingDoc(['s']))
    const receive = (f.client as unknown as { receive(data: unknown): void })
    const original = receive.receive.bind(f.client)
    receive.receive = data => { const d = data as { replyBytes?: unknown; requestId?: string }; if (d.requestId === 'nosize') delete d.replyBytes; original(data) }
    try {
      const before = f.remote.mainLedger.snapshot().byPool.drawings
      await f.remote.request(read('nosize'))
      expect(f.remote.mainLedger.snapshot().byPool.drawings - before).toBeGreaterThanOrEqual(16 * 1024 * 1024)
    } finally { await f.remote.close() }
  })

  it('admits the walk workspace in chunks before it grows, and walks without Object.keys arrays', () => {
    const graph = Array.from({ length: 10_000 }, (_, i) => ({ i }))
    const admitted: number[] = []
    const keys = vi.spyOn(Object, 'keys')
    try {
      expect(bytes.measurePayloadBytes(graph, n => { admitted.push(n); return true })).toBe(16 + 10_000 * 8 + 10_000 * (24 + 16 + 8))
      expect(keys).not.toHaveBeenCalled()
    } finally { keys.mockRestore() }
    // 10,001 objects: a first chunk before the walk starts, then chunks doubling up to 4,096 entries.
    expect(admitted.map(n => n / bytes.MEASURE_ENTRY_BYTES)).toEqual([16, 16, 32, 64, 128, 256, 512, 1024, 2048, 4096, 4096])
    expect(admitted.reduce((a, b) => a + b, 0) / bytes.MEASURE_ENTRY_BYTES).toBeGreaterThanOrEqual(10_001)
    let calls = 0
    expect(bytes.measurePayloadBytes(graph, () => ++calls < 2)).toBeUndefined()
    expect(calls).toBe(2)
    // Confirmation r3 F1: no unadmitted allowance, even for a one-object walk.
    expect(bytes.measurePayloadBytes({ a: 1 }, () => false), 'nothing is walked before the first chunk is admitted').toBeUndefined()
  })

  it('funds every walked entry: 127 zero-length views under an 8,704 B pool are admitted at 17,408 B or refused', () => {
    const buffer = new ArrayBuffer(0)
    const value = Array.from({ length: 127 }, () => new Uint8Array(buffer))
    const pool = new MemoryPool({ limitBytes: 8704 }), share = pool.charge()
    let admitted = 0
    const measured = bytes.measurePayloadBytes(value, n => share.set(admitted += n))
    if (measured !== undefined) expect(admitted).toBeGreaterThanOrEqual(128 * bytes.MEASURE_ENTRY_BYTES)
    expect(measured).toBeUndefined()
    share.release()
  })

  it('still sizes the kernel-built payload-free replies when the pool admits nothing', () => {
    const command = read('full', 1 << 20)
    const cancelled = { kind: 'cancelled', commandKind: 'read', requestId: 'full', owner: command.owner, diagnostics: [{ code: 'query-timeout', message: 'Query timed out' }] } as WorkerReply
    const out = admitReply(command, cancelled, 1 << 20, () => false, true)
    expect(out).toMatchObject({ kind: 'cancelled', diagnostics: [{ code: 'query-timeout' }] })
    expect(out.replyBytes).toBe(bytes.measurePayloadBytes(out))
    const payload = { ...cancelled, kind: 'complete', result: { kind: 'read', result: { op: 'drawings', result: { completeness: 'complete', drawings: Array.from({ length: 100 }, () => ({})), images: [], diagnostics: [], bytes: 0 } } } } as unknown as WorkerReply
    const refused = admitReply(command, payload, 1 << 20, () => false, false)
    expect(refused).toMatchObject({ kind: 'quota', limitBytes: 1 << 20 })
    expect('requiredBytes' in refused).toBe(false)
    expect(refused.replyBytes).toBe(bytes.measurePayloadBytes(refused))
  })

  it('refuses a reply whose workspace the worker pool cannot admit as a sizeless quota, which main retries', async () => {
    const doc = docWithCharts(300)
    // The payload fits the worker pool; its 41,000-object walk workspace does not.
    const f = await openRemote(doc, 8 * 1024 * 1024)
    const posted: WorkerReply[] = []
    try {
      await f.kernel.handle(read('pressed', 32 * 1024 * 1024), out => posted.push(out))
      expect(posted).toHaveLength(1)
      expect(posted[0]).toMatchObject({ kind: 'quota', limitBytes: 32 * 1024 * 1024 })
      expect('requiredBytes' in posted[0]).toBe(false)
      expect(f.kernel.lifecycleStats().transientBytes).toBe(0)
      const out = await f.remote.request(read('pressed-main'))
      expect(out).toMatchObject({ kind: 'complete', result: { result: { result: { completeness: 'incomplete', error: { code: 'resource-limit', retryable: true } } } } })
      // Confirmation r3 F2: the refusal names the worker's memory, not main's.
      const refusal = drawingsOf(out as WorkerReply)!
      expect(refusal.error?.message).toMatch(/worker/i)
      expect(refusal.error?.message).not.toMatch(/main-thread/)
      expect(refusal.diagnostics[0].message).toBe(refusal.error?.message)
    } finally { await f.remote.close() }
  }, 30000)
})

// Confirmation r4 (Codex, f32a676): the constant-space tree sizing is only for
// replies the kernel builds itself; a provider graph shares objects.
describe('r4: provider results always take the exact admitted walk', () => {
  /** Two drawings that share one 100,000-number dash array (as resolveDrawingStyle retains it). */
  async function sharedDash() {
    const f = await kernelFixture()
    const dash = Array.from({ length: 100_000 }, (_, i) => i % 7)
    const node = (i: number) => ({ xEmu: 0, yEmu: 0, widthEmu: 1, heightEmu: 1, source: { partPath: 'd', treePath: `sp[${i}]`, element: 'sp', representation: 'native' }, style: { line: { dash } } })
    vi.spyOn(f.source.store('s'), 'readDrawings').mockResolvedValue({ completeness: 'complete', images: [], diagnostics: [], bytes: 0, drawings: [node(0), node(1)] } as never)
    return f
  }

  it('delivers a shared-array reply that fits, stamped with its exact deduplicated size', async () => {
    const f = await sharedDash()
    const { posted: [probe] } = await f.read('dash', 1 << 30)
    const exact = probe.replyBytes!
    expect(exact).toBe(bytes.measurePayloadBytes(probe))
    // A tree count holds the shared array twice (about 1.6 MB more).
    expect(bytes.smallTreeBytes(probe)! - exact).toBeGreaterThan(1_500_000)
    const { posted: [out] } = await f.read('dash', exact)
    expect(drawingsOf(out)?.completeness).toBe('complete')
    expect(out.replyBytes).toBe(exact)
  })

  it('never sizes a provider result with the tree shortcut; kernel-built refusals still are', async () => {
    const f = await sharedDash()
    const tree = vi.spyOn(bytes, 'smallTreeBytes')
    try {
      await f.read('dash', 1 << 30)
      expect(tree.mock.calls.filter(([v]) => (v as { kind?: string })?.kind === 'complete')).toEqual([])
      tree.mockClear()
      await f.read('tiny', 10)
      await f.read('quota', 4096)
      expect(tree.mock.calls.some(([v]) => (v as { kind?: string })?.kind === 'quota')).toBe(true)
    } finally { tree.mockRestore() }
  })

  it('stops scanning a tree as soon as it is ineligible', () => {
    let reads = 0
    const items = Array.from({ length: 10_000 }, () => ({}))
    const watched = new Proxy(items, { get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver) } })
    expect(bytes.smallTreeBytes(watched)).toBeUndefined()
    expect(reads).toBeLessThanOrEqual(65)
  })
})
