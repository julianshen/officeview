// @vitest-environment node
/** Spec §10.2 (codex D4, E1-E5): one worker walker measures a reply and
 * collects its transferables in a single checkpointed, admitted pass; main
 * charges the stamp without walking; the owner is checked after the walk. */
import { afterEach, describe, expect, it, vi } from 'vitest'
import * as payload from '../src/xlsx/view/payloadBytes'
import * as protocol from '../src/worker/viewerProtocol'
import { ViewerClient } from '../src/worker/viewerClient'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { StopWalk, UNPOOLED_WALK_BYTES, walkAdmission, walkPayload } from '../src/worker/replyMeasure'
import { boundedQueryStore } from '../src/worker/viewerQueryBudget'
import { queryPayloadBytes } from '../src/xlsx/view/queryPayload'
import { CheckpointBudget } from '../src/xlsx/view/scheduler'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { initialIdentity } from '../src/worker/viewerProtocol'
import { workbookOf } from './helpers/xlsx-find'
import { kernelFixture } from './helpers/drawing-grant'
import { WorkerViewportPlans } from '../src/components/xlsx/workerViewport'
import { accountQueryBytes } from '../src/worker/viewerQueryBudget'
import { FindError } from '../src/xlsx/view/find'
import type { WorkerReply } from '../src/xlsx/view/contracts'

vi.mock('../src/xlsx/view/payloadBytes', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/view/payloadBytes')>()
  return { ...actual, estimatePayloadBytes: vi.fn(actual.estimatePayloadBytes), measurePayloadBytes: vi.fn(actual.measurePayloadBytes) }
})
vi.mock('../src/xlsx/view/queryPayload', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/view/queryPayload')>()
  return { ...actual, queryPayloadBytes: vi.fn(actual.queryPayloadBytes) }
})
vi.mock('../src/worker/viewerProtocol', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/worker/viewerProtocol')>()
  return { ...actual, transferablesOf: vi.fn(actual.transferablesOf) }
})
const queryPayloadBytesMock = () => vi.mocked(queryPayloadBytes)
const estimate = vi.mocked(payload.estimatePayloadBytes), measure = vi.mocked(payload.measurePayloadBytes), transfers = vi.mocked(protocol.transferablesOf)
afterEach(() => { vi.restoreAllMocks() })

const frozen = (onYield: () => void = () => {}) => new CheckpointBudget(() => {}, undefined, { now: () => 0, yield: async () => { onYield() } })
const settle = async <T>(run: T | Promise<T>): Promise<T> => run instanceof Promise ? await run : run
/** An array whose elements count every read. */
function counted(length: number) {
  const reads = { count: 0 }
  const items = Array.from({ length }, () => 0)
  for (let i = 0; i < length; i++) Object.defineProperty(items, i, { get() { reads.count++; return 0 }, enumerable: true })
  return { items, reads }
}

describe('the one payload walker', () => {
  it('measures exactly as measurePayloadBytes and collects transfers as transferablesOf (shared buffers once)', async () => {
    const shared = new ArrayBuffer(64)
    const value = { a: new Uint8Array(shared, 0, 8), b: [new Uint32Array(shared, 8, 2), new Float64Array(4)], c: { d: new ArrayBuffer(16), s: 'text', n: 3 } }
    const outcome = await settle(walkPayload(value, () => true, frozen()))
    expect(outcome.kind).toBe('measured')
    if (outcome.kind !== 'measured') return
    expect(outcome.bytes).toBe(payload.measurePayloadBytes(value))
    expect(new Set(outcome.transfer)).toEqual(new Set(protocol.transferablesOf(value)))
    expect(outcome.transfer).toHaveLength(3)
    const page = { items: [{ row: 0, cells: Array.from({ length: 300 }, (_, i) => ({ row: 0, col: i, value: `v${i}`, buf: new Uint8Array(2) })) }] }
    const query = await settle(walkPayload(page, () => true, frozen(), { model: 'query', limit: 8 * 1024 * 1024 }))
    expect(query).toMatchObject({ kind: 'measured', bytes: queryPayloadBytes(page, 8 * 1024 * 1024) })
  })
  it('an 8,192-element getter array first yields within 4,096 reads, continuing one iterator (codex E2)', async () => {
    const { items, reads } = counted(8192)
    let first = -1
    const outcome = await settle(walkPayload(items, () => true, frozen(() => { if (first < 0) first = reads.count })))
    expect(outcome.kind).toBe('measured')
    expect(first).toBeGreaterThan(0)
    expect(first).toBeLessThanOrEqual(4096)
    expect(reads.count).toBe(8192) // each element read once: no restart
  })
  it('completes a small payload synchronously, checked at its end; a stop is not an error', () => {
    expect(walkPayload({ a: [1, 2, 3] }, () => true, frozen())).toMatchObject({ kind: 'measured' })
    const stop = new CheckpointBudget(() => { throw new StopWalk() }, undefined, { now: () => 0, yield: async () => {} })
    expect(walkPayload({ a: [1, 2, 3] }, () => true, stop)).toEqual({ kind: 'stopped' })
  })
  it('checks within 256 visits', async () => {
    const { items, reads } = counted(8192)
    let stopAt = -1
    const budget = new CheckpointBudget(() => { if (reads.count > 0) { stopAt = reads.count; throw new StopWalk() } }, undefined, { now: () => 0, yield: async () => {} })
    expect(await settle(walkPayload(items, () => true, budget))).toEqual({ kind: 'stopped' })
    expect(stopAt).toBeLessThanOrEqual(256)
  })
})

/** Kernel and remote over a one-sheet object document whose readRows answers `cells`. */
async function harness(cells: () => { row: number; col: number; value: unknown; styleIndex: number }[], afterDispatch?: (kernel: ViewerKernel) => void) {
  const document = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }])
  const source = createObjectWorkbookSource(document, 'doc'), store = source.store('s')
  vi.spyOn(store, 'readRows').mockImplementation(async () => ({ items: [{ row: 0, cells: cells() }], completeness: 'complete', stamp: store.stamp }) as never)
  const kernel: ViewerKernel = new ViewerKernel({ parse: async () => document, openSource: async () => source, ...(afterDispatch ? { afterDispatch: command => { if (command.kind === 'read') afterDispatch(kernel) } } : {}) })
  await kernel.handle({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
  return { kernel, document, source, close: () => kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
}
const rowRead = (requestId = 'r') => ({ kind: 'read' as const, requestId, owner: initialIdentity('doc', 's'), query: { op: 'rows' as const, range: { r0: 0, r1: 0, c0: 0, c1: 8191 } } })
const wide = () => Array.from({ length: 8192 }, (_, col) => ({ row: 0, col, value: col, styleIndex: 0 }))
/** Holds the first reply-walk pause until released. */
function gate() {
  let begin!: () => void, release!: () => void
  const began = new Promise<void>(resolve => { begin = resolve }), held = new Promise<void>(resolve => { release = resolve })
  const native = CheckpointBudget.prototype.pause
  const pauses = { reply: 0 }
  let gated = false
  vi.spyOn(CheckpointBudget.prototype, 'pause').mockImplementation(async function (this: CheckpointBudget) {
    if ((this as unknown as { phase: string }).phase === 'reply') {
      pauses.reply++
      if (!gated) { gated = true; begin(); await held }
    }
    return native.call(this)
  })
  return { began, pauses, release: () => release() }
}

describe('kernel publication through the walker', () => {
  it('a read whose owner a newer commit replaced during the walk is stale (codex E1)', async () => {
    const h = await harness(wide), g = gate(), replies: WorkerReply[] = []
    const pending = h.kernel.handle(rowRead(), out => replies.push(out))
    try {
      await g.began
      const sort: WorkerReply[] = []
      await h.kernel.handle({ kind: 'sort', requestId: 'sort', owner: initialIdentity('doc', 's'), spec: null }, out => sort.push(out))
      expect(sort.at(-1)?.kind).toBe('complete')
      g.release(); await pending
      expect(replies.at(-1)).toMatchObject({ kind: 'error', code: 'stale-source' })
      // The walk stopped at its first checkpoint after the commit, not after the whole reply.
      expect(g.pauses.reply).toBe(1)
    } finally { g.release(); await pending; await h.close() }
  })
  it('a cancel during the walk does not hide an already committed sort (codex E1)', async () => {
    const document = workbookOf([{ sheetId: 's', cells: Array.from({ length: 2048 }, (_, row) => ({ row, col: 0, value: row })) }])
    for (const row of document.sheets[0].rows) row.heightPt = 25 + row.index % 2
    const kernel = new ViewerKernel({ parse: async () => document })
    await kernel.handle({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
    const g = gate(), replies: WorkerReply[] = []
    const pending = kernel.handle({ kind: 'sort', requestId: 'sort', owner: initialIdentity('doc', 's'), spec: null }, out => replies.push(out))
    try {
      await g.began
      expect(kernel.session('doc')!.isCurrent(initialIdentity('doc', 's'))).toBe(false)
      await kernel.handle({ kind: 'cancel', requestId: 'cancel', targetRequestId: 'sort' }, () => {})
      g.release(); await pending
      expect(replies.at(-1)).toMatchObject({ kind: 'complete', result: { kind: 'view' } })
    } finally { g.release(); await pending; await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
  }, 15_000)
  it('workspace pressure posts a typed retryable refusal and no data (codex E3)', async () => {
    const value = 'x'.repeat(10_000)
    let pressure: { release(): void } | undefined
    const h = await harness(() => Array.from({ length: 8192 }, (_, col) => ({ row: 0, col, value, styleIndex: 0 })), kernel => {
      const memory = kernel.session('doc')!.memory, charge = memory.charge()
      expect(charge.set(memory.limitBytes - memory.totalBytes)).toBe(true)
      pressure = charge
    })
    const replies: WorkerReply[] = []
    try {
      await h.kernel.handle(rowRead(), out => replies.push(out))
      const last = replies.at(-1) as WorkerReply & { code?: string; retryable?: boolean; result?: unknown; replyBytes?: number }
      expect(last).toMatchObject({ kind: 'error', code: 'resource-limit', retryable: true })
      expect(last.result).toBeUndefined()
      expect(last.replyBytes).toBeGreaterThan(0)
    } finally { pressure?.release(); await h.close() }
  })
  it('posting reuses the walker\'s transfer list: no second traversal (codex E5)', async () => {
    const h = await harness(wide)
    try {
      transfers.mockClear()
      const replies: WorkerReply[] = []
      await h.kernel.handle(rowRead(), out => replies.push(out))
      expect(replies.at(-1)).toMatchObject({ kind: 'complete' })
      expect(transfers.mock.calls.filter(([message]) => (message as { requestId?: string }).requestId === 'r')).toHaveLength(0)
    } finally { await h.close() }
  })
})

describe('bounded query facade accounting (codex E4)', () => {
  it('an 8,192-cell query page is accounted in steps, first yielding within 4,096 visits', async () => {
    const document = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }])
    const store = createObjectWorkbookSource(document, 'doc').store('s')
    let visits = 0, first = -1, pauses = 0
    const cells = Array.from({ length: 8192 }, (_, col) => ({ row: 0, col, styleIndex: 0, get value() { visits++; return col } }))
    vi.spyOn(store, 'readRows').mockResolvedValue({ items: [{ row: 0, cells }], completeness: 'complete', stamp: store.stamp } as never)
    const native = CheckpointBudget.prototype.pause
    vi.spyOn(CheckpointBudget.prototype, 'pause').mockImplementation(async function (this: CheckpointBudget) {
      if ((this as unknown as { phase: string }).phase === 'query') { pauses++; if (first < 0) first = visits }
      return native.call(this)
    })
    const pool = { charge: () => ({ set: () => true, release: () => {} }) } as never
    const page = await boundedQueryStore(store, 8 * 1024 * 1024, undefined, 20_000, false, false, pool).readRows({ r0: 0, r1: 0, c0: 0, c1: 8191 }, undefined, { signal: new AbortController().signal, stamp: { ...store.stamp } })
    expect(page.items[0].cells).toHaveLength(8192)
    expect(visits).toBe(8192)
    expect(pauses).toBeGreaterThanOrEqual(2)
    expect(first).toBeLessThanOrEqual(4096)
  })
})

describe('main charges a stamped reply without walking it', () => {
  it('an 8,192-cell readRows reply: no main walk, and the charge covers the reply', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: Array.from({ length: 8192 }, (_, i) => ({ row: Math.floor(i / 64), col: i % 64, value: `cell-${i}` })) }])
    const source = createObjectWorkbookSource(doc, 'stamp')
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source })
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'stamp', source: { kind: 'object', document: doc }, profile: 'D' })
    const session = kernel.session('stamp')!
    const remote = new RemoteWorkbook(client, { docId: 'stamp', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map([['s', session.layout('s')]]) })
    try {
      const store = remote.store('s')
      estimate.mockClear(); measure.mockClear()
      const page = await store.readRows({ r0: 0, r1: 127, c0: 0, c1: 63 }, undefined, { signal: new AbortController().signal, stamp: { ...store.stamp } })
      expect(page.error).toBeUndefined()
      expect(page.items.reduce((sum, item) => sum + item.cells.length, 0)).toBe(8192)
      expect(estimate).not.toHaveBeenCalled()
      expect(measure).not.toHaveBeenCalled()
      const charged = remote.mainLedger.snapshot().byPool.replies
      expect(charged).toBeGreaterThanOrEqual(payload.measurePayloadBytes(page)!)
      // The worker's stamp, not the defensive unstamped charge (the whole main budget).
      expect(charged).toBeLessThan(MAIN_THREAD_BUDGETS.D)
    } finally { await remote.release() }
  })
})

describe('residuals closed on the one walker', () => {
  const keyRun = (count: number) => {
    const reads = { count: 0 }
    const texts = Array.from({ length: count }, (_, i) => `key-${i}`)
    const counted: string[] = []
    for (let i = 0; i < count; i++) Object.defineProperty(counted, i, { enumerable: true, get() { reads.count++; return texts[i] } })
    return { reads, set: { refId: 'ref-1', scope: 'scope', run: { count, kinds: new Uint8Array(count).fill(2), numbers: new Float64Array(count), texts: counted } } }
  }
  const filter = (keySets: unknown[]) => ({ kind: 'filter' as const, requestId: 'f', owner: initialIdentity('doc', 's'), nowMs: 0,
    spec: { range: { r0: 0, r1: 1, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, columns: new Map() }, keySets } as never)
  it('kernel key-set admission is accounted in walker steps (first yield within 4,096 visits)', async () => {
    const h = await harness(wide), { reads, set } = keyRun(8192)
    const marks: number[] = [0]
    const native = CheckpointBudget.prototype.pause
    vi.spyOn(CheckpointBudget.prototype, 'pause').mockImplementation(async function (this: CheckpointBudget) {
      if ((this as unknown as { phase: string }).phase === 'query') marks.push(reads.count)
      return native.call(this)
    })
    try {
      await h.kernel.handle(filter([set]), () => {})
      marks.push(reads.count)
      // Both accountings (the command's key sets, then each installed set) read every key.
      expect(reads.count).toBe(2 * 8192)
      expect(Math.max(...marks.slice(1).map((at, i) => at - marks[i]))).toBeLessThanOrEqual(4096)
    } finally { await h.close() }
  })
  it('kernel key-set accounting the pool cannot admit is a sizeless quota', async () => {
    const h = await harness(wide), { set } = keyRun(8192)
    const memory = h.kernel.session('doc')!.memory, pressure = memory.charge()
    expect(pressure.set(memory.limitBytes - memory.totalBytes)).toBe(true)
    const replies: WorkerReply[] = []
    try {
      await h.kernel.handle(filter([set]), out => replies.push(out))
      const last = replies.at(-1) as WorkerReply & { requiredBytes?: number }
      expect(last).toMatchObject({ kind: 'quota', resource: 'memory' })
      expect(last.requiredBytes).toBeUndefined()
    } finally { pressure.release(); await h.close() }
  })
  it('the query facade admits its walk to the pool: a refusal is a typed retryable resource-limit and no page', async () => {
    const document = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }])
    const store = createObjectWorkbookSource(document, 'doc').store('s')
    vi.spyOn(store, 'readRows').mockResolvedValue({ items: [{ row: 0, cells: wide() }], completeness: 'complete', stamp: store.stamp } as never)
    const refusing = { charge: () => ({ set: () => false, release: () => {} }) } as never
    const error = await boundedQueryStore(store, 8 * 1024 * 1024, undefined, 20_000, false, false, refusing)
      .readRows({ r0: 0, r1: 0, c0: 0, c1: 8191 }, undefined, { signal: new AbortController().signal, stamp: { ...store.stamp } }).then(() => null, cause => cause)
    expect(error).toBeInstanceOf(FindError)
    expect(error).toMatchObject({ code: 'resource-limit', retryable: true })
  })
  it('FC-T2: a granted drawings reply is measured and its transfers collected by the walker, once', async () => {
    const { read } = await kernelFixture(64 * 1024)
    measure.mockClear(); transfers.mockClear()
    const { posted: [full] } = await read('g1', 1 << 30)
    expect(full).toMatchObject({ kind: 'complete' })
    expect(full.replyBytes).toBe(payload.measurePayloadBytes(full))
    const own = (calls: unknown[][]) => calls.filter(([message]) => (message as { requestId?: string }).requestId === 'g1')
    expect(own(transfers.mock.calls)).toHaveLength(0)
    expect(own(measure.mock.calls)).toHaveLength(1) // the assertion's own call above
  })
})

describe('confirmation round (codex F1-F5)', () => {
  it('F1: an open reply walks against the new document session pool', async () => {
    const document = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }])
    ;(document as { cellStyles?: unknown[] }).cellStyles = Array.from({ length: 2048 }, () => ({ fontSizePt: 11 }))
    let charges: ReturnType<typeof vi.spyOn> | undefined
    const kernel: ViewerKernel = new ViewerKernel({ parse: async () => document, afterDispatch: command => { if (command.kind === 'open') charges = vi.spyOn(kernel.session('doc')!.memory, 'charge') } })
    await kernel.handle({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
    try { expect(charges).toHaveBeenCalled() } finally { await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
  })
  it('F1: a walk with no pool has a small fixed allowance, never an unlimited one', () => {
    const admit = walkAdmission(undefined)
    expect(admit(UNPOOLED_WALK_BYTES)).toBe(true)
    expect(admit(1)).toBe(false)
  })
  it('F2: the query walk holds one frame per open container, never pending children, all admitted', async () => {
    const { items, reads } = counted(10_000)
    let held = 0, first = -1
    const outcome = await settle(walkPayload(items, more => { held += more; return true }, frozen(() => { if (first < 0) first = reads.count }), { model: 'query', limit: 8 * 1024 * 1024 }))
    expect(outcome).toMatchObject({ kind: 'measured', bytes: queryPayloadBytes(Array.from({ length: 10_000 }, () => 0), 8 * 1024 * 1024), peakFrames: 1 })
    expect(first).toBeLessThanOrEqual(4096)
    if (outcome.kind !== 'measured') return
    expect(outcome.peakFrames).toBeLessThanOrEqual(outcome.admittedFrames)
    // One visited entry chunk (the array) plus its frames, both admitted.
    expect(held).toBeGreaterThanOrEqual(payload.MEASURE_FIRST_CHUNK_ENTRIES * payload.MEASURE_ENTRY_BYTES + outcome.admittedFrames * 256)
    const deep: unknown[] = []
    let at = deep
    for (let depth = 0; depth < 40; depth++) { const next: unknown[] = []; at.push(next, 1, 'x'); at = next }
    const nested = await settle(walkPayload(deep, () => true, frozen(), { model: 'query', limit: 8 * 1024 * 1024 }))
    expect(nested).toMatchObject({ kind: 'measured', bytes: queryPayloadBytes(deep, 8 * 1024 * 1024) })
    if (nested.kind === 'measured') expect(nested.peakFrames).toBeLessThanOrEqual(nested.admittedFrames)
  })
  it('F3: a pool refusal after a sort committed still delivers the committed view', async () => {
    const document = workbookOf([{ sheetId: 's', cells: Array.from({ length: 2048 }, (_, row) => ({ row, col: 0, value: row })) }])
    for (const row of document.sheets[0].rows) row.heightPt = 25 + row.index % 2
    let pressure: { release(): void } | undefined
    const kernel: ViewerKernel = new ViewerKernel({ parse: async () => document, afterDispatch: command => {
      if (command.kind !== 'sort') return
      const memory = kernel.session('doc')!.memory, charge = memory.charge()
      expect(charge.set(memory.limitBytes - memory.totalBytes)).toBe(true)
      pressure = charge
    } })
    await kernel.handle({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
    const replies: WorkerReply[] = []
    try {
      await kernel.handle({ kind: 'sort', requestId: 'sort', owner: initialIdentity('doc', 's'), spec: null }, out => replies.push(out))
      expect(kernel.session('doc')!.isCurrent(initialIdentity('doc', 's'))).toBe(false)
      const last = replies.at(-1)!
      expect(last).toMatchObject({ kind: 'complete', result: { kind: 'view' } })
      // Measured without deduplication: never less than the exact size.
      expect(last.replyBytes).toBeGreaterThanOrEqual(payload.measurePayloadBytes({ ...last, replyBytes: 0 })!)
    } finally { pressure?.release(); await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
  })
  it('F4: main computes a 5,000-cell plan\'s bytes as it builds it, with no whole-plan walk', () => {
    const style = { fontSizePt: 11, bold: true }
    const plans = new WorkerViewportPlans({ styleOf: () => style } as never)
    const cells = Array.from({ length: 5000 }, (_, i) => ({ kind: 'cell', source: { row: Math.floor(i / 100), col: i % 100 }, bounds: { x: 0, y: 0, width: 64, height: 20 }, cell: { id: i, row: Math.floor(i / 100), col: i % 100, value: `v${i}`, styleIndex: 1 } }))
    const proof = new Map(cells.map(cell => [cell.source.row * 16384 + cell.source.col, { left: 0, right: 0 }]))
    const result = { kind: 'viewport', completeness: 'complete', content: { kind: 'detail', cells }, merges: [], rowLines: [], colLines: [] }
    const queryBytes = vi.mocked(queryPayloadBytes)
    queryBytes.mockClear()
    const { plan, bytes } = (plans as unknown as { decode(...args: unknown[]): { plan: { cells: unknown[] }; bytes: number } }).decode(result, { width: 100, height: 100 }, proof)
    expect(plan.cells).toHaveLength(5000)
    // Each measure covered one placement (or the cell-free shell), never the plan.
    const whole = queryBytes.mock.calls.filter(([value]) => (value as { cells?: unknown[] })?.cells?.length === 5000)
    expect(whole).toHaveLength(0)
    expect(queryBytes.mock.calls.length).toBe(5001)
    expect(bytes).toBe(queryPayloadBytes(plan, Number.POSITIVE_INFINITY))
  })
  it('F4: main accounts an 8,192-interval overview request in steps (RemoteSheetStore.queryOverview)', async () => {
    let visits = 0, first = -1
    const intervals = Array.from({ length: 8192 }, (_, i) => ({ get r0() { visits++; return i * 2 }, r1: i * 2 }))
    const request = { maxGroups: 1, groups: [{ groupId: 'g', rows: { kind: 'intervals' as const, intervals }, cols: [{ c0: 0, c1: 0 }] }] }
    const native = CheckpointBudget.prototype.pause
    vi.spyOn(CheckpointBudget.prototype, 'pause').mockImplementation(async function (this: CheckpointBudget) {
      if ((this as unknown as { phase: string }).phase === 'query' && first < 0) first = visits
      return native.call(this)
    })
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }])
    const source = createObjectWorkbookSource(doc, 'ov')
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source })
    const port = connectInProcess(kernel), client = new ViewerClient(port)
    await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'ov', source: { kind: 'object', document: doc }, profile: 'D' })
    const remote = new RemoteWorkbook(client, { docId: 'ov', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map([['s', kernel.session('ov')!.layout('s')]]) })
    try {
      const store = remote.store('s')
      queryPayloadBytesMock().mockClear()
      await store.queryOverview(request as never, { signal: new AbortController().signal, stamp: { ...store.stamp } })
      // The main side's first pause comes within 4,096 interval reads; no synchronous whole-request walk on main.
      expect(first).toBeGreaterThanOrEqual(0)
      expect(first).toBeLessThanOrEqual(4096)
      expect(queryPayloadBytesMock().mock.calls.filter(([value]) => value === request)).toHaveLength(0)
    } finally { await remote.release() }
  })
  it('F5: a stopped query walk rethrows StopWalk, and keyBytes reports stale, not quota', async () => {
    await expect(Promise.resolve().then(() => accountQueryBytes([1], 1024, () => { throw new StopWalk() }))).rejects.toBeInstanceOf(StopWalk)
    const document = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 1 }] }])
    const kernel = new ViewerKernel({ parse: async () => document })
    await kernel.handle({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document }, profile: 'D' }, () => {})
    try {
      const stopping = new CheckpointBudget(() => { throw new StopWalk() }, undefined, { now: () => 0, yield: async () => {} })
      expect(await (kernel as unknown as { keyBytes(...args: unknown[]): Promise<unknown> }).keyBytes(kernel.session('doc'), [1], stopping)).toBe('stale')
    } finally { await kernel.handle({ kind: 'close', requestId: 'close', docId: 'doc' }, () => {}) }
  })
})
