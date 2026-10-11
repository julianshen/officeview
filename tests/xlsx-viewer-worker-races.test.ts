/**
 * S4.2 cooperative scheduling, cancellation and atomic publication.
 *
 * Every interleaving is driven explicitly: gated source reads (abort-aware, like
 * a real store), an injected clock and an injected yield. No assertion depends
 * on elapsed time. Race scenarios record an event log (commands, test actions
 * and every worker reply in emission order) and assert the exact sequence; the
 * logs are rerun to prove they are reproducible.
 */
import { describe, expect, it } from 'vitest'
import type {
  FilterSpec, SortSpec, SourceStamp, SourceStore, ViewCommitDelta, WorkerCommand, WorkerReply,
} from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'
import { formatA1, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createIntervalSet, projectVisibility, projectVisibilitySteps, PROJECTION_CHUNK } from '../src/xlsx/view/visibility'
import { createRowOrder, pinRowPermutation } from '../src/xlsx/view/order'
import { YIELD_INTERVAL_MS, YIELD_WORK_UNITS } from '../src/xlsx/view/scheduler'
import { initialIdentity, initialWorkbookOwner } from '../src/worker/viewerProtocol'
import { sameStamp, type FilterEvaluator, type WorkbookSource } from '../src/worker/viewerSession'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { connectInProcess, createInThreadPort, ViewerKernel } from '../src/worker/viewerKernel'
import { createViewerClient, selectViewerHost, ViewerPublications } from '../src/worker/viewerClient'

type Value = number | string | null
function objectDoc(sheets: { sheetId: string; rows: Value[][] }[]): XlsxDocument {
  return {
    sheets: sheets.map(sheet => ({
      name: `Sheet ${sheet.sheetId}`, sheetId: sheet.sheetId,
      rows: sheet.rows.map((values, row) => ({
        index: row,
        cells: values.flatMap((value, col) => value === null ? [] : [{
          ref: formatA1({ row, col }), row, col, value, styleIndex: 0, hasCachedValue: true,
        }]),
      })),
      cols: [], merges: [], mergeRanges: [],
    })),
    tables: [],
  }
}
const column = (values: Value[]): Value[][] => values.map(value => [value])
const descending = (n: number): Value[][] => column(Array.from({ length: n }, (_v, i) => n - i))
const sortSpec = (r1: number, order: 'asc' | 'desc' = 'asc'): SortSpec => ({
  range: { r0: 0, r1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order }], caseSensitive: false,
})
const filterSpec = (r1: number): FilterSpec => ({
  range: { r0: 0, r1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]),
})
const settle = () => new Promise(resolve => setTimeout(resolve, 0))
async function until(condition: () => boolean, what = 'condition'): Promise<void> {
  for (let i = 0; i < 200; i++) {
    if (condition()) return
    await settle()
  }
  throw new Error(`Timed out waiting for ${what}`)
}

/** Abort-aware gate: a blocked read resumes on open() or when its request aborts. */
class Gate {
  hits = 0
  private release: (() => void) | null = null
  private blocked: Promise<void> | null = null
  close(): void { if (!this.blocked) this.blocked = new Promise(resolve => { this.release = resolve }) }
  open(): void { this.release?.(); this.release = null; this.blocked = null }
  get waiting(): boolean { return this.blocked !== null }
  async pass(signal?: AbortSignal): Promise<void> {
    this.hits++
    if (!this.blocked) return
    await Promise.race([this.blocked, new Promise<void>(resolve => {
      if (signal?.aborted) resolve()
      signal?.addEventListener('abort', () => resolve(), { once: true })
    })])
  }
}
/** Source with a gated key column and source-published value generations. */
function controlledSource(doc: XlsxDocument, docId: string) {
  const inner = createObjectWorkbookSource(doc, docId)
  const gate = new Gate()
  const gens = new Map<string, number>()
  let gateColumn: number | null = 0
  let releases = 0
  const stores = new Map<string, SourceStore>()
  const source: WorkbookSource = {
    valuePolicy: inner.valuePolicy, sheets: inner.sheets, geometry: id => inner.geometry(id),
    coverageDiagnostics: id => inner.coverageDiagnostics?.(id) ?? [],
    release: async () => { releases++; await inner.release() },
    store(sheetId) {
      let store = stores.get(sheetId)
      if (store) return store
      const base = inner.store(sheetId)
      const stamp = (): SourceStamp => Object.freeze({ ...base.stamp, valueGen: gens.get(sheetId) ?? 0 })
      const read = <T extends { stamp: SourceStamp }>(options: { stamp: SourceStamp; signal: AbortSignal }, call: (o: { stamp: SourceStamp; signal: AbortSignal }) => Promise<T>) => {
        const captured = stamp()
        const mapped = { ...options, stamp: options.stamp && sameStamp(options.stamp, captured) ? base.stamp : { ...base.stamp, valueGen: -1 } }
        return call(mapped).then(page => ({ ...page, stamp: captured }))
      }
      store = Object.create(base) as SourceStore
      Object.defineProperty(store, 'stamp', { get: stamp })
      store.queryColumn = async (request, options) => {
        if (request.col === gateColumn) await gate.pass(options.signal)
        return read(options, o => base.queryColumn(request, o))
      }
      store.readRows = (range, cursor, options) => read(options, o => base.readRows(range, cursor, o))
      store.queryStructural = (rect, options) => read(options, o => base.queryStructural(rect, o))
      stores.set(sheetId, store)
      return store
    },
  }
  return {
    source, gate,
    get releases() { return releases },
    gateOn(col: number | null) { gateColumn = col },
    publishValues(sheetId: string) { gens.set(sheetId, (gens.get(sheetId) ?? 0) + 1) },
  }
}
type Controlled = ReturnType<typeof controlledSource>

/** Deterministic scheduler: an injected clock and a yield that can hold a phase. */
class Steps {
  time = 0
  tick = 0
  checks = 0
  readonly phases: string[] = []
  readonly gaps: number[] = []
  private hold: string | null = null
  private held: (() => void)[] = []
  now = () => { this.checks++; return (this.time += this.tick) }
  yield = (phase: string): Promise<void> => {
    this.phases.push(phase)
    this.gaps.push(this.checks)
    this.checks = 0
    if (phase !== this.hold) return Promise.resolve()
    return new Promise<void>(resolve => this.held.push(resolve))
  }
  holdAt(phase: string | null): void { this.hold = phase }
  get waiting(): number { return this.held.length }
  releaseHeld(): void { this.hold = null; const held = this.held; this.held = []; for (const resume of held) resume() }
  count(phase: string): number { return this.phases.filter(entry => entry === phase).length }
}

/** Ordered event log: test actions and worker replies in emission order. */
class RaceLog {
  readonly events: string[] = []
  act(event: string): void { this.events.push(`> ${event}`) }
  sink = (out: WorkerReply): void => { this.events.push(describeReply(out)) }
}
function describeReply(out: WorkerReply): string {
  let detail = ''
  if (out.kind === 'progress') detail = ` ${out.completed}/${out.total ?? '?'}`
  else if (out.kind === 'complete' && out.result.kind === 'view') detail = ` rev${out.result.delta.owner.viewGen} g${out.result.delta.workbookOwner.globalViewGen}`
  else if (out.kind === 'error') detail = ` ${out.code}`
  else if (out.kind === 'cancelled' && out.diagnostics.length) detail = ` ${out.diagnostics.map(d => d.code).join(',')}`
  return `${out.requestId} ${out.commandKind}:${out.kind}${detail}`
}
function deltaOf(out: WorkerReply): ViewCommitDelta {
  if (out.kind !== 'complete' || out.result.kind !== 'view') throw new Error(`expected view, got ${describeReply(out)}`)
  return out.result.delta
}
function lastReply(log: RaceLog, requestId: string): string | undefined {
  return [...log.events].reverse().find(event => event.startsWith(`${requestId} `))
}

interface Rig { kernel: ViewerKernel; steps: Steps; sources: Controlled[]; log: RaceLog }
async function rig(doc: XlsxDocument, options: { filterEvaluator?: FilterEvaluator; tick?: number; afterDispatch?: (command: WorkerCommand, rig: Rig) => void | Promise<void> } = {}): Promise<Rig> {
  const steps = new Steps()
  steps.tick = options.tick ?? 0
  const sources: Controlled[] = []
  const log = new RaceLog()
  const self = {} as Rig
  const kernel = new ViewerKernel({
    parse: async () => doc,
    openSource: async (_handle, docId) => { const controlled = controlledSource(doc, docId); sources.push(controlled); return controlled.source },
    afterDispatch: options.afterDispatch ? command => options.afterDispatch!(command, self) : undefined,
  }, { now: () => steps.now(), yield: phase => steps.yield(phase), ...(options.filterEvaluator ? { filterEvaluator: options.filterEvaluator } : {}) })
  Object.assign(self, { kernel, steps, sources, log })
  await kernel.handle({ kind: 'open', requestId: 'open-1', docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }, log.sink)
  return self
}
const send = (r: Rig, command: WorkerCommand) => r.kernel.handle(command, r.log.sink)

describe('S4.2 cooperative yields on the shared budget', () => {
  it('shares the sort engine budget constants', () => {
    expect(YIELD_WORK_UNITS).toBe(4096)
    expect(YIELD_INTERVAL_MS).toBe(10)
  })

  it('preparation yields every 4,096 entries on a fixed clock', async () => {
    // 10,000 populated rows then one trailing blank run: 10,001 entries.
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]))
    await send(r, { kind: 'prepareColumns', requestId: 'p', owner: initialIdentity('doc', 'A'), cols: new Uint16Array([0]) })
    expect(lastReply(r.log, 'p')).toBe('p prepareColumns:complete')
    expect(r.steps.count('prepare')).toBe(2)
    expect(Math.max(...r.steps.gaps)).toBeLessThanOrEqual(YIELD_WORK_UNITS + 1)
  })

  it('preparation yields on the 10 ms clock before 4,096 entries, whichever first', async () => {
    // 3 ms per clock sample: the budget is spent on every 4th entry (12 ms).
    const slow = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]), { tick: 3 })
    await send(slow, { kind: 'prepareColumns', requestId: 'p', owner: initialIdentity('doc', 'A'), cols: new Uint16Array([0]) })
    expect(slow.steps.count('prepare')).toBe(Math.floor(10_001 / 4))
    // 10 ms per sample: every entry spends the budget.
    const slower = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]), { tick: 10 })
    await send(slower, { kind: 'prepareColumns', requestId: 'p', owner: initialIdentity('doc', 'A'), cols: new Uint16Array([0]) })
    expect(slower.steps.count('prepare')).toBe(10_001)
  })

  it('filter projection scans yield every 4,096 rows and equal the synchronous projection', async () => {
    const n = 20_000
    const matches = (row: number) => row % 2 === 0
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(n) }]), { filterEvaluator: async () => ({ kind: 'complete', matches }) })
    await send(r, { kind: 'filter', requestId: 'f', owner: initialIdentity('doc', 'A'), spec: filterSpec(n - 1), nowMs: 0 })
    expect(lastReply(r.log, 'f')).toBe('f filter:complete rev1 g1')
    expect(r.steps.count('project')).toBe(Math.floor(n / YIELD_WORK_UNITS))
    const session = r.kernel.session('doc')!
    const view = session.view('A')
    const reference = projectVisibility(session.source.geometry('A').baseHidden, filterSpec(n - 1), view.order, matches).visibility
    expect([...view.visibility.words]).toEqual([...reference.words])
  })

  it('sort guard and sort kernel sample the injected session clock and stay inside the budget', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]))
    await send(r, { kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortSpec(9_999) })
    expect(lastReply(r.log, 's')).toBe('s sort:complete rev1 g1')
    for (const phase of ['read', 'merge', 'publish']) expect(r.steps.count(phase)).toBeGreaterThan(0)
    // Fixed clock: only the work-unit budget can trigger a yield. The guard hands
    // over to the sort kernel at a yield, so no stretch exceeds one budget. A gap
    // counts clock samples: 4,096 work units plus the post-yield reset sample and,
    // at a kernel handoff, the next kernel's construction sample.
    expect(Math.max(...r.steps.gaps)).toBeLessThanOrEqual(YIELD_WORK_UNITS + 2)
    expect(r.steps.count('guard')).toBeGreaterThan(0)
  })
})

describe('S4.2 stale results never replace the current view', () => {
  async function sortAgainst(event: (r: Rig) => Promise<void>, doc = objectDoc([{ sheetId: 'A', rows: descending(5) }, { sheetId: 'B', rows: descending(3) }])) {
    const r = await rig(doc)
    r.sources[0].gate.close()
    r.log.act('sort A blocked on its key read')
    const sorting = send(r, { kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortSpec(4) })
    await until(() => r.sources[0].gate.hits > 0, 'key read')
    await event(r)
    r.log.act('release key read')
    r.sources[0].gate.open()
    await sorting
    return r
  }

  it('after another view commit (e.g. a selection draft) the delayed sort replies stale', async () => {
    const r = await sortAgainst(async r => {
      const session = r.kernel.session('doc')!
      const view = session.view('A')
      r.log.act('commit selection draft on A')
      await session.commitView('A', { order: view.order, visibility: view.visibility, sort: null, filter: null, selection: { active: { row: 1, col: 0 }, anchor: { row: 1, col: 0 }, ranges: [] } })
    })
    expect(r.log.events).toEqual([
      'open-1 open:complete',
      '> sort A blocked on its key read',
      's sort:progress 0/3',
      's sort:progress 1/3',
      '> commit selection draft on A',
      '> release key read',
      's sort:error stale-source',
    ])
    const view = r.kernel.session('doc')!.view('A')
    expect(view.revision).toBe(1)
    expect(view.order.kind).toBe('identity')
    expect(view.selection.active).toEqual({ row: 1, col: 0 })
  })

  it('a sheet switch is not a cancel: the sort still commits its own sheet and leaves the other untouched', async () => {
    const r = await sortAgainst(async r => {
      r.log.act('activate B')
      await send(r, { kind: 'activateSheet', requestId: 'act', owner: initialWorkbookOwner('doc', [{ sheetId: 'A' }, { sheetId: 'B' }]), sheetId: 'B' })
    })
    expect(r.log.events).toEqual([
      'open-1 open:complete',
      '> sort A blocked on its key read',
      's sort:progress 0/3',
      's sort:progress 1/3',
      '> activate B',
      'act activateSheet:complete',
      '> release key read',
      's sort:progress 2/3',
      's sort:progress 3/3',
      's sort:complete rev1 g1',
    ])
    const session = r.kernel.session('doc')!
    expect(session.activeSheetId).toBe('B')
    expect(session.view('B')).toMatchObject({ revision: 0 })
    expect(session.view('B').order.kind).toBe('identity')
    expect([...session.view('A').order.bands[0].sourceAtOffset]).toEqual([4, 3, 2, 1, 0])
  })

  it('after a value publication the delayed sort replies stale and publishes nothing', async () => {
    const r = await sortAgainst(async r => {
      r.log.act('publish values for A')
      r.sources[0].publishValues('A')
      r.kernel.session('doc')!.publishValues('A')
    })
    expect(lastReply(r.log, 's')).toBe('s sort:error stale-source')
    const session = r.kernel.session('doc')!
    expect(session.view('A')).toMatchObject({ revision: 0, valueGen: 1 })
    expect(session.view('A').order.kind).toBe('identity')
    expect(session.identity('A')).toEqual({ docId: 'doc', sheetId: 'A', valueGen: 1, fontGen: 0, viewGen: 0 })
    expect(session.workbookOwner()).toMatchObject({ globalViewGen: 0, workbookValueGen: 1 })
  })

  it('after a font change the delayed sort replies stale; the font owner advances without a view revision', async () => {
    const r = await sortAgainst(async r => {
      r.log.act('publish fonts')
      r.kernel.session('doc')!.publishFonts()
    })
    expect(r.log.events.slice(-2)).toEqual(['> release key read', 's sort:error stale-source'])
    const session = r.kernel.session('doc')!
    expect(session.identity('A')).toEqual({ docId: 'doc', sheetId: 'A', valueGen: 0, fontGen: 1, viewGen: 0 })
    expect(session.globalViewGen).toBe(0)
  })

  it('after close the delayed sort is cancelled with the document and the source is released once', async () => {
    const r = await sortAgainst(async r => {
      r.log.act('close')
      await send(r, { kind: 'close', requestId: 'c', docId: 'doc' })
    })
    expect(r.log.events).toEqual([
      'open-1 open:complete',
      '> sort A blocked on its key read',
      's sort:progress 0/3',
      's sort:progress 1/3',
      '> close',
      'c close:complete',
      '> release key read',
      's sort:cancelled document-closed',
    ])
    expect(r.sources[0].releases).toBe(1)
    expect(r.kernel.session('doc')).toBeUndefined()
  })

  it('after close and reopen of the same docId the old request never publishes into the new epoch', async () => {
    const r = await sortAgainst(async r => {
      r.log.act('close and reopen')
      await send(r, { kind: 'close', requestId: 'c', docId: 'doc' })
      await send(r, { kind: 'open', requestId: 'open-2', docId: 'doc', source: { kind: 'object', document: objectDoc([]) }, profile: 'D' })
    })
    expect(lastReply(r.log, 's')).toBe('s sort:cancelled document-closed')
    const session = r.kernel.session('doc')!
    expect(session.epoch).toBe(2)
    expect(session.globalViewGen).toBe(0)
    expect(session.view('A').order.kind).toBe('identity')
  })

  it('a completion already committed in a closed epoch is withheld even when the docId reopened before emission', async () => {
    let reopened = false
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(5) }]), {
      afterDispatch: async (command, r) => {
        if (command.kind !== 'sort' || reopened) return
        reopened = true
        r.log.act('close and reopen before emission')
        await send(r, { kind: 'close', requestId: 'c', docId: 'doc' })
        await send(r, { kind: 'open', requestId: 'open-2', docId: 'doc', source: { kind: 'object', document: objectDoc([]) }, profile: 'D' })
      },
    })
    await send(r, { kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortSpec(4) })
    expect(r.log.events).toEqual([
      'open-1 open:complete',
      's sort:progress 0/3',
      's sort:progress 1/3',
      's sort:progress 2/3',
      's sort:progress 3/3',
      '> close and reopen before emission',
      'c close:complete',
      'open-2 open:complete',
      's sort:error stale-source',
    ])
    expect(r.kernel.session('doc')!.globalViewGen).toBe(0)
  })

  it('a delayed value-index page after a view commit replies stale', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(5) }]))
    r.sources[0].gateOn(0)
    r.sources[0].gate.close()
    const owner = initialIdentity('doc', 'A')
    const query = send(r, { kind: 'valueQuery', requestId: 'q', owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, limit: 8 } })
    await until(() => r.sources[0].gate.hits > 0, 'value read')
    r.sources[0].gateOn(null)
    r.log.act('filter commits')
    await send(r, { kind: 'filter', requestId: 'f', owner, spec: null, nowMs: 0 })
    r.sources[0].gate.open()
    await query
    expect(r.log.events.slice(1)).toEqual([
      '> filter commits',
      'f filter:progress 0/3',
      'f filter:progress 1/3',
      'f filter:progress 2/3',
      'f filter:progress 3/3',
      'f filter:complete rev1 g1',
      'q valueQuery:error stale-source',
    ])
  })

  it('race logs are reproducible: the same scenario yields the same event sequence', async () => {
    const scenario = async () => (await sortAgainst(async r => {
      r.log.act('activate B')
      await send(r, { kind: 'activateSheet', requestId: 'act', owner: initialWorkbookOwner('doc', [{ sheetId: 'A' }, { sheetId: 'B' }]), sheetId: 'B' })
    })).log.events
    expect(await scenario()).toEqual(await scenario())
  })
})

describe('S4.2 cancellation and reclaim', () => {
  it('cancelling a sort preserves the previous committed view', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]))
    const first = deltaOf((await collect(r, { kind: 'sort', requestId: 's1', owner: initialIdentity('doc', 'A'), spec: sortSpec(9_999) })))
    const before = r.kernel.session('doc')!.view('A')
    r.steps.holdAt('merge')
    const second = send(r, { kind: 'sort', requestId: 's2', owner: first.owner, spec: sortSpec(9_999, 'desc') })
    await until(() => r.steps.waiting > 0, 'merge yield')
    r.log.act('cancel s2')
    await send(r, { kind: 'cancel', requestId: 'x', targetRequestId: 's2' })
    r.steps.releaseHeld()
    await second
    expect(r.log.events.slice(-3)).toEqual(['> cancel s2', 'x cancel:cancelled', 's2 sort:cancelled cancel-requested'])
    const after = r.kernel.session('doc')!.view('A')
    expect(after.revision).toBe(before.revision)
    expect(after.order).toBe(before.order)
    expect(after.visibility).toBe(before.visibility)
  })

  it('a superseding draft on the same sheet cancels the older draft, which never commits', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(5) }]))
    r.sources[0].gate.close()
    const owner = initialIdentity('doc', 'A')
    const older = send(r, { kind: 'sort', requestId: 's1', owner, spec: sortSpec(4, 'desc') })
    await until(() => r.sources[0].gate.hits > 0, 'older key read')
    r.sources[0].gateOn(null)
    r.log.act('superseding sort')
    await send(r, { kind: 'sort', requestId: 's2', owner, spec: sortSpec(4) })
    r.sources[0].gate.open()
    await older
    expect(r.log.events.slice(1)).toEqual([
      's1 sort:progress 0/3',
      's1 sort:progress 1/3',
      '> superseding sort',
      's2 sort:progress 0/3',
      's2 sort:progress 1/3',
      's1 sort:cancelled superseded',
      's2 sort:progress 2/3',
      's2 sort:progress 3/3',
      's2 sort:complete rev1 g1',
    ])
    expect([...r.kernel.session('doc')!.view('A').order.bands[0].sourceAtOffset]).toEqual([4, 3, 2, 1, 0])
  })

  it('a stale command never supersedes a current in-flight draft', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(5) }]))
    const committed = deltaOf(await collect(r, { kind: 'filter', requestId: 'f', owner: initialIdentity('doc', 'A'), spec: null, nowMs: 0 }))
    r.sources[0].gate.close()
    const current = send(r, { kind: 'sort', requestId: 's1', owner: committed.owner, spec: sortSpec(4) })
    await until(() => r.sources[0].gate.hits > 0, 'key read')
    await send(r, { kind: 'sort', requestId: 's-stale', owner: initialIdentity('doc', 'A'), spec: sortSpec(4, 'desc') })
    r.sources[0].gate.open()
    await current
    expect(lastReply(r.log, 's-stale')).toBe('s-stale sort:error stale-source')
    expect(lastReply(r.log, 's1')).toBe('s1 sort:complete rev2 g2')
  })

  it('range replacement cancels the older value query on the same column', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(30) }]))
    r.sources[0].gate.close()
    const owner = initialIdentity('doc', 'A')
    const older = send(r, { kind: 'valueQuery', requestId: 'q1', owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 9 }] }, limit: 20 } })
    await until(() => r.sources[0].gate.hits > 0, 'first range')
    r.sources[0].gateOn(null)
    r.log.act('replace range')
    await send(r, { kind: 'valueQuery', requestId: 'q2', owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 5, r1: 20 }] }, limit: 20 } })
    // A different column is an independent query, not a replacement.
    await send(r, { kind: 'valueQuery', requestId: 'q3', owner, query: { col: 1, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1 }] }, limit: 4 } })
    r.sources[0].gate.open()
    await older
    expect(r.log.events.slice(1)).toEqual([
      '> replace range',
      'q2 valueQuery:complete',
      'q1 valueQuery:cancelled range-replaced',
      'q3 valueQuery:complete',
    ])
  })

  it('cancelling a preparation reclaims its staged bytes and leaves the working set unchanged', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]))
    const session = r.kernel.session('doc')!
    r.steps.holdAt('prepare')
    const preparing = send(r, { kind: 'prepareColumns', requestId: 'p', owner: initialIdentity('doc', 'A'), cols: new Uint16Array([0]) })
    await until(() => r.steps.waiting > 0, 'prepare yield')
    const charged = session.memoryStats().transientBytes
    expect(charged).toBeGreaterThan(4096 * 22)
    await send(r, { kind: 'cancel', requestId: 'x', targetRequestId: 'p' })
    r.steps.releaseHeld()
    await preparing
    expect(lastReply(r.log, 'p')).toBe('p prepareColumns:cancelled cancel-requested')
    expect(session.memoryStats()).toMatchObject({ transientBytes: 0, preparedBytes: 0 })
    expect(session.memoryStats().peakTransientBytes).toBeGreaterThanOrEqual(charged)
    expect(session.preparedStats()).toEqual({ columns: 0, bytes: 0 })
  })

  it('a superseded sort reclaims its private draft (order, inverse and mask) before anything commits', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(5_000) }]))
    const session = r.kernel.session('doc')!
    r.steps.holdAt('project')
    const owner = initialIdentity('doc', 'A')
    const older = send(r, { kind: 'sort', requestId: 's1', owner, spec: sortSpec(4_999) })
    await until(() => r.steps.waiting > 0, 'projection yield')
    // Order + inverse (2 x 4 B per row) plus the private mask draft are charged.
    expect(session.memoryStats().transientBytes).toBeGreaterThanOrEqual(5_000 * 8)
    expect(session.view('A').revision).toBe(0)
    await send(r, { kind: 'sort', requestId: 's2', owner, spec: null })
    r.steps.releaseHeld()
    await older
    expect(lastReply(r.log, 's1')).toBe('s1 sort:cancelled superseded')
    expect(lastReply(r.log, 's2')).toBe('s2 sort:complete rev1 g1')
    expect(session.memoryStats().transientBytes).toBe(0)
  })

  it('closing the document cancels every in-flight job and reclaims their charges', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]))
    const session = r.kernel.session('doc')!
    r.steps.holdAt('prepare')
    const preparing = send(r, { kind: 'prepareColumns', requestId: 'p', owner: initialIdentity('doc', 'A'), cols: new Uint16Array([0]) })
    await until(() => r.steps.waiting > 0, 'prepare yield')
    await send(r, { kind: 'close', requestId: 'c', docId: 'doc' })
    r.steps.releaseHeld()
    await preparing
    expect(lastReply(r.log, 'p')).toBe('p prepareColumns:cancelled document-closed')
    expect(session.memoryStats().transientBytes).toBe(0)
  })
})

describe('S4.2 one view transaction', () => {
  it('every scheduler step observes the whole old view or the whole new view, never a mix', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(9_000) }]), { filterEvaluator: async () => ({ kind: 'complete', matches: row => row % 3 !== 0 }) })
    const session = r.kernel.session('doc')!
    const snapshots: string[] = []
    const fingerprint = () => {
      const view = session.view('A'), owner = session.workbookOwner()
      return JSON.stringify([view.revision, owner.globalViewGen, view.order.kind, view.visibility.rank(MAX_ROWS), view.counts, view.filter !== null, view.sort !== null])
    }
    const observe = r.steps.yield
    r.steps.yield = phase => { snapshots.push(fingerprint()); return observe(phase) }
    const base = fingerprint()
    const sorted = deltaOf(await collect(r, { kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortSpec(8_999) }))
    const afterSort = fingerprint()
    deltaOf(await collect(r, { kind: 'filter', requestId: 'f', owner: sorted.owner, spec: filterSpec(8_999), nowMs: 0 }))
    const afterFilter = fingerprint()
    expect(new Set(snapshots)).toEqual(new Set([base, afterSort]))
    expect(snapshots.indexOf(afterSort)).toBeGreaterThan(snapshots.lastIndexOf(base))
    expect(afterFilter).not.toBe(afterSort)
    expect(session.view('A').counts).toEqual({ visible: MAX_ROWS - 3_000, total: MAX_ROWS })
  })

  it('commitDraft is one validated transaction: a stale or closed base changes nothing', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(3) }]))
    const session = r.kernel.session('doc')!
    const view = session.view('A')
    const draft = { order: view.order, visibility: view.visibility, sort: null, filter: null }
    const committed = session.commitDraft(initialIdentity('doc', 'A'), draft)
    expect(committed?.identity.viewGen).toBe(1)
    expect(session.commitDraft(initialIdentity('doc', 'A'), draft)).toBeNull()
    expect(session.view('A').revision).toBe(1)
    await session.release()
    expect(session.commitDraft(session.identity('A'), draft)).toBeNull()
  })

  it('the committed delta carries the counts committed with the view', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10) }]), { filterEvaluator: async () => ({ kind: 'complete', matches: row => row < 4 }) })
    const delta = deltaOf(await collect(r, { kind: 'filter', requestId: 'f', owner: initialIdentity('doc', 'A'), spec: filterSpec(9), nowMs: 0 }))
    expect(delta.counts).toEqual(r.kernel.session('doc')!.view('A').counts)
    expect(delta.counts).toEqual({ visible: MAX_ROWS - 6, total: MAX_ROWS })
  })
})

describe('S4.2 progress from completed authoritative batches', () => {
  it('preparation reports rows covered by completed pages, monotonic, before its terminal reply', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]))
    await send(r, { kind: 'prepareColumns', requestId: 'p', owner: initialIdentity('doc', 'A'), cols: new Uint16Array([0]) })
    expect(r.log.events.slice(1)).toEqual([
      `p prepareColumns:progress 0/${MAX_ROWS}`,
      `p prepareColumns:progress 4096/${MAX_ROWS}`,
      `p prepareColumns:progress 8192/${MAX_ROWS}`,
      `p prepareColumns:progress ${MAX_ROWS}/${MAX_ROWS}`,
      'p prepareColumns:complete',
    ])
  })

  it('no progress is published after cancellation or after the owner changed', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]))
    r.steps.holdAt('prepare')
    const preparing = send(r, { kind: 'prepareColumns', requestId: 'p', owner: initialIdentity('doc', 'A'), cols: new Uint16Array([0]) })
    await until(() => r.steps.waiting > 0, 'prepare yield')
    r.log.act('publish fonts')
    r.kernel.session('doc')!.publishFonts()
    r.steps.releaseHeld()
    await preparing
    expect(r.log.events.slice(1)).toEqual([
      `p prepareColumns:progress 0/${MAX_ROWS}`,
      '> publish fonts',
      'p prepareColumns:error stale-source',
    ])
  })

  it('input is acknowledged immediately: other commands are answered while a long job is held', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }, { sheetId: 'B', rows: descending(2) }]))
    r.steps.holdAt('merge')
    const sorting = send(r, { kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortSpec(9_999) })
    await until(() => r.steps.waiting > 0, 'merge yield')
    r.log.act('input while sorting')
    await send(r, { kind: 'activateSheet', requestId: 'act', owner: initialWorkbookOwner('doc', [{ sheetId: 'A' }, { sheetId: 'B' }]), sheetId: 'B' })
    await send(r, { kind: 'valueQuery', requestId: 'q', owner: initialIdentity('doc', 'B'), query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1 }] }, limit: 4 } })
    r.log.act('release sort')
    r.steps.releaseHeld()
    await sorting
    expect(r.log.events.slice(1)).toEqual([
      's sort:progress 0/3',
      's sort:progress 1/3',
      '> input while sorting',
      'act activateSheet:complete',
      'q valueQuery:complete',
      '> release sort',
      's sort:progress 2/3',
      's sort:progress 3/3',
      's sort:complete rev1 g1',
    ])
  })

  it('the client surfaces progress through onProgress; progress never settles a request', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]))
    const port = connectInProcess(r.kernel)
    const client = createViewerClient(port, 'c')
    const seen: string[] = []
    const id = client.nextRequestId()
    const off = client.onProgress(id, progress => seen.push(`${progress.completed}/${progress.total}`))
    client.onProgress('never-sent', () => seen.push('foreign'))
    let settled = false
    r.steps.holdAt('prepare')
    const pending = client.send({ kind: 'prepareColumns', requestId: id, owner: initialIdentity('doc', 'A'), cols: new Uint16Array([0]) })
    void pending.then(() => { settled = true })
    await until(() => r.steps.waiting > 0, 'prepare yield')
    await settle()
    // Held inside the first page: only the acknowledgment has been reported.
    expect(seen).toEqual([`0/${MAX_ROWS}`])
    expect(settled).toBe(false)
    r.steps.releaseHeld()
    expect((await pending).kind).toBe('complete')
    expect(seen).toEqual([`0/${MAX_ROWS}`, `4096/${MAX_ROWS}`, `8192/${MAX_ROWS}`, `${MAX_ROWS}/${MAX_ROWS}`])
    off()
    client.dispose()
  })
})

describe('S4.2 client publication state', () => {
  const sheets = [{ sheetId: 'A' }, { sheetId: 'B' }]
  const viewReply = (requestId: string, viewGen: number, globalViewGen: number, visible = MAX_ROWS, docId = 'doc'): WorkerReply<'sort'> => ({
    kind: 'complete', commandKind: 'sort', requestId, owner: initialIdentity(docId, 'A'), diagnostics: [],
    result: { kind: 'view', delta: {
      bands: [], visibility: [], metrics: { rows: [], cols: [] }, counts: { visible, total: MAX_ROWS },
      selection: { active: null, anchor: null, ranges: [] },
      owner: { docId, sheetId: 'A', valueGen: 0, fontGen: 0, viewGen },
      workbookOwner: { docId, globalViewGen, workbookValueGen: 0, fontGen: 0, sheetVector: [] },
    } },
  })
  const cancelled = (requestId: string): WorkerReply<'sort'> => ({ kind: 'cancelled', commandKind: 'sort', requestId, owner: initialIdentity('doc', 'A'), diagnostics: [] })

  it('input is acknowledged at once as pending and the prior view stays visible until a newer commit', () => {
    const pubs = new ViewerPublications()
    pubs.open('doc', sheets)
    const t1 = pubs.begin('doc', 'A', 's1')
    expect(pubs.get('doc', 'A')!.status).toEqual({ kind: 'pending', requestId: 's1', progress: null })
    expect(pubs.get('doc', 'A')!.counts).toBeNull()
    expect(pubs.progress(t1, { completed: 1, total: 3 })).toBe(true)
    expect(pubs.settle(t1, viewReply('s1', 1, 1, 100))).toBe('applied')
    expect(pubs.get('doc', 'A')).toMatchObject({ status: { kind: 'ready' }, counts: { visible: 100 }, owner: { viewGen: 1 } })
    // A newer request is pending; an older delayed completion arrives late.
    const t3 = pubs.begin('doc', 'A', 's3')
    expect(pubs.settle(t3, viewReply('s3', 3, 3, 300))).toBe('applied')
    const late = pubs.begin('doc', 'A', 's2')
    const t4 = pubs.begin('doc', 'A', 's4')
    expect(pubs.settle(late, viewReply('s2', 2, 2, 200))).toBe('ignored')
    expect(pubs.progress(late, { completed: 2, total: 3 })).toBe(false)
    expect(pubs.get('doc', 'A')).toMatchObject({ counts: { visible: 300 }, owner: { viewGen: 3 }, status: { kind: 'pending', requestId: 's4' } })
    // Cancellation keeps the prior committed view and records one notice.
    expect(pubs.settle(t4, cancelled('s4'))).toBe('status')
    expect(pubs.get('doc', 'A')).toMatchObject({ counts: { visible: 300 }, status: { kind: 'ready' }, notice: { requestId: 's4', kind: 'cancelled' } })
  })

  it('search geometry is owned by the exact view and font generation', () => {
    const pubs = new ViewerPublications()
    pubs.open('doc', sheets)
    const t1 = pubs.begin('doc', 'A', 's1')
    pubs.settle(t1, viewReply('s1', 1, 1))
    const owner1 = pubs.get('doc', 'A')!.owner
    expect(pubs.publishGeometry('doc', 'A', owner1, pubs.fontGen('doc'), { hits: 1 })).toBe(true)
    // A font change invalidates geometry but never order or counts.
    pubs.fontChanged('doc')
    expect(pubs.get('doc', 'A')).toMatchObject({ geometry: null, owner: { viewGen: 1 } })
    expect(pubs.publishGeometry('doc', 'A', owner1, 0, { hits: 'old font' })).toBe(false)
    expect(pubs.publishGeometry('doc', 'A', owner1, 1, { hits: 2 })).toBe(true)
    // An applied view change invalidates geometry-dependent search layout.
    const t2 = pubs.begin('doc', 'A', 's2')
    pubs.settle(t2, viewReply('s2', 2, 2))
    expect(pubs.get('doc', 'A')!.geometry).toBeNull()
    expect(pubs.publishGeometry('doc', 'A', owner1, 1, { hits: 'old view' })).toBe(false)
  })

  it('a reopened docId starts a new epoch; replies from the previous epoch are ignored', () => {
    const pubs = new ViewerPublications()
    expect(pubs.open('doc', sheets)).toBe(1)
    const old = pubs.begin('doc', 'A', 's1')
    pubs.close('doc')
    expect(pubs.get('doc', 'A')).toBeUndefined()
    expect(pubs.open('doc', sheets)).toBe(2)
    // Identical owner numbers, but the ticket belongs to epoch 1.
    expect(pubs.settle(old, viewReply('s1', 1, 1))).toBe('ignored')
    expect(pubs.get('doc', 'A')).toMatchObject({ epoch: 2, owner: { viewGen: 0 }, status: { kind: 'ready' } })
  })
})

describe('S4.2 no-worker support', () => {
  const port = { postMessage: () => {}, addEventListener: () => {}, removeEventListener: () => {} }
  it('large mode refuses missing mandatory worker/storage features with a typed reason', () => {
    expect(selectViewerHost({ mode: 'large', worker: null, persistentStorage: true, inThread: () => port })).toEqual({
      kind: 'refused', code: 'unsupported', reason: 'worker-required', message: expect.any(String),
    })
    expect(selectViewerHost({ mode: 'large', worker: () => port, persistentStorage: false, inThread: () => port })).toEqual({
      kind: 'refused', code: 'unsupported', reason: 'storage-required', message: expect.any(String),
    })
    expect(selectViewerHost({ mode: 'large', worker: () => port, persistentStorage: true, inThread: () => port })).toMatchObject({ kind: 'worker' })
  })

  it('small admitted mode falls back to the same chunked kernels in-thread', async () => {
    let inThread = 0
    const host = selectViewerHost({ mode: 'small', worker: null, persistentStorage: false, inThread: () => { inThread++; return port } })
    expect(host).toMatchObject({ kind: 'in-thread' })
    expect(inThread).toBe(1)
    expect(selectViewerHost({ mode: 'small', worker: () => port, persistentStorage: false, inThread: () => port })).toMatchObject({ kind: 'worker' })
    // The in-thread port runs the resident kernel with the shared cooperative budget.
    const steps = new Steps()
    const doc = objectDoc([{ sheetId: 'A', rows: descending(5_000) }])
    const client = createViewerClient(createInThreadPort({ parse: async () => doc }, { now: steps.now, yield: steps.yield }))
    expect((await client.send({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })).kind).toBe('complete')
    expect(deltaOf(await client.send({ kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortSpec(4_999) })).owner.viewGen).toBe(1)
    expect(steps.count('merge')).toBeGreaterThan(0)
    // Paged (large) sources need the worker/storage stack: a typed refusal, not a crash.
    const paged = await client.send({ kind: 'open', requestId: 'o2', docId: 'big', source: { kind: 'paged', handleId: 'h' }, profile: 'D' })
    expect(paged).toMatchObject({ kind: 'error', code: 'unsupported', diagnostics: [{ code: 'worker-required' }] })
    client.dispose()
  })
})

describe('S4.2 review fixes', () => {
  it('pi#1 preparation refuses column pages out of ascending contiguous row order (typed page-error, nothing staged)', async () => {
    const doc = objectDoc([{ sheetId: 'A', rows: descending(20) }])
    for (const mangle of [
      (items: readonly unknown[]) => [...items].reverse(),
      (items: readonly unknown[]) => [items[0], ...items],
    ]) {
      const inner = createObjectWorkbookSource(doc, 'doc')
      const base = inner.store('A')
      const store = Object.create(base) as SourceStore
      store.queryColumn = async (request, options) => {
        const page = await base.queryColumn(request, options)
        return { ...page, items: mangle(page.items) as typeof page.items }
      }
      const source: WorkbookSource = { ...inner, sheets: inner.sheets, geometry: id => inner.geometry(id), store: () => store, release: () => inner.release() }
      const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
      const log = new RaceLog()
      await kernel.handle({ kind: 'open', requestId: 'o', docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' }, log.sink)
      await kernel.handle({ kind: 'prepareColumns', requestId: 'p', owner: initialIdentity('doc', 'A'), cols: new Uint16Array([0]) }, log.sink)
      expect(lastReply(log, 'p')).toBe('p prepareColumns:error page-error')
      expect(kernel.session('doc')!.preparedStats()).toEqual({ columns: 0, bytes: 0 })
      expect(kernel.session('doc')!.memoryStats().transientBytes).toBe(0)
    }
  })

  it('pi#2 the filter geometry guard steps the shared budget: one unit per mask word, changed slot, source word and band', async () => {
    // Filter rows 0..1 hiding row 0: 32,768 mask words compared, 1 changed slot,
    // 1 source word scanned, 1 covered band queried, no drawings.
    const guardUnits = MAX_ROWS / 32 + 1 + 1 + 1
    const filter = { filterEvaluator: (async () => ({ kind: 'complete', matches: (row: number) => row !== 0 })) as FilterEvaluator }
    const fixed = await rig(objectDoc([{ sheetId: 'A', rows: descending(2) }]), filter)
    await send(fixed, { kind: 'filter', requestId: 'f', owner: initialIdentity('doc', 'A'), spec: filterSpec(1), nowMs: 0 })
    expect(lastReply(fixed.log, 'f')).toBe('f filter:complete rev1 g1')
    // The 2 projected rows share the operation budget with the guard.
    expect(fixed.steps.count('guard')).toBe(Math.floor((2 + guardUnits) / YIELD_WORK_UNITS))
    expect(Math.max(...fixed.steps.gaps)).toBeLessThanOrEqual(YIELD_WORK_UNITS + 1)
    const slow = await rig(objectDoc([{ sheetId: 'A', rows: descending(2) }]), { ...filter, tick: YIELD_INTERVAL_MS })
    await send(slow, { kind: 'filter', requestId: 'f', owner: initialIdentity('doc', 'A'), spec: filterSpec(1), nowMs: 0 })
    expect(slow.steps.count('guard')).toBe(guardUnits)
  })

  it('pi#2 a cancel during the geometry guard stops it at its next yield', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(2) }]), { filterEvaluator: async () => ({ kind: 'complete', matches: row => row !== 0 }) })
    r.steps.holdAt('guard')
    const filtering = send(r, { kind: 'filter', requestId: 'f', owner: initialIdentity('doc', 'A'), spec: filterSpec(1), nowMs: 0 })
    await until(() => r.steps.waiting > 0, 'guard yield')
    await send(r, { kind: 'cancel', requestId: 'x', targetRequestId: 'f' })
    const before = r.steps.count('guard')
    r.steps.releaseHeld()
    await filtering
    expect(lastReply(r.log, 'f')).toBe('f filter:cancelled cancel-requested')
    expect(r.steps.count('guard')).toBe(before)
    expect(r.kernel.session('doc')!.view('A').revision).toBe(0)
  })

  it('pi#3 a sort/filter completion is withheld when a newer commit superseded its delta before emission', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(5) }]), {
      afterDispatch: async (command, r) => {
        if (command.kind !== 'sort') return
        const session = r.kernel.session('doc')!
        const view = session.view('A')
        r.log.act('commit selection draft before emission')
        await session.commitView('A', { order: view.order, visibility: view.visibility, sort: view.sort, filter: view.filter })
      },
    })
    await send(r, { kind: 'sort', requestId: 's', owner: initialIdentity('doc', 'A'), spec: sortSpec(4) })
    expect(r.log.events.slice(-2)).toEqual(['> commit selection draft before emission', 's sort:error stale-source'])
  })

  it('pi#4 a queued preparation acknowledges only once its transaction is admitted', async () => {
    const r = await rig(objectDoc([{ sheetId: 'A', rows: descending(10_000) }]))
    r.steps.holdAt('prepare')
    const owner = initialIdentity('doc', 'A')
    const first = send(r, { kind: 'prepareColumns', requestId: 'p1', owner, cols: new Uint16Array([0]) })
    await until(() => r.steps.waiting > 0, 'prepare yield')
    const second = send(r, { kind: 'prepareColumns', requestId: 'p2', owner, cols: new Uint16Array([1]) })
    await settle()
    r.log.act('release p1')
    r.steps.releaseHeld()
    await Promise.all([first, second])
    expect(r.log.events.slice(1)).toEqual([
      `p1 prepareColumns:progress 0/${MAX_ROWS}`,
      '> release p1',
      `p1 prepareColumns:progress 4096/${MAX_ROWS}`,
      `p1 prepareColumns:progress 8192/${MAX_ROWS}`,
      `p1 prepareColumns:progress ${MAX_ROWS}/${MAX_ROWS}`,
      // p2 is admitted when p1's transaction ends, just before p1's reply is emitted.
      `p2 prepareColumns:progress 0/${MAX_ROWS}`,
      'p1 prepareColumns:complete',
      `p2 prepareColumns:progress ${MAX_ROWS}/${MAX_ROWS}`,
      'p2 prepareColumns:complete',
    ])
  })

  it('opencode#1 a cancellation notice keeps the worker cancel reason', () => {
    const pubs = new ViewerPublications()
    pubs.open('doc', [{ sheetId: 'A' }])
    const ticket = pubs.begin('doc', 'A', 's1')
    pubs.settle(ticket, { kind: 'cancelled', commandKind: 'sort', requestId: 's1', owner: initialIdentity('doc', 'A'), diagnostics: [{ code: 'superseded', message: 'newer draft' }] })
    expect(pubs.get('doc', 'A')!.notice).toEqual({ requestId: 's1', kind: 'cancelled', code: 'superseded' })
  })

  it('opencode#2 the synchronous projection and the cooperative steps are one implementation', () => {
    const n = 10_000
    const order = createRowOrder([pinRowPermutation(0, n - 1, () => false, Uint32Array.from({ length: n }, (_v, i) => n - 1 - i))])
    const hidden = createIntervalSet([{ r0: 3, r1: 7 }])
    const spec = filterSpec(n - 1)
    const matches = (row: number) => row % 5 !== 0
    const steps = projectVisibilitySteps(hidden, spec, order, matches)
    const chunks: number[] = []
    let next = steps.next()
    for (; !next.done; next = steps.next()) chunks.push(next.value)
    // Band rows plus filter body rows, reported in bounded chunks.
    expect(chunks.reduce((sum, units) => sum + units, 0)).toBe(2 * n)
    expect(Math.max(...chunks)).toBeLessThanOrEqual(PROJECTION_CHUNK)
    expect(YIELD_WORK_UNITS % PROJECTION_CHUNK).toBe(0)
    expect([...next.value.visibility.words]).toEqual([...projectVisibility(hidden, spec, order, matches).visibility.words])
  })
})

async function collect(r: Rig, command: WorkerCommand): Promise<WorkerReply> {
  const replies: WorkerReply[] = []
  await r.kernel.handle(command, out => { replies.push(out); r.log.sink(out) })
  return replies[replies.length - 1]
}
