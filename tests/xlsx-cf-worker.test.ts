import { describe, expect, it } from 'vitest'
import { createCfMemory, createCfWorkBudget, type CfWorkBudget } from '../src/xlsx/cf/limits'
import type { CfSheetRules, CfStamp, CfLiteral, CfValueSource, CfValuePage, CfLease } from '../src/xlsx/cf/model'
import type { SourceRect } from '../src/xlsx/view/coordinates'
import {
  estimateCfPreparationBytes,
  measureCfCellBytes,
  measureCfDefinitionBytes,
  measureCfDiagnosticBytes,
  measureCfPropertiesBytes,
  cfRulePropertyBytes,
  versionCfRules,
  CF_REPLY_ENVELOPE_BYTES,
  type CfReadRequest,
  type CfReadResult,
  type CfReadStamp,
  type CfCellStateWire,
  type CfSharedDefinition,
} from '../src/xlsx/view/contracts'
import { prepareCfRemote, pagedStoreCfPages, type CfPagedReadHost } from '../src/xlsx/view/pagedSource'
import { RemoteCfCache, type CfReadHost } from '../src/worker/viewerRemoteCf'
import { boundedQueryStore } from '../src/worker/viewerQueryBudget'
import { interruptibleStore } from '../src/worker/viewerAbort'
import type { SourceStore } from '../src/xlsx/view/contracts'
import type { CacheIdentity, WorkerReply } from '../src/xlsx/view/contracts'
import type { CommandOf } from '../src/worker/viewerProtocol'

const SHEET = 's'
const stamp = (gen = 1, extra: Partial<CfStamp> = {}): CfStamp => ({
  owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
})
const rect = (r0: number, r1: number, c0 = 0, c1 = 0): SourceRect => ({ r0, r1, c0, c1 })
const frozenBudget = (): CfWorkBudget => createCfWorkBudget({ now: () => 0 })
function rule(order: number, props: Record<string, string | boolean | null>, writes: string[], stop = false): CfSheetRules['rules'][number] {
  return {
    id: order, priority: 1, documentOrder: order,
    ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }],
    stopIfTrue: stop, writes: new Set(writes), dxf: props,
    definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
  } as unknown as CfSheetRules['rules'][number]
}
const oneRule: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'fill.pattern': 'solid', 'fill.fg': 'FF00FF00' }, ['fill'])] }
const nc = (value: number): CfLiteral => ({ kind: 'number', value })

/** Bare async CF source for provider tests: no admission of its own. */
function fakeSource(
  cells: { row: number; col: number; value: CfLiteral }[],
  opts: { bounds?: { firstRow: number; lastRow: number; firstCol: number; lastCol: number } | null; stamp?: CfStamp; gate?: Promise<void>; visits?: { n: number }; release?: () => void } = {},
): CfValueSource & { live: { stamp: CfStamp } } {
  const live = { stamp: opts.stamp ?? stamp(1) }
  return {
    live,
    bounds: async (_id: string) => {
      if (opts.bounds === null) return undefined
      if (opts.bounds !== undefined) return opts.bounds
      const m = Math.max(0, ...cells.map(c => c.row))
      const mc = Math.max(0, ...cells.map(c => c.col))
      return { firstRow: 0, lastRow: m, firstCol: 0, lastCol: mc }
    },
    stamp: (_id: string) => live.stamp,
    pages: async function* (_id: string) {
      if (opts.gate) await opts.gate
      yield {
        value: {
          rows: Uint32Array.from(cells, c => { if (opts.visits) opts.visits.n += 1; return c.row }),
          cols: new Uint16Array(cells.map(c => c.col)),
          values: cells.map(c => c.value),
        },
        bytes: 0,
        release() { opts.release?.() },
      } as unknown as CfLease<CfValuePage>
    },
  } as unknown as CfValueSource & { live: { stamp: CfStamp } }
}
const cfRequest = (rects: readonly SourceRect[], st: CfStamp, grantBytes: number): CfReadRequest => ({
  rects,
  stamp: { owner: st.owner, sheetId: st.sheetId, workbookValueGen: st.workbookValueGen, cfGen: st.cfGen, viewGen: st.viewGen, structureGen: st.structureGen, clockGen: st.clockGen },
  grantBytes,
})

describe('CF reply estimator', () => {
  it('measures a synthetic 1024 B payload exactly', () => {
    const def = { id: 'd0', version: 1, properties: { a: 'x' }, writes: ['a'] }
    expect(measureCfDefinitionBytes(def)).toBe(154)
    const cell = { row: 0, col: 0, properties: { a: 'x', pad: 'f'.repeat(238) } }
    expect(measureCfCellBytes(cell)).toBe(614)
    expect(CF_REPLY_ENVELOPE_BYTES + measureCfDefinitionBytes(def) + measureCfCellBytes(cell)).toBe(1024)
  })

  it('does not hide shared references', () => {
    const big = 'N'.repeat(10000)
    const one = [{ row: 0, col: 0, properties: { 'font.name': big } }]
    const hundred = Array.from({ length: 100 }, (_, i) => ({ row: i, col: 0, properties: { 'font.name': big } }))
    const oneBytes = one.reduce((n, s) => n + measureCfCellBytes(s), 0)
    const hundredBytes = hundred.reduce((n, s) => n + measureCfCellBytes(s), 0)
    expect(hundredBytes - oneBytes).toBeGreaterThan(99 * 10000)
    const def = { id: 'big', version: 7, properties: { 'numFmt.code': big }, writes: ['numFmt'] }
    const defBytes = measureCfDefinitionBytes(def)
    expect(defBytes).toBeGreaterThan(20000)
    expect(defBytes).toBeLessThan(21000)
  })

  it('refuses unrepresentable estimates instead of overflowing', () => {
    expect(estimateCfPreparationBytes(Number.MAX_SAFE_INTEGER, 1 << 30, 0, 0, 0)).toBeUndefined()
    expect(estimateCfPreparationBytes(-1, 10, 10, 10, 256)).toBeUndefined()
    expect(estimateCfPreparationBytes(1, 10, 10, -5, 256)).toBeUndefined()
    // 256 + 0 + 50 + 2*100 = 506.
    expect(estimateCfPreparationBytes(2, 100, 50, 0, 256)).toBe(506)
    expect(measureCfDiagnosticBytes({ kind: 'resource-limit', feature: 'cf-test', message: 'm'.repeat(100) })).toBe(24 + (16 + 28) + (16 + 14) + (16 + 200))
  })

  it('versions rule sets deterministically by content', () => {
    const a = versionCfRules('[{"id":0,"dxf":{"a":"b"}}]')
    expect(versionCfRules('[{"id":0,"dxf":{"a":"b"}}]')).toBe(a)
    expect(versionCfRules('[{"id":0,"dxf":{"a":"c"}}]')).not.toBe(a)
    expect(versionCfRules('')).toBe(0x811c9dc5)
  })

  it('mirrors the service propertyBytes accounting for admission parity', () => {
    const rules = [
      { dxf: { 'fill.pattern': 'solid' }, writes: ['fill'] },
      { dxf: { 'font.bold': true }, writes: ['font.bold'] },
    ] as const
    // 192 + 4*len(key) + (string? len*2 : 16) per dxf entry; 128 + 4*len per write.
    // 'fill.pattern'(12):192+48+10=250; writes ['fill']:128+16=144;
    // 'font.bold'(9):192+36+16=244; writes:128+36=164. Total 802.
    expect(cfRulePropertyBytes(rules)).toBe(250 + 144 + 244 + 164)
    expect(cfRulePropertyBytes([])).toBe(0)
  })
})

describe('CF remote preparation grants and admission', () => {
  it('grant minus one refuses before payload allocation; base cells stay readable', async () => {
    // The admission bound is the provider's wire preflight estimate; grant
    // minus one must refuse before any page is ever visited, and the exact
    // grant must complete with balanced memory. The 1,024 B figure itself
    // is pinned at the measurement layer above.
    const props = { 'fill.pattern': 'solid', 'fill.fg': 'FF00FF00' }
    const cells = [{ row: 0, col: 0, value: nc(1) }]
    const rules: CfSheetRules = { diagnostics: [], rules: [rule(0, props, ['fill'])] }
    const defs = rules.rules.map((r, i) => ({ id: `rule-${i}`, version: 1, properties: { ...r.dxf }, writes: [...r.writes] }))
    const definitionBytes = defs.reduce((n, d) => n + measureCfDefinitionBytes(d), 0)
    let perCell = 40
    for (const r of rules.rules) perCell += measureCfPropertiesBytes((r.dxf ?? {}) as Record<string, string | number | boolean | null>)
    const exact = estimateCfPreparationBytes(1, perCell, definitionBytes, 0, CF_REPLY_ENVELOPE_BYTES)
    if (exact === undefined) throw new Error('expected a bounded estimate')
    const visits = { n: 0 }
    const tight = createCfMemory({})
    const refused = await prepareCfRemote(cfRequest([rect(0, 0)], stamp(1), exact - 1), SHEET, fakeSource(cells, { visits }), rules, tight, { budget: frozenBudget })
    expect(refused).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    if (refused.kind !== 'unresolved') throw new Error('expected unresolved')
    expect(refused.requiredBytes).toBe(exact)
    expect(visits.n).toBe(0)
    expect(tight.currentBytes).toBe(0)
    const roomy = createCfMemory({})
    const ready = await prepareCfRemote(cfRequest([rect(0, 0)], stamp(1), exact), SHEET, fakeSource(cells, { visits }), rules, roomy, { budget: frozenBudget })
    expect(ready.kind).toBe('ready')
    expect(visits.n).toBeGreaterThan(0)
    expect(roomy.currentBytes).toBe(0)
  })

  it('an exactly-full base reply omits inline diagnostics instead of refusing', async () => {
    // A rule set carrying a parser diagnostic produces a notice on the roomy
    // reply; a grant that fits only the base still completes ready with the
    // diagnostic omitted (never silently partial data).
    const st = stamp(1)
    const rules: CfSheetRules = {
      diagnostics: [{ kind: 'unsupported', feature: 'cf-test-parse', message: 'synthetic parser notice' }],
      rules: [rule(0, { 'fill.pattern': 'solid' }, ['fill'])],
    }
    const src = () => fakeSource([{ row: 0, col: 0, value: nc(5) }])
    const wide = await prepareCfRemote(cfRequest([rect(0, 0)], st, 1 << 24), SHEET, src(), rules, createCfMemory({}), { budget: frozenBudget })
    expect(wide.kind).toBe('ready')
    if (wide.kind !== 'ready') throw new Error('expected ready')
    expect(wide.diagnostics.length).toBeGreaterThan(0)
    expect(wide.states).toHaveLength(1)
    expect(wide.states[0]!.unresolvedReason).toBeUndefined()
    const base = CF_REPLY_ENVELOPE_BYTES
      + wide.definitions.reduce((n: number, d) => n + measureCfDefinitionBytes(d), 0)
      + wide.states.reduce((n: number, s) => n + measureCfCellBytes(s), 0)
    const tight = await prepareCfRemote(cfRequest([rect(0, 0)], st, base), SHEET, src(), rules, createCfMemory({}), { budget: frozenBudget })
    expect(tight.kind).toBe('ready')
    if (tight.kind !== 'ready') throw new Error('expected ready with omitted diagnostic')
    expect(tight.diagnostics).toEqual([])
    expect(tight.states).toHaveLength(1)
  })

  it('rejects invalid grants and ranges before touching the source', async () => {
    const memory = createCfMemory({})
    let touched = 0
    const src = fakeSource([{ row: 0, col: 0, value: nc(1) }], { visits: { n: 0 } })
    const wrapped = {
      ...src,
      bounds: async (_id: string) => { touched++; return rect(0, 0) },
    } as never
    const badGrant = await prepareCfRemote(cfRequest([rect(0, 0)], stamp(1), 0), SHEET, wrapped, oneRule, memory, { budget: frozenBudget })
    expect(badGrant).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    const badRect = await prepareCfRemote(
      cfRequest([{ r0: 5, r1: 3, c0: 0, c1: 0 }], stamp(1), 100),
      SHEET, wrapped, oneRule, memory, { budget: frozenBudget })
    expect(badRect.kind).toBe('unresolved')
    expect(touched).toBe(0)
    expect(memory.currentBytes).toBe(0)
  })

  it('stale sheet identity is dropped and refunded', async () => {
    const memory = createCfMemory({})
    const src = fakeSource([{ row: 0, col: 0, value: nc(1) }])
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], stamp(1), 1_000_000), 'other-sheet', src, oneRule, memory, { budget: frozenBudget })
    expect(out).toMatchObject({ kind: 'unresolved', reason: 'cf-stale-generation' })
    expect(memory.currentBytes).toBe(0)
  })

  it('stale workbook generation is dropped and refunded with the live stamp', async () => {
    const memory = createCfMemory({})
    const live = { stamp: stamp(2) }
    const src = fakeSource([{ row: 0, col: 0, value: nc(1) }], { stamp: live.stamp })
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], stamp(1), 1_000_000), SHEET, src, oneRule, memory, { budget: frozenBudget })
    expect(out).toMatchObject({ kind: 'unresolved', reason: 'cf-stale-generation' })
    if (out.kind !== 'unresolved') throw new Error('expected unresolved')
    expect(out.stamp.workbookValueGen).toBe(2)
    expect(memory.currentBytes).toBe(0)
  })

  it('unimported ranges stay pending without running any scan', async () => {
    const memory = createCfMemory({})
    const counter = { n: 0 }
    const src = fakeSource(
      Array.from({ length: 50 }, (_, i) => ({ row: i, col: 0, value: nc(i) })),
      { bounds: null, visits: counter },
    )
    const gated = {
      ...src,
      pages: async function* (...args: [string, readonly never[], CfWorkBudget, AbortSignal]) {
        counter.n += 1000
        yield* src.pages(...args)
      },
    }
    const out = await prepareCfRemote(cfRequest([rect(0, 49)], stamp(1), 8 * 1024 * 1024), SHEET, gated as never, oneRule, memory, { budget: frozenBudget })
    expect(out).toMatchObject({ kind: 'pending', reason: 'cf-pending-data' })
    if (out.kind !== 'pending') throw new Error('expected pending')
    expect(out.pendingSheetIds).toEqual([SHEET])
    expect(counter.n).toBe(0)
    expect(memory.currentBytes).toBe(0)
  })

  it('paged store bridge chunks, filters and admits bounded pages', async () => {
    // 130 cells minus unavailable rows split into 64-cell admitted pages;
    // out-of-scope and unavailable cells never enter a page; every page is
    // admitted before construction and released by the consumer.
    const memory = createCfMemory({})
    const released: number[] = []
    const store: CfPagedReadHost = {
      stamp: { docId: 'doc', sheetId: SHEET, valueGen: 1 },
      readRows: async (range, cursor) => {
        if (cursor !== undefined) throw new Error('unexpected cursor in single-page fixture')
        const items: { row: number; cells: { id: number; row: number; col: number; value: number; styleIndex: number; unavailable?: true }[] }[] = []
        for (let row = 0; row < 130; row++) {
          if (row < range.r0 || row > range.r1) continue
          if (row % 10 === 9) items.push({ row, cells: [{ id: row * 16384, row, col: 0, value: row, styleIndex: 0, unavailable: true }] })
          else items.push({ row, cells: [{ id: row * 16384, row, col: 0, value: row, styleIndex: 0 }] })
        }
        return { items, completeness: 'complete' as const, stamp: store.stamp }
      },
    }
    const pages = pagedStoreCfPages(store, memory)
    const seen: Array<{ row: number; col: number; value: unknown }> = []
    for await (const lease of pages('s', [{ firstRow: 0, lastRow: 129, firstCol: 0, lastCol: 0 }], frozenBudget(), new AbortController().signal)) {
      expect(lease.value.rows.length).toBeLessThanOrEqual(64)
      expect(lease.value.rows.length).toBe(lease.value.cols.length)
      expect(lease.value.rows.length).toBe(lease.value.values.length)
      for (let i = 0; i < lease.value.rows.length; i++) {
        const v = lease.value.values[i] as { kind: string; value?: unknown }
        seen.push({ row: lease.value.rows[i]!, col: lease.value.cols[i]!, value: v.kind === 'number' ? v.value : v })
      }
      lease.release()
      released.push(lease.value.rows.length)
    }
    // 130 cells minus 13 unavailable = 117 visited, split 64 + 53.
    expect(seen).toHaveLength(117)
    expect(seen.every(c => (c.row as number) % 10 !== 9)).toBe(true)
    expect(released).toEqual([64, 53])
    expect(memory.currentBytes).toBe(0)
  })

  it('adapter withholds while importing and serves empty sheets without scanning', async () => {
    const { PagedXlsxSheetAdapter } = await import('../src/xlsx/view/pagedSource')
    const sourceStamp = { docId: 'doc', sheetId: SHEET, valueGen: 1 }
    const fakeSheet = (over: Record<string, unknown>) => ({
      sheetId: SHEET, name: SHEET, state: 'loading' as const, importedThroughRow: -1,
      metadata: {}, tables: undefined, arrayCells: [], drawingAnchors: { records: [], unresolved: 0 },
      mergeRects: [], invalidMerges: 0, used: undefined, conditionalFormatting: undefined,
      ...over,
    })
    const loading = new PagedXlsxSheetAdapter(sourceStamp, {} as never, fakeSheet({}) as never, undefined, {}, false)
    const pending = await loading.readConditionalFormatting(
      cfRequest([rect(0, 9)], stamp(1), 1 << 20), createCfMemory({}), { budget: frozenBudget },
    )
    expect(pending).toMatchObject({ kind: 'pending', reason: 'cf-pending-data', pendingSheetIds: [SHEET] })
    const complete = new PagedXlsxSheetAdapter(sourceStamp, {} as never, fakeSheet({ state: 'complete', importedThroughRow: 1048575 }) as never, undefined, {}, false)
    const empty = await complete.readConditionalFormatting(
      cfRequest([rect(0, 9)], stamp(1), 1 << 20), createCfMemory({}), { budget: frozenBudget },
    )
    expect(empty).toMatchObject({ kind: 'ready', states: [] })
    await loading.release().catch(() => {})
    await complete.release().catch(() => {})
  })

  it('abort, timeout and close each release exactly once with a gated provider', async () => {
    function deferred<T = void>() { let resolve!: (v: T | Promise<T>) => void; const promise = new Promise<T>(r => { resolve = r }); return { promise, resolve } }
    const ways = ['abort', 'timeout', 'close'] as const
    for (const way of ways) {
      const memory = createCfMemory({})
      const gate = deferred<void>()
      const releases = { n: 0 }
      const src = fakeSource(
        Array.from({ length: 200 }, (_, i) => ({ row: i, col: 0, value: nc(i) })),
        { gate: gate.promise, release: () => { releases.n++ } },
      )
      const controller = new AbortController()
      const pending = prepareCfRemote(cfRequest([rect(0, 199)], stamp(1), 8 * 1024 * 1024), SHEET, src, oneRule, memory, { signal: controller.signal, budget: frozenBudget })
      // Let the scan start and block on the gate (admitted memory held).
      await new Promise(r => setTimeout(r, 0))
      expect(memory.currentBytes).toBeGreaterThan(0)
      const held = memory.currentBytes
      if (way === 'timeout') await new Promise(r => setTimeout(r, 5))
      // The gated provider ignores the abort; the charge persists until it
      // settles, then unwinds exactly once without publishing.
      controller.abort()
      gate.resolve()
      await expect(pending).rejects.toMatchObject({ code: 'cancelled' })
      await new Promise(r => setTimeout(r, 10))
      expect(releases.n).toBe(1)
      expect(memory.currentBytes).toBe(0)
      expect(memory.peakBytes).toBeGreaterThanOrEqual(held)
    }
  })
})

describe('CF remote cache (viewerRemoteCf)', () => {
  const identity = (sheetId = SHEET, valueGen = 1): CacheIdentity => ({ docId: 'doc', sheetId, valueGen, fontGen: 0, viewGen: 0 })
  const cfStamp = (gen = 1, extra: Partial<CfReadStamp> = {}): CfReadStamp => ({
    owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
  })
  const wireCell = (row: number, col: number, props: Record<string, string | number | boolean | null> = { 'fill.pattern': 'solid' }): CfCellStateWire => ({ row, col, properties: props })
  const wireDef = (id = 'rule-0'): CfSharedDefinition => ({ id, version: 7, properties: { 'fill.pattern': 'solid' }, writes: ['fill'] })
  const wireResult = (over: Partial<CfReadResult> = {}): CfReadResult => ({
    sheetId: SHEET, stamp: cfStamp(1), kind: 'ready', states: [wireCell(0, 0)], definitions: [wireDef()],
    definitionsVersion: 7, diagnostics: [], ...over,
  })
  function cfReply(result: CfReadResult, bytes: number, requestId = 'r1', owner: CacheIdentity = identity()): WorkerReply<'read'> {
    return {
      commandKind: 'read', requestId, owner, diagnostics: [], replyBytes: bytes, kind: 'complete',
      result: { kind: 'read', result: { op: 'conditionalFormatting', result } },
    } as WorkerReply<'read'>
  }

  interface FakeHost extends CfReadHost {
    requests: { command: CommandOf<'read'>; signal?: AbortSignal }[]
    granted: Map<string, number>
    releasedReplies: WorkerReply<'read'>[]
    changedCalls: { sheetId: string; rects: readonly SourceRect[]; reason?: string }[]
    respond: (command: CommandOf<'read'>, signal?: AbortSignal) => Promise<WorkerReply<'read'> | symbol>
    hasCf: boolean
    load: 'complete' | 'loading' | 'incomplete'
    ledger: Map<string | object, number>
  }
  function testHost(overrides: Partial<FakeHost> & { respond?: FakeHost['respond']; grantBytes?: number | null } = {}): FakeHost {
    const ledger = new Map<string | object, number>()
    const host = {
      requests: [],
      granted: new Map<string, number>(),
      releasedReplies: [],
      changedCalls: [],
      hasCf: true,
      load: 'complete' as const,
      ledger,
      identity: (sheetId: string) => identity(sheetId),
      loadState: () => ({ completeness: (host.load ?? 'complete') as 'complete' | 'loading' | 'incomplete' }),
      hasConditionalFormatting: () => host.hasCf,
      grant: (requestId: string) => {
        if (overrides.grantBytes === null) return null
        const bytes = overrides.grantBytes ?? 1 << 20
        host.granted.set(requestId, bytes)
        ledger.set(`grant:${requestId}`, bytes)
        return bytes
      },
      releaseGrant: (requestId: string) => {
        host.granted.delete(requestId)
        ledger.delete(`grant:${requestId}`)
      },
      request: async (command: CommandOf<'read'>, signal?: AbortSignal) => {
        host.requests.push({ command, signal })
        return host.respond(command, signal)
      },
      nextRequestId: (() => { let n = 0; return () => `r${++n}` })(),
      release: (reply: WorkerReply<'read'>) => {
        host.releasedReplies.push(reply)
        ledger.delete(reply)
      },
      changed: (sheetId: string, rects: readonly SourceRect[], reason?: 'updated' | 'evicted' | 'status') => {
        void reason
        host.changedCalls.push({ sheetId, rects, reason })
      },
      respond: overrides.respond ?? (async (command: CommandOf<'read'>) => {
        const q = command.query
        const grant = host.granted.get(command.requestId) ?? 1 << 20
        const result = wireResult()
        const reply = cfReply(result, 1024, command.requestId, command.owner as CacheIdentity)
        host.granted.delete(command.requestId)
        ledger.delete(`grant:${command.requestId}`)
        const moved = Math.min(1024, grant)
        ledger.set(reply, moved)
        void q
        return reply
      }),
    } as unknown as FakeHost
    Object.assign(host, { grantBytes: undefined })
    if (overrides.grantBytes !== undefined) (host as unknown as Record<string, unknown>).grantBytesOverride = overrides.grantBytes
    return host
  }
  const flush = () => new Promise(r => setTimeout(r, 0))

  it('loads once for two viewport consumers and releases once', async () => {
    const host = testHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const a = cache.acquire(SHEET, [rect(0, 4)], cfStamp(1))
    const b = cache.acquire(SHEET, [rect(0, 4)], cfStamp(1))
    expect(host.requests).toHaveLength(1)
    await flush()
    expect(a.state.state).toBe('ready')
    expect(b.state.state).toBe('ready')
    expect(a.state.result).toBe(b.state.result)
    expect(host.releasedReplies).toHaveLength(0)
    a.release()
    expect(host.releasedReplies).toHaveLength(0)
    b.release()
    cache.invalidate(SHEET)
    await flush()
    expect(host.releasedReplies).toHaveLength(1)
    expect(cache.retainedBytes).toBe(0)
    cache.close()
  })

  it('drops and refunds each stale stamp dimension independently', async () => {
    const dims = [
      { owner: 'other' },
      { workbookValueGen: 2 },
      { cfGen: 2 },
      { structureGen: 2 },
    ] as const
    for (const dim of dims) {
      const host = testHost()
      const cache = new RemoteCfCache(host, 1 << 20)
      const current = { ...cfStamp(1), ...dim, sheetId: SHEET }
      void current
      const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
      await flush()
      expect(held.state.state).toBe('ready')
      expect(cache.retainedBytes).toBeGreaterThan(0)
      const next = cache.acquire(SHEET, [rect(0, 0)], { ...cfStamp(1), ...dim })
      await flush()
      expect(next.state.state).toBe('ready')
      expect(next.state.result).not.toBe(held.state.result)
      // Old version stays charged while its consumer holds it...
      expect(host.releasedReplies).toHaveLength(0)
      held.release()
      // ...then releases exactly once after the last release + invalidation.
      cache.invalidate(SHEET)
      await flush()
      expect(host.releasedReplies).toHaveLength(1)
      next.release()
      cache.invalidate(SHEET)
      await flush()
      expect(host.releasedReplies).toHaveLength(2)
      expect(cache.retainedBytes).toBe(0)
      cache.close()
    }
  })

  it('definition replacement keeps the old pinned version charged until last release', async () => {
    const host = testHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const old = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(old.state.state).toBe('ready')
    const fresh = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1, { cfGen: 2 }))
    await flush()
    expect(fresh.state.state).toBe('ready')
    expect(host.requests).toHaveLength(2)
    // Both versions charged while both are held.
    expect(cache.retainedBytes).toBeGreaterThan(0)
    expect(host.releasedReplies).toHaveLength(0)
    old.release()
    cache.invalidate(SHEET)
    await flush()
    // Old entry detached on replacement; released exactly once now.
    expect(host.releasedReplies).toHaveLength(1)
    fresh.release()
    cache.invalidate(SHEET)
    await flush()
    expect(host.releasedReplies).toHaveLength(2)
    expect(cache.retainedBytes).toBe(0)
    cache.close()
  })

  it('evicts the eldest entry exactly once at the budget edge', async () => {
    const host = testHost()
    const cache = new RemoteCfCache(host, 1500)
    const a = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(a.state.state).toBe('ready')
    expect(cache.retainedBytes).toBe(1024)
    a.release()
    // Released but still cached: a second consumer reuses it without a request.
    const a2 = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(host.requests).toHaveLength(1)
    expect(a2.state.state).toBe('ready')
    a2.release()
    // A new scope evicts the eldest unreferenced entry, releasing once.
    const b = cache.acquire(SHEET, [rect(10, 10)], cfStamp(1))
    await flush()
    expect(b.state.state).toBe('ready')
    expect(host.releasedReplies).toHaveLength(1)
    expect(cache.retainedBytes).toBe(1024)
    // While everything is pinned, a new scope fails closed instead of
    // evicting live data or exceeding the budget (its reply still released).
    const c = cache.acquire(SHEET, [rect(20, 20)], cfStamp(1))
    await flush()
    expect(c.state.state).toBe('failed')
    expect(c.state.diagnostics[0]!.feature).toBe('cf-cache-budget')
    expect(cache.retainedBytes).toBe(1024)
    expect(host.releasedReplies).toHaveLength(2)
    b.release()
    c.release()
    cache.close()
    expect(host.releasedReplies).toHaveLength(3)
    expect(cache.retainedBytes).toBe(0)
  })

  it('abort, timeout and close each release exactly once with live-ledger assertions', async () => {
    for (const way of ['abort', 'timeout', 'close'] as const) {
      const deferreds: { resolve: (v: WorkerReply<'read'> | symbol) => void }[] = []
      const host = testHost({
        respond: (_command, signal) => new Promise<WorkerReply<'read'> | symbol>(resolve => {
          deferreds.push({ resolve })
          signal?.addEventListener('abort', () => resolve(Symbol('cancelled')), { once: true })
        }),
      })
      const cache = new RemoteCfCache(host, 1 << 20)
      const controller = new AbortController()
      const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1), controller.signal)
      await flush()
      expect(held.state.state).toBe('loading')
      expect(host.granted.size).toBe(1)
      expect(host.ledger.size).toBeGreaterThan(0)
      if (way === 'timeout') await new Promise(r => setTimeout(r, 5))
      // The waiter detaches; the shared build is cancelled only when its
      // entry drops (evict/invalidate/close). Resolving the late reply must
      // release, never publish nor double-count.
      if (way === 'close') cache.close()
      else controller.abort()
      for (const d of deferreds) d.resolve(Symbol('cancelled'))
      await flush()
      expect(host.granted.size).toBe(0)
      held.release()
      if (way !== 'close') cache.close()
      expect(host.ledger.size).toBe(0)
    }
  })

  it('unimported ranges stay pending and invalidate only affected tiles', async () => {
    const host = testHost()
    host.load = 'loading'
    const cache = new RemoteCfCache(host, 1 << 20)
    const waiting = cache.get(SHEET, [rect(0, 49)], cfStamp(1))
    expect(waiting.state).toBe('waiting')
    expect(host.requests).toHaveLength(0)
    host.load = 'complete'
    const loading = cache.get(SHEET, [rect(0, 49)], cfStamp(1))
    expect(loading.state).toBe('loading')
    expect(host.requests).toHaveLength(1)
    await flush()
    expect(cache.get(SHEET, [rect(0, 49)], cfStamp(1)).state).toBe('ready')
    await flush()
    // One notification, bounded to the affected tiles (not the whole sheet).
    const updates = host.changedCalls.filter(c => c.reason === 'updated')
    expect(updates).toHaveLength(1)
    expect(updates[0]!.rects).toEqual([rect(0, 49)])
    cache.close()
  })

  it('notifications are deferred outside render and union bounded rects', async () => {
    const host = testHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const a = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    const b = cache.acquire(SHEET, [rect(1, 1)], cfStamp(1))
    // Synchronous get/acquire never notifies.
    expect(host.changedCalls).toHaveLength(0)
    await flush()
    expect(a.state.state).toBe('ready')
    expect(b.state.state).toBe('ready')
    // Both completions flushed in one microtask with both rects.
    expect(host.changedCalls.filter(c => c.reason === 'updated')).toHaveLength(1)
    expect(host.changedCalls[0]!.rects).toEqual(expect.arrayContaining([rect(0, 0), rect(1, 1)]))
    a.release()
    b.release()
    cache.close()
  })

  it('absent CF capability never touches the worker', () => {
    const host = testHost()
    host.hasCf = false
    const cache = new RemoteCfCache(host, 1 << 20)
    const state = cache.get(SHEET, [rect(0, 0)], cfStamp(1))
    expect(state.state).toBe('waiting')
    expect(host.requests).toHaveLength(0)
    expect(host.granted.size).toBe(0)
    cache.close()
  })
})

describe('CF store wrappers', () => {
  const baseStamp = { docId: 'doc', sheetId: SHEET, valueGen: 1 }
  const cfReq = (grantBytes = 1024) => ({
    rects: [rect(0, 0)],
    stamp: { owner: 'o', sheetId: SHEET, workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1 },
    grantBytes,
  })

  it('bounded wrapper preserves an absent CF capability without touching the stamp', () => {
    // A strict mock whose stamp getter throws: wrapping must only probe the
    // optional method, never the stamp, and the facade must stay absent.
    let stampTouched = false
    const mock = {
      get stamp() { stampTouched = true; throw new Error('stamp touched') },
      readConditionalFormatting: undefined as never,
    } as unknown as SourceStore
    const wrapped = boundedQueryStore(mock, 1024)
    expect('readConditionalFormatting' in wrapped).toBe(false)
    expect(stampTouched).toBe(false)
  })

  it('bounded wrapper forwards a present CF capability', async () => {
    const inner = {
      stamp: baseStamp,
      readConditionalFormatting: async () => ({ sheetId: SHEET, stamp: cfStampShim(), kind: 'ready', states: [], definitions: [], definitionsVersion: 1, diagnostics: [] }),
    } as unknown as SourceStore
    const wrapped = boundedQueryStore(inner, 1024)
    expect(typeof wrapped.readConditionalFormatting).toBe('function')
    const out = await wrapped.readConditionalFormatting!(cfReq(), undefined, {})
    expect(out).toMatchObject({ kind: 'ready' })
  })

  it('interruptible wrapper preserves an absent CF capability without touching the stamp', () => {
    let stampTouched = false
    const mock = {
      get stamp() { stampTouched = true; throw new Error('stamp touched') },
      readConditionalFormatting: undefined as never,
    } as unknown as SourceStore
    const wrapped = interruptibleStore(mock)
    expect('readConditionalFormatting' in wrapped).toBe(false)
    expect(stampTouched).toBe(false)
  })

  it('interruptible wrapper forwards a present CF capability and keeps the abort contract', async () => {
    const seen: AbortSignal[] = []
    const inner = {
      stamp: baseStamp,
      readConditionalFormatting: async (_req: unknown, _mem: unknown, opts?: { signal?: AbortSignal }) => {
        if (opts?.signal) seen.push(opts.signal)
        return { sheetId: SHEET, stamp: cfStampShim(), kind: 'ready', states: [], definitions: [], definitionsVersion: 1, diagnostics: [] }
      },
    } as unknown as SourceStore
    const wrapped = interruptibleStore(inner)
    const controller = new AbortController()
    const out = await wrapped.readConditionalFormatting!(cfReq(), undefined, { signal: controller.signal })
    expect(out).toMatchObject({ kind: 'ready' })
    expect(seen).toHaveLength(1)
    controller.abort()
    const cancelled = await wrapped.readConditionalFormatting!(cfReq(), undefined, { signal: controller.signal })
    expect(cancelled).toMatchObject({ kind: 'unresolved' })
  })
  function cfStampShim() {
    return { owner: 'o', sheetId: SHEET, workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1 }
  }
})

describe('CF worker round trip (in-process kernel)', () => {
  // Exercises session dispatch, the kernel grant envelope and the generic
  // publish path with CF payloads — no real worker thread required.
  async function kernelHarness(cellCount = 1) {
    const { ViewerClient } = await import('../src/worker/viewerClient')
    const { connectInProcess, ViewerKernel } = await import('../src/worker/viewerKernel')
    const { createObjectWorkbookSource } = await import('../src/worker/viewerObjectSource')
    const { initialIdentity } = await import('../src/worker/viewerProtocol')
    const { workbookOf } = await import('./helpers/xlsx-find')
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 5 }] }])
    const base = createObjectWorkbookSource(doc, 'doc')
    const calls: { grant: number }[] = []
    const wrapped = {
      ...base,
      store: (id: string) => ({
        ...base.store(id),
        readConditionalFormatting: async (request: CfReadRequest) => {
          calls.push({ grant: request.grantBytes })
          const states = Array.from({ length: cellCount }, (_, i) => ({ row: i, col: 0, properties: { 'fill.pattern': 'solid', 'note': 'n'.repeat(50) } }))
          const definitions = [{ id: 'rule-0', version: 9, properties: { 'fill.pattern': 'solid' }, writes: ['fill'] }]
          return {
            sheetId: id,
            stamp: { ...request.stamp },
            kind: 'ready' as const,
            states,
            definitions,
            definitionsVersion: 9,
            diagnostics: [],
          }
        },
      }),
    }
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => wrapped }, { yield: async () => {} })
    const port = connectInProcess(kernel)
    const client = new ViewerClient(port)
    const open = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
    if (open.kind !== 'complete') throw new Error('open failed')
    return { client, kernel, calls, initialIdentity }
  }

  it('serves a CF read end to end and stamps its exact size', async () => {
    const h = await kernelHarness()
    const owner = h.initialIdentity('doc', 's')
    const out = await h.client.send({
      kind: 'read', requestId: h.client.nextRequestId(), owner,
      query: { op: 'conditionalFormatting', rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], stamp: { owner: 'owner', sheetId: 's', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 1, clockGen: 0 }, grantBytes: 1 << 20 },
    })
    expect(out.kind).toBe('complete')
    if (out.kind !== 'complete' || out.result.kind !== 'read' || out.result.result.op !== 'conditionalFormatting') throw new Error('unexpected reply shape')
    expect(out.result.result.result.kind).toBe('ready')
    expect(out.result.result.result.states).toHaveLength(1)
    expect(Number.isSafeInteger(out.replyBytes) && (out.replyBytes ?? 0) > 0).toBe(true)
  })

  it('replaces an over-grant reply with a typed whole-reply quota', async () => {
    const h = await kernelHarness(500)
    const owner = h.initialIdentity('doc', 's')
    const stamp = { owner: 'owner', sheetId: 's', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 1, clockGen: 0 }
    const out = await h.client.send({
      kind: 'read', requestId: h.client.nextRequestId(), owner,
      query: { op: 'conditionalFormatting', rects: [{ r0: 0, r1: 499, c0: 0, c1: 0 }], stamp, grantBytes: 4096 },
    })
    // Above the envelope (so the kernel dispatches) but below the measured
    // reply: the kernel replaces the over-grant reply with a payload-free
    // quota naming its measured size.
    expect(out.kind).toBe('quota')
    if (out.kind !== 'quota') throw new Error('expected quota')
    expect(typeof out.requiredBytes).toBe('number')
    expect(out.limitBytes).toBe(4096)
  })

  it('refuses a malformed grant without running the provider', async () => {
    const h = await kernelHarness()
    const owner = h.initialIdentity('doc', 's')
    const out = await h.client.send({
      kind: 'read', requestId: h.client.nextRequestId(), owner,
      query: { op: 'conditionalFormatting', rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], stamp: { owner: 'owner', sheetId: 's', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 1, clockGen: 0 }, grantBytes: 0 },
    })
    expect(h.calls).toHaveLength(0)
    // Malformed grants surface as a typed error or an unresolved refusal;
    // either way the provider never ran.
    expect(out.kind === 'error' || (out.kind === 'complete' && out.result.kind === 'read')).toBe(true)
  })
})

