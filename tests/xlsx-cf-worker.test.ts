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
  type CfReadOptions,
  type CfCellStateWire,
  type CfSharedDefinition,
} from '../src/xlsx/view/contracts'
import { prepareCfRemote, pagedStoreCfPages, cfRulesVersion, releaseSharedCfDefinitions, type CfPagedReadHost } from '../src/xlsx/view/pagedSource'
import { RemoteCfCache, entryMetaBytes, type CfReadHost } from '../src/worker/viewerRemoteCf'
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
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
  opts: { bounds?: { firstRow: number; lastRow: number; firstCol: number; lastCol: number } | null; stamp?: CfStamp; gate?: Promise<void>; gateAfterChunks?: number; visits?: { n: number }; release?: () => void } = {},
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
      if (opts.gate && (opts.gateAfterChunks ?? 0) <= 0) await opts.gate
      let yielded = 0
      for (let i = 0; i < cells.length; i += 64) {
        const chunk = cells.slice(i, i + 64)
        yield {
          value: {
            rows: Uint32Array.from(chunk, c => { if (opts.visits) opts.visits.n += 1; return c.row }),
            cols: new Uint16Array(chunk.map(c => c.col)),
            values: chunk.map(c => c.value),
          },
          bytes: 0,
          release() { opts.release?.() },
        } as unknown as CfLease<CfValuePage>
        yielded++
        if (opts.gate && yielded >= (opts.gateAfterChunks ?? Number.MAX_SAFE_INTEGER)) await opts.gate
      }
    },
  } as unknown as CfValueSource & { live: { stamp: CfStamp } }
}
const cfRequest = (rects: readonly SourceRect[], st: CfStamp, grantBytes: number): CfReadRequest => ({
  rects,
  stamp: { owner: st.owner, sheetId: st.sheetId, workbookValueGen: st.workbookValueGen, cfGen: st.cfGen, viewGen: st.viewGen, structureGen: st.structureGen, clockGen: st.clockGen },
  grantBytes,
})
/** Fake source whose stamp carries the live version for the given rules. */
function versionedSrc(cells: { row: number; col: number; value: CfLiteral }[], rules: CfSheetRules, opts: { bounds?: { firstRow: number; lastRow: number; firstCol: number; lastCol: number } | null; visits?: { n: number }; gate?: Promise<void>; release?: () => void } = {}) {
  return fakeSource(cells, { ...opts, stamp: stamped(rules) })
}
/** Stamp carrying the live content hash for the given rules (what a synced client sends). */
const stamped = (rules: CfSheetRules, gen = 1, extra: Partial<CfStamp> = {}): CfStamp =>
  stamp(gen, { ...extra, cfGen: cfRulesVersion(rules) })

/** Expected retained bytes of one ready cache entry under single-charge
 * accounting: admitted scope/stamp metadata plus the walked states (with
 * framing and diagnostics, without shared definitions). Mirrors the
 * production split instead of hardcoding sizes. */
const entryBytesFor = (sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp, result: CfReadResult): number =>
  entryMetaBytes(sheetId, rects, stamp) + measurePayloadBytes({ ...result, definitions: [] }, () => true)!

describe('CF reply estimator', () => {
  it('measures a synthetic 2048 B payload exactly', () => {
    const def = { id: 'd0', version: 1, properties: { a: 'x' }, writes: ['a'] }
    expect(measureCfDefinitionBytes(def)).toBe(216)
    const cell = { row: 0, col: 0, properties: { a: 'x', pad: 'f'.repeat(251) } }
    expect(measureCfCellBytes(cell)).toBe(680)
    expect(CF_REPLY_ENVELOPE_BYTES + measureCfDefinitionBytes(def) + measureCfCellBytes(cell)).toBe(2048)
  })

  it('the envelope bound covers realistic reply wrappers', async () => {
    const { measurePayloadBytes } = await import('../src/xlsx/view/payloadBytes')
    const wrap = (requestId: string) => ({
      commandKind: 'read', requestId,
      owner: { docId: 'doc', sheetId: 's', valueGen: 1, fontGen: 0, viewGen: 0 },
      diagnostics: [], replyBytes: 0, kind: 'complete',
      result: { kind: 'read', result: { op: 'conditionalFormatting', result: {
        sheetId: 's', stamp: { owner: 'o', sheetId: 's', workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1 },
        kind: 'ready', states: [], definitions: [], definitionsVersion: 1, diagnostics: [],
      } } },
    })
    // Typical short ids fit comfortably; the kernel walk backs up outliers.
    expect(measurePayloadBytes(wrap('r1'), () => true)!).toBeLessThanOrEqual(CF_REPLY_ENVELOPE_BYTES)
    expect(measurePayloadBytes(wrap('request-123'), () => true)!).toBeLessThanOrEqual(CF_REPLY_ENVELOPE_BYTES)
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
    expect(measureCfDiagnosticBytes({ kind: 'resource-limit', feature: 'cf-test', message: 'm'.repeat(100) })).toBe(24 + (16 + (16 + 28)) + (16 + (16 + 14)) + (16 + (16 + 200)))
  })

  it('versions rule sets deterministically by content', () => {
    const a = versionCfRules('[{"id":0,"dxf":{"a":"b"}}]')
    expect(versionCfRules('[{"id":0,"dxf":{"a":"b"}}]')).toBe(a)
    expect(versionCfRules('[{"id":0,"dxf":{"a":"c"}}]')).not.toBe(a)
    expect(versionCfRules('')).toBe(0x811c9dc5)
  })

  it('wire measures match the kernel payload walker exactly', async () => {
    // The provider preflight and the kernel publish gate must agree byte for
    // byte; this pins the shared model against the repository walker.
    const { measurePayloadBytes } = await import('../src/xlsx/view/payloadBytes')
    const cases: Array<[string, Record<string, unknown> | unknown[], (v: never) => number]> = [
      ['cell', { row: 0, col: 1, properties: { a: 'x', n: 5, b: true, z: null } }, measureCfCellBytes as (v: never) => number],
      ['cell-tainted', { row: 2, col: 3, properties: { 'fill.pattern': 'solid' }, unresolvedWrites: ['fill', 'font.bold'], unresolvedReason: 'r' }, measureCfCellBytes as (v: never) => number],
      ['def', { id: 'rule-9', version: 3, properties: { a: 'b' }, writes: ['a', 'b'] }, measureCfDefinitionBytes as (v: never) => number],
      ['diag', { kind: 'unsupported', feature: 'cf-x', message: 'hello' }, (v: never) => measureCfDiagnosticBytes(v as never)],
    ]
    for (const [name, value, fn] of cases) {
      const mine = fn(value as never)
      const walked = measurePayloadBytes(value, () => true)
      expect(mine, name).toBe(walked)
    }
    expect(measureCfPropertiesBytes({ a: 'x', bb: 1 })).toBe(measurePayloadBytes({ a: 'x', bb: 1 }, () => true))
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
    // The admission bound is the provider's preflight estimate: a walked
    // template (exact framing + definitions) plus a bounded per-cell term
    // that matches the walker exactly for full wins. Grant minus one must
    // refuse before any page is ever visited, and the exact grant must
    // complete with balanced memory.
    const props = { 'fill.pattern': 'solid', 'fill.fg': 'FF00FF00' }
    const cells = [{ row: 0, col: 0, value: nc(1) }]
    const rules: CfSheetRules = { diagnostics: [], rules: [rule(0, props, ['fill'])] }
    const defs = rules.rules.map((r, i) => ({ id: `rule-${i}`, version: 1, properties: { ...r.dxf }, writes: [...r.writes] }))
    const templateBytes = measurePayloadBytes({ sheetId: SHEET, stamp: stamp(1), kind: 'ready', states: [], definitions: defs, definitionsVersion: 1, diagnostics: [] }, () => true)
    if (templateBytes === undefined) throw new Error('expected measurable template')
    let perCell = 96
    for (const r of rules.rules) perCell += measureCfPropertiesBytes((r.dxf ?? {}) as Record<string, string | number | boolean | null>)
    const exact = templateBytes + estimateCfPreparationBytes(1, perCell, 0, 0, CF_REPLY_ENVELOPE_BYTES)!;
    if (exact === undefined) throw new Error('expected a bounded estimate')
    const st = stamped(rules)
    const visits = { n: 0 }
    const tight = createCfMemory({})
    const refused = await prepareCfRemote(cfRequest([rect(0, 0)], st, exact - 1), SHEET, versionedSrc(cells, rules, { visits }), rules, tight, { budget: frozenBudget })
    expect(refused).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    if (refused.kind !== 'unresolved') throw new Error('expected unresolved')
    expect(refused.requiredBytes).toBe(exact)
    expect(visits.n).toBe(0)
    expect(tight.currentBytes).toBe(0)
    const roomy = createCfMemory({})
    const ready = await prepareCfRemote(cfRequest([rect(0, 0)], st, exact), SHEET, versionedSrc(cells, rules, { visits }), rules, roomy, { budget: frozenBudget })
    expect(ready.kind).toBe('ready')
    expect(visits.n).toBeGreaterThan(0)
    // The committed share owns the definitions until explicitly released.
    releaseSharedCfDefinitions(rules)
    expect(roomy.currentBytes).toBe(0)
  })

  it('an exactly-full base reply omits inline diagnostics instead of refusing', async () => {
    // A rule set carrying a parser diagnostic produces a notice on the roomy
    // reply; a grant that fits only the base still completes ready with the
    // diagnostic omitted (never silently partial data).
    const rules: CfSheetRules = {
      diagnostics: [{ kind: 'unsupported', feature: 'cf-test-parse', message: 'synthetic parser notice' }],
      rules: [rule(0, { 'fill.pattern': 'solid' }, ['fill'])],
    }
    const st = stamped(rules)
    const src = () => versionedSrc([{ row: 0, col: 0, value: nc(5) }], rules)
    const wide = await prepareCfRemote(cfRequest([rect(0, 0)], st, 1 << 24), SHEET, src(), rules, createCfMemory({}), { budget: frozenBudget })
    expect(wide.kind).toBe('ready')
    if (wide.kind !== 'ready') throw new Error('expected ready')
    expect(wide.diagnostics.length).toBeGreaterThan(0)
    expect(wide.states).toHaveLength(1)
    expect(wide.states[0]!.unresolvedReason).toBeUndefined()
    // Base is the walked probe result without diagnostics plus the envelope —
    // the same measure the provider checks — so an exactly-fitting grant
    // exercises the omit path instead of refusing.
    const probeBase = measurePayloadBytes({ ...wide, diagnostics: [] }, () => true)
    if (probeBase === undefined) throw new Error('expected measurable probe')
    const base = probeBase + CF_REPLY_ENVELOPE_BYTES
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

  it('stale definitions version is dropped and refunded with the live stamp', async () => {
    // A known-but-wrong cfGen is withheld (not painted); the reply carries the
    // live version so the caller re-requests current definitions.
    const memory = createCfMemory({})
    const st = stamped(oneRule)
    const stale = { ...st, cfGen: st.cfGen + 1, workbookValueGen: st.workbookValueGen }
    const src = fakeSource([{ row: 0, col: 0, value: nc(1) }], { stamp: st })
    let visits = 0
    const watched = { ...src, pages: async function* (...args: [string, readonly never[], CfWorkBudget, AbortSignal]) { visits++; yield* src.pages(...args) } }
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], stale, 1_000_000), SHEET, watched as never, oneRule, memory, { budget: frozenBudget })
    expect(out).toMatchObject({ kind: 'unresolved', reason: 'cf-stale-generation' })
    if (out.kind !== 'unresolved') throw new Error('expected unresolved')
    expect(out.stamp.cfGen).toBe(stamped(oneRule).cfGen)
    expect(visits).toBe(0)
    expect(memory.currentBytes).toBe(0)
  })

  it('unknown definitions version adopts live and completes', async () => {
    // cfGen 0 means "unknown": the provider builds against the live rule set
    // and reports the adopted version, enabling cold-start handshakes.
    const memory = createCfMemory({})
    const unknown = { ...stamp(1), cfGen: 0 }
    const src = fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: stamped(oneRule) })
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], unknown, 1 << 24), SHEET, src, oneRule, memory, { budget: frozenBudget })
    expect(out.kind).toBe('ready')
    if (out.kind !== 'ready') throw new Error('expected ready')
    expect(out.definitionsVersion).toBe(cfRulesVersion(oneRule))
    expect(out.stamp.cfGen).toBe(cfRulesVersion(oneRule))
    // The committed share owns the definitions until explicitly released.
    releaseSharedCfDefinitions(oneRule)
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
    const out = await prepareCfRemote(cfRequest([rect(0, 49)], stamped(oneRule), 8 * 1024 * 1024), SHEET, gated as never, oneRule, memory, { budget: frozenBudget })
    expect(out).toMatchObject({ kind: 'pending', reason: 'cf-pending-data' })
    if (out.kind !== 'pending') throw new Error('expected pending')
    expect(out.pendingSheetIds).toEqual([SHEET])
    expect(counter.n).toBe(0)
    expect(memory.currentBytes).toBe(0)
  })

  it('paged store bridge chunks and admits bounded pages', async () => {
    // 130 available cells split into 64-cell admitted pages; out-of-scope
    // cells never enter a page; every page is admitted before construction
    // and released by the consumer. Unavailable cells are covered below.
    const memory = createCfMemory({})
    const released: number[] = []
    const store: CfPagedReadHost = {
      stamp: { docId: 'doc', sheetId: SHEET, valueGen: 1 },
      readRows: async (range, cursor) => {
        if (cursor !== undefined) throw new Error('unexpected cursor in single-page fixture')
        const items: { row: number; cells: { id: number; row: number; col: number; value: number; styleIndex: number }[] }[] = []
        for (let row = 0; row < 130; row++) {
          if (row < range.r0 || row > range.r1) continue
          items.push({ row, cells: [{ id: row * 16384, row, col: 0, value: row, styleIndex: 0 }] })
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
    expect(seen).toHaveLength(130)
    expect(released).toEqual([64, 64, 2])
    expect(memory.currentBytes).toBe(0)
  })

  it('withholds on incomplete pages instead of publishing partial statistics', async () => {
    // A page that is still importing must not certify its partial rows: the
    // bridge withholds incomplete pages as pending rather than letting the
    // service compute a partial mean.
    const memory = createCfMemory({})
    const store: CfPagedReadHost = {
      stamp: { docId: 'doc', sheetId: SHEET, valueGen: 1 },
      readRows: async () => ({
        items: Array.from({ length: 10 }, (_, row) => ({ row, cells: [{ id: row * 16384, row, col: 0, value: row, styleIndex: 0 }] })),
        completeness: 'incomplete' as const,
        stamp: store.stamp,
      }),
    }
    const source: CfValueSource = {
      bounds: async () => ({ firstRow: 0, lastRow: 49, firstCol: 0, lastCol: 0 }),
      stamp: (_id: string) => ({ owner: 'owner', sheetId: SHEET, workbookValueGen: 1, cfGen: cfRulesVersion(oneRule), viewGen: 1, structureGen: 1, clockGen: 1 }),
      pages: pagedStoreCfPages(store, memory),
    }
    const out = await prepareCfRemote(cfRequest([rect(0, 49)], stamped(oneRule), 8 * 1024 * 1024), SHEET, source, oneRule, memory, { budget: frozenBudget })
    expect(out.kind).toBe('pending')
    if (out.kind !== 'pending') throw new Error('expected pending, not partial statistics')
    expect(out.pendingSheetIds).toEqual([SHEET])
    expect(memory.currentBytes).toBe(0)
  })

  it('withholds on unavailable cells instead of treating them as blank', async () => {
    // An unavailable value is unknown, not blank: erasing it would let blank
    // predicates match falsely and corrupt statistics. The bridge refuses to
    // build pages from it, and the preparation withholds (never publishes a
    // partial result), with memory fully refunded.
    const memory = createCfMemory({})
    const store: CfPagedReadHost = {
      stamp: { docId: 'doc', sheetId: SHEET, valueGen: 1 },
      readRows: async (range) => {
        const items: { row: number; cells: { id: number; row: number; col: number; value: number | null; styleIndex: number; unavailable?: true }[] }[] = []
        for (let row = range.r0; row <= range.r1; row++) {
          if (row === 0) items.push({ row, cells: [{ id: 0, row, col: 0, value: 5, styleIndex: 0 }] })
          else items.push({ row, cells: [{ id: row * 16384, row, col: 0, value: null, styleIndex: 0, unavailable: true }] })
        }
        return { items, completeness: 'complete' as const, stamp: store.stamp }
      },
    }
    const source = {
      bounds: async () => ({ firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 0 }),
      stamp: (_id: string) => stamped(oneRule),
      pages: pagedStoreCfPages(store, memory),
    }
    const out = await prepareCfRemote(cfRequest([rect(0, 1)], stamped(oneRule), 8 * 1024 * 1024), SHEET, source, oneRule, memory, { budget: frozenBudget })
    expect(out.kind).toBe('unresolved')
    if (out.kind !== 'unresolved') throw new Error('expected withheld preparation')
    expect(out.reason).toBe('cf-unavailable-values')
    expect(out.states).toEqual([])
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
      const calls = { n: 0 }
      const releases = { n: 0 }
      const base = fakeSource(
        Array.from({ length: 200 }, (_, i) => ({ row: i, col: 0, value: nc(i) })),
        { stamp: stamped(oneRule), release: () => { releases.n++ } },
      )
      // Gate the first page so the scan provably blocks inside next() with
      // admitted workspace held; aborting then draining must unwind exactly
      // once without publishing.
      let gated = false
      const src = {
        ...base,
        pages: async function* (...args: [string, readonly never[], CfWorkBudget, AbortSignal]) {
          calls.n++
          if (!gated) { gated = true; await gate.promise }
          yield* base.pages(...args)
        },
      }
      const controller = new AbortController()
      const pending = prepareCfRemote(cfRequest([rect(0, 199)], stamped(oneRule), 8 * 1024 * 1024), SHEET, src as never, oneRule, memory, { signal: controller.signal, budget: frozenBudget })
      // Wait until the service is blocked inside pages.next() (deterministic:
      // the gate is the only await before the first yield).
      for (let i = 0; i < 1000 && calls.n === 0; i++) await new Promise(r => setTimeout(r, 1))
      expect(calls.n).toBeGreaterThan(0)
      expect(memory.currentBytes).toBeGreaterThan(0)
      const held = memory.currentBytes
      if (way === 'timeout') await new Promise(r => setTimeout(r, 5))
      // 'close' on a one-shot preparation is an abort: there is no persistent
      // handle to close, only the in-flight cancellation to honor.
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
    admit: boolean
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
      admit: true,
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
      admitEntry: (key: string, bytes: number) => {
        if (overrides.admit === false) return null
        ledger.set(`entry:${key}`, bytes)
        return () => { ledger.delete(`entry:${key}`) }
      },
      changed: (sheetId: string, rects: readonly SourceRect[], reason?: 'updated' | 'evicted' | 'status') => {
        void reason
        host.changedCalls.push({ sheetId, rects, reason })
      },
      respond: overrides.respond ?? (async (command: CommandOf<'read'>) => {
        const q = command.query as { op: 'conditionalFormatting'; rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const grant = host.granted.get(command.requestId) ?? 1 << 20
        // Echo the requested identity like a correct worker (which validates
        // before preparing); tests that need mismatches override respond.
        const result = wireResult({ stamp: { ...q.stamp } })
        const reply = cfReply(result, 1024, command.requestId, command.owner as CacheIdentity)
        host.granted.delete(command.requestId)
        ledger.delete(`grant:${command.requestId}`)
        const moved = Math.min(1024, grant)
        ledger.set(reply, moved)
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
    // The reply charge settles at publish; dropping releases the components.
    expect(host.releasedReplies).toHaveLength(1)
    a.release()
    expect(host.releasedReplies).toHaveLength(1)
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
      // Each reply settles at its own publish; drops release components.
      expect(host.releasedReplies).toHaveLength(2)
      held.release()
      // ...then releases exactly once after the last release + invalidation.
      cache.invalidate(SHEET)
      await flush()
      expect(host.releasedReplies).toHaveLength(2)
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
    expect(host.releasedReplies).toHaveLength(2)
    old.release()
    cache.invalidate(SHEET)
    await flush()
    // Old entry detached on replacement; components released, no new settle.
    expect(host.releasedReplies).toHaveLength(2)
    fresh.release()
    cache.invalidate(SHEET)
    await flush()
    expect(host.releasedReplies).toHaveLength(2)
    expect(cache.retainedBytes).toBe(0)
    cache.close()
  })

  it('evicts the eldest entry exactly once at the budget edge', async () => {
    const host = testHost()
    // Single-charge budget: one entry (metadata + walked states) plus the
    // shared definitions, with too little headroom for a second entry. The
    // shared array is counted once no matter how many scopes use it.
    const probe = wireResult()
    const oneEntry = entryBytesFor(SHEET, [rect(0, 0)], cfStamp(1), probe)
    const defsBytes = measurePayloadBytes(probe.definitions, () => true)!
    const cache = new RemoteCfCache(host, oneEntry + defsBytes + 100)
    const a = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(a.state.state).toBe('ready')
    expect(cache.retainedBytes).toBe(oneEntry + defsBytes)
    a.release()
    // Released but still cached: a second consumer reuses it without a request.
    const a2 = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(host.requests).toHaveLength(1)
    expect(a2.state.state).toBe('ready')
    const sharedDefs = a2.state.result!.definitions
    a2.release()
    // A new scope evicts the eldest unreferenced entry; the shared
    // definitions survive the eviction (new consumer of the same version).
    const b = cache.acquire(SHEET, [rect(10, 10)], cfStamp(1))
    await flush()
    expect(b.state.state).toBe('ready')
    expect(b.state.result!.definitions).toBe(sharedDefs)
    expect(cache.retainedBytes).toBe(oneEntry + defsBytes)
    expect(host.releasedReplies).toHaveLength(2)
    // While everything is pinned, a new scope fails closed instead of
    // evicting live data or exceeding the budget (its reply still released).
    const c = cache.acquire(SHEET, [rect(20, 20)], cfStamp(1))
    await flush()
    expect(c.state.state).toBe('failed')
    expect(c.state.diagnostics[0]!.feature).toBe('cf-cache-budget')
    // The withheld entry keeps its admitted metadata; the shared version
    // stays charged exactly once for the live entry.
    expect(cache.retainedBytes).toBe(oneEntry + defsBytes + entryMetaBytes(SHEET, [rect(20, 20)], cfStamp(1)))
    expect(host.releasedReplies).toHaveLength(3)
    b.release()
    c.release()
    cache.close()
    expect(host.releasedReplies).toHaveLength(3)
    expect(cache.retainedBytes).toBe(0)
  })

  it('abort, timeout and close each release exactly once with live-ledger assertions', async () => {
    for (const way of ['abort', 'timeout', 'close'] as const) {
      // Transport-faithful fake: abort resolves the waiter at once, but the
      // grant stays with the transport until the deferred producer drains
      // (settling it, like viewerRemoteSource.request's sent.then). The
      // cache must never refund early.
      const producers = new Map<string, (v: WorkerReply<'read'> | symbol) => void>()
      const host = testHost({
        respond: (command, signal) => new Promise<WorkerReply<'read'> | symbol>(resolve => {
          producers.set(command.requestId, resolve)
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
      // entry drops (evict/invalidate/close). Resolving the late waiter must
      // never publish, and the grant stays held until the producer drains.
      const requestId = [...host.granted.keys()][0]!
      if (way === 'close') cache.close()
      else controller.abort()
      for (const [, resolve] of producers) resolve(Symbol('cancelled'))
      await flush()
      expect(host.granted.has(requestId)).toBe(true)
      held.release()
      // The late producer drains: the transport settles the grant into the
      // late reply and settles that reply at once (its waiter is gone) —
      // net grant gone, no leak, no double release.
      host.granted.delete(requestId)
      host.ledger.delete(`grant:${requestId}`)
      await flush()
      expect(host.granted.size).toBe(0)
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

  it('drops a reply whose stamp does not match the request', async () => {
    // A worker that answers with a different generation triggers one
    // bounded rediscovery: the first miss rekeys and reloads once, the
    // second miss fails closed naming the discovered version, and every
    // reply releases exactly once (no retry loop).
    const host = testHost({
      respond: async (command) => {
        const q = command.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const wrong = { ...wireResult(), stamp: { ...q.stamp, cfGen: q.stamp.cfGen + 999 }, definitionsVersion: 1 as const, states: [] as never[] };
        const reply = cfReply(wrong as CfReadResult, 1024, command.requestId, command.owner as CacheIdentity)
        host.granted.delete(command.requestId)
        host.ledger.delete(`grant:${command.requestId}`)
        host.ledger.set(reply, 1024)
        return reply
      },
    })
    const cache = new RemoteCfCache(host, 1 << 20)
    const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    await flush()
    expect(held.state.state).toBe('failed')
    // One bounded retry: the first miss (cfGen 1000) reloaded, the second
    // (cfGen 1999) failed closed and named the discovered version.
    expect(host.requests).toHaveLength(2)
    expect(held.state.diagnostics[0]).toMatchObject({ feature: 'cf-stale-generation' })
    expect(held.state.diagnostics[0]!.message).toContain('1999')
    expect(host.releasedReplies).toHaveLength(2)
    held.release()
    cache.close()
    expect(host.ledger.size).toBe(0)
  })

  it('keys entries on the full stamp including view and clock generations', async () => {
    const host = testHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const a = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(a.state.state).toBe('ready')
    // Same scope and value generation but a new view generation is a
    // different key: no stale hit, a fresh preparation is built.
    const b = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1, { viewGen: 2 }))
    await flush()
    expect(b.state.state).toBe('ready')
    expect(b.state.result).not.toBe(a.state.result)
    expect(host.requests).toHaveLength(2)
    a.release()
    b.release()
    cache.invalidate(SHEET)
    await flush()
    expect(host.releasedReplies).toHaveLength(2)
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

describe('CF fix round 1 (review-t7-r1)', () => {
  const identity = (sheetId = SHEET, valueGen = 1): CacheIdentity => ({ docId: 'doc', sheetId, valueGen, fontGen: 0, viewGen: 0 })
  const cfStamp = (gen = 1, extra: Partial<CfReadStamp> = {}): CfReadStamp => ({
    owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
  })
  const wireResult = (over: Partial<CfReadResult> = {}): CfReadResult => ({
    sheetId: SHEET, stamp: cfStamp(1), kind: 'ready', states: [{ row: 0, col: 0, properties: { 'fill.pattern': 'solid' } }],
    definitions: [{ id: 'rule-0', version: 7, properties: { 'fill.pattern': 'solid' }, writes: ['fill'] }],
    definitionsVersion: 7, diagnostics: [], ...over,
  })
  const cfReply = (result: CfReadResult, bytes: number, requestId = 'r1', owner: CacheIdentity = identity()): WorkerReply<'read'> => ({
    commandKind: 'read', requestId, owner, diagnostics: [], replyBytes: bytes, kind: 'complete',
    result: { kind: 'read', result: { op: 'conditionalFormatting', result } },
  } as WorkerReply<'read'>)
  const flush = () => new Promise(r => setTimeout(r, 0))

  it('F1: refuses on a zero-byte account before touching rule content', async () => {
    // Normalized serialization, definition copies and diagnostics must all
    // be admitted before they are built: a getter on a rule value observes
    // any pre-admission serialization or copying (key enumeration alone,
    // which the admission bound uses, never evaluates values).
    let touched = 0
    const dxfProbe: Record<string, unknown> = {}
    Object.defineProperty(dxfProbe, 'fill.pattern', { enumerable: true, configurable: true, get() { touched++; return 'solid' } })
    const evil = { ...rule(0, {}, ['fill']), dxf: dxfProbe } as unknown as CfSheetRules['rules'][number]
    const rules: CfSheetRules = { diagnostics: [], rules: [evil] }
    const reqStamp = stamped(rules)
    touched = 0
    const memory = createCfMemory({ limitBytes: 0 })
    const source = fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: reqStamp })
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], reqStamp, 8 * 1024 * 1024), SHEET, source, rules, memory, { budget: frozenBudget })
    expect(out.kind).toBe('unresolved')
    expect(touched).toBe(0)
    expect(memory.peakBytes).toBe(0)
    expect(memory.currentBytes).toBe(0)
  })

  it('F1: the page bridge admits its workspace before touching the store', async () => {
    // Chunk entries and materialized values must live inside an admitted
    // hold: with an account too small for the workspace the bridge refuses
    // before the first readRows call, not after materializing a page.
    const memory = createCfMemory({ limitBytes: 100 })
    let reads = 0
    const store: CfPagedReadHost = {
      stamp: { docId: 'doc', sheetId: SHEET, valueGen: 1 },
      readRows: async (range) => {
        reads++
        const items: { row: number; cells: { id: number; row: number; col: number; value: number; styleIndex: number }[] }[] = []
        for (let row = range.r0; row <= range.r1; row++) items.push({ row, cells: [{ id: row * 16384, row, col: 0, value: row, styleIndex: 0 }] })
        return { items, completeness: 'complete' as const, stamp: store.stamp }
      },
    }
    const pages = pagedStoreCfPages(store, memory)
    const drain = async () => {
      for await (const lease of pages('s', [{ firstRow: 0, lastRow: 9, firstCol: 0, lastCol: 0 }], frozenBudget(), new AbortController().signal)) lease.release()
    }
    await expect(drain()).rejects.toMatchObject({ code: 'resource-limit' })
    expect(reads).toBe(0)
    expect(memory.peakBytes).toBe(0)
  })

  it('F1: the session passes one shared memory account to the provider', async () => {
    // The worker pool must be able to refuse CF construction: the session
    // connects its shared CF account (wired to the document allocation when
    // present) instead of leaving the provider on an isolated per-request one.
    const { ViewerClient } = await import('../src/worker/viewerClient')
    const { connectInProcess, ViewerKernel } = await import('../src/worker/viewerKernel')
    const { createObjectWorkbookSource } = await import('../src/worker/viewerObjectSource')
    const { initialIdentity } = await import('../src/worker/viewerProtocol')
    const { workbookOf } = await import('./helpers/xlsx-find')
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 5 }] }])
    const base = createObjectWorkbookSource(doc, 'doc')
    const seen: unknown[] = []
    const wrapped = {
      ...base,
      store: (id: string) => ({
        ...base.store(id),
        readConditionalFormatting: async (request: CfReadRequest, memory: unknown) => {
          seen.push(memory)
          return {
            sheetId: id, stamp: { ...request.stamp }, kind: 'ready' as const,
            states: [], definitions: [], definitionsVersion: 0, diagnostics: [],
          }
        },
      }),
    }
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => wrapped }, { yield: async () => {} })
    const port = connectInProcess(kernel)
    const client = new ViewerClient(port)
    const open = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
    if (open.kind !== 'complete') throw new Error('open failed')
    const owner = initialIdentity('doc', 's')
    const query = (requestId: string) => ({
      kind: 'read' as const, requestId, owner,
      query: { op: 'conditionalFormatting' as const, rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], stamp: { owner: 'owner', sheetId: 's', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 1, clockGen: 0 }, grantBytes: 1 << 20 },
    })
    await client.send(query(client.nextRequestId()))
    await client.send(query(client.nextRequestId()))
    expect(seen).toHaveLength(2)
    expect(seen[0]).toBeDefined()
    // One session-scoped account, so concurrent preparations share the cap.
    expect(seen[0]).toBe(seen[1])
  })

  it('F2: abort keeps the grant with the transport until the producer drains', async () => {
    // Models viewerRemoteSource.request faithfully: aborting the in-flight
    // entry resolves the public promise with CANCELLED at once, but the
    // grant stays held until the late reply arrives and the transport
    // settles it. The cache must not refund early and let other work
    // consume the headroom first.
    const CANCELLED = Symbol('cf-test-cancelled')
    const granted = new Map<string, number>()
    const ledger = new Map<string | object, number>()
    const producers = new Map<string, (v: WorkerReply<'read'> | symbol) => void>()
    const releasedReplies: WorkerReply<'read'>[] = []
    const requests: string[] = []
    let n = 0
    const host = {
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: (requestId: string) => { granted.set(requestId, 1 << 20); ledger.set(`grant:${requestId}`, 1 << 20); return 1 << 20 },
      releaseGrant: (requestId: string) => { granted.delete(requestId); ledger.delete(`grant:${requestId}`) },
      request: (command: CommandOf<'read'>, signal?: AbortSignal) => new Promise<WorkerReply<'read'> | symbol>(resolve => {
        requests.push(command.requestId)
        producers.set(command.requestId, resolve)
        signal?.addEventListener('abort', () => resolve(CANCELLED as unknown as symbol), { once: true })
      }),
      nextRequestId: () => `r${++n}`,
      release: (reply: WorkerReply<'read'>) => { releasedReplies.push(reply); ledger.delete(reply) },
      admitEntry: (key: string, bytes: number) => { ledger.set(`entry:${key}`, bytes); return () => { ledger.delete(`entry:${key}`) } },
      changed: () => {},
      // The transport settling the late reply: the grant is adopted into the
      // reply charge and settled at once (its waiter is gone) — net zero,
      // but only once the producer actually drains.
      settleProducer: (requestId: string, reply?: WorkerReply<'read'>) => {
        const resolve = producers.get(requestId)
        producers.delete(requestId)
        granted.delete(requestId)
        ledger.delete(`grant:${requestId}`)
        if (reply && resolve) {
          ledger.set(reply, 1024)
          resolve(reply)
        }
      },
    }
    const cache = new RemoteCfCache(host, 1 << 20)
    cache.get(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    const requestId = [...granted.keys()][0]!
    expect(requests).toHaveLength(1)
    // Eviction aborts the in-flight entry: the waiter resolves CANCELLED,
    // but the transport still owns the grant — no early refund for other
    // work to consume before the reply drains.
    cache.invalidate(SHEET)
    await flush()
    expect(granted.has(requestId)).toBe(true)
    host.settleProducer(requestId)
    await flush()
    expect(granted.has(requestId)).toBe(false)
    // A live re-get starts a fresh request and settles exactly once.
    cache.get(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(requests).toHaveLength(2)
    const liveId = requests[1]!
    const reply = cfReply(wireResult(), 1024, liveId)
    host.settleProducer(liveId, reply)
    await flush()
    expect(cache.get(SHEET, [rect(0, 0)], cfStamp(1)).state).toBe('ready')
    cache.invalidate(SHEET)
    await flush()
    expect(releasedReplies).toHaveLength(1)
    expect(ledger.size).toBe(0)
    cache.close()
  })

  function readyHost(budgetBytes = 1 << 20) {
    const granted = new Map<string, number>()
    const ledger = new Map<string | object, number>()
    const releasedReplies: WorkerReply<'read'>[] = []
    const changedCalls: { sheetId: string; rects: readonly SourceRect[]; reason?: string }[] = []
    let n = 0
    return {
      granted, ledger, releasedReplies, changedCalls,
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: (requestId: string) => { granted.set(requestId, budgetBytes); ledger.set(`grant:${requestId}`, budgetBytes); return budgetBytes },
      releaseGrant: (requestId: string) => { granted.delete(requestId); ledger.delete(`grant:${requestId}`) },
      request: async (command: CommandOf<'read'>) => {
        const q = command.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const result = wireResult({ stamp: { ...q.stamp } })
        const reply = cfReply(result, 1024, command.requestId, command.owner as CacheIdentity)
        granted.delete(command.requestId)
        ledger.delete(`grant:${command.requestId}`)
        ledger.set(reply, 1024)
        return reply
      },
      nextRequestId: () => `r${++n}`,
      release: (reply: WorkerReply<'read'>) => { releasedReplies.push(reply); ledger.delete(reply) },
      admitEntry: (key: string, bytes: number) => { ledger.set(`entry:${key}`, bytes); return () => { ledger.delete(`entry:${key}`) } },
      changed: (sheetId: string, rects: readonly SourceRect[], reason?: 'updated' | 'evicted' | 'status') => { changedCalls.push({ sheetId, rects, reason }) },
    }
  }

  it('F6: entry metadata is charged and detached versions stay charged', async () => {
    const host = readyHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(held.state.state).toBe('ready')
    // Single charge: entry metadata plus the walked states; the shared
    // definitions live once in the registry, not per entry.
    const probe = wireResult()
    const one = entryBytesFor(SHEET, [rect(0, 0)], cfStamp(1), probe)
    const defs = measurePayloadBytes(probe.definitions, () => true)!
    expect(cache.retainedBytes).toBe(one + defs)
    const fresh = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1, { cfGen: 2 }))
    await flush()
    expect(fresh.state.state).toBe('ready')
    // Both versions charged while both are held; the shared version counted once.
    expect(fresh.state.result!.definitions).toBe(held.state.result!.definitions)
    expect(cache.retainedBytes).toBe(2 * one + defs)
    expect(host.releasedReplies).toHaveLength(2)
    held.release()
    cache.invalidate(SHEET)
    await flush()
    expect(cache.retainedBytes).toBe(one + defs)
    expect(host.releasedReplies).toHaveLength(2)
    fresh.release()
    cache.invalidate(SHEET)
    await flush()
    expect(host.releasedReplies).toHaveLength(2)
    expect(cache.retainedBytes).toBe(0)
    cache.close()
  })

  it('F6: the pinned entry cap refuses closed instead of growing without bound', async () => {
    // Every entry pinned by a live consumer counts toward the cap: once 64
    // entries are live, a new scope fails closed without sending a request.
    // The shared definitions are counted once no matter how many scopes use them.
    const host = readyHost()
    const probe = wireResult()
    const one = entryBytesFor(SHEET, [rect(0, 0)], cfStamp(1), probe)
    const defs = measurePayloadBytes(probe.definitions, () => true)!
    const cache = new RemoteCfCache(host, 64 * one + defs + (1 << 20))
    const held: { release(): void }[] = []
    for (let i = 0; i < 64; i++) {
      const h = cache.acquire(SHEET, [rect(i, i)], cfStamp(1))
      held.push(h)
    }
    await flush()
    await flush()
    // One settle per publish; drops release components without re-settling.
    expect(host.releasedReplies).toHaveLength(64)
    expect(cache.retainedBytes).toBe(64 * one + defs)
    expect(cache.sharedDefinitionVersions).toBe(1)
    const extra = cache.acquire(SHEET, [rect(1000, 1000)], cfStamp(1))
    await flush()
    expect(extra.state.state).toBe('failed')
    expect(extra.state.diagnostics[0]).toMatchObject({ feature: 'cf-cache-budget' })
    extra.release()
    for (const h of held) h.release()
    cache.invalidate(SHEET)
    await flush()
    expect(host.releasedReplies).toHaveLength(64)
    expect(cache.retainedBytes).toBe(0)
    expect(cache.sharedDefinitionVersions).toBe(0)
    cache.close()
  })

  it('F6: close keeps held entries charged until their last release', async () => {
    const host = readyHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(held.state.state).toBe('ready')
    cache.close()
    // The held preparation stays charged across close; no new reads start.
    // (The reply settled at publish; the retained states keep their charge.)
    expect(cache.retainedBytes).toBeGreaterThan(0)
    expect(host.releasedReplies).toHaveLength(1)
    // New consumers after close get nothing and start no requests.
    const after = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    expect(after.state.state).toBe('waiting')
    after.release()
    held.release()
    expect(host.releasedReplies).toHaveLength(1)
    expect(cache.retainedBytes).toBe(0)
  })

  it('F6: version replacement under a byte cap fails closed with pinned charged', async () => {
    // A tight budget that fits one preparation plus the shared definitions
    // but not a second states payload: replacing the version withholds the
    // new one while the pinned old one stays charged (shared defs once).
    const host = readyHost()
    const probe = wireResult()
    const statesBytes = measurePayloadBytes({ ...probe, definitions: [] }, () => true)!
    const meta = entryMetaBytes(SHEET, [rect(0, 0)], cfStamp(1))
    const defsBytes = measurePayloadBytes(probe.definitions, () => true)!
    const one = meta + statesBytes
    const cache = new RemoteCfCache(host, one + defsBytes + meta + 10)
    const old = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(old.state.state).toBe('ready')
    expect(cache.retainedBytes).toBe(one + defsBytes)
    const next = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1, { cfGen: 2 }))
    await flush()
    expect(next.state.state).toBe('failed')
    expect(next.state.diagnostics[0]).toMatchObject({ feature: 'cf-cache-budget' })
    // The withheld reply released at once; the shared version stays charged
    // exactly once for the pinned old entry, plus the new entry's metadata.
    expect(host.releasedReplies).toHaveLength(2)
    expect(cache.retainedBytes).toBe(one + defsBytes + meta)
    old.release()
    next.release()
    cache.invalidate(SHEET)
    await flush()
    expect(host.releasedReplies).toHaveLength(2)
    expect(cache.retainedBytes).toBe(0)
    cache.close()
  })

  it('F7: empty scopes refuse tiny grants and serve blanks through the service', async () => {
    const { PagedXlsxSheetAdapter } = await import('../src/xlsx/view/pagedSource')
    const sourceStamp = { docId: 'doc', sheetId: SHEET, valueGen: 1 }
    const fakeSheet = (over: Record<string, unknown>) => ({
      sheetId: SHEET, name: SHEET, state: 'complete' as const, importedThroughRow: 1048575,
      metadata: {}, tables: undefined, arrayCells: [], drawingAnchors: { records: [], unresolved: 0 },
      mergeRects: [], invalidMerges: 0, used: undefined, conditionalFormatting: undefined,
      ...over,
    })
    // A rule over blank cells is grant-gated like any other preparation:
    // a one-byte grant refuses instead of returning ready.
    const tiny = new PagedXlsxSheetAdapter(sourceStamp, {} as never, fakeSheet({ conditionalFormatting: oneRule }) as never, undefined, {}, false)
    const refused = await tiny.readConditionalFormatting(
      cfRequest([rect(0, 0)], stamped(oneRule), 1), createCfMemory({}), { budget: frozenBudget },
    )
    expect(refused.kind).toBe('unresolved')
    expect(refused).toMatchObject({ reason: 'cf-resource-limit' })
    await tiny.release().catch(() => {})
    // And an empty scope with room evaluates through the common service
    // (no shortcut): blank cells compose to a ready preparation.
    const memory = createCfMemory({})
    const source = fakeSource([], { stamp: stamped(oneRule), bounds: { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 16383 } })
    const ready = await prepareCfRemote(cfRequest([rect(0, 0)], stamped(oneRule), 8 * 1024 * 1024), SHEET, source, oneRule, memory, { budget: frozenBudget })
    expect(ready.kind).toBe('ready')
    // The committed share owns the definitions until explicitly released.
    releaseSharedCfDefinitions(oneRule)
    expect(memory.currentBytes).toBe(0)
  })

  it('F8: identical rule sets share one frozen definition array', async () => {
    const memory = createCfMemory({})
    const mkSource = () => fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: stamped(oneRule) })
    const a = await prepareCfRemote(cfRequest([rect(0, 0)], stamped(oneRule), 8 * 1024 * 1024), SHEET, mkSource(), oneRule, memory, { budget: frozenBudget })
    const b = await prepareCfRemote(cfRequest([rect(5, 5)], stamped(oneRule), 8 * 1024 * 1024), SHEET, mkSource(), oneRule, memory, { budget: frozenBudget })
    expect(a.kind).toBe('ready')
    expect(b.kind).toBe('ready')
    if (a.kind !== 'ready' || b.kind !== 'ready') throw new Error('expected ready preparations')
    // Two viewport scopes, unchanged rules: one shared definition array.
    expect(b.definitions).toBe(a.definitions)
    expect(Object.isFrozen(a.definitions)).toBe(true)
    // A changed rule set builds a new array; the old one stays intact.
    const changed: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'fill.pattern': 'dashed' }, ['fill'])] }
    const cSource = fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: stamped(changed) })
    const c = await prepareCfRemote(cfRequest([rect(0, 0)], stamped(changed), 8 * 1024 * 1024), SHEET, cSource, changed, memory, { budget: frozenBudget })
    expect(c.kind).toBe('ready')
    if (c.kind !== 'ready') throw new Error('expected ready preparation')
    expect(c.definitions).not.toBe(a.definitions)
    expect(a.definitions).toHaveLength(1)
    // Both committed shares own their definitions until explicitly released.
    releaseSharedCfDefinitions(oneRule)
    releaseSharedCfDefinitions(changed)
    expect(memory.currentBytes).toBe(0)
  })

  it('F9: CF dirty-rectangle handoff stays bounded', async () => {
    const { accumulateCfDirtyRects } = await import('../src/worker/viewerRemoteSource')
    const into = new Map<string, SourceRect[]>()
    for (let i = 0; i < 200; i++) accumulateCfDirtyRects(into, SHEET, [rect(i, i)])
    const kept = into.get(SHEET) ?? []
    // Bounded: overflow coalesces to covering rectangles along the way.
    expect(kept.length).toBeLessThanOrEqual(64)
    expect(kept.length).toBeGreaterThan(0)
    // Nothing is lost: every notified row stays covered by some kept rect.
    for (let i = 0; i < 200; i++) expect(kept.some(r => r.r0 <= i && r.r1 >= i)).toBe(true)
  })
})

describe('CF fix round 2 (review-t7-r2)', () => {
  const identity = (sheetId = SHEET, valueGen = 1): CacheIdentity => ({ docId: 'doc', sheetId, valueGen, fontGen: 0, viewGen: 0 })
  const cfStamp = (gen = 1, extra: Partial<CfReadStamp> = {}): CfReadStamp => ({
    owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
  })
  const wireResult = (over: Partial<CfReadResult> = {}): CfReadResult => ({
    sheetId: SHEET, stamp: cfStamp(1), kind: 'ready', states: [{ row: 0, col: 0, properties: { 'fill.pattern': 'solid' } }],
    definitions: [{ id: 'rule-0', version: 7, properties: { 'fill.pattern': 'solid' }, writes: ['fill'] }],
    definitionsVersion: 7, diagnostics: [], ...over,
  })
  const cfReply = (result: CfReadResult, bytes: number, requestId = 'r1', owner: CacheIdentity = identity()): WorkerReply<'read'> => ({
    commandKind: 'read', requestId, owner, diagnostics: [], replyBytes: bytes, kind: 'complete',
    result: { kind: 'read', result: { op: 'conditionalFormatting', result } },
  } as WorkerReply<'read'>)
  const flush = () => new Promise(r => setTimeout(r, 0))

  it('R2-1: the shared CF account refuses when the ordinary worker pool is full', async () => {
    // The session-scoped CF account must be adapted to the real session
    // MemoryPool, not an independent ledger: filling the worker pool denies
    // CF construction before any build, and one account is shared.
    const { DocumentSession } = await import('../src/worker/viewerSession')
    const session = new DocumentSession('d', { sheets: [{ sheetId: SHEET }], store: () => ({}) } as never)
    try {
      const account = await (session as unknown as { cfAccount(): Promise<import('../src/xlsx/cf/limits').CfMemory> }).cfAccount()
      const again = await (session as unknown as { cfAccount(): Promise<import('../src/xlsx/cf/limits').CfMemory> }).cfAccount()
      expect(again).toBe(account)
      const hold = session.memory.charge()
      expect(hold.set(session.memory.limitBytes)).toBe(true)
      try {
        expect(account.reserve(1048576).kind).toBe('unresolved')
        expect(account.currentBytes).toBe(0)
      } finally {
        hold.release()
      }
      // With room, the same shared account admits.
      const admitted = account.reserve(1048576)
      expect(admitted.kind).toBe('reserved')
      if (admitted.kind === 'reserved') admitted.release()
      expect(account.currentBytes).toBe(0)
    } finally {
      await session.release().catch(() => {})
    }
  })

  it('R2-2a: the adapter refuses on a starved account before hashing rules', async () => {
    // Even the adapter's generation hash is construction: a value getter
    // observes any pre-admission read, and a zero-byte account must refuse
    // before it fires.
    const { PagedXlsxSheetAdapter } = await import('../src/xlsx/view/pagedSource')
    let touched = 0
    const dxfProbe: Record<string, unknown> = {}
    Object.defineProperty(dxfProbe, 'fill.pattern', { enumerable: true, configurable: true, get() { touched++; return 'solid' } })
    const evil = { ...rule(0, {}, ['fill']), dxf: dxfProbe } as unknown as CfSheetRules['rules'][number]
    const rules: CfSheetRules = { diagnostics: [], rules: [evil] }
    const st = stamped(rules)
    touched = 0
    const sourceStamp = { docId: 'doc', sheetId: SHEET, valueGen: 1 }
    const fakeSheet = {
      sheetId: SHEET, name: SHEET, state: 'complete' as const, importedThroughRow: 1048575,
      metadata: {}, tables: undefined, arrayCells: [], drawingAnchors: { records: [], unresolved: 0 },
      mergeRects: [], invalidMerges: 0, used: undefined, conditionalFormatting: rules,
    }
    const adapter = new PagedXlsxSheetAdapter(sourceStamp, {} as never, fakeSheet as never, undefined, {}, false)
    try {
      const memory = createCfMemory({ limitBytes: 0 })
      const out = await adapter.readConditionalFormatting(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), memory, { budget: frozenBudget })
      expect(out.kind).toBe('unresolved')
      expect(touched).toBe(0)
      expect(memory.peakBytes).toBe(0)
      expect(memory.currentBytes).toBe(0)
    } finally {
      await adapter.release().catch(() => {})
    }
  })

  it('R2-2b: serialization is sized exactly before it is built', async () => {
    // A flat per-key allowance cannot bound strings: an 8 KiB format code
    // serializes to ~16.8 KiB, so the hold must already cover it when the
    // string is built. Escapable content (quotes, backslashes, controls)
    // must be covered too, or the byte-exact backstop withholds.
    const longRules: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'numFmt.code': 'X'.repeat(8192) }, ['numFmt'])] }
    const st = stamped(longRules)
    const memory = createCfMemory({})
    const source = fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: st })
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), SHEET, source, longRules, memory, { budget: frozenBudget })
    expect(out.kind).toBe('ready')
    // The peak hold covered the serialized rules, not just a flat allowance.
    expect(memory.peakBytes).toBeGreaterThanOrEqual(8192 * 2)
    // The committed share owns the definitions until explicitly released.
    releaseSharedCfDefinitions(longRules)
    expect(memory.currentBytes).toBe(0)
    const escapeRules: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'numFmt.code': 'a"b\\c\x01d'.repeat(1000) }, ['numFmt'])] }
    const est = stamped(escapeRules)
    const ememory = createCfMemory({})
    const esource = fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: est })
    const eout = await prepareCfRemote(cfRequest([rect(0, 0)], est, 8 * 1024 * 1024), SHEET, esource, escapeRules, ememory, { budget: frozenBudget })
    expect(eout.kind).toBe('ready')
    releaseSharedCfDefinitions(escapeRules)
    expect(ememory.currentBytes).toBe(0)
  })

  it('R2-2d: a ready provider transfers a live outgoing lease to the caller', async () => {
    // No window may exist where the live states sit uncharged: with an
    // outbox the account stays charged after return until the caller
    // releases; a second scope reuses the committed share without a new
    // lease on its own account.
    const st = stamped(oneRule)
    const mkSource = () => fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: st })
    const memory = createCfMemory({})
    const releases: Array<() => void> = []
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), SHEET, mkSource(), oneRule, memory,
      { budget: frozenBudget, retainOutgoing: release => { releases.push(release) } })
    expect(out.kind).toBe('ready')
    expect(releases).toHaveLength(1)
    expect(memory.currentBytes).toBeGreaterThan(0)
    // Draining the outgoing lease keeps the table's own charge; releasing
    // the share settles everything.
    releases[0]!()
    expect(memory.currentBytes).toBeGreaterThan(0)
    releaseSharedCfDefinitions(oneRule)
    expect(memory.currentBytes).toBe(0)
    const plain = createCfMemory({})
    const fresh: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'fill.pattern': 'dashed' }, ['fill'])] }
    const unleased = await prepareCfRemote(cfRequest([rect(0, 0)], stamped(fresh), 8 * 1024 * 1024), SHEET,
      fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: stamped(fresh) }), fresh, plain, { budget: frozenBudget })
    expect(unleased.kind).toBe('ready')
    // No outbox: the committed table entry still owns the shared definitions.
    expect(plain.currentBytes).toBeGreaterThan(0)
    releaseSharedCfDefinitions(fresh)
    expect(plain.currentBytes).toBe(0)
    releaseSharedCfDefinitions(oneRule)
    expect(memory.currentBytes).toBe(0)
  })

  it('R2-3: cancellation and build failures release the admitted hold', async () => {
    // The reservation starts before the bounds await and normalization, so
    // every throw on the way out must unwind it: a cancelled bounds lookup
    // rejects with nothing charged, and a throwing rule value withholds
    // with nothing charged.
    const gate = (() => { let resolve!: () => void; const promise = new Promise<void>(r => { resolve = r }); return { promise, resolve } })()
    const cancelling = fakeSource([{ row: 0, col: 0, value: nc(1) }], { stamp: stamped(oneRule) })
    const controller = new AbortController()
    const suspending = {
      ...cancelling,
      bounds: async () => { await gate.promise; return cancelling.bounds('s') },
    }
    const cmemory = createCfMemory({})
    const pending = prepareCfRemote(
      cfRequest([rect(0, 0)], stamped(oneRule), 8 * 1024 * 1024), SHEET, suspending, oneRule, cmemory,
      { budget: frozenBudget, signal: controller.signal },
    )
    await new Promise(r => setTimeout(r, 0))
    controller.abort()
    gate.resolve()
    await expect(pending).rejects.toMatchObject({ code: 'cancelled' })
    expect(cmemory.currentBytes).toBe(0)
    const throwing = fakeSource([{ row: 0, col: 0, value: nc(1) }], { stamp: stamped(oneRule) })
    const badDxf: Record<string, unknown> = {}
    Object.defineProperty(badDxf, 'fill.pattern', { enumerable: true, configurable: true, get() { throw new Error('probe construction failure') } })
    const badRules: CfSheetRules = { diagnostics: [], rules: [{ ...rule(0, {}, ['fill']), dxf: badDxf } as unknown as CfSheetRules['rules'][number]] }
    const bmemory = createCfMemory({})
    const bout = await prepareCfRemote(
      cfRequest([rect(0, 0)], stamped(oneRule), 8 * 1024 * 1024), SHEET, throwing, badRules, bmemory, { budget: frozenBudget },
    )
    expect(bout.kind).toBe('unresolved')
    expect(bmemory.currentBytes).toBe(0)
    expect(bmemory.peakBytes).toBeGreaterThan(0)
  })

  it('R2-5: shared definitions are deep-frozen, committed on success, and released on replacement', async () => {
    // Cross-scope identity without ownership leaks: frozen at every level,
    // cached only by successful preparations, and dropped when the sheet
    // replaces (or the store releases) its rules.
    const { PagedXlsxSheetAdapter, sharedCfDefinitionStats } =
      await import('../src/xlsx/view/pagedSource')
    // A dedicated rule set keeps the hit counts deterministic across tests
    // that share the module-level fixtures.
    const base: CfSheetRules = { diagnostics: [], rules: [rule(7, { 'fill.pattern': 'solid' }, ['fill'])] }
    const st = stamped(base)
    const mkSource = () => fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: st })
    const first = await prepareCfRemote(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), SHEET, mkSource(), base, createCfMemory({}), { budget: frozenBudget })
    expect(first.kind).toBe('ready')
    if (first.kind !== 'ready') throw new Error('expected ready')
    // Deep freeze: array, definitions, properties and writes alike.
    expect(Object.isFrozen(first.definitions)).toBe(true)
    expect(Object.isFrozen(first.definitions[0])).toBe(true)
    expect(Object.isFrozen(first.definitions[0]!.properties)).toBe(true)
    expect(Object.isFrozen(first.definitions[0]!.writes)).toBe(true)
    expect(sharedCfDefinitionStats(base)).toMatchObject({ hits: 1 })
    // A second scope reuses the committed identity.
    const second = await prepareCfRemote(cfRequest([rect(5, 5)], st, 8 * 1024 * 1024), SHEET, mkSource(), base, createCfMemory({}), { budget: frozenBudget })
    expect(second.kind).toBe('ready')
    if (second.kind !== 'ready') throw new Error('expected ready')
    expect(second.definitions).toBe(first.definitions)
    expect(sharedCfDefinitionStats(base)).toMatchObject({ hits: 2 })
    // Refusals never populate the table: a grant-starved scope with fresh
    // rules leaves no share behind.
    const fresh: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'fill.pattern': 'dashed' }, ['fill'])] }
    const refused = await prepareCfRemote(cfRequest([rect(0, 0)], stamped(fresh), 1), SHEET, fakeSource([], { stamp: stamped(fresh) }), fresh, createCfMemory({}), { budget: frozenBudget })
    expect(refused.kind).toBe('unresolved')
    expect(sharedCfDefinitionStats(fresh)).toBeUndefined()
    // Replacement drops the old share without touching pages: the adapter
    // observes the base rules, then replacement rules, releasing the old
    // identity on a grant-starved call that never reaches the page bridge.
    const sourceStamp = { docId: 'doc', sheetId: SHEET, valueGen: 1 }
    const sheet = {
      sheetId: SHEET, name: SHEET, state: 'complete' as const, importedThroughRow: 1048575,
      metadata: {}, tables: undefined, arrayCells: [], drawingAnchors: { records: [], unresolved: 0 },
      mergeRects: [], invalidMerges: 0, used: { r0: 0, r1: 0, c0: 0, c1: 0 }, conditionalFormatting: base as CfSheetRules | undefined,
    }
    const adapter = new PagedXlsxSheetAdapter(sourceStamp, {} as never, sheet as never, undefined, {}, false)
    const replacement: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'fill.pattern': 'solid' }, ['fill'])] }
    try {
      const seen = await adapter.readConditionalFormatting(cfRequest([rect(0, 0)], st, 1), createCfMemory({}), { budget: frozenBudget })
      expect(seen.kind).toBe('unresolved')
      expect(sharedCfDefinitionStats(base)).toBeDefined()
      const committed = await prepareCfRemote(cfRequest([rect(0, 0)], stamped(replacement), 8 * 1024 * 1024), SHEET,
        fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: stamped(replacement) }), replacement, createCfMemory({}), { budget: frozenBudget })
      expect(committed.kind).toBe('ready')
      sheet.conditionalFormatting = replacement
      const rst = stamped(replacement)
      const third = await adapter.readConditionalFormatting(cfRequest([rect(0, 0)], rst, 1), createCfMemory({}), { budget: frozenBudget })
      expect(third.kind).toBe('unresolved')
      expect(sharedCfDefinitionStats(base)).toBeUndefined()
      expect(sharedCfDefinitionStats(replacement)).toBeDefined()
    } finally {
      await adapter.release().catch(() => {})
    }
    expect(sharedCfDefinitionStats(base)).toBeUndefined()
    expect(sharedCfDefinitionStats(replacement)).toBeUndefined()
  })

  it('R2-5: shared definitions survive a structured-clone boundary per scope', async () => {
    // The main-side contract across the worker boundary: each scope's reply
    // clones cleanly and versions identically; bytes stay charged per reply.
    const memory = createCfMemory({})
    const mkSource = () => fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: stamped(oneRule) })
    const a = await prepareCfRemote(cfRequest([rect(0, 0)], stamped(oneRule), 8 * 1024 * 1024), SHEET, mkSource(), oneRule, memory, { budget: frozenBudget })
    const b = await prepareCfRemote(cfRequest([rect(5, 5)], stamped(oneRule), 8 * 1024 * 1024), SHEET, mkSource(), oneRule, memory, { budget: frozenBudget })
    expect(a.kind).toBe('ready')
    expect(b.kind).toBe('ready')
    if (a.kind !== 'ready' || b.kind !== 'ready') throw new Error('expected ready')
    const ca = structuredClone(a)
    const cb = structuredClone(b)
    expect(ca.definitionsVersion).toBe(cb.definitionsVersion)
    expect(ca.definitions).toStrictEqual(cb.definitions)
    expect(ca.states).toHaveLength(1)
    expect(cb.states).toHaveLength(1)
    // The committed share owns the definitions until explicitly released.
    releaseSharedCfDefinitions(oneRule)
    expect(memory.currentBytes).toBe(0)
  })

  it('R2-7: an over-cap sentinel with no rules withholds instead of ready', async () => {
    // Empty retained rules are not proof of absence: capture/parse overflow
    // marks the sheet unknown, and the common service would withhold on it.
    const sentinel: CfSheetRules = { rules: [], diagnostics: [], unresolvedSheet: { reason: 'cf-over-cap', writes: 'ALL' } }
    const st = { ...stamped(oneRule), cfGen: 0 }
    const memory = createCfMemory({})
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), SHEET, fakeSource([], { stamp: st }), sentinel, memory, { budget: frozenBudget })
    expect(out.kind).toBe('unresolved')
    // The sentinel vocabulary is preserved, like the service snapshot: the
    // sheet scope is over cap, not merely resource-limited.
    expect(out).toMatchObject({ reason: 'cf-over-cap' })
    expect(out.diagnostics[0]).toMatchObject({ feature: 'cf-sheet-scope-unresolved' })
    expect(out.states).toEqual([])
    expect(memory.currentBytes).toBe(0)
    // Populated sheets with the sentinel withhold the same way.
    const pstmemory = createCfMemory({})
    const pout = await prepareCfRemote(
      cfRequest([rect(0, 1)], st, 8 * 1024 * 1024), SHEET,
      fakeSource([{ row: 0, col: 0, value: nc(1) }, { row: 1, col: 0, value: nc(2) }], { stamp: st }),
      sentinel, pstmemory, { budget: frozenBudget },
    )
    expect(pout.kind).toBe('unresolved')
    expect(pout).toMatchObject({ reason: 'cf-over-cap' })
    expect(pstmemory.currentBytes).toBe(0)
  })

  function realProviderHost() {
    // Cache/transport integration through the real provider (not a fake
    // echoing the request stamp): requests run prepareCfRemote, and the
    // fake transport settles grants into reply charges on arrival.
    const granted = new Map<string, number>()
    const ledger = new Map<string | object, number>()
    const releasedReplies: WorkerReply<'read'>[] = []
    const requests: CommandOf<'read'>[] = []
    let n = 0
    const live = stamped(oneRule)
    const host = {
      granted, ledger, releasedReplies, requests,
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: (requestId: string) => { granted.set(requestId, 8 * 1024 * 1024); ledger.set(`grant:${requestId}`, 8 * 1024 * 1024); return 8 * 1024 * 1024 },
      releaseGrant: (requestId: string) => { granted.delete(requestId); ledger.delete(`grant:${requestId}`) },
      request: async (command: CommandOf<'read'>) => {
        requests.push(command)
        const q = command.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const source = fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: live })
        const result = await prepareCfRemote(
          { rects: q.rects, stamp: q.stamp, grantBytes: q.grantBytes }, SHEET, source, oneRule, createCfMemory({}), { budget: frozenBudget },
        )
        const reply = cfReply(result, measurePayloadBytes(result, () => true) ?? q.grantBytes, command.requestId, command.owner as CacheIdentity)
        granted.delete(command.requestId)
        ledger.delete(`grant:${command.requestId}`)
        ledger.set(reply, 1024)
        return reply
      },
      nextRequestId: () => `r${++n}`,
      release: (reply: WorkerReply<'read'>) => { releasedReplies.push(reply); ledger.delete(reply) },
      admitEntry: (key: string, bytes: number) => { ledger.set(`entry:${key}`, bytes); return () => { ledger.delete(`entry:${key}`) } },
      changed: () => {},
    }
    return { host, live, ledger, releasedReplies, requests, granted }
  }

  it('R2-4: a cold cfGen discovers and publishes the live version', async () => {
    // cfGen 0 means "unknown, adopt live": the first real reply is adopted
    // (rekeyed) and published — not dropped as stale — ending ready with a
    // live stamp the painter can keep using.
    const { host, live } = realProviderHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const cold = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1, { cfGen: 0 }))
    await flush()
    await flush()
    expect(cold.state.state).toBe('ready')
    expect(host.requests).toHaveLength(1)
    expect(cold.state.result!.stamp.cfGen).toBe(live.cfGen)
    expect(live.cfGen).not.toBe(0)
    // A live re-get hits the adopted entry without a new request.
    const again = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1, { cfGen: live.cfGen }))
    await flush()
    expect(again.state.state).toBe('ready')
    expect(host.requests).toHaveLength(1)
    cold.release()
    again.release()
    cache.invalidate(SHEET)
    await flush()
    cache.close()
    expect(host.ledger.size).toBe(0)
  })

  it('R2-4: a stale known version reloads once under the adopted stamp', async () => {
    // A provider that raced ahead does not strand the entry: one bounded
    // rediscovery adopts the live version and completes.
    const { host, live } = realProviderHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const stale = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1, { cfGen: 424242 }))
    await flush()
    await flush()
    await flush()
    expect(stale.state.state).toBe('ready')
    expect(host.requests).toHaveLength(2)
    expect(stale.state.result!.stamp.cfGen).toBe(live.cfGen)
    stale.release()
    cache.invalidate(SHEET)
    await flush()
    cache.close()
    expect(host.ledger.size).toBe(0)
  })

  it('R2-4: a persistently wrong version fails closed after one retry', async () => {
    // An adversarial worker that never matches still terminates: one
    // adoption, one reload, then failure naming the discovered version —
    // never a retry loop, and every reply released exactly once.
    const granted = new Map<string, number>()
    const ledger = new Map<string | object, number>()
    const releasedReplies: WorkerReply<'read'>[] = []
    const requests: CommandOf<'read'>[] = []
    let n = 0
    const host = {
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: (requestId: string) => { granted.set(requestId, 1 << 20); ledger.set(`grant:${requestId}`, 1 << 20); return 1 << 20 },
      releaseGrant: (requestId: string) => { granted.delete(requestId); ledger.delete(`grant:${requestId}`) },
      request: async (command: CommandOf<'read'>) => {
        requests.push(command)
        const q = command.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const result = wireResult({ stamp: { ...q.stamp, cfGen: q.stamp.cfGen + 999 } })
        const reply = cfReply(result, 1024, command.requestId, command.owner as CacheIdentity)
        granted.delete(command.requestId)
        ledger.delete(`grant:${command.requestId}`)
        ledger.set(reply, 1024)
        return reply
      },
      nextRequestId: () => `r${++n}`,
      release: (reply: WorkerReply<'read'>) => { releasedReplies.push(reply); ledger.delete(reply) },
      admitEntry: (key: string, bytes: number) => { ledger.set(`entry:${key}`, bytes); return () => { ledger.delete(`entry:${key}`) } },
      changed: () => {},
    }
    const cache = new RemoteCfCache(host, 1 << 20)
    const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1, { cfGen: 5 }))
    await flush()
    await flush()
    await flush()
    expect(held.state.state).toBe('failed')
    // First miss (cfGen 1004) reloaded once; the second (cfGen 2003) failed
    // closed and named the discovered version — then nothing more happens.
    expect(requests).toHaveLength(2)
    await flush()
    expect(requests).toHaveLength(2)
    expect(held.state.diagnostics[0]).toMatchObject({ feature: 'cf-stale-generation' })
    expect(held.state.diagnostics[0]!.message).toContain('2003')
    expect(releasedReplies).toHaveLength(2)
    held.release()
    cache.invalidate(SHEET)
    await flush()
    cache.close()
    expect(ledger.size).toBe(0)
  })

  it('R2-6: entry metadata is ledger-admitted before allocation and released with the entry', async () => {
    // Waiting entries hold a ledger charge for their key/scope/stamp copies;
    // a 256-rect scope is admitted with its variable size (not a flat 512);
    // refusal creates nothing; drops settle everything.
    const granted = new Map<string, number>()
    const ledger = new Map<string | object, number>()
    let n = 0
    const host = {
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: (requestId: string) => { granted.set(requestId, 1 << 20); return 1 << 20 },
      releaseGrant: (requestId: string) => { granted.delete(requestId) },
      request: async (): Promise<WorkerReply<'read'> | symbol> => new Promise(() => {}),
      nextRequestId: () => `r${++n}`,
      release: () => {},
      admitEntry: (key: string, bytes: number) => { ledger.set(`entry:${key}`, bytes); return () => { ledger.delete(`entry:${key}`) } },
      changed: () => {},
    }
    const cache = new RemoteCfCache(host, 1 << 20)
    const waiting = cache.get(SHEET, [rect(0, 0)], cfStamp(1))
    // The entry is admitted at creation (waiting instant); the load starts
    // right after on a complete sheet while its request never resolves here.
    expect(['waiting', 'loading']).toContain(waiting.state)
    // The entry's metadata is on the ledger before any request completes.
    expect(ledger.size).toBeGreaterThan(0)
    const wide = cache.get(SHEET, Array.from({ length: 256 }, (_, i) => rect(i, i)), cfStamp(1))
    // Loading already (its request never resolves here); the admission
    // below is what this step proves, not the load state.
    expect(['waiting', 'loading']).toContain(wide.state)
    const { entryMetaBytes: metaOf } = await import('../src/worker/viewerRemoteCf')
    const narrowMeta = metaOf(SHEET, [rect(0, 0)], cfStamp(1))
    const wideMeta = metaOf(SHEET, Array.from({ length: 256 }, (_, i) => rect(i, i)), cfStamp(1))
    expect(wideMeta).toBeGreaterThan(narrowMeta)
    expect([...ledger.values()].reduce((sum, v) => sum + (typeof v === 'number' ? v : 0), 0)).toBe(narrowMeta + wideMeta)
    cache.invalidate(SHEET)
    await flush()
    expect(ledger.size).toBe(0)
    cache.close()
  })

  it('R2-6: admission refusal creates no entry and sends no request', async () => {
    const requests: CommandOf<'read'>[] = []
    let n = 0
    const host = {
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: () => 1 << 20,
      releaseGrant: () => {},
      request: async (command: CommandOf<'read'>) => { requests.push(command); throw new Error('must not send') },
      nextRequestId: () => `r${++n}`,
      release: () => {},
      admitEntry: () => null,
      changed: () => {},
    }
    const cache = new RemoteCfCache(host, 1 << 20)
    const failed = cache.get(SHEET, [rect(0, 0)], cfStamp(1))
    expect(failed.state).toBe('failed')
    expect(failed.diagnostics[0]).toMatchObject({ feature: 'cf-cache-budget' })
    expect(requests).toHaveLength(0)
    expect(cache.retainedEntries).toBe(0)
    expect(cache.retainedBytes).toBe(0)
    cache.close()
  })

  it('R2-2d: the session releases the transferred outgoing lease after drain', async () => {
    // The provider hands the session a live outgoing lease; the round trip
    // ends with it released exactly once — no window, no leak.
    const { ViewerClient } = await import('../src/worker/viewerClient')
    const { connectInProcess, ViewerKernel } = await import('../src/worker/viewerKernel')
    const { createObjectWorkbookSource } = await import('../src/worker/viewerObjectSource')
    const { initialIdentity } = await import('../src/worker/viewerProtocol')
    const { workbookOf } = await import('./helpers/xlsx-find')
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 5 }] }])
    const base = createObjectWorkbookSource(doc, 'doc')
    const releases: Array<() => void> = []
    let released = 0
    const wrapped = {
      ...base,
      store: (id: string) => ({
        ...base.store(id),
        readConditionalFormatting: async (request: CfReadRequest, memory: unknown, options?: CfReadOptions) => {
          void memory
          options?.retainOutgoing?.(() => { released++; releases.push(() => {}) })
          return {
            sheetId: id, stamp: { ...request.stamp }, kind: 'ready' as const,
            states: [], definitions: [], definitionsVersion: 0, diagnostics: [],
          }
        },
      }),
    }
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => wrapped }, { yield: async () => {} })
    const port = connectInProcess(kernel)
    const client = new ViewerClient(port)
    const open = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
    if (open.kind !== 'complete') throw new Error('open failed')
    const owner = initialIdentity('doc', 's')
    const out = await client.send({
      kind: 'read', requestId: client.nextRequestId(), owner,
      query: { op: 'conditionalFormatting' as const, rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], stamp: { owner: 'owner', sheetId: 's', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 1, clockGen: 0 }, grantBytes: 1 << 20 },
    })
    expect(out.kind).toBe('complete')
    // Posting resolves the client send before the kernel's drain releases
    // the outgoing leases; a flush lets the drain run first.
    await flush()
    expect(releases).toHaveLength(1)
    expect(released).toBe(1)
  })
})

describe('CF fix round 3 (review-t7-r3)', () => {
  const identity = (sheetId = SHEET, valueGen = 1): CacheIdentity => ({ docId: 'doc', sheetId, valueGen, fontGen: 0, viewGen: 0 })
  const cfStamp = (gen = 1, extra: Partial<CfReadStamp> = {}): CfReadStamp => ({
    owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
  })
  const flush = () => new Promise(r => setTimeout(r, 0))
  const wireResult = (over: Partial<CfReadResult> = {}): CfReadResult => ({
    sheetId: SHEET, stamp: cfStamp(1), kind: 'ready', states: [{ row: 0, col: 0, properties: { 'fill.pattern': 'solid' } }],
    definitions: [{ id: 'rule-0', version: 7, properties: { 'fill.pattern': 'solid' }, writes: ['fill'] }],
    definitionsVersion: 7, diagnostics: [], ...over,
  })
  const cfReply = (result: CfReadResult, bytes: number, requestId = 'r1', owner: CacheIdentity = identity()): WorkerReply<'read'> => ({
    commandKind: 'read', requestId, owner, diagnostics: [], replyBytes: bytes, kind: 'complete',
    result: { kind: 'read', result: { op: 'conditionalFormatting', result } },
  } as WorkerReply<'read'>)
  /** Cache host with a ledger-backed main-ledger admission, mirroring the
   * production wiring (layout pool, per-owner release). */
  function grantHost(budgetBytes = 1 << 20) {
    const granted = new Map<string, number>()
    const ledger = new Map<string | object, number>()
    const releasedReplies: WorkerReply<'read'>[] = []
    const changedCalls: { sheetId: string; rects: readonly SourceRect[]; reason?: string }[] = []
    let n = 0
    return {
      granted, ledger, releasedReplies, changedCalls,
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: (requestId: string) => { granted.set(requestId, budgetBytes); ledger.set(`grant:${requestId}`, budgetBytes); return budgetBytes },
      releaseGrant: (requestId: string) => { granted.delete(requestId); ledger.delete(`grant:${requestId}`) },
      request: async (command: CommandOf<'read'>) => {
        const q = command.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const result = wireResult({ stamp: { ...q.stamp } })
        const reply = cfReply(result, 1024, command.requestId, command.owner as CacheIdentity)
        granted.delete(command.requestId)
        ledger.delete(`grant:${command.requestId}`)
        ledger.set(reply, 1024)
        return reply
      },
      nextRequestId: () => `r${++n}`,
      release: (reply: WorkerReply<'read'>) => { releasedReplies.push(reply); ledger.delete(reply) },
      admitEntry: (key: string, bytes: number) => { ledger.set(`entry:${key}`, bytes); return () => { ledger.delete(`entry:${key}`) } },
      changed: (sheetId: string, rects: readonly SourceRect[], reason?: 'updated' | 'evicted' | 'status') => { changedCalls.push({ sheetId, rects, reason }) },
    }
  }
  const longRules = (code: string): CfSheetRules => ({ diagnostics: [], rules: [rule(0, { 'numFmt.code': code }, ['numFmt'])] })
  const fakeSheetWith = (rules: CfSheetRules | undefined) => ({
    sheetId: SHEET, name: SHEET, state: 'complete' as const, importedThroughRow: 1048575,
    metadata: {}, tables: undefined, arrayCells: [], drawingAnchors: { records: [], unresolved: 0 },
    mergeRects: [], invalidMerges: 0, used: { r0: 0, r1: 0, c0: 0, c1: 0 }, conditionalFormatting: rules,
  })
  /** Runs with JSON.stringify hooked: every normalized-rules serialization
   * records the account bytes held at that moment (the probe pattern), so
   * tests prove admission precedes serialization and count how often the
   * version is computed. Always restores the original. */
  async function withNormalizeProbe<T>(memory: { readonly currentBytes: number }, run: (calls: { held: number; serializedBytes: number }[]) => Promise<T>): Promise<{ result: T; calls: { held: number; serializedBytes: number }[] }> {
    const calls: { held: number; serializedBytes: number }[] = []
    const stringify = JSON.stringify
    JSON.stringify = (function (this: unknown, value: unknown, ...args: unknown[]) {
      const out = (stringify as (...a: unknown[]) => string).call(this, value, ...args)
      if (Array.isArray(value) && (value as unknown[]).length > 0 && ((value as { dxf?: unknown }[])[0] as { dxf?: unknown })?.dxf !== undefined) {
        calls.push({ held: memory.currentBytes, serializedBytes: out.length * 2 })
      }
      return out
    }) as typeof JSON.stringify
    try {
      const result = await run(calls)
      return { result, calls }
    } finally {
      JSON.stringify = stringify
    }
  }

  it('R3-1: the adapter admits the exact size before its own normalization', async () => {
    // The adapter's hash walk must be exactly sized like the provider's: an
    // 8 KiB format code is serialized exactly once, with the hold above the
    // serialized size — and a tight account refuses before serializing.
    const { PagedXlsxSheetAdapter } = await import('../src/xlsx/view/pagedSource')
    const { createCfMemory: makeMemory } = await import('../src/xlsx/cf/limits')
    const rules = longRules('X'.repeat(8192))
    const st = stamped(rules)
    const sourceStamp = { docId: 'doc', sheetId: SHEET, valueGen: 1 }
    const adapter = new PagedXlsxSheetAdapter(sourceStamp, {} as never, fakeSheetWith(rules) as never, undefined, {}, false)
    try {
      const memory = makeMemory({})
      const { calls } = await withNormalizeProbe(memory, () =>
        adapter.readConditionalFormatting(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), memory, { budget: frozenBudget }))
      // Exactly one normalization (adapter computes the version once and the
      // provider reuses the admitted result), held above the serialized size.
      expect(calls).toHaveLength(1)
      expect(calls[0]!.serializedBytes).toBeGreaterThan(8192 * 2)
      expect(calls[0]!.held).toBeGreaterThanOrEqual(calls[0]!.serializedBytes)
      expect(memory.currentBytes).toBe(0)
      // A tight account refuses in exact sizing, before any serialization.
      const tight = makeMemory({ limitBytes: 8000 })
      const refused = await withNormalizeProbe(tight, () =>
        adapter.readConditionalFormatting(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), tight, { budget: frozenBudget }))
      expect(refused.calls).toHaveLength(0)
      expect(refused.result.kind).toBe('unresolved')
      expect(tight.currentBytes).toBe(0)
    } finally {
      await adapter.release().catch(() => {})
    }
  })

  it('R3-2: the session walk is admitted on the shared account', async () => {
    // The outgoing measure's traversal workspace must be admitted on the
    // shared account, not merely capped: the walk reserves and grows there,
    // and a full pool refuses the walk itself instead of allocating
    // visited-set/stack storage first and failing only at the charge.
    const { DocumentSession } = await import('../src/worker/viewerSession')
    const session = new DocumentSession('d', { sheets: [{ sheetId: SHEET }], store: () => ({}) } as never)
    try {
      const owner = (session as unknown as { identity(sheetId: string): { docId: string; sheetId: string; valueGen: number; fontGen: number; viewGen: number } }).identity(SHEET)
      const query = {
        op: 'conditionalFormatting' as const,
        rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
        stamp: { owner: 'o', sheetId: SHEET, workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 0 },
        grantBytes: 1 << 20,
      }
      const account = await (session as unknown as { cfAccount(): Promise<import('../src/xlsx/cf/limits').CfMemory> }).cfAccount()
      let reserves = 0
      const origReserve = account.reserve.bind(account)
      account.reserve = ((bytes: number) => { reserves++; return origReserve(bytes) }) as typeof account.reserve
      try {
        // Roomy pool: the walk admits its workspace on the shared account.
        const roomy = await session.read(owner, query, new AbortController().signal)
        expect(roomy).toMatchObject({ op: 'conditionalFormatting', result: { kind: 'unresolved', reason: 'cf-unavailable' } })
        expect(reserves).toBeGreaterThan(0)
        // Full pool: the walk itself is refused, failing closed with no build.
        const hold = session.memory.charge()
        expect(hold.set(session.memory.limitBytes)).toBe(true)
        try {
          const out = await session.read(owner, query, new AbortController().signal)
          expect(out).toMatchObject({ op: 'conditionalFormatting', result: { kind: 'unresolved', reason: 'cf-resource-limit' } })
        } finally {
          hold.release()
        }
      } finally {
        account.reserve = origReserve
      }
    } finally {
      await session.release().catch(() => {})
    }
  })

  it('R3-3: the shared-definition entry owns its bytes until explicit release', async () => {
    // The table carries its own retained lease: after the outgoing result
    // lease drains, the account still holds the definitions charge, and only
    // an explicit release settles it.
    const { sharedCfDefinitionStats, releaseSharedCfDefinitions } =
      await import('../src/xlsx/view/pagedSource')
    const rules: CfSheetRules = { diagnostics: [], rules: [rule(3, { 'fill.pattern': 'solid' }, ['fill'])] }
    const st = stamped(rules)
    const mkSource = () => fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: st })
    const memory = createCfMemory({})
    const releases: Array<() => void> = []
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), SHEET, mkSource(), rules, memory,
      { budget: frozenBudget, retainOutgoing: release => { releases.push(release) } })
    expect(out.kind).toBe('ready')
    expect(releases).toHaveLength(1)
    const held = memory.currentBytes
    expect(held).toBeGreaterThan(0)
    expect(sharedCfDefinitionStats(rules)).toMatchObject({ charged: true })
    // Draining the outgoing lease keeps the table charge: bytes drop but
    // stay above zero, and the entry remains cached.
    releases[0]!()
    const retained = memory.currentBytes
    expect(retained).toBeGreaterThan(0)
    expect(retained).toBeLessThan(held)
    expect(sharedCfDefinitionStats(rules)).toBeDefined()
    releaseSharedCfDefinitions(rules)
    expect(memory.currentBytes).toBe(0)
    expect(sharedCfDefinitionStats(rules)).toBeUndefined()
  })

  it('R3-3: two viewport consumers lease one shared definitions array', async () => {
    // Version-keyed shared ownership with independent state leases: two
    // scopes read the identical frozen array; dropping the first keeps the
    // second live; dropping the last releases the version and settles all.
    const host = grantHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const a = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    const b = cache.acquire(SHEET, [rect(5, 5)], cfStamp(1))
    await flush()
    await flush()
    expect(a.state.state).toBe('ready')
    expect(b.state.state).toBe('ready')
    expect(a.state.result!.definitions).toBe(b.state.result!.definitions)
    expect(Object.isFrozen(a.state.result!.definitions)).toBe(true)
    expect(cache.sharedDefinitionVersions).toBe(1)
    // Each publish settles its reply at once; drops only release components.
    expect(host.releasedReplies).toHaveLength(2)
    // Drop the first scope: the shared array survives for the second.
    a.release()
    cache.invalidate(SHEET, [rect(0, 0)])
    await flush()
    expect(cache.sharedDefinitionVersions).toBe(1)
    expect(cache.get(SHEET, [rect(5, 5)], cfStamp(1)).state).toBe('ready')
    expect(host.releasedReplies).toHaveLength(2)
    // Drop the last scope: the version releases and everything settles.
    b.release()
    cache.invalidate(SHEET)
    await flush()
    expect(cache.sharedDefinitionVersions).toBe(0)
    expect(cache.retainedBytes).toBe(0)
    expect(host.releasedReplies).toHaveLength(2)
    expect(host.ledger.size).toBe(0)
    cache.close()
  })

  it('R3-4: a ready cache hit needs no new admission', async () => {
    // Lookup happens before admission: with the ledger refusing further
    // reservations, a warm hit still serves the ready entry without a new
    // admission or request, while a new scope fails closed sending nothing.
    const requests: CommandOf<'read'>[] = []
    const ledger = new Map<string | object, number>()
    let n = 0
    let permit = true
    let admissions = 0
    const host = {
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: () => 1 << 20,
      releaseGrant: () => {},
      request: async (command: CommandOf<'read'>) => {
        requests.push(command)
        const q = command.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const result = wireResult({ stamp: { ...q.stamp } })
        const reply = cfReply(result, 1024, command.requestId, command.owner as CacheIdentity)
        ledger.set(reply, 1024)
        return reply
      },
      nextRequestId: () => `r${++n}`,
      release: (reply: WorkerReply<'read'>) => { ledger.delete(reply) },
      admitEntry: (key: string, bytes: number) => {
        if (!permit) return null
        admissions++
        ledger.set(`entry:${key}`, bytes)
        return () => { ledger.delete(`entry:${key}`) }
      },
      changed: () => {},
    }
    const cache = new RemoteCfCache(host, 1 << 20)
    const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    expect(held.state.state).toBe('ready')
    expect(requests).toHaveLength(1)
    const admitted = admissions
    expect(admitted).toBeGreaterThan(0)
    // Freeze the ledger: the warm hit must not admit, request, or churn.
    permit = false
    const warm = cache.get(SHEET, [rect(0, 0)], cfStamp(1))
    expect(warm.state).toBe('ready')
    expect(warm.result).toBe(held.state.result)
    expect(requests).toHaveLength(1)
    expect(admissions).toBe(admitted)
    // A new scope under pressure fails closed without sending.
    const cold = cache.get(SHEET, [rect(9, 9)], cfStamp(1))
    expect(cold.state).toBe('failed')
    expect(requests).toHaveLength(1)
    permit = true
    held.release()
    cache.invalidate(SHEET)
    await flush()
    cache.close()
    expect(ledger.size).toBe(0)
  })
})

describe('CF fix round 4 (review-t7-r4)', () => {
  const identity = (sheetId = SHEET, valueGen = 1): CacheIdentity => ({ docId: 'doc', sheetId, valueGen, fontGen: 0, viewGen: 0 })
  const cfStamp = (gen = 1, extra: Partial<CfReadStamp> = {}): CfReadStamp => ({
    owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
  })
  const flush = () => new Promise(r => setTimeout(r, 0))
  const cfReply = (result: CfReadResult, bytes: number, requestId = 'r1', owner: CacheIdentity = identity()): WorkerReply<'read'> => ({
    commandKind: 'read', requestId, owner, diagnostics: [], replyBytes: bytes, kind: 'complete',
    result: { kind: 'read', result: { op: 'conditionalFormatting', result } },
  } as WorkerReply<'read'>)

  it('R4-1: a refused definition reservation publishes no share entry', async () => {
    // Probe pattern: refuse exactly the definition-sized reservation on a
    // real account. The preparation still completes (the outgoing lease
    // owns the result), but no share entry is published and nothing stays
    // charged after the outgoing lease drains.
    const { sharedCfDefinitionStats, releaseSharedCfDefinitions } =
      await import('../src/xlsx/view/pagedSource')
    const dxf = { 'numFmt.code': 'Y'.repeat(9000) }
    // Learn the exact shared-definitions size with a dry run: without an
    // outbox the only retained charge afterwards is the table lease.
    const probeRules: CfSheetRules = { diagnostics: [], rules: [rule(0, dxf, ['numFmt'])] }
    const probeSt = stamped(probeRules)
    const probeMem = createCfMemory({})
    const probe = await prepareCfRemote(cfRequest([rect(0, 0)], probeSt, 8 * 1024 * 1024), SHEET,
      fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: probeSt }), probeRules, probeMem, { budget: frozenBudget })
    expect(probe.kind).toBe('ready')
    const defsBytes = probeMem.currentBytes
    expect(defsBytes).toBeGreaterThan(0)
    releaseSharedCfDefinitions(probeRules)
    expect(probeMem.currentBytes).toBe(0)
    const rules: CfSheetRules = { diagnostics: [], rules: [rule(0, { ...dxf }, ['numFmt'])] }
    const st = stamped(rules)
    const mkSource = () => fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: st })
    const memory = createCfMemory({})
    let refusals = 0
    const refusing = {
      get currentBytes() { return memory.currentBytes },
      get constructionBytes() { return memory.constructionBytes },
      get retainedBytes() { return memory.retainedBytes },
      get peakBytes() { return memory.peakBytes },
      get closed() { return memory.closed },
      close() { memory.close() },
      reserve: (bytes: number) => {
        if (bytes === defsBytes) { refusals++; return { kind: 'unresolved' as const, reason: 'cf-budget-exceeded' } }
        return memory.reserve(bytes)
      },
    } as unknown as import('../src/xlsx/cf/limits').CfMemory
    const releases: Array<() => void> = []
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), SHEET, mkSource(), rules, refusing,
      { budget: frozenBudget, retainOutgoing: release => { releases.push(release) } })
    expect(out.kind).toBe('ready')
    expect(releases).toHaveLength(1)
    // Exactly the definition reservation was refused — nothing else is
    // admitted at exactly that size on this account.
    expect(refusals).toBe(1)
    expect(sharedCfDefinitionStats(rules)).toBeUndefined()
    releases[0]!()
    expect(memory.currentBytes).toBe(0)
    releaseSharedCfDefinitions(rules)
  })

  it('R4-1: shared admission fails closed at each stage without holding', async () => {
    // Unit coverage for the three independent failure paths of the shared
    // copy admission: refused reservation, unmeasurable walk, failed adopt.
    const { admitSharedCopy, createCfMemory } = await import('../src/xlsx/cf/limits')
    const value = { id: 'r', properties: { 'fill.pattern': 'solid' }, writes: ['fill'] }
    // Success path admits and releases cleanly.
    const roomy = createCfMemory({})
    const lease = admitSharedCopy(roomy, value)
    expect(lease).toBeDefined()
    expect(roomy.currentBytes).toBeGreaterThan(0)
    lease!.release()
    expect(roomy.currentBytes).toBe(0)
    // Refused reservation admits nothing.
    const denied = createCfMemory({ limitBytes: 0 })
    expect(admitSharedCopy(denied, value)).toBeUndefined()
    expect(denied.currentBytes).toBe(0)
    expect(denied.peakBytes).toBe(0)
    // Unmeasurable walk (workspace cannot grow) admits nothing.
    const noGrow = {
      get currentBytes() { return 0 }, get constructionBytes() { return 0 }, get retainedBytes() { return 0 },
      get peakBytes() { return 0 }, get closed() { return false }, close() {},
      reserve: () => ({ kind: 'reserved' as const, bytes: 0, grow: () => false, adopt: (): never => { throw new Error('unreachable') }, release: () => {} }),
    } as unknown as import('../src/xlsx/cf/limits').CfMemory
    expect(admitSharedCopy(noGrow, value)).toBeUndefined()
    // Failed adoption releases the reservation it admitted.
    let releasedHold = false
    const noAdopt = {
      get currentBytes() { return 0 }, get constructionBytes() { return 0 }, get retainedBytes() { return 0 },
      get peakBytes() { return 0 }, get closed() { return false }, close() {},
      reserve: () => ({
        kind: 'reserved' as const, bytes: 0, grow: () => true,
        adopt: (): never => { throw new Error('adoption refused') },
        release: () => { releasedHold = true },
      }),
    } as unknown as import('../src/xlsx/cf/limits').CfMemory
    expect(admitSharedCopy(noAdopt, value)).toBeUndefined()
    expect(releasedHold).toBe(true)
  })

  it('R4-2: concurrent cold publications share one owner and leak nothing', async () => {
    // Two preparations racing past the same table miss must not overwrite
    // each other's definition lease: the recheck publishes once, the second
    // reuses the winner, and releasing the share settles everything.
    const rules: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'fill.pattern': 'solid' }, ['fill'])] }
    const st = stamped(rules)
    const memory = createCfMemory({})
    const { sharedCfDefinitionStats, releaseSharedCfDefinitions } =
      await import('../src/xlsx/view/pagedSource')
    let starts = 0
    let unlock!: () => void
    const gate = new Promise<void>(resolve => { unlock = resolve })
    const base = fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: st })
    const src = {
      ...base,
      pages: (async function* (...args: [string, readonly never[], import('../src/xlsx/cf/limits').CfWorkBudget, AbortSignal]) {
        starts++
        await gate
        yield* base.pages(...args)
      }) as CfValueSource['pages'],
    }
    const ends: Array<() => void> = []
    const call = () => prepareCfRemote(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), SHEET, src, rules, memory,
      { budget: frozenBudget, retainOutgoing: release => { ends.push(release) } })
    const p1 = call()
    const p2 = call()
    while (starts < 2) await new Promise(r => setTimeout(r, 1))
    unlock()
    const [a, b] = await Promise.all([p1, p2])
    expect(a.kind).toBe('ready')
    expect(b.kind).toBe('ready')
    if (a.kind !== 'ready' || b.kind !== 'ready') throw new Error('expected ready')
    // One shared entry for both racers (the second reused the winner via the
    // atomic recheck instead of committing its own lease).
    expect(a.definitions).toBe(b.definitions)
    expect(sharedCfDefinitionStats(rules)).toMatchObject({ hits: 1 })
    for (const end of ends) end()
    // Exactly one definition lease survives the drains.
    expect(memory.currentBytes).toBeGreaterThan(0)
    releaseSharedCfDefinitions(rules)
    expect(memory.currentBytes).toBe(0)
    expect(sharedCfDefinitionStats(rules)).toBeUndefined()
  })

  it('R4-3: the receiver walk is admitted before it traverses', async () => {
    // Admission precedes traversal: with the walk-scratch reservation
    // refused, the traced value is never read, yet the entry still
    // completes ready with its own definitions (fail-soft sharing).
    const events: string[] = []
    const props: Record<string, unknown> = {}
    Object.defineProperty(props, 'fill.fg', {
      enumerable: true, configurable: true,
      get() { events.push('walk read'); return 'green' },
    })
    let n = 0
    const host = {
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: () => 1 << 20,
      releaseGrant: () => {},
      request: async (command: CommandOf<'read'>) => {
        const q = command.query as { stamp: CfReadStamp }
        const result: CfReadResult = {
          sheetId: SHEET, stamp: { ...q.stamp }, kind: 'ready',
          states: [{ row: 0, col: 0, properties: {} }],
          definitions: [{ id: 'r', version: 99, properties: props as Record<string, string | number | boolean | null>, writes: ['fill'] }],
          definitionsVersion: 99, diagnostics: [],
        }
        const reply = cfReply(result, 2048, command.requestId, command.owner as CacheIdentity)
        return reply
      },
      nextRequestId: () => `r${++n}`,
      release: () => {},
      admitEntry: (key: string) => {
        if (key.startsWith('cf-defs-walk:')) { events.push('walk scratch refused'); return null }
        return () => {}
      },
      changed: () => {},
    }
    const cache = new RemoteCfCache(host, 1 << 20)
    const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    await flush()
    expect(held.state.state).toBe('ready')
    // No traversal happened: the admission refusal came first.
    expect(events).toEqual(['walk scratch refused'])
    expect(cache.sharedDefinitionVersions).toBe(0)
    held.release()
    cache.invalidate(SHEET)
    await flush()
    cache.close()
  })

  it('R4-3: sharing admits walk scratch transiently and the registry durably', async () => {    // A full main ledger still serves ready entries: the walk scratch is
    // admitted transiently (released with the walk) while the version
    // charge is retained durably; full release settles everything.
    const ledger = new Map<string | object, number>()
    let current = 0
    let walkAdmitted = 0
    const cap = 1 << 20
    let n = 0
    const props: Record<string, unknown> = {}
    for (let i = 0; i < 64; i++) props[`k${i}`] = `v${i}`
    const host = {
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: () => 1 << 20,
      releaseGrant: () => {},
      request: async (command: CommandOf<'read'>) => {
        const q = command.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const result: CfReadResult = {
          sheetId: SHEET, stamp: { ...q.stamp }, kind: 'ready',
          states: [{ row: 0, col: 0, properties: {} }],
          definitions: [{ id: 'r', version: 77, properties: props as Record<string, string | number | boolean | null>, writes: ['fill'] }],
          definitionsVersion: 77, diagnostics: [],
        }
        const reply = cfReply(result, 65536, command.requestId, command.owner as CacheIdentity)
        ledger.set(reply, 65536)
        current += 65536
        return reply
      },
      nextRequestId: () => `r${++n}`,
      release: (reply: WorkerReply<'read'>) => { ledger.delete(reply); current -= 65536 },
      admitEntry: (key: string, bytes: number) => {
        if (current + bytes > cap) return null
        if (key.startsWith('cf-defs-walk:')) walkAdmitted += bytes
        ledger.set(key, bytes)
        current += bytes
        return () => { ledger.delete(key); current -= bytes }
      },
      changed: () => {},
    }
    const cache = new RemoteCfCache(host, 1 << 20)
    const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    await flush()
    expect(held.state.state).toBe('ready')
    expect(cache.sharedDefinitionVersions).toBe(1)
    // The walk scratch was admitted (then released with the walk) while the
    // version charge stayed retained — not the other way around.
    expect(walkAdmitted).toBeGreaterThan(0)
    expect([...ledger.keys()].some(k => typeof k === 'string' && k.startsWith('cf-defs:77'))).toBe(true)
    expect([...ledger.keys()].some(k => typeof k === 'string' && k.startsWith('cf-defs-walk:'))).toBe(false)
    held.release()
    cache.invalidate(SHEET)
    await flush()
    cache.close()
    expect(ledger.size).toBe(0)
    expect(current).toBe(0)
    expect(cache.sharedDefinitionVersions).toBe(0)
  })
})

describe('CF fix round 5 (review-t7-r5)', () => {
  const identity = (sheetId = SHEET, valueGen = 1): CacheIdentity => ({ docId: 'doc', sheetId, valueGen, fontGen: 0, viewGen: 0 })
  const cfStamp = (gen = 1, extra: Partial<CfReadStamp> = {}): CfReadStamp => ({
    owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
  })
  const wireResult = (over: Partial<CfReadResult> = {}): CfReadResult => ({
    sheetId: SHEET, stamp: cfStamp(1), kind: 'ready', states: [{ row: 0, col: 0, properties: { 'fill.pattern': 'solid' } }],
    definitions: [{ id: 'rule-0', version: 7, properties: { 'fill.pattern': 'solid' }, writes: ['fill'] }],
    definitionsVersion: 7, diagnostics: [], ...over,
  })
  const cfReply = (result: CfReadResult, bytes: number, requestId = 'r1', owner: CacheIdentity = identity()): WorkerReply<'read'> => ({
    commandKind: 'read', requestId, owner, diagnostics: [], replyBytes: bytes, kind: 'complete',
    result: { kind: 'read', result: { op: 'conditionalFormatting', result } },
  } as WorkerReply<'read'>)
  const flush = () => new Promise(r => setTimeout(r, 0))
  /** Cache host with a ledger-backed main-ledger admission, mirroring the
   * production wiring (layout pool, per-owner release). */
  function grantHost(budgetBytes = 1 << 20) {
    const granted = new Map<string, number>()
    const ledger = new Map<string | object, number>()
    const releasedReplies: WorkerReply<'read'>[] = []
    let n = 0
    return {
      granted, ledger, releasedReplies,
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: (requestId: string) => { granted.set(requestId, budgetBytes); ledger.set(`grant:${requestId}`, budgetBytes); return budgetBytes },
      releaseGrant: (requestId: string) => { granted.delete(requestId); ledger.delete(`grant:${requestId}`) },
      request: async (command: CommandOf<'read'>) => {
        const q = command.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const result = wireResult({ stamp: { ...q.stamp } })
        const reply = cfReply(result, 1024, command.requestId, command.owner as CacheIdentity)
        granted.delete(command.requestId)
        ledger.delete(`grant:${command.requestId}`)
        ledger.set(reply, 1024)
        return reply
      },
      nextRequestId: () => `r${++n}`,
      release: (reply: WorkerReply<'read'>) => { releasedReplies.push(reply); ledger.delete(reply) },
      admitEntry: (key: string, bytes: number) => { ledger.set(`entry:${key}`, bytes); return () => { ledger.delete(`entry:${key}`) } },
      changed: () => {},
    }
  }
  it('R5-1: a failed outgoing adoption publishes no share entry', async () => {
    // Commit ordering: the table entry is created only after the outgoing
    // adopt succeeds. Refusing exactly the outgoing adoption withholds the
    // result with nothing retained and no share entry — while an unconditional
    // commit would have leaked a charged entry on this failed request.
    const { sharedCfDefinitionStats, releaseSharedCfDefinitions } =
      await import('../src/xlsx/view/pagedSource')
    const rules: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'fill.pattern': 'solid' }, ['fill'])] }
    const st = stamped(rules)
    const mkSource = () => fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: st })
    const backing = createCfMemory({})
    let adoptedResult = 0
    const refusing = {
      get currentBytes() { return backing.currentBytes },
      get constructionBytes() { return backing.constructionBytes },
      get retainedBytes() { return backing.retainedBytes },
      get peakBytes() { return backing.peakBytes },
      get closed() { return backing.closed },
      close() { backing.close() },
      reserve: (bytes: number) => {
        const hold = backing.reserve(bytes)
        if (hold.kind !== 'reserved') return hold
        const wrapped = {
          kind: 'reserved' as const,
          get bytes() { return hold.bytes },
          grow: (next: number) => hold.grow(next),
          adopt: (<T>(value: T, retainedBytes?: number) => {
            if (value !== null && typeof value === 'object' && 'states' in value && 'definitions' in value) {
              adoptedResult++
              throw new Error('outgoing adoption refused')
            }
            return hold.adopt(value, retainedBytes)
          }) as import('../src/xlsx/cf/limits').CfReservation['adopt'],
          release: () => hold.release(),
        }
        return wrapped
      },
    } as unknown as import('../src/xlsx/cf/limits').CfMemory
    const out = await prepareCfRemote(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), SHEET, mkSource(), rules, refusing, { budget: frozenBudget })
    expect(adoptedResult).toBe(1)
    expect(out.kind).toBe('unresolved')
    expect(out).toMatchObject({ reason: 'cf-resource-limit' })
    expect(sharedCfDefinitionStats(rules)).toBeUndefined()
    expect(backing.currentBytes).toBe(0)
    releaseSharedCfDefinitions(rules)
  })

  it('R5-1: a tight budget withholds cleanly and retries stay clean', async () => {
    // An account far below the working set withholds before any share can
    // exist — repeatedly, with nothing retained and no table entry.
    const { sharedCfDefinitionStats, releaseSharedCfDefinitions } =
      await import('../src/xlsx/view/pagedSource')
    const rules: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'fill.fg': 'FF00FF00' }, ['fill'])] }
    const st = stamped(rules)
    const mkSource = () => fakeSource([{ row: 0, col: 0, value: nc(1) }], { stamp: st })
    const memory = createCfMemory({ limitBytes: 8000 })
    for (let i = 0; i < 3; i++) {
      const out = await prepareCfRemote(cfRequest([rect(0, 0)], st, 100000), SHEET, mkSource(), rules, memory, { budget: frozenBudget })
      expect(out.kind).toBe('unresolved')
      expect(sharedCfDefinitionStats(rules)).toBeUndefined()
      expect(memory.currentBytes).toBe(0)
    }
    releaseSharedCfDefinitions(rules)
  })

  it('R5-1: a failed handoff rolls back its newly committed share', async () => {
    // Refusal after the shared lease is reserved: when the outgoing handoff
    // throws, the adopted lease unwinds and exactly the entry this call
    // committed is released — no newly retained cache owner on a failed
    // request, and a rechecked winner is never touched.
    const { sharedCfDefinitionStats, releaseSharedCfDefinitions } =
      await import('../src/xlsx/view/pagedSource')
    const rules: CfSheetRules = { diagnostics: [], rules: [rule(0, { 'fill.pattern': 'solid' }, ['fill'])] }
    const st = stamped(rules)
    const mkSource = () => fakeSource([{ row: 0, col: 0, value: nc(5) }], { stamp: st })
    const memory = createCfMemory({})
    await expect(prepareCfRemote(cfRequest([rect(0, 0)], st, 8 * 1024 * 1024), SHEET, mkSource(), rules, memory, {
      budget: frozenBudget,
      retainOutgoing: () => { throw new Error('drain failed') },
    })).rejects.toThrow('drain failed')
    expect(sharedCfDefinitionStats(rules)).toBeUndefined()
    expect(memory.currentBytes).toBe(0)
    releaseSharedCfDefinitions(rules)
  })

  it('R5-note: shared definitions are charged once across viewport consumers', async () => {
    // Single-charge accounting: two scopes with the same definitions version
    // hold one registry charge plus one states charge each — the second scope
    // adds no new definitions charge. Dropping both scopes settles everything.
    const host = grantHost()
    const cache = new RemoteCfCache(host, 1 << 20)
    const probe = wireResult()
    const statesBytes = measurePayloadBytes({ ...probe, definitions: [] }, () => true)!
    const meta = entryMetaBytes(SHEET, [rect(0, 0)], cfStamp(1))
    const defsBytes = measurePayloadBytes(probe.definitions, () => true)!
    const a = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    const b = cache.acquire(SHEET, [rect(5, 5)], cfStamp(1))
    await flush()
    await flush()
    expect(a.state.state).toBe('ready')
    expect(b.state.state).toBe('ready')
    expect(a.state.result!.definitions).toBe(b.state.result!.definitions)
    // Two states charges, one shared definitions charge, two metadata charges.
    expect(cache.retainedBytes).toBe(2 * (meta + statesBytes) + defsBytes)
    expect(cache.sharedDefinitionVersions).toBe(1)
    a.release()
    cache.invalidate(SHEET, [rect(0, 0)])
    await flush()
    expect(cache.sharedDefinitionVersions).toBe(1)
    expect(cache.retainedBytes).toBe(meta + statesBytes + defsBytes)
    b.release()
    cache.invalidate(SHEET)
    await flush()
    expect(cache.sharedDefinitionVersions).toBe(0)
    expect(cache.retainedBytes).toBe(0)
    expect(host.ledger.size).toBe(0)
    cache.close()
  })

  it('R5-note: unmeasurable components fall back to the full reply charge', async () => {
    // If the split walk cannot be admitted, the entry keeps the transport's
    // full reply charge instead of publishing partially owned state: still
    // fail-closed, never double-counted.
    const granted = new Map<string, number>()
    const ledger = new Map<string | object, number>()
    const releasedReplies: WorkerReply<'read'>[] = []
    const requests: CommandOf<'read'>[] = []
    let n = 0
    const host = {
      identity, loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: (requestId: string) => { granted.set(requestId, 1 << 20); return 1 << 20 },
      releaseGrant: (requestId: string) => { granted.delete(requestId) },
      request: async (command: CommandOf<'read'>) => {
        requests.push(command)
        const q = command.query as { stamp: CfReadStamp }
        const result = wireResult({ stamp: { ...q.stamp } })
        const reply = cfReply(result, 1024, command.requestId, command.owner as CacheIdentity)
        ledger.set(reply, 1024)
        return reply
      },
      nextRequestId: () => `r${++n}`,
      release: (reply: WorkerReply<'read'>) => { releasedReplies.push(reply); ledger.delete(reply) },
      // Refuse every walk scratch chunk: the split measurement cannot run.
      admitEntry: (key: string, bytes: number) => {
        if (key.startsWith('cf-walk-')) return null
        ledger.set(`entry:${key}`, bytes)
        return () => { ledger.delete(`entry:${key}`) }
      },
      changed: () => {},
    }
    const cache = new RemoteCfCache(host, 1 << 20)
    const held = cache.acquire(SHEET, [rect(0, 0)], cfStamp(1))
    await flush()
    await flush()
    // Ready with the full reply charge: no sharing, no split.
    expect(held.state.state).toBe('ready')
    expect(held.state.result!.definitions).not.toBeUndefined()
    expect(cache.sharedDefinitionVersions).toBe(0)
    expect(cache.retainedBytes).toBe(1024 + entryMetaBytes(SHEET, [rect(0, 0)], cfStamp(1)))
    held.release()
    cache.invalidate(SHEET)
    await flush()
    expect(cache.retainedBytes).toBe(0)
    expect(ledger.size).toBe(0)
    cache.close()
  })
})

