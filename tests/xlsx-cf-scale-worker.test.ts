import { describe, expect, it } from 'vitest'
import { createCfMemory, createCfWorkBudget, type CfWorkBudget } from '../src/xlsx/cf/limits'
import type { CfSheetRules, CfStamp, CfLiteral, CfValueSource, CfValuePage, CfLease, CfRule } from '../src/xlsx/cf/model'
import type { SourceRect } from '../src/xlsx/view/coordinates'
import {
  CF_REPLY_ENVELOPE_BYTES,
  type CfReadRequest,
  type CfReadResult,
  type CfReadStamp,
  type CacheIdentity,
  type WorkerReply,
} from '../src/xlsx/view/contracts'
import { prepareCfRemote, releaseSharedCfDefinitions, cfRulesVersion } from '../src/xlsx/view/pagedSource'
import { RemoteCfCache, type CfReadHost } from '../src/worker/viewerRemoteCf'
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import type { CommandOf } from '../src/worker/viewerProtocol'
import { withInjectedCfEvidence } from '../src/xlsx/cf/evidence'

const SHEET = 'scale-worker'
const stamp = (gen = 1, extra: Partial<CfStamp> = {}): CfStamp => ({
  owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
})
const rect = (r0: number, r1: number, c0 = 0, c1 = 0): SourceRect => ({ r0, r1, c0, c1 })
const frozenBudget = (): CfWorkBudget => createCfWorkBudget({ now: () => 0 })
const nc = (value: number): CfLiteral => ({ kind: 'number', value })

function scaleRule(firstRow = 0, lastRow = 999, firstCol = 0, lastCol = 0): CfRule {
  return {
    id: 1,
    priority: 1,
    documentOrder: 0,
    ranges: [{ firstRow, lastRow, firstCol, lastCol }],
    stopIfTrue: false,
    writes: new Set(['colorScale', 'fill']),
    definition: {
      kind: 'colorScale',
      points: [
        { bound: { type: 'min' }, argb: 0xfff8696b },
        { bound: { type: 'max' }, argb: 0xff63be7b },
      ],
      dialect: 'base',
      formulas: [],
    },
  }
}

function dxfBaseRule(firstRow = 0, lastRow = 999, firstCol = 0, lastCol = 0): CfRule {
  return {
    id: 2,
    priority: 2,
    documentOrder: 1,
    ranges: [{ firstRow, lastRow, firstCol, lastCol }],
    stopIfTrue: false,
    writes: new Set(['font.bold']),
    dxf: { 'font.bold': true },
    definition: {
      kind: 'cellIs',
      operator: 'greaterThan',
      operands: [{ kind: 'number', value: 0 }],
      dialect: 'base',
      formulas: [],
    },
  }
}

function fakeSource(
  cells: { row: number; col: number; value: CfLiteral }[],
  opts: { bounds?: { firstRow: number; lastRow: number; firstCol: number; lastCol: number } | null; stamp?: CfStamp; visits?: { n: number } } = {},
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
      for (let i = 0; i < cells.length; i += 64) {
        const chunk = cells.slice(i, i + 64)
        yield {
          value: {
            rows: Uint32Array.from(chunk, c => { if (opts.visits) opts.visits.n += 1; return c.row }),
            cols: new Uint16Array(chunk.map(c => c.col)),
            values: chunk.map(c => c.value),
          },
          bytes: 0,
          release() {},
        } as unknown as CfLease<CfValuePage>
      }
    },
  } as unknown as CfValueSource & { live: { stamp: CfStamp } }
}

const cfRequest = (rects: readonly SourceRect[], st: CfStamp, grantBytes: number): CfReadRequest => ({
  rects,
  stamp: { owner: st.owner, sheetId: st.sheetId, workbookValueGen: st.workbookValueGen, cfGen: st.cfGen, viewGen: st.viewGen, structureGen: st.structureGen, clockGen: st.clockGen },
  grantBytes,
})

const scaleEvidence = [
  'scale:thresholds-min-max',
  'scale:interpolation-color-space',
  'scale:interpolation-rounding-ladder',
  'scale:participation-types',
]

describe('CF Stage 2 Task 15 — Scale worker payload and invalidation', () => {
  it('scale payload fits the grant or is omitted', async () => {
    // 1,000-cell viewport: rows 0..999, col 0
    const cellCount = 1000
    const cells = Array.from({ length: cellCount }, (_, i) => ({ row: i, col: 0, value: nc(i) }))
    const rules: CfSheetRules = {
      diagnostics: [],
      rules: [scaleRule(0, cellCount - 1, 0, 0), dxfBaseRule(0, cellCount - 1, 0, 0)],
    }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })

    await withInjectedCfEvidence(scaleEvidence, async () => {
      // Wide grant to observe full payload size
      const wideMemory = createCfMemory({})
      const wide = await prepareCfRemote(
        cfRequest([rect(0, cellCount - 1)], st, 1 << 24),
        SHEET,
        fakeSource(cells, { stamp: st }),
        rules,
        wideMemory,
        { budget: frozenBudget },
      )
      expect(wide.kind).toBe('ready')
      if (wide.kind !== 'ready') throw new Error('expected ready')
      expect(wide.states).toHaveLength(cellCount)
      expect(wide.states[0]?.scale).toBeDefined()
      expect(wide.states[0]?.scale?.argb).toBeDefined()

      // Exact measure of the full reply carrying scale
      const fullMeasured = measurePayloadBytes(wide, () => true)!
      const exactGrant = fullMeasured + CF_REPLY_ENVELOPE_BYTES

      // At grant - 1: base cells still return, scale field is omitted, with bounded diagnostic
      const tightMemory = createCfMemory({})
      const tight = await prepareCfRemote(
        cfRequest([rect(0, cellCount - 1)], st, exactGrant - 1),
        SHEET,
        fakeSource(cells, { stamp: st }),
        rules,
        tightMemory,
        { budget: frozenBudget },
      )

      expect(tight.kind).toBe('ready')
      if (tight.kind !== 'ready') throw new Error('expected ready on tight grant')
      // Base cells still return!
      expect(tight.states).toHaveLength(cellCount)
      // Scale field is omitted!
      expect(tight.states[0]?.scale).toBeUndefined()
      // Base properties still return!
      expect(tight.states[1]?.properties['font.bold']).toBe(true)
      // Bounded diagnostic only if it fits
      expect(tight.diagnostics.length).toBeGreaterThan(0)
      expect(tight.diagnostics.some(d => d.feature.includes('grant-required'))).toBe(true)

      // At exactly the grant: reply carries scale
      const exactMemory = createCfMemory({})
      const exact = await prepareCfRemote(
        cfRequest([rect(0, cellCount - 1)], st, exactGrant),
        SHEET,
        fakeSource(cells, { stamp: st }),
        rules,
        exactMemory,
        { budget: frozenBudget },
      )
      expect(exact.kind).toBe('ready')
      if (exact.kind !== 'ready') throw new Error('expected ready on exact grant')
      expect(exact.states).toHaveLength(cellCount)
      expect(exact.states[0]?.scale).toBeDefined()

      releaseSharedCfDefinitions(rules)
    })
  })

  it('the advertised requiredBytes of a refused scale read is enough to carry every scale (retry at exactly that grant)', async () => {
    const cellCount = 1000
    const cells = Array.from({ length: cellCount }, (_, i) => ({ row: i, col: 0, value: nc(i) }))
    const rules: CfSheetRules = { diagnostics: [], rules: [scaleRule(0, cellCount - 1, 0, 0)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    await withInjectedCfEvidence(scaleEvidence, async () => {
      const refused = await prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, 4096), SHEET, fakeSource(cells, { stamp: st }), rules, createCfMemory({}), { budget: frozenBudget })
      const required = (refused as { requiredBytes?: number }).requiredBytes
      expect(refused.kind).not.toBe('ready')
      expect(required).toBeGreaterThan(4096)
      const retry = await prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, required!), SHEET, fakeSource(cells, { stamp: st }), rules, createCfMemory({}), { budget: frozenBudget })
      expect(retry.kind).toBe('ready')
      if (retry.kind !== 'ready') throw new Error('expected ready')
      expect(retry.states.filter(state => state.scale !== undefined)).toHaveLength(cellCount)
      releaseSharedCfDefinitions(rules)
    })
  })

  it('clone copies are counted on both sides until adoption', async () => {
    // Setup cache host that monitors ledger allocations for grant and replies
    const granted = new Map<string, number>()
    const ledger = new Map<string | object, number>()
    const releasedReplies: WorkerReply<'read'>[] = []
    let reqId = 0
    const budgetBytes = 1 << 20

    const host: CfReadHost = {
      identity: () => ({ docId: 'doc', sheetId: SHEET, valueGen: 1, fontGen: 1, viewGen: 1 }),
      loadState: () => ({ completeness: 'complete' as const }),
      hasConditionalFormatting: () => true,
      grant: (requestId: string) => {
        granted.set(requestId, budgetBytes)
        ledger.set(`grant:${requestId}`, budgetBytes)
        return budgetBytes
      },
      releaseGrant: (requestId: string) => {
        granted.delete(requestId)
        ledger.delete(`grant:${requestId}`)
      },
      request: async (command: CommandOf<'read'>) => {
        const q = command.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        const result: CfReadResult = {
          stamp: { ...q.stamp },
          sheetId: SHEET,
          kind: 'ready',
          states: [{ row: 0, col: 0, properties: {}, scale: { ruleId: 1, argb: 0xff63be7b } }],
          definitions: [],
          definitionsVersion: 1,
          diagnostics: [],
        }
        const replyBytes = measurePayloadBytes(result, () => true)!
        const reply: WorkerReply<'read'> = {
          commandKind: 'read',
          kind: 'complete',
          requestId: command.requestId,
          owner: command.owner as CacheIdentity,
          replyBytes,
          diagnostics: [],
          result: { kind: 'read', result: { op: 'conditionalFormatting', result } },
        }
        granted.delete(command.requestId)
        ledger.delete(`grant:${command.requestId}`)
        ledger.set(reply, replyBytes)
        return reply
      },
      nextRequestId: () => `req-${++reqId}`,
      release: (reply: WorkerReply<'read'>) => {
        releasedReplies.push(reply)
        ledger.delete(reply)
      },
      admitEntry: (key: string, bytes: number) => {
        ledger.set(`entry:${key}`, bytes)
        return () => { ledger.delete(`entry:${key}`) }
      },
      changed: () => {},
    }

    const cache = new RemoteCfCache(host, 1 << 20)
    const st = { owner: 'doc', sheetId: SHEET, workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1 }

    const handle = cache.acquire(SHEET, [rect(0, 0)], st)
    await new Promise(r => setTimeout(r, 10))

    expect(handle.state.state).toBe('ready')
    // After adoption, the reply was released
    expect(releasedReplies).toHaveLength(1)
    // The entry states are admitted in the ledger
    expect([...ledger.keys()].some(k => typeof k === 'string' && k.startsWith('entry:cf-states:'))).toBe(true)

    handle.release()
    cache.close()
  })

  it('max edit recolours the whole scope in one generation', async () => {
    // Shell at g1: values 0..10. Max is 10.
    const rule = scaleRule(0, 4, 0, 0)
    const rules: CfSheetRules = { diagnostics: [], rules: [rule] }
    const version = cfRulesVersion(rules)
    const stG1 = stamp(1, { cfGen: version })
    const cellsG1 = [
      { row: 0, col: 0, value: nc(0) },
      { row: 1, col: 0, value: nc(5) },
      { row: 2, col: 0, value: nc(10) },
    ]

    await withInjectedCfEvidence(scaleEvidence, async () => {
      const memory = createCfMemory({})
      const resG1 = await prepareCfRemote(
        cfRequest([rect(0, 2)], stG1, 1 << 20),
        SHEET,
        fakeSource(cellsG1, { stamp: stG1 }),
        rules,
        memory,
        { budget: frozenBudget },
      )
      expect(resG1.kind).toBe('ready')
      if (resG1.kind !== 'ready') throw new Error('expected ready')

      const colorAt10G1 = resG1.states.find(s => s.row === 2)?.scale?.argb
      // At g1, value 10 is max (10/10) -> top color (green 0xff63be7b)
      expect(colorAt10G1).toBe(0xff63be7b)

      // Commit value 20 at g2 (so max becomes 20)
      const stG2 = stamp(2, { cfGen: version })
      const cellsG2 = [
        { row: 0, col: 0, value: nc(0) },
        { row: 1, col: 0, value: nc(5) },
        { row: 2, col: 0, value: nc(10) },
        { row: 3, col: 0, value: nc(20) }, // new max 20
      ]

      // Stale generation check: request with g1 against source with g2 withholds
      const staleCheck = await prepareCfRemote(
        cfRequest([rect(0, 3)], stG1, 1 << 20),
        SHEET,
        fakeSource(cellsG2, { stamp: stG2 }),
        rules,
        memory,
        { budget: frozenBudget },
      )
      expect(staleCheck.kind).toBe('unresolved')
      expect(staleCheck.reason).toBe('cf-stale-generation')

      // Prepare fresh at g2
      const resG2 = await prepareCfRemote(
        cfRequest([rect(0, 3)], stG2, 1 << 20),
        SHEET,
        fakeSource(cellsG2, { stamp: stG2 }),
        rules,
        memory,
        { budget: frozenBudget },
      )
      expect(resG2.kind).toBe('ready')
      if (resG2.kind !== 'ready') throw new Error('expected ready at g2')

      // In g2: value 10 is no longer max (it is 10/20 = 50% midpoint) -> recoloured!
      const colorAt10G2 = resG2.states.find(s => s.row === 2)?.scale?.argb
      expect(colorAt10G2).not.toBe(colorAt10G1)
      // New max (20) gets top color
      const colorAt20G2 = resG2.states.find(s => s.row === 3)?.scale?.argb
      expect(colorAt20G2).toBe(0xff63be7b)

      releaseSharedCfDefinitions(rules)
    })
  })

  it('unimported scope pages withhold then repaint once on the paged path', async () => {
    const rules: CfSheetRules = { diagnostics: [], rules: [scaleRule(0, 10, 0, 0)] }
    const version = cfRulesVersion(rules)
    const st = stamp(1, { cfGen: version })

    await withInjectedCfEvidence(scaleEvidence, async () => {
      const memory = createCfMemory({})
      // Source whose bounds are initially undefined (unimported)
      let imported = false
      const source: CfValueSource = {
        bounds: async () => imported ? { firstRow: 0, lastRow: 10, firstCol: 0, lastCol: 0 } : undefined,
        stamp: () => st,
        pages: async function* () {
          if (!imported) return
          yield {
            value: {
              rows: Uint32Array.from([0, 10]),
              cols: new Uint16Array([0, 0]),
              values: [nc(0), nc(100)],
            },
            bytes: 0,
            release() {},
          } as unknown as CfLease<CfValuePage>
        },
      }

      // 1. Unimported: withholds with pendingSheetIds, base fill, no permanent diagnostic
      const withheld = await prepareCfRemote(
        cfRequest([rect(0, 10)], st, 1 << 20),
        SHEET,
        source,
        rules,
        memory,
        { budget: frozenBudget },
      )
      expect(withheld.kind).toBe('pending')
      expect(withheld.reason).toBe('cf-pending-data')
      expect(withheld.pendingSheetIds).toEqual([SHEET])
      // Diagnostics must NOT report a permanent error
      expect(withheld.diagnostics.every(d => d.kind !== 'unsupported')).toBe(true)

      // 2. On arrival (imported = true): targeted invalidation / reload repaints once
      imported = true
      const arrived = await prepareCfRemote(
        cfRequest([rect(0, 10)], st, 1 << 20),
        SHEET,
        source,
        rules,
        memory,
        { budget: frozenBudget },
      )
      expect(arrived.kind).toBe('ready')
      expect(arrived.states.length).toBeGreaterThan(0)
      expect(arrived.states[0]?.scale).toBeDefined()

      releaseSharedCfDefinitions(rules)
    })
  })
})
