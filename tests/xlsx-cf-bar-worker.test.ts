import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { CF_BAR_CELL_BYTES, CF_SCALE_CELL_BYTES, createCfMemory, createCfWorkBudget, type CfWorkBudget } from '../src/xlsx/cf/limits'
import type { CfBarDescriptor, CfDefinition, CfLease, CfLiteral, CfRule, CfSheetRules, CfStamp, CfValuePage, CfValueSource } from '../src/xlsx/cf/model'
import type { SourceRect } from '../src/xlsx/view/coordinates'
import {
  CF_REPLY_ENVELOPE_BYTES,
  estimateCfPreparationBytes,
  measureCfCellBytes,
  measureCfPropertiesBytes,
  type CacheIdentity,
  type CfCellStateWire,
  type CfReadRequest,
  type CfReadResult,
  type CfReadStamp,
  type WorkerReply,
} from '../src/xlsx/view/contracts'
import { CfPageFault, cfRulesVersion, prepareCfRemote, releaseSharedCfDefinitions } from '../src/xlsx/view/pagedSource'
import { RemoteCfCache, type CfReadHost } from '../src/worker/viewerRemoteCf'
import { admitReply } from '../src/worker/viewerReplyLimit'
import type { CommandOf } from '../src/worker/viewerProtocol'
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { injectCfBarLayout, withInjectedCfEvidence } from '../src/xlsx/cf/evidence'
import { createCfService } from '../src/xlsx/cf/service'
import { residentSource } from '../src/xlsx/cf/render'
import { prepareCfTextScope, prepareCfWireScope } from '../src/xlsx/cf/text'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import type { XlsxSheet } from '../src/xlsx/types'

// Every bar in this file comes from injected evidence and an injected layout (Task 21 addendum §1): nothing Stage 3 is
// published, so production withholds every bar. The numbers below are hypotheses run under stubs, never Excel facts.

const SHEET = 'bar-worker'
const REPO_ROOT = resolve(__dirname, '..')
const BLUE = 0xff0000ff, RED = 0xffff0000, BLACK = 0xff000000
const stamp = (gen = 1, extra: Partial<CfStamp> = {}): CfStamp => ({
  owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
})
const rect = (r0: number, r1: number, c0 = 0, c1 = 0): SourceRect => ({ r0, r1, c0, c1 })
const frozenBudget = (): CfWorkBudget => createCfWorkBudget({ now: () => 0 })
const nc = (value: number): CfLiteral => ({ kind: 'number', value })

/** Every evaluation capability these bars can need, plus the parse gates of the Excel fixture (stubs only). */
const BAR_EVIDENCE = [
  'bar:bounds-min-max', 'bar:bounds-num', 'bar:bounds-percent', 'bar:bounds-percentile', 'bar:bounds-automin', 'bar:bounds-automax',
  'bar:lengths-default-10-90', 'bar:lengths-explicit-0-100', 'bar:render-solid', 'bar:render-border',
  'bar:axis-automatic', 'bar:axis-middle', 'bar:axis-none', 'bar:direction-ltr', 'bar:direction-rtl', 'bar:direction-context-ltr-sheet',
  'bar:negative-fill-color', 'bar:negative-fill-same-as-positive', 'bar:negative-border-color', 'bar:negative-border-same-as-positive',
  'bar:participation-types',
  'bar:settings-precondition-probe', 'bar:settings-schema-defaults', 'bar:settings-paired-precedence',
  'bar:pairing-paired', 'bar:pairing-standalone', 'bar:pairing-base-only',
  'bar:colors-theme-tint', 'bar:colors-indexed', 'bar:bounds-positional', 'bar:bounds-out-of-domain', 'bar:degenerate-inverted',
]
const SCALE_EVIDENCE = ['scale:thresholds-min-max', 'scale:interpolation-color-space', 'scale:interpolation-rounding-ladder', 'scale:participation-types']

/** The layout the service reads from evidence (S3-B), injected the way Task 23's export would publish it. */
const LAYOUT = {
  allPositive: { origin: 'span', lengthBasis: 'cell', settledBy: ['bar:lengths-explicit-0-100'] },
  mixedSign: { origin: 'axis', lengthBasis: 'side', settledBy: ['bar:axis-automatic', 'bar:axis-middle'] },
}
/** Bars under stub evidence: the injected set joins the production evidence, so Stage 1/2 rules keep theirs. */
const withBars = <T>(fn: () => Promise<T>, extra: readonly string[] = []): Promise<T> => {
  injectCfBarLayout(LAYOUT)
  return withInjectedCfEvidence([...BAR_EVIDENCE, ...SCALE_EVIDENCE, ...extra], fn)
}
afterEach(() => { injectCfBarLayout(null) })

type BarDef = Extract<CfDefinition, { kind: 'dataBar' }>
function barDef(overrides: Partial<BarDef> = {}): BarDef {
  return {
    kind: 'dataBar', dialect: 'base', formulas: [],
    bounds: [{ type: 'min' }, { type: 'max' }],
    fillArgb: BLUE, negativeFillArgb: RED, axisArgb: BLACK,
    minLength: 0, maxLength: 100, gradient: false, border: false,
    negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: true,
    axisPosition: 'automatic', direction: 'leftToRight', showValue: true, sheetRightToLeft: false,
    ...overrides,
  }
}
function barRule(firstRow = 0, lastRow = 999, id = 1, priority = 1): CfRule {
  return { id, priority, documentOrder: priority - 1, ranges: [{ firstRow, lastRow, firstCol: 0, lastCol: 0 }], stopIfTrue: false, writes: new Set(['dataBar']), definition: barDef() }
}
function scaleRule(firstRow = 0, lastRow = 999): CfRule {
  return {
    id: 1, priority: 1, documentOrder: 0, ranges: [{ firstRow, lastRow, firstCol: 0, lastCol: 0 }], stopIfTrue: false,
    writes: new Set(['colorScale', 'fill']),
    definition: { kind: 'colorScale', points: [{ bound: { type: 'min' }, argb: 0xfff8696b }, { bound: { type: 'max' }, argb: 0xff63be7b }], dialect: 'base', formulas: [] },
  }
}
function dxfBaseRule(firstRow = 0, lastRow = 999): CfRule {
  return {
    id: 2, priority: 2, documentOrder: 1, ranges: [{ firstRow, lastRow, firstCol: 0, lastCol: 0 }], stopIfTrue: false,
    writes: new Set(['font.bold']), dxf: { 'font.bold': true },
    definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
  }
}

/** Column A values; the source resolves number formats (resident-like), so bars can participate (S3-G). */
function columnSource(
  values: number[],
  opts: { stamp?: () => CfStamp; visits?: { n: number }; gate?: () => Promise<void>; fault?: () => CfPageFault | undefined } = {},
): CfValueSource {
  // Values are read at each call, so a test can commit a new value between preparations.
  const lits = { get length() { return values.length } }
  const at = (row: number): CfLiteral => nc(values[row]!)
  return {
    numberFormats: 'resolved',
    bounds: async () => ({ firstRow: 0, lastRow: Math.max(0, lits.length - 1), firstCol: 0, lastCol: 0 }),
    stamp: () => opts.stamp?.() ?? stamp(1),
    pages: async function* (_id: string, rects: readonly { firstRow: number; lastRow: number; firstCol: number; lastCol: number }[]) {
      await opts.gate?.()
      const fault = opts.fault?.()
      if (fault) throw fault
      const rows: number[] = [], vals: CfLiteral[] = []
      for (const r of rects) for (let row = r.firstRow; row <= Math.min(r.lastRow, lits.length - 1); row++) {
        if (r.firstCol <= 0 && r.lastCol >= 0) { rows.push(row); vals.push(at(row)); if (opts.visits) opts.visits.n++ }
      }
      yield { value: { rows: Uint32Array.from(rows), cols: new Uint16Array(rows.length), values: vals }, bytes: 0, release() {} } as unknown as CfLease<CfValuePage>
    },
  } as CfValueSource
}

const cfRequest = (rects: readonly SourceRect[], st: CfStamp, grantBytes: number): CfReadRequest => ({
  rects,
  stamp: { owner: st.owner, sheetId: st.sheetId, workbookValueGen: st.workbookValueGen, cfGen: st.cfGen, viewGen: st.viewGen, structureGen: st.structureGen, clockGen: st.clockGen },
  grantBytes,
})
const ready = (result: CfReadResult): CfReadResult => {
  expect(result.kind, JSON.stringify({ reason: result.reason, diagnostics: result.diagnostics })).toBe('ready')
  return result
}
const bars = (result: CfReadResult) => result.states.filter(s => s.bar !== undefined)
/** The production decode of a wire bar: `rtl` back to the descriptor's `direction`. */
const decode = (bar: NonNullable<CfCellStateWire['bar']>): CfBarDescriptor => {
  const { rtl, ...rest } = bar
  return { ...rest, direction: rtl ? 'rightToLeft' : 'leftToRight' }
}
/** The fixed framing a preparation measures before it reads a page: an empty ready result carrying the definitions. */
const templateBytes = (rules: CfSheetRules, st: CfStamp): number => measurePayloadBytes({
  sheetId: SHEET, stamp: { ...st }, kind: 'ready', states: [],
  definitions: rules.rules.map((rule, index) => ({ id: `rule-${typeof rule.id === 'number' ? rule.id : index}`, version: st.cfGen, properties: { ...(rule.dxf ?? {}) }, writes: [...rule.writes] })),
  definitionsVersion: st.cfGen, diagnostics: [],
}, () => true)!
/** Per-cell wire bound the preflight prices with (Stage 2 formula; visuals added from the measurer constants). */
const perCellBase = (rules: CfSheetRules): number => rules.rules.reduce((n, r) => n + measureCfPropertiesBytes((r.dxf ?? {}) as Record<string, string | number | boolean | null>), 96)
const command = (grantBytes: number, st: CfStamp = stamp(1), rects: readonly SourceRect[] = [rect(0, 999)]): CommandOf<'read'> => ({
  kind: 'read', requestId: 'r1', owner: { docId: 'doc', sheetId: SHEET, valueGen: 1, fontGen: 0, viewGen: 0 },
  query: { op: 'conditionalFormatting', rects, stamp: { ...st }, grantBytes },
} as unknown as CommandOf<'read'>)
const wireReply = (cmd: CommandOf<'read'>, result: CfReadResult): WorkerReply => ({
  commandKind: 'read', kind: 'complete', requestId: cmd.requestId, owner: cmd.owner, diagnostics: [],
  result: { kind: 'read', result: { op: 'conditionalFormatting', result } },
} as unknown as WorkerReply)
const budgetOf = (): FormatBudget & { held: number } => {
  const b = { held: 0, reserve(n: number) { b.held += n }, release(n: number) { b.held -= n } }
  return b
}

describe('CF Stage 3 Task 21 — bar wire under the receiver grant (S3-S)', () => {
  it('bar payload fits the grant or is omitted', async () => {
    const cellCount = 1000
    const values = Array.from({ length: cellCount }, (_, i) => i)
    const rules: CfSheetRules = { diagnostics: [], rules: [barRule(0, cellCount - 1), dxfBaseRule(0, cellCount - 1)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    const run = (grant: number, memory = createCfMemory({}), visits?: { n: number }, with_: CfSheetRules = rules) =>
      prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, grant), SHEET, columnSource(values, { stamp: () => st, visits }), with_, memory, { budget: frozenBudget })
    await withBars(async () => {
      // Priced before the payload is built: the advertised bound carries CF_BAR_CELL_BYTES per cell, and a tight grant
      // refuses without reading a single page.
      const pricedVisits = { n: 0 }
      const refused = await run(4096, createCfMemory({}), pricedVisits)
      expect(refused.kind).not.toBe('ready')
      expect(pricedVisits.n).toBe(0)
      const perCellBar = perCellBase(rules) + CF_BAR_CELL_BYTES
      expect(refused.requiredBytes).toBe(templateBytes(rules, st) + estimateCfPreparationBytes(cellCount, perCellBar, 0, 0, 0)! + CF_REPLY_ENVELOPE_BYTES)
      // The working hold grows to cover the bars before any state exists: a pool that admits the base but not the
      // bars refuses before a page is read (the grant itself is wide).
      const baseBound = templateBytes(rules, st) + cellCount * perCellBase(rules) + CF_REPLY_ENVELOPE_BYTES
      const heldVisits = { n: 0 }
      const pressed = await run(1 << 24, createCfMemory({ limitBytes: baseBound + (cellCount * CF_BAR_CELL_BYTES) / 2 }), heldVisits)
      expect(pressed.kind).not.toBe('ready')
      expect(heldVisits.n).toBe(0)
      // Refused at the working-hold growth itself: the only site that carries the visual-priced bound.
      expect(pressed.diagnostics.map(d => d.feature)).toEqual(['cf-budget-exceeded'])
      expect(pressed.requiredBytes).toBe(refused.requiredBytes)

      // Positive control: a wide grant carries one bar per scope cell.
      const wide = ready(await run(1 << 24))
      expect(wide.states).toHaveLength(cellCount)
      expect(bars(wide)).toHaveLength(cellCount)
      expect(wide.states[999]?.bar).toMatchObject({ ruleId: 1, startFraction: 0, endFraction: 1, rtl: false })
      expect(wide.states[1]?.properties['font.bold']).toBe(true)
      const exactGrant = measurePayloadBytes(wide, () => true)! + CF_REPLY_ENVELOPE_BYTES

      // At grant − 1: base cells return, every bar field is omitted, and cf-grant-required rides along because it fits.
      const tight = ready(await run(exactGrant - 1))
      expect(tight.states).toHaveLength(cellCount)
      expect(tight.states.every(s => !('bar' in s))).toBe(true)
      expect(tight.states[1]?.properties['font.bold']).toBe(true)
      expect(tight.diagnostics.some(d => d.feature === 'cf-grant-required')).toBe(true)
      expect(measurePayloadBytes(tight, () => true)! + CF_REPLY_ENVELOPE_BYTES).toBeLessThanOrEqual(exactGrant - 1)
      // A grant that admits the stripped base but not the diagnostics (64 retained parse diagnostics of 256 units ride
      // along): base cells still return, every bar is omitted, and no diagnostic is sent because none fits.
      const noisy: CfSheetRules = { rules: rules.rules, diagnostics: Array.from({ length: 64 }, (_, i) => ({ kind: 'unsupported' as const, feature: `cf-noise-${i}` as const, message: 'x'.repeat(256) })) }
      const strippedBase = measurePayloadBytes({ ...tight, diagnostics: [] }, () => true)! + CF_REPLY_ENVELOPE_BYTES
      const strippedGrant = Math.max(strippedBase, baseBound)
      const noisyStripped = ready(await run(1 << 24, createCfMemory({}), undefined, noisy))
      const strippedWithDiagnostics = measurePayloadBytes({ ...noisyStripped, states: tight.states }, () => true)! + CF_REPLY_ENVELOPE_BYTES
      expect(strippedWithDiagnostics, 'precondition: the diagnostics do not fit').toBeGreaterThan(strippedGrant)
      const bare = ready(await run(strippedGrant, createCfMemory({}), undefined, noisy))
      expect(bare.states).toHaveLength(cellCount)
      expect(bare.states.every(s => !('bar' in s))).toBe(true)
      expect(bare.states[1]?.properties['font.bold']).toBe(true)
      expect(bare.diagnostics).toEqual([])

      // At exactly the grant the reply carries every bar, identical to the wide reply.
      const exact = ready(await run(exactGrant))
      expect(exact.states).toEqual(wide.states)

      // The kernel admission never replaces a provider reply that the provider fitted to its grant.
      for (const [grant, result] of [[exactGrant - 1, tight], [exactGrant, exact], [strippedGrant, bare]] as const) {
        const cmd = command(grant, st)
        const out = admitReply(cmd, wireReply(cmd, result), grant, () => true, false)
        expect(out.kind, `grant ${grant}`).toBe('complete')
        expect(out.replyBytes).toBeLessThanOrEqual(grant)
      }
      releaseSharedCfDefinitions(rules)
    })
  })

  it('a sheet with scales and bars omits the bars first, then the scales (ordering pinned; see report)', async () => {
    const cellCount = 1000
    const values = Array.from({ length: cellCount }, (_, i) => i)
    const rules: CfSheetRules = { diagnostics: [], rules: [scaleRule(0, cellCount - 1), { ...barRule(0, cellCount - 1, 3, 2) }] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    const run = (grant: number) => prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, grant), SHEET, columnSource(values, { stamp: () => st }), rules, createCfMemory({}), { budget: frozenBudget })
    await withBars(async () => {
      const wide = ready(await run(1 << 24))
      expect(bars(wide)).toHaveLength(cellCount)
      expect(wide.states.filter(s => s.scale !== undefined)).toHaveLength(cellCount)
      const full = measurePayloadBytes(wide, () => true)! + CF_REPLY_ENVELOPE_BYTES
      const grantMessages = (r: CfReadResult) => r.diagnostics.filter(d => d.feature === 'cf-grant-required').map(d => d.message)
      // One byte short: only the bars go, every scale stays.
      const noBars = ready(await run(full - 1))
      expect(noBars.states.every(s => !('bar' in s))).toBe(true)
      expect(noBars.states.filter(s => s.scale !== undefined)).toHaveLength(cellCount)
      expect(grantMessages(noBars)).toEqual(['Conditional formatting data bars omitted due to receiver grant limit.'])
      // Short of the scales too: both visuals go, the base cells return.
      const scalesOnly = measurePayloadBytes({ ...noBars, diagnostics: [] }, () => true)! + CF_REPLY_ENVELOPE_BYTES
      const neither = ready(await run(scalesOnly - 1))
      expect(neither.states).toHaveLength(cellCount)
      expect(neither.states.every(s => !('bar' in s) && !('scale' in s))).toBe(true)
      expect(grantMessages(neither)).toEqual([
        'Conditional formatting data bars omitted due to receiver grant limit.',
        'Conditional formatting scale fill omitted due to receiver grant limit.',
      ])
      releaseSharedCfDefinitions(rules)
    }, ['bar:composition-with-scale'])
  })

  it('the advertised requiredBytes of a refused bar read carries every bar (retry at exactly that grant)', async () => {
    const cellCount = 1000
    const values = Array.from({ length: cellCount }, (_, i) => i - 500)
    const rules: CfSheetRules = { diagnostics: [], rules: [barRule(0, cellCount - 1)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    const run = (grant: number) => prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, grant), SHEET, columnSource(values, { stamp: () => st }), rules, createCfMemory({}), { budget: frozenBudget })
    await withBars(async () => {
      const refused = await run(4096)
      expect(refused.kind).not.toBe('ready')
      const required = refused.requiredBytes!
      expect(required).toBeGreaterThan(4096)
      const retry = ready(await run(required))
      // Mixed sign under the automatic axis: every cell carries its bar, negative ones included.
      expect(bars(retry)).toHaveLength(cellCount)
      expect(retry.states.some(s => s.bar?.isNegative === true)).toBe(true)
      expect(retry.diagnostics.some(d => d.feature === 'cf-grant-required')).toBe(false)
      releaseSharedCfDefinitions(rules)
    })
  })

  it('clone copies of bar states are counted on both sides until adoption', async () => {
    // Two scopes: a small one the cache splits into per-entry states plus shared definitions, and a 200-cell one whose
    // walk workspace (136 B per object) outgrows the reply-sized walk cap, so the entry keeps the full reply charge.
    for (const [cellCount, path] of [[2, 'split'], [200, 'reply']] as const) {
      const values = Array.from({ length: cellCount }, (_, i) => i)
      const rules: CfSheetRules = { diagnostics: [], rules: [barRule(0, cellCount - 1)] }
      const st = stamp(1, { cfGen: cfRulesVersion(rules) })
      await withBars(async () => {
        // Worker side: the outgoing result stays charged on the provider account until the post releases it.
        const memory = createCfMemory({})
        const releases: Array<() => void> = []
        const produced = ready(await prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, 1 << 22), SHEET, columnSource(values, { stamp: () => st }), rules, memory,
          { budget: frozenBudget, retainOutgoing: release => releases.push(release) }))
        expect(bars(produced)).toHaveLength(cellCount)
        const wire = measurePayloadBytes(produced, () => true)!
        expect(releases).toHaveLength(1)
        const held = memory.currentBytes
        expect(held).toBeGreaterThanOrEqual(wire)
        // Main side: the structured clone arrives charged as the reply.
        const clone = structuredClone(produced)
        const replyBytes = measurePayloadBytes(clone, () => true)!
        expect(replyBytes).toBe(wire)
        const ledger = new Map<string | object, number>()
        const replyHeldAtStatesAdmission: boolean[] = []
        let reqId = 0
        const host: CfReadHost = {
          identity: () => ({ docId: 'doc', sheetId: SHEET, valueGen: 1, fontGen: 1, viewGen: 1 }),
          loadState: () => ({ completeness: 'complete' as const }),
          hasConditionalFormatting: () => true,
          grant: requestId => { ledger.set(`grant:${requestId}`, 1 << 22); return 1 << 22 },
          releaseGrant: requestId => { ledger.delete(`grant:${requestId}`) },
          request: async cmd => {
            const reply = { commandKind: 'read', kind: 'complete', requestId: cmd.requestId, owner: cmd.owner as CacheIdentity, replyBytes, diagnostics: [],
              result: { kind: 'read', result: { op: 'conditionalFormatting', result: clone } } } as WorkerReply<'read'>
            ledger.delete(`grant:${cmd.requestId}`)
            ledger.set(reply, replyBytes)
            return reply
          },
          nextRequestId: () => `req-${++reqId}`,
          release: reply => { ledger.delete(reply) },
          admitEntry: (key, bytes) => {
            if (key.startsWith('cf-states:')) replyHeldAtStatesAdmission.push([...ledger.keys()].some(k => typeof k === 'object'))
            ledger.set(`entry:${key}`, bytes)
            return () => { ledger.delete(`entry:${key}`) }
          },
          changed: () => {},
        }
        const cache = new RemoteCfCache(host, 1 << 22)
        const handle = cache.acquire(SHEET, [rect(0, cellCount - 1)], { ...st })
        await new Promise(r => setTimeout(r, 10))
        expect(handle.state.state).toBe('ready')
        expect(handle.state.result?.states.filter(s => s.bar !== undefined)).toHaveLength(cellCount)
        const replyHeld = [...ledger.keys()].some(k => typeof k === 'object')
        const statesKey = [...ledger.keys()].find(k => typeof k === 'string' && k.startsWith('entry:cf-states:'))
        if (path === 'split') {
          // The states are admitted while the reply charge is still held; the reply ends only at adoption.
          expect(replyHeldAtStatesAdmission).toEqual([true])
          expect(replyHeld, 'the reply charge ends at adoption').toBe(false)
          // The adopted states charge is the walked size of the states WITH their bars: one measured bar per cell.
          const statesCharge = ledger.get(statesKey!)!
          expect(statesCharge).toBe(measurePayloadBytes({ ...clone, definitions: [] }, () => true))
          const stripped = { ...clone, definitions: [], states: clone.states.map(({ bar: _bar, ...rest }) => rest) }
          const perBar = clone.states.reduce((n, s) => n + measureCfCellBytes(s) - measureCfCellBytes({ ...s, bar: undefined }), 0)
          expect(statesCharge - measurePayloadBytes(stripped, () => true)!).toBe(perBar)
          expect(perBar).toBeGreaterThan(0)
        } else if (statesKey === undefined) {
          // No split today (the walk workspace outgrows the reply-sized cap): the entry keeps the transport's reply
          // charge, which is the measured clone with every bar.
          expect(replyHeld).toBe(true)
          expect([...ledger.entries()].find(([k]) => typeof k === 'object')?.[1]).toBe(replyBytes)
        } else {
          // Should the cache split larger replies later, the invariant is the same: one charge, the measured states.
          expect(replyHeld).toBe(false)
          expect(ledger.get(statesKey)).toBe(measurePayloadBytes({ ...clone, definitions: [] }, () => true))
        }
        handle.release()
        cache.close()
        expect(ledger.size).toBe(0)
        // The worker share ends only when the post releases it.
        for (const release of releases) release()
        expect(memory.currentBytes).toBeLessThanOrEqual(held - wire)
        releaseSharedCfDefinitions(rules)
      })
    }
  })

  it('wire bars equal the resident descriptor', async () => {
    const bytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/stage3-bar-axis-negative.xlsx')))
    await withBars(async () => {
      const sheet = (await parseXlsx(await OfficePackage.load(bytes))).sheets[0]!
      const rules = sheet.conditionalFormatting!
      const sheetId = sheet.sheetId ?? sheet.name
      // B2:I6: every bar of the fixture (automatic, middle and no axis; negative fills and borders; all-positive).
      const scope = [rect(1, 5, 1, 8)]
      const st: CfStamp = { owner: 'fixture', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
      const residentMemory = createCfMemory({})
      const service = createCfService(rules, residentMemory, {})
      const lease = await service.prepareCf({ stamp: st, sourceRects: [{ firstRow: 1, lastRow: 5, firstCol: 1, lastCol: 8 }], signal: new AbortController().signal, grantBytes: 8 << 20 },
        residentSource(sheet, st, residentMemory, 'engine'))
      const wireMemory = createCfMemory({})
      const wire = ready(await prepareCfRemote(cfRequest(scope, st, 1 << 22), sheetId, residentSource(sheet, st, wireMemory, 'engine'), rules, wireMemory, { budget: frozenBudget }))
      // The paged main thread decodes a structured clone of the reply through the production scope.
      const budget = budgetOf()
      const paged = await prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), scope, { budget })
      expect(paged).toBeDefined()
      let compared = 0, negative = 0
      for (let row = 1; row <= 5; row++) for (let col = 1; col <= 8; col++) {
        const at = lease.value.at(row, col)
        const resident = at.kind === 'ready' ? at.value.visual?.dataBar : undefined
        const state = wire.states.find(s => s.row === row && s.col === col)
        const decoded = state?.bar ? decode(state.bar) : undefined
        expect(decoded, `wire ${row},${col}`).toEqual(resident)
        expect(paged!.barAt?.(row, col), `paged ${row},${col}`).toEqual(resident)
        if (resident) { compared++; if (resident.isNegative) negative++ }
      }
      // Positive controls: the mixed-sign columns resolve, negative bars included.
      expect(compared).toBeGreaterThanOrEqual(20)
      expect(negative).toBeGreaterThanOrEqual(4)
      paged!.dispose()
      expect(budget.held).toBe(0)
      lease.release(); service.dispose()
      releaseSharedCfDefinitions(rules)
    })
  })

  it('max edit re-lengths every bar of the scope in one generation', async () => {
    const rules: CfSheetRules = { diagnostics: [], rules: [barRule(0, 3)] }
    const version = cfRulesVersion(rules)
    await withBars(async () => {
      // Resident shell at g1: A1:A4 = 0, 5, 10, 0 (max 10).
      const sheet = {
        name: 'Bars', sheetId: SHEET, cols: [], merges: [], mergeRanges: [],
        rows: [0, 5, 10, 0].map((value, r) => ({ index: r, cells: [{ row: r, col: 0, ref: `A${r + 1}`, styleIndex: 0, value }] })),
        conditionalFormatting: rules,
      } as unknown as XlsxSheet
      const source = createObjectWorkbookSource({ sheets: [sheet] }, 'bar-live')
      const store = source.store(SHEET)
      const budget = budgetOf()
      const prepare = () => prepareCfTextScope(sheet, { rects: [rect(0, 3)], signal: new AbortController().signal, budget, store })
      const r1 = (await prepare())!
      expect(r1.barAt?.(2, 0)?.endFraction).toBe(1)
      expect(r1.barAt?.(1, 0)?.endFraction).toBe(0.5)

      // Worker shell at g1, decoded on the paged main thread with generation liveness.
      let liveGen = 1
      let gate: Promise<void> | undefined
      const workerValues = [0, 5, 10, 0]
      const src = columnSource(workerValues, { stamp: () => stamp(liveGen, { cfGen: version }), gate: () => gate ?? Promise.resolve() })
      const read = (gen: number) => prepareCfRemote(cfRequest([rect(0, 3)], stamp(gen, { cfGen: version }), 1 << 20), SHEET, src, rules, createCfMemory({}), { budget: frozenBudget })
      const g1Wire = ready(await read(1))
      const w1 = (await prepareCfWireScope(async () => ({ result: g1Wire, release() {} }), [rect(0, 3)], { budget, current: () => liveGen === 1 }))!
      expect(w1.barAt?.(2, 0)?.endFraction).toBe(1)

      // Commit 20 at g2 on both shells while the g2 preparation is pending.
      sheet.rows[3]!.cells[0]!.value = 20
      store.noteValuesChanged?.()
      workerValues[3] = 20
      liveGen = 2
      let open!: () => void
      gate = new Promise(r => { open = r })
      const pending = read(2)
      const residentPending = prepare()
      for (let row = 0; row <= 3; row++) {
        expect(r1.barAt?.(row, 0), `resident g1 bar at g2, row ${row}`).toBeUndefined()
        expect(w1.barAt?.(row, 0), `worker g1 bar at g2, row ${row}`).toBeUndefined()
      }
      open()
      const g2Wire = ready(await pending)
      const r2 = (await residentPending)!
      const w2 = (await prepareCfWireScope(async () => ({ result: g2Wire, release() {} }), [rect(0, 3)], { budget, current: () => liveGen === 2 }))!
      // Every scope cell gets its g2 bar: max is now 20.
      for (const [row, value] of [[0, 0], [1, 5], [2, 10], [3, 20]] as const) {
        expect(r2.barAt?.(row, 0)?.endFraction, `resident row ${row}`).toBe(value / 20)
        expect(w2.barAt?.(row, 0)?.endFraction, `worker row ${row}`).toBe(value / 20)
      }
      for (const scope of [r1, w1, r2, w2]) scope.dispose()
      expect(budget.held).toBe(0)
      releaseSharedCfDefinitions(rules)
    }, ['bar:participation-dateserial'])
  })

  it('unimported scope pages withhold bars then repaint once', async () => {
    const rules: CfSheetRules = { diagnostics: [], rules: [barRule(0, 9)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    await withBars(async () => {
      // Bounds are known but the scope's pages are still importing: the provider faults `incomplete`.
      let imported = false
      const src = columnSource(Array.from({ length: 10 }, (_, i) => i), {
        stamp: () => st, fault: () => imported ? undefined : new CfPageFault('incomplete', SHEET, 'pages still importing'),
      })
      const provide = async (cmd: CommandOf<'read'>): Promise<CfReadResult> => {
        const q = cmd.query as { rects: readonly SourceRect[]; stamp: CfReadStamp; grantBytes: number }
        return prepareCfRemote({ rects: q.rects, stamp: q.stamp, grantBytes: q.grantBytes }, SHEET, src, rules, createCfMemory({}), { budget: frozenBudget })
      }
      const direct = await provide(command(1 << 20, st, [rect(0, 9)]))
      expect(direct.kind).toBe('pending')
      expect(direct.pendingSheetIds).toEqual([SHEET])
      expect(direct.states).toHaveLength(0)
      expect(direct.diagnostics.every(d => d.kind !== 'unsupported')).toBe(true)

      const changed: string[] = []
      let reqId = 0
      const host: CfReadHost = {
        identity: () => ({ docId: 'doc', sheetId: SHEET, valueGen: 1, fontGen: 0, viewGen: 0 }),
        loadState: () => ({ completeness: 'complete' as const }),
        hasConditionalFormatting: () => true,
        grant: () => 1 << 20,
        releaseGrant: () => {},
        request: async cmd => {
          const result = await provide(cmd)
          return { commandKind: 'read', kind: 'complete', requestId: cmd.requestId, owner: cmd.owner as CacheIdentity, replyBytes: measurePayloadBytes(result, () => true)!, diagnostics: [],
            result: { kind: 'read', result: { op: 'conditionalFormatting', result } } } as WorkerReply<'read'>
        },
        nextRequestId: () => `req-${++reqId}`,
        release: () => {},
        admitEntry: () => () => {},
        changed: (_sheetId, _rects, reason) => { changed.push(reason ?? 'updated') },
      }
      const cache = new RemoteCfCache(host, 1 << 22)
      const handle = cache.acquire(SHEET, [rect(0, 9)], { ...st })
      await new Promise(r => setTimeout(r, 10))
      // The page shows no bar and no permanent diagnostic, and nothing is invalidated while pending.
      expect(handle.state.state).toBe('waiting')
      expect(handle.state.result).toBeUndefined()
      expect(handle.state.diagnostics).toEqual([])
      expect(handle.state.pendingSheetIds).toEqual([SHEET])
      expect(changed).toEqual([])
      // Arrival: the next notification wakes the waiting entry, which repaints exactly once with every bar.
      imported = true
      cache.wakeWaiting()
      await new Promise(r => setTimeout(r, 10))
      expect(handle.state.state).toBe('ready')
      expect(handle.state.result?.states.filter(s => s.bar !== undefined)).toHaveLength(10)
      expect(changed).toEqual(['updated'])
      cache.wakeWaiting()
      await new Promise(r => setTimeout(r, 10))
      expect(changed).toEqual(['updated'])
      handle.release(); cache.close()
      releaseSharedCfDefinitions(rules)
    })
  })

  it('no-CF and no-bar sheets do zero bar work', async () => {
    const cellCount = 1000
    const values = Array.from({ length: cellCount }, (_, i) => i)
    const rules: CfSheetRules = { diagnostics: [], rules: [scaleRule(0, cellCount - 1), dxfBaseRule(0, cellCount - 1)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    const run = (grant: number) => prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, grant), SHEET, columnSource(values, { stamp: () => st }), rules, createCfMemory({}), { budget: frozenBudget })
    await withBars(async () => {
      // A scale-only sheet: no state carries a `bar` key, at a wide grant and when its scales are omitted.
      const wide = ready(await run(1 << 24))
      expect(wide.states.filter(s => s.scale !== undefined)).toHaveLength(cellCount)
      expect(wide.states.every(s => !('bar' in s))).toBe(true)
      const tight = ready(await run(measurePayloadBytes(wide, () => true)! + CF_REPLY_ENVELOPE_BYTES - 1))
      expect(tight.states.every(s => !('bar' in s) && !('scale' in s))).toBe(true)
      expect(tight.diagnostics.filter(d => d.feature === 'cf-grant-required').map(d => d.message)).toEqual(['Conditional formatting scale fill omitted due to receiver grant limit.'])
      // Its byte estimate is the Stage 2 estimate: the scale visual only, no bar bytes.
      const refused = await run(4096)
      expect(refused.requiredBytes).toBe(templateBytes(rules, st) + cellCount * (perCellBase(rules) + CF_SCALE_CELL_BYTES) + CF_REPLY_ENVELOPE_BYTES)
      releaseSharedCfDefinitions(rules)
      // A no-CF sheet returns an empty ready reply.
      const none: CfSheetRules = { diagnostics: [], rules: [] }
      const empty = ready(await prepareCfRemote(cfRequest([rect(0, 9)], stamp(1, { cfGen: 0 }), 1 << 20), SHEET, columnSource(values), none, createCfMemory({}), { budget: frozenBudget }))
      expect(empty.states).toEqual([])
    })
  })
})
