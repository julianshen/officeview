import { readdirSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { CF_BAR_CELL_BYTES, CF_ICON_CELL_BYTES, createCfMemory, createCfWorkBudget, type CfWorkBudget } from '../src/xlsx/cf/limits'
import { CF_ICON_SET_NAMES } from '../src/xlsx/cf/model'
import type { CfDefinition, CfIconDescriptor, CfLease, CfLiteral, CfRule, CfSheetRules, CfStamp, CfValuePage, CfValueSource } from '../src/xlsx/cf/model'
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
import { VERIFIED_CAPABILITIES, injectCfBarLayout, injectCfIconSelections, withInjectedCfEvidence } from '../src/xlsx/cf/evidence'
import { createCfService } from '../src/xlsx/cf/service'
import { residentSource } from '../src/xlsx/cf/render'
import { descriptorOfIcon, prepareCfTextScope, prepareCfWireScope, wireIcon, type CfCopiedIconProbe } from '../src/xlsx/cf/text'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import type { XlsxSheet } from '../src/xlsx/types'

// Stage 4 Task 27: icon states on the wire (Ruling S4-Q). Every icon in this file comes from injected evidence: nothing
// Stage 4 is published, so production withholds every icon. The numbers are hypotheses run under stubs, never Excel facts.

const SHEET = 'icon-worker'
const REPO_ROOT = resolve(__dirname, '..')
const BLUE = 0xff0000ff, RED = 0xffff0000, BLACK = 0xff000000
const stamp = (gen = 1, extra: Partial<CfStamp> = {}): CfStamp => ({
  owner: 'owner', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...extra,
})
const rect = (r0: number, r1: number, c0 = 0, c1 = 0): SourceRect => ({ r0, r1, c0, c1 })
const frozenBudget = (): CfWorkBudget => createCfWorkBudget({ now: () => 0 })
const nc = (value: number): CfLiteral => ({ kind: 'number', value })

/** Every evaluation, scope and composition capability these icons can need (stubs only). */
const ICON_EVIDENCE = [
  'icon:families-base', 'icon:families-x14', 'icon:thresholds-num', 'icon:thresholds-percent', 'icon:thresholds-percentile-inc-exc',
  'icon:thresholds-gte-true', 'icon:thresholds-gte-false', 'icon:thresholds-mixed', 'icon:reverse', 'icon:reverse-custom',
  'icon:participation-types', 'icon:participation-dateserial', 'icon:composition-with-bar', 'icon:composition-with-scale',
]
const BAR_EVIDENCE = [
  'bar:bounds-min-max', 'bar:bounds-num', 'bar:bounds-percent', 'bar:bounds-percentile', 'bar:bounds-automin', 'bar:bounds-automax',
  'bar:lengths-default-10-90', 'bar:lengths-explicit-0-100', 'bar:render-solid', 'bar:render-border',
  'bar:axis-automatic', 'bar:axis-middle', 'bar:axis-none', 'bar:direction-ltr', 'bar:direction-rtl', 'bar:direction-context-ltr-sheet',
  'bar:negative-fill-color', 'bar:negative-fill-same-as-positive', 'bar:negative-border-color', 'bar:negative-border-same-as-positive',
  'bar:participation-types', 'bar:participation-dateserial',
  'bar:settings-precondition-probe', 'bar:settings-schema-defaults', 'bar:settings-paired-precedence',
  'bar:pairing-paired', 'bar:pairing-standalone', 'bar:pairing-base-only',
  'bar:colors-theme-tint', 'bar:colors-indexed', 'bar:bounds-positional', 'bar:bounds-out-of-domain', 'bar:degenerate-inverted',
  'bar:geometry-insets', 'bar:geometry-snapping', 'bar:geometry-narrow-clip',
]
const LAYOUT = {
  allPositive: { origin: 'span', lengthBasis: 'cell', settledBy: ['bar:lengths-explicit-0-100'] },
  mixedSign: { origin: 'axis', lengthBasis: 'side', settledBy: ['bar:axis-automatic', 'bar:axis-middle'] },
}
const withIcons = <T>(fn: () => Promise<T>, extra: readonly string[] = []): Promise<T> => {
  injectCfBarLayout(LAYOUT)
  return withInjectedCfEvidence([...ICON_EVIDENCE, ...BAR_EVIDENCE, ...extra], fn)
}
afterEach(() => { injectCfBarLayout(null); injectCfIconSelections(null) })

type IconDef = Extract<CfDefinition, { kind: 'iconSet' }>
type BarDef = Extract<CfDefinition, { kind: 'dataBar' }>
const th = (type: 'num' | 'percent' | 'percentile', value: number, gte = true) => ({ bound: { type, value }, gte })
function iconDef(overrides: Partial<IconDef> = {}): IconDef {
  return {
    kind: 'iconSet', dialect: 'base', formulas: [], set: '3Arrows', thresholds: [th('percent', 0), th('percent', 33), th('percent', 67)],
    reverse: false, showValue: true, percentAttr: true, sheetRightToLeft: false, ...overrides,
  }
}
function iconRule(firstRow = 0, lastRow = 999, id = 1, priority = 1, def: IconDef = iconDef()): CfRule {
  return { id, priority, documentOrder: priority - 1, ranges: [{ firstRow, lastRow, firstCol: 0, lastCol: 0 }], stopIfTrue: false, writes: new Set(['iconSet']), definition: def }
}
function barRule(firstRow = 0, lastRow = 999, id = 3, priority = 2): CfRule {
  const definition: BarDef = {
    kind: 'dataBar', dialect: 'base', formulas: [], bounds: [{ type: 'min' }, { type: 'max' }],
    fillArgb: BLUE, negativeFillArgb: RED, axisArgb: BLACK, minLength: 0, maxLength: 100, gradient: false, border: false,
    negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: true,
    axisPosition: 'automatic', direction: 'leftToRight', showValue: true, sheetRightToLeft: false,
  }
  return { id, priority, documentOrder: priority - 1, ranges: [{ firstRow, lastRow, firstCol: 0, lastCol: 0 }], stopIfTrue: false, writes: new Set(['dataBar']), definition }
}
function dxfBaseRule(firstRow = 0, lastRow = 999, id = 2, priority = 3): CfRule {
  return {
    id, priority, documentOrder: priority - 1, ranges: [{ firstRow, lastRow, firstCol: 0, lastCol: 0 }], stopIfTrue: false,
    writes: new Set(['font.bold']), dxf: { 'font.bold': true },
    definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
  }
}
const CUSTOM = [{ set: '3Flags' as const, iconId: 0 }, { set: 'NoIcons' as const, iconId: 0 }, { set: '5Boxes' as const, iconId: 4 }]

/** Column A values; the source resolves number formats (resident-like), so icons can participate (S4-G). */
function columnSource(
  values: number[],
  opts: { stamp?: () => CfStamp; visits?: { n: number }; gate?: () => Promise<void>; fault?: () => CfPageFault | undefined } = {},
): CfValueSource {
  const lits = { get length() { return values.length } }
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
        if (r.firstCol <= 0 && r.lastCol >= 0) { rows.push(row); vals.push(nc(values[row]!)); if (opts.visits) opts.visits.n++ }
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
const icons = (result: CfReadResult) => result.states.filter(s => s.icon !== undefined)
const bars = (result: CfReadResult) => result.states.filter(s => s.bar !== undefined)
/** The production decode of a wire icon: `set` back to its family name through the frozen table. */
const decode = (icon: NonNullable<CfCellStateWire['icon']>): CfIconDescriptor => ({ ...icon, set: CF_ICON_SET_NAMES[icon.set]! })
const templateBytes = (rules: CfSheetRules, st: CfStamp): number => measurePayloadBytes({
  sheetId: SHEET, stamp: { ...st }, kind: 'ready', states: [],
  definitions: rules.rules.map((rule, index) => ({ id: `rule-${typeof rule.id === 'number' ? rule.id : index}`, version: st.cfGen, properties: { ...(rule.dxf ?? {}) }, writes: [...rule.writes] })),
  definitionsVersion: st.cfGen, diagnostics: [],
}, () => true)!
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
const probe = (scope: unknown) => scope as CfCopiedIconProbe
/** A cache host that serves one recorded reply; returns the cache and the ledger it charges. */
function hostFor(clone: CfReadResult, replyBytes: number, opts: { grant?: number; keys?: string[] } = {}) {
  const ledger = new Map<string | object, number>()
  let reqId = 0
  const host: CfReadHost = {
    identity: () => ({ docId: 'doc', sheetId: SHEET, valueGen: 1, fontGen: 1, viewGen: 1 }),
    loadState: () => ({ completeness: 'complete' as const }),
    hasConditionalFormatting: () => true,
    grant: requestId => { ledger.set(`grant:${requestId}`, opts.grant ?? 1 << 22); return opts.grant ?? 1 << 22 },
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
    admitEntry: (key, bytes) => { opts.keys?.push(key); ledger.set(`entry:${key}`, bytes); return () => { ledger.delete(`entry:${key}`) } },
    changed: () => {},
  }
  return { host, ledger }
}

describe('CF Stage 4 Task 27 — icon wire under the receiver grant (S4-Q)', () => {
  it('icon payload fits the grant or is omitted', async () => {
    const cellCount = 1000
    const values = Array.from({ length: cellCount }, (_, i) => i)
    const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, cellCount - 1), barRule(0, cellCount - 1), dxfBaseRule(0, cellCount - 1)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    const run = (grant: number, memory = createCfMemory({}), visits?: { n: number }, with_: CfSheetRules = rules) =>
      prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, grant), SHEET, columnSource(values, { stamp: () => st, visits }), with_, memory, { budget: frozenBudget })
    await withIcons(async () => {
      // Priced before the payload is built: the advertised bound carries CF_ICON_CELL_BYTES per cell, and a tight grant
      // refuses without reading a single page.
      const pricedVisits = { n: 0 }
      const refused = await run(4096, createCfMemory({}), pricedVisits)
      expect(refused.kind).not.toBe('ready')
      expect(pricedVisits.n).toBe(0)
      const perCellVisuals = perCellBase(rules) + CF_BAR_CELL_BYTES + CF_ICON_CELL_BYTES
      expect(refused.requiredBytes).toBe(templateBytes(rules, st) + estimateCfPreparationBytes(cellCount, perCellVisuals, 0, 0, 0)! + CF_REPLY_ENVELOPE_BYTES)
      // The working hold covers the icons before any state exists: a pool that admits the base and the bars but not the
      // icons refuses before a page is read (the grant itself is wide).
      const baseBound = templateBytes(rules, st) + cellCount * perCellBase(rules) + CF_REPLY_ENVELOPE_BYTES
      const heldVisits = { n: 0 }
      const pressed = await run(1 << 24, createCfMemory({ limitBytes: baseBound + cellCount * CF_BAR_CELL_BYTES + (cellCount * CF_ICON_CELL_BYTES) / 2 }), heldVisits)
      expect(pressed.kind).not.toBe('ready')
      expect(heldVisits.n).toBe(0)
      expect(pressed.diagnostics.map(d => d.feature)).toEqual(['cf-budget-exceeded'])
      expect(pressed.requiredBytes).toBe(refused.requiredBytes)

      // Positive control: a wide grant carries one icon per scope cell, bars included.
      const wide = ready(await run(1 << 24))
      expect(wide.states).toHaveLength(cellCount)
      expect(icons(wide)).toHaveLength(cellCount)
      expect(bars(wide)).toHaveLength(cellCount)
      expect(wide.states[999]?.icon).toEqual({ ruleId: 1, set: CF_ICON_SET_NAMES.indexOf('3Arrows'), iconId: 2, bucket: 2, buckets: 3, showValue: true })
      expect(wide.states[1]?.properties['font.bold']).toBe(true)
      const exactGrant = measurePayloadBytes(wide, () => true)! + CF_REPLY_ENVELOPE_BYTES

      // At grant − 1: base cells and bars return, every icon field is omitted, and cf-grant-required rides along because it fits.
      const tight = ready(await run(exactGrant - 1))
      expect(tight.states).toHaveLength(cellCount)
      expect(tight.states.every(s => !('icon' in s))).toBe(true)
      expect(bars(tight)).toHaveLength(cellCount)
      expect(tight.states[1]?.properties['font.bold']).toBe(true)
      expect(tight.diagnostics.filter(d => d.feature === 'cf-grant-required').map(d => d.message)).toEqual(['Conditional formatting icon sets omitted due to receiver grant limit.'])
      expect(measurePayloadBytes(tight, () => true)! + CF_REPLY_ENVELOPE_BYTES).toBeLessThanOrEqual(exactGrant - 1)
      // A grant that admits the stripped base but not the diagnostics: base cells and bars return, no icon, no diagnostic.
      const noisy: CfSheetRules = { rules: rules.rules, diagnostics: Array.from({ length: 64 }, (_, i) => ({ kind: 'unsupported' as const, feature: `cf-noise-${i}` as const, message: 'x'.repeat(256) })) }
      const strippedBase = measurePayloadBytes({ ...tight, diagnostics: [] }, () => true)! + CF_REPLY_ENVELOPE_BYTES
      const strippedGrant = Math.max(strippedBase, baseBound)
      const noisyStripped = ready(await run(1 << 24, createCfMemory({}), undefined, noisy))
      const strippedWithDiagnostics = measurePayloadBytes({ ...noisyStripped, states: tight.states }, () => true)! + CF_REPLY_ENVELOPE_BYTES
      expect(strippedWithDiagnostics, 'precondition: the diagnostics do not fit').toBeGreaterThan(strippedGrant)
      const bare = ready(await run(strippedGrant, createCfMemory({}), undefined, noisy))
      expect(bare.states).toHaveLength(cellCount)
      expect(bare.states.every(s => !('icon' in s))).toBe(true)
      expect(bars(bare)).toHaveLength(cellCount)
      expect(bare.diagnostics).toEqual([])

      // At exactly the grant the reply carries every icon, identical to the wide reply.
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

  it('a sheet with icons and bars omits the icons first, then the bars (ordering pinned)', async () => {
    const cellCount = 1000
    const values = Array.from({ length: cellCount }, (_, i) => i)
    const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, cellCount - 1), barRule(0, cellCount - 1)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    const run = (grant: number) => prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, grant), SHEET, columnSource(values, { stamp: () => st }), rules, createCfMemory({}), { budget: frozenBudget })
    await withIcons(async () => {
      const wide = ready(await run(1 << 24))
      expect(icons(wide)).toHaveLength(cellCount)
      expect(bars(wide)).toHaveLength(cellCount)
      const full = measurePayloadBytes(wide, () => true)! + CF_REPLY_ENVELOPE_BYTES
      const grantMessages = (r: CfReadResult) => r.diagnostics.filter(d => d.feature === 'cf-grant-required').map(d => d.message)
      // One byte short: only the icons go, every bar stays.
      const noIcons = ready(await run(full - 1))
      expect(noIcons.states.every(s => !('icon' in s))).toBe(true)
      expect(bars(noIcons)).toHaveLength(cellCount)
      expect(grantMessages(noIcons)).toEqual(['Conditional formatting icon sets omitted due to receiver grant limit.'])
      // Short of the bars too: both go, the base cells return.
      const barsOnly = measurePayloadBytes({ ...noIcons, diagnostics: [] }, () => true)! + CF_REPLY_ENVELOPE_BYTES
      const neither = ready(await run(barsOnly - 1))
      expect(neither.states.every(s => !('icon' in s) && !('bar' in s))).toBe(true)
      expect(grantMessages(neither)).toEqual([
        'Conditional formatting icon sets omitted due to receiver grant limit.',
        'Conditional formatting data bars omitted due to receiver grant limit.',
      ])
      releaseSharedCfDefinitions(rules)
    })
  })

  it('Task 30 (carried from Task 27): with icons, bars and scales the grant omits icons, then bars, then scales', async () => {
    const cellCount = 1000
    const values = Array.from({ length: cellCount }, (_, i) => i)
    const scale: CfRule = {
      id: 5, priority: 3, documentOrder: 2, ranges: [{ firstRow: 0, lastRow: cellCount - 1, firstCol: 0, lastCol: 0 }], stopIfTrue: false, writes: new Set(['colorScale']),
      definition: { kind: 'colorScale', points: [{ bound: { type: 'min' }, argb: 0xfff8696b }, { bound: { type: 'max' }, argb: 0xff63be7b }], dialect: 'base', formulas: [] } as CfDefinition,
    }
    const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, cellCount - 1), barRule(0, cellCount - 1), scale] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    const run = (grant: number) => prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, grant), SHEET, columnSource(values, { stamp: () => st }), rules, createCfMemory({}), { budget: frozenBudget })
    const scales = (r: CfReadResult) => r.states.filter(x => x.scale !== undefined)
    const grantMessages = (r: CfReadResult) => r.diagnostics.filter(d => d.feature === 'cf-grant-required').map(d => d.message)
    const fits = (r: CfReadResult) => measurePayloadBytes({ ...r, diagnostics: [] }, () => true)! + CF_REPLY_ENVELOPE_BYTES
    await withIcons(async () => {
      const wide = ready(await run(1 << 24))
      expect([icons(wide).length, bars(wide).length, scales(wide).length]).toEqual([cellCount, cellCount, cellCount])
      const noIcons = ready(await run(fits(wide) - 1))
      expect([icons(noIcons).length, bars(noIcons).length, scales(noIcons).length]).toEqual([0, cellCount, cellCount])
      expect(grantMessages(noIcons)).toEqual(['Conditional formatting icon sets omitted due to receiver grant limit.'])
      const noBars = ready(await run(fits(noIcons) - 1))
      expect([icons(noBars).length, bars(noBars).length, scales(noBars).length]).toEqual([0, 0, cellCount])
      expect(grantMessages(noBars)).toEqual([
        'Conditional formatting icon sets omitted due to receiver grant limit.',
        'Conditional formatting data bars omitted due to receiver grant limit.',
      ])
      const baseOnly = ready(await run(fits(noBars) - 1))
      expect([icons(baseOnly).length, bars(baseOnly).length, scales(baseOnly).length]).toEqual([0, 0, 0])
      expect(baseOnly.states).toHaveLength(cellCount)
      releaseSharedCfDefinitions(rules)
    }, ['scale:thresholds-min-max', ...[...VERIFIED_CAPABILITIES].filter(id => id.startsWith('scale:'))])
  })

  it('the advertised requiredBytes of a refused icon read carries every icon (retry at exactly that grant)', async () => {
    const cellCount = 1000
    const values = Array.from({ length: cellCount }, (_, i) => i)
    const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, cellCount - 1)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    const run = (grant: number) => prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, grant), SHEET, columnSource(values, { stamp: () => st }), rules, createCfMemory({}), { budget: frozenBudget })
    await withIcons(async () => {
      const refused = await run(4096)
      expect(refused.kind).not.toBe('ready')
      const required = refused.requiredBytes!
      expect(required).toBeGreaterThan(4096)
      const retry = ready(await run(required))
      expect(icons(retry)).toHaveLength(cellCount)
      expect(retry.diagnostics.some(d => d.feature === 'cf-grant-required')).toBe(false)
      releaseSharedCfDefinitions(rules)
    })
  })

  it('clone copies of icon states are counted on both sides until adoption', async () => {
    for (const [cellCount, path] of [[2, 'split'], [200, 'reply']] as const) {
      const values = Array.from({ length: cellCount }, (_, i) => i)
      const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, cellCount - 1)] }
      const st = stamp(1, { cfGen: cfRulesVersion(rules) })
      await withIcons(async () => {
        const memory = createCfMemory({})
        const releases: Array<() => void> = []
        const produced = ready(await prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, 1 << 22), SHEET, columnSource(values, { stamp: () => st }), rules, memory,
          { budget: frozenBudget, retainOutgoing: release => releases.push(release) }))
        expect(icons(produced)).toHaveLength(cellCount)
        const wire = measurePayloadBytes(produced, () => true)!
        expect(releases).toHaveLength(1)
        const held = memory.currentBytes
        expect(held).toBeGreaterThanOrEqual(wire)
        const clone = structuredClone(produced)
        const replyBytes = measurePayloadBytes(clone, () => true)!
        expect(replyBytes).toBe(wire)
        const keys: string[] = []
        const { host, ledger } = hostFor(clone, replyBytes, { keys })
        const cache = new RemoteCfCache(host, 1 << 22)
        const handle = cache.acquire(SHEET, [rect(0, cellCount - 1)], { ...st })
        await new Promise(r => setTimeout(r, 10))
        expect(handle.state.state).toBe('ready')
        expect(handle.state.result?.states.filter(s => s.icon !== undefined)).toHaveLength(cellCount)
        const replyHeld = [...ledger.keys()].some(k => typeof k === 'object')
        const statesKey = [...ledger.keys()].find(k => typeof k === 'string' && k.startsWith('entry:cf-states:'))
        if (path === 'split') {
          expect(replyHeld, 'the reply charge ends at adoption').toBe(false)
          // The adopted states charge is the walked size of the states WITH their icons: one measured icon per cell.
          expect(ledger.get(statesKey!)).toBe(measurePayloadBytes({ ...clone, definitions: [] }, () => true))
          const stripped = { ...clone, definitions: [], states: clone.states.map(({ icon: _icon, ...rest }) => rest) }
          const perIcon = clone.states.reduce((n, s) => n + measureCfCellBytes(s) - measureCfCellBytes({ ...s, icon: undefined }), 0)
          expect(ledger.get(statesKey!)! - measurePayloadBytes(stripped, () => true)!).toBe(perIcon)
          expect(perIcon).toBe(cellCount * CF_ICON_CELL_BYTES)
        } else if (statesKey === undefined) {
          expect(replyHeld).toBe(true)
          expect([...ledger.entries()].find(([k]) => typeof k === 'object')?.[1]).toBe(replyBytes)
        } else {
          expect(replyHeld).toBe(false)
          expect(ledger.get(statesKey)).toBe(measurePayloadBytes({ ...clone, definitions: [] }, () => true))
        }
        handle.release()
        cache.close()
        expect(ledger.size).toBe(0)
        for (const release of releases) release()
        expect(memory.currentBytes).toBeLessThanOrEqual(held - wire)
        releaseSharedCfDefinitions(rules)
      })
    }
  })
})

/** A hand-built resident sheet whose column A holds `values`, with the given rules. */
const residentSheet = (values: readonly number[], rules: CfSheetRules): XlsxSheet => ({
  name: 'Icons', sheetId: SHEET, cols: [], merges: [], mergeRanges: [],
  rows: values.map((value, r) => ({ index: r, cells: [{ row: r, col: 0, ref: `A${r + 1}`, styleIndex: 0, value }] })),
  conditionalFormatting: rules,
}) as unknown as XlsxSheet

describe('CF Stage 4 Task 27 — wire decode, immutability and the S4-Q boundary', () => {
  it('wire icons equal the resident descriptor', async () => {
    const values = Array.from({ length: 12 }, (_, i) => i)
    const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, 11, 1, 1, iconDef({ custom: CUSTOM }))] }
    const sheet = residentSheet(values, rules)
    const st: CfStamp = { owner: 'fixture', sheetId: SHEET, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
    await withIcons(async () => {
      const residentMemory = createCfMemory({})
      const service = createCfService(rules, residentMemory, {})
      const lease = await service.prepareCf({ stamp: st, sourceRects: [{ firstRow: 0, lastRow: 11, firstCol: 0, lastCol: 0 }], signal: new AbortController().signal, grantBytes: 8 << 20 },
        residentSource(sheet, st, residentMemory, 'engine'))
      const wireMemory = createCfMemory({})
      const wire = ready(await prepareCfRemote(cfRequest([rect(0, 11)], st, 1 << 22), SHEET, residentSource(sheet, st, wireMemory, 'engine'), rules, wireMemory, { budget: frozenBudget }))
      const budget = budgetOf()
      const paged = await prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), [rect(0, 11)], { budget })
      expect(paged).toBeDefined()
      let compared = 0, noIcons = 0
      for (let row = 0; row <= 11; row++) {
        const at = lease.value.at(row, 0)
        const resident = at.kind === 'ready' ? at.value.visual?.iconSet : undefined
        const state = wire.states.find(s => s.row === row && s.col === 0)
        expect(state?.icon ? decode(state.icon) : undefined, `wire ${row}`).toEqual(resident)
        expect(probe(paged!).copiedIconAt?.(row, 0), `paged ${row}`).toEqual(resident)
        if (resident) { compared++; if (resident.set === 'NoIcons') noIcons++ }
      }
      // Positive controls: every cell resolves, and the NoIcons bucket travels as set 20 (never as an omitted field).
      expect(compared).toBe(12)
      expect(noIcons).toBeGreaterThanOrEqual(3)
      expect(wire.states.find(s => s.icon?.set === CF_ICON_SET_NAMES.indexOf('NoIcons'))?.icon?.iconId).toBe(0)
      paged!.dispose()
      expect(budget.held).toBe(0)
      lease.release(); service.dispose()
      releaseSharedCfDefinitions(rules)
    })
  })

  it('the main thread consumes no per-rule icon data (S4-Q, owner-accepted)', async () => {
    const values = Array.from({ length: 12 }, (_, i) => i)
    const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, 11, 1, 1, iconDef({ custom: CUSTOM }))] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    await withIcons(async () => {
      const produced = ready(await prepareCfRemote(cfRequest([rect(0, 11)], st, 1 << 22), SHEET, columnSource(values, { stamp: () => st }), rules, createCfMemory({}), { budget: frozenBudget }))
      expect(icons(produced)).toHaveLength(12)
      const clone = structuredClone(produced)
      const replyBytes = measurePayloadBytes(clone, () => true)!
      const { host, ledger } = hostFor(clone, replyBytes)
      const cache = new RemoteCfCache(host, 1 << 22)
      const handle = cache.acquire(SHEET, [rect(0, 11)], { ...st })
      await new Promise(r => setTimeout(r, 10))
      expect(handle.state.state).toBe('ready')
      const adopted = handle.state.result!
      const WIRE_KEYS = ['bucket', 'buckets', 'iconId', 'ruleId', 'set', 'showValue']
      for (const state of adopted.states) {
        expect(Object.keys(state.icon ?? {}).sort()).toEqual(state.icon ? WIRE_KEYS : [])
        for (const v of Object.values(state.icon ?? {})) expect(['number', 'boolean']).toContain(typeof v)
      }
      // The adopted entry carries no icon definition, no thresholds, no cfIcon list and no shared icon table.
      const text = JSON.stringify({ ...adopted, states: adopted.states.map(({ icon: _icon, ...rest }) => rest) })
      for (const needle of ['thresholds', 'cfIcon', 'custom', 'percentile', '3Flags', '5Boxes', 'NoIcons']) expect(text, needle).not.toContain(needle)
      for (const def of adopted.definitions) {
        expect(Object.keys(def).sort()).toEqual(['id', 'properties', 'version', 'writes'])
        expect(def.properties).toEqual({})
      }
      // Every CopiedState of the paged scope is the plain wire descriptor: the six wire fields, no more.
      const budget = budgetOf()
      const paged = (await prepareCfWireScope(async () => ({ result: structuredClone(adopted), release() {} }), [rect(0, 11)], { budget }))!
      for (let row = 0; row <= 11; row++) {
        const copied = probe(paged).copiedIconAt?.(row, 0)
        expect(copied, `row ${row}`).toBeDefined()
        expect(Object.keys(copied!).sort()).toEqual(WIRE_KEYS)
      }
      paged.dispose()
      handle.release(); cache.close()
      expect(ledger.size).toBe(0)
      releaseSharedCfDefinitions(rules)
    })
    // Source grep: no consumer on the main-thread side reads a rule's icon definition.
    const files = ['src/xlsx/cf/text.ts', 'src/worker/viewerRemoteCf.ts', 'src/xlsx/view/renderViewport.ts']
    const walk = (dir: string): string[] => readdirSync(resolve(REPO_ROOT, dir), { recursive: true, encoding: 'utf8' })
      .filter(f => /\.tsx?$/.test(f)).map(f => `${dir}/${f}`)
    for (const file of [...files, ...walk('src/components')]) {
      const source = readFileSync(resolve(REPO_ROOT, file), 'utf8')
      expect(source, `${file} reads definition.custom`).not.toMatch(/definition\.custom/)
      expect(source, `${file} reads definition.thresholds`).not.toMatch(/definition\.thresholds/)
      expect(source, `${file} reads kind === 'iconSet'`).not.toMatch(/kind\s*===\s*'iconSet'/)
    }
  })

  it('CF_ICON_SET_NAMES is an immutable wire table', async () => {
    // A reorder is a breaking wire change: the numbers on the wire index this table.
    expect([...CF_ICON_SET_NAMES]).toEqual([
      '3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2',
      '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating',
      '5Quarters', '3Triangles', '3Stars', '5Boxes', 'NoIcons',
    ])
    expect(CF_ICON_SET_NAMES).toHaveLength(21)
    expect(Object.isFrozen(CF_ICON_SET_NAMES)).toBe(true)
    expect(() => (CF_ICON_SET_NAMES as unknown as string[]).push('X')).toThrow()
    expect(() => { (CF_ICON_SET_NAMES as unknown as string[])[0] = 'X' }).toThrow()
    expect((CF_ICON_SET_NAMES as unknown as string[]).length).toBe(21)
    // Encode/decode round trip through the table for every entry.
    CF_ICON_SET_NAMES.forEach((set, index) => {
      const wire = wireIcon({ ruleId: 1, set, iconId: 0, bucket: 0, buckets: 3, showValue: true })
      expect(wire.set).toBe(index)
      expect(descriptorOfIcon(wire)?.set).toBe(set)
    })
    // Out-of-range, negative and non-integer sets decode to no icon, never to a guessed set.
    const good = { ruleId: 1, set: 0, iconId: 0, bucket: 0, buckets: 3, showValue: true }
    for (const set of [-1, 21, 2.5, 1e9, Number.NaN, Number.POSITIVE_INFINITY]) expect(descriptorOfIcon({ ...good, set }), `set ${set}`).toBeUndefined()
    expect(descriptorOfIcon({ ...good, set: 0 })).toEqual({ ...good, set: '3Arrows' })
    expect(descriptorOfIcon({ ...good, set: 20 })).toEqual({ ...good, set: 'NoIcons' })
    // Malformed fields fail closed the way malformed bar fields do: a bad number, a bad flag, a bucket past the count.
    for (const bad of [{ iconId: -1 }, { iconId: 1.5 }, { bucket: 3 }, { bucket: -1 }, { buckets: 0 }, { buckets: 6 }, { ruleId: -1 }, { ruleId: Number.NaN }, { showValue: 'yes' }, { set: '0' }, { iconId: Number.NaN }]) {
      expect(descriptorOfIcon({ ...good, ...bad } as never), JSON.stringify(bad)).toBeUndefined()
    }
    // The paged scope decodes a hostile reply to no icon on exactly those cells, leaving every base property.
    const values = [0, 5, 10]
    const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, 2), dxfBaseRule(0, 2)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    await withIcons(async () => {
      const wire = ready(await prepareCfRemote(cfRequest([rect(0, 2)], st, 1 << 20), SHEET, columnSource(values, { stamp: () => st }), rules, createCfMemory({}), { budget: frozenBudget }))
      expect(icons(wire)).toHaveLength(3)
      const sets = [-1, 21, 2.5]
      const hostile: CfReadResult = { ...structuredClone(wire), states: wire.states.map((s, i) => ({ ...s, icon: { ...s.icon!, set: sets[i]! } })) }
      const budget = budgetOf()
      const paged = (await prepareCfWireScope(async () => ({ result: hostile, release() {} }), [rect(0, 2)], { budget }))!
      for (let row = 0; row <= 2; row++) expect(probe(paged).copiedIconAt?.(row, 0), `row ${row}`).toBeUndefined()
      expect(paged.effectiveStyle(1, 0, undefined)).toBeDefined()
      paged.dispose()
      expect(budget.held).toBe(0)
      releaseSharedCfDefinitions(rules)
    })
  })
})

describe('CF Stage 4 Task 27 — generations, pending pages and no-icon sheets', () => {
  it('max edit re-buckets every icon of the scope in one generation', async () => {
    const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, 3)] }
    const version = cfRulesVersion(rules)
    await withIcons(async () => {
      // Resident shell at g1: A1:A4 = 0, 60, 100, 10 (max 100; 60 sits in bucket 1 under percent 33/67).
      const sheet = residentSheet([0, 60, 100, 10], rules)
      const source = createObjectWorkbookSource({ sheets: [sheet] }, 'icon-live')
      const store = source.store(SHEET)
      const budget = budgetOf()
      const prepare = () => prepareCfTextScope(sheet, { rects: [rect(0, 3)], signal: new AbortController().signal, budget, store })
      const r1 = (await prepare())!
      expect(probe(r1).copiedIconAt?.(1, 0)?.bucket).toBe(1)
      expect(probe(r1).copiedIconAt?.(2, 0)?.bucket).toBe(2)

      // Worker shell at g1, decoded on the paged main thread with generation liveness.
      let liveGen = 1
      let gate: Promise<void> | undefined
      const workerValues = [0, 60, 100, 10]
      const src = columnSource(workerValues, { stamp: () => stamp(liveGen, { cfGen: version }), gate: () => gate ?? Promise.resolve() })
      const read = (gen: number) => prepareCfRemote(cfRequest([rect(0, 3)], stamp(gen, { cfGen: version }), 1 << 20), SHEET, src, rules, createCfMemory({}), { budget: frozenBudget })
      const g1Wire = ready(await read(1))
      const w1 = (await prepareCfWireScope(async () => ({ result: g1Wire, release() {} }), [rect(0, 3)], { budget, current: () => liveGen === 1 }))!
      expect(probe(w1).copiedIconAt?.(1, 0)?.bucket).toBe(1)

      // Commit max 50 at g2 on both shells while the g2 preparation is pending: the 100 becomes 50, so 60 is the new max.
      sheet.rows[2]!.cells[0]!.value = 50
      store.noteValuesChanged?.()
      workerValues[2] = 50
      liveGen = 2
      let open!: () => void
      gate = new Promise(r => { open = r })
      const pending = read(2)
      const residentPending = prepare()
      for (let row = 0; row <= 3; row++) {
        expect(probe(r1).copiedIconAt?.(row, 0), `resident g1 icon at g2, row ${row}`).toBeUndefined()
        expect(probe(w1).copiedIconAt?.(row, 0), `worker g1 icon at g2, row ${row}`).toBeUndefined()
      }
      open()
      const g2Wire = ready(await pending)
      const r2 = (await residentPending)!
      const w2 = (await prepareCfWireScope(async () => ({ result: g2Wire, release() {} }), [rect(0, 3)], { budget, current: () => liveGen === 2 }))!
      // Every scope cell gets its g2 bucket: thresholds are now [0, 19.8, 40.2] over 0..60, so 60 moves to bucket 2.
      for (const [row, bucket] of [[0, 0], [1, 2], [2, 2], [3, 0]] as const) {
        expect(probe(r2).copiedIconAt?.(row, 0)?.bucket, `resident row ${row}`).toBe(bucket)
        expect(probe(w2).copiedIconAt?.(row, 0)?.bucket, `worker row ${row}`).toBe(bucket)
      }
      for (const scope of [r1, w1, r2, w2]) scope.dispose()
      expect(budget.held).toBe(0)
      releaseSharedCfDefinitions(rules)
    })
  })

  it('unimported scope pages withhold icons then repaint once', async () => {
    const rules: CfSheetRules = { diagnostics: [], rules: [iconRule(0, 9)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    await withIcons(async () => {
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
      // No icon, no inset (no scope states at all), no permanent diagnostic, and nothing invalidated while pending.
      expect(handle.state.state).toBe('waiting')
      expect(handle.state.result).toBeUndefined()
      expect(handle.state.diagnostics).toEqual([])
      expect(handle.state.pendingSheetIds).toEqual([SHEET])
      expect(changed).toEqual([])
      imported = true
      cache.wakeWaiting()
      await new Promise(r => setTimeout(r, 10))
      expect(handle.state.state).toBe('ready')
      expect(handle.state.result?.states.filter(s => s.icon !== undefined)).toHaveLength(10)
      expect(changed).toEqual(['updated'])
      cache.wakeWaiting()
      await new Promise(r => setTimeout(r, 10))
      expect(changed).toEqual(['updated'])
      handle.release(); cache.close()
      releaseSharedCfDefinitions(rules)
    })
  })

  it('no-icon sheets do zero icon work', async () => {
    const cellCount = 1000
    const values = Array.from({ length: cellCount }, (_, i) => i)
    const rules: CfSheetRules = { diagnostics: [], rules: [barRule(0, cellCount - 1, 1, 1), dxfBaseRule(0, cellCount - 1)] }
    const st = stamp(1, { cfGen: cfRulesVersion(rules) })
    const run = (grant: number) => prepareCfRemote(cfRequest([rect(0, cellCount - 1)], st, grant), SHEET, columnSource(values, { stamp: () => st }), rules, createCfMemory({}), { budget: frozenBudget })
    await withIcons(async () => {
      // A bar-only sheet: no state carries an `icon` key, at a wide grant and when its bars are omitted.
      const wide = ready(await run(1 << 24))
      expect(bars(wide)).toHaveLength(cellCount)
      expect(wide.states.every(s => !('icon' in s))).toBe(true)
      const tight = ready(await run(measurePayloadBytes(wide, () => true)! + CF_REPLY_ENVELOPE_BYTES - 1))
      expect(tight.states.every(s => !('icon' in s) && !('bar' in s))).toBe(true)
      expect(tight.diagnostics.filter(d => d.feature === 'cf-grant-required').map(d => d.message)).toEqual(['Conditional formatting data bars omitted due to receiver grant limit.'])
      // Its byte estimate is the Stage 3 estimate: the bar visual only, no icon bytes.
      const refused = await run(4096)
      expect(refused.requiredBytes).toBe(templateBytes(rules, st) + cellCount * (perCellBase(rules) + CF_BAR_CELL_BYTES) + CF_REPLY_ENVELOPE_BYTES)
      releaseSharedCfDefinitions(rules)
    })
  })
})
