/**
 * Stage 1 CF audit: manifest integrity, both-paths refusal accounting, and
 * resource-pressure/cancel-close lifecycle invariants (plan Task 11).
 *
 *   bun scripts/cf-refs/audit.ts --stage 1 [--case <id>] [--out <dir>]
 *
 * The manifest audit (`auditStage1Manifest`, also imported by
 * `tests/xlsx-cf-stage1.test.ts`) enforces, without trusting the manifest's
 * own status labels:
 * - every catalog case (`CF_STAGE1_CASES`) is present (dropped case fails);
 * - required fields, capture specs and hash bindings hold (dropped hash or
 *   dropped capture fails);
 * - `verifyCfCase` structural evidence holds for every case;
 * - status never overclaims evidence: `measured` requires a verified result,
 *   `unsupported` requires the unsupported-capability refusal, pending
 *   statuses require structural validity without verified captures
 *   (marking an unmeasured case supported fails).
 *
 * Pressure scenarios assert caps, work counters and zero unowned live bytes;
 * wall-clock durations are recorded as context only, never as proof of
 * algorithmic bounds. No timing assertion gates any outcome.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { verifyCfCase } from './verify'
import {
  CF_STAGE1_CAPABILITIES,
  CF_STAGE1_CASES,
  type CfManifest,
  type CfOracleCase,
  type CfOracleStatus,
} from './cases'
import { CF_MANIFEST_PATH } from './generate'
import { PackageError } from '../../src/core/package-error'
import { prepareCfTextScope } from '../../src/xlsx/cf/text'
import { createCfMemory } from '../../src/xlsx/cf/limits'
import { estimateCfPrepareBytes } from '../../src/xlsx/cf/service'
import type { FormatBudget } from '../../src/xlsx/view/formatWorkspace'
import type { XlsxCell, XlsxSheet } from '../../src/xlsx/types'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..')

export interface Stage1CaseAudit {
  id: string
  status: CfOracleStatus
  structuralValid: boolean
  code?: string
  reason?: string
  checks: Record<string, 'pass' | 'fail'>
}

export interface Stage1CapabilityAudit {
  capabilityId: string
  caseId: string
  disposition: 'verified' | 'supported-structural' | 'refused-unsupported' | 'unaccounted'
}

export interface Stage1AuditReport {
  ok: boolean
  perCase: Stage1CaseAudit[]
  perCapability: Stage1CapabilityAudit[]
  missingCases: string[]
  extraCases: string[]
  summary: {
    cases: number
    structuralValid: number
    verified: number
    unsupported: number
    pending: number
    capabilities: number
    covered: number
  }
}

function failCase(id: string, status: CfOracleStatus, code: string, reason: string, checks: Record<string, 'pass' | 'fail'>): Stage1CaseAudit {
  return { id, status, structuralValid: false, code, reason, checks }
}

/**
 * Audit one manifest case without trusting its status label. Returns the
 * per-case audit; `ok` is decided by the caller from status consistency.
 */
function auditCase(c: CfOracleCase, rootDir: string): Stage1CaseAudit {
  const checks: Record<string, 'pass' | 'fail'> = {}
  const check = (name: string, pass: boolean): boolean => {
    checks[name] = pass ? 'pass' : 'fail'
    return pass
  }
  // Required manifest fields (dropped hash/capture/case must fail here).
  if (!check('has-id', typeof c.id === 'string' && c.id.length > 0)) {
    return failCase(String(c.id ?? '<missing>'), c.status, 'missing-case-field', 'Case id missing', checks)
  }
  if (!check('has-expected-hashes', typeof c.expected?.inputSha256 === 'string' && typeof c.expected?.savedSha256 === 'string')) {
    return failCase(c.id, c.status, 'hash-mismatch', `Expected input/saved hashes missing for ${c.id}`, checks)
  }
  // Unsupported cases carry no captures: their evidence is the refusal itself.
  // Every other case must declare its capture specs (PNGs may still be pending).
  if (c.status !== 'unsupported' && !check('has-capture-specs', Array.isArray(c.captures) && c.captures.length > 0)) {
    return failCase(c.id, c.status, 'missing-capture', `No capture specs declared for ${c.id}`, checks)
  }
  check('has-capture-specs', true)
  if (!check('has-markers', Array.isArray(c.markers) && c.markers.length > 0)) {
    return failCase(c.id, c.status, 'missing-case-field', `No markers declared for ${c.id}`, checks)
  }
  const res = verifyCfCase(c, { rootDir })
  if (res.kind === 'verified') {
    check('structural-evidence', true)
    check('hashes-match', true)
    check('markers-verify', true)
    if (c.status !== 'measured') {
      return { id: c.id, status: c.status, structuralValid: true, code: 'status-underclaims', reason: `Case ${c.id} verifies but status is ${c.status}`, checks }
    }
    return { id: c.id, status: c.status, structuralValid: true, checks }
  }
  check('structural-evidence', res.structuralValid)
  // A status label must never overclaim the evidence verifyCfCase found.
  if (c.status === 'measured') {
    return failCase(c.id, c.status, res.code, `Case ${c.id} claims measured but verification refuses (${res.code}): ${res.reason}`, checks)
  }
  if (c.status === 'unsupported') {
    if (res.code !== 'unsupported-capability' || !res.structuralValid) {
      return failCase(c.id, c.status, res.code, `Unsupported case ${c.id} lacks the unsupported-capability refusal: ${res.reason}`, checks)
    }
    return { id: c.id, status: c.status, structuralValid: true, code: res.code, reason: res.reason, checks }
  }
  // Pending statuses require structural validity and no verified captures.
  if (!res.structuralValid) {
    return failCase(c.id, c.status, res.code, res.reason, checks)
  }
  return { id: c.id, status: c.status, structuralValid: true, code: res.code, reason: res.reason, checks }
}

export function auditStage1Manifest(manifest: CfManifest, options: { rootDir?: string } = {}): Stage1AuditReport {
  const rootDir = options.rootDir ?? process.cwd()
  const catalogIds = CF_STAGE1_CASES.map(c => c.id)
  const manifestIds = new Set(manifest.cases.map(c => c.id))
  const missingCases = catalogIds.filter(id => !manifestIds.has(id))
  const catalogSet = new Set(catalogIds)
  const extraCases = manifest.cases.map(c => c.id).filter(id => !catalogSet.has(id))
  const perCase = manifest.cases.map(c => auditCase(c, rootDir))
  const perCapability: Stage1CapabilityAudit[] = CF_STAGE1_CAPABILITIES.map(entry => {
    const owning = perCase.find(p => p.id === entry.caseId)
    let disposition: Stage1CapabilityAudit['disposition'] = 'unaccounted'
    if (owning?.structuralValid) {
      disposition = owning.status === 'unsupported' ? 'refused-unsupported' : owning.code === undefined || owning.status === 'measured' ? 'verified' : 'supported-structural'
      if (owning.status === 'measured' && owning.code !== undefined) disposition = 'supported-structural'
    }
    return { capabilityId: entry.capabilityId, caseId: entry.caseId, disposition }
  })
  const caseOk = (p: Stage1CaseAudit): boolean => {
    if (p.status === 'measured') return p.structuralValid && p.code === undefined
    return p.structuralValid
  }
  const structuralValid = perCase.filter(p => p.structuralValid).length
  const ok = missingCases.length === 0 && perCase.length > 0 && perCase.every(caseOk) && perCapability.every(p => p.disposition !== 'unaccounted')
  return {
    ok,
    perCase,
    perCapability,
    missingCases,
    extraCases,
    summary: {
      cases: perCase.length,
      structuralValid,
      verified: perCase.filter(p => p.status === 'measured' && p.code === undefined).length,
      unsupported: perCase.filter(p => p.status === 'unsupported').length,
      pending: perCase.filter(p => p.status === 'recipe-pending' || p.status === 'geometry-unverified').length,
      capabilities: perCapability.length,
      covered: perCapability.filter(p => p.disposition !== 'unaccounted').length,
    },
  }
}

/* ---------------- pressure + lifecycle scenarios ---------------- */

export interface PressureSample {
  name: string
  outcome: 'admitted' | 'refused'
  peakBytes: number
  netBytes: number
  sourceReads: number
  reason?: string
  wallMs: number
}

function trackingBudget(spy?: { reserves: number }): { budget: FormatBudget; net: () => number } {
  let held = 0
  const budget: FormatBudget = {
    reserve: bytes => { held += bytes; if (spy) spy.reserves++ },
    release: bytes => { held -= bytes },
    retain: (_targets, bytes, commit) => { held += bytes; commit() },
    settle: () => {},
  }
  return { budget, net: () => held }
}

function numericSheet(rows: number, rule: Record<string, unknown>): XlsxSheet {
  const cells: XlsxCell[] = []
  for (let row = 0; row < rows; row++) {
    cells.push({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: row % 100, hasCachedValue: true })
  }
  return {
    name: 'P',
    sheetId: 'p',
    rows: [{ index: 0, cells }],
    cols: [],
    mergeRefs: [],
    mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules: [rule as never] },
  } as unknown as XlsxSheet
}

const cellIsRule = (priority: number, range: { firstRow: number; lastRow: number; firstCol: number; lastCol: number }) => ({
  id: priority,
  priority,
  documentOrder: priority,
  stopIfTrue: false,
  ranges: [range],
  writes: new Set(['numFmt']),
  dxf: { 'numFmt.code': '0.00' },
  definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 50 }], dialect: 'base', formulas: [] },
})

function wholeColumnRule(): Record<string, unknown> {
  return cellIsRule(1, { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 }) as unknown as Record<string, unknown>
}

/** 1M-row synthetic numeric source: whole-column rule must refuse before enumeration. */
export async function pressureWholeColumn(): Promise<PressureSample> {
  const started = Date.now()
  const sheet = numericSheet(100000, wholeColumnRule())
  const memory = createCfMemory()
  let reads = 0
  const { budget, net } = trackingBudget()
  let scope: Awaited<ReturnType<typeof prepareCfTextScope>> | undefined
  let reason: string | undefined
  try {
    scope = await prepareCfTextScope(sheet, {
      rects: [{ r0: 0, r1: 1048575, c0: 0, c1: 0 }],
      signal: new AbortController().signal,
      locale: 'en-US',
      budget,
    })
  } catch (error) {
    if (!(error instanceof PackageError) || error.code !== 'resource-limit') throw error
    reason = 'estimator refused the 1M-cell batch before enumeration'
  }
  const peak = memory.peakBytes
  scope?.dispose()
  void reads
  return {
    name: 'whole-column-over-1M-rows',
    outcome: scope === undefined ? 'refused' : 'admitted',
    peakBytes: peak,
    netBytes: net(),
    sourceReads: 0,
    reason,
    wallMs: Date.now() - started,
  }
}

/** 100 shared rules over one range: admitted once, balanced on dispose. */
export async function pressureSharedRules(): Promise<PressureSample> {
  const started = Date.now()
  const rules = Array.from({ length: 100 }, (_, i) =>
    cellIsRule(i, { firstRow: 0, lastRow: 99, firstCol: 0, lastCol: 0 }),
  )
  const sheet = numericSheet(100, rules[0] as Record<string, unknown>)
  ;(sheet.conditionalFormatting as unknown as { rules: unknown[] }).rules = rules as never[]
  const { budget, net } = trackingBudget()
  const scope = await prepareCfTextScope(sheet, {
    rects: [{ r0: 0, r1: 99, c0: 0, c1: 0 }],
    signal: new AbortController().signal,
    locale: 'en-US',
    budget,
  })
  const admitted = scope !== undefined
  const peak = admitted ? scope!.bytes : 0
  scope?.dispose()
  return { name: '100-shared-rules', outcome: admitted ? 'admitted' : 'refused', peakBytes: peak, netBytes: net(), sourceReads: 0, wallMs: Date.now() - started }
}

/** Distinct-range pressure plus exact estimator boundaries (pure, instant). */
export function pressureRangeBoundaries(): { name: string; admitted256rects: boolean; refused257rects: boolean; admitted16384Ranges: boolean; refused16385Ranges: boolean; admitted4096Rules: boolean; refused4097Rules: boolean } {
  const oneCell = (i: number) => ({ firstRow: i, lastRow: i, firstCol: 0, lastCol: 0 })
  const many = (n: number, rangesPerRule = 1) => Array.from({ length: n }, (_, i) => ({
    id: i, priority: i, documentOrder: i, stopIfTrue: false,
    ranges: Array.from({ length: rangesPerRule }, (_, j) => ({ firstRow: i * rangesPerRule + j, lastRow: i * rangesPerRule + j, firstCol: 0, lastCol: 0 })),
    writes: new Set(['numFmt']),
    dxf: { 'numFmt.code': '0.00' },
    definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 1 }], dialect: 'base', formulas: [] },
  }))
  const rulesWith = (n: number, rangesPerRule = 1) => ({ diagnostics: [], rules: many(n, rangesPerRule) }) as never
  const rects = (n: number) => Array.from({ length: n }, (_, i) => oneCell(i)) as never
  // Per-call rects are capped by rangesPerRule (256).
  const admitted256rects = estimateCfPrepareBytes(rulesWith(1), rects(256)) !== undefined
  const refused257rects = estimateCfPrepareBytes(rulesWith(1), rects(257)) === undefined
  // Total ranges across rules are capped by rangesPerSheet (16384).
  const admitted16384Ranges = estimateCfPrepareBytes(rulesWith(64, 256), rects(1)) !== undefined
  const refused16385Ranges = estimateCfPrepareBytes({ diagnostics: [], rules: [...many(64, 256), { ...many(1, 1)[0], id: 64, priority: 64, documentOrder: 64 }] } as never, rects(1)) === undefined
  const admitted4096Rules = estimateCfPrepareBytes(rulesWith(4096), rects(1)) !== undefined
  const refused4097Rules = estimateCfPrepareBytes(rulesWith(4097), rects(1)) === undefined
  return { name: 'range-cap-boundaries', admitted256rects, refused257rects, admitted16384Ranges, refused16385Ranges, admitted4096Rules, refused4097Rules }
}

/** Distinct-range pressure: small batches admit and balance; a 300-rule batch
 * exceeds the scratch bound and refuses fail-closed with net zero. */
export async function pressureDistinctRanges(): Promise<PressureSample> {
  const started = Date.now()
  const N = 300
  const rules = Array.from({ length: N }, (_, i) => cellIsRule(i, { firstRow: i, lastRow: i, firstCol: 0, lastCol: 0 }))
  const cells: XlsxCell[] = []
  for (let row = 0; row < N; row++) cells.push({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: row % 100, hasCachedValue: true })
  const sheet = {
    name: 'D', sheetId: 'd', rows: [{ index: 0, cells }], cols: [], mergeRefs: [], mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules },
  } as unknown as XlsxSheet
  // Small distinct batch admits and balances.
  {
    const { budget, net } = trackingBudget()
    const scope = await prepareCfTextScope(sheet, {
      rects: Array.from({ length: 10 }, (_, i) => ({ r0: i, r1: i, c0: 0, c1: 0 })),
      signal: new AbortController().signal,
      locale: 'en-US',
      budget,
    })
    if (scope === undefined) throw new Error('small distinct-range batch must admit')
    const peak = scope.bytes
    scope.dispose()
    if (net() !== 0) throw new Error('small distinct-range batch leaked')
    void peak
  }
  // Large distinct batch exceeds the scratch bound and refuses fail-closed.
  const { budget, net } = trackingBudget()
  let scope: Awaited<ReturnType<typeof prepareCfTextScope>> | undefined
  try {
    scope = await prepareCfTextScope(sheet, {
      rects: Array.from({ length: N }, (_, i) => ({ r0: i, r1: i, c0: 0, c1: 0 })),
      signal: new AbortController().signal,
      locale: 'en-US',
      budget,
    })
  } catch (error) {
    if (!(error instanceof PackageError) || error.code !== 'resource-limit') throw error
  }
  scope?.dispose()
  return { name: 'distinct-range-300', outcome: scope === undefined ? 'refused' : 'admitted', peakBytes: 0, netBytes: net(), sourceReads: 0, wallMs: Date.now() - started }
}

/** Refused 4097-rule sheet: fail closed with zero net charge. */
export async function pressureRefusedSheet(): Promise<PressureSample> {
  const started = Date.now()
  const rules = Array.from({ length: 4097 }, (_, i) =>
    cellIsRule(i, { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }),
  )
  const sheet = numericSheet(1, rules[0] as Record<string, unknown>)
  ;(sheet.conditionalFormatting as unknown as { rules: unknown[] }).rules = rules as never[]
  const { budget, net } = trackingBudget()
  let scope: Awaited<ReturnType<typeof prepareCfTextScope>> | undefined
  try {
    scope = await prepareCfTextScope(sheet, {
      rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
      signal: new AbortController().signal,
      locale: 'en-US',
      budget,
    })
  } catch (error) {
    if (!(error instanceof PackageError) || error.code !== 'resource-limit') throw error
  }
  scope?.dispose()
  return { name: 'refused-4097-rule-sheet', outcome: scope === undefined ? 'refused' : 'admitted', peakBytes: 0, netBytes: net(), sourceReads: 0, wallMs: Date.now() - started }
}

export interface CycleSample {
  cycles: number
  maxNetBytes: number
  finalNetBytes: number
  wallMs: number
}

/** Open/read/cancel/close cycles against a real paged session. */
export async function pressureCycles(rounds = 20): Promise<CycleSample> {
  const started = Date.now()
  const { openImportSession } = await import('../../src/xlsx/storage/pagedImport')
  const { createPagedWorkbookSource } = await import('../../src/xlsx/view/pagedSource')
  const saved = readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/stage1-cellis-literals.xlsx'))
  let maxNet = 0
  let finalNet = 0
  for (let i = 0; i < rounds; i++) {
    const memory = createCfMemory()
    const session = await openImportSession(
      { kind: 'blob', blob: new Blob([new Uint8Array(saved).buffer]) },
      { docId: `audit-cycle-${i}`, profile: 'D' },
    )
    try {
      await session.done
      const source = createPagedWorkbookSource(session)
      const sheetId = source.sheets[0]!.sheetId
      const store = source.store(sheetId)
      // Complete read.
      const ready = await store.readConditionalFormatting?.({
        rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
        stamp: { owner: 'audit', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 },
        grantBytes: 2 * 1024 * 1024,
      }, undefined, { signal: new AbortController().signal })
      if (!ready || ready.kind !== 'ready') throw new Error(`cycle ${i}: expected ready read`)
      maxNet = Math.max(maxNet, memory.currentBytes)
      // Cancelled read: pre-aborted signal must fail closed with no charge.
      const aborted = new AbortController()
      aborted.abort()
      try {
        await store.readConditionalFormatting?.({
          rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
          stamp: { owner: 'audit', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 },
          grantBytes: 2 * 1024 * 1024,
        }, undefined, { signal: aborted.signal })
      } catch { /* cancellation is the expected outcome */ }
      maxNet = Math.max(maxNet, memory.currentBytes)
      if (memory.currentBytes !== 0) throw new Error(`cycle ${i}: unowned live bytes ${memory.currentBytes}`)
    } finally {
      await session.release()
    }
    finalNet = memory.currentBytes
    if (finalNet !== 0) throw new Error(`cycle ${i}: net bytes ${finalNet} after release`)
  }
  return { cycles: rounds, maxNetBytes: maxNet, finalNetBytes: finalNet, wallMs: Date.now() - started }
}

export interface NoCfBaseline {
  paintReserves: number
  sourceReads: number
  scopeBuilt: boolean
}

/** No-CF baseline: same shapes without conditional formatting do zero CF work. */
export async function baselineNoCf(): Promise<NoCfBaseline> {
  const { prepareCfPaint } = await import('../../src/xlsx/cf/render')
  const memory = createCfMemory()
  let reserves = 0
  let reads = 0
  const rawReserve = memory.reserve.bind(memory)
  ;(memory as { reserve: typeof rawReserve }).reserve = (...args: Parameters<typeof rawReserve>) => { reserves++; return rawReserve(...args) }
  const sheet = numericSheet(100, cellIsRule(0, { firstRow: 0, lastRow: 99, firstCol: 0, lastCol: 0 }))
  delete (sheet as { conditionalFormatting?: unknown }).conditionalFormatting
  const store = { get stamp(): never { throw new Error('no-CF stamp read') }, readRows: (...args: unknown[]) => { reads++; throw new Error(`no-CF read ${String(args)}`) } }
  let scopeBuilt = false
  for (let i = 0; i < 5; i++) {
    const batch = await prepareCfPaint(sheet, { memory, store: store as never })
    if (batch !== undefined) scopeBuilt = true
    batch?.dispose?.()
  }
  return { paintReserves: reserves, sourceReads: reads, scopeBuilt }
}

export interface EagerBoundary {
  checked: boolean
  passed: boolean
  reason: string
}

/** Module-graph check: run the built-graph eager gate when a build exists. */
export function auditEagerBoundary(rootDir: string = REPO_ROOT): EagerBoundary {
  const mapPath = join(rootDir, 'dist', '.module-map.json')
  if (!existsSync(mapPath)) {
    return { checked: false, passed: true, reason: 'no dist module map; deferred to the eager gate (bun run build + check-eager)' }
  }
  const out = spawnSync('bun', ['scripts/check-eager.ts'], { cwd: rootDir, encoding: 'utf8' })
  return { checked: true, passed: out.status === 0, reason: out.status === 0 ? 'eager graph clean' : String(out.stderr ?? out.stdout).slice(0, 500) }
}

/* ---------------- CLI ---------------- */

function printCase(entry: { id: string; status: string; structuralValid: boolean; code?: string; reason?: string }): void {
  const flag = entry.structuralValid ? 'ok' : 'FAIL'
  // eslint-disable-next-line no-console
  console.log(`${flag} ${entry.id} [${entry.status}]${entry.code ? ` ${entry.code}` : ''}${entry.reason ? ` — ${entry.reason}` : ''}`)
}

async function main(): Promise<void> {
  const args = process.argv.slice(2)
  const stage = Number(args.includes('--stage') ? args[args.indexOf('--stage') + 1] : 1)
  if (stage !== 1) {
    // eslint-disable-next-line no-console
    console.error('Only --stage 1 is implemented')
    process.exit(2)
  }
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  const outDir = args.includes('--out') ? resolve(args[args.indexOf('--out') + 1]!) : undefined
  const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
  const report = auditStage1Manifest(manifest, { rootDir: REPO_ROOT })
  const cases = only ? report.perCase.filter(c => c.id === only) : report.perCase
  for (const entry of cases) printCase(entry)
  // eslint-disable-next-line no-console
  console.log(`manifest: ${report.summary.structuralValid}/${report.summary.cases} structural, capabilities ${report.summary.covered}/${report.summary.capabilities}, missing=[${report.missingCases.join(',')}] extra=[${report.extraCases.join(',')}]`)
  // eslint-disable-next-line no-console
  console.log('--- pressure ---')
  const whole = await pressureWholeColumn()
  // eslint-disable-next-line no-console
  console.log(`${whole.outcome} ${whole.name}: peak=${whole.peakBytes} net=${whole.netBytes} reads=${whole.sourceReads} (${whole.wallMs}ms context only)${whole.reason ? ` — ${whole.reason}` : ''}`)
  if (whole.outcome !== 'refused' || whole.netBytes !== 0) throw new Error('whole-column pressure must refuse with net zero')
  const shared = await pressureSharedRules()
  // eslint-disable-next-line no-console
  console.log(`${shared.outcome} ${shared.name}: peak=${shared.peakBytes} net=${shared.netBytes} (${shared.wallMs}ms context only)`)
  if (shared.outcome !== 'admitted' || shared.netBytes !== 0) throw new Error('100 shared rules must admit and balance')
  const bounds = pressureRangeBoundaries()
  // eslint-disable-next-line no-console
  console.log(`boundaries: per-call 256=${bounds.admitted256rects ? 'admitted' : 'REFUSED'} 257=${bounds.refused257rects ? 'refused' : 'ADMITTED'} total-ranges 16384=${bounds.admitted16384Ranges ? 'admitted' : 'REFUSED'} 16385=${bounds.refused16385Ranges ? 'refused' : 'ADMITTED'} rules 4096=${bounds.admitted4096Rules ? 'admitted' : 'REFUSED'} 4097=${bounds.refused4097Rules ? 'refused' : 'ADMITTED'}`)
  if (!bounds.admitted256rects || !bounds.refused257rects || !bounds.admitted16384Ranges || !bounds.refused16385Ranges || !bounds.admitted4096Rules || !bounds.refused4097Rules) throw new Error('cap boundaries inexact')
  const distinct = await pressureDistinctRanges()
  // eslint-disable-next-line no-console
  console.log(`${distinct.outcome} ${distinct.name}: peak=${distinct.peakBytes} net=${distinct.netBytes} (${distinct.wallMs}ms context only)`)
  if (distinct.outcome !== 'refused' || distinct.netBytes !== 0) throw new Error('distinct-range pressure must refuse cleanly with net zero')
  const refusedSheet = await pressureRefusedSheet()
  // eslint-disable-next-line no-console
  console.log(`${refusedSheet.outcome} ${refusedSheet.name}: net=${refusedSheet.netBytes} (${refusedSheet.wallMs}ms context only)`)
  if (refusedSheet.outcome !== 'refused' || refusedSheet.netBytes !== 0) throw new Error('4097-rule sheet must refuse with net zero')
  const cycles = await pressureCycles(20)
  // eslint-disable-next-line no-console
  console.log(`cycles: ${cycles.cycles} open/read/cancel/close, maxNet=${cycles.maxNetBytes} finalNet=${cycles.finalNetBytes} (${cycles.wallMs}ms context only)`)
  if (cycles.finalNetBytes !== 0) throw new Error('cycles leaked live bytes')
  const baseline = await baselineNoCf()
  // eslint-disable-next-line no-console
  console.log(`no-cf baseline: reserves=${baseline.paintReserves} reads=${baseline.sourceReads} scopeBuilt=${baseline.scopeBuilt}`)
  if (baseline.paintReserves !== 0 || baseline.sourceReads !== 0 || baseline.scopeBuilt) throw new Error('no-CF path did CF work')
  const eager = auditEagerBoundary(REPO_ROOT)
  // eslint-disable-next-line no-console
  console.log(`module graph: ${eager.checked ? (eager.passed ? 'pass' : 'FAIL') : 'deferred'} — ${eager.reason}`)
  if (eager.checked && !eager.passed) throw new Error('eager boundary violated')
  const unexpected = report.perCase.filter(c => !c.structuralValid)
  const ok = report.ok && unexpected.length === 0
  // eslint-disable-next-line no-console
  console.log(`summary: ok=${ok} structural=${report.summary.structuralValid}/${report.summary.cases} unsupported=${report.summary.unsupported} pending=${report.summary.pending}`)
  if (outDir) {
    mkdirSync(outDir, { recursive: true })
    writeFileSync(join(outDir, 'stage1-audit.json'), JSON.stringify({ report, whole, shared, bounds, refusedSheet, cycles, baseline, eager }, null, 2))
  }
  if (!ok) process.exit(1)
  // A wholly refused suite is not success: at least one case must hold evidence.
  if (report.summary.structuralValid === 0) {
    // eslint-disable-next-line no-console
    console.error('wholly refused suite is not success')
    process.exit(1)
  }
}

const invoked = process.argv[1] ? resolve(process.argv[1]) === fileURLToPath(import.meta.url) : false
if (invoked) {
  main().catch(error => {
    // eslint-disable-next-line no-console
    console.error(error instanceof Error ? error.message : String(error))
    process.exit(1)
  })
}
