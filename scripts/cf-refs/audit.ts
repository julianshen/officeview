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
  CF_STAGE2_CAPABILITIES,
  CF_STAGE2_CASES,
  type CfManifest,
  type CfOracleCase,
  type CfOracleStatus,
} from './cases'
import { CF_MANIFEST_PATH } from './generate'
import { PackageError } from '../../src/core/package-error'
import { prepareCfTextScope } from '../../src/xlsx/cf/text'
import { createCfMemory, createCfWorkBudget } from '../../src/xlsx/cf/limits'
import { estimateCfPrepareBytes } from '../../src/xlsx/cf/service'
import { residentSource } from '../../src/xlsx/cf/render'
import { createCfStatistics } from '../../src/xlsx/cf/statistics'
import { computeMetrics } from '../../src/xlsx/render'
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
  /** Stage 1 completion requires verified independent semantic evidence for
   * every supported capability. Structural integrity (`ok`) without verified
   * captures reports blocked, never green. */
  stageComplete: boolean
  blockedCapabilities: string[]
  summary: {
    cases: number
    structuralValid: number
    verified: number
    unsupported: number
    pending: number
    capabilities: number
    covered: number
    verifiedCapabilities: number
  }
}

/**
 * Pure stage gate: complete only when every supported capability is verified.
 * Refused-unsupported capabilities never block; supported-structural (measured
 * structurally but without verified independent evidence) and unaccounted
 * capabilities block with an explicit list.
 */
export function stageGateStatus(perCapability: Stage1CapabilityAudit[]): { complete: boolean; blocked: string[] } {
  const blocked = perCapability
    .filter(entry => entry.disposition === 'supported-structural' || entry.disposition === 'unaccounted')
    .map(entry => entry.capabilityId)
  return { complete: blocked.length === 0, blocked }
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
  const stage1ManifestCases = manifest.cases.filter(c => c.stage === 1 || (!c.stage && !c.id.startsWith('stage2')))
  const catalogIds = CF_STAGE1_CASES.map(c => c.id)
  const manifestIds = new Set(stage1ManifestCases.map(c => c.id))
  const missingCases = catalogIds.filter(id => !manifestIds.has(id))
  const catalogSet = new Set(catalogIds)
  const extraCases = stage1ManifestCases.map(c => c.id).filter(id => !catalogSet.has(id))
  const perCase = stage1ManifestCases.map(c => auditCase(c, rootDir))
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
  const gate = stageGateStatus(perCapability)
  return {
    ok,
    perCase,
    perCapability,
    missingCases,
    extraCases,
    stageComplete: ok && gate.complete,
    blockedCapabilities: gate.blocked,
    summary: {
      cases: perCase.length,
      structuralValid,
      verified: perCase.filter(p => p.status === 'measured' && p.code === undefined).length,
      unsupported: perCase.filter(p => p.status === 'unsupported').length,
      pending: perCase.filter(p => p.status === 'recipe-pending' || p.status === 'geometry-unverified').length,
      capabilities: perCapability.length,
      covered: perCapability.filter(p => p.disposition !== 'unaccounted').length,
      verifiedCapabilities: perCapability.filter(p => p.disposition === 'verified').length,
    },
  }
}

export type Stage2CaseAudit = Stage1CaseAudit
export type Stage2CapabilityAudit = Stage1CapabilityAudit
export type Stage2AuditReport = Stage1AuditReport

export function auditStage2Manifest(manifest: CfManifest, options: { rootDir?: string } = {}): Stage2AuditReport {
  const rootDir = options.rootDir ?? process.cwd()
  const stage2ManifestCases = manifest.cases.filter(c => c.stage === 2 || c.id.startsWith('stage2'))
  const catalogIds = CF_STAGE2_CASES.map(c => c.id)
  const manifestIds = new Set(stage2ManifestCases.map(c => c.id))
  const missingCases = catalogIds.filter(id => !manifestIds.has(id))
  const catalogSet = new Set(catalogIds)
  const extraCases = stage2ManifestCases.map(c => c.id).filter(id => !catalogSet.has(id))
  const perCase = stage2ManifestCases.map(c => auditCase(c, rootDir))
  const perCapability: Stage2CapabilityAudit[] = CF_STAGE2_CAPABILITIES.map(entry => {
    const owning = perCase.find(p => p.id === entry.caseId)
    let disposition: Stage2CapabilityAudit['disposition'] = 'unaccounted'
    if (owning?.structuralValid) {
      disposition = owning.status === 'unsupported' ? 'refused-unsupported' : owning.code === undefined || owning.status === 'measured' ? 'verified' : 'supported-structural'
      if (owning.status === 'measured' && owning.code !== undefined) disposition = 'supported-structural'
    }
    return { capabilityId: entry.capabilityId, caseId: entry.caseId, disposition }
  })
  const caseOk = (p: Stage2CaseAudit): boolean => {
    if (p.status === 'measured') return p.structuralValid && p.code === undefined
    return p.structuralValid
  }
  const structuralValid = perCase.filter(p => p.structuralValid).length
  const ok = missingCases.length === 0 && perCase.length > 0 && perCase.every(caseOk) && perCapability.every(p => p.disposition !== 'unaccounted')
  const gate = stageGateStatus(perCapability)
  return {
    ok,
    perCase,
    perCapability,
    missingCases,
    extraCases,
    stageComplete: ok && gate.complete,
    blockedCapabilities: gate.blocked,
    summary: {
      cases: perCase.length,
      structuralValid,
      verified: perCase.filter(p => p.status === 'measured' && p.code === undefined).length,
      unsupported: perCase.filter(p => p.status === 'unsupported').length,
      pending: perCase.filter(p => p.status === 'recipe-pending' || p.status === 'geometry-unverified').length,
      capabilities: perCapability.length,
      covered: perCapability.filter(p => p.disposition !== 'unaccounted').length,
      verifiedCapabilities: perCapability.filter(p => p.disposition === 'verified').length,
    },
  }
}

/* ---------------- pressure + lifecycle scenarios ---------------- */

/** PressureSample.sourceReads semantics (all observed, never literal):
 * resident preparations count distinct cells visited via counting wrappers
 * (0 when the estimator refuses pre-enumeration); worker cycles count
 * readConditionalFormatting calls issued. */
export interface PressureSample {
  name: string
  outcome: 'admitted' | 'refused'
  peakBytes: number
  netBytes: number
  sourceReads: number
  /** Cells with observed CF outcomes from a real scope (never a literal). */
  cellsWithResults: number
  reason?: string
  wallMs: number
}

/** Budget that observes real holds: net for leak detection, peak for work. */
export function createTrackingBudget(spy?: { reserves: number }): { budget: FormatBudget; net: () => number; peak: () => number } {
  let held = 0
  let max = 0
  const budget: FormatBudget = {
    reserve: bytes => { held += bytes; if (held > max) max = held; if (spy) spy.reserves++ },
    release: bytes => { held -= bytes },
    retain: (_targets, bytes, commit) => { held += bytes; if (held > max) max = held; commit() },
    settle: () => {},
  }
  return { budget, net: () => held, peak: () => max }
}

/**
 * Gate a pressure sample against observed reality. Net bytes must always be
 * zero (no leak); the outcome must match; admitted work must show observed
 * peak and cells (tautological zeros fail). Throws on violation.
 */
export function checkPressureSample(sample: PressureSample, options: { expectOutcome: 'admitted' | 'refused'; requireWork?: boolean }): void {
  if (sample.outcome !== options.expectOutcome) {
    throw new Error(`${sample.name}: outcome ${sample.outcome}, expected ${options.expectOutcome}${sample.reason ? ` (${sample.reason})` : ''}`)
  }
  if (sample.netBytes !== 0) {
    throw new Error(`${sample.name}: netBytes ${sample.netBytes} leaked (peak ${sample.peakBytes})`)
  }
  if (options.expectOutcome === 'admitted' && options.requireWork !== false) {
    if (sample.peakBytes <= 0) throw new Error(`${sample.name}: admitted but peakBytes ${sample.peakBytes} observed no work`)
    if (sample.sourceReads <= 0) throw new Error(`${sample.name}: admitted but sourceReads ${sample.sourceReads} observed no source visits`)
    // cellsWithResults is reported, not gated: an admitted no-match scope
    // legitimately yields zero outcomes after real work.
  }
  if (options.expectOutcome === 'refused' && sample.sourceReads !== 0) {
    throw new Error(`${sample.name}: refused but sourceReads ${sample.sourceReads} observed enumeration`)
  }
}

export function numericSheet(rows: number, rule: Record<string, unknown>): XlsxSheet {
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

export const cellIsRule = (priority: number, range: { firstRow: number; lastRow: number; firstCol: number; lastCol: number }) => ({
  id: priority,
  priority,
  documentOrder: priority,
  stopIfTrue: false,
  ranges: [range],
  writes: new Set(['numFmt']),
  dxf: { 'numFmt.code': '0.00' },
  definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 50 }], dialect: 'base', formulas: [] },
})

export const scaleRule = (priority: number, range: { firstRow: number; lastRow: number; firstCol: number; lastCol: number }, points?: unknown[]) => ({
  id: priority,
  priority,
  documentOrder: priority,
  stopIfTrue: false,
  ranges: [range],
  writes: new Set(['colorScale', 'fill']),
  definition: {
    kind: 'colorScale',
    points: points ?? [
      { bound: { type: 'min' }, argb: 0xfff8696b },
      { bound: { type: 'max' }, argb: 0xff63be7b },
    ],
    dialect: 'base',
    formulas: [],
  },
})

function wholeColumnRule(): Record<string, unknown> {
  return cellIsRule(1, { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 }) as unknown as Record<string, unknown>
}

/** Whole-column rule over a 100k-cell sheet: must refuse before enumeration. */
export async function pressureWholeColumn(): Promise<PressureSample> {
  const started = Date.now()
  const inner = numericSheet(100000, wholeColumnRule())
  // Count every data enumeration: the estimator must refuse from rules+rects
  // alone, never touching rows or cells.
  let rowReads = 0
  const realRows = inner.rows
  const sheet = {
    ...inner,
    get rows() { rowReads++; return realRows },
  } as unknown as XlsxSheet
  const { budget, net, peak } = createTrackingBudget()
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
    reason = 'estimator refused the 1M-row range before enumeration'
  }
  scope?.dispose()
  return {
    name: 'whole-column-over-1M-rows',
    outcome: scope === undefined ? 'refused' : 'admitted',
    peakBytes: peak(),
    netBytes: net(),
    sourceReads: rowReads,
    cellsWithResults: 0,
    reason,
    wallMs: Date.now() - started,
  }
}

/**
 * Wrap sheet cells to count distinct cells visited during preparation.
 * Snapshot the count right after the preparation returns, before probing the
 * scope (probing reads scope state, not sheet cells). Proxies are transparent
 * to the preparation.
 */
export function countingSheetCells(sheet: XlsxSheet): { sheet: XlsxSheet; visits: () => number } {
  const touched = new Set<string>()
  const rows = sheet.rows.map(band => ({
    ...band,
    cells: band.cells.map(cell => new Proxy(cell, {
      get: (t, p, r) => {
        if (typeof p === 'string') touched.add(`${t.row}:${t.col}`)
        return Reflect.get(t, p, r)
      },
    })),
  }))
  return { sheet: { ...sheet, rows } as unknown as XlsxSheet, visits: () => touched.size }
}

/** 100 shared rules over one range: admitted once, balanced on dispose. */
export async function pressureSharedRules(): Promise<PressureSample> {
  const started = Date.now()
  const rules = Array.from({ length: 100 }, (_, i) =>
    cellIsRule(i, { firstRow: 0, lastRow: 99, firstCol: 0, lastCol: 0 }),
  )
  const inner = numericSheet(100, rules[0] as Record<string, unknown>)
  ;(inner.conditionalFormatting as unknown as { rules: unknown[] }).rules = rules as never[]
  const { sheet, visits } = countingSheetCells(inner)
  const { budget, net, peak } = createTrackingBudget()
  const scope = await prepareCfTextScope(sheet, {
    rects: [{ r0: 0, r1: 99, c0: 0, c1: 0 }],
    signal: new AbortController().signal,
    locale: 'en-US',
    budget,
  })
  const admitted = scope !== undefined
  const visited = visits()
  // Observed outcomes from the real scope: values 0..99 against `> 50`
  // match on 51..99 (49 cells carry the winning numFmt).
  let matched = 0
  if (scope) {
    for (let row = 0; row < 100; row++) {
      if (scope.describe(row, 0) !== undefined) matched++
    }
  }
  const peakBytes = peak()
  scope?.dispose()
  return { name: '100-shared-rules', outcome: admitted ? 'admitted' : 'refused', peakBytes, netBytes: net(), sourceReads: visited, cellsWithResults: matched, wallMs: Date.now() - started }
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
  const inner = {
    name: 'D', sheetId: 'd', rows: [{ index: 0, cells }], cols: [], mergeRefs: [], mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules },
  } as unknown as XlsxSheet
  const { sheet, visits } = countingSheetCells(inner)
  // Small distinct batch admits and balances (gated on observed work).
  {
    const { budget, net, peak } = createTrackingBudget()
    const scope = await prepareCfTextScope(sheet, {
      rects: Array.from({ length: 10 }, (_, i) => ({ r0: i, r1: i, c0: 0, c1: 0 })),
      signal: new AbortController().signal,
      locale: 'en-US',
      budget,
    })
    if (scope === undefined) throw new Error('small distinct-range batch must admit')
    const visited = visits()
    let matched = 0
    for (let row = 0; row < 10; row++) {
      if (scope.describe(row, 0) !== undefined) matched++
    }
    const peakBytes = peak()
    scope.dispose()
    const small: PressureSample = {
      name: 'distinct-range-10', outcome: 'admitted', peakBytes, netBytes: net(),
      sourceReads: visited, cellsWithResults: matched, wallMs: 0,
    }
    checkPressureSample(small, { expectOutcome: 'admitted' })
  }
  // Large distinct batch exceeds the scratch bound and refuses fail-closed.
  const visitedBefore = visits()
  const { budget, net, peak } = createTrackingBudget()
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
  return { name: 'distinct-range-300', outcome: scope === undefined ? 'refused' : 'admitted', peakBytes: peak(), netBytes: net(), sourceReads: visits() - visitedBefore, cellsWithResults: 0, wallMs: Date.now() - started }
}

/** Refused 4097-rule sheet: fail closed with zero net charge. */
export async function pressureRefusedSheet(): Promise<PressureSample> {
  const started = Date.now()
  const rules = Array.from({ length: 4097 }, (_, i) =>
    cellIsRule(i, { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }),
  )
  const inner = numericSheet(1, rules[0] as Record<string, unknown>)
  ;(inner.conditionalFormatting as unknown as { rules: unknown[] }).rules = rules as never[]
  let rowReads = 0
  const realRows = inner.rows
  const sheet = { ...inner, get rows() { rowReads++; return realRows } } as unknown as XlsxSheet
  const { budget, net, peak } = createTrackingBudget()
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
  return { name: 'refused-4097-rule-sheet', outcome: scope === undefined ? 'refused' : 'admitted', peakBytes: peak(), netBytes: net(), sourceReads: rowReads, cellsWithResults: 0, wallMs: Date.now() - started }
}

/** Statistics visits: a top10-style scan over 200 stored cells must visit all
 * of them (product counter, not an audit literal) and balance on release. */
export async function pressureStatisticsVisits(): Promise<PressureSample> {
  const started = Date.now()
  const sheet = numericSheet(200, cellIsRule(0, { firstRow: 0, lastRow: 199, firstCol: 0, lastCol: 0 }))
  const memory = createCfMemory()
  const stamp = { owner: 'audit', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
  const source = residentSource(sheet, stamp, memory, 'engine')
  const session = createCfStatistics(memory)
  const work = createCfWorkBudget()
  const outs = await session.prepareStatistics(
    [[{ firstRow: 0, lastRow: 199, firstCol: 0, lastCol: 0 }]],
    stamp,
    source,
    work,
    { signal: new AbortController().signal },
  )
  const ready = outs.find(o => o.kind === 'ready')
  if (!ready || ready.kind !== 'ready') throw new Error('statistics scan must admit')
  const stats = ready.value.value
  const visited = stats.storedCells
  const peak = memory.peakBytes
  ready.value.release()
  session.dispose()
  return {
    name: 'statistics-200-cells',
    outcome: 'admitted',
    peakBytes: peak,
    netBytes: memory.currentBytes,
    sourceReads: visited,
    cellsWithResults: stats.spanCells,
    wallMs: Date.now() - started,
  }
}

/** 1M-row synthetic A:A scope with 100 scales sharing statistics (one scan, Lesson 3). */
export async function pressureScale1MRowSharedStatistics(): Promise<PressureSample> {
  const started = Date.now()
  const sheet = numericSheet(100, scaleRule(0, { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 }))
  const memory = createCfMemory()
  const stamp = { owner: 'audit', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
  const source = residentSource(sheet, stamp, memory, 'engine')
  const session = createCfStatistics(memory, { retainEntries: true })
  const work = createCfWorkBudget({ noTimeLimit: true, maxUnits: 10_000_000 })
  const wholeCol = [{ firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 }]
  const scopes = Array.from({ length: 100 }, () => wholeCol)
  const outs = await session.prepareStatistics(
    scopes,
    stamp,
    source,
    work,
    { signal: new AbortController().signal },
  )
  const ready = outs.find(o => o.kind === 'ready')
  if (!ready || ready.kind !== 'ready') throw new Error('statistics scan must admit')
  const stats = ready.value.value
  const visited = stats.storedCells
  const peak = memory.peakBytes
  for (const o of outs) {
    if (o.kind === 'ready') o.value.release()
  }
  session.dispose()
  return {
    name: 'scale-1M-row-shared-statistics',
    outcome: 'admitted',
    peakBytes: peak,
    netBytes: memory.currentBytes,
    sourceReads: visited,
    cellsWithResults: stats.spanCells,
    wallMs: Date.now() - started,
  }
}

/** Refused percentile copy with min/max still rendering. */
export async function pressureRefusedPercentileWithMinMax(): Promise<PressureSample> {
  const started = Date.now()
  const cells: XlsxCell[] = []
  for (let row = 0; row < 10; row++) {
    cells.push({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: (row + 1) * 10, hasCachedValue: true })
  }
  for (let row = 0; row < 10; row++) {
    cells.push({ ref: `B${row + 1}`, row, col: 1, styleIndex: 0, value: (row + 1) * 10, hasCachedValue: true })
  }
  const minMaxRule = scaleRule(1, { firstRow: 0, lastRow: 9, firstCol: 0, lastCol: 0 })
  const refusedPercentileRule = {
    id: 2,
    priority: 2,
    documentOrder: 1,
    stopIfTrue: false,
    ranges: [{ firstRow: 0, lastRow: 9, firstCol: 1, lastCol: 1 }],
    writes: new Set(['colorScale', 'fill']),
    definition: {
      kind: 'colorScale',
      points: [
        { bound: { type: 'percentile', value: -10 }, argb: 0xfff8696b },
        { bound: { type: 'max' }, argb: 0xff63be7b },
      ],
      dialect: 'base',
      formulas: [],
    },
  }
  const inner: XlsxSheet = {
    name: 'P',
    sheetId: 'p',
    rows: [{ index: 0, cells }],
    cols: [],
    mergeRefs: [],
    mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules: [minMaxRule as never, refusedPercentileRule as never] },
  } as unknown as XlsxSheet
  const { sheet, visits } = countingSheetCells(inner)
  const { budget, net, peak } = createTrackingBudget()
  const scope = await prepareCfTextScope(sheet, {
    rects: [
      { r0: 0, r1: 9, c0: 0, c1: 0 },
      { r0: 0, r1: 9, c0: 1, c1: 1 },
    ],
    signal: new AbortController().signal,
    locale: 'en-US',
    budget,
  })
  if (scope === undefined) throw new Error('scope must admit')
  const visited = visits()
  let renderedCells = 0
  for (let row = 0; row < 10; row++) {
    const descA = scope.describe(row, 0)
    if (descA !== undefined && descA.text !== 'Conditional formatting unavailable') {
      renderedCells++
    }
    const descB = scope.describe(row, 1)
    if (descB?.text !== 'Conditional formatting unavailable') {
      throw new Error(`percentile cell B${row + 1} must withhold as unavailable, got ${descB?.text}`)
    }
  }
  const peakBytes = peak()
  scope.dispose()
  return {
    name: 'refused-percentile-with-min-max',
    outcome: 'admitted',
    peakBytes,
    netBytes: net(),
    sourceReads: visited,
    cellsWithResults: renderedCells,
    wallMs: Date.now() - started,
  }
}

/** Whole-column scope CF ranges do not extend extents (Lesson 4). */
export async function pressureScaleWholeColumnExtents(): Promise<PressureSample> {
  const started = Date.now()
  const dataRows = 10
  const cells: XlsxCell[] = []
  for (let r = 0; r < dataRows; r++) {
    cells.push({ ref: `A${r + 1}`, row: r, col: 0, styleIndex: 0, value: (r + 1) * 10, hasCachedValue: true })
  }
  const wholeColumnScale = scaleRule(1, { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 })
  const inner: XlsxSheet = {
    name: 'WCExtents',
    sheetId: 'wc',
    rows: Array.from({ length: dataRows }, (_, r) => ({
      index: r,
      cells: [cells[r]!],
    })),
    cols: [],
    mergeRefs: [],
    mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules: [wholeColumnScale as never] },
  } as unknown as XlsxSheet
  const m = computeMetrics(inner)
  if (m.rowHeightsPx.length !== dataRows || m.heightPx > 1000) {
    throw new Error(`CF scale extended metrics extent: ${m.rowHeightsPx.length} rows, height ${m.heightPx}px`)
  }
  const { sheet, visits } = countingSheetCells(inner)
  const { budget, net, peak } = createTrackingBudget()
  const scope = await prepareCfTextScope(sheet, {
    rects: [{ r0: 0, r1: dataRows - 1, c0: 0, c1: 0 }],
    signal: new AbortController().signal,
    locale: 'en-US',
    budget,
  })
  if (scope === undefined) throw new Error('scope must admit')
  const visited = visits()
  let renderedCells = 0
  for (let r = 0; r < dataRows; r++) {
    const desc = scope.describe(r, 0)
    if (desc !== undefined && desc.text !== 'Conditional formatting unavailable') renderedCells++
  }
  const peakBytes = peak()
  scope.dispose()
  return {
    name: 'scale-whole-column-extents',
    outcome: 'admitted',
    peakBytes,
    netBytes: net(),
    sourceReads: visited,
    cellsWithResults: renderedCells,
    wallMs: Date.now() - started,
  }
}

export interface CycleSample {
  cycles: number
  reads: number
  maxPeakBytes: number
  maxNetBytes: number
  finalNetBytes: number
  wallMs: number
}

/** Open/read/cancel/close cycles against a real paged session. A fresh memory
 * observes every round: peak proves real work, and complete teardown (store
 * release for shared definitions plus session release for pages) must return
 * to a zero baseline — stability alone is not cleanup. Pre-aborted reads must
 * throw cancellation; a mid-flight abort must settle without charge and leave
 * the session usable. */
export async function pressureCycles(rounds = 20): Promise<CycleSample> {
  const started = Date.now()
  const { openImportSession } = await import('../../src/xlsx/storage/pagedImport')
  const { createPagedWorkbookSource } = await import('../../src/xlsx/view/pagedSource')
  const saved = readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/stage1-cellis-literals.xlsx'))
  let maxPeak = 0
  let maxNet = 0
  let finalNet = 0
  let reads = 0
  for (let i = 0; i < rounds; i++) {
    // Fresh memory per round (caller-scoped, as in real use).
    const memory = createCfMemory()
    const session = await openImportSession(
      { kind: 'blob', blob: new Blob([new Uint8Array(saved).buffer]) },
      { docId: `audit-cycle-${i}`, profile: 'D' },
    )
    let store: ReturnType<ReturnType<typeof createPagedWorkbookSource>['store']> | undefined
    try {
      await session.done
      const source = createPagedWorkbookSource(session)
      const sheetId = source.sheets[0]!.sheetId
      store = source.store(sheetId)
      const stamp = (owner: string) => ({ owner, sheetId, workbookValueGen: store!.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 })
      // Complete read through the audited memory: peak proves real work.
      reads++
      const ready = await store.readConditionalFormatting?.({
        rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
        stamp: stamp('audit'),
        grantBytes: 2 * 1024 * 1024,
      }, memory, { signal: new AbortController().signal })
      if (!ready || ready.kind !== 'ready') throw new Error(`cycle ${i}: expected ready read`)
      if (memory.peakBytes <= 0) throw new Error(`cycle ${i}: ready read observed no allocation`)
      maxPeak = Math.max(maxPeak, memory.peakBytes)
      maxNet = Math.max(maxNet, memory.currentBytes)
      // Cancelled read: pre-aborted signal must throw cancellation, never
      // success, with no charge.
      const aborted = new AbortController()
      aborted.abort()
      let cancelled: unknown
      try {
        reads++
        await store.readConditionalFormatting?.({
          rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
          stamp: stamp('audit'),
          grantBytes: 2 * 1024 * 1024,
        }, memory, { signal: aborted.signal })
      } catch (error) {
        cancelled = error
      }
      if (!(cancelled instanceof PackageError) || cancelled.code !== 'cancelled') {
        const cancelledCode = typeof (cancelled as unknown as { code?: unknown })?.code === 'string'
          ? (cancelled as unknown as { code: string }).code
          : 'no-code'
        const got = cancelled instanceof Error ? cancelled.constructor.name + '(' + cancelledCode + ')' : String(cancelled)
        throw new Error(`cycle ${i}: pre-aborted read must throw cancelled, got ${got}`)
      }
      maxNet = Math.max(maxNet, memory.currentBytes)
      // Mid-flight abort: settle without charge and leave the session usable.
      const mid = new AbortController()
      reads++
      const inflight = store.readConditionalFormatting?.({
        rects: [{ r0: 0, r1: 99, c0: 0, c1: 0 }],
        stamp: stamp('audit'),
        grantBytes: 2 * 1024 * 1024,
      }, memory, { signal: mid.signal })
      mid.abort()
      let midOutcome: string
      try {
        const midResult = await inflight
        midOutcome = midResult?.kind ?? 'missing'
      } catch (error) {
        midOutcome = error instanceof PackageError ? `threw:${error.code}` : `threw:${error instanceof Error ? error.constructor.name : String(error)}`
      }
      if (midOutcome === 'ready') throw new Error(`cycle ${i}: mid-flight abort returned ready`)
      // Session still usable after aborts.
      reads++
      const again = await store.readConditionalFormatting?.({
        rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
        stamp: stamp('audit'),
        grantBytes: 2 * 1024 * 1024,
      }, memory, { signal: new AbortController().signal })
      if (!again || again.kind !== 'ready') throw new Error(`cycle ${i}: session unusable after abort (mid-flight outcome ${midOutcome})`)
      maxPeak = Math.max(maxPeak, memory.peakBytes)
    } finally {
      // Complete teardown: store release drops shared definitions and drawing
      // holds; session release removes pages. Caller-scoped memory must return
      // to the released baseline — no orphaned charges.
      await store?.release()
      await session.release()
    }
    finalNet = memory.currentBytes
    if (finalNet !== 0) throw new Error(`cycle ${i}: net bytes ${finalNet} after release`)
  }
  return { cycles: rounds, reads, maxPeakBytes: maxPeak, maxNetBytes: maxNet, finalNetBytes: finalNet, wallMs: Date.now() - started }
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
  if (stage !== 1 && stage !== 2) {
    // eslint-disable-next-line no-console
    console.error('Only --stage 1 and --stage 2 are implemented')
    process.exit(2)
  }
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  const outDir = args.includes('--out') ? resolve(args[args.indexOf('--out') + 1]!) : undefined
  const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
  const report = stage === 1
    ? auditStage1Manifest(manifest, { rootDir: REPO_ROOT })
    : auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
  if (args.includes('--write')) {
    if (stage === 1) {
      const verified = report.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()
      const content = `/**
 * Verified conditional formatting capabilities for Stage 1.
 * Generated via bun scripts/cf-refs/audit.ts --stage 1 --write.
 * Checked by tests/xlsx-cf-stage1.test.ts.
 */
export const VERIFIED_CAPABILITIES: ReadonlySet<string> = new Set([
${verified.map(id => `  '${id}',`).join('\n')}
])
`
      writeFileSync(resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), content, 'utf8')
      // eslint-disable-next-line no-console
      console.log(`wrote ${verified.length} verified capabilities to src/xlsx/cf/verifiedCapabilities.ts`)
    } else {
      const stage1Report = auditStage1Manifest(manifest, { rootDir: REPO_ROOT })
      const stage1Verified = stage1Report.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()
      const stage2Verified = report.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()
      const allVerified = [...new Set([...stage1Verified, ...stage2Verified])].sort()
      const content = `/**
 * Verified conditional formatting capabilities for Stage 1 and Stage 2.
 * Generated via bun scripts/cf-refs/audit.ts --stage 2 --write.
 * Checked by tests/xlsx-cf-stage1.test.ts and tests/xlsx-cf-stage2.test.ts.
 */
export const VERIFIED_STAGE1_CAPABILITIES: ReadonlySet<string> = new Set([
${stage1Verified.map(id => `  '${id}',`).join('\n')}
])

export const VERIFIED_STAGE2_CAPABILITIES: ReadonlySet<string> = new Set([
${stage2Verified.map(id => `  '${id}',`).join('\n')}
])

export const VERIFIED_CAPABILITIES: ReadonlySet<string> = new Set([
${allVerified.map(id => `  '${id}',`).join('\n')}
])
`
      writeFileSync(resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), content, 'utf8')
      // eslint-disable-next-line no-console
      console.log(`wrote ${allVerified.length} verified capabilities (${stage1Verified.length} S1, ${stage2Verified.length} S2) to src/xlsx/cf/verifiedCapabilities.ts`)
    }
  }
  const cases = only ? report.perCase.filter(c => c.id === only) : report.perCase
  for (const entry of cases) printCase(entry)
  // eslint-disable-next-line no-console
  console.log(`manifest: ${report.summary.structuralValid}/${report.summary.cases} structural, capabilities ${report.summary.covered}/${report.summary.capabilities}, missing=[${report.missingCases.join(',')}] extra=[${report.extraCases.join(',')}]`)
  // eslint-disable-next-line no-console
  console.log('--- pressure ---')

  let stage1Pressure: { whole: PressureSample; shared: PressureSample; bounds: ReturnType<typeof pressureRangeBoundaries>; refusedSheet: PressureSample; stats: PressureSample } | undefined
  let stage2Pressure: { scaleShared: PressureSample; percentile: PressureSample; extents: PressureSample } | undefined

  if (stage === 1) {
    const whole = await pressureWholeColumn()
    // eslint-disable-next-line no-console
    console.log(`${whole.outcome} ${whole.name}: peak=${whole.peakBytes} net=${whole.netBytes} reads=${whole.sourceReads} cells=${whole.cellsWithResults} (${whole.wallMs}ms context only)${whole.reason ? ` — ${whole.reason}` : ''}`)
    checkPressureSample(whole, { expectOutcome: 'refused' })
    const shared = await pressureSharedRules()
    // eslint-disable-next-line no-console
    console.log(`${shared.outcome} ${shared.name}: peak=${shared.peakBytes} net=${shared.netBytes} reads=${shared.sourceReads} cells=${shared.cellsWithResults} (${shared.wallMs}ms context only)`)
    checkPressureSample(shared, { expectOutcome: 'admitted' })
    const bounds = pressureRangeBoundaries()
    // eslint-disable-next-line no-console
    console.log(`boundaries: per-call 256=${bounds.admitted256rects ? 'admitted' : 'REFUSED'} 257=${bounds.refused257rects ? 'refused' : 'ADMITTED'} total-ranges 16384=${bounds.admitted16384Ranges ? 'admitted' : 'REFUSED'} 16385=${bounds.refused16385Ranges ? 'refused' : 'ADMITTED'} rules 4096=${bounds.admitted4096Rules ? 'admitted' : 'REFUSED'} 4097=${bounds.refused4097Rules ? 'refused' : 'ADMITTED'}`)
    if (!bounds.admitted256rects || !bounds.refused257rects || !bounds.admitted16384Ranges || !bounds.refused16385Ranges || !bounds.admitted4096Rules || !bounds.refused4097Rules) throw new Error('cap boundaries inexact')
    const refusedSheet = await pressureRefusedSheet()
    // eslint-disable-next-line no-console
    console.log(`${refusedSheet.outcome} ${refusedSheet.name}: peak=${refusedSheet.peakBytes} net=${refusedSheet.netBytes} reads=${refusedSheet.sourceReads} (${refusedSheet.wallMs}ms context only)`)
    checkPressureSample(refusedSheet, { expectOutcome: 'refused' })
    const stats = await pressureStatisticsVisits()
    // eslint-disable-next-line no-console
    console.log(`${stats.outcome} ${stats.name}: peak=${stats.peakBytes} net=${stats.netBytes} visits=${stats.sourceReads} span=${stats.cellsWithResults} (${stats.wallMs}ms context only)`)
    checkPressureSample(stats, { expectOutcome: 'admitted' })
    stage1Pressure = { whole, shared, bounds, refusedSheet, stats }
  } else {
    const scaleShared = await pressureScale1MRowSharedStatistics()
    // eslint-disable-next-line no-console
    console.log(`${scaleShared.outcome} ${scaleShared.name}: peak=${scaleShared.peakBytes} net=${scaleShared.netBytes} reads=${scaleShared.sourceReads} cells=${scaleShared.cellsWithResults} (${scaleShared.wallMs}ms context only)`)
    checkPressureSample(scaleShared, { expectOutcome: 'admitted' })
    const percentile = await pressureRefusedPercentileWithMinMax()
    // eslint-disable-next-line no-console
    console.log(`${percentile.outcome} ${percentile.name}: peak=${percentile.peakBytes} net=${percentile.netBytes} reads=${percentile.sourceReads} cells=${percentile.cellsWithResults} (${percentile.wallMs}ms context only)`)
    checkPressureSample(percentile, { expectOutcome: 'admitted' })
    const extents = await pressureScaleWholeColumnExtents()
    // eslint-disable-next-line no-console
    console.log(`${extents.outcome} ${extents.name}: peak=${extents.peakBytes} net=${extents.netBytes} reads=${extents.sourceReads} cells=${extents.cellsWithResults} (${extents.wallMs}ms context only)`)
    checkPressureSample(extents, { expectOutcome: 'admitted' })
    stage2Pressure = { scaleShared, percentile, extents }
  }

  const distinct = await pressureDistinctRanges()
  // eslint-disable-next-line no-console
  console.log(`${distinct.outcome} ${distinct.name}: peak=${distinct.peakBytes} net=${distinct.netBytes} reads=${distinct.sourceReads} cells=${distinct.cellsWithResults} (${distinct.wallMs}ms context only)`)
  checkPressureSample(distinct, { expectOutcome: 'refused' })

  const cycles = await pressureCycles(20)
  // eslint-disable-next-line no-console
  console.log(`cycles: ${cycles.cycles} open/read/cancel/close, reads=${cycles.reads} maxPeak=${cycles.maxPeakBytes} maxNet=${cycles.maxNetBytes} finalNet=${cycles.finalNetBytes} (released baseline, ${cycles.wallMs}ms context only)`)
  if (cycles.maxPeakBytes <= 0) throw new Error('cycles observed no allocation')
  if (cycles.reads < cycles.cycles * 4) throw new Error(`cycles issued ${cycles.reads} reads, expected at least ${cycles.cycles * 4} (ready + cancelled + mid-flight + retry per cycle)`)

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
  console.log(`summary: ok=${ok} structural=${report.summary.structuralValid}/${report.summary.cases} unsupported=${report.summary.unsupported} pending=${report.summary.pending} stageComplete=${report.stageComplete} blockedCapabilities=${report.blockedCapabilities.length}`)

  if (outDir) {
    mkdirSync(outDir, { recursive: true })
    const outData = stage === 1
      ? { report, ...stage1Pressure, distinct, cycles, baseline, eager }
      : { report, ...stage2Pressure, distinct, cycles, baseline, eager }
    writeFileSync(join(outDir, `stage${stage}-audit.json`), JSON.stringify(outData, null, 2))
  }

  if (!ok) process.exit(1)
  if (!report.stageComplete) {
    // eslint-disable-next-line no-console
    console.error(`BLOCKED: Stage ${stage} semantic evidence incomplete — ${report.blockedCapabilities.length} supported capabilities lack verified captures: ${report.blockedCapabilities.slice(0, 8).join(',')}${report.blockedCapabilities.length > 8 ? ',…' : ''}`)
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
