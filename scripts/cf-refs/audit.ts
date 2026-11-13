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
import {
  BAR_AUTO_COLOUR_CAPABILITIES,
  BAR_GRADIENT_CAPABILITY,
  BAR_LENGTH_RULES,
  drawsBarAutoColour,
  drawsBarGradient,
  isBarAutoColourCapability,
  isCapabilityVerified,
  judgeBarAutoColour,
  loadBarScreenCapture,
  measureCaseBars,
  judgeBarGradient,
  sampleBarAutoColours,
  sampleBarGradients,
  verifyCfCase,
  type BarAutoColourObservation,
  type BarGradientObservation,
} from './verify'
import {
  CF_STAGE1_CAPABILITIES,
  CF_STAGE1_CASES,
  CF_STAGE2_CAPABILITIES,
  CF_STAGE2_NONDISCRIMINATING,
  CF_STAGE2_CASES,
  CF_STAGE3_CAPABILITIES,
  CF_STAGE3_CASES,
  CF_STAGE3_NONDISCRIMINATING,
  CF_STAGE3_UNSUPPORTED,
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
  disposition: 'verified' | 'supported-structural' | 'refused-unsupported' | 'unverified-nondiscriminating' | 'unaccounted'
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
    /** Measured but non-discriminating (Ruling F-2): withheld in production until re-measured. */
    unverifiedCapabilities?: number
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
      // A measured but non-discriminating capability is never reported verified (Ruling F-2).
      if (disposition === 'verified' && CF_STAGE2_NONDISCRIMINATING[entry.capabilityId] !== undefined) disposition = 'unverified-nondiscriminating'
    }
    return { capabilityId: entry.capabilityId, caseId: entry.caseId, disposition }
  })
  const stage2Ids = new Set(CF_STAGE2_CAPABILITIES.map(entry => entry.capabilityId))
  const unknownNonDiscriminating = Object.keys(CF_STAGE2_NONDISCRIMINATING).filter(id => !stage2Ids.has(id) || !CF_STAGE2_NONDISCRIMINATING[id]!.trim())
  const caseOk = (p: Stage2CaseAudit): boolean => {
    if (p.status === 'measured') return p.structuralValid && p.code === undefined
    return p.structuralValid
  }
  const structuralValid = perCase.filter(p => p.structuralValid).length
  const ok = missingCases.length === 0 && perCase.length > 0 && perCase.every(caseOk) && perCapability.every(p => p.disposition !== 'unaccounted') && unknownNonDiscriminating.length === 0
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
      unverifiedCapabilities: perCapability.filter(p => p.disposition === 'unverified-nondiscriminating').length,
    },
  }
}

export type Stage3CaseAudit = Stage1CaseAudit
export type Stage3CapabilityAudit = Stage1CapabilityAudit
export interface Stage3AuditReport extends Stage1AuditReport {
  /** Ruling S3-F as amended: the measured ARGB of each VERIFIED automatic-colour capability (empty until one verifies). */
  autoBarColours: Readonly<Record<string, string>>
  /** Ruling S3-B: the one layout variant per scope sign that the verified captures select, with why any scope has none. */
  barLayout: BarLayoutDerivation
  /** Ruling S3-M (Task 22b): the verified gradient model's parameters, pooled across every capture (empty until verified). */
  barGradient: BarGradientExport
  /** The pooled gradient judgement, verified or not (for the report and the docs). */
  barGradientJudgement: string
  /** Task 22b: the verified rule for Excel's default 10/90 lengths (empty until bar:lengths-default-10-90 verifies). */
  barLengths: BarLengthsExport
}

/** The length rule production may apply to 10/90 bars: a `BAR_LENGTH_RULES` id matched by the owning capture. */
export type BarLengthsExport = Partial<Record<'default-10-90', { rule: string; settledBy: string[] }>>

/** The generated export of the verified 10/90 length rule (fixed key order, so `--write` stays idempotent). */
export function renderBarLengthsBlock(lengths: BarLengthsExport): string {
  const r = lengths['default-10-90']
  const rows = r ? `  'default-10-90': { rule: '${r.rule}', settledBy: [${r.settledBy.map(c => `'${c}'`).join(', ')}] },\n` : ''
  return `export const VERIFIED_STAGE3_BAR_LENGTHS: Readonly<Record<string, { rule: string; settledBy: readonly string[] }>> = {\n${rows}}\n`
}

/** The gradient model production may draw: `bar-length`, from the fill toward white by `endBlend` over the bar length. */
export type BarGradientExport = Partial<Record<'bar-length', { endBlend: number; settledBy: string[] }>>

/** A `CfBarLayoutVariant` (src/xlsx/cf/dataBar.ts) and the verified capabilities that select it. */
export interface BarLayoutSelection { origin: 'span' | 'axis'; lengthBasis: 'cell' | 'side'; settledBy: string[] }
export type BarLayoutScope = 'allPositive' | 'mixedSign'
export interface BarLayoutDerivation {
  layout: Partial<Record<BarLayoutScope, BarLayoutSelection>>
  findings: string[]
}

/**
 * Ruling S3-B: the layout each length/axis hypothesis of the bar predictor (`predictBarCell`, verify.ts) embodies.
 * - With no axis the predictor starts every bar at the cell's leading edge and applies minLength/maxLength to the whole
 *   cell: `{ span, cell }` (the `effective` lengths hypothesis).
 * - With an axis it predicts each side separately (`v / v_max` of the positive side, `v / v_min` of the negative side,
 *   `max(|v_min|, |v_max|)` for `middle`): `{ axis, side }`. It predicts axes only at 0/100 lengths, where the side and
 *   cell bases coincide; production therefore draws an axis only at 0/100 lengths.
 * A matched hypothesis absent from a capability's map (for example `base` or `full` lengths) selects no variant.
 */
export const BAR_LAYOUT_SETTLING: Readonly<Record<string, { scope: BarLayoutScope; variants: Readonly<Record<string, { origin: 'span' | 'axis'; lengthBasis: 'cell' | 'side' }>> }>> = {
  'bar:lengths-explicit-0-100': { scope: 'allPositive', variants: { effective: { origin: 'span', lengthBasis: 'cell' } } },
  // Every 10/90 length rule (Task 22b) starts at the leading edge and is a share of the whole cell: span/cell.
  'bar:lengths-default-10-90': { scope: 'allPositive', variants: Object.fromEntries(BAR_LENGTH_RULES.map(r => [r, { origin: 'span', lengthBasis: 'cell' }])) },
  'bar:axis-automatic': { scope: 'mixedSign', variants: { automatic: { origin: 'axis', lengthBasis: 'side' } } },
  'bar:axis-middle': { scope: 'mixedSign', variants: { middle: { origin: 'axis', lengthBasis: 'side' } } },
}

/**
 * Selects one variant per scope sign: every VERIFIED settling capability of the scope, pooled across every case's matched
 * hypotheses, must map to the same variant. Any match that maps to no variant or to a different one leaves the scope
 * without a variant (bars of that sign stay withheld) and records a finding; capability dispositions are untouched.
 */
export function deriveBarLayout(
  perCapability: readonly Pick<Stage3CapabilityAudit, 'capabilityId' | 'disposition'>[],
  matchedByCase: readonly Readonly<Record<string, string>>[],
): BarLayoutDerivation {
  const layout: BarLayoutDerivation['layout'] = {}
  const findings: string[] = []
  for (const scope of ['allPositive', 'mixedSign'] as const) {
    const settling = Object.keys(BAR_LAYOUT_SETTLING).filter(cap => BAR_LAYOUT_SETTLING[cap]!.scope === scope &&
      perCapability.some(p => p.capabilityId === cap && p.disposition === 'verified')).sort()
    if (settling.length === 0) continue
    const seen = new Map<string, string[]>()
    for (const matched of matchedByCase) {
      for (const cap of settling) {
        const hyp = matched[cap]
        if (hyp === undefined) continue
        const v = BAR_LAYOUT_SETTLING[cap]!.variants[hyp]
        const key = v ? `${v.origin}/${v.lengthBasis}` : `unmapped:${cap}=${hyp}`
        seen.set(key, [...(seen.get(key) ?? []), `${cap}=${hyp}`])
      }
    }
    const keys = [...seen.keys()]
    if (keys.length === 1 && !keys[0]!.startsWith('unmapped:')) {
      const [origin, lengthBasis] = keys[0]!.split('/') as ['span' | 'axis', 'cell' | 'side']
      layout[scope] = { origin, lengthBasis, settledBy: settling }
    } else {
      findings.push(`${scope}: no single layout variant (${keys.length === 0 ? 'no matched hypothesis' : [...seen].map(([k, v]) => `${k} <- ${v.join(', ')}`).join('; ')})`)
    }
  }
  return { layout, findings }
}

/** The generated export of the selected layout (fixed key order, so `--write` stays idempotent). */
export function renderBarLayoutBlock(layout: BarLayoutDerivation['layout']): string {
  const rows = (['allPositive', 'mixedSign'] as const).filter(scope => layout[scope]).map(scope => {
    const v = layout[scope]!
    return `  ${scope}: { origin: '${v.origin}', lengthBasis: '${v.lengthBasis}', settledBy: [${v.settledBy.map(c => `'${c}'`).join(', ')}] },\n`
  }).join('')
  return `export const VERIFIED_STAGE3_BAR_LAYOUT: Readonly<Record<string, { origin: string; lengthBasis: string; settledBy: readonly string[] }>> = {\n${rows}}\n`
}

/**
 * Every automatic bar colour that any measured Stage 3 capture draws, pooled across cases (refs are `<case>:<ref>`):
 * the cross-case half of Ruling S3-F as amended. Recomputed from each hash-bound PNG and Excel-resaved workbook.
 */
export function barAutoColoursAcrossCases(manifest: CfManifest, rootDir: string = REPO_ROOT): Record<string, BarAutoColourObservation> {
  const pooled: Record<string, BarAutoColourObservation> = {}
  for (const c of manifest.cases.filter(x => (x.stage === 3 || x.id.startsWith('stage3')) && x.status === 'measured')) {
    const capture = loadBarScreenCapture(c, rootDir, parts => drawsBarAutoColour(c, parts))
    if (!capture) continue
    const measured = measureCaseBars(c, capture.png, capture.geom, capture.parts)
    const observed = sampleBarAutoColours(c, capture.parts, capture.png, capture.geom, measured instanceof Map ? measured : undefined)
    for (const [cap, o] of Object.entries(observed)) {
      const into = (pooled[cap] ??= { samples: [], unsampled: [], others: [] })
      into.samples.push(...o.samples.map(s => ({ ref: `${c.id}:${s.ref}`, argb: s.argb })))
      into.unsampled.push(...o.unsampled.map(u => `${c.id}: ${u}`))
      if (o.others === undefined || into.others === undefined) into.others = undefined
      else for (const x of o.others) if (!into.others.includes(x)) into.others.push(x)
    }
  }
  return pooled
}

/** The colours production may consume: only a capability the audit reports verified AND whose pooled captures agree. */
export function verifiedAutoBarColours(
  perCapability: readonly Stage3CapabilityAudit[],
  pooled: Readonly<Record<string, BarAutoColourObservation>>,
): Record<string, string> {
  const out: Record<string, string> = {}
  for (const cap of Object.keys(BAR_AUTO_COLOUR_CAPABILITIES).sort()) {
    if (perCapability.find(p => p.capabilityId === cap)?.disposition !== 'verified') continue
    const judged = judgeBarAutoColour(pooled[cap])
    if (judged.status === 'discriminating' && judged.matched) out[cap] = judged.matched
  }
  return out
}

/** The generated export of the measured automatic colours (sorted, so `--write` stays idempotent). */
export function renderAutoBarColoursBlock(colours: Readonly<Record<string, string>>): string {
  const rows = Object.keys(colours).sort().map(cap => `  '${cap}': '${colours[cap]}',\n`).join('')
  return `export const VERIFIED_STAGE3_AUTO_BAR_COLORS: Readonly<Record<string, string>> = {\n${rows}}\n`
}

/**
 * Every gradient bar of the measured shape that any measured Stage 3 capture draws, pooled across cases (refs are
 * `<case>:<ref>`): the cross-case half of the gradient model (one blend fraction must hold for every fill and length).
 */
export function barGradientsAcrossCases(manifest: CfManifest, rootDir: string = REPO_ROOT): BarGradientObservation {
  const pooled: BarGradientObservation = { samples: [], unsampled: [] }
  for (const c of manifest.cases.filter(x => (x.stage === 3 || x.id.startsWith('stage3')) && x.status === 'measured')) {
    const capture = loadBarScreenCapture(c, rootDir, drawsBarGradient)
    if (!capture) continue
    const o = sampleBarGradients(c, capture.parts, capture.png, capture.geom)
    pooled.samples.push(...o.samples.map(s => ({ ...s, ref: `${c.id}:${s.ref}` })))
    pooled.unsampled.push(...o.unsampled.map(u => `${c.id}: ${u}`))
  }
  return pooled
}

/** The generated export of the verified gradient model (fixed key order, so `--write` stays idempotent). */
export function renderBarGradientBlock(gradient: BarGradientExport): string {
  const g = gradient['bar-length']
  const rows = g ? `  'bar-length': { endBlend: ${g.endBlend}, settledBy: [${g.settledBy.map(c => `'${c}'`).join(', ')}] },\n` : ''
  return `export const VERIFIED_STAGE3_BAR_GRADIENT: Readonly<Record<string, { endBlend: number; settledBy: readonly string[] }>> = {\n${rows}}\n`
}

/**
 * Registry entries the audit cannot honour: an unsupported (Ruling S3-CLOSE) id that is not a Stage 3 capability, has
 * no reason, or is also registered non-discriminating (one capability has one disposition). Any of them fails `ok`.
 */
export function unknownStage3Registrations(
  unsupported: Readonly<Record<string, string>> = CF_STAGE3_UNSUPPORTED,
  nonDiscriminating: Readonly<Record<string, string>> = CF_STAGE3_NONDISCRIMINATING,
): string[] {
  const ids = new Set(CF_STAGE3_CAPABILITIES.map(entry => entry.capabilityId))
  return Object.keys(unsupported).filter(id => !ids.has(id) || !unsupported[id]!.trim() || nonDiscriminating[id] !== undefined).sort()
}

/**
 * Stage 3 (data bars). Unlike Stages 1/2 a measured case can verify only some of its capabilities, so each capability's
 * disposition comes from `verifyCfCase`'s per-capability result and the cross-case gates (pixel calibration, precondition
 * probe), never from the case-level status: a case that verifies still reports every capability it did not
 * discriminate as `supported-structural` (blocks the stage) or `unverified-nondiscriminating` (Lesson 3).
 */
export function auditStage3Manifest(manifest: CfManifest, options: { rootDir?: string } = {}): Stage3AuditReport {
  const rootDir = options.rootDir ?? process.cwd()
  const stage3ManifestCases = manifest.cases.filter(c => c.stage === 3 || c.id.startsWith('stage3'))
  const catalogIds = CF_STAGE3_CASES.map(c => c.id)
  const manifestIds = new Set(stage3ManifestCases.map(c => c.id))
  const missingCases = catalogIds.filter(id => !manifestIds.has(id))
  const catalogSet = new Set(catalogIds)
  const extraCases = stage3ManifestCases.map(c => c.id).filter(id => !catalogSet.has(id))
  const perCase = stage3ManifestCases.map(c => auditCase(c, rootDir))
  const evidence = new Map(stage3ManifestCases.map(c => [c.id, verifyCfCase(c, { rootDir })] as const))
  const perCapability: Stage3CapabilityAudit[] = CF_STAGE3_CAPABILITIES.map(entry => {
    const owning = perCase.find(p => p.id === entry.caseId)
    let disposition: Stage3CapabilityAudit['disposition'] = 'unaccounted'
    if (owning?.structuralValid) {
      // Ruling S3-CLOSE: a registered capability of a measured case is unsupported in Stage 3, whatever its capture shows.
      if (owning.status === 'unsupported' || CF_STAGE3_UNSUPPORTED[entry.capabilityId] !== undefined) disposition = 'refused-unsupported'
      else {
        disposition = 'supported-structural'
        const res = evidence.get(entry.caseId)
        if (res?.kind === 'verified') {
          if (CF_STAGE3_NONDISCRIMINATING[entry.capabilityId] !== undefined || res.nonDiscriminatingCapabilities?.includes(entry.capabilityId)) {
            disposition = 'unverified-nondiscriminating'
          } else if (res.verifiedCapabilities.includes(entry.capabilityId) && isCapabilityVerified(entry.capabilityId, manifest, { rootDir })) {
            disposition = 'verified'
          }
        }
      }
    }
    return { capabilityId: entry.capabilityId, caseId: entry.caseId, disposition }
  })
  // An automatic colour verifies only when every capture that draws it agrees with the owning case's measurement.
  const pooled = perCapability.some(p => isBarAutoColourCapability(p.capabilityId) && p.disposition === 'verified')
    ? barAutoColoursAcrossCases(manifest, rootDir)
    : {}
  for (const p of perCapability) {
    if (!isBarAutoColourCapability(p.capabilityId) || p.disposition !== 'verified') continue
    const own = evidence.get(p.caseId)
    const owned = own?.kind === 'verified' ? own.matchedHypotheses?.[p.capabilityId] : undefined
    const judged = judgeBarAutoColour(pooled[p.capabilityId])
    if (judged.status !== 'discriminating' || judged.matched === undefined || owned === undefined || !sameArgbWithin(owned, judged.matched)) {
      p.disposition = 'supported-structural'
    }
  }
  const autoBarColours = verifiedAutoBarColours(perCapability, pooled)
  // The gradient model verifies only when every capture's gradient bars agree with the owning case's (one blend fraction
  // over every fill and length, the fixed-end alternative refuted by a second fill).
  const gradientEntry = perCapability.find(p => p.capabilityId === BAR_GRADIENT_CAPABILITY)
  const pooledGradient = judgeBarGradient(gradientEntry?.disposition === 'verified' ? barGradientsAcrossCases(manifest, rootDir) : undefined)
  const barGradient: BarGradientExport = {}
  if (gradientEntry?.disposition === 'verified') {
    const own = evidence.get(gradientEntry.caseId)
    const owned = own?.kind === 'verified' ? own.matchedHypotheses?.[BAR_GRADIENT_CAPABILITY] : undefined
    if (pooledGradient.status === 'discriminating' && pooledGradient.matched === 'bar-length' && owned === 'bar-length' && pooledGradient.endBlend !== undefined) {
      barGradient['bar-length'] = { endBlend: pooledGradient.endBlend, settledBy: [BAR_GRADIENT_CAPABILITY] }
    } else gradientEntry.disposition = 'supported-structural'
  }
  // Task 22b: the 10/90 rule is the owning capture's matched candidate, and only once the capability is verified.
  const barLengths: BarLengthsExport = {}
  const lengthsEntry = perCapability.find(p => p.capabilityId === 'bar:lengths-default-10-90')
  if (lengthsEntry?.disposition === 'verified') {
    const own = evidence.get(lengthsEntry.caseId)
    const rule = own?.kind === 'verified' ? own.matchedHypotheses?.['bar:lengths-default-10-90'] : undefined
    if (rule !== undefined && (BAR_LENGTH_RULES as readonly string[]).includes(rule)) barLengths['default-10-90'] = { rule, settledBy: ['bar:lengths-default-10-90'] }
    else lengthsEntry.disposition = 'supported-structural'
  }
  const matchedByCase = [...evidence.values()].flatMap(r => r.kind === 'verified' && r.matchedHypotheses ? [r.matchedHypotheses] : [])
  const barLayout = deriveBarLayout(perCapability, matchedByCase)
  const stage3Ids = new Set(CF_STAGE3_CAPABILITIES.map(entry => entry.capabilityId))
  const unknownNonDiscriminating = [
    ...Object.keys(CF_STAGE3_NONDISCRIMINATING).filter(id => !stage3Ids.has(id) || !CF_STAGE3_NONDISCRIMINATING[id]!.trim()),
    ...unknownStage3Registrations(),
  ]
  const caseOk = (p: Stage3CaseAudit): boolean => {
    if (p.status === 'measured') return p.structuralValid && p.code === undefined
    return p.structuralValid
  }
  const structuralValid = perCase.filter(p => p.structuralValid).length
  const ok = missingCases.length === 0 && perCase.length > 0 && perCase.every(caseOk) && perCapability.every(p => p.disposition !== 'unaccounted') && unknownNonDiscriminating.length === 0
  const gate = stageGateStatus(perCapability)
  return {
    autoBarColours,
    barLayout,
    barGradient,
    barGradientJudgement: `${pooledGradient.status}: ${pooledGradient.reason}`,
    barLengths,
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
      unverifiedCapabilities: perCapability.filter(p => p.disposition === 'unverified-nondiscriminating').length,
    },
  }
}

/**
 * The exact text `--stage N --write` puts in `src/xlsx/cf/verifiedCapabilities.ts`, as a pure function of the manifest
 * so the write is testable and idempotent (Lesson 4): the same manifest always renders the same bytes, and the Stage 1/2
 * sets never change when Stage 3 is rendered.
 */
export interface RenderVerifiedOptions {
  /** Ruling S3-WRITE: publish the audited Stage 3 set (Task 23 only). Without it the committed Stage 3 export stays as is. */
  publishStage3?: boolean
  /** The current `verifiedCapabilities.ts`, whose Stage 3 export is carried over verbatim unless `publishStage3`. */
  committedSource?: string
}

/** Per-channel ARGB agreement within the bar colour tolerance (8, the calibration cell's). */
const sameArgbWithin = (a: string, b: string, tol = 8): boolean => [2, 4, 6].every(k => Math.abs(parseInt(a.slice(k, k + 2), 16) - parseInt(b.slice(k, k + 2), 16)) <= tol)

const exportedColours = (text: string | undefined): string | undefined => {
  const m = text === undefined ? null : /export const VERIFIED_STAGE3_AUTO_BAR_COLORS: Readonly<Record<string, string>> = \{\n[\s\S]*?\n?\}\n/.exec(text)
  return m ? m[0] : undefined
}

const exportedLayout = (text: string | undefined): string | undefined => {
  const m = text === undefined ? null : /export const VERIFIED_STAGE3_BAR_LAYOUT: [^\n]* = \{\n[\s\S]*?\n?\}\n/.exec(text)
  return m ? m[0] : undefined
}

const exportedGradient = (text: string | undefined): string | undefined => {
  const m = text === undefined ? null : /export const VERIFIED_STAGE3_BAR_GRADIENT: [^\n]* = \{\n[\s\S]*?\n?\}\n/.exec(text)
  return m ? m[0] : undefined
}

const exportedLengths = (text: string | undefined): string | undefined => {
  const m = text === undefined ? null : /export const VERIFIED_STAGE3_BAR_LENGTHS: [^\n]* = \{\n[\s\S]*?\n?\}\n/.exec(text)
  return m ? m[0] : undefined
}

const exportedSet = (text: string | undefined, name: string): string[] | undefined => {
  const m = text === undefined ? null : new RegExp(`export const ${name}: ReadonlySet<string> = new Set\\(\\[([\\s\\S]*?)\\]\\)`).exec(text)
  return m ? [...m[1]!.matchAll(/'([^']+)'/g)].map(x => x[1]!) : undefined
}

/**
 * Every `--write`, whatever `--stage`, renders ALL stages' exports, each from its own audit (Task 18c fix 1: a
 * Stage-1-only rendering once dropped `VERIFIED_STAGE2_CAPABILITIES` and broke `evidence.ts`). Stage 3 is published only
 * with `publishStage3` (Ruling S3-WRITE, Task 23); otherwise its committed export, or its absence, is kept verbatim.
 */
export function renderVerifiedCapabilitiesSource(manifest: CfManifest, rootDir: string = REPO_ROOT, options: RenderVerifiedOptions = {}): string {
  const verifiedOf = (report: Stage1AuditReport): string[] => report.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()
  const stage1Verified = verifiedOf(auditStage1Manifest(manifest, { rootDir }))
  const stage2Verified = verifiedOf(auditStage2Manifest(manifest, { rootDir }))
  const stage3Report = options.publishStage3 ? auditStage3Manifest(manifest, { rootDir }) : undefined
  const stage3Verified = stage3Report
    ? verifiedOf(stage3Report)
    : exportedSet(options.committedSource, 'VERIFIED_STAGE3_CAPABILITIES')
  // Ruling S3-F as amended: the measured automatic colours travel only beside a published Stage 3 set.
  const autoColoursBlock = stage3Report ? renderAutoBarColoursBlock(stage3Report.autoBarColours) : exportedColours(options.committedSource)
  // Ruling S3-B: the selected layout travels the same way, only beside a published Stage 3 set.
  const layoutBlock = stage3Report ? renderBarLayoutBlock(stage3Report.barLayout.layout) : exportedLayout(options.committedSource)
  // Ruling S3-M (Task 22b): the verified gradient model travels the same way.
  const gradientBlock = stage3Report ? renderBarGradientBlock(stage3Report.barGradient) : exportedGradient(options.committedSource)
  const lengthsBlock = stage3Report ? renderBarLengthsBlock(stage3Report.barLengths) : exportedLengths(options.committedSource)
  const list = (ids: readonly string[]): string => ids.map(id => `  '${id}',`).join('\n')
  const allVerified = [...new Set([...stage1Verified, ...stage2Verified, ...(stage3Verified ?? [])])].sort()
  const block = (name: string, ids: readonly string[]): string => `export const ${name}: ReadonlySet<string> = new Set([\n${list(ids)}\n])\n`
  return `/**
 * Verified conditional formatting capabilities, every stage from its own audit (Stage 3 only once published, Ruling S3-WRITE).
 * Generated via bun scripts/cf-refs/audit.ts --write (any --stage; --publish-stage3 publishes Stage 3).
 * Checked by tests/xlsx-cf-stage1.test.ts, tests/xlsx-cf-stage2.test.ts and tests/xlsx-cf-oracle.test.ts.
 */
${block('VERIFIED_STAGE1_CAPABILITIES', stage1Verified)}
${block('VERIFIED_STAGE2_CAPABILITIES', stage2Verified)}
${stage3Verified !== undefined ? `${block('VERIFIED_STAGE3_CAPABILITIES', stage3Verified)}\n` : ''}${stage3Verified !== undefined && autoColoursBlock !== undefined ? `${autoColoursBlock}\n` : ''}${stage3Verified !== undefined && layoutBlock !== undefined ? `${layoutBlock}\n` : ''}${stage3Verified !== undefined && gradientBlock !== undefined ? `${gradientBlock}\n` : ''}${stage3Verified !== undefined && lengthsBlock !== undefined ? `${lengthsBlock}\n` : ''}${block('VERIFIED_CAPABILITIES', allVerified)}`
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

/** A 1M-row DECLARED A:A scope holding only 100 stored cells, requested by 100 scales that share one statistics scan
 * (Lesson 3). This proves the sparse scan (reads = stored cells, not 1M) and the shared scan; it is NOT a populated
 * 1M-row scope (final review M5). */
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
    name: 'scale-sparse-1M-declared-shared-statistics',
    outcome: 'admitted',
    peakBytes: peak,
    netBytes: memory.currentBytes,
    sourceReads: visited,
    cellsWithResults: stats.spanCells,
    wallMs: Date.now() - started,
  }
}

/** An invalid (out-of-range) percentile bound withholds its own scale (cf-invalid-bound) while a min/max scale over
 * another column still renders. It is NOT a sorted-copy reservation refusal: the cf-resource-limit refusal path is
 * covered in tests/xlsx-cf-statistics.test.ts (final review M5). */
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
    name: 'invalid-percentile-bound-withheld-with-min-max-rendered',
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
  if (stage !== 1 && stage !== 2 && stage !== 3) {
    // eslint-disable-next-line no-console
    console.error('Only --stage 1, --stage 2 and --stage 3 are implemented')
    process.exit(2)
  }
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  const outDir = args.includes('--out') ? resolve(args[args.indexOf('--out') + 1]!) : undefined
  const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
  const report = stage === 1
    ? auditStage1Manifest(manifest, { rootDir: REPO_ROOT })
    : stage === 2
      ? auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
      : auditStage3Manifest(manifest, { rootDir: REPO_ROOT })
  if (args.includes('--write')) {
    const target = resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts')
    const content = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, {
      committedSource: existsSync(target) ? readFileSync(target, 'utf8') : undefined,
      publishStage3: args.includes('--publish-stage3'),
    })
    writeFileSync(target, content, 'utf8')
    const count = (name: string): number => (new RegExp(`export const ${name}: ReadonlySet<string> = new Set\\(\\[([\\s\\S]*?)\\]\\)`).exec(content)?.[1]?.match(/'[^']+'/g) ?? []).length
    const s3 = content.includes('VERIFIED_STAGE3_CAPABILITIES') ? `, ${count('VERIFIED_STAGE3_CAPABILITIES')} S3${args.includes('--publish-stage3') ? ' published' : ' as committed'}` : ', S3 unpublished'
    const detail = ` (${count('VERIFIED_STAGE1_CAPABILITIES')} S1, ${count('VERIFIED_STAGE2_CAPABILITIES')} S2${s3})`
    // eslint-disable-next-line no-console
    console.log(`wrote ${count('VERIFIED_CAPABILITIES')} verified capabilities${detail} to src/xlsx/cf/verifiedCapabilities.ts`)
  }
  const cases = only ? report.perCase.filter(c => c.id === only) : report.perCase
  for (const entry of cases) printCase(entry)
  // eslint-disable-next-line no-console
  console.log(`manifest: ${report.summary.structuralValid}/${report.summary.cases} structural, capabilities ${report.summary.covered}/${report.summary.capabilities}, missing=[${report.missingCases.join(',')}] extra=[${report.extraCases.join(',')}]`)
  if ('barLayout' in report) {
    const l = (report as Stage3AuditReport).barLayout
    const show = (scope: BarLayoutScope): string => { const v = l.layout[scope]; return v ? `${v.origin}/${v.lengthBasis} (${v.settledBy.join(', ')})` : 'none' }
    // eslint-disable-next-line no-console
    console.log(`bar layout (S3-B): allPositive=${show('allPositive')} mixedSign=${show('mixedSign')}${l.findings.length ? ` findings=[${l.findings.join('; ')}]` : ''}`)
    const r3 = report as Stage3AuditReport
    const g = r3.barGradient['bar-length']
    // eslint-disable-next-line no-console
    console.log(`bar gradient (S3-M): ${g ? `bar-length endBlend=${g.endBlend}` : 'none'} (${r3.barGradientJudgement})`)
    const d = r3.barLengths['default-10-90']
    // eslint-disable-next-line no-console
    console.log(`bar 10/90 lengths (Task 22b): ${d ? `${d.rule} (${d.settledBy.join(', ')})` : 'none selected'}`)
  }
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
  console.log(`summary: ok=${ok} structural=${report.summary.structuralValid}/${report.summary.cases} unsupported=${report.summary.unsupported} pending=${report.summary.pending} unverifiedNonDiscriminating=${report.summary.unverifiedCapabilities ?? 0} stageComplete=${report.stageComplete} blockedCapabilities=${report.blockedCapabilities.length}`)

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
