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
import { CF_FORMULA_PRODUCTION_READINGS } from '../../src/xlsx/cf/formulaReadings'
import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  BAR_AUTO_COLOUR_CAPABILITIES,
  BAR_CAPABILITY_HYPOTHESES,
  BAR_GRADIENT_BORDER_CAPABILITY,
  BAR_GRADIENT_CAPABILITY,
  BAR_LENGTH_RULES,
  isCellInSqref,
  fitBarGradientBlend,
  judgeBarGradientBorder,
  drawsBarAutoColour,
  drawsBarGradient,
  isBarAutoColourCapability,
  isCapabilityVerified,
  judgeBarAutoColour,
  loadBarScreenCapture,
  measureCaseBars,
  measuredBarForms,
  readZipPartsSync,
  primeVerifiedCases,
  stage5GatesHold,
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
  CF_STAGE4_CAPABILITIES,
  CF_STAGE4_CASES,
  CF_STAGE4_NONDISCRIMINATING,
  CF_STAGE4_UNSUPPORTED,
  CF_STAGE5_CAPABILITIES,
  CF_STAGE5_CASES,
  CF_STAGE5_NONDISCRIMINATING,
  CF_STAGE5_REUSED_CAPTURES,
  CF_STAGE5_UNSUPPORTED,
  iconCapabilityKind,
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
import { CF_ICON_PAINT_READINGS } from '../../src/xlsx/cf/iconPaint'
import { CF_ICON_GLYPH_CLASSES, CF_ICON_GLYPH_WITHHELD } from '../../src/xlsx/cf/iconGlyphs'
import type { FormatBudget } from '../../src/xlsx/view/formatWorkspace'
import type { XlsxCell, XlsxSheet } from '../../src/xlsx/types'
import type { CfLease, CfStamp, CfValuePage, CfValueSource } from '../../src/xlsx/cf/model'

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
function auditCase(c: CfOracleCase, rootDir: string, verified?: ReturnType<typeof verifyCfCase>): Stage1CaseAudit {
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
  const res = verified ?? verifyCfCase(c, { rootDir })
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
  /** Ruling S3-GRAD-BORDER: the verified bordered-gradient shape and its blend (empty until it verifies). */
  barGradientBorder: BarGradientBorderExport
  /** Ruling S3-AUTO-BOUNDS (Task 22c): the verified rule of each automatic bound (empty until it verifies). */
  barAutoBounds: BarAutoBoundsExport
}

/** What an automatic bound resolves to: `zero-anchored` (min(0, v_min) / max(0, v_max)) or `min-max` (v_min / v_max). */
export type BarAutoBoundsExport = Partial<Record<'autoMin' | 'autoMax', { rule: string; settledBy: string[] }>>

/** The generated export of the verified automatic-bound rules (fixed key order, so `--write` stays idempotent). */
export function renderBarAutoBoundsBlock(auto: BarAutoBoundsExport): string {
  const rows = (['autoMin', 'autoMax'] as const).filter(k => auto[k]).map(k => `  ${k}: { rule: '${auto[k]!.rule}', settledBy: [${auto[k]!.settledBy.map(c => `'${c}'`).join(', ')}] },\n`).join('')
  return `export const VERIFIED_STAGE3_BAR_AUTO_BOUNDS: Readonly<Record<string, { rule: string; settledBy: readonly string[] }>> = {\n${rows}}\n`
}

/** The bordered-gradient shape production may draw (a `BAR_GRADIENT_BORDER_HYPOTHESES` id), with its fitted blend. */
export type BarGradientBorderExport = Record<string, { endBlend?: number; settledBy: string[] }>

/** The generated export of the verified bordered-gradient shape (one key at most, so `--write` stays idempotent). */
export function renderBarGradientBorderBlock(border: BarGradientBorderExport): string {
  const rows = Object.keys(border).sort().map(shape => {
    const v = border[shape]!
    return `  '${shape}': { ${v.endBlend !== undefined ? `endBlend: ${v.endBlend}, ` : ''}settledBy: [${v.settledBy.map(c => `'${c}'`).join(', ')}] },\n`
  }).join('')
  return `export const VERIFIED_STAGE3_BAR_GRADIENT_BORDER: Readonly<Record<string, { endBlend?: number; settledBy: readonly string[] }>> = {\n${rows}}\n`
}

/** The length rule production may apply to 10/90 bars: a `BAR_LENGTH_RULES` id matched by the owning capture. */
export type BarLengthsExport = Partial<Record<'default-10-90', { rule: string; measured: string[]; settledBy: string[] }>>

/** The generated export of the verified 10/90 length rule (fixed key order, so `--write` stays idempotent). */
export function renderBarLengthsBlock(lengths: BarLengthsExport): string {
  const r = lengths['default-10-90']
  const list = (xs: readonly string[]): string => `[${xs.map(c => `'${c}'`).join(', ')}]`
  const rows = r ? `  'default-10-90': { rule: '${r.rule}', measured: ${list(r.measured)}, settledBy: ${list(r.settledBy)} },\n` : ''
  return `export const VERIFIED_STAGE3_BAR_LENGTHS: Readonly<Record<string, { rule: string; measured: readonly string[]; settledBy: readonly string[] }>> = {\n${rows}}\n`
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
  // Each case is verified once per audit: the case audit, the per-capability evidence and the cross-case gates share it.
  const evidence = new Map(stage3ManifestCases.map(c => [c.id, verifyCfCase(c, { rootDir })] as const))
  primeVerifiedCases(manifest, { rootDir }, evidence)
  const perCase = stage3ManifestCases.map(c => auditCase(c, rootDir, evidence.get(c.id)))
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
  // Ruling S3-GRAD-BORDER: the owning capture's matched shape, recomputed with its fitted blend, only once verified.
  const barGradientBorder: BarGradientBorderExport = {}
  const borderEntry = perCapability.find(p => p.capabilityId === BAR_GRADIENT_BORDER_CAPABILITY)
  if (borderEntry?.disposition === 'verified') {
    const own = evidence.get(borderEntry.caseId)
    const owned = own?.kind === 'verified' ? own.matchedHypotheses?.[BAR_GRADIENT_BORDER_CAPABILITY] : undefined
    const owningCase = manifest.cases.find(c => c.id === borderEntry.caseId)
    const capture = owningCase ? loadBarScreenCapture(owningCase, rootDir) : undefined
    const refs = owningCase ? new Set(owningCase.expected.cells.filter(cell => cell.matched && owningCase.markers.some(m => m.capabilityId === BAR_GRADIENT_BORDER_CAPABILITY && isCellInSqref(cell.ref, m.sqref))).map(cell => cell.ref.toUpperCase())) : undefined
    const blend = owningCase && capture ? fitBarGradientBlend(sampleBarGradients(owningCase, capture.parts, capture.png, capture.geom).samples) : undefined
    const judged = owningCase && capture ? judgeBarGradientBorder(sampleBarGradients(owningCase, capture.parts, capture.png, capture.geom, refs, 'bordered'), blend) : undefined
    if (judged?.status === 'discriminating' && judged.matched !== undefined && judged.matched === owned) {
      barGradientBorder[judged.matched] = { ...(judged.endBlend !== undefined ? { endBlend: judged.endBlend } : {}), settledBy: [BAR_GRADIENT_BORDER_CAPABILITY] }
    } else borderEntry.disposition = 'supported-structural'
  }
  // Task 22b: the 10/90 rule is the owning capture's matched candidate, and only once the capability is verified.
  const barLengths: BarLengthsExport = {}
  const lengthsEntry = perCapability.find(p => p.capabilityId === 'bar:lengths-default-10-90')
  if (lengthsEntry?.disposition === 'verified') {
    const own = evidence.get(lengthsEntry.caseId)
    const rule = own?.kind === 'verified' ? own.matchedHypotheses?.['bar:lengths-default-10-90'] : undefined
    // Task 22c (review I2): the rule travels with the pairing form x fill type combinations its capture measured.
    const owningCase = manifest.cases.find(c => c.id === lengthsEntry.caseId)
    const savedParts = owningCase && existsSync(resolve(rootDir, owningCase.saved)) ? readZipPartsSync(new Uint8Array(readFileSync(resolve(rootDir, owningCase.saved)))) : null
    const measured = owningCase && savedParts ? measuredBarForms(owningCase, savedParts, 'bar:lengths-default-10-90') : []
    if (rule !== undefined && (BAR_LENGTH_RULES as readonly string[]).includes(rule) && measured.length > 0) barLengths['default-10-90'] = { rule, measured, settledBy: ['bar:lengths-default-10-90'] }
    else lengthsEntry.disposition = 'supported-structural'
  }
  // Ruling S3-AUTO-BOUNDS (Task 22c): each automatic bound's matched rule, only once verified and only while the 10/90
  // rule its hypotheses were predicted under is the one exported (else the bound falls back to supported-structural).
  const barAutoBounds: BarAutoBoundsExport = {}
  for (const [cap, key] of [['bar:bounds-automin', 'autoMin'], ['bar:bounds-automax', 'autoMax']] as const) {
    const entry = perCapability.find(p => p.capabilityId === cap)
    if (entry?.disposition !== 'verified') continue
    const own = evidence.get(entry.caseId)
    const rule = own?.kind === 'verified' ? own.matchedHypotheses?.[cap] : undefined
    const assumed = BAR_CAPABILITY_HYPOTHESES[cap]?.knobs?.lengthRule
    const lengthsOk = assumed === undefined || barLengths['default-10-90']?.rule === assumed
    if (rule !== undefined && (rule === 'zero-anchored' || rule === 'min-max') && lengthsOk) barAutoBounds[key] = { rule, settledBy: [cap] }
    else entry.disposition = 'supported-structural'
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
    barGradientBorder,
    barAutoBounds,
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

export type Stage4CaseAudit = Stage1CaseAudit
export type Stage4CapabilityAudit = Stage1CapabilityAudit
/** Stage 4 (icon sets): the per-capability dispositions, like Stage 3's, plus the recorded object-model probe answer. */
export interface Stage4AuditReport extends Stage1AuditReport {
  /** The `stage4-icon-om-reachability` sidecar's recorded rows (Ruling S4-A, record only), when captured. */
  iconOmReachability?: unknown
  /** The reading Excel matched for each capability whose owning case verified it with a named hypothesis. */
  matchedReadings: Readonly<Record<string, string>>
}

/**
 * Stage 4 (icon sets, Task 24), following `auditStage3Manifest`. A capability's disposition comes from `verifyCfCase`'s
 * per-capability result and the cross-case gates in `isCapabilityVerified` (pixel calibration, the glyph key, the family
 * order), never from the case-level status. Record-only and registry non-discriminating capabilities are
 * `unverified-nondiscriminating`; the unresolved shapes are `refused-unsupported`.
 */
export function auditStage4Manifest(manifest: CfManifest, options: { rootDir?: string } = {}): Stage4AuditReport {
  const rootDir = options.rootDir ?? process.cwd()
  const cases = manifest.cases.filter(c => c.stage === 4 || c.id.startsWith('stage4'))
  const catalogIds = CF_STAGE4_CASES.map(c => c.id)
  const manifestIds = new Set(cases.map(c => c.id))
  const missingCases = catalogIds.filter(id => !manifestIds.has(id))
  const catalogSet = new Set(catalogIds)
  const extraCases = cases.map(c => c.id).filter(id => !catalogSet.has(id))
  const evidence = new Map(cases.map(c => [c.id, verifyCfCase(c, { rootDir })] as const))
  primeVerifiedCases(manifest, { rootDir }, evidence)
  const perCase = cases.map(c => auditCase(c, rootDir, evidence.get(c.id)))
  const perCapability: Stage4CapabilityAudit[] = CF_STAGE4_CAPABILITIES.map(entry => {
    const owning = perCase.find(p => p.id === entry.caseId)
    let disposition: Stage4CapabilityAudit['disposition'] = 'unaccounted'
    if (owning?.structuralValid) {
      if (owning.status === 'unsupported' || CF_STAGE4_UNSUPPORTED[entry.capabilityId] !== undefined) disposition = 'refused-unsupported'
      else if (CF_STAGE4_NONDISCRIMINATING[entry.capabilityId] !== undefined || iconCapabilityKind(entry.capabilityId) === 'record') disposition = 'unverified-nondiscriminating'
      else {
        disposition = 'supported-structural'
        const res = evidence.get(entry.caseId)
        if (res?.kind === 'verified') {
          if (res.nonDiscriminatingCapabilities?.includes(entry.capabilityId)) disposition = 'unverified-nondiscriminating'
          else if (res.verifiedCapabilities.includes(entry.capabilityId) && isCapabilityVerified(entry.capabilityId, manifest, { rootDir })) disposition = 'verified'
        }
      }
    }
    return { capabilityId: entry.capabilityId, caseId: entry.caseId, disposition }
  })
  const matchedReadings: Record<string, string> = {}
  for (const p of perCapability) {
    const res = evidence.get(p.caseId)
    const matched = p.disposition === 'verified' && res?.kind === 'verified' ? res.matchedHypotheses?.[p.capabilityId] : undefined
    if (typeof matched === 'string') matchedReadings[p.capabilityId] = matched
  }
  const ids = new Set(CF_STAGE4_CAPABILITIES.map(e => e.capabilityId))
  const unknownNonDiscriminating = Object.keys(CF_STAGE4_NONDISCRIMINATING).filter(id => !ids.has(id) || !CF_STAGE4_NONDISCRIMINATING[id]!.trim())
  const caseOk = (p: Stage4CaseAudit): boolean => (p.status === 'measured' ? p.structuralValid && p.code === undefined : p.structuralValid)
  const ok = missingCases.length === 0 && perCase.length > 0 && perCase.every(caseOk) && perCapability.every(p => p.disposition !== 'unaccounted') && unknownNonDiscriminating.length === 0
  const gate = stageGateStatus(perCapability)
  const omCase = cases.find(c => c.id === 'stage4-icon-om-reachability')
  const omSidecar = omCase?.captures[0]?.sidecarPath ? resolve(rootDir, omCase.captures[0].sidecarPath) : undefined
  let iconOmReachability: unknown
  if (omSidecar && existsSync(omSidecar)) {
    try {
      iconOmReachability = (JSON.parse(readFileSync(omSidecar, 'utf8')) as Record<string, unknown>).iconOmReachability
    } catch {
      iconOmReachability = undefined
    }
  }
  return {
    ok,
    perCase,
    perCapability,
    missingCases,
    extraCases,
    stageComplete: ok && gate.complete,
    blockedCapabilities: gate.blocked,
    ...(iconOmReachability !== undefined ? { iconOmReachability } : {}),
    matchedReadings,
    summary: {
      cases: perCase.length,
      structuralValid: perCase.filter(p => p.structuralValid).length,
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

/* ---------------- Stage 4 publication gate (Task 30) ---------------- */

/**
 * Verified Stage 4 capabilities that are never published, each by its ruling: production does not implement Excel's
 * matched reading, or the views would disagree. Publishing one would claim a behaviour production lacks.
 */
export const CF_STAGE4_NEVER_PUBLISH: Readonly<Record<string, string>> = {
  'icon:scope-merged-cell':
    'Ruling S4-MERGED-UNPUBLISHED: the worker and forced-paged scopes carry no merge knowledge, so publishing would make the resident view paint icons the paged views withhold; merged scopes keep withholding.',
  'icon:geometry-short-row-clip':
    'Ruling S4-SHORTROW-WITHHELD: Excel shrinks the box in a short row (matched `shrink`) by a law the captures do not determine; production does not implement shrink, so short rows withhold.',
  'icon:inset-overflow':
    'Excel draws the icon and a `#` fill (matched `icon-and-text`), whose layout is unmeasured; production withholds every overflowing host.',
}

/**
 * The reading production implements for each Stage 4 capability whose owning case verifies a named hypothesis: the
 * geometry, inset and showValue readings are `CF_ICON_PAINT_READINGS` (`src/xlsx/cf/iconPaint.ts`, pinned to behaviour in
 * tests/xlsx-cf-icon-align.test.ts); the rest are the evaluator's and composer's (`iconSet.ts`, `compose.ts`,
 * `service.ts`, `parse.ts`), checked against Excel's own workbooks by the no-injection bucket conformance in
 * tests/xlsx-cf-oracle-conformance.test.ts. A list is a selection: production reads the matched hypothesis from
 * `VERIFIED_STAGE4_ICON_SELECTIONS` and implements exactly the listed ones (tests/xlsx-cf-stage4.test.ts pins each).
 * A capability with a matched reading but no entry here is withheld (fail closed).
 */
export const CF_STAGE4_PRODUCTION_READINGS: Readonly<Record<string, string | readonly string[]>> = {
  ...CF_ICON_PAINT_READINGS,
  // parse.ts: the base and x14 dialects parse to one model (`as-modelled`), with the schema defaults applied.
  'icon:dialect-standalone': 'as-modelled',
  'icon:dialect-base': 'as-modelled',
  'icon:settings-schema-defaults': 'as-modelled',
  // iconSet.ts `evaluateIconSet`: bucket i shows iconId i (k-1-i under reverse); a custom list maps bucket i to entry i;
  // a NoIcons entry draws nothing.
  'icon:families-base': 'bucket-order',
  'icon:families-x14': 'bucket-order',
  'icon:reverse': 'reversed',
  'icon:custom-mapping': 'custom-order',
  'icon:custom-noicons': 'no-icon',
  // iconSet.ts `resolveIconThresholds`: num literal, percent of [min, max], per-type mixing, gte as written.
  'icon:thresholds-num': 'literal',
  'icon:thresholds-percent': 'percent-of-range',
  'icon:thresholds-mixed': 'per-type',
  'icon:thresholds-gte-true': 'as-written',
  'icon:thresholds-gte-false': 'as-written',
  // Selections (`iconSelectionEvidence`): each branch production implements.
  'icon:thresholds-percentile-inc-exc': ['inc', 'exc'],
  'icon:thresholds-floor': ['below-floor-no-icon', 'below-floor-bucket0'],
  'icon:degenerate-equal-values': ['as-model'],
  'icon:degenerate-single-cell': ['as-model'],
  'icon:degenerate-non-monotonic': ['as-model'],
  'icon:bounds-out-of-domain': ['as-model'],
  'icon:composition-noicons-slot': ['first-paints', 'second-paints'],
  'icon:reverse-custom': ['reverse-then-map', 'map-then-mirror', 'ignored'],
  // iconSet.ts / service.ts participation: non-numbers draw no icon, date serials evaluate, error cells are ignored by
  // the statistics (as for bars, bar:participation-error-presence).
  'icon:participation-types': 'no-icon-for-non-numbers',
  'icon:participation-dateserial': 'evaluated',
  'icon:participation-error-presence': 'error-ignored',
  // service.ts `blockScope` and the shared statistics (as for bars): a multi-range rule is one union scope, hidden and
  // autofilter-hidden rows stay in it, a whole-column rule scopes its stored cells.
  'icon:scope-whole-column': 'whole-column',
  'icon:scope-disjoint': 'union',
  'icon:scope-hidden-row': 'include-hidden',
  'icon:scope-autofilter-hidden-row': 'include-hidden',
  // compose.ts: of two matching icons (or duplicate icon blocks) the first in priority order paints; an icon and a bar
  // both paint; a matching stopIfTrue rule ahead halts the icon. (`icon:composition-unresolved-stub` matched `halted`,
  // but production cannot evaluate the Stage 5 expression stub and withholds every cell below it: no entry, so the
  // gate withholds it, as Stage 3 did for bar:composition-unresolved-stub.)
  'icon:scope-duplicate-blocks': 'first-paints',
  'icon:composition-overlapping-icons': 'first-paints',
  'icon:composition-with-bar': 'both-paint',
  'icon:composition-matching-stopIfTrue-ahead': 'halted',
}

export interface Stage4Publication {
  /** The ids `--publish-stage4` writes to `VERIFIED_STAGE4_CAPABILITIES` (sorted). */
  published: string[]
  /** Every other Stage 4 id, with the reason it stays withheld. */
  withheld: Record<string, string>
  /** `VERIFIED_STAGE4_ICON_SELECTIONS`: the matched reading of each published selection capability. */
  selections: Record<string, string>
}

/**
 * The Stage 4 publication gate (Task 30; carried from the Task 25/26 review and Ruling S4-PROD-ALIGN): a capability
 * publishes only when it is verified, no ruling withholds it, the reading Excel matched is the one production
 * implements, and (a glyph capability) its glyph class passed the OD-9(b) approximation check. Pure in the report.
 */
export function stage4Publication(report: Stage4AuditReport): Stage4Publication {
  const published: string[] = []
  const withheld: Record<string, string> = {}
  const selections: Record<string, string> = {}
  for (const p of report.perCapability) {
    const id = p.capabilityId
    if (p.disposition !== 'verified') { withheld[id] = p.disposition; continue }
    if (CF_STAGE4_NEVER_PUBLISH[id] !== undefined) { withheld[id] = `never-publish: ${CF_STAGE4_NEVER_PUBLISH[id]}`; continue }
    const glyph = /^icon:glyph-(\d[A-Za-z0-9]+)-(\d)$/.exec(id)
    if (glyph) {
      const cls = Object.prototype.hasOwnProperty.call(CF_ICON_GLYPH_CLASSES, `${glyph[1]}:${glyph[2]}`) ? CF_ICON_GLYPH_CLASSES[`${glyph[1]}:${glyph[2]}`] : undefined
      if (cls === undefined) { withheld[id] = 'glyph-unknown-pair'; continue }
      if (CF_ICON_GLYPH_WITHHELD.has(cls)) { withheld[id] = `glyph-approximation-failed: class ${cls} is in CF_ICON_GLYPH_WITHHELD`; continue }
    }
    const matched = report.matchedReadings[id]
    if (matched !== undefined) {
      const prod = Object.prototype.hasOwnProperty.call(CF_STAGE4_PRODUCTION_READINGS, id) ? CF_STAGE4_PRODUCTION_READINGS[id] : undefined
      if (prod === undefined) { withheld[id] = `reading-unimplemented: Excel matched ${matched}, production names no reading`; continue }
      if (typeof prod === 'string') {
        if (prod !== matched) { withheld[id] = `reading-mismatch: Excel matched ${matched}, production implements ${prod}`; continue }
      } else {
        if (!prod.includes(matched)) { withheld[id] = `reading-unimplemented: Excel matched ${matched}, production implements ${prod.join(' | ')}`; continue }
        selections[id] = matched
      }
    }
    published.push(id)
  }
  published.sort()
  return { published, withheld, selections }
}

/* ---------------- Stage 5 audit and publication gate (Task 31) ---------------- */

export type Stage5CaseAudit = Stage1CaseAudit
export type Stage5CapabilityAudit = Stage1CapabilityAudit
/** Stage 5 (formulas): the per-capability dispositions, like Stage 4's, plus the reused Stage 1-4 captures (S5-STUB). */
export interface Stage5AuditReport extends Stage1AuditReport {
  /** The Produces name of `blockedCapabilities` (the same list). */
  blocked: readonly string[]
  /** The reading Excel matched for each verified capability (the hypothesis its configurations selected). */
  matchedReadings: Readonly<Record<string, string>>
  /** `CF_STAGE5_REUSED_CAPTURES` entry (`caseId:capabilityId`) -> whether its measured capture evidences every ref. */
  reused: Readonly<Record<string, 'evidenced' | 'pending'>>
}

/**
 * Fix round 1 #7: whether a reused Stage 1-4 case is verified and its capture records evidence at each named cell, read in
 * that stage's own representation (no Stage 1-4 fixture is rewritten):
 * - Stage 1/2: the sidecar's `observedCells` row; an Excel DisplayFormat fill must equal the expected fill, otherwise the
 *   Stage 1 pixel contrast against the control (>= 16, the verifier's own threshold) must hold for a matched cell;
 * - Stage 3: the sidecar's `observedBars` row (a measured bar, or a measured absence);
 * - Stage 4: the sidecar's `observedIcons` row, or a `pixel-classified` bucket on the cell.
 */
export function reusedCaptureEvidenced(manifest: CfManifest, r: (typeof CF_STAGE5_REUSED_CAPTURES)[number], rootDir: string): boolean {
  const c = manifest.cases.find(x => x.id === r.caseId)
  if (!c || c.status !== 'measured' || verifyCfCase(c, { rootDir }).kind !== 'verified') return false
  const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
  let sidecar: Record<string, unknown> = {}
  try {
    sidecar = cap ? (JSON.parse(readFileSync(resolve(rootDir, cap.sidecarPath), 'utf8')) as Record<string, unknown>) : {}
  } catch {
    return false
  }
  const rows = (key: string): Array<Record<string, unknown>> => (Array.isArray(sidecar[key]) ? (sidecar[key] as Array<Record<string, unknown>>) : [])
  const rowOf = (key: string, ref: string): Record<string, unknown> | undefined => rows(key).find(o => typeof o?.ref === 'string' && o.ref.toUpperCase() === ref.toUpperCase())
  return r.refs.every(ref => {
    const e = c.expected.cells.find(x => x.ref === ref)
    if (!e) return false
    if (c.stage === 1 || c.stage === 2) {
      const o = rowOf('observedCells', ref)
      if (!o) return false
      if (o.fillSource === 'excel-display-format') return e.fillRgb === undefined || String(o.fillRgb).toUpperCase() === e.fillRgb.toUpperCase()
      return e.matched === true && typeof o.bgContrastDelta === 'number' && o.bgContrastDelta >= 16
    }
    if (c.stage === 3) return rowOf('observedBars', ref) !== undefined
    if (c.stage === 4) return rowOf('observedIcons', ref) !== undefined || e.iconEvidence?.bucket === 'pixel-classified'
    return false
  })
}

/**
 * Stage 5 (Task 31), following `auditStage4Manifest`. A capability's disposition comes from `verifyCfCase`'s
 * per-capability semantic evaluation and the Stage 5 gates in `isCapabilityVerified`, never from the case-level status.
 * Unsupported shapes are `refused-unsupported`; twin registries are `unverified-nondiscriminating`; a capability that
 * re-reads a Stage 1-4 capture (S5-STUB) also needs that capture to evidence its cells.
 */
export function auditStage5Manifest(manifest: CfManifest, options: { rootDir?: string } = {}): Stage5AuditReport {
  const rootDir = options.rootDir ?? process.cwd()
  const cases = manifest.cases.filter(c => c.stage === 5 || c.id.startsWith('stage5'))
  const catalogIds = CF_STAGE5_CASES.map(c => c.id)
  const manifestIds = new Set(cases.map(c => c.id))
  const missingCases = manifest.cases.length === 0 ? [] : catalogIds.filter(id => !manifestIds.has(id))
  const catalogSet = new Set(catalogIds)
  const extraCases = cases.map(c => c.id).filter(id => !catalogSet.has(id))
  const evidence = new Map(cases.map(c => [c.id, verifyCfCase(c, { rootDir })] as const))
  primeVerifiedCases(manifest, { rootDir }, evidence)
  const perCase = cases.map(c => auditCase(c, rootDir, evidence.get(c.id)))
  const reused: Record<string, 'evidenced' | 'pending'> = {}
  for (const r of CF_STAGE5_REUSED_CAPTURES) reused[`${r.caseId}:${r.capabilityId}`] = reusedCaptureEvidenced(manifest, r, rootDir) ? 'evidenced' : 'pending'
  const reuseHolds = (cap: string): boolean => CF_STAGE5_REUSED_CAPTURES.filter(r => r.capabilityId === cap).every(r => reused[`${r.caseId}:${r.capabilityId}`] === 'evidenced')
  const ownedByCase = (cap: string): boolean => CF_STAGE5_CASES.some(c => c.markers.some(m => m.capabilityId === cap) || (c.formulaConfigurations ?? []).some(x => x.capabilityId === cap))
  const perCapability: Stage5CapabilityAudit[] = (manifest.cases.length === 0 ? [] : CF_STAGE5_CAPABILITIES).map(entry => {
    const owning = perCase.find(p => p.id === entry.caseId)
    let disposition: Stage5CapabilityAudit['disposition'] = 'unaccounted'
    if (owning?.structuralValid) {
      if (owning.status === 'unsupported' || CF_STAGE5_UNSUPPORTED[entry.capabilityId] !== undefined) disposition = 'refused-unsupported'
      else if (CF_STAGE5_NONDISCRIMINATING[entry.capabilityId] !== undefined) disposition = 'unverified-nondiscriminating'
      // A reuse-only capability (S5-STUB) rests on the Stage 1-4 capture and the Stage 5 gates (formula:truth-boolean).
      else if (!ownedByCase(entry.capabilityId)) disposition = reuseHolds(entry.capabilityId) && stage5GatesHold(entry.capabilityId, manifest, { rootDir }) ? 'verified' : 'supported-structural'
      else {
        disposition = 'supported-structural'
        const res = evidence.get(entry.caseId)
        if (res?.kind === 'verified') {
          if (res.nonDiscriminatingCapabilities?.includes(entry.capabilityId)) disposition = 'unverified-nondiscriminating'
          else if (res.verifiedCapabilities.includes(entry.capabilityId) && isCapabilityVerified(entry.capabilityId, manifest, { rootDir }) && reuseHolds(entry.capabilityId)) disposition = 'verified'
        }
      }
    }
    return { capabilityId: entry.capabilityId, caseId: entry.caseId, disposition }
  })
  const matchedReadings: Record<string, string> = {}
  for (const p of perCapability) {
    if (p.disposition !== 'verified') continue
    const res = evidence.get(p.caseId)
    // A reuse-only capability (S5-STUB) has no configuration in its owning case, so its reading is always the recorded
    // Stage 1-4 outcome, whether or not that case's own capture verifies (Task 39: the composition rebuild verified it).
    const matched = !ownedByCase(p.capabilityId) ? 'recorded' : res?.kind === 'verified' ? res.matchedHypotheses?.[p.capabilityId] : undefined
    if (typeof matched === 'string') matchedReadings[p.capabilityId] = matched
  }
  const ids = new Set(CF_STAGE5_CAPABILITIES.map(e => e.capabilityId))
  const unknownRegistrations = [...Object.keys(CF_STAGE5_NONDISCRIMINATING), ...Object.keys(CF_STAGE5_UNSUPPORTED)].filter(id => !ids.has(id))
  const caseOk = (p: Stage5CaseAudit): boolean => (p.status === 'measured' ? p.structuralValid && p.code === undefined : p.structuralValid)
  const ok = missingCases.length === 0 && perCase.every(caseOk) && perCapability.every(p => p.disposition !== 'unaccounted') && unknownRegistrations.length === 0
  const gate = stageGateStatus(perCapability)
  return {
    ok,
    perCase,
    perCapability,
    missingCases,
    extraCases,
    stageComplete: ok && perCase.length > 0 && gate.complete,
    blockedCapabilities: gate.blocked,
    blocked: gate.blocked,
    matchedReadings,
    reused,
    summary: {
      cases: perCase.length,
      structuralValid: perCase.filter(p => p.structuralValid).length,
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
 * The reading production implements for each Stage 5 capability (the Stage 4 `CF_STAGE4_PRODUCTION_READINGS` pattern):
 * `CF_FORMULA_PRODUCTION_READINGS` (`src/xlsx/cf/formulaReadings.ts`), the one table production gates on, added by Task 35
 * together with the code and tests that implement each reading (tests/xlsx-cf-formula-rules.test.ts pins every entry to a
 * catalog hypothesis). A list is a selection, read from `VERIFIED_STAGE5_FORMULA_SELECTIONS`, which `--publish-stage5` renders
 * beside the published set (`renderFormulaSelectionsBlock`, Task 39).
 * A capability with a matched reading but no entry (`formula:truth-range`, `formula:fn-today`, ...) is withheld.
 */
export const CF_STAGE5_PRODUCTION_READINGS: Readonly<Record<string, string | readonly string[]>> = { ...CF_FORMULA_PRODUCTION_READINGS }

/**
 * Task 35 fix round 2 (ruling a): capabilities whose Stage 5 fixture production cannot confirm, with the only witnesses
 * that may fail there. The publication gate refuses them whatever the audit verified (tests/xlsx-cf-formula-readings.test.ts).
 */
export const CF_STAGE5_UNCONFIRMABLE: Readonly<Record<string, Readonly<{ reason: string; witnesses: readonly string[]
  /** Task 37 fix round 1: production withholds the configuration's designated cells (fail closed) rather than painting them,
   * so the narrow check accepts withheld cells; it still refuses any painted cell that contradicts the reading. */
  designatedWithheld?: true }>>> = {
  'bar:bounds-formula': {
    reason: 'stage5-cfvo-formulas cfvo-bar-formula: the bar column holds 5..100 against formula bounds 20..80, a whole-scope cf-value-beyond-bounds (Ruling S3-E), and its num twin W needs bar:bounds-num (CF_STAGE3_NONDISCRIMINATING), so production cannot reproduce the fixture (Task 37 fix round 1)',
    witnesses: ['W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8', 'W9'], designatedWithheld: true,
  },
  'formula:active-present-bar': {
    reason: 'stage5-active-present-bar: a standalone x14 bar, and bar:pairing-standalone is CF_STAGE3_UNSUPPORTED, so the activated bar and its twin withhold in production (Task 37 fix round 1)',
    witnesses: ['C2', 'C3', 'C4', 'C5', 'C6', 'C7'], designatedWithheld: true,
  },
  'formula:composition-stop-icon': {
    reason: "stage5-composition stop-icon: its continues witnesses L3/L5/L7 (a red fill plus an icon; column K before Task 39's re-lettering) are withheld in production while icon:composition-dxf-fill is unpublished, so production cannot confirm the fixture's reading (S5-WITNESS)",
    witnesses: ['L3', 'L5', 'L7'],
  },
}

export interface Stage5Publication {
  /** The ids `--publish-stage5` writes to `VERIFIED_STAGE5_CAPABILITIES` (sorted). */
  published: string[]
  /** Every other Stage 5 id, with the reason it stays withheld. */
  withheld: Record<string, string>
  /** The matched reading of each published selection capability (a list entry in the production readings). */
  selections: Record<string, string>
}

/**
 * The Stage 5 publication gate: a capability publishes only when it is verified and the reading Excel matched is the one
 * production implements (`readings`, by default `CF_STAGE5_PRODUCTION_READINGS`). Pure in the report.
 */
export function stage5Publication(report: Stage5AuditReport, readings: Readonly<Record<string, string | readonly string[]>> = CF_STAGE5_PRODUCTION_READINGS): Stage5Publication {
  const published: string[] = []
  const withheld: Record<string, string> = {}
  const selections: Record<string, string> = {}
  for (const p of report.perCapability) {
    const id = p.capabilityId
    if (p.disposition !== 'verified') { withheld[id] = p.disposition; continue }
    const unconfirmable = CF_STAGE5_UNCONFIRMABLE[id]
    if (unconfirmable) { withheld[id] = `fixture-unconfirmable: ${unconfirmable.reason}`; continue }
    const matched = report.matchedReadings[id]
    const prod = Object.prototype.hasOwnProperty.call(readings, id) ? readings[id] : undefined
    if (matched === undefined) { withheld[id] = 'reading-unrecorded: a verified Stage 5 capability names the reading it matched'; continue }
    if (prod === undefined) { withheld[id] = `reading-unimplemented: Excel matched ${matched}, production names no reading`; continue }
    if (typeof prod === 'string') {
      if (prod !== matched) { withheld[id] = `reading-mismatch: Excel matched ${matched}, production implements ${prod}`; continue }
    } else {
      if (!prod.includes(matched)) { withheld[id] = `reading-unimplemented: Excel matched ${matched}, production implements ${prod.join(' | ')}`; continue }
      selections[id] = matched
    }
    published.push(id)
  }
  published.sort()
  return { published, withheld, selections }
}

/** `VERIFIED_STAGE4_ICON_SELECTIONS` as `--publish-stage4` renders it (keys sorted, so the write is idempotent). */
export function renderIconSelectionsBlock(selections: Readonly<Record<string, string>>): string {
  const keys = Object.keys(selections).sort()
  return `export const VERIFIED_STAGE4_ICON_SELECTIONS: Readonly<Record<string, string>> = {\n${keys.map(k => `  '${k}': '${selections[k]}',`).join('\n')}\n}\n`
}

/** `VERIFIED_STAGE5_FORMULA_SELECTIONS` as `--publish-stage5` renders it beside the Stage 5 set (keys sorted, idempotent). */
export function renderFormulaSelectionsBlock(selections: Readonly<Record<string, string>>): string {
  const keys = Object.keys(selections).sort()
  return `export const VERIFIED_STAGE5_FORMULA_SELECTIONS: Readonly<Record<string, string>> = {\n${keys.map(k => `  '${k}': '${selections[k]}',`).join('\n')}\n}\n`
}

const exportedFormulaSelections = (text: string | undefined): string | undefined => {
  const m = text === undefined ? null : /export const VERIFIED_STAGE5_FORMULA_SELECTIONS: [^\n]* = \{\n[\s\S]*?\n?\}\n/.exec(text)
  return m ? m[0] : undefined
}

const exportedSelections = (text: string | undefined): string | undefined => {
  const m = text === undefined ? null : /export const VERIFIED_STAGE4_ICON_SELECTIONS: [^\n]* = \{\n[\s\S]*?\n?\}\n/.exec(text)
  return m ? m[0] : undefined
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
  /** Mirrors Ruling S3-WRITE for Stage 4: publish the audited Stage 4 set (Task 30 only); otherwise the committed Stage 4 export, or its absence, is kept verbatim. */
  publishStage4?: boolean
  /**
   * Mirrors S3-WRITE and `publishStage4` for Stage 5: publish the gated Stage 5 set (`stage5Publication`), and only when the
   * Stage 5 audit reports `stageComplete`; otherwise the committed Stage 5 export, or its absence, is kept verbatim.
   */
  publishStage5?: boolean
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

const exportedGradientBorder = (text: string | undefined): string | undefined => {
  const m = text === undefined ? null : /export const VERIFIED_STAGE3_BAR_GRADIENT_BORDER: [^\n]* = \{\n[\s\S]*?\n?\}\n/.exec(text)
  return m ? m[0] : undefined
}

const exportedAutoBounds = (text: string | undefined): string | undefined => {
  const m = text === undefined ? null : /export const VERIFIED_STAGE3_BAR_AUTO_BOUNDS: [^\n]* = \{\n[\s\S]*?\n?\}\n/.exec(text)
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
  const borderBlock = stage3Report ? renderBarGradientBorderBlock(stage3Report.barGradientBorder) : exportedGradientBorder(options.committedSource)
  const autoBoundsBlock = stage3Report ? renderBarAutoBoundsBlock(stage3Report.barAutoBounds) : exportedAutoBounds(options.committedSource)
  // Task 30: only the gated set publishes (`stage4Publication`), with the selections production reads beside it.
  const stage4Gate = options.publishStage4 ? stage4Publication(auditStage4Manifest(manifest, { rootDir })) : undefined
  const stage4Verified = stage4Gate ? stage4Gate.published : exportedSet(options.committedSource, 'VERIFIED_STAGE4_CAPABILITIES')
  const selectionsBlock = stage4Gate ? renderIconSelectionsBlock(stage4Gate.selections) : exportedSelections(options.committedSource)
  // Task 31: Stage 5 publishes only a complete audit's gated set; an incomplete one leaves the committed export as is.
  const stage5Report = options.publishStage5 ? auditStage5Manifest(manifest, { rootDir }) : undefined
  const stage5Gate = stage5Report?.stageComplete ? stage5Publication(stage5Report) : undefined
  const stage5Verified = stage5Gate ? stage5Gate.published : exportedSet(options.committedSource, 'VERIFIED_STAGE5_CAPABILITIES')
  // Task 39: the matched reading of each published selection capability travels beside the set (evidence.ts reads it).
  const formulaSelectionsBlock = stage5Gate ? renderFormulaSelectionsBlock(stage5Gate.selections) : exportedFormulaSelections(options.committedSource)
  const list = (ids: readonly string[]): string => ids.map(id => `  '${id}',`).join('\n')
  const allVerified = [...new Set([...stage1Verified, ...stage2Verified, ...(stage3Verified ?? []), ...(stage4Verified ?? []), ...(stage5Verified ?? [])])].sort()
  const block = (name: string, ids: readonly string[]): string => `export const ${name}: ReadonlySet<string> = new Set([\n${list(ids)}\n])\n`
  return `/**
 * Verified conditional formatting capabilities, every stage from its own audit (Stage 3 only once published, Ruling S3-WRITE).
 * Generated via bun scripts/cf-refs/audit.ts --write (any --stage; --publish-stage3 publishes Stage 3).
 * Checked by tests/xlsx-cf-stage1.test.ts, tests/xlsx-cf-stage2.test.ts and tests/xlsx-cf-oracle.test.ts.
 */
${block('VERIFIED_STAGE1_CAPABILITIES', stage1Verified)}
${block('VERIFIED_STAGE2_CAPABILITIES', stage2Verified)}
${stage3Verified !== undefined ? `${block('VERIFIED_STAGE3_CAPABILITIES', stage3Verified)}\n` : ''}${stage3Verified !== undefined && autoColoursBlock !== undefined ? `${autoColoursBlock}\n` : ''}${stage3Verified !== undefined && layoutBlock !== undefined ? `${layoutBlock}\n` : ''}${stage3Verified !== undefined && gradientBlock !== undefined ? `${gradientBlock}\n` : ''}${stage3Verified !== undefined && lengthsBlock !== undefined ? `${lengthsBlock}\n` : ''}${stage3Verified !== undefined && borderBlock !== undefined ? `${borderBlock}\n` : ''}${stage3Verified !== undefined && autoBoundsBlock !== undefined ? `${autoBoundsBlock}\n` : ''}${stage4Verified !== undefined ? `${block('VERIFIED_STAGE4_CAPABILITIES', stage4Verified)}\n` : ''}${stage4Verified !== undefined && selectionsBlock !== undefined ? `${selectionsBlock}\n` : ''}${stage5Verified !== undefined ? `${block('VERIFIED_STAGE5_CAPABILITIES', stage5Verified)}\n` : ''}${stage5Verified !== undefined && formulaSelectionsBlock !== undefined ? `${formulaSelectionsBlock}\n` : ''}${block('VERIFIED_CAPABILITIES', allVerified)}`
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
  const work = createCfWorkBudget({ scan: true, maxUnits: 10_000_000 })
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

/* ---------------- Stage 3 bar pressure (Task 23) ---------------- */

/** A paired (x14) solid min/max data bar at 0/100 lengths, left to right, no border or axis colour: the shape the Stage 3
 * evidence verifies for all-positive data, so the scenarios below exercise painted bars, not withheld ones. */
export const barRule = (priority: number, range: { firstRow: number; lastRow: number; firstCol: number; lastCol: number }, bounds?: unknown[]) => ({
  id: priority,
  priority,
  documentOrder: priority,
  stopIfTrue: false,
  ranges: [range],
  writes: new Set(['dataBar']),
  definition: {
    kind: 'dataBar',
    bounds: bounds ?? [{ type: 'min' }, { type: 'max' }],
    fillArgb: 0xff638ec6,
    minLength: 0,
    maxLength: 100,
    gradient: false,
    border: false,
    negativeBarColorSameAsPositive: false,
    negativeBarBorderColorSameAsPositive: true,
    axisPosition: 'automatic',
    direction: 'leftToRight',
    showValue: true,
    sheetRightToLeft: false,
    pairing: 'x14',
    dialect: 'x14',
    formulas: [],
  },
})

/** Whether a resident scope announces a painted data bar at the cell (Ruling S3-P label). */
const announcesBar = (scope: { describe(row: number, col: number): { text: string } | undefined }, row: number, col: number): boolean =>
  scope.describe(row, col)?.text.startsWith('Data bar') === true

/** Wrap sheet cells to count EVERY read of a cell's value (not distinct cells): a scan repeated per rule multiplies it. */
export function countingValueReads(sheet: XlsxSheet): { sheet: XlsxSheet; reads: () => number } {
  let reads = 0
  const rows = sheet.rows.map(band => ({
    ...band,
    cells: band.cells.map(cell => new Proxy(cell, {
      get: (t, p, r) => {
        if (p === 'value') reads++
        return Reflect.get(t, p, r)
      },
    })),
  }))
  return { sheet: { ...sheet, rows } as unknown as XlsxSheet, reads: () => reads }
}

/**
 * A 1M-row DECLARED `A:A` scope (100 stored cells) under 100 data bars and 100 colour scales. Two proofs:
 * - statistics: the 200 rule scopes are one scope, so one sparse scan serves them all (`sourceReads` = 100 stored cells);
 * - production: the resident text scope paints the first bar on every stored cell (S3-I-OVERLAP; whole-column and
 *   duplicate-block bars are verified), and every value read during preparation is counted. A per-rule scan would read
 *   each value 200 times; the shared scan keeps the total within a small multiple of the 100 cells.
 * It is NOT a populated 1M-row scope (as Stage 2's M5 note).
 */
export async function pressureBar1MRowSharedScan(): Promise<PressureSample & { statisticsReads: number; valueReads: number; estimatorRefusedAll: boolean }> {
  const started = Date.now()
  const wholeCol = { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 }
  const rules = [
    ...Array.from({ length: 100 }, (_, i) => barRule(i, wholeCol)),
    ...Array.from({ length: 100 }, (_, i) => scaleRule(100 + i, wholeCol)),
  ]
  // Values 5 and 15 alternate: every bar is either the zero-length minimum or the full span, so none rounds away (the
  // unmeasured minimum-visible-width branch would withhold it).
  const cells: XlsxCell[] = Array.from({ length: 100 }, (_, row) => ({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: (row % 2) * 10 + 5, hasCachedValue: true }))
  const inner = { name: 'P', sheetId: 'p', rows: [{ index: 0, cells }], cols: [], mergeRefs: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules } } as unknown as XlsxSheet
  // Statistics: one scan for 200 identical scopes.
  const memory = createCfMemory()
  const stamp = { owner: 'audit', sheetId: inner.sheetId ?? inner.name, workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
  const session = createCfStatistics(memory, { retainEntries: true })
  const outs = await session.prepareStatistics(rules.map(r => r.ranges), stamp, residentSource(inner, stamp, memory, 'engine'),
    createCfWorkBudget({ scan: true, maxUnits: 10_000_000 }), { signal: new AbortController().signal })
  const ready = outs.find(o => o.kind === 'ready')
  if (!ready || ready.kind !== 'ready') throw new Error('bar statistics scan must admit')
  const statisticsReads = ready.value.value.storedCells
  const statsPeak = memory.peakBytes
  for (const o of outs) if (o.kind === 'ready') o.value.release()
  session.dispose()
  if (memory.currentBytes !== 0) throw new Error(`bar statistics leaked ${memory.currentBytes} B`)
  // Production: the resident text scope. Its byte estimate is O(rules x cells) (every rule's property bytes per cell),
  // so 200 rules over all 100 cells exceed the 8 MiB scratch bound and refuse before any read; 40 cells admit.
  const { sheet, reads } = countingValueReads(inner)
  const { budget, net, peak } = createTrackingBudget()
  let estimatorRefusedAll = false
  try {
    await prepareCfTextScope(sheet, { rects: [{ r0: 0, r1: 99, c0: 0, c1: 0 }], signal: new AbortController().signal, locale: 'en-US', budget })
  } catch (error) {
    if (!(error instanceof PackageError) || error.code !== 'resource-limit') throw error
    estimatorRefusedAll = true
  }
  if (reads() !== 0 || net() !== 0) throw new Error(`the refused 100-cell request read ${reads()} values and left ${net()} B`)
  const rows = 40
  const scope = await prepareCfTextScope(sheet, {
    rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }],
    signal: new AbortController().signal,
    locale: 'en-US',
    budget,
  })
  const valueReads = reads()
  let painted = 0
  // The bar paints (`barAt`); the cell announces unavailable formatting because its whole-column scales withhold
  // (scale:scope-whole-column is unverified).
  if (scope) for (let row = 0; row < rows; row++) if (scope.barAt?.(row, 0) !== undefined) painted++
  const peakBytes = Math.max(peak(), statsPeak)
  scope?.dispose()
  return {
    name: 'bar-sparse-1M-declared-shared-scan',
    outcome: scope === undefined ? 'refused' : 'admitted',
    peakBytes,
    netBytes: net(),
    sourceReads: statisticsReads,
    cellsWithResults: painted,
    statisticsReads,
    valueReads,
    estimatorRefusedAll,
    wallMs: Date.now() - started,
  }
}

/** An invalid (out-of-range) percentile bar withholds as unavailable while a min/max bar over another column still
 * paints (it is not a sorted-copy refusal: that path is covered in tests/xlsx-cf-statistics.test.ts). */
export async function pressureRefusedPercentileBarWithMinMax(): Promise<PressureSample> {
  const started = Date.now()
  const cells: XlsxCell[] = []
  for (let col = 0; col < 2; col++) for (let row = 0; row < 10; row++) {
    cells.push({ ref: `${'AB'[col]}${row + 1}`, row, col, styleIndex: 0, value: (row + 1) * 10, hasCachedValue: true })
  }
  const minMax = barRule(1, { firstRow: 0, lastRow: 9, firstCol: 0, lastCol: 0 })
  const percentile = { ...barRule(2, { firstRow: 0, lastRow: 9, firstCol: 1, lastCol: 1 }, [{ type: 'percentile', value: -10 }, { type: 'max' }]), documentOrder: 1 }
  const inner = {
    name: 'P', sheetId: 'p', rows: [{ index: 0, cells }], cols: [], mergeRefs: [], mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules: [minMax, percentile] },
  } as unknown as XlsxSheet
  const { sheet, visits } = countingSheetCells(inner)
  const { budget, net, peak } = createTrackingBudget()
  const scope = await prepareCfTextScope(sheet, {
    rects: [{ r0: 0, r1: 9, c0: 0, c1: 0 }, { r0: 0, r1: 9, c0: 1, c1: 1 }],
    signal: new AbortController().signal,
    locale: 'en-US',
    budget,
  })
  if (scope === undefined) throw new Error('scope must admit')
  const visited = visits()
  let painted = 0
  for (let row = 0; row < 10; row++) {
    if (announcesBar(scope, row, 0)) painted++
    const b = scope.describe(row, 1)?.text
    if (b !== 'Conditional formatting unavailable') throw new Error(`percentile bar B${row + 1} must withhold as unavailable, got ${b}`)
    if (scope.barAt?.(row, 1) !== undefined) throw new Error(`percentile bar B${row + 1} must not report a bar`)
  }
  const peakBytes = peak()
  scope.dispose()
  return { name: 'invalid-percentile-bar-withheld-with-min-max-bar-painted', outcome: 'admitted', peakBytes, netBytes: net(), sourceReads: visited, cellsWithResults: painted, wallMs: Date.now() - started }
}

/** A whole-column bar does not extend the sheet's extents, and its scope paints the stored cells only. */
export async function pressureBarWholeColumnExtents(): Promise<PressureSample> {
  const started = Date.now()
  const dataRows = 10
  const inner = {
    name: 'WCBar', sheetId: 'wcbar',
    rows: Array.from({ length: dataRows }, (_, r) => ({ index: r, cells: [{ ref: `A${r + 1}`, row: r, col: 0, styleIndex: 0, value: (r + 1) * 10, hasCachedValue: true }] })),
    cols: [], mergeRefs: [], mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules: [barRule(1, { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 })] },
  } as unknown as XlsxSheet
  const m = computeMetrics(inner)
  if (m.rowHeightsPx.length !== dataRows || m.heightPx > 1000) {
    throw new Error(`CF bar extended metrics extent: ${m.rowHeightsPx.length} rows, height ${m.heightPx}px`)
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
  let painted = 0
  for (let r = 0; r < dataRows; r++) if (announcesBar(scope, r, 0)) painted++
  const peakBytes = peak()
  scope.dispose()
  return { name: 'bar-whole-column-extents', outcome: 'admitted', peakBytes, netBytes: net(), sourceReads: visited, cellsWithResults: painted, wallMs: Date.now() - started }
}

/** Distinct-range bar pressure: 150 bars over distinct two-cell ranges; a 10-rect batch admits, paints and balances;
 * 300 distinct rects exceed the per-call rect cap and refuse before enumeration with net zero. */
export async function pressureBarDistinctRanges(): Promise<PressureSample> {
  const started = Date.now()
  const N = 300
  // Two cells per rule: a single-cell scope is degenerate (S3-E keeps it unresolved), a pair is an ordinary min/max bar.
  const rules = Array.from({ length: N / 2 }, (_, i) => barRule(i, { firstRow: 2 * i, lastRow: 2 * i + 1, firstCol: 0, lastCol: 0 }))
  const cells: XlsxCell[] = []
  for (let row = 0; row < N; row++) cells.push({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: (row % 2) * 10 + 5, hasCachedValue: true })
  const inner = { name: 'DB', sheetId: 'db', rows: [{ index: 0, cells }], cols: [], mergeRefs: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules } } as unknown as XlsxSheet
  const { sheet, visits } = countingSheetCells(inner)
  {
    const { budget, net, peak } = createTrackingBudget()
    const scope = await prepareCfTextScope(sheet, { rects: Array.from({ length: 10 }, (_, i) => ({ r0: i, r1: i, c0: 0, c1: 0 })), signal: new AbortController().signal, locale: 'en-US', budget })
    if (scope === undefined) throw new Error('small distinct bar batch must admit')
    let painted = 0
    for (let row = 0; row < 10; row++) if (announcesBar(scope, row, 0)) painted++
    if (painted !== 10) throw new Error(`small distinct bar batch painted ${painted}/10`)
    const peakBytes = peak()
    scope.dispose()
    checkPressureSample({ name: 'bar-distinct-range-10', outcome: 'admitted', peakBytes, netBytes: net(), sourceReads: visits(), cellsWithResults: painted, wallMs: 0 }, { expectOutcome: 'admitted' })
  }
  const visitedBefore = visits()
  const { budget, net, peak } = createTrackingBudget()
  let scope: Awaited<ReturnType<typeof prepareCfTextScope>> | undefined
  try {
    scope = await prepareCfTextScope(sheet, { rects: Array.from({ length: N }, (_, i) => ({ r0: i, r1: i, c0: 0, c1: 0 })), signal: new AbortController().signal, locale: 'en-US', budget })
  } catch (error) {
    if (!(error instanceof PackageError) || error.code !== 'resource-limit') throw error
  }
  scope?.dispose()
  return { name: 'bar-distinct-range-300', outcome: scope === undefined ? 'refused' : 'admitted', peakBytes: peak(), netBytes: net(), sourceReads: visits() - visitedBefore, cellsWithResults: 0, wallMs: Date.now() - started }
}

/** A synthetic column of `cells` numbers that charges one work unit per examined cell, as `residentSource` does. */
function syntheticColumnSource(cells: number, stamp: CfStamp): CfValueSource & { reads: () => number } {
  let reads = 0
  return {
    numberFormats: 'resolved',
    stamp: () => stamp,
    async bounds() { return { firstRow: 0, lastRow: Math.max(0, cells - 1), firstCol: 0, lastCol: 0 } },
    async *pages(_id, _rects, budget, signal) {
      const rows = new Uint32Array(1), cols = new Uint16Array(1)
      for (let r = 0; r < cells; r++) {
        if (signal.aborted) throw new Error('CF cancelled')
        const work = budget.spend(1)
        if (work === 'refused') throw new Error('CF work limit')
        if (work === 'yield') budget.resume()
        reads++
        rows[0] = r; cols[0] = 0
        yield { value: { rows, cols, values: [{ kind: 'number', value: (r % 100) + 1 }] }, bytes: 256, release() {}, transfer() { return { kind: 'unresolved', reason: 'borrowed-page' } } } as CfLease<CfValuePage>
      }
    },
    reads: () => reads,
  }
}

/**
 * The 1,000,000-unit work cap (CF_LIMITS.workUnits) on a bar's statistics scan, under the default statistics budget a
 * service builds (no time limit, default units). A scanned cell costs far more than the one unit the page charges (the
 * statistics accounting, its checkpoints and its buffer growth: about 66 units per cell at 10,000 cells, measured), so
 * 10,000 cells admit and 100,000 refuse `cf-resource-limit` part-way, with nothing left charged. Synthetic one-cell pages
 * keep the column cheap to build.
 */
export async function pressureBarUnitCap(): Promise<{ name: string; admittedCells: number; admittedReads: number; admittedUnits: number; refusedCells: number; refusedReads: number; refusedUnits: number; refusedReason: string; netBytes: number; wallMs: number }> {
  const started = Date.now()
  const stamp = { owner: 'audit', sheetId: 'cap', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
  const run = async (cells: number) => {
    const memory = createCfMemory()
    const session = createCfStatistics(memory)
    const source = syntheticColumnSource(cells, stamp)
    const budget = createCfWorkBudget({ scan: true })
    const [out] = await session.prepareStatistics([[{ firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 }]], stamp, source,
      budget, { signal: new AbortController().signal })
    const result = { kind: out!.kind, reason: out!.kind === 'unresolved' ? out!.reason : '', reads: source.reads(), units: budget.units }
    if (out!.kind === 'ready') out!.value.release()
    session.dispose()
    return { ...result, net: memory.currentBytes }
  }
  const admittedCells = 10_000, refusedCells = 100_000
  const admitted = await run(admittedCells)
  if (admitted.kind !== 'ready') throw new Error(`bar scan of ${admittedCells} cells must admit, got ${admitted.kind} ${admitted.reason}`)
  const refused = await run(refusedCells)
  if (refused.kind !== 'unresolved') throw new Error(`bar scan of ${refusedCells} cells must refuse at the unit cap`)
  return { name: 'bar-statistics-unit-cap', admittedCells, admittedReads: admitted.reads, admittedUnits: admitted.units, refusedCells, refusedReads: refused.reads,
    refusedUnits: refused.units, refusedReason: refused.reason, netBytes: admitted.net + refused.net, wallMs: Date.now() - started }
}

/* ---------------- Stage 4 (icon sets) pressure scenarios (Task 30) ---------------- */

/** An icon-set rule (3Arrows, base dialect) over `range`, with `thresholds` (default percent 0/33/67). */
export const iconRuleOf = (priority: number, range: { firstRow: number; lastRow: number; firstCol: number; lastCol: number }, thresholds?: unknown[]) => ({
  id: priority,
  priority,
  documentOrder: priority,
  stopIfTrue: false,
  ranges: [range],
  writes: new Set(['iconSet']),
  definition: {
    kind: 'iconSet', dialect: 'base', formulas: [], set: '3Arrows',
    thresholds: thresholds ?? [{ bound: { type: 'percent', value: 0 }, gte: true }, { bound: { type: 'percent', value: 33 }, gte: true }, { bound: { type: 'percent', value: 67 }, gte: true }],
    reverse: false, showValue: true, percentAttr: true, sheetRightToLeft: false,
  },
})

/** Counts the icon descriptors production's resident service evaluates over `rects` (one lease per rect, under the
 * production evidence), with the bytes it held. */
async function residentIcons(sheet: XlsxSheet, rects: ReadonlyArray<{ firstRow: number; lastRow: number; firstCol: number; lastCol: number }>, memory = createCfMemory()): Promise<{ icons: number; unresolved: number; byCol: Record<number, string[]>; peak: number; net: number }> {
  const { createCfService } = await import('../../src/xlsx/cf/service')
  const { CF_LIMITS } = await import('../../src/xlsx/cf/limits')
  const service = createCfService(sheet.conditionalFormatting!, memory)
  const stamp = { owner: 'audit', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
  const source = residentSource(sheet, stamp, memory, 'engine')
  let icons = 0, unresolved = 0
  const byCol: Record<number, string[]> = {}
  for (const rect of rects) {
    const lease = await service.prepareCf({ stamp, sourceRects: [rect], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
    for (let row = rect.firstRow; row <= rect.lastRow; row++) for (let col = rect.firstCol; col <= rect.lastCol; col++) {
      const o = lease.value.at(row, col)
      ;(byCol[col] ??= []).push(o.kind === 'ready' ? (o.value.visual?.iconSet ? 'icon' : `ready:${[...(o.value.unresolvedWrites ?? [])].join('+')}`) : o.kind === 'unresolved' ? o.reason : o.kind)
      if (o.kind === 'ready' && o.value.visual?.iconSet) icons++
      else if (o.kind === 'unresolved') unresolved++
    }
    lease.release()
  }
  service.dispose()
  return { icons, unresolved, byCol, peak: memory.peakBytes, net: memory.currentBytes }
}

/**
 * A 1M-row DECLARED `A:A` scope (100 stored cells) under 100 icon sets, 100 data bars and 100 colour scales. The 300
 * rule scopes are one scope, so one sparse scan serves them all (`statisticsReads` = 100 stored cells); production's
 * resident service then evaluates the first icon set on every requested cell (S4-I first-paints), with and without the
 * bars and scales beside it. It is NOT a populated 1M-row scope.
 */
export async function pressureIcon1MRowSharedScan(): Promise<PressureSample & { statisticsReads: number; iconOnlyCells: number }> {
  const started = Date.now()
  const wholeCol = { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 }
  const rules = [
    ...Array.from({ length: 100 }, (_, i) => iconRuleOf(i, wholeCol)),
    ...Array.from({ length: 100 }, (_, i) => barRule(100 + i, wholeCol)),
    ...Array.from({ length: 100 }, (_, i) => scaleRule(200 + i, wholeCol)),
  ]
  const cells: XlsxCell[] = Array.from({ length: 100 }, (_, row) => ({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: (row % 2) * 10 + 5, hasCachedValue: true }))
  const sheetOf = (r: unknown[]) => ({ name: 'P', sheetId: 'p', rows: [{ index: 0, cells }], cols: [], mergeRefs: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules: r } } as unknown as XlsxSheet)
  const inner = sheetOf(rules)
  const memory = createCfMemory()
  const stamp = { owner: 'audit', sheetId: 'p', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
  const session = createCfStatistics(memory, { retainEntries: true })
  const outs = await session.prepareStatistics(rules.map(r => r.ranges), stamp, residentSource(inner, stamp, memory, 'engine'),
    createCfWorkBudget({ scan: true, maxUnits: 10_000_000 }), { signal: new AbortController().signal })
  const ready = outs.find(o => o.kind === 'ready')
  if (!ready || ready.kind !== 'ready') throw new Error('icon statistics scan must admit')
  const statisticsReads = ready.value.value.storedCells
  for (const o of outs) if (o.kind === 'ready') o.value.release()
  session.dispose()
  if (memory.currentBytes !== 0) throw new Error(`icon statistics leaked ${memory.currentBytes} B`)
  const rects = Array.from({ length: 5 }, (_, i) => ({ firstRow: i * 8, lastRow: i * 8 + 7, firstCol: 0, lastCol: 0 }))
  const all = await residentIcons(inner, rects)
  // The same 100 icon sets without the scales and bars: every requested cell carries the first icon.
  const iconOnly = await residentIcons(sheetOf(rules.slice(0, 100)), rects)
  return {
    name: 'icon-sparse-1M-declared-shared-scan',
    outcome: 'admitted',
    peakBytes: Math.max(memory.peakBytes, all.peak, iconOnly.peak),
    netBytes: memory.currentBytes + all.net + iconOnly.net,
    sourceReads: statisticsReads,
    cellsWithResults: all.icons,
    statisticsReads,
    iconOnlyCells: iconOnly.icons,
    wallMs: Date.now() - started,
  }
}

/** A refused percentile icon set (its sorted copy exceeds the admitted memory) withholds while a `num` icon set over
 * another column still evaluates: the refusal is per rule, never the sheet's. Control: with room for the copy, the
 * percentile set evaluates too. */
export async function pressureRefusedPercentileIconWithNum(): Promise<PressureSample & { percentileWithheld: number; controlPercentileIcons: number }> {
  const started = Date.now()
  const rows = 5_000
  const cells: XlsxCell[] = []
  for (let row = 0; row < rows; row++) {
    cells.push({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: row % 10, hasCachedValue: true })
    cells.push({ ref: `B${row + 1}`, row, col: 1, styleIndex: 0, value: row, hasCachedValue: true })
  }
  const num = iconRuleOf(1, { firstRow: 0, lastRow: rows - 1, firstCol: 0, lastCol: 0 }, [{ bound: { type: 'num', value: 0 }, gte: true }, { bound: { type: 'num', value: 3 }, gte: true }, { bound: { type: 'num', value: 6 }, gte: true }])
  const percentile = iconRuleOf(2, { firstRow: 0, lastRow: rows - 1, firstCol: 1, lastCol: 1 }, [{ bound: { type: 'percentile', value: 0 }, gte: true }, { bound: { type: 'percentile', value: 33 }, gte: true }, { bound: { type: 'percentile', value: 67 }, gte: true }])
  const inner = { name: 'P', sheetId: 'p', rows: [{ index: 0, cells }], cols: [], mergeRefs: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules: [num, percentile] } } as unknown as XlsxSheet
  const rects = [{ firstRow: 0, lastRow: 9, firstCol: 0, lastCol: 1 }]
  // 400,000 B admits the streaming num scope and the batch, not the 5,000-value sorted copy; 4,000,000 B admits both.
  const got = await residentIcons(inner, rects, createCfMemory({ limitBytes: 400_000 }))
  const control = await residentIcons(inner, rects, createCfMemory({ limitBytes: 4_000_000 }))
  const a = got.byCol[0] ?? [], b = got.byCol[1] ?? []
  if (a.filter(x => x === 'icon').length !== 10) throw new Error(`the num icon set must evaluate its 10 cells, got ${a.join(',')}`)
  const percentileWithheld = b.filter(x => x.startsWith('ready:') && x.includes('iconSet')).length
  if (percentileWithheld !== 10) throw new Error(`the refused percentile set must withhold its 10 cells, got ${b.join(',')}`)
  const controlPercentileIcons = (control.byCol[1] ?? []).filter(x => x === 'icon').length
  return { name: 'refused-percentile-icon-with-num-icon-rendering', outcome: 'admitted', peakBytes: Math.max(got.peak, control.peak), netBytes: got.net + control.net, sourceReads: rows, cellsWithResults: 10, percentileWithheld, controlPercentileIcons, wallMs: Date.now() - started }
}

/** A whole-column icon set does not extend the sheet's extents, and evaluates its stored cells only. */
export async function pressureIconWholeColumnExtents(): Promise<PressureSample> {
  const started = Date.now()
  const dataRows = 10
  const inner = {
    name: 'WCIcon', sheetId: 'wcicon',
    rows: Array.from({ length: dataRows }, (_, r) => ({ index: r, cells: [{ ref: `A${r + 1}`, row: r, col: 0, styleIndex: 0, value: (r + 1) * 10, hasCachedValue: true }] })),
    cols: [], mergeRefs: [], mergeRanges: [],
    conditionalFormatting: { diagnostics: [], rules: [iconRuleOf(1, { firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 })] },
  } as unknown as XlsxSheet
  const m = computeMetrics(inner)
  if (m.rowHeightsPx.length !== dataRows || m.heightPx > 1000) throw new Error(`CF icon set extended metrics extent: ${m.rowHeightsPx.length} rows, height ${m.heightPx}px`)
  const { sheet, visits } = countingSheetCells(inner)
  const got = await residentIcons(sheet, [{ firstRow: 0, lastRow: dataRows + 9, firstCol: 0, lastCol: 0 }])
  return { name: 'icon-whole-column-extents', outcome: 'admitted', peakBytes: got.peak, netBytes: got.net, sourceReads: visits(), cellsWithResults: got.icons, wallMs: Date.now() - started }
}

/**
 * Glyph-admission pressure (Ruling S4-R): one batch whose icons use every class the glyph-key sidecar records. Each
 * admitted class is decoded once and charged; withheld classes are never decoded; a memory too small for any class
 * refuses every admission and decodes nothing; release returns every byte.
 */
export async function pressureIconGlyphAdmission(): Promise<{ name: string; classes: number; admitted: number; decodes: number; withheldDecoded: number; refusedDecodes: number; peakBytes: number; netBytes: number; wallMs: number }> {
  const started = Date.now()
  const { admitBatchGlyphs } = await import('../../src/xlsx/cf/iconPaint')
  const { CF_ICON_GLYPH_CLASSES: classes, CF_ICON_GLYPH_WITHHELD: withheld, cfDecodeCalls, cfDecodedGlyphCount } = await import('../../src/xlsx/cf/iconGlyphs')
  const pairs = Object.keys(classes).map(k => { const [set, id] = k.split(':'); return { ruleId: 1, set: set as never, iconId: Number(id), bucket: 0, buckets: Number(set![0]), showValue: true } })
  const admittedClasses = new Set(Object.values(classes).filter(c => !withheld.has(c)))
  const memory = createCfMemory()
  const before = cfDecodeCalls()
  const batch = admitBatchGlyphs(pairs, memory, () => createCfWorkBudget({}))
  const decodes = cfDecodeCalls() - before
  const decoded = cfDecodedGlyphCount()
  const peak = memory.peakBytes
  batch.release()
  const tiny = createCfMemory({ limitBytes: 64 })
  const beforeRefused = cfDecodeCalls()
  const refused = admitBatchGlyphs(pairs, tiny, () => createCfWorkBudget({}))
  const refusedDecodes = cfDecodeCalls() - beforeRefused
  refused.release()
  return { name: 'icon-glyph-admission-every-class', classes: new Set(Object.values(classes)).size, admitted: decoded, decodes, withheldDecoded: decoded - admittedClasses.size, refusedDecodes, peakBytes: peak, netBytes: memory.currentBytes + tiny.currentBytes, wallMs: Date.now() - started }
}

/** Distinct-range icon pressure: 150 icon sets over distinct two-cell ranges; a 10-rect batch admits and evaluates; 300
 * distinct rects exceed the per-call rect cap and refuse before enumeration with net zero. */
export async function pressureIconDistinctRanges(): Promise<PressureSample> {
  const started = Date.now()
  const N = 300
  const rules = Array.from({ length: N / 2 }, (_, i) => iconRuleOf(i, { firstRow: 2 * i, lastRow: 2 * i + 1, firstCol: 0, lastCol: 0 }))
  const cells: XlsxCell[] = []
  for (let row = 0; row < N; row++) cells.push({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: (row % 2) * 10 + 5, hasCachedValue: true })
  const inner = { name: 'DI', sheetId: 'di', rows: [{ index: 0, cells }], cols: [], mergeRefs: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules } } as unknown as XlsxSheet
  const { sheet, visits } = countingSheetCells(inner)
  const small = await residentIcons(sheet, Array.from({ length: 10 }, (_, i) => ({ firstRow: i, lastRow: i, firstCol: 0, lastCol: 0 })))
  if (small.icons !== 10 || small.net !== 0) throw new Error(`small distinct icon batch evaluated ${small.icons}/10 with ${small.net} B left`)
  const visitedBefore = visits()
  const { budget, net, peak } = createTrackingBudget()
  let scope: Awaited<ReturnType<typeof prepareCfTextScope>> | undefined
  try {
    scope = await prepareCfTextScope(sheet, { rects: Array.from({ length: N }, (_, i) => ({ r0: i, r1: i, c0: 0, c1: 0 })), signal: new AbortController().signal, locale: 'en-US', budget })
  } catch (error) {
    if (!(error instanceof PackageError) || error.code !== 'resource-limit') throw error
  }
  scope?.dispose()
  return { name: 'icon-distinct-range-300', outcome: scope === undefined ? 'refused' : 'admitted', peakBytes: peak(), netBytes: net(), sourceReads: visits() - visitedBefore, cellsWithResults: 0, wallMs: Date.now() - started }
}

/* ---------------- Stage 4 per-case validation records (Task 30 fix round 2, codex P2) ---------------- */

/** One validation row per measured Stage 4 case and variant (a designated cell, its control included). */
export interface Stage4CaseRecord {
  caseId: string
  /** The variant: the designated cell. */
  ref: string
  status: string
  /** The cell's evidence labels (`iconEvidence` bucket/geometry), or `control`. */
  evidenceKind: string
  /** What Excel drew: the sidecar's ink box in cell-relative px at 1 px/pt and its observed class (null: no icon). */
  measured: { box?: { l: number; t: number; r: number; b: number }; glyphClass: string | null | undefined }
  /** What production does: its evaluated class (`withheld` when the slot is withheld), whether it paints, and the
   * painted frame's ink box (frame plus the class's glyph-key ink offset), cell-relative px at 1 px/pt. */
  produced: { glyphClass: string | null | 'withheld'; paints: boolean; box?: { l: number; t: number; r: number; b: number } }
  /** The mode the row was produced in. */
  mode: 'resident'
  stamp: { cfGen: number; workbookValueGen: number }
  /** The case's CF memory peak (bytes) and the work units its preparations spent. */
  bytePeak: number
  workUnits: number
}

/** The per-case/variant records of every measured Stage 4 case, under production evidence (nothing injected). */
export async function stage4CaseRecords(manifest: CfManifest, rootDir: string = REPO_ROOT): Promise<Stage4CaseRecord[]> {
  const { OfficePackage } = await import('../../src/core/zip')
  const { parseXlsx } = await import('../../src/xlsx/parse')
  const { createCfService } = await import('../../src/xlsx/cf/service')
  const { CF_LIMITS } = await import('../../src/xlsx/cf/limits')
  const { cfRuntimeEvidence } = await import('../../src/xlsx/cf/predicates')
  const { prepareCfPaint, cfRowPtOfSheet, cfRowUniformOfSheet } = await import('../../src/xlsx/cf/render')
  const { iconTextOf, layoutIcon } = await import('../../src/xlsx/cf/iconPaint')
  const { cfRulesVersion } = await import('../../src/xlsx/view/pagedSource')
  const { createCellLayoutHost } = await import('../../src/xlsx/view/renderViewport')
  const { observedIconClass, iconTwinSetMembers } = await import('./verify')
  const { createCanvas } = await import('canvas')
  const ctx = createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
  // Each class's ink offset inside the frame, from the glyph key (20 px rows: frame 2 px in, 1 px above the bottom).
  const keyCase = manifest.cases.find(c => c.id === 'stage4-icon-glyph-key')!
  const key = JSON.parse(readFileSync(resolve(rootDir, keyCase.captures[0]!.sidecarPath), 'utf8')) as { observedIcons: Array<{ cellLeftPx: number; cellTopPx: number; cellHeightPx: number; icon: null | { leftPx: number; topPx: number; rightPx: number; bottomPx: number; glyphClass: string } }> }
  const inkOffset = new Map<string, { l: number; t: number; r: number; b: number }>()
  for (const o of key.observedIcons) {
    if (!o.icon || iconTwinSetMembers(o.icon.glyphClass)) continue
    const fx = o.cellLeftPx + 2, fy = o.cellTopPx + o.cellHeightPx - 1 - 12
    if (!inkOffset.has(o.icon.glyphClass)) inkOffset.set(o.icon.glyphClass, { l: o.icon.leftPx - fx, t: o.icon.topPx - fy, r: o.icon.rightPx - fx, b: o.icon.bottomPx - fy })
  }
  const refRC = (ref: string) => { const m = /^([A-Z]+)(\d+)$/.exec(ref)!; return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 } }
  const out: Stage4CaseRecord[] = []
  for (const c of manifest.cases.filter(x => x.stage === 4 && x.status === 'measured')) {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(rootDir, c.saved)))))
    const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
    const sidecar = c.captures[0] && existsSync(resolve(rootDir, c.captures[0].sidecarPath)) ? JSON.parse(readFileSync(resolve(rootDir, c.captures[0].sidecarPath), 'utf8')) as { observedIcons?: CfObservedIconRow[] } : undefined
    const observed = new Map((sidecar?.observedIcons ?? []).map(o => [o.ref.toUpperCase(), o] as const))
    const memory = createCfMemory()
    let units = 0
    const budgets: CfWorkBudgetLike[] = []
    const budget = (scan: boolean) => () => { const b = createCfWorkBudget(scan ? { scan: true } : {}); budgets.push(b); return b }
    const cf = sheet.conditionalFormatting
    const service = cf ? createCfService(cf, memory, { evidence: cfRuntimeEvidence, budget: budget(false), statisticsBudget: budget(true) }) : undefined
    const stamp = { owner: 'audit-records', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cf ? cfRulesVersion(cf) : 0, viewGen: 0, structureGen: 0, clockGen: 0 }
    const source = residentSource(sheet, stamp, memory, 'engine')
    const batch = cf ? await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, layoutContext: () => ctx }) : undefined
    const m = computeMetrics(sheet)
    const rowPtOf = cfRowPtOfSheet(sheet), rowUniformOf = cfRowUniformOfSheet(sheet)
    try {
      for (const exp of c.expected.cells) {
        const { row, col } = refRC(exp.ref)
        let produced: Stage4CaseRecord['produced'] = { glyphClass: 'withheld', paints: false }
        if (service) {
          const lease = await service.prepareCf({ stamp, sourceRects: [{ firstRow: row, lastRow: row, firstCol: col, lastCol: col }], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
          const o = lease.value.at(row, col)
          lease.release()
          const tainted = o.kind === 'ready' && (o.value.unresolvedWrites?.size ?? 0) > 0
          const icon = o.kind === 'ready' ? o.value.visual?.iconSet : undefined
          produced.glyphClass = icon ? (icon.set === 'NoIcons' ? null : CF_ICON_GLYPH_CLASSES[`${icon.set}:${icon.iconId}`] ?? `unknown:${icon.set}:${icon.iconId}`)
            : o.kind === 'no-match' || (o.kind === 'ready' && !tainted) ? null : 'withheld'
          const painted = batch?.iconAt(row, col)
          if (painted) {
            const cell = sheet.rows.find(r => r.index === row)?.cells.find(x => x.col === col)
            const width = m.colWidthsPx[col] ?? 64, height = m.rowHeightsPx[row] ?? 20
            const body = cell ? createCellLayoutHost(ctx).layoutBody(cell, width, height) : undefined
            const textWidth = body?.text?.kind === 'layout' ? body.text.layout.lines[0]!.segments.reduce((n, sg) => n + sg.width, 0) : 0
            const rowPt = rowPtOf(row), rowUniform = rowUniformOf(row)
            const host = { x: m.colWidthsPx.slice(0, col).reduce((a, b) => a + b, 0), y: m.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0), width, height, textWidth, ...(cell ? { cell } : {}), ...(rowPt !== undefined ? { rowPt } : {}), ...(rowUniform !== undefined ? { rowUniform } : {}) }
            const laid = layoutIcon(painted, host, iconTextOf(cell)!, { zoom: 3 / 4, dpr: 1, output: 'device' })
            const off = inkOffset.get(CF_ICON_GLYPH_CLASSES[`${painted.set}:${painted.iconId}`] ?? '')
            const X0 = Math.floor(host.x * 3 / 4 + 0.5), Y0 = Math.floor(host.y * 3 / 4 + 0.5)
            produced = { ...produced, paints: true, ...(laid && off ? { box: { l: laid.box.x - X0 + off.l, t: laid.box.y - Y0 + off.t, r: laid.box.x - X0 + off.r, b: laid.box.y - Y0 + off.b } } : {}) }
          }
        }
        const obs = observed.get(exp.ref.toUpperCase())
        out.push({
          caseId: c.id, ref: exp.ref, status: c.status,
          evidenceKind: exp.iconEvidence ? `${exp.iconEvidence.bucket}/${exp.iconEvidence.geometry}` : 'control',
          measured: { glyphClass: obs ? observedIconClass(obs as never) : undefined, ...(obs?.icon ? { box: { l: obs.icon.leftPx - obs.cellLeftPx, t: obs.icon.topPx - obs.cellTopPx, r: obs.icon.rightPx - obs.cellLeftPx, b: obs.icon.bottomPx - obs.cellTopPx } } : {}) },
          produced, mode: 'resident', stamp: { cfGen: stamp.cfGen, workbookValueGen: stamp.workbookValueGen },
          bytePeak: 0, workUnits: 0,
        })
      }
    } finally {
      batch?.dispose()
      service?.dispose()
    }
    for (const b of budgets) units += b.units
    for (const r of out) if (r.caseId === c.id) { r.bytePeak = memory.peakBytes; r.workUnits = units }
  }
  return out
}
// Fix round 2 (R2): the agreement and the slot-specific justifications live in ./stage5-withhold (shared with
// cf-compare and the conformance test).
export { stage5CellAgreement, type Stage5Excel, type Stage5Production } from './stage5-withhold'
import { stage5CellAgreement, stage5Justifications, type Stage5Excel, type Stage5Justification, type Stage5Production } from './stage5-withhold'

export interface Stage5CaseRecord {
  caseId: string
  /** The variant: the designated cell (or its control). */
  ref: string
  status: string
  /** Whether the row was compared with Excel; when not, `notCompared` says why (an unmeasured case). */
  compared: boolean
  notCompared?: string
  /** `display-format` under a formula rule, `literal` under literal twins/controls only, `control` with no rule. */
  evidenceKind: 'display-format' | 'literal' | 'control'
  /** The formula capabilities of the rules covering the cell. */
  capabilities: string[]
  excel: Stage5Excel | undefined
  production: Stage5Production
  /** The rule- and slot-specific justifications for the cell's unresolved slots (stage5Justifications). */
  justification: Stage5Justification[]
  /** Every resolved slot is Excel's, every unresolved slot is justified for that slot, a measured visual agrees. */
  agrees: boolean
  disagreement?: string
  mode: 'resident'
  stamp: { cfGen: number; workbookValueGen: number }
  /** The case's CF memory peak (bytes) and the work units its preparations spent. */
  bytePeak: number
  workUnits: number
}

/** The per-case/variant records of every Stage 5 case under production evidence (nothing injected). */
export async function stage5CaseRecords(manifest: CfManifest, rootDir: string = REPO_ROOT): Promise<Stage5CaseRecord[]> {
  const { OfficePackage } = await import('../../src/core/zip')
  const { parseXlsx } = await import('../../src/xlsx/parse')
  const { createCfService } = await import('../../src/xlsx/cf/service')
  const { CF_LIMITS } = await import('../../src/xlsx/cf/limits')
  const { cfRuntimeEvidence } = await import('../../src/xlsx/cf/predicates')
  const { cfWorkbookOfDocument, cfDocumentEvaluationOf } = await import('../../src/xlsx/cf/render')
  const { cfEvaluationContextFor } = await import('../../src/xlsx/cf/formula')
  const { resolveCfCellStyle } = await import('../../src/xlsx/cf/dxf')
  const { cfRulesVersion } = await import('../../src/xlsx/view/pagedSource')
  const { formulaObservationOf, decodePngSync, computeWorksheetRangeGeometry, readZipPartsSync, stage5IconKey, describeS5Visual } = await import('./verify')
  const refRC = (ref: string) => { const m = /^([A-Z]+)(\d+)$/.exec(ref)!; return { row: Number(m[2]) - 1, col: [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 } }
  const iconKey = stage5IconKey(rootDir)
  const out: Stage5CaseRecord[] = []
  for (const c of manifest.cases.filter(x => x.stage === 5)) {
    const kindOf = (ref: string): { covering: typeof c.markers; formulaCaps: string[]; kind: Stage5CaseRecord['evidenceKind'] } => {
      const covering = c.markers.filter(m => isCellInSqref(ref, m.sqref))
      const formulaCaps = [...new Set(covering.filter(m => m.formula !== undefined).map(m => m.capabilityId))].sort()
      return { covering, formulaCaps, kind: formulaCaps.length > 0 ? 'display-format' : covering.length > 0 ? 'literal' : 'control' }
    }
    const blank: Stage5Production = { fill: 'none', strike: false, unresolved: [], visual: 'none' }
    if (c.status !== 'measured') {
      // Not compared: the case holds no verified Excel evidence (a recorded outcome, a refused shape, or pending).
      const refs = c.expected.cells.length > 0 ? c.expected.cells.map(e => e.ref) : [...new Set(c.markers.map(m => m.sqref.split(/[: ]/)[0]!))]
      for (const ref of refs) {
        const k = kindOf(ref)
        out.push({ caseId: c.id, ref, status: c.status, compared: false, notCompared: `${c.status}: ${c.reason ?? 'no verified Excel capture'}`.slice(0, 400),
          evidenceKind: k.kind, capabilities: k.formulaCaps, excel: undefined, production: blank, justification: [], agrees: true, mode: 'resident', stamp: { cfGen: 0, workbookValueGen: 0 }, bytePeak: 0, workUnits: 0 })
      }
      continue
    }
    const savedBytes = new Uint8Array(readFileSync(resolve(rootDir, c.saved)))
    const doc = await parseXlsx(await OfficePackage.load(savedBytes))
    const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
    const cap = c.captures.find(x => x.appearance === 'screen') ?? c.captures[0]
    const sidecar = cap && existsSync(resolve(rootDir, cap.sidecarPath)) ? JSON.parse(readFileSync(resolve(rootDir, cap.sidecarPath), 'utf8')) as Record<string, unknown> : undefined
    const png = cap && cap.pngPath && existsSync(resolve(rootDir, cap.pngPath)) ? decodePngSync(new Uint8Array(readFileSync(resolve(rootDir, cap.pngPath)))) ?? undefined : undefined
    const parts = readZipPartsSync(savedBytes) ?? undefined
    const geom = parts && sidecar && Array.isArray(sidecar.colWidthsPt) ? computeWorksheetRangeGeometry(parts, c.range, { colWidthsPt: sidecar.colWidthsPt as number[], rowHeightsPt: sidecar.rowHeightsPt as number[] }) ?? undefined : undefined
    const observation = sidecar ? formulaObservationOf(c, sidecar, png, geom, iconKey) : undefined
    const memory = createCfMemory()
    const budgets: CfWorkBudgetLike[] = []
    const budget = (scan: boolean) => () => { const b = createCfWorkBudget(scan ? { scan: true } : {}); budgets.push(b); return b }
    const cf = sheet.conditionalFormatting
    // The clock cases were measured on their anchor day; production reads the same day through the evaluation clock.
    const now = c.clockRequirement ? () => Date.parse(`${c.clockRequirement!.anchorLocalDate}T04:00:00+08:00`) : Date.now
    const service = cf ? createCfService(cf, memory, { evidence: cfRuntimeEvidence, budget: budget(false), statisticsBudget: budget(true),
      evaluation: cfEvaluationContextFor(cfDocumentEvaluationOf(doc), now), sheetName: sheet.name, workbook: cfWorkbookOfDocument(doc.sheets, memory, 'engine') }) : undefined
    const stamp = { owner: 'audit-records', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cf ? cfRulesVersion(cf) : 0, viewGen: 0, structureGen: 0, clockGen: 0 }
    const source = residentSource(sheet, stamp, memory, 'engine')
    try {
      for (const exp of c.expected.cells) {
        const k = kindOf(exp.ref)
        const fill = observation?.fill(exp.ref)
        const visualKind = k.covering.some(m => m.ruleType === 'iconSet') ? 'icon' : k.covering.some(m => m.ruleType === 'dataBar') ? 'bar' : undefined
        const measured = visualKind && observation?.visual ? observation.visual(exp.ref, visualKind) : undefined
        const visualToken = (v: NonNullable<typeof measured>): string => (v.kind === 'none' ? 'none' : v.kind === 'bar' ? 'bar' : `icon:${(v as { cls: string }).cls}`)
        const excel: Stage5Excel | undefined = fill ? { fill: fill.fill, strike: fill.strike, ...(measured ? { visual: visualToken(measured) } : {}) } : undefined
        let production: Stage5Production = { ...blank }
        if (service) {
          const { row, col } = refRC(exp.ref)
          const lease = await service.prepareCf({ stamp, sourceRects: [{ firstRow: row, lastRow: row, firstCol: col, lastCol: col }], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, source)
          const r = lease.value.at(row, col)
          lease.release()
          if (r.kind === 'unresolved') production = { fill: 'none', strike: false, unresolved: ['ALL'], visual: 'withheld' }
          else if (r.kind === 'ready') {
            const st = resolveCfCellStyle(undefined, r.value)
            const unresolved = [...(r.value.unresolvedWrites ?? [])].sort()
            const v = r.value.visual
            // A zero-length bar (the column minimum) draws nothing, as Excel's measurement reads it; a withheld visual slot
            // (`dataBar`/`iconSet` unresolved) is `withheld`.
            const bar = v?.dataBar as { startFraction?: number; endFraction?: number } | undefined
            const drawnBar = bar !== undefined && Math.abs((bar.endFraction ?? 0) - (bar.startFraction ?? 0)) > 1e-9
            const visual = drawnBar ? 'bar' : v?.iconSet ? `icon:${v.iconSet.set}:${v.iconSet.iconId}` : unresolved.some(w => w === 'ALL' || w === 'dataBar' || w === 'iconSet') ? 'withheld' : 'none'
            production = { fill: st.fillColor ?? 'none', strike: st.strike === true, unresolved, visual }
          }
        }
        // Production names icons by family:id, the capture by glyph class; map through the published class table.
        if (production.visual.startsWith('icon:') && excel?.visual?.startsWith('icon:')) {
          const { CF_ICON_GLYPH_CLASSES } = await import('../../src/xlsx/cf/iconGlyphs')
          const cls = CF_ICON_GLYPH_CLASSES[production.visual.slice(5)]
          if (cls) production = { ...production, visual: `icon:${cls}` }
        }
        const justification = production.unresolved.length > 0 || production.visual === 'withheld' ? stage5Justifications(c, exp.ref) : []
        const verdict = stage5CellAgreement(production, excel, justification)
        out.push({
          caseId: c.id, ref: exp.ref, status: c.status, compared: true,
          evidenceKind: k.kind, capabilities: k.formulaCaps, excel, production, justification,
          agrees: verdict.agrees, ...(verdict.disagreement ? { disagreement: verdict.disagreement } : {}),
          mode: 'resident', stamp: { cfGen: stamp.cfGen, workbookValueGen: stamp.workbookValueGen }, bytePeak: 0, workUnits: 0,
        })
      }
    } finally { service?.dispose() }
    void describeS5Visual
    const units = budgets.reduce((n, b) => n + b.units, 0)
    for (const r of out) if (r.caseId === c.id) { r.bytePeak = memory.peakBytes; r.workUnits = units }
  }
  return out
}
type CfWorkBudgetLike = { readonly units: number }
type CfObservedIconRow = { ref: string; cellLeftPx: number; cellTopPx: number; icon: null | { leftPx: number; topPx: number; rightPx: number; bottomPx: number; glyphClass: string } }

export interface CycleSample {
  cycles: number
  reads: number
  maxPeakBytes: number
  maxNetBytes: number
  finalNetBytes: number
  /** Bar states the ready reads carried (Stage 3 cycles read a bar range). */
  barStates: number
  /** Icon states the ready reads carried (Stage 4 cycles read an icon range). */
  iconStates: number
  wallMs: number
}

/** Open/read/cancel/close cycles against a real paged session. A fresh memory
 * observes every round: peak proves real work, and complete teardown (store
 * release for shared definitions plus session release for pages) must return
 * to a zero baseline — stability alone is not cleanup. Pre-aborted reads must
 * throw cancellation; a mid-flight abort must settle without charge and leave
 * the session usable. */
export async function pressureCycles(rounds = 20, options: { fixture?: string; rect?: { r0: number; r1: number; c0: number; c1: number }; requireBars?: boolean; requireIcons?: boolean } = {}): Promise<CycleSample> {
  const started = Date.now()
  const { openImportSession } = await import('../../src/xlsx/storage/pagedImport')
  const { createPagedWorkbookSource } = await import('../../src/xlsx/view/pagedSource')
  const saved = readFileSync(join(REPO_ROOT, options.fixture ?? 'tests/fixtures/xlsx-cf/stage1-cellis-literals.xlsx'))
  const first = options.rect ?? { r0: 0, r1: 0, c0: 0, c1: 0 }
  let barStates = 0
  let iconStates = 0
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
        rects: [first],
        stamp: stamp('audit'),
        grantBytes: 2 * 1024 * 1024,
      }, memory, { signal: new AbortController().signal })
      if (!ready || ready.kind !== 'ready') throw new Error(`cycle ${i}: expected ready read`)
      const bars = ready.states.filter(s => s.bar !== undefined).length
      if (options.requireBars && bars === 0) throw new Error(`cycle ${i}: the ready read carried no bar state`)
      barStates += bars
      const icons = ready.states.filter(s => s.icon !== undefined).length
      if (options.requireIcons && icons === 0) throw new Error(`cycle ${i}: the ready read carried no icon state`)
      iconStates += icons
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
  return { cycles: rounds, reads, maxPeakBytes: maxPeak, maxNetBytes: maxNet, finalNetBytes: finalNet, barStates, iconStates, wallMs: Date.now() - started }
}

export interface NoCfBaseline {
  paintReserves: number
  sourceReads: number
  scopeBuilt: boolean
  /** Glyph decodes during the no-CF preparations (Stage 4: zero icon work). */
  glyphDecodes?: number
}

/** No-CF baseline: same shapes without conditional formatting do zero CF work. */
export async function baselineNoCf(rule: Record<string, unknown> = cellIsRule(0, { firstRow: 0, lastRow: 99, firstCol: 0, lastCol: 0 }) as unknown as Record<string, unknown>): Promise<NoCfBaseline> {
  const { prepareCfPaint } = await import('../../src/xlsx/cf/render')
  const memory = createCfMemory()
  let reserves = 0
  let reads = 0
  const rawReserve = memory.reserve.bind(memory)
  ;(memory as { reserve: typeof rawReserve }).reserve = (...args: Parameters<typeof rawReserve>) => { reserves++; return rawReserve(...args) }
  const sheet = numericSheet(100, rule)
  delete (sheet as { conditionalFormatting?: unknown }).conditionalFormatting
  const store = { get stamp(): never { throw new Error('no-CF stamp read') }, readRows: (...args: unknown[]) => { reads++; throw new Error(`no-CF read ${String(args)}`) } }
  let scopeBuilt = false
  const { cfDecodeCalls } = await import('../../src/xlsx/cf/iconGlyphs')
  const decodesBefore = cfDecodeCalls()
  for (let i = 0; i < 5; i++) {
    const batch = await prepareCfPaint(sheet, { memory, store: store as never })
    if (batch !== undefined) scopeBuilt = true
    batch?.dispose?.()
  }
  return { paintReserves: reserves, sourceReads: reads, scopeBuilt, glyphDecodes: cfDecodeCalls() - decodesBefore }
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

/** Task 39 fix round 1 (F4): writes `content` to a temporary sibling, then renames it over `target` (one atomic step). */
export function writeFileAtomic(target: string, content: string): void {
  const tmp = `${target}.tmp-${process.pid}-${Date.now()}`
  try {
    writeFileSync(tmp, content, 'utf8')
    renameSync(tmp, target)
  } finally {
    if (existsSync(tmp)) rmSync(tmp, { force: true })
  }
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
  if (stage !== 1 && stage !== 2 && stage !== 3 && stage !== 4 && stage !== 5) {
    // eslint-disable-next-line no-console
    console.error('Only --stage 1, --stage 2, --stage 3, --stage 4 and --stage 5 are implemented')
    process.exit(2)
  }
  const only = args.includes('--case') ? args[args.indexOf('--case') + 1] : undefined
  const outDir = args.includes('--out') ? resolve(args[args.indexOf('--out') + 1]!) : undefined
  const manifest = JSON.parse(readFileSync(CF_MANIFEST_PATH, 'utf8')) as CfManifest
  const report = stage === 1
    ? auditStage1Manifest(manifest, { rootDir: REPO_ROOT })
    : stage === 2
      ? auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
      : stage === 3
        ? auditStage3Manifest(manifest, { rootDir: REPO_ROOT })
        : stage === 4
          ? auditStage4Manifest(manifest, { rootDir: REPO_ROOT })
          : auditStage5Manifest(manifest, { rootDir: REPO_ROOT })
  // Task 39 fix round 1 (F4): `--write` renders now and writes only after every check below has passed, atomically
  // (a temporary sibling and one rename). Any failure (a thrown scenario, ok=false, an incomplete stage) leaves the
  // target byte-unchanged. `--target` redirects the write (tests); CF_AUDIT_INJECT_FAILURE=after-render is a test seam.
  let pendingWrite: { target: string; content: string; summary: string } | undefined
  if (args.includes('--write')) {
    const target = args.includes('--target') ? resolve(args[args.indexOf('--target') + 1]!) : resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts')
    const committed = resolve(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts')
    const content = renderVerifiedCapabilitiesSource(manifest, REPO_ROOT, {
      committedSource: existsSync(committed) ? readFileSync(committed, 'utf8') : undefined,
      publishStage3: args.includes('--publish-stage3'),
      publishStage4: args.includes('--publish-stage4'),
      publishStage5: args.includes('--publish-stage5'),
    })
    if (process.env.CF_AUDIT_INJECT_FAILURE === 'after-render') throw new Error('injected failure after render (test seam)')
    const count = (name: string): number => (new RegExp(`export const ${name}: ReadonlySet<string> = new Set\\(\\[([\\s\\S]*?)\\]\\)`).exec(content)?.[1]?.match(/'[^']+'/g) ?? []).length
    const s3 = content.includes('VERIFIED_STAGE3_CAPABILITIES') ? `, ${count('VERIFIED_STAGE3_CAPABILITIES')} S3${args.includes('--publish-stage3') ? ' published' : ' as committed'}` : ', S3 unpublished'
    const s4 = content.includes('VERIFIED_STAGE4_CAPABILITIES') ? `, ${count('VERIFIED_STAGE4_CAPABILITIES')} S4${args.includes('--publish-stage4') ? ' published' : ' as committed'}` : ', S4 unpublished'
    const s5 = content.includes('VERIFIED_STAGE5_CAPABILITIES') ? `, ${count('VERIFIED_STAGE5_CAPABILITIES')} S5${args.includes('--publish-stage5') ? ' published' : ' as committed'}` : ', S5 unpublished'
    const detail = ` (${count('VERIFIED_STAGE1_CAPABILITIES')} S1, ${count('VERIFIED_STAGE2_CAPABILITIES')} S2${s3}${s4}${s5})`
    pendingWrite = { target, content, summary: `${count('VERIFIED_CAPABILITIES')} verified capabilities${detail}` }
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
    const ab = r3.barAutoBounds
    // eslint-disable-next-line no-console
    console.log(`bar automatic bounds (S3-AUTO-BOUNDS): autoMin=${ab.autoMin?.rule ?? 'none'} autoMax=${ab.autoMax?.rule ?? 'none'}`)
    const bb = Object.entries(r3.barGradientBorder)
    // eslint-disable-next-line no-console
    console.log(`bar bordered gradient (S3-GRAD-BORDER): ${bb.length ? bb.map(([k, v]) => `${k}${v?.endBlend !== undefined ? ` endBlend=${v.endBlend}` : ''}`).join(', ') : 'none selected'}`)
  }
  // eslint-disable-next-line no-console
  console.log('--- pressure ---')

  let stage1Pressure: { whole: PressureSample; shared: PressureSample; bounds: ReturnType<typeof pressureRangeBoundaries>; refusedSheet: PressureSample; stats: PressureSample } | undefined
  let stage2Pressure: { scaleShared: PressureSample; percentile: PressureSample; extents: PressureSample } | undefined
  let stage4Pressure: Record<string, unknown> | undefined
  let stage5Pressure: Record<string, unknown> | undefined
  let stage3Pressure: { barShared: Awaited<ReturnType<typeof pressureBar1MRowSharedScan>>; barPercentile: PressureSample; barExtents: PressureSample; barDistinct: PressureSample; barUnitCap: Awaited<ReturnType<typeof pressureBarUnitCap>> } | undefined

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
  } else if (stage === 3) {
    const barShared = await pressureBar1MRowSharedScan()
    // eslint-disable-next-line no-console
    console.log(`${barShared.outcome} ${barShared.name}: peak=${barShared.peakBytes} net=${barShared.netBytes} statisticsReads=${barShared.statisticsReads} valueReads=${barShared.valueReads} barsPainted=${barShared.cellsWithResults}/40 estimatorRefused100=${barShared.estimatorRefusedAll} (${barShared.wallMs}ms context only)`)
    checkPressureSample(barShared, { expectOutcome: 'admitted' })
    if (barShared.statisticsReads !== 100) throw new Error(`200 identical scopes must share one 100-cell scan, read ${barShared.statisticsReads}`)
    if (barShared.cellsWithResults !== 40) throw new Error(`the first bar must paint all 40 requested cells, painted ${barShared.cellsWithResults}`)
    if (!barShared.estimatorRefusedAll) throw new Error('the 100-cell text request of 200 rules must refuse at the byte estimate')
    if (barShared.valueReads > 1000) throw new Error(`preparation read ${barShared.valueReads} values: the scan is not shared`)
    const barPercentile = await pressureRefusedPercentileBarWithMinMax()
    // eslint-disable-next-line no-console
    console.log(`${barPercentile.outcome} ${barPercentile.name}: peak=${barPercentile.peakBytes} net=${barPercentile.netBytes} reads=${barPercentile.sourceReads} barsPainted=${barPercentile.cellsWithResults} (${barPercentile.wallMs}ms context only)`)
    checkPressureSample(barPercentile, { expectOutcome: 'admitted' })
    if (barPercentile.cellsWithResults !== 10) throw new Error(`the min/max bar must paint 10 cells, painted ${barPercentile.cellsWithResults}`)
    const barExtents = await pressureBarWholeColumnExtents()
    // eslint-disable-next-line no-console
    console.log(`${barExtents.outcome} ${barExtents.name}: peak=${barExtents.peakBytes} net=${barExtents.netBytes} reads=${barExtents.sourceReads} barsPainted=${barExtents.cellsWithResults} (${barExtents.wallMs}ms context only)`)
    checkPressureSample(barExtents, { expectOutcome: 'admitted' })
    if (barExtents.cellsWithResults !== 10) throw new Error(`the whole-column bar must paint its 10 stored cells, painted ${barExtents.cellsWithResults}`)
    const barDistinct = await pressureBarDistinctRanges()
    // eslint-disable-next-line no-console
    console.log(`${barDistinct.outcome} ${barDistinct.name}: peak=${barDistinct.peakBytes} net=${barDistinct.netBytes} reads=${barDistinct.sourceReads} (${barDistinct.wallMs}ms context only; the 10-rect batch admitted and painted 10)`)
    checkPressureSample(barDistinct, { expectOutcome: 'refused' })
    const barUnitCap = await pressureBarUnitCap()
    // eslint-disable-next-line no-console
    console.log(`${barUnitCap.name}: ${barUnitCap.admittedCells} cells admitted (reads=${barUnitCap.admittedReads}, units=${barUnitCap.admittedUnits}), ${barUnitCap.refusedCells} cells refused ${barUnitCap.refusedReason} (reads=${barUnitCap.refusedReads}, units=${barUnitCap.refusedUnits}) net=${barUnitCap.netBytes} (${barUnitCap.wallMs}ms context only)`)
    if (barUnitCap.netBytes !== 0) throw new Error(`unit-cap scans leaked ${barUnitCap.netBytes} B`)
    stage3Pressure = { barShared, barPercentile, barExtents, barDistinct, barUnitCap }
  } else if (stage === 4) {
    const line = (p: PressureSample, extra = ''): void => {
      // eslint-disable-next-line no-console
      console.log(`${p.outcome} ${p.name}: peak=${p.peakBytes} net=${p.netBytes} reads=${p.sourceReads} cells=${p.cellsWithResults}${extra} (${p.wallMs}ms context only)`)
    }
    const iconShared = await pressureIcon1MRowSharedScan()
    line(iconShared, ` statisticsReads=${iconShared.statisticsReads} iconOnlyCells=${iconShared.iconOnlyCells}`)
    checkPressureSample(iconShared, { expectOutcome: 'admitted' })
    if (iconShared.statisticsReads !== 100) throw new Error(`300 identical scopes must share one 100-cell scan, read ${iconShared.statisticsReads}`)
    if (iconShared.cellsWithResults !== 40 || iconShared.iconOnlyCells !== 40) throw new Error(`the first icon set must evaluate all 40 requested cells, got ${iconShared.cellsWithResults}/${iconShared.iconOnlyCells}`)
    const iconPercentile = await pressureRefusedPercentileIconWithNum()
    line(iconPercentile, ` percentileWithheld=${iconPercentile.percentileWithheld} controlPercentileIcons=${iconPercentile.controlPercentileIcons}`)
    checkPressureSample(iconPercentile, { expectOutcome: 'admitted' })
    const iconExtents = await pressureIconWholeColumnExtents()
    line(iconExtents)
    checkPressureSample(iconExtents, { expectOutcome: 'admitted' })
    if (iconExtents.cellsWithResults !== 10) throw new Error(`the whole-column icon set must evaluate its 10 stored cells, got ${iconExtents.cellsWithResults}`)
    const glyphs = await pressureIconGlyphAdmission()
    // eslint-disable-next-line no-console
    console.log(`${glyphs.name}: classes=${glyphs.classes} admitted=${glyphs.admitted} decodes=${glyphs.decodes} withheldDecoded=${glyphs.withheldDecoded} refusedDecodes=${glyphs.refusedDecodes} peak=${glyphs.peakBytes} net=${glyphs.netBytes} (${glyphs.wallMs}ms context only)`)
    if (glyphs.decodes !== glyphs.admitted || glyphs.withheldDecoded !== 0 || glyphs.refusedDecodes !== 0 || glyphs.netBytes !== 0) throw new Error('glyph admission decoded outside its admission or leaked')
    const iconDistinct = await pressureIconDistinctRanges()
    line(iconDistinct)
    checkPressureSample(iconDistinct, { expectOutcome: 'refused' })
    const unitCap = await pressureBarUnitCap()
    // eslint-disable-next-line no-console
    console.log(`statistics unit cap (family-agnostic scan, shared by icons): ${unitCap.admittedCells} cells admitted (units=${unitCap.admittedUnits}), ${unitCap.refusedCells} refused ${unitCap.refusedReason} (reads=${unitCap.refusedReads}) net=${unitCap.netBytes}`)
    if (unitCap.netBytes !== 0) throw new Error(`unit-cap scans leaked ${unitCap.netBytes} B`)
    const caseRecords = await stage4CaseRecords(manifest, REPO_ROOT)
    // eslint-disable-next-line no-console
    console.log(`stage 4 case records: ${caseRecords.length} rows over ${new Set(caseRecords.map(r => r.caseId)).size} measured cases (resident mode); painted ${caseRecords.filter(r => r.produced.paints).length}`)
    stage4Pressure = { iconShared, iconPercentile, iconExtents, glyphs, iconDistinct, unitCap, caseRecords }
  } else if (stage === 5) {
    const r5 = report as Stage5AuditReport
    // eslint-disable-next-line no-console
    console.log(`stage 5 reused captures (S5-STUB): ${Object.entries(r5.reused).map(([k, v]) => `${k}=${v}`).join(', ')}`)
    // Task 39: the formula resource scenarios (scripts/cf-refs/pressure-stage5.ts), each with zero unowned live bytes.
    const p5 = await import('./pressure-stage5')
    const fail = (msg: string): never => { throw new Error(`stage 5 pressure: ${msg}`) }
    const shared = await p5.pressureFormulaSharedSnapshot()
    // eslint-disable-next-line no-console
    console.log(`${shared.name}: rules=${shared.rules} attempts=${shared.attempts} snapshotEntries=${shared.snapshotEntries}/${shared.targetCells} ready=${shared.ready} peak=${shared.peakBytes} net=${shared.netBytes} (${shared.wallMs}ms context only)`)
    if (shared.attempts !== 1 || shared.snapshotEntries !== shared.targetCells || shared.ready !== shared.rules * shared.targetCells) fail('100 rules must share one snapshot of the target cells')
    const hoists = await p5.pressureFormulaHoists()
    // eslint-disable-next-line no-console
    console.log(`${hoists.name}: visible=${hoists.visibleCells} columnScans=${hoists.columnScans} clamped=${hoists.clampedScans} calls=${hoists.calls} maxCallUnits=${hoists.maxCallUnits} refused=${hoists.refused}; ~1M populated: ${hoists.millionOutcomes.join(',')} calls=${hoists.millionCalls}; relative SUM(A1:A200): ${hoists.relativeOutcomes.join(',')} calls=${hoists.relativeCalls}; peak=${hoists.peakBytes} net=${hoists.netBytes} (${hoists.wallMs}ms context only)`)
    if (hoists.columnScans !== 2 || hoists.clampedScans !== 2 || hoists.refused !== 0 || hoists.maxCallUnits > p5.CHECKPOINT_UNITS) fail('hoists must scan the populated column once per rule and keep every call within the checkpoint')
    if (hoists.millionOutcomes.some(o => o !== 'unresolved:cf-resource-limit') || hoists.millionCalls !== 0) fail('a ~1M populated column must refuse at the scan budget')
    if (hoists.relativeOutcomes.join() !== 'unresolved:cf-resource-limit' || hoists.relativeCalls !== 0) fail('a relative-range SUM(A1:A200) must refuse before any call')
    const cap = await p5.pressureFormulaSnapshotCap()
    // eslint-disable-next-line no-console
    console.log(`${cap.name}: cap=${cap.cap} atCap=${cap.atCap} (entries=${cap.atCapEntries}) cap+1=${cap.overCap} peak=${cap.peakBytes} net=${cap.netBytes} (${cap.wallMs}ms context only)`)
    if (cap.atCap !== 'ready' || cap.atCapEntries !== cap.cap || cap.overCap !== 'unresolved:cf-resource-limit') fail('the snapshot cap must admit exactly cap cells')
    const sheets = await p5.pressureFormulaReferencedSheets()
    // eslint-disable-next-line no-console
    console.log(`${sheets.name}: 16=${sheets.at16} 17(compile)=${sheets.compile17} union18=${sheets.union18.join(',')} net=${sheets.netBytes}`)
    if (sheets.at16 !== 'ready' || !sheets.compile17.startsWith('refused:') || sheets.union18.join() !== 'ready,unresolved:cf-resource-limit') fail('16 sheets admit, 17 refuse')
    const beside = await p5.pressureFormulaRefusedBesideReady()
    // eslint-disable-next-line no-console
    console.log(`${beside.name}: ready=${beside.readyCells}/10 refused=${beside.refusedCells}/10 (${beside.refusedReasons.join(',')}) net=${beside.netBytes}`)
    if (beside.readyCells !== 10 || beside.refusedCells !== 10) fail('a refused footprint must not take the ready rule down')
    const cyc = await p5.pressureFormulaCrossSheetCycles(20)
    // eslint-disable-next-line no-console
    console.log(`${cyc.name}: cycles=${cyc.cycles} ready=${cyc.ready} pending=${cyc.pending} closedMidRead=${cyc.cancelled} maxNetBetween=${cyc.maxNetBetween} net=${cyc.netBytes}`)
    if (cyc.ready !== 20 || cyc.pending !== 20 || cyc.cancelled !== 20 || cyc.maxNetBetween !== 0) fail('cross-sheet cycles must hold nothing between close and open')
    for (const p of [shared, hoists, cap, sheets, beside, cyc]) if (p.netBytes !== 0) fail(`${p.name} leaked ${p.netBytes} B`)
    const caseRecords5 = await stage5CaseRecords(manifest, REPO_ROOT)
    const compared5 = caseRecords5.filter(r => r.compared)
    const disagree = compared5.filter(r => !r.agrees)
    // eslint-disable-next-line no-console
    console.log(`stage 5 case records: ${caseRecords5.length} rows over ${new Set(caseRecords5.map(r => r.caseId)).size} cases (${new Set(compared5.map(r => r.caseId)).size} measured and compared, resident); formula cells ${compared5.filter(r => r.evidenceKind === 'display-format').length}, painted as Excel with nothing unresolved ${compared5.filter(r => r.evidenceKind === 'display-format' && r.production.unresolved.length === 0 && r.production.fill === r.excel?.fill).length}, with justified unresolved slots ${compared5.filter(r => r.production.unresolved.length > 0).length}, visual evidence compared ${compared5.filter(r => r.excel?.visual !== undefined).length}, disagreeing ${disagree.length}${disagree.length ? ` (${disagree.slice(0, 5).map(r => `${r.caseId}:${r.ref} ${r.disagreement}`).join('; ')})` : ''}`)
    if (disagree.length > 0) fail(`${disagree.length} Stage 5 cells disagree with Excel or are withheld without a recorded justification`)
    const noFormula = await p5.pressureNoFormulaBaseline()
    // eslint-disable-next-line no-console
    console.log(`no-formula baseline (literal CF kept): compiles=${noFormula.literal.compiles} snapshots=${noFormula.literal.snapshots} clockReads=${noFormula.literal.clockReads} painted=${noFormula.literal.painted}; control (one expression): compiles=${noFormula.control.compiles} snapshots=${noFormula.control.snapshots} clockReads=${noFormula.control.clockReads}`)
    if (noFormula.literal.compiles + noFormula.literal.snapshots + noFormula.literal.clockReads !== 0 || noFormula.literal.painted === 0) fail('the no-formula baseline did formula work (or painted nothing)')
    if (noFormula.control.compiles === 0 || noFormula.control.snapshots === 0 || noFormula.control.clockReads === 0) fail('the no-formula positive control moved no counter')
    stage5Pressure = { shared, hoists, cap, sheets, beside, cyc, caseRecords: caseRecords5, noFormula }
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

  // Stage 3 cycles read a bar range of a Stage 3 workbook (B2:B12 of the default-lengths case), so every ready read
  // carries bar states.
  const cycles = stage === 3
    ? await pressureCycles(20, { fixture: 'tests/fixtures/xlsx-cf/stage3-bar-lengths-default.xlsx', rect: { r0: 1, r1: 11, c0: 1, c1: 1 }, requireBars: true })
    : stage === 4
      // Stage 4 cycles read the thresholds case's icon range (B2:B12), so every ready read carries icon states.
      ? await pressureCycles(20, { fixture: 'tests/fixtures/xlsx-cf/stage4-icon-thresholds.xlsx', rect: { r0: 1, r1: 11, c0: 1, c1: 1 }, requireIcons: true })
      : await pressureCycles(20)
  // eslint-disable-next-line no-console
  console.log(`cycles: ${cycles.cycles} open/read/cancel/close, reads=${cycles.reads} barStates=${cycles.barStates} iconStates=${cycles.iconStates} maxPeak=${cycles.maxPeakBytes} maxNet=${cycles.maxNetBytes} finalNet=${cycles.finalNetBytes} (released baseline, ${cycles.wallMs}ms context only)`)
  if (cycles.maxPeakBytes <= 0) throw new Error('cycles observed no allocation')
  if (cycles.reads < cycles.cycles * 4) throw new Error(`cycles issued ${cycles.reads} reads, expected at least ${cycles.cycles * 4} (ready + cancelled + mid-flight + retry per cycle)`)

  const baseline = await baselineNoCf(stage === 3 ? barRule(0, { firstRow: 0, lastRow: 99, firstCol: 0, lastCol: 0 }) as unknown as Record<string, unknown>
    : stage === 4 ? iconRuleOf(0, { firstRow: 0, lastRow: 99, firstCol: 0, lastCol: 0 }) as unknown as Record<string, unknown> : undefined)
  // eslint-disable-next-line no-console
  console.log(`no-cf baseline: reserves=${baseline.paintReserves} reads=${baseline.sourceReads} scopeBuilt=${baseline.scopeBuilt} glyphDecodes=${baseline.glyphDecodes}`)
  if (baseline.paintReserves !== 0 || baseline.sourceReads !== 0 || baseline.scopeBuilt || (baseline.glyphDecodes ?? 0) !== 0) throw new Error('no-CF path did CF work')

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
      : stage === 3
        ? { report, ...stage3Pressure, distinct, cycles, baseline, eager }
        : stage === 4
          ? { report, publication: stage4Publication(report as Stage4AuditReport), ...stage4Pressure, distinct, cycles, baseline, eager }
          : stage === 5
            ? { report, publication: stage5Publication(report as Stage5AuditReport), ...stage5Pressure, distinct, cycles, baseline, eager }
            : { report, ...stage2Pressure, distinct, cycles, baseline, eager }
    writeFileSync(join(outDir, `stage${stage}-audit.json`), JSON.stringify(outData, null, 2))
  }

  if (pendingWrite) {
    if (ok && report.stageComplete) {
      writeFileAtomic(pendingWrite.target, pendingWrite.content)
      // eslint-disable-next-line no-console
      console.log(`wrote ${pendingWrite.summary} to ${relative(REPO_ROOT, pendingWrite.target)}`)
    } else {
      // eslint-disable-next-line no-console
      console.log(`not written: the audit failed, ${relative(REPO_ROOT, pendingWrite.target)} is unchanged`)
    }
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
