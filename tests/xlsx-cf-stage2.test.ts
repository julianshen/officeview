// @vitest-environment node
/**
 * Task 17: Stage 2 validation, resource audit and PR evidence.
 *
 * Verifies that:
 * - every enabled scale capability has verified evidence;
 * - every unresolved scale shape refuses atomically on resident and wire paths;
 * - manifest mutations fail the audit (dropping fillSource, marking pending variant
 *   supported, enabling interpolation variant without its case, dropped hash/case);
 * - Stage 2 pressure scenarios enforce algorithmic bounds, sharing, extent clamping,
 *   and zero unowned live bytes across cancel/close cycles.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join, resolve } from 'node:path'
import {
  auditStage2Manifest,
  stageGateStatus,
  pressureDistinctRanges,
  pressureCycles,
  baselineNoCf,
  pressureScale1MRowSharedStatistics,
  pressureRefusedPercentileWithMinMax,
  pressureScaleWholeColumnExtents,
  checkPressureSample,
  type Stage2AuditReport,
} from '../scripts/cf-refs/audit'
import { CF_STAGE2_CAPABILITIES, CF_STAGE2_NONDISCRIMINATING, type CfManifest } from '../scripts/cf-refs/cases'
import { prepareCfTextScope, prepareCfWireScope } from '../src/xlsx/cf/text'
import { createCfService } from '../src/xlsx/cf/service'
import { createCfMemory, CF_LIMITS } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { VERIFIED_CAPABILITIES } from '../src/xlsx/cf/evidence'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import type { XlsxSheet } from '../src/xlsx/types'
import type { CfReadResult } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const MANIFEST_PATH = new URL('./fixtures/xlsx-cf/manifest.json', import.meta.url)

function loadManifest(): CfManifest {
  return JSON.parse(readFileSync(MANIFEST_PATH, 'utf8')) as CfManifest
}

function trackingBudget(): { budget: FormatBudget; net: () => number } {
  let held = 0
  const budget: FormatBudget = {
    reserve: bytes => { held += bytes },
    release: bytes => { held -= bytes },
    retain: (_targets, bytes, commit) => { held += bytes; commit() },
    settle: () => {},
  }
  return { budget, net: () => held }
}

async function sheetFromSaved(inputPath: string, savedPath: string): Promise<XlsxSheet> {
  void inputPath
  const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, savedPath)))
  const doc = await parseXlsx(await OfficePackage.load(bytes))
  const sheet = doc.sheets.find(s => s.name === 'Sheet1') ?? doc.sheets[0]!
  if (!sheet) throw new Error(`no sheet in ${savedPath}`)
  return sheet as unknown as XlsxSheet
}

function baseOf(sheet: XlsxSheet, row: number, col: number): unknown {
  for (const band of sheet.rows) {
    const cell = band.cells.find(candidate => candidate.row === row && candidate.col === col)
    if (cell) return cell.style
  }
  return undefined
}

describe('every enabled scale capability has verified evidence', () => {
  it('the manifest audit passes integrity with full catalog coverage and structural evidence', () => {
    const manifest = loadManifest()
    const report: Stage2AuditReport = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect(report.ok).toBe(true)
    const covered = new Set(report.perCapability.map(entry => entry.capabilityId))
    const missing = CF_STAGE2_CAPABILITIES.map(entry => entry.capabilityId).filter(id => !covered.has(id))
    expect(missing).toEqual([])
    for (const entry of report.perCase) {
      if (entry.status === 'measured') expect(entry.structuralValid, entry.id).toBe(true)
    }
    const unsupported = report.perCase.find(entry => entry.id === 'stage2-scale-unresolved-shapes')
    expect(unsupported?.status).toBe('unsupported')
    expect(unsupported?.code).toBe('unsupported-capability')
  })

  it('structural integrity with verified semantic evidence completes Stage 2', () => {
    const manifest = loadManifest()
    const report: Stage2AuditReport = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    expect(report.ok).toBe(true)
    expect(report.stageComplete).toBe(true)
    // 41 measured scale capabilities; 5 are measured but NON-DISCRIMINATING (Ruling F-2) and are reported unverified.
    expect(report.summary.verifiedCapabilities).toBe(36)
    expect(report.summary.unverifiedCapabilities).toBe(5)
    expect(report.blockedCapabilities.length).toBe(0)
  })

  it('VERIFIED_CAPABILITIES contains all verified Stage 2 scale capabilities', () => {
    const manifest = loadManifest()
    const report: Stage2AuditReport = auditStage2Manifest(manifest, { rootDir: REPO_ROOT })
    const auditVerified = new Set(
      report.perCapability.filter(c => c.disposition === 'verified').map(c => c.capabilityId),
    )
    expect(auditVerified.size).toBe(36)
    for (const cap of auditVerified) {
      expect(VERIFIED_CAPABILITIES.has(cap), `VERIFIED_CAPABILITIES must contain ${cap}`).toBe(true)
    }
    // The non-discriminating five are accounted for, carry a reason, and are never in the production set.
    const unverified = report.perCapability.filter(c => c.disposition === 'unverified-nondiscriminating').map(c => c.capabilityId).sort()
    expect(unverified).toEqual(Object.keys(CF_STAGE2_NONDISCRIMINATING).sort())
    for (const cap of unverified) {
      expect(CF_STAGE2_NONDISCRIMINATING[cap]!.length, `${cap} needs a reason`).toBeGreaterThan(40)
      expect(VERIFIED_CAPABILITIES.has(cap), `${cap} must not be production-verified`).toBe(false)
    }
  })

  it('production runtime evidence gate allows verified and refuses unverified/unsupported capabilities', () => {
    expect(cfRuntimeEvidence.allows('scale:interpolation-color-space')).toBe(true)
    expect(cfRuntimeEvidence.allows('scale:thresholds-min-max')).toBe(true)
    expect(cfRuntimeEvidence.allows('scale:colors-theme-tint')).toBe(true)
    expect(cfRuntimeEvidence.allows('scale:scope-whole-column')).toBe(true)
    expect(cfRuntimeEvidence.allows('scale:composition-dxf-p1-scale-p2')).toBe(true)

    expect(cfRuntimeEvidence.allows('scale:unresolved-2cfvo-3colors')).toBe(false)
    expect(cfRuntimeEvidence.allows('scale:unresolved-cfvo-without-type')).toBe(false)
    expect(cfRuntimeEvidence.allows('scale:unresolved-cfvo-formula')).toBe(false)
    expect(cfRuntimeEvidence.allows('scale:unresolved-active-present')).toBe(false)
    expect(cfRuntimeEvidence.allows('scale:unresolved-x14-standalone')).toBe(false)
    expect(cfRuntimeEvidence.allows('scale:nonexistent')).toBe(false)
  })
})

describe('stage gate status distinguishes verified completion from blocked evidence', () => {
  it('all verified completes; any supported-structural or unaccounted blocks', () => {
    expect(stageGateStatus([
      { capabilityId: 'scale:a', caseId: 'c', disposition: 'verified' },
      { capabilityId: 'scale:b', caseId: 'c', disposition: 'verified' },
    ])).toEqual({ complete: true, blocked: [] })
    expect(stageGateStatus([
      { capabilityId: 'scale:a', caseId: 'c', disposition: 'verified' },
      { capabilityId: 'scale:b', caseId: 'c', disposition: 'supported-structural' },
      { capabilityId: 'scale:c', caseId: 'u', disposition: 'refused-unsupported' },
    ])).toEqual({ complete: false, blocked: ['scale:b'] })
    expect(stageGateStatus([
      { capabilityId: 'scale:a', caseId: 'c', disposition: 'verified' },
      { capabilityId: 'scale:z', caseId: 'c', disposition: 'unaccounted' },
    ])).toEqual({ complete: false, blocked: ['scale:z'] })
  })
})

describe('every unresolved scale shape refuses atomically on both paths', () => {
  it('all product-marked unresolved rules withhold on resident and wire scopes', async () => {
    const manifest = loadManifest()
    const stage2Cases = manifest.cases.filter(c => c.stage === 2)
    const seen = new Map<string, { ruleType: string; reason: string }>()
    let unresolvedCount = 0
    let probedCells = 0

    for (const c of stage2Cases) {
      const sheet = await sheetFromSaved(c.input, c.saved)
      const rules = sheet.conditionalFormatting?.rules ?? []
      const unresolved = rules.filter(rule => {
        const definition = rule.definition as { kind?: string } | undefined
        return definition?.kind === 'unresolved'
      })
      if (unresolved.length === 0) continue

      const oracleMemory = createCfMemory()
      const oracle = createCfService(sheet.conditionalFormatting!, oracleMemory, { evidence: cfRuntimeEvidence })
      const gen = cfRulesVersion(sheet.conditionalFormatting)
      const stamp = { owner: 'cf-text', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: gen, viewGen: 0, structureGen: 0, clockGen: 0 }
      const source = residentSource(sheet, stamp, oracleMemory, 'engine')

      for (const rule of unresolved) {
        const definition = rule.definition as { type?: string; reason?: string }
        unresolvedCount++
        const key = `${definition.type ?? 'unknown'}:${definition.reason ?? 'unknown'}`
        if (!seen.has(key)) seen.set(key, { ruleType: definition.type ?? 'unknown', reason: definition.reason ?? 'unknown' })

        const rects = rule.ranges.map(r => ({ r0: r.firstRow, r1: Math.min(r.lastRow, r.firstRow + 4), c0: r.firstCol, c1: r.firstCol }))
        const lease = await oracle.prepareCf(
          { stamp, sourceRects: rects.map(r => ({ firstRow: r.r0, lastRow: r.r1, firstCol: r.c0, lastCol: r.c1 })), signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes },
          source,
          oracleMemory,
        )

        try {
          const cells: { row: number; col: number; kind: string; hasProperties: boolean }[] = []
          for (const r of rects) {
            for (let row = r.r0; row <= r.r1; row++) {
              const outcome = lease.value.at(row, r.c0)
              if (outcome.kind === 'unresolved' && outcome.reason === 'cf-outside-preparation') continue
              const hasProperties = outcome.kind === 'ready' && Object.keys(outcome.value.properties).length > 0
              cells.push({ row, col: r.c0, kind: outcome.kind, hasProperties })
              probedCells++
            }
          }

          expect(cells.length, `${c.id} pri${(rule as { priority: number }).priority} probed cells`).toBeGreaterThan(0)

          const { budget, net } = trackingBudget()
          const scope = await prepareCfTextScope(sheet, {
            rects,
            signal: new AbortController().signal,
            locale: 'en-US',
            budget,
          })

          if (cells.length > 0) expect(scope, `${c.id} resident scope`).toBeDefined()
          if (scope) {
            for (const cell of cells) {
              const where = `${c.id} resident ${cell.row},${cell.col}`
              const base = baseOf(sheet, cell.row, cell.col)
              const desc = scope.describe(cell.row, cell.col)
              const styled = scope.effectiveStyle(cell.row, cell.col, base as never)

              if (cell.kind === 'unresolved') {
                expect(desc?.text, where).toBe('Conditional formatting unavailable')
                expect(styled, where).toBe(base)
              }
            }
            scope.dispose()
          }
          expect(net()).toBe(0)
        } finally {
          lease.release()
        }
      }
      oracle.dispose()
      expect(oracleMemory.currentBytes).toBe(0)
    }

    expect(unresolvedCount).toBeGreaterThan(0)
    expect(probedCells).toBeGreaterThan(0)
  })

  it('parse-marked unresolved scale cells withhold through the real paged worker route', async () => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
    const manifest = loadManifest()
    const unresolvedCase = manifest.cases.find(c => c.id === 'stage2-scale-unresolved-shapes')!
    expect(unresolvedCase).toBeDefined()

    const sheet = await sheetFromSaved(unresolvedCase.input, unresolvedCase.saved)
    const unresolved = (sheet.conditionalFormatting?.rules ?? []).filter(rule => {
      const definition = (rule as { definition?: { kind?: string } }).definition
      return definition?.kind === 'unresolved'
    })
    expect(unresolved.length).toBeGreaterThan(0)

    const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, unresolvedCase.saved)))
    const session = await openImportSession(
      { kind: 'blob', blob: new Blob([bytes.buffer as ArrayBuffer]) },
      { docId: `stage2-wire-${unresolvedCase.id}`, profile: 'D' },
    )

    try {
      await session.done
      const source = createPagedWorkbookSource(session)
      const sheetId = source.sheets[0]!.sheetId
      const store = source.store(sheetId)
      const memory = createCfMemory()

      for (const rule of unresolved) {
        const ranges = (rule as unknown as { ranges: readonly { firstRow: number; lastRow: number; firstCol: number; lastCol: number }[] }).ranges
        const rects = ranges.map(r => ({ r0: r.firstRow, r1: Math.min(r.lastRow, r.firstRow + 4), c0: r.firstCol, c1: r.firstCol }))
        if (rects.length === 0) continue

        for (const r of rects) {
          await store.readRows?.({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1 }, undefined, { signal: new AbortController().signal, stamp: store.stamp } as never).catch(() => undefined)
        }

        const result = await store.readConditionalFormatting?.({
          rects,
          stamp: { owner: 'audit', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 },
          grantBytes: 2 * 1024 * 1024,
        }, memory, { signal: new AbortController().signal })

        if (result?.kind === 'ready') {
          const { budget: wireBudget, net: wireNet } = trackingBudget()
          const scope = await prepareCfWireScope(
            async () => ({ result: result as CfReadResult, release: () => {} }),
            rects,
            { locale: 'en-US', budget: wireBudget },
          )
          try {
            expect(scope).toBeDefined()
            for (const r of rects) {
              for (let row = r.r0; row <= r.r1; row++) {
                const base = baseOf(sheet, row, r.c0)
                const desc = scope!.describe(row, r.c0)
                const styled = scope!.effectiveStyle(row, r.c0, base as never)
                expect(desc?.text === undefined || desc.text === 'Conditional formatting unavailable').toBe(true)
                if (styled !== base) expect(desc?.text).toBe('Conditional formatting unavailable')
              }
            }
          } finally {
            scope?.dispose()
          }
          expect(wireNet()).toBe(0)
        }
      }
    } finally {
      await session.release()
    }
  }, 120000)
})

describe('manifest mutations fail the Stage 2 audit', () => {
  const rootDir = REPO_ROOT

  it('dropping fillSource fails', () => {
    const manifest = loadManifest()
    const targetCase = manifest.cases.find(c => c.id === 'stage2-scale-thresholds')!
    const targetCell = targetCase.expected.cells.find(cell => cell.scaleApplied)!
    delete (targetCell as unknown as Record<string, unknown>).fillSource
    const report = auditStage2Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
  })

  it('marking a pending variant supported fails', () => {
    const manifest = loadManifest()
    const targetCase = manifest.cases.find(c => c.id === 'stage2-scale-thresholds')!
    targetCase.expected.provenance = 'seed-hypothesis' as never
    const report = auditStage2Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
  })

  it('enabling an interpolation variant without its case fails', () => {
    const manifest = loadManifest()
    manifest.cases = manifest.cases.filter(c => c.id !== 'stage2-scale-interpolation')
    const report = auditStage2Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
    expect(report.missingCases).toContain('stage2-scale-interpolation')
  })

  it('dropped hash fails', () => {
    const manifest = loadManifest()
    const targetCase = manifest.cases.find(c => c.id === 'stage2-scale-colors')!
    delete (targetCase.expected as unknown as Record<string, unknown>).savedSha256
    const report = auditStage2Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
  })

  it('dropped case fails', () => {
    const manifest = loadManifest()
    manifest.cases = manifest.cases.filter(c => c.id !== 'stage2-scale-unresolved-shapes')
    const report = auditStage2Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
    expect(report.missingCases).toContain('stage2-scale-unresolved-shapes')
  })
})

describe('Stage 2 pressure scenarios and resource bounds', () => {
  it('a 1M-row DECLARED A:A scope with 100 stored cells and 100 scales shares one sparse scan (Lesson 3, M5)', async () => {
    const sample = await pressureScale1MRowSharedStatistics()
    checkPressureSample(sample, { expectOutcome: 'admitted' })
    expect(sample.sourceReads).toBe(100) // 1 scan of 100 stored cells
    expect(sample.netBytes).toBe(0)
  })

  it('an invalid percentile bound withholds its scale while min/max still renders (M5; resource refusal lives in the statistics tests)', async () => {
    const sample = await pressureRefusedPercentileWithMinMax()
    checkPressureSample(sample, { expectOutcome: 'admitted' })
    expect(sample.netBytes).toBe(0)
    expect(sample.cellsWithResults).toBeGreaterThan(0) // min/max rendered
  })

  it('whole-column scope CF ranges do not extend extents (Lesson 4)', async () => {
    const sample = await pressureScaleWholeColumnExtents()
    checkPressureSample(sample, { expectOutcome: 'admitted' })
    expect(sample.netBytes).toBe(0)
  })

  it('distinct-range pressure up to 1,000,000-unit pass cap', async () => {
    const distinct = await pressureDistinctRanges()
    checkPressureSample(distinct, { expectOutcome: 'refused' })
    expect(distinct.netBytes).toBe(0)
  })

  it('20 open/read/cancel/close cycles against Stage 2 workbook return to zero baseline', async () => {
    const cycles = await pressureCycles(20)
    expect(cycles.cycles).toBe(20)
    expect(cycles.finalNetBytes).toBe(0)
    expect(cycles.maxPeakBytes).toBeGreaterThan(0)
  })

  it('zero scale work or reads on a no-CF baseline', async () => {
    const baseline = await baselineNoCf()
    expect(baseline.paintReserves).toBe(0)
    expect(baseline.sourceReads).toBe(0)
    expect(baseline.scopeBuilt).toBe(false)
  })
})
