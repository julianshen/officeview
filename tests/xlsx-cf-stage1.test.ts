// @vitest-environment node
/**
 * Task 11: Stage 1 semantic manifest audit and both-paths atomic refusal.
 *
 * Oracle provenance: expected outcomes below come from the repository's own
 * structural verification (`verifyCfCase`) and the product's fail-closed
 * composition — never from Excel pixel captures (still `recipe-pending`;
 * screen-locked host) and never from seed-hypothesis cell values. Pixel
 * comparison stays pending until measured captures exist; this test pins
 * that unmeasured capabilities are never *claimed* as verified and that
 * every product-marked unresolved rule withholds atomically on both paths.
 *
 * RED first: imports `../scripts/cf-refs/audit`, which does not exist yet.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join, resolve } from 'node:path'
import { auditStage1Manifest, stageGateStatus, type Stage1AuditReport } from '../scripts/cf-refs/audit'
import { CF_STAGE1_CAPABILITIES, type CfManifest } from '../scripts/cf-refs/cases'
import { prepareCfTextScope, prepareCfWireScope } from '../src/xlsx/cf/text'
import { createCfService } from '../src/xlsx/cf/service'
import { createCfMemory, CF_LIMITS } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE1_CAPABILITIES } from '../src/xlsx/cf/evidence'
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

describe('every supported Stage 1 semantic capability has verified oracle evidence', () => {
  it('the manifest audit passes integrity with full catalog coverage and structural evidence', () => {
    const manifest = loadManifest()
    const report: Stage1AuditReport = auditStage1Manifest(manifest, { rootDir: REPO_ROOT })
    expect(report.ok).toBe(true)
    // Every catalog capability is accounted for: supported-structural or refused.
    const covered = new Set(report.perCapability.map(entry => entry.capabilityId))
    const missing = CF_STAGE1_CAPABILITIES.map(entry => entry.capabilityId).filter(id => !covered.has(id))
    expect(missing).toEqual([])
    // Every recipe-pending case carries verified structural evidence.
    for (const entry of report.perCase) {
      if (entry.status === 'recipe-pending') expect(entry.structuralValid, entry.id).toBe(true)
    }
    // The known-unsupported case refuses with the unsupported capability code.
    const unsupported = report.perCase.find(entry => entry.id === 'stage1-dxf-layering-2')
    expect(unsupported?.code).toBe('unsupported-capability')
  })

  it('structural integrity with verified semantic evidence completes Stage 1', () => {
    const manifest = loadManifest()
    const report: Stage1AuditReport = auditStage1Manifest(manifest, { rootDir: REPO_ROOT })
    // All 12 supported Stage 1 cases now carry verified independent semantic evidence
    // from host Excel captures: the stage gate reports complete.
    expect(report.ok).toBe(true)
    expect(report.stageComplete).toBe(true)
    expect(report.summary.verifiedCapabilities).toBe(110)
    expect(report.blockedCapabilities.length).toBe(0)
  })

  it('VERIFIED_CAPABILITIES matches the audit verified capabilities exactly (C1)', () => {
    const manifest = loadManifest()
    const report: Stage1AuditReport = auditStage1Manifest(manifest, { rootDir: REPO_ROOT })
    const auditVerified = new Set(
      report.perCapability.filter(c => c.disposition === 'verified').map(c => c.capabilityId),
    )
    const stage1Verified = VERIFIED_STAGE1_CAPABILITIES ?? VERIFIED_CAPABILITIES
    expect(stage1Verified).toEqual(auditVerified)
    expect(stage1Verified.size).toBe(110)
  })

  it('every evidence.allows literal under src/ resolves to a catalog ID (C1)', () => {
    const catalogIds = new Set(CF_STAGE1_CAPABILITIES.map(c => c.capabilityId))
    const cfDir = join(REPO_ROOT, 'src/xlsx/cf')
    const files = readdirSync(cfDir).filter(f => f.endsWith('.ts'))
    const literalPattern = /evidence\.allows\(\s*['"]([^'"]+)['"]/g
    const checkedLiterals: string[] = []

    for (const file of files) {
      const content = readFileSync(join(cfDir, file), 'utf8')
      let match: RegExpExecArray | null
      while ((match = literalPattern.exec(content)) !== null) {
        const id = match[1]!
        checkedLiterals.push(id)
        expect(catalogIds.has(id), `evidence.allows('${id}') in ${file} not found in catalog`).toBe(true)
      }
    }
    expect(checkedLiterals.length).toBeGreaterThan(15)
  })

  it('production runtime evidence gate allows verified and refuses unverified/unsupported capabilities (C1)', () => {
    expect(cfRuntimeEvidence.allows('cellIs:between-reversed')).toBe(true)
    expect(cfRuntimeEvidence.allows('text:containsText-canonical')).toBe(true)
    expect(cfRuntimeEvidence.allows('timePeriod:today')).toBe(true)
    expect(cfRuntimeEvidence.allows('layout:hidden-row-statistics')).toBe(true)
    expect(cfRuntimeEvidence.allows('priority:equal-priority-first')).toBe(true)

    expect(cfRuntimeEvidence.allows('dxf:unsupported-gradientFill')).toBe(false)
    expect(cfRuntimeEvidence.allows('dxf:unsupported-patternFill')).toBe(false)
    expect(cfRuntimeEvidence.allows('nonexistent:capability')).toBe(false)
  })
})

describe('stage gate status distinguishes verified completion from blocked evidence', () => {
  it('all verified completes; any supported-structural or unaccounted blocks', () => {
    expect(stageGateStatus([
      { capabilityId: 'a', caseId: 'c', disposition: 'verified' },
      { capabilityId: 'b', caseId: 'c', disposition: 'verified' },
    ])).toEqual({ complete: true, blocked: [] })
    expect(stageGateStatus([
      { capabilityId: 'a', caseId: 'c', disposition: 'verified' },
      { capabilityId: 'b', caseId: 'c', disposition: 'supported-structural' },
      { capabilityId: 'c', caseId: 'u', disposition: 'refused-unsupported' },
    ])).toEqual({ complete: false, blocked: ['b'] })
    expect(stageGateStatus([
      { capabilityId: 'a', caseId: 'c', disposition: 'verified' },
      { capabilityId: 'z', caseId: 'c', disposition: 'unaccounted' },
    ])).toEqual({ complete: false, blocked: ['z'] })
  })
})

describe('every unresolved capability refuses atomically on both paths', () => {
  it('all product-marked unresolved rules withhold on resident and wire scopes', async () => {
    const manifest = loadManifest()
    const stage1Cases = manifest.cases.filter(c => c.stage === 1 || !c.stage)
    const seen = new Map<string, { ruleType: string; reason: string }>()
    let unresolvedCount = 0
    let probedCells = 0
    for (const c of stage1Cases) {
      const sheet = await sheetFromSaved(c.input, c.saved)
      const rules = sheet.conditionalFormatting?.rules ?? []
      const unresolved = rules.filter(rule => {
        const definition = rule.definition as { kind?: string } | undefined
        return definition?.kind === 'unresolved'
      })
      if (unresolved.length === 0) continue
      // Oracle: the real composition over the same inputs both builders use.
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
          // Record the composed outcome per cell. The service lease is the
          // reference composition for exact per-cell agreement; the
          // fail-closed invariant below holds without consulting it.
          const cells: { row: number; col: number; kind: string; hasProperties: boolean; hasWrites: boolean }[] = []
          for (const r of rects) {
              for (let row = r.r0; row <= r.r1; row++) {
                const outcome = lease.value.at(row, r.c0)
                if (outcome.kind === 'unresolved' && outcome.reason === 'cf-outside-preparation') continue
              // A resolved colour-scale visual is a winning paint just like winning properties
              // (Stage 2): the fill must reach the effective style on every path.
              const hasProperties = outcome.kind === 'ready'
                && (Object.keys(outcome.value.properties).length > 0 || outcome.value.visual?.colorScale !== undefined)
              const writes = outcome.kind === 'ready' ? outcome.value.unresolvedWrites : undefined
              const hasWrites = writes !== undefined && (writes as ReadonlySet<string>).size !== undefined
                ? (writes as ReadonlySet<string>).size > 0
                : Array.isArray(writes) ? writes.length > 0 : false
              cells.push({ row, col: r.c0, kind: outcome.kind, hasProperties, hasWrites })
              probedCells++
            }
          }
          // Observable contract per composed outcome:
          // - unresolved: announced once, style is exactly base (same ref).
          // - ready: withheld aspects announce unavailable; winning properties
          //   apply (style !== base) iff the outcome has properties.
          // - no-match: silent, base style.
          const expectCell = (
            describe: ((row: number, col: number) => { text?: string } | undefined),
            effectiveStyle: (row: number, col: number, base: unknown) => unknown,
            cell: { row: number; col: number; kind: string; hasProperties: boolean; hasWrites: boolean },
            base: unknown,
            where: string,
          ) => {
            const desc = describe(cell.row, cell.col)
            const styled = effectiveStyle(cell.row, cell.col, base)
            if (cell.kind === 'unresolved') {
              expect(desc?.text, where).toBe('Conditional formatting unavailable')
              expect(styled, where).toBe(base)
            } else if (cell.kind === 'ready') {
              if (cell.hasWrites || !cell.hasProperties) {
                if (desc !== undefined) expect(desc.text, where).toBe('Conditional formatting unavailable')
              } else {
                expect(desc?.text, where).not.toBe('Conditional formatting unavailable')
              }
              if (cell.hasProperties) expect(styled, where).not.toBe(base)
              else expect(styled, where).toBe(base)
            } else {
              expect(desc, where).toBeUndefined()
              expect(styled, where).toBe(base)
            }
          }
          expect(cells.length, `${c.id} pri${(rule as { priority: number }).priority} probed cells`).toBeGreaterThan(0)
          // Resident path must agree with the oracle outcome per cell.
          {
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
                expectCell(
                  (row, col) => scope.describe(row, col) ?? undefined,
                  (row, col, base) => scope.effectiveStyle(row, col, base as never),
                  cell,
                  baseOf(sheet, cell.row, cell.col),
                  where,
                )
                // Service-free fail-closed invariant: a cell overlapping a
                // parse-marked unresolved rule never presents CF properties as
                // cleanly available.
                const base = baseOf(sheet, cell.row, cell.col)
                const desc = scope.describe(cell.row, cell.col)
                const styled = scope.effectiveStyle(cell.row, cell.col, base as never)
                expect(desc?.text === undefined || desc.text === 'Conditional formatting unavailable', where).toBe(true)
                if (styled !== base) expect(desc?.text, where).toBe('Conditional formatting unavailable')
              }
              scope.dispose()
            }
            expect(net()).toBe(0)
          }
        } finally {
          lease.release()
        }
      }
      oracle.dispose()
      expect(oracleMemory.currentBytes).toBe(0)
    }
    // The sweep must actually exercise unresolved rules, across mechanisms.
    expect(unresolvedCount).toBeGreaterThan(0)
    expect(probedCells).toBeGreaterThan(0)
    const reasons = new Set([...seen.values()].map(v => v.reason))
    expect([...reasons].sort()).toEqual(['cf-unsupported-dxf', 'cf-unsupported-rule'])
    const types = new Set([...seen.values()].map(v => v.ruleType))
    for (const expected of ['expression', 'cellIs', 'containsText']) {
      expect(types.has(expected), `mechanism ${expected} covered`).toBe(true)
    }
  })
})

describe('unresolved rules withhold through the real paged worker route', () => {
  it('parse-marked unresolved cells are unavailable with base style on the worker read path', async () => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
    const { createCfMemory } = await import('../src/xlsx/cf/limits')
    const manifest = loadManifest()
    const stage1Cases = manifest.cases.filter(c => c.stage === 1 || !c.stage)
    let rules = 0
    let cells = 0
    let readyReads = 0
    let pendingReads = 0
    for (const c of stage1Cases) {
      const sheet = await sheetFromSaved(c.input, c.saved)
      const unresolved = (sheet.conditionalFormatting?.rules ?? []).filter(rule => {
        const definition = (rule as { definition?: { kind?: string } }).definition
        return definition?.kind === 'unresolved'
      })
      if (unresolved.length === 0) continue
      const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))
      const session = await openImportSession(
        { kind: 'blob', blob: new Blob([bytes.buffer as ArrayBuffer]) },
        { docId: `stage1-wire-${c.id}`, profile: 'D' },
      )
      try {
        await session.done
        const source = createPagedWorkbookSource(session)
        const sheetId = source.sheets[0]!.sheetId
        const store = source.store(sheetId)
        const memory = createCfMemory()
        for (const rule of unresolved) {
          rules++
          const ranges = (rule as unknown as { ranges: readonly { firstRow: number; lastRow: number; firstCol: number; lastCol: number }[] }).ranges
          const rects = ranges.map(r => ({ r0: r.firstRow, r1: Math.min(r.lastRow, r.firstRow + 4), c0: r.firstCol, c1: r.firstCol }))
          if (rects.length === 0) continue
          // Warm the pages first so the CF read evaluates actual states rather
          // than pending on unimported ranges; a residual pending is still a
          // correct fail-closed withhold (tracked for retry, never properties).
          for (const r of rects) {
            await store.readRows?.({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1 }, undefined, { signal: new AbortController().signal, stamp: store.stamp } as never).catch(() => undefined)
          }
          const { budget, net } = trackingBudget()
          const result = await store.readConditionalFormatting?.({
            rects,
            stamp: { owner: 'audit', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 },
            grantBytes: 2 * 1024 * 1024,
          }, memory, { signal: new AbortController().signal })
          if (result?.kind !== 'ready') {
            // Withheld without states: pending (import-tracked), unresolved or
            // error. None may expose CF properties as available.
            expect(result?.kind === 'pending' || result?.kind === 'unresolved' || result?.kind === 'error', `${c.id} withhold kind`).toBe(true)
            if (result?.kind === 'pending') expect(result.pendingSheetIds?.length ?? 0, `${c.id} pending tracked`).toBeGreaterThan(0)
            pendingReads++
            continue
          }
          readyReads++
          const { budget: wireBudget, net: wireNet } = trackingBudget()
          const scope = await prepareCfWireScope(
            async () => ({ result: result as CfReadResult, release: () => {} }),
            rects,
            { locale: 'en-US', budget: wireBudget },
          )
          try {
            expect(scope, `${c.id} wire scope`).toBeDefined()
            for (const r of rects) {
              for (let row = r.r0; row <= r.r1; row++) {
                cells++
                const where = `${c.id} wire ${row},${r.c0}`
                const desc = scope!.describe(row, r.c0)
                const styled = scope!.effectiveStyle(row, r.c0, baseOf(sheet, row, r.c0) as never)
                // Independent fail-closed contract (no product composition in
                // the oracle): a cell overlapping a parse-marked unresolved
                // rule must never present CF properties as cleanly available.
                const base = baseOf(sheet, row, r.c0)
                expect(desc?.text === undefined || desc.text === 'Conditional formatting unavailable', where).toBe(true)
                if (styled !== base) expect(desc?.text, where).toBe('Conditional formatting unavailable')
              }
            }
          } finally {
            scope?.dispose()
          }
          expect(wireNet()).toBe(0)
          expect(net()).toBe(0)
          void budget
        }
      } finally {
        await session.release()
      }
    }
    expect(rules).toBeGreaterThan(0)
    expect(cells).toBeGreaterThan(0)
    // Both withhold paths must be exercised: ready-with-invariant and
    // pending-tracked. (If all reads pend, page warming is broken.)
    expect(readyReads).toBeGreaterThan(0)
    expect(pendingReads).toBeGreaterThan(0)
  }, 120000)
})

describe('manifest mutations fail the audit', () => {
  const rootDir = REPO_ROOT

  it('dropped hash fails', () => {
    const manifest = loadManifest()
    delete (manifest.cases[0]!.expected as unknown as Record<string, unknown>).savedSha256
    const report = auditStage1Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
    expect(report.perCase.find(entry => entry.id === manifest.cases[0]!.id)?.code).toBe('hash-mismatch')
  })

  it('dropped capture fails', () => {
    const manifest = loadManifest()
    manifest.cases[1]!.captures = []
    const report = auditStage1Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
  })

  it('dropped case fails', () => {
    const manifest = loadManifest()
    manifest.cases = manifest.cases.filter(c => c.id !== 'stage1-dxf-layering-2')
    const report = auditStage1Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
    expect(report.missingCases).toContain('stage1-dxf-layering-2')
  })

  it('marking an unmeasured case supported fails', () => {
    const manifest = loadManifest()
    const target = manifest.cases[0]!
    target.expected.provenance = 'seed-hypothesis' as never
    const report = auditStage1Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
  })
})
