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
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join, resolve } from 'node:path'
import { auditStage1Manifest, type Stage1AuditReport } from '../scripts/cf-refs/audit'
import { CF_STAGE1_CAPABILITIES, type CfManifest } from '../scripts/cf-refs/cases'
import { prepareCfTextScope, prepareCfWireScope } from '../src/xlsx/cf/text'
import { createCfService } from '../src/xlsx/cf/service'
import { createCfMemory, CF_LIMITS } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import type { XlsxSheet } from '../src/xlsx/types'
import type { CfReadResult, CfCellStateWire } from '../src/xlsx/view/contracts'
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

describe('every supported Stage 1 semantic capability has verified oracle evidence', () => {
  it('the manifest audit passes with full catalog coverage and structural evidence', () => {
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
})

describe('every unresolved capability refuses atomically on both paths', () => {
  it('all product-marked unresolved rules withhold on resident and wire scopes', async () => {
    const manifest = loadManifest()
    const seen = new Map<string, { ruleType: string; reason: string }>()
    let unresolvedCount = 0
    let probedCells = 0
    const baseOf = (sheet: XlsxSheet, row: number, col: number) => {
      for (const band of sheet.rows) {
        const cell = band.cells.find(candidate => candidate.row === row && candidate.col === col)
        if (cell) return cell.style
      }
      return undefined
    }
    for (const c of manifest.cases) {
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
          // Record the composed outcome per cell plus the wire states for it.
          const cells: { row: number; col: number; kind: string; hasProperties: boolean; hasWrites: boolean }[] = []
          const wireStates: CfCellStateWire[] = []
          for (const r of rects) {
              for (let row = r.r0; row <= r.r1; row++) {
                const outcome = lease.value.at(row, r.c0)
                if (outcome.kind === 'unresolved' && outcome.reason === 'cf-outside-preparation') continue
              const hasProperties = outcome.kind === 'ready' && Object.keys(outcome.value.properties).length > 0
              const writes = outcome.kind === 'ready' ? outcome.value.unresolvedWrites : undefined
              const hasWrites = writes !== undefined && (writes as ReadonlySet<string>).size !== undefined
                ? (writes as ReadonlySet<string>).size > 0
                : Array.isArray(writes) ? writes.length > 0 : false
              cells.push({ row, col: r.c0, kind: outcome.kind, hasProperties, hasWrites })
              probedCells++
              if (outcome.kind === 'no-match') continue
              if (outcome.kind === 'unresolved') {
                wireStates.push({ row, col: r.c0, properties: {}, unresolvedWrites: undefined, unresolvedReason: outcome.reason })
              } else {
                wireStates.push({
                  row, col: r.c0,
                  properties: { ...outcome.value.properties },
                  unresolvedWrites: outcome.value.unresolvedWrites ? [...outcome.value.unresolvedWrites] : undefined,
                  unresolvedReason: undefined,
                })
              }
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
            if (wireStates.length > 0) expect(scope, `${c.id} resident scope`).toBeDefined()
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
              }
              scope.dispose()
            }
            expect(net()).toBe(0)
          }
          // Wire path must agree with the same oracle outcomes.
          {
            const { budget, net } = trackingBudget()
            const result: CfReadResult = {
              stamp: { owner: 'audit', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: gen, viewGen: 0, structureGen: 0, clockGen: 0 },
              sheetId: sheet.sheetId ?? sheet.name,
              kind: 'ready',
              states: wireStates,
              definitions: [],
              definitionsVersion: 1,
              diagnostics: [],
            }
            const scope = await prepareCfWireScope(
              async () => ({ result, release: () => {} }),
              rects,
              { locale: 'en-US', budget },
            )
            if (wireStates.length > 0) expect(scope, `${c.id} wire scope`).toBeDefined()
            if (scope) {
              for (const cell of cells) {
                const where = `${c.id} wire ${cell.row},${cell.col}`
                expectCell(
                  (row, col) => scope.describe(row, col) ?? undefined,
                  (row, col, base) => scope.effectiveStyle(row, col, base as never),
                  cell,
                  baseOf(sheet, cell.row, cell.col),
                  where,
                )
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
    expect([...reasons].sort()).toEqual(['cf-invalid-rule', 'cf-unsupported-dxf', 'cf-unsupported-rule'])
    const types = new Set([...seen.values()].map(v => v.ruleType))
    for (const expected of ['expression', 'cellIs', 'containsText', 'top10', 'aboveAverage']) {
      expect(types.has(expected), `mechanism ${expected} covered`).toBe(true)
    }
  })
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
    const target = manifest.cases.find(c => c.status === 'recipe-pending')!
    target.status = 'measured'
    const report = auditStage1Manifest(manifest, { rootDir })
    expect(report.ok).toBe(false)
  })
})
