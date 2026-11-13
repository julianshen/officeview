// @vitest-environment node
/**
 * Task 23: Stage 3 (data bars) validation, resource audit and PR evidence.
 *
 * Verifies that:
 * - every enabled bar capability has verified evidence, and production publishes exactly the audit's Stage 3 result
 *   (Ruling S3-WRITE) while the Stage 1/2 sets stay their own audits' results;
 * - every unresolved bar shape refuses atomically on the resident and the wire paths;
 * - corpus workbooks parse to identical CF rules on the resident and the streamed paths;
 * - the Stage 3 pressure scenarios hold their bounds with zero unowned live bytes.
 *
 * Manifest-format mutations live in tests/xlsx-cf-oracle.test.ts (Task 18) and are not duplicated here.
 */
import { describe, expect, it } from 'vitest'
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join, resolve } from 'node:path'
import {
  auditStage1Manifest,
  auditStage2Manifest,
  auditStage3Manifest,
  baselineNoCf,
  barRule,
  checkPressureSample,
  pressureBar1MRowSharedScan,
  pressureBarDistinctRanges,
  pressureBarUnitCap,
  pressureBarWholeColumnExtents,
  pressureCycles,
  pressureRefusedPercentileBarWithMinMax,
} from '../scripts/cf-refs/audit'
import { CF_STAGE3_CAPABILITIES, CF_STAGE3_NONDISCRIMINATING, CF_STAGE3_UNSUPPORTED, type CfManifest } from '../scripts/cf-refs/cases'
import * as verifiedExports from '../src/xlsx/cf/verifiedCapabilities'
import { VERIFIED_CAPABILITIES, VERIFIED_STAGE1_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES } from '../src/xlsx/cf/evidence'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { prepareCfTextScope, prepareCfWireScope } from '../src/xlsx/cf/text'
import { createCfService } from '../src/xlsx/cf/service'
import { createCfMemory, CF_LIMITS } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRulesVersion } from '../src/xlsx/view/pagedSource'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import type { XlsxSheet } from '../src/xlsx/types'
import type { CfReadResult } from '../src/xlsx/view/contracts'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const CORPUS = resolve(REPO_ROOT, 'corpus')
const manifest = (): CfManifest => JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
const UNAVAILABLE = 'Conditional formatting unavailable'

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

async function savedSheet(path: string): Promise<XlsxSheet> {
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, path)))))
  return (doc.sheets.find(s => s.name === 'Sheet1') ?? doc.sheets[0]!) as unknown as XlsxSheet
}

const colOf = (letters: string): number => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
/** The single rect of an `A1:B2` sqref. */
function sqrefRect(sqref: string): { firstRow: number; lastRow: number; firstCol: number; lastCol: number } {
  const [a, b = a] = sqref.split(':').map(ref => /^([A-Z]+)(\d+)$/.exec(ref)!)
  return { firstRow: Number(a![2]) - 1, lastRow: Number(b![2]) - 1, firstCol: colOf(a![1]!), lastCol: colOf(b![1]!) }
}

function baseOf(sheet: XlsxSheet, row: number, col: number): unknown {
  for (const band of sheet.rows) {
    const cell = band.cells.find(candidate => candidate.row === row && candidate.col === col)
    if (cell) return cell.style
  }
  return undefined
}

describe('every enabled bar capability has verified evidence', () => {
  const audit = auditStage3Manifest(manifest(), { rootDir: REPO_ROOT })
  const by = (d: string) => audit.perCapability.filter(p => p.disposition === d).map(p => p.capabilityId).sort()

  it('the Stage 3 audit is complete, with every bar capability accounted for', () => {
    expect(audit.ok).toBe(true)
    expect(audit.stageComplete).toBe(true)
    expect(audit.blockedCapabilities).toEqual([])
    expect(audit.perCapability.map(p => p.capabilityId).sort()).toEqual(CF_STAGE3_CAPABILITIES.map(c => c.capabilityId).sort())
    for (const p of audit.perCapability) {
      expect(['verified', 'unverified-nondiscriminating', 'refused-unsupported', 'supported-structural'], p.capabilityId).toContain(p.disposition)
    }
    expect(by('verified')).toHaveLength(51)
    expect(by('unverified-nondiscriminating')).toEqual(Object.keys(CF_STAGE3_NONDISCRIMINATING).sort())
    expect(by('refused-unsupported')).toEqual([...CF_STAGE3_CAPABILITIES.filter(c => c.capabilityId.startsWith('bar:unresolved-')).map(c => c.capabilityId), ...Object.keys(CF_STAGE3_UNSUPPORTED)].sort())
    expect(by('supported-structural')).toEqual([])
  })

  it('production publishes exactly the audited Stage 3 set and its exports (Ruling S3-WRITE)', () => {
    const published = Reflect.get(verifiedExports, 'VERIFIED_STAGE3_CAPABILITIES') as ReadonlySet<string> | undefined
    expect(published, 'VERIFIED_STAGE3_CAPABILITIES is published').toBeDefined()
    expect([...published!].sort()).toEqual(by('verified'))
    for (const cap of by('verified')) {
      expect(VERIFIED_CAPABILITIES.has(cap), cap).toBe(true)
      expect(cfRuntimeEvidence.allows(cap), cap).toBe(true)
    }
    for (const cap of [...by('unverified-nondiscriminating'), ...by('refused-unsupported')]) {
      expect(VERIFIED_CAPABILITIES.has(cap), `${cap} must not be production-verified`).toBe(false)
      expect(cfRuntimeEvidence.allows(cap), cap).toBe(false)
    }
    expect(Reflect.get(verifiedExports, 'VERIFIED_STAGE3_BAR_LAYOUT')).toEqual(audit.barLayout.layout)
    expect(Reflect.get(verifiedExports, 'VERIFIED_STAGE3_AUTO_BAR_COLORS')).toEqual(audit.autoBarColours)
    expect(Reflect.get(verifiedExports, 'VERIFIED_STAGE3_BAR_GRADIENT')).toEqual(audit.barGradient)
    expect(Reflect.get(verifiedExports, 'VERIFIED_STAGE3_BAR_GRADIENT_BORDER')).toEqual(audit.barGradientBorder)
    expect(Reflect.get(verifiedExports, 'VERIFIED_STAGE3_BAR_LENGTHS')).toEqual(audit.barLengths)
    expect(Reflect.get(verifiedExports, 'VERIFIED_STAGE3_BAR_AUTO_BOUNDS')).toEqual(audit.barAutoBounds)
  })

  it('the Stage 1/2 sets are their own audits and the combined set is exactly the three stages', () => {
    const verifiedOf = (r: { perCapability: readonly { capabilityId: string; disposition: string }[] }) => r.perCapability.filter(p => p.disposition === 'verified').map(p => p.capabilityId).sort()
    expect([...VERIFIED_STAGE1_CAPABILITIES].sort()).toEqual(verifiedOf(auditStage1Manifest(manifest(), { rootDir: REPO_ROOT })))
    expect([...VERIFIED_STAGE2_CAPABILITIES].sort()).toEqual(verifiedOf(auditStage2Manifest(manifest(), { rootDir: REPO_ROOT })))
    const s3 = Reflect.get(verifiedExports, 'VERIFIED_STAGE3_CAPABILITIES') as ReadonlySet<string>
    expect([...s3].every(id => id.startsWith('bar:'))).toBe(true)
    expect([...VERIFIED_CAPABILITIES].sort()).toEqual([...new Set([...VERIFIED_STAGE1_CAPABILITIES, ...VERIFIED_STAGE2_CAPABILITIES, ...s3])].sort())
  })
})

describe('every unresolved bar shape refuses atomically on both paths', () => {
  it('every parse-marked unresolved bar withholds on the resident scope: unavailable, base style, no bar', async () => {
    const stage3 = manifest().cases.filter(c => c.stage === 3)
    const reasons = new Set<string>()
    let targetsSeen = 0, probed = 0, sheetLevel = 0
    for (const c of stage3) {
      // Excel strips the unresolved shapes on resave, so their case is read as authored (the workbook a viewer gets).
      const sheet = await savedSheet(c.status === 'unsupported' ? c.input : c.saved)
      const cf = sheet.conditionalFormatting!
      // A sheet-level refusal (an orphan x14 or a duplicate GUID taints the whole sheet) covers every marker range.
      const targets: { label: string; ranges: readonly { firstRow: number; lastRow: number; firstCol: number; lastCol: number }[] }[] = cf.unresolvedSheet
        ? c.markers.map(m => ({ label: `${m.capabilityId} (sheet: ${cf.unresolvedSheet!.reason})`, ranges: [sqrefRect(m.sqref)] }))
        : cf.rules.filter(rule => {
          const d = rule.definition as { kind?: string; type?: string }
          return d.kind === 'unresolved' && d.type === 'dataBar'
        }).map(rule => ({ label: `pri${rule.priority} ${(rule.definition as { reason: string }).reason}`, ranges: rule.ranges }))
      if (targets.length === 0) continue
      if (cf.unresolvedSheet) { sheetLevel++; reasons.add(cf.unresolvedSheet.reason) }
      else for (const rule of cf.rules) if ((rule.definition as { kind?: string }).kind === 'unresolved') reasons.add((rule.definition as { reason: string }).reason)
      const oracleMemory = createCfMemory()
      const oracle = createCfService(cf, oracleMemory, { evidence: cfRuntimeEvidence })
      const stamp = { owner: 'cf-text', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(cf), viewGen: 0, structureGen: 0, clockGen: 0 }
      const source = residentSource(sheet, stamp, oracleMemory, 'engine')
      for (const target of targets) {
        targetsSeen++
        const rects = target.ranges.map(r => ({ r0: r.firstRow, r1: Math.min(r.lastRow, r.firstRow + 4), c0: r.firstCol, c1: r.firstCol }))
        const lease = await oracle.prepareCf(
          { stamp, sourceRects: rects.map(r => ({ firstRow: r.r0, lastRow: r.r1, firstCol: r.c0, lastCol: r.c1 })), signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes },
          source,
          oracleMemory,
        )
        try {
          const { budget, net } = trackingBudget()
          const scope = await prepareCfTextScope(sheet, { rects, signal: new AbortController().signal, locale: 'en-US', budget })
          // A sheet-level refusal leaves no rule to scope: the text scope is absent and every cell keeps its base style.
          if (cf.unresolvedSheet) {
            expect(scope, `${c.id} sheet-level refusal builds no scope`).toBeUndefined()
            for (const r of rects) for (let row = r.r0; row <= r.r1; row++) {
              expect(lease.value.at(row, r.c0), `${c.id} ${target.label} ${row},${r.c0}`).toEqual({ kind: 'unresolved', reason: cf.unresolvedSheet.reason })
              probed++
            }
            expect(net()).toBe(0)
            continue
          }
          expect(scope, `${c.id} resident scope`).toBeDefined()
          for (const r of rects) for (let row = r.r0; row <= r.r1; row++) {
            const outcome = lease.value.at(row, r.c0)
            if (outcome.kind === 'unresolved' && outcome.reason === 'cf-outside-preparation') continue
            const where = `${c.id} ${target.label} ${row},${r.c0}`
            // The oracle never yields a bar for an unresolved rule's cell.
            expect(outcome.kind === 'ready' && outcome.value.visual?.dataBar !== undefined, where).toBe(false)
            expect(scope!.barAt?.(row, r.c0), where).toBeUndefined()
            if (outcome.kind === 'unresolved') {
              expect(scope!.describe(row, r.c0)?.text, where).toBe(UNAVAILABLE)
              expect(scope!.effectiveStyle(row, r.c0, baseOf(sheet, row, r.c0) as never), where).toBe(baseOf(sheet, row, r.c0))
            }
            probed++
          }
          scope!.dispose()
          expect(net()).toBe(0)
        } finally {
          lease.release()
        }
      }
      oracle.dispose()
      expect(oracleMemory.currentBytes).toBe(0)
    }
    expect(sheetLevel).toBe(1)
    expect(targetsSeen).toBeGreaterThan(15)
    expect(probed).toBeGreaterThan(60)
    expect(reasons.size).toBeGreaterThan(1)
  })

  it('the unresolved-shapes workbook withholds every bar through the real paged worker route', async () => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
    const c = manifest().cases.find(x => x.id === 'stage3-bar-unresolved-shapes')!
    // The authored input: Excel's resave strips these shapes.
    const sheet = await savedSheet(c.input)
    // An orphan x14 and a duplicate GUID taint the whole sheet: every marker range must withhold.
    expect(sheet.conditionalFormatting?.unresolvedSheet?.reason).toBe('cf-invalid-x14-link')
    const unresolved = c.markers.map(m => ({ priority: m.capabilityId, ranges: [sqrefRect(m.sqref)] }))
    expect(unresolved.length).toBeGreaterThan(5)
    const session = await openImportSession(
      { kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(join(REPO_ROOT, c.input))).buffer as ArrayBuffer]) },
      { docId: 'stage3-wire-unresolved', profile: 'D' },
    )
    let probed = 0
    try {
      await session.done
      const source = createPagedWorkbookSource(session)
      const sheetId = source.sheets[0]!.sheetId
      const store = source.store(sheetId)
      const memory = createCfMemory()
      for (const rule of unresolved) {
        const rects = rule.ranges.map(r => ({ r0: r.firstRow, r1: Math.min(r.lastRow, r.firstRow + 4), c0: r.firstCol, c1: r.firstCol }))
        const result = await store.readConditionalFormatting?.({
          rects,
          stamp: { owner: 'audit', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 },
          grantBytes: 2 * 1024 * 1024,
        }, memory, { signal: new AbortController().signal })
        // The worker refuses the whole read with the sheet's reason and sends no state, so no bar.
        expect(result?.kind, `${rule.priority}`).toBe('unresolved')
        expect(result?.reason, `${rule.priority}`).toBe('cf-invalid-x14-link')
        expect(result!.states.filter(s => s.bar !== undefined), `${rule.priority}`).toEqual([])
        // The wire scope refuses the unresolved read outright (the engine keeps base formatting), charging nothing.
        const { budget, net } = trackingBudget()
        await expect(prepareCfWireScope(async () => ({ result: result as CfReadResult, release: () => {} }), rects, { locale: 'en-US', budget }))
          .rejects.toThrow('CF wire preparation is unresolved')
        for (const r of rects) probed += r.r1 - r.r0 + 1
        expect(net()).toBe(0)
      }
      memory.close()
    } finally {
      await session.release()
    }
    expect(probed).toBeGreaterThan(20)
  }, 120000)
})

describe('published bars agree on the resident and the paged worker paths', () => {
  it('every measured Stage 3 case carries the same bar descriptors through the real paged route as the resident service', async () => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
    let compared = 0, bars = 0
    for (const c of manifest().cases.filter(x => x.stage === 3 && x.status === 'measured')) {
      const sheet = await savedSheet(c.saved)
      const cf = sheet.conditionalFormatting!
      const memory = createCfMemory()
      const merges = (sheet.mergeRanges ?? []).map(m => ({ firstRow: m.minRow, lastRow: m.maxRow, firstCol: m.minCol, lastCol: m.maxCol }))
      const service = createCfService(cf, memory, { evidence: cfRuntimeEvidence, merges })
      const stamp = { owner: 'resident', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(cf), viewGen: 0, structureGen: 0, clockGen: 0 }
      const rect = sqrefRect(c.range)
      // The resident outcome over the case range, in paint batches of at most 256 cells (the cap per call).
      const resident = new Map<string, unknown>()
      for (let r0 = rect.firstRow; r0 <= rect.lastRow; r0 += 8) {
        const part = { ...rect, firstRow: r0, lastRow: Math.min(rect.lastRow, r0 + 7) }
        const lease = await service.prepareCf({ stamp, sourceRects: [part], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(sheet, stamp, memory, 'engine'))
        for (let row = part.firstRow; row <= part.lastRow; row++) for (let col = part.firstCol; col <= part.lastCol; col++) {
          const o = lease.value.at(row, col)
          const bar = o.kind === 'ready' ? o.value.visual?.dataBar : undefined
          if (bar) resident.set(`${row},${col}`, { start: bar.startFraction, end: bar.endFraction, axis: bar.axisFraction, fill: bar.fillColor, neg: bar.isNegative })
        }
        lease.release()
      }
      service.dispose()
      const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(readFileSync(join(REPO_ROOT, c.saved))).buffer as ArrayBuffer]) }, { docId: `stage3-agree-${c.id}`, profile: 'D' })
      const wire = new Map<string, unknown>()
      try {
        await session.done
        const source = createPagedWorkbookSource(session)
        const sheetId = source.sheets.find(x => x.name === c.sheet)?.sheetId ?? source.sheets[0]!.sheetId
        const store = source.store(sheetId)
        const wireMemory = createCfMemory()
        for (let r0 = rect.firstRow; r0 <= rect.lastRow; r0 += 8) {
          const rects = [{ r0, r1: Math.min(rect.lastRow, r0 + 7), c0: rect.firstCol, c1: rect.lastCol }]
          const result = await store.readConditionalFormatting?.({ rects, stamp: { owner: 'audit', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: 2 * 1024 * 1024 }, wireMemory, { signal: new AbortController().signal })
          expect(result?.kind, `${c.id} rows ${r0}+`).toBe('ready')
          for (const st of result!.states) if (st.bar) wire.set(`${st.row},${st.col}`, { start: st.bar.startFraction, end: st.bar.endFraction, axis: st.bar.axisFraction, fill: st.bar.fillColor, neg: st.bar.isNegative })
        }
        await store.release()
      } finally {
        await session.release()
      }
      expect([...wire.entries()].sort(), c.id).toEqual([...resident.entries()].sort())
      compared++
      bars += resident.size
    }
    expect(compared).toBe(14)
    expect(bars).toBeGreaterThan(400)
  }, 300000)
})

// The Stage 3 corpus inventory (docs/validation/xlsx-cf.md §4): no corpus workbook carries a data bar.
const CORPUS_DATABAR_FILES = 0
const corpusFiles = existsSync(CORPUS) ? readdirSync(CORPUS).filter(f => f.endsWith('.xlsx') && !f.startsWith('~$')).sort() : []

describe.skipIf(!existsSync(CORPUS) || corpusFiles.length === 0)('corpus dataBar files parse identically on both paths', () => {
  it('the dataBar file count equals the §4 inventory', async () => {
    const { readZipPartsSync } = await import('../scripts/cf-refs/verify')
    let withBars = 0
    for (const file of corpusFiles) {
      const parts = readZipPartsSync(new Uint8Array(readFileSync(join(CORPUS, file))))
      expect(parts, file).not.toBeNull()
      const sheets = [...parts!.entries()].filter(([name]) => /^xl\/worksheets\/[^/]+\.xml$/.test(name))
      if (sheets.some(([, xml]) => /<(x14:)?dataBar\b/.test(new TextDecoder().decode(xml)))) withBars++
    }
    expect(withBars).toBe(CORPUS_DATABAR_FILES)
  })

  it.each(corpusFiles)('%s: resident and streamed CfSheetRules deep-equal, nothing throws', async file => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const bytes = new Uint8Array(readFileSync(join(CORPUS, file)))
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: `stage3-corpus-${file}`, profile: 'D' })
    try {
      await session.done.catch(() => undefined)
      for (const paged of session.workbook.sheets) {
        if (paged.kind !== 'worksheet') continue
        const resident = doc.sheets.find(s => s.name === paged.name)
        expect(paged.conditionalFormatting, `${file} ${paged.name}`).toEqual(resident?.conditionalFormatting)
      }
    } finally {
      await session.release()
    }
  }, 120000)
})

// No corpus workbook carries CF at all (§4), so the corpus check above compares absent rules. The Stage 3 oracle
// workbooks and Excel's own dialog-authored bar are the bar writer shapes this repository holds: they run the same check.
const BAR_WORKBOOKS = [...new Set(manifest().cases.filter(c => c.stage === 3).flatMap(c => [c.saved, c.input])), 'tests/fixtures/xlsx-cf/excel-authored-databar.xlsx']
describe('Stage 3 oracle workbooks parse identically on both paths', () => {
  it.each(BAR_WORKBOOKS)('%s', async path => {
    const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
    const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, path)))
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: `stage3-both-${path}`, profile: 'D' })
    try {
      await session.done
      let withCf = 0
      for (const paged of session.workbook.sheets) {
        if (paged.kind !== 'worksheet') continue
        const resident = doc.sheets.find(s => s.name === paged.name)
        expect(paged.conditionalFormatting, `${path} ${paged.name}`).toEqual(resident?.conditionalFormatting)
        if (resident?.conditionalFormatting) withCf++
      }
      expect(withCf, path).toBeGreaterThan(0)
    } finally {
      await session.release()
    }
  }, 120000)
})

describe('Stage 3 pressure scenarios and resource bounds', () => {
  it('a 1M-row DECLARED A:A scope with 100 bars and 100 scales shares one scan and paints the first bar', async () => {
    const sample = await pressureBar1MRowSharedScan()
    checkPressureSample(sample, { expectOutcome: 'admitted' })
    expect(sample.statisticsReads).toBe(100)
    // The text scope's byte estimate refuses all 100 cells under 200 rules before reading; 40 admit and paint.
    expect(sample.estimatorRefusedAll).toBe(true)
    expect(sample.cellsWithResults).toBe(40)
    // A per-rule scan would read each of the 100 values 200 times.
    expect(sample.valueReads).toBeLessThanOrEqual(1000)
    expect(sample.netBytes).toBe(0)
  })

  it('a refused percentile bar withholds while a min/max bar still paints', async () => {
    const sample = await pressureRefusedPercentileBarWithMinMax()
    checkPressureSample(sample, { expectOutcome: 'admitted' })
    expect(sample.cellsWithResults).toBe(10)
    expect(sample.netBytes).toBe(0)
  })

  it('a whole-column bar does not extend extents (isolation)', async () => {
    const sample = await pressureBarWholeColumnExtents()
    checkPressureSample(sample, { expectOutcome: 'admitted' })
    expect(sample.cellsWithResults).toBe(10)
  })

  it('distinct-range bar pressure refuses past the rect cap, and a bar scan refuses past the 1,000,000-unit cap', async () => {
    const distinct = await pressureBarDistinctRanges()
    checkPressureSample(distinct, { expectOutcome: 'refused' })
    const cap = await pressureBarUnitCap()
    expect(CF_LIMITS.workUnits).toBe(1_000_000)
    expect(cap.admittedReads).toBe(cap.admittedCells)
    expect(cap.admittedUnits).toBeLessThanOrEqual(CF_LIMITS.workUnits)
    expect(cap.refusedReason).toBe('cf-resource-limit')
    // Refused part-way: the scan stopped at the cap, before reading every cell.
    expect(cap.refusedReads).toBeLessThan(cap.refusedCells)
    expect(cap.refusedUnits).toBeLessThanOrEqual(CF_LIMITS.workUnits)
    expect(cap.netBytes).toBe(0)
  }, 120000)

  it('20 open/read/cancel/close cycles over a bar range return to zero', async () => {
    const cycles = await pressureCycles(20, { fixture: 'tests/fixtures/xlsx-cf/stage3-bar-lengths-default.xlsx', rect: { r0: 1, r1: 11, c0: 1, c1: 1 }, requireBars: true })
    expect(cycles.cycles).toBe(20)
    expect(cycles.barStates).toBeGreaterThanOrEqual(20)
    expect(cycles.finalNetBytes).toBe(0)
    expect(cycles.maxPeakBytes).toBeGreaterThan(0)
  }, 120000)

  it('zero bar work or reads on a no-CF baseline', async () => {
    const baseline = await baselineNoCf(barRule(0, { firstRow: 0, lastRow: 99, firstCol: 0, lastCol: 0 }) as unknown as Record<string, unknown>)
    expect(baseline).toEqual({ paintReserves: 0, sourceReads: 0, scopeBuilt: false })
  })
})
