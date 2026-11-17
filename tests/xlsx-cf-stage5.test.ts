// @vitest-environment node
/**
 * Task 39: Stage 5 (formula conditional formatting) validation, publication gate and corpus gate.
 *
 * Verifies that:
 * - every enabled formula capability has verified evidence: production holds a Stage 5 id only when the dual gate held
 *   (`auditStage5Manifest` reports `stageComplete` and `stage5Publication` published it), and every Excel outcome recorded
 *   on a tested construct is a registered refusal;
 * - every unresolved formula shape refuses atomically on the resident and the paged (worker import) paths;
 * - corpus and oracle formula workbooks parse to identical CF rules on both paths (the file count is the §4 inventory,
 *   never a vacuous pass).
 *
 * Covers the `formula:*` rows and the formula rows of `scale:`/`bar:`/`icon:` only.
 */
import { describe, expect, it } from 'vitest'
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join, resolve } from 'node:path'
import { auditStage5Manifest, CF_STAGE5_PRODUCTION_READINGS, renderFormulaSelectionsBlock, stage5Publication, type Stage5AuditReport } from '../scripts/cf-refs/audit'
import { CF_STAGE5_CAPABILITIES, CF_STAGE5_NONDISCRIMINATING, CF_STAGE5_UNSUPPORTED, type CfManifest } from '../scripts/cf-refs/cases'
import * as verifiedExports from '../src/xlsx/cf/verifiedCapabilities'
import { VERIFIED_CAPABILITIES } from '../src/xlsx/cf/evidence'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const CORPUS = resolve(REPO_ROOT, 'corpus')
const manifest = (): CfManifest => JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest

/** The formula rows this task covers: `formula:*` and the formula rows of the visual families. */
const FORMULA_FAMILY_ROWS = new Set(['scale:thresholds-formula', 'bar:bounds-formula', 'icon:thresholds-formula', 'scale:dialect-standalone'])
const isFormulaRow = (id: string): boolean => id.startsWith('formula:') || FORMULA_FAMILY_ROWS.has(id)
const publishedStage5 = (): ReadonlySet<string> | undefined => Reflect.get(verifiedExports, 'VERIFIED_STAGE5_CAPABILITIES') as ReadonlySet<string> | undefined

/**
 * The Excel outcomes recorded on the construct a case tests (Ruling S5-REPAIR, S5-J, OD-16): each must be a registered
 * refusal, so no later capture can verify it by accident.
 */
const RECORDED_EXCEL_REFUSALS = [
  'formula:active-present-scale',
  'formula:active-present-bar',
  'formula:active-present-icon',
  'formula:predicate-timeperiod-divergent',
  'formula:cross-sheet-missing',
  'formula:cfvo-nonnumeric',
  // Re-deferred after the composition rebuild: the icon-only twin K2 is unmeasurable (icon drawn at cell top + 4 px).
  'formula:composition-stop-icon',
  'formula:composition-indirect-ahead-icon',
]

describe('every enabled formula capability has verified evidence', () => {
  const audit = auditStage5Manifest(manifest(), { rootDir: REPO_ROOT })
  const gate = stage5Publication(audit)
  const rows = audit.perCapability.filter(p => isFormulaRow(p.capabilityId))
  const by = (d: string): string[] => rows.filter(p => p.disposition === d).map(p => p.capabilityId).sort()

  it('the Stage 5 audit accounts for every formula row', () => {
    expect(audit.ok).toBe(true)
    expect(rows.map(p => p.capabilityId).sort()).toEqual(CF_STAGE5_CAPABILITIES.map(c => c.capabilityId).filter(isFormulaRow).sort())
    expect(rows.length).toBeGreaterThan(70)
    for (const p of rows) expect(['verified', 'unverified-nondiscriminating', 'refused-unsupported', 'supported-structural'], p.capabilityId).toContain(p.disposition)
    expect(by('unverified-nondiscriminating')).toEqual(expect.arrayContaining(Object.keys(CF_STAGE5_NONDISCRIMINATING)))
    expect(by('refused-unsupported')).toEqual(expect.arrayContaining(Object.keys(CF_STAGE5_UNSUPPORTED)))
    // The audit's blocked list is exactly the supported-structural rows (nothing unaccounted).
    expect([...audit.blockedCapabilities].sort()).toEqual(audit.perCapability.filter(p => p.disposition === 'supported-structural').map(p => p.capabilityId).sort())
  })

  it('every Excel outcome recorded on a tested construct is a registered refusal (S5-REPAIR, S5-J)', () => {
    for (const id of RECORDED_EXCEL_REFUSALS) {
      expect(CF_STAGE5_UNSUPPORTED[id], `${id} registered`).toBeDefined()
      expect(rows.find(p => p.capabilityId === id)?.disposition, id).toBe('refused-unsupported')
    }
  })

  it('the publication gate: verified only, matched reading = production reading, a partition of the Stage 5 ids', () => {
    const verified = new Set(by('verified'))
    expect([...gate.published, ...Object.keys(gate.withheld)].sort()).toEqual(audit.perCapability.map(p => p.capabilityId).sort())
    for (const id of gate.published) {
      expect(verified.has(id) || !isFormulaRow(id), id).toBe(true)
      const matched = audit.matchedReadings[id]!
      const prod = CF_STAGE5_PRODUCTION_READINGS[id]
      expect(prod, `${id}: production names a reading`).toBeDefined()
      if (typeof prod === 'string') expect(prod, id).toBe(matched)
      else expect(prod, id).toContain(matched)
    }
    // Non-vacuous: the gate publishes and withholds for each reason today.
    expect(gate.published.length).toBeGreaterThan(20)
    const reasons = Object.values(gate.withheld)
    for (const prefix of ['reading-unimplemented: ', 'reading-mismatch: ', 'refused-unsupported', 'unverified-nondiscriminating']) {
      expect(reasons.some(r => r.startsWith(prefix)), prefix).toBe(true)
    }
  })

  it('the gate withholds a verified id whose matched reading production does not implement (mutations)', () => {
    const id = gate.published.find(x => typeof CF_STAGE5_PRODUCTION_READINGS[x] === 'string')!
    expect(id).toBeDefined()
    const mutate = (reading: string): Stage5AuditReport => ({ ...audit, matchedReadings: { ...audit.matchedReadings, [id]: reading } })
    expect(stage5Publication(mutate('invented')).withheld[id]).toMatch(/^reading-mismatch: /)
    const extra: Stage5AuditReport = { ...audit, perCapability: [...audit.perCapability, { capabilityId: 'formula:invented', caseId: 'x', disposition: 'verified' }], matchedReadings: { ...audit.matchedReadings, 'formula:invented': 'r' } }
    expect(stage5Publication(extra).withheld['formula:invented']).toMatch(/^reading-unimplemented: /)
  })

  it('dual gate: production holds a Stage 5 id only when the audit is complete and the gate published it', () => {
    const s5 = publishedStage5()
    if (!audit.stageComplete) {
      // Lesson 4: an incomplete audit publishes nothing, whatever the gate would select.
      expect(s5, 'VERIFIED_STAGE5_CAPABILITIES must not exist while the Stage 5 audit is incomplete').toBeUndefined()
      for (const p of audit.perCapability) {
        expect(VERIFIED_CAPABILITIES.has(p.capabilityId), p.capabilityId).toBe(false)
        expect(cfRuntimeEvidence.allows(p.capabilityId), p.capabilityId).toBe(false)
      }
      // Non-vacuous: the gate would publish ids, and still none reaches production.
      expect(gate.published.length).toBeGreaterThan(0)
      return
    }
    expect(s5, 'VERIFIED_STAGE5_CAPABILITIES is published').toBeDefined()
    expect([...s5!].sort()).toEqual(gate.published)
    for (const id of gate.published) expect(cfRuntimeEvidence.allows(id), id).toBe(true)
    for (const id of Object.keys(gate.withheld)) expect(cfRuntimeEvidence.allows(id), id).toBe(false)
  })
})

/** The authored and the Stage 5 oracle workbooks, as the viewer receives them (the saved file, or the input if Excel never resaved it). */
const STAGE5_WORKBOOKS = (() => {
  const m = manifest()
  const paths = m.cases.filter(c => c.stage === 5).flatMap(c => [c.saved, c.input])
  return [...new Set(['tests/fixtures/xlsx-cf/excel-authored-formulas.xlsx', ...paths])].filter(p => existsSync(join(REPO_ROOT, p)))
})()

async function bothPaths(path: string, bytes: Uint8Array, id: string): Promise<{ resident: Awaited<ReturnType<typeof parseXlsx>>; paged: Array<{ name: string; cf: unknown }> }> {
  const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
  const resident = await parseXlsx(await OfficePackage.load(bytes))
  const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.slice().buffer as ArrayBuffer]) }, { docId: id, profile: 'D' })
  try {
    await session.done
    return { resident, paged: session.workbook.sheets.filter(s => s.kind === 'worksheet').map(s => ({ name: s.name, cf: s.conditionalFormatting })) }
  } finally {
    await session.release()
  }
  void path
}

describe('every unresolved formula shape refuses atomically on both paths', () => {
  const cases = manifest().cases.filter(c => c.stage === 5 && c.status === 'unsupported')
  it('nine recorded unresolved shapes', () => {
    expect(cases.map(c => c.id).sort()).toEqual([
      'array-constant', 'cellis-text', 'duplicate-formula', 'external', 'indirect', 'name', 'spill', 'table', 'union',
    ].map(s => `stage5-unresolved-shapes-${s}`))
  })
  it.each(cases.map(c => c.id))('%s: every marked rule is unresolved on the resident and the paged path, and the two agree', async id => {
    const c = manifest().cases.find(x => x.id === id)!
    // Excel repaired, prompted or stripped these shapes, so the authored input is the workbook a viewer gets.
    const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, c.input)))
    const { resident, paged } = await bothPaths(c.input, bytes, `stage5-unresolved-${id}`)
    const sheet = resident.sheets.find(s => s.name === c.sheet) ?? resident.sheets[0]!
    const cf = sheet.conditionalFormatting
    expect(cf, `${id} has conditional formatting`).toBeDefined()
    const unresolvedMarkers = c.markers.filter(m => m.capabilityId.startsWith('formula:unresolved-'))
    expect(unresolvedMarkers.length, `${id} marks its shape`).toBeGreaterThan(0)
    for (const m of unresolvedMarkers) {
      const hit = cf!.unresolvedSheet !== undefined || cf!.rules.some(r => r.priority === m.priority && (r.definition as { kind?: string }).kind === 'unresolved')
      expect(hit, `${id} ${m.capabilityId} p${m.priority} is unresolved`).toBe(true)
      // Atomic: no rule with that priority resolved into a paintable definition.
      expect(cf!.rules.filter(r => r.priority === m.priority && (r.definition as { kind?: string }).kind !== 'unresolved'), `${id} p${m.priority}`).toEqual([])
    }
    for (const p of paged) expect(p.cf, `${id} ${p.name}`).toEqual(resident.sheets.find(s => s.name === p.name)?.conditionalFormatting)
  }, 120000)
})

// The Stage 5 corpus inventory (docs/validation/xlsx-cf.md §4): no corpus workbook carries a formula-bearing rule.
const CORPUS_FORMULA_FILES = 0
const corpusFiles = existsSync(CORPUS) ? readdirSync(CORPUS).filter(f => f.endsWith('.xlsx') && !f.startsWith('~$')).sort() : []
const FORMULA_CF = /<(?:x14:)?cfRule\b[^>]*>(?:(?!<\/(?:x14:)?cfRule>)[\s\S])*?(?:<formula>|<xm:f>|cfvo type="formula")/
/** Whether a workbook's worksheet parts (inflated) hold a formula-bearing CF rule. */
async function hasFormulaCf(bytes: Uint8Array): Promise<boolean> {
  const { readZipPartsSync } = await import('../scripts/cf-refs/verify')
  const parts = readZipPartsSync(bytes)
  expect(parts).not.toBeNull()
  return [...parts!.entries()].some(([name, xml]) => /^xl\/worksheets\/[^/]+\.xml$/.test(name) && FORMULA_CF.test(new TextDecoder().decode(xml)))
}

/** Every Stage 1-5 capability plus the Stage 5 selections: the evidence a published Stage 5 would give the parser. */
const ALL_STAGE5 = (): string[] => [...VERIFIED_CAPABILITIES, ...CF_STAGE5_CAPABILITIES.map(c => c.capabilityId)]
const SELECTIONS = { 'formula:relative-wrap': 'wrap', 'formula:cellis-operand-blank': 'blank-zero' }
async function withStage5Evidence<T>(fn: () => Promise<T>): Promise<T> {
  const { withExactCfEvidence, injectCfFormulaSelections } = await import('../src/xlsx/cf/evidence')
  injectCfFormulaSelections(SELECTIONS)
  try { return await withExactCfEvidence(ALL_STAGE5(), fn) } finally { injectCfFormulaSelections(null) }
}
const resolvedFormulaRules = (cf: unknown): number => ((cf as { rules?: Array<{ definition: { kind?: string; formula?: unknown; operands?: unknown[]; dependencySheets?: unknown } }> } | undefined)?.rules ?? [])
  .filter(r => r.definition.kind !== 'unresolved' && (r.definition.formula !== undefined || (r.definition.operands ?? []).some(o => (o as { formula?: unknown }).formula !== undefined))).length

describe('formula CF files parse identically on both paths', () => {
  it('the corpus formula-bearing file count equals the §4 inventory (0 is recorded, never a pass on its own)', async () => {
    // Positive control: the scanner inflates the worksheet parts, so it must flag Excel's dialog-authored workbook.
    expect(await hasFormulaCf(new Uint8Array(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/excel-authored-formulas.xlsx'))))).toBe(true)
    // A missing corpus/ is recorded by the count below; the oracle workbooks are the unconditional part of the gate.
    let withFormulas = 0
    for (const file of corpusFiles) if (await hasFormulaCf(new Uint8Array(readFileSync(join(CORPUS, file))))) withFormulas++
    expect(withFormulas).toBe(CORPUS_FORMULA_FILES)
    expect(STAGE5_WORKBOOKS).toContain('tests/fixtures/xlsx-cf/excel-authored-formulas.xlsx')
    expect(STAGE5_WORKBOOKS.length).toBeGreaterThan(30)
  })

  it.each(corpusFiles)('corpus %s: resident and streamed CfSheetRules deep-equal', async file => {
    const bytes = new Uint8Array(readFileSync(join(CORPUS, file)))
    const { resident, paged } = await bothPaths(file, bytes, `stage5-corpus-${file}`)
    for (const p of paged) expect(p.cf, `${file} ${p.name}`).toEqual(resident.sheets.find(s => s.name === p.name)?.conditionalFormatting)
  }, 120000)

  it.each(STAGE5_WORKBOOKS)('%s', async path => {
    const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, path)))
    const { resident, paged } = await bothPaths(path, bytes, `stage5-both-${path}`)
    expect(paged.length).toBeGreaterThan(0)
    for (const p of paged) expect(p.cf, `${path} ${p.name}`).toEqual(resident.sheets.find(s => s.name === p.name)?.conditionalFormatting)
  }, 120000)

  // Without Stage 5 evidence every formula rule takes the pre-Stage-5 path on both sides, so parity above compares
  // unresolved with unresolved. Under the evidence a publication would give, the formulas compile: parity must still hold.
  const resolvedTotal: number[] = []
  it.each(STAGE5_WORKBOOKS)('under Stage 5 evidence: %s', async path => {
    const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, path)))
    const { resident, paged } = await withStage5Evidence(() => bothPaths(path, bytes, `stage5-ev-${path}`))
    for (const p of paged) {
      const r = resident.sheets.find(s => s.name === p.name)?.conditionalFormatting
      expect(p.cf, `${path} ${p.name}`).toEqual(r)
      resolvedTotal.push(resolvedFormulaRules(r))
    }
  }, 120000)
  it('under Stage 5 evidence formulas resolve (non-vacuous), and the x14 cross-sheet cellIs resolves on both paths', async () => {
    expect(resolvedTotal.reduce((a, b) => a + b, 0)).toBeGreaterThan(50)
    const path = 'tests/fixtures/xlsx-cf/inputs/stage5-cross-sheet.xlsx'
    const { resident, paged } = await withStage5Evidence(() => bothPaths(path, new Uint8Array(readFileSync(join(REPO_ROOT, path))), 'stage5-ev-cross-operand'))
    for (const cf of [resident.sheets.find(s => s.name === 'Sheet1')?.conditionalFormatting, paged.find(p => p.name === 'Sheet1')?.cf]) {
      const rule = (cf as { rules: Array<{ priority: number; definition: { kind: string; operands?: Array<{ formula?: unknown }> } }> }).rules.find(r => r.priority === 3)!
      expect(rule.definition.kind).toBe('cellIs')
      expect(rule.definition.operands?.[0]?.formula).toBeDefined()
    }
  }, 120000)
})

describe('the Stage 5 formula selections travel with the published set (Task 35 parked item)', () => {
  it('renders VERIFIED_STAGE5_FORMULA_SELECTIONS sorted, in the shape evidence.ts reads', async () => {
    const { renderFormulaSelectionsBlock } = await import('../scripts/cf-refs/audit')
    const block = renderFormulaSelectionsBlock({ 'formula:relative-wrap': 'wrap', 'formula:cellis-operand-blank': 'blank-zero' })
    expect(block).toBe("export const VERIFIED_STAGE5_FORMULA_SELECTIONS: Readonly<Record<string, string>> = {\n  'formula:cellis-operand-blank': 'blank-zero',\n  'formula:relative-wrap': 'wrap',\n}\n")
    expect(renderFormulaSelectionsBlock({})).toBe('export const VERIFIED_STAGE5_FORMULA_SELECTIONS: Readonly<Record<string, string>> = {\n\n}\n')
  })

  it('a --write without --publish-stage5 carries the committed Stage 5 set and its selections verbatim, and omits both when absent', async () => {
    const { renderVerifiedCapabilitiesSource } = await import('../scripts/cf-refs/audit')
    // The file without any Stage 5 export (Task 39 published one), so the carry is tested from a known absence.
    const current = readFileSync(join(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
      .replace(/export const VERIFIED_STAGE5_CAPABILITIES: [^\n]*\n[\s\S]*?\n\]\)\n\n/, '')
      .replace(/export const VERIFIED_STAGE5_FORMULA_SELECTIONS: [^\n]*\n[\s\S]*?\n\}\n\n/, '')
    const s5 = "export const VERIFIED_STAGE5_CAPABILITIES: ReadonlySet<string> = new Set([\n  'formula:relative-wrap',\n])\n"
    const sel = "export const VERIFIED_STAGE5_FORMULA_SELECTIONS: Readonly<Record<string, string>> = {\n  'formula:relative-wrap': 'wrap',\n}\n"
    const committed = current.replace(/export const VERIFIED_CAPABILITIES:/, `${s5}\n${sel}\nexport const VERIFIED_CAPABILITIES:`)
    const out = renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: committed })
    expect(out).toContain(s5)
    expect(out).toContain(sel)
    expect(out.indexOf(sel)).toBeGreaterThan(out.indexOf(s5))
    const bare = renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: current })
    expect(bare).not.toContain('VERIFIED_STAGE5_FORMULA_SELECTIONS')
    expect(bare).not.toContain('VERIFIED_STAGE5_CAPABILITIES')
    // An incomplete audit publishes nothing, even when asked (Lesson 4).
    const audit = auditStage5Manifest(manifest(), { rootDir: REPO_ROOT })
    if (!audit.stageComplete) {
      expect(renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: current, publishStage5: true })).toBe(bare)
    } else {
      // A complete audit publishes the gate's set and its selections block beside it.
      const pub = renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: current, publishStage5: true })
      expect(pub).toContain(renderFormulaSelectionsBlock(stage5Publication(audit).selections))
    }
  }, 120000)
})

describe('Stage 5 pressure scenarios and resource bounds (zero unowned live bytes)', () => {
  it('100 formula rules over a declared 1M-row A:A share one snapshot of the target cells', async () => {
    const { pressureFormulaSharedSnapshot } = await import('../scripts/cf-refs/pressure-stage5')
    const s = await pressureFormulaSharedSnapshot()
    expect(s).toMatchObject({ rules: 100, attempts: 1, snapshotEntries: 40, ready: 4000, netBytes: 0 })
  })
  it('SUM($A:$A) and COUNTIF($A:$A,A2)>1 hoist over 10,000 populated rows; ~1M populated refuse at the scan budget; SUM(A1:A200) refuses before any call', async () => {
    const { pressureFormulaHoists, CHECKPOINT_UNITS } = await import('../scripts/cf-refs/pressure-stage5')
    const s = await pressureFormulaHoists()
    expect(s).toMatchObject({ visibleCells: 5000, columnScans: 2, clampedScans: 2, refused: 0, millionCalls: 0, relativeCalls: 0, netBytes: 0 })
    expect(s.maxCallUnits).toBeLessThanOrEqual(CHECKPOINT_UNITS)
    expect(s.calls).toBeGreaterThanOrEqual(5000)
    expect(s.millionOutcomes).toEqual(['unresolved:cf-resource-limit', 'unresolved:cf-resource-limit'])
    expect(s.relativeOutcomes).toEqual(['unresolved:cf-resource-limit'])
  }, 120000)
  it('the dependency snapshot admits exactly its cap and refuses cap + 1', async () => {
    const { pressureFormulaSnapshotCap } = await import('../scripts/cf-refs/pressure-stage5')
    const s = await pressureFormulaSnapshotCap()
    expect(s).toMatchObject({ cap: 262_144, atCap: 'ready', atCapEntries: 262_144, overCap: 'unresolved:cf-resource-limit', netBytes: 0 })
  }, 120000)
  it('16 referenced sheets evaluate; 17 refuse at compile; a union of 18 refuses only the rule that crosses 16', async () => {
    const { pressureFormulaReferencedSheets } = await import('../scripts/cf-refs/pressure-stage5')
    const s = await pressureFormulaReferencedSheets()
    expect(s.at16).toBe('ready')
    expect(s.compile17).toMatch(/^refused:/)
    expect(s.union18).toEqual(['ready', 'unresolved:cf-resource-limit'])
    expect(s.netBytes).toBe(0)
  })
  it('a refused footprint beside a ready rule leaves the ready rule evaluating', async () => {
    const { pressureFormulaRefusedBesideReady } = await import('../scripts/cf-refs/pressure-stage5')
    expect(await pressureFormulaRefusedBesideReady()).toMatchObject({ readyCells: 10, refusedCells: 10, refusedReasons: ['unresolved:cf-resource-limit'], netBytes: 0 })
  })
  it('20 open/read/cancel/close cycles, including a close during a cross-sheet pending read, hold nothing between cycles', async () => {
    const { pressureFormulaCrossSheetCycles } = await import('../scripts/cf-refs/pressure-stage5')
    expect(await pressureFormulaCrossSheetCycles(20)).toMatchObject({ cycles: 20, ready: 20, pending: 20, cancelled: 20, maxNetBetween: 0, netBytes: 0 })
  })
})

describe('reuse-only capabilities keep their recorded reading whatever the owning case verifies (Task 39 gate defect)', () => {
  it('the S5-STUB successors are matched `recorded` and publish with a verified stage5-composition', () => {
    const audit = auditStage5Manifest(manifest(), { rootDir: REPO_ROOT })
    const stubs = audit.perCapability.filter(p => p.capabilityId.startsWith('formula:composition-stub-successor-'))
    expect(stubs).toHaveLength(4)
    for (const p of stubs) {
      expect(p.disposition, p.capabilityId).toBe('verified')
      expect(audit.matchedReadings[p.capabilityId], p.capabilityId).toBe('recorded')
    }
    const gate = stage5Publication(audit)
    for (const p of stubs) expect(gate.published, p.capabilityId).toContain(p.capabilityId)
  }, 120000)
})

describe('per-case and per-variant Stage 5 audit rows (Task 39; fix round 1, F3)', () => {
  it('a resolved paint is compared even when another slot is unresolved; a visual is compared where Excel measured one', async () => {
    const { stage5CellAgreement } = await import('../scripts/cf-refs/audit')
    const RED = 'FFFFC7CE', GREEN = 'FFC6EFCE'
    // Codex's synthetic: production paints a green fill with an unresolved font.strike; Excel shows red. Wrong paint.
    expect(stage5CellAgreement({ fill: GREEN, strike: false, unresolved: ['font.strike'], visual: 'none' }, { fill: RED, strike: false }, ['x']).agrees).toBe(false)
    // The same unresolved strike with Excel's own fill is fine only with a justification for the unresolved slot.
    expect(stage5CellAgreement({ fill: RED, strike: false, unresolved: ['font.strike'], visual: 'none' }, { fill: RED, strike: false }, ['x']).agrees).toBe(true)
    expect(stage5CellAgreement({ fill: RED, strike: false, unresolved: ['font.strike'], visual: 'none' }, { fill: RED, strike: false }, []).agrees).toBe(false)
    // A wholly withheld fill needs a justification.
    expect(stage5CellAgreement({ fill: 'none', strike: false, unresolved: ['fill'], visual: 'none' }, { fill: RED, strike: false }, []).agrees).toBe(false)
    expect(stage5CellAgreement({ fill: 'none', strike: false, unresolved: ['fill'], visual: 'none' }, { fill: RED, strike: false }, ['x']).agrees).toBe(true)
    // Visual evidence: a bar where Excel measured none (or the reverse) disagrees; unmeasured is not compared.
    expect(stage5CellAgreement({ fill: 'none', strike: false, unresolved: [], visual: 'bar' }, { fill: 'none', strike: false, visual: 'none' }, []).agrees).toBe(false)
    expect(stage5CellAgreement({ fill: 'none', strike: false, unresolved: [], visual: 'bar' }, { fill: 'none', strike: false, visual: 'bar' }, []).agrees).toBe(true)
    expect(stage5CellAgreement({ fill: 'none', strike: false, unresolved: [], visual: 'bar' }, { fill: 'none', strike: false }, []).agrees).toBe(true)
  })

  it('rows for every Stage 5 case: measured cases compared cell by cell, the others say why they are not', async () => {
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const m = manifest()
    const rows = await stage5CaseRecords(m, REPO_ROOT)
    const all = m.cases.filter(c => c.stage === 5)
    expect(all).toHaveLength(31)
    expect(new Set(rows.map(r => r.caseId))).toEqual(new Set(all.map(c => c.id)))
    const measured = all.filter(c => c.status === 'measured')
    expect(rows.filter(r => r.compared).length).toBe(measured.reduce((n, c) => n + c.expected.cells.length, 0))
    for (const r of rows.filter(x => !x.compared)) expect(r.notCompared, `${r.caseId}:${r.ref}`).toMatch(/\S/)
    expect(rows.filter(r => r.compared && !r.agrees).map(r => `${r.caseId}:${r.ref} excel=${JSON.stringify(r.excel)} production=${JSON.stringify(r.production)}`)).toEqual([])
    const formula = rows.filter(r => r.compared && r.evidenceKind === 'display-format')
    expect(formula.filter(r => r.production.fill === r.excel?.fill && r.production.fill !== 'none' && r.production.unresolved.length === 0).length).toBeGreaterThan(100)
    expect(formula.filter(r => r.production.unresolved.length > 0).length).toBeGreaterThan(0)
    expect(rows.filter(r => r.excel?.visual !== undefined).length, 'bar or icon evidence compared').toBeGreaterThan(0)
    for (const r of rows) expect(r.mode).toBe('resident')
  }, 300000)
})

describe('publication is validated first, then written atomically (Task 39 fix round 1, F4)', () => {
  it('a failure after rendering leaves the target byte-unchanged; a passing run replaces it in one rename', async () => {
    const { mkdtempSync, writeFileSync: write, readFileSync: read, readdirSync: list } = await import('node:fs')
    const { tmpdir } = await import('node:os')
    const { spawnSync } = await import('node:child_process')
    const dir = mkdtempSync(join(tmpdir(), 'cf-publish-'))
    const target = join(dir, 'verifiedCapabilities.ts')
    write(target, 'SENTINEL\n')
    const run = (env: Record<string, string>) => spawnSync('npx', ['tsx', 'scripts/cf-refs/audit.ts', '--stage', '1', '--write', '--target', target], { cwd: REPO_ROOT, encoding: 'utf8', env: { ...process.env, ...env } })
    const failed = run({ CF_AUDIT_INJECT_FAILURE: 'after-render' })
    expect(failed.status, failed.stdout + failed.stderr).not.toBe(0)
    expect(read(target, 'utf8')).toBe('SENTINEL\n')
    expect(list(dir)).toEqual(['verifiedCapabilities.ts'])
    const ok = run({})
    expect(ok.status, ok.stdout + ok.stderr).toBe(0)
    expect(read(target, 'utf8')).toBe(read(join(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8'))
    expect(list(dir)).toEqual(['verifiedCapabilities.ts'])
  }, 300000)
  it('writeFileAtomic replaces the file through a temporary sibling and leaves nothing behind', async () => {
    const { writeFileAtomic } = await import('../scripts/cf-refs/audit')
    const { mkdtempSync, readFileSync: read, readdirSync: list, writeFileSync: write } = await import('node:fs')
    const { tmpdir } = await import('node:os')
    const dir = mkdtempSync(join(tmpdir(), 'cf-atomic-'))
    const target = join(dir, 'x.ts')
    write(target, 'old')
    writeFileAtomic(target, 'new')
    expect(read(target, 'utf8')).toBe('new')
    expect(list(dir)).toEqual(['x.ts'])
  })
})

describe('the no-formula baseline keeps literal CF and does no formula work (Task 39 fix round 1, F5)', () => {
  it('literal cellIs CF: 0 formula compiles, 0 snapshot preparations, 0 clock reads; one expression rule moves all three', async () => {
    const { pressureNoFormulaBaseline } = await import('../scripts/cf-refs/pressure-stage5')
    const { literal, control } = await pressureNoFormulaBaseline()
    expect(literal.painted, 'the literal CF still paints (the baseline keeps CF)').toBeGreaterThan(0)
    expect({ compiles: literal.compiles, snapshots: literal.snapshots, clockReads: literal.clockReads }).toEqual({ compiles: 0, snapshots: 0, clockReads: 0 })
    expect(control.compiles, 'positive control: compiles').toBeGreaterThan(0)
    expect(control.snapshots, 'positive control: snapshots').toBeGreaterThan(0)
    expect(control.clockReads, 'positive control: clock reads').toBeGreaterThan(0)
  }, 120000)
})

describe('enumerated withholding justifications are backed by the workbook (Task 39 fix round 1, F1)', () => {
  it('every enumerated refusal names a cell whose rule reads a range the S5-HOIST policy refuses', async () => {
    const { STAGE5_ENUMERATED_REFUSALS, stage5WithholdJustification } = await import('../scripts/cf-refs/stage5-withhold')
    const keys = Object.keys(STAGE5_ENUMERATED_REFUSALS)
    expect(keys.length).toBe(18)
    const c = manifest().cases.find(x => x.id === 'stage5-formula-functions')!
    for (const k of keys) {
      const ref = k.split(':')[1]!
      expect(c.markers.some(m => m.formulas?.includes('COUNTIF($A:$A,A2)>1') && m.sqref.split(':')[0]!.startsWith('H')), k).toBe(true)
      expect(stage5WithholdJustification(c, ref).join(), k).toMatch(/S5-HOIST policy/)
    }
    // The recorded condition: column A of the saved workbook holds a boolean.
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
    const a = doc.sheets[0]!.rows.flatMap(r => r.cells).filter(x => x.col === 0)
    expect(a.some(x => typeof (x as { value?: unknown }).value === 'boolean')).toBe(true)
    // A published configuration with no such record is never justified (composition B2: every capability published).
    const comp = manifest().cases.find(x => x.id === 'stage5-composition')!
    expect(stage5WithholdJustification(comp, 'B2')).toEqual([])
  })
})

describe('cf-compare interactive runs production inputs on Stage 5 cases (Task 39 fix round 1, F2)', () => {
  it.each(['stage5-formula-authored', 'stage5-cross-sheet'])('%s: the formula cells are compared on production\'s paint, never skipped as unresolved', async id => {
    const { compareInteractiveViewport } = await import('../scripts/cf-compare')
    const c = manifest().cases.find(x => x.id === id)!
    const r = await compareInteractiveViewport(c, REPO_ROOT)
    expect(r.outcome, r.reason).toBe('pass')
    expect(r.stage5, 'Stage 5 cell accounting').toBeDefined()
    // Every formula cell Excel formats is painted by production, or withheld with a recorded justification.
    expect(r.stage5!.painted, 'formula cells painted').toBeGreaterThan(0)
    expect(r.stage5!.painted + r.stage5!.justifiedWithheld).toBe(r.stage5!.formulaMatched)
  }, 120000)
  it('an unresolved formula cell with a verified expectation and no recorded justification fails', async () => {
    const { verifyCellExpectations } = await import('../scripts/cf-compare')
    const c = manifest().cases.find(x => x.id === 'stage5-composition')!
    // Production withholds everything: the first verified formula cell with no recorded justification fails.
    const msg = verifyCellExpectations(c, () => ({ applied: false, unresolved: true }))
    expect(msg).toMatch(/unresolved without a recorded justification/)
  })
})
