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
import { auditStage5Manifest, CF_STAGE5_PRODUCTION_READINGS, stage5Publication, type Stage5AuditReport } from '../scripts/cf-refs/audit'
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

describe('formula CF files parse identically on both paths', () => {
  it('the corpus formula-bearing file count equals the §4 inventory (0 is recorded, never a pass on its own)', () => {
    // A missing corpus/ is recorded by the count below; the oracle workbooks are the unconditional part of the gate.
    let withFormulas = 0
    for (const file of corpusFiles) {
      const bytes = new Uint8Array(readFileSync(join(CORPUS, file)))
      const text = new TextDecoder('latin1').decode(bytes)
      if (FORMULA_CF.test(text)) withFormulas++
    }
    expect(withFormulas).toBe(CORPUS_FORMULA_FILES)
    // The unconditional part: Excel's dialog-authored workbook and every Stage 5 workbook.
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
    const current = readFileSync(join(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8')
    const s5 = "export const VERIFIED_STAGE5_CAPABILITIES: ReadonlySet<string> = new Set([\n  'formula:relative-wrap',\n])\n"
    const sel = "export const VERIFIED_STAGE5_FORMULA_SELECTIONS: Readonly<Record<string, string>> = {\n  'formula:relative-wrap': 'wrap',\n}\n"
    const committed = current.replace(/export const VERIFIED_CAPABILITIES:/, `${s5}\n${sel}\nexport const VERIFIED_CAPABILITIES:`)
    const out = renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: committed })
    expect(out).toContain(s5)
    expect(out).toContain(sel)
    expect(out.indexOf(sel)).toBeGreaterThan(out.indexOf(s5))
    const bare = renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: current })
    expect(bare).not.toContain('VERIFIED_STAGE5_FORMULA_SELECTIONS')
    // An incomplete audit publishes nothing, even when asked (Lesson 4).
    if (!auditStage5Manifest(manifest(), { rootDir: REPO_ROOT }).stageComplete) {
      expect(renderVerifiedCapabilitiesSource(manifest(), REPO_ROOT, { committedSource: current, publishStage5: true })).toBe(bare)
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
