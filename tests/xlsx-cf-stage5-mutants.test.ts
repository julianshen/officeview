// @vitest-environment node
/**
 * Task 39 fix round 2 (codex re-review, rulings R1 and R2): in-memory mutants that the Stage 5 validation must catch.
 * - R1: the CF compiler refuses the supported `$A2>0` (cf-resource-limit). A compile refusal justifies nothing unless an
 *   independently recorded expectation names it, so stage5-composition B2's lost fill must make the audit rows and the
 *   conformance verdict disagree.
 * - R2: justifications are slot- and rule-specific. stage5-composition L3's icon justification (icon:composition-dxf-fill)
 *   excuses only the icon slot: the fill rule refused beside it, or a green fill beside the withheld icon (Excel: red),
 *   must disagree in the audit, the conformance verdict and cf-compare.
 */
import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const mutant = vi.hoisted(() => ({ refuse: undefined as string | undefined }))
vi.mock('../src/xlsx/cf/formula', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/cf/formula')>()
  const refused = { refused: 'cf-resource-limit' as const, reason: 'mutant: refused at compile' }
  return {
    ...actual,
    compileCfFormula: (text: string, ...rest: Parameters<typeof actual.compileCfFormula> extends [string, ...infer R] ? R : never) =>
      (mutant.refuse !== undefined && text.trim() === mutant.refuse ? refused : actual.compileCfFormula(text, ...rest)),
    compileCfFormulaWith: (text: string, ...rest: Parameters<typeof actual.compileCfFormulaWith> extends [string, ...infer R] ? R : never) =>
      (mutant.refuse !== undefined && text.trim() === mutant.refuse ? refused : actual.compileCfFormulaWith(text, ...rest)),
  }
})

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const manifest = () => JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as import('../scripts/cf-refs/cases').CfManifest
async function loadCase(id: string) {
  const { OfficePackage } = await import('../src/core/zip')
  const { parseXlsx } = await import('../src/xlsx/parse')
  const c = manifest().cases.find(x => x.id === id)!
  const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
  return { c, doc, sheet: doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]! }
}

describe('R1: a compile refusal is no justification unless independently recorded', () => {
  it('the unmutated composition B2 agrees; with $A2>0 refused at compile, the audit row and the conformance verdict disagree', async () => {
    const { stage5SlotVerdict } = await import('./helpers/cf-stage5-verdict')
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const clean = await loadCase('stage5-composition')
    expect(await stage5SlotVerdict(clean.c, clean.doc, clean.sheet, 'B2', REPO_ROOT)).toBeUndefined()
    mutant.refuse = '$A2>0'
    try {
      const m = await loadCase('stage5-composition')
      // Production's own parse refused the rule (the mutant reached it), and the helper found no recorded expectation.
      expect(m.sheet.conditionalFormatting!.rules.some(r => (r.definition as { kind: string }).kind === 'unresolved')).toBe(true)
      expect(await stage5SlotVerdict(m.c, m.doc, m.sheet, 'B2', REPO_ROOT)).toMatch(/without a recorded justification/)
      const rows = await stage5CaseRecords({ ...manifest(), cases: [m.c] }, REPO_ROOT)
      const b2 = rows.find(r => r.ref === 'B2')!
      expect(b2.agrees, JSON.stringify(b2)).toBe(false)
    } finally { mutant.refuse = undefined }
  }, 120000)
  it('every recorded compile refusal still refuses exactly as recorded', async () => {
    const { STAGE5_EXPECTED_COMPILE_REFUSALS } = await import('../scripts/cf-refs/stage5-withhold')
    const { compileCfFormula, isCfFormulaRefusal } = await import('../src/xlsx/cf/formula')
    const { createCfMemory } = await import('../src/xlsx/cf/limits')
    expect(STAGE5_EXPECTED_COMPILE_REFUSALS.length).toBeGreaterThan(0)
    for (const e of STAGE5_EXPECTED_COMPILE_REFUSALS) {
      const c = manifest().cases.find(x => x.id === e.caseId)!
      expect(c.markers.some(m => m.formulas?.includes(e.formula)), e.formula).toBe(true)
      const out = compileCfFormula(e.formula, { row: 1, col: 1 }, createCfMemory())
      expect(isCfFormulaRefusal(out) && out.refused, e.formula).toBe(e.refused)
    }
  })
})

describe('R2: justifications are slot- and rule-specific', () => {
  it('composition L3: the literal fill rule refused beside the justified icon slot disagrees (audit and conformance verdict)', async () => {
    const { stage5SlotVerdict } = await import('./helpers/cf-stage5-verdict')
    const { stage5Justifications } = await import('../scripts/cf-refs/stage5-withhold')
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const m = await loadCase('stage5-composition')
    expect(await stage5SlotVerdict(m.c, m.doc, m.sheet, 'L3', REPO_ROOT), 'unmutated').toBeUndefined()
    // The only justification at L3 excuses the icon rule's slots.
    for (const j of stage5Justifications(m.c, 'L3')) expect(j.slots, j.reason).not.toContain('fill')
    const lit = m.c.markers.find(x => x.ruleType === 'cellIs' && x.sqref === 'L2:L7')!
    const rule = m.sheet.conditionalFormatting!.rules.find(r => r.priority === lit.priority)!
    ;(rule as { definition: unknown }).definition = { kind: 'unresolved', type: 'cellIs', reason: 'cf-resource-limit' }
    expect(await stage5SlotVerdict(m.c, m.doc, m.sheet, 'L3', REPO_ROOT)).toMatch(/fill unresolved without a recorded justification/)
    void stage5CaseRecords
  }, 120000)
  it('composition L3: a green fill beside an unresolved icon disagrees with Excel\'s red (agreement, and cf-compare)', async () => {
    const { stage5CellAgreement, stage5Justifications } = await import('../scripts/cf-refs/stage5-withhold')
    const { verifyCaseCells } = await import('../scripts/cf-compare')
    const c = manifest().cases.find(x => x.id === 'stage5-composition')!
    const RED = 'FFFFC7CE', GREEN = 'FFC6EFCE'
    expect(stage5CellAgreement({ fill: GREEN, strike: false, unresolved: ['iconSet'], visual: 'withheld' }, { fill: RED, strike: false }, stage5Justifications(c, 'L3')).agrees).toBe(false)
    // cf-compare: production is Excel's everywhere except L3, which shows green with the icon slot unresolved.
    const { stage5ObservedFills, stage5ObservedStrikes } = await import('../scripts/cf-compare')
    const fills = stage5ObservedFills(c, REPO_ROOT)!, strikes = stage5ObservedStrikes(c, REPO_ROOT)
    const refOf = (row: number, col: number) => `${String.fromCharCode(65 + col)}${row + 1}`
    const msg = verifyCaseCells(c, REPO_ROOT, (row, col) => {
      const ref = refOf(row, col)
      if (ref === 'L3') return { applied: true, unresolved: true, unresolvedSlots: ['iconSet'], style: { fillColor: GREEN } }
      const f = fills.get(ref)
      const painted = f !== undefined && f !== 'FFFFFFFF'
      return { applied: painted || strikes.get(ref) === true, unresolved: false, unresolvedSlots: [], ...(painted || strikes.get(ref) ? { style: { ...(painted ? { fillColor: f } : {}), ...(strikes.get(ref) ? { strike: true } : {}) } } : {}) }
    })
    expect(msg).toMatch(/L3.*fill FFC6EFCE vs Excel FFFFC7CE/)
  })
})

describe('fix round 3 (S1, S2): justification scope follows the compositor\'s real effect, property by property', () => {
  const observed = async (c: import('../scripts/cf-refs/cases').CfOracleCase) => {
    const { stage5ObservedFills, stage5ObservedStrikes } = await import('../scripts/cf-compare')
    return { fills: stage5ObservedFills(c, REPO_ROOT)!, strikes: stage5ObservedStrikes(c, REPO_ROOT) }
  }
  /** cf-compare over a case whose production is Excel's everywhere except at `ref`. */
  async function compareWith(c: import('../scripts/cf-refs/cases').CfOracleCase, ref: string, at: import('../scripts/cf-compare').ActualCellOutcome): Promise<string | undefined> {
    const { verifyCaseCells } = await import('../scripts/cf-compare')
    const { fills, strikes } = await observed(c)
    return verifyCaseCells(c, REPO_ROOT, (row, col) => {
      const r = `${String.fromCharCode(65 + col)}${row + 1}`
      if (r === ref) return at
      const f = fills.get(r)
      const painted = f !== undefined && f !== 'FFFFFFFF'
      return { applied: painted || strikes.get(r) === true, unresolved: false, unresolvedSlots: [], ...(painted || strikes.get(r) ? { style: { ...(painted ? { fillColor: f } : {}), ...(strikes.get(r) ? { strike: true } : {}) } } : {}) }
    })
  }

  it('S1: composition I3 with its stopping expression refused (no fill, ALL unresolved) disagrees: audit, conformance verdict, cf-compare', async () => {
    const { stage5SlotVerdict } = await import('./helpers/cf-stage5-verdict')
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const { stage5CellAgreement, stage5Justifications } = await import('../scripts/cf-refs/stage5-withhold')
    const m = await loadCase('stage5-composition')
    expect(await stage5SlotVerdict(m.c, m.doc, m.sheet, 'I3', REPO_ROOT), 'unmutated').toBeUndefined()
    // The missing formula:composition-stop-icon excuses the lower icon's slot only, never the stopping rule's own fill.
    for (const j of stage5Justifications(m.c, 'I3')) { expect(j.slots, j.reason).not.toContain('fill'); expect(j.slots, j.reason).not.toContain('ALL') }
    expect(stage5CellAgreement({ fill: 'none', strike: false, unresolved: ['ALL'], visual: 'withheld' }, { fill: 'FFFFC7CE', strike: false }, stage5Justifications(m.c, 'I3')).agrees).toBe(false)
    const expr = m.c.markers.find(x => x.ruleType === 'expression' && x.sqref === 'I2:I7')!
    const rule = m.sheet.conditionalFormatting!.rules.find(r => r.priority === expr.priority)!
    ;(rule as { definition: unknown }).definition = { kind: 'unresolved', type: 'expression', reason: 'cf-resource-limit' }
    expect(await stage5SlotVerdict(m.c, m.doc, m.sheet, 'I3', REPO_ROOT)).toMatch(/without a recorded justification/)
    expect(await compareWith(m.c, 'I3', { applied: false, unresolved: true, unresolvedSlots: ['ALL'] })).toMatch(/I3/)
    // The audit rows on production's own parse with the expression refused (the compiler mocked to refuse $A2>0).
    mutant.refuse = '$A2>0'
    try {
      const rows = await stage5CaseRecords({ ...manifest(), cases: [m.c] }, REPO_ROOT)
      const i3 = rows.find(r => r.ref === 'I3')!
      expect(i3.production.unresolved, JSON.stringify(i3.production)).toContain('ALL')
      expect(i3.agrees, JSON.stringify(i3)).toBe(false)
    } finally { mutant.refuse = undefined }
  }, 120000)

  it('S2: composition Q3 with only font.bold unresolved and strikethrough resolved false (Excel: struck) disagrees everywhere', async () => {
    const { stage5CellAgreement, stage5Justifications } = await import('../scripts/cf-refs/stage5-withhold')
    const c = manifest().cases.find(x => x.id === 'stage5-composition')!
    const { strikes } = await observed(c)
    expect(strikes.get('Q3'), 'Excel struck Q3').toBe(true)
    // An unrelated bold refusal never suppresses the strikethrough comparison (property-level masks).
    expect(stage5CellAgreement({ fill: 'none', strike: false, unresolved: ['font.bold'], visual: 'none' }, { fill: 'none', strike: true }, stage5Justifications(c, 'Q3')).agrees).toBe(false)
    expect(await compareWith(c, 'Q3', { applied: false, unresolved: true, unresolvedSlots: ['font.bold'] })).toMatch(/Q3/)
    // The conformance verdict path: the same production token through its helper.
    const { stage5VerdictOf } = await import('./helpers/cf-stage5-verdict')
    expect(stage5VerdictOf(c, 'Q3', { fill: 'none', strike: false, unresolved: ['font.bold'], visual: 'none' }, REPO_ROOT)).toMatch(/strikethrough/)
  })
})
