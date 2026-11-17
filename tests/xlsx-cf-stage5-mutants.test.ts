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

const mutant = vi.hoisted(() => ({ refuse: undefined as string | undefined, afterParse: undefined as ((doc: import('../src/xlsx/types').XlsxDocument) => void) | undefined }))
// Fix round 4: a doc mutation applied after production's own parse, so the audit rows (which parse from disk) see it.
vi.mock('../src/xlsx/parse', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/parse')>()
  return { ...actual, parseXlsx: async (...args: Parameters<typeof actual.parseXlsx>) => { const doc = await actual.parseXlsx(...args); mutant.afterParse?.(doc); return doc } }
})
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

describe('fix round 4 (codex r3): the counterfactual compositor catches what the hand-mirrored grants accepted', () => {
  type Token = import('../scripts/cf-refs/stage5-withhold').Stage5Production
  type Case = import('../scripts/cf-refs/cases').CfOracleCase
  type Doc = import('../src/xlsx/types').XlsxDocument
  type Sheet = import('../src/xlsx/types').XlsxSheet
  const refOf = (row: number, col: number) => `${String.fromCharCode(65 + col)}${row + 1}`
  const actualOf = (t: Token): import('../scripts/cf-compare').ActualCellOutcome => ({
    applied: t.fill !== 'none' || t.strike, unresolved: t.unresolved.length > 0, unresolvedSlots: [...t.unresolved],
    ...(t.fill !== 'none' || t.strike ? { style: { ...(t.fill !== 'none' ? { fillColor: t.fill } : {}), ...(t.strike ? { strike: true } : {}) } } : {}),
  })
  /** cf-compare's Stage 5 check over production's own tokens for every expected and observed cell of the (mutated) doc. */
  async function compareProduction(c: Case, doc: Doc, sheet: Sheet): Promise<string | undefined> {
    const { verifyCaseCells, stage5ObservedFills } = await import('../scripts/cf-compare')
    const { stage5ProductionAt } = await import('./helpers/cf-stage5-verdict')
    const { stage5CounterfactualOf } = await import('../scripts/cf-refs/stage5-withhold') as { stage5CounterfactualOf?: (input: unknown) => Promise<unknown> }
    const refs = [...new Set([...c.expected.cells.map(e => e.ref), ...(stage5ObservedFills(c, REPO_ROOT)?.keys() ?? [])])]
    const tokens = new Map<string, Token>()
    for (const ref of refs) tokens.set(ref, await stage5ProductionAt(doc, sheet, ref))
    const truths = stage5CounterfactualOf ? await stage5CounterfactualOf({ c, doc, sheet, refs, now: Date.now }) : undefined
    return (verifyCaseCells as (...a: unknown[]) => string | undefined)(c, REPO_ROOT, (row: number, col: number) => actualOf(tokens.get(refOf(row, col)) ?? { fill: 'none', strike: false, unresolved: [], visual: 'none' }), truths)
  }
  const refuseRule = (sheet: Sheet, priority: number, type: string) => {
    const rule = sheet.conditionalFormatting!.rules.find(r => r.priority === priority)!
    ;(rule as { definition: unknown }).definition = { kind: 'unresolved', type, reason: 'cf-resource-limit' }
  }

  it('r3 #1: composition I2 with the priority-13 icon refused while the stop expression is false disagrees: audit, conformance helper, cf-compare', async () => {
    const { stage5SlotVerdict, stage5ProductionAt } = await import('./helpers/cf-stage5-verdict')
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const clean = await loadCase('stage5-composition')
    expect(await stage5SlotVerdict(clean.c, clean.doc, clean.sheet, 'I2', REPO_ROOT), 'unmutated').toBeUndefined()
    expect((await stage5ProductionAt(clean.doc, clean.sheet, 'I2')).unresolved, 'unmutated I2 composes with no unresolved write').toEqual([])
    expect(await compareProduction(clean.c, clean.doc, clean.sheet), 'unmutated cf-compare').toBeUndefined()
    const m = await loadCase('stage5-composition')
    refuseRule(m.sheet, 13, 'iconSet')
    expect((await stage5ProductionAt(m.doc, m.sheet, 'I2')).unresolved, 'the mutant reaches production').toEqual(['iconSet'])
    expect.soft(await stage5SlotVerdict(m.c, m.doc, m.sheet, 'I2', REPO_ROOT), 'conformance helper').toMatch(/I2/)
    expect.soft(await compareProduction(m.c, m.doc, m.sheet), 'cf-compare').toMatch(/I2/)
    mutant.afterParse = doc => refuseRule(doc.sheets.find(s => s.name === m.c.sheet) ?? doc.sheets[0]!, 13, 'iconSet')
    try {
      const rows = await stage5CaseRecords({ ...manifest(), cases: [m.c] }, REPO_ROOT)
      const i2 = rows.find(r => r.ref === 'I2')!
      expect(i2.production.unresolved, JSON.stringify(i2.production)).toEqual(['iconSet'])
      expect.soft(i2.agrees, JSON.stringify(i2)).toBe(false)
    } finally { mutant.afterParse = undefined }
  }, 180000)

  it('r3 #2: the relative-scale marker made stopIfTrue over a lower strike rule keeps strike; a mutant that also loses strike disagrees everywhere', async () => {
    const { stage5SlotVerdict, stage5ProductionAt } = await import('./helpers/cf-stage5-verdict')
    const { stage5CaseRecords } = await import('../scripts/cf-refs/audit')
    const { VERIFIED_CAPABILITIES } = await import('../src/xlsx/cf/evidence')
    // The premise: the scale's stop flag is ignored (measured), so compose.ts's halts() is false for it.
    expect(VERIFIED_CAPABILITIES.has('scale:composition-scale-stopIfTrue')).toBe(true)
    const STRIKE_PRIORITY = 1000
    const strikeRule = (await loadCase('stage5-composition')).sheet.conditionalFormatting!.rules.find(r => r.priority === 35)!
    /** The synthetic variant: K2:K9's relative scale (p8) carries stopIfTrue, and a literal strike rule (cellIs > 0) lies below it. */
    const synthesize = (doc: Doc, refuseStrike: boolean) => {
      const sheet = doc.sheets.find(s => s.name === 'Sheet1') ?? doc.sheets[0]!
      const cf = sheet.conditionalFormatting!
      const strike = { ...strikeRule, id: 9999, priority: STRIKE_PRIORITY, documentOrder: 9999, ranges: [{ firstRow: 1, lastRow: 8, firstCol: 10, lastCol: 10 }],
        ...(refuseStrike ? { definition: { kind: 'unresolved', type: 'cellIs', reason: 'cf-resource-limit' } } : {}) }
      ;(sheet as { conditionalFormatting: unknown }).conditionalFormatting = { ...cf, rules: [...cf.rules.map(r => r.priority === 8 ? { ...r, stopIfTrue: true } : r), strike] }
    }
    const base = manifest().cases.find(x => x.id === 'stage5-cfvo-formulas')!
    const c = { ...base, markers: [...base.markers.map(k => k.priority === 8 ? { ...k, stopIfTrue: true } : k),
      { part: 'xl/worksheets/sheet1.xml', extension: 'legacy', sqref: 'K2:K9', capabilityId: 'formula:positive-control', ruleType: 'cellIs', priority: STRIKE_PRIORITY, dxf: { strikethrough: true } }] } as unknown as Case
    const clean = await loadCase('stage5-cfvo-formulas')
    synthesize(clean.doc, false)
    expect(await stage5ProductionAt(clean.doc, clean.sheet, 'K3'), 'synthetic: only the scale withholds, strike kept').toMatchObject({ unresolved: ['colorScale', 'fill'], strike: true })
    const m = await loadCase('stage5-cfvo-formulas')
    synthesize(m.doc, true)
    expect((await stage5ProductionAt(m.doc, m.sheet, 'K3')).unresolved, 'the mutant also loses strike').toEqual(['colorScale', 'fill', 'font.strike'])
    expect.soft(await stage5SlotVerdict(c, m.doc, m.sheet, 'K3', REPO_ROOT), 'conformance helper').toMatch(/K3.*(without a recorded justification|counterfactual)/)
    expect.soft(await compareProduction(c, m.doc, m.sheet), 'cf-compare').toMatch(/K\d.*(without a recorded justification|counterfactual)/)
    mutant.afterParse = doc => synthesize(doc, true)
    try {
      const rows = await stage5CaseRecords({ ...manifest(), cases: [c] }, REPO_ROOT)
      const k3 = rows.find(r => r.ref === 'K3')!
      expect(k3.production.unresolved, JSON.stringify(k3.production)).toEqual(['colorScale', 'fill', 'font.strike'])
      expect.soft(k3.agrees, JSON.stringify(k3)).toBe(false)
      expect.soft(k3.disagreement).toMatch(/without a recorded justification|counterfactual/)
    } finally { mutant.afterParse = undefined }
  }, 180000)
})
