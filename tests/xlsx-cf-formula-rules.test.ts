// @vitest-environment node
/**
 * Stage 5, Task 35: expression rules, formula operands and divergent predicates evaluated through the real CF service
 * (Rulings S5-B..S5-G, S5-STUB), and the resident and paged adapters agreeing on them (S5-P). Production publishes no
 * `formula:*` capability yet (the Excel captures have not run), so every data-driven Lesson 1 assertion below asserts
 * withheld today and starts comparing with Excel's DisplayFormat readback automatically once a capture is published. The
 * paint paths are exercised under injected evidence only.
 */
import { describe, expect, it } from 'vitest'
import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { cfDocumentEvaluationOf, cfEvaluationContextFor, compileCfFormula, isCfFormulaRefusal } from '../src/xlsx/cf/formula'
import { prepareCfPaint, residentSource } from '../src/xlsx/cf/render'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { createCfService } from '../src/xlsx/cf/service'
import { CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import { VERIFIED_CAPABILITIES, injectCfFormulaSelections, withExactCfEvidence, withInjectedCfEvidence } from '../src/xlsx/cf/evidence'
import { CF_FORMULA_PRODUCTION_READINGS, cfFormulaReading } from '../src/xlsx/cf/formulaReadings'
import { cfRulesVersion, createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { isCellInSqref } from '../scripts/cf-refs/verify'
import { CF_STAGE5_CAPABILITIES, CF_STAGE5_FIXTURES, CF_STAGE5_REUSED_CAPTURES, type CfManifest } from '../scripts/cf-refs/cases'
import { CF_STAGE5_PRODUCTION_READINGS } from '../scripts/cf-refs/audit'
import type { CfCellState, CfOutcome, CfValueSource } from '../src/xlsx/cf/model'
import { FORMULA_CAPS, GREEN, RED, SELECTIONS, parseRules, prepare, rc, sheetOf, shown } from './helpers/cf-formula-rules'

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const manifest = (): CfManifest => JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
const expr = (sqref: string, priority: number, formula: string, o: { dxf?: number; stop?: boolean; type?: string; attrs?: string } = {}): string =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="${o.type ?? 'expression'}" priority="${priority}" dxfId="${o.dxf ?? 0}"${o.stop ? ' stopIfTrue="1"' : ''}${o.attrs ?? ''}><formula>${formula}</formula></cfRule></conditionalFormatting>`
const cellIs = (sqref: string, priority: number, operator: string, formulas: readonly string[], dxf = 0): string =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="cellIs" priority="${priority}" dxfId="${dxf}" operator="${operator}">${formulas.map(f => `<formula>${f}</formula>`).join('')}</cfRule></conditionalFormatting>`
const fill = (o: CfOutcome<CfCellState> | undefined): string => shown(o).withheld ? 'withheld' : shown(o).fill

/* ------------------------------------------------------------------ */
/* Lesson 1: production against Excel's DisplayFormat, data-driven.    */
/* ------------------------------------------------------------------ */

type Token = Readonly<{ fill: string; strike: boolean }>
/**
 * The Lesson 1 verdict for one designated DXF cell, pure in its inputs. A cell whose capabilities are not all verified, or
 * that has no DisplayFormat observation, may carry no formula-sourced property. A verified, observed cell must show
 * exactly Excel's fill and strikethrough. Returns a violation, or undefined.
 */
export function stage5CellVerdict(cell: Readonly<{ key: string; verified: boolean; observed?: Token }>, got: Token, formulaSourced: boolean): string | undefined {
  if (!cell.verified || !cell.observed) return formulaSourced ? `${cell.key}: a formula-sourced property paints while the cell is ${cell.verified ? 'unobserved' : 'withheld'}` : undefined
  return got.fill === cell.observed.fill && got.strike === cell.observed.strike ? undefined
    : `${cell.key}: production shows ${got.fill}${got.strike ? '+strike' : ''}, Excel's DisplayFormat ${cell.observed.fill}${cell.observed.strike ? '+strike' : ''}`
}

describe('expression rule paints exactly the DisplayFormat cells (Lesson 1)', () => {
  it('the verdict: withheld cells carry nothing; a verified cell is Excel\'s observation, never computed by the function under test', () => {
    const red = { fill: RED, strike: false }, none = { fill: 'none', strike: false }
    expect(stage5CellVerdict({ key: 'a', verified: false }, none, false)).toBeUndefined()
    expect(stage5CellVerdict({ key: 'a', verified: false }, red, true)).toMatch(/withheld/)
    expect(stage5CellVerdict({ key: 'a', verified: true }, red, true)).toMatch(/unobserved/)
    expect(stage5CellVerdict({ key: 'a', verified: true, observed: red }, red, true)).toBeUndefined()
    expect(stage5CellVerdict({ key: 'a', verified: true, observed: none }, none, false)).toBeUndefined()
    expect(stage5CellVerdict({ key: 'a', verified: true, observed: none }, red, true)).toMatch(/Excel's DisplayFormat none/)
    expect(stage5CellVerdict({ key: 'a', verified: true, observed: red }, none, false)).toMatch(/production shows none/)
    expect(stage5CellVerdict({ key: 'a', verified: true, observed: { fill: 'none', strike: true } }, none, false)).toMatch(/\+strike/)
  })

  it('production (the real route, with the document\'s evaluation context) matches every verified cell and withholds the rest', async () => {
    const stage5 = manifest().cases.filter(c => c.stage === 5)
    expect(stage5.length).toBeGreaterThan(20)
    let covered = 0, compared = 0, preparedSheets = 0, preparedCells = 0
    const violations: string[] = []
    for (const c of stage5) {
      const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
      const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
      const formulaMarkers = (ref: string) => c.markers.filter(m => m.formula !== undefined && isCellInSqref(ref, m.sqref))
      // DXF-slot cells only: a visual's formula bound is compared by crop in Task 37.
      const designated = c.expected.cells.filter(e => { const ms = formulaMarkers(e.ref); return ms.length > 0 && ms.every(m => m.ruleType !== 'colorScale' && m.ruleType !== 'dataBar' && m.ruleType !== 'iconSet') })
      if (!sheet.conditionalFormatting) { covered += designated.length; continue }
      preparedSheets++
      let observations: Record<string, { displayFormatFill?: string; baseFill?: string; strikethrough?: boolean }> = {}
      const capture = c.status === 'measured' ? c.captures.find(x => x.appearance === 'screen') ?? c.captures[0] : undefined
      if (capture) observations = (JSON.parse(readFileSync(join(REPO_ROOT, capture.sidecarPath), 'utf8')) as { formulaObservations?: typeof observations }).formulaObservations ?? {}
      const memory = createCfMemory()
      const evaluation = cfEvaluationContextFor(cfDocumentEvaluationOf(doc), Date.now)
      const service = createCfService(sheet.conditionalFormatting, memory, { evidence: cfRuntimeEvidence, evaluation, sheetName: sheet.name })
      const stamp = { owner: 'lesson1', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting), viewGen: 0, structureGen: 0, clockGen: 0 }
      const rects = designated.map(e => { const { row, col } = rc(e.ref); return { firstRow: row, lastRow: row, firstCol: col, lastCol: col } })
      if (rects.length === 0) { service.dispose(); continue }
      const lease = await service.prepareCf({ stamp, sourceRects: rects, signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(sheet, stamp, memory, 'engine'))
      for (const e of designated) {
        covered++; preparedCells++
        const { row, col } = rc(e.ref)
        const outcome = lease.value.at(row, col)
        const got = shown(outcome)
        const formulaSourced = outcome.kind === 'ready' && Object.keys(outcome.value.properties).length > 0
        const verified = c.status === 'measured' && formulaMarkers(e.ref).every(m => VERIFIED_CAPABILITIES.has(m.capabilityId))
        const rec = observations[e.ref]
        const observed = rec && typeof rec.displayFormatFill === 'string'
          ? { fill: rec.displayFormatFill.toUpperCase() === (rec.baseFill ?? 'FFFFFFFF').toUpperCase() ? 'none' : rec.displayFormatFill.toUpperCase(), strike: rec.strikethrough === true } : undefined
        if (verified && observed) compared++
        const v = stage5CellVerdict({ key: `${c.id}:${e.ref}`, verified, ...(observed ? { observed } : {}) }, got, formulaSourced)
        if (v) violations.push(v)
      }
      lease.release(); service.dispose()
    }
    expect(violations).toEqual([])
    expect(covered).toBeGreaterThan(400)
    // Non-vacuous: the service ran on (nearly) every Stage 5 sheet, not only the parse-refused branch.
    expect(preparedSheets).toBeGreaterThanOrEqual(25)
    expect(preparedCells).toBeGreaterThan(400)
    // No Stage 5 capture is published yet: nothing is compared, everything above asserted withheld. A published capture
    // moves cells into `compared` with no edit here.
    const anyPublished = [...VERIFIED_CAPABILITIES].some(id => id.startsWith('formula:'))
    if (!anyPublished) expect(compared).toBe(0)
  }, 180_000)

  it('under injected evidence a synthetic expression paints exactly where its formula is true', async () => {
    const rules = parseRules(expr('B1:B4', 1, '$A1&gt;10'), FORMULA_CAPS)
    const sheet = sheetOf(rules, { A1: 5, A2: 10, A3: 11, A4: 'x', B1: 1, B2: 1, B3: 1, B4: 1 })
    const out = await prepare(sheet, ['B1', 'B2', 'B3', 'B4'], FORMULA_CAPS)
    // "x" > 10 is TRUE in Excel's comparison order (text above numbers), as the engine evaluates it.
    expect(['B1', 'B2', 'B3', 'B4'].map(r => fill(out.get(r)))).toEqual(['none', 'none', RED, RED])
    // Production evidence: withheld, whatever the formula computes.
    const plain = await prepare(sheet, ['B3'], [])
    expect(fill(plain.get('B3'))).toBe('withheld')
  })
})

describe('the formula arm paces between cells (S5-N round-2 #2)', () => {
  // One row of targets: the resident source paces per row, so every yield honoured here is the service's own, per cell.
  const rules = () => parseRules(expr('B1:G1', 1, 'SUM(A$2:A$4)&gt;0'), FORMULA_CAPS)
  const cells: Record<string, number> = {}
  for (const col of 'ABCDEFG') for (let r = 1; r <= 4; r++) cells[`${col}${r}`] = r
  /** A pacing budget that records spends; `onCall` runs on a formula call's admission (the only spend of 10+ units). */
  const recording = (onCall: (n: number) => 'continue' | 'yield') => {
    const log: string[] = []
    let paused = false, used = 0
    const budget = {
      get units() { return used }, get remaining() { return CF_LIMITS.activeUnits - used },
      spend(n: number): 'continue' | 'yield' | 'refused' { used += n; if (n >= 10) { log.push(`call:${n}`); const r = onCall(n); if (r === 'yield') paused = true; return r } return paused ? 'yield' : 'continue' },
      resume() { if (paused) log.push('resume'); paused = false },
    }
    return { budget, log }
  }
  const run = async (budget: ReturnType<typeof recording>['budget'], signal: AbortSignal) => {
    const sheet = sheetOf(rules(), cells)
    const memory = createCfMemory()
    const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: { allows: cap => FORMULA_CAPS.includes(cap) }, evaluation: cfEvaluationContextFor(cfDocumentEvaluationOf({} as never), () => 0), budget: () => budget })
    const stamp = { owner: 't', sheetId: 'sheet1', workbookValueGen: 0, cfGen: 1, viewGen: 0, structureGen: 0, clockGen: 0 }
    try {
      // The source pages the row in one page and never touches the pacing budget, so any resume is the service's own.
      const resident = residentSource(sheet, stamp, memory, 'engine')
      const source: CfValueSource = { ...resident, stamp: resident.stamp, bounds: resident.bounds, async *pages(id, rects, b, sig) {
        // Scans (the dependency snapshot) read through the resident source on their own budgets.
        if (b !== budget) { yield* resident.pages(id, rects, b, sig); return }
        const cols = [1, 2, 3, 4, 5, 6]
        yield { value: { rows: Uint32Array.from(cols.map(() => 0)), cols: Uint16Array.from(cols), values: cols.map(() => ({ kind: 'number' as const, value: 1 })) }, bytes: 0, release() {}, transfer() { throw new Error('unused') } }
      } }
      const lease = await service.prepareCf({ stamp, sourceRects: [{ firstRow: 0, lastRow: 0, firstCol: 1, lastCol: 6 }], signal, grantBytes: CF_LIMITS.scratchBytes }, source)
      lease.release()
    } finally { service.dispose() }
  }
  it('a yield earned inside a call is honoured (resume) before the next cell\'s call', async () => {
    const { budget, log } = recording(() => 'yield')
    await run(budget, new AbortController().signal)
    const calls = log.filter(x => x.startsWith('call:')).length
    expect(calls).toBe(6)
    // Every call after the first is preceded by a resume: the latch was honoured between cells, never inside a call.
    for (let i = 1; i < log.length; i++) if (log[i]!.startsWith('call:')) expect(log[i - 1], log.join(' ')).toBe('resume')
  })
  it('an abort during a call is checked before the next cell: no later call runs', async () => {
    const controller = new AbortController()
    const { budget, log } = recording(() => { controller.abort(); return 'continue' })
    let afterAbort = 0
    const spend = budget.spend.bind(budget)
    budget.spend = (n: number) => { if (controller.signal.aborted) afterAbort++; return spend(n) }
    await expect(run(budget, controller.signal)).rejects.toMatchObject({ code: 'cancelled' })
    expect(log.filter(x => x.startsWith('call:'))).toHaveLength(1)
    // Nothing is spent on the pacing budget after the abort: the very next checkpoint throws.
    expect(afterAbort).toBe(0)
  })
})

describe('manual calculation withholds formula rules (S5-L)', () => {
  it('under calcMode="manual" a formula rule is withheld, whatever capabilities are held; auto evaluates it', async () => {
    const sheet = sheetOf(parseRules(expr('B1', 1, '$A1&gt;10'), FORMULA_CAPS), { A1: 11, B1: 1 })
    const caps = [...FORMULA_CAPS, 'formula:calc-manual-saved-values', 'formula:calc-manual-today']
    expect(fill((await prepare(sheet, ['B1'], caps)).get('B1'))).toBe(RED)
    const manual = await prepare(sheet, ['B1'], caps, { evaluation: { calcMode: 'manual', valuePolicy: 'engine', dateSystem: '1900', timeZone: 'UTC' } })
    expect(fill(manual.get('B1'))).toBe('withheld')
    const noTable = await prepare(sheet, ['B1'], caps, { evaluation: { calcMode: 'autoNoTable', valuePolicy: 'engine', dateSystem: '1900', timeZone: 'UTC' } })
    expect(fill(noTable.get('B1'))).toBe(RED)
  })
})

/* ------------------------------------------------------------------ */
/* Formula operands (Ruling S5-F).                                     */
/* ------------------------------------------------------------------ */

describe('formula operands reuse the Stage 1 comparators', () => {
  it('cellIs > $F$1 with F1 = 100 separates 99/100/101 exactly as literal 100 does', async () => {
    const rules = parseRules(cellIs('A1:A3', 1, 'greaterThan', ['$F$1']) + cellIs('C1:C3', 2, 'greaterThan', ['100']), FORMULA_CAPS)
    const sheet = sheetOf(rules, { A1: 99, A2: 100, A3: 101, C1: 99, C2: 100, C3: 101, F1: 100 })
    const out = await prepare(sheet, ['A1', 'A2', 'A3', 'C1', 'C2', 'C3'], FORMULA_CAPS)
    const formula = ['A1', 'A2', 'A3'].map(r => fill(out.get(r))), literal = ['C1', 'C2', 'C3'].map(r => fill(out.get(r)))
    expect(literal).toEqual(['none', 'none', RED])
    expect(formula).toEqual(literal)
    // Without the operand kind's capability the operand rule is withheld, the literal one is not.
    const gated = await prepare(sheet, ['A3', 'C3'], FORMULA_CAPS.filter(c => c !== 'formula:cellis-operand-reference'))
    expect([fill(gated.get('A3')), fill(gated.get('C3'))]).toEqual(['withheld', RED])
  })

  it('an operand #N/A is unresolved on that cell only', async () => {
    const rules = parseRules(cellIs('A1:A3', 1, 'greaterThan', ['B1']), FORMULA_CAPS)
    const sheet = sheetOf(rules, { A1: 5, A2: 5, A3: 5, B1: 1, B2: { err: '#N/A' }, B3: 9 })
    const out = await prepare(sheet, ['A1', 'A2', 'A3'], FORMULA_CAPS)
    expect(['A1', 'A2', 'A3'].map(r => fill(out.get(r)))).toEqual([RED, 'withheld', 'none'])
    expect(out.get('A2')).toMatchObject({ kind: 'ready', value: { unresolvedWrites: expect.anything() } })
    // Never a fallback: the Stage 1 comparator decides a number against an error literal as no match (`equal`), but an
    // operand that evaluates to an error refuses the cell (spec line 126).
    const eq = parseRules(cellIs('A1:A3', 1, 'equal', ['B1']), FORMULA_CAPS)
    const eqOut = await prepare(sheetOf(eq, { A1: 1, A2: 5, A3: 9, B1: 1, B2: { err: '#N/A' }, B3: 8 }), ['A1', 'A2', 'A3'], FORMULA_CAPS)
    expect(['A1', 'A2', 'A3'].map(r => fill(eqOut.get(r)))).toEqual([RED, 'withheld', 'none'])
  })

  it('a reversed between follows cellIs:between-reversed', async () => {
    const rules = parseRules(cellIs('A1:A3', 1, 'between', ['$F$1', '$F$2']), FORMULA_CAPS)
    const sheet = sheetOf(rules, { A1: 98, A2: 100, A3: 102, F1: 101, F2: 99 })
    const without = await prepare(sheet, ['A1', 'A2', 'A3'], FORMULA_CAPS)
    expect(['A1', 'A2', 'A3'].map(r => fill(without.get(r)))).toEqual(['withheld', 'withheld', 'withheld'])
    const withCap = await prepare(sheet, ['A1', 'A2', 'A3'], [...FORMULA_CAPS, 'cellIs:between-reversed'])
    expect(['A1', 'A2', 'A3'].map(r => fill(withCap.get(r)))).toEqual(['none', RED, 'none'])
  })

  it('blank and text operands follow their own capabilities', async () => {
    const rules = parseRules(cellIs('A1:A2', 1, 'greaterThan', ['$F$1']) + cellIs('B1', 2, 'equal', ['$F$2']), FORMULA_CAPS)
    const sheet = sheetOf(rules, { A1: 1, A2: -1, B1: 'ab', F2: 'ab' })
    const none = await prepare(sheet, ['A1', 'B1'], FORMULA_CAPS)
    expect([fill(none.get('A1')), fill(none.get('B1'))]).toEqual(['withheld', 'withheld'])
    const caps = [...FORMULA_CAPS, 'formula:cellis-operand-blank', 'formula:cellis-operand-text']
    const zero = await prepare(sheet, ['A1', 'A2', 'B1'], caps, { selections: { ...SELECTIONS, 'formula:cellis-operand-blank': 'blank-zero' } })
    // A text operand stays withheld even with its capability: production registers no text-operand reading (the Stage 1
    // comparator does not reproduce the fixture's text-compare reading, tests/xlsx-cf-formula-readings.test.ts).
    expect(['A1', 'A2', 'B1'].map(r => fill(zero.get(r)))).toEqual([RED, 'none', 'withheld'])
    // An unlisted selection is never honoured.
    const unlisted = await prepare(sheet, ['A1'], caps, { selections: { ...SELECTIONS, 'formula:cellis-operand-blank': 'blank-ignored' } })
    expect(fill(unlisted.get('A1'))).toBe('withheld')
  })
})

/* ------------------------------------------------------------------ */
/* Divergent predicates (Ruling S5-G).                                 */
/* ------------------------------------------------------------------ */

describe('divergent predicates follow the formula (S5-G)', () => {
  const xml = expr('D2:D4', 1, 'ISBLANK(D2)', { type: 'containsBlanks' })
    + expr('G2:G3', 2, 'NOT(ISBLANK(G2))', { type: 'notContainsBlanks' })
    + expr('J2:J3', 3, 'ISNA(J2)', { type: 'containsErrors' })
    + expr('M2:M3', 4, 'SEARCH("zz",M2)', { type: 'containsText', attrs: ' operator="containsText" text="ab"' })
  const cells = { D3: '', D4: 1, G2: 0, J2: { err: '#N/A' }, J3: { err: '#DIV/0!' }, M2: 'zz', M3: 'ab' } as const
  it('ISBLANK matches the unstored blank and not "", NOT(ISBLANK) matches 0, ISNA matches #N/A and not #DIV/0!; SEARCH stays withheld', async () => {
    const rules = parseRules(xml, FORMULA_CAPS)
    expect(rules.rules.map(r => r.definition.kind)).toEqual(['expression', 'expression', 'expression', 'unresolved'])
    const sheet = sheetOf(rules, cells)
    const refs = ['D2', 'D3', 'D4', 'G2', 'G3', 'J2', 'J3', 'M2', 'M3']
    const out = await prepare(sheet, refs, FORMULA_CAPS)
    expect(Object.fromEntries(refs.map(r => [r, fill(out.get(r))]))).toEqual({
      D2: RED, D3: 'none', D4: 'none', G2: RED, G3: 'none', J2: RED, J3: 'none', M2: 'withheld', M3: 'withheld',
    })
  })
  it('without formula:predicate-divergent at evaluation the divergent rules are withheld', async () => {
    const rules = parseRules(xml, FORMULA_CAPS)
    const out = await prepare(sheetOf(rules, cells), ['D2', 'G2', 'J2'], FORMULA_CAPS.filter(c => c !== 'formula:predicate-divergent'))
    expect(['D2', 'G2', 'J2'].map(r => fill(out.get(r)))).toEqual(['withheld', 'withheld', 'withheld'])
  })
})

/* ------------------------------------------------------------------ */
/* timePeriod readings (Rulings S5-TP-READING, S5-TP-MULTI).          */
/* ------------------------------------------------------------------ */

describe('timePeriod formula readings', () => {
  it('S5-TP-READING: the divergent timePeriod reading production registers is the formula reading', () => {
    expect(CF_FORMULA_PRODUCTION_READINGS['formula:predicate-timeperiod-divergent']).toBe('formula')
    expect(CF_STAGE5_PRODUCTION_READINGS['formula:predicate-timeperiod-divergent']).toBe('formula')
  })
  it('S5-TP-MULTI: a multi-range canonical timePeriod rule stays withheld until formula:anchor-multi-range is evidenced', async () => {
    const tp = (sqref: string, formula: string) => `<conditionalFormatting sqref="${sqref}"><cfRule type="timePeriod" priority="1" dxfId="0" timePeriod="today"><formula>${formula}</formula></cfRule></conditionalFormatting>`
    const single = parseRules(tp('B2:B4', 'FLOOR(B2,1)=TODAY()'), [])
    expect(single.rules[0]!.definition.kind).toBe('timePeriod') // canonical single range: the Stage 1 attribute predicate
    for (const caps of [[], FORMULA_CAPS, [...FORMULA_CAPS, 'formula:predicate-timeperiod-divergent']]) {
      const multi = parseRules(tp('D5:D8 B2:B4', 'FLOOR(D5,1)=TODAY()'), caps)
      expect(multi.rules[0]!.definition.kind, caps.join(',')).toBe('unresolved')
    }
  })
})

/* ------------------------------------------------------------------ */
/* Stub successors (Ruling S5-STUB).                                   */
/* ------------------------------------------------------------------ */

describe('stub successors repaint only where Excel did (S5-STUB)', () => {
  const SUCCESSORS: Readonly<Record<string, readonly string[]>> = {
    'stage1-blanks-errors': ['D2', 'G6', 'J8', 'M6'],
    'stage1-priority-composition': ['G2', 'G3'],
    'stage2-scale-composition': ['L2', 'L3', 'L4', 'L5'],
    'stage3-bar-composition': ['I2', 'I3', 'I4'],
    'stage4-icon-composition': ['L2', 'L3', 'L4', 'L5'],
  }
  it('the inventory names exactly the reused captures', () => {
    for (const r of CF_STAGE5_REUSED_CAPTURES) expect(SUCCESSORS[r.caseId], r.caseId).toEqual(expect.arrayContaining([...r.refs]))
  })
  it('every stub cell is Excel\'s recorded outcome or withheld, under production and under injected successors', async () => {
    const all = [...FORMULA_CAPS, 'formula:composition-stop-dxf', 'formula:composition-stop-scale', 'formula:composition-stop-bar', 'formula:composition-stop-icon',
      'formula:composition-stub-successor-dxf', 'formula:composition-stub-successor-scale', 'formula:composition-stub-successor-bar', 'formula:composition-stub-successor-icon']
    const painted: string[] = []
    for (const inject of [[], all]) {
      await withInjectedCfEvidence(inject, async () => {
        injectCfFormulaSelections(SELECTIONS)
        try {
          for (const [caseId, refs] of Object.entries(SUCCESSORS)) {
            const c = manifest().cases.find(x => x.id === caseId)!
            const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))))
            const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
            const paint = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence, evaluation: cfDocumentEvaluationOf(doc) })
            for (const ref of refs) {
              const exp = c.expected.cells.find(e => e.ref === ref)!
              const { row, col } = rc(ref)
              const st = paint!.layer.styleAt(row, col, undefined)
              const got = Boolean(st && (st.fillColor || st.color || st.bold !== undefined || st.strike))
              if (!got) continue // withheld
              painted.push(`${inject.length > 0 ? 'injected' : 'production'}:${caseId}:${ref}`)
              // Painted: Excel's recorded outcome. For a Stage 3/4 cell the expected fill is the non-visual rule's.
              expect(exp.matched === true || c.stage >= 3, `${caseId}:${ref} paints where Excel recorded no match`).toBe(true)
              if (exp.fillRgb) expect(st!.fillColor, `${caseId}:${ref} fill`).toBe(exp.fillRgb)
            }
            paint!.dispose()
          }
        } finally { injectCfFormulaSelections(null) }
      })
    }
    // Production paints none of them; under the injected successors the stage1-blanks-errors cells repaint as Excel did.
    expect(painted.filter(p => p.startsWith('production:'))).toEqual([])
    expect(painted).toEqual(expect.arrayContaining(['injected:stage1-blanks-errors:D2', 'injected:stage1-blanks-errors:G6', 'injected:stage1-blanks-errors:J8', 'injected:stage1-blanks-errors:M6']))
  }, 120_000)

  it('the re-pointing commit touches only the inventoried conformance lines', ctx => {
    let sha = ''
    try {
      sha = execFileSync('git', ['log', '--format=%H', '--fixed-strings', '--grep=S5-STUB re-pointing', '--', 'tests/xlsx-cf-oracle-conformance.test.ts'], { cwd: REPO_ROOT, encoding: 'utf8' }).trim().split('\n')[0] ?? ''
    } catch { sha = '' }
    // A history without the commit (CI's shallow checkout, or main after a squash merge) cannot be checked: an explicit
    // skip. The commit is named in the Task 35 report and is checked on the branch.
    if (!sha) { ctx.skip(); return }
    const files = execFileSync('git', ['show', '--format=', '--name-only', sha], { cwd: REPO_ROOT, encoding: 'utf8' }).trim().split('\n')
    expect(files).toEqual(['tests/xlsx-cf-oracle-conformance.test.ts'])
    const diff = execFileSync('git', ['show', '--format=', '--unified=0', sha], { cwd: REPO_ROOT, encoding: 'utf8' })
    // Old-file windows of the inventoried assertions (STAGE5_OR_DEFERRED_WITHHELD and its use, the L2:L5 findings and their
    // Stage 2 use, STAGE3_NONBAR_WITHHELD and its count, the stage 4 stub and its count, and the counters).
    const windows: ReadonlyArray<readonly [number, number]> = [[107, 118], [190, 200], [222, 228], [250, 258], [280, 290], [350, 368], [545, 615]]
    const hunks = [...diff.matchAll(/^@@ -(\d+)(?:,(\d+))? \+\d+(?:,\d+)? @@/gm)].map(m => [Number(m[1]), Number(m[1]) + Math.max(Number(m[2] ?? 1) - 1, 0)] as const)
    expect(hunks.length).toBeGreaterThan(0)
    for (const [a, b] of hunks) expect(windows.some(([lo, hi]) => a >= lo && b <= hi), `hunk at old lines ${a}-${b}`).toBe(true)
  })

  it('src/xlsx/cf/verifiedCapabilities.ts is unchanged from main at this task', ctx => {
    // Retires itself once Stage 5 publishes (Task 39 writes VERIFIED_STAGE5_CAPABILITIES into the file).
    if (readFileSync(join(REPO_ROOT, 'src/xlsx/cf/verifiedCapabilities.ts'), 'utf8').includes('VERIFIED_STAGE5_CAPABILITIES')) { ctx.skip(); return }
    // The branch's base: origin/main when fetched (a local main can lag behind it), else main. A ref without the file (a
    // stale local main) cannot be the base and is skipped.
    const git = (...args: string[]): string => execFileSync('git', args, { cwd: REPO_ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] })
    const base = ['origin/main', 'main'].find(ref => { try { git('cat-file', '-e', `${ref}:src/xlsx/cf/verifiedCapabilities.ts`); return true } catch { return false } })
    if (!base) { ctx.skip(); return }
    expect(git('diff', base, '--', 'src/xlsx/cf/verifiedCapabilities.ts')).toBe('')
  })
})

/* ------------------------------------------------------------------ */
/* Capability ids and readings.                                        */
/* ------------------------------------------------------------------ */

describe('every formula capability production queries is a catalog capability with a catalog reading', () => {
  const catalog = new Set(CF_STAGE5_CAPABILITIES.map(e => e.capabilityId))
  const hypotheses = new Map<string, Set<string>>()
  for (const fx of CF_STAGE5_FIXTURES) for (const cfg of fx.configurations) for (const h of Object.keys(cfg.hypotheses)) hypotheses.set(cfg.capabilityId, new Set([...(hypotheses.get(cfg.capabilityId) ?? []), h]))
  const reuseOnly = new Set(CF_STAGE5_REUSED_CAPTURES.map(r => r.capabilityId).filter(id => !CF_STAGE5_FIXTURES.some(fx => fx.configurations.some(c => c.capabilityId === id))))
  it('each production reading is a hypothesis of its catalog capability (or the recorded reuse)', () => {
    for (const [id, reading] of Object.entries(CF_FORMULA_PRODUCTION_READINGS)) {
      expect(catalog.has(id), id).toBe(true)
      for (const r of typeof reading === 'string' ? [reading] : reading) {
        if (reuseOnly.has(id)) expect(r, id).toBe('recorded')
        else expect(hypotheses.get(id)?.has(r), `${id}: ${r}`).toBe(true)
      }
    }
    // The audit spreads the same table, so publication and production agree.
    for (const [id, reading] of Object.entries(CF_FORMULA_PRODUCTION_READINGS)) expect(CF_STAGE5_PRODUCTION_READINGS[id]).toEqual(reading)
  })
  it('every literal and templated formula:* id in the CF sources is a catalog id', () => {
    const src = ['service.ts', 'compose.ts', 'formula.ts', 'parse.ts', 'formulaReadings.ts', 'predicates.ts'].map(f => readFileSync(join(REPO_ROOT, 'src/xlsx/cf', f), 'utf8')).join('\n')
    const literal = [...src.matchAll(/'(formula:[a-z0-9-]+)'/g)].map(m => m[1]!)
    const families = ['logical', 'information', 'text', 'math', 'aggregate', 'conditional-aggregate', 'date', 'row-column', 'today']
    const templated = [...families.map(f => `formula:fn-${f}`), ...['number', 'boolean', 'text', 'blank', 'error', 'range'].map(k => `formula:truth-${k}`),
      ...['dxf', 'scale', 'bar', 'icon'].map(k => `formula:composition-stop-${k}`), 'formula:dialect-base', 'formula:dialect-x14']
    expect(src).toMatch(/`formula:composition-stop-\$\{kind\}`/)
    for (const id of [...literal, ...templated]) expect(catalog.has(id), id).toBe(true)
  })
  it('an unregistered capability is never honoured, whatever evidence holds it', () => {
    const all = (): boolean => true
    for (const id of ['formula:truth-range', 'formula:fn-today', 'formula:anchor-multi-range', 'formula:cross-sheet-reference']) expect(cfFormulaReading(id, all), id).toBeUndefined()
    expect(cfFormulaReading('formula:truth-number', all)).toBe('number-nonzero-true')
    expect(cfFormulaReading('formula:truth-number', () => false)).toBeUndefined()
    // A selection needs its matched reading: none is published, so none resolves.
    expect(cfFormulaReading('formula:truth-text', all)).toBeUndefined()
  })
})

/* ------------------------------------------------------------------ */
/* Every view agrees (Ruling S5-P): resident and paged adapters.       */
/* ------------------------------------------------------------------ */

describe('formula rules evaluate identically on the resident and paged adapters', () => {
  const PARITY = ['stage5-formula-truthiness', 'stage5-formula-functions', 'stage5-cellis-operands', 'stage5-predicate-formulas', 'stage5-composition', 'stage5-formula-anchor']
  it('the Stage 5 fixtures give deep-equal per-cell states under injected evidence (non-vacuous)', async () => {
    const caps = [...FORMULA_CAPS, 'formula:composition-stop-dxf', 'formula:composition-overlapping', 'formula:composition-expression-vs-scale', 'formula:hoist-aggregate', 'formula:hoist-countif-frequency']
    let paintedCells = 0, compared = 0
    await withInjectedCfEvidence(caps, async () => {
      injectCfFormulaSelections(SELECTIONS)
      try {
        for (const id of PARITY) {
          const c = manifest().cases.find(x => x.id === id)!
          const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, c.saved)))
          const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
          const sheet = doc.sheets.find(s => s.name === c.sheet) ?? doc.sheets[0]!
          const refs = c.expected.cells.map(e => e.ref).filter(r => c.markers.some(m => m.formula !== undefined && isCellInSqref(r, m.sqref)))
          const rects = refs.map(r => { const { row, col } = rc(r); return { r0: row, r1: row, c0: col, c1: col } })
          // Resident leg.
          const memory = createCfMemory()
          const evaluation = cfEvaluationContextFor(cfDocumentEvaluationOf(doc), Date.now)
          const service = createCfService(sheet.conditionalFormatting!, memory, { evidence: cfRuntimeEvidence, evaluation, sheetName: sheet.name })
          const stamp = { owner: 'parity', sheetId: sheet.sheetId ?? sheet.name, workbookValueGen: 0, cfGen: cfRulesVersion(sheet.conditionalFormatting!), viewGen: 0, structureGen: 0, clockGen: 0 }
          const lease = await service.prepareCf({ stamp, sourceRects: rects.map(r => ({ firstRow: r.r0, lastRow: r.r1, firstCol: r.c0, lastCol: r.c1 })), signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(sheet, stamp, memory, 'cached'))
          const resident = new Map<string, unknown>()
          for (const r of refs) {
            const { row, col } = rc(r)
            const o = lease.value.at(row, col)
            if (o.kind === 'no-match') continue
            resident.set(`${row},${col}`, o.kind === 'unresolved' ? { unresolved: true } : { properties: { ...o.value.properties }, unresolvedWrites: [...(o.value.unresolvedWrites ?? [])].sort() })
            if (o.kind === 'ready' && Object.keys(o.value.properties).length > 0) paintedCells++
          }
          lease.release(); service.dispose()
          // Paged leg: the worker route over the same bytes.
          const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes.buffer as ArrayBuffer]) }, { docId: `parity-${id}`, profile: 'D' })
          try {
            await session.done
            const source = createPagedWorkbookSource(session)
            const sheetId = source.sheets[0]!.sheetId
            const store = source.store(sheetId)
            const result = await store.readConditionalFormatting!({ rects, stamp: { owner: 'parity', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: 8 * 1024 * 1024 }, createCfMemory(), { signal: new AbortController().signal })
            expect(result?.kind, id).toBe('ready')
            const paged = new Map<string, unknown>()
            for (const s of result!.states) paged.set(`${s.row},${s.col}`, s.unresolvedReason !== undefined ? { unresolved: true } : { properties: { ...s.properties }, unresolvedWrites: [...(s.unresolvedWrites ?? [])].sort() })
            expect(paged, id).toEqual(resident)
            compared += resident.size
          } finally { await session.release() }
        }
      } finally { injectCfFormulaSelections(null) }
    })
    expect(compared).toBeGreaterThan(100)
    expect(paintedCells).toBeGreaterThan(20)
  }, 180_000)
})

void compileCfFormula; void isCfFormulaRefusal; void GREEN; void withExactCfEvidence
