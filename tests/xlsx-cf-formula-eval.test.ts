// @vitest-environment node
/**
 * Stage 5 Task 34: translation (S5-C), footprint (S5-N caps), result shape and truth readings (S5-D), the values a
 * formula reads (S5-E), and the discard rules (S5-E round-1 #1, Review Focus 1). Expectations come from the brief and the
 * engine's own evaluation on the full store (tests/helpers/cf-formula.ts), never from the CF evaluator under test.
 */
import { describe, expect, it } from 'vitest'
import '../src/xlsx/formula/functions'
import type { AstNode } from '../src/xlsx/formula/types'
import {
  cfElementToLiteral, cfFormulaFootprint, cfLiteralToElement, cfTruth, translateCfFormula,
  type CfResult, type CfTruthReading,
} from '../src/xlsx/cf/formula'
import { CF_FORMULA_NODE_BYTES, CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import type { CfLiteral } from '../src/xlsx/cf/model'
import { a1, compile, rectOf } from './helpers/cf-formula'

const ready = <T>(o: { kind: string; value?: T }): T => { expect(o.kind).toBe('ready'); return o.value as T }
type CellNode = Extract<AstNode, { type: 'cell' }>
const cellOf = (node: AstNode): CellNode => {
  const walk = (n: AstNode): CellNode | undefined => n.type === 'cell' ? n : n.type === 'binary' ? walk(n.left) ?? walk(n.right) : n.type === 'unary' ? walk(n.expr) : n.type === 'call' ? n.args.map(walk).find(Boolean) : undefined
  const found = walk(node)
  if (!found) throw new Error('no cell node')
  return found
}

describe('translation follows the measured edge mode and pinned axes (S5-C)', () => {
  const at = (f: ReturnType<typeof compile>, target: string, edge: 'wrap' | 'ref-error' | undefined = 'ref-error') => {
    const t = a1(target), anchor = f.anchor!
    return translateCfFormula(f, t.row - anchor.row, t.col - anchor.col, edge, createCfMemory())
  }
  it('=A2>0 anchored at B2 reads A5 at B5; $A2 at C5 reads $A5; A$2 keeps row 2', () => {
    expect(cellOf(ready<{ value: AstNode }>(at(compile('A2>0', 'B2'), 'B5')).value).ref).toMatchObject({ row: 4, col: 0 })
    expect(cellOf(ready<{ value: AstNode }>(at(compile('$A2>0', 'B2'), 'C5')).value).ref).toMatchObject({ row: 4, col: 0, absCol: true })
    expect(cellOf(ready<{ value: AstNode }>(at(compile('A$2>0', 'B2'), 'C5')).value).ref).toMatchObject({ row: 1, col: 1, absRow: true })
    // A fully pinned reference is never moved, and its node is shared rather than rebuilt.
    const pinned = compile('$A$2>0', 'B2')
    expect(cellOf(ready<{ value: AstNode }>(at(pinned, 'C9')).value)).toBe(cellOf(pinned.ast))
  })
  it('past the last row: wrap reads A1, ref-error reads #REF!, and no edge evidence is unresolved', () => {
    // A1048577 is A1048576 one row further: `=A1048576` anchored at A1048575, evaluated at A1048576.
    const f = compile('A1048576>0', { row: 1048574, col: 0 })
    const wrap = translateCfFormula(f, 1, 0, 'wrap', createCfMemory())
    expect(cellOf(ready<{ value: AstNode }>(wrap).value).ref).toMatchObject({ row: 0, col: 0 })
    const refError = ready<{ value: AstNode }>(translateCfFormula(f, 1, 0, 'ref-error', createCfMemory())).value
    expect(refError).toMatchObject({ type: 'binary', left: { type: 'error', error: '#REF!' } })
    expect(translateCfFormula(f, 1, 0, undefined, createCfMemory())).toEqual({ kind: 'unresolved', reason: 'cf-formula-edge-unmeasured' })
    // Inside the sheet no edge mode is needed.
    expect(translateCfFormula(f, -3, 0, undefined, createCfMemory()).kind).toBe('ready')
  })
  it('the translated tree is reserved before allocation: one byte short builds nothing', () => {
    const f = compile('A2>B2+C2', 'D2')
    const bound = f.nodes * CF_FORMULA_NODE_BYTES
    const short = createCfMemory({ limitBytes: bound - 1 })
    expect(translateCfFormula(f, 3, 0, 'ref-error', short)).toEqual({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(short.peakBytes).toBe(0)
    const exact = createCfMemory({ limitBytes: bound })
    const lease = ready<{ value: AstNode; release(): void; bytes: number }>(translateCfFormula(f, 3, 0, 'ref-error', exact))
    expect(exact.peakBytes).toBe(bound)
    expect(lease.bytes).toBeLessThanOrEqual(bound)
    lease.release()
    expect(exact.currentBytes).toBe(0)
  })
})

describe('footprint over a target rect is the shifted union, clamped', () => {
  it('=A2>B1 anchored at C2 over C2:C9 is {A2:A9, B1:B8}', () => {
    expect(cfFormulaFootprint(compile('A2>B1', 'C2'), [rectOf('C2:C9')], a1('C2'), 'ref-error')).toEqual([{ rect: rectOf('A2:A9') }, { rect: rectOf('B1:B8') }])
  })
  it('COUNTIF($A:$A,A2) on a 50-row sheet is A1:A50 plus A2:A9', () => {
    expect(cfFormulaFootprint(compile('COUNTIF($A:$A,A2)', 'B2'), [rectOf('B2:B9')], a1('B2'), 'ref-error', { bounds: rectOf('A1:B50') }))
      .toEqual([{ rect: rectOf('A1:A50') }, { rect: rectOf('A2:A9') }])
  })
  it(`${CF_LIMITS.dependencyCells} cells admit and one more is 'over-cap'`, () => {
    expect(CF_LIMITS.dependencyCells).toBe(262_144) // Ruling S5-N caps table
    const sized = (cells: number) => cfFormulaFootprint(compile(`SUM(A1:A${cells})`, 'B1'), [rectOf('B1')], a1('B1'), 'ref-error')
    expect(sized(262_144)).toEqual([{ rect: rectOf('A1:A262144') }])
    expect(sized(262_145)).toBe('over-cap')
  })
  it('a target shifted past the edge keeps only on-sheet cells (ref-error) or adds the wrapped rows (wrap)', () => {
    const f = compile('A1048575>0', a1('B1'))
    const targets = [rectOf('B1:B4')]
    expect(cfFormulaFootprint(f, targets, a1('B1'), 'ref-error')).toEqual([{ rect: { firstRow: 1048574, lastRow: 1048575, firstCol: 0, lastCol: 0 } }])
    expect(cfFormulaFootprint(f, targets, a1('B1'), 'wrap')).toEqual([{ rect: { firstRow: 1048574, lastRow: 1048575, firstCol: 0, lastCol: 0 } }, { rect: rectOf('A1:A2') }])
  })
})

describe('truth readings by observable outcome (S5-D)', () => {
  const r = (...readings: CfTruthReading[]): ReadonlySet<CfTruthReading> => new Set(readings)
  const at = { row: 1, col: 1 }
  const res = (kind: Exclude<CfResult['kind'], 'range'>, value: unknown): CfResult => ({ kind, value } as CfResult)
  it('each kind needs its own reading; evidence for one never authorises another', () => {
    expect(cfTruth(res('number', 2), at, r('number-nonzero-true'))).toEqual({ kind: 'ready', value: true })
    expect(cfTruth(res('number', 0), at, r('number-nonzero-true'))).toEqual({ kind: 'ready', value: false })
    expect(cfTruth(res('blank', null), at, r('number-nonzero-true')).kind).toBe('unresolved')
    expect(cfTruth(res('blank', null), at, r('blank-no-match'))).toEqual({ kind: 'ready', value: false })
    expect(cfTruth(res('boolean', true), at, r('number-nonzero-true')).kind).toBe('unresolved')
    expect(cfTruth(res('boolean', true), at, r('boolean'))).toEqual({ kind: 'ready', value: true })
    expect(cfTruth(res('text', 'abc'), at, r('text-true'))).toEqual({ kind: 'ready', value: true })
    expect(cfTruth(res('text', 'abc'), at, r('text-no-match'))).toEqual({ kind: 'ready', value: false })
    expect(cfTruth(res('error', { kind: 'formula-error', code: '#N/A' }), at, r('error-no-match'))).toEqual({ kind: 'ready', value: false })
    expect(cfTruth(res('error', { kind: 'formula-error', code: '#N/A' }), at, r('error-halts'))).toEqual({ kind: 'ready', value: 'halt' })
  })
  it('two readings for one kind are ambiguous and unresolved; a non-finite number is unresolved', () => {
    expect(cfTruth(res('text', 'x'), at, r('text-true', 'text-no-match')).kind).toBe('unresolved')
    expect(cfTruth(res('error', { kind: 'formula-error', code: '#N/A' }), at, r('error-halts', 'error-no-match')).kind).toBe('unresolved')
    expect(cfTruth(res('number', Infinity), at, r('number-nonzero-true')).kind).toBe('unresolved')
  })
})

describe('values a formula reads equal the host literals (S5-E)', () => {
  it('blank, "", #DIV/0!, a date serial, TRUE and "1" round-trip', () => {
    const cases: Array<[CfLiteral, unknown]> = [
      [{ kind: 'blank' }, null], [{ kind: 'string', value: '' }, ''], [{ kind: 'error', value: '#DIV/0!' }, { kind: 'formula-error', code: '#DIV/0!' }],
      [{ kind: 'number', value: 46332, dateSerial: true }, 46332], [{ kind: 'boolean', value: true }, true], [{ kind: 'string', value: '1' }, '1'],
    ]
    for (const [literal, element] of cases) {
      expect(cfLiteralToElement(literal)).toEqual(element)
      // The literal comes back without `dateSerial`: a formula reads the serial as a number (S5-E).
      const { dateSerial: _d, ...plain } = literal as { dateSerial?: true }
      expect(cfElementToLiteral(cfLiteralToElement(literal)!)).toEqual(plain)
    }
  })
  it('an error code the engine does not know is unrepresentable, never a guessed error', () => {
    expect(cfLiteralToElement({ kind: 'error', value: '#GETTING_DATA' })).toBeUndefined()
    expect(cfElementToLiteral(Infinity)).toBeUndefined()
    expect(cfElementToLiteral({ kind: 'formula-matrix', rows: 1, cols: 1, values: [[1]] })).toBeUndefined()
  })
})

/* ------------------------------------------------------------------ */
/* Evaluation over the snapshot.                                       */
/* ------------------------------------------------------------------ */
import { readFileSync, existsSync } from 'node:fs'
import { vi } from 'vitest'
import * as evaluatorModule from '../src/xlsx/formula/evaluator'
import { chargeMaterialization } from '../src/xlsx/formula/evaluator'
import { composeCell } from '../src/xlsx/cf/compose'
import { evaluatePredicate } from '../src/xlsx/cf/predicates'
import { prepareCfDependencies } from '../src/xlsx/cf/dependencies'
import { cfFormulaSemantics, compileCfFormula, createCfInvariantMemo, evaluateCfFormula, isCfFormulaRefusal } from '../src/xlsx/cf/formula'
import { createCfUnitCeiling, createCfWorkBudget } from '../src/xlsx/cf/limits'
import type { CfRule } from '../src/xlsx/cf/model'
import { EVALUATION, cfEval, engineEval, fakeSource, formulaOf, formulaRule, batchOf, type Cells } from './helpers/cf-formula'
import { cfTreeReadsMatchView, planCfHoist } from '../src/xlsx/cf/formula'

vi.mock('../src/xlsx/formula/evaluator', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/formula/evaluator')>()
  return { ...actual, evaluateNode: vi.fn(actual.evaluateNode), coerceToBoolean: vi.fn(actual.coerceToBoolean) }
})
const evaluateNodeSpy = vi.mocked(evaluatorModule.evaluateNode)
const coerceSpy = vi.mocked(evaluatorModule.coerceToBoolean)
const all = (...readings: CfTruthReading[]): ReadonlySet<CfTruthReading> => new Set(readings)
const valueOf = (o: { kind: string; value?: unknown }) => { expect(o.kind).toBe('ready'); return o.value as CfResult }

describe('shape is gated before projection (S5-D, round-1 #7)', () => {
  it('=A2 over an unstored A2 is a blank result read from the snapshot, never 0, and number evidence leaves it unresolved', async () => {
    evaluateNodeSpy.mockClear(); coerceSpy.mockClear()
    const result = valueOf(await cfEval('A2', 'B2', { A3: 1 }, 'B2'))
    expect(result).toEqual({ kind: 'blank', value: null })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    expect(cfTruth(result, a1('B2'), all('number-nonzero-true')).kind).toBe('unresolved')
    expect(cfTruth(result, a1('B2'), all('blank-no-match'))).toEqual({ kind: 'ready', value: false })
    expect(coerceSpy).not.toHaveBeenCalled()
  })
  it('=A1:A3 at row 2 is a range result: unresolved on number evidence alone, A2 under range-implicit, A1 under range-top-left', async () => {
    coerceSpy.mockClear()
    const cells: Cells = { A1: 0, A2: 5, A3: 0 }
    const result = valueOf(await cfEval('A1:A3', 'B2', cells, 'B2'))
    expect(result).toMatchObject({ kind: 'range', area: rectOf('A1:A3') })
    expect(cfTruth(result, a1('B2'), all('number-nonzero-true')).kind).toBe('unresolved')
    expect(cfTruth(result, a1('B2'), all('range-implicit', 'number-nonzero-true'))).toEqual({ kind: 'ready', value: true }) // A2 = 5
    expect(cfTruth(result, a1('B2'), all('range-top-left', 'number-nonzero-true'))).toEqual({ kind: 'ready', value: false }) // A1 = 0
    expect(cfTruth(result, a1('B2'), all('range-no-match', 'number-nonzero-true'))).toEqual({ kind: 'ready', value: false })
    expect(coerceSpy).not.toHaveBeenCalled()
  })
  it('formula.ts never imports evaluateFormulaInternal (source text)', () => {
    const source = readFileSync('src/xlsx/cf/formula.ts', 'utf8')
    const imports = source.match(/^import[^;]*?from '[^']+'/gms)!.join('\n')
    for (const name of ['evaluateFormulaInternal', 'coerceToBoolean']) {
      expect(imports, name).not.toMatch(new RegExp(`\\b${name}\\b`))
      expect(source, name).not.toMatch(new RegExp(`\\b${name}\\(`))
    }
    expect(imports).toMatch(/\bevaluateNode\b/)
  })
})

describe('values a formula reads equal the host literals (S5-E): ISBLANK and =A2=""', () => {
  it('=ISBLANK(A2) matches the unstored blank and not "" (Excel-backed: cases.ts stage1-blanks-errors D2)', async () => {
    expect(valueOf(await cfEval('ISBLANK(A2)', 'B2', { A3: 1 }, 'B2'))).toEqual({ kind: 'boolean', value: true })
    expect(valueOf(await cfEval('ISBLANK(A2)', 'B2', { A2: '' }, 'B2'))).toEqual({ kind: 'boolean', value: false })
  })
  it('=A2="" is asserted only against the stage5-formula-functions observation, never the engine\'s own answer', async () => {
    const manifest = JSON.parse(readFileSync('tests/fixtures/xlsx-cf/manifest.json', 'utf8')) as { cases: Array<{ id: string; status: string; captures: Array<{ sidecarPath: string }> }> }
    const entry = manifest.cases.find(c => c.id === 'stage5-formula-functions')!
    const sidecar = entry.captures[0]?.sidecarPath
    const captured = entry.status !== 'recipe-pending' && sidecar !== undefined && existsSync(sidecar)
    const result = valueOf(await cfEval('A2=""', 'L2', { A3: 1 }, 'L2'))
    if (!captured) {
      // No Excel observation exists yet (Task 31: recipe-pending): production holds no truth reading, so the cell withholds.
      expect(entry.status).toBe('recipe-pending')
      expect(cfTruth(result, a1('L2'), new Set()).kind).toBe('unresolved')
      return
    }
    const observed = JSON.parse(readFileSync(sidecar!, 'utf8')) as { cells?: Record<string, { fill?: string }> }
    const fill = observed.cells?.L2?.fill
    expect(fill, 'the sidecar observes L2').toBeDefined()
    expect(cfTruth(result, a1('L2'), all('boolean'))).toEqual({ kind: 'ready', value: fill !== 'none' })
  })
})

describe('a recorded unsupported feature discards a masked result (round-1 #1)', () => {
  const RED = { 'fill.pattern': 'solid', 'fill.bg': 'FFFF0000' }, GREEN = { 'fill.pattern': 'solid', 'fill.bg': 'FF00FF00', 'font.bold': true }
  const greenRule = (): CfRule => ({ id: 900, priority: 2, documentOrder: 2, stopIfTrue: false, ranges: [rectOf('B2:B4')], writes: new Set(['fill', 'font']), dxf: GREEN,
    definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] } })
  const compose = (p1: CfRule, outcome: { kind: string }, p2 = greenRule()) => composeCell([p1, p2], [
    outcome.kind === 'ready' ? cfTruth((outcome as unknown as { value: CfResult }).value, a1('B2'), all('boolean', 'number-nonzero-true')) : outcome as never,
    evaluatePredicate(p2, a1('B2'), { kind: 'number', value: 5 }, undefined, { allows: () => true }),
  ], { hasEvidence: () => true })
  it('IFERROR(COUNTIF(A:A,TRUE),TRUE): the engine returns TRUE while recording; CF discards it and the cell shows neither fill', async () => {
    const cells: Cells = { A1: 1, A2: true, B2: 5 }
    const engine = engineEval('IFERROR(COUNTIF(A:A,TRUE),TRUE)', cells, a1('B2'))
    expect(engine.value).toBe(true)
    expect(engine.features).toEqual(['conditional-boolean-criterion-unverified'])
    for (const writesFont of [false, true]) {
      const p1 = formulaRule('IFERROR(COUNTIF(A:A,TRUE),TRUE)', 'B2:B4', { priority: 1, dxf: writesFont ? { ...RED, 'font.italic': true } : RED })
      const batch = await batchOf([p1], fakeSource(cells))
      const outcome = batch.evaluate(0, 1, 1)
      batch.dispose()
      expect(outcome).toEqual({ kind: 'unresolved', reason: 'cf-formula-unsupported' })
      const state = compose(p1, outcome)
      expect(state.properties['fill.bg']).toBeUndefined()
      expect(state.properties['font.bold'], `p1 writes font: ${writesFont}`).toBe(writesFont ? undefined : true)
    }
  })
  it('a cycle flag and a materialization refusal swallowed by IFERROR discard the result the same way', async () => {
    const cells: Cells = { A2: 5 }
    const probes = {
      cycle: (ctx: { hasCycle?: boolean }) => { ctx.hasCycle = true },
      materialization: (ctx: Parameters<typeof chargeMaterialization>[0]) => { chargeMaterialization(ctx, 1001, 1000) },
    }
    for (const [name, probe] of Object.entries(probes)) {
      const rule = formulaRule('IFERROR(A2>0,TRUE)', 'B2:B4', { priority: 1, dxf: RED })
      const batch = await batchOf([rule], fakeSource(cells), { probe: probe as never })
      const outcome = batch.evaluate(0, 1, 1)
      batch.dispose()
      expect(outcome, name).toEqual({ kind: 'unresolved', reason: name === 'cycle' ? 'cf-formula-unsupported' : 'cf-resource-limit' })
      expect(compose(rule, outcome).properties['fill.bg'], name).toBeUndefined()
    }
    // The same rule without a probe is an ordinary positive: the discard is the recording, not the formula.
    expect(valueOf(await cfEval('IFERROR(A2>0,TRUE)', 'B2:B4', cells, 'B2'))).toEqual({ kind: 'boolean', value: true })
  })
})

describe('a read outside the snapshot taints and discards (Review Focus 1)', () => {
  it('a dynamic branch reading a cell outside the footprint is cf-formula-dependency, although its value would be TRUE', async () => {
    const cells: Cells = { A2: 1, B2: 0 }
    const f = compile('IF(A2>0,TRUE,B2)', 'C2')
    expect(engineEval('IF(A2>0,TRUE,B2)', cells, a1('C2')).value).toBe(true)
    const memory = createCfMemory(), signal = new AbortController().signal
    const snap = await prepareCfDependencies([{ rect: rectOf('A2') }], fakeSource(cells), undefined, createCfWorkBudget({ signal, scan: true }), memory, signal, { sheetId: 's1' })
    const snapshot = ready<{ value: Parameters<typeof evaluateCfFormula>[2] }>(snap).value
    evaluateNodeSpy.mockClear()
    const outcome = evaluateCfFormula(f, a1('C2'), snapshot, {
      sheetName: 'Sheet1', evaluation: EVALUATION, semantics: cfFormulaSemantics(EVALUATION), edge: 'ref-error', materialization: { limit: 1e6, used: 0 },
      budget: createCfWorkBudget({ signal, scan: true }), memory, invariants: createCfInvariantMemo(memory), ceiling: createCfUnitCeiling(),
    })
    expect(outcome).toEqual({ kind: 'unresolved', reason: 'cf-formula-dependency' })
    expect(evaluateNodeSpy).not.toHaveBeenCalled()
    // The snapshot itself never serves a missing cell as blank: it taints and reads #N/A.
    snapshot.reset()
    expect(snapshot.read(undefined, 1, 1)).toEqual({ kind: 'formula-error', code: '#N/A' })
    expect(snapshot.tainted).toBe(true)
    snapshot.dispose()
    expect(memory.currentBytes).toBe(0)
  })
  it('an unrepresentable snapshot value taints at read time even when IFERROR masks the read (second line of defence)', async () => {
    // Covered statically (A2 is in the footprint), so only the read-time taint can catch it.
    const masked = await cfEval('IFERROR(A2,TRUE)', 'B2', { A2: { error: '#GETTING_DATA' as never } }, 'B2')
    expect(masked).toEqual({ kind: 'unresolved', reason: 'cf-formula-dependency' })
    expect(await cfEval('A2', 'B2', { A2: { error: '#GETTING_DATA' as never } }, 'B2')).toEqual({ kind: 'unresolved', reason: 'cf-formula-dependency' })
    expect(valueOf(await cfEval('IFERROR(A2,TRUE)', 'B2', { A2: { error: '#N/A' } }, 'B2'))).toEqual({ kind: 'boolean', value: true })
  })
  it('a cached-unavailable cell is unresolved, and a pending source is cf-pending-data', async () => {
    const source = fakeSource({ A2: 1 })
    source.unavailable.add('A2')
    const batch = await batchOf([formulaRule('A2>0', 'B2')], source)
    expect(batch.evaluate(0, 1, 1)).toEqual({ kind: 'unresolved', reason: 'cf-formula-dependency' })
    batch.dispose()
    expect(await cfEval('A2>0', 'B2', { A2: 1 }, 'B2', {}).then(o => o.kind)).toBe('ready')
    const pending = await batchOf([formulaRule('A2>0', 'B2')], fakeSource({ A2: 1 }, { pending: true }))
    expect(pending.evaluate(0, 1, 1)).toEqual({ kind: 'unresolved', reason: 'cf-pending-data' })
    pending.dispose()
  })
  it('a viewFor request not fully covered is refused before any view build or stored() visit', async () => {
    const memory = createCfMemory(), signal = new AbortController().signal
    const snap = await prepareCfDependencies([{ rect: rectOf('A1:A2') }], fakeSource({ A1: 1, A2: 2, A3: 3 }), undefined, createCfWorkBudget({ signal, scan: true }), memory, signal, { sheetId: 's1' })
    const snapshot = ready<{ value: Parameters<typeof evaluateCfFormula>[2] }>(snap).value
    const spend = vi.fn(() => 'continue' as const)
    expect(snapshot.viewFor([{ rect: rectOf('A1:A3') }], { spend, resume() {}, units: 0 })).toEqual({ kind: 'unresolved', reason: 'cf-formula-dependency' })
    expect(spend).not.toHaveBeenCalled()
    snapshot.dispose()
  })
})

describe('SUMIF/AVERAGEIF read the same cells on the view as on the full store (S5-SUMIF-SHAPE regressions)', () => {
  const cells: Cells = { A1: null, A2: null, B1: 0, B2: 7 }
  it('a sum range of another shape is refused at compile, never evaluated on a view', () => {
    for (const text of ['SUMIF(A1:A2,"",B1)=0', 'AVERAGEIF(A1:A2,"",B1)=0', 'SUMIF(A:A,"",B1)=0']) {
      const out = compileCfFormula(text, a1('C1'), createCfMemory())
      expect(isCfFormulaRefusal(out) && out.refused, text).toBe('cf-unsupported-rule')
    }
  })
  it('a same-shape sum range evaluates on the per-call view exactly as the engine does on the full store', async () => {
    for (const text of ['SUMIF(A1:A2,"",B1:B2)', 'AVERAGEIF(A1:A2,"",B1:B2)', 'SUMIF(A:A,"",B:B)', 'SUMIF(A1:A2,"",B1:B2)=0']) {
      const engine = engineEval(text, cells, a1('C1'))
      expect(engine.features, text).toEqual([])
      const result = valueOf(await cfEval(text, 'C1', cells, 'C1', { hoist: false }))
      expect(result.kind === 'range' ? undefined : result.value, text).toEqual(engine.value)
    }
    expect(engineEval('SUMIF(A1:A2,"",B1:B2)', cells, a1('C1')).value).toBe(7)
  })
})

describe('formula results stay finite near ±1e308 (Lesson 6)', () => {
  it('=A2*10>0 with A2 = 1e308 is #NUM! in the engine, so CF never reads it as TRUE', async () => {
    expect(engineEval('A2*10>0', { A2: 1e308 }, a1('B2')).value).toEqual({ kind: 'formula-error', code: '#NUM!' })
    const result = valueOf(await cfEval('A2*10>0', 'B2', { A2: 1e308 }, 'B2'))
    expect(result).toEqual({ kind: 'error', value: { kind: 'formula-error', code: '#NUM!' } })
    for (const readings of [all('boolean', 'number-nonzero-true'), all('boolean', 'error-no-match'), all('boolean', 'error-halts')]) {
      expect(cfTruth(result, a1('B2'), readings)).not.toEqual({ kind: 'ready', value: true })
    }
  })
})

describe('the translated tree reads exactly the call\'s view areas (Task 34 fix 2: view-engine tree guard)', () => {
  const resultOf = (v: unknown) => (typeof v === 'boolean' ? { kind: 'boolean', value: v } : typeof v === 'number' ? { kind: 'number', value: v } : { kind: 'error', value: v })
  it('one end-to-end evaluation per edge mode equals the engine', async () => {
    const cells: Cells = { A1: 5, A1048573: -1, A1048576: 2, A4: 3 }
    // Past the last row: `=A1048576>0` anchored at B1048575, evaluated at B1048576, reads A1048577.
    const past = a1('B1048576')
    const wrap = await cfEval('A1048576>0', 'B1048575:B1048576', cells, 'B1048576', { edge: 'wrap' })
    expect(wrap).toEqual({ kind: 'ready', value: resultOf(engineEval('A1>0', cells, past).value) })
    const refError = await cfEval('A1048576>0', 'B1048575:B1048576', cells, 'B1048576', { edge: 'ref-error' })
    expect(refError).toEqual({ kind: 'ready', value: resultOf(engineEval('#REF!>0', cells, past).value) })
    // No edge evidence: inside the sheet the call needs none.
    const inside = await cfEval('A2>0', 'B1:B3', cells, 'B3', { edge: undefined })
    expect(inside).toEqual({ kind: 'ready', value: resultOf(engineEval('A4>0', cells, a1('B3')).value) })
  })
  it('=SUM(A2>AVERAGE($A:$A))>0 with the hoist substituted paints as the engine does', async () => {
    const cells: Cells = Object.fromEntries(Array.from({ length: 30 }, (_, i) => [`A${i + 1}`, (i * 7) % 11]))
    const text = 'SUM(A2>AVERAGE($A:$A))>0'
    const rule = formulaRule(text, 'B2:B30')
    expect(planCfHoist(formulaOf(rule))).toHaveLength(1)
    const batch = await batchOf([rule], fakeSource(cells))
    let marks = 0
    for (let row = 1; row < 30; row++) {
      const engine = engineEval(text.replace('A2', `A${row + 1}`), cells, { row, col: 1 })
      expect(engine.features).toEqual([])
      const o = batch.evaluate(0, row, 1)
      expect(o, `B${row + 1}`).toEqual({ kind: 'ready', value: resultOf(engine.value) })
      if (engine.value === true) marks++
    }
    expect(marks).toBeGreaterThan(0)
    batch.dispose()
  })
  it('a view that holds every cell a shifted tree could read (no hoist, a pinned range around the target) paints as the engine', async () => {
    // The pinned range puts A1:A10 in the view, so a tree misread by one row would still read a cell the view holds and
    // would not taint: only the guard keeps such a call from painting the wrong value.
    const cells: Cells = { A1: 1, A2: 9, A3: 2, A4: 8, A5: 3, A6: 7, A7: 4, A8: 6, A9: 5, A10: 5 }
    const text = 'A2*10>SUM($A$1:$A$10)'
    const batch = await batchOf([formulaRule(text, 'B2:B9')], fakeSource(cells), { hoist: false })
    for (let row = 1; row < 9; row++) {
      const engine = engineEval(text.replace('A2', `A${row + 1}`), cells, { row, col: 1 })
      expect(batch.evaluate(0, row, 1), `B${row + 1}`).toEqual({ kind: 'ready', value: resultOf(engine.value) })
    }
    batch.dispose()
  })
  it('the guard bites: a tree that reads outside or beside the view areas is refused', () => {
    const f = compile('A2>AVERAGE($A:$A)', 'B2')
    const plan = planCfHoist(f)
    const at = (dRow: number, substitute?: (n: AstNode) => AstNode | undefined) => ready<{ value: AstNode }>(translateCfFormula(f, dRow, 0, 'ref-error', createCfMemory(), substitute)).value
    const hoisted = (n: AstNode): AstNode | undefined => (n === plan[0]!.node ? { type: 'number', value: 4 } : undefined)
    const viewAreas = [rectOf('A5')] // the per-cell shape at B5: A5 only (the hoisted column is not in the view)
    expect(cfTreeReadsMatchView(at(3, hoisted), viewAreas)).toBe(true)
    // Mutant 1: the hoist not substituted, so the engine would traverse $A:$A outside the view.
    expect(cfTreeReadsMatchView(at(3), viewAreas)).toBe(false)
    // Mutant 2: an off-by-one translation reads A6, a cell the view may well hold: only the guard catches it.
    expect(cfTreeReadsMatchView(at(4, hoisted), viewAreas)).toBe(false)
    // Mutant 3: a view missing an area the tree reads.
    expect(cfTreeReadsMatchView(at(3, hoisted), [])).toBe(false)
  })
})
