// @vitest-environment node
/**
 * Task 38 (Rulings S5-L, S5-M, S5-P, S5-Q; Lesson 13): every view takes the same calculation policy, the same clock and the
 * same workbook, and answers the same for paint, accessibility and copy/find, or withholds identically. Each leg prepares the
 * same workbook through its own production route (tests/helpers/cf-views.ts). Stage 5 capabilities are unverified (no Excel
 * capture yet), so real paint is exercised under `withPaint` (additive injected evidence, never a fabricated observation) and
 * identical withholding is asserted again under production evidence.
 */
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import * as serviceModule from '../src/xlsx/cf/service'
import * as composeModule from '../src/xlsx/cf/compose'
import * as workbookModule from '../src/xlsx/formula/workbook'
import { cfDocumentEvaluationOf, cfDocumentEvaluationOfPaged, type CfDocumentEvaluation } from '../src/xlsx/cf/formula'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import { residentSource } from '../src/xlsx/cf/render'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import {
  GREEN, NOW, RED, UTC_AUTO, cellIsRule, cfBlock, exprRule, fixtureBytes, frameHosts, load, rcOf, residentLeg, staticLeg, textLeg, withPaint, workbookBytes, workerLeg,
  type Loaded, type Obs,
} from './helpers/cf-views'

vi.mock('../src/xlsx/cf/service', async importOriginal => {
  const original = await importOriginal<typeof import('../src/xlsx/cf/service')>()
  return { ...original, createCfService: vi.fn(original.createCfService) }
})
vi.mock('../src/xlsx/cf/compose', async importOriginal => {
  const original = await importOriginal<typeof import('../src/xlsx/cf/compose')>()
  return { ...original, composeCellAsync: vi.fn(original.composeCellAsync) }
})
vi.mock('../src/xlsx/formula/workbook', async importOriginal => {
  const original = await importOriginal<typeof import('../src/xlsx/formula/workbook')>()
  return { ...original, recomputeDependents: vi.fn(original.recomputeDependents) }
})
const createService = vi.mocked(serviceModule.createCfService)
const recompute = vi.mocked(workbookModule.recomputeDependents)
const compose = vi.mocked(composeModule.composeCellAsync)
beforeEach(() => { createService.mockClear(); recompute.mockClear(); compose.mockClear() })
/** The outcome the service handed the compositor for the rule of `priority`, one per prepared cell. */
const outcomesOf = (priority: number): unknown[] => compose.mock.calls.flatMap(([rules, outcomes]) => rules.flatMap((rule, i) => rule.priority === priority ? [outcomes[i]] : []))

// The paged leg reads the ambient zone once per workbook: pin it so the legs share one zone on every machine.
const previousTz = process.env.TZ
beforeAll(() => { process.env.TZ = 'UTC' })
afterAll(() => { if (previousTz === undefined) delete process.env.TZ; else process.env.TZ = previousTz })

const fills = (m: ReadonlyMap<string, Obs>): Record<string, string | undefined> => Object.fromEntries([...m].map(([k, v]) => [k, v.fill]))
const descs = (m: ReadonlyMap<string, Obs>): Record<string, string | undefined> => Object.fromEntries([...m].map(([k, v]) => [k, v.desc]))
const UNAVAILABLE = 'cf-unavailable|Conditional formatting unavailable'
const FORMULA_CELLS = ['C2', 'C3', 'C4', 'C5', 'C6']
const CONTROL_CELLS = ['A2', 'A3', 'A4', 'A5', 'A6']
const REFS = [...CONTROL_CELLS, ...FORMULA_CELLS]

/** A2:A6 = 10..50 under a Stage 1 cellIs (>25 green); B2:B6 = `A*2` formulas with NO cached value; C2:C6 holds the formula
 * rule `$B2>50` (red), whose footprint is the uncached column: true on rows 4-6 where the engine value is known. */
const policyBytes = (): Uint8Array => workbookBytes(
  { A2: 10, A3: 20, A4: 30, A5: 40, A6: 50, B2: { f: 'A2*2' }, B3: { f: 'A3*2' }, B4: { f: 'A4*2' }, B5: { f: 'A5*2' }, B6: { f: 'A6*2' } },
  cfBlock('C2:C6', exprRule('$B2&gt;50', 1, 0)) + cfBlock('A2:A6', cellIsRule('greaterThan', '25', 2, 1)))

describe('every route takes one evaluation context or withholds (S5-L, round-1 #5)', () => {
  it('under the cached policy every leg withholds the uncached footprint identically, and Stage 1-4 paint is unchanged', async () => {
    await withPaint(async () => {
      const l = await load(policyBytes(), 'cached')
      const paintLegs = [await staticLeg(l, REFS), await residentLeg(l, REFS)]
      const scopeLegs = [await textLeg(l, REFS), await textLeg(l, REFS, { resident: true }), (await workerLeg(l, REFS)).scope]
      for (const leg of [...paintLegs, ...scopeLegs]) {
        expect(Object.values(fills(leg)).slice(5), 'no formula cell paints').toEqual([undefined, undefined, undefined, undefined, undefined])
        expect(fills(leg).A4, 'the Stage 1 control paints').toBe(GREEN)
      }
      for (const leg of scopeLegs) {
        expect(FORMULA_CELLS.map(r => leg.get(r)!.desc), 'every scope announces the formula cells as unavailable').toEqual(FORMULA_CELLS.map(() => UNAVAILABLE))
        const applied = 'cf:fill|Conditional formatting applied: fill'
        expect(CONTROL_CELLS.map(r => leg.get(r)!.desc)).toEqual([undefined, undefined, applied, applied, applied])
      }
      expect(descs(scopeLegs[0]!)).toEqual(descs(scopeLegs[1]!))
      expect(descs(scopeLegs[0]!)).toEqual(descs(scopeLegs[2]!))
    })
  })

  it('under the engine policy the static, resident and text legs evaluate the same footprint identically', async () => {
    await withPaint(async () => {
      const l = await load(policyBytes(), 'engine')
      const legs = [await staticLeg(l, REFS), await residentLeg(l, REFS), await textLeg(l, REFS), await textLeg(l, REFS, { resident: true })]
      for (const leg of legs) expect(FORMULA_CELLS.map(r => leg.get(r)!.fill)).toEqual([undefined, undefined, RED, RED, RED])
      for (const leg of legs.slice(1)) expect(fills(leg)).toEqual(fills(legs[0]!))
      expect(descs(legs[2]!)).toEqual(descs(legs[3]!))
    })
  })

  it('a route called without the context leaves every formula rule unresolved and Stage 1-4 rules unchanged', async () => {
    await withPaint(async () => {
      const l = await load(policyBytes(), 'engine')
      const plain = await staticLeg(l, REFS, { evaluation: UTC_AUTO })
      const legs = [await staticLeg(l, REFS, { evaluation: false }), await residentLeg(l, REFS, { evaluation: false }), await textLeg(l, REFS, { evaluation: false }), await textLeg(l, REFS, { evaluation: false, resident: true })]
      expect(FORMULA_CELLS.map(r => plain.get(r)!.fill), 'the policy-carrying twin evaluates').toEqual([undefined, undefined, RED, RED, RED])
      for (const leg of legs) {
        expect(FORMULA_CELLS.map(r => leg.get(r)!.fill)).toEqual(FORMULA_CELLS.map(() => undefined))
        expect(CONTROL_CELLS.map(r => leg.get(r)!.fill)).toEqual(CONTROL_CELLS.map(r => plain.get(r)!.fill))
      }
      for (const leg of legs.slice(2)) expect(FORMULA_CELLS.map(r => leg.get(r)!.desc)).toEqual(FORMULA_CELLS.map(() => UNAVAILABLE))
      // The worker's own route, handed no context, names the reason on the wire.
      const resident = await load(policyBytes(), 'cached')
      const rules = resident.sheet.conditionalFormatting!, memory = createCfMemory()
      const stamp = { owner: 'ctx', sheetId: resident.sheetId, workbookValueGen: 0, cfGen: cfRulesVersion(rules), viewGen: 0, structureGen: 0, clockGen: 0 }
      compose.mockClear()
      const wire = await prepareCfRemote({ rects: [{ r0: 1, r1: 5, c0: 2, c1: 2 }], stamp, grantBytes: 1 << 22 }, resident.sheetId, residentSource(resident.sheet, stamp, memory, 'cached'), rules, memory, {})
      expect(wire.kind).toBe('ready')
      if (wire.kind === 'ready') expect(wire.states.map(s => s.unresolvedWrites)).toEqual(FORMULA_CELLS.map(() => ['fill']))
      expect(outcomesOf(1)).toEqual(FORMULA_CELLS.map(() => ({ kind: 'unresolved', reason: 'cf-evaluation-context-missing' })))
    })
  })

  it('a route whose own value policy disagrees with the document policy withholds like one without the context', async () => {
    await withPaint(async () => {
      const l = await load(policyBytes(), 'engine')
      for (const leg of [await staticLeg(l, REFS, { valuePolicy: 'cached' }), await residentLeg(l, REFS, { valuePolicy: 'cached' }), await textLeg(l, REFS, { valuePolicy: 'cached' })]) {
        expect(FORMULA_CELLS.map(r => leg.get(r)!.fill)).toEqual(FORMULA_CELLS.map(() => undefined))
      }
    })
  })

  it('production evidence withholds every formula cell identically on every leg (nothing in Stage 5 is verified)', async () => {
    const l = await load(policyBytes(), 'cached')
    const legs = [await staticLeg(l, REFS), await residentLeg(l, REFS), await textLeg(l, REFS), await textLeg(l, REFS, { resident: true }), (await workerLeg(l, REFS)).scope]
    for (const leg of legs) expect(FORMULA_CELLS.map(r => leg.get(r)!.fill)).toEqual(FORMULA_CELLS.map(() => undefined))
    for (const leg of legs.slice(2)) expect(FORMULA_CELLS.map(r => leg.get(r)!.desc)).toEqual(FORMULA_CELLS.map(() => UNAVAILABLE))
  })

  it('a paged manual-mode workbook reports calcMode manual and withholds exactly as the resident one does', async () => {
    const bytes = fixtureBytes('stage5-calc-manual')
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'manual-paged', profile: 'D' })
    try {
      await session.done
      expect(session.workbook.calcMode).toBe('manual')
      expect(cfDocumentEvaluationOfPaged(session.workbook).calcMode).toBe('manual')
    } finally { await session.release() }
    const l = await load(bytes, 'cached')
    expect(cfDocumentEvaluationOf(l.doc).calcMode).toBe('manual')
    const refs = ['C2', 'C3', 'C4', 'C5']
    await withPaint(async () => {
      const resident = await residentLeg(l, refs), worker = await workerLeg(l, refs)
      expect(fills(resident)).toEqual({ C2: undefined, C3: undefined, C4: undefined, C5: undefined })
      expect(fills(worker.scope)).toEqual(fills(resident))
      expect(refs.map(r => worker.scope.get(r)!.desc)).toEqual(refs.map(() => UNAVAILABLE))
    }, REGISTERED_WITHOUT_MANUAL)
  })

  it('no route ever recomputes dependents', async () => {
    await withPaint(async () => {
      const l = await load(policyBytes(), 'engine')
      await staticLeg(l, REFS); await residentLeg(l, REFS); await textLeg(l, REFS, { resident: true })
      await workerLeg(await load(policyBytes(), 'cached'), REFS)
    })
    expect(recompute).not.toHaveBeenCalled()
  })
})

import { REGISTERED } from './helpers/cf-views'
const REGISTERED_WITHOUT_MANUAL = REGISTERED.filter(id => !id.startsWith('formula:calc-manual-'))
void NOW; void RED; void frameHosts; void rcOf; void createService; void ({} as Loaded) ; void ({} as CfDocumentEvaluation)
