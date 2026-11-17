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
import { VERIFIED_CAPABILITIES, injectCfFormulaSelections } from '../src/xlsx/cf/evidence'
import {
  formulaBoundsBytes, GREEN, NOW, REGISTERED, firstSelections, RED, UTC_AUTO, cellIsRule, cfBlock, exprRule, fixtureBytes, frameHosts, load, rcOf, residentLeg, staticLeg, textLeg, withPaint, workbookBytes, workerLeg,
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
    const refs = ['C2', 'C3', 'C4', 'C5']
    await withPaint(async () => {
      // Parsed under the evidence too: the rules are resolved, so it is the manual gate that withholds them.
      const l = await load(bytes, 'cached')
      expect(cfDocumentEvaluationOf(l.doc).calcMode).toBe('manual')
      expect(l.sheet.conditionalFormatting!.rules.map(r => r.definition.kind).slice(0, 2)).toEqual(['expression', 'expression'])
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


/* ------------------------------------------------------------------ */
/* One clock per preparation (S5-M).                                    */
/* ------------------------------------------------------------------ */

const REGISTERED_WITHOUT_MANUAL = REGISTERED.filter(id => !id.startsWith('formula:calc-manual-'))
/** The unix day number of the local calendar date: the interactive midnight tick's `clockGen` (`localDaySerial`). */
const unixDay = (epochMs: number, zone: string): number => {
  const p = new Intl.DateTimeFormat('en-CA', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(epochMs)).split('-').map(Number)
  return Math.floor(Date.UTC(p[0]!, p[1]! - 1, p[2]!) / 86_400_000)
}
/** Rows 2-6 hold the date serials of today-2 .. today+2 (UTC date of NOW); C holds `=$A2=TODAY()`, D `=$A2<NOW()`. */
const SERIAL_NOW = Math.floor(NOW / 86_400_000) + 25569
const clockBytes = (): Uint8Array => workbookBytes(
  Object.fromEntries([-2, -1, 0, 1, 2].map((d, i) => [`A${i + 2}`, SERIAL_NOW + d])),
  cfBlock('C2:C6', exprRule('$A2=TODAY()', 1, 0)) + cfBlock('D2:D6', exprRule('$A2&lt;NOW()', 2, 1)))
const CLOCK_REFS = ['C2', 'C3', 'C4', 'C5', 'C6']
const preparedStamps: Array<{ clockGen: number; clock: { epochNowMs: number; clockGen: number } | undefined; dateEpoch: number | undefined }> = []
beforeEach(async () => {
  preparedStamps.length = 0
  const actual = await vi.importActual<typeof import('../src/xlsx/cf/service')>('../src/xlsx/cf/service')
  createService.mockImplementation((rules, memory, options) => {
    const service = actual.createCfService(rules, memory, options)
    const prepareCf = service.prepareCf.bind(service)
    return { ...service, prepareCf: (request, ...rest) => {
      preparedStamps.push({ clockGen: request.stamp.clockGen, clock: options?.evaluation?.clock, dateEpoch: options?.dateContext?.epochNowMs })
      return prepareCf(request, ...rest)
    } }
  })
})

describe('one clock per preparation on every leg (S5-M)', () => {
  for (const zone of ['UTC', 'Asia/Taipei']) {
    it(`static, text, resident and worker give the same TODAY() outcome and the same clockGen (${zone})`, async () => {
      // 20:00 UTC on Nov 16 is already Nov 17 in Taipei: the zone, not the UTC date, decides.
      const now = zone === 'UTC' ? NOW : Date.UTC(2026, 10, 16, 20)
      const previous = process.env.TZ
      process.env.TZ = zone
      try {
        await withPaint(async () => {
          const l = await load(clockBytes(), 'cached')
          const policy = { ...cfDocumentEvaluationOf(l.doc) }
          expect(policy.timeZone).toBe(zone)
          const today = SERIAL_NOW + (zone === 'UTC' ? 0 : 1), expected = [-2, -1, 0, 1, 2].map(d => SERIAL_NOW + d === today ? RED : undefined)
          const options = { now: () => now }
          preparedStamps.length = 0
          const legs = { static: await staticLeg(l, CLOCK_REFS, options), resident: await residentLeg(l, CLOCK_REFS, options), text: await textLeg(l, CLOCK_REFS, options), textResident: await textLeg(l, CLOCK_REFS, { ...options, resident: true }), worker: (await workerLeg(l, CLOCK_REFS, options)).scope }
          for (const [name, leg] of Object.entries(legs)) expect(CLOCK_REFS.map(r => leg.get(r)!.fill), name).toEqual(expected)
          // One clockGen on every leg: the local day number of the injected instant, in the document's zone.
          expect(preparedStamps.length).toBeGreaterThanOrEqual(5)
          for (const p of preparedStamps) {
            expect(p.clock?.epochNowMs).toBe(now)
            expect(p.clock?.clockGen).toBe(unixDay(now, zone))
            expect(p.clockGen).toBe(unixDay(now, zone))
          }
        }, REGISTERED_WITHOUT_MANUAL)
      } finally { if (previous === undefined) delete process.env.TZ; else process.env.TZ = previous }
    })
  }

  it('the clock a preparation hands the predicates is the clock it hands the formulas (one instant)', async () => {
    await withPaint(async () => {
      const l = await load(clockBytes(), 'cached')
      preparedStamps.length = 0
      await staticLeg(l, CLOCK_REFS, { now: () => NOW }); await residentLeg(l, CLOCK_REFS, { now: () => NOW }); await textLeg(l, CLOCK_REFS, { now: () => NOW })
      await workerLeg(l, CLOCK_REFS, { now: () => NOW })
      expect(preparedStamps.length).toBeGreaterThanOrEqual(4)
      for (const p of preparedStamps) expect(p.dateEpoch, 'timePeriod rules read the formula clock').toBe(p.clock?.epochNowMs)
    })
  })

  it('a sheet without a clock-dependent formula keeps its stamp clock-free on every leg', async () => {
    await withPaint(async () => {
      const l = await load(policyBytes(), 'cached')
      preparedStamps.length = 0
      await staticLeg(l, REFS, { now: () => NOW }); await residentLeg(l, REFS, { now: () => NOW }); await textLeg(l, REFS, { now: () => NOW }); await workerLeg(l, REFS, { now: () => NOW })
      expect(preparedStamps.length).toBeGreaterThanOrEqual(4)
      for (const p of preparedStamps) expect(p.clockGen).toBe(0)
    })
  })

  it('NOW() stays refused: its rule never paints on any leg, with every registered capability held', async () => {
    await withPaint(async () => {
      const l = await load(clockBytes(), 'cached')
      const refs = ['D2', 'D3', 'D4', 'D5', 'D6']
      const legs = [await staticLeg(l, refs), await residentLeg(l, refs), await textLeg(l, refs), (await workerLeg(l, refs)).scope]
      for (const leg of legs) expect(refs.map(r => leg.get(r)!.fill)).toEqual(refs.map(() => undefined))
      expect(refs.map(r => legs[2]!.get(r)!.desc)).toEqual(refs.map(() => UNAVAILABLE))
      expect(refs.map(r => legs[3]!.get(r)!.desc)).toEqual(refs.map(() => UNAVAILABLE))
    }, REGISTERED_WITHOUT_MANUAL)
  })
})

describe('manual calculation withholds until measured (S5-L)', () => {
  const SAVED = ['C2', 'C3', 'C4', 'C5'], TODAYS = ['H2', 'H3', 'H4', 'H5']
  const legsOf = async (l: Loaded, refs: string[]) => ({ static: await staticLeg(l, refs), resident: await residentLeg(l, refs), text: await textLeg(l, refs), textResident: await textLeg(l, refs, { resident: true }), worker: (await workerLeg(l, refs)).scope })
  it('without formula:calc-manual-saved-values every leg withholds; with it every leg follows the saved values', async () => {
    // Saved values 999, 0, 50, 1 against `A2>20`: rows 2 and 4. A recompute (10, 8, 2, 61) would paint row 5 instead.
    await withPaint(async () => {
      const l = await load(fixtureBytes('stage5-calc-manual'), 'cached')
      for (const [name, leg] of Object.entries(await legsOf(l, SAVED))) expect(SAVED.map(r => leg.get(r)!.fill), name).toEqual([undefined, undefined, undefined, undefined])
      const text = await textLeg(l, SAVED)
      expect(SAVED.map(r => text.get(r)!.desc)).toEqual(SAVED.map(() => UNAVAILABLE))
    }, REGISTERED_WITHOUT_MANUAL)
    await withPaint(async () => {
      const l = await load(fixtureBytes('stage5-calc-manual'), 'cached')
      const legs = await legsOf(l, SAVED)
      for (const [name, leg] of Object.entries(legs)) {
        const f = SAVED.map(r => leg.get(r)!.fill)
        expect(f[0], `${name} C2`).toBeDefined(); expect(f[1], `${name} C3`).toBeUndefined(); expect(f[2], `${name} C4`).toBe(f[0]); expect(f[3], `${name} C5`).toBeUndefined()
      }
      for (const leg of Object.values(legs)) expect(fills(leg)).toEqual(fills(legs.static))
    })
    // Production evidence: nothing in Stage 5 is verified, so the same cells withhold.
    const l = await load(fixtureBytes('stage5-calc-manual'), 'cached')
    for (const leg of Object.values(await legsOf(l, SAVED))) expect(SAVED.map(r => leg.get(r)!.fill)).toEqual([undefined, undefined, undefined, undefined])
  })
  it('a TODAY() rule under manual calculation stays withheld on every leg even with the saved-values reading and the capability held', async () => {
    await withPaint(async () => {
      const l = await load(fixtureBytes('stage5-calc-manual'), 'cached')
      const legs = await legsOf(l, TODAYS)
      for (const [name, leg] of Object.entries(legs)) expect(TODAYS.map(r => leg.get(r)!.fill), name).toEqual([undefined, undefined, undefined, undefined])
      for (const name of ['text', 'textResident', 'worker'] as const) expect(TODAYS.map(r => legs[name].get(r)!.desc), name).toEqual(TODAYS.map(() => UNAVAILABLE))
    }, [...REGISTERED, 'formula:calc-manual-today'])
  })
})

/* ------------------------------------------------------------------ */
/* Static, print, export, resident, worker and forced-paged agree (Lesson 13). */
/* ------------------------------------------------------------------ */

import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { isCellInSqref } from '../scripts/cf-refs/verify'
import type { CfManifest } from '../scripts/cf-refs/cases'
import { REPO_ROOT } from './helpers/cf-views'

const MATRIX = ['stage5-formula-authored', 'stage5-formula-functions', 'stage5-predicate-formulas', 'stage5-cellis-operands', 'stage5-composition', 'stage5-cross-sheet']
const manifest = (): CfManifest => JSON.parse(readFileSync(join(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
const markedCells = (id: string): string[] => {
  const c = manifest().cases.find(x => x.id === id)!
  return c.expected.cells.map(e => e.ref).filter(r => c.markers.some(m => m.formula !== undefined && isCellInSqref(r, m.sqref)))
}
const paintToken = (o: Obs): string => `${o.fill ?? '-'}${o.strike ? '+strike' : ''}${o.fmt ? `|${o.fmt}` : ''}`
const tokens = (m: ReadonlyMap<string, Obs>, f: (o: Obs) => string = paintToken): string[] => [...m.entries()].map(([k, v]) => `${k}:${f(v)}`)

describe('static, print, export, resident, worker and forced-paged agree for paint, a11y and copy/find (Lesson 13)', () => {
  it('every leg shows each Stage 5 fixture cell identically, and the cross-sheet fixture reads its other sheet on all of them', async () => {
    let painted = 0, described = 0, compared = 0
    for (const id of MATRIX) {
      await withPaint(async () => {
        const l = await load(fixtureBytes(id), 'cached', manifest().cases.find(x => x.id === id)!.sheet)
        const refs = markedCells(id)
        // The same frame hosts on every scope, so bars and icons are decided on one geometry (not on the legs' own).
        const hosts = frameHosts()
        const worker = await workerLeg(l, refs, { hosts })
        expect(worker.wire.kind, id).toBe('ready')
        const paint = { static: await staticLeg(l, refs), resident: await residentLeg(l, refs), text: await textLeg(l, refs, { hosts }), textResident: await textLeg(l, refs, { resident: true, hosts }), forcedPaged: worker.scope }
        const reference = tokens(paint.static)
        for (const [name, leg] of Object.entries(paint)) expect(tokens(leg), `${id}: ${name} paints what static paints`).toEqual(reference)
        const announced = { text: tokens(paint.text, o => o.desc ?? '-'), textResident: tokens(paint.textResident, o => o.desc ?? '-'), forcedPaged: tokens(paint.forcedPaged, o => o.desc ?? '-') }
        expect(announced.textResident, `${id}: resident announces what text announces`).toEqual(announced.text)
        expect(announced.forcedPaged, `${id}: forced-paged announces what text announces`).toEqual(announced.text)
        // The raw worker wire carries the same fills the scopes derive.
        for (const r of refs) { const st = worker.states.get(r); if (st?.properties['fill.bg']) expect(paint.forcedPaged.get(r)!.fill).toBe(String(st.properties['fill.bg']).toUpperCase()) }
        painted += reference.filter(t => !t.endsWith(':-')).length
        described += announced.text.filter(t => !t.endsWith(':-')).length
        compared += refs.length * 5
      })
    }
    expect(compared).toBeGreaterThan(200)
    expect(painted, 'real paint is compared, not only withholding').toBeGreaterThan(20)
    expect(described).toBeGreaterThan(20)
  })

  it('under production evidence every leg withholds identically, and withholds what the evidenced run paints (no Stage 5 capability is verified)', async () => {
    let withheldPaint = 0, evidencedPaint = 0
    for (const id of MATRIX) {
      const sheetName = manifest().cases.find(x => x.id === id)!.sheet
      const refs = markedCells(id), hosts = frameHosts()
      const production = await load(fixtureBytes(id), 'cached', sheetName)
      const legs = [await staticLeg(production, refs), await residentLeg(production, refs), await textLeg(production, refs, { hosts }), await textLeg(production, refs, { resident: true, hosts }), (await workerLeg(production, refs, { hosts })).scope]
      for (const leg of legs) expect(tokens(leg), id).toEqual(tokens(legs[0]!))
      expect(tokens(legs[3]!, o => o.desc ?? '-'), id).toEqual(tokens(legs[2]!, o => o.desc ?? '-'))
      expect(tokens(legs[4]!, o => o.desc ?? '-'), id).toEqual(tokens(legs[2]!, o => o.desc ?? '-'))
      withheldPaint += tokens(legs[0]!).filter(t => !t.endsWith(':-')).length
      await withPaint(async () => { evidencedPaint += tokens(await staticLeg(await load(fixtureBytes(id), 'cached', sheetName), refs)).filter(t => !t.endsWith(':-')).length })
    }
    expect(evidencedPaint, 'the evidence is what adds paint').toBeGreaterThan(withheldPaint + 20)
  })

  it('a leg without the workbook source withholds the cross-sheet cells identically on paint, a11y and copy/find', async () => {
    const id = 'stage5-cross-sheet'
    await withPaint(async () => {
      const l = await load(fixtureBytes(id), 'cached', manifest().cases.find(x => x.id === id)!.sheet)
      const refs = markedCells(id)
      const withBook = await staticLeg(l, refs)
      expect(tokens(withBook).some(t => !t.endsWith(':-')), 'with the workbook the cross-sheet rules paint').toBe(true)
      const bare = { static: await staticLeg(l, refs, { workbook: false }), resident: await residentLeg(l, refs, { workbook: false }), text: await textLeg(l, refs, { workbook: false }), textResident: await textLeg(l, refs, { workbook: false, resident: true }) }
      for (const [name, leg] of Object.entries(bare)) expect(tokens(leg).every(t => t.endsWith(':-')), `${name}: paints nothing`).toBe(true)
      expect(tokens(bare.textResident, o => o.desc ?? '-')).toEqual(tokens(bare.text, o => o.desc ?? '-'))
      expect(tokens(bare.text, o => o.desc ?? '-').some(t => t.endsWith(UNAVAILABLE)), 'a11y names the withheld cells').toBe(true)
    })
  })
})

describe('formula cfvo bounds announce and paint the same on every scope, forced-paged included (Task 37 follow-up)', () => {
  it('a data bar, a colour scale and an icon set bounded by formulas are described identically by text, resident and forced-paged scopes', async () => {
    const { withIcons } = await import('./helpers/cf-icons')
    await withIcons(async () => {
      injectCfFormulaSelections(firstSelections())
      try {
        const l = await load(formulaBoundsBytes(), 'cached')
        const refs = ['B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'C2', 'C5', 'C8', 'D2', 'D5', 'D8'], hosts = frameHosts()
        const text = await textLeg(l, refs, { hosts }), resident = await textLeg(l, refs, { hosts, resident: true }), forced = await workerLeg(l, refs, { hosts })
        expect(tokens(resident, o => o.desc ?? '-')).toEqual(tokens(text, o => o.desc ?? '-'))
        expect(tokens(forced.scope, o => o.desc ?? '-')).toEqual(tokens(text, o => o.desc ?? '-'))
        expect(tokens(forced.scope)).toEqual(tokens(text))
        const all = tokens(text, o => o.desc ?? '-').join('\n')
        expect(all, 'a formula-bounded bar announces its position').toMatch(/B8:cf:dataBar\|Data bar, 100% of range/)
        expect(all, 'a formula-bounded scale announces its position').toMatch(/C8:cf:colorScale\|Colour scale, 100% of range/)
        // The icon's bucket comes from the formula threshold on the wire the forced-paged leg reads (D2 = 0, D5 = 40 under $F$1 = 50;
        // D8 = 100 above it). Whether its frame paints is the icon host tests' business (S4-ROW-EXACT, glyph table): here every
        // scope agrees on it, which the equalities above pin.
        const buckets = (forced.wire as unknown as { states: Array<{ row: number; col: number; icon?: { bucket: number; buckets: number } }> }).states.filter(x => x.col === 3).map(x => `${x.row}:${x.icon?.bucket}/${x.icon?.buckets}`)
        expect(buckets).toEqual(['1:0/3', '4:0/3', '7:2/3'])
      } finally { injectCfFormulaSelections(null) }
    }, { extra: [...VERIFIED_CAPABILITIES, ...REGISTERED, 'bar:participation-dateserial'] })
  })
  it('under production evidence every scope announces the same unavailable formatting, and none paints', async () => {
    const l = await load(formulaBoundsBytes(), 'cached')
    const refs = ['B4', 'C4', 'D4'], hosts = frameHosts()
    const text = await textLeg(l, refs, { hosts }), forced = await workerLeg(l, refs, { hosts })
    expect(tokens(forced.scope, o => o.desc ?? '-')).toEqual(tokens(text, o => o.desc ?? '-'))
    expect(tokens(text, o => o.desc ?? '-').every(t => t.endsWith(UNAVAILABLE) || t.endsWith(':-'))).toBe(true)
  })
})

describe('formula descriptors carry no formula text (S5-Q, Lesson 10)', () => {
  /** C holds `=$A2>1` (rows 3-4 match) and E the literal twin `>1`, both with the same DXF; A and E hold 1, 2, 3. */
  const twinBytes = (dxf: number): Uint8Array => workbookBytes(
    { A2: 1, A3: 2, A4: 3, C2: 1, C3: 2, C4: 3, E2: 1, E3: 2, E4: 3 },
    cfBlock('C2:C4', exprRule('$A2&gt;1', 1, dxf)) + cfBlock('E2:E4', cellIsRule('greaterThan', '1', 2, dxf)))
  const FORMULA_TEXT = /[=$!]|[A-Z]{2,}\(|\b[A-Z]{1,3}\d+\b|\bformula\b|TODAY|COUNTIF|ISBLANK/
  for (const dxf of [0, 2]) {
    it(`a matching expression announces exactly what its Stage 1 twin announces, on every scope (dxf ${dxf})`, async () => {
      await withPaint(async () => {
        const l = await load(twinBytes(dxf), 'cached')
        const refs = ['C2', 'C3', 'C4', 'E2', 'E3', 'E4'], hosts = frameHosts()
        const worker = await workerLeg(l, refs, { hosts })
        const scopes = [await textLeg(l, refs, { hosts }), await textLeg(l, refs, { hosts, resident: true }), worker.scope]
        for (const scope of scopes) {
          for (const [formula, literal] of [['C2', 'E2'], ['C3', 'E3'], ['C4', 'E4']] as const) {
            expect(scope.get(formula), `${formula} equals ${literal}`).toEqual(scope.get(literal))
          }
          expect(scope.get('C3')!.desc, 'the match is announced').toBeDefined()
          expect(scope.get('C2')!.desc, 'no match, nothing announced').toBeUndefined()
          for (const r of refs) { const d = scope.get(r)!.desc; if (d) { expect(d.split('|')[1]!.length).toBeLessThanOrEqual(128); expect(d, r).not.toMatch(FORMULA_TEXT) } }
        }
        // Copy and find follow the effective number format exactly as Stage 1 does.
        if (dxf === 2) for (const scope of scopes) { expect(scope.get('C3')!.fmt).toBe('0.00'); expect(scope.get('E3')!.fmt).toBe('0.00'); expect(scope.get('C2')!.fmt).toBeUndefined() }
      })
    })
  }
})

describe('blank cells match formula rules in every view (Stage 1 B2/Z5 pattern, a CF-only sheet)', () => {
  const blankBytes = (): Uint8Array => workbookBytes({}, cfBlock('B2', exprRule('ISBLANK(B2)', 1, 1)) + cfBlock('Z5', exprRule('ISBLANK(Z5)', 2, 1)))
  it('=ISBLANK(B2) over unstored cells paints on the static, resident and text legs, and reaches the worker and its forced-paged scope', async () => {
    await withPaint(async () => {
      const l = await load(blankBytes(), 'cached')
      const refs = ['B2', 'Z5']
      const worker = await workerLeg(l, refs)
      const legs = [await staticLeg(l, refs), await residentLeg(l, refs), await textLeg(l, refs), await textLeg(l, refs, { resident: true }), worker.scope]
      for (const leg of legs) expect(refs.map(r => leg.get(r)!.fill), 'unstored B2 and Z5 paint').toEqual([GREEN, GREEN])
      expect(worker.states.get('B2')!.properties['fill.bg']).toBe(GREEN)
      expect(worker.states.get('Z5')!.properties['fill.bg']).toBe(GREEN)
    })
  })
  it('the same rule is withheld identically under production evidence', async () => {
    const l = await load(blankBytes(), 'cached')
    const refs = ['B2', 'Z5']
    const worker = await workerLeg(l, refs)
    for (const leg of [await staticLeg(l, refs), await textLeg(l, refs), worker.scope]) expect(refs.map(r => leg.get(r)!.fill)).toEqual([undefined, undefined])
    expect(refs.map(r => worker.scope.get(r)!.desc)).toEqual([UNAVAILABLE, UNAVAILABLE])
  })
})
void NOW; void RED; void frameHosts; void rcOf; void ({} as Loaded); void ({} as CfDocumentEvaluation); void GREEN; void cfDocumentEvaluationOfPaged
