/**
 * PR #74 bot fix (S3-P): a painted data bar announces the VALUE's position between the resolved min and max thresholds
 * (0 to 100), never the bar's painted endpoint. Default 10/90 lengths and an asymmetric axis shift the geometry, not the
 * value: the descriptor carries `valueFraction` beside its geometry, and the announcement and its cache key use it.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { prepareCfTextScope, prepareCfWireScope } from '../src/xlsx/cf/text'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { injectCfBarLengths } from '../src/xlsx/cf/evidence'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import type { CfStamp } from '../src/xlsx/cf/model'
import type { CfReadResult } from '../src/xlsx/view/contracts'
import type { XlsxSheet } from '../src/xlsx/types'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import { barSheet, paintScopeBars, withBars } from './helpers/cf-bars'

const budget = (): FormatBudget => ({ reserve() {}, release() {}, retain: (_t, _b, commit) => commit(), settle() {} })
const scopeOf = (sheet: XlsxSheet, rows: number, locale: string) => prepareCfTextScope(sheet, {
  rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }], signal: new AbortController().signal, locale, budget: budget(),
})
const withLocale = (sheet: XlsxSheet, locale: string): XlsxSheet => ({ ...sheet, semantics: { locale, dateSystem: '1900' } } as unknown as XlsxSheet)
const X14_DEFAULT = { minLength: 10, maxLength: 90, pairing: 'x14' as const }
const ALL_FORMS = ['base-only/gradient', 'base-only/solid', 'x14/gradient', 'x14/solid']
const HALF_MIN_X14 = { 'default-10-90': { rule: 'half-min@x14', measured: ALL_FORMS, settledBy: ['bar:lengths-default-10-90'] } }
const EXTRA = ['bar:lengths-default-10-90']
const UNAVAILABLE = { key: 'cf-unavailable', text: '無法使用條件式格式設定' }

afterEach(() => injectCfBarLengths(null))

const wireOf = async (sheet: XlsxSheet, rows: number): Promise<CfReadResult> => {
  const rules = sheet.conditionalFormatting!, sheetId = sheet.sheetId!
  const st: CfStamp = { owner: 'paged', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
  const memory = createCfMemory({})
  return prepareCfRemote({ rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }], stamp: { ...st }, grantBytes: 1 << 22 }, sheetId, residentSource(sheet, st, memory, 'engine'), rules, memory, { budget: () => createCfWorkBudget({ now: () => 0 }) }) as Promise<CfReadResult>
}
const wireScope = (wire: CfReadResult, rows: number, locale: string) => prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }], { budget: { reserve() {}, release() {} }, locale })

describe('a bar announces the value position, not the painted endpoint (PR #74 P2)', () => {
  it('x14 default 10/90 lengths over 0, 50, 100 announce 0%, 50%, 100% in en and zh-TW', async () => {
    injectCfBarLengths(HALF_MIN_X14)
    await withBars(async () => {
      for (const [locale, text] of [['en-US', (n: number) => `Data bar, ${n}% of range`], ['zh-TW', (n: number) => `資料橫條 ${n}%`]] as const) {
        const scope = (await scopeOf(withLocale(barSheet([0, 50, 100], { def: X14_DEFAULT }), locale), 3, locale))!
        expect([0, 1, 2].map(r => scope.describe(r, 0)?.text), locale).toEqual([text(0), text(50), text(100)])
        // The painted geometry is the Excel-verified one (0.05, 0.45, 0.85): only the announcement is the value position.
        expect([0, 1, 2].map(r => Number(scope.barAt!(r, 0)!.endFraction.toFixed(12))), locale).toEqual([0.05, 0.45, 0.85])
        expect([0, 1, 2].map(r => scope.barAt!(r, 0)!.valueFraction), locale).toEqual([0, 0.5, 1])
        scope.dispose()
      }
    }, { extra: EXTRA })
  })

  it('a mixed-sign bar with asymmetric bounds announces the position between min and max (axis value reads 25%, not 0%)', async () => {
    await withBars(async () => {
      const scope = (await scopeOf(barSheet([-25, 0, 75]), 3, 'en-US'))!
      expect([0, 1, 2].map(r => scope.describe(r, 0)?.text)).toEqual(['Data bar, 0% of range', 'Data bar, 25% of range', 'Data bar, 100% of range'])
      expect([0, 1, 2].map(r => scope.barAt!(r, 0)!.valueFraction)).toEqual([0, 0.25, 1])
      scope.dispose()
    })
  })

  it('the descriptor cache keys by the value position: equal geometry with different values never shares text, and equal values share', async () => {
    await withBars(async () => {
      const sheet = barSheet([0, 50, 100])
      const wire = await wireOf(sheet, 3)
      const states = wire.states.filter(s => s.bar) as { bar?: NonNullable<CfReadResult['states'][number]['bar']> }[]
      expect(states.length).toBe(3)
      // Cells 1 and 2 get the SAME painted geometry but different values; cell 0 gets cell 1's value with other geometry.
      const geometry = { ...states[1]!.bar!, valueFraction: 0.2 }
      states[1]!.bar = geometry
      states[2]!.bar = { ...geometry, valueFraction: 0.8 }
      states[0]!.bar = { ...states[0]!.bar!, valueFraction: 0.2 }
      const scope = (await wireScope(wire, 3, 'en-US'))!
      await paintScopeBars(scope, 3, 1)
      expect([0, 1, 2].map(r => scope.describe(r, 0)?.text)).toEqual(['Data bar, 20% of range', 'Data bar, 20% of range', 'Data bar, 80% of range'])
      scope.dispose()
    })
  })

  it('a malformed valueFraction on the wire fails closed: no bar is reported or announced', async () => {
    await withBars(async () => {
      const wire = await wireOf(barSheet([0, 50, 100]), 3)
      const bad: unknown[] = [Number.NaN, Infinity, -0.01, 1.01, '0.5', undefined, null]
      for (const value of bad) {
        const mutated = structuredClone(wire)
        const cell = mutated.states.find(s => s.bar && s.row === 1)!
        ;(cell.bar as unknown as Record<string, unknown>).valueFraction = value
        const scope = (await wireScope(mutated, 3, 'zh-TW'))!
        await paintScopeBars(scope, 3, 1)
        expect(scope.barAt?.(1, 0), String(value)).toBeUndefined()
        expect(scope.describe(1, 0), String(value)).toEqual(UNAVAILABLE)
        expect(scope.describe(2, 0)?.key, `${String(value)}: the untouched bar of the same page`).toBe('cf:dataBar')
        scope.dispose()
      }
    })
  })
})
