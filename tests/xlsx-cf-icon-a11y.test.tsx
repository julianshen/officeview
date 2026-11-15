/**
 * Task 29, Ruling S4-N: a painted icon announces a bounded positional label (OD-10: 「圖示 b/k」 / "Icon b of k"), derived
 * from the descriptor's bucket semantics, never from paint geometry, glyph data or the glyph class. The value text stays
 * in the accessible name, copy and find, and the icon adds no tab stop. Icons paint here under INJECTED stubs only
 * (tests/helpers/cf-icons.ts); production paints none yet (S4-ORDER).
 */
import { describe, expect, it } from 'vitest'
import { render } from '@testing-library/react'
import { createCanvas } from 'canvas'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { SemanticPresentationCache, withCfDescription } from '../src/components/xlsx/formatPresentation'
import { prepareCfTextScope, prepareCfWireScope, wireIcon } from '../src/xlsx/cf/text'
import { describeCfWin } from '../src/xlsx/cf/describe'
import { CF_LIMITS, createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import type { CfStamp } from '../src/xlsx/cf/model'
import type { CfReadResult } from '../src/xlsx/view/contracts'
import type { XlsxSheet } from '../src/xlsx/types'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import { gridFixture } from './helpers/xlsx-grid-fixture'
import { iconSheet, withIcons } from './helpers/cf-icons'
import { barRule } from './helpers/cf-bars'

const UNAVAILABLE_EN = { key: 'cf-unavailable', text: 'Conditional formatting unavailable' }
const layoutContext = () => createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
function trackingBudget(): { budget: FormatBudget; settled: object[] } {
  const settled: object[] = []
  return { budget: { reserve() {}, release() {}, retain: (_t, _b, commit) => commit(), settle: target => { settled.push(target) } }, settled }
}
const scopeOf = (sheet: XlsxSheet, rows: number, locale = 'zh-TW') => prepareCfTextScope(sheet, {
  rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }], signal: new AbortController().signal, locale, budget: trackingBudget().budget, layoutContext,
})
const wireOf = async (sheet: XlsxSheet, rows: number): Promise<CfReadResult> => {
  const rules = sheet.conditionalFormatting!, sheetId = sheet.sheetId!
  const st: CfStamp = { owner: 'paged', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
  const memory = createCfMemory({})
  const wire = await prepareCfRemote({ rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }], stamp: { ...st }, grantBytes: 1 << 22 }, sheetId, residentSource(sheet, st, memory, 'engine'), rules, memory, { budget: () => createCfWorkBudget({ now: () => 0 }) })
  expect(wire.kind).toBe('ready')
  return wire as CfReadResult
}

describe('icon cell announces a bounded positional label, without a tab stop (S4-N, Lesson 10)', () => {
  it('describeCfWin labels bucket 1 of 3 positionally in zh-TW and en, whatever the glyph class or reverse', () => {
    for (const icon of [{ bucket: 1, buckets: 3 }, { bucket: 1, buckets: 3, set: '3Flags', iconId: 2 }] as const) {
      expect(describeCfWin({}, undefined, undefined, 'zh-TW', undefined, undefined, icon)).toEqual({ key: 'cf:iconSet', text: '圖示 2/3' })
      expect(describeCfWin({}, undefined, undefined, 'en-US', undefined, undefined, icon)).toEqual({ key: 'cf:iconSet', text: 'Icon 2 of 3' })
    }
    // Bar + icon: one bounded label, the bar's then the icon's, by the locale joiner.
    expect(describeCfWin({}, undefined, undefined, 'zh-TW', undefined, { fraction: 0.72 }, { bucket: 1, buckets: 3 })).toEqual({ key: 'cf:dataBar+iconSet', text: '資料橫條 72%、圖示 2/3' })
    expect(describeCfWin({}, undefined, undefined, 'en-US', undefined, { fraction: 0.72 }, { bucket: 1, buckets: 3 })).toEqual({ key: 'cf:dataBar+iconSet', text: 'Data bar, 72% of range, Icon 2 of 3' })
    for (const locale of ['zh-TW', 'en-US']) for (const b of [0, 1, 2, 3, 4]) {
      const d = describeCfWin({}, undefined, undefined, locale, undefined, { fraction: 1 }, { bucket: b, buckets: 5 })!
      expect(d.text.length).toBeLessThanOrEqual(CF_LIMITS.descriptorUnits)
      expect(d.text).not.toMatch(/#|rgb|Arrows|Flags|ff|min|max|%.*%/i)
    }
    // An unresolved icon announces only the unavailable descriptor; a malformed bucket adds nothing.
    expect(describeCfWin({}, new Set(['iconSet']), undefined, 'en-US', undefined, undefined, { bucket: 1, buckets: 3 })).toEqual(UNAVAILABLE_EN)
    expect(describeCfWin({}, undefined, undefined, 'en-US', undefined, undefined, { bucket: 3, buckets: 3 })).toBeUndefined()
  })

  it('a prepared scope announces each icon at its own bucket; NoIcons adds nothing; the descriptor cache is keyed by the icon', async () => {
    await withIcons(async () => {
      const zh = (await scopeOf(iconSheet([10, 50, 90]), 3))!
      expect([0, 1, 2].map(r => zh.describe(r, 0))).toEqual([
        { key: 'cf:iconSet', text: '圖示 1/3' }, { key: 'cf:iconSet', text: '圖示 2/3' }, { key: 'cf:iconSet', text: '圖示 3/3' },
      ])
      zh.dispose()
      // reverse flips the glyph, never the positional label (bucket semantics).
      const reversed = (await scopeOf(iconSheet([10, 50, 90], { def: { reverse: true } }), 3, 'en-US'))!
      expect(reversed.iconAt?.(0, 0)?.iconId).toBe(2)
      expect(reversed.describe(0, 0)).toEqual({ key: 'cf:iconSet', text: 'Icon 1 of 3' })
      reversed.dispose()
      const custom = [{ set: '3Arrows' as const, iconId: 0 }, { set: 'NoIcons' as const, iconId: 0 }, { set: '3Flags' as const, iconId: 2 }]
      const mixed = (await scopeOf(iconSheet([10, 50, 90], { def: { custom } }), 3, 'en-US'))!
      expect(mixed.describe(1, 0), 'NoIcons adds nothing').toBeUndefined()
      expect(mixed.iconAt?.(1, 0)).toBeUndefined()
      expect(mixed.describe(2, 0)).toEqual({ key: 'cf:iconSet', text: 'Icon 3 of 3' })
      mixed.dispose()
    })
  })

  it('a bar and an icon on one cell announce one joined label', async () => {
    await withIcons(async () => {
      const sheet = iconSheet([0, 72, 100], { def: { thresholds: [{ bound: { type: 'num', value: 0 }, gte: true }, { bound: { type: 'num', value: 40 }, gte: true }, { bound: { type: 'num', value: 80 }, gte: true }] }, extraRules: [barRule({ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }, {}, 2)] })
      const scope = (await scopeOf(sheet, 3))!
      expect(scope.describe(1, 0)).toEqual({ key: 'cf:dataBar+iconSet', text: '資料橫條 72%、圖示 2/3' })
      scope.dispose()
    })
  })

  it('a bar and an icon whose showValue disagree announce unavailable (Ruling S4-SHOWVALUE-TOKEN)', async () => {
    await withIcons(async () => {
      const sheet = iconSheet([0, 72, 100], { def: { showValue: false }, extraRules: [barRule({ firstRow: 0, lastRow: 2, firstCol: 0, lastCol: 0 }, { showValue: true }, 2)] })
      const scope = (await scopeOf(sheet, 3, 'en-US'))!
      expect(scope.describe(1, 0)).toEqual(UNAVAILABLE_EN)
      scope.dispose()
    })
  })

  it('a withheld icon announces unavailable, never an icon; production announces no icon at all', async () => {
    await withIcons(async () => {
      const scope = (await scopeOf(iconSheet([10, 50, 90]), 3, 'en-US'))!
      expect(scope.iconAt?.(1, 0)).toBeUndefined()
      expect(scope.describe(1, 0)).toEqual(UNAVAILABLE_EN)
      scope.dispose()
    }, { glyphs: () => undefined })
    const production = (await scopeOf(iconSheet([10, 50, 90]), 3, 'en-US'))!
    expect(production.describe(1, 0)?.text ?? '').not.toMatch(/Icon/)
    production.dispose()
  })

  it('a malformed wire icon is marked on its copied state, so it announces unavailable, like a rejected bar', async () => {
    await withIcons(async () => {
      const wire = await wireOf(iconSheet([10, 50, 90]), 3)
      const broken = structuredClone(wire)
      for (const state of broken.states) if (state.row === 1 && state.icon) (state as { icon: unknown }).icon = { ...state.icon, iconId: 9 }
      const scope = (await prepareCfWireScope(async () => ({ result: broken, release() {} }), [{ r0: 0, r1: 2, c0: 0, c1: 0 }], { budget: { reserve() {}, release() {} }, locale: 'en-US' }))!
      expect(scope.describe(1, 0)).toEqual(UNAVAILABLE_EN)
      // The paged painter's coverage includes icon cells.
      expect(scope.regions).toEqual(expect.arrayContaining([{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 0 }, { minRow: 1, maxRow: 1, minCol: 0, maxCol: 0 }]))
      scope.dispose()
      expect(wireIcon({ ruleId: 1, set: '3Arrows', iconId: 1, bucket: 1, buckets: 3, showValue: true }).set).toBe(0)
    })
  })

  it('the value text stays in the accessible name, the grid keeps one tab stop, and an icon win misses the presentation cache', async () => {
    await withIcons(async () => {
      const sheet = iconSheet([10, 50, 90], { def: { showValue: false } })
      const scope = (await scopeOf(sheet, 3))!
      const desc = scope.describe(1, 0)!
      expect(desc).toEqual({ key: 'cf:iconSet', text: '圖示 2/3' })
      const name = withCfDescription('50', desc.text)
      expect(name).toBe('50; 圖示 2/3')
      const ui = render(<GridSemantics {...gridFixture({ cellPresentation: () => ({ text: name }) })} />)
      try {
        const cells = ui.getAllByRole('gridcell', { name: /圖示 2\/3/ })
        expect(cells.length).toBeGreaterThan(0)
        expect(cells[0]!.textContent).toContain('50')
        const tabbable = [...ui.container.querySelectorAll('[tabindex]')].filter(node => node.getAttribute('tabindex') !== '-1')
        expect(tabbable).toHaveLength(1)
        expect(tabbable[0]!.getAttribute('role')).toBe('grid')
      } finally { ui.unmount() }
      const { budget, settled } = trackingBudget()
      const cache = new SemanticPresentationCache(sheet, budget)
      const cell = sheet.rows[1]!.cells[0]!
      cache.get(cell, { gen: scope.gen, style: undefined, descriptorKey: 'cf:fill', description: 'fill' })
      const hit = cache.get(cell, { gen: scope.gen, style: undefined, descriptorKey: desc.key, description: desc.text })
      expect(settled).toHaveLength(1)
      expect(hit).toEqual({ text: '50', description: '圖示 2/3' })
      cache.dispose(); scope.dispose()
    }, { extra: ['icon:showvalue-canvas-suppression'] })
  })
})
