/**
 * Task 22, Ruling S3-P: a painted data bar announces a fixed localized label with the value's position in its range;
 * the value text stays in the accessible name, and the bar adds no tab stop. Bars paint here under INJECTED evidence
 * only (tests/helpers/cf-bars.ts, exact injection); production paints from the Stage 3 set Task 23 published.
 *
 * The zh-TW term 「資料橫條」 is Excel's zh-TW name for data bars: the orchestrator selected it by that exact name in Excel
 * 16.114 zh-TW's Conditional Formatting style menu when authoring tests/fixtures/xlsx-cf/excel-authored-databar.xlsx.
 */
import { describe, expect, it } from 'vitest'
import { render } from '@testing-library/react'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { SemanticPresentationCache, withCfDescription } from '../src/components/xlsx/formatPresentation'
import { prepareCfTextScope, prepareCfWireScope } from '../src/xlsx/cf/text'
import { describeCfWin } from '../src/xlsx/cf/describe'
import { CF_LIMITS, createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { residentSource } from '../src/xlsx/cf/render'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import type { CfStamp } from '../src/xlsx/cf/model'
import type { XlsxSheet } from '../src/xlsx/types'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import { gridFixture } from './helpers/xlsx-grid-fixture'
import { barSheet, paintScopeBars, withBars } from './helpers/cf-bars'
import { VERIFIED_STAGE1_CAPABILITIES, VERIFIED_STAGE2_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'

function trackingBudget(): { budget: FormatBudget; settled: object[] } {
  const settled: object[] = []
  const budget: FormatBudget = { reserve() {}, release() {}, retain: (_t, _b, commit) => commit(), settle: target => { settled.push(target) } }
  return { budget, settled }
}
const scopeOf = (sheet: XlsxSheet, rows: number, locale = 'zh-TW') => prepareCfTextScope(sheet, {
  rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }], signal: new AbortController().signal, locale, budget: trackingBudget().budget,
})
const withLocale = (sheet: XlsxSheet, locale: string): XlsxSheet => ({ ...sheet, semantics: { locale, dateSystem: '1900' } } as unknown as XlsxSheet)

describe('bar cell announces a bounded label with its position, without a tab stop', () => {
  it('describeCfWin labels a 0.72 bar in zh-TW and en, bounded, with no colour and no thresholds', () => {
    const zh = describeCfWin({}, undefined, undefined, 'zh-TW', undefined, { fraction: 0.72 })
    expect(zh).toEqual({ key: 'cf:dataBar', text: '資料橫條 72%' })
    const en = describeCfWin({}, undefined, undefined, 'en-US', undefined, { fraction: 0.72 })
    expect(en).toEqual({ key: 'cf:dataBar', text: 'Data bar, 72% of range' })
    for (const d of [zh!, en!]) {
      expect(d.text.length).toBeLessThanOrEqual(CF_LIMITS.descriptorUnits)
      expect(d.text.length).toBeLessThanOrEqual(128)
      expect(d.text).not.toMatch(/#|FF0000FF|0000FF|rgb|min|max/i)
    }
    // Bounded to 0–100, and it precedes a scale and DXF aspects (one label per cell).
    expect(describeCfWin({}, undefined, undefined, 'en-US', undefined, { fraction: 1.5 })?.text).toBe('Data bar, 100% of range')
    expect(describeCfWin({ 'font.bold': true }, undefined, undefined, 'en-US', { fraction: 0.1 }, { fraction: 0.72 })?.key).toBe('cf:dataBar')
    // An unresolved bar announces only the existing unavailable descriptor.
    expect(describeCfWin({}, new Set(['dataBar']), undefined, 'zh-TW', undefined, { fraction: 0.72 })).toEqual({ key: 'cf-unavailable', text: '無法使用條件式格式設定' })
  })

  it('a prepared scope announces each bar at its own position: positive, negative and all-negative ranges', async () => {
    await withBars(async () => {
      const zh = (await scopeOf(withLocale(barSheet([0, 72, 100]), 'zh-TW'), 3, 'zh-TW'))!
      expect(zh.describe(1, 0)).toEqual({ key: 'cf:dataBar', text: '資料橫條 72%' })
      // The descriptor cache is keyed by the bar too: the next cell's identical (empty) properties get their own text.
      expect(zh.describe(2, 0)).toEqual({ key: 'cf:dataBar', text: '資料橫條 100%' })
      expect(zh.describe(0, 0)).toEqual({ key: 'cf:dataBar', text: '資料橫條 0%' })
      zh.dispose()
      // Mixed sign under an automatic axis: the value's position is the bar end away from the axis.
      const mixed = (await scopeOf(barSheet([-50, 0, 50, 100]), 4, 'en-US'))!
      expect([0, 1, 2, 3].map(r => mixed.describe(r, 0)?.text)).toEqual(['Data bar, 0% of range', 'Data bar, 33% of range', 'Data bar, 67% of range', 'Data bar, 100% of range'])
      mixed.dispose()
    })
    await withBars(async () => {
      const negative = (await scopeOf(barSheet([-100, -60, -20]), 3, 'en-US'))!
      expect([0, 1, 2].map(r => negative.describe(r, 0)?.text)).toEqual(['Data bar, 0% of range', 'Data bar, 50% of range', 'Data bar, 100% of range'])
      negative.dispose()
    }, { extra: ['bar:lengths-all-negative'] })
  })

  it('the value text stays in the accessible name, the grid keeps one tab stop, and a bar win misses the presentation cache', async () => {
    await withBars(async () => {
      const sheet = withLocale(barSheet([0, 72, 100], { def: { showValue: false } }), 'zh-TW')
      const scope = (await scopeOf(sheet, 3))!
      const desc = scope.describe(1, 0)!
      expect(desc.key).toBe('cf:dataBar')
      // showValue="0" hides only canvas text: the accessible name still carries the value.
      const name = withCfDescription('72', desc.text)
      expect(name).toBe('72; 資料橫條 72%')
      const ui = render(<GridSemantics {...gridFixture({ cellPresentation: () => ({ text: name }) })} />)
      try {
        const cells = ui.getAllByRole('gridcell', { name: /資料橫條 72%/ })
        expect(cells.length).toBeGreaterThan(0)
        expect(cells[0]!.textContent).toContain('72')
        const tabbable = [...ui.container.querySelectorAll('[tabindex]')].filter(node => node.getAttribute('tabindex') !== '-1')
        expect(tabbable).toHaveLength(1)
        expect(tabbable[0]!.getAttribute('role')).toBe('grid')
      } finally { ui.unmount() }
      // The same generation and style with another descriptor key is a miss: the stale entry is settled and rebuilt.
      const { budget, settled } = trackingBudget()
      const cache = new SemanticPresentationCache(sheet, budget)
      const cell = sheet.rows[1]!.cells[0]!
      cache.get(cell, { gen: scope.gen, style: undefined, descriptorKey: 'cf:fill', description: 'fill' })
      expect(settled).toHaveLength(0)
      const hit = cache.get(cell, { gen: scope.gen, style: undefined, descriptorKey: desc.key, description: desc.text })
      expect(settled).toHaveLength(1)
      expect(hit).toEqual({ text: '72', description: '資料橫條 72%' })
      cache.dispose(); scope.dispose()
    }, { extra: ['bar:showvalue-canvas-suppression'] })
  })

  it('an unresolved or withheld bar announces only the unavailable descriptor', async () => {
    // Without Stage 3 evidence (production's Stage 1/2 sets only, as before Task 23 published) the bar is unresolved.
    await withExactCfEvidence([...VERIFIED_STAGE1_CAPABILITIES, ...VERIFIED_STAGE2_CAPABILITIES], async () => {
      const plain = (await scopeOf(barSheet([0, 72, 100]), 3))!
      expect(plain.describe(1, 0)).toEqual({ key: 'cf-unavailable', text: '無法使用條件式格式設定' })
      plain.dispose()
    })
    // Production evidence since Task 23: the same bar paints and announces its position.
    const published = (await scopeOf(barSheet([0, 72, 100]), 3))!
    expect(published.describe(1, 0)).toEqual({ key: 'cf:dataBar', text: '資料橫條 72%' })
    published.dispose()
    // Evaluated, but the painter withholds it (no paint gates): announced the same way, never as a bar.
    await withBars(async () => {
      const scope = (await scopeOf(barSheet([0, 72, 100]), 3, 'en-US'))!
      expect(scope.barAt?.(1, 0)).toBeUndefined()
      expect(scope.describe(1, 0)).toEqual({ key: 'cf-unavailable', text: 'Conditional formatting unavailable' })
      scope.dispose()
    }, { paint: false })
  })

  it('the paged wire scope describes and covers bar-only cells', async () => {
    await withBars(async () => {
      const sheet = barSheet([0, 72, 100])
      const rules = sheet.conditionalFormatting!, sheetId = sheet.sheetId!
      const st: CfStamp = { owner: 'paged', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
      const memory = createCfMemory({}), rects = [{ r0: 0, r1: 2, c0: 0, c1: 0 }]
      const wire = await prepareCfRemote({ rects, stamp: { ...st }, grantBytes: 1 << 22 }, sheetId, residentSource(sheet, st, memory, 'engine'), rules, memory, { budget: () => createCfWorkBudget({ now: () => 0 }) })
      expect(wire.kind).toBe('ready')
      const scope = (await prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), rects, { budget: { reserve() {}, release() {} }, locale: 'en-US' }))!
      // Before any paint the paged scope cannot know whether painting withholds the bar: never announced as a bar.
      expect(scope.describe(1, 0)).toEqual({ key: 'cf-unavailable', text: 'Conditional formatting unavailable' })
      await paintScopeBars(scope, 3, 1)
      expect(scope.describe(1, 0)).toEqual({ key: 'cf:dataBar', text: 'Data bar, 72% of range' })
      // The paged painter's coverage includes the cells whose only state is a bar.
      expect(scope.regions).toEqual(expect.arrayContaining([{ minRow: 1, maxRow: 1, minCol: 0, maxCol: 0 }]))
      scope.dispose()
    })
  })
})

/**
 * Fix round 2: a cell whose bar painting withholds announces no data bar and `barAt` does not report it, by the
 * painter's own predicate (`barWithheld`). Two reasons: a non-white axis over a filled host (its dash gap is
 * unmeasured), and a non-zero share with no visible pixel at Excel's capture scale (minimum visible width unmeasured).
 */
describe('announcement tracks paint withholding', () => {
  const UNAVAILABLE = { key: 'cf-unavailable', text: 'Conditional formatting unavailable' }
  // [-50, 50, 100] under an automatic black axis; the cells carry a yellow static fill.
  const filledAxisSheet = (): XlsxSheet => {
    const sheet = barSheet([-50, 50, 100])
    for (const row of sheet.rows) for (const cell of row.cells) cell.style = { numFmtId: 0, fillColor: 'FFFFFF00' }
    return sheet
  }
  // 0.001 of [0, 1000]: a share that rounds to no pixel at the capture scale.
  const subPixelSheet = (): XlsxSheet => barSheet([0, 0.001, 1000])

  it('resident: the text scope and the paint batch withhold both reasons', async () => {
    const { prepareCfPaint } = await import('../src/xlsx/cf/render')
    const { cfRuntimeEvidence } = await import('../src/xlsx/cf/predicates')
    await withBars(async () => {
      for (const [name, sheet, row] of [['filled axis', filledAxisSheet(), 1], ['sub-pixel', subPixelSheet(), 1]] as const) {
        const scope = (await scopeOf(sheet, 3, 'en-US'))!
        expect(scope.barAt?.(row, 0), name).toBeUndefined()
        expect(scope.describe(row, 0), name).toEqual(UNAVAILABLE)
        expect(scope.describe(2, 0)?.key, `${name}: the full bar of the same rule`).toBe(name === 'sub-pixel' ? 'cf:dataBar' : 'cf-unavailable')
        scope.dispose()
        const batch = (await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence }))!
        expect(batch.barAt(row, 0), name).toBeUndefined()
        batch.dispose()
      }
      // Positive control: the same axis bars on unfilled cells are announced and reported.
      const plain = (await scopeOf(barSheet([-50, 50, 100]), 3, 'en-US'))!
      expect(plain.describe(1, 0)).toEqual({ key: 'cf:dataBar', text: 'Data bar, 67% of range' })
      expect(plain.barAt?.(1, 0)).toBeDefined()
      plain.dispose()
    })
  })

  it('paged: no bar is announced before it paints, nor after painting withholds it', async () => {
    await withBars(async () => {
      for (const [name, sheet, fill] of [['filled axis', barSheet([-50, 50, 100]), '#FFFF00'], ['sub-pixel', subPixelSheet(), undefined], ['control', barSheet([-50, 50, 100]), undefined]] as const) {
        const rules = sheet.conditionalFormatting!, sheetId = sheet.sheetId!
        const st: CfStamp = { owner: 'paged', sheetId, workbookValueGen: 1, cfGen: cfRulesVersion(rules), viewGen: 1, structureGen: 1, clockGen: 1 }
        const memory = createCfMemory({}), rects = [{ r0: 0, r1: 2, c0: 0, c1: 0 }]
        const wire = await prepareCfRemote({ rects, stamp: { ...st }, grantBytes: 1 << 22 }, sheetId, residentSource(sheet, st, memory, 'engine'), rules, memory, { budget: () => createCfWorkBudget({ now: () => 0 }) })
        const scope = (await prepareCfWireScope(async () => ({ result: structuredClone(wire), release() {} }), rects, { budget: { reserve() {}, release() {} }, locale: 'en-US' }))!
        expect(scope.describe(1, 0), `${name} before paint`).toEqual(UNAVAILABLE)
        const { createCanvas } = await import('canvas')
        const ctx = createCanvas(200, 60).getContext('2d') as unknown as CanvasRenderingContext2D
        scope.paintBars?.(ctx, (row, col) => ({ x: col * 200, y: row * 20, width: 200, height: 20, ...(fill ? { fill } : {}) }), { r0: 0, c0: 0, r1: 2, c1: 0 }, { zoom: 1, dpr: 1 })
        if (name === 'control') {
          expect(scope.describe(1, 0)).toEqual({ key: 'cf:dataBar', text: 'Data bar, 67% of range' })
          expect(scope.barAt?.(1, 0)).toBeDefined()
        } else {
          expect(scope.describe(1, 0), `${name} after paint`).toEqual(UNAVAILABLE)
          expect(scope.barAt?.(1, 0), name).toBeUndefined()
        }
        scope.dispose()
      }
    })
  })
})
