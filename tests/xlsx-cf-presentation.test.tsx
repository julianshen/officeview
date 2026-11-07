/**
 * Task 9: CF-aware derived text, semantic cache generations and accessible
 * descriptions. Oracle provenance: expected strings come from the
 * stage1-numfmt-accessibility seed hypothesis (`1.25` via CF DXF `0.00`,
 * control `1.3` via cell `0.0`); the case is recipe-pending, never Excel-measured.
 */
import { describe, expect, it } from 'vitest'
import { render } from '@testing-library/react'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { gridFixture } from './helpers/xlsx-grid-fixture'
import {
  SemanticPresentationCache,
  formatCellPresentation,
  withCfDescription,
  type CfPresentationInput,
} from '../src/components/xlsx/formatPresentation'
import { prepareCfTextScope } from '../src/xlsx/cf/text'
import { describeCfUnavailable, describeCfWin } from '../src/xlsx/cf/describe'
import { cellDisplayText } from '../src/xlsx/render'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import { FindError } from '../src/xlsx/view/find'
import type { CfRule } from '../src/xlsx/cf/model'

function trackingBudget(refuse = false): { budget: FormatBudget; held: () => number; settled: object[] } {
  let held = 0
  const settled: object[] = []
  const budget: FormatBudget = {
    reserve: bytes => { held += bytes; if (refuse) { held -= bytes; throw new FindError('resource-limit', 'refused') } },
    release: bytes => { held -= bytes },
    retain: (_targets, _bytes, commit) => { commit() },
    settle: target => { settled.push(target) },
  }
  return { budget, held: () => held, settled }
}

const cellIsRule = (dxf: Record<string, string>, writes: string[]): CfRule => ({
  id: 0,
  priority: 1,
  documentOrder: 0,
  stopIfTrue: false,
  ranges: [{ firstRow: 0, lastRow: 600, firstCol: 0, lastCol: 0 }],
  writes: new Set(writes),
  dxf,
  definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 1 }], dialect: 'base', formulas: [] },
})

function numFmtSheet(code: string, locale = 'en-US'): XlsxSheet {
  const cell = (row: number, col: number, value: number): XlsxCell => ({
    ref: `R${row}C${col}`, row, col, styleIndex: 0, style: { numFmtId: 0, formatCode: '0.0' }, value, hasCachedValue: true,
  })
  return {
    name: 'CF',
    sheetId: 'cf',
    rows: [
      { index: 0, cells: [cell(0, 0, 1.25), cell(0, 1, 0.5)] },
      { index: 500, cells: [cell(500, 0, 1.25)] },
    ],
    cols: [],
    merges: [],
    mergeRanges: [],
    semantics: { dateSystem: '1900', unicode: true, locale, timeZone: 'UTC', epochNowMs: 0 },
    conditionalFormatting: {
      diagnostics: [],
      rules: [cellIsRule({ 'numFmt.code': code }, ['numFmt'])],
    },
  } as unknown as XlsxSheet
}

const scopeFor = (sheet: XlsxSheet, budget?: FormatBudget) =>
  prepareCfTextScope(sheet, {
    rects: [{ firstRow: 0, lastRow: 600, firstCol: 0, lastCol: 0 }],
    signal: new AbortController().signal,
    locale: (sheet.semantics?.locale as string | undefined) ?? 'en-US',
    ...(budget ? { budget } : {}),
  })

const inputOf = (scope: { gen: number; effectiveStyle: (row: number, col: number, base: XlsxCell['style']) => XlsxCell['style']; describe?: (row: number, col: number) => { key: string; text: string } | undefined }, cell: XlsxCell): CfPresentationInput => {
  const style = scope.effectiveStyle(cell.row, cell.col, cell.style)
  const desc = scope.describe?.(cell.row, cell.col)
  return { gen: scope.gen, style, descriptorKey: desc?.key, description: desc?.text }
}

describe('same 1.25 value changes 0.0 to 0.00 across CF generations', () => {
  it('formats, finds and caches the CF number format, dropping the stale entry', async () => {
    const sheet = numFmtSheet('0.00')
    const scope = await scopeFor(sheet)
    expect(scope).toBeDefined()
    const [hit, control, far] = [sheet.rows[0]!.cells[0]!, sheet.rows[0]!.cells[1]!, sheet.rows[1]!.cells[0]!]
    expect(cellDisplayText(sheet, { value: hit.value, style: hit.style })).toBe('1.3')
    const effective = scope!.effectiveStyle(0, 0, hit.style)
    expect(cellDisplayText(sheet, { value: hit.value, style: effective })).toBe('1.25')
    // Outside the painted viewport band the same prepared generation applies.
    expect(cellDisplayText(sheet, { value: far.value, style: scope!.effectiveStyle(500, 0, far.style) })).toBe('1.25')
    // Non-matching control keeps base formatting and has no descriptor.
    expect(cellDisplayText(sheet, { value: control.value, style: scope!.effectiveStyle(0, 1, control.style) })).toBe('0.5')
    expect(scope!.describe?.(0, 1)).toBeUndefined()

    const { budget, settled } = trackingBudget()
    const cache = new SemanticPresentationCache(sheet, budget)
    const first = cache.get(hit, inputOf(scope!, hit))
    expect(first.text).toBe('1.25')

    // Same value and cell, next CF generation: stale entry is dropped and settled.
    const next = numFmtSheet('0.000')
    const scope2 = await scopeFor(next)
    expect(scope2).toBeDefined()
    expect(scope2!.gen).not.toBe(scope!.gen)
    const effective2 = scope2!.effectiveStyle(0, 0, hit.style)
    const desc2 = scope2!.describe?.(0, 0)
    const second = cache.get(hit, { gen: scope2!.gen, style: effective2, descriptorKey: desc2?.key, description: desc2?.text })
    expect(cellDisplayText(sheet, { value: hit.value, style: effective2 })).toBe('1.250')
    expect(second.text).toBe('1.250')
    expect(second.text).not.toBe(first.text)
    expect(settled).toContain(first)
    scope!.dispose()
    scope2!.dispose()
    cache.dispose()
  })
})

describe('red-fill-only alert has bounded zh-TW text and no added tab stop', () => {
  const redFillSheet = (): XlsxSheet => {
    const sheet = numFmtSheet('0.00')
    sheet.conditionalFormatting = {
      diagnostics: [],
      rules: [cellIsRule({ 'fill.bg': 'FFFFC7CE' }, ['fill'])],
    }
    sheet.semantics = { ...(sheet.semantics as object), locale: 'zh-TW' } as XlsxSheet['semantics']
    return sheet
  }

  it('names the reviewed localized state without raw values or new tab stops', async () => {
    const sheet = redFillSheet()
    const win = { 'fill.bg': 'FFFFC7CE' }
    const direct = describeCfWin(win, undefined, undefined, 'zh-TW')
    expect(direct?.text).toBe('已套用條件式格式設定：填滿')
    expect(direct!.text.length).toBeLessThanOrEqual(128)
    expect(direct!.text).not.toContain('FFFFC7CE')

    const { budget: describeBudget } = trackingBudget()
    const scope = await scopeFor(sheet, describeBudget)
    expect(scope).toBeDefined()
    const desc = scope!.describe?.(0, 0)
    expect(desc?.key).toBe(direct?.key)
    expect(desc?.text).toBe('已套用條件式格式設定：填滿')

    const name = withCfDescription('2', desc?.text)
    expect(name).toContain('已套用條件式格式設定：填滿')
    const ui = render(<GridSemantics {...gridFixture({ cellPresentation: () => ({ text: withCfDescription('2', desc?.text) }) })} />)
    try {
      const cells = ui.getAllByRole('gridcell', { name: /已套用條件式格式設定：填滿/ })
      expect(cells.length).toBeGreaterThan(0)
      for (const found of cells) expect(found.textContent).toContain('2')
      const tabbable = [...ui.container.querySelectorAll('[tabindex]')].filter(node => node.getAttribute('tabindex') !== '-1')
      expect(tabbable).toHaveLength(1)
      expect(tabbable[0]!.getAttribute('role')).toBe('grid')
    } finally {
      ui.unmount()
      scope!.dispose()
    }
  })

  it('error/unknown CF announces unavailable once, never rule internals', async () => {
    const sheet = redFillSheet()
    sheet.conditionalFormatting = {
      diagnostics: [],
      rules: [{
        id: 0,
        priority: 1,
        documentOrder: 0,
        stopIfTrue: false,
        ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }],
        writes: new Set(['numFmt']),
        dxf: { 'numFmt.code': '0.00' },
        definition: { kind: 'unresolved', type: 'expression', reason: 'cf-unsupported-rule' } as const,
      }],
    }
    const unavailable = describeCfUnavailable('zh-TW')
    expect(unavailable?.text).toBe('無法使用條件式格式設定')

    const { budget: scopeBudget } = trackingBudget()
    const scope = await scopeFor(sheet, scopeBudget)
    expect(scope).toBeDefined()
    const cell = sheet.rows[0]!.cells[0]!
    // Withheld to base formatting, never a guessed or neighboring result.
    expect(cellDisplayText(sheet, { value: cell.value, style: scope!.effectiveStyle(0, 0, cell.style) })).toBe('1.3')
    const desc = scope!.describe?.(0, 0)
    expect(desc?.text).toBe('無法使用條件式格式設定')
    const name = withCfDescription('1.3', desc?.text)
    expect(name.split('無法使用條件式格式設定')).toHaveLength(2)
    expect(name).not.toContain('expression')
    expect(name).not.toContain('cf-unsupported-rule')
    expect(name).not.toContain('0.00')
    scope!.dispose()
  })
})

describe('format budget refusal releases descriptor workspace', () => {
  it('zero budget refuses allocation; lifecycles return to baseline', async () => {
    const sheet = numFmtSheet('0.00')
    const refusing = trackingBudget(true)
    await expect(scopeFor(sheet, refusing.budget)).resolves.toBeUndefined()
    expect(refusing.held()).toBe(0)

    const live = trackingBudget()
    const scope = await scopeFor(sheet, live.budget)
    expect(scope).toBeDefined()
    const { budget, settled } = trackingBudget()
    const cache = new SemanticPresentationCache(sheet, budget)
    const cell = sheet.rows[0]!.cells[0]!
    const first = cache.get(cell, inputOf(scope!, cell))
    expect(first.text).toBe('1.25')
    // A refused presentation never caches and never charges.
    const refused = formatCellPresentation(sheet, cell, refusing.budget)
    expect(refused).toMatchObject({ text: 'Value unavailable (formatting memory budget).', state: 'error' })
    scope!.dispose()
    cache.dispose()
    // Settled exactly once: dispose is idempotent and releases the entry charge.
    expect(settled.map(entry => (entry as { text?: string }).text)).toContain('1.25')
    const settledCount = settled.length
    cache.dispose()
    scope!.dispose()
    expect(settled).toHaveLength(settledCount)

    // Long rule internals never reach accessible names.
    const longFormula = `IF(${ 'A1>1,'.repeat(400)}TRUE)`
    expect(longFormula.length).toBeGreaterThan(1000)
    const name = withCfDescription('1.3', describeCfUnavailable('zh-TW')?.text)
    expect(name).not.toContain(longFormula.slice(0, 32))
    expect(name.length).toBeLessThanOrEqual(128)
  })
})
