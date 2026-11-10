// @vitest-environment node
/**
 * I2: derived CF text scopes (copy, find, accessibility) walk their rects in chunks, but the statistical
 * scope of a scale / top10 / aboveAverage / duplicate rule is scanned ONCE per generation, not once per chunk.
 */
import { describe, expect, it } from 'vitest'
import { prepareCfTextScope } from '../src/xlsx/cf/text'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import type { XlsxSheet } from '../src/xlsx/types'

const budget: FormatBudget = { reserve: () => {}, release: () => {}, retain: (_t, _b, commit) => { commit() }, settle: () => {} } as unknown as FormatBudget

function countingSheet(rows: number, definition: Record<string, unknown>, writes: string[]): { sheet: XlsxSheet; reads: () => number } {
  let reads = 0
  const bands = Array.from({ length: rows }, (_, r) => ({
    index: r,
    cells: [Object.defineProperty({ row: r, col: 0, ref: `A${r + 1}`, styleIndex: 0 } as Record<string, unknown>, 'value', { get() { reads++; return r + 1 }, enumerable: true })],
  }))
  const sheet = {
    name: 'S', sheetId: 's', cols: [], merges: [], mergeRanges: [], rows: bands,
    conditionalFormatting: { diagnostics: [], rules: [{
      id: 1, priority: 1, documentOrder: 0, stopIfTrue: false,
      ranges: [{ firstRow: 0, lastRow: rows - 1, firstCol: 0, lastCol: 0 }],
      writes: new Set(writes), ...(definition.kind === 'top10' ? { dxf: { 'fill.bg': 'FF00FF00' } } : {}), definition,
    }] },
  } as unknown as XlsxSheet
  return { sheet, reads: () => reads }
}

describe('derived CF text scopes share one statistics session across chunks (I2)', () => {
  const ROWS = 20_000, REQUEST = 5_000
  const request = { rects: [{ r0: 0, r1: REQUEST - 1, c0: 0, c1: 0 }], signal: new AbortController().signal, budget }

  it('a colour-scale scope is scanned once for a multi-chunk request', async () => {
    const { sheet, reads } = countingSheet(ROWS, {
      kind: 'colorScale', dialect: 'base', formulas: [],
      points: [{ bound: { type: 'min' }, argb: 0xfff8696b }, { bound: { type: 'max' }, argb: 0xff63be7b }],
    }, ['colorScale', 'fill'])
    const scope = await prepareCfTextScope(sheet, request)
    expect(scope).toBeDefined()
    // 20k statistics scan + 5k evaluation reads (+ slack); five chunks used to scan 5 x 20k.
    expect(reads()).toBeLessThanOrEqual(ROWS + REQUEST + 1_000)
    scope?.dispose()
  })

  it('a top10 scope is scanned once for a multi-chunk request', async () => {
    const { sheet, reads } = countingSheet(ROWS, { kind: 'top10', dialect: 'base', rank: 3, percent: false, bottom: false, formulas: [] }, ['fill'])
    const scope = await prepareCfTextScope(sheet, request)
    expect(scope).toBeDefined()
    expect(reads()).toBeLessThanOrEqual(ROWS + REQUEST + 1_000)
    scope?.dispose()
  })
})
