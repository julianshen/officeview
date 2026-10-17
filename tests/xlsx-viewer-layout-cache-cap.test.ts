/**
 * Confirmation r5 F3: the layout cache's byte budget is a real bound. A record
 * larger than the whole budget is painted uncached and never retained, and the
 * cache evicts before inserting, so retained bytes never exceed the budget.
 */
import { describe, expect, it } from 'vitest'
import { LayoutCache, estimateLayoutBytes, type CellBodyLayout, type CellLayoutHost, type PlacedCell } from '../src/xlsx/view/textLayout'
import type { CacheIdentity } from '../src/xlsx/view/contracts'

const owner: CacheIdentity = { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
const placedOf = (col: number, text: string): PlacedCell => ({ cell: { ref: 'A1', row: 0, col, value: text, styleIndex: 0 }, source: { row: 0, col }, x: 0, y: 0, width: 64, height: 20 })
const host: CellLayoutHost = { layoutBody: (cell, w, h) => ({ x: 0, y: 0, w, h, formattedText: String(cell.value) }) as CellBodyLayout }

describe('LayoutCache byte budget', () => {
  it('never retains a record larger than the whole budget; it is still returned for painting', () => {
    const cache = new LayoutCache(100, 4096)
    const big = cache.layout(owner, placedOf(0, 'x'.repeat(4000)), host)
    expect(big.formattedText).toHaveLength(4000)
    expect(cache.stats.bytes).toBe(0)
    expect(cache.stats.entries).toBe(0)
  })
  it('evicts before inserting: retained bytes stay within the budget after every insertion', () => {
    const cache = new LayoutCache(100, 6000)
    let max = 0
    for (let col = 0; col < 40; col++) {
      cache.layout(owner, placedOf(col, 'y'.repeat(400 + col * 10)), host)
      max = Math.max(max, cache.stats.bytes)
    }
    expect(max).toBeLessThanOrEqual(6000)
    expect(estimateLayoutBytes('k', { x: 0, y: 0, w: 1, h: 1, formattedText: 'y'.repeat(400) })).toBeGreaterThan(0)
  })
})
