import { describe, expect, test } from 'vitest'
import type { XlsxCellStyle } from '../src/xlsx/types'
import type { CacheIdentity } from '../src/xlsx/view/contracts'
import { LayoutCache, type CellLayoutHost, type PlacedCell } from '../src/xlsx/view/textLayout'

const owner: CacheIdentity = { docId: 'doc', sheetId: 'sheet', valueGen: 0, fontGen: 0, viewGen: 0 }

function fixture(certified?: boolean) {
  let styleReads = 0, builds = 0
  const style: XlsxCellStyle = { numFmtId: 0 }
  const placed: PlacedCell = {
    cell: {
      ref: 'A1', row: 0, col: 0, value: 'text', styleIndex: 0,
      get style() { styleReads++; return style },
    },
    source: { row: 0, col: 0 }, x: 0, y: 0, width: 64, height: 24,
  }
  const host: CellLayoutHost = {
    layoutBody: (_cell, w, h, overflow) => {
      builds++
      return { x: 0, y: 0, w, h, formattedText: certified ? 'text' : String(overflow?.right) }
    },
    ...(certified === undefined ? {} : { canReuseOverflow: () => certified }),
  }
  return { placed, host, reads: () => styleReads, builds: () => builds }
}

describe('slice10: layout-key hot path', () => {
  test.each([true, false, undefined])('reads the cell style once per layout lookup (certification %s)', certified => {
    const f = fixture(certified), cache = new LayoutCache()
    let previous: ReturnType<LayoutCache['layout']> | undefined
    for (const [index, right] of [100, 200, 200].entries()) {
      const reads = f.reads()
      const body = cache.layout(owner, { ...f.placed, overflow: { left: 0, right } }, f.host)
      expect(f.reads() - reads).toBe(1)
      expect(body.formattedText).toBe(certified ? 'text' : String(right))
      if (previous && (certified || index === 2)) expect(body).toBe(previous)
      previous = body
    }
    expect(f.builds()).toBe(certified ? 1 : 2)
    expect(cache.stats.entries).toBe(certified ? 1 : 2)
    expect(cache.stats.reused).toBe(certified ? 2 : 1)
  })

  test.each([true, false, undefined])('charges only retained allowance digits (certification %s)', certified => {
    const f = fixture(certified), cache = new LayoutCache(1)
    cache.layout(owner, { ...f.placed, overflow: { left: 0, right: 1 } }, f.host)
    const bytes = cache.stats.bytes
    cache.layout(owner, { ...f.placed, overflow: { left: 0, right: 123456789 } }, f.host)
    // A shared body retains no allowance text; exact entries retain eight more
    // UTF-16 digits both in their key and in this host's formatted value.
    expect(cache.stats.bytes - bytes).toBe(certified ? 0 : 2 * 8 * 2)
    expect(cache.stats.entries).toBe(1)
  })
})
