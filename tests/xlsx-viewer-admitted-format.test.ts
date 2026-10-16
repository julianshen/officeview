/**
 * Repaint memory admitted per record (spec §§5.6, 12.4), follow-ups:
 * - a cache miss formats its text ONCE: the admitted (observed) formatting
 *   pass feeds the canonical layout;
 * - find-hit geometry builds its record inside an admitted batch and falls
 *   back to the whole-cell highlight when it is not admitted.
 */
import { describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import * as format from '../src/xlsx/formula/format'
import { createCellLayoutHost } from '../src/xlsx/view/renderViewport'
import { LayoutCache, type LayoutAdmission, type PlacedCell } from '../src/xlsx/view/textLayout'
import { hitGeometry, type FindHit } from '../src/xlsx/view/find'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import type { CacheIdentity } from '../src/xlsx/view/contracts'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'

const owner: CacheIdentity = { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
const host = () => createCellLayoutHost(createCanvas(400, 40).getContext('2d') as unknown as CanvasRenderingContext2D, sheet)
function admission(limit = Infinity) {
  let held = 0
  const api: LayoutAdmission & { held: () => number } = {
    admit: bytes => { if (held + bytes > limit) return false; held += bytes; return true },
    release: bytes => { held -= bytes },
    hold: bytes => { held += bytes },
    cacheChanged: () => {},
    held: () => held,
  }
  return api
}

describe('a cache miss formats its text once', () => {
  it('the admitted formatting pass feeds the canonical layout (formatCommon runs once per miss, none per hit)', () => {
    const spy = vi.spyOn(format, 'formatCommon')
    const cache = new LayoutCache()
    const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value: 3, styleIndex: 0, style: { numFmtId: 164, formatCode: '0.00" units"' } }
    const placed: PlacedCell = { cell, source: { row: 0, col: 0 }, x: 0, y: 0, width: 400, height: 20 }
    const gate = admission()
    cache.beginBatch(gate)
    const body = cache.layout(owner, placed, host())
    cache.endBatch()
    expect(body.formattedText).toBe('3.00 units')
    expect(spy).toHaveBeenCalledTimes(1)
    cache.beginBatch(gate)
    cache.layout(owner, placed, host())
    cache.endBatch()
    expect(spy).toHaveBeenCalledTimes(1)
    expect(gate.held()).toBe(0)
    spy.mockRestore()
  })
})

describe('find-hit geometry is admitted', () => {
  const metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 1_048_576, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: 16_384, defaultSizeCss: 64 }) })
  const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value: `needle ${'a\n'.repeat(9000)}`, styleIndex: 0, style: { numFmtId: 0, wrapText: true } }
  const hit = { sheetId: 's', source: { row: 0, col: 0 }, start: 0, end: 6, text: cell.value } as unknown as FindHit
  it('builds the record inside an admitted batch and returns the whole-cell highlight when it is not admitted', () => {
    const canonical = host()
    const layouts = vi.spyOn(canonical, 'layoutBody')
    const refused = hitGeometry(hit, cell, metrics, owner, canonical, new LayoutCache(), undefined, admission(64 * 1024))!
    expect(layouts).not.toHaveBeenCalled()
    expect(refused.record).toBeUndefined()
    expect(refused.rects).toEqual([{ x: 0, y: 0, width: 64, height: 20 }])
    const gate = admission()
    const admitted = hitGeometry(hit, cell, metrics, owner, canonical, new LayoutCache(), undefined, gate)!
    expect(layouts).toHaveBeenCalledTimes(1)
    expect(admitted.record).toBeDefined()
    expect(gate.held()).toBe(0)
  })
})
