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
import * as cellFormat from '../src/xlsx/cellFormat'
import { createAdmittedLayoutContext } from '../src/components/xlsx/measureContext'
import { cellDisplayText } from '../src/xlsx/render'
import { createCellLayoutHost } from '../src/xlsx/view/renderViewport'
import { LayoutCache, estimateLayoutBytes, type LayoutAdmission, type PlacedCell } from '../src/xlsx/view/textLayout'
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

describe('confirmation r7', () => {
  it('R4 an authored numeric format without a built-in id is formatted once per miss', () => {
    const spy = vi.spyOn(format, 'formatCommon')
    const cache = new LayoutCache()
    const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value: 3, styleIndex: 0, style: { formatCode: '0.00" u"' } as XlsxCell['style'] }
    const gate = admission()
    cache.beginBatch(gate)
    const body = cache.layout(owner, { cell, source: { row: 0, col: 0 }, x: 0, y: 0, width: 400, height: 20 }, host())
    cache.endBatch()
    expect(body.formattedText).toBe('3.00 u')
    expect(spy).toHaveBeenCalledTimes(1)
    spy.mockRestore()
  })

  it('R1 a warm hit used by the batch keeps its bytes owned by the batch when a later miss evicts it', () => {
    const fake: import('../src/xlsx/view/textLayout').CellLayoutHost = { layoutBody: (cell, w, h) => ({ x: 0, y: 0, w, h, formattedText: String(cell.value) }) as never }
    const placed = (col: number) => ({ cell: { ref: '', row: 0, col, value: 'z'.repeat(1000), styleIndex: 0 } as XlsxCell, source: { row: 0, col }, x: 0, y: 0, width: 64, height: 20 })
    const one = estimateLayoutBytes(`|0,0|64|20||${''}|string:${'z'.repeat(1000)}`, { x: 0, y: 0, w: 64, h: 20, formattedText: 'z'.repeat(1000) })
    const cache = new LayoutCache(100, Math.floor(one * 2.5))
    const gate = admission()
    cache.beginBatch(gate); cache.layout(owner, placed(0), fake); cache.layout(owner, placed(1), fake); cache.endBatch()
    expect(gate.held()).toBe(0)
    cache.beginBatch(gate)
    cache.layout(owner, placed(0), fake) // warm hit, now held by the batch's records
    cache.layout(owner, placed(1), fake) // warm hit
    cache.layout(owner, placed(2), fake) // miss: evicts the LRU hit (col 0)
    // The evicted hit is still alive in this batch: its bytes moved to the batch hold.
    expect(gate.held()).toBeGreaterThanOrEqual(one * 0.9)
    cache.endBatch()
    expect(gate.held()).toBe(0)
  })

  it('R3 a formatting refusal evicts idle layout entries and retries before failing', () => {
    let cached = 0, held = 0
    const limit = { value: Infinity }
    const gate: LayoutAdmission = {
      admit: bytes => { if (held + cached + bytes > limit.value) return false; held += bytes; return true },
      release: bytes => { held -= bytes }, hold: bytes => { held += bytes }, cacheChanged: delta => { cached += delta },
    }
    const cache = new LayoutCache()
    cache.beginBatch(gate)
    for (let col = 1; col <= 48; col++) cache.layout(owner, { cell: { ref: '', row: 0, col, value: 'i'.repeat(4000), styleIndex: 0 }, source: { row: 0, col }, x: 0, y: 0, width: 64, height: 20 }, host())
    cache.endBatch()
    expect(cached).toBeGreaterThan(0)
    limit.value = cached + 8 * 1024
    const heavy: XlsxCell = { ref: 'A1', row: 0, col: 0, value: 'w'.repeat(64), styleIndex: 0, style: { numFmtId: 164, formatCode: `0;0;0;${'@'.repeat(64)}` } }
    cache.beginBatch(gate)
    const body = cache.layout(owner, { cell: heavy, source: { row: 0, col: 0 }, x: 0, y: 0, width: 64, height: 20 }, host())
    cache.endBatch()
    expect(body.formattedText.length).toBe(64 * 64)
    expect(cache.stats.evicted).toBeGreaterThan(0)
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

describe('merge review (integration B formatting through E admission)', () => {
  it('a conditional/extended format admits before building any section body (refusal first)', () => {
    const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value: 3, styleIndex: 0, style: { numFmtId: 164, formatCode: `[>0]"${'q'.repeat(65_536)}"0` } }
    const order: string[] = []
    const refuse: LayoutAdmission = { admit: () => { order.push('admit'); return false }, release: () => {}, hold: () => {}, cacheChanged: () => {} }
    const slice = String.prototype.slice
    const spy = vi.spyOn(String.prototype, 'slice').mockImplementation(function (this: string, start?: number, end?: number) {
      const out = slice.call(this, start, end)
      if (out.length > 60_000) order.push('big-slice')
      return out
    })
    try { expect(() => host().paintedText!(cell, 64, refuse)).toThrow() } finally { spy.mockRestore() }
    expect(order[0]).toBe('admit')
    expect(order).not.toContain('big-slice')
  })

  for (const formatCode of ['[>0]0.00" units";0', '0.00"General"']) {
    it(`a width-independent format formats once per miss (${formatCode})`, () => {
      const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value: 3, styleIndex: 0, style: { numFmtId: 164, formatCode } }
      const extended = vi.spyOn(cellFormat, 'formatExtended'), common = vi.spyOn(format, 'formatCommon')
      const calls = () => extended.mock.calls.length + common.mock.calls.length
      cellDisplayText(sheet, cell)
      const onePass = calls()
      extended.mockClear(); common.mockClear()
      const canonical = host()
      const layouts = vi.spyOn(canonical, 'layoutBody')
      const cache = new LayoutCache()
      const gate = admission()
      cache.beginBatch(gate)
      cache.layout(owner, { cell, source: { row: 0, col: 0 }, x: 0, y: 0, width: 400, height: 20 }, canonical)
      cache.endBatch()
      // One formatting pass (the admitted one), fed to the layout as its text.
      expect(calls()).toBe(onePass)
      expect(layouts.mock.calls[0][4]).toBeDefined()
      extended.mockRestore(); common.mockRestore()
    })
  }

  it('a width-dependent (General) number refits inside the layout under the same observer', () => {
    const canonical = host()
    const layouts = vi.spyOn(canonical, 'layoutBody')
    const cache = new LayoutCache()
    const gate = admission()
    cache.beginBatch(gate)
    cache.layout(owner, { cell: { ref: 'A1', row: 0, col: 0, value: 1234.56789, styleIndex: 0 }, source: { row: 0, col: 0 }, x: 0, y: 0, width: 40, height: 20 }, canonical)
    cache.endBatch()
    expect(layouts.mock.calls[0][4]).toBeUndefined()
    expect(layouts.mock.calls[0][5]).toBeDefined()
    expect(gate.held()).toBe(0)
  })

  it('both admitted measurement canvases inherit the document language, like painting tiles', () => {
    const root = document.documentElement, previous = root.getAttribute('lang')
    root.setAttribute('lang', 'zh-TW')
    const created: HTMLCanvasElement[] = []
    const make = document.createElement.bind(document)
    const spy = vi.spyOn(document, 'createElement').mockImplementation(((tag: string) => { const el = make(tag); if (tag === 'canvas') created.push(el as HTMLCanvasElement); return el }) as typeof document.createElement)
    try {
      createAdmittedLayoutContext(() => {})
      expect(created).toHaveLength(2)
      expect(created.map(canvas => canvas.getAttribute('lang'))).toEqual(['zh-TW', 'zh-TW'])
    } finally { spy.mockRestore(); if (previous === null) root.removeAttribute('lang'); else root.setAttribute('lang', previous) }
  })
})
