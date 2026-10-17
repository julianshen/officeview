import { expect, it, vi } from 'vitest'
import { SheetTileRenderer, paneSurfaceBytes } from '../src/components/xlsx/sheetTiles'
import { buildSheetGeometry, MergeQueryCache } from '../src/xlsx/view/viewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import type { XlsxSheet } from '../src/xlsx/types'

const viewport = { width: 200, height: 140 }
function frame() {
  const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('ready geometry required')
  return { sheet, geometry, layoutCache: new LayoutCache(), merges: new MergeQueryCache(), identity: { docId: 'ce', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }, origin: { x: 100, y: 100 }, viewport, zoom: 1, dpr: 1, extent: { width: 10000, height: 10000 }, frozen: { rows: 1, cols: 1, left: 64, top: 20 } }
}
function target() { const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 140; return canvas.getContext('2d')! }
const samples = (ctx: CanvasRenderingContext2D) => [[10, 2], [100, 2], [10, 100], [100, 100]].map(([x, y]) => [...ctx.getImageData(x, y, 1, 1).data])

it('one real pending tile is painted through all4clips, then closed once without retention', () => {
  let live = 0, peak = 0, made = 0, closed = 0
  const renderer = new SheetTileRenderer('D', paneSurfaceBytes(viewport, 1), () => {
    const canvas = document.createElement('canvas'); canvas.width = 512; canvas.height = 512
    made++; live++; peak = Math.max(peak, live); let open = true
    return { canvas, width: 512, height: 512, close() { if (open) { open = false; live--; closed++; canvas.width = 0; canvas.height = 0 } } }
  })
  const ctx = target(), draw = vi.spyOn(ctx, 'drawImage')
  try {
    renderer.present(ctx, { ...frame(), pending: box => [box] })
    expect(draw).toHaveBeenCalledTimes(4)
    expect(made).toBe(1); expect(peak).toBe(1); expect(closed).toBe(1); expect(live).toBe(0)
    expect(samples(ctx)).toEqual(Array.from({ length: 4 }, () => [236, 239, 243, 255]))
    expect(renderer.stats.pendingTiles).toBe(1)
  } finally { renderer.dispose() }
})

it('a refused tile draws neutral pixels through all4clips without allocating a surface', () => {
  const make = vi.fn(() => { throw new Error('a refused tile must not allocate') })
  const renderer = new SheetTileRenderer('D', paneSurfaceBytes(viewport, 1), make)
  const ctx = target(), draw = vi.spyOn(ctx, 'drawImage')
  try {
    const result = renderer.present(ctx, { ...frame(), layoutAdmission: { admit: () => false, release: vi.fn(), hold: vi.fn(), cacheChanged: vi.fn() } })
    expect(make).not.toHaveBeenCalled(); expect(draw).not.toHaveBeenCalled()
    expect(samples(ctx)).toEqual(Array.from({ length: 4 }, () => [236, 239, 243, 255]))
    expect(result.resourceNotice).toMatch(/memory/i)
  } finally { renderer.dispose() }
})
