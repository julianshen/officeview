import { describe, expect, it } from 'vitest'
import { SheetTileRenderer } from '../src/components/xlsx/sheetTiles'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { createDocumentOverflowCache } from '../src/xlsx/view/documentOverflow'
import { bindOverflowCache } from '../src/xlsx/view/overflowCache'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { buildSheetGeometry, MergeQueryCache, planViewport } from '../src/xlsx/view/viewport'
import { workbookOf } from './helpers/xlsx-find'

function refused() {
  const sheet = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'known text' }] }]).sheets[0]
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw Error('geometry')
  const ledger = new MainMemoryLedger(), cache = createDocumentOverflowCache(ledger, 'D')
  bindOverflowCache(geometry, cache)
  ledger.set('copy', 'in-flight', MAIN_THREAD_BUDGETS.D - 1000)
  return { sheet, geometry, ledger, cache }
}

function ink(ctx: CanvasRenderingContext2D, y: number, height: number): number {
  const data = ctx.getImageData(0, y, 200, height).data
  let count = 0
  for (let i = 0; i < data.length; i += 4) if (Math.min(data[i], data[i + 1], data[i + 2]) < 128) count++
  return count
}

describe('round 6: overflow resource rendering keeps exact and unknown coverage', () => {
  it('keeps the exact active record and counts actual coarse groups after cold-cache refusal', () => {
    const f = refused()
    const plan = planViewport(f.geometry, { box: { x: 0, y: 0, width: 200, height: 80 }, zoom: 1, deviceScale: 1 }, 'D', { activeCell: { row: 0, col: 0 } })
    expect(plan.kind).toBe('coarse')
    if (plan.kind !== 'coarse') throw Error('expected resource coarse')
    expect(plan.diagnostics.some(d => d.code === 'overflow-memory')).toBe(true)
    expect(plan.active).toMatchObject({ source: { row: 0, col: 0 }, cell: { value: 'known text' }, x: 0, y: 0, width: 64, height: 20 })
    expect(plan.work.semanticNodes).toBe(1)
    expect(plan.groups.length).toBeGreaterThan(0)
    expect(plan.work.overviewGroups).toBe(plan.groups.length)
    expect(plan.work.records).toBe(0)
    expect(f.ledger.snapshot().byPool.layout).toBe(0)
  })

  it('paints pending coverage and Loading on a refused real tile, reports memory notice, and retains no tile', () => {
    const f = refused(), texts: string[] = []
    const renderer = new SheetTileRenderer('D', 0, () => {
      const canvas = document.createElement('canvas')
      canvas.width = 512; canvas.height = 512
      const ctx = canvas.getContext('2d')!, draw = ctx.fillText.bind(ctx)
      ctx.fillText = (text, x, y) => { texts.push(text); draw(text, x, y) }
      return { canvas, width: 512, height: 512, close() { canvas.width = 0; canvas.height = 0 } }
    })
    const pane = document.createElement('canvas')
    pane.width = 200; pane.height = 80
    const ctx = pane.getContext('2d')!
    try {
      const result = renderer.present(ctx, { sheet: f.sheet, geometry: f.geometry, identity: { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }, origin: { x: 0, y: 0 }, viewport: { width: 200, height: 80 }, zoom: 1, dpr: 1, extent: { width: f.geometry.metrics.cols.totalCss(), height: f.geometry.metrics.rows.totalCss() }, layoutCache: new LayoutCache(), merges: new MergeQueryCache(), pending: () => [{ x: 0, y: 20, width: 512, height: 20 }] })
      expect(result.resourceNotice).toContain('exceeds the document memory budget')
      expect(texts).toContain('Loading…')
      expect(renderer.stats.pendingTiles).toBe(1)
      expect(ink(ctx, 8, 12)).toBeGreaterThan(0)
      expect(ink(ctx, 25, 13)).toBeGreaterThan(0)
      expect(renderer.cache.stats().tiles).toBe(0)
      expect(f.cache.bytes).toBe(0)
    } finally { renderer.dispose(); f.cache.close() }
  })
})
