import { expect, it } from 'vitest'
import { fireEvent, render } from '@testing-library/react'
import { createCanvas } from 'canvas'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { gridFixture } from './helpers/xlsx-grid-fixture'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { createScrollWindow, hitTestViewport, logicalToViewport, setScrollOffset } from '../src/xlsx/view/scrollWindow'
import { emptyGridSelection, selectCell, rowContains } from '../src/xlsx/view/selection'
import { buildSheetGeometry, MergeQueryCache } from '../src/xlsx/view/viewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { SheetTileRenderer } from '../src/components/xlsx/sheetTiles'
import type { XlsxSheet } from '../src/xlsx/types'

const visibility = visibilityFromHidden(createIntervalSet([]))
const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 100, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: 100, defaultSizeCss: 64 }) })
const viewport = { width: 640, height: 400 }
const selectionContains = (selection: ReturnType<typeof emptyGridSelection>, at: { row: number; col: number }) => selection.ranges.some(part => rowContains(part.rows, at.row) && part.cols.some(range => at.col >= range.c0 && at.col <= range.c1))

it('shift-click from a pinned corner to the scrolling body uses the source endpoints', () => {
  const window = setScrollOffset(createScrollWindow(projection, viewport, { frozen: { rows: 1, cols: 1 } }), projection, viewport, { x: 300, y: 400 })
  const selection = selectCell(emptyGridSelection(), projection, { row: 0, col: 0 })
  let next = selection
  const ui = render(<GridSemantics {...gridFixture({ projection, viewport, window, selection, view: { baseHidden: createIntervalSet([]), visibility, filter: null, sort: null }, onChange: t => { next = t.selection } })} />)
  fireEvent.click(ui.getByRole('grid'), { clientX: 44 + 70, clientY: 44 + 30, shiftKey: true, detail: 1 })
  expect(next.active).toEqual({ row: 21, col: 5 })
  expect(selectionContains(next, { row: 20, col: 4 })).toBe(true)
  expect(selectionContains(next, { row: 22, col: 5 })).toBe(false)
  ui.unmount()
})

it('visible frozen and body cell centers round-trip with variable dimensions at fractional zoom', () => {
  const metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 100, defaultSizeCss: 20, overrides: [{ start: 0, end: 0, sizeCss: 26 }] }), cols: createAxisMetrics({ logicalCount: 100, defaultSizeCss: 64, overrides: [{ start: 0, end: 0, sizeCss: 78 }] }) })
  for (const zoom of [0.5, 0.83, 1.25, 2]) {
    const window = setScrollOffset(createScrollWindow(metrics, viewport, { zoom, frozen: { rows: 1, cols: 1 } }), metrics, viewport, { x: 300, y: 400 })
    for (const at of [{ row: 0, col: 0 }, { row: 0, col: 7 }, { row: 25, col: 0 }, { row: 25, col: 7 }]) {
      const box = metrics.sourceCellToLogical(at)!
      const point = logicalToViewport(window, { x: box.x + box.width / 2, y: box.y + box.height / 2 })
      expect(hitTestViewport(window, metrics, viewport, point)).toEqual(at)
    }
  }
})

function paint(sheet: XlsxSheet, offset = { x: 32, y: 10 }, freeze = { rows: 1, cols: 1, left: 64, top: 20 }, zoom = 1) {
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('Expected ready geometry')
  const canvas = createCanvas(viewport.width, viewport.height)
  const renderer = new SheetTileRenderer('D', 0, () => {
    const tile = createCanvas(512, 512)
    return { canvas: tile as unknown as HTMLCanvasElement, width: 512, height: 512, close: () => { tile.width = 0; tile.height = 0 } }
  })
  renderer.present(canvas.getContext('2d') as unknown as CanvasRenderingContext2D, { sheet, geometry, identity: { docId: 'review', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }, origin: offset, viewport, zoom, dpr: 1, extent: { width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() }, frozen: freeze, layoutCache: new LayoutCache(), merges: new MergeQueryCache() })
  renderer.dispose()
  return (x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data)
}

it('a merge crossing both splits paints in all four clipped regions', () => {
  const pixels = paint({ name: 'S', rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, value: '', styleIndex: 0, style: { numFmtId: 0, fillColor: '#ff0000' } }] }], cols: [], merges: ['A1:C3'], mergeRanges: [{ minRow: 0, minCol: 0, maxRow: 2, maxCol: 2 }] })
  for (const point of [[10, 10], [100, 10], [10, 30], [100, 30]]) expect(pixels(...point as [number, number])).toEqual([255, 0, 0, 255])
  expect(pixels(201, 31)).toEqual([255, 255, 255, 255])
})

it('a drawing crossing both splits paints in all four clipped regions', () => {
  const pixels = paint({ name: 'S', rows: [], cols: [], merges: [], mergeRanges: [], drawings: [{ xEmu: 0, yEmu: 0, widthEmu: 192 * 9525, heightEmu: 60 * 9525, geometry: 'rect', fill: '#00ff00', source: { partPath: 'd', treePath: 's', element: 'sp', representation: 'native' } }] })
  for (const point of [[10, 10], [100, 10], [10, 30], [100, 30]]) expect(pixels(...point as [number, number])).toEqual([0, 255, 0, 255])
  expect(pixels(201, 31)).toEqual([255, 255, 255, 255])
})

it('reusing tiles across frozen regions does not erase detail on a pane with 4000 visible cells', () => {
  const sheet: XlsxSheet = { name: 'Dense', cols: [{ min: 0, max: 131, widthChars: 11 / 7 }, { min: 63, max: 68, widthChars: 0 }], merges: [], mergeRanges: [], rows: Array.from({ length: 64 }, (_, row) => ({ index: row, customHeight: true, heightPt: 12, cells: Array.from({ length: 126 }, (_, offset) => { const col = offset < 63 ? offset : offset + 6; return { row, col, ref: '', value: 1, styleIndex: 0, style: { numFmtId: 0, fillColor: '#ff0000', color: '#ff0000' } } }) })) }
  const plain = paint(sheet, { x: 0, y: 0 }, { rows: 0, cols: 0, left: 0, top: 0 }, 0.5)
  expect(plain(3, 100)).toEqual([255, 0, 0, 255])
  expect(plain(3, 3)).toEqual([255, 0, 0, 255])
  const pixels = paint(sheet, { x: 0, y: 0 }, { rows: 1, cols: 1, left: 16, top: 16 }, 0.5)
  expect(pixels(100, 100)).toEqual([255, 0, 0, 255])
  expect.soft(pixels(3, 100)).toEqual([255, 0, 0, 255])
  expect.soft(pixels(3, 3)).toEqual([255, 0, 0, 255])
})
