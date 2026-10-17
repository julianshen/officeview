import { expect, it } from 'vitest'
import { autoFitSize } from '../src/components/xlsx/sheetResize'
import { AUTOFIT_MAX_CELLS, AUTOFIT_MAX_ROWS } from '../src/xlsx/view/axisSizes'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'
const cell = (row: number, col: number, value: string, wrap = false): XlsxCell => ({ ref: '', row, col, value, styleIndex: 0, ...(wrap ? { style: { numFmtId: 0, wrapText: true } } : {}) })
const sheet = (): XlsxSheet => ({ name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] })

it('column auto-fit samples visible pinned rows before the scrolling-body row cap', () => {
  const s = sheet(), title = 'T'.repeat(40), bodyRow = 10_000
  s.rows = [{ index: 0, cells: [cell(0, 0, title)] }, { index: bodyRow, cells: [cell(bodyRow, 0, 'body')] }]
  const geometry = buildSheetGeometry(s)
  if (geometry.kind !== 'ready') throw new Error('Expected ready geometry')
  const pane = { x: 0, y: geometry.metrics.rows.start(bodyRow), width: 64, height: (AUTOFIT_MAX_ROWS + 1) * 20 }
  const pinned = { x: 0, y: 0, width: 64, height: 20 }
  const seen: string[] = []
  const input = { geometry, sheet: s, pane, regions: [pinned, pane, { ...pane, x: 64 }], knownRows: () => [0, bodyRow], measure: (_font: string, text: string) => { seen.push(text); return text.length * 7 } }
  expect(autoFitSize('column', 0, input)).toBe(287)
  expect(seen.filter(t => t === title)).toHaveLength(1)
  expect(seen.filter(t => t === 'body')).toHaveLength(1)
})

it('row auto-fit includes pinned columns in its bounded visible-first cell sample', () => {
  const s = sheet(), bodyCol = 2_000
  const cells = [cell(0, 0, 'pinned', true), ...Array.from({ length: AUTOFIT_MAX_CELLS }, (_, i) => cell(0, bodyCol + i, 'body', true))]
  s.rows = [{ index: 0, cells }]
  const geometry = buildSheetGeometry(s)
  if (geometry.kind !== 'ready') throw new Error('Expected ready geometry')
  const pane = { x: geometry.metrics.cols.start(bodyCol), y: 0, width: AUTOFIT_MAX_CELLS * 64, height: 20 }
  const seen: number[] = []
  const input = { geometry, sheet: s, pane, regions: [{ x: 0, y: 0, width: 64, height: 20 }, pane], knownRows: () => [0], measure: (_font: string, text: string) => text.length * 7, layoutHeight: (c: XlsxCell) => { seen.push(c.col); return c.col === 0 ? 300 : 1 } }
  expect(autoFitSize('row', 0, input)).toBe(306)
  expect(seen).toHaveLength(AUTOFIT_MAX_CELLS)
  expect(seen.filter(col => col === 0)).toHaveLength(1)
})
