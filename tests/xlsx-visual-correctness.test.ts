/**
 * Visual-review stream B (PR #11): XLSX cell-rendering corrections toward
 * Excel. Paint is recorded through a fixed-metric context so assertions do not
 * depend on host fonts.
 */
import { describe, expect, it } from 'vitest'
import { renderSheet, type GridMetrics } from '../src/xlsx/render'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'

interface PaintOp { op: string; args: unknown[]; fillStyle: unknown; strokeStyle: unknown; lineWidth: unknown; dash: number[]; font: string }
/** A canvas stand-in: every glyph is 7px wide; state at each call is recorded. */
function recordingContext(glyph: (text: string, font: string) => number = text => text.length * 7) {
  const ops: PaintOp[] = []
  const state: Record<string, unknown> = { fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1, font: '10px sans-serif', textBaseline: 'alphabetic' }
  let dash: number[] = []
  const stack: Array<{ state: Record<string, unknown>; dash: number[] }> = []
  const record = (op: string, args: unknown[]) => ops.push({ op, args, fillStyle: state.fillStyle, strokeStyle: state.strokeStyle, lineWidth: state.lineWidth, dash: [...dash], font: String(state.font) })
  const methods: Record<string, (...args: unknown[]) => unknown> = {
    save: () => { stack.push({ state: { ...state }, dash: [...dash] }) },
    restore: () => { const top = stack.pop(); if (top) { Object.assign(state, top.state); dash = top.dash } },
    setLineDash: (d: unknown) => { dash = [...(d as number[])] },
    getLineDash: () => [...dash],
    measureText: (text: unknown) => ({ width: glyph(String(text), String(state.font)) }),
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
  }
  const ctx = new Proxy(state, {
    get(target, key: string) {
      if (key in methods) return methods[key]
      if (key in target) return target[key]
      return (...args: unknown[]) => { record(key, args) }
    },
    set(target, key: string, value) { target[key] = value; return true },
  }) as unknown as CanvasRenderingContext2D
  return { ctx, ops }
}

function sheetOf(cells: Array<Partial<XlsxCell> & { row: number; col: number }>, extra: Partial<XlsxSheet> = {}): XlsxSheet {
  const rows = new Map<number, XlsxCell[]>()
  for (const c of cells) {
    const cell: XlsxCell = { ref: '', styleIndex: 0, value: null, ...c }
    rows.set(c.row, [...(rows.get(c.row) ?? []), cell])
  }
  return { name: 'S', cols: [], merges: [], mergeRanges: [], rows: [...rows].sort((a, b) => a[0] - b[0]).map(([index, cs]) => ({ index, cells: cs })), ...extra }
}
const grid = (cols: number[], rows: number[]): GridMetrics => ({ colWidthsPx: cols, rowHeightsPx: rows, widthPx: cols.reduce((a, b) => a + b, 0), heightPx: rows.reduce((a, b) => a + b, 0) })
const style = (s: Partial<XlsxCellStyle>): XlsxCellStyle => ({ numFmtId: 0, ...s })

describe('#1 SpreadsheetML ARGB alpha byte is ignored (Excel treats 00RRGGBB as opaque)', () => {
  it('paints an openpyxl 00RRGGBB fill and font colour opaque in the static path', () => {
    const { ctx, ops } = recordingContext()
    const sheet = sheetOf([{ row: 0, col: 0, value: 'x', style: style({ fillColor: '00FFF2CC', color: '00FFFFFF' }) }])
    renderSheet(sheet, ctx, grid([64], [20]))
    expect(ops.some(o => o.op === 'fillRect' && o.fillStyle === '#FFF2CC')).toBe(true)
    expect(ops.find(o => o.op === 'fillText')?.fillStyle).toBe('#FFFFFF')
  })
})

describe('#11 worker-served object documents (compatibility mode) keep their cell styles', () => {
  it('the object source publishes the workbook style table, indexed by cellXfs', async () => {
    const bytes = await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', t: 'str', v: 'head', style: 1 }, { ref: 'B1', v: 3.14159, style: 2 }] }] }])
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const source = createObjectWorkbookSource(doc, 'd')
    const formatting = source.formatting?.()
    // The main-thread remote source styles cells only through this table
    // (viewerRemoteSource styleOf); without it fills, fonts and borders vanish.
    expect(formatting?.styles[1]).toMatchObject({ fillColor: 'FFFFFF00', bold: true, borders: { left: 'thin' } })
    expect(formatting?.styles[2]).toMatchObject({ numFmtId: 2 })
    expect(formatting?.styles[1]).toEqual(doc.sheets[0].rows[0].cells[0].style)
  })
})
