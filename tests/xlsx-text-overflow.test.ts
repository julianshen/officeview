/**
 * S13 Excel fidelity: text overflows only into EMPTY adjacent cells — to the
 * right for left/general alignment, to the left for right alignment, both ways
 * for center — and clips at the first non-empty cell (a value, a formula that
 * returns "", any merged cell) or the sheet edge. Hidden (zero-width) columns
 * are skipped. Merged cells never overflow. Shared by the static renderer and
 * the interactive viewer (layoutCellBody + textOverflowAllowance).
 *
 * Shrink-to-fit (alignment@shrinkToFit, unwrapped cells): the font scales down
 * until the text fits; numbers show '#' only when even the minimum size does
 * not fit. Find/accessible values (cellDisplayText) never change.
 *
 * Fixed 7 px glyphs at 10 pt (scaled by the font size) pin measurement;
 * padding is 3 px a side; default columns here are 50 px.
 */
import { describe, expect, it } from 'vitest'
import { cellDisplayText, layoutCellBody, renderSheet, rowOverflowFromCells, textMeasurer, textOverflowAllowance } from '../src/xlsx/render'
import { formatA1, MAX_COLS } from '../src/xlsx/view/coordinates'
import type { XlsxCell, XlsxCellStyle, XlsxMergeRange, XlsxSheet } from '../src/xlsx/types'

const W = 50
interface Painted { text: string; font: string; clip: { x0: number; x1: number } | null }

/** A 2D context stub that records each fillText with its font and the
 * horizontal clip in force (axis-aligned rects only; no transforms used). */
function recorder() {
  const painted: Painted[] = []
  let clip: Painted['clip'] = null
  let pending: Painted['clip'] = null
  const stack: Array<Painted['clip']> = []
  const state = {
    font: '',
    measureText: (s: string) => ({ width: s.length * 7 * (Number(/([\d.]+)pt/.exec(state.font)?.[1] ?? 10) / 10) }),
    fillText: (s: string) => painted.push({ text: s, font: state.font, clip }),
    save: () => stack.push(clip),
    restore: () => { clip = stack.pop() ?? null },
    rect: (x: number, _y: number, w: number) => { pending = { x0: x, x1: x + w } },
    clip: () => { if (pending) clip = clip ? { x0: Math.max(clip.x0, pending.x0), x1: Math.min(clip.x1, pending.x1) } : pending },
  }
  const ctx = new Proxy(state, { get: (o, k) => k in o ? o[k as keyof typeof o] : () => {} }) as unknown as CanvasRenderingContext2D
  return { ctx, painted }
}

function cellAt(col: number, value: XlsxCell['value'], extra: Partial<XlsxCell> = {}): XlsxCell {
  return { ref: '', row: 0, col, styleIndex: 0, value, ...extra }
}
function sheetOf(cells: XlsxCell[], merges: XlsxMergeRange[] = [], cols = 6, hidden: number[] = []): { sheet: XlsxSheet; widths: number[] } {
  const widths = Array.from({ length: cols }, (_, c) => hidden.includes(c) ? 0 : W)
  return { sheet: { name: 'S', cols: [], merges: [], mergeRanges: merges, rows: [{ index: 0, cells }] }, widths }
}
function paint(cells: XlsxCell[], merges: XlsxMergeRange[] = [], hidden: number[] = []): Painted[] {
  const { sheet, widths } = sheetOf(cells, merges, 6, hidden)
  const out = recorder()
  const width = widths.reduce((a, b) => a + b, 0)
  renderSheet(sheet, out.ctx, { colWidthsPx: widths, rowHeightsPx: [20, 20], widthPx: width, heightPx: 40 })
  return out.painted
}
const LONG = 'abcdefghijklmnopq' // 17 glyphs = 119 px at 10 pt
const textOf = (painted: Painted[], text: string) => painted.find(p => p.text === text)!

describe('text overflow into empty neighbours only (static renderer)', () => {
  it('left/general text overflows across empty cells and clips at the first value', () => {
    const p = textOf(paint([cellAt(0, LONG), cellAt(2, 'x')]), LONG)
    expect(p.clip).toEqual({ x0: 0, x1: 100 }) // A and empty B; C holds "x"
  })
  it('a value in the next cell clips at the cell edge', () => {
    expect(textOf(paint([cellAt(0, LONG), cellAt(1, 'x')]), LONG).clip).toEqual({ x0: 0, x1: 50 })
  })
  it('a formula that returns "" blocks like a value', () => {
    expect(textOf(paint([cellAt(0, LONG), cellAt(1, '', { formula: '""' })]), LONG).clip).toEqual({ x0: 0, x1: 50 })
  })
  it('a cell holding an empty-string constant blocks too (Excel: not blank)', () => {
    expect(textOf(paint([cellAt(0, LONG), cellAt(1, '')]), LONG).clip).toEqual({ x0: 0, x1: 50 })
  })
  it('a merged neighbour blocks, even when its anchor is in another row', () => {
    const merge: XlsxMergeRange = { minRow: 0, maxRow: 1, minCol: 1, maxCol: 2 }
    expect(textOf(paint([cellAt(0, LONG)], [merge]), LONG).clip).toEqual({ x0: 0, x1: 50 })
  })
  it('text that fits its empty overflow run is not clipped', () => {
    expect(textOf(paint([cellAt(0, LONG)]), LONG).clip).toBeNull() // B..F empty: 300 px of room
  })
  it('hidden columns are skipped: they neither block nor add width', () => {
    // B hidden (holding a value), C empty, D holds "x": the run is C only.
    const p = textOf(paint([cellAt(0, LONG), cellAt(1, 'hidden value'), cellAt(3, 'x')], [], [1]), LONG)
    expect(p.clip).toEqual({ x0: 0, x1: 100 }) // A (0-50) + hidden B (0 px) + C (50-100)
  })
  it('formatting-only cells (no value, no formula) are empty and do not block', () => {
    expect(textOf(paint([cellAt(0, LONG), cellAt(1, null, { style: { numFmtId: 0, fillColor: 'FFFF00' } as XlsxCellStyle }), cellAt(2, 'x')]), LONG).clip).toEqual({ x0: 0, x1: 100 })
  })
})

describe('textOverflowAllowance', () => {
  const colStart = (c: number) => c * W
  const cellsAt = (...cols: number[]) => cols.map(col => ({ col, value: 'v' as XlsxCell['value'] }))
  const input = (cells: Array<{ col: number; value: XlsxCell['value'] }> = [], merges: Array<{ minCol: number; maxCol: number }> = []) => rowOverflowFromCells(cells, merges, colStart, MAX_COLS - 1)
  it('right alignment overflows left until the previous value; center both ways', () => {
    expect(textOverflowAllowance(input(cellsAt(0, 2, 5)), 2)).toEqual({ left: 50, right: 100 }) // B empty; D,E empty, F value
    expect(textOverflowAllowance(input(cellsAt(1, 2, 3)), 2)).toEqual({ left: 0, right: 0 })
  })
  it('formatting-only cells in between are skipped, not blockers (unsorted input too)', () => {
    const cells = [{ col: 4, value: 'x' as XlsxCell['value'] }, { col: 1, value: null }, { col: 2, value: 'v' as XlsxCell['value'] }, { col: 3, value: null }]
    expect(textOverflowAllowance(input(cells), 2)).toEqual({ left: 100, right: 50 })
  })
  it('duplicate source columns do not block right-aligned text from overflowing left', () => {
    const cells = [cellAt(2, LONG, { style: { numFmtId: 0, horizontal: 'right' } as XlsxCellStyle }), cellAt(2, 'duplicate'), cellAt(4, 'blocker')]
    expect(textOverflowAllowance(input(cells), 2)).toEqual({ left: 100, right: 50 })
    const { ctx } = recorder()
    const layout = layoutCellBody(ctx, sheetOf(cells).sheet, cells[0], 100, 0, 50, 20, textMeasurer(ctx), textOverflowAllowance(input(cells), 2))
    const clip = layout.text?.kind === 'layout' ? layout.text.clip : null
    expect(clip!.x).toBeLessThan(100)
    expect(clip!.x + clip!.width).toBeCloseTo(150, 9)
  })
  it('unsorted repeated columns use the nearest strictly different blocker on both sides', () => {
    expect(textOverflowAllowance(input(cellsAt(5, 2, 0, 2, 5, 0, 2)), 2)).toEqual({ left: 50, right: 100 })
  })
  it('stops at the worksheet edges (A on the left, the last column on the right)', () => {
    expect(textOverflowAllowance(input(), MAX_COLS - 1).right).toBe(0)
    expect(textOverflowAllowance(input(), 0).left).toBe(0)
    expect(textOverflowAllowance(input(), 3).left).toBe(150)
  })
  it('a merged cell never overflows', () => {
    expect(textOverflowAllowance(input([], [{ minCol: 2, maxCol: 3 }]), 2)).toEqual({ left: 0, right: 0 })
  })
})

/** Review round 4 finding 2 (adapted from the reviewer's repro): an allowance
 * must not scan formatting-only cells. Counts the index reads one allowance
 * performs on its (once-per-row) input. */
describe('text-overflow allowance cost on a wide styled row', () => {
  const lastCol = 16_384
  function readsFor(cells: Array<{ col: number; value: XlsxCell['value'] }>, at: number): number {
    const built = rowOverflowFromCells(cells, [], (c: number) => c * 64, lastCol)
    let reads = 0
    const counted = new Proxy(built.blockers as number[], {
      get: (t, k) => { if (typeof k === 'string' && /^\d+$/.test(k)) reads++; return (t as unknown as Record<string, unknown>)[k as never] },
    })
    textOverflowAllowance({ ...built, blockers: counted }, at)
    return reads
  }
  it('one allowance does not grow with the formatting-only cells before the blocker', () => {
    const counts: number[] = []
    for (const n of [1_000, 2_000, 4_000, 8_000, 16_383]) {
      const cells: Array<{ col: number; value: XlsxCell['value'] }> = Array.from({ length: n }, (_, i) => ({ col: i, value: null }))
      cells.push({ col: n, value: 'blocker' })
      counts.push(readsFor(cells, 0))
    }
    expect(counts[counts.length - 1] / counts[0], `reads ${counts.join(', ')}`).toBeLessThan(4)
  })
  it('a whole viewport of text cells in a styled row stays bounded', () => {
    const styled: Array<{ col: number; value: XlsxCell['value'] }> = Array.from({ length: lastCol - 1 }, (_, i) => ({ col: i + 1, value: null }))
    const cells = [...styled, { col: lastCol - 1, value: 'blocker' as XlsxCell['value'] }]
    for (let k = 0; k < 64; k++) cells[k * 16] = { col: k * 16 + 1, value: 'text' }
    let total = 0
    for (let k = 0; k < 64; k++) total += readsFor(cells, k * 16 + 1)
    expect(total, `64 allowances cost ${total} index reads`).toBeLessThan(64 * 32)
  })
})

describe('explicit alignment layouts carry the allowance as their clip', () => {
  const measure = (ctx: CanvasRenderingContext2D) => textMeasurer(ctx)
  it('right-aligned text clips at the previous value; centered text at both blockers', () => {
    const { ctx } = recorder()
    const sheet = sheetOf([]).sheet
    const right = layoutCellBody(ctx, sheet, cellAt(2, LONG, { style: { numFmtId: 0, horizontal: 'right' } as XlsxCellStyle }), 100, 0, 50, 20, measure(ctx), { left: 50, right: 0 })
    expect(right.text?.kind === 'layout' && right.text.clip).toMatchObject({ x: 50, width: 100 })
    const center = layoutCellBody(ctx, sheet, cellAt(1, LONG, { style: { numFmtId: 0, horizontal: 'center' } as XlsxCellStyle }), 50, 0, 50, 20, measure(ctx), { left: 50, right: 0 })
    const clip = center.text?.kind === 'layout' ? center.text.clip : null
    // Left room (50 px) covers the left half; the right side is blocked at the cell edge.
    expect(clip!.x).toBeGreaterThanOrEqual(0)
    expect(clip!.x).toBeLessThan(50)
    expect(clip!.x + clip!.width).toBeCloseTo(100, 9)
  })
})

describe('review round 4 regressions', () => {
  it('each cell of a malformed row uses the merges of its OWN source row', () => {
    // <row r="1"> holding C3 (another row) first, then A1. B1:B2 is merged.
    const merge: XlsxMergeRange = { minRow: 0, maxRow: 1, minCol: 1, maxCol: 1 }
    const cells = [{ ...cellAt(2, LONG), ref: 'C3', row: 2 }, { ...cellAt(0, LONG), ref: 'A1', row: 0 }]
    const sheet: XlsxSheet = { name: 'S', cols: [], merges: [], mergeRanges: [merge], rows: [{ index: 0, cells }] }
    const out = recorder()
    renderSheet(sheet, out.ctx, { colWidthsPx: Array(6).fill(W), rowHeightsPx: [20, 20, 20], widthPx: 300, heightPx: 60 })
    const a1 = out.painted.filter(p => p.text === LONG).find(p => p.clip?.x0 === 0)
    expect(a1?.clip).toEqual({ x0: 0, x1: 50 }) // blocked by the B1:B2 merge on row 1
  })
})

describe('review round 5: mixed-source grouping cost', () => {
  it.each([128, 256, 512])('partitions %i diagonal explicit refs in linear source-ref reads', count => {
    let reads = 0
    const cells = Array.from({ length: count }, (_, row) => {
      const ref = formatA1({ row, col: row })
      return { ...cellAt(row, LONG), row, get ref() { reads++; return ref } }
    })
    const sheet: XlsxSheet = { name: 'S', cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells }] }
    const out = recorder()
    renderSheet(sheet, out.ctx, { colWidthsPx: Array(count).fill(4), rowHeightsPx: Array(count).fill(4), widthPx: count * 4, heightPx: count * 4 })
    expect(out.painted).toHaveLength(count)
    expect(reads, `${count} cells cost ${reads} source-ref reads`).toBeLessThanOrEqual(count * 6)
  })
})

describe('review round 3 regressions', () => {
  it('a numeric value is judged by its selected section, not by an "@" in another section', () => {
    expect(paint([cellAt(0, 123456789, { style: { numFmtId: 164, formatCode: '0.00;[Red]-0.00;0;@' } as XlsxCellStyle })]).map(p => p.text)).toEqual(['######'])
  })
  it('a quoted or escaped "#" literal that fits is painted as authored', () => {
    expect(paint([cellAt(0, 42, { style: { numFmtId: 164, formatCode: '"#"' } as XlsxCellStyle })]).map(p => p.text)).toEqual(['#'])
    expect(paint([cellAt(0, 42, { style: { numFmtId: 164, formatCode: '\\#' } as XlsxCellStyle })]).map(p => p.text)).toEqual(['#'])
  })
  it('a number shown through a text ("@") section keeps text behaviour', () => {
    expect(paint([cellAt(0, 123456789, { style: { numFmtId: 49, formatCode: '@' } as XlsxCellStyle })]).map(p => p.text)).not.toEqual(['######'])
  })
  it('fill alignment never overflows', () => {
    const { ctx } = recorder()
    const fill = layoutCellBody(ctx, sheetOf([]).sheet, cellAt(0, LONG, { style: { numFmtId: 0, horizontal: 'fill' } as XlsxCellStyle }), 0, 0, 50, 20, textMeasurer(ctx), { left: 0, right: 250 })
    expect(fill.text?.kind === 'layout' && fill.text.clip).toMatchObject({ x: 0, width: 50 })
  })
  it('overflow text is painted after every fill, so a formatted empty neighbour does not erase it', () => {
    const ops: string[] = []
    const { sheet, widths } = sheetOf([cellAt(0, LONG), cellAt(1, null, { style: { numFmtId: 0, fillColor: 'FFFF00' } as XlsxCellStyle })])
    const state = { font: '', fillStyle: '', measureText: (t: string) => ({ width: t.length * 7 }), fillRect: (x: number) => ops.push(`fillRect@${x}:${state.fillStyle}`), fillText: (t: string) => ops.push(`text:${t}`) }
    const ctx = new Proxy(state, { get: (o, k) => k in o ? o[k as keyof typeof o] : () => {} }) as unknown as CanvasRenderingContext2D
    renderSheet(sheet, ctx, { colWidthsPx: widths, rowHeightsPx: [20], widthPx: 300, heightPx: 20 })
    const fillB = ops.findIndex(op => op.startsWith('fillRect@50:') && !op.endsWith('#ffffff'))
    expect(fillB).toBeGreaterThan(-1)
    expect(ops.indexOf(`text:${LONG}`)).toBeGreaterThan(fillB)
  })
})

describe('shrink to fit', () => {
  const shrink = (extra: Partial<XlsxCellStyle> = {}) => ({ style: { numFmtId: 0, shrinkToFit: true, ...extra } as XlsxCellStyle })
  const pt = (font: string) => Number(/([\d.]+)pt/.exec(font)?.[1])
  it('a number that would show # shrinks instead', () => {
    const p = paint([cellAt(0, -12345.67, shrink({ numFmtId: 4 }))])
    expect(p.map(x => x.text)).toEqual(['-12,345.67'])
    expect(pt(p[0].font)).toBeLessThan(10)
    expect('-12,345.67'.length * 7 * pt(p[0].font) / 10).toBeLessThanOrEqual(W - 6 + 1e-9)
  })
  it('a date shrinks to fit', () => {
    const p = paint([cellAt(0, 45079, shrink({ formatCode: 'yyyy-mm-dd' }))])
    expect(p.map(x => x.text)).toEqual(['2023-06-02'])
    expect(pt(p[0].font)).toBeLessThan(10)
  })
  it('text shrinks instead of overflowing', () => {
    const p = paint([cellAt(0, 'abcdefghij', shrink())]) // 70 px at 10 pt -> fits 44 px at 6.2 pt
    expect(p[0].text).toBe('abcdefghij')
    expect(pt(p[0].font)).toBeCloseTo(6.2, 9)
    expect(p[0].clip).toBeNull()
  })
  it('text that cannot fit even at the floor is held to its cell, never spilled', () => {
    const p = paint([cellAt(0, LONG, shrink())]) // would need 3.6 pt
    expect(pt(p[0].font)).toBe(4)
    expect(p[0].clip).toEqual({ x0: 0, x1: 50 })
  })
  it('an authored size below the floor is never enlarged', () => {
    const fits = paint([cellAt(0, 'abc', shrink({ fontSizePt: 2 }))])
    expect(pt(fits[0].font)).toBe(2)
    const tooLong = paint([cellAt(0, 'x'.repeat(40), shrink({ fontSizePt: 2 }))]) // 56 px at 2 pt
    expect(pt(tooLong[0].font)).toBeLessThanOrEqual(2)
    expect(tooLong[0].clip).toEqual({ x0: 0, x1: 50 })
  })
  it('numbers and dates that cannot fit at the floor show #', () => {
    const { sheet } = sheetOf([cellAt(0, 45079, shrink({ formatCode: 'yyyy-mm-dd' }))])
    const out = recorder()
    renderSheet(sheet, out.ctx, { colWidthsPx: [20], rowHeightsPx: [20], widthPx: 20, heightPx: 20 }) // 14 px interior; date needs 3.0 pt
    expect(out.painted.map(x => x.text)).toEqual(['##'])
  })
  it('text that already fits keeps its size', () => {
    expect(pt(paint([cellAt(0, 'ab', shrink())])[0].font)).toBe(10)
  })
  it('below the minimum readable size a number falls back to #', () => {
    // 6 px column: interior 0 px. Nothing fits at any size.
    const { sheet } = sheetOf([cellAt(0, 123456789, shrink({ numFmtId: 3 }))])
    const out = recorder()
    renderSheet(sheet, out.ctx, { colWidthsPx: [12], rowHeightsPx: [20], widthPx: 12, heightPx: 20 })
    expect(out.painted.map(x => x.text)).toEqual(['#'])
  })
  it('the accessible / find value is unchanged', () => {
    expect(cellDisplayText(sheetOf([]).sheet, { value: -12345.67, style: { numFmtId: 4, shrinkToFit: true } as XlsxCellStyle })).toBe('-12,345.67')
  })
})
