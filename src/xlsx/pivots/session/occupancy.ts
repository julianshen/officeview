/**
 * Ruling S3-AC (Task 29 fix round 1): the structure a growing pivot treats as occupied, unmeasured and so failing
 * closed — every merge, the cells of every array formula and spill range (an extent that cannot be known reaches the
 * sheet's edge), and cells with a comment or note (when a sheet's notes could not be read, every cell). The session
 * asks only about cells outside the pivot's own area, so a pivot's own merges never block it.
 */
import { RectangleIndex } from '../../view/structure'
import { MAX_COLS, MAX_ROWS, type SourceRect } from '../../view/coordinates'
import { cellCoords } from '../stub'
import type { XlsxSheet } from '../../types'

export interface StructuralBlockers { rects: SourceRect[]; notes: readonly number[]; notesUnknown: boolean }
export interface StructuralOccupancy { occupied(row: number, col: number): boolean; readonly bytes: number }

export function structuralOccupancy(blockers: StructuralBlockers): StructuralOccupancy {
  if (blockers.notesUnknown) return { occupied: () => true, bytes: 64 }
  const index = RectangleIndex.build(blockers.rects.map(rect => ({ rect, value: true })))
  const notes = new Set(blockers.notes)
  return {
    occupied: (row, col) => notes.has(row * MAX_COLS + col) || index.first({ r0: row, r1: row, c0: col, c1: col }) !== undefined,
    bytes: 64 + index.estimatedBytes + notes.size * 24,
  }
}

const rangeOf = (ref: string | undefined): SourceRect | undefined => {
  if (!ref) return undefined
  const [first, last = first] = ref.replace(/\$/g, '').split(':')
  const a = cellCoords(first!), b = cellCoords(last)
  return a && b ? { r0: Math.min(a.row, b.row) - 1, r1: Math.max(a.row, b.row) - 1, c0: Math.min(a.col, b.col) - 1, c1: Math.max(a.col, b.col) - 1 } : undefined
}
const toEdge = (row: number, col: number): SourceRect => ({ r0: row, r1: MAX_ROWS - 1, c0: col, c1: MAX_COLS - 1 })

/** A resident sheet's blockers; `spills` are the engine's footprints when it evaluated the sheet (`engineSpillFootprints`). */
export function sheetBlockers(sheet: XlsxSheet, spills?: readonly { anchor: { row: number; col: number }; r0: number; r1: number; c0: number; c1: number }[]): StructuralBlockers {
  const rects: SourceRect[] = sheet.mergeRanges.map(m => ({ r0: m.minRow, r1: m.maxRow, c0: m.minCol, c1: m.maxCol }))
  const engine = new Set((spills ?? []).map(f => f.anchor.row * MAX_COLS + f.anchor.col))
  for (const f of spills ?? []) rects.push({ r0: f.r0, r1: f.r1, c0: f.c0, c1: f.c1 })
  for (const row of sheet.rows) for (const cell of row.cells) {
    if (cell.arrayRef === undefined && !cell.dynamicArray) continue
    const saved = rangeOf(cell.arrayRef)
    if (saved) rects.push(saved)
    else if (!engine.has(cell.row * MAX_COLS + cell.col)) rects.push(toEdge(cell.row, cell.col))
  }
  return { rects, notes: sheet.noteCells ?? [], notesUnknown: sheet.notesUnknown === true }
}

/** A paged sheet's blockers, from its import records. */
export function pagedBlockers(sheet: { mergeRects: readonly SourceRect[]; arrayCells: readonly { row: number; col: number; saved?: SourceRect; dynamic: boolean }[]; noteCells?: readonly number[]; notesUnknown?: boolean }): StructuralBlockers {
  const rects: SourceRect[] = sheet.mergeRects.map(rect => ({ ...rect }))
  for (const a of sheet.arrayCells) rects.push(a.saved ?? (a.dynamic ? toEdge(a.row, a.col) : { r0: a.row, r1: a.row, c0: a.col, c1: a.col }))
  return { rects, notes: sheet.noteCells ?? [], notesUnknown: sheet.notesUnknown === true }
}
