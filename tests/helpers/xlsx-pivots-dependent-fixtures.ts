import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { loadAll } from './xlsx-pivots-load'
import { openPivotSession } from '../../src/xlsx/pivots/session/session'
import { sheetCellReader } from '../../src/xlsx/pivots/engine/faithful'
import { pivotItemLabel, type PivotLayout } from '../../src/xlsx/pivots/engine/place'
import { formulaError } from '../../src/xlsx/formula/evaluator'
import type { XlsxDocument, XlsxCell } from '../../src/xlsx/types'
import type { EvaluationValue } from '../../src/xlsx/formula/types'
import type { LivePivotDataContext } from '../../src/xlsx/pivots/session/getpivotdata'

export const F = `${resolve(__dirname, '../fixtures/xlsx-pivots')}/`
export const SEED = F + 'recipes/getpivotdata-seed.xlsx'
export async function load(file = SEED): Promise<XlsxDocument> { return parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file)))) }
export function cellAt(doc: XlsxDocument, sheet: string, ref: string): XlsxCell | undefined {
  return doc.sheets.find(s => s.name === sheet)!.rows.flatMap(r => r.cells).find(c => c.ref === ref)
}
export function addCell(doc: XlsxDocument, sheet: string, ref: string, col: number, row: number, formula: string, value: number | string) {
  const s = doc.sheets.find(x => x.name === sheet)!
  let r = s.rows.find(x => x.index === row)
  if (!r) { r = { index: row, cells: [] }; s.rows.push(r); s.rows.sort((a, b) => a.index - b.index) }
  r.cells.push({ ref, col, row, value, formula, styleIndex: 0, hasCachedValue: true } as XlsxCell)
  r.cells.sort((a, b) => a.col - b.col)
}
/** Open the seed's pivot session; optionally hide one Region item by label; return the committed live context. */
export async function liveContext(hideRegion?: string, file = SEED) {
  const loaded = (await loadAll(file))[0]!
  const opened = openPivotSession({ def: loaded.def, cache: loaded.cache, store: loaded.store, date1904: loaded.date1904, saved: sheetCellReader(loaded.sheet), occupied: () => false, linkedCharts: [] })
  if (!opened.ok) throw new Error('session did not open')
  const session = opened.value
  let layout: PivotLayout = session.layout
  const savedRect = session.layout.rect
  if (hideRegion) {
    const items = loaded.def.fields[0]!.items!
    const k = items.findIndex((_, i) => { const v = pivotItemLabel(loaded.def, loaded.cache, 0, i, loaded.date1904).value; return v.kind === 'text' && v.value === hideRegion })
    const applied = session.apply({ kind: 'items', field: 0, hidden: [k] })
    if (!applied.ok) throw new Error(JSON.stringify(applied.diagnostic))
    if (!session.commit(applied.value)) throw new Error('commit refused')
    layout = applied.value.layout
  }
  const ctx: LivePivotDataContext & { rects: { r0: number; r1: number; c0: number; c1: number }[] } = {
    sheet: loaded.sheet.name, def: session.def, cache: loaded.cache, layout, date1904: loaded.date1904,
    rects: [savedRect, layout.rect].map(r => ({ r0: r.minRow, r1: r.maxRow, c0: r.minCol, c1: r.maxCol })),
  }
  return { ctx, session, loaded, savedRect }
}
/** InteractiveSheetDoc.tsx's readCellValue, verbatim. */
export function viewerReadCellValue(contexts: ReadonlyArray<LivePivotDataContext & { rects: readonly { r0: number; r1: number; c0: number; c1: number }[]; overlay?: import('../../src/xlsx/view/contracts').ValueOverlay }>) {
  return (name: string, col: number, row: number): EvaluationValue | undefined => {
    const overlay = [...contexts].reverse().find(context => context.overlay)?.overlay
    const current = overlay?.cell(row, col)
    if (current !== undefined) {
      if (current === null) return null
      if (current.valueIsError) return formulaError(String(current.value) as Parameters<typeof formulaError>[0])
      return current.value
    }
    for (const context of contexts) {
      if (context.sheet.toLocaleLowerCase() !== name.toLocaleLowerCase()) continue
      const cell = context.layout.cells.find(candidate => candidate.row === row && candidate.col === col)
      if (!cell) {
        if (context.rects.some(rect => row >= rect.r0 && row <= rect.r1 && col >= rect.c0 && col <= rect.c1)) return null
        continue
      }
      if (cell.value.kind === 'empty') return null
      if (cell.value.kind === 'error') return formulaError(cell.value.value as Parameters<typeof formulaError>[0])
      return cell.value.value
    }
    return undefined
  }
}
export const changedOf = (doc: XlsxDocument, sheet: string, rects: readonly { r0: number; r1: number; c0: number; c1: number }[]) =>
  rects.map(r => ({ sheetId: doc.sheets.findIndex(s => s.name === sheet), rect: { minRow: r.r0, minCol: r.c0, maxRow: r.r1, maxCol: r.c1 } }))
