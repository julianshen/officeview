/** Tiny eager bridge into the lazy pivots chunk: definitions at parse time (Ruling AD), display regions and buttons. */
import type { XlsxSheet } from './types'
import type { Rect } from './view/textLayout'
import type { PackageReader } from './pivots/load'
import type { RegionStyleMode, StyleRegion } from './tableStyles/regionLayer'
import type { StylesExtras, TableStyleDiagnostic } from './tableStyles/model'
import type { ThemeContext } from '../drawing/style'

export interface PivotEngine {
  regions(sheet: XlsxSheet, extras: StylesExtras, theme: ThemeContext, report: (diagnostic: TableStyleDiagnostic) => void): StyleRegion[]
  paintButtons(ctx: CanvasRenderingContext2D, sheet: XlsxSheet, cellRect: (row: number, col: number) => Rect | undefined, visible: { r0: number; c0: number; r1: number; c1: number }, mode: RegionStyleMode): void
}
let engine: PivotEngine | undefined
let pending: Promise<typeof import('./pivots/index')> | undefined
export function registerPivotEngine(value: PivotEngine): void { engine = value }
function load(): Promise<typeof import('./pivots/index')> {
  return pending ??= import('./pivots/index').catch(error => { pending = undefined; throw error })
}
export function pivotEngine(): PivotEngine | undefined { return engine }
export function ensurePivotEngine(sheets: readonly XlsxSheet[]): Promise<void> {
  return sheets.some(sheet => sheet.pivotTables?.length) ? load().then(() => {}, () => {}) : Promise.resolve()
}

/** Parse time: the definitions of every pivot the sheets locate, kept on their sheets (no part text retained). */
export async function attachPivotTables(pkg: PackageReader, sheets: readonly XlsxSheet[]): Promise<void> {
  const owners = sheets.filter(sheet => sheet.pivotRegions?.length)
  if (!owners.length) return
  try {
    const book = await (await load()).loadPivotWorkbook(pkg, new Map(owners.map(sheet => [sheet.name, sheet.pivotRegions!])))
    for (const sheet of owners) {
      const tables = book.tables.filter(table => table.sheet === sheet.name)
      if (tables.length) sheet.pivotTables = tables
    }
    for (const { kind, feature, message, partPath } of book.diagnostics) {
      const sheet = owners.find(owner => owner.pivotRegions!.some(region => region.partPath === partPath)) ?? owners[0]!
      ;(sheet.diagnostics ??= []).push({ kind, feature, message })
    }
  } catch {
    for (const sheet of owners) (sheet.diagnostics ??= []).push({ kind: 'unsupported', feature: 'pivot-engine-unavailable', message: 'Pivot definitions could not be loaded; pivot cells paint with their stored formatting only' })
  }
}

/** The pivot button pass (after cell content, before drawings); false when the chunk is not loaded. */
export function paintPivotButtonsIn(ctx: CanvasRenderingContext2D, sheet: XlsxSheet, cellRect: (row: number, col: number) => Rect | undefined, visible: { r0: number; c0: number; r1: number; c1: number }, mode: RegionStyleMode): boolean {
  if (!engine) return false
  engine.paintButtons(ctx, sheet, cellRect, visible, mode)
  return true
}
