/** Tiny eager bridge into the table-style painter chunk. */
import type { XlsxSheet } from './types'
import type { RegionStyleLayer } from './tableStyles/regionLayer'
import { sheetHasRegionStyles } from './pivots/stub'
import { ensurePivotEngine, pivotEngine } from './pivotLoader'

type RegionFactory = (sheet: XlsxSheet) => RegionStyleLayer | undefined
let factory: RegionFactory | undefined
let pending: Promise<void> | undefined

export function registerTableStyleEngine(value: RegionFactory): void { factory = value }
function load(): Promise<void> {
  return pending ??= import('./tableStyles/regionLayer').then(() => {}).catch(error => { pending = undefined; throw error })
}
export function preloadTableStyles(): Promise<void> { return load() }
export async function ensureTableStyleEngine(sheets: readonly XlsxSheet[]): Promise<void> {
  if (!sheets.some(sheetHasRegionStyles)) return
  try { await preloadTableStyles() } catch { return }
  // Pivot regions join the layer from the pivots chunk; a failed load leaves the tables.
  await ensurePivotEngine(sheets)
  for (const sheet of sheets) {
    if (!sheetHasRegionStyles(sheet)) continue
    if (sheet.regionStylesResolved) continue
    sheet.regionStyles = build(sheet)
    sheet.regionStylesResolved = true
  }
}

/** The factory, failing closed per sheet: a throw leaves that sheet unstyled with one diagnostic, never the workbook. */
function build(sheet: XlsxSheet): RegionStyleLayer | undefined {
  try { return factory?.(sheet) } catch (cause) {
    const pivot = Boolean(sheet.pivotTables?.length)
    const feature = pivot ? 'pivot-display-failed' : 'table-style-failed'
    const detail = cause instanceof Error ? `: ${cause.message.slice(0, 128)}` : ''
    const list = (sheet.diagnostics ??= [])
    if (!list.some(item => item.feature === feature)) {
      list.push({ kind: 'unsupported', feature, message: `${pivot ? 'Pivot' : 'Table'} styles of "${sheet.name}" could not be built${detail}; cells paint with their stored formatting only.` })
    }
    sheet.regionStylesResolved = true
    return undefined
  }
}

export function isTableStyleEngineLoaded(): boolean { return factory !== undefined }

/** Synchronous static renderer access: only a prior preload/getPaintables call can enable it. */
export function regionStylesFor(sheet: XlsxSheet): RegionStyleLayer | undefined {
  if (!sheetHasRegionStyles(sheet)) return undefined
  if (sheet.regionStylesResolved) return sheet.regionStyles
  if (!factory) return undefined
  sheet.regionStyles = build(sheet)
  // Without the pivots chunk a pivot sheet's layer is tables-only: not final. A failed build is final.
  sheet.regionStylesResolved ||= !sheet.pivotTables?.length || pivotEngine() !== undefined
  return sheet.regionStyles
}
