/** Tiny eager bridge into the table-style painter chunk. */
import type { XlsxSheet } from './types'
import type { RegionStyleLayer } from './tableStyles/regionLayer'
import { sheetHasRegionStyles } from './pivots/stub'

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
  for (const sheet of sheets) {
    if (!sheetHasRegionStyles(sheet)) continue
    if (sheet.regionStylesResolved) continue
    sheet.regionStyles = factory?.(sheet)
    sheet.regionStylesResolved = true
  }
}

export function isTableStyleEngineLoaded(): boolean { return factory !== undefined }

/** Synchronous static renderer access: only a prior preload/getPaintables call can enable it. */
export function regionStylesFor(sheet: XlsxSheet): RegionStyleLayer | undefined {
  if (!sheetHasRegionStyles(sheet)) return undefined
  if (sheet.regionStylesResolved) return sheet.regionStyles
  if (!factory) return undefined
  sheet.regionStyles = factory(sheet)
  sheet.regionStylesResolved = true
  return sheet.regionStyles
}
