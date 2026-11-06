/** Tiny bridge only: a CF-free sheet creates no account, cache or lazy import. */
import type { XlsxSheet } from './types'
type CfFoundation = typeof import('./cf/limits')
let pending: Promise<CfFoundation> | undefined
export function preloadCf(): Promise<CfFoundation> {
  return pending ??= import('./cf/limits').catch(error => { pending = undefined; throw error })
}
export async function ensureCf<T = CfFoundation>(
  sheets: readonly Pick<XlsxSheet, 'conditionalFormatting'>[],
  load: () => Promise<T> = preloadCf as () => Promise<T>,
): Promise<T | undefined> {
  return sheets.some(sheet => sheet.conditionalFormatting !== undefined) ? load() : undefined
}
