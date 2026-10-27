import type { XlsxSheet } from './types'
import type { Rect } from './view/textLayout'
import type { SparkScale } from './sparklines/geometry'
import type { SparklinePainter } from './sparklines/paint'

let painter: SparklinePainter | undefined
let pending: Promise<void> | undefined
export function registerSparklinePainter(value: SparklinePainter): void { painter = value }
function load(): Promise<void> {
  return pending ??= import('./sparklines/index').then(() => {}).catch(error => { pending = undefined; throw error })
}
export function preloadSparklines(): Promise<void> { return load() }
export function ensureSparklinePainter(sheets: readonly XlsxSheet[]): Promise<void> {
  return sheets.some(sheet => sheet.resolvedSparklines?.size) ? load().catch(() => {}) : Promise.resolve()
}
export function paintSparklinesIn(ctx: CanvasRenderingContext2D, sheet: XlsxSheet, cellRect: (row: number, col: number) => Rect | undefined, visible: { r0: number; c0: number; r1: number; c1: number }, scale: SparkScale): boolean {
  return painter?.(ctx, sheet, cellRect, visible, scale) ?? false
}
