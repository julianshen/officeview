/**
 * Excel's column-width model (ECMA-376 §18.3.1.13, §18.3.1.81): widths are
 * stored in characters of the workbook default font's maximum digit width
 * (MDW), padding included, and convert to pixels as
 * trunc(((256 * width + trunc(128 / mdw)) / 256) * mdw). Calibri 11 at 96 dpi
 * has an MDW of 7px.
 */

/** Column pixel model of one sheet: the workbook MDW and the sheet's default width. */
export interface ColumnMetrics { maxDigitWidthPx: number; defaultWidthPx: number }
export const DEFAULT_COLUMN_METRICS: ColumnMetrics = Object.freeze({ maxDigitWidthPx: 7, defaultWidthPx: 64 })

/** Digit advance (em) of common default faces; others use Calibri's. */
const DIGIT_EM: Array<[RegExp, number]> = [
  [/^(arial|helvetica|liberation sans|arimo)/i, 1139 / 2048],
  [/^(times|liberation serif|tinos)/i, 1024 / 2048],
  [/^(cambria|caladea)/i, 1137 / 2048],
]
const CALIBRI_DIGIT_EM = 1038 / 2048

/** The default font's maximum digit width in whole pixels at 96 dpi. */
export function maxDigitWidthPx(fontName?: string, sizePt = 11): number {
  const em = DIGIT_EM.find(([pattern]) => pattern.test(fontName ?? ''))?.[1] ?? CALIBRI_DIGIT_EM
  const size = Number.isFinite(sizePt) && sizePt > 0 ? sizePt : 11
  return Math.max(1, Math.round(em * size * 96 / 72))
}

/** A stored column width (characters incl. padding) in pixels. Invalid
 * (negative or non-finite) widths stay invalid for the callers' diagnostics. */
export function columnWidthPx(width: number, mdw: number): number {
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw)
}

/** The sheet's default column width: defaultColWidth when authored, else
 * baseColWidth characters plus padding, which Excel rounds up to a multiple of 8px. */
export function defaultColumnWidthPx(format: { baseColWidth?: number; defaultColWidth?: number }, mdw: number): number {
  if (format.defaultColWidth !== undefined && Number.isFinite(format.defaultColWidth) && format.defaultColWidth >= 0) return columnWidthPx(format.defaultColWidth, mdw)
  const base = format.baseColWidth !== undefined && Number.isFinite(format.baseColWidth) && format.baseColWidth >= 0 ? format.baseColWidth : 8
  const width = Math.trunc((base * mdw + 5) / mdw * 256) / 256
  return Math.ceil(columnWidthPx(width, mdw) / 8) * 8
}

/** Parse `<sheetFormatPr>` attributes. */
export function sheetFormat(a: Record<string, unknown>): { baseColWidth?: number; defaultColWidth?: number } {
  const num = (v: unknown) => typeof v === 'string' && v !== '' && Number.isFinite(Number(v)) ? Number(v) : undefined
  const baseColWidth = num(a.baseColWidth), defaultColWidth = num(a.defaultColWidth)
  return { ...(baseColWidth !== undefined ? { baseColWidth } : {}), ...(defaultColWidth !== undefined ? { defaultColWidth } : {}) }
}
