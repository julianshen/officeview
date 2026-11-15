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

/** `<sheetFormatPr>`: the column defaults, and the default row height with its flags (Task 28 round 5). */
export type SheetFormat = { baseColWidth?: number; defaultColWidth?: number; defaultRowHeight?: number; customHeight?: boolean; zeroHeight?: boolean; thickTop?: boolean; thickBottom?: boolean }
/** Parse `<sheetFormatPr>` attributes. */
export function sheetFormat(a: Record<string, unknown>): SheetFormat {
  const num = (v: unknown) => typeof v === 'string' && v !== '' && Number.isFinite(Number(v)) ? Number(v) : undefined
  const flag = (v: unknown) => (v === '1' || v === 'true' ? true : v === '0' || v === 'false' ? false : undefined)
  const out: SheetFormat = {}
  const baseColWidth = num(a.baseColWidth), defaultColWidth = num(a.defaultColWidth), defaultRowHeight = num(a.defaultRowHeight)
  if (baseColWidth !== undefined) out.baseColWidth = baseColWidth
  if (defaultColWidth !== undefined) out.defaultColWidth = defaultColWidth
  if (defaultRowHeight !== undefined) out.defaultRowHeight = defaultRowHeight
  for (const k of ['customHeight', 'zeroHeight', 'thickTop', 'thickBottom'] as const) {
    const f = flag(a[k])
    if (f !== undefined) out[k] = f
  }
  return out
}
/**
 * The authored height (pt) of a row without `ht` (Ruling S4-DEFAULT-ROW): `defaultRowHeight`, unless `zeroHeight`
 * hides such rows or `thickTop`/`thickBottom` change their height; `customHeight` only marks the value as set by hand.
 * Undefined: not known (absent or invalid).
 */
export function defaultRowPtOf(format: SheetFormat | undefined): number | undefined {
  const h = format?.defaultRowHeight
  if (h === undefined || !Number.isFinite(h) || h < 0 || format?.zeroHeight || format?.thickTop || format?.thickBottom) return undefined
  return h
}
