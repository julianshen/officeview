import { isDateFormat } from '../charts/dates'

/**
 * Whether a cell style displays its number as a date or time (data-bar participation, Ruling S3-G): an authored format
 * code with a date or time token, or a built-in date/time id (14-22, 27-36 and 50-58 locale dates, 45-47 times). A
 * missing style is the General format. Shared by the resident and paged CF sources, so both classify identically.
 */
export function isDateNumberStyle(style: { readonly numFmtId: number; readonly formatCode?: string } | undefined): boolean {
  if (!style) return false
  if (style.formatCode !== undefined) return isDateFormat(style.formatCode)
  const id = style.numFmtId
  return (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58)
}
