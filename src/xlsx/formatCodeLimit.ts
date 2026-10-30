/** Excel's own limit on a custom number-format code (UTF-16 units). Excel
 * refuses longer codes, so an authored one is shown as General (documented in
 * docs/formula-compatibility.md); every formatter therefore parses at most
 * this much per authored code: a cell xf's (parse.ts) or a table-style or
 * pivot-format dxf's (tableStyles/parseStylesExtras.ts). */
export const MAX_FORMAT_CODE_UNITS = 255

/** An authored format code, or null when it is over the limit and shows as General. */
export function admittedFormatCode(code: string): string | null {
  return code.length > MAX_FORMAT_CODE_UNITS ? null : code
}
