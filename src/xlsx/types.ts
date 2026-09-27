/** XLSX document model. */
export interface XlsxCell {
  ref: string
  row: number // 0-based
  col: number // 0-based
  /** Raw value; strings resolved from sharedStrings already. */
  value: string | number | boolean | null
  /** Style index into cellXfs. */
  styleIndex: number
  formula?: string
  /** Resolved style (from cellXfs + fonts/fills/borders). */
  style?: XlsxCellStyle
}

export interface XlsxCellStyle {
  numFmtId: number
  bold?: boolean
  italic?: boolean
  fontSizePt?: number
  color?: string
  fillColor?: string
  borders?: { left?: string; right?: string; top?: string; bottom?: string }
}

export interface XlsxRow {
  index: number
  heightPt?: number
  customHeight?: boolean
  cells: XlsxCell[]
}

export interface XlsxColumnSpec {
  min: number
  max: number
  widthChars?: number
  hidden?: boolean
}

export interface XlsxMergeRange {
  minRow: number
  minCol: number
  maxRow: number
  maxCol: number
}

export interface XlsxSheet {
  name: string
  rows: XlsxRow[]
  cols: XlsxColumnSpec[]
  merges: string[]
  /** Parsed merge ranges (same info as merges, structured). */
  mergeRanges: XlsxMergeRange[]
}

export interface XlsxDocument {
  sheets: XlsxSheet[]
}
