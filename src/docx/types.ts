/**
 * DOCX document model — a flattened, renderer-friendly view of the OOXML
 * document part. Layout (line breaking, page flow) operates on this.
 */

export interface DocxTextRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  fontFamily?: string
  fontSizePt?: number
  color?: string
  highlight?: string
  tabWidths?: number[]
  breakBefore?: boolean
}

export type ParagraphAlign = 'left' | 'center' | 'right' | 'justify'

export interface DocxParagraph {
  runs: DocxTextRun[]
  align: ParagraphAlign
  /** Indents in twips. */
  indentLeftTwips?: number
  indentRightTwips?: number
  indentFirstLineTwips?: number
  /** Spacing in twips. */
  spacingBeforeTwips?: number
  spacingAfterTwips?: number
  lineSpacing?: { rule: 'auto' | 'exact' | 'atLeast'; value: number }
  outlineLevel?: number
}

export interface DocxPageMargins {
  topTwips: number
  rightTwips: number
  bottomTwips: number
  leftTwips: number
  headerTwips: number
  footerTwips: number
  gutterTwips: number
}

export interface DocxSection {
  margins: DocxPageMargins
  pageSize: { widthTwips: number; heightTwips: number; orientation: 'portrait' | 'landscape' }
  paragraphs: DocxParagraph[]
  /** All flow content in document order (paragraphs and tables). */
  blocks: DocxBlock[]
}

export interface DocxDocument {
  sections: DocxSection[]
  defaultFontFamily: string
  defaultFontSizePt: number
  styleDefaults: Map<string, { fontFamily?: string; fontSizePt?: number }>
}

export type TableCellBorder = { style?: string; color?: string }

export interface DocxTableBorders {
  top?: TableCellBorder
  bottom?: TableCellBorder
  left?: TableCellBorder
  right?: TableCellBorder
  insideH?: TableCellBorder
  insideV?: TableCellBorder
}

export interface DocxTableCellMargins {
  topTwips: number
  rightTwips: number
  bottomTwips: number
  leftTwips: number
}

export interface DocxTableCell {
  paragraphs: DocxParagraph[]
  /** Number of grid columns this cell spans (default 1). */
  gridSpan: number
  /** Vertical merge state: 'restart' starts a merged region, 'continue' extends it. */
  vMerge?: 'restart' | 'continue'
  /** Shading fill color (hex RGB). */
  fill?: string
  /** Borders override for this cell. */
  borders?: DocxTableBorders
}

export interface DocxTableRow {
  cells: DocxTableCell[]
  heightTwips?: number
  heightRule?: 'atLeast' | 'exact' | 'auto'
}

export interface DocxTable {
  gridColsTwips: number[]
  rows: DocxTableRow[]
  fill?: string
  borders?: DocxTableBorders
  cellMargins: DocxTableCellMargins
}

/** A flow content block: paragraph or table, in document order. */
export type DocxBlock =
  | { kind: 'p'; paragraph: DocxParagraph }
  | { kind: 'table'; table: DocxTable }
