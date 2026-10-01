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
  /** Field instruction (e.g. 'PAGE', 'NUMPAGES') resolved at paint time. */
  field?: string
}

export type ParagraphAlign = 'left' | 'center' | 'right' | 'justify'

/** Positioning of a floating (wp:anchor) drawing. */
export interface DocxFloating {
  /** w:behindDoc — drawn beneath the text layer instead of over it. */
  behindDoc: boolean
  /** w:relativeHeight — z-order among floating images on a page. */
  relativeHeight: number
  /** Text wrapping mode declared on the anchor. */
  wrap: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom'
  posH: { relativeFrom: string; offsetEmu: number; align?: string }
  posV: { relativeFrom: string; offsetEmu: number; align?: string }
}

export interface DocxImage {
  /** Raw encoded bytes (png/jpeg/gif/webp). */
  data: Uint8Array
  mime?: string
  /** Display size in EMU from wp:extent. */
  widthEmu: number
  heightEmu: number
  /** Present for wp:anchor drawings; absent means inline (in the text flow). */
  floating?: DocxFloating
}

export interface DocxParagraph {
  runs: DocxTextRun[]
  images: DocxImage[]
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
  /** Resolved list marker (e.g. "1.", "a)", "•") from w:numPr. */
  listMarker?: string
  /** List level (0-based) from w:numPr/w:ilvl. */
  listLevel?: number
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
  /** Header paragraphs (w:headerReference), painted on every page. */
  header?: DocxParagraph[]
  /** Footer paragraphs (w:footerReference), painted on every page. */
  footer?: DocxParagraph[]
  /** w:titlePg — the first page uses these instead. */
  titlePg?: boolean
  firstHeader?: DocxParagraph[]
  firstFooter?: DocxParagraph[]
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
  /** w:tcW (type=dxa) — used to derive column widths when tblGrid is absent. */
  widthTwips?: number
  /** w:vAlign — vertical text placement within the cell. */
  vAlign?: 'top' | 'center' | 'bottom'
}

export interface DocxTableRow {
  cells: DocxTableCell[]
  heightTwips?: number
  heightRule?: 'atLeast' | 'exact' | 'auto'
  /** w:trPr/w:tblHeader — repeat this row at the top of each continuation page. */
  isHeader?: boolean
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
