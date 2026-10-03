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
  /** Vertical space reserved by wp:effectExtent, outside the image. */
  effectExtentEmu?: { top: number; bottom: number }
  /** Present for wp:anchor drawings; absent means inline (in the text flow). */
  floating?: DocxFloating
  /** Vector content stored in the package rather than an encoded picture. */
  drawing?: DocxDrawing
}

export interface DocxDrawingShape {
  xEmu: number
  yEmu: number
  widthEmu: number
  heightEmu: number
  geometry: string
  rotationDeg?: number
  fill?: string
  line?: { color: string; widthEmu: number }
  paragraphs: Array<{ runs: DocxTextRun[]; align: ParagraphAlign }>
  fontFamily: string
  textColor?: string
}

export type DocxDrawing =
  | { kind: 'diagram'; shapes: DocxDrawingShape[] }
  | {
      kind: 'ink'
      strokes: Array<{ points: Array<[number, number]>; widthEmu: number; color: string }>
      widthEmu: number
      heightEmu: number
    }
  | {
      kind: 'chart'
      title?: string
      categories: string[]
      series: Array<{ name: string; values: Array<number | undefined>; color: string }>
      min?: number
      max?: number
      majorUnit?: number
      gapWidth: number
      overlap: number
      legend: boolean
      fontFamily: string
      fontSizePt: number
    }
  | {
      kind: 'textbox'
      paragraphs: DocxParagraph[]
      vertical: boolean
      fontFamily: string
      fontSizePt: number
      insets: { left: number; top: number; right: number; bottom: number }
      fill?: string
      line?: { color: string; widthEmu: number }
    }

export type DocxInline = { kind: 'text'; run: DocxTextRun } | { kind: 'image'; image: DocxImage }

export interface DocxParagraph {
  runs: DocxTextRun[]
  images: DocxImage[]
  /** Text and drawings in source order, when the paragraph contains drawings. */
  inline?: DocxInline[]
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

export type TableCellBorder = { style?: string; color?: string; widthPt?: number }

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
  /** No grid or cell widths: size columns from their content during layout. */
  autoWidth?: boolean
  rows: DocxTableRow[]
  fill?: string
  borders?: DocxTableBorders
  cellMargins: DocxTableCellMargins
}

/** A flow content block: paragraph or table, in document order. */
export type DocxBlock = { kind: 'p'; paragraph: DocxParagraph } | { kind: 'table'; table: DocxTable }
