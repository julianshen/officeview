/** PPTX document model. All geometry in EMU (rendered via emuToPx). */
export interface PptxTextRun {
  text: string
  bold?: boolean
  italic?: boolean
  fontSizePt?: number
  color?: string
  fontFamily?: string
}

export interface PptxParagraph {
  runs: PptxTextRun[]
  align: 'left' | 'center' | 'right' | 'justify'
  bullet?: boolean
  level: number
}

export interface PptxTextBody {
  paragraphs: PptxParagraph[]
  /** Anchor text vertically within shape: t/ctr/b. */
  anchor: 't' | 'ctr' | 'b'
  /** Inset in EMU (defaults 91440 l/r, 45720 t/b). */
  insetLeftEmu: number
  insetRightEmu: number
  insetTopEmu: number
  insetBottomEmu: number
  wrap: boolean
}

/** An image part referenced by a shape (p:blipFill -> a:blip r:embed). */
export interface PptxImageRef {
  /** Raw encoded bytes (png/jpeg/gif/webp). */
  data: Uint8Array
  mime?: string
  /** a:srcRect crop, as 0..1 fractions. */
  srcRect?: { l: number; t: number; r: number; b: number }
}

/** A table cell inside a:p:graphicFrame -> a:tbl. */
export interface PptxTableCell {
  paragraphs: PptxParagraph[]
  /** Columns spanned (a:tc@gridSpan, default 1). */
  gridSpan: number
  /** Rows spanned (a:tc@rowSpan, default 1). */
  rowSpan: number
  /** Cells absorbed by a merge (a:tc@hMerge/@vMerge) carry no content. */
  merged?: boolean
  /** Solid fill color from a:tcPr/a:solidFill. */
  fill?: string
}

export interface PptxTableRow {
  cells: PptxTableCell[]
  /** a:tr@h in EMU (optional; rows auto-size otherwise). */
  heightEmu?: number
}

export interface PptxTableBorder {
  color: string
  widthEmu?: number
}

export type PptxTableBorders = Partial<Record<'left' | 'right' | 'top' | 'bottom' | 'insideH' | 'insideV', PptxTableBorder>>

export interface PptxTable {
  /** a:tblGrid/a:gridCol@w in EMU. */
  colWidthsEmu: number[]
  rows: PptxTableRow[]
  /** a:tblPr/a:tableStyleId, resolved against ppt/tableStyles.xml. */
  styleId?: string
  /** a:tblPr@firstRow — apply the style's firstRow banding. */
  firstRow?: boolean
  /** a:tblPr@bandRow — apply the style's horizontal row banding. */
  bandRow?: boolean
  /** Resolved per-region fills from the table style. */
  styleFills?: {
    firstRow?: string
    band1?: string
    band2?: string
    wholeTable?: string
    lastRow?: string
    firstCol?: string
  }
  /** Text color the style applies to the first row (headers are often white). */
  firstRowTextColor?: string
  firstRowBold?: boolean
  styleBorders?: PptxTableBorders
  firstRowBorders?: PptxTableBorders
}

export interface PptxShape {
  /** Geometry in EMU, relative to slide origin. */
  xEmu: number
  yEmu: number
  widthEmu: number
  heightEmu: number
  /** Inherited from a slide layout when the shape carries no explicit xfrm. */
  placeholder?: { type?: string; idx?: number }
  geometry: 'rect' | 'ellipse' | 'roundRect' | 'other'
  fill?: string
  line?: { color: string; widthEmu?: number }
  textBody?: PptxTextBody
  rotationDeg?: number
  /** Picture content, when this shape is a p:pic. */
  image?: PptxImageRef
  /** Table content, when this shape is a p:graphicFrame wrapping an a:tbl. */
  table?: PptxTable
  /** Index into the document-wide image list (PptxDocument.images). */
  imageIndex?: number
}

export interface PptxSlide {
  index: number
  widthEmu: number
  heightEmu: number
  shapes: PptxShape[]
  /** Slide background fill as CSS hex (p:bg solid fill, else layout/master fallback). Absent means white. */
  background?: string
}

export interface PptxDocument {
  slideWidthEmu: number
  slideHeightEmu: number
  slides: PptxSlide[]
  /** Unique images across all slides, in first-use order (matches shape.imageIndex). */
  images: PptxImageRef[]
}
