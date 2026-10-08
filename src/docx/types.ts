/**
 * DOCX document model — a flattened, renderer-friendly view of the OOXML
 * document part. Layout (line breaking, page flow) operates on this.
 */
import type { DrawingContent, DrawingContentShape } from '../drawing/content'
import type { ParsedDrawingTextBody } from '../drawing/text-parse'
import type { ImageSelection, SvgCandidate, SvgVerdict } from '../core/svg'
import type { DrawingCoverageEntry } from '../drawing/coverage'
import type { TextWarp } from '../drawing/text'

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
  /** Explicit DrawingML / w14:noFill suppresses glyph paint while retaining logical text. */
  noFill?: boolean
  /** WordArt fill: gradient */
  textFill?: { kind: 'gradient'; stops: Array<{ position: number; color: string }>; angle: number }
  /** WordArt outline: stroked centered on glyph edge */
  textOutline?: { color: string; widthPx: number }
  /** WordArt outer shadow */
  textShadow?: { color: string; blurPx: number; offsetX: number; offsetY: number }
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
  /** Selected vector payload relationship, when this image carries one. */
  referenceId?: string
  /** svgBlip candidate preferred over the raster `data` when preflight allows. */
  svg?: SvgCandidate
  /** Preflight verdict when the primary `data` bytes are themselves SVG. */
  primarySvgVerdict?: SvgVerdict
  /** Explicitly false when an SVG candidate exists but no raster fallback does. */
  hasRaster?: boolean
  /** Owner-part media path hint for reliable SVG detection. */
  pathHint?: string
  /** Parse/decode selection record shared with the drawing coverage entry. */
  imageSelection?: ImageSelection
}

/** Compatibility aliases: Word owns its paragraph/text layout. */
export type DocxDrawingShape = DrawingContentShape<ParsedDrawingTextBody>
export type DocxDrawing = DrawingContent<DocxParagraph, ParsedDrawingTextBody>

export type DocxInline = { kind: 'text'; run: DocxTextRun } | { kind: 'image'; image: DocxImage }

export interface DocxParagraph {
  /** Formatting of the paragraph mark, used for an empty line. */
  paragraphMark?: Partial<DocxTextRun>
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
  /** Preset WordArt warp configuration for text routed through this paragraph */
  textWarp?: TextWarp
  /** Unsupported appearance or warp diagnostics associated with this paragraph */
  diagnostics?: Array<{ kind: 'unsupported-text-alignment' | 'unsupported-text-appearance' | 'unsupported-text-warp'; feature: string; message: string }>
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
  /** This section's start relative to the previous section (default nextPage). */
  type?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn'
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
  /** Ordered repeated content; provided lists (even empty) override paragraph arrays. */
  headerBlocks?: DocxBlock[]
  footerBlocks?: DocxBlock[]
  firstHeaderBlocks?: DocxBlock[]
  firstFooterBlocks?: DocxBlock[]
  titlePg?: boolean
  firstHeader?: DocxParagraph[]
  firstFooter?: DocxParagraph[]
}

export interface DocxDocument {
  drawingCoverage?: DrawingCoverageEntry[]
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
  /** Resolved cell-specific padding, inherited field by field. */
  margins?: DocxTableCellMargins
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
  /** w:textDirection — vertical cell text flow. Absent/invalid means lrTb (ordinary horizontal). */
  textDirection?: 'lrTb' | 'tbRl' | 'btLr' | 'lrTbV' | 'tbRlV' | 'tbLrV'
}

export interface DocxTableRow {
  /** Grid columns omitted before/after a ragged row. */
  gridBefore?: number
  gridAfter?: number
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
