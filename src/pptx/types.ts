import type { EmbeddedFontFace, FontDiagnostic } from '../core/fonts/types'
import type { ContentDiagnostic } from '../drawing/parts'
import type { DrawingCoverageEntry } from '../drawing/coverage'
import type { DrawingContent } from '../drawing/content'
import type { GeometryIssue } from '../drawing/geometry'
import type { SceneGroupTransform, SceneImage, SceneNode } from '../drawing/scene'
import type { DrawingFill, DrawingIssue, DrawingLine, ThemeContext } from '../drawing/style'
import type { DrawingTextBody, DrawingTextParagraph, DrawingTextRun, DrawingTextSpacing, DrawingTextStyle, DrawingTabStop } from '../drawing/text'

/** PPTX document model. All placement geometry in EMU (rendered via emuToPx). */
export interface PptxTextStyle extends DrawingTextStyle {}

export interface PptxTextRun extends DrawingTextRun {
  /** Source rPr only; absent properties must remain inheritable by table regions. */
  directProperties?: PptxTextStyle
  propertySources?: Partial<Record<keyof PptxTextStyle, 'default' | 'master' | 'layout' | 'placeholder' | 'list' | 'paragraph' | 'end' | 'run'>>
}

export interface PptxTextSpacing extends DrawingTextSpacing {}
export interface PptxTabStop extends DrawingTabStop {}

export interface PptxParagraph extends DrawingTextParagraph { runs: PptxTextRun[] }

export interface PptxTextBody extends DrawingTextBody {
  diagnostics?: PptxDiagnostic[]
  paragraphs: PptxParagraph[]
}

/** An image part referenced by a shape (p:blipFill -> a:blip r:embed). */
export interface PptxImageRef extends SceneImage {
  /** Raw encoded bytes (png/jpeg/gif/webp); required for PPTX package assets. */
  data: Uint8Array
}

export interface PptxSource {
  partPath: string
  treePath: string
  element: string
  id?: string
  name?: string
  representation: 'native' | 'choice' | 'fallback'
  reason?: string
  feature?: string
}

export type PptxDiagnostic = (GeometryIssue | DrawingIssue | ContentDiagnostic | {
  kind: 'invalid-transform' | 'missing-representation' | 'fallback-representation' | 'missing-image' | 'unsupported-object' | 'deferred-animation' | 'unsupported-text-alignment'
  message: string
  feature?: string
}) & { source?: PptxSource }

export interface PptxGroupTransform extends SceneGroupTransform {}

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
  /** Direct DrawingML paint; explicit none/transparent paint overrides style fill. */
  drawingFill?: DrawingFill
  drawingBorders?: Partial<Record<'left' | 'right' | 'top' | 'bottom', DrawingLine>>
  margins?: { leftEmu?: number; rightEmu?: number; topEmu?: number; bottomEmu?: number }
  anchor?: PptxTextBody['anchor']
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

export interface PptxShape extends SceneNode<PptxTextBody, PptxImageRef> {
  /** Inherited from a slide layout when the shape carries no explicit xfrm. */
  placeholder?: { type?: string; idx?: number }
  geometry: 'rect' | 'ellipse' | 'roundRect' | 'other'
  /** Original preset name, including names the engine does not recognize. */
  presetName?: string
  /** Plain text from an unpainted compatibility Choice, retained for later audit/search. */
  sourceTextBody?: PptxTextBody
  children?: PptxShape[]
  group?: PptxGroupTransform
  source?: PptxSource
  diagnostics?: PptxDiagnostic[]
  /** Table content remains owned and painted by the PPTX adapter. */
  content?: DrawingContent<PptxParagraph, PptxTextBody>
  table?: PptxTable
}

export interface PptxSlide {
  index: number
  widthEmu: number
  heightEmu: number
  shapes: PptxShape[]
  /** Resolved solid background from the slide, layout, or master. Absent means white. */
  background?: string
  theme?: ThemeContext
  diagnostics?: PptxDiagnostic[]
}

export interface PptxDocument {
  drawingCoverage?: DrawingCoverageEntry[]
  embeddedFonts?: EmbeddedFontFace[]
  fontDiagnostics?: FontDiagnostic[]
  slideWidthEmu: number
  slideHeightEmu: number
  slides: PptxSlide[]
  /** Unique images across all slides, in first-use order (matches shape.imageIndex). */
  images: PptxImageRef[]
}
