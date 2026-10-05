/** XLSX document model. */
import type { SceneImage, SceneNode } from '../drawing/scene'
import type { DrawingContent } from '../drawing/content'
import type { ContentDiagnostic } from '../drawing/parts'
import type { DrawingCoverageEntry } from '../drawing/coverage'
import type { PptxParagraph, PptxTextBody } from '../pptx/types'
import type { ThemeContext } from '../drawing/style'
import type { EmbeddedFontFace, FontDiagnostic } from '../core/fonts/types'

export interface XlsxDrawingSource { partPath: string; treePath: string; element: string; id?: string; name?: string; referenceId?: string; representation: 'native' | 'choice' | 'fallback'; reason?: string; emptySelection?: boolean }
export interface XlsxDrawing extends SceneNode<PptxTextBody, SceneImage> {
  source: XlsxDrawingSource
  children?: XlsxDrawing[]
  content?: DrawingContent<PptxParagraph, PptxTextBody>
}
export interface XlsxImage extends SceneImage { data: Uint8Array }
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
  /** Normalized alignment textRotation: 1-180 as authored, or 255 stacked. Absent, 0 and invalid values stay undefined (horizontal). */
  textRotation?: number
  /** Alignment horizontal/vertical as authored (e.g. left/center/right, top/center/bottom). Absent stays undefined. */
  horizontal?: string
  vertical?: string
  /** wrapText as authored. Absent stays undefined. */
  wrapText?: boolean
}

export interface XlsxRow {
  index: number
  heightPt?: number
  customHeight?: boolean
  hidden?: boolean
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
  drawingCoverage?: DrawingCoverageEntry[]
  name: string
  sourcePartPath?: string
  rows: XlsxRow[]
  cols: XlsxColumnSpec[]
  merges: string[]
  /** Parsed merge ranges (same info as merges, structured). */
  mergeRanges: XlsxMergeRange[]
  drawings?: XlsxDrawing[]
  drawingMarkers?: { maxCol: number; maxRow: number }
  drawingDiagnostics?: ContentDiagnostic[]
  drawingTheme?: ThemeContext
}

export interface XlsxDocument {
  drawingCoverage?: DrawingCoverageEntry[]
  sheets: XlsxSheet[]
  images?: XlsxImage[]
  embeddedFonts?: EmbeddedFontFace[]
  fontDiagnostics?: FontDiagnostic[]
}
