import type { XlsxFrozenPane } from './frozenPane'
/** XLSX document model. */
import type { SceneImage, SceneNode } from '../drawing/scene'
import type { DrawingContent } from '../drawing/content'
import type { ContentDiagnostic } from '../drawing/parts'
import type { DrawingCoverageEntry } from '../drawing/coverage'
import type { PptxParagraph, PptxTextBody } from '../pptx/types'
import type { ThemeContext } from '../drawing/style'
import type { EmbeddedFontFace, FontDiagnostic } from '../core/fonts/types'
import type {
  CalcSettings,
  DefinedNameMetadata,
  ResolvedSemantics,
  TableMetadata,
  WorkbookSheetIdentity,
} from './formula/types'
import type { DeclaredSheetFilters } from './view/declaredFilters'

export interface XlsxDrawingSource { partPath: string; treePath: string; element: string; id?: string; name?: string; referenceId?: string; representation: 'native' | 'choice' | 'fallback'; reason?: string; emptySelection?: boolean; /** An authored fallback's original feature (spec 6.6). */ feature?: 'chart' | 'ChartEx' }
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
  /** Whether a cached or calculated value is available. Undefined retains legacy null-based behavior. */
  hasCachedValue?: boolean
  /** Distinguishes real cached/calculated errors from identical-looking text; false is explicit text. */
  valueIsError?: boolean
  formula?: string
  sharedFormula?: { si: number; ref?: string }
  /** Legacy fixed array formula: OOXML `<f t="array" ref="A1:B2">` output range. */
  arrayRef?: string
  /** Dynamic array verified through cm -> cellMetadata -> XLDAPR future record.
   * arrayRef then records the saved cache extent, not a fixed output size. */
  dynamicArray?: boolean
  /** Raw OOXML `<f t="...">` formula type (e.g. `dataTable`); array/shared are
   * also represented by arrayRef/sharedFormula. Never inferred from content. */
  formulaType?: string
  /** Calculate always flag from <f ca="1"> per ECMA-376 Part 1 §18.3.1.40. */
  ca?: boolean
  /** Resolved style (from cellXfs + fonts/fills/borders). */
  style?: XlsxCellStyle
}

export interface XlsxCellStyle {
  numFmtId: number
  /** Authored custom numFmt formatCode (from styles.xml numFmts) when present. */
  formatCode?: string
  bold?: boolean
  italic?: boolean
  fontSizePt?: number
  color?: string
  fillColor?: string
  borders?: { left?: string; right?: string; top?: string; bottom?: string }
  /** SpreadsheetML ARGB per border side (absent: black). */
  borderColors?: { left?: string; right?: string; top?: string; bottom?: string }
  /** Normalized alignment textRotation: 1-180 as authored, or 255 stacked. Absent, 0 and invalid values stay undefined (horizontal). */
  textRotation?: number
  /** Alignment horizontal/vertical as authored (e.g. left/center/right, top/center/bottom). Absent stays undefined. */
  horizontal?: string
  vertical?: string
  /** wrapText as authored. Absent stays undefined. */
  wrapText?: boolean
  /** alignment@shrinkToFit as authored (ignored when wrapping). Absent stays undefined. */
  shrinkToFit?: boolean
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
/** Print page setup as authored (all optional; Excel defaults apply downstream). */
export interface XlsxPageSetup {
  paperSizeId?: number
  orientation?: 'portrait' | 'landscape'
  /** Explicit print scale percent (10-400). Absent means 100. */
  scale?: number
  /** Fit-to-page targets (0/undefined = unbounded in that axis). */
  fitToWidth?: number
  fitToHeight?: number
  /** sheetPr/pageSetUpPr fitToPage flag. */
  fitToPage?: boolean
}
/** Page margins in inches (Excel defaults when the element is absent). */
export interface XlsxPageMargins {
  left: number
  right: number
  top: number
  bottom: number
  header: number
  footer: number
}

/** A worksheet drawing marker: zero-based cell plus EMU offsets inside it. */
export interface XlsxAnchorMarker { col: number; row: number; colOffEmu: number; rowOffEmu: number }
/** Raw anchor placement as authored; occupied bounds are resolved later against sheet geometry. */
export type XlsxDrawingAnchorRecord =
  | { id: string; kind: 'oneCell'; from: XlsxAnchorMarker; cx: number; cy: number }
  | { id: string; kind: 'twoCell'; from: XlsxAnchorMarker; to: XlsxAnchorMarker }
  | { id: string; kind: 'absolute'; x: number; y: number; cx: number; cy: number }
/** `unresolved` counts anchors or drawing parts whose placement could not be read (fail-closed). */
export interface XlsxDrawingAnchors { records: XlsxDrawingAnchorRecord[]; unresolved: number }

export interface XlsxSheet {
  /** Parsed x14 sparkline groups (resident worksheet path). */
  sparklineGroups?: import('./sparklines/model').XlsxSparklineGroup[]
  /** Render-ready sparkline cache, filled once after parse (resident path): row -> col -> resolved sparkline. */
  resolvedSparklines?: import('./sparklines/resolve').ResolvedSparklines
  drawingCoverage?: DrawingCoverageEntry[]
  /** Authored frozen panes of the first sheet view (interactive viewer only). */
  frozen?: XlsxFrozenPane
  name: string
  sourcePartPath?: string
  /** Resolved workbook semantics (date system/locale/zone/clock) so the
   * renderer can decode serials with the SAME model as evaluation without
   * callers injecting it manually. Optional for hand-built sheets. */
  semantics?: ResolvedSemantics
  /** Stable workbook identity (original order, never a filtered index). */
  sheetId?: string
  /** Tables owned by this sheet (also present in the document-wide inventory). */
  tables?: TableMetadata[]
  /** Bounded pivot/control region references used by region-style rendering. */
  pivotRegions?: import('./pivots/stub').PivotRegionRef[]
  slicerRefs?: import('./pivots/stub').ControlPartRef[]
  timelineRefs?: import('./pivots/stub').ControlPartRef[]
  workbookIndex?: number
  kind?: 'worksheet' | 'other'
  /** Excel column pixel model (workbook MDW, sheet default width); absent:
   * Calibri 11 (MDW 7px, 64px default). */
  columnMetrics?: import('./columnWidth').ColumnMetrics
  rows: XlsxRow[]
  cols: XlsxColumnSpec[]
  merges: string[]
  /** Parsed merge ranges (same info as merges, structured). */
  mergeRanges: XlsxMergeRange[]
  drawings?: XlsxDrawing[]
  drawingMarkers?: { maxCol: number; maxRow: number }
  /** Bounded DrawingML anchor metadata for the viewer's structural guards
   * (S3.3): one record per anchor, independent of content/image parsing.
   * Undefined means never extracted (hand-built sheets), never "no anchors". */
  drawingAnchors?: XlsxDrawingAnchors
  drawingDiagnostics?: ContentDiagnostic[]
  /** Render-time unsupported-format diagnostics (same shape as doc.diagnostics),
   * deduped across repeated paints. No product warning UI. */
  diagnostics?: Array<{ kind: string; feature: string; message: string }>
  drawingTheme?: ThemeContext
  pageSetup?: XlsxPageSetup
  pageMargins?: XlsxPageMargins
}

export interface XlsxDocument {
  drawingCoverage?: DrawingCoverageEntry[]
  sheets: XlsxSheet[]
  images?: XlsxImage[]
  embeddedFonts?: EmbeddedFontFace[]
  fontDiagnostics?: FontDiagnostic[]
  /** Formula-evaluation diagnostics (unsupported constructs reached while a
   * recalculation ran; cached values retained). Deduped by feature+message. */
  diagnostics?: Array<{ kind: string; feature: string; message: string }>
  /** B0 workbook metadata (original order/identities, names, tables, settings). */
  workbookSheets?: WorkbookSheetIdentity[]
  definedNames?: DefinedNameMetadata[]
  tables?: TableMetadata[]
  /** Declared file filter/sort state as diagnostics only (S8.3, spec §7.5):
   * present only when a sheet declares worksheet/table autoFilter or
   * sortState. Supported file filters persist as hidden rows; nothing here
   * becomes an editable viewer rule. */
  declaredFilters?: DeclaredSheetFilters[]
  calc?: CalcSettings
  /** Value policy (spec §4.4), set only on saved-values documents (parseXlsx valuePolicy 'cached');
   * absent means 'engine'. Read through workbookValuePolicy / documentValuePolicy. */
  valuePolicy?: 'engine' | 'cached'
  /** True when the package actually ships an `xl/calcChain.xml` part, detected
   * from the workbook relationship inventory — never inferred from a producer
   * or guessed. `undefined` means the model did not declare it. */
  calcChainPresent?: boolean
  semantics?: ResolvedSemantics
  /** Resolved cellXfs table (index = cell styleIndex), published to worker
   * viewer sources so remote cells style exactly like resident ones. */
  cellStyles?: readonly XlsxCellStyle[]
}
