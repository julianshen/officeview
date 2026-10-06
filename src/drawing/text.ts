/** Source-preserving DrawingML text model shared by document adapters. */
export type TextDirection = 'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl'
export interface DrawingTextStyle {
  bold?: boolean
  italic?: boolean
  fontSizePt?: number
  color?: string
  /** Explicit DrawingML a:noFill suppresses glyph paint while retaining logical text. */
  noFill?: boolean
  /**
   * WordArt fill. Absent means the flat `color` above. Gradient/pattern never
   * affect advances or recording — paint-only.
   */
  textFill?: { kind: 'gradient'; stops: Array<{ position: number; color: string }>; angle: number }
    | { kind: 'pattern'; preset: PatternPreset; fg: string; bg: string }
  /** WordArt outline: stroked centered on the glyph edge after the fill. */
  textOutline?: { color: string; widthPx: number }
  /**
   * WordArt outer shadow. Offsets follow canvas CTM semantics in rotated
   * frames; Word-parity of shadow direction there is a validation item.
   */
  textShadow?: { color: string; blurPx: number; offsetX: number; offsetY: number }
  fontFamily?: string
  fontFamilyEastAsia?: string
  fontFamilyComplexScript?: string
  characterSpacingPt?: number
  language?: string
}
/**
 * WordArt pattern presets paint can tile (diagonal families and grids).
 * Parse accepts exactly this set and defers the rest with a diagnostic;
 * keep both sides on this list.
 */
export const SUPPORTED_PATTERN_PRESETS: ReadonlySet<string> = new Set([
  'dkUpDiag', 'dkDnDiag', 'ltUpDiag', 'ltDnDiag', 'smGrid', 'lgGrid',
])
/** Type-enforced twin of the set above: adding a preset here without
 * extending paintPatternTile is a compile error, not a silent grid. */
export type PatternPreset = 'dkUpDiag' | 'dkDnDiag' | 'ltUpDiag' | 'ltDnDiag' | 'smGrid' | 'lgGrid'
export interface DrawingTextRun extends DrawingTextStyle {
  text: string
  directProperties?: DrawingTextStyle
  propertySources?: Partial<Record<keyof DrawingTextStyle, 'default' | 'master' | 'layout' | 'placeholder' | 'list' | 'paragraph' | 'end' | 'run'>>
}
/** Original source-run provenance for a placed cluster, even across style boundaries. */
export interface SourceRunRef { runIndex: number; start: number; end: number; style: DrawingTextRun }
export interface DrawingTextSpacing { kind: 'points' | 'percent'; value: number }
/** Physical flow objects have no source characters or source-run identity. */
export interface DrawingInlineSlot { id: number; sourceOffset: number; width: number; height: number }
export interface DrawingTabStop { positionEmu: number; align: 'left' | 'center' | 'right' | 'decimal' }
export interface DrawingTextParagraph {
  runs: DrawingTextRun[]
  inlineSlots?: DrawingInlineSlot[]
  /** Stacked text uses one em cell per grapheme, including authored spaces. */
  graphemeCells?: boolean
  /** Excel wraps authored whitespace as physical cells, retaining every space. */
  wrapWhitespace?: boolean
  align: 'left' | 'center' | 'right' | 'justify'
  sourceAlign?: string
  bullet?: boolean
  level: number
  bulletCharacter?: string
  marginLeftEmu?: number
  marginRightEmu?: number
  indentEmu?: number
  defaultTabSizeEmu?: number
  tabStops?: DrawingTabStop[]
  lineSpacing?: DrawingTextSpacing
  /** Word's 240ths/twips line rules depend on measured normal font metrics. */
  wordLineSpacing?: { rule: 'auto' | 'exact' | 'atLeast'; value: number }
  spaceBefore?: DrawingTextSpacing
  spaceAfter?: DrawingTextSpacing
  defaultProperties?: DrawingTextStyle
  endProperties?: DrawingTextStyle
}
export interface DrawingTextBody {
  paragraphs: DrawingTextParagraph[]
  direction?: TextDirection
  anchor: 't' | 'ctr' | 'b'
  insetLeftEmu: number
  insetRightEmu: number
  insetTopEmu: number
  insetBottomEmu: number
  wrap: boolean
}
export interface LocalAffine { a: number; b: number; c: number; d: number; e: number; f: number }
