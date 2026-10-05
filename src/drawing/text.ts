/** Source-preserving DrawingML text model shared by document adapters. */
export type TextDirection = 'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl'
export interface DrawingTextStyle {
  bold?: boolean
  italic?: boolean
  fontSizePt?: number
  color?: string
  /** Explicit DrawingML a:noFill suppresses glyph paint while retaining logical text. */
  noFill?: boolean
  fontFamily?: string
  fontFamilyEastAsia?: string
  fontFamilyComplexScript?: string
  characterSpacingPt?: number
  language?: string
}
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
