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
  textWarp?: TextWarp
}
export interface LocalAffine { a: number; b: number; c: number; d: number; e: number; f: number }

export type TextWarpPreset =
  | 'textNoShape'
  | 'textPlain'
  | 'textStop'
  | 'textTriangle'
  | 'textTriangleInverted'
  | 'textChevron'
  | 'textChevronInverted'
  | 'textRingInside'
  | 'textRingOutside'
  | 'textArchUp'
  | 'textArchDown'
  | 'textCircle'
  | 'textButton'
  | 'textArchUpPour'
  | 'textArchDownPour'
  | 'textCirclePour'
  | 'textButtonPour'
  | 'textCurveUp'
  | 'textCurveDown'
  | 'textCanUp'
  | 'textCanDown'
  | 'textWave1'
  | 'textWave2'
  | 'textDoubleWave1'
  | 'textWave4'
  | 'textInflate'
  | 'textDeflate'
  | 'textInflateBottom'
  | 'textDeflateBottom'
  | 'textInflateTop'
  | 'textDeflateTop'
  | 'textDeflateInflate'
  | 'textDeflateInflateDeflate'
  | 'textFadeRight'
  | 'textFadeLeft'
  | 'textFadeUp'
  | 'textFadeDown'
  | 'textSlantUp'
  | 'textSlantDown'
  | 'textCascadeUp'
  | 'textCascadeDown'

/**
 * WordArt preset text warps with complete geometry implementations.
 * Presets outside this set emit an 'unsupported-text-warp' diagnostic
 * and fall back gracefully to unwarped text rendering.
 */
export const SUPPORTED_TEXT_WARP_PRESETS: ReadonlySet<string> = new Set<TextWarpPreset>([
  'textArchUp',
  'textArchDown',
  'textCircle',
  'textWave1',
  'textWave2',
  'textInflate',
  'textDeflate',
  'textSlantUp',
  'textSlantDown',
  'textCurveUp',
  'textCurveDown',
])

export interface TextWarp {
  preset: TextWarpPreset
  adjustments?: Record<string, number>
}

/** ECMA-376 Part 1 §20.1.9.22 default adjust values for preset text warps */
export const DEFAULT_WARP_ADJUSTMENTS: Readonly<Record<string, Record<string, number>>> = {
  textArchUp: { adj: 10800000 },
  textArchDown: { adj: 10800000 },
  textCircle: { adj: 10800000 },
  textButton: { adj: 10800000 },
  textArchUpPour: { adj: 10800000 },
  textArchDownPour: { adj: 10800000 },
  textCirclePour: { adj: 10800000 },
  textButtonPour: { adj: 10800000 },
  textRingInside: { adj: 10800000 },
  textRingOutside: { adj: 10800000 },
  textCurveUp: { adj: 25000 },
  textCurveDown: { adj: 25000 },
  textCanUp: { adj: 25000 },
  textCanDown: { adj: 25000 },
  textWave1: { adj1: 0, adj2: 50000 },
  textWave2: { adj1: 0, adj2: 50000 },
  textDoubleWave1: { adj1: 0, adj2: 50000 },
  textWave4: { adj1: 0, adj2: 50000 },
  textInflate: { adj: 50000 },
  textDeflate: { adj: 50000 },
  textInflateBottom: { adj: 50000 },
  textDeflateBottom: { adj: 50000 },
  textInflateTop: { adj: 50000 },
  textDeflateTop: { adj: 50000 },
  textDeflateInflate: { adj: 50000 },
  textDeflateInflateDeflate: { adj: 50000 },
  textSlantUp: { adj: 25000 },
  textSlantDown: { adj: 25000 },
  textFadeRight: { adj: 50000 },
  textFadeLeft: { adj: 50000 },
  textFadeUp: { adj: 50000 },
  textFadeDown: { adj: 50000 },
  textCascadeUp: { adj: 25000 },
  textCascadeDown: { adj: 25000 },
  textStop: { adj: 0 },
  textTriangle: { adj: 50000 },
  textTriangleInverted: { adj: 50000 },
  textChevron: { adj: 50000 },
  textChevronInverted: { adj: 50000 },
}

