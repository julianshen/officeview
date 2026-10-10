import { attrs, getChildren, type XmlNode } from '../core/xml'
import { TEXT_WARP_CATALOG } from './text-warp-catalog'
import { textWarpSeedValues, evaluateGuides } from './geometry'
import { SUPPORTED_PATTERN_PRESETS, type PatternPreset } from './pattern'
export { SUPPORTED_PATTERN_PRESETS, type PatternPreset }
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
export interface DrawingTextAutofitNormal {
  kind: 'normal'
  /** Font scale factor in range [0.01, 1.0]. E.g. 0.8 for 80% / 80000. */
  fontScale?: number
  /** Line spacing reduction factor in range [0.0, 1.0]. E.g. 0.2 for 20% / 20000. */
  lnSpcReduction?: number
}

export type DrawingTextAutofit =
  /** a:noAutofit: No text scaling; text underflows or overflows fixed box. */
  | { kind: 'none' }
  /**
   * a:spAutoFit: In authoring applications, the shape boundary grows to fit text.
   * In OfficeView's fixed-bounds canvas model, text renders unscaled (scale 1.0)
   * within the authored geometry without artificial shrinking.
   */
  | { kind: 'shape' }
  /** a:normAutofit: Normal text scaling; fonts shrink to fit the text frame. */
  | DrawingTextAutofitNormal

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
  autofit?: DrawingTextAutofit
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

/** ECMA-376 Part 1 §20.1.9.22 default adjust values for preset text warps.
 *
 * Supported (regulated) presets take their defaults from the pinned official
 * catalog — avLst DEFAULT FORMULAS evaluated at a nominal box — so parse-side
 * models and resolve-side geometry can never disagree (matrix G-case parity).
 * Unsupported legacy presets keep their previous table entries (documented;
 * not routed through geometry, so their values are display-only). */
const OFFICIAL_SUPPORTED_DEFAULTS: Readonly<Record<string, Record<string, number>>> = Object.fromEntries(
  Object.entries(TEXT_WARP_CATALOG).map(([preset, definition]) => {
    // Official avLst defaults evaluate through the SAME shared guide resolver
    // the warp engine uses (plain vals and seeded formulas both); never the
    // shape preset catalog.
    const values = evaluateGuides(
      definition.adjustments as unknown as ReadonlyArray<[string, string]>,
      definition.guides as unknown as ReadonlyArray<[string, string]>,
      textWarpSeedValues(100, 100),
      {},
      [],
    )
    const out: Record<string, number> = {}
    for (const [name] of definition.adjustments) {
      const value = values.get(name)
      if (value !== undefined) out[name] = value
    }
    return [preset, out]
  }),
)

const LEGACY_DEFAULTS: Readonly<Record<string, Record<string, number>>> = {
  textDoubleWave1: { adj1: 0, adj2: 50000 },
  textWave4: { adj1: 0, adj2: 50000 },
  textInflateBottom: { adj: 50000 },
  textDeflateBottom: { adj: 50000 },
  textInflateTop: { adj: 50000 },
  textDeflateTop: { adj: 50000 },
  textDeflateInflate: { adj: 50000 },
  textDeflateInflateDeflate: { adj: 50000 },
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

/** Parse-side default table. Regulated (official-geometry) presets are
 * overridden from the pinned catalog below, so this hand table can never
 * shadow the authoritative values. */
export const DEFAULT_WARP_ADJUSTMENTS: Readonly<Record<string, Record<string, number>>> = {
  ...LEGACY_DEFAULTS,
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

/** Regulated presets ALWAYS take catalog defaults (single source of truth for
 * geometry and parse-like callers): the hand entries above are display-only. */
Object.assign(DEFAULT_WARP_ADJUSTMENTS as Record<string, Record<string, number>>, OFFICIAL_SUPPORTED_DEFAULTS)

/** Parses guide adjustments from an <a:avLst> XML node. */
export function parseAdjustGuides(avLst: XmlNode | undefined): Record<string, number> {
  const adjustments: Record<string, number> = {}
  if (!avLst) return adjustments
  for (const gd of getChildren(avLst, 'gd')) {
    const ga = attrs(gd)
    if (ga.name && ga.fmla) {
      const raw = ga.fmla.startsWith('val ') ? ga.fmla.slice(4).trim() : ga.fmla.trim()
      if (raw.length > 0) {
        const valNum = Number(raw)
        if (Number.isFinite(valNum)) {
          adjustments[ga.name] = valNum
        }
      }
    }
  }
  return adjustments
}

/**
 * Evaluates whether a textPlain warp node has a nondefault adjustment.
 * Default is adj = 50000 (unwarped). If adj is absent or 50000, returns undefined.
 * If nondefault adj is specified, returns the numeric adjustment.
 */
export function parseTextPlainAdjustment(prstWarpNode: XmlNode | undefined): number | undefined {
  if (!prstWarpNode) return undefined
  const avLst = getChildren(prstWarpNode, 'avLst')[0]
  const guides = parseAdjustGuides(avLst)
  const adj = guides['adj']
  if (adj !== undefined && adj !== 50000) {
    return adj
  }
  return undefined
}

