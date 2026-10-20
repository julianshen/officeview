/**
 * Chart line formatting, colour precedence and point overrides (spec 5.4, 5.8, 5.11; lazy chunk).
 *
 * Colour precedence for a series (or, under varyColors, a point) without an explicit `spPr` colour:
 *
 * 1. the chart style part (`cs:chartStyle`): its `dataPoint` (or `dataPointLine` for line and scatter)
 *    entry: an explicit colour wins; a `phClr` colour takes the `styleClr` colour (an index into the colour
 *    part, or `auto`: the next level) with the entry's own transforms applied;
 * 2. the colour part (`cs:colorStyle`): `cycle` (and unknown methods) cycle its colour list per series, each
 *    further cycle taking the part's next `cs:variation`; `withinLinear`/`acrossLinear` (and *Reversed) select
 *    the first / modulo colour of the variation-expanded total set with an approximated brightness ramp (see `linearBrightness`);
 * 3. `c:style` (1-48): the gallery column ((n-1) mod 8) picks greyscale (0), the six accents (1), or a single
 *    accent (2-7). The single-colour columns spread a shade/tint over the series in (-70%, 70%), as LibreOffice's
 *    `oox` import approximates Office 2007 (`DetailFormatterBase::getPhColor`); not verified against Office;
 * 4. the theme accent cycle: accent1-6, each later cycle taking the next variation of Office's default colour
 *    part (colors1.xml, id 10): none, lumMod 60%, lumMod 80% + lumOff 20%, lumMod 80%, lumMod 60% + lumOff 40%,
 *    lumMod 50%, lumMod 70% + lumOff 30%, lumMod 70%, lumMod 50% + lumOff 50%.
 */
import { attrs, getChild, type XmlNode } from '../core/xml'
import { resolveDrawingColor, type ColorDefinition, type ThemeContext } from '../drawing/style'
import { applyFillStyle, hexColor, parseFillStyle, toCssColor, type ChartFillStyle, type Box } from './fills'
import type { ChartMarkerStyle } from './markers'
import { chargeChartLine, type ChartCharge } from './own'

export interface ChartLineStyle {
  /** CSS px (the `w` EMUs at 96 dpi). */
  width?: number
  color?: string
  fill?: ChartFillStyle
  dash?: string
  cap?: CanvasLineCap
  join?: CanvasLineJoin
  noFill?: boolean
}

export interface ChartPointOverride {
  fill?: ChartFillStyle
  color?: string
  line?: ChartLineStyle
  marker?: ChartMarkerStyle
  explosion?: number
  invertIfNegative?: boolean
}

type Transform = ColorDefinition['transforms'][number]
/** A colour part (`cs:colorStyle`), read before the chart is parsed (charts/styleParts.ts). */
export interface ChartColorPart { method: string; colors: ColorDefinition[]; variations: Transform[][] }
/** One entry of a chart style part: its spPr colour (fill, or line for `dataPointLine`) and the fillRef/lnRef styleClr. */
export interface ChartStyleEntry { color?: ColorDefinition; styleColor?: 'auto' | number }
/** A style part's text defaults (`cs:title`: `cs:defRPr` size/bold/italic, the `cs:fontRef` or defRPr colour; spec 5.10). */
export interface ChartStyleText { sizePt?: number; bold?: boolean; italic?: boolean; color?: ColorDefinition; /** `cs:fontRef idx` (`major`/`minor`) or a defRPr `a:latin` typeface. */ font?: string; underline?: boolean }
export interface ChartStylePart { dataPoint?: ChartStyleEntry; dataPointLine?: ChartStyleEntry; title?: ChartStyleText }
export interface ChartStyleParts { color?: ChartColorPart; style?: ChartStylePart }

export interface SeriesColorOptions {
  /** The series index (c:idx) or, under varyColors, the point index. */
  index: number
  /** How many series (or points) share the cycle: the single-colour `c:style` spread depends on it. */
  count?: number
  /** Line and scatter series read the chart style part's `dataPointLine` entry. */
  line?: boolean
  styleId?: number
  parts?: ChartStyleParts
  theme?: ThemeContext
}

export const DEFAULT_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']
const t = (name: string, value: number): Transform => ({ name, value: String(value) })
/** Office's default colour part (colors1.xml, `meth="cycle" id="10"`): the variation of each later accent cycle. */
export const OFFICE_ACCENT_VARIATIONS: readonly Transform[][] = [
  [], [t('lumMod', 60000)], [t('lumMod', 80000), t('lumOff', 20000)], [t('lumMod', 80000)], [t('lumMod', 60000), t('lumOff', 40000)],
  [t('lumMod', 50000)], [t('lumMod', 70000), t('lumOff', 30000)], [t('lumMod', 70000)], [t('lumMod', 50000), t('lumOff', 50000)],
]

const PRESET_DASHES: Record<string, number[]> = {
  solid: [], dot: [1, 3], dash: [4, 3], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
}
const CAPS: Record<string, CanvasLineCap> = { flat: 'butt', rnd: 'round', sq: 'square' }

/** Parse a DrawingML `a:ln` (charged through `charge` before it is built). */
export function parseLineStyle(ln: XmlNode | undefined, theme?: ThemeContext, charge?: ChartCharge): ChartLineStyle | undefined {
  if (!ln) return undefined
  chargeChartLine(charge)
  const noFill = Boolean(getChild(ln, 'noFill'))
  const w = Number(attrs(ln).w)
  const cap = CAPS[attrs(ln).cap ?? '']
  const dash = attrs(getChild(ln, 'prstDash')).val
  const join: CanvasLineJoin | undefined = getChild(ln, 'round') ? 'round' : getChild(ln, 'bevel') ? 'bevel' : getChild(ln, 'miter') ? 'miter' : undefined
  // The line's own fill (its stroke colour), charged with its kept stops before it is built.
  const fill = noFill ? undefined : parseFillStyle(ln, theme, charge)
  return {
    ...(attrs(ln).w !== undefined && Number.isFinite(w) && w >= 0 ? { width: w / 12700 * 96 / 72 } : {}),
    ...(cap ? { cap } : {}),
    ...(join ? { join } : {}),
    ...(dash && PRESET_DASHES[dash] ? { dash } : {}),
    ...(fill?.kind === 'solid' ? { color: fill.color } : {}),
    ...(fill && fill.kind !== 'none' ? { fill } : {}),
    noFill,
  }
}

/** Whether a line carries a stroke colour of its own (a solid colour or another fill). */
export const lineHasColor = (line: ChartLineStyle | undefined): boolean => Boolean(line && !line.noFill && (line.color || line.fill))

/**
 * Applies an authored line over the stroke state the caller already set (the series colour and default
 * width/join): only what the `a:ln` states replaces it, so a width without a colour keeps the series colour.
 * Returns false for a `noFill` line (nothing to stroke).
 */
export function applyLineStyle(ctx: CanvasRenderingContext2D, line: ChartLineStyle | undefined, bounds?: Box): boolean {
  if (!line) return true
  if (line.noFill) return false
  if (line.width !== undefined) ctx.lineWidth = line.width
  if (line.cap) ctx.lineCap = line.cap
  if (line.join) ctx.lineJoin = line.join
  if (line.dash) ctx.setLineDash((PRESET_DASHES[line.dash] ?? []).map(d => d * ctx.lineWidth))
  if (line.fill && line.fill.kind !== 'solid' && bounds) {
    const previous = ctx.fillStyle
    if (applyFillStyle(ctx, line.fill, bounds)) ctx.strokeStyle = ctx.fillStyle
    ctx.fillStyle = previous
  } else if (line.color) ctx.strokeStyle = toCssColor(line.color, line.fill?.kind === 'solid' ? line.fill.alpha : undefined)
  return true
}

const scheme = (value: string, transforms: Transform[] = []): ColorDefinition => ({ kind: 'scheme', value, transforms })
/** A colour with extra transforms appended (a variation), resolved; a theme-less accent falls back to Office's default. */
function resolve(color: ColorDefinition, theme: ThemeContext | undefined, extra: Transform[] = [], placeholder?: ColorDefinition): string | undefined {
  const withExtra = (c: ColorDefinition): ColorDefinition => ({ ...c, transforms: [...c.transforms, ...extra] })
  // A scheme colour reads the theme's resolved palette (as the accent cycle always has); without a theme slot,
  // an accent takes Office's default accent and tx1/dk1 black.
  const fallback = (c: ColorDefinition): ColorDefinition => {
    if (c.kind !== 'scheme' || c.value === 'phClr') return c
    const name = c.value ?? '', mapped = theme?.colorMap?.[name] ?? name
    const known = (theme?.palette?.[mapped] ?? theme?.palette?.[name])?.replace(/^#/, '')
    if (known) return { kind: 'srgb', value: known, transforms: c.transforms }
    if (theme?.colors?.[mapped]) return c
    const slot = /^accent([1-6])$/.exec(name)
    if (slot) return { kind: 'srgb', value: DEFAULT_ACCENTS[Number(slot[1]) - 1], transforms: c.transforms }
    if (mapped === 'dk1' || mapped === 'tx1') return { kind: 'srgb', value: '000000', transforms: c.transforms }
    return c
  }
  // A partial theme (no colour scheme) resolves like none.
  const usable = theme?.colors && theme.colorMap ? theme : undefined
  const base = placeholder ? resolveDrawingColor(fallback(withExtra(placeholder)), usable) : undefined
  return hexColor(resolveDrawingColor(fallback(placeholder ? color : withExtra(color)), usable, base))
}

/**
 * The brightness ramp of the linear methods. MS-ODRAWXML 2.8.4.2 specifies the colour selection and that
 * brightness progresses with index/count (reversed for the *Reversed methods) but gives no formula: this
 * symmetric lumMod/lumOff ramp, darker (-LINEAR_SPREAD) to lighter (+LINEAR_SPREAD) across the objects, is an
 * APPROXIMATION, not verified against Office.
 */
export const LINEAR_SPREAD = 0.5
function linearBrightness(index: number, count: number, reversed: boolean): Transform[] {
  const progress = count > 1 ? Math.min(index, count - 1) / (count - 1) : 0.5
  const b = Math.round(((reversed ? 1 - progress : progress) * 2 - 1) * LINEAR_SPREAD * 100000)
  return b < 0 ? [t('lumMod', 100000 + b)] : b > 0 ? [t('lumMod', 100000 - b), t('lumOff', b)] : []
}

/**
 * Level 2: the colour part's colour for `index` of `count` objects (MS-ODRAWXML 2.8.4.2): `withinLinear` takes the
 * first colour of the variation-expanded total set (2.8.3.2) for every object, `acrossLinear` the colour modulo
 * that total set, both with the brightness ramp above
 * (the *Reversed methods reverse it); `cycle` and unknown methods (2.8.4.1) cycle with the part's variations.
 */
function colorPartColor(part: ChartColorPart, index: number, count: number): ColorDefinition | undefined {
  if (!part.colors.length) return undefined
  const n = part.colors.length, total = n * Math.max(1, part.variations.length)
  // Logical slot k of the variation-expanded total set (2.8.3.2; colour fastest, then variation), never materialized.
  // Transform order: the colour's own, then the variation, then (linear methods) the brightness approximation.
  const slot = (k: number, extra: Transform[] = []): ColorDefinition => {
    const color = part.colors[k % n], variation = part.variations.length ? part.variations[Math.floor(k / n) % part.variations.length] : []
    return { ...color, transforms: [...color.transforms, ...variation, ...extra] }
  }
  const linear = /^(within|across)Linear(Reversed)?$/.exec(part.method)
  if (linear) return slot(linear[1] === 'within' ? 0 : index % total, linearBrightness(index, count, Boolean(linear[2])))
  return slot(index)
}

/** Levels 3 and 4: `c:style`, then the theme accent cycle with Office's default variations. */
function styleColor(index: number, count: number, styleId: number | undefined): ColorDefinition {
  const column = styleId !== undefined && Number.isInteger(styleId) && styleId >= 1 && styleId <= 48 ? (styleId - 1) % 8 : 1
  if (column !== 1) {
    // One colour per series: shade the leading series, tint the trailing ones, over (-70%, 70%) (LibreOffice oox).
    const v = Math.round(((index + 1) / (Math.max(count, index + 1) + 1) * 1.4 - 0.7) * 100000)
    const base = column === 0 ? 'tx1' : `accent${column - 1}`
    return scheme(base, v < 0 ? [t('shade', 100000 + v)] : v > 0 ? [t('tint', 100000 - v)] : [])
  }
  const cycle = Math.floor(index / 6)
  return scheme(`accent${(index % 6) + 1}`, [...OFFICE_ACCENT_VARIATIONS[cycle % OFFICE_ACCENT_VARIATIONS.length]])
}

/** The automatic colour of a series (or point) with no explicit `spPr` colour (spec 5.11; levels 1-4 above). */
export function resolveSeriesColor(options: SeriesColorOptions): string {
  const { theme, parts } = options
  const index = Number.isSafeInteger(options.index) && options.index >= 0 ? options.index : 0
  const count = options.count ?? index + 1
  const below = (): ColorDefinition => (parts?.color && colorPartColor(parts.color, index, count)) || styleColor(index, count, options.styleId)
  const entry = options.line ? parts?.style?.dataPointLine ?? parts?.style?.dataPoint : parts?.style?.dataPoint
  if (entry && (entry.color || typeof entry.styleColor === 'number')) {
    // The styleClr colour (an index into the colour part, or auto) is the entry's phClr.
    const indexed = typeof entry.styleColor === 'number' && parts?.color?.colors.length ? parts.color.colors[entry.styleColor % parts.color.colors.length] : undefined
    const ph = indexed ?? below()
    const hex = entry.color ? resolve(entry.color, theme, [], entry.color.value === 'phClr' ? ph : undefined) : resolve(ph, theme)
    if (hex) return hex
  }
  return resolve(below(), theme) ?? DEFAULT_ACCENTS[index % 6]
}
