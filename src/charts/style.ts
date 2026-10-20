/**
 * Chart styles, colour precedence, line styles and point overrides (spec §5.4, §5.8, §5.11).
 */
import { attrs, getChild, type XmlNode } from '../core/xml'
import { resolveDrawingColor, type ThemeContext } from '../drawing/style'
import { parseFillStyle, applyFillStyle, type ChartFillStyle, type Box } from './fills'
import type { ChartMarkerStyle } from './markers'

export interface ChartLineStyle {
  width?: number
  color?: string
  fill?: ChartFillStyle
  dash?: string | number[]
  cap?: CanvasLineCap
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

export interface SeriesColorOptions {
  spPrColor?: string
  chartStyleColor?: string
  colorStyleColor?: string
  styleId?: number
  seriesIndex: number
  theme?: ThemeContext
}

export const DEFAULT_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

const PRESET_DASHES: Record<string, number[]> = {
  solid: [],
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
}

const CAPS_MAP: Record<string, CanvasLineCap> = {
  flat: 'butt',
  rnd: 'round',
  sq: 'square',
}

function hexOf(color: { r: number; g: number; b: number } | undefined): string | undefined {
  if (!color) return undefined
  return [color.r, color.g, color.b]
    .map(c => Math.max(0, Math.min(255, Math.round(c))).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()
}

/** Parse line style from DrawingML a:ln node. */
export function parseLineStyle(ln: XmlNode | undefined, theme?: ThemeContext): ChartLineStyle | undefined {
  if (!ln) return undefined
  const noFill = Boolean(getChild(ln, 'noFill'))

  const lnAttrs = attrs(ln)
  const wRaw = lnAttrs.w
  // w in EMUs: 1 pt = 12700 EMUs, 1 pt = 96/72 px = 4/3 px
  const width = wRaw !== undefined && Number.isFinite(Number(wRaw)) ? (Number(wRaw) / 12700) * (96 / 72) : undefined
  const capRaw = lnAttrs.cap
  const cap = capRaw && CAPS_MAP[capRaw] ? CAPS_MAP[capRaw] : undefined

  const prstDash = getChild(ln, 'prstDash')
  const dash = prstDash ? attrs(prstDash).val ?? 'solid' : undefined

  const fill = !noFill ? parseFillStyle(ln, theme) : undefined
  const solidColor = fill?.kind === 'solid' ? fill.color : undefined

  return {
    ...(width !== undefined ? { width } : {}),
    ...(cap ? { cap } : {}),
    ...(dash ? { dash } : {}),
    ...(solidColor ? { color: solidColor } : {}),
    ...(fill ? { fill } : {}),
    noFill,
  }
}

/** Apply line stroke style to 2D canvas context. */
export function applyLineStyle(
  ctx: CanvasRenderingContext2D,
  line: ChartLineStyle | undefined,
  bounds?: Box,
): boolean {
  if (!line || line.noFill) return false

  const width = line.width ?? 1
  ctx.lineWidth = width
  ctx.lineCap = line.cap ?? 'butt'

  // Dash
  let dashArray: number[] = []
  if (Array.isArray(line.dash)) {
    dashArray = line.dash
  } else if (typeof line.dash === 'string') {
    dashArray = PRESET_DASHES[line.dash] ?? []
  }
  ctx.setLineDash(dashArray.map(d => d * width))

  // Color / fill
  if (line.fill && bounds) {
    applyFillStyle(ctx, line.fill, bounds)
    ctx.strokeStyle = ctx.fillStyle
  } else if (line.color) {
    ctx.strokeStyle = line.color.startsWith('#') ? line.color : `#${line.color}`
  } else {
    ctx.strokeStyle = '#000000'
  }

  return true
}

/**
 * Colour precedence (spec §5.11):
 * explicit spPr -> chartStyle part -> colorStyle part -> c:style index -> theme accents with lumMod variants after 6.
 */
export function resolveSeriesColor(options: SeriesColorOptions): string {
  const { spPrColor, chartStyleColor, colorStyleColor, seriesIndex, theme } = options

  if (spPrColor) return spPrColor.replace(/^#/, '').toUpperCase()
  if (chartStyleColor) return chartStyleColor.replace(/^#/, '').toUpperCase()
  if (colorStyleColor) return colorStyleColor.replace(/^#/, '').toUpperCase()

  const s = Math.abs(Number.isSafeInteger(seriesIndex) ? seriesIndex : 0)
  const cycle = Math.floor(s / 6)
  const slot = s % 6
  const schemeName = `accent${slot + 1}`

  if (cycle === 0) {
    const raw = theme?.palette[schemeName]?.replace(/^#/, '') ?? DEFAULT_ACCENTS[slot]
    return raw.toUpperCase()
  }

  // lumMod/lumOff variations after 6 series (Office accent cycling)
  // Cycle 1 (series 6-11): lumMod: 60%, lumOff: 40% (lighter variant)
  // Cycle 2 (series 12-17): lumMod: 40% (darker variant)
  // Cycle 3 (series 18-23): lumMod: 80% (slight tint)
  const transforms: Array<{ name: string; value?: string }> = []
  if (cycle % 3 === 1) {
    transforms.push({ name: 'lumMod', value: '60000' }, { name: 'lumOff', value: '40000' })
  } else if (cycle % 3 === 2) {
    transforms.push({ name: 'lumMod', value: '40000' })
  } else {
    transforms.push({ name: 'lumMod', value: '75000' }, { name: 'lumOff', value: '25000' })
  }

  const resolved = resolveDrawingColor({ kind: 'scheme', value: schemeName, transforms }, theme)
  const hex = hexOf(resolved)
  if (hex) return hex

  // Fallback if no theme available: manual HSL adjust on DEFAULT_ACCENTS
  const baseHex = DEFAULT_ACCENTS[slot]
  const r = parseInt(baseHex.slice(0, 2), 16)
  const g = parseInt(baseHex.slice(2, 4), 16)
  const b = parseInt(baseHex.slice(4, 6), 16)
  const factor = cycle % 2 === 1 ? 0.7 : 0.5
  const modR = Math.round(r * factor).toString(16).padStart(2, '0')
  const modG = Math.round(g * factor).toString(16).padStart(2, '0')
  const modB = Math.round(b * factor).toString(16).padStart(2, '0')
  return `${modR}${modG}${modB}`.toUpperCase()
}

/** Invert a colour for negative bar fill when invertIfNegative is active. */
export function invertColor(hex: string): string {
  const clean = hex.replace(/^#/, '')
  if (clean.length !== 6) return 'FFFFFF'
  const r = 255 - parseInt(clean.slice(0, 2), 16)
  const g = 255 - parseInt(clean.slice(2, 4), 16)
  const b = 255 - parseInt(clean.slice(4, 6), 16)
  return [r, g, b].map(c => Math.max(0, Math.min(255, c)).toString(16).padStart(2, '0')).join('').toUpperCase()
}
