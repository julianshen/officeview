/**
 * Chart markers: symbols, auto cycle, size clamp 2-72, marker spPr (spec §5.4).
 */
import { attrs, getChild, type XmlNode } from '../core/xml'
import type { ThemeContext } from '../drawing/style'
import { parseFillStyle, type ChartFillStyle } from './fills'
import { parseLineStyle, type ChartLineStyle } from './style'

export type MarkerSymbol =
  | 'circle'
  | 'square'
  | 'diamond'
  | 'triangle'
  | 'x'
  | 'star'
  | 'dash'
  | 'dot'
  | 'plus'
  | 'none'
  | 'auto'

export const MARKER_AUTO_CYCLE: readonly MarkerSymbol[] = [
  'square',
  'diamond',
  'triangle',
  'x',
  'star',
  'dot',
  'dash',
  'circle',
  'plus',
] as const

export interface ChartMarkerStyle {
  symbol?: MarkerSymbol
  size?: number
  fill?: ChartFillStyle
  line?: ChartLineStyle
}

/** Clamp marker size to 2-72 per schema (spec §5.4). */
export function clampMarkerSize(size: number | undefined, defaultSize = 5): number {
  if (size === undefined || !Number.isFinite(size)) return defaultSize
  return Math.max(2, Math.min(72, Math.round(size)))
}

/** Resolve marker symbol against the series auto cycle. */
export function resolveMarkerSymbol(symbol: MarkerSymbol | undefined, seriesIndex: number): MarkerSymbol {
  if (!symbol || symbol === 'auto') {
    const idx = Math.abs(Number.isSafeInteger(seriesIndex) ? seriesIndex : 0)
    return MARKER_AUTO_CYCLE[idx % MARKER_AUTO_CYCLE.length]
  }
  return symbol
}

const KNOWN_SYMBOLS = new Set<string>([
  'circle', 'square', 'diamond', 'triangle', 'x', 'star', 'dash', 'dot', 'plus', 'none', 'auto',
])

export function parseMarker(node: XmlNode | undefined, theme?: ThemeContext): ChartMarkerStyle | undefined {
  if (!node) return undefined
  const symVal = attrs(getChild(node, 'symbol')).val
  const symbol = symVal && KNOWN_SYMBOLS.has(symVal) ? (symVal as MarkerSymbol) : undefined
  const sizeRaw = attrs(getChild(node, 'size')).val
  const size = sizeRaw !== undefined && Number.isFinite(Number(sizeRaw)) ? clampMarkerSize(Number(sizeRaw)) : undefined
  const spPr = getChild(node, 'spPr')
  const fill = parseFillStyle(spPr, theme)
  const line = parseLineStyle(spPr ? getChild(spPr, 'ln') : undefined, theme)

  if (!symbol && size === undefined && !fill && !line) return undefined
  return {
    ...(symbol ? { symbol } : {}),
    ...(size !== undefined ? { size } : {}),
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {}),
  }
}

/** Draw a marker at (cx, cy) with the given symbol and radius r = size / 2. */
export function paintMarker(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  symbol: MarkerSymbol,
  size: number,
  hasFill = true,
  hasStroke = false,
): void {
  if (symbol === 'none' || size <= 0) return
  const r = size / 2

  ctx.beginPath()
  switch (symbol) {
    case 'circle':
      ctx.arc(cx, cy, r, 0, Math.PI * 2)
      if (hasFill) ctx.fill()
      if (hasStroke) ctx.stroke()
      break

    case 'square':
      ctx.rect(cx - r, cy - r, size, size)
      if (hasFill) ctx.fill()
      if (hasStroke) ctx.stroke()
      break

    case 'diamond':
      ctx.moveTo(cx, cy - r)
      ctx.lineTo(cx + r, cy)
      ctx.lineTo(cx, cy + r)
      ctx.lineTo(cx - r, cy)
      ctx.closePath()
      if (hasFill) ctx.fill()
      if (hasStroke) ctx.stroke()
      break

    case 'triangle':
      ctx.moveTo(cx, cy - r)
      ctx.lineTo(cx + r, cy + r)
      ctx.lineTo(cx - r, cy + r)
      ctx.closePath()
      if (hasFill) ctx.fill()
      if (hasStroke) ctx.stroke()
      break

    case 'x':
      ctx.moveTo(cx - r, cy - r)
      ctx.lineTo(cx + r, cy + r)
      ctx.moveTo(cx - r, cy + r)
      ctx.lineTo(cx + r, cy - r)
      ctx.stroke()
      break

    case 'plus':
      ctx.moveTo(cx - r, cy)
      ctx.lineTo(cx + r, cy)
      ctx.moveTo(cx, cy - r)
      ctx.lineTo(cx, cy + r)
      ctx.stroke()
      break

    case 'dash':
      ctx.moveTo(cx - r, cy)
      ctx.lineTo(cx + r, cy)
      ctx.stroke()
      break

    case 'dot':
      ctx.arc(cx, cy, Math.max(1, r * 0.45), 0, Math.PI * 2)
      ctx.fill()
      break

    case 'star': {
      // 5-point star
      for (let i = 0; i < 10; i++) {
        const rad = (i * Math.PI) / 5 - Math.PI / 2
        const d = i % 2 === 0 ? r : r * 0.45
        const px = cx + d * Math.cos(rad)
        const py = cy + d * Math.sin(rad)
        if (i === 0) ctx.moveTo(px, py)
        else ctx.lineTo(px, py)
      }
      ctx.closePath()
      if (hasFill) ctx.fill()
      if (hasStroke) ctx.stroke()
      break
    }

    default:
      // Default to square
      ctx.rect(cx - r, cy - r, size, size)
      if (hasFill) ctx.fill()
      if (hasStroke) ctx.stroke()
      break
  }
}
