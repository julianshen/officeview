/**
 * Chart fills: solid, gradient (linear and path), pattern, and noFill (spec §5.4).
 * Reuses Canvas pattern tile rasterizer from src/drawing/pattern.ts.
 */
import { attrs, getChild, getChildren, type XmlNode } from '../core/xml'
import { getPatternTile } from '../drawing/pattern'
import { parseDrawingColor, resolveDrawingColor, type ThemeContext, type DrawingColor } from '../drawing/style'
import { CHART_GRADIENT_STOP_LIMIT, chargeChartFill, type ChartCharge } from './own'

export interface SolidFill {
  kind: 'solid'
  color: string
  alpha?: number
}

export interface GradientStop {
  position: number
  color: string
  alpha?: number
}

export interface GradientFill {
  kind: 'gradient'
  gradient: 'linear' | 'path'
  angle?: number
  stops: GradientStop[]
}

export interface PatternFill {
  kind: 'pattern'
  preset: string
  fgColor: string
  bgColor: string
}

export interface NoFill {
  kind: 'none'
}

export type ChartFillStyle = SolidFill | GradientFill | PatternFill | NoFill

export interface Box {
  x: number
  y: number
  w: number
  h: number
}

/** The one hex helper of the chart formatting code (CSS hex without `#`, upper case). */
export function hexColor(color: DrawingColor | undefined): string | undefined {
  if (!color) return undefined
  const r = Math.max(0, Math.min(255, Math.round(color.r))).toString(16).padStart(2, '0')
  const g = Math.max(0, Math.min(255, Math.round(color.g))).toString(16).padStart(2, '0')
  const b = Math.max(0, Math.min(255, Math.round(color.b))).toString(16).padStart(2, '0')
  return `${r}${g}${b}`.toUpperCase()
}

/** A CSS colour for a hex colour and its DrawingML alpha (opaque `#hex` when there is none). */
export function toCssColor(color: string, alpha?: number): string {
  if (color.startsWith('#') || color.startsWith('rgb')) {
    if (alpha !== undefined && alpha < 1 && color.startsWith('#') && color.length === 7) {
      const r = parseInt(color.slice(1, 3), 16)
      const g = parseInt(color.slice(3, 5), 16)
      const b = parseInt(color.slice(5, 7), 16)
      return `rgba(${r}, ${g}, ${b}, ${alpha})`
    }
    return color
  }
  if (alpha !== undefined && alpha < 1 && /^[0-9A-Fa-f]{6}$/.test(color)) {
    const r = parseInt(color.slice(0, 2), 16)
    const g = parseInt(color.slice(2, 4), 16)
    const b = parseInt(color.slice(4, 6), 16)
    return `rgba(${r}, ${g}, ${b}, ${alpha})`
  }
  return `#${color}`
}

/**
 * Parse the fill of an spPr (or a direct fill container). Each retained fill is charged through `charge`
 * (charts/own.ts) before it is built; a gradient keeps at most CHART_GRADIENT_STOP_LIMIT stops, each charged.
 */
export function parseFillStyle(node: XmlNode | undefined, theme?: ThemeContext, charge?: ChartCharge): ChartFillStyle | undefined {
  if (!node) return undefined
  if (getChild(node, 'noFill')) { chargeChartFill(charge); return { kind: 'none' } }

  const solidNode = getChild(node, 'solidFill')
  if (solidNode) {
    const resolved = resolveDrawingColor(parseDrawingColor(solidNode), theme)
    const hex = hexColor(resolved)
    if (hex) { chargeChartFill(charge); return { kind: 'solid', color: hex, ...(resolved && resolved.a < 1 ? { alpha: resolved.a } : {}) } }
  }

  const gradNode = getChild(node, 'gradFill')
  if (gradNode) {
    const gs = getChildren(getChild(gradNode, 'gsLst'), 'gs').slice(0, CHART_GRADIENT_STOP_LIMIT)
    chargeChartFill(charge, Math.max(gs.length, 2))
    const stops: GradientStop[] = []
    for (const stop of gs) {
      const posRaw = Number(attrs(stop).pos)
      const resolved = resolveDrawingColor(parseDrawingColor(stop), theme)
      stops.push({ position: Number.isFinite(posRaw) ? Math.max(0, Math.min(1, posRaw / 100000)) : 0, color: hexColor(resolved) ?? '000000', ...(resolved && resolved.a < 1 ? { alpha: resolved.a } : {}) })
    }
    const lin = getChild(gradNode, 'lin')
    const angleRaw = Number(lin ? attrs(lin).ang : undefined)
    return {
      kind: 'gradient',
      gradient: getChild(gradNode, 'path') ? 'path' : 'linear',
      angle: Number.isFinite(angleRaw) ? angleRaw / 60000 : 90,
      stops: stops.length ? stops : [{ position: 0, color: 'FFFFFF' }, { position: 1, color: '000000' }],
    }
  }

  const pattNode = getChild(node, 'pattFill')
  if (pattNode) {
    const fgNode = getChild(pattNode, 'fgClr'), bgNode = getChild(pattNode, 'bgClr')
    chargeChartFill(charge)
    return {
      kind: 'pattern',
      preset: attrs(pattNode).prst ?? 'pct50',
      fgColor: hexColor(resolveDrawingColor(fgNode ? parseDrawingColor(fgNode) : undefined, theme)) ?? '000000',
      bgColor: hexColor(resolveDrawingColor(bgNode ? parseDrawingColor(bgNode) : undefined, theme)) ?? 'FFFFFF',
    }
  }

  return undefined
}

/** Apply fill style to canvas context. Returns true if a fill is active, false if noFill / transparent. */
export function applyFillStyle(ctx: CanvasRenderingContext2D, fill: ChartFillStyle | undefined, bounds: Box): boolean {
  if (!fill || fill.kind === 'none') return false

  if (fill.kind === 'solid') {
    ctx.fillStyle = toCssColor(fill.color, fill.alpha)
    return true
  }

  if (fill.kind === 'pattern') {
    const fgCss = toCssColor(fill.fgColor)
    const bgCss = toCssColor(fill.bgColor)
    const tile = getPatternTile(ctx, fill.preset, fgCss, bgCss)
    if (tile) {
      try {
        const pattern = ctx.createPattern(tile, 'repeat')
        if (pattern) {
          ctx.fillStyle = pattern
          return true
        }
      } catch {
        // Fallback to fgColor on pattern failure
      }
    }
    ctx.fillStyle = fgCss
    return true
  }

  if (fill.kind === 'gradient') {
    const w = Math.max(1, bounds.w)
    const h = Math.max(1, bounds.h)
    let grad: CanvasGradient
    if (fill.gradient === 'path') {
      const cx = bounds.x + w / 2
      const cy = bounds.y + h / 2
      const r = Math.max(w, h) / 2
      grad = ctx.createRadialGradient(cx, cy, 0, cx, cy, r)
    } else {
      const rad = ((fill.angle ?? 90) * Math.PI) / 180
      const dx = Math.cos(rad)
      const dy = Math.sin(rad)
      const span = Math.abs(w * dx) + Math.abs(h * dy)
      const cx = bounds.x + w / 2
      const cy = bounds.y + h / 2
      const x0 = cx - (dx * span) / 2
      const y0 = cy - (dy * span) / 2
      const x1 = cx + (dx * span) / 2
      const y1 = cy + (dy * span) / 2
      grad = ctx.createLinearGradient(x0, y0, x1, y1)
    }
    const stops = [...fill.stops].sort((a, b) => a.position - b.position)
    for (const stop of stops) {
      grad.addColorStop(stop.position, toCssColor(stop.color, stop.alpha))
    }
    ctx.fillStyle = grad
    return true
  }

  return false
}
