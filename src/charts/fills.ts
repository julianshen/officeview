/**
 * Chart fills: solid, gradient (linear and path), pattern, and noFill (spec §5.4).
 * Reuses Canvas pattern tile rasterizer from src/drawing/pattern.ts.
 */
import { attrs, getChild, getChildren, type XmlNode } from '../core/xml'
import { getPatternTile } from '../drawing/pattern'
import { parseDrawingColor, resolveDrawingColor, type ThemeContext, type DrawingColor } from '../drawing/style'

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

function hexColor(color: DrawingColor | undefined): string | undefined {
  if (!color) return undefined
  const r = Math.max(0, Math.min(255, Math.round(color.r))).toString(16).padStart(2, '0')
  const g = Math.max(0, Math.min(255, Math.round(color.g))).toString(16).padStart(2, '0')
  const b = Math.max(0, Math.min(255, Math.round(color.b))).toString(16).padStart(2, '0')
  return `${r}${g}${b}`.toUpperCase()
}

function toCssColor(color: string, alpha?: number): string {
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

/** Parse fill from spPr node or direct fill container. */
export function parseFillStyle(node: XmlNode | undefined, theme?: ThemeContext): ChartFillStyle | undefined {
  if (!node) return undefined
  if (getChild(node, 'noFill')) return { kind: 'none' }

  // 1. solidFill
  const solidNode = getChild(node, 'solidFill')
  if (solidNode) {
    const parsed = parseDrawingColor(solidNode)
    const resolved = resolveDrawingColor(parsed, theme)
    const hex = hexColor(resolved)
    if (hex) return { kind: 'solid', color: hex, alpha: resolved?.a }
  }

  // 2. gradFill
  const gradNode = getChild(node, 'gradFill')
  if (gradNode) {
    const gsLst = getChild(gradNode, 'gsLst')
    const stops: GradientStop[] = []
    if (gsLst) {
      for (const gs of getChildren(gsLst, 'gs')) {
        const posRaw = attrs(gs).pos
        const pos = posRaw !== undefined ? Math.max(0, Math.min(1, Number(posRaw) / 100000)) : 0
        const parsed = parseDrawingColor(gs)
        const resolved = resolveDrawingColor(parsed, theme)
        const hex = hexColor(resolved) ?? '000000'
        stops.push({ position: pos, color: hex, alpha: resolved?.a })
      }
    }
    const lin = getChild(gradNode, 'lin')
    const path = getChild(gradNode, 'path')
    const angleRaw = lin ? attrs(lin).ang : undefined
    const angle = angleRaw !== undefined ? Number(angleRaw) / 60000 : 90
    return {
      kind: 'gradient',
      gradient: path ? 'path' : 'linear',
      angle,
      stops: stops.length ? stops : [{ position: 0, color: 'FFFFFF' }, { position: 1, color: '000000' }],
    }
  }

  // 3. pattFill
  const pattNode = getChild(node, 'pattFill')
  if (pattNode) {
    const preset = attrs(pattNode).prst ?? 'pct50'
    const fgNode = getChild(pattNode, 'fgClr')
    const bgNode = getChild(pattNode, 'bgClr')
    const fgResolved = resolveDrawingColor(fgNode ? parseDrawingColor(fgNode) : undefined, theme)
    const bgResolved = resolveDrawingColor(bgNode ? parseDrawingColor(bgNode) : undefined, theme)
    return {
      kind: 'pattern',
      preset,
      fgColor: hexColor(fgResolved) ?? '000000',
      bgColor: hexColor(bgResolved) ?? 'FFFFFF',
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
