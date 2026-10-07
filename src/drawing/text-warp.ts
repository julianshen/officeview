import type { LocalAffine, TextWarp, TextWarpPreset } from './text'
import { DEFAULT_WARP_ADJUSTMENTS } from './text'

export interface WarpBox {
  x: number
  y: number
  width: number
  height: number
}

export interface WarpTransform {
  x: number
  y: number
  rotation: number
  scaleX: number
  scaleY: number
  matrix?: LocalAffine
}

/**
 * Angle presets use adjustments measured in 60,000ths of a degree (0 to 360 degrees = 21,600,000).
 * Percentage/ratio presets use adjustments measured in 1/100,000ths (0 to 100,000).
 */
const ANGLE_PRESETS: ReadonlySet<TextWarpPreset> = new Set<TextWarpPreset>([
  'textArchUp',
  'textArchDown',
  'textCircle',
  'textButton',
  'textArchUpPour',
  'textArchDownPour',
  'textCirclePour',
  'textButtonPour',
])

export function clampAdjustment(preset: TextWarpPreset, name: string, value: number): number {
  if (ANGLE_PRESETS.has(preset) && name.startsWith('adj')) {
    return Math.max(0, Math.min(21600000, value))
  }
  // Standard percentage/ratio adjustment range [0, 100000]
  return Math.max(0, Math.min(100000, value))
}

export function resolveAdjustment(warp: TextWarp, name: string, fallback: number): number {
  const authored = warp.adjustments?.[name]
  const defaults = DEFAULT_WARP_ADJUSTMENTS[warp.preset]
  const raw = authored ?? defaults?.[name] ?? fallback
  return clampAdjustment(warp.preset, name, raw)
}

/**
 * Computes the local warp transformation for a point or glyph origin within a text bounding box.
 */
export function computeWarpTransform(
  warp: TextWarp,
  box: WarpBox,
  point: { x: number; y: number }
): WarpTransform {
  const w = Math.max(box.width, 1)
  const h = Math.max(box.height, 1)
  const u = Math.max(0, Math.min(1, (point.x - box.x) / w))
  const midY = box.y + h / 2
  const v = point.y - midY

  switch (warp.preset) {
    case 'textArchUp': {
      const adj = resolveAdjustment(warp, 'adj', 10800000)
      const theta = (adj / 60000) * (Math.PI / 180)
      if (theta < 1e-4) {
        return { x: point.x, y: point.y, rotation: 0, scaleX: 1, scaleY: 1 }
      }
      const alpha = (u - 0.5) * theta
      const r = w / theta
      const radialDist = r - v
      const xc = box.x + w / 2
      const yc = midY + r
      const x = xc + radialDist * Math.sin(alpha)
      const y = yc - radialDist * Math.cos(alpha)
      return { x, y, rotation: alpha, scaleX: 1, scaleY: 1 }
    }

    case 'textArchDown': {
      const adj = resolveAdjustment(warp, 'adj', 10800000)
      const theta = (adj / 60000) * (Math.PI / 180)
      if (theta < 1e-4) {
        return { x: point.x, y: point.y, rotation: 0, scaleX: 1, scaleY: 1 }
      }
      const alpha = (u - 0.5) * theta
      const r = w / theta
      const radialDist = r + v
      const xc = box.x + w / 2
      const yc = midY - r
      const x = xc + radialDist * Math.sin(alpha)
      const y = yc + radialDist * Math.cos(alpha)
      return { x, y, rotation: -alpha, scaleX: 1, scaleY: 1 }
    }

    case 'textCircle': {
      const adj = resolveAdjustment(warp, 'adj', 10800000)
      const theta = (adj / 60000) * (Math.PI / 180)
      const rBase = Math.min(w, h) / 2
      const r = Math.max(1, rBase - v)
      const xc = box.x + w / 2
      const yc = midY
      const phi = -Math.PI / 2 + u * theta
      const x = xc + r * Math.cos(phi)
      const y = yc + r * Math.sin(phi)
      const rotation = u * theta
      return { x, y, rotation, scaleX: 1, scaleY: 1 }
    }

    case 'textWave1': {
      const adj2 = resolveAdjustment(warp, 'adj2', 50000)
      const amp = (adj2 / 100000) * (h / 4)
      const dy = amp * Math.sin(2 * Math.PI * u)
      const slope = ((2 * Math.PI * amp) / w) * Math.cos(2 * Math.PI * u)
      const rotation = Math.atan(slope)
      return { x: point.x, y: point.y + dy, rotation, scaleX: 1, scaleY: 1 }
    }

    case 'textWave2': {
      const adj2 = resolveAdjustment(warp, 'adj2', 50000)
      const amp = (adj2 / 100000) * (h / 4)
      const dy = -amp * Math.sin(2 * Math.PI * u)
      const slope = ((-2 * Math.PI * amp) / w) * Math.cos(2 * Math.PI * u)
      const rotation = Math.atan(slope)
      return { x: point.x, y: point.y + dy, rotation, scaleX: 1, scaleY: 1 }
    }

    case 'textInflate': {
      const adj = resolveAdjustment(warp, 'adj', 50000)
      const k = adj / 100000
      const scaleY = 1 + k * Math.sin(Math.PI * u)
      const y = midY + v * scaleY
      return { x: point.x, y, rotation: 0, scaleX: 1, scaleY }
    }

    case 'textDeflate': {
      const adj = resolveAdjustment(warp, 'adj', 50000)
      const k = adj / 100000
      const scaleY = Math.max(0.1, 1 - k * 0.8 * Math.sin(Math.PI * u))
      const y = midY + v * scaleY
      return { x: point.x, y, rotation: 0, scaleX: 1, scaleY }
    }

    case 'textSlantUp': {
      const adj = resolveAdjustment(warp, 'adj', 25000)
      const s = adj / 100000
      const dy = -s * (u - 0.5) * h
      const rotation = Math.atan((-s * h) / w)
      return { x: point.x, y: point.y + dy, rotation, scaleX: 1, scaleY: 1 }
    }

    case 'textSlantDown': {
      const adj = resolveAdjustment(warp, 'adj', 25000)
      const s = adj / 100000
      const dy = s * (u - 0.5) * h
      const rotation = Math.atan((s * h) / w)
      return { x: point.x, y: point.y + dy, rotation, scaleX: 1, scaleY: 1 }
    }

    case 'textCurveUp': {
      const adj = resolveAdjustment(warp, 'adj', 25000)
      const k = adj / 100000
      const dy = -k * 4 * u * (1 - u) * (h / 2)
      const slope = ((-k * 4 * (1 - 2 * u) * (h / 2)) / w)
      return { x: point.x, y: point.y + dy, rotation: Math.atan(slope), scaleX: 1, scaleY: 1 }
    }

    case 'textCurveDown': {
      const adj = resolveAdjustment(warp, 'adj', 25000)
      const k = adj / 100000
      const dy = k * 4 * u * (1 - u) * (h / 2)
      const slope = ((k * 4 * (1 - 2 * u) * (h / 2)) / w)
      return { x: point.x, y: point.y + dy, rotation: Math.atan(slope), scaleX: 1, scaleY: 1 }
    }

    default: {
      return { x: point.x, y: point.y, rotation: 0, scaleX: 1, scaleY: 1 }
    }
  }
}

/**
 * Applies a warp transform to a Canvas 2D context.
 */
export function applyWarpTransform(ctx: CanvasRenderingContext2D, t: WarpTransform): void {
  ctx.translate(t.x, t.y)
  if (t.rotation !== 0) {
    ctx.rotate(t.rotation)
  }
  if (t.scaleX !== 1 || t.scaleY !== 1) {
    ctx.scale(t.scaleX, t.scaleY)
  }
}
