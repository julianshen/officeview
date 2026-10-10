/** DrawingML pattern presets and Canvas 2D tile rasterizer. */

export const SUPPORTED_PATTERN_PRESETS: ReadonlySet<string> = new Set([
  'dkUpDiag', 'dkDnDiag', 'ltUpDiag', 'ltDnDiag', 'smGrid', 'lgGrid',
  'pct5', 'pct10', 'pct20', 'pct25', 'pct30', 'pct40', 'pct50',
  'pct60', 'pct70', 'pct75', 'pct80', 'pct90',
  'horz', 'vert', 'grid', 'check', 'trellis',
])

export type PatternPreset =
  | 'dkUpDiag' | 'dkDnDiag' | 'ltUpDiag' | 'ltDnDiag' | 'smGrid' | 'lgGrid'
  | 'pct5' | 'pct10' | 'pct20' | 'pct25' | 'pct30' | 'pct40' | 'pct50'
  | 'pct60' | 'pct70' | 'pct75' | 'pct80' | 'pct90'
  | 'horz' | 'vert' | 'grid' | 'check' | 'trellis'

const BAYER_8X8: readonly number[] = [
   0, 32,  8, 40,  2, 34, 10, 42,
  48, 16, 56, 24, 50, 18, 58, 26,
  12, 44,  4, 36, 14, 46,  6, 38,
  60, 28, 52, 20, 62, 30, 54, 22,
   3, 35, 11, 43,  1, 33,  9, 41,
  51, 19, 59, 27, 49, 17, 57, 25,
  15, 47,  7, 39, 13, 45,  5, 37,
  63, 31, 55, 23, 61, 29, 53, 21,
]

function generateBayerMask(fgCount: number): readonly number[] {
  const mask: number[] = [0, 0, 0, 0, 0, 0, 0, 0]
  for (let y = 0; y < 8; y++) {
    let row = 0
    for (let x = 0; x < 8; x++) {
      if (BAYER_8X8[y * 8 + x] < fgCount) {
        row |= (1 << (7 - x))
      }
    }
    mask[y] = row
  }
  return mask
}

const STATIC_PATTERN_MASKS: Readonly<Record<string, readonly number[]>> = {
  // Percentage presets generated from 8x8 Bayer dither thresholds
  pct5: generateBayerMask(3),
  pct10: generateBayerMask(6),
  pct20: generateBayerMask(13),
  pct25: generateBayerMask(16),
  pct30: generateBayerMask(19),
  pct40: generateBayerMask(26),
  pct50: generateBayerMask(32),
  pct60: generateBayerMask(38),
  pct70: generateBayerMask(45),
  pct75: generateBayerMask(48),
  pct80: generateBayerMask(51),
  pct90: generateBayerMask(58),
  // horz: horizontal rules repeating every 4 pixels (rows 0, 4)
  horz: [0xFF, 0x00, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00],
  // vert: vertical rules repeating every 4 pixels (cols 0, 4)
  vert: [0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88],
  // grid: orthogonal grid repeating every 4 pixels (rows 0, 4 + cols 0, 4)
  grid: [0xFF, 0x88, 0x88, 0x88, 0xFF, 0x88, 0x88, 0x88],
  // check: 4x4 alternating checkerboard
  check: [0xF0, 0xF0, 0xF0, 0xF0, 0x0F, 0x0F, 0x0F, 0x0F],
  // trellis: standard lattice mesh
  trellis: [0xFF, 0x11, 0x55, 0x11, 0xFF, 0x11, 0x55, 0x11],
}
// Masks are shared module state handed out for inspection; freeze them so a
// consumer can never mutate (and poison) tile generation for everyone else.
for (const mask of Object.values(STATIC_PATTERN_MASKS)) Object.freeze(mask);
Object.freeze(STATIC_PATTERN_MASKS);

/**
 * Returns the 8x8 bitmask (8 row bytes) for the given preset if defined.
 */
export function getPatternBitmask(preset: string): readonly number[] | undefined {
  return STATIC_PATTERN_MASKS[preset]
}

/**
 * 2-color pattern tile renderer for supported diagonal, grid, and bitmask families.
 */
export function paintPatternTile(ctx: CanvasRenderingContext2D, preset: string, fg: string, bg: string, size: number): void {
  ctx.save()
  try {
    ctx.fillStyle = bg
    ctx.fillRect(0, 0, size, size)
    const mask = getPatternBitmask(preset)
    if (mask) {
      ctx.fillStyle = fg
      const scale = size / 8
      ctx.beginPath()
      for (let y = 0; y < 8; y++) {
        const row = mask[y]
        if (row === 0) continue
        for (let x = 0; x < 8; x++) {
          if ((row & (1 << (7 - x))) !== 0) {
            ctx.rect(x * scale, y * scale, scale, scale)
          }
        }
      }
      ctx.fill()
    } else if (preset.endsWith('UpDiag') || preset.endsWith('DnDiag')) {
      const up = preset.endsWith('UpDiag')
      ctx.strokeStyle = fg
      ctx.lineWidth = preset.startsWith('dk') ? Math.max(2, size / 3) : Math.max(1, size / 8)
      ctx.beginPath()
      for (const o of [-size, 0, size]) {
        if (up) { ctx.moveTo(o, size); ctx.lineTo(o + size, 0) }
        else { ctx.moveTo(o, 0); ctx.lineTo(o + size, size) }
      }
      ctx.stroke()
    } else if (preset === 'smGrid' || preset === 'lgGrid') {
      ctx.strokeStyle = fg
      const cell = preset === 'smGrid' ? size / 2 : size
      ctx.lineWidth = Math.max(1, size / 8)
      ctx.beginPath()
      for (let k = 0; k <= size + 0.5; k += cell) {
        ctx.moveTo(k, 0); ctx.lineTo(k, size)
        ctx.moveTo(0, k); ctx.lineTo(size, k)
      }
      ctx.stroke()
    }
    // Unknown presets paint background only: every production caller gates on
    // SUPPORTED_PATTERN_PRESETS, so reaching here means a future caller
    // bypassed the gate — a blank tile degrades visibly instead of a
    // misleading grid.
  } finally {
    ctx.restore()
  }
}

interface TileSurface {
  width: number
  height: number
  getContext(kind: '2d'): unknown
}

function createTileSurface(ctx: CanvasRenderingContext2D, size: number): TileSurface | undefined {
  if (typeof OffscreenCanvas !== 'undefined') {
    try {
      return new OffscreenCanvas(size, size)
    } catch {
      return undefined
    }
  }
  const canvas = ctx.canvas
  if (!canvas) return undefined
  const owner = canvas.ownerDocument
  if (typeof owner?.createElement === 'function') {
    const s = owner.createElement('canvas')
    s.width = size
    s.height = size
    return s
  }
  try {
    const Constructor = canvas.constructor as unknown as new (width: number, height: number) => TileSurface
    const surface = new Constructor(size, size)
    return typeof surface.getContext === 'function' ? surface : undefined
  } catch {
    return undefined
  }
}

export const MAX_PATTERN_TILES = 64
const patternTiles = new Map<string, CanvasImageSource>()

/** Returns cached or newly rendered 8x8 pattern tile image. True LRU with hit promotion. */
export function getPatternTile(ctx: CanvasRenderingContext2D, preset: string, fg: string, bg: string, size = 8): CanvasImageSource | undefined {
  // Size joins the key: tiles are rasterized per size, so differently sized
  // requests for the same preset/colors must not share one bitmap.
  const key = `${preset}\n${fg}\n${bg}\n${size}`
  const existing = patternTiles.get(key)
  if (existing) {
    patternTiles.delete(key)
    patternTiles.set(key, existing)
    return existing
  }
  const surface = createTileSurface(ctx, size)
  if (!surface) return undefined
  const sCtx = (surface as { getContext(kind: '2d'): unknown }).getContext('2d') as CanvasRenderingContext2D | undefined
  if (!sCtx) return undefined
  paintPatternTile(sCtx, preset, fg, bg, size)
  if (patternTiles.size >= MAX_PATTERN_TILES) {
    const oldest = patternTiles.keys().next()
    if (!oldest.done) patternTiles.delete(oldest.value)
  }
  const img = surface as unknown as CanvasImageSource
  patternTiles.set(key, img)
  return img
}

/** Evict a cache entry if creating a pattern from it fails (avoids poisoning). */
export function evictPatternTile(preset: string, fg: string, bg: string, size = 8): void {
  const key = `${preset}\n${fg}\n${bg}\n${size}`
  patternTiles.delete(key)
}

/** Test / lifecycle helper to reset the LRU cache. */
export function clearPatternTileCache(): void {
  patternTiles.clear()
}

/** Test inspection helper for cache size. */
export function patternTileCacheSize(): number {
  return patternTiles.size
}

/** Test inspection helper for key presence. */
export function hasPatternTile(preset: string, fg: string, bg: string, size = 8): boolean {
  const key = `${preset}\n${fg}\n${bg}\n${size}`
  return patternTiles.has(key)
}
