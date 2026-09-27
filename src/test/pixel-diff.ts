/**
 * Pixel comparison utilities for the golden-image harness.
 * Node-oriented (uses the `canvas` package for PNG IO) but the compare logic
 * is pure and works on any RGBA buffers.
 */

export interface Bitmap {
  width: number
  height: number
  /** RGBA, 4 bytes per pixel. */
  data: Uint8ClampedArray
}

export interface DiffResult {
  width: number
  height: number
  /** Count of pixels where any channel delta > threshold. */
  changedPixels: number
  /** Total pixels compared. */
  totalPixels: number
  /** changedPixels / totalPixels, 0..1. */
  ratio: number
  /** Largest per-channel delta seen. */
  maxDelta: number
  /** RGBA mask: red where changed, transparent elsewhere (4B/px). */
  mask: Uint8ClampedArray
}

/**
 * Compare two bitmaps. A pixel counts as changed when any channel differs by
 * more than `threshold` (0-255). Default threshold 8 forgives minor
 * antialiasing jitter from font rasterizers.
 */
export function diffBitmaps(a: Bitmap, b: Bitmap, threshold = 8): DiffResult {
  if (a.width !== b.width || a.height !== b.height) {
    throw new Error(`dimension mismatch: ${a.width}x${a.height} vs ${b.width}x${b.height}`)
  }
  const n = a.width * a.height
  const mask = new Uint8ClampedArray(n * 4)
  let changed = 0
  let maxDelta = 0
  for (let i = 0; i < n; i++) {
    const o = i * 4
    let d = 0
    for (let ch = 0; ch < 3; ch++) {
      const delta = Math.abs(a.data[o + ch] - b.data[o + ch])
      if (delta > d) d = delta
    }
    if (d > maxDelta) maxDelta = d
    if (d > threshold) {
      changed++
      mask[o] = 255
      mask[o + 3] = 255
    } else {
      mask[o + 3] = a.data[o + 3] // faint ghost of the original alpha
    }
  }
  return { width: a.width, height: a.height, changedPixels: changed, totalPixels: n, ratio: changed / n, maxDelta, mask }
}

/** --- PNG IO (node-canvas) --- */

function createImageData(bitmap: Bitmap): import('canvas').ImageData {
  const { createImageData } = (require('canvas') as { createImageData: (arr: Uint8ClampedArray, w: number, h: number) => import('canvas').ImageData })
  return createImageData(bitmap.data, bitmap.width, bitmap.height)
}

export async function bitmapFromCanvas(canvas: import('canvas').Canvas): Promise<Bitmap> {
  const ctx = canvas.getContext('2d')
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height)
  return { width: canvas.width, height: canvas.height, data: img.data }
}

export async function savePng(bitmap: Bitmap, path: string): Promise<void> {
  const { createCanvas } = await import('canvas')
  const canvas = createCanvas(bitmap.width, bitmap.height)
  const ctx = canvas.getContext('2d')
  ctx.putImageData(createImageData(bitmap), 0, 0)
  const { writeFileSync } = await import('fs')
  writeFileSync(path, canvas.toBuffer('image/png'))
}

export async function loadPng(path: string): Promise<Bitmap> {
  const { loadImage, createCanvas } = await import('canvas')
  const img = await loadImage(path)
  const canvas = createCanvas(img.width, img.height)
  const ctx = canvas.getContext('2d')
  ctx.drawImage(img, 0, 0)
  return bitmapFromCanvas(canvas)
}

/** Render paintables (all units) stacked? No — returns one bitmap per unit. */
export async function renderPaintables(
  paintables: Array<{ spec: { widthPx: number; heightPx: number }; paint: (ctx: CanvasRenderingContext2D) => void }>,
): Promise<Bitmap[]> {
  const { createCanvas } = await import('canvas')
  return paintables.map(({ spec, paint }) => {
    const canvas = createCanvas(Math.ceil(spec.widthPx), Math.ceil(spec.heightPx))
    const ctx = canvas.getContext('2d')!
    paint(ctx as never)
    return { width: canvas.width, height: canvas.height, data: ctx.getImageData(0, 0, canvas.width, canvas.height).data }
  })
}
