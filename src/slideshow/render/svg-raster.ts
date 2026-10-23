/**
 * Bounded SVG rasterisation for the slideshow (PR #18 review 2, P1).
 *
 * An SVG's intrinsic size is whatever its author declared (20000 × 20000 is
 * one attribute away), so the show never lets a decoder pick the bitmap size.
 * The caller sizes the target from where the picture is drawn, leases
 * `w × h × 4` bytes, and only then calls `rasteriseSvg`, which draws the
 * vector image into a canvas of exactly that size. The decoded SVG image
 * stays vector (an `Image` element in browsers, an rsvg-backed image under
 * node-canvas) and is rasterised by the draw at the target size.
 */
import { scanSelfContainedSvg } from '../../core/svg'
import type { PptxShape } from '../../pptx/types'
import type { CanvasFactory } from '../types'

/** EMU per CSS px. */
const EMU_PER_PX = 9525

/**
 * The largest drawn size, in slide px, of image `index` among `nodes`: each use's box through its group
 * transforms, widened by a source crop (`srcRect`) so the visible part keeps its resolution. Undefined
 * when no node draws it.
 */
export function imageDrawSize(nodes: readonly PptxShape[], index: number): { w: number; h: number } | undefined {
  let w = 0, h = 0
  const walk = (list: readonly PptxShape[], sx: number, sy: number) => {
    for (const n of list) {
      if (n.group) {
        const g = n.group
        const kx = g.chExt.width > 0 ? g.ext.width / g.chExt.width : 1
        const ky = g.chExt.height > 0 ? g.ext.height / g.chExt.height : 1
        walk(n.children ?? [], sx * (Number.isFinite(kx) ? kx : 1), sy * (Number.isFinite(ky) ? ky : 1))
        continue
      }
      if (n.imageIndex !== index) continue
      const crop = (n as { image?: { srcRect?: { l: number; t: number; r: number; b: number } } }).image?.srcRect
      const fx = crop ? 1 - Math.max(0, crop.l) - Math.max(0, crop.r) : 1
      const fy = crop ? 1 - Math.max(0, crop.t) - Math.max(0, crop.b) : 1
      w = Math.max(w, (Math.abs(n.widthEmu) * sx) / EMU_PER_PX / (fx > 0.01 ? fx : 1))
      h = Math.max(h, (Math.abs(n.heightEmu) * sy) / EMU_PER_PX / (fy > 0.01 ? fy : 1))
    }
  }
  walk(nodes, 1, 1)
  return w > 0 && h > 0 ? { w, h } : undefined
}

type VectorImage = CanvasImageSource & { width: number; height: number }

async function loadSvg(bytes: Uint8Array): Promise<VectorImage> {
  if (typeof document !== 'undefined' && typeof createImageBitmap === 'function') {
    const copy = new Uint8Array(bytes.byteLength)
    copy.set(bytes)
    const url = URL.createObjectURL(new Blob([copy], { type: 'image/svg+xml' }))
    try {
      const image = new Image()
      image.src = url
      if (typeof image.decode === 'function') await image.decode()
      else await new Promise<void>((resolve, reject) => { image.onload = () => resolve(); image.onerror = () => reject(new Error('svg image load failed')) })
      return image as unknown as VectorImage
    } finally {
      URL.revokeObjectURL(url)
    }
  }
  const mod = (await import(/* @vite-ignore */ 'canvas')) as unknown as { loadImage: (src: Uint8Array | Buffer) => Promise<VectorImage> }
  return mod.loadImage(typeof Buffer !== 'undefined' ? Buffer.from(bytes) : bytes)
}

/**
 * Draws the SVG into a new `w` × `h` canvas (the size the caller has already leased). Undefined when the
 * bytes fail the self-contained preflight or cannot be decoded or drawn.
 */
export async function rasteriseSvg(bytes: Uint8Array, w: number, h: number, factory: CanvasFactory): Promise<HTMLCanvasElement | undefined> {
  if (!scanSelfContainedSvg(bytes).ok) return undefined
  let image: VectorImage
  try { image = await loadSvg(bytes) } catch { return undefined }
  const canvas = factory(w, h) as HTMLCanvasElement
  const ctx = canvas.getContext('2d')
  if (!ctx) return undefined
  try { ctx.drawImage(image, 0, 0, w, h) } catch { return undefined }
  return canvas
}
