/**
 * The `@napi-rs/canvas` CanvasBackend (spec §9.2, Stage-0 T1 choice) for
 * Stage-1 scripts and the recorder pool.
 */
import { createCanvas, loadImage } from '@napi-rs/canvas'
import type { CanvasBackend, CanvasLike, ImageLike } from '../../src/core/canvas-backend'

const buffer = (bytes: Uint8Array) => Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)

export const napiBackend: CanvasBackend = {
  name: '@napi-rs/canvas',
  createCanvas: (w, h) => createCanvas(w, h) as unknown as CanvasLike,
  async imageFromPixels(w, h, rgba) {
    const canvas = createCanvas(w, h), ctx = canvas.getContext('2d')
    const pixels = ctx.createImageData(w, h); pixels.data.set(rgba); ctx.putImageData(pixels, 0, 0)
    return await loadImage(canvas.toBuffer('image/png')) as unknown as ImageLike
  },
  decodeImage: async (bytes) => await loadImage(buffer(bytes)) as unknown as ImageLike,
  rasteriseSvg: async (bytes) => await loadImage(buffer(bytes)) as unknown as ImageLike,
}
