/** Decodes an investigator's capture (PNG or JPEG) to straight-alpha RGBA. */
import { createCanvas, loadImage } from '@napi-rs/canvas'
import type { Bitmap } from '../mark/types'

export async function decodeImage(bytes: Uint8Array): Promise<Bitmap> {
  const image = await loadImage(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength))
  const { width, height } = image
  if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width < 1 || height < 1) throw new Error('Invalid image dimensions')
  const ctx = createCanvas(width, height).getContext('2d')
  ctx.drawImage(image, 0, 0)
  return { width, height, data: new Uint8ClampedArray(ctx.getImageData(0, 0, width, height).data) }
}
