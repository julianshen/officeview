/**
 * Image decoding for embedded office media (png/jpeg/gif/webp).
 * Browser: createImageBitmap. Node/bun: canvas package's loadImage.
 */
export type DecodedImage = CanvasImageSource & { width: number; height: number }

export async function decodeImage(bytes: Uint8Array, mimeHint?: string): Promise<DecodedImage> {
  if (typeof document !== 'undefined' && typeof createImageBitmap === 'function') {
    const copy = new Uint8Array(bytes.byteLength)
    copy.set(bytes)
    const blob = new Blob([copy], { type: mimeHint ?? 'image/png' })
    const bitmap = await createImageBitmap(blob)
    return bitmap as unknown as DecodedImage
  }
  const mod = (await import(/* @vite-ignore */ 'canvas')) as unknown as {
    loadImage: (src: Uint8Array | Buffer) => Promise<DecodedImage>
  }
  // node-canvas wants a Buffer
  return mod.loadImage(typeof Buffer !== 'undefined' ? Buffer.from(bytes) : bytes)
}

/** Sniff image mime from magic bytes. */
export function sniffImageMime(bytes: Uint8Array): string | undefined {
  if (bytes[0] === 0x89 && bytes[1] === 0x50) return 'image/png'
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return 'image/jpeg'
  if (bytes[0] === 0x47 && bytes[1] === 0x49) return 'image/gif'
  if (bytes[8] === 0x57 && bytes[9] === 0x45) return 'image/webp'
  return undefined
}
