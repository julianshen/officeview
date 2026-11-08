/**
 * `CanvasBackend` (spec §7.3, §9.2): the one seam for every canvas the paint
 * pipeline constructs outside the caller's own context. A recorder-pool
 * worker uses exactly one canvas library per process, so the seam is
 * installed process-wide (`installCanvasBackend`) rather than threaded
 * through every painter call. Each construction site asks
 * `installedCanvasBackend()` first and otherwise runs its historical
 * fallback unchanged, so nothing changes until a backend is installed.
 * `createCanvas` is synchronous because three sites are (scratch surfaces,
 * pattern tiles, the local text measurer).
 */

/** The 2-D surface shape the painters rely on. */
export interface CanvasLike {
  width: number
  height: number
  getContext(type: '2d'): CanvasRenderingContext2D | null
}

/** A decoded raster image `drawImage` accepts. */
export type ImageLike = CanvasImageSource & { width: number; height: number }

export interface CanvasBackend {
  /** Human-readable id for diagnostics, e.g. '@napi-rs/canvas'. */
  readonly name: string
  createCanvas(width: number, height: number): CanvasLike
  /** Reconstitute a decoded image, preserving native image sampling semantics. */
  imageFromPixels(width: number, height: number, rgba: Uint8Array): Promise<ImageLike>
  /** Decode raster bytes (png/jpeg/gif/webp/bmp). */
  decodeImage(bytes: Uint8Array, mime?: string): Promise<ImageLike>
  /** Rasterise SVG bytes into an image `drawImage` accepts. */
  rasteriseSvg(bytes: Uint8Array): Promise<ImageLike>
}

let installed: CanvasBackend | undefined

/** Install the process-wide backend (recorder pool, scripts, tests).
 * Returns the previous one so callers can restore it. */
export function installCanvasBackend(backend: CanvasBackend | undefined): CanvasBackend | undefined {
  const previous = installed
  installed = backend
  return previous
}

/** The explicitly installed backend, if any. Clone sites (which copy the
 * constructor of an existing canvas so measurement uses the painting
 * engine) only defer to the seam when one is installed. */
export function installedCanvasBackend(): CanvasBackend | undefined {
  return installed
}

/** The installed backend, else `browserBackend`. */
export function canvasBackend(): CanvasBackend {
  return installed ?? browserBackend
}

/** Browser default for code that needs a backend with none installed (the
 * replayer): OffscreenCanvas, else a DOM canvas. Decoding uses
 * `createImageBitmap`. Headless callers must install a backend. */
export const browserBackend: CanvasBackend = {
  name: 'browser',
  createCanvas(width, height) {
    if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(width, height) as unknown as CanvasLike
    if (typeof document !== 'undefined') return Object.assign(document.createElement('canvas'), { width, height })
    throw new Error('no canvas available: install a CanvasBackend (installCanvasBackend)')
  },
  async imageFromPixels(width, height, rgba) {
    return await createImageBitmap(new ImageData(new Uint8ClampedArray(rgba), width, height)) as unknown as ImageLike
  },
  async decodeImage(bytes, mime) {
    return await createImageBitmap(new Blob([bytes as BlobPart], mime ? { type: mime } : undefined)) as unknown as ImageLike
  },
  async rasteriseSvg(bytes) {
    return await createImageBitmap(new Blob([bytes as BlobPart], { type: 'image/svg+xml' })) as unknown as ImageLike
  },
}
