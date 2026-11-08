/**
 * `CanvasBackend` (spec §7.3, §9.2): the one seam for every canvas the paint
 * pipeline constructs outside the caller's own context. A recorder-pool
 * worker uses exactly one canvas library per process, so the seam is
 * installed process-wide (`installCanvasBackend`) rather than threaded
 * through every painter call. With nothing installed, `defaultBackend`
 * reproduces the historical per-site fallbacks exactly
 * (OffscreenCanvas → DOM → the node `canvas` package).
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
  createCanvas(width: number, height: number): CanvasLike | Promise<CanvasLike>
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

/** The installed backend, else the historical default. */
export function canvasBackend(): CanvasBackend {
  return installed ?? defaultBackend
}

type NodeCanvas = {
  createCanvas(w: number, h: number): CanvasLike
  loadImage(src: Uint8Array): Promise<ImageLike>
}
// @vite-ignore: resolved only in node/bun (golden harness, scripts), never bundled
const nodeCanvas = async () => (await import(/* @vite-ignore */ 'canvas')) as unknown as NodeCanvas
const nodeBytes = (bytes: Uint8Array) => (typeof Buffer !== 'undefined' ? Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength) : bytes)

/** Today's behaviour, unchanged: browser surfaces first, the node `canvas`
 * package only where no DOM/OffscreenCanvas exists. */
export const defaultBackend: CanvasBackend = {
  name: 'default',
  async createCanvas(width, height) {
    if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(width, height) as unknown as CanvasLike
    if (typeof document !== 'undefined') return Object.assign(document.createElement('canvas'), { width, height })
    return (await nodeCanvas()).createCanvas(width, height)
  },
  async decodeImage(bytes, mime) {
    if (typeof createImageBitmap === 'function' && typeof Blob === 'function') {
      return await createImageBitmap(new Blob([bytes as BlobPart], mime ? { type: mime } : undefined)) as unknown as ImageLike
    }
    return (await nodeCanvas()).loadImage(nodeBytes(bytes))
  },
  async rasteriseSvg(bytes) {
    return (await nodeCanvas()).loadImage(nodeBytes(bytes))
  },
}
