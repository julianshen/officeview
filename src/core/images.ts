/**
 * Image decoding for embedded office media (png/jpeg/gif/webp/svg).
 * Browser: createImageBitmap. Node/bun: canvas package's loadImage.
 *
 * Every SVG byte stream is preflighted (`guardSvgBytes`) before it reaches ANY
 * decoder — including the low-level public `decodeImage`/`decodeSvgImage`
 * entry points, a bare `a:blip` that references an SVG part directly, and a
 * hostile primary fallback behind a preferred (but failing) safe svgBlip.
 */
import { guardSvgBytes, looksLikeSvg, scanSelfContainedSvg, type ImageSelection, type SvgCandidate, type SvgVerdict } from './svg'

export type DecodedImage = CanvasImageSource & { width: number; height: number }

export type ImageDecodeFn = (bytes: Uint8Array, mime?: string) => Promise<CanvasImageSource | undefined>

/** Complete candidate set for one picture identity (raster primary + optional SVG). */
export interface ImageCandidates {
  /** Primary a:blip bytes: raster fallback, or the SVG itself for a bare SVG part. */
  data?: Uint8Array
  mime?: string
  /** Owner-part media path hint, used for reliable SVG detection. */
  pathHint?: string
  /** Owner-part media path (structural images); used when pathHint is absent. */
  partPath?: string
  /** Retained svgBlip candidate (accepted OR rejected). */
  svg?: SvgCandidate
  /** Preflight verdict when the primary bytes are themselves SVG. */
  primarySvgVerdict?: SvgVerdict
  /** Explicitly false when the picture carries an SVG candidate but no raster. */
  hasRaster?: boolean
  /** Shared parse/decode selection record (also referenced by drawing coverage). */
  imageSelection?: ImageSelection
  /** Recolour the decoded pixels (`a:duotone`). */
  duotone?: Duotone
  /** Paint a raster that is exactly 2 × 2 pixels as one colour, as PowerPoint does (see flatTwoByTwoPixels). */
  flatTwoByTwo?: boolean
}

/** `a:duotone`'s two colours, resolved: luminance 0 maps to `dark`, luminance 1 to `light`. Channels 0–255, alpha 0–1. */
export interface Duotone {
  dark: { r: number; g: number; b: number; a: number }
  light: { r: number; g: number; b: number; a: number }
}

/**
 * Recolour RGBA pixels in place (ECMA-376 20.1.8.23 duotone): each pixel's luminance (Rec. 601
 * weights, as LibreOffice's duotone filter) is the position on the line from `dark` to `light`.
 * The pixel keeps its alpha, scaled by the interpolated colour's alpha.
 */
export function duotonePixels(data: Uint8ClampedArray, dark: Duotone['dark'], light: Duotone['light']): void {
  for (let i = 0; i < data.length; i += 4) {
    const t = (0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]) / 255
    data[i] = dark.r + (light.r - dark.r) * t
    data[i + 1] = dark.g + (light.g - dark.g) * t
    data[i + 2] = dark.b + (light.b - dark.b) * t
    data[i + 3] = data[i + 3] * (dark.a + (light.a - dark.a) * t)
  }
}

type Surface = { width: number; height: number; getContext(type: '2d'): unknown }
/** A canvas from the same backend `decodeImage` decodes with. */
async function surface(width: number, height: number): Promise<Surface> {
  if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(width, height) as unknown as Surface
  if (typeof document !== 'undefined' && typeof createImageBitmap === 'function') return Object.assign(document.createElement('canvas'), { width, height })
  const mod = (await import(/* @vite-ignore */ 'canvas')) as unknown as { createCanvas: (w: number, h: number) => Surface }
  return mod.createCanvas(width, height)
}

/**
 * Reserves memory for a decode step's temporary buffers: a held reservation, or undefined to refuse.
 * The recolour step asks for its canvas and pixel buffer (2 × width × height × 4 bytes) before allocating them.
 */
export type ImageScratch = (bytes: number) => { release(): void } | undefined

/**
 * The decoded image recoloured on a canvas of its own size; the source is released. Undefined (source
 * released) when `scratch` refuses the buffers or no 2d context is available.
 */
async function applyDuotone(image: CanvasImageSource, duotone: Duotone, scratch?: ImageScratch): Promise<CanvasImageSource | undefined> {
  const { width, height } = image as unknown as { width: number; height: number }
  if (!(width > 0 && height > 0)) return image
  const held = scratch?.(2 * width * height * 4)
  if (scratch && !held) { releaseDecodedImage(image); return undefined }
  try {
    const canvas = await surface(width, height)
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null
    if (!ctx) return undefined
    ctx.drawImage(image, 0, 0, width, height)
    const pixels = ctx.getImageData(0, 0, width, height)
    duotonePixels(pixels.data, duotone.dark, duotone.light)
    ctx.putImageData(pixels, 0, 0)
    return canvas as unknown as CanvasImageSource
  } finally {
    releaseDecodedImage(image)
    held?.release()
  }
}

/**
 * The one colour PowerPoint paints for a stretched 2 × 2 picture, written over all four RGBA pixels in
 * place (tdf146223; probe decks exported as movies from PowerPoint 16.114). A PNG gets the plain mean.
 * A JPEG gets the bilinear sample 13/32 of a pixel right of and below the top-left pixel's centre:
 * weights (19/32)² top-left, 19 × 13/32² top-right and bottom-left, (13/32)² bottom-right; every
 * one-white-pixel probe matched within two levels. Other formats are unmeasured and take the mean.
 * Colours are weighted by alpha, so transparent pixels add no colour.
 */
export function flatTwoByTwoPixels(data: Uint8ClampedArray, jpeg: boolean): void {
  const weights = jpeg ? [361, 247, 247, 169].map(w => w / 1024) : [0.25, 0.25, 0.25, 0.25]
  let r = 0, g = 0, b = 0, a = 0
  weights.forEach((w, i) => {
    const wa = w * data[i * 4 + 3]
    r += wa * data[i * 4]; g += wa * data[i * 4 + 1]; b += wa * data[i * 4 + 2]; a += wa
  })
  const flat = a > 0 ? [r / a, g / a, b / a, a] : [0, 0, 0, 0]
  for (let i = 0; i < 4; i++) data.set(flat, i * 4)
}

/** A 2 × 2 decoded image flattened on a canvas of its own size (source released); any other size is returned as is. */
async function flattenTwoByTwo(image: CanvasImageSource, jpeg: boolean): Promise<CanvasImageSource> {
  const { width, height } = image as unknown as { width: number; height: number }
  if (width !== 2 || height !== 2) return image
  // 32 bytes of canvas and pixels: too small to charge to a workspace.
  const canvas = await surface(2, 2)
  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null
  if (!ctx) return image
  ctx.drawImage(image, 0, 0, 2, 2)
  const pixels = ctx.getImageData(0, 0, 2, 2)
  flatTwoByTwoPixels(pixels.data, jpeg)
  ctx.putImageData(pixels, 0, 0)
  releaseDecodedImage(image)
  return canvas as unknown as CanvasImageSource
}

function svgRejection(bytes: Uint8Array, mime?: string, pathHint?: string): string | undefined {
  const verdict = guardSvgBytes(bytes, mime, pathHint)
  return verdict.ok ? undefined : verdict.reason
}

export async function decodeImage(bytes: Uint8Array, mimeHint?: string): Promise<DecodedImage> {
  const rejected = svgRejection(bytes, mimeHint)
  if (rejected) throw new Error(`svg-rejected:${rejected}`)
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

/**
 * Browser-safe SVG decode. `createImageBitmap` acceptance of SVG is not
 * guaranteed, so a blob-URL `Image` element is the explicit fallback. The
 * object URL is always revoked. SVG bytes are preflighted first.
 */
export async function decodeSvgImage(bytes: Uint8Array, mimeHint?: string): Promise<DecodedImage> {
  const verdict = scanSelfContainedSvg(bytes)
  if (!verdict.ok) throw new Error(`svg-rejected:${verdict.reason}`)
  if (typeof document !== 'undefined' && typeof createImageBitmap === 'function') {
    try {
      return await decodeImage(bytes, mimeHint ?? 'image/svg+xml')
    } catch {
      return await decodeSvgViaImageElement(bytes)
    }
  }
  return decodeImage(bytes, mimeHint ?? 'image/svg+xml')
}

async function decodeSvgViaImageElement(bytes: Uint8Array): Promise<DecodedImage> {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  const blob = new Blob([copy], { type: 'image/svg+xml' })
  const url = URL.createObjectURL(blob)
  try {
    const image = new Image()
    image.src = url
    if (typeof image.decode === 'function') await image.decode()
    else await new Promise<void>((resolve, reject) => { image.onload = () => resolve(); image.onerror = () => reject(new Error('svg image load failed')) })
    return image as unknown as DecodedImage
  } finally {
    URL.revokeObjectURL(url)
  }
}

async function tryDecode(decode: ImageDecodeFn, bytes: Uint8Array, mime?: string): Promise<CanvasImageSource | undefined> {
  try { return (await decode(bytes, mime)) ?? undefined } catch { return undefined }
}

/** Release a decoded bitmap when the owning paint lease is disposed. */
export function releaseDecodedImage(image: CanvasImageSource | undefined): void {
  const close = (image as { close?: () => void } | undefined)?.close
  if (typeof close === 'function') { try { close.call(image) } catch { /* already released */ } }
}

/**
 * Select and decode one embedded representation.
 *
 * Preference derives from the RETAINED candidates (never from a previously
 * mutated selection), so a repeat render retries SVG after an earlier decoder
 * failure. Rejected candidates are retained so distinct pairs never collapse
 * and an ordinary raster never inherits another source's rejection reason. A
 * custom decoder returning `undefined` is a decode failure.
 */
export async function decodeImageAsset(asset: ImageCandidates, decode?: ImageDecodeFn, scratch?: ImageScratch): Promise<CanvasImageSource | undefined> {
  const image = await decodeRepresentation(asset, decode)
  const toned = image && asset.duotone ? await applyDuotone(image, asset.duotone, scratch) : image
  // Rasters only: a usable SVG candidate is preferred, and a vector has no pixels to flatten.
  const raster = asset.data && !asset.svg?.verdict.ok
  return toned && asset.flatTwoByTwo && raster ? flattenTwoByTwo(toned, sniffImageMime(asset.data!) === 'image/jpeg') : toned
}

async function decodeRepresentation(asset: ImageCandidates, decode?: ImageDecodeFn): Promise<CanvasImageSource | undefined> {
  // Parse-time reason is IMMUTABLE: recomputed each render from retained facts,
  // so a stale decode failure can never survive a later success.
  const parseReason = asset.svg && !asset.svg.verdict.ok ? asset.svg.verdict.reason
    : asset.primarySvgVerdict && !asset.primarySvgVerdict.ok ? asset.primarySvgVerdict.reason
    : undefined
  const record = (representation: ImageSelection['representation'], reason: string | undefined): void => {
    if (!asset.imageSelection) return
    asset.imageSelection.phase = 'decoded'
    asset.imageSelection.representation = representation
    asset.imageSelection.reason = reason
  }
  // Primary fallback bytes. When a svgBlip candidate exists, `data` is the
  // raster fallback; otherwise a primary that looks like SVG is the SVG itself.
  const primaryLooksSvg = !asset.svg && looksLikeSvg(asset.data, asset.mime, asset.pathHint ?? asset.partPath)
  const rasterBytes = asset.svg
    ? (asset.hasRaster === false ? undefined : asset.data)
    : (primaryLooksSvg ? undefined : asset.data)
  // Guard the fallback byte stream too: a hostile primary SVG must never reach a decoder.
  const rasterIsSvg = rasterBytes ? looksLikeSvg(rasterBytes, asset.mime, asset.pathHint ?? asset.partPath) : false
  const rasterVerdict: SvgVerdict | undefined = rasterIsSvg ? scanSelfContainedSvg(rasterBytes) : undefined
  const rasterUsable = !!rasterBytes && (!rasterIsSvg || rasterVerdict?.ok === true)
  const svgBytes = asset.svg?.bytes ?? (primaryLooksSvg ? asset.data : undefined)
  const svgVerdict: SvgVerdict | undefined = asset.svg ? asset.svg.verdict : (asset.primarySvgVerdict ?? (svgBytes ? scanSelfContainedSvg(svgBytes) : undefined))
  const decodeRaster = async (reason: string | undefined): Promise<CanvasImageSource | undefined> => {
    if (!rasterUsable) return undefined
    const image = await tryDecode(decode ?? decodeImage, rasterBytes!, rasterIsSvg ? 'image/svg+xml' : asset.mime)
    if (image) { record(rasterIsSvg ? 'svg' : 'raster', reason); return image }
    return undefined
  }

  if (svgBytes) {
    if (svgVerdict?.ok) {
      const image = await tryDecode(decode ?? decodeSvgImage, svgBytes, 'image/svg+xml')
      if (image) { record('svg', undefined); return image }
      const raster = await decodeRaster('svg-decode-failed')
      if (raster) return raster
      record('none', rasterVerdict && !rasterVerdict.ok ? rasterVerdict.reason : 'svg-decode-failed')
      return undefined
    }
    // Rejected SVG candidate (or a byte-less missing/external/link-only state):
    // fall back to the guarded raster once, preserving the source reason.
    const reason = svgVerdict && !svgVerdict.ok ? svgVerdict.reason : 'svg-rejected'
    const raster = await decodeRaster(reason)
    if (raster) return raster
    record('none', reason)
    return undefined
  }

  if (rasterBytes) {
    if (rasterIsSvg && rasterVerdict?.ok !== true) { record('none', rasterVerdict && !rasterVerdict.ok ? rasterVerdict.reason : 'svg-rejected'); return undefined }
    const image = await tryDecode(decode ?? (rasterIsSvg ? decodeSvgImage : decodeImage), rasterBytes, rasterIsSvg ? 'image/svg+xml' : asset.mime)
    if (image) { record(rasterIsSvg ? 'svg' : 'raster', parseReason); return image }
    record('none', 'raster-decode-failed')
    return undefined
  }
  record('none', 'no-image')
  return undefined
}

/** Sniff image mime from magic bytes; SVG (text/gzip) is detected last. */
export function sniffImageMime(bytes: Uint8Array): string | undefined {
  if (bytes[0] === 0x89 && bytes[1] === 0x50) return 'image/png'
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return 'image/jpeg'
  if (bytes[0] === 0x47 && bytes[1] === 0x49) return 'image/gif'
  if (bytes[8] === 0x57 && bytes[9] === 0x45) return 'image/webp'
  if (looksLikeSvg(bytes)) return 'image/svg+xml'
  return undefined
}
