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
export async function decodeImageAsset(asset: ImageCandidates, decode?: ImageDecodeFn): Promise<CanvasImageSource | undefined> {
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
