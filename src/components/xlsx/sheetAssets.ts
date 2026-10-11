/**
 * Font/image lease for the interactive sheet viewport (plan S5.2; spec §12.4).
 *
 * Unlike the static paintables, which decode every workbook image up front, the
 * sheet viewport leases ONE sheet's images. Every picture is admitted before any
 * decode from dimensions it can prove: a raster header (S1.1b
 * `probeImageHeader`/`imageLimitsFor`) or an SVG root's intrinsic size. An SVG
 * whose rasterized size is unknown or too large never reaches the decoder: its
 * raster fallback is used when that is admissible, otherwise it is refused.
 *
 * Decoded pictures live in a SheetImageCache shared by every lease of one viewer:
 * a decode reserves twice its RGBA size (surface plus decoder transient) and
 * keeps one RGBA size while retained, all against one aggregate limit, so leases
 * that overlap during a font or sheet swap never exceed it, and a picture two
 * leases need is decoded once. Refused pictures paint as the renderer's blank
 * placeholder and are counted.
 */
import type { XlsxDocument, XlsxDrawing } from '../../xlsx/types'
import type { ProfileId } from '../../xlsx/view/contracts'
import type { RenderSheetAssets } from '../../xlsx/render'
import { acquireFonts } from '../../core/fonts/register'
import { withFallbackFonts } from '../../core/fonts/fallback'
import type { FontDiagnostic } from '../../core/fonts/types'
import { releaseDecodedImage, type ImageCandidates, type ImageDecodeFn } from '../../core/images'
import { looksLikeSvg } from '../../core/svg'
import { collectCachedPictures, decodeImageAssets, type PaintOptions } from '../../render/paint'
import { IMAGE_PROBE_BYTES, imageLimitsFor, probeImageHeader, type ImageLimits } from '../../xlsx/storage/imageAdmission'

/** The aggregate image budget of one viewer: two pictures at the per-image cap. */
export function imageBudgetFor(profile: ProfileId): number { return imageLimitsFor(profile).maxDecodedBytes * 2 }

interface CacheEntry { bytes: number; refs: number; promise: Promise<CanvasImageSource | undefined>; settled: boolean; image?: CanvasImageSource }
/** Refcounted decoded pictures under one aggregate byte limit. */
export class SheetImageCache {
  chargedBytes = 0
  peakBytes = 0
  private readonly entries = new Map<object, CacheEntry>()
  constructor(readonly limitBytes: number) {}

  /** A retained or in-flight decode of `asset`, or null when the reservation
   * (2 x `bytes` while decoding) does not fit what the other leases hold. */
  acquire(key: object, asset: ImageCandidates, bytes: number, decode?: ImageDecodeFn): Promise<CanvasImageSource | undefined> | null {
    const found = this.entries.get(key)
    if (found) { found.refs++; return found.promise }
    if (this.chargedBytes + bytes * 2 > this.limitBytes) return null
    this.charge(bytes * 2)
    const entry: CacheEntry = { bytes, refs: 1, settled: false, promise: Promise.resolve(undefined) }
    entry.promise = decodeImageAssets([asset], decode).then(([image]) => image, () => undefined).then(image => {
      entry.settled = true
      this.charge(-bytes) // the decoder transient is gone; the surface stays
      if (entry.refs === 0) { this.drop(key, entry, image); return undefined }
      entry.image = image
      return image
    })
    this.entries.set(key, entry)
    return entry.promise
  }

  release(key: object): void {
    const entry = this.entries.get(key)
    if (!entry || --entry.refs > 0) return
    // An in-flight decode is dropped when it settles.
    if (entry.settled) this.drop(key, entry, entry.image)
  }

  private drop(asset: object, entry: CacheEntry, image: CanvasImageSource | undefined): void {
    if (this.entries.get(asset) !== entry) return
    this.entries.delete(asset)
    this.charge(-entry.bytes)
    releaseDecodedImage(image)
  }

  private charge(delta: number): void {
    this.chargedBytes += delta
    this.peakBytes = Math.max(this.peakBytes, this.chargedBytes)
  }
}

export interface SheetAssetsOptions extends PaintOptions {
  /** Sheet whose pictures are leased (default 0). */
  sheetIndex?: number
  profile?: ProfileId
  /** Shared by every lease of one viewer; a private cache is created otherwise. */
  imageCache?: SheetImageCache
}
export interface SheetAssetsLease {
  assets: RenderSheetAssets
  fontDiagnostics: FontDiagnostic[]
  /** RGBA bytes of this lease's admitted pictures, and the aggregate limit. */
  decodedBytes: number
  imageBudgetBytes: number
  /** Pictures of this sheet refused by admission or the aggregate limit. */
  refusedImages: number
  dispose: () => void
}

function imageIndexes(nodes: readonly XlsxDrawing[], out = new Set<number>()): Set<number> {
  for (const node of nodes) {
    if (node.imageIndex !== undefined) out.add(node.imageIndex)
    if (node.children) imageIndexes(node.children as XlsxDrawing[], out)
  }
  return out
}

const UNIT_PX: Record<string, number> = { '': 1, px: 1, pt: 4 / 3, pc: 16, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 }
/** Intrinsic raster size of an SVG document (CSS px), or null when it cannot be
 * proven from the root element (compressed, relative units without a viewBox). */
export function svgIntrinsicSize(bytes: Uint8Array): { width: number; height: number } | null {
  const text = new TextDecoder().decode(bytes.subarray(0, IMAGE_PROBE_BYTES))
  const root = /<svg\b[^>]*>/i.exec(text)?.[0]
  if (!root) return null
  const attr = (name: string) => new RegExp(`\\s${name}\\s*=\\s*["']([^"']*)["']`, 'i').exec(root)?.[1]
  const length = (value: string | undefined): number | null | undefined => {
    if (value === undefined) return undefined
    const match = /^\s*([0-9]*\.?[0-9]+(?:e[+-]?\d+)?)\s*(px|pt|pc|in|cm|mm)?\s*$/i.exec(value)
    if (!match) return null // %, em, ...: relative to a context the decoder supplies
    return Number(match[1]) * UNIT_PX[(match[2] ?? '').toLowerCase()]
  }
  const box = attr('viewBox')?.trim().split(/[\s,]+/).map(Number)
  const view = box && box.length === 4 && box.every(Number.isFinite) && box[2] > 0 && box[3] > 0 ? { width: box[2], height: box[3] } : null
  const w = length(attr('width')), h = length(attr('height'))
  // A relative length (%, em) has no intrinsic value without a viewBox.
  if ((w === null || h === null) && !view) return null
  let width = typeof w === 'number' ? w : w === null ? view!.width : undefined
  let height = typeof h === 'number' ? h : h === null ? view!.height : undefined
  // Browsers size an <svg> without width/height at 300x150 (or by its viewBox ratio).
  if (width === undefined && height === undefined) { width = 300; height = view ? 300 * view.height / view.width : 150 }
  else if (width === undefined) width = view ? height! * view.width / view.height : 300
  else if (height === undefined) height = view ? width * view.height / view.width : 150
  if (!(width! > 0) || !(height! > 0) || !Number.isFinite(width!) || !Number.isFinite(height!)) return null
  return { width: Math.ceil(width!), height: Math.ceil(height!) }
}

function fits(width: number, height: number, limits: ImageLimits): number | null {
  const pixels = width * height
  return pixels <= limits.maxPixels && pixels * 4 <= limits.maxDecodedBytes ? pixels * 4 : null
}
/** The representation to decode and its RGBA bytes, or null (refused). An SVG
 * that cannot be bounded is withdrawn so only an admissible raster is decoded. */
function admissible(asset: ImageCandidates, limits: ImageLimits): { asset: ImageCandidates; bytes: number } | null {
  const svgBytes = asset.svg?.bytes ?? (asset.data && looksLikeSvg(asset.data, asset.mime, asset.pathHint ?? asset.partPath) ? asset.data : undefined)
  const svgSize = svgBytes ? svgIntrinsicSize(svgBytes) : null
  const svgFits = svgSize ? fits(svgSize.width, svgSize.height, limits) : null
  const rasterBytes = asset.svg ? (asset.hasRaster === false ? undefined : asset.data) : svgBytes ? undefined : asset.data
  const probe = rasterBytes ? probeImageHeader(rasterBytes.subarray(0, IMAGE_PROBE_BYTES)) : null
  const rasterFits = probe && !('error' in probe) ? fits(probe.width, probe.height, limits) : null
  const svgPreferred = Boolean(asset.svg ? asset.svg.verdict.ok : svgBytes)
  if (svgBytes && svgPreferred && svgFits !== null) return { asset, bytes: Math.max(svgFits, rasterFits ?? 0) }
  if (rasterFits === null) return null
  // Withdraw an unbounded SVG candidate: decodeImageAsset then takes the raster.
  if (asset.svg) return { asset: { ...asset, svg: { ...asset.svg, verdict: { ok: false, reason: 'svg-size-unadmitted' } } }, bytes: rasterFits }
  return { asset, bytes: rasterFits }
}

export function prepareSheetAssets(doc: XlsxDocument, options: SheetAssetsOptions = {}): Promise<SheetAssetsLease> {
  const profile = options.profile ?? 'D', limits = imageLimitsFor(profile)
  const cache = options.imageCache ?? new SheetImageCache(imageBudgetFor(profile))
  return acquireFonts(doc, options.registerFont).then(async lease => {
    const held: ImageCandidates[] = []
    const releaseHeld = () => { for (const asset of held.splice(0)) cache.release(asset) }
    try {
      const sheet = doc.sheets[options.sheetIndex ?? 0]
      const indexed = doc.images ?? []
      const drawings = sheet?.drawings ?? []
      const wanted = [...imageIndexes(drawings)].filter(index => index >= 0 && index < indexed.length).sort((a, b) => a - b)
      const cached = collectCachedPictures(drawings)
      let decodedBytes = 0, refused = 0
      // Keyed by the original picture, so overlapping leases share one decode.
      const take = (original: ImageCandidates): Promise<CanvasImageSource | undefined> | undefined => {
        const admitted = admissible(original, limits)
        const pending = admitted && cache.acquire(original, admitted.asset, admitted.bytes, options.decodeImage)
        if (!admitted || !pending) { refused++; return undefined }
        held.push(original)
        decodedBytes += admitted.bytes
        return pending
      }
      const images: Array<CanvasImageSource | undefined> = new Array(indexed.length).fill(undefined)
      await Promise.all(wanted.map(async index => { images[index] = await take(indexed[index]) }))
      const byObject = new Map<object, CanvasImageSource | undefined>()
      await Promise.all(cached.map(async image => { byObject.set(image, await take(image)) }))
      let released = false
      return {
        assets: { images, imageFor: image => byObject.get(image), resolveFont: withFallbackFonts(lease.resolve, options.fallbackFonts) },
        fontDiagnostics: lease.diagnostics,
        decodedBytes,
        imageBudgetBytes: cache.limitBytes,
        refusedImages: refused,
        dispose: () => {
          if (released) return
          released = true
          lease.dispose()
          releaseHeld()
        },
      }
    } catch (error) { releaseHeld(); lease.dispose(); throw error }
  })
}
