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
import { ensureChartPainter, hasXlsxCharts, type ChartReceiverLease } from '../../xlsx/chartLoader'
import { ensureSparklinePainter } from '../../xlsx/sparklineLoader'
import { acquireFonts } from '../../core/fonts/register'
import { withFallbackFonts } from '../../core/fonts/fallback'
import type { FontDiagnostic } from '../../core/fonts/types'
import { releaseDecodedImage, type ImageCandidates, type ImageDecodeFn } from '../../core/images'
import { looksLikeSvg } from '../../core/svg'
import { chartReceiver, collectCachedPictures, decodeImageAssets, sheetPaintable, type Paintable, type PaintableArray, type PaintOptions } from '../../render/paint'
import type { MainMemoryLedger } from '../../xlsx/view/mainLedger'
import type { ChartParentLedger } from '../../charts/workspace'
import { normalizeWatermark, type WatermarkOptions } from '../../core/watermark'
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
  /** The document whose chart context receives transferred models (default `doc`; a viewer passes its own document
   * so successive composed sheet documents share one chart facet). */
  chartOwner?: object
}
/**
 * The viewer's live main-thread ledger as a chart parent budget (spec 6.4): chart paint charges are admitted against
 * the same hard main budget as the viewer's drawings and recorded under `drawings`/`charts`, refunded on release.
 */
export function mainChartParent(ledger: Pick<MainMemoryLedger, 'bytes' | 'set'>, budgetBytes: number): ChartParentLedger {
  let held = 0
  return {
    reserve: bytes => { if (ledger.bytes + bytes > budgetBytes) return false; held += bytes; ledger.set('drawings', 'charts', held); return true },
    release: (_tag, bytes) => { held = Math.max(0, held - bytes); ledger.set('drawings', 'charts', held) },
  }
}
export interface SheetAssetsLease {
  assets: RenderSheetAssets
  fontDiagnostics: FontDiagnostic[]
  /** RGBA bytes of this lease's admitted pictures, and the aggregate limit. */
  decodedBytes: number
  imageBudgetBytes: number
  /** Pictures of this sheet refused by admission or the aggregate limit. */
  refusedImages: number
  /** The chart chunk did not load: this sheet's charts paint as cards. */
  chartsUnavailable?: true
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
const SIZING_EXACT = new Set(['width', 'height', 'viewBox', 'style'])
const SIZING = new Set([...SIZING_EXACT].map(name => name.toLowerCase()))
/** Index just past the doctype starting at `at` (internal subset and quotes honoured), or -1. */
function skipDoctype(text: string, at: number): number {
  let depth = 0, quote = ''
  for (let i = at; i < text.length; i++) {
    const c = text[i]
    if (quote) { if (c === quote) quote = ''; continue }
    if (c === '"' || c === "'") quote = c
    else if (c === '[') depth++
    else if (c === ']') depth--
    else if (c === '>' && depth <= 0) return i + 1
  }
  return -1
}
/**
 * Intrinsic raster size of an SVG document (CSS px), or null when it cannot be
 * proven. The real root element is found by skipping the XML declaration,
 * processing instructions, comments and a doctype (never a regex over markup
 * that may sit inside them). The size must come from absolute `width` and
 * `height` attributes on that root; anything CSS or the decoder's context can
 * change withdraws the proof: a `style` attribute, any `<style>` element or
 * `xml-stylesheet` instruction, relative units (%, em, ...), missing dimensions,
 * entities, a prefixed root or compressed bytes.
 */
export function svgIntrinsicSize(bytes: Uint8Array): { width: number; height: number } | null {
  let text: string
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes) } catch { return null }
  let i = text.charCodeAt(0) === 0xfeff ? 1 : 0
  for (;;) {
    while (i < text.length && /\s/.test(text[i])) i++
    if (text.startsWith('<?', i)) {
      const end = text.indexOf('?>', i + 2)
      if (end < 0 || /^<\?xml-stylesheet\b/i.test(text.slice(i, end))) return null
      i = end + 2
    } else if (text.startsWith('<!--', i)) {
      const end = text.indexOf('-->', i + 4)
      if (end < 0) return null
      i = end + 3
    } else if (/^<!DOCTYPE\b/i.test(text.slice(i, i + 10))) {
      i = skipDoctype(text, i)
      if (i < 0) return null
    } else break
  }
  const name = /^<([A-Za-z_][\w.:-]*)/.exec(text.slice(i, i + 64))
  if (!name || name[1] !== 'svg') return null
  // Root attributes, quote-aware, up to the end of the start tag.
  const attrs = new Map<string, string>()
  const pattern = /\s+([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/y
  let at = i + name[0].length
  for (;;) {
    pattern.lastIndex = at
    const match = pattern.exec(text)
    if (!match) break
    // XML attribute names are case-sensitive: keep them exact. A duplicate name or a
    // case variant of a sizing attribute makes the decoder's choice ambiguous.
    const key = match[1]
    if (attrs.has(key)) return null
    if (SIZING.has(key.toLowerCase()) && !SIZING_EXACT.has(key)) return null
    attrs.set(key, match[2] ?? match[3])
    at = pattern.lastIndex
  }
  if (!/^\s*\/?>/.test(text.slice(at, at + 8))) return null
  if (attrs.has('style')) return null
  const rest = text.slice(at)
  if (/<(?:[\w.-]+:)?style[\s>/]/i.test(rest) || /<\?xml-stylesheet\b/i.test(rest)) return null
  const length = (value: string | undefined): number | null => {
    if (value === undefined) return null
    const match = /^\s*([0-9]*\.?[0-9]+(?:e[+-]?\d+)?)\s*(px|pt|pc|in|cm|mm)?\s*$/i.exec(value)
    return match ? Number(match[1]) * UNIT_PX[(match[2] ?? '').toLowerCase()] : null
  }
  const width = length(attrs.get('width')), height = length(attrs.get('height'))
  if (width === null || height === null || !(width > 0) || !(height > 0) || !Number.isFinite(width) || !Number.isFinite(height)) return null
  return { width: Math.ceil(width), height: Math.ceil(height) }
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
    // This lease's receiving chart context (per document, parent and profile), carried by its assets, released with it.
    let charts: ChartReceiverLease | undefined
    const releaseHeld = () => { for (const asset of held.splice(0)) cache.release(asset); charts?.release() }
    try {
      const sheet = doc.sheets[options.sheetIndex ?? 0]
      const indexed = doc.images ?? []
      const drawings = sheet?.drawings ?? []
      const wanted = [...imageIndexes(drawings)].filter(index => index >= 0 && index < indexed.length).sort((a, b) => a - b)
      const cached = collectCachedPictures(drawings)
      // The receiving document's chart budget for transferred models (spec 6.4): a supplied context as is; else one
      // per viewer document, sized by the viewer profile and parented to the viewer's live ledger (chartParent).
      charts = await ensureChartPainter(drawings, chartReceiver(options.chartOwner ?? doc, { chartProfile: profile === 'M' ? 'mobile' : 'desktop', ...options }))
      await ensureSparklinePainter(sheet ? [sheet] : [])
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
        assets: { images, imageFor: image => byObject.get(image), resolveFont: withFallbackFonts(lease.resolve, options.fallbackFonts), ...(charts ? { chartContext: charts.context } : {}) },
        fontDiagnostics: lease.diagnostics,
        decodedBytes,
        imageBudgetBytes: cache.limitBytes,
        refusedImages: refused,
        ...(!charts && hasXlsxCharts(drawings) ? { chartsUnavailable: true as const } : {}),
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

export interface PrintUnitsOptions extends SheetAssetsOptions {
  /** Check before acquiring assets, including programmatic callers. */
  allowed?: () => boolean
  /** Lease factory; the viewer passes its own so print shares its image budget. */
  assetLoader?: typeof prepareSheetAssets
}
/**
 * Static print units for every sheet through the S2.3 budgeted path, with each
 * sheet's pictures leased through the same admission and shared image budget
 * as the viewport (never the unrestricted workbook-wide decode). Sheets are
 * leased one at a time; release everything with dispose() after printing.
 */
/** Print units plus, per unit, the pictures admission refused (shown as notices). */
export interface PrintUnits extends PaintableArray { refusedImages: number[] }
export async function preparePrintUnits(doc: XlsxDocument, options: PrintUnitsOptions = {}): Promise<PrintUnits> {
  const permitted = () => options.allowed?.() ?? true
  const denied = () => Object.assign(new Error('Printing is not permitted.'), { code: 'permission-denied' })
  if (!permitted()) throw denied()
  const { assetLoader = prepareSheetAssets, ...lease } = options
  const watermark = normalizeWatermark(options.watermark as WatermarkOptions | undefined)
  const leases: SheetAssetsLease[] = []
  const release = () => { for (const held of leases.splice(0)) held.dispose() }
  try {
    const units: Paintable[] = []
    for (let sheetIndex = 0; sheetIndex < doc.sheets.length; sheetIndex++) {
      if (!permitted()) throw denied()
      const held = await assetLoader(doc, { ...lease, sheetIndex })
      leases.push(held)
      if (!permitted()) throw denied()
      units.push(sheetPaintable(doc.sheets[sheetIndex], held.assets, watermark))
    }
    let released = false
    return Object.assign(units, {
      fontDiagnostics: leases[0]?.fontDiagnostics ?? [],
      refusedImages: leases.map(held => held.refusedImages),
      dispose: () => { if (!released) { released = true; release() } },
    })
  } catch (error) { release(); throw error }
}
