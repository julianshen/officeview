/**
 * Font/image lease for the interactive sheet viewport (plan S5.2; spec §12.4).
 *
 * Unlike the static paintables, which decode every workbook image up front, the
 * sheet viewport leases ONE sheet's images: each picture is admitted from its
 * header dimensions (S1.1b `probeImageHeader`/`imageLimitsFor`) before any
 * decode, and charged twice its RGBA size (surface plus decoder transient)
 * against a bounded aggregate. Refused pictures paint as the renderer's blank
 * placeholder and are counted. Inactive sheets' images are never decoded.
 */
import type { XlsxDocument, XlsxDrawing } from '../../xlsx/types'
import type { ProfileId } from '../../xlsx/view/contracts'
import type { RenderSheetAssets } from '../../xlsx/render'
import { acquireFonts } from '../../core/fonts/register'
import { withFallbackFonts } from '../../core/fonts/fallback'
import type { FontDiagnostic } from '../../core/fonts/types'
import type { ImageCandidates } from '../../core/images'
import { collectCachedPictures, decodeImageAssets, releaseDecoded, type PaintOptions } from '../../render/paint'
import { IMAGE_PROBE_BYTES, imageLimitsFor, probeImageHeader } from '../../xlsx/storage/imageAdmission'

export interface SheetAssetsOptions extends PaintOptions {
  /** Sheet whose pictures are leased (default 0). */
  sheetIndex?: number
  profile?: ProfileId
}
export interface SheetAssetsLease {
  assets: RenderSheetAssets
  fontDiagnostics: FontDiagnostic[]
  /** Charged bytes (2 x RGBA per admitted picture) and the aggregate cap. */
  decodedBytes: number
  imageBudgetBytes: number
  /** Pictures of this sheet refused by header admission or the aggregate cap. */
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

/** Estimated decode charge of one picture, or null when it cannot be bounded. */
function decodeCharge(asset: ImageCandidates, limits: { maxPixels: number; maxDecodedBytes: number }): number | null {
  // An SVG representation rasterizes at an intrinsic size its bytes do not
  // bound up front: charge a fixed quarter of the per-image cap.
  if (asset.svg?.verdict.ok || asset.primarySvgVerdict?.ok) return Math.ceil(limits.maxDecodedBytes / 4) * 2
  if (!asset.data) return null
  const probe = probeImageHeader(asset.data.subarray(0, IMAGE_PROBE_BYTES))
  if ('error' in probe) return null
  const pixels = probe.width * probe.height
  if (pixels > limits.maxPixels || pixels * 4 > limits.maxDecodedBytes) return null
  return pixels * 4 * 2
}

export function prepareSheetAssets(doc: XlsxDocument, options: SheetAssetsOptions = {}): Promise<SheetAssetsLease> {
  const limits = imageLimitsFor(options.profile ?? 'D')
  // Aggregate: two images at the per-image cap including their transients.
  const imageBudgetBytes = limits.maxDecodedBytes * 2
  return acquireFonts(doc, options.registerFont).then(async lease => {
    let decoded: Array<CanvasImageSource | undefined> = []
    try {
      const sheet = doc.sheets[options.sheetIndex ?? 0]
      const indexed = doc.images ?? []
      const drawings = sheet?.drawings ?? []
      const wanted = [...imageIndexes(drawings)].filter(index => index >= 0 && index < indexed.length).sort((a, b) => a - b)
      const cached = collectCachedPictures(drawings)
      let charged = 0, refused = 0
      const admit = (asset: ImageCandidates): boolean => {
        const charge = decodeCharge(asset, limits)
        if (charge === null || charged + charge > imageBudgetBytes) { refused++; return false }
        charged += charge
        return true
      }
      const admittedIndexes = wanted.filter(index => admit(indexed[index]))
      const admittedCached = cached.filter(image => admit(image))
      decoded = await decodeImageAssets([...admittedIndexes.map(index => indexed[index]), ...admittedCached], options.decodeImage)
      const images: Array<CanvasImageSource | undefined> = new Array(indexed.length).fill(undefined)
      admittedIndexes.forEach((index, at) => { images[index] = decoded[at] })
      const byObject = new Map(admittedCached.map((image, at) => [image, decoded[admittedIndexes.length + at]]))
      let released = false
      return {
        assets: { images, imageFor: image => byObject.get(image), resolveFont: withFallbackFonts(lease.resolve, options.fallbackFonts) },
        fontDiagnostics: lease.diagnostics,
        decodedBytes: charged,
        imageBudgetBytes,
        refusedImages: refused,
        dispose: () => {
          if (released) return
          released = true
          lease.dispose()
          releaseDecoded(decoded)
        },
      }
    } catch (error) { releaseDecoded(decoded); lease.dispose(); throw error }
  })
}
