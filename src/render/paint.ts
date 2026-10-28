import { acquireFonts, type RegisterFont } from '../core/fonts/register'
import { withFallbackFonts, type FallbackFontsOptions } from '../core/fonts/fallback'
import type { FontDiagnostic } from '../core/fonts/types'
/**
 * Shared "document → canvas paintables" extraction. Single source of truth
 * used by <OfficeDoc> (React) and the pixel-diff/golden harness.
 */
import { collectDocImages, createMeasurer, layoutDocx, pageChartBound, pageChartBoxes, renderPages } from '../docx/layout'
import { fontFamilyCss as fontFamilyCssDocx } from '../docx/styles'
import { decodeImageAsset, releaseDecodedImage, type ImageCandidates, type ImageDecodeFn } from '../core/images'
import { inheritedCanvasLang } from '../core/text-metrics'
import { normalizeWatermark, paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'
import { computeMetrics, renderSheet, type RenderSheetAssets, type StaticExportLimit } from '../xlsx/render'
import { renderSlide, slideMetrics } from '../pptx/render'
import type { DocxDocument } from '../docx/types'
import type { ContentImageAsset, DrawingContent, DrawingContentShape } from '../drawing/content'
import type { XlsxDocument, XlsxSheet } from '../xlsx/types'
import { CHART_ENGINE_UNAVAILABLE, ensureChartEngine, ensureChartPainter, hasXlsxCharts, loadChartModule, type ChartReceiverLease } from '../xlsx/chartLoader'
import { ensureSparklinePainter } from '../xlsx/sparklineLoader'
import type { PptxDocument } from '../pptx/types'
import type { DrawingCoverageEntry } from '../drawing/coverage'
import type { ImageSelection } from '../core/svg'

export interface PageSpec {
  widthPx: number
  heightPx: number
}

/** Each extraction owns a lease. Release when its paint closures are no longer used. */
export interface PaintableArray extends Array<Paintable> {
  dispose: () => void
  fontDiagnostics: FontDiagnostic[]
  /** Chart preparation diagnostics by unit (spec 6.6): missed picture deadlines, description limits/refusals. */
  chartDiagnostics?: ChartPaintables['diagnostics']
}
function withLease(pages: Paintable[], dispose: () => void = () => {}, fontDiagnostics: FontDiagnostic[] = [], decoded: Array<CanvasImageSource | undefined> = [], charts?: Pick<ChartPaintables, 'diagnostics'>): PaintableArray {
  let released = false
  return Object.assign(pages, { fontDiagnostics, ...(charts?.diagnostics.length ? { chartDiagnostics: charts.diagnostics } : {}), dispose: () => {
    if (released) return
    released = true
    dispose()
    releaseDecoded(decoded)
  } })
}
/** Each decoded bitmap is closed once, even when shared by several assets. */
/** @internal */
export function releaseDecoded(decoded: Array<CanvasImageSource | undefined>): void {
  for (const image of new Set(decoded)) releaseDecodedImage(image)
}
export interface Paintable {
  spec: PageSpec
  /** Paints one white-background unit into a ctx scaled so 1 unit = 1 spec px. */
  paint: (ctx: CanvasRenderingContext2D) => void
  /** Set when the unit is a typed refusal (spec §11.5 `large-export-limit`): the
   * static bitmap cannot hold the sheet, so `paint` draws a notice instead of a
   * partial grid that would look complete. */
  refusal?: readonly StaticExportLimit[]
  /** The unit's charts in reading order: accessible descriptions with their placement in spec px (spec 6.6). A
   * notice entry ends a list that could not hold every chart (over the per-page limit, or refused memory). */
  charts?: readonly PaintableChart[]
  /** The unit's true chart count (may exceed the described entries). */
  chartTotal?: number
}
export type PaintableChart = import('../charts/paintables').PaintableChart
type ChartPaintables = import('../charts/paintables').ChartPaintables
type ChartPageSource = import('../charts/paintables').ChartPageSource
/** Pre-paint chart preparation through the lazy engine: user-shape pictures decoded in this realm within a bounded,
 * cancellable wait (a model transferred from a worker realm shows them at its first paint) and each unit's chart
 * descriptions, admitted to the receiving chart workspace before they are built. A failure leaves the pages without
 * descriptions, never without pages. */
async function chartMetadata(lease: ChartReceiverLease | undefined, source: ChartPageSource, options: PaintOptions | undefined, describe = true): Promise<ChartPaintables | undefined> {
  if (!lease) return undefined
  let any = false
  for (let page = 0; page < source.pages && !any; page++) any = source.bound(page) > 0
  if (!any) return undefined
  try {
    return await (await loadChartModule()).prepareChartPaintables(lease.context, source, {
      describe, ...(options?.signal ? { signal: options.signal } : {}), ...(options?.chartClock ? { clock: options.chartClock } : {}), ...(options?.chartDeadlineMs !== undefined ? { deadlineMs: options.chartDeadlineMs } : {}),
    })
  } catch { return undefined }
}
/** The unit's chart fields of a Paintable. */
const chartFields = (metadata: ChartPaintables | undefined, index: number): Pick<Paintable, 'charts' | 'chartTotal'> => {
  const page = metadata?.page(index)
  // A refused page whose record could not be charged keeps its static notice, with no total.
  return page?.charts.length ? { charts: page.charts, ...(page.total ? { chartTotal: page.total } : {}) } : {}
}
type ChartShape = {
  content?: DrawingContent<unknown, unknown>; children?: ChartShape[]; xEmu?: number; yEmu?: number; widthEmu?: number; heightEmu?: number
  rotationDeg?: number; flipH?: boolean; flipV?: boolean
  group?: { ext: { width: number; height: number }; chOff: { x: number; y: number }; chExt: { width: number; height: number } }
}
function countSceneCharts(nodes: readonly ChartShape[], depth = 0): number {
  let n = 0
  if (depth > 64) return 0
  for (const node of nodes) { if (node.content?.kind === 'xlsxChart') n++; if (node.children?.length) n += countSceneCharts(node.children, depth + 1) }
  return n
}
/** A 2D affine (canvas order: x' = a x + c y + e, y' = b x + d y + f). */
type Affine = readonly [number, number, number, number, number, number]
const multiply = (m: Affine, n: Affine): Affine => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]]
/**
 * A node's own transform, composed exactly as the scene painter applies it (drawing/scene-paint.ts): translate to
 * its centre, rotate, flip, translate back by half its size; its box is then (0, 0, w, h) in that space.
 */
function nodeAffine(parent: Affine, node: ChartShape): Affine {
  const x = node.xEmu ?? 0, y = node.yEmu ?? 0, w = node.widthEmu ?? 0, h = node.heightEmu ?? 0
  const r = (node.rotationDeg ?? 0) * Math.PI / 180, cos = Math.cos(r), sin = Math.sin(r)
  let m = multiply(parent, [1, 0, 0, 1, x + w / 2, y + h / 2])
  if (r) m = multiply(m, [cos, sin, -sin, cos, 0, 0])
  if (node.flipH || node.flipV) m = multiply(m, [node.flipH ? -1 : 1, 0, 0, node.flipV ? -1 : 1, 0, 0])
  return multiply(m, [1, 0, 0, 1, -w / 2, -h / 2])
}
/** A slide's (or sheet's) charts in z-order, each at the page-space bounding box of its own transformed box. */
function placedSceneCharts(nodes: readonly ChartShape[], scale: number, out: import('../charts/paintables').PlacedChart[] = [], parent: Affine = [1, 0, 0, 1, 0, 0], depth = 0): import('../charts/paintables').PlacedChart[] {
  if (depth > 64) return out
  for (const node of nodes) {
    const m = nodeAffine(parent, node), w = node.widthEmu ?? 0, h = node.heightEmu ?? 0
    if (node.content?.kind === 'xlsxChart') {
      const xs: number[] = [], ys: number[] = []
      for (const [px, py] of [[0, 0], [w, 0], [0, h], [w, h]]) { xs.push(m[0] * px + m[2] * py + m[4]); ys.push(m[1] * px + m[3] * py + m[5]) }
      const x0 = Math.min(...xs), y0 = Math.min(...ys)
      out.push({ chart: node.content.chart, x: x0 * scale, y: y0 * scale, width: (Math.max(...xs) - x0) * scale, height: (Math.max(...ys) - y0) * scale })
    }
    if (node.children?.length) {
      // A group's children are in its child space: scaled by ext/chExt and offset by chOff, as the painter does.
      const g = node.group
      const valid = g && g.chExt.width > 0 && g.chExt.height > 0 && [g.ext.width, g.ext.height, g.chOff.x, g.chOff.y].every(Number.isFinite)
      const child = valid ? multiply(multiply(m, [g.ext.width / g.chExt.width, 0, 0, g.ext.height / g.chExt.height, 0, 0]), [1, 0, 0, 1, -g.chOff.x, -g.chOff.y]) : m
      placedSceneCharts(node.children, scale, out, child, depth + 1)
    }
  }
  return out
}
const sceneSource = (units: ReadonlyArray<readonly ChartShape[]>, scale: (unit: number) => number): ChartPageSource => ({
  pages: units.length, bound: unit => countSceneCharts(units[unit]), collect: unit => placedSceneCharts(units[unit], scale(unit)),
})

/** @internal One static sheet unit (S2.3 budgeted path): the whole-sheet raster,
 * or a typed `large-export-limit` refusal unit. Shared by getPaintables and the
 * sheet viewer's print preparation, which supplies admitted image assets. */
export function sheetPaintable(sheet: XlsxSheet, assets: RenderSheetAssets, watermark?: ResolvedWatermark): Paintable {
  const m = computeMetrics(sheet)
  if (m.exportLimits?.length) return refusalPaintable(sheet, m.exportLimits, watermark)
  return {
    spec: { widthPx: m.widthPx, heightPx: m.heightPx },
    paint: (ctx) => renderSheet(sheet, ctx, m, watermark, assets),
  }
}

const REFUSAL_SPEC: PageSpec = { widthPx: 480, heightPx: 96 }
export const LARGE_EXPORT_MESSAGE = 'This sheet is too large to export as one image. Choose a range or print selected pages.'
/** A static unit that explains the refusal instead of painting a truncated grid. */
function refusalPaintable(sheet: XlsxSheet, limits: readonly StaticExportLimit[], watermark?: ResolvedWatermark): Paintable {
  // Same deduped diagnostics channel the renderer uses for unsupported formats.
  sheet.diagnostics ??= []
  for (const limit of limits) {
    if (!sheet.diagnostics.some(d => d.feature === 'large-export-limit' && d.message === limit.message)) {
      sheet.diagnostics.push({ kind: 'large-export-limit', feature: 'large-export-limit', message: limit.message })
    }
  }
  return {
    spec: { ...REFUSAL_SPEC },
    refusal: limits,
    paint: (ctx) => {
      ctx.save()
      try {
        ctx.fillStyle = '#ffffff'
        ctx.fillRect(0, 0, REFUSAL_SPEC.widthPx, REFUSAL_SPEC.heightPx)
        if (watermark) paintWatermark(ctx, REFUSAL_SPEC, watermark)
        ctx.fillStyle = '#444444'
        ctx.textBaseline = 'alphabetic'
        ctx.font = 'bold 13px sans-serif'
        ctx.fillText(`Sheet "${sheet.name}" is too large for a static image.`, 12, 28)
        ctx.font = '12px sans-serif'
        ctx.fillText('This sheet is too large to export as one image.', 12, 52)
        ctx.fillText('Choose a range or print selected pages.', 12, 72)
      } finally { ctx.restore() }
    },
  }
}

async function measurerFromDoc(resolve: (family: string) => string): Promise<ReturnType<typeof createMeasurer>> {
  // headless-safe: DOM canvas when available (browser/jsdom), node-canvas fallback
  let ctx: CanvasRenderingContext2D | null = null
  try {
    if (typeof document !== 'undefined') {
      const c = document.createElement('canvas')
      c.width = 4
      c.height = 4
      ctx = c.getContext('2d')
    }
  } catch {
    // fall through to node-canvas
  }
  if (!ctx) {
    // @vite-ignore: resolved only in node/bun (golden harness), never bundled
    const { createCanvas } = (await import(/* @vite-ignore */ 'canvas')) as unknown as { createCanvas: (w: number, h: number) => { getContext: (t: string) => CanvasRenderingContext2D | null } }
    ctx = createCanvas(4, 4).getContext('2d')
  }
  if (!ctx) throw new Error('no 2d context available for measurement')
  // Pages paint on canvases attached to this document; measure under the same
  // language so generic font families resolve to the same face.
  const lang = inheritedCanvasLang()
  if (lang && 'lang' in ctx) (ctx as { lang: string }).lang = lang
  return createMeasurer(ctx, resolve)
}

/** Options for stamping a watermark across every painted unit. */
export interface PaintOptions {
  /** Optional headless font adapter; its identity is the registration cache key. */
  registerFont?: RegisterFont
  /**
   * Applied to every page/sheet/slide. Resolved once here rather than per
   * canvas, so a bad watermark is rejected before any painting happens.
   */
  watermark?: WatermarkOptions | ResolvedWatermark
  /** Optional image decoder override (tests count decodes). Defaults to decodeImage. */
  decodeImage?: ImageDecodeFn
  /**
   * Optional explicit fallback chain/region, composed onto every resolver.
   * Absent by default: resolvers pass through byte-identical, so default
   * rendering (and goldens) never move; pinned chains are host-deterministic.
   */
  fallbackFonts?: FallbackFontsOptions
  /**
   * The receiving document's chart budget (spec 6.4): transferred chart models are charged to this context, or to
   * one created once per document with `chartProfile` (the viewer profile; mobile is 32 MiB; default desktop).
   */
  chartContext?: import('../charts/context').ChartDocumentContext
  chartProfile?: import('../charts/workspace').ChartProfile
  /** The receiving document's live parent budget (an interactive viewer's main-thread ledger) for that created context. */
  chartParent?: import('../charts/workspace').ChartParentLedger
  /** Aborts pre-paint chart preparation (the viewer unmounted or changed document). */
  signal?: AbortSignal
  /** @internal The chart picture deadline's clock and length (tests drive them; default real time, 1.5 s). */
  chartClock?: import('../charts/paintables').ChartClock
  /** @internal */
  chartDeadlineMs?: number
  /** @internal The slide appearance chunk's loader (tests inject a failure). */
  __appearanceLoader?: () => Promise<{ enrichSlideAppearance: (doc: PptxDocument) => Promise<void> }>
}
/** The lazy slide appearance enrichment (slideshow spec 8): backgrounds and master/layout artwork. */
const loadAppearance = () => import('../pptx/appearance')

/** Byte equality for shared-asset decode reuse (length pre-checked by callers). */
function bytesEqual(a: Uint8Array | undefined, b: Uint8Array | undefined): boolean {
  if (a === b) return true
  if (!a || !b || a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

/**
 * Decode one bitmap per COMPLETE candidate pair (raster bytes + SVG bytes).
 * Same PNG with a different SVG is a distinct asset; identical pairs decode
 * once. The selected representation/reason is propagated to every alias's
 * selection record so coverage stays consistent.
 */
/** @internal Shared with the interactive sheet asset lease (components/xlsx/sheetAssets). */
export async function decodeImageAssets(assets: Array<ImageCandidates & { drawing?: unknown }>, decode?: ImageDecodeFn): Promise<Array<CanvasImageSource | undefined>> {
  // Complete-pair identity: raster bytes + SVG bytes + SVG verdict state +
  // primary-SVG verdict state + fallback availability. Missing/external
  // candidates (byte-less) and ordinary rasters therefore never collapse.
  const svgState = (svg: ImageCandidates['svg']): string => svg ? (svg.verdict.ok ? 'ok' : `rej:${(svg.verdict as { reason: string }).reason}`) : 'none'
  const primaryState = (verdict: ImageCandidates['primarySvgVerdict']): string => verdict ? (verdict.ok ? 'ok' : `rej:${(verdict as { reason: string }).reason}`) : 'none'
  const tone = (duotone: ImageCandidates['duotone']): string => duotone ? JSON.stringify([duotone.dark, duotone.light]) : ''
  const canonical: Array<{ data?: Uint8Array; svg?: Uint8Array; svgState: string; primaryState: string; svgOnly: boolean; duotone: string; decoded: Promise<CanvasImageSource | undefined> }> = []
  const representative: number[] = []
  for (const asset of assets) {
    if (asset.drawing || (!asset.data && !asset.svg)) { representative.push(-1); continue }
    const state = svgState(asset.svg), primary = primaryState(asset.primarySvgVerdict), svgOnly = asset.hasRaster === false, duotone = tone(asset.duotone)
    const found = canonical.findIndex(entry => entry.svgState === state && entry.primaryState === primary && entry.svgOnly === svgOnly && entry.duotone === duotone && bytesEqual(entry.data, asset.data) && bytesEqual(entry.svg, asset.svg?.bytes))
    if (found >= 0) { representative.push(found); continue }
    canonical.push({ data: asset.data, svg: asset.svg?.bytes, svgState: state, primaryState: primary, svgOnly, duotone, decoded: decodeImageAsset(asset, decode) })
    representative.push(canonical.length - 1)
  }
  const decoded = await Promise.all(assets.map((_, index) => representative[index] < 0 ? undefined : canonical[representative[index]].decoded))
  // Propagate the decoded outcome from each group's representative (the asset
  // whose selection decodeImageAsset mutated) to every alias's selection.
  const groups = new Map<number, number[]>()
  for (let index = 0; index < representative.length; index++) {
    const rep = representative[index]
    if (rep < 0) continue
    const list = groups.get(rep) ?? []
    list.push(index)
    groups.set(rep, list)
  }
  for (const members of groups.values()) {
    const lead = assets[members[0]].imageSelection
    if (!lead) continue
    for (const index of members) {
      const selection = assets[index].imageSelection
      if (selection && selection !== lead) { selection.phase = lead.phase; selection.representation = lead.representation; selection.reason = lead.reason }
    }
  }
  return decoded
}

type FallbackCarrier = { chartFallback?: { unsupported?: string }; imageIndex?: number; image?: { imageSelection?: ImageSelection }; children?: FallbackCarrier[] }
const savedCoverage = new WeakMap<DrawingCoverageEntry, Pick<DrawingCoverageEntry, 'selectedRepresentation' | 'reason'>>()
/** A chart's selected fallback picture that failed to decode paints the chart's card (spec 6.3). Its coverage
 * records that, recomputed at each preparation so a later successful decode restores the fallback entry. */
function recordChartFallbacks(lists: Array<DrawingCoverageEntry[] | undefined>, failed: Map<ImageSelection, string>): void {
  for (const entry of lists.flatMap(list => list ?? [])) {
    const reason = entry.imageSelection && failed.get(entry.imageSelection), saved = savedCoverage.get(entry)
    if (reason !== undefined) {
      if (!saved) savedCoverage.set(entry, { selectedRepresentation: entry.selectedRepresentation, reason: entry.reason })
      entry.selectedRepresentation = 'chart-placeholder'
      entry.reason = `chart-fallback-unusable: ${reason}`
    } else if (saved) { Object.assign(entry, saved); savedCoverage.delete(entry) }
  }
}
/** Scene nodes (PPTX shapes, XLSX drawings) whose chart-fallback picture did not decode. */
function failedSceneFallbacks(nodes: FallbackCarrier[], decoded: Array<CanvasImageSource | undefined>, failed = new Map<ImageSelection, string>()): Map<ImageSelection, string> {
  for (const node of nodes) {
    if (node.chartFallback && node.imageIndex !== undefined && !decoded[node.imageIndex] && node.image?.imageSelection) failed.set(node.image.imageSelection, node.chartFallback.unsupported ?? '')
    failedSceneFallbacks(node.children ?? [], decoded, failed)
  }
  return failed
}

interface ContentCarrier { content?: DrawingContent<unknown, unknown>; children?: ContentCarrier[] }
/** Append cached pictures after indexed assets so ordinary scene image indices
 * stay stable. Identity deduplication also bounds handwritten cyclic models. */
/** @internal */
export function collectCachedPictures(nodes: ContentCarrier[]): ContentImageAsset[] {
  const images: ContentImageAsset[] = [], seenImages = new Set<ContentImageAsset>()
  const seenNodes = new Set<ContentCarrier>(), seenShapes = new Set<DrawingContentShape<unknown>>()
  const shapes = (items: DrawingContentShape<unknown>[]) => {
    for (const shape of items) {
      if (seenShapes.has(shape)) continue
      seenShapes.add(shape)
      if (shape.image && !seenImages.has(shape.image)) { seenImages.add(shape.image); images.push(shape.image) }
      if (shape.children) shapes(shape.children)
    }
  }
  const walk = (items: ContentCarrier[]) => {
    for (const node of items) {
      if (seenNodes.has(node)) continue
      seenNodes.add(node)
      if (node.content?.kind === 'diagram') shapes(node.content.shapes)
      if (node.children) walk(node.children)
    }
  }
  walk(nodes)
  return images
}

/** The receiving document for its transferred chart models (spec 6.4). */
export const chartReceiver = (doc: object, options?: PaintOptions) => ({ owner: doc, ...(options?.chartProfile ? { profile: options.chartProfile } : {}), ...(options?.chartContext ? { context: options.chartContext } : {}), ...(options?.chartParent ? { parent: options.chartParent } : {}) })
/** Prepare embedded fonts before measurement/painting. Low-level renderSlide stays synchronous. */
export function getPaintables(
  doc: DocxDocument | XlsxDocument | PptxDocument,
  options?: PaintOptions,
): Promise<PaintableArray> {
  const watermark = normalizeWatermark(options?.watermark as WatermarkOptions | undefined)
  if ('sections' in doc) {
    return acquireFonts(doc, options?.registerFont).then(async (lease) => {
      // Decoded before the lease exists: a later failure must close them too.
      let decoded: Array<CanvasImageSource | undefined> = []
      // The receiving chart context these pages charge their transferred charts to (released with the pages).
      let charts: ChartReceiverLease | undefined
      let metadata: ChartPaintables | undefined
      try {
        const baseFamilyCss = (family: string) => fontFamilyCssDocx(lease.resolve(family))
        const measure = await measurerFromDoc(withFallbackFonts(baseFamilyCss, options?.fallbackFonts))
        const pages = layoutDocx(doc, measure)
        // One decode per unique candidate pair; failures degrade to a blank slot.
        const docImages = collectDocImages(doc)
        charts = await ensureChartEngine(docImages.map(image => ({ content: image.drawing })), chartReceiver(doc, options))
        decoded = await decodeImageAssets(docImages, options?.decodeImage)
        const decodedByObject = new Map<ContentImageAsset, CanvasImageSource | undefined>(docImages.map((img, index) => [img, decoded[index]]))
        const failed = new Map<ImageSelection, string>()
        docImages.forEach((image, index) => { if (image.chartFallback && !decoded[index] && image.imageSelection) failed.set(image.imageSelection, image.chartFallback.unsupported ?? '') })
        recordChartFallbacks([doc.drawingCoverage], failed)
        // Each page paints on its own canvas, so PAGE/NUMPAGES fields must be
        // resolved against the whole document, not the single-page array.
        const total = pages.length
        metadata = await chartMetadata(charts, {
          pages: pages.length,
          bound: index => pageChartBound(pages[index]),
          collect: index => pageChartBoxes(pages[index], measure, index + 1, total).map(box => ({ chart: (box.drawing as { chart: import('../charts/model').ChartModel }).chart, x: box.xPx, y: box.yPx, width: box.widthPx, height: box.heightPx })),
        }, options)
        return withLease(pages.map((page, index) => ({
          spec: { widthPx: Math.ceil(page.widthPx), heightPx: Math.ceil(page.heightPx) },
          ...chartFields(metadata, index),
          paint: (ctx) => {
            ctx.fillStyle = '#ffffff'
            ctx.fillRect(0, 0, page.widthPx, page.heightPx)
            renderPages([page], ctx, decoded, {
              pageNumberStart: index + 1,
              totalPages: total,
              resolveFont: lease.resolve,
              ...(watermark ? { watermark } : {}),
              assets: {
                imageFor: (image) => decodedByObject.get(image),
                ...(charts ? { chartContext: charts.context } : {}),
                ...(options?.fallbackFonts ? { fallbackFonts: options.fallbackFonts } : {}),
                resolveFont: lease.resolve,
              },
            })
          },
        })), () => { metadata?.release(); charts?.release(); lease.dispose() }, lease.diagnostics, decoded, metadata)
      } catch (error) { releaseDecoded(decoded); metadata?.release(); charts?.release(); lease.dispose(); throw error }
    })
  }
  if ('sheets' in doc) {
    return acquireFonts(doc, options?.registerFont).then(async lease => {
      let decoded: Array<CanvasImageSource | undefined> = []
      let charts: ChartReceiverLease | undefined
      let metadata: ChartPaintables | undefined
      try {
        const cached = collectCachedPictures(doc.sheets.flatMap(sheet => sheet.drawings ?? []))
        const drawings = doc.sheets.flatMap(sheet => sheet.drawings ?? [])
        charts = await ensureChartPainter(drawings, chartReceiver(doc, options))
        await ensureSparklinePainter(doc.sheets)
        const indexed = doc.images ?? []
        decoded = await decodeImageAssets([...indexed, ...cached], options?.decodeImage)
        const decodedByObject = new Map(cached.map((image, index) => [image, decoded[indexed.length + index]]))
        recordChartFallbacks([doc.drawingCoverage, ...doc.sheets.map(sheet => sheet.drawingCoverage)], failedSceneFallbacks(doc.sheets.flatMap(sheet => sheet.drawings ?? []), decoded))
        // Sheet charts are described by the sheet viewer; here only their pictures are prepared (no placement).
        metadata = await chartMetadata(charts, sceneSource(doc.sheets.map(sheet => sheet.drawings ?? []), () => 1 / 9525), options, false)
        metadata?.release()
        return withLease(doc.sheets.map((sheet): Paintable =>
          sheetPaintable(sheet, { images: decoded, imageFor: image => decodedByObject.get(image), resolveFont: withFallbackFonts(lease.resolve, options?.fallbackFonts), ...(charts ? { chartContext: charts.context } : {}) }, watermark),
        ), () => { charts?.release(); lease.dispose() }, lease.diagnostics, decoded, charts || !hasXlsxCharts(drawings) ? metadata : { diagnostics: [CHART_ENGINE_UNAVAILABLE] })
      } catch (error) { releaseDecoded(decoded); metadata?.release(); charts?.release(); lease.dispose(); throw error }
    })
  }
  // A failed appearance chunk leaves the pre-Stage-0 appearance, diagnosed, never a failed paint.
  return (options?.__appearanceLoader ?? loadAppearance)().then(m => m.enrichSlideAppearance(doc)).catch(() => {
    doc.diagnostics ??= []
    if (!doc.diagnostics.some(d => d.kind === 'appearance-unavailable')) doc.diagnostics.push({ kind: 'appearance-unavailable', message: 'Slide backgrounds and master artwork could not be loaded; slides paint without them' })
  }).then(() => acquireFonts(doc, options?.registerFont)).then(async lease => {
    let decoded: Array<CanvasImageSource | undefined> = []
    let charts: ChartReceiverLease | undefined
    let metadata: ChartPaintables | undefined
    try {
      const sm = slideMetrics(doc)
      const scenes = doc.slides.flatMap(slide => [...(slide.underlay ?? []), ...slide.shapes])
      const cached = collectCachedPictures(scenes)
      charts = await ensureChartEngine(doc.slides.flatMap(slide => slide.shapes), chartReceiver(doc, options))
      decoded = await decodeImageAssets([...doc.images, ...cached], options?.decodeImage)
      const decodedByObject = new Map(cached.map((image, index) => [image, decoded[doc.images.length + index]]))
      recordChartFallbacks([doc.drawingCoverage], failedSceneFallbacks(scenes, decoded))
      metadata = await chartMetadata(charts, sceneSource(doc.slides.map(slide => slide.shapes), unit => sm.widthPx / (doc.slides[unit].widthEmu || 1)), options)
      return withLease(doc.slides.map((slide, index) => ({
        spec: { widthPx: sm.widthPx, heightPx: sm.heightPx },
        ...chartFields(metadata, index),
        paint: ctx => renderSlide(slide, ctx, sm, decoded, watermark, withFallbackFonts(lease.resolve, options?.fallbackFonts), { imageFor: image => decodedByObject.get(image), ...(charts ? { chartContext: charts.context } : {}) }),
      })), () => { metadata?.release(); charts?.release(); lease.dispose() }, lease.diagnostics, decoded, metadata)
    } catch (error) { releaseDecoded(decoded); metadata?.release(); charts?.release(); lease.dispose(); throw error }
  })
}
