/**
 * Interactive viewport painter (spec §§5.3, 11.1, 11.2; plans S2.3/S5.1/S5.2).
 * Lives outside `render.ts` so the static renderer, which every consumer loads
 * eagerly, never reaches the viewport planner; only the interactive sheet
 * viewport (lazy chunk) and tests import it.
 */
import type { XlsxSheet } from '../types'
import type { CacheIdentity, LogicalViewport, WorkerDiagnostic } from './contracts'
import type { SourceCell } from './coordinates'
import { COARSE_NOTICE, emptyWork, planViewport, type CoarseGroup, type FrameBudget, type MergeQueryCache, type SheetGeometryResult, type ViewportPlan, type ViewportProfile, type ViewportWork } from './viewport'
import { buildCellRecords, LayoutCache, type CellLayout, type CellLayoutHost } from './textLayout'
import { layoutCellBody, paintCellLayout, paintSheetDrawings, textMeasurer, type RenderSheetAssets } from '../render'
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../../core/watermark'
import { CSS_PX_PER_INCH, EMU_PER_INCH } from '../../core/geometry'

/** The host side of canonical records (spec §11.2): the same formatting and text
 * layout the static renderer uses, laid out in a box at the origin. Measurement
 * is cached per host by face and text. */
export function createCellLayoutHost(ctx: CanvasRenderingContext2D, sheet: Partial<XlsxSheet> = {}): CellLayoutHost {
  const measured = textMeasurer(ctx)
  // Format diagnostics land on the caller's sheet when one is given.
  const target = sheet as XlsxSheet
  return {
    layoutBody: (cell, w, h, overflow) => layoutCellBody(ctx, target, cell, 0, 0, w, h, measured, overflow),
    canReuseOverflow: (body, overflow) => {
      const text = body.text
      if (!text) return true
      // The canonical plain path uses allowances only to add clipX. If neither
      // allowance clips the text band, its formatting and paint data are equal.
      // Plain-path clipping normally fails the extent check too. Keep clipX
      // explicit as a conservative certificate: shrink-to-fit can force
      // NO_OVERFLOW independently of the requested allowance. Such bodies
      // need not share keys. Complex layouts retain exact keys (absent vs
      // extended clips can differ).
      return text.kind === 'plain' && !text.clipX && (!overflow || (
        Math.max(0, body.x - text.tx) <= overflow.left
        && Math.max(0, text.tx + text.width - (body.x + body.w)) <= overflow.right
      ))
    },
  }
}

export interface ViewportPaintResult {
  pending?: boolean
  kind: ViewportPlan['kind']; diagnostics: readonly WorkerDiagnostic[]
  /** Canonical records the frame painted (spec §11.2), in paint order. */
  records: readonly CellLayout[]
  /** Summarized sub-detail pixel groups (whole frame when `kind` is coarse). */
  coarse: readonly CoarseGroup[]
  /** The exact active-cell record, painted in any mode. */
  active?: CellLayout
  work: ViewportWork
}
export interface ViewportPaintOptions {
  /** Worker-planned records; canonical layout and paint remain on this host. */
  preparedPlan?: ViewportPlan
  profile?: ViewportProfile
  /** Publication owner of the records; layouts are cached under it. */
  owner?: CacheIdentity
  /** Cross-frame formatted-value/layout cache (plan S5.1). */
  cache?: LayoutCache
  merges?: MergeQueryCache
  /** Shared per-frame budget across the tiles of one frame (see PlanOptions). */
  frameBudget?: FrameBudget
  activeCell?: SourceCell
  /** Painted in viewport space: protection marks follow the visible pane. */
  watermark?: WatermarkOptions | ResolvedWatermark
}
const VIEWPORT_NOTICE: Record<Exclude<ViewportPlan['kind'], 'detail'>, string> = {
  pending: 'Loading…',
  coarse: COARSE_NOTICE,
  diagnostic: 'This sheet cannot be laid out',
}
/** Coarse summaries without a resolved aggregate paint a neutral "pending" tone,
 * never a white (asserted blank) region. */
const COARSE_FILL = '#eceff3'

/**
 * Interactive full-coordinate seam (spec §§5.3, 11.1, 11.2, 11.5). ctx state: 1
 * unit = 1 logical CSS px at the viewport's top-left; the caller applies zoom/DPR.
 * Geometry comes from the full S2.1 axes/projection (`buildSheetGeometry`), so a
 * source row past the legacy 16,384-row cap paints at its true offset or sorted
 * display slot. The planner admits detail first; the host lays each admitted
 * cell out once (cached across frames by its complete identity) into canonical
 * records, and the painter replays those records at tile-local coordinates
 * derived in Float64 (`logical - origin`), never through a global translation.
 * Pending, coarse and diagnostic frames paint a visible notice and return their
 * state; none of them paints a cell at the origin or throws for a typed axis failure.
 */
export function renderSheetViewport(
  sheet: XlsxSheet,
  ctx: CanvasRenderingContext2D,
  geometry: SheetGeometryResult | undefined,
  viewport: LogicalViewport,
  prepared?: RenderSheetAssets,
  options: ViewportPaintOptions = {},
): ViewportPaintResult {
  const plan = options.preparedPlan ?? planViewport(geometry, viewport, options.profile, { activeCell: options.activeCell, merges: options.merges, frameBudget: options.frameBudget })
  const { x: vx, y: vy, width, height } = viewport.box
  const paintable = Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0
  const owner = options.owner ?? { docId: '', sheetId: sheet.sheetId ?? sheet.name, valueGen: 0, fontGen: 0, viewGen: 0 }
  const cache = options.cache ?? new LayoutCache()
  const notice = (text: string) => {
    ctx.fillStyle = '#666666'
    ctx.font = '12px sans-serif'
    ctx.textBaseline = 'alphabetic'
    ctx.fillText(text, 8, 20)
  }
  ctx.save()
  try {
    if (paintable) {
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, width, height)
      if (options.watermark) paintWatermark(ctx, { widthPx: width, heightPx: height }, options.watermark)
    }
    if (plan.kind === 'pending' || plan.kind === 'diagnostic') {
      if (paintable) notice(VIEWPORT_NOTICE[plan.kind])
      return { kind: plan.kind, diagnostics: plan.diagnostics, records: [], coarse: [], work: emptyWork() }
    }
    const work = plan.work
    const built = cache.stats.built, reused = cache.stats.reused
    // The host lays a cell out in a box at the origin: records are box-relative.
    const host = createCellLayoutHost(ctx, sheet)
    const records = plan.kind === 'detail' ? buildCellRecords(plan.cells, owner, host, cache) : []
    const active = plan.active ? buildCellRecords([plan.active], owner, host, cache)[0] : undefined
    work.layoutsBuilt = cache.stats.built - built
    work.layoutsReused = cache.stats.reused - reused
    const coarse = plan.kind === 'detail' ? plan.coarse : plan.groups
    const paintBorders: Array<() => void> = []
    // Two passes: every fill first, then text and borders, so overflow into an
    // empty formatted neighbour is drawn over its fill (S13, as Excel does).
    const replay = (record: CellLayout, phase: 'fill' | 'content') => {
      // Float64 local offset; the canvas only ever sees tile-local coordinates.
      const lx = record.logicalBounds.x - vx, ly = record.logicalBounds.y - vy
      // Border strokes clamp to the sheet extent (paintCellLayout); expressed in the
      // record's box-local frame.
      const localExtent = { width: plan.extent.width - record.logicalBounds.x, height: plan.extent.height - record.logicalBounds.y }
      ctx.save()
      ctx.translate(lx, ly)
      const border = paintCellLayout(ctx, record.body, localExtent, phase)
      ctx.restore()
      if (border) paintBorders.push(() => { ctx.save(); ctx.translate(lx, ly); border(); ctx.restore() })
    }
    for (const group of coarse) {
      ctx.fillStyle = COARSE_FILL
      ctx.fillRect(group.bounds.x - vx, group.bounds.y - vy, group.bounds.width, group.bounds.height)
    }
    // The active cell keeps its exact record even inside a summarized region;
    // it joins the records' layer: under gridlines, borders and drawings.
    const extraActive = active && !records.some(record => record.id === active.id) ? active : undefined
    for (const record of records) replay(record, 'fill')
    if (extraActive) replay(extraActive, 'fill')
    for (const record of records) replay(record, 'content')
    if (extraActive) replay(extraActive, 'content')
    if (plan.kind === 'detail') {
      ctx.strokeStyle = '#d0d0d0'
      ctx.lineWidth = 1
      ctx.beginPath()
      // Segments overshoot the pane by one pixel so their butt caps fall outside it,
      // exactly like the static path's full-length lines.
      const yTop = Math.max(0, vy - 1), yBottom = Math.min(plan.extent.height, vy + height + 1)
      const xLeft = Math.max(0, vx - 1), xRight = Math.min(plan.extent.width, vx + width + 1)
      for (const x of plan.colLines) {
        const hidden = plan.merges.filter(merge => merge.x < x && x < merge.x + merge.width).map(merge => [merge.y, merge.y + merge.height] as [number, number])
        for (const [y1, y2] of openSegments(yTop, yBottom, hidden)) {
          work.gridlineSpans++
          ctx.moveTo(x - vx + 0.5, y1 - vy + 0.5)
          ctx.lineTo(x - vx + 0.5, y2 - vy + 0.5)
        }
      }
      for (const y of plan.rowLines) {
        const hidden = plan.merges.filter(merge => merge.y < y && y < merge.y + merge.height).map(merge => [merge.x, merge.x + merge.width] as [number, number])
        for (const [x1, x2] of openSegments(xLeft, xRight, hidden)) {
          work.gridlineSpans++
          ctx.moveTo(x1 - vx + 0.5, y - vy + 0.5)
          ctx.lineTo(x2 - vx + 0.5, y - vy + 0.5)
        }
      }
      ctx.stroke()
      for (const paint of paintBorders) paint()
      if (plan.drawings.length) {
        // Rebase each top-level placement by the Float64 pane origin (EMU) instead
        // of translating the canvas by worksheet coordinates: near the last row
        // a global translation would reach ~2.5e7 px, beyond Float32 transform
        // precision. Group children are relative to their parent and unchanged.
        const emuPerPx = EMU_PER_INCH / CSS_PX_PER_INCH, dx = vx * emuPerPx, dy = vy * emuPerPx
        paintSheetDrawings(sheet, plan.drawings.map(node => ({ ...node, xEmu: node.xEmu - dx, yEmu: node.yEmu - dy })), ctx, prepared)
      }
    }
    let pendingNotice = false
    if (plan.kind === 'detail') for (const region of plan.pending ?? []) {
      const box = region.bounds
      ctx.fillStyle = COARSE_FILL
      ctx.fillRect(box.x - vx, box.y - vy, box.width, box.height)
      if (!pendingNotice && box.x < vx + width && box.x + box.width > vx && box.y < vy + height && box.y + box.height > vy) {
        ctx.fillStyle = '#666666'
        ctx.font = '12px sans-serif'
        ctx.textBaseline = 'alphabetic'
        ctx.fillText(region.error ? 'Cells could not be loaded' : 'Loading…', Math.max(8, box.x - vx + 8), Math.max(16, box.y - vy + 16))
        pendingNotice = true
      }
    }
    // A coarse frame draws no gridlines; the active record's border still paints.
    if (plan.kind !== 'detail') for (const paint of paintBorders) paint()
    if (paintable && coarse.length) notice(plan.diagnostics.find(d => d.code === 'overflow-memory')?.message ?? COARSE_NOTICE)
    return { kind: plan.kind, ...(plan.kind === 'detail' && plan.pending?.length ? { pending: true } : {}), diagnostics: plan.diagnostics, records, coarse, ...(active ? { active } : {}), work }
  } finally { ctx.restore() }
}

/** [start,end] minus hidden intervals, in order. */
function openSegments(start: number, end: number, hidden: Array<[number, number]>): Array<[number, number]> {
  if (!(end > start)) return []
  hidden.sort((a, b) => a[0] - b[0])
  const segs: Array<[number, number]> = []
  let pos = start
  for (const [h1, h2] of hidden) {
    if (h1 > pos) segs.push([pos, Math.min(h1, end)])
    pos = Math.max(pos, h2)
    if (pos >= end) break
  }
  if (pos < end) segs.push([pos, end])
  return segs
}
