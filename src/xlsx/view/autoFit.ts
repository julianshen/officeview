/**
 * The bounded auto-fit reader (spec §5.6 amendment), React-free so the stream server's pool worker runs the same
 * engine as the local viewer (Stage 4b P9). Moved verbatim from `src/components/xlsx/sheetResize.tsx`, which
 * re-exports it; local imports are unchanged.
 */
import type { LogicalCellBox } from './contracts'
import { MAX_ROWS } from './coordinates'
import type { SheetGeometry } from './viewport'
import { AUTOFIT_MAX_CELLS, AUTOFIT_MAX_ROWS, RESIZE_LIMITS, autoFitRowCandidates, cellFont, clampSize, fitColumnWidth, fitRowHeight, type Measure } from './axisSizes'
import { createCellLayoutHost } from './renderViewport'
import { createAdmittedLayoutContext } from '../../components/xlsx/measureContext'
import type { CanvasLike } from '../../core/canvas-backend'
import { TILE_CANVAS_BYTES, shapeCacheBytes } from './mainLedger'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../types'
import { cellDisplayText } from '../render'
import type { FormatAllocation } from '../formula/format'


type FitCell = Pick<XlsxCell, 'col' | 'value'> & { style?: XlsxCellStyle; styleIndex?: number }
export interface AutoFitInput {
  geometry: SheetGeometry
  sheet: XlsxSheet
  /** The pane in logical CSS px (visible rows/columns come first). */
  pane: LogicalCellBox
  /** Disjoint visible regions in priority order (pinned strips before body). */
  regions?: readonly LogicalCellBox[]
  /** Rows the source already holds on the main thread, never a read. */
  knownRows: (limit: number) => Iterable<number>
  styleOf?: (styleIndex: number) => XlsxCellStyle | undefined
  measure: Measure
  /** Formatting allocation observer (FitMeasure.allocation): every builder admitted first. */
  allocation?: FormatAllocation
  /** Canonical painter height of a wrapped or rotated body (FitMeasure.layoutHeight). */
  layoutHeight?: (cell: XlsxCell, widthCss: number) => number | undefined
}
function cellIn(cells: readonly FitCell[] | undefined, col: number): FitCell | undefined {
  if (!cells) return undefined
  let low = 0, high = cells.length
  while (low < high) {
    const mid = (low + high) >>> 1
    if (cells[mid].col < col) low = mid + 1
    else high = mid
  }
  return cells[low]?.col === col ? cells[low] : cells.find(cell => cell.col === col)
}
/** The bounded auto-fit size (axisSizes AUTOFIT_MAX_ROWS / AUTOFIT_MAX_CELLS):
 * a column fits its widest formatted value among the visible and known rows; a
 * row fits its tallest populated cell. Merged cells spanning the resized
 * direction are skipped, as Excel does. */
export function autoFitSize(kind: 'row' | 'column', index: number, input: AutoFitInput): number {
  const { geometry, sheet, pane, measure } = input
  const regions = input.regions ?? [pane]
  const metrics = geometry.metrics
  const styleOf = (cell: FitCell) => cell.style ?? (cell.styleIndex !== undefined ? input.styleOf?.(cell.styleIndex) : undefined)
  const text = (cell: FitCell, style: XlsxCellStyle | undefined) => cellDisplayText(sheet, { value: cell.value, style }, input.allocation)
  if (kind === 'column') {
    const visible = function* () {
      for (const region of regions) for (const run of metrics.rows.visibleBetween(region.y, region.y + region.height).runs) for (let slot = run.start; slot <= run.end; slot++) yield metrics.order.displayToSource(slot)
    }
    const rows = autoFitRowCandidates(visible(), input.knownRows(4 * AUTOFIT_MAX_ROWS), row => row >= 0 && row < MAX_ROWS && metrics.sourceToVisibleRank(row) !== null)
    const cells = function* () {
      for (const row of rows) {
        const cell = cellIn(geometry.cellsOf(row) as readonly FitCell[] | undefined, index)
        if (!cell || geometry.merges.firstWhere({ r0: row, r1: row, c0: index, c1: index }, range => range.minCol !== range.maxCol)) continue
        const style = styleOf(cell), shown = text(cell, style), font = cellFont(style).font
        // Excel parity: a wrapping cell never asks for its unwrapped width; it
        // contributes its longest unbreakable word (its own line breaks stay).
        if (style?.wrapText === true) { for (const word of shown.split(/\s+/)) if (word) yield { text: word, font } }
        else yield { text: shown, font }
      }
    }
    return clampSize('col', fitColumnWidth(cells(), measure))
  }
  const row = index
  const populated = (geometry.cellsOf(row) ?? []) as readonly FitCell[]
  const inPane = (cell: FitCell) => { const x = metrics.cols.start(cell.col); return regions.some(region => x < region.x + region.width && x + metrics.cols.size(cell.col) > region.x) }
  const ordered = [...populated.filter(inPane), ...populated.filter(cell => !inPane(cell))].slice(0, AUTOFIT_MAX_CELLS)
  const cells = function* () {
    for (const cell of ordered) {
      const widthCss = metrics.cols.size(cell.col)
      if (widthCss <= 0 || geometry.merges.firstWhere({ r0: row, r1: row, c0: cell.col, c1: cell.col }, range => range.minRow !== range.maxRow)) continue
      const style = styleOf(cell), { font, sizePx } = cellFont(style), wrap = style?.wrapText === true
      const heightPx = wrap || style?.textRotation ? input.layoutHeight?.({ ref: '', row, col: cell.col, value: cell.value, styleIndex: cell.styleIndex ?? 0, ...(style ? { style } : {}) }, widthCss) : undefined
      yield { text: text(cell, style), font, sizePx, wrap, widthCss, ...(heightPx !== undefined ? { heightPx } : {}) }
    }
  }
  return clampSize('row', fitRowHeight(cells(), measure))
}

/** An auto-fit was refused because its measurement scratch would exceed its
 * admitted memory (the caller announces it; nothing is resized). */
export class AutoFitRefusedError extends Error {
  constructor() { super('Auto-fit needs more memory than is available; resize by dragging or with the keyboard instead.'); this.name = 'AutoFitRefusedError' }
}
/** Retained JS bytes per measured (font, text) cache entry beyond its UTF-16 payload. */
const MEASURE_ENTRY_BYTES = 96
/** A canonical body layout's transient JS bytes beyond its text (§11.2 records). */
const LAYOUT_HEAP_BYTES = 512
export interface FitCharge { heap: number; native: number }
export interface FitMeasure {
  measure: Measure
  /** Height of a wrapped or rotated cell body from the canonical painter
   * layout, as its transformed (rotated) extent. */
  layoutHeight?: (cell: XlsxCell, widthCss: number) => number | undefined
  /** Formatting observer: each builder is admitted (heap) before it is made. */
  allocation: FormatAllocation
  bytes(): number
}
/** One auto-fit's measurement owner: a 1×1 scratch canvas whose 2D context and
 * caches live only for that fit (the caller drops them after it). Costs are
 * reserved BEFORE they exist, with the calibrated §12.4 model: the context
 * (TILE_CANVAS_BYTES) before it is created, the browser shaping cache
 * (shapeCacheBytes per font) before each new text is measured, and the JS
 * cache entries. `admit({ heap, native })` receives the running totals and
 * charges them; a refusal throws AutoFitRefusedError and nothing more is
 * created. With an injected `measureText` no context is created or charged.
 * `createCanvas` supplies the admitted pair's canvases where there is no
 * document (the stream server's canvas backend); everything else, the
 * canonical `layoutHeight` included, is unchanged. */
export function createFitMeasure(options: { admit: (charge: FitCharge) => boolean; measureText?: Measure; sheet?: XlsxSheet; createCanvas?: () => CanvasLike }): FitMeasure {
  let ctx: CanvasRenderingContext2D | null | undefined = options.measureText ? null : undefined
  let host: ReturnType<typeof createCellLayoutHost> | null = null
  const fonts = new Map<string, { entries: number; units: number }>()
  const cache = new Map<string, number>()
  let heap = 0, native = 0
  const reserve = (moreHeap: number, moreNative: number) => {
    if (!options.admit({ heap: heap + moreHeap, native: native + moreNative })) throw new AutoFitRefusedError()
    heap += moreHeap; native += moreNative
  }
  /** Reserves one new shaped text in `font`'s cache before the browser shapes it. */
  const shaped = new Map<string, Set<string>>()
  /** `cache` names the context whose shaping cache grows (ours, or the
   * canonical layout's local measurement context). */
  const reserveShape = (font: string, text: string, cacheName = '') => {
    font = `${cacheName}\u0000${font}`
    const texts = shaped.get(font)
    if (texts?.has(text)) return
    const current = fonts.get(font) ?? { entries: 0, units: 0 }
    const next = { entries: current.entries + 1, units: current.units + text.length }
    reserve(font.length * 2 + text.length * 2 + MEASURE_ENTRY_BYTES, shapeCacheBytes(next.entries, next.units) - shapeCacheBytes(current.entries, current.units))
    fonts.set(font, next)
    if (texts) texts.add(text)
    else shaped.set(font, new Set([text]))
  }
  const context = (): CanvasRenderingContext2D | null => {
    if (ctx !== undefined) return ctx
    // Both admitted contexts (the base and the canonical layout's local
    // identity context) are reserved before they exist; each measures only
    // through its own admitting wrapper (no global state is touched).
    reserve(0, 2 * TILE_CANVAS_BYTES)
    ctx = createAdmittedLayoutContext((cacheName, font, text) => reserveShape(font, text, cacheName), options.createCanvas)
    return ctx
  }
  // Formatting builders are admitted (heap) before they are made; the fit's
  // owner releases everything when the fit ends.
  const allocation: FormatAllocation = { allocate: bytes => reserve(bytes, 0), release: () => {} }
  const measure: Measure = (font, text) => {
    const key = `${font}\n${text}`
    let width = cache.get(key)
    if (width !== undefined) return width
    const c = context()
    reserve(key.length * 2 + MEASURE_ENTRY_BYTES, 0)
    if (c) {
      c.font = font
      width = c.measureText(text).width
    } else width = options.measureText ? options.measureText(font, text) : text.length * (Number(/([\d.]+)pt/.exec(font)?.[1] ?? 10) * 96 / 72) * 0.55
    cache.set(key, width)
    return width
  }
  const layoutHeight = options.sheet && !options.measureText ? (cell: XlsxCell, widthCss: number): number | undefined => {
    const c = context()
    if (!c) return undefined
    // As paint does: text that cannot change with the width feeds the layout,
    // so only a General number is formatted a second time.
    let widthDependent = false
    const text = cellDisplayText(options.sheet!, cell, allocation, () => { widthDependent = true })
    reserve(text.length * 4 + LAYOUT_HEAP_BYTES, 0)
    host ??= createCellLayoutHost(c, options.sheet!)
    const body = host.layoutBody(cell, widthCss, RESIZE_LIMITS.row.max, undefined, widthDependent ? undefined : text, allocation)
    if (body.text?.kind !== 'layout') return undefined
    return body.text.rotation ? rotatedExtent(body.text.layout, c) : body.text.layout.height
  } : undefined
  return { measure, ...(layoutHeight ? { layoutHeight } : {}), allocation, bytes: () => heap + native }
}
/** Vertical extent of a laid-out body whose segments carry rotated baseline
 * transforms (render.ts rotatedLayout), from the canonical glyph band of each
 * segment's font: measured 'Mg' ascent/descent, at least .85em/.25em
 * (render.ts cellTextMeasurer / textLayoutBounds), for any angle. */
function rotatedExtent(layout: { lines: ReadonlyArray<{ baseline: number; segments: ReadonlyArray<{ text: string; x: number; width: number; style: { fontSizePt?: number; bold?: boolean; italic?: boolean }; transform?: { a: number; b: number; c: number; d: number; e: number; f: number } }> }> }, ctx: CanvasRenderingContext2D): number {
  const bands = new Map<string, { ascent: number; descent: number }>()
  const band = (style: { fontSizePt?: number; bold?: boolean; italic?: boolean }) => {
    const { font, sizePx } = cellFont(style)
    let value = bands.get(font)
    if (!value) {
      ctx.font = font
      const m = ctx.measureText('Mg')
      bands.set(font, value = { ascent: Math.max(m.actualBoundingBoxAscent ?? 0, sizePx * .85), descent: Math.max(m.actualBoundingBoxDescent ?? 0, sizePx * .25) })
    }
    return value
  }
  let top = Infinity, bottom = -Infinity
  for (const line of layout.lines) for (const segment of line.segments) {
    if (segment.text === '\n') continue
    const t = segment.transform ?? { a: 1, b: 0, c: 0, d: 1, e: segment.x, f: line.baseline }
    const { ascent, descent } = band(segment.style)
    for (const [x, y] of [[0, -ascent], [segment.width, -ascent], [segment.width, descent], [0, descent]]) {
      const py = t.b * x + t.d * y + t.f
      top = Math.min(top, py); bottom = Math.max(bottom, py)
    }
  }
  return bottom > top ? bottom - top : 0
}

