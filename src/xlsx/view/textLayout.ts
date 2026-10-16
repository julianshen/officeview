/**
 * Canonical cell layout and text records (S5.1, spec §§11.1–11.2).
 *
 * One record per painted cell carries its source identity, Float64 logical
 * geometry, clip, fill/border/text paint data and the logical text a search,
 * copy or accessibility consumer needs. The pixel backend replays records; it
 * never recomputes formatting or text layout. Measurement and formatting stay
 * the host's (render.ts) so both paths format and shape a cell identically.
 */
import type { XlsxCell, XlsxCellStyle, XlsxMergeRange } from '../types'
import type { PptxTextBody } from '../../pptx/types'
import type { TextLayout } from '../../drawing/text-layout'
import type { CacheIdentity, ViewportCell } from './contracts'
import { MAX_COLS, type CellId, type SourceCell } from './coordinates'
import { graphemes } from '../../core/text-recording'

export interface Rect { x: number; y: number; width: number; height: number }

/** Box-relative paint data of one cell body, computed by the host in a box whose
 * top-left is (x, y). The static renderer computes it at absolute sheet
 * coordinates and paints immediately; the viewer computes it at the origin,
 * caches it, and replays it under a tile-local translation. */
export interface CellBodyLayout {
  x: number; y: number; w: number; h: number
  fill?: string
  /** Display text after host formatting (General width fitting included). */
  formattedText: string
  text?:
    | { kind: 'plain'; text: string; font: string; color: string; tx: number; ty: number; width: number
      /** Horizontal clip where blocked neighbours stop overflow (S13); absent: unclipped. */
      clipX?: { x0: number; x1: number } }
    | { kind: 'layout'; text: string; font: string; color: string; body: PptxTextBody; layout: TextLayout; clip: Rect;
      rotation?: number; stacked?: boolean; wrap: boolean; lines: number }
  borders?: XlsxCellStyle['borders']
}

/** The host owns measurement and formatting (spec §11.2: no second text model). */
export interface CellLayoutHost {
  layoutBody(cell: XlsxCell, width: number, height: number, overflow?: { left: number; right: number }, preformatted?: string): CellBodyLayout
  /** Opt-in proof that this exact body remains valid under the given allowance.
   * Only certified bodies share a key across allowances; other hosts retain
   * their complete allowance-dependent formatting/layout identity. */
  canReuseOverflow?(body: CellBodyLayout, overflow: PlacedCell['overflow']): boolean
  /** The painted text formatted ONCE under `admission` (every formatting
   * allocation is admitted before it is made; a refusal throws
   * LayoutAdmissionError): an upper bound of its UTF-16 units for a box of
   * `width`, and the formatted text itself when it does not depend on the box
   * (fed to layoutBody so a cache miss formats once). `release` returns the
   * formatting reservation once the layout is built. */
  paintedText?(cell: XlsxCell, width: number, admission: LayoutAdmission): { units: number; text?: string; release(): void }
}

export interface FillRecord { color: string | null }
export interface BorderRecord { side: 'left' | 'right' | 'top' | 'bottom'; style: string }
/** One shaped run: its UTF-16 source interval, grapheme boundaries (absolute
 * code-unit offsets, lone surrogates preserved) and transform flags. */
export interface ShapedRun {
  text: string; start: number; end: number
  font: string; color: string
  graphemeBoundaries: readonly number[]
  rotation?: number; stacked?: boolean; wrap?: boolean; lines?: number
}
export interface TextRecord {
  cellId: CellId; source: SourceCell
  rawText: string; formattedText: string
  logicalBounds: Rect; clip: Rect
  runs: readonly ShapedRun[]
  valueGen: number; fontGen: number
}
export interface CellLayout {
  id: CellId; source: SourceCell; owner: CacheIdentity
  /** Float64 logical CSS geometry; merge anchors span their whole range. */
  logicalBounds: Rect; clip: Rect
  fill: FillRecord; borders: readonly BorderRecord[]
  text: readonly TextRecord[]
  mergeOwner?: CellId
  drawingRefs?: readonly string[]
  /** Box-relative replay data for the Canvas2D (or any) pixel backend. */
  body: CellBodyLayout
}

/** A cell the planner admitted, in logical CSS geometry. */
export interface PlacedCell { cell: XlsxCell; source: SourceCell; x: number; y: number; width: number; height: number; merge?: XlsxMergeRange
  /** Room the text may use past the box into empty neighbours (S13); a layout
   * key unless the host certifies an identical body under another allowance. */
  overflow?: { left: number; right: number } }

export function cellIdOf(source: SourceCell): CellId {
  return source.row * MAX_COLS + source.col
}

export interface LayoutCacheStats { built: number; reused: number; invalidations: number; entries: number; bytes: number; evicted: number }

/** Formatted values and text layouts across frames (plan S5.1). The scope is the
 * document, sheet and font generation: a font publication (`fontGen`) or another
 * document/sheet drops every layout. Within a scope an entry is keyed by the cell
 * identity, everything the host formats/shapes from (value, error/formula flags,
 * resolved style) and the box dimensions, so a value or geometry change rebuilds
 * only the records it touches. Zoom/DPR are not keys: layout is in logical CSS
 * pixels and the caller applies the scale. LRU bounded by entry count and by
 * estimated retained bytes (text, runs, laid-out lines), which are charged to the
 * §12.4 main-thread pool: the default 8 MiB is an eighth of desktop's 64 MiB. */
/** Repaint memory admitted per record at construction (spec §§5.6, 12.4).
 * `admit` reserves bytes against the main budget in O(1) (false: they do not
 * fit); `release` returns a reservation; `hold` moves bytes the batch already
 * owns into the reservation without a check (net-zero transfers); `cacheChanged`
 * reports the cache's retained-byte delta so its ledger charge follows at once. */
export interface LayoutAdmission {
  admit(bytes: number): boolean
  release(bytes: number): void
  hold(bytes: number): void
  cacheChanged(delta: number): void
}
/** A record could not be admitted even after evicting the cache: the caller
 * paints the tile's declared coarse/notice path instead. */
export class LayoutAdmissionError extends Error {
  constructor() { super('Not enough memory to lay out this text; zoom in or close other views.'); this.name = 'LayoutAdmissionError' }
}
/** Upper bound of estimateLayoutBytes for a record whose painted text has at
 * most `units` UTF-16 units: per unit at most one line and one segment (each
 * with its object, transform, runs and boundaries) plus the record's own
 * strings; the fixed entry/body overheads; and the key. */
export function layoutUpperBytes(units: number, keyLength: number): number {
  return 1536 + 700 * units + 2 * keyLength
}

export class LayoutCache {
  readonly stats: LayoutCacheStats = { built: 0, reused: 0, invalidations: 0, entries: 0, bytes: 0, evicted: 0 }
  private scope: string | undefined
  private readonly entries = new Map<string, CellBodyLayout>()
  private readonly sizes = new Map<string, number>()
  /** The current paint batch: keys it built that the cache retains, and the
   * bytes it holds for bodies it keeps alive outside the cache. */
  private batch: { admission: LayoutAdmission; keys: Set<string>; held: number } | null = null
  constructor(readonly maxEntries = 65_536, readonly maxBytes = DEFAULT_LAYOUT_CACHE_BYTES) {}
  /** Starts a paint batch: until endBatch, every body built for it is admitted
   * before construction and stays accounted while the batch holds it. */
  beginBatch(admission: LayoutAdmission): void { this.endBatch(); this.batch = { admission, keys: new Set(), held: 0 } }
  endBatch(): void {
    const batch = this.batch
    this.batch = null
    if (batch && batch.held) batch.admission.release(batch.held)
  }
  private evictOldest(): boolean {
    const oldest = this.entries.keys().next().value
    if (oldest === undefined) return false
    const bytes = this.sizes.get(oldest) ?? 0
    this.entries.delete(oldest)
    this.sizes.delete(oldest)
    this.stats.bytes -= bytes
    this.stats.evicted++
    const batch = this.batch
    if (batch) {
      batch.admission.cacheChanged(-bytes)
      // A body this batch still paints leaves the cache, not memory.
      if (batch.keys.delete(oldest)) { batch.admission.hold(bytes); batch.held += bytes }
    }
    return true
  }
  layout(owner: CacheIdentity, placed: PlacedCell, host: CellLayoutHost): CellBodyLayout {
    const scope = `${owner.docId}\u0000${owner.sheetId}\u0000${owner.fontGen}`
    if (scope !== this.scope) {
      if (this.scope !== undefined) this.stats.invalidations++
      if (this.batch && this.stats.bytes) this.batch.admission.cacheChanged(-this.stats.bytes)
      this.entries.clear()
      this.sizes.clear()
      this.stats.bytes = 0
      this.scope = scope
    }
    const commonKey = layoutKey(placed)
    // Keep one budgeted entry, not aliases for every tile's capped allowance.
    const sharedKey = host.canReuseOverflow ? `shared|${commonKey}` : undefined
    const shared = sharedKey === undefined ? undefined : this.entries.get(sharedKey)
    let key: string
    if (shared && host.canReuseOverflow!(shared, placed.overflow)) key = sharedKey!
    else {
      const overflow = placed.overflow
      key = `${overflow ? `${overflow.left}:${overflow.right}` : ''}|${commonKey}`
    }
    const hit = this.entries.get(key)
    if (hit) {
      this.entries.delete(key)
      this.entries.set(key, hit)
      this.stats.reused++
      return hit
    }
    const batch = this.batch
    let upper = 0
    let painted: ReturnType<NonNullable<CellLayoutHost['paintedText']>> | undefined
    if (batch) {
      // Admission BEFORE construction: the painted text, formatted once under
      // the batch's allocation observer, then the record's bound, evicting
      // least recently used entries until it fits.
      painted = host.paintedText?.(placed.cell, placed.width, batch.admission)
      upper = layoutUpperBytes(painted ? painted.units : rawTextOf(placed.cell.value).length, key.length + (sharedKey?.length ?? 0))
      try { while (!batch.admission.admit(upper)) if (!this.evictOldest()) throw new LayoutAdmissionError() }
      catch (error) { painted?.release(); throw error }
    }
    let built: CellBodyLayout
    try { built = host.layoutBody(placed.cell, placed.width, placed.height, placed.overflow, painted?.text) }
    catch (error) { if (batch) batch.admission.release(upper); throw error }
    finally { painted?.release() }
    this.stats.built++
    if (sharedKey !== undefined && host.canReuseOverflow!(built, placed.overflow)) key = sharedKey
    const bytes = estimateLayoutBytes(key, built)
    const replaced = this.sizes.get(key)
    if (replaced !== undefined) {
      this.entries.delete(key); this.sizes.delete(key); this.stats.bytes -= replaced
      if (batch) { batch.admission.cacheChanged(-replaced); batch.keys.delete(key) }
    }
    if (batch) {
      // The construction reservation becomes the body's own accounting: held
      // by the batch when it is not retained, by the cache charge when it is.
      if (bytes <= upper) batch.admission.release(upper - bytes)
      else batch.admission.hold(bytes - upper)
    }
    // The byte budget is a real bound: a record larger than the whole budget is
    // painted uncached (returned, never retained), and least recently used
    // entries are evicted BEFORE an insertion, so retained bytes never exceed it.
    if (bytes > this.maxBytes) {
      if (batch) batch.held += bytes
      this.stats.entries = this.entries.size
      return built
    }
    while (this.entries.size && (this.entries.size + 1 > this.maxEntries || this.stats.bytes + bytes > this.maxBytes)) this.evictOldest()
    this.entries.set(key, built)
    this.stats.bytes += bytes
    this.sizes.set(key, bytes)
    this.stats.entries = this.entries.size
    if (batch) {
      batch.admission.release(bytes)
      batch.admission.cacheChanged(bytes)
      batch.keys.add(key)
    }
    return built
  }
}

export const DEFAULT_LAYOUT_CACHE_BYTES = 8 * 1024 * 1024
/** Conservative retained-size estimate: UTF-16 strings at 2 bytes per unit, the
 * record's and each laid-out segment's grapheme-boundary arrays (charged 16 bytes per entry: slot plus boxed
 * number in the worst case), a 128-byte fixed entry/map overhead, and object
 * overheads for the body, each laid-out line and segment. */
export function estimateLayoutBytes(key: string, body: CellBodyLayout): number {
  let bytes = 128 + 256 + 2 * key.length + 2 * body.formattedText.length
  const text = body.text
  if (text) {
    bytes += 128 + 2 * (text.text.length + text.font.length) + 16 * (text.text.length + 1)
    if (text.kind === 'layout') {
      bytes += 512 + 2 * text.text.length // body/paragraph copy of the source text
      for (const line of text.layout.lines) {
        bytes += 160 + 8 * line.segments.length
        // Each segment holds its own transform, source-run list and grapheme
        // boundaries (charged like the record's: 16 bytes per entry, plus slack).
        for (const segment of line.segments) bytes += 384 + 2 * segment.text.length + 24 * (segment.text.length + 1)
      }
    }
  }
  return bytes
}

const styleKeys = new WeakMap<XlsxCellStyle, string>()
function styleKey(style: XlsxCellStyle | undefined): string {
  if (!style) return ''
  let key = styleKeys.get(style)
  if (key === undefined) styleKeys.set(style, key = JSON.stringify(style))
  return key
}
function layoutKey({ cell, source, width, height }: PlacedCell): string {
  const value = cell.value === null ? 'n' : `${typeof cell.value}:${String(cell.value)}`
  return `${source.row},${source.col}|${width}|${height}|${cell.valueIsError === undefined ? '' : cell.valueIsError ? 1 : 0}${cell.formula === undefined ? '' : 'f'}|${styleKey(cell.style)}|${value}`
}

/** Raw logical text of a value: what search/copy see before display formatting. */
export function rawTextOf(value: XlsxCell['value']): string {
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return String(value)
}

/** UTF-16 grapheme boundaries including both ends; lone surrogates stay their
 * own cluster (no replacement decoding). */
export function graphemeBoundaries(text: string): number[] {
  const out = graphemes(text).map(cluster => cluster.start)
  out.push(text.length)
  return out
}

const runCache = new WeakMap<CellBodyLayout, ShapedRun>()
function shapedRun(body: CellBodyLayout): ShapedRun | undefined {
  const text = body.text
  if (!text) return undefined
  let run = runCache.get(body)
  if (!run) {
    run = { text: text.text, start: 0, end: text.text.length, font: text.font, color: text.color, graphemeBoundaries: graphemeBoundaries(text.text) }
    if (text.kind === 'layout') Object.assign(run, text.rotation !== undefined ? { rotation: text.rotation } : {}, text.stacked ? { stacked: true } : {}, { wrap: text.wrap, lines: text.lines })
    runCache.set(body, run)
  }
  return run
}

const SIDES = ['left', 'right', 'top', 'bottom'] as const

/** Canonical records for admitted cells in planner order. Geometry stays Float64
 * logical CSS; the body is box-relative and replayed by the pixel backend. */
export function buildCellRecords(placed: readonly PlacedCell[], owner: CacheIdentity, host: CellLayoutHost, cache: LayoutCache): CellLayout[] {
  const records: CellLayout[] = []
  for (const item of placed) {
    const body = cache.layout(owner, item, host)
    const id = cellIdOf(item.source)
    const logicalBounds = { x: item.x, y: item.y, width: item.width, height: item.height }
    const text: TextRecord[] = []
    const run = shapedRun(body)
    if (run && body.text) {
      // The record's clip is the painted clip (S13: overflow widens or cuts it),
      // translated from the box-local body into logical coordinates.
      const clip = body.text.kind === 'layout'
        ? { x: item.x + (body.text.clip.x - body.x), y: item.y + (body.text.clip.y - body.y), width: body.text.clip.width, height: body.text.clip.height }
        // Plain text overflows like the static renderer: its own ink band, cut
        // where a blocked neighbour stops the overflow (S13).
        : plainClip(item.x - body.x, body.text, item.y, item.height)
      text.push({ cellId: id, source: item.source, rawText: rawTextOf(item.cell.value), formattedText: body.formattedText, logicalBounds, clip, runs: [run], valueGen: owner.valueGen, fontGen: owner.fontGen })
    }
    const borders: BorderRecord[] = []
    for (const side of SIDES) { const style = body.borders?.[side]; if (style) borders.push({ side, style }) }
    records.push({
      id, source: item.source, owner, logicalBounds, clip: logicalBounds,
      fill: { color: body.fill ?? null }, borders, text,
      ...(item.merge ? { mergeOwner: cellIdOf({ row: item.merge.minRow, col: item.merge.minCol }) } : {}),
      body,
    })
  }
  return records
}

function plainClip(dx: number, text: { tx: number; width: number; clipX?: { x0: number; x1: number } }, y: number, height: number) {
  const x0 = text.clipX ? Math.max(text.tx, text.clipX.x0) : text.tx
  const x1 = text.clipX ? Math.min(text.tx + text.width, text.clipX.x1) : text.tx + text.width
  return { x: dx + x0, y, width: Math.max(0, x1 - x0), height }
}

/** Store-path detail cells as host cells at their logical boxes: resolved and
 * permanently unavailable values only (unknown/error records stay placeholders,
 * never text). `styleOf` resolves a style index when the caller has styles. */
export function placedFromViewportCells(cells: readonly ViewportCell[], styleOf?: (styleIndex: number) => XlsxCellStyle | undefined): PlacedCell[] {
  const out: PlacedCell[] = []
  for (const item of cells) {
    if (item.kind !== 'cell' && item.kind !== 'unavailable') continue
    const source = item.cell
    const style = styleOf?.(source.styleIndex)
    const cell: XlsxCell = {
      ref: '', row: item.source.row, col: item.source.col, value: item.kind === 'unavailable' ? null : source.value, styleIndex: source.styleIndex,
      ...(source.hasCachedValue !== undefined ? { hasCachedValue: source.hasCachedValue } : {}),
      ...(source.valueIsError !== undefined ? { valueIsError: source.valueIsError } : {}),
      ...(source.formula !== undefined ? { formula: source.formula } : {}),
      ...(style ? { style } : {}),
    }
    out.push({ cell, source: item.source, x: item.bounds.x, y: item.bounds.y, width: item.bounds.width, height: item.bounds.height })
  }
  return out
}
