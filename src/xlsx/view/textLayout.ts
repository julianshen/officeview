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
import type { CacheIdentity } from './contracts'
import { MAX_COLS, type CellId, type SourceCell } from './coordinates'

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
    | { kind: 'plain'; text: string; font: string; color: string; tx: number; ty: number; width: number }
    | { kind: 'layout'; text: string; font: string; color: string; body: PptxTextBody; layout: TextLayout; clip: Rect;
      rotation?: number; stacked?: boolean; wrap: boolean; lines: number }
  borders?: XlsxCellStyle['borders']
}

/** The host owns measurement and formatting (spec §11.2: no second text model). */
export interface CellLayoutHost {
  layoutBody(cell: XlsxCell, width: number, height: number): CellBodyLayout
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
export interface PlacedCell { cell: XlsxCell; source: SourceCell; x: number; y: number; width: number; height: number; merge?: XlsxMergeRange }

export function cellIdOf(source: SourceCell): CellId {
  return source.row * MAX_COLS + source.col
}

export interface LayoutCacheStats { built: number; reused: number; invalidations: number; entries: number }

/** Formatted-value/text-layout cache across frames (plan S5.1). Deliberately
 * wrong boundary: never reuses. */
export class LayoutCache {
  readonly stats: LayoutCacheStats = { built: 0, reused: 0, invalidations: 0, entries: 0 }
  constructor(readonly maxEntries = 65_536) {}
  layout(_owner: CacheIdentity, placed: PlacedCell, host: CellLayoutHost): CellBodyLayout {
    this.stats.built++
    return host.layoutBody(placed.cell, placed.width, placed.height)
  }
}

/** Canonical records for admitted cells. Deliberately wrong boundary: empty. */
export function buildCellRecords(_placed: readonly PlacedCell[], _owner: CacheIdentity, _host: CellLayoutHost, _cache: LayoutCache): CellLayout[] {
  return []
}
