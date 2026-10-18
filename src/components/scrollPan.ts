/**
 * Native scroll and transform pan as one coordinate model (spec §6.11, R2-5).
 * Eager and pure: the zoom handlers in <OfficeDoc> and the lazy page tracker
 * both map the viewport through these, so the retained scrollTop of an
 * overflow: hidden container can never skew tracking or goTo.
 */
import type { PageSpec } from '../render/paint'

export interface ScrollPanState { scrollTop: number; panY: number; zoom: number; padding: number }
export interface PageOffsets { top: number[]; height: number[] }

/** Content y (current layout px, before the zoom transform) under viewport y. */
export function contentYAt(viewY: number, s: ScrollPanState): number {
  return (viewY + s.scrollTop - s.padding - s.panY) / s.zoom
}

/** The panY that puts content y at viewport y = 0: the inverse of contentYAt. */
export function panYForContentTop(contentY: number, s: ScrollPanState): number {
  return s.scrollTop - s.padding - contentY * s.zoom
}

/** Page tops and heights at displayWidth; gaps are fixed px. */
export function pageOffsets(specs: readonly PageSpec[], displayWidth: number, gap: number): PageOffsets {
  const top: number[] = []
  const height: number[] = []
  let y = 0
  for (const s of specs) {
    const h = s.heightPx * displayWidth / s.widthPx
    top.push(y)
    height.push(h)
    y += h + gap
  }
  return { top, height }
}

/** A content point in one layout mapped into another by page index plus
 * page-local offset (scaled); gap offsets stay unscaled. */
export function remapContentY(y: number, from: PageOffsets, to: PageOffsets): number {
  const n = from.top.length
  if (n === 0 || y < from.top[0]) return y
  let i = n - 1
  while (i > 0 && y < from.top[i]) i--
  const local = y - from.top[i]
  return local <= from.height[i]
    ? to.top[i] + local * (to.height[i] / from.height[i])
    : to.top[i] + to.height[i] + local - from.height[i]
}

/** Native scroll → pan at nextZoom, keeping contentYAt(anchorY) fixed. */
export function enterZoom(s: ScrollPanState, nextZoom: number, anchorY: number): { panY: number; scrollTop: 0 } {
  return { panY: anchorY - s.padding - contentYAt(anchorY, s) * nextZoom, scrollTop: 0 }
}

/** Pan → native scroll at zoom 1, keeping contentYAt(anchorY) fixed (clamped). */
export function exitZoom(s: ScrollPanState, anchorY: number, maxScrollTop: number): { panY: 0; scrollTop: number } {
  const top = contentYAt(anchorY, s) + s.padding - anchorY
  return { panY: 0, scrollTop: Math.min(Math.max(0, top), Math.max(0, maxScrollTop)) }
}
