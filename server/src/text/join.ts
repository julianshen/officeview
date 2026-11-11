/**
 * R1 — the visible-text join (Stage 3a). The text index (recording pass) and
 * the occlusion evidence (serving paint pass) are separate passes with
 * different geometry; this decides which index spans are PROVABLY visible.
 * Fail closed: a span the join cannot resolve is withheld, never approximated.
 *
 * - A span's box (page px) is the bounds of `placement.visual.polygon` when
 *   present; else the bounds of its local band (x..x+width, top..bottom)
 *   projected through `placement.transform` (`placement.clip` ignored — run ink
 *   is unclipped too); else `{x, y - fontSize, width, fontSize * 1.25}`.
 * - Run ink (device px) is divided by `scale`. A run is a candidate when its ink
 *   centre lies in the span box inflated by 2 px, or the span box centre lies in
 *   the run ink.
 * - Served iff ≥ 1 candidate, every candidate servable, and every non-whitespace
 *   code point of the span is covered by the multiset of the candidates' code points.
 * - A whitespace-only span is served iff at least one nearest same-line
 *   (baselines within 1 px) non-whitespace neighbour exists, before or after it
 *   in list order, and every such neighbour is served.
 */
import { projectTextPoint, type TextSpan } from '../../../src/core/search'
import type { RunDetail } from '../occlusion'

export interface JoinResult { served: TextSpan[]; withheld: number }

interface Box { x0: number; y0: number; x1: number; y1: number }
const INFLATE = 2
const LINE_TOLERANCE = 1

function bounds(points: Iterable<{ x: number; y: number }>): Box {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
  for (const p of points) { x0 = Math.min(x0, p.x); y0 = Math.min(y0, p.y); x1 = Math.max(x1, p.x); y1 = Math.max(y1, p.y) }
  return { x0, y0, x1, y1 }
}

function spanBox(span: TextSpan): Box {
  const p = span.placement
  if (p?.visual?.polygon.length) return bounds(p.visual.polygon)
  if (p) {
    const l = p.x, r = p.x + p.width
    return bounds([[l, p.top], [r, p.top], [r, p.bottom], [l, p.bottom]].map(([x, y]) => projectTextPoint(p.transform, x!, y!)))
  }
  return { x0: span.x, y0: span.y - span.fontSize, x1: span.x + span.width, y1: span.y - span.fontSize + span.fontSize * 1.25 }
}

/** Non-finite boxes contain nothing (NaN comparisons are false). */
const contains = (b: Box, x: number, y: number) => x >= b.x0 && x <= b.x1 && y >= b.y0 && y <= b.y1
const isWhitespace = (s: string) => s.trim() === ''

function covered(text: string, candidates: RunDetail[]): boolean {
  const pool = new Map<string, number>()
  for (const r of candidates) for (const cp of r.text) pool.set(cp, (pool.get(cp) ?? 0) + 1)
  for (const cp of text) {
    if (isWhitespace(cp)) continue
    const n = pool.get(cp) ?? 0
    if (n === 0) return false
    pool.set(cp, n - 1)
  }
  return true
}

export function joinSpans(spans: TextSpan[], runs: RunDetail[], scale: number): JoinResult {
  if (!(Number.isFinite(scale) && scale > 0)) return { served: [], withheld: spans.length }
  const inks: Box[] = runs.map(({ ink }) => ({ x0: ink.x / scale, y0: ink.y / scale, x1: (ink.x + ink.width) / scale, y1: (ink.y + ink.height) / scale }))

  const verdict: Array<boolean | undefined> = spans.map((span) => {
    if (isWhitespace(span.text)) return undefined // decided by its neighbours below
    const b = spanBox(span)
    const grown: Box = { x0: b.x0 - INFLATE, y0: b.y0 - INFLATE, x1: b.x1 + INFLATE, y1: b.y1 + INFLATE }
    const cx = (b.x0 + b.x1) / 2, cy = (b.y0 + b.y1) / 2
    const candidates = runs.filter((_, i) => {
      const k = inks[i]!
      return contains(grown, (k.x0 + k.x1) / 2, (k.y0 + k.y1) / 2) || contains(k, cx, cy)
    })
    return candidates.length > 0 && candidates.every((r) => r.servable) && covered(span.text, candidates)
  })

  const sameLine = (a: TextSpan, b: TextSpan) => Math.abs(a.y - b.y) <= LINE_TOLERANCE
  const neighbour = (i: number, step: 1 | -1): number | undefined => {
    for (let j = i + step; j >= 0 && j < spans.length; j += step)
      if (verdict[j] !== undefined && sameLine(spans[i]!, spans[j]!)) return j
    return undefined
  }
  const served: TextSpan[] = []
  spans.forEach((span, i) => {
    let ok = verdict[i]
    if (ok === undefined) {
      const found = [neighbour(i, -1), neighbour(i, 1)].filter((j): j is number => j !== undefined)
      ok = found.length > 0 && found.every((j) => verdict[j] === true)
    }
    if (ok) served.push(span)
  })
  return { served, withheld: spans.length - served.length }
}
