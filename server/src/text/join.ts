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
 * - Served iff ≥ 1 candidate, every candidate servable, and ordered
 *   correspondence: with whitespace removed, the span's code points T are a
 *   contiguous substring of the candidates' texts concatenated in paint (`seq`)
 *   order, and the candidates' unconsumed budget is ≥ |T|. A run's budget starts
 *   at its stripped code-point length; a served span consumes |T| from its
 *   candidates in seq order. Spans are processed in input order, so one painted
 *   run never authorises more text than it painted (no anagrams, no reuse).
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

/** Code points with whitespace removed (no normalisation). */
const stripped = (s: string) => Array.from(s).filter((cp) => !isWhitespace(cp))

/** Contiguous code-point substring test. */
function containsRun(hay: string[], needle: string[]): boolean {
  outer: for (let i = 0; i + needle.length <= hay.length; i++) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer
    return true
  }
  return false
}

export function joinSpans(spans: TextSpan[], runs: RunDetail[], scale: number): JoinResult {
  if (!(Number.isFinite(scale) && scale > 0)) return { served: [], withheld: spans.length }
  const budget = runs.map((r) => stripped(r.text).length)
  const inks: Box[] = runs.map(({ ink }) => ({ x0: ink.x / scale, y0: ink.y / scale, x1: (ink.x + ink.width) / scale, y1: (ink.y + ink.height) / scale }))

  const verdict: Array<boolean | undefined> = spans.map((span) => {
    if (isWhitespace(span.text)) return undefined // decided by its neighbours below
    const b = spanBox(span)
    const grown: Box = { x0: b.x0 - INFLATE, y0: b.y0 - INFLATE, x1: b.x1 + INFLATE, y1: b.y1 + INFLATE }
    const cx = (b.x0 + b.x1) / 2, cy = (b.y0 + b.y1) / 2
    const candidates = runs.flatMap((_, i) => {
      const k = inks[i]!
      return contains(grown, (k.x0 + k.x1) / 2, (k.y0 + k.y1) / 2) || contains(k, cx, cy) ? [i] : []
    }).sort((a, b) => runs[a]!.seq - runs[b]!.seq)
    if (!candidates.length || !candidates.every((i) => runs[i]!.servable)) return false
    const t = stripped(span.text)
    const left = candidates.reduce((n, i) => n + budget[i]!, 0)
    if (left < t.length || !containsRun(candidates.flatMap((i) => stripped(runs[i]!.text)), t)) return false
    let need = t.length
    for (const i of candidates) { const take = Math.min(need, budget[i]!); budget[i] = budget[i]! - take; need -= take }
    return true
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
