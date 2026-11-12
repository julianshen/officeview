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
 * - Served iff ≥ 1 candidate, every candidate servable, and positional
 *   correspondence: with whitespace removed, the span's code points T occur
 *   contiguously in the candidates' code points concatenated in paint (`seq`)
 *   order, at an occurrence whose positions are all unconsumed (the first such,
 *   left to right). Serving marks exactly those positions consumed. Spans are
 *   processed in input order, so each painted code point authorises at most one
 *   span's code point (no anagrams, no reuse).
 * - A whitespace-only span is served iff at least one nearest same-line
 *   (baselines within 1 px) non-whitespace neighbour exists, before or after it
 *   in list order, and every such neighbour is served.
 *
 * Diagnostic (Task 9, counts only, never sent on the wire): an optional
 * `causes` tally receives one count per withheld span — `noCandidate` (no run
 * in reach: e.g. glyph-path text), `occluded[reason]` (a candidate run is not
 * servable; the first such in paint order names the reason), `mismatch`
 * (candidates found, characters do not correspond), `whitespace` (a blank span
 * whose neighbours are missing or withheld), `badScale` (all spans, scale
 * not finite and positive). It does not change any verdict.
 */
import { projectTextPoint, type TextSpan } from '../../../src/core/search'
import type { OcclusionReason, RunDetail } from '../occlusion'

export interface JoinResult { served: TextSpan[]; withheld: number }
/** Withheld spans by cause (see the module comment); sums to `withheld`. */
export interface WithheldCauses { noCandidate: number; occluded: Partial<Record<OcclusionReason | 'unspecified', number>>; mismatch: number; whitespace: number; badScale: number }
export const emptyCauses = (): WithheldCauses => ({ noCandidate: 0, occluded: {}, mismatch: 0, whitespace: 0, badScale: 0 })

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

/** First start of `needle` in `hay` whose positions are all unconsumed, or -1. */
function freeOccurrence(hay: Array<{ cp: string; run: number; at: number }>, needle: string[], consumed: boolean[][]): number {
  outer: for (let i = 0; i + needle.length <= hay.length; i++) {
    for (let j = 0; j < needle.length; j++) {
      const h = hay[i + j]!
      if (h.cp !== needle[j] || consumed[h.run]![h.at]) continue outer
    }
    return i
  }
  return -1
}

export function joinSpans(spans: TextSpan[], runs: RunDetail[], scale: number, causes?: WithheldCauses): JoinResult {
  if (!(Number.isFinite(scale) && scale > 0)) { if (causes) causes.badScale += spans.length; return { served: [], withheld: spans.length } }
  const runCps = runs.map((r) => stripped(r.text))
  const consumed = runCps.map((cps) => cps.map(() => false))
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
    if (!candidates.length) { if (causes) causes.noCandidate++; return false }
    const hidden = candidates.find((i) => !runs[i]!.servable)
    if (hidden !== undefined) {
      if (causes) { const why = runs[hidden]!.reason ?? 'unspecified'; causes.occluded[why] = (causes.occluded[why] ?? 0) + 1 }
      return false
    }
    const t = stripped(span.text)
    const hay = candidates.flatMap((run) => runCps[run]!.map((cp, at) => ({ cp, run, at })))
    const start = freeOccurrence(hay, t, consumed)
    if (start < 0) { if (causes) causes.mismatch++; return false }
    for (const h of hay.slice(start, start + t.length)) consumed[h.run]![h.at] = true
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
    else if (causes && verdict[i] === undefined) causes.whitespace++
  })
  return { served, withheld: spans.length - served.length }
}
