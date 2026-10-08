/**
 * Text selection over the rendered canvas.
 *
 * Everything here is pure and works off the text index (see `search.ts`), which
 * already carries per-span geometry for every format — so hit-testing, word
 * expansion, range math and copy-text assembly need no DOM and no per-format
 * knowledge. A selection may span pages: a range is just an ordered pair of
 * carets, and the helpers slice it per page.
 */

import type { HighlightRect } from './overlay'
import type { IndexLine, TextIndex, TextIndexPage, TextPlacement } from './search'
import { pointInTextClip, projectTextPoint, rectsForRange, visibleTextPolygon } from './search'
import { snapGrapheme, type LogicalTextSource } from './text-recording'

export type { HighlightRect }

/** A caret position: which page, which visual line, which character. */
export interface CaretPos {
  pageIndex: number
  lineIndex: number
  charIndex: number
}

/** An ordered selection; `start` is always the earlier end. */
export interface SelectionRange {
  start: CaretPos
  end: CaretPos
}

/** One line's worth of a selection, on a particular page. */
export interface SelectionLineSlice {
  pageIndex: number
  lineIndex: number
  /** inclusive start character */
  from: number
  /** exclusive end character */
  to: number
}

function pageOf(index: TextIndex, pageIndex: number): TextIndexPage | undefined {
  return index.pages.find((p) => p.index === pageIndex)
}

/** Among containing vertical bands, prefer the text extent nearest x. */
function containingLineAt(page: TextIndexPage, x: number, y: number, accepts: (line: IndexLine) => boolean = () => true): { line: IndexLine; lineIndex: number } | undefined {
  let best: { line: IndexLine; lineIndex: number } | undefined
  let bestDistance = Infinity
  page.lines.forEach((line, lineIndex) => {
    if (y < line.top || y > line.bottom || !accepts(line)) return
    const left = Math.min(...line.spans.map(span => span.x))
    const right = Math.max(...line.spans.map(span => span.x + span.width))
    const distance = Math.max(left - x, x - right, 0)
    if (distance < bestDistance) { best = { line, lineIndex }; bestDistance = distance }
  })
  return best
}

/** Nearest vertical band; horizontal distance disambiguates overlapping bands. */
function lineAtY(page: TextIndexPage, y: number, x: number): { line: IndexLine; lineIndex: number } | undefined {
  if (page.lines.length === 0) return undefined
  const containing = containingLineAt(page, x, y)
  if (containing) return containing
  // no line contains y: snap to the vertically closest one
  let best = 0
  let bestDist = Infinity
  page.lines.forEach((line, i) => {
    const centre = (line.top + line.bottom) / 2
    const dist = Math.abs(centre - y)
    if (dist < bestDist) {
      bestDist = dist
      best = i
    }
  })
  return { line: page.lines[best], lineIndex: best }
}

/** Character offset within a line for an x coordinate. */
function charAtX(line: IndexLine, x: number): number {
  let offset = 0
  let lastEnd = 0
  for (const span of line.spans) {
    const spanStart = offset
    const spanEnd = offset + span.text.length
    offset = spanEnd
    if (span.text.length === 0) continue
    const spanRight = span.x + span.width
    if (x < span.x) return spanStart
    if (x <= spanRight) {
      const ratio = span.width > 0 ? (x - span.x) / span.width : 0
      // snap to the nearer edge, like every text editor
      const raw = spanStart + Math.round(ratio * span.text.length)
      return snapGrapheme(line.text, Math.min(spanEnd, Math.max(spanStart, raw)))
    }
    lastEnd = spanEnd
  }
  return lastEnd
}

/** Containment in a small convex mesh quad (same-side cross-product test). */
function pointInConvexQuad(
  quad: readonly [{ x: number; y: number }, { x: number; y: number }, { x: number; y: number }, { x: number; y: number }],
  x: number,
  y: number,
): boolean {
  if (!Number.isFinite(x) || !Number.isFinite(y)) return false
  if (!quad.every(p => Number.isFinite(p.x) && Number.isFinite(p.y))) return false
  const doubleArea = Math.abs(
    (quad[0].x * quad[1].y - quad[1].x * quad[0].y) +
    (quad[1].x * quad[2].y - quad[2].x * quad[1].y) +
    (quad[2].x * quad[3].y - quad[3].x * quad[2].y) +
    (quad[3].x * quad[0].y - quad[0].x * quad[3].y)
  )
  if (doubleArea <= 1e-9) return false
  let sign = 0
  for (let i = 0; i < 4; i++) {
    const p = quad[i], q = quad[(i + 1) % 4]
    const cross = (q.x - p.x) * (y - p.y) - (q.y - p.y) * (x - p.x)
    if (Math.abs(cross) <= 1e-9) continue
    const s = Math.sign(cross)
    if (sign !== 0 && s !== sign) return false
    sign = s
  }
  return sign !== 0
}

/** Union containment over mapped mesh cells (robust to global band winding). */
function pointInWarpCells(
  cells: ReadonlyArray<readonly [{ x: number; y: number }, { x: number; y: number }, { x: number; y: number }, { x: number; y: number }]>,
  x: number,
  y: number,
): boolean {
  return cells.some(cell => pointInConvexQuad(cell, x, y))
}

/** Union containment over mapped mesh triangles (authoritative: triangles stay
 * convex even when envelope bands twist and quads bowtie). */
function pointInWarpTris(
  tris: ReadonlyArray<readonly [{ x: number; y: number }, { x: number; y: number }, { x: number; y: number }]>,
  x: number,
  y: number,
): boolean {
  return tris.some(tri => pointInConvexQuad([tri[0], tri[1], tri[2], tri[2]], x, y))
}

function pointInSpanVisual(visual: NonNullable<TextPlacement['visual']>, x: number, y: number): boolean {
  return visual.tris?.length
    ? pointInWarpTris(visual.tris, x, y)
    : visual.cells?.length
      ? pointInWarpCells(visual.cells, x, y)
      : pointInTextClip(visual.polygon, x, y)
}

/** Inverse-map a hit to each span's local band so rotation/flips keep their character progression. */
function transformedHit(page: TextIndexPage, pageIndex: number, x: number, y: number, strict: boolean): CaretPos | undefined {
  let best: CaretPos | undefined
  let distance = Infinity
  for (const [lineIndex, line] of page.lines.entries()) {
    let offset = 0
    for (const span of line.spans) {
      const spanStart = offset
      offset += span.text.length
      if (!span.text.length) continue
      const p = span.placement ?? { x: span.x, y: span.y, width: span.width, top: line.top, bottom: line.bottom, transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } }
      if (p.clip && (!pointInTextClip(p.clip, x, y) || !visibleTextPolygon(p).length)) continue
      if (p.visual) {
        // Curved ink: containment on the mapped mesh-cell union (page space),
        // falling back to the mapped polygon; the caret resolves to the nearest
        // mapped cluster center snapped into the canonical grapheme grid.
        const contained = pointInSpanVisual(p.visual, x, y)
        if (strict && !contained) continue
        const logical = span.logical
        const boundaries = logical?.graphemeBoundaries
        if (p.visual.clusters?.length && boundaries?.length && logical) {
          let bestCluster = 0, bestDelta = Infinity
          p.visual.clusters.forEach((c, ci) => {
            const d = (c.x - x) ** 2 + (c.y - y) ** 2
            if (d < bestDelta) { bestDelta = d; bestCluster = ci }
          })
          const effDistance = Math.sqrt(bestDelta)
          if (effDistance >= distance) continue
          // Signed progression along the reading direction (works for vertical
          // frames too): project the hit onto the reading axis through the
          // selected cluster. The before/after split is at the cluster's OWN
          // center — matching the ordinary nearest-caret policy — not at a
          // neighbor Voronoi border: nearest-center selection already places
          // the query inside that cluster's region, so the sign alone
          // decides. The terminal cluster has no next center, so its axis
          // continues from the previous center; a single cluster uses the
          // whole-band cap axis (orientation-agnostic). Boundaries are
          // ABSOLUTE paragraph-source offsets; the span's absolute start maps
          // them onto this line exactly once.
          const spanSourceStart = logical.start
          const clusterCenter = p.visual.clusters[bestCluster]
          const hasNext = bestCluster + 1 < p.visual.clusters.length
          const hasPrev = bestCluster > 0
          const axisEnd = hasNext ? p.visual.clusters[bestCluster + 1] : clusterCenter
          const axisStart = hasNext ? clusterCenter : hasPrev ? p.visual.clusters[bestCluster - 1] : undefined
          let axisX: number, axisY: number
          if (axisStart) {
            axisX = axisEnd.x - axisStart.x; axisY = axisEnd.y - axisStart.y
          } else if (p.visual.cells && p.visual.cells.length > 0) {
            // Single cluster: whole-band reading axis from the start cap to
            // the end cap, recovered from the first/last mesh cells. This is
            // orientation-agnostic (vertical columns progress along the
            // column, not along the replay context's +x), unlike the span's
            // recording-time transform.
            const first = p.visual.cells[0], last = p.visual.cells[p.visual.cells.length - 1]
            const sx = (first[0].x + first[3].x) / 2, sy = (first[0].y + first[3].y) / 2
            const ex = (last[1].x + last[2].x) / 2, ey = (last[1].y + last[2].y) / 2
            axisX = ex - sx; axisY = ey - sy
          } else {
            // No neighbor and no cells: the span's baseline tangent is the
            // reading axis.
            const m = p.transform
            const len = Math.hypot(m.a, m.b) || 1
            axisX = m.a / len; axisY = m.b / len
          }
          const axisLen2 = axisX * axisX + axisY * axisY
          const signed = axisLen2 > 1e-12
            ? ((x - clusterCenter.x) * axisX + (y - clusterCenter.y) * axisY) / Math.sqrt(axisLen2)
            : 0
          const afterCenter = signed > 1e-9
          const boundary = boundaries[Math.min(bestCluster, boundaries.length - 1)]
          const nextBoundary = boundaries[Math.min(bestCluster + 1, boundaries.length - 1)]
          const chosen = afterCenter && nextBoundary > boundary ? nextBoundary : boundary
          const charIndex = spanStart + (chosen - spanSourceStart)
          best = { pageIndex, lineIndex, charIndex: snapGrapheme(line.text, charIndex) }
          distance = effDistance
          continue
        }
        if (strict) continue
        // fall through to the inverse-affine estimate when strict is off
      }
      const m = p.transform, determinant = m.a * m.d - m.b * m.c
      if (!Number.isFinite(determinant) || determinant === 0) continue
      const dx = x - m.e, dy = y - m.f
      const localX = (m.d * dx - m.c * dy) / determinant
      const localY = (m.a * dy - m.b * dx) / determinant
      const tolerance = 2 / Math.hypot(m.a, m.b)
      const inside = localX >= p.x - tolerance && localX <= p.x + p.width + tolerance && localY >= p.top && localY <= p.bottom
      if (strict && !inside) continue
      const clampedX = Math.max(p.x, Math.min(p.x + p.width, localX))
      const clampedY = Math.max(p.top, Math.min(p.bottom, localY))
      const closest = projectTextPoint(m, clampedX, clampedY)
      const dist = inside ? 0 : Math.hypot(closest.x - x, closest.y - y)
      if (!Number.isFinite(dist) || dist >= distance) continue
      const ratio = p.width > 0 ? (clampedX - p.x) / p.width : 0
      best = { pageIndex, lineIndex, charIndex: snapGrapheme(line.text, spanStart + Math.round(ratio * span.text.length)) }
      distance = dist
    }
  }
  return best
}

function positiveAxisAligned(line: IndexLine): boolean {
  if (line.spans.some(span => span.logical?.flow === 'vertical' || span.placement?.visual)) return false
  return line.spans.every(span => {
    const m = span.placement?.transform
    return !m || m.a > 0 && m.d > 0 && Math.abs(m.b) <= 1e-10 && Math.abs(m.c) <= 1e-10
  })
}

/**
 * Map a point in page coordinates to the nearest caret position.
 *
 * `strict` returns undefined unless the point actually lands on the line's
 * text (inside its vertical band *and* its horizontal extent) — used to tell
 * "clicked on text" from "clicked on blank page" (which zooms instead).
 */
export function hitTest(
  index: TextIndex,
  pageIndex: number,
  x: number,
  y: number,
  options: { strict?: boolean } = {},
): CaretPos | undefined {
  const page = pageOf(index, pageIndex)
  if (!page) return undefined
  if (page.lines.some(line => line.spans.some(span => span.placement))) {
    // Ordinary drags stay on a containing line's vertical band even past its text edges.
    // An unrelated rotated span elsewhere on the page must not change that behavior.
    if (!options.strict) {
      // Ordinary pages preserve the original y-first drag policy, including gaps between bands.
      if (page.lines.every(positiveAxisAligned)) {
        const eligible = page.lines.filter(line => !line.spans.length || line.spans.some(span => !span.placement?.clip || pointInTextClip(span.placement.clip, x, y) && visibleTextPolygon(span.placement).length > 0))
        const hit = lineAtY({ ...page, lines: eligible }, y, x)
        return hit ? { pageIndex, lineIndex: page.lines.indexOf(hit.line), charIndex: charAtX(hit.line, x) } : undefined
      }
      // A visible rotated/reflected scope can share an ordinary line's page-y
      // band. Actual local-band containment wins before off-edge drag snapping.
      const contained = transformedHit(page, pageIndex, x, y, true)
      if (contained) return contained
      const hit = containingLineAt(page, x, y, line => positiveAxisAligned(line) && (!line.spans.length || line.spans.some(span => !span.placement?.clip || pointInTextClip(span.placement.clip, x, y) && visibleTextPolygon(span.placement).length > 0)))
      if (hit) return { pageIndex, lineIndex: hit.lineIndex, charIndex: charAtX(hit.line, x) }
    }
    return transformedHit(page, pageIndex, x, y, options.strict ?? false)
  }
  const hit = lineAtY(page, y, x)
  if (!hit) return undefined
  if (options.strict) {
    const { line } = hit
    const left = Math.min(...line.spans.map((s) => s.x))
    const right = Math.max(...line.spans.map((s) => s.x + s.width))
    if (y < line.top || y > line.bottom || x < left - 2 || x > right + 2) return undefined
  }
  return { pageIndex, lineIndex: hit.lineIndex, charIndex: charAtX(hit.line, x) }
}

function posLess(a: CaretPos, b: CaretPos): boolean {
  if (a.pageIndex !== b.pageIndex) return a.pageIndex < b.pageIndex
  if (a.lineIndex !== b.lineIndex) return a.lineIndex < b.lineIndex
  return a.charIndex < b.charIndex
}

/** Order two carets into a forward range (works across pages). */
export function normalizeRange(a: CaretPos, b: CaretPos): SelectionRange {
  return posLess(a, b) || samePos(a, b) ? { start: a, end: b } : { start: b, end: a }
}

function samePos(a: CaretPos, b: CaretPos): boolean {
  return a.pageIndex === b.pageIndex && a.lineIndex === b.lineIndex && a.charIndex === b.charIndex
}

export function isEmptyRange(range: SelectionRange | undefined): boolean {
  if (!range) return true
  return samePos(range.start, range.end)
}

const WORD_CHARS = /[\p{L}\p{N}_'-]/u

/** Expand a caret to the word it sits in (double-click behaviour). */
export function wordRangeAt(line: IndexLine, charIndex: number): { start: number; end: number } {
  const text = line.text
  if (text.length === 0) return { start: 0, end: 0 }
  const isWord = (i: number) => i >= 0 && i < text.length && WORD_CHARS.test(text[i])
  let start = snapGrapheme(text, Math.min(Math.max(0, charIndex), text.length), 'floor')
  let end = start
  if (start > 0 && !isWord(start - 1) && isWord(start)) start -= 1
  if (!isWord(start) && isWord(start - 1)) start -= 1
  while (start > 0 && isWord(start - 1)) start--
  while (end < text.length && isWord(end)) end++
  if (end === start && start < text.length) {
    // not inside a word (punctuation/space): take the run of same-kind chars
    end = start + 1
    while (end < text.length && !WORD_CHARS.test(text[end]) && text[end] !== ' ') end++
  }
  return { start: snapGrapheme(text, start, 'floor'), end: snapGrapheme(text, end, 'ceil') }
}

/** The whole line a caret sits on (triple-click behaviour). */
export function lineRangeAt(line: IndexLine, _charIndex: number): { start: number; end: number } {
  return { start: 0, end: line.text.length }
}

/**
 * Split a range into per-line character slices, ordered by page then line.
 * Lines that are only *touched* (zero characters on them) are dropped, so
 * dragging from the end of one page to the start of the next does not paint
 * an empty band.
 */
export function selectionSlices(index: TextIndex, range: SelectionRange): SelectionLineSlice[] {
  // Defensive: the documented precondition is that `start` is the earlier end,
  // but this is a public export, so normalize rather than silently returning
  // nothing for a denormalized range.
  const ordered = posLess(range.start, range.end) || samePos(range.start, range.end) ? range : { start: range.end, end: range.start }
  const slices: SelectionLineSlice[] = []
  const pages = [...index.pages].sort((a, b) => a.index - b.index)
  for (const page of pages) {
    if (page.index < ordered.start.pageIndex || page.index > ordered.end.pageIndex) continue
    const from = page.index === ordered.start.pageIndex ? ordered.start.lineIndex : 0
    const to = page.index === ordered.end.pageIndex ? ordered.end.lineIndex : page.lines.length - 1
    const isStart = (li: number): boolean => page.index === ordered.start.pageIndex && li === ordered.start.lineIndex
    const isEnd = (li: number): boolean => page.index === ordered.end.pageIndex && li === ordered.end.lineIndex
    for (let li = from; li <= to && li < page.lines.length; li++) {
      const line = page.lines[li]
      const charFrom = isStart(li) ? ordered.start.charIndex : 0
      const charTo = isEnd(li) ? ordered.end.charIndex : line.text.length
      // zero-length touch points (page end -> next page start) paint nothing
      if (charTo > charFrom || line.text.length === 0 && line.spans[0]?.logical && !isEnd(li)) slices.push({ pageIndex: page.index, lineIndex: li, from: charFrom, to: charTo })
    }
  }
  return slices
}

/** Highlight rectangles for the part of a selection that lands on one page. */
export function rectsForSelectionOnPage(index: TextIndex, pageIndex: number, range: SelectionRange): HighlightRect[] {
  const page = pageOf(index, pageIndex)
  if (!page) return []
  const out: HighlightRect[] = []
  for (const slice of selectionSlices(index, range)) {
    if (slice.pageIndex !== pageIndex) continue
    const line = page.lines[slice.lineIndex]
    if (!line) continue
    out.push(...rectsForRange(line, slice.from, slice.to))
  }
  return out
}

/**
 * Text for a selection. Lines join with a newline; a page boundary also
 * breaks the line, so copied text reads as continuous prose.
 */
export function textForRange(index: TextIndex, range: SelectionRange): string {
  const slices = selectionSlices(index, range)
  const parts: Array<{ text: string; source?: LogicalTextSource; start?: number; end?: number }> = []
  for (let i = 0; i < slices.length; i++) {
    const slice = slices[i], page = pageOf(index, slice.pageIndex), line = page?.lines[slice.lineIndex]
    if (!line) continue
    const from = snapGrapheme(line.text, slice.from, 'floor'), to = snapGrapheme(line.text, slice.to, 'ceil')
    const logical = line.spans[0]?.logical
    if (logical && line.spans.every(span => span.logical?.source === logical.source)) {
      let offset = 0
      let sourceFrom = logical.start, sourceTo = logical.start
      for (const span of line.spans) {
        const spanStart = offset; offset += span.text.length
        if (from >= spanStart && from <= offset) sourceFrom = span.logical!.start + from - spanStart
        if (to >= spanStart && to <= offset) sourceTo = span.logical!.start + to - spanStart
      }
      const previous = parts[parts.length - 1]
      if (previous?.source === logical.source) {
        previous.end = sourceTo; previous.text = logical.source.text.slice(previous.start, sourceTo)
      } else parts.push({ source: logical.source, start: sourceFrom, end: sourceTo, text: logical.source.text.slice(sourceFrom, sourceTo) })
    } else {
      const piece = line.text.slice(from, to)
      parts.push({ text: i === slices.length - 1 ? piece : piece.replace(/\s+$/, '') })
    }
  }
  return parts.map(part => part.text).join('\n')
}
