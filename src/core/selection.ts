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
import type { IndexLine, TextIndex, TextIndexPage } from './search'
import { rectsForRange } from './search'

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

/** Nearest line to a y coordinate, preferring ones that contain it. */
function lineAtY(page: TextIndexPage, y: number): { line: IndexLine; lineIndex: number } | undefined {
  if (page.lines.length === 0) return undefined
  let containing: { line: IndexLine; lineIndex: number } | undefined
  for (let i = 0; i < page.lines.length; i++) {
    const line = page.lines[i]
    if (y >= line.top && y <= line.bottom) {
      containing = { line, lineIndex: i }
      break
    }
  }
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
      return Math.min(spanEnd, Math.max(spanStart, raw))
    }
    lastEnd = spanEnd
  }
  return lastEnd
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
  const hit = lineAtY(page, y)
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
  let start = Math.min(Math.max(0, charIndex), text.length)
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
  return { start, end }
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
  const slices: SelectionLineSlice[] = []
  const pages = [...index.pages].sort((a, b) => a.index - b.index)
  for (const page of pages) {
    if (page.index < range.start.pageIndex || page.index > range.end.pageIndex) continue
    const from = page.index === range.start.pageIndex ? range.start.lineIndex : 0
    const to = page.index === range.end.pageIndex ? range.end.lineIndex : page.lines.length - 1
    const isStart = (li: number): boolean => page.index === range.start.pageIndex && li === range.start.lineIndex
    const isEnd = (li: number): boolean => page.index === range.end.pageIndex && li === range.end.lineIndex
    for (let li = from; li <= to && li < page.lines.length; li++) {
      const line = page.lines[li]
      const charFrom = isStart(li) ? range.start.charIndex : 0
      const charTo = isEnd(li) ? range.end.charIndex : line.text.length
      // zero-length touch points (page end -> next page start) paint nothing
      if (charTo > charFrom) slices.push({ pageIndex: page.index, lineIndex: li, from: charFrom, to: charTo })
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
  const parts: string[] = []
  for (let i = 0; i < slices.length; i++) {
    const slice = slices[i]
    const page = pageOf(index, slice.pageIndex)
    const line = page?.lines[slice.lineIndex]
    if (!line) continue
    const piece = line.text.slice(slice.from, slice.to)
    // trim trailing whitespace on every line except the last, the way a
    // browser trims a dragged selection
    const isLast = i === slices.length - 1
    parts.push(isLast ? piece : piece.replace(/\s+$/, ''))
  }
  return parts.join('\n')
}
