/**
 * Text selection over the rendered canvas.
 *
 * Everything here is pure and works off the text index (see `search.ts`), which
 * already carries per-span geometry for every format — so hit-testing, word
 * expansion and copy-text assembly need no DOM and no per-format knowledge.
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

/** An ordered selection within one page. */
export interface SelectionRange {
  pageIndex: number
  startLine: number
  startChar: number
  endLine: number
  endChar: number
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

/** Order two carets into a forward range. */
export function normalizeRange(a: CaretPos, b: CaretPos): SelectionRange {
  const [start, end] = posLess(a, b) ? [a, b] : [b, a]
  return {
    pageIndex: start.pageIndex,
    startLine: start.lineIndex,
    startChar: start.charIndex,
    endLine: end.lineIndex,
    endChar: end.charIndex,
  }
}

export function isEmptyRange(range: SelectionRange | undefined): boolean {
  if (!range) return true
  return range.startLine === range.endLine && range.startChar === range.endChar
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

/** Highlight rectangles for a selection, in page coordinates. */
export function rectsForSelection(index: TextIndex, range: SelectionRange): HighlightRect[] {
  const page = pageOf(index, range.pageIndex)
  if (!page) return []
  const out: HighlightRect[] = []
  for (let li = range.startLine; li <= range.endLine && li < page.lines.length; li++) {
    const line = page.lines[li]
    const from = li === range.startLine ? range.startChar : 0
    const to = li === range.endLine ? range.endChar : line.text.length
    if (to <= from) continue
    out.push(...rectsForRange(line, from, to))
  }
  return out
}

/**
 * Text for a selection: one line per visual line, with trailing spaces trimmed
 * (what a browser puts on the clipboard for a dragged selection).
 */
export function textForRange(index: TextIndex, range: SelectionRange): string {
  const page = pageOf(index, range.pageIndex)
  if (!page) return ''
  const parts: string[] = []
  for (let li = range.startLine; li <= range.endLine && li < page.lines.length; li++) {
    const line = page.lines[li]
    const from = li === range.startLine ? range.startChar : 0
    const to = li === range.endLine ? range.endChar : line.text.length
    if (to < from) continue
    const lastLine = li === range.endLine
    parts.push((lastLine ? line.text.slice(from, to) : line.text.slice(from)).replace(/\s+$/, ''))
  }
  return parts.join('\n')
}

/** Per-line char ranges covered by a selection (used for drag feedback). */
export function lineRangesForSelection(range: SelectionRange, lineCount: number): Array<{ lineIndex: number; from: number; to: number }> {
  const out: Array<{ lineIndex: number; from: number; to: number }> = []
  for (let li = range.startLine; li <= range.endLine && li < lineCount; li++) {
    out.push({
      lineIndex: li,
      from: li === range.startLine ? range.startChar : 0,
      to: li === range.endLine ? range.endChar : Number.MAX_SAFE_INTEGER,
    })
  }
  return out
}
