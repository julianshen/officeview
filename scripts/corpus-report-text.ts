import { createRecordingContext, type TextIndex, type TextSpan } from '../src/core/search'
import type { Paintable } from '../src/render/paint'
import { computeMetrics, layoutCellBody, paintCellLayout, textMeasurer, mergeEdges, rowOverflowFromCells, textOverflowAllowance } from '../src/xlsx/render'
import { cellSourceRow } from '../src/xlsx/view/sheetCells'
import { NO_OVERFLOW, type RowOverflowInput } from '../src/xlsx/view/textOverflow'
import { RectangleIndex } from '../src/xlsx/view/structure'
import { isValidCoordinate, MAX_COLS } from '../src/xlsx/view/coordinates'
import type { XlsxCell, XlsxMergeRange, XlsxDocument } from '../src/xlsx/types'

type Stratum = 'string' | 'other'
type Status = 'comparable' | 'unavailable' | 'refused'
type Coverage = Record<Stratum, { expected: number; matched: number }>
type Exclusion = { reason: 'collapsed' | 'merge-covered' | 'outside-geometry'; cells: number; sourceStringChars: number }
interface SheetCoverage {
  sheetIndex: number
  sheetName: string
  coverage: Coverage
  status: Status
  diagnostics: string[]
  collapsedGeometry: boolean
}
export interface XlsxTextCoverage {
  kind: 'xlsx-recording-preservation'
  allSourceStringChars: number
  eligibleCells: number
  eligibleSourceStringChars: number
  formattedExpectedChars: number
  expectedSpanChars: number
  matchedSpanChars: number
  recordedSpanChars: number
  unmatchedRecordedChars: number
  expectedSpans: number
  matchedSpans: number
  recordedSpans: number
  unmatchedRecordedSpans: number
  coverage: Coverage
  sheets: SheetCoverage[]
  excluded: Exclusion[]
  clipping: 'not-assessed'
  diagnostics: string[]
  status: Status
}
const coverage = (): Coverage => ({ string: { expected: 0, matched: 0 }, other: { expected: 0, matched: 0 } })
const chars = (text: string): number => text.trim().length
function prefix(sizes: number[]): number[] {
  const out = [0]
  for (const size of sizes) out.push(out[out.length - 1] + size)
  return out
}
function key(span: TextSpan): string | undefined {
  if (![span.x, span.y, span.width, span.fontSize].every(Number.isFinite) || span.width < 0 || span.fontSize < 0) return undefined
  return JSON.stringify([span.text, span.x, span.y, span.width, span.fontSize])
}
// Mirrors renderSheet's compact staticMergeIndex, including cover-before-anchor
// and last-anchor precedence, without expanding any covered cells.
function mergesFor(ranges: readonly XlsxMergeRange[]) {
  const anchors = new Map<string, XlsxMergeRange>()
  const entries = []
  for (const range of ranges) {
    anchors.set(`${range.minRow}:${range.minCol}`, range)
    if (isValidCoordinate(range.minRow, range.minCol) && isValidCoordinate(range.maxRow, range.maxCol) && range.minRow <= range.maxRow && range.minCol <= range.maxCol)
      entries.push({ rect: { r0: range.minRow, r1: range.maxRow, c0: range.minCol, c1: range.maxCol }, value: range })
  }
  const index = RectangleIndex.build(entries)
  return {
    inRow: (row: number) => isValidCoordinate(row, 0) ? index.collect({ r0: row, r1: row, c0: 0, c1: MAX_COLS - 1 }) : [],
    anchorAt: (row: number, col: number) => anchors.get(`${row}:${col}`),
    covers: (row: number, col: number) => isValidCoordinate(row, col) && index.firstWhere({ r0: row, r1: row, c0: col, c1: col }, r => r.minRow !== row || r.minCol !== col) !== undefined,
  }
}

/** Model enumeration versus outer paint/index replay. Shared format/layout/replay
 * cannot certify source parsing, native appearance or physical clipping. */
export function measureXlsxTextCoverage(doc: XlsxDocument, pages: readonly Paintable[], index: TextIndex | undefined, context: CanvasRenderingContext2D): XlsxTextCoverage {
  const result: XlsxTextCoverage = {
    kind: 'xlsx-recording-preservation', allSourceStringChars: 0, eligibleCells: 0,
    eligibleSourceStringChars: 0, formattedExpectedChars: 0, expectedSpanChars: 0,
    matchedSpanChars: 0, recordedSpanChars: 0, unmatchedRecordedChars: 0,
    expectedSpans: 0, matchedSpans: 0, recordedSpans: 0, unmatchedRecordedSpans: 0,
    coverage: coverage(), sheets: [], excluded: [], clipping: 'not-assessed', diagnostics: [], status: 'comparable',
  }
  const exclude = (reason: Exclusion['reason'], sourceStringChars: number) => {
    let entry = result.excluded.find(e => e.reason === reason)
    if (!entry) result.excluded.push(entry = { reason, cells: 0, sourceStringChars: 0 })
    entry.cells++; entry.sourceStringChars += sourceStringChars
  }
  const diagnose = (unit: SheetCoverage, message: string, status?: Status) => {
    if (unit.diagnostics.length < 12) unit.diagnostics.push(message)
    if (status && unit.status !== 'refused') unit.status = status
  }
  if (pages.length !== doc.sheets.length || !index || index.pages.length !== doc.sheets.length) {
    result.status = 'unavailable'
    result.diagnostics.push('unavailable page/sheet correspondence or text index')
  }
  for (const [sheetIndex, sheet] of doc.sheets.entries()) {
    const m = computeMetrics(sheet), x = prefix(m.colWidthsPx), y = prefix(m.rowHeightsPx)
    // Zero totals establish collapse only when every effective axis is valid;
    // malformed negative sizes can cancel a visible row's positive extent.
    const validSize = (size: number) => Number.isFinite(size) && size >= 0
    const validGeometry = [m.widthPx, m.heightPx].every(validSize) && m.colWidthsPx.every(validSize) && m.rowHeightsPx.every(validSize)
    const limits = [...(m.exportLimits ?? []), ...(pages[sheetIndex]?.refusal ?? [])]
    const unit: SheetCoverage = { sheetIndex, sheetName: sheet.name, coverage: coverage(), status: 'comparable', diagnostics: [], collapsedGeometry: validGeometry && limits.length === 0 && (m.widthPx === 0 || m.heightPx === 0) }
    result.sheets.push(unit)
    if (!validGeometry) diagnose(unit, 'invalid effective static geometry: axes and sheet extent must be finite and nonnegative', 'unavailable')
    for (const diagnostic of m.diagnostics ?? []) diagnose(unit, `${diagnostic.code}: ${diagnostic.message}`)
    if (limits.length) diagnose(unit, `static export refused: ${limits.map(l => l.message).join('; ')}`, 'refused')
    const page = index?.pages[sheetIndex], paintable = pages[sheetIndex]
    if (!page || page.index !== sheetIndex || !paintable) diagnose(unit, 'unavailable actual page/sheet correspondence', 'unavailable')
    if (!limits.length && paintable && (paintable.spec.widthPx !== m.widthPx || paintable.spec.heightPx !== m.heightPx)) diagnose(unit, 'unavailable static page geometry correspondence', 'unavailable')
    const expected: Array<{ span: TextSpan; stratum: Stratum }> = []
    const measured = textMeasurer(context), edges = mergeEdges(sheet, m, x, y), merges = mergesFor(sheet.mergeRanges ?? [])
    const rowBottom = new Map<number, number>()
    for (const range of sheet.mergeRanges ?? []) {
      const bottom = edges.row(range.maxRow + 1)
      if (bottom > (rowBottom.get(range.minRow) ?? -Infinity)) rowBottom.set(range.minRow, bottom)
    }
    for (const row of sheet.rows) {
      const sourceCells = new Map<number, XlsxCell[]>(), overflow = new Map<number, RowOverflowInput>()
      for (const cell of row.cells) {
        const source = cellSourceRow(row, cell), cells = sourceCells.get(source) ?? []
        cells.push(cell); sourceCells.set(source, cells)
      }
      for (const cell of row.cells) {
        const sourceChars = typeof cell.value === 'string' ? chars(cell.value) : 0
        result.allSourceStringChars += sourceChars
        if (limits.length || !validGeometry) continue // Refusal/invalid geometry cannot become successful exclusions.
        const source = cellSourceRow(row, cell)
        if (!Number.isInteger(source) || source < 0 || source >= m.rowHeightsPx.length || !Number.isInteger(cell.col) || cell.col < 0 || cell.col >= m.colWidthsPx.length) {
          exclude('outside-geometry', sourceChars)
          diagnose(unit, `outside static geometry: ${cell.ref}`, 'unavailable'); continue
        }
        const top = y[source], h = m.rowHeightsPx[source]
        // Same full-sheet intersection and expanded row-bottom guard as renderSheet.
        if (Math.max(top + h, rowBottom.get(source) ?? -Infinity) <= 0 || top >= m.heightPx) { exclude('collapsed', sourceChars); continue }
        if (merges.covers(cell.row, cell.col)) { exclude('merge-covered', sourceChars); continue }
        const range = merges.anchorAt(cell.row, cell.col), left = x[cell.col]
        const w = range ? edges.col(range.maxCol + 1) - left : m.colWidthsPx[cell.col]
        const height = range ? edges.row(range.maxRow + 1) - top : h
        if (![left, top, w, height].every(Number.isFinite)) {
          exclude('outside-geometry', sourceChars); diagnose(unit, `nonfinite static geometry: ${cell.ref}`, 'unavailable'); continue
        }
        if (!(w > 0 && height > 0)) { exclude('collapsed', sourceChars); continue }
        if (!(left < m.widthPx && left + w > 0 && top < m.heightPx && top + height > 0)) { exclude('outside-geometry', sourceChars); continue }
        result.eligibleCells++; result.eligibleSourceStringChars += sourceChars
        let input = overflow.get(source)
        if (!range && !input) {
          input = rowOverflowFromCells(sourceCells.get(source)!, merges.inRow(source), edges.col, MAX_COLS - 1)
          overflow.set(source, input)
        }
        context.save()
        try {
          context.textAlign = 'left'; context.textBaseline = 'alphabetic'
          const body = layoutCellBody(context, sheet, cell, left, top, w, height, measured, range ? NO_OVERFLOW : textOverflowAllowance(input!, cell.col))
          result.formattedExpectedChars += chars(body.formattedText)
          const spans: TextSpan[] = []
          paintCellLayout(createRecordingContext(context, spans), body, { width: m.widthPx, height: m.heightPx }, 'content')
          if (chars(body.formattedText) > 0 && spans.length === 0) diagnose(unit, `nonblank expected text produced no recording: ${cell.ref}`, 'unavailable')
          for (const span of spans) expected.push({ span, stratum: typeof cell.value === 'string' ? 'string' : 'other' })
        } finally { context.restore() }
      }
    }
    const remaining = new Map<string, TextSpan[]>()
    for (const span of page?.lines.flatMap(l => l.spans) ?? []) {
      result.recordedSpans++; result.recordedSpanChars += chars(span.text)
      const k = key(span)
      if (k === undefined) {
        diagnose(unit, 'nonfinite actual span geometry', 'unavailable')
        result.unmatchedRecordedSpans++; result.unmatchedRecordedChars += chars(span.text); continue
      }
      const bucket = remaining.get(k) ?? []
      bucket.push(span); remaining.set(k, bucket)
    }
    const expectedKeys = new Map<string, Array<{ span: TextSpan; stratum: Stratum }>>()
    for (const item of expected) {
      const n = chars(item.span.text), k = key(item.span)
      result.expectedSpans++; result.expectedSpanChars += n; unit.coverage[item.stratum].expected += n
      if (k === undefined) { diagnose(unit, 'nonfinite expected span geometry', 'unavailable'); continue }
      const bucket = expectedKeys.get(k) ?? []
      bucket.push(item); expectedKeys.set(k, bucket)
    }
    let missingSpans = 0
    for (const [k, items] of expectedKeys) {
      const actual = remaining.get(k) ?? []
      if (new Set(items.map(i => i.stratum)).size > 1 && actual.length < items.length) {
        diagnose(unit, 'ambiguous mixed-stratum key with insufficient multiplicity', 'unavailable')
        missingSpans += items.length; continue
      }
      for (const { span, stratum } of items) {
        if (actual.pop()) {
          const n = chars(span.text)
          result.matchedSpans++; result.matchedSpanChars += n; unit.coverage[stratum].matched += n
        } else missingSpans++
      }
    }
    if (missingSpans) diagnose(unit, `${missingSpans} unmatched expected span(s)`)
    for (const bucket of remaining.values()) for (const span of bucket) {
      result.unmatchedRecordedSpans++; result.unmatchedRecordedChars += chars(span.text)
    }
    for (const stratum of ['string', 'other'] as const) {
      result.coverage[stratum].expected += unit.coverage[stratum].expected
      result.coverage[stratum].matched += unit.coverage[stratum].matched
    }
    if (unit.status === 'refused' || result.status === 'comparable' && unit.status === 'unavailable') result.status = unit.status
    if (result.diagnostics.length < 24) result.diagnostics.push(...unit.diagnostics.slice(0, 24 - result.diagnostics.length).map(d => `sheet ${sheetIndex} (${sheet.name}): ${d}`))
  }
  // Unexpected extra pages remain recorded diagnostics, never coverage.
  for (const page of index?.pages.slice(doc.sheets.length) ?? []) for (const line of page.lines) for (const span of line.spans) {
    result.recordedSpans++; result.recordedSpanChars += chars(span.text)
    result.unmatchedRecordedSpans++; result.unmatchedRecordedChars += chars(span.text)
  }
  return result
}
export function hasRecordingLoss(result: XlsxTextCoverage): boolean {
  return result.status !== 'comparable' || result.sheets.some(s => (['string', 'other'] as const).some(k => s.coverage[k].expected > 0 && s.coverage[k].matched < s.coverage[k].expected * 0.5))
}
