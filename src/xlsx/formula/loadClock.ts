/**
 * The across-midnight follow-up (brief cf-midnight): the clock a resident document was loaded under, and the cells whose
 * values may move with the clock.
 *
 * A resident document's formulas are evaluated once at parse (`evaluateWorkbookFormulas`). A `ca` or uncached `TODAY()`
 * cell then names the load day for as long as the document stays open, while a clock-reading conditional-formatting rule
 * is re-evaluated at every preparation. Once the local day differs from the load day, the cell and the rule can name two
 * days, a state Excel was never measured in. The CF service withholds those cells (`src/xlsx/cf/midnight.ts`); this module
 * records what it needs, eagerly at parse, through the same `now` seam the load-time evaluation reads.
 *
 * The taint is every cell whose value could move with the clock: a formula cell that calls a volatile function, any formula
 * cell that cannot be read statically (fail closed), every formula that reads one of them (the static reference graph,
 * defined names expanded), and the spill followers of a tainted owner. It is computed from the formula text, so it is the
 * same for an evaluated and a saved-values-only document: a cell Excel would recalculate when the day changes is tainted
 * whether or not this viewer recalculated it at load.
 */
import type { XlsxDocument, XlsxSheet } from '../types'

/** Tainted cells by lower-cased sheet name, each `row * 16384 + col` (0-based), or `'all'` (every cell, fail closed). */
export type DocumentClockTaint = ReadonlyMap<string, ReadonlySet<number>> | 'all'
export interface DocumentLoadClock {
  /** The instant the load-time evaluation read (the parse's one clock sample). */
  readonly epochMs: number
  /** The zone and date system the load-time evaluation read it in: the day a `ca` TODAY() cell names. */
  readonly timeZone: string
  readonly dateSystem: '1900' | '1904'
  readonly taint: DocumentClockTaint
}

const records = new WeakMap<object, DocumentLoadClock>()

/** Records a document's load clock against the document and each of its sheets (the CF routes hold a sheet). */
export function recordDocumentLoadClock(doc: Pick<XlsxDocument, 'sheets'>, clock: DocumentLoadClock): void {
  const frozen = Object.freeze({ epochMs: clock.epochMs, timeZone: clock.timeZone, dateSystem: clock.dateSystem, taint: clock.taint })
  records.set(doc, frozen)
  for (const sheet of doc.sheets) records.set(sheet, frozen)
}

/** The load clock of a document or one of its sheets; undefined for a document no parse recorded. */
export function documentLoadClockOf(owner: object): DocumentLoadClock | undefined {
  return records.get(owner)
}

/**
 * A recalculation after load that sampled the clock (a pivot commit's `recomputeDependents`) may have moved some cells to
 * another instant: the taint becomes every cell (fail closed). The load instant is kept.
 */
export function markDocumentClockUnknown(doc: Pick<XlsxDocument, 'sheets'>): void {
  const prior = records.get(doc)
  if (prior && prior.taint !== 'all') recordDocumentLoadClock(doc, { ...prior, taint: 'all' })
}

/** Whether a document (or sheet) holds any clock-tainted cell (true when unknown). */
export function documentClockTainted(owner: object): boolean {
  const record = records.get(owner)
  if (!record) return false
  return record.taint === 'all' || [...record.taint.values()].some(set => set.size > 0)
}

/** Whether `rect` (0-based, inclusive) of `sheetName` holds a tainted cell. */
export function clockTaintHas(taint: DocumentClockTaint, sheetName: string, rect: Readonly<{ firstRow: number; lastRow: number; firstCol: number; lastCol: number }>): boolean {
  if (taint === 'all') return true
  const cells = taint.get(sheetName.toLowerCase())
  if (!cells || cells.size === 0) return false
  const area = (rect.lastRow - rect.firstRow + 1) * (rect.lastCol - rect.firstCol + 1)
  if (area <= cells.size) {
    for (let row = rect.firstRow; row <= rect.lastRow; row++) for (let col = rect.firstCol; col <= rect.lastCol; col++) if (cells.has(row * 16384 + col)) return true
    return false
  }
  for (const id of cells) {
    const row = Math.floor(id / 16384), col = id % 16384
    if (row >= rect.firstRow && row <= rect.lastRow && col >= rect.firstCol && col <= rect.lastCol) return true
  }
  return false
}

/** The sheet a key's lower-cased name belongs to (used by tests and the survey). */
export function sheetTaintOf(taint: DocumentClockTaint, sheet: Pick<XlsxSheet, 'name'>): ReadonlySet<number> | 'all' {
  return taint === 'all' ? 'all' : taint.get(sheet.name.toLowerCase()) ?? new Set()
}

/**
 * A paged workbook shows saved values and holds no reference graph: it may hold clock-tainted cells unless its completed
 * import counted no clock-sensitive formula (no volatile or `ca` formula; `view/values.ts` isClockSensitive).
 */
export function pagedWorkbookClockTainted(workbook: Readonly<{ completeness?: string; clockFormulas?: number }>): boolean {
  return !(workbook.completeness === 'complete' && workbook.clockFormulas === 0)
}
