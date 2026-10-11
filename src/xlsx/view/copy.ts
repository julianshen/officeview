import type { AuthoritativeCell, CacheIdentity, GridSelection, ProfileId, RowOrder, RowVisibility, SourceStore } from './contracts'
import { cellId, validateSourceRect } from './coordinates'
import { rowSetIntervals, sameIdentity, sameStamp } from '../../worker/viewerSession'
import { MAIN_THREAD_BUDGETS } from './tileCache'
import { copyText, type CopyOutcome } from '../../core/clipboard'

export const CLIPBOARD_BYTES = 8 * 1024 * 1024
/** The in-process SourceStore does not expose the S4.3 live reservation ledger.
 * Reserve at most 3/8 of its main-thread profile for output buffers/projection:
 * D=24 MiB, M=12 MiB. A worker caller can pass its admitted reservation instead. */
export function copySessionBudget(profile: ProfileId): number { return MAIN_THREAD_BUDGETS[profile] * 3 / 8 }
export const DEFAULT_COPY_SESSION_BYTES = copySessionBudget('D')
const METADATA_BYTES = 4 * 1024 * 1024
const BATCH_CELLS = 256
export interface CopyInput {
  store: SourceStore
  owner: CacheIdentity
  currentOwner(): CacheIdentity
  selection: GridSelection
  order: RowOrder
  visibility: RowVisibility
  columns: RowVisibility
  allowed(): boolean
  signal: AbortSignal
  format?: (cell: AuthoritativeCell) => string
  onPending?: () => void
  /** Peak output buffers plus compact projection, independent of the row count. */
  sessionBudgetBytes?: number
}
export type CopyResult = { kind: 'ready'; text: string; bytes: number } | { kind: 'refused'; code: string; message: string } | { kind: 'cancelled' | 'stale' | 'denied' }
interface Run { lo: number; hi: number }
interface Block { rows: Run; cols: Run }
class CopyFailure extends Error {
  constructor(readonly code: string, message: string) { super(message) }
}
const limit = () => new CopyFailure('resource-limit', 'Copy is too large. Select a smaller range.')
function check(input: CopyInput): void {
  if (input.signal.aborted) throw new CopyFailure('cancelled', 'Copy cancelled.')
  if (!input.allowed()) throw new CopyFailure('denied', 'Copy is not permitted.')
  if (!sameIdentity(input.owner, input.currentOwner()) || !sameStamp(input.owner, input.store.stamp)) throw new CopyFailure('stale', 'Source or view changed; copy again.')
}
function merged(runs: Run[]): Run[] {
  const out: Run[] = []
  for (const run of runs.sort((a, b) => a.lo - b.lo)) {
    const last = out[out.length - 1]
    if (last && last.hi + 1 >= run.lo) last.hi = Math.max(last.hi, run.hi)
    else out.push({ ...run })
  }
  return out
}
/** Identity spans remain intervals; only partial permutation bands are visited.
 * No cell values or selected row×column product is materialized by preflight. */
function projectedRows(input: CopyInput, part: GridSelection['ranges'][number]): Run[] {
  const intervals = rowSetIntervals(part.rows)
  if (!intervals) throw new CopyFailure('unsupported', 'Copy needs the worker-owned selection membership.')
  const out: Run[] = []
  const push = (lo: number, hi: number) => {
    const first = input.visibility.rank(lo), end = input.visibility.rank(hi + 1)
    if (end > first) out.push({ lo: first, hi: end - 1 })
    if (out.length * 64 > METADATA_BYTES) throw limit()
  }
  for (const { r0, r1 } of intervals) {
    let at = r0
    for (const band of input.order.bands) {
      if (band.r1 < at || band.r0 > r1) continue
      if (at < band.r0) push(at, band.r0 - 1)
      const lo = Math.max(at, band.r0), hi = Math.min(r1, band.r1)
      if (lo === band.r0 && hi === band.r1) push(lo, hi)
      else for (let row = lo; row <= hi; row++) {
        if (row % 4096 === 0) check(input)
        const slot = input.order.sourceToDisplay(row)
        push(slot, slot)
      }
      at = hi + 1
    }
    if (at <= r1) push(at, r1)
  }
  return merged(out)
}
/** Range-count tree over at most 16,384 visible columns. Whole covered/empty
 * subtrees stop traversal, so collecting the union visits output runs rather
 * than rescanning every selection rectangle at every row boundary. */
class ColumnCoverage {
  private readonly counts: Int32Array
  private readonly covered: Uint32Array
  constructor(private readonly size: number) {
    this.counts = new Int32Array(size * 4)
    this.covered = new Uint32Array(size * 4)
  }
  change(run: Run, delta: number, node = 1, lo = 0, hi = this.size - 1): void {
    if (run.lo <= lo && hi <= run.hi) this.counts[node] += delta
    else {
      const mid = Math.floor((lo + hi) / 2)
      if (run.lo <= mid) this.change(run, delta, node * 2, lo, mid)
      if (run.hi > mid) this.change(run, delta, node * 2 + 1, mid + 1, hi)
    }
    this.covered[node] = this.counts[node] > 0 ? hi - lo + 1 : lo === hi ? 0 : this.covered[node * 2] + this.covered[node * 2 + 1]
  }
  runs(): Run[] {
    const out: Run[] = []
    const visit = (node: number, lo: number, hi: number) => {
      if (!this.covered[node]) return
      if (this.covered[node] === hi - lo + 1) {
        const last = out[out.length - 1]
        if (last && last.hi + 1 === lo) last.hi = hi
        else out.push({ lo, hi })
        return
      }
      const mid = Math.floor((lo + hi) / 2)
      visit(node * 2, lo, mid); visit(node * 2 + 1, mid + 1, hi)
    }
    visit(1, 0, this.size - 1)
    return out
  }
}
/** Union in visible-rank space: hidden gaps disappear, overlapping parts dedupe.
 * Sort row events once; each changes column coverage in O(log C). Disjoint
 * rectangles use top-left order and a blank line between blocks. */
function blocksFor(input: CopyInput): Block[] {
  const events: { row: number; cols: Run; delta: number }[] = []
  let columnCount = 0
  for (const part of input.selection.ranges) {
    const rows = projectedRows(input, part)
    const cols = merged(part.cols.flatMap(({ c0, c1 }) => {
      validateSourceRect({ r0: 0, r1: 0, c0, c1 })
      const lo = input.columns.rank(c0), end = input.columns.rank(c1 + 1)
      return end > lo ? [{ lo, hi: end - 1 }] : []
    }))
    if ((events.length + rows.length * cols.length * 2) * 32 > METADATA_BYTES) throw limit()
    for (const r of rows) for (const c of cols) {
      columnCount = Math.max(columnCount, c.hi + 1)
      events.push({ row: r.lo, cols: c, delta: 1 }, { row: r.hi + 1, cols: c, delta: -1 })
    }
  }
  if (!events.length) return []
  events.sort((a, b) => a.row - b.row)
  const coverage = new ColumnCoverage(columnCount)
  const blocks: Block[] = [], lastByColumns = new Map<string, Block>()
  let lowerBytes = 0
  for (let i = 0; i < events.length;) {
    check(input)
    const lo = events[i].row
    do { coverage.change(events[i].cols, events[i].delta); i++ } while (i < events.length && events[i].row === lo)
    if (i === events.length) break
    const hi = events[i].row - 1
    for (const c of coverage.runs()) {
      const key = `${c.lo}:${c.hi}`, last = lastByColumns.get(key), rows = hi - lo + 1, cols = c.hi - c.lo + 1
      if (last && last.rows.hi + 1 === lo) { last.rows.hi = hi; lowerBytes += 2 * rows * cols }
      else {
        lowerBytes += 2 * (rows * (cols - 1) + rows - 1) + (blocks.length ? 4 : 0)
        const block = { rows: { lo, hi }, cols: { ...c } }
        blocks.push(block); lastByColumns.set(key, block)
      }
    }
    // Output-derived rejection, independent of the number of row edges visited.
    if (lowerBytes > CLIPBOARD_BYTES || blocks.length * 128 > METADATA_BYTES) throw limit()
  }
  return blocks.sort((a, b) => a.rows.lo - b.rows.lo || a.cols.lo - b.cols.lo)
}
function abortable<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
  return new Promise((resolve, reject) => {
    const abort = () => reject(new CopyFailure('cancelled', 'Copy cancelled.'))
    if (signal.aborted) { abort(); return }
    signal.addEventListener('abort', abort, { once: true })
    promise.then(resolve, reject).finally(() => signal.removeEventListener('abort', abort))
  })
}
async function read(input: CopyInput, row: number, col: number): Promise<string> {
  const at = { row, col }, options = { stamp: input.owner, signal: input.signal }
  let completeCoverage = false
  for (;;) {
    check(input)
    const result = await abortable(input.store.readCell(at, options), input.signal)
    check(input)
    if (result.kind === 'unknown') {
      if (completeCoverage) throw new CopyFailure('page-error', 'Copy source returned unknown values after confirming complete coverage.')
      input.onPending?.()
      const range = { r0: row, r1: row, c0: col, c1: col }
      const complete = await abortable(input.store.ensureComplete(range, options), input.signal)
      check(input)
      if (!sameStamp(complete.stamp, input.owner)) throw new CopyFailure('stale', 'Source changed; copy again.')
      if (complete.error) throw new CopyFailure(complete.error.code, complete.error.message)
      // A progressive source may return loading until the next import batch.
      completeCoverage = complete.completeness === 'complete'
      if (!completeCoverage) await abortable(new Promise<void>(resolve => setTimeout(resolve, 25)), input.signal)
      continue
    }
    if (result.kind === 'blank') return ''
    if (result.kind === 'error') throw new CopyFailure(result.code, `Copy could not read a cell (${result.code}).`)
    if (result.cell.row !== row || result.cell.col !== col || result.cell.id !== cellId(row, col)) throw new CopyFailure('page-error', 'Copy source returned a different cell.')
    if (result.kind === 'unavailable' || result.cell.unavailable) throw new CopyFailure('formula-value-unavailable', 'Copy contains a formula whose saved value is unavailable.')
    return input.format ? input.format(result.cell) : String(result.cell.value ?? '')
  }
}
/** TSV escaping quotes fields containing tab, CR, LF or quote; embedded quotes
 * double. Strings stay raw UTF-16, including lone surrogates. */
function escaped(text: string, remainingUnits: number): string {
  if (text.length > remainingUnits) throw limit()
  if (!/[\t\r\n"]/.test(text)) return text
  let quotes = 0
  for (let i = 0; i < text.length; i++) if (text[i] === '"') quotes++
  if (text.length + quotes + 2 > remainingUnits) throw limit()
  return `"${text.replace(/"/g, '""')}"`
}
export async function serializeSelection(input: CopyInput): Promise<CopyResult> {
  try {
    check(input)
    const blocks = blocksFor(input)
    const lower = blocks.reduce((sum, block) => {
      const r = block.rows.hi - block.rows.lo + 1, c = block.cols.hi - block.cols.lo + 1
      return sum + 2 * (r * (c - 1) + r - 1)
    }, Math.max(0, blocks.length - 1) * 4)
    if (lower > CLIPBOARD_BYTES) throw limit()
    const sessionBytes = input.sessionBudgetBytes ?? DEFAULT_COPY_SESSION_BYTES
    const metadata = blocks.length * 128
    if (lower * 2 + metadata > sessionBytes) throw limit()
    const chunks: string[] = []
    let buffer = '', units = 0, batches = 0
    const append = (text: string) => {
      const next = units + text.length
      if (next * 2 > CLIPBOARD_BYTES || next * 4 + metadata > sessionBytes) throw limit()
      units = next; buffer += text
      if (buffer.length >= 65536) { chunks.push(buffer); buffer = '' }
    }
    for (let b = 0; b < blocks.length; b++) {
      if (b) append('\n\n')
      const block = blocks[b], width = block.cols.hi - block.cols.lo + 1
      const cells = (block.rows.hi - block.rows.lo + 1) * width
      for (let offset = 0; offset < cells; offset += BATCH_CELLS) {
        check(input)
        const n = Math.min(BATCH_CELLS, cells - offset)
        const values: string[] = []
        let batchBytes = 0
        for (let i = 0; i < n; i++) {
          const index = offset + i, rank = block.rows.lo + Math.floor(index / width)
          const row = input.order.displayToSource(input.visibility.select(rank)!)
          const col = input.columns.select(block.cols.lo + index % width)!
          const text = await read(input, row, col)
          batchBytes += text.length * 2
          if (units * 4 + batchBytes + metadata > sessionBytes || units * 2 + batchBytes > CLIPBOARD_BYTES) throw limit()
          values.push(text)
        }
        check(input)
        for (let i = 0; i < n; i++) {
          const index = offset + i
          if (index) append(index % width ? '\t' : '\n')
          append(escaped(values[i], CLIPBOARD_BYTES / 2 - units))
        }
        if (++batches % 16 === 0) await abortable(new Promise<void>(resolve => setTimeout(resolve, 0)), input.signal)
      }
    }
    if (buffer) chunks.push(buffer)
    check(input)
    return { kind: 'ready', text: chunks.join(''), bytes: units * 2 }
  } catch (error) {
    if (error instanceof CopyFailure) {
      if (error.code === 'cancelled' || error.code === 'stale' || error.code === 'denied') return { kind: error.code }
      return { kind: 'refused', code: error.code, message: error.message }
    }
    return { kind: 'refused', code: 'page-error', message: error instanceof Error ? error.message : 'Copy unavailable; retry.' }
  }
}
export async function copySelection(input: CopyInput): Promise<CopyResult | { kind: 'copied'; outcome: CopyOutcome }> {
  const result = await serializeSelection(input)
  if (result.kind !== 'ready') return result
  try { check(input) } catch (error) { return { kind: (error as CopyFailure).code as 'cancelled' | 'stale' | 'denied' } }
  const outcome = await copyText(result.text, { allowed: () => input.allowed() && sameIdentity(input.owner, input.currentOwner()) && sameStamp(input.owner, input.store.stamp), signal: input.signal })
  return outcome === 'failed' ? { kind: 'refused', code: 'clipboard-error', message: 'Copy failed. Try copying again.' } : { kind: 'copied', outcome }
}
