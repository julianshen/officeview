import type { AuthoritativeCell, CacheIdentity, GridSelection, ProfileId, RowOrder, RowVisibility, SourceStore } from './contracts'
import { cellId, validateSourceRect } from './coordinates'
import { rowSetIntervals, sameIdentity, sameStamp } from '../../worker/viewerSession'
import { MAIN_THREAD_BUDGETS } from './tileCache'
import { copyText, type CopyOutcome } from '../../core/clipboard'
import type { FormatAllocation } from '../formula/format'

export const CLIPBOARD_BYTES = 8 * 1024 * 1024
/** Per-operation ceiling, D=24 MiB / M=12 MiB; incremental caller admission
 * shares the document budget rather than reserving this full allowance. */
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
  /** Canonical production formatting admits construction through the context.
   * Legacy callbacks keep their API and own any unobserved construction. */
  format?: (cell: AuthoritativeCell, allocation?: FormatAllocation) => string
  onPending?: () => void
  /** Peak output buffers plus compact projection, independent of the row count. */
  sessionBudgetBytes?: number
  /** S10.2 main ledger: the session's live bytes by the same formula as the budget. */
  onBytes?: (bytes: number) => void
  /** Atomic absolute reservation; refusal leaves the previous charge intact. */
  admitBytes?: (bytes: number) => boolean
  /** Adopt reserved canonical diagnostics into their document lifetime. */
  retainBytes?: (targets: readonly object[], bytes: number, commit: () => void) => void
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
/** One writer, shared admission, and scoped storage estimates. Source strings
 * and selection membership remain borrowed; only copy-owned storage is added. */
class CopyMemory {
  bytes = 0
  constructor(private readonly input: CopyInput) {}
  set(next: number): void {
    if (!Number.isSafeInteger(next) || next < 0 || next > (this.input.sessionBudgetBytes ?? DEFAULT_COPY_SESSION_BYTES)) throw limit()
    if (this.input.admitBytes && !this.input.admitBytes(next)) throw limit()
    this.bytes = next
    this.input.onBytes?.(next)
  }
  add(bytes: number): void { this.set(this.bytes + bytes) }
}
function sorted<T>(items: T[], compare: (a: T, b: T) => number, memory: CopyMemory): void {
  const scratch = items.length * 8
  memory.add(scratch)
  items.sort(compare)
  memory.add(-scratch)
}
function merged(runs: Run[], memory: CopyMemory): Run[] {
  sorted(runs, (a, b) => a.lo - b.lo, memory)
  let size = 0
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i], last = runs[size - 1]
    if (last && last.hi + 1 >= run.lo) last.hi = Math.max(last.hi, run.hi)
    else runs[size++] = run
  }
  const released = (runs.length - size) * 64
  runs.length = size
  memory.add(-released)
  return runs
}
/** Identity spans remain intervals; only partial permutation bands are visited. */
function projectedRows(input: CopyInput, part: GridSelection['ranges'][number], memory: CopyMemory): Run[] {
  let decoded = 0
  let intervals = rowSetIntervals(part.rows, () => {
    if (decoded + 64 > METADATA_BYTES) throw limit()
    memory.add(64); decoded += 64
  })
  if (!intervals) throw new CopyFailure('unsupported', 'Copy needs the worker-owned selection membership.')
  const out: Run[] = []
  const push = (lo: number, hi: number) => {
    const first = input.visibility.rank(lo), end = input.visibility.rank(hi + 1)
    if (end > first) {
      if ((out.length + 1) * 64 > METADATA_BYTES) throw limit()
      memory.add(64)
      out.push({ lo: first, hi: end - 1 })
    }
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
  intervals = null
  memory.add(-decoded)
  return merged(out, memory)
}
/** Range-count tree over at most 16,384 visible columns. */
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
  runs(memory: CopyMemory): Run[] {
    const out: Run[] = []
    const visit = (node: number, lo: number, hi: number) => {
      if (!this.covered[node]) return
      if (this.covered[node] === hi - lo + 1) {
        const last = out[out.length - 1]
        if (last && last.hi + 1 === lo) last.hi = hi
        else { memory.add(64); out.push({ lo, hi }) }
        return
      }
      const mid = Math.floor((lo + hi) / 2)
      visit(node * 2, lo, mid); visit(node * 2 + 1, mid + 1, hi)
    }
    visit(1, 0, this.size - 1)
    return out
  }
}
/** Union in visible-rank space, with admission before every owned growth. */
function blocksFor(input: CopyInput, memory: CopyMemory): Block[] {
  const events: { row: number; cols: Run; delta: number }[] = []
  let columnCount = 0, columnBytes = 0
  for (const part of input.selection.ranges) {
    const rows = projectedRows(input, part, memory), cols: Run[] = []
    for (const { c0, c1 } of part.cols) {
      validateSourceRect({ r0: 0, r1: 0, c0, c1 })
      const lo = input.columns.rank(c0), end = input.columns.rank(c1 + 1)
      if (end > lo) {
        if ((cols.length + 1) * 64 > METADATA_BYTES) throw limit()
        memory.add(64); cols.push({ lo, hi: end - 1 })
      }
    }
    merged(cols, memory)
    if ((events.length + rows.length * cols.length * 2) * 32 > METADATA_BYTES) throw limit()
    for (const r of rows) for (const c of cols) {
      columnCount = Math.max(columnCount, c.hi + 1)
      memory.add(64)
      events.push({ row: r.lo, cols: c, delta: 1 }, { row: r.hi + 1, cols: c, delta: -1 })
    }
    const rowBytes = rows.length * 64
    rows.length = 0
    memory.add(-rowBytes)
    if (rowBytes) columnBytes += cols.length * 64
    else { const bytes = cols.length * 64; cols.length = 0; memory.add(-bytes) }
  }
  if (!events.length) return []
  sorted(events, (a, b) => a.row - b.row, memory)
  memory.add(32 * columnCount)
  let coverage: ColumnCoverage | null = new ColumnCoverage(columnCount)
  const blocks: Block[] = []
  let lastByColumns: Map<string, Block> | null = new Map(), lowerBytes = 0, mapBytes = 0
  for (let i = 0; i < events.length;) {
    check(input)
    const lo = events[i].row
    do { coverage.change(events[i].cols, events[i].delta); i++ } while (i < events.length && events[i].row === lo)
    if (i === events.length) break
    const hi = events[i].row - 1, runs = coverage.runs(memory)
    let lookupBytes = 0
    for (const c of runs) {
      if (lookupBytes) memory.add(-lookupBytes)
      // Two bounded column ranks need at most 44 transient UTF-16 bytes.
      // Map entries/retained keys have their own 64 B estimate, separate from
      // the unchanged 128 B block estimate retained by native clipboard jobs.
      memory.add(64); lookupBytes = 64
      const key = `${c.lo}:${c.hi}`, last = lastByColumns.get(key), rows = hi - lo + 1, cols = c.hi - c.lo + 1
      if (last && last.rows.hi + 1 === lo) { last.rows.hi = hi; lowerBytes += 2 * rows * cols }
      else {
        lowerBytes += 2 * (rows * (cols - 1) + rows - 1) + (blocks.length ? 4 : 0)
        if (lowerBytes > CLIPBOARD_BYTES || (blocks.length + 1) * 128 > METADATA_BYTES) throw limit()
        memory.add(128 + (last ? 0 : 64))
        if (!last) mapBytes += 64
        const block = { rows: { lo, hi }, cols: { ...c } }
        blocks.push(block); lastByColumns.set(key, block)
      }
    }
    memory.add(-lookupBytes)
    const runBytes = runs.length * 64
    runs.length = 0
    memory.add(-runBytes)
    if (lowerBytes > CLIPBOARD_BYTES) throw limit()
  }
  coverage = null
  lastByColumns.clear(); lastByColumns = null
  memory.add(-mapBytes)
  const eventBytes = events.length * 32
  events.length = 0
  memory.add(-32 * columnCount - eventBytes - columnBytes)
  sorted(blocks, (a, b) => a.rows.lo - b.rows.lo || a.cols.lo - b.cols.lo, memory)
  return blocks
}
function abortable<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
  return new Promise((resolve, reject) => {
    const abort = () => reject(new CopyFailure('cancelled', 'Copy cancelled.'))
    if (signal.aborted) { abort(); return }
    signal.addEventListener('abort', abort, { once: true })
    promise.then(resolve, reject).finally(() => signal.removeEventListener('abort', abort))
  })
}
async function read(input: CopyInput, row: number, col: number, memory: CopyMemory): Promise<{ text: string; owned: number }> {
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
    if (result.kind === 'blank') return { text: '', owned: 0 }
    if (result.kind === 'error') throw new CopyFailure(result.code, `Copy could not read a cell (${result.code}).`)
    if (result.cell.row !== row || result.cell.col !== col || result.cell.id !== cellId(row, col)) throw new CopyFailure('page-error', 'Copy source returned a different cell.')
    if (result.kind === 'unavailable' || result.cell.unavailable) throw new CopyFailure('formula-value-unavailable', 'Copy contains a formula whose saved value is unavailable.')
    let scratch = 0, pendingRelease = 0, borrowed = false
    const flushRelease = () => {
      if (pendingRelease) { memory.add(-pendingRelease); scratch -= pendingRelease; pendingRelease = 0 }
    }
    const allocation: FormatAllocation = {
      allocate: bytes => { flushRelease(); memory.add(bytes); scratch += bytes },
      // A builder's old argument remains an alias until its helper returns.
      // Defer refund to the next allocation (after assignment) or formatter exit.
      release: bytes => { pendingRelease += bytes },
      borrowed: text => { borrowed = text === result.cell.value },
      retain: (targets, bytes, commit) => {
        flushRelease()
        if (!input.retainBytes) { commit(); return }
        input.retainBytes(targets, bytes, commit)
        memory.bytes -= bytes
        scratch -= bytes
        input.onBytes?.(memory.bytes)
      },
    }
    const value = result.cell.value
    if (!input.format && typeof value !== 'string') allocation.allocate(64)
    const text = input.format ? input.format(result.cell, allocation) : String(value ?? '')
    // Legacy callbacks may ignore the context; their owned result is charged on
    // return. Production uses the canonical guarded formatter before allocation.
    flushRelease()
    const owned = typeof value === 'string' && text === value && (scratch === 0 || borrowed) ? 0 : text.length * 2
    memory.add(owned - scratch)
    return { text, owned }
  }
}
/** Count raw UTF-16 before replace/wrapper/concat can construct escaped text. */
function escapedUnits(text: string): number {
  let quotes = 0, quote = false
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (ch === '"') quotes++
    if (ch === '"' || ch === '\t' || ch === '\r' || ch === '\n') quote = true
  }
  return quote ? text.length + quotes + 2 : text.length
}
export async function serializeSelection(input: CopyInput): Promise<CopyResult> {
  const memory = new CopyMemory(input)
  let ready = false
  try {
    check(input)
    const blocks = blocksFor(input, memory), metadata = blocks.length * 128
    const lower = blocks.reduce((sum, block) => {
      const r = block.rows.hi - block.rows.lo + 1, c = block.cols.hi - block.cols.lo + 1
      return sum + 2 * (r * (c - 1) + r - 1)
    }, Math.max(0, blocks.length - 1) * 4)
    if (lower > CLIPBOARD_BYTES || lower * 2 + metadata > (input.sessionBudgetBytes ?? DEFAULT_COPY_SESSION_BYTES)) throw limit()
    const chunks: string[] = []
    let buffer = '', units = 0, batches = 0, slots = 0, formatted = 0
    const reserve = (nextUnits = units) => memory.set(nextUnits * 4 + metadata + slots + formatted + chunks.length * 8)
    const flush = () => {
      if (!buffer) return
      memory.add(8)
      chunks.push(buffer); buffer = ''
    }
    const append = (text: string, field = false) => {
      const length = field ? escapedUnits(text) : text.length, next = units + length
      if (next * 2 > CLIPBOARD_BYTES) throw limit()
      reserve(next)
      if (field && length !== text.length) buffer += `"${text.replace(/"/g, '""')}"`
      else buffer += text
      units = next
      if (buffer.length >= 65536) flush()
    }
    for (let b = 0; b < blocks.length; b++) {
      if (b) append('\n\n')
      const block = blocks[b], width = block.cols.hi - block.cols.lo + 1
      const cells = (block.rows.hi - block.rows.lo + 1) * width
      for (let offset = 0; offset < cells; offset += BATCH_CELLS) {
        check(input)
        const n = Math.min(BATCH_CELLS, cells - offset)
        const values: { text: string; owned: number }[] = []
        for (let i = 0; i < n; i++) {
          const index = offset + i, rank = block.rows.lo + Math.floor(index / width)
          const row = input.order.displayToSource(input.visibility.select(rank)!)
          const col = input.columns.select(block.cols.lo + index % width)!
          // Reference slot plus owned value record, before read can format.
          memory.add(32); slots += 32
          const value = await read(input, row, col, memory)
          formatted += value.owned
          if (value.text.length * 2 + units * 2 > CLIPBOARD_BYTES) throw limit()
          values.push(value)
        }
        check(input)
        for (let i = 0; i < n; i++) {
          const index = offset + i
          if (index) append(index % width ? '\t' : '\n')
          append(values[i].text, true)
          const owned = values[i].owned
          values[i].text = ''; values[i].owned = 0
          formatted -= owned
          reserve()
        }
        values.length = 0
        slots = 0
        reserve()
        if (++batches % 16 === 0) await abortable(new Promise<void>(resolve => setTimeout(resolve, 0)), input.signal)
      }
    }
    flush()
    check(input)
    reserve()
    const text = chunks.join('')
    chunks.length = 0
    memory.set(units * 4 + metadata)
    ready = true
    return { kind: 'ready', text, bytes: units * 2 }
  } catch (error) {
    if (error instanceof CopyFailure) {
      if (error.code === 'cancelled' || error.code === 'stale' || error.code === 'denied') return { kind: error.code }
      return { kind: 'refused', code: error.code, message: error.message }
    }
    return { kind: 'refused', code: 'page-error', message: error instanceof Error ? error.message : 'Copy unavailable; retry.' }
  } finally {
    // Failed serialization has no submitted endpoint; all local buffers have
    // become unreachable. A ready result stays charged until copy settlement.
    if (!ready && memory.bytes) memory.set(0)
  }
}
export async function copySelection(input: CopyInput): Promise<CopyResult | { kind: 'copied'; outcome: CopyOutcome }> {
  const result = await serializeSelection(input)
  if (result.kind !== 'ready') return result
  try { check(input) } catch (error) { return { kind: (error as CopyFailure).code as 'cancelled' | 'stale' | 'denied' } }
  const outcome = await copyText(result.text, { allowed: () => input.allowed() && sameIdentity(input.owner, input.currentOwner()) && sameStamp(input.owner, input.store.stamp), signal: input.signal })
  return outcome === 'failed' ? { kind: 'refused', code: 'clipboard-error', message: 'Copy failed. Try copying again.' } : { kind: 'copied', outcome }
}
