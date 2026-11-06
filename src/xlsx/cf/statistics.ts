/**
 * Measured range statistics shared across CF rules (Stage 1, Task 5).
 *
 * `prepareStatistics` scans each unique scope exactly once through any
 * `CfValueSource` (resident or paged-shaped), clamps whole-column spans to the
 * source's populated bounds, and memoises the bundle under
 * (sheetId, normalised scope, hidden policy, value generation) in a bounded
 * LRU. Every allocation is admitted before it is built; aggregate work is
 * spent from the shared `CfWorkBudget` (refusal withholds, never publishes
 * partial results) and cancellation aborts the scan without caching anything.
 *
 * Lease model: every returned lease carries its own admitted charge for the
 * shared value. The cache holds no separate charge — dropping the last lease
 * returns the account to zero, while an evicted-but-leased value stays alive
 * (and charged) through its live leases.
 */
import { PackageError } from '../../core/package-error'
import type {
  CfLease,
  CfLiteral,
  CfOutcome,
  CfStamp,
  CfValuePage,
  CfValueSource,
  SourceRect,
} from './model'
import type { XlsxCell } from '../types'
import type { CfMemory, CfWorkBudget } from './limits'

/** Stable, type-tagged frequency key. Blanks and errors are never counted. */
export function cfValueKey(value: CfLiteral): string | undefined {
  switch (value.kind) {
    case 'number': return `n:${value.value}`
    case 'string': return `s:${value.value.toLowerCase()}`
    case 'boolean': return `b:${value.value ? 1 : 0}`
    case 'blank':
    case 'error':
      return undefined
  }
}

export interface CfRangeStatistics {
  readonly scopeKey: string
  readonly numericCount: number
  readonly sum: number
  readonly mean: number | undefined
  readonly min: number | undefined
  readonly max: number | undefined
  readonly stddevPop: number | undefined
  /** Ascending numeric values; retained for top/bottom cutoffs. */
  readonly sortedAsc: Float64Array
  /** Frequency by `cfValueKey`; blanks and errors are never counted. */
  readonly frequencies: ReadonlyMap<string, number>
  readonly spanCells: number
  readonly storedCells: number
}

export interface CfStatisticsOptions {
  /** Rows hidden by metadata/filter to exclude ('include' mirrors Excel AVERAGE). */
  readonly hidden?: 'include' | 'exclude'
  readonly hiddenRows?: readonly SourceRect[]
  readonly signal?: AbortSignal
  readonly maxEntries?: number
}

export interface CfStatisticsSession {
  prepareStatistics(
    ruleScopes: readonly (readonly SourceRect[])[],
    stamp: CfStamp,
    source: CfValueSource,
    budget: CfWorkBudget,
    options?: Omit<CfStatisticsOptions, 'maxEntries'>,
  ): Promise<Array<CfOutcome<CfLease<CfRangeStatistics>>>>
  readonly size: number
  dispose(): void
}

const DEFAULT_MAX_ENTRIES = 64

function cancelled(): PackageError {
  return new PackageError('cancelled', 'CF statistics preparation cancelled')
}

function scopeKey(rects: readonly SourceRect[]): string {
  return [...rects]
    .map(r => `${r.firstRow},${r.lastRow},${r.firstCol},${r.lastCol}`)
    .sort()
    .join(';')
}

function hiddenKey(policy: 'include' | 'exclude', hiddenRows: readonly SourceRect[]): string {
  if (policy === 'include' || hiddenRows.length === 0) return policy
  return `${policy}@${scopeKey(hiddenRows)}`
}

/** Worst-case workspace for a span: two numeric arrays, plus dedup slots. */
export function estimateStatisticsBytes(spanCells: number, needsDedup: boolean): number {
  return 256 + spanCells * 16 + (needsDedup ? spanCells * 64 : 0)
}

function rectsOverlap(a: SourceRect, b: SourceRect): boolean {
  return a.firstRow <= b.lastRow && b.firstRow <= a.lastRow && a.firstCol <= b.lastCol && b.firstCol <= a.lastCol
}

function clampRect(rect: SourceRect, bounds: SourceRect): SourceRect | undefined {
  const firstRow = Math.max(rect.firstRow, bounds.firstRow)
  const lastRow = Math.min(rect.lastRow, bounds.lastRow)
  const firstCol = Math.max(rect.firstCol, bounds.firstCol)
  const lastCol = Math.min(rect.lastCol, bounds.lastCol)
  if (firstRow > lastRow || firstCol > lastCol) return undefined
  return { firstRow, lastRow, firstCol, lastCol }
}

function rectArea(rect: SourceRect): number {
  return (rect.lastRow - rect.firstRow + 1) * (rect.lastCol - rect.firstCol + 1)
}

function inRects(row: number, col: number, rects: readonly SourceRect[]): boolean {
  for (const r of rects) {
    if (row >= r.firstRow && row <= r.lastRow && col >= r.firstCol && col <= r.lastCol) return true
  }
  return false
}

function isCancel(error: unknown): boolean {
  return error instanceof PackageError && error.code === 'cancelled'
}

function isBudgetRefusal(error: unknown): boolean {
  return (
    error instanceof PackageError &&
    (error.reason === 'memory-budget' || (error.code === 'resource-limit' && error.reason === undefined))
  )
}

/**
 * Cooperative stable mergesort over an admitted scratch buffer.
 * Comparisons are work units; the caller yields at checkpoints.
 */
async function cooperativeSort(
  values: Float64Array,
  scratch: Float64Array,
  budget: CfWorkBudget,
  signal: AbortSignal,
  yieldTurn: () => Promise<void>,
): Promise<void> {
  const n = values.length
  if (n < 2) return
  let pending = 0
  const check = async (): Promise<void> => {
    if (signal.aborted) throw cancelled()
    const r = budget.spend(pending)
    pending = 0
    if (r === 'refused') {
      if (signal.aborted) throw cancelled()
      throw new PackageError('resource-limit', 'CF statistics sort exceeded the shared work budget')
    }
    if (r === 'yield') {
      await yieldTurn()
      if (signal.aborted) throw cancelled()
      budget.resume()
    }
  }
  for (let width = 1; width < n; width *= 2) {
    for (let lo = 0; lo < n; lo += 2 * width) {
      const mid = Math.min(lo + width, n)
      const hi = Math.min(lo + 2 * width, n)
      let i = lo, j = mid, k = lo
      while (i < mid && j < hi) {
        if (++pending >= 64) await check()
        if (values[j]! < values[i]!) scratch[k++] = values[j++]!
        else scratch[k++] = values[i++]!
      }
      while (i < mid) scratch[k++] = values[i++]!
      while (j < hi) scratch[k++] = values[j++]!
      values.set(scratch.subarray(lo, hi), lo)
    }
  }
  if (pending > 0) await check()
}

interface CacheEntry {
  stats: CfRangeStatistics
  bytes: number
  refs: number
}

interface BuiltStats {
  stats: CfRangeStatistics
  bytes: number
}

export function createCfStatistics(memory: CfMemory, options: { maxEntries?: number } = {}): CfStatisticsSession {
  const maxEntries = options.maxEntries ?? DEFAULT_MAX_ENTRIES
  const cache = new Map<string, CacheEntry>()

  function releaseRef(key: string): void {
    const entry = cache.get(key)
    if (entry === undefined) return
    entry.refs--
    if (entry.refs <= 0) cache.delete(key)
  }

  function evict(): void {
    // LRU: Map preserves insertion order; the first key is the least recently used.
    // Eviction drops the key only; live caller leases keep the value alive and charged.
    for (const key of cache.keys()) {
      cache.delete(key)
      if (cache.size <= maxEntries) break
    }
  }

  function dispose(): void {
    cache.clear()
  }

  function adoptStats(key: string, stats: CfRangeStatistics, bytes: number): CfOutcome<CfLease<CfRangeStatistics>> {
    const reservation = memory.reserve(bytes)
    // A refused outcome holds nothing, so the cache is untouched.
    if (reservation.kind === 'unresolved') return { kind: 'unresolved', reason: 'cf-resource-limit' }
    const adopted = reservation.adopt(stats, bytes)
    let entry = cache.get(key)
    if (entry === undefined) {
      entry = { stats, bytes, refs: 0 }
      cache.set(key, entry)
      if (cache.size > maxEntries) evict()
    }
    entry.refs++
    let settled = false
    const releaseOnce = (): void => {
      if (settled) return
      settled = true
      try { adopted.release() } catch { /* release-only path */ }
      releaseRef(key)
    };
    const lease: CfLease<CfRangeStatistics> = {
      value: stats,
      bytes,
      release: releaseOnce,
      transfer: (destination) => {
        const incoming = destination.reserve(bytes)
        if (incoming.kind === 'unresolved') return incoming
        let moved: CfLease<CfRangeStatistics>
        try {
          moved = incoming.adopt(stats) as CfLease<CfRangeStatistics>
        } catch (error) {
          try { incoming.release() } catch { /* release-only path */ }
          throw error
        }
        releaseOnce()
        return { kind: 'ready', value: moved }
      },
    }
    return { kind: 'ready', value: lease }
  }

  async function build(
    clamped: readonly SourceRect[],
    spanCells: number,
    needsDedup: boolean,
    stamp: CfStamp,
    source: CfValueSource,
    budget: CfWorkBudget,
    signal: AbortSignal,
    hidden: 'include' | 'exclude',
    hiddenRows: readonly SourceRect[],
    key: string,
  ): Promise<CfOutcome<BuiltStats>> {
    // Reserve the worst-case numeric workspace before building any array.
    const workspace = estimateStatisticsBytes(spanCells, needsDedup)
    const hold = memory.reserve(workspace)
    if (hold.kind === 'unresolved') return { kind: 'unresolved', reason: 'cf-resource-limit' }
    try {
      if (signal.aborted) throw cancelled()
      const values = new Float64Array(spanCells)
      const scratch = new Float64Array(spanCells)
      const seen = needsDedup ? new Set<number>() : undefined
      const frequencies = new Map<string, number>()
      let sum = 0
      let sumSq = 0
      let min = Infinity
      let max = -Infinity
      let stored = 0
      let numericCount = 0
      let pendingSpend = 0
      const spend = async (units: number): Promise<boolean> => {
        if (signal.aborted) throw cancelled()
        pendingSpend += units
        if (pendingSpend < 64) return true
        const r = budget.spend(pendingSpend)
        pendingSpend = 0
        if (r === 'refused') {
          if (signal.aborted) throw cancelled()
          return false
        }
        if (r === 'yield') {
          await Promise.resolve()
          if (signal.aborted) throw cancelled()
          budget.resume()
        }
        return true
      }
      for (const rect of clamped) {
        const pages = source.pages(stamp.sheetId, [rect], budget, signal)
        for await (const page of pages) {
          try {
            const rows = page.value.rows
            const cols = page.value.cols
            const items = page.value.values
            for (let i = 0; i < items.length; i++) {
              const row = rows[i]!
              const col = cols[i]!
              stored++
              if (hidden === 'exclude' && inRects(row, col, hiddenRows)) continue
              const value = items[i]!
              if (value.kind === 'blank') continue
              if (!(await spend(1))) return { kind: 'unresolved', reason: 'cf-resource-limit' }
              if (value.kind === 'number') {
                if (seen !== undefined) {
                  const id = row * 16384 + col
                  if (seen.has(id)) continue
                  seen.add(id)
                }
                values[numericCount++] = value.value
                sum += value.value
                sumSq += value.value * value.value
                if (value.value < min) min = value.value
                if (value.value > max) max = value.value
              } else if (value.kind === 'string' || value.kind === 'boolean') {
                const k = cfValueKey(value)!
                const prior = frequencies.get(k)
                if (prior === undefined) {
                  // Admit the new key, entry slot and worst-case string before retaining.
                  const textLen = value.kind === 'string' ? value.value.length : 0
                  if (!hold.grow(hold.bytes + 64 + 16 + textLen * 2)) {
                    return { kind: 'unresolved', reason: 'cf-resource-limit' }
                  }
                  frequencies.set(k, 1)
                } else {
                  frequencies.set(k, prior + 1)
                }
                if (!(await spend(1))) return { kind: 'unresolved', reason: 'cf-resource-limit' }
                if (value.kind === 'string') {
                  if (!(await spend(1 + (value.value.length >>> 3)))) return { kind: 'unresolved', reason: 'cf-resource-limit' }
                }
              }
              // Error values are never counted: not numeric, not frequency keys.
            }
          } finally {
            try { page.release() } catch { /* release-only path */ }
          }
        }
      }
      if (signal.aborted) throw cancelled()
      if (pendingSpend > 0) {
        const tail = budget.spend(pendingSpend)
        pendingSpend = 0
        if (tail === 'refused') {
          if (signal.aborted) throw cancelled()
          return { kind: 'unresolved', reason: 'cf-resource-limit' }
        }
        if (tail === 'yield') {
          await Promise.resolve()
          if (signal.aborted) throw cancelled()
          budget.resume()
        }
      }
      // Cooperative ascending sort for top/bottom cutoffs; raises budget refusal, never partial state.
      await cooperativeSort(values.subarray(0, numericCount), scratch.subarray(0, numericCount), budget, signal, () => Promise.resolve())
      if (signal.aborted) throw cancelled()
      const mean = numericCount > 0 ? sum / numericCount : undefined
      const variance = numericCount > 0 ? Math.max(0, sumSq / numericCount - (mean ?? 0) * (mean ?? 0)) : undefined
      const retained = values.slice(0, numericCount)
      let freqBytes = 0
      for (const [k, v] of frequencies) freqBytes += 64 + k.length * 2 + 8 + `${v}`.length * 2
      const bytes = 256 + retained.length * 8 + freqBytes
      const stats: CfRangeStatistics = {
        scopeKey: key,
        numericCount,
        sum,
        mean,
        min: numericCount > 0 ? min : undefined,
        max: numericCount > 0 ? max : undefined,
        stddevPop: variance === undefined ? undefined : Math.sqrt(variance),
        sortedAsc: retained,
        frequencies,
        spanCells,
        storedCells: stored,
      }
      return { kind: 'ready', value: { stats, bytes } }
    } catch (error) {
      if (isCancel(error) || signal.aborted) throw cancelled()
      if (isBudgetRefusal(error)) return { kind: 'unresolved', reason: 'cf-resource-limit' }
      throw error
    } finally {
      try { hold.release() } catch { /* release-only path */ }
    }
  }

  async function prepareStatistics(
    ruleScopes: readonly (readonly SourceRect[])[],
    stamp: CfStamp,
    source: CfValueSource,
    budget: CfWorkBudget,
    options: Omit<CfStatisticsOptions, 'maxEntries'> = {},
  ): Promise<Array<CfOutcome<CfLease<CfRangeStatistics>>>> {
    const signal = options.signal ?? new AbortController().signal
    const hidden = options.hidden ?? 'include'
    const hiddenRows = options.hiddenRows ?? []
    const bounds = source.bounds(stamp.sheetId)
    const out: Array<CfOutcome<CfLease<CfRangeStatistics>>> = []
    const pending = new Map<string, Promise<CfOutcome<BuiltStats>>>()
    for (const scope of ruleScopes) {
      if (signal.aborted) throw cancelled()
      const key = `${stamp.sheetId}|${stamp.workbookValueGen}|${scopeKey(scope)}|${hiddenKey(hidden, hiddenRows)}`
      const hit = cache.get(key)
      if (hit !== undefined) {
        // Warm hit: refresh recency and share the value. No scan, no loops.
        cache.delete(key)
        cache.set(key, hit)
        out.push(adoptStats(key, hit.stats, hit.bytes))
        continue
      }
      let job = pending.get(key)
      if (job === undefined) {
        job = (async (): Promise<CfOutcome<BuiltStats>> => {
          if (bounds === undefined) {
            return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: [stamp.sheetId] }
          }
          const clamped: SourceRect[] = []
          let spanCells = 0
          for (const r of scope) {
            const c = clampRect(r, bounds)
            if (c !== undefined) {
              clamped.push(c)
              spanCells += rectArea(c)
            }
          }
          if (clamped.length === 0 || spanCells === 0) {
            return {
              kind: 'ready',
              value: {
                stats: {
                  scopeKey: key, numericCount: 0, sum: 0, mean: undefined, min: undefined, max: undefined,
                  stddevPop: undefined, sortedAsc: new Float64Array(0), frequencies: new Map(),
                  spanCells: 0, storedCells: 0,
                },
                bytes: 256,
              },
            }
          }
          let needsDedup = false
          for (let a = 0; a < clamped.length && !needsDedup; a++) {
            for (let b = a + 1; b < clamped.length; b++) {
              if (rectsOverlap(clamped[a]!, clamped[b]!)) { needsDedup = true; break }
            }
          }
          return build(clamped, spanCells, needsDedup, stamp, source, budget, signal, hidden, hiddenRows, key)
        })()
        pending.set(key, job)
      }
      const built = await job
      if (built.kind !== 'ready') {
        out.push(built)
        continue
      }
      out.push(adoptStats(key, built.value.stats, built.value.bytes))
    }
    pending.clear()
    return out
  }

  return { prepareStatistics, get size() { return cache.size }, dispose }
}

/** Resident adapter: array-backed cells behind the shared CfValueSource contract. */
export function createResidentCfSource(
  cells: readonly XlsxCell[],
  sheetId: string,
  stamp: CfStamp,
  memory: CfMemory,
): CfValueSource {
  const byRow = new Map<number, XlsxCell[]>()
  for (const cell of cells) {
    const list = byRow.get(cell.row) ?? []
    list.push(cell)
    byRow.set(cell.row, list)
  }
  let minRow = Infinity, maxRow = -Infinity, minCol = Infinity, maxCol = -Infinity
  for (const cell of cells) {
    if (cell.row < minRow) minRow = cell.row
    if (cell.row > maxRow) maxRow = cell.row
    if (cell.col < minCol) minCol = cell.col
    if (cell.col > maxCol) maxCol = cell.col
  }
  const literalBytes = (value: CfLiteral): number => {
    switch (value.kind) {
      case 'number': return 8
      case 'boolean': return 8
      case 'blank': return 8
      case 'string': return 16 + value.value.length * 2
      case 'error': return 16 + value.value.length * 2
    }
  }
  const toLiteral = (cell: XlsxCell): CfLiteral => {
    if (cell.value === null) return { kind: 'blank' }
    if (cell.valueIsError) return { kind: 'error', value: String(cell.value) }
    if (typeof cell.value === 'number') return { kind: 'number', value: cell.value }
    if (typeof cell.value === 'boolean') return { kind: 'boolean', value: cell.value }
    return { kind: 'string', value: cell.value }
  }
  return {
    bounds: (_id: string) => {
      if (cells.length === 0) return undefined
      return { firstRow: minRow, lastRow: maxRow, firstCol: minCol, lastCol: maxCol }
    },
    stamp: (_id: string) => stamp,
    pages: async function* (id: string, rects: readonly SourceRect[], budget: CfWorkBudget, signal: AbortSignal) {
      if (id !== sheetId) throw new Error('CF resident source queried for the wrong sheet')
      for (const rect of rects) {
        if (signal.aborted) throw cancelled()
        const rows: number[] = []
        const cols: number[] = []
        const values: CfLiteral[] = []
        for (const [row, list] of byRow) {
          if (row < rect.firstRow || row > rect.lastRow) continue
          for (const cell of list) {
            if (cell.col < rect.firstCol || cell.col > rect.lastCol) continue
            rows.push(row)
            cols.push(cell.col)
            values.push(toLiteral(cell))
          }
        }
        const spent = budget.spend(values.length)
        if (spent === 'refused') {
          if (signal.aborted) throw cancelled()
          throw new PackageError('resource-limit', 'CF resident page exceeded the shared work budget')
        }
        if (spent === 'yield') {
          await Promise.resolve()
          if (signal.aborted) throw cancelled()
          budget.resume()
        }
        const bytes = 128 + rows.length * 4 + cols.length * 2 + values.reduce((n, v) => n + literalBytes(v), 0)
        const hold = memory.reserve(bytes)
        if (hold.kind === 'unresolved') {
          throw new PackageError('resource-limit', 'CF resident page admission refused', undefined, 'memory-budget')
        }
        yield hold.adopt<CfValuePage>({ rows: new Uint32Array(rows), cols: new Uint16Array(cols), values }, bytes)
      }
    },
  }
}
