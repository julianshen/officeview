/**
 * Measured range statistics shared across CF rules (Stage 1, Task 5, fix round 1).
 *
 * `prepareStatistics` scans each unique scope exactly once through any
 * `CfValueSource` (resident or paged-shaped), clamps whole-column spans to the
 * source's populated bounds, and memoises the bundle under
 * (sheetId, owner, structure generation, value generation, normalised scope,
 * hidden policy) in a bounded LRU. Every allocation is admitted before it is
 * built; aggregate work is spent from the shared `CfWorkBudget` (refusal
 * withholds, never publishes partial results) and cancellation aborts the scan
 * without caching anything.
 *
 * Ownership: each cache entry adopts exactly one charge for its retained
 * footprint; per-scope outcomes are lightweight wrappers sharing that charge
 * through a refcount. Dropping the last live lease releases the entry, and
 * eviction drops only the key while live leases keep the graph alive and
 * charged. Concurrent preparations for the same key coalesce onto one build;
 * waiters observe their own abort promptly without disturbing the shared job.
 */
import { PackageError } from '../../core/package-error'
import { hostTask } from '../view/scheduler'
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
import { CF_LIMITS } from './limits'

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
  /** Distinct declared cells (union area of the scope rectangles). */
  readonly declaredCells: number
  /** Distinct cells inside the populated clamp actually scanned. */
  readonly spanCells: number
  /** Distinct stored cells visited (post overlap-dedup). */
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
/** Initial numeric capacity per scan; arrays grow additively with admission. */
const INITIAL_NUMERIC_CAPACITY = 1024
const NUMERIC_GROW_CHUNK = 4096

/** Initial workspace for a span: base record, one values/scratch pair, headroom. */
export function estimateStatisticsInitialBytes(spanCells: number): number {
  return 256 + Math.min(Math.max(spanCells, 1), INITIAL_NUMERIC_CAPACITY) * 8 * 2 + 256
}

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

/** Cooperative insertion sort: in-place, no allocation, spending one unit
 * per comparison and per shift before performing it, so large inputs
 * refuse through the shared budget instead of running uncounted. The
 * batched spender keeps checkpoints within the 64-unit granularity. */
async function cooperativeInsertionSort<T>(
  values: T[],
  compare: (a: T, b: T) => number,
  spend: SpendState,
): Promise<boolean> {
  for (let i = 1; i < values.length; i++) {
    const current = values[i]!
    let j = i - 1
    while (j >= 0) {
      if (!(await spendUnits(spend, 1))) return false
      if (compare(current, values[j]!) >= 0) break
      if (!(await spendUnits(spend, 1))) return false
      values[j + 1] = values[j]
      j--
    }
    values[j + 1] = current
  }
  return true
}

/** Exact union area via sweep-line: exact for any rectangle count, with no
 * naive-sum fallback. Cooperative: spends per examined rectangle pair,
 * per sort comparison and shift, and per merge iteration — all before or
 * during execution, never as post-hoc bulk charges — so large scopes stay
 * interruptible, budgeted within the checkpoint granularity, and refused
 * rather than published over budget. The caller admits the simultaneous
 * workspace (edges, coordinate list, interval arrays, sort scratch) and
 * tracks it in the session cap; this function only sweeps. Returns
 * undefined when the shared budget refuses. */
async function sweepUnion(
  rects: readonly SourceRect[],
  spend: SpendState,
): Promise<number | undefined> {
  const cleaned = rects.filter(r => r.firstRow <= r.lastRow && r.firstCol <= r.lastCol)
  if (cleaned.length === 0) return 0
  if (cleaned.length === 1) return rectArea(cleaned[0]!)
  const xs = new Set<number>()
  for (const r of cleaned) {
    if (!(await spendUnits(spend, 1))) return undefined
    xs.add(r.firstCol)
    xs.add(r.lastCol + 1)
  }
  if (!(await flushSpend(spend))) return undefined
  const xList: number[] = [...xs]
  if (!(await cooperativeInsertionSort(xList, (a, b) => a - b, spend))) return undefined
  if (!(await flushSpend(spend))) return undefined
  let total = 0
  const iv: Array<[number, number]> = []
  for (let xi = 0; xi < xList.length - 1; xi++) {
    if (!(await spendUnits(spend, 1))) return undefined
    const x0 = xList[xi]!
    const x1 = xList[xi + 1]!
    iv.length = 0
    for (const r of cleaned) {
      // Each inner rectangle examination is shared work: spend, checkpoint
      // cancellation, and yield cooperatively through the batched spender.
      if (!(await spendUnits(spend, 1))) return undefined
      if (r.firstCol <= x0 && r.lastCol + 1 >= x1) iv.push([r.firstRow, r.lastRow + 1])
    }
    if (iv.length === 0) continue
    if (!(await cooperativeInsertionSort(iv, (a, b) => a[0] - b[0] || a[1] - b[1], spend))) return undefined
    let cs = iv[0]![0]
    let ce = iv[0]![1]
    for (let j = 1; j < iv.length; j++) {
      if (!(await spendUnits(spend, 1))) return undefined
      const s = iv[j]![0]
      const e = iv[j]![1]
      if (s <= ce) { if (e > ce) ce = e }
      else { total += (x1 - x0) * (ce - cs); cs = s; ce = e }
    }
    total += (x1 - x0) * (ce - cs)
    if (!(await flushSpend(spend))) return undefined
  }
  if (!(await flushSpend(spend))) return undefined
  return total
}

function inRects(row: number, col: number, rects: readonly SourceRect[]): boolean {
  for (const r of rects) {
    if (row >= r.firstRow && row <= r.lastRow && col >= r.firstCol && col <= r.lastCol) return true
  }
  return false
}

function isCancel(error: unknown): boolean {
  return (
    error instanceof PackageError && error.code === 'cancelled'
  )
}

function isBudgetRefusal(error: unknown): boolean {
  return (
    error instanceof PackageError &&
    (error.reason === 'memory-budget' || (error.code === 'resource-limit' && error.reason === undefined))
  )
}

function stampsMatch(live: CfStamp, requested: CfStamp): boolean {
  return (
    live.sheetId === requested.sheetId &&
    live.owner === requested.owner &&
    live.structureGen === requested.structureGen &&
    live.workbookValueGen === requested.workbookValueGen
  )
}

interface SpendState {
  pending: number
  signal: AbortSignal
  budget: CfWorkBudget
}

async function spendUnits(state: SpendState, units: number): Promise<boolean> {
  if (state.signal.aborted) throw cancelled()
  state.pending += units
  if (state.pending < 64) return true
  const spent = state.budget.spend(state.pending)
  state.pending = 0
  if (spent === 'refused') {
    if (state.signal.aborted) throw cancelled()
    return false
  }
  if (spent === 'yield') {
    await hostTask()
    if (state.signal.aborted) throw cancelled()
    state.budget.resume()
  }
  return true
}

async function flushSpend(state: SpendState): Promise<boolean> {
  if (state.pending === 0) return true
  // Force-spend the remainder instead of dropping tails below the batch
  // threshold; tiny preparations must still count their work.
  const units = state.pending
  state.pending = 0
  const spent = state.budget.spend(units)
  if (spent === 'refused') {
    if (state.signal.aborted) throw cancelled()
    return false
  }
  if (spent === 'yield') {
    await hostTask()
    if (state.signal.aborted) throw cancelled()
    state.budget.resume()
  }
  return true
}

/** Copy with bounded work slices so large copies stay interruptible. */
async function copyInto(
  dst: Float64Array,
  dstOffset: number,
  src: Float64Array | ArrayLike<number>,
  length: number,
  state: SpendState,
): Promise<boolean> {
  for (let i = 0; i < length; i++) {
    dst[dstOffset + i] = src[i]!
    if ((i & 63) === 63 && !(await spendUnits(state, 64))) return false
  }
  return flushSpend(state)
}

/**
 * Cooperative stable mergesort over admitted buffers. Comparisons are work
 * units; checkpoints yield to the host so cancellation is observed promptly.
 */
async function cooperativeSort(
  values: Float64Array,
  scratch: Float64Array,
  length: number,
  state: SpendState,
): Promise<boolean> {
  for (let width = 1; width < length; width *= 2) {
    for (let lo = 0; lo < length; lo += 2 * width) {
      const mid = Math.min(lo + width, length)
      const hi = Math.min(lo + 2 * width, length)
      let i = lo, j = mid, k = lo
      while (i < mid && j < hi) {
        // Comparisons are work units (load, compare, branch, counter).
        if (!(await spendUnits(state, 4))) return false
        if (values[j]! < values[i]!) scratch[k++] = values[j++]!
        else scratch[k++] = values[i++]!
      }
      while (i < mid) {
        scratch[k++] = values[i++]!
        if ((k & 63) === 0 && !(await spendUnits(state, 64))) return false
      }
      while (j < hi) {
        scratch[k++] = values[j++]!
        if ((k & 63) === 0 && !(await spendUnits(state, 64))) return false
      }
      for (let t = lo; t < hi; t++) {
        values[t] = scratch[t]!
        if (((t - lo) & 63) === 63 && !(await spendUnits(state, 64))) return false
      }
      if (state.signal.aborted) throw cancelled()
    }
  }
  return flushSpend(state)
}

interface CacheEntry {
  stats: CfRangeStatistics
  bytes: number
  refs: number
  lease: CfLease<CfRangeStatistics>
}

interface Waiter {
  done: boolean
  signal: AbortSignal
  onAbort: () => void
  resolve: (o: CfOutcome<CfLease<CfRangeStatistics>>) => void
  reject: (e: unknown) => void
}

interface InflightJob {
  waiters: Set<Waiter>
  controller: AbortController
  build: Promise<BuiltOutcome>
  abort(): void
}

type BuiltOutcome =
  | { kind: 'ready'; value: { stats: CfRangeStatistics; bytes: number; candidate: CfLease<CfRangeStatistics> } }
  | { kind: 'unresolved'; reason: string; pendingSheetIds?: readonly string[] }

export function createCfStatistics(memory: CfMemory, options: { maxEntries?: number } = {}): CfStatisticsSession {
  const maxEntries = options.maxEntries ?? DEFAULT_MAX_ENTRIES
  const sessionCap = CF_LIMITS.statisticsBytes
  const cache = new Map<string, CacheEntry>()
  const inflight = new Map<string, InflightJob>()
  const workspaces = new Map<object, number>()
  // Populated-bounds memo per sheet generation: bounds() rescans the whole
  // sheet, so warm cache hits must not touch it at all, and concurrent
  // misses in one generation share a single scan.
  const boundsMemo = new Map<string, SourceRect | undefined>()
  let retainedHeld = 0

  const withinCap = (additional: number): number => {
    let total = retainedHeld + additional
    for (const [, bytes] of workspaces) total += bytes
    return total
  }
  const checkCap = (additional: number): boolean => withinCap(additional) <= sessionCap

  function trackRelease(bytes: number): void {
    retainedHeld -= bytes
    if (retainedHeld < 0) retainedHeld = 0
  }

  function evict(): void {
    // LRU: Map preserves insertion order; the first key is the least recently used.
    // Eviction drops the key only; live leases keep the shared charge.
    for (const key of cache.keys()) {
      cache.delete(key)
      if (cache.size <= maxEntries) break
    }
  }

  function dispose(): void {
    cache.clear()
    boundsMemo.clear()
  }

  function dropEntry(key: string, entry: CacheEntry): void {
    if (cache.get(key) === entry) cache.delete(key)
    try { entry.lease.release() } catch { /* release-only path */ }
    trackRelease(entry.bytes)
  }

  function makeOutcome(key: string, entry: CacheEntry): CfOutcome<CfLease<CfRangeStatistics>> {
    entry.refs++
    let settled = false
    const release = (): void => {
      if (settled) return
      settled = true
      entry.refs--
      if (entry.refs <= 0) dropEntry(key, entry)
    }
    const lease: CfLease<CfRangeStatistics> = {
      value: entry.stats,
      bytes: entry.bytes,
      release,
      transfer: (destination) => {
        const incoming = destination.reserve(entry.bytes)
        if (incoming.kind === 'unresolved') return incoming
        let moved: CfLease<CfRangeStatistics>
        try {
          moved = incoming.adopt(entry.stats) as CfLease<CfRangeStatistics>
        } catch (error) {
          try { incoming.release() } catch { /* release-only path */ }
          throw error
        }
        release()
        return { kind: 'ready', value: moved }
      },
    }
    return { kind: 'ready', value: lease }
  }

  async function build(
    clamped: readonly SourceRect[],
    spanCells: number,
    declaredCells: number,
    needsDedup: boolean,
    stamp: CfStamp,
    source: CfValueSource,
    budget: CfWorkBudget,
    signal: AbortSignal,
    hidden: 'include' | 'exclude',
    hiddenRows: readonly SourceRect[],
    key: string,
  ): Promise<BuiltOutcome> {
    // Admit the initial workspace before building anything.
    const initial = 256 + Math.min(Math.max(spanCells, 1), INITIAL_NUMERIC_CAPACITY) * 8 * 2 + 256
    if (!checkCap(initial)) {
      return { kind: 'unresolved', reason: 'cf-resource-limit' }
    }
    const hold = memory.reserve(initial)
    if (hold.kind === 'unresolved') {
      return { kind: 'unresolved', reason: 'cf-resource-limit' }
    }
    workspaces.set(hold, initial)
    let adoptedLease: CfLease<CfRangeStatistics> | undefined
    const growHold = (delta: number): boolean => {
      // The session cap binds the full new total: every tracked workspace
      // hold (including this one) plus retained plus the increment.
      if (!checkCap(delta)) return false
      if (!hold.grow(hold.bytes + delta)) return false
      workspaces.set(hold, hold.bytes)
      return true
    }
    try {
      if (signal.aborted) throw cancelled()
      let capacity = Math.min(Math.max(spanCells, 1), INITIAL_NUMERIC_CAPACITY)
      let values = new Float64Array(capacity)
      let scratch = new Float64Array(capacity)
      const seen = needsDedup ? new Set<number>() : undefined
      const frequencies = new Map<string, number>()
      // Neumaier compensated summation: exact for cancellation patterns.
      let sum = 0
      let compensation = 0
      let numericCount = 0
      let min = Infinity
      let max = -Infinity
      let stored = 0
      const spend: SpendState = { pending: 0, signal, budget }
      const ensureCapacity = async (needed: number): Promise<boolean> => {
        if (needed <= values.length) return true
        let next = values.length
        while (next < needed) next += NUMERIC_GROW_CHUNK
        const delta = (next - values.length) * 8 * 2
        // Admit the replacement peak: both old arrays stay referenced
        // across the awaited copies, so the hold must cover old + new.
        // The slack persists in the hold (conservative) and is freed when
        // the workspace is adopted or released.
        if (!growHold(delta + values.byteLength + scratch.byteLength)) return false
        const grown = new Float64Array(next)
        const grownScratch = new Float64Array(next)
        if (!(await copyInto(grown, 0, values.subarray(0, numericCount), numericCount, spend))) return false
        if (!(await copyInto(grownScratch, 0, scratch.subarray(0, numericCount), numericCount, spend))) return false
        values = grown
        scratch = grownScratch
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
              if (!(await spendUnits(spend, 1))) {
                return { kind: 'unresolved', reason: 'cf-resource-limit' }
              }
              if (hidden === 'exclude' && inRects(row, col, hiddenRows)) continue
              const value = items[i]!
              // Overlap dedup covers every stored kind including explicit
              // blanks, so overlapping rects count each position once.
              // Each new id grows the dedup set and is admitted first.
              if (needsDedup) {
                const id = row * 16384 + col
                if (seen!.has(id)) continue
                if (!growHold(64)) {
                  return { kind: 'unresolved', reason: 'cf-resource-limit' }
                }
                seen!.add(id)
              }
              // Distinct, non-excluded stored cells (blanks included).
              stored++
              if (value.kind === 'blank') continue
              if (value.kind === 'number') {
                if (!(await ensureCapacity(numericCount + 1))) {
                  return { kind: 'unresolved', reason: 'cf-resource-limit' }
                }
                values[numericCount++] = value.value
                // Neumaier update.
                const t = sum + value.value
                if (Math.abs(sum) >= Math.abs(value.value)) compensation += sum - t + value.value
                else compensation += value.value - t + sum
                sum = t
                if (value.value < min) min = value.value
                if (value.value > max) max = value.value
                // Duplicate detection counts numbers too: admit the slot and
                // the short 'n:<digits>' key before constructing it.
                const nkey = 'n:' + String(value.value)
                if (!frequencies.has(nkey)) {
                  if (!growHold(128)) {
                    return { kind: 'unresolved', reason: 'cf-resource-limit' }
                  }
                }
                frequencies.set(nkey, (frequencies.get(nkey) ?? 0) + 1)
              } else if (value.kind === 'string' || value.kind === 'boolean') {
                // Admit the key text before constructing it.
                const textLen = value.kind === 'string' ? value.value.length : 0
                if (!growHold(64 + textLen * 2 + 16)) {
                  return { kind: 'unresolved', reason: 'cf-resource-limit' }
                }
                const k = cfValueKey(value)!
                const prior = frequencies.get(k)
                if (prior === undefined) {
                  if (!growHold(64)) {
                    return { kind: 'unresolved', reason: 'cf-resource-limit' }
                  }
                  frequencies.set(k, 1)
                } else {
                  frequencies.set(k, prior + 1)
                }
                if (!(await spendUnits(spend, 2))) {
                  return { kind: 'unresolved', reason: 'cf-resource-limit' }
                }
                if (value.kind === 'string') {
                  if (!(await spendUnits(spend, 1 + (value.value.length >>> 3)))) {
                    return { kind: 'unresolved', reason: 'cf-resource-limit' }
                  }
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
      if (!(await flushSpend(spend))) {
        return { kind: 'unresolved', reason: 'cf-resource-limit' }
      }
      sum += compensation
      // Live source check before publication: withhold on mid-read changes.
      const live = source.stamp(stamp.sheetId)
      if (!stampsMatch(live, stamp)) {
        return { kind: 'unresolved', reason: 'cf-stale-generation', pendingSheetIds: [stamp.sheetId] }
      }
      // Cooperative ascending sort for top/bottom cutoffs.
      if (!(await cooperativeSort(values, scratch, numericCount, {
        pending: 0, signal, budget,
      }))) {
        return { kind: 'unresolved', reason: 'cf-resource-limit' }
      }
      if (signal.aborted) throw cancelled()
      const mean = Number.isFinite(sum) && numericCount > 0 ? sum / numericCount : undefined
      // Two-pass population variance around the compensated mean.
      let devSq = 0
      let devComp = 0
      if (mean !== undefined) {
        for (let i = 0; i < numericCount; i++) {
          const dev = values[i]! - mean
          if (!(await spendUnits(spend, 1))) {
            return { kind: 'unresolved', reason: 'cf-resource-limit' }
          }
          const sq = dev * dev
          if (!Number.isFinite(sq)) { devSq = Infinity; break }
          const t = devSq + sq
          devComp += devSq - t + sq
          devSq = t
        }
        if (!(await flushSpend(spend))) {
          return { kind: 'unresolved', reason: 'cf-resource-limit' }
        }
      }
      const variance = mean === undefined || !Number.isFinite(devSq + devComp) ? undefined : Math.max(0, (devSq + devComp) / numericCount)
      // Admit the retained copy before building it.
      if (!growHold(numericCount * 8 + 64)) {
        return { kind: 'unresolved', reason: 'cf-resource-limit' }
      }
      const retained = new Float64Array(numericCount)
      if (!(await copyInto(retained, 0, values.subarray(0, numericCount), numericCount, spend))) {
        return { kind: 'unresolved', reason: 'cf-resource-limit' }
      }
      if (signal.aborted) throw cancelled()
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
        declaredCells,
        spanCells,
        storedCells: stored,
      }
      // Atomically transfer the workspace peak into the adopted candidate:
      // no refund-then-rereserve window where the result sits uncharged.
      // The hold stops counting as workspace here; the check below therefore
      // compares only retained bytes plus the new adoption against the cap.
      workspaces.delete(hold)
      if (!checkCap(bytes)) {
        try { hold.release() } catch { /* release-only path */ }
        return { kind: 'unresolved', reason: 'cf-resource-limit' }
      }
      const adopted = hold.adopt(stats, bytes)
      retainedHeld += bytes
      adoptedLease = adopted
      return { kind: 'ready', value: { stats, bytes, candidate: adopted } }
    } catch (error) {
      if (isCancel(error) || signal.aborted) throw cancelled()
      if (isBudgetRefusal(error)) return { kind: 'unresolved', reason: 'cf-resource-limit' }
      throw error
    } finally {
      if (adoptedLease === undefined) {
        workspaces.delete(hold)
        try { hold.release() } catch { /* release-only path */ }
      }
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
    const out: Array<CfOutcome<CfLease<CfRangeStatistics>>> = []
    const checkLive = (requested: CfStamp): boolean => {
      try {
        return stampsMatch(source.stamp(requested.sheetId), requested)
      } catch {
        return false
      }
    }
    // Populated bounds memo keyed by the full source/owner/structure identity
    // the statistics key uses: a structure change, owner replacement, or new
    // value generation must never reuse older geometry. Pending (undefined)
    // answers are never stored, so a later complete source at the same
    // generation is re-queried instead of staying pending indefinitely.
    // The underlying scan is cooperative (budgeted, abortable, host-yielding).
    const getBounds = async (
      sheetId: string,
      identity: CfStamp,
      budget: CfWorkBudget,
      signal: AbortSignal,
    ): Promise<SourceRect | undefined> => {
      const memoKey = `${sheetId}|${identity.owner}|${identity.structureGen}|${identity.workbookValueGen}`
      if (!boundsMemo.has(memoKey)) {
        const fresh = await source.bounds(sheetId, budget, signal)
        if (fresh !== undefined) boundsMemo.set(memoKey, fresh)
        return fresh
      }
      return boundsMemo.get(memoKey)
    }
    const staleOutcome = (): CfOutcome<CfLease<CfRangeStatistics>> => ({
      kind: 'unresolved', reason: 'cf-stale-generation', pendingSheetIds: [stamp.sheetId],
    })
    const releaseUnpublished = (candidate: CfLease<CfRangeStatistics>, bytes: number): void => {
      try { candidate.release() } catch { /* release-only path */ }
      trackRelease(bytes)
    }
    const settleWaiter = (w: Waiter, outcome: CfOutcome<CfLease<CfRangeStatistics>>): void => {
      if (w.done) return
      w.done = true
      w.signal.removeEventListener('abort', w.onAbort)
      w.resolve(outcome)
    }
    // Central job settlement: the job owns the adopted candidate until a
    // live waiter takes it. Exactly one of: shared entry created and handed
    // out, candidate released unclaimed, or error propagated to waiters.
    const settleJob = async (key: string, job: InflightJob, built: BuiltOutcome): Promise<void> => {
      if (built.kind === 'ready') {
        // Final liveness gate: re-read the source generation after a host
        // yield, so a change queued by the pre-sort stamp read lands first.
        // A stale candidate is released, never published, for every waiter.
        let first: CfStamp | undefined
        let second: CfStamp | undefined
        try { first = source.stamp(stamp.sheetId) } catch { first = undefined }
        await hostTask()
        try { second = source.stamp(stamp.sheetId) } catch { second = undefined }
        if (inflight.get(key) === job) inflight.delete(key)
        const live = [...job.waiters]
        job.waiters.clear()
        if (first === undefined || second === undefined || !stampsMatch(first, stamp) || !stampsMatch(second, stamp)) {
          releaseUnpublished(built.value.candidate, built.value.bytes)
          for (const w of live) settleWaiter(w, staleOutcome())
          return
        }
        const active = live.filter(w => !w.done)
        if (active.length === 0) {
          // No live waiter can take the adopted result: release it instead
          // of caching an unowned charge.
          releaseUnpublished(built.value.candidate, built.value.bytes)
          return
        }
        let entry = cache.get(key)
        if (entry === undefined) {
          entry = { stats: built.value.stats, bytes: built.value.bytes, lease: built.value.candidate, refs: 0 }
          cache.set(key, entry)
          if (cache.size > maxEntries) evict()
        } else {
          // Same key completed twice (inflight dedupe should prevent this):
          // prefer the cached entry and release the duplicate candidate.
          releaseUnpublished(built.value.candidate, built.value.bytes)
        }
        for (const w of active) {
          if (w.done) continue
          w.done = true
          w.signal.removeEventListener('abort', w.onAbort)
          w.resolve(makeOutcome(key, entry))
        }
        return
      }
      if (inflight.get(key) === job) inflight.delete(key)
      const live = [...job.waiters]
      job.waiters.clear()
      for (const w of live) settleWaiter(w, built)
    }
    const failJob = (key: string, job: InflightJob, error: unknown): void => {
      if (inflight.get(key) === job) inflight.delete(key)
      for (const w of [...job.waiters]) {
        if (w.done) continue
        w.done = true
        w.signal.removeEventListener('abort', w.onAbort)
        w.reject(isCancel(error) || w.signal.aborted ? cancelled() : error)
      }
      job.waiters.clear()
    }
    for (const scope of ruleScopes) {
      if (signal.aborted) {
        for (const o of out) if (o.kind === 'ready') o.value.release()
        throw cancelled()
      }
      const key = `${stamp.sheetId}|${stamp.owner}|${stamp.structureGen}|${stamp.workbookValueGen}|${scopeKey(scope)}|${hiddenKey(hidden, hiddenRows)}`
      const hit = cache.get(key)
      if (hit !== undefined) {
        // Warm hit: revalidate the live source before publishing, then share
        // the adopted charge. No loops, no new scan.
        if (!checkLive(stamp)) {
          out.push({ kind: 'unresolved', reason: 'cf-stale-generation', pendingSheetIds: [stamp.sheetId] })
          continue
        }
        cache.delete(key)
        cache.set(key, hit)
        out.push(makeOutcome(key, hit))
        continue
      }
      if (!checkLive(stamp)) {
        out.push({ kind: 'unresolved', reason: 'cf-stale-generation', pendingSheetIds: [stamp.sheetId] })
        continue
      }
      let job = inflight.get(key)
      if (job === undefined) {
        const controller = new AbortController()
        const waiters = new Set<Waiter>()
        const jobPromise: Promise<BuiltOutcome> = (async (): Promise<BuiltOutcome> => {
          // Refuse before touching the source when the account cannot hold
          // even the minimum unit: a cold unbounded scan must not run on an
          // exhausted ledger. Bounds stay lazy and memoized per full source
          // identity so warm hits never rescan the sheet.
          if (!checkCap(256)) {
            return { kind: 'unresolved', reason: 'cf-resource-limit' }
          }
          const probe = memory.reserve(256)
          if (probe.kind === 'unresolved') {
            return { kind: 'unresolved', reason: 'cf-resource-limit' }
          }
          probe.release()
          const uspend: SpendState = { pending: 0, signal: controller.signal, budget }
          // Admitted sweep workspace, tracked in the session cap for its
          // lifetime: edges, coordinate list, interval arrays and sort
          // scratch for one sweep. The clamped set itself lives in the
          // separate clamp hold below and coexists through both sweeps.
          const countedUnion = async (rects: readonly SourceRect[]): Promise<number | undefined> => {
            // Simultaneous bound from the payload model: a 2n-number edge
            // set, a 2n-number coordinate list, up to n interval pairs, and
            // sort scratch at the interval-sort boundary, plus headroom.
            const need = 272 + rects.length * 208
            if (!checkCap(need)) return undefined
            const hold = memory.reserve(need)
            if (hold.kind === 'unresolved') return undefined
            workspaces.set(hold, need)
            try {
              return await sweepUnion(rects, uspend)
            } finally {
              workspaces.delete(hold)
              try { hold.release() } catch { /* release-only path */ }
            }
          }
          // Clamped rectangles are constructed objects the sweeps and the
          // scan below read throughout the job: admit them (objects plus
          // array, with growth headroom) before construction and hold that
          // charge until the job settles.
          const clampNeed = 16 + scope.length * 136
          if (!checkCap(clampNeed)) {
            return { kind: 'unresolved', reason: 'cf-resource-limit' }
          }
          const clampHold = memory.reserve(clampNeed)
          if (clampHold.kind === 'unresolved') {
            return { kind: 'unresolved', reason: 'cf-resource-limit' }
          }
          workspaces.set(clampHold, clampNeed)
          try {
            let bounds: SourceRect | undefined
            const clamped: SourceRect[] = []
            let clampedSum = 0
            let declared: number | undefined
            let span: number | undefined
            try {
              bounds = await getBounds(stamp.sheetId, stamp, budget, controller.signal)
              if (bounds === undefined) {
                return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: [stamp.sheetId] }
              }
              for (const r of scope) {
                if (!(await spendUnits(uspend, 1))) {
                  return { kind: 'unresolved', reason: 'cf-resource-limit' }
                }
                const c = clampRect(r, bounds)
                if (c !== undefined) { clamped.push(c); clampedSum += rectArea(c) }
              }
              if (!(await flushSpend(uspend))) {
                return { kind: 'unresolved', reason: 'cf-resource-limit' }
              }
              declared = await countedUnion(scope)
              if (declared === undefined) {
                return { kind: 'unresolved', reason: 'cf-resource-limit' }
              }
              span = await countedUnion(clamped)
              if (span === undefined) {
                return { kind: 'unresolved', reason: 'cf-resource-limit' }
              }
              if (!(await flushSpend(uspend))) {
                return { kind: 'unresolved', reason: 'cf-resource-limit' }
              }
            } catch (error) {
              if (isCancel(error) || controller.signal.aborted) throw cancelled()
              if (isBudgetRefusal(error)) return { kind: 'unresolved', reason: 'cf-resource-limit' }
              throw error
            }
            if (clamped.length === 0 || span === 0) {
              const empty: CfRangeStatistics = {
                scopeKey: key, numericCount: 0, sum: 0, mean: undefined, min: undefined, max: undefined,
                stddevPop: undefined, sortedAsc: new Float64Array(0), frequencies: new Map(),
                declaredCells: declared, spanCells: span, storedCells: 0,
              }
              if (!checkCap(256)) {
                return { kind: 'unresolved', reason: 'cf-resource-limit' }
              }
              const hold = memory.reserve(256)
              if (hold.kind === 'unresolved') {
                return { kind: 'unresolved', reason: 'cf-resource-limit' }
              }
              const candidate = hold.adopt(empty, 256)
              retainedHeld += 256
              return { kind: 'ready', value: { stats: empty, bytes: 256, candidate } }
            }
            // Overlap exists exactly when the union is smaller than the sum.
            const needsDedup = span < clampedSum
            try {
              return await build(clamped, span, declared, needsDedup, stamp, source, budget, controller.signal, hidden, hiddenRows, key)
            } catch (error) {
              if (isCancel(error) || controller.signal.aborted) throw cancelled()
              if (isBudgetRefusal(error)) return { kind: 'unresolved', reason: 'cf-resource-limit' }
              throw error
            }
          } finally {
            // The clamped workspace lives through sweeps and the scan that
            // reads it; release only when the job settles.
            workspaces.delete(clampHold)
            try { clampHold.release() } catch { /* release-only path */ }
          }
        })()
        const abort = (): void => {
          try { controller.abort() } catch { /* already aborted */ }
        }
        job = { waiters, controller, build: jobPromise, abort }
        inflight.set(key, job)
        void job.build.then(
          async (built) => { await settleJob(key, job!, built) },
          (error: unknown) => { failJob(key, job!, error) },
        )
      }
      // Join the shared build; an abort here detaches only this waiter.
      // The job owns the adopted candidate; settlement hands a shared entry
      // to live waiters or releases the candidate when none remain.
      const outcome: CfOutcome<CfLease<CfRangeStatistics>> = await new Promise<CfOutcome<CfLease<CfRangeStatistics>>>((resolve, reject) => {
        const waiter: Waiter = {
          done: false, signal,
          onAbort: () => {},
          resolve: (o) => resolve(o),
          reject: (e) => reject(e),
        }
        const onAbort = (): void => {
          if (waiter.done) return
          waiter.done = true
          job!.waiters.delete(waiter)
          signal.removeEventListener('abort', onAbort)
          if (job!.waiters.size === 0) job!.abort()
          reject(cancelled())
        }
        waiter.onAbort = onAbort
        if (signal.aborted) {
          onAbort()
          return
        }
        job!.waiters.add(waiter)
        signal.addEventListener('abort', onAbort, { once: true })
      }).catch((error: unknown) => {
        // Unwind this call's unreturned leases, then propagate.
        for (const o of out) if (o.kind === 'ready') o.value.release()
        throw error
      })
      out.push(outcome)
    }
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
  const toLiteral = (cell: XlsxCell): CfLiteral => {
    if (cell.value === null) return { kind: 'blank' }
    if (cell.valueIsError) return { kind: 'error', value: String(cell.value) }
    if (typeof cell.value === 'number') return { kind: 'number', value: cell.value }
    if (typeof cell.value === 'boolean') return { kind: 'boolean', value: cell.value }
    return { kind: 'string', value: cell.value }
  }
  return {
    bounds: async (_id: string, budget?: CfWorkBudget, signal?: AbortSignal): Promise<SourceRect | undefined> => {
      if (cells.length === 0) return undefined
      // Cooperative bounds scan: every examined cell spends shared work so
      // cold bounds stay budgeted, and cancellation lands at the next host
      // checkpoint — the loop genuinely awaits the host on yield slices.
      // Served from the memoized session result on repeats, never rescanned.
      let minRow = Infinity, maxRow = -Infinity, minCol = Infinity, maxCol = -Infinity
      for (const cell of cells) {
        if (signal?.aborted) throw cancelled()
        if (budget !== undefined) {
          const spent = budget.spend(1)
          if (spent === 'refused') {
            if (signal?.aborted) throw cancelled()
            throw new PackageError('resource-limit', 'CF resident bounds exceeded the shared work budget', undefined, 'memory-budget')
          }
          if (spent === 'yield') {
            await hostTask()
            if (signal?.aborted) throw cancelled()
            budget.resume()
          }
        }
        if (cell.row < minRow) minRow = cell.row
        if (cell.row > maxRow) maxRow = cell.row
        if (cell.col < minCol) minCol = cell.col
        if (cell.col > maxCol) maxCol = cell.col
      }
      return { firstRow: minRow, lastRow: maxRow, firstCol: minCol, lastCol: maxCol }
    },
    stamp: (_id: string) => stamp,
    pages: async function* (id: string, rects: readonly SourceRect[], budget: CfWorkBudget, signal: AbortSignal) {
      if (id !== sheetId) throw new Error('CF resident source queried for the wrong sheet')
      for (const rect of rects) {
        let chunk: XlsxCell[] = []
        const flush = async function* (): AsyncGenerator<CfLease<CfValuePage>> {
          if (chunk.length === 0) return
          const spent = budget.spend(chunk.length)
          if (spent === 'refused') {
            if (signal.aborted) throw cancelled()
            throw new PackageError('resource-limit', 'CF resident page exceeded the shared work budget')
          }
          // Admit the full input-derived page bound before building it:
          // the wrapper and array headers, typed-array backing buffers,
          // values slots, and every per-cell CfLiteral object with its
          // kind and value payloads, matching measurePayloadBytes of the
          // adopted page from below. The 640 base covers the ref staging
          // array and array headers.
          let bytes = 640 + 24 + 3 * 16 + 16 + chunk.length * (8 + 4 + 2)
          for (const cell of chunk) {
            const v = cell.value
            if (v === null) bytes += 24 + 16 + (16 + 5 * 2)
            else if (typeof v === 'number') bytes += 24 + 2 * 16 + (16 + 6 * 2) + 8
            else if (typeof v === 'boolean') bytes += 24 + 2 * 16 + (16 + 7 * 2) + 8
            else if (cell.valueIsError) bytes += 24 + 2 * 16 + (16 + 5 * 2) + (16 + String(v).length * 2)
            else bytes += 24 + 2 * 16 + (16 + 6 * 2) + (16 + v.length * 2)
          }
          const hold = memory.reserve(bytes)
          if (hold.kind === 'unresolved') {
            throw new PackageError('resource-limit', 'CF resident page admission refused', undefined, 'memory-budget')
          }
          try {
            const rows = new Uint32Array(chunk.length)
            const cols = new Uint16Array(chunk.length)
            const values = new Array<CfLiteral>(chunk.length)
            for (let i = 0; i < chunk.length; i++) {
              const cell = chunk[i]!
              rows[i] = cell.row
              cols[i] = cell.col
              values[i] = toLiteral(cell)
            }
            yield hold.adopt<CfValuePage>({ rows, cols, values }, bytes)
          } catch (error) {
            try { hold.release() } catch { /* release-only path */ }
            throw error
          }
          if (spent === 'yield') {
            await hostTask()
            if (signal.aborted) throw cancelled()
            budget.resume()
          }
          chunk = []
        }
        for (const cell of cells) {
          if (signal.aborted) throw cancelled()
          // Spend for every examined cell, including nonmatching ones that
          // are filtered out below: examination itself is shared work.
          const spent = budget.spend(1)
          if (spent === 'refused') {
            if (signal.aborted) throw cancelled()
            throw new PackageError('resource-limit', 'CF resident page exceeded the shared work budget')
          }
          if (spent === 'yield') {
            await hostTask()
            if (signal.aborted) throw cancelled()
            budget.resume()
          }
          if (cell.row < rect.firstRow || cell.row > rect.lastRow || cell.col < rect.firstCol || cell.col > rect.lastCol) {
            continue
          }
          chunk.push(cell)
          if (chunk.length >= 64) yield* flush()
        }
        yield* flush()
      }
    },
  }
}
