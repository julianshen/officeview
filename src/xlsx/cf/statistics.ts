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

/** Exact union area via coordinate compression (rect counts here are small). */
function unionArea(rects: readonly SourceRect[]): number {
  if (rects.length === 0) return 0
  if (rects.length === 1) return rectArea(rects[0]!)
  if (rects.length > 64) {
    let total = 0
    for (const r of rects) total += rectArea(r)
    return total
  }
  const xs = new Set<number>()
  const ys = new Set<number>()
  for (const r of rects) {
    xs.add(r.firstCol)
    xs.add(r.lastCol + 1)
    ys.add(r.firstRow)
    ys.add(r.lastRow + 1)
  }
  const xList = [...xs].sort((a, b) => a - b)
  const yList = [...ys].sort((a, b) => a - b)
  let total = 0
  for (let xi = 0; xi < xList.length - 1; xi++) {
    for (let yi = 0; yi < yList.length - 1; yi++) {
      const x0 = xList[xi]!
      const y0 = yList[yi]!
      let covered = false
      for (const r of rects) {
        if (r.firstCol <= x0 && x0 <= r.lastCol && r.firstRow <= y0 && y0 <= r.lastRow) {
          covered = true
          break
        }
      }
      if (covered) total += (xList[xi + 1]! - x0) * (yList[yi + 1]! - y0)
    }
  }
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
  return spendUnits(state, 0)
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
  let retainedHeld = 0

  const withinCap = (additional: number, exclude?: object): number => {
    let total = retainedHeld + additional
    for (const [hold, bytes] of workspaces) if (hold !== exclude) total += bytes
    return total
  }
  const checkCap = (additional: number, exclude?: object): boolean => withinCap(additional, exclude) <= sessionCap

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
      if (!checkCap(delta, hold)) return false
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
        if (!growHold(delta)) return false
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
              if (needsDedup && value.kind !== 'blank') {
                const id = row * 16384 + col
                if (seen!.has(id)) continue
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
    const bounds = source.bounds(stamp.sheetId)
    const out: Array<CfOutcome<CfLease<CfRangeStatistics>>> = []
    const checkLive = (requested: CfStamp): boolean => {
      try {
        return stampsMatch(source.stamp(requested.sheetId), requested)
      } catch {
        return false
      }
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
          if (bounds === undefined) {
            return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: [stamp.sheetId] }
          }
          const clamped: SourceRect[] = []
          for (const r of scope) {
            const c = clampRect(r, bounds)
            if (c !== undefined) clamped.push(c)
          }
          const declared = unionArea(scope)
          const span = unionArea(clamped)
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
          let needsDedup = false
          for (let a = 0; a < clamped.length && !needsDedup; a++) {
            for (let b = a + 1; b < clamped.length; b++) {
              if (rectsOverlap(clamped[a]!, clamped[b]!)) { needsDedup = true; break }
            }
          }
          try {
            return await build(clamped, span, declared, needsDedup, stamp, source, budget, controller.signal, hidden, hiddenRows, key)
          } catch (error) {
            if (isCancel(error) || controller.signal.aborted) throw cancelled()
            if (isBudgetRefusal(error)) return { kind: 'unresolved', reason: 'cf-resource-limit' }
            throw error
          }
        })()
        const abort = (): void => {
          try { controller.abort() } catch { /* already aborted */ }
        }
        job = { waiters, controller, build: jobPromise, abort }
        inflight.set(key, job)
        void jobPromise.then(
          () => { if (inflight.get(key) === job) inflight.delete(key) },
          () => { if (inflight.get(key) === job) inflight.delete(key) },
        )
      }
      // Join the shared build; an abort here detaches only this waiter.
      const outcome: CfOutcome<CfLease<CfRangeStatistics>> = await new Promise<CfOutcome<CfLease<CfRangeStatistics>>>((resolve, reject) => {
        const waiter: Waiter = { done: false }
        const onAbort = (): void => {
          if (waiter.done) return
          waiter.done = true
          job!.waiters.delete(waiter)
          signal.removeEventListener('abort', onAbort)
          if (job!.waiters.size === 0) job!.abort()
          reject(cancelled())
        }
        if (signal.aborted) {
          onAbort()
          return
        }
        waiter.done = false
        job!.waiters.add(waiter)
        signal.addEventListener('abort', onAbort, { once: true })
        job!.build.then(
          (built) => {
            signal.removeEventListener('abort', onAbort)
            if (waiter.done) return
            waiter.done = true
            job!.waiters.delete(waiter)
            if (built.kind !== 'ready') {
              resolve(built)
              return
            }
            let entry = cache.get(key)
            if (entry === undefined) {
              entry = { stats: built.value.stats, bytes: built.value.bytes, lease: built.value.candidate, refs: 0 }
              cache.set(key, entry)
              if (cache.size > maxEntries) evict()
            }
            resolve(makeOutcome(key, entry))
          },
          (error: unknown) => {
            signal.removeEventListener('abort', onAbort)
            if (waiter.done) return
            waiter.done = true
            job!.waiters.delete(waiter)
            if (job!.waiters.size === 0) job!.abort()
            reject(isCancel(error) || signal.aborted ? cancelled() : error)
          },
        )
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
    bounds: (_id: string) => {
      if (cells.length === 0) return undefined
      let minRow = Infinity, maxRow = -Infinity, minCol = Infinity, maxCol = -Infinity
      for (const cell of cells) {
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
          // Admit staging workspace plus the typed page before building either.
          // The 640 base covers the ref staging array and array headers.
          let bytes = 640 + chunk.length * 8
          for (const cell of chunk) {
            const v = cell.value
            if (typeof v === 'string') bytes += 16 + v.length * 2
            else if (cell.valueIsError) bytes += 16 + String(v).length * 2
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
          if (cell.row < rect.firstRow || cell.row > rect.lastRow || cell.col < rect.firstCol || cell.col > rect.lastCol) {
            continue
          }
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
          chunk.push(cell)
          if (chunk.length >= 64) yield* flush()
        }
        yield* flush()
      }
    },
  }
}
