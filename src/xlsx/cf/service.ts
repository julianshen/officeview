import { PackageError } from '../../core/package-error'
import { hostTask } from '../view/scheduler'
import { createCfCache, sameCfStamp } from './cache'
import { composeCellAsync } from './compose'
import { resolveCfCellStyle } from './dxf'
import { CF_LIMITS, createCfWorkBudget, type CfMemory, type CfWorkBudget } from './limits'
import type { CfCellState, CfLease, CfLiteral, CfOutcome, CfPrepared, CfPrepareRequest, CfRule, CfSheetRules, CfValueSource, SourceRect } from './model'
import { cfRuntimeEvidence, evaluatePredicate, type CfEvidenceGate } from './predicates'
import { createCfStatistics, type CfRangeStatistics, type CfStatisticsOptions, type CfStatisticsSession } from './statistics'
const quota = () => new PackageError('resource-limit', 'CF preparation exceeds its admitted batch')
const cancel = () => new PackageError('cancelled', 'CF preparation cancelled or superseded')
const blank: CfLiteral = Object.freeze({ kind: 'blank' })
const inside = (r: SourceRect, row: number, col: number) => row >= r.firstRow && row <= r.lastRow && col >= r.firstCol && col <= r.lastCol
function statistical(rule: CfRule) {
  const d = rule.definition
  return d.kind === 'top10' || d.kind === 'aboveAverage' || (d.kind === 'predicate' && (d.type === 'uniqueValues' || d.type === 'duplicateValues'))
}

/** A caller-owned, single-generation statistics lease set. The caller admits
 * the rule-index framing, releases ready results, then disposes the session. */
export interface CfPaintStatistics {
  session: CfStatisticsSession
  stamp: import('./model').CfStamp
  results: Map<CfRule, CfOutcome<CfLease<CfRangeStatistics>>>
}

/** Admitted preparation bound for rules over rects (Task 9): the same arithmetic
 * `prepareCf` enforces, without side effects. Returns undefined when any cap is
 * exceeded or the bound is unrepresentable, so callers can reserve upfront in
 * their own aggregate account before constructing anything. */
export function estimateCfPrepareBytes(rules: CfSheetRules, sourceRects: readonly SourceRect[]): number | undefined {
  const valid = (n: number): boolean => Number.isSafeInteger(n) && n >= 0
  if (sourceRects.length > CF_LIMITS.rangesPerRule || rules.rules.length > CF_LIMITS.rules) return undefined
  let cells = 0
  for (const rect of sourceRects) {
    if (![rect.firstRow, rect.lastRow, rect.firstCol, rect.lastCol].every(valid)
      || rect.firstRow > rect.lastRow || rect.firstCol > rect.lastCol
      || rect.lastRow > 1048575 || rect.lastCol > 16383) return undefined
    cells += (rect.lastRow - rect.firstRow + 1) * (rect.lastCol - rect.firstCol + 1)
    if (cells > CF_LIMITS.preparedCells) return undefined
  }
  let propertyBytes = 0, ranges = 0
  for (const rule of rules.rules) {
    ranges += rule.ranges.length
    if (rule.ranges.length > CF_LIMITS.rangesPerRule || ranges > CF_LIMITS.rangesPerSheet) return undefined
    if (rule.dxf) for (const key of Object.keys(rule.dxf)) {
      const v = (rule.dxf as Record<string, unknown>)[key]
      propertyBytes += 192 + key.length * 4 + (typeof v === 'string' ? v.length * 2 : 16)
    }
    for (const key of rule.writes) propertyBytes += 128 + String(key).length * 4
  }
  const bound = 2048 + sourceRects.length * 160 + rules.rules.length * 640 + cells * (512 + propertyBytes)
  return Number.isSafeInteger(bound) && bound <= CF_LIMITS.scratchBytes ? bound : undefined
}
/** Lazy per-sheet engine. Parser-owned rules must remain leased by the caller for
 * this engine's lifetime; published snapshots retain only copied winning deltas. */
export function createCfService(sheet: CfSheetRules, memory: CfMemory,
  options: { evidence?: CfEvidenceGate; budget?: () => CfWorkBudget; statistics?: CfPaintStatistics; statisticsOptions?: CfStatisticsOptions; dateContext?: import('./predicates').CfDateContext } = {}) {
  const cache = createCfCache(memory)
  const evidence = options.evidence ?? cfRuntimeEvidence
  async function prepareCf(request: CfPrepareRequest, source: CfValueSource, account = memory): Promise<CfLease<CfPrepared>> {
    if (account !== memory) throw quota()
    if (options.statistics && !sameCfStamp(options.statistics.stamp, request.stamp)) throw cancel()
    cache.activate(request.stamp) // invalidate old paint even if the new request refuses
    if (request.signal.aborted) throw cancel()
    if (!Number.isSafeInteger(request.grantBytes) || request.grantBytes < 256 || request.sourceRects.length > CF_LIMITS.rangesPerRule) throw quota()
    // Single source of truth with estimateCfPrepareBytes: the same caps and the
    // same bound, enforced here against the caller's grant as well.
    const bound = estimateCfPrepareBytes(sheet, request.sourceRects)
    if (bound === undefined || bound > request.grantBytes) throw quota()
    // Admit temporary encoding before map/join; cache admits retained key ownership
    // synchronously before this temporary hold ends.
    const keyHold = memory.reserve(256 + request.sourceRects.length * 160)
    if (keyHold.kind !== 'reserved') throw quota()
    try {
      const key = request.sourceRects.map(r => `${r.firstRow},${r.lastRow},${r.firstCol},${r.lastCol}`).join(';') + `|${request.grantBytes}`
      return cache.prepare(key, request.stamp, request.signal, async signal => {
        const hold = memory.reserve(bound)
        if (hold.kind !== 'reserved') throw quota()
        const budget = options.budget?.() ?? createCfWorkBudget({ signal })
        const statistics = options.statistics?.session ?? createCfStatistics(memory)
        const stats = options.statistics?.results ?? new Map<CfRule, CfOutcome<CfLease<CfRangeStatistics>>>()
        const leases: CfLease<CfRangeStatistics>[] = []
        let adopted = false
        const work = async (units = 1) => {
          if (signal.aborted) throw cancel()
          for (let remaining = units; remaining > 0;) {
            const amount = Math.min(remaining, 64), result = budget.spend(amount)
            if (result === 'refused') throw quota()
            remaining -= amount
            if (result === 'yield') { await hostTask(); if (signal.aborted) throw cancel(); budget.resume() }
          }
        }
        try {
          const stamp = Object.freeze({ ...request.stamp })
          const rects = request.sourceRects.map(r => Object.freeze({ ...r }))
          // A charged sorted rule index, without mutating the parser's ordering.
          const ordered = [...sheet.rules]
          for (let i = 1; i < ordered.length; i++) {
            const value = ordered[i]!; let j = i - 1
            while (j >= 0) {
              await work()
              const other = ordered[j]!
              if (other.priority < value.priority || (other.priority === value.priority && other.documentOrder >= value.documentOrder)) break
              ordered[j + 1] = other; j--
            }
            ordered[j + 1] = value
          }
          const data = new Map<number, CfOutcome<CfCellState> | undefined>()
          for (const rect of rects) for (let row = rect.firstRow; row <= rect.lastRow; row++) for (let col = rect.firstCol; col <= rect.lastCol; col++) {
            await work(); data.set(row * 16384 + col, undefined)
          }
          const snapshot = (reason?: string): CfLease<CfPrepared> => {
            if (signal.aborted || !sameCfStamp(source.stamp(stamp.sheetId), stamp)) throw cancel()
            const prepared: CfPrepared = Object.freeze({ stamp,
              at(row: number, col: number): CfOutcome<CfCellState> {
                if (!Number.isInteger(row) || !Number.isInteger(col) || col < 0 || col >= 16384 || row < 0 || row >= 1048576) return { kind: 'unresolved', reason: 'cf-outside-preparation' }
                if (!sameCfStamp(source.stamp(stamp.sheetId), stamp)) return { kind: 'unresolved', reason: 'cf-stale-generation' }
                const id = row * 16384 + col
                if (!data.has(id)) return { kind: 'unresolved', reason: 'cf-outside-preparation' }
                return reason ? { kind: 'unresolved', reason } : data.get(id)!
              }, dispose() {},
            })
            const result = hold.adopt(prepared) // conservative bound retained; no refund of live closures/maps
            adopted = true
            return result
          }
          if (sheet.unresolvedSheet) return snapshot(sheet.unresolvedSheet.reason)
          if (await source.bounds(stamp.sheetId, budget, signal) === undefined) return snapshot('cf-pending-data')
          for (const rule of ordered) if (statistical(rule) && !stats.has(rule)) {
            const [out] = await statistics.prepareStatistics([rule.ranges], stamp, source, budget, { signal, ...options.statisticsOptions })
            stats.set(rule, out!)
            if (!options.statistics && out?.kind === 'ready') leases.push(out.value)
          }
          const evaluate = async (row: number, col: number, value: CfLiteral): Promise<CfOutcome<CfCellState>> => {
            const textBytes = value.kind === 'string' ? 256 + value.value.length * 8 : 0
            if (textBytes > CF_LIMITS.scratchBytes) return { kind: 'unresolved', reason: 'cf-resource-limit' }
            const textHold = memory.reserve(textBytes)
            if (textHold.kind !== 'reserved') return { kind: 'unresolved', reason: 'cf-resource-limit' }
            try {
              const matching: CfRule[] = [], outcomes: CfOutcome<boolean>[] = []
              for (const rule of ordered) {
                await work()
                let covered = false
                for (const rect of rule.ranges) { await work(); if (inside(rect, row, col)) { covered = true; break } }
                if (!covered) continue
                matching.push(rule)
                if (value.kind === 'string') await work(value.value.length + 1)
                const preparedStats = stats.get(rule)
                outcomes.push(preparedStats?.kind === 'unresolved' ? preparedStats
                  : evaluatePredicate(rule, { row, col }, value, preparedStats?.kind === 'ready' ? preparedStats.value.value : undefined, evidence, options.dateContext))
              }
              // The pure compositor implements the candidate ordering contract; enabling
              // multi-positive precedence in production still requires its oracle.
              let positives = 0
              const priorities = new Set<number>()
              for (let i = 0; i < outcomes.length; i++) {
                await work()
                const outcome = outcomes[i]!
                if (outcome.kind !== 'ready' || !outcome.value) continue
                const priority = matching[i]!.priority
                if (priorities.has(priority) && !evidence.allows('priority:equal-priority-first')) return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
                priorities.add(priority)
                if (++positives > 1 && (!evidence.allows('priority:fill-winner')
                  || !evidence.allows('priority:orthogonal-font-bold'))) return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
              }
              return Object.freeze({ kind: 'ready', value: await composeCellAsync(matching, outcomes, work) })
            } finally { textHold.release() }
          }
          for await (const page of source.pages(stamp.sheetId, rects, budget, signal)) {
            try {
              const { rows, cols, values } = page.value
              if (rows.length !== cols.length || rows.length !== values.length) throw quota()
              for (let i = 0; i < rows.length; i++) {
                await work(); const row = rows[i]!, col = cols[i]!, id = row * 16384 + col
                if (data.has(id) && data.get(id) === undefined) data.set(id, await evaluate(row, col, values[i]!))
              }
            } finally { page.release() }
          }
          for (const [id, state] of data) if (state === undefined) { await work(); data.set(id, await evaluate(Math.floor(id / 16384), id % 16384, blank)) }
          return snapshot()
        } finally {
          for (const lease of leases) lease.release()
          if (!options.statistics) statistics.dispose()
          if (!adopted) hold.release()
        }
      })
    } finally { keyHold.release() }
  }
  return { prepareCf, at: cache.at, dispose: cache.dispose, resolveCfCellStyle }
}
