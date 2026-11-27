/**
 * The across-midnight follow-up (brief cf-midnight, owner choice 2026-11-27): a viewer left open across local midnight.
 *
 * A clock-reading rule (`timePeriod`, a formula calling TODAY() or NOW()) is re-evaluated at every preparation, while a
 * cell whose value was computed from the clock at load (a `ca` TODAY() cell) keeps the load day. Once the live local day
 * differs from the load day, the rule and the cell can name two days, a state Excel was never measured in. Excel's own
 * behaviour, with a workbook open and no edit, is one of three pre-registered readings of `clock:across-midnight`:
 * - `frozen`: nothing changes until a recalculation; rules and cells stay on the load day;
 * - `rules-tick`: clock rules move to the new day, clock cells stay on the load day (production before this follow-up);
 * - `full-recalc`: clock cells and rules both move to the new day.
 *
 * Unregistered (the default), production withholds, from the first preparation on another day until the document is
 * reloaded, every cell of a clock-reading rule and every cell whose rule reads a clock-tainted cell (`cf-unmeasured-across-
 * midnight`). A registered reading (with its evidence) replaces the withhold on every path that holds the context:
 * - `rules-tick` removes the gate (the clock rules read the live clock, the cells keep their load values);
 * - `frozen` evaluates every rule of the preparation at the load instant (predicates and formulas alike);
 * - `full-recalc` lets the clock rules read the live clock and keeps the tainted cells withheld: no route recalculates a
 *   cell after load (a resident document is evaluated once at parse, a paged one shows saved values), so the new-day cell
 *   values Excel would show are unknown here (narrowing, never a guess).
 */
import { epochToSerial } from '../formula/serial'
import { clockTaintHas, documentLoadClockOf, pagedWorkbookClockTainted } from '../formula/loadClock'
import type { CfDateContext } from './predicates'

export const CF_ACROSS_MIDNIGHT_CAP = 'clock:across-midnight'
export const CF_ACROSS_MIDNIGHT_READINGS = ['frozen', 'rules-tick', 'full-recalc'] as const
export type CfAcrossMidnightReading = (typeof CF_ACROSS_MIDNIGHT_READINGS)[number]
/** What production applies across midnight: a registered reading, or the default withhold. */
export type CfAcrossMidnight = CfAcrossMidnightReading | 'withhold'
export const CF_ACROSS_MIDNIGHT_REASON = 'cf-unmeasured-across-midnight'
/** The reading production implements, registered from the capture of stage1-clock-across-midnight: none yet. */
export const CF_ACROSS_MIDNIGHT_PRODUCTION_READINGS: Readonly<Partial<Record<string, CfAcrossMidnightReading>>> = Object.freeze({})

let proposedReadings: Readonly<Partial<Record<string, CfAcrossMidnightReading>>> | null = null
/** Test seam: rehearse a proposed registration as if it were in the table above. Null clears. Production never calls it. */
export function injectCfAcrossMidnightReadings(readings: Readonly<Partial<Record<string, CfAcrossMidnightReading>>> | null): void {
  proposedReadings = readings ? Object.freeze({ ...readings }) : null
}
/** The registered reading of the id (undefined: unregistered), a proposed one (the test seam) first. */
export function cfAcrossMidnightReading(): CfAcrossMidnightReading | undefined {
  const table = proposedReadings && Object.prototype.hasOwnProperty.call(proposedReadings, CF_ACROSS_MIDNIGHT_CAP) ? proposedReadings : CF_ACROSS_MIDNIGHT_PRODUCTION_READINGS
  const r = table[CF_ACROSS_MIDNIGHT_CAP]
  return r !== undefined && (CF_ACROSS_MIDNIGHT_READINGS as readonly string[]).includes(r) ? r : undefined
}
/** What production applies: the registered reading with its evidence, else the withhold. */
export function cfAcrossMidnight(checkEvidence: (cap: string) => boolean): CfAcrossMidnight {
  const r = cfAcrossMidnightReading()
  return r !== undefined && checkEvidence(CF_ACROSS_MIDNIGHT_CAP) ? r : 'withhold'
}

/** A rectangle of source cells (0-based, inclusive). */
export type CfMidnightRect = Readonly<{ firstRow: number; lastRow: number; firstCol: number; lastCol: number }>
/**
 * The cells whose values may move with the clock (`formula/loadClock.ts`): none, every cell (unknown, fail closed), or a
 * query by sheet (undefined: the host sheet) and rectangle.
 */
export type CfClockTaint = Readonly<{ kind: 'none' }> | Readonly<{ kind: 'all' }> | Readonly<{ kind: 'cells'; has: (sheet: string | undefined, rect: CfMidnightRect) => boolean }>
/** A clock as a preparation reads it: its local day, and its instant when it has one (a pinned serial has none). */
export type CfMidnightClock = Readonly<{ day: number; epochMs?: number }>
/**
 * One route's across-midnight context: the load clock, given this preparation's live clock and zone (a resident document's
 * is its parse's; a paged workbook's is its first judged preparation's), and the taint.
 */
export interface CfMidnightContext {
  readonly load: (live: CfMidnightClock, timeZone: string, dateSystem: '1900' | '1904') => CfMidnightClock
  readonly taint: CfClockTaint
}
/** The serial offset of the unix epoch in a date system (`cfClockDay` subtracts it). */
export const cfEpochSerialOffset = (dateSystem: '1900' | '1904'): number => dateSystem === '1904' ? 24_107 : 25_569

/**
 * The local day number of an instant in a zone (the unix day of the local calendar date): the numbering of the
 * interactive midnight tick (`localDaySerial`) and of `CfEvaluationContext.clock.clockGen`. A zone the engine cannot
 * resolve falls back to the UTC day.
 */
export function cfClockDay(epochMs: number, timeZone: string, dateSystem: '1900' | '1904'): number {
  const serial = epochToSerial(epochMs, timeZone, dateSystem)
  return typeof serial === 'number' ? Math.floor(serial) - (dateSystem === '1904' ? 24_107 : 25_569) : Number.isFinite(epochMs) ? Math.floor(epochMs / 86_400_000) : 0
}
/** The day a predicate date context names (a pinned serial first, else its instant), or undefined when it names none. */
export function cfDateContextDay(context: CfDateContext | undefined): number | undefined {
  if (!context) return undefined
  const dateSystem = context.dateSystem ?? '1900'
  if (context.clockSerial !== undefined) return Number.isFinite(context.clockSerial) ? Math.floor(context.clockSerial) - (dateSystem === '1904' ? 24_107 : 25_569) : undefined
  const epoch = context.epochNowMs
  return epoch === undefined || !Number.isFinite(epoch) ? undefined : cfClockDay(epoch, context.timeZone ?? 'UTC', dateSystem)
}

/**
 * A resident sheet's context: its document's load clock, recorded at parse (`formula/loadClock.ts`), read at each
 * preparation (a later recalculation can widen the taint). Undefined for a sheet no parse recorded.
 */
export function cfResidentMidnight(sheet: Readonly<{ name: string }>): CfMidnightContext | undefined {
  const record = documentLoadClockOf(sheet)
  if (!record) return undefined
  const taint = record.taint
  return {
    // The load day is the day the load-time evaluation named (its own zone and date system: a `ca` TODAY() cell's day).
    load: () => ({ day: cfClockDay(record.epochMs, record.timeZone, record.dateSystem), epochMs: record.epochMs }),
    taint: taint === 'all' ? { kind: 'all' } : [...taint.values()].some(cells => cells.size > 0)
      ? { kind: 'cells', has: (other, rect) => clockTaintHas(taint, other ?? sheet.name, rect) } : { kind: 'none' },
  }
}

const pagedLoadClocks = new WeakMap<object, CfMidnightClock>()
/**
 * A paged workbook's context. A paged workbook shows saved values and recalculates nothing, so its load instant is its
 * first judged preparation's clock (the day its clock rules were first painted on), kept per workbook; its taint is none only when the
 * completed import counted no clock-sensitive formula (no volatile or `ca` formula, `view/values.ts` isClockSensitive), else every cell (the worker
 * holds no reference graph: fail closed).
 */
export function cfPagedMidnight(workbook: object & Readonly<{ completeness?: string; clockFormulas?: number }>): CfMidnightContext {
  return {
    load: live => {
      let at = pagedLoadClocks.get(workbook)
      if (at === undefined) pagedLoadClocks.set(workbook, at = Object.freeze({ ...live }))
      return at
    },
    taint: pagedWorkbookClockTainted(workbook) ? { kind: 'all' } : { kind: 'none' },
  }
}
