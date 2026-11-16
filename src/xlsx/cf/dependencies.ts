/**
 * The same-sheet dependency snapshot of CF formula evaluation (Stage 5 Task 34; Rulings S5-E, S5-N, S5-SUMIF-SHAPE).
 *
 * One snapshot per preparation holds only the per-cell footprint of its formula rules (hoisted ranges excluded), read
 * once from the host `CfValueSource` and charged per entry (work units before each insertion, bytes before each growth,
 * old plus new at the growth peak). Each per-cell call gets a restricted `SparseCellStore` view built from the snapshot
 * index for that call's own areas: coverage is validated before any traversal, the view's build, every read, every
 * `stored()` element and the sort on entering `stored()` spend from the call's own meter first, and a read outside the
 * view or of an unavailable value taints the call instead of serving a guessed blank.
 *
 * Cross-sheet values are Task 36: a qualified atom refuses here (`cf-formula-dependency`).
 *
 * The batch driver below runs the per-cell calls between cooperative checkpoints: a call never yields inside the
 * engine; the loop awaits `hostTask()` between cells whenever the pacing budget said `'yield'`, and checks cancellation
 * there. It is the seam Task 35 plugs into the service's page loop.
 */
import { PackageError } from '../../core/package-error'
import type { ElementValue } from '../formula/types'
import type { SparseCellStore, StoredCell } from '../formula/refs'
import { hostTask } from '../view/scheduler'
import { measurePayloadBytes } from '../view/payloadBytes'
import {
  cfFormulaFootprint, cfFormulaSemantics, cfLiteralToElement, createCfInvariantMemo, evaluateCfFormula, planCfHoist,
  type CfCallWorkBreakdown, type CfEdgeMode, type CfEvaluationContext, type CfFootprintEntry, type CfFormulaContext, type CfFormulaPacing,
  type CfFormulaProbe, type CfHoist, type CfHoisted, type CfPayloadMetrics, type CfResult,
} from './formula'
import {
  CF_DEPENDENCY_CELL_BYTES, CF_LIMITS, CF_WORK_REFUSED, createCfUnitCeiling, createCfWorkBudget,
  type CfMemory, type CfUnitCeiling, type CfWorkBudget,
} from './limits'
import type { CfFormula, CfLease, CfOutcome, CfRule, CfStamp, CfValueSource, SourceRect } from './model'
import type { CfStatisticsSession } from './statistics'

/** The other sheets' stores (Task 36). Values and generation come from the same store. */
export interface CfWorkbookSource { sheetIdOf(name: string): string | undefined; store(sheetId: string): unknown }
/** One call's restricted view: its store for the engine, its measured payload, and a charged single-cell reader. */
export type CfDependencyView = Readonly<{ store: SparseCellStore; metrics: CfPayloadMetrics; read(row: number, col: number): ElementValue }>
export interface CfDependencySnapshot {
  readonly sheetId: string
  /** Stored entries held. */
  readonly entries: number
  /** A snapshot read (uncharged): taints on a cell outside the snapshot, another sheet, or an unavailable value. */
  read(sheet: string | undefined, row: number, col: number): ElementValue
  readonly tainted: boolean
  reset(): void
  dispose(): void
  /** The per-call store over the call's effective areas. Refused before any traversal when an area is not fully covered. */
  viewFor(effective: readonly CfFootprintEntry[], budget: CfWorkBudget): CfOutcome<CfDependencyView>
}

const UNAVAILABLE: unique symbol = Symbol('cf-unavailable')
type Slot = ElementValue | typeof UNAVAILABLE
const NA: ElementValue = Object.freeze({ kind: 'formula-error', code: '#N/A' })
const cancel = (): PackageError => new PackageError('cancelled', 'CF formula preparation cancelled or superseded')
const unresolved = <T>(reason: string): CfOutcome<T> => ({ kind: 'unresolved', reason })
const inRect = (r: SourceRect, row: number, col: number): boolean => row >= r.firstRow && row <= r.lastRow && col >= r.firstCol && col <= r.lastCol
const clampTo = (r: SourceRect, b: SourceRect): SourceRect | undefined => {
  const out = { firstRow: Math.max(r.firstRow, b.firstRow), lastRow: Math.min(r.lastRow, b.lastRow), firstCol: Math.max(r.firstCol, b.firstCol), lastCol: Math.min(r.lastCol, b.lastCol) }
  return out.firstRow > out.lastRow || out.firstCol > out.lastCol ? undefined : out
}
const containsRect = (o: SourceRect, i: SourceRect): boolean => o.firstRow <= i.firstRow && o.lastRow >= i.lastRow && o.firstCol <= i.firstCol && o.lastCol >= i.lastCol
const area = (r: SourceRect): number => (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1)
const sortUnits = (n: number): number => (n <= 1 ? 0 : n * Math.ceil(Math.log2(n)))
/** One column of the snapshot index: its stored rows ascending, with their values and string units. */
interface Column { rows: number[]; values: Slot[]; units: number[]; sorted: boolean }
/** One new column's bookkeeping, DERIVED from the payload measurer (a map entry holding an empty column record). */
const CF_DEPENDENCY_COLUMN_BYTES: number = measurePayloadBytes(new Map([[0, { rows: [], values: [], units: [], sorted: true }]]))! - measurePayloadBytes(new Map())!
const INITIAL_CAPACITY = 16
function lowerBound(rows: readonly number[], value: number): number {
  let a = 0, b = rows.length
  while (a < b) { const m = (a + b) >>> 1; if (rows[m]! < value) a = m + 1; else b = m }
  return a
}

/**
 * Build the snapshot of `atoms` (each a footprint rect; whole columns already clamped to the populated bounds) from the
 * host source. Work: one unit per stored entry, spent on `budget` before the entry is inserted (the source's own scan
 * pays from `options.scanBudget`, as statistics scans do). Bytes: the declared table, one column record per new column
 * and the entry table at its capacity, reserved on `memory` before each growth, at old plus new capacity at the peak.
 * The lease holds the final table. A refusal publishes nothing and returns the ledger to its baseline.
 */
export async function prepareCfDependencies(atoms: readonly CfFootprintEntry[], host: CfValueSource, workbook: CfWorkbookSource | undefined,
  budget: CfWorkBudget, memory: CfMemory, signal: AbortSignal, options: Readonly<{ sheetId: string; scanBudget?: () => CfWorkBudget }>): Promise<CfOutcome<CfLease<CfDependencySnapshot>>> {
  void workbook // Task 36 (S5-K): cross-sheet stores
  if (signal.aborted) throw cancel()
  if (atoms.some(a => a.sheet !== undefined)) return unresolved('cf-formula-dependency')
  const sheetId = options.sheetId
  const scan = (): CfWorkBudget => options.scanBudget?.() ?? createCfWorkBudget({ signal, scan: true })
  let bounds: SourceRect | undefined
  try { bounds = await host.bounds(sheetId, scan(), signal) } catch (error) { return sourceFailure(error, signal) }
  if (signal.aborted) throw cancel()
  if (!bounds) return unresolved('cf-pending-data')
  const declared = atoms.map(a => Object.freeze({ ...a.rect }))
  const scanRects = declared.map(r => clampTo(r, bounds!)).filter((r): r is SourceRect => r !== undefined)
  let cells = 0
  for (const r of scanRects) cells += area(r)
  if (cells > CF_LIMITS.dependencyCells) return unresolved('cf-resource-limit')
  const base = measurePayloadBytes({ declared, bounds })!
  const hold = memory.reserve(base)
  if (hold.kind !== 'reserved') return unresolved('cf-resource-limit')
  const columns = new Map<number, Column>()
  let count = 0, capacity = INITIAL_CAPACITY, stringBytes = 0, columnBytes = 0, adopted = false
  // The table is reserved at its capacity; growing to 2x holds old plus new until the copy is done (Lesson 15).
  if (!hold.grow(base + capacity * CF_DEPENDENCY_CELL_BYTES)) { hold.release(); return unresolved('cf-resource-limit') }
  const settle = (): boolean => hold.grow(base + columnBytes + stringBytes + capacity * CF_DEPENDENCY_CELL_BYTES)
  try {
    if (scanRects.length > 0) {
      try {
        for await (const page of host.pages(sheetId, scanRects, scan(), signal)) {
          try {
            const { rows, cols, values } = page.value
            for (let i = 0; i < values.length; i++) {
              if (signal.aborted) throw cancel()
              const row = rows[i]!, col = cols[i]!
              if (!scanRects.some(r => inRect(r, row, col))) continue
              const spent = budget.spend(1)
              if (spent === 'refused') return unresolved('cf-resource-limit')
              if (spent === 'yield') { await hostTask(); if (signal.aborted) throw cancel(); budget.resume() }
              const literal = values[i]!
              const element = cfLiteralToElement(literal)
              const units = literal.kind === 'string' ? literal.value.length : 0
              let column = columns.get(col)
              if (count + 1 > capacity) {
                const next = capacity * 2
                if (!hold.grow(base + columnBytes + stringBytes + (capacity + next) * CF_DEPENDENCY_CELL_BYTES)) return unresolved('cf-resource-limit')
                capacity = next
              }
              if (!column) {
                columnBytes += CF_DEPENDENCY_COLUMN_BYTES
                if (!settle()) return unresolved('cf-resource-limit')
                column = { rows: [], values: [], units: [], sorted: true }
                columns.set(col, column)
              }
              stringBytes += units * 2
              if (!settle()) return unresolved('cf-resource-limit')
              const last = column.rows.length - 1
              if (last >= 0 && column.rows[last] === row) continue // the same cell met twice
              if (last >= 0 && column.rows[last]! > row) column.sorted = false
              column.rows.push(row); column.values.push(element === undefined ? UNAVAILABLE : element); column.units.push(units)
              count++
            }
          } finally { page.release() }
        }
      } catch (error) {
        if (error instanceof PackageError && error.code === 'cancelled') throw error
        if (signal.aborted) throw cancel()
        return sourceFailure(error, signal)
      }
    }
    // An out-of-order source is sorted here, charged n log n before the work.
    for (const column of columns.values()) {
      if (column.sorted) continue
      const spent = budget.spend(sortUnits(column.rows.length))
      if (spent === 'refused') return unresolved('cf-resource-limit')
      const order = column.rows.map((_, i) => i).sort((a, b) => column.rows[a]! - column.rows[b]!)
      const rowsOut: number[] = [], valuesOut: Slot[] = [], unitsOut: number[] = []
      for (const i of order) {
        if (rowsOut.length && rowsOut[rowsOut.length - 1] === column.rows[i]) { count--; continue }
        rowsOut.push(column.rows[i]!); valuesOut.push(column.values[i]!); unitsOut.push(column.units[i]!)
      }
      column.rows = rowsOut; column.values = valuesOut; column.units = unitsOut; column.sorted = true
      if (spent === 'yield') { await hostTask(); if (signal.aborted) throw cancel(); budget.resume() }
    }
    if (signal.aborted) throw cancel()
    const snapshot = createSnapshot(sheetId, bounds, declared, columns, count, () => lease.release())
    const lease = hold.adopt(snapshot as CfDependencySnapshot, base + columnBytes + stringBytes + capacity * CF_DEPENDENCY_CELL_BYTES)
    adopted = true
    return { kind: 'ready', value: lease }
  } finally {
    if (!adopted) hold.release()
  }
}
function sourceFailure(error: unknown, signal: AbortSignal): CfOutcome<never> {
  if (error instanceof PackageError && error.code === 'cancelled') throw error
  if (signal.aborted) throw cancel()
  const message = error instanceof Error ? error.message : ''
  if (message === 'CF cancelled') throw cancel()
  if (message === 'CF work limit' || (error instanceof PackageError && error.code === 'resource-limit')) return unresolved('cf-resource-limit')
  // 'CF unavailable formula value' (render.ts residentSource) and anything else: the footprint is not readable now.
  return unresolved('cf-formula-dependency')
}

function createSnapshot(sheetId: string, bounds: SourceRect, declared: readonly SourceRect[], columns: Map<number, Column>, count: number, release: () => void): CfDependencySnapshot {
  const colKeys = [...columns.keys()].sort((a, b) => a - b)
  let tainted = false, disposed = false
  const lookup = (row: number, col: number): Slot | undefined => {
    const column = columns.get(col)
    if (!column) return undefined
    const i = lowerBound(column.rows, row)
    return column.rows[i] === row ? column.values[i] : undefined
  }
  /** A declared rect is the call geometry as written (a whole column stays a whole column); only its part inside the
   * populated bounds was scanned, and a cell outside the bounds is known blank (bounds are the populated extent, as
   * statistics clamp). An area is covered only when one declared rect contains it whole. */
  const covered = (r: SourceRect): boolean => declared.some(d => containsRect(d, r))
  const value = (row: number, col: number): ElementValue => {
    const slot = lookup(row, col)
    if (slot === UNAVAILABLE) { tainted = true; return NA }
    return slot === undefined ? null : slot
  }
  /** Visits the stored cells of `r` in row-major order; `visit` returns false to stop. */
  const eachIn = (r: SourceRect, visit: (row: number, col: number, units: number) => boolean): void => {
    const c = clampTo(r, bounds)
    if (!c) return
    const perColumn: Array<{ col: number; at: number; end: number; column: Column }> = []
    for (let k = lowerBound(colKeys, c.firstCol); k < colKeys.length && colKeys[k]! <= c.lastCol; k++) {
      const column = columns.get(colKeys[k]!)!
      perColumn.push({ col: colKeys[k]!, at: lowerBound(column.rows, c.firstRow), end: lowerBound(column.rows, c.lastRow + 1), column })
    }
    for (;;) {
      let best: (typeof perColumn)[number] | undefined
      for (const p of perColumn) if (p.at < p.end && (!best || p.column.rows[p.at]! < best.column.rows[best.at]!)) best = p
      if (!best) return
      if (!visit(best.column.rows[best.at]!, best.col, best.column.units[best.at]!)) return
      best.at++
    }
  }
  const countIn = (r: SourceRect, cap: number): number => {
    const c = clampTo(r, bounds)
    if (!c) return 0
    let n = 0
    for (let k = lowerBound(colKeys, c.firstCol); k < colKeys.length && colKeys[k]! <= c.lastCol && n <= cap; k++) {
      const column = columns.get(colKeys[k]!)!
      n += lowerBound(column.rows, c.lastRow + 1) - lowerBound(column.rows, c.firstRow)
    }
    return n
  }
  const snapshot: CfDependencySnapshot = {
    sheetId, entries: count,
    get tainted() { return tainted },
    reset() { tainted = false },
    dispose() { if (!disposed) { disposed = true; columns.clear(); release() } },
    read(sheet, row, col) {
      if (disposed || sheet !== undefined || !declared.some(d => inRect(d, row, col))) { tainted = true; return NA }
      return value(row, col)
    },
    viewFor(effective, budget) {
      if (disposed) return unresolved('cf-formula-dependency')
      // Coverage of every effective area BEFORE any traversal (S5-SUMIF-SHAPE): an uncovered area is never served empty.
      for (const e of effective) if (e.sheet !== undefined || !covered(e.rect)) return unresolved('cf-formula-dependency')
      const cap = CF_LIMITS.checkpointUnits
      const distinct: SourceRect[] = []
      const indexOf = effective.map(e => {
        const i = distinct.findIndex(d => d.firstRow === e.rect.firstRow && d.lastRow === e.rect.lastRow && d.firstCol === e.rect.firstCol && d.lastCol === e.rect.lastCol)
        if (i >= 0) return i
        distinct.push(e.rect)
        return distinct.length - 1
      })
      // Every stored cell of an area is read at least once by its call, so more than the per-call cap refuses now.
      let total = 0
      for (const r of distinct) { total += countIn(r, cap); if (total > cap) return unresolved('cf-resource-limit') }
      const keys = new Set<number>()
      const perArea = distinct.map(r => {
        let cells = 0, stringUnits = 0, longestString = 0
        eachIn(r, (row, col, units) => { cells++; stringUnits += units; longestString = Math.max(longestString, units); keys.add(row * 16384 + col); return true })
        return { cells, stringUnits, longestString }
      })
      const viewCells = keys.size
      if (budget.spend(viewCells) === 'refused') return unresolved('cf-resource-limit')
      const metrics: CfPayloadMetrics = { viewCells, atoms: indexOf.map(i => perArea[i]!) }
      const inView = (row: number, col: number): boolean => distinct.some(r => inRect(r, row, col))
      const charge = (units: number): void => { if (budget.spend(units) === 'refused') throw CF_WORK_REFUSED }
      let ordered: number[] | undefined
      const store: SparseCellStore = {
        read(address) {
          if (address.sheetId !== sheetId || !inView(address.row, address.col)) { tainted = true; return { value: NA, origin: 'input' } }
          charge(1)
          const slot = lookup(address.row, address.col)
          if (slot === UNAVAILABLE) { tainted = true; return { value: NA, origin: 'input' } }
          return slot === undefined ? undefined : { value: slot, origin: 'input' }
        },
        stored(): Iterable<StoredCell> {
          // Entering stored() pays for the native sort that follows (refs.ts `populated`), before it happens.
          charge(sortUnits(viewCells))
          ordered ??= [...keys].sort((a, b) => a - b)
          const list = ordered
          return (function* () {
            for (const key of list) {
              charge(1)
              const row = Math.floor(key / 16384), col = key % 16384
              const slot = lookup(row, col)
              yield { address: { sheetId, row, col }, value: slot === UNAVAILABLE || slot === undefined ? NA : slot, origin: 'input' as const }
            }
          })()
        },
      }
      return { kind: 'ready', value: { store, metrics, read(row, col) {
        if (!inView(row, col)) { tainted = true; return NA }
        charge(1)
        return value(row, col)
      } } }
    },
  }
  return snapshot
}

/* ------------------------------------------------------------------ */
/* The batch driver: one snapshot, hoists, bounded calls, yields.      */
/* ------------------------------------------------------------------ */

export interface CfFormulaTask { readonly rule: CfRule; readonly formula: CfFormula }
export interface CfFormulaBatchOptions {
  readonly source: CfValueSource; readonly sheetId: string; readonly sheetName: string; readonly stamp: CfStamp
  /** Absent: every formula task is unresolved (`cf-evaluation-context-missing`, Ruling S5-L). */
  readonly evaluation: CfEvaluationContext | undefined
  readonly edge: CfEdgeMode | undefined
  readonly memory: CfMemory; readonly signal: AbortSignal
  /** The preparation's per-call pacing budget (its `'yield'` is honoured between cells). */
  readonly budget: CfWorkBudget
  /** One preparation ceiling for every snapshot attempt and every call (default 2 x activeUnits). */
  readonly ceiling?: CfUnitCeiling
  /** A fresh budget per snapshot attempt, capped at what the ceiling has left. */
  readonly attemptBudget?: (maxUnits: number) => CfWorkBudget
  /** Source scans (bounds, pages, hoist scans) pay from their own scan budgets. */
  readonly scanBudget?: () => CfWorkBudget
  /** Hoisting (S5-HOIST) is on only with its evidence; off, a hoistable call pays per cell (and over 128 refuses). */
  readonly hoist?: boolean
  readonly statistics?: CfStatisticsSession
  /** The cells to prepare for (default: every task's ranges). */
  readonly targets?: readonly SourceRect[]
  readonly onCall?: (breakdown: CfCallWorkBreakdown) => void
  readonly probe?: CfFormulaProbe
}
export interface CfFormulaBatch {
  /** Snapshot attempts made (1 shared, plus one per task when the shared one was refused). */
  readonly attempts: number
  readonly ceiling: CfUnitCeiling
  readonly snapshotEntries: number
  outcome(task: number): CfOutcome<void>
  evaluate(task: number, row: number, col: number): CfOutcome<CfResult>
  /** Cells in order; between cells the loop honours a pending yield and cancellation. Throws `cancelled`, disposing. */
  run(cells: Iterable<Readonly<{ row: number; col: number }>>, onResult: (task: number, row: number, col: number, outcome: CfOutcome<CfResult>) => void): Promise<void>
  dispose(): void
}

export async function prepareCfFormulaBatch(tasks: readonly CfFormulaTask[], options: CfFormulaBatchOptions): Promise<CfFormulaBatch> {
  const { signal, memory, evaluation } = options
  const ceiling = options.ceiling ?? createCfUnitCeiling()
  const scanBudget = (): CfWorkBudget => options.scanBudget?.() ?? createCfWorkBudget({ signal, scan: true })
  const attemptBudget = (): CfWorkBudget => options.attemptBudget?.(ceiling.remaining) ?? createCfWorkBudget({ signal, scan: true, maxUnits: ceiling.remaining })
  const taskOutcome: Array<CfOutcome<void>> = tasks.map(() => ({ kind: 'ready', value: undefined }))
  const snapshots: Array<CfLease<CfDependencySnapshot> | undefined> = tasks.map(() => undefined)
  const owned: Array<CfLease<CfDependencySnapshot>> = []
  const hoistLeases: Array<CfLease<CfHoisted>> = []
  const hoists: Array<Readonly<{ plan: readonly CfHoist[]; values: CfHoisted }> | undefined> = tasks.map(() => undefined)
  const invariants = createCfInvariantMemo(memory)
  const pacing: CfFormulaPacing = { yieldPending: false }
  let attempts = 0, disposed = false
  const dispose = (): void => {
    if (disposed) return
    disposed = true
    for (const lease of owned) { lease.value.dispose() }
    for (const lease of hoistLeases) lease.release()
    invariants.dispose()
  }
  try {
    if (!evaluation) {
      for (let i = 0; i < tasks.length; i++) taskOutcome[i] = unresolved('cf-evaluation-context-missing')
    } else {
      let bounds: SourceRect | undefined
      try { bounds = await options.source.bounds(options.sheetId, scanBudget(), signal) } catch (error) { bounds = undefined; if (signal.aborted) throw cancel(); void error }
      if (signal.aborted) throw cancel()
      const footprints: Array<readonly CfFootprintEntry[]> = tasks.map(() => [])
      const cappedCells: number[] = tasks.map(() => 0)
      for (let i = 0; i < tasks.length; i++) {
        const { formula, rule } = tasks[i]!
        if (!bounds) { taskOutcome[i] = unresolved('cf-pending-data'); continue }
        if (formula.dependencySheets.length > 0) { taskOutcome[i] = unresolved('cf-formula-dependency'); continue }
        if (!formula.anchor) { taskOutcome[i] = unresolved('cf-unmeasured-capability'); continue }
        const plan = options.hoist ? planCfHoist(formula) : []
        if (plan.length > 0) {
          if (!options.statistics) { taskOutcome[i] = unresolved('cf-resource-limit'); continue }
          const out = await options.statistics.prepareHoists(rule, plan, options.stamp, options.source, scanBudget(), { signal })
          if (signal.aborted) throw cancel()
          if (out.kind !== 'ready') { taskOutcome[i] = out.kind === 'unresolved' ? out : unresolved('cf-resource-limit'); continue }
          hoistLeases.push(out.value)
          hoists[i] = { plan, values: out.value.value }
        }
        const targets = options.targets ? rule.ranges.flatMap(r => options.targets!.map(t => clampTo(r, t)).filter((x): x is SourceRect => x !== undefined)) : rule.ranges
        const exclude = new Set(plan.map(h => h.node))
        // The cap counts the populated clamp; the snapshot declares the geometry as written (S5-SUMIF-SHAPE: a whole
        // column stays a whole column, so `countAbsent` and coverage see the true area).
        const capped = cfFormulaFootprint(formula, targets, formula.anchor, options.edge, { bounds, exclude })
        const declared = cfFormulaFootprint(formula, targets, formula.anchor, options.edge, { exclude, uncapped: true })
        if (capped === 'over-cap' || declared === 'over-cap') { taskOutcome[i] = unresolved('cf-resource-limit'); continue }
        footprints[i] = declared
        cappedCells[i] = capped.reduce((n, e) => n + area(e.rect), 0)
      }
      // One shared snapshot of the admitted tasks' footprints, in priority order, within the cell cap.
      const union: CfFootprintEntry[] = []
      let cells = 0
      const live: number[] = []
      for (let i = 0; i < tasks.length; i++) {
        if (taskOutcome[i]!.kind !== 'ready') continue
        const add = cappedCells[i]!
        if (cells + add > CF_LIMITS.dependencyCells) { taskOutcome[i] = unresolved('cf-resource-limit'); continue }
        cells += add; live.push(i); union.push(...footprints[i]!)
      }
      const attempt = async (atoms: readonly CfFootprintEntry[]): Promise<CfOutcome<CfLease<CfDependencySnapshot>>> => {
        attempts++
        const b = attemptBudget()
        try {
          const out = await prepareCfDependencies(atoms, options.source, undefined, b, memory, signal, { sheetId: options.sheetId, scanBudget })
          if (out.kind === 'ready') owned.push(out.value)
          return out
        } finally {
          // Every attempt draws from the one preparation ceiling, whatever its outcome.
          if (!ceiling.take(Math.min(b.units, ceiling.remaining))) ceiling.take(ceiling.remaining)
        }
      }
      if (live.length > 0) {
        const shared = await attempt(union)
        if (shared.kind === 'ready') for (const i of live) snapshots[i] = shared.value
        else if (live.length === 1) taskOutcome[live[0]!] = shared
        else for (const i of live) {
          // A refused shared snapshot is retried per rule, each attempt on a fresh budget from the same ceiling (S5-N).
          const own = await attempt(footprints[i]!)
          if (own.kind === 'ready') snapshots[i] = own.value
          else taskOutcome[i] = own
        }
      }
    }
  } catch (error) { dispose(); throw error }
  const materialization = { limit: CF_LIMITS.materializationCells, used: 0 }
  const contexts: Array<CfFormulaContext | undefined> = tasks.map((_, i) => evaluation && snapshots[i] ? {
    sheetName: options.sheetName, evaluation, semantics: cfFormulaSemantics(evaluation), edge: options.edge, materialization, budget: options.budget,
    memory, invariants, ceiling, pacing, ...(hoists[i] ? { hoists: hoists[i] } : {}), ...(options.probe ? { probe: options.probe } : {}), ...(options.onCall ? { onCall: options.onCall } : {}),
  } : undefined)
  const covers = (task: number, row: number, col: number): boolean => tasks[task]!.rule.ranges.some(r => inRect(r, row, col))
  const batch: CfFormulaBatch = {
    get attempts() { return attempts },
    ceiling,
    get snapshotEntries() { return owned.reduce((n, l) => n + l.value.entries, 0) },
    outcome: task => taskOutcome[task]!,
    evaluate(task, row, col) {
      if (disposed) return unresolved('cf-stale-generation')
      const out = taskOutcome[task]!
      if (out.kind !== 'ready') return out
      if (!covers(task, row, col)) return { kind: 'no-match' }
      return evaluateCfFormula(tasks[task]!.formula, { row, col }, snapshots[task]!.value, contexts[task]!)
    },
    async run(cells, onResult) {
      try {
        for (const cell of cells) {
          if (signal.aborted) throw cancel()
          for (let t = 0; t < tasks.length; t++) if (covers(t, cell.row, cell.col)) onResult(t, cell.row, cell.col, batch.evaluate(t, cell.row, cell.col))
          // Between cells, never inside a call (round-2 #2).
          if (pacing.yieldPending) {
            pacing.yieldPending = false
            await hostTask()
            if (signal.aborted) throw cancel()
            options.budget.resume()
          }
        }
      } catch (error) { dispose(); throw error }
    },
    dispose,
  }
  return batch
}
