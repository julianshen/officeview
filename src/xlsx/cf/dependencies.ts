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
 * Cross-sheet values (Task 36, Ruling S5-K): a qualified atom is read from the referenced sheet's own source, resolved
 * through the workbook seam `CfWorkbookSource`, whose values and generation come from one overlay-aware store. Each sheet
 * is its own snapshot (its own bounds, its own keys); the composite dispatches reads and views by sheet, so two sheets'
 * `A1` never share a key and a call's view cardinality counts every sheet it reads.
 *
 * The batch driver below runs the per-cell calls between cooperative checkpoints: a call never yields inside the
 * engine; the loop awaits `hostTask()` between cells whenever the pacing budget said `'yield'`, and checks cancellation
 * there. It is the seam Task 35 plugs into the service's page loop.
 */
import { PackageError } from '../../core/package-error'
import type { AstNode, ElementValue } from '../formula/types'
import type { SparseCellStore, StoredCell } from '../formula/refs'
import { hostTask } from '../view/scheduler'
import { measurePayloadBytes } from '../view/payloadBytes'
import {
  CF_TEXT_UNIT_CHARS, cfFormulaFootprint, cfFormulaSemantics, cfLiteralToElement, createCfInvariantMemo, evaluateCfFormula, planCfHoist,
  type CfCallWorkBreakdown, type CfEdgeMode, type CfEvaluationContext, type CfFootprintEntry, type CfFormulaContext, type CfFormulaPacing,
  type CfFormulaProbe, type CfHoist, type CfHoisted, type CfPayloadMetrics, type CfResult,
} from './formula'
import {
  CF_DEPENDENCY_CELL_BYTES, CF_DEPENDENCY_KEY_BYTES, CF_LIMITS, CF_VIEW_WORKSPACE_BYTES, CF_WORK_REFUSED, createCfUnitCeiling, createCfWorkBudget,
  type CfMemory, type CfReservation, type CfUnitCeiling, type CfWorkBudget,
} from './limits'
import type { CfFormula, CfLease, CfOutcome, CfRule, CfSheetRules, CfStamp, CfValueSource, SourceRect } from './model'
import type { CfStatisticsSession } from './statistics'

/**
 * One referenced sheet of the workbook seam (Task 36, Ruling S5-K): its values and its live value generation come from the
 * same canonical, overlay-aware store, so a value read and the generation it is published under cannot disagree.
 */
export interface CfWorkbookSheet {
  readonly sheetId: string
  readonly name: string
  readonly values: CfValueSource
  /** `'pending'` while the sheet is still importing, `'failed'` when its import ended incomplete. */
  state(): 'ready' | 'pending' | 'failed'
  /** The live value generation, or undefined when this view cannot name it (`cf-dependency-generation-unknown`). */
  valueGen(): number | undefined
}
/** The other sheets of the host's workbook (Task 36). Absent on a leg: every cross-sheet rule withholds (S5-P). */
export interface CfWorkbookSource {
  sheetIdOf(name: string): string | undefined
  sheet(sheetId: string): CfWorkbookSheet | undefined
}
/** Ruling S5-K caps table: at most 16 distinct referenced sheets per sheet's rules (one formula's own cap is enforced at
 * compile, `CF_LIMITS.referencedSheets`; this one bounds the union over the sheet's tasks). */
export const CF_REFERENCED_SHEETS_LIMIT: number = CF_LIMITS.referencedSheets
const lower = (s: string | undefined): string | undefined => s?.toLowerCase()
/** The other sheets `rules` name, case-insensitive distinct in first-appearance order, the host excluded. Reads only the
 * rule origins' `dependencySheets` (no formula is walked), so a sheet without cross-sheet rules does no work. */
export function cfReferencedSheetNames(rules: CfSheetRules | undefined, hostName: string | undefined): string[] {
  const out: string[] = []
  if (!rules) return out
  const seen = new Set<string>()
  const host = lower(hostName)
  for (const rule of rules.rules) {
    const d = rule.definition
    const names = d.kind === 'unresolved' ? undefined : d.dependencySheets
    if (!names) continue
    for (const name of names) {
      const key = name.toLowerCase()
      if (key === host || seen.has(key)) continue
      seen.add(key); out.push(name)
    }
  }
  return out
}
/**
 * The admitted dependency set (fix round 1, #7): the first 16 distinct referenced names in rule order, known or not. It
 * is the one set resolution, `dependencyGen`, the reply's `dependencies` and the batch's admission all use; a rule naming
 * a sheet outside it is refused (`cf-resource-limit`) and nothing is resolved for that sheet.
 */
export function cfAdmittedSheetNames(rules: CfSheetRules | undefined, hostName: string | undefined): string[] {
  return cfReferencedSheetNames(rules, hostName).slice(0, CF_REFERENCED_SHEETS_LIMIT)
}
/** The admitted referenced sheets the workbook knows, resolved once (unknown names are skipped: their rules never paint). */
function resolveReferenced(rules: CfSheetRules | undefined, hostName: string | undefined, workbook: CfWorkbookSource): CfWorkbookSheet[] {
  const out: CfWorkbookSheet[] = []
  for (const name of cfAdmittedSheetNames(rules, hostName)) {
    const id = workbook.sheetIdOf(name)
    const sheet = id === undefined ? undefined : workbook.sheet(id)
    if (sheet && !out.some(s => s.sheetId === sheet.sheetId)) out.push(sheet)
  }
  return out
}
const sumGen = (sheets: readonly CfWorkbookSheet[]): number => {
  let gen = 0
  for (const s of sheets) { const g = s.valueGen(); if (g !== undefined && Number.isSafeInteger(g)) gen += g }
  return gen
}
/**
 * `CfStamp.dependencyGen` (Ruling S5-K): the sum of the referenced sheets' live value generations, the host excluded; 0
 * when no rule names another sheet, without reading the workbook. A sheet whose generation is unknown contributes nothing
 * (its rules are refused, `cf-dependency-generation-unknown`).
 */
export function cfDependencyGenOf(rules: CfSheetRules | undefined, hostName: string | undefined, workbook: CfWorkbookSource | undefined): number {
  if (!workbook || cfReferencedSheetNames(rules, hostName).length === 0) return 0
  return sumGen(resolveReferenced(rules, hostName, workbook))
}
/** The referenced sheets' `{ sheetId, valueGen }` pairs, for the worker reply (`CfReadResult.dependencies`). */
export function cfDependencyIdentities(rules: CfSheetRules | undefined, hostName: string | undefined, workbook: CfWorkbookSource | undefined): Array<{ sheetId: string; valueGen: number }> {
  if (!workbook || cfReferencedSheetNames(rules, hostName).length === 0) return []
  const out: Array<{ sheetId: string; valueGen: number }> = []
  for (const s of resolveReferenced(rules, hostName, workbook)) { const g = s.valueGen(); if (g !== undefined) out.push({ sheetId: s.sheetId, valueGen: g }) }
  return out
}
/**
 * The host source with a live `dependencyGen` on every `stamp()` (Ruling S5-K): `CfPrepared.at` re-reads it, so a
 * referenced sheet's commit makes the old batch stale synchronously, exactly as a host commit does. Without a cross-sheet
 * rule (or without a workbook) the host source is returned unchanged and its stamp is never touched.
 */
export function withCfDependencyGen(host: CfValueSource, rules: CfSheetRules | undefined, hostName: string | undefined, workbook: CfWorkbookSource | undefined): CfValueSource {
  if (!workbook || cfReferencedSheetNames(rules, hostName).length === 0) return host
  const sheets = resolveReferenced(rules, hostName, workbook)
  return {
    bounds: (id, budget, signal) => host.bounds(id, budget, signal),
    pages: (id, rects, budget, signal) => host.pages(id, rects, budget, signal),
    ...(host.numberFormats ? { numberFormats: host.numberFormats } : {}),
    stamp: id => ({ ...host.stamp(id), dependencyGen: sumGen(sheets) }),
  }
}
/** One call's restricted view: its store for the engine, its measured payload, and a charged single-cell reader. */
export type CfDependencyView = Readonly<{ store: SparseCellStore; metrics: CfPayloadMetrics
  /** A charged single-cell read; `sheet` names a referenced sheet (absent: the host). */
  read(row: number, col: number, sheet?: string): ElementValue }>
export interface CfDependencySnapshot {
  readonly sheetId: string
  /** Every sheet the snapshot holds, the host first (its name absent when the caller gave none); the engine's
   * `ReferenceServices` resolve names and ids from exactly this list. */
  readonly sheets: readonly Readonly<{ sheetId: string; name?: string }>[]
  /** Stored entries held. */
  readonly entries: number
  /** A snapshot read (uncharged): taints on a cell outside the snapshot, another sheet, or an unavailable value. */
  read(sheet: string | undefined, row: number, col: number): ElementValue
  readonly tainted: boolean
  reset(): void
  dispose(): void
  /** The per-call store over the call's effective areas. Refused before any traversal when an area is not fully covered,
   * and before anything is built unless `workspace` already holds `CF_VIEW_WORKSPACE_BYTES` (the caller reserves it). */
  viewFor(effective: readonly CfFootprintEntry[], budget: CfWorkBudget, workspace: CfReservation): CfOutcome<CfDependencyView>
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
/** `flags`: 0 a non-string, 1 an ASCII string, 2 a string with a non-ASCII unit (round 3, finding 4). */
interface Column { rows: number[]; values: Slot[]; units: number[]; flags: number[]; sorted: boolean }
/** One new column's bookkeeping, DERIVED from the payload measurer (a map entry holding an empty column record). */
export const CF_DEPENDENCY_COLUMN_BYTES: number = measurePayloadBytes(new Map([[0, { rows: [], values: [], units: [], flags: [], sorted: true }]]))! - measurePayloadBytes(new Map())!
const INITIAL_CAPACITY = 16
const isAsciiText = (s: string): boolean => { for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) >= 0x80) return false; return true }
const W = 16384
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
  budget: CfWorkBudget, memory: CfMemory, signal: AbortSignal, options: Readonly<{ sheetId: string; sheetName?: string; scanBudget?: () => CfWorkBudget }>): Promise<CfOutcome<CfLease<CfDependencySnapshot>>> {
  if (signal.aborted) throw cancel()
  // A qualifier naming the host is the host (case-insensitive); any other qualifier is a referenced sheet (S5-K).
  const hostKey = lower(options.sheetName)
  const own: CfFootprintEntry[] = [], others = new Map<string, { sheet: CfWorkbookSheet; name: string; atoms: CfFootprintEntry[] }>()
  for (const a of atoms) {
    if (a.sheet === undefined || lower(a.sheet) === hostKey) { own.push({ rect: a.rect }); continue }
    // No workbook, or no host name to tell the host from another sheet: a qualified atom cannot be read (S5-P).
    if (!workbook || hostKey === undefined) return unresolved('cf-formula-dependency')
    const key = a.sheet.toLowerCase()
    let group = others.get(key)
    if (!group) {
      const id = workbook.sheetIdOf(a.sheet)
      const sheet = id === undefined ? undefined : workbook.sheet(id)
      if (!sheet) return unresolved('cf-formula-dependency')
      if (others.size >= CF_REFERENCED_SHEETS_LIMIT) return unresolved('cf-resource-limit')
      const state = sheet.state()
      if (state === 'pending') return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: [sheet.sheetId] }
      if (state === 'failed') return unresolved('cf-read-failed')
      if (sheet.valueGen() === undefined) return unresolved('cf-dependency-generation-unknown')
      others.set(key, group = { sheet, name: a.sheet, atoms: [] })
    }
    group.atoms.push({ rect: a.rect })
  }
  const hostSnap = await prepareSheetSnapshot(own, host, budget, memory, signal, { sheetId: options.sheetId, ...(options.scanBudget ? { scanBudget: options.scanBudget } : {}) }, false)
  if (others.size === 0 || hostSnap.kind !== 'ready') return hostSnap
  const subs = new Map<string, CfLease<CfDependencySnapshot>>()
  const releaseAll = (): void => { hostSnap.value.value.dispose(); for (const l of subs.values()) l.value.dispose() }
  try {
    for (const [key, group] of others) {
      const out = await prepareSheetSnapshot(group.atoms, group.sheet.values, budget, memory, signal, { sheetId: group.sheet.sheetId, ...(options.scanBudget ? { scanBudget: options.scanBudget } : {}) }, true)
      if (out.kind !== 'ready') { releaseAll(); return out }
      subs.set(key, out.value)
    }
  } catch (error) { releaseAll(); throw error }
  const composite = composeSnapshots(hostSnap.value.value, options.sheetName, [...others].map(([key, g]) => ({ key, name: g.name, snapshot: subs.get(key)!.value })))
  const leases = [hostSnap.value, ...subs.values()]
  const bytes = leases.reduce((n, l) => n + l.bytes, 0)
  return { kind: 'ready', value: { value: composite, bytes, release: () => composite.dispose(),
    transfer: () => ({ kind: 'unresolved', reason: 'cf-dependency-snapshot-transfer' }) } }
}

/** One sheet's snapshot (the Task 34 single-sheet table). `dependency`: a referenced sheet, whose pending page or fault
 * names it (`pendingSheetIds`) instead of reading as a permanent dependency failure. */
async function prepareSheetSnapshot(atoms: readonly CfFootprintEntry[], host: CfValueSource,
  budget: CfWorkBudget, memory: CfMemory, signal: AbortSignal, options: Readonly<{ sheetId: string; scanBudget?: () => CfWorkBudget }>, dependency: boolean): Promise<CfOutcome<CfLease<CfDependencySnapshot>>> {
  const sheetId = options.sheetId
  const failure = (error: unknown): CfOutcome<never> => sourceFailure(error, signal, dependency ? sheetId : undefined)
  const scan = (): CfWorkBudget => options.scanBudget?.() ?? createCfWorkBudget({ signal, scan: true })
  let bounds: SourceRect | undefined
  try { bounds = await host.bounds(sheetId, scan(), signal) } catch (error) { return failure(error) }
  if (signal.aborted) throw cancel()
  if (!bounds) return dependency ? { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: [sheetId] } : unresolved('cf-pending-data')
  const declared = atoms.map(a => Object.freeze({ ...a.rect }))
  const scanRects = declared.map(r => clampTo(r, bounds!)).filter((r): r is SourceRect => r !== undefined)
  let cells = 0
  for (const r of scanRects) cells += area(r)
  if (cells > CF_LIMITS.dependencyCells) return unresolved('cf-resource-limit')
  const base = measurePayloadBytes({ declared, bounds })!
  const hold = memory.reserve(base)
  if (hold.kind !== 'reserved') return unresolved('cf-resource-limit')
  const columns = new Map<number, Column>()
  // The row-major key index (row * 16384 + col), in arrival order; rebuilt (charged) when the source was out of order.
  let keys: number[] = [], keysSorted = true, lastKey = -1
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
              const literal = values[i]!
              // One unit per entry, plus the ASCII scan of a string at the text rate (round 3, finding 4).
              const spent = budget.spend(1 + (literal.kind === 'string' ? Math.ceil(literal.value.length / CF_TEXT_UNIT_CHARS) : 0))
              if (spent === 'refused') return unresolved('cf-resource-limit')
              if (spent === 'yield') { await hostTask(); if (signal.aborted) throw cancel(); budget.resume() }
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
                column = { rows: [], values: [], units: [], flags: [], sorted: true }
                columns.set(col, column)
              }
              stringBytes += units * 2
              if (!settle()) return unresolved('cf-resource-limit')
              const last = column.rows.length - 1
              if (last >= 0 && column.rows[last] === row) continue // the same cell met twice
              if (last >= 0 && column.rows[last]! > row) column.sorted = false
              column.rows.push(row); column.values.push(element === undefined ? UNAVAILABLE : element); column.units.push(units)
              column.flags.push(literal.kind !== 'string' ? 0 : isAsciiText(literal.value) ? 1 : 2)
              const key = row * W + col
              if (key <= lastKey) keysSorted = false
              lastKey = key
              keys.push(key)
              count++
            }
          } finally { page.release() }
        }
      } catch (error) {
        if (error instanceof PackageError && error.code === 'cancelled') throw error
        if (signal.aborted) throw cancel()
        return failure(error)
      }
    }
    // An out-of-order source is sorted here, charged n log n before the work.
    for (const column of columns.values()) {
      if (column.sorted) continue
      const spent = budget.spend(sortUnits(column.rows.length))
      if (spent === 'refused') return unresolved('cf-resource-limit')
      const order = column.rows.map((_, i) => i).sort((a, b) => column.rows[a]! - column.rows[b]!)
      const rowsOut: number[] = [], valuesOut: Slot[] = [], unitsOut: number[] = [], flagsOut: number[] = []
      for (const i of order) {
        if (rowsOut.length && rowsOut[rowsOut.length - 1] === column.rows[i]) { count--; continue }
        rowsOut.push(column.rows[i]!); valuesOut.push(column.values[i]!); unitsOut.push(column.units[i]!); flagsOut.push(column.flags[i]!)
      }
      column.rows = rowsOut; column.values = valuesOut; column.units = unitsOut; column.flags = flagsOut; column.sorted = true
      if (spent === 'yield') { await hostTask(); if (signal.aborted) throw cancel(); budget.resume() }
    }
    if (!keysSorted) {
      // Rebuilt from the normalised columns (so it follows both column dedupes), then sorted: n log n charged first, and
      // the new index held beside the old one until the old is dropped.
      const spent = budget.spend(count + sortUnits(count))
      if (spent === 'refused') return unresolved('cf-resource-limit')
      if (!hold.grow(base + columnBytes + stringBytes + capacity * CF_DEPENDENCY_CELL_BYTES + count * CF_DEPENDENCY_KEY_BYTES)) return unresolved('cf-resource-limit')
      const rebuilt: number[] = []
      for (const [col, column] of columns) for (const row of column.rows) rebuilt.push(row * W + col)
      rebuilt.sort((a, b) => a - b)
      keys = rebuilt
      if (spent === 'yield') { await hostTask(); if (signal.aborted) throw cancel(); budget.resume() }
    }
    if (signal.aborted) throw cancel()
    const snapshot = createSnapshot(sheetId, bounds, declared, columns, keys, count, () => lease.release())
    const lease = hold.adopt(snapshot as CfDependencySnapshot, base + columnBytes + stringBytes + capacity * CF_DEPENDENCY_CELL_BYTES)
    adopted = true
    return { kind: 'ready', value: lease }
  } finally {
    if (!adopted) hold.release()
  }
}
function sourceFailure(error: unknown, signal: AbortSignal, dependencySheetId?: string): CfOutcome<never> {
  if (error instanceof PackageError && error.code === 'cancelled') throw error
  if (signal.aborted) throw cancel()
  // A referenced sheet's page fault (pagedSource `CfPageFault`, matched by name: no worker import here): an unimported
  // page is pending on that sheet, a page whose generation moved is stale; neither is a permanent dependency failure.
  if (dependencySheetId !== undefined && error instanceof Error && error.name === 'CfPageFault') {
    const fault = (error as Error & { fault?: string }).fault
    if (fault === 'incomplete') return { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: [dependencySheetId] }
    if (fault === 'stale') return unresolved('cf-stale-generation')
  }
  const message = error instanceof Error ? error.message : ''
  if (message === 'CF cancelled') throw cancel()
  if (message === 'CF work limit' || (error instanceof PackageError && error.code === 'resource-limit')) return unresolved('cf-resource-limit')
  // 'CF unavailable formula value' (render.ts residentSource) and anything else: the footprint is not readable now.
  return unresolved('cf-formula-dependency')
}

function createSnapshot(sheetId: string, bounds: SourceRect, declared: readonly SourceRect[], columns: Map<number, Column>, keys: readonly number[], count: number, release: () => void): CfDependencySnapshot {
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
  const entryAt = (row: number, col: number): [number, number] => {
    const column = columns.get(col)!
    const i = lowerBound(column.rows, row)
    return [column.units[i]!, column.flags[i]!]
  }
  /**
   * Visits the stored cells of `r` (finding 1). The walk takes the cheaper of two indexes, each priced in O(log n)
   * first: by column (the populated columns in the rect's column span) or by row (the keys in its row band). Every step
   * that finds no cell (an empty column, or a key past the rect's columns that makes the walk jump) calls `step`, which
   * charges it and may refuse; every stored cell calls `hit`. Neither walk touches a column or key outside the rect's
   * span, so a sparse call never scans the snapshot.
   */
  const walkIn = (r: SourceRect, step: () => boolean, hit: (row: number, col: number, units: number, flag: number) => boolean): boolean => {
    const c = clampTo(r, bounds)
    if (!c) return true
    const k0 = lowerBound(colKeys, c.firstCol), k1 = lowerBound(colKeys, c.lastCol + 1)
    const b0 = lowerBound(keys, c.firstRow * W), b1 = lowerBound(keys, (c.lastRow + 1) * W)
    if (k1 - k0 <= b1 - b0) {
      for (let k = k0; k < k1; k++) {
        const col = colKeys[k]!, column = columns.get(col)!
        const at = lowerBound(column.rows, c.firstRow), end = lowerBound(column.rows, c.lastRow + 1)
        if (at === end && !step()) return false
        for (let i = at; i < end; i++) if (!hit(column.rows[i]!, col, column.units[i]!, column.flags[i]!)) return false
      }
      return true
    }
    const last = c.lastRow * W + c.lastCol
    for (let i = lowerBound(keys, c.firstRow * W + c.firstCol); i < keys.length && keys[i]! <= last;) {
      const key = keys[i]!, row = Math.floor(key / W), col = key % W
      if (col < c.firstCol) { if (!step()) return false; i = lowerBound(keys, row * W + c.firstCol); continue }
      if (col > c.lastCol) { if (!step()) return false; i = lowerBound(keys, (row + 1) * W + c.firstCol); continue }
      const [units, flag] = entryAt(row, col)
      if (!hit(row, col, units, flag)) return false
      i++
    }
    return true
  }
  const snapshot: CfDependencySnapshot = {
    sheetId, entries: count, sheets: [{ sheetId }],
    get tainted() { return tainted },
    reset() { tainted = false },
    dispose() { if (!disposed) { disposed = true; columns.clear(); release() } },
    read(sheet, row, col) {
      if (disposed || sheet !== undefined || !declared.some(d => inRect(d, row, col))) { tainted = true; return NA }
      return value(row, col)
    },
    viewFor(effective, budget, workspace) {
      if (disposed) return unresolved('cf-formula-dependency')
      // Nothing is built until the caller's reservation holds the worst-case workspace (finding 1).
      if (workspace.bytes < CF_VIEW_WORKSPACE_BYTES) return unresolved('cf-resource-limit')
      const cap = CF_LIMITS.checkpointUnits
      if (effective.length > cap) return unresolved('cf-resource-limit')
      let distinctCount = 0
      // Coverage of every effective area BEFORE any traversal (S5-SUMIF-SHAPE): an uncovered area is never served empty.
      for (const e of effective) if (e.sheet !== undefined || !covered(e.rect)) return unresolved('cf-formula-dependency')
      const distinct: SourceRect[] = []
      const indexOf = effective.map(e => {
        const i = distinct.findIndex(d => d.firstRow === e.rect.firstRow && d.lastRow === e.rect.lastRow && d.firstCol === e.rect.firstCol && d.lastCol === e.rect.lastCol)
        if (i >= 0) return i
        distinct.push(e.rect); distinctCount++
        return distinct.length - 1
      })
      // The workspace is priced at CF_LIMITS.formulaAtoms distinct areas (one per admitted atom).
      if (distinctCount > CF_LIMITS.formulaAtoms) return unresolved('cf-resource-limit')
      // Every stored cell of an area is read at least once by its call, so more than the per-call cap refuses; each
      // empty index step is charged as it is taken (the call's `index` term).
      const keySet = new Set<number>()
      let hits = 0, index = 0, refused = false
      const step = (): boolean => { if (budget.spend(1) === 'refused') { refused = true; return false } index++; return true }
      const perArea: Array<{ cells: number; stringUnits: number; longestString: number; strings: number; nonAscii: number }> = []
      for (const r of distinct) {
        const m = { cells: 0, stringUnits: 0, longestString: 0, strings: 0, nonAscii: 0 }
        const ok = walkIn(r, step, (row, col, units, flag) => {
          if (++hits > cap) return false
          // Each new view key is the view build's unit, spent as it is added (round 3): a walk refused part-way has
          // still paid for what it built.
          const key = row * W + col
          if (!keySet.has(key)) { if (budget.spend(1) === 'refused') { refused = true; return false } keySet.add(key) }
          m.cells++; m.stringUnits += units; m.longestString = Math.max(m.longestString, units)
          if (flag > 0) m.strings++
          if (flag === 2) m.nonAscii++
          return true
        })
        if (!ok || refused) return unresolved('cf-resource-limit')
        perArea.push(m)
      }
      const keys = keySet
      const viewCells = keys.size // already spent, one per key
      const metrics: CfPayloadMetrics = { viewCells, index, atoms: indexOf.map(i => perArea[i]!) }
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
              const row = Math.floor(key / W), col = key % W
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

/**
 * The host snapshot plus one per referenced sheet (Ruling S5-K). Reads and views dispatch by sheet; a view's metrics sum
 * every sheet's (its `viewCells` is the call's whole cardinality, so the bound prices every sheet it reads) and its atoms
 * keep the call's order. The distinct-area total across sheets stays within the one workspace the caller reserved.
 */
function composeSnapshots(host: CfDependencySnapshot, hostName: string | undefined, others: readonly { key: string; name: string; snapshot: CfDependencySnapshot }[]): CfDependencySnapshot {
  const hostKey = lower(hostName)
  const byKey = new Map(others.map(o => [o.key, o.snapshot] as const))
  const byId = new Map<string, CfDependencySnapshot>([[host.sheetId, host], ...others.map(o => [o.snapshot.sheetId, o.snapshot] as const)])
  const all = [host, ...others.map(o => o.snapshot)]
  let disposed = false, tainted = false
  const subOf = (sheet: string | undefined): CfDependencySnapshot | undefined => (sheet === undefined || lower(sheet) === hostKey ? host : byKey.get(sheet.toLowerCase()))
  return {
    sheetId: host.sheetId,
    sheets: [{ sheetId: host.sheetId, ...(hostName !== undefined ? { name: hostName } : {}) }, ...others.map(o => ({ sheetId: o.snapshot.sheetId, name: o.name }))],
    get entries() { return all.reduce((n, s) => n + s.entries, 0) },
    get tainted() { return tainted || all.some(s => s.tainted) },
    reset() { tainted = false; for (const s of all) s.reset() },
    dispose() { if (disposed) return; disposed = true; for (const s of all) s.dispose() },
    read(sheet, row, col) {
      const sub = disposed ? undefined : subOf(sheet)
      if (!sub) { tainted = true; return NA }
      return sub.read(undefined, row, col)
    },
    viewFor(effective, budget, workspace) {
      if (disposed) return unresolved('cf-formula-dependency')
      const groups = new Map<CfDependencySnapshot, { entries: CfFootprintEntry[]; at: number[] }>()
      const distinct = new Set<string>()
      for (let i = 0; i < effective.length; i++) {
        const e = effective[i]!, sub = subOf(e.sheet)
        if (!sub) return unresolved('cf-formula-dependency')
        distinct.add(`${sub.sheetId}|${e.rect.firstRow}:${e.rect.lastRow}:${e.rect.firstCol}:${e.rect.lastCol}`)
        let g = groups.get(sub)
        if (!g) groups.set(sub, g = { entries: [], at: [] })
        g.entries.push({ rect: e.rect }); g.at.push(i)
      }
      // One workspace for every sheet's view: the distinct areas of the whole call are capped as one view's are.
      if (distinct.size > CF_LIMITS.formulaAtoms) return unresolved('cf-resource-limit')
      const atoms: Array<CfPayloadMetrics['atoms'][number]> = new Array(effective.length)
      let viewCells = 0, index = 0
      const views = new Map<string, CfDependencyView>()
      for (const [sub, g] of groups) {
        const v = sub.viewFor(g.entries, budget, workspace)
        if (v.kind !== 'ready') return v
        views.set(sub.sheetId, v.value)
        viewCells += v.value.metrics.viewCells; index += v.value.metrics.index ?? 0
        g.at.forEach((pos, j) => { atoms[pos] = v.value.metrics.atoms[j]! })
      }
      const viewOf = (sheet: string | undefined): CfDependencyView | undefined => { const sub = subOf(sheet); return sub ? views.get(sub.sheetId) : undefined }
      const store: SparseCellStore = {
        read(address) {
          const v = byId.has(address.sheetId) ? views.get(address.sheetId) : undefined
          if (!v) { tainted = true; return { value: NA, origin: 'input' } }
          return v.store.read(address)
        },
        stored(): Iterable<StoredCell> {
          // Each sheet's view charges its own sort and elements on entry (their sum never exceeds the bound's v log v).
          const parts = [...views.values()].map(v => v.store.stored())
          return (function* () { for (const p of parts) yield* p })()
        },
      }
      return { kind: 'ready', value: { store, metrics: { viewCells, index, atoms }, read(row, col, sheet) {
        const v = viewOf(sheet)
        if (!v) { tainted = true; return NA }
        return v.read(row, col)
      } } }
    },
  }
}

/* ------------------------------------------------------------------ */
/* The batch driver: one snapshot, hoists, bounded calls, yields.      */
/* ------------------------------------------------------------------ */

export interface CfFormulaTask { readonly rule: CfRule; readonly formula: CfFormula }
export interface CfFormulaBatchOptions {
  /** `sheetName`: the host's name, the formula context's `currentSheet`. Absent (the sheet-id fallback), a formula naming
   * any sheet withholds: the host cannot be told apart from another sheet. */
  readonly source: CfValueSource; readonly sheetId: string; readonly sheetName?: string; readonly stamp: CfStamp
  /** The other sheets (Task 36, Ruling S5-K). Absent: every cross-sheet task withholds (`cf-formula-dependency`). */
  readonly workbook?: CfWorkbookSource
  /** An unknown sheet name reads `#REF!` (only under the `formula:cross-sheet-missing` reading); absent: unresolved. */
  readonly missingSheet?: 'ref-error'
  /** The sheet's admitted dependency set (`cfAdmittedSheetNames`); a task naming any other sheet is `cf-resource-limit`.
   * Absent (a direct caller), the batch admits the first 16 distinct names its tasks meet. */
  readonly admittedSheets?: readonly string[]
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
  readonly hoist?: boolean | ((plan: readonly CfHoist[]) => boolean)
  readonly statistics?: CfStatisticsSession
  /** The cells to prepare for (default: every task's ranges). */
  readonly targets?: readonly SourceRect[]
  readonly onCall?: (breakdown: CfCallWorkBreakdown) => void
  readonly onSettled?: (spent: number, breakdown: CfCallWorkBreakdown) => void
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

/** The value generations a batch read (the fields the statistics session's source recheck compares; a rule change makes
 * a new batch, so cfGen, which a source stamp does not own, is not compared). */
const sameDataGeneration = (a: CfStamp, b: CfStamp): boolean => a.owner === b.owner && a.sheetId === b.sheetId && a.workbookValueGen === b.workbookValueGen
  && a.structureGen === b.structureGen && (a.dependencyGen ?? 0) === (b.dependencyGen ?? 0)
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
  /** Lower-cased unknown sheet names that read `#REF!` (only under `missingSheet: 'ref-error'`). */
  const missing = new Set<string>()
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
      const hostKey = lower(options.sheetName)
      // Ruling S5-K: the referenced sheets of the whole sheet's tasks, at most 16; their bounds clamp whole-axis atoms.
      const referenced = new Map<string, CfWorkbookSheet>()
      const admitted = options.admittedSheets ? new Set(options.admittedSheets.map(n => n.toLowerCase())) : undefined
      const depBounds = new Map<string, SourceRect>()
      for (let i = 0; i < tasks.length; i++) {
        const { formula, rule } = tasks[i]!
        if (!bounds) { taskOutcome[i] = unresolved('cf-pending-data'); continue }
        if (!formula.anchor) { taskOutcome[i] = unresolved('cf-unmeasured-capability'); continue }
        const others = formula.dependencySheets.filter(name => name.toLowerCase() !== hostKey)
        if (formula.dependencySheets.length > 0 && hostKey === undefined) { taskOutcome[i] = unresolved('cf-formula-dependency'); continue }
        if (others.length > 0) {
          if (!options.workbook) { taskOutcome[i] = unresolved('cf-formula-dependency'); continue }
          let refused: CfOutcome<void> | undefined
          for (const name of others) {
            const key = name.toLowerCase()
            if (referenced.has(key) || missing.has(key)) continue
            if (admitted && !admitted.has(key)) { refused = unresolved('cf-resource-limit'); break }
            const id = options.workbook.sheetIdOf(name)
            const sheet = id === undefined ? undefined : options.workbook.sheet(id)
            if (!sheet) {
              // Decided before invocation: an unknown sheet never reaches the engine as a silent `#REF!` no-match.
              if (options.missingSheet === 'ref-error') { missing.add(key); continue }
              refused = unresolved('cf-formula-dependency'); break
            }
            if (referenced.size >= CF_REFERENCED_SHEETS_LIMIT) { refused = unresolved('cf-resource-limit'); break }
            if (sheet.valueGen() === undefined) { refused = unresolved('cf-dependency-generation-unknown'); break }
            const state = sheet.state()
            if (state === 'pending') { refused = { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: [sheet.sheetId] }; break }
            if (state === 'failed') { refused = unresolved('cf-read-failed'); break }
            let b: SourceRect | undefined
            try { b = await sheet.values.bounds(sheet.sheetId, scanBudget(), signal) } catch (error) { b = undefined; if (signal.aborted) throw cancel(); void error }
            if (signal.aborted) throw cancel()
            if (!b) { refused = { kind: 'unresolved', reason: 'cf-pending-data', pendingSheetIds: [sheet.sheetId] }; break }
            referenced.set(key, sheet); depBounds.set(key, b)
          }
          if (refused) { taskOutcome[i] = refused; continue }
        }
        // Per task (Task 35): a predicate decides from the plan's hoist kinds, so an unevidenced kind is never hoisted.
        // Each hoist is scanned from its own sheet's store under that store's generation (Task 36 fix round 1, #8): the
        // host's hoists through the host source and stamp, a qualified range's through the referenced sheet's canonical
        // source and its live stamp (the statistics session's source recheck then compares that sheet's generation), each
        // scan paying from its own scan budget. The values merge into one hoist record for the call.
        const planned = options.hoist ? planCfHoist(formula) : []
        const plan = planned.length > 0 && (typeof options.hoist === 'function' ? options.hoist(planned) : options.hoist === true) ? planned : []
        if (plan.length > 0) {
          if (!options.statistics) { taskOutcome[i] = unresolved('cf-resource-limit'); continue }
          const groups = new Map<string, { sheet?: CfWorkbookSheet; hoists: CfHoist[] }>()
          let unroutable = false
          for (const h of plan) {
            const key = h.range.sheet === undefined || h.range.sheet.toLowerCase() === hostKey ? '' : h.range.sheet.toLowerCase()
            const sheet = key === '' ? undefined : referenced.get(key)
            if (key !== '' && !sheet) { unroutable = true; break }
            let g = groups.get(key)
            if (!g) groups.set(key, g = { ...(sheet ? { sheet } : {}), hoists: [] })
            g.hoists.push(h.range.sheet === undefined ? h : { ...h, range: { rect: h.range.rect } })
          }
          if (unroutable) { taskOutcome[i] = unresolved('cf-formula-dependency'); continue }
          const scalars = new Map<AstNode, ElementValue>(), frequencies = new Map<AstNode, ReadonlyMap<string, number>>(), longest = new Map<AstNode, number>()
          let refusedHoist: CfOutcome<void> | undefined, only: CfHoisted | undefined
          for (const g of groups.values()) {
            const stamp = g.sheet ? g.sheet.values.stamp(g.sheet.sheetId) : options.stamp
            const out = await options.statistics.prepareHoists(rule, g.hoists, stamp, g.sheet ? g.sheet.values : options.source, scanBudget(), { signal })
            // Owned before any further check, so a cancellation here disposes it with the rest.
            if (out.kind === 'ready') hoistLeases.push(out.value)
            if (signal.aborted) throw cancel()
            if (out.kind !== 'ready') { refusedHoist = out.kind === 'unresolved' ? out : unresolved('cf-resource-limit'); break }
            const v = out.value.value
            // One sheet (the common case): its leased record is the call's, nothing new is built. Several: the merged
            // maps hold references to leased values, one entry per hoist (at most the plan's length).
            if (groups.size === 1) { only = v; break }
            for (const [k, x] of v.scalars) scalars.set(k, x)
            for (const [k, x] of v.frequencies) frequencies.set(k, x)
            for (const [k, x] of v.longest) longest.set(k, x)
          }
          if (refusedHoist) { taskOutcome[i] = refusedHoist; continue }
          hoists[i] = { plan, values: only ?? { scalars, frequencies, longest } }
        }
        const targets = options.targets ? rule.ranges.flatMap(r => options.targets!.map(t => clampTo(r, t)).filter((x): x is SourceRect => x !== undefined)) : rule.ranges
        // The cap counts the populated clamp; the snapshot declares the geometry as written (S5-SUMIF-SHAPE: a whole
        // column stays a whole column, so `countAbsent` and coverage see the true area).
        const boundsOf = (sheet: string | undefined): SourceRect | undefined => (sheet === undefined || sheet.toLowerCase() === hostKey ? bounds : depBounds.get(sheet.toLowerCase()))
        const capped = cfFormulaFootprint(formula, targets, formula.anchor, options.edge, { bounds, boundsOf, hoists: plan })
        const declared = cfFormulaFootprint(formula, targets, formula.anchor, options.edge, { hoists: plan, uncapped: true })
        if (capped === 'over-cap' || declared === 'over-cap') { taskOutcome[i] = unresolved('cf-resource-limit'); continue }
        // A missing sheet's atoms read `#REF!` and are never snapshotted.
        footprints[i] = declared.filter(e => e.sheet === undefined || !missing.has(e.sheet.toLowerCase()))
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
          const out = await prepareCfDependencies(atoms, options.source, options.workbook, b, memory, signal, { sheetId: options.sheetId, ...(options.sheetName !== undefined ? { sheetName: options.sheetName } : {}), scanBudget })
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
    sheetName: options.sheetName ?? options.sheetId, evaluation, semantics: cfFormulaSemantics(evaluation), edge: options.edge, materialization, budget: options.budget,
    ...(missing.size > 0 ? { missingSheets: missing } : {}),
    memory, invariants, ceiling, pacing, ...(hoists[i] ? { hoists: hoists[i] } : {}), ...(options.probe ? { probe: options.probe } : {}), ...(options.onCall ? { onCall: options.onCall } : {}), ...(options.onSettled ? { onSettled: options.onSettled } : {}),
  } : undefined)
  const covers = (task: number, row: number, col: number): boolean => tasks[task]!.rule.ranges.some(r => inRect(r, row, col))
  const batch: CfFormulaBatch = {
    get attempts() { return attempts },
    ceiling,
    get snapshotEntries() { return owned.reduce((n, l) => n + l.value.entries, 0) },
    outcome: task => taskOutcome[task]!,
    evaluate(task, row, col) {
      // A live batch outlives the session's cache entries (its hoist leases are refcounted), so it checks its own
      // currency: a commit to the host sheet, a referenced sheet, the structure or the rules makes it stale.
      if (disposed || !sameDataGeneration(options.source.stamp(options.sheetId), options.stamp)) return unresolved('cf-stale-generation')
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
