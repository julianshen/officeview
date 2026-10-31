/**
 * The pivot session (plan Task 25, spec §6.4, §7.1, §7.6): what both paths run when a pivot becomes interactive. The
 * resident viewer runs it on the main thread; the worker runs it for the paged path (Task 27). Lazy chunk `pivots`.
 *
 * Open, in order (a refusal keeps the saved cells; the pivot stays display-only with that diagnostic):
 * 1. eligibility (`engineEligibility`, §7.6: records part, source type, OLAP);
 * 2. the faithfulness check (`checkSavedLayout`, §6.4) on the decoded store;
 * 3. the caption rule, read off the saved state's implicit captions (captions.ts);
 * 4. the chart gate on the saved layout: a pivot with linked charts and no gate refuses `pivot-chart-mapping`, so it is
 *    never interactive before Task 30's gate exists.
 *
 * Apply is pure (the session changes only on `commit`), and refuses, as a typed result: a session released; an
 * interaction kind with no passing oracle (`ORACLE_VERIFIED_KINDS`); the delta's own refusal (delta.ts); any engine
 * refusal; a caption with no pattern (`pivot-caption-locale`); an overlap; the chart gate's refusal.
 *
 * Rulings recorded here (Task 25):
 * - Overlap (§7.1): a cell the new rectangle or page cells cover is "this pivot's" when it lies in the pivot's SAVED
 *   area (the overlay replaces those saved cells) or its current area; any other cell for which `occupied` is true
 *   refuses the change with `pivot-overlap`. `occupied` answers for the sheet as loaded (non-empty cells; Task 29 adds
 *   formatting-only cells if Task 24 measures that Excel counts them).
 * - `vacated` is the current area (rectangle and page cells) minus the new one, as disjoint row-band rectangles.
 * - Revision 0 and `reset()` carry the saved definition object itself and the saved state's layout, its implicit
 *   captions filled with the saved text; `reset()` is pure like `apply` (the caller commits it, as Task 27 stages it).
 * - `commit` accepts only a result this session issued, for exactly the next revision.
 * - Diagnostics of the faithfulness check: a mismatch is `pivot-faithful-mismatch` (first differing cell); an engine
 *   refusal inside the check keeps the engine's feature; any other unverified state is `pivot-faithful-unverified`.
 */
import type { ChartModel } from '../../../charts/model'
import type { SheetRect } from '../../tableStyles/model'
import type { PivotCacheDefinition, PivotDiagnostic, PivotTableDefinition } from '../model'
import type { EngineResult } from '../engine/axes'
import { computePivot } from '../engine/compute'
import { checkSavedLayout, engineEligibility, type CellReader, type FaithfulResult } from '../engine/faithful'
import { liveDefinition, type PivotLayout } from '../engine/place'
import type { RecordStore } from '../engine/records'
import { initialViewState, type PivotViewState } from '../engine/view'
import { captionRule, resolveCaptions, savedCaptions } from './captions'
import { applyDelta, type PivotDelta, type PivotDeltaKind } from './delta'

export interface PivotSessionInput {
  def: PivotTableDefinition; cache: PivotCacheDefinition; store: RecordStore; date1904: boolean; saved: CellReader
  occupied: (row: number, col: number) => boolean
  linkedCharts: readonly ChartModel[]   // charts whose c:pivotSource names this pivot (the client collects them)
  chartGate?: ChartGate                 // Task 30; absent → a pivot with linked charts refuses pivot-chart-mapping at open
  /** The store's lease (runtime.ts), released with the session. */
  lease?: { release(): void }
  /** The interaction kinds apply accepts; default `ORACLE_VERIFIED_KINDS`. */
  verifiedKinds?: ReadonlySet<PivotDeltaKind>
}
export type ChartGate = (charts: readonly ChartModel[], layout: PivotLayout, def: PivotTableDefinition) => EngineResult<readonly ChartModel[]>
export interface PivotApplyResult { revision: number; view: PivotViewState; def: PivotTableDefinition; layout: PivotLayout; vacated: readonly SheetRect[]; charts: readonly ChartModel[] }
export interface PivotSession {
  readonly revision: number; readonly view: PivotViewState; readonly def: PivotTableDefinition; readonly layout: PivotLayout
  apply(delta: PivotDelta): EngineResult<PivotApplyResult>   // pure: does not change the session
  commit(result: PivotApplyResult): boolean                   // false when result.revision !== revision + 1
  reset(): PivotApplyResult                                    // back to the saved state, as the next revision
  /** Ends the session and releases its store lease (once). Later applies refuse `pivot-session-released`. */
  release(): void
}

/**
 * The interaction kinds whose Excel oracle passed (spec §9: a kind with no passing oracle case is absent from the UI
 * and refused here). Each is pinned by tests/xlsx-pivots-interaction-oracle.test.ts, which asserts that every kind
 * listed has at least one passing case: today the Stage 2 cases that are Excel applying one UI change to the engine
 * seed (collapse: layout-collapsed-item; items: quirk-hidden-row-item, quirk-cross-hidden-item; sort: sort-text-*,
 * sort-number-descending, sort-value-*; filter: filter-label-*, filter-value-*, filter-top-*). `collapseField`,
 * `collapseAll`, `clearFilters` and `page` wait for Task 24's cases.
 */
export const ORACLE_VERIFIED_KINDS: ReadonlySet<PivotDeltaKind> = new Set<PivotDeltaKind>(['collapse', 'items', 'sort', 'filter'])

/** Excel's message for a change that would overlap existing data (§7.1; spec text, until Task 24 measures it). */
export const PIVOT_OVERLAP_MESSAGE = 'A PivotTable report cannot overlap another PivotTable report.'

const refuse = (feature: string, message: string): { ok: false; diagnostic: PivotDiagnostic } => ({ ok: false, diagnostic: { kind: 'unsupported', feature, message } })
const cellName = (row: number, col: number) => `R${row + 1}C${col + 1}`

function faithfulRefusal(def: PivotTableDefinition, result: Exclude<FaithfulResult, { status: 'ok' }>): { ok: false; diagnostic: PivotDiagnostic } {
  const name = `Pivot table ${def.name}`
  if (result.status === 'mismatch') {
    const { row, col, reason } = result.firstDiff
    return refuse('pivot-faithful-mismatch', `${name} differs from its saved cells at ${cellName(row, col)}: ${reason}`)
  }
  if (result.reason === 'unsupported' && result.detail?.startsWith('pivot-')) return refuse(result.detail, `${name} cannot be recomputed (${result.detail})`)
  return refuse('pivot-faithful-unverified', `${name} cannot be checked against its saved cells (${result.reason}${result.detail ? `: ${result.detail}` : ''})`)
}

const inRect = (rect: SheetRect, row: number, col: number) => row >= rect.minRow && row <= rect.maxRow && col >= rect.minCol && col <= rect.maxCol
/** The cells a layout covers: its rectangle, then its page cells (the only cells outside it) as row runs. */
function areaOf(layout: PivotLayout): SheetRect[] {
  const out: SheetRect[] = [layout.rect]
  for (const cell of layout.cells) {
    if (inRect(layout.rect, cell.row, cell.col)) continue
    const last = out[out.length - 1]!
    if (out.length > 1 && last.minRow === cell.row && last.maxCol === cell.col - 1) last.maxCol = cell.col
    else out.push({ minRow: cell.row, minCol: cell.col, maxRow: cell.row, maxCol: cell.col })
  }
  return out
}
/** The column runs of row `row` in [c0, c1] that no rect of `cover` contains. */
function uncovered(row: number, c0: number, c1: number, cover: readonly SheetRect[]): Array<[number, number]> {
  const spans = cover.filter(rect => row >= rect.minRow && row <= rect.maxRow && rect.maxCol >= c0 && rect.minCol <= c1)
    .map(rect => [Math.max(rect.minCol, c0), Math.min(rect.maxCol, c1)] as [number, number]).sort((a, b) => a[0] - b[0])
  const out: Array<[number, number]> = []
  let at = c0
  for (const [from, to] of spans) {
    if (from > at) out.push([at, from - 1])
    at = Math.max(at, to + 1)
  }
  if (at <= c1) out.push([at, c1])
  return out
}
/** `area` minus `cover`, as row-band rectangles (consecutive rows with the same runs merge). */
function subtract(area: readonly SheetRect[], cover: readonly SheetRect[]): SheetRect[] {
  const out: SheetRect[] = []
  for (const rect of area) {
    let open: SheetRect[] = []
    for (let row = rect.minRow; row <= rect.maxRow; row++) {
      const runs = uncovered(row, rect.minCol, rect.maxCol, cover)
      const same = open.length === runs.length && open.every((band, k) => band.minCol === runs[k]![0] && band.maxCol === runs[k]![1])
      if (same) { for (const band of open) band.maxRow = row; continue }
      out.push(...open)
      open = runs.map(([minCol, maxCol]) => ({ minRow: row, minCol, maxRow: row, maxCol }))
    }
    out.push(...open)
  }
  return out
}

export function openPivotSession(input: PivotSessionInput): EngineResult<PivotSession> {
  const { def, cache, store, date1904, saved, occupied, linkedCharts, chartGate, lease } = input
  const verified = input.verifiedKinds ?? ORACLE_VERIFIED_KINDS
  const eligible = engineEligibility(cache, def, { sourceReadable: false })
  if (!eligible.ok) return eligible
  const check = checkSavedLayout(store, cache, def, saved, { date1904 })
  if (check.status !== 'ok') return faithfulRefusal(def, check)
  const view0 = initialViewState(def, cache)
  const computed = computePivot(store, cache, def, view0, { date1904 })
  if (!computed.ok) return computed
  const rule = captionRule(def, saved, computed.value)
  const layout0 = savedCaptions(computed.value, saved)
  const gate = (layout: PivotLayout, live: PivotTableDefinition): EngineResult<readonly ChartModel[]> => {
    if (!linkedCharts.length) return { ok: true, value: [] }
    if (!chartGate) return refuse('pivot-chart-mapping', `Pivot table ${def.name} has ${linkedCharts.length} linked chart(s); pivot charts cannot follow its changes yet`)
    return chartGate(linkedCharts, layout, live)
  }
  const charts0 = gate(layout0, def)
  if (!charts0.ok) return charts0
  const savedArea = areaOf(layout0)

  const issued = new WeakSet<PivotApplyResult>()
  const areas = new WeakMap<PivotLayout, SheetRect[]>([[layout0, savedArea]])
  const areaFor = (layout: PivotLayout) => { let area = areas.get(layout); if (!area) areas.set(layout, area = areaOf(layout)); return area }
  let state: PivotApplyResult = { revision: 0, view: view0, def, layout: layout0, vacated: [], charts: charts0.value }
  let released = false

  const session: PivotSession = {
    get revision() { return state.revision },
    get view() { return state.view },
    get def() { return state.def },
    get layout() { return state.layout },
    apply(delta) {
      if (released) return refuse('pivot-session-released', `the session of pivot table ${def.name} was released`)
      if (!verified.has(delta.kind)) return refuse('pivot-interaction-unverified', `the ${delta.kind} interaction has no passing Excel oracle case`)
      const view = applyDelta(state.view, delta, def, cache)
      if (!view.ok) return view
      const next = computePivot(store, cache, def, view.value, { date1904 })
      if (!next.ok) return next
      const layout = resolveCaptions(next.value, rule)
      if (!layout.ok) return layout
      const area = areaFor(layout.value), current = areaFor(state.layout)
      const own = [...savedArea, ...current]
      for (const rect of area) for (let row = rect.minRow; row <= rect.maxRow; row++) {
        for (const [c0, c1] of uncovered(row, rect.minCol, rect.maxCol, own)) for (let col = c0; col <= c1; col++) {
          if (occupied(row, col)) return refuse('pivot-overlap', `${PIVOT_OVERLAP_MESSAGE} (${cellName(row, col)})`)
        }
      }
      const live = liveDefinition(def, view.value, next.value.axes, layout.value)
      const charts = gate(layout.value, live)
      if (!charts.ok) return charts
      const result: PivotApplyResult = { revision: state.revision + 1, view: view.value, def: live, layout: layout.value, vacated: subtract(current, area), charts: charts.value }
      issued.add(result)
      return { ok: true, value: result }
    },
    commit(result) {
      if (released || !issued.has(result) || result.revision !== state.revision + 1) return false
      state = result
      return true
    },
    reset() {
      const result: PivotApplyResult = { revision: state.revision + 1, view: view0, def, layout: layout0, vacated: subtract(areaFor(state.layout), savedArea), charts: charts0.value }
      issued.add(result)
      return result
    },
    release() {
      if (released) return
      released = true
      lease?.release()
    },
  }
  return { ok: true, value: session }
}
