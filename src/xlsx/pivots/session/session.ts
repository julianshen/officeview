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
 *   refuses the change with `pivot-overlap`. `occupied` answers for the sheet as loaded: a cell with a value, or a
 *   cell of another pivot. A formatting-only cell is not occupied: the pivot grows over it and its format replaces
 *   the cell's (Ruling S3-V, owner-overlap-format). Excel 16.114 turns an overlapping pivot into a #SPILL! anchor
 *   with no alert (owner-overlap-value, -pivot); the viewer keeps §7.1's refusal, the pivot unchanged, with its own
 *   message naming the blocking cell (Ruling S3-U, `pivotOverlapMessage`).
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
import { RECORD_LIMITS, type RecordLimits, type RecordStore } from '../engine/records'
import { initialViewState, type PivotViewState } from '../engine/view'
import { captionRule, MEASURED_CAPTION_TABLES, resolveCaptions, savedCaptions, type CaptionTable } from './captions'
import { applyDelta, type PivotDelta, type PivotDeltaKind } from './delta'
import { recordsFromSource, refreshDefinition, refreshedCacheBytes, savedCaptionLayout, worksheetSourceRange } from './refresh'
import { savedPivotLayout } from '../savedLayout'

export interface PivotSessionInput {
  /** `store`: the decoded records; absent for a refreshOnLoad pivot rebuilt from `source` (§7.7). */
  def: PivotTableDefinition; cache: PivotCacheDefinition; store?: RecordStore; date1904: boolean; saved: CellReader
  occupied: (row: number, col: number) => boolean
  linkedCharts: readonly ChartModel[]   // charts whose c:pivotSource names this pivot (the client collects them)
  chartGate?: ChartGate                 // Task 30; absent → a pivot with linked charts refuses pivot-chart-mapping at open
  /** The store's lease (runtime.ts), released with the session. */
  lease?: { release(): void }
  /** The interaction kinds apply accepts; default `ORACLE_VERIFIED_KINDS`. */
  verifiedKinds?: ReadonlySet<PivotDeltaKind>
  /** The measured caption tables the locale is detected among; default `MEASURED_CAPTION_TABLES`. */
  captionTables?: Readonly<Partial<Record<'zh-TW', CaptionTable>>>
  /** Task 32: the cells of the cache's source sheet. Given, a refreshOnLoad pivot with a worksheet-range source is
   * rebuilt from it (refresh.ts) instead of its saved records; absent, such a pivot refuses `pivot-refresh-on-load`. */
  source?: CellReader
  /** The rebuild's record limits (default `RECORD_LIMITS.D`). */
  limits?: RecordLimits
  /** Admits a rebuilt store, and the rebuilt cache's `metadataBytes` (`refreshedCacheBytes`), before the session
   * holds them (the caller's ledger or pool), for the session's life; undefined refuses
   * `pivot-records-limit`. The returned hold is released with the session. */
  admitStore?: (store: RecordStore, metadataBytes: number) => { release(): void } | undefined
}
export type ChartGate = (charts: readonly ChartModel[], layout: PivotLayout, def: PivotTableDefinition) => EngineResult<readonly ChartModel[]>
/** `notice`: the one-line notice a refreshed pivot shows with its first change (§7.7), on every result issued before
 * the session's first commit. */
export interface PivotApplyResult { revision: number; view: PivotViewState; def: PivotTableDefinition; layout: PivotLayout; vacated: readonly SheetRect[]; charts: readonly ChartModel[]; notice?: string }
export interface PivotSession {
  readonly revision: number; readonly view: PivotViewState; readonly def: PivotTableDefinition; readonly layout: PivotLayout
  /** The cache the session computes from: the saved one, or the one a refresh rebuilt (its definitions index it). */
  readonly cache: PivotCacheDefinition
  apply(delta: PivotDelta): EngineResult<PivotApplyResult>   // pure: does not change the session
  commit(result: PivotApplyResult): boolean                   // false when result.revision !== revision + 1
  reset(): PivotApplyResult                                    // back to the saved state, as the next revision
  /** Ends the session and releases its store lease (once). Later applies refuse `pivot-session-released`. */
  release(): void
}

/**
 * The interaction kinds whose Excel oracle passed (spec §9: a kind with no passing oracle case is absent from the UI
 * and refused here). Each is pinned by tests/xlsx-pivots-interaction-oracle.test.ts, which asserts that every kind
 * listed has at least one passing case. As of Task 24's oracle (44d9dbe8) every built task-25 delta case passes:
 * - collapse: interaction-collapse-item-* (compact, outline, tabular; subtotals top and bottom), interaction-expand-item,
 *   layout-collapsed-item, interaction-sort-then-collapse;
 * - collapseField / collapseAll: interaction-collapse-field-*, interaction-collapse-all-*, interaction-expand-field/-all;
 * - items: quirk-hidden-row-item, quirk-cross-hidden-item, interaction-filter-all-hidden, interaction-items-*;
 * - page: interaction-page-select-other, -all, -multi, -multi-to-all;
 * - filter (label, value, top-10 and date): the Stage 2 filter-* cases, interaction-filter-date-between;
 * - clearFilters: interaction-clear-label/value/top10-filter, -clear-filters-hidden-items, -filter-date-clear;
 * - sort: the Stage 2 sort-* cases, interaction-sort-value-ascending-ties.
 * A shape the engine or the delta refuses stays refused inside a listed kind (a second filter family on a field, a
 * dynamic date filter, a nested ascending value-sort tie, …).
 */
export const ORACLE_VERIFIED_KINDS: ReadonlySet<PivotDeltaKind> = new Set<PivotDeltaKind>(['collapse', 'collapseField', 'collapseAll', 'items', 'page', 'filter', 'clearFilters', 'sort'])

/** The refusal of a change that would overlap a cell outside the pivot (Ruling S3-U: Excel shows no alert, it turns
 * the pivot into a #SPILL! anchor; the viewer refuses and names the blocking cell). */
export function pivotOverlapMessage(row: number, col: number, sheet: string): string {
  let letters = ''
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters
  return `This change would overlap cell ${letters}${row + 1} on ${sheet}.`
}

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

/** The one-line notice of a refreshed pivot (§7.7). */
export const PIVOT_REFRESH_NOTICE = (name: string) => `Pivot table ${name} was refreshed from the workbook's data, as Excel does when it opens this file.`

interface Opened { def: PivotTableDefinition; cache: PivotCacheDefinition; store: RecordStore; view0: PivotViewState; layout0: PivotLayout; def0: PivotTableDefinition; rule: ReturnType<typeof captionRule>; displayed?: SheetRect[]; hold?: { release(): void } }

/** The saved route: the faithfulness check on the decoded records, then the saved state with its saved captions. */
function openSaved(input: PivotSessionInput, store: RecordStore): EngineResult<Opened> {
  const { def, cache, date1904, saved } = input
  const check = checkSavedLayout(store, cache, def, saved, { date1904 })
  if (check.status !== 'ok') return faithfulRefusal(def, check)
  const view0 = initialViewState(def, cache)
  const computed = computePivot(store, cache, def, view0, { date1904 })
  if (!computed.ok) return computed
  const rule = captionRule(def, saved, computed.value, input.captionTables ?? MEASURED_CAPTION_TABLES)
  return { ok: true, value: { def, cache, store, view0, layout0: savedCaptions(computed.value, saved), def0: def, rule } }
}

/**
 * The §7.7 route (Task 32): rebuild the cache and records from the source range, re-point the definition's items,
 * and lay out the refreshed saved view. The saved cells predate the source, so there is no faithfulness check; the
 * caption rule reads them through the saved lines' own layout, and the refreshed view is shown only from the first
 * committed change (until then the saved area is what the sheet shows).
 */
function openRefreshed(input: PivotSessionInput, source: CellReader): EngineResult<Opened> {
  const { def, cache, date1904, saved } = input
  const name = `Pivot table ${def.name}`
  const range = worksheetSourceRange(cache)
  if (!range) return refuse('pivot-refresh-on-load', `${name} refreshes on load and its source is not a worksheet range`)
  const rebuilt = recordsFromSource(cache, source, range.ref, date1904, input.limits ?? RECORD_LIMITS.D)
  if (!rebuilt.ok) return rebuilt
  const live = refreshDefinition(def, cache, rebuilt.value.cache)
  if (!live.ok) return live
  const savedLines = savedCaptionLayout(def, cache, date1904)
  if (!savedLines) return refuse('pivot-refresh-on-load', `${name}'s saved layout cannot be read for its captions`)
  const rule = captionRule(def, saved, savedLines, input.captionTables ?? MEASURED_CAPTION_TABLES)
  const { cache: refreshed, store } = rebuilt.value
  const view0 = initialViewState(live.value, refreshed)
  const computed = computePivot(store, refreshed, live.value, view0, { date1904 })
  if (!computed.ok) return computed
  const layout0 = resolveCaptions(computed.value, rule)
  if (!layout0.ok) return layout0
  let hold: { release(): void } | undefined
  if (input.admitStore) {
    hold = input.admitStore(store, refreshedCacheBytes(refreshed))
    if (!hold) return refuse('pivot-records-limit', `${name}'s refreshed records (${store.bytes} bytes) could not be admitted`)
  }
  const def0 = liveDefinition(live.value, view0, computed.value.axes, layout0.value)
  return { ok: true, value: { def: live.value, cache: refreshed, store, view0, layout0: layout0.value, def0, rule, displayed: savedPivotLayout(def)?.rects ?? [], ...(hold ? { hold } : {}) } }
}

export function openPivotSession(input: PivotSessionInput): EngineResult<PivotSession> {
  const { date1904, occupied, linkedCharts, chartGate, lease } = input
  const verified = input.verifiedKinds ?? ORACLE_VERIFIED_KINDS
  const readable = input.source !== undefined && worksheetSourceRange(input.cache) !== undefined
  const eligible = engineEligibility(input.cache, input.def, { sourceReadable: readable })
  if (!eligible.ok) return eligible
  if (!eligible.refresh && !input.store) return refuse('pivot-records-absent', `Pivot table ${input.def.name} has no decoded records`)
  const route = eligible.refresh ? openRefreshed(input, input.source!) : openSaved(input, input.store!)
  if (!route.ok) return route
  const { def, cache, store, view0, layout0, def0, rule, hold } = route.value
  const notice0 = eligible.refresh ? PIVOT_REFRESH_NOTICE(String(input.def.name)) : undefined
  const gate = (layout: PivotLayout, live: PivotTableDefinition): EngineResult<readonly ChartModel[]> => {
    if (!linkedCharts.length) return { ok: true, value: [] }
    if (!chartGate) return refuse('pivot-chart-mapping', `Pivot table ${def.name} has ${linkedCharts.length} linked chart(s); pivot charts cannot follow its changes yet`)
    return chartGate(linkedCharts, layout, live)
  }
  const charts0 = gate(layout0, def0)
  if (!charts0.ok) { hold?.release(); return charts0 }
  // What the sheet shows before the first commit: the saved area (a refreshed pivot's saved cells are stale).
  const savedArea = route.value.displayed ?? areaOf(layout0)

  const issued = new WeakSet<PivotApplyResult>()
  const areas = new WeakMap<PivotLayout, SheetRect[]>()
  const areaFor = (layout: PivotLayout) => { let area = areas.get(layout); if (!area) areas.set(layout, area = areaOf(layout)); return area }
  let state: PivotApplyResult = { revision: 0, view: view0, def: def0, layout: layout0, vacated: [], charts: charts0.value }
  /** The area the sheet shows: the saved area until the first commit. */
  let shown: SheetRect[] = savedArea
  let notice = notice0
  let released = false

  const session: PivotSession = {
    get revision() { return state.revision },
    get view() { return state.view },
    get def() { return state.def },
    get layout() { return state.layout },
    get cache() { return cache },
    apply(delta) {
      if (released) return refuse('pivot-session-released', `the session of pivot table ${def.name} was released`)
      if (!verified.has(delta.kind)) return refuse('pivot-interaction-unverified', `the ${delta.kind} interaction has no passing Excel oracle case`)
      const view = applyDelta(state.view, delta, def, cache)
      if (!view.ok) return view
      const next = computePivot(store, cache, def, view.value, { date1904 })
      if (!next.ok) return next
      const layout = resolveCaptions(next.value, rule)
      if (!layout.ok) return layout
      const area = areaFor(layout.value), current = shown
      const own = [...savedArea, ...current]
      for (const rect of area) for (let row = rect.minRow; row <= rect.maxRow; row++) {
        for (const [c0, c1] of uncovered(row, rect.minCol, rect.maxCol, own)) for (let col = c0; col <= c1; col++) {
          if (occupied(row, col)) return refuse('pivot-overlap', pivotOverlapMessage(row, col, def.sheet))
        }
      }
      const live = liveDefinition(def, view.value, next.value.axes, layout.value)
      const charts = gate(layout.value, live)
      if (!charts.ok) return charts
      const result: PivotApplyResult = { revision: state.revision + 1, view: view.value, def: live, layout: layout.value, vacated: subtract(current, area), charts: charts.value, ...(notice ? { notice } : {}) }
      issued.add(result)
      return { ok: true, value: result }
    },
    commit(result) {
      if (released || !issued.has(result) || result.revision !== state.revision + 1) return false
      state = result
      shown = areaFor(result.layout)
      notice = undefined
      return true
    },
    reset() {
      const result: PivotApplyResult = { revision: state.revision + 1, view: view0, def: def0, layout: layout0, vacated: subtract(shown, areaFor(layout0)), charts: charts0.value, ...(notice ? { notice } : {}) }
      issued.add(result)
      return result
    },
    release() {
      if (released) return
      released = true
      lease?.release()
      hold?.release()
    },
  }
  return { ok: true, value: session }
}
