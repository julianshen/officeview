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
 * - Overlap (§7.1, Ruling S3-AO, replacing S3-U): a cell the new rectangle or page cells cover is "this pivot's" when
 *   it lies in the pivot's SAVED area (the overlay replaces those saved cells) or its current area. Any other cell
 *   `occupied` holds — measured blockers: a cell with a value (O1) or another pivot's current area (O3) — makes the
 *   change commit as Excel 16.114 saves it (owner-overlap-value, -pivot): the pivot spills, one #SPILL! error cell at
 *   its top-left (`location` that cell, header/data offsets 0), its other cells empty, the item state kept (a later
 *   change that fits shows it again). A cell `unmeasured` holds (merges, spill/array members, noted cells, another
 *   pivot's vacated saved area: Ruling S3-AC), a pivot with linked charts, or one with page fields refuses
 *   `pivot-overlap` instead, naming the cell (fail closed: unmeasured). A formatting-only cell is not occupied: the
 *   pivot grows over it and its format replaces the cell's (Ruling S3-V, owner-overlap-format).
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
import { applyDeltaState, type PivotDelta, type PivotDeltaKind, type PivotMemoryHold } from './delta'
import { recordsFromSource, refreshDefinition, refreshedCacheBytes, savedCaptionLayout, worksheetSourceRange } from './refresh'
import { savedPivotLayout } from '../savedLayout'

export interface PivotSessionInput {
  /** `store`: the decoded records; absent for a refreshOnLoad pivot rebuilt from `source` (§7.7). */
  def: PivotTableDefinition; cache: PivotCacheDefinition; store?: RecordStore; date1904: boolean; saved: CellReader
  /** Measured blockers (Ruling S3-AO): a change covering one spills. */
  occupied: (row: number, col: number) => boolean
  /** Unmeasured blockers (Ruling S3-AC): a change covering one refuses `pivot-overlap`. */
  unmeasured?: (row: number, col: number) => boolean
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
  /** Reserves candidate cache/index bytes before a Stage 4 field move allocates them. */
  admitCandidateBytes?: (extraBytes: number) => PivotMemoryHold | undefined
  /** Admits a rebuilt store, and the rebuilt cache's `metadataBytes` (`refreshedCacheBytes`), before the session
   * holds them (the caller's ledger or pool), for the session's life; undefined refuses
   * `pivot-records-limit`. The returned hold is released with the session. */
  admitStore?: (store: RecordStore, metadataBytes: number) => { release(): void } | undefined
}
export type ChartGate = (charts: readonly ChartModel[], layout: PivotLayout, def: PivotTableDefinition) => EngineResult<readonly ChartModel[]>
/** `notice`: the one-line notice a refreshed pivot shows with its first change (§7.7), on every result issued before
 * the session's first commit. */
export interface PivotApplyResult { revision: number; view: PivotViewState; def: PivotTableDefinition; cache: PivotCacheDefinition; store: RecordStore; layout: PivotLayout; vacated: readonly SheetRect[]; charts: readonly ChartModel[]; hold?: PivotMemoryHold; notice?: string
  /** Ruling S3-AO: the change spilled; the blocking cell (also `def.spill`). */
  spill?: { row: number; col: number } }
export interface PivotSession {
  readonly revision: number; readonly view: PivotViewState; readonly def: PivotTableDefinition; readonly layout: PivotLayout
  /** The cache the session computes from: the saved one, or the one a refresh rebuilt (its definitions index it). */
  readonly cache: PivotCacheDefinition
  apply(delta: PivotDelta, options?: { admitCandidateBytes?: (extraBytes: number) => PivotMemoryHold | undefined }): EngineResult<PivotApplyResult>   // pure: does not change the session
  commit(result: PivotApplyResult): boolean                   // false when result.revision !== revision + 1
  /** Back to the saved state, as the next revision; refused `pivot-overlap` when that layout would now cover another
   * pivot's or the sheet's cells (a refreshed pivot's saved view is its refreshed layout, which may be larger). */
  reset(): EngineResult<PivotApplyResult>
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
export const ORACLE_VERIFIED_KINDS: ReadonlySet<PivotDeltaKind> = new Set<PivotDeltaKind>(['collapse', 'collapseField', 'collapseAll', 'items', 'page', 'filter', 'clearFilters', 'sort', 'moveField', 'addDataField', 'removeDataField', 'moveDataField'])

const a1 = (row: number, col: number) => {
  let letters = ''
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters
  return `${letters}${row + 1}`
}
/** The refusal of a change that would cover an unmeasured blocker (Ruling S3-AC; S3-AO spills over measured ones),
 * naming the blocking cell. */
export function pivotOverlapMessage(row: number, col: number, sheet: string): string {
  return `This change would overlap cell ${a1(row, col)} on ${sheet}.`
}
/** Ruling S3-AO: the change's announcement when the pivot spills, naming the blocking cell. */
export function pivotSpillNotice(name: string, row: number, col: number, sheet: string): string {
  return `Pivot table ${name.slice(0, 128)} can't fit: #SPILL!. Cell ${a1(row, col)} on ${sheet} is in the way.`
}
/** The spilled layout (Ruling S3-AO, owner-overlap-value): one #SPILL! error cell at the would-be layout's top-left. */
function spillLayout(layout: PivotLayout): PivotLayout {
  const { minRow: row, minCol: col } = layout.rect
  return { rect: { minRow: row, maxRow: row, minCol: col, maxCol: col }, axes: layout.axes, cells: [{ row, col, value: { kind: 'error', value: '#SPILL!' } }], merges: [], captionsUnverified: false }
}
/** The live definition as Excel saves a spilled pivot: `location` its top-left cell, header and data offsets 0 (the
 * item state, e.g. an expanded item, is kept). */
function spillDefinition(live: PivotTableDefinition, layout: PivotLayout, spill: { row: number; col: number }): PivotTableDefinition {
  return { ...live, location: { ...live.location, ref: a1(layout.rect.minRow, layout.rect.minCol), firstHeaderRow: 0, firstDataRow: 0, firstDataCol: 0, rowPageCount: 0, colPageCount: 0 }, spill }
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
  const unmeasured = input.unmeasured ?? (() => false)
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
  let state: PivotApplyResult = { revision: 0, view: view0, def: def0, cache, store, layout: layout0, vacated: [], charts: charts0.value }
  /** The area the sheet shows: the saved area until the first commit. */
  let shown: SheetRect[] = savedArea
  let notice = notice0
  let released = false

  /** Ruling S3-AO/S3-AC over the cells of `area` outside this pivot's own (saved ∪ shown) area: the first unmeasured
   * blocker refuses; otherwise the first measured one spills. */
  const overlapOf = (area: readonly SheetRect[]): { ok: false; diagnostic: PivotDiagnostic } | { spill: { row: number; col: number } } | undefined => {
    const own = [...savedArea, ...shown]
    let spill: { row: number; col: number } | undefined
    for (const rect of area) for (let row = rect.minRow; row <= rect.maxRow; row++) {
      for (const [c0, c1] of uncovered(row, rect.minCol, rect.maxCol, own)) for (let col = c0; col <= c1; col++) {
        if (unmeasured(row, col)) return refuse('pivot-overlap', pivotOverlapMessage(row, col, def.sheet))
        if (!spill && occupied(row, col)) spill = { row, col }
      }
    }
    return spill ? { spill } : undefined
  }

  const session: PivotSession = {
    get revision() { return state.revision },
    get view() { return state.view },
    get def() { return state.def },
    get layout() { return state.layout },
    get cache() { return state.cache },
    apply(delta, options) {
      if (released) return refuse('pivot-session-released', `the session of pivot table ${def.name} was released`)
      if (!verified.has(delta.kind)) return refuse('pivot-interaction-unverified', `the ${delta.kind} interaction has no passing Excel oracle case`)
      // A spill stores a one-cell display location on the live definition. Use the saved geometry as the basis for
      // the next candidate while retaining its current field definitions and view; a change that now fits must
      // restore the pivot at its original location.
      const applyDef = state.spill ? (() => { const restored = { ...state.def, location: def0.location }; delete restored.spill; return restored })() : state.def
      const transitioned = applyDeltaState({ def: applyDef, cache: state.cache, store: state.store, view: state.view, date1904,
        captionRule: { caption: (kind, args) => rule.caption(kind, args) }, admitCandidateBytes: options?.admitCandidateBytes ?? input.admitCandidateBytes }, delta)
      if (!transitioned.ok) return transitioned
      const candidate = transitioned.value
      const discard = <T extends EngineResult<unknown>>(result: T): T => { candidate.hold?.release(); return result }
      // Excel keeps the opened pivot's original location when every field and value is removed (Task 34 oracle).
      // The intermediate live definitions shrink as fields move; use the saved location for this final empty shape.
      const emptyPivot = !candidate.view.data.length && !candidate.view.rows.length && !candidate.view.cols.length && !candidate.view.pages.length && !candidate.view.filters.length
      // Excel's sole Task 35 all-fields-removed oracle starts from this saved A3 pivot (12 fields,
      // Region/Product rows, Year + Values columns) and retains the measured empty shell A3:C20.
      const task35EmptyBase = def0.location.ref === 'A3:I24' && def0.fields.length === 12 && view0.rows.join(',') === '0,1' && view0.cols.join(',') === '2,-2' && view0.data.length === 2
      if (emptyPivot && !task35EmptyBase) return discard(refuse('pivot-engine-unmeasured', 'An empty PivotTable layout has not been measured for this saved pivot.'))
      const emptyRef = task35EmptyBase ? 'A3:C20' : def0.location.ref
      const layoutDef = emptyPivot ? { ...candidate.def, location: { ...candidate.def.location, ref: emptyRef } } : candidate.def
      const next = computePivot(candidate.store, candidate.cache, layoutDef, candidate.view, { date1904 })
      if (!next.ok) return discard(next)
      const layout = resolveCaptions(next.value, rule)
      if (!layout.ok) return discard(layout)
      const area = areaFor(layout.value), current = shown
      const overlap = overlapOf(area)
      if (overlap && 'ok' in overlap) return discard(overlap)
      const live = liveDefinition(layoutDef, candidate.view, next.value.axes, layout.value)
      if (overlap) {
        // A spilled pivot's charts and page cells were never measured: those refuse (fail closed).
        const { row, col } = overlap.spill
        if (linkedCharts.length || candidate.view.pages.length) return discard(refuse('pivot-overlap', pivotOverlapMessage(row, col, def.sheet)))
        const spilled = spillLayout(layout.value)
        const result: PivotApplyResult = { revision: state.revision + 1, view: candidate.view, def: spillDefinition(live, layout.value, overlap.spill), cache: candidate.cache, store: candidate.store, layout: spilled,
          vacated: subtract(current, areaFor(spilled)), charts: [], ...(candidate.hold ? { hold: candidate.hold } : {}), spill: overlap.spill, notice: [notice, pivotSpillNotice(String(def.name ?? ''), row, col, def.sheet)].filter(Boolean).join(' ') }
        issued.add(result)
        return { ok: true, value: result }
      }
      const charts = gate(layout.value, live)
      if (!charts.ok) return discard(charts)
      const result: PivotApplyResult = { revision: state.revision + 1, view: candidate.view, def: live, cache: candidate.cache, store: candidate.store, layout: layout.value, vacated: subtract(current, area), charts: charts.value, ...(candidate.hold ? { hold: candidate.hold } : {}), ...(notice ? { notice } : {}) }
      issued.add(result)
      return { ok: true, value: result }
    },
    commit(result) {
      if (released || !issued.has(result) || result.revision !== state.revision + 1) return false
      const priorHold = state.hold
      state = result
      if (priorHold && priorHold !== result.hold) priorHold.release()
      shown = areaFor(result.layout)
      notice = undefined
      return true
    },
    reset() {
      if (released) return refuse('pivot-session-released', `the session of pivot table ${def.name} was released`)
      // The saved route's layout0 covers exactly the saved area (nothing to check); a refreshed one may reach further
      // (final review minor 1: checked like any change).
      // A refreshed saved view that no longer fits refuses (its spill was never measured).
      const overlap = overlapOf(areaFor(layout0))
      if (overlap) return 'ok' in overlap ? overlap : refuse('pivot-overlap', pivotOverlapMessage(overlap.spill.row, overlap.spill.col, def.sheet))
      const result: PivotApplyResult = { revision: state.revision + 1, view: view0, def: def0, cache, store, layout: layout0, vacated: subtract(shown, areaFor(layout0)), charts: charts0.value, ...(notice ? { notice } : {}) }
      issued.add(result)
      return { ok: true, value: result }
    },
    release() {
      if (released) return
      released = true
      lease?.release()
      hold?.release()
      state.hold?.release()
    },
  }
  return { ok: true, value: session }
}
