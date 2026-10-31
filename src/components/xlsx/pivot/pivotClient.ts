/**
 * The pivot client (plan Task 28, spec §6.6–6.8, §7.1, §7.6, §7.9): what the interactive viewer calls when a pivot's
 * button is used. It opens Task 25's session on first use, applies a delta or a reset, commits it, and hands the
 * viewer a commit to install. Lazy: loaded with the pivot controls (this module statically imports the `pivots`
 * chunk; nothing on the eager path imports it).
 *
 * Two implementations behind one interface:
 * - **resident** (`createResidentPivotClient`): the session runs on the main thread over the runtime the main-thread
 *   parse retained (`pivotRuntimeFor(doc)`), its store leased against the viewer's main ledger;
 * - **remote** (`createRemotePivotClient`, the Task 27 seam): a typed adapter over the worker protocol's
 *   `RemotePivotSessions`. Until that port exists, a paged or worker-parsed workbook gets `unavailablePivotClient`:
 *   its buttons are painted and named "not interactive", honestly.
 *
 * `sheetId` in this interface is the sheet's position in `doc.sheets` (the plan's `number`), not the viewer's string
 * sheet id; the viewer maps between them.
 *
 * Rulings recorded here (Task 28):
 * - The overlay a commit carries is the sheet's composite over every committed pivot. Each pivot owns its saved area
 *   (`savedPivotLayout(savedDef).rects`, page cells included) and its current area: a cell of the current layout reads
 *   the layout's value, any other cell of either area reads blank (`null`). Owning the whole saved area, not only the
 *   last change's `vacated`, keeps rows an earlier change vacated blank through every later change.
 * - Overlap (`occupied`, §7.1, refined by Task 29): a cell inside another pivot's current area (or, for a pivot never
 *   opened, its saved rectangles) is occupied; a cell another pivot vacated is free; anything else is occupied when
 *   the sheet as loaded has a value there (`sheetCellReader`, as the oracle does).
 * - Linked charts: every chart on any sheet whose `c:pivotSource` names this pivot (the name after the last `!`),
 *   whatever sheet it names: over-linking only makes a pivot display-only, never stale (fail closed, §6.6).
 */
import type { XlsxDocument, XlsxSheet } from '../../../xlsx/types'
import type { ChartModel } from '../../../charts/model'
import type { ProfileId, ValueOverlay, ValueOverlayCell } from '../../../xlsx/view/contracts'
import type { SourceRect } from '../../../xlsx/view/coordinates'
import { cellDisplayText } from '../../../xlsx/render'
import { pivotRuntimeFor } from '../../../xlsx/pivotLoader'
import type { RuntimeLedger } from '../../../xlsx/pivots/session/runtime'
import {
  PIVOT_VALUES_FIELD,
  type PivotCacheDefinition, type PivotDiagnostic, type PivotFieldSortType, type PivotFilter, type PivotTableDefinition,
} from '../../../xlsx/pivots/model'
import {
  openPivotSession, ORACLE_VERIFIED_KINDS, pivotItemLabel, RECORD_LIMITS, savedPivotLayout, sheetCellReader,
  type ChartGate, type PivotDelta, type PivotDeltaKind, type PivotLayout, type PivotSession, type PivotViewState,
} from '../../../xlsx/pivots/index'
import { sheetCharts } from '../sheetCharts'

export type { PivotDelta, PivotDeltaKind }

/** One item of a field, by its pivot-item index (the index a delta names). */
export interface PivotMenuItem { index: number; label: string; hidden: boolean; collapsed: boolean }
export interface PivotMenuField {
  index: number
  name: string
  axis: 'row' | 'col' | 'page' | 'none'
  /** The field's data items (subtotal items are not listed). */
  items: PivotMenuItem[]
  sort: PivotFieldSortType
  /** The data field a value sort follows. */
  byData?: number
  filters: readonly PivotFilter[]
  /** What the label/date filter compares: the cache field's item types. */
  values: 'text' | 'number' | 'date'
  /** A level lies below this field on its axis (collapse entire field applies). */
  expandable: boolean
  /** A page field's selection. */
  page?: readonly number[] | 'all'
}
export interface PivotMenuModel {
  name: string
  fields: PivotMenuField[]
  data: Array<{ index: number; name: string }>
  /** The id a new filter takes (Excel's filter ids are unique in the pivot). */
  nextFilterId: number
}
/** What a commit installs in the viewer (spec §7.9): the live definition and, on the resident path, the sheet's
 * composite overlay (`null` once every pivot of the sheet is back to its saved view: the saved cells show again,
 * their own styles included). The paged path's overlay lives in the worker (Task 27); its commit carries none. */
export interface PivotCommit { sheetId: number; pivot: number; revision: number; def: PivotTableDefinition; overlay?: ValueOverlay | null }
export type PivotOpenResult =
  | { ok: true; revision: number; def: PivotTableDefinition; verifiedKinds: ReadonlySet<PivotDeltaKind>; model: PivotMenuModel }
  | { ok: false; diagnostic: PivotDiagnostic }
export type PivotApplyOutcome =
  | { ok: true; revision: number; def: PivotTableDefinition; model: PivotMenuModel; commit: PivotCommit }
  | { ok: false; diagnostic: PivotDiagnostic }
/** Field and item names for button labels before any session opens (no records are read). */
export interface PivotLabels { field(field: number): string; item(field: number, item: number): string | undefined }

export interface PivotClient {
  open(sheetId: number, pivot: number, signal: AbortSignal): Promise<PivotOpenResult>
  apply(sheetId: number, pivot: number, delta: PivotDelta | 'reset'): Promise<PivotApplyOutcome>
  release(sheetId: number, pivot: number): void
  /** A refusal known without opening (no runtime, no cache, a remote path without its port). */
  preflight?(sheetId: number, pivot: number): PivotDiagnostic | undefined
  labels?(sheetId: number, pivot: number): PivotLabels | undefined
  /** Releases every session (viewer unmount). */
  dispose?(): void
}

const refusal = (feature: string, message: string): { ok: false; diagnostic: PivotDiagnostic } => ({ ok: false, diagnostic: { kind: 'unsupported', feature, message } })
const nameOf = (def: PivotTableDefinition | undefined) => `Pivot table ${String(def?.name ?? '').slice(0, 128)}`

// ---- the remote seam (Task 27) ---------------------------------------------------------------------------------------

/**
 * What the paged path must supply (Task 27's `RemotePivotSessions`, plan "Interfaces"): open, apply and release over
 * the worker protocol, the apply resolving only after the worker committed the staged candidate (`pivotCommit`), so
 * the viewer installs a committed revision. The worker owns the overlay; the main side publishes merges and extent
 * through `RemoteWorkbook` and bumps its `valueGen` (plan Task 27 "On the commit reply"). The menu model (item
 * labels, the view state) is built in the worker from the cache it holds.
 */
export interface RemotePivotPort {
  open(sheetId: number, pivot: number, signal: AbortSignal): Promise<PivotOpenResult>
  apply(sheetId: number, pivot: number, delta: PivotDelta | 'reset'): Promise<PivotApplyOutcome>
  release(sheetId: number, pivot: number): void
  labels?(sheetId: number, pivot: number): PivotLabels | undefined
}
/** The paged client: the port as-is, the viewer's install run on each committed outcome. */
export function createRemotePivotClient(port: RemotePivotPort, install: (commit: PivotCommit) => void): PivotClient {
  return {
    open: (sheetId, pivot, signal) => port.open(sheetId, pivot, signal),
    async apply(sheetId, pivot, delta) {
      const outcome = await port.apply(sheetId, pivot, delta)
      if (outcome.ok) install(outcome.commit)
      return outcome
    },
    release: (sheetId, pivot) => port.release(sheetId, pivot),
    ...(port.labels ? { labels: port.labels.bind(port) } : {}),
  }
}
/** Not interactive here, with one reason (the paged path until Task 27's port lands). */
export function unavailablePivotClient(feature: string, message: string): PivotClient {
  const diagnostic: PivotDiagnostic = { kind: 'unsupported', feature, message }
  return {
    open: async () => ({ ok: false, diagnostic }),
    apply: async () => ({ ok: false, diagnostic }),
    release: () => {},
    preflight: () => diagnostic,
  }
}
export const REMOTE_PIVOTS_UNAVAILABLE = { feature: 'pivot-remote-unavailable', message: 'Pivot tables are shown as saved: changing them is not available in the large-file view yet.' } as const

// ---- the menu model --------------------------------------------------------------------------------------------------

const FUNCTION_NAMES: Record<string, string> = {
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp',
}
function fieldName(def: PivotTableDefinition, cache: PivotCacheDefinition | undefined, field: number): string {
  if (field === PIVOT_VALUES_FIELD) return def.dataCaption || 'Values'
  return def.fields[field]?.name ?? cache?.fields[field]?.name ?? `Field ${field + 1}`
}
function itemText(sheet: XlsxSheet, def: PivotTableDefinition, cache: PivotCacheDefinition, field: number, item: number, date1904: boolean): string {
  const label = pivotItemLabel(def, cache, field, item, date1904)
  if (label.value.kind === 'empty') return ''
  if (label.value.kind !== 'number') return label.value.value
  const shared = cache.fields[field]?.sharedItems
  const dates = Boolean(shared?.containsDate && !shared.containsString && !shared.containsNumber)
  // A custom format id has no format code here: dates fall back to the built-in short date, numbers to General.
  const numFmtId = label.numFmtId !== undefined && label.numFmtId < 164 ? label.numFmtId : dates ? 14 : 0
  try { return cellDisplayText(sheet, { value: label.value.value, style: { numFmtId } }) } catch { return String(label.value.value) }
}
function valuesOf(cache: PivotCacheDefinition, field: number): PivotMenuField['values'] {
  const shared = cache.fields[field]?.sharedItems
  if (!shared || shared.containsString || shared.containsMixedTypes) return 'text'
  if (shared.containsDate && !shared.containsNumber) return 'date'
  if (shared.containsNumber && !shared.containsDate) return 'number'
  return 'text'
}
export function menuModel(sheet: XlsxSheet, def: PivotTableDefinition, cache: PivotCacheDefinition, view: PivotViewState, date1904: boolean): PivotMenuModel {
  const axisOf = (field: number): PivotMenuField['axis'] => view.rows.includes(field) ? 'row' : view.cols.includes(field) ? 'col' : view.pages.some(page => page.field === field) ? 'page' : 'none'
  const below = (fields: readonly number[], field: number) => { const at = fields.indexOf(field); return at >= 0 && fields.slice(at + 1).length > 0 }
  const used = [...new Set([...view.rows, ...view.cols, ...view.pages.map(page => page.field)])].filter(field => field !== PIVOT_VALUES_FIELD)
  const fields = used.map((field): PivotMenuField => {
    const state = view.fields[field]
    const list = def.fields[field]?.items
    const indices = list ? list.flatMap((item, k) => item.type === 'data' ? [k] : []) : (state?.hidden ?? []).map((_, k) => k)
    const scope = state?.autoSortScope?.references.find(ref => ref.field === PIVOT_VALUES_FIELD)?.x[0]
    const page = view.pages.find(entry => entry.field === field)
    return {
      index: field, name: fieldName(def, cache, field), axis: axisOf(field),
      items: indices.map(k => ({ index: k, label: itemText(sheet, def, cache, field, k, date1904), hidden: state?.hidden[k] ?? false, collapsed: state?.collapsed[k] ?? false })),
      sort: state?.sortType ?? 'manual', ...(scope !== undefined ? { byData: scope } : {}),
      filters: view.filters.filter(filter => filter.fld === field), values: valuesOf(cache, field),
      expandable: below(view.rows, field) || below(view.cols, field),
      ...(page ? { page: page.selected } : {}),
    }
  })
  const data = view.data.map((entry, index) => ({ index, name: entry.name || `${FUNCTION_NAMES[entry.fn] ?? entry.fn} of ${fieldName(def, cache, entry.field)}` }))
  const nextFilterId = view.filters.reduce((max, filter) => Math.max(max, filter.id), 0) + 1
  return { name: String(def.name ?? ''), fields, data, nextFilterId }
}

// ---- the overlay -----------------------------------------------------------------------------------------------------

const toSource = (rect: { minRow: number; maxRow: number; minCol: number; maxCol: number }): SourceRect => ({ r0: rect.minRow, r1: rect.maxRow, c0: rect.minCol, c1: rect.maxCol })
const within = (rects: readonly SourceRect[], row: number, col: number) => rects.some(rect => row >= rect.r0 && row <= rect.r1 && col >= rect.c0 && col <= rect.c1)
/** The cells a layout covers: its rectangle, then one cell per page cell outside it. */
function layoutArea(layout: PivotLayout): SourceRect[] {
  const out = [toSource(layout.rect)]
  for (const cell of layout.cells) if (!within(out, cell.row, cell.col)) out.push({ r0: cell.row, r1: cell.row, c0: cell.col, c1: cell.col })
  return out
}
interface PivotPart { saved: readonly SourceRect[]; current: readonly SourceRect[]; layout: PivotLayout }
function overlayCell(layout: PivotLayout): (row: number, col: number) => ValueOverlayCell | null {
  const cells = new Map<number, Map<number, ValueOverlayCell>>()
  for (const cell of layout.cells) {
    const value = cell.value
    if (value.kind === 'empty') continue
    let line = cells.get(cell.row)
    if (!line) cells.set(cell.row, line = new Map())
    line.set(cell.col, value.kind === 'error' ? { value: value.value, valueIsError: true } : { value: value.value, ...(cell.numFmtId !== undefined ? { numFmtId: cell.numFmtId } : {}) })
  }
  return (row, col) => cells.get(row)?.get(col) ?? null
}
/** The sheet's composite overlay over every committed pivot part (see the header's ruling). */
export function compositeOverlay(parts: readonly PivotPart[]): ValueOverlay {
  const owned = parts.map(part => ({ rects: [...part.saved, ...part.current], cell: overlayCell(part.layout), current: part.current }))
  const rects = owned.flatMap(part => part.rects)
  const protectedRects = owned.flatMap(part => part.current)
  const merges = parts.flatMap(part => part.layout.merges.map(toSource))
  const usedRows = rects.reduce((max, rect) => Math.max(max, rect.r1 + 1), 0)
  const usedCols = rects.reduce((max, rect) => Math.max(max, rect.c1 + 1), 0)
  return Object.freeze({
    rects, protectedRects, merges, usedRows, usedCols,
    cell(row: number, col: number) {
      // The pivot whose current area holds the cell answers first, so a pivot growing into a cell another one vacated wins.
      for (const part of owned) if (within(part.current, row, col)) return part.cell(row, col)
      for (const part of owned) if (within(part.rects, row, col)) return null
      return undefined
    },
  })
}

// ---- the resident client ---------------------------------------------------------------------------------------------

export interface ResidentPivotClientOptions {
  doc: XlsxDocument
  /** The file's definitions per sheet (default: the document's as the client is made); sessions open on these. */
  saved?: ReadonlyArray<readonly PivotTableDefinition[] | undefined>
  ledger: RuntimeLedger
  /** Installs a committed change in the viewer (live definition, region styles, overlay, tiles). */
  install(commit: PivotCommit): void
  profile?: ProfileId
  /** Default `ORACLE_VERIFIED_KINDS`; the menus read the open result's set. */
  verifiedKinds?: ReadonlySet<PivotDeltaKind>
  /** Task 30's pivot chart gate; absent, a pivot with linked charts refuses at open. */
  chartGate?: ChartGate
}
interface OpenPivot { session: PivotSession; savedDef: PivotTableDefinition; cache: PivotCacheDefinition; saved: SourceRect[]; current: SourceRect[]; date1904: boolean }

/** The chart's pivot name: `[book]Sheet!Name` → `Name`. */
const linkedName = (chart: ChartModel) => { const name = chart.pivot?.name ?? ''; return name.slice(name.lastIndexOf('!') + 1) }

export function createResidentPivotClient(options: ResidentPivotClientOptions): PivotClient {
  const { doc, ledger } = options
  const verified = options.verifiedKinds ?? ORACLE_VERIFIED_KINDS
  const open = new Map<string, OpenPivot>()
  const opening = new Map<string, Promise<PivotOpenResult>>()
  const key = (sheetId: number, pivot: number) => `${sheetId}:${pivot}`
  const sheetOf = (sheetId: number) => doc.sheets[sheetId]
  const savedDefs = doc.sheets.map((sheet, index) => [...(options.saved ? options.saved[index] ?? [] : sheet.pivotTables ?? [])])

  const preflight = (sheetId: number, pivot: number): PivotDiagnostic | undefined => {
    const def = savedDefs[sheetId]?.[pivot]
    if (!def) return { kind: 'unsupported', feature: 'pivot-delta-invalid', message: `Sheet ${sheetId + 1} has no pivot table ${pivot + 1}` }
    const runtime = pivotRuntimeFor(doc)
    if (!runtime) return { kind: 'unsupported', feature: 'pivot-runtime-unavailable', message: `${nameOf(def)} is shown as saved: this workbook was not loaded for pivot interaction.` }
    if (!runtime.cache(def)) return { kind: 'unsupported', feature: 'pivot-cache-absent', message: `${nameOf(def)} is shown as saved: its pivot cache is not in the file.` }
    return undefined
  }
  const parts = (sheetId: number): PivotPart[] => {
    const out: PivotPart[] = []
    // A pivot back at its saved view (reset carries the saved definition) owns no overlay: the saved cells show.
    for (const [id, entry] of open) if (id.startsWith(`${sheetId}:`) && entry.session.def !== entry.savedDef) out.push({ saved: entry.saved, current: entry.current, layout: entry.session.layout })
    return out
  }
  const occupiedFor = (sheetId: number, pivot: number) => {
    const sheet = sheetOf(sheetId)!
    const saved = sheetCellReader(sheet)
    const others = savedDefs[sheetId]!.map((def, index) => ({ index, def })).filter(entry => entry.index !== pivot)
    const savedRects = new Map(others.map(entry => [entry.index, (savedPivotLayout(entry.def)?.rects ?? []).map(toSource)]))
    return (row: number, col: number): boolean => {
      for (const { index } of others) {
        const live = open.get(key(sheetId, index))
        if (live && live.session.def !== live.savedDef) {
          if (within(live.current, row, col)) return true
          if (within(live.saved, row, col)) return false
        } else if (within(savedRects.get(index)!, row, col)) return true
      }
      return saved(row, col) !== undefined
    }
  }
  const linkedCharts = (def: PivotTableDefinition): ChartModel[] => {
    const name = String(def.name ?? '')
    return doc.sheets.flatMap(sheet => sheetCharts(sheet.drawings)).filter(chart => chart.pivot && linkedName(chart) === name)
  }
  const modelOf = (sheetId: number, entry: OpenPivot) => menuModel(sheetOf(sheetId)!, entry.session.def, entry.cache, entry.session.view, entry.date1904)

  const client: PivotClient = {
    preflight,
    labels(sheetId, pivot) {
      const def = sheetOf(sheetId)?.pivotTables?.[pivot]
      const cache = def ? pivotRuntimeFor(doc)?.cache(def) : undefined
      if (!def) return undefined
      const date1904 = pivotRuntimeFor(doc)?.date1904 ?? false
      return {
        field: field => fieldName(def, cache, field),
        item: (field, item) => cache ? itemText(sheetOf(sheetId)!, def, cache, field, item, date1904) : undefined,
      }
    },
    open(sheetId, pivot, signal) {
      const id = key(sheetId, pivot)
      const entry = open.get(id)
      if (entry) return Promise.resolve({ ok: true, revision: entry.session.revision, def: entry.session.def, verifiedKinds: verified, model: modelOf(sheetId, entry) })
      const pending = opening.get(id)
      if (pending) return pending
      const started = (async (): Promise<PivotOpenResult> => {
        const refused = preflight(sheetId, pivot)
        if (refused) return { ok: false, diagnostic: refused }
        const runtime = pivotRuntimeFor(doc)!
        const sheet = sheetOf(sheetId)!, def = savedDefs[sheetId]![pivot]!, cache = runtime.cache(def)!
        const leased = await runtime.lease(cache, ledger, signal, RECORD_LIMITS[options.profile ?? 'D'])
        if (!leased.ok) return { ok: false, diagnostic: leased.diagnostic }
        const date1904 = runtime.date1904
        const session = openPivotSession({
          def, cache, store: leased.lease.store, date1904, saved: sheetCellReader(sheet), occupied: occupiedFor(sheetId, pivot),
          linkedCharts: linkedCharts(def), lease: leased.lease, verifiedKinds: verified, ...(options.chartGate ? { chartGate: options.chartGate } : {}),
        })
        if (!session.ok) { leased.lease.release(); return session }
        const saved = savedPivotLayout(def)
        const current = layoutArea(session.value.layout)
        const opened: OpenPivot = { session: session.value, savedDef: def, cache, saved: saved ? saved.rects.map(toSource) : current, current, date1904 }
        if (signal.aborted) { session.value.release(); return refusal('pivot-session-released', `${nameOf(def)}: opening was cancelled`) }
        open.set(id, opened)
        return { ok: true, revision: 0, def: session.value.def, verifiedKinds: verified, model: modelOf(sheetId, opened) }
      })().finally(() => opening.delete(id))
      opening.set(id, started)
      return started
    },
    async apply(sheetId, pivot, delta) {
      const entry = open.get(key(sheetId, pivot))
      if (!entry) return refusal('pivot-session-released', `${nameOf(savedDefs[sheetId]?.[pivot])} has no open session`)
      const { session } = entry
      let result
      if (delta === 'reset') result = session.reset()
      else {
        const applied = session.apply(delta)
        if (!applied.ok) return applied
        result = applied.value
      }
      if (!session.commit(result)) return refusal('pivot-session-released', `${nameOf(entry.savedDef)}: the change is out of date`)
      entry.current = layoutArea(session.layout)
      const owned = parts(sheetId)
      const commit: PivotCommit = { sheetId, pivot, revision: session.revision, def: session.def, overlay: owned.length ? compositeOverlay(owned) : null }
      options.install(commit)
      return { ok: true, revision: session.revision, def: session.def, model: modelOf(sheetId, entry), commit }
    },
    release(sheetId, pivot) {
      const id = key(sheetId, pivot)
      open.get(id)?.session.release()
      open.delete(id)
    },
    dispose() {
      for (const entry of open.values()) entry.session.release()
      open.clear()
    },
  }
  return client
}

// ---- filters as the menus build them (the shapes Excel saves; tests/xlsx-pivots-interaction-oracle.test.ts) ----------

type CustomOperator = NonNullable<PivotFilter['autoFilter']['columns'][number]['customFilters']>['filters'][number]['operator']
const column = (and: boolean, rules: Array<[CustomOperator, string]>): PivotFilter['autoFilter']['columns'][number] =>
  ({ colId: 0, hiddenButton: false, showButton: true, customFilters: { and, filters: rules.map(([operator, val]) => ({ operator, val })) } })
const filterOf = (fld: number, id: number, type: PivotFilter['type'], col: PivotFilter['autoFilter']['columns'][number], extra: Partial<PivotFilter> = {}): PivotFilter =>
  ({ fld, type, evalOrder: -1, id, ...extra, autoFilter: { ref: 'A1', columns: [col] } })
/** Excel's wildcards (`*`, `?`, `~`) in literal text, escaped. */
export const escapeWildcards = (text: string) => text.replace(/[~*?]/g, match => `~${match}`)

export type LabelOp = 'eq' | 'neq' | 'begins' | 'ends' | 'contains' | 'not-contains' | 'between'
export function labelFilter(fld: number, id: number, op: LabelOp, first: string, second?: string, wildcards = false): PivotFilter {
  const a = wildcards ? first : escapeWildcards(first), b = second === undefined ? undefined : wildcards ? second : escapeWildcards(second)
  const strings = b === undefined ? { stringValue1: first } : { stringValue1: first, stringValue2: second }
  switch (op) {
    case 'eq': return filterOf(fld, id, 'captionEqual', column(false, [['equal', a]]), strings)
    case 'neq': return filterOf(fld, id, 'captionNotEqual', column(false, [['notEqual', a]]), strings)
    case 'begins': return filterOf(fld, id, 'captionBeginsWith', column(false, [['equal', `${a}*`]]), strings)
    case 'ends': return filterOf(fld, id, 'captionEndsWith', column(false, [['equal', `*${a}`]]), strings)
    case 'contains': return filterOf(fld, id, 'captionContains', column(false, [['equal', `*${a}*`]]), strings)
    case 'not-contains': return filterOf(fld, id, 'captionNotContains', column(false, [['notEqual', `*${a}*`]]), strings)
    case 'between': return filterOf(fld, id, 'captionBetween', column(true, [['greaterThanOrEqual', first], ['lessThanOrEqual', second ?? first]]), strings)
  }
}
export type ValueOp = 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'between'
const VALUE_TYPES: Record<ValueOp, PivotFilter['type']> = { eq: 'valueEqual', neq: 'valueNotEqual', gt: 'valueGreaterThan', gte: 'valueGreaterThanOrEqual', lt: 'valueLessThan', lte: 'valueLessThanOrEqual', between: 'valueBetween' }
const VALUE_RULES: Record<Exclude<ValueOp, 'between'>, CustomOperator> = { eq: 'equal', neq: 'notEqual', gt: 'greaterThan', gte: 'greaterThanOrEqual', lt: 'lessThan', lte: 'lessThanOrEqual' }
export function valueFilter(fld: number, id: number, data: number, op: ValueOp, first: number, second?: number): PivotFilter {
  const col = op === 'between' ? column(true, [['greaterThanOrEqual', String(first)], ['lessThanOrEqual', String(second ?? first)]]) : column(false, [[VALUE_RULES[op], String(first)]])
  return filterOf(fld, id, VALUE_TYPES[op], col, { iMeasureFld: data })
}
export type DateOp = 'eq' | 'neq' | 'before' | 'after' | 'between'
/** Dates as 1900-system serials (the engine refuses date filters in a 1904 workbook). */
export function dateFilter(fld: number, id: number, op: DateOp, first: number, second?: number): PivotFilter {
  switch (op) {
    case 'eq': return filterOf(fld, id, 'dateEqual', column(false, [['equal', String(first)]]))
    case 'neq': return filterOf(fld, id, 'dateNotEqual', column(false, [['notEqual', String(first)]]))
    case 'before': return filterOf(fld, id, 'dateOlderThan', column(false, [['lessThan', String(first)]]))
    case 'after': return filterOf(fld, id, 'dateNewerThan', column(false, [['greaterThan', String(first)]]))
    case 'between': return filterOf(fld, id, 'dateBetween', column(true, [['greaterThanOrEqual', String(first)], ['lessThanOrEqual', String(second ?? first)]]))
  }
}
/** Top 10 (Excel's fields: top, count, kind and data field). Bottom is not offered: it refuses until measured. */
export function top10Filter(fld: number, id: number, data: number, kind: 'count' | 'percent' | 'sum', val: number): PivotFilter {
  return filterOf(fld, id, kind, { colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: kind === 'percent', val } }, { iMeasureFld: data })
}
