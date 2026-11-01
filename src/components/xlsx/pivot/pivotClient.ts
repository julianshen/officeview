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
 *   `RemotePivotSessions`, for a paged or worker-parsed workbook (`remotePivotPort`).
 *
 * `sheetId` in this interface is the sheet's position in `doc.sheets` (the plan's `number`), not the viewer's string
 * sheet id; the viewer maps between them.
 *
 * Rulings recorded here (Task 28):
 * - The overlay a commit carries is the sheet's composite over every committed pivot. Each pivot owns its saved area
 *   (`savedPivotLayout(savedDef).rects`, page cells included) and its current area: a cell of the current layout reads
 *   the layout's value, any other cell of either area reads blank (`null`). Owning the whole saved area, not only the
 *   last change's `vacated`, keeps rows an earlier change vacated blank through every later change.
 * - Protection (Ruling S3-G): the overlay's `protectedRects` are every committed pivot's saved area ∪ its current area,
 *   so a sheet sort or filter never reaches the base cells a vacated (blank) saved area still holds.
 * - Overlap (`occupied`, §7.1, refined by Task 29): a cell inside another pivot's current area (or, for a pivot never
 *   opened, its saved rectangles) is occupied; a cell another pivot vacated is free; anything else is occupied when
 *   the sheet as loaded has a value there (`sheetCellReader`, as the oracle does).
 * - Linked charts: every chart on any sheet whose `c:pivotSource` names this pivot by its sheet and name
 *   (`chartLinksPivot`; pivot names repeat across sheets, Task 27 r2 R2-4). They go
 *   through Task 30's `pivotChartGate` (lazily loaded); a commit carries the refilled models, in the same order as
 *   `installPivotCharts` walks them.
 */
import type { XlsxCell, XlsxCellStyle, XlsxDocument, XlsxSheet } from '../../../xlsx/types'
import type { StyledValueOverlayCell } from '../../../xlsx/view/viewport'
import type { ChartModel } from '../../../charts/model'
import type { ProfileId, ValueOverlay, ValueOverlayCell } from '../../../xlsx/view/contracts'
import type { SourceRect } from '../../../xlsx/view/coordinates'
import { cellDisplayText } from '../../../xlsx/render'
import { pivotRuntimeFor } from '../../../xlsx/pivotLoader'
import { MAIN_QUOTA_MESSAGE, type RuntimeLedger } from '../../../xlsx/pivots/session/runtime'
import { MAIN_THREAD_BUDGETS } from '../../../xlsx/view/tileCache'
import type { PivotCacheDefinition, PivotDiagnostic, PivotFilter, PivotTableDefinition } from '../../../xlsx/pivots/model'
import type { SavedPivotLayout } from '../../../xlsx/pivots/savedLayout'
import {
  macDigitWidth, viewerFitTolerance, openPivotSession, ORACLE_VERIFIED_KINDS, PIVOT_FIT_INDENT_PT, pivotAutofit, pivotColumnFits, RECORD_LIMITS, savedPivotLayout, sheetCellReader, worksheetSourceRange,
  type ChartGate, type PivotDelta, type PivotDeltaKind, type PivotLayout, type PivotSession, type PivotViewState,
} from '../../../xlsx/pivots/index'
import { sheetCharts } from '../sheetCharts'
import { linkPivotCharts } from '../../../xlsx/pivotChartLink'
import { pivotFieldName, pivotMenuItemLabel, pivotMenuModel, type PivotMenuField, type PivotMenuItem, type PivotMenuModel } from '../../../xlsx/pivots/session/menu'
import { registerGetPivotData } from '../../../xlsx/pivots/session/getpivotdata'
import { sheetBlockers, structuralOccupancy } from '../../../xlsx/pivots/session/occupancy'
import { engineSpillFootprints } from '../../../xlsx/formula/workbook'
import { estimatePayloadBytes } from '../../../xlsx/view/payloadBytes'
import { OwnedRects, uniqueRectList } from '../../../xlsx/view/rectIndex'

registerGetPivotData()

export type { PivotDelta, PivotDeltaKind }

export type { PivotMenuField, PivotMenuItem, PivotMenuModel }
/** What a commit installs in the viewer (spec §7.9): the live definition and, on the resident path, the sheet's
 * composite overlay (`null` once every pivot of the sheet is back to its saved view: the saved cells show again,
 * their own styles included). The paged path's overlay lives in the worker (Task 27); its commit carries none. */
export interface PivotCommit {
  sheetId: number; pivot: number; revision: number; def: PivotTableDefinition; overlay?: ValueOverlay | null
  /** True when this commit restores this pivot to its saved view. */
  reset?: boolean
  /** Resident: after this commit no pivot of the document shows a live layout (every one is back to its saved view),
   * so the viewer may put back every saved dependent value at once (Ruling S3-AF). */
  allSaved?: boolean
  /** Resident-only live formula context; never sent through the worker protocol. */
  layout?: PivotLayout; cache?: PivotCacheDefinition; date1904?: boolean
  /** Saved and current rectangles affected by this commit, in source coordinates. */
  changedRects?: readonly SourceRect[]
  /** The pivot's linked charts refilled for this revision (Task 30's gate), in the order the client collected them
   * (`doc.sheets` × drawing order); the viewer installs them with `installPivotCharts`. */
  charts?: readonly ChartModel[]
  /** Paged (Task 29): the pivot's shown cells before and after the change, for the main thread's column autofit. */
  fit?: { before: PivotShownLayout; after: PivotShownLayout }
}
/** A pivot's definition and its layout cells as the worker replies carry them. */
export interface PivotShownLayout { def: PivotTableDefinition; cells: readonly { row: number; col: number; value: ValueOverlayCell | null }[] }
export type PivotOpenResult =
  | { ok: true; revision: number; def: PivotTableDefinition; verifiedKinds: ReadonlySet<PivotDeltaKind>; model: PivotMenuModel }
  | { ok: false; diagnostic: PivotDiagnostic }
/** `notice`: a refreshOnLoad pivot's one-line notice, on its first committed change (§7.7, Task 32). */
export type PivotApplyOutcome =
  | { ok: true; revision: number; def: PivotTableDefinition; model: PivotMenuModel; commit: PivotCommit; notice?: string }
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
 * What the paged path supplies: open, apply and release over the worker protocol, the apply resolving only after the
 * worker committed the staged candidate (`pivotCommit`), so the viewer installs a committed revision. The worker owns
 * the overlay; `RemoteWorkbook` publishes its merges, extent, valueGen and the refilled pivot charts (S3-M) on the
 * acknowledged commit. The menu model is built in the worker, where the cache is (`session/menu.ts`).
 */
export interface RemotePivotPort {
  open(sheetId: number, pivot: number, signal: AbortSignal): Promise<PivotOpenResult>
  apply(sheetId: number, pivot: number, delta: PivotDelta | 'reset'): Promise<PivotApplyOutcome>
  release(sheetId: number, pivot: number): void
  labels?(sheetId: number, pivot: number): PivotLabels | undefined
  dispose?(): void
}
/** Task 27's `RemotePivotSessions` as this module reads it (its `sheetId` is the worker's sheet id as a number). */
export interface RemotePivotSessionsLike {
  open(sheetId: number, pivot: number, signal: AbortSignal): Promise<{ kind: 'open'; revision: number; def: PivotTableDefinition; menu: PivotMenuModel; cells?: PivotShownLayout['cells'] } | { kind: 'refused'; message: string }>
  apply(sheetId: number, pivot: number, delta: PivotDelta | 'reset'): Promise<{ kind: 'committed'; revision: number; def: PivotTableDefinition; menu: PivotMenuModel; notice?: string; cells?: PivotShownLayout['cells'] } | { kind: 'stale' | 'cancelled' } | { kind: 'refused'; message: string }>
  release(sheetId: number, pivot: number): void
}
const OPEN_FIRST = 'Open the pivot before changing it.'
/**
 * The port over `RemotePivotSessions`. Sheets are addressed here by position (`doc.sheets`), there by the worker's
 * numeric sheet id. A session the workbook released (a sheet switch releases the others, `releaseExcept`) is opened
 * again before an apply; a stale or cancelled apply is refused with a retry message (a newer one won). The menu
 * model's numbers and dates are labelled on this thread (`formatMenu`).
 */
/** Main-ledger admission of fit cells the main thread keeps (codex review T29-3): false refuses (nothing is kept). */
export interface PivotFitMemory { set(owner: string, bytes: number): boolean; release(owner: string): void }
let fitMemorySerial = 0
export function remotePivotPort(sessions: RemotePivotSessionsLike, sheets: { id(sheetId: number): number | undefined; sheet(sheetId: number): XlsxSheet | undefined }, memory?: PivotFitMemory): RemotePivotPort {
  const models = new Map<string, PivotMenuModel>(), opened = new Set<string>()
  // The cells each open pivot shows (its last committed or opened layout): the "before" of the next change's autofit,
  // charged while kept; one the budget refuses is not kept (that pivot's next change is not autofitted).
  const shown = new Map<string, PivotShownLayout>(), serial = ++fitMemorySerial
  // The port's own lifetime: a reopen it makes on its own (a session the workbook released) ends with it (minor 14).
  const lifetime = new AbortController()
  const owner = (k: string) => `pivot-shown:${serial}:${k}`
  const keep = (k: string, layout: PivotShownLayout | undefined) => {
    if (layout && (!memory || memory.set(owner(k), estimatePayloadBytes(layout.cells)))) { shown.set(k, layout); return }
    shown.delete(k); memory?.release(owner(k))
  }
  const key = (sheetId: number, pivot: number) => `${sheetId}:${pivot}`
  const refusedAs = (feature: string, message: string): { ok: false; diagnostic: PivotDiagnostic } => ({ ok: false, diagnostic: { kind: 'unsupported', feature, message } })
  const port: RemotePivotPort = {
    async open(sheetId, pivot, signal) {
      const id = sheets.id(sheetId)
      if (id === undefined) return refusedAs('pivot-remote-unavailable', REMOTE_PIVOTS_UNAVAILABLE.message)
      const result = await sessions.open(id, pivot, signal)
      if (result.kind === 'refused') return refusedAs('pivot-remote-refused', result.message)
      const model = formatMenu(result.menu, sheets.sheet(sheetId))
      models.set(key(sheetId, pivot), model); opened.add(key(sheetId, pivot))
      keep(key(sheetId, pivot), result.cells ? { def: result.def, cells: result.cells } : undefined)
      return { ok: true, revision: result.revision, def: result.def, verifiedKinds: ORACLE_VERIFIED_KINDS, model }
    },
    async apply(sheetId, pivot, delta) {
      const id = sheets.id(sheetId)
      if (id === undefined) return refusedAs('pivot-remote-unavailable', REMOTE_PIVOTS_UNAVAILABLE.message)
      let result = await sessions.apply(id, pivot, delta)
      if (result.kind === 'refused' && result.message === OPEN_FIRST) {
        const reopened = await sessions.open(id, pivot, lifetime.signal)
        if (reopened.kind === 'refused') return refusedAs('pivot-remote-refused', reopened.message)
        keep(key(sheetId, pivot), reopened.cells ? { def: reopened.def, cells: reopened.cells } : undefined)
        result = await sessions.apply(id, pivot, delta)
      }
      if (result.kind === 'refused') return refusedAs('pivot-remote-refused', result.message)
      if (result.kind !== 'committed') return refusedAs('pivot-session-released', 'The pivot table changed meanwhile; try again.')
      const model = formatMenu(result.menu, sheets.sheet(sheetId))
      models.set(key(sheetId, pivot), model)
      const before = shown.get(key(sheetId, pivot)), after = result.cells ? { def: result.def, cells: result.cells } : undefined
      keep(key(sheetId, pivot), after)
      return { ok: true, revision: result.revision, def: result.def, model, commit: { sheetId, pivot, revision: result.revision, def: result.def, ...(before && after ? { fit: { before, after } } : {}) }, ...(result.notice ? { notice: result.notice } : {}) }
    },
    release(sheetId, pivot) {
      const k = key(sheetId, pivot), id = sheets.id(sheetId)
      models.delete(k); keep(k, undefined)
      if (opened.delete(k) && id !== undefined) sessions.release(id, pivot)
    },
    labels(sheetId, pivot) {
      const model = models.get(key(sheetId, pivot))
      if (!model) return undefined
      return {
        field: field => model.fields.find(entry => entry.index === field)?.name ?? `Field ${field + 1}`,
        item: (field, item) => model.fields.find(entry => entry.index === field)?.items.find(entry => entry.index === item)?.label,
      }
    },
    dispose() { lifetime.abort(); for (const k of [...opened]) { const [sheetId, pivot] = k.split(':').map(Number) as [number, number]; port.release(sheetId, pivot) } },
  }
  return port
}
/** The paged client: the port, the viewer's install run on each committed outcome. */
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
    dispose: () => port.dispose?.(),
  }
}
/** A sheet the worker cannot address (no numeric sheet id): its pivots stay as saved. */
const REMOTE_PIVOTS_UNAVAILABLE = { feature: 'pivot-remote-unavailable', message: 'Pivot tables on this sheet are shown as saved: the large-file view cannot address the sheet.' } as const

// ---- the menu model --------------------------------------------------------------------------------------------------

/** A number or date item's label in its format (the model carries the value; the menus show this). */
function formatItem(sheet: XlsxSheet | undefined, item: Pick<PivotMenuItem, 'label' | 'number'>): string {
  if (!item.number || !sheet) return item.label
  try { return cellDisplayText(sheet, { value: item.number.value, style: { numFmtId: item.number.numFmtId } }) } catch { return item.label }
}
/** The model with its number and date items labelled in their formats. */
export function formatMenu(model: PivotMenuModel, sheet: XlsxSheet | undefined): PivotMenuModel {
  return { ...model, fields: model.fields.map(field => ({ ...field, items: field.items.map(item => item.number ? { ...item, label: formatItem(sheet, item) } : item) })) }
}
export const menuModel = (sheet: XlsxSheet, def: PivotTableDefinition, cache: PivotCacheDefinition, view: PivotViewState, date1904: boolean): PivotMenuModel =>
  formatMenu(pivotMenuModel(def, cache, view, date1904), sheet)
const fieldName = pivotFieldName
const itemText = (sheet: XlsxSheet, def: PivotTableDefinition, cache: PivotCacheDefinition, field: number, item: number, date1904: boolean): string =>
  formatItem(sheet, pivotMenuItemLabel(def, cache, field, item, date1904))

// ---- the overlay -----------------------------------------------------------------------------------------------------

const toSource = (rect: { minRow: number; maxRow: number; minCol: number; maxCol: number }): SourceRect => ({ r0: rect.minRow, r1: rect.maxRow, c0: rect.minCol, c1: rect.maxCol })
const within = (rects: readonly SourceRect[], row: number, col: number) => rects.some(rect => row >= rect.r0 && row <= rect.r1 && col >= rect.c0 && col <= rect.c1)
/** The cells a layout covers: its rectangle, then one cell per page cell outside it. */
function layoutArea(layout: PivotLayout): SourceRect[] {
  const out = [toSource(layout.rect)]
  for (const cell of layout.cells) if (!within(out, cell.row, cell.col)) out.push({ r0: cell.row, r1: cell.row, c0: cell.col, c1: cell.col })
  return out
}
/** `style`: the style a cell of the layout paints with (alignment, font, fill), its number format aside. */
interface PivotPart { saved: readonly SourceRect[]; current: readonly SourceRect[]; layout: PivotLayout; style?: (row: number, col: number) => XlsxCellStyle | undefined }
function overlayCell(part: PivotPart): (row: number, col: number) => ValueOverlayCell | null {
  const cells = new Map<number, Map<number, StyledValueOverlayCell>>()
  for (const cell of part.layout.cells) {
    const value = cell.value
    if (value.kind === 'empty') continue
    let line = cells.get(cell.row)
    if (!line) cells.set(cell.row, line = new Map())
    const style = part.style?.(cell.row, cell.col)
    line.set(cell.col, value.kind === 'error' ? { value: value.value, valueIsError: true, ...(style ? { style } : {}) }
      : { value: value.value, ...(cell.numFmtId !== undefined ? { numFmtId: cell.numFmtId } : {}), ...(style ? { style } : {}) })
  }
  return (row, col) => cells.get(row)?.get(col) ?? null
}
/** A pivot cell's role for its style (Task 28): its place in the header rows, or its column and its row line's kind
 * and depth in the body; page cells by their offset. */
function roleOf(layout: SavedPivotLayout, row: number, col: number): string | undefined {
  const { rect, dataRow0 } = layout
  const c = col - rect.minCol
  if (row < rect.minRow || row > rect.maxRow || col < rect.minCol || col > rect.maxCol) return `page|${row - rect.minRow}|${c}`
  if (row < dataRow0) return `head|${row - rect.minRow}|${c}`
  const line = layout.rows[row - dataRow0]
  return line ? `body|${c}|${line.type}|${line.items.length}` : undefined
}
/**
 * The style a pivot cell keeps (ruling, Task 28): Excel writes one alignment, indent and font per pivot area (the
 * pivot style itself is the region layer's), so a cell of the live layout takes the saved style of a cell in the same
 * role (header row; column and line kind and depth in the body), else the saved cell's own at that position, else
 * the workbook's default cell format (`cellXfs[0]`), as a cell Excel newly writes gets.
 */
function savedStyles(doc: XlsxDocument, sheet: XlsxSheet, savedDef: PivotTableDefinition, saved: readonly SourceRect[]): (live: SavedPivotLayout | undefined) => (row: number, col: number) => XlsxCellStyle | undefined {
  let rows: Map<number, XlsxSheet['rows'][number]> | undefined
  let roles: Map<string, XlsxCellStyle> | undefined
  const fallback = doc.cellStyles?.[0]
  const own = (row: number, col: number) => {
    rows ??= new Map(sheet.rows.map(entry => [entry.index, entry]))
    return rows.get(row)?.cells.find(cell => cell.col === col)?.style
  }
  const byRole = () => {
    if (roles) return roles
    roles = new Map()
    const layout = savedPivotLayout(savedDef)
    if (!layout) return roles
    for (const rect of layout.rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) {
      const key = roleOf(layout, row, col), style = own(row, col)
      if (key !== undefined && style && !roles.has(key)) roles.set(key, style)
    }
    return roles
  }
  return live => (row, col) => {
    const key = live ? roleOf(live, row, col) : undefined
    const role = key !== undefined ? byRole().get(key) : undefined
    if (role) return role
    if (within(saved, row, col)) { const style = own(row, col); if (style) return style }
    return fallback
  }
}
/** The sheet's composite overlay over every committed pivot part (see the header's ruling). */
/** `indexBytes`: the retained bytes of its two cell indexes (the caller charges them while it keeps the overlay). */
export function compositeOverlay(parts: readonly PivotPart[]): ValueOverlay & { readonly indexBytes: number } {
  const owned = parts.map(part => ({ rects: [...part.saved, ...part.current], cell: overlayCell(part), current: part.current }))
  const rects = owned.flatMap(part => part.rects)
  // Ruling S3-G: sheet sort and filter stay refused over every saved pivot area and every current one (page cells
  // included); a vacated saved area reads blank but is still the pivot's.
  const protectedRects = uniqueRectList(rects)
  const merges = parts.flatMap(part => part.layout.merges.map(toSource))
  // Indexed by part (Task 33 perf): a read is one point query, not a walk over every pivot.
  const currents = new OwnedRects(owned.flatMap((part, k) => part.current.map(rect => ({ rect, owner: k }))))
  const areas = new OwnedRects(owned.flatMap((part, k) => part.rects.map(rect => ({ rect, owner: k }))))
  return Object.freeze({
    rects, protectedRects, merges, indexBytes: currents.estimatedBytes + areas.estimatedBytes,
    cell(row: number, col: number) {
      // The pivot whose current area holds the cell answers first, so a pivot growing into a cell another one vacated
      // wins; among several, the first part (as the parts are listed).
      const holders = currents.at(row, col)
      if (holders.length) return owned[Math.min(...holders)]!.cell(row, col)
      return areas.has(row, col) ? null : undefined
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
  /** The main-thread budget the record stores and their decodes are admitted against (final review I3); default
   * `MAIN_THREAD_BUDGETS[profile]`, as the viewer's other main-thread admissions. */
  budgetBytes?: number
}
interface OpenPivot { session: PivotSession; savedDef: PivotTableDefinition; cache: PivotCacheDefinition; saved: SourceRect[]; current: SourceRect[]; date1904: boolean; style: ReturnType<typeof savedStyles> }


let refreshBuilds = 0
export function createResidentPivotClient(options: ResidentPivotClientOptions): PivotClient {
  const { doc, ledger } = options
  const verified = options.verifiedKinds ?? ORACLE_VERIFIED_KINDS
  const open = new Map<string, OpenPivot>()
  const opening = new Map<string, Promise<PivotOpenResult>>()
  const key = (sheetId: number, pivot: number) => `${sheetId}:${pivot}`
  const sheetOf = (sheetId: number) => doc.sheets[sheetId]
  const savedDefs = doc.sheets.map((sheet, index) => [...(options.saved ? options.saved[index] ?? [] : sheet.pivotTables ?? [])])
  // Final review I3: a session's record store, the decode that builds it, and a refresh's build are admitted against
  // the main-thread budget before they are charged, as the worker admits them to its pool; refused, the open is.
  const budget = options.budgetBytes ?? MAIN_THREAD_BUDGETS[options.profile ?? 'D']
  /** Charges `bytes` under `owner` (replacing its earlier charge) only when the ledger stays within the budget; the
   * caller zeroes the owner when the work ends. */
  const admitAs = (owner: string) => {
    let held = 0
    return (bytes: number): boolean => {
      if (ledger.bytes - held + bytes > budget) return false
      ledger.set('document', owner, held = bytes)
      return true
    }
  }
  /** A decode's or a build's refusal as the main thread says it (the engine words it for the worker). */
  const mainQuota = (diagnostic: PivotDiagnostic): PivotDiagnostic => diagnostic.feature === 'pivot-worker-quota' ? { ...diagnostic, feature: 'pivot-records-limit', message: MAIN_QUOTA_MESSAGE } : diagnostic

  const preflight = (sheetId: number, pivot: number): PivotDiagnostic | undefined => {
    const def = savedDefs[sheetId]?.[pivot]
    if (!def) return { kind: 'unsupported', feature: 'pivot-delta-invalid', message: `Sheet ${sheetId + 1} has no pivot table ${pivot + 1}` }
    const runtime = pivotRuntimeFor(doc)
    if (!runtime) return { kind: 'unsupported', feature: 'pivot-runtime-unavailable', message: `${nameOf(def)} is shown as saved: this workbook was not loaded for pivot interaction.` }
    if (!runtime.cache(def)) return { kind: 'unsupported', feature: 'pivot-cache-absent', message: `${nameOf(def)} is shown as saved: its pivot cache is not in the file.` }
    return undefined
  }
  /** Whether a pivot shows its live layout: it committed a change, and is not back at its saved view (a refreshed
   * pivot's reset shows the refreshed view; an uncommitted refreshed session still shows the saved cells). */
  const shows = (entry: OpenPivot) => entry.session.revision > 0 && entry.session.def !== entry.savedDef
  const parts = (sheetId: number): PivotPart[] => {
    const out: PivotPart[] = []
    // A pivot back at its saved view (reset carries the saved definition) owns no overlay: the saved cells show.
    for (const [id, entry] of open) if (id.startsWith(`${sheetId}:`) && shows(entry)) out.push({ saved: entry.saved, current: entry.current, layout: entry.session.layout, style: entry.style(savedPivotLayout(entry.session.def)) })
    return out
  }
  // Ruling S3-AC: every pivot's saved area (vacated cells too, which it protects while it is open and gets back on a
  // reset) and the current area of each one showing a change, indexed per sheet (Task 33 perf: one point query per
  // cell, not a walk over every pivot); rebuilt after a commit or a release.
  const blockers = new Map<number, OwnedRects<number>>()
  const blockersOf = (sheetId: number): OwnedRects<number> => {
    let index = blockers.get(sheetId)
    if (index) return index
    const entries: { rect: SourceRect; owner: number }[] = []
    savedDefs[sheetId]!.forEach((def, pivot) => { for (const rect of savedPivotLayout(def)?.rects ?? []) entries.push({ rect: toSource(rect), owner: pivot }) })
    for (const [id, entry] of open) {
      const [sheet, pivot] = id.split(':').map(Number) as [number, number]
      if (sheet === sheetId && shows(entry)) for (const rect of entry.current) entries.push({ rect, owner: pivot })
    }
    blockers.set(sheetId, index = new OwnedRects(entries))
    // Charged while the client keeps it (perf round 2); replaced after a commit or release, it is collected.
    ledger.chargeWhileReachable('structure', [index], index.estimatedBytes)
    return index
  }
  const occupiedFor = (sheetId: number, pivot: number) => {
    const sheet = sheetOf(sheetId)!
    const saved = sheetCellReader(sheet)
    // Merges, array and spill ranges and noted cells (S3-AC), charged while the session holds the check.
    const structure = structuralOccupancy(sheetBlockers(sheet, engineSpillFootprints(doc)?.get(sheet)))
    const occupied = (row: number, col: number): boolean =>
      blockersOf(sheetId).find(row, col, owner => owner !== pivot) !== undefined || saved(row, col) !== undefined || structure.occupied(row, col)
    ledger.chargeWhileReachable('structure', [occupied], structure.bytes)
    return occupied
  }
  // Ruling S3-Q: an unqualified source naming same-named pivots on several sheets links none, and they refuse.
  const linkedCharts = (def: PivotTableDefinition) => linkPivotCharts(doc.sheets.flatMap(sheet => sheetCharts(sheet.drawings)), { sheet: def.sheet, name: String(def.name ?? '') },
    doc.sheets.flatMap(sheet => (sheet.pivotTables ?? []).map(table => ({ sheet: table.sheet, name: String(table.name ?? '') }))))
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
        const limits = RECORD_LIMITS[options.profile ?? 'D']
        // §7.7 (Task 32): a refreshOnLoad pivot whose source range is a sheet of this workbook is rebuilt from that
        // sheet's cells by the session; its saved records (stale, or absent with saveData="0") are not read.
        const range = cache.refreshOnLoad ? worksheetSourceRange(cache) : undefined
        const sourceSheet = range ? doc.sheets.find(candidate => candidate.name === range.sheet) : undefined
        let leased: { release(): void; store?: import('../../../xlsx/pivots/engine/records').RecordStore } | undefined
        if (!sourceSheet) {
          // The decode's arrays are charged as they grow, under their own owner, until the store's own charge takes over.
          const decoding = `pivot-decode:${++refreshBuilds}`
          let result: Awaited<ReturnType<typeof runtime.lease>>
          try {
            result = await runtime.lease(cache, ledger, signal, { ...limits, admitBytes: admitAs(decoding) },
              bytes => { ledger.set('document', decoding, 0); return ledger.bytes + bytes <= budget })
          } finally { ledger.set('document', decoding, 0) }
          if (!result.ok) return { ok: false, diagnostic: mainQuota(result.diagnostic) }
          leased = result.lease
        } else runtime.chargeRetained(ledger)
        const date1904 = runtime.date1904
        // Task 30's chart gate, loaded only for a pivot with linked charts (its module statically imports the charts chunk).
        const links = linkedCharts(def)
        if (links.ambiguous) { leased?.release(); return refusal('pivot-chart-mapping', `${nameOf(def)} cannot follow its charts: a chart names "${String(def.name ?? '')}" without its sheet, and more than one sheet has a pivot table of that name.`) }
        const linked = links.charts
        let chartGate = options.chartGate
        if (!chartGate && linked.length) {
          try { chartGate = (await import('../../../xlsx/pivots/session/charts')).pivotChartGate } catch { /* no gate: the session refuses pivot-chart-mapping */ }
          if (signal.aborted) { leased?.release(); return refusal('pivot-session-released', `${nameOf(def)}: opening was cancelled`) }
        }
        // A refresh's construction (the decoder's arrays and the build's distinct texts and items) is charged to the
        // document pool as it grows, under its own owner, until the session holds the store (codex review T32-1).
        const building = `pivot-refresh-build:${++refreshBuilds}`
        const session = openPivotSession({
          def, cache, ...(leased?.store ? { store: leased.store } : {}), date1904, saved: sheetCellReader(sheet), occupied: occupiedFor(sheetId, pivot),
          linkedCharts: linked, ...(leased ? { lease: leased } : {}), verifiedKinds: verified, ...(chartGate ? { chartGate } : {}),
          ...(sourceSheet ? {
            source: sheetCellReader(sourceSheet), limits: { ...limits, admitBytes: admitAs(building) },
            // The rebuilt store and cache metadata are admitted, then charged to the document pool while the session
            // holds them.
            admitStore: (store, metadataBytes) => {
              ledger.set('document', building, 0)
              if (ledger.bytes + store.bytes + metadataBytes > budget) return undefined
              ledger.chargeWhileHeld('document', [store], store.bytes + metadataBytes); return ledger.hold(store) ?? undefined
            },
          } : {}),
        })
        ledger.set('document', building, 0)
        if (!session.ok) { leased?.release(); return { ok: false, diagnostic: mainQuota(session.diagnostic) } }
        const saved = savedPivotLayout(def)
        const savedArea = saved ? saved.rects.map(toSource) : layoutArea(session.value.layout)
        // Until its first commit the sheet shows the saved cells (a refreshed pivot's revision 0 is not shown yet).
        const current = savedArea
        const opened: OpenPivot = { session: session.value, savedDef: def, cache: session.value.cache, saved: savedArea, current, date1904, style: savedStyles(doc, sheet, def, savedArea) }
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
      const applied = delta === 'reset' ? session.reset() : session.apply(delta)
      if (!applied.ok) return applied
      const result = applied.value
      if (!session.commit(result)) return refusal('pivot-session-released', `${nameOf(entry.savedDef)}: the change is out of date`)
      // Ruling S3-AF: the dependents to recompute read the saved area, the area shown before this commit and the one
      // shown after (a shrink or reset vacates cells the previous layout filled).
      const before = entry.current
      entry.current = layoutArea(session.layout)
      blockers.delete(sheetId)
      const owned = parts(sheetId)
      const overlay = owned.length ? compositeOverlay(owned) : null
      // Its cell indexes are charged while the viewer keeps the overlay (perf round 2).
      if (overlay) ledger.chargeWhileReachable('structure', [overlay], overlay.indexBytes)
      const commit: PivotCommit = { sheetId, pivot, revision: session.revision, def: session.def, overlay, ...(delta === 'reset' ? { reset: true } : {}),
        layout: result.layout, cache: entry.cache, date1904: entry.date1904,
        changedRects: [...entry.saved, ...before, ...entry.current], allSaved: ![...open.values()].some(shows),
        ...(result.charts.length ? { charts: result.charts } : {}) }
      options.install(commit)
      return { ok: true, revision: session.revision, def: session.def, model: modelOf(sheetId, entry), commit, ...(result.notice ? { notice: result.notice } : {}) }
    },
    release(sheetId, pivot) {
      const id = key(sheetId, pivot)
      open.get(id)?.session.release()
      open.delete(id)
      blockers.delete(sheetId)
    },
    dispose() {
      for (const entry of open.values()) entry.session.release()
      open.clear()
      blockers.clear()
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

// ---- column autofit after a change (Task 29, §7.8) -------------------------------------------------------------------

/** A pivot cell as the viewer shows it: its value and its own style (the pivot style resolves over it). */
export interface PivotShownCell { value: unknown; style?: XlsxCellStyle }
/**
 * The columns a commit resizes, in characters (session/autofit.ts's rule). The fits are measured on this thread:
 * `before` and `after` read the cells shown before and after the change, each with the region styles of its own
 * definition (`styleAt`); `measure` is the painter's text measure (CSS px for a font and a text).
 */
export function pivotAutofitColumns(input: {
  sheet: XlsxSheet
  before: { def: PivotTableDefinition; read(row: number, col: number): PivotShownCell | undefined; styleAt?(row: number, col: number, own: XlsxCellStyle | undefined): XlsxCellStyle | undefined }
  after: { def: PivotTableDefinition; read(row: number, col: number): PivotShownCell | undefined; styleAt?(row: number, col: number, own: XlsxCellStyle | undefined): XlsxCellStyle | undefined }
  measure(font: string, text: string): number
  font(style: XlsxCellStyle | undefined): string
  mdwPx: number
  current(col: number): number | undefined
  userSet(col: number): boolean
}): Map<number, number> {
  const mdw = macDigitWidth(input.mdwPx)
  const fits = (side: typeof input.before) => {
    const layout = savedPivotLayout(side.def)
    if (!layout) return new Map<number, number>()
    const cells = function* () { for (const rect of layout.rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) yield { row, col } }
    return pivotColumnFits({ def: side.def, cells: cells() }, { textPt(row, col) {
      const cell = side.read(row, col)
      if (!cell || cell.value === null || cell.value === undefined || cell.value === '') return undefined
      const style = side.styleAt?.(row, col, cell.style) ?? cell.style
      const text = cellDisplayText(input.sheet, { value: cell.value as XlsxCell['value'], ...(style ? { style } : {}) })
      const own = (style as { indentPx?: number } | undefined)?.indentPx ? 0 : (style?.indent ?? 0) * PIVOT_FIT_INDENT_PT
      return input.measure(input.font(style), text) * 0.75 + own
    } }, mdw)
  }
  return pivotAutofit(fits(input.before), fits(input.after), input.current, input.userSet, viewerFitTolerance)
}

/** A paged commit's fit waiting for its sheet's refreshed pivot display (its styles; Task 29). */
export interface PendingPivotFit<S> { pivot: number; fit: NonNullable<PivotCommit['fit']>; styleAt: S; styleGen: string }
/**
 * The paged fits waiting per sheet, one per pivot (codex review T29-6): a second commit of the same pivot before the
 * display refreshed keeps the first's "before" and takes the new "after" (the cells shown between were never fitted);
 * another pivot's commit waits beside it. `take` hands over, in commit order, a sheet's fits whose display moved on.
 */
export function pendingPivotFits<S>(memory?: PivotFitMemory) {
  const bySheet = new Map<string, Map<number, PendingPivotFit<S>>>(), serial = ++fitMemorySerial
  // Each waiting fit's cells (before and after) are charged until it is fitted or dropped (codex review T29-3).
  const owner = (sheetId: string, pivot: number) => `pivot-fit:${serial}:${sheetId}:${pivot}`
  const drop = (sheetId: string, fits: Map<number, PendingPivotFit<S>>, pivot: number) => {
    fits.delete(pivot); memory?.release(owner(sheetId, pivot))
    if (!fits.size) bySheet.delete(sheetId)
  }
  return {
    get size() { let count = 0; for (const fits of bySheet.values()) count += fits.size; return count },
    sheets: () => [...bySheet.keys()],
    add(sheetId: string, pending: PendingPivotFit<S>): void {
      let fits = bySheet.get(sheetId)
      if (!fits) bySheet.set(sheetId, fits = new Map())
      const earlier = fits.get(pending.pivot)
      const next = earlier ? { ...pending, fit: { before: earlier.fit.before, after: pending.fit.after }, styleAt: earlier.styleAt } : pending
      if (memory && !memory.set(owner(sheetId, pending.pivot), estimatePayloadBytes(next.fit.before.cells) + estimatePayloadBytes(next.fit.after.cells))) { drop(sheetId, fits, pending.pivot); return }
      fits.set(pending.pivot, next)
    },
    /** The ready fits, removed; each stays charged until the caller's `release()` (in `finally`, after fitting:
     * codex re-review T29-R2-3). */
    take(sheetId: string, styleGen: string): Array<PendingPivotFit<S> & { release(): void }> {
      const fits = bySheet.get(sheetId)
      if (!fits) return []
      const ready = [...fits.values()].filter(pending => pending.styleGen !== styleGen)
      for (const pending of ready) fits.delete(pending.pivot)
      if (!fits.size) bySheet.delete(sheetId)
      return ready.map(pending => {
        let held = true
        return { ...pending, release: () => { if (held) { held = false; memory?.release(owner(sheetId, pending.pivot)) } } }
      })
    },
    /** Drops a sheet's waiting fits and their charges (its pivot display can no longer arrive: no autofit). */
    discard(sheetId: string): void {
      const fits = bySheet.get(sheetId)
      if (fits) for (const pivot of [...fits.keys()]) drop(sheetId, fits, pivot)
    },
    clear(): void {
      for (const [sheetId, fits] of bySheet) for (const pivot of [...fits.keys()]) drop(sheetId, fits, pivot)
    },
  }
}
export type PendingPivotFits<S> = ReturnType<typeof pendingPivotFits<S>>
