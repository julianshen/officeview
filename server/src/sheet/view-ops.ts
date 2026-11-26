/**
 * Stage 4c (S4 §6, plan P3, P5, P6, P13): the `sheet-view` pool job. Sort, filter, value lists and range detection
 * run through the local engine (`viewerControls`) over the key view (`keyDocument`), one port per
 * `(sheetIndex, clock)`.
 *
 * - The port's `view`, `snapshot` and `publish` read and write the job's current base; one job runs per port at a
 *   time.
 * - A derived view is a pure function of `(key view, sort, sortFilter, filter)` (P2). `derive` replays
 *   filter(`sortFilter`), sort(`sort`), filter(`filter`) from the identity view; a refused or failed replay step
 *   throws (fail closed, P3).
 * - C1 (parent §12): under `'visible-only'`, `publish` keeps a filter's body row visible iff it matches and is not an
 *   authored collapsed row (`collapsedRows`, the snapshot geometry's zero-size row segments). Under `'full'` the local
 *   mask stands.
 * - Value lists (P6) merge local entries by full display text, in local order; an entry's id is its position in the
 *   merged unsearched list. A values rule's positions expand back to local ordinals through the same merge, and a
 *   fresh draft builds the local `ValuesFilter`. No committed key run outlives its job: every draft is closed when its
 *   op ends.
 */
import type { CacheIdentity, ColumnRule, FilterRange, FilterSpec, GridSelection, OperationRefusal, RowOrder, SheetView, SortSpec, TypedCondition } from '../../../src/xlsx/view/contracts'
import type { SourceRect } from '../../../src/xlsx/view/coordinates'
import type { XlsxDocument, XlsxSheet } from '../../../src/xlsx/types'
import type { WireDetection, WireRange, WireSelection, WireSortSpec, WireTyped, SheetOrder, ValuesPage } from '../../../src/stream/sheet-wire'
import { createViewerControlResources, viewerControls, type ViewerControlResources, type ViewerControls } from '../../../src/components/xlsx/viewerControlsClient'
import type { ControlResult, ControlSnapshot, DraftQuery, ValuesDraftPort } from '../../../src/components/xlsx/controlsClient'
import { MainMemoryLedger } from '../../../src/xlsx/view/mainLedger'
import { createIdentityOrder } from '../../../src/xlsx/view/order'
import { emptyGridSelection } from '../../../src/xlsx/view/selection'
import { createObjectWorkbookSource } from '../../../src/xlsx/view/source'
import { VALUE_DISPLAY_LIMIT } from '../../../src/xlsx/view/valueIndex'
import { CompactRowVisibility, createRowVisibility, visibilityFromHidden } from '../../../src/xlsx/view/visibility'
import type { WorkbookSource } from '../../../src/worker/viewerSession'
import type { SheetGeometryResult } from '../../../src/xlsx/view/viewport'
import { clockEpochMs, type SheetClock } from './clock'
import { canonicalView, isIdentityView, type CanonicalView, type ResolvedFilter, type ResolvedRule } from './view-spec'
import type { WireSizes } from './view-state'

export type SheetViewJob =
  | { op: 'derive'; sheetIndex: number; view: CanonicalView }
  | { op: 'sort'; sheetIndex: number; base: CanonicalView; sort: WireSortSpec | null }
  | { op: 'filter'; sheetIndex: number; base: CanonicalView; filter: ResolvedFilter | null }
  | { op: 'values'; sheetIndex: number; view: CanonicalView; range: WireRange; col: number; search?: { text: string; wildcards: boolean }; offset: number }
  | { op: 'detect'; sheetIndex: number; kind: 'sort' | 'filter'; selection: WireSelection; override?: { header: 0 | 1 } }
  /** P9: the sheet's extent under `view` at `sizes`, answered by the host from its tile geometry (never by a port). */
  | { op: 'extent'; sheetIndex: number; view: CanonicalView; sizes: WireSizes }
/** P13: a local refusal as it travels: its reason, a source rect for `merge`, `fixed-array`, `dynamic-spill`,
 * `outside-column-data` and `pivot`, and `count`/`at` for `formula-value-unavailable`. Never a `drawingId` or message. */
export interface WireOperationRefusal { reason: string; range?: SourceRect; count?: number; at?: [number, number] }
export type SheetViewReply =
  | { op: 'view'; view: CanonicalView; order: SheetOrder }
  | { op: 'view'; refused: WireOperationRefusal }
  | { op: 'values'; page: Omit<ValuesPage, 'list' | 'cursor'>; next?: number; count: number }
  | { op: 'detect'; detection: WireDetection }
  | { op: 'detect'; refused: WireOperationRefusal }
  | { op: 'extent'; colsCss: number; rowsCss: number }

/** Entries per values answer (P6, `VALUE_LIST_LIMITS.page`). */
export const VALUES_ANSWER_ENTRIES = 500

/** A job the engine could not answer (stale, cancelled, incomplete, a refused replay step): the caller refuses it
 * `policy-refused`. */
export class SheetViewFailed extends Error { constructor(m: string) { super(`sheet-view: ${m}`); this.name = 'SheetViewFailed' } }

type Runs = ReadonlyArray<{ start: number; end: number }>

/** One derived view (P3): the canonical view, the local view state the engine committed, and its wire order. */
export interface DerivedView {
  view: CanonicalView
  order: RowOrder
  visibility: CompactRowVisibility
  sort: SortSpec | null
  filter: FilterSpec | null
  reply: SheetOrder
}
/** The document's derived-view cache, keyed by the canonical view (the host scopes it to `(sheetIndex, clock)`). */
export interface DerivedCache { get(view: CanonicalView): DerivedView | undefined; put(d: DerivedView): void }

export interface SheetViewPort {
  readonly sheetIndex: number
  /** One job at a time: `fn` runs after every earlier job of this port settled. */
  run<T>(fn: () => Promise<T>): Promise<T>
  job(job: SheetViewJob, cache: DerivedCache): Promise<SheetViewReply>
  /** The bytes the port's main ledger charges now (its engine's working set between jobs). */
  ledgerBytes(): number
  dispose(): Promise<void>
}

const QUIET: DraftQuery = { text: '', offset: VALUE_DISPLAY_LIMIT }

const filterRangeOf = (r: WireRange): FilterRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, headerRows: r.header, totalsRows: r.totals })
const sameRange = (a: WireRange, b: WireRange) => a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1 && a.header === b.header && a.totals === b.totals
const sortOf = (s: WireSortSpec): SortSpec => ({ range: filterRangeOf(s.range), keys: s.keys.map((k) => ({ col: k.col, order: k.order })), caseSensitive: s.caseSensitive })
const typedOf = (t: WireTyped): TypedCondition => ({ ...t } as TypedCondition)

/** The zero-size row runs of `geometry`, as `sheet-layout`'s `collapsedRows` reads them. */
export function collapsedRowsOf(geometry: SheetGeometryResult): Runs | undefined {
  if (geometry.kind !== 'ready') return undefined
  return geometry.metrics.rows.overrides.filter((s) => s.sizeCss === 0).map((s) => ({ start: s.start, end: s.end }))
}

/** P13. `incomplete` cannot occur on a resident key view: it fails the job. */
export function wireRefusal(r: OperationRefusal): WireOperationRefusal {
  switch (r.reason) {
    case 'merge': case 'fixed-array': case 'dynamic-spill': case 'outside-column-data': case 'pivot':
      return { reason: r.reason, range: { r0: r.range.r0, r1: r.range.r1, c0: r.range.c0, c1: r.range.c1 } }
    case 'formula-value-unavailable': return { reason: r.reason, count: r.count, at: [r.firstAddress.row, r.firstAddress.col] }
    case 'incomplete': throw new SheetViewFailed('incomplete key view')
    default: return { reason: r.reason }
  }
}

function gridSelection(s: WireSelection): GridSelection {
  return {
    ...emptyGridSelection(),
    active: s.active ? { row: s.active[0], col: s.active[1] } : null,
    anchor: s.anchor ? { row: s.anchor[0], col: s.anchor[1] } : null,
    ranges: s.ranges.map((r) => ({ rows: { kind: 'intervals' as const, intervals: r.rows.map(([r0, r1]) => ({ r0, r1 })) }, cols: r.cols.map(([c0, c1]) => ({ c0, c1 })) })),
  }
}

/** One merged entry: its full display text, local's label, the local ordinals it merges, and whether all are errors. */
interface Merged { text: string; label: string; ordinals: number[]; error: boolean }
interface MergedList { entries: Merged[]; idOf: Map<string, number>; truncated: boolean }

/** Ends the main-ledger lease of a draft page the job has consumed (its entries are copied into the merged list). */
type Settle = (page: object) => void

/** Every displayable local entry of `draft`, merged by full display text in local (first-occurrence) order. */
async function mergedList(draft: ValuesDraftPort, signal: AbortSignal, settle: Settle): Promise<MergedList> {
  const entries: Merged[] = [], idOf = new Map<string, number>()
  let offset = 0, displayed = 0, truncated = false
  do {
    const page = await draft.read({ text: '', offset }, signal)
    displayed = page.displayed; truncated = page.truncated
    for (const e of page.entries) {
      const text = e.labelTruncated ? await draft.inspect(e.ordinal, signal) : e.label
      let id = idOf.get(text)
      if (id === undefined) { id = entries.length; idOf.set(text, id); entries.push({ text, label: e.label, ordinals: [], error: true }) }
      const m = entries[id]!
      m.ordinals.push(e.ordinal)
      if (e.kind !== 'error') m.error = false
    }
    settle(page)
    if (page.entries.length === 0) break
    offset += page.entries.length
  } while (offset < displayed)
  return { entries, idOf, truncated }
}

/** The merged ids `search` matches, ascending. */
async function searchIds(draft: ValuesDraftPort, list: MergedList, search: { text: string; wildcards: boolean }, signal: AbortSignal, settle: Settle): Promise<number[]> {
  const ids = new Set<number>()
  let offset = 0, displayed = 0
  do {
    const page = await draft.read({ text: search.text, wildcards: search.wildcards, offset }, signal)
    displayed = page.displayed
    for (const e of page.entries) {
      const id = list.idOf.get(e.labelTruncated ? await draft.inspect(e.ordinal, signal) : e.label)
      if (id !== undefined) ids.add(id)
    }
    settle(page)
    if (page.entries.length === 0) break
    offset += page.entries.length
  } while (offset < displayed)
  return [...ids].sort((a, b) => a - b)
}

/** Non-blank cells of `col` in body rows `[r0, r1]` of `sheet`, compared with the body's row count. */
function blanksPresent(sheet: XlsxSheet, col: number, r0: number, r1: number): boolean {
  let filled = 0
  for (const row of sheet.rows) {
    if (row.index < r0 || row.index > r1) continue
    for (const cell of row.cells) if (cell.col === col && cell.row === row.index && cell.value !== null && cell.value !== undefined) { filled++; break }
  }
  return filled < r1 - r0 + 1
}

export interface PortInputs {
  sheetIndex: number
  policy: 'visible-only' | 'full'
  clock: SheetClock
  /** `keyDocument(servable workbook, {sheetIndex: key view})`: the key view's sheet in slot `sheetIndex`. */
  doc: XlsxDocument
  /** The snapshot's authored geometry (C1's collapsed rows). */
  geometry: SheetGeometryResult
}

/** A `viewerControls` port over `inputs.doc` (P3). The caller owns the key view and releases it after `dispose`. */
export async function openSheetViewPort(inputs: PortInputs): Promise<SheetViewPort> {
  const { sheetIndex, policy, clock, doc } = inputs
  const collapsed = collapsedRowsOf(inputs.geometry)
  if (!collapsed) throw new SheetViewFailed('sheet geometry unavailable')
  const sheet = doc.sheets[sheetIndex]!
  const docId = `sheet-view|${sheetIndex}|${clock.day}|${clock.timeZone}`
  const source: WorkbookSource = createObjectWorkbookSource(doc, docId)
  const sheetId = source.sheets[sheetIndex]!.sheetId
  const baseHidden = source.geometry(sheetId).baseHidden
  const store = source.store(sheetId)
  const context = () => ({ locale: sheet.semantics?.locale ?? 'en-US', dateSystem: sheet.semantics?.dateSystem ?? '1900', timeZone: clock.timeZone, nowMs: clockEpochMs(clock) })
  let state: SheetView = { sheetId, revision: 0, order: createIdentityOrder(), baseHidden, visibility: visibilityFromHidden(baseHidden), sort: null, filter: null, selection: emptyGridSelection() }
  let viewGen = 0
  const owner = (): CacheIdentity => ({ docId, sheetId, valueGen: store.stamp.valueGen, fontGen: 0, viewGen })
  const snapshot = (): ControlSnapshot => ({ owner: owner(), filter: state.filter, sort: state.sort, records: { visible: 0, total: 0, completeness: 'complete' } })
  /** C1: clears every collapsed body row of `filter` (a display slot through `order`). */
  const c1 = (visibility: CompactRowVisibility, order: RowOrder, filter: FilterSpec): CompactRowVisibility => {
    const lo = filter.range.r0 + filter.range.headerRows, hi = filter.range.r1 - filter.range.totalsRows
    let words: Uint32Array | undefined
    for (const run of collapsed) {
      for (let r = Math.max(lo, run.start); r <= Math.min(hi, run.end); r++) {
        const d = order.sourceToDisplay(r)
        words ??= visibility.words.slice()
        words[d >>> 5] = (words[d >>> 5]! & ~(1 << (d & 31))) >>> 0
      }
    }
    return words ? new CompactRowVisibility(visibility.logicalCount, words) : visibility
  }
  const compact = (v: SheetView['visibility']): CompactRowVisibility => v instanceof CompactRowVisibility ? v : createRowVisibility(1048576, (slot) => v.visible(slot))
  const publish = (next: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>): ControlSnapshot => {
    // P11: a committed view's order and mask are charged once, by the server's `DocumentLedger` hold, never also on
    // this port's ledger past its job; an op that reads them later holds them for its own run.
    ledger.settle(next.order)
    ledger.settle(next.visibility)
    let visibility = compact(next.visibility)
    if (policy === 'visible-only' && next.filter) visibility = c1(visibility, next.order, next.filter)
    state = { ...state, order: next.order, visibility, filter: next.filter, sort: next.sort, revision: state.revision + 1 }
    viewGen++
    return snapshot()
  }
  // One main ledger per port, with no per-port document charge: the server charges documents and held views to its
  // own `DocumentLedger` (Task 4).
  const ledger = new MainMemoryLedger()
  const settle: Settle = (page) => ledger.settle(page)
  const resources: ViewerControlResources = createViewerControlResources('D', undefined, ledger)
  let controls: ViewerControls
  try {
    controls = await viewerControls({ doc, resources, source, sheetId, snapshot, view: () => state, publish, context, ledger }, new AbortController().signal)
  } catch (e) { resources.dispose(); await source.release(); throw e }
  const client = controls.client
  const identity: DerivedView = derived(IDENTITY, state.order, compact(state.visibility), null, null)

  function derived(view: CanonicalView, order: RowOrder, visibility: CompactRowVisibility, sort: SortSpec | null, filter: FilterSpec | null): DerivedView {
    return { view, order, visibility, sort, filter, reply: sheetOrderOf(sheetIndex, view, order, visibility) }
  }
  const setBase = (d: DerivedView) => {
    state = { ...state, order: d.order, visibility: d.visibility, sort: d.sort, filter: d.filter, revision: state.revision + 1 }
    viewGen++
  }
  const capture = (view: CanonicalView): DerivedView => derived(view, state.order, compact(state.visibility), state.sort, state.filter)

  /** A fresh draft over `(range, col)`, closed by the caller. */
  const draftOf = (range: WireRange, col: number, signal: AbortSignal) => client.values(owner(), filterRangeOf(range), col, null, signal)
  /** The local rule of a values rule: its positions through the merge, then a fresh draft's candidate (null: every
   * value and the blanks are selected, which is no rule). The draft joins `drafts`, closed when the op ends. */
  const valuesRule = async (rule: Extract<ResolvedRule, { kind: 'values' }>, drafts: ValuesDraftPort[], signal: AbortSignal): Promise<ColumnRule | null> => {
    const draft = await draftOf(rule.list.range, rule.list.col, signal)
    drafts.push(draft)
    const list = await mergedList(draft, signal, settle)
    const include = rule.mode === 'include'
    settle(await draft.change({ kind: 'all', selected: !include, includeBlanks: false }, QUIET, signal))
    for (const position of rule.positions) {
      const entry = list.entries[position]
      if (!entry) throw new SheetViewFailed('values position past its list')
      for (const ordinal of entry.ordinals) settle(await draft.change({ kind: 'toggle', ordinal, selected: include }, QUIET, signal))
    }
    settle(await draft.change({ kind: 'blanks', selected: rule.blanks }, QUIET, signal))
    return draft.candidate(signal)
  }
  const localFilter = async (f: ResolvedFilter, drafts: ValuesDraftPort[], signal: AbortSignal): Promise<FilterSpec | null> => {
    const columns = new Map<number, ColumnRule>()
    for (const [col, rule] of f.columns) {
      const local = rule.kind === 'values' ? await valuesRule(rule, drafts, signal)
        : rule.kind === 'compound' ? { kind: 'compound' as const, connective: rule.connective, first: typedOf(rule.first), second: typedOf(rule.second) }
        : typedOf(rule)
      if (local) columns.set(col, local)
    }
    return columns.size ? { range: filterRangeOf(f.range), columns } : null
  }
  /** One filter op on the current state; the drafts it opened are closed whatever the outcome. */
  const filterOp = async (f: ResolvedFilter | null, signal: AbortSignal): Promise<ControlResult> => {
    const drafts: ValuesDraftPort[] = []
    try { return await client.filter(owner(), f ? await localFilter(f, drafts, signal) : null, signal) }
    finally { for (const d of drafts) d.cancel() }
  }
  const accepted = (r: ControlResult, step: string) => {
    if (r.kind === 'committed' || r.kind === 'unchanged') return
    throw new SheetViewFailed(r.kind === 'refused' ? `replay ${step} refused: ${r.refusal.reason}` : `replay ${step}: ${r.kind}`)
  }
  /** P3: the derived view of `v`, from the cache or by replay from the identity view. */
  const derive = async (v: CanonicalView, cache: DerivedCache, signal: AbortSignal): Promise<DerivedView> => {
    const c = canonicalView(v)
    if (isIdentityView(c)) return identity
    const hit = cache.get(c)
    if (hit) return hit
    setBase(identity)
    if (c.sort && c.sortFilter) accepted(await filterOp(c.sortFilter, signal), 'sortFilter')
    if (c.sort) accepted(await client.sort(owner(), sortOf(c.sort), signal), 'sort')
    if (c.filter && !(c.sort && JSON.stringify(c.filter) === JSON.stringify(c.sortFilter))) accepted(await filterOp(c.filter, signal), 'filter')
    else if (!c.filter && c.sortFilter) accepted(await filterOp(null, signal), 'clear')
    const d = capture(c)
    cache.put(d)
    return d
  }
  const commitReply = (r: ControlResult, next: CanonicalView, cache: DerivedCache): SheetViewReply => {
    if (r.kind === 'refused') return { op: 'view', refused: wireRefusal(r.refusal) }
    if (r.kind !== 'committed' && r.kind !== 'unchanged') throw new SheetViewFailed(r.kind)
    const d = capture(next)
    cache.put(d)
    return { op: 'view', view: d.view, order: d.reply }
  }

  async function values(job: Extract<SheetViewJob, { op: 'values' }>, signal: AbortSignal): Promise<SheetViewReply> {
    const draft = await draftOf(job.range, job.col, signal)
    try {
      const list = await mergedList(draft, signal, settle)
      const ids = job.search ? await searchIds(draft, list, job.search, signal, settle) : list.entries.map((_, i) => i)
      const filter = canonicalView(job.view).filter
      const rule = filter && sameRange(filter.range, job.range) ? filter.columns.find(([col]) => col === job.col)?.[1] : undefined
      const values = rule?.kind === 'values' && rule.list.col === job.col && sameRange(rule.list.range, job.range) ? rule : undefined
      const positions = new Set(values?.positions ?? [])
      const selected = (id: number) => values ? (values.mode === 'include') === positions.has(id) : true
      const items = ids.slice(job.offset, job.offset + VALUES_ANSWER_ENTRIES).map((id) => ({ id, text: list.entries[id]!.label, selected: selected(id), error: list.entries[id]!.error }))
      const body0 = job.range.r0 + job.range.header, body1 = job.range.r1 - job.range.totals
      const page: Omit<ValuesPage, 'list' | 'cursor'> = {
        items, blanks: { present: blanksPresent(sheet, job.col, body0, body1), selected: values ? values.blanks : true },
        total: ids.length, truncated: list.truncated,
      }
      if (job.offset === 0) {
        const c = client.classify(owner(), filterRangeOf(job.range), job.col, signal)
        if (c.kind !== 'loading') page.classification = { kind: c.kind, hasNumeric: c.hasNumeric }
      }
      const next = job.offset + VALUES_ANSWER_ENTRIES
      return { op: 'values', page, count: list.entries.length, ...(next < ids.length ? { next } : {}) }
    } finally { draft.cancel() }
  }

  async function detect(job: Extract<SheetViewJob, { op: 'detect' }>, signal: AbortSignal): Promise<SheetViewReply> {
    setBase(identity)
    state = { ...state, selection: gridSelection(job.selection) }
    const r = await client.resolve(owner(), job.kind, signal, job.override?.header)
    if (r.kind === 'refused') return { op: 'detect', refused: wireRefusal(r.refusal) }
    if (r.kind !== 'complete') throw new SheetViewFailed(`detect: ${r.kind}`)
    const g = r.range
    return { op: 'detect', detection: { range: { r0: g.r0, r1: g.r1, c0: g.c0, c1: g.c1, header: g.headerRows, totals: g.totalsRows }, source: r.source.kind, header: r.header, overrideIgnored: r.overrideIgnored } }
  }

  let queue: Promise<unknown> = Promise.resolve()
  let disposed = false
  const port: SheetViewPort = {
    sheetIndex,
    run<T>(fn: () => Promise<T>): Promise<T> {
      const next = queue.then(fn, fn)
      queue = next.catch(() => undefined)
      return next
    },
    job(job, cache) {
      return port.run(async () => {
        if (disposed) throw new SheetViewFailed('port disposed')
        const signal = new AbortController().signal
        switch (job.op) {
          case 'derive': {
            const d = await derive(job.view, cache, signal)
            return { op: 'view', view: d.view, order: d.reply }
          }
          case 'sort': {
            const base = await derive(job.base, cache, signal)
            setBase(base)
            const next: CanonicalView = canonicalView(job.sort ? { sort: job.sort, sortFilter: base.view.filter, filter: base.view.filter } : { sort: null, sortFilter: null, filter: base.view.filter })
            return commitReply(await client.sort(owner(), job.sort ? sortOf(job.sort) : null, signal), next, cache)
          }
          case 'filter': {
            const base = await derive(job.base, cache, signal)
            setBase(base)
            const next = canonicalView({ sort: base.view.sort, sortFilter: base.view.sortFilter, filter: job.filter })
            return commitReply(await filterOp(job.filter, signal), next, cache)
          }
          case 'values': return values(job, signal)
          case 'detect': return detect(job, signal)
          case 'extent': throw new SheetViewFailed('extent is answered by the host')
        }
      })
    },
    ledgerBytes: () => ledger.bytes,
    dispose() {
      return port.run(async () => {
        if (disposed) return
        disposed = true
        controls.dispose()
        resources.dispose()
        await source.release()
      })
    },
  }
  return port
}

const IDENTITY: CanonicalView = Object.freeze({ sort: null, sortFilter: null, filter: null })

/** The union row span of the view's ranges, or null for the identity view. */
function spanOf(v: CanonicalView): { r0: number; r1: number } | null {
  const ranges = [v.sort?.range, v.sortFilter?.range, v.filter?.range].filter((r): r is WireRange => r !== undefined)
  if (!ranges.length) return null
  return { r0: Math.min(...ranges.map((r) => r.r0)), r1: Math.max(...ranges.map((r) => r.r1)) }
}

/**
 * P7's order of a derived view. `head.viewGen` is 0: the session stamps its own. `counts` are over the body rows of
 * the filter range (else the sort range): `total` rows, `visible` of them shown. `bands` are the committed
 * `RowOrder.bands`; `mask` covers the union row span of the view's ranges, bit i (LSB-first) = display slot `r0 + i`
 * (every other slot shows as the identity view does). The identity view has no band and no mask.
 */
export function sheetOrderOf(sheetIndex: number, view: CanonicalView, order: RowOrder, visibility: CompactRowVisibility): SheetOrder {
  const body = view.filter?.range ?? view.sort?.range
  let total = 0, visible = 0
  if (body) {
    for (let r = body.r0 + body.header; r <= body.r1 - body.totals; r++) {
      total++
      if (visibility.visible(order.sourceToDisplay(r))) visible++
    }
  }
  const span = spanOf(view)
  let mask: SheetOrder['mask'] = null
  if (span) {
    const words = new Uint32Array(Math.ceil((span.r1 - span.r0 + 1) / 32))
    for (let i = 0; i <= span.r1 - span.r0; i++) if (visibility.visible(span.r0 + i)) words[i >>> 5] = (words[i >>> 5]! | (1 << (i & 31))) >>> 0
    mask = { r0: span.r0, r1: span.r1, words }
  }
  return {
    head: {
      sheetId: String(sheetIndex), viewGen: 0, counts: { visible, total },
      sort: view.sort ? { range: view.sort.range, keys: view.sort.keys } : null,
      filterSummary: view.filter ? view.filter.columns.map(([col, rule]) => ({ col, kind: rule.kind })) : [],
    },
    bands: order.bands.map((b) => ({ r0: b.r0, r1: b.r1, sourceAtOffset: b.sourceAtOffset })),
    mask,
  }
}
