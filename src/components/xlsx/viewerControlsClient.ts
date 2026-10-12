/** In-process S9 transport: stage in a private, budgeted session and publish the
 * resulting order/mask together. The uploaded document and live view stay untouched
 * while a dialog is pending. Worker transports use the same ControlsClient port. */
import type { CacheIdentity, FilterKeySet, FilterRange, FilterSpec, OperationRefusal, PagedKeySetRef, ProfileId, SheetView, SortSpec } from '../../xlsx/view/contracts'
import type { XlsxCell, XlsxDocument } from '../../xlsx/types'
import { cellDisplayText } from '../../xlsx/render'
import { classifyColumn, FilterEngine, type FilterContext, type PreparedValue, type ColumnClassification } from '../../xlsx/view/filter'
import { compileValuesMembership, ValueIndexCache, valueScope, valueSourceFromStore, type ValueIndexIdentity, type ValueLanePage, ValueIndexError } from '../../xlsx/view/valueIndex'
import { CompactRowVisibility, createRowVisibility } from '../../xlsx/view/visibility'
import { MemoryValueRunStore, VALUE_RUN_PAGE, type ValueRunStore } from '../../xlsx/view/valueRuns'
import { StructuralIndex } from '../../xlsx/view/structure'
import { DocumentSession, type OperationOutcome, type WorkbookSource, sameStamp } from '../../worker/viewerSession'
import { createControlsClient, type ControlSnapshot, type ControlsClient, type StagedControl, type ControlProgress, sameControlOwner } from './controlsClient'

/** Document-owned committed keys survive port eviction. Both indexes and runs
 * share one bounded ledger; indexes are rebuildable, committed keys are not. */
export interface ViewerControlResources { cache: ValueIndexCache; runs: ValueRunStore; dispose(): void }
export function createViewerControlResources(profile: ProfileId = 'D'): ViewerControlResources {
  const cache = new ValueIndexCache({ limitBytes: (profile === 'M' ? 8 : 16) * 1024 * 1024 })
  const store = new MemoryValueRunStore({ limitBytes: 8 * 1024 * 1024, account: cache.reserve() })
  const refs = new Map<string, PagedKeySetRef>()
  let closed = false
  const runs: ValueRunStore = {
    async put(...args) {
      if (closed) throw new ValueIndexError('cancelled', 'Controls disposed.')
      const ref = await store.put(...args)
      if (closed) { await store.release(ref); throw new ValueIndexError('cancelled', 'Controls disposed.') }
      refs.set(ref.refId, ref)
      return ref
    },
    read: (...args) => store.read(...args),
    async release(ref) { refs.delete(ref.refId); await store.release(ref) },
  }
  return { cache, runs, dispose() {
    closed = true
    cache.invalidate()
    for (const ref of refs.values()) void store.release(ref)
    refs.clear()
  } }
}

export interface ViewerControlsOptions {
  doc: XlsxDocument
  resources?: ViewerControlResources
  source: WorkbookSource
  sheetId: string
  snapshot(): ControlSnapshot
  view(): SheetView
  publish(view: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>): ControlSnapshot
  /** Worker-owned sheet views (S5.4): sort/filter commit in the resident worker,
   * which then publishes its view to this side; nothing is staged here. */
  remote?: RemoteViewPort
  /** Resolved style of a style id when the document carries no cells (worker-backed sheets). */
  styleOf?: (index: number) => XlsxCell['style']
}
export type RemoteViewCommit =
  | { kind: 'committed'; view: Pick<SheetView, 'order' | 'filter' | 'sort'> & { visibility?: SheetView['visibility'] } }
  | { kind: 'refused'; refusal: OperationRefusal }
  | { kind: 'stale' | 'cancelled' | 'incomplete' }
  | { kind: 'error'; message: string }
export interface RemoteViewPort {
  sort(spec: SortSpec | null, signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteViewCommit>
  filter(spec: FilterSpec | null, keySets: readonly FilterKeySet[], signal: AbortSignal, progress?: (completed: number, total: number) => void): Promise<RemoteViewCommit>
}
export interface ViewerControls { client: ControlsClient; dispose(): void }
export async function viewerControls(options: ViewerControlsOptions, signal: AbortSignal): Promise<ViewerControls> {
  const { doc, source, sheetId } = options
  const sheet = doc.sheets[source.sheets.findIndex(s => s.sheetId === sheetId)], store = source.store(sheetId)
  const context = (): FilterContext => ({ locale: sheet.semantics?.locale ?? 'en-US', dateSystem: sheet.semantics?.dateSystem ?? '1900', timeZone: sheet.semantics?.timeZone ?? Intl.DateTimeFormat().resolvedOptions().timeZone, nowMs: Date.now() })
  const resources = options.resources ?? createViewerControlResources()
  const { cache, runs } = resources
  const stamp = { ...store.stamp }
  let structure: StructuralIndex
  try {
    structure = new StructuralIndex(await store.queryStructural({ r0: 0, r1: 1048575, c0: 0, c1: 16383 }, { signal, stamp }))
    signal.throwIfAborted()
  } catch (error) { if (!options.resources) resources.dispose(); throw error }
  const styles = new Map<number, XlsxCell['style']>()
  for (const row of sheet.rows) for (const cell of row.cells) styles.set(cell.styleIndex, cell.style)
  const styleAt = (index: number) => options.styleOf ? options.styleOf(index) : styles.get(index)
  const classifications = new Map<string, ColumnClassification>()
  const scope = (owner: CacheIdentity, range: FilterRange, col: number) => JSON.stringify([owner.valueGen, range, col])
  const prepared = (page: ValueLanePage, index: number): PreparedValue => {
    const tag = page.tags[index], value = page.numbers[index], style = styleAt(page.styles?.[index] ?? 0)
    if (tag === 5) return { kind: 'blank' }
    if (tag === 7) return { kind: 'unavailable' }
    if (tag === 4) return { kind: 'error', value: page.texts[index]! }
    if (tag === 2) return { kind: 'string', value: page.texts[index]! }
    const displayText = cellDisplayText(sheet, { value: tag === 3 ? Boolean(value) : value, style })
    return tag === 3 ? { kind: 'boolean', value: Boolean(value), displayText } : { kind: 'number', value, displayText, numFmtId: style?.numFmtId, formatCode: style?.formatCode }
  }
  const pages = async (owner: CacheIdentity, range: FilterRange, col: number, abort: AbortSignal, accept: (page: ValueLanePage, firstRow: number) => Promise<void> | void) => {
    const r0 = range.r0 + range.headerRows, r1 = range.r1 - range.totalsRows, size = 4096
    const laneSource = valueSourceFromStore(store, col, { r0, r1 }, size)
    for (let index = 0; index < laneSource.pageCount; index++) {
      abort.throwIfAborted()
      let complete = false
      for (;;) {
        const page = await laneSource.readPage(index, abort)
        if (!sameControlOwner(owner, options.snapshot().owner) || !sameStamp(owner, store.stamp)) throw new ValueIndexError('stale-source', 'Source or view changed; retry.')
        if (page.kind === 'error') throw new ValueIndexError(page.failure.code, page.failure.message)
        if (page.kind === 'values') { await accept(page, r0 + index * size); break }
        if (complete) throw new ValueIndexError('page-error', 'Values stayed unknown after source completion.')
        const coverage = await store.ensureComplete({ r0: r0 + index * size, r1: Math.min(r1, r0 + (index + 1) * size - 1), c0: col, c1: col }, { stamp: owner, signal: abort })
        if (coverage.error) throw new ValueIndexError(coverage.error.code, coverage.error.message)
        complete = coverage.completeness === 'complete'
        if (!complete) throw new ValueIndexError('incomplete', 'Source values are still loading; retry.')
      }
    }
  }
  const classify = async (owner: CacheIdentity, range: FilterRange, col: number, abort: AbortSignal) => {
    const key = scope(owner, range, col)
    if (classifications.has(key)) return
    let numeric = false, text = false, date = false, loading = false
    await pages(owner, range, col, abort, page => {
      function* values() { for (let i = 0; i < page.count; i++) yield prepared(page, i) }
      const part = classifyColumn(values(), true)
      numeric ||= part.hasNumeric; text ||= part.kind === 'text'; date ||= part.kind === 'date'; loading ||= part.kind === 'loading'
    })
    if (classifications.size >= 64) classifications.delete(classifications.keys().next().value!)
    classifications.set(key, { kind: loading ? 'loading' : text ? 'text' : numeric ? date ? 'date' : 'number' : 'empty', hasNumeric: numeric })
  }
  const identity = (owner: CacheIdentity, range: FilterRange, col: number): ValueIndexIdentity => ({ ...store.stamp, range, col, classification: classifications.get(scope(owner, range, col)) ?? { kind: 'loading', hasNumeric: false }, classificationGen: owner.valueGen })
  /** Committed value-list keys of a filter, read from the main-owned key store
   * for shipment with a worker filter command (bounded by that store). */
  const shipKeys = async (owner: CacheIdentity, filter: FilterSpec, abort: AbortSignal): Promise<FilterKeySet[]> => {
    const out: FilterKeySet[] = []
    for (const [col, rule] of filter.columns) {
      if (rule.kind !== 'values') continue
      const ref = rule.selection.mode === 'include' ? rule.selection.keys : rule.selection.mode === 'allExcept' ? rule.selection.excluded : null
      if (!ref) continue
      await classify(owner, filter.range, col, abort)
      const scope = valueScope(identity(owner, filter.range, col))
      const kinds = new Uint8Array(ref.count), numbers = new Float64Array(ref.count), texts: (string | undefined)[] = []
      for (let offset = 0; offset < ref.count; offset += VALUE_RUN_PAGE) {
        const run = await runs.read(ref, offset, Math.min(VALUE_RUN_PAGE, ref.count - offset), scope, abort)
        kinds.set(run.kinds.subarray(0, run.count), offset)
        numbers.set(run.numbers.subarray(0, run.count), offset)
        for (let i = 0; i < run.count; i++) texts.push(run.texts[i])
      }
      out.push({ refId: ref.refId, scope, run: { count: ref.count, kinds, numbers, texts } })
    }
    return out
  }
  const stageRemote = async (remote: RemoteViewPort, kind: 'filter' | 'sort', spec: FilterSpec | SortSpec | null, owner: CacheIdentity, abort: AbortSignal, progress?: (value: ControlProgress) => void): Promise<StagedControl> => {
    const report = (completed: number, total: number) => progress?.({ completed, total, message: kind === 'sort' ? 'Sorting view' : 'Filtering view' })
    const commit = kind === 'sort' ? await remote.sort(spec as SortSpec | null, abort, report)
      : await remote.filter(spec as FilterSpec | null, spec ? await shipKeys(owner, spec as FilterSpec, abort) : [], abort, report)
    switch (commit.kind) {
      case 'committed': {
        const view = commit.view
        return { kind: 'ready', committed: true, publish: () => options.publish({ order: view.order, visibility: view.visibility ?? options.view().visibility, filter: view.filter, sort: view.sort }) }
      }
      case 'error': return { kind: 'refused', refusal: { reason: 'source-error', code: 'page-error', retryable: true, message: commit.message } }
      default: return commit
    }
  }
  const stage = async (kind: 'filter' | 'sort', spec: Parameters<ControlsClient['filter']>[1] | Parameters<ControlsClient['sort']>[1], owner: CacheIdentity, abort: AbortSignal, progress?: (value: ControlProgress) => void): Promise<StagedControl> => {
    if (options.remote) return stageRemote(options.remote, kind, spec, owner, abort, progress)
    const before = options.view(), filterContext = context()
    // Borrow the source: closing a private draft releases only its own caches.
    const borrowed: WorkbookSource = { ...source, bindMemory: undefined, release: async () => {} }
    const engine = new DocumentSession(owner.docId, borrowed, { filterEvaluator: async ({ spec: filter, signal: request, cooperate }) => {
      const flags = new Uint8Array(filter.range.r1 - filter.range.r0 + 1).fill(1)
      for (const [col, rule] of filter.columns) {
        await classify(owner, filter.range, col, request)
        const domain = identity(owner, filter.range, col)
        const match = rule.kind === 'values' ? (await compileValuesMembership(rule, domain, runs, { signal: request })).matches : new FilterEngine(filterContext).compileRule(rule, domain.classification, { signal: request })
        await pages(owner, filter.range, col, request, async (page, first) => {
          for (let i = 0; i < page.count; i++) {
            request.throwIfAborted()
            if (!match(prepared(page, i))) flags[first + i - filter.range.r0] = 0
            if (cooperate?.due(1)) await cooperate.pause('filter')
          }
        })
      }
      return { kind: 'complete', matches: row => Boolean(flags[row - filter.range.r0]) }
    } })
    try {
      if (store.stamp.valueGen > engine.identity(sheetId).valueGen) engine.publishValues(sheetId)
      const visibility = before.visibility instanceof CompactRowVisibility ? before.visibility : createRowVisibility(1048576, slot => before.visibility.visible(slot))
      engine.commitDraft(engine.identity(sheetId), { ...before, visibility, matches: row => before.visibility.visible(before.order.sourceToDisplay(row)) })
      const hooks = { progress: (completed: number, total: number) => progress?.({ completed, total, message: kind === 'sort' ? 'Sorting view' : 'Filtering view' }) }
      const outcome = kind === 'sort' ? await engine.sort(engine.identity(sheetId), spec as Parameters<ControlsClient['sort']>[1], abort, hooks)
        : await engine.filter(engine.identity(sheetId), spec as Parameters<ControlsClient['filter']>[1], Date.now(), abort, hooks)
      if (outcome.kind !== 'committed') return failure(outcome)
      const draft = outcome.view
      return { kind: 'ready', publish: () => options.publish(draft) }
    } finally { await engine.release() }
  }
  const client = createControlsClient({ snapshot: options.snapshot, context, cache, runs,
    rangeRequest: col => {
      const selection = options.view().selection
      const active = selection.active ? { ...selection.active, col: col ?? selection.active.col } : { row: store.extent.used?.r0 ?? 0, col: col ?? 0 }
      // A header gesture moves the active cell to its column. A leftover
      // single-cell selection is no explicit range (S7.1), and kept here it
      // would refuse as a selection that does not contain the active cell.
      const ranges = col !== undefined && isSingleCell(selection) ? [] : selection.ranges
      return { store, structure, selection: { ...selection, active, ranges } }
    },
    domain: (owner, range, col) => ({ identity: identity(owner, range, col), source: valueSourceFromStore(store, col, { r0: range.r0 + range.headerRows, r1: range.r1 - range.totalsRows }), label: (kind, value, styleIndex) => cellDisplayText(sheet, { value: kind === 'boolean' ? Boolean(value) : value, style: styleAt(styleIndex) }) }),
    stageFilter: (spec, owner, abort, progress) => stage('filter', spec, owner, abort, progress), stageSort: (spec, owner, abort, progress) => stage('sort', spec, owner, abort, progress),
  })
  const resolve = client.resolve, values = client.values
  client.resolve = async (owner, kind, abort, override, col) => {
    // A worker-backed sheet opens while its import still runs, so the snapshot
    // taken above can be 'loading'. Re-query until it is complete (or failed);
    // a stuck snapshot would refuse every later sort/filter as incomplete.
    if (structure.completeness !== 'complete' && structure.failure === undefined) {
      structure = new StructuralIndex(await store.queryStructural({ r0: 0, r1: 1048575, c0: 0, c1: 16383 }, { signal: abort, stamp: { ...store.stamp } }))
      abort.throwIfAborted()
    }
    const resolved = await resolve(owner, kind, abort, override, col)
    if (resolved.kind === 'complete' && kind === 'filter') await classify(owner, resolved.range, col ?? resolved.range.c0, abort)
    return resolved
  }
  client.values = async (owner, range, col, applied, abort) => { await classify(owner, range, col, abort); return values(owner, range, col, applied, abort) }
  return { client, dispose() {
    classifications.clear()
    cache.invalidate(identity => identity.sheetId === sheetId)
    if (!options.resources) resources.dispose()
  } }
}
function isSingleCell(selection: SheetView['selection']): boolean {
  if (selection.ranges.length !== 1) return false
  const { rows, cols } = selection.ranges[0]
  const oneRow = rows.kind === 'intervals' ? rows.intervals.length === 1 && rows.intervals[0].r0 === rows.intervals[0].r1 : rows.kind === 'list' ? rows.rows.length === 1 : false
  return oneRow && cols.length === 1 && cols[0].c0 === cols[0].c1
}
function failure(outcome: Exclude<OperationOutcome, { kind: 'committed' }>): StagedControl {
  if (outcome.kind === 'cancelled' || outcome.kind === 'stale' || outcome.kind === 'incomplete' || outcome.kind === 'refused') return outcome
  return outcome.kind === 'error' ? { kind: 'refused', refusal: { reason: 'source-error', ...outcome.error } } : { kind: 'refused', refusal: { reason: 'resource-limit', message: 'View exceeds the session memory budget.' } }
}
