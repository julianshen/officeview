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
import type { StructuralIndex } from '../../xlsx/view/structure'
import { MainControlMemory, controlSourceStore, controlStructure, controlView } from '../../xlsx/view/mainControlMemory'
import { MainMemoryLedger, estimatePayloadBytes } from '../../xlsx/view/mainLedger'
import type { RangeDetection, RangeDetectionRequest } from '../../xlsx/view/rangeDetection'
import { DocumentSession, type OperationOutcome, type WorkbookSource, sameStamp } from '../../worker/viewerSession'
import { DisplayTextLaneCache, needsDisplayText } from '../../worker/displayTextLane'
import { createControlsClient, type ControlSnapshot, type ControlsClient, type StagedControl, type ControlProgress, sameControlOwner } from './controlsClient'

/** onUsage observes usage; pass a ledger as the third factory argument for accounting.
 * Ledger-writing legacy callbacks must migrate rather than mirror that account.
 * Document-owned committed keys survive port eviction. Both indexes and runs
 * share one bounded ledger; indexes are rebuildable, committed keys are not. */
export interface ViewerControlResources { cache: ValueIndexCache; runs: ValueRunStore; readonly memory: MainControlMemory; readonly closed: boolean; readonly cacheOwner: string; adoptCache(targets: readonly object[], bytes: number, commit: () => void): void; bindLedger(ledger: MainMemoryLedger, usage?: 'observer'): void; dispose(): void }
export function createViewerControlResources(profile: ProfileId = 'D', onUsage?: (bytes: number) => void, ledger?: MainMemoryLedger): ViewerControlResources {
  let memory = new MainControlMemory(ledger ?? new MainMemoryLedger(), profile)
  const headerAccount = memory.reserve()
  headerAccount.charge(1024)
  try {
    let account = memory.reserve(), bound = ledger !== undefined, activated = false
    let reportCloseError: ((cause: unknown) => void) | undefined
    const cache = new ValueIndexCache({ limitBytes: (profile === 'M' ? 8 : 16) * 1024 * 1024,
      beforeGrowth: bytes => !closed && account.trySet(bytes),
      adopt: (targets, bytes) => account.adopt(targets, bytes),
      onUsage: bytes => {
        account.trySet(bytes)
        try { onUsage?.(bytes) } catch (cause) { reportCloseError?.(cause); throw cause }
      },
    })
    const store = new MemoryValueRunStore({ limitBytes: 8 * 1024 * 1024, account: cache.reserve(), readerAllocation: {
      get closed() { return memory.closed }, get limitBytes() { return memory.limitBytes }, reserve: () => memory.reserve(), canReserve: bytes => memory.canReserve(bytes), owns: target => memory.owns(target),
    } })
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
    const resource: ViewerControlResources = { cache, runs, get closed() { return closed }, adoptCache: (targets, bytes, commit) => account.adopt(targets, bytes, commit), get cacheOwner() { return account.owner }, get memory() { activated = true; return memory }, bindLedger(ledger, usage) {
      if (memory.ledger === ledger) return
      if (onUsage && usage !== 'observer') throw new RangeError('Declare resourceUsage: observer, or pass the ledger to createViewerControlResources instead of a ledger-writing onUsage callback')
      if (bound || activated || cache.usedBytes) throw new RangeError('Controls resources belong to another main ledger')
      const rebound = new MainControlMemory(ledger, profile), header = rebound.reserve()
      header.charge(1024)
      try { header.adopt([resource, cache, store, refs, runs]); memory.ledger.settle(resource) }
      finally { header.release() }
      account.release()
      memory = rebound
      account = memory.reserve()
      bound = true
    }, dispose() {
      if (reportCloseError) return
      closed = true
      memory.close()
      let failed = false, error: unknown
      const report = (cause: unknown) => { if (!failed) { failed = true; error = cause } }
      reportCloseError = report
      try {
        try { cache.invalidate() } catch (cause) { report(cause) }
        // This store's release performs its refund/notification synchronously,
        // then returns a promise. Report the original observer error from this
        // synchronous close and consume only its rejected cleanup promise.
        for (const ref of refs.values()) void store.release(ref).catch(() => {})
        refs.clear()
      } finally { reportCloseError = undefined }
      if (failed) throw error
    } }
    headerAccount.adopt([resource, cache, store, refs, runs])
    return resource
  } catch (cause) {
    // Failed registration leaves the original reservation intact; successful
    // adoption consumed it. Refund only the still-unadopted factory header.
    try { headerAccount.release() } catch { /* Preserve the construction/adoption error. */ }
    throw cause
  }
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
  /** Display-text lane cache (S13.5; tests, embedders). Default: one per port,
   * charged to the controls caches' main-thread share, closed on dispose. */
  displayLanes?: DisplayTextLaneCache
  /** S10.2: the document's main-thread ledger (structural snapshots are charged while reachable). */
  ledger?: MainMemoryLedger
  /** Explicit migration for callback-only resources: callbacks remain observers. */
  resourceUsage?: 'observer'
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
  const sheet = doc.sheets[source.sheets.findIndex(s => s.sheetId === sheetId)]
  const context = (): FilterContext => ({ locale: sheet.semantics?.locale ?? 'en-US', dateSystem: sheet.semantics?.dateSystem ?? '1900', timeZone: sheet.semantics?.timeZone ?? Intl.DateTimeFormat().resolvedOptions().timeZone, nowMs: Date.now() })
  const resources = options.resources ?? createViewerControlResources('D', undefined, options.ledger)
  if (options.ledger) resources.bindLedger(options.ledger, options.resourceUsage)
  if (resources.closed) throw new ValueIndexError('cancelled', 'Controls resources disposed')
  let disposed = false
  const memory = resources.memory
  if (!options.ledger && !memory.owns(doc)) {
    const ownedDoc = memory.reserve()
    try {
      ownedDoc.charge(estimatePayloadBytes(doc))
      memory.ledger.adoptWhileReachable('controls', ownedDoc.owner, 'document', [doc], ownedDoc.bytes)
    } finally { ownedDoc.release() }
  }
  const store = controlSourceStore(source.store(sheetId), memory)
  const { cache, runs } = resources
  const stamp = { ...store.stamp }
  let structure: StructuralIndex
  try {
    const result = await store.queryStructural({ r0: 0, r1: 1048575, c0: 0, c1: 16383 }, { signal, stamp })
    signal.throwIfAborted()
    structure = controlStructure(result, memory)
    // The frozen index cloned the reply; the index alone carries its charge now.
    store.settle?.(result)
    signal.throwIfAborted()
  } catch (error) { if (!options.resources) resources.dispose(); throw error }
  // Explicit owners of each structural index (AC-LIFE-02): the port while the
  // index is current, and every resolver whose range request handed it out,
  // until that resolver settles. The last holder ends the charge, exactly once.
  const holders = new Map<StructuralIndex, number>()
  const hold = (index: StructuralIndex) => { holders.set(index, (holders.get(index) ?? 0) + 1); return index }
  const letGo = (index: StructuralIndex) => {
    const left = holders.get(index)! - 1
    if (left) holders.set(index, left)
    else { holders.delete(index); memory.settle?.(index) }
  }
  hold(structure)
  const portMetadata = memory.reserve()
  let styles: Map<number, XlsxCell['style']>, classifications: Map<string, ColumnClassification>, jobs: Set<AbortController>
  try {
    // Lookup maps, their style references, and resolver/page/evaluator closures.
    portMetadata.charge(2048)
    styles = new Map<number, XlsxCell['style']>()
    for (const row of sheet.rows) for (const cell of row.cells) {
      if (!styles.has(cell.styleIndex)) portMetadata.charge(48)
      styles.set(cell.styleIndex, cell.style)
    }
    classifications = new Map<string, ColumnClassification>()
    jobs = new Set<AbortController>()
    portMetadata.adopt([styles, classifications, jobs])
  } catch (error) { if (!options.resources) resources.dispose(); throw error }
  finally { portMetadata.release() }
  const styleAt = (index: number) => options.styleOf ? options.styleOf(index) : styles.get(index)
  const scope = (owner: CacheIdentity, range: FilterRange, col: number) => JSON.stringify([owner.valueGen, range, col])
  // Display text comes from an on-demand lane (S13.5), present only for a text
  // rule on a text/mixed column; classification and other rules read raw values.
  const prepared = (page: ValueLanePage, index: number, displayText?: string): PreparedValue => {
    const tag = page.tags[index], value = page.numbers[index], style = styleAt(page.styles?.[index] ?? 0)
    if (tag === 5) return { kind: 'blank' }
    if (tag === 7) return { kind: 'unavailable' }
    if (tag === 4) return { kind: 'error', value: page.texts[index]! }
    if (tag === 2) return { kind: 'string', value: page.texts[index]! }
    return tag === 3 ? { kind: 'boolean', value: Boolean(value), displayText } : { kind: 'number', value, displayText, numFmtId: style?.numFmtId, formatCode: style?.formatCode }
  }
  const display = (page: ValueLanePage, index: number, accept: (text: string) => void): void => {
    const tag = page.tags[index]
    if (tag === 5 || tag === 7 || tag === 4 || tag === 2) return undefined
    const value = page.numbers[index]
    const scratch = laneAccount ? null : memory.reserve()
    try {
      const text = cellDisplayText(sheet, { value: tag === 3 ? Boolean(value) : value, style: styleAt(page.styles?.[index] ?? 0) }, {
        allocate: bytes => { if (scratch) scratch.charge(bytes); else { laneAccount!.charge(bytes); formatCredit += bytes } },
        release: bytes => { if (scratch) scratch.release(bytes); else { laneAccount!.release(bytes); formatCredit -= bytes } },
        retain: (targets, bytes, commit) => {
          if (scratch) scratch.adopt(targets, bytes, commit)
          else { resources.adoptCache(targets, bytes, commit); laneAccount!.release(bytes); formatCredit -= bytes }
        },
      })
      accept(text)
    } finally { scratch?.release(); if (formatCredit) { laneAccount!.release(formatCredit); formatCredit = 0 } }
  }
  // The lane shares the controls caches' main-thread share (§12.4): its bytes
  // are charged to the value-index cache's ledger (16 MiB D / 8 MiB M, with
  // value indexes and committed key runs), which reports into the main
  // ledger's 'controls' pool for exactly as long as the lane is retained. Idle
  // indexes are shed first; then this cache evicts its own unleased lanes.
  const laneAccount = options.displayLanes ? null : cache.reserve()
  // Canonical formatting and sink.add run synchronously. Its admitted credit
  // moves into the lane without charging the same retained string twice.
  let formatCredit = 0
  const lanes = options.displayLanes ?? new DisplayTextLaneCache({
    limitBytes: cache.limitBytes,
    budget: { reserve: bytes => {
      const credit = Math.min(formatCredit, bytes)
      formatCredit -= credit
      laneAccount!.charge(bytes - credit)
    }, release: bytes => laneAccount!.release(bytes) },
  })
  // The style resolver is this port's style owner: lanes key by its identity.
  const styleOwner: object = options.styleOf ?? styles
  const pages = async (owner: CacheIdentity, range: FilterRange, col: number, abort: AbortSignal, accept: (page: ValueLanePage, firstRow: number) => Promise<void> | void) => {
    const r0 = range.r0 + range.headerRows, r1 = range.r1 - range.totalsRows, size = 4096
    const laneSource = valueSourceFromStore(store, col, { r0, r1 }, size, memory)
    for (let index = 0; index < laneSource.pageCount; index++) {
      abort.throwIfAborted()
      let complete = false
      for (;;) {
        if (disposed || resources.closed) throw new ValueIndexError('cancelled', 'Controls disposed')
        const page = await laneSource.readPage(index, abort)
        try {
        if (disposed || resources.closed) throw new ValueIndexError('cancelled', 'Controls disposed')
        if (!sameControlOwner(owner, options.snapshot().owner) || !sameStamp(owner, store.stamp)) throw new ValueIndexError('stale-source', 'Source or view changed; retry.')
        if (page.kind === 'error') throw new ValueIndexError(page.failure.code, page.failure.message)
        // The consumer folds the page into its own (charged) state, then the
        // reply it came from stops being charged.
        if (page.kind === 'values') { await accept(page, r0 + index * size); break }
        if (complete) throw new ValueIndexError('page-error', 'Values stayed unknown after source completion.')
        const coverage = await store.ensureComplete({ r0: r0 + index * size, r1: Math.min(r1, r0 + (index + 1) * size - 1), c0: col, c1: col }, { stamp: owner, signal: abort })
        if (coverage.error) throw new ValueIndexError(coverage.error.code, coverage.error.message)
        complete = coverage.completeness === 'complete'
        if (!complete) throw new ValueIndexError('incomplete', 'Source values are still loading; retry.')
        } finally { if (page.kind === 'values') page.settle?.() }
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
  const stageOwned = async (kind: 'filter' | 'sort', spec: Parameters<ControlsClient['filter']>[1] | Parameters<ControlsClient['sort']>[1], owner: CacheIdentity, abort: AbortSignal, progress?: (value: ControlProgress) => void): Promise<StagedControl> => {
    if (disposed || resources.closed) return { kind: 'cancelled' }
    if (options.remote) return stageRemote(options.remote, kind, spec, owner, abort, progress)
    const before = options.view(), filterContext = context()
    // This operation reads the displayed view until it ends: it holds it too.
    const releaseBefore = before.visibility instanceof CompactRowVisibility ? controlView(memory, before.order, before.visibility, true) : () => {}
    // Borrow the source: closing a private draft releases only its own caches.
    const borrowed: WorkbookSource = { ...source, store: id => controlSourceStore(source.store(id), memory), bindMemory: undefined, release: async () => {} }
    const flagsAccount = memory.reserve(), sessionMetadata = memory.reserve()
    // Session/pool descriptors, their empty maps, and bounded draft/guard records.
    // The extra 512 B conservatively covers the state, helper closures and
    // third weak adoption target; admit it before creating those records.
    // A refusal here releases the hold this operation took on its view.
    try { sessionMetadata.charge(4096 + 512) } catch (error) { releaseBefore(); throw error }
    const state: { producer: 'new' | 'running' | 'done'; consumerDone: boolean; drained: boolean; flags: Uint8Array | undefined } = {
      producer: 'new', consumerDone: false, drained: false, flags: undefined,
    }
    const drain = () => {
      if (state.drained || !state.consumerDone || state.producer === 'running') return
      state.drained = true
      state.flags = undefined
      flagsAccount.release()
    }
    // Keep the retained handle outside the evaluator's cooperative request scope.
    const createMatcher = (flags: typeof state, offset: number) => (row: number) => {
      if (!flags.flags) throw new ValueIndexError('cancelled', 'Evaluator matcher released')
      return Boolean(flags.flags[row - offset])
    }
    let engine: DocumentSession
    try { engine = new DocumentSession(owner.docId, borrowed, { profile: memory.profile, workerBudgetBytes: memory.limitBytes, allocation: memory, filterEvaluator: async ({ spec: filter, signal: request, cooperate }) => {
      if (state.producer !== 'new' || state.consumerDone || state.drained || request.aborted || disposed || resources.closed) throw new ValueIndexError('cancelled', 'Evaluator closed')
      state.producer = 'running'
      try {
        flagsAccount.charge(filter.range.r1 - filter.range.r0 + 1 + 128)
        state.flags = new Uint8Array(filter.range.r1 - filter.range.r0 + 1).fill(1)
        const r0 = filter.range.r0 + filter.range.headerRows, r1 = filter.range.r1 - filter.range.totalsRows
        for (const [col, rule] of filter.columns) {
          await classify(owner, filter.range, col, request)
          const domain = identity(owner, filter.range, col)
          const membership = rule.kind === 'values' ? await compileValuesMembership(rule, domain, runs, { cache, signal: request }) : undefined
          const match = membership?.matches ?? new FilterEngine(filterContext).compileRule(rule, domain.classification, { signal: request })
          let lease
          try {
            lease = needsDisplayText(rule, domain.classification)
              ? await lanes.lease({ docId: owner.docId, sheetId: owner.sheetId, valueGen: owner.valueGen, styleOwner, col, r0, r1 }, sink => pages(owner, filter.range, col, request, async (page, first) => {
                for (let i = 0; i < page.count; i++) {
                  request.throwIfAborted()
                  display(page, i, text => sink.add(first + i - r0, text))
                  if (cooperate?.due(1)) await cooperate.pause('filter')
                }
              }))
              : undefined
          } catch (error) {
            membership?.release()
            if (error instanceof ValueIndexError && error.code === 'quota') return { kind: 'refused', refusal: { reason: 'resource-limit', message: error.message } }
            throw error
          }
          try {
            await pages(owner, filter.range, col, request, async (page, first) => {
              for (let i = 0; i < page.count; i++) {
                request.throwIfAborted()
                if (!match(prepared(page, i, lease?.lane.text(first + i - r0)))) state.flags![first + i - filter.range.r0] = 0
                if (cooperate?.due(1)) await cooperate.pause('filter')
              }
            })
          } finally { lease?.release(); membership?.release() }
        }
        if (state.consumerDone || request.aborted || disposed || resources.closed) throw new ValueIndexError('cancelled', 'Evaluator closed')
        return { kind: 'complete', matches: createMatcher(state, filter.range.r0) }
      } finally { state.producer = 'done'; drain() }
    } }) } catch (error) { state.consumerDone = true; drain(); sessionMetadata.release(); releaseBefore(); throw error }
    let primaryFailure = false, primaryError: unknown
    try {
      if (store.stamp.valueGen > engine.identity(sheetId).valueGen) engine.publishValues(sheetId)
      let visibility: CompactRowVisibility
      if (before.visibility instanceof CompactRowVisibility) visibility = before.visibility
      else {
        const converted = memory.reserve()
        try {
          converted.charge(2 * Math.ceil(1048576 / 32) * 4 + 4 * (Math.ceil(1048576 / 1024) + 1))
          visibility = createRowVisibility(1048576, slot => before.visibility.visible(slot))
          converted.adopt([visibility, visibility.words, visibility.words.buffer, visibility.summaries, visibility.summaries.buffer], visibility.words.byteLength + visibility.summaries.byteLength, undefined, true)
        } finally { converted.release() }
      }
      engine.commitDraft(engine.identity(sheetId), { ...before, visibility, matches: row => before.visibility.visible(before.order.sourceToDisplay(row)) })
      const hooks = { progress: (completed: number, total: number) => progress?.({ completed, total, message: kind === 'sort' ? 'Sorting view' : 'Filtering view' }) }
      const outcome = kind === 'sort' ? await engine.sort(engine.identity(sheetId), spec as Parameters<ControlsClient['sort']>[1], abort, hooks)
        : await engine.filter(engine.identity(sheetId), spec as Parameters<ControlsClient['filter']>[1], Date.now(), abort, hooks)
      if (disposed || resources.closed || abort.aborted) return { kind: 'cancelled' }
      if (outcome.kind !== 'committed') return failure(outcome)
      const draft = outcome.view
      return { kind: 'ready', publish: () => disposed || resources.closed ? null : options.publish(draft) }
    } catch (error) {
      // When the operation has been aborted or the client disposed/closed, any concurrent
      // failure during in-flight teardown is treated as cancelled to avoid noisy unwinds
      // on torn-down instances.
      if (abort.aborted || disposed || resources.closed) return { kind: 'cancelled' }
      if (error instanceof ValueIndexError && error.code === 'quota') return { kind: 'refused', refusal: { reason: 'resource-limit', message: error.message } }
      primaryFailure = true
      primaryError = error
      throw error
    } finally {
      let cleanupFailed = false, cleanupError: unknown
      try { await engine.release() }
      catch (error) { cleanupFailed = true; cleanupError = error }
      releaseBefore()
      try {
        state.consumerDone = true
        drain()
        // Atomic adoption keeps the explicit reservation on registration failure.
        // If adopt throws, sessionMetadata.release() is skipped intentionally: this permanent
        // conservative over-charge guarantees memory is never under-accounted if target registration fails.
        sessionMetadata.adopt([engine, engine.memory, state])
        sessionMetadata.release()
      } catch (error) {
        if (!cleanupFailed) { cleanupFailed = true; cleanupError = error }
      }
      // Cleanup still completes; surface any secondary cleanup error on the primary failure without replacing it.
      if (cleanupFailed) {
        if (primaryFailure) {
          // If primaryError is a primitive or already defines `suppressed`, the secondary cleanup
          // error is intentionally dropped here to preserve the primary failure semantics and avoid
          // clobbering an existing suppressed error chain.
          if (typeof primaryError === 'object' && primaryError !== null && !('suppressed' in primaryError)) {
            try { (primaryError as { suppressed?: unknown }).suppressed = cleanupError } catch {}
          }
        } else {
          throw cleanupError
        }
      }
    }
  }
  const stage: typeof stageOwned = async (kind, spec, owner, outer, progress) => {
    if (disposed || resources.closed) return { kind: 'cancelled' }
    const metadata = memory.reserve()
    metadata.charge(256)
    const controller = new AbortController(), forward = () => controller.abort()
    if (outer.aborted) controller.abort()
    outer.addEventListener('abort', forward, { once: true })
    jobs.add(controller)
    try { return await stageOwned(kind, spec, owner, controller.signal, progress) }
    finally {
      outer.removeEventListener('abort', forward)
      jobs.delete(controller)
      metadata.adopt([controller, controller.signal])
      metadata.release()
    }
  }
  const client = createControlsClient({ allocation: memory, snapshot: options.snapshot, context, cache, runs,
    rangeRequest: col => {
      const selection = options.view().selection
      const active = selection.active ? { ...selection.active, col: col ?? selection.active.col } : { row: store.extent.used?.r0 ?? 0, col: col ?? 0 }
      // A header gesture moves the active cell to its column. A leftover
      // single-cell selection is no explicit range (S7.1), and kept here it
      // would refuse as a selection that does not contain the active cell.
      const ranges = col !== undefined && isSingleCell(selection) ? [] : selection.ranges
      return { store, structure, selection: { ...selection, active, ranges } }
    },
    ...(isRemoteStore(store) ? { detectRange: (kind: 'filter' | 'sort', request: RangeDetectionRequest) => store.detectRange(kind, request.selection, request.override ?? null, request.options) } : {}),
    domain: (owner, range, col) => ({ identity: identity(owner, range, col), source: valueSourceFromStore(store, col, { r0: range.r0 + range.headerRows, r1: range.r1 - range.totalsRows }, 4096, memory), label: (kind, value, styleIndex, account) => cellDisplayText(sheet, { value: kind === 'boolean' ? Boolean(value) : value, style: styleAt(styleIndex) }, account ? {
      allocate: bytes => account.charge(bytes), release: bytes => account.release(bytes),
      retain: (targets, bytes, commit) => {
        const held = account as import('../../xlsx/view/allocationAccount').AllocationReservation
        if (held.adopt) held.adopt(targets, bytes, commit)
        else commit()
      },
    } : undefined) }),
    stageFilter: (spec, owner, abort, progress) => stage('filter', spec, owner, abort, progress), stageSort: (spec, owner, abort, progress) => stage('sort', spec, owner, abort, progress),
  })
  const resolve = client.resolve, values = client.values
  client.resolve = async (owner, kind, abort, override, col) => {
    if (disposed || resources.closed) throw new ValueIndexError('cancelled', 'Controls disposed')
    // A worker-backed sheet opens while its import still runs, so the snapshot
    // taken above can be 'loading'. Re-query until it is complete (or failed);
    // a stuck snapshot would refuse every later sort/filter as incomplete.
    if (structure.completeness !== 'complete' && structure.failure === undefined) {
      const result = await store.queryStructural({ r0: 0, r1: 1048575, c0: 0, c1: 16383 }, { signal: abort, stamp: { ...store.stamp } })
      abort.throwIfAborted()
      if (disposed || resources.closed) throw new ValueIndexError('cancelled', 'Controls disposed')
      const stale = structure
      structure = hold(controlStructure(result, memory))
      store.settle?.(result)
      letGo(stale)
        abort.throwIfAborted()
    }
    // The inner resolve reads `structure` synchronously through rangeRequest
    // and may use it again after awaiting header rows.
    const reading = hold(structure)
    let resolved: Awaited<ReturnType<typeof resolve>>
    try { resolved = await resolve(owner, kind, abort, override, col) } finally { letGo(reading) }
    if (resolved.kind === 'complete' && kind === 'filter') await classify(owner, resolved.range, col ?? resolved.range.c0, abort)
    return resolved
  }
  client.values = async (owner, range, col, applied, abort) => {
    if (disposed || resources.closed) throw new ValueIndexError('cancelled', 'Controls disposed')
    await classify(owner, range, col, abort)
    if (disposed || resources.closed) throw new ValueIndexError('cancelled', 'Controls disposed')
    return values(owner, range, col, applied, abort)
  }
  return { client, dispose() {
    const first = !disposed
    disposed = true
    for (const job of jobs) job.abort()
    classifications.clear()
    if (options.displayLanes) lanes.invalidate(key => key.sheetId === sheetId)
    // A lane still read by an in-flight filter keeps its controls charge until
    // that reader exits; the account closes once the cache is drained.
    else lanes.close(() => laneAccount?.close())
    cache.invalidate(identity => identity.sheetId === sheetId)
    // Explicit owners (AC-LIFE-02): a disposed port's lookup metadata ends now,
    // and its structure as soon as no in-flight resolver still reads it;
    // closures a host still retains must not keep them charged.
    memory.settle?.(styles)
    if (first) letGo(structure)
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

/** A worker-backed sheet store that can run range detection next to its pages. */
function isRemoteStore(store: unknown): store is { detectRange(kind: 'filter' | 'sort', selection: RangeDetectionRequest['selection'], override: RangeDetectionRequest['override'] | null, options: RangeDetectionRequest['options']): Promise<RangeDetection> } {
  return typeof (store as { detectRange?: unknown }).detectRange === 'function'
}
