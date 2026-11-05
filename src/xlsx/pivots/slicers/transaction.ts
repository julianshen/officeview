/** Workbook-scoped staging primitives for linked slicer/timeline changes.
 * This module is lazy with the rest of `xlsx/pivots`; it is not imported by the eager graph. */

import { MAX_LINKED_CONTROL_ITEMS, MAX_LINKED_CONTROL_TARGETS } from '../../view/contracts'
export { MAX_LINKED_CONTROL_ITEMS, MAX_LINKED_CONTROL_TARGETS } from '../../view/contracts'

export type LinkedControlOperation =
  | { kind: 'slicerSelection'; name: string; selectedItemIds: readonly number[] }
  | { kind: 'slicerClear'; name: string }
  | { kind: 'timelineSelection'; name: string; start: number; end: number }
  | { kind: 'timelineClear'; name: string }
  | { kind: 'tableSlicerSelection'; name: string; selectedItemIds: readonly number[] }

export function validateLinkedControlOperation(value: unknown): value is LinkedControlOperation {
  if (typeof value !== 'object' || value === null) return false
  const operation = value as Record<string, unknown>
  if (typeof operation.name !== 'string' || operation.name.length < 1 || operation.name.length > 256) return false
  const exactKeys = (keys: readonly string[]) => Object.keys(operation).length === keys.length && keys.every(key => Object.prototype.hasOwnProperty.call(operation, key))
  if (operation.kind === 'slicerClear' || operation.kind === 'timelineClear') return exactKeys(['kind', 'name'])
  if (operation.kind === 'timelineSelection') return exactKeys(['kind', 'name', 'start', 'end']) && Number.isSafeInteger(operation.start) && Number.isSafeInteger(operation.end)
    && Number(operation.start) <= Number(operation.end)
  if (operation.kind !== 'slicerSelection' && operation.kind !== 'tableSlicerSelection') return false
  if (!exactKeys(['kind', 'name', 'selectedItemIds'])) return false
  const ids = operation.selectedItemIds
  if (!Array.isArray(ids) || ids.length > MAX_LINKED_CONTROL_ITEMS) return false
  const seen = new Set<number>()
  for (const id of ids) {
    if (!Number.isSafeInteger(id) || id < 0 || seen.has(id)) return false
    seen.add(id)
  }
  return true
}

/** The sole click normalization permitted by S5-C: Excel's measured Multi-Select-on
 * click from the all-selected state deselects the clicked item. Other click states
 * (including toggling Multi-Select, an already-selected item, or modifier clicks)
 * have no helper and remain unavailable until their owner recipes are run. */
export function measuredDeselectFromAll(name: string, allItemIds: readonly number[], currentlySelectedItemIds: readonly number[], clickedItemId: number): LinkedControlOperation | undefined {
  if (!name || !Number.isSafeInteger(clickedItemId) || !allItemIds.length || allItemIds.length > MAX_LINKED_CONTROL_ITEMS
    || currentlySelectedItemIds.length !== allItemIds.length) return undefined
  const seen = new Set<number>()
  for (const itemId of allItemIds) {
    if (!Number.isSafeInteger(itemId) || itemId < 0 || seen.has(itemId)) return undefined
    seen.add(itemId)
  }
  if (!seen.has(clickedItemId)) return undefined
  const current = new Set(currentlySelectedItemIds)
  if (current.size !== seen.size || [...seen].some(itemId => !current.has(itemId))) return undefined
  return { kind: 'slicerSelection', name, selectedItemIds: allItemIds.filter(itemId => itemId !== clickedItemId) }
}

export interface ControlMemoryHold { release(): void }
export interface ControlMemoryAdmission { readonly limitBytes?: number; reserve(bytes: number): ControlMemoryHold | undefined }
export interface LinkedControlEstimate { candidateBytes: number; replyBytes: number }
export type LinkedControlStage<C> =
  | { kind: 'staged'; value: C; bytes: number }
  | { kind: 'refused'; message: string }

/** `publish` must be synchronous and non-throwing: commit first validates every target, then publishes the complete
 * candidate set without yielding. `discard` releases any target-local construction holds on every refusal path. */
export interface LinkedControlTarget<O, C> {
  /** Stable identity, e.g. JSON.stringify([sheetId, pivotIndex]). */
  readonly id: string
  estimate(operation: O): LinkedControlEstimate
  stage(operation: O, context: { revision: number; grantBytes: number; signal: AbortSignal }): LinkedControlStage<C> | Promise<LinkedControlStage<C>>
  current(candidate: C, revision: number): boolean
  publish(candidate: C, revision: number): void
  discard(candidate: C): void
}

export type LinkedControlOutcome =
  | { kind: 'staged'; revision: number; targetIds: readonly string[]; bytes: number }
  | { kind: 'committed'; revision: number; targetIds: readonly string[] }
  | { kind: 'stale' | 'cancelled' | 'released' }
  | { kind: 'refused'; feature: string; message: string; requiredBytes?: number; limitBytes?: number }

interface Candidate<O, C> { target: LinkedControlTarget<O, C>; value: C }
interface Stage<O, C> { revision: number; operation: O; hold: ControlMemoryHold; candidates: Candidate<O, C>[]; controller: AbortController; disposed: boolean; building: boolean; externalSignal: AbortSignal; abortListener: () => void }

/** Stages all linked outputs under one admission, then validates the whole set before a single synchronous publish. */
export class LinkedControlTransaction<O, C> {
  private staged?: Stage<O, C>
  private committedHold?: ControlMemoryHold
  private committedRevision = 0
  private latestRevision = 0
  private released = false
  private epoch = 0
  private readonly targetIds: readonly string[]

  constructor(private readonly targets: readonly LinkedControlTarget<O, C>[], private readonly memory: ControlMemoryAdmission) {
    const ids = new Set(targets.map(target => target.id))
    if (!targets.length || targets.length > MAX_LINKED_CONTROL_TARGETS || ids.size !== targets.length || targets.some(target => !target.id)) {
      throw new RangeError(`Linked control target count must be 1..${MAX_LINKED_CONTROL_TARGETS} with unique stable ids`)
    }
    this.targetIds = targets.map(target => target.id)
  }

  async stage(operation: O, options: { revision: number; grantBytes: number; signal: AbortSignal }): Promise<LinkedControlOutcome> {
    const { revision, grantBytes, signal } = options
    if (this.released) return { kind: 'released' }
    if (signal.aborted) return { kind: 'cancelled' }
    if (!validateLinkedControlOperation(operation)) return { kind: 'refused', feature: 'pivot-control-operation-invalid', message: 'The linked control selection is invalid or exceeds its item limit.' }
    if (!Number.isSafeInteger(revision) || revision <= this.latestRevision || !Number.isSafeInteger(grantBytes) || grantBytes < 0) return { kind: 'stale' }
    this.dropStage()
    this.latestRevision = revision
    const estimates: LinkedControlEstimate[] = []
    let requiredBytes = 0
    try {
      for (const target of this.targets) {
        const estimate = target.estimate(operation)
        if (!Number.isSafeInteger(estimate.candidateBytes) || estimate.candidateBytes < 0
          || !Number.isSafeInteger(estimate.replyBytes) || estimate.replyBytes < 0) {
          return { kind: 'refused', feature: 'pivot-control-estimate-invalid', message: 'A linked control target returned an invalid memory estimate.' }
        }
        requiredBytes += estimate.candidateBytes + estimate.replyBytes
        estimates.push(estimate)
      }
    } catch (error) {
      return { kind: 'refused', feature: 'pivot-control-estimate-failed', message: error instanceof Error ? error.message.slice(0, 256) : 'Linked control memory could not be estimated.' }
    }
    if (!Number.isSafeInteger(requiredBytes) || requiredBytes > grantBytes) {
      return { kind: 'refused', feature: 'pivot-control-memory-limit', message: 'Linked control candidates exceed the receiver memory grant.', requiredBytes, limitBytes: grantBytes }
    }
    const hold = this.memory.reserve(requiredBytes)
    if (!hold) return { kind: 'refused', feature: 'pivot-control-memory-limit', message: 'Linked control candidates exceed the workbook memory budget.', requiredBytes, ...(this.memory.limitBytes === undefined ? {} : { limitBytes: this.memory.limitBytes }) }

    const epoch = ++this.epoch, controller = new AbortController()
    const candidates: Candidate<O, C>[] = []
    const stage: Stage<O, C> = { revision, operation, hold, candidates, controller, disposed: false, building: true, externalSignal: signal,
      abortListener: () => { controller.abort(); if (!stage.building && this.staged === stage) this.discardCandidates(stage) } }
    signal.addEventListener('abort', stage.abortListener, { once: true })
    this.staged = stage
    try {
      for (let index = 0; index < this.targets.length; index++) {
        if (controller.signal.aborted || this.released || this.epoch !== epoch || this.staged !== stage) {
          this.discardCandidates(stage)
          return { kind: 'cancelled' }
        }
        const target = this.targets[index]!, estimate = estimates[index]!
        const result = await target.stage(operation, { revision, grantBytes: estimate.candidateBytes + estimate.replyBytes, signal: controller.signal })
        if (stage.disposed || controller.signal.aborted || this.released || this.epoch !== epoch || this.staged !== stage) {
          if (result.kind === 'staged') target.discard(result.value)
          this.discardCandidates(stage)
          return { kind: 'cancelled' }
        }
        if (result.kind === 'refused') {
          this.discardCandidates(stage)
          return { kind: 'refused', feature: 'pivot-control-target-refused', message: result.message.slice(0, 256) }
        }
        if (!Number.isSafeInteger(result.bytes) || result.bytes < 0 || result.bytes > estimate.candidateBytes + estimate.replyBytes) {
          target.discard(result.value)
          this.discardCandidates(stage)
          return { kind: 'refused', feature: 'pivot-control-reply-limit', message: 'A linked control target exceeded its admitted candidate and reply bytes.', requiredBytes: result.bytes, limitBytes: estimate.candidateBytes + estimate.replyBytes }
        }
        candidates.push({ target, value: result.value })
      }
      if (controller.signal.aborted || this.released || this.epoch !== epoch || this.staged !== stage) {
        this.discardCandidates(stage)
        return { kind: 'cancelled' }
      }
      stage.building = false
      return { kind: 'staged', revision, targetIds: this.targetIds, bytes: requiredBytes }
    } catch (error) {
      this.discardCandidates(stage)
      if (controller.signal.aborted || signal.aborted) return { kind: 'cancelled' }
      return { kind: 'refused', feature: 'pivot-control-stage-failed', message: error instanceof Error ? error.message.slice(0, 256) : 'A linked control target could not be staged.' }
    } finally {
      // A cancelled producer may ignore its signal. Keep its construction grant
      // until the awaited producer has returned and its late candidate is discarded.
      stage.building = false
      if (stage.disposed) stage.hold.release()
      if (this.staged !== stage) signal.removeEventListener('abort', stage.abortListener)
    }
  }

  commit(revision: number): LinkedControlOutcome {
    if (this.released) return { kind: 'released' }
    const stage = this.staged
    if (!stage || stage.building) return { kind: 'stale' }
    if (stage.revision !== revision || stage.controller.signal.aborted) { this.dropStage(); return { kind: 'stale' } }
    if (!stage.candidates.every(({ target, value }) => target.current(value, revision))) {
      this.dropStage()
      return { kind: 'stale' }
    }
    stage.externalSignal.removeEventListener('abort', stage.abortListener)
    // No await or callback that can re-enter between validation and publication.
    for (const { target, value } of stage.candidates) target.publish(value, revision)
    this.staged = undefined
    const previous = this.committedHold
    this.committedHold = stage.hold
    this.committedRevision = revision
    previous?.release()
    return { kind: 'committed', revision, targetIds: this.targetIds }
  }

  cancel(revision: number): boolean {
    if (this.staged?.revision !== revision) return false
    this.dropStage()
    return true
  }

  release(): void {
    if (this.released) return
    this.released = true
    this.epoch++
    this.dropStage()
    this.committedHold?.release()
    this.committedHold = undefined
  }

  get revision(): number { return this.committedRevision }
  get closed(): boolean { return this.released }

  private discardCandidates(stage: Stage<O, C>): void {
    if (stage.disposed) return
    stage.disposed = true
    stage.externalSignal.removeEventListener('abort', stage.abortListener)
    if (this.staged === stage) this.staged = undefined
    stage.controller.abort()
    for (const candidate of stage.candidates.splice(0)) candidate.target.discard(candidate.value)
    if (!stage.building) stage.hold.release()
  }
  private dropStage(): void {
    const stage = this.staged
    if (stage) this.discardCandidates(stage)
  }
}

export type SharedControlIndexResult<V> =
  | { kind: 'acquired'; value: V; lease: { release(): void } }
  | { kind: 'refused'; requiredBytes: number; limitBytes?: number }

/** One admitted source index shared by every active control using a cache. */
export class SharedControlIndexCache<K, V> {
  private readonly entries = new Map<K, { value: V; refs: number; hold: ControlMemoryHold }>()
  constructor(private readonly memory: ControlMemoryAdmission, private readonly maxEntries = 64) {}

  acquire(key: K, bytes: number, build: () => V): SharedControlIndexResult<V> {
    const existing = this.entries.get(key)
    if (existing) {
      existing.refs++
      return { kind: 'acquired', value: existing.value, lease: this.lease(key, existing) }
    }
    if (!Number.isSafeInteger(bytes) || bytes < 0 || this.entries.size >= this.maxEntries) {
      return { kind: 'refused', requiredBytes: bytes, ...(this.memory.limitBytes === undefined ? {} : { limitBytes: this.memory.limitBytes }) }
    }
    const hold = this.memory.reserve(bytes)
    if (!hold) return { kind: 'refused', requiredBytes: bytes, ...(this.memory.limitBytes === undefined ? {} : { limitBytes: this.memory.limitBytes }) }
    try {
      const value = build()
      const entry = { value, refs: 1, hold }
      this.entries.set(key, entry)
      return { kind: 'acquired', value, lease: this.lease(key, entry) }
    } catch (error) {
      hold.release()
      throw error
    }
  }

  releaseAll(): void {
    for (const entry of this.entries.values()) entry.hold.release()
    this.entries.clear()
  }

  private lease(key: K, entry: { value: V; refs: number; hold: ControlMemoryHold }): { release(): void } {
    let live = true
    return { release: () => {
      if (!live) return
      live = false
      if (this.entries.get(key) !== entry) return
      entry.refs--
      if (entry.refs === 0) { this.entries.delete(key); entry.hold.release() }
    } }
  }
}

export interface LinkedPivotDescriptor { sheetId: string; pivot: number; name: string; cacheId: number; layoutKey?: string }
export interface SlicerPivotLink { name: string; kind: 'tabular'; pivotTables: readonly string[]; pivotCacheId: number }
export type ControlTargetsResult = { ok: true; targets: readonly LinkedPivotDescriptor[] } | { ok: false; feature: string; message: string }

export interface WorkerControlLink {
  readonly name: string
  readonly kind: 'slicer' | 'timeline'
  /** Workbook pivot cache id, resolved from the control's x14 cache id. */
  readonly cacheId?: number
  readonly pivotTables: readonly string[]
}
export interface WorkerControlTarget { readonly name: string; readonly cacheId: number }
export type WorkerControlLinkResult = { ok: true; link: WorkerControlLink } | { ok: false; feature: string; message: string }

/** Validates the saved workbook linkage again at the worker boundary; caller-supplied targets are not authority. */
export function resolveWorkerControlLink(selection: LinkedControlOperation, links: readonly WorkerControlLink[], targets: readonly WorkerControlTarget[]): WorkerControlLinkResult {
  const kind = selection.kind.startsWith('timeline') ? 'timeline' : 'slicer'
  const matches = links.filter(link => link.name === selection.name && link.kind === kind)
  // Repeated drawings of one cache emit identical links; they are one linkage.
  // Genuinely conflicting links (another cache or other targets) stay ambiguous.
  const distinct = [...new Map(matches.map(link => [JSON.stringify([link.kind, link.cacheId, [...link.pivotTables].sort()]), link])).values()]
  if (distinct.length !== 1) return { ok: false,
    feature: distinct.length ? 'pivot-control-linkage-ambiguous' : 'pivot-control-linkage-unknown',
    message: distinct.length ? `Control ${selection.name} has ambiguous saved linkage.` : `Control ${selection.name} has no known saved workbook linkage.` }
  const link = distinct[0]!
  if (!Number.isSafeInteger(link.cacheId) || link.cacheId! < 0 || !link.pivotTables.length || link.pivotTables.length > MAX_LINKED_CONTROL_TARGETS
    || link.pivotTables.some(name => !name) || new Set(link.pivotTables).size !== link.pivotTables.length) {
    return { ok: false, feature: 'pivot-control-linkage-unsupported', message: `Control ${selection.name} has incomplete or unsupported saved linkage.` }
  }
  if (targets.length !== link.pivotTables.length) return { ok: false, feature: 'pivot-control-target-unlinked', message: `The requested pivots do not match all pivots linked to ${selection.name}.` }
  const wanted = new Set(link.pivotTables)
  const seen = new Set<string>()
  for (const target of targets) {
    if (target.cacheId !== link.cacheId) return { ok: false, feature: 'pivot-control-linkage-cache-mismatch', message: `A target of ${selection.name} uses a different pivot cache.` }
    if (!wanted.has(target.name)) return { ok: false, feature: 'pivot-control-target-unlinked', message: `Pivot ${target.name} is not linked to ${selection.name}.` }
    if (seen.has(target.name)) return { ok: false, feature: 'pivot-control-linkage-ambiguous', message: `Pivot ${target.name} is duplicated in the linked-control request.` }
    seen.add(target.name)
  }
  return { ok: true, link }
}

/** Resolves only the pivot tables listed by a saved slicer cache; omitted pivots on that cache stay independent. */
export function resolveSlicerPivotTargets(link: SlicerPivotLink, pivots: readonly LinkedPivotDescriptor[]): ControlTargetsResult {
  if (!link.name || link.kind !== 'tabular' || !Number.isSafeInteger(link.pivotCacheId) || link.pivotCacheId < 0) {
    return { ok: false, feature: 'pivot-control-linkage-unsupported', message: 'The slicer cache linkage is unsupported.' }
  }
  if (!link.pivotTables.length || link.pivotTables.length > MAX_LINKED_CONTROL_TARGETS) {
    return { ok: false, feature: 'pivot-control-target-limit', message: 'The slicer target list is empty or exceeds the linked-pivot limit.' }
  }
  if (pivots.length > 10_000) return { ok: false, feature: 'pivot-control-target-limit', message: 'The workbook pivot target index exceeds its bound.' }
  const byNameAndCache = new Map<string, LinkedPivotDescriptor[]>()
  for (const pivot of pivots) {
    const key = JSON.stringify([pivot.cacheId, pivot.name]), list = byNameAndCache.get(key)
    if (list) list.push(pivot); else byNameAndCache.set(key, [pivot])
  }
  const targets: LinkedPivotDescriptor[] = []
  const seen = new Set<string>()
  for (const name of link.pivotTables) {
    const matches = byNameAndCache.get(JSON.stringify([link.pivotCacheId, name])) ?? []
    if (matches.length !== 1) return { ok: false, feature: matches.length ? 'pivot-control-target-ambiguous' : 'pivot-control-target-missing', message: `Slicer ${link.name} target ${name} is ${matches.length ? 'ambiguous' : 'missing'} or uses another cache.` }
    const target = matches[0]!, id = JSON.stringify([target.sheetId, target.pivot])
    if (seen.has(id)) return { ok: false, feature: 'pivot-control-target-duplicate', message: `Slicer ${link.name} contains a duplicate target ${name}.` }
    seen.add(id); targets.push(target)
  }
  if (targets.length > 1) {
    const layoutKeys = targets.map(target => target.layoutKey)
    if (layoutKeys.some(key => !key) || new Set(layoutKeys).size !== 1) return { ok: false, feature: 'pivot-control-linked-layout-unverified', message: `Slicer ${link.name} links pivots with different or unverified layouts.` }
  }
  return { ok: true, targets }
}
