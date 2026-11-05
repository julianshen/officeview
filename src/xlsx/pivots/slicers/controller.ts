import type { SlicerModel, TimelineLevel, TimelineModel } from './model'

export type SlicerAction = { kind: 'selection'; selectedItemIds: readonly number[] } | { kind: 'clear' }
export type SlicerApplyResult = { kind: 'committed'; revision: number } | { kind: 'stale' | 'cancelled' } | { kind: 'refused'; message: string }
export interface SlicerControllerHost {
  apply(model: SlicerModel, action: SlicerAction, signal: AbortSignal): Promise<SlicerApplyResult>
  announce(message: string): void
  onChange?(model: SlicerModel): void
  /** Reserve an upper bound before copying the committed item state. */
  admit?(bytes: number): { release(): void } | undefined
}
const UNVERIFIED = 'This slicer action has not been measured in Excel and is not available.'

/** View-only committed slicer selection. It publishes a new model only after the transaction commits. */
export class SlicerController {
  private readonly models = new Map<string, SlicerModel>()
  private readonly filteredFields = new Map<string, boolean>()
  private readonly listeners = new Set<() => void>()
  private readonly holds = new Map<string, { release(): void }>()
  private disposed = false
  private queue: Promise<unknown> = Promise.resolve()

  constructor(models: readonly SlicerModel[], private readonly host: SlicerControllerHost) {
    for (const model of models) if (!this.models.has(model.name)) { this.models.set(model.name, model); this.reindexField(model) }
  }

  register(models: readonly SlicerModel[]): void {
    if (this.disposed) return
    for (const model of models) {
      const current = this.models.get(model.name)
      if (!current || (model.revision ?? 0) > (current.revision ?? 0)) {
        this.models.set(model.name, model)
        this.reindexField(model)
      }
    }
  }

  model(name: string): SlicerModel | undefined { return this.models.get(name) }
  multiSelect(_name: string): boolean { return false }
  subscribe(listener: () => void): () => void { this.listeners.add(listener); return () => { this.listeners.delete(listener) } }
  private publish(model: SlicerModel): void {
    if (this.disposed) return
    this.models.set(model.name, model)
    this.reindexField(model)
    this.host.onChange?.(model)
    for (const listener of this.listeners) listener()
  }

  applySlicer(name: string, selection: readonly number[], signal: AbortSignal): Promise<boolean> {
    return this.enqueue(name, { kind: 'selection', selectedItemIds: [...selection] }, signal)
  }
  clearSlicer(name: string, signal: AbortSignal): Promise<boolean> { return this.enqueue(name, { kind: 'clear' }, signal) }

  /** S5-C: the Multi-Select toggle and its click semantics are still owner recipes, so it fails closed. */
  toggleMultiSelect(name: string, _enabled: boolean): boolean {
    if (!this.disposed && this.models.has(name)) this.host.announce(UNVERIFIED)
    return false
  }

  isFieldFiltered(cacheId: number | undefined, fieldIndex: number): boolean {
    return this.filteredFields.get(`${cacheId ?? -1}:${fieldIndex}`) ?? false
  }

  private reindexField(model: SlicerModel): void {
    if (model.cacheId === undefined || model.sourceFieldIndex === undefined) return
    const key = `${model.cacheId}:${model.sourceFieldIndex}`
    // Multiple slicers can address one cache field. Its filter indicator remains on while any is restricted.
    let filtered = false
    for (const candidate of this.models.values()) {
      if (candidate.name === model.name || candidate.cacheId !== model.cacheId || candidate.sourceFieldIndex !== model.sourceFieldIndex) continue
      const all = candidate.allItemIds ?? candidate.items.map(item => item.id)
      const selected = new Set(candidate.selectedItemIds ?? candidate.items.filter(item => item.selected).map(item => item.id))
      if (all.some(id => !selected.has(id))) { filtered = true; break }
    }
    if (!filtered) {
      const all = model.allItemIds ?? model.items.map(item => item.id)
      const selected = new Set(model.selectedItemIds ?? model.items.filter(item => item.selected).map(item => item.id))
      filtered = all.some(id => !selected.has(id))
    }
    this.filteredFields.set(key, filtered)
  }

  private enqueue(name: string, action: SlicerAction, signal: AbortSignal): Promise<boolean> {
    const run = this.queue.then(async () => {
      const before = this.models.get(name)
      if (this.disposed || signal.aborted || !before) return false
      const itemCount = before.allItemIds?.length ?? before.items.length
      if (!itemCount || itemCount > 10_000 || before.items.length > 10_000) {
        this.host.announce('This slicer selection is unavailable because its item identities are incomplete.')
        return false
      }
      // Includes transient id vectors/sets and the retained model, before any copies.
      const candidateBytes = 256 + before.items.length * 128 + itemCount * 128
      const hold = this.host.admit?.(candidateBytes)
      if (this.host.admit && !hold) { this.host.announce('The slicer selection exceeds the viewer memory budget.'); return false }
      const all = before.allItemIds ?? before.items.map(item => item.id)
      const selectable = before.items.map(item => item.id)
      const allSet = new Set(all)
      if (allSet.size !== all.length || all.some(id => !Number.isSafeInteger(id) || id < 0)) {
        hold?.release()
        this.host.announce('This slicer selection is unavailable because its item identities are incomplete.')
        return false
      }
      const currentSelected = new Set(before.selectedItemIds ?? before.items.filter(item => item.selected).map(item => item.id))
      let selected: ReadonlySet<number>
      if (action.kind === 'clear') selected = allSet
      else {
        const ids = action.selectedItemIds, wanted = new Set(ids)
        const wasAll = all.every(id => currentSelected.has(id))
        const alreadyAllSelected = wanted.size === all.length && all.every(id => wanted.has(id))
        // The measured click-one case starts from All and selects exactly one item.
        if (!wasAll || alreadyAllSelected || ids.length !== 1 || wanted.size !== ids.length || ids.some(id => !allSet.has(id) || !selectable.includes(id))) {
        hold?.release(); this.host.announce(UNVERIFIED)
          return false
        }
        selected = wanted
      }
      let result: SlicerApplyResult
      try { result = await this.host.apply(before, action, signal) }
      catch (error) { result = { kind: 'refused', message: error instanceof Error ? error.message.slice(0, 256) : 'The slicer change failed.' } }
      if (signal.aborted || this.disposed || result.kind !== 'committed') {
        hold?.release()
        if (!signal.aborted && !this.disposed) this.host.announce(result.kind === 'refused' ? result.message : 'The slicer changed meanwhile; try again.')
        return false
      }
      const current = this.models.get(name)
      if (current !== before) { hold?.release(); this.host.announce('The slicer changed meanwhile; try again.'); return false }
      this.holds.get(name)?.release()
      if (hold) this.holds.set(name, hold)
      this.publish({ ...before, selectedItemIds: [...selected], revision: Math.max((before.revision ?? 0) + 1, result.revision), items: before.items.map(item => ({ ...item, selected: selected.has(item.id) })) })
      this.host.announce('Slicer selection applied (view only; the file is unchanged).')
      return true
    }).catch(error => { if (!this.disposed) this.host.announce(error instanceof Error ? error.message.slice(0, 256) : 'Slicer change failed.'); return false })
    this.queue = run
    return run
  }

  dispose(): void { this.disposed = true; this.listeners.clear(); this.models.clear(); this.filteredFields.clear(); for (const hold of this.holds.values()) hold.release(); this.holds.clear() }
}

export type TimelineAction = { kind: 'selection'; start: number; end: number } | { kind: 'clear' }
export type TimelineApplyResult = { kind: 'committed'; revision: number; filterId?: number } | { kind: 'stale' | 'cancelled' } | { kind: 'refused'; message: string }
export interface TimelineControllerHost {
  apply(model: TimelineModel, action: TimelineAction, signal: AbortSignal): Promise<TimelineApplyResult>
  announce(message: string): void
  onChange?(model: TimelineModel): void
  /** Reserve the bounded committed model replacement before constructing it. */
  admit?(bytes: number): { release(): void } | undefined
}
const TIMELINE_LEVELS: readonly TimelineLevel[] = ['years', 'quarters', 'months', 'days']
const DAY_MS = 86_400_000
function periodStart(level: TimelineLevel, value: number): number {
  const date = new Date(value), year = date.getUTCFullYear(), month = date.getUTCMonth(), day = date.getUTCDate()
  if (level === 'years') return Date.UTC(year, 0, 1)
  if (level === 'quarters') return Date.UTC(year, month - month % 3, 1)
  if (level === 'months') return Date.UTC(year, month, 1)
  return Date.UTC(year, month, day)
}
function periodEnd(level: TimelineLevel, value: number): number {
  const date = new Date(value), year = date.getUTCFullYear(), month = date.getUTCMonth(), day = date.getUTCDate()
  if (level === 'years') return Date.UTC(year + 1, 0, 1)
  if (level === 'quarters') return Date.UTC(year, month + 3 - month % 3, 1)
  if (level === 'months') return Date.UTC(year, month + 1, 1)
  return Date.UTC(year, month, day + 1)
}

/** Timeline state is view-only; a range is published only after the workbook-scoped linked transaction commits. */
export class TimelineController {
  private readonly models = new Map<string, TimelineModel>()
  private readonly listeners = new Set<() => void>()
  private readonly holds = new Map<string, { release(): void }>()
  private disposed = false
  private queue: Promise<unknown> = Promise.resolve()

  constructor(models: readonly TimelineModel[], private readonly host: TimelineControllerHost) {
    for (const model of models) if (!this.models.has(model.name)) this.models.set(model.name, model)
  }
  register(models: readonly TimelineModel[]): void {
    if (this.disposed) return
    for (const model of models) if (!this.models.has(model.name)) this.models.set(model.name, model)
  }
  model(name: string): TimelineModel | undefined { return this.models.get(name) }
  report(message: string): void { if (!this.disposed) this.host.announce(message) }
  subscribe(listener: () => void): () => void { this.listeners.add(listener); return () => { this.listeners.delete(listener) } }
  private publish(model: TimelineModel): void {
    if (this.disposed) return
    this.models.set(model.name, model)
    this.host.onChange?.(model)
    for (const listener of this.listeners) listener()
  }

  selectTimeline(name: string, startMs: number, endMs: number, signal: AbortSignal): Promise<boolean> {
    return this.enqueue(name, signal, model => {
      if (!Number.isSafeInteger(startMs) || !Number.isSafeInteger(endMs)) return undefined
      const first = periodStart(model.level, Math.min(startMs, endMs)), last = periodStart(model.level, Math.max(startMs, endMs))
      if (!Number.isFinite(first) || !Number.isFinite(last) || first >= model.range.end || last >= model.range.end) return undefined
      const end = Math.min(periodEnd(model.level, last) - DAY_MS, model.range.end - DAY_MS)
      return end < first ? undefined : { action: { kind: 'selection', start: first, end } as const, selection: { start: first, end } }
    })
  }

  clearTimeline(name: string, signal: AbortSignal): Promise<boolean> {
    return this.enqueue(name, signal, _model => ({ action: { kind: 'clear' }, clear: true }))
  }

  setTimelineLevel(name: string, level: TimelineLevel, signal: AbortSignal): Promise<boolean> {
    const run = this.queue.then(async () => {
      const before = this.models.get(name)
      if (this.disposed || signal.aborted || !before) return false
      if (!TIMELINE_LEVELS.includes(level)) { this.host.announce('This timeline level is not supported.'); return false }
      if (before.level === level) return true
      const hold = this.host.admit?.(128)
      if (this.host.admit && !hold) { this.host.announce('The timeline change exceeds the viewer memory budget.'); return false }
      try {
        if (signal.aborted || this.disposed || this.models.get(name) !== before) return false
        this.holds.get(name)?.release()
        if (hold) this.holds.set(name, hold)
        this.publish({ ...before, level, revision: Math.max((before.revision ?? 0) + 1, 0) })
        this.host.announce('Timeline display level changed (view only; the file is unchanged).')
        return true
      } finally { if (signal.aborted || this.disposed || this.models.get(name) === before) hold?.release() }
    }).catch(error => { if (!this.disposed) this.host.announce(error instanceof Error ? error.message.slice(0, 256) : 'Timeline level change failed.'); return false })
    this.queue = run
    return run
  }

  dispose(): void { this.disposed = true; this.listeners.clear(); this.models.clear(); for (const hold of this.holds.values()) hold.release(); this.holds.clear() }

  private enqueue(name: string, signal: AbortSignal, prepare: (model: TimelineModel) => { action: TimelineAction; selection?: { start: number; end: number }; clear?: true } | undefined): Promise<boolean> {
    const run = this.queue.then(async () => {
      const before = this.models.get(name)
      if (this.disposed || signal.aborted || !before) return false
      const prepared = prepare(before)
      if (!prepared) { this.host.announce('This timeline range is outside the supported calendar range.'); return false }
      const hold = this.host.admit?.(128)
      if (this.host.admit && !hold) { this.host.announce('The timeline change exceeds the viewer memory budget.'); return false }
      let result: TimelineApplyResult
      try { result = await this.host.apply(before, prepared.action, signal) }
      catch (error) { result = { kind: 'refused', message: error instanceof Error ? error.message.slice(0, 256) : 'The timeline change failed.' } }
      if (signal.aborted || this.disposed || result.kind !== 'committed' || this.models.get(name) !== before) {
        hold?.release()
        if (!signal.aborted && !this.disposed) this.host.announce(result.kind === 'refused' ? result.message : 'The timeline changed meanwhile; try again.')
        return false
      }
      this.holds.get(name)?.release()
      if (hold) this.holds.set(name, hold)
      const revision = Math.max((before.revision ?? 0) + 1, result.revision)
      const next: TimelineModel = prepared.clear
        ? (({ selection: _selection, filterId: _filterId, ...rest }) => ({ ...rest, revision }))(before)
        : { ...before, selection: prepared.selection, revision, ...(result.filterId === undefined ? {} : { filterId: result.filterId }) }
      this.publish(next)
      this.host.announce('Timeline selection applied (view only; the file is unchanged).')
      return true
    }).catch(error => { if (!this.disposed) this.host.announce(error instanceof Error ? error.message.slice(0, 256) : 'Timeline change failed.'); return false })
    this.queue = run
    return run
  }
}

/** Compares unique source-item indices without depending on menu order. */
export function sameSlicerHiddenItems(expected: readonly number[], actual: readonly number[]): boolean {
  if (expected.length !== actual.length) return false
  const hidden = new Set(actual)
  return expected.every(index => hidden.has(index))
}
