import type { SlicerModel } from './model'

export type SlicerAction = { kind: 'selection'; selectedItemIds: readonly number[] } | { kind: 'clear' }
export type SlicerApplyResult = { kind: 'committed'; revision: number } | { kind: 'stale' | 'cancelled' } | { kind: 'refused'; message: string }
export interface SlicerControllerHost {
  apply(model: SlicerModel, action: SlicerAction, signal: AbortSignal): Promise<SlicerApplyResult>
  announce(message: string): void
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

  model(name: string): SlicerModel | undefined { return this.models.get(name) }
  multiSelect(_name: string): boolean { return false }
  subscribe(listener: () => void): () => void { this.listeners.add(listener); return () => { this.listeners.delete(listener) } }
  private publish(model: SlicerModel): void {
    if (this.disposed) return
    this.models.set(model.name, model)
    this.reindexField(model)
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
      const all = before.allItemIds ?? before.items.map(item => item.id)
      const selectable = before.items.map(item => item.id)
      const allSet = new Set(all)
      if (!all.length || all.length > 10_000 || allSet.size !== all.length || all.some(id => !Number.isSafeInteger(id) || id < 0)) {
        this.host.announce('This slicer selection is unavailable because its item identities are incomplete.')
        return false
      }
      const candidateBytes = 256 + before.items.length * 64 + all.length * 8
      const hold = this.host.admit?.(candidateBytes)
      if (this.host.admit && !hold) { this.host.announce('The slicer selection exceeds the viewer memory budget.'); return false }
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
