import type { CacheIdentity, ColumnRule, FilterRange, FilterSpec, OperationRefusal, SortSpec, ValuesFilter } from '../../xlsx/view/contracts'
import { reduceSortCommand } from '../../xlsx/view/sort'
import { canonicalSearchKey, type ColumnClassification, type FilterContext } from '../../xlsx/view/filter'
import { detectOperationRange, admitSortRange, type RangeDetection, type RangeDetectionRequest } from '../../xlsx/view/rangeDetection'
import { ValueSelectionDraft, ValueIndexError, type ValueSearchResult, type ValueColumnSource, type ValueEntryView, type ValueIndexCache, type ValueIndexIdentity, type ValueLabeler, type ValueSelectedCount, type ValueSelectionState } from '../../xlsx/view/valueIndex'
import type { ValueRunStore } from '../../xlsx/view/valueRuns'

export interface ControlSnapshot {
  owner: CacheIdentity
  filter: FilterSpec | null
  sort: SortSpec | null
  records: { visible: number; total: number; completeness: 'complete' }
}
export type ControlFailure = { kind: 'refused'; refusal: OperationRefusal } | { kind: 'incomplete' } | { kind: 'stale' } | { kind: 'cancelled' }
export type ControlResult = ControlFailure | { kind: 'committed' | 'unchanged'; snapshot: ControlSnapshot }
export interface ControlProgress { completed: number; total: number; message: string }
export interface DraftQuery { text: string; wildcards?: boolean; offset: number }
export interface DraftPage {
  entries: readonly (ValueEntryView & { selected: ValueSelectionState })[]
  total: number
  displayed: number
  truncated: boolean
  count: ValueSelectedCount
  blanks: boolean
  complete: boolean
}
export type ValueDraftAction = { kind: 'toggle'; ordinal: number; selected: boolean } | { kind: 'all'; selected: boolean; includeBlanks: boolean } | { kind: 'blanks'; selected: boolean }
export interface ValuesDraftPort {
  readonly identity: ValueIndexIdentity
  read(query: DraftQuery, signal: AbortSignal, progress?: (value: ControlProgress) => void): Promise<DraftPage>
  change(action: ValueDraftAction, query: DraftQuery, signal: AbortSignal): Promise<DraftPage>
  candidate(signal: AbortSignal): Promise<ValuesFilter | null>
  inspect(ordinal: number, signal: AbortSignal): Promise<string>
  transfer(): void
  cancel(): void
}
export interface ControlsClient {
  owner(): CacheIdentity
  context(): FilterContext
  resolve(owner: CacheIdentity, kind: 'filter' | 'sort', signal: AbortSignal, override?: 0 | 1, col?: number): Promise<RangeDetection>
  /** Domain classification for typed filters; independent of building the value universe. */
  classify(owner: CacheIdentity, range: FilterRange, col: number, signal: AbortSignal): ColumnClassification
  values(owner: CacheIdentity, range: FilterRange, col: number, applied: ValuesFilter | null, signal: AbortSignal): Promise<ValuesDraftPort>
  filter(owner: CacheIdentity, spec: FilterSpec | null, signal: AbortSignal, progress?: (value: ControlProgress) => void, lease?: ValuesDraftPort): Promise<ControlResult>
  sort(owner: CacheIdentity, spec: SortSpec | null, signal: AbortSignal, progress?: (value: ControlProgress) => void): Promise<ControlResult>
}
/** `committed`: the stage already committed in the worker that owns the view
 * (S5.4) and its view was adopted, so the owner it was staged under is
 * superseded by design; publication only reports it. */
export type StagedControl = ControlFailure | { kind: 'ready'; publish: () => ControlSnapshot | null; committed?: true }
export interface ControlsClientOptions {
  snapshot(): ControlSnapshot
  context(): FilterContext
  rangeRequest(col?: number): Omit<RangeDetectionRequest, 'options' | 'override'>
  domain(owner: CacheIdentity, range: FilterRange, col: number): { identity: ValueIndexIdentity; source: ValueColumnSource; label?: ValueLabeler }
  cache: ValueIndexCache
  runs: ValueRunStore
  stageFilter(spec: FilterSpec | null, owner: CacheIdentity, signal: AbortSignal, progress?: (value: ControlProgress) => void): Promise<StagedControl>
  stageSort(spec: SortSpec | null, owner: CacheIdentity, signal: AbortSignal, progress?: (value: ControlProgress) => void): Promise<StagedControl>
}
export function sameControlOwner(a: CacheIdentity, b: CacheIdentity): boolean {
  return a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen && a.fontGen === b.fontGen && a.viewGen === b.viewGen
}
/** Same owner except the view generation (the only part a committed stage moves). */
export function sameSourceOwner(a: CacheIdentity, b: CacheIdentity): boolean {
  return a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen && a.fontGen === b.fontGen
}
export function sameFilterRange(a: FilterRange, b: FilterRange): boolean {
  return a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1 && a.headerRows === b.headerRows && a.totalsRows === b.totalsRows
}
export function filterWithColumn(applied: FilterSpec | null, range: FilterRange, col: number, rule: ColumnRule | null): FilterSpec | null {
  const columns = new Map(applied && sameFilterRange(applied.range, range) ? applied.columns : [])
  if (rule) columns.set(col, rule)
  else columns.delete(col)
  return columns.size ? { range: { ...range }, columns } : null
}
function ruleRef(rule: ValuesFilter | null) {
  if (!rule) return null
  return rule.selection.mode === 'include' ? rule.selection.keys : rule.selection.mode === 'allExcept' ? rule.selection.excluded : null
}
function equalFilter(a: FilterSpec | null, b: FilterSpec | null): boolean {
  if (!a || !b) return a === b
  return sameFilterRange(a.range, b.range) && a.columns.size === b.columns.size && [...a.columns].every(([col, rule]) => JSON.stringify(rule) === JSON.stringify(b.columns.get(col)))
}
/** In-process reference transport. Production proxies these bounded draft calls
 * to a resident worker; publish is synchronous and must only commit model data. */
export function createControlsClient(options: ControlsClientOptions): ControlsClient {
  const valid = (owner: CacheIdentity, signal: AbortSignal) => !signal.aborted && sameControlOwner(owner, options.snapshot().owner)
  const assert = (owner: CacheIdentity, signal: AbortSignal) => {
    if (signal.aborted) throw new ValueIndexError('cancelled', 'Operation cancelled')
    if (!valid(owner, signal)) throw new ValueIndexError('stale-source', 'Source or view changed; reopen and retry.')
  }
  const apply = async (kind: 'filter' | 'sort', owner: CacheIdentity, spec: FilterSpec | SortSpec | null, signal: AbortSignal, progress?: (value: ControlProgress) => void, lease?: ValuesDraftPort): Promise<ControlResult> => {
    if (signal.aborted) return { kind: 'cancelled' }
    if (!valid(owner, signal)) return { kind: 'stale' }
    const before = options.snapshot()
    if (kind === 'filter' ? equalFilter(before.filter, spec as FilterSpec | null) : JSON.stringify(before.sort) === JSON.stringify(spec)) {
      return { kind: 'unchanged', snapshot: before }
    }
    const staged = kind === 'filter' ? await options.stageFilter(spec as FilterSpec | null, owner, signal, progress) : await options.stageSort(spec as SortSpec | null, owner, signal, progress)
    const committed = staged.kind === 'ready' && staged.committed === true
    if (!committed && signal.aborted) return { kind: 'cancelled' }
    if (!committed && !valid(owner, signal)) return { kind: 'stale' }
    // A worker-committed stage moved only its own view generation: the source,
    // value and font generations it was staged under must still be current.
    if (committed && !sameSourceOwner(owner, options.snapshot().owner)) return { kind: 'stale' }
    if (staged.kind !== 'ready') return staged
    // No await between this ownership check and publication/lease transfer.
    const snapshot = staged.publish()
    if (!snapshot) return { kind: 'stale' }
    lease?.transfer()
    if (kind === 'filter') {
      const retained = new Set([...snapshot.filter?.columns.values() ?? []].flatMap(rule => rule.kind === 'values' && ruleRef(rule) ? [ruleRef(rule)!.refId] : []))
      for (const rule of before.filter?.columns.values() ?? []) if (rule.kind === 'values') {
        const ref = ruleRef(rule)
        if (ref && !retained.has(ref.refId)) void options.runs.release(ref)
      }
    }
    return { kind: 'committed', snapshot }
  }
  return {
    owner: () => options.snapshot().owner, context: options.context,
    async resolve(owner, kind, signal, override, col) {
      assert(owner, signal)
      const request = options.rangeRequest(col)
      const result = await (kind === 'sort' ? admitSortRange : detectOperationRange)({ ...request, options: { stamp: { docId: owner.docId, sheetId: owner.sheetId, valueGen: owner.valueGen }, signal }, override: override === undefined ? null : { headerRows: override } })
      assert(owner, signal)
      return result
    },
    classify(owner, range, col, signal) {
      assert(owner, signal)
      return options.domain(owner, range, col).identity.classification
    },
    async values(owner, range, col, applied, signal) {
      assert(owner, signal)
      const domain = options.domain(owner, range, col), index = options.cache.open(domain.identity, domain.source)
      let draft = await ValueSelectionDraft.open(index, applied, options.runs, signal), closed = false
      let search: ValueSearchResult | undefined, queryKey = '', latestRule = applied
      const owned = new Map<string, NonNullable<ReturnType<typeof ruleRef>>>()
      const check = (input: AbortSignal) => { assert(owner, input); if (closed) throw new ValueIndexError('closed', 'Draft closed') }
      const page = (): DraftPage => {
        const result = index.page({ offset: currentQuery.offset, limit: 200, search, label: domain.label })
        const count = draft.selectedCount()
        return { ...result, entries: result.entries.map(entry => ({ ...entry, selected: draft.isSelected(entry.ordinal) })), count, blanks: draft.blanks,
          complete: count.complete && (!search || search.complete) && index.completeness === 'complete' }
      }
      let currentQuery: DraftQuery = { text: '', offset: 0 }
      const read = async (query: DraftQuery, input: AbortSignal, progress?: (value: ControlProgress) => void): Promise<DraftPage> => {
        check(input)
        await index.build({ signal: input, onProgress: p => progress?.({ ...p, message: 'Loading all values' }) })
        check(input)
        await draft.prepare({ signal: input, onProgress: p => progress?.({ completed: p.searched, total: p.total, message: 'Preparing selection' }) })
        check(input)
        const key = JSON.stringify([query.text, query.wildcards])
        if (index.completeness !== 'loading' && (key !== queryKey || search?.searchedEntries !== index.size)) {
          search = await draft.search({ text: query.text, wildcards: query.wildcards, locale: options.context().locale }, { signal: input, label: domain.label, onProgress: p => progress?.({ completed: p.searched, total: p.total, message: 'Searching all values' }) })
          queryKey = key
        }
        check(input)
        currentQuery = query
        return page()
      }
      return {
        identity: domain.identity, read,
        async change(action, query, input) {
          await read(query, input)
          if (action.kind === 'toggle') draft.toggle(action.ordinal, action.selected)
          else if (action.kind === 'blanks') draft.setBlanks(action.selected)
          else draft.selectAll(action.selected, canonicalSearchKey(query.text, options.context().locale) ? search : undefined, { includeBlanks: action.includeBlanks })
          await draft.prepare({ signal: input })
          check(input)
          return page()
        },
        async candidate(input) {
          check(input)
          const result = await draft.commit(input)
          if (result.kind === 'committed' && ruleRef(result.rule)) owned.set(ruleRef(result.rule)!.refId, ruleRef(result.rule)!)
          try {
            check(input)
            latestRule = result.rule
            draft = await ValueSelectionDraft.open(index, latestRule, options.runs, input)
            check(input)
            search = undefined; queryKey = ''
            return latestRule
          } catch (error) {
            for (const ref of owned.values()) void options.runs.release(ref)
            owned.clear()
            throw error
          }
        },
        async inspect(ordinal, input) {
          check(input)
          const text = index.entryText(ordinal)
          if (text && (text.length + 2) * 2 > 65536) throw new ValueIndexError('quota', 'Value exceeds the 64 KiB inspection budget')
          const key = index.key(ordinal)
          if (key.length * 2 > 65536) throw new ValueIndexError('quota', 'Value exceeds the 64 KiB inspection budget')
          return key
        },
        transfer() {
          const retained = ruleRef(latestRule)?.refId
          for (const ref of owned.values()) if (ref.refId !== retained) void options.runs.release(ref)
          owned.clear()
        },
        cancel() {
          closed = true; draft.cancel()
          for (const ref of owned.values()) void options.runs.release(ref)
          owned.clear()
        },
      }
    },
    filter: (owner, spec, signal, progress, lease) => apply('filter', owner, spec, signal, progress, lease),
    sort: (owner, spec, signal, progress) => apply('sort', owner, spec, signal, progress),
  }
}

/** Header commands use the same range admission and eight-key reducer as S7. */
export async function cycleHeaderSort(client: ControlsClient, snapshot: ControlSnapshot, col: number, shift: boolean, signal: AbortSignal): Promise<ControlResult> {
  const range = await client.resolve(snapshot.owner, 'sort', signal, undefined, col)
  if (range.kind !== 'complete') return range.kind === 'refused' ? range : { kind: range.kind === 'cancelled' ? 'cancelled' : 'incomplete' }
  const next = reduceSortCommand(snapshot.sort, { kind: 'cycle', range: range.range, col, shift })
  if (next.notice) return { kind: 'refused', refusal: { reason: 'invalid-rule', message: next.notice.reason === 'too-many-sort-keys' ? 'Choose at most eight sort levels.' : 'Invalid sort command.' } }
  return client.sort(snapshot.owner, next.sort, signal)
}
