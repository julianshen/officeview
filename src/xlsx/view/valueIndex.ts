import type { FilterRange, PagedKeySetRef, SourceFailure, SourceStamp, SourceStore, ValueSearchPredicate, ValuesFilter } from './contracts'
import type { ColumnClassification, PreparedValue } from './filter'
import type { FilterLimits } from './textFilter'
import { ValueIndexError, valueKeyString, KEY_NUMBER, type ValueRunStore } from './valueRuns'

export { ValueIndexError } from './valueRuns'

/** Source lane tags; 0 means an uncovered cell, which is a page error. */
export const SOURCE_NUMBER = 1, SOURCE_STRING = 2, SOURCE_BOOLEAN = 3, SOURCE_ERROR = 4
export const SOURCE_BLANK = 5, SOURCE_UNKNOWN = 6, SOURCE_UNAVAILABLE = 7
export const VALUE_DISPLAY_LIMIT = 10_000
export const VALUE_PAGE_LIMIT = 500
export const VALUE_LABEL_LIMIT = 256
export interface ValueLanePage {
  kind: 'values'
  pageId: string
  count: number
  tags: Uint8Array
  numbers: Float64Array
  texts: readonly (string | undefined)[]
  styles?: Uint32Array
}
export type ValuePageRead = ValueLanePage | { kind: 'unknown'; pageId: string } | { kind: 'error'; failure: SourceFailure }
export interface ValueColumnSource {
  readonly stamp: SourceStamp
  readonly pageCount: number
  readPage(page: number, signal: AbortSignal): Promise<ValuePageRead>
}
export interface ValueIndexIdentity {
  docId: string
  sheetId: string
  valueGen: number
  range: FilterRange
  col: number
  classification: ColumnClassification
  classificationGen: number
}
export type ValueCompleteness = 'complete' | 'loading' | 'incomplete'
export interface ValueIndexLedger {
  sourcePageReads: number
  mergedPages: number
  unknownPages: number
  allocations: number
  allocatedBytes: number
  chargedBytes: number
  peakBytes: number
}
export type ValueLabeler = (kind: 'number' | 'date' | 'boolean', value: number, style: number) => string
export type ValueEntryKind = 'number' | 'date' | 'string' | 'boolean' | 'error'
export interface ValueEntryView { ordinal: number; key: string; kind: ValueEntryKind; label: string; labelTruncated: boolean; count: number }
export interface ValueSearchResult {
  query: ValueSearchPredicate
  owner: SourceStamp
  matches: Uint32Array
  count: number
  searchedEntries: number
  complete: boolean
}
export interface ValuePageRequest { offset: number; limit: number; search?: ValueSearchResult; label?: ValueLabeler }
export interface ValuePage { entries: readonly ValueEntryView[]; total: number; displayed: number; truncated: boolean; completeness: ValueCompleteness }
export interface ValueBuildOptions { signal?: AbortSignal }
export interface ValueSearchOptions { signal?: AbortSignal; label?: ValueLabeler; limits?: FilterLimits; yield?: () => Promise<void>; resume?: ValueSearchResult }

export class ValueIndexCache {
  readonly limitBytes: number
  private readonly indexes = new Map<string, ValueIndex>()
  constructor(options: { limitBytes?: number } = {}) { this.limitBytes = options.limitBytes ?? 512 * 1024 * 1024 }
  get usedBytes(): number { return 0 }
  get size(): number { return this.indexes.size }
  open(identity: ValueIndexIdentity, source: ValueColumnSource): ValueIndex {
    const key = `${identity.docId}|${identity.sheetId}|${identity.col}`
    let index = this.indexes.get(key)
    if (!index) { index = new ValueIndex(identity, source); this.indexes.set(key, index) }
    return index
  }
  invalidate(_predicate: (identity: ValueIndexIdentity) => boolean = () => true): void {}
}

export class ValueIndex {
  readonly ledger: ValueIndexLedger = { sourcePageReads: 0, mergedPages: 0, unknownPages: 0, allocations: 0, allocatedBytes: 0, chargedBytes: 0, peakBytes: 0 }
  blankCount = 0
  unavailableCount = 0
  disposed = false
  private readonly keys: string[] = []
  private readonly labels: string[] = []
  private done = false
  constructor(readonly identity: ValueIndexIdentity, readonly source: ValueColumnSource) {}
  get size(): number { return this.keys.length }
  get completeness(): ValueCompleteness { return this.done ? 'complete' : 'loading' }
  get stamp(): SourceStamp { return { docId: this.identity.docId, sheetId: this.identity.sheetId, valueGen: this.identity.valueGen } }
  async build(options: ValueBuildOptions = {}): Promise<ValueCompleteness> {
    const seen = new Set(this.keys)
    for (let page = 0; page < this.source.pageCount; page++) {
      const read = await this.source.readPage(page, options.signal ?? new AbortController().signal)
      this.ledger.sourcePageReads++
      if (read.kind !== 'values') continue
      for (let i = 0; i < read.count; i++) {
        if (read.tags[i] === SOURCE_BLANK) { this.blankCount++; continue }
        const key = read.tags[i] === SOURCE_STRING ? `s:${read.texts[i]}` : `n:${read.numbers[i]}`
        // Deliberately wrong: the display cap is applied to construction.
        if (!seen.has(key) && this.keys.length < VALUE_DISPLAY_LIMIT) { seen.add(key); this.keys.push(key); this.labels.push(key.slice(2)) }
      }
    }
    this.done = true
    return 'complete'
  }
  key(ordinal: number): string { return this.keys[ordinal] }
  find(value: PreparedValue): number { return this.keys.indexOf(value.kind === 'string' ? `s:${value.value}` : `n:${'value' in value ? value.value : ''}`) }
  page(request: ValuePageRequest): ValuePage {
    const entries = this.keys.slice(request.offset, request.offset + request.limit).map((key, at) => ({
      ordinal: request.offset + at, key, kind: 'string' as const, label: this.labels[request.offset + at], labelTruncated: false, count: 1,
    }))
    return { entries, total: this.keys.length, displayed: this.keys.length, truncated: false, completeness: this.completeness }
  }
  async search(query: ValueSearchPredicate, _options: ValueSearchOptions = {}): Promise<ValueSearchResult> {
    const matches = new Uint32Array(Math.ceil(this.keys.length / 32))
    let count = 0
    for (let i = 0; i < Math.min(this.keys.length, VALUE_DISPLAY_LIMIT); i++) {
      if (this.labels[i].includes(query.text)) { matches[Math.floor(i / 32)] += 2 ** (i % 32); count++ }
    }
    return { query, owner: this.stamp, matches, count, searchedEntries: this.keys.length, complete: true }
  }
}

export interface ValueDraftPatchSummary { kind: 'keys' | 'search'; selected: boolean; count: number }
export type ValueCommit = { kind: 'unchanged' | 'committed'; rule: ValuesFilter | null }
export class ValueSelectionDraft {
  closed = false
  private readonly selected = new Map<number, boolean>()
  private showBlanks = true
  private constructor(readonly index: ValueIndex, readonly committed: ValuesFilter | null, readonly store: ValueRunStore) {}
  static async open(index: ValueIndex, committed: ValuesFilter | null, store: ValueRunStore): Promise<ValueSelectionDraft> {
    return new ValueSelectionDraft(index, committed, store)
  }
  get blanks(): boolean { return this.showBlanks }
  isSelected(ordinal: number): boolean { return this.selected.get(ordinal) ?? true }
  toggle(ordinal: number, selected: boolean): void { this.selected.set(ordinal, selected) }
  setBlanks(selected: boolean): void { this.showBlanks = selected }
  selectAll(selected: boolean, _search?: ValueSearchResult, _options: { includeBlanks?: boolean } = {}): void {
    for (let i = 0; i < Math.min(this.index.size, VALUE_DISPLAY_LIMIT); i++) this.selected.set(i, selected)
  }
  search(query: ValueSearchPredicate, options: ValueSearchOptions = {}): Promise<ValueSearchResult> { return this.index.search(query, options) }
  selectedCount(): { count: number; complete: boolean } {
    let count = 0
    for (let i = 0; i < this.index.size; i++) if (this.isSelected(i)) count++
    return { count, complete: true }
  }
  describe(): { base: 'all' | 'none'; patches: readonly ValueDraftPatchSummary[]; blanks: boolean } {
    return { base: 'all', patches: [], blanks: this.showBlanks }
  }
  async commit(): Promise<ValueCommit> {
    // Deliberately wrong: only the displayed prefix's selected keys are stored.
    const keys: number[] = []
    for (let i = 0; i < Math.min(this.index.size, VALUE_DISPLAY_LIMIT); i++) if (this.isSelected(i)) keys.push(i)
    const run = { count: keys.length, kinds: new Uint8Array(keys.length).fill(KEY_NUMBER), numbers: new Float64Array(keys), texts: [] }
    const ref = await this.store.put(run, this.index.stamp)
    return { kind: 'committed', rule: { kind: 'values', selection: { mode: 'include', keys: ref }, showBlanks: this.showBlanks } }
  }
  cancel(): void { this.closed = true }
}

export async function compileValuesMembership(
  _rule: ValuesFilter, _classification: ColumnClassification, _store: ValueRunStore, _owner: SourceStamp, _signal?: AbortSignal,
): Promise<(value: PreparedValue) => boolean> {
  return () => true
}

export function valueSourceFromStore(store: SourceStore, col: number, body: { r0: number; r1: number }, pageRows = 4096): ValueColumnSource {
  return {
    stamp: store.stamp,
    pageCount: Math.ceil((body.r1 - body.r0 + 1) / pageRows),
    readPage: async () => ({ kind: 'unknown', pageId: `${col}` }),
  }
}
export type { PagedKeySetRef }
void valueKeyString
void ValueIndexError
