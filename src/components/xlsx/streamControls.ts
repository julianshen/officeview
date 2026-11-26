/**
 * Stage 4c (P15, P16): the controls port of a `sheetSort` stream. It speaks the `ControlsClient` the sort and filter
 * menus already use, over the server's `range-detect`, `filter-values`, `sheet-sort` and `sheet-filter`
 * (`StreamWorkbook`). Nothing here indexes, sorts or filters a cell: there is no value index, no document session and
 * no structural query on a stream (the client holds no cell data, S4 §6.1).
 *
 * - `resolve` is the server's detection; a filter detection with a column also fetches that column's first value page,
 *   so `classify` (synchronous) has the server's classification when the dropdown asks for it.
 * - `values` is a draft `{base, flipped, blanks}` over the server's list ids (P15). Its search is sent 250 ms after the
 *   last keystroke, and a newer one supersedes it. `inspect` returns the label as sent.
 * - `filter` and `sort` build the wire specs: a column the user did not edit goes as `committed`, the draft gives the
 *   values rule, and a number or date condition is refused here when the session does not serve it (never sent).
 * - One `'sheet'` op is in flight per stream; the rest wait their turn, and one aborted while waiting is never sent.
 *
 * Reachable only from the lazy `InteractiveSheetDoc` chunk, never from the package index.
 */
import type { CacheIdentity, ColumnRule, FilterRange, FilterSpec, OperationRefusal, SheetView, ValuesFilter } from '../../xlsx/view/contracts'
import { canonicalSearchKey, type ColumnClassification, type FilterContext } from '../../xlsx/view/filter'
import type { DetectedRange } from '../../xlsx/view/rangeDetection'
import type { ValueIndexIdentity } from '../../xlsx/view/valueIndex'
import { VALUE_LIST_LIMITS, type ValuesPage, type WireDetection, type WireFilterSpec, type WireRange, type WireRule } from '../../stream/sheet-wire'
import { sameControlOwner, sameFilterRange, type ControlProgress, type ControlResult, type ControlSnapshot, type ControlsClient, type DraftPage, type DraftQuery, type ValueDraftAction, type ValuesDraftPort } from './controlsClient'
import type { RemoteCommit } from '../../worker/viewerRemoteSource'
import type { ViewerControls } from './viewerControlsClient'
import type { StreamWorkbook } from './streamWorkbook'
import { wireSelection } from './streamText'
import { TYPED_CONDITIONS_UNAVAILABLE } from './FilterDropdown'

/** P16: a value-list search is sent this long after the last keystroke. */
export const STREAM_SEARCH_SETTLE_MS = 250
/** P15: a `stale-filter-values` refusal; the list is fetched again. */
export const STREAM_VALUES_CHANGED = 'Values changed; review the list and try again.'
/** P15: a values rule past 4,096 intervals each way. */
export const STREAM_TOO_MANY_VALUES = 'Too many values selected for this viewing mode.'
/** The dropdown pages its list 200 entries at a time (`FilterDropdown`). */
const DRAFT_PAGE = 200
const LOADING: ColumnClassification = Object.freeze({ kind: 'loading', hasNumeric: false })

const toWireRange = (r: FilterRange): WireRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, header: r.headerRows ? 1 : 0, totals: r.totalsRows ? 1 : 0 })
const toFilterRange = (r: WireRange): FilterRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, headerRows: r.header, totalsRows: r.totals })
const rangeKey = (r: FilterRange, col: number) => JSON.stringify([r.r0, r.r1, r.c0, r.c1, r.headerRows, r.totalsRows, col])
const TYPED_KINDS = new Set(['number', 'date', 'datePreset'])
/** A number or date condition, alone or inside a compound (refused on a protected stream, P16). */
const typedRule = (rule: ColumnRule) => rule.kind === 'compound' ? TYPED_KINDS.has(rule.first.kind) || TYPED_KINDS.has(rule.second.kind) : TYPED_KINDS.has(rule.kind)
/** The local placeholder of a values rule whose selection lives on the server (the view shows the column filtered). */
const valuesPlaceholder = (showBlanks: boolean): ValuesFilter => ({ kind: 'values', selection: { mode: 'queryPatches', base: 'all', patches: [] }, showBlanks })
const abortError = () => Object.assign(new Error('Operation cancelled'), { name: 'AbortError' })
const codeOf = (error: unknown) => (error as { code?: unknown } | null)?.code
const isAbort = (error: unknown) => codeOf(error) === 'cancelled' || (error as { name?: unknown } | null)?.name === 'AbortError'

/** One `'sheet'` op in flight per stream (P16): each waits for the one before it to settle. */
const queues = new WeakMap<StreamWorkbook, { tail: Promise<void> }>()
function sheetOp<T>(stream: StreamWorkbook, signal: AbortSignal | undefined, run: () => Promise<T>): Promise<T> {
  let queue = queues.get(stream)
  if (!queue) queues.set(stream, queue = { tail: Promise.resolve() })
  const turn = queue.tail.then(() => { if (signal?.aborted) throw abortError(); return run() })
  queue.tail = turn.then(() => {}, () => {})
  return turn
}

/** P15: the view's filter after a refetch names its columns only (`filterHead`); the viewer holds a values placeholder
 * per column, which `filter` re-sends as `committed`. Stable per sheet and viewGen. */
const placeholders = new WeakMap<StreamWorkbook, Map<string, { viewGen: number; spec: FilterSpec | null }>>()
export function streamFilterPlaceholder(stream: StreamWorkbook, sheetId: string): FilterSpec | null {
  let held = placeholders.get(stream)
  if (!held) placeholders.set(stream, held = new Map())
  const viewGen = stream.identity(sheetId).viewGen, cached = held.get(sheetId)
  if (cached && cached.viewGen === viewGen) return cached.spec
  const head = stream.filterHead(sheetId)
  const spec = head ? { range: toFilterRange(head.range), columns: new Map(head.columns.map(c => [c.col, valuesPlaceholder(true)] as const)) } : null
  held.set(sheetId, { viewGen, spec })
  return spec
}

export interface StreamControlsOptions {
  stream: StreamWorkbook
  sheetId: string
  snapshot(): ControlSnapshot
  view(): SheetView
  publish(view: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>): ControlSnapshot
  pace?: { setTimer?: typeof setTimeout }
}

/** The state of one value list as the server lists it: entries by id (a position in the unsearched list). */
interface ListState {
  list: string
  total: number
  truncated: boolean
  /** Ids 0 … `loaded - 1` of the unsearched list are held; `cursor` asks for the next page. */
  loaded: number
  cursor?: string
  entries: Map<number, { text: string; selected: boolean; error: boolean }>
  blanks: { present: boolean; selected: boolean }
}
interface SearchState { key: string; text: string; wildcards: boolean; ids: number[]; cursor?: string }

export function streamControls(o: StreamControlsOptions): ViewerControls {
  const { stream, sheetId } = o
  const setTimer = o.pace?.setTimer ?? setTimeout
  let disposed = false
  /** Classifications and prefetched first pages by (range, col), from the last first page of each. */
  const classifications = new Map<string, ColumnClassification>()
  const firstPages = new Map<string, { viewGen: number; page: ValuesPage }>()
  /** The wire rule of each values rule a draft produced (`candidate`), and the drafts' re-list hooks. */
  const drafted = new WeakMap<ColumnRule, WireRule>()
  const relists = new WeakMap<ValuesDraftPort, () => Promise<void>>()
  const context = (): FilterContext => ({ locale: 'en-US', dateSystem: '1900', timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone, nowMs: Date.now() })
  const current = (owner: CacheIdentity) => sameControlOwner(owner, o.snapshot().owner)
  const assert = (owner: CacheIdentity, signal: AbortSignal) => {
    if (disposed || signal.aborted) throw abortError()
    if (!current(owner)) throw new Error('Source or view changed; reopen and retry.')
  }
  const firstPage = (range: FilterRange, col: number, signal: AbortSignal) => sheetOp(stream, signal, () => stream.values(sheetId, { range: toWireRange(range), col }, signal)).then(page => {
    if (page.classification) classifications.set(rangeKey(range, col), page.classification)
    return page
  })

  const detection = (d: WireDetection, owner: CacheIdentity): DetectedRange => ({
    kind: 'complete', range: toFilterRange(d.range), stamp: { docId: owner.docId, sheetId: owner.sheetId, valueGen: owner.valueGen },
    source: d.source === 'table' ? { kind: 'table', id: '' } : { kind: d.source }, header: d.header, overrideIgnored: d.overrideIgnored,
  })
  const resolve: ControlsClient['resolve'] = async (owner, kind, signal, override, col) => {
    assert(owner, signal)
    const selection = o.view().selection
    // A header gesture moves the active cell to its column, and a leftover single cell is no explicit range (S7.1).
    const active = selection.active ? { ...selection.active, col: col ?? selection.active.col } : { row: stream.layout(sheetId).used?.r0 ?? 0, col: col ?? 0 }
    const single = selection.ranges.length === 1 && selection.ranges[0]!.cols.length === 1 && selection.ranges[0]!.cols[0]!.c0 === selection.ranges[0]!.cols[0]!.c1
      && (selection.ranges[0]!.rows.kind === 'intervals' ? selection.ranges[0]!.rows.intervals.length === 1 && selection.ranges[0]!.rows.intervals[0]!.r0 === selection.ranges[0]!.rows.intervals[0]!.r1 : selection.ranges[0]!.rows.kind === 'list' && selection.ranges[0]!.rows.rows.length === 1)
    const wire = wireSelection({ ...selection, active, ranges: col !== undefined && single ? [] : selection.ranges })
    if (wire === 'too-large') return { kind: 'refused', refusal: { reason: 'invalid-range', message: 'Selection too large for this viewing mode.' } }
    let answer: WireDetection
    try {
      answer = await sheetOp(stream, signal, () => stream.detect(sheetId, kind, wire, override === undefined ? undefined : { header: override }, signal))
    } catch (error) {
      const refusal = (error as { refusal?: OperationRefusal }).refusal
      if (codeOf(error) === 'view-refused' && refusal) return { kind: 'refused', refusal }
      if (isAbort(error)) return { kind: 'cancelled' }
      throw error
    }
    assert(owner, signal)
    const detected = detection(answer, owner)
    if (kind === 'filter' && col !== undefined && col >= detected.range.c0 && col <= detected.range.c1) {
      // The dropdown classifies synchronously before it lists: the first page carries the classification.
      try {
        const page = await firstPage(detected.range, col, signal)
        firstPages.set(rangeKey(detected.range, col), { viewGen: owner.viewGen, page })
      } catch (error) { if (isAbort(error)) return { kind: 'cancelled' } }
      assert(owner, signal)
    }
    return detected
  }

  const values: ControlsClient['values'] = async (owner, range, col, applied, signal) => {
    assert(owner, signal)
    const key = rangeKey(range, col)
    let list: ListState | undefined, search: SearchState | undefined, fresh: DraftPage | null = null, closed = false, timer: ReturnType<typeof setTimeout> | undefined
    // P15: `served` reads each id as the server listed it (the column's committed rule); `all`/`none` after Select All.
    let base: 'all' | 'none' | 'served' = applied ? 'served' : 'all', flipped = new Set<number>(), blanks = true, edited = false
    const identity: ValueIndexIdentity = { docId: owner.docId, sheetId: owner.sheetId, valueGen: owner.valueGen, range: { ...range }, col, get classification() { return classifications.get(key) ?? LOADING }, classificationGen: 0 }
    const check = (input: AbortSignal) => { if (closed) throw new Error('Draft closed'); assert(owner, input) }
    const take = (page: ValuesPage, into: ListState) => {
      for (const item of page.items) into.entries.set(item.id, { text: item.text, selected: item.selected, error: item.error })
    }
    const adopt = (page: ValuesPage) => {
      list = { list: page.list, total: page.total, truncated: page.truncated, loaded: page.items.length, ...(page.cursor ? { cursor: page.cursor } : {}), entries: new Map(), blanks: page.blanks }
      take(page, list)
      base = applied ? 'served' : 'all'; flipped = new Set(); edited = false; search = undefined
      blanks = applied ? page.blanks.selected : true
    }
    const load = async (input: AbortSignal) => {
      if (list) return list
      const pre = firstPages.get(key)
      firstPages.delete(key)
      adopt(pre && pre.viewGen === owner.viewGen ? pre.page : await firstPage(range, col, input))
      return list!
    }
    const more = async (cursor: string, input: AbortSignal) => sheetOp(stream, input, () => stream.values(sheetId, { cursor }, input))
    /** Fetches unsearched pages until `want` ids are held (or the list ends). */
    const fill = async (want: number, input: AbortSignal) => {
      const l = await load(input)
      while (l.loaded < Math.min(want, l.total) && l.cursor) {
        const page = await more(l.cursor, input)
        check(input)
        if (page.list !== l.list) throw Object.assign(new Error(STREAM_VALUES_CHANGED), { code: 'stale-filter-values' })
        take(page, l)
        l.loaded += page.items.length
        l.cursor = page.cursor
      }
      return l
    }
    const fillSearch = async (s: SearchState, want: number, input: AbortSignal) => {
      while (s.ids.length < want && s.cursor) {
        const page = await more(s.cursor, input)
        check(input)
        take(page, list!)
        s.ids.push(...page.items.map(item => item.id))
        s.cursor = page.cursor
      }
    }
    const settle = (input: AbortSignal) => new Promise<void>((done, fail) => {
      if (input.aborted) { fail(abortError()); return }
      const onAbort = () => { if (timer !== undefined) clearTimeout(timer); timer = undefined; fail(abortError()) }
      timer = setTimer(() => { timer = undefined; input.removeEventListener('abort', onAbort); done() }, STREAM_SEARCH_SETTLE_MS)
      input.addEventListener('abort', onAbort, { once: true })
    })
    const baseSelected = (id: number) => base === 'all' ? true : base === 'none' ? false : list!.entries.get(id)?.selected ?? false
    const selected = (id: number) => baseSelected(id) !== flipped.has(id)
    const set = (id: number, on: boolean) => { if (on === baseSelected(id)) flipped.delete(id); else flipped.add(id) }
    const listed = () => Math.min(list!.total, VALUE_LIST_LIMITS.entries)
    const count = () => {
      const l = list!, n = listed()
      if (base === 'all') return { count: n - flipped.size, pending: 0, complete: true }
      if (base === 'none') return { count: flipped.size, pending: 0, complete: true }
      let c = 0
      for (let id = 0; id < l.loaded; id++) if (selected(id)) c++
      return { count: c, pending: 0, complete: l.loaded >= n }
    }
    let query: DraftQuery = { text: '', offset: 0 }
    const page = (): DraftPage => {
      const l = list!, ids = search ? search.ids : Array.from({ length: Math.max(0, Math.min(l.loaded, query.offset + DRAFT_PAGE) - query.offset) }, (_, i) => query.offset + i)
      const shown = search ? ids.slice(query.offset, query.offset + DRAFT_PAGE) : ids
      const displayed = search ? search.ids.length + (search.cursor ? 1 : 0) : listed()
      return {
        entries: shown.map(id => { const e = l.entries.get(id)!; return { ordinal: id, kind: e.error ? 'error' as const : 'string' as const, label: e.text, labelTruncated: false, count: 1, selected: selected(id) } }),
        total: listed(), displayed, truncated: l.truncated, count: count(), blanks, complete: true,
      }
    }
    const read = async (q: DraftQuery, input: AbortSignal, progress?: (value: ControlProgress) => void): Promise<DraftPage> => {
      check(input)
      progress?.({ completed: 0, total: 1, message: 'Loading all values' })
      await load(input)
      check(input)
      const text = canonicalSearchKey(q.text, context().locale) ? q.text : ''
      const searchKey = text ? JSON.stringify([text, Boolean(q.wildcards)]) : ''
      if (!searchKey) search = undefined
      else if (search?.key !== searchKey) {
        await settle(input)
        check(input)
        const answer = await sheetOp(stream, input, () => stream.values(sheetId, { range: toWireRange(range), col, search: { text, wildcards: Boolean(q.wildcards) } }, input))
        check(input)
        take(answer, list!)
        search = { key: searchKey, text, wildcards: Boolean(q.wildcards), ids: answer.items.map(item => item.id), ...(answer.cursor ? { cursor: answer.cursor } : {}) }
      }
      if (search) await fillSearch(search, q.offset + DRAFT_PAGE, input)
      else await fill(q.offset + DRAFT_PAGE, input)
      check(input)
      query = q
      return page()
    }
    const relist = async () => {
      list = undefined; search = undefined; query = { text: '', offset: 0 }
      const signal = new AbortController().signal
      await load(signal)
      fresh = page()
    }
    const port: ValuesDraftPort = {
      identity, read,
      async change(action: ValueDraftAction, q: DraftQuery, input: AbortSignal) {
        await read(q, input)
        edited = true
        if (action.kind === 'toggle') set(action.ordinal, action.selected)
        else if (action.kind === 'blanks') blanks = action.selected
        else if (search) {
          // P15: Select All under a search takes every page of that search (at most 10,000) and flips those ids.
          await fillSearch(search, VALUE_LIST_LIMITS.entries, input)
          check(input)
          for (const id of search.ids) set(id, action.selected)
          if (action.includeBlanks) blanks = action.selected
        } else { base = action.selected ? 'all' : 'none'; flipped = new Set(); blanks = action.selected }
        return page()
      },
      async candidate(input) {
        check(input)
        const l = await load(input)
        if (!edited && base === 'served') {
          const rule = valuesPlaceholder(blanks)
          drafted.set(rule, { kind: 'committed' })
          return rule
        }
        // Every id's selection must be known: a served base needs the whole list.
        if (base === 'served') await fill(listed(), input)
        check(input)
        const n = listed(), included: Array<[number, number]> = [], excluded: Array<[number, number]> = []
        let on = 0
        for (let id = 0; id < n; id++) {
          const into = selected(id) ? (on++, included) : excluded, last = into[into.length - 1]
          if (last && last[1] === id - 1) last[1] = id
          else into.push([id, id])
        }
        if (on === n && (blanks || !l.blanks.present)) return null
        // P15: the smaller of the two (fewer intervals, then fewer ids; a tie goes as exclude).
        const smaller = included.length !== excluded.length ? included.length < excluded.length : on < n - on
        const mode = smaller ? 'include' : 'exclude', ids = smaller ? included : excluded
        if (ids.length > VALUE_LIST_LIMITS.intervals) throw new Error(STREAM_TOO_MANY_VALUES)
        const rule = valuesPlaceholder(blanks)
        drafted.set(rule, { kind: 'values', list: l.list, mode, ids, blanks })
        return rule
      },
      async inspect(ordinal, input) {
        check(input)
        return list?.entries.get(ordinal)?.text ?? ''
      },
      relisted() { const out = fresh; fresh = null; return out },
      transfer() {},
      cancel() { closed = true; if (timer !== undefined) clearTimeout(timer); timer = undefined },
    }
    relists.set(port, relist)
    return port
  }

  const progressOf = (progress?: (value: ControlProgress) => void) => progress ? (completed: number, total: number) => progress({ completed, total, message: 'Applying' }) : undefined
  const outcome = async (run: () => Promise<RemoteCommit>, lease?: ValuesDraftPort): Promise<ControlResult> => {
    let result: RemoteCommit
    try { result = await run() } catch (error) {
      if (isAbort(error)) return { kind: 'cancelled' }
      if (codeOf(error) === 'stale-filter-values') {
        // The list the draft named is gone: fetch it again, so the dropdown shows the new one.
        const relist = lease ? relists.get(lease) : undefined
        if (relist) await relist().catch(() => {})
        return { kind: 'refused', refusal: { reason: 'invalid-rule', message: STREAM_VALUES_CHANGED } }
      }
      throw error
    }
    switch (result.kind) {
      case 'committed': return { kind: 'committed', snapshot: o.publish({ order: result.view.order, ...(result.view.visibility ? { visibility: result.view.visibility } : {}), filter: result.view.filter, sort: result.view.sort } as Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>) }
      case 'refused': return { kind: 'refused', refusal: result.refusal }
      case 'error': return { kind: 'refused', refusal: { reason: 'unsupported', message: result.message } }
      default: return { kind: result.kind }
    }
  }

  const filter: ControlsClient['filter'] = async (owner, spec, signal, progress, lease) => {
    if (disposed || signal.aborted) return { kind: 'cancelled' }
    if (!current(owner)) return { kind: 'stale' }
    const before = o.snapshot().filter
    let wire: WireFilterSpec | null = null
    if (spec) {
      const same = Boolean(before && sameFilterRange(before.range, spec.range))
      let changed = !same || before!.columns.size !== spec.columns.size
      const columns: WireFilterSpec['columns'] = []
      for (const [col, rule] of [...spec.columns].sort((a, b) => a[0] - b[0])) {
        let w = drafted.get(rule)
        if (!w && same && before!.columns.get(col) === rule) w = { kind: 'committed' }
        if (!w) {
          if (rule.kind === 'values') return { kind: 'refused', refusal: { reason: 'invalid-rule', message: 'A values filter needs its value list on a streamed sheet.' } }
          if (!stream.typedFilters && typedRule(rule)) return { kind: 'refused', refusal: { reason: 'unsupported', message: TYPED_CONDITIONS_UNAVAILABLE } }
          w = rule
        }
        if (w.kind !== 'committed') changed = true
        columns.push([col, w])
      }
      if (!changed) return { kind: 'unchanged', snapshot: o.snapshot() }
      wire = { range: toWireRange(spec.range), columns }
    } else if (!before) return { kind: 'unchanged', snapshot: o.snapshot() }
    const asked = wire
    return outcome(() => sheetOp(stream, signal, () => stream.filter(stream.identity(sheetId), asked, [], signal, progressOf(progress))), lease)
  }

  const sort: ControlsClient['sort'] = async (owner, spec, signal, progress) => {
    if (disposed || signal.aborted) return { kind: 'cancelled' }
    if (!current(owner)) return { kind: 'stale' }
    if (JSON.stringify(o.snapshot().sort) === JSON.stringify(spec)) return { kind: 'unchanged', snapshot: o.snapshot() }
    return outcome(() => sheetOp(stream, signal, () => stream.sort(stream.identity(sheetId), spec, signal, progressOf(progress))))
  }

  const client: ControlsClient = {
    owner: () => o.snapshot().owner, context, resolve,
    classify(owner, range, col, signal) {
      assert(owner, signal)
      return classifications.get(rangeKey(range, col)) ?? LOADING
    },
    values, filter, sort,
  }
  return { client, dispose() { disposed = true; classifications.clear(); firstPages.clear() } }
}

