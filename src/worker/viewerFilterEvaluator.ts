/**
 * Worker-side filter rule evaluation for worker-owned (paged) sheets (plan
 * S5.4; spec §§7, 10.2). The in-process S9 transport stages filters in a
 * private main-thread session; a paged sheet's view lives in the resident
 * worker, so the same rule semantics run here over the session store: typed and
 * compound rules through FilterEngine, value-list rules through the committed
 * keys the main thread ships with the command (its key store is main-owned).
 *
 * Column values stream in bounded lane pages (4,096 rows); an unknown page waits
 * for the import frontier through ensureComplete, never reads as blank. The
 * matcher is a source-row flag lane over the filter range only.
 */
import type { FilterKeySet, FilterSpec, PagedKeySetRef, SourceStamp, SourceStore, WorkbookFormatting } from '../xlsx/view/contracts'
import { classifyColumn, FilterEngine, type ColumnClassification, type FilterContext, type PreparedValue } from '../xlsx/view/filter'
import { FilterDiagnostic } from '../xlsx/view/textFilter'
import { compileValuesMembership, valueSourceFromStore, type ValueIndexIdentity, type ValueLanePage } from '../xlsx/view/valueIndex'
import { ValueIndexError, type ValueKeyRun, type ValueRunStore } from '../xlsx/view/valueRuns'
import { cellDisplayText } from '../xlsx/render'
import type { XlsxSheet } from '../xlsx/types'
import type { FilterEvaluation, FilterEvaluator } from './viewerSession'
import { semanticsFor } from '../xlsx/view/semanticsFor'

const LANE_PAGE_ROWS = 4096

/** Read-only key store over the keys shipped with one filter command. */
function shippedRuns(keySets: ReadonlyMap<string, FilterKeySet>): ValueRunStore {
  return {
    async put() { throw new ValueIndexError('unsupported', 'Worker filter keys are read-only') },
    async read(ref: PagedKeySetRef, offset: number, limit: number, scope: string): Promise<ValueKeyRun> {
      const set = keySets.get(ref.refId)
      if (!set) throw new ValueIndexError('invalid-rule', 'The value-list keys of this filter were not supplied; reopen the filter and retry.')
      if (set.scope !== scope || set.run.count !== ref.count) throw new ValueIndexError('stale-source', 'Value-list keys belong to another column domain; reopen the filter and retry.')
      const end = Math.min(set.run.count, offset + limit)
      return { count: end - offset, kinds: set.run.kinds.slice(offset, end), numbers: set.run.numbers.slice(offset, end), texts: set.run.texts.slice(offset, end) }
    },
    async release() {},
  }
}

/** Prepared values exactly as the main-thread S9 transport builds them: the
 * workbook style table gives display text, number format and date semantics. */
function preparer(formatting: WorkbookFormatting | undefined): (page: ValueLanePage, index: number) => PreparedValue {
  const sheet = { name: '', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semanticsFor(formatting?.dateSystem ?? '1900') } as XlsxSheet
  const styles = formatting?.styles ?? []
  return (page, index) => {
    const tag = page.tags[index], value = page.numbers[index], style = styles[page.styles?.[index] ?? 0]
    if (tag === 5) return { kind: 'blank' }
    if (tag === 7) return { kind: 'unavailable' }
    if (tag === 4) return { kind: 'error', value: page.texts[index]! }
    if (tag === 2) return { kind: 'string', value: page.texts[index]! }
    const displayText = cellDisplayText(sheet, { value: tag === 3 ? Boolean(value) : value, style })
    return tag === 3 ? { kind: 'boolean', value: Boolean(value), displayText } : { kind: 'number', value, displayText, numFmtId: style?.numFmtId, formatCode: style?.formatCode }
  }
}

/** `keysFor` returns the keys shipped with the command now being evaluated. */
export function createWorkerFilterEvaluator(keysFor: (stamp: SourceStamp) => ReadonlyMap<string, FilterKeySet>, formatting?: WorkbookFormatting, locale = 'en-US'): FilterEvaluator {
  const prepared = preparer(formatting)
  return async ({ store, spec, nowMs, signal, cooperate }): Promise<FilterEvaluation> => {
    const stamp = { ...store.stamp }
    const context: FilterContext = { locale, dateSystem: formatting?.dateSystem ?? '1900', timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone, nowMs }
    const r0 = spec.range.r0 + spec.range.headerRows, r1 = spec.range.r1 - spec.range.totalsRows
    const flags = new Uint8Array(Math.max(0, spec.range.r1 - spec.range.r0 + 1)).fill(1)
    try {
      for (const [col, rule] of spec.columns) {
        const classification = await classify(store, stamp, col, r0, r1, signal, prepared)
        if (classification.kind === 'loading') return { kind: 'incomplete', missing: [{ r0, r1, c0: col, c1: col }] }
        const identity: ValueIndexIdentity = { ...stamp, range: spec.range, col, classification, classificationGen: stamp.valueGen }
        const match = rule.kind === 'values'
          ? (await compileValuesMembership(rule, identity, shippedRuns(keysFor(stamp)), { signal })).matches
          : new FilterEngine(context).compileRule(rule, classification, { signal })
        await pages(store, stamp, col, r0, r1, signal, async (page, first) => {
          for (let i = 0; i < page.count; i++) {
            if (signal.aborted) throw new ValueIndexError('cancelled', 'Filter cancelled')
            if (!match(prepared(page, i))) flags[first + i - spec.range.r0] = 0
            if (cooperate?.due(1)) await cooperate.pause('filter')
          }
        })
      }
    } catch (cause) {
      return outcome(cause, spec)
    }
    return { kind: 'complete', matches: row => row < spec.range.r0 || row > spec.range.r1 || Boolean(flags[row - spec.range.r0]) }
  }
}

async function classify(store: SourceStore, stamp: SourceStamp, col: number, r0: number, r1: number, signal: AbortSignal, prepared: (page: ValueLanePage, index: number) => PreparedValue): Promise<ColumnClassification> {
  let numeric = false, text = false, date = false, loading = false
  await pages(store, stamp, col, r0, r1, signal, page => {
    function* values() { for (let i = 0; i < page.count; i++) yield prepared(page, i) }
    const part = classifyColumn(values(), true)
    numeric ||= part.hasNumeric; text ||= part.kind === 'text'; date ||= part.kind === 'date'; loading ||= part.kind === 'loading'
  })
  return { kind: loading ? 'loading' : text ? 'text' : numeric ? date ? 'date' : 'number' : 'empty', hasNumeric: numeric }
}

async function pages(store: SourceStore, stamp: SourceStamp, col: number, r0: number, r1: number, signal: AbortSignal, accept: (page: ValueLanePage, firstRow: number) => Promise<void> | void): Promise<void> {
  if (r1 < r0) return
  const source = valueSourceFromStore(store, col, { r0, r1 }, LANE_PAGE_ROWS)
  for (let index = 0; index < source.pageCount; index++) {
    let complete = false
    for (;;) {
      if (signal.aborted) throw new ValueIndexError('cancelled', 'Filter cancelled')
      const page = await source.readPage(index, signal)
      if (store.stamp.valueGen !== stamp.valueGen) throw new ValueIndexError('stale-source', 'Source values changed; retry.')
      if (page.kind === 'error') throw new ValueIndexError(page.failure.code === 'cancelled' ? 'cancelled' : 'page-error', page.failure.message)
      if (page.kind === 'values') { await accept(page, r0 + index * LANE_PAGE_ROWS); break }
      if (complete) throw new ValueIndexError('page-error', 'Values stayed unknown after source completion.')
      const range = { r0: r0 + index * LANE_PAGE_ROWS, r1: Math.min(r1, r0 + (index + 1) * LANE_PAGE_ROWS - 1), c0: col, c1: col }
      const coverage = await store.ensureComplete(range, { stamp, signal })
      if (coverage.error) throw new ValueIndexError(coverage.error.code === 'cancelled' ? 'cancelled' : 'page-error', coverage.error.message)
      complete = coverage.completeness === 'complete'
      if (!complete) throw new FilterDiagnostic('incomplete', 'Source values are still loading; retry.')
    }
  }
}

function outcome(cause: unknown, spec: FilterSpec): FilterEvaluation {
  const range = { r0: spec.range.r0, r1: spec.range.r1, c0: spec.range.c0, c1: spec.range.c1 }
  if (cause instanceof FilterDiagnostic) {
    if (cause.code === 'cancelled') return { kind: 'cancelled' }
    if (cause.code === 'incomplete') return { kind: 'incomplete', missing: [range] }
    if (cause.code === 'resource-limit') return { kind: 'refused', refusal: { reason: 'resource-limit', message: cause.message } }
    return { kind: 'refused', refusal: { reason: 'invalid-rule', message: cause.message } }
  }
  if (cause instanceof ValueIndexError) {
    if (cause.code === 'cancelled') return { kind: 'cancelled' }
    if (cause.code === 'stale-source') return { kind: 'error', error: { code: 'stale-source', retryable: true, message: cause.message } }
    if (cause.code === 'quota') return { kind: 'refused', refusal: { reason: 'resource-limit', message: cause.message } }
    if (cause.code === 'invalid-rule' || cause.code === 'unsupported') return { kind: 'refused', refusal: { reason: cause.code === 'unsupported' ? 'unsupported' : 'invalid-rule', message: cause.message } }
    return { kind: 'error', error: { code: 'page-error', retryable: true, message: cause.message } }
  }
  throw cause
}
