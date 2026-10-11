/**
 * S8.4 kernel scenario runners (dev harness only; runs in a module Worker).
 *
 * Every scenario measures the DELIVERED kernels as they are, over the S0.2
 * deterministic fixture stores, and separates three boundaries (spec §12.6):
 *
 * - kernel: the S7/S8 kernel call alone — sortRowOrder over the session's
 *   prepared store, FilterEngine.projectFilter over a prepared numeric lane,
 *   or a cold ValueIndex build plus its first bounded page;
 * - query completion: the worker DocumentSession operation that publishes an
 *   atomic view (guards, kernel, mask projection, metrics, commit), measured
 *   inside the worker;
 * - input-to-present: measured by the page around the whole round-trip.
 *
 * Nothing here is imported by the library (src/debug is outside the build).
 */
import { createFixture, type FixtureCell, type FixtureHandle } from '../../../tests/helpers/xlsx-viewer-fixtures'
import type { ColumnRule, FilterSpec, SortSpec, SourceStore } from '../../xlsx/view/contracts'
import { sortRowOrder } from '../../xlsx/view/sort'
import { createHostYield } from '../../xlsx/view/scheduler'
import { FilterEngine, classifyColumn, type ColumnClassification, type PreparedValue } from '../../xlsx/view/filter'
import { ValueIndexCache, VALUE_PAGE_LIMIT, valueSourceFromStore, type ValueIndexIdentity } from '../../xlsx/view/valueIndex'
import { createIdentityOrder } from '../../xlsx/view/order'
import { createIntervalSet } from '../../xlsx/view/visibility'
import { READ_ROWS_PAGE_ROWS, buildGeometry } from '../../xlsx/view/source'
import { valueStatus } from '../../xlsx/view/values'
import { DocumentSession, type FilterEvaluator, type WorkbookSource } from '../../worker/viewerSession'
import { initialIdentity } from '../../worker/viewerProtocol'
import type { CacheIdentity } from '../../xlsx/view/contracts'
import { SCENARIOS, type ScenarioDef, type ScenarioId } from './scenarios'

/** Rows shown by the page's stand-in viewport after each query. */
export const VIEWPORT_ROWS = 40

export interface ScenarioInfo {
  id: ScenarioId
  fixture: 'F2' | 'F4'
  rows: number
  cols: number
  bodyRows: number
  description: string
}
export interface KernelSample { ms: number; extras: Record<string, number> }
export interface QuerySample { ms: number; visible: number; total: number; viewport: string[][] }
export interface ScenarioRunner {
  info: ScenarioInfo
  /** One-shot preparation, reported separately (never folded into kernel time). */
  prepare(): Promise<{ metrics: Record<string, number>; notes: string[] }>
  kernel(): Promise<KernelSample>
  query(): Promise<QuerySample>
}

const now = (): number => performance.now()
const signal = (): AbortSignal => new AbortController().signal

/** Body band of a fixture: F2 carries one header and one totals row. */
function bodyRange(def: ScenarioDef, rows: number): SortSpec['range'] {
  return def.fixture === 'F2'
    ? { r0: 0, r1: rows - 1, c0: 0, c1: 7, headerRows: 1, totalsRows: 1 }
    : { r0: 0, r1: rows - 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }
}

/** F2 fields lowest-cardinality first (flag, name, note, date, ratio, amount,
 * code, id): with a unique key first the comparator would never reach lanes 2–8. */
const F2_EIGHT_KEYS = [2, 3, 7, 1, 5, 4, 6, 0]

/** Materializes F4 cells once so repeated sorts measure the kernel, not the
 * fixture's per-read string synthesis (1/8 of F4 strings are 200–700 units). */
function residentF4(handle: FixtureHandle): { store: SourceStore; cells: FixtureCell[]; ms: number } {
  const started = now()
  const cells: FixtureCell[] = new Array(handle.rowCount)
  for (let row = 0; row < handle.rowCount; row++) cells[row] = handle.cellAt(row, 0)!
  const ms = now() - started
  const store = handle.createStore({ read: { cellAt: (row, col) => (col === 0 ? cells[row] ?? null : null) } })
  return { store, cells, ms }
}

/** The S0.2 fixture store bounds readRows by rectangle AREA, while the
 * resident adapters bound it by populated cells (source.ts readRows). The sort
 * guard reads the whole outside-column band, so give the fixture store the
 * adapters' sparse semantics: rows outside the used extent carry no cells and
 * a page covers at most READ_ROWS_PAGE_ROWS rows of the used columns. */
function withSparseRows(store: SourceStore): SourceStore {
  const used = store.extent.used!
  const readRows: SourceStore['readRows'] = async (range, cursor, options) => {
    const c0 = Math.max(range.c0, used.c0), c1 = Math.min(range.c1, used.c1)
    const last = Math.min(range.r1, used.r1)
    const first = cursor === undefined ? Math.max(range.r0, used.r0) : Number(cursor.slice('sparse:'.length))
    if (cursor !== undefined && (!cursor.startsWith('sparse:') || !Number.isSafeInteger(first) || first < range.r0 || first > last)) {
      return { items: [], completeness: 'incomplete', stamp: store.stamp, error: { code: 'invalid-cursor', retryable: false, message: 'Invalid sparse row cursor' } }
    }
    if (c0 > c1 || first > last) return { items: [], completeness: 'complete', stamp: store.stamp }
    const end = Math.min(last, first + READ_ROWS_PAGE_ROWS - 1)
    const page = await store.readRows({ r0: first, r1: end, c0, c1 }, undefined, options)
    if (page.error || page.nextCursor !== undefined) return { ...page, nextCursor: undefined, completeness: 'incomplete' }
    return { ...page, ...(end < last ? { nextCursor: `sparse:${end + 1}` } : {}) }
  }
  return new Proxy(store, { get: (target, key, receiver) => (key === 'readRows' ? readRows : Reflect.get(target, key, receiver)) })
}

function workbookSource(store: SourceStore, name: string): WorkbookSource {
  const geometry = buildGeometry([], [])
  return {
    valuePolicy: store.valuePolicy,
    sheets: Object.freeze([Object.freeze({ sheetId: store.stamp.sheetId, name })]),
    store: () => store,
    geometry: () => geometry,
    coverageDiagnostics: () => [],
    declaredFilters: () => [],
    valueStatus: () => valueStatus(store.valuePolicy, 0),
    release: async () => {},
  }
}

function displayText(value: FixtureCell['value'] | undefined): string {
  if (value === null || value === undefined) return ''
  const text = String(value)
  return text.length > 40 ? `${text.slice(0, 40)}…` : text
}

class SessionHarness {
  readonly session: DocumentSession
  owner: CacheIdentity
  constructor(store: SourceStore, name: string, filterEvaluator?: FilterEvaluator) {
    this.session = new DocumentSession(store.stamp.docId, workbookSource(withSparseRows(store), name), { profile: 'D', locale: 'en', filterEvaluator })
    this.owner = initialIdentity(store.stamp.docId, store.stamp.sheetId)
  }
  get sheetId(): string { return this.owner.sheetId }
}

function viewportOf(view: { order: { displayToSource(row: number): number }; visibility: { select(ordinal: number): number | null } }, cellAt: (row: number, col: number) => FixtureCell | null, cols: number): string[][] {
  const rows: string[][] = []
  for (let ordinal = 0; ordinal < VIEWPORT_ROWS; ordinal++) {
    const slot = view.visibility.select(ordinal)
    if (slot === null) break
    const source = view.order.displayToSource(slot)
    const row: string[] = []
    for (let col = 0; col < cols; col++) row.push(displayText(cellAt(source, col)?.value))
    rows.push(row)
  }
  return rows
}

function sortRunner(def: ScenarioDef, rows: number): ScenarioRunner {
  const handle = createFixture(def.fixture, { rowCount: rows })
  let store: SourceStore, cellAt: (row: number, col: number) => FixtureCell | null, residentMs: number | null = null
  if (def.fixture === 'F4') {
    const resident = residentF4(handle)
    store = resident.store
    cellAt = (row, col) => (col === 0 ? resident.cells[row] ?? null : null)
    residentMs = resident.ms
  } else {
    store = handle.createStore()
    cellAt = (row, col) => handle.cellAt(row, col)
  }
  const keys = def.fixture === 'F4' ? [0] : def.id.endsWith('8key') ? F2_EIGHT_KEYS : [1]
  const spec: SortSpec = Object.freeze({
    range: bodyRange(def, rows),
    keys: Object.freeze(keys.map(col => Object.freeze({ col, order: 'asc' as const }))),
    caseSensitive: false,
  })
  const host = new SessionHarness(store, def.fixture)
  const cols = def.fixture === 'F2' ? 8 : 1
  return {
    info: { id: def.id, fixture: def.fixture, rows, cols, bodyRows: rows - spec.range.headerRows - spec.range.totalsRows, description: def.description },
    async prepare() {
      const metrics: Record<string, number> = {}, notes: string[] = []
      if (residentMs !== null) {
        metrics['resident-source-materialize'] = residentMs
        notes.push('F4 cells are materialized once (resident-source-materialize) so repeated sorts do not re-synthesize fixture strings')
      }
      // The session's prepared-column cache is the "prepared complete columns"
      // boundary. Its D budget may refuse; the session then reads the source.
      const started = now()
      const outcome = await host.session.prepareColumns(host.owner, keys, signal())
      metrics['prepare-columns'] = now() - started
      notes.push(`prepareColumns(${keys.join(',')}) -> ${outcome.kind}${outcome.kind === 'quota' ? ` (required ${outcome.requiredBytes} B > limit ${outcome.limitBytes} B; sorts read the source store)` : ''}`)
      metrics['prepared-columns-admitted'] = outcome.kind === 'complete' ? keys.length : 0
      return { metrics, notes }
    },
    async kernel() {
      const prepared = host.session.preparedStore(host.sheetId)
      // Phase attribution through the kernel's own yield hook, delegating to the
      // same host yield the kernel would create: the work before each
      // pause(phase) is charged to that phase, the remainder to `tail`.
      const hostYield = createHostYield()
      const phases: Record<string, number> = {}
      let mark = 0
      const started = now()
      mark = started
      let result: Awaited<ReturnType<typeof sortRowOrder>>
      try {
        result = await sortRowOrder(prepared, spec, {
          locale: 'en', residentWorkerBytes: host.session.memory.totalBytes, maxWorkerBytes: host.session.memory.limitBytes,
          yield: async (phase) => {
            const at = now()
            phases[phase] = (phases[phase] ?? 0) + (at - mark)
            await hostYield.yield()
            mark = now()
            phases.yielded = (phases.yielded ?? 0) + (mark - at)
          },
        })
      } finally { hostYield.dispose() }
      const ended = now()
      phases.tail = ended - mark
      const ms = ended - started
      if (result.kind !== 'complete') throw new Error(`sortRowOrder did not complete: ${JSON.stringify(result)}`)
      const extras: Record<string, number> = { 'kernel-peak-bytes': result.ledger.peakBytes }
      for (const [phase, value] of Object.entries(phases)) extras[`kernel-phase:${phase}`] = value
      return { ms, extras }
    },
    async query() {
      const started = now()
      const outcome = await host.session.sort(host.owner, spec, signal())
      const ms = now() - started
      if (outcome.kind !== 'committed') throw new Error(`session.sort did not commit: ${JSON.stringify(outcome)}`)
      host.owner = outcome.view.identity
      return { ms, visible: outcome.view.counts.visible, total: outcome.view.counts.total, viewport: viewportOf(outcome.view, cellAt, cols) }
    },
  }
}

const BLANK: PreparedValue = Object.freeze({ kind: 'blank' })
const NUMBER_TAG = 1, STRING_TAG = 2, BOOLEAN_TAG = 3, ERROR_TAG = 4

function filterRunner(def: ScenarioDef, rows: number): ScenarioRunner {
  const handle = createFixture('F2', { rowCount: rows })
  const store = handle.createStore()
  const col = 4
  const range = bodyRange(def, rows)
  // Prepared lane for the rule column: typed tags + numbers, built once.
  const tags = new Uint8Array(rows), numbers = new Float64Array(rows), texts = new Map<number, string>()
  const read = (row: number, at: number): PreparedValue => {
    if (at !== col) throw new RangeError(`unprepared column ${at}`)
    switch (tags[row]) {
      case NUMBER_TAG: return { kind: 'number', value: numbers[row] }
      case BOOLEAN_TAG: return { kind: 'boolean', value: numbers[row] === 1 }
      case STRING_TAG: return { kind: 'string', value: texts.get(row)! }
      case ERROR_TAG: return { kind: 'error', value: texts.get(row)! }
      default: return BLANK
    }
  }
  // Number rules never read the date context; it is supplied for completeness.
  const engine = new FilterEngine({ locale: 'en', dateSystem: '1900', timeZone: 'UTC', nowMs: 0 })
  let columns: ReadonlyMap<number, ColumnClassification> = new Map()
  // About half of the body: amount = row * 1.5.
  const rule: ColumnRule = { kind: 'number', op: 'gt', first: Math.floor(rows / 2) * 1.5 }
  const spec: FilterSpec = Object.freeze({ range, columns: new Map([[col, rule]]) })
  const filterEvaluator: FilterEvaluator = async ({ spec: normalized }) => {
    const matches = engine.compileFilter(normalized, columns, read)
    return { kind: 'complete', matches: row => matches(row) }
  }
  const host = new SessionHarness(store, 'F2', filterEvaluator)
  return {
    info: { id: def.id, fixture: 'F2', rows, cols: 8, bodyRows: rows - 2, description: def.description },
    async prepare() {
      const started = now()
      for (let row = range.r0 + range.headerRows; row <= range.r1 - range.totalsRows; row++) {
        const cell = handle.cellAt(row, col)
        if (cell === null || cell.value === null) continue
        if (cell.valueIsError) { tags[row] = ERROR_TAG; texts.set(row, String(cell.value)) }
        else if (typeof cell.value === 'number') { tags[row] = NUMBER_TAG; numbers[row] = cell.value }
        else if (typeof cell.value === 'boolean') { tags[row] = BOOLEAN_TAG; numbers[row] = cell.value ? 1 : 0 }
        else { tags[row] = STRING_TAG; texts.set(row, cell.value) }
      }
      const laneMs = now() - started
      const classifyStarted = now()
      function* body(): Generator<PreparedValue> {
        for (let row = range.r0 + range.headerRows; row <= range.r1 - range.totalsRows; row++) yield read(row, col)
      }
      columns = new Map([[col, classifyColumn(body(), true)]])
      const classifyMs = now() - classifyStarted
      return {
        metrics: { 'prepare-lane': laneMs, 'classify-column': classifyMs },
        notes: [`rule column ${col} classified ${columns.get(col)!.kind}; lane bytes ${tags.byteLength + numbers.byteLength}`],
      }
    },
    async kernel() {
      const started = now()
      const result = engine.projectFilter(spec, createIntervalSet([]), createIdentityOrder(), columns, read)
      const ms = now() - started
      const visible = result.visibility.rank(result.visibility.logicalCount)
      return { ms, extras: { 'kernel-visible-rows': visible } }
    },
    async query() {
      const started = now()
      const outcome = await host.session.filter(host.owner, spec, Date.now(), signal())
      const ms = now() - started
      if (outcome.kind !== 'committed') throw new Error(`session.filter did not commit: ${JSON.stringify(outcome)}`)
      host.owner = outcome.view.identity
      return { ms, visible: outcome.view.counts.visible, total: outcome.view.counts.total, viewport: viewportOf(outcome.view, (row, at) => handle.cellAt(row, at), 8) }
    },
  }
}

function valueIndexRunner(def: ScenarioDef, rows: number): ScenarioRunner {
  const handle = createFixture('F4', { rowCount: rows })
  let store: SourceStore | null = null
  const range = bodyRange(def, rows)
  const identity: ValueIndexIdentity = {
    ...handle.createStore().stamp, range, col: 0,
    classification: { kind: 'text', hasNumeric: false }, classificationGen: 1,
  }
  async function coldBuild(): Promise<{ ms: number; size: number; page: { label: string }[] }> {
    // A fresh cache per run: no prior prepared index exists (spec §12.2 "cold").
    const started = now()
    const cache = new ValueIndexCache()
    const index = cache.open(identity, valueSourceFromStore(store!, 0, { r0: range.r0, r1: range.r1 }))
    const status = await index.build()
    if (status !== 'complete') throw new Error(`value index build ended ${status}`)
    const page = index.page({ offset: 0, limit: VALUE_PAGE_LIMIT })
    const ms = now() - started
    const size = index.size
    index.dispose()
    return { ms, size, page: page.entries.slice(0, VIEWPORT_ROWS) as { label: string }[] }
  }
  return {
    info: { id: def.id, fixture: 'F4', rows, cols: 1, bodyRows: rows, description: def.description },
    async prepare() {
      const resident = residentF4(handle)
      store = resident.store
      return {
        metrics: { 'resident-source-materialize': resident.ms },
        notes: ['F4 cells are materialized once so every cold build reads the same resident column; the index itself is rebuilt from nothing each run'],
      }
    },
    async kernel() {
      const { ms, size } = await coldBuild()
      return { ms, extras: { 'index-unique-values': size } }
    },
    async query() {
      const { ms, size, page } = await coldBuild()
      return { ms, visible: page.length, total: size, viewport: page.map(entry => [entry.label]) }
    },
  }
}

export function createScenarioRunner(id: ScenarioId, rowsOverride?: number): ScenarioRunner {
  const def = SCENARIOS[id]
  const rows = rowsOverride ?? def.rows
  if (!Number.isSafeInteger(rows) || rows < 3) throw new RangeError(`invalid scenario row count ${rows}`)
  switch (def.kind) {
    case 'sort': return sortRunner(def, rows)
    case 'filter': return filterRunner(def, rows)
    case 'value-index': return valueIndexRunner(def, rows)
  }
}
