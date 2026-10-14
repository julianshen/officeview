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
import type { CacheIdentity, ProfileId } from '../../xlsx/view/contracts'
import { WORKER_BUDGETS } from '../../xlsx/view/tileCache'
import { SCENARIOS, type ScenarioDef, type ScenarioId } from './scenarios'

/** Rows shown by the page's stand-in viewport after each query. */
export const VIEWPORT_ROWS = 40

export interface ScenarioInfo {
  id: ScenarioId
  /** Admission profile whose worker/prepared budgets the session enforces. */
  profile: ProfileId
  /** False: resident harness data is NOT charged to the worker budget, so the
   * run is not an admitted-profile measurement (diagnostic timing only). */
  admitted: boolean
  /** Sort/filter key columns (empty for the value index). */
  keys: readonly number[]
  fixture: 'F2' | 'F4'
  rows: number
  cols: number
  bodyRows: number
  description: string
}
export interface KernelSample { ms: number; extras: Record<string, number> }
export interface QuerySample { ms: number; visible: number; total: number; viewport: string[][]; extras: Record<string, number> }
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

function keysOf(def: ScenarioDef): number[] {
  return [...def.keys]
}

/** Conservative retained bytes of one materialized fixture cell, on the same
 * basis the sort ledger charges strings: raw UTF-16 + 40 B string header and
 * slot, plus a 64 B cell object. */
function residentCellBytes(cell: FixtureCell): number {
  return 64 + (typeof cell.value === 'string' ? cell.value.length * 2 + 40 : 8)
}

/** Materializes F4 cells once so repeated sorts measure the kernel, not the
 * fixture's per-read string synthesis (1/8 of F4 strings are 200–700 units). */
function residentF4(handle: FixtureHandle): { store: SourceStore; cells: FixtureCell[]; ms: number; bytes: number } {
  const started = now()
  const cells: FixtureCell[] = new Array(handle.rowCount)
  let bytes = 0
  for (let row = 0; row < handle.rowCount; row++) {
    const cell = handle.cellAt(row, 0)!
    cells[row] = cell
    bytes += residentCellBytes(cell)
  }
  const ms = now() - started
  const store = handle.createStore({ read: { cellAt: (row, col) => (col === 0 ? cells[row] ?? null : null) } })
  return { store, cells, ms, bytes }
}

export class ResidentBudgetError extends RangeError {
  constructor(readonly requiredBytes: number, readonly limitBytes: number) {
    super(`resident harness data (${requiredBytes} B) does not fit the worker budget (${limitBytes} B)`)
  }
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
  return { ...plainStore(store), readRows }
}

/** A plain-object view of a store. NOT a Proxy: the kernels read `stamp` on
 * every cooperative check (per comparison), and each Proxy layer measurably
 * slows them (a harness artifact, not kernel cost). Fixture stamps are
 * constant for the store's lifetime, so a data property is exact. This
 * kernel-only fixture facade does not expose optional drawing capabilities. */
function plainStore(store: SourceStore): SourceStore {
  return {
    stamp: store.stamp,
    valuePolicy: store.valuePolicy,
    extent: store.extent,
    readCell: (at, options) => store.readCell(at, options),
    readRows: (range, cursor, options) => store.readRows(range, cursor, options),
    queryColumn: (request, options) => store.queryColumn(request, options),
    scanPopulated: (request, options) => store.scanPopulated(request, options),
    queryStructural: (rect, options) => store.queryStructural(rect, options),
    queryOverview: (request, options) => store.queryOverview(request, options),
    ensureComplete: (rect, options) => store.ensureComplete(rect, options),
    release: () => store.release(),
  }
}

/** Resident harness data (materialized source cells, prepared filter lanes)
 * is charged to the document pool's `source` tier, pinned and non-evictable,
 * so kernel admission sees the aggregate worker budget (spec §12.4). */
function workbookSource(store: SourceStore, name: string, residentBytes: () => number): WorkbookSource {
  const geometry = buildGeometry([], [])
  return {
    // Charged, pinned resident cells own their strings for the session lifetime.
    retainsKeyStrings: () => residentBytes() > 0,
    bindMemory(pool) {
      const bytes = residentBytes()
      if (bytes <= 0) return
      const entry = pool.retain(`source:harness:${store.stamp.sheetId}`, { tier: 'source', bytes, owner: store.stamp.sheetId, evictable: false })
      if (!entry) throw new ResidentBudgetError(bytes, pool.limitBytes)
      entry.pin()
    },
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

/** Counts raw source column queries: a genuinely prepared warm sort makes none. */
function countingColumns(store: SourceStore): { store: SourceStore; reads: () => number; reset: () => void } {
  let reads = 0
  const queryColumn: SourceStore['queryColumn'] = (request, options) => { reads++; return store.queryColumn(request, options) }
  return {
    store: { ...plainStore(store), queryColumn },
    reads: () => reads,
    reset: () => { reads = 0 },
  }
}

class SessionHarness {
  readonly session: DocumentSession
  owner: CacheIdentity
  private readonly counter: ReturnType<typeof countingColumns>
  readonly ownsStrings: boolean
  constructor(store: SourceStore, name: string, profile: ProfileId, residentBytes: number, filterEvaluator?: FilterEvaluator) {
    this.counter = countingColumns(withSparseRows(store))
    this.ownsStrings = residentBytes > 0
    this.session = new DocumentSession(store.stamp.docId, workbookSource(this.counter.store, name, () => residentBytes), { profile, locale: 'en', filterEvaluator })
    this.owner = initialIdentity(store.stamp.docId, store.stamp.sheetId)
  }
  get sourceReads(): number { return this.counter.reads() }
  resetReads(): void { this.counter.reset() }
  /** Prepares key columns as one batch; if the batch is refused, stages them
   * one column at a time (a batch's staging peak exceeds the retained
   * capacity). Returns how many keys are prepared; source-read counts in the
   * kernel/query extras verify that none was evicted afterwards. */
  async prepareKeys(keys: readonly number[], notes: string[]): Promise<number> {
    const batch = await this.session.prepareColumns(this.owner, keys, signal())
    notes.push(`prepareColumns(${keys.join(',')}) -> ${batch.kind}${batch.kind === 'quota' ? ` (required ${batch.requiredBytes} B > limit ${batch.limitBytes} B)` : ''}`)
    if (batch.kind === 'complete') return keys.length
    // Staging one column can evict an earlier one (LRU) to fit the column
    // count or bytes, so stop as soon as a staged key is no longer resident.
    const staged: number[] = []
    for (const col of keys) {
      const outcome = await this.session.prepareColumns(this.owner, [col], signal())
      notes.push(`staged prepareColumns(${col}) -> ${outcome.kind}`)
      if (outcome.kind !== 'complete') break
      staged.push(col)
      const evicted: number[] = []
      for (const done of staged) if (!(await this.isPrepared(done))) evicted.push(done)
      if (evicted.length) {
        notes.push(`staging column ${col} evicted prepared column(s) ${evicted.join(',')}`)
        return staged.length - evicted.length
      }
    }
    return staged.length
  }
  /** A one-row prepared-store read touches the raw source only when the column is not prepared. */
  private async isPrepared(col: number): Promise<boolean> {
    const before = this.counter.reads()
    await this.session.preparedStore(this.sheetId).queryColumn(
      { col, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, limit: 1 },
      { signal: signal(), stamp: this.session.preparedStore(this.sheetId).stamp },
    )
    return this.counter.reads() === before
  }
  get sheetId(): string { return this.owner.sheetId }
  /** Aggregate worker pool accounting after an operation (resident + retained + transient peak). */
  poolExtras(): Record<string, number> {
    const stats = this.session.memory.stats()
    return { 'worker-pool-peak-bytes': stats.peakBytes, 'worker-pool-source-bytes': stats.byTier.source }
  }
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

function sortRunner(def: ScenarioDef, rows: number, profile: ProfileId, admitted: boolean): ScenarioRunner {
  const handle = createFixture(def.fixture, { rowCount: rows })
  let store: SourceStore, cellAt: (row: number, col: number) => FixtureCell | null, residentMs: number | null = null, residentBytes = 0
  if (def.fixture === 'F4') {
    const resident = residentF4(handle)
    store = resident.store
    cellAt = (row, col) => (col === 0 ? resident.cells[row] ?? null : null)
    residentMs = resident.ms
    residentBytes = resident.bytes
  } else {
    store = handle.createStore()
    cellAt = (row, col) => handle.cellAt(row, col)
  }
  const keys = keysOf(def)
  const spec: SortSpec = Object.freeze({
    range: bodyRange(def, rows),
    keys: Object.freeze(keys.map(col => Object.freeze({ col, order: 'asc' as const }))),
    caseSensitive: false,
  })
  const host = new SessionHarness(store, def.fixture, profile, admitted ? residentBytes : 0)
  const cols = def.fixture === 'F2' ? 8 : 1
  return {
    info: { id: def.id, profile, admitted, keys, fixture: def.fixture, rows, cols, bodyRows: rows - spec.range.headerRows - spec.range.totalsRows, description: def.description },
    async prepare() {
      const metrics: Record<string, number> = {}, notes: string[] = []
      if (residentMs !== null) {
        metrics['resident-source-materialize'] = residentMs
        metrics['resident-source-bytes'] = residentBytes
        notes.push(admitted
          ? `F4 cells are materialized once (${residentBytes} B, charged to the worker pool source tier) so repeated sorts do not re-synthesize fixture strings`
          : `UNADMITTED run: F4 cells are materialized once (${residentBytes} B) but NOT charged to the worker budget; this is diagnostic timing, not an admitted-profile measurement`)
      }
      // The session's prepared-column cache is the "prepared complete columns"
      // boundary. Keys it cannot hold are read from the source: a FALLBACK
      // diagnostic, not the warm boundary (source reads are recorded).
      const started = now()
      notes.push(`profile ${profile}`)
      const prepared = await host.prepareKeys(keys, notes)
      metrics['prepare-columns'] = now() - started
      metrics['prepared-columns-admitted'] = prepared
      metrics['prepared-retained-bytes'] = host.session.preparedStats().bytes
      metrics['prepared-staging-peak-bytes'] = host.session.memoryStats().peakTransientBytes
      if (prepared < keys.length) notes.push(`only ${prepared} of ${keys.length} key columns prepared: FALLBACK, not the prepared warm boundary`)
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
      host.resetReads()
      try {
        result = await sortRowOrder(prepared, spec, {
          locale: 'en', residentWorkerBytes: host.session.memory.totalBytes, maxWorkerBytes: host.session.memory.limitBytes,
          ...(host.ownsStrings ? { sourceRetainsStrings: true } : {}),
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
      const extras: Record<string, number> = {
        'kernel-peak-bytes': result.ledger.peakBytes, 'kernel-aggregate-peak-bytes': result.ledger.aggregatePeakBytes,
        'kernel-dictionary-bytes': result.ledger.dictionaryBytes, 'kernel-source-column-reads': host.sourceReads,
      }
      for (const [phase, value] of Object.entries(phases)) extras[`kernel-phase:${phase}`] = value
      return { ms, extras }
    },
    async query() {
      host.resetReads()
      const started = now()
      const outcome = await host.session.sort(host.owner, spec, signal())
      const ms = now() - started
      if (outcome.kind !== 'committed') throw new Error(`session.sort did not commit: ${JSON.stringify(outcome)}`)
      host.owner = outcome.view.identity
      return { ms, visible: outcome.view.counts.visible, total: outcome.view.counts.total, viewport: viewportOf(outcome.view, cellAt, cols), extras: { ...host.poolExtras(), 'source-column-reads': host.sourceReads } }
    },
  }
}

const BLANK: PreparedValue = Object.freeze({ kind: 'blank' })
const NUMBER_TAG = 1, STRING_TAG = 2, BOOLEAN_TAG = 3, ERROR_TAG = 4

function filterRunner(def: ScenarioDef, rows: number, profile: ProfileId, admitted: boolean): ScenarioRunner {
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
  // The prepared rule lane (tags + numbers) is resident worker data.
  const host = new SessionHarness(store, 'F2', profile, admitted ? tags.byteLength + numbers.byteLength : 0, filterEvaluator)
  return {
    info: { id: def.id, profile, admitted, keys: [col], fixture: 'F2', rows, cols: 8, bodyRows: rows - 2, description: def.description },
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
        metrics: { 'prepare-lane': laneMs, 'classify-column': classifyMs, 'resident-lane-bytes': tags.byteLength + numbers.byteLength },
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
      return { ms, visible: outcome.view.counts.visible, total: outcome.view.counts.total, viewport: viewportOf(outcome.view, (row, at) => handle.cellAt(row, at), 8), extras: host.poolExtras() }
    },
  }
}

function valueIndexRunner(def: ScenarioDef, rows: number, profile: ProfileId, admitted: boolean): ScenarioRunner {
  const handle = createFixture('F4', { rowCount: rows })
  let store: SourceStore | null = null, residentBytes = 0
  const range = bodyRange(def, rows)
  const identity: ValueIndexIdentity = {
    ...handle.createStore().stamp, range, col: 0,
    classification: { kind: 'text', hasNumeric: false }, classificationGen: 1,
  }
  async function coldBuild(): Promise<{ ms: number; size: number; usedBytes: number; page: { label: string }[] }> {
    // A fresh cache per run: no prior prepared index exists (spec §12.2 "cold").
    // Its budget is the profile's worker pool minus the resident source (§12.4).
    const started = now()
    const cache = new ValueIndexCache({ limitBytes: WORKER_BUDGETS[profile] - (admitted ? residentBytes : 0) })
    const index = cache.open(identity, valueSourceFromStore(store!, 0, { r0: range.r0, r1: range.r1 }))
    const status = await index.build()
    if (status !== 'complete') throw new Error(`value index build ended ${status}`)
    const page = index.page({ offset: 0, limit: VALUE_PAGE_LIMIT })
    const ms = now() - started
    const size = index.size, usedBytes = cache.usedBytes
    index.dispose()
    return { ms, size, usedBytes, page: page.entries.slice(0, VIEWPORT_ROWS) as { label: string }[] }
  }
  return {
    info: { id: def.id, profile, admitted, keys: [], fixture: 'F4', rows, cols: 1, bodyRows: rows, description: def.description },
    async prepare() {
      const resident = residentF4(handle)
      if (admitted && resident.bytes > WORKER_BUDGETS[profile]) throw new ResidentBudgetError(resident.bytes, WORKER_BUDGETS[profile])
      store = resident.store
      residentBytes = resident.bytes
      return {
        metrics: { 'resident-source-materialize': resident.ms, 'resident-source-bytes': resident.bytes },
        notes: [`profile ${profile}: F4 cells are materialized once (${resident.bytes} B, deducted from the value-index cache budget) so every cold build reads the same resident column; the index itself is rebuilt from nothing each run`],
      }
    },
    async kernel() {
      const { ms, size, usedBytes } = await coldBuild()
      return { ms, extras: { 'index-unique-values': size, 'index-used-bytes': usedBytes, 'worker-aggregate-bytes': usedBytes + residentBytes } }
    },
    async query() {
      const { ms, size, usedBytes, page } = await coldBuild()
      return { ms, visible: page.length, total: size, viewport: page.map(entry => [entry.label]), extras: { 'worker-aggregate-bytes': usedBytes + residentBytes } }
    },
  }
}

export function createScenarioRunner(id: ScenarioId, rowsOverride?: number, profile: ProfileId = 'D', admitted = true): ScenarioRunner {
  const def = SCENARIOS[id]
  const rows = rowsOverride ?? def.rows
  if (!Number.isSafeInteger(rows) || rows < 3) throw new RangeError(`invalid scenario row count ${rows}`)
  switch (def.kind) {
    case 'sort': return sortRunner(def, rows, profile, admitted)
    case 'filter': return filterRunner(def, rows, profile, admitted)
    case 'value-index': return valueIndexRunner(def, rows, profile, admitted)
  }
}
