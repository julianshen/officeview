// @vitest-environment node
/**
 * The refreshOnLoad refresh path (plan Task 32, spec §7.7): a `refreshOnLoad="1"` pivot whose source is a worksheet
 * range in this workbook is rebuilt from that range on its first interaction, as Excel's refresh does, and compared
 * with Excel's refreshed saves (Task 24: recipes/refresh-seed*.xlsx → interaction/refresh-on-load*.xlsx).
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { loadPivotWorkbook } from '../src/xlsx/pivots/load'
import { decodeRecords, RECORD_LIMITS } from '../src/xlsx/pivots/engine/records'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { recordsFromSource, worksheetSourceRange } from '../src/xlsx/pivots/session/refresh'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { once } from './helpers/xlsx-pivots-load'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const file = (name: string) => resolve(FIXTURES, `${name}.xlsx`)

interface Book { doc: XlsxDocument; def: PivotTableDefinition; cache: PivotCacheDefinition; pkg: OfficePackage; sheet: XlsxSheet; source: (name: string) => XlsxSheet | undefined }
const books = new Map<string, Promise<Book>>()
function book(name: string): Promise<Book> {
  let pending = books.get(name)
  if (!pending) books.set(name, pending = (async () => {
    const pkg = await OfficePackage.load(new Uint8Array(readFileSync(file(name))))
    const doc = await parseXlsx(pkg)
    const workbook = await loadPivotWorkbook(pkg, new Map(doc.sheets.filter(s => s.pivotRegions?.length).map(s => [s.name, s.pivotRegions!])))
    const def = workbook.tables[0]!
    return { doc, def, cache: workbook.caches.get(def.cacheId)!, pkg, sheet: doc.sheets.find(s => s.name === def.sheet)!, source: n => doc.sheets.find(s => s.name === n) }
  })())
  return pending
}
async function savedStore(b: Book) {
  const bytes = await b.pkg.bytes(b.cache.recordsPartPath!)
  const result = await decodeRecords(once(bytes!), b.cache, RECORD_LIMITS.D, false)
  if (!result.ok) throw new Error(result.diagnostic.message)
  return result.store
}
function rebuild(b: Book) {
  const range = worksheetSourceRange(b.cache)!
  const result = recordsFromSource(b.cache, sheetCellReader(b.source(range.sheet)!), range.ref, false, RECORD_LIMITS.D)
  if (!result.ok) throw new Error(`refused ${result.diagnostic.feature}: ${result.diagnostic.message}`)
  return result.value
}

describe('recordsFromSource', () => {
  it('reads the cache source range', async () => {
    const seed = await book('recipes/refresh-seed')
    expect(seed.cache.refreshOnLoad).toBe(true)
    expect(worksheetSourceRange(seed.cache)).toEqual({ sheet: 'Data', ref: { minRow: 0, minCol: 0, maxRow: 56, maxCol: 11 } })
  })

  it.each([
    ['recipes/refresh-seed', 'interaction/refresh-on-load'],
    ['recipes/refresh-seed-nodata', 'interaction/refresh-on-load-nodata'],
  ])('%s rebuilt from its source equals Excel\'s refreshed cache (%s)', async (seedName, caseName) => {
    const seed = await book(seedName), refreshed = await book(caseName)
    const rebuilt = rebuild(seed)
    // Fields, shared items (order, types, flags, min/max) and the record count, exactly as Excel's refresh wrote them.
    expect(rebuilt.cache.fields).toEqual(refreshed.cache.fields)
    expect(rebuilt.cache.recordCount).toBe(refreshed.cache.recordCount)
    // Excel's refresh moved North to the front of Region's shared items (its first record is now North).
    expect(rebuilt.cache.fields[0]!.sharedItems!.items.map(item => item.type === 'string' ? item.value : '')).toEqual(['North', 'East', 'South', 'West', 'Central'])
    // The records equal Excel's refreshed records part (the nodata case has none: compare with the refreshed save of the same source).
    const withRecords = refreshed.cache.recordsPartPath ? refreshed : await book('interaction/refresh-on-load')
    expect(rebuilt.store).toEqual(await savedStore(withRecords))
    // The input cache is untouched.
    expect(seed.cache.fields[0]!.sharedItems!.items[0]).toEqual({ type: 'string', value: 'East' })
  })

  it('refuses a source whose columns no longer match the cache, past the record limits, or with unmeasured values', async () => {
    const seed = await book('recipes/refresh-seed')
    const range = worksheetSourceRange(seed.cache)!
    const cells = sheetCellReader(seed.source('Data')!)
    const feature = (result: { ok: boolean; diagnostic?: { feature: string } }) => result.ok ? 'ok' : result.diagnostic!.feature
    const renamed = (row: number, col: number) => row === 0 && col === 3 ? { kind: 'text' as const, value: 'Revenue' } : cells(row, col)
    expect(feature(recordsFromSource(seed.cache, renamed, range.ref, false, RECORD_LIMITS.D))).toBe('pivot-refresh-on-load')
    expect(feature(recordsFromSource(seed.cache, cells, { ...range.ref, maxCol: 10 }, false, RECORD_LIMITS.D))).toBe('pivot-refresh-on-load')
    expect(feature(recordsFromSource(seed.cache, cells, range.ref, false, { ...RECORD_LIMITS.D, maxRecords: 10 }))).toBe('pivot-records-limit')
    // A boolean value (Excel's refresh of one is not measured).
    const bool = (row: number, col: number) => row === 3 && col === 0 ? { kind: 'bool' as const, value: true } : cells(row, col)
    expect(feature(recordsFromSource(seed.cache, bool, range.ref, false, RECORD_LIMITS.D))).toBe('pivot-refresh-on-load')
    // Text differing only in ASCII case is one value, spelled as first seen (Excel saved Code's "AB" records as "ab").
    const cased = (row: number, col: number) => row === 3 && col === 0 ? { kind: 'text' as const, value: 'north' } : cells(row, col)
    const folded = recordsFromSource(seed.cache, cased, range.ref, false, RECORD_LIMITS.D)
    if (!folded.ok) throw new Error(folded.diagnostic.message)
    expect(folded.value.cache.fields[0]!.sharedItems!.items).toHaveLength(5)
    // Case beyond ASCII is not measured.
    const accented = (row: number, col: number) => row === 3 && col === 8 ? { kind: 'text' as const, value: 'É' } : row === 4 && col === 8 ? { kind: 'text' as const, value: 'é' } : cells(row, col)
    expect(feature(recordsFromSource(seed.cache, accented, range.ref, false, RECORD_LIMITS.D))).toBe('pivot-refresh-on-load')
    // An unreadable source (a paged read that did not complete).
    expect(feature(recordsFromSource(seed.cache, Object.assign(cells, { incomplete: 'read-failed' as const }), range.ref, false, RECORD_LIMITS.D))).toBe('pivot-refresh-on-load')
  })
})

// ---- the session's refresh route --------------------------------------------------------------------------------
import { openPivotSession, type PivotSessionInput } from '../src/xlsx/pivots/session/session'
import { compareWithCase, describeOutcome } from './helpers/xlsx-pivots-oracle'

const EAST = 0, REGION = 0
async function refreshInput(b: Book, over: Partial<PivotSessionInput> = {}): Promise<PivotSessionInput> {
  const range = worksheetSourceRange(b.cache)
  return {
    def: b.def, cache: b.cache, ...(b.cache.recordsPartPath ? { store: await savedStore(b) } : {}), date1904: false,
    saved: sheetCellReader(b.sheet), occupied: () => false, linkedCharts: [],
    ...(range ? { source: sheetCellReader(b.source(range.sheet)!) } : {}), ...over,
  }
}
const featureOf = (result: { ok: boolean; diagnostic?: { feature: string } }) => result.ok ? 'ok' : result.diagnostic!.feature

describe('session refresh route (§7.7)', () => {
  it.each([
    ['recipes/refresh-seed', 'interaction/refresh-on-load', []],
    ['recipes/refresh-seed-nodata', 'interaction/refresh-on-load-nodata', []],
    ['recipes/refresh-seed', 'interaction/refresh-on-load-collapse', [{ kind: 'collapse', field: REGION, item: EAST, collapsed: true }]],
  ] as const)('refreshed cache and layout equal Excel\'s saved output: %s → %s', async (seedName, caseName, deltas) => {
    const seed = await book(seedName), target = await book(caseName)
    const opened = openPivotSession(await refreshInput(seed))
    if (!opened.ok) throw new Error(`open refused ${opened.diagnostic.feature}: ${opened.diagnostic.message}`)
    const session = opened.value
    expect(session.cache.fields).toEqual(target.cache.fields)
    for (const delta of deltas) {
      const applied = session.apply(delta)
      if (!applied.ok) throw new Error(`apply refused ${applied.diagnostic.feature}: ${applied.diagnostic.message}`)
      expect(session.commit(applied.value)).toBe(true)
    }
    expect(describeOutcome(compareWithCase(session, target))).toBe('ok')
  })

  it('an unreadable or external source is not interactive (pivot-refresh-on-load)', async () => {
    const seed = await book('recipes/refresh-seed'), nodata = await book('recipes/refresh-seed-nodata')
    // No source reader: the saved records are older than the source.
    expect(featureOf(openPivotSession(await refreshInput(seed, { source: undefined })))).toBe('pivot-refresh-on-load')
    // A source that is a defined name or table, or in another workbook.
    const named = { ...seed.cache, source: { ...seed.cache.source, worksheet: { name: 'SalesData' } } }
    expect(featureOf(openPivotSession(await refreshInput(seed, { cache: named })))).toBe('pivot-refresh-on-load')
    const external = { ...seed.cache, source: { ...seed.cache.source, worksheet: { ...seed.cache.source.worksheet!, relId: 'rId9' } } }
    expect(featureOf(openPivotSession(await refreshInput(seed, { cache: external })))).toBe('pivot-refresh-on-load')
    // A read that did not complete.
    const cells = sheetCellReader(seed.source('Data')!)
    expect(featureOf(openPivotSession(await refreshInput(seed, { source: Object.assign((row: number, col: number) => cells(row, col), { incomplete: 'read-failed' as const }) })))).toBe('pivot-refresh-on-load')
    // saveData="0" with no readable source: §7.6's missing records apply (CODEX-R2-2).
    expect(featureOf(openPivotSession(await refreshInput(nodata, { source: undefined })))).toBe('pivot-records-absent')
  })

  it('display stays the saved cells in every mode', async () => {
    const seed = await book('recipes/refresh-seed'), target = await book('interaction/refresh-on-load')
    // Load never refreshes: the sheet keeps Excel's stale saved cells (East's 2022 Sales differs after the refresh).
    const cell = (b: Book, row: number, col: number) => sheetCellReader(b.sheet)(row, col)
    expect(cell(seed, 5, 1)).not.toEqual(cell(target, 5, 1))
    expect(seed.sheet.pivotTables![0]!.rowItems).toEqual(seed.def.rowItems)
    // One load diagnostic says Excel would refresh it on open (§7.7); a pivot without refreshOnLoad has none.
    expect(seed.sheet.diagnostics?.filter(d => d.feature === 'pivot-refresh-on-load')).toHaveLength(1)
    expect((await book('engine/baseline')).sheet.diagnostics?.some(d => d.feature === 'pivot-refresh-on-load') ?? false).toBe(false)
    // Opening a session changes nothing shown: the first change is measured from the saved area, and the saved
    // definition and cells are untouched.
    const before = JSON.stringify(seed.sheet.rows)
    const session = openPivotSession(await refreshInput(seed))
    if (!session.ok) throw new Error(session.diagnostic.message)
    expect(JSON.stringify(seed.sheet.rows)).toBe(before)
    const collapsed = session.value.apply({ kind: 'collapse', field: REGION, item: EAST, collapsed: true })
    if (!collapsed.ok) throw new Error(collapsed.diagnostic.message)
    expect(collapsed.value.vacated).toEqual([{ minRow: 21, minCol: 0, maxRow: 23, maxCol: 8 }])
  })

  it('the notice appears once', async () => {
    const seed = await book('recipes/refresh-seed')
    const session = openPivotSession(await refreshInput(seed))
    if (!session.ok) throw new Error(session.diagnostic.message)
    const first = session.value.apply({ kind: 'collapse', field: REGION, item: EAST, collapsed: true })
    if (!first.ok) throw new Error(first.diagnostic.message)
    expect(first.value.notice).toMatch(/refreshed from the workbook's data/)
    // Until a change is committed every candidate carries it; after, none does.
    const again = session.value.apply({ kind: 'collapse', field: REGION, item: EAST, collapsed: true })
    expect(again.ok && again.value.notice).toBe(first.value.notice)
    expect(session.value.commit(first.value)).toBe(true)
    const later = session.value.apply({ kind: 'collapse', field: REGION, item: EAST, collapsed: false })
    expect(later.ok && later.value.notice).toBeUndefined()
    expect(session.value.reset().notice).toBeUndefined()
    // A pivot that does not refresh never shows it.
    const plain = await book('engine/baseline')
    const other = openPivotSession({ def: plain.def, cache: plain.cache, store: await savedStore(plain), date1904: false, saved: sheetCellReader(plain.sheet), occupied: () => false, linkedCharts: [] })
    if (!other.ok) throw new Error(other.diagnostic.message)
    const applied = other.value.apply({ kind: 'collapse', field: REGION, item: EAST, collapsed: true })
    expect(applied.ok && applied.value.notice).toBeUndefined()
  })

  it('the rebuilt store is admitted before the session holds it, and released with it', async () => {
    const seed = await book('recipes/refresh-seed')
    let admitted = 0, released = 0
    const session = openPivotSession(await refreshInput(seed, { admitStore: store => { admitted = store.bytes; return { release: () => { released++ } } } }))
    if (!session.ok) throw new Error(session.diagnostic.message)
    expect(admitted).toBeGreaterThan(0)
    session.value.release()
    session.value.release()
    expect(released).toBe(1)
    expect(featureOf(openPivotSession(await refreshInput(seed, { admitStore: () => undefined })))).toBe('pivot-records-limit')
  })
})

// ---- the resident viewer's client -------------------------------------------------------------------------------
import { createResidentPivotClient, type PivotCommit } from '../src/components/xlsx/pivot/pivotClient'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { pivotRuntimeFor } from '../src/xlsx/pivotLoader'

describe('resident client refresh route', () => {
  it.each([['recipes/refresh-seed'], ['recipes/refresh-seed-nodata']])('%s: refreshed on the first change from the Data sheet, with one notice', async seedName => {
    const pkg = await OfficePackage.load(new Uint8Array(readFileSync(file(seedName))))
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets.findIndex(s => s.pivotTables?.length)
    const installs: PivotCommit[] = []
    const ledger = new MainMemoryLedger()
    const client = createResidentPivotClient({ doc, ledger, install: commit => { installs.push(commit) } })
    const opened = await client.open(sheet, 0, new AbortController().signal)
    if (!opened.ok) throw new Error(`${opened.diagnostic.feature}: ${opened.diagnostic.message}`)
    // Opening shows nothing new: no install until the first change.
    expect(installs).toHaveLength(0)
    // The rebuilt store is charged to the document pool (with the runtime's retained copy).
    expect(ledger.snapshot().byPool.document).toBeGreaterThan(pivotRuntimeFor(doc)!.retainedBytes)
    const applied = await client.apply(sheet, 0, { kind: 'collapse', field: REGION, item: EAST, collapsed: true })
    if (!applied.ok) throw new Error(applied.diagnostic.message)
    expect(applied.notice).toMatch(/refreshed from the workbook's data/)
    // The overlay holds Excel's refreshed-and-collapsed cells.
    const target = await book('interaction/refresh-on-load-collapse')
    const saved = sheetCellReader(target.sheet)
    for (const [row, col] of [[5, 1], [6, 0], [20, 8]] as const) expect(installs[0]!.overlay!.cell(row, col)?.value, `R${row + 1}C${col + 1}`).toEqual((saved(row, col) as { value: unknown }).value)
    const second = await client.apply(sheet, 0, { kind: 'collapse', field: REGION, item: EAST, collapsed: false })
    expect(second.ok && second.notice).toBeUndefined()
    client.dispose?.()
    expect(ledger.snapshot().byPool.document).toBe(pivotRuntimeFor(doc)!.retainedBytes)
  })
})

describe('the refresh notice in the viewer', () => {
  it('the controller announces the notice with the first change only', async () => {
    const { PivotController, PIVOT_APPLIED } = await import('../src/components/xlsx/pivot/PivotButtonLayer')
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file('recipes/refresh-seed')))))
    const sheet = doc.sheets.findIndex(s => s.pivotTables?.length)
    const announced: string[] = []
    const client = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: () => {} })
    const controller = new PivotController({ client, announce: message => { announced.push(message) }, openMenu: () => {}, closeMenu: () => {} })
    expect(await controller.apply(sheet, 0, { kind: 'collapse', field: REGION, item: EAST, collapsed: true })).toBe(false)
    // The controller opens first (as the buttons do).
    expect((await client.open(sheet, 0, new AbortController().signal)).ok).toBe(true)
    expect(await controller.apply(sheet, 0, { kind: 'collapse', field: REGION, item: EAST, collapsed: true })).toBe(true)
    expect(await controller.apply(sheet, 0, { kind: 'collapse', field: REGION, item: EAST, collapsed: false })).toBe(true)
    const changes = announced.filter(message => message.endsWith(PIVOT_APPLIED))
    expect(changes).toHaveLength(2)
    expect(changes[0]).toMatch(/^Pivot table .* was refreshed from the workbook's data/)
    expect(changes[1]).toBe(PIVOT_APPLIED)
    controller.dispose()
    client.dispose?.()
  })
})

// ---- fix round 1 (codex review of Task 32) ----------------------------------------------------------------------
import { setFlagsFromString } from 'node:v8'
import { runInNewContext } from 'node:vm'
setFlagsFromString('--expose_gc')
const collectGarbage = runInNewContext('gc') as () => void
/** Heap growth while `run` executes, garbage included (no collection in between): a proxy for its transient peak. */
async function heapGrowth<T>(run: () => T | Promise<T>): Promise<{ value: T; grown: number }> {
  collectGarbage(); collectGarbage()
  const before = process.memoryUsage().heapUsed
  const value = await run()
  return { value, grown: process.memoryUsage().heapUsed - before }
}
const tall = (cache: PivotCacheDefinition, rows: number): PivotCacheDefinition =>
  ({ ...cache, source: { ...cache.source, worksheet: { ...cache.source.worksheet!, ref: `A1:L${rows + 1}` } } })

describe('refresh fix round 1', () => {
  it('a sparse tall source builds without an unbounded transient matrix, and its construction is admitted as it grows', async () => {
    const seed = await book('recipes/refresh-seed')
    const cache = tall(seed.cache, 100_000), range = worksheetSourceRange(cache)!
    const cells = sheetCellReader(seed.source('Data')!)
    let peak = 0
    const limits = { ...RECORD_LIMITS.D, admitBytes: (bytes: number) => { peak = Math.max(peak, bytes); return true } }
    const { value: result, grown } = await heapGrowth(() => recordsFromSource(cache, cells, range.ref, false, limits))
    if (!result.ok) throw new Error(result.diagnostic.message)
    const stored = result.value.store.bytes
    expect(result.value.store.count).toBe(100_000)
    // Admission covered at least the store; the heap grew by no more than about twice what was admitted (the decoder's
    // doubling arrays), never the ~90 MB boxed-value matrix the review measured for this 12.9 MB store.
    expect(peak).toBeGreaterThanOrEqual(stored)
    expect(grown).toBeLessThan(2 * peak + 8 * 1024 * 1024)
  })

  it('a source too large for the record limits refuses before reading its rows', async () => {
    const seed = await book('recipes/refresh-seed')
    const cache = tall(seed.cache, 3_999_999), range = worksheetSourceRange(cache)!
    const base = sheetCellReader(seed.source('Data')!)
    let reads = 0
    const counted = (row: number, col: number) => { reads++; return base(row, col) }
    // Within D's record count (4,000,000) but not its 64 MiB: the store's least size is known from the shape alone.
    const result = recordsFromSource(cache, counted, range.ref, false, RECORD_LIMITS.D)
    expect(result.ok ? 'ok' : result.diagnostic.feature).toBe('pivot-records-limit')
    expect(reads).toBeLessThanOrEqual(12)
    // An admission refusal stops construction too.
    const refused = recordsFromSource(tall(seed.cache, 100_000), base, worksheetSourceRange(tall(seed.cache, 100_000))!.ref, false, { ...RECORD_LIMITS.D, admitBytes: () => false })
    expect(refused.ok ? 'ok' : refused.diagnostic.feature).toMatch(/^pivot-(records-limit|worker-quota)$/)
  })

  it('an out-of-range date serial in a dates-only field refuses pivot-refresh-on-load instead of throwing', async () => {
    const seed = await book('recipes/refresh-seed')
    const range = worksheetSourceRange(seed.cache)!, cells = sheetCellReader(seed.source('Data')!)
    for (const serial of [1e100, -1, Number.MAX_VALUE, 2958466]) {
      const odd = (row: number, col: number) => row === 3 && col === 5 ? { kind: 'number' as const, value: serial } : cells(row, col)
      const result = recordsFromSource(seed.cache, odd, range.ref, false, RECORD_LIMITS.D)
      expect(result.ok ? 'ok' : result.diagnostic.feature, String(serial)).toBe('pivot-refresh-on-load')
    }
  })

  it('the resident and worker refresh paths charge construction and stay bounded on a sparse tall source', async () => {
    const pkg = await OfficePackage.load(new Uint8Array(readFileSync(file('recipes/refresh-seed'))))
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets.findIndex(s => s.pivotTables?.length)
    const runtimeCache = pivotRuntimeFor(doc)!.cache(doc.sheets[sheet]!.pivotTables![0]!)!
    runtimeCache.source = tall(runtimeCache, 100_000).source
    const ledger = new MainMemoryLedger()
    const sets = vi.spyOn(ledger, 'set')
    const client = createResidentPivotClient({ doc, ledger, install: () => {} })
    const { value: opened, grown } = await heapGrowth(() => client.open(sheet, 0, new AbortController().signal))
    // The construction was charged while it grew, under its own owner, and that charge ended with the open.
    const building = sets.mock.calls.filter(([, owner]) => String(owner).startsWith('pivot-refresh-build:'))
    expect(Math.max(...building.map(([, , bytes]) => bytes))).toBeGreaterThan(1024 * 1024)
    expect(building.at(-1)![2]).toBe(0)
    // Either outcome is bounded: refused by a typed limit, or open with its construction charged as it was built.
    if (opened.ok) {
      const peak = ledger.snapshot().peakBytes
      expect(peak).toBeGreaterThan(pivotRuntimeFor(doc)!.retainedBytes)
      expect(grown).toBeLessThan(3 * peak + 16 * 1024 * 1024)
    } else expect(opened.diagnostic.feature).toMatch(/^pivot-(records-limit|caption-locale|refresh-on-load)$/)
    client.dispose?.()
    expect(ledger.snapshot().byPool.document).toBe(pivotRuntimeFor(doc)!.retainedBytes)
  })

  it('the load diagnostic stays true after the first refresh commit', async () => {
    const seed = await book('recipes/refresh-seed')
    const message = seed.sheet.diagnostics!.find(d => d.feature === 'pivot-refresh-on-load')!.message
    expect(message).not.toMatch(/shows its saved cells/)
    expect(message).toMatch(/refreshed from the workbook's data on its first change/)
  })
})

// ---- fix round 2 (codex re-review of Task 32) -------------------------------------------------------------------
import { refreshedCacheBytes } from '../src/xlsx/pivots/session/refresh'

describe('refresh fix round 2', () => {
  it('an oversized source text is admitted (and refused) before it is case-folded', async () => {
    const seed = await book('recipes/refresh-seed')
    const range = worksheetSourceRange(seed.cache)!, cells = sheetCellReader(seed.source('Data')!)
    const huge = 'A'.repeat(40_000_000)
    const reader = (row: number, col: number) => row === 3 && col === 8 ? { kind: 'text' as const, value: huge } : cells(row, col)
    let admitted = 0
    const limits = { ...RECORD_LIMITS.D, admitBytes: (bytes: number) => { admitted = Math.max(admitted, bytes); return true } }
    const { value: result, grown } = await heapGrowth(() => recordsFromSource(seed.cache, reader, range.ref, false, limits))
    expect(result.ok ? 'ok' : result.diagnostic.feature).toBe('pivot-records-limit')
    // Refused on its reservation: no folded copies of the 40M characters were made.
    expect(grown).toBeLessThan(16 * 1024 * 1024)
    // A text within the limit is admitted before it is kept.
    const long = 'B'.repeat(1_000_000)
    let before = 0
    const fits = recordsFromSource(seed.cache, (row, col) => row === 3 && col === 8 ? { kind: 'text', value: long } : cells(row, col), range.ref, false,
      { ...RECORD_LIMITS.D, admitBytes: bytes => { before = Math.max(before, bytes); return true } })
    expect(fits.ok).toBe(true)
    expect(before).toBeGreaterThanOrEqual(2 * long.length)
  })

  it('the refreshed cache\'s metadata stays charged with the store for the session\'s life', async () => {
    const seed = await book('recipes/refresh-seed')
    const range = worksheetSourceRange(seed.cache)!
    const rebuilt = recordsFromSource(seed.cache, sheetCellReader(seed.source('Data')!), range.ref, false, RECORD_LIMITS.D)
    if (!rebuilt.ok) throw new Error(rebuilt.diagnostic.message)
    const meta = refreshedCacheBytes(rebuilt.value.cache)
    // Every listed item counts (Region 5, Product 4, Year 3).
    expect(meta).toBeGreaterThanOrEqual(12 * 64)
    let admitted: [number, number] | undefined, released = 0
    const session = openPivotSession(await refreshInput(seed, { admitStore: (store, metadataBytes) => { admitted = [store.bytes, metadataBytes]; return { release: () => { released++ } } } }))
    if (!session.ok) throw new Error(session.diagnostic.message)
    expect(admitted).toEqual([rebuilt.value.store.bytes, meta])
    session.value.release()
    expect(released).toBe(1)
    // The resident client charges both, until the session ends.
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file('recipes/refresh-seed')))))
    const sheet = doc.sheets.findIndex(s => s.pivotTables?.length)
    const ledger = new MainMemoryLedger()
    const client = createResidentPivotClient({ doc, ledger, install: () => {} })
    expect((await client.open(sheet, 0, new AbortController().signal)).ok).toBe(true)
    expect(ledger.snapshot().byPool.document).toBe(pivotRuntimeFor(doc)!.retainedBytes + rebuilt.value.store.bytes + meta)
    client.dispose?.()
    expect(ledger.snapshot().byPool.document).toBe(pivotRuntimeFor(doc)!.retainedBytes)
  })
})
