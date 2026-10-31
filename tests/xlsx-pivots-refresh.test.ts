// @vitest-environment node
/**
 * The refreshOnLoad refresh path (plan Task 32, spec §7.7): a `refreshOnLoad="1"` pivot whose source is a worksheet
 * range in this workbook is rebuilt from that range on its first interaction, as Excel's refresh does, and compared
 * with Excel's refreshed saves (Task 24: recipes/refresh-seed*.xlsx → interaction/refresh-on-load*.xlsx).
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
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
