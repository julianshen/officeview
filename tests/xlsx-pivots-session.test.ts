// @vitest-environment node
/**
 * The pivot session (plan Task 25, spec §6.4, §7.1, §7.6, §7.7): the pure view deltas, the caption rule, open / apply /
 * commit / reset, and the resident runtime that keeps caches and streams records on demand.
 */
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { readdirSync, readFileSync } from 'node:fs'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { pivotRuntimeFor } from '../src/xlsx/pivotLoader'
import { decodeRecords, RECORD_LIMITS } from '../src/xlsx/pivots/engine/records'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { openPivotSession, ORACLE_VERIFIED_KINDS, type ChartGate, type PivotSession, type PivotSessionInput } from '../src/xlsx/pivots/session/session'
import type { ChartModel } from '../src/charts/model'
import { applyDelta, type PivotDelta } from '../src/xlsx/pivots/session/delta'
import { captionRule, resolveCaptions, savedCaptions } from '../src/xlsx/pivots/session/captions'
import { sheetCellReader, type CellReader } from '../src/xlsx/pivots/engine/faithful'
import type { PivotLayout } from '../src/xlsx/pivots/engine/place'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import { initialViewState, type PivotViewState } from '../src/xlsx/pivots/engine/view'
import type { PivotFilter } from '../src/xlsx/pivots/model'
import { loadAll, once, type Loaded } from './helpers/xlsx-pivots-load'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const ENGINE = resolve(FIXTURES, 'engine')
const cached = new Map<string, Promise<Loaded[]>>()
const load = (name: string) => { const path = resolve(ENGINE, `${name}.xlsx`); let p = cached.get(path); if (!p) cached.set(path, p = loadAll(path)); return p }
const baseline = async () => (await load('baseline'))[0]!
/** Field indices of the engine seed (scripts/xlsx-refs/engine-cases.ts `F`). */
const REGION = 0, PRODUCT = 1, YEAR = 2

const labelFilter = (fld: number, value: string, id = 1): PivotFilter => ({
  fld, type: 'captionBeginsWith', evalOrder: -1, id, stringValue1: value,
  autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, customFilters: { and: false, filters: [{ operator: 'equal', val: `${value}*` }] } }] },
})
const valueFilter = (fld: number, above: number, id = 2): PivotFilter => ({
  fld, type: 'valueGreaterThan', evalOrder: -1, id, iMeasureFld: 0,
  autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, customFilters: { and: false, filters: [{ operator: 'greaterThan', val: String(above) }] } }] },
})
const topFilter = (fld: number, count: number, id = 3): PivotFilter => ({
  fld, type: 'count', evalOrder: -1, id, iMeasureFld: 0,
  autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: count, filterVal: count } }] },
})
const apply = (p: Loaded, view: PivotViewState, delta: PivotDelta) => applyDelta(view, delta, p.def, p.cache)
function ok<T>(result: { ok: true; value: T } | { ok: false; diagnostic: { feature: string; message: string } }): T {
  if (!result.ok) throw new Error(`refused ${result.diagnostic.feature}: ${result.diagnostic.message}`)
  return result.value
}
const refusal = (result: { ok: boolean; diagnostic?: { feature: string } }) => result.ok ? 'ok' : result.diagnostic!.feature

describe('applyDelta maps each kind onto the view', () => {
  it('collapse marks collapsed[item]; collapseField and collapseAll mark every item', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const one = ok(apply(p, view, { kind: 'collapse', field: REGION, item: 0, collapsed: true }))
    expect(one.fields[REGION]!.collapsed).toEqual([true, false, false, false, false, false])
    // Pure: the input view is untouched.
    expect(view.fields[REGION]!.collapsed).toEqual([false, false, false, false, false, false])
    expect(ok(apply(p, one, { kind: 'collapse', field: REGION, item: 0, collapsed: false })).fields[REGION]!.collapsed).toEqual(view.fields[REGION]!.collapsed)
    // A field's data items only: the default-subtotal item keeps its state.
    expect(ok(apply(p, view, { kind: 'collapseField', field: REGION, collapsed: true })).fields[REGION]!.collapsed).toEqual([true, true, true, true, true, false])
    const all = ok(apply(p, view, { kind: 'collapseAll', collapsed: true }))
    expect(all.fields[REGION]!.collapsed).toEqual([true, true, true, true, true, false])
    expect(all.fields[PRODUCT]!.collapsed).toEqual([true, true, true, true, false])
    expect(all.fields[YEAR]!.collapsed).toEqual([true, true, true, false])
    // Fields on no axis are left alone.
    expect(all.fields[3]).toBe(view.fields[3])
  })

  it('collapse refuses an item that is not a data item of the field', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    expect(refusal(apply(p, view, { kind: 'collapse', field: REGION, item: 5, collapsed: true }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'collapse', field: REGION, item: 9, collapsed: true }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'collapse', field: 99, item: 0, collapsed: true }))).toBe('pivot-delta-invalid')
  })

  it('items sets hidden, and refuses hiding every item', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const next = ok(apply(p, view, { kind: 'items', field: REGION, hidden: [3, 4] }))
    expect(next.fields[REGION]!.hidden).toEqual([false, false, false, true, true, false])
    expect(ok(apply(p, next, { kind: 'items', field: REGION, hidden: [] })).fields[REGION]!.hidden).toEqual(view.fields[REGION]!.hidden)
    expect(refusal(apply(p, view, { kind: 'items', field: REGION, hidden: [0, 1, 2, 3, 4] }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'items', field: REGION, hidden: [5] }))).toBe('pivot-delta-invalid')
  })

  it('page sets pages, only on a page field', async () => {
    const p = (await load('quirk-page-single'))[0]!
    const view = initialViewState(p.def, p.cache)
    expect(view.pages).toEqual([{ field: PRODUCT, selected: [0] }])
    expect(ok(apply(p, view, { kind: 'page', field: PRODUCT, selected: 'all' })).pages).toEqual([{ field: PRODUCT, selected: 'all' }])
    expect(ok(apply(p, view, { kind: 'page', field: PRODUCT, selected: [1, 2] })).pages).toEqual([{ field: PRODUCT, selected: [1, 2] }])
    expect(refusal(apply(p, view, { kind: 'page', field: PRODUCT, selected: [] }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'page', field: PRODUCT, selected: [4] }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'page', field: REGION, selected: 'all' }))).toBe('pivot-delta-invalid')
  })

  it('filter replaces that field\'s filter of the same family; another family refuses as the engine does', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const first = ok(apply(p, view, { kind: 'filter', field: PRODUCT, filter: labelFilter(PRODUCT, 'B') }))
    expect(first.filters).toEqual([labelFilter(PRODUCT, 'B')])
    // Same family (label or date; value or top-10): replaced in place.
    const replaced = ok(apply(p, first, { kind: 'filter', field: PRODUCT, filter: labelFilter(PRODUCT, 'C', 7) }))
    expect(replaced.filters).toEqual([labelFilter(PRODUCT, 'C', 7)])
    const value = ok(apply(p, view, { kind: 'filter', field: REGION, filter: valueFilter(REGION, 50000) }))
    expect(ok(apply(p, value, { kind: 'filter', field: REGION, filter: topFilter(REGION, 2) })).filters).toEqual([topFilter(REGION, 2)])
    // Other fields keep their filters.
    expect(ok(apply(p, first, { kind: 'filter', field: REGION, filter: valueFilter(REGION, 50000) })).filters).toEqual([labelFilter(PRODUCT, 'B'), valueFilter(REGION, 50000)])
    // The seed saves multipleFieldFilters="0": a second filter of the other family on one field is the engine's
    // "more than one filter" refusal (pivot-engine-unmeasured), never a silent replace.
    expect(p.def.multipleFieldFilters).toBe(false)
    expect(refusal(apply(p, first, { kind: 'filter', field: PRODUCT, filter: valueFilter(PRODUCT, 1000) }))).toBe('pivot-engine-unmeasured')
    // With multiple filters per field allowed, the engine computes one label and one value filter on a field: both kept.
    const multi = { ...p, def: { ...p.def, multipleFieldFilters: true } }
    const both = ok(apply(multi, first, { kind: 'filter', field: PRODUCT, filter: valueFilter(PRODUCT, 1000) }))
    expect(both.filters).toEqual([labelFilter(PRODUCT, 'B'), valueFilter(PRODUCT, 1000)])
    expect(computePivot(p.store, p.cache, multi.def, both, { date1904: false }).ok).toBe(true)
    // The filter must name the delta's field.
    expect(refusal(apply(p, view, { kind: 'filter', field: REGION, filter: labelFilter(PRODUCT, 'B') }))).toBe('pivot-delta-invalid')
  })

  it('clearFilters removes the field\'s filters and its hidden items', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const filtered = ok(apply(p, ok(apply(p, ok(apply(p, view, { kind: 'filter', field: PRODUCT, filter: labelFilter(PRODUCT, 'B') })),
      { kind: 'filter', field: REGION, filter: valueFilter(REGION, 50000) })), { kind: 'items', field: PRODUCT, hidden: [3] }))
    const cleared = ok(apply(p, filtered, { kind: 'clearFilters', field: PRODUCT }))
    expect(cleared.filters).toEqual([valueFilter(REGION, 50000)])
    expect(cleared.fields[PRODUCT]!.hidden).toEqual(view.fields[PRODUCT]!.hidden)
  })

  it('sort sets sortType and autoSortScope', async () => {
    const p = await baseline()
    const view = initialViewState(p.def, p.cache)
    const label = ok(apply(p, view, { kind: 'sort', field: PRODUCT, order: 'descending' }))
    expect(label.fields[PRODUCT]!.sortType).toBe('descending')
    expect(label.fields[PRODUCT]!.autoSortScope).toBeUndefined()
    const byValue = ok(apply(p, view, { kind: 'sort', field: REGION, order: 'descending', byData: 0 }))
    expect(byValue.fields[REGION]!.sortType).toBe('descending')
    // The scope Excel saves for "sort by Sum of Sales" (sort-value-top-level).
    const saved = (await load('sort-value-top-level'))[0]!.def.fields[REGION]!.autoSortScope
    expect(byValue.fields[REGION]!.autoSortScope).toEqual(saved)
    // A label sort after a value sort drops the scope.
    expect(ok(apply(p, byValue, { kind: 'sort', field: REGION, order: 'ascending' })).fields[REGION]!.autoSortScope).toBeUndefined()
    expect(refusal(apply(p, view, { kind: 'sort', field: REGION, order: 'descending', byData: 2 }))).toBe('pivot-delta-invalid')
    expect(refusal(apply(p, view, { kind: 'sort', field: 3, order: 'descending' }))).toBe('pivot-delta-invalid')
  })
})

/** The saved state laid out without saved captions: its implicit captions keep their slots. */
const savedLayout = (p: Loaded, view: PivotViewState = initialViewState(p.def, p.cache)) => ok(computePivot(p.store, p.cache, p.def, view, { date1904: p.date1904 }))
const textAt = (layout: PivotLayout, row: number, col: number) => { const cell = layout.cells.find(c => c.row === row && c.col === col); return cell?.value.kind === 'text' ? cell.value.value : undefined }
/** Region with its subtotals at the bottom: a view the seed does not save, needing the "<item> Total" kind. */
const bottomSubtotals = (view: PivotViewState): PivotViewState => ({ ...view, fields: view.fields.map((f, k) => k === REGION ? { ...f, subtotalTop: false } : f) })

describe('captions', () => {
  it('captions: zh-TW table from the file', async () => {
    const p = await baseline()
    const layout = savedLayout(p), saved = sheetCellReader(p.sheet)
    // The saved implicit captions ("Sum of Sales 的加總", "Sum of Units 的加總") all render through the table.
    const table = { 'zh-TW': { valuesTotal: '{data} 的加總', subtotal: '{item} 合計' } }
    const rule = captionRule(p.def, saved, layout, table)
    expect(rule.locale).toBe('zh-TW')
    // A kind the file never shows comes from the measured table.
    expect(rule.caption('subtotal', { item: 'East' })).toBe('East 合計')
    expect(ok(resolveCaptions(savedLayout(p, bottomSubtotals(initialViewState(p.def, p.cache))), rule)).captionsUnverified).toBe(false)
    // A table the saved captions contradict is not this pivot's locale.
    expect(captionRule(p.def, saved, layout, { 'zh-TW': { valuesTotal: 'Total {data}' } }).locale).toBe('file')
    // Until Task 24 measures it the shipped table is empty: no pivot matches it.
    expect(captionRule(p.def, saved, layout).locale).toBe('file')
  })

  it('a pivot with no matching table uses its own patterns', async () => {
    const p = await baseline()
    const rule = captionRule(p.def, sheetCellReader(p.sheet), savedLayout(p))
    expect(rule.locale).toBe('file')
    expect(rule.caption('valuesTotal', { data: 'Sum of Units' })).toBe('Sum of Units 的加總')
    expect(rule.caption('subtotal', { item: 'East' })).toBeUndefined()
    // Another Excel save of the same cache shows the bottom subtotal: its own pattern reproduces its saved cells.
    const bottom = (await load('layout-subtotals-bottom-compact'))[0]!
    const own = captionRule(bottom.def, sheetCellReader(bottom.sheet), savedLayout(bottom))
    expect(own.caption('subtotal', { item: 'East' })).toBe('East 合計')
  })

  it('nested implicit arguments render through the same rule', () => {
    const slot = (row: number, kind: 'dataField' | 'valuesTotal', args: object) => ({ row, col: 0, kind, args } as never)
    const layout = {
      rect: { minRow: 0, minCol: 0, maxRow: 1, maxCol: 0 }, cells: [], merges: [], captionsUnverified: true,
      captionSlots: [slot(0, 'dataField', { fn: 'sum', field: 'Sales' }), slot(1, 'valuesTotal', { data: { kind: 'dataField', args: { fn: 'sum', field: 'Sales' } } })],
    } as unknown as PivotLayout
    const texts = ['加總 - Sales', '加總 - Sales 的加總']
    const reader: CellReader = (row, col) => col === 0 && texts[row] ? { kind: 'text', value: texts[row]! } : undefined
    const rule = captionRule({} as never, reader, layout)
    expect(rule.caption('valuesTotal', { data: { kind: 'dataField', args: { fn: 'sum', field: 'Units' } } })).toBe('加總 - Units 的加總')
    // The function is part of the pattern: Count has no sample.
    expect(rule.caption('dataField', { fn: 'count', field: 'Units' })).toBeUndefined()
  })

  it('a needed kind with no pattern refuses pivot-caption-locale', async () => {
    const p = await baseline()
    const rule = captionRule(p.def, sheetCellReader(p.sheet), savedLayout(p))
    const layout = savedLayout(p, bottomSubtotals(initialViewState(p.def, p.cache)))
    expect(layout.captionSlots!.some(slot => slot.kind === 'subtotal')).toBe(true)
    const refused = resolveCaptions(layout, rule)
    expect(refused.ok ? 'ok' : refused.diagnostic.feature).toBe('pivot-caption-locale')
    // The case that saved that layout knows the pattern: the same layout then shows exactly its saved captions.
    const bottom = (await load('layout-subtotals-bottom-compact'))[0]!
    const filled = ok(resolveCaptions(layout, captionRule(bottom.def, sheetCellReader(bottom.sheet), savedLayout(bottom))))
    const saved = sheetCellReader(bottom.sheet)
    for (const slot of layout.captionSlots!) expect(textAt(filled, slot.row, slot.col)).toBe((saved(slot.row, slot.col) as { value: string }).value)
  })

  it('every engine oracle case\'s own patterns reproduce its saved captions', async () => {
    let slots = 0
    for (const name of readdirSync(ENGINE).filter(file => file.endsWith('.xlsx')).sort()) {
      for (const p of await loadAll(resolve(ENGINE, name), 'skip')) {
        const computed = computePivot(p.store, p.cache, p.def, initialViewState(p.def, p.cache), { date1904: p.date1904 })
        if (!computed.ok) continue
        const layout = computed.value, saved = sheetCellReader(p.sheet)
        // Every implicit cell carries its slot.
        expect(layout.captionSlots!.length, name).toBe(layout.cells.filter(cell => cell.caption === 'implicit').length)
        const filled = ok(resolveCaptions(layout, captionRule(p.def, saved, layout)))
        expect(filled, name).toEqual(savedCaptions(layout, saved))
        for (const slot of layout.captionSlots!) { slots++; expect(textAt(filled, slot.row, slot.col), `${name} R${slot.row + 1}C${slot.col + 1}`).toBe((saved(slot.row, slot.col) as { value: string }).value) }
      }
    }
    expect(slots).toBeGreaterThan(100)
  })
})

/** A resident package that records every read of a part's bytes: `bytes`/`text` calls and zip-entry reads. */
async function countingPackage(file: string) {
  const pkg = await OfficePackage.load(new Uint8Array(readFileSync(file)))
  const reads: string[] = []
  const bytes = pkg.bytes.bind(pkg), text = pkg.text.bind(pkg), entry = pkg.entry.bind(pkg)
  pkg.bytes = async (path, max) => { reads.push(`bytes:${path}`); return bytes(path, max) }
  pkg.text = async path => { reads.push(`text:${path}`); return text(path) }
  pkg.entry = path => {
    const found = entry(path)
    if (!found) return found
    const handle = found as unknown as { internalStream(type: string): unknown; async(type: string): Promise<unknown> }
    const stream = handle.internalStream.bind(found), whole = handle.async.bind(found)
    handle.internalStream = type => { reads.push(`stream:${path}`); return stream(type) }
    handle.async = type => { reads.push(`async:${path}`); return whole(type) }
    return found
  }
  return { pkg, reads }
}
const recordsReads = (reads: readonly string[], path: string) => reads.filter(read => read.endsWith(`:${path}`))

describe('runtime', () => {
  it('runtime retains caches and streams records on resident load', async () => {
    const { pkg, reads } = await countingPackage(resolve(ENGINE, 'baseline.xlsx'))
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets.find(s => s.pivotTables?.length)!
    const def = sheet.pivotTables![0]!
    const runtime = pivotRuntimeFor(doc)!
    expect(runtime).toBeDefined()
    // Reachable from the document and from its sheet; never stored on the structured-cloneable sheet.
    expect(pivotRuntimeFor(sheet)).toBe(runtime)
    expect(Object.values(sheet).includes(runtime)).toBe(false)
    const cache = runtime.cache(def)!
    expect(cache.cacheId).toBe(def.cacheId)
    expect(cache.partPath).toBe(def.cachePartPath)
    const path = cache.recordsPartPath!
    // A sheet that never opens a session reads no record bytes.
    expect(recordsReads(reads, path)).toEqual([])
    const ledger = new MainMemoryLedger()
    expect(ledger.snapshot().byPool.document).toBe(0)

    const result = await runtime.records(cache, new AbortController().signal)
    expect(result.ok).toBe(true)
    // Streamed from the zip entry (never pkg.bytes, which would keep the inflated part in the package's cache).
    expect(recordsReads(reads, path)).toEqual([`stream:${path}`])
    const direct = await decodeRecords(once((await (await OfficePackage.load(new Uint8Array(readFileSync(resolve(ENGINE, 'baseline.xlsx'))))).bytes(path))!), cache, RECORD_LIMITS.D, false)
    expect(result).toEqual(direct)
    // records() itself charges nothing; a lease charges the store once while any lease holds it.
    expect(ledger.snapshot().byPool.document).toBe(0)
  })

  it('pivots on one cache share one leased store, charged to the main ledger until the last release', async () => {
    const { pkg, reads } = await countingPackage(resolve(ENGINE, 'baseline.xlsx'))
    const doc = await parseXlsx(pkg)
    const def = doc.sheets.find(s => s.pivotTables?.length)!.pivotTables![0]!
    const runtime = pivotRuntimeFor(doc)!, cache = runtime.cache(def)!, path = cache.recordsPartPath!
    const ledger = new MainMemoryLedger(), signal = new AbortController().signal
    const [a, b] = await Promise.all([runtime.lease(cache, ledger, signal), runtime.lease(cache, ledger, signal)])
    if (!a.ok || !b.ok) throw new Error('lease refused')
    expect(a.lease.store).toBe(b.lease.store)
    expect(recordsReads(reads, path)).toHaveLength(1)
    const bytes = a.lease.store.bytes
    expect(bytes).toBeGreaterThan(0)
    expect(ledger.snapshot().byPool.document).toBe(bytes)
    a.lease.release()
    a.lease.release()
    expect(ledger.snapshot().byPool.document).toBe(bytes)
    b.lease.release()
    expect(ledger.snapshot().byPool.document).toBe(0)
    // After the last release the store is dropped: the next lease decodes again and charges again.
    const c = await runtime.lease(cache, ledger, signal)
    if (!c.ok) throw new Error('lease refused')
    expect(c.lease.store).not.toBe(a.lease.store)
    expect(recordsReads(reads, path)).toHaveLength(2)
    expect(ledger.snapshot().byPool.document).toBe(bytes)
    c.lease.release()
    expect(ledger.snapshot().byPool.document).toBe(0)
  })

  it('records refuse over the cap and without a records part; an aborted lease is cancelled', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(ENGINE, 'baseline.xlsx')))))
    const runtime = pivotRuntimeFor(doc)!, cache = runtime.cache(doc.sheets.find(s => s.pivotTables?.length)!.pivotTables![0]!)!
    const signal = new AbortController().signal
    const capped = await runtime.records(cache, signal, { ...RECORD_LIMITS.M, maxBytes: 64 })
    expect(capped.ok ? 'ok' : capped.diagnostic.feature).toBe('pivot-records-limit')
    const ledger = new MainMemoryLedger()
    const leased = await runtime.lease(cache, ledger, signal, { ...RECORD_LIMITS.M, maxBytes: 64 })
    expect(leased.ok ? 'ok' : leased.diagnostic.feature).toBe('pivot-records-limit')
    expect(ledger.snapshot().byPool.document).toBe(0)
    const absent = await runtime.records({ ...cache, recordsPartPath: undefined }, signal)
    expect(absent.ok ? 'ok' : absent.diagnostic.feature).toBe('pivot-records-absent')
    const aborted = new AbortController()
    aborted.abort()
    await expect(runtime.lease(cache, ledger, aborted.signal)).rejects.toMatchObject({ code: 'cancelled' })
    expect(ledger.snapshot().byPool.document).toBe(0)
  })
})

/** Session input for a loaded pivot: its own saved cells, an empty sheet around it, no linked chart. */
const inputOf = (p: Loaded, over: Partial<PivotSessionInput> = {}): PivotSessionInput => ({
  def: p.def, cache: p.cache, store: p.store, date1904: p.date1904, saved: sheetCellReader(p.sheet), occupied: () => false, linkedCharts: [], ...over,
})
function open(input: PivotSessionInput): PivotSession {
  const session = openPivotSession(input)
  if (!session.ok) throw new Error(`open refused ${session.diagnostic.feature}: ${session.diagnostic.message}`)
  return session.value
}
const openRefusal = (input: PivotSessionInput) => { const r = openPivotSession(input); return r.ok ? 'ok' : r.diagnostic.feature }
const COLLAPSE_EAST: PivotDelta = { kind: 'collapse', field: REGION, item: 0, collapsed: true }
const EXPAND_EAST: PivotDelta = { kind: 'collapse', field: REGION, item: 0, collapsed: false }

describe('session', () => {
  it('open refuses each §7.6 case with its feature', async () => {
    const p = await baseline()
    expect(openRefusal(inputOf(p))).toBe('ok')
    // The records part is absent (saveData="0").
    expect(openRefusal(inputOf(p, { cache: { ...p.cache, saveData: false } }))).toBe('pivot-records-absent')
    // An OLAP cache, or a non-worksheet source.
    expect(openRefusal(inputOf(p, { cache: { ...p.cache, olap: true } }))).toBe('pivot-source-unsupported')
    expect(openRefusal(inputOf(p, { cache: { ...p.cache, source: { ...p.cache.source, type: 'consolidation' } } }))).toBe('pivot-source-unsupported')
    // An unknown construct in the definition.
    expect(openRefusal(inputOf(p, { def: { ...p.def, unsupported: ['unknown-element:x'] } }))).toBe('pivot-definition-unsupported')
    // A calculated field the formula engine cannot evaluate (a cell reference).
    const calc = (await load('calc-constant'))[0]!
    const formula = calc.cache.fields.findIndex(field => field.formula !== undefined)
    expect(openRefusal(inputOf(calc))).toBe('ok')
    expect(openRefusal(inputOf(calc, { cache: { ...calc.cache, fields: calc.cache.fields.map((f, k) => k === formula ? { ...f, formula: 'A1*2' } : f) } }))).toBe('pivot-calculated-field')
    // Calculated items, before they are verified.
    expect(openRefusal(inputOf((await load('calc-item'))[0]!))).toBe('pivot-calculated-item')
    // The faithfulness check fails: one saved value differs.
    const base = sheetCellReader(p.sheet)
    expect(openRefusal(inputOf(p, { saved: (row, col) => row === 5 && col === 1 ? { kind: 'number', value: 1 } : base(row, col) }))).toBe('pivot-faithful-mismatch')
    // ...or cannot run: no saved rowItems.
    expect(openRefusal(inputOf(p, { def: { ...p.def, rowItems: undefined } }))).toBe('pivot-faithful-unverified')
    // The record store over its cap never reaches open: the lease refuses it (runtime test above).
    const runtime = pivotRuntimeFor(await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(resolve(ENGINE, 'baseline.xlsx'))))))!
    const over = await runtime.lease(runtime.cache(p.def)!, new MainMemoryLedger(), new AbortController().signal, { ...RECORD_LIMITS.M, maxRecords: 8 })
    expect(over.ok ? 'ok' : over.diagnostic.feature).toBe('pivot-records-limit')
  })

  it('apply is pure; commit rejects a skipped revision; reset returns the saved layout as the next revision', async () => {
    const p = await baseline()
    const session = open(inputOf(p))
    const saved = { revision: session.revision, view: session.view, def: session.def, layout: session.layout }
    expect(saved.revision).toBe(0)
    expect(saved.def).toBe(p.def)
    // The saved state shows what was saved, its implicit captions included.
    expect(textAt(saved.layout, 3, 7)).toBe('Sum of Sales 的加總')
    const first = ok(session.apply(COLLAPSE_EAST))
    expect(first.revision).toBe(1)
    expect({ revision: session.revision, view: session.view, def: session.def, layout: session.layout }).toEqual(saved)
    expect(first.layout.rect.maxRow).toBe(saved.layout.rect.maxRow - 3)
    expect(first.def.location.ref).toBe('A3:I21')
    // A result two revisions ahead, or one not issued by this session, is refused.
    const skipped = { ...first, revision: 2 }
    expect(session.commit(skipped)).toBe(false)
    expect(session.commit({ ...first })).toBe(false)
    expect(session.commit(first)).toBe(true)
    expect(session.revision).toBe(1)
    expect(session.layout).toBe(first.layout)
    // A result of the old revision is now stale.
    const stale = ok(open(inputOf(p)).apply(COLLAPSE_EAST))
    expect(session.commit(stale)).toBe(false)
    expect(session.commit(first)).toBe(false)
    const reset = session.reset()
    expect(reset.revision).toBe(2)
    expect(reset.layout).toBe(saved.layout)
    expect(reset.view).toBe(saved.view)
    expect(reset.def).toBe(p.def)
    expect(session.revision).toBe(1)
    expect(session.commit(reset)).toBe(true)
    expect(session.revision).toBe(2)
    expect(session.layout).toBe(saved.layout)
  })

  it('vacated covers the cells a collapse frees', async () => {
    const p = await baseline()
    const session = open(inputOf(p))
    const collapsed = ok(session.apply(COLLAPSE_EAST))
    // A3:I24 → A3:I21: rows 22–24 are freed.
    expect(collapsed.vacated).toEqual([{ minRow: 21, minCol: 0, maxRow: 23, maxCol: 8 }])
    expect(session.commit(collapsed)).toBe(true)
    // Growing back frees nothing, and cells of the pivot's own saved area are never an overlap (they hold its saved cells).
    const grown = session.apply(EXPAND_EAST)
    expect(ok(grown).vacated).toEqual([])
    const owned = open(inputOf(p, { occupied: (row, col) => row >= 2 && row <= 23 && col <= 8 }))
    expect(owned.commit(ok(owned.apply(COLLAPSE_EAST)))).toBe(true)
    expect(ok(owned.apply(EXPAND_EAST)).layout.rect.maxRow).toBe(23)
    expect(session.reset().vacated).toEqual([])
  })

  it('a change whose new cells cover a non-empty cell outside the pivot refuses as an overlap', async () => {
    const p = (await load('layout-collapsed-item'))[0]!
    // Saved A3:I21; expanding East grows it to A3:I24.
    const blocked = open(inputOf(p, { occupied: (row, col) => row === 22 && col === 4 }))
    const refused = blocked.apply(EXPAND_EAST)
    expect(refused.ok ? 'ok' : refused.diagnostic.feature).toBe('pivot-overlap')
    expect(blocked.revision).toBe(0)
    expect(ok(open(inputOf(p, { occupied: (row, col) => row === 24 && col === 4 })).apply(EXPAND_EAST)).layout.rect.maxRow).toBe(23)
  })

  it('a pivot with a linked chart and no chart gate refuses pivot-chart-mapping at open, before any apply', async () => {
    const p = await baseline()
    const chart = { pivot: { name: `[baseline.xlsx]Pivot!${p.def.name}` } } as unknown as ChartModel
    expect(openRefusal(inputOf(p, { linkedCharts: [chart] }))).toBe('pivot-chart-mapping')
    // With a gate: open runs it on the saved layout, every apply on the new one; its refilled charts ride on the result.
    const refilled = { ...chart } as ChartModel
    const seen: number[] = []
    const gate: ChartGate = (charts, layout) => {
      seen.push(layout.rect.maxRow)
      expect(charts).toEqual([chart])
      return layout.rect.maxRow >= 23 ? { ok: true, value: [refilled] }
        : { ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-chart-mapping', message: 'no measured mapping for this layout' } }
    }
    const session = open(inputOf(p, { linkedCharts: [chart], chartGate: gate }))
    expect(seen).toEqual([23])
    const refused = session.apply(COLLAPSE_EAST)
    expect(refused.ok ? 'ok' : refused.diagnostic.feature).toBe('pivot-chart-mapping')
    expect(session.revision).toBe(0)
    const sorted = ok(session.apply({ kind: 'sort', field: PRODUCT, order: 'descending' }))
    expect(sorted.charts).toEqual([refilled])
    expect(sorted.charts[0]).toBe(refilled)
  })

  it('apply refuses an interaction kind with no passing oracle, a caption with no rule, and engine refusals', async () => {
    const p = await baseline()
    const session = open(inputOf(p))
    for (const kind of ['collapseAll', 'collapseField', 'clearFilters', 'page'] as const) expect(ORACLE_VERIFIED_KINDS.has(kind), kind).toBe(false)
    const unverified = session.apply({ kind: 'collapseAll', collapsed: true })
    expect(unverified.ok ? 'ok' : unverified.diagnostic.feature).toBe('pivot-interaction-unverified')
    const widened = open(inputOf(p, { verifiedKinds: new Set([...ORACLE_VERIFIED_KINDS, 'collapseAll']) }))
    // Five regions and the grand total under three header rows: A3:…11.
    expect(ok(widened.apply({ kind: 'collapseAll', collapsed: true })).layout.rect.maxRow).toBe(10)
    // The delta's own refusal, then the engine's (a dynamic date filter, CODEX-4).
    const invalid = session.apply({ kind: 'collapse', field: REGION, item: 9, collapsed: true })
    expect(invalid.ok ? 'ok' : invalid.diagnostic.feature).toBe('pivot-delta-invalid')
    const dynamic = session.apply({ kind: 'filter', field: YEAR, filter: { ...valueFilter(YEAR, 0), type: 'thisYear' } })
    expect(dynamic.ok ? 'ok' : dynamic.diagnostic.feature).toBe('pivot-filter-dynamic-date')
    // Captions: quirk-page-single saves no "(All)" caption, so selecting every item cannot be captioned.
    const page = (await load('quirk-page-single'))[0]!
    const pages = open(inputOf(page, { verifiedKinds: new Set([...ORACLE_VERIFIED_KINDS, 'page']) }))
    const all = pages.apply({ kind: 'page', field: PRODUCT, selected: 'all' })
    expect(all.ok ? 'ok' : all.diagnostic.feature).toBe('pivot-caption-locale')
  })

  it('release releases the lease once and closes the session', async () => {
    const p = await baseline()
    let releases = 0
    const session = open(inputOf(p, { lease: { release: () => { releases++ } } }))
    session.release()
    session.release()
    expect(releases).toBe(1)
    const after = session.apply(COLLAPSE_EAST)
    expect(after.ok ? 'ok' : after.diagnostic.feature).toBe('pivot-session-released')
  })
})
