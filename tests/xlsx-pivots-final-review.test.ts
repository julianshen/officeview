// @vitest-environment node
/**
 * Stage 2 final whole-branch review (review-final.md): each finding's repro, run against the Excel-saved oracle cases.
 * Ruling S2-AA: behaviour no Excel case settles refuses with a typed diagnostic; measured behaviour stays ok.
 */
import { describe, expect, it } from 'vitest'
import { resolve } from 'node:path'
import { loadAll, type Loaded } from './helpers/xlsx-pivots-load'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import { checkSavedLayout, sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { initialViewState, type DataFieldState, type FieldState, type PivotViewState } from '../src/xlsx/pivots/engine/view'
import { parseXmlOrdered } from '../src/core/xml'
import { parseCacheDefinition } from '../src/xlsx/pivots/parseCacheDefinition'
import { parsePivotTable } from '../src/xlsx/pivots/parsePivotTable'
import { VALUE_TAG, type RecordStore } from '../src/xlsx/pivots/engine/records'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'

const load = async (name: string): Promise<Loaded> => (await loadAll(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/engine/${name}.xlsx`)))[0]!
const run = (p: Loaded, view: PivotViewState, def: PivotTableDefinition = p.def, cache: PivotCacheDefinition = p.cache, store: RecordStore = p.store) =>
  computePivot(store, cache, def, view, { date1904: p.date1904 })
const refusal = (result: ReturnType<typeof run>) => result.ok ? 'ok' : `${result.diagnostic.feature}: ${result.diagnostic.message}`
const withData = (view: PivotViewState, k: number, patch: Partial<DataFieldState>): PivotViewState =>
  ({ ...view, data: view.data.map((data, index) => index === k ? { ...data, ...patch } : data) })
const withField = (view: PivotViewState, k: number, patch: Partial<FieldState>): PivotViewState =>
  ({ ...view, fields: view.fields.map((field, index) => index === k ? { ...field, ...patch } : field) })

describe('I1: error-valued records', () => {
  it('refuses every function but Sum and Count, named subtotals and show-as over an error record', async () => {
    const p = await load('quirk-error-value')
    const view = initialViewState(p.def, p.cache)
    expect(refusal(run(p, view))).toBe('ok')
    for (const fn of ['average', 'max', 'min', 'product', 'countNums', 'stdDev'] as const)
      expect(refusal(run(p, withData(view, 0, { fn }))), fn).toMatch(/^pivot-engine-unmeasured: .*error-valued records/)
    expect(refusal(run(p, withData(view, 0, { fn: 'count' }))), 'count').toBe('ok')
    // Show-as reading an error cell (the parent of East × 2023 is #N/A).
    expect(refusal(run(p, withData(view, 0, { showAs: 'percentOfParentRow', baseField: 1 })))).toMatch(/^pivot-engine-unmeasured: .*error-valued records/)
    // A Sum field under an Average named subtotal.
    expect(refusal(run(p, withField(view, view.rows[0]!, { subtotals: ['avg'] })))).toMatch(/^pivot-engine-unmeasured: .*error-valued records/)
  })
})

describe('I2 and I3: % of parent total', () => {
  it('refuses a base field on the column axis instead of blank cells', async () => {
    const p = await load('values-showas-percent-of-parent-column')
    const view = initialViewState(p.def, p.cache)
    expect(refusal(run(p, withData(view, 0, { showAs: 'percentOfParent', baseField: 2 })))).toMatch(/^pivot-engine-unmeasured: % of parent total with base field 2 on the columns/)
  })

  it('refuses a cell without a parent line (grand totals off) instead of a blank', async () => {
    const p = await load('values-showas-percent-of-parent-row-two-levels')
    expect(refusal(run(p, initialViewState(p.def, p.cache)))).toBe('ok')
    const def = { ...p.def, rowGrandTotals: false }
    expect(refusal(run(p, initialViewState(def, p.cache), def))).toMatch(/^pivot-engine-unmeasured: % of parent .* has no parent line/)
  })
})

describe('I4: showAll record-less items', () => {
  it('refuses a listed child item that no passing record has', async () => {
    const p = await load('quirk-cross-hidden-item-showall')
    const view = initialViewState(p.def, p.cache)
    const product = (p.store.columns[1] as { index: Uint32Array }).index
    const year = (p.store.columns[2] as { index: Uint32Array }).index.slice()
    for (let r = 0; r < p.store.count; r++) if (product[r] === 2) year[r] = 0   // every Cherries record in 2022
    const columns = p.store.columns.slice(); columns[2] = { kind: 'shared', index: year }
    const hidden2022 = withField(view, 2, { hidden: view.fields[2]!.hidden.map((_, k) => k === 0) })
    expect(refusal(run(p, hidden2022, p.def, p.cache, { ...p.store, columns }))).toMatch(/^pivot-engine-unmeasured: an item without records under showAll lists field 1 item 2/)
  })
})

describe('I5: error and missing captions', () => {
  it('refuses showError or a missing caption, and computes with the defaults Excel writes', async () => {
    const zero = await load('values-showas-percent-diff-zero-base')
    expect(refusal(run(zero, initialViewState(zero.def, zero.cache)))).toBe('ok')
    const showError = { ...zero.def, showError: true, errorCaption: 'n/a' }
    expect(refusal(run(zero, initialViewState(showError, zero.cache), showError))).toMatch(/^pivot-engine-unmeasured: error captions/)
    const baseline = await load('baseline')
    const missing = { ...baseline.def, showMissing: true, missingCaption: '0' }
    expect(refusal(run(baseline, initialViewState(missing, baseline.cache), missing))).toMatch(/^pivot-engine-unmeasured: missing-value captions/)
    const hideMissing = { ...baseline.def, showMissing: false }
    expect(refusal(run(baseline, initialViewState(hideMissing, baseline.cache), hideMissing))).toMatch(/^pivot-engine-unmeasured: missing-value captions/)
  })
})

describe('I6: filters and sorts on a show-as data field', () => {
  it('refuses a value filter and a value sort whose measure has show-as', async () => {
    const filter = await load('filter-value-greater-than')
    const view = initialViewState(filter.def, filter.cache)
    expect(refusal(run(filter, view))).toBe('ok')
    for (const showAs of ['percentOfTotal', 'rankDescending', 'difference'] as const)
      expect(refusal(run(filter, withData(view, 0, { showAs, baseField: 0, baseItem: 1048828 }))), showAs).toMatch(/^pivot-engine-unmeasured: .*show-as/)
    const sort = await load('sort-value-top-level')
    const sortView = initialViewState(sort.def, sort.cache)
    expect(refusal(run(sort, sortView))).toBe('ok')
    for (const showAs of ['rankDescending', 'difference'] as const)
      expect(refusal(run(sort, { ...sortView, data: sortView.data.map(data => ({ ...data, showAs, baseField: 0, baseItem: 1048828 })) })), showAs).toMatch(/^pivot-engine-unmeasured: .*show-as/)
  })
})

describe('I7: show-as sibling scans are bounded', () => {
  it('refuses a show-as mode whose sibling scan exceeds the work cap, and computes the same pivot as normal', () => {
    const N = 2000, records = N * 6
    const items = (n: number, prefix: string) => Array.from({ length: n }, (_, i) => `<s v="${prefix}${i}"/>`).join('')
    const cache = parseCacheDefinition(parseXmlOrdered(`<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="${records}"><cacheSource type="worksheet"/><cacheFields count="3">
<cacheField name="R" databaseField="1"><sharedItems count="${N}">${items(N, 'R')}</sharedItems></cacheField>
<cacheField name="C" databaseField="1"><sharedItems count="12">${items(12, 'C')}</sharedItems></cacheField>
<cacheField name="D" databaseField="1"><sharedItems count="0"/></cacheField></cacheFields></pivotCacheDefinition>`), 1, 'xl/pivotCache/pivotCacheDefinition1.xml', [])
    const store: RecordStore = { count: records, bytes: 0, columns: [
      { kind: 'shared', index: Uint32Array.from({ length: records }, (_, i) => i % N) },
      { kind: 'shared', index: Uint32Array.from({ length: records }, (_, i) => (i * 7) % 12) },
      { kind: 'inline', tag: new Uint8Array(records).fill(VALUE_TAG.number), num: Float64Array.from({ length: records }, (_, i) => (i % 97) + 1), text: [], textIndex: new Uint32Array(records) },
    ] }
    const def = (showAs: string, baseItem: number) => parsePivotTable(parseXmlOrdered(`<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="1" dataCaption="Data">
<location ref="A3:N100" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields count="3"><pivotField axis="axisRow" showAll="0"/><pivotField axis="axisCol" showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields>
<rowFields count="1"><field x="0"/></rowFields><colFields count="1"><field x="1"/></colFields>
<dataFields count="1"><dataField name="S" fld="2" subtotal="sum" showDataAs="${showAs}" baseField="0" baseItem="${baseItem}"/></dataFields></pivotTableDefinition>`),
    { sheet: 'Sheet1', region: { partPath: 'xl/pivotTables/pivotTable1.xml', ref: 'A3:N100' } }, [])!
    const runShape = (showAs: string, baseItem: number) => {
      const d = def(showAs, baseItem)
      return computePivot(store, cache, d, initialViewState(d, cache), { date1904: false })
    }
    expect(runShape('normal', 0).ok).toBe(true)
    expect(runShape('percentOfTotal', 0).ok).toBe(true)
    for (const [showAs, baseItem, mode] of [['difference', 1048828, 'difference'], ['runTotal', 0, 'runTotal'], ['percent', 3, 'percent']] as const) {
      const result = runShape(showAs, baseItem)
      expect(result.ok ? 'ok' : `${result.diagnostic.feature}: ${result.diagnostic.message}`, showAs).toMatch(new RegExp(`^pivot-values-limit: show-as ${mode} .*sibling`))
    }
  })
})

describe('M1: error-valued item labels', () => {
  it('places an error-valued item label as text, as Excel saves it', async () => {
    const p = await load('sort-error-item')
    const def = structuredClone(p.def)
    for (const field of def.fields) field.sortType = 'manual'
    expect(checkSavedLayout(p.store, p.cache, def, sheetCellReader(p.sheet), { date1904: p.date1904 })).toEqual({ status: 'ok' })
    expect(checkSavedLayout(p.store, p.cache, p.def, sheetCellReader(p.sheet), { date1904: p.date1904 })).toEqual({ status: 'ok' })
  })
})

describe('M2: date grouping finer than months', () => {
  it('refuses days, hours, minutes and seconds', async () => {
    const p = await load('group-date-months')
    const field = p.cache.fields.findIndex(f => f.fieldGroup?.rangePr)
    for (const groupBy of ['days', 'hours', 'minutes', 'seconds'] as const) {
      const cache = structuredClone(p.cache)
      const group = cache.fields[field]!.fieldGroup!
      group.rangePr!.groupBy = groupBy
      group.groupItems = [group.groupItems![0]!, ...Array.from({ length: 31 }, (_, k) => ({ type: 'string' as const, value: String(k + 1) })), group.groupItems!.at(-1)!]
      const def = structuredClone(p.def)
      def.fields[field]!.items = [...group.groupItems.map((_, x) => ({ ...def.fields[field]!.items![0]!, type: 'data' as const, x })), ...def.fields[field]!.items!.filter(item => item.type !== 'data')]
      expect(refusal(run(p, initialViewState(def, cache), def, cache)), groupBy).toMatch(new RegExp(`^pivot-engine-unmeasured: date grouping by ${groupBy}`))
    }
  })
})

describe('M3: calculated items on any field', () => {
  it('refuses a calculated item on a field the pivot does not use', async () => {
    const p = await load('baseline')
    const view = initialViewState(p.def, p.cache)
    const used = new Set([...view.rows, ...view.cols, ...view.pages.map(page => page.field)])
    const unused = p.def.fields.findIndex((_, k) => !used.has(k) && !view.data.some(data => data.field === k))
    expect(unused).toBeGreaterThanOrEqual(0)
    const def = structuredClone(p.def)
    def.fields[unused]!.items = [...def.fields[unused]!.items ?? [], { ...p.def.fields[0]!.items![0]!, type: 'data', x: 0, calculated: true }]
    expect(refusal(run(p, initialViewState(def, p.cache), def))).toMatch(/^pivot-calculated-item: /)
  })
})

describe('M4: layout shapes beyond the measured cases', () => {
  it('refuses a collapsed item above Values with more than one unfilled level', async () => {
    const p = await load('layout-collapsed-item-above-values')
    const view = initialViewState(p.def, p.cache)
    expect(refusal(run(p, view))).toBe('ok')
    const deeper = { ...view, rows: [0, 1, 2, -2], cols: [] }
    expect(refusal(run(p, deeper))).toMatch(/^pivot-engine-unmeasured: a collapsed item above Values with 2 unfilled levels/)
  })

  it('refuses mergeItem beyond two tabular row fields and one column field', async () => {
    const p = await load('layout-merge-labels')
    const view = initialViewState(p.def, p.cache)
    expect(refusal(run(p, view))).toBe('ok')
    // Three tabular row fields with totals only after the outermost items: the old gate's line test admits it.
    const threeRows = withField({ ...view, rows: [0, 1, 2], cols: [-2] }, 1, { defaultSubtotal: false })
    expect(refusal(run(p, threeRows))).toMatch(/^pivot-engine-unmeasured: merged item labels/)
  })

  it('refuses the showAll record-less listing in tabular form or below the first row level', async () => {
    const p = await load('quirk-cross-hidden-item-showall')
    const view = initialViewState(p.def, p.cache)
    expect(refusal(run(p, view))).toBe('ok')
    expect(refusal(run(p, withField(view, 0, { outline: false })))).toMatch(/^pivot-engine-unmeasured: a record-less showAll item/)
    expect(refusal(run(p, { ...view, rows: [0, 1, 2], cols: [-2] }))).toMatch(/^pivot-engine-unmeasured: a record-less showAll item/)
  })
})

describe('final review round 2 residuals', () => {
  it('R2-I1: refuses a calculated field over a field with error-valued records (no error read as 0)', async () => {
    const p = await load('quirk-error-value')
    const cache = structuredClone(p.cache), def = structuredClone(p.def)
    cache.fields.push({ ...cache.fields[3]!, name: 'Double', databaseField: false, formula: 'Sales*2', sharedItems: undefined })
    def.fields.push({ ...def.fields[3]!, items: undefined, axis: undefined, dataField: true })
    def.dataFields[0] = { ...def.dataFields[0]!, fld: cache.fields.length - 1, name: 'Sum of Double' }
    expect(refusal(run(p, initialViewState(def, cache), def, cache))).toMatch(/^pivot-engine-unmeasured: calculated data field 0 sums field 3, which has error-valued records/)
  })

  it('R2-M1: refuses the record-less showAll listing when a value or top-10 filter is in the view', async () => {
    for (const name of ['filter-value-nested']) {
      const p = await load(name)
      const view = initialViewState(p.def, p.cache)
      expect(refusal(run(p, view)), name).toBe('ok')
      expect(refusal(run(p, withField(view, view.rows[0]!, { showAll: true }))), name)
        .toMatch(/^pivot-engine-unmeasured: a record-less showAll item under a value or top-10 filter/)
    }
    // A label filter with the same listing is measured (filter-label-cross-field-showall) and stays ok.
    const label = await load('filter-label-cross-field-showall')
    expect(refusal(run(label, initialViewState(label.def, label.cache)))).toBe('ok')
  })
})

describe('PR #56 bot review: calculated slots, overflow and page-cell bounds', () => {
  const numbers = (values: number[]) => ({
    kind: 'inline' as const, tag: new Uint8Array(values.length).fill(VALUE_TAG.number), num: Float64Array.from(values), text: [], textIndex: new Uint32Array(values.length),
  })
  const view = (data: Array<Partial<DataFieldState>>, fields: number) => ({
    rows: [0], cols: [], pages: [], filters: [],
    data: data.map(d => ({ field: 0, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: 'Calc', numFmtId: 0, ...d })),
    fields: Array.from({ length: fields }, () => ({ hidden: [false, false], collapsed: [false, false], sortType: 'manual', defaultSubtotal: false, subtotals: [],
      subtotalTop: false, compact: false, outline: true, insertBlankRow: false, fillDownLabels: false, showAll: false })),
  }) as unknown as PivotViewState
  const def = (fields: number) => ({
    location: { ref: 'A3:B5', firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 1 }, rowGrandTotals: true, colGrandTotals: true,
    fields: Array.from({ length: fields }, () => ({ items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }, { type: 'default' }] })),
    pageFields: [], dataFields: [], unsupported: [], pageWrap: 0, pageOverThenDown: false,
  }) as unknown as PivotTableDefinition

  it('a dependency-free calculated field (=1) over a shared-date row field computes', () => {
    const cache = { source: { type: 'worksheet' }, olap: false, fields: [
      { name: 'When', databaseField: true, sharedItems: { inline: false, items: [{ type: 'date', value: '2024-01-01T00:00:00' }, { type: 'date', value: '2024-02-01T00:00:00' }] } },
      { name: 'One', databaseField: false, formula: '1' },
    ] } as unknown as PivotCacheDefinition
    const store: RecordStore = { count: 2, bytes: 0, columns: [{ kind: 'shared', index: Uint32Array.of(0, 1) }, undefined as never] }
    const result = computePivot(store, cache, def(2), view([{ field: 1 }], 2), { date1904: false })
    expect(result.ok, result.ok ? '' : result.diagnostic.message).toBe(true)
    // Column B holds the values (column A's numbers are the date labels' serials): two items and the grand total.
    if (result.ok) expect(result.value.cells.filter(cell => cell.col === 1 && cell.value.kind === 'number').map(cell => (cell.value as { value: number }).value)).toEqual([1, 1, 1])
  })

  it('percent of total on a calculated field refuses an overflowing share instead of placing Infinity', () => {
    const cache = { source: { type: 'worksheet' }, olap: false, fields: [
      { name: 'Item', databaseField: true, sharedItems: { inline: false, items: [{ type: 'string', value: 'A' }, { type: 'string', value: 'B' }] } },
      { name: 'Amount', databaseField: true, sharedItems: { inline: true, items: [] } },
      { name: 'Calc', databaseField: false, formula: '1/(Amount*10)' },
    ] } as unknown as PivotCacheDefinition
    const store: RecordStore = { count: 2, bytes: 0, columns: [{ kind: 'shared', index: Uint32Array.of(0, 1) }, numbers([1e-308, 1e307]), undefined as never] }
    const result = computePivot(store, cache, def(3), view([{ field: 2, showAs: 'percentOfTotal' }], 3), { date1904: false })
    expect(result.ok ? 'ok' : result.diagnostic.feature).toBe('pivot-engine-unmeasured')
  })

  it('page cells wrapping past column XFD refuse even when the main rectangle fits', async () => {
    const [p] = await loadAll(resolve(process.cwd(), 'tests/fixtures/xlsx-pivots/engine/baseline.xlsx'))
    const base = initialViewState(p!.def, p!.cache)
    const pagesView = { ...base, rows: [], cols: [], pages: [0, 1, 2].map(field => ({ field, selected: 'all' as const })) }
    const near = { ...p!.def, pageWrap: 3, pageOverThenDown: true, location: { ...p!.def.location, ref: 'XFB10:XFC11' } }
    const result = computePivot(p!.store, p!.cache, near, pagesView, { date1904: false })
    expect(result.ok ? 'ok' : result.diagnostic.feature).toBe('pivot-values-limit')
    const room = { ...near, location: { ...near.location, ref: 'B10:C11' } }
    expect(computePivot(p!.store, p!.cache, room, pagesView, { date1904: false }).ok).toBe(true)
  })
})

