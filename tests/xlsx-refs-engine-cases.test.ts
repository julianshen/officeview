import { describe, expect, it } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { OfficePackage } from '../src/core/zip'
import {
  ENGINE_CASES, ENGINE_ROWS, ERROR_ROW, F, SOURCE_LAST_ROW, engineSeedStandIn, fieldAttr, generateEngineRows, rowField, seedDataLines,
  totals, xmlDiffSummary, type EngineRow, type PivotParts,
} from '../scripts/xlsx-refs/engine-cases'

const RECIPES = join(__dirname, 'fixtures/xlsx-pivots/recipes')
const ENGINE_SEED = join(RECIPES, 'engine-seed.xlsx')

async function parts(file: string): Promise<PivotParts> {
  const pkg = await OfficePackage.load(readFileSync(file))
  return { pivot: (await pkg.text('xl/pivotTables/pivotTable1.xml'))!, cache: (await pkg.text('xl/pivotCache/pivotCacheDefinition1.xml'))! }
}
/** The engine seed's parts: Excel's own once `build-engine.ts --seed` has run, else the stand-in derived from the
 * real pivot-seed.xlsx parts. */
async function seedParts(): Promise<PivotParts> {
  return existsSync(ENGINE_SEED) ? parts(ENGINE_SEED) : engineSeedStandIn(await parts(join(RECIPES, 'pivot-seed.xlsx')))
}
function parse(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  const error = doc.getElementsByTagName('parsererror')[0]
  if (error) throw new Error(`not well-formed: ${error.textContent}`)
  return doc
}
const kids = (el: Element, name: string) => Array.from(el.children).filter(c => c.localName === name)
const one = (doc: Document, name: string) => doc.getElementsByTagName(name)[0]
/** Every `count` the edits maintain agrees with its children, and pivot fields line up with cache fields. */
function expectConsistent(p: PivotParts): void {
  const pivot = parse(p.pivot), cache = parse(p.cache)
  const cacheFields = kids(one(cache, 'cacheFields')!, 'cacheField')
  expect(Number(one(cache, 'cacheFields')!.getAttribute('count'))).toBe(cacheFields.length)
  const pivotFields = kids(one(pivot, 'pivotFields')!, 'pivotField')
  expect(Number(one(pivot, 'pivotFields')!.getAttribute('count'))).toBe(pivotFields.length)
  expect(pivotFields.length).toBe(cacheFields.length)
  for (const f of pivotFields) for (const items of kids(f, 'items')) expect(Number(items.getAttribute('count'))).toBe(kids(items, 'item').length)
  for (const name of ['rowFields', 'colFields', 'dataFields', 'filters', 'pageFields']) {
    const el = one(pivot, name)
    if (el) expect(Number(el.getAttribute('count')), name).toBe(el.children.length)
  }
  for (const shared of Array.from(cache.getElementsByTagName('sharedItems'))) {
    if (shared.getAttribute('count') !== null) expect(Number(shared.getAttribute('count'))).toBe(shared.children.length)
  }
  // Each data field and each pivot item points inside its field.
  for (const d of Array.from(pivot.getElementsByTagName('dataField'))) expect(Number(d.getAttribute('fld'))).toBeLessThan(cacheFields.length)
}

describe('engine seed source', () => {
  const rows = ENGINE_ROWS
  const sum = (key: (r: EngineRow) => string, value: (r: EngineRow) => number | null) => totals(rows, key, value)
  it('is the literal output of the seeded generator', () => {
    expect(rows).toEqual(generateEngineRows())
    expect(SOURCE_LAST_ROW).toBe(rows.length + 1)
    expect(ERROR_ROW).toBe(SOURCE_LAST_ROW + 1)
  })
  it('keeps Central to Elderberries on Retail, and Elderberries to Central', () => {
    for (const r of rows) {
      expect(r[0] === 'Central').toBe(r[1] === 'Elderberries')
      if (r[0] === 'Central') expect(r[6]).toBe('Retail')
    }
  })
  it('engineers the Units zeros and ties, and keeps Sales aggregates distinct and non-zero', () => {
    const units = sum(r => r[0], r => r[4])
    expect(units.get('Central')).toBe(0)
    expect(sum(r => `${r[0]}/${r[1]}/${r[2]}`, r => r[4]).get('South/Cherries/2022')).toBe(0)
    expect(units.get('North')).toBe(units.get('South'))
    expect(units.get('East')!).toBeGreaterThan(units.get('North')!)
    expect(units.get('West')!).toBeLessThan(units.get('North')!)
    const regionSales = [...sum(r => r[0], r => r[3]).values()]
    expect(new Set(regionSales).size).toBe(regionSales.length)
    expect(regionSales.every(v => v > 0)).toBe(true)
    for (const region of ['East', 'North', 'South', 'West']) {
      const products = [...sum(r => r[0] === region ? r[1] : '', r => r[0] === region ? r[3] : null).values()]
      expect(new Set(products).size).toBe(products.length)
    }
  })
  it('has the Sales, Date, Channel, Net Amount and Stamp quirks the cases rely on', () => {
    const sales = rows.map(r => r[3])
    expect(sales.filter(v => v === null)).toHaveLength(3)
    expect(sales.filter(v => v !== null && v < 0).length).toBeGreaterThanOrEqual(3)
    expect(sales.filter(v => v === 1234.5)).toHaveLength(2)
    const dates = rows.map(r => r[5])
    expect(dates.filter(v => v === null)).toHaveLength(4)
    const years = new Set(dates.filter((v): v is number => v !== null).map(v => new Date(Date.UTC(1899, 11, 30) + v * 864e5).getUTCFullYear()))
    expect([...years].sort()).toEqual([2019, 2020, 2021, 2022, 2023, 2024])
    expect(new Set(rows.map(r => r[6]))).toEqual(new Set(['Online', 'Retail', 42, null]))
    expect(rows.filter(r => r[0] === 'West' && r[1] === 'Apples' && r[2] === 2023 && r[7] !== null).map(r => r[7])).toEqual([0.1, 0.2, 0.3])
    expect(rows.map(r => r[11]).filter(v => v !== null)).toEqual([0.5, 1, 59, 60, 61, 45000.25])
    expect(new Set(rows.map(r => r[8]))).toEqual(new Set(['A 1', 'A1', 'a-b', 'a_b', '$x', '@x', '#x', '-x', "a'b", 'ab', 'a b', 'AB', 'x', '_x', '~x']))
  })
  it('writes the whole table, the error formula and the blanks through AppleScript', () => {
    const lines = seedDataLines('worksheet "Data" of workbook "w"')
    const write = lines.find(l => l.startsWith('set value of range "A1:L'))!
    expect(write).toContain(`"A1:L${ERROR_ROW}"`)
    expect(write.match(/\{"/g)).toHaveLength(rows.length + 2)
    expect(lines).toContain(`set formula of range "D${ERROR_ROW}" of worksheet "Data" of workbook "w" to "=NA()"`)
    const blanks = rows.reduce((n, r) => n + r.filter(v => v === null).length, 0) + 2   // the error record: Net Amount, Stamp
    const cleared = lines.filter(l => l.startsWith('clear range (range "') && !l.includes('A1:Z500')).map(l => /"([A-L])(\d+)(?::[A-L](\d+))?"/.exec(l)!)
    expect(cleared.reduce((n, m) => n + (m[3] ? Number(m[3]) - Number(m[2]) + 1 : 1), 0)).toBe(blanks)
    expect(cleared.some(m => m[1] === 'D' && Number(m[2]) <= ERROR_ROW && Number(m[3] ?? m[2]) >= ERROR_ROW)).toBe(false)
  })
})

describe('engine case manifest', () => {
  const ids = ENGINE_CASES.map(c => c.id)
  const byId = new Map(ENGINE_CASES.map(c => [c.id, c]))
  it('has unique ids and a buildable route for every case', () => {
    expect(new Set(ids).size).toBe(ids.length)
    for (const c of ENGINE_CASES) {
      expect([17, 18, 21, 22]).toContain(c.task)
      if (c.route === 'applescript') expect(c.lines, c.id).toBeTypeOf('function')
      if (c.route === 'xml' || c.fallback === 'xml') expect(c.xmlEdit, c.id).toBeTypeOf('function')
      if (c.route === 'recipe') expect(c.recipe, c.id).toBeTruthy()
      expect(c.markers.length + (c.cacheMarkers?.length ?? 0), `${c.id} has no marker`).toBeGreaterThan(0)
    }
  })
  it('covers every group of the brief and every gate that refuses with pivot-engine-unmeasured', () => {
    expect(new Set(ENGINE_CASES.map(c => c.group))).toEqual(new Set(['layout', 'data-quirks', 'values', 'sort-filter-group', 'calculated']))
    const required = [
      // Brief, step 3
      'layout-compact', 'layout-outline', 'layout-tabular', 'three-row-fields', 'collapsed-item', 'subtotal-top', 'subtotal-bottom', 'grand-totals-row-only',
      'grand-totals-column-only', 'grand-totals-off', 'values-on-rows', 'values-on-columns', 'blank-rows', 'repeat-labels', 'named-subtotals',
      'blank-items', 'error-value', 'mixed-type-items', 'hidden-item', 'page-single', 'page-multiple', 'cross-field-visibility', 'show-all',
      'fn-sum', 'fn-count', 'fn-average', 'fn-max', 'fn-min', 'fn-product', 'fn-countNums', 'fn-stdDev', 'fn-stdDevp', 'fn-var', 'fn-varp',
      'base-previous', 'base-next', 'percentOfParentRow', 'rank-ties', 'runTotal', 'index', 'percentDiff', 'addition-order',
      'sort-ascending', 'sort-descending', 'sort-number', 'sort-text', 'sort-blanks', 'months', 'weekdays', 'sort-by-value', 'nested-sort',
      'wildcard', 'between', 'value-filter', 'top-count', 'top-percent', 'top-sum', 'top10-ties', 'date-filter',
      'group-years', 'group-quarters', 'group-months', 'group-out-of-range', 'number-group', 'manual-group', 'group-field-with-base',
      'spaces-in-name', 'division-by-zero', 'calculated-item',
      // task-15-extra-cases.md
      'off-axis-hidden-item', 'merge-item', 'collapsed-above-values', 'time-only', 'dates-before-1900-03-01', 'collation', 'collation-space-pair',
      'fractional-interval', 'nested-filter', 'classic-layout', 'single-data-field', 'no-record-cells',
    ]
    const features = new Set(ENGINE_CASES.flatMap(c => c.features))
    expect(required.filter(f => !features.has(f))).toEqual([])
    // Each unmeasured gate in src/xlsx/pivots/engine (by its message) → the cases that measure it.
    const gates: Record<string, string[]> = {
      'merged item labels (mergeItem)': ['layout-merge-labels'],
      'classic pivot layout (gridDropZones)': ['layout-classic-single-data'],
      'cross-field filter visibility with showAll': ['filter-label-cross-field-showall'],
      'is on no axis but has hidden items': ['quirk-off-axis-hidden-item'],
      'renamed item captions for label filter': ['filter-label-renamed-item'],
      'criteria or item type': ['filter-label-collation', 'filter-label-mixed-items'],
      'nested … filter needs a parent-group oracle': ['filter-value-nested', 'filter-top-nested'],
      'top-10 … thresholds / percent mode': ['filter-top-percent', 'filter-top-sum'],
      'showAll with … filter': ['filter-value-showall'],
      'top-10 threshold ties': ['filter-top-count-ties'],
      'sorted group field': ['sort-grouped-field'],
      'autoSortScope for nested field': ['sort-value-within-groups'],
      'autoSortScope ties': ['sort-value-ties'],
      'autoSortScope placement of no-data items': ['sort-value-showall-no-data'],
      'a collapsed item above the Values field': ['layout-collapsed-item-above-values'],
      'mixed-type item order': ['sort-mixed-blank-ascending', 'sort-mixed-blank-descending'],
      'error-item order': ['sort-error-item'],
      'text collation / space-versus-word collation': ['sort-collation-ascending', 'sort-collation-descending', 'filter-label-collation'],
      'custom-list order': ['sort-custom-list-months', 'sort-custom-list-weekdays', 'sort-custom-list-off'],
      'fractional numeric group start or interval': ['group-number-fractional'],
      'out-of-range number group values': ['group-number-out-of-range'],
      'out-of-range date group labels': ['group-date-out-of-range'],
      'a reference to calculated field': ['calc-references-calculated'],
      'uses no field': ['calc-constant'],
      'calculated data field … with function': ['calc-function-count'],
      'show-as on calculated data field': ['calc-show-as'],
      'a named subtotal of a calculated data field': ['calc-named-subtotals'],
      'text values; their sum in a calculated field': ['calc-text-field'],
      'a calculated cell without records': ['calc-cell-without-records'],
      'calculated data field gives error': ['calc-division-by-zero'],
      'a filter or sort by calculated data field': ['calc-value-sort'],
      'error-valued records': ['quirk-error-value', 'quirk-error-value-count'],
      'shared date values need the workbook date-system context': ['values-fn-max-of-date'],
      'sample variance with fewer than two numbers': ['values-fn-stdev', 'values-fn-var'],
      'filtered-out records require measured visible-total semantics': ['values-average-hidden-item', 'quirk-hidden-row-item'],
      'record-based subtotal/grand-total behavior': ['values-fn-count', 'values-fn-max', 'values-fn-product', 'values-fn-varp'],
      'intersecting row and column named subtotals': ['values-named-subtotals-intersecting'],
      'show-as applied to a named subtotal': ['values-showas-on-named-subtotals'],
      'rank tie handling': ['values-showas-rank-descending-ties'],
      'the index show-as zero denominator': ['values-showas-index-zero-total'],
      'difference-from values for the selected base item': ['values-showas-difference-item'],
      'show-as division by zero': ['values-showas-percent-diff-zero-base', 'values-showas-percent-of-row-zero-total'],
      'parent-total behavior': ['values-showas-percent-of-parent-region', 'values-showas-percent-of-parent-row-two-levels'],
    }
    expect(Object.values(gates).flat().filter(id => !byId.has(id))).toEqual([])
  })
})

// Cases whose AppleScript prep changes the cache shape the edit reads: Excel-saved engine-seed.xlsx keeps an off-axis
// field's text inline (no shared items) until the prep puts the field on an axis, which the stand-in already models.
const PREPPED_SHAPE = new Set(['quirk-off-axis-hidden-item'])

describe('engine case xml edits', async () => {
  const realSeed = await seedParts()
  const standIn = engineSeedStandIn(await parts(join(RECIPES, 'pivot-seed.xlsx')))
  const seedFor = (id: string) => PREPPED_SHAPE.has(id) ? standIn : realSeed
  const seed = realSeed
  const xmlCases = ENGINE_CASES.filter(c => c.xmlEdit)
  it('starts from consistent seed parts', () => expectConsistent(seed))
  it.each(xmlCases.map(c => [c.id, c] as const))('%s: well-formed, consistent and carrying its marker', (_, c) => {
    const seed = seedFor(c.id)
    const next = c.xmlEdit!({ ...seed })
    expect(next.pivot !== seed.pivot || next.cache !== seed.cache).toBe(true)
    expectConsistent(next)
    // At least one of the case's markers is written by the edit itself (the rest may come from the AppleScript prep).
    const fresh = [...c.markers.map(m => [m, seed.pivot, next.pivot] as const), ...(c.cacheMarkers ?? []).map(m => [m, seed.cache, next.cache] as const)]
      .filter(([m, before, after]) => !m.test(before) && m.test(after))
    expect(fresh.length, `${c.id}: no marker added by the edit`).toBeGreaterThan(0)
  })
  it('writes Excel-shaped pieces for the trickier edits', () => {
    const edit = (id: string) => ENGINE_CASES.find(c => c.id === id)!.xmlEdit!({ ...seedFor(id) })
    const named = edit('layout-named-subtotals-compact').pivot
    expect(named).toMatch(fieldAttr(F.Region, 'defaultSubtotal', '0'))
    expect(named).toMatch(/<item t="sum"\/><item t="count"\/><item t="avg"\/><\/items>/)
    const calcItem = edit('calc-item')
    expect(calcItem.cache).toMatch(/<sharedItems [^>]*><s v="Apples"\/>[\s\S]*<s v="Apples and Bananas"\/><\/sharedItems>/)
    expect(calcItem.cache).toMatch(/<\/cacheFields><calculatedItems count="1">/)
    const group = edit('group-manual')
    expect(group.cache).toMatch(/<cacheField name="Product2" numFmtId="0" databaseField="0"><fieldGroup base="1"><discretePr count="4"><x v="0"\/><x v="0"\/>/)
    expect(group.pivot).toMatch(/<rowFields count="3"><field x="0"\/><field x="12"\/><field x="1"\/><\/rowFields>/)
    expect(group.pivot).not.toMatch(/<rowItems/)
    const calc = edit('calc-references-calculated')
    expect(calc.cache).toMatch(/<cacheField name="Price per Unit" numFmtId="0" formula="Sales\/Units" databaseField="0"\/><cacheField name="Double Price" numFmtId="0" formula="&apos;Price per Unit&apos;\*2|formula="'Price per Unit'\*2"/)
    expect(calc.pivot).toMatch(/<dataField name="Sum of Double Price" fld="13"/)
    const filters = edit('filter-label-cross-field-showall').pivot
    expect(filters).toMatch(/showLastColumn="1"\/><filters count="1"><filter fld="1" type="captionBeginsWith" evalOrder="-1" id="\d+" stringValue1="A"><autoFilter ref="A1"><filterColumn colId="0"><customFilters><customFilter val="A\*"\/>/)
    expect(filters).toMatch(fieldAttr(F.Region, 'showAll', '1'))
    const offAxis = edit('quirk-off-axis-hidden-item').pivot
    expect(offAxis).toMatch(/<item h="1" x="\d+"\/>/)
    const x14 = edit('values-showas-rank-descending-ties').pivot
    expect(x14).toMatch(/<dataField name="Sum of Units" fld="4" baseField="0" baseItem="0"><extLst><ext uri="\{E15A36E0-9728-4e99-A89B-3F7291B0FE68\}"[^>]*><x14:dataField pivotShowAs="rankDescending"\/><\/ext><\/extLst><\/dataField>/)
  })
})

describe('engine marker helpers', async () => {
  const seed = await parts(join(RECIPES, 'pivot-seed.xlsx'))
  it('address one pivotField by index on the real seed', () => {
    expect(seed.pivot).toMatch(fieldAttr(0, 'axis', 'axisRow'))
    expect(seed.pivot).toMatch(fieldAttr(2, 'axis', 'axisCol'))
    expect(seed.pivot).toMatch(fieldAttr(3, 'dataField', '1'))
    expect(seed.pivot).not.toMatch(fieldAttr(0, 'axis', 'axisCol'))
    expect(seed.pivot).not.toMatch(fieldAttr(3, 'axis', 'axisRow'))
    expect(seed.pivot).toMatch(rowField(1))
    expect(seed.pivot).not.toMatch(rowField(2))
  })
  it('summarises a diff as added and removed start tags', () => {
    expect(xmlDiffSummary('<a><b x="1"/></a>', '<a><b x="2"/><c/></a>')).toEqual({ added: ['<b x="2"/>', '<c/>'], removed: ['<b x="1"/>'] })
  })
})
