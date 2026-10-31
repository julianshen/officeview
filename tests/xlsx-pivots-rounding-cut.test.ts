// @vitest-environment node
/**
 * PR #56 bot round 9 and ruling S2-AJ: value filters, top-10 filters and value sorts decide on item aggregates summed
 * in record order, while the displayed totals use the S2-AB leaf rollup. A decision closer to its cut than the
 * aggregate's rounding bound refuses; exact and well-separated decisions compute. Top Percent cannot overflow.
 */
import { describe, expect, it } from 'vitest'
import { computeAxes } from '../src/xlsx/pivots/engine/axes'
import { VALUE_TAG, type RecordStore } from '../src/xlsx/pivots/engine/records'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { PivotViewState } from '../src/xlsx/pivots/engine/view'

const fieldState = (count: number, extra: Record<string, unknown> = {}) => ({
  hidden: new Array(count).fill(false), collapsed: new Array(count).fill(false), sortType: 'manual', defaultSubtotal: true,
  subtotals: [], subtotalTop: false, compact: false, outline: true, insertBlankRow: false, fillDownLabels: false, showAll: false, ...extra,
})
/** Region (rows) × Channel (columns) records with one numeric field, and optionally a calculated field over it. */
function pivot(records: Array<[region: number, channel: number, value: number]>, formula?: string) {
  const regions = ['A', 'B', 'C'], channels = ['X', 'Y'], count = records.length
  const fields: unknown[] = [
    { name: 'Region', databaseField: true, sharedItems: { inline: false, items: regions.map(value => ({ type: 'string', value })) } },
    { name: 'Channel', databaseField: true, sharedItems: { inline: false, items: channels.map(value => ({ type: 'string', value })) } },
    { name: 'D', databaseField: true, sharedItems: { inline: true, items: [] } },
  ]
  if (formula) fields.push({ name: 'Calc', databaseField: false, formula })
  const cache = { source: { type: 'worksheet' }, olap: false, fields } as unknown as PivotCacheDefinition
  const items = (n: number) => Array.from({ length: n }, (_, x) => ({ type: 'data', x, showDetails: true, hidden: false }))
  const def = { fields: [{ items: items(3) }, { items: items(2) }, {}, {}], rowGrandTotals: true, colGrandTotals: true, customListSort: true } as unknown as PivotTableDefinition
  const store: RecordStore = { count, bytes: 0, columns: [
    { kind: 'shared', index: Uint32Array.from(records, r => r[0]) },
    { kind: 'shared', index: Uint32Array.from(records, r => r[1]) },
    { kind: 'inline', tag: new Uint8Array(count).fill(VALUE_TAG.number), num: Float64Array.from(records, r => r[2]), text: [], textIndex: new Uint32Array(count) },
  ] }
  const view = (data: number, extra: Partial<Record<keyof PivotViewState, unknown>> = {}, region: Record<string, unknown> = {}) => ({
    rows: [0], cols: [1], pages: [], filters: [],
    fields: [fieldState(3, region), fieldState(2), fieldState(0), fieldState(0)],
    data: [{ field: data, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: 'Sum' }], ...extra,
  }) as unknown as PivotViewState
  return { cache, def, store, view }
}
const valueFilter = (type: string, value: string, value2?: string) =>
  ({ fld: 0, type, evalOrder: 0, id: 1, iMeasureFld: 0, stringValue1: value, ...(value2 !== undefined ? { stringValue2: value2 } : {}), autoFilter: { columns: [{ colId: 0 }] } })
const topFilter = (type: 'count' | 'percent' | 'sum', val: number) =>
  ({ fld: 0, type, evalOrder: 0, id: 1, iMeasureFld: 0, autoFilter: { columns: [{ colId: 0, top10: { top: true, percent: type === 'percent', val } }] } })
const rowItems = (result: ReturnType<typeof computeAxes>) => result.ok
  ? result.value.rows.filter(line => line.type === 'data').map(line => line.items[0])
  : `${result.diagnostic.feature}: ${result.diagnostic.message}`

describe('ruling S2-AJ: decisions near their cut refuse', () => {
  it('refuses the bot\'s cancellation case instead of keeping A', () => {
    // A's leaf cells are X = 1e16 + 1 = 1e16 and Y = −1e16, so its displayed subtotal is 0; record order gives 1.
    const p = pivot([[0, 0, 1e16], [0, 1, -1e16], [0, 0, 1], [1, 0, 5]])
    const result = computeAxes(p.store, p.cache, p.def, p.view(2, { filters: [valueFilter('valueGreaterThan', '0')] }))
    expect(rowItems(result)).toMatch(/^pivot-engine-unmeasured: .*too close to the cut for the summation order to be settled/)
  })

  it('computes exact integer sums at the cut, ties keeping item order', () => {
    const p = pivot([[0, 0, 3], [0, 1, 2], [1, 0, 5], [2, 1, 4], [2, 0, 1]])
    // A = 5, B = 5, C = 5: exactly on the cut, exact sums.
    expect(rowItems(computeAxes(p.store, p.cache, p.def, p.view(2, { filters: [valueFilter('valueGreaterThanOrEqual', '5')] })))).toEqual([0, 1, 2])
    expect(rowItems(computeAxes(p.store, p.cache, p.def, p.view(2, { filters: [topFilter('count', 1)] })))).toEqual([0, 1, 2])
    const sorted = p.view(2, {}, { sortType: 'descending', autoSortScope: { references: [{ field: -2, x: [0] }] } })
    expect(rowItems(computeAxes(p.store, p.cache, p.def, sorted))).toEqual([0, 1, 2])
  })

  it('computes decimal sums far from the cut', () => {
    const p = pivot([[0, 0, 1.25], [0, 1, 0.1], [1, 0, 7.75], [1, 1, 0.2], [2, 1, 3.3]])
    expect(rowItems(computeAxes(p.store, p.cache, p.def, p.view(2, { filters: [valueFilter('valueGreaterThan', '3')] })))).toEqual([1, 2])
    expect(rowItems(computeAxes(p.store, p.cache, p.def, p.view(2, { filters: [topFilter('count', 2)] })))).toEqual([1, 2])
    const sorted = p.view(2, {}, { sortType: 'descending', autoSortScope: { references: [{ field: -2, x: [0] }] } })
    expect(rowItems(computeAxes(p.store, p.cache, p.def, sorted))).toEqual([1, 2, 0])
  })

  it('refuses a decimal sum, a near tie in a value sort and a top-N cut within their bounds', () => {
    // 0.1 + 0.2 = 0.30000000000000004: within its bound of 0.3.
    const p = pivot([[0, 0, 0.1], [0, 1, 0.2], [1, 0, 1], [2, 0, 0.3]])
    expect(rowItems(computeAxes(p.store, p.cache, p.def, p.view(2, { filters: [valueFilter('valueGreaterThan', '0.3')] }))))
      .toMatch(/too close to the cut/)
    expect(rowItems(computeAxes(p.store, p.cache, p.def, p.view(2, { filters: [valueFilter('valueBetween', '0', '0.3')] }))))
      .toMatch(/too close to the cut/)
    // A (0.30000000000000004, inexact) and C (0.3, exact single value) are a near tie.
    const sorted = p.view(2, {}, { sortType: 'descending', autoSortScope: { references: [{ field: -2, x: [0] }] } })
    expect(rowItems(computeAxes(p.store, p.cache, p.def, sorted))).toMatch(/too close to the cut/)
    expect(rowItems(computeAxes(p.store, p.cache, p.def, p.view(2, { filters: [topFilter('count', 2)] })))).toMatch(/too close to the cut/)
  })

  it('refuses a calculated field whose value interval straddles the cut', () => {
    // 'D'*2 over 0.1 + 0.2: 0.6000000000000001 against 0.6.
    const p = pivot([[0, 0, 0.1], [0, 1, 0.2], [1, 0, 1]], "'D'*2")
    expect(rowItems(computeAxes(p.store, p.cache, p.def, p.view(3, { filters: [valueFilter('valueGreaterThan', '0.6')] }))))
      .toMatch(/too close to the cut/)
    // Far from the cut it computes.
    expect(rowItems(computeAxes(p.store, p.cache, p.def, p.view(3, { filters: [valueFilter('valueGreaterThan', '1')] })))).toEqual([1])
  })
})

describe('bot round 9: Top Percent does not overflow', () => {
  it('keeps the right items when total × percent would overflow, and refuses an overflowing total', () => {
    // Item sums 8e306 and 2e306 at 30 %: the threshold is 3e306, so only A is kept (total × 30 overflowed to ∞).
    const big = pivot([[0, 0, 8e306], [1, 0, 2e306]])
    expect(rowItems(computeAxes(big.store, big.cache, big.def, big.view(2, { filters: [topFilter('percent', 30)] })))).toEqual([0])
    const overflow = pivot([[0, 0, 1.5e308], [1, 0, 1e308]])
    expect(rowItems(computeAxes(overflow.store, overflow.cache, overflow.def, overflow.view(2, { filters: [topFilter('percent', 30)] }))))
      .toMatch(/^pivot-values-limit: /)
    const normal = pivot([[0, 0, 60], [1, 0, 30], [2, 0, 10]])
    expect(rowItems(computeAxes(normal.store, normal.cache, normal.def, normal.view(2, { filters: [topFilter('percent', 50)] })))).toEqual([0])
  })
})
