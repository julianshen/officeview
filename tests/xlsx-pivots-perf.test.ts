// @vitest-environment node
/**
 * Super-linear work in the engine (perf-audit.md, items 1–9 and the show-as footnote): each repro scaled so the
 * pre-fix code takes at least 3× the bound on the development machine, while the fixed code passes it easily, and
 * the result is asserted unchanged where it is not a refusal.
 */
import { describe, expect, it } from 'vitest'
import { computeAxes, type PivotAxes } from '../src/xlsx/pivots/engine/axes'
import { computeValues } from '../src/xlsx/pivots/engine/values'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import { liveDefinition } from '../src/xlsx/pivots/engine/place'
import { checkSavedLayout } from '../src/xlsx/pivots/engine/faithful'
import { compileCalculatedField } from '../src/xlsx/pivots/engine/calculated'
import { VALUE_TAG, type RecordStore } from '../src/xlsx/pivots/engine/records'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { PivotViewState } from '../src/xlsx/pivots/engine/view'

interface Spec { items?: number; labels?: string[]; numeric?: Float64Array; formula?: string; name?: string; hidden?: (k: number) => boolean; showAll?: boolean; defaultSubtotal?: boolean; sortType?: string; index?: (r: number) => number }
const fieldState = (count: number, s: Spec = {}) => ({
  hidden: Array.from({ length: count }, (_, k) => s.hidden?.(k) ?? false), collapsed: new Array(count).fill(false), sortType: s.sortType ?? 'manual',
  defaultSubtotal: s.defaultSubtotal ?? false, subtotals: [], subtotalTop: false, compact: false, outline: true,
  insertBlankRow: false, fillDownLabels: false, showAll: s.showAll ?? false,
})
/** Synthetic pivots, built as in tests/xlsx-pivots-accumulator-budget.test.ts (the audit's harness). */
function build(count: number, specs: Spec[]) {
  const cacheFields: unknown[] = [], defFields: Array<Record<string, unknown>> = [], columns: unknown[] = [], states: Array<ReturnType<typeof fieldState> & Record<string, unknown>> = []
  specs.forEach((s, f) => {
    const name = s.name ?? `F${f}`
    if (s.formula !== undefined) {
      cacheFields.push({ name, databaseField: false, formula: s.formula }); defFields.push({ items: undefined }); columns.push(undefined); states.push(fieldState(0, s)); return
    }
    if (s.numeric) {
      cacheFields.push({ name, databaseField: true, sharedItems: { inline: true, items: [] } }); defFields.push({ items: undefined })
      columns.push({ kind: 'inline', tag: new Uint8Array(count).fill(VALUE_TAG.number), num: s.numeric, text: [], textIndex: new Uint32Array(count) })
      states.push(fieldState(0, s)); return
    }
    const n = s.labels?.length ?? s.items!
    cacheFields.push({ name, databaseField: true, sharedItems: { inline: false, items: Array.from({ length: n }, (_, i) => ({ type: 'string', value: s.labels?.[i] ?? `item${i}` })) } })
    defFields.push({ name, items: Array.from({ length: n }, (_, x) => ({ type: 'data', x, showDetails: true, hidden: s.hidden?.(x) ?? false })) })
    columns.push({ kind: 'shared', index: Uint32Array.from({ length: count }, (_, r) => s.index!(r)) })
    states.push(fieldState(n, s))
  })
  defFields.forEach((d, f) => Object.assign(d, { defaultSubtotal: specs[f]!.defaultSubtotal ?? false, subtotals: [], sortType: 'manual', subtotalTop: false, compact: false, outline: true, insertBlankRow: false, showAll: specs[f]!.showAll ?? false }))
  const cache = { source: { type: 'worksheet' }, olap: false, fields: cacheFields } as unknown as PivotCacheDefinition
  const def = { fields: defFields, rowGrandTotals: true, colGrandTotals: true, unsupported: [], location: { ref: 'A3' }, showHeaders: true, pageWrap: 0, dataCaption: 'Values',
    outlineData: true, compactData: true, customListSort: true, showMissing: true, pageOverThenDown: false, pageFields: [], dataFields: [], rowFields: [], colFields: [], filters: [] } as unknown as PivotTableDefinition
  const store = { count, bytes: 0, columns } as unknown as RecordStore
  return { cache, def, store, states }
}
const view = (states: unknown[], v: Partial<Record<keyof PivotViewState, unknown>>) => ({ rows: [], cols: [], pages: [], filters: [], data: [], fields: states, ...v }) as unknown as PivotViewState
const sumOf = (field: number, extra: Record<string, unknown> = {}) => ({ field, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: `S${field}`, ...extra })
const timed = <T>(fn: () => T): { value: T; ms: number } => { const t0 = performance.now(); const value = fn(); return { value, ms: performance.now() - t0 } }
const labelFilter = (type: string, extra: Record<string, unknown>) => ({ fld: 0, type, evalOrder: 0, id: 1, ...extra, autoFilter: { columns: [{ colId: 0, ...(extra.column as object ?? {}) }] } })
const feature = (result: { ok: boolean; diagnostic?: { feature: string } }) => result.ok ? 'ok' : result.diagnostic!.feature

describe('perf audit: label filters (item 1)', () => {
  it('refuses a wildcard filter past its whole-filter work budget instead of a DP per item', () => {
    // Before: 2.9 s (M=1600 labels of 400 chars, an 8000-character never-matching pattern).
    const M = 1600, labels = Array.from({ length: M }, (_, i) => `L${i}_`.padEnd(400, 'x'))
    const { cache, def, store, states } = build(M, [{ labels, index: r => r }])
    const filter = labelFilter('captionContains', { stringValue1: 'x?'.repeat(3999) + 'q' })
    const { value, ms } = timed(() => computeAxes(store, cache, def, view(states, { rows: [0], filters: [filter] })))
    expect(feature(value)).toBe('pivot-values-limit')
    expect(ms).toBeLessThan(800)
  })

  it('validates an ordered criterion once, with an unchanged result', () => {
    // Before: 5.3 s (a 400,000-character captionGreaterThan criterion over 1000 items).
    const M = 1000, labels = Array.from({ length: M }, (_, i) => `L${i}_xxxxx`)
    const { cache, def, store, states } = build(M, [{ labels, index: r => r }])
    const filter = labelFilter('captionGreaterThan', { stringValue1: 'a'.repeat(400_000) })
    const { value, ms } = timed(() => computeAxes(store, cache, def, view(states, { rows: [0], filters: [filter] })))
    expect(value.ok && value.value.rows.length).toBe(M + 1)   // every L… label sorts after a…a
    expect(ms).toBeLessThan(800)
  })

  it('matches a <filters> value list through one lower-cased set, with an unchanged result', () => {
    // Before: 8.4 s (16,000 values × 16,000 items).
    const M = 16_000, labels = Array.from({ length: M }, (_, i) => `L${i}`)
    const { cache, def, store, states } = build(M, [{ labels, index: r => r }])
    const values = Array.from({ length: M }, (_, i) => i % 2 ? `none${i}` : `l${i}`)   // every even label, case-folded
    const filter = labelFilter('captionEqual', { column: { filters: { blank: false, values } } })
    const { value, ms } = timed(() => computeAxes(store, cache, def, view(states, { rows: [0], filters: [filter] })))
    expect(value.ok && value.value.rows.length).toBe(M / 2 + 1)
    expect(ms).toBeLessThan(1500)
  })

  it('refuses more custom rules than the schema allows instead of comparing each per item', () => {
    // Before: 18.6 s (8000 rules × 8000 items). CT_CustomFilters holds at most two rules.
    const M = 8000, labels = Array.from({ length: M }, (_, i) => `L${i}_xxxxx`)
    const { cache, def, store, states } = build(M, [{ labels, index: r => r }])
    const filters = Array.from({ length: M }, (_, i) => ({ operator: 'greaterThan', val: `a${i}` }))
    const filter = labelFilter('captionEqual', { column: { customFilters: { and: false, filters } } })
    const { value, ms } = timed(() => computeAxes(store, cache, def, view(states, { rows: [0], filters: [filter] })))
    expect(feature(value)).toBe('pivot-engine-unmeasured')
    expect(ms).toBeLessThan(1000)
  })
})

describe('perf audit: faithfulness check walks its rects (item 2)', () => {
  it('checks 8000 lines under 400 page fields without the bounding-box scan, still ok', () => {
    // Before: 7.9 s.
    const L = 8000, n = 400
    const specs: Spec[] = [{ items: L, index: r => r }]
    for (let k = 0; k < n; k++) specs.push({ items: 1, index: () => 0 })
    specs.push({ numeric: Float64Array.from({ length: L }, (_, r) => r) })
    const { cache, def, store, states } = build(L, specs)
    Object.assign(def, { pageWrap: 1 })
    const v = view(states, { rows: [0], data: [sumOf(n + 1)], pages: Array.from({ length: n }, (_, k) => ({ field: k + 1, selected: 'all' })) })
    const layout = computePivot(store, cache, def, v, { date1904: false })
    expect(layout.ok).toBe(true)
    if (!layout.ok) return
    const saved = liveDefinition(def, v, layout.value.axes, layout.value)
    const cells = new Map(layout.value.cells.map(cell => [`${cell.row},${cell.col}`, cell.value]))
    const reader = (row: number, col: number) => {
      const cell = cells.get(`${row},${col}`)
      if (!cell || cell.kind === 'empty') return undefined
      return cell as { kind: 'number'; value: number } | { kind: 'text'; value: string }
    }
    const { value, ms } = timed(() => checkSavedLayout(store, cache, saved, reader, { date1904: false }))
    expect(value).toEqual({ status: 'ok' })
    expect(ms).toBeLessThan(1500)
  })
})

describe('perf audit: Values-axis target lists (item 3)', () => {
  it('reads each slot\'s own subtotal lines, not every data field\'s', () => {
    // Before: 1.0 s (16 subtotalled levels above Values, 150 Count data fields).
    const k = 16, D = 150
    const specs: Spec[] = Array.from({ length: k }, () => ({ items: 1, index: () => 0, defaultSubtotal: true }))
    specs.push({ items: 1, index: () => 0 }, { numeric: Float64Array.of(1) })
    const { cache, def, store, states } = build(1, specs)
    const v = view(states, { rows: [...Array.from({ length: k + 1 }, (_, i) => i), -2], data: Array.from({ length: D }, () => sumOf(k + 1, { fn: 'count' })) })
    const axes = computeAxes(store, cache, def, v)
    expect(axes.ok).toBe(true)
    if (!axes.ok) return
    const { value, ms } = timed(() => computeValues(store, cache, def, v, axes.value))
    expect(value.ok).toBe(true)
    if (value.ok) {
      const grand = axes.value.rows.findIndex(line => line.type === 'grand' && line.data === D - 1)
      expect(value.value.get(grand, 0, D - 1)).toEqual({ kind: 'number', value: 1 })
    }
    expect(ms).toBeLessThan(250)
  })
})

describe('perf audit: calculated fields (items 4 and 7)', () => {
  it('sums each referenced field once, not once per data slot, with unchanged values', () => {
    // Before: 4.6 s (80 calculated fields over 80 fields, Values on columns, 40,000 records).
    const R = 40_000, D = 80
    const specs: Spec[] = [{ items: 10, index: r => r % 10 }]
    for (let i = 0; i < D; i++) specs.push({ numeric: Float64Array.from({ length: R }, (_, r) => r + i), name: `n ${i}` })
    for (let i = 0; i < D; i++) specs.push({ formula: `'n ${i}'*2`, name: `c ${i}` })
    const { cache, def, store, states } = build(R, specs)
    const v = view(states, { rows: [0], cols: [-2], data: Array.from({ length: D }, (_, i) => sumOf(1 + D + i)) })
    const axes = computeAxes(store, cache, def, v)
    expect(axes.ok).toBe(true)
    if (!axes.ok) return
    const { value, ms } = timed(() => computeValues(store, cache, def, v, axes.value))
    expect(value.ok).toBe(true)
    const grand = axes.value.rows.length - 1
    if (value.ok) for (const i of [0, 40, 79]) {
      const col = axes.value.cols.findIndex(line => line.data === i)
      expect(value.value.get(grand, col, i)).toEqual({ kind: 'number', value: 2 * (R * (R - 1) / 2 + R * i) })
    }
    expect(ms).toBeLessThan(1000)
  })

  it('with the Values field on an axis, sums only the slots that name a field, with unchanged values', () => {
    const R = 400, D = 4
    const specs: Spec[] = [{ items: 2, index: r => r % 2 }]
    for (let i = 0; i < D; i++) specs.push({ numeric: Float64Array.from({ length: R }, (_, r) => r * (i + 1)), name: `n ${i}` })
    for (let i = 0; i < D; i++) specs.push({ formula: `'n ${i}'+'n ${(i + 1) % D}'`, name: `c ${i}` })
    const { cache, def, store, states } = build(R, specs)
    const v = view(states, { rows: [0], cols: [-2], data: Array.from({ length: D }, (_, i) => sumOf(1 + D + i)) })
    const axes = computeAxes(store, cache, def, v)
    expect(axes.ok).toBe(true)
    if (!axes.ok) return
    const value = computeValues(store, cache, def, v, axes.value)
    expect(value.ok).toBe(true)
    if (!value.ok) return
    const total = R * (R - 1) / 2, grand = axes.value.rows.length - 1
    for (let i = 0; i < D; i++) {
      const col = axes.value.cols.findIndex(line => line.data === i)
      expect(value.value.get(grand, col, i)).toEqual({ kind: 'number', value: total * (i + 1) + total * ((i + 1) % D + 1) })
    }
  })

  it('mixes ordinary and calculated slots, and formulas over one or two fields, with unchanged values', () => {
    const R = 300
    const specs: Spec[] = [{ items: 3, index: r => r % 3 }, { numeric: Float64Array.from({ length: R }, (_, r) => r), name: 'a' },
      { numeric: Float64Array.from({ length: R }, (_, r) => 2 * r), name: 'b' }, { formula: "'a'*3", name: 'one' }, { formula: "'a'+'b'", name: 'two' }]
    const { cache, def, store, states } = build(R, specs)
    const v = view(states, { rows: [0], cols: [-2], data: [sumOf(2), sumOf(4), sumOf(3)] })
    const axes = computeAxes(store, cache, def, v)
    expect(axes.ok).toBe(true)
    if (!axes.ok) return
    const value = computeValues(store, cache, def, v, axes.value)
    expect(value.ok).toBe(true)
    if (!value.ok) return
    const total = R * (R - 1) / 2, grand = axes.value.rows.length - 1, col = (d: number) => axes.value.cols.findIndex(line => line.data === d)
    expect([0, 1, 2].map(d => value.value.get(grand, col(d), d))).toEqual([
      { kind: 'number', value: 2 * total }, { kind: 'number', value: total + 2 * total }, { kind: 'number', value: 3 * total },
    ])
  })

  it('compiles a chain of 1600 calculated fields in linear time, with unchanged values', () => {
    // Before: compiling the chain 1.1 s, computeValues 1.2 s.
    const C = 1600
    const specs: Spec[] = [{ items: 1, index: () => 0 }, { numeric: Float64Array.of(1), name: 'n' }]
    for (let i = 0; i < C; i++) specs.push({ formula: i === 0 ? "'n'+1" : `'c ${i - 1}'+1`, name: `c ${i}` })
    const { cache, def, store, states } = build(1, specs)
    const v = view(states, { rows: [0], data: Array.from({ length: C }, (_, i) => sumOf(2 + i)) })
    const axes = computeAxes(store, cache, def, v)
    expect(axes.ok).toBe(true)
    if (!axes.ok) return
    const { value, ms } = timed(() => computeValues(store, cache, def, v, axes.value))
    expect(value.ok && value.value.get(0, 0, 0)).toEqual({ kind: 'number', value: 2 })
    // The last link alone: n + C.
    const last = view(states, { rows: [0], data: [sumOf(1 + C)] })
    const lastAxes = computeAxes(store, cache, def, last)
    const lastValue = lastAxes.ok ? computeValues(store, cache, def, last, lastAxes.value) : lastAxes
    expect(lastValue.ok && lastValue.value.get(0, 0, 0)).toEqual({ kind: 'number', value: 1 + C })
    const fresh = build(1, specs)
    const compile = timed(() => { for (let i = 0; i < C; i++) compileCalculatedField(specs[2 + i]!.formula!, fresh.cache, 2 + i) })
    expect(ms).toBeLessThan(300)
    expect(compile.ms).toBeLessThan(300)
  })
})

describe('perf audit: showAll listings (items 5 and 8)', () => {
  it('filters a showAll field\'s order once per axis, not once per parent, with unchanged lines', () => {
    // Before: about 1.7 s (16,000 parents × 16,000 items, a label filter keeping one).
    const P = 16_000
    const { cache, def, store, states } = build(P, [{ items: P, index: r => r }, { items: P, index: () => 0, showAll: true }])
    const filter = { fld: 1, type: 'captionEqual', evalOrder: 0, id: 1, stringValue1: 'item0', autoFilter: { columns: [{ colId: 0 }] } }
    const { value, ms } = timed(() => computeAxes(store, cache, def, view(states, { rows: [0, 1], filters: [filter] })))
    expect(value.ok && value.value.rows.length).toBe(2 * P + 1)
    expect(ms).toBeLessThan(500)
  })

  it('counts a parent\'s no-data items under a showAll value sort without scanning the order', () => {
    // Before: about 0.7 s (16,000 parents, 16,000 hidden items, 2 visible).
    const P = 16_000, H = 16_000, R = 2 * P
    const { cache, def, store, states } = build(R, [{ items: P, index: r => r >> 1 }, { items: H + 2, index: r => r & 1, showAll: true, sortType: 'descending', hidden: k => k >= 2 }, { numeric: Float64Array.from({ length: R }, (_, r) => r) }])
    states[1]!.autoSortScope = { references: [{ field: -2, x: [0] }] }
    const { value, ms } = timed(() => computeAxes(store, cache, def, view(states, { rows: [0, 1], data: [sumOf(2)] })))
    expect(value.ok && value.value.rows.length).toBe(3 * P + 1)
    if (value.ok) expect(value.value.rows.slice(0, 3).map(line => line.items)).toEqual([[0], [0, 1], [0, 0]])   // descending: record 1 > record 0
    expect(ms).toBeLessThan(230)
  })
})

describe('perf audit: filters per field (item 6)', () => {
  it('refuses more than one label or value filter on a field instead of a record pass each', () => {
    // Before: 1.5 s (label) and 1.9 s (value) for 1600 filters over 100,000 records.
    const R = 100_000, M = 1000, K = 1600
    const { cache, def, store, states } = build(R, [{ items: M, index: r => r % M }, { numeric: Float64Array.from({ length: R }, (_, r) => r) }])
    const labels = Array.from({ length: K }, (_, i) => ({ fld: 0, type: 'captionNotEqual', evalOrder: i, id: i, stringValue1: `zz${i}`, autoFilter: { columns: [{ colId: 0 }] } }))
    const values = Array.from({ length: K }, (_, i) => ({ fld: 0, type: 'valueGreaterThan', evalOrder: i, id: i, iMeasureFld: 0, stringValue1: '-1', autoFilter: { columns: [{ colId: 0 }] } }))
    for (const filters of [labels, values]) {
      const { value, ms } = timed(() => computeAxes(store, cache, def, view(states, { rows: [0], data: [sumOf(1)], filters })))
      expect(feature(value)).toBe('pivot-engine-unmeasured')
      expect(ms).toBeLessThan(400)
    }
    // One label and one value filter on the field still compute.
    const both = computeAxes(store, cache, def, view(states, { rows: [0], data: [sumOf(1)], filters: [labels[0]!, { ...values[0]!, evalOrder: 1 }] }))
    expect(both.ok && both.value.rows.length).toBe(M + 1)
  })
})

describe('perf audit: liveDefinition lookups (item 9)', () => {
  it('writes a page field multi-selection through a set, with unchanged items', () => {
    // Before: about 4 s (159,999 of 160,000 items selected).
    const M = 160_000
    const { cache, def, store, states } = build(M, [{ items: 2, index: r => r % 2 }, { items: M, index: r => r }, { numeric: Float64Array.from({ length: M }, (_, r) => r) }])
    const v = view(states, { rows: [0], data: [sumOf(2)], pages: [{ field: 1, selected: Array.from({ length: M - 1 }, (_, k) => k + 1) }] })
    const layout = computePivot(store, cache, def, v, { date1904: false })
    expect(layout.ok).toBe(true)
    if (!layout.ok) return
    const { value, ms } = timed(() => liveDefinition(def, v, layout.value.axes, layout.value))
    const items = value.fields[1]!.items!
    expect([items[0]!.hidden, items[1]!.hidden, items[M - 1]!.hidden]).toEqual([true, false, false])
    expect(ms).toBeLessThan(330)
  })
})

describe('perf audit footnote: the show-as scan cap is one total', () => {
  it('refuses four running-total fields whose scans together pass the cap', () => {
    // Before: 2.9 s (each of the four was under the per-field cap).
    const L = 7000
    const { cache, def, store, states } = build(L, [{ items: L, index: r => r }, { numeric: Float64Array.from({ length: L }, (_, r) => r) }])
    const v = view(states, { rows: [0], data: Array.from({ length: 4 }, () => sumOf(1, { showAs: 'runTotal', baseField: 0 })) })
    const axes = computeAxes(store, cache, def, v)
    expect(axes.ok).toBe(true)
    if (!axes.ok) return
    const { value, ms } = timed(() => computeValues(store, cache, def, v, axes.value as PivotAxes))
    expect(feature(value)).toBe('pivot-values-limit')
    expect(ms).toBeLessThan(900)
  })
})
