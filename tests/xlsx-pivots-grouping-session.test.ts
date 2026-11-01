// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { loadAll } from './helpers/xlsx-pivots-load'
import { groupColumn } from '../src/xlsx/pivots/engine/groups'
import { PIVOT_DELTA_KINDS, type PivotDelta } from '../src/xlsx/pivots/session/delta'
import { openPivotSession } from '../src/xlsx/pivots/session/session'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'

const ROOT = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const interaction = (name: string) => resolve(ROOT, `interaction/${name}.xlsx`)
const pivot = (loaded: Awaited<ReturnType<typeof loadAll>>[number], verified = new Set(PIVOT_DELTA_KINDS)) => {
  const saved = sheetCellReader(loaded.sheet)
  const opened = openPivotSession({ def: loaded.def, cache: loaded.cache, store: loaded.store, date1904: loaded.date1904,
    saved, occupied: () => false, linkedCharts: [], verifiedKinds: verified })
  if (!opened.ok) throw new Error(`${opened.diagnostic.feature}: ${opened.diagnostic.message}`)
  return opened.value
}

describe('Task 37 grouping', () => {
  it('groups days, N-day intervals, hours, minutes and seconds including serial 60', async () => {
    for (const [file, groupBy] of [
      ['s4-group-date-days-single', 'days'], ['s4-group-date-days-7', 'days'],
      ['s4-group-date-time-hms', 'hours'], ['s4-group-date-time-hms', 'minutes'], ['s4-group-date-time-hms', 'seconds'],
    ] as const) {
      const loaded = (await loadAll(interaction(file)))[0]!
      const field = loaded.cache.fields.findIndex(item => item.fieldGroup?.rangePr?.groupBy === groupBy)
      expect(field, `${file} should contain the saved grouped cache field`).toBeGreaterThanOrEqual(0)
      const grouped = groupColumn(loaded.store, loaded.cache, field)
      expect(grouped.ok, grouped.ok ? undefined : `${grouped.diagnostic.feature}: ${grouped.diagnostic.message}`).toBe(true)
    }
  })

  it('maps a missing date record into the measured lower bucket for a day group', async () => {
    const loaded = (await loadAll(interaction('s4-group-date-days-single')))[0]!
    const field = loaded.cache.fields.findIndex(item => item.fieldGroup?.rangePr?.groupBy === 'days')
    const base = loaded.cache.fields[field]!.fieldGroup!.base!
    const original = loaded.cache.fields[base]!.sharedItems!
    const missingIndex = original.items.length
    const cache = { ...loaded.cache, fields: loaded.cache.fields.map((item, index) => index === base
      ? { ...item, sharedItems: { ...original, containsBlank: true, items: [...original.items, { type: 'missing' as const }] } } : item) }
    const column = loaded.store.columns[base]!
    if (column.kind !== 'shared') throw new Error('Expected a shared date column')
    const indices = new Uint32Array(column.index)
    indices[0] = missingIndex
    const store = { ...loaded.store, columns: loaded.store.columns.map((item, index) => index === base ? { ...column, index: indices } : item) }
    const grouped = groupColumn(store, cache, field)
    expect(grouped.ok, grouped.ok ? undefined : `${grouped.diagnostic.feature}: ${grouped.diagnostic.message}`).toBe(true)
    if (!grouped.ok) return
    const lower = loaded.cache.fields[field]!.fieldGroup!.groupItems!.findIndex(item => item.type === 'string' && item.value.startsWith('<'))
    expect(lower).toBeGreaterThanOrEqual(0)
    expect(grouped.value[0]).toBe(lower)
  })

  it('groups number ranges with blank, negative, below-start and above-end source records', async () => {
    const loaded = (await loadAll(interaction('s4-group-number-sales-blanks-manual')))[0]!
    const field = loaded.cache.fields.findIndex(item => item.fieldGroup?.rangePr?.groupBy === 'range')
    expect(field).toBeGreaterThanOrEqual(0)
    const grouped = groupColumn(loaded.store, loaded.cache, field)
    expect(grouped.ok, grouped.ok ? undefined : `${grouped.diagnostic.feature}: ${grouped.diagnostic.message}`).toBe(true)
  })

  it('applies a multi-level date group as a pure candidate, then commits its cache and axis hierarchy', async () => {
    const seed = (await loadAll(interaction('interaction-filter-date-base')))[0]!
    const session = pivot(seed)
    const oldCache = session.cache, oldRows = session.view.rows
    const delta: PivotDelta = { kind: 'groupDate', field: 5, periods: ['months', 'quarters', 'years'] }
    const candidate = session.apply(delta)
    expect(candidate.ok, candidate.ok ? undefined : `${candidate.diagnostic.feature}: ${candidate.diagnostic.message}`).toBe(true)
    if (!candidate.ok) return
    expect(session.cache).toBe(oldCache)
    expect(session.view.rows).toBe(oldRows)
    expect(candidate.value.view.rows).toEqual(expect.arrayContaining([expect.any(Number)]))
    expect(candidate.value.cache.fields.length).toBeGreaterThan(oldCache.fields.length)
    expect(session.commit(candidate.value)).toBe(true)
    expect(session.cache).toBe(candidate.value.cache)
    expect(session.view.rows.slice(-3)).toEqual([14, 13, 12])
    const regroup = session.apply({ kind: 'groupDate', field: 14, periods: ['months', 'years'] })
    expect(regroup.ok, regroup.ok ? undefined : `${regroup.diagnostic.feature}: ${regroup.diagnostic.message}`).toBe(true)
    if (!regroup.ok) return
    expect(session.commit(regroup.value)).toBe(true)
    expect(session.view.rows.slice(-2)).toEqual([13, 12])
    const ungrouped = session.apply({ kind: 'ungroup', field: 5 })
    expect(ungrouped.ok, ungrouped.ok ? undefined : `${ungrouped.diagnostic.feature}: ${ungrouped.diagnostic.message}`).toBe(true)
    if (!ungrouped.ok) return
    expect(session.commit(ungrouped.value)).toBe(true)
    expect(session.view.rows).toContain(5)
    expect(session.cache.fields).toHaveLength(oldCache.fields.length)
  })

  it('groups and ungroups a number field without mutating the opened cache', async () => {
    const loaded = (await loadAll(interaction('s4-move-sales-only-rows')))[0]!
    const session = pivot(loaded)
    const oldCache = session.cache
    const grouped = session.apply({ kind: 'groupNumber', field: 3, start: 0, end: 5000, by: 1000 })
    expect(grouped.ok, grouped.ok ? undefined : `${grouped.diagnostic.feature}: ${grouped.diagnostic.message}`).toBe(true)
    if (!grouped.ok) return
    expect(grouped.value.cache.fields[3]?.fieldGroup?.groupItems).toHaveLength(7)
    expect(session.cache).toBe(oldCache)
    expect(session.commit(grouped.value)).toBe(true)
    const ungrouped = session.apply({ kind: 'ungroup', field: 3 })
    expect(ungrouped.ok, ungrouped.ok ? undefined : `${ungrouped.diagnostic.feature}: ${ungrouped.diagnostic.message}`).toBe(true)
    if (!ungrouped.ok) return
    expect(ungrouped.value.cache.fields[3]?.fieldGroup).toBeUndefined()
    expect(session.commit(ungrouped.value)).toBe(true)
  })
})
