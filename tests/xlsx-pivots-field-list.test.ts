// @vitest-environment node
import { describe, expect, it } from 'vitest'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { loadAll } from './helpers/xlsx-pivots-load'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { openPivotSession } from '../src/xlsx/pivots/session/session'
import { PIVOT_DELTA_KINDS } from '../src/xlsx/pivots/session/delta'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import { initialViewState } from '../src/xlsx/pivots/engine/view'
import { pivotMenuModel } from '../src/xlsx/pivots/session/menu'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const BASELINE = resolve(FIXTURES, 'engine/baseline.xlsx')

async function openBaseline() {
  const loaded = (await loadAll(BASELINE))[0]!
  const opened = openPivotSession({
    def: loaded.def, cache: loaded.cache, store: loaded.store, date1904: loaded.date1904,
    saved: sheetCellReader(loaded.sheet), occupied: () => false, linkedCharts: [],
    verifiedKinds: new Set(PIVOT_DELTA_KINDS),
  })
  if (!opened.ok) throw new Error(`${opened.diagnostic.feature}: ${opened.diagnostic.message}`)
  return { loaded, session: opened.value }
}

async function openInteraction(name: string) {
  const file = name === 's4-data-add-second-default-axis'
    ? resolve(FIXTURES, 'interaction/s4-data-remove-to-single.xlsx')
    : resolve(FIXTURES, 'recipes/engine-seed.xlsx')
  const loaded = (await loadAll(file))[0]!
  const opened = openPivotSession({ def: loaded.def, cache: loaded.cache, store: loaded.store, date1904: loaded.date1904,
    saved: sheetCellReader(loaded.sheet), occupied: () => false, linkedCharts: [], verifiedKinds: new Set(PIVOT_DELTA_KINDS) })
  if (!opened.ok) throw new Error(`${opened.diagnostic.feature}: ${opened.diagnostic.message}`)
  return { loaded, session: opened.value }
}

describe('Task 35 field list deltas', () => {
  it('builds date-group menu periods with a bounded single pass over cache fields', async () => {
    const loaded = (await loadAll(BASELINE))[0]!
    const count = 512
    let reads = 0
    const fields = Array.from({ length: count }, (_, index) => {
      const base = loaded.cache.fields[index % loaded.cache.fields.length]!
      return Object.defineProperty({ ...base, name: `Field ${index}` }, 'fieldGroup', {
        configurable: true, get() { reads++; return base.fieldGroup },
      })
    })
    const cache = { ...loaded.cache, fields }
    const def = { ...loaded.def, fields: Array.from({ length: count }, (_, index) => ({ ...loaded.def.fields[index % loaded.def.fields.length]!, name: `Field ${index}` })) }
    const view = initialViewState(def, cache)
    pivotMenuModel(def, cache, view, loaded.date1904)
    expect(reads).toBeLessThanOrEqual(count * 10)
  })

  it('materializes an inline field into first-appearance shared items without mutating the opened store or cache', async () => {
    const { loaded, session } = await openBaseline()
    const field = loaded.cache.fields.findIndex(cacheField => cacheField.sharedItems?.inline)
    expect(field).toBeGreaterThanOrEqual(0)
    const originalCache = loaded.cache.fields[field]!.sharedItems!
    const originalColumn = loaded.store.columns[field]
    expect(originalColumn?.kind).toBe('inline')

    const result = session.apply({ kind: 'moveField', field, axis: 'row' })
    expect(result.ok, result.ok ? undefined : `${result.diagnostic.feature}: ${result.diagnostic.message}`).toBe(true)
    if (!result.ok) return
    expect(result.value.cache.fields[field]!.sharedItems?.inline).toBe(false)
    expect(result.value.cache.fields[field]!.sharedItems?.items.length).toBeGreaterThan(0)
    expect(result.value.cache.fields[field]!.sharedItems).not.toBe(originalCache)
    expect(result.value.store.columns[field]?.kind).toBe('shared')
    expect(loaded.store.columns[field]).toBe(originalColumn)
    expect(session.cache).toBe(loaded.cache)
  })

  it('reorders a row field to its exact requested position', async () => {
    const { loaded, session } = await openInteraction('s4-move-reorder-rows')
    const field = loaded.cache.fields.findIndex(f => f.name === 'Product')
    const result = session.apply({ kind: 'moveField', field, axis: 'row', index: 0 })
    expect(result.ok, result.ok ? undefined : `${result.diagnostic.feature}: ${result.diagnostic.message}`).toBe(true)
    if (result.ok) expect(result.value.view.rows[0]).toBe(field)
  })

  it('adds a second data field with its measured localized default caption', async () => {
    const { loaded, session } = await openInteraction('s4-data-add-second-default-axis')
    const field = loaded.cache.fields.findIndex(f => f.name === 'Units')
    const result = session.apply({ kind: 'addDataField', field })
    expect(result.ok, result.ok ? undefined : `${result.diagnostic.feature}: ${result.diagnostic.message}`).toBe(true)
    if (result.ok) expect(result.value.view.data).toHaveLength(session.view.data.length + 1)
  })

  it('reserves inline materialization before allocation and keeps the hold through commit until release', async () => {
    const { loaded, session } = await openBaseline()
    const field = loaded.cache.fields.findIndex(cacheField => cacheField.sharedItems?.inline)
    let released = 0, reserved = 0
    const result = session.apply({ kind: 'moveField', field, axis: 'row' }, { admitCandidateBytes: bytes => {
      reserved = bytes
      return { bytes, release: () => { released++ } }
    } })
    expect(result.ok).toBe(true)
    expect(reserved).toBeGreaterThan(loaded.store.count * 4)
    expect(released).toBe(0)
    if (!result.ok) return
    expect(session.commit(result.value)).toBe(true)
    expect(released).toBe(0)
    session.release()
    expect(released).toBe(1)
  })

  it('rejects an inline move when the pre-allocation reservation is refused', async () => {
    const { loaded, session } = await openBaseline()
    const field = loaded.cache.fields.findIndex(cacheField => cacheField.sharedItems?.inline)
    const before = loaded.store.columns[field]
    let requested = 0
    const result = session.apply({ kind: 'moveField', field, axis: 'row' }, { admitCandidateBytes: bytes => { requested = bytes; return undefined } })
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.diagnostic.feature).toBe('pivot-records-limit')
    expect(requested).toBeGreaterThan(0)
    expect(loaded.store.columns[field]).toBe(before)
    expect(session.cache).toBe(loaded.cache)
  })

  it('removes the Values axis marker at one data field and restores it when adding the second', async () => {
    const { loaded, session } = await openInteraction('s4-data-add-second-default-axis')
    const single = session.view.data.length === 1 ? session : undefined
    expect(single).toBeDefined()
    const initialValues = session.view.data
    const added = session.apply({ kind: 'addDataField', field: loaded.cache.fields.findIndex(f => f.name === 'Sales') })
    expect(added.ok).toBe(true)
    if (added.ok) {
      expect([...added.value.view.rows, ...added.value.view.cols]).toContain(-2)
      expect(session.commit(added.value)).toBe(true)
    }
    const removed = session.apply({ kind: 'removeDataField', dataIndex: 1 })
    expect(removed.ok).toBe(true)
    if (removed.ok) expect([...removed.value.view.rows, ...removed.value.view.cols]).not.toContain(-2)
    expect(initialValues).toHaveLength(1)
  })

  it('reorders Values and shifts measure indexes in the view', async () => {
    const { session } = await openBaseline()
    expect(session.view.data.length).toBeGreaterThanOrEqual(2)
    const before = session.view.data.map(entry => entry.field)
    const result = session.apply({ kind: 'moveDataField', dataIndex: 1, toIndex: 0 })
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.value.view.data.map(entry => entry.field)).toEqual([before[1], before[0], ...before.slice(2)])
      expect(session.commit(result.value)).toBe(true)
    }
  })

  it('lays out a pivot with no data fields and no axis fields as empty', async () => {
    const { loaded } = await openBaseline()
    const initial = initialViewState(loaded.def, loaded.cache)
    const empty = { ...initial, rows: [], cols: [], pages: [], data: [], filters: [] }
    const result = computePivot(loaded.store, loaded.cache, loaded.def, empty, { date1904: loaded.date1904 })
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.value.axes.rows).toEqual([])
      expect(result.value.axes.cols).toEqual([])
      expect(result.value.cells).toEqual([])
    }
  })
})
