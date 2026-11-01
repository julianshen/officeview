// @vitest-environment node
/**
 * Task 36: Value field settings deltas and rules (plan Stage 4, spec §6.7, S4-H, S4-I).
 */
import { describe, expect, it } from 'vitest'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { readFileSync } from 'node:fs'
import { once, loadAll } from './helpers/xlsx-pivots-load'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { openPivotSession } from '../src/xlsx/pivots/session/session'
import { PIVOT_DELTA_KINDS, type PivotValueShowAs } from '../src/xlsx/pivots/session/delta'
import type { PivotDataFunction } from '../src/xlsx/pivots/model'
import { getPivotData } from '../src/xlsx/pivots/session/getpivotdata'
import type { AstNode } from '../src/xlsx/formula/types'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { loadPivotWorkbook } from '../src/xlsx/pivots/load'
import { decodeRecords, RECORD_LIMITS } from '../src/xlsx/pivots/engine/records'
import { sheetCharts } from '../src/components/xlsx/sheetCharts'
import { pivotChartGate } from '../src/xlsx/pivots/session/charts'
import { captionRule } from '../src/xlsx/pivots/session/captions'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import { initialViewState } from '../src/xlsx/pivots/engine/view'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const BASELINE = resolve(FIXTURES, 'engine/baseline.xlsx')
const NO_TOTALS = resolve(FIXTURES, 'interaction/s4-values-no-totals-base.xlsx')
const GPD_SEED = resolve(FIXTURES, 'recipes/getpivotdata-seed.xlsx')
const CHART_SEED = resolve(FIXTURES, 'recipes/pivot-chart-seed.xlsx')

async function openSession(file: string, options?: { verifiedKinds?: ReadonlySet<string>; locale?: 'zh-TW' | 'file' }) {
  const loaded = (await loadAll(file))[0]!
  const saved = sheetCellReader(loaded.sheet)
  const opened = openPivotSession({
    def: loaded.def,
    cache: loaded.cache,
    store: loaded.store,
    date1904: loaded.date1904,
    saved,
    occupied: () => false,
    linkedCharts: [],
    verifiedKinds: options?.verifiedKinds ? new Set(options.verifiedKinds as any) : new Set(PIVOT_DELTA_KINDS),
  })
  if (!opened.ok) throw new Error(`${opened.diagnostic.feature}: ${opened.diagnostic.message}`)
  return { loaded, session: opened.value }
}

describe('Task 36 value field settings deltas', () => {
  it('valueSettings applies all 11 summary functions and all 15 showAs modes', async () => {
    const { session } = await openSession(NO_TOTALS)
    const functions: PivotDataFunction[] = [
      'sum', 'count', 'average', 'max', 'min', 'product', 'countNums',
      'stdDev', 'stdDevp', 'var', 'varp',
    ]

    for (const fn of functions) {
      const result = session.apply({ kind: 'valueSettings', dataIndex: 0, fn, name: `Fn ${fn}` })
      expect(result.ok, `fn ${fn} failed: ${result.ok ? '' : result.diagnostic.message}`).toBe(true)
      if (result.ok) {
        expect(result.value.view.data[0]?.fn).toBe(fn)
      }
    }

    const showAsCases: Array<{ showAs: PivotValueShowAs; baseField?: number; baseItem?: number }> = [
      { showAs: 'normal' },
      { showAs: 'difference', baseField: 0, baseItem: 1048828 },
      { showAs: 'percent', baseField: 0, baseItem: 1048828 },
      { showAs: 'percentDiff', baseField: 0, baseItem: 1048829 },
      { showAs: 'runTotal', baseField: 0 },
      { showAs: 'percentOfRow' },
      { showAs: 'percentOfCol' },
      { showAs: 'percentOfTotal' },
      { showAs: 'index' },
      { showAs: 'percentOfParent', baseField: 0 },
      { showAs: 'percentOfRunningTotal', baseField: 0 },
      { showAs: 'rankAscending', baseField: 0 },
      { showAs: 'rankDescending', baseField: 0 },
    ]

    for (const item of showAsCases) {
      const result = session.apply({ kind: 'valueSettings', dataIndex: 0, ...item })
      expect(result.ok, `showAs ${item.showAs} failed: ${result.ok ? '' : result.diagnostic.message}`).toBe(true)
      if (result.ok) {
        expect(result.value.view.data[0]?.showAs).toBe(item.showAs)
      }
    }

    // percentOfParentRow requires a hierarchy with parent row lines (e.g. baseline has Region & Product on rows)
    const { session: baselineSession } = await openSession(BASELINE)
    const rowRes = baselineSession.apply({ kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfParentRow' })
    expect(rowRes.ok, `showAs percentOfParentRow failed: ${rowRes.ok ? '' : rowRes.diagnostic.message}`).toBe(true)
    if (rowRes.ok) expect(rowRes.value.view.data[0]?.showAs).toBe('percentOfParentRow')

    // percentOfParentCol requires parent column lines (e.g. values-showas-percent-of-parent-column has Year & Region on cols)
    const PARENT_COL = resolve(FIXTURES, 'engine/values-showas-percent-of-parent-column.xlsx')
    const { session: colSession } = await openSession(PARENT_COL)
    const colRes = colSession.apply({ kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfParentCol' })
    expect(colRes.ok, `showAs percentOfParentCol failed: ${colRes.ok ? '' : colRes.diagnostic.message}`).toBe(true)
    if (colRes.ok) expect(colRes.value.view.data[0]?.showAs).toBe('percentOfParentCol')
  })

  it('valueSettings caption rules: auto zh-TW, duplicate suffix, showAs preservation and fail-closed locale', async () => {
    const { session } = await openSession(NO_TOTALS)
    // 1. zh-TW auto-caption when name omitted and fn changed
    const autoResult = session.apply({ kind: 'valueSettings', dataIndex: 0, fn: 'average' })
    expect(autoResult.ok).toBe(true)
    if (autoResult.ok) {
      expect(autoResult.value.view.data[0]?.name).toBe('平均值 - Sales')
    }

    // 2. Duplicate disambiguation suffix
    // We add Sales as a third data field (index 2)
    const withThreeData = session.apply({ kind: 'addDataField', field: 3 })
    expect(withThreeData.ok).toBe(true)
    if (withThreeData.ok) {
      session.commit(withThreeData.value)
      // First field has name '平均值 - Sales'
      const s1 = session.apply({ kind: 'valueSettings', dataIndex: 0, fn: 'average', name: '平均值 - Sales' })
      expect(s1.ok).toBe(true)
      if (s1.ok) session.commit(s1.value)
      // Third field (Sales) changed to average without name -> gets suffix 2
      const s2 = session.apply({ kind: 'valueSettings', dataIndex: 2, fn: 'average' })
      expect(s2.ok).toBe(true)
      if (s2.ok) {
        expect(s2.value.view.data[2]?.name).toBe('平均值 - Sales2')
      }
    }

    // 3. ShowAs preservation: name preserved when fn is unchanged
    const customNamed = session.apply({ kind: 'valueSettings', dataIndex: 0, fn: 'sum', name: 'My Custom Sales' })
    expect(customNamed.ok).toBe(true)
    if (customNamed.ok) session.commit(customNamed.value)
    // Now change only showAs
    const showAsOnly = session.apply({ kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfTotal' })
    expect(showAsOnly.ok).toBe(true)
    if (showAsOnly.ok) {
      expect(showAsOnly.value.view.data[0]?.name).toBe('My Custom Sales')
    }

    // 4. Fail-closed locale: a file without zh-TW samples whose saved def is English (def.name === 'PivotTable1')
    // When fn changes without name, it must reject with pivot-caption-locale
    const engLoaded = (await loadAll(NO_TOTALS))[0]!
    const engDef = { ...engLoaded.def, name: 'PivotTable1' }
    const engSession = openPivotSession({
      def: engDef,
      cache: engLoaded.cache,
      store: engLoaded.store,
      date1904: engLoaded.date1904,
      saved: sheetCellReader(engLoaded.sheet),
      occupied: () => false,
      linkedCharts: [],
      verifiedKinds: new Set(PIVOT_DELTA_KINDS),
    })
    expect(engSession.ok).toBe(true)
    if (engSession.ok) {
      const failClosed = engSession.value.apply({ kind: 'valueSettings', dataIndex: 0, fn: 'average' })
      expect(failClosed.ok).toBe(false)
      if (!failClosed.ok) {
        expect(failClosed.diagnostic.feature).toBe('pivot-caption-locale')
      }
    }

    // 5. User-renamed Chinese pivot with contradictory English data fields (def.name === '銷售分析')
    // When saved names remain 'Sum of Sales' / 'Sum of Units' and no implicit caption samples exist,
    // captionRule must return locale: 'file' and refuse unmeasured default captions.
    const customZhLoaded = (await loadAll(NO_TOTALS))[0]!
    const customZhDef = { ...customZhLoaded.def, name: '銷售分析' }
    const view0 = initialViewState(customZhDef, customZhLoaded.cache)
    const computed = computePivot(customZhLoaded.store, customZhLoaded.cache, customZhDef, view0, { date1904: customZhLoaded.date1904 })
    expect(computed.ok).toBe(true)
    if (computed.ok) {
      const customZhRule = captionRule(customZhDef, sheetCellReader(customZhLoaded.sheet), computed.value)
      expect(customZhRule.locale).toBe('file')
    }

    const customZhSession = openPivotSession({
      def: customZhDef,
      cache: customZhLoaded.cache,
      store: customZhLoaded.store,
      date1904: customZhLoaded.date1904,
      saved: sheetCellReader(customZhLoaded.sheet),
      occupied: () => false,
      linkedCharts: [],
      verifiedKinds: new Set(PIVOT_DELTA_KINDS),
    })
    expect(customZhSession.ok).toBe(true)
    if (customZhSession.ok) {
      const refused = customZhSession.value.apply({ kind: 'valueSettings', dataIndex: 0, fn: 'average' })
      expect(refused.ok).toBe(false)
      if (!refused.ok) {
        expect(refused.diagnostic.feature).toBe('pivot-caption-locale')
      }
      // Supplying an explicit custom name succeeds
      const explicit = customZhSession.value.apply({ kind: 'valueSettings', dataIndex: 0, fn: 'average', name: '平均銷售額' })
      expect(explicit.ok).toBe(true)
    }
  })

  it('valueSettings rejects blank or duplicate custom names and out-of-scope custom numFmtIds with pivot-delta-invalid', async () => {
    const { session } = await openSession(BASELINE)

    // Blank name rejected
    const blank = session.apply({ kind: 'valueSettings', dataIndex: 0, name: '' })
    expect(blank.ok).toBe(false)
    if (!blank.ok) expect(blank.diagnostic.feature).toBe('pivot-delta-invalid')

    const whitespace = session.apply({ kind: 'valueSettings', dataIndex: 0, name: '   ' })
    expect(whitespace.ok).toBe(false)
    if (!whitespace.ok) expect(whitespace.diagnostic.feature).toBe('pivot-delta-invalid')

    // Duplicate name across other data fields rejected
    // BASELINE has 2 data fields: "Sum of Sales" (index 0) and "Sum of Units" (index 1)
    const duplicate = session.apply({ kind: 'valueSettings', dataIndex: 0, name: 'Sum of Units' })
    expect(duplicate.ok).toBe(false)
    if (!duplicate.ok) expect(duplicate.diagnostic.feature).toBe('pivot-delta-invalid')

    // Re-applying own name is permitted
    const sameName = session.apply({ kind: 'valueSettings', dataIndex: 0, name: 'Sum of Sales' })
    expect(sameName.ok).toBe(true)

    // Out-of-scope custom numFmtId rejected
    const unknownId = session.apply({ kind: 'valueSettings', dataIndex: 0, numFmtId: 999 })
    expect(unknownId.ok).toBe(false)
    if (!unknownId.ok) expect(unknownId.diagnostic.feature).toBe('pivot-delta-invalid')

    const nonPresetId = session.apply({ kind: 'valueSettings', dataIndex: 0, numFmtId: 5 })
    expect(nonPresetId.ok).toBe(false)
    if (!nonPresetId.ok) expect(nonPresetId.diagnostic.feature).toBe('pivot-delta-invalid')

    // Valid preset numFmtId accepted
    const validPreset = session.apply({ kind: 'valueSettings', dataIndex: 0, numFmtId: 4 })
    expect(validPreset.ok).toBe(true)

    // numFmtId: null clears format
    const cleared = session.apply({ kind: 'valueSettings', dataIndex: 0, numFmtId: null })
    expect(cleared.ok).toBe(true)
    if (cleared.ok) {
      expect(cleared.value.view.data[0]?.numFmtId).toBeUndefined()
    }
  })

  it('variance functions succeed without totals and fail closed with announcement when totals exist', async () => {
    // 1. Without totals: succeeds
    const { session: noTotalsSession } = await openSession(NO_TOTALS)
    const varianceFns: PivotDataFunction[] = ['stdDev', 'stdDevp', 'var', 'varp']
    for (const fn of varianceFns) {
      const res = noTotalsSession.apply({ kind: 'valueSettings', dataIndex: 0, fn, name: `Stat ${fn}` })
      expect(res.ok, `variance fn ${fn} without totals should succeed`).toBe(true)
    }

    // 2. With totals: fails closed with pivot-engine-unmeasured
    const { session: baselineSession } = await openSession(BASELINE)
    for (const fn of varianceFns) {
      const res = baselineSession.apply({ kind: 'valueSettings', dataIndex: 0, fn, name: `Stat ${fn}` })
      expect(res.ok, `variance fn ${fn} with totals should fail`).toBe(false)
      if (!res.ok) {
        expect(res.diagnostic.feature).toBe('pivot-engine-unmeasured')
      }
    }
  })

  it('chart-linked pivot refuses non-zero numFmtId cleanly', async () => {
    const pkg = await OfficePackage.load(new Uint8Array(readFileSync(CHART_SEED)))
    const doc = await parseXlsx(pkg)
    const regions = new Map(doc.sheets.filter(s => s.pivotRegions?.length).map(s => [s.name, s.pivotRegions!]))
    const wb = await loadPivotWorkbook(pkg, regions)
    const def = wb.tables[0]!, cache = wb.caches.get(def.cacheId)!
    const records = await decodeRecords(once((await pkg.bytes(cache.recordsPartPath!))!), cache, RECORD_LIMITS.D, false)
    if (!records.ok) throw new Error(records.diagnostic.message)
    const sheet = doc.sheets.find(s => s.name === def.sheet)!
    const charts = sheetCharts(sheet.drawings)

    const opened = openPivotSession({
      def,
      cache,
      store: records.store,
      date1904: false,
      saved: sheetCellReader(sheet),
      occupied: () => false,
      linkedCharts: charts,
      chartGate: pivotChartGate,
      verifiedKinds: new Set(PIVOT_DELTA_KINDS),
    })
    expect(opened.ok).toBe(true)
    if (!opened.ok) return
    const chartSession = opened.value

    // General numFmtId: 0 is accepted
    const generalRes = chartSession.apply({ kind: 'valueSettings', dataIndex: 0, numFmtId: 0 })
    expect(generalRes.ok).toBe(true)

    // Non-zero numFmtId is refused cleanly with pivot-chart-mapping
    const formattedRes = chartSession.apply({ kind: 'valueSettings', dataIndex: 0, numFmtId: 4 })
    expect(formattedRes.ok).toBe(false)
    if (!formattedRes.ok) {
      expect(formattedRes.diagnostic.feature).toBe('pivot-chart-mapping')
    }
  })

  it('GETPIVOTDATA resolves renamed data field after valueSettings commit', async () => {
    const loaded = (await loadAll(GPD_SEED))[0]!
    const opened = openPivotSession({
      def: loaded.def,
      cache: loaded.cache,
      store: loaded.store,
      date1904: loaded.date1904,
      saved: sheetCellReader(loaded.sheet),
      occupied: () => false,
      linkedCharts: [],
      verifiedKinds: new Set(PIVOT_DELTA_KINDS),
    })
    expect(opened.ok).toBe(true)
    if (!opened.ok) return
    const session = opened.value

    // Apply moveField and valueSettings to rename dataIndex 1 to "Avg Units" (average)
    const moved = session.apply({ kind: 'moveField', field: 2, axis: 'row' })
    expect(moved.ok).toBe(true)
    if (!moved.ok) return
    session.commit(moved.value)

    const renamed = session.apply({ kind: 'valueSettings', dataIndex: 1, fn: 'average', name: 'Avg Units' })
    expect(renamed.ok).toBe(true)
    if (!renamed.ok) return
    session.commit(renamed.value)

    // Verify GETPIVOTDATA resolves "Avg Units"
    const liveContext = {
      sheet: loaded.sheet.name,
      def: session.def,
      cache: session.cache,
      layout: session.layout,
      date1904: loaded.date1904,
    }

    const anchorArg: AstNode = { type: 'cell', ref: { sheet: loaded.sheet.name, row: 4, col: 0, absCol: false, absRow: false } }
    const nameArg: AstNode = { type: 'string', value: 'Avg Units' }
    const res = getPivotData([liveContext], [nameArg, anchorArg], arg => arg.type === 'string' ? arg.value : null)
    expect(typeof res === 'number' || (typeof res === 'object' && res !== null && !('code' in res))).toBe(true)
    expect(res).not.toEqual({ kind: 'formula-error', code: '#NAME?' })
  })
})
