// @vitest-environment node
/**
 * A sparkline workbook written by XlsxWriter 3.3.0 (scripts/sparkline-refs/xlsxwriter-case.py), an independent writer:
 * the resident and paged parsers must read XML this repository did not generate, and agree (spec section 3).
 * XlsxWriter writes groups in reverse insertion order and declares x14 on <ext>, unlike generate.ts.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import type { XlsxSparklineGroup } from '../src/xlsx/sparklines/model'

const bytes = new Uint8Array(readFileSync(resolve(process.cwd(), 'tests/fixtures/xlsx-sparklines/xlsxwriter-written.xlsx')))

const byHost = (groups: XlsxSparklineGroup[]) => new Map(groups.flatMap(g => g.sparklines.map(s => [`${s.host?.row},${s.host?.col}`, { g, s }] as const)))
/** J = column 9 (0-based), row r (1-based) → row r − 1. */
const at = (map: ReturnType<typeof byHost>, row: number) => map.get(`${row - 1},9`)!

describe('XlsxWriter-written sparkline workbook', () => {
  it('parses every group with the options XlsxWriter wrote', async () => {
    const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
    const groups = doc.sheets[0].sparklineGroups ?? []
    expect(groups).toHaveLength(10)
    expect(groups.flatMap(g => g.sparklines)).toHaveLength(11)
    expect((doc.sheets[0].diagnostics ?? []).filter(d => d.feature.startsWith('sparkline-'))).toEqual([])
    const map = byHost(groups)

    expect(at(map, 1).g).toMatchObject({ type: 'line', markers: true })
    expect(at(map, 1).s.data).toEqual({ sheet: 'Sheet2', r0: 0, c0: 0, r1: 0, c1: 7 })
    expect(at(map, 2).g).toMatchObject({ type: 'column', high: true, low: true, first: true, last: true, negative: true })
    expect(at(map, 3).g).toMatchObject({ type: 'stacked', negative: true, displayXAxis: true })
    expect(at(map, 4).g).toMatchObject({ rightToLeft: true, first: true, last: true })
    expect(at(map, 4).g.colors.series).toBe('FFE965E0')
    expect(at(map, 5).g).toMatchObject({ displayEmptyCellsAs: 'span' })
    expect(at(map, 6).g).toMatchObject({ dateAxis: true, lineWeightPt: 2.25, dateRange: { sheet: 'Sheet2', r0: 12, c0: 0, r1: 12, c1: 7 } })
    expect(at(map, 7).g).toMatchObject({ type: 'column', minAxisType: 'custom', maxAxisType: 'custom', manualMin: 0, manualMax: 100 })
    expect(at(map, 8).g).toMatchObject({ displayHidden: true })
    expect(at(map, 9).g).toBe(at(map, 10).g)
    expect(at(map, 9).g).toMatchObject({ minAxisType: 'group', maxAxisType: 'group' })
    // Excel's style presets are theme colours with tints: style 1 is accent1 darkened 50 %. XlsxWriter ships the
    // Office 2007 theme (accent1 4F81BD), so this pins that the workbook's own theme is used, not the Office default.
    expect(at(map, 1).g.colors.series).toBe('FF254061')
    expect(at(map, 11).g).toMatchObject({ high: true, low: true })
  })

  it('reads identically on the paged path', async () => {
    const resident = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId: 'spark-xlsxwriter', profile: 'D' })
    try {
      expect((await session.done).kind).toBe('complete')
      expect(session.workbook.sheets.map(s => s.metadata.sparklineGroups ?? [])).toEqual(resident.sheets.map(s => s.sparklineGroups ?? []))
    } finally {
      await session.release()
    }
  })
})
