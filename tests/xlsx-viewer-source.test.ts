/**
 * S1.2 source adapters: typed column lanes and source-space overview on the
 * object adapter (plan S1.2 RED; spec §4.1). The paged adapter's parity with
 * these answers is pinned in xlsx-viewer-source-pages.test.ts.
 */
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import {
  LANE_BLANK, LANE_BOOLEAN, LANE_ERROR, LANE_NUMBER, LANE_STRING, LANE_UNAVAILABLE,
  type ColumnLanes, type OverviewGroup, type SourceStore,
} from '../src/xlsx/view/contracts'
import { SOURCE_BLANK, SOURCE_BOOLEAN, SOURCE_ERROR, SOURCE_NUMBER, SOURCE_STRING, SOURCE_UNAVAILABLE, valueSourceFromStore } from '../src/xlsx/view/valueIndex'
import { xlsxBytes } from './helpers/xlsx-package'

async function storeOf(sheetData: string, sharedStrings?: string[], policy: 'engine' | 'cached' = 'cached'): Promise<SourceStore> {
  const pkg = await OfficePackage.load(xlsxBytes({ sheetData, sharedStrings, styles: '<cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>' }))
  const doc = await parseXlsx(pkg, { valuePolicy: policy })
  return createObjectWorkbookSource(doc, 'doc-1', { valuePolicy: policy }).store('1')
}
const opts = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
const text = (lanes: ColumnLanes, i: number) => String.fromCharCode(...lanes.text.subarray(lanes.textStart[i], lanes.textStart[i + 1]))

// Column A: number, shared string, boolean, cached error, error-looking text,
// style-only record, uncached formula (unavailable under 'cached'), lone surrogate.
const COLUMN_A = [
  '<row r="1"><c r="A1"><v>1.5</v></c></row>',
  '<row r="2"><c r="A2" t="s"><v>0</v></c></row>',
  '<row r="4"><c r="A4" t="b"><v>1</v></c></row>',
  '<row r="5"><c r="A5" t="e"><v>#DIV/0!</v></c></row>',
  '<row r="6"><c r="A6" t="str"><v>#VALUE!</v></c></row>',
  '<row r="7"><c r="A7" s="2"/></row>',
  '<row r="8"><c r="A8"><f>B8*2</f></c></row>',
  '<row r="9"><c r="A9" t="inlineStr"><is><t>x_xD800_y</t></is></c></row>',
].join('')

describe('typed column lanes (S1.2 follow-up, mandatory before S7/S8)', () => {
  it('answers encoding=lanes with one typed lane run covering the rows, raw UTF-16 kept', async () => {
    const store = await storeOf(COLUMN_A, ['héllo'])
    const page = await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 11 }] }, limit: 100, encoding: 'lanes' }, opts(store))
    expect(page.error).toBeUndefined()
    expect(page.completeness).toBe('complete')
    expect(page.items).toHaveLength(1)
    const lanes = page.items[0] as ColumnLanes
    expect(lanes.kind).toBe('lanes')
    expect([lanes.r0, lanes.r1]).toEqual([0, 11])
    expect([...lanes.rows]).toEqual([0, 1, 3, 4, 5, 6, 7, 8])
    expect([...lanes.tags]).toEqual([LANE_NUMBER, LANE_STRING, LANE_BOOLEAN, LANE_ERROR, LANE_STRING, LANE_BLANK, LANE_UNAVAILABLE, LANE_STRING])
    expect(lanes.numbers[0]).toBe(1.5)
    expect(lanes.numbers[2]).toBe(1)
    expect(text(lanes, 1)).toBe('héllo')
    expect(text(lanes, 3)).toBe('#DIV/0!')
    expect(text(lanes, 4)).toBe('#VALUE!')
    expect(lanes.styles[5]).toBe(2)
    const lone = text(lanes, 7)
    expect(lone.length).toBe(3)
    expect(lone.charCodeAt(1)).toBe(0xd800)
  })

  it('bounds one lanes page by limit rows and continues with a cursor', async () => {
    const store = await storeOf(COLUMN_A, ['héllo'])
    const first = await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 11 }] }, limit: 5, encoding: 'lanes' }, opts(store))
    const a = first.items[0] as ColumnLanes
    expect([a.kind, a.r0, a.r1]).toEqual(['lanes', 0, 4])
    expect(first.nextCursor).toBeDefined()
    const second = await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 11 }] }, limit: 5, encoding: 'lanes', cursor: first.nextCursor }, opts(store))
    const b = second.items[0] as ColumnLanes
    expect([b.r0, b.r1]).toEqual([5, 9])
  })

  it('default encoding never yields lanes (existing consumers are unchanged)', async () => {
    const store = await storeOf(COLUMN_A, ['héllo'])
    const page = await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 11 }] }, limit: 100 }, opts(store))
    expect(page.items.some((item) => item.kind === 'lanes')).toBe(false)
  })

  it('the value index reads the same page from lanes as from cell entries', async () => {
    const store = await storeOf(COLUMN_A, ['héllo'])
    const read = await valueSourceFromStore(store, 0, { r0: 0, r1: 11 }).readPage(0, new AbortController().signal)
    expect(read.kind).toBe('values')
    if (read.kind !== 'values') return
    expect([...read.tags]).toEqual([SOURCE_NUMBER, SOURCE_STRING, SOURCE_BLANK, SOURCE_BOOLEAN, SOURCE_ERROR, SOURCE_STRING, SOURCE_BLANK, SOURCE_UNAVAILABLE, SOURCE_STRING, SOURCE_BLANK, SOURCE_BLANK, SOURCE_BLANK])
    expect(read.texts[1]).toBe('héllo')
    expect(read.texts[8]?.charCodeAt(1)).toBe(0xd800)
    expect(read.styles?.[6]).toBe(2)
  })

  it('lane tags share the value index numbering', () => {
    expect([LANE_NUMBER, LANE_STRING, LANE_BOOLEAN, LANE_ERROR, LANE_BLANK, LANE_UNAVAILABLE]).toEqual([SOURCE_NUMBER, SOURCE_STRING, SOURCE_BOOLEAN, SOURCE_ERROR, SOURCE_BLANK, SOURCE_UNAVAILABLE])
  })
})

describe('object adapter queryOverview (source-space aggregates, §4.1)', () => {
  const GRID = [
    '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row>',
    '<row r="2"><c r="A2" s="1"><v>3</v></c></row>',
    '<row r="3"><c r="A3" s="1"><v>4</v></c><c r="B3" s="1"><v>5</v></c></row>',
    '<row r="5"><c r="C5" s="2"/></row>',
  ].join('')
  const byId = (items: readonly OverviewGroup[]) => new Map(items.map((g) => [g.groupId, g]))

  it('noncontiguous source rows {3,1} exclude row 2 and blank spans are analytic', async () => {
    const store = await storeOf(GRID)
    const page = await store.queryOverview({
      maxGroups: 16,
      groups: [
        { groupId: 'rows-1-3', rows: { kind: 'list', r0: 0, r1: 2, rows: new Uint32Array([0, 2]) }, cols: [{ c0: 0, c1: 1 }] },
        { groupId: 'bbox', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 2 }] }, cols: [{ c0: 0, c1: 1 }] },
        { groupId: 'far', rows: { kind: 'intervals', intervals: [{ r0: 1000, r1: 1_048_575 }] }, cols: [{ c0: 0, c1: 16383 }] },
        { groupId: 'style-only', rows: { kind: 'intervals', intervals: [{ r0: 4, r1: 4 }] }, cols: [{ c0: 2, c1: 2 }] },
      ],
    }, opts(store))
    expect(page.error).toBeUndefined()
    expect(page.completeness).toBe('complete')
    const g = byId(page.items)
    expect(g.get('rows-1-3')).toEqual({ groupId: 'rows-1-3', populatedCount: 4, countKind: 'exact', style: { kind: 'uniform', styleIndex: 1 }, content: 'populated' })
    expect(g.get('bbox')).toMatchObject({ populatedCount: 5, content: 'mixed', style: { kind: 'mixed' } })
    expect(g.get('far')).toEqual({ groupId: 'far', populatedCount: 0, countKind: 'exact', style: { kind: 'uniform', styleIndex: 0 }, content: 'blank' })
    expect(g.get('style-only')).toEqual({ groupId: 'style-only', populatedCount: 0, countKind: 'exact', style: { kind: 'uniform', styleIndex: 2 }, content: 'blank' })
  })

  it('bounds output by maxGroups with a cursor, and rejects stale stamps and cancellation', async () => {
    const store = await storeOf(GRID)
    const groups = Array.from({ length: 5 }, (_, i) => ({ groupId: `g${i}`, rows: { kind: 'intervals' as const, intervals: [{ r0: i, r1: i }] }, cols: [{ c0: 0, c1: 0 }] }))
    const first = await store.queryOverview({ maxGroups: 2, groups }, opts(store))
    expect(first.items.map((g) => g.groupId)).toEqual(['g0', 'g1'])
    expect(first.nextCursor).toBeDefined()
    const rest = await store.queryOverview({ maxGroups: 2, groups, cursor: first.nextCursor }, opts(store))
    expect(rest.items.map((g) => g.groupId)).toEqual(['g2', 'g3'])
    const stale = await store.queryOverview({ maxGroups: 2, groups }, { signal: new AbortController().signal, stamp: { ...store.stamp, valueGen: 99 } })
    expect(stale.error?.code).toBe('stale-source')
    const controller = new AbortController()
    controller.abort()
    const cancelled = await store.queryOverview({ maxGroups: 2, groups }, { signal: controller.signal, stamp: { ...store.stamp } })
    expect(cancelled.error?.code).toBe('cancelled')
    const tooMany = await store.queryOverview({ maxGroups: 4096, groups }, opts(store))
    expect(tooMany.error?.code).toBe('resource-limit')
  })
})
