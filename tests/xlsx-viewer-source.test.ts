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
import { xlsxBytes, xlsxEntries } from './helpers/xlsx-package'
import { buildZip } from './helpers/zip-builder'

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

// ---------------------------------------------------------------------------
// Small/large adapter parity (spec §4.2): the same file through the object
// adapter (parsed with valuePolicy 'cached') and through the paged importer +
// PagedXlsxSheetAdapter answers every SourceStore question identically.

import { readFileSync, readdirSync, existsSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import type { WorkbookSource } from '../src/worker/viewerSession'
import type { ColumnEntry, AuthoritativeCell } from '../src/xlsx/view/contracts'

const HERE = dirname(fileURLToPath(import.meta.url))
const FULL = { r0: 0, r1: 1_048_575, c0: 0, c1: 16_383 }

async function objectSource(bytes: Uint8Array): Promise<WorkbookSource> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  return createObjectWorkbookSource(doc, 'doc', { valuePolicy: 'cached' })
}
async function pagedSource(bytes: Uint8Array, pageRows = 7) {
  const store = new SourcePageStore({ docId: 'doc', backend: memoryPageBackend() })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'doc', profile: 'D', store, pageRows })
  const outcome = await session.done
  return { session, outcome, source: createPagedWorkbookSource(session) }
}
async function allRows(store: SourceStore) {
  const out: { row: number; cells: readonly AuthoritativeCell[] }[] = []
  let cursor: string | undefined
  for (let guard = 0; guard < 100_000; guard++) {
    const page = await store.readRows(FULL, cursor, opts(store))
    expect(page.error).toBeUndefined()
    out.push(...page.items)
    cursor = page.nextCursor
    if (cursor === undefined) { expect(page.completeness).toBe('complete'); break }
  }
  return out
}
/** Column answer as canonical runs: adjacent blank runs merged (page cuts may split them). */
async function column(store: SourceStore, col: number, encoding?: 'lanes') {
  const out: unknown[] = []
  let cursor: string | undefined
  for (let guard = 0; guard < 100_000; guard++) {
    const page = await store.queryColumn({ col, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 1_048_575 }] }, limit: encoding ? 300_000 : 64, cursor, encoding }, opts(store))
    expect(page.error).toBeUndefined()
    expect(page.completeness).toBe('complete')
    for (const item of page.items as ColumnEntry[]) {
      if (item.kind === 'lanes') {
        for (let i = 0; i < item.rows.length; i++) out.push(['lane', item.rows[i], item.tags[i], item.numbers[i], item.styles[i], String.fromCharCode(...item.text.subarray(item.textStart[i], item.textStart[i + 1]))])
        continue
      }
      const last = out[out.length - 1] as { kind?: string; r1?: number } | undefined
      if (item.kind === 'blank' && last?.kind === 'blank' && last.r1! + 1 === item.r0) { last.r1 = item.r1; continue }
      out.push({ ...item })
    }
    cursor = page.nextCursor
    if (cursor === undefined) break
  }
  return out
}
async function scanAll(store: SourceStore) {
  const out: AuthoritativeCell[] = []
  for await (const page of store.scanPopulated({ range: FULL, limit: 500 }, opts(store))) {
    expect(page.error).toBeUndefined()
    out.push(...page.items)
  }
  return out
}
async function expectParity(name: string, bytes: Uint8Array) {
  const small = await objectSource(bytes)
  const large = await pagedSource(bytes)
  expect(large.outcome.kind, `${name}: paged import outcome`).toBe('complete')
  expect(large.source.sheets).toEqual(small.sheets)
  expect(large.source.declaredFilters?.()).toEqual(small.declaredFilters?.())
  let cells = 0
  for (const { sheetId } of small.sheets) {
    const a = small.store(sheetId), b = large.source.store(sheetId)
    expect(b.extent, `${name}/${sheetId} extent`).toEqual(a.extent)
    const rowsA = await allRows(a), rowsB = await allRows(b)
    expect(rowsB, `${name}/${sheetId} readRows`).toEqual(rowsA)
    cells += rowsA.reduce((n, r) => n + r.cells.length, 0)
    expect(await scanAll(b), `${name}/${sheetId} scanPopulated`).toEqual(await scanAll(a))
    const used = a.extent.used
    const lastCol = used ? Math.min(used.c1, used.c0 + 6) : 0
    for (let col = used?.c0 ?? 0; col <= lastCol; col++) {
      expect(await column(b, col), `${name}/${sheetId} col ${col}`).toEqual(await column(a, col))
      expect(await column(b, col, 'lanes'), `${name}/${sheetId} lanes ${col}`).toEqual(await column(a, col, 'lanes'))
    }
    const structA = await a.queryStructural(FULL, opts(a)), structB = await b.queryStructural(FULL, opts(b))
    expect({ ...structB, stamp: null }, `${name}/${sheetId} structure`).toEqual({ ...structA, stamp: null })
    for (const at of [{ row: 0, col: 0 }, { row: used?.r1 ?? 5, col: used?.c1 ?? 5 }, { row: 1_048_575, col: 16_383 }, ...rowsA.slice(0, 40).flatMap(r => r.cells.map(c => ({ row: c.row, col: c.col })))]) {
      expect(await b.readCell(at, opts(b)), `${name}/${sheetId} readCell ${at.row},${at.col}`).toEqual(await a.readCell(at, opts(a)))
    }
    const groups = [
      { groupId: 'all', rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 1_048_575 }] }, cols: [{ c0: 0, c1: 16_383 }] },
      { groupId: 'head', rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 9 }] }, cols: [{ c0: 0, c1: 3 }] },
      { groupId: 'odd', rows: { kind: 'list' as const, r0: 1, r1: 9, rows: new Uint32Array([1, 3, 5, 7, 9]) }, cols: [{ c0: 0, c1: 1 }, { c0: 3, c1: 3 }] },
    ]
    const ovA = await a.queryOverview({ groups, maxGroups: 16 }, opts(a)), ovB = await b.queryOverview({ groups, maxGroups: 16 }, opts(b))
    expect({ ...ovB, stamp: null }, `${name}/${sheetId} overview`).toEqual({ ...ovA, stamp: null })
    const gA = small.geometry(sheetId), gB = large.source.geometry(sheetId)
    expect(gB.rowOverrides).toEqual(gA.rowOverrides)
    expect(gB.colOverrides).toEqual(gA.colOverrides)
    expect(gB.baseHidden.intervals).toEqual(gA.baseHidden.intervals)
  }
  await large.source.release()
  return cells
}

const EDGE = xlsxBytes({
  sharedStrings: ['héllo', 'x_xD83D__xDE00_y'],
  styles: '<cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="0"/></cellXfs>',
  beforeSheetData: '<cols><col min="2" max="4" width="20" customWidth="1"/></cols>',
  sheetData: [
    COLUMN_A,
    '<row r="10" ht="30" customHeight="1"><c r="A10"><f t="shared" ref="A10:A12" si="0">B10+1</f><v>3</v></c><c r="B10" t="s"><v>1</v></c><c r="C10" t="s"><v>99</v></c></row>',
    '<row r="11" hidden="1"><c r="A11"><f t="shared" si="0"/><v>4</v></c><c r="D11" t="d"><v>2024-01-02</v></c></row>',
    '<row r="12" hidden="1"><c r="A12"><f t="shared" si="0"/></c><c r="B12" t="e"/><c r="C12" t="str"><v></v></c><c r="D12"><v>abc</v></c></row>',
    '<row r="13"><c r="A13"><f t="array" ref="A13:B14">ROW(1:2)</f><v>1</v></c><c r="B13" ca="1"><v>7</v></c></row>',
    '<row r="20"/>',
    '<row r="300000"><c r="XFD300000" t="inlineStr"><is><t>a</t><r><t>b</t><t>ignored</t></r><rPh><t>phon</t></rPh></is></c></row>',
    '<row r="1048576"><c r="A1048576"><v>-0</v></c><c r="XFD1048576"><v>1E308</v></c></row>',
  ].join(''),
  afterSheetData: '<autoFilter ref="A1:D12"><filterColumn colId="1"><filters blank="1"><filter val="x"/></filters></filterColumn><filterColumn colId="0"><customFilters and="1"><customFilter operator="greaterThan" val="1"/><customFilter val="5"/></customFilters></filterColumn></autoFilter><sortState ref="A2:D12"><sortCondition ref="A2:A12" descending="1"/></sortState><mergeCells count="2"><mergeCell ref="B1:C2"/><mergeCell ref="E5:F9"/></mergeCells>',
})

describe('small/large adapter parity (S1.2, same file, cached policy)', () => {
  it('edge package: values, shared/inline strings, errors, unavailable formulas, merges, filters, geometry', async () => {
    expect(await expectParity('edge', EDGE)).toBeGreaterThan(20)
  })

  it('S13.4: oversized values stored as blob records read identically to the object adapter (raw UTF-16, errors, lanes)', async () => {
    const big = (fill: string) => `${fill.repeat(70 * 1024)}_xD800_${fill.repeat(30 * 1024)}`
    const bytes = xlsxBytes({
      sharedStrings: ['short'],
      sheetData: [
        `<row r="1"><c r="A1" t="inlineStr"><is><t>${big('a')}</t></is></c><c r="B1" t="s"><v>0</v></c></row>`,
        `<row r="2"><c r="A2" t="str"><f>REPT("b",3)</f><v>${big('b')}</v></c><c r="B2" t="e"><v>${big('c')}</v></c></row>`,
        `<row r="3"><c r="A3"><v>${big('d')}</v></c><c r="B3"><v>4</v></c></row>`,
      ].join(''),
    })
    const probe = await pagedSource(bytes)
    expect(probe.session.workbook.store.stats.blobs).toBe(4)
    await probe.source.release()
    expect(await expectParity('oversized', bytes)).toBe(6)
  })

  const fixtureDir = join(HERE, 'fixtures', 'xlsx-viewer')
  for (const file of readdirSync(fixtureDir).filter(f => f.endsWith('.xlsx'))) {
    it(`generated fixture ${file}`, async () => {
      expect(await expectParity(file, new Uint8Array(readFileSync(join(fixtureDir, file))))).toBeGreaterThan(0)
    })
  }

  const corpusDir = join(HERE, '..', 'corpus')
  const corpus = existsSync(corpusDir) ? readdirSync(corpusDir).filter(f => f.endsWith('.xlsx') && !f.startsWith('~$')) : []
  /** Files the paged path refuses with a typed reason (they open through the compatibility path). */
  const UNSERVABLE: Record<string, string> = { 'poi-59746_NoRowNums.xlsx': 'rows without row numbers' }
  for (const file of corpus) {
    it.skipIf(!existsSync(join(corpusDir, file)))(`corpus ${file}`, async () => {
      const bytes = new Uint8Array(readFileSync(join(corpusDir, file)))
      const probe = await pagedSource(bytes)
      if (UNSERVABLE[file]) {
        expect(probe.outcome.kind).toBe('incomplete')
        await probe.source.release()
        return
      }
      await probe.source.release()
      await expectParity(file, bytes)
    })
  }
})

describe('review round 1 (opencode) parity findings', () => {
  it('NIT 4: several direct <t> in one <si> concatenate on both adapters', async () => {
    const withSst = xlsxBytesRaw('<si><t>ab</t><t>cd</t></si>', '<row r="1"><c r="A1" t="s"><v>0</v></c></row>')
    const small = (await objectSource(withSst)).store('1'), large = (await pagedSource(withSst)).source.store('1')
    const a = await small.readCell({ row: 0, col: 0 }, opts(small)), b = await large.readCell({ row: 0, col: 0 }, opts(large))
    expect(b).toEqual(a)
    expect(a).toMatchObject({ kind: 'cell', cell: { value: 'abcd' } })
  })

  it('NIT 5: a garbage merge ref is dropped by both adapters with the same typed diagnostic', async () => {
    const bytes = xlsxBytes({ sheetData: rowsOf(3), afterSheetData: '<mergeCells count="2"><mergeCell ref="garbage"/><mergeCell ref="B1:C2"/></mergeCells>' })
    const small = await objectSource(bytes), large = (await pagedSource(bytes)).source
    for (const source of [small, large]) {
      const store = source.store('1')
      const s = await store.queryStructural(FULL, opts(store))
      expect(s.merges.map((m) => m.id)).toEqual(['merge:0:1:1:2'])
      expect(s.completeness).toBe('complete')
      expect(source.coverageDiagnostics?.('1')).toContainEqual({ code: 'invalid-merge-ref', message: 'Ignored 1 merge with an unreadable range' })
    }
  })
})
const rowsOf = (n: number) => Array.from({ length: n }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}"><v>${r}</v></c></row>`).join('')
function xlsxBytesRaw(siXml: string, sheetData: string): Uint8Array {
  const entries = xlsxEntries({ sheetData })
  entries.push({ name: 'xl/sharedStrings.xml', data: `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${siXml}</sst>` })
  return buildZip(entries).bytes
}
