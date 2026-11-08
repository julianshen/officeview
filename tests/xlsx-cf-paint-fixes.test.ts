// @vitest-environment node
import { expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { getPaintables } from '../src/render/paint'
import { prepareSheetAssets, preparePrintUnits, mainChartParent } from '../src/components/xlsx/sheetAssets'
import type { XlsxSheet } from '../src/xlsx/types'
function sheet(): XlsxSheet & { conditionalFormatting: { diagnostics: []; rules: import('../src/xlsx/cf/model').CfRule[] } } {
  return { name: 'CF', sheetId: '1', merges: [], mergeRanges: [], cols: [], rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 10, styleIndex: 0 }] }], conditionalFormatting: { diagnostics: [], rules: [{ id: 0, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }], writes: new Set(['fill']), dxf: { 'fill.bg': 'FF00FF00' }, definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 5 }], dialect: 'base', formulas: [] } }] } }
}
it('resolves symbolic CF colors against the workbook palette', async () => {
  const s = sheet(); s.conditionalFormatting!.rules[0] = { ...s.conditionalFormatting!.rules[0], dxf: { 'fill.bg': 'theme:4:0' } }
  s.tableStyleSource = { colors: { theme: { accent1: '#ff0000' } } } as unknown as XlsxSheet['tableStyleSource']
  const batch = await prepareCfPaint(s)
  try { expect(batch?.layer.styleAt(0, 0, undefined)?.fillColor?.toLowerCase()).toMatch(/ff0000$/) } finally { batch?.dispose() }
})
it('prepares more than 5000 cells in bounded batches', async () => {
  const s = sheet(); s.conditionalFormatting!.rules[0] = { ...s.conditionalFormatting!.rules[0], ranges: [{ firstRow: 0, lastRow: 500, firstCol: 0, lastCol: 9 }] }
  const batch = await prepareCfPaint(s)
  try { expect(batch?.layer.styleAt(0, 0, undefined)?.fillColor).toBe('FF00FF00') } finally { batch?.dispose() }
})
it('does not count overlapping applies-to cells twice', async () => {
  const s = sheet(), r = s.conditionalFormatting!.rules[0]
  s.conditionalFormatting!.rules = [{ ...r, ranges: [{ firstRow: 0, lastRow: 299, firstCol: 0, lastCol: 9 }] }, { ...r, id: 1, priority: 2, ranges: [{ firstRow: 1, lastRow: 299, firstCol: 0, lastCol: 9 }] }]
  const batch = await prepareCfPaint(s)
  try { expect(batch?.layer.styleAt(0, 0, undefined)?.fillColor).toBe('FF00FF00') } finally { batch?.dispose() }
})
it('accepts 257 single-cell rules within sheet limits', async () => {
  const s = sheet(), r = s.conditionalFormatting!.rules[0]
  s.conditionalFormatting!.rules = Array.from({ length: 257 }, (_, row) => ({ ...r, id: row, priority: row + 1, ranges: [{ firstRow: row, lastRow: row, firstCol: 0, lastCol: 0 }] }))
  const batch = await prepareCfPaint(s)
  try { expect(batch?.layer.styleAt(0, 0, undefined)?.fillColor).toBe('FF00FF00') } finally { batch?.dispose() }
})
it('whole columns cover blanks inside merge-expanded static bounds', async () => {
  const s = sheet(), r = s.conditionalFormatting!.rules[0]
  s.mergeRanges = [{ minRow: 0, maxRow: 4, minCol: 0, maxCol: 0 }]
  s.conditionalFormatting!.rules[0] = { ...r, ranges: [{ firstRow: 0, lastRow: 1048575, firstCol: 1, lastCol: 1 }], definition: { kind: 'predicate', type: 'containsBlanks', dialect: 'base', formulas: [] } }
  const b = await prepareCfPaint(s)
  try { expect(b?.layer.styleAt(4, 1, undefined)?.fillColor).toBe('FF00FF00') } finally { b?.dispose() }
})
it('open-ended column tails never expand the static sheet', async () => {
  const s = sheet(); s.rows.push({ index: 2, cells: [{ row: 2, col: 0, ref: 'A3', value: 20, styleIndex: 0 }] })
  s.conditionalFormatting!.rules[0] = { ...s.conditionalFormatting!.rules[0], ranges: [{ firstRow: 1, lastRow: 1048575, firstCol: 0, lastCol: 0 }] }
  expect(computeMetrics(s).rowHeightsPx).toHaveLength(3)
  const p = await getPaintables({ sheets: [s] })
  try { expect(p[0].refusal).toBeUndefined() } finally { p.dispose() }
})
it('a finite rule which adds no extent preserves legacy tall sheet rendering', async () => {
  const s = sheet(); s.rows.push({ index: 999, cells: [{ row: 999, col: 0, ref: 'A1000', value: 20, styleIndex: 0 }] })
  expect(computeMetrics(s).exportLimits).toBeUndefined()
  expect(() => renderSheet(s, createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D)).not.toThrow()
})
it('uncached shared formula followers never match a blank rule', async () => {
  const s = sheet(); s.rows[0].cells[0] = { row: 0, col: 0, ref: 'A1', value: null, styleIndex: 0, sharedFormula: { si: 0 }, hasCachedValue: false }
  s.conditionalFormatting!.rules[0] = { ...s.conditionalFormatting!.rules[0], definition: { kind: 'predicate', type: 'containsBlanks', dialect: 'base', formulas: [] } }
  const b = await prepareCfPaint(s)
  try { expect(b?.layer.styleAt(0, 0, undefined)?.fillColor).toBeUndefined() } finally { b?.dispose() }
})
it('initial refusal emits a bounded deduplicated diagnostic', async () => {
  const s = sheet(), memory = createCfMemory({ limitBytes: 0 })
  await prepareCfPaint(s, { memory }); await prepareCfPaint(s, { memory })
  expect(s.diagnostics?.filter(d => d.feature === 'cf-paint-unavailable')).toHaveLength(1)
  s.diagnostics = Array.from({ length: 64 }, () => ({ kind: 'unsupported', feature: 'existing', message: 'existing' }))
  await prepareCfPaint(s, { memory }); expect(s.diagnostics).toHaveLength(64)
})
it('asset CF admission uses the supplied main ledger parent, including print', async () => {
  const s = sheet(); let bytes = 100
  const parent = mainChartParent({ get bytes() { return bytes }, set(_pool, _owner, n) { bytes = 100 + n } }, 100)
  const a = await prepareSheetAssets({ sheets: [s] }, { chartParent: parent })
  try { expect(a.assets.conditionalFormatting).toBeUndefined(); expect(bytes).toBe(100) } finally { a.dispose() }
  const p = await preparePrintUnits({ sheets: [s] }, { chartParent: parent }); p.dispose(); expect(bytes).toBe(100)
})
it('visible blank rows beyond stored bounds use an admitted viewport request', async () => {
  const { buildSheetGeometry } = await import('../src/xlsx/view/viewport')
  const s = sheet(), r = s.conditionalFormatting.rules[0]
  s.conditionalFormatting.rules[0] = { ...r, ranges: [{ firstRow: 0, lastRow: 1048575, firstCol: 1, lastCol: 1 }], definition: { kind: 'predicate', type: 'containsBlanks', dialect: 'base', formulas: [] } }
  const g = buildSheetGeometry(s); if (g.kind !== 'ready') throw Error('geometry')
  const a = await prepareSheetAssets({ sheets: [s] }, { cfViewport: { metrics: g.metrics, boxes: [{ x: 0, y: 0, width: 300, height: 120 }] } })
  try { expect(a.assets.conditionalFormatting?.layer.styleAt(5, 1, undefined)?.fillColor).toBe('FF00FF00') } finally { a.dispose() }
})
it('paged import exposes the same explicit Z5 used extent as the resident source', async () => {
  const { xlsxBytes } = await import('./helpers/xlsx-package')
  const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
  const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
  const { parseXlsx } = await import('../src/xlsx/parse')
  const { OfficePackage } = await import('../src/core/zip')
  const { createObjectWorkbookSource } = await import('../src/xlsx/view/source')
  const bytes = xlsxBytes({ sheetData: '<row r="1"><c r="A1"><v>10</v></c><c r="H1"><v>20</v></c></row>', styles: '<cellXfs count="1"><xf numFmtId="0"/></cellXfs><dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs>', afterSheetData: '<conditionalFormatting sqref="Z5"><cfRule type="containsBlanks" priority="1" dxfId="0"/></conditionalFormatting>' })
  const resident = createObjectWorkbookSource(await parseXlsx(await OfficePackage.load(bytes)), 'cf-extent')
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'cf-extent', profile: 'D' })
  const source = createPagedWorkbookSource(session)
  try {
    expect((await session.done).kind).toBe('complete')
    expect(source.store('1')!.extent.used).toEqual(resident.store('1')!.extent.used)
    expect(source.store('1')!.extent.used).toMatchObject({ r1: 4, c1: 25 })
  } finally { await source.release(); await resident.release() }
})
it('static sparse paint constructs bounded surfaces while reaching the last blank', async () => {
  const { syntheticCellWork } = await import('../src/xlsx/render')
  const s = sheet(), r = s.conditionalFormatting.rules[0]
  s.conditionalFormatting.rules[0] = { ...r, ranges: [{ firstRow: 0, lastRow: 500, firstCol: 0, lastCol: 9 }], definition: { kind: 'predicate', type: 'containsBlanks', dialect: 'base', formulas: [] } }
  const p = await getPaintables({ sheets: [s] })
  try {
    const ctx = createCanvas(p[0].spec.widthPx, p[0].spec.heightPx).getContext('2d')
    syntheticCellWork.peakLive = 0
    p[0].paint(ctx as unknown as CanvasRenderingContext2D)
    expect([...ctx.getImageData(9 * 64 + 5, 500 * 20 + 5, 1, 1).data]).toEqual([0, 255, 0, 255])
    // Actual concurrent blank records, independent of background-fill/tile count.
    expect(syntheticCellWork.peakLive).toBeGreaterThan(0)
    expect(syntheticCellWork.peakLive).toBeLessThanOrEqual(256)
    expect(syntheticCellWork.live).toBe(0)
  } finally { p.dispose() }
})
it('an out-of-order empty row cannot hide a later populated value', async () => {
  const s = sheet(); s.rows.push({ index: 100, cells: [] }, { index: 2, cells: [{ row: 2, col: 0, ref: 'A3', value: 10, styleIndex: 0 }] })
  s.conditionalFormatting.rules[0] = { ...s.conditionalFormatting.rules[0], ranges: [{ firstRow: 2, lastRow: 2, firstCol: 0, lastCol: 0 }] }
  const b = await prepareCfPaint(s)
  try { expect(b?.layer.styleAt(2, 0, undefined)?.fillColor).toBe('FF00FF00') } finally { b?.dispose() }
})
it('print and overlapping assets share one account and refund all parent charges', async () => {
  const s = sheet(); let held = 0, peak = 0
  const parent = { reserve(n: number) { held += n; peak = Math.max(peak, held); return true }, release(_tag: string, n: number) { held -= n } }
  const doc = { sheets: [s, { ...s, name: 'Second', sheetId: '2' }] }
  const { paintMemoryFor } = await import('../src/xlsx/cf/render')
  const a = await prepareSheetAssets(doc, { chartParent: parent })
  const before = held
  const p = await preparePrintUnits(doc, { chartParent: parent })
  expect(held).toBeGreaterThan(before)
  expect(paintMemoryFor(doc, parent).currentBytes).toBe(held)
  expect(peak).toBeLessThanOrEqual(16 * 1024 ** 2)
  p.dispose(); expect(held).toBe(before)
  a.dispose(); expect(held).toBe(0)
})
it('whole-row coverage uses the static renderer column ceiling', async () => {
  const s = sheet(), r = s.conditionalFormatting.rules[0]
  s.rows[0].cells.push({ row: 0, col: 1100, ref: 'API1', value: 10, styleIndex: 0 })
  s.conditionalFormatting.rules[0] = { ...r, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 16383 }] }
  const b = await prepareCfPaint(s)
  try { expect(b?.layer.styleAt(0, 1100, undefined)?.fillColor).toBe('FF00FF00') } finally { b?.dispose() }
})

it('a timePeriod=lastWeek rule whose stored formula is wrong/off-by-one still paints exactly lastWeek', async () => {
  const { parseXml } = await import('../src/core/xml')
  const { parseCfSheet } = await import('../src/xlsx/cf/parse')
  const xmlCanonical = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<conditionalFormatting sqref="G2:G15">` +
    `<cfRule type="timePeriod" timePeriod="lastWeek" priority="1" dxfId="0">` +
    `<formula>AND(TODAY()-ROUNDDOWN(G2,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(G2,0)&lt;(WEEKDAY(TODAY())+7))</formula>` +
    `</cfRule></conditionalFormatting></worksheet>`
  const xmlOffByOne = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<conditionalFormatting sqref="G2:G15">` +
    `<cfRule type="timePeriod" timePeriod="lastWeek" priority="1" dxfId="0">` +
    `<formula>AND(TODAY()-ROUNDDOWN(G2,0)&gt;(WEEKDAY(TODAY())-1),TODAY()-ROUNDDOWN(G2,0)&lt;(WEEKDAY(TODAY())+6))</formula>` +
    `</cfRule></conditionalFormatting></worksheet>`
  const styles = { colors: {}, dxfNodes: [parseXml('<dxf><fill><patternFill patternType="solid"><bgColor rgb="FF00FF00"/></patternFill></fill></dxf>')] }
  const memory1 = createCfMemory()
  const memory2 = createCfMemory()
  const parsedCanonical = parseCfSheet(parseXml(xmlCanonical), styles, memory1)
  const parsedOffByOne = parseCfSheet(parseXml(xmlOffByOne), styles, memory2)
  try {
    // R1c: both resolve to timePeriod=lastWeek from the attribute
    expect(parsedCanonical.value.rules[0]!.definition).toMatchObject({ kind: 'timePeriod', period: 'lastWeek' })
    expect(parsedOffByOne.value.rules[0]!.definition).toMatchObject({ kind: 'timePeriod', period: 'lastWeek' })
    // Both produce identical paint batches (evaluating the attribute, not stored formulas)
    const sheetCanonical: XlsxSheet = { name: 'S1', sheetId: '1', merges: [], mergeRanges: [], cols: [], rows: [{ index: 1, cells: [{ row: 1, col: 6, ref: 'G2', value: 46325, styleIndex: 0 }] }], conditionalFormatting: parsedCanonical.value }
    const sheetOffByOne: XlsxSheet = { name: 'S2', sheetId: '2', merges: [], mergeRanges: [], cols: [], rows: [{ index: 1, cells: [{ row: 1, col: 6, ref: 'G2', value: 46325, styleIndex: 0 }] }], conditionalFormatting: parsedOffByOne.value }
    const batchCanonical = await prepareCfPaint(sheetCanonical)
    const batchOffByOne = await prepareCfPaint(sheetOffByOne)
    try {
      expect(batchOffByOne?.layer.styleAt(1, 6, undefined)).toEqual(batchCanonical?.layer.styleAt(1, 6, undefined))
    } finally {
      batchCanonical?.dispose()
      batchOffByOne?.dispose()
    }
  } finally {
    parsedCanonical.release()
    parsedOffByOne.release()
  }
})
