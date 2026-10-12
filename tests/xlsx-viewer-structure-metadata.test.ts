/**
 * S3.3 parser structural metadata extraction and wiring: real XLSX bytes with
 * DrawingML anchors, saved/engine dynamic-array spills, fixed arrays, tables
 * (headerRowCount=0, totals) and giant merges produce identical structural
 * records in the resident (object) and paged adapters, and the worker session
 * feeds them to the S3.2 guard: an unaffected row band is allowed, a band that
 * would displace an anchor or split a spill is refused with the typed reason,
 * and unreadable structure keeps the operation waiting (never allowed).
 */
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { StructuralIndex } from '../src/xlsx/view/structure'
import { initialIdentity } from '../src/worker/viewerProtocol'
import { ViewerKernel } from '../src/worker/viewerKernel'
import type { FilterEvaluator, WorkbookSource } from '../src/worker/viewerSession'
import type { FilterSpec, SortSpec, SourceStore, StructuralResult, WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { DRAWING_PART, EMU_PER_PX, absolute, dataRows, oneCell, structureBytes, twoCell, wrappedAnchor } from './helpers/xlsx-structure-fixtures'

const FULL = { r0: 0, r1: 1_048_575, c0: 0, c1: 16_383 }
const opts = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })

async function adapters(bytes: Uint8Array): Promise<{ cached: WorkbookSource; engine: WorkbookSource; paged: WorkbookSource }> {
  const cached = createObjectWorkbookSource(await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' }), 'doc')
  const engine = createObjectWorkbookSource(await parseXlsx(await OfficePackage.load(bytes)), 'doc')
  const store = new SourcePageStore({ docId: 'doc', backend: memoryPageBackend() })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'doc', profile: 'D', store, pageRows: 8 })
  expect((await session.done).kind).toBe('complete')
  return { cached, engine, paged: createPagedWorkbookSource(session) }
}
async function structural(source: WorkbookSource, rect = FULL): Promise<StructuralResult> {
  const store = source.store('1')
  return store.queryStructural(rect, opts(store))
}
const strip = (result: StructuralResult) => ({ ...result, stamp: null })

// ---- fixtures -------------------------------------------------------------

/** Row 10 (index 9) is 30pt tall and row 11 (index 10) is hidden, so the
 * one-cell anchor's 50 px extent ends in row index 11: 180+40+0 → 220..240. */
const DRAWINGS = structureBytes({
  sheetData: dataRows(30, () => '', r => r === 9 ? ' ht="30" customHeight="1"' : r === 10 ? ' hidden="1"' : ''),
  drawing: oneCell(3, 9, 100, 50, 2) + twoCell([3, 14], [5, 16], 3) + absolute(0, 0, 10, 10, 4),
})
const ONE_CELL = `${DRAWING_PART}#0`, TWO_CELL = `${DRAWING_PART}#1`, ABSOLUTE = `${DRAWING_PART}#2`
const EXPECTED_DRAWINGS = [
  { id: ONE_CELL, kind: 'oneCell', from: { row: 9, col: 3, offsetXEmu: 0, offsetYEmu: 0 }, extent: { x: 0, y: 0, width: 100 * EMU_PER_PX, height: 50 * EMU_PER_PX }, bounds: { r0: 9, r1: 11, c0: 3, c1: 4 } },
  {
    id: TWO_CELL, kind: 'twoCell', from: { row: 14, col: 3, offsetXEmu: 0, offsetYEmu: 0 }, to: { row: 16, col: 5, offsetXEmu: 0, offsetYEmu: 0 },
    extent: { x: 0, y: 0, width: 128 * EMU_PER_PX, height: 40 * EMU_PER_PX }, bounds: { r0: 14, r1: 16, c0: 3, c1: 5 },
  },
  { id: ABSOLUTE, kind: 'absolute', extent: { x: 0, y: 0, width: 10 * EMU_PER_PX, height: 10 * EMU_PER_PX } },
]

/** Saved dynamic spill D20:D23 (children carry saved values, no formula) and a fixed array F25:G25. */
const SAVED_SPILLS = structureBytes({
  dynamic: true,
  sheetData: dataRows(30, r => r === 19 ? '<c r="D20" cm="1"><f t="array" ref="D20:D23">_xlfn.SEQUENCE(4)</f><v>1</v></c>'
    : r >= 20 && r <= 22 ? `<c r="D${r + 1}"><v>${r - 18}</v></c>`
      : r === 24 ? '<c r="F25"><f t="array" ref="F25:G25">{1,2}</f><v>1</v></c><c r="G25"><v>2</v></c>' : ''),
})
const SAVED_SPILL_RECORDS = [
  { anchor: { row: 19, col: 3 }, kind: 'dynamic', saved: { r0: 19, r1: 22, c0: 3, c1: 3 }, evaluated: { r0: 19, r1: 22, c0: 3, c1: 3 } },
  { anchor: { row: 24, col: 5 }, kind: 'fixed', saved: { r0: 24, r1: 24, c0: 5, c1: 6 } },
]

/** No cached value: the engine recalculates SEQUENCE(6) past the saved D20:D23 footprint. */
const ENGINE_SPILL = structureBytes({
  dynamic: true,
  sheetData: dataRows(30, r => r === 19 ? '<c r="D20" cm="1"><f t="array" ref="D20:D23">_xlfn.SEQUENCE(6)</f></c>' : ''),
})

const TABLE_XML = '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="7" name="T1" displayName="T1" ref="A1:B6" headerRowCount="0" totalsRowCount="1"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>'
const TABLES = structureBytes({ sheetData: dataRows(10), tables: [TABLE_XML] })

const GIANT_MERGE = structureBytes({ sheetData: dataRows(30), afterSheetData: '<mergeCells count="1"><mergeCell ref="D25:XFD1048576"/></mergeCells>' })

// ---- structural records in both adapters ---------------------------------

describe('S3.3 structural records from real XLSX bytes (resident and paged, identical)', () => {
  it('drawing anchors: oneCell resolves its occupied bounds through custom/hidden rows; twoCell and absolute keep their kind', async () => {
    const { cached, engine, paged } = await adapters(DRAWINGS)
    const a = await structural(cached)
    expect(a.completeness).toBe('complete')
    expect(a.missing).toEqual([])
    expect(a.drawings).toEqual(EXPECTED_DRAWINGS)
    expect(strip(await structural(paged))).toEqual(strip(a))
    expect(strip(await structural(engine))).toEqual(strip(a))
    // A narrow rect still finds the anchor by its occupied band, not only its from-row.
    expect((await structural(paged, { r0: 11, r1: 11, c0: 4, c1: 4 })).drawings.map(d => d.id)).toEqual([ONE_CELL, ABSOLUTE])
    for (const source of [cached, paged]) expect(source.coverageDiagnostics?.('1')).toEqual([])
    await paged.release()
  })

  it('anchors are placement-only: unreadable picture content and rows past the static paint cap still yield records', async () => {
    const brokenPicture = '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="9" name="p"/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rMissing"/></xdr:blipFill><xdr:spPr/></xdr:pic>'
    const bytes = structureBytes({
      sheetData: dataRows(5),
      drawing: oneCell(2, 20_000, 64, 20) + `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>${brokenPicture}<xdr:clientData/></xdr:twoCellAnchor>`,
    })
    const { cached, paged } = await adapters(bytes)
    const a = await structural(cached)
    expect(a.completeness).toBe('complete')
    expect(a.drawings.map(d => [d.id, d.kind, d.kind === 'absolute' ? null : d.bounds])).toEqual([
      // An extent ending exactly on a cell boundary keeps the next cell (over-approximation is the safe side).
      [`${DRAWING_PART}#0`, 'oneCell', { r0: 20_000, r1: 20_001, c0: 2, c1: 3 }],
      [`${DRAWING_PART}#1`, 'twoCell', { r0: 2, r1: 3, c0: 0, c1: 1 }],
    ])
    expect(strip(await structural(paged))).toEqual(strip(a))
    await paged.release()
  })

  it('saved dynamic spill and fixed array extents are complete and identical under the cached policy', async () => {
    const { cached, paged } = await adapters(SAVED_SPILLS)
    const a = await structural(cached)
    expect(a.completeness).toBe('complete')
    expect(a.spills).toEqual(SAVED_SPILL_RECORDS)
    expect(strip(await structural(paged))).toEqual(strip(a))
    // Child-only rect: the anchor lies outside, the extent still matches.
    expect((await structural(paged, { r0: 21, r1: 21, c0: 0, c1: 16_383 })).spills.map(s => s.anchor)).toEqual([{ row: 19, col: 3 }])
    await paged.release()
  })

  it('an engine-evaluated spill reports its evaluated extent beyond the saved footprint (object, engine policy)', async () => {
    const { engine, paged } = await adapters(ENGINE_SPILL)
    const a = await structural(engine)
    expect(a.completeness).toBe('complete')
    expect(a.spills).toEqual([{ anchor: { row: 19, col: 3 }, kind: 'dynamic', saved: { r0: 19, r1: 22, c0: 3, c1: 3 }, evaluated: { r0: 19, r1: 24, c0: 3, c1: 3 } }])
    // Enumerated policy difference: the cached source shows only the saved footprint.
    expect((await structural(paged)).spills).toEqual([{ anchor: { row: 19, col: 3 }, kind: 'dynamic', saved: { r0: 19, r1: 22, c0: 3, c1: 3 }, evaluated: { r0: 19, r1: 22, c0: 3, c1: 3 } }])
    await paged.release()
  })

  it('tables keep headerRowCount=0 and totals rows; giant merges stay one interval record', async () => {
    const t = await adapters(TABLES)
    const table = await structural(t.cached)
    expect(table.tables).toEqual([{ id: '7', range: { r0: 0, r1: 5, c0: 0, c1: 1 }, headerRows: 0, totalsRows: 1 }])
    expect(strip(await structural(t.paged))).toEqual(strip(table))
    await t.paged.release()
    const m = await adapters(GIANT_MERGE)
    const merge = await structural(m.cached)
    expect(merge.merges).toEqual([{ id: 'merge:24:3:1048575:16383', range: { r0: 24, r1: 1_048_575, c0: 3, c1: 16_383 }, anchor: { row: 24, col: 3 } }])
    expect(strip(await structural(m.paged))).toEqual(strip(merge))
    expect(new StructuralIndex(merge).recordCount).toBe(1)
    await m.paged.release()
  })

  it('unreadable structure stays incomplete in both adapters: invalid anchor, wrapped anchor, dangling drawing, dynamic array without ref', async () => {
    const cases: Array<[string, Uint8Array, { r0: number; r1: number; c0: number; c1: number }]> = [
      ['invalid anchor', structureBytes({ sheetData: dataRows(5), drawing: '<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:clientData/></xdr:oneCellAnchor>' }), FULL],
      ['wrapped anchor', structureBytes({ sheetData: dataRows(5), drawing: wrappedAnchor(oneCell(1, 1, 10, 10)) }), FULL],
      ['dangling drawing', structureBytes({ sheetData: dataRows(5), danglingDrawing: true }), FULL],
      ['drawing part with a foreign root', structureBytes({ sheetData: dataRows(5), drawingPart: '<?xml version="1.0"?><!-- not a drawing --><notADrawing/>' }), FULL],
      ['dynamic without ref', structureBytes({ dynamic: true, sheetData: dataRows(5, r => r === 2 ? '<c r="D3" cm="1"><f t="array">_xlfn.SEQUENCE(2)</f><v>1</v></c>' : '') }), { r0: 2, r1: 1_048_575, c0: 3, c1: 16_383 }],
    ]
    for (const [name, bytes, missing] of cases) {
      const { cached, paged } = await adapters(bytes)
      const a = await structural(cached), b = await structural(paged)
      expect(a.completeness, name).toBe('incomplete')
      expect(a.missing, name).toContainEqual(missing)
      expect(strip(b), name).toEqual(strip(a))
      expect(cached.coverageDiagnostics?.('1').map(d => d.code), name).toContain('coverage-unresolvable')
      await paged.release()
    }
  })
})

describe('S3.3 fail-closed sources without a known anchor or spill extent', () => {
  it('paged: anchors beyond the structural record cap are unresolved, not dropped', async () => {
    const bytes = structureBytes({ sheetData: dataRows(5), drawing: oneCell(1, 1, 10, 10, 2) + oneCell(1, 3, 10, 10, 3) })
    const store = new SourcePageStore({ docId: 'doc', backend: memoryPageBackend() })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'doc', profile: 'D', store, limits: { paged: { structuralRecords: 1 } } })
    expect((await session.done).kind).toBe('complete')
    const paged = createPagedWorkbookSource(session)
    const result = await structural(paged)
    expect(result.completeness).toBe('incomplete')
    expect(result.missing).toContainEqual(FULL)
    expect(result.drawings.map(d => d.id)).toEqual([`${DRAWING_PART}#0`])
    expect(paged.coverageDiagnostics?.('1').map(d => d.code)).toContain('coverage-unresolvable')
    await paged.release()
  })

  it('object, engine policy without evaluation state: a dynamic array extent stays unknown even with a saved ref', async () => {
    const doc = await parseXlsx(await OfficePackage.load(structureBytes({ sheetData: dataRows(5) })))
    // A hand-built dynamic anchor the engine never evaluated (no spill state for this document).
    const cell = doc.sheets[0].rows[2].cells[0]
    Object.assign(cell, { formula: '_xlfn.SEQUENCE(2)', dynamicArray: true, arrayRef: 'A3:A4' })
    // A new document object carries no engine spill state (that state is keyed by document identity).
    const fresh = { ...doc, sheets: doc.sheets.map(sheet => ({ ...sheet })) }
    const source = createObjectWorkbookSource(fresh, 'doc')
    expect(source.valuePolicy).toBe('engine')
    const result = await structural(source)
    expect(result.completeness).toBe('incomplete')
    expect(result.missing).toContainEqual({ r0: 2, r1: 1_048_575, c0: 0, c1: 16_383 })
    expect(source.coverageDiagnostics?.('1')).toContainEqual(expect.objectContaining({ code: 'coverage-unresolvable', source: { row: 2, col: 0 } }))
  })

  it('object: a sheet with parsed drawings but no extracted anchors is unresolved', async () => {
    const doc = await parseXlsx(await OfficePackage.load(DRAWINGS), { valuePolicy: 'cached' })
    expect(doc.sheets[0].drawings?.length).toBeGreaterThan(0)
    delete doc.sheets[0].drawingAnchors
    const source = createObjectWorkbookSource(doc, 'doc')
    const result = await structural(source)
    expect(result.completeness).toBe('incomplete')
    expect(result.missing).toContainEqual(FULL)
    expect(source.coverageDiagnostics?.('1').map(d => d.code)).toContain('coverage-unresolvable')
  })
})

// ---- through the worker session (S3.2 guard) ------------------------------

const quiet = { yield: async () => {} }
let nextId = 0
const rid = () => `s33-${++nextId}`
async function run<K extends WorkerCommand['kind']>(kernel: ViewerKernel, command: Extract<WorkerCommand, { kind: K }>): Promise<WorkerReply<K>> {
  const replies: WorkerReply[] = []
  await kernel.handle(command, out => { replies.push(out) })
  return replies[replies.length - 1] as WorkerReply<K>
}
const sort = (r0: number, r1: number, extra: Partial<SortSpec['range']> = {}): SortSpec => ({
  range: { r0, r1, c0: 0, c1: 1, headerRows: 0, totalsRows: 0, ...extra }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false,
})
let hiddenRow = -1
/** Hides one chosen source row; everything else matches. */
const hideOne: FilterEvaluator = async () => ({ kind: 'complete', matches: (row: number) => row !== hiddenRow })
const filter = (r0: number, r1: number): FilterSpec => ({ range: { r0, r1, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: -1 }]]) })

/** Opens the same bytes resident (object parse) or paged (admission forced to paged). */
async function open(mode: 'object' | 'paged', bytes: Uint8Array) {
  const kernel = new ViewerKernel({
    parse: async (input: Uint8Array) => parseXlsx(await OfficePackage.load(input)),
    ...(mode === 'paged' ? { importLimits: { object: { rows: 1 } } } : {}),
  }, { ...quiet, filterEvaluator: hideOne })
  const docId = `${mode}-${++nextId}`
  const out = await run(kernel, { kind: 'open', requestId: rid(), docId, source: { kind: 'blob', blob: new Blob([bytes as BlobPart]) }, profile: 'D' })
  expect(out.kind, JSON.stringify(out)).toBe('complete')
  if (out.kind === 'complete' && out.result.kind === 'open') expect(out.result.valuePolicy).toBe(mode === 'paged' ? 'cached' : 'engine')
  const owner = initialIdentity(docId, '1')
  // A paged open resolves at the first committed page; structure is final once the sheet completes.
  for (let i = 0; mode === 'paged' && i < 200; i++) {
    const read = await run(kernel, { kind: 'read', requestId: rid(), owner, query: { op: 'structural', rect: FULL } })
    if (read.kind === 'complete' && read.result.kind === 'read' && read.result.result.op === 'structural' && read.result.result.result.completeness !== 'loading') break
    await new Promise(resolve => setTimeout(resolve, 1))
  }
  return { kernel, owner }
}
async function sortOutcome(mode: 'object' | 'paged', bytes: Uint8Array, spec: SortSpec) {
  const { kernel, owner } = await open(mode, bytes)
  return run(kernel, { kind: 'sort', requestId: rid(), owner, spec })
}
async function filterOutcome(mode: 'object' | 'paged', bytes: Uint8Array, spec: FilterSpec, hide: number) {
  const { kernel, owner } = await open(mode, bytes)
  hiddenRow = hide
  return run(kernel, { kind: 'filter', requestId: rid(), owner, spec, nowMs: 0 })
}
const committed = (out: WorkerReply) => expect(out, JSON.stringify(out)).toMatchObject({ kind: 'complete', result: { kind: 'view' } })
const refused = (out: WorkerReply, refusal: object) =>
  expect(out, JSON.stringify(out)).toMatchObject({ kind: 'error', code: 'operation-refused', retryable: false, refusal })

for (const mode of ['object', 'paged'] as const) {
  describe(`S3.3 guard decisions through the worker session (${mode})`, () => {
    it('sort: a band without anchors is allowed; a band over a cell anchor is refused drawing-impact; absolute never blocks', async () => {
      committed(await sortOutcome(mode, DRAWINGS, sort(0, 5)))
      refused(await sortOutcome(mode, DRAWINGS, sort(0, 9)), { reason: 'drawing-impact', drawingId: ONE_CELL })
      // Hidden row 10 sits inside the one-cell span: the occupied band, not only the from-row, counts.
      refused(await sortOutcome(mode, DRAWINGS, sort(11, 13)), { reason: 'drawing-impact', drawingId: ONE_CELL })
      refused(await sortOutcome(mode, DRAWINGS, sort(15, 20)), { reason: 'drawing-impact', drawingId: TWO_CELL })
      committed(await sortOutcome(mode, DRAWINGS, sort(18, 29)))
    })

    it('filter: hiding a row above an anchor is refused (AC-SEM-05); hiding below every anchor is allowed', async () => {
      refused(await filterOutcome(mode, DRAWINGS, filter(0, 29), 2), { reason: 'drawing-impact', drawingId: ONE_CELL })
      committed(await filterOutcome(mode, DRAWINGS, filter(18, 29), 25))
      // The authored-hidden row 10 stays hidden outside the filter range, so the committed
      // geometry is not "all visible": a filter below the anchors leaves every span unchanged.
      committed(await filterOutcome(mode, DRAWINGS, filter(12, 29), 25))
      // Inside a filter range visibility is the filter's: unhiding row 10 grows the one-cell span.
      refused(await filterOutcome(mode, DRAWINGS, filter(0, 29), 25), { reason: 'drawing-impact', drawingId: ONE_CELL })
    })

    it('saved spills: an unaffected band is allowed; child-only dynamic and fixed-array bands are refused', async () => {
      committed(await sortOutcome(mode, SAVED_SPILLS, sort(0, 10)))
      refused(await sortOutcome(mode, SAVED_SPILLS, sort(21, 22)), { reason: 'dynamic-spill', range: { r0: 19, r1: 22, c0: 3, c1: 3 } })
      refused(await sortOutcome(mode, SAVED_SPILLS, sort(24, 26)), { reason: 'fixed-array', range: { r0: 24, r1: 24, c0: 5, c1: 6 } })
      refused(await filterOutcome(mode, SAVED_SPILLS, filter(15, 29), 16), { reason: 'dynamic-spill' })
    })

    it('tables with headerRowCount=0 sort their body; a giant merge refuses only the bands it covers', async () => {
      committed(await sortOutcome(mode, TABLES, sort(0, 5, { totalsRows: 1 })))
      committed(await sortOutcome(mode, GIANT_MERGE, sort(0, 10)))
      refused(await sortOutcome(mode, GIANT_MERGE, sort(20, 29)), { reason: 'merge', range: { r0: 24, r1: 1_048_575, c0: 3, c1: 16_383 } })
    })

    it('unreadable drawing structure keeps sort waiting as incomplete, never allowed', async () => {
      const wrapped = structureBytes({ sheetData: dataRows(5), drawing: wrappedAnchor(oneCell(1, 1, 10, 10)) })
      expect(await sortOutcome(mode, wrapped, sort(0, 4))).toMatchObject({ kind: 'incomplete', commandKind: 'sort' })
    })
  })
}

describe('S3.3 engine spill through the resident session', () => {
  it('a sort band touching only engine-generated children is refused with the evaluated extent', async () => {
    refused(await sortOutcome('object', ENGINE_SPILL, sort(23, 26)), { reason: 'dynamic-spill', range: { r0: 19, r1: 24, c0: 3, c1: 3 } })
    committed(await sortOutcome('object', ENGINE_SPILL, sort(0, 10)))
  })
})
