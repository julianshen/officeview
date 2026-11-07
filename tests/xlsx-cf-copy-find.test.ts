// @vitest-environment node
/**
 * Task 9: copy/find use bounded per-batch CF preparation in resident and paged
 * modes, including cells outside the current viewport. Expected strings follow
 * the stage1-numfmt-accessibility seed hypothesis (`1.25` via CF DXF `0.00`,
 * control `1.3` via cell `0.0`); the case is recipe-pending, never measured.
 */
import { describe, expect, it } from 'vitest'
import { WorkbookFind, type FindSheetInput } from '../src/xlsx/view/find'
import { serializeSelection, type CopyInput } from '../src/xlsx/view/copy'
import { cfTextScopeFromWire, prepareCfTextScope, CF_TEXT_SCOPE_GRANT_BYTES } from '../src/xlsx/cf/text'
import { cellDisplayText } from '../src/xlsx/render'
import type { AuthoritativeCell, SourceStore, WorkbookOwner } from '../src/xlsx/view/contracts'
import type { FormatAllocation } from '../src/xlsx/formula/format'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import type { EffectiveFormatScope } from '../src/xlsx/view/formatWorkspace'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { initialWorkbookOwner } from '../src/worker/viewerProtocol'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS } from '../src/xlsx/view/coordinates'
import { DocumentSession } from '../src/worker/viewerSession'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { xlsxBytes } from './helpers/xlsx-package'
import { drain } from './helpers/xlsx-find'

const cellIsRule = {
  id: 0,
  priority: 1,
  documentOrder: 0,
  stopIfTrue: false,
  ranges: [{ firstRow: 0, lastRow: 600, firstCol: 0, lastCol: 1 }],
  writes: new Set(['numFmt']),
  dxf: { 'numFmt.code': '0.00' },
  definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 1 }], dialect: 'base', formulas: [] },
}

function cfDoc(): XlsxDocument {
  const cell = (row: number, col: number, value: number): XlsxCell => ({
    ref: `R${row}C${col}`, row, col, styleIndex: 0, style: { numFmtId: 0, formatCode: '0.0' }, value, hasCachedValue: true,
  })
  return {
    sheets: [{
      name: 'CF',
      sheetId: 'cf',
      rows: [
        { index: 0, cells: [cell(0, 0, 1.25), cell(0, 1, 0.5)] },
        { index: 500, cells: [cell(500, 0, 1.25)] },
      ],
      cols: [],
      merges: [],
      mergeRanges: [],
      conditionalFormatting: { diagnostics: [], rules: [cellIsRule] },
    }],
    tables: [],
  } as unknown as XlsxDocument
}

const docCell = (sheet: XlsxSheet, row: number, col: number): XlsxCell | undefined =>
  sheet.rows.find(candidate => candidate.index === row)?.cells.find(candidate => candidate.col === col)

function residentSheets(docId: string, source: { sheets: readonly { sheetId: string }[]; store: (sheetId: string) => SourceStore; geometry: (sheetId: string) => { baseHidden: { intervals: readonly { r0: number; r1: number }[] } } }, owner: WorkbookOwner): FindSheetInput[] {
  return source.sheets.map(({ sheetId }) => {
    const entry = owner.sheetVector.find(item => item.sheetId === sheetId)!
    const authored = source.geometry(sheetId).baseHidden.intervals.map(({ r0, r1 }) => ({ r0, r1 }))
    return {
      sheetId,
      store: source.store(sheetId),
      identity: { docId, sheetId, valueGen: entry.valueGen, fontGen: owner.fontGen, viewGen: entry.revision },
      order: createIdentityOrder(),
      visibility: visibilityFromHidden(createIntervalSet(authored)),
    }
  })
}

const residentPrepare = (sheet: XlsxSheet) =>
  (sheetId: string, range: { r0: number; r1: number; c0: number; c1: number }, signal: AbortSignal) => {
    if (sheetId !== sheet.sheetId) return Promise.resolve(undefined)
    return prepareCfTextScope(sheet, {
      rects: [{ firstRow: range.r0, lastRow: range.r1, firstCol: range.c0, lastCol: range.c1 }],
      signal,
      locale: 'en-US',
    })
  }

const residentTextOf = (sheet: XlsxSheet) =>
  (cell: AuthoritativeCell, _sheetId: string, allocation?: FormatAllocation, scope?: EffectiveFormatScope): string => {
    const base = docCell(sheet, cell.row, cell.col)?.style
    const style = scope?.effectiveStyle(cell.row, cell.col, base) ?? base
    return cellDisplayText(sheet, { value: cell.value, style }, allocation)
  }

describe('resident copy/find follow the CF number format outside the viewport', () => {
  it('finds 1.25 (not 1.3) in near and far cells', async () => {
    const doc = cfDoc()
    const source = createObjectWorkbookSource(doc, 'cf-doc')
    const owner = initialWorkbookOwner('cf-doc', source.sheets)
    const sheet = doc.sheets[0]! as unknown as XlsxSheet
    const session = new WorkbookFind(
      { owner, sheets: residentSheets('cf-doc', source, owner) },
      { text: '1.25', caseSensitive: false },
      { textOf: residentTextOf(sheet), prepareScope: residentPrepare(sheet) },
    )
    const hits = await drain(session)
    expect(hits.map(hit => [hit.source.row, hit.source.col, hit.text])).toEqual([[0, 0, '1.25'], [500, 0, '1.25']])

    const stale = new WorkbookFind(
      { owner, sheets: residentSheets('cf-doc', source, owner) },
      { text: '1.3', caseSensitive: false },
      { textOf: residentTextOf(sheet), prepareScope: residentPrepare(sheet) },
    )
    expect((await drain(stale)).map(hit => [hit.source.row, hit.source.col])).toEqual([])
  })

  it('copies CF-formatted text in near and far selections', async () => {
    const doc = cfDoc()
    const source = createObjectWorkbookSource(doc, 'cf-copy')
    const owner = initialWorkbookOwner('cf-copy', source.sheets)
    const entry = owner.sheetVector[0]!
    const sheet = doc.sheets[0]! as unknown as XlsxSheet
    const store = source.store('cf')
    const base: CopyInput = {
      store,
      owner: { ...store.stamp, fontGen: 0, viewGen: entry.revision },
      currentOwner: () => ({ ...store.stamp, fontGen: 0, viewGen: entry.revision }),
      selection: { active: null, anchor: null, ranges: [] },
      order: createIdentityOrder(),
      visibility: visibilityFromHidden(createIntervalSet([])),
      columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
      allowed: () => true,
      signal: new AbortController().signal,
      format: (cell, allocation, scope) => {
        const base = docCell(sheet, cell.row, cell.col)?.style
        return cellDisplayText(sheet, { value: cell.value, style: scope?.effectiveStyle(cell.row, cell.col, base) ?? base }, allocation)
      },
      prepareScope: (range, signal) => residentPrepare(sheet)('cf', range, signal),
    }
    const near = await serializeSelection({ ...base, selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 1 }] }] } })
    expect(near).toMatchObject({ kind: 'ready' })
    expect(near.kind === 'ready' ? near.text : '').toBe('1.25\t0.5')
    const far = await serializeSelection({ ...base, selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 500, r1: 500 }] }, cols: [{ c0: 0, c1: 0 }] }] } })
    expect(far).toMatchObject({ kind: 'ready' })
    expect(far.kind === 'ready' ? far.text : '').toBe('1.25')
  })
})

/** Paged workbook: A1=1.25 and A501=1.25 under a cellIs>1 rule with DXF 0.00. */
function pagedBytes(): Uint8Array {
  return xlsxBytes({
    sheetData: '<row r="1"><c r="A1" s="0"><v>1.25</v></c><c r="B1" s="0"><v>0.5</v></c></row><row r="501"><c r="A501" s="0"><v>1.25</v></c></row>',
    styles: '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0"/></numFmts><cellXfs count="1"><xf numFmtId="164"/></cellXfs><dxfs count="1"><dxf><numFmt numFmtId="164" formatCode="0.00"/></dxf></dxfs>',
    afterSheetData: '<conditionalFormatting sqref="A1:A501"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>',
  })
}

describe('paged copy/find follow the CF number format outside the viewport', () => {
  it('worker find hits CF-formatted text in near and far cells', async () => {
    const bytes = pagedBytes()
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer]) }, { docId: 'cf-paged-find', profile: 'D' })
    try {
      await session.done
      const worker = new DocumentSession('cf-paged-find', createPagedWorkbookSource(session))
      try {
        const result = await worker.search(worker.workbookOwner(), { text: '1.25', caseSensitive: false }, undefined, new AbortController().signal)
        expect(result).toMatchObject({ complete: true })
        if (!('hits' in result)) throw new Error('search did not complete')
        expect(result.hits.map(hit => [hit.source.row, hit.source.col, hit.text]).sort()).toEqual([[0, 0, '1.25'], [500, 0, '1.25']])
      } finally {
        await worker.release()
      }
    } finally {
      await session.release()
    }
  })

  it('main-thread paged copy formats near and far selections through wire states', async () => {
    const bytes = pagedBytes()
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer]) }, { docId: 'cf-paged-copy', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session)
      const sheetId = source.sheets[0]!.sheetId
      const store = source.store(sheetId)
      const identity = { docId: 'cf-paged-copy', sheetId, valueGen: store.stamp.valueGen, fontGen: 0, viewGen: 0 }
      const prepareScope = async (range: { r0: number; r1: number; c0: number; c1: number }, signal: AbortSignal) => {
        if (!store.readConditionalFormatting) return undefined
        const result = await store.readConditionalFormatting({
          rects: [{ r0: range.r0, r1: range.r1, c0: range.c0, c1: range.c1 }],
          stamp: { owner: 'cf-paged-copy', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 },
          grantBytes: CF_TEXT_SCOPE_GRANT_BYTES,
        }, undefined, { signal })
        if (result.kind !== 'ready') return undefined
        return cfTextScopeFromWire(result, {})
      }
      const format = (cell: AuthoritativeCell, allocation?: FormatAllocation, scope?: EffectiveFormatScope): string =>
        source.displayText!(cell, sheetId, allocation, undefined, scope)
      const base: CopyInput = {
        store,
        owner: identity,
        currentOwner: () => identity,
        selection: { active: null, anchor: null, ranges: [] },
        order: createIdentityOrder(),
        visibility: visibilityFromHidden(createIntervalSet([])),
        columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
        allowed: () => true,
        signal: new AbortController().signal,
        format,
        prepareScope,
      }
      const near = await serializeSelection({ ...base, selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 1 }] }] } })
      expect(near).toMatchObject({ kind: 'ready' })
      expect(near.kind === 'ready' ? near.text : '').toBe('1.25\t0.5')
      const far = await serializeSelection({ ...base, selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 500, r1: 500 }] }, cols: [{ c0: 0, c1: 0 }] }] } })
      expect(far).toMatchObject({ kind: 'ready' })
      expect(far.kind === 'ready' ? far.text : '').toBe('1.25')
    } finally {
      await session.release()
    }
  })
})
