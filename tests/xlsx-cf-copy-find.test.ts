// @vitest-environment node
/**
 * Task 9: copy/find use bounded per-batch CF preparation in resident and paged
 * modes, including cells outside the current viewport. Expected strings follow
 * the stage1-numfmt-accessibility seed hypothesis (`1.25` via CF DXF `0.00`,
 * control `1.3` via cell `0.0`); the case is recipe-pending, never measured.
 * Sparse batches prepare exact-cover rects, never dense hulls, so scanning and
 * hit materialization always agree.
 */
import { describe, expect, it } from 'vitest'
import { WorkbookFind, type FindSheetInput } from '../src/xlsx/view/find'
import { serializeSelection, type CopyInput } from '../src/xlsx/view/copy'
import { prepareCfTextScope, prepareCfWireScope, CF_TEXT_SCOPE_GRANT_BYTES } from '../src/xlsx/cf/text'
import { cellDisplayText } from '../src/xlsx/render'
import type { AuthoritativeCell, SourceStore, WorkbookOwner } from '../src/xlsx/view/contracts'
import type { FormatAllocation } from '../src/xlsx/formula/format'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import type { EffectiveFormatScope, FormatBudget } from '../src/xlsx/view/formatWorkspace'
import { FindError } from '../src/xlsx/view/find'
import type { SourceRect } from '../src/xlsx/view/coordinates'
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

function trackingBudget(refuse = false): { budget: FormatBudget; net: () => number; retained: () => number } {
  let held = 0
  let kept = 0
  const budget: FormatBudget = {
    reserve: bytes => { held += bytes; if (refuse) { held -= bytes; throw new FindError('resource-limit', 'refused') } },
    release: bytes => { held -= bytes },
    retain: (_targets, bytes, commit) => { kept += bytes; commit() },
    settle: () => {},
  }
  return { budget, net: () => held, retained: () => kept }
}

/** Refuses once held bytes would exceed the limit (deterministic pressure). */
function limitedBudget(limit: number): { budget: FormatBudget; net: () => number } {
  let held = 0
  const budget: FormatBudget = {
    reserve: bytes => {
      if (held + bytes > limit) throw new FindError('resource-limit', 'refused')
      held += bytes
    },
    release: bytes => { held -= bytes },
  }
  return { budget, net: () => held }
}

const cellIsRule = {
  id: 0,
  priority: 1,
  documentOrder: 0,
  stopIfTrue: false,
  ranges: [{ firstRow: 0, lastRow: 100000, firstCol: 0, lastCol: 1 }],
  writes: new Set(['numFmt']),
  dxf: { 'numFmt.code': '0.00' },
  definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 1 }], dialect: 'base', formulas: [] },
}

function cfDoc(farRow = 500): XlsxDocument {
  const cell = (row: number, col: number, value: number): XlsxCell => ({
    ref: `R${row}C${col}`, row, col, styleIndex: 0, style: { numFmtId: 0, formatCode: '0.0' }, value, hasCachedValue: true,
  })
  return {
    sheets: [{
      name: 'CF',
      sheetId: 'cf',
      rows: [
        { index: 0, cells: [cell(0, 0, 1.25), cell(0, 1, 0.5)] },
        { index: farRow, cells: [cell(farRow, 0, 1.25)] },
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

const residentPrepare = (sheet: XlsxSheet, budget: FormatBudget) =>
  (sheetId: string, rects: readonly SourceRect[], signal: AbortSignal) => {
    if (sheetId !== sheet.sheetId) return Promise.resolve(undefined)
    return prepareCfTextScope(sheet, { rects, signal, locale: 'en-US', budget })
  }

const residentTextOf = (sheet: XlsxSheet) =>
  (cell: AuthoritativeCell, _sheetId: string, allocation?: FormatAllocation, scope?: EffectiveFormatScope): string => {
    const base = docCell(sheet, cell.row, cell.col)?.style
    const style = scope?.effectiveStyle(cell.row, cell.col, base) ?? base
    return cellDisplayText(sheet, { value: cell.value, style }, allocation)
  }

describe('resident copy/find follow the CF number format outside the viewport', () => {
  it('finds 1.25 (not 1.3) in near and far cells', async () => {
    const { budget, net, retained } = trackingBudget()
    const doc = cfDoc()
    const source = createObjectWorkbookSource(doc, 'cf-doc')
    const owner = initialWorkbookOwner('cf-doc', source.sheets)
    const sheet = doc.sheets[0]! as unknown as XlsxSheet
    const session = new WorkbookFind(
      { owner, sheets: residentSheets('cf-doc', source, owner) },
      { text: '1.25', caseSensitive: false },
      { textOf: residentTextOf(sheet), prepareScope: residentPrepare(sheet, budget), formatBudget: budget },
    )
    const hits = await drain(session)
    expect(hits.map(hit => [hit.source.row, hit.source.col, hit.text])).toEqual([[0, 0, '1.25'], [500, 0, '1.25']])
    // Everything held is owned by the emitted hits: no unowned charge remains.
    expect(net()).toBe(retained())
    expect(retained()).toBe(hits.reduce((sum, hit) => sum + hit.text.length * 2 + 128, 0))

    const stale = new WorkbookFind(
      { owner, sheets: residentSheets('cf-doc', source, owner) },
      { text: '1.3', caseSensitive: false },
      { textOf: residentTextOf(sheet), prepareScope: residentPrepare(sheet, budget), formatBudget: budget },
    )
    expect((await drain(stale)).map(hit => [hit.source.row, hit.source.col])).toEqual([])
    // Still only the first session's hits stay charged; the empty search leaks nothing.
    expect(net()).toBe(retained())
  })

  it('sparse batches prepare exact cover, never a dense hull', async () => {
    const { budget, net, retained } = trackingBudget()
    const doc = cfDoc(10_000)
    const source = createObjectWorkbookSource(doc, 'cf-sparse')
    const owner = initialWorkbookOwner('cf-sparse', source.sheets)
    const sheet = doc.sheets[0]! as unknown as XlsxSheet
    const seen: SourceRect[][] = []
    const session = new WorkbookFind(
      { owner, sheets: residentSheets('cf-sparse', source, owner) },
      { text: '1.25', caseSensitive: false },
      {
        textOf: residentTextOf(sheet),
        formatBudget: budget,
        prepareScope: (sheetId, rects, signal) => {
          seen.push([...rects])
          return residentPrepare(sheet, budget)(sheetId, rects, signal)
        },
      },
    )
    const hits = await drain(session)
    // Three populated cells tens of thousands of rows apart: scan and
    // single-cell materialization agree, with no stale-source error.
    expect(hits.map(hit => [hit.source.row, hit.source.col, hit.text])).toEqual([[0, 0, '1.25'], [10000, 0, '1.25']])
    for (const rects of seen) {
      const area = rects.reduce((sum, rect) => sum + (rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1), 0)
      expect(area).toBeLessThanOrEqual(5000)
    }
    expect(net()).toBe(retained())
    expect(retained()).toBe(hits.reduce((sum, hit) => sum + hit.text.length * 2 + 128, 0))
  })

  it('a refused batch and its materialization agree on base text', async () => {
    // A bounded budget admits a single-cell preparation (~7.5 KiB: workspace,
    // admission framing, shared statistics session, one-cell service bound)
    // but refuses the three-cell scan batch (~9.7 KiB). Both paths must use
    // base text rather than disagreeing with a false stale.
    const { budget, net } = limitedBudget(9000)
    const doc = cfDoc(10_000)
    const source = createObjectWorkbookSource(doc, 'cf-limited')
    const owner = initialWorkbookOwner('cf-limited', source.sheets)
    const sheet = doc.sheets[0]! as unknown as XlsxSheet
    const textOf = residentTextOf(sheet)
    const prepare = residentPrepare(sheet, budget)
    // The boundary is real: a single cell prepares fine under this budget.
    const single = await prepareCfTextScope(sheet, { rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }], signal: new AbortController().signal, locale: 'en-US', budget })
    expect(single).toBeDefined()
    single!.dispose()
    expect(net()).toBe(0)
    const found = new WorkbookFind(
      { owner, sheets: residentSheets('cf-limited', source, owner) },
      { text: '1.25', caseSensitive: false },
      { textOf, prepareScope: prepare, formatBudget: budget },
    )
    try {
      expect(await drain(found)).toEqual([])
      const base = new WorkbookFind(
        { owner, sheets: residentSheets('cf-limited', source, owner) },
        { text: '1.3', caseSensitive: false },
        { textOf, prepareScope: prepare, formatBudget: budget },
      )
      try {
        // Base matches survive materialization; no FindError('stale', ...) escapes.
        const hits = await drain(base)
        expect(hits.map(hit => [hit.source.row, hit.source.col, hit.text])).toEqual([[0, 0, '1.3'], [10000, 0, '1.3']])
        // Each refused scan admits its exact-cover rects (two per search here);
        // closing both sessions returns the aggregate to baseline.
        expect(net()).toBe(4 * 64)
      } finally {
        base.close()
      }
    } finally {
      found.close()
    }
    expect(net()).toBe(0)
  })

  it('zero headroom refuses the search instead of growing an uncharged history', async () => {
    const refusing = trackingBudget(true)
    const doc = cfDoc(10_000)
    const source = createObjectWorkbookSource(doc, 'cf-zero')
    const owner = initialWorkbookOwner('cf-zero', source.sheets)
    const sheet = doc.sheets[0]! as unknown as XlsxSheet
    const denied = new WorkbookFind(
      { owner, sheets: residentSheets('cf-zero', source, owner) },
      { text: '1.25', caseSensitive: false },
      { textOf: residentTextOf(sheet), prepareScope: residentPrepare(sheet, refusing.budget), formatBudget: refusing.budget },
    )
    await expect(drain(denied)).rejects.toMatchObject({ name: 'FindError', code: 'resource-limit' })
    expect(refusing.net()).toBe(0)
    denied.close()
  })

  it('copies CF-formatted text in near and far selections', async () => {
    const { budget, net } = trackingBudget()
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
        const baseStyle = docCell(sheet, cell.row, cell.col)?.style
        return cellDisplayText(sheet, { value: cell.value, style: scope?.effectiveStyle(cell.row, cell.col, baseStyle) ?? baseStyle }, allocation)
      },
      prepareScope: (rects, signal) => residentPrepare(sheet, budget)('cf', rects, signal),
    }
    const near = await serializeSelection({ ...base, selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 1 }] }] } })
    expect(near).toMatchObject({ kind: 'ready' })
    expect(near.kind === 'ready' ? near.text : '').toBe('1.25\t0.5')
    const far = await serializeSelection({ ...base, selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 500, r1: 500 }] }, cols: [{ c0: 0, c1: 0 }] }] } })
    expect(far).toMatchObject({ kind: 'ready' })
    expect(far.kind === 'ready' ? far.text : '').toBe('1.25')
    expect(net()).toBe(0)
  })

  it('shared statistics workspace is admitted in the aggregate budget', async () => {
    let peak = 0
    let held = 0
    const budget: FormatBudget = {
      reserve: bytes => { held += bytes; peak = Math.max(peak, held) },
      release: bytes => { held -= bytes },
    }
    const cells = Array.from({ length: 5000 }, (_, row) => ({ ref: `A${row + 1}`, row, col: 0, styleIndex: 0, value: row, hasCachedValue: true }))
    const sheet = {
      name: 'CF',
      sheetId: 'cf',
      rows: cells.map(cell => ({ index: cell.row, cells: [cell] })),
      cols: [],
      merges: [],
      mergeRanges: [],
      conditionalFormatting: {
        diagnostics: [],
        rules: [{
          id: 0,
          priority: 1,
          documentOrder: 0,
          stopIfTrue: false,
          ranges: [{ firstRow: 0, lastRow: 4999, firstCol: 0, lastCol: 0 }],
          writes: new Set(['numFmt']),
          dxf: { 'numFmt.code': '0.00' },
          definition: { kind: 'top10', rank: 3, bottom: false, percent: false },
        }],
      },
    } as unknown as XlsxSheet
    // Preparing only the last cell still builds statistics over the rule scope:
    // two 40,960-byte typed arrays must be admitted before they exist.
    const scope = await prepareCfTextScope(sheet, {
      rects: [{ r0: 4999, r1: 4999, c0: 0, c1: 0 }],
      signal: new AbortController().signal,
      locale: 'en-US',
      budget,
    })
    expect(scope).toBeDefined()
    expect(peak).toBeGreaterThanOrEqual(2 * 40960)
    expect(cellDisplayText(sheet, { value: 4999, style: scope!.effectiveStyle(4999, 0, undefined) })).toBe('4999.00')
    scope!.dispose()
    expect(held).toBe(0)
  })

  it('resident and wire scopes agree on partial wins with blocked aspects', async () => {    const { budget, net } = trackingBudget()
    const doc = cfDoc()
    const sheet = doc.sheets[0]! as unknown as XlsxSheet
    // A win on numFmt with fill blocked: the surviving property still formats.
    const states = [{ row: 0, col: 0, properties: { 'numFmt.code': '0.00' }, unresolvedWrites: ['fill'] as readonly string[] }]
    const scope = await prepareCfWireScope(
      async () => ({
        result: {
          kind: 'ready' as const,
          stamp: { owner: 't', sheetId: 'cf', workbookValueGen: 0, cfGen: 7, viewGen: 0, structureGen: 0, clockGen: 0 },
          sheetId: 'cf',
          definitions: [],
          definitionsVersion: 42,
          diagnostics: [],
          states,
        },
        release: () => {},
      }),
      [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
      { budget },
    )
    expect(scope).toBeDefined()
    const base: XlsxCell['style'] = { numFmtId: 0, formatCode: '0.0' }
    const wireText = cellDisplayText(sheet, { value: 1.25, style: scope!.effectiveStyle(0, 0, base) })
    // The resident service resolves the same winning delta.
    const resident = await prepareCfTextScope(sheet, {
      rects: [{ r0: 0, r1: 0, c0: 0, c1: 0 }],
      signal: new AbortController().signal,
      locale: 'en-US',
      budget,
    })
    const residentText = cellDisplayText(sheet, { value: 1.25, style: resident!.effectiveStyle(0, 0, base) })
    expect(wireText).toBe('1.25')
    expect(residentText).toBe(wireText)
    scope!.dispose()
    resident!.dispose()
    expect(net()).toBe(0)
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
    const { budget, net } = trackingBudget()
    const bytes = pagedBytes()
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer]) }, { docId: 'cf-paged-copy', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session)
      const sheetId = source.sheets[0]!.sheetId
      const store = source.store(sheetId)
      const identity = { docId: 'cf-paged-copy', sheetId, valueGen: store.stamp.valueGen, fontGen: 0, viewGen: 0 }
      const prepareScope = async (rects: readonly SourceRect[], signal: AbortSignal) => {
        if (!store.readConditionalFormatting) return undefined
        return prepareCfWireScope(
          async chunk => {
            const result = await store.readConditionalFormatting!({
              rects: chunk,
              stamp: { owner: 'cf-paged-copy', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 },
              grantBytes: CF_TEXT_SCOPE_GRANT_BYTES,
            }, undefined, { signal })
            return { result, release: () => {} }
          },
          rects,
          { budget },
        )
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
      expect(net()).toBe(0)
    } finally {
      await session.release()
    }
  })
})
