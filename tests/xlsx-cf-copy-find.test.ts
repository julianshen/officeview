// @vitest-environment node
/**
 * Task 9: copy/find use bounded per-batch CF preparation in resident and paged
 * modes, including cells outside the current viewport. Expected strings follow
 * the stage1-numfmt-accessibility seed hypothesis (`1.25` via CF DXF `0.00`,
 * control `1.3` via cell `0.0`); the case is recipe-pending, never measured.
 * Sparse batches prepare exact-cover rects, never dense hulls, so scanning and
 * hit materialization always agree.
 */
import { describe, expect, it, vi } from 'vitest'
import { WorkbookFind, type FindSheetInput } from '../src/xlsx/view/find'
import { serializeSelection, type CopyInput } from '../src/xlsx/view/copy'
import { prepareCfTextScope, prepareCfWireScope, CF_TEXT_SCOPE_GRANT_BYTES } from '../src/xlsx/cf/text'
import { injectCfBarAutoBounds } from '../src/xlsx/cf/evidence'
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
import { BAR_SHOWVALUE_EVIDENCE, barRule, withBars } from './helpers/cf-bars'

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

  it('a refused worker batch stays on base text without a false stale', async () => {
    // A pressured provider refuses multi-cell preparations but admits single
    // cells: scan falls back to base text and records the refusal, so the
    // later single-cell materialization agrees instead of throwing stale.
    const bytes = pagedBytes()
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer]) }, { docId: 'cf-paged-refused', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session)
      const sheetId = source.sheets[0]!.sheetId
      const store = source.store(sheetId)
      const read = store.readConditionalFormatting!.bind(store)
      const probe = vi.spyOn(store, 'readConditionalFormatting').mockImplementation(async (request, memory, options) => {
        const cells = request.rects.reduce((sum, rect) => sum + (rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1), 0)
        if (cells > 1) {
          return {
            sheetId,
            stamp: { ...request.stamp },
            kind: 'unresolved' as const,
            reason: 'cf-budget-exceeded',
            states: [],
            definitions: [],
            definitionsVersion: 42,
            diagnostics: [],
          }
        }
        return read(request, memory, options)
      })
      const worker = new DocumentSession('cf-paged-refused', source)
      try {
        const found = await worker.search(worker.workbookOwner(), { text: '1.25', caseSensitive: false }, undefined, new AbortController().signal)
        expect(found).toMatchObject({ complete: true })
        if (!('hits' in found)) throw new Error('search did not complete')
        expect(found.hits).toEqual([])
        const base = await worker.search(worker.workbookOwner(), { text: '1.3', caseSensitive: false }, undefined, new AbortController().signal)
        expect(base).toMatchObject({ complete: true })
        if (!('hits' in base)) throw new Error('search did not complete')
        expect(base.hits.map(hit => [hit.source.row, hit.source.col, hit.text]).sort()).toEqual([[0, 0, '1.3'], [500, 0, '1.3']])
      } finally {
        probe.mockRestore()
        await worker.release()
      }
    } finally {
      await session.release()
    }
  })

  it('a pending worker import stays on base text without a false stale', async () => {
    // A pending import reports no states: the scan records the refusal and
    // materialization agrees on base text instead of throwing a false stale.
    const bytes = pagedBytes()
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer]) }, { docId: 'cf-paged-pending', profile: 'D' })
    try {
      await session.done
      const source = createPagedWorkbookSource(session)
      const sheetId = source.sheets[0]!.sheetId
      const store = source.store(sheetId)
      const real = store.readConditionalFormatting!.bind(store)
      const probe = vi.spyOn(store, 'readConditionalFormatting').mockImplementation(async (request, memory, options) => {
        const cells = request.rects.reduce((sum, rect) => sum + (rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1), 0)
        if (cells > 1) {
          return {
            sheetId,
            stamp: { ...request.stamp },
            kind: 'pending' as const,
            pendingSheetIds: [sheetId],
            states: [],
            definitions: [],
            definitionsVersion: 0,
            diagnostics: [],
          }
        }
        return real(request, memory, options)
      })
      const worker = new DocumentSession('cf-paged-pending', source)
      try {
        const found = await worker.search(worker.workbookOwner(), { text: '1.25', caseSensitive: false }, undefined, new AbortController().signal)
        expect(found).toMatchObject({ complete: true })
        if (!('hits' in found)) throw new Error('search did not complete')
        expect(found.hits).toEqual([])
        const base = await worker.search(worker.workbookOwner(), { text: '1.3', caseSensitive: false }, undefined, new AbortController().signal)
        expect(base).toMatchObject({ complete: true })
        if (!('hits' in base)) throw new Error('search did not complete')
        expect(base.hits.map(hit => [hit.source.row, hit.source.col, hit.text])).toEqual([[0, 0, '1.3'], [500, 0, '1.3']])
      } finally {
        probe.mockRestore()
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

/**
 * Task 22, Ruling S3-P: `showValue="0"` hides only the canvas text. Copy, find and the accessible name keep the formatted
 * value. These pass on the base too (nothing there reads `showValue`): they are characterization, each with a positive
 * control that the cells really are painted `showValue="0"` bars, and a mutant formatter that honours the canvas
 * suppression and must fail. Bars come from injected evidence only (tests/helpers/cf-bars.ts).
 */
describe('showValue=0 value still copies and finds (S3-P)', () => {
  // A1 = 1.25 and A501 = 3.75 under a '0.0' format ("1.3", "3.8"), one showValue="0" min/max bar over A1:A501.
  const barRows = (sheet: XlsxSheet): void => {
    sheet.conditionalFormatting = { diagnostics: [], rules: [barRule({ firstRow: 0, lastRow: 500, firstCol: 0, lastCol: 0 }, { showValue: false })] }
  }
  function barDoc(): XlsxDocument {
    const doc = cfDoc()
    const sheet = doc.sheets[0]! as unknown as XlsxSheet
    sheet.rows = [{ index: 0, cells: [{ ...sheet.rows[0]!.cells[0]!, value: 1.25 }] }, { index: 500, cells: [{ ...sheet.rows[1]!.cells[0]!, value: 3.75 }] }]
    barRows(sheet)
    return doc
  }
  const BAR_XML = '<conditionalFormatting sqref="A1:A501"><cfRule type="dataBar" priority="1"><dataBar showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar>'
    + '<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{3B3B0022-0000-4000-8000-000000000001}</x14:id></ext></extLst></cfRule></conditionalFormatting>'
    + '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    + '<x14:cfRule type="dataBar" id="{3B3B0022-0000-4000-8000-000000000001}"><x14:dataBar minLength="0" maxLength="100" gradient="0" direction="leftToRight"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>'
    + '<xm:sqref>A1:A501</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>'
  const barBytes = (): Uint8Array => xlsxBytes({
    sheetData: '<row r="1"><c r="A1" s="0"><v>1.25</v></c></row><row r="501"><c r="A501" s="0"><v>3.75</v></c></row>',
    styles: '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0"/></numFmts><cellXfs count="1"><xf numFmtId="164"/></cellXfs>',
    afterSheetData: BAR_XML,
  })
  /** A 200 px host at 100 %: where the canvas would paint these cells. */
  const HOST = { x: 0, y: 0, width: 200, height: 20 }, AT_100 = { zoom: 1, dpr: 1 }
  const SUPPRESS = { extra: [BAR_SHOWVALUE_EVIDENCE, 'bar:bounds-automin', 'bar:bounds-automax', 'bar:participation-dateserial'] }

  it('resident find and copy return the formatted value of a showValue="0" bar cell in near and far rows', async () => {
    await withBars(async () => {
      const { budget } = trackingBudget()
      const doc = barDoc()
      const source = createObjectWorkbookSource(doc, 'cf-bar-find')
      const owner = initialWorkbookOwner('cf-bar-find', source.sheets)
      const sheet = doc.sheets[0]! as unknown as XlsxSheet
      // Positive control: both cells are painted bars whose canvas text is hidden.
      const scope = (await prepareCfTextScope(sheet, { rects: [{ r0: 0, r1: 500, c0: 0, c1: 0 }].flatMap(r => [{ ...r, r1: 0 }, { ...r, r0: 500 }]), signal: new AbortController().signal, budget }))!
      for (const row of [0, 500]) {
        expect(scope.barAt?.(row, 0)?.showValue, `row ${row}`).toBe(false)
        expect(scope.hidesTextAt?.(row, 0, HOST, AT_100), `row ${row}`).toBe(true)
      }
      scope.dispose()
      for (const [text, want] of [['1.3', [[0, 0, '1.3']]], ['3.8', [[500, 0, '3.8']]]] as const) {
        const session = new WorkbookFind({ owner, sheets: residentSheets('cf-bar-find', source, owner) }, { text, caseSensitive: false },
          { textOf: residentTextOf(sheet), prepareScope: residentPrepare(sheet, budget), formatBudget: budget })
        expect((await drain(session)).map(hit => [hit.source.row, hit.source.col, hit.text])).toEqual(want)
      }
      const store = source.store('cf'), entry = owner.sheetVector[0]!
      const format: CopyInput['format'] = (cell, allocation, scope) => {
        const baseStyle = docCell(sheet, cell.row, cell.col)?.style
        return cellDisplayText(sheet, { value: cell.value, style: scope?.effectiveStyle(cell.row, cell.col, baseStyle) ?? baseStyle }, allocation)
      }
      const input = (r0: number, fmt: CopyInput['format'] = format): CopyInput => ({
        store, owner: { ...store.stamp, fontGen: 0, viewGen: entry.revision }, currentOwner: () => ({ ...store.stamp, fontGen: 0, viewGen: entry.revision }),
        selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0, r1: r0 }] }, cols: [{ c0: 0, c1: 0 }] }] },
        order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])), columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
        allowed: () => true, signal: new AbortController().signal, format: fmt,
        prepareScope: (rects, signal) => residentPrepare(sheet, budget)('cf', rects, signal),
      })
      for (const [row, want] of [[0, '1.3'], [500, '3.8']] as const) {
        const copied = await serializeSelection(input(row))
        expect(copied.kind === 'ready' ? copied.text : copied.kind, `row ${row}`).toBe(want)
        // Mutant: a formatter that applies the canvas suppression to copy fails this test.
        const mutant = await serializeSelection(input(row, (cell, allocation, scope) => scope?.hidesTextAt?.(cell.row, cell.col, HOST, AT_100) ? '' : format(cell, allocation, scope)))
        expect(mutant.kind === 'ready' ? mutant.text : mutant.kind, `mutant row ${row}`).not.toBe(want)
      }
    }, SUPPRESS)
  })

  it('worker find and main-thread paged copy return the formatted value of a showValue="0" bar cell in near and far rows', async () => {
    // The fixture is Excel's preset shape (x14 autoMin/autoMax): its bounds resolve only through a verified rule
    // (Task 22c, S3-AUTO-BOUNDS), stubbed here beside the stubbed capabilities.
    injectCfBarAutoBounds({ autoMin: { rule: 'zero-anchored', settledBy: ['bar:bounds-automin'] }, autoMax: { rule: 'zero-anchored', settledBy: ['bar:bounds-automax'] } })
    try { await withBars(async () => {
      const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(barBytes()).buffer]) }, { docId: 'cf-bar-paged', profile: 'D' })
      try {
        await session.done
        const source = createPagedWorkbookSource(session)
        const sheetId = source.sheets[0]!.sheetId
        const store = source.store(sheetId)
        const read = async (rects: readonly SourceRect[], signal: AbortSignal) => store.readConditionalFormatting!({
          rects, stamp: { owner: 'cf-bar-paged', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: CF_TEXT_SCOPE_GRANT_BYTES,
        }, undefined, { signal })
        // Positive control: the worker's wire states carry the showValue="0" bar on both rows.
        const wire = await read([{ r0: 0, r1: 0, c0: 0, c1: 0 }, { r0: 500, r1: 500, c0: 0, c1: 0 }], new AbortController().signal)
        expect(wire?.kind).toBe('ready')
        expect(wire!.states.filter(s => s.bar?.showValue === false).map(s => s.row).sort((a, b) => a - b)).toEqual([0, 500])
        const { budget } = trackingBudget()
        const prepareScope = (rects: readonly SourceRect[], signal: AbortSignal) =>
          prepareCfWireScope(async chunk => ({ result: await read(chunk, signal), release: () => {} }), rects, { budget })
        const format: CopyInput['format'] = (cell, allocation, scope) => source.displayText!(cell, sheetId, allocation, undefined, scope)
        const identity = { docId: 'cf-bar-paged', sheetId, valueGen: store.stamp.valueGen, fontGen: 0, viewGen: 0 }
        for (const [row, want] of [[0, '1.3'], [500, '3.8']] as const) {
          const input: CopyInput = {
            store, owner: identity, currentOwner: () => identity,
            selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: row, r1: row }] }, cols: [{ c0: 0, c1: 0 }] }] },
            order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])), columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
            allowed: () => true, signal: new AbortController().signal, format, prepareScope,
          }
          const copied = await serializeSelection(input)
          expect(copied.kind === 'ready' ? copied.text : JSON.stringify(copied), `row ${row}`).toBe(want)
          const scope = (await prepareScope([{ r0: row, r1: row, c0: 0, c1: 0 }], new AbortController().signal))!
          expect(scope.hidesTextAt?.(row, 0, HOST, AT_100), `paged scope hides the canvas text at row ${row}`).toBe(true)
          scope.dispose()
          const mutant = await serializeSelection({ ...input, format: (cell, allocation, s) => s?.hidesTextAt?.(cell.row, cell.col, HOST, AT_100) ? '' : format(cell, allocation, s) })
          expect(mutant.kind === 'ready' ? mutant.text : mutant.kind, `mutant row ${row}`).not.toBe(want)
        }
        // Last: the worker session releases the source it searches.
        const worker = new DocumentSession('cf-bar-paged', source)
        try {
          for (const [text, want] of [['1.3', [[0, 0, '1.3']]], ['3.8', [[500, 0, '3.8']]]] as const) {
            const found = await worker.search(worker.workbookOwner(), { text, caseSensitive: false }, undefined, new AbortController().signal)
            if (!('hits' in found)) throw new Error('search did not complete')
            expect(found.hits.map(hit => [hit.source.row, hit.source.col, hit.text])).toEqual(want)
          }
        } finally { await worker.release() }
      } finally { await session.release() }
    }, SUPPRESS) } finally { injectCfBarAutoBounds(null) }
  })
})
