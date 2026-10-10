import { describe, expect, expectTypeOf, it } from 'vitest'
import * as contractConstants from '../src/xlsx/view/contracts'
import {
  MAX_ROWS, MAX_COLS, MAX_CELL_ID, CoordinateError,
  validateCoordinate, isValidCoordinate, cellId, cellIdToCoordinate,
  parseA1, formatA1, columnToLetters, lettersToColumn,
  validateSourceRect, isValidSourceRect, sourceRectCellCount,
} from '../src/xlsx/view/coordinates'
import {
  MAX_OVERVIEW_GROUPS_DESKTOP, MAX_OVERVIEW_GROUPS_MOBILE, MAX_PHYSICAL_SCROLL_CSS_PX,
} from '../src/xlsx/view/contracts'
import type {
  CellRead, SourceStore, Page, SourceStamp, OverviewDomain, OverviewGroup,
  CompactSourceRowSet, WorkerCommand, WorkerReply, CacheIdentity, WorkbookOwner,
  ValuePolicy, OperationRefusal, SourceErrorCode, ReadOptions, SourceRowInterval,
  RowOrder, RowPermutationBand, RowVisibility, AxisMetrics, SheetMetrics, ScrollWindow,
  StructuralResult, SheetView, WorkerResult, SourceFailure, AuthoritativeCell, SparseRowPage,
  ColumnEntry, ViewCommitDelta, PagedKeySetRef, WorkerReplyBase,
} from '../src/xlsx/view/contracts'

function expectCode(action: () => unknown, code: CoordinateError['code']) {
  expect(action).toThrow(CoordinateError)
  try {
    action()
  } catch (error) {
    expect(error).toMatchObject({ code })
  }
}

describe('coordinate contracts (spec §3, AC-CAP-01)', () => {
  it('declares the full exact safe-number domain', () => {
    expect(MAX_ROWS).toBe(1_048_576)
    expect(MAX_COLS).toBe(16_384)
    expect(MAX_CELL_ID).toBe(17_179_869_183)
    expect(Number.isSafeInteger(MAX_CELL_ID)).toBe(true)
  })

  it('preserves the plan boundary assertions verbatim', () => {
    expect(cellId(1_048_575, 16_383)).toBe(17_179_869_183)
    expect(cellId(262_144, 0)).not.toBe(cellId(0, 0))
    expect(() => validateCoordinate(1_048_576, 0)).toThrow()
  })

  it('never aliases at signed or unsigned 32-bit boundaries', () => {
    expect(cellId(131_072, 0)).toBe(2_147_483_648)
    expect(cellId(262_144, 0)).toBe(4_294_967_296)
  })

  for (const row of [0, 16_384, 16_385, 131_072, 262_144, 1_048_575]) {
    for (const col of [0, 4_095, 4_096, 16_383]) {
      it(`round-trips source row ${row}, column ${col}`, () => {
        expect(validateCoordinate(row, col)).toBeUndefined()
        expect(isValidCoordinate(row, col)).toBe(true)
        expect(cellId(row, col)).toBe(row * 16_384 + col)
        expect(cellIdToCoordinate(cellId(row, col))).toEqual({ row, col })
        expect(parseA1(formatA1({ row, col }))).toEqual({ row, col })
      })
    }
  }

  it.each([-1, 0.5, NaN, Infinity, -Infinity, Number.MAX_SAFE_INTEGER, 1_048_576])(
    'rejects invalid row %s with a stable code', (row) => {
      expectCode(() => validateCoordinate(row, 0), 'invalid-coordinate')
      expectCode(() => cellId(row, 0), 'invalid-coordinate')
      expectCode(() => formatA1({ row, col: 0 }), 'invalid-coordinate')
      expect(isValidCoordinate(row, 0)).toBe(false)
    },
  )
  it.each([-1, 0.5, NaN, Infinity, -Infinity, Number.MAX_SAFE_INTEGER, 16_384])(
    'rejects invalid column %s with a stable code', (col) => {
      expectCode(() => validateCoordinate(0, col), 'invalid-coordinate')
      expectCode(() => cellId(0, col), 'invalid-coordinate')
      expectCode(() => columnToLetters(col), 'invalid-coordinate')
      expect(isValidCoordinate(0, col)).toBe(false)
    },
  )
  it.each([-1, 0.5, NaN, Infinity, -Infinity, 17_179_869_184, Number.MAX_SAFE_INTEGER])(
    'rejects invalid cell ID %s', (id) => {
      expectCode(() => cellIdToCoordinate(id), 'invalid-cell-id')
    },
  )

  it.each([
    ['A1', 0, 0], ['XFD1048576', 1_048_575, 16_383], ['A16386', 16_385, 0],
    ['xfd1048576', 1_048_575, 16_383], ['$A$1', 0, 0], ['$A1', 0, 0], ['A$1', 0, 0],
  ] as const)('parses and normalizes %s', (address, row, col) => {
    expect(parseA1(address)).toEqual({ row, col })
    expect(formatA1({ row, col })).toBe(address.replace(/\$/g, '').toUpperCase())
  })
  it.each(['XFE1', 'A1048577', 'A0', 'A01', '', ' A1', 'A1 ', '1A', 'A-1', 'A1.5', 'A1:B2', 'Sheet1!A1', '$$A1', 'A$$1', 'A1\n', 'A999999999999999999999999999'])(
    'rejects malformed or out-of-domain A1 address %s', (address) => {
      expectCode(() => parseA1(address), 'invalid-a1')
    },
  )
  it.each([['A', 0], ['Z', 25], ['AA', 26], ['AZ', 51], ['BA', 52], ['XFD', 16_383]] as const)(
    'round-trips column letters %s', (letters, col) => {
      expect(columnToLetters(col)).toBe(letters)
      expect(lettersToColumn(letters)).toBe(col)
      expect(lettersToColumn(letters.toLowerCase())).toBe(col)
    },
  )
  it.each(['', 'XFE', '$A', 'A1', ' A', 'A\n', 'AAAAAAAAAAAAAAAAAAAA'])('rejects invalid column letters %s', (letters) => {
    expectCode(() => lettersToColumn(letters), 'invalid-a1')
  })

  it('round-trips every legal column label without rollover aliases', () => {
    const labels = new Set<string>()
    for (let col = 0; col < MAX_COLS; col++) {
      const label = columnToLetters(col)
      expect(lettersToColumn(label)).toBe(col)
      expect(parseA1(`${label}1048576`)).toEqual({ row: MAX_ROWS - 1, col })
      labels.add(label)
    }
    expect(labels.size).toBe(MAX_COLS)
  })

  it('validates inclusive rectangles and exact full-sheet counts', () => {
    const full = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
    expect(validateSourceRect(full)).toBeUndefined()
    expect(isValidSourceRect(full)).toBe(true)
    expect(sourceRectCellCount(full)).toBe(17_179_869_184)
    expect(Number.isSafeInteger(sourceRectCellCount(full))).toBe(true)
    expect(sourceRectCellCount({ r0: 9, r1: 9, c0: 2, c1: 2 })).toBe(1)
    expect(sourceRectCellCount({ r0: 9, r1: 11, c0: 2, c1: 4 })).toBe(9)
  })
  it.each([
    { r0: 1, r1: 0, c0: 0, c1: 0 }, { r0: 0, r1: 0, c0: 1, c1: 0 },
    ...[-1, 0.5, NaN, Infinity, -Infinity, Number.MAX_SAFE_INTEGER].flatMap((invalid) =>
      ['r0', 'r1', 'c0', 'c1'].map((key) => ({ r0: 0, r1: 0, c0: 0, c1: 0, [key]: invalid }))),
    { r0: 0, r1: MAX_ROWS, c0: 0, c1: 0 }, { r0: 0, r1: 0, c0: 0, c1: MAX_COLS },
  ])('rejects invalid rectangle %o', (rect) => {
    expectCode(() => validateSourceRect(rect), 'invalid-source-rect')
    expectCode(() => sourceRectCellCount(rect), 'invalid-source-rect')
    expect(isValidSourceRect(rect)).toBe(false)
  })

  it('rejects negative zero rather than publishing a signed-zero identity', () => {
    expect(Object.is(-0, 0)).toBe(false)
    expect(isValidCoordinate(-0, 0)).toBe(false)
    expect(isValidCoordinate(0, -0)).toBe(false)
    expectCode(() => validateCoordinate(-0, -0), 'invalid-coordinate')
    expectCode(() => cellId(-0, 0), 'invalid-coordinate')
    expectCode(() => cellIdToCoordinate(-0), 'invalid-cell-id')
    expect(Object.is(cellIdToCoordinate(0).row, -0)).toBe(false)
    expect(Object.is(cellIdToCoordinate(0).col, -0)).toBe(false)
  })

  it.each(['A\n', 'A1\n', 'XFD\n', 'A\t'])(
    'rejects trailing whitespace without multiline regex guards: %j', (text) => {
      expect(/^[A-Za-z]{1,3}$/.exec(text)).toBeNull()
      expect(/^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/.exec(text)).toBeNull()
      expectCode(() => parseA1(text), 'invalid-a1')
      expectCode(() => lettersToColumn(text), 'invalid-a1')
    },
  )

  it.each([undefined, null, '1', {}, [], Symbol('coordinate'), 1n])(
    'rejects non-number IDs with the existing typed error: %s', (input) => {
      // The cast exercises calls from untyped consumers at the API boundary.
      expectCode(() => cellIdToCoordinate(input as number), 'invalid-cell-id')
    },
  )
  it.each([undefined, null, 'A1', 1, [], Symbol('rectangle')])(
    'returns false for non-rectangles and validates them with a typed error: %s', (input) => {
      const rect = input as unknown as Parameters<typeof validateSourceRect>[0]
      expect(isValidSourceRect(rect)).toBe(false)
      expectCode(() => validateSourceRect(rect), 'invalid-source-rect')
      expectCode(() => sourceRectCellCount(rect), 'invalid-source-rect')
    },
  )

  it.todo('S0.2 full-extent sparse provider: expect(await source.readCell({ row: 1_048_575, col: 16_383 }, readOptions)).toMatchObject({ kind: \'cell\' })')
})

describe('type contract freeze (spec §§4–5, 10)', () => {
  const stamp = { docId: 'doc-v1', sheetId: 'sheet-1', valueGen: 0 } satisfies SourceStamp
  const emptyPage = <T>(): Page<T> => ({ items: [], completeness: 'complete', stamp })
  // Only a compile-time fake; adapter behavior and the edge-cell fixture belong to S0.2/S1.
  const source = {
    stamp,
    valuePolicy: 'engine',
    extent: { rows: 1_048_576, cols: 16_384, used: { r0: 0, r1: 0, c0: 0, c1: 0 } },
    async readCell() { return { kind: 'unknown', pageId: 'page-1' } as const },
    async readRows() { return emptyPage() },
    async queryColumn() { return emptyPage() },
    async *scanPopulated() { yield emptyPage() },
    async queryStructural() {
      return { stamp, completeness: 'complete' as const, merges: [], tables: [], spills: [], drawings: [], missing: [] }
    },
    async queryOverview() { return emptyPage() },
    async ensureComplete(range) { return { range, stamp, completeness: 'complete' as const, missing: [] } },
    async release() {},
  } satisfies SourceStore

  it('requires all asynchronous source methods and literal full extent', () => {
    expectTypeOf(source).toExtend<SourceStore>()
    expectTypeOf<SourceStore['extent']['rows']>().toEqualTypeOf<1048576>()
    expectTypeOf<SourceStore['extent']['cols']>().toEqualTypeOf<16384>()
    expectTypeOf<ReturnType<SourceStore['readCell']>>().toEqualTypeOf<Promise<CellRead>>()
    expectTypeOf<Parameters<SourceStore['readCell']>>().toEqualTypeOf<[
      at: { row: number; col: number }, options: ReadOptions,
    ]>()
    expectTypeOf<Parameters<SourceStore['readRows']>>().toEqualTypeOf<[
      range: { r0: number; r1: number; c0: number; c1: number },
      cursor: string | undefined, options: ReadOptions,
    ]>()
    expectTypeOf<ReturnType<SourceStore['queryStructural']>>().toEqualTypeOf<Promise<StructuralResult>>()
    expectTypeOf<ReturnType<SourceStore['release']>>().toEqualTypeOf<Promise<void>>()
    expect(source.extent.used).toEqual({ r0: 0, r1: 0, c0: 0, c1: 0 })
    // @ts-expect-error Source extents cannot be truncated to the used range.
    const truncated: SourceStore['extent'] = { rows: 1, cols: 1 }
    expect(truncated).toBeDefined()
  })

  function classify(read: CellRead): string {
    switch (read.kind) {
      case 'cell': return `cell:${read.cell.id}`
      case 'blank': return 'blank'
      case 'unknown': return `pending:${read.pageId}`
      case 'error': return `error:${read.code}`
      case 'unavailable': return read.reason
      default: {
        const exhaustive: never = read
        return exhaustive
      }
    }
  }

  it('keeps cell, blank, unknown, error and unavailable exhaustive and distinct', () => {
    const cell = { id: 0, row: 0, col: 0, value: null, formula: '1+1', hasCachedValue: false, styleIndex: 0 }
    const reads = [
      { kind: 'cell', cell }, { kind: 'blank' }, { kind: 'unknown', pageId: 'page-1' },
      { kind: 'error', code: 'page-error', retryable: true },
      { kind: 'unavailable', cell, reason: 'formula-value-unavailable' },
    ] satisfies CellRead[]
    expect(reads.map(classify)).toEqual(['cell:0', 'blank', 'pending:page-1', 'error:page-error', 'formula-value-unavailable'])
    expectTypeOf<CellRead['kind']>().toEqualTypeOf<'cell' | 'blank' | 'unknown' | 'error' | 'unavailable'>()
    expectTypeOf<ValuePolicy>().toEqualTypeOf<'engine' | 'cached'>()
    const refusal = { reason: 'formula-value-unavailable', count: 1, firstAddress: { row: 0, col: 0 } } satisfies OperationRefusal
    expect(refusal.count).toBe(1)
  })

  it('bounds overview/scroll output while preserving exact disjoint source membership', () => {
    expect(MAX_OVERVIEW_GROUPS_DESKTOP).toBe(2048)
    expect(MAX_OVERVIEW_GROUPS_MOBILE).toBe(1024)
    expect(MAX_PHYSICAL_SCROLL_CSS_PX).toBe(1_000_000)
    const memberships = [
      { kind: 'intervals', intervals: [{ r0: 1, r1: 1 }, { r0: 3, r1: 3 }] },
      { kind: 'list', r0: 1, r1: 3, rows: new Uint32Array([1, 3]) },
      { kind: 'bitset', r0: 1, r1: 3, words: new Uint32Array([5]), count: 2 },
      { kind: 'ref', refId: 'rows-v1', count: 2, stamp },
    ] satisfies CompactSourceRowSet[]
    const domain = { groupId: 'group-v1', rows: memberships[0], cols: [{ c0: 0, c1: 0 }] } satisfies OverviewDomain
    const group = { groupId: domain.groupId, populatedCount: 2, countKind: 'exact', style: { kind: 'mixed' }, content: 'populated' } satisfies OverviewGroup
    expect(group.populatedCount).toBe(2)
    expectTypeOf<Extract<keyof OverviewGroup, 'x' | 'y' | 'width' | 'height' | 'text'>>().toEqualTypeOf<never>()
    // @ts-expect-error An enclosing rectangle is not exact disjoint row membership.
    const boundingRows: CompactSourceRowSet = { r0: 1, r1: 3 }
    expect(boundingRows).toBeDefined()
  })

  it('freezes compact row state, full-axis geometry and atomic view shapes', () => {
    expectTypeOf<RowOrder>().toEqualTypeOf<{
      kind: 'identity' | 'bands'
      bands: readonly RowPermutationBand[]
      displayToSource(displayRow: number): number
      sourceToDisplay(sourceRow: number): number
    }>()
    expectTypeOf<RowPermutationBand>().toEqualTypeOf<{
      r0: number; r1: number
      sourceAtOffset: Uint32Array; displayOffsetOfSource: Uint32Array
    }>()
    expectTypeOf<RowVisibility>().toEqualTypeOf<{
      visible(displayRow: number): boolean
      rank(endExclusive: number): number
      select(visibleOrdinal: number): number | null
      next(displayRow: number, direction: -1 | 1): number | null
    }>()
    expectTypeOf<ReturnType<AxisMetrics['locate']>>().toEqualTypeOf<number | null>()
    expectTypeOf<ReturnType<AxisMetrics['visibleBetween']>['runs']>().toExtend<Iterable<unknown>>()
    expectTypeOf<SheetMetrics['order']>().toEqualTypeOf<RowOrder>()
    expectTypeOf<ReturnType<SheetMetrics['sourceCellToLogical']>>().toEqualTypeOf<{
      x: number; y: number; width: number; height: number
    } | null>()
    expectTypeOf<SheetView['baseHidden']['intervals']>().toEqualTypeOf<readonly SourceRowInterval[]>()
    expectTypeOf<SheetView['selection']['active']>().toEqualTypeOf<{ row: number; col: number } | null>()
    expectTypeOf<ScrollWindow>().toEqualTypeOf<{
      logicalOrigin: { x: number; y: number }
      physicalPosition: { x: number; y: number }
      physicalExtent: { width: number; height: number }
      zoom: number; deviceScale: number
    }>()
    expectTypeOf<Extract<WorkerResult, { kind: 'view' }>['delta']['owner']>().toEqualTypeOf<CacheIdentity>()
    expectTypeOf<Extract<WorkerResult, { kind: 'view' }>['delta']['workbookOwner']>().toEqualTypeOf<WorkbookOwner>()
  })

  it('requires generation-complete owners on completed sheet and workbook publications', () => {
    type DocumentOnlySheetReply = {
      kind: 'complete'; commandKind: 'viewport'; requestId: string; owner: { docId: string }; diagnostics: []
      result: Extract<WorkerResult, { kind: 'viewport' }>
    }
    type LocalOnlySearchReply = {
      kind: 'complete'; commandKind: 'search'; requestId: string; owner: CacheIdentity; diagnostics: []
      result: Extract<WorkerResult, { kind: 'search' }>
    }
    expectTypeOf<DocumentOnlySheetReply extends WorkerReply ? true : false>().toEqualTypeOf<false>()
    expectTypeOf<LocalOnlySearchReply extends WorkerReply ? true : false>().toEqualTypeOf<false>()
  })

  it('carries typed failures and independently unloaded overview groups in every source page', () => {
    const failure = { code: 'quota', retryable: true, message: 'Storage full' } satisfies SourceFailure
    const groups = [
      { groupId: 'known', content: 'blank', populatedCount: 0, countKind: 'exact', style: { kind: 'uniform', styleIndex: 0 } },
      { groupId: 'pending', content: 'unknown', pageId: 'page-2' },
    ] satisfies OverviewGroup[]
    const page = { items: groups, completeness: 'incomplete', stamp, error: failure } satisfies Page<OverviewGroup>
    function summarize(group: OverviewGroup): string {
      switch (group.content) {
        case 'unknown': return `loading:${group.pageId}`
        case 'blank': return 'blank'
        case 'populated': return `populated:${group.populatedCount}`
        case 'mixed': return 'mixed'
        default: { const exhaustive: never = group; return exhaustive }
      }
    }
    expect(page.items.map(summarize)).toEqual(['blank', 'loading:page-2'])
    expect(page.error.code).toBe('quota')
    // Unknown aggregates must not invent counts/styles before their page loads.
    expectTypeOf<Extract<OverviewGroup, { content: 'unknown' }>>().toEqualTypeOf<{
      groupId: string; content: 'unknown'; pageId: string
    }>()
    const failures = [
      { items: [], completeness: 'incomplete', stamp, error: failure } satisfies Page<SparseRowPage>,
      { items: [], completeness: 'incomplete', stamp, error: failure } satisfies Page<ColumnEntry>,
      { items: [], completeness: 'incomplete', stamp, error: failure } satisfies Page<AuthoritativeCell>,
    ]
    expect(failures.every(result => result.error === failure)).toBe(true)
  })

  it('keeps availability explicit in sparse rows/scans and fixes document value policy', () => {
    const engineCell = { id: 0, row: 0, col: 0, value: 2, formula: '1+1', hasCachedValue: false, styleIndex: 0 } satisfies AuthoritativeCell
    const cachedCell = { ...engineCell, value: null, unavailable: true, valueProvenance: 'cached' } satisfies AuthoritativeCell
    const row = { row: 0, cells: [engineCell, cachedCell] } satisfies SparseRowPage
    const scan = { items: [cachedCell], completeness: 'complete', stamp } satisfies Page<AuthoritativeCell>
    expect(row.cells.map(cell => 'unavailable' in cell ? cell.unavailable : undefined)).toEqual([undefined, true])
    expect(scan.items[0].unavailable).toBe(true)
    expect(source.valuePolicy).toBe('engine')
    expectTypeOf<SourceStore['valuePolicy']>().toEqualTypeOf<ValuePolicy>()
    expectTypeOf<AuthoritativeCell['unavailable']>().toEqualTypeOf<true | undefined>()
    expectTypeOf<Extract<keyof AuthoritativeCell, 'style'>>().toEqualTypeOf<never>()
    expectTypeOf<Extract<keyof SparseRowPage, 'range' | 'limit' | 'cursor'>>().toEqualTypeOf<never>()
  })

  it('pins source error codes at runtime and treats source code and reply kind as orthogonal', () => {
    const codes = [
      'invalid-coordinate', 'invalid-source-rect', 'invalid-cursor', 'stale-source',
      'cancelled', 'incomplete', 'quota', 'resource-limit', 'unsupported',
      'parse-error', 'storage-error', 'page-error',
    ] satisfies SourceErrorCode[]
    expect(Reflect.get(contractConstants, 'SOURCE_ERROR_CODES')).toEqual(codes)
    function classify(code: SourceErrorCode): string {
      switch (code) {
        case 'invalid-coordinate': case 'invalid-source-rect': case 'invalid-cursor': case 'stale-source':
        case 'cancelled': case 'incomplete': case 'quota': case 'resource-limit': case 'unsupported':
        case 'parse-error': case 'storage-error': case 'page-error': return code
        default: { const exhaustive: never = code; return exhaustive }
      }
    }
    expect(codes.map(classify)).toEqual(codes)
    const owner = { ...stamp, fontGen: 0, viewGen: 0 }
    const reply = { commandKind: 'viewport', kind: 'error', requestId: 'r1', owner, diagnostics: [], code: 'quota', retryable: true } satisfies WorkerReply
    expect(reply.kind).toBe('error')
    expect(reply.code).toBe('quota')
    expectTypeOf<Extract<CompactSourceRowSet, { kind: 'ref' }>['stamp']>().toEqualTypeOf<PagedKeySetRef['owner']>()
  })

  it('includes absolute sheet zoom and DPR in viewport ownership and progressive detail/coarse replies', () => {
    expect(Reflect.get(contractConstants, 'SHEET_MIN_ZOOM')).toBe(0.25)
    expect(Reflect.get(contractConstants, 'SHEET_MAX_ZOOM')).toBe(4)
    const bounds = { x: 0, y: 0, width: 40, height: 20 }
    const results = [
      { kind: 'viewport', completeness: 'loading', content: { kind: 'detail', cells: [{ source: { row: 0, col: 0 }, bounds, kind: 'unknown', pageId: 'p1' }] } },
      { kind: 'viewport', completeness: 'loading', content: { kind: 'overview', groups: [{ groupId: 'g1', bounds, content: 'unknown', pageId: 'p2' }] } },
      { kind: 'viewport', completeness: 'loading', content: { kind: 'unknown', pageId: 'p3' } },
      { kind: 'viewport', completeness: 'incomplete', content: { kind: 'unknown', pageId: 'p4' }, error: { code: 'page-error', retryable: true, message: 'Unavailable page' } },
    ] satisfies Extract<WorkerResult, { kind: 'viewport' }>[]
    function mode(result: Extract<WorkerResult, { kind: 'viewport' }>): string {
      const content = result.content
      switch (content.kind) {
        case 'detail': return content.cells[0].kind
        case 'overview': return content.groups[0].content
        case 'unknown': return content.kind
        default: { const exhaustive: never = content; return exhaustive }
      }
    }
    expect(results.map(mode)).toEqual(['unknown', 'unknown', 'unknown', 'unknown'])
    const command = { kind: 'viewport', requestId: 'r1', owner: { ...stamp, viewGen: 0, fontGen: 0 }, window: { box: bounds, zoom: 0.25, deviceScale: 2 } } satisfies WorkerCommand
    expect(command.window.zoom * command.window.deviceScale).toBe(0.5)
  })

  it('correlates progress/failure/refusal ownership with the command kind', () => {
    type BadProgress = { kind: 'progress'; commandKind: 'sort'; requestId: string; diagnostics: []; owner: { docId: string }; completed: number }
    type BadSearchFailure = { kind: 'error'; commandKind: 'search'; requestId: string; diagnostics: []; owner: CacheIdentity; code: 'page-error'; retryable: true }
    type BadResultKind = { kind: 'complete'; commandKind: 'sort'; requestId: string; diagnostics: []; owner: CacheIdentity; result: Extract<WorkerResult, { kind: 'search' }> }
    expectTypeOf<BadProgress extends WorkerReply ? true : false>().toEqualTypeOf<false>()
    expectTypeOf<BadSearchFailure extends WorkerReply ? true : false>().toEqualTypeOf<false>()
    expectTypeOf<BadResultKind extends WorkerReply ? true : false>().toEqualTypeOf<false>()
    type WorkbookOwnedSort = Omit<BadProgress, 'owner'> & { owner: WorkbookOwner }
    expectTypeOf<WorkbookOwnedSort extends WorkerReplyBase<'sort' | 'search'> ? true : false>().toEqualTypeOf<false>()
    const owner = { ...stamp, viewGen: 0, fontGen: 0 }
    const refusal = { reason: 'merge', range: { r0: 0, r1: 1, c0: 0, c1: 1 } } satisfies OperationRefusal
    const reply = { kind: 'error', commandKind: 'sort', requestId: 'r1', owner, diagnostics: [], code: 'operation-refused', retryable: false, refusal } satisfies WorkerReply
    expect(reply.owner).toEqual(owner)
    expect(reply.refusal).toEqual(refusal)
    expect('missing' in reply).toBe(false)
    type BogusRefusal = Omit<typeof reply, 'code'> & { code: 'page-error' }
    expectTypeOf<BogusRefusal extends WorkerReply ? true : false>().toEqualTypeOf<false>()
  })

  it('pins ascending absolute lists and LSB-first bounded masks to the same exact members', () => {
    type EncodedRows = Exclude<CompactSourceRowSet, { kind: 'ref' }>
    function decode(set: EncodedRows): number[] {
      switch (set.kind) {
        case 'intervals': {
          const rows: number[] = []
          let previousEnd = -1
          for (const interval of set.intervals) {
            expect(interval.r0).toBeGreaterThan(previousEnd)
            expect(interval.r1).toBeGreaterThanOrEqual(interval.r0)
            for (let row = interval.r0; row <= interval.r1; row++) rows.push(row)
            previousEnd = interval.r1
          }
          return rows
        }
        case 'list': {
          const rows = Array.from(set.rows)
          rows.forEach((row, index) => {
            expect(row).toBeGreaterThanOrEqual(set.r0)
            expect(row).toBeLessThanOrEqual(set.r1)
            if (index > 0) expect(row).toBeGreaterThan(rows[index - 1])
          })
          return rows
        }
        case 'bitset': {
          const length = set.r1 - set.r0 + 1
          expect(set.words.length).toBe(Math.ceil(length / 32))
          const rows: number[] = []
          for (let i = 0; i < set.words.length * 32; i++) {
            // Bitwise operations encode one mask word, never a cell identity.
            const bit = (set.words[Math.floor(i / 32)] >>> (i % 32)) & 1
            if (i >= length) expect(bit).toBe(0)
            else if (bit) rows.push(set.r0 + i)
          }
          expect(rows.length).toBe(set.count)
          return rows
        }
        default: { const exhaustive: never = set; return exhaustive }
      }
    }
    for (const rows of [[1, 3], [0, 1, 3], [100, 101, 131, 132, 133]]) {
      const r0 = rows[0], r1 = rows[rows.length - 1]
      const words = rows.length === 5 ? new Uint32Array([0x80000003, 3]) : new Uint32Array([rows[0] === 0 ? 11 : 5])
      const list = { kind: 'list', r0, r1, rows: new Uint32Array(rows) } satisfies CompactSourceRowSet
      const bitset = { kind: 'bitset', r0, r1, words, count: rows.length } satisfies CompactSourceRowSet
      expect(decode(list)).toEqual(decode(bitset))
      expect(decode(bitset)).toEqual(rows)
      expect(decode(bitset)).not.toContain(2)
      const visibility = [bitset] satisfies ViewCommitDelta['visibility']
      expect(visibility[0].count).toBe(rows.length)
    }
    const intervals = { kind: 'intervals', intervals: [{ r0: 0, r1: 1 }, { r0: 3, r1: 3 }] } satisfies CompactSourceRowSet
    expect(decode(intervals)).toEqual([0, 1, 3])
    expectTypeOf<ViewCommitDelta['visibility'][number]>().toEqualTypeOf<Extract<CompactSourceRowSet, { kind: 'bitset' }>>()
  })

  it('pins forward absolute source rows and inverse band-relative display offsets', () => {
    for (const r0 of [0, 10]) {
      const band = { r0, r1: r0 + 3, sourceAtOffset: new Uint32Array([r0, r0 + 3, r0 + 1, r0 + 2]), displayOffsetOfSource: new Uint32Array([0, 2, 3, 1]) } satisfies RowPermutationBand
      band.sourceAtOffset.forEach((sourceRow, displayOffset) => {
        expect(band.displayOffsetOfSource[sourceRow - band.r0]).toBe(displayOffset)
      })
      expect(Array.from(band.sourceAtOffset.slice(1, 3)).map(row => row - r0).sort()).toEqual([1, 3])
      expect(Array.from(band.sourceAtOffset.slice(0, 3)).map(row => row - r0).sort()).toEqual([0, 1, 3])
    }
  })

  it('provides DTO v1 command/reply examples with complete generation owners', () => {
    const owner = { ...stamp, fontGen: 0, viewGen: 1 } satisfies CacheIdentity
    const workbook = {
      docId: stamp.docId, globalViewGen: 1, workbookValueGen: 0, fontGen: 0,
      sheetVector: [{ sheetId: stamp.sheetId, revision: 1, valueGen: 0 }],
    } satisfies WorkbookOwner
    const commands = [
      { kind: 'open', requestId: 'r0', docId: stamp.docId, source: { kind: 'blob', blob: new Blob() }, profile: 'D' },
      { kind: 'activateSheet', requestId: 'r1', owner: workbook, sheetId: stamp.sheetId },
      { kind: 'prepareColumns', requestId: 'r2', owner, cols: new Uint16Array([16_383]) },
      { kind: 'sort', requestId: 'r3', owner, spec: null },
      { kind: 'filter', requestId: 'r4', owner, spec: null, nowMs: 0 },
      { kind: 'viewport', requestId: 'r5', owner, window: { box: { x: 0, y: 0, width: 800, height: 600 }, zoom: 0.25, deviceScale: 2 } },
      { kind: 'valueQuery', requestId: 'r6', owner, query: { col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 10 }] }, limit: 200 } },
      { kind: 'search', requestId: 'r7', owner: workbook, query: { text: 'value', caseSensitive: false } },
      { kind: 'cancel', requestId: 'r8', targetRequestId: 'r7' },
      { kind: 'close', requestId: 'r9', docId: stamp.docId },
    ] satisfies WorkerCommand[]
    const base = { commandKind: 'sort' as const, requestId: 'r3', owner, diagnostics: [] }
    const replies = [
      { ...base, kind: 'progress', completed: 5, total: 10 },
      { ...base, kind: 'complete', result: { kind: 'view', delta: {
        bands: [], visibility: [], metrics: { rows: [], cols: [] },
        counts: { visible: 10, total: 10 }, selection: { active: null, anchor: null, ranges: [] },
        owner, workbookOwner: workbook,
      } } },
      { ...base, kind: 'cancelled' }, { ...base, kind: 'incomplete', missing: [] },
      { ...base, kind: 'quota', resource: 'storage', limitBytes: 1024 },
      { ...base, kind: 'error', code: 'page-error', retryable: true },
    ] satisfies WorkerReply[]
    expect(commands).toHaveLength(10)
    expect(replies).toHaveLength(6)
    expectTypeOf<WorkerReply['kind']>().toEqualTypeOf<'progress' | 'complete' | 'cancelled' | 'incomplete' | 'quota' | 'error'>()
    expectTypeOf<Extract<SourceErrorCode, 'quota' | 'cancelled' | 'invalid-cursor'>>().toEqualTypeOf<'quota' | 'cancelled' | 'invalid-cursor'>()
    // @ts-expect-error Local viewGen alone cannot own an aggregate workbook query.
    const badSearch: WorkerCommand = { kind: 'search', requestId: 'bad', owner, query: { text: '', caseSensitive: false } }
    expect(badSearch).toBeDefined()
  })
})
