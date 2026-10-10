/** S3.1 structural range index and chunk indexes (spec sections 3, 4.5, 6.1, 6.3). */
import { describe, expect, it } from 'vitest'
import { MAX_COLS, MAX_ROWS, type SourceRect } from '../src/xlsx/view/coordinates'
import type { DrawingAnchor, MergeExtent, SpillExtent, StructuralResult, TableExtent } from '../src/xlsx/view/contracts'
import {
  ColumnChunkIndex,
  RectangleIndex,
  RowChunkIndex,
  StructuralIndex,
  currentRegion,
  findOutsideColumnData,
  fromStoreScan,
  type CellAnchoredDrawing,
  type PopulatedCellSource,
  type PopulatedPage,
  type StructuralPage,
} from '../src/xlsx/view/structure'
import { createFixture, createPrng, readOptionsFor, type FixtureHandle } from './helpers/xlsx-viewer-fixtures'

const stamp = { docId: 'test-doc', sheetId: 'test-sheet', valueGen: 0 }
const fullDomain: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
function rectsIntersect(a: SourceRect, b: SourceRect): boolean {
  return a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1
}

async function f5Index(): Promise<{ handle: FixtureHandle; index: StructuralIndex }> {
  const handle = createFixture('F5')
  const store = handle.createStore()
  const result = await store.queryStructural(fullDomain, readOptionsFor(store))
  expect(result.completeness).toBe('complete')
  return { handle, index: new StructuralIndex(result) }
}

/** In-memory populated-cell source over a fixture handle with explicit pending regions. */
function fixtureSource(handle: FixtureHandle, pending: readonly SourceRect[] = [], pageSize = 64): PopulatedCellSource {
  const keyOf = (range: SourceRect): string => `${range.r0},${range.r1},${range.c0},${range.c1}`
  return {
    async iterate(range: SourceRect, cursor?: string): Promise<PopulatedPage> {
      if (cursor !== undefined) {
        const parts = cursor.split(':')
        const offset = Number(parts[2])
        if (parts.length !== 3 || parts[0] !== 'f5pop' || parts[1] !== keyOf(range)
          || !Number.isSafeInteger(offset) || offset < 0 || parts[2] !== String(offset)) {
          throw new RangeError('Invalid populated cursor')
        }
        const cells = handle.scanCells(range, offset, pageSize, handle, [...pending])
        const missing = pending.filter(region => rectsIntersect(region, range))
        return {
          items: cells.map(({ row, col }) => ({ row, col })),
          completeness: missing.length > 0 ? 'loading' : 'complete',
          missing,
          nextCursor: cells.length === pageSize ? `f5pop:${keyOf(range)}:${offset + pageSize}` : undefined,
        }
      }
      const cells = handle.scanCells(range, 0, pageSize, handle, [...pending])
      const missing = pending.filter(region => rectsIntersect(region, range))
      return {
        items: cells.map(({ row, col }) => ({ row, col })),
        completeness: missing.length > 0 ? 'loading' : 'complete',
        missing,
        nextCursor: cells.length === pageSize ? `f5pop:${keyOf(range)}:${pageSize}` : undefined,
      }
    },
  }
}

describe('structural index over F5 records', () => {
  it('indexes merges, tables, spills and cell-anchored drawings with exact bounds', async () => {
    const { index } = await f5Index()
    const page = index.query(fullDomain)
    expect(page.completeness).toBe('complete')
    expect(page.merges.map(merge => merge.range)).toEqual([
      { r0: 0, r1: 1, c0: 0, c1: 1 },
      { r0: 99, r1: 199999, c0: 1, c1: 100 },
    ])
    expect(page.tables.map(table => [table.range, table.headerRows, table.totalsRows])).toEqual([
      [{ r0: 4, r1: 20, c0: 0, c1: 3 }, 1, 1],
      [{ r0: 24, r1: 30, c0: 0, c1: 2 }, 0, 0],
    ])
    expect(page.spills.map(spill => [spill.anchor, spill.kind])).toEqual([
      [{ row: 32, col: 0 }, 'fixed'],
      [{ row: 40, col: 3 }, 'dynamic'],
    ])
    expect(page.drawings.map(drawing => drawing.id)).toEqual(['drawing-1', 'drawing-2'])
  })

  it('finds a spill through its child cells even though the child has no formula', async () => {
    const { handle, index } = await f5Index()
    expect(handle.cellAt(45, 3)?.formula).toBeUndefined()
    const page = index.query({ r0: 44, r1: 46, c0: 0, c1: 3 })
    expect(page.spills.map(spill => spill.anchor)).toEqual([{ row: 40, col: 3 }])
  })

  it('finds a fixed spill through a child-only extent with the anchor outside the query', async () => {
    const { index } = await f5Index()
    const page = index.query({ r0: 33, r1: 33, c0: 1, c1: 1 })
    expect(page.spills.map(spill => [spill.anchor, spill.kind])).toEqual([[{ row: 32, col: 0 }, 'fixed']])
  })

  it('preserves an explicit headerRows 0 table', async () => {
    const { index } = await f5Index()
    const page = index.query({ r0: 24, r1: 30, c0: 0, c1: 2 })
    expect(page.tables).toHaveLength(1)
    expect(page.tables[0].headerRows).toBe(0)
    expect(page.tables[0].totalsRows).toBe(0)
  })

  it('resolves one-cell anchor bounds and keeps absolute anchors out of spatial queries', async () => {
    const { index } = await f5Index()
    const page = index.query({ r0: 70, r1: 74, c0: 2, c1: 4 })
    expect(page.drawings).toHaveLength(1)
    expect(page.drawings[0]).toMatchObject({
      id: 'drawing-1', kind: 'oneCell', from: { row: 70, col: 2 }, bounds: { r0: 70, r1: 74, c0: 2, c1: 4 },
    })
    const full = index.query(fullDomain)
    expect(full.drawings.map(drawing => drawing.id)).toEqual(['drawing-1', 'drawing-2'])
    // The S0.2 store reports source-indexed records only, so absolute anchors
    // never reach the index through it (placement is an S5 concern); the
    // separation itself is covered by the hand-built case below.
    expect(index.absoluteDrawings).toEqual([])
  })
})

describe('giant-rectangle memory behavior', () => {
  function singleMergeIndex(range: SourceRect): StructuralIndex {
    const result: StructuralResult = {
      stamp, completeness: 'complete', tables: [], spills: [], drawings: [], missing: [],
      merges: [{ id: 'giant', range, anchor: { row: range.r0, col: range.c0 } }],
    }
    return new StructuralIndex(result)
  }
  it('keeps a full-domain merge as one record with bounded heap growth', () => {
    singleMergeIndex(fullDomain)
    const before = process.memoryUsage().heapUsed
    const index = singleMergeIndex(fullDomain)
    expect(index.recordCount).toBe(1)
    expect(index.query({ r0: 500000, r1: 500100, c0: 0, c1: MAX_COLS - 1 }).merges).toHaveLength(1)
    expect(index.query({ r0: 0, r1: 0, c0: 0, c1: 0 }).merges).toHaveLength(1)
    const after = process.memoryUsage().heapUsed
    expect(after - before).toBeLessThan(16 * 1024 * 1024)
  })
  it('gives narrow-column giant rectangles no phantom coverage', () => {
    const index = singleMergeIndex({ r0: 0, r1: MAX_ROWS - 1, c0: 5, c1: 5 })
    expect(index.query({ r0: 100, r1: 200, c0: 6, c1: 10 }).merges).toHaveLength(0)
    expect(index.query({ r0: 100, r1: 200, c0: 5, c1: 5 }).merges).toHaveLength(1)
    expect(index.query({ r0: MAX_ROWS - 1, r1: MAX_ROWS - 1, c0: 5, c1: 5 }).merges).toHaveLength(1)
  })
})

describe('bounded pages and cursors', () => {
  async function drain(index: StructuralIndex, rect: SourceRect, limit: number): Promise<StructuralPage> {
    const merges: MergeExtent[] = []
    const tables: TableExtent[] = []
    const spills: SpillExtent[] = []
    const drawings: CellAnchoredDrawing[] = []
    let cursor: string | undefined
    let pages = 0
    for (;;) {
      const page = index.query(rect, cursor === undefined ? { limit } : { limit, cursor })
      for (const merge of page.merges) merges.push(merge)
      for (const table of page.tables) tables.push(table)
      for (const spill of page.spills) spills.push(spill)
      for (const drawing of page.drawings) drawings.push(drawing)
      pages++
      if (page.nextCursor === undefined) break
      cursor = page.nextCursor
      if (pages > 100) throw new Error('cursor did not terminate')
    }
    return {
      completeness: 'complete',
      missing: [],
      merges: merges,
      tables: tables,
      spills: spills,
      drawings: drawings,
    }
  }
  it('pages through hits in deterministic order and rejects foreign cursors', async () => {
    const { index } = await f5Index()
    const full = index.query(fullDomain)
    expect(full.nextCursor).toBeUndefined()
    const first = index.query(fullDomain, { limit: 2 })
    expect([...first.merges, ...first.tables, ...first.spills, ...first.drawings]).toHaveLength(2)
    expect(first.nextCursor).toBeDefined()
    const drained = await drain(index, fullDomain, 2)
    expect(drained.merges.map(merge => merge.id)).toEqual(full.merges.map(merge => merge.id))
    expect(drained.tables.map(table => table.id)).toEqual(full.tables.map(table => table.id))
    expect(drained.spills.map(spill => spill.anchor)).toEqual(full.spills.map(spill => spill.anchor))
    expect(drained.drawings.map(drawing => drawing.id)).toEqual(full.drawings.map(drawing => drawing.id))
    expect(() => index.query(fullDomain, { limit: 2, cursor: 'garbage' })).toThrow(RangeError)
    expect(() => index.query({ r0: 0, r1: 10, c0: 0, c1: 10 }, { limit: 2, cursor: first.nextCursor })).toThrow(RangeError)
    expect(() => index.query(fullDomain, { limit: 0 })).toThrow(RangeError)
    expect(() => index.query(fullDomain, { limit: 1.5 })).toThrow(RangeError)
  })
})

describe('rectangle index differential vs brute force', () => {
  it('matches a list scan on thousands of random rects including domain edges', () => {
    const random = createPrng(20261010)
    const integer = (max: number): number => Math.floor(random() * (max + 1))
    const records = Array.from({ length: 600 }, (_, id) => {
      const edge = random()
      const r0 = edge < 0.1 ? 0 : edge < 0.15 ? MAX_ROWS - 1 : integer(MAX_ROWS - 1)
      const c0 = edge < 0.1 ? 0 : edge < 0.15 ? MAX_COLS - 1 : integer(MAX_COLS - 1)
      const giant = random() < 0.1
      const r1 = giant ? MAX_ROWS - 1 : Math.min(MAX_ROWS - 1, r0 + integer(5000))
      const c1 = giant ? MAX_COLS - 1 : Math.min(MAX_COLS - 1, c0 + integer(200))
      return { rect: { r0, r1, c0, c1 }, value: { id } }
    })
    const index = RectangleIndex.build(records)
    expect(index.size).toBe(600)
    let checked = 0
    for (let query = 0; query < 5000; query++) {
      const r0 = integer(MAX_ROWS - 1)
      const c0 = integer(MAX_COLS - 1)
      const rect: SourceRect = {
        r0,
        r1: Math.min(MAX_ROWS - 1, r0 + integer(query % 7 === 0 ? MAX_ROWS - 1 : 3000)),
        c0,
        c1: Math.min(MAX_COLS - 1, c0 + integer(query % 11 === 0 ? MAX_COLS - 1 : 500)),
      }
      const expected = records.filter(entry => rectsIntersect(entry.rect, rect)).map(entry => entry.value.id).sort((a, b) => a - b)
      const page = index.query(rect, { limit: 100000 })
      expect(page.items.map(item => item.id).sort((a, b) => a - b)).toEqual(expected)
      checked++
    }
    expect(checked).toBe(5000)
  })
})

describe('query immutability', () => {
  it('freezes returned records and never mutates the input result', async () => {
    const { index } = await f5Index()
    const before = JSON.stringify(index.query(fullDomain))
    const page = index.query({ r0: 0, r1: 50, c0: 0, c1: 10 })
    expect(Object.isFrozen(page.merges)).toBe(true)
    expect(Object.isFrozen(page.tables)).toBe(true)
    expect(Object.isFrozen(page.spills)).toBe(true)
    expect(Object.isFrozen(page.drawings)).toBe(true)
    expect(JSON.stringify(index.query(fullDomain))).toBe(before)
  })
  it('rejects invalid query rects instead of scanning', async () => {
    const { index } = await f5Index()
    expect(() => index.query({ r0: 10, r1: 5, c0: 0, c1: 0 })).toThrow()
    expect(() => index.query({ r0: 0, r1: MAX_ROWS, c0: 0, c1: 0 })).toThrow()
  })
})

describe('row and column chunk indexes', () => {
  it('answers outside-column occupancy from sparse cells without full-column scans', async () => {
    const handle = createFixture('F5')
    let calls = 0
    const source: PopulatedCellSource = {
      async iterate(range: SourceRect, cursor?: string): Promise<PopulatedPage> {
        calls++
        expect(range).toEqual({ r0: 0, r1: 9, c0: 0, c1: MAX_COLS - 1 })
        return fixtureSource(handle, [], 1000).iterate(range, cursor)
      },
    }
    const built = await RowChunkIndex.build(source, { r0: 0, r1: 9 })
    expect(built.complete).toBe(true)
    expect(built.index.columnsOf(1)).toEqual([{ c0: 2, c1: 2 }, { c0: 5, c1: 5 }, { c0: 8, c1: 8 }])
    expect(built.index.outsideColumnData(0, 1)).toEqual({ kind: 'occupied', at: { row: 0, col: 3 } })
    expect(built.index.outsideColumnData(0, 9)).toEqual({ kind: 'clear' })
    expect(calls).toBeLessThanOrEqual(2)
  })
  it('finds data separated by entirely blank columns', async () => {
    const cells = [
      { row: 2, col: 3 }, { row: 4, col: 3 },
    ]
    const source: PopulatedCellSource = {
      async iterate(range: SourceRect): Promise<PopulatedPage> {
        return {
          items: cells.filter(cell => cell.row >= range.r0 && cell.row <= range.r1),
          completeness: 'complete',
          missing: [],
        }
      },
    }
    const built = await RowChunkIndex.build(source, { r0: 0, r1: 9 })
    expect(built.index.outsideColumnData(0, 1)).toEqual({ kind: 'occupied', at: { row: 2, col: 3 } })
  })
  it('streams outside-only rects with the page limit to every request', async () => {
    const seen: Array<{ rect: SourceRect; limit: number | undefined }> = []
    const source: PopulatedCellSource = {
      async iterate(range: SourceRect, _cursor?: string, limit?: number): Promise<PopulatedPage> {
        seen.push({ rect: { ...range }, limit })
        return { items: [], completeness: 'complete', missing: [] }
      },
    }
    expect(await findOutsideColumnData(source, { r0: 0, r1: 9 }, 2, 5, 23)).toEqual({ kind: 'clear' })
    expect(seen).toEqual([
      { rect: { r0: 0, r1: 9, c0: 0, c1: 1 }, limit: 23 },
      { rect: { r0: 0, r1: 9, c0: 6, c1: MAX_COLS - 1 }, limit: 23 },
    ])
    await RowChunkIndex.build(source, { r0: 0, r1: 9 }, 17)
    expect(seen[seen.length - 1]).toEqual({ rect: { r0: 0, r1: 9, c0: 0, c1: MAX_COLS - 1 }, limit: 17 })
  })
  it('reports clear for empty bands and never fakes safety for pending pages', async () => {
    const handle = createFixture('F5')
    const empty: PopulatedCellSource = {
      async iterate(): Promise<PopulatedPage> {
        return { items: [], completeness: 'complete', missing: [] }
      },
    }
    const clear = await RowChunkIndex.build(empty, { r0: 60, r1: 65 })
    expect(clear.complete).toBe(true)
    expect(clear.index.outsideColumnData(0, 1)).toEqual({ kind: 'clear' })
    const pending: SourceRect = { r0: 60, r1: 65, c0: 0, c1: MAX_COLS - 1 }
    const waiting = await RowChunkIndex.build(fixtureSource(handle, [pending]), { r0: 60, r1: 65 })
    expect(waiting.complete).toBe(false)
    expect(waiting.missing).toEqual([pending])
    expect(waiting.index.outsideColumnData(0, 1)).toEqual({ kind: 'unknown', missing: [pending] })
  })
  it('treats style-only marks as non-data by construction', async () => {
    const styleOnly: PopulatedCellSource = {
      async iterate(): Promise<PopulatedPage> {
        return { items: [], completeness: 'complete', missing: [] }
      },
    }
    const built = await RowChunkIndex.build(styleOnly, { r0: 4, r1: 4 })
    expect(built.index.outsideColumnData(0, 0)).toEqual({ kind: 'clear' })
  })
  it('transposes populated rows per column', async () => {
    const handle = createFixture('F5')
    const built = await ColumnChunkIndex.build(fixtureSource(handle), { r0: 0, r1: 9 })
    expect(built.complete).toBe(true)
    expect(built.index.rowsOf(0)).toEqual([{ r0: 0, r1: 0 }, { r0: 3, r1: 4 }, { r0: 6, r1: 6 }, { r0: 9, r1: 9 }])
    expect(built.index.rowsOf(9)).toEqual([{ r0: 0, r1: 0 }, { r0: 3, r1: 3 }, { r0: 6, r1: 6 }, { r0: 9, r1: 9 }])
    expect(built.index.rowsOf(10)).toEqual([])
  })
})

describe('sparse current-region detection', () => {
  it('expands through populated adjacency and stops at blank separators', async () => {
    const handle = createFixture('F5')
    const source = fixtureSource(handle)
    const band = { r0: 0, r1: 60 }
    const rows = (await RowChunkIndex.build(source, band)).index
    const cols = (await ColumnChunkIndex.build(source, band)).index
    expect(currentRegion({ row: 45, col: 3 }, rows, cols)).toEqual({ r0: 44, r1: 46, c0: 3, c1: 3 })
    expect(currentRegion({ row: 0, col: 0 }, rows, cols)).toEqual({ r0: 0, r1: 0, c0: 0, c1: 0 })
    expect(currentRegion({ row: 10, col: 10 }, rows, cols)).toBeNull()
  })
  it('excludes diagonally touching cells', async () => {
    const cells = [
      { row: 2, col: 2 }, { row: 2, col: 3 }, { row: 3, col: 2 }, { row: 3, col: 3 },
      { row: 4, col: 4 },
    ]
    const source: PopulatedCellSource = {
      async iterate(range: SourceRect): Promise<PopulatedPage> {
        return {
          items: cells.filter(cell => cell.row >= range.r0 && cell.row <= range.r1),
          completeness: 'complete',
          missing: [],
        }
      },
    }
    const band = { r0: 0, r1: 9 }
    const rows = (await RowChunkIndex.build(source, band)).index
    const cols = (await ColumnChunkIndex.build(source, band)).index
    expect(currentRegion({ row: 2, col: 2 }, rows, cols)).toEqual({ r0: 2, r1: 3, c0: 2, c1: 3 })
    expect(currentRegion({ row: 4, col: 4 }, rows, cols)).toEqual({ r0: 4, r1: 4, c0: 4, c1: 4 })
  })
})

describe('store-backed populated source', () => {
  it('propagates store cursors across pages instead of truncating the band', async () => {
    const handle = createFixture('F5')
    const store = handle.createStore()
    const source = fromStoreScan(store, readOptionsFor(store), 2)
    const first = await source.iterate({ r0: 0, r1: 99, c0: 0, c1: MAX_COLS - 1 })
    expect(first.items.length).toBeLessThanOrEqual(2)
    expect(first.nextCursor).toBeDefined()
    const second = await source.iterate({ r0: 0, r1: 99, c0: 0, c1: MAX_COLS - 1 }, first.nextCursor)
    expect(second.items.length).toBeGreaterThan(0)
    const built = await RowChunkIndex.build(source, { r0: 0, r1: 99 })
    expect(built.complete).toBe(true)
    expect(built.index.populatedRows()).toEqual([
      { r0: 0, r1: 9 }, { r0: 32, r1: 32 }, { r0: 40, r1: 40 }, { r0: 44, r1: 46 },
    ])
  })
  it('rejects a source that repeats its cursor instead of looping forever', async () => {
    const looping: PopulatedCellSource = {
      async iterate(): Promise<PopulatedPage> {
        return { items: [], completeness: 'complete', missing: [], nextCursor: 'loop' }
      },
    }
    await expect(RowChunkIndex.build(looping, { r0: 0, r1: 9 })).rejects.toThrow(/cursor/i)
  })
  it('feeds chunk indexes from scanPopulated with small pages and cursors', async () => {
    const handle = createFixture('F5')
    const store = handle.createStore()
    const source = fromStoreScan(store, readOptionsFor(store), 8)
    const built = await RowChunkIndex.build(source, { r0: 0, r1: 9 })
    expect(built.complete).toBe(true)
    expect(built.missing).toEqual([])
    expect(built.index.outsideColumnData(0, 1)).toEqual({ kind: 'occupied', at: { row: 0, col: 3 } })
    const cols = await ColumnChunkIndex.build(fromStoreScan(store, readOptionsFor(store), 8), { r0: 44, r1: 46 })
    expect(cols.index.rowsOf(3)).toEqual([{ r0: 44, r1: 46 }])
  })
  it('surfaces pending pages as loading with the requested range missing', async () => {
    const handle = createFixture('F5')
    const pending: SourceRect = { r0: 0, r1: 9, c0: 0, c1: 100 }
    const store = handle.createStore({ pending: [pending] })
    const source = fromStoreScan(store, readOptionsFor(store))
    const page = await source.iterate({ r0: 0, r1: 9, c0: 0, c1: MAX_COLS - 1 })
    expect(page.completeness).toBe('loading')
    expect(page.missing).toHaveLength(1)
    const built = await RowChunkIndex.build(source, { r0: 0, r1: 9 })
    expect(built.complete).toBe(false)
    expect(built.index.outsideColumnData(0, 1).kind).toBe('unknown')
  })
  it('throws rather than faking pages for failed regions and foreign cursors', async () => {
    const handle = createFixture('F5')
    const failed = handle.createStore({
      errors: [{ range: { r0: 0, r1: 9, c0: 0, c1: MAX_COLS - 1 }, code: 'page-error', retryable: false }],
    })
    await expect(RowChunkIndex.build(fromStoreScan(failed, readOptionsFor(failed)), { r0: 0, r1: 9 }))
      .rejects.toThrow(/page-error/)
    const store = handle.createStore()
    const source = fromStoreScan(store, readOptionsFor(store))
    await expect(source.iterate({ r0: 0, r1: 9, c0: 0, c1: 9 }, 'foreign-cursor')).rejects.toThrow(RangeError)
    await expect(source.iterate({ r0: 9, r1: 0, c0: 0, c1: 9 })).rejects.toThrow()
  })
})

describe('absolute drawing retention', () => {
  it('keeps absolute anchors with their EMU box outside the spatial index', () => {
    const absolute: DrawingAnchor = { id: 'abs-1', kind: 'absolute', extent: { x: 100, y: 200, width: 300, height: 400 } }
    const result: StructuralResult = {
      stamp, completeness: 'complete', merges: [], tables: [], spills: [], drawings: [absolute], missing: [],
    }
    const index = new StructuralIndex(result)
    expect(index.query(fullDomain).drawings).toEqual([])
    expect(index.absoluteDrawings).toEqual([absolute])
    expect(index.recordCount).toBe(0)
  })
})
