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
  overlappingRange,
  overlappingRows,
  type CellAnchoredDrawing,
  type ColumnIntervalView,
  type PopulatedCellSource,
  type PopulatedPage,
  type RowIntervalView,
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
  it('keeps a full-domain merge as one record with bounded visits', () => {
    const index = singleMergeIndex(fullDomain)
    expect(index.recordCount).toBe(1)
    let visits = 0
    const probe = (n: number) => { visits += n }
    expect(index.query({ r0: 500000, r1: 500100, c0: 0, c1: MAX_COLS - 1 }, { countVisits: probe }).merges).toHaveLength(1)
    expect(index.query({ r0: 0, r1: 0, c0: 0, c1: 0 }, { countVisits: probe }).merges).toHaveLength(1)
    expect(visits).toBeLessThanOrEqual(16)
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

describe('first-hit search and resumable visits', () => {
  it('finds the first hit without draining the index', async () => {
    const { index } = await f5Index()
    const band: SourceRect = { r0: 0, r1: 199999, c0: 0, c1: MAX_COLS - 1 }
    expect(index.firstWhere(band, hit => hit.kind === 'merge')).toMatchObject({
      kind: 'merge', record: { id: 'merge-0', range: { r0: 0, r1: 1, c0: 0, c1: 1 } },
    })
    expect(index.firstWhere(band, hit => hit.kind === 'spill')).toMatchObject({
      kind: 'spill', record: { anchor: { row: 32, col: 0 } },
    })
    expect(index.firstWhere({ r0: 500000, r1: 600000, c0: 0, c1: 0 }, () => true)).toBeUndefined()
  })
  it('firstWhere skips non-qualifying hits and keeps scanning', async () => {
    const { index } = await f5Index()
    const hit = index.firstWhere(fullDomain, entry => entry.kind === 'merge' && entry.record.range.r0 > 10)
    expect(hit).toMatchObject({ kind: 'merge', record: { id: 'merge-1' } })
  })
  it('resumed pages never re-collect prior hits', () => {
    const random = createPrng(303)
    const records = Array.from({ length: 2000 }, (_, id) => {
      const r0 = Math.floor(random() * 100000)
      return { rect: { r0, r1: r0 + Math.floor(random() * 50), c0: Math.floor(random() * 100), c1: 0 }, value: { id } }
    }).map(entry => ({ rect: { ...entry.rect, c1: entry.rect.c0 + 5 }, value: entry.value }))
    const index = RectangleIndex.build(records)
    const rect: SourceRect = { r0: 0, r1: 99999, c0: 0, c1: 99 }
    const full = index.query(rect, { limit: 1000000 })
    expect(full.nextCursor).toBeUndefined()
    let cursor: string | undefined
    const seen: number[] = []
    let pages = 0
    let visits = 0
    for (;;) {
      const page = index.query(rect, cursor === undefined ? { limit: 7, countVisits: n => { visits += n } } : { limit: 7, cursor, countVisits: n => { visits += n } })
      for (const item of page.items) seen.push(item.id)
      pages++
      if (page.nextCursor === undefined) break
      cursor = page.nextCursor
      if (pages > 1000) throw new Error('cursor did not terminate')
    }
    expect(seen.sort((a, b) => a - b)).toEqual(full.items.map(item => item.id).sort((a, b) => a - b))
    expect(pages).toBeGreaterThan(2)
    expect(visits).toBeLessThan(8 * (seen.length + pages * 64))
  })
  it('split-spill unions verify exact extents: gap-only queries miss', () => {
    const spill = {
      anchor: { row: 10, col: 0 }, kind: 'dynamic' as const,
      saved: { r0: 10, r1: 12, c0: 0, c1: 0 }, evaluated: { r0: 20, r1: 22, c0: 0, c1: 0 },
    }
    const result: StructuralResult = {
      stamp, completeness: 'complete', merges: [], tables: [], spills: [spill], drawings: [], missing: [],
    }
    const index = new StructuralIndex(result)
    expect(index.recordCount).toBe(1)
    expect(index.query({ r0: 14, r1: 18, c0: 0, c1: 0 }).spills).toHaveLength(0)
    expect(index.query({ r0: 10, r1: 10, c0: 0, c1: 0 }).spills).toHaveLength(1)
    expect(index.query({ r0: 21, r1: 21, c0: 0, c1: 0 }).spills).toHaveLength(1)
    expect(index.firstWhere({ r0: 14, r1: 18, c0: 0, c1: 0 }, hit => hit.kind === 'spill')).toBeUndefined()
  })
})

describe('column-dimension visit bound', () => {
  it('pins visits per hit on 1M small records with random small queries', () => {
    const random = createPrng(777)
    const count = 1000000
    const records = new Array(count)
    for (let id = 0; id < count; id++) {
      const r0 = Math.floor(random() * MAX_ROWS)
      const c0 = Math.floor(random() * MAX_COLS)
      records[id] = {
        rect: {
          r0, r1: Math.min(MAX_ROWS - 1, r0 + Math.floor(random() * 20)),
          c0, c1: Math.min(MAX_COLS - 1, c0 + Math.floor(random() * 20)),
        },
        value: { id },
      }
    }
    const index = RectangleIndex.build(records)
    expect(index.size).toBe(count)
    for (let query = 0; query < 30; query++) {
      const r0 = Math.floor(random() * MAX_ROWS)
      const c0 = Math.floor(random() * MAX_COLS)
      const rect: SourceRect = {
        r0, r1: Math.min(MAX_ROWS - 1, r0 + 30), c0, c1: Math.min(MAX_COLS - 1, c0 + 30),
      }
      let visits = 0
      const page = index.query(rect, { limit: 100000, countVisits: n => { visits += n } })
      expect(page.nextCursor).toBeUndefined()
      expect(visits).toBeLessThanOrEqual(8 * (page.items.length + 64))
    }
  })
  it('bounds the adversarial full-height column case by row hits, not hits', () => {
    const stripes = Array.from({ length: MAX_COLS }, (_v, col) => ({
      rect: { r0: 0, r1: MAX_ROWS - 1, c0: col, c1: col }, value: { id: col },
    }))
    const index = RectangleIndex.build(stripes)
    let visits = 0
    const page = index.query({ r0: 100, r1: 200, c0: 5, c1: 5 }, { countVisits: n => { visits += n } })
    expect(page.items.map(item => item.id)).toEqual([5])
    expect(visits).toBeLessThanOrEqual(4 * MAX_COLS)
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
  it('clones an own "__proto__" key as data without changing the prototype', () => {
    const hostile = JSON.parse(
      '{"id":"m","range":{"r0":0,"r1":1,"c0":0,"c1":1},"anchor":{"row":0,"col":0},"__proto__":{"polluted":true}}',
    ) as MergeExtent
    const index = new StructuralIndex({ stamp, completeness: 'complete', merges: [hostile], tables: [], spills: [], drawings: [], missing: [] })
    const record = index.query({ r0: 0, r1: 0, c0: 0, c1: 0 }).merges[0]
    expect(Object.getPrototypeOf(record)).toBe(Object.prototype)
    expect((record as unknown as { polluted?: boolean }).polluted).toBeUndefined()
    expect(Object.hasOwn(record, '__proto__')).toBe(true)
    expect(Object.isFrozen(record)).toBe(true)
    expect(Object.getOwnPropertyDescriptor(record, '__proto__')?.value).toEqual({ polluted: true })
    expect(Object.isFrozen(Object.getOwnPropertyDescriptor(record, '__proto__')?.value)).toBe(true)
    expect(({} as { polluted?: boolean }).polluted).toBeUndefined()
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
  it('transposes populated rows per column', async () => {
    const handle = createFixture('F5')
    const built = await ColumnChunkIndex.build(fixtureSource(handle), { r0: 0, r1: 9 })
    expect(built.complete).toBe(true)
    expect(built.index.rowsOf(0)).toEqual([{ r0: 0, r1: 0 }, { r0: 3, r1: 4 }, { r0: 6, r1: 6 }, { r0: 9, r1: 9 }])
    expect(built.index.rowsOf(9)).toEqual([{ r0: 0, r1: 0 }, { r0: 3, r1: 3 }, { r0: 6, r1: 6 }, { r0: 9, r1: 9 }])
    expect(built.index.rowsOf(10)).toEqual([])
  })
})

describe('interval overlap primitives', () => {
  it('finds overlapping sub-arrays at the boundaries', () => {
    const intervals = [{ c0: 0, c1: 0 }, { c0: 4, c1: 9 }, { c0: 10, c1: 10 }, { c0: 20, c1: 30 }]
    expect(overlappingRange([], 0, 5)).toEqual([])
    expect(overlappingRange(intervals, 1, 3)).toEqual([])
    expect(overlappingRange(intervals, 0, 0)).toEqual([{ c0: 0, c1: 0 }])
    expect(overlappingRange(intervals, 9, 10)).toEqual([{ c0: 4, c1: 9 }, { c0: 10, c1: 10 }])
    expect(overlappingRange(intervals, 0, 100)).toEqual(intervals)
    expect(overlappingRange(intervals, 31, 40)).toEqual([])
    const rows = [{ r0: 2, r1: 5 }, { r0: 9, r1: 9 }]
    expect(overlappingRows(rows, 5, 9)).toEqual(rows)
    expect(overlappingRows(rows, 6, 8)).toEqual([])
    expect(overlappingRows(rows, 0, 1)).toEqual([])
  })
})

describe('Excel current-region detection', () => {
  function handViews(cells: ReadonlyArray<{ row: number; col: number }>, band: { r0: number; r1: number }) {
    const source: PopulatedCellSource = {
      async iterate(range: SourceRect): Promise<PopulatedPage> {
        return {
          items: cells.filter(cell => cell.row >= range.r0 && cell.row <= range.r1),
          completeness: 'complete',
          missing: [],
        }
      },
    }
    return { source, band }
  }
  async function views(cells: ReadonlyArray<{ row: number; col: number }>, band: { r0: number; r1: number }) {
    const { source } = handViews(cells, band)
    return {
      rows: (await RowChunkIndex.build(source, band)).index,
      cols: (await ColumnChunkIndex.build(source, band)).index,
    }
  }
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
  it('absorbs diagonally touching cells like Excel (A1+B2 -> A1:B2)', async () => {
    const { rows, cols } = await views([{ row: 0, col: 0 }, { row: 1, col: 1 }], { r0: 0, r1: 9 })
    expect(currentRegion({ row: 0, col: 0 }, rows, cols)).toEqual({ r0: 0, r1: 1, c0: 0, c1: 1 })
    expect(currentRegion({ row: 1, col: 1 }, rows, cols)).toEqual({ r0: 0, r1: 1, c0: 0, c1: 1 })
  })
  it('grows through ring adjacency (A1:C1+A2+C3 -> A1:C3)', async () => {
    const cells = [
      { row: 0, col: 0 }, { row: 0, col: 1 }, { row: 0, col: 2 }, { row: 1, col: 0 }, { row: 2, col: 2 },
    ]
    const { rows, cols } = await views(cells, { r0: 0, r1: 9 })
    expect(currentRegion({ row: 0, col: 0 }, rows, cols)).toEqual({ r0: 0, r1: 2, c0: 0, c1: 2 })
  })
  it('absorbs diagonal chains and keeps L-shapes and holes as bounding boxes', async () => {
    const chain = [{ row: 10, col: 10 }, { row: 11, col: 11 }, { row: 12, col: 12 }]
    const grown = await views(chain, { r0: 0, r1: 20 })
    expect(currentRegion({ row: 10, col: 10 }, grown.rows, grown.cols)).toEqual({ r0: 10, r1: 12, c0: 10, c1: 12 })
    const ell = [{ row: 30, col: 30 }, { row: 31, col: 30 }, { row: 31, col: 31 }]
    const ellViews = await views(ell, { r0: 20, r1: 40 })
    expect(currentRegion({ row: 30, col: 30 }, ellViews.rows, ellViews.cols)).toEqual({ r0: 30, r1: 31, c0: 30, c1: 31 })
    const donut = [
      { row: 50, col: 50 }, { row: 50, col: 51 }, { row: 50, col: 52 },
      { row: 51, col: 50 }, { row: 51, col: 52 },
      { row: 52, col: 50 }, { row: 52, col: 51 }, { row: 52, col: 52 },
    ]
    const donutViews = await views(donut, { r0: 40, r1: 60 })
    expect(currentRegion({ row: 50, col: 50 }, donutViews.rows, donutViews.cols)).toEqual({ r0: 50, r1: 52, c0: 50, c1: 52 })
  })
  it('returns unknown for incomplete indexes instead of bounding partial data', async () => {
    const handle = createFixture('F5')
    const pending: SourceRect = { r0: 0, r1: 60, c0: 0, c1: MAX_COLS - 1 }
    const source = fixtureSource(handle, [pending])
    const band = { r0: 0, r1: 60 }
    const rows = (await RowChunkIndex.build(source, band)).index
    const cols = (await ColumnChunkIndex.build(source, band)).index
    expect(rows.complete).toBe(false)
    expect(currentRegion({ row: 0, col: 0 }, rows, cols)).toMatchObject({ kind: 'unknown' })
  })
  it('returns unknown when the region touches the band edge', async () => {
    const cells = [{ row: 5, col: 5 }, { row: 6, col: 5 }, { row: 7, col: 5 }, { row: 8, col: 5 }, { row: 9, col: 5 }]
    const { rows, cols } = await views(cells, { r0: 0, r1: 9 })
    const outcome = currentRegion({ row: 7, col: 5 }, rows, cols)
    expect(outcome).toMatchObject({ kind: 'unknown' })
    const wide = await views(cells, { r0: 0, r1: 20 })
    expect(currentRegion({ row: 7, col: 5 }, wide.rows, wide.cols)).toEqual({ r0: 5, r1: 9, c0: 5, c1: 5 })
  })
  it('returns unknown when the column index covers a narrower band', async () => {
    const { source, band } = handViews([{ row: 5, col: 5 }], { r0: 0, r1: 9 })
    const rows = (await RowChunkIndex.build(source, band)).index
    const narrow = (await ColumnChunkIndex.build(source, { r0: 6, r1: 9 })).index
    expect(currentRegion({ row: 5, col: 5 }, rows, narrow)).toMatchObject({ kind: 'unknown' })
  })
  it('bounds interval probes on wide combs (counting probe)', async () => {
    const stripes = 2000
    const cells = [{ row: 4, col: 0 }, { row: 6, col: 0 }]
    for (let k = 0; k < stripes; k++) cells.push({ row: 5, col: k * 2 })
    const { source, band } = handViews(cells, { r0: 0, r1: 9 })
    const rows = (await RowChunkIndex.build(source, band)).index
    const cols = (await ColumnChunkIndex.build(source, band)).index
    let accesses = 0
    const counted = (intervals: readonly { c0: number; c1: number }[]) => new Proxy(intervals, {
      get(target, property, receiver) {
        if (typeof property === 'string' && /^(0|[1-9][0-9]*)$/.test(property)) accesses++
        return Reflect.get(target, property, receiver)
      },
    })
    const countedRows: RowIntervalView = {
      band: rows.band, complete: rows.complete, missing: rows.missing,
      columnsOf: row => counted(rows.columnsOf(row) as { c0: number; c1: number }[]),
    }
    const countedCols: ColumnIntervalView = {
      band: cols.band, complete: cols.complete, missing: cols.missing,
      rowsOf: col => cols.rowsOf(col),
    }
    expect(currentRegion({ row: 4, col: 0 }, countedRows, countedCols)).toEqual({ r0: 4, r1: 6, c0: 0, c1: 0 })
    expect(accesses).toBeLessThan(stripes / 4)
  })
  it('handles a huge sparse sheet without blowup (counting probe)', async () => {
    const random = createPrng(4242)
    const cells = Array.from({ length: 5000 }, () => ({ row: Math.floor(random() * MAX_ROWS), col: Math.floor(random() * MAX_COLS) }))
    cells.push({ row: 100, col: 100 }, { row: 100, col: 101 }, { row: 101, col: 100 })
    const { source, band } = handViews(cells, { r0: 0, r1: MAX_ROWS - 1 })
    const rows = (await RowChunkIndex.build(source, band)).index
    const cols = (await ColumnChunkIndex.build(source, band)).index
    let calls = 0
    const countedRows: RowIntervalView = {
      band: rows.band, complete: rows.complete, missing: rows.missing,
      columnsOf: row => { calls++; return rows.columnsOf(row) },
    }
    const countedCols: ColumnIntervalView = {
      band: cols.band, complete: cols.complete, missing: cols.missing,
      rowsOf: col => { calls++; return cols.rowsOf(col) },
    }
    expect(currentRegion({ row: 100, col: 100 }, countedRows, countedCols)).toEqual({ r0: 100, r1: 101, c0: 100, c1: 101 })
    expect(calls).toBeLessThan(100)
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

describe('structural result error handling', () => {
  it('treats a set result error as incomplete instead of authorizing', () => {
    const result: StructuralResult = {
      stamp, completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [],
      error: { code: 'page-error', retryable: false, message: 'structural page failed' },
    }
    expect(new StructuralIndex(result).completeness).toBe('incomplete')
  })
})

describe('record ownership', () => {
  it('isolates the index from later caller mutation', () => {
    const spill = {
      anchor: { row: 32, col: 0 }, kind: 'fixed' as const,
      saved: { r0: 32, r1: 34, c0: 0, c1: 1 } as SourceRect,
    }
    const result: StructuralResult = {
      stamp, completeness: 'complete', merges: [], tables: [], spills: [spill], drawings: [], missing: [],
    }
    const index = new StructuralIndex(result)
    const query: SourceRect = { r0: 33, r1: 33, c0: 1, c1: 1 }
    expect(index.query(query).spills).toHaveLength(1)
    spill.saved = { r0: 32, r1: 34, c0: 5, c1: 5 }
    spill.anchor.col = 9
    expect(index.query(query).spills).toHaveLength(1)
    expect(index.query(query).spills[0].anchor).toEqual({ row: 32, col: 0 })
  })
  it('freezes returned records so mutation attempts throw', async () => {
    const { index } = await f5Index()
    const page = index.query({ r0: 0, r1: 1, c0: 0, c1: 1 })
    expect(page.merges).toHaveLength(1)
    expect(() => { (page.merges[0] as unknown as { id: string }).id = 'x' }).toThrow()
    expect(() => { (page.merges[0].range as unknown as { r1: number }).r1 = 999 }).toThrow()
    expect(() => { (page.merges[0].anchor as unknown as { row: number }).row = 7 }).toThrow()
    expect(page.merges[0]).toMatchObject({ id: 'merge-0', range: { r0: 0, r1: 1, c0: 0, c1: 1 } })
  })
})
