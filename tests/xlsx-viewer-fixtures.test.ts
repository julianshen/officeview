/** S0.2 deterministic fixture family (spec §12.5, plan S0.2). */
import { expect, it, describe } from 'vitest'
import { readFileSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxCell } from '../src/xlsx/types'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import type { SourceStore, OverviewQuery } from '../src/xlsx/view/contracts'
import {
  createPrng, hashSeed, cellHash, createFixture, f4StringAt, f4Family, F4_STRING_COUNT, f8Cases, f8CorpusRefs, f7Spec, f7Sheets,
  sha256Hex, readOptionsFor, headQueryHashes, encodeXlsxString, FIXTURE_SEEDS, STYLE_DATE,
  type FixtureId, type FixtureHandle, type FixtureParams, type SyntheticStoreOptions, type CellReader,
} from './helpers/xlsx-viewer-fixtures'

const FIXTURE_DIR = join(__dirname, 'fixtures', 'xlsx-viewer')

function makeStore(handle: FixtureHandle, options?: SyntheticStoreOptions): { store: SourceStore } {
  return { store: handle.createStore(options) }
}

describe('seeded PRNG (no Math.random)', () => {
  it('repeats the same sequence for the same seed', () => {
    const a = createPrng(12345)
    const b = createPrng(12345)
    expect(Array.from({ length: 100 }, () => a())).toEqual(Array.from({ length: 100 }, () => b()))
  })

  it('emits values in [0, 1) and differs across seeds', () => {
    const gen = createPrng(7)
    const values = Array.from({ length: 1000 }, () => gen())
    expect(values.every((v) => v >= 0 && v < 1)).toBe(true)
    expect(new Set(values).size).toBeGreaterThan(900)
    const other = Array.from({ length: 100 }, () => createPrng(8)())
    expect(values.slice(0, 100)).not.toEqual(other)
  })

  it('derives stable numeric seeds from fixture ids', () => {
    expect(hashSeed('F1')).toBe(hashSeed('F1'))
    expect(hashSeed('F1')).not.toBe(hashSeed('F2'))
    expect(Number.isSafeInteger(hashSeed('F1'))).toBe(true)
  })

  it('matches the documented mulberry32 reference vector', () => {
    const gen = createPrng(1)
    expect(gen().toFixed(10)).toBe('0.6270739406')
  })
})

describe('F1 full-extents sparse corners', () => {
  it('reads all four edge cells including XFD1048576 (resolves contracts-test todo)', async () => {
    const handle = createFixture('F1', { nnz: 500 })
    const { store } = makeStore(handle)
    const options = readOptionsFor(store)
    for (const at of [{ row: 0, col: 0 }, { row: 0, col: MAX_COLS - 1 }, { row: MAX_ROWS - 1, col: 0 }, { row: MAX_ROWS - 1, col: MAX_COLS - 1 }]) {
      expect(await store.readCell(at, options)).toMatchObject({ kind: 'cell' })
    }
  })

  it('holds exactly the declared NNZ over the full axes at reduced scale', () => {
    const handle = createFixture('F1', { nnz: 500 })
    expect(handle.rowCount).toBe(MAX_ROWS)
    expect(handle.colCount).toBe(MAX_COLS)
    expect(handle.nnz()).toBe(500)
  })

  it('declares exactly 100000 cells including corners with full column spread', () => {
    const handle = createFixture('F1')
    expect(handle.nnz()).toBe(100_000)
    const cols = new Set<number>()
    for (const ref of handle.enumerateHead(100_000)) cols.add(ref.col)
    expect(cols.size).toBeGreaterThan(8000)
  })

  it('is deterministic for the same seed', () => {
    expect(createFixture('F1', { nnz: 500 }).canonicalSamples())
      .toBe(createFixture('F1', { nnz: 500 }).canonicalSamples())
  })
})

describe('F2 tall table', () => {
  it('exposes header, body, totals and the max-row edge at reduced scale', () => {
    const handle = createFixture('F2', { rowCount: 1000 })
    expect(handle.cellAt(0, 0)?.value).toBe('id')
    expect(handle.cellAt(0, 3)?.value).toBe('name')
    const totals = handle.cellAt(999, 0)
    expect(totals?.value).toBe('total')
    expect(handle.cellAt(999, 1)?.formula).toBeDefined()
    expect(handle.cellAt(500, 4)).toBeNull()
    const err = handle.cellAt(505, 5)
    expect(err?.valueIsError).toBe(true)
    expect(handle.cellAt(10, 2)?.value).toEqual(expect.any(Boolean))
    expect(handle.cellAt(10, 1)?.styleIndex).toBe(STYLE_DATE)
  })

  it('matches brute-force NNZ with the closed form at reduced scale', () => {
    const handle = createFixture('F2', { rowCount: 1000 })
    let count = 0
    for (let row = 0; row < 1000; row++) for (let col = 0; col < 8; col++) {
      if (handle.cellAt(row, col) !== null) count++
    }
    expect(handle.nnz()).toBe(count)
  })

  it('counts degenerate row counts exactly', () => {
    expect(createFixture('F2', { rowCount: 0 }).nnz()).toBe(0)
    expect(createFixture('F2', { rowCount: 1 }).nnz()).toBe(8)
  })

  it('derives seeded values from (seed,row,col) while keeping structure fixed', () => {
    const a = createFixture('F2', { rowCount: 500, seed: 11 })
    const b = createFixture('F2', { rowCount: 500, seed: 12 })
    expect(a.canonicalSamples()).not.toBe(b.canonicalSamples())
    expect(a.cellAt(0, 0)?.value).toBe('id')
    expect(a.cellAt(499, 0)?.value).toBe('total')
    expect(a.cellAt(3, 4)).toBeNull()
    for (const id of ['F1', 'F3', 'F4', 'F5', 'F7'] as FixtureId[]) {
      expect(createFixture(id, { seed: 21 }).canonicalSamples())
        .not.toBe(createFixture(id, { seed: 22 }).canonicalSamples())
    }
    expect(cellHash(11, 5, 3)).toBe(cellHash(11, 5, 3))
    expect(cellHash(11, 5, 3)).not.toBe(cellHash(12, 5, 3))
  })
})

describe('F3 wide table', () => {
  it('covers the dense rectangle at reduced scale', () => {
    const handle = createFixture('F3', { rowCount: 4, colCount: 64 })
    expect(handle.nnz()).toBe(256)
    expect(handle.cellAt(0, 63)).not.toBeNull()
    expect(handle.cellAt(4, 0)).toBeNull()
  })

  it('keeps the sparse-wide variant sparse but XFD-addressable', () => {
    const handle = createFixture('F3', { rowCount: 256, variant: 'sparse-wide' })
    expect(handle.nnz()).toBeLessThan(256 * 512)
    expect(handle.nnz()).toBeGreaterThan(0)
    expect(handle.cellAt(0, MAX_COLS - 1)).not.toBeNull()
  })
})

describe('F4 string domain', () => {
  it('yields unique strings with accents, case, natural suffixes and lone surrogates', () => {
    const first = Array.from({ length: 5000 }, (_, i) => f4StringAt(i))
    expect(new Set(first).size).toBe(5000)
    expect(first.some((s) => /[\u0300-\u036f]/.test(s.normalize('NFD')))).toBe(true)
    const lower = first.findIndex((s) => s === s.toLowerCase() && /[a-z]/.test(s))
    expect(lower).toBeGreaterThanOrEqual(0)
    expect(first.some((s) => /item-1\d/.test(s))).toBe(true)
    expect(first.some((s) => /[\uD800-\uDFFF]/.test(s) && ![...s].some((ch) => ch.codePointAt(0)! > 0xffff))).toBe(true)
  })

  it('round-trips lone-surrogate code units exactly', () => {
    const lone = f4StringAt(3)
    expect([...lone].length).toBe(lone.length)
    expect(lone).toBe('\uDE00-lone-3')
  })

  it('declares one million unique strings', () => {
    expect(F4_STRING_COUNT).toBe(1_000_000)
    const handle = createFixture('F4')
    expect(handle.nnz()).toBe(1_000_000)
  })
})

describe('F5 structural metadata', () => {
  it('describes giant merges, both table shapes, spills and all drawing kinds', () => {
    const handle = createFixture('F5')
    const structure = handle.structure()
    expect(structure.merges.length).toBeGreaterThanOrEqual(1)
    expect(Math.max(...structure.merges.map((m) => m.r1 - m.r0))).toBeGreaterThan(100_000)
    expect(structure.tables).toContainEqual(expect.objectContaining({ headerRows: 0, totalsRows: 0 }))
    expect(structure.tables).toContainEqual(expect.objectContaining({ headerRows: 1, totalsRows: 1 }))
    expect(structure.spills.map((s) => s.kind).sort()).toEqual(['dynamic', 'fixed'])
    expect(structure.spills.every((s) => s.saved !== undefined || s.evaluated !== undefined)).toBe(true)
    expect(structure.drawings.map((d) => d.kind).sort()).toEqual(['absolute', 'oneCell', 'twoCell'])
  })

  it('answers structural queries only for intersecting records', async () => {
    const handle = createFixture('F5')
    const { store } = makeStore(handle)
    const options = readOptionsFor(store)
    const hit = await store.queryStructural({ r0: 0, r1: 10, c0: 0, c1: 10 }, options)
    expect(hit.merges.length).toBeGreaterThanOrEqual(1)
    expect(hit.completeness).toBe('complete')
    const miss = await store.queryStructural({ r0: 500_000, r1: 500_010, c0: 8000, c1: 8010 }, options)
    expect([...miss.merges, ...miss.tables, ...miss.spills, ...miss.drawings]).toHaveLength(0)
  })
})

describe('F6 full-axis synthetic provider', () => {
  it('reports every address readable including both max corners', async () => {
    const handle = createFixture('F6')
    const { store } = makeStore(handle)
    const options = readOptionsFor(store)
    expect(await store.readCell({ row: 0, col: 0 }, options)).toMatchObject({ kind: 'cell' })
    expect(await store.readCell({ row: MAX_ROWS - 1, col: MAX_COLS - 1 }, options)).toMatchObject({ kind: 'cell' })
    expect(await store.readCell({ row: 123456, col: 7890 }, options)).toMatchObject({ kind: 'cell' })
  })

  it('answers bounded windows without enumerating the domain', async () => {
    const inner = createFixture('F6')
    let visits = 0
    const handle: FixtureHandle = { ...inner, cellAt: (row, col) => { visits++; return inner.cellAt(row, col) } }
    const { store } = makeStore(handle)
    const options = readOptionsFor(store)
    visits = 0
    await store.readRows({ r0: 1000, r1: 1099, c0: 0, c1: 7 }, undefined, options)
    expect(visits).toBeLessThanOrEqual(100 * 8)
    visits = 0
    await store.queryOverview({
      groups: [{ groupId: 'g', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: MAX_ROWS - 1 }] }, cols: [{ c0: 0, c1: 7 }] }],
      maxGroups: 8,
    }, options)
    expect(visits).toBe(0)
  })
})

describe('F7 view stress descriptors', () => {
  it('provides eight sort keys, hidden intervals and multiple sheets', () => {
    const spec = f7Spec()
    expect(spec.sortCols).toHaveLength(8)
    expect(spec.hiddenIntervals.length).toBeGreaterThanOrEqual(2)
    expect(spec.sheetCount).toBeGreaterThan(1)
  })
})

describe('F8 compatibility parity', () => {
  it('distinguishes cached errors from error-looking text', () => {
    const cases = f8Cases()
    const byNote = new Map(cases.map((c) => [c.note, c]))
    expect(byNote.get('cached-error')?.cell.valueIsError).toBe(true)
    expect(byNote.get('error-looking-text')?.cell.valueIsError).toBe(false)
    expect(byNote.get('error-looking-text')?.cell.value).toBe('#DIV/0!')
  })

  it('marks cache-less formulas unavailable only under the cached policy', async () => {
    const engine = createFixture('F8').createStore({ valuePolicy: 'engine' })
    const cached = createFixture('F8').createStore({ valuePolicy: 'cached' })
    const target = f8Cases().find((c) => c.note === 'cache-less-formula')!.address
    expect(await engine.readCell(target, readOptionsFor(engine))).toMatchObject({ kind: 'cell' })
    expect(await cached.readCell(target, readOptionsFor(cached))).toMatchObject({ kind: 'unavailable' })
  })
})

describe('synthetic store contract semantics', () => {
  it('separates unknown, blank, cell and error reads', async () => {
    const handle = createFixture('F1', { nnz: 200 })
    const pending = { r0: 500, r1: 599, c0: 0, c1: 7 }
    const { store } = makeStore(handle, { pending: [pending] })
    const options = readOptionsFor(store)
    expect(await store.readCell({ row: 550, col: 3 }, options)).toMatchObject({ kind: 'unknown' })
    const blank = await store.readCell({ row: 700, col: 700 }, options)
    expect(blank).toEqual({ kind: 'blank' })
    expect(await store.readCell({ row: 0, col: 0 }, options)).toMatchObject({ kind: 'cell' })
  })

  it('rejects stale stamps and aborted reads', async () => {
    const { store } = makeStore(createFixture('F1', { nnz: 50 }))
    const stale = { ...store.stamp, valueGen: store.stamp.valueGen + 1 }
    const aborted = new AbortController()
    aborted.abort()
    await expect(store.readCell({ row: 0, col: 0 }, { signal: aborted.signal, stamp: store.stamp })).rejects.toThrow()
    expect(await store.readCell({ row: 0, col: 0 }, { signal: new AbortController().signal, stamp: stale })).toMatchObject({ kind: 'error', code: 'stale-source' })
    const page = await store.readRows({ r0: 0, r1: 5, c0: 0, c1: 0 }, undefined, { signal: new AbortController().signal, stamp: stale })
    expect(page.completeness).toBe('incomplete')
    expect(page.error?.code).toBe('stale-source')
  })

  it('reports missing ranges from ensureComplete while pending', async () => {
    const pending = { r0: 10, r1: 20, c0: 0, c1: 3 }
    const { store } = makeStore(createFixture('F1', { nnz: 50 }), { pending: [pending] })
    const options = readOptionsFor(store)
    const incomplete = await store.ensureComplete({ r0: 0, r1: 30, c0: 0, c1: 3 }, options)
    expect(incomplete.completeness).toBe('loading')
    expect(incomplete.missing).toHaveLength(1)
    const done = await store.ensureComplete({ r0: 0, r1: 5, c0: 0, c1: 3 }, options)
    expect(done.completeness).toBe('complete')
    expect(done.missing).toHaveLength(0)
  })

  it('refuses unbounded readRows ranges with an explicit resource limit', async () => {
    const { store } = makeStore(createFixture('F6'))
    const page = await store.readRows({ r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }, undefined, readOptionsFor(store))
    expect(page.completeness).toBe('incomplete')
    expect(page.error?.code).toBe('resource-limit')
  })

  it('returns analytic overview groups with exact disjoint membership', async () => {
    const { store } = makeStore(createFixture('F1', { nnz: 300 }))
    const options = readOptionsFor(store)
    const page = await store.queryOverview({
      groups: [
        { groupId: 'odd', rows: { kind: 'intervals', intervals: [{ r0: 1, r1: 1 }, { r0: 3, r1: 3 }] }, cols: [{ c0: 0, c1: 0 }] },
        { groupId: 'blank-span', rows: { kind: 'intervals', intervals: [{ r0: 900_000, r1: 900_010 }] }, cols: [{ c0: 0, c1: 0 }] },
      ],
      maxGroups: 2048,
    }, options)
    expect(page.completeness).toBe('complete')
    expect(page.items).toHaveLength(2)
    const odd = page.items.find((g) => g.groupId === 'odd')!
    if (odd.content === 'unknown') throw new Error('overview group must be resolved')
    expect(odd.countKind).toBe('exact')
    const span = page.items.find((g) => g.groupId === 'blank-span')!
    if (span.content === 'unknown') throw new Error('blank span must be analytic')
    expect(span.content).toBe('blank')
    expect(span.populatedCount).toBe(0)
  })
})

describe('canonical checksums', () => {
  it('hashes the known SHA-256 vector', () => {
    expect(sha256Hex('abc')).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
  })

  it('keeps isPopulated/styleOf in agreement with cellAt', () => {
    const grids: Array<{ handle: FixtureHandle; rows: number; cols: number }> = [
      { handle: createFixture('F2', { rowCount: 200 }), rows: 200, cols: 8 },
      { handle: createFixture('F3', { rowCount: 8, colCount: 32 }), rows: 8, cols: 32 },
      { handle: createFixture('F3', { rowCount: 8, colCount: 32, variant: 'sparse-wide' }), rows: 8, cols: 32 },
      { handle: createFixture('F5'), rows: 0, cols: 0 },
      { handle: createFixture('F7', { rowCount: 200 }), rows: 200, cols: 8 },
    ]
    for (const { handle, rows, cols } of grids) {
      for (let row = 0; row < rows; row++) for (let col = 0; col < cols; col++) {
        const cell = handle.cellAt(row, col)
        expect(handle.isPopulated(row, col)).toBe(cell !== null)
        if (cell !== null) expect(handle.styleOf(row, col)).toBe(cell.styleIndex)
      }
    }
    const f5 = createFixture('F5')
    for (const [row, col] of [[32, 0], [40, 3], [45, 3], [4, 2], [0, 0], [9, 9]] as Array<[number, number]>) {
      const cell = f5.cellAt(row, col)
      expect(f5.isPopulated(row, col)).toBe(cell !== null)
      if (cell !== null) expect(f5.styleOf(row, col)).toBe(cell.styleIndex)
    }
  })

  it('keeps canonical samples stable for every fixture', () => {
    for (const id of ['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8'] as FixtureId[]) {
      expect(createFixture(id).canonicalSamples()).toBe(createFixture(id).canonicalSamples())
    }
  })
})

describe('fixture manifest', () => {
  it('declares F1-F8 with files that match disk bytes and hashes', () => {
    const manifestPath = join(FIXTURE_DIR, 'manifest.json')
    expect(existsSync(manifestPath)).toBe(true)
    const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as {
      fixtures: Array<{ id: string; seed: number; files: Array<{ path: string; bytes: number; sha256: string }> }>
    }
    expect(manifest.fixtures.map((f) => f.id).sort()).toEqual(['F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8'])
    for (const fixture of manifest.fixtures) {
      expect(fixture.seed).toBe(FIXTURE_SEEDS[fixture.id as FixtureId])
      for (const file of fixture.files) {
        const disk = join(FIXTURE_DIR, file.path)
        expect(existsSync(disk)).toBe(true)
        const bytes = readFileSync(disk)
        expect(bytes.length).toBe(file.bytes)
        expect(createHash('sha256').update(bytes).digest('hex')).toBe(file.sha256)
      }
    }
    const total = manifest.fixtures.flatMap((f) => f.files).reduce((sum, f) => sum + f.bytes, 0)
    expect(total).toBeLessThan(200 * 1024)
  })

  it('recomputes every deterministic manifest field from the builders', async () => {
    const manifest = JSON.parse(readFileSync(join(FIXTURE_DIR, 'manifest.json'), 'utf8')) as {
      fixtures: Array<{
        id: FixtureId; seed: number; rowCount: number; colCount: number; nnz: number
        used: { r0: number; r1: number; c0: number; c1: number }; samplesSha256: string
        queryHashes: { readRowsHead: string; overviewHead: string }
        package: string | null; packageParams: FixtureParams | null; packageNnz: number | null
        files: Array<{ path: string }>
      }>
    }
    expect('measured' in manifest).toBe(false)
    for (const entry of manifest.fixtures) {
      const handle = createFixture(entry.id)
      expect(entry.seed).toBe(handle.seed)
      expect(entry.rowCount).toBe(handle.rowCount)
      expect(entry.colCount).toBe(handle.colCount)
      expect(entry.nnz).toBe(handle.nnz())
      expect(entry.used).toEqual(handle.used)
      expect(entry.samplesSha256).toBe(sha256Hex(handle.canonicalSamples()))
      expect(entry.queryHashes).toEqual(await headQueryHashes(handle))
      if (entry.package) {
        const probe = createFixture(entry.id, entry.packageParams ?? {})
        const expected = entry.id === 'F5'
          ? probe.scanCells({ r0: 0, r1: 4, c0: 0, c1: 3 }, 0, 1_000_000).length
          : probe.enumerateHead(1_000_000).length
        expect(entry.packageNnz).toBe(expected)
      }
    }
  }, 120_000)
})

describe('generated-package agreement', () => {
  it('matches parsed small-package cells against the synthetic store', async () => {
    const manifest = JSON.parse(readFileSync(join(FIXTURE_DIR, 'manifest.json'), 'utf8')) as {
      fixtures: Array<{ id: FixtureId; package: string | null; packageFixture: FixtureId | null; packageParams: FixtureParams | null }>
    }
    for (const entry of manifest.fixtures) {
      if (!entry.package || !entry.packageFixture) continue
      const bytes = readFileSync(join(FIXTURE_DIR, entry.package))
      const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(bytes)))
      const handle = createFixture(entry.packageFixture, entry.packageParams ?? {})
      const { store } = makeStore(handle)
      const options = readOptionsFor(store)
      const cells = doc.sheets[0].rows.flatMap((row) => row.cells)
      expect(cells.length).toBeGreaterThan(0)
      for (const parsed of cells) {
        const expected = handle.cellAt(parsed.row, parsed.col)
        if (expected?.formula !== undefined && !expected.hasCachedValue) {
          // Engine policy: the evaluator materializes an error cache on load
          // (spec §4.4 engine-vs-cached difference, pinned for S1.3).
          expect(parsed.formula).toBe(expected.formula)
          expect(parsed.value).toBe('#NAME?')
          expect(parsed.valueIsError).toBe(true)
          continue
        }
        const read = await store.readCell({ row: parsed.row, col: parsed.col }, options)
        expect(read).toMatchObject({ kind: 'cell' })
        if (read.kind !== 'cell') throw new Error('expected cell')
        assertParsedCellMatches(parsed, read.cell)
      }
    }
  }, 60_000)

  function assertParsedCellMatches(parsed: XlsxCell, cell: { value: unknown; formula?: string; styleIndex: number; hasCachedValue?: boolean; valueIsError?: boolean }) {
    expect(cell.value).toEqual(parsed.value)
    expect(cell.formula ?? undefined).toEqual(parsed.formula ?? undefined)
    expect(cell.styleIndex).toBe(parsed.styleIndex)
    expect(cell.hasCachedValue ?? false).toBe(parsed.hasCachedValue ?? false)
    expect(cell.valueIsError ?? false).toBe(parsed.valueIsError ?? false)
  }

  it('round-trips lone surrogates from committed packages losslessly', async () => {
    const bytes = readFileSync(join(FIXTURE_DIR, 'f8-small.xlsx'))
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(bytes)))
    const values = doc.sheets[0].rows.flatMap((row) => row.cells.map((c) => c.value))
    const lone = values.find((v) => v === '\uDE00-lone')
    expect(lone).toBe('\uDE00-lone')
  })

  it('is bidirectional: parsed count matches packageNnz and no store cell is missing', async () => {
    const manifest = JSON.parse(readFileSync(join(FIXTURE_DIR, 'manifest.json'), 'utf8')) as {
      fixtures: Array<{
        id: FixtureId; package: string | null; packageFixture: FixtureId | null
        packageParams: FixtureParams | null; packageNnz: number | null
        coverage: { r0: number; r1: number; c0: number; c1: number } | null
        extraPackages: Array<{ path: string; fixture: FixtureId; params: FixtureParams | null; nnz: number; coverage: { r0: number; r1: number; c0: number; c1: number } }>
      }>
    }
    const jobs: Array<{ path: string; fixture: FixtureId; params: FixtureParams | null; nnz: number; coverage: { r0: number; r1: number; c0: number; c1: number } }> = []
    for (const entry of manifest.fixtures) {
      if (entry.package && entry.packageFixture && entry.packageNnz !== null && entry.coverage) {
        jobs.push({ path: entry.package, fixture: entry.packageFixture, params: entry.packageParams, nnz: entry.packageNnz, coverage: entry.coverage })
      }
      for (const extra of entry.extraPackages ?? []) jobs.push(extra)
    }
    expect(jobs.length).toBeGreaterThan(0)
    for (const job of jobs) {
      const bytes = readFileSync(join(FIXTURE_DIR, job.path))
      const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(bytes)))
      const cells = doc.sheets[0].rows.flatMap((row) => row.cells)
      expect(cells.length).toBe(job.nnz)
      const parsedKeys = new Set(cells.map((c) => `${c.row}:${c.col}`))
      const handle = createFixture(job.fixture, job.params ?? {})
      let covered: Array<{ row: number; col: number }>
      try {
        covered = handle.enumerateHead(1_000_000)
          .filter((ref) => ref.row >= job.coverage.r0 && ref.row <= job.coverage.r1 && ref.col >= job.coverage.c0 && ref.col <= job.coverage.c1)
      } catch {
        const range = { r0: job.coverage.r0, r1: job.coverage.r1, c0: job.coverage.c0, c1: job.coverage.c1 }
        covered = handle.scanCells(range, 0, 1_000_000)
      }
      const missing = covered.filter((ref) => !parsedKeys.has(`${ref.row}:${ref.col}`))
      expect(missing).toEqual([])
    }
  }, 60_000)
})

describe('queryColumn blank runs (exact membership)', () => {
  async function columnEntries(handle: FixtureHandle, col: number, rows: { r0: number; r1: number } | number[]) {
    const { store } = makeStore(handle)
    const options = readOptionsFor(store)
    const set = Array.isArray(rows)
      ? { kind: 'list' as const, r0: rows[0], r1: rows[rows.length - 1], rows: new Uint32Array(rows) }
      : { kind: 'intervals' as const, intervals: [{ r0: rows.r0, r1: rows.r1 }] }
    const page = await store.queryColumn({ col, rows: set, limit: 100, cursor: undefined }, options)
    return page.items
  }

  it('resets blank runs at every non-blank cell (F5 col 0 rows 1..5)', async () => {
    const handle = createFixture('F5')
    const items = await columnEntries(handle, 0, { r0: 1, r1: 5 })
    expect(items).toEqual([
      { kind: 'blank', r0: 1, r1: 2 },
      { kind: 'cell', row: 3, cell: expect.objectContaining({ value: 30 + (FIXTURE_SEEDS.F5 % 5) }) },
      { kind: 'cell', row: 4, cell: expect.objectContaining({ value: 'head-0' }) },
      { kind: 'blank', r0: 5, r1: 5 },
    ])
  })

  it('never merges blanks across gaps in non-contiguous row sets', async () => {
    const handle = createFixture('F5')
    const items = await columnEntries(handle, 1, [1, 9])
    expect(items).toEqual([
      { kind: 'blank', r0: 1, r1: 1 },
      { kind: 'blank', r0: 9, r1: 9 },
    ])
  })
})

describe('bounded domains (used clamping, analytic summaries)', () => {
  it('resolves a full-axes F2 overview exactly without enumerating 2^34', async () => {
    const handle = createFixture('F2')
    const { store } = makeStore(handle)
    const options = readOptionsFor(store)
    const page = await store.queryOverview({
      groups: [{ groupId: 'all', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: MAX_ROWS - 1 }] }, cols: [{ c0: 0, c1: 7 }] }],
      maxGroups: 2048,
    }, options)
    expect(page.completeness).toBe('complete')
    const group = page.items[0]
    if (group.content === 'unknown') throw new Error('full F2 domain must resolve')
    expect(group.populatedCount).toBe(handle.nnz())
    expect(group.countKind).toBe('exact')
  })

  it('resolves full-domain F3/F4/F5 overviews exactly', async () => {
    for (const [id, expected] of [['F3', 4_194_304], ['F4', 1_000_000]] as Array<[FixtureId, number]>) {
      const handle = createFixture(id)
      const { store } = makeStore(handle)
      const options = readOptionsFor(store)
      const page = await store.queryOverview({
        groups: [{ groupId: 'all', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: MAX_ROWS - 1 }] }, cols: [{ c0: 0, c1: handle.colCount - 1 }] }],
        maxGroups: 2048,
      }, options)
      const group = page.items[0]
      if (group.content === 'unknown') throw new Error(`${id} must resolve`)
      expect(group.populatedCount).toBe(expected)
    }
    const f5 = createFixture('F5')
    const { store } = makeStore(f5)
    const page = await store.queryOverview({
      groups: [{ groupId: 'all', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: MAX_ROWS - 1 }] }, cols: [{ c0: 0, c1: MAX_COLS - 1 }] }],
      maxGroups: 2048,
    }, readOptionsFor(store))
    const group = page.items[0]
    if (group.content === 'unknown') throw new Error('F5 must resolve')
    expect(group.populatedCount).toBe(f5.nnz())
  })
})

describe('counting store reads (no vacuous wrappers)', () => {
  function counting(handle: FixtureHandle) {
    const visits = { cellAt: 0, isPopulated: 0, styleOf: 0 }
    const read: CellReader = {
      cellAt: (row, col) => { visits.cellAt++; return handle.cellAt(row, col) },
      isPopulated: (row, col) => { visits.isPopulated++; return handle.isPopulated(row, col) },
      styleOf: (row, col) => { visits.styleOf++; return handle.styleOf(row, col) },
    }
    return { read, visits }
  }

  it('visits a tight positive bound for F6 readRows windows', async () => {
    const handle = createFixture('F6')
    const { read, visits } = counting(handle)
    const store = handle.createStore({ read })
    await store.readRows({ r0: 1000, r1: 1099, c0: 0, c1: 7 }, undefined, readOptionsFor(store))
    expect(visits.cellAt).toBeGreaterThan(0)
    expect(visits.cellAt).toBeLessThanOrEqual(800)
  })

  it('answers F6 full-domain overview and structural queries with zero cell visits', async () => {
    const handle = createFixture('F6')
    const { read, visits } = counting(handle)
    const store = handle.createStore({ read })
    const options = readOptionsFor(store)
    await store.queryOverview({
      groups: [{ groupId: 'g', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: MAX_ROWS - 1 }] }, cols: [{ c0: 0, c1: 7 }] }],
      maxGroups: 8,
    }, options)
    await store.queryStructural({ r0: 0, r1: 10, c0: 0, c1: 10 }, options)
    expect(visits.cellAt + visits.isPopulated + visits.styleOf).toBe(0)
  })

  it('bounds F6 scan and column visits by the requested limit', async () => {
    const handle = createFixture('F6')
    const { read, visits } = counting(handle)
    const store = handle.createStore({ read })
    const options = readOptionsFor(store)
    for await (const page of store.scanPopulated({ range: { r0: 0, r1: 999, c0: 0, c1: 7 }, limit: 500, cursor: undefined }, options)) {
      expect(page.items).toHaveLength(500)
    }
    expect(visits.cellAt).toBeLessThanOrEqual(500)
    const column = await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 999 }] }, limit: 100, cursor: undefined }, options)
    expect(column.items.length).toBeLessThanOrEqual(100)
    expect(visits.cellAt).toBeLessThanOrEqual(600)
  })

  it('runs F5/F2/F3/F4 overview and scan within visit budgets', async () => {
    for (const id of ['F5', 'F2', 'F3', 'F4'] as FixtureId[]) {
      const handle = createFixture(id)
      const { read, visits } = counting(handle)
      const store = handle.createStore({ read })
      const options = readOptionsFor(store)
      const before = visits.cellAt + visits.isPopulated + visits.styleOf
      await store.queryOverview({
        groups: [{ groupId: 'g', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 999 }] }, cols: [{ c0: 0, c1: 7 }] }],
        maxGroups: 8,
      }, options)
      for await (const page of store.scanPopulated({ range: { r0: 0, r1: 99, c0: 0, c1: 7 }, limit: 200, cursor: undefined }, options)) {
        expect(page.items.length).toBeLessThanOrEqual(200)
      }
      const spent = visits.cellAt + visits.isPopulated + visits.styleOf - before
      expect(spent).toBeLessThan(100 * 8 + 1000)
    }
  })
})

describe('F4 collision families', () => {
  it('pairs twins that differ only by case, accent or trailing zero', () => {
    expect(f4StringAt(1)).toBe('café-0')
    expect(f4StringAt(9)).toBe('cafe-0')
    expect(f4StringAt(2)).toBe('Alpha-0')
    expect(f4StringAt(10)).toBe('alpha-0')
    expect(f4StringAt(4)).toBe('price-0-1.5')
    expect(f4StringAt(12)).toBe('price-0-1.50')
    expect(f4StringAt(11)).toBe('\uD800-lone-11')
  })

  it('classifies every index at full scale with balanced families', () => {
    const counts = new Map<string, number>()
    for (let index = 0; index < F4_STRING_COUNT; index++) {
      const family = f4Family(index)
      counts.set(family, (counts.get(family) ?? 0) + 1)
    }
    for (const family of ['natural', 'accent', 'case', 'lone', 'collide', 'long', 'astral', 'plain']) {
      expect(counts.get(family)).toBeGreaterThan(100_000)
    }
    expect(counts.get('blob')).toBe(1)
  })
})

describe('overview classification and request validation', () => {
  it('classifies by cells, not rows: one blank cell in a row means mixed', async () => {
    const { store } = makeStore(createFixture('F2', { rowCount: 100 }))
    const page = await store.queryOverview({
      groups: [{ groupId: 'r3', rows: { kind: 'intervals', intervals: [{ r0: 3, r1: 3 }] }, cols: [{ c0: 0, c1: 7 }] }],
      maxGroups: 8,
    }, readOptionsFor(store))
    const group = page.items[0]
    if (group.content === 'unknown') throw new Error('must resolve')
    expect(group.content).toBe('mixed')
    expect(group.populatedCount).toBe(7)
  })

  it('uses the corner cell (not the sampled value) for corner aggregates', async () => {
    const { store } = makeStore(createFixture('F1', { nnz: 200 }))
    const page = await store.queryOverview({
      groups: [{ groupId: 'a1', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 0 }] }],
      maxGroups: 8,
    }, readOptionsFor(store))
    const group = page.items[0]
    if (group.content === 'unknown') throw new Error('must resolve')
    expect(group).toMatchObject({ content: 'populated', style: { kind: 'uniform', styleIndex: 0 } })
  })

  it('reports uniform styles for uniform F6 regions including overrides', async () => {
    const { store } = makeStore(createFixture('F6'))
    const options = readOptionsFor(store)
    const uniform = await store.queryOverview({
      groups: [{ groupId: 'u', rows: { kind: 'intervals', intervals: [{ r0: 100, r1: 100 }] }, cols: [{ c0: 5, c1: 5 }] }],
      maxGroups: 8,
    }, options)
    const single = uniform.items[0]
    if (single.content === 'unknown') throw new Error('must resolve')
    expect(single.style).toEqual({ kind: 'uniform', styleIndex: 2 })
    const origin = await store.queryOverview({
      groups: [{ groupId: 'o', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 1 }] }],
      maxGroups: 8,
    }, options)
    const mixed = origin.items[0]
    if (mixed.content === 'unknown') throw new Error('must resolve')
    expect(mixed.style).toEqual({ kind: 'mixed' })
  })

  it('rejects out-of-range members, inverted intervals and bad maxGroups with Page.error', async () => {
    const { store } = makeStore(createFixture('F2', { rowCount: 100 }))
    const options = readOptionsFor(store)
    const cols = [{ c0: 0, c1: 0 }]
    const bad: OverviewQuery[] = [
      { groups: [{ groupId: 'o', rows: { kind: 'intervals', intervals: [{ r0: MAX_ROWS, r1: MAX_ROWS + 10 }] }, cols }], maxGroups: 8 },
      { groups: [{ groupId: 'i', rows: { kind: 'intervals', intervals: [{ r0: 10, r1: 5 }] }, cols }], maxGroups: 8 },
      { groups: [{ groupId: 'z', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 5 }] }, cols }], maxGroups: 0 },
      { groups: [{ groupId: 'h', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 5 }] }, cols }], maxGroups: 99999 },
    ]
    for (const request of bad) {
      const page = await store.queryOverview(request, options)
      expect(page.completeness).toBe('incomplete')
      expect(page.error).toBeDefined()
      expect(page.items).toHaveLength(0)
    }
  })
})

describe('five-state matrix across every query method', () => {
  function fiveStateStore() {
    const handle = createFixture('F8')
    const pending = [{ r0: 10, r1: 10, c0: 0, c1: 2 }]
    const errors = [{ range: { r0: 11, r1: 11, c0: 0, c1: 2 }, code: 'page-error' as const, retryable: true }]
    const store = handle.createStore({ valuePolicy: 'cached', pending, errors })
    return { handle, store, options: readOptionsFor(store) }
  }

  it('reads cell, blank, unknown, unavailable and error kinds', async () => {
    const { store, options } = fiveStateStore()
    expect(await store.readCell({ row: 0, col: 0 }, options)).toMatchObject({ kind: 'cell' })
    expect(await store.readCell({ row: 9, col: 9 }, options)).toEqual({ kind: 'blank' })
    expect(await store.readCell({ row: 10, col: 1 }, options)).toMatchObject({ kind: 'unknown' })
    expect(await store.readCell({ row: 1, col: 1 }, options)).toMatchObject({ kind: 'unavailable' })
    expect(await store.readCell({ row: 11, col: 1 }, options)).toMatchObject({ kind: 'error', code: 'page-error' })
  })

  it('propagates error and pending state through readRows pages', async () => {
    const { store, options } = fiveStateStore()
    const page = await store.readRows({ r0: 9, r1: 12, c0: 0, c1: 2 }, undefined, options)
    expect(page.completeness).toBe('incomplete')
    expect(page.error?.code).toBe('page-error')
    expect(page.items.every((item) => item.row < 10 || item.row > 11)).toBe(true)
  })

  it('propagates error and pending state through column pages', async () => {
    const { store, options } = fiveStateStore()
    const page = await store.queryColumn({
      col: 1,
      rows: { kind: 'intervals', intervals: [{ r0: 9, r1: 12 }] },
      limit: 10, cursor: undefined,
    }, options)
    expect(page.items.some((entry) => entry.kind === 'unknown')).toBe(true)
    expect(page.items.some((entry) => entry.kind === 'error')).toBe(true)
    expect(page.completeness).toBe('incomplete')
    expect(page.error?.code).toBe('page-error')
  })

  it('propagates error and pending state through scans, completion and structural queries', async () => {
    const { store, options } = fiveStateStore()
    for await (const page of store.scanPopulated({ range: { r0: 9, r1: 12, c0: 0, c1: 2 }, limit: 50, cursor: undefined }, options)) {
      expect(page.completeness).toBe('incomplete')
      expect(page.error?.code).toBe('page-error')
    }
    const completion = await store.ensureComplete({ r0: 9, r1: 12, c0: 0, c1: 2 }, options)
    expect(completion.completeness).toBe('incomplete')
    expect(completion.error?.code).toBe('page-error')
    expect(completion.missing).toHaveLength(1)
    const structural = await store.queryStructural({ r0: 9, r1: 12, c0: 0, c1: 2 }, options)
    expect(structural.completeness).toBe('incomplete')
    expect(structural.error?.code).toBe('page-error')
    expect(structural.missing).toHaveLength(1)
    const pendingOnly = await store.queryStructural({ r0: 9, r1: 10, c0: 0, c1: 2 }, options)
    expect(pendingOnly.completeness).toBe('loading')
    expect(pendingOnly.error).toBeUndefined()
    expect(pendingOnly.missing).toHaveLength(1)
  })

  it('reports overview unknown groups for pending domains and errors for error domains', async () => {
    const { store, options } = fiveStateStore()
    const page = await store.queryOverview({
      groups: [
        { groupId: 'p', rows: { kind: 'intervals', intervals: [{ r0: 10, r1: 10 }] }, cols: [{ c0: 0, c1: 0 }] },
        { groupId: 'ok', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 0 }] },
      ],
      maxGroups: 8,
    }, options)
    expect(page.items[0]).toMatchObject({ content: 'unknown' })
    expect(page.items[1].content).not.toBe('unknown')
    const errored = await store.queryOverview({
      groups: [{ groupId: 'e', rows: { kind: 'intervals', intervals: [{ r0: 11, r1: 11 }] }, cols: [{ c0: 0, c1: 0 }] }],
      maxGroups: 8,
    }, options)
    expect(errored.completeness).toBe('incomplete')
    expect(errored.error?.code).toBe('page-error')
  })

  it('pages readRows cursors through the whole range exactly once', async () => {
    const handle = createFixture('F2', { rowCount: 1000 })
    const { store } = makeStore(handle)
    const options = readOptionsFor(store)
    const seen: number[] = []
    let cursor: string | undefined
    for (let guard = 0; guard < 10; guard++) {
      const page = await store.readRows({ r0: 0, r1: 999, c0: 0, c1: 0 }, cursor, options)
      seen.push(...page.items.map((item) => item.row))
      cursor = page.nextCursor
      if (cursor === undefined) break
    }
    expect(seen).toEqual(Array.from({ length: 1000 }, (_, i) => i))
    expect(cursor).toBeUndefined()
  })
})

describe('F7 stress descriptors', () => {
  it('uses low-cardinality seeded sort keys with real ties', () => {
    const handle = createFixture('F7', { rowCount: 1000 })
    const distinct = new Set<number>()
    for (let row = 0; row < 1000; row++) distinct.add(handle.cellAt(row, 0)?.value as number)
    expect(distinct.size).toBeLessThanOrEqual(3)
    expect(distinct.size).toBeGreaterThan(0)
  })

  it('pins a max-row outlier and reaches the maximum row', () => {
    const handle = createFixture('F7')
    expect(handle.cellAt(MAX_ROWS - 1, 0)).not.toBeNull()
    expect(handle.used.r1).toBe(MAX_ROWS - 1)
  })

  it('builds max-dimension sparse sheets with distinct seeds', () => {
    const sheets = f7Sheets(3)
    expect(sheets).toHaveLength(3)
    for (const sheet of sheets) {
      expect(sheet.rowCount).toBe(MAX_ROWS)
      expect(sheet.colCount).toBe(MAX_COLS)
      expect(sheet.nnz()).toBeLessThanOrEqual(1000)
    }
    expect(sheets[0].canonicalSamples()).not.toBe(sheets[1].canonicalSamples())
  })

  it('describes row-interval hiding plus tiny sizes, zoom and slow jobs', () => {
    const spec = f7Spec()
    expect(spec.hiddenIntervals).toEqual([
      { r0: 5, r1: 5 }, { r0: 10, r1: 12 }, { r0: 100, r1: 199 },
    ])
    expect(spec.tinyRowHeights.length).toBeGreaterThan(0)
    expect(spec.zoomLevels).toEqual([0.25, 1, 2, 4])
    expect(spec.deviceScales).toEqual([1, 2])
    expect(spec.slowJobs[0]).toMatchObject({ kind: 'sort', rows: 100_000 })
  })
})

describe('structural identity and absolute placement', () => {
  it('keeps merge and table ids stable across query rects', async () => {
    const { store } = makeStore(createFixture('F5'))
    const options = readOptionsFor(store)
    const wide = await store.queryStructural({ r0: 0, r1: 100, c0: 0, c1: 10 }, options)
    const narrow = await store.queryStructural({ r0: 0, r1: 1, c0: 0, c1: 1 }, options)
    expect(narrow.merges).toHaveLength(1)
    expect(narrow.merges[0].id).toBe(wide.merges[0].id)
  })

  it('documents absolute drawings as viewport-level, excluded from rect queries', async () => {
    const handle = createFixture('F5')
    expect(handle.structure().drawings.some((d) => d.kind === 'absolute')).toBe(true)
    const { store } = makeStore(handle)
    const result = await store.queryStructural({ r0: 0, r1: 100, c0: 0, c1: 10 }, readOptionsFor(store))
    expect(result.drawings.some((d) => d.kind === 'absolute')).toBe(false)
    expect(result.drawings.map((d) => d.kind).sort()).toEqual(['oneCell', 'twoCell'])
  })

  it('resolves nonzero drawing extents from nominal geometry', async () => {
    const { store } = makeStore(createFixture('F5'))
    const result = await store.queryStructural({ r0: 70, r1: 74, c0: 2, c1: 4 }, readOptionsFor(store))
    expect(result.drawings).toHaveLength(1)
    expect(result.drawings[0].extent.width).toBeGreaterThan(0)
    expect(result.drawings[0].extent.height).toBeGreaterThan(0)
  })
})

describe('bound cursors and limits', () => {
  it('rejects cursors from a different query with invalid-cursor', async () => {
    const { store } = makeStore(createFixture('F2', { rowCount: 500 }))
    const options = readOptionsFor(store)
    const first = await store.readRows({ r0: 0, r1: 299, c0: 0, c1: 0 }, undefined, options)
    expect(first.nextCursor).toBeDefined()
    const foreign = await store.readRows({ r0: 200, r1: 299, c0: 0, c1: 0 }, first.nextCursor, options)
    expect(foreign.completeness).toBe('incomplete')
    expect(foreign.error?.code).toBe('invalid-cursor')
    const column = await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 9 }] }, limit: 5, cursor: first.nextCursor }, options)
    expect(column.error?.code).toBe('invalid-cursor')
  })

  it('rejects non-positive limits with Page.error instead of looping', async () => {
    const { store } = makeStore(createFixture('F2', { rowCount: 100 }))
    const options = readOptionsFor(store)
    for (const limit of [0, -1]) {
      const column = await store.queryColumn({ col: 0, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 9 }] }, limit, cursor: undefined }, options)
      expect(column.completeness).toBe('incomplete')
      expect(column.error).toBeDefined()
      expect(column.nextCursor).toBeUndefined()
      for await (const page of store.scanPopulated({ range: { r0: 0, r1: 9, c0: 0, c1: 0 }, limit, cursor: undefined }, options)) {
        expect(page.completeness).toBe('incomplete')
        expect(page.error).toBeDefined()
        expect(page.nextCursor).toBeUndefined()
      }
    }
  })
})

describe('string encoding (ST_Xstring)', () => {
  it('encodes lone surrogates, carriage returns and guarded escapes', () => {
    expect(encodeXlsxString('\uD800-lone')).toBe('_xD800_-lone')
    expect(encodeXlsxString('\uDE00-lone')).toBe('_xDE00_-lone')
    expect(encodeXlsxString('a\rb')).toBe('a_x000D_b')
    expect(encodeXlsxString('a\r\nb')).toBe('a_x000D_\nb')
    expect(encodeXlsxString('_x0041_')).toBe('_x005F_x0041_')
    expect(encodeXlsxString('😀')).toBe('😀')
    expect(encodeXlsxString('a&b<c>d')).toBe('a&amp;b&lt;c&gt;d')
  })
})

describe('corpus parity references', () => {
  it('points at existing committed oracle tests', () => {
    const refs = f8CorpusRefs()
    expect(refs.length).toBeGreaterThanOrEqual(2)
    for (const ref of refs) {
      expect(existsSync(join(__dirname, '..', ref))).toBe(true)
    }
  })
})
