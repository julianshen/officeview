/** S0.2 deterministic fixture family (spec §12.5, plan S0.2). */
import { expect, it, describe } from 'vitest'
import { readFileSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxCell } from '../src/xlsx/types'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import type { SourceStore } from '../src/xlsx/view/contracts'
import {
  createPrng, hashSeed, createFixture, f4StringAt, F4_STRING_COUNT, f8Cases, f7Spec,
  sha256Hex, readOptionsFor, FIXTURE_SEEDS, STYLE_DATE,
  type FixtureId, type FixtureHandle, type FixtureParams, type SyntheticStoreOptions,
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

  it('holds the declared NNZ over the full axes at reduced scale', () => {
    const handle = createFixture('F1', { nnz: 500 })
    expect(handle.rowCount).toBe(MAX_ROWS)
    expect(handle.colCount).toBe(MAX_COLS)
    expect(handle.nnz()).toBeGreaterThanOrEqual(4)
    expect(handle.nnz()).toBeLessThanOrEqual(504)
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
    expect(handle.cellAt(10, 2)?.value).toBe(true)
    expect(handle.cellAt(11, 2)?.value).toBe(false)
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
  it('provides eight sort keys, hidden interleavings and size overrides', () => {
    const spec = f7Spec()
    expect(spec.sortCols).toHaveLength(8)
    expect(spec.hiddenIntervals.length).toBeGreaterThanOrEqual(2)
    expect(spec.sizeOverrides.length).toBeGreaterThanOrEqual(1)
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
      fixtures: Array<{ id: string; seed: number; nnz: number; files: Array<{ path: string; bytes: number; sha256: string }> }>
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
        expectNailCells(parsed, read.cell)
      }
    }
  }, 60_000)

  function expectNailCells(parsed: XlsxCell, cell: { value: unknown; formula?: string; styleIndex: number; hasCachedValue?: boolean; valueIsError?: boolean }) {
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
})
