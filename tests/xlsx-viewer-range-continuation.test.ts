/**
 * S10.1 reproduced gap (real Chromium, F2 at 1,000 and 100,010 rows, object and
 * worker-backed paths alike): a header sort with no selection resolved the
 * current region A1:H6 instead of the whole table, so the sort and the record
 * counts covered five body rows ("Visible body 5 / complete body 5").
 *
 * Cause: the SourceStore contract (spec §4.1) returns an AsyncIterable of
 * pages, and the shipped adapters (object, paged, remote) continue by
 * iteration without a `nextCursor`. Range detection read one page and took a
 * missing cursor as "exhausted", so a full strip page proved a row blank. The
 * synthetic fixture store pages by cursor, which is why the S7.1 suite passed.
 * These tests run the shipped object adapter over real parsed packages.
 */
import { describe, expect, it } from 'vitest'
import { admitSortRange, detectOperationRange, Scanner } from '../src/xlsx/view/rangeDetection'
import { StructuralIndex } from '../src/xlsx/view/structure'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { SourceStore } from '../src/xlsx/view/contracts'
import { colName, xlsxBytes } from './helpers/xlsx-package'

const FULL = { r0: 0, r1: 1048575, c0: 0, c1: 16383 }

/** Header row of text plus `rows` numeric body rows over `cols` columns. */
function table(rows: number, cols: number): string {
  const out: string[] = []
  for (let r = 0; r <= rows; r++) {
    const cells: string[] = []
    for (let c = 0; c < cols; c++) {
      const ref = `${colName(c)}${r + 1}`
      cells.push(r === 0 ? `<c r="${ref}" t="inlineStr"><is><t>h${c}</t></is></c>` : `<c r="${ref}"><v>${r * 10 + c}</v></c>`)
    }
    out.push(`<row r="${r + 1}">${cells.join('')}</row>`)
  }
  return out.join('')
}

async function objectStore(sheetData: string): Promise<{ store: SourceStore; structure: StructuralIndex }> {
  const doc = await parseXlsx(await OfficePackage.load(xlsxBytes({ sheetData })))
  const source = createObjectWorkbookSource(doc, 'doc-range')
  const store = source.store(source.sheets[0].sheetId)
  const options = { signal: new AbortController().signal, stamp: { ...store.stamp } }
  return { store, structure: new StructuralIndex(await store.queryStructural(FULL, options)) }
}
const request = (store: SourceStore, structure: StructuralIndex, row: number, col: number) => ({
  store, structure, selection: { active: { row, col }, anchor: { row, col }, ranges: [] },
  options: { signal: new AbortController().signal, stamp: { ...store.stamp } },
})

describe('range detection over iteration-continued store pages', () => {
  it('a header sort from A1 resolves the whole 8-column table, not the first strip page', async () => {
    const { store, structure } = await objectStore(table(999, 8))
    const resolved = await admitSortRange(request(store, structure, 0, 0))
    expect(resolved).toMatchObject({ kind: 'complete', range: { r0: 0, r1: 999, c0: 0, c1: 7, headerRows: 1 } })
  })

  it('a body active cell resolves the same table', async () => {
    const { store, structure } = await objectStore(table(999, 8))
    expect(await detectOperationRange(request(store, structure, 500, 3))).toMatchObject({ kind: 'complete', range: { r0: 0, r1: 999, c0: 0, c1: 7 } })
  })
})

describe('current-region growth cost', () => {
  it('a tall 8-column table resolves in a number of store pages bounded by its strips, not its rows', async () => {
    const { store, structure } = await objectStore(table(20_000, 8))
    let pages = 0
    const scan = store.scanPopulated.bind(store)
    store.scanPopulated = async function* (query, options) { for await (const page of scan(query, options)) { pages++; yield page } }
    const resolved = await detectOperationRange(request(store, structure, 0, 0))
    expect(resolved).toMatchObject({ kind: 'complete', range: { r0: 0, r1: 20_000, c0: 0, c1: 7 } })
    // Strips gallop to 1024 lines: about log2(1024) + 20,000/1024 strip pages
    // per side plus their read-ahead, never one probe per row.
    expect(pages).toBeLessThan(200)
  }, 60_000)
})

describe('detection scanner continuation scope', () => {
  it('a continuation token resumed with another range or limit is a RangeError, like any PopulatedCellSource cursor', async () => {
    const { store } = await objectStore(table(30, 8))
    const options = { signal: new AbortController().signal, stamp: { ...store.stamp } }
    const scanner = new Scanner(store, options, { tick: async () => {}, spend: () => false, yield: async () => {}, reserve: async () => {} } as never, Number.POSITIVE_INFINITY)
    const rect = { r0: 0, r1: 30, c0: 0, c1: 7 }
    const first = await scanner.page(rect, 8)
    expect(first.nextCursor).toBeDefined()
    await expect(scanner.page({ ...rect, r1: 20 }, 8, first.nextCursor)).rejects.toThrow(RangeError)
    const again = await scanner.page(rect, 8)
    await expect(scanner.page(rect, 4, again.nextCursor)).rejects.toThrow(RangeError)
    const third = await scanner.page(rect, 8)
    const next = await scanner.page(rect, 8, third.nextCursor)
    expect(next.items[0]?.row).toBe(1) // the same scope continues where the first page stopped
    scanner.close()
  })
})
