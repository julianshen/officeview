// @vitest-environment node
/** SourcePageStore admits a record before reading it from storage (spec §12.4), for every reader. */
import { expect, it } from 'vitest'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { chartWorkbookBytes } from './helpers/xlsx-chart-fixture'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, type PageBackend } from '../src/xlsx/storage/sourcePages'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { MemoryPool } from '../src/xlsx/view/tileCache'

const col = (n: number) => { let out = ''; for (n++; n; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + (n - 1) % 26) + out; return out }
async function fixture(limitBytes: number) {
  const dir = await mkdtemp(join(tmpdir(), 'read-admission-')), file = (k: string) => join(dir, encodeURIComponent(k))
  const reads: Array<{ key: string; bytes: number; liveAtLeast: number }> = []
  let pool: MemoryPool | undefined
  // eslint-disable-next-line prefer-const
  let store: SourcePageStore
  const backend: PageBackend = {
    kind: 'opfs', put: async (k, b) => { await writeFile(file(k), b) }, delete: async k => { await rm(file(k), { force: true }) }, removeAll: async () => { await rm(dir, { recursive: true, force: true }) },
    get: async k => {
      try {
        const bytes = new Uint8Array(await readFile(file(k)))
        // The incoming buffer counts once: the pool total already holds its reservation.
        if (pool) reads.push({ key: k, bytes: bytes.byteLength, liveAtLeast: pool.stats().totalBytes + Math.max(0, bytes.byteLength - (store.readingBytes ?? 0)) })
        return bytes
      } catch { return undefined }
    },
  }
  store = new SourcePageStore({ docId: 'read-admission', backend })
  // 12 one-row pages of 2,000 numbers each (~48 KB per page on disk).
  const cells = Array.from({ length: 12 }, (_, r) => Array.from({ length: 2000 }, (_, c) => ({ ref: `${col(c)}${r + 1}`, v: r * 2000 + c }))).flat()
  const bytes = await chartWorkbookBytes('<chartSpace><chart><plotArea><barChart><ser><val><numLit><pt idx="0"><v>1</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>', { cells })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'read-admission', profile: 'D', store, pageRows: 1 })
  expect((await session.done).kind).toBe('complete')
  const source = createPagedWorkbookSource(session), adapter = source.store('1') as any
  pool = new MemoryPool({ limitBytes }); source.bindMemory!(pool)
  const read = (row: number) => adapter.readCell({ row, col: 1999 }, { stamp: adapter.stamp, signal: new AbortController().signal })
  return { pool, store, reads, read, done: async () => { await source.release(); pool!.close(); await rm(dir, { recursive: true, force: true }) } }
}

it('cell reads make room and reserve each page before it is read, staying within the pool', { timeout: 30_000 }, async () => {
  // Two pages fit beside the index; a third must evict before it is read.
  const h = await fixture(120_000)
  try {
    for (let row = 0; row < 12; row++) expect(await h.read(row)).toMatchObject({ kind: 'cell', cell: { value: row * 2000 + 1999 } })
    expect(h.reads.length).toBeGreaterThanOrEqual(12)
    expect(h.reads.every(r => r.liveAtLeast <= h.pool.limitBytes)).toBe(true)
    expect(h.store.readingBytes).toBe(0)
  } finally { await h.done() }
})

it('a page that cannot be admitted is refused with a typed resource result, never read', { timeout: 30_000 }, async () => {
  const h = await fixture(36_000)
  try {
    expect(await h.read(0)).toMatchObject({ kind: 'error', code: 'resource-limit' })
    expect(h.reads).toEqual([])
    expect(h.store.readingBytes).toBe(0)
  } finally { await h.done() }
})
