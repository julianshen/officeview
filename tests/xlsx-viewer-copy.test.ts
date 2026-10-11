import { describe, expect, it, vi } from 'vitest'
import { serializeSelection, CLIPBOARD_BYTES, copySessionBudget, type CopyInput } from '../src/xlsx/view/copy'
import { createIdentityOrder, createRowOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { cellId, MAX_ROWS, MAX_COLS } from '../src/xlsx/view/coordinates'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { workbookOf } from './helpers/xlsx-find'
import type { CellRead, SingleCellPayload } from '../src/xlsx/view/contracts'

const part = (r0: number, r1: number, c0 = 0, c1 = 0) => ({ rows: { kind: 'intervals' as const, intervals: [{ r0, r1 }] }, cols: [{ c0, c1 }] })
function input(read?: (row: number, col: number) => CellRead): CopyInput {
  const doc = workbookOf([{ sheetId: 'A', cells: [] }])
  const store = createObjectWorkbookSource(doc, 'copy').store('A')
  if (read) store.readCell = vi.fn(async at => read(at.row, at.col))
  const owner = { ...store.stamp, fontGen: 0, viewGen: 0 }
  return { store, owner, currentOwner: () => owner, selection: { active: null, anchor: null, ranges: [part(0, 2)] }, order: createIdentityOrder(), visibility: visibilityFromHidden(createIntervalSet([])), columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS), allowed: () => true, signal: new AbortController().signal }
}
const value = (row: number, col: number, text: string): CellRead => ({ kind: 'cell', cell: { row, col, id: cellId(row, col), value: text, styleIndex: 0 } })
const textOf = async (f: CopyInput) => { const out = await serializeSelection(f); expect(out.kind).toBe('ready'); return out.kind === 'ready' ? out.text : '' }

describe('displayed selection copy', () => {
  it('preserves disjoint source membership after a later sort and skips hidden rows/columns', async () => {
    const f = input((r, c) => value(r, c, `${r}:${c}`))
    f.selection.ranges = [{ rows: { kind: 'list', r0: 0, r1: 3, rows: Uint32Array.of(0, 1, 3) }, cols: [{ c0: 0, c1: 2 }] }]
    f.order = createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: Uint32Array.of(3, 2, 0, 1) }])
    f.visibility = visibilityFromHidden(createIntervalSet([{ r0: 1, r1: 1 }]))
    f.columns = visibilityFromHidden(createIntervalSet([{ r0: 1, r1: 1 }]), MAX_COLS)
    expect(await textOf(f)).toBe('3:0\t3:2\n0:0\t0:2\n1:0\t1:2')
  })
  it('dedupes overlapping ranges and orders disjoint blocks by display top-left', async () => {
    const f = input((r, c) => value(r, c, `${r}:${c}`))
    f.selection.ranges = [part(4, 4, 3, 3), part(0, 1, 0, 1), part(0, 1, 0, 1)]
    expect(await textOf(f)).toBe('0:0\t0:1\n1:0\t1:1\n\n4:3')
  })
  it('copies far corners, formatted values, literal errors and quoted TSV without repairing surrogates', async () => {
    const f = input((r, c) => value(r, c, 'a\tb\n"\ud800'))
    f.selection.ranges = [part(MAX_ROWS - 1, MAX_ROWS - 1, MAX_COLS - 1, MAX_COLS - 1)]
    expect(await textOf(f)).toBe('"a\tb\n""\ud800"')
    f.format = () => '#N/A'
    expect(await textOf(f)).toBe('#N/A')
    f.format = () => '$1.00'
    expect(await textOf(f)).toBe('$1.00')
  })
  it('refuses 5 MiB ASCII atomically and counts lone surrogates as raw UTF-16', async () => {
    const f = input((r, c) => value(r, c, 'x'.repeat(5 * 1024 * 1024)))
    f.selection.ranges = [part(0, 0)]
    expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
    f.store.readCell = async at => value(at.row, at.col, '\ud800'.repeat(CLIPBOARD_BYTES / 2))
    const out = await serializeSelection(f)
    expect(out).toMatchObject({ kind: 'ready', bytes: CLIPBOARD_BYTES })
  })
  it('rejects the fully visible full sheet from separator counts before any source reads', async () => {
    const f = input(() => ({ kind: 'blank' }))
    f.selection.ranges = [part(0, MAX_ROWS - 1, 0, MAX_COLS - 1)]
    f.order = { ...f.order }; const toSource = vi.spyOn(f.order, 'displayToSource')
    expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
    expect(f.store.readCell).not.toHaveBeenCalled()
    expect(toSource).not.toHaveBeenCalled()
  })
  it('accepts full logical selection when only one visible row and column remain', async () => {
    const f = input((r, c) => value(r, c, 'corner'))
    f.selection.ranges = [part(0, MAX_ROWS - 1, 0, MAX_COLS - 1)]
    f.visibility = visibilityFromHidden(createIntervalSet([{ r0: 0, r1: MAX_ROWS - 2 }]))
    f.columns = visibilityFromHidden(createIntervalSet([{ r0: 0, r1: MAX_COLS - 2 }]), MAX_COLS)
    expect(await textOf(f)).toBe('corner')
    expect(f.store.readCell).toHaveBeenCalledTimes(1)
  })
  it('accepts one million short rows with serial value reads and no row-count cap', async () => {
    const f = input((r, c) => value(r, c, 'x'))
    f.selection.ranges = [part(0, 999999)]
    let active = 0, peak = 0, reads = 0
    f.store.readCell = async at => { active++; peak = Math.max(peak, active); reads++; await Promise.resolve(); active--; return value(at.row, at.col, 'x') }
    const out = await serializeSelection(f)
    expect(out).toMatchObject({ kind: 'ready', bytes: 3999998 })
    expect(reads).toBe(1000000)
    expect(peak).toBe(1)
  }, 30000)
  it('keeps unknown values pending, then resumes; cancellation/stale ownership prevent publication', async () => {
    const f = input((r, c) => value(r, c, 'loaded'))
    f.selection.ranges = [part(0, 0)]
    let loaded = false, complete!: () => void
    f.store.readCell = async at => loaded ? value(at.row, at.col, 'loaded') : { kind: 'unknown', pageId: 'p' }
    f.store.ensureComplete = async range => { await new Promise<void>(resolve => { complete = resolve }); loaded = true; return { stamp: f.store.stamp, range, completeness: 'complete', missing: [] } }
    const pending = vi.fn(); f.onPending = pending
    const promise = serializeSelection(f)
    await vi.waitFor(() => expect(pending).toHaveBeenCalled())
    complete(); expect(await promise).toMatchObject({ kind: 'ready', text: 'loaded' })
    const controller = new AbortController(); f.signal = controller.signal; controller.abort()
    expect(await serializeSelection(f)).toMatchObject({ kind: 'cancelled' })
    f.signal = new AbortController().signal; f.currentOwner = () => ({ ...f.owner, viewGen: 1 })
    expect(await serializeSelection(f)).toMatchObject({ kind: 'stale' })
  })
  it('refuses unavailable formulas and typed source errors rather than copying blank', async () => {
    const f = input(() => ({ kind: 'unavailable', cell: (value(0, 0, '') as { kind: 'cell'; cell: SingleCellPayload }).cell, reason: 'formula-value-unavailable' }))
    f.selection.ranges = [part(0, 0)]
    expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'formula-value-unavailable' })
    f.store.readCell = async () => ({ kind: 'error', code: 'storage-error', retryable: false })
    expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'storage-error' })
  })
})

it('programmatic grid copy writes once after complete admission and never writes denied or oversized data', async () => {
  const { copySelection } = await import('../src/xlsx/view/copy')
  const write = vi.fn(async () => {})
  vi.stubGlobal('navigator', { clipboard: { writeText: write } })
  try {
    const f = input((r, c) => value(r, c, 'x'))
    f.allowed = () => false
    expect(await copySelection(f)).toMatchObject({ kind: 'denied' }); expect(write).not.toHaveBeenCalled()
    f.allowed = () => true
    expect(await copySelection(f)).toMatchObject({ kind: 'copied' }); expect(write).toHaveBeenCalledExactlyOnceWith('x\nx\nx')
    f.selection.ranges = [part(0, MAX_ROWS - 1, 0, MAX_COLS - 1)]
    expect(await copySelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' }); expect(write).toHaveBeenCalledTimes(1)
  } finally { vi.unstubAllGlobals() }
})

it('rejects unknown cells after the source confirmed complete coverage instead of retrying indefinitely', async () => {
  const f = input(); f.selection.ranges = [part(0, 0)]
  let reads = 0
  f.store.readCell = async () => ++reads <= 2 ? { kind: 'unknown', pageId: 'p' } : { kind: 'blank' }
  f.store.ensureComplete = async range => ({ range, stamp: f.store.stamp, completeness: 'complete', missing: [] })
  expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'page-error' })
  expect(reads).toBe(2)
})

describe('fragmented selections within the clipboard budget', () => {
  it.each([1000, 4096])('copies %i disjoint singleton parts', async count => {
    const f = input((r, c) => value(r, c, 'x'))
    f.selection.ranges = Array.from({ length: count }, (_, i) => part(i * 2, i * 2))
    expect(await textOf(f)).toBe(Array(count).fill('x').join('\n\n'))
    expect(f.store.readCell).toHaveBeenCalledTimes(count)
  })
  it.each(['list', 'bitset'] as const)('copies 4096 fragmented rows from a %s', async kind => {
    const f = input((r, c) => value(r, c, 'x'))
    f.selection.ranges = [{ cols: [{ c0: 0, c1: 0 }], rows: kind === 'list'
      ? { kind, r0: 0, r1: 8191, rows: Uint32Array.from({ length: 4096 }, (_, i) => i * 2) }
      : { kind, space: 'source', r0: 0, r1: 8191, words: new Uint32Array(256).fill(0x55555555), count: 4096 } }]
    expect(await textOf(f)).toBe(Array(4096).fill('x').join('\n\n'))
    expect(f.store.readCell).toHaveBeenCalledTimes(4096)
  })
  it('copies 4096 parts whose row sets each contain disjoint members', async () => {
    const f = input((r, c) => value(r, c, 'x'))
    f.selection.ranges = Array.from({ length: 4096 }, (_, i) => ({ cols: [{ c0: 0, c1: 0 }], rows: { kind: 'list' as const, r0: i * 4, r1: i * 4 + 2, rows: Uint32Array.of(i * 4, i * 4 + 2) } }))
    expect(await textOf(f)).toBe(Array(8192).fill('x').join('\n\n'))
    expect(f.store.readCell).toHaveBeenCalledTimes(8192)
  })
})


it('scales copy output/projection allowance with the S4.3 main-thread profile', () => {
  expect(copySessionBudget('D')).toBe(24 * 1024 * 1024)
  expect(copySessionBudget('M')).toBe(12 * 1024 * 1024)
})
