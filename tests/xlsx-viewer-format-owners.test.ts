import { expect, it, vi } from 'vitest'
import { cellDisplayText } from '../src/xlsx/render'
import { withFormatWorkspace } from '../src/xlsx/view/formatWorkspace'
import { WorkbookFind, FindError } from '../src/xlsx/view/find'
import { drain, findInputOf, workbookOf } from './helpers/xlsx-find'
import type { XlsxSheet } from '../src/xlsx/types'
import { createWorkerFilterEvaluator } from '../src/worker/viewerFilterEvaluator'
import { DisplayTextLaneCache } from '../src/worker/displayTextLane'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { DocumentSession } from '../src/worker/viewerSession'
import * as display from '../src/xlsx/render'
import { formatColor } from '../src/xlsx/cellFormat'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { mainFormatBudget } from '../src/components/xlsx/mainFormatBudget'
import { formatCellPresentation, SemanticPresentationCache } from '../src/components/xlsx/formatPresentation'
import { buildSheetGeometry, emptyWork, MergeQueryCache, planViewport } from '../src/xlsx/view/viewport'

const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
function budget(limit = Infinity) {
  let held = 0, peak = 0
  return {
    reserve(bytes: number) { if (held + bytes > limit) throw new FindError('resource-limit', 'Formatting memory refused'); held += bytes; peak = Math.max(peak, held) },
    release(bytes: number) { held -= bytes },
    get held() { return held }, get peak() { return peak },
  }
}

it('keeps format workspace through its consumer and drains on success/refusal', () => {
  const owner = budget()
  withFormatWorkspace(owner, allocation => {
    const text = cellDisplayText(sheet, { value: 3, style: { numFmtId: 164, formatCode: '[Red]0.00' } }, allocation)
    expect(text).toBe('3.00')
    expect(owner.held).toBeGreaterThan(0)
  })
  expect(owner.held).toBe(0)
  const tiny = budget(128)
  expect(() => withFormatWorkspace(tiny, allocation => cellDisplayText(sheet, { value: 'x'.repeat(10000), style: { numFmtId: 164, formatCode: '0;0;0;"v"' + '@'.repeat(32) } }, allocation))).toThrow('Formatting memory refused')
  expect(tiny.held).toBe(0)
})

it('Find forwards its operation observer and releases nonmatching builders', async () => {
  const doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 3 }] }])
  const owner = budget()
  const find = new WorkbookFind(findInputOf(doc).input, { text: 'missing', caseSensitive: false }, {
    formatBudget: owner,
    textOf: (cell, _id, allocation) => {
      expect(allocation).toBeDefined()
      return cellDisplayText(sheet, { value: cell.value, style: { numFmtId: 164, formatCode: '[Red]0.00' } }, allocation)
    },
  })
  expect(await drain(find)).toHaveLength(0)
  expect(owner.held).toBe(0)
  expect(owner.peak).toBeGreaterThan(0)
})

it('Find refuses repeated text placeholders under its owner budget before publishing', async () => {
  const doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'x'.repeat(10000) }] }])
  const owner = budget(50000)
  const find = new WorkbookFind(findInputOf(doc).input, { text: 'x', caseSensitive: false }, {
    formatBudget: owner,
    textOf: (cell, _id, allocation) => cellDisplayText(sheet, { value: cell.value, style: { numFmtId: 164, formatCode: '0;0;0;"v"' + '@'.repeat(32) } }, allocation),
  })
  await expect(find.nextPage()).rejects.toMatchObject({ code: 'resource-limit' })
  expect(owner.held).toBe(0)
  expect(owner.peak).toBeLessThanOrEqual(50000)
})

it('worker filter builders use the lane owner and refuse huge unused sections', async () => {
  const code = '[>0]0;[<=0]"' + 'q'.repeat(32768) + '"0'
  const style = { numFmtId: 164, formatCode: code }
  const source = createObjectWorkbookSource({ sheets: [{ ...sheet, rows: [
    { index: 0, cells: [{ ref: 'A1', row: 0, col: 0, value: 3, styleIndex: 0, style }] },
    { index: 1, cells: [{ ref: 'A2', row: 1, col: 0, value: 'text', styleIndex: 0 }] },
  ] }] }, 'd')
  const spy = vi.spyOn(display, 'cellDisplayText')
  const lanes = new DisplayTextLaneCache({ limitBytes: 10000 })
  try {
    const evaluate = createWorkerFilterEvaluator(() => new Map(), { dateSystem: '1900', styles: [style] }, 'en-US', { lanes })
    const result = await evaluate({ store: source.store('1'), spec: { range: { r0: 0, r1: 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, columns: new Map([[0, { kind: 'text', op: 'contains', first: '3', wildcards: false }]]) }, nowMs: 0, signal: new AbortController().signal })
    expect(spy.mock.calls.length).toBeGreaterThan(0)
    expect(spy.mock.calls.every(call => !!call[2])).toBe(true)
    expect(result).toMatchObject({ kind: 'refused', refusal: { reason: 'resource-limit' } })
    expect(lanes.bytes).toBe(0)
  } finally { spy.mockRestore(); lanes.close(); await source.release() }
})

it('colour parsing admits four Section records and slots even for empty bodies', () => {
  let held = 0, sections = 0, observed = 0
  const push = Array.prototype.push
  Array.prototype.push = function(this: unknown[], ...items: unknown[]) {
    for (const item of items) if (item && typeof item === 'object' && 'body' in item && 'unsupported' in item) { sections++; observed = held }
    return push.apply(this, items)
  }
  try {
    formatColor(3, '[Red];;;', { allocate: n => { held += n }, release: n => { held -= n } })
    expect(sections).toBe(4)
    expect(observed).toBeGreaterThanOrEqual(sections * 72)
  } finally { Array.prototype.push = push }
})

it('semantic summaries remain bounded, refuse safely and preserve the source value', () => {
  const value = 'x'.repeat(10000), cell = { ref: 'A1', row: 0, col: 0, styleIndex: 0, value, style: { numFmtId: 164, formatCode: '0;0;0;"v"' + '@'.repeat(32) } }
  const owner = budget()
  const presentation = formatCellPresentation(sheet, cell, owner)
  expect(presentation.text).toHaveLength(256)
  expect(presentation.text.startsWith('vxxx')).toBe(true)
  expect(owner.held).toBe(0)
  const tiny = budget(50000)
  expect(formatCellPresentation(sheet, cell, tiny)).toMatchObject({ state: 'error', text: expect.stringContaining('formatting memory budget') })
  expect(tiny.held).toBe(0)
  expect(cell.value).toBe(value)
})

it('retained diagnostic and Find hit strings transfer to the document lifetime', async () => {
  const registered: { target: object; token: string }[] = []
  const ledger = new MainMemoryLedger({ register: (target, token) => registered.push({ target, token }) })
  const owner = mainFormatBudget(ledger, 1000000)
  const target: XlsxSheet = { ...sheet }
  withFormatWorkspace(owner, allocation => cellDisplayText(target, { value: 3, style: { numFmtId: 164, formatCode: '[Purple]0' } }, allocation))
  expect(target.diagnostics?.length).toBeGreaterThan(0)
  expect(ledger.snapshot().byPool.layout).toBe(0)
  expect(ledger.snapshot().byPool.document).toBeGreaterThan(0)
  const diagnosticBytes = ledger.bytes
  const doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'value' }] }])
  const hits = await drain(new WorkbookFind(findInputOf(doc).input, { text: 'value', caseSensitive: false }, { formatBudget: owner }))
  expect(hits).toHaveLength(1)
  expect(ledger.bytes).toBeGreaterThan(diagnosticBytes)
  expect(ledger.snapshot().byPool.layout).toBe(0)
  for (const item of registered.filter(item => item.target !== target)) ledger.collected(item.token)
  expect(ledger.bytes).toBe(diagnosticBytes)
  for (const item of registered.filter(item => item.target === target)) ledger.collected(item.token)
  expect(ledger.bytes).toBe(0)
})

it('semantic cache reuses admitted presentations on an unrelated rerender', () => {
  const owner = budget(), cache = new SemanticPresentationCache(sheet, owner)
  const cell = { ref: 'A1', row: 0, col: 0, styleIndex: 0, value: 3, style: { numFmtId: 164, formatCode: '[Red]0.00' } }
  const spy = vi.spyOn(display, 'cellDisplayText')
  try {
    expect(cache.get(cell).text).toBe('3.00')
    const peak = owner.peak
    expect(cache.get(cell).text).toBe('3.00')
    expect(owner.peak).toBe(peak)
    expect(spy).toHaveBeenCalledTimes(1)
    cell.value = 4
    expect(cache.get(cell).text).toBe('4.00')
    cell.style.formatCode = '"v"0'
    expect(cache.get(cell).text).toBe('v4')
    expect(spy).toHaveBeenCalledTimes(3)
  } finally { spy.mockRestore() }
})

it('Find cancellation drains the active formatting workspace', async () => {
  const doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 3 }] }])
  const owner = budget(), abort = new AbortController()
  const find = new WorkbookFind(findInputOf(doc).input, { text: '3', caseSensitive: false }, {
    signal: abort.signal, formatBudget: owner,
    textOf: (cell, _id, allocation) => {
      const text = cellDisplayText(sheet, { value: cell.value, style: { numFmtId: 164, formatCode: '[Red]0.00' } }, allocation)
      abort.abort()
      return text
    },
  })
  await expect(find.nextPage()).rejects.toMatchObject({ code: 'cancelled' })
  expect(owner.held).toBe(0)
})

it('worker resident Find forwards the observer and keeps source diagnostics until session release', async () => {
  const target: XlsxSheet = { ...sheet, sheetId: '1', rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, value: 3, styleIndex: 0, style: { numFmtId: 164, formatCode: '[Purple]0' } }] }] }
  const session = new DocumentSession('d', createObjectWorkbookSource({ sheets: [target] }, 'd'))
  const spy = vi.spyOn(display, 'cellDisplayText')
  try {
    const before = session.memoryStats().transientBytes
    const result = await session.search(session.workbookOwner(), { text: '3', caseSensitive: false }, undefined, new AbortController().signal)
    expect(result).toMatchObject({ complete: true, hits: [{ text: '3' }] })
    expect(spy.mock.calls.length).toBeGreaterThan(0)
    expect(spy.mock.calls.every(call => !!call[2])).toBe(true)
    expect(target.diagnostics?.length).toBe(1)
    expect(session.memoryStats().transientBytes).toBeGreaterThan(before)
  } finally { spy.mockRestore(); await session.release() }
  expect(session.memoryStats().transientBytes).toBe(0)
})

it('a globally empty merge index shares its empty answer without growing per-rectangle caches', () => {
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('fixture geometry must be ready')
  const cache = new MergeQueryCache(), work = emptyWork()
  const first = cache.query(geometry, { r0: 0, r1: 1, c0: 0, c1: 1 }, 100, work)
  const second = cache.query(geometry, { r0: 2, r1: 3, c0: 2, c1: 3 }, 100, work)
  expect(first.items).toEqual([])
  expect(second).toBe(first)
  expect(cache.bytes).toBe(0)
  expect(work.mergeQueries).toBe(0)
})

it('merge pages are refused before the index builder and retained pages transfer to the merge pool', () => {
  const geometry = buildSheetGeometry({ ...sheet, merges: ['A1000:B1000'], mergeRanges: [{ minRow: 999, maxRow: 999, minCol: 0, maxCol: 1 }] })
  if (geometry.kind !== 'ready') throw new Error('fixture geometry must be ready')
  const ledger = new MainMemoryLedger({ register: () => {} })
  const tiny = new MergeQueryCache(64, mainFormatBudget(ledger, 128, 'merges'))
  const spy = vi.spyOn(geometry.merges, 'query')
  try {
    expect(tiny.query(geometry, { r0: 0, r1: 1, c0: 0, c1: 1 }, 100, emptyWork())).toMatchObject({ memoryRefused: true })
    expect(spy).not.toHaveBeenCalled()
    expect(ledger.bytes).toBe(0)
    const cache = new MergeQueryCache(64, mainFormatBudget(ledger, 100000, 'merges'))
    expect(cache.query(geometry, { r0: 999, r1: 999, c0: 0, c1: 1 }, 100, emptyWork()).items).toHaveLength(1)
    expect(ledger.snapshot().byPool.merges).toBeGreaterThan(0)
    expect(ledger.snapshot().byPool.layout).toBe(0)
  } finally { spy.mockRestore() }
})

it('merge warm hits reserve nothing and eviction keeps a held page alias charged', () => {
  const geometry = buildSheetGeometry({ ...sheet, merges: ['A1:B1'], mergeRanges: [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 1 }] })
  if (geometry.kind !== 'ready') throw new Error('fixture geometry must be ready')
  const ledger = new MainMemoryLedger({ register: () => {} }), owner = mainFormatBudget(ledger, 100000, 'merges')
  const reserve = vi.spyOn(owner, 'reserve'), cache = new MergeQueryCache(1, owner)
  const rect = { r0: 0, r1: 1, c0: 0, c1: 1 }, work = emptyWork()
  const first = cache.query(geometry, rect, 100, work), alias = first.items
  const admitted = reserve.mock.calls.length, bytes = ledger.bytes
  expect(cache.query(geometry, rect, 100, work)).toBe(first)
  expect(reserve.mock.calls.length).toBe(admitted)
  cache.query(geometry, { ...rect, r0: 2, r1: 3 }, 100, work)
  expect(alias).toHaveLength(1)
  expect(ledger.bytes).toBeGreaterThan(bytes)
  expect(ledger.bytes).toBeGreaterThan(cache.bytes)
  expect(ledger.snapshot().byPool.layout).toBe(0)
})

it('merge query exceptions drain workspace and refusal keeps active/coarse guarantees', () => {
  const target = { ...sheet, rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, styleIndex: 0, value: 3 }] }], merges: ['A1000:B1000'], mergeRanges: [{ minRow: 999, maxRow: 999, minCol: 0, maxCol: 1 }] }
  const geometry = buildSheetGeometry(target)
  if (geometry.kind !== 'ready') throw new Error('fixture geometry must be ready')
  const ledger = new MainMemoryLedger({ register: () => {} })
  const cache = new MergeQueryCache(64, mainFormatBudget(ledger, 100000, 'merges'))
  const spy = vi.spyOn(geometry.merges, 'query').mockImplementation(() => { throw new Error('index failure') })
  try {
    expect(() => cache.query(geometry, { r0: 0, r1: 1, c0: 0, c1: 1 }, 100, emptyWork())).toThrow('index failure')
    expect(ledger.bytes).toBe(0)
    expect(cache.bytes).toBe(0)
  } finally { spy.mockRestore() }
  const refused = new MergeQueryCache(64, mainFormatBudget(ledger, 128, 'merges'))
  const plan = planViewport(geometry, { box: { x: 0, y: 0, width: 64, height: 20 }, zoom: 1, deviceScale: 1 }, 'D', { merges: refused, activeCell: { row: 0, col: 0 } })
  expect(plan.kind).toBe('coarse')
  if (plan.kind !== 'coarse') throw new Error('expected coarse refusal')
  expect(plan.active?.source).toEqual({ row: 0, col: 0 })
  expect(plan.groups.length).toBeGreaterThan(0)
  expect(plan.work.overviewGroups).toBe(plan.groups.length)
  expect(plan.diagnostics).toContainEqual({ code: 'layout-memory', message: expect.stringContaining('memory budget') })
  expect(ledger.bytes).toBe(0)
})

it('overlapping searches keep one diagnostic charge until the actual source close', async () => {
  const target: XlsxSheet = { ...sheet, sheetId: '1', rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, styleIndex: 0, value: 3, style: { numFmtId: 164, formatCode: '[Purple]0' } }] }] }
  const source = createObjectWorkbookSource({ sheets: [target] }, 'd'), original = source.displayText!
  const session = new DocumentSession('d', source)
  let replacement: ReturnType<typeof session.search> | undefined, diagnosticBytes = 0
  source.displayText = (cell, id, allocation) => original(cell, id, allocation && { ...allocation, retain(targets, bytes, commit) {
    allocation.retain!(targets, bytes, commit)
    diagnosticBytes += bytes
    replacement ??= session.search(session.workbookOwner(), { text: '3', caseSensitive: true }, undefined, new AbortController().signal)
  } })
  try {
    await session.search(session.workbookOwner(), { text: '3', caseSensitive: false }, undefined, new AbortController().signal)
    expect(await replacement).toMatchObject({ complete: true, hits: [{ text: '3' }] })
    expect(target.diagnostics).toHaveLength(1)
    expect(session.memoryStats().transientBytes).toBe(diagnosticBytes)
    await session.search(session.workbookOwner(), { text: '3', caseSensitive: false }, undefined, new AbortController().signal)
    expect(session.memoryStats().transientBytes).toBe(diagnosticBytes)
    let close!: () => void, began!: () => void
    const gate = new Promise<void>(resolve => { close = resolve }), started = new Promise<void>(resolve => { began = resolve })
    const originalRelease = source.release.bind(source)
    source.release = async () => { began(); await gate; await originalRelease() }
    const closing = session.release()
    await started
    expect(session.memoryStats().transientBytes).toBe(diagnosticBytes)
    close()
    await closing
    expect(session.memoryStats().transientBytes).toBe(0)
  } finally { await session.release() }
})

it('closing a search keeps its diagnostic share until an abort-ignoring producer drains', async () => {
  const target: XlsxSheet = { ...sheet, sheetId: '1', rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, styleIndex: 0, value: 3, style: { numFmtId: 164, formatCode: '[Purple]0' } }] }] }
  const source = createObjectWorkbookSource({ sheets: [target] }, 'd'), store = source.store('1')
  const read = store.readCell.bind(store)
  let drain!: () => void, began!: () => void
  const gate = new Promise<void>(resolve => { drain = resolve }), started = new Promise<void>(resolve => { began = resolve })
  store.readCell = async (at, options) => { began(); await gate; return read(at, options) }
  const session = new DocumentSession('d', source)
  const search = session.search(session.workbookOwner(), { text: '3', caseSensitive: false }, undefined, new AbortController().signal)
  try {
    await started
    expect(target.diagnostics).toHaveLength(1)
    await session.release()
    await search
    expect(session.memoryStats().transientBytes).toBeGreaterThan(0)
    drain()
    for (let i = 0; i < 3; i++) await new Promise(resolve => setTimeout(resolve, 0))
    expect(session.memoryStats().transientBytes).toBe(0)
  } finally { drain(); await session.release(); await search }
})
