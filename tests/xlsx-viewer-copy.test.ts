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
    const before = JSON.stringify({ selection: f.selection, bands: f.order.bands, visibility: f.visibility, columns: f.columns })
    expect(await textOf(f)).toBe('3:0\t3:2\n0:0\t0:2\n1:0\t1:2')
    expect(JSON.stringify({ selection: f.selection, bands: f.order.bands, visibility: f.visibility, columns: f.columns })).toBe(before)
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

it('admits projection before allocating coverage or reading values', async () => {
  const f = input((r, c) => value(r, c, 'x'))
  f.selection.ranges = [part(0, 0, 0, 191)]
  const requests: number[] = []
  Object.assign(f, { admitBytes: (bytes: number) => { requests.push(bytes); return bytes < 32 * 192 } })
  expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
  expect(requests.some(bytes => bytes >= 32 * 192)).toBe(true)
  expect(f.store.readCell).not.toHaveBeenCalled()
})

it('reserves tiny field and quoting growth before constructing escaped text', async () => {
  const f = input((r, c) => value(r, c, '"'.repeat(200)))
  f.selection.ranges = [part(0, 0)]
  const replace = String.prototype.replace, escaped = vi.fn()
  vi.spyOn(String.prototype, 'replace').mockImplementation(function (this: string, ...args: Parameters<typeof replace>) {
    if (String(this) === '"'.repeat(200)) escaped()
    return replace.apply(this, args)
  })
  Object.assign(f, { admitBytes: (bytes: number) => bytes < 128 + 4 * 402 })
  try {
    expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
    expect(escaped).not.toHaveBeenCalled()
  } finally { vi.restoreAllMocks() }
})

it('keeps onBytes notification compatible and never reports denied growth', async () => {
  const f = input((r, c) => value(r, c, 'x'))
  f.selection.ranges = [part(0, 0)]
  const readings: number[] = []
  f.onBytes = bytes => readings.push(bytes)
  Object.assign(f, { admitBytes: () => false })
  expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
  expect(readings).toEqual([])
  expect(f.store.readCell).not.toHaveBeenCalled()
})

it('guards canonical fourth-section expansion before constructing its full output', async () => {
  const { cellDisplayText } = await import('../src/xlsx/render')
  const doc = workbookOf([{ sheetId: 'A', cells: [] }]), sheet = doc.sheets[0]
  const f = input((r, c) => value(r, c, 'x'.repeat(256)))
  f.selection.ranges = [part(0, 0)]
  const code = '0;0;0;' + '@'.repeat(1024), requests: number[] = []
  f.format = (cell, allocation) => cellDisplayText(sheet, { value: cell.value, style: { numFmtId: 0, formatCode: code } }, allocation)
  Object.assign(f, { admitBytes: (bytes: number) => { requests.push(bytes); return bytes < 150_000 } })
  expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
  expect(requests.find(bytes => bytes >= 150_000)).toBeLessThan(200_000)
  Object.assign(f, { admitBytes: () => true })
  expect(await textOf(f)).toBe('x'.repeat(256 * 1024))
})

it('shares live headroom between two suspended serializers and leaves refused growth unchanged', async () => {
  const { MainMemoryLedger } = await import('../src/xlsx/view/mainLedger')
  const { MAIN_THREAD_BUDGETS } = await import('../src/xlsx/view/tileCache')
  const ledger = new MainMemoryLedger(), totals: number[] = [], denials: [number, number][] = []
  ledger.set('controls', 'pressure', MAIN_THREAD_BUDGETS.D - 720)
  const pending: (() => void)[] = []
  const make = (owner: string) => {
    const f = input((r, c) => value(r, c, 'x'.repeat(100)))
    f.selection.ranges = [part(0, 0)]
    const read = f.store.readCell.bind(f.store)
    f.store.readCell = async (...args) => { await new Promise<void>(resolve => pending.push(resolve)); return read(...args) }
    let held = 0
    f.admitBytes = next => {
      const before = ledger.bytes
      if (next > held && before - held + next > MAIN_THREAD_BUDGETS.D) { denials.push([before, ledger.bytes]); return false }
      ledger.set('copy', owner, next); held = next; totals.push(ledger.bytes); return true
    }
    return f
  }
  const a = serializeSelection(make('A')), b = serializeSelection(make('B'))
  await vi.waitFor(() => expect(pending.length).toBe(2))
  expect(ledger.snapshot().byPool.copy).toBe(320)
  pending[0]()
  expect(await a).toMatchObject({ kind: 'ready', text: 'x'.repeat(100) })
  pending[1]()
  expect(await b).toMatchObject({ kind: 'refused', code: 'resource-limit' })
  expect(denials).toHaveLength(1)
  expect(denials[0][0]).toBe(denials[0][1])
  expect(ledger.snapshot().byPool.copy).toBe(528)
  expect(totals.every(bytes => bytes <= MAIN_THREAD_BUDGETS.D)).toBe(true)
  ledger.release('copy', 'A')
  expect(ledger.snapshot().byPool.copy).toBe(0)
})

it.each(['intervals', 'list', 'bitset'] as const)('admits each newly decoded %s run before row projection', async kind => {
  const f = input((r, c) => value(r, c, 'x'))
  f.selection.ranges = [{ cols: [{ c0: 0, c1: 0 }], rows: kind === 'intervals'
    ? { kind, intervals: [{ r0: 0, r1: 0 }, { r0: 2, r1: 2 }] }
    : kind === 'list' ? { kind, r0: 0, r1: 2, rows: Uint32Array.of(0, 2) }
    : { kind, space: 'source', r0: 0, r1: 2, words: Uint32Array.of(5), count: 2 } }]
  const rank = vi.spyOn(f.visibility, 'rank')
  f.admitBytes = next => next < 128
  expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
  expect(rank).not.toHaveBeenCalled()
  expect(f.store.readCell).not.toHaveBeenCalled()
})

it('counts newly formed row runs, events and sorting scratch before source reads', async () => {
  const f = input((r, c) => value(r, c, 'x'))
  f.selection.ranges = [{ cols: [{ c0: 0, c1: 1 }, { c0: 3, c1: 3 }], rows: { kind: 'list', r0: 0, r1: 2, rows: Uint32Array.of(0, 2) } }]
  f.order = createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: Uint32Array.of(2, 1, 0, 3) }])
  const requests: number[] = []
  f.admitBytes = next => { requests.push(next); return next < 512 }
  expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
  expect(requests.some(next => next >= 512)).toBe(true)
  expect(f.store.readCell).not.toHaveBeenCalled()
})

it.each([480, 479])('small serializer needs its measured %i-byte construction capacity', async capacity => {
  const f = input((r, c) => value(r, c, 'beta'))
  f.selection.ranges = [part(0, 0)]
  f.admitBytes = next => next <= capacity
  const out = await serializeSelection(f)
  expect(out).toMatchObject(capacity === 480 ? { kind: 'ready', text: 'beta' } : { kind: 'refused', code: 'resource-limit' })
})

it('refunds admitted scratch after formatter and observer exceptions', async () => {
  for (const failure of ['formatter', 'observer']) {
    const f = input((r, c) => value(r, c, 'beta')), readings: number[] = []
    f.selection.ranges = [part(0, 0)]
    let held = 0
    f.admitBytes = next => { held = next; readings.push(next); return true }
    if (failure === 'formatter') f.format = (_, allocation) => { allocation?.allocate(512); throw new Error('format failed') }
    else f.onBytes = bytes => { if (bytes > 0) throw new Error('observer failed') }
    expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'page-error' })
    expect(Math.max(...readings)).toBeGreaterThan(0)
    expect(held).toBe(0)
  }
})

it('keeps canonical raw string passthrough borrowed after parsing a format', async () => {
  const { cellDisplayText } = await import('../src/xlsx/render')
  const sheet = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0], f = input((r, c) => value(r, c, 'x'.repeat(100_000)))
  f.selection.ranges = [part(0, 0)]
  f.format = (cell, allocation) => cellDisplayText(sheet, { value: cell.value, style: { numFmtId: 0, formatCode: '0' } }, allocation)
  const requests: number[] = []
  f.admitBytes = next => { requests.push(next); return true }
  expect(await textOf(f)).toBe('x'.repeat(100_000))
  expect(Math.max(...requests)).toBeLessThan(401_000)
})

it('adopts new canonical diagnostics while duplicates retain only temporary copy storage', async () => {
  const { cellDisplayText } = await import('../src/xlsx/render')
  const { MainMemoryLedger } = await import('../src/xlsx/view/mainLedger')
  const watched: { target: object; token: string }[] = []
  const ledger = new MainMemoryLedger({ register: (target, token) => watched.push({ target, token }) })
  const sheet = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0]
  const f = input((r, c) => value(r, c, 'beta'))
  f.selection.ranges = [part(0, 0)]
  f.format = (cell, allocation) => cellDisplayText(sheet, { value: cell.value, style: { numFmtId: 0, formatCode: '0;0;0;yyyy' } }, allocation)
  let held = 0
  f.admitBytes = next => { ledger.set('copy', 'job', next); held = next; return true }
  f.retainBytes = (targets, bytes, commit) => { ledger.adoptWhileReachable('copy', 'job', 'document', targets, bytes, commit); held -= bytes }
  expect(await textOf(f)).toBe('beta')
  const retained = ledger.snapshot().byPool.document
  expect(retained).toBeGreaterThan(100)
  expect(held).toBe(144)
  expect(sheet.diagnostics).toHaveLength(1)
  ledger.release('copy', 'job')
  expect(ledger.bytes).toBe(retained)
  expect(await textOf(f)).toBe('beta')
  ledger.release('copy', 'job')
  expect(ledger.bytes).toBe(retained)
  expect(sheet.diagnostics).toHaveLength(1)
  expect(watched).toHaveLength(2)
  const token = watched[0].token
  // Collection watcher reports the two independently retained targets ending.
  ledger.collected(token)
  expect(ledger.bytes).toBe(retained)
  ledger.collected(token)
  expect(ledger.bytes).toBe(0)
})

it('denies canonical diagnostic growth before publication with a typed whole-copy refusal', async () => {
  const { cellDisplayText } = await import('../src/xlsx/render')
  const sheet = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0], f = input((r, c) => value(r, c, 'beta'))
  sheet.diagnostics = [{ kind: 'test', feature: 'existing', message: 'keep' }]
  const before = JSON.stringify(sheet.diagnostics)
  f.selection.ranges = [part(0, 0)]
  f.format = (cell, allocation) => cellDisplayText(sheet, { value: cell.value, style: { numFmtId: 0, formatCode: '0;0;0;yyyy' } }, allocation)
  f.admitBytes = next => next <= 384
  expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
  expect(JSON.stringify(sheet.diagnostics)).toBe(before)
})

it('reserves an unquoted tiny concat before coercing its formatted input', async () => {
  const f = input((r, c) => value(r, c, 'tiny'))
  f.selection.ranges = [part(0, 0)]
  let held = 0
  const observed: number[] = [], text = new String('tiny')
  // A boxed string instruments the concat coercion boundary; source values
  // remain ordinary valid primitive strings throughout this serializer test.
  Object.defineProperty(text, Symbol.toPrimitive, { value: () => { observed.push(held); return 'tiny' } })
  f.format = () => text as unknown as string
  f.admitBytes = next => { held = next; return true }
  expect(await textOf(f)).toBe('tiny')
  expect(observed).toEqual([184]) // metadata128 + slots32 + formatted8 + output16
})

it('reserves chunk slots and the final join before either constructs output', async () => {
  const f = input((r, c) => value(r, c, 'x'.repeat(100_000)))
  f.selection.ranges = [part(0, 0)]
  let atJoinReservation = 0
  f.admitBytes = next => next !== 400_136 || ++atJoinReservation < 2
  const join = Array.prototype.join, joined: number[] = []
  vi.spyOn(Array.prototype, 'join').mockImplementation(function (this: unknown[], ...args: Parameters<typeof join>) {
    if (args[0] === '' && this.length === 1 && this[0] === 'x'.repeat(100_000)) joined.push(1)
    return join.apply(this, args)
  })
  try {
    expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
    expect(joined).toEqual([])
  } finally { vi.restoreAllMocks() }
})

it('refused coverage storage constructs neither typed coverage array', async () => {
  const f = input((r, c) => value(r, c, 'x'))
  f.selection.ranges = [part(0, 0, 0, 191)]
  const Original = Int32Array, constructed: number[] = []
  vi.stubGlobal('Int32Array', class extends Original { constructor(length: number) { constructed.push(length); super(length) } })
  f.admitBytes = next => next < 32 * 192
  try {
    expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'resource-limit' })
    expect(constructed).toEqual([])
  } finally { vi.unstubAllGlobals() }
})

it('failed canonical diagnostic publication leaves existing diagnostics and ownership unchanged', async () => {
  const { cellDisplayText } = await import('../src/xlsx/render')
  const { MainMemoryLedger } = await import('../src/xlsx/view/mainLedger')
  const ledger = new MainMemoryLedger(), sheet = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0]
  sheet.diagnostics = Object.freeze([{ kind: 'test', feature: 'existing', message: 'keep' }]) as unknown as NonNullable<typeof sheet.diagnostics>
  const before = JSON.stringify(sheet.diagnostics), f = input((r, c) => value(r, c, 'beta'))
  f.selection.ranges = [part(0, 0)]
  f.format = (cell, allocation) => cellDisplayText(sheet, { value: cell.value, style: { numFmtId: 0, formatCode: '0;0;0;yyyy' } }, allocation)
  f.admitBytes = next => { ledger.set('copy', 'job', next); return true }
  f.retainBytes = (targets, bytes, commit) => { ledger.adoptWhileReachable('copy', 'job', 'document', targets, bytes, commit) }
  expect(await serializeSelection(f)).toMatchObject({ kind: 'refused', code: 'page-error' })
  expect(JSON.stringify(sheet.diagnostics)).toBe(before)
  expect(ledger.snapshot().byPool).toMatchObject({ copy: 0, document: 0 })
})

it('keeps canonical newly formatted strings charged across the next batch source read', async () => {
  const { cellDisplayText } = await import('../src/xlsx/render')
  const sheet = workbookOf([{ sheetId: 'A', cells: [] }]).sheets[0], f = input((r, c) => value(r, c, 'x'.repeat(256)))
  f.selection.ranges = [part(0, 0, 0, 1)]
  f.format = (cell, allocation) => cellDisplayText(sheet, { value: cell.value, style: { numFmtId: 0, formatCode: '0;0;0;' + '@'.repeat(64) } }, allocation)
  let held = 0, atSecondRead = 0
  f.admitBytes = next => { held = next; return true }
  const read = f.store.readCell.bind(f.store)
  f.store.readCell = async (...args) => { if (args[0].col === 1) atSecondRead = held; return read(...args) }
  expect(await textOf(f)).toBe('x'.repeat(16384) + '\t' + 'x'.repeat(16384))
  expect(atSecondRead).toBe(32960) // owned field32768 + metadata128 + two record/slots64
  expect(held).toBe(131204)
})
