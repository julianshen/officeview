/**
 * Stage 4b Task 11 (S4 §6.7, §6.11; plan P10, P12a, P15, Open item 2, Review Focus 2): the a11y grid's text on a 4b
 * stream, fetched from the server by `manifest.caps.a11y`, paced, and cached by `(sheetId, viewGen, cell)`; and the
 * §6.11 notices in the viewer's notice region.
 * - `createStreamA11yText` is driven directly with fake timers (rest, pacing, retry, stale, dispose).
 * - The viewer cases mount `RemoteSheetView` over a fake tiles-mode `RemoteOfficeDoc` (real timers while the grid
 *   mounts, fake ones once it is ready).
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import RemoteSheetView, { streamShellFor } from '../src/components/remote/RemoteSheetView'
import { TEXT_UNAVAILABLE } from '../src/components/remote/textUnavailable'
import { createStreamA11yText, STREAM_A11Y_MAX_CELLS } from '../src/components/xlsx/streamA11y'
import { StreamSessionError } from '../src/stream/remote/errors'
import type { RemoteOfficeDoc, RemoteSheetOps, RemoteSheets } from '../src/stream/remote/remote-doc'
import type { TilesRequest, TilesResult } from '../src/stream/remote/sheet-channel'
import type { SheetFindAnswer, SheetViewChange, StreamSheetLayout } from '../src/stream/sheet-wire'
import { cellIdToCoordinate } from '../src/xlsx/view/coordinates'

afterEach(() => { cleanup(); vi.useRealTimers(); vi.clearAllMocks() })

type A11yFetch = NonNullable<RemoteSheets['view']>['a11y']
type Cell = { row: number; col: number; text: string }
/** Serves `R{row}C{col}` for cells whose row + col is even; the rest are blank (left out, as the server does). */
const served = (rows: Array<[number, number]>, cols: [number, number]): Cell[] => {
  const out: Cell[] = []
  for (const [r0, r1] of rows) for (let row = r0; row <= r1; row++) for (let col = cols[0]; col <= cols[1]; col++) if ((row + col) % 2 === 0) out.push({ row, col, text: `R${row}C${col}` })
  return out
}
const cellsOf = (rows: Array<[number, number]>, cols: [number, number]) => rows.reduce((n, [a, b]) => n + (b - a + 1), 0) * (cols[1] - cols[0] + 1)
const flush = async () => { for (let i = 0; i < 5; i++) await Promise.resolve() }

// ── the controller ─────────────────────────────────────────────────────────

describe('createStreamA11yText (unit)', () => {
  it('a 250 ms rest fetches the window once; the cached window is never fetched again', async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async (_s, _g, rows, cols) => served(rows, cols))
    const t = createStreamA11yText({ mode: 'unit', a11y })
    t.showWindow('0', 0, [[0, 9]], [0, 4])
    expect(t.pending('0', 0, { row: 0, col: 0 })).toBe(true)
    await vi.advanceTimersByTimeAsync(249)
    expect(a11y).not.toHaveBeenCalled()
    await vi.advanceTimersByTimeAsync(1)
    expect(a11y).toHaveBeenCalledTimes(1)
    expect(a11y).toHaveBeenCalledWith('0', 0, [[0, 9]], [0, 4])
    await flush()
    expect(t.textAt('0', 0, { row: 2, col: 2 })).toBe('R2C2')
    // A blank cell is cached as "no text": not loading, not fetched again.
    expect(t.textAt('0', 0, { row: 0, col: 1 })).toBe('')
    expect(t.pending('0', 0, { row: 0, col: 1 })).toBe(false)
    t.showWindow('0', 0, [[5, 20]], [0, 4])
    await vi.advanceTimersByTimeAsync(300)
    // Only the uncached rows 10..20 are asked.
    expect(a11y).toHaveBeenCalledTimes(2)
    expect(a11y.mock.calls[1]).toEqual(['0', 0, [[10, 20]], [0, 4]])
    await flush()
    t.showWindow('0', 0, [[0, 9]], [0, 4])
    await vi.advanceTimersByTimeAsync(1000)
    expect(a11y).toHaveBeenCalledTimes(2)
    t.dispose()
  })

  it('a moving window waits for the rest; only the last one is fetched', async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async (_s, _g, rows, cols) => served(rows, cols))
    const t = createStreamA11yText({ mode: 'unit', a11y })
    for (let i = 0; i < 10; i++) { t.showWindow('0', 0, [[i * 10, i * 10 + 9]], [0, 4]); await vi.advanceTimersByTimeAsync(100) }
    expect(a11y).not.toHaveBeenCalled()
    await vi.advanceTimersByTimeAsync(150)
    expect(a11y.mock.calls).toEqual([['0', 0, [[90, 99]], [0, 4]]])
    t.dispose()
  })

  it(`windows split into rects of at most ${STREAM_A11Y_MAX_CELLS} cells, uncached cells only`, async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async (_s, _g, rows, cols) => served(rows, cols))
    const t = createStreamA11yText({ mode: 'unit', a11y })
    t.showWindow('0', 0, [[0, 99], [200, 249]], [0, 39])
    for (let i = 0; i < 10; i++) await vi.advanceTimersByTimeAsync(300)
    const asked = a11y.mock.calls.map(([, , rows, cols]) => cellsOf(rows, cols))
    for (const n of asked) expect(n).toBeLessThanOrEqual(STREAM_A11Y_MAX_CELLS)
    expect(asked.reduce((a, b) => a + b, 0)).toBe(150 * 40)
    // Each row of the window is asked exactly once.
    const rows = a11y.mock.calls.flatMap(([, , rs]) => rs.flatMap(([a, b]) => Array.from({ length: b - a + 1 }, (_, i) => a + i)))
    expect(new Set(rows).size).toBe(rows.length)
    // A diagonal move asks only the new cells: the old rows' new columns and the new rows' full span.
    a11y.mockClear()
    t.showWindow('0', 0, [[0, 9]], [0, 3])
    await vi.advanceTimersByTimeAsync(300)
    t.showWindow('0', 0, [[95, 104]], [2, 41])
    for (let i = 0; i < 4; i++) await vi.advanceTimersByTimeAsync(300)
    expect(a11y.mock.calls.map(([, , r, c]) => [r, c])).toEqual([[[[95, 99]], [40, 41]], [[[100, 104]], [2, 41]]])
    t.dispose()
  })

  it('pacing: 40 settled scrolls to new windows in 60 s send at most 30 requests', async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async (_s, _g, rows, cols) => served(rows, cols))
    const t = createStreamA11yText({ mode: 'unit', a11y })
    for (let i = 0; i < 40; i++) {
      t.showWindow('0', 0, [[i * 50, i * 50 + 29]], [0, 9])
      await vi.advanceTimersByTimeAsync(1450)
    }
    expect(a11y.mock.calls.length).toBeLessThanOrEqual(30)
    expect(a11y.mock.calls.length).toBeGreaterThanOrEqual(25)
    // The budget frees up as the minute rolls on: the last window is fetched.
    await vi.advanceTimersByTimeAsync(60_000)
    expect(a11y.mock.calls.at(-1)).toEqual(['0', 0, [[1950, 1979]], [0, 9]])
    t.dispose()
  })

  it('rate-limited: one retry after 1 s', async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async () => { throw new StreamSessionError('rate-limited') })
    const t = createStreamA11yText({ mode: 'unit', a11y })
    t.showWindow('0', 0, [[0, 9]], [0, 4])
    await vi.advanceTimersByTimeAsync(250)
    expect(a11y).toHaveBeenCalledTimes(1)
    await vi.advanceTimersByTimeAsync(999)
    expect(a11y).toHaveBeenCalledTimes(1)
    a11y.mockImplementation(async (_s, _g, rows, cols) => served(rows, cols))
    await vi.advanceTimersByTimeAsync(1)
    expect(a11y).toHaveBeenCalledTimes(2)
    await flush()
    expect(t.textAt('0', 0, { row: 0, col: 0 })).toBe('R0C0')
    // A second refusal is not retried: the cells stop loading.
    a11y.mockImplementation(async () => { throw new StreamSessionError('rate-limited') })
    t.showWindow('0', 0, [[10, 19]], [0, 4])
    await vi.advanceTimersByTimeAsync(5000)
    expect(a11y).toHaveBeenCalledTimes(4)
    expect(t.pending('0', 0, { row: 10, col: 0 })).toBe(false)
    expect(t.textAt('0', 0, { row: 10, col: 0 })).toBeUndefined()
    t.dispose()
  })

  it('stale-view drops the reply; the next rest under the new viewGen refetches', async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async (_s, g, rows, cols) => { if (g === 0) throw new StreamSessionError('stale-view'); return served(rows, cols) })
    const t = createStreamA11yText({ mode: 'unit', a11y })
    t.showWindow('0', 0, [[0, 9]], [0, 4])
    await vi.advanceTimersByTimeAsync(250)
    await flush()
    expect(t.textAt('0', 0, { row: 0, col: 0 })).toBeUndefined()
    expect(t.pending('0', 0, { row: 0, col: 0 })).toBe(false)
    t.showWindow('0', 1, [[0, 9]], [0, 4])
    await vi.advanceTimersByTimeAsync(250)
    await flush()
    expect(a11y.mock.calls.map(c => c[1])).toEqual([0, 1])
    expect(t.textAt('0', 1, { row: 0, col: 0 })).toBe('R0C0')
    t.dispose()
  })

  it('a reply filed under an old viewGen never answers the new one', async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async (_s, _g, rows, cols) => served(rows, cols))
    const t = createStreamA11yText({ mode: 'unit', a11y })
    t.showWindow('0', 0, [[0, 9]], [0, 4])
    await vi.advanceTimersByTimeAsync(250)
    await flush()
    expect(t.textAt('0', 0, { row: 0, col: 0 })).toBe('R0C0')
    expect(t.textAt('0', 1, { row: 0, col: 0 })).toBeUndefined()
    t.showWindow('0', 1, [[0, 9]], [0, 4])
    await vi.advanceTimersByTimeAsync(250)
    expect(a11y.mock.calls.map(c => c[1])).toEqual([0, 1])
    t.dispose()
  })

  it('a late reply after dispose neither caches nor notifies; other failures stop loading', async () => {
    vi.useFakeTimers()
    let release!: (cells: Cell[]) => void
    const a11y = vi.fn<A11yFetch>(() => new Promise(r => { release = r }))
    const t = createStreamA11yText({ mode: 'unit', a11y })
    const changed = vi.fn()
    t.onChange(changed)
    t.showWindow('0', 0, [[0, 0]], [0, 0])
    await vi.advanceTimersByTimeAsync(250)
    changed.mockClear()
    t.dispose()
    release([{ row: 0, col: 0, text: 'late' }])
    await flush()
    expect(changed).not.toHaveBeenCalled()
    expect(t.textAt('0', 0, { row: 0, col: 0 })).toBeUndefined()
    const failing = vi.fn<A11yFetch>(async () => { throw new StreamSessionError('session-rotated') })
    const u = createStreamA11yText({ mode: 'unit', a11y: failing })
    u.showWindow('0', 0, [[0, 0]], [0, 0])
    await vi.advanceTimersByTimeAsync(5000)
    expect(failing).toHaveBeenCalledTimes(1)
    expect(u.pending('0', 0, { row: 0, col: 0 })).toBe(false)
    u.dispose()
  })
})

describe('createStreamA11yText (focused, off)', () => {
  it('focused: a 400 ms rest fetches the active cell once (1 × 1); a cached cell is never fetched again', async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async (_s, _g, rows, cols) => served(rows, cols))
    const t = createStreamA11yText({ mode: 'focused', a11y })
    t.showWindow('0', 0, [[0, 30]], [0, 9])
    for (let col = 0; col < 5; col++) { t.focus('0', 0, { row: 1, col }); await vi.advanceTimersByTimeAsync(399) }
    expect(a11y).not.toHaveBeenCalled()
    await vi.advanceTimersByTimeAsync(1)
    expect(a11y.mock.calls).toEqual([['0', 0, [[1, 1]], [4, 4]]])
    await flush()
    // A blank cell (1, 4) is cached too.
    expect(t.textAt('0', 0, { row: 1, col: 4 })).toBe('')
    t.focus('0', 0, { row: 1, col: 3 })
    await vi.advanceTimersByTimeAsync(400)
    await flush()
    expect(t.textAt('0', 0, { row: 1, col: 3 })).toBe('R1C3')
    t.focus('0', 0, { row: 1, col: 4 })
    await vi.advanceTimersByTimeAsync(2000)
    t.focus('0', 0, { row: 1, col: 3 })
    await vi.advanceTimersByTimeAsync(2000)
    expect(a11y).toHaveBeenCalledTimes(2)
    t.dispose()
  })

  it('focused: a quota refusal stops fetching', async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async () => { throw new StreamSessionError('copy-quota-exhausted') })
    const t = createStreamA11yText({ mode: 'focused', a11y })
    t.focus('0', 0, { row: 0, col: 0 })
    await vi.advanceTimersByTimeAsync(400)
    t.focus('0', 0, { row: 0, col: 1 })
    await vi.advanceTimersByTimeAsync(400)
    expect(a11y).toHaveBeenCalledTimes(1)
    expect(t.pending('0', 0, { row: 0, col: 1 })).toBe(false)
    t.dispose()
  })

  it('off fetches nothing', async () => {
    vi.useFakeTimers()
    const a11y = vi.fn<A11yFetch>(async () => [])
    const t = createStreamA11yText({ mode: 'off', a11y })
    t.showWindow('0', 0, [[0, 9]], [0, 4])
    t.focus('0', 0, { row: 0, col: 0 })
    await vi.advanceTimersByTimeAsync(5000)
    expect(a11y).not.toHaveBeenCalled()
    expect(t.pending('0', 0, { row: 0, col: 0 })).toBe(false)
    expect(t.textAt('0', 0, { row: 0, col: 0 })).toBeUndefined()
  })
})

// ── the viewer ─────────────────────────────────────────────────────────────

const AUTHOR_HIDDEN = "Some rows are hidden by the author and can't be shown."
const HIDDEN_CELLS = "Hidden cells can't be shown in this viewing mode."

function wireOf(sheetId: string, o: Partial<StreamSheetLayout> = {}): StreamSheetLayout {
  return {
    sheetId, viewGen: 0, defaults: { rowPx: 20, colPx: 64 }, rows: [], cols: [], collapsedRows: [], collapsedCols: [],
    merges: [], filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 40, totalRows: 40 },
    used: { r0: 0, c0: 0, r1: 39, c1: 9 }, ...o,
  } as StreamSheetLayout
}
function fakeOps(a11y: A11yFetch): RemoteSheetOps {
  return {
    find: vi.fn(async (): Promise<SheetFindAnswer> => ({ hits: [], total: 0, truncated: false })),
    copy: vi.fn(async () => ''),
    a11y: vi.fn(a11y),
    fit: vi.fn(async () => 64),
    resize: vi.fn(async (req: { sheetId: string; viewGen: number }): Promise<SheetViewChange> => ({ sheetId: req.sheetId, viewGen: req.viewGen + 1, change: { reset: true } })),
  }
}
/** A tiles-mode `RemoteOfficeDoc` (`sheetView: 1`) whose manifest can change mid-session. */
function remoteDoc(o: { a11y?: 'unit' | 'focused' | 'off'; copy?: boolean; wire?: (sheetId: string) => StreamSheetLayout; list?: Array<{ sheetId: string; name: string }>; fetch?: A11yFetch } = {}) {
  const list = o.list ?? [{ sheetId: '0', name: 'First' }]
  const ops = fakeOps(o.fetch ?? (async (_s, _g, rows, cols) => served(rows, cols)))
  const sheets: RemoteSheets = {
    list,
    layout: vi.fn(async (sheetId: string) => (o.wire ?? wireOf)(sheetId)),
    tiles: vi.fn((_req: TilesRequest, _onTile: unknown, signal?: AbortSignal): Promise<TilesResult> =>
      new Promise(resolve => signal?.addEventListener('abort', () => resolve({ refused: [], cancelled: true }), { once: true }))),
    view: ops,
  }
  const listeners = new Set<() => void>()
  const manifest = {
    v: 1, docRef: 'doc-ref:4b-a11y', name: 'book.xlsx', kind: 'xlsx', units: { count: 1, list: [{ w: 300, h: 200 }] },
    can: { search: true, copy: o.copy ?? true, print: false, textA11y: true, present: false, notes: false, links: false },
    caps: { copy: null, a11y: o.a11y ?? 'unit', print: { maxPages: null, maxSessionPages: null } },
    recorder: { scalesBuckets: [1, 2] }, notices: [],
    sheetMode: 'tiles', sheets: list,
    tiles: { px: 512, zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4], dprs: [1, 2], palettes: ['light', 'dark'], maxBatch: 64 },
    sheetView: 1,
  }
  const doc = {
    kind: 'remote', manifest,
    onManifestChange: vi.fn((listener: () => void) => { listeners.add(listener); return () => { listeners.delete(listener) } }),
    units: async () => [{ widthPx: 300, heightPx: 200 }],
    paint: vi.fn(async () => {}), search: vi.fn(), select: vi.fn(), copy: vi.fn(), a11y: vi.fn(), print: vi.fn(), paintNotes: vi.fn(async () => {}),
    close: vi.fn(), ended: new Promise(() => {}), sheets,
  }
  /** A resume's new manifest (S3 item 2). */
  const resume = (change: { a11y?: 'unit' | 'focused' | 'off'; copy?: boolean }) => {
    if (change.a11y) manifest.caps.a11y = change.a11y
    if (change.copy !== undefined) manifest.can.copy = change.copy
    for (const l of [...listeners]) l()
  }
  const view = render(<RemoteSheetView document={doc as unknown as RemoteOfficeDoc} />)
  return { doc, ops, view, resume, workbook: streamShellFor(doc as unknown as RemoteOfficeDoc).source }
}

const gridCells = (v: ReturnType<typeof render>) => [...v.container.querySelectorAll<HTMLElement>('[role="gridcell"]')].map(el => {
  const at = cellIdToCoordinate(Number(el.id.replace(/^.*-cell-/, '')))
  return { ...at, el, text: el.textContent ?? '' }
})
async function ready(v: ReturnType<typeof render>) {
  const grid = await v.findByRole('grid', {}, { timeout: 20_000 })
  await waitFor(() => expect(v.container.querySelectorAll('[role="gridcell"]').length).toBeGreaterThan(10))
  return grid
}
const notices = (v: ReturnType<typeof render>) => v.container.querySelector('[role="status"][data-officeview-notice]')?.textContent ?? ''
const count = (text: string, part: string) => text.split(part).length - 1
/** Real time passes (the viewer cases run on real timers; the rests are 250 and 400 ms). */
const wait = async (ms: number) => { await act(async () => { await new Promise(r => setTimeout(r, ms)) }) }
const inWindow = (call: Parameters<A11yFetch>, at: { row: number; col: number }) =>
  call[2].some(([a, b]) => at.row >= a && at.row <= b) && at.col >= call[3][0] && at.col <= call[3][1]
const a1 = (row: number, col: number) => `${String.fromCharCode(65 + col)}${row + 1}`
const expected = (row: number, col: number) => `${a1(row, col)}, value ${(row + col) % 2 === 0 ? `R${row}C${col}` : ''}`

describe('stream a11y in the viewer', { timeout: 30_000 }, () => {
  it('unit mode fills the visible window once', async () => {
    const { view, ops } = remoteDoc({ a11y: 'unit' })
    const grid = await ready(view)
    await waitFor(() => expect(ops.a11y).toHaveBeenCalledTimes(1))
    const call = vi.mocked(ops.a11y).mock.calls[0]!
    expect(call[1]).toBe(0)
    for (const c of gridCells(view)) expect(inWindow(call, c)).toBe(true)
    // Each gridcell's hidden text carries the served text, in the local accessible form ("A1, value R0C0").
    await waitFor(() => expect(gridCells(view)[0]!.text).not.toContain('Loading…'))
    for (const c of gridCells(view)) expect(c.text).toBe(expected(c.row, c.col))
    // Moving the cursor away and back (the same window): nothing more is asked.
    grid.focus()
    fireEvent.keyDown(grid, { key: 'ArrowRight' })
    fireEvent.keyDown(grid, { key: 'ArrowLeft' })
    await wait(600)
    expect(ops.a11y).toHaveBeenCalledTimes(1)
    expect(count(view.container.textContent ?? '', TEXT_UNAVAILABLE)).toBe(0)
  })

  it('focused fetches the rested active cell once', async () => {
    const { view, ops } = remoteDoc({ a11y: 'focused' })
    const grid = await ready(view)
    const active = () => gridCells(view).find(c => c.el.id === grid.getAttribute('aria-activedescendant'))!
    grid.focus()
    // Five moves well inside 400 ms of each other: nothing is asked.
    for (let i = 0; i < 5; i++) { fireEvent.keyDown(grid, { key: 'ArrowDown' }); await wait(50) }
    expect(ops.a11y).not.toHaveBeenCalled()
    const rested = active()
    await waitFor(() => expect(ops.a11y).toHaveBeenCalledTimes(1))
    expect(vi.mocked(ops.a11y).mock.calls[0]!.slice(2)).toEqual([[[rested.row, rested.row]], [rested.col, rested.col]])
    await waitFor(() => expect(active().text).toBe(expected(rested.row, rested.col)))
    fireEvent.keyDown(grid, { key: 'ArrowRight' })
    await waitFor(() => expect(ops.a11y).toHaveBeenCalledTimes(2))
    await waitFor(() => expect(active().text).toBe(expected(rested.row, rested.col + 1)))
    // Returning to cached cells asks nothing.
    fireEvent.keyDown(grid, { key: 'ArrowLeft' })
    await wait(600)
    fireEvent.keyDown(grid, { key: 'ArrowRight' })
    await wait(600)
    expect(ops.a11y).toHaveBeenCalledTimes(2)
  })

  it('off fetches nothing and says so once', async () => {
    const { view, ops } = remoteDoc({ a11y: 'off', copy: true })
    await ready(view)
    await wait(600)
    expect(ops.a11y).not.toHaveBeenCalled()
    expect(view.getAllByRole('status').filter(s => s.textContent === TEXT_UNAVAILABLE)).toHaveLength(1)
    expect(count(view.container.textContent ?? '', TEXT_UNAVAILABLE)).toBe(1)
    for (const c of gridCells(view)) expect(c.text).toBe('')
  })

  it('copy off says text unavailable once; a11y on still fills the grid', async () => {
    const { view, ops } = remoteDoc({ a11y: 'unit', copy: false })
    await ready(view)
    await waitFor(() => expect(ops.a11y).toHaveBeenCalledTimes(1))
    expect(count(view.container.textContent ?? '', TEXT_UNAVAILABLE)).toBe(1)
  })

  it('resize invalidates by viewGen', async () => {
    const { view, ops, workbook } = remoteDoc({ a11y: 'unit' })
    await ready(view)
    await waitFor(() => expect(ops.a11y).toHaveBeenCalledTimes(1))
    act(() => workbook.applyView({ sheetId: '0', viewGen: 1, change: { axis: 'col', index: 9, size: 90 } }))
    await waitFor(() => expect(vi.mocked(ops.a11y).mock.calls.length).toBeGreaterThanOrEqual(2))
    await wait(600)
    const calls = vi.mocked(ops.a11y).mock.calls
    // Every later ask is under viewGen 1, and together they cover the shown window again (each cell once).
    expect(calls.slice(1).every(c => c[1] === 1)).toBe(true)
    const keys = calls.slice(1).flatMap(c => c[2].flatMap(([a, b]) => Array.from({ length: b - a + 1 }, (_, i) => a + i)).flatMap(r => Array.from({ length: c[3][1] - c[3][0] + 1 }, (_, j) => `${r}:${c[3][0] + j}`)))
    expect(new Set(keys).size).toBe(keys.length)
    for (const c of gridCells(view)) expect(keys).toContain(`${c.row}:${c.col}`)
    await waitFor(() => expect(gridCells(view)[0]!.text).toBe('A1, value R0C0'))
  })

  it('a11y turned off on resume clears the text', async () => {
    const { view, ops, resume } = remoteDoc({ a11y: 'unit' })
    await ready(view)
    await waitFor(() => expect(gridCells(view).some(c => c.text === 'A1, value R0C0')).toBe(true))
    act(() => resume({ a11y: 'off' }))
    for (const c of gridCells(view)) expect(c.text).toBe('')
    expect(count(view.container.textContent ?? '', TEXT_UNAVAILABLE)).toBe(1)
    const grid = view.getByRole('grid')
    grid.focus()
    fireEvent.keyDown(grid, { key: 'PageDown' })
    await wait(600)
    expect(ops.a11y).toHaveBeenCalledTimes(1)
    for (const c of gridCells(view)) expect(c.text).toBe('')
  })

  it('notices: author-hidden once per sheet; a hidden Go To target says so and opens no inspection', async () => {
    const { view } = remoteDoc({
      a11y: 'unit',
      list: [{ sheetId: '0', name: 'First' }, { sheetId: '1', name: 'Second' }],
      wire: id => wireOf(id, id === '0' ? { authorHidden: true, collapsedRows: [{ start: 5, end: 5 }] } : {}),
    })
    await ready(view)
    await waitFor(() => expect(count(notices(view), AUTHOR_HIDDEN)).toBe(1))
    expect(count(view.container.textContent ?? '', AUTHOR_HIDDEN)).toBe(1)
    expect(view.container.textContent).not.toContain(HIDDEN_CELLS)
    fireEvent.change(view.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A6' } })
    fireEvent.click(view.getByRole('button', { name: 'Go To' }))
    await waitFor(() => expect(count(view.container.textContent ?? '', HIDDEN_CELLS)).toBe(1))
    // A second attempt does not repeat it, and no inspection opens.
    fireEvent.click(view.getByRole('button', { name: 'Go To' }))
    await wait(0)
    expect(count(view.container.textContent ?? '', HIDDEN_CELLS)).toBe(1)
    expect(view.container.textContent).not.toContain('Read-only inspection')
    expect(view.container.textContent).not.toContain('read-only.')
    // The other sheet has no author-hidden notice.
    fireEvent.change(view.getByRole('combobox', { name: 'Worksheet' }), { target: { value: (view.getByRole('option', { name: 'Second' }) as HTMLOptionElement).value } })
    await waitFor(() => expect(notices(view)).not.toContain(AUTHOR_HIDDEN))
  })
})
