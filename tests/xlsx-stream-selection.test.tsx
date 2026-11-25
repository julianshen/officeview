/**
 * Stage 4b Task 10 (S4 §6.7, §6.11; plan P6, P7, P11–P12a, P15, P17): the viewer on a 4b stream (`sheetView: 1`).
 * - Selection, keyboard, header and Go To navigation run the local code: the parity cases mount a local viewer on
 *   `textGrid`'s visible-only text-view document and a stream viewer over `buildStreamLayout` of the same snapshot, and
 *   compare the committed selections and the active descendant (the cell part of its id: the id prefix carries the
 *   viewer instance, docId, sheetId and viewGen, which differ by construction). P17: the panes are the authored 1 × 1.
 * - Resize and fit go to the server (`fit`, then `resize`); key repeat folds into one pending target.
 * - Find waits for typing to rest; Copy and Ctrl+C go to `copy`; there is no local text index (parent §8.3).
 * - 4a streams (no `sheetView`) still select nothing and have no Find or Copy.
 *
 * The local viewer has no Ctrl+A handler and no drag-selection gesture: the scripts keep both steps, and they are
 * equal no-ops on both viewers (the corner "Select worksheet" button is the select-all gesture).
 */
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'

const spies = vi.hoisted(() => ({ workbookFind: vi.fn(), copySelection: vi.fn() }))
vi.mock('../src/xlsx/view/find', async original => {
  const actual = await original<typeof import('../src/xlsx/view/find')>()
  // FindController builds its default session inside the module, so the prototype's `nextPage` is spied too (below).
  class SpiedFind extends actual.WorkbookFind {
    constructor(...args: ConstructorParameters<typeof actual.WorkbookFind>) { spies.workbookFind(...args); super(...args) }
  }
  return { ...actual, WorkbookFind: SpiedFind }
})
vi.mock('../src/xlsx/view/copy', async original => {
  const actual = await original<typeof import('../src/xlsx/view/copy')>()
  return { ...actual, copySelection: vi.fn((...args: Parameters<typeof actual.copySelection>) => { spies.copySelection(...args); return actual.copySelection(...args) }) }
})

import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import RemoteSheetView from '../src/components/remote/RemoteSheetView'
import { StreamWorkbook, streamShellDoc } from '../src/components/xlsx/streamWorkbook'
import { WorkbookFind } from '../src/xlsx/view/find'
import { StreamSessionError } from '../src/stream/remote/errors'
import type { RemoteOfficeDoc, RemoteSheetOps, RemoteSheets } from '../src/stream/remote/remote-doc'
import type { TilesRequest, TilesResult } from '../src/stream/remote/sheet-channel'
import { decodeSheetLayout, encodeSheetLayout, type SheetFindAnswer, type SheetViewChange, type StreamSheetLayout } from '../src/stream/sheet-wire'
import type { GridSelection } from '../src/xlsx/view/contracts'
import { cellId, cellIdToCoordinate } from '../src/xlsx/view/coordinates'
import type { XlsxDocument } from '../src/xlsx/types'
import { buildSheetGeometry, type SheetGeometry } from '../src/xlsx/view/viewport'
import { streamSheetGeometry } from '../src/xlsx/view/streamGeometry'
import { NO_SESSION_SIZES } from '../src/xlsx/view/axisSizes'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { prepareSheetAssets } from '../src/components/xlsx/sheetAssets'
import { projectPackage } from '../server/src/policy'
import { buildSheetSnapshot, servableSheet, type SheetSnapshot } from '../server/src/sheet/snapshot'
import { sheetTextView, textDocument, type TextView } from '../server/src/sheet/text-view'
import { buildStreamLayout } from '../server/src/sheet/layout'
import { clockEpochMs, type SheetClock } from '../server/src/sheet/clock'
import { textGrid, TEXT_GRID } from '../server/test/sheet-fixtures'

// The module's own WorkbookFind (the mock's base class): FindController constructs that one.
const nextPage = vi.spyOn(Object.getPrototypeOf(WorkbookFind.prototype) as InstanceType<typeof WorkbookFind>, 'nextPage')
Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: vi.fn(async () => {}) } })

afterEach(() => { cleanup(); vi.clearAllMocks() })

// ── the textGrid fixture: local text-view document and the stream's layout of the same snapshot ──
const CLOCK: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
let fixture: { local: XlsxDocument; wire: StreamSheetLayout; geometry: SheetGeometry; snap: SheetSnapshot; view: TextView }
beforeAll(async () => {
  // DocumentHost cannot load under vitest (server/src/fonts.ts reads Bun's `import.meta.dir`), so the snapshot is built
  // the way `DocumentHost.sheetSnapshot` builds it: the projected package parsed under 'cached', its servable sheets.
  const projected = await projectPackage(await OfficePackage.load(await textGrid()), 'xlsx', 'visible-only')
  const saved = await parseXlsx(projected.pkg, { valuePolicy: 'cached' })
  const workbook: XlsxDocument = { ...saved, sheets: saved.sheets.map(s => servableSheet(s, 'visible-only').sheet) }
  const snap = await buildSheetSnapshot(saved, 0, 'visible-only', CLOCK, () => prepareSheetAssets(workbook, { sheetIndex: 0, cfClock: { epochNowMs: clockEpochMs(CLOCK), timeZone: CLOCK.timeZone }, chartOwner: saved }))
  const view = await sheetTextView(snap, saved.sheets[0]!, NO_SESSION_SIZES, { workbook, clock: CLOCK, budget: { reserve: () => {}, release: () => {} } })
  const local = textDocument(workbook, new Map([[0, view]]))
  const wire = decodeSheetLayout(JSON.parse(JSON.stringify(encodeSheetLayout(buildStreamLayout(snap, '0')))))
  const built = buildSheetGeometry(local.sheets[0]!)
  if (built.kind !== 'ready') throw new Error('textGrid geometry is a diagnostic')
  fixture = { local, wire, geometry: built, snap, view }
}, 60_000)
afterAll(() => { fixture?.view.release(); fixture?.snap.release() })

// ── stream fakes ──
function simpleWire(sheetId: string): StreamSheetLayout {
  return {
    sheetId, viewGen: 0, defaults: { rowPx: 20, colPx: 64 }, rows: [], cols: [], collapsedRows: [], collapsedCols: [],
    merges: [], filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 40, totalRows: 40 },
    used: { r0: 0, c0: 0, r1: 39, c1: 9 },
  } as StreamSheetLayout
}
/** The geometry of `simpleWire` (the server's, as the stream viewer builds it). */
const simple = () => streamSheetGeometry(simpleWire('0')) as unknown as SheetGeometry
type Ops = { [K in keyof RemoteSheetOps]: ReturnType<typeof vi.fn> & RemoteSheetOps[K] }
function fakeOps(o: Partial<RemoteSheetOps> = {}): Ops {
  return {
    find: vi.fn(o.find ?? (async (): Promise<SheetFindAnswer> => ({ hits: [{ sheetIndex: 0, row: 2, col: 1 }], total: 1, truncated: false }))),
    copy: vi.fn(o.copy ?? (async () => 'a\tb')),
    a11y: vi.fn(o.a11y ?? (async () => [])),
    fit: vi.fn(o.fit ?? (async () => 120)),
    resize: vi.fn(o.resize ?? (async (req: { sheetId: string; viewGen: number; axis?: 'row' | 'col'; index?: number; size?: number; reset?: true }): Promise<SheetViewChange> =>
      ({ sheetId: req.sheetId, viewGen: req.viewGen + 1, change: req.reset ? { reset: true } : { axis: req.axis!, index: req.index!, size: req.size! } }))),
  } as Ops
}
function fakeSheets(o: { wire?: (sheetId: string) => StreamSheetLayout; ops?: Ops | null; list?: Array<{ sheetId: string; name: string }> } = {}) {
  const calls: TilesRequest[] = []
  const list = o.list ?? [{ sheetId: '0', name: 'First' }]
  const ops = o.ops === null ? undefined : o.ops ?? fakeOps()
  const sheets: RemoteSheets = {
    list,
    layout: vi.fn(async (sheetId: string) => (o.wire ?? simpleWire)(sheetId)),
    tiles: vi.fn((req: TilesRequest, _onTile: unknown, signal?: AbortSignal): Promise<TilesResult> => {
      calls.push(req)
      return new Promise(resolve => signal?.addEventListener('abort', () => resolve({ refused: [], cancelled: true }), { once: true }))
    }),
    ...(ops ? { view: ops } : {}),
  }
  return { sheets, ops, calls }
}
function mountStream(o: Parameters<typeof fakeSheets>[0] & { onSelection?: (sheetId: string, s: GridSelection) => void; allowCopy?: boolean } = {}) {
  const fake = fakeSheets(o)
  const workbook = new StreamWorkbook(fake.sheets, 'doc-ref')
  const doc = streamShellDoc(fake.sheets.list)
  const view = render(<InteractiveSheetDoc doc={doc} source={workbook} allowCopy={o.allowCopy ?? false} {...(o.onSelection ? { onSelection: o.onSelection } : {})} />)
  return { ...fake, workbook, view }
}
/** A RemoteOfficeDoc in tiles mode whose manifest can change mid-session (`setCan`). */
function remoteDoc(o: { sheetView?: boolean; copy?: boolean; caps?: unknown; ops?: Ops } = {}) {
  const fake = fakeSheets({ ops: o.sheetView === false ? null : o.ops ?? fakeOps() })
  const listeners = new Set<() => void>()
  const manifest = {
    v: 1, docRef: 'doc-ref:4b', name: 'book.xlsx', kind: 'xlsx', units: { count: 1, list: [{ w: 300, h: 200 }] },
    can: { search: true, copy: o.copy ?? true, print: false, textA11y: false, present: false, notes: false, links: false },
    caps: { copy: o.caps ?? null, a11y: 'off', print: { maxPages: null, maxSessionPages: null } },
    recorder: { scalesBuckets: [1, 2] }, notices: [],
    sheetMode: 'tiles', sheets: fake.sheets.list,
    tiles: { px: 512, zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4], dprs: [1, 2], palettes: ['light', 'dark'], maxBatch: 64 },
    ...(o.sheetView === false ? {} : { sheetView: 1 }),
  }
  const doc = {
    kind: 'remote', manifest,
    onManifestChange: vi.fn((listener: () => void) => { listeners.add(listener); return () => { listeners.delete(listener) } }),
    units: async () => [{ widthPx: 300, heightPx: 200 }],
    paint: vi.fn(async () => {}), search: vi.fn(), select: vi.fn(), copy: vi.fn(), a11y: vi.fn(), print: vi.fn(), paintNotes: vi.fn(async () => {}),
    close: vi.fn(), ended: new Promise(() => {}), sheets: fake.sheets,
  }
  const setCan = (can: Partial<typeof manifest.can>) => { Object.assign(manifest.can, can); for (const l of [...listeners]) l() }
  const view = render(<RemoteSheetView document={doc as unknown as RemoteOfficeDoc} />)
  return { ...fake, doc, view, setCan }
}

// ── DOM helpers ──
const gridOf = (v: ReturnType<typeof render>) => v.findByRole('grid', {}, { timeout: 20_000 })
const statusText = (v: ReturnType<typeof render>) => v.container.querySelector('[data-officeview-ui="viewer-status"]')?.textContent ?? ''
const cellPart = (id: string | null) => id === null ? null : id.replace(/^.*-cell-/, '')
/** The cell part of a gridcell id (`cellId`). */
const idOf = (row: number, col: number) => String(cellId(row, col))
const rowColOf = (part: string | null) => part === null ? null : cellIdToCoordinate(Number(part))
/** The client point of a cell's top-left corner (+2 px) in a grid at its initial window (frozen 1 × 1 at the origin). */
function cellPoint(grid: HTMLElement, geometry: SheetGeometry, row: number, col: number) {
  const area = grid.firstElementChild as HTMLElement
  const gutter = parseFloat(area.style.left), header = parseFloat(area.style.top)
  const m = geometry.metrics
  return { clientX: gutter + m.cols.start(col) + 2, clientY: header + m.rows.start(m.order.sourceToDisplay(row)) + 2 }
}
async function ready(v: ReturnType<typeof render>) {
  const grid = await gridOf(v)
  await waitFor(() => expect(v.container.querySelectorAll('[role="gridcell"]').length).toBeGreaterThan(10))
  return grid
}

type Step = { kind: 'key'; key: string; shift?: boolean; ctrl?: boolean } | { kind: 'click'; row: number; col: number; shift?: boolean; ctrl?: boolean }
  | { kind: 'header'; axis: 'column' | 'row'; index: number; shift?: boolean } | { kind: 'corner' } | { kind: 'drag'; from: [number, number]; to: [number, number] }
/** Runs a script on one viewer; the committed selections and the active descendant after each step. */
async function script(v: ReturnType<typeof render>, grid: HTMLElement, steps: Step[], selections: GridSelection[]) {
  const active: Array<string | null> = []
  for (const step of steps) {
    await act(async () => {
      if (step.kind === 'key') { grid.focus(); fireEvent.keyDown(grid, { key: step.key, shiftKey: !!step.shift, ctrlKey: !!step.ctrl }) }
      else if (step.kind === 'click') fireEvent.click(grid, { ...cellPoint(grid, fixture.geometry, step.row, step.col), shiftKey: !!step.shift, ctrlKey: !!step.ctrl })
      else if (step.kind === 'header') {
        const button = grid.querySelector<HTMLButtonElement>(step.axis === 'column' ? `[data-header-column="${step.index}"]` : `[data-header-row="${step.index}"]`)
        if (!button) throw new Error(`no ${step.axis} header ${step.index}`)
        fireEvent.click(button, { shiftKey: !!step.shift })
      } else if (step.kind === 'corner') fireEvent.click(v.getAllByRole('button', { name: 'Select worksheet' }).find(b => grid.contains(b))!)
      else {
        const from = cellPoint(grid, fixture.geometry, ...step.from), to = cellPoint(grid, fixture.geometry, ...step.to)
        fireEvent.pointerDown(grid, { ...from, pointerId: 1 })
        fireEvent.pointerMove(grid, { ...to, pointerId: 1 })
        fireEvent.pointerUp(grid, { ...to, pointerId: 1 })
        fireEvent.click(grid, to)
      }
      await new Promise(r => setTimeout(r, 0))
    })
    active.push(cellPart(grid.getAttribute('aria-activedescendant')))
  }
  return { selections: [...selections], active }
}
async function parity(steps: Step[]) {
  const localSel: GridSelection[] = [], streamSel: GridSelection[] = []
  const local = render(<InteractiveSheetDoc doc={fixture.local} onSelection={(_id, s) => localSel.push(s)} />)
  const localGrid = await ready(local)
  const a = await script(local, localGrid, steps, localSel)
  cleanup()
  const s = mountStream({ list: TEXT_GRID.sheets.map((name, i) => ({ sheetId: String(i), name })), wire: () => fixture.wire, onSelection: (_id, sel) => streamSel.push(sel) })
  const streamGrid = await ready(s.view)
  await waitFor(() => expect(s.workbook.loadState('0').completeness).toBe('complete'))
  const b = await script(s.view, streamGrid, steps, streamSel)
  return { local: a, stream: b }
}

describe('4b stream selection parity (P17: under the authored frozen panes)', { timeout: 60_000 }, () => {
  it('the fixture: hidden row 7, hidden column F, merge A5:C6, frozen 1 × 1 on both sides', () => {
    expect(fixture.wire.collapsedRows).toEqual([{ start: TEXT_GRID.hiddenRow, end: TEXT_GRID.hiddenRow }])
    expect(fixture.wire.collapsedCols).toEqual([{ start: TEXT_GRID.hiddenCol, end: TEXT_GRID.hiddenCol }])
    expect(fixture.wire.merges).toEqual([{ r0: 4, r1: 5, c0: 0, c1: 2 }])
    expect(fixture.wire.frozen).toMatchObject({ rows: 1, cols: 1 })
    expect(fixture.local.sheets[0]!.frozen).toMatchObject({ rows: 1, cols: 1 })
  })

  it('selection parity: keyboard', async () => {
    const steps: Step[] = [
      // Arrow ×3 → Shift+ArrowRight ×2 → Ctrl+ArrowDown → Shift+Ctrl+ArrowRight → PageDown → Home → Ctrl+End → Ctrl+A.
      { kind: 'key', key: 'ArrowDown' }, { kind: 'key', key: 'ArrowDown' }, { kind: 'key', key: 'ArrowRight' },
      { kind: 'key', key: 'ArrowRight', shift: true }, { kind: 'key', key: 'ArrowRight', shift: true },
      { kind: 'key', key: 'ArrowDown', ctrl: true }, { kind: 'key', key: 'ArrowRight', shift: true, ctrl: true },
      { kind: 'key', key: 'PageDown' }, { kind: 'key', key: 'Home' }, { kind: 'key', key: 'End', ctrl: true }, { kind: 'key', key: 'a', ctrl: true },
      { kind: 'key', key: 'Home', ctrl: true }, { kind: 'corner' },
    ]
    // A run over the merge A5:C6, hidden row 7 and hidden column F, under the frozen 1 × 1 (first, at the initial window).
    steps.unshift(
      { kind: 'click', row: 3, col: 0 }, { kind: 'key', key: 'ArrowDown' }, { kind: 'key', key: 'ArrowDown' }, { kind: 'key', key: 'ArrowDown' },
      { kind: 'key', key: 'ArrowRight' }, { kind: 'key', key: 'ArrowRight' }, { kind: 'key', key: 'ArrowRight' }, { kind: 'key', key: 'ArrowRight' }, { kind: 'key', key: 'ArrowRight' },
      { kind: 'key', key: 'ArrowUp', shift: true }, { kind: 'key', key: 'ArrowUp', shift: true }, { kind: 'key', key: 'ArrowLeft', shift: true },
      { kind: 'key', key: 'Escape' },
    )
    const { local, stream } = await parity(steps)
    // The script moved the local viewer (it is not a run of no-ops), through the merge and past the hidden lines.
    expect(local.selections.length).toBeGreaterThan(15)
    expect(local.active).toContain(idOf(4, 0))
    expect(local.active.some(id => rowColOf(id)?.row === TEXT_GRID.hiddenRow)).toBe(false)
    expect(local.active.some(id => rowColOf(id)?.col === TEXT_GRID.hiddenCol)).toBe(false)
    expect(stream.selections).toEqual(local.selections)
    expect(stream.active).toEqual(local.active)
  })

  it('selection parity: mouse', async () => {
    const steps: Step[] = [
      { kind: 'click', row: 2, col: 2 }, { kind: 'click', row: 7, col: 4, shift: true }, { kind: 'click', row: 1, col: 6, ctrl: true },
      { kind: 'header', axis: 'column', index: 1 }, { kind: 'header', axis: 'column', index: 3, shift: true }, { kind: 'header', axis: 'row', index: 9 },
      { kind: 'drag', from: [2, 2], to: [4, 4] },
    ]
    const { local, stream } = await parity(steps)
    expect(local.active[0]).toBe(idOf(2, 2))
    expect(local.selections.length).toBeGreaterThanOrEqual(steps.length)
    expect(stream.selections).toEqual(local.selections)
    expect(stream.active).toEqual(local.active)
  })
})

describe('4b stream resize and fit (P11, P12a)', { timeout: 60_000 }, () => {
  it('resize routes to the server', async () => {
    const s = mountStream()
    const grid = await ready(s.view)
    await waitFor(() => expect(s.calls.length).toBeGreaterThan(0))
    await act(async () => { fireEvent.click(grid, cellPoint(grid, simple(), 1, 1)) })
    expect(cellPart(grid.getAttribute('aria-activedescendant'))).toBe(idOf(1, 1))
    grid.focus()
    await act(async () => { fireEvent.keyDown(grid, { key: 'ArrowRight', altKey: true, shiftKey: true }) })
    await waitFor(() => expect(s.ops!.resize).toHaveBeenCalledTimes(1))
    expect(s.ops!.resize.mock.calls[0]![0]).toEqual({ sheetId: '0', viewGen: 0, axis: 'col', index: 1, size: 72 })
    await waitFor(() => expect(statusText(s.view)).toContain('Column B width 72 pixels. View only; the file is unchanged.'))
    await waitFor(() => expect(s.calls.some(c => c.viewGen === 1)).toBe(true))
    expect(s.workbook.identity('0').viewGen).toBe(1)
    // The viewer's geometry is the server's: column B is 72 px now.
    expect(s.workbook.geometry('0').cols.size(1)).toBe(72)

    // Alt+Shift+Enter: fit, then resize with the fitted size.
    await act(async () => { fireEvent.keyDown(grid, { key: 'Enter', altKey: true, shiftKey: true }) })
    await waitFor(() => expect(s.ops!.resize).toHaveBeenCalledTimes(2))
    expect(s.ops!.fit).toHaveBeenCalledTimes(1)
    expect(s.ops!.fit.mock.calls[0]![0]).toMatchObject({ sheetId: '0', viewGen: 1, axis: 'col', index: 1 })
    expect(s.ops!.fit.mock.calls[0]![0].pane).toMatchObject({ x: 0, y: 0 })
    expect(s.ops!.fit.mock.invocationCallOrder[0]!).toBeLessThan(s.ops!.resize.mock.invocationCallOrder[1]!)
    expect(s.ops!.resize.mock.calls[1]![0]).toEqual({ sheetId: '0', viewGen: 1, axis: 'col', index: 1, size: 120 })
    await waitFor(() => expect(statusText(s.view)).toContain('Column B width 120 pixels, fitted to content. View only; the file is unchanged.'))

    // Reset sizes sends `reset: true`.
    const reset = s.view.getByRole('button', { name: 'Reset sizes' })
    await waitFor(() => expect((reset as HTMLButtonElement).disabled).toBe(false))
    await act(async () => { fireEvent.click(reset) })
    await waitFor(() => expect(s.ops!.resize).toHaveBeenCalledTimes(3))
    expect(s.ops!.resize.mock.calls[2]![0]).toEqual({ sheetId: '0', viewGen: 2, reset: true })
    await waitFor(() => expect(statusText(s.view)).toContain('Column widths and row heights restored to the workbook sizes.'))
  })

  it('key-repeat folds into one pending resize', async () => {
    let release!: (c: SheetViewChange) => void
    const ops = fakeOps()
    ops.resize.mockImplementationOnce(() => new Promise<SheetViewChange>(r => { release = r }))
      .mockImplementationOnce(async () => { throw new StreamSessionError('rate-limited') })
    const s = mountStream({ ops })
    const grid = await ready(s.view)
    await act(async () => { fireEvent.click(grid, cellPoint(grid, simple(), 1, 1)) })
    grid.focus()
    for (let i = 0; i < 5; i++) await act(async () => { fireEvent.keyDown(grid, { key: 'ArrowRight', altKey: true, shiftKey: true }) })
    expect(ops.resize).toHaveBeenCalledTimes(1)
    expect(ops.resize.mock.calls[0]![0]).toMatchObject({ size: 72 })
    await act(async () => { release({ sheetId: '0', viewGen: 1, change: { axis: 'col', index: 1, size: 72 } }) })
    await waitFor(() => expect(ops.resize).toHaveBeenCalledTimes(2))
    expect(ops.resize.mock.calls[1]![0]).toEqual({ sheetId: '0', viewGen: 1, axis: 'col', index: 1, size: 104 })
    await waitFor(() => expect(statusText(s.view)).toContain('Too many requests. Try again shortly.'))
    await act(async () => { await new Promise(r => setTimeout(r, 30)) })
    expect(ops.resize).toHaveBeenCalledTimes(2)
  })

  it('resize refusals speak P12a\'s strings', async () => {
    const ops = fakeOps()
    ops.resize.mockRejectedValueOnce(new StreamSessionError('view-state-limit')).mockRejectedValueOnce(new StreamSessionError('coverage-unknown'))
      .mockRejectedValueOnce(new StreamSessionError('hidden-target'))
    const s = mountStream({ ops })
    const grid = await ready(s.view)
    await act(async () => { fireEvent.click(grid, cellPoint(grid, simple(), 1, 1)) })
    grid.focus()
    const expected = [
      'Column B was not resized: this sheet already has the maximum number of resized columns. Reset sizes to start over.',
      'Column B can\'t be resized in this viewing mode.',
      'Column B is hidden; it cannot be resized.',
    ]
    for (const [i, text] of expected.entries()) {
      await act(async () => { fireEvent.keyDown(grid, { key: 'ArrowRight', altKey: true, shiftKey: true }) })
      await waitFor(() => expect(ops.resize).toHaveBeenCalledTimes(i + 1))
      await waitFor(() => expect(statusText(s.view)).toContain(text))
    }
  })
})

describe('4b stream find (P6, P12a, P15, F15)', { timeout: 60_000 }, () => {
  it('find waits for typing to stop', async () => {
    const r = remoteDoc()
    await ready(r.view)
    const box = await r.view.findByRole('searchbox')
    const word = 'needle'
    let last = 0
    for (let i = 1; i <= word.length; i++) {
      await act(async () => { fireEvent.change(box, { target: { value: word.slice(0, i) } }) })
      last = performance.now()
      if (i < word.length) await act(async () => { await new Promise(res => setTimeout(res, 20)) })
    }
    expect(r.ops!.find).not.toHaveBeenCalled()
    await waitFor(() => expect(r.ops!.find).toHaveBeenCalledTimes(1), { timeout: 2000 })
    const at = performance.now()
    expect(at - last).toBeGreaterThanOrEqual(240)
    expect(r.ops!.find.mock.calls[0]![0]).toEqual({ query: 'needle', matchCase: false })
    await act(async () => { await new Promise(res => setTimeout(res, 400)) })
    expect(r.ops!.find).toHaveBeenCalledTimes(1)
    // The hit is selected (no hit text rects: the cell's selection shows it).
    const grid = r.view.getByRole('grid')
    await waitFor(() => expect(cellPart(grid.getAttribute('aria-activedescendant'))).toBe(idOf(2, 1)))
  })

  it('a truncated answer says "Showing the first 5,000 matches."', async () => {
    const r = remoteDoc({ ops: fakeOps({ find: async () => ({ hits: [{ sheetIndex: 0, row: 4, col: 2 }], total: 1, truncated: true }) }) })
    await ready(r.view)
    await act(async () => { fireEvent.change(await r.view.findByRole('searchbox'), { target: { value: 'x' } }) })
    await waitFor(() => expect(statusText(r.view)).toContain('Showing the first 5,000 matches.'), { timeout: 2000 })
  })

  it('a hit hidden by the filter is said, never inspected (P15)', async () => {
    const r = remoteDoc({ ops: fakeOps({ find: async () => ({ hits: [{ sheetIndex: 0, row: 4, col: 2, hidden: 'filter' }], total: 1, truncated: false }) }) })
    await ready(r.view)
    await act(async () => { fireEvent.change(await r.view.findByRole('searchbox'), { target: { value: 'x' } }) })
    await waitFor(() => expect(statusText(r.view)).toContain('This match is in a row hidden by the filter.'), { timeout: 2000 })
    expect(statusText(r.view)).not.toContain('Original cell')
    expect(r.view.getByRole('grid').getAttribute('aria-activedescendant')).toBeNull()
  })
})

describe('4b stream copy (S3 item 10, Review Focus 2)', { timeout: 60_000 }, () => {
  it('copy off disables Copy and Ctrl+C', async () => {
    const off = remoteDoc({ copy: false })
    const grid = await ready(off.view)
    await act(async () => { fireEvent.click(grid, cellPoint(grid, simple(), 1, 1)) })
    expect(grid.getAttribute('aria-activedescendant')).not.toBeNull()
    expect(off.view.queryByRole('button', { name: 'Copy selection' })).toBeNull()
    await act(async () => { fireEvent.keyDown(grid, { key: 'c', ctrlKey: true }) })
    expect(off.ops!.copy).not.toHaveBeenCalled()
    cleanup()

    // Copy turned off mid-session (a resume's manifest change).
    const on = remoteDoc({ copy: true })
    const g = await ready(on.view)
    await act(async () => { fireEvent.click(g, cellPoint(g, simple(), 1, 1)) })
    expect(on.view.getByRole('button', { name: 'Copy selection' })).toBeTruthy()
    await act(async () => { on.setCan({ copy: false }) })
    await waitFor(() => expect(on.view.queryByRole('button', { name: 'Copy selection' })).toBeNull())
    await act(async () => { fireEvent.keyDown(g, { key: 'c', ctrlKey: true }) })
    expect(on.ops!.copy).not.toHaveBeenCalled()
  })

  it('no local text index', async () => {
    // Control: a local viewer's find does page a WorkbookFind (the spy sees it).
    const local = render(<InteractiveSheetDoc doc={fixture.local} />)
    await ready(local)
    await act(async () => { fireEvent.change(await local.findByRole('searchbox'), { target: { value: 'needle' } }) })
    await waitFor(() => expect(nextPage).toHaveBeenCalled())
    cleanup()
    vi.clearAllMocks()

    const r = remoteDoc()
    const grid = await ready(r.view)
    await act(async () => { fireEvent.change(await r.view.findByRole('searchbox'), { target: { value: 'needle' } }) })
    await waitFor(() => expect(r.ops!.find).toHaveBeenCalledTimes(1), { timeout: 2000 })
    await act(async () => { fireEvent.click(grid, cellPoint(grid, simple(), 1, 1)) })
    await act(async () => { fireEvent.click(r.view.getByRole('button', { name: 'Copy selection' })) })
    await waitFor(() => expect(r.ops!.copy).toHaveBeenCalledTimes(1))
    await waitFor(() => expect(statusText(r.view)).toContain('Selection copied.'))
    expect(r.ops!.copy.mock.calls[0]![0]).toBe('0')
    expect(r.ops!.copy.mock.calls[0]![1]).toBe(0)
    expect(spies.workbookFind).toHaveBeenCalledTimes(0)
    expect(nextPage).toHaveBeenCalledTimes(0)
    expect(spies.copySelection).toHaveBeenCalledTimes(0)
    expect(r.ops!.find).toHaveBeenCalledTimes(1)
    expect(r.ops!.copy).toHaveBeenCalledTimes(1)
    expect(vi.mocked(navigator.clipboard.writeText)).toHaveBeenCalledWith('a\tb')
  })

  it('copy refusals speak S3\'s strings', async () => {
    const ops = fakeOps()
    ops.copy.mockRejectedValueOnce(new StreamSessionError('copy-too-large', 'copy-too-large', { limit: 3, requested: 4 }))
      .mockRejectedValueOnce(new StreamSessionError('copy-quota-exhausted', 'copy-quota-exhausted', { limit: 5, requested: 6 }))
    const r = remoteDoc({ ops, caps: { maxParagraphs: 3, maxSessionParagraphs: 5 } })
    const grid = await ready(r.view)
    await act(async () => { fireEvent.click(grid, cellPoint(grid, simple(), 1, 1)) })
    grid.focus()
    await act(async () => { fireEvent.keyDown(grid, { key: 'c', ctrlKey: true }) })
    await waitFor(() => expect(statusText(r.view)).toContain('Selection too large to copy (limit: 3 paragraphs).'))
    await act(async () => { fireEvent.keyDown(grid, { key: 'c', ctrlKey: true }) })
    await waitFor(() => expect(statusText(r.view)).toContain('Copy limit for this session reached (5 paragraphs).'))
    expect(ops.copy).toHaveBeenCalledTimes(2)
  })
})

describe('4a streams (no sheetView)', { timeout: 60_000 }, () => {
  it('4a streams are unchanged', async () => {
    const s = mountStream({ ops: null, allowCopy: true })
    expect(s.workbook.textOps).toBe(false)
    const grid = await ready(s.view)
    await act(async () => { fireEvent.click(grid, cellPoint(grid, simple(), 1, 1)) })
    grid.focus()
    await act(async () => { fireEvent.keyDown(grid, { key: 'ArrowDown' }); fireEvent.keyDown(grid, { key: 'ArrowRight' }) })
    expect(grid.getAttribute('aria-activedescendant')).toBeNull()
    expect(grid.getAttribute('aria-multiselectable')).toBeNull()
    expect(s.view.container.querySelectorAll('[role="gridcell"][aria-selected]')).toHaveLength(0)
    expect(s.view.queryByRole('searchbox')).toBeNull()
    expect(s.view.queryByRole('button', { name: /copy/i })).toBeNull()
    expect(s.view.queryByRole('button', { name: 'Reset sizes' })).toBeNull()
  })
})
