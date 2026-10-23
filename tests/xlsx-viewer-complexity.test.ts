/**
 * S10.1 complexity audit (spec §3 invariants 1–7; AC-CAP-01, AC-VIEW-03,
 * AC-RENDER-01, AC-PERF-02 "bounded complexity").
 *
 * The instruments sit at boundaries, not inside the implementation:
 *  - an allocation ledger over the global typed-array constructors, `Array(n)`
 *    and Map#set / Set#add on officeview-built collections;
 *  - the V8 heap retained after a full collection across each warmed-up window,
 *    which is what bounds plain objects and literal arrays (no constructor to
 *    proxy): a per-row/per-cell object table or per-row array set stays
 *    reachable and exceeds the bound (self-checked with the review probe).
 *    Residual: transient garbage freed inside the window is not bounded here;
 *  - canvas backing-store sizes (no full-sheet canvas) and the owned DOM size;
 *  - the worker port: every command, every read rect (no all-column read per
 *    row, no unbounded row band) and the kernel's parse spy (no per-command
 *    reparse or workbook clone);
 *  - the source store of the F6 full-axis provider (only the window is read).
 * Bounds are derived from the viewport/window, never from the implementation's
 * internal counts; the "would be" figures name what each anti-pattern costs.
 */
import v8 from 'node:v8'
import vm from 'node:vm'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createElement } from 'react'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { CacheIdentity, SourceStore, WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import type { SheetFrameStats } from '../src/components/xlsx/InteractiveSheetDoc'
import { queryViewport, VIEWPORT_DETAIL_LIMITS } from '../src/xlsx/view/viewport'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createFixture } from './helpers/xlsx-viewer-fixtures'
import { xlsxBytes } from './helpers/xlsx-package'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

// The object path admits one row, so the worker-backed tests below are paged
// on the main thread and in the worker alike (as in the S5.4 suite).
vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})

afterEach(() => { cleanup(); vi.restoreAllMocks() })

// ---------------------------------------------------------------- instruments

const TYPED = ['Float64Array', 'Float32Array', 'Uint32Array', 'Int32Array', 'Uint16Array', 'Int16Array', 'Uint8Array', 'Uint8ClampedArray', 'Int8Array', 'BigInt64Array', 'BigUint64Array'] as const
interface Ledger { typedElements: number; largestTyped: number; inserts: number; arrayElements: number; retainedBytes: number }
// Plain objects and literal arrays have no constructor to proxy, so their cost
// is bounded by what stays reachable: the V8 heap after a full collection,
// before and after `run`, while the viewer stays mounted. A per-coordinate
// object table or per-row array set survives the window and shows up here.
v8.setFlagsFromString('--expose-gc')
/** Null where the runtime refuses to expose a collector: the retained-heap
 * assertions are then skipped with a visible reason, never silently passed. */
const collect: (() => void) | null = (() => {
  try { return vm.runInNewContext('gc') as () => void } catch { return null }
})()
const HEAP_SKIP = 'retained-heap bound skipped: the runtime does not expose gc'
function settledHeap(): number {
  if (!collect) return 0
  for (let i = 0; i < 3; i++) collect()
  return process.memoryUsage().heapUsed
}
/** Asserts the retained-heap bound when a collector exists. */
function expectRetainedBelow(ledger: Ledger, bound: number): void {
  if (collect) expect(ledger.retainedBytes).toBeLessThan(bound)
}
/** Counts typed-array elements, `Array(n)` lengths and Map/Set insertions while
 * `run` executes, and the heap it leaves retained. */
async function measure(run: () => Promise<void>): Promise<Ledger> {
  const ledger: Ledger = { typedElements: 0, largestTyped: 0, inserts: 0, arrayElements: 0, retainedBytes: 0 }
  const heapBefore = settledHeap()
  const scope = globalThis as unknown as Record<string, unknown>
  const saved = TYPED.map(name => [name, scope[name]] as const)
  for (const [name, ctor] of saved) {
    scope[name] = new Proxy(ctor as new (...args: unknown[]) => ArrayLike<unknown>, {
      construct(target, args, newTarget) {
        const made = Reflect.construct(target, args, newTarget === scope[name] ? target : newTarget)
        ledger.typedElements += made.length
        ledger.largestTyped = Math.max(ledger.largestTyped, made.length)
        return made
      },
    })
  }
  const savedArray = Array
  scope.Array = new Proxy(Array, {
    construct(target, args, newTarget) {
      const made = Reflect.construct(target, args, newTarget === scope.Array ? target : newTarget) as unknown[]
      ledger.arrayElements += made.length
      return made
    },
    apply(target, thisArg, args) {
      const made = Reflect.apply(target, thisArg, args) as unknown[]
      ledger.arrayElements += made.length
      return made
    },
  })
  const mapSet = Map.prototype.set, setAdd = Set.prototype.add
  // Only collections constructed by officeview's own modules count: jsdom's CSS
  // engine and React keep their own maps, which a real browser does not run in JS.
  const owned = new WeakSet<object>()
  // The caller is read from structured call sites, never a formatted `.stack`:
  // the runner's source-map support decodes and caches a file's whole map the
  // first time a formatted stack crosses it, and that cache (MiBs, growing with
  // every source file the window touches) would be charged to the measured code.
  const ours = () => {
    const prepare = Error.prepareStackTrace
    Error.prepareStackTrace = (_error, sites) => sites
    try {
      const file = (new Error().stack as unknown as NodeJS.CallSite[] | undefined)?.[2]?.getFileName() ?? ''
      return file.includes('/src/') && !file.includes('node_modules')
    } finally { Error.prepareStackTrace = prepare }
  }
  const savedCollections = [['Map', Map], ['Set', Set]] as const
  for (const [name, ctor] of savedCollections) {
    scope[name] = new Proxy(ctor, {
      construct(target, args, newTarget) {
        const made = Reflect.construct(target, args, newTarget === scope[name] ? target : newTarget) as object
        if (ours()) { owned.add(made); ledger.inserts += (made as Map<unknown, unknown>).size }
        return made
      },
    })
  }
  Map.prototype.set = function (this: Map<unknown, unknown>, key, value) { if (owned.has(this)) ledger.inserts++; return mapSet.call(this, key, value) }
  Set.prototype.add = function (this: Set<unknown>, value) { if (owned.has(this)) ledger.inserts++; return setAdd.call(this, value) }
  const traceLimit = Error.stackTraceLimit
  Error.stackTraceLimit = 4 // the caller frame is all `ours` reads
  try { await run() } finally {
    Error.stackTraceLimit = traceLimit
    for (const [name, ctor] of saved) scope[name] = ctor
    for (const [name, ctor] of savedCollections) scope[name] = ctor
    scope.Array = savedArray
    Map.prototype.set = mapSet
    Set.prototype.add = setAdd
  }
  ledger.retainedBytes = Math.max(0, settledHeap() - heapBefore)
  return ledger
}

/** Largest canvas backing store and owned DOM size inside a container. */
function surfaces(root: HTMLElement): { maxCanvasPixels: number; nodes: number; gridcells: number } {
  const canvases = [...root.querySelectorAll('canvas')]
  return {
    maxCanvasPixels: Math.max(0, ...canvases.map(canvas => canvas.width * canvas.height)),
    nodes: root.querySelectorAll('*').length,
    gridcells: root.querySelectorAll('[role="gridcell"]').length,
  }
}

function frames() {
  const all: SheetFrameStats[] = []
  return { all, onFrame: (stats: SheetFrameStats) => { all.push(stats) }, last: () => all[all.length - 1] }
}
const internal = (props: Record<string, unknown>) => props as object
function go(ui: ReturnType<typeof render>, at: string, space: 'display' | 'source' = 'display') {
  fireEvent.change(ui.getByRole('combobox', { name: 'Address space' }), { target: { value: space } })
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: at } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  const grid = ui.getByRole('grid')
  act(() => grid.focus())
  return grid
}
const activeId = (ui: ReturnType<typeof render>) => {
  const id = ui.getByRole('grid').getAttribute('aria-activedescendant')
  return id ? Number(/-cell-(\d+)$/.exec(id)?.[1]) : null
}

/** A port with dedicated-worker semantics over a real kernel (structured clones). */
function workerHarness() {
  const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
  const kernel = new ViewerKernel({ parse }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  const posted: WorkerCommand[] = []
  const port = {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
      posted.push(command)
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const data = structuredClone(out, { transfer: replyTransfer })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
  return { port, parse, posted }
}
/** Every read rect the viewer asked the worker for. */
function readRects(posted: readonly WorkerCommand[]) {
  const out: Array<{ op: string; r0: number; r1: number; c0: number; c1: number }> = []
  for (const command of posted) {
    if (command.kind !== 'read') continue
    const query = command.query as { op: string; range?: { r0: number; r1: number; c0: number; c1: number }; rect?: { r0: number; r1: number; c0: number; c1: number } }
    const rect = query.range ?? query.rect
    if (rect) out.push({ op: query.op, ...rect })
  }
  return out
}

// ------------------------------------------------------------------ fixtures

const inline = (ref: string, text: string) => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`
/** One-cell sparse corners: A1 and XFD1048576 only (CAP-02 "one-cell sparse sheet"). */
const CORNERS = `<row r="1">${inline('A1', 'origin')}</row><row r="1048576">${inline('XFD1048576', 'corner')}</row>`
/** Rows just past the 2^31 and 2^32 cell-ID boundaries (AC-CAP-01). */
const BOUNDARIES = `<row r="1">${inline('A1', 'origin')}</row><row r="131073">${inline('A131073', 'two-pow-31')}</row><row r="262145">${inline('A262145', 'two-pow-32')}</row>`
async function objectDoc(sheetData: string, afterSheetData = '') {
  return parseXlsx(await OfficePackage.load(xlsxBytes({ sheetData, ...(afterSheetData ? { afterSheetData } : {}) })))
}
function table(rows: number): string {
  const out = [`<row r="1">${inline('A1', 'id')}${inline('B1', 'score')}${inline('C1', 'name')}</row>`]
  for (let r = 2; r <= rows + 1; r++) out.push(`<row r="${r}"><c r="A${r}"><v>${r}</v></c><c r="B${r}"><v>${(r * 7919) % 1000}</v></c>${inline(`C${r}`, `n${r % 17}`)}</row>`)
  return out.join('')
}

// A row-sized typed lane is MAX_ROWS (1,048,576) elements and a dense cell
// lane would be 2^34; a per-row/per-cell Map or Set needs one insert per
// coordinate. The bounds below sit far under the first and are generous for
// the viewport's own working set (tiles, visible records, React).
const ROW_LANE = MAX_ROWS
const VIEWPORT_ALLOCATION = 1 << 19
const VIEWPORT_INSERTS = 50_000
// Retained heap after a warmed-up window (jsdom/React noise settles near
// 0.3 MiB per gesture round; the largest viewport result, 2,048 coarse groups,
// retains ~0.78 MB; it read ~2 MiB while the ledger's formatted stacks charged
// the runner's source-map cache to the window). The review probe below, a 65,536-entry plain-object cell table plus
// 1,024 plain row arrays, retains ~7 MiB and must exceed it; a per-row table over
// the 1,048,576-row axis would be larger still.
const VIEWPORT_RETAINED = 4 * 1024 * 1024

describe('instrument self-check (the ledger is not vacuous)', () => {
  // Where gc is unavailable, this skipped entry names why the heap bounds did not run.
  if (!collect) it.skip(HEAP_SKIP, () => {})
  it('sees a row-sized typed lane and per-entry inserts into an officeview-built collection', async () => {
    const ledger = await measure(async () => {
      void new Uint32Array(ROW_LANE)
      // 60,000 single-row size overrides: axes.ts builds a boundary Set with one entry per edge.
      createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20, overrides: Array.from({ length: 60_000 }, (_, i) => ({ start: i * 2, end: i * 2, sizeCss: 30 })) })
    })
    expect(ledger.largestTyped).toBe(ROW_LANE)
    expect(ledger.inserts).toBeGreaterThanOrEqual(VIEWPORT_INSERTS)
  })

  it.skipIf(!collect)('sees a retained plain-object cell table and plain row arrays (no constructor to proxy)', async () => {
    const kept: unknown[] = []
    const ledger = await measure(async () => {
      const cells: Record<string, { v: number }> = {}
      for (let i = 0; i < 65_536; i++) cells[`c${i}`] = { v: i }
      const rows: number[][] = []
      for (let r = 0; r < 1024; r++) { const row: number[] = []; for (let c = 0; c < 64; c++) row.push(r * 64 + c + 0.5); rows.push(row) }
      kept.push(cells, rows)
    })
    expect(ledger.inserts).toBe(0)
    expect(ledger.typedElements).toBe(0)
    expect(ledger.retainedBytes).toBeGreaterThan(VIEWPORT_RETAINED)
    expect(kept).toHaveLength(2)
  })
})

describe('resident (object) viewer: navigation and paint stay viewport-bounded', () => {
  it('Go To XFD1048576 on a one-cell corner sheet allocates no row lane, map or full-sheet canvas', async () => {
    const doc = await objectDoc(CORNERS)
    const f = frames()
    const ui = render(createElement(OfficeDoc, { document: doc, interactiveSheets: true, ...internal({ onFrame: f.onFrame }) }))
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /origin/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    // Warm-up: the same gestures three times, so first-use module/JIT state and
    // the bounded tile/layout caches are full before the measured window.
    // The warm-up gestures are cold first gestures (module, JIT and cache fill), so they take the first-paint wait.
    // Under load a cold warm-up Go To was seen to be dropped outright (PR #19 round 10 CI and a local loop of 16
    // parallel runs: the selection never moved, aria-activedescendant stayed unset for 29 s), so a warm-up gesture
    // that has not landed is sent once more before the test gives up. The measured window below is unchanged.
    const reach = async (at: string, landed: () => void) => {
      go(ui, at)
      try { await waitFor(landed, { timeout: FIRST_PAINT_WAIT.timeout / 2 }) }
      catch { go(ui, at); await waitFor(landed, FIRST_PAINT_WAIT) }
    }
    for (let round = 0; round < 3; round++) {
      await reach('XFD1048576', () => expect(ui.getByRole('gridcell', { name: /corner/ })).toBeTruthy())
      await reach('A1048576', () => expect(activeId(ui)).toBe((MAX_ROWS - 1) * MAX_COLS))
      await reach('A1', () => expect(ui.getByRole('gridcell', { name: /origin/ })).toBeTruthy())
    }
    const ledger = await measure(async () => {
      go(ui, 'XFD1048576')
      await waitFor(() => expect(ui.getByRole('gridcell', { name: /corner/ })).toBeTruthy())
      go(ui, 'A1048576')
      await waitFor(() => expect(activeId(ui)).toBe((MAX_ROWS - 1) * MAX_COLS))
    })
    expect(ledger.largestTyped).toBeLessThan(ROW_LANE)
    expect(ledger.typedElements).toBeLessThan(VIEWPORT_ALLOCATION)
    expect(ledger.inserts).toBeLessThan(VIEWPORT_INSERTS)
    expectRetainedBelow(ledger, VIEWPORT_RETAINED)
    const s = surfaces(ui.container)
    const pane = f.last()!
    // A full-sheet canvas would be 2^34 x 1280 px; every backing store here is a pane or tile.
    expect(s.maxCanvasPixels).toBeLessThanOrEqual(4096 * 4096)
    expect(pane.peakBitmapBytes).toBeLessThanOrEqual(pane.budgetBytes)
    expect(s.gridcells).toBeLessThan(2000)
    expect(s.nodes).toBeLessThan(5000)
    // Exact identity at the last coordinate: 2^34 - 1, not a 32-bit alias.
    go(ui, 'XFD1048576')
    await waitFor(() => expect(activeId(ui)).toBe(2 ** 34 - 1))
  }, 60_000)
  // Timing-only budget, documented (C1 flake hardening): this test asserts allocation bounds, never time, but its
  // run time is pure CPU work (real React renders of the grid in jsdom: three warm-up rounds of three Go To gestures,
  // then the measured window), with no timer or clock to drive. It takes ~2.6 s alone on the Mac and timed out at
  // 15 s under full-suite load (10.7 s observed); a 2-CPU Linux runner is about 2x slower again. 60 s keeps a wide
  // margin while still failing on a viewer that never paints. The measured-window waits stay at the default: by then
  // every module and cache is warm.

  it('cell IDs across the 2^31 and 2^32 boundaries stay exact end to end (no bitwise packing)', FIRST_PAINT_TEST, async () => {
    const doc = await objectDoc(BOUNDARIES)
    const ui = render(createElement(OfficeDoc, { document: doc, interactiveSheets: true }))
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /origin/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    go(ui, 'A131073')
    await waitFor(() => expect(ui.getByRole('gridcell', { name: 'A131073, value two-pow-31' })).toBeTruthy())
    expect(activeId(ui)).toBe(2 ** 31) // a signed 32-bit result would be negative
    go(ui, 'A262145')
    await waitFor(() => expect(ui.getByRole('gridcell', { name: 'A262145, value two-pow-32' })).toBeTruthy())
    expect(activeId(ui)).toBe(2 ** 32) // 32-bit packing would alias row 0 (id 0)
  })

  it('a giant narrow-column merge (D1:D1048576) paints and refuses a sort without a coverage map', async () => {
    const doc = await objectDoc(table(40), '<mergeCells count="1"><mergeCell ref="D1:D1048576"/></mergeCells>')
    const ui = render(createElement(OfficeDoc, { document: doc, interactiveSheets: true }))
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /^A1, value id/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    const ledger = await measure(async () => {
      go(ui, 'D524288')
      await waitFor(() => expect(ui.getByRole('grid').getAttribute('aria-activedescendant')).toBeTruthy())
      go(ui, 'A1')
      fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
      await waitFor(() => expect(ui.container.textContent).toMatch(/Operation refused: merge/))
    })
    // A covered-cell map of this merge alone would hold 1,048,576 entries.
    expect(ledger.inserts).toBeLessThan(VIEWPORT_INSERTS)
    expectRetainedBelow(ledger, VIEWPORT_RETAINED)
    expect(ledger.largestTyped).toBeLessThan(ROW_LANE)
  }, 15_000) // ~0.8 s alone (2.4 s with formatted stacks); margin for suite load
})

describe('worker-backed viewer: bounded reads, one parse, no per-command reparse or clone', () => {
  // No warm-up: the retained-heap gate is the whole-file run, where the earlier
  // tests have loaded the shared modules (~0.5 MB retained here). Run alone with
  // `-t` in a fresh process, this window also holds cold first-use module state
  // (~7.5-8 MB) and is expected to exceed VIEWPORT_RETAINED.
  it('open, corner navigation and a sort issue bounded reads and never re-open or re-parse the workbook', async () => {
    const h = workerHarness()
    const blob = new Blob([xlsxBytes({ sheetData: table(3000) + `<row r="1048576">${inline('XFD1048576', 'corner')}</row>` }) as BlobPart])
    const ui = render(createElement(OfficeFile, { data: blob, interactiveSheets: true, viewerWorker: () => h.port }))
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /^A1, value id/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    const ledger = await measure(async () => {
      go(ui, 'XFD1048576')
      await waitFor(() => expect(ui.getByRole('gridcell', { name: /corner/ })).toBeTruthy())
      go(ui, 'A1')
      await waitFor(() => expect(ui.getByRole('gridcell', { name: /^A1, value id/ })).toBeTruthy())
      fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
      await waitFor(() => expect(ui.container.textContent).toMatch(/View-only sort applied/), { timeout: 15_000 })
    })
    const kinds = h.posted.map(command => command.kind)
    expect(kinds.filter(kind => kind === 'open')).toHaveLength(1)
    expect(h.parse).not.toHaveBeenCalled() // paged import; never a whole-workbook object parse
    expect(kinds.filter(kind => kind === 'sort')).toHaveLength(1)
    // Each painted row read is a bounded band over a bounded column span: never
    // the 16,384-column row nor the 1,048,576-row column.
    for (const rect of readRects(h.posted).filter(r => r.op === 'rows')) {
      expect(rect.c1 - rect.c0 + 1).toBeLessThanOrEqual(1024)
      expect(rect.r1 - rect.r0 + 1).toBeLessThanOrEqual(4096)
    }
    // The main thread allocated no row lane for the corner or the sort (the sort's
    // permutation lives in the worker; the adopted view is band-sized, 3,001 rows).
    expect(ledger.largestTyped).toBeLessThan(ROW_LANE)
    expect(ledger.inserts).toBeLessThan(VIEWPORT_INSERTS)
    expectRetainedBelow(ledger, VIEWPORT_RETAINED)
  }, 30_000)
})

describe('F6 full-axis provider: the geometry window reads only what it shows', () => {
  const owner = (store: SourceStore): CacheIdentity => ({ docId: store.stamp.docId, sheetId: store.stamp.sheetId, valueGen: store.stamp.valueGen, fontGen: 0, viewGen: 0 })
  const metricsOf = (rowCss: number, colCss: number) => createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: rowCss }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: colCss }) })
  function counted() {
    const fixture = createFixture('F6')
    const counts = { cellAt: 0, isPopulated: 0 }
    const store = fixture.createStore({ read: {
      cellAt: (row, col) => { counts.cellAt++; return fixture.cellAt(row, col) },
      isPopulated: (row, col) => { counts.isPopulated++; return fixture.isPopulated(row, col) },
    } })
    return { store, counts }
  }

  it('detail at XFD1048576: visits and allocations scale with the 1440x900 window, not 2^34', async () => {
    const { store, counts } = counted()
    const metrics = metricsOf(20, 64)
    const box = { x: metrics.cols.totalCss() - 1440, y: metrics.rows.totalCss() - 900, width: 1440, height: 900 }
    let outcome: Awaited<ReturnType<typeof queryViewport>> | undefined
    const ledger = await measure(async () => { outcome = await queryViewport(store, metrics, { box, zoom: 1, deviceScale: 2 }, owner(store), { signal: new AbortController().signal }) })
    expect(outcome?.kind).toBe('viewport')
    const window = Math.ceil(900 / 20 + 1) * Math.ceil(1440 / 64 + 1) // 46 x 24 visible cells
    expect(counts.cellAt + counts.isPopulated).toBeLessThanOrEqual(2 * window)
    expect(ledger.largestTyped).toBeLessThan(ROW_LANE)
    expect(ledger.typedElements).toBeLessThan(VIEWPORT_ALLOCATION)
    expect(ledger.inserts).toBeLessThan(VIEWPORT_INSERTS)
    expectRetainedBelow(ledger, VIEWPORT_RETAINED)
  })

  it('tiny sizes at 25% zoom: the whole pane is analytic coarse groups with no per-cell visit', async () => {
    const { store, counts } = counted()
    const metrics = metricsOf(0.5, 1) // 7,200 x 5,760 logical cells fit the pane at 25%
    let outcome: Awaited<ReturnType<typeof queryViewport>> | undefined
    const ledger = await measure(async () => { outcome = await queryViewport(store, metrics, { box: { x: 0, y: 0, width: 1440 * 4, height: 900 * 4 }, zoom: 0.25, deviceScale: 2 }, owner(store), { signal: new AbortController().signal }) })
    expect(outcome?.kind).toBe('viewport')
    if (outcome?.kind !== 'viewport') return
    expect(outcome.groups.length).toBeLessThanOrEqual(VIEWPORT_DETAIL_LIMITS.D.overviewGroups)
    // 41 million covered cells; the summary visits none of them one by one.
    expect(counts.cellAt).toBe(0)
    expect(ledger.inserts).toBeLessThan(VIEWPORT_INSERTS)
    expectRetainedBelow(ledger, VIEWPORT_RETAINED)
    expect(ledger.typedElements).toBeLessThan(VIEWPORT_ALLOCATION)
  })
})
