import { describe, expect, it } from 'vitest'
import { frozenPaneFrom } from '../src/xlsx/frozenPane'
import { paneRegions, visibleSpans, ensureAxisVisible, FROZEN_VIEWPORT_FRACTION } from '../src/xlsx/view/frozenPanes'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createScrollWindow, hitTestViewport, logicalToViewport, navigateKey, navigateTo, setScrollOffset, setScrollZoom, viewportToLogical } from '../src/xlsx/view/scrollWindow'

const visibility = visibilityFromHidden(createIntervalSet([]))
const metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order: createIdentityOrder(), visibility })
const viewport = { width: 640, height: 400 }
const view = { visibility, baseHidden: createIntervalSet([]), filter: null, selection: { ranges: [], active: null, anchor: null } } as never
const frozenB2 = { rows: 1, cols: 1 }

describe('authored frozen panes: parsing', () => {
  it('reads frozen and frozenSplit panes, ignores plain splits and bad counts', () => {
    expect(frozenPaneFrom({ xSplit: '1', ySplit: '1', topLeftCell: 'B2', state: 'frozen' })).toEqual({ rows: 1, cols: 1, topRow: 1, leftCol: 1 })
    expect(frozenPaneFrom({ ySplit: '3', topLeftCell: 'A20', state: 'frozenSplit' })).toEqual({ rows: 3, cols: 0, topRow: 19 })
    expect(frozenPaneFrom({ xSplit: '1200', ySplit: '800', state: 'split' })).toBeUndefined()
    expect(frozenPaneFrom({ xSplit: '1', ySplit: '1' })).toBeUndefined()
    expect(frozenPaneFrom({ xSplit: '-2', ySplit: 'x', state: 'frozen' })).toBeUndefined()
    // A top-left cell inside the frozen strip is not a scroll position.
    expect(frozenPaneFrom({ ySplit: '4', topLeftCell: 'A2', state: 'frozen' })).toEqual({ rows: 4, cols: 0 })
  })
})

describe('frozen pane viewport math', () => {
  it('pins the leading strips: frozen cells map with a zero origin, the body keeps the scroll origin', () => {
    const window = setScrollOffset(createScrollWindow(metrics, viewport, { frozen: frozenB2 }), metrics, viewport, { x: 300, y: 400 })
    expect(window.frozen).toEqual({ rows: 1, cols: 1, left: 64, top: 20 })
    expect(logicalToViewport(window, { x: 0, y: 0 })).toEqual({ x: 0, y: 0 })
    expect(logicalToViewport(window, { x: 364, y: 420 })).toEqual({ x: 64, y: 20 })
    expect(viewportToLogical(window, { x: 10, y: 10 })).toEqual({ x: 10, y: 10 })
    expect(viewportToLogical(window, { x: 100, y: 100 })).toEqual({ x: 400, y: 500 })
    // Hit tests: A1 in the corner, row 1 over a scrolled column, column A over a scrolled row.
    expect(hitTestViewport(window, metrics, viewport, { x: 10, y: 10 })).toEqual({ row: 0, col: 0 })
    expect(hitTestViewport(window, metrics, viewport, { x: 70, y: 10 })).toEqual({ row: 0, col: 5 })
    expect(hitTestViewport(window, metrics, viewport, { x: 10, y: 30 })).toEqual({ row: 21, col: 0 })
    expect(hitTestViewport(window, metrics, viewport, { x: 70, y: 30 })).toEqual({ row: 21, col: 5 })
  })

  it('at origin 0 the body starts right after the strip (no duplicate row/column)', () => {
    const window = createScrollWindow(metrics, viewport, { frozen: frozenB2 })
    expect(hitTestViewport(window, metrics, viewport, { x: 64, y: 20 })).toEqual({ row: 1, col: 1 })
    expect(paneRegions(window.frozen, window.logicalOrigin, window.zoom, viewport)).toEqual([
      { kind: 'body', clip: { x: 64, y: 20, width: 576, height: 380 }, origin: { x: 0, y: 0 } },
      { kind: 'top', clip: { x: 64, y: 0, width: 576, height: 20 }, origin: { x: 0, y: 0 } },
      { kind: 'left', clip: { x: 0, y: 20, width: 64, height: 380 }, origin: { x: 0, y: 0 } },
      { kind: 'corner', clip: { x: 0, y: 0, width: 64, height: 20 }, origin: { x: 0, y: 0 } },
    ])
    expect(visibleSpans(window.frozen, { x: 300, y: 400 }, 1, viewport)).toEqual({ x: [[0, 64], [364, 940]], y: [[0, 20], [420, 800]] })
  })

  it('Go To puts a body target at the top-left of the scrolling body and keeps the origin for a frozen target', () => {
    const start = setScrollOffset(createScrollWindow(metrics, viewport, { frozen: frozenB2 }), metrics, viewport, { x: 300, y: 400 })
    const context = { metrics, viewport, view }
    const far = navigateTo(start, context, { addressSpace: 'source', at: { row: 49, col: 9 } })
    expect(far.window.logicalOrigin).toEqual({ x: 9 * 64 - 64, y: 49 * 20 - 20 })
    expect(logicalToViewport(far.window, { x: 9 * 64, y: 49 * 20 })).toEqual({ x: 64, y: 20 })
    const corner = navigateTo(start, context, { addressSpace: 'source', at: { row: 0, col: 0 } })
    expect(corner.window.logicalOrigin).toEqual({ x: 300, y: 400 })
  })

  it('keyboard navigation scrolls a cell out from under the strip, never for frozen cells', () => {
    const start = setScrollOffset(createScrollWindow(metrics, viewport, { frozen: frozenB2 }), metrics, viewport, { x: 0, y: 400 })
    const context = { metrics, viewport, view }
    // Row 21 (y 420) is the first body row at origin 400; ArrowUp to row 20 (y 400) must scroll to 380.
    const up = navigateKey(start, context, 'ArrowUp', { row: 21, col: 3 })
    expect(up.target).toEqual({ row: 20, col: 3 })
    expect(up.window.logicalOrigin.y).toBe(380)
    const intoFrozen = navigateKey(setScrollOffset(start, metrics, viewport, { x: 0, y: 0 }), context, 'ArrowUp', { row: 1, col: 3 })
    expect(intoFrozen.target).toEqual({ row: 0, col: 3 })
    expect(intoFrozen.window.logicalOrigin.y).toBe(0)
    expect(ensureAxisVisible(400, 0, 20, 400, 20)).toBe(400)
  })

  it('does not pin a strip wider than the pane allows, and re-evaluates on zoom', () => {
    const wide = createScrollWindow(metrics, viewport, { frozen: { rows: 0, cols: 7 } })
    expect(7 * 64 > viewport.width * FROZEN_VIEWPORT_FRACTION).toBe(true)
    expect(wide.frozen?.left).toBe(0)
    expect(hitTestViewport(setScrollOffset(wide, metrics, viewport, { x: 640, y: 0 }), metrics, viewport, { x: 10, y: 10 })?.col).toBe(10)
    const zoomedOut = setScrollZoom(wide, metrics, viewport, 0.5, { x: 0, y: 0 })
    expect(zoomedOut.frozen?.left).toBe(7 * 64)
  })

  it('a zoom anchored inside the strip keeps the scroll origin', () => {
    const window = setScrollOffset(createScrollWindow(metrics, viewport, { frozen: frozenB2 }), metrics, viewport, { x: 300, y: 400 })
    const zoomed = setScrollZoom(window, metrics, viewport, 2, { x: 10, y: 10 })
    expect(zoomed.logicalOrigin).toEqual({ x: 300, y: 400 })
  })
})

describe('authored frozen panes: both adapters', async () => {
  const JSZip = (await import('jszip')).default
  const { buildXlsx } = await import('../src/testdata/ooxml-builders')
  const { OfficePackage } = await import('../src/core/zip')
  const { parseXlsx } = await import('../src/xlsx/parse')
  const { createObjectWorkbookSource } = await import('../src/xlsx/view/source')
  const { createPagedWorkbookSource } = await import('../src/xlsx/view/pagedSource')
  const { openImportSession } = await import('../src/xlsx/storage/pagedImport')
  async function frozenWorkbook(): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }, { r: 30, cells: [{ ref: 'C30', v: 2 }] }] }]))
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    zip.file('xl/worksheets/sheet1.xml', xml.replace('<sheetData>', '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="C12" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="C12"/></sheetView></sheetViews><sheetData>'))
    return zip.generateAsync({ type: 'uint8array' })
  }
  it('the resident parser, object source and paged import report the same panes', async () => {
    const bytes = await frozenWorkbook()
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const expected = { rows: 2, cols: 1, topRow: 11, leftCol: 2 }
    expect(doc.sheets[0].frozen).toEqual(expected)
    expect(createObjectWorkbookSource(doc, 'R').geometry(doc.sheets[0].sheetId ?? '1').frozen).toEqual(expected)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'P', profile: 'D' })
    expect((await session.done).kind).toBe('complete')
    const paged = (createPagedWorkbookSource as (...args: unknown[]) => ReturnType<typeof createPagedWorkbookSource>)(session)
    expect(paged.geometry('1').frozen).toEqual(expected)
  })
})
