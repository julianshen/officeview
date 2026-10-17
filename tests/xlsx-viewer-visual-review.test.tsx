/**
 * PR #11 visual review, stream A (interactive viewer UX): regressions for the
 * header strip (#9, #21), the status line (#18), the find overlay (#20), the
 * popup/toolbar layout (#13), the chrome font (#19) and spreadsheet print (#12).
 * Real-browser layout evidence lives in tests/browser/interactive-viewer-visual.spec.ts.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { GridHeaders, GridRowHeader, HEADER_COLORS, rowGutterWidth, GRID_HEADER_SIZE } from '../src/components/xlsx/GridHeaders'
import { GridSemantics, STATUS_BAR_HEIGHT } from '../src/components/xlsx/GridSemantics'
import { paintFindOverlay } from '../src/components/xlsx/selectionOverlay'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { workbookOf } from './helpers/xlsx-find'
import { gridFixture } from './helpers/xlsx-grid-fixture'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'
import type { Paintable, PaintableArray } from '../src/render/paint'

afterEach(() => { cleanup(); vi.restoreAllMocks() })

function luminance(hex: string): number {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255).map(c => c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4)
  return 0.2126 * r + 0.7152 * g + 0.0722 * b
}
const contrast = (a: string, b: string) => { const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x); return (hi + 0.05) / (lo + 0.05) }

describe('#21 header contrast', () => {
  it('header labels, action glyphs and active state meet WCAG AA on the header strip; tracks are >= 3:1', () => {
    expect(contrast(HEADER_COLORS.text, HEADER_COLORS.background)).toBeGreaterThanOrEqual(4.5)
    expect(contrast(HEADER_COLORS.action, HEADER_COLORS.background)).toBeGreaterThanOrEqual(4.5)
    expect(contrast(HEADER_COLORS.active, HEADER_COLORS.background)).toBeGreaterThanOrEqual(4.5)
    expect(contrast(HEADER_COLORS.track, HEADER_COLORS.background)).toBeGreaterThanOrEqual(3)
  })
  it('header buttons carry the explicit label colour instead of the UA button colour', () => {
    const props = gridFixture()
    const ui = render(<GridHeaders projection={props.projection} window={props.window} viewport={props.viewport} columns={[0]} selection={props.selection} sort={null} onSortCycle={() => {}} onFilter={() => {}} />)
    expect(ui.getByRole('button', { name: /^Select Column A/ }).style.color).toBe(rgb(HEADER_COLORS.text))
  })
})
const rgb = (hex: string) => `rgb(${[1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16)).join(', ')})`

describe('#9 header actions and labels', () => {
  const headers = (zoom: number, coarse = false) => {
    const props = gridFixture()
    const window = createScrollWindow(props.projection, props.viewport, { zoom })
    return render(<GridHeaders projection={props.projection} window={window} viewport={props.viewport} columns={[0, 2]} selection={props.selection} sort={null}
      touchControls={coarse} controlMode onSortCycle={() => {}} onFilter={() => {}} onSortDialog={() => {}} />)
  }
  it('a wide column shows its label at full width beside compact action glyphs (never four equal slices)', () => {
    const ui = headers(1)
    const label = ui.getByRole('button', { name: /^Select Column A/ }), sort = ui.getByRole('button', { name: 'Cycle sort column A' })
    expect(label.style.flex).toBe('1 1 auto')
    expect(sort.style.position).not.toBe('absolute')
    expect(parseFloat(sort.style.width)).toBeLessThanOrEqual(16)
  })
  it('at 50% zoom the action glyphs are hidden visually but stay in the accessibility tree and keyboard order', () => {
    const ui = headers(0.5)
    const sort = ui.getByRole('button', { name: 'Cycle sort column A' }), filter = ui.getByRole('button', { name: /^Filter column A/ })
    for (const button of [sort, filter]) {
      expect(button.style.position).toBe('absolute')
      expect(button.style.width).toBe('1px')
      expect(button.tabIndex).toBe(0)
    }
    // Focus (control mode) reveals the actions: the header grows over its neighbour.
    act(() => sort.focus())
    const header = sort.closest('[role=columnheader]') as HTMLElement
    expect(header.style.zIndex).toBe('3')
    expect(parseFloat(header.style.width)).toBeGreaterThan(40)
    expect(sort.style.position).not.toBe('absolute')
  })
  it('touch headers keep 44px targets (stream F: an 80px column opens its actions from a disclosure)', () => {
    const ui = headers(1, true)
    const disclosure = ui.getByRole('button', { name: 'Actions for column A' })
    expect(parseFloat(disclosure.style.minHeight)).toBeGreaterThanOrEqual(GRID_HEADER_SIZE)
    fireEvent.click(disclosure)
    expect(parseFloat(ui.getByRole('button', { name: 'Cycle sort column A' }).style.minHeight)).toBeGreaterThanOrEqual(GRID_HEADER_SIZE)
  })
  it('the row gutter widens to fit seven-digit row numbers', () => {
    expect(rowGutterWidth(1)).toBe(GRID_HEADER_SIZE)
    expect(rowGutterWidth(9999)).toBe(GRID_HEADER_SIZE)
    const wide = rowGutterWidth(1048576)
    expect(wide).toBeGreaterThan(GRID_HEADER_SIZE)
    const props = gridFixture()
    const headerProps = { projection: props.projection, window: props.window, viewport: props.viewport, columns: [0], selection: props.selection, sort: null, gutter: wide }
    const ui = render(<div role="grid"><div role="row"><GridRowHeader props={headerProps} slot={0} /></div></div>)
    const button = ui.getByRole('button', { name: 'Select row 1' })
    expect(button.style.width).toBe(`${wide}px`)
    expect(parseFloat((button.firstElementChild as HTMLElement).style.fontSize)).toBeLessThanOrEqual(12)
  })
  it('a row cut by the pane edge keeps its label where the whole row puts it (no re-centring while scrolling in)', () => {
    const props = gridFixture() // 24px rows, 96px viewport
    const window = createScrollWindow(props.projection, props.viewport, { offset: { x: 0, y: 12 } })
    const headerProps = { projection: props.projection, window, viewport: props.viewport, columns: [0], selection: props.selection, sort: null }
    const ui = render(<div role="grid"><div role="row"><GridRowHeader props={headerProps} slot={0} /></div></div>)
    const button = ui.getByRole('button', { name: 'Select row 1' })
    // Row 1 spans y -12..12: its clamped box starts at 0, its centre stays at 0 (-12 + 24/2).
    expect(parseFloat(button.style.top)).toBe(0)
    expect((button.firstElementChild as HTMLElement).style.height).toBe('24px')
  })
})

describe('#18 status line', () => {
  it('is a visible fixed-height bar inside the grid wrapper', () => {
    const ui = render(<GridSemantics {...gridFixture()} />)
    const status = ui.container.querySelector('[data-officeview-ui=viewer-status]') as HTMLElement
    expect(status.style.height).toBe(`${STATUS_BAR_HEIGHT}px`)
    expect(status.style.position).toBe('absolute')
    expect(status.style.bottom).toBe('0px')
  })
  it('is one text run, so a changing count never moves the notices after it (no layout shift)', () => {
    const ui = render(<GridSemantics {...gridFixture()} announcement="Selection cleared." />)
    const status = ui.container.querySelector('[data-officeview-ui=viewer-status]') as HTMLElement
    expect(status.childNodes).toHaveLength(1)
    expect(status.childNodes[0].nodeType).toBe(Node.TEXT_NODE)
    expect(status.title).toBe(status.textContent)
  })
  it('in the viewer pane it fills its positioned parent, so the status line follows the pane at once', () => {
    const ui = render(<GridSemantics {...gridFixture()} fill />)
    const wrapper = ui.container.firstElementChild as HTMLElement
    expect(wrapper.style.position).toBe('absolute')
    expect(wrapper.style.inset).toBe('0px')
  })
  it('a cached document with no affected formulas states the saved-values policy without "0 formulas ... (file too large)"', async () => {
    const { createObjectWorkbookSource } = await import('../src/xlsx/view/source')
    const { valueStatus } = await import('../src/xlsx/view/values')
    const doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'x' }] }]); doc.valuePolicy = 'cached'
    const base = createObjectWorkbookSource(doc, 'zero')
    const source = { ...base, valueStatus: () => valueStatus('cached', 0), openDiagnostics: [] }
    const ui = render(<InteractiveSheetDoc doc={doc} source={source} />)
    await waitFor(() => expect(ui.getByRole('status').textContent).toContain('Showing saved values; formulas are not recalculated'), FIRST_PAINT_WAIT)
    expect(ui.getByRole('status').textContent).not.toMatch(/0 formulas|file too large/)
  }, FIRST_PAINT_TEST.timeout)
})

describe('layout stability: horizontal scroll', () => {
  it('moves rows by transform only: gridcell lefts and the row header position stay fixed', () => {
    const props = gridFixture()
    const at = (x: number) => createScrollWindow(props.projection, props.viewport, { offset: { x, y: 0 } })
    const ui = render(<GridSemantics {...props} window={at(0)} />)
    const first = ui.container.querySelectorAll<HTMLElement>('[role=row] [role=gridcell]')[1] // not the edge cell
    const col = first.getAttribute('aria-colindex')
    const before = { left: first.style.left, row: (first.parentElement as HTMLElement).style.transform }
    ui.rerender(<GridSemantics {...props} window={at(37)} />)
    const after = document.getElementById(first.id) as HTMLElement
    expect(after.getAttribute('aria-colindex')).toBe(col)
    expect(after.style.left).toBe(before.left) // no layout move
    const row = after.parentElement as HTMLElement
    expect(row.style.transform).not.toBe(before.row) // the row moved by transform
    const x = Number(/translate\(([-\d.e]+)px/.exec(row.style.transform)![1])
    expect((row.querySelector('[role=rowheader]') as HTMLElement).style.transform).toBe(`translateX(${-x}px)`)
  })
})

describe('#20 find overlay', () => {
  it('paints the active match as an inset fill that never touches the cell edges (no underline-like stroke)', () => {
    const calls: Array<[string, number[]]> = []
    const ctx = new Proxy({}, { get: (_t, key) => key === 'save' || key === 'restore' ? () => {} : typeof key === 'string' && /Rect$/.test(key) ? (...args: number[]) => calls.push([key, args]) : undefined, set: () => true }) as unknown as CanvasRenderingContext2D
    const rect = { x: 10, y: 20, width: 60, height: 20 }
    paintFindOverlay(ctx, [rect], [rect])
    expect(calls.filter(([name]) => name === 'strokeRect')).toHaveLength(0)
    const fills = calls.filter(([name]) => name === 'fillRect').map(([, args]) => args)
    expect(fills).toContainEqual([10, 20, 60, 20]) // every hit: the exact band
    const active = fills.find(args => args[0] !== 10)!
    expect(active[0]).toBeGreaterThan(10); expect(active[1]).toBeGreaterThan(20)
    expect(active[0] + active[2]).toBeLessThan(70); expect(active[1] + active[3]).toBeLessThan(40)
  })
})

describe('#19 chrome font', () => {
  it('the viewer chrome uses an explicit system font stack and its native controls inherit it', async () => {
    const ui = render(<InteractiveSheetDoc doc={workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'x' }] }])} />)
    const wrapper = ui.container.querySelector('[data-officeview-sheet-instance]') as HTMLElement
    expect(wrapper.style.fontFamily).toMatch(/system-ui/)
    const rule = [...document.head.querySelectorAll('style[data-officeview-sheet-font]')].map(style => style.textContent).join('')
    expect(rule).toMatch(/button.*input.*select|font:\s*inherit/)
  })
})

describe('#13 popup and toolbar layout', () => {
  it('the column popup opens below the header row and is bounded by the cell area', () => {
    const props = gridFixture({ onFilter: () => {} })
    const ui = render(<GridSemantics {...props} popup={{ invoker: null, close: () => {}, content: <p>filter</p> }} />)
    const popup = ui.container.querySelector('[data-officeview-ui=grid-popup]') as HTMLElement
    expect(parseFloat(popup.style.top)).toBeGreaterThanOrEqual(GRID_HEADER_SIZE)
    expect(parseFloat(popup.style.maxHeight)).toBeLessThanOrEqual(props.viewport.height)
  })
  it('the search row is laid out in the toolbar, not floated over the worksheet navigator', async () => {
    const ui = render(<InteractiveSheetDoc doc={workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'x' }] }])} />)
    const search = ui.getByTestId('officeview-search')
    expect(search.closest('[data-officeview-ui=sheet-toolbar]')).not.toBeNull()
  })
})

describe('#12 spreadsheet print', () => {
  const unit = (width: number, height: number, painted: string[], name: string): Paintable => ({ spec: { widthPx: width, heightPx: height } as Paintable['spec'], paint: () => { painted.push(name) } })
  const loaderOf = (pages: Paintable[]) => async () => Object.assign(pages, { dispose: () => {} }) as unknown as PaintableArray
  it('renders each sheet at 2x device scale with a sheet-name heading and a page break before every later sheet', async () => {
    const print = vi.spyOn(window, 'print').mockImplementation(() => {}), painted: string[] = []
    const doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'x' }] }, { sheetId: 'B', cells: [{ row: 0, col: 0, value: 'y' }] }])
    const ui = render(<InteractiveSheetDoc doc={doc} printLoader={loaderOf([unit(100, 50, painted, 'a'), unit(120, 60, painted, 'b')])} />)
    fireEvent.click(ui.getByRole('button', { name: 'Print' }))
    await waitFor(() => expect(print).toHaveBeenCalled(), FIRST_PAINT_WAIT)
    const pages = ui.getByTestId('officeview-sheet-print-pages')
    const canvases = [...pages.querySelectorAll('canvas')]
    expect(canvases.map(c => [c.width, c.height])).toEqual([[200, 100], [240, 120]])
    expect(canvases.map(c => c.style.width)).toEqual(['100px', '120px'])
    const titles = [...pages.querySelectorAll('[data-officeview-print-sheet-title]')].map(h => h.textContent)
    expect(titles).toEqual(doc.sheets.map(sheet => sheet.name))
    const units = [...pages.children] as HTMLElement[]
    expect(units[0].style.breakBefore).not.toBe('page')
    expect(units[1].style.breakBefore).toBe('page')
    expect(painted).toEqual(['a', 'b'])
  }, FIRST_PAINT_TEST.timeout)
  it('falls back to 1x for a sheet whose 2x store does not fit the print budget', async () => {
    const print = vi.spyOn(window, 'print').mockImplementation(() => {})
    const ui = render(<InteractiveSheetDoc doc={workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'x' }] }])} printLoader={loaderOf([unit(3000, 3000, [], 'a')])} />)
    fireEvent.click(ui.getByRole('button', { name: 'Print' }))
    await waitFor(() => expect(print).toHaveBeenCalled(), FIRST_PAINT_WAIT)
    const canvas = ui.getByTestId('officeview-sheet-print-pages').querySelector('canvas')!
    expect([canvas.width, canvas.height]).toEqual([3000, 3000])
  }, FIRST_PAINT_TEST.timeout)
  it('print rules turn the viewer background white and let page breaks apply (block, not flex)', () => {
    render(<InteractiveSheetDoc doc={workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'x' }] }])} />)
    const css = [...document.head.querySelectorAll('style[data-officeview-sheet-print-media]')].map(style => style.textContent).join('')
    expect(css).toMatch(/@media print\{.*\[data-officeview-root\]\{[^}]*background:#fff!important/)
    expect(css).toMatch(/display:block!important/)
  })
})
