import { describe, expect, it } from 'vitest'
import { render, within } from '@testing-library/react'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { MAX_ROWS, MAX_COLS } from '../src/xlsx/view/coordinates'
import { createScrollWindow, navigateTo } from '../src/xlsx/view/scrollWindow'
import { emptyGridSelection, selectCell, selectHeader } from '../src/xlsx/view/selection'
import { gridFixture } from './helpers/xlsx-grid-fixture'

describe('owned virtual grid semantics (S6.2)', () => {
  it('owns header/body roles with chrome-inclusive counts, gaps and worksheet names', () => {
    const ui = render(<GridSemantics {...gridFixture()} />), grid = ui.getByRole('grid')
    expect(grid.getAttribute('aria-rowcount')).toBe('1048577')
    expect(grid.getAttribute('aria-colcount')).toBe('16385')
    const rows = within(grid).getAllByRole('row')
    expect(rows.map(row => row.getAttribute('aria-rowindex'))).toEqual(['1', '2', '3', '5', '6'])
    expect(new Set(rows.map(row => row.getAttribute('aria-rowindex'))).size).toBe(rows.length)
    const headers = within(rows[0]).getAllByRole('columnheader')
    expect(headers.map(cell => cell.getAttribute('aria-colindex'))).toEqual(['1', '2', '4', '5'])
    const active = ui.getByRole('gridcell', { name: 'A2, value 10; original A4' })
    expect(active.closest('[role="row"]')?.getAttribute('aria-rowindex')).toBe('3')
    expect(active.getAttribute('aria-colindex')).toBe('2')
    expect(active.getAttribute('aria-selected')).toBe('true')
    expect(grid.getAttribute('aria-activedescendant')).toBe(active.id)
    expect(rows.every(row => row.querySelector('[role="rowheader"], [role="columnheader"]'))).toBe(true)
    expect(ui.getAllByRole('status')).toHaveLength(1)
    expect(ui.getByRole('status').textContent).toContain('2 of 3 records')
    expect(grid.querySelectorAll('[tabindex="0"]')).toHaveLength(0)
    expect(grid.tabIndex).toBe(0)
  })
  it('represents XFD1048576 at the exact last semantic indices with bounded DOM', () => {
    const props = gridFixture(), target = { row: MAX_ROWS - 1, col: MAX_COLS - 1 }
    const selection = selectCell(props.selection, props.projection, target)
    const result = navigateTo(props.window, { metrics: props.projection, viewport: props.viewport, view: { ...props.view, selection } }, { addressSpace: 'source', at: target })
    const ui = render(<GridSemantics {...props} selection={selection} window={result.window} />)
    const cell = ui.getByRole('gridcell', { name: 'XFD1048576, value 10' })
    expect(cell.getAttribute('aria-colindex')).toBe('16385')
    expect(cell.closest('[role="row"]')?.getAttribute('aria-rowindex')).toBe('1048577')
    expect(ui.container.querySelectorAll('*').length).toBeLessThan(150)
    expect(ui.getByRole('grid').getAttribute('aria-activedescendant')).toBe(cell.id)
  })
  it('uses unique instance/generation IDs and drops a hidden or null active descendant', () => {
    const props = gridFixture(), ui = render(<><GridSemantics {...props} /><GridSemantics {...props} /></>)
    const ids = [...ui.container.querySelectorAll('[id]')].map(node => node.id)
    expect(ids.length).toBeGreaterThan(0)
    expect(new Set(ids).size).toBe(ids.length)
    const before = ui.getAllByRole('grid')[0].getAttribute('aria-activedescendant')
    ui.rerender(<GridSemantics {...props} owner={{ ...props.owner, valueGen: 2, viewGen: 2 }} />)
    const grid = ui.getByRole('grid'), next = grid.getAttribute('aria-activedescendant')!
    expect(next).not.toBe(before)
    expect(document.getElementById(next)?.closest('[role="grid"]')).toBe(grid)
    ui.rerender(<GridSemantics {...props} selection={{ ...props.selection, active: null, anchor: null }} />)
    expect(ui.getByRole('grid').hasAttribute('aria-activedescendant')).toBe(false)
  })
  it('keeps the active descendant ID across value/font republish and keeps it short for long doc IDs', () => {
    const docId = `https://example.test/${'very long path/'.repeat(40)}book.xlsx?sig=abc%20def`
    const props = gridFixture({ owner: { docId, sheetId: 'Sheet 1', valueGen: 1, fontGen: 1, viewGen: 1 } })
    const ui = render(<GridSemantics {...props} />), grid = ui.getByRole('grid')
    const before = grid.getAttribute('aria-activedescendant')!
    expect(before.length).toBeLessThan(64)
    expect(before).not.toMatch(/[%\s/]/)
    ui.rerender(<GridSemantics {...props} owner={{ ...props.owner, valueGen: 7, fontGen: 3 }} />)
    expect(grid.getAttribute('aria-activedescendant')).toBe(before)
    expect(document.getElementById(before)?.closest('[role="grid"]')).toBe(grid)
  })
  it('removes a stale descendant for an all-hidden filter body and preserves exact membership', () => {
    const props = gridFixture(), visibility = visibilityFromHidden(createIntervalSet([{ r0: 1, r1: 3 }]))
    const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: props.projection.cols, order: props.projection.order, visibility })
    const filter = { range: { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1 as const, totalsRows: 0 as const }, columns: new Map([[0, { kind: 'number' as const, op: 'gt' as const, first: 100 }]]) }
    const ui = render(<GridSemantics {...props} projection={projection} view={{ ...props.view, visibility, filter }} />)
    expect(ui.getByRole('grid').hasAttribute('aria-activedescendant')).toBe(false)
    expect(ui.getByRole('status').textContent).toContain('No visible records.')
    expect(ui.getByRole('status').textContent).toContain('0 of 3 records')
    expect(props.selection.ranges).toHaveLength(1)
  })
  it('keeps one transient owned active cell then removes its extra row after viewport transfer', () => {
    const props = gridFixture(), selection = selectCell(props.selection, props.projection, { row: 1000, col: 1000 })
    const ui = render(<GridSemantics {...props} selection={selection} />)
    const active = ui.getByRole('gridcell', { name: 'ALM1001, value 10' })
    expect(active.closest('[role="row"]')?.getAttribute('data-transient-active')).toBe('true')
    const result = navigateTo(props.window, { metrics: props.projection, viewport: props.viewport, view: { ...props.view, selection } }, { addressSpace: 'source', at: selection.active! })
    ui.rerender(<GridSemantics {...props} selection={selection} window={result.window} />)
    expect(ui.container.querySelector('[data-transient-active="true"]')).toBeNull()
    expect(ui.getAllByRole('gridcell', { name: 'ALM1001, value 10' })).toHaveLength(1)
  })
  it('bounds transient active CSS geometry to the pane, not worksheet extent', () => {
    const props = gridFixture(), selection = selectCell(props.selection, props.projection, { row: MAX_ROWS - 1, col: MAX_COLS - 1 })
    const ui = render(<GridSemantics {...props} selection={selection} />)
    const transient = ui.container.querySelector<HTMLElement>('[data-transient-active="true"]')!
    expect(Math.abs(parseFloat(transient.style.top) || 0)).toBeLessThanOrEqual(props.viewport.height + 44)
  })
  it('clips giant-cell DOM geometry without changing worksheet metrics', () => {
    const props = gridFixture()
    const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 1000000 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 1000000 }) })
    const active = selectCell(emptyGridSelection(), projection, { row: 0, col: 0 })
    const ui = render(<GridSemantics {...props} projection={projection} selection={active} window={createScrollWindow(projection, props.viewport)} />)
    for (const row of ui.container.querySelectorAll<HTMLElement>('[role="row"]')) expect(parseFloat(row.style.height)).toBeLessThanOrEqual(props.viewport.height)
    for (const col of ui.container.querySelectorAll<HTMLElement>('[role="columnheader"]')) expect(parseFloat(col.style.width)).toBeLessThanOrEqual(props.viewport.width)
    expect(projection.rows.size(0)).toBe(1000000)
    expect(projection.cols.size(0)).toBe(1000000)
  })
  it('exposes full-column selection and multi-key sort priorities on owned headers', () => {
    const props = gridFixture(), selection = selectHeader(props.selection, props.projection, 'column', 2)
    const sort = { range: { r0: 0, r1: 4, c0: 0, c1: 2, headerRows: 1 as const, totalsRows: 0 as const }, keys: [{ col: 2, order: 'desc' as const }, { col: 0, order: 'asc' as const }], caseSensitive: false }
    const ui = render(<GridSemantics {...props} selection={selection} view={{ ...props.view, sort }} />)
    const header = ui.getByRole('columnheader', { name: /Column C/ })
    expect(header.getAttribute('aria-selected')).toBe('true')
    expect(header.getAttribute('aria-sort')).toBe('descending')
    expect(header.getAttribute('aria-label')).toContain('sort priority 1, descending')
    expect(ui.getByRole('columnheader', { name: /Column A/ }).getAttribute('aria-label')).toContain('sort priority 2, ascending')
    expect(ui.getByRole('columnheader', { name: /Column A/ }).hasAttribute('aria-sort')).toBe(false)
  })
  it('exposes full-row selection on the owned rowheader of its display slot', () => {
    const props = gridFixture(), selection = selectHeader(props.selection, props.projection, 'row', 3)
    const ui = render(<GridSemantics {...props} selection={selection} />)
    // source row 3 sits in display slot 1, labelled Row 2
    expect(ui.getByRole('rowheader', { name: 'Row 2' }).getAttribute('aria-selected')).toBe('true')
    for (const name of ['Row 1', 'Row 4', 'Row 5']) expect(ui.getByRole('rowheader', { name }).getAttribute('aria-selected')).toBe('false')
  })
  it('does not double-count overlapping selection parts as a full header', () => {
    const props = gridFixture()
    // 600000 + 448576 rows sum to the axis length, but their union leaves a gap
    const selection = { active: { row: 0, col: 2 }, anchor: { row: 0, col: 2 }, ranges: [
      { rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 599999 }] }, cols: [{ c0: 2, c1: 2 }] },
      { rows: { kind: 'intervals' as const, intervals: [{ r0: 400000, r1: 848575 }] }, cols: [{ c0: 2, c1: 2 }] },
      // 10000 + 6384 columns on row 1 likewise sum to 16384 with gaps
      { rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 1 }, { c0: 3, c1: 10000 }] },
      { rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 5000, c1: 11383 }] },
    ] }
    const ui = render(<GridSemantics {...props} selection={selection} />)
    expect(ui.getByRole('columnheader', { name: 'Column C' }).getAttribute('aria-selected')).toBe('false')
    expect(ui.getByRole('rowheader', { name: 'Row 1' }).getAttribute('aria-selected')).toBe('false')
    const full = { ...selection, ranges: [...selection.ranges, { rows: { kind: 'intervals' as const, intervals: [{ r0: 500000, r1: MAX_ROWS - 1 }] }, cols: [{ c0: 2, c1: 2 }] }] }
    ui.rerender(<GridSemantics {...props} selection={full} />)
    expect(ui.getByRole('columnheader', { name: 'Column C' }).getAttribute('aria-selected')).toBe('true')
    const list = (rows: number[]) => ({ kind: 'list' as const, r0: rows[0], r1: rows[rows.length - 1], rows: Uint32Array.from(rows) })
    const mixed = (tail: number[]) => ({ ...selection, ranges: [{ rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: MAX_ROWS - 3 }] }, cols: [{ c0: 2, c1: 2 }] }, { rows: list(tail), cols: [{ c0: 2, c1: 2 }] }] })
    ui.rerender(<GridSemantics {...props} selection={mixed([MAX_ROWS - 2, MAX_ROWS - 1])} />)
    expect(ui.getByRole('columnheader', { name: 'Column C' }).getAttribute('aria-selected')).toBe('true')
    ui.rerender(<GridSemantics {...props} selection={mixed([MAX_ROWS - 3, MAX_ROWS - 1])} />)
    expect(ui.getByRole('columnheader', { name: 'Column C' }).getAttribute('aria-selected')).toBe('false')
  })
  it('renders worker row refs with unknown, not thrown, membership', () => {
    const props = gridFixture(), stamp = { docId: 'doc', sheetId: 's', valueGen: 1 }
    const selection = { ...props.selection, ranges: [{ rows: { kind: 'ref' as const, refId: 'r1', count: 5, stamp }, cols: [{ c0: 0, c1: 0 }] }] }
    const ui = render(<GridSemantics {...props} selection={selection} />)
    const active = ui.getByRole('gridcell', { name: 'A2, value 10; original A4' })
    expect(active.hasAttribute('aria-selected')).toBe(false)
    expect(ui.getByRole('columnheader', { name: 'Column A' }).hasAttribute('aria-selected')).toBe(false)
    expect(ui.getByRole('rowheader', { name: 'Row 2' }).hasAttribute('aria-selected')).toBe(false)
    expect(ui.getByRole('columnheader', { name: 'Column C' }).getAttribute('aria-selected')).toBe('false')
  })
  it('keeps an admitted 25% desktop viewport in detail and applies the mobile frame budget', () => {
    const props = gridFixture(), viewport = { width: 1440, height: 900 }
    const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }) })
    const window = createScrollWindow(projection, viewport, { zoom: 0.25, deviceScale: 2 })
    const selection = selectCell(emptyGridSelection(), projection, { row: 0, col: 0 })
    const ui = render(<GridSemantics {...props} profile="D" projection={projection} viewport={viewport} window={window} selection={selection} />)
    expect(ui.container.querySelectorAll('[role="gridcell"]').length).toBe(10800)
    expect(ui.container.querySelector('[role="grid"]')?.getAttribute('aria-description')).not.toContain('Coarse')
    ui.rerender(<GridSemantics {...props} profile="M" projection={projection} viewport={viewport} window={window} selection={selection} />)
    expect(ui.container.querySelectorAll('[role="gridcell"]').length).toBe(1)
    expect(ui.container.querySelector('[role="grid"]')?.getAttribute('aria-description')).toContain('Coarse')
  }, 30000)
  it('uses an explicit coarse description and only exact active semantics for subpixel density', () => {
    const props = gridFixture(), projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 0.00001 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 0.00001 }) })
    const selection = selectCell(emptyGridSelection(), projection, { row: 0, col: 0 })
    const ui = render(<GridSemantics {...props} projection={projection} selection={selection} window={createScrollWindow(projection, props.viewport)} />)
    expect(ui.container.querySelectorAll('*').length).toBeLessThan(30)
    expect(ui.getAllByRole('gridcell')).toHaveLength(1)
    expect(ui.getByRole('status').textContent).toContain('Zoom in or inspect a cell for detail')
  })
})
