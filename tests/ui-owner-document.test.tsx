/**
 * Final review M6: a viewer rendered into another document (a popup window)
 * closes its menus on an outside pointerdown in THAT document. The listeners
 * hang off the owner document, not the global `document`.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, within } from '@testing-library/react'
import { MoreMenu } from '../src/components/paged/MoreMenu'
import { PagedToolbar } from '../src/components/paged/PagedToolbar'
import { GridHeaders } from '../src/components/xlsx/GridHeaders'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { emptyGridSelection, selectCell } from '../src/xlsx/view/selection'

afterEach(() => cleanup())

const frames: HTMLIFrameElement[] = []
afterEach(() => { frames.splice(0).forEach((f) => f.remove()) })

/** A same-origin frame's document stands in for a popup window's. */
function otherDocument() {
  const frame = document.createElement('iframe')
  document.body.append(frame)
  frames.push(frame)
  const doc = frame.contentDocument!
  const container = doc.createElement('div')
  doc.body.append(container)
  return { doc, container }
}

describe('outside-click listeners use the owner document', () => {
  it('MoreMenu closes on a pointerdown elsewhere in its own document', async () => {
    const { doc, container } = otherDocument()
    const ui = render(<MoreMenu buttonClass="" items={[{ key: 'a', label: 'Alpha', icon: null, onSelect: vi.fn() }]} />, { container, baseElement: doc.body })
    const more = within(doc.body).getByRole('button', { name: 'More options' })
    fireEvent.click(more)
    expect(more.getAttribute('aria-expanded')).toBe('true')
    await act(async () => { fireEvent.pointerDown(doc.body) })
    expect(more.getAttribute('aria-expanded')).toBe('false')
    ui.unmount()
  })

  it('MoreMenu arrow keys read focus from its own document', () => {
    const { doc, container } = otherDocument()
    const items = ['Alpha', 'Beta', 'Gamma'].map((label) => ({ key: label, label, icon: null, onSelect: vi.fn() }))
    const ui = render(<MoreMenu buttonClass="" items={items} />, { container, baseElement: doc.body })
    fireEvent.click(within(doc.body).getByRole('button', { name: 'More options' }))
    const menu = within(doc.body).getByRole('menu')
    const [alpha, beta] = within(menu).getAllByRole('menuitem')
    expect(doc.activeElement).toBe(alpha)
    fireEvent.keyDown(menu, { key: 'ArrowDown' })
    expect(doc.activeElement).toBe(beta)
    fireEvent.keyDown(menu, { key: 'ArrowUp' })
    expect(doc.activeElement).toBe(alpha)
    ui.unmount()
  })

  const toolbarProps = {
    kind: 'document' as const, current: 1, total: 3, onPrev() {}, onNext() {}, zoom: 1, onZoomIn() {}, onZoomOut() {}, onResetZoom() {},
    onFit() {}, fitActive: false, copyState: 'idle' as const, showSearch: true, showZoom: true, onEscapeToDocument() {},
    search: <input type="search" aria-label="Search document" />, compact: true, focusSearchRequest: 1,
  }

  it('PagedToolbar Escape in the compact Find input collapses Find in its own document', () => {
    const { doc, container } = otherDocument()
    render(<PagedToolbar {...toolbarProps} />, { container, baseElement: doc.body })
    const input = within(doc.body).getByRole('searchbox')
    expect(doc.activeElement).toBe(input)
    fireEvent.keyDown(input, { key: 'Escape' })
    expect(within(doc.body).queryByRole('searchbox')).toBeNull()
    expect(doc.activeElement).toBe(doc.querySelector('[data-ov-control="find"]'))
  })

  it('PagedToolbar Ctrl+C in the Find input never reaches the allowCopy={false} guard in its own document', () => {
    const { doc, container } = otherDocument()
    // Stands in for OfficeDoc's region guard (allowCopy={false} prevents Ctrl/Cmd+C).
    const guard = vi.fn((e: React.KeyboardEvent) => { if ((e.ctrlKey || e.metaKey) && e.key === 'c') e.preventDefault() })
    render(<div onKeyDown={guard}><PagedToolbar {...toolbarProps} /></div>, { container, baseElement: doc.body })
    const input = within(doc.body).getByRole('searchbox')
    expect(fireEvent.keyDown(input, { key: 'c', ctrlKey: true })).toBe(true)
    expect(guard).not.toHaveBeenCalled()
  })

  it.each([[64, 'header-actions'], [30, 'header-lane']])('GridHeaders (%ipx columns) closes its %s group on a pointerdown elsewhere in its own document', (width, kind) => {
    const { doc, container } = otherDocument()
    const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: width }) })
    const viewport = { width: 600, height: 400 }
    render(<GridHeaders projection={projection} window={createScrollWindow(projection, viewport)} viewport={viewport} columns={[0, 1, 2]}
      selection={selectCell(emptyGridSelection(), projection, { row: 0, col: 0 })} sort={null} touchControls onSortCycle={vi.fn()} onFilter={vi.fn()} onSelect={vi.fn()} onSortDialog={vi.fn()} />, { container, baseElement: doc.body })
    fireEvent.click(within(doc.body).getAllByRole('button', { name: /^Actions for column/ })[0])
    expect(doc.querySelector(`[data-officeview-ui=${kind}]`)).not.toBeNull()
    fireEvent.pointerDown(doc.body)
    expect(doc.querySelector(`[data-officeview-ui=${kind}]`)).toBeNull()
  })
})
