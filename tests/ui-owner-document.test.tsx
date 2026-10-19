/**
 * Final review M6: a viewer rendered into another document (a popup window)
 * closes its menus on an outside pointerdown in THAT document. The listeners
 * hang off the owner document, not the global `document`.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, within } from '@testing-library/react'
import { MoreMenu } from '../src/components/paged/MoreMenu'
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
