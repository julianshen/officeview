import { describe, expect, it, vi } from 'vitest'
import { useState } from 'react'
import { createRoot } from 'react-dom/client'
import { act, fireEvent, render, waitFor } from '@testing-library/react'
import { GridSemantics, type GridPopup, type GridSemanticsProps, type GridTransition } from '../src/components/xlsx/GridSemantics'
import { GridHeaders, GridRowHeader } from '../src/components/xlsx/GridHeaders'
import { SourceSelection, emptyGridSelection } from '../src/xlsx/view/selection'
import { gridFixture } from './helpers/xlsx-grid-fixture'

function StatefulGrid({ parentKey = () => {}, initial = {} }: { parentKey?: (key: string) => void; initial?: Partial<GridSemanticsProps> }) {
  const [props, setProps] = useState(() => gridFixture(initial)), [popup, setPopup] = useState<GridPopup>()
  return <div onKeyDown={event => parentKey(event.key)}><GridSemantics {...props} onChange={(next: GridTransition) => setProps({ ...props, ...next })}
    popup={popup} onFilter={(_col, invoker) => setPopup({ invoker, close: () => setPopup(undefined), content: <form onSubmit={event => event.preventDefault()}><input autoFocus aria-label="Numeric filter" defaultValue="-10" /><label>Show blanks<input type="checkbox" /></label><button type="button" onClick={() => setPopup(undefined)}>Close filter</button></form> })} /></div>
}

describe('grid keyboard and control ownership', () => {
  it('moves viewport and selection together; Shift uses source-safe display membership', () => {
    const parent = vi.fn(), ui = render(<StatefulGrid parentKey={parent} />), grid = ui.getByRole('grid')
    grid.focus()
    fireEvent.keyDown(grid, { key: 'ArrowDown', shiftKey: true })
    expect(ui.getByRole('gridcell', { name: 'A4, value 10; original A3' }).getAttribute('aria-selected')).toBe('true')
    expect(parent).not.toHaveBeenCalled()
    fireEvent.keyDown(grid, { key: 'End', ctrlKey: true })
    const last = ui.getByRole('gridcell', { name: 'XFD1048576, value 10' })
    expect(grid.getAttribute('aria-activedescendant')).toBe(last.id)
    expect(last.closest('[data-transient-active]')).toBeNull()
    expect(document.activeElement).toBe(grid)
    expect(ui.container.querySelectorAll('*').length).toBeLessThan(160)
    expect(ui.getByRole('status').textContent).toContain('XFD1048576')
  })
  it('handles arrows only while the grid wrapper owns actual focus and preserves workbook find', () => {
    const parent = vi.fn(), change = vi.fn(), props = gridFixture(), ui = render(<div onKeyDown={e => parent(e.key)}><GridSemantics {...props} onChange={change} /></div>)
    const grid = ui.getByRole('grid')
    fireEvent.keyDown(grid, { key: 'ArrowDown' })
    expect(change).not.toHaveBeenCalled()
    expect(parent).toHaveBeenCalledWith('ArrowDown')
    grid.focus()
    fireEvent.keyDown(grid, { key: 'f', ctrlKey: true })
    fireEvent.keyDown(grid, { key: 'f', metaKey: true })
    expect(parent.mock.calls.filter(([key]) => key === 'f')).toHaveLength(2)
  })
  it('enters native header controls with F2 and returns to a valid active grid cell on Escape', () => {
    const ui = render(<StatefulGrid />), grid = ui.getByRole('grid')
    grid.focus(); fireEvent.keyDown(grid, { key: 'F2' })
    const header = ui.getByRole('button', { name: 'Select Column A' })
    expect(document.activeElement).toBe(header)
    expect(grid.tabIndex).toBe(-1)
    expect(header.tabIndex).toBe(0)
    fireEvent.keyDown(header, { key: 'ArrowRight' })
    expect(document.activeElement).toBe(header)
    fireEvent.keyDown(header, { key: 'Escape' })
    expect(document.activeElement).toBe(grid)
    expect(grid.tabIndex).toBe(0)
    expect(document.getElementById(grid.getAttribute('aria-activedescendant')!)?.closest('[role="grid"]')).toBe(grid)
    expect(grid.querySelectorAll('[tabindex="0"]')).toHaveLength(0)
  })
  it('applies Escape in popup → selection → parent order and retains native filter inputs', () => {
    const parent = vi.fn(), ui = render(<StatefulGrid parentKey={parent} />), grid = ui.getByRole('grid')
    grid.focus(); fireEvent.keyDown(grid, { key: 'ArrowDown', altKey: true })
    const input = ui.getByRole('textbox', { name: 'Numeric filter' })
    expect(document.activeElement).toBe(input)
    for (const key of ['0', '-', 'ArrowDown', 'Enter', ' ']) fireEvent.keyDown(input, { key })
    expect(parent).not.toHaveBeenCalled()
    fireEvent.keyDown(input, { key: 'Escape' })
    expect(ui.queryByRole('dialog')).toBeNull()
    expect(grid.getAttribute('aria-activedescendant')).toBeTruthy()
    expect(document.activeElement).toBe(grid)
    fireEvent.keyDown(grid, { key: 'Escape' })
    expect(grid.hasAttribute('aria-activedescendant')).toBe(false)
    expect(parent).not.toHaveBeenCalled()
    fireEvent.keyDown(grid, { key: 'Escape' })
    expect(parent).toHaveBeenCalledWith('Escape')
  })
  it('returns popup focus to an owned invoker, or the grid when the invoker is detached', () => {
    const props = gridFixture(), close = vi.fn(), ui = render(<GridSemantics {...props} />)
    const invoker = ui.getByRole('button', { name: 'Select Column C' })
    for (const detached of [false, true]) {
      const popup = { close, invoker: detached ? document.createElement('button') : invoker, content: <input autoFocus aria-label="Rule" /> }
      ui.rerender(<GridSemantics {...props} popup={popup} />)
      fireEvent.keyDown(ui.getByRole('textbox', { name: 'Rule' }), { key: 'Escape' })
      ui.rerender(<GridSemantics {...props} />)
      expect(document.activeElement).toBe(detached ? ui.getByRole('grid') : invoker)
      expect(document.getElementById(ui.getByRole('grid').getAttribute('aria-activedescendant')!)?.closest('[role="grid"]')).toBe(ui.getByRole('grid'))
    }
    expect(close).toHaveBeenCalledTimes(2)
  })
  it('restores focus when a popup closes through its native button as well as Escape', () => {
    const ui = render(<StatefulGrid />), grid = ui.getByRole('grid')
    act(() => grid.focus()); fireEvent.keyDown(grid, { key: 'ArrowDown', altKey: true })
    const close = ui.getByRole('button', { name: 'Close filter' })
    act(() => close.focus()); fireEvent.click(close)
    expect(ui.queryByRole('dialog')).toBeNull()
    expect(document.activeElement).toBe(grid)
    expect(document.getElementById(grid.getAttribute('aria-activedescendant')!)?.closest('[role="grid"]')).toBe(grid)
  })
  it('does not restore a popup invoker captured from an older sheet generation', () => {
    function SheetChangingPopup() {
      const [props, setProps] = useState(gridFixture), [popup, setPopup] = useState<GridPopup>()
      return <><button onClick={() => setProps({ ...props, owner: { ...props.owner, sheetId: 'new-sheet', viewGen: 2 } })}>Replace sheet</button>
        <GridSemantics {...props} popup={popup} onFilter={(_col, grid) => setPopup({ invoker: grid.querySelector<HTMLButtonElement>('[data-header-column="0"]'), close: () => setPopup(undefined), content: <input autoFocus aria-label="Rule" /> })} /></>
    }
    const ui = render(<SheetChangingPopup />), grid = ui.getByRole('grid')
    act(() => grid.focus()); fireEvent.keyDown(grid, { key: 'ArrowDown', altKey: true })
    expect(ui.getByRole('dialog')).toBeTruthy()
    fireEvent.click(ui.getByRole('button', { name: 'Replace sheet' }))
    expect(ui.queryByRole('dialog')).toBeNull()
    expect(document.activeElement).toBe(grid)
    expect(grid.tabIndex).toBe(0)
  })
  it('header click/Ctrl/Shift reuse S6.1 membership instead of allocating a full axis', () => {
    const props = gridFixture(), changes: GridTransition[] = [], ui = render(<GridSemantics {...props} onChange={change => changes.push(change)} />)
    fireEvent.click(ui.getByRole('button', { name: 'Select Column C' }))
    expect(changes).toHaveLength(1)
    expect(new SourceSelection(changes[0].selection.ranges).contains({ row: 1048575, col: 2 })).toBe(true)
    expect(changes[0].selection.ranges).toHaveLength(1)
    ui.rerender(<GridSemantics {...props} selection={changes[0].selection} onChange={change => changes.push(change)} />)
    fireEvent.click(ui.getByRole('button', { name: 'Select Column A' }), { shiftKey: true })
    const selected = new SourceSelection(changes[1].selection.ranges)
    expect(selected.contains({ row: 1048575, col: 0 })).toBe(true)
    expect(selected.contains({ row: 1048575, col: 1 })).toBe(false)
    ui.rerender(<GridSemantics {...props} selection={changes[1].selection} onChange={change => changes.push(change)} />)
    fireEvent.click(ui.getByRole('button', { name: 'Select Column A' }), { ctrlKey: true })
    expect(new SourceSelection(changes[2].selection.ranges).contains({ row: 1048575, col: 0 })).toBe(false)
    expect(new SourceSelection(changes[2].selection.ranges).contains({ row: 1048575, col: 2 })).toBe(true)
  })
})

/** Mounts a real grid in its own React root inside OfficeDoc's region, the way
 * S5 will nest it: native events reach OfficeDoc only if the grid lets them bubble. */
async function officeDocWithGrid(options: { allowCopy?: boolean; allowPrint?: boolean } = {}) {
  const { OfficeDoc } = await import('../src/components/OfficeDoc')
  const { parseDocx } = await import('../src/docx/parse')
  const { OfficePackage } = await import('../src/core/zip')
  const { buildDocx } = await import('../src/testdata/ooxml-builders')
  const doc = await parseDocx(await OfficePackage.load(await buildDocx([{ runs: [{ text: 'alpha beta gamma' }] }])))
  const ui = render(<OfficeDoc document={doc} {...options} />)
  await waitFor(() => expect(ui.container.querySelector('canvas')).toBeTruthy())
  const canvas = ui.container.querySelector('canvas') as HTMLCanvasElement
  canvas.getBoundingClientRect = () => ({ left: 0, top: 0, width: 816, height: 1056, right: 816, bottom: 1056, x: 0, y: 0, toJSON: () => ({}) }) as DOMRect
  const host = document.createElement('div')
  ui.getByRole('region').append(host)
  const root = createRoot(host)
  act(() => root.render(<StatefulGrid initial={{ selection: emptyGridSelection() }} />))
  const grid = host.querySelector<HTMLElement>('[role="grid"]')!
  return { ui, canvas, grid, cleanup: () => act(() => root.unmount()) }
}

describe('OfficeDoc as the grid ancestor', () => {
  it('still swallows Ctrl/Cmd+P on the grid when printing is off', async () => {
    const { grid, cleanup } = await officeDocWithGrid({ allowPrint: false })
    act(() => grid.focus())
    for (const modifier of ['metaKey', 'ctrlKey']) {
      const event = new KeyboardEvent('keydown', { key: 'p', [modifier]: true, bubbles: true, cancelable: true })
      act(() => { grid.dispatchEvent(event) })
      expect(event.defaultPrevented).toBe(true)
    }
    cleanup()
  })
  it('still swallows Ctrl/Cmd+C on the grid when copying is off', async () => {
    const { grid, cleanup } = await officeDocWithGrid({ allowCopy: false })
    act(() => grid.focus())
    for (const modifier of ['metaKey', 'ctrlKey']) {
      const event = new KeyboardEvent('keydown', { key: 'c', [modifier]: true, bubbles: true, cancelable: true })
      act(() => { grid.dispatchEvent(event) })
      expect(event.defaultPrevented).toBe(true)
    }
    cleanup()
  })
  it('falls an unhandled grid Escape through to the viewer selection clear', async () => {
    const { ui, canvas, grid, cleanup } = await officeDocWithGrid()
    const copyButton = () => ui.container.querySelector('[data-testid="officeview-copy-selection"]')
    fireEvent.doubleClick(canvas, { clientX: 160, clientY: 104, bubbles: true })
    await waitFor(() => expect(copyButton()).toBeTruthy())
    act(() => grid.focus())
    fireEvent.keyDown(grid, { key: 'Escape' })
    await waitFor(() => expect(copyButton()).toBeNull())
    cleanup()
  })
})

describe('viewer gesture isolation', () => {
  it('OfficeDoc ignores editable/UI key, pointer, double-click and wheel events', async () => {
    const { OfficeDoc } = await import('../src/components/OfficeDoc')
    const { parseDocx } = await import('../src/docx/parse')
    const { OfficePackage } = await import('../src/core/zip')
    const { buildDocx } = await import('../src/testdata/ooxml-builders')
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([{ runs: [{ text: 'baseline' }] }])))
    const ui = render(<OfficeDoc document={doc} />)
    const { waitFor } = await import('@testing-library/react')
    await waitFor(() => expect(ui.container.querySelector('canvas')).toBeTruthy())
    fireEvent.click(ui.getByRole('button', { name: 'Zoom in' }))
    const content = ui.container.querySelector('[data-testid="officeview-content"]') as HTMLElement
    await waitFor(() => expect(content.style.transform).toContain('scale(1.5)'))
    const before = content.style.transform
    const region = ui.getByRole('region'), targets = [document.createElement('input'), document.createElement('textarea'), document.createElement('select'), document.createElement('div'), document.createElement('div')]
    targets[3].setAttribute('contenteditable', 'true'); targets[4].setAttribute('data-officeview-ui', 'headers')
    for (const target of targets) {
      region.append(target)
      for (const key of ['0', '-', '+']) fireEvent.keyDown(target, { key })
      fireEvent.doubleClick(target, { clientX: 100, clientY: 100 })
      fireEvent.pointerDown(target, { pointerId: 1, clientX: 10, clientY: 10 })
      fireEvent.pointerMove(target, { pointerId: 1, clientX: 200, clientY: 200 })
      fireEvent.pointerUp(target, { pointerId: 1 })
      fireEvent.wheel(target, { ctrlKey: true, deltaY: -100 })
    }
    await new Promise(resolve => setTimeout(resolve, 30))
    expect(content.style.transform).toBe(before)
    fireEvent.keyDown(targets[0], { key: 'f', metaKey: true })
    expect(document.activeElement).toBe(ui.getByRole('searchbox', { name: 'Search document' }))
  })
  it('isolates headers and preserves native checkbox Space/form Enter defaults', () => {
    const pointer = vi.fn(), dbl = vi.fn(), key = vi.fn(), ui = render(<div onPointerDown={pointer} onDoubleClick={dbl} onKeyDown={key}><StatefulGrid /></div>)
    const header = ui.getByRole('button', { name: 'Select Column A' })
    fireEvent.pointerDown(header); fireEvent.doubleClick(header)
    expect(pointer).not.toHaveBeenCalled()
    expect(dbl).not.toHaveBeenCalled()
    act(() => header.focus())
    const native = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    act(() => { header.dispatchEvent(native) })
    expect(native.defaultPrevented).toBe(false)
    expect(key).not.toHaveBeenCalled()
    const grid = ui.getByRole('grid')
    act(() => grid.focus()); fireEvent.keyDown(grid, { key: 'ArrowDown', altKey: true })
    const checkbox = ui.getByRole('checkbox', { name: 'Show blanks' }) as HTMLInputElement
    act(() => checkbox.focus())
    const space = new KeyboardEvent('keydown', { key: ' ', bubbles: true, cancelable: true })
    act(() => { checkbox.dispatchEvent(space) })
    expect(space.defaultPrevented).toBe(false)
    fireEvent.click(checkbox)
    expect(checkbox.checked).toBe(true)
    const input = ui.getByRole('textbox', { name: 'Numeric filter' })
    act(() => input.focus())
    const enter = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })
    act(() => { input.dispatchEvent(enter) })
    expect(enter.defaultPrevented).toBe(false)
    expect(key).not.toHaveBeenCalled()

  })
  it('header strips isolate pointer/double-click even when mounted through another S5 owner', () => {
    const props = gridFixture(), pointer = vi.fn(), dbl = vi.fn()
    const headerProps = { projection: props.projection, window: props.window, viewport: props.viewport, columns: [0], selection: props.selection, sort: null }
    const ui = render(<div onPointerDown={pointer} onDoubleClick={dbl}><div role="grid"><GridHeaders {...headerProps} /><div role="row"><GridRowHeader props={headerProps} slot={0} /></div></div></div>)
    for (const button of ui.getAllByRole('button')) { fireEvent.pointerDown(button); fireEvent.doubleClick(button) }
    expect(pointer).not.toHaveBeenCalled()
    expect(dbl).not.toHaveBeenCalled()
  })
  it('routes overlapping 44px row hit-slop by physical row, preserving 24px geometry', () => {
    const props = gridFixture(), change = vi.fn(), ui = render(<GridSemantics {...props} onChange={change} />)
    const grid = ui.getByRole('grid')
    grid.getBoundingClientRect = () => ({ left: 0, top: 0, width: 284, height: 140, right: 284, bottom: 140, x: 0, y: 0, toJSON: () => ({}) })
    const row1 = ui.getByRole('button', { name: 'Select row 1' })
    expect(row1.style.minHeight).toBe('44px')
    expect(row1.style.minWidth).toBe('44px')
    expect(row1.closest('[role="row"]')?.getAttribute('style')).toContain('height: 24px')
    // Its expanded button overlaps the second row: the visual coordinate owns the hit.
    fireEvent.click(row1, { detail: 1, clientX: 20, clientY: 44 + 25 })
    expect(change).toHaveBeenCalledOnce()
    expect(change.mock.calls[0][0].selection.active).toEqual({ row: 3, col: 0 })
    expect(new SourceSelection(change.mock.calls[0][0].selection.ranges).contains({ row: 0, col: 0 })).toBe(false)
  })
})
