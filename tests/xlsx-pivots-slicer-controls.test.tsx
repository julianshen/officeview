import { act, cleanup, fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { SlicerController } from '../src/xlsx/pivots/slicers/controller'
import { SlicerLayer } from '../src/components/xlsx/pivot/SlicerLayer'
import { PX_PER_PT } from '../src/xlsx/pivots/buttons'
import { slicerLayout } from '../src/xlsx/pivots/slicers/paint'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { readFileSync } from 'node:fs'
import type { SlicerModel } from '../src/xlsx/pivots/slicers/model'
import type { ScrollWindow } from '../src/xlsx/view/contracts'

function cellText(container: HTMLElement, row: number, col: number): string | undefined {
  return container.querySelector(`[role=row][aria-rowindex="${row + 2}"]`)?.querySelector(`[role=gridcell][aria-colindex="${col + 2}"]`)?.textContent ?? undefined
}

const model: SlicerModel = {
  name: 'Slicer_Product', cacheName: 'Slicer_Product', cacheId: 1, sourceFieldIndex: 2, revision: 0, allItemIds: [0, 1, 2], selectedItemIds: [0, 1, 2],
  caption: 'Product', showCaption: true, columnCount: 1, rowHeightEmu: 209550, style: 'SlicerStyleLight1', startItem: 0,
  items: [{ id: 0, label: 'Apples', selected: true, noData: false }, { id: 1, label: 'Bananas', selected: true, noData: false }, { id: 2, label: 'Cherries', selected: true, noData: true }],
  drives: [{ pivotTable: 'Pivot1' }],
}
const window: ScrollWindow = { logicalOrigin: { x: 0, y: 0 }, physicalPosition: { x: 0, y: 0 }, physicalExtent: { width: 500, height: 500 }, zoom: 1, deviceScale: 1 }
const box = { x: 0, y: 0, width: 140, height: 140 }
afterEach(cleanup)

describe('interactive slicer layer', () => {
  it('exposes selected/no-data state, keeps the grid as the only tab stop, and refuses unmeasured modifiers', () => {
    const announce = vi.fn(), apply = vi.fn(async () => ({ kind: 'committed' as const, revision: 1 }))
    const controller = new SlicerController([model], { apply, announce })
    render(<SlicerLayer model={model} box={box} window={window} viewport={{ width: 400, height: 300 }} controller={controller} />)
    const apples = screen.getByRole('button', { name: 'Apples, selected' })
    expect(apples.getAttribute('tabindex')).toBe('-1')
    expect(screen.getByRole('button', { name: 'Cherries, selected, no data' })).toBeTruthy()
    fireEvent.click(apples, { metaKey: true })
    expect(apply).not.toHaveBeenCalled()
    expect(announce).toHaveBeenCalled()
    expect(screen.getByRole('button', { name: 'Multi-Select unavailable' }).getAttribute('aria-pressed')).toBe('false')
  })

  it('maps Arrow navigation and Enter to the measured single-item action, and Escape restores the grid focus', async () => {
    const apply = vi.fn(async (_model: SlicerModel, _action: { kind: string; selectedItemIds?: readonly number[] }, _signal: AbortSignal) => ({ kind: 'committed' as const, revision: 1 }))
    const controller = new SlicerController([model], { apply, announce: vi.fn() })
    const invoker = document.createElement('div'); invoker.tabIndex = 0; document.body.append(invoker); invoker.focus()
    render(<SlicerLayer model={model} box={box} window={window} viewport={{ width: 400, height: 300 }} controller={controller} invoker={invoker} />)
    const apples = screen.getByRole('button', { name: 'Apples, selected' })
    apples.focus()
    fireEvent.keyDown(apples, { key: 'ArrowDown' })
    expect(document.activeElement).toBe(screen.getByRole('button', { name: 'Bananas, selected' }))
    await act(async () => { fireEvent.keyDown(document.activeElement!, { key: 'Enter' }); await Promise.resolve() })
    expect(apply.mock.calls[0]?.[0].name).toBe('Slicer_Product')
    expect(apply.mock.calls[0]?.[1]).toEqual({ kind: 'selection', selectedItemIds: [1] })
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' })
    expect(document.activeElement).toBe(invoker)
    invoker.remove()
  })

  it('keeps selection only after a committed reply and the measured Clear Filter restores All', async () => {
    const apply = vi.fn(async (_model: SlicerModel, action: { kind: string }) => ({ kind: 'committed' as const, revision: action.kind === 'clear' ? 2 : 1 }))
    const controller = new SlicerController([model], { apply, announce: vi.fn() })
    render(<SlicerLayer model={model} box={box} window={window} viewport={{ width: 400, height: 300 }} controller={controller} />)
    fireEvent.click(screen.getByRole('button', { name: 'Bananas, selected' }))
    await waitFor(() => expect(screen.getByRole('button', { name: 'Apples, not selected' }).getAttribute('aria-pressed')).toBe('false'))
    fireEvent.click(screen.getByRole('button', { name: 'Clear Filter' }))
    await waitFor(() => expect(screen.getByRole('button', { name: 'Apples, selected' }).getAttribute('aria-pressed')).toBe('true'))
  })

  it('covers the painted items at 200% zoom without double scaling', () => {
    const controller = new SlicerController([model], { apply: vi.fn(), announce: vi.fn() })
    const zoomed: ScrollWindow = { ...window, zoom: 2 }
    render(<SlicerLayer model={model} box={box} window={zoomed} viewport={{ width: 800, height: 600 }} controller={controller} />)
    // the layout is worksheet-logical; the viewport applies zoom exactly once
    const logical = slicerLayout(model, box, PX_PER_PT).items[0]!
    const apples = screen.getByRole('button', { name: 'Apples, selected' }) as HTMLElement
    expect(parseFloat(apples.style.left)).toBeCloseTo(2 * logical.rect.x, 5)
    expect(parseFloat(apples.style.width)).toBeCloseTo(2 * logical.rect.width, 5)
  })

  it('heals a slicer selection after its pivot is reset to the saved view', async () => {
    const doc = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(`${__dirname}/fixtures/xlsx-pivots/controls/controls-seed.xlsx`))))
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    const group = await ui.findByRole('group', { name: /Product slicer/i })
    fireEvent.click(within(group).getByRole('button', { name: 'Apples, selected' }))
    await waitFor(() => expect(within(group).getByRole('button', { name: 'Bananas, not selected' })).toBeTruthy())
    // reset the pivot through its field menu
    let reset = false
    for (const button of ui.container.querySelectorAll<HTMLButtonElement>('button[data-officeview-ui="pivot-button"]')) {
      fireEvent.click(button)
      const menu = await ui.findByRole('dialog').catch(() => null)
      const resetButton = menu && within(menu).queryByRole('button', { name: 'Reset to saved view' })
      if (resetButton) { fireEvent.click(resetButton); reset = true; break }
      const cancelButton = menu && within(menu).queryByRole('button', { name: 'Cancel' })
      if (cancelButton) fireEvent.click(cancelButton)
    }
    expect(reset).toBe(true)
    // the pivot is restored while the slicer model still holds Apples only
    await waitFor(() => expect(cellText(ui.container, 7, 4)).toContain('150'))
    fireEvent.click(within(group).getByRole('button', { name: 'Clear Filter' }))
    await waitFor(() => expect(within(group).getByRole('button', { name: 'Apples, selected' }).getAttribute('aria-pressed')).toBe('true'))
    ui.unmount()
  }, 60_000)
})
