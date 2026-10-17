import { afterEach, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render } from '@testing-library/react'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { gridFixture } from './helpers/xlsx-grid-fixture'
afterEach(() => { cleanup(); vi.restoreAllMocks() })
it('routes resize keys from the actual inline touch disclosure to its source column', () => {
  const previous = window.matchMedia
  window.matchMedia = vi.fn(query => ({ matches: query.includes('pointer: coarse'), media: query, onchange: null, addEventListener: vi.fn(), removeEventListener: vi.fn(), addListener: vi.fn(), removeListener: vi.fn(), dispatchEvent: () => true })) as typeof window.matchMedia
  try {
    const resize = vi.fn()
    const ui = render(<GridSemantics {...gridFixture()} onResize={resize} onSortCycle={vi.fn()} />)
    const grid = ui.getByRole('grid')
    act(() => grid.focus())
    fireEvent.keyDown(grid, { key: 'F2' })
    const disclosure = ui.getByRole('button', { name: 'Actions for column A' })
    expect(disclosure.dataset.headerDisclosure).toBe('0')
    act(() => disclosure.focus())
    expect(fireEvent.keyDown(disclosure, { key: 'ArrowRight', altKey: true, shiftKey: true })).toBe(false)
    expect(resize).toHaveBeenCalledExactlyOnceWith({ kind: 'column', index: 0, delta: 8 })
  } finally { window.matchMedia = previous }
})
