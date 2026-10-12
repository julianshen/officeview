/**
 * S13 guided VoiceOver pass in Safari 26.5.2 (docs/validation/interactive-viewer.md, S13):
 * V1 — every cell value was spoken twice ("B3, value 2023-06-09 2023-06-09"):
 *      the gridcell's aria-label carries the value AND its visually hidden
 *      text child was still exposed, so VoiceOver read name + contents.
 * V2 — Shift+Arrow extension announced only the target cell, never that a
 *      range is selected. Spec §9.2: names carry selection descriptions and
 *      the live region reports selection.
 */
import { describe, expect, it } from 'vitest'
import { fireEvent, render } from '@testing-library/react'
import { GridSemantics, gridNavigationTransition, type GridTransition } from '../src/components/xlsx/GridSemantics'
import { selectCell, selectHeader } from '../src/xlsx/view/selection'
import { gridFixture } from './helpers/xlsx-grid-fixture'

describe('V1: a cell value is exposed once', () => {
  it('no gridcell exposes text beside its accessible name', () => {
    const ui = render(<GridSemantics {...gridFixture()} />)
    const cells = [...ui.container.querySelectorAll('[role=gridcell]')]
    expect(cells.length).toBeGreaterThan(0)
    for (const cell of cells) {
      // Text nodes reachable without crossing aria-hidden would be read after the name.
      const exposed = [...cell.querySelectorAll('*')].filter(node => node.textContent && !node.closest('[aria-hidden=true]'))
      expect(exposed.map(node => node.textContent)).toEqual([])
    }
    // The name contract is unchanged.
    expect(ui.getByRole('gridcell', { name: 'A2, value 10; original A4' })).toBeTruthy()
  })
})

describe('V2: range selection is announced', () => {
  it('Shift+Arrow reports the selected range in the live announcement', () => {
    const props = gridFixture()
    const transition = gridNavigationTransition(props, 'ArrowRight', { extend: true })
    // Column B is collapsed in the fixture, so the extension lands on C.
    expect(transition.announcement).toBe('C2, value 10; original C4; selected A2:C2')
    const plain = gridNavigationTransition(props, 'ArrowRight')
    expect(plain.announcement).toBe('C2, value 10; original C4')
  })

  it('a keyboard Shift+Arrow on the grid publishes that announcement', () => {
    const seen: GridTransition[] = []
    const ui = render(<GridSemantics {...gridFixture({ onChange: transition => seen.push(transition) })} />)
    const grid = ui.getByRole('grid')
    grid.focus()
    fireEvent.keyDown(grid, { key: 'ArrowRight', shiftKey: true })
    expect(seen.at(-1)?.announcement).toBe('C2, value 10; original C4; selected A2:C2')
  })

  it('cells of a multi-cell selection carry a "selected" description; a lone active cell does not', () => {
    const props = gridFixture()
    const single = render(<GridSemantics {...props} />)
    expect(single.getByRole('gridcell', { name: 'A2, value 10; original A4' }).getAttribute('aria-description') ?? '').not.toContain('selected')
    single.unmount()
    const extended = selectCell(props.selection, props.projection, { row: 3, col: 2 }, { extend: true })
    const ui = render(<GridSemantics {...props} selection={extended} />)
    for (const name of ['A2, value 10; original A4', 'C2, value 10; original C4']) {
      const cell = ui.getByRole('gridcell', { name })
      expect(cell.getAttribute('aria-selected')).toBe('true')
      expect(cell.getAttribute('aria-description')).toContain('selected')
    }
  })
})

describe('V2 review round 3: truthful extension wording and cardinality', () => {
  it('extending a column-header selection announces whole columns, not a cell range', () => {
    const props = gridFixture()
    const columnA = selectHeader(props.selection, props.projection, 'column', 0)
    const transition = gridNavigationTransition({ ...props, selection: columnA }, 'ArrowRight', { extend: true })
    expect(transition.announcement).toMatch(/; selected columns A:C$/)
  })
  it('extending a row-header selection announces rows by their displayed labels', () => {
    const props = gridFixture()
    const row = selectHeader(props.selection, props.projection, 'row', 3) // source row 4 shows as row 2
    const transition = gridNavigationTransition({ ...props, selection: row }, 'ArrowDown', { extend: true })
    expect(transition.announcement).toMatch(/; selected rows 2:\d+$/)
  })
  it('a single-cell bitset selection is one cell: no "selected" description', () => {
    const props = gridFixture()
    const at = { row: 3, col: 0 }
    const bitset = { active: at, anchor: at, ranges: [{ rows: { kind: 'bitset' as const, space: 'source' as const, r0: 3, r1: 3, words: Uint32Array.of(1), count: 1 }, cols: [{ c0: 0, c1: 0 }] }] }
    const ui = render(<GridSemantics {...props} selection={bitset} />)
    expect(ui.getByRole('gridcell', { name: 'A2, value 10; original A4' }).getAttribute('aria-description') ?? '').not.toContain('selected')
  })
})
