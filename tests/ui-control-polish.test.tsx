/**
 * Task 8 fix round 1 (jsdom half; computed styles in
 * tests/browser/ui-control-polish.spec.ts): selects carry the drawn chevron
 * and room for it, a disabled primary drops the accent fill, dialog buttons
 * keep a gap, and dialog controls use the accent focus ring.
 */
import { describe, expect, it, vi } from 'vitest'
import { render } from '@testing-library/react'
import { CHECKBOX_CLS, CONTROL_BTN_CLS, CONTROL_CLS, CONTROL_PRIMARY_BTN_CLS, SELECT_CLS } from '../src/components/xlsx/controlUi'
import { SheetNavigator } from '../src/components/xlsx/SheetNavigator'
import { GridHeaders } from '../src/components/xlsx/GridHeaders'
import { gridFixture } from './helpers/xlsx-grid-fixture'

const has = (cls: string, ...want: string[]) => want.filter(w => !cls.split(/\s+/).includes(w))
const FOCUS = ['ov:focus-visible:outline-2', 'ov:focus-visible:outline-solid', 'ov:focus-visible:outline-accent']

describe('sheet control polish classes', () => {
  it('selects draw a chevron and reserve room for it', () => {
    expect(has(SELECT_CLS, 'ov:select-chevron', 'ov:px-2', 'ov:pr-6', ...FOCUS)).toEqual([])
  })
  it('the disabled primary button has no accent fill', () => {
    expect(has(CONTROL_PRIMARY_BTN_CLS, 'ov:disabled:bg-hover', 'ov:disabled:text-muted', ...FOCUS)).toEqual([])
  })
  it('dialog buttons and checkboxes keep a gap; inputs use the accent ring', () => {
    expect(has(CONTROL_BTN_CLS, 'ov:mr-1', ...FOCUS)).toEqual([])
    expect(has(CONTROL_PRIMARY_BTN_CLS, 'ov:mr-1')).toEqual([])
    expect(has(CHECKBOX_CLS, 'ov:mr-2')).toEqual([])
    expect(has(CONTROL_CLS, ...FOCUS)).toEqual([])
  })
  it('the sheet navigator selects use the chevron', () => {
    const ui = render(<SheetNavigator sheets={[{ id: 's1', name: 'One' }, { id: 's2', name: 'Two' }]} activeSheetId="s1" documentId="d" onSheetChange={vi.fn()} onNavigate={vi.fn()} />)
    for (const name of ['Worksheet', 'Address space']) {
      const select = ui.getByRole('combobox', { name })
      expect(has(select.className, 'ov:select-chevron', 'ov:pr-6', ...FOCUS), name).toEqual([])
    }
  })
  it('sheet navigator label text is visually hidden on a coarse pointer (controls keep aria-labels)', () => {
    const ui = render(<SheetNavigator sheets={[{ id: 's1', name: 'One' }]} activeSheetId="s1" documentId="d" onSheetChange={vi.fn()} onNavigate={vi.fn()} />)
    for (const text of ['Worksheet', 'Address space', 'Go To cell']) {
      const span = [...ui.container.querySelectorAll('label > span')].find(el => el.textContent?.trim() === text)
      expect(span?.className, text).toContain('ov:coarse:sr-only')
      expect(ui.getByLabelText(text, { selector: 'select, input' })).toBeTruthy()
    }
  })
  it('desktop header actions are quiet unless active; the header is the hover group', () => {
    const p = gridFixture()
    const ui = render(<div role="grid"><GridHeaders projection={p.projection} window={p.window} viewport={p.viewport} columns={[0, 2]} selection={p.selection} sort={null} controlMode
      filter={{ range: { r0: 0, r1: 3, c0: 0, c1: 2, headerRows: 1, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }}
      onSelect={vi.fn()} onSelectAll={vi.fn()} onSortCycle={vi.fn()} onFilter={vi.fn()} onSortDialog={vi.fn()} /></div>)
    const cls = (name: string | RegExp) => ui.getByRole('button', { name }).className
    expect(ui.getByRole('columnheader', { name: /^Column C/ }).className).toContain('ov:group')
    for (const name of ['Cycle sort column C', /^Filter column C/, 'More sort options for C']) {
      expect(has(cls(name), 'ov:opacity-0', 'ov:group-hover:opacity-100', 'ov:group-focus-within:opacity-100', 'ov:focus-visible:opacity-100'), String(name)).toEqual([])
    }
    // Column A has an applied filter: its funnel is never quiet.
    expect(cls(/^Filter column A/)).not.toContain('ov:opacity-0')
  })
})
