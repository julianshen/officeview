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
})
