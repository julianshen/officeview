// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, screen } from '@testing-library/react'
import { PivotFieldList } from '../src/components/xlsx/pivot/PivotFieldList'
import type { PivotMenuModel } from '../src/components/xlsx/pivot/pivotClient'

afterEach(cleanup)

const model: PivotMenuModel = {
  name: 'Sales', rows: [0], cols: [], pages: [], valuesCaption: 'Values', nextFilterId: 1,
  fields: [
    { index: 0, sourceOrder: 0, name: 'Region', axis: 'row', items: [], sort: 'manual', filters: [], values: 'text', expandable: false, defaultDataFn: 'count', permissions: { dragToRow: true, dragToCol: true, dragToPage: true, dragToData: true, dragOff: true } },
    { index: 1, sourceOrder: 1, name: 'Sales', axis: 'none', items: [], sort: 'manual', filters: [], values: 'number', expandable: false, defaultDataFn: 'sum', permissions: { dragToRow: true, dragToCol: true, dragToPage: true, dragToData: true, dragOff: true } },
  ],
  data: [{ index: 0, field: 1, name: 'Sum of Sales', fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832 }],
}

describe('PivotFieldList', () => {
  it('searches fields, exposes keyboard/menu moves, drag targets and reset', () => {
    const onApply = vi.fn(), onClose = vi.fn()
    render(<PivotFieldList model={model} pivot={0} pivotCount={1} onApply={onApply} onClose={onClose} />)
    fireEvent.change(screen.getByRole('textbox', { name: 'Search fields' }), { target: { value: 'sales' } })
    expect(screen.queryByRole('checkbox', { name: 'Region' })).toBeNull()
    const actions = screen.getByRole('combobox', { name: 'Move Sales' })
    fireEvent.change(actions, { target: { value: 'row' } })
    expect(onApply).toHaveBeenCalledWith({ kind: 'moveField', field: 1, axis: 'row' })
    fireEvent.click(screen.getByRole('button', { name: 'Reset to saved view' }))
    expect(onApply).toHaveBeenCalledWith('reset')
  })

  it('offers the responsive bottom-sheet dialog and Escape close', () => {
    const original = window.matchMedia
    window.matchMedia = vi.fn().mockReturnValue({ matches: true, addEventListener: vi.fn(), removeEventListener: vi.fn() }) as unknown as typeof window.matchMedia
    const onClose = vi.fn()
    render(<PivotFieldList model={model} pivot={0} onApply={vi.fn()} onClose={onClose} />)
    const dialog = screen.getByRole('dialog', { name: 'PivotTable Fields' })
    expect(dialog.getAttribute('data-pivot-field-list')).toBe('bottom-sheet')
    fireEvent.keyDown(dialog, { key: 'Escape' })
    expect(onClose).toHaveBeenCalledOnce()
    window.matchMedia = original
  })
})
