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
    expect(document.activeElement).toBe(screen.getByRole('textbox', { name: 'Search fields' }))
    fireEvent.change(screen.getByRole('textbox', { name: 'Search fields' }), { target: { value: 'sales' } })
    expect(screen.queryByRole('checkbox', { name: 'Region' })).toBeNull()
    const actions = screen.getByRole('combobox', { name: 'Move Sales' })
    fireEvent.change(actions, { target: { value: 'row' } })
    expect(onApply).toHaveBeenCalledWith({ kind: 'moveField', field: 1, axis: 'row' })
    fireEvent.click(screen.getByRole('button', { name: 'Reset to saved view' }))
    expect(onApply).toHaveBeenCalledWith('reset')
  })

  it('uses a linear number of field lookup checks while rendering a large field list', () => {
    const count = 512
    let predicates = 0
    const fields = new Proxy(Array.from({ length: count }, (_, index) => ({
      ...model.fields[0]!, index, sourceOrder: index, name: `Field ${index}`, axis: 'none' as const,
    })), {
      get(target, property, receiver) {
        if (property === 'find') return (predicate: (value: (typeof target)[number], index: number, array: typeof target) => unknown) =>
          Array.prototype.find.call(target, (value: (typeof target)[number], index: number, array: typeof target) => {
            predicates++
            return predicate(value, index, array)
          })
        return Reflect.get(target, property, receiver)
      },
    })
    render(<PivotFieldList model={{ ...model, fields }} pivot={0} onApply={vi.fn()} onClose={vi.fn()} />)
    expect(predicates).toBeLessThanOrEqual(count * 4)
    expect(screen.getAllByRole('checkbox')).toHaveLength(200)
    expect(screen.getByRole('status').textContent).toContain('512 fields')
  })

  it('sorts the field list and reorders Values chips with their existing data-field identity', () => {
    const onApply = vi.fn()
    const sortedModel: PivotMenuModel = {
      ...model, fieldListSortAscending: true, rows: [0], cols: [-2],
      fields: [
        { ...model.fields[0]!, index: 0, sourceOrder: 0, name: 'Zulu', axis: 'none' },
        { ...model.fields[1]!, index: 1, sourceOrder: 1, name: 'Alpha', axis: 'none' },
      ],
      data: [
        { index: 0, field: 0, name: 'Sum of Zulu', fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832 },
        { index: 1, field: 1, name: 'Sum of Alpha', fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832 },
      ],
    }
    render(<PivotFieldList model={sortedModel} pivot={0} onApply={onApply} onClose={vi.fn()} />)
    expect(screen.getAllByRole('checkbox').map(node => node.getAttribute('aria-label'))).toEqual(['Alpha', 'Zulu'])
    const values = screen.getByRole('group', { name: 'Values' })
    const chips = values.querySelectorAll<HTMLElement>('[draggable="true"]')
    expect(chips).toHaveLength(2)
    const dataTransfer = { setData: vi.fn(), getData: vi.fn().mockReturnValue('0'), effectAllowed: 'move' }
    fireEvent.change(screen.getByRole('combobox', { name: 'Actions for Sum of Zulu' }), { target: { value: 'down' } })
    expect(onApply).toHaveBeenCalledWith({ kind: 'moveDataField', dataIndex: 0, toIndex: 1 })
    onApply.mockClear()
    fireEvent.dragStart(chips[0]!, { dataTransfer })
    fireEvent.drop(chips[1]!, { dataTransfer })
    expect(onApply).toHaveBeenCalledWith({ kind: 'moveDataField', dataIndex: 0, toIndex: 1 })

    onApply.mockClear()
    const columns = screen.getByRole('group', { name: 'Columns' })
    const marker = columns.querySelector<HTMLElement>('[draggable="true"]')!
    fireEvent.dragStart(marker, { dataTransfer })
    fireEvent.drop(screen.getByRole('group', { name: 'Rows' }), { dataTransfer })
    expect(onApply).toHaveBeenCalledWith({ kind: 'moveField', field: -2, axis: 'row', index: 1 })

    onApply.mockClear()
    const alpha = screen.getByRole('checkbox', { name: 'Alpha' }).parentElement!.parentElement!
    fireEvent.dragStart(alpha, { dataTransfer })
    fireEvent.drop(screen.getByRole('group', { name: 'Rows' }), { dataTransfer })
    expect(onApply).toHaveBeenCalledWith({ kind: 'moveField', field: 1, axis: 'row', index: 1 })

    onApply.mockClear()
    const rowChip = screen.getByRole('group', { name: 'Rows' }).querySelector<HTMLElement>('[draggable="true"]')!
    fireEvent.dragStart(rowChip, { dataTransfer })
    fireEvent.drop(screen.getByRole('group', { name: 'Columns' }), { dataTransfer })
    expect(onApply).toHaveBeenCalledWith({ kind: 'moveField', field: 0, axis: 'col', index: 1 })
  })

  it('offers the responsive bottom-sheet dialog and Escape close', () => {
    const original = window.matchMedia
    window.matchMedia = vi.fn().mockReturnValue({ matches: true, addEventListener: vi.fn(), removeEventListener: vi.fn() }) as unknown as typeof window.matchMedia
    const onClose = vi.fn()
    render(<PivotFieldList model={model} pivot={0} onApply={vi.fn()} onClose={onClose} />)
    const dialog = screen.getByRole('dialog', { name: 'PivotTable Fields' })
    expect(dialog.getAttribute('data-pivot-field-list')).toBe('bottom-sheet')
    for (const node of dialog.querySelectorAll<HTMLElement>('button,input:not([type="checkbox"]),select')) {
      expect(Number.parseFloat(node.style.minHeight || '0')).toBeGreaterThanOrEqual(44)
      expect(Number.parseFloat(node.style.minWidth || '0')).toBeGreaterThanOrEqual(44)
    }
    fireEvent.keyDown(dialog, { key: 'Escape' })
    expect(onClose).toHaveBeenCalledOnce()
    window.matchMedia = original
  })
})
