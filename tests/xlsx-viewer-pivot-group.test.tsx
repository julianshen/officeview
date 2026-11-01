// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react'
import { PivotFieldList } from '../src/components/xlsx/pivot/PivotFieldList'
import { PivotFieldMenu } from '../src/components/xlsx/pivot/PivotFieldMenu'
import type { PivotMenuModel } from '../src/components/xlsx/pivot/pivotClient'

afterEach(cleanup)

const field = (index: number, name: string, group: NonNullable<PivotMenuModel['fields'][number]['group']>) => ({
  index, sourceOrder: index, name, axis: 'row' as const, items: [], sort: 'manual' as const, filters: [], values: 'date' as const,
  expandable: false, group, defaultDataFn: 'count' as const,
  permissions: { dragToRow: true, dragToCol: true, dragToPage: true, dragToData: true, dragOff: true },
})
const model: PivotMenuModel = {
  name: 'Grouping', rows: [0, 1, 2], cols: [], pages: [], valuesCaption: 'Values', nextFilterId: 1, date1904: false,
  fields: [
    field(0, 'Date', { kind: 'date', grouped: false, baseField: 0, bounds: { min: 43831, max: 43862 }, date: { periods: [], autoStart: true, autoEnd: true, daysInterval: 1 } }),
    field(1, 'Grouped Date', { kind: 'date', grouped: true, baseField: 1, date: { periods: ['years', 'months'], startDate: 43831, endDate: 43862, autoStart: false, autoEnd: false, daysInterval: 1 } }),
    field(2, 'Manual Group', { grouped: true, baseField: 2 }),
  ], data: [],
}

describe('pivot grouping UI', () => {
  it('shows verified Group and Ungroup entries in the field menu only for supported fields', () => {
    const onApply = vi.fn()
    render(<PivotFieldMenu owner={{ docId: 'doc', sheetId: '1', valueGen: 0, fontGen: 0, viewGen: 0 }} model={model}
      verified={new Set(['groupDate', 'ungroup'])} field={0} onApply={onApply} onClose={vi.fn()} onNotice={vi.fn()} />)
    expect(screen.getByRole('button', { name: 'Group…' })).toBeTruthy()
    expect(screen.queryByRole('button', { name: 'Ungroup' })).toBeNull()
    cleanup()
    render(<PivotFieldMenu owner={{ docId: 'doc', sheetId: '1', valueGen: 0, fontGen: 0, viewGen: 0 }} model={model}
      verified={new Set(['ungroup'])} field={1} onApply={onApply} onClose={vi.fn()} onNotice={vi.fn()} />)
    expect(screen.queryByRole('button', { name: 'Group…' })).toBeNull()
    expect(screen.getByRole('button', { name: 'Ungroup' })).toBeTruthy()
  })

  it('offers Group only for verified date/number fields and Ungroup only for grouped fields', () => {
    render(<PivotFieldList model={model} pivot={0} verified={new Set(['groupDate', 'ungroup'])} onApply={vi.fn()} onClose={vi.fn()} />)
    const dateActions = screen.getByRole('combobox', { name: 'Actions for Date' })
    const groupedActions = screen.getByRole('combobox', { name: 'Actions for Grouped Date' })
    const manualActions = screen.getByRole('combobox', { name: 'Actions for Manual Group' })
    expect([...dateActions.querySelectorAll('option')].map(option => option.textContent)).toContain('Group…')
    expect([...dateActions.querySelectorAll('option')].map(option => option.textContent)).not.toContain('Ungroup')
    expect([...groupedActions.querySelectorAll('option')].map(option => option.textContent)).toContain('Ungroup')
    expect([...manualActions.querySelectorAll('option')].map(option => option.textContent)).not.toContain('Group…')
    expect([...manualActions.querySelectorAll('option')].map(option => option.textContent)).toContain('Ungroup')
  })

  it('does not offer grouping actions for fields in Filters', () => {
    const pageModel: PivotMenuModel = { ...model, rows: [1, 2], pages: [0] }
    render(<PivotFieldList model={pageModel} pivot={0} verified={new Set(['groupDate', 'ungroup'])} onApply={vi.fn()} onClose={vi.fn()} />)
    const actions = screen.getByRole('combobox', { name: 'Actions for Date' })
    expect([...actions.querySelectorAll('option')].map(option => option.textContent)).not.toContain('Group…')
    expect([...actions.querySelectorAll('option')].map(option => option.textContent)).not.toContain('Ungroup')
  })

  it('announces when the grouping apply is refused', async () => {
    const onApply = vi.fn().mockResolvedValue(false)
    render(<PivotFieldList model={model} pivot={0} verified={new Set(['groupDate'])} onApply={onApply} onClose={vi.fn()} />)
    fireEvent.change(screen.getByRole('combobox', { name: 'Actions for Date' }), { target: { value: 'group' } })
    fireEvent.click(screen.getByLabelText('Months'))
    await act(async () => { fireEvent.click(screen.getByRole('button', { name: 'OK' })) })
    expect(await screen.findByRole('alert')).toBeTruthy()
    expect(screen.getByRole('dialog')).toBeTruthy()
  })

  it('prefills date bounds, gates day interval, applies regroup values, and restores focus on Escape', async () => {
    const onApply = vi.fn().mockResolvedValue(true)
    render(<PivotFieldList model={model} pivot={0} verified={new Set(['groupDate', 'ungroup'])} onApply={onApply} onClose={vi.fn()} />)
    const action = screen.getByRole('combobox', { name: 'Actions for Grouped Date' })
    action.focus()
    fireEvent.change(action, { target: { value: 'group' } })
    const startingAt = screen.getByLabelText('Starting at') as HTMLInputElement
    expect(startingAt.value).toBe('2020-01-01')
    expect((screen.getByLabelText('Ending at') as HTMLInputElement).value).toBe('2020-02-01')
    const days = screen.getByLabelText('Days') as HTMLInputElement
    const interval = screen.getByLabelText('Number of days') as HTMLInputElement
    expect(interval.disabled).toBe(true)
    fireEvent.click(days)
    fireEvent.click(screen.getByLabelText('Months'))
    fireEvent.click(screen.getByLabelText('Years'))
    expect(interval.disabled).toBe(false)
    fireEvent.click(screen.getByLabelText('Years'))
    fireEvent.change(interval, { target: { value: '3' } })
    await act(async () => { fireEvent.click(screen.getByRole('button', { name: 'OK' })) })
    expect((await screen.findByRole('alert')).textContent).toContain('Number of days must be 1')
    fireEvent.change(interval, { target: { value: '1' } })
    await act(async () => { fireEvent.click(screen.getByRole('button', { name: 'OK' })) })
    await vi.waitFor(() => expect(onApply).toHaveBeenCalledWith({ kind: 'groupDate', field: 1, periods: ['days', 'years'], startDate: 43831, endDate: 43862, daysInterval: 1 }))
    await vi.waitFor(() => expect(screen.queryByRole('dialog', { name: 'Group Grouped Date' })).toBeNull())

    action.focus()
    fireEvent.change(action, { target: { value: 'group' } })
    const dialog = screen.getByRole('dialog', { name: 'Group Grouped Date' })
    const focusable = dialog.querySelectorAll<HTMLElement>('button:not(:disabled),input:not(:disabled)')
    const last = focusable.item(focusable.length - 1)!
    const first = focusable.item(0)!
    const focusFirst = vi.spyOn(first, 'focus')
    last.focus()
    const tab = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })
    last.dispatchEvent(tab)
    expect(tab.defaultPrevented).toBe(true)
    expect(focusFirst).toHaveBeenCalledOnce()
    fireEvent.keyDown(dialog, { key: 'Escape' })
    expect(screen.queryByRole('dialog', { name: 'Group Grouped Date' })).toBeNull()
    expect(document.activeElement).toBe(action)
  })

  it('validates number bounds and positive interval before apply', async () => {
    const onApply = vi.fn().mockResolvedValue(true)
    const numeric: PivotMenuModel = { ...model, rows: [0], fields: [field(0, 'Amount', { kind: 'number', grouped: false, baseField: 0, bounds: { min: -5, max: 19 }, number: { by: 5, autoStart: true, autoEnd: true } })] }
    render(<PivotFieldList model={numeric} pivot={0} verified={new Set(['groupNumber'])} onApply={onApply} onClose={vi.fn()} />)
    fireEvent.change(screen.getByRole('combobox', { name: 'Actions for Amount' }), { target: { value: 'group' } })
    expect((screen.getByLabelText('Starting at') as HTMLInputElement).value).toBe('-5')
    expect((screen.getByLabelText('Ending at') as HTMLInputElement).value).toBe('19')
    fireEvent.change(screen.getByLabelText('By'), { target: { value: '0' } })
    await act(async () => { fireEvent.click(screen.getByRole('button', { name: 'OK' })) })
    expect((await screen.findByRole('alert')).textContent).toContain('By must be greater than zero')
    expect(onApply).not.toHaveBeenCalled()
  })
})
