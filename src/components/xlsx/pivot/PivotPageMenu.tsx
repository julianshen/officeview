/**
 * A pivot page field's (report filter's) menu (plan Task 28, spec §6.7): (All), one item, or several items when
 * "Select multiple items" is on; clearing the field's filter; Reset to saved view. The same shell, lifetime and 44 px
 * targets as the field menu. Entries follow the session's verified kinds (`page`, `clearFilters`).
 */
import { useState, type ReactElement } from 'react'
import { controlTarget, CHECKBOX_CLS } from '../controlUi'
import { MenuButton, PivotMenuShell, usePivotMenuLifetime, type PivotMenuProps } from './PivotFieldMenu'

export function PivotPageMenu(props: PivotMenuProps): ReactElement {
  usePivotMenuLifetime(props)
  const { model, verified, onApply } = props
  const field = model.fields.find(entry => entry.index === props.field)
  const initial = field?.page ?? 'all'
  const [multiple, setMultiple] = useState(initial !== 'all' && initial.length > 1)
  const [selected, setSelected] = useState<'all' | number[]>(initial === 'all' ? 'all' : [...initial])
  if (!field) return <PivotMenuShell title="Pivot table" onClose={props.onClose}><p>This field is no longer a report filter.</p><MenuButton onClick={props.onClose}>Close</MenuButton></PivotMenuShell>
  const isOn = (item: number) => selected !== 'all' && selected.includes(item)
  const toggle = (item: number, on: boolean) => setSelected(current => {
    const list = current === 'all' ? [] : current.filter(entry => entry !== item)
    return on ? [...list, item].sort((a, b) => a - b) : list
  })
  const row = { ...controlTarget, display: 'flex', alignItems: 'center' } as const
  const empty = selected !== 'all' && selected.length === 0
  return <PivotMenuShell title={field.name} onClose={props.onClose}>
    {verified.has('page') && <fieldset><legend>Show</legend>
      <label style={row}><input type="checkbox" className={CHECKBOX_CLS} checked={multiple} aria-label="Select multiple items"
        onChange={event => { setMultiple(event.target.checked); if (!event.target.checked && selected !== 'all' && selected.length > 1) setSelected([selected[0]!]) }} />Select multiple items</label>
      {multiple ? <>
        <label style={row}><input type="checkbox" className={CHECKBOX_CLS} aria-label="(All)" checked={selected === 'all'} onChange={event => setSelected(event.target.checked ? 'all' : [])} />(All)</label>
        {field.items.map(item => <label key={item.index} style={row}><input type="checkbox" className={CHECKBOX_CLS} aria-label={item.label || '(blank)'} checked={selected === 'all' || isOn(item.index)}
          onChange={event => { if (selected === 'all') setSelected(field.items.map(entry => entry.index).filter(entry => entry !== item.index)); else toggle(item.index, event.target.checked) }} />{item.label || '(blank)'}</label>)}
      </> : <div role="radiogroup" aria-label={`${field.name} item`}>
        <label style={row}><input type="radio" name={`pivot-page-${field.index}`} aria-label="(All)" checked={selected === 'all'} onChange={() => setSelected('all')} />(All)</label>
        {field.items.map(item => <label key={item.index} style={row}><input type="radio" name={`pivot-page-${field.index}`} aria-label={item.label || '(blank)'} checked={isOn(item.index)} onChange={() => setSelected([item.index])} />{item.label || '(blank)'}</label>)}
      </div>}
      {empty && <p>Select at least one item.</p>}
      <MenuButton primary disabled={empty} onClick={() => onApply({ kind: 'page', field: field.index, selected })}>OK</MenuButton>
    </fieldset>}
    {verified.has('clearFilters') && <MenuButton onClick={() => onApply({ kind: 'clearFilters', field: field.index })}>{`Clear filter from ${field.name}`}</MenuButton>}
    <div><MenuButton onClick={() => onApply('reset')}>Reset to saved view</MenuButton><MenuButton onClick={props.onClose}>Cancel</MenuButton></div>
  </PivotMenuShell>
}
