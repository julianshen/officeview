/**
 * A pivot row or column field's menu (plan Task 28, spec §6.7): sort by label or by a value column, expand and
 * collapse, label/date, value and top-10 filters, clearing the field's filters, the item checklist (search, select
 * all, clear) and Reset to saved view. A sibling of `FilterDropdown` (which is bound to the sheet `ControlsClient`):
 * the same shell (desktop dropdown, bottom sheet on a coarse pointer or a narrow pane), 44 px targets, the owner
 * lifetime rule, and `ConditionEditor` for the conditions.
 *
 * The menu builds a `PivotDelta` and hands it to `onApply`; the host closes the menu before it applies (the caller
 * commits), so a commit's new owner never reads as "changed under an open menu". An entry exists only for an
 * interaction kind in the session's verified set (spec §9: absent, not greyed); Reset is the session's own reset and
 * always present. Bottom-N is not offered: it refuses until Excel's behaviour is measured.
 */
import { useEffect, useMemo, useRef, useState, type CSSProperties, type KeyboardEvent, type ReactElement, type ReactNode } from 'react'
import type { CacheIdentity } from '../../../xlsx/view/contracts'
import type { ColumnClassification } from '../../../xlsx/view/filter'
import { encodeCivil } from '../../../xlsx/formula/serial'
import { ConditionEditor, conditionInput, typedCondition, type ConditionInput } from '../ConditionEditor'
import { useControlMedia } from '../controlMedia'
import { controlTarget, CONTROL_BTN_CLS, CONTROL_CLS, CONTROL_PRIMARY_BTN_CLS, CHECKBOX_CLS, SELECT_CLS } from '../controlUi'
import { dateFilter, labelFilter, top10Filter, valueFilter, type DateOp, type LabelOp, type PivotDelta, type PivotDeltaKind, type PivotMenuField, type PivotMenuModel, type ValueOp } from './pivotClient'

export interface PivotMenuProps {
  /** The sheet's owner: a change while the menu is open closes it (FilterDropdown's lifetime rule). */
  owner: CacheIdentity
  model: PivotMenuModel
  verified: ReadonlySet<PivotDeltaKind>
  field: number
  onApply(delta: PivotDelta | 'reset'): void
  onClose(): void
  onNotice(message: string): void
}

/** The owner lifetime: an owner change while open closes the menu with FilterDropdown's notice. */
export function usePivotMenuLifetime(props: Pick<PivotMenuProps, 'owner' | 'onClose' | 'onNotice'>): void {
  const key = JSON.stringify(props.owner), previous = useRef(key), latest = useRef(props)
  latest.current = props
  useEffect(() => {
    if (previous.current === key) return
    previous.current = key
    latest.current.onNotice('Source or view changed; reopen and retry.')
    latest.current.onClose()
  }, [key])
}

/** The menu's shell: a dropdown under its button, or a bottom sheet on a coarse pointer or a narrow pane. */
export function PivotMenuShell({ title, onClose, children }: { title: string; onClose(): void; children: ReactNode }): ReactElement {
  const bottomSheet = useControlMedia('(pointer: coarse), (max-width: 640px)')
  const root = useRef<HTMLDivElement>(null)
  useEffect(() => { root.current?.querySelector<HTMLElement>('button, input, select')?.focus() }, [])
  const onKeyDown = (event: KeyboardEvent) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') return
    event.stopPropagation()
    if (event.key === 'Escape') { event.preventDefault(); onClose() }
  }
  const style: CSSProperties = {
    position: bottomSheet ? 'fixed' : 'relative', bottom: bottomSheet ? 0 : undefined, left: bottomSheet ? 0 : undefined, zIndex: bottomSheet ? 4 : undefined,
    width: bottomSheet ? '100%' : 'min(420px, calc(100vw - 16px))', maxHeight: bottomSheet ? '85dvh' : undefined, overflowY: bottomSheet ? 'auto' : undefined,
    padding: 12, paddingBottom: 'calc(12px + env(safe-area-inset-bottom, 0px))', boxSizing: 'border-box',
  }
  return <div ref={root} data-pivot-menu={bottomSheet ? 'bottom-sheet' : 'dropdown'} data-officeview-ui="pivot-menu" onKeyDown={onKeyDown}
    onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
    className={bottomSheet ? 'ov:rounded-t-[12px] ov:shadow-popover ov:bg-surface ov:border-t ov:border-border' : undefined} style={style}>
    <h2>{title}</h2>
    {children}
  </div>
}

export const MenuButton = ({ children, onClick, primary, disabled, label }: { children: ReactNode; onClick(): void; primary?: boolean; disabled?: boolean; label?: string }) =>
  <button type="button" className={primary ? CONTROL_PRIMARY_BTN_CLS : CONTROL_BTN_CLS} style={controlTarget} disabled={disabled} aria-label={label} onClick={onClick}>{children}</button>

/** The item checklist (FilterDropdown's values list): search, Select all and Clear act on the matching items. */
export function ItemChecklist({ field, onApply }: { field: PivotMenuField; onApply(hidden: number[]): void }): ReactElement {
  const [text, setText] = useState('')
  const [shown, setShown] = useState(() => new Set(field.items.filter(item => !item.hidden).map(item => item.index)))
  const matching = useMemo(() => { const needle = text.trim().toLowerCase(); return field.items.filter(item => !needle || item.label.toLowerCase().includes(needle)) }, [field, text])
  const setAll = (selected: boolean) => setShown(current => { const next = new Set(current); for (const item of matching) if (selected) next.add(item.index); else next.delete(item.index); return next })
  const none = shown.size === 0
  return <fieldset><legend>Items</legend>
    <label>Search <input aria-label="Search items" className={CONTROL_CLS} style={controlTarget} value={text} onChange={event => setText(event.target.value)} /></label>
    <div><MenuButton onClick={() => setAll(true)}>Select all</MenuButton><MenuButton onClick={() => setAll(false)}>Clear</MenuButton></div>
    {/* Bounded DOM: at most 200 matching items are listed; search narrows a longer field. */}
    <div style={{ maxHeight: 240, overflowY: 'auto' }}>{matching.slice(0, 200).map(item => <label key={item.index} style={{ ...controlTarget, display: 'flex', alignItems: 'center' }}>
      <input type="checkbox" className={CHECKBOX_CLS} aria-label={item.label || '(blank)'} checked={shown.has(item.index)}
        onChange={event => setShown(current => { const next = new Set(current); if (event.target.checked) next.add(item.index); else next.delete(item.index); return next })} />
      <span>{item.label || '(blank)'}</span>
    </label>)}</div>
    {matching.length > 200 && <p>Showing the first 200 of {matching.length} items; search to narrow.</p>}
    {none && <p>Select at least one item.</p>}
    <MenuButton primary disabled={none} onClick={() => onApply(field.items.filter(item => !shown.has(item.index)).map(item => item.index))}>OK</MenuButton>
  </fieldset>
}

const TEXT: ColumnClassification = { kind: 'text', hasNumeric: false }
const NUMBER: ColumnClassification = { kind: 'number', hasNumeric: true }
const DATE: ColumnClassification = { kind: 'date', hasNumeric: true }
const serialOf = (text: string): number => {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text)
  if (!match) throw new RangeError('Enter a calendar date as YYYY-MM-DD.')
  const serial = encodeCivil({ year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) }, '1900')
  if (typeof serial !== 'number') throw new RangeError('Enter a valid calendar date.')
  return serial
}

export function PivotFieldMenu(props: PivotMenuProps): ReactElement {
  usePivotMenuLifetime(props)
  const { model, verified, onApply } = props
  // A compact label column holds several fields: the menu can switch between the fields of the button's axis.
  const start = model.fields.find(entry => entry.index === props.field)
  const siblings = model.fields.filter(entry => entry.axis === start?.axis && (entry.axis === 'row' || entry.axis === 'col'))
  const [chosen, setChosen] = useState(props.field)
  const field = model.fields.find(entry => entry.index === chosen) ?? start
  const [label, setLabel] = useState<ConditionInput>(() => conditionInput(undefined, 'text'))
  const [date, setDate] = useState<ConditionInput>(() => conditionInput(undefined, 'date'))
  const [value, setValue] = useState<ConditionInput>(() => conditionInput(undefined, 'number'))
  const [valueData, setValueData] = useState(0), [sortData, setSortData] = useState(field?.byData ?? 0)
  const [top, setTop] = useState({ count: '10', kind: 'count' as 'count' | 'percent' | 'sum', data: 0 })
  const [error, setError] = useState('')
  if (!field) return <PivotMenuShell title="Pivot table" onClose={props.onClose}><p>This field is no longer in the pivot table.</p><MenuButton onClick={props.onClose}>Close</MenuButton></PivotMenuShell>
  const has = (kind: PivotDeltaKind) => verified.has(kind)
  const id = model.nextFilterId
  const attempt = (build: () => PivotDelta) => { try { setError(''); onApply(build()) } catch (cause) { setError(cause instanceof Error ? cause.message : 'Check the condition.') } }
  const dataSelect = (value: number, onChange: (next: number) => void, name: string) => <label>{name} <select aria-label={name} className={SELECT_CLS} style={controlTarget} value={value} onChange={event => onChange(Number(event.target.value))}>
    {model.data.map(entry => <option key={entry.index} value={entry.index}>{entry.name}</option>)}</select></label>
  const onAxis = field.axis === 'row' || field.axis === 'col'
  return <PivotMenuShell title={field.name} onClose={props.onClose}>
    {siblings.length > 1 && <label>Select field <select aria-label="Select field" className={SELECT_CLS} style={controlTarget} value={field.index} onChange={event => setChosen(Number(event.target.value))}>
      {siblings.map(entry => <option key={entry.index} value={entry.index}>{entry.name}</option>)}</select></label>}
    {has('sort') && onAxis && <div role="group" aria-label="Sort">
      <MenuButton onClick={() => onApply({ kind: 'sort', field: field.index, order: 'ascending' })}>Sort ascending</MenuButton>
      <MenuButton onClick={() => onApply({ kind: 'sort', field: field.index, order: 'descending' })}>Sort descending</MenuButton>
      {model.data.length > 0 && <div>{dataSelect(sortData, setSortData, 'Sort by value of')}
        <MenuButton onClick={() => onApply({ kind: 'sort', field: field.index, order: 'ascending', byData: sortData })}>Sort ascending by value</MenuButton>
        <MenuButton onClick={() => onApply({ kind: 'sort', field: field.index, order: 'descending', byData: sortData })}>Sort descending by value</MenuButton></div>}
    </div>}
    {has('collapseField') && field.expandable && <div role="group" aria-label="Expand and collapse">
      <MenuButton onClick={() => onApply({ kind: 'collapseField', field: field.index, collapsed: false })}>Expand entire field</MenuButton>
      <MenuButton onClick={() => onApply({ kind: 'collapseField', field: field.index, collapsed: true })}>Collapse entire field</MenuButton>
    </div>}
    {has('collapseAll') && <div role="group" aria-label="Every field">
      <MenuButton onClick={() => onApply({ kind: 'collapseAll', collapsed: false })}>Expand all</MenuButton>
      <MenuButton onClick={() => onApply({ kind: 'collapseAll', collapsed: true })}>Collapse all</MenuButton>
    </div>}
    {has('filter') && <>
      {field.values === 'date' ? <div>
        <ConditionEditor value={date} onChange={setDate} classification={DATE} kinds={['date']} index={1} invalid={false} label="Date filter" />
        <MenuButton onClick={() => attempt(() => { const rule = typedCondition(date); if (rule.kind !== 'date') throw new RangeError('Enter a calendar date.')
          const day = (d: { year: number; month: number; day: number }) => serialOf(`${String(d.year).padStart(4, '0')}-${String(d.month).padStart(2, '0')}-${String(d.day).padStart(2, '0')}`)
          return { kind: 'filter', field: field.index, filter: dateFilter(field.index, id, rule.op as DateOp, day(rule.first), rule.second ? day(rule.second) : undefined) } })}>Apply date filter</MenuButton>
      </div> : <div>
        <ConditionEditor value={label} onChange={setLabel} classification={TEXT} kinds={['text']} index={1} invalid={false} label="Label filter" />
        <MenuButton onClick={() => attempt(() => { const rule = typedCondition(label); if (rule.kind !== 'text') throw new RangeError('Enter a label.')
          if (!rule.first) throw new RangeError('Enter a label to compare with.')
          return { kind: 'filter', field: field.index, filter: labelFilter(field.index, id, rule.op as LabelOp, rule.first, undefined, rule.wildcards) } })}>Apply label filter</MenuButton>
      </div>}
      {model.data.length > 0 && <div>
        {dataSelect(valueData, setValueData, 'Value filter on')}
        <ConditionEditor value={value} onChange={setValue} classification={NUMBER} kinds={['number']} index={2} invalid={false} label="Value filter" />
        <MenuButton onClick={() => attempt(() => { const rule = typedCondition(value); if (rule.kind !== 'number') throw new RangeError('Enter a number.')
          return { kind: 'filter', field: field.index, filter: valueFilter(field.index, id, valueData, rule.op as ValueOp, rule.first, rule.second) } })}>Apply value filter</MenuButton>
      </div>}
      {model.data.length > 0 && <fieldset><legend>Top 10</legend>
        {/* Bottom refuses until measured, so Excel's Top/Bottom choice is shown as text: Top only. */}
        <p>Show the top items</p>
        <label>Count <input aria-label="Top 10 count" inputMode="decimal" className={CONTROL_CLS} style={controlTarget} value={top.count} onChange={event => setTop({ ...top, count: event.target.value })} /></label>
        <label>Of <select aria-label="Top 10 kind" className={SELECT_CLS} style={controlTarget} value={top.kind} onChange={event => setTop({ ...top, kind: event.target.value as typeof top.kind })}>
          <option value="count">Items</option><option value="percent">Percent</option><option value="sum">Sum</option></select></label>
        {dataSelect(top.data, data => setTop({ ...top, data }), 'Top 10 by')}
        <MenuButton onClick={() => attempt(() => { const count = Number(top.count); if (!top.count.trim() || !Number.isFinite(count) || count <= 0) throw new RangeError('Enter a positive number.')
          return { kind: 'filter', field: field.index, filter: top10Filter(field.index, id, top.data, top.kind, count) } })}>Apply top 10</MenuButton>
      </fieldset>}
    </>}
    {has('clearFilters') && <MenuButton onClick={() => onApply({ kind: 'clearFilters', field: field.index })}>{`Clear filter from ${field.name}`}</MenuButton>}
    {has('items') && <ItemChecklist key={field.index} field={field} onApply={hidden => onApply({ kind: 'items', field: field.index, hidden })} />}
    {error && <p role="alert">{error}</p>}
    <div><MenuButton onClick={() => onApply('reset')}>Reset to saved view</MenuButton><MenuButton onClick={props.onClose}>Cancel</MenuButton></div>
  </PivotMenuShell>
}
