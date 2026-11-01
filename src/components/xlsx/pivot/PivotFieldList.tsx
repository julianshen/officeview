/** Lazy Stage 4 field list. It keeps the sheet grid's single Tab stop and offers every move through buttons as well as drag. */
import { useEffect, useMemo, useRef, useState, type DragEvent, type KeyboardEvent, type ReactElement } from 'react'
import { useControlMedia } from '../controlMedia'
import { controlTarget, CONTROL_CLS } from '../controlUi'
import type { PivotDelta, PivotMenuModel } from './pivotClient'
import type { PivotDeltaKind } from './pivotClient'
import { PivotGroupDialog } from './PivotGroupDialog'
import { PivotValueSettingsDialog } from './PivotValueSettingsDialog'
import { buildPivotFieldListIndex } from './fieldListIndex'

export interface PivotFieldListProps {
  model: PivotMenuModel
  pivot: number
  pivotCount?: number
  onPivot?(pivot: number): void
  onApply(delta: PivotDelta | 'reset'): Promise<boolean> | void
  verified?: ReadonlySet<PivotDeltaKind>
  onClose(): void
  invoker?: HTMLElement | null
}

type Area = 'page' | 'col' | 'row' | 'data'
type DragSource = { kind: 'field'; field: number } | { kind: 'chip'; field: number; area: Area; index: number }
const areas: Array<{ id: Area; label: string }> = [
  { id: 'page', label: 'Filters' }, { id: 'col', label: 'Columns' }, { id: 'row', label: 'Rows' }, { id: 'data', label: 'Values' },
]
const FIELD_LIST_RENDER_LIMIT = 200

export function PivotFieldList({ model, pivot, pivotCount = 1, onPivot, onApply, onClose, invoker, verified }: PivotFieldListProps): ReactElement {
  const mobile = useControlMedia('(pointer: coarse), (max-width: 640px)')
  const root = useRef<HTMLElement>(null)
  const [search, setSearch] = useState('')
  const [dragging, setDragging] = useState<DragSource | null>(null)
  const [groupField, setGroupField] = useState<number | null>(null)
  const [valueSettingsDataIndex, setValueSettingsDataIndex] = useState<number | null>(null)
  const fieldIndex = useMemo(() => buildPivotFieldListIndex(model), [model.fields, model.data])
  const orderedFields = useMemo(() => {
    return [...model.fields].sort(model.fieldListSortAscending ? (a, b) => a.name.localeCompare(b.name) || a.sourceOrder - b.sourceOrder : (a, b) => a.sourceOrder - b.sourceOrder)
  }, [model.fields, model.fieldListSortAscending])
  const fields = useMemo(() => {
    const query = search.trim().toLocaleLowerCase()
    return orderedFields.filter(field => !query || field.name.toLocaleLowerCase().includes(query))
  }, [orderedFields, search])
  const shownFields = fields.slice(0, FIELD_LIST_RENDER_LIMIT)
  const dataFor = (field: number) => fieldIndex.dataIndicesByField.get(field) ?? []
  const isPlaced = (field: number) => fieldIndex.placedFields.has(field)
  const move = (field: number, area: Area, index?: number) => {
    if (field === -2) {
      if (area === 'row' || area === 'col') onApply({ kind: 'moveField', field, axis: area, ...(index === undefined ? {} : { index }) })
      return
    }
    const entry = fieldIndex.fieldsByIndex.get(field)
    if (!entry) return
    if (area === 'data') {
      if (entry.permissions.dragToData) onApply({ kind: 'addDataField', field, ...(index === undefined ? {} : { index }) })
      return
    }
    const axis = area === 'page' ? 'page' : area === 'col' ? 'col' : 'row'
    const allowed = axis === 'page' ? entry.permissions.dragToPage : axis === 'col' ? entry.permissions.dragToCol : entry.permissions.dragToRow
    if (allowed) onApply({ kind: 'moveField', field, axis, ...(index === undefined ? {} : { index }) })
  }
  const removeEverywhere = (field: number) => {
    const entry = fieldIndex.fieldsByIndex.get(field)
    if (!entry?.permissions.dragOff) return
    for (const index of [...dataFor(field)].reverse()) onApply({ kind: 'removeDataField', dataIndex: index })
    if (entry.axis !== 'none') onApply({ kind: 'moveField', field, axis: 'none' })
  }
  const chips = (area: Area) => area === 'page' ? model.pages : area === 'row' ? model.rows : area === 'col' ? model.cols : model.data.map((_, index) => index)
  const fieldOf = (area: Area, chip: number) => area === 'data' ? model.data[chip]?.field : chip
  const drop = (event: DragEvent, area: Area, index: number) => {
    event.preventDefault()
    event.stopPropagation()
    const source = dragging ?? { kind: 'field' as const, field: Number(event.dataTransfer.getData('text/plain')) }
    if (source.kind === 'chip' && source.area === 'data' && area === 'data') {
      const toIndex = Math.max(0, Math.min(model.data.length - 1, index))
      onApply({ kind: 'moveDataField', dataIndex: source.index, toIndex })
    } else if (Number.isInteger(source.field) && (source.field >= 0 || source.field === -2)) {
      const targetLength = chips(area).length
      const targetIndex = source.kind === 'chip' && source.area === area
        ? Math.max(0, Math.min(targetLength - 1, index)) : index
      move(source.field, area, targetIndex)
    }
    setDragging(null)
  }
  const close = () => { onClose(); requestAnimationFrame(() => invoker?.focus()) }
  useEffect(() => { root.current?.querySelector<HTMLElement>('input[aria-label="Search fields"]')?.focus() }, [])
  const keyDown = (event: KeyboardEvent<HTMLElement>) => {
    event.stopPropagation()
    if (event.key === 'Escape') { event.preventDefault(); close(); return }
    if (mobile && event.key === 'Tab') {
      const nodes = [...(root.current?.querySelectorAll<HTMLElement>('button:not(:disabled),input:not(:disabled),select:not(:disabled)') ?? [])]
      if (!nodes.length) return
      const first = nodes[0]!, last = nodes[nodes.length - 1]!
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus() }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
    }
  }
  return <>
    {mobile && <div aria-hidden="true" data-officeview-ui="pivot-field-list-backdrop" onPointerDown={event => { if (event.target === event.currentTarget) close() }} style={{ position: 'fixed', inset: 0, zIndex: 9, background: '#0005' }} />}
    <section ref={root} role={mobile ? 'dialog' : 'region'} aria-label="PivotTable Fields" aria-modal={mobile || undefined}
    data-pivot-field-list={mobile ? 'bottom-sheet' : 'docked'} data-officeview-ui="pivot-field-list"
    onKeyDown={keyDown}
    style={{ ...(mobile ? { position: 'fixed', insetInline: 0, bottom: 0, zIndex: 10, maxHeight: '85dvh', overflowY: 'auto', borderRadius: '12px 12px 0 0', boxShadow: '0 -4px 24px #0003' } : { width: 300, flex: '0 0 300px', overflowY: 'auto', borderInlineStart: '1px solid var(--ov-color-border)' }), background: 'var(--ov-color-surface, #fff)', color: 'var(--ov-color-text, #222)', boxSizing: 'border-box', padding: 12 }}>
    <header style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
      <h2 style={{ margin: 0, fontSize: 16 }}>PivotTable Fields</h2>
      {pivotCount > 1 && <label>Pivot <select aria-label="Choose pivot" value={pivot} onChange={event => onPivot?.(Number(event.target.value))}>{Array.from({ length: pivotCount }, (_, i) => <option key={i} value={i}>{i + 1}</option>)}</select></label>}
      <button type="button" style={controlTarget} aria-label="Close field list" onClick={close}>Close</button>
    </header>
    <label style={{ display: 'block', marginBlock: 10 }}>Search fields <input aria-label="Search fields" className={CONTROL_CLS} style={{ ...controlTarget, width: '100%' }} value={search} onChange={event => setSearch(event.target.value)} /></label>
    <fieldset><legend>Choose fields to add to report</legend>
      {shownFields.map(field => <div key={field.index} draggable={field.permissions.dragOff || field.permissions.dragToData || field.permissions.dragToRow || field.permissions.dragToCol || field.permissions.dragToPage}
        onDragStart={event => { setDragging({ kind: 'field', field: field.index }); event.dataTransfer.setData('text/plain', String(field.index)) }} onDragEnd={() => setDragging(null)}
        style={{ display: 'flex', alignItems: 'center', gap: 4, minHeight: 44 }}>
        <label style={{ ...controlTarget, display: 'flex', alignItems: 'center', gap: 6, flex: 1, minWidth: 0 }}><input type="checkbox" aria-label={field.name} checked={isPlaced(field.index)} disabled={isPlaced(field.index) ? !field.permissions.dragOff : !field.permissions.dragToData && !field.permissions.dragToRow}
          onChange={event => { if (event.target.checked) field.defaultDataFn === 'sum' && field.permissions.dragToData ? move(field.index, 'data') : move(field.index, 'row'); else removeEverywhere(field.index) }} />{field.name}</label>
        <select aria-label={`Move ${field.name}`} style={{ ...controlTarget, maxWidth: 130 }} value="" onChange={event => { const value = event.target.value as Area | ''; if (value) move(field.index, value); event.currentTarget.value = '' }}>
          <option value="">Move to…</option>
          {field.permissions.dragToPage && !field.calculated && <option value="page">Move to Filters</option>}{field.permissions.dragToCol && !field.calculated && <option value="col">Move to Columns</option>}{field.permissions.dragToRow && !field.calculated && <option value="row">Move to Rows</option>}{field.permissions.dragToData && <option value="data">Move to Values</option>}
        </select>
      </div>)}
      {fields.length > shownFields.length && <p role="status">Showing {shownFields.length} of {fields.length} fields. Search to narrow the list.</p>}
    </fieldset>
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, marginBlock: 12 }}>
      {areas.map(area => <fieldset key={area.id} aria-label={area.label} onDragOver={event => event.preventDefault()} onDrop={event => drop(event, area.id, chips(area.id).length)} style={{ minWidth: 0, minHeight: 84 }}>
        <legend>{area.label}</legend>
        {chips(area.id).map((chip, index) => {
          const field = fieldOf(area.id, chip), meta = field === undefined ? undefined : fieldIndex.fieldsByIndex.get(field), label = area.id === 'data' ? model.data[chip]?.name : meta?.name
          if (field === undefined) return null
          if (field === -2) {
            const axis = area.id === 'row' ? 'row' : 'col'
            return <div key={`${field}:${index}`} draggable onDragStart={event => { setDragging({ kind: 'chip', field, area: area.id, index }); event.dataTransfer.setData('text/plain', String(field)) }} onDragEnd={() => setDragging(null)} onDragOver={event => event.preventDefault()} onDrop={event => drop(event, area.id, index)}><span>{model.valuesCaption}</span><select aria-label={`Actions for ${model.valuesCaption}`} style={controlTarget} value="" onChange={event => { const action = event.target.value; if (action === 'up' || action === 'down' || action === 'begin' || action === 'end') {
              const list = chips(area.id), target = action === 'up' ? Math.max(0, index - 1) : action === 'down' ? Math.min(list.length - 1, index + 1) : action === 'begin' ? 0 : list.length - 1; move(-2, area.id, target)
            } else if (action === 'other') move(-2, area.id === 'row' ? 'col' : 'row'); event.currentTarget.value = '' }}>
              <option value="">Actions</option><option value="up">Move Up</option><option value="down">Move Down</option><option value="begin">Move to Beginning</option><option value="end">Move to End</option><option value="other">Move to {axis === 'row' ? 'Columns' : 'Rows'}</option>
            </select></div>
          }
          if (!meta) return null
          return <div key={`${field}:${index}`} draggable onDragStart={event => { setDragging({ kind: 'chip', field, area: area.id, index }); event.dataTransfer.setData('text/plain', String(field)) }} onDragEnd={() => setDragging(null)} onDragOver={event => event.preventDefault()} onDrop={event => drop(event, area.id, index)} style={{ display: 'flex', alignItems: 'center', gap: 2 }}>
            <span>{label}</span>
            <select aria-label={`Actions for ${label}`} style={controlTarget} value="" onChange={event => { const action = event.target.value; if (action === 'up' || action === 'down' || action === 'begin' || action === 'end') {
              if (area.id === 'data') { const toIndex = action === 'up' ? Math.max(0, index - 1) : action === 'down' ? Math.min(model.data.length - 1, index + 1) : action === 'begin' ? 0 : model.data.length - 1; onApply({ kind: 'moveDataField', dataIndex: index, toIndex }) }
              else { const list = chips(area.id), nextIndex = action === 'up' ? Math.max(0, index - 1) : action === 'down' ? Math.min(list.length - 1, index + 1) : action === 'begin' ? 0 : list.length - 1; move(field, area.id, nextIndex) }
            } else if (action === 'remove') area.id === 'data' ? onApply({ kind: 'removeDataField', dataIndex: index }) : onApply({ kind: 'moveField', field, axis: 'none' }); else if (action === 'valueSettings') setValueSettingsDataIndex(index); else if (action === 'group') setGroupField(field); else if (action === 'ungroup') onApply({ kind: 'ungroup', field }); else if (action) move(field, action as Area) ; event.currentTarget.value = '' }}>
              <option value="">Actions</option>
              {area.id === 'data' && (!verified || verified.has('valueSettings')) && <option value="valueSettings">Value field settings…</option>}
              {(area.id === 'row' || area.id === 'col') && meta.group?.kind === 'date' && verified?.has('groupDate') && <option value="group">Group…</option>}{(area.id === 'row' || area.id === 'col') && meta.group?.kind === 'number' && verified?.has('groupNumber') && <option value="group">Group…</option>}{(area.id === 'row' || area.id === 'col') && meta.group?.grouped && verified?.has('ungroup') && <option value="ungroup">Ungroup</option>}
              {meta.permissions.dragToPage && !meta.calculated && <option value="page">Move to Filters</option>}{meta.permissions.dragToCol && !meta.calculated && <option value="col">Move to Columns</option>}{meta.permissions.dragToRow && !meta.calculated && <option value="row">Move to Rows</option>}{meta.permissions.dragToData && <option value="data">Move to Values</option>}
              <option value="up">Move Up</option><option value="down">Move Down</option><option value="begin">Move to Beginning</option><option value="end">Move to End</option>{meta.permissions.dragOff && <option value="remove">Remove Field</option>}
            </select>
          </div>
        })}
      </fieldset>)}
    </div>
    <footer style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}><button type="button" style={controlTarget} onClick={() => onApply('reset')}>Reset to saved view</button><button type="button" style={controlTarget} onClick={close}>Close</button></footer>
    </section>
    {groupField !== null && fieldIndex.fieldsByIndex.get(groupField)?.group?.kind && <PivotGroupDialog field={fieldIndex.fieldsByIndex.get(groupField)!} date1904={model.date1904 ?? false}
      onApply={delta => Promise.resolve(onApply(delta)).then(result => result !== false)} onClose={() => setGroupField(null)} />}
    {valueSettingsDataIndex !== null && model.data[valueSettingsDataIndex] && <PivotValueSettingsDialog model={model} dataIndex={valueSettingsDataIndex}
      onApply={delta => Promise.resolve(onApply(delta)).then(result => result !== false)} onClose={() => setValueSettingsDataIndex(null)} />}
  </>
}
