/** Accessible, bounded Group dialog for Task 37 (lazy pivots chunk). */
import { useEffect, useRef, useState, type ReactElement, type KeyboardEvent } from 'react'
import type { PivotDelta } from './pivotClient'
import type { PivotDateGroupPeriod } from '../../../xlsx/pivots/session/delta'
import type { PivotMenuField } from './pivotClient'
import { controlTarget, CONTROL_CLS } from '../controlUi'
import { decodeSerial, encodeCivil } from '../../../xlsx/formula/serial'

export interface PivotGroupDialogProps {
  field: PivotMenuField
  date1904: boolean
  onApply: (delta: Extract<PivotDelta, { kind: 'groupDate' | 'groupNumber' }>) => Promise<boolean>
  onClose: () => void
}
const PERIODS: Array<[PivotDateGroupPeriod, string]> = [['seconds', 'Seconds'], ['minutes', 'Minutes'], ['hours', 'Hours'], ['days', 'Days'], ['months', 'Months'], ['quarters', 'Quarters'], ['years', 'Years']]
const dateInput = (serial: number | undefined, date1904: boolean): string => {
  if (serial === undefined || !Number.isFinite(serial)) return ''
  const decoded = decodeSerial(serial, date1904 ? '1904' : '1900')
  return 'civil' in decoded ? `${String(decoded.civil.year).padStart(4, '0')}-${String(decoded.civil.month).padStart(2, '0')}-${String(decoded.civil.day).padStart(2, '0')}` : ''
}
const dateSerial = (text: string, date1904: boolean): number => {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text)
  if (!match) throw new RangeError('Enter a date in YYYY-MM-DD format.')
  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])))
  if (date.toISOString().slice(0, 10) !== text) throw new RangeError('Enter a valid calendar date.')
  const serial = encodeCivil({ year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) }, date1904 ? '1904' : '1900')
  if (typeof serial !== 'number') throw new RangeError('Enter a valid calendar date.')
  return serial
}
export function PivotGroupDialog({ field, date1904, onApply, onClose }: PivotGroupDialogProps): ReactElement {
  const root = useRef<HTMLDivElement>(null), opener = useRef<HTMLElement | null>(null)
  const info = field.group!
  const isDate = info.kind === 'date'
  const [periods, setPeriods] = useState<PivotDateGroupPeriod[]>(() => [...(info.date?.periods ?? [])])
  const [autoStart, setAutoStart] = useState(info.date?.autoStart ?? info.number?.autoStart ?? true)
  const [autoEnd, setAutoEnd] = useState(info.date?.autoEnd ?? info.number?.autoEnd ?? true)
  const [startDate, setStartDate] = useState(dateInput(info.date?.startDate ?? info.bounds?.min, date1904))
  const [endDate, setEndDate] = useState(dateInput(info.date?.endDate ?? info.bounds?.max, date1904))
  const [startNumber, setStartNumber] = useState(String(info.number?.start ?? info.bounds?.min ?? ''))
  const [endNumber, setEndNumber] = useState(String(info.number?.end ?? info.bounds?.max ?? ''))
  const [by, setBy] = useState(String(info.number?.by ?? 1)), [days, setDays] = useState(String(info.date?.daysInterval ?? 1))
  const [error, setError] = useState(''), [busy, setBusy] = useState(false)
  useEffect(() => {
    opener.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
    root.current?.querySelector<HTMLElement>('input,button')?.focus()
    return () => opener.current?.focus()
  }, [])
  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    event.stopPropagation()
    if (event.key === 'Escape') { event.preventDefault(); onClose(); return }
    if (event.key === 'Tab') {
      const nodes = [...(root.current?.querySelectorAll<HTMLElement>('button:not(:disabled),input:not(:disabled)') ?? [])]
      if (!nodes.length) return
      const first = nodes[0]!, last = nodes[nodes.length - 1]!
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus() }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
    }
  }
  const apply = async () => {
    try {
      setError(''); setBusy(true)
      if (isDate) {
        const selected = PERIODS.filter(([period]) => periods.includes(period)).map(([period]) => period)
        if (!selected.length) throw new RangeError('Choose at least one date period.')
        if (selected.includes('days') && selected.length > 1 && Number(days) !== 1) throw new RangeError('Number of days must be 1 when Days is combined with another period.')
        const daysInterval = selected.includes('days') ? Number(days) : 1
        const start = autoStart ? undefined : dateSerial(startDate, date1904), end = autoEnd ? undefined : dateSerial(endDate, date1904)
        if (start !== undefined && end !== undefined && start > end) throw new RangeError('Starting at must not be after Ending at.')
        if (selected.includes('days') && (!Number.isInteger(daysInterval) || daysInterval < 1 || daysInterval > 32767)) throw new RangeError('Number of days must be between 1 and 32767.')
        const result = await onApply({ kind: 'groupDate', field: info.baseField, periods: selected, ...(start !== undefined ? { startDate: start } : {}), ...(end !== undefined ? { endDate: end } : {}), daysInterval })
        if (result) onClose()
        else setError('Grouping could not be applied.')
      } else {
        const start = autoStart ? undefined : Number(startNumber), end = autoEnd ? undefined : Number(endNumber), interval = Number(by)
        if (start !== undefined && !Number.isFinite(start) || end !== undefined && !Number.isFinite(end)) throw new RangeError('Enter finite numeric bounds.')
        if (start !== undefined && end !== undefined && start > end) throw new RangeError('Starting at must not be greater than Ending at.')
        if (!Number.isFinite(interval) || interval <= 0) throw new RangeError('By must be greater than zero.')
        const result = await onApply({ kind: 'groupNumber', field: info.baseField, ...(start !== undefined ? { start } : {}), ...(end !== undefined ? { end } : {}), by: interval })
        if (result) onClose()
        else setError('Grouping could not be applied.')
      }
    } catch (cause) { setError(cause instanceof Error ? cause.message : 'Check the grouping values.') }
    finally { setBusy(false) }
  }
  return <div ref={root} role="dialog" aria-modal="true" aria-labelledby="pivot-group-title" data-officeview-ui="pivot-group-dialog" onKeyDown={onKeyDown}
    onPointerDown={event => event.stopPropagation()} style={{ position: 'fixed', inset: 0, zIndex: 20, display: 'grid', placeItems: 'center', background: '#0005' }}>
    <section style={{ background: 'var(--ov-color-surface, white)', color: 'var(--ov-color-text, #222)', padding: 16, width: 'min(420px, calc(100vw - 24px))', maxHeight: '85dvh', overflowY: 'auto' }}>
      <h2 id="pivot-group-title">Group {field.name}</h2>
      <fieldset><legend>Starting at</legend><label><input type="checkbox" checked={autoStart} onChange={event => setAutoStart(event.target.checked)} /> Auto</label><input aria-label="Starting at" type={isDate ? 'date' : 'number'} className={CONTROL_CLS} style={controlTarget} disabled={autoStart} value={isDate ? startDate : startNumber} onChange={event => isDate ? setStartDate(event.target.value) : setStartNumber(event.target.value)} /></fieldset>
      <fieldset><legend>Ending at</legend><label><input type="checkbox" checked={autoEnd} onChange={event => setAutoEnd(event.target.checked)} /> Auto</label><input aria-label="Ending at" type={isDate ? 'date' : 'number'} className={CONTROL_CLS} style={controlTarget} disabled={autoEnd} value={isDate ? endDate : endNumber} onChange={event => isDate ? setEndDate(event.target.value) : setEndNumber(event.target.value)} /></fieldset>
      {isDate ? <fieldset><legend>By</legend>{PERIODS.map(([period, title]) => <label key={period} style={{ display: 'block' }}><input type="checkbox" checked={periods.includes(period)} onChange={event => setPeriods(old => event.target.checked ? [...old, period] : old.filter(item => item !== period))} /> {title}</label>)}
        <label>Number of days <input aria-label="Number of days" type="number" min="1" max="32767" disabled={periods.length !== 1 || periods[0] !== 'days'} className={CONTROL_CLS} style={controlTarget} value={days} onChange={event => setDays(event.target.value)} /></label></fieldset>
        : <label>By <input aria-label="By" type="number" min="0.000000000001" className={CONTROL_CLS} style={controlTarget} value={by} onChange={event => setBy(event.target.value)} /></label>}
      {error && <p role="alert">{error}</p>}
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}><button type="button" style={controlTarget} disabled={busy} onClick={() => void apply()}>OK</button><button type="button" style={controlTarget} onClick={onClose}>Cancel</button></div>
      <span hidden>{date1904 ? '1904 date system' : '1900 date system'}</span>
    </section>
  </div>
}
