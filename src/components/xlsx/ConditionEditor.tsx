import type { ReactElement } from 'react'
import type { CalendarDay, DateOp, DatePreset, NumberOp, TextOp, TypedCondition } from '../../xlsx/view/contracts'
import type { ColumnClassification } from '../../xlsx/view/filter'
import { controlTarget } from './controlUi'

export interface ConditionInput { kind: TypedCondition['kind']; op: string; first: string; second: string; wildcards: boolean; preset: DatePreset }
export function conditionInput(rule?: TypedCondition, kind: TypedCondition['kind'] = 'text'): ConditionInput {
  const day = (value: CalendarDay) => `${String(value.year).padStart(4, '0')}-${String(value.month).padStart(2, '0')}-${String(value.day).padStart(2, '0')}`
  return { kind: rule?.kind ?? kind, op: rule && 'op' in rule ? rule.op : 'eq', first: rule && 'first' in rule ? typeof rule.first === 'object' ? day(rule.first) : String(rule.first) : '',
    second: rule && 'second' in rule && rule.second !== undefined ? typeof rule.second === 'object' ? day(rule.second) : String(rule.second) : '', wildcards: rule?.kind === 'text' && rule.wildcards, preset: rule?.kind === 'datePreset' ? rule.preset : 'today' }
}
function calendar(text: string): CalendarDay {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text)
  if (!match) throw new RangeError('Enter a calendar date as YYYY-MM-DD.')
  return { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) }
}
function number(text: string): number { if (!text.trim() || !Number.isFinite(Number(text))) throw new RangeError('Enter a finite number.'); return Number(text) }
export function typedCondition(input: ConditionInput): TypedCondition {
  switch (input.kind) {
    case 'datePreset': return { kind: 'datePreset', preset: input.preset }
    case 'text': return { kind: 'text', op: input.op as TextOp, first: input.first, wildcards: input.wildcards }
    case 'number': return { kind: 'number', op: input.op as NumberOp, first: number(input.first), ...(input.op === 'between' ? { second: number(input.second) } : {}) }
    case 'date': return { kind: 'date', op: input.op as DateOp, first: calendar(input.first), ...(input.op === 'between' ? { second: calendar(input.second) } : {}) }
  }
}
const operators = { number: ['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'between'], text: ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'], date: ['eq', 'neq', 'before', 'after', 'between'], datePreset: [] }
const names: Record<string, string> = { eq: 'Equals', neq: 'Does not equal', gt: 'Greater than', gte: 'Greater than or equal', lt: 'Less than', lte: 'Less than or equal', between: 'Between (inclusive)', begins: 'Begins with', ends: 'Ends with', contains: 'Contains', 'not-contains': 'Does not contain', before: 'Before', after: 'After' }
export function ConditionEditor({ value, onChange, classification, index, invalid }: { value: ConditionInput; onChange: (value: ConditionInput) => void; classification: ColumnClassification; index: number; invalid: boolean }): ReactElement {
  const types: TypedCondition['kind'][] = classification.kind === 'date' ? ['date', 'datePreset', 'number'] : classification.kind === 'number' ? ['number'] : classification.hasNumeric ? ['text', 'number'] : ['text']
  const prefix = `Condition ${index}`
  return <fieldset><legend>{prefix}</legend>
    <label>Type <select aria-label={`${prefix} type`} value={value.kind} style={controlTarget} onChange={event => onChange(conditionInput(undefined, event.target.value as TypedCondition['kind']))}>
      {types.map(kind => <option key={kind} value={kind}>{kind === 'datePreset' ? 'Date preset' : kind === 'date' ? 'Calendar date' : kind === 'number' ? 'Number' : 'Text'}</option>)}</select></label>
    {value.kind === 'datePreset' ? <label>Period <select aria-label={`${prefix} preset`} value={value.preset} style={controlTarget} onChange={event => onChange({ ...value, preset: event.target.value as DatePreset })}>
      {(['today', 'yesterday', 'tomorrow', 'thisWeek', 'thisMonth', 'thisYear'] as const).map(preset => <option key={preset} value={preset}>{preset}</option>)}</select></label> : <>
      <label>Operator <select aria-label={`${prefix} operator`} value={value.op} style={controlTarget} onChange={event => onChange({ ...value, op: event.target.value })}>{operators[value.kind].map(op => <option key={op} value={op}>{names[op]}</option>)}</select></label>
      <label>{value.kind === 'date' ? 'Calendar date' : 'Value'} <input aria-label={`${prefix} first`} aria-invalid={invalid} value={value.first} placeholder={value.kind === 'date' ? 'YYYY-MM-DD' : undefined} style={controlTarget} onChange={event => onChange({ ...value, first: event.target.value })} /></label>
      {value.op === 'between' && <label>Through <input aria-label={`${prefix} second`} aria-invalid={invalid} value={value.second} style={controlTarget} onChange={event => onChange({ ...value, second: event.target.value })} /></label>}
      {value.kind === 'text' && <label style={{ ...controlTarget, display: 'inline-flex', alignItems: 'center' }}><input type="checkbox" checked={value.wildcards} onChange={event => onChange({ ...value, wildcards: event.target.checked })} />Use * / ? wildcards (~ escapes)</label>}
    </>}
  </fieldset>
}
