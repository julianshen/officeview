/** Accessible, bounded Value Field Settings dialog (Task 36, spec §6.7, Ruling S4-H, S4-I). */
import { useEffect, useMemo, useRef, useState, type CSSProperties, type KeyboardEvent, type ReactElement } from 'react'
import type { PivotDelta, PivotMenuModel } from './pivotClient'
import type { PivotDataFunction, PivotValueShowAs } from '../../../xlsx/pivots/session/delta'
import { controlTarget, CONTROL_CLS, SELECT_CLS, CONTROL_PRIMARY_BTN_CLS, CONTROL_BTN_CLS } from '../controlUi'

// Built-in number-format presets offered in Value Field Settings (Ruling S4-I):
export const PIVOT_VALUE_NUMFMT_PRESETS: ReadonlyArray<{ readonly numFmtId: number; readonly label: string; readonly sample: string }> = [
  { numFmtId: 0, label: 'General', sample: '1234.5' },
  { numFmtId: 1, label: '0', sample: '1235' },
  { numFmtId: 2, label: '0.00', sample: '1234.50' },
  { numFmtId: 3, label: '#,##0', sample: '1,235' },
  { numFmtId: 4, label: '#,##0.00', sample: '1,234.50' },
  { numFmtId: 9, label: '0%', sample: '12%' },
  { numFmtId: 10, label: '0.00%', sample: '12.35%' },
  { numFmtId: 11, label: '0.00E+00', sample: '1.23E+03' },
  { numFmtId: 14, label: 'yyyy-mm-dd', sample: '2024-03-15' },
  { numFmtId: 37, label: '#,##0;(#,##0)', sample: '(1,235)' },
  { numFmtId: 38, label: '#,##0;[Red](#,##0)', sample: '(1,235)' },
  { numFmtId: 39, label: '#,##0.00;(#,##0.00)', sample: '(1,234.50)' },
  { numFmtId: 40, label: '#,##0.00;[Red](#,##0.00)', sample: '(1,234.50)' },
  { numFmtId: 44, label: 'Accounting', sample: '$ 1,234.50' },
  { numFmtId: 49, label: 'Text (@)', sample: '1234.5' },
]

export interface PivotValueSettingsDialogProps {
  model: PivotMenuModel
  dataIndex: number
  onApply: (delta: Extract<PivotDelta, { kind: 'valueSettings' }>) => Promise<boolean>
  onClose: () => void
}

const FUNCTIONS: Array<{ fn: PivotDataFunction; label: string }> = [
  { fn: 'sum', label: 'Sum' },
  { fn: 'count', label: 'Count' },
  { fn: 'average', label: 'Average' },
  { fn: 'max', label: 'Max' },
  { fn: 'min', label: 'Min' },
  { fn: 'product', label: 'Product' },
  { fn: 'countNums', label: 'Count Numbers' },
  { fn: 'stdDev', label: 'StdDev' },
  { fn: 'stdDevp', label: 'StdDevp' },
  { fn: 'var', label: 'Var' },
  { fn: 'varp', label: 'Varp' },
]

const SHOW_AS_MODES: Array<{ mode: PivotValueShowAs; label: string }> = [
  { mode: 'normal', label: 'No Calculation' },
  { mode: 'difference', label: 'Difference From' },
  { mode: 'percent', label: '% Of' },
  { mode: 'percentDiff', label: '% Difference From' },
  { mode: 'runTotal', label: 'Running Total In' },
  { mode: 'percentOfRow', label: '% of Row Total' },
  { mode: 'percentOfCol', label: '% of Column Total' },
  { mode: 'percentOfTotal', label: '% of Grand Total' },
  { mode: 'index', label: 'Index' },
  { mode: 'percentOfParent', label: '% of Parent Total' },
  { mode: 'percentOfParentRow', label: '% of Parent Row Total' },
  { mode: 'percentOfParentCol', label: '% of Parent Column Total' },
  { mode: 'percentOfRunningTotal', label: '% Running Total In' },
  { mode: 'rankAscending', label: 'Rank Smallest to Largest' },
  { mode: 'rankDescending', label: 'Rank Largest to Smallest' },
]

export const requiresBaseField = (mode: PivotValueShowAs): boolean =>
  ['percentOfParent', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRunningTotal', 'rankAscending', 'rankDescending'].includes(mode)

export const requiresBaseItem = (mode: PivotValueShowAs): boolean =>
  ['difference', 'percent', 'percentDiff'].includes(mode)

export function PivotValueSettingsDialog({ model, dataIndex, onApply, onClose }: PivotValueSettingsDialogProps): ReactElement {
  const root = useRef<HTMLDivElement>(null)
  const opener = useRef<HTMLElement | null>(null)

  const currentData = model.data[dataIndex]
  const sourceFieldMeta = model.fields.find(f => f.index === currentData?.field)
  const isCalculated = Boolean(sourceFieldMeta?.calculated)

  const [activeTab, setActiveTab] = useState<'summary' | 'showAs'>('summary')
  const [customName, setCustomName] = useState(currentData?.name ?? '')
  const [fn, setFn] = useState<PivotDataFunction>(isCalculated ? 'sum' : (currentData?.fn ?? 'sum'))
  const [showAs, setShowAs] = useState<PivotValueShowAs>(currentData?.showAs ?? 'normal')

  const availableBaseFields = useMemo(() => {
    if (showAs === 'percentOfParent') {
      return model.fields.filter(f => model.rows.includes(f.index) && f.index !== -2)
    }
    return model.fields.filter(f => (model.rows.includes(f.index) || model.cols.includes(f.index)) && f.index !== -2)
  }, [model.fields, model.rows, model.cols, showAs])

  const initialBaseField = currentData?.baseField !== undefined && currentData.baseField !== -1
    ? currentData.baseField
    : (availableBaseFields[0]?.index ?? 0)
  const [baseField, setBaseField] = useState<number>(initialBaseField)

  const chosenBaseFieldMeta = model.fields.find(f => f.index === baseField)
  const [baseItem, setBaseItem] = useState<number>(currentData?.baseItem ?? 1048828)

  const [numFmtId, setNumFmtId] = useState<string>(
    currentData?.numFmtId !== undefined ? String(currentData.numFmtId) : ''
  )
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)

  useEffect(() => {
    opener.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
    root.current?.querySelector<HTMLElement>('input,button,select')?.focus()
    return () => opener.current?.focus()
  }, [])

  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    event.stopPropagation()
    if (event.key === 'Escape') {
      event.preventDefault()
      onClose()
      return
    }
    if (event.key === 'Tab') {
      const nodes = [...(root.current?.querySelectorAll<HTMLElement>('button:not(:disabled),input:not(:disabled),select:not(:disabled)') ?? [])]
      if (!nodes.length) return
      const first = nodes[0]!
      const last = nodes[nodes.length - 1]!
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault()
        last.focus()
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault()
        first.focus()
      }
    }
  }

  const apply = async () => {
    try {
      setError('')
      const trimmed = customName.trim()
      if (!trimmed) {
        setError('The custom name cannot be blank.')
        return
      }
      const otherNames = model.data.filter((_, i) => i !== dataIndex).map(d => d.name).filter(Boolean)
      if (otherNames.some(n => n.toLowerCase() === trimmed.toLowerCase())) {
        setError(`A data field named "${trimmed}" already exists.`)
        return
      }

      setBusy(true)
      const delta: Extract<PivotDelta, { kind: 'valueSettings' }> = {
        kind: 'valueSettings',
        dataIndex,
        fn,
        showAs,
        name: trimmed,
        numFmtId: numFmtId === '' ? null : Number(numFmtId),
        ...(requiresBaseField(showAs) ? { baseField } : {}),
        ...(requiresBaseItem(showAs) ? { baseItem } : {}),
      }

      const success = await onApply(delta)
      if (success) {
        onClose()
      } else {
        setError('The value field settings could not be applied.')
      }
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Check the value field settings.')
    } finally {
      setBusy(false)
    }
  }

  const allowedFunctions = isCalculated ? FUNCTIONS.filter(f => f.fn === 'sum') : FUNCTIONS
  const allowedShowAs = isCalculated ? SHOW_AS_MODES.filter(m => m.mode === 'normal' || m.mode === 'percentOfTotal') : SHOW_AS_MODES

  const dialogStyle: CSSProperties = {
    position: 'fixed',
    inset: 0,
    zIndex: 20,
    display: 'grid',
    placeItems: 'center',
    background: '#0005',
  }

  const contentStyle: CSSProperties = {
    background: 'var(--ov-color-surface, white)',
    color: 'var(--ov-color-text, #222)',
    padding: 16,
    width: 'min(440px, calc(100vw - 24px))',
    maxHeight: '85dvh',
    overflowY: 'auto',
    borderRadius: 8,
    boxShadow: '0 4px 24px #0004',
  }

  return (
    <div
      ref={root}
      role="dialog"
      aria-modal="true"
      aria-labelledby="pivot-value-settings-title"
      data-officeview-ui="pivot-value-settings-dialog"
      onKeyDown={onKeyDown}
      onPointerDown={event => event.stopPropagation()}
      style={dialogStyle}
    >
      <section style={contentStyle}>
        <h2 id="pivot-value-settings-title" style={{ margin: '0 0 12px 0', fontSize: 18 }}>Value Field Settings</h2>

        <div style={{ marginBottom: 12 }}>
          <label style={{ display: 'block', marginBottom: 4, fontWeight: 500 }}>
            Source Name
            <input
              type="text"
              readOnly
              disabled
              value={sourceFieldMeta?.name ?? ''}
              className={CONTROL_CLS}
              style={{ ...controlTarget, width: '100%', marginTop: 2, background: 'var(--ov-color-surface-disabled, #eee)' }}
            />
          </label>
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={{ display: 'block', marginBottom: 4, fontWeight: 500 }}>
            Custom Name
            <input
              aria-label="Custom Name"
              type="text"
              value={customName}
              onChange={event => setCustomName(event.target.value)}
              className={CONTROL_CLS}
              style={{ ...controlTarget, width: '100%', marginTop: 2 }}
            />
          </label>
        </div>

        <div role="tablist" aria-label="Settings tabs" style={{ display: 'flex', gap: 8, borderBottom: '1px solid var(--ov-color-border, #ccc)', marginBottom: 12 }}>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'summary'}
            onClick={() => setActiveTab('summary')}
            style={{
              ...controlTarget,
              padding: '6px 12px',
              border: 'none',
              background: 'transparent',
              borderBottom: activeTab === 'summary' ? '2px solid var(--ov-color-primary, #0066cc)' : '2px solid transparent',
              fontWeight: activeTab === 'summary' ? 600 : 400,
              cursor: 'pointer',
            }}
          >
            Summarize Values By
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={activeTab === 'showAs'}
            onClick={() => setActiveTab('showAs')}
            style={{
              ...controlTarget,
              padding: '6px 12px',
              border: 'none',
              background: 'transparent',
              borderBottom: activeTab === 'showAs' ? '2px solid var(--ov-color-primary, #0066cc)' : '2px solid transparent',
              fontWeight: activeTab === 'showAs' ? 600 : 400,
              cursor: 'pointer',
            }}
          >
            Show Values As
          </button>
        </div>

        {activeTab === 'summary' && (
          <div role="tabpanel" aria-label="Summarize Values By" style={{ marginBottom: 12 }}>
            <fieldset style={{ border: '1px solid var(--ov-color-border, #ccc)', borderRadius: 4, padding: 8 }}>
              <legend>Summarize value field by</legend>
              <div style={{ maxHeight: 180, overflowY: 'auto' }}>
                {allowedFunctions.map(item => (
                  <label key={item.fn} style={{ display: 'block', padding: '4px 0', cursor: 'pointer' }}>
                    <input
                      type="radio"
                      name="summaryFunction"
                      value={item.fn}
                      checked={fn === item.fn}
                      onChange={() => setFn(item.fn)}
                      style={{ marginRight: 6 }}
                    />
                    {item.label}
                  </label>
                ))}
              </div>
            </fieldset>
          </div>
        )}

        {activeTab === 'showAs' && (
          <div role="tabpanel" aria-label="Show Values As" style={{ marginBottom: 12 }}>
            <label style={{ display: 'block', marginBottom: 8 }}>
              Show values as
              <select
                aria-label="Show values as"
                value={showAs}
                onChange={event => setShowAs(event.target.value as PivotValueShowAs)}
                className={SELECT_CLS}
                style={{ ...controlTarget, width: '100%', marginTop: 2 }}
              >
                {allowedShowAs.map(item => (
                  <option key={item.mode} value={item.mode}>{item.label}</option>
                ))}
              </select>
            </label>

            {requiresBaseField(showAs) && (
              <label style={{ display: 'block', marginBottom: 8 }}>
                Base field
                <select
                  aria-label="Base field"
                  value={baseField}
                  onChange={event => setBaseField(Number(event.target.value))}
                  className={SELECT_CLS}
                  style={{ ...controlTarget, width: '100%', marginTop: 2 }}
                >
                  {availableBaseFields.map(field => (
                    <option key={field.index} value={field.index}>{field.name}</option>
                  ))}
                </select>
              </label>
            )}

            {requiresBaseItem(showAs) && (
              <label style={{ display: 'block', marginBottom: 8 }}>
                Base item
                <select
                  aria-label="Base item"
                  value={baseItem}
                  onChange={event => setBaseItem(Number(event.target.value))}
                  className={SELECT_CLS}
                  style={{ ...controlTarget, width: '100%', marginTop: 2 }}
                >
                  <option value={1048828}>(previous)</option>
                  <option value={1048829}>(next)</option>
                  {chosenBaseFieldMeta?.items.map(item => (
                    <option key={item.index} value={item.index}>{item.label || '(blank)'}</option>
                  ))}
                </select>
              </label>
            )}
          </div>
        )}

        <div style={{ marginBottom: 16 }}>
          <label style={{ display: 'block', marginBottom: 4, fontWeight: 500 }}>
            Number Format
            <select
              aria-label="Number Format"
              value={numFmtId}
              onChange={event => setNumFmtId(event.target.value)}
              className={SELECT_CLS}
              style={{ ...controlTarget, width: '100%', marginTop: 2 }}
            >
              <option value="">Default (none)</option>
              {PIVOT_VALUE_NUMFMT_PRESETS.map(preset => (
                <option key={preset.numFmtId} value={preset.numFmtId}>
                  {preset.label} ({preset.sample})
                </option>
              ))}
              {currentData?.numFmtId !== undefined && currentData.numFmtId >= 164 && (
                <option value={currentData.numFmtId}>
                  Saved custom format (#{currentData.numFmtId})
                </option>
              )}
            </select>
          </label>
        </div>

        {error && <p role="alert" style={{ color: 'var(--ov-color-error, #d32f2f)', margin: '8px 0' }}>{error}</p>}

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
          <button
            type="button"
            className={CONTROL_PRIMARY_BTN_CLS}
            style={controlTarget}
            disabled={busy}
            onClick={() => void apply()}
          >
            OK
          </button>
          <button
            type="button"
            className={CONTROL_BTN_CLS}
            style={controlTarget}
            onClick={onClose}
          >
            Cancel
          </button>
        </div>
      </section>
    </div>
  )
}
