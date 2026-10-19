import { useEffect, useId, useState, type ReactElement } from 'react'
import type { NavigatorRequest } from '../../xlsx/view/coordinates'
import { navigatorRequestFromA1 } from '../../xlsx/view/scrollWindow'

export interface SheetNavigatorProps {
  documentId: string
  sheets: readonly { id: string; name: string }[]
  activeSheetId: string
  onSheetChange: (sheetId: string) => void
  onNavigate: (request: NavigatorRequest) => void
  onNotice?: (message: string) => void
  disabled?: boolean
}
/** S5 mounts this controlled shell with its active worker session/viewport. */
export function SheetNavigator({ documentId, sheets, activeSheetId, onSheetChange, onNavigate, onNotice, disabled }: SheetNavigatorProps): ReactElement {
  const [address, setAddress] = useState(''), [space, setSpace] = useState<'display' | 'source'>('display'), [error, setError] = useState(''), [invalid, setInvalid] = useState(false)
  const errorId = useId()
  useEffect(() => { setAddress(''); setError(''); setInvalid(false) }, [documentId, activeSheetId])
  const control = { minHeight: 44, minWidth: 44 } as const
  const row = { display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '4px 8px' } as const
  const controlCls = 'ov:rounded-control ov:border ov:border-border ov:bg-field ov:text-text ov:enabled:hover:bg-hover ov:disabled:text-muted ov:disabled:opacity-50 ov:disabled:cursor-default ov:px-2 ov:py-1 ov:text-xs'
  return <nav aria-label="Worksheet navigation" data-officeview-ui="sheet-navigator" style={row}
    onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
    onKeyDown={event => { if (!((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f')) event.stopPropagation() }}>
    <label>Worksheet <select aria-label="Worksheet" value={activeSheetId} disabled={disabled} className={controlCls} style={control} onChange={event => onSheetChange(event.target.value)}>
      {sheets.map(sheet => <option key={sheet.id} value={sheet.id}>{sheet.name}</option>)}
    </select></label>
    <form style={row} onSubmit={event => {
      event.preventDefault()
      if (disabled) return
      let request: NavigatorRequest
      try { request = navigatorRequestFromA1(address, space) } catch {
        const notice = 'Enter a valid cell address through XFD1048576.'
        setError(notice)
        setInvalid(true)
        onNotice?.(notice)
        return
      }
      setError('')
      setInvalid(false)
      try { onNavigate(request) }
      catch (error) {
        const notice = error instanceof Error ? error.message.slice(0, 256) : 'Navigation failed. Please retry.'
        setError(notice)
        onNotice?.(notice)
      }
    }}>
      <label>Address space <select aria-label="Address space" value={space} disabled={disabled} className={controlCls} style={control} onChange={event => setSpace(event.target.value === 'source' ? 'source' : 'display')}>
        <option value="display">Displayed cell</option><option value="source">Original cell</option>
      </select></label>
      <label>Go To cell <input aria-label="Go To cell" value={address} onChange={event => { setAddress(event.target.value); setError(''); setInvalid(false) }} maxLength={16}
        disabled={disabled} className={controlCls} style={control} aria-invalid={invalid} aria-describedby={error ? errorId : undefined} placeholder="A1" /></label>
      <button type="submit" disabled={disabled} className={`${controlCls} ov:font-medium`} style={control}>Go To</button>
      {error && <span id={errorId}>{error}</span>}
    </form>
  </nav>
}
