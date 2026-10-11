import { useEffect, useLayoutEffect, useRef, useState, type ReactElement } from 'react'
import type { SortSpec } from '../../xlsx/view/contracts'
import { columnToLetters, lettersToColumn } from '../../xlsx/view/coordinates'
import type { DetectedRange } from '../../xlsx/view/rangeDetection'
import { sameFilterRange } from './controlsClient'
import { controlTarget, failureText, useControlLifetime, type ControlProps } from './controlUi'
export interface SortDialogProps extends ControlProps { applied: SortSpec | null; col?: number }
interface Level { id: number; column: string; order: 'asc' | 'desc' }
export function SortDialog(props: SortDialogProps): ReactElement | null {
  const life = useControlLifetime(props), root = useRef<HTMLDivElement>(null), ids = useRef(0), focus = useRef<string | null>(null)
  const [range, setRange] = useState<DetectedRange | null>(null), [levels, setLevels] = useState<Level[]>([]), [caseSensitive, setCaseSensitive] = useState(props.applied?.caseSensitive ?? false)
  const [busy, setBusy] = useState(true), [error, setError] = useState(''), [progress, setProgress] = useState('Resolving source range')
  const resolve = (override?: 0 | 1, reset = false) => {
    const job = life.begin(); setBusy(true); setError('')
    void props.client.resolve(job.owner, 'sort', job.signal, override, props.col).then(result => {
      if (!job.current()) return
      if (result.kind !== 'complete') { const message = result.kind === 'refused' ? failureText(result) : 'Source range is incomplete; retry.'; setError(message); props.onNotice(message); setBusy(false); return }
      setRange(result)
      if (reset || !levels.length) {
        const applied = props.applied && sameFilterRange(props.applied.range, result.range) ? props.applied : null
        setLevels((applied?.keys ?? [{ col: props.col !== undefined && props.col >= result.range.c0 && props.col <= result.range.c1 ? props.col : result.range.c0, order: 'asc' as const }]).map(key => ({ id: ++ids.current, column: columnToLetters(key.col), order: key.order })))
      }
      setBusy(false)
    }).catch(error => { if (job.current()) { const message = life.error(error); setError(message); props.onNotice(message); setBusy(false) } })
  }
  useEffect(() => { if (props.open !== false) { setLevels([]); setCaseSensitive(props.applied?.caseSensitive ?? false); resolve(undefined, true) } }, [props.open, life.ownerKey])
  useLayoutEffect(() => {
    if (focus.current) { root.current?.querySelector<HTMLElement>(focus.current)?.focus(); focus.current = null }
    else if (levels.length === 1 && !root.current?.contains(document.activeElement)) root.current?.querySelector<HTMLInputElement>('input')?.focus()
  }, [levels])
  let spec: SortSpec | undefined, validation = ''
  if (range) {
    try {
      if (!levels.length || levels.length > 8) throw new RangeError('Choose 1 to 8 sort levels.')
      const keys = levels.map(level => ({ col: lettersToColumn(level.column.trim().toUpperCase()), order: level.order }))
      if (new Set(keys.map(key => key.col)).size !== keys.length || keys.some(key => key.col < range.range.c0 || key.col > range.range.c1)) throw new RangeError('Choose unique columns inside the resolved range.')
      spec = { range: range.range, keys, caseSensitive }
    } catch (error) { validation = life.error(error) }
  }
  const add = () => {
    if (!range || levels.length >= 8) return
    const used = new Set(levels.map(level => level.column.trim().toUpperCase()))
    let col = range.range.c0
    while (col <= range.range.c1 && used.has(columnToLetters(col))) col++
    if (col > range.range.c1) return
    const id = ++ids.current
    focus.current = `[data-sort-level="${id}"] input`
    setLevels([...levels, { id, column: columnToLetters(col), order: 'asc' }])
  }
  const move = (index: number, delta: -1 | 1) => {
    const next = [...levels], at = index + delta
    if (at < 0 || at >= next.length) return
    focus.current = `[data-sort-level="${next[index].id}"] [data-move="${delta}"]`
    ;[next[index], next[at]] = [next[at], next[index]]
    setLevels(next)
  }
  const apply = () => {
    if (!spec || busy || life.invalid) return
    const job = life.begin(); setBusy(true); setError('')
    void props.client.sort(job.owner, spec, job.signal, p => { if (job.current()) setProgress(`${p.message}: ${p.completed} / ${p.total}`) }).then(result => {
      if (!job.accepts(result)) return
      if (!life.result(result, 'View-only sort applied; formulas unchanged.')) { setError(failureText(result as Parameters<typeof failureText>[0])); setBusy(false) }
    }).catch(error => { if (job.current()) { const message = life.error(error); setError(message); props.onNotice(message); setBusy(false) } })
  }
  if (props.open === false) return null
  return <div ref={root} role={props.embedded ? undefined : 'dialog'} aria-label="Sort worksheet" aria-modal={props.embedded ? undefined : true} data-officeview-ui="sort-dialog"
    onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
    onKeyDown={event => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') { life.cancel(false); return }
      event.stopPropagation()
      if (event.key === 'Escape') { event.preventDefault(); life.cancel() }
      else if (event.key === 'Tab') {
        const controls = [...root.current!.querySelectorAll<HTMLElement>('input:not(:disabled), select:not(:disabled), button:not(:disabled), [tabindex="0"]')]
        const first = controls[0], last = controls[controls.length - 1]
        if (event.shiftKey && (document.activeElement === first || !root.current!.contains(document.activeElement))) { event.preventDefault(); last?.focus() }
        else if (!event.shiftKey && (document.activeElement === last || !root.current!.contains(document.activeElement))) { event.preventDefault(); first?.focus() }
      }
    }} style={{ width: 'min(560px, calc(100vw - 16px))', maxHeight: '80vh', overflowY: 'auto', background: '#fff', border: '1px solid #888', padding: '12px 12px calc(12px + env(safe-area-inset-bottom))', boxSizing: 'border-box' }}>
    <form onSubmit={event => { event.preventDefault(); apply() }}>
      <h2>Sort worksheet</h2><p>View-only sort; formulas unchanged.</p>
      {levels.map((level, index) => <fieldset key={level.id} data-sort-level={level.id}><legend>Level {index + 1}</legend>
        <label>Column letter <input aria-label={`Column level ${index + 1}`} style={controlTarget} value={level.column} aria-invalid={Boolean(validation)} onChange={event => setLevels(levels.map(item => item.id === level.id ? { ...item, column: event.target.value } : item))} /></label>
        <label>Direction <select aria-label={`Direction level ${index + 1}`} style={controlTarget} value={level.order} onChange={event => setLevels(levels.map(item => item.id === level.id ? { ...item, order: event.target.value as 'asc' | 'desc' } : item))}><option value="asc">Ascending</option><option value="desc">Descending</option></select></label>
        <button type="button" style={controlTarget} onClick={() => { focus.current = levels.length > 1 ? `[data-sort-level="${levels[index === 0 ? 1 : 0].id}"] input` : '[data-add-level]'; setLevels(levels.filter(item => item.id !== level.id)) }}>Remove level {index + 1}</button>
        <button type="button" data-move={-1} style={controlTarget} disabled={index === 0} onClick={() => move(index, -1)}>Move level {index + 1} up</button><button type="button" data-move={1} style={controlTarget} disabled={index === levels.length - 1} onClick={() => move(index, 1)}>Move level {index + 1} down</button>
      </fieldset>)}
      <button type="button" data-add-level="" style={controlTarget} disabled={busy || levels.length >= 8 || Boolean(range && levels.length >= range.range.c1 - range.range.c0 + 1)} onClick={add}>Add level</button>
      <label style={{ ...controlTarget, display: 'flex', alignItems: 'center' }}><input type="checkbox" aria-label="Case sensitive" checked={caseSensitive} onChange={event => setCaseSensitive(event.target.checked)} />Case sensitive</label>
      {range && <><label style={{ ...controlTarget, display: 'flex', alignItems: 'center' }}><input type="checkbox" aria-label="Range has a header row" checked={range.range.headerRows === 1} disabled={busy || range.source.kind === 'table'} onChange={event => resolve(event.target.checked ? 1 : 0)} />Range has a header row</label>
        {range.source.kind === 'table' && <p>Table metadata protects {range.range.headerRows} header row(s) and {range.range.totalsRows} totals row(s).</p>}</>}
      {validation && <p>{validation}</p>}{error && <p>{error}</p>}{busy && <div role="progressbar" aria-label={progress}>{progress}</div>}
      {!busy && !range && <button type="button" style={controlTarget} onClick={() => resolve()}>Retry range</button>}
      <button type="submit" style={controlTarget} disabled={busy || !spec || life.invalid}>OK</button><button type="button" style={controlTarget} onClick={() => life.cancel()}>Cancel</button>
    </form>
  </div>
}
