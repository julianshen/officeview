import { useEffect, useRef, useState, type ReactElement, type CSSProperties } from 'react'
import type { ColumnRule, FilterSpec, SortSpec } from '../../xlsx/view/contracts'
import { FilterEngine, type ColumnClassification } from '../../xlsx/view/filter'
import { lowerDateCondition } from '../../xlsx/view/datePredicates'
import { decodeSerial } from '../../xlsx/formula/serial'
import { columnToLetters } from '../../xlsx/view/coordinates'
import type { DetectedRange } from '../../xlsx/view/rangeDetection'
import { filterWithColumn, sameFilterRange, type DraftPage, type DraftQuery, type ValueDraftAction, type ValuesDraftPort } from './controlsClient'
import { controlTarget, CONTROL_CLS, CONTROL_BTN_CLS, CONTROL_PRIMARY_BTN_CLS, CHECKBOX_CLS, failureText, useControlLifetime, type ControlProps } from './controlUi'
import { useControlMedia } from './controlMedia'
import { ConditionEditor, conditionInput, typedCondition } from './ConditionEditor'
export interface FilterDropdownProps extends ControlProps { col: number; applied: FilterSpec | null; snapshotHidden?: boolean; appliedSort?: SortSpec | null }
/** Reserved height of the loaded page's controls ((Blanks), the selected-count
 * line, Previous/Next values) so their arrival moves nothing below them. */
const PAGE_BLOCK_MIN_HEIGHT = 160
export function FilterDropdown(props: FilterDropdownProps): ReactElement | null {
  const bottomSheet = useControlMedia('(pointer: coarse), (max-width: 640px)')
  const life = useControlLifetime(props), session = useRef<ValuesDraftPort | null>(null), searchInput = useRef<HTMLInputElement>(null), changingRef = useRef(false)
  const [range, setRange] = useState<DetectedRange | null>(null), [page, setPage] = useState<DraftPage | null>(null), [query, setQuery] = useState<DraftQuery>({ text: '', offset: 0 })
  const [changing, setChanging] = useState(false), [domainClass, setDomainClass] = useState<ColumnClassification | null>(null)
  const [busy, setBusy] = useState(true), [progress, setProgress] = useState('Loading values'), [error, setError] = useState(''), [inspection, setInspection] = useState(''), [includeBlanks, setIncludeBlanks] = useState(false)
  const appliedRule = props.applied?.columns.get(props.col)
  const initial = appliedRule?.kind === 'compound' ? appliedRule.first : appliedRule && appliedRule.kind !== 'values' ? appliedRule : undefined
  const [method, setMethod] = useState<'values' | 'custom'>(initial ? 'custom' : 'values'), [first, setFirst] = useState(() => conditionInput(initial)), [second, setSecond] = useState(() => conditionInput(appliedRule?.kind === 'compound' ? appliedRule.second : undefined))
  const [compound, setCompound] = useState(appliedRule?.kind === 'compound'), [connective, setConnective] = useState<'and' | 'or'>(appliedRule?.kind === 'compound' ? appliedRule.connective : 'and')
  // Typed rules need only the domain identity, so a failed value universe
  // (quota, storage) must not hide the custom builder.
  const classification: ColumnClassification = session.current?.identity.classification ?? domainClass ?? { kind: 'loading', hasNumeric: false }
  life.dispose.current = () => { session.current?.cancel(); session.current = null }
  const noticeError = (message: string) => { setError(message); props.onNotice(message) }
  const initialize = () => {
    const job = life.begin(); changingRef.current = false; setChanging(false); setBusy(true); setError(''); setPage(null); setRange(null); setDomainClass(null)
    setMethod(initial ? 'custom' : 'values'); setFirst(conditionInput(initial)); setSecond(conditionInput(appliedRule?.kind === 'compound' ? appliedRule.second : undefined))
    setCompound(appliedRule?.kind === 'compound'); setConnective(appliedRule?.kind === 'compound' ? appliedRule.connective : 'and'); setInspection(''); setIncludeBlanks(false)
    session.current?.cancel(); session.current = null
    void (async () => {
      const resolved = await props.client.resolve(job.owner, 'filter', job.signal, undefined, props.col)
      if (!job.current()) return
      if (resolved.kind !== 'complete') { noticeError(resolved.kind === 'refused' ? failureText(resolved) : 'Source range is incomplete; retry.'); setBusy(false); return }
      if (props.col < resolved.range.c0 || props.col > resolved.range.c1) throw new RangeError('Column is outside the resolved filter range.')
      setRange(resolved)
      const domain = props.client.classify(job.owner, resolved.range, props.col, job.signal)
      setDomainClass(domain)
      if (!initial) { const kind = domain.kind === 'number' ? 'number' : domain.kind === 'date' ? 'date' : 'text'; setFirst(conditionInput(undefined, kind)); setSecond(conditionInput(undefined, kind)) }
      const rule = props.applied && sameFilterRange(props.applied.range, resolved.range) && appliedRule?.kind === 'values' ? appliedRule : null
      const port = await props.client.values(job.owner, resolved.range, props.col, rule, job.signal)
      if (!job.current()) { port.cancel(); return }
      session.current = port
      const result = await port.read({ text: '', offset: 0 }, job.signal, p => { if (job.current()) setProgress(`${p.message}: ${p.completed} / ${p.total}`) })
      if (job.current()) { setQuery({ text: '', offset: 0 }); setPage(result); setBusy(false) }
    })().catch(error => { if (job.current()) { noticeError(life.error(error)); setBusy(false) } })
  }
  useEffect(() => { if (props.open !== false) { initialize(); searchInput.current?.focus() } }, [props.open, life.ownerKey, props.col])
  const refresh = (next: DraftQuery, action?: ValueDraftAction) => {
    const port = session.current
    if (!port || changingRef.current) return
    if (action) { changingRef.current = true; setChanging(true) }
    const job = life.begin(); setBusy(true); setError(''); setQuery(next)
    void (action ? port.change(action, next, job.signal) : port.read(next, job.signal, p => { if (job.current()) setProgress(`${p.message}: ${p.completed} / ${p.total}`) }))
      .then(result => { if (job.current()) { changingRef.current = false; setChanging(false); setPage(result); setBusy(false) } }).catch(error => { if (job.current()) { changingRef.current = false; setChanging(false); noticeError(life.error(error)); setBusy(false) } })
  }
  let custom: ColumnRule | undefined, validation = method === 'custom' && classification.kind === 'loading' ? 'Column classification is pending; wait and retry.' : ''
  if (method === 'custom' && classification.kind !== 'loading') {
    try {
      const one = typedCondition(first), two = compound ? typedCondition(second) : undefined
      custom = two ? { kind: 'compound', connective, first: one, second: two } : one
      new FilterEngine(props.client.context()).compileRule(custom, classification)
    } catch (error) { validation = life.error(error) }
  }
  const periods: string[] = []
  if (method === 'custom' && custom && !validation) {
    for (const rule of custom.kind === 'compound' ? [custom.first, custom.second] : [custom]) if (rule.kind === 'datePreset') {
      const context = props.client.context(), interval = lowerDateCondition(rule, context)
      const day = (serial: number) => { const decoded = decodeSerial(serial, context.dateSystem); if ('kind' in decoded) return 'unavailable'; const { year, month, day } = decoded.civil; return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}` }
      periods.push(`${rule.preset}: ${day(interval.start)} through ${day(interval.end - 1)} (${context.timeZone}); re-resolved when applied.`)
    }
  }
  const canApply = Boolean(range) && (method === 'values' ? Boolean(page?.complete) : Boolean(custom) && !validation)
  const apply = (shortcut?: 'clear' | 'clearAll' | 'asc' | 'desc') => {
    if (busy || life.invalid || (!shortcut && !canApply)) return
    const job = life.begin(); setBusy(true); setError('')
    void (async () => {
      let outcome
      if (shortcut === 'asc' || shortcut === 'desc') {
        const detected = await props.client.resolve(job.owner, 'sort', job.signal, undefined, props.col)
        if (!job.current()) return
        if (detected.kind !== 'complete') { noticeError(detected.kind === 'refused' ? failureText(detected) : 'Source range is incomplete; retry.'); setBusy(false); return }
        outcome = await props.client.sort(job.owner, { range: detected.range, keys: [{ col: props.col, order: shortcut }], caseSensitive: props.appliedSort?.caseSensitive ?? false }, job.signal)
      } else {
        let next: FilterSpec | null
        if (shortcut === 'clearAll') next = null
        else if (shortcut === 'clear') next = props.applied ? filterWithColumn(props.applied, props.applied.range, props.col, null) : null
        else {
          const rule = method === 'values' ? await session.current!.candidate(job.signal) : custom!
          if (!job.current()) return
          next = filterWithColumn(props.applied, range!.range, props.col, rule)
        }
        outcome = await props.client.filter(job.owner, next, job.signal, p => { if (job.current()) setProgress(`${p.message}: ${p.completed} / ${p.total}`) }, !shortcut && method === 'values' ? session.current! : undefined)
      }
      if (!job.accepts(outcome)) return
      const replacement = !shortcut && props.applied && range && !sameFilterRange(props.applied.range, range.range)
      const text = shortcut === 'asc' || shortcut === 'desc' ? 'View-only sort applied; formulas unchanged.' : shortcut ? `Filter cleared.${props.snapshotHidden ? " File's saved hidden rows remain hidden." : ''}` : `${outcome.kind === 'unchanged' ? 'Filter unchanged.' : 'Filter applied.'}${replacement ? ' Previous filter range replaced.' : ''}`
      if (!life.result(outcome, text)) { setError(failureText(outcome as Parameters<typeof failureText>[0])); setBusy(false) }
    })().catch(error => { if (job.current()) { noticeError(life.error(error)); setBusy(false) } })
  }
  if (props.open === false) return null
  return <div role={props.embedded ? undefined : 'dialog'} aria-label={`Filter column ${columnToLetters(props.col)}`} data-officeview-ui="filter-dropdown"
    className={bottomSheet ? "ov:rounded-t-[12px] ov:shadow-popover ov:bg-surface ov:border-t ov:border-border" : "ov:rounded-panel ov:shadow-popover ov:bg-surface ov:border ov:border-border"}
    onPointerDown={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
    onKeyDown={event => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') return
      event.stopPropagation()
      if (event.key === 'Escape') { event.preventDefault(); life.cancel() }
      else if (event.key === 'Enter' && event.target === searchInput.current) { event.preventDefault(); refresh({ ...query, offset: 0 }) }
    }} style={{ position: bottomSheet ? 'fixed' : 'relative', bottom: bottomSheet ? 0 : undefined, left: bottomSheet ? 0 : undefined, zIndex: bottomSheet ? 4 : undefined, width: bottomSheet ? '100%' : 'min(420px, calc(100vw - 16px))',
      // Embedded on desktop, the bounded grid popup is the one scroll container (no nested scroller taller than the pane).
      maxHeight: bottomSheet ? '85dvh' : props.embedded ? undefined : '80vh', overflowY: bottomSheet || !props.embedded ? 'auto' : undefined, padding: '12px', paddingBottom: 'calc(12px + var(--officeview-safe-bottom))', '--officeview-safe-bottom': 'env(safe-area-inset-bottom, 0px)', boxSizing: 'border-box' } as CSSProperties}>
    <form onSubmit={event => { event.preventDefault(); apply() }}>
      <h2>Filter column {columnToLetters(props.col)}</h2>
      <div><button type="button" className={CONTROL_BTN_CLS} style={controlTarget} disabled={busy || !range || life.invalid} onClick={() => apply('asc')}>Sort ascending</button><button type="button" className={CONTROL_BTN_CLS} style={controlTarget} disabled={busy || !range || life.invalid} onClick={() => apply('desc')}>Sort descending</button></div>
      <button type="button" className={CONTROL_BTN_CLS} style={controlTarget} disabled={busy || life.invalid} onClick={() => apply('clear')}>Clear column</button><button type="button" className={CONTROL_BTN_CLS} style={controlTarget} disabled={busy || life.invalid} onClick={() => apply('clearAll')}>Clear all filters</button>
      {/* Space reserved for what the value list's arrival fills in, so nothing below moves (no layout shift). */}
      <p style={{ minHeight: '2.6em' }}>{range && <>{range.source.kind === 'table' ? `Table metadata: ${range.range.headerRows} header row(s), ${range.range.totalsRows} totals row(s).` : 'Resolved source range.'}{props.applied && !sameFilterRange(props.applied.range, range.range) ? ' OK will replace the previous filter range.' : ''}</>}</p>
      <label>Filter method <select aria-label="Filter method" className={CONTROL_CLS} style={controlTarget} value={method} onChange={event => setMethod(event.target.value as 'values' | 'custom')}><option value="values">Values and blanks</option>{(method === 'custom' || !['loading', 'empty'].includes(classification.kind)) && <option value="custom">{classification.kind === 'loading' ? 'Custom filters (classification pending)' : classification.kind === 'date' ? 'Date filters' : classification.kind === 'number' ? 'Number filters' : 'Text / mixed filters'}</option>}</select></label>
      {method === 'values' ? <>
        <label>Search <input ref={searchInput} aria-label="Search all values" className={CONTROL_CLS} disabled={changing || life.invalid} style={controlTarget} value={query.text} onChange={event => refresh({ text: event.target.value, offset: 0, wildcards: query.wildcards })} /></label>
        <label style={{ ...controlTarget, display: 'flex', alignItems: 'center' }}><input type="checkbox" className={CHECKBOX_CLS} disabled={changing || life.invalid} checked={Boolean(query.wildcards)} onChange={event => refresh({ ...query, offset: 0, wildcards: event.target.checked })} />Search with wildcards</label>
        <p style={{ minHeight: '1.3em' }}>{page?.truncated ? 'Showing first 10,000; search all values.' : ''}</p>
        <button type="button" className={CONTROL_BTN_CLS} style={controlTarget} disabled={busy || !page || life.invalid} onClick={() => refresh(query, { kind: 'all', selected: true, includeBlanks })}>Select All</button>
        <button type="button" className={CONTROL_BTN_CLS} style={controlTarget} disabled={busy || !page || life.invalid} onClick={() => refresh(query, { kind: 'all', selected: false, includeBlanks })}>Deselect All</button>
        {query.text && <label style={{ ...controlTarget, display: 'flex', alignItems: 'center' }}><input type="checkbox" className={CHECKBOX_CLS} disabled={changing || life.invalid} checked={includeBlanks} onChange={event => setIncludeBlanks(event.target.checked)} />Include blanks in searched Select All</label>}
        <div style={{ height: 280, overflowY: 'auto' }}>{page?.entries.map(entry => <label key={entry.ordinal} data-value-ordinal={entry.ordinal} style={{ ...controlTarget, display: 'flex', alignItems: 'center' }}>
          <input type="checkbox" className={CHECKBOX_CLS} aria-label={`${entry.label}${entry.kind === 'error' ? ' (error)' : ''}`} checked={entry.selected === true} disabled={busy || entry.selected === 'pending' || life.invalid}
            ref={input => { if (input) input.indeterminate = entry.selected === 'pending' }} onChange={event => refresh(query, { kind: 'toggle', ordinal: entry.ordinal, selected: event.target.checked })} />
          <span>{entry.label}{entry.kind === 'error' ? ' (error)' : ''}{entry.selected === 'pending' ? ' — selection pending' : ''}</span>
          <button type="button" aria-label={`Inspect value ${entry.ordinal + 1}`} className={CONTROL_BTN_CLS} style={controlTarget} disabled={busy || life.invalid} onClick={() => { if (busy || life.invalid) return; const job = life.begin(); void session.current!.inspect(entry.ordinal, job.signal).then(value => { if (job.current()) setInspection(value) }).catch(error => { if (job.current()) noticeError(life.error(error)) }) }}>Inspect</button>
        </label>)}</div>
        <div style={{ minHeight: PAGE_BLOCK_MIN_HEIGHT }}>{page && <><label style={{ ...controlTarget, display: 'flex', alignItems: 'center' }}><input type="checkbox" className={CHECKBOX_CLS} aria-label="(Blanks)" checked={page.blanks} disabled={busy || life.invalid} onChange={event => refresh(query, { kind: 'blanks', selected: event.target.checked })} />(Blanks)</label>
          <p>{page.count.count} selected values{page.count.pending ? `; ${page.count.pending} pending` : ''}; {page.count.complete ? 'complete universe' : 'universe incomplete'}.</p>
          <button type="button" className={CONTROL_BTN_CLS} style={controlTarget} disabled={busy || query.offset === 0} onClick={() => refresh({ ...query, offset: Math.max(0, query.offset - 200) })}>Previous values</button><button type="button" className={CONTROL_BTN_CLS} style={controlTarget} disabled={busy || query.offset + 200 >= page.displayed} onClick={() => refresh({ ...query, offset: query.offset + 200 })}>Next values</button></>}</div>
      </> : <>
        <ConditionEditor value={first} onChange={setFirst} classification={classification} index={1} invalid={Boolean(validation)} />
        <label style={{ ...controlTarget, display: 'flex', alignItems: 'center' }}><input type="checkbox" className={CHECKBOX_CLS} aria-label="Use two conditions" checked={compound} onChange={event => setCompound(event.target.checked)} />Use two conditions</label>
        {compound && <><label>Combine <select aria-label="Combine conditions" className={CONTROL_CLS} style={controlTarget} value={connective} onChange={event => setConnective(event.target.value as 'and' | 'or')}><option value="and">All (AND)</option><option value="or">Either (OR)</option></select></label><ConditionEditor value={second} onChange={setSecond} classification={classification} index={2} invalid={Boolean(validation)} /></>}
        {classification.kind === 'date' && <p>Calendar dates in {props.client.context().timeZone}; periods include times and resolve at apply time.</p>}
      </>}
      {periods.map(period => <p key={period}>{period}</p>)}
      {inspection && <pre>{inspection}</pre>}{validation && <p>{validation}</p>}{error && <p>{error}</p>}
      <div style={{ minHeight: controlTarget.minHeight }}>
        {busy && <div role="progressbar" aria-label={progress}>{progress}</div>}
        {!busy && (!page?.complete || error) && <button type="button" className={CONTROL_BTN_CLS} style={controlTarget} onClick={initialize}>Retry values</button>}
      </div>
      {/* OK/Cancel stay in view at the bottom of the scrolling popup. */}
      <div className="ov:bg-surface" style={{ position: 'sticky', bottom: 'calc(-12px - var(--officeview-safe-bottom))', margin: '0 -12px calc(-12px - var(--officeview-safe-bottom))', padding: '4px 12px calc(12px + var(--officeview-safe-bottom))' }}><button type="submit" className={CONTROL_PRIMARY_BTN_CLS} style={controlTarget} disabled={busy || !canApply || life.invalid}>OK</button><button type="button" className={CONTROL_BTN_CLS} style={controlTarget} onClick={() => life.cancel()}>Cancel</button></div>
    </form>
  </div>
}
