/**
 * Stage 3b print in the remote view (spec §4.2): the Print button, the
 * page-range dialog under a cap, the progress line and the notices. Only
 * server-drawn print rasters (`RemoteOfficeDoc.print`) are printed, through
 * the shared hidden-iframe surface (`printRasters`, loaded on demand); the
 * on-screen canvases never print. Every element is `data-ov-remote-chrome`.
 */
import { useEffect, useRef, useState, type CSSProperties, type KeyboardEvent, type ReactElement, type RefObject } from 'react'
import type { Manifest } from '../../stream/manifest'
import type { RemoteOfficeDoc } from '../../stream/remote/remote-doc'

const FAILED = 'The document could not be printed. Please try again.'

/** The notice for a refused or failed print job. */
export function printFailure(e: unknown): string {
  const { code, limit } = (e ?? {}) as { code?: unknown; limit?: unknown }
  const n = typeof limit === 'number' ? limit : undefined
  if (code === 'print-too-large') return n !== undefined ? `Too many pages to print at once (limit: ${n} pages).` : 'This document is too large to print from the viewer.'
  if (code === 'print-quota-exhausted' && n !== undefined) return `Print limit for this session reached (${n} pages).`
  if (code === 'rate-limited') return 'Too many requests. Try again shortly.'
  return FAILED
}

const capped = (caps: NonNullable<Manifest['caps']['print']>) => caps.maxPages !== null || caps.maxSessionPages !== null
/** A 1-based page number in 1..count, or undefined. */
const pageNo = (s: string, count: number) => {
  const n = s.trim() === '' ? NaN : Number(s)
  return Number.isInteger(n) && n >= 1 && n <= count ? n : undefined
}

export interface RemotePrint {
  /** The button, dialog, progress line and notice (null when print is not allowed and nothing is showing). */
  ui: ReactElement | null
  /** Ctrl/Cmd+P: starts the flow when allowed; ignored while a job runs or the dialog is open. */
  start(): void
}

const BAR: CSSProperties = { position: 'absolute', top: 8, left: 12, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', maxWidth: 'calc(50% - 12px)' }
const SURFACE: CSSProperties = { background: 'var(--ov-color-chrome, #fff)', color: 'var(--ov-color-text, inherit)', border: '1px solid var(--ov-color-border, #ccc)', borderRadius: 'var(--ov-radius-control, 6px)', padding: '4px 8px' }
const BUTTON: CSSProperties = { ...SURFACE, cursor: 'pointer', font: 'inherit', minHeight: 28 }

export function useRemotePrint(doc: RemoteOfficeDoc, o: { allowPrint: boolean; title?: string; themeRoot: RefObject<HTMLElement | null> }): RemotePrint {
  // A resume may replace the manifest (caps and can.print): re-render and re-read.
  const [, setManifestRev] = useState(0)
  useEffect(() => doc.onManifestChange?.(() => setManifestRev((n) => n + 1)), [doc])
  // A session that ends (revoked, expired, closed…) can print nothing more:
  // print is then off for this document, which stops the network job and a
  // surface still importing or loading its images (the frame never prints).
  // Keyed by the document, so an old document's late end changes nothing.
  const [endedDoc, setEndedDoc] = useState<RemoteOfficeDoc>()
  useEffect(() => {
    let live = true
    void doc.ended.then(() => { if (live) setEndedDoc(doc) })
    return () => { live = false }
  }, [doc])
  const ended = endedDoc === doc
  // caps.print is null only when the server has print off: a `can.print`
  // that says otherwise is contradictory and fails closed.
  const allowed = o.allowPrint && !ended && doc.manifest.can.print && doc.manifest.caps.print !== null
  const caps = doc.manifest.caps.print
  const count = doc.manifest.units.list.length

  const [range, setRange] = useState<{ from: string; to: string }>()
  const [progress, setProgress] = useState<{ done: number; total: number }>()
  const [notice, setNotice] = useState<string>()
  const job = useRef<AbortController | null>(null)
  // The last job handed to the print surface: its frame lives until afterprint.
  const surface = useRef<AbortController | null>(null)
  const stopAll = () => { job.current?.abort(); job.current = null; surface.current?.abort(); surface.current = null }
  const printBtn = useRef<HTMLButtonElement | null>(null)
  const fromInput = useRef<HTMLInputElement | null>(null)
  const dialogRef = useRef<HTMLDivElement | null>(null)

  // Losing permission (or a new document) aborts a running job silently and closes the dialog.
  useEffect(() => {
    if (allowed) return
    stopAll()
    setRange(undefined); setProgress(undefined)
  }, [allowed])
  // After a session end the viewer's ended banner is the only message.
  useEffect(() => { if (ended) setNotice(undefined) }, [ended])
  useEffect(() => stopAll, [doc]) // eslint-disable-line react-hooks/exhaustive-deps
  // A replaced document starts clean: the old job's progress, range and notice
  // are reset during render (the old job is aborted by the cleanup above, and
  // its late callbacks are dropped by `live()`).
  const [shownDoc, setShownDoc] = useState(doc)
  if (shownDoc !== doc) { setShownDoc(doc); setRange(undefined); setProgress(undefined); setNotice(undefined) }
  useEffect(() => { if (range) fromInput.current?.focus() }, [range !== undefined]) // eslint-disable-line react-hooks/exhaustive-deps

  const run = (from: number, to: number) => {
    if (job.current || !allowed || count < 1) return
    surface.current?.abort(); surface.current = null
    const ac = new AbortController()
    job.current = ac
    const live = () => !ac.signal.aborted && job.current === ac
    const finish = () => { if (job.current === ac) { job.current = null; setProgress(undefined) } }
    setNotice(undefined)
    setProgress({ done: 0, total: to - from + 1 })
    void (async () => {
      try {
        const pages = await doc.print({ from, to }, { signal: ac.signal, onProgress: (done, total) => { if (live()) setProgress({ done, total }) } })
        if (!live()) return
        const m = await import('../paged/printDocument')
        // The frame is hosted outside the view's root, which is display:none in
        // print media; the frame's own print() prints only the frame document.
        const root = o.themeRoot.current?.ownerDocument.body
        if (!live() || !root) return
        surface.current = ac
        await m.printRasters(pages, root, {
          signal: ac.signal, ...(o.title !== undefined ? { title: o.title } : {}),
          onNotice: (text) => { if (!ac.signal.aborted) setNotice(text) },
        })
      } catch (e) {
        // An abort (Cancel, unmount, permission lost) shows nothing.
        if (live()) setNotice(printFailure(e))
      } finally { finish() }
    })()
  }

  const closeDialog = () => { setRange(undefined); (printBtn.current ?? o.themeRoot.current)?.focus() }
  const start = () => {
    if (!allowed || job.current || range || count < 1) return
    if (!caps || !capped(caps)) { run(0, count - 1); return }
    setRange({ from: '1', to: String(Math.min(count, caps.maxPages ?? count)) })
  }

  const from = range ? pageNo(range.from, count) : undefined
  const to = range ? pageNo(range.to, count) : undefined
  const valid = from !== undefined && to !== undefined && from <= to
  const submit = () => { if (!valid) return; setRange(undefined); run(from - 1, to - 1) }
  const onDialogKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeDialog() }
    else if (e.key === 'Enter' && (e.target as Element).tagName === 'INPUT') { e.preventDefault(); e.stopPropagation(); submit() }
    else if (e.key === 'Tab') {
      // Focus stays inside the modal (the PivotGroupDialog pattern).
      const nodes = [...(dialogRef.current?.querySelectorAll<HTMLButtonElement | HTMLInputElement>('button, input') ?? [])].filter((n) => !n.disabled)
      if (!nodes.length) return
      const first = nodes[0]!, last = nodes[nodes.length - 1]!, active = dialogRef.current?.ownerDocument.activeElement
      if (e.shiftKey && active === first) { e.preventDefault(); last.focus() }
      else if (!e.shiftKey && active === last) { e.preventDefault(); first.focus() }
    }
  }

  const numberInput = (label: 'From' | 'To', key: 'from' | 'to') => (
    <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>{label}
      <input type="number" min={1} max={count} step={1} value={range![key]} ref={key === 'from' ? fromInput : undefined}
        onChange={(e) => { const v = e.target.value; setRange((r) => r && { ...r, [key]: v }) }}
        style={{ ...SURFACE, width: '5em', background: 'var(--ov-color-field, #fff)' }} />
    </label>
  )
  const running = progress !== undefined
  const ui = allowed || notice ? (
    <div data-ov-remote-chrome="" style={BAR}>
      {allowed ? <button type="button" data-ov-print-button="" ref={printBtn} disabled={running} onClick={start} style={BUTTON}>Print</button> : null}
      {running ? (
        <span style={{ ...SURFACE, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
          <span aria-live="polite">Preparing pages {progress.done} of {progress.total}…</span>
          <button type="button" onClick={() => { job.current?.abort() ; job.current = null; setProgress(undefined) }} style={BUTTON}>Cancel</button>
        </span>
      ) : null}
      {notice ? <span role="alert" data-ov-print-notice="" style={{ ...SURFACE, overflowWrap: 'anywhere' }}>{notice}</span> : null}
      {range && allowed ? (
        <div ref={dialogRef} role="dialog" aria-modal="true" aria-label="Print pages" data-ov-remote-chrome="" onKeyDown={onDialogKey}
          style={{ ...SURFACE, position: 'absolute', top: '100%', left: 0, marginTop: 8, padding: 12, display: 'flex', flexDirection: 'column', gap: 8, minWidth: 220, zIndex: 7, boxShadow: 'var(--ov-shadow-popover, 0 4px 12px rgba(0,0,0,.2))' }}>
          <strong>Print pages</strong>
          <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>{numberInput('From', 'from')}{numberInput('To', 'to')}</div>
          {caps?.maxPages != null ? <div>Up to {caps.maxPages} pages per print.</div> : null}
          {caps?.maxSessionPages != null ? <div>Up to {caps.maxSessionPages} pages in this session.</div> : null}
          <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
            <button type="button" onClick={closeDialog} style={BUTTON}>Cancel</button>
            <button type="button" disabled={!valid} onClick={submit} style={BUTTON}>Print</button>
          </div>
        </div>
      ) : null}
    </div>
  ) : null
  return { ui, start }
}
