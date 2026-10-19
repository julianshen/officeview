/**
 * The Word/PowerPoint top toolbar (spec §2 item 8, §6.5): file name, page
 * navigation, zoom out / readout / zoom in / fit-to-width, Find, copy and
 * print. Compact (coarse pointer or a narrow pane): 44px targets, no file
 * name, Find collapsed to an icon button, and Fit/Copy/Print in a More menu.
 * Presentational: all document state arrives as props (see PagedChrome).
 */
import { useLayoutEffect, useRef, useState, type ReactElement, type ReactNode } from 'react'
import { Ban, Check, ChevronLeft, ChevronRight, Copy, Minus, MoveHorizontal, Plus, Printer, Search } from 'lucide-react'
import { Icon } from '../Icon'
import { MoreMenu, type MoreItem } from './MoreMenu'
import { m } from 'motion/react'
import { Presence, Tick, tick, useMotionProps } from '../motion/presets'
import { PAGED_TOOLBAR_HEIGHT } from './metrics'
import type { PagedKind } from './PagedStatus'

export interface PagedToolbarProps {
  title?: string
  kind: PagedKind
  /** 1-based page (or slide) in view. */
  current: number
  total: number
  onPrev(): void
  onNext(): void
  /** Shown zoom (z0 while fit mode is on). */
  zoom: number
  onZoomIn(): void
  onZoomOut(): void
  /** The readout button resets the zoom (aria-label "Reset zoom"). */
  onResetZoom(): void
  onFit(): void
  fitActive: boolean
  /** The inline Find bar. */
  search?: ReactNode
  /** Omitted when copying is not allowed. */
  onCopy?: () => void
  /** False while there is no selection to copy. */
  copyEnabled?: boolean
  copyState: 'idle' | 'done' | 'failed'
  /** Omitted when printing is not allowed. */
  onPrint?: () => void
  compact: boolean
  showSearch: boolean
  showZoom: boolean
  /** A non-zero value not seen before expands Find and focuses its input. */
  focusSearchRequest: number
  onEscapeToDocument(): void
}

const BLOCKED = 'Copy blocked — press ⌘C'
const COPY_STATE = { idle: 'idle', done: 'copied', failed: 'blocked' } as const
const MOTION = 'ov:transition-colors ov:duration-150 ov:motion-reduce:transition-none'
const FOCUS = 'ov:focus-visible:outline-2 ov:focus-visible:outline-solid ov:focus-visible:outline-offset-1 ov:focus-visible:outline-accent'
const BASE = `ov:inline-flex ov:items-center ov:justify-center ov:gap-1 ov:shrink-0 ov:py-0 ov:px-1 ov:rounded-control ov:bg-transparent ov:border-0 ov:text-icon ov:cursor-pointer ov:enabled:hover:bg-hover ov:disabled:text-muted ov:disabled:cursor-default ov:aria-pressed:bg-header-active ov:aria-pressed:text-accent ${MOTION} ${FOCUS}`
/** Regular: 28px controls in a 40px bar (44px on a coarse pointer); compact: 44px. */
const button = (compact: boolean) => `${BASE} ${compact ? 'ov:min-h-11 ov:min-w-11' : 'ov:h-7 ov:min-w-7 ov:coarse:min-h-11 ov:coarse:min-w-11'}`
const DIVIDER = 'ov:w-px ov:h-5 ov:mx-1 ov:shrink-0 ov:bg-border'
/** Compact drops the group dividers: 44px targets need the width at 375px. */
const divider = (compact: boolean) => (compact ? null : <span className={DIVIDER} aria-hidden="true" />)

/** Where focus goes when the layout crosses the compact threshold. */
const COUNTERPART: Record<string, string> = { fit: 'more', copy: 'more', print: 'more', more: 'fit', find: 'search' }

export function PagedToolbar(p: PagedToolbarProps): ReactElement {
  const { compact, kind } = p
  const root = useRef<HTMLDivElement | null>(null)
  const [expanded, setExpanded] = useState(false)
  const focusKey = useRef<string | null>(null)
  const pendingFocus = useRef<string | null>(null)
  const seenRequest = useRef(0)
  const btn = button(compact)
  const copied = useMotionProps(tick)
  const unit = kind === 'slides' ? 'slide' : 'page'
  const findOpen = p.showSearch && (!compact || expanded)

  const control = (key: string): HTMLElement | null =>
    root.current?.querySelector<HTMLElement>(key === 'search' ? 'input[type=search]' : `[data-ov-control="${key}"]`) ?? null

  // Cmd/Ctrl-F (possibly queued while the chunk loaded): expand and focus Find.
  // A mounted input is focused at once (an already-open Find re-renders
  // nothing, so a deferred focus would never run); otherwise Find expands first.
  useLayoutEffect(() => {
    const req = p.focusSearchRequest
    if (!p.showSearch || req <= 0 || req === seenRequest.current) return
    seenRequest.current = req
    const input = control('search') as HTMLInputElement | null
    if (input) { input.focus(); input.select(); return }
    setExpanded(true)
    pendingFocus.current = 'search'
  }, [p.focusSearchRequest, p.showSearch])

  // Crossing the compact threshold: focus moves to the counterpart control,
  // and a focused Find stays open.
  const wasCompact = useRef(compact)
  useLayoutEffect(() => {
    if (wasCompact.current === compact) return
    wasCompact.current = compact
    const key = focusKey.current
    if (!key || root.current?.contains(document.activeElement)) return
    const next = control(key) ? key : COUNTERPART[key] ?? key
    if (compact && next === 'search') setExpanded(true)
    if (!compact && key === 'search') setExpanded(false)
    pendingFocus.current = next
  }, [compact])

  useLayoutEffect(() => {
    const key = pendingFocus.current
    if (!key) return
    const el = control(key)
    if (!el) return
    pendingFocus.current = null
    el.focus()
    if (el instanceof HTMLInputElement) el.select()
  })

  const onFocus = (e: React.FocusEvent) => {
    const t = e.target as HTMLElement
    focusKey.current = t instanceof HTMLInputElement || t.closest('[data-testid="officeview-search"]') ? 'search' : t.closest('[data-ov-control]')?.getAttribute('data-ov-control') ?? null
  }
  const onBlur = (e: React.FocusEvent) => {
    const next = e.relatedTarget as Node | null
    // focus left the toolbar (an unmounting control blurs with no target and
    // is no longer connected; keep its key for the counterpart move)
    if (next ? !root.current?.contains(next) : (e.target as Node).isConnected) focusKey.current = null
  }
  const onFindKeyDown = (e: React.KeyboardEvent) => {
    if (!(e.target instanceof HTMLInputElement)) return
    // The query is host-entered text, not protected document content.
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'c') { e.stopPropagation(); return }
    if (e.key !== 'Escape') return
    e.stopPropagation()
    e.preventDefault()
    if (compact) { setExpanded(false); pendingFocus.current = 'find' }
    else p.onEscapeToDocument()
  }

  const fitIcon = <Icon icon={MoveHorizontal} name="fit-width" />
  const copyIcon = p.copyState === 'done' ? <Icon icon={Check} name="check" /> : p.copyState === 'failed' ? <Icon icon={Ban} name="ban" /> : <Icon icon={Copy} name="copy" />
  const more: MoreItem[] = [
    // Fit belongs to the zoom group: hidden with it (showZoomControls={false})
    ...(p.showZoom ? [{ key: 'fit', label: 'Fit to width', icon: fitIcon, onSelect: p.onFit, checked: p.fitActive }] : []),
    ...(p.onCopy ? [{ key: 'copy', label: 'Copy selection', icon: copyIcon, onSelect: p.onCopy, disabled: !p.copyEnabled }] : []),
    ...(p.onPrint ? [{ key: 'print', label: 'Print', icon: <Icon icon={Printer} name="printer" />, onSelect: p.onPrint }] : []),
  ]

  return (
    <div
      ref={root}
      role="toolbar"
      aria-label="Document toolbar"
      data-testid="officeview-toolbar"
      data-officeview-chrome=""
      data-compact={compact ? '' : undefined}
      onFocus={onFocus}
      onBlur={onBlur}
      onKeyDown={(e) => {
        if (!p.showSearch || !(e.metaKey || e.ctrlKey) || e.key.toLowerCase() !== 'f') return
        e.preventDefault()
        setExpanded(true)
        pendingFocus.current = 'search'
        const input = control('search')
        if (input) { input.focus(); (input as HTMLInputElement).select() }
      }}
      style={{ height: PAGED_TOOLBAR_HEIGHT[compact ? 'compact' : 'regular'] }}
      className={`ov:flex ov:items-center ov:shrink-0 ov:min-w-0 ${compact ? 'ov:gap-0 ov:px-1 ov:py-0.5' : 'ov:gap-1 ov:px-2 ov:py-1.5'} ov:bg-chrome ov:text-text ov:border-b ov:border-border ov:text-base`}
    >
      {!compact && p.title && (
        <span data-testid="officeview-title" title={p.title} className="ov:min-w-0 ov:shrink ov:truncate ov:font-[600] ov:px-1">{p.title}</span>
      )}
      <span className="ov:flex-1" />
      {!(compact && findOpen) && (
        <div className="ov:flex ov:items-center ov:gap-0 ov:shrink-0">
          <button type="button" data-ov-control="prev" aria-label={`Previous ${unit}`} disabled={p.current <= 1} onClick={p.onPrev} className={btn}><Icon icon={ChevronLeft} name="chevron-left" /></button>
          <span data-testid="officeview-page-position" aria-label={p.total > 0 ? `${kind === 'slides' ? 'Slide' : 'Page'} ${p.current} of ${p.total}` : 'Loading pages'} className="ov:px-1 ov:whitespace-nowrap ov:tabular-nums ov:text-text">
            {/* no numeric position until the pages are known */}
            {p.total > 0 ? <>{kind === 'slides' && !compact ? 'Slide ' : ''}{p.current} / {p.total}</> : '—'}
          </span>
          <button type="button" data-ov-control="next" aria-label={`Next ${unit}`} disabled={p.current >= p.total} onClick={p.onNext} className={btn}><Icon icon={ChevronRight} name="chevron-right" /></button>
        </div>
      )}
      {p.showZoom && !(compact && findOpen) && (
        <>
          {divider(compact)}
          <div data-testid="officeview-zoom-controls" className="ov:flex ov:items-center ov:gap-0 ov:shrink-0">
            <button type="button" data-ov-control="zoom-out" aria-label="Zoom out" disabled={p.zoom <= 1} onClick={p.onZoomOut} className={btn}><Icon icon={Minus} name="minus" /></button>
            <button type="button" data-ov-control="zoom-reset" aria-label="Reset zoom" onClick={p.onResetZoom} className={`${btn} ov:min-w-11 ov:text-sm ov:font-[600] ov:tabular-nums ov:text-text`}><Tick k={Math.round(p.zoom * 100)}>{Math.round(p.zoom * 100)}%</Tick></button>
            <button type="button" data-ov-control="zoom-in" aria-label="Zoom in" onClick={p.onZoomIn} className={btn}><Icon icon={Plus} name="plus" /></button>
            {!compact && <button type="button" data-ov-control="fit" aria-label="Fit to width" aria-pressed={p.fitActive} onClick={p.onFit} className={btn}>{fitIcon}</button>}
          </div>
        </>
      )}
      {p.showSearch && (
        <>
          {divider(compact)}
          {compact && (
            <button type="button" data-ov-control="find" aria-label="Find" aria-expanded={expanded} onClick={() => { const open = !expanded; setExpanded(open); pendingFocus.current = open ? 'search' : null }} className={btn}>
              <Icon icon={Search} name="search" />
            </button>
          )}
          {findOpen && <div className="ov:flex ov:min-w-0 ov:shrink" onKeyDown={onFindKeyDown}>{p.search}</div>}
        </>
      )}
      {compact ? (
        more.length > 0 && <MoreMenu items={more} buttonClass={btn} />
      ) : (p.onCopy || p.onPrint) && (
        <>
          {divider(compact)}
          {p.onCopy && (
            <button
              type="button"
              data-ov-control="copy"
              data-testid="officeview-copy-selection"
              data-state={COPY_STATE[p.copyState]}
              aria-label={p.copyState === 'failed' ? BLOCKED : 'Copy selection'}
              title={p.copyState === 'failed' ? BLOCKED : 'Copy selection'}
              disabled={!p.copyEnabled}
              onClick={p.onCopy}
              className={`${btn} ov:data-[state=copied]:text-accent`}
            >
              {copyIcon}
              <Presence>{p.copyState === 'done' && <m.span key="copied" {...copied} className="ov:text-sm ov:text-text">Copied</m.span>}</Presence>
            </button>
          )}
          {p.onPrint && <button type="button" data-ov-control="print" aria-label="Print" onClick={p.onPrint} className={btn}><Icon icon={Printer} name="printer" /></button>}
        </>
      )}
      {/* Polite live region: the copy outcome is announced. */}
      <span data-testid="officeview-copy-status" role="status" aria-live="polite" className="ov:sr-only">
        {p.copyState === 'done' ? 'Copied' : p.copyState === 'failed' ? BLOCKED : ''}
      </span>
    </div>
  )
}
