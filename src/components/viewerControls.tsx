/**
 * Viewer chrome shared by the page viewer and the interactive sheet viewport:
 * the Ctrl/Cmd+F search bar and the zoom control cluster.
 */
import type { ReactElement } from 'react'
import { ChevronDown, ChevronUp, Minus, Plus } from 'lucide-react'
import { Icon } from './Icon'
import { Tick } from './motion/presets'

/** Shared control styling (spec §6.1, §6.8). Every control declares colour,
 * background, padding and border as ov: utilities, which beat host rules such
 * as `.app button` (control-utility rule). Coarse pointers get 44px targets
 * (spec §513); desktop sizes are main's. */
const TARGET = 'ov:coarse:min-h-11 ov:coarse:min-w-11'
const MOTION = 'ov:transition-colors ov:duration-150 ov:motion-reduce:transition-none'
const FOCUS = 'ov:focus-visible:outline-2 ov:focus-visible:outline-solid ov:focus-visible:outline-offset-1 ov:focus-visible:outline-accent'
const SURFACE = 'ov:bg-surface ov:border ov:border-border ov:rounded-control ov:shadow-popover'
const ICON_BUTTON = `ov:inline-flex ov:items-center ov:justify-center ov:shrink-0 ov:p-0 ov:text-icon ov:cursor-pointer ov:enabled:hover:bg-hover ov:disabled:text-muted ov:disabled:cursor-default ${MOTION} ${FOCUS} ${TARGET}`

export function SearchBar({
  inputRef,
  query,
  onQuery,
  count,
  active,
  indexing,
  onPrev,
  onNext,
  countLabel,
  note,
  failure,
  variant = 'floating',
}: {
  inputRef: React.RefObject<HTMLInputElement | null>
  query: string
  onQuery: (q: string) => void
  count: number
  active: number
  indexing: boolean
  onPrev: () => void
  onNext: () => void
  /** Shown instead of `count`, e.g. a lower bound "200+" while results are partial. */
  countLabel?: string
  /** Qualifier appended in parentheses (unavailable values, incomplete sheets). */
  note?: string
  /** A failed search has no final match count. */
  failure?: string
  /** 'floating' overlays the viewer (sheet viewer); 'inline' sits in a toolbar. */
  variant?: 'floating' | 'inline'
}): ReactElement {
  const qualifier = note ? ` (${note})` : ''
  const status = indexing ? 'indexing…' : failure ?? (count === 0 ? (query ? `no matches${qualifier}` : '') : `${active + 1} of ${countLabel ?? count}${qualifier}`)
  const arrow = `ov:h-7 ov:w-7 ov:rounded-control ov:bg-transparent ov:border ov:border-border ${ICON_BUTTON}`
  return (
    <div
      data-testid="officeview-search"
      data-officeview-chrome=""
      data-variant={variant}
      // floating: never wider than the viewer, the input shrinks and
      // status/buttons wrap; inline: one row that shrinks inside the toolbar
      className={variant === 'inline'
        ? 'ov:flex ov:flex-nowrap ov:items-center ov:gap-1 ov:min-w-0 ov:shrink ov:p-0 ov:text-text ov:bg-transparent ov:border-0'
        : `ov:absolute ov:top-[calc(env(safe-area-inset-top)+8px)] ov:right-3 ov:z-6 ov:max-w-[calc(100%-24px)] ov:box-border ov:flex ov:flex-wrap ov:justify-end ov:items-center ov:gap-1.5 ov:p-1.5 ov:text-text ${SURFACE}`}
    >
      <input
        ref={inputRef}
        type="search"
        value={query}
        placeholder="Search…"
        aria-label="Search document"
        onChange={(e) => onQuery(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') { e.preventDefault(); if (e.shiftKey) onPrev(); else onNext() }
          else if (e.key === 'Escape') { e.preventDefault(); onQuery('') }
        }}
        className={`ov:h-7 ov:w-37 ov:min-w-0 ov:shrink ov:basis-24 ov:grow ov:max-w-37 ov:py-0 ov:px-2 ov:rounded-control ov:border ov:border-border ov:bg-field ov:text-text ov:text-[14px] ov:placeholder:text-muted ov:outline-none ov:focus-visible:border-accent ov:focus-visible:outline-2 ov:focus-visible:outline-solid ov:focus-visible:outline-offset-0 ov:focus-visible:outline-accent ${MOTION} ${TARGET}`}
      />
      <span
        data-testid="officeview-search-status"
        className="ov:min-w-16 ov:text-sm ov:text-muted ov:text-center"
      >
        {status}
      </span>
      <button type="button" onClick={onPrev} disabled={count === 0} aria-label="Previous match" className={arrow}><Icon icon={ChevronUp} name="chevron-up" /></button>
      <button type="button" onClick={onNext} disabled={count === 0} aria-label="Next match" className={arrow}><Icon icon={ChevronDown} name="chevron-down" /></button>
    </div>
  )
}

export function ZoomControls({
  zoom,
  onIn,
  onOut,
  onReset,
  atFit,
}: {
  zoom: number
  onIn: () => void
  onOut: () => void
  onReset: () => void
  atFit: boolean
}): ReactElement {
  const btn = `ov:w-9 ov:h-9 ${SURFACE} ${ICON_BUTTON} ov:disabled:opacity-50`
  return (
    <div
      data-testid="officeview-zoom-controls"
      data-officeview-chrome=""
      className="ov:absolute ov:right-3 ov:bottom-[calc(env(safe-area-inset-bottom)+12px)] ov:z-5 ov:flex ov:flex-col ov:gap-2"
    >
      <button type="button" onClick={onIn} aria-label="Zoom in" className={btn}><Icon icon={Plus} name="plus" size={20} /></button>
      <button type="button" onClick={onOut} aria-label="Zoom out" disabled={atFit} className={btn}><Icon icon={Minus} name="minus" size={20} /></button>
      <button
        type="button"
        onClick={onReset}
        aria-label="Reset zoom"
        className={`ov:w-9 ov:h-7 ov:text-xs ov:font-[600] ov:tabular-nums ov:text-text ${SURFACE} ov:inline-flex ov:items-center ov:justify-center ov:p-0 ov:cursor-pointer ov:hover:bg-hover ${MOTION} ${FOCUS} ${TARGET}`}
      >
        <Tick k={Math.round(zoom * 100)}>{Math.round(zoom * 100)}%</Tick>
      </button>
    </div>
  )
}
