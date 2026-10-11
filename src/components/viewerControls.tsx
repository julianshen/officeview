/**
 * Viewer chrome shared by the page viewer and the interactive sheet viewport:
 * the Ctrl/Cmd+F search bar and the zoom control cluster.
 */
import type { CSSProperties, ReactElement } from 'react'

export function SearchBar({
  inputRef,
  query,
  onQuery,
  count,
  active,
  indexing,
  onPrev,
  onNext,
}: {
  inputRef: React.RefObject<HTMLInputElement | null>
  query: string
  onQuery: (q: string) => void
  count: number
  active: number
  indexing: boolean
  onPrev: () => void
  onNext: () => void
}): ReactElement {
  const status = indexing ? 'indexing…' : count === 0 ? (query ? 'no matches' : '') : `${active + 1} of ${count}`
  const btn: CSSProperties = {
    height: 28,
    minWidth: 28,
    padding: '0 8px',
    borderRadius: 6,
    border: '1px solid rgba(0,0,0,0.2)',
    background: count > 0 ? '#f0f0f0' : 'rgba(0,0,0,0.06)',
    color: count > 0 ? '#111' : '#888',
    fontSize: 14,
    cursor: count > 0 ? 'pointer' : 'default',
  }
  return (
    <div
      data-testid="officeview-search"
      style={{
        position: 'absolute',
        top: 'calc(env(safe-area-inset-top) + 8px)',
        right: 12,
        display: 'flex',
        alignItems: 'center',
        gap: 6,
        padding: 6,
        borderRadius: 10,
        background: 'rgba(255,255,255,0.95)',
        boxShadow: '0 2px 8px rgba(0,0,0,0.3)',
        zIndex: 6,
      }}
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
        style={{
          height: 28,
          width: 148,
          padding: '0 8px',
          borderRadius: 6,
          border: '1px solid rgba(0,0,0,0.2)',
          fontSize: 14,
          outline: 'none',
        }}
      />
      <span
        data-testid="officeview-search-status"
        style={{ minWidth: 64, fontSize: 12, color: '#555', textAlign: 'center' }}
      >
        {status}
      </span>
      <button type="button" onClick={onPrev} disabled={count === 0} aria-label="Previous match" style={btn}>↑</button>
      <button type="button" onClick={onNext} disabled={count === 0} aria-label="Next match" style={btn}>↓</button>
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
  const btn: CSSProperties = {
    width: 36,
    height: 36,
    borderRadius: 18,
    border: '1px solid rgba(0,0,0,0.15)',
    background: '#fff',
    color: '#111',
    fontSize: 18,
    lineHeight: 1,
    cursor: 'pointer',
    boxShadow: '0 1px 4px rgba(0,0,0,0.25)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  }
  return (
    <div
      data-testid="officeview-zoom-controls"
      style={{
        position: 'absolute',
        right: 12,
        bottom: 'calc(env(safe-area-inset-bottom) + 12px)',
        display: 'flex',
        flexDirection: 'column',
        gap: 8,
        zIndex: 5,
      }}
    >
      <button type="button" onClick={onIn} aria-label="Zoom in" style={btn}>+</button>
      <button type="button" onClick={onOut} aria-label="Zoom out" disabled={atFit} style={{ ...btn, opacity: atFit ? 0.5 : 1 }}>−</button>
      <button
        type="button"
        onClick={onReset}
        aria-label="Reset zoom"
        style={{ ...btn, width: 36, height: 28, fontSize: 11, fontWeight: 600 }}
      >
        {Math.round(zoom * 100)}%
      </button>
    </div>
  )
}
