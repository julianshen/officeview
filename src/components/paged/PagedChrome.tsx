/**
 * Paged-viewer chrome, loaded lazily by <OfficeDoc> so none of it weighs on
 * the eager entry. For now it holds the floating Find bar, zoom cluster and
 * copy pill, moved verbatim from OfficeDoc; the toolbar task replaces them.
 * All state and key handling stays in OfficeDoc, which only hands props in.
 */
import { useEffect, type ComponentProps, type ReactElement } from 'react'
import { SearchBar, ZoomControls } from '../viewerControls'

export interface FloatingControlsProps {
  search: ComponentProps<typeof SearchBar> | null
  zoom: ComponentProps<typeof ZoomControls> | null
  /** Non-null only while copying is allowed and a non-empty selection exists. */
  copy: { state: 'idle' | 'done' | 'failed'; onCopy(): void } | null
  /** Bumped by a Cmd/Ctrl-F press that found no search input (the chunk was
   * still loading); a non-zero value focuses Find once it mounts. */
  focusSearchRequest: number
}

// Last request honoured per viewer (its input ref is stable), so a remount or
// re-showing Find does not replay an old Cmd/Ctrl-F and steal focus.
const honoured = new WeakMap<object, number>()

export function FloatingControls({ search, zoom, copy, focusSearchRequest }: FloatingControlsProps): ReactElement {
  const inputRef = search?.inputRef
  useEffect(() => {
    const input = inputRef?.current
    if (!inputRef || !input || focusSearchRequest <= (honoured.get(inputRef) ?? 0)) return
    honoured.set(inputRef, focusSearchRequest)
    input.focus()
    input.select()
  }, [focusSearchRequest, inputRef])
  return (
    <>
      {search && <SearchBar {...search} />}
      {copy && (
        <button
          type="button"
          data-testid="officeview-copy-selection"
          data-state={copy.state === 'done' ? 'copied' : copy.state}
          onClick={copy.onCopy}
          style={{
            position: 'absolute',
            bottom: 'calc(env(safe-area-inset-bottom) + 16px)',
            left: '50%',
            transform: 'translateX(-50%)',
            padding: '8px 16px',
            borderRadius: 18,
            border: 'none',
            background: copy.state === 'failed' ? '#8a1f11' : copy.state === 'done' ? '#1a7f37' : 'rgba(20,22,28,0.92)',
            color: '#fff',
            fontSize: 14,
            cursor: 'pointer',
            boxShadow: '0 2px 10px rgba(0,0,0,0.35)',
            zIndex: 6,
          }}
        >
          {copy.state === 'done' ? 'Copied ✓' : copy.state === 'failed' ? 'Copy blocked — press ⌘C' : 'Copy selection'}
        </button>
      )}
      {zoom && <ZoomControls {...zoom} />}
    </>
  )
}
