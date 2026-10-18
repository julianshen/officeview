/**
 * Paged-viewer chrome, loaded lazily by <OfficeDoc> so none of it weighs on
 * the eager entry. For now it holds the floating Find bar, zoom cluster and
 * copy pill (moved from OfficeDoc, restyled in Task 2); the toolbar task
 * replaces them.
 * All state and key handling stays in OfficeDoc, which only hands props in.
 */
import { useEffect, type ComponentProps, type ReactElement } from 'react'
import { Ban, Check, Command, Copy } from 'lucide-react'
import { SearchBar, ZoomControls } from '../viewerControls'
import { Icon } from '../Icon'

const BLOCKED = 'Copy blocked — press ⌘C'
/** data-state per copy outcome; the pill's look keys off it. */
const STATE = { idle: 'idle', done: 'copied', failed: 'blocked' } as const
const PILL = [
  'ov:absolute ov:bottom-[calc(env(safe-area-inset-bottom)+16px)] ov:left-1/2 ov:-translate-x-1/2 ov:z-6',
  'ov:inline-flex ov:items-center ov:gap-1.5 ov:px-4 ov:py-2 ov:rounded-full ov:text-[14px] ov:whitespace-nowrap ov:cursor-pointer',
  'ov:bg-surface ov:text-text ov:border ov:border-border ov:shadow-popover ov:hover:bg-hover',
  'ov:data-[state=copied]:bg-accent ov:data-[state=copied]:text-on-accent ov:data-[state=copied]:border-accent',
  'ov:data-[state=blocked]:border-accent',
  'ov:transition-colors ov:duration-150 ov:motion-reduce:transition-none',
  'ov:focus-visible:outline-2 ov:focus-visible:outline-solid ov:focus-visible:outline-offset-2 ov:focus-visible:outline-accent',
  'ov:coarse:min-h-11 ov:coarse:min-w-11',
].join(' ')

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
          data-state={STATE[copy.state]}
          onClick={copy.onCopy}
          aria-label={copy.state === 'failed' ? BLOCKED : undefined}
          className={PILL}
        >
          {copy.state === 'done'
            ? <><Icon icon={Check} name="check" />Copied</>
            : copy.state === 'failed'
              ? <><Icon icon={Ban} name="ban" />Copy blocked — press<Icon icon={Command} name="command" size={14} />C</>
              : <><Icon icon={Copy} name="copy" />Copy selection</>}
        </button>
      )}
      {/* Polite live region, mounted with the controls so a state change is announced. */}
      <span data-testid="officeview-copy-status" role="status" aria-live="polite" className="ov:sr-only">
        {copy?.state === 'done' ? 'Copied' : copy?.state === 'failed' ? BLOCKED : ''}
      </span>
      {zoom && <ZoomControls {...zoom} />}
    </>
  )
}
