/**
 * The paged status line (spec §6.5): page or slide position, plus the word
 * count for Word-like documents. Its height comes from the eager metrics, so
 * it replaces the placeholder slot without moving anything.
 */
import type { ReactElement } from 'react'
import { PAGED_STATUS_HEIGHT } from './metrics'

export type PagedKind = 'document' | 'slides' | 'sheet-pages'

export function PagedStatus({ kind, current, total, words }: { kind: PagedKind; current: number; total: number; words: number | null }): ReactElement {
  return (
    <div
      data-testid="officeview-status"
      data-officeview-chrome=""
      style={{ height: PAGED_STATUS_HEIGHT }}
      className="ov:flex ov:items-center ov:gap-2 ov:px-3 ov:text-sm ov:text-muted ov:bg-chrome ov:border-t ov:border-border ov:whitespace-nowrap ov:overflow-hidden ov:tabular-nums ov:select-none"
    >
      <span>{kind === 'slides' ? 'Slide' : 'Page'} {current} of {total}</span>
      {kind === 'document' && (
        <>
          {' '}<span aria-hidden="true">·</span>{' '}
          {/* fixed min-width: the count arriving never moves the line */}
          <span data-testid="officeview-word-count" className="ov:min-w-20">{words === null ? '—' : words.toLocaleString('en-US')} words</span>
        </>
      )}
    </div>
  )
}
