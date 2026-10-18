/**
 * Paged-viewer chrome, loaded lazily by <OfficeDoc> (spec §6.10): the top
 * toolbar, the status line (portalled into OfficeDoc's fixed-height slot
 * below the pages), page tracking, the word count and the Find bar. Zoom
 * state and the keyboard paths stay eager in OfficeDoc, which hands props in.
 */
import { useEffect, type ReactElement, type RefObject } from 'react'
import { createPortal } from 'react-dom'
import { stepMatch, type SearchMatch } from '../../core/search'
import type { Transform } from '../../core/zoom'
import type { OfficeDocSource } from '../OfficeDoc'
import { contentYAt, pageOffsets } from '../scrollPan'
import { SearchBar } from '../viewerControls'
import { PAGE_PADDING_PX } from './metrics'
import { PagedStatus, type PagedKind } from './PagedStatus'
import { PagedToolbar } from './PagedToolbar'
import { usePageTracking, type PagedLayout } from './usePageTracking'
import { useWordCount } from './useWordCount'

export { PagedToolbar } from './PagedToolbar'
export { PagedStatus } from './PagedStatus'

export interface PagedChromeProps {
  document: OfficeDocSource
  title?: string
  compact: boolean
  /** OfficeDoc's reserved status-line slot below the scroll container. */
  statusSlot: HTMLElement | null
  containerRef: RefObject<HTMLDivElement | null>
  layout: PagedLayout
  transform: Transform
  setTransform(t: Transform): void
  /** The zoom shown in the readout (z0 while fit mode is on). */
  zoom: number
  onZoom(direction: 1 | -1): void
  onReset(): void
  fitActive: boolean
  onFit(): void
  showZoom: boolean
  search: {
    inputRef: RefObject<HTMLInputElement | null>
    query: string
    onQuery(q: string): void
    matches: SearchMatch[]
    active: number
    setActive(update: (current: number) => number): void
    indexing: boolean
  } | null
  /** Undefined when copying is not allowed. */
  copy?: { state: 'idle' | 'done' | 'failed'; enabled: boolean; onCopy(): void }
  onPrint?: () => void
  /** Bumped by a Cmd/Ctrl-F press that found no search input. */
  focusSearchRequest: number
}

export function kindOf(doc: OfficeDocSource): PagedKind {
  return 'slides' in doc ? 'slides' : 'sections' in doc ? 'document' : 'sheet-pages'
}

// Last Find request honoured per viewer (its container ref is stable), so a
// remount or re-enabled search never replays an old Cmd/Ctrl-F.
const honoured = new WeakMap<object, number>()

export default function PagedChrome(p: PagedChromeProps): ReactElement {
  const { containerRef, layout, transform, search } = p
  const tracking = usePageTracking(containerRef, layout, transform, p.setTransform)
  const words = useWordCount(p.document)
  const kind = kindOf(p.document)
  const total = layout.specs.length
  const current = Math.min(tracking.current, Math.max(0, total - 1)) + 1

  // A request made while search was off is dropped, not replayed later.
  const request = search && p.focusSearchRequest > (honoured.get(containerRef) ?? 0) ? p.focusSearchRequest : 0
  useEffect(() => { honoured.set(containerRef, Math.max(honoured.get(containerRef) ?? 0, p.focusSearchRequest)) }, [p.focusSearchRequest, containerRef])

  // Bring the active match into view inside the viewer (never scrollIntoView).
  const match = search?.matches[search.active]
  useEffect(() => {
    const el = containerRef.current
    const spec = match && layout.specs[match.pageIndex]
    if (!el || !match || !spec) return
    const o = pageOffsets(layout.specs, layout.displayWidth, layout.gap)
    const scale = layout.displayWidth / spec.widthPx
    const rect = match.rects[0]
    const y = o.top[match.pageIndex] + (rect ? rect.y * scale : 0)
    const s = { scrollTop: el.scrollTop, panY: transform.panY, zoom: transform.zoom, padding: PAGE_PADDING_PX }
    const top = contentYAt(0, s)
    const bottom = contentYAt(el.clientHeight, s)
    if (y >= top && y + (rect ? rect.height * scale : 0) <= bottom) return
    tracking.reveal(Math.max(o.top[match.pageIndex], y - (bottom - top) / 3))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [match])

  const go = (index: number) => tracking.goTo(Math.max(0, Math.min(total - 1, index)))
  return (
    <>
      <PagedToolbar
        title={p.title}
        kind={kind}
        current={current}
        total={total}
        onPrev={() => go(current - 2)}
        onNext={() => go(current)}
        zoom={p.zoom}
        onZoomIn={() => p.onZoom(1)}
        onZoomOut={() => p.onZoom(-1)}
        onResetZoom={p.onReset}
        onFit={p.onFit}
        fitActive={p.fitActive}
        search={search && (
          <SearchBar
            variant="inline"
            inputRef={search.inputRef}
            query={search.query}
            onQuery={search.onQuery}
            count={search.matches.length}
            active={search.active}
            indexing={search.indexing}
            onPrev={() => search.setActive((c) => stepMatch(search.matches, c, -1))}
            onNext={() => search.setActive((c) => stepMatch(search.matches, c, 1))}
          />
        )}
        {...(p.copy ? { onCopy: p.copy.onCopy, copyEnabled: p.copy.enabled } : {})}
        copyState={p.copy?.state ?? 'idle'}
        {...(p.onPrint ? { onPrint: p.onPrint } : {})}
        compact={p.compact}
        showSearch={!!search}
        showZoom={p.showZoom}
        focusSearchRequest={request}
        onEscapeToDocument={() => containerRef.current?.focus()}
      />
      {p.statusSlot && createPortal(<PagedStatus kind={kind} current={current} total={total} words={words} />, p.statusSlot)}
    </>
  )
}
