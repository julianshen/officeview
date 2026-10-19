import { useCanvasColors, type CanvasColors } from '../canvasColors'
/**
 * Paged-viewer chrome, loaded lazily by <OfficeDoc> (spec §6.10): the top
 * toolbar, the status line (portalled into OfficeDoc's fixed-height slot
 * below the pages), page tracking, the word count and the Find bar. Zoom
 * state and the keyboard paths stay eager in OfficeDoc, which hands props in.
 */
import { useEffect, useMemo, useRef, useState, type ReactElement, type RefObject } from 'react'
import { createPortal } from 'react-dom'
import type { Paintable } from '../../render/paint'
import { stepMatch, type SearchMatch } from '../../core/search'
import type { Transform } from '../../core/zoom'
import type { OfficeDocSource } from '../OfficeDoc'
import { contentYAt, pageOffsets } from '../scrollPan'
import { SearchBar } from '../viewerControls'
import { MotionRoot } from '../motion/presets'
import { warnIfStylesMissing } from '../styleWarning'
import { PAGE_PADDING_PX } from './metrics'
import { PagedStatus, type PagedKind } from './PagedStatus'
import { PagedToolbar } from './PagedToolbar'
import { usePageTracking, type PagedLayout } from './usePageTracking'
import { useWordCount } from './useWordCount'

export { PagedToolbar } from './PagedToolbar'
export { PagedStatus } from './PagedStatus'

export interface PagedChromeProps {
  onPalette?(palette: CanvasColors & {key: string}): void
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
  /** Print is offered only when allowed. */
  allowPrint: boolean
  pages: readonly Paintable[]
  owner: object
  showToolbar: boolean
  /** A queued keyboard print, tagged with its document/layout owner. */
  print?: { owner: object }
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
  const { containerRef, search } = p
  const themeRoot = useMemo(() => ({ get current() { return containerRef.current?.closest('[data-ov-theme-root]') ?? null } }), [containerRef])
  // Root attribute observation handles prop-driven theme changes too.
  const palette = useCanvasColors(themeRoot, 'system')
  useEffect(() => { p.onPalette?.({...palette.colors, key: palette.key}) }, [palette, p.onPalette])
  // Once per page: is officeview.css loaded? (spec §5, R10)
  useEffect(() => { const root = themeRoot.current; if (root) warnIfStylesMissing(root) }, [themeRoot])
  const [notice, setNotice] = useState('')
  useEffect(() => {
    if (p.allowPrint) return
    const root = containerRef.current?.closest<HTMLElement>('[data-ov-theme-root]')
    if (!root || getComputedStyle(root).getPropertyValue('--ov-sentinel').trim() === '1') return
    const id = root.getAttribute('data-ov-instance')
    if (id === null) return
    // CSS string escaping, not identifier escaping: React ids may contain
    // punctuation, and an embedding host may supply quotes or control bytes.
    const quoted = id.replace(/[\x00-\x1f\x7f"\\]/g, c => `\\${c.codePointAt(0)!.toString(16)} `)
    const style = root.ownerDocument.createElement('style')
    style.setAttribute('data-ov-print-guard', '')
    style.textContent = `@media print{[data-ov-instance="${quoted}"]{display:none!important}}`
    root.ownerDocument.head.append(style)
    return () => style.remove()
  }, [containerRef, p.allowPrint, p.owner])
  const printing = useRef<AbortController | null>(null)
  const beginPrint = () => {
    printing.current?.abort()
    if (!p.allowPrint || !p.pages.length) return
    const root = containerRef.current?.closest<HTMLElement>('[data-ov-theme-root]')
    if (!root) return
    const controller = new AbortController()
    printing.current = controller
    setNotice('')
    void printLoader.load().then(m => {
      if (!controller.signal.aborted) return m.printDocument(p.pages, root, {
        signal: controller.signal, title: p.title,
        onNotice: text => { if (!controller.signal.aborted) setNotice(text) },
      })
    }).catch(() => {
      if (!controller.signal.aborted) setNotice('The document could not be printed. Please try again.')
    })
  }
  useEffect(() => {
    setNotice('')
    return () => { printing.current?.abort() }
  }, [p.owner, p.pages, p.allowPrint])
  useEffect(() => {
    if (p.print?.owner === p.owner) beginPrint()
    // Pages may arrive after a shortcut. Restoring permission must never
    // replay a cancelled request; the separate owner effect aborts it.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [p.print, p.owner, p.pages])
  // A request made while search was off is dropped, not replayed later.
  const request = search && p.focusSearchRequest > (honoured.get(containerRef) ?? 0) ? p.focusSearchRequest : 0
  useEffect(() => { honoured.set(containerRef, Math.max(honoured.get(containerRef) ?? 0, p.focusSearchRequest)) }, [p.focusSearchRequest, containerRef])

  return <>
    {p.showToolbar && <VisibleChrome {...p} focusSearchRequest={request} beginPrint={beginPrint} />}
    <div data-ov-print-notice="" data-officeview-chrome=""
      style={{ position: 'absolute', bottom: p.showToolbar ? 32 : 8, left: 8, right: 8, zIndex: 10, pointerEvents: notice ? 'auto' : 'none' }}
      className={notice ? 'ov:flex ov:items-center ov:gap-2 ov:p-2 ov:bg-chrome ov:text-text ov:border ov:border-border ov:rounded-control' : undefined}>
      <span role="status" className="ov:min-w-0" style={{ overflowWrap: 'anywhere' }}>{notice}</span>
      {notice && <button type="button" aria-label="Dismiss print notice" onClick={() => setNotice('')}
        className="ov:shrink-0 ov:min-w-11 ov:min-h-11 ov:p-0 ov:bg-transparent ov:border-0 ov:text-icon ov:rounded-control ov:hover:bg-hover ov:focus-visible:outline-2 ov:focus-visible:outline-accent">×</button>}
    </div>
  </>
}

// The headless print/guard service never mounts these observers or the word
// counter. Hiding the toolbar unmounts this subtree and runs their cleanups.
function VisibleChrome(p: PagedChromeProps & { beginPrint(): void }): ReactElement {
  const { containerRef, layout, transform, search, beginPrint } = p
  const tracking = usePageTracking(containerRef, layout, transform, p.setTransform)
  const words = useWordCount(p.document)
  const kind = kindOf(p.document)
  const total = layout.specs.length
  const current = Math.min(tracking.current, Math.max(0, total - 1)) + 1

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
    <MotionRoot>
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
        {...(p.allowPrint ? { onPrint: beginPrint } : {})}
        compact={p.compact}
        showSearch={!!search}
        showZoom={p.showZoom}
        focusSearchRequest={p.focusSearchRequest}
        onEscapeToDocument={() => containerRef.current?.focus()}
      />
      {p.statusSlot && createPortal(<PagedStatus kind={kind} current={current} total={total} words={words} />, p.statusSlot)}
    </MotionRoot>
  )
}

export const printLoader = {load: () => import('./printDocument')}
