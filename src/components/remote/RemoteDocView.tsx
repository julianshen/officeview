/**
 * The remote document view (§6.3/§8.2/§8.3), loaded lazily from <OfficeDoc>.
 * Each unit is a canvas sized in CSS px at scale 1; a unit is painted at the
 * scale bucket for devicePixelRatio when it nears the viewport. Stage 3a adds
 * server-answered search, selection and copy and the screen-reader text layer
 * (RemoteText.tsx); Stage 3b prints server-drawn print rasters
 * (RemotePrint.tsx), never these canvases. Stage 3c presents a streamed deck
 * (Present button, F5, the `present()` handle): the eager start in present.ts
 * over this view's show host, then the lazy remote show (stream/remote/show.ts).
 * There is no zoom.
 * With copy and text a11y both off the view says "text unavailable in this
 * viewing mode" once.
 */
import { useEffect, useId, useLayoutEffect, useRef, useState, type CSSProperties, type MutableRefObject, type ReactElement } from 'react'
import type { RemoteOfficeDoc } from '../../stream/remote/remote-doc'
import { presentRemote } from '../../stream/remote/present-start'
import { defaultPlatform, type OfficeViewerHandle, type PresentOptions, type PresentPlatform, type SlideShowDiagnostic, type SlideShowMotion } from '../present'
import { themeRootProps, type ThemeMode } from '../theme'
import { BUTTON, useRemotePrint } from './RemotePrint'
import { UnitWrap, useRemoteText } from './RemoteText'

export interface RemoteDocViewProps {
  document: RemoteOfficeDoc
  theme?: ThemeMode
  className?: string
  style?: CSSProperties
  background?: string
  pageGapPx?: number
  allowPrint?: boolean
  title?: string
  /** As on <OfficeDoc>: the host may withhold presenting; the server's `can.present` decides too. */
  allowPresent?: boolean
  presentShortcuts?: boolean
  slideShowMotion?: SlideShowMotion
  onSlideShowStart?: () => void
  onSlideShowEnd?: () => void
  /** Streamed slide index (hidden slides not counted, R19). */
  onSlideShowSlideChange?: (slideIndex: number, buildIndex: number) => void
  onSlideShowDiagnostic?: (diagnostic: SlideShowDiagnostic) => void
  /** <OfficeDoc>'s ref handle reaches this view's start through here while it is mounted. */
  presentRef?: MutableRefObject<OfficeViewerHandle['present'] | null>
  /** @internal Tests inject the full-screen/popup platform. */
  __presentPlatform?: PresentPlatform
}

export const TEXT_UNAVAILABLE = 'text unavailable in this viewing mode'
/** Belt and braces under the always-hidden root: the chrome (search bar,
 * print controls, hit/selection boxes, screen-reader layer, status line, all
 * marked data-ov-remote-chrome) never prints. Owned by this view, so local
 * documents' print behaviour is untouched. */
const PRINT_CSS = '@media print { [data-ov-remote-chrome] { display: none !important; } }'
/** The view's own print guard, so hiding the screen canvases never depends on
 * officeview.css being loaded: this instance's root is display:none in print
 * media. CSS string escaping (as in PagedChrome), not identifier escaping:
 * React ids may contain punctuation. */
export function remotePrintGuard(instance: string): string {
  const quoted = instance.replace(/[\x00-\x1f\x7f"\\]/g, (c) => `\\${c.codePointAt(0)!.toString(16)} `)
  return `@media print{[data-ov-theme-root][data-ov-instance="${quoted}"]{display:none!important}}`
}

/** Host props that do not apply to a remote document are ignored: the baked
 * mark is authoritative (no `watermark`), copy follows the server's `can.copy`,
 * and print follows the server's `can.print`. */
export default function RemoteDocView({
  document: doc, theme, className, style, background, pageGapPx = 16, allowPrint = true, title,
  allowPresent = true, presentShortcuts = true, slideShowMotion = 'auto', onSlideShowStart, onSlideShowEnd, onSlideShowSlideChange, onSlideShowDiagnostic,
  presentRef, __presentPlatform = defaultPlatform,
}: RemoteDocViewProps): ReactElement {
  const canvases = useRef<Array<HTMLCanvasElement | null>>([])
  const units = doc.manifest.units.list
  const scroller = useRef<HTMLDivElement | null>(null)
  const text = useRemoteText(doc, scroller)
  const root = useRef<HTMLDivElement | null>(null)
  const instance = useId()
  const present = useRemotePresent(doc, scroller, canvases, pageGapPx, {
    allowPresent, slideShowMotion, platform: __presentPlatform,
    callbacks: { onStart: onSlideShowStart, onEnd: onSlideShowEnd, onSlideChange: onSlideShowSlideChange, onDiagnostic: onSlideShowDiagnostic },
  })
  useLayoutEffect(() => {
    if (!presentRef) return
    presentRef.current = present.start
    return () => { presentRef.current = null }
  })
  const presentButton = present.offered
    ? <button type="button" data-ov-present-button="" onClick={() => void present.start({ from: 'beginning' }).catch(() => {})} style={BUTTON}>Present</button>
    : null
  const print = useRemotePrint(doc, { allowPrint, themeRoot: root, lead: presentButton, ...(title !== undefined ? { title } : {}) })
  // The observer below is created per document; it reaches the latest a11y handler through this ref.
  const a11yVisible = useRef(text.a11yVisible)
  a11yVisible.current = text.a11yVisible
  const textOff = !doc.manifest.can.copy && !doc.manifest.can.textA11y
  // §10: a session that ends (revoked, expired, oversized frame…) says so once;
  // recovery is a host-driven re-open.
  const [ended, setEnded] = useState<string>()
  useEffect(() => {
    let live = true
    void doc.ended.then((e) => { if (live && e) setEnded(e.code) })
    return () => { live = false }
  }, [doc])
  useEffect(() => {
    const painted = new Set<number>(), ac = new AbortController()
    const dpr = typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1
    const paint = (i: number): Promise<void> => {
      const c = canvases.current[i]
      if (!c || painted.has(i)) return Promise.resolve()
      painted.add(i)
      return doc.paint(i, 1, dpr, c, ac.signal).then(() => {
        if (ac.signal.aborted) return
        c.dataset.painted = '1'
      }, () => { painted.delete(i) })
    }
    if (typeof IntersectionObserver !== 'function') { units.forEach((_, i) => { paint(i); a11yVisible.current(i, true) }); return () => ac.abort() }
    const io = new IntersectionObserver((entries) => {
      for (const e of entries) {
        const i = Number((e.target as HTMLElement).dataset.unit)
        // Screen-reader text follows only the units actually in view (each 'unit'-mode answer is audited).
        a11yVisible.current(i, e.isIntersecting)
        if (!e.isIntersecting) continue
        // The unit entering view and its neighbours.
        for (const j of [i - 1, i, i + 1]) if (j >= 0 && j < units.length) paint(j)
      }
    }, { rootMargin: '200px' })
    for (const c of canvases.current) if (c) io.observe(c)
    return () => { io.disconnect(); ac.abort() }
  }, [doc, units])
  return (
    // The screen canvases never print, by any path (browser menu included):
    // data-ov-print is always "false" on the themed instance root, which the
    // print guard (officeview.css) matches as
    // [data-ov-theme-root][data-ov-instance][data-ov-print="false"]. Printing
    // is server-drawn print rasters only (RemotePrint.tsx).
    <div ref={root} {...themeRootProps(theme, instance)} data-ov-print="false" className={className} style={style}
      onKeyDown={(e) => {
        // F5, Shift+F5 from the current slide, Alt+F5 Presenter View (as the local viewer);
        // Ctrl/Cmd+F5 stays the browser's hard reload.
        if (e.key === 'F5' && presentShortcuts && present.offered && !(e.ctrlKey || e.metaKey)) {
          e.preventDefault()
          void present.start({ from: e.shiftKey ? 'current' : 'beginning', ...(e.altKey ? { presenterView: true } : {}) }).catch(() => {})
          return
        }
        if ((e.metaKey || e.ctrlKey) && !e.altKey && e.key.toLowerCase() === 'p') {
          // Always swallowed: the browser's own print would print the (hidden) screen.
          e.preventDefault()
          print.start()
          return
        }
        text.onKeyDown(e)
      }}>
      <style data-ov-print-guard="">{remotePrintGuard(instance)}</style>
      <style>{PRINT_CSS}</style>
      <div ref={scroller} data-officeview-root="" role="document" aria-label={title ?? doc.manifest.name} tabIndex={0}
        style={{ overflow: 'auto', height: '100%', background: background ?? 'var(--ov-color-backdrop)' }}>
        {/* The floating search bar stays at the top of the scroller (zero height: no layout change). */}
        {text.bar || print.ui ? <div data-ov-remote-chrome="" style={{ position: 'sticky', top: 0, height: 0, zIndex: 6 }}>{print.ui}{text.bar}</div> : null}
        {/* One line is always reserved: a copy message coming or going must not
            shift the units under a drag in progress. */}
        <div role="status" data-officeview-notice="" data-ov-remote-chrome="" style={{ minHeight: '1lh' }}>{text.message ?? (textOff ? TEXT_UNAVAILABLE : '')}</div>
        {present.notice ? <div role="alert" data-ov-present-notice="" data-ov-remote-chrome="">{present.notice}</div> : null}
        {ended ? <div role="alert" data-stream-error={ended}>Viewing session ended ({ended}).</div> : null}
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: pageGapPx, padding: pageGapPx }}>
          {units.map((u, i) => (
            <UnitWrap key={i} index={i} w={u.w} h={u.h} boxes={text.boxes(i)} layer={text.layer(i)} {...(text.onPointerDown ? { onPointerDown: text.onPointerDown } : {})}>
              <canvas data-unit={i} ref={(el) => { canvases.current[i] = el }} aria-hidden="true"
                style={{ display: 'block', width: u.w, maxWidth: '100%', height: 'auto', aspectRatio: `${u.w} / ${u.h}`, background: '#fff', boxShadow: 'var(--ov-shadow-page, 0 1px 3px rgba(0,0,0,.2))' }} />
            </UnitWrap>
          ))}
        </div>
      </div>
      {/* The show host (slideshow spec 4.1): always mounted, empty and hidden, inside the theme root. */}
      <div ref={present.host} data-officeview-show="" tabIndex={-1} style={{ visibility: 'hidden' }} />
    </div>
  )
}

/**
 * Presenting a streamed deck: offered while the server's `can.present`, the
 * host's `allowPresent` and the session all allow it (re-read on a resume).
 * `start` calls `presentRemote` (the eager prologue) synchronously (the
 * gesture rule); the start belongs to the document epoch (a new document,
 * unmount, `allowPresent` withdrawn or the session's end cancels it, even
 * while the show module is still loading, or ends the show). The navigation
 * callbacks speak unit (page) indices.
 */
function useRemotePresent(doc: RemoteOfficeDoc, scroller: MutableRefObject<HTMLDivElement | null>, canvases: MutableRefObject<Array<HTMLCanvasElement | null>>, gap: number, o: {
  allowPresent: boolean; slideShowMotion: SlideShowMotion; platform: PresentPlatform
  callbacks: Parameters<typeof presentRemote>[0]['callbacks']
}) {
  const host = useRef<HTMLDivElement | null>(null)
  const epoch = useRef<AbortController | null>(null)
  const gen = useRef(0)
  const [notice, setNotice] = useState('')
  const [, setManifestRev] = useState(0)
  useEffect(() => doc.onManifestChange?.(() => setManifestRev((n) => n + 1)), [doc])
  const [endedDoc, setEndedDoc] = useState<RemoteOfficeDoc>()
  useEffect(() => {
    let live = true
    void doc.ended.then(() => { if (live) setEndedDoc(doc) })
    return () => { live = false }
  }, [doc])
  useLayoutEffect(() => {
    const c = epoch.current = new AbortController()
    gen.current++
    // A terminal end (revoked, expired, closed) aborts this document's epoch only.
    void doc.ended.then(() => c.abort())
    return () => c.abort()
  }, [doc, o.allowPresent])
  const allow = o.allowPresent && endedDoc !== doc
  const start = (opts?: PresentOptions) => {
    setNotice('')
    const r = scroller.current
    const back = r?.ownerDocument.activeElement ?? null
    const run = presentRemote({
      doc, host: host.current!, opts, platform: o.platform, generation: gen.current, allow, motion: o.slideShowMotion, signal: epoch.current!.signal,
      callbacks: {
        ...o.callbacks,
        currentSlide() {
          if (!r) return 0
          const box = r.getBoundingClientRect(), mid = box.top + box.height / 2
          const list = canvases.current.slice(0, doc.manifest.units.list.length)
          const i = list.findIndex((c) => c && c.getBoundingClientRect().bottom > mid)
          return i < 0 ? Math.max(0, list.length - 1) : i
        },
        scrollToSlide(i) {
          const c = canvases.current[i]
          // Its own scrollTop only: never scrollIntoView, so host ancestors never scroll.
          if (r && c) r.scrollTop += c.getBoundingClientRect().top - r.getBoundingClientRect().top - gap
        },
        restoreFocus() {
          const own = back?.ownerDocument
          const el = back?.isConnected && back !== own?.body && back !== own?.documentElement ? back : r
          ;(el as HTMLElement | null)?.focus?.()
        },
      },
    })
    run.catch((e) => e.code == 'load-failed' && setNotice(e.message))
    return run
  }
  return { host, start, notice, offered: allow && doc.manifest.can.present }
}
