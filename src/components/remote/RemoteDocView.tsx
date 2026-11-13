/**
 * The remote document view (§6.3/§8.2/§8.3), loaded lazily from <OfficeDoc>.
 * Each unit is a canvas sized in CSS px at scale 1; a unit is painted at the
 * scale bucket for devicePixelRatio when it nears the viewport. Stage 3a adds
 * server-answered search, selection and copy and the screen-reader text layer
 * (RemoteText.tsx); Stage 3b prints server-drawn print rasters
 * (RemotePrint.tsx), never these canvases. There is no zoom or presentation.
 * With copy and text a11y both off the view says "text unavailable in this
 * viewing mode" once.
 */
import { useEffect, useId, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import type { RemoteOfficeDoc } from '../../stream/remote/remote-doc'
import { themeRootProps, type ThemeMode } from '../theme'
import { useRemotePrint } from './RemotePrint'
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
}

export const TEXT_UNAVAILABLE = 'text unavailable in this viewing mode'
/** Belt and braces under the always-hidden root: the chrome (search bar,
 * print controls, hit/selection boxes, screen-reader layer, status line, all
 * marked data-ov-remote-chrome) never prints. Owned by this view, so local
 * documents' print behaviour is untouched. */
const PRINT_CSS = '@media print { [data-ov-remote-chrome] { display: none !important; } }'

/** Host props that do not apply to a remote document are ignored: the baked
 * mark is authoritative (no `watermark`), copy follows the server's `can.copy`,
 * and print follows the server's `can.print`. */
export default function RemoteDocView({ document: doc, theme, className, style, background, pageGapPx = 16, allowPrint = true, title }: RemoteDocViewProps): ReactElement {
  const canvases = useRef<Array<HTMLCanvasElement | null>>([])
  const units = doc.manifest.units.list
  const scroller = useRef<HTMLDivElement | null>(null)
  const text = useRemoteText(doc, scroller)
  const root = useRef<HTMLDivElement | null>(null)
  const print = useRemotePrint(doc, { allowPrint, themeRoot: root, ...(title !== undefined ? { title } : {}) })
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
    <div ref={root} {...themeRootProps(theme, useId())} data-ov-print="false" className={className} style={style}
      onKeyDown={(e) => {
        if ((e.metaKey || e.ctrlKey) && !e.altKey && e.key.toLowerCase() === 'p') {
          // Always swallowed: the browser's own print would print the (hidden) screen.
          e.preventDefault()
          print.start()
          return
        }
        text.onKeyDown(e)
      }}>
      <style>{PRINT_CSS}</style>
      <div ref={scroller} data-officeview-root="" role="document" aria-label={title ?? doc.manifest.name} tabIndex={0}
        style={{ overflow: 'auto', height: '100%', background: background ?? 'var(--ov-color-backdrop)' }}>
        {/* The floating search bar stays at the top of the scroller (zero height: no layout change). */}
        {text.bar || print.ui ? <div data-ov-remote-chrome="" style={{ position: 'sticky', top: 0, height: 0, zIndex: 6 }}>{print.ui}{text.bar}</div> : null}
        {/* One line is always reserved: a copy message coming or going must not
            shift the units under a drag in progress. */}
        <div role="status" data-officeview-notice="" data-ov-remote-chrome="" style={{ minHeight: '1lh' }}>{text.message ?? (textOff ? TEXT_UNAVAILABLE : '')}</div>
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
    </div>
  )
}
