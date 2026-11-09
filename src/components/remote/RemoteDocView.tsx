/**
 * The remote document view (§6.3/§8.2/§8.3), loaded lazily from <OfficeDoc>.
 * Each unit is a canvas sized in CSS px at scale 1; a unit is painted at the
 * scale bucket for devicePixelRatio when it nears the viewport. Stage 1 has
 * no text layer, search, copy, zoom or presentation for a remote document:
 * the view says "text unavailable in this viewing mode" once.
 */
import { useEffect, useId, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import type { RemoteOfficeDoc } from '../../stream/remote/remote-doc'
import { themeRootProps, type ThemeMode } from '../theme'

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

/** Host props that do not apply to a remote document are ignored: the baked
 * mark is authoritative (no `watermark`), copy does not exist in Stage 1, and
 * print follows the server's `can.print`. */
export default function RemoteDocView({ document: doc, theme, className, style, background, pageGapPx = 16, allowPrint = true, title }: RemoteDocViewProps): ReactElement {
  allowPrint = allowPrint && doc.manifest.can.print
  const canvases = useRef<Array<HTMLCanvasElement | null>>([])
  const units = doc.manifest.units.list
  // §10: a session that ends (revoked, expired, oversized frame…) says so once;
  // recovery is a host-driven re-open.
  const [ended, setEnded] = useState<string>()
  useEffect(() => {
    let live = true
    void doc.ended.then((e) => { if (live && e) setEnded(e.code) })
    return () => { live = false }
  }, [doc])
  // Printing waits for every unit: a unit never near the viewport has no pixels yet.
  const [paintedCount, setPaintedCount] = useState(0)
  const paintAll = useRef<() => Promise<void>>(async () => {})
  useEffect(() => {
    const painted = new Set<number>(), done = new Set<number>(), ac = new AbortController()
    const dpr = typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1
    setPaintedCount(0)
    const paint = (i: number): Promise<void> => {
      const c = canvases.current[i]
      if (!c || painted.has(i)) return Promise.resolve()
      painted.add(i)
      return doc.paint(i, 1, dpr, c, ac.signal).then(() => {
        if (ac.signal.aborted) return
        c.dataset.painted = '1'; done.add(i); setPaintedCount(done.size)
      }, () => { painted.delete(i) })
    }
    paintAll.current = async () => { await Promise.all(units.map((_, i) => paint(i))) }
    const beforePrint = () => { void paintAll.current() }
    window.addEventListener('beforeprint', beforePrint)
    ac.signal.addEventListener('abort', () => window.removeEventListener('beforeprint', beforePrint))
    if (typeof IntersectionObserver !== 'function') { units.forEach((_, i) => paint(i)); return () => ac.abort() }
    const io = new IntersectionObserver((entries) => {
      for (const e of entries) {
        if (!e.isIntersecting) continue
        const i = Number((e.target as HTMLElement).dataset.unit)
        // The unit entering view and its neighbours.
        for (const j of [i - 1, i, i + 1]) if (j >= 0 && j < units.length) paint(j)
      }
    }, { rootMargin: '200px' })
    for (const c of canvases.current) if (c) io.observe(c)
    return () => { io.disconnect(); ac.abort() }
  }, [doc, units])
  return (
    // data-ov-print sits on the themed instance root: the print guard
    // (officeview.css) matches [data-ov-theme-root][data-ov-instance][data-ov-print="false"].
    <div {...themeRootProps(theme, useId())} data-ov-print={allowPrint && paintedCount === units.length} className={className} style={style}
      onKeyDown={(e) => {
        if (!(e.metaKey || e.ctrlKey) || e.key.toLowerCase() !== 'p') return
        // Like the local viewer, Ctrl/Cmd+P prints this document — once every unit has pixels.
        e.preventDefault()
        if (allowPrint) void paintAll.current().then(() => { requestAnimationFrame(() => window.print()) })
      }}>
      <div data-officeview-root="" role="document" aria-label={title ?? doc.manifest.name}
        style={{ overflow: 'auto', height: '100%', background: background ?? 'var(--ov-color-backdrop)' }}>
        <div role="status" data-officeview-notice="">{TEXT_UNAVAILABLE}</div>
        {ended ? <div role="alert" data-stream-error={ended}>Viewing session ended ({ended}).</div> : null}
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: pageGapPx, padding: pageGapPx }}>
          {units.map((u, i) => (
            <canvas key={i} data-unit={i} ref={(el) => { canvases.current[i] = el }} aria-hidden="true"
              style={{ width: u.w, maxWidth: '100%', height: 'auto', aspectRatio: `${u.w} / ${u.h}`, background: '#fff', boxShadow: 'var(--ov-shadow-page, 0 1px 3px rgba(0,0,0,.2))' }} />
          ))}
        </div>
      </div>
    </div>
  )
}
