/**
 * The remote document view (§6.3/§8.2/§8.3), loaded lazily from <OfficeDoc>.
 * Each unit is a canvas sized in CSS px at scale 1; a unit is painted at the
 * scale bucket for devicePixelRatio when it nears the viewport. Stage 1 has
 * no text layer, search, copy, zoom or presentation for a remote document:
 * the view says "text unavailable in this viewing mode" once.
 */
import { useEffect, useId, useRef, type CSSProperties, type ReactElement } from 'react'
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
  useEffect(() => {
    const painted = new Set<number>(), ac = new AbortController()
    const dpr = typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1
    const paint = (i: number) => {
      const c = canvases.current[i]
      if (!c || painted.has(i)) return
      painted.add(i)
      doc.paint(i, 1, dpr, c, ac.signal).catch(() => { painted.delete(i) })
    }
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
    <div {...themeRootProps(theme, useId())} className={className} style={style}>
      <div data-officeview-root="" data-ov-print={allowPrint} role="document" aria-label={title ?? doc.manifest.name}
        style={{ overflow: 'auto', height: '100%', background: background ?? 'var(--ov-color-backdrop)' }}>
        <div role="status" data-officeview-notice="">{TEXT_UNAVAILABLE}</div>
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
