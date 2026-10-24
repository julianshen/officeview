import { useEffect, useRef, useState } from 'react'
import { getPaintables, loadOfficeFile, type Paintable } from 'officeview'

/** Render page thumbnails without the viewer: load, extract paintable pages, draw each to a canvas. */
export default function PageThumbnails({ url }: { url: string }) {
  const [pages, setPages] = useState<Paintable[]>([])

  useEffect(() => {
    let cancelled = false
    let dispose = () => {}
    loadOfficeFile({ url })
      .then((doc) => getPaintables(doc))
      .then((paintables) => {
        dispose = paintables.dispose
        if (cancelled) dispose()
        else setPages([...paintables])
      })
    return () => { cancelled = true; dispose() }
  }, [url])

  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, padding: 16, overflow: 'auto', alignContent: 'flex-start' }}>
      {pages.map((page, i) => <Thumbnail key={i} page={page} label={`Page ${i + 1}`} />)}
    </div>
  )
}

function Thumbnail({ page, label, width = 180 }: { page: Paintable; label: string; width?: number }) {
  const canvas = useRef<HTMLCanvasElement>(null)
  useEffect(() => {
    const el = canvas.current!
    const scale = (width / page.spec.widthPx) * devicePixelRatio
    el.width = Math.round(page.spec.widthPx * scale)
    el.height = Math.round(page.spec.heightPx * scale)
    const ctx = el.getContext('2d')!
    ctx.scale(scale, scale)
    page.paint(ctx)
  }, [page, width])
  return (
    <figure style={{ margin: 0, font: '13px system-ui', color: '#605e5c', textAlign: 'center' }}>
      <canvas ref={canvas} style={{ width, display: 'block', boxShadow: '0 1px 4px rgba(0,0,0,.2)', background: '#fff' }} />
      <figcaption style={{ marginTop: 6 }}>{label}</figcaption>
    </figure>
  )
}
